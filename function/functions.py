from statistics import mean, stdev
from subprocess import check_output
import datetime
import scipy.stats as stats
import PySimpleGUI as sg
import csv
import base64

# Carga imagenes del layout e icono
from image.icono import *


# -----------Icono-----------
icon_bytes = base64.b64decode(icon)

# -------------------------Mensajes popup-------------------------
# Popup de error
def error_popup(message):
    sg.Window('Error', [[sg.T('{}'.format(message))], [sg.B('OK', bind_return_key=True, size=(4, 1))]],
              element_justification='c', icon=icon_bytes, modal=True).read(close=True)


# Popup informativo
def info_popup(message):
    sg.Window('Informacion', [[sg.T('{}'.format(message))], [sg.B('OK', bind_return_key=True, size=(4, 1))]],
              element_justification='c', icon=icon_bytes, modal=True).read(close=True)


# Ventana con lista de archivos no procesados
def error_files_popup(files):
    sg.Window('Error', [[sg.T('Los siguientes archivos no pudieron procesarse:')], [
        sg.Multiline(default_text=files, write_only=True, expand_x=True, expand_y=True, size=(30, 10))],
                        [sg.Push(), sg.B('OK', bind_return_key=True, size=(4, 1))]], resizable=True,
              icon=icon_bytes, modal=True).read(close=True)


# Ventana informativa con el listado de voltajes del autozero
def autozero_popup(values):
    sg.Window('Resultados Autozero', [[sg.T('Lista de valores de tension utilizados para el Autozero:')], [
        sg.Multiline(default_text=values, write_only=True, expand_x=True, expand_y=True, size=(30, 10))],
                                      [sg.Push(), sg.B('OK', bind_return_key=True, size=(4, 1))]], resizable=True,
              icon=icon_bytes, modal=True).read(close=True)


# -------------------------Funciones de procesamiento-------------------------
# Devuelve las variables necesarias para establecer el formato CSV seleccionado.
def formato_csv(option):
    # La opcion "0" lee el formato de "separacion de listas" y el "simbolo decimal" del registro de windows.
    if option == 0:
        # Envia el comando al CMD y luego se aisla el valor de los parametros buscados.
        try:
            salida = check_output('Reg Query "HKEY_CURRENT_USER\Control Panel\International" /v sList',
                                  shell=True)
            salida = salida.decode("utf-8").split("\n")
            seplist = salida[2].replace('    sList    REG_SZ    ', '').replace('\r', '')
            salida = check_output('Reg Query "HKEY_CURRENT_USER\Control Panel\International" /v sDecimal',
                                  shell=True)
            salida = salida.decode("utf-8").split("\n")
            decsep = salida[2].replace('    sDecimal    REG_SZ    ', '').replace('\r', '')
            info_popup('Según el registro del sistema el separdor de LISTAS es "{}"'.format(seplist) +
                        ' y el simbolo DECIMAL es "{}"'.format(seplist))
        except Exception as e:
            # Ante falla de la deteccion automatica se avisa y se usan valores por default.
            print(e)
            error_popup('''Fallo el modo automatico, se establecera por defecto separdor de LISTAS "," y el simbolo DECIMAL "."''')
            seplist = ','
            decsep = '.'
    elif option == 1:
        seplist = ','
        decsep = '.'
    elif option == 2:
        seplist = ';'
        decsep = ','
    else:
        # Caso remoto de suceder.
        sg.popup('Algo raro paso en la eleccion del formato de salida del CSV', title='Error',
                 keep_on_top=True, icon=icon_bytes)
        seplist = ','
        decsep = '.'
    return seplist, decsep


# Determinacion del voltaje de referencia de cada sensor del instrumento.
def reference_voltage(path):
    # Diccionario por defecto de los voltajes de referencia de los sensores. Maximo SAPY 32 sensores. Valor pod default 1.
    vref = {f'V{i}': 1 for i in range(1, 33)}
    with open(path) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=';')
        # Extraigo todas las filas
        data_row = []  # Incializacion variable donde se guardan los datos en bruto del CSV.
        for csv_row in csv_reader:
            data_row.append(csv_row)
    data_row.pop(-1)  # Se elimina ultima fila con el caracter #
    # Determinaciòn tipo de formato de entrada. Con o sin tiempo incluido.
    if data_row[0][0] == '>T':
        format_input_file = 'B'
    else:
        format_input_file = 'A'
    # Calculo de los voltajes de referencia para diferentes formatos de datos
    if format_input_file == 'A':
        # Calculo el valor promedio de voltaje de cada sensor del archivo seleccionado.
        for i in range(int(len(data_row) / 2)):
            numbsenor = data_row[2 * i][1]  # Se utiliza la estrategia de que los valores vienen en pares
            values = data_row[i * 2][2:-1]  # Se extraen los valores de los voltajes.
            values = [float(i.replace(',', '.')) for i in values]  # Convierto valores de la lista a float.
            averang = sum(values) / len(values)  # Calculo el promedio de Vout.
            averang = float('%.6f' % averang)  # Reducir el numero de cifras a 6
            vref["V{}".format(numbsenor)] = averang  # Modifico valor del diccionario para el sensor especifico.
    elif 'B':
        # Determinacion del numero de cada sensor utilizado.
        # Notas codigo: En la posicion 3 empiezan los valores de  presión ; El -1 es por la presencia de ">"
        header = [int(line.replace("toma_", "")) for line in data_row[0][3:-1]]
        data_row.pop(0)  # Se elimina el encabezado
        count = 0  # Contador utilizado para determinar el numero de sensor procesado
        for i in range(3, len(data_row[0]) - 1):
            # Valor de referencia del sensor analizado.
            numbsenor = header[count]
            # Se suma el contador ya que se definio el "numbsenor"
            count += 1
            values = []  # Reinicio de la variable
            for j in range(len(data_row)):
                values.append(float(data_row[j][i].replace(',', '.')))
            averang = sum(values) / len(values)  # Calculo el promedio de Vout.
            averang = float('%.6f' % averang)  # Reducir el numero de cifras a 6
            vref["V{}".format(numbsenor)] = averang  # Modifico valor del diccionario para el sensor especifico.
    return vref


# Procesamiento de las presiones y la incertidumbre
def data_process(data_csv, vref, filename, nivconf):
    data_out = {}  # Inicializacion variable donde se guardan los resultados de cada archivo csv procesado.
    data_out.update({'Archivo': filename})  # Guardado nombre de archivo
    # --------------Procesamiento de los datos en bruto--------------
    data = []  # Inicializacion variable de guardado de los datos.
    data_csv.pop(-1)  # Se elimina ultima fila con el caracter #
    # Determinaciòn tipo de formato de entrada. Con o sin tiempo incluido.
    if data_csv[0][0] == '>T':
        format_input_file = 'B'
    else:
        format_input_file = 'A'
    # Calculo de presion para diferentes formatos de datos
    if format_input_file == 'A':
        for line in data_csv:
            # Se elimina el primer (M o V) y el ultimo (>) elemento
            line = line[1:-1]
            # Conversion de string a float de todos los valores del CSV.
            data_buffer = [float(i.replace(',', '.')) for i in line]
            data.append(data_buffer)
        # --------------Procesamiento de las presiones--------------
        for i in range(int(len(data) / 2)):  # Se utiliza la estrategia de que los valores de las tomas vienen en pares
            # Valor de referencia del sensor analizado.
            numbsenor = int(data[2 * i][0])
            V0 = vref["V{}".format(numbsenor)]  # Extraigo del diccionario el valor de referencia.
            # Calculo las presiones para el sensor indicado.
            data_pressure = []  # Inicializo la variable donde guardo las presiones.
            for j in range(1, len(data[0])):
                Vout = data[i * 2][j]
                Vs = data[i * 2 + 1][j]
                value = (((Vout - V0) / (Vs * 0.2)) * 1000)  # Calculo de presion en Pascales
                value = float('%.4f' % value)  # Reduccion a 4 cifras.
                data_pressure.append(value)
            # Guardado de datos en variable de salida
            data_out.update({"Presion-Sensor {}".format(numbsenor): data_pressure})  # Agregado de datos de presiones
    else:
        # Determinacion de los numero de sensores guardados
        # Notas codigo: En la posicion 3 empiezan los sensores de presión ; El -1 es por la presencia de ">"
        header = [int(line.replace("toma_", "")) for line in data_csv[0][3:-1]]
        data_csv.pop(0)  # Se elimina el encabezado
        # Extracion y conversion del tiempo en segundos. Se redondea a 4 cifras
        time_value = [round(float(line[1])* 1e-6, 4) for line in data_csv]
        # Guardado de datos del tiempo
        data_out.update({"Tiempo medicion": time_value})
        del time_value
        # Se analiza los datos por columna usando un contandor para determinar el numero de sensor de cada columna
        count = 0  # Contador utilizado para determinar numero de sensor
        for i in range(3, len(data_csv[0]) - 1):
            # Valor de referencia del sensor analizado.
            numbsenor = header[count]
            V0 = vref["V{}".format(numbsenor)]  # Extraigo del diccionario el valor de referencia.
            # Se suma el contador ya que se definio el "numbsenor"
            count += 1
            # Calculo las presiones para el sensor indicado.
            pressure = []  # Inicializo la variable donde guardo las presiones.
            for j in range(len(data_csv)):
                Vout = float(data_csv[j][i].replace(',', '.'))
                Vs = float(data_csv[j][2].replace(',', '.'))
                value = (((Vout - V0) / (Vs * 0.2)) * 1000)  # Calculo de presion en Pascales
                value = float('%.4f' % value)  # Reduccion a 4 cifras
                pressure.append(value)
            # Guardado de datos en variable de salida
            data_out.update({"Presion-Sensor {}".format(numbsenor): pressure})  # Agregado de datos de presiones
    # -------------- Calculo de la incertidumbre --------------
    # Se determina el numero de sensores de los keys a partir del diccionario "data_out"
    pressure_list = [k for k in list(data_out.keys()) if 'Presion-Sensor' in k]
    crit = 10  # Criterio de contribucion dominante. Se eligio 10 veces superior.
    for i in pressure_list:
        # Extraigo datos de presiones.
        data_raw = data_out[i]
        # Numero del sensor. Se obtiene del key del diccionario
        numb_probe = i.replace('Presion-Sensor ','')
        # Calculo de incertidumbre.
        sample = len(data_raw)  # Numero de muestras.
        data_out.update({"Muestras-{}".format(numb_probe): sample})
        averange = mean(data_raw)  # Estimado de la medicion.
        data_out.update({"Promedio-{}".format(numb_probe): averange})
        if sample > 1:
            typea = stdev(data_raw) / (sample ** 0.5)  # Desviación típica experimental.
            data_out.update({"Tipo A-{}".format(numb_probe): typea})
            # Se diferencia el calculo del componente Tipo B para los diferentes sensores.
            typeb = averange * 0.015 / (3 ** 0.5)  # Componente Tipo B debido a la calibración del sensor de presion.
            data_out.update({"Tipo B-presion-{}".format(numb_probe): typeb})
            ucomb = (typea ** 2 + typeb ** 2) ** 0.5  # Incertidumbre combinada.
            data_out.update({"Incertidumbre Combinada-{}".format(numb_probe): ucomb})
            # Analisis de la contribucion dominante para la determinacion de la incertidumbre expandida.
            try:  # Se evita la division por cero.
                rel_tipe = typea / typeb
            except Exception as e:
                print(e)
                rel_tipe = 1e10
            # Analisis del tipo de distribucion
            if rel_tipe > crit:
                k = stats.t.ppf((1 + nivconf) / 2, sample - 1)  # t student doble cola. t.ppf(alfa, gl)
                distrib = 't-student con {} GL'.format(sample - 1)
                # Guardado datos
                data_out.update({"Coeficiente Expansion-{})".format(numb_probe): k})
                data_out.update({"Tipo distribucion-{}".format(numb_probe): distrib})
            elif typea / typeb < 1 / crit:
                k = (3 ** 0.5) * nivconf  # Distribucion rectangular k=raiz(3)*p
                distrib = 'Rectangular'
                # Guardado datos
                data_out.update({"Coeficiente Expansion-{}".format(numb_probe): k})
                data_out.update({"Tipo distribucion-{}".format(numb_probe): distrib})
            else:
                k = stats.norm.ppf((1 + nivconf) / 2)  # Cumple teorema limite central. Distribución Normal.
                distrib = 'Normal TCLimite'
                # Guardado datos
                data_out.update({"Coeficiente Expansion-{})".format(numb_probe): k})
                data_out.update({"Tipo distribucion-{}".format(numb_probe): distrib})
            # Incertidumbre expandida
            uexpand = k * ucomb
            data_out.update({"Uexpandida ({}%)-{}".format(nivconf * 100, numb_probe): uexpand})
        else:
            # En mediciones de un solo valor no es posible calcular la incertidumbre. Se aplica N/A a todos los datos
            data_out.update({"Tipo A-{}".format(numb_probe): 'N/A'})
            data_out.update({"Tipo B-presion-{}".format(numb_probe): 'N/A'})
            data_out.update({"Incertidumbre Combinada-{}".format(numb_probe): 'N/A'})
            data_out.update({"Coeficiente Expansion-{}".format(numb_probe): 'N/A'})
            data_out.update({"Tipo distribucion-{}".format(numb_probe): 'N/A'})
            data_out.update({"Uexpandida ({}%)-{}".format(nivconf * 100, numb_probe): 'N/A'})
    return data_out


# ------------------------- Guardados de archivos CSV -------------------------
# Guardado de los datos de las presiones en archivo CSV
def save_csv_pressure(save_pressure, path, seplist, decsep):
    save_data = []  # Variable buffer para grabacion de datos.
    # Determinar la longitud mas larga de las listas de presiones.
    # Puede existir mediciones con numeros diferentes de muestras.
    max_len = 0
    for i in range(len(save_pressure)):
        # Maxima longitud de save_pressure["Presion-Sensor x"] siendo "Presion-Sensor x" el primer sensor guardado
        # en el diccionario.
        long = len(save_pressure[i][list(save_pressure[i].keys())[1]])
        if max_len < long:
            max_len = long
    # Listado de sensores usados.
    for i in range(len(save_pressure)):
        # Listado de keys para cada formato de entrada (con o sin tiempo)
        if 'Tiempo medicion' in list(save_pressure[i].keys()):  # Si tiene tiempo se agrega los datos.
            list_sensor = ['Tiempo medicion']
            list_sensor.extend([l for l in list(save_pressure[i].keys()) if 'Presion-Sensor' in l])
        else:
            list_sensor = [l for l in list(save_pressure[i].keys()) if 'Presion-Sensor' in l]
        # Armado de la estructura de datos para ser guardada de cada sensor.
        for j in list_sensor:
            save_data_buffer = [save_pressure[i]["Archivo"], j]  # Agrego nombre del archivo y el nombre del sensor/tiempo.
            save_data_buffer.extend(save_pressure[i][j])   # Agregado de los datos de presion
            # Si el largo de la lista es menor a "max_len" se agregan string vacios "". Todas las listas deben tener
            # la misma lingitud.
            # Nota: Esto permite trasponer los datos en columnas al guardar el CSV sino generaria un error mientras se
            # graba cada linea.
            if len(save_data_buffer) < max_len + 2:  # el 2 es por el agregado del nombre de archivo y el sensor/tiempo.
                save_data_buffer.extend(["" for i in range(max_len+2-len(save_data_buffer))])
            save_data.append(save_data_buffer)
        del(list_sensor, save_data_buffer)
    # Grabado de los datos obtenidos. Se transpone la variable "save_data"
    date_file_name = datetime.datetime.now().strftime("%H-%M-%S_%d-%m-%Y")  # Hora y dia de guardado. Utilizado para guardado de los archivos CSV
    save_file_name = path +'/presiones_{}.csv'.format(date_file_name)
    with open(save_file_name, "w", newline='') as f:
        writer = csv.writer(f, delimiter=seplist)
        # Transposicion de la lista de listados. Conversion de los datos al formato CSV elegido.
        buffer = [[str(line[i]).replace('.', decsep) for line in save_data] for i in range(len(save_data[0]))]
        for line_csv in buffer:
            writer.writerow(line_csv)
    f.close()  # Cerrado del archivo CSV

# Guardado de los datos de las incertidumbres en archivo CSV
def save_csv_incert(save_uncert, conf_level, path, seplist, decsep):
    # Grabado de los datos obtenidos y apertura del archivo a guardar los datos de incertidumbre.
    date_file_name = datetime.datetime.now().strftime("%H-%M-%S_%d-%m-%Y")  # Hora y dia de guardado. Utilizado para guardado de los archivos CSV
    save_file_name = path + '/incertidumbre_{}.csv'.format(date_file_name)
    with open(save_file_name, "w", newline='') as f:
        writer = csv.writer(f, delimiter=seplist)
        for i in range(len(save_uncert)):
            # Determinacion de los encabezados de los datos
            header = [l for l in list(save_uncert[i].keys()) if 'Presion-Sensor' in l]
            header.insert(0, save_uncert[i]['Archivo'])  # Se inserta el nombre de archivo en el primer espacio del encabezado
            # Listado de variables a guardar. Se analiza los keys del primer diccionario unicamente.
            sample_list = [l for l in list(save_uncert[i].keys()) if 'Muestras-' in l]  # Listado de Tomas - Muestras
            sample_list.sort()
            averange_list = [l for l in list(save_uncert[i].keys()) if 'Promedio-' in l]  # Listado de Tomas - Promedio
            averange_list.sort()
            type_a_list = [l for l in list(save_uncert[i].keys()) if 'Tipo A' in l]  # Listado de Tomas - Uexpandida
            type_a_list.sort()
            type_b_list = [l for l in list(save_uncert[i].keys()) if 'Tipo B-presion' in l]  # Listado de Tomas - Uexpandida
            type_b_list.sort()
            comb_uncert_list = [l for l in list(save_uncert[i].keys()) if 'Incertidumbre Combinada' in l]  # Listado de Tomas - Uexpandida
            comb_uncert_list.sort()
            k_list = [l for l in list(save_uncert[i].keys()) if 'Coeficiente Expansion-' in l]  # Listado de Tomas - K
            k_list.sort()
            exp_list = [l for l in list(save_uncert[i].keys()) if 'Uexpandida ' in l]  # Listado de Tomas - Uexpandida
            exp_list.sort()
            distrib_list = [l for l in list(save_uncert[i].keys()) if 'Tipo distribucion-' in l]  # Listado de Tomas - Promedio
            distrib_list.sort()
            # ----------------- Grabado de los datos al CSV -----------------
            buffer = []  # Reinicio de la variable. Guarda temporalmente los datos antes de pasarlo al CSV.
            writer.writerow(header)  # Guardado del encabezado
            # ---Guardado de Muestras---
            buffer = [save_uncert[i][l] for l in sample_list]
            # Convierto los decimales al formato elegido.
            buffer = [str(buffer[i]).replace('.', decsep) for i in range(len(buffer))]
            buffer.insert(0, 'Numero de muestras')
            writer.writerow(buffer)
            # ---Guardado de Promedios---
            buffer = [save_uncert[i][l] for l in averange_list]
            # Convierto los decimales al formato elegido.
            buffer = [str(buffer[i]).replace('.', decsep) for i in range(len(buffer))]
            buffer.insert(0, 'Promedio')
            writer.writerow(buffer)
            # ---Guardado de Incertidumbre Tipo A---
            buffer = [save_uncert[i][l] for l in type_a_list]
            # Convierto los decimales al formato elegido.
            buffer = [str(buffer[i]).replace('.', decsep) for i in range(len(buffer))]
            buffer.insert(0, 'Tipo A')
            writer.writerow(buffer)
            # ---Guardado de Incertidumbre Tipo B---
            buffer = [save_uncert[i][l] for l in type_b_list]
            # Convierto los decimales al formato elegido.
            buffer = [str(buffer[i]).replace('.', decsep) for i in range(len(buffer))]
            buffer.insert(0, 'Tipo B')
            writer.writerow(buffer)
            # ---Guardado de Incertidumbre combinada---
            buffer = [save_uncert[i][l] for l in comb_uncert_list]
            # Convierto los decimales al formato elegido.
            buffer = [str(buffer[i]).replace('.', decsep) for i in range(len(buffer))]
            buffer.insert(0, 'Incertidumbre Combinada')
            writer.writerow(buffer)
            # ---Guardado del Coeficiente de expansion---
            buffer = [save_uncert[i][l] for l in k_list]
            # Convierto los decimales al formato elegido.
            buffer = [str(buffer[i]).replace('.', decsep) for i in range(len(buffer))]
            buffer.insert(0, 'Coeficiente de expansion')
            writer.writerow(buffer)
            # ---Guardado de Uexpandida---
            buffer = [save_uncert[i][l] for l in exp_list]
            # Convierto los decimales al formato elegido.
            buffer = [str(buffer[i]).replace('.', decsep) for i in range(len(buffer))]
            buffer.insert(0, 'Uexpandida ({}%)'.format(conf_level * 100))
            writer.writerow(buffer)
            # ---Guardado del Tipo de distribucion---
            buffer = [save_uncert[i][l] for l in distrib_list]
            # Convierto los decimales al formato elegido.
            buffer = [str(buffer[i]).replace('.', decsep) for i in range(len(buffer))]
            buffer.insert(0, 'Tipo de distribucion')
            writer.writerow(buffer)
            writer.writerow(['##########' for i in range(len(header))])  # Division entre puntos del traverser
        # Escritura de nota de archivos de incertidumbre.
        writer.writerow([])
        writer.writerow(
            ['Importante: El analisis de incertidumbre realizado solo incluye incertidumbre por repetividad (Tipo A)'])
        writer.writerow(
            ['y la incertidumbre debido a la calibracion del instrumento (Tipo B), se debe realizar un analisis'])
        writer.writerow(['de otras fuentes de incertidumbre.'])
    f.close()  # Cerrado del archivo CSV

# Developed by P