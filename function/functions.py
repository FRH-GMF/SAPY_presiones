from statistics import mean, stdev
from subprocess import check_output
from datetime import datetime
from scipy.stats import t
from scipy.stats import norm
import FreeSimpleGUI as sg
import csv
import base64
import os

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
            error_popup(
                '''Fallo el modo automatico, se establecera por defecto separdor de LISTAS "," y el simbolo DECIMAL "."''')
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


# Reorganiza la información recibida de los sensores del SAPY
def pre_process_csv(file_path):
    data_csv = []
    # Paso 1: Extracción de datos del CSV
    with open(file_path, mode='r', newline='', encoding='utf-8') as csvfile:
        csv_reader = csv.reader(csvfile, delimiter=';')
        for csv_row in csv_reader:
            data_csv.append(csv_row)
        # Paso 2: Se eliminan las filas que no tengan información de los sensores de presión.
        # Se elimina el final de archivo "#", si existe.
        if "#" in data_csv[-1]:
            data_csv.pop(-1)
        # Se analiza los encabezados de las primeras 5 filas y se eliminan los que no tengan datos como ">T", ">M"
        # o ">V". Posicion, angulo u otro dato es eliminado
        data_csv_buffer = [sub_list for sub_list in data_csv[:5]
                           if sub_list and (sub_list[0].startswith(">M") or
                                            sub_list[0].startswith(">V") or
                                            sub_list[0].startswith(">T"))]
        data_csv = data_csv_buffer + data_csv[5:]  # Junto las filas filtradas y el resto de los datos

        # Paso 3: Se determina la estructura de datos de ingreso. Por número o por tiempo de muestreo.
        # Se analiza usando el encabezado de filas.
        row_header_summary = [row_header[0] for row_header in data_csv if row_header]
        # Elimino encabezados duplicados
        row_header = list(set(row_header_summary))
        # Analizo encabezado y la estructura definida
        time_structure_flag = False
        sample_structure_flag = False
        for i in row_header:
            if ">T" in i:
                time_structure_flag = True
                break
            if ">M" in i or ">V" in i:
                sample_structure_flag = True
                break

        # ANÁLISIS ESTRUCTURA MUESTREO
        if sample_structure_flag:
            data_dictionary = {}  # Inicio variable de salida tipo diccionario
            for i in data_csv:

                # Paso 4: Se cargan los datos de cada sensor de cada banco al diccionario de salida.
                # Genero el nombre de encabezado
                if ">M" in i[0]:
                    header = i[0].replace(">M", "sensor_") + "_" + i[1]
                elif ">V" in i[0]:
                    header = i[0].replace(">V", "tension_") + "_" + i[1]
                # El ultimo elemento de la fila "<" se elimina con [2:-1]. Los dos primeros elementos son el número
                # de banco y el número de sensor.
                data_dictionary.update({header: [float(row.replace(",", ".")) for row in i[2:-1]]})

        # ANÁLISIS ESTRUCTURA TIEMPO
        if time_structure_flag:
            # Paso 4: Determinar la cantidad de bancos usados. Máximo 5 bancos.
            banc_codes = [row[0] for row in data_csv[:5]]
            # Remuevo los valores duplicados y determino la cantidad de bancos
            number_banc = len(list(set(banc_codes)))

            # Paso 5: Se modifican los encabezados de "tiempo" y "tension" en función del número de banco
            for i in range(number_banc):
                banc_name = data_csv[i][0].replace(">T", "")
                data_csv[i][1] = 'tiempo_' + banc_name
                data_csv[i][2] = 'tension_' + banc_name

            # Paso 6: Desentrelazar los datos de los bancos y guardar en forma de listas.
            #  Se coloca en forma continua los datos de ambos bancos y no en forma entralazada.
            data_list = []
            for i in range(0, len(data_csv), number_banc):
                group = []
                for sublist in data_csv[i:i + number_banc]:
                    group.extend(sublist[1:-1])
                data_list.append(group)

            # Paso 7: Se transponen los datos de manera que cada encabezado es el key del diccionario y la lista
            # adjunta al key son los datos
            headers = data_list[0]  # Extraer los encabezados de datos
            data_dictionary = {header: [float(row[i]) for row in data_list[1:]] for i, header in enumerate(headers)}
        return data_dictionary


# Determinacion del voltaje de referencia de cada sensor del instrumento.
def reference_voltage(path):
    # Paso 1: Pre-procesamiento del archivo csv
    data_dictionary_csv = pre_process_csv(path)
    # Paso 2: Crear diccionario con los voltajes de referencia de los sensores con valor 1.
    zero_voltage = {f"V_{i}_{j}": float(1) for i in range(5) for j in range(1, 13)}
    # Paso 3: Obtengo los valores promedio de voltaje de cada sensor y actualizo "reference_voltage"
    for key, value in data_dictionary_csv.items():
        if 'sensor' in key:
            v_average = round(sum(value) / len(value), 6)  # Reducir el numero de cifras a 6
            # Solo para debugging
            # print(f"{key} = {sum(value)}")
            # print(len(value))
            sensor_name = key.replace('sensor', 'V')
            zero_voltage[sensor_name] = v_average
    return zero_voltage


# Procesamiento de las presiones y la incertidumbre
def data_process(data_csv, volt_reference, filename, nivconf):
    data_out = {}  # Inicialización variable donde se guardan los resultados de cada archivo csv procesado.
    data_out.update({'Archivo': filename})  # Guardado nombre de archivo

    # --------------Procesamiento de las presiones--------------
    # Paso 1: Convierto los datos de voltaje a presiones, segun fabricante del sensor.
    for key, data_list in data_csv.items():

        # Paso 1-a:Se determina el tipo de estructura de datos utilizada. Por tiempo o muestreo
        time_structure_flag = False
        sample_structure_flag = False
        if any("tiempo" in key for key in data_csv):
            time_structure_flag = True
        else:
            sample_structure_flag = True

        # Paso 1-b: Si el key indica ser "tiempo" se guarda la variable como se indica en el key del diccionario.
        if 'tiempo' in key:
            data_out[key] = data_list  # Agrego el tiempo del muestreo

        # Paso 1-c: Si es sensor se procesa todos los voltajes a presiones y se guarda.
        if 'sensor' in key:
            # Numero de banco y sensor
            banc_number = int(key.split('_')[1])  # Se determina el numero de banco del sensor
            sensor_number = key.replace('sensor_', '')

            # Los voltajes de placa varian en función de si la estructura es por muestreo o por tiempo.
            # En caso del tiempo hay una tension única para varios sensores y por muestreo uno por cada sensor.
            if time_structure_flag:
                Vs = data_csv['tension_{}'.format(banc_number)]  # Lista de valores de voltaje de placa del banco
            elif sample_structure_flag:
                Vs = data_csv['tension_{}'.format(sensor_number)]  # Lista de valores de voltaje de placa del banco

            V0 = volt_reference["V_{}".format(sensor_number)]  # Extraigo el voltaje de referencia del sensor.
            Vout = data_list  # Valores de voltaje del sensor.
            # Finalmente se calculan las presiones
            data_pressure = []  # Inicializo la variable donde guardo las presiones.
            for i in range(len(Vout)):
                pressure = (((Vout[i] - V0) / (Vs[i] * 0.2)) * 1000)  # Calculo de presion en Pascales
                pressure = float('%.4f' % pressure)  # Reduccion a 4 cifras.
                data_pressure.append(pressure)
            # Guardado de datos en variable de salida
            data_out.update(
                {"Presion-Sensor_{}".format(sensor_number): data_pressure})  # Agregado de datos de presiones

    # -------------- Calculo de la incertidumbre --------------
    # Se determina el numero de sensores de los keys a partir del diccionario "data_out"
    pressure_list = [k for k in list(data_out.keys()) if 'Presion-Sensor' in k]
    crit = 10  # Criterio de contribucion dominante. Se eligio 10 veces superior.
    for i in pressure_list:
        # Extraigo datos de presiones.
        data_raw = data_out[i]
        # Numero del sensor. Se obtiene del key del diccionario
        numb_probe = i.replace('Presion-Sensor_', '')
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
                k = t.ppf((1 + nivconf) / 2, sample - 1)  # t student doble cola. t.ppf(alfa, gl)
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
                k = norm.ppf((1 + nivconf) / 2)  # Cumple teorema limite central. Distribución Normal.
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
    # Paso 1: Determinar la longitud mas larga de las listas de presiones.
    # Puede existir mediciones con numeros diferentes de muestras.
    max_len = 0
    for i in range(len(save_pressure)):
        # Maxima longitud de save_pressure["Presion-Sensor x"] siendo "Presion-Sensor x" el primer sensor guardado
        # en el diccionario.
        long = len(save_pressure[i][list(save_pressure[i].keys())[1]])
        if max_len < long:
            max_len = long
    # Paso 2: Armado de lista de listas con los datos calculados
    data_csv_output = []
    for i in range(len(save_pressure)):
        file_name = save_pressure[i]['Archivo']  # Guardo nombre de archivo
        for key, data in save_pressure[i].items():
            data_csv_output_buffer = [file_name]  # Se agrega nombre de archivo a cada variable
            if 'tiempo' in key:
                data_csv_output_buffer.extend([key.replace('tiempo', 'Tiempo Muestreo')])  # Se agrega encabezado
                data_csv_output_buffer.extend(data)  # Se agrega los datos de salida
                # Agregan elementos vacios para que la listas tengan todas las mismas longitudes
                data_csv_output_buffer.extend([''] * (max_len + 2 - len(data_csv_output_buffer)))
                data_csv_output.append(data_csv_output_buffer)
            elif 'Presion' in key:
                data_csv_output_buffer.extend([key])  # Se agrega encabezado
                data_csv_output_buffer.extend(data)  # Se agrega los datos de salida
                # Agregan elementos vacios para que la listas tengan todas las mismas longitudes
                data_csv_output_buffer.extend([''] * (max_len + 2 - len(data_csv_output_buffer)))
                data_csv_output.append(data_csv_output_buffer)

        # Paso 3: Se agregan espacios vacios para generar listas de igual longitud
        # Si el largo de la lista es menor a "max_len" se agregan string vacios "". Todas las listas deben tener
        # la misma longitud.
        # Nota: Esto permite trasponer los datos en columnas al guardar el CSV sino generaría un error mientras se
        # graba cada linea.

    # Paso 4: Grabado de los datos obtenidos. Se transpone la variable "save_data"
    date_file_name = datetime.now().strftime(
        "%H-%M-%S_%d-%m-%Y")  # Hora y dia de guardado. Utilizado para guardado de los archivos CSV
    save_file_name = path + '/presiones_{}.csv'.format(date_file_name)
    with open(save_file_name, "w", newline='') as f:
        writer = csv.writer(f, delimiter=seplist)
        # Transposicion de la lista de listados. Conversion de los datos al formato CSV elegido.
        buffer = [[str(line[i]).replace('.', decsep) for line in data_csv_output] for i in
                  range(len(data_csv_output[0]))]
        [[str(row[i]).replace('.', decsep) for row in data_csv_output] for i in range(len(data_csv_output[0]))]
        # Corrige la generacion de ",csv" en el nombre de archivo
        buffer[0] = [line.replace(",csv", ".csv") for line in buffer[0]]
        for line_csv in buffer:
            writer.writerow(line_csv)
    f.close()  # Cerrado del archivo CSV


# Guardado de los datos de las incertidumbres en archivo CSV
def save_csv_incert(save_uncert, conf_level, path, seplist, decsep):
    # Grabado de los datos obtenidos y apertura del archivo a guardar los datos de incertidumbre.
    date_file_name = datetime.now().strftime(
        "%H-%M-%S_%d-%m-%Y")  # Hora y dia de guardado. Utilizado para guardado de los archivos CSV
    save_file_name = path + '/incertidumbre_{}.csv'.format(date_file_name)
    with open(save_file_name, "w", newline='') as f:
        writer = csv.writer(f, delimiter=seplist)
        for i in range(len(save_uncert)):
            # Determinacion de los encabezados de los datos
            header = [l for l in list(save_uncert[i].keys()) if 'Presion-Sensor' in l]
            header.insert(0, save_uncert[i][
                'Archivo'])  # Se inserta el nombre de archivo en el primer espacio del encabezado
            # Listado de variables a guardar. Se analiza los keys del primer diccionario unicamente.
            sample_list = [l for l in list(save_uncert[i].keys()) if 'Muestras-' in l]  # Listado de Tomas - Muestras
            averange_list = [l for l in list(save_uncert[i].keys()) if 'Promedio-' in l]  # Listado de Tomas - Promedio
            type_a_list = [l for l in list(save_uncert[i].keys()) if 'Tipo A' in l]  # Listado de Tomas - Uexpandida
            type_b_list = [l for l in list(save_uncert[i].keys()) if
                           'Tipo B-presion' in l]  # Listado de Tomas - Uexpandida
            comb_uncert_list = [l for l in list(save_uncert[i].keys()) if
                                'Incertidumbre Combinada' in l]  # Listado de Tomas - Uexpandida
            k_list = [l for l in list(save_uncert[i].keys()) if 'Coeficiente Expansion-' in l]  # Listado de Tomas - K
            exp_list = [l for l in list(save_uncert[i].keys()) if 'Uexpandida ' in l]  # Listado de Tomas - Uexpandida
            distrib_list = [l for l in list(save_uncert[i].keys()) if
                            'Tipo distribucion-' in l]  # Listado de Tomas - Promedio
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


# Debug Code
if __name__ == "__main__":

    # Interface grafica de prueba
    layout = [
        [sg.Text("Seleccionar archivo de prueba:"), sg.Input(key="-ARCHIVO-"), sg.FileBrowse()],
        [sg.Button("Aceptar"), sg.Button("Cancelar")]
    ]

    # Create the window
    window = sg.Window("File Selector", layout)

    # Event loop
    while True:
        event, values = window.read()

        if event == sg.WINDOW_CLOSED or event == "Cancelar":
            print("Operacion cancelada.")
            break
        elif event == "Aceptar":
            file_path = values["-ARCHIVO-"]
            if file_path:
                # Prueba Modulo "Voltaje de referencia"
                print("Archivo Seleccionado:", file_path)
                data = pre_process_csv(file_path)
                print(data)
                vref = reference_voltage(file_path)
                print(vref)
                filename = os.path.basename(file_path)
                print(data_process(data, vref, filename, 0.95))
                break
            else:
                sg.popup("Favor de seleccionar un archivo.")

    # Close the window
    window.close()

# Developed by P
