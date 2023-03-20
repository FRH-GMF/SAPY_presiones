from statistics import mean, stdev
from subprocess import check_output
import scipy.stats as stats
import PySimpleGUI as sg
import csv


# -------------------------Mensajes popup-------------------------
def error_popup(message):
    sg.Window('Error', [[sg.T('{}'.format(message))],
                        [sg.B('OK', bind_return_key=True, size=(4, 1))]],
              element_justification='c').read(close=True)


def info_popup(message):
    sg.Window('Informacion', [[sg.T('{}'.format(message))],
                              [sg.B('OK', bind_return_key=True, size=(4, 1))]],
              element_justification='c').read(close=True)


def error_files_popup(files):
    sg.Window('Error', [[sg.T('Los siguientes archivos no pudieron procesarse:')],
                        [sg.Multiline(default_text=files, write_only=True, expand_x=True, expand_y=True,
                                      size=(30, 10))],
                        [sg.Push(), sg.B('OK', bind_return_key=True, size=(4, 1))]], resizable=True).read(close=True)


def autozero_popup(values):
    sg.Window('Resultados Autozero', [[sg.T('Se lista los valores de tension utilizados para el Autozero:')],
                        [sg.Multiline(default_text=values, write_only=True, expand_x=True, expand_y=True,
                                      size=(30, 10))],
                        [sg.Push(), sg.B('OK', bind_return_key=True, size=(4, 1))]], resizable=True).read(close=True)


# -------------------------Funciones de procesamiento-------------------------
def formato_csv(option):
    # La opcion "0" lee el formato de "separacion de listas" y el "simbolo decimal" del registro de windows.
    if option == 0:
        # Envia el comando al CMD y luego se aisla el valor del parametro.
        salida = check_output('Reg Query "HKEY_CURRENT_USER\Control Panel\International" /v sList',
                              shell=True)
        salida = salida.decode("utf-8").split("\n")
        seplist = salida[2].replace('    sList    REG_SZ    ', '').replace('\r', '')
        salida = check_output('Reg Query "HKEY_CURRENT_USER\Control Panel\International" /v sDecimal',
                              shell=True)
        salida = salida.decode("utf-8").split("\n")
        decsep = salida[2].replace('    sDecimal    REG_SZ    ', '').replace('\r', '')
        error_popup('Según el registro del sistema el separdor de LISTAS es "{}"'.format(seplist) +
                    ' y el simbolo DECIMAL es "{}"'.format(seplist))
    elif option == 1:
        seplist = ','
        decsep = '.'
    elif option == 2:
        seplist = ';'
        decsep = ','
    else:
        sg.popup('Algo raro paso en la eleccion del formato de salida del CSV', title='Error',
                 keep_on_top=True)
        seplist = ''
        decsep = ''
    return seplist, decsep


# Determinacion del voltaje de referencia de cada toma del instrumento.
def reference_voltage(path):
    # Diccionario por defecto de los voltajes de referencia de las tomas. Maximo SAPY 32 tomas.
    vref = {}
    for i in range(1, 33):
        vref["V{}".format(i)] = 1
    with open(path) as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=';')
        # Extraigo todas las filas
        data_row = []  # Incializacion variable donde se guardan los datos en bruto del CSV.
        for csv_row in csv_reader:
            data_row.append(csv_row)
    data_row.pop(-1)  # Se elimina ultima fila con el caracter #
    # Calculo el valor promedio de voltaje de cada toma del archivo seleccionado.
    for i in range(int(len(data_row) / 2)):
        numbsenor = data_row[2 * i][1]  # Se utiliza la estrategia de que los valores vienen en pares
        values = data_row[i * 2][2:-1]  # Se extraen los valores de los voltajes.
        values = [float(i.replace(',', '.')) for i in values]  # Convierto valores de la lista a flotacion.
        averang = sum(values) / len(values)  # Calculo el promedio de Vout.
        averang = float('%.6f' % averang)  # Reducir el numero de cifras a 6
        vref["V{}".format(numbsenor)] = averang  # Modifico valor del diccionario para la toma especifica.
    return vref


# Se preprocesan los datos inicialmente y luego se determinan las presiones y las incertidumbres
def data_process(data_csv, vref, filename, nivconf):
    # --------------Procesamiento de los datos en bruto--------------
    data = []  # Inicializacion variable de guardado de datos del csv procesado
    data_csv.pop(-1)  # Se elimina ultima fila con el caracter #
    for i in range(len(data_csv)):
        data_buffer = []  # Reinicio variable de guardado
        # Procesamiento de los datos de la fila extraida.
        data_csv[i].pop(-1)  # Se elimina el ultimo elemento con el caracter <
        data_csv[i].pop(0)  # Se elimina el primer elemento con el caracter M o V
        for j in range(len(data_csv[i])):
            data_buffer.append(float(data_csv[i][j].replace(',', '.')))  # Conversion de string a float.
        data.append(data_buffer)
    # --------------Procesamiento de las presiones--------------
    data_pressure = []  # Inicializo la variable de salida.
    for i in range(int(len(data) / 2)):  # Se utiliza la estrategia de que los valores de las tomas vienen en pares
        # Valor de referencia del sensor analizado.
        numbsenor = int(data[2 * i][0])
        V0 = vref["V{}".format(numbsenor)]  # Extraigo del diccionario el valor de referencia.
        # Calculo las presiones para la toma indicada.
        # Inicializo la variable donde guardo las presiones junto con el nombre de archivo y el numero de toma.
        pressure = [filename, "Toma {}".format(numbsenor)]
        for j in range(1, len(data[0])):
            Vout = data[i * 2][j]
            Vs = data[i * 2 + 1][j]
            value = (((Vout - V0) / (Vs * 0.2)) * 1000)
            value = float('%.4f' % value)  # Reduccion de cifras
            pressure.append(value)
        data_pressure.append(pressure)
    # --------------Calculo de la incertidumbre--------------
    data_uncert = []  # Inicializo la variable de salida.
    crit = 10  # Criterio de contribucion dominante. Se eligio 10 veces superior.
    for i in range(len(data_pressure)):
        # Genero encabezado de los datos
        uncert = data_pressure[i][0:2]
        # Extraigo datos numericos.
        data_raw = data_pressure[i][2:]
        # Calculo de incertidumbre.
        sample = len(data_raw)  # Numero de muestras.
        averange = mean(data_raw)  # Estimado de la medicion.
        if sample > 1:
            typea = stdev(data_raw) / (sample ** 0.5)  # Desviación típica experimental.
            # Se diferencia el calculo del componente Tipo B para los diferentes sensores.
            typeb = averange * 0.015 / (3 ** 0.5)  # Componente Tipo B debido a la calibración del sensor de presion.
            ucomb = (typea ** 2 + typeb ** 2) ** 0.5  # Incertidumbre combinada.
            # Analisis de la contribucion dominante para la determinacion de la incertidumbre expandida.
            # El siguiente codigo evita la division por cero.
            try:
                rel_tipe = typea / typeb
            except Exception as e:
                print(e)
                rel_tipe = 1e10
            # Analisis del tipo de distribucion
            if rel_tipe > crit:
                k = stats.t.ppf((1 + nivconf) / 2, sample - 1)  # t student doble cola. t.ppf(alfa, gl)
                distrib = 't-student con {} GL'.format(sample - 1)
            elif typea / typeb < 1 / crit:
                k = (3 ** 0.5) * nivconf  # Distribucion rectangular k=raiz(3)*p
                distrib = 'Rectangular'
            else:
                k = stats.norm.ppf((1 + nivconf) / 2)  # Cumple teorema limite central. Distribución Normal.
                distrib = 'Normal TCLimite'
            uexpand = k * ucomb  # Incertidumbre expandida
            # Reduccion de cifras
            averange = float('%.2f' % averange)
            uexpand = float('%.4f' % uexpand)
            k = float('%.4f' % k)
        else:
            # En mediciones de un solo valor no es posible calcular la incertidumbre
            averange = float('%.2f' % averange)
            uexpand = 'N/A'
            distrib = 'N/A'
            k = 'N/A'
        # Organizo los datos calculados
        uncert.extend([averange, uexpand, sample, distrib, k])
        data_uncert.append(uncert)
    return data_pressure, data_uncert


# -------------------------Guardados de archivos CSV-------------------------
def save_csv_pressure(save_pressure, path, seplist, decsep):
    # Determinacion de la longitud mas larga de las listas de valores.
    # Puede existir mediciones con numeros diferentes de muestras y es necesario determinar la mayor.
    long = 0
    for i in range(len(save_pressure)):
        if long < len(save_pressure[i]):
            long = len(save_pressure[i])
    # Grabado de los datos obtenidos.
    with open(path + '/presiones.csv', "w", newline='') as f:
        writer = csv.writer(f, delimiter=seplist)
        for i in range(long):
            buffer = []  # Reinicio de la variable. Guarda temporalmente los datos antes de pasarlo al CSV.
            for j in range(len(save_pressure)):
                try:
                    buffer.append(save_pressure[j][i])  # transpone los datos de las listas.
                except Exception as e:
                    print(e)
                    buffer.append('')  # Cuando las longitudes de datos son diferentes se agregan espacios vacios.
            if i > 1:  # Evita que se hagan reemplazos en los encabezados de los archivos.
                # Convierto los decimales al formato elegido.
                buffer = [str(buffer[i]).replace('.', decsep) for i in range(len(buffer))]
            writer.writerow(buffer)
        f.close()  # Cerrado del archivo CSV


def save_csv_incert(save_uncert, conf_level, path, seplist, decsep):
    with open(path + '/incertidumbre.csv', "w", newline='') as f:
        writer = csv.writer(f, delimiter=seplist)
        # Encabezado de las filas.
        column = ['', '', 'Promedio', 'Uexpandida ({}%)'.format(conf_level * 100), 'Numero de muestras',
                  'Distribución final', 'Coeficiente de expansion']
        for i in range(len(save_uncert[0])):
            buffer = []  # Reinicio de la variable. Guarda temporalmente los datos antes de pasarlo al CSV.
            for j in range(len(save_uncert)):
                buffer.append(save_uncert[j][i])  # transpone los datos de las listas.
            if i > 1:  # Evita que se hagan reemplazos en los encabezados de los archivos.
                # Convierto los decimales al formato elegido.
                buffer = [str(buffer[i]).replace('.', decsep) for i in range(len(buffer))]
            buffer.insert(0, column[i])  # Inserto el encabezado de las filas.
            writer.writerow(buffer)
        # Escritura de nota de archivos de incertidumbre.
        writer.writerow([])
        writer.writerow(
            ['Importante: El analisis de incertidumbre realizado solo incluye incertidumbre por repetividad (Tipo A)'])
        writer.writerow(
            ['y la incertidumbre debido a la calibracion del instrumento (Tipo B), se debe realizar un analisis'])
        writer.writerow(['de otras fuentes de incertidumbre.'])
        f.close()  # Cerrado del archivo CSV
