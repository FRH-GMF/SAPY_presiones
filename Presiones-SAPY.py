import os
import base64

# Carga imagenes del layout e icono
from image.logo import *
from image.icono import *

# Cargar funciones
from function.functions import *


# -----------Icono-----------
icon_bytes = base64.b64decode(icon)

# -------------Layout-------------
# Theme
sg.theme('SystemDefaultForReal')

# Secciones del Layout
# Armado del frame de la configuracion CSV
frame = [[sg.Radio("Automatico (solo Windows)", "grupo-CSV")],
         [sg.Radio("Separador de listas: COMA - Simbolo decimal: PUNTO", "grupo-CSV", default=True)],
         [sg.Radio("Separador de listas: PUNTO Y COMA - Simbolo decimal: COMA", "grupo-CSV")]]

# Armado del logo junto al nivel de confianza.
cola = [[sg.Image(source=logo, subsample=3, tooltip='Laboratorio de Aerodinamica y Fluidos')]]
colb = [[sg.Text('Nivel de confianza:'),
         sg.Combo(values=['68%', '95%', '99%'], key='-CONF-', default_value=['95%'], s=(5, 1),
                  readonly=True, background_color='white')]]

# Armado del frame del autozero
frama_a = [[sg.Combo(values=[], key='-AUTOZERO-', enable_events=True, expand_x=True)],
           [sg.Checkbox('Informar el resultado del autozero', key='-INFAUTOZERO-')]]

# Coluna izquierda
col1 = [[sg.Text("Ingresar carpeta de trabajo")],
        [sg.Input(key='-FOLDER-', enable_events=True, size=(50, 1)), sg.FolderBrowse(button_text='Buscar')],
        [sg.Frame('Elegir archivo del Autozero', frama_a, vertical_alignment='center',
                  pad=(0, (8, 4)), expand_x=True)],
        [sg.Frame('Formato de salida del CSV', frame, vertical_alignment='center', expand_y=True)],
        [sg.Column(cola), sg.Column(colb, vertical_alignment='top')]]

# Coluna derecha
col2 = [[sg.Text('Seleccione los archivos a procesar')], [sg.Button('Todos', key='-TODOS-', size=(7, 1)),
                                                          sg.Button('Ninguno', key='-NINGUNO-', size=(7, 1))],
        [sg.Listbox(values=['No hay archivos CSV'], key='-FILE LIST-', size=(35, 16), enable_events=True,
                    select_mode=sg.LISTBOX_SELECT_MODE_EXTENDED)],
        [sg.Button('Procesar', key='-PROCESS-', enable_events=True, font=("Arial", 13), size=(9, 1),
                   pad=(10, (10, 5))), sg.Button('Salir', font=("Arial", 13), size=(9, 1), pad=(10, (10, 5))),
         sg.Push()]]

layout = [[sg.Column(col1), sg.Column(col2)]]

# Generacion de 2 ventanas (windows)
# La primera es el programa principal, la segunda es para ventanas de avisos de progreso.
window1 = sg.Window("Procesamiento de presiones – SAPY - Version 1.1", layout, resizable=False, icon=icon_bytes,
                    finalize=True)  # Ventana de Principal
window2 = None  # Ventana de progreso

# Inicializacion de variables. Evita que se generen errores en el loop.
fnames = ['No hay archivos CSV']
vref = []

# -------------Loop de evento-------------
while True:
    # Se utiliza el sistema multi-window.
    window, event, values = sg.read_all_windows()

    # Lectura de la carpeta de trabajo.
    if event == '-FOLDER-':
        # Guardado de la ubicacion de la carpeta de trabajo
        folder = values['-FOLDER-']
        try:  # Existe la carpeta sino devuelve listado vacio de archivos
            file_list = os.listdir(folder)
        except Exception as e:
            print(e)
            file_list = []
        # Busca en la carpeta de trabajo los archivos que sean solo CSV.
        fnames = [f for f in file_list if os.path.isfile(
            os.path.join(folder, f)) and f.lower().endswith(".csv")]
        # Si mo hay archivos CSV en la carpeta devuelve advertencia en el ListBox.
        if not fnames:
            fnames = ['No hay archivos CSV']
        # Actualiza los nombres de los archivos al ListBox y al Combo.
        window['-FILE LIST-'].update(fnames)
        window['-AUTOZERO-'].update(values=fnames)

    # Selecciona todos los archivos del listado.
    if event == '-TODOS-':
        window['-FILE LIST-'].update(set_to_index=[i for i in range(len(fnames))])

    # Deselecciona todos los archivos del listado
    if event == '-NINGUNO-':
        window['-FILE LIST-'].update(set_to_index=[])

    # Boton Procesar
    if event == '-PROCESS-':
        can_process = True  # Flag de que no existen errores.

        # Carga de datos provenientes de la interfaz grafica.
        # Nombre de la carpeta de trabajo
        path_folder = values['-FOLDER-']
        # Nombre del archivo del Autozero referencia
        autozero_file = values['-AUTOZERO-']
        # Formato de salida CSV
        if values[0]:
            option = 0
        elif values[1]:
            option = 1
        else:
            option = 2
        seplist, decsep = formato_csv(option)  # De la opcion seleccionada obtengo el formato definido
        # Nivel de Confianza. Se definieron 68.27%, 95% y 99%. Los que se suelen usar.
        conf_level = values['-CONF-']
        if conf_level == '68%':
            conf_level = float(0.6827)
        elif conf_level == '95%':
            conf_level = float(0.95)
        elif conf_level == '99%':
            conf_level = float(0.99)

        # ------ Comprobacióm de errores ------
        #  La carpeta no existe
        if not os.path.exists(path_folder) and can_process:
            error_popup('La carpeta seleccionada no existe')
            can_process = False
        #  No se seleccionaron archivos. Que no sea vacio ni con el mensaje del listbox.
        if (values['-FILE LIST-'] == [] or values['-FILE LIST-'] == ['No hay archivos CSV']) and can_process:
            error_popup('No se selecciono ningun archivo')
            can_process = False
        if autozero_file == '' and can_process:
            error_popup('No se selecciono el archivo de referencia')
            can_process = False

        # Calculo de los voltajes de referencia.
        if can_process:  # Si no hubo errores se continua.
            try:  # Intenta procesar el archivo sino genera un mensaje de error.
                path = path_folder + '/' + autozero_file
                vref = reference_voltage(path)
                # Listado de los valores de voltaje del autozero. Se activa por el checkbox de la interfaz.
                if values['-INFAUTOZERO-'] and can_process:
                    values_vref = []
                    for toma_volt, value in vref.items():
                        values_vref.append(str(toma_volt) + ': {}'.format(round(value, 4)))
                    autozero_popup('\n'.join(values_vref))
            except Exception as e:
                print(e)
                vref = []  # Evita tomar valores de una instancia anterior
                # Aviso de archivo de Autozero no procesable
                error_popup('El archivo del Autozero no es procesable')
                can_process = False

        # Creacion/verificacion de la carpeta "Resultados".
        save_path_folder = path_folder + '/Resultados'  # Linea para cambiar el nombre de la carpeta de salida
        if not os.path.isdir(save_path_folder):
            try:
                # Prueba generar la cerpeta "Resultados".
                os.mkdir(save_path_folder)
            except Exception as e:
                print(e)
                # Aviso la carpeta de salida no pudo crearse
                error_popup('No se pudo crear la carpeta "Resultados"')
                can_process = False

        # Si no hay errores se prosigue
        if can_process:
            # Armado listado de archivos seleccionados
            file_list = values['-FILE LIST-']
            file_path_list = [path_folder + '/' + i for i in file_list]

            # Procesamiento de los archivos
            save_data = []  # Inicializo variable donde se guardan los datos.
            error_files_list = []  # Inicializo variable donde se guardan los archivos con fallas.
            # Barra de progreso del calculo
            window2 = sg.Window('Procesando', [[sg.Text('Procesando ... 0%', key='-PROGRESS VALUE-')], [
                sg.ProgressBar(len(file_path_list), orientation='horizontal', style='xpnative', size=(20, 20),
                               k='-PROGRESS-')]], finalize=True, icon=icon_bytes, modal=True)
            for i in range(len(file_path_list)):
                data = []  # Reinicio de la variable donde se guardan los datos del CSV.
                # Actualizacion barra de progreso
                window2['-PROGRESS-'].update(current_count=i)
                window2['-PROGRESS VALUE-'].update('Procesando ... {}%'.format(int(((i+1)/len(file_path_list))*100)))
                # Se abre cada archivo seleccionado en la interfaz.
                with open(file_path_list[i]) as csv_file:
                    csv_reader = csv.reader(csv_file, delimiter=';')
                    # Extraccion de todas las filas del archivo CSV
                    for csv_row in csv_reader:
                        data.append(csv_row)
                    try:
                        # Calculo de las presiones y la incertidumbre
                        data_calc = data_process(data, vref, file_list[i], conf_level)
                        # Union de los datos procesados de cada archivo.
                        save_data.append(data_calc)
                    except Exception as e:
                        print(e)
                        # Se agrega el nombre de archivo que no pudo procesarse.
                        error_files_list.append(file_list[i])
            # Se cierra la ventana de progreso
            window2.close()

            # Prevención de error cuando todos los archivos fallan en procesarse.
            if len(save_data) == 0:
                info_popup('No se llego a procesar ningun archivo')
            else:
                # ---------- Guardado de los archivos ----------
                try:
                    # Ventana de aviso de guardado de archivos
                    window2 = sg.Window('', [[sg.Text('Guardando archivos CSV')]], no_titlebar=True,
                                        background_color='grey', finalize=True, modal=True)
                    save_csv_pressure(save_data, save_path_folder, seplist, decsep)
                    save_csv_incert(save_data, conf_level, save_path_folder, seplist, decsep)
                    info_popup('Los archivos de salida se guardaron con exito')
                except Exception as e:
                    print(e)
                    info_popup('Existen problemas en el guardado de los archivos de salida')
                # Se cierra la ventana de aviso de guardado de archivos.
                window2.close()
                
            # Aviso de archivos no procesados
            if error_files_list:
                error_files_popup('\n'.join(error_files_list))

    # Salida del programa
    if event == "Salir" or event == sg.WIN_CLOSED:
        break

window.close()

# Developed by P