import os
import sys
import math
import pandas as pd


def pf_enviroment(dig_path):
    sys.path.append(dig_path)
    os.environ['PATH'] += f';{dig_path}'
    print(f"PowerFactory environment initialized with path: {dig_path}")


def activate_project(app, project_name):
    project = app.ActivateProject(project_name)
    if project == 0:
        print(f"Proyecto '{project_name}' activado con éxito.")
    else:
        print(f"No se pudo activar el proyecto '{project_name}'. Verifica si el nombre es correcto.")
    return project


def list_and_select_study_case(app, study_case_name: str):
    study_case_fldr = app.GetProjectFolder('study')
    study_cases = study_case_fldr.GetContents('*.Intcase', 0)
    print('List of study cases:')
    for case in study_cases:
        print(case.loc_name)
    selected_case = next((case for case in study_cases if case.loc_name == study_case_name), None)
    if selected_case:
        selected_case.Activate()
        print(f"Study case '{study_case_name}' activated.")
    else:
        print(f"Study case '{study_case_name}' not found.")
    return app.GetActiveStudyCase()


def list_and_activate_operation_scenario(app, scenario_name: str):
    operation_scenarios_folder = app.GetProjectFolder('scen')
    print('List of operation scenarios:')
    for scenario in operation_scenarios_folder.GetChildren(1):
        print(scenario.loc_name)
    selected_scenario = next((s for s in operation_scenarios_folder.GetChildren(1) if s.loc_name == scenario_name), None)
    if selected_scenario:
        selected_scenario.Activate()
        print(f"Operation scenario '{scenario_name}' activated.")
    else:
        print(f"Operation scenario '{scenario_name}' not found.")
    return selected_scenario


def create_static_generator(app, network_data, hoja_name, barra_name, potencia_activa, factor_potencia):
    print(f"Creando generador estático en la barra '{barra_name}' con potencia activa {potencia_activa} MW y factor de potencia {factor_potencia}.")
    potencia_aparente = potencia_activa / factor_potencia
    potencia_reactiva_max = math.sqrt(potencia_aparente*2 - potencia_activa*2)
    potencia_reactiva_min = -potencia_reactiva_max
    print(f"Buscando hoja '{hoja_name}' en 'Network Data'.")
    hoja = network_data.GetContents(hoja_name, 1)[0]
    if not hoja:
        print(f"No se encontró la hoja '{hoja_name}' en 'Network Data'.")
        return None, None, None, None
    bus = None
    for b in app.GetCalcRelevantObjects('*.ElmTerm'):
        if b.loc_name == barra_name:
            bus = b
            break
    if bus:
        print(f"Barra '{barra_name}' encontrada. Creando cubículo y generador estático.")
        cubicle = bus.CreateObject('StaCubic', 'Cubicle_Generador')
        cubicle.bus1 = bus
        if cubicle:
            print(f"Cubículo creado en la barra '{barra_name}' y conectado a la barra.")
        else:
            print(f"Error al crear el cubículo en la barra '{barra_name}'.")
        switcher = cubicle.CreateObject('StaSwitch', 'Switcher_Generador')
        switcher.on_off = 1
        static_generator = hoja.CreateObject('ElmGenstat', 'Generador_Estatico')
        static_generator.SetAttribute('sgn', potencia_aparente)
        static_generator.SetAttribute('e:pgini', potencia_activa)
        static_generator.SetAttribute('cosn', factor_potencia)
        static_generator.SetAttribute('av_mode', 'constv')
        static_generator.term = cubicle
        static_generator.SetAttribute('usetp', 1)
        static_generator.SetAttribute('cQ_max', potencia_reactiva_max)
        static_generator.SetAttribute('cQ_min', potencia_reactiva_min)
        cubicle.obj_id = static_generator
        print(f"Ejecutando flujo de potencia para la barra '{barra_name}'.")
        power_flow = app.GetFromStudyCase('ComLdf')
        power_flow.Execute()
        bus_voltage = bus.GetAttribute('m:u')
        potencia_activa_generada = static_generator.GetAttribute('m:P:bus1')
        potencia_reactiva_generada = static_generator.GetAttribute('m:Q:bus1')
        print(f"Generador estático creado: Voltaje barra = {bus_voltage}, P generada = {potencia_activa_generada}, Q generada = {potencia_reactiva_generada}")
        return bus_voltage, potencia_activa_generada, potencia_reactiva_generada, static_generator, cubicle
    else:
        print(f"Barra '{barra_name}' no encontrada.")
        return None, None, None, None


def run_contingency_analysis(app):
    print("Iniciando análisis de contingencia N-1.")
    app.ClearOutputWindow()
    contingency_analysis = app.GetFromStudyCase('*.ComSimoutage')
    contingency_analysis.iopt_Linear = 0
    contingency_analysis.loadmax = 50
    contingency_analysis.vlmin = 0.9
    contingency_analysis.vlmax = 1.1
    contingency_analysis.vmax_step = 5
    contingency_analysis.Execute()
    elmres = app.GetFromStudyCase('Contingency Analysis AC.ElmRes')
    comres = app.GetFromStudyCase('ComRes')
    comres.iopt_exp = 6
    comres.iopt_csel = 0
    comres.pResult = elmres
    comres.f_name = r'Resultados.csv'
    comres.Execute()
    print("Cargando resultados del archivo 'Resultados.csv'.")
    df = pd.read_csv('Resultados.csv', encoding='latin1', low_memory=False)
    return df


def process_cargabilidad(df):
    print("Procesando resultados de cargabilidad.")
    line_names = df.columns
    last_row = df.iloc[-1]
    line_load_df = pd.DataFrame({'Linea': line_names, 'Cargabilidad_Maxima': last_row})
    line_load_df = line_load_df[line_load_df['Cargabilidad_Maxima'] != '----']
    line_load_df['Cargabilidad_Maxima'] = pd.to_numeric(line_load_df['Cargabilidad_Maxima'], errors='coerce')
    line_load_df = line_load_df.dropna(subset=['Cargabilidad_Maxima'])
    line_load_df = line_load_df[~line_load_df['Linea'].str.contains('Study Cases', case=False)]
    line_load_df['Linea'] = line_load_df['Linea'].apply(lambda x: x.split('\\')[-1])
    line_load_df['Linea'] = line_load_df['Linea'].str.replace('.ElmLne', '', regex=False)
    line_load_df = line_load_df[~line_load_df['Linea'].str.contains('69 kV|34.5 kV|4.16 kV', case=False)]
    line_load_df = line_load_df.sort_values(by='Cargabilidad_Maxima', ascending=False).reset_index(drop=True)
    print("Procesamiento de cargabilidad completado.")
    return line_load_df


def optimize_generators_for_substations(app, substations, network_data, hoja, initial_potencia=1, factor_potencia=0.95, max_cargabilidad=110, threshold_inconvergence=10):
    """
    Itera sobre cada subestación, agregando un generador, corriendo el análisis de contingencia y verificando el 
    máximo de potencia que puede soportar sin exceder el 110% de cargabilidad en ninguna línea.
    """
    results = []  # Lista para almacenar los resultados

    for substation in substations:
        current_potencia = initial_potencia
        print(f"Optimizando generador para la subestación '{substation}'.")

        # Crear generador inicial
        bus_voltage, p_gen, q_gen, static_generator, cubicle = create_static_generator(app, network_data, hoja, substation, current_potencia, factor_potencia)
        if static_generator is None:
            print(f"Error: No se pudo crear el generador en la subestación '{substation}'.")
            continue

        last_max_line_load = None

        while True:
            # Ejecutar el análisis de contingencia N-1 y obtener los resultados
            print("Ejecutando análisis de contingencia.")
            df = run_contingency_analysis(app)
            line_load_df = process_cargabilidad(df)

            # Verificar si alguna línea excede el 110% de cargabilidad
            max_line_load = line_load_df['Cargabilidad_Maxima'].max()
            max_line = line_load_df[line_load_df['Cargabilidad_Maxima'] == max_line_load]['Linea'].values[0]

            print(f"Subestación {substation}: Potencia actual = {current_potencia} MW, Max cargabilidad = {max_line_load}%, Línea = {max_line}")

            # Detectar inconvergencia si el incremento supera el umbral definido y si la cargabilidad es mayor a 110%
            if last_max_line_load is not None and (max_line_load - last_max_line_load) > threshold_inconvergence and max_line_load > max_cargabilidad:
                print(f"Advertencia: Inconvergencia detectada en la subestación '{substation}'. La cargabilidad saltó más del {threshold_inconvergence}% y superó el 110%. Aumentando potencia y volviendo a intentar.")
                # Aumentar la potencia en 1 MW y continuar
                current_potencia += 1
                potencia_aparente = current_potencia / factor_potencia
                potencia_reactiva_max = math.sqrt(potencia_aparente*2 - current_potencia*2)
                potencia_reactiva_min = -potencia_reactiva_max
                static_generator.SetAttribute('sgn', potencia_aparente)
                static_generator.SetAttribute('e:pgini', current_potencia)
                static_generator.SetAttribute('cQ_max', potencia_reactiva_max)
                static_generator.SetAttribute('cQ_min', potencia_reactiva_min)

                power_flow = app.GetFromStudyCase('ComLdf')
                power_flow.Execute()

                continue  # Seguir con la próxima iteración con la nueva potencia

            last_max_line_load = max_line_load

            if max_line_load > max_cargabilidad:
                print(f"Subestación {substation}: Potencia máxima segura = {current_potencia - 1} MW, Línea crítica = {max_line}")

                # Guardar el resultado en la lista
                results.append({
                    'Subestacion': substation,
                    'Potencia Maxima': current_potencia - 1,
                    'Linea Critica': max_line,
                    'Cargabilidad Maxima': last_max_line_load
                })

                # Eliminar el generador estático y el cubículo
                print(f"Eliminando generador estático y cubículo en la subestación '{substation}'.")
                static_generator.Delete()
                cubicle.Delete()

                break

            # Aumentar la potencia del generador estático existente
            current_potencia += 1
            potencia_aparente = current_potencia / factor_potencia
            potencia_reactiva_max = math.sqrt(potencia_aparente*2 - current_potencia*2)
            potencia_reactiva_min = -potencia_reactiva_max
            static_generator.SetAttribute('sgn', potencia_aparente)
            static_generator.SetAttribute('e:pgini', current_potencia)
            static_generator.SetAttribute('cQ_max', potencia_reactiva_max)
            static_generator.SetAttribute('cQ_min', potencia_reactiva_min)

            # Ejecutar el flujo de potencia nuevamente para aplicar los cambios
            power_flow = app.GetFromStudyCase('ComLdf')
            power_flow.Execute()

            print(f"Subestación {substation}: Aumentando potencia a {current_potencia} MW.")

    # Convertir los resultados a un DataFrame
    df_results = pd.DataFrame(results)
    print("Resultados finales:")
    print(df_results)

    return df_results 