"""
Contingency Analysis Module

This module contains functions for running contingency analysis,
processing results, and optimizing generators for substations.
"""

import pandas as pd
from .generators import create_static_generator, update_generator_power, delete_generator, calculate_power_limits


def run_contingency_analysis(app):
    """
    Run N-1 contingency analysis in PowerFactory and export results to CSV.
    
    Args:
        app: PowerFactory application object
        
    Returns:
        pandas.DataFrame: Contingency analysis results
        
    Example:
        >>> df = run_contingency_analysis(app)
    """
    print("Iniciando análisis de contingencia N-1.")
    app.ClearOutputWindow()
    
    # Configure contingency analysis
    contingency_analysis = app.GetFromStudyCase('*.ComSimoutage')
    contingency_analysis.iopt_Linear = 0
    contingency_analysis.loadmax = 50
    contingency_analysis.vlmin = 0.9
    contingency_analysis.vlmax = 1.1
    contingency_analysis.vmax_step = 5
    contingency_analysis.Execute()
    
    # Export results
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
    """
    Process contingency analysis results to extract line loading information.
    
    Args:
        df (pandas.DataFrame): Raw contingency analysis results
        
    Returns:
        pandas.DataFrame: Processed line loading data sorted by maximum loading
        
    Example:
        >>> line_load_df = process_cargabilidad(df)
    """
    print("Procesando resultados de cargabilidad.")
    
    line_names = df.columns
    last_row = df.iloc[-1]
    line_load_df = pd.DataFrame({'Linea': line_names, 'Cargabilidad_Maxima': last_row})
    
    # Clean data
    line_load_df = line_load_df[line_load_df['Cargabilidad_Maxima'] != '----']
    line_load_df['Cargabilidad_Maxima'] = pd.to_numeric(line_load_df['Cargabilidad_Maxima'], errors='coerce')
    line_load_df = line_load_df.dropna(subset=['Cargabilidad_Maxima'])
    line_load_df = line_load_df[~line_load_df['Linea'].str.contains('Study Cases', case=False)]
    
    # Clean line names
    line_load_df['Linea'] = line_load_df['Linea'].apply(lambda x: x.split('\\')[-1])
    line_load_df['Linea'] = line_load_df['Linea'].str.replace('.ElmLne', '', regex=False)
    line_load_df = line_load_df[~line_load_df['Linea'].str.contains('69 kV|34.5 kV|4.16 kV', case=False)]
    
    # Sort by maximum loading
    line_load_df = line_load_df.sort_values(by='Cargabilidad_Maxima', ascending=False).reset_index(drop=True)
    
    print("Procesamiento de cargabilidad completado.")
    return line_load_df


def optimize_generators_for_substations(app, substations, network_data, hoja, initial_potencia=1, factor_potencia=0.95, max_cargabilidad=110, threshold_inconvergence=10):
    """
    Optimize generators for each substation by finding the maximum safe power
    that doesn't exceed the specified loading threshold.
    
    Args:
        app: PowerFactory application object
        substations (list): List of substation names to optimize
        network_data: PowerFactory network data folder object
        hoja (str): Name of the sheet/folder for generator creation
        initial_potencia (float): Initial power in MW (default: 1)
        factor_potencia (float): Power factor (default: 0.95)
        max_cargabilidad (float): Maximum allowed loading percentage (default: 110)
        threshold_inconvergence (float): Threshold for detecting convergence issues (default: 10)
        
    Returns:
        pandas.DataFrame: Results with maximum safe power for each substation
        
    Example:
        >>> results = optimize_generators_for_substations(app, substations, network_data, 'Grid')
    """
    results = []  # List to store results
    
    for substation in substations:
        current_potencia = initial_potencia
        print(f"Optimizando generador para la subestación '{substation}'.")
        
        # Create initial generator
        bus_voltage, p_gen, q_gen, static_generator, cubicle = create_static_generator(
            app, network_data, hoja, substation, current_potencia, factor_potencia
        )
        
        if static_generator is None:
            print(f"Error: No se pudo crear el generador en la subestación '{substation}'.")
            continue
        
        last_max_line_load = None
        
        while True:
            # Run contingency analysis and get results
            print("Ejecutando análisis de contingencia.")
            df = run_contingency_analysis(app)
            line_load_df = process_cargabilidad(df)
            
            # Check if any line exceeds the loading threshold
            max_line_load = line_load_df['Cargabilidad_Maxima'].max()
            max_line = line_load_df[line_load_df['Cargabilidad_Maxima'] == max_line_load]['Linea'].values[0]
            
            print(f"Subestación {substation}: Potencia actual = {current_potencia} MW, Max cargabilidad = {max_line_load}%, Línea = {max_line}")
            
            # Detect convergence issues
            if (last_max_line_load is not None and 
                (max_line_load - last_max_line_load) > threshold_inconvergence and 
                max_line_load > max_cargabilidad):
                
                print(f"Advertencia: Inconvergencia detectada en la subestación '{substation}'. "
                      f"La cargabilidad saltó más del {threshold_inconvergence}% y superó el {max_cargabilidad}%. "
                      f"Aumentando potencia y volviendo a intentar.")
                
                # Increase power and continue
                current_potencia += 1
                update_generator_power(static_generator, current_potencia, factor_potencia)
                
                # Run power flow
                power_flow = app.GetFromStudyCase('ComLdf')
                power_flow.Execute()
                continue
            
            last_max_line_load = max_line_load
            
            # Check if loading threshold is exceeded
            if max_line_load > max_cargabilidad:
                print(f"Subestación {substation}: Potencia máxima segura = {current_potencia - 1} MW, Línea crítica = {max_line}")
                
                # Save result
                results.append({
                    'Subestacion': substation,
                    'Potencia Maxima': current_potencia - 1,
                    'Linea Critica': max_line,
                    'Cargabilidad Maxima': last_max_line_load
                })
                
                # Delete generator and cubicle
                print(f"Eliminando generador estático y cubículo en la subestación '{substation}'.")
                delete_generator(static_generator, cubicle)
                break
            
            # Increase generator power
            current_potencia += 1
            update_generator_power(static_generator, current_potencia, factor_potencia)
            
            # Run power flow to apply changes
            power_flow = app.GetFromStudyCase('ComLdf')
            power_flow.Execute()
            
            print(f"Subestación {substation}: Aumentando potencia a {current_potencia} MW.")
    
    # Convert results to DataFrame
    df_results = pd.DataFrame(results)
    print("Resultados finales:")
    print(df_results)
    
    return df_results
