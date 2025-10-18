"""
Contingency Analysis Module

This module contains functions for running contingency analysis,
processing results, and optimizing generators for substations.
"""

import pandas as pd
import numpy as np
from .generators import create_static_generator, update_generator_power, delete_generator, calculate_power_limits, cleanup_existing_generator


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
    if contingency_analysis is None:
        print("Error: No se encontró el objeto de análisis de contingencia")
        return pd.DataFrame()
    
    print("Configurando análisis de contingencia...")
    contingency_analysis.iopt_Linear = 0
    contingency_analysis.loadmax = 200  # Increased from 50 to 200 to capture more lines
    contingency_analysis.vlmin = 0.9
    contingency_analysis.vlmax = 1.1
    contingency_analysis.vmax_step = 5
    
    # Try to configure contingency analysis to analyze all lines
    try:
        # Set to analyze all lines
        contingency_analysis.iopt_auto = 1  # Auto-select all elements
        
        # Also try to set the analysis to include all lines
        contingency_analysis.iopt_auto = 1
        contingency_analysis.iopt_auto = 1  # Ensure auto-select is enabled
        
        # Try to set the analysis to include all transmission lines
        try:
            # Set the analysis to include all transmission lines
            contingency_analysis.iopt_auto = 1
        except:
            pass
            
        # Try to configure the analysis to include all transmission lines
        try:
            # Set the analysis to include all transmission lines
            contingency_analysis.iopt_auto = 1
        except Exception as e:
            pass
            
    except Exception as e:
        pass
    
    # Ensure load flow considers reactive power limits
    try:
        power_flow = app.GetFromStudyCase('ComLdf')
        if power_flow is not None:
            power_flow.iopt_lim = 1  # Consider reactive power limits
    except Exception as e:
        pass

    print("Ejecutando análisis de contingencia...")
    contingency_analysis.Execute()
    
    # Export results
    elmres = app.GetFromStudyCase('Contingency Analysis AC.ElmRes')
    if elmres is None:
        print("Error: No se encontraron resultados de análisis de contingencia")
        return pd.DataFrame()
    
    comres = app.GetFromStudyCase('ComRes')
    if comres is None:
        print("Error: No se encontró el objeto ComRes para exportar")
        return pd.DataFrame()
    
    print("Exportando resultados a CSV...")
    comres.iopt_exp = 6
    comres.iopt_csel = 0
    comres.pResult = elmres
    comres.f_name = r'Resultados.csv'
    
    comres.Execute()
    
    print("Cargando resultados del archivo 'Resultados.csv'.")
    
    # Check if CSV file exists and has content
    import os
    if not os.path.exists('Resultados.csv'):
        print("Error: CSV file does not exist!")
        return pd.DataFrame()
    
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
    print("Procesando resultados de cargabilidad desde Resultados.csv...")
    
    # Use the new CSV parser to get the data
    from .csv_parser import parse_contingency_results
    
    # Parse the CSV file directly
    result_df = parse_contingency_results('Resultados.csv')
    
    if result_df.empty:
        print("⚠️  No hay datos de cargabilidad válidos")
        return pd.DataFrame(columns=['Linea', 'Cargabilidad_Maxima'])
    
    # Convert to the expected format
    line_load_df = result_df[['Linea', 'Max_Loading_Percent']].copy()
    line_load_df.columns = ['Linea', 'Cargabilidad_Maxima']
    
    print(f"Encontradas {len(line_load_df)} líneas con datos de cargabilidad máxima.")
    print("Procesamiento de cargabilidad completado.")
    return line_load_df


def show_iteration_details(df, iteration_num, substation, current_potencia):
    """Show detailed iteration results from CSV data"""
    print(f"\n📊 DETALLES DE ITERACIÓN {iteration_num} - {substation} ({current_potencia} MW)")
    print("="*60)
    
    # Use the new CSV parser to get the data
    from .csv_parser import get_max_loading_summary
    
    # Get the summary from the CSV file
    summary = get_max_loading_summary('Resultados.csv')
    
    if summary['all_lines'].empty:
        print("⚠️  No hay datos de cargabilidad válidos")
        print(f"\n🎯 RESUMEN: Potencia actual = {current_potencia} MW")
        print("Max cargabilidad = 0.00%, Línea crítica = N/A")
        return pd.DataFrame(columns=['Linea', 'Cargabilidad_Maxima'])
    
    # Show top 10 most loaded lines
    print("Top 10 líneas más cargadas:")
    for i, row in summary['all_lines'].head(10).iterrows():
        print(f"  {i+1}. {row['Linea']}: {row['Max_Loading_Percent']:.2f}%")
    
    # Show summary
    print(f"\n🎯 RESUMEN: Potencia actual = {current_potencia} MW")
    print(f"Max cargabilidad = {summary['max_loading']:.2f}%, Línea crítica = {summary['critical_line']}")
    
    # Convert to the expected format
    line_load_df = summary['all_lines'][['Linea', 'Max_Loading_Percent']].copy()
    line_load_df.columns = ['Linea', 'Cargabilidad_Maxima']
    
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
        
        iteration_num = 0
        while True:
            iteration_num += 1
            print(f"\n--- Iteración {iteration_num} para {substation} ---")
            
            # Run power flow first to get actual power values
            print("Ejecutando flujo de potencia para verificar sistema...")
            power_flow = app.GetFromStudyCase('ComLdf')
            if power_flow is not None:
                power_flow.iopt_lim = 1  # Consider reactive power limits
            power_flow.Execute()
            
            # Get actual power values from the generator after load flow
            bus_terminal = static_generator.term
            actual_p_gen = static_generator.GetAttribute('m:P:bus1')
            actual_q_gen = static_generator.GetAttribute('m:Q:bus1')
            actual_bus_voltage = bus_terminal.GetAttribute('m:u')
            
            print(f"📊 Estado actual: P = {actual_p_gen:.2f} MW, Q = {actual_q_gen:.2f} MVar, V = {actual_bus_voltage:.4f} pu")
            
            # Check voltage limits
            if actual_bus_voltage < 0.9 or actual_bus_voltage > 1.1:
                print(f"⚠️  Voltaje fuera de límites: {actual_bus_voltage:.4f} pu (límites: 0.9 - 1.1 pu)")
                print(f"🎯 Límite de voltaje alcanzado: Potencia máxima segura = {current_potencia - 1} MW")
                
                # Save result
                results.append({
                    'Subestacion': substation,
                    'Potencia_Maxima': current_potencia - 1,
                    'Linea_Critica': 'Voltage Limit',
                    'Cargabilidad_Maxima': f'V={actual_bus_voltage:.4f}'
                })
                
                # Delete generator and cubicle
                print(f"Eliminando generador estático y cubículo en la subestación '{substation}'.")
                delete_generator(static_generator, cubicle)
                break
            
            # Run contingency analysis and get results
            print("Ejecutando análisis de contingencia...")
            df = run_contingency_analysis(app)
            
            # Show detailed iteration results
            line_load_df = show_iteration_details(df, iteration_num, substation, current_potencia)
            
            # Check if any line exceeds the loading threshold
            if not line_load_df.empty:
                max_line_load = line_load_df['Cargabilidad_Maxima'].max()
                max_line = line_load_df[line_load_df['Cargabilidad_Maxima'] == max_line_load]['Linea'].values[0]
            else:
                print("⚠️  No hay datos de cargabilidad válidos")
                max_line_load = 0
                max_line = "N/A"
            
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
                continue
            
            last_max_line_load = max_line_load
            
            # Check if loading threshold is exceeded
            if max_line_load > max_cargabilidad:
                print(f"Subestación {substation}: Potencia máxima segura = {current_potencia - 1} MW, Línea crítica = {max_line}")
                
                # Save result
                results.append({
                    'Subestacion': substation,
                    'Potencia_Maxima': current_potencia - 1,
                    'Linea_Critica': max_line,
                    'Cargabilidad_Maxima': last_max_line_load
                })
                
                # Delete generator and cubicle
                print(f"Eliminando generador estático y cubículo en la subestación '{substation}'.")
                delete_generator(static_generator, cubicle)
                break
            
            # Increase generator power
            current_potencia += 1
            update_generator_power(static_generator, current_potencia, factor_potencia)
            
            print(f"🔄 Aumentando potencia a {current_potencia} MW...")
    
    # Convert results to DataFrame
    df_results = pd.DataFrame(results)
    print("Resultados finales:")
    print(df_results)
    
    return df_results
