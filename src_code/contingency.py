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
    
    # Debug: Check contingency analysis configuration
    print(f"Debug: Contingency analysis loadmax: {contingency_analysis.loadmax}")
    print(f"Debug: Contingency analysis vlmin: {contingency_analysis.vlmin}")
    print(f"Debug: Contingency analysis vlmax: {contingency_analysis.vlmax}")
    
    # Try to configure contingency analysis to analyze all lines
    try:
        # Set to analyze all lines
        contingency_analysis.iopt_auto = 1  # Auto-select all elements
        print("Debug: Set contingency analysis to auto-select all elements")
        
        # Also try to set the analysis to include all lines
        contingency_analysis.iopt_auto = 1
        contingency_analysis.iopt_auto = 1  # Ensure auto-select is enabled
        
        # Try to set the analysis to include all transmission lines
        try:
            contingency_analysis.iopt_auto = 1
            print("Debug: Auto-select enabled")
        except:
            print("Debug: Could not enable auto-select")
            
        # Try to configure the analysis to include all transmission lines
        try:
            # Set the analysis to include all transmission lines
            contingency_analysis.iopt_auto = 1
            print("Debug: Configured to analyze all transmission lines")
        except Exception as e:
            print(f"Debug: Could not configure transmission lines: {e}")
            
    except Exception as e:
        print(f"Debug: Could not set auto-select option: {e}")
    
    print("Ejecutando análisis de contingencia...")
    contingency_analysis.Execute()
    
    # Debug: Check if contingencies were analyzed
    print("Debug: Checking contingency analysis results...")
    print(f"Debug: Contingency analysis completed")
    
    # Try to get information about what was analyzed
    try:
        # Check if there are any contingencies configured
        print(f"Debug: Checking contingency analysis configuration...")
        print(f"Debug: Contingency analysis object: {contingency_analysis}")
        
        # Try to get the number of contingencies
        try:
            num_contingencies = contingency_analysis.GetNumberOfContingencies()
            print(f"Debug: Number of contingencies: {num_contingencies}")
        except:
            print("Debug: Could not get number of contingencies")
            
    except Exception as e:
        print(f"Debug: Error checking contingency analysis: {e}")
    
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
    
    # Debug: Check export configuration
    print(f"Debug: Export format: {comres.iopt_exp}")
    print(f"Debug: Export selection: {comres.iopt_csel}")
    print(f"Debug: Export filename: {comres.f_name}")
    
    comres.Execute()
    
    print("Cargando resultados del archivo 'Resultados.csv'.")
    
    # Debug: Check if CSV file exists and has content
    import os
    if os.path.exists('Resultados.csv'):
        file_size = os.path.getsize('Resultados.csv')
        print(f"Debug: CSV file exists, size: {file_size} bytes")
    else:
        print("Debug: CSV file does not exist!")
        return pd.DataFrame()
    
    df = pd.read_csv('Resultados.csv', encoding='latin1', low_memory=False)
    
    print(f"Debug: CSV loaded with shape {df.shape}")
    print(f"Debug: First few rows:\n{df.head(3)}")
    
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
    
    # Get the line names row (row 0) and headers row (row 1)
    line_names_row = df.iloc[0].values
    headers = df.iloc[1].values
    
    # Debug: Print first few headers to see what we're working with
    print(f"Debug: First 10 headers: {headers[:10]}")
    print(f"Debug: Looking for 'Max. loading in %' in headers...")
    
    # Find columns that contain "Max. loading in %" parameter
    max_loading_indices = []
    line_names = []
    
    for i, header in enumerate(headers):
        if header == "Max. loading in %":
            max_loading_indices.append(i)
            # Get the actual line name from the same column in row 0
            line_name = line_names_row[i]
            if line_name == '   ----' or line_name == '----' or pd.isna(line_name):
                line_name = f"Line_{i:03d}"
            line_names.append(line_name)
    
    print(f"Encontradas {len(max_loading_indices)} líneas con datos de cargabilidad máxima.")
    
    if len(max_loading_indices) == 0:
        print("⚠️  No se encontraron datos válidos de cargabilidad.")
        print("Debug: Available headers:", [h for h in headers if 'loading' in str(h).lower() or 'max' in str(h).lower()])
        print("Debug: All headers:", headers[:20])  # Show first 20 headers
        return pd.DataFrame(columns=['Linea', 'Cargabilidad_Maxima'])
    
    # Extract maximum loading data for each line
    line_load_data = []
    
    for i, (idx, line_name) in enumerate(zip(max_loading_indices, line_names)):
        # Get all data for this line (skip first 2 rows which are line names and headers)
        line_data = df.iloc[2:, idx].values
        
        # Debug: Print first few data values for this line
        print(f"Debug: Line {line_name} - First 5 data values: {line_data[:5]}")
        
        # Convert to numeric, handling '----' as NaN
        numeric_data = []
        for val in line_data:
            if val == '   ----' or val == '----' or pd.isna(val):
                numeric_data.append(np.nan)
            else:
                try:
                    numeric_data.append(float(val))
                except (ValueError, TypeError):
                    numeric_data.append(np.nan)
        
        # Find the maximum loading for this line
        max_loading = np.nanmax(numeric_data) if not all(np.isnan(numeric_data)) else np.nan
        
        if not np.isnan(max_loading):
            # Clean line name
            clean_name = line_name.split('\\')[-1] if '\\' in str(line_name) else str(line_name)
            clean_name = clean_name.replace('.ElmLne', '')
            
            # Include all lines (remove voltage filtering for now)
            line_load_data.append({
                'Linea': clean_name,
                'Cargabilidad_Maxima': max_loading
            })
            print(f"Debug: Found valid data for {clean_name}: {max_loading}%")
        else:
            print(f"Debug: No valid data for {line_name} - all values are NaN or '----'")
    
    if not line_load_data:
        print("⚠️  No hay datos de cargabilidad válidos")
        print("Debug: All lines showed '----' values, which means no loading data available")
        print("Debug: This suggests that either:")
        print("  1. No contingencies were analyzed")
        print("  2. The contingencies were analyzed but no lines exceeded the loading threshold")
        print("  3. The contingency analysis configuration needs adjustment")
        return pd.DataFrame(columns=['Linea', 'Cargabilidad_Maxima'])
    
    # Create DataFrame and sort by maximum loading
    line_load_df = pd.DataFrame(line_load_data)
    line_load_df = line_load_df.sort_values(by='Cargabilidad_Maxima', ascending=False).reset_index(drop=True)
    
    print("Procesamiento de cargabilidad completado.")
    return line_load_df


def show_iteration_details(df, iteration_num, substation, current_potencia):
    """Show detailed iteration results from CSV data"""
    print(f"\n📊 DETALLES DE ITERACIÓN {iteration_num} - {substation} ({current_potencia} MW)")
    print("="*60)
    
    # Process the loading data
    line_load_df = process_cargabilidad(df)
    
    if not line_load_df.empty:
        print("Top 10 líneas más cargadas:")
        for i, row in line_load_df.head(10).iterrows():
            print(f"  {i+1}. {row['Linea']}: {row['Cargabilidad_Maxima']:.2f}%")
        
        max_line_load = line_load_df['Cargabilidad_Maxima'].max()
        max_line = line_load_df[line_load_df['Cargabilidad_Maxima'] == max_line_load]['Linea'].values[0]
        
        print(f"\n🎯 RESUMEN: Potencia actual = {current_potencia} MW")
        print(f"Max cargabilidad = {max_line_load:.2f}%, Línea crítica = {max_line}")
    else:
        print("⚠️  No hay datos de cargabilidad válidos")
        print(f"\n🎯 RESUMEN: Potencia actual = {current_potencia} MW")
        print("Max cargabilidad = 0.00%, Línea crítica = N/A")
    
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
            power_flow.Execute()
            
            # Get actual power values from the generator after load flow
            actual_p_gen = static_generator.GetAttribute('c:p')
            actual_q_gen = static_generator.GetAttribute('c:q')
            bus_terminal = static_generator.term
            actual_bus_voltage = bus_terminal.GetAttribute('m:u')
            
            print(f"📊 Estado actual: P = {actual_p_gen:.2f} MW, Q = {actual_q_gen:.2f} MVar, V = {actual_bus_voltage:.4f} pu")
            
            # Check voltage limits
            if actual_bus_voltage < 0.9 or actual_bus_voltage > 1.1:
                print(f"⚠️  Voltaje fuera de límites: {actual_bus_voltage:.4f} pu (límites: 0.9 - 1.1 pu)")
                print(f"🎯 Límite de voltaje alcanzado: Potencia máxima segura = {current_potencia - 1} MW")
                
                # Save result
                results.append({
                    'Subestacion': substation,
                    'Potencia Maxima': current_potencia - 1,
                    'Linea Critica': 'Voltage Limit',
                    'Cargabilidad Maxima': f'V={actual_bus_voltage:.4f}'
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
            
            print(f"🔄 Aumentando potencia a {current_potencia} MW...")
    
    # Convert results to DataFrame
    df_results = pd.DataFrame(results)
    print("Resultados finales:")
    print(df_results)
    
    return df_results
