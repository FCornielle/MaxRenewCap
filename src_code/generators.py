"""
Static Generator Management Module

This module contains functions for creating and managing static generators
in PowerFactory, including power calculations and generator configuration.
"""

import math


def calculate_power_limits(potencia_activa, factor_potencia):
    """
    Calculate apparent power and reactive power limits for a generator.
    
    Args:
        potencia_activa (float): Active power in MW
        factor_potencia (float): Power factor (0-1)
        
    Returns:
        tuple: (potencia_aparente, potencia_reactiva_max, potencia_reactiva_min)
        
    Example:
        >>> s, q_max, q_min = calculate_power_limits(10, 0.95)
    """
    potencia_aparente = potencia_activa / factor_potencia
    potencia_reactiva_max = math.sqrt(potencia_aparente**2 - potencia_activa**2)
    potencia_reactiva_min = -potencia_reactiva_max
    
    return potencia_aparente, potencia_reactiva_max, potencia_reactiva_min


def create_static_generator(app, network_data, hoja_name, barra_name, potencia_activa, factor_potencia):
    """
    Create a static generator in PowerFactory at the specified bus.
    
    Args:
        app: PowerFactory application object
        network_data: PowerFactory network data folder object
        hoja_name (str): Name of the sheet/folder where to create the generator
        barra_name (str): Name of the bus where to connect the generator
        potencia_activa (float): Active power in MW
        factor_potencia (float): Power factor (0-1)
        
    Returns:
        tuple: (bus_voltage, potencia_activa_generada, potencia_reactiva_generada, static_generator, cubicle)
               Returns (None, None, None, None, None) if creation fails
        
    Example:
        >>> voltage, p, q, gen, cubicle = create_static_generator(app, network_data, 'Grid', 'Bus1', 10, 0.95)
    """
    print(f"Creando generador estático en la barra '{barra_name}' con potencia activa {potencia_activa} MW y factor de potencia {factor_potencia}.")
    
    # Calculate power limits
    potencia_aparente, potencia_reactiva_max, potencia_reactiva_min = calculate_power_limits(potencia_activa, factor_potencia)
    print(f"Límites calculados: S = {potencia_aparente:.2f} MVA, Q_max = {potencia_reactiva_max:.2f} MVar, Q_min = {potencia_reactiva_min:.2f} MVar")
    
    print(f"Buscando hoja '{hoja_name}' en 'Network Data'.")
    
    # Look for the Grid data folder (not the network diagram)
    hoja = None
    data_folders = network_data.GetContents('*', 1)
    
    for folder in data_folders:
        print(f"Debug: Found folder: {folder.loc_name} (type: {folder.GetClassName()})")
        if folder.loc_name == hoja_name:
            # Check if it's a data folder (not a network diagram)
            if folder.GetClassName() in ['IntFolder', 'IntPrjfolder', 'ElmNet']:
                print(f"Debug: Found Grid data folder: '{folder.loc_name}' (type: {folder.GetClassName()})")
                # Test if we can create generators in this specific folder
                try:
                    test_gen = folder.CreateObject('ElmGenstat', 'Test_Gen_Temp')
                    if test_gen:
                        print(f"Debug: ✅ Can create generators in Grid folder: '{folder.loc_name}'")
                        test_gen.Delete()
                        hoja = folder
                        break
                    else:
                        print(f"Debug: ❌ Cannot create generators in Grid folder '{folder.loc_name}'")
                except Exception as e:
                    print(f"Debug: ❌ Error testing Grid folder '{folder.loc_name}': {e}")
            elif folder.GetClassName() == 'IntGrfnet':
                print(f"Debug: Found Grid network diagram: '{folder.loc_name}' (type: {folder.GetClassName()})")
                print("Debug: This is a network diagram, not a data folder. Looking for Grid data folder...")
                # Look for the actual Grid data folder that contains generators
                for data_folder in data_folders:
                    if data_folder.loc_name == hoja_name and data_folder.GetClassName() in ['IntFolder', 'IntPrjfolder', 'ElmNet']:
                        print(f"Debug: Found Grid data folder: '{data_folder.loc_name}' (type: {data_folder.GetClassName()})")
                        # Test if we can create generators in this folder
                        try:
                            test_gen = data_folder.CreateObject('ElmGenstat', 'Test_Gen_Temp')
                            if test_gen:
                                print(f"Debug: ✅ Can create generators in Grid folder: '{data_folder.loc_name}'")
                                test_gen.Delete()
                                hoja = data_folder
                                break
                            else:
                                print(f"Debug: ❌ Cannot create generators in Grid folder '{data_folder.loc_name}'")
                        except Exception as e:
                            print(f"Debug: ❌ Error testing Grid folder '{data_folder.loc_name}': {e}")
                break
    
    # If no suitable Grid folder found, use Network Data directly
    if not hoja:
        print("Debug: No suitable Grid data folder found. Using Network Data directly...")
        hoja = network_data
    
    # Find the bus
    bus = None
    for b in app.GetCalcRelevantObjects('*.ElmTerm'):
        if b.loc_name == barra_name:
            bus = b
            break
    
    if not bus:
        print(f"Barra '{barra_name}' no encontrada.")
        return None, None, None, None, None
    
    print(f"Barra '{barra_name}' encontrada. Creando cubículo y generador estático.")
    
    # Clean up any existing generator for this bus
    cleanup_existing_generator(app, barra_name)
    
    # Create unique names based on bus name
    cubicle_name = f'Cubicle_Gen_{barra_name}'
    switcher_name = f'Switch_Gen_{barra_name}'
    generator_name = f'Gen_Estatico_{barra_name}'
    
    # Create cubicle
    cubicle = bus.CreateObject('StaCubic', cubicle_name)
    if not cubicle:
        print(f"Error al crear el cubículo '{cubicle_name}' en la barra '{barra_name}'.")
        return None, None, None, None, None
    
    cubicle.bus1 = bus
    print(f"Cubículo '{cubicle_name}' creado en la barra '{barra_name}' y conectado a la barra.")
    
    # Create switcher
    switcher = cubicle.CreateObject('StaSwitch', switcher_name)
    if not switcher:
        print(f"Error al crear el switcher '{switcher_name}' en el cubículo.")
        cubicle.Delete()
        return None, None, None, None, None
    
    switcher.on_off = 1
    
    # Debug: Check what objects can be created in cubicle
    print(f"Debug: Cubicle type: {cubicle.GetClassName()}")
    print(f"Debug: Cubicle name: {cubicle.loc_name}")
    
    # Debug: Check what objects can be created in the Grid folder
    print(f"Debug: Grid folder type: {hoja.GetClassName()}")
    print(f"Debug: Grid folder name: {hoja.loc_name}")
    
    # Debug: Check existing generators in the project
    print("Debug: Checking existing generators in project...")
    existing_gens = app.GetCalcRelevantObjects('*.ElmGenstat')
    print(f"Debug: Found {len(existing_gens)} existing static generators")
    for gen in existing_gens[:5]:  # Show first 5
        print(f"Debug: - {gen.loc_name} (type: {gen.GetClassName()})")
    
    # Debug: Final check of the hoja we'll use
    print(f"Debug: Final hoja type: {hoja.GetClassName()}")
    print(f"Debug: Final hoja name: {hoja.loc_name}")
    
    # Try to create static generator in the Grid folder first
    print(f"Intentando crear generador en la hoja '{hoja_name}'...")
    static_generator = hoja.CreateObject('ElmGenstat', generator_name)
    if static_generator is None:
        print(f"Error: No se pudo crear el generador estático '{generator_name}' en la hoja '{hoja_name}'.")
        print("Intentando crear en el cubículo como alternativa...")
        
        # Try creating in the cubicle as fallback
        static_generator = cubicle.CreateObject('ElmGenstat', generator_name)
        if static_generator is None:
            print(f"Error: No se pudo crear el generador estático '{generator_name}' en el cubículo tampoco.")
            print("Posibles causas:")
            print("1. El nombre ya existe")
            print("2. No hay permisos para crear objetos en esta ubicación")
            print("3. El tipo de objeto no es válido en este contexto")
            # Clean up created objects
            switcher.Delete()
            cubicle.Delete()
            return None, None, None, None, None
        else:
            print(f"✅ Generador creado exitosamente en el cubículo.")
    else:
        print(f"✅ Generador creado exitosamente en la hoja '{hoja_name}'.")
    
    static_generator.SetAttribute('sgn', potencia_aparente)
    static_generator.SetAttribute('e:pgini', potencia_activa)
    static_generator.SetAttribute('cosn', factor_potencia)
    static_generator.SetAttribute('av_mode', 'constv')  # Constant voltage mode for reactive power management
    static_generator.term = cubicle
    static_generator.SetAttribute('usetp', 1)
    
    # Set reactive power limits using correct PowerFactory attributes
    static_generator.SetAttribute('cQ_max', potencia_reactiva_max)
    static_generator.SetAttribute('cQ_min', potencia_reactiva_min)
    
    # Set voltage reference for constv mode (should be close to bus voltage)
    static_generator.SetAttribute('usetp', 1)  # Enable the generator
    static_generator.SetAttribute('av_mode', 'constv')  # Ensure constv mode
    static_generator.SetAttribute('e:usetp', 1)  # Enable in calculation
    cubicle.obj_id = static_generator
    
    # Run power flow
    print(f"Ejecutando flujo de potencia para la barra '{barra_name}'.")
    power_flow = app.GetFromStudyCase('ComLdf')
    power_flow.Execute()
    
    # Get results
    bus_voltage = bus.GetAttribute('m:u')
    potencia_activa_generada = static_generator.GetAttribute('c:p')
    potencia_reactiva_generada = static_generator.GetAttribute('c:q')
    
    print(f"Generador estático creado: Voltaje barra = {bus_voltage}, P generada = {potencia_activa_generada}, Q generada = {potencia_reactiva_generada}")
    
    return bus_voltage, potencia_activa_generada, potencia_reactiva_generada, static_generator, cubicle


def update_generator_power(static_generator, potencia_activa, factor_potencia):
    """
    Update the power settings of an existing static generator.
    
    Args:
        static_generator: PowerFactory static generator object
        potencia_activa (float): New active power in MW
        factor_potencia (float): New power factor (0-1)
        
    Example:
        >>> update_generator_power(gen, 15, 0.9)
    """
    potencia_aparente, potencia_reactiva_max, potencia_reactiva_min = calculate_power_limits(potencia_activa, factor_potencia)
    
    static_generator.SetAttribute('sgn', potencia_aparente)
    static_generator.SetAttribute('e:pgini', potencia_activa)
    static_generator.SetAttribute('cosn', factor_potencia)
    
    # Set reactive power limits using correct PowerFactory attributes
    static_generator.SetAttribute('cQ_max', potencia_reactiva_max)
    static_generator.SetAttribute('cQ_min', potencia_reactiva_min)
    
    # Ensure constv mode is maintained for reactive power control
    static_generator.SetAttribute('av_mode', 'constv')
    static_generator.SetAttribute('usetp', 1)
    static_generator.SetAttribute('e:usetp', 1)


def delete_generator(static_generator, cubicle):
    """
    Delete a static generator and its associated cubicle.
    
    Args:
        static_generator: PowerFactory static generator object
        cubicle: PowerFactory cubicle object
        
    Example:
        >>> delete_generator(gen, cubicle)
    """
    if static_generator:
        static_generator.Delete()
    if cubicle:
        cubicle.Delete()


def cleanup_existing_generator(app, bus_name):
    """
    Clean up any existing generator and cubicle for a specific bus.
    
    Args:
        app: PowerFactory application object
        bus_name (str): Name of the bus to clean up
        
    Example:
        >>> cleanup_existing_generator(app, 'Bus1')
    """
    print(f"Limpiando generadores existentes para la barra '{bus_name}'...")
    
    # Find the bus
    bus = None
    for b in app.GetCalcRelevantObjects('*.ElmTerm'):
        if b.loc_name == bus_name:
            bus = b
            break
    
    if not bus:
        print(f"Barra '{bus_name}' no encontrada.")
        return
    
    # Look for existing cubicles with generator pattern
    cubicle_name = f'Cubicle_Gen_{bus_name}'
    generator_name = f'Gen_Estatico_{bus_name}'
    
    # Get all cubicles and generators from the entire project
    all_cubicles = app.GetCalcRelevantObjects('*.StaCubic')
    all_generators = app.GetCalcRelevantObjects('*.ElmGenstat')
    
    # Also search in the project folders
    try:
        project_folders = app.GetProjectFolder('net').GetContents('*', 1)
        for folder in project_folders:
            if folder.GetClassName() in ['IntFolder', 'IntPrjfolder']:
                folder_cubicles = folder.GetContents('*.StaCubic', 1)
                folder_generators = folder.GetContents('*.ElmGenstat', 1)
                all_cubicles.extend(folder_cubicles)
                all_generators.extend(folder_generators)
    except Exception as e:
        print(f"Debug: Error searching project folders: {e}")
    
    # Find and delete matching cubicles
    cubicles_deleted = 0
    for cubicle in all_cubicles:
        if cubicle_name in cubicle.loc_name:
            print(f"Eliminando cubículo existente '{cubicle.loc_name}' en la barra '{bus_name}'.")
            try:
                cubicle.Delete()
                cubicles_deleted += 1
            except Exception as e:
                print(f"Error al eliminar cubículo '{cubicle.loc_name}': {e}")
    
    # Find and delete matching generators
    generators_deleted = 0
    for gen in all_generators:
        if generator_name in gen.loc_name:
            print(f"Eliminando generador existente '{gen.loc_name}'.")
            try:
                gen.Delete()
                generators_deleted += 1
            except Exception as e:
                print(f"Error al eliminar generador '{gen.loc_name}': {e}")
    
    print(f"Limpieza completada: {cubicles_deleted} cubículos y {generators_deleted} generadores eliminados.")


def cleanup_all_test_generators(app):
    """
    Clean up all test generators and cubicles created by this script.
    
    Args:
        app: PowerFactory application object
        
    Example:
        >>> cleanup_all_test_generators(app)
    """
    print("Limpiando todos los generadores de prueba existentes...")
    
    # Get all static generators
    all_generators = app.GetCalcRelevantObjects('*.ElmGenstat')
    generators_to_delete = []
    
    for gen in all_generators:
        if 'Gen_Estatico_Bus' in gen.loc_name:
            generators_to_delete.append(gen)
            print(f"Encontrado generador de prueba: '{gen.loc_name}'")
    
    # Delete generators
    for gen in generators_to_delete:
        print(f"Eliminando generador: '{gen.loc_name}'")
        gen.Delete()
    
    # Get all cubicles
    all_cubicles = app.GetCalcRelevantObjects('*.StaCubic')
    cubicles_to_delete = []
    
    for cubicle in all_cubicles:
        if 'Cubicle_Gen_Bus' in cubicle.loc_name:
            cubicles_to_delete.append(cubicle)
            print(f"Encontrado cubículo de prueba: '{cubicle.loc_name}'")
    
    # Delete cubicles
    for cubicle in cubicles_to_delete:
        print(f"Eliminando cubículo: '{cubicle.loc_name}'")
        cubicle.Delete()
    
    print(f"Limpieza completada. Eliminados {len(generators_to_delete)} generadores y {len(cubicles_to_delete)} cubículos.")
