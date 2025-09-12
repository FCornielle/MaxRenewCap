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
    
    print(f"Buscando hoja '{hoja_name}' en 'Network Data'.")
    hoja = network_data.GetContents(hoja_name, 1)[0]
    if not hoja:
        print(f"No se encontró la hoja '{hoja_name}' en 'Network Data'.")
        return None, None, None, None, None
    
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
    
    # Create cubicle
    cubicle = bus.CreateObject('StaCubic', 'Cubicle_Generador')
    cubicle.bus1 = bus
    if cubicle:
        print(f"Cubículo creado en la barra '{barra_name}' y conectado a la barra.")
    else:
        print(f"Error al crear el cubículo en la barra '{barra_name}'.")
        return None, None, None, None, None
    
    # Create switcher
    switcher = cubicle.CreateObject('StaSwitch', 'Switcher_Generador')
    switcher.on_off = 1
    
    # Create static generator
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
    
    # Run power flow
    print(f"Ejecutando flujo de potencia para la barra '{barra_name}'.")
    power_flow = app.GetFromStudyCase('ComLdf')
    power_flow.Execute()
    
    # Get results
    bus_voltage = bus.GetAttribute('m:u')
    potencia_activa_generada = static_generator.GetAttribute('m:P:bus1')
    potencia_reactiva_generada = static_generator.GetAttribute('m:Q:bus1')
    
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
    static_generator.SetAttribute('cQ_max', potencia_reactiva_max)
    static_generator.SetAttribute('cQ_min', potencia_reactiva_min)


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
