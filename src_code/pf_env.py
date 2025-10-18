"""
PowerFactory Environment Setup Module

This module contains functions for setting up the PowerFactory environment,
managing projects, study cases, and operation scenarios.
"""

import os
import sys


def pf_enviroment(dig_path):
    """
    Initialize PowerFactory environment by adding the path to sys.path and PATH.
    
    Args:
        dig_path (str): Path to PowerFactory Python installation directory
        
    Example:
        >>> pf_enviroment(r'C:\\Program Files\\DIgSILENT\\PowerFactory 2021 SP2\\Python\\3.9')
    """
    sys.path.append(dig_path)
    os.environ['PATH'] += f';{dig_path}'
    print(f"PowerFactory environment initialized with path: {dig_path}")


def initialize_powerfactory():
    """
    Initialize and connect to PowerFactory application.
    
    Returns:
        PowerFactory application object or None if failed
        
    Example:
        >>> app = initialize_powerfactory()
    """
    try:
        import powerfactory as pf
        
        # Try to get the application
        app = pf.GetApplication()
        
        # If app is None, try alternative methods
        if app is None:
            print("First attempt failed, trying alternative connection methods...")
            
            # Try to get application with different method
            try:
                app = pf.GetApplication()
                if app is not None:
                    print("PowerFactory application connected on second attempt!")
                    return app
            except:
                pass
            
            # If still None, try to initialize PowerFactory
            try:
                print("Attempting to initialize PowerFactory...")
                pf.Initialize()
                app = pf.GetApplication()
                if app is not None:
                    print("PowerFactory initialized and connected successfully!")
                    return app
            except Exception as init_error:
                print(f"PowerFactory initialization failed: {init_error}")
            
            print("PowerFactory is not running or not properly initialized.")
            print("Please make sure:")
            print("1. PowerFactory is running")
            print("2. Your project is loaded in PowerFactory")
            print("3. You have a valid PowerFactory license")
            print("4. The PowerFactory path is correct")
            return None
        else:
            print("PowerFactory application connected successfully!")
            return app
            
    except ImportError as e:
        print(f"Error importing PowerFactory: {e}")
        print("Please check if PowerFactory is installed and the path is correct.")
        return None
    except Exception as e:
        print(f"Error initializing PowerFactory: {e}")
        return None


def activate_project(app, project_name):
    """
    Activate a PowerFactory project by name.
    
    Args:
        app: PowerFactory application object
        project_name (str): Name of the project to activate
        
    Returns:
        int: Project activation result (0 = success)
        
    Example:
        >>> project = activate_project(app, "39 Bus New England System")
    """
    if app is None:
        print("Error: PowerFactory application is not initialized.")
        return None
    
    try:
        project = app.ActivateProject(project_name)
        if project == 0:
            print(f"Proyecto '{project_name}' activado con exito.")
        else:
            print(f"No se pudo activar el proyecto '{project_name}'. Verifica si el nombre es correcto.")
        return project
    except Exception as e:
        print(f"Error activating project '{project_name}': {e}")
        return None


def list_and_select_study_case(app, study_case_name: str):
    """
    List all available study cases and activate the specified one.
    
    Args:
        app: PowerFactory application object
        study_case_name (str): Name of the study case to activate
        
    Returns:
        PowerFactory study case object: The activated study case
        
    Example:
        >>> study_case = list_and_select_study_case(app, "1. Power Flow")
    """
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
    """
    List all available operation scenarios and activate the specified one.
    
    Args:
        app: PowerFactory application object
        scenario_name (str): Name of the operation scenario to activate
        
    Returns:
        PowerFactory scenario object: The activated operation scenario or None if not found
        
    Example:
        >>> scenario = list_and_activate_operation_scenario(app, "Operation Scenario blank space")
    """
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
