#!/usr/bin/env python3
"""
PowerFactory Generator Optimization Tool

This script provides a complete solution for optimizing static generators
in PowerFactory by running N-1 contingency analysis and finding the maximum
safe power that can be injected at each substation without exceeding
loading limits.

Author: PowerFactory Python Scripting
Version: 1.0
"""

import os
import sys
from src_code.pf_env import pf_enviroment, initialize_powerfactory, activate_project, list_and_select_study_case, list_and_activate_operation_scenario
from src_code.generators import cleanup_all_test_generators
from src_code.contingency import optimize_generators_for_substations
from src_code.io_utils import create_results_summary
from src_code.plotting import create_optimization_dashboard


def main():
    """
    Main function to run the PowerFactory generator optimization process.
    """
    print("=" * 60)
    print("PowerFactory Generator Optimization Tool")
    print("=" * 60)
    
    # Configuration
    DIG_PATH = r'C:\Program Files\DIgSILENT\PowerFactory 2021 SP2\Python\3.9'
    PROJECT_NAME = "39 Bus New England System"
    STUDY_CASE_NAME = "1. Power Flow"
    OPERATION_SCENARIO_NAME = "Ene_Bloque_3_dem_min_diurna_2028"
    HOJA_NAME = "NORTE"
    
    # Optimization parameters
    INITIAL_POTENCIA = 1  # MW
    FACTOR_POTENCIA = 0.95
    MAX_CARGABILIDAD = 110  # %
    THRESHOLD_INCONVERGENCE = 10  # %
    
    # Substations to analyze (uncomment the ones you want to test)
    SUBSTATIONS = [
        # "138 kV Agua Clara",
        # "138 kV Dajabón",
        # "138 kV Dominicana Azul",
        # "138 kV El Naranjo B2 Sec A",
        # "138 kV Esperanza",
        # "138 kV Frontera",
        # "138 kV Gaspar Hernandez",
        # "138 kV Guanillo",
        # "138 kV Guayubín",
        # "138 kV La Gallera Sec B",
        # "138 kV La Vega II Sec A",
        # "138 kV Los Guzmancito",
        # "138 kV Matrisol",
        # "138 kV Moca",
        # "138 kV Monción",
        # "138 kV Monte Cristi",
        # "138 kV Nagua 2",
        # "138 kV Navarrete 2 B1 Sec A",
        # "138 kV Nibaje",
        # "138 kV PSF Manzanillo",
        # "138 kV Palamara B1",
        # "138 kV Payita",
        # "138 kV Pimentel Sec B",
        # "138 kV Playa Dorada",
        # "138 kV Puerto Plata 2 B1 Sec A",
        # "138 kV Río San Juan",
        # "138 kV SFM BT",
        # "138 kV Sajoma",
        # "138 kV Salcedo",
        # "138 kV Santiago Norte Sec A",
        # "138 kV Santiago Rodríguez",
        # "138 kV Solsur",
        # "138 kV Sánchez",
        # "138 kV Tavera B2",
        # "138 kV Valverde Mao BT",
        # "138 kV ZF Santiago B1",
        # "138 kV Pimentel Energy",
        # "bonao2",
        "bonao3",
        "canabacoa",
        "guayubin"
    ]
    
    try:
        # Step 1: Initialize PowerFactory environment
        print("\n1. Initializing PowerFactory environment...")
        pf_enviroment(DIG_PATH)
        
        # Step 2: Connect to PowerFactory application
        print("\n2. Connecting to PowerFactory application...")
        app = initialize_powerfactory()
        if app is None:
            print("Error: Could not connect to PowerFactory. Please check your installation and license.")
            return
        
        # Step 3: Activate project
        print(f"\n3. Activating project '{PROJECT_NAME}'...")
        project_result = activate_project(app, PROJECT_NAME)
        if project_result != 0:
            print(f"Error: Could not activate project '{PROJECT_NAME}'.")
            return
        
        # Step 4: Activate study case
        print(f"\n4. Activating study case '{STUDY_CASE_NAME}'...")
        study_case = list_and_select_study_case(app, STUDY_CASE_NAME)
        if study_case is None:
            print(f"Error: Could not activate study case '{STUDY_CASE_NAME}'.")
            return
        
        # Step 5: Activate operation scenario
        print(f"\n5. Activating operation scenario '{OPERATION_SCENARIO_NAME}'...")
        scenario = list_and_activate_operation_scenario(app, OPERATION_SCENARIO_NAME)
        if scenario is None:
            print(f"Error: Could not activate operation scenario '{OPERATION_SCENARIO_NAME}'.")
            return
        
        # Step 6: Get network data
        print("\n6. Getting network data...")
        project = app.GetActiveProject()
        network_data = project.GetContents('Network Model.IntPrjfolder\\Network Data', 1)[0]
        if network_data is None:
            print("Error: Could not access network data.")
            return
        
        # Step 7: Clean up any existing test generators
        print("\n7. Cleaning up existing test generators...")
        cleanup_all_test_generators(app)
        
        # Step 8: Run optimization
        print(f"\n8. Running optimization for {len(SUBSTATIONS)} substations...")
        print(f"   - Initial power: {INITIAL_POTENCIA} MW")
        print(f"   - Power factor: {FACTOR_POTENCIA}")
        print(f"   - Max loading: {MAX_CARGABILIDAD}%")
        print(f"   - Convergence threshold: {THRESHOLD_INCONVERGENCE}%")
        
        results_df = optimize_generators_for_substations(
            app=app,
            substations=SUBSTATIONS,
            network_data=network_data,
            hoja=HOJA_NAME,
            initial_potencia=INITIAL_POTENCIA,
            factor_potencia=FACTOR_POTENCIA,
            max_cargabilidad=MAX_CARGABILIDAD,
            threshold_inconvergence=THRESHOLD_INCONVERGENCE
        )
        
        # Step 9: Create results summary
        print("\n9. Creating results summary...")
        if results_df is not None and not results_df.empty:
            files_created = create_results_summary(results_df, output_dir='output')
            print(f"Results saved to: {files_created}")
            
            # Step 10: Create visualization dashboard
            print("\n10. Creating optimization dashboard...")
            try:
                create_optimization_dashboard(results_df, output_dir='plots')
                print("Dashboard created successfully!")
            except Exception as e:
                print(f"Warning: Could not create dashboard: {e}")
        else:
            print("No results to summarize.")
        
        print("\n" + "=" * 60)
        print("Optimization process completed successfully!")
        print("=" * 60)
        
    except Exception as e:
        print(f"\nError during execution: {e}")
        print("Please check your PowerFactory setup and try again.")
        return 1
    
    return 0


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
