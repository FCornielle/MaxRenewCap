"""
PowerFactory Generator Optimization Package

This package provides tools for optimizing static generators in PowerFactory
by running N-1 contingency analysis and finding the maximum safe power
that can be injected at each substation without exceeding loading limits.

Modules:
    - pf_env: PowerFactory environment setup and project management
    - generators: Static generator creation and management
    - contingency: Contingency analysis and optimization algorithms
    - io_utils: Input/output utilities and file management
    - plotting: Visualization and dashboard creation

Author: PowerFactory Python Scripting
Version: 1.0
"""

__version__ = "1.0.0"
__author__ = "PowerFactory Python Scripting"

# Import main functions for easy access
from .pf_env import (
    pf_enviroment,
    initialize_powerfactory,
    activate_project,
    list_and_select_study_case,
    list_and_activate_operation_scenario
)

from .generators import (
    create_static_generator,
    update_generator_power,
    delete_generator,
    calculate_power_limits,
    cleanup_existing_generator,
    cleanup_all_test_generators
)

from .contingency import (
    run_contingency_analysis,
    process_cargabilidad,
    optimize_generators_for_substations,
    show_iteration_details
)
from .csv_parser import (
    parse_contingency_results,
    get_max_loading_summary
)

from .io_utils import (
    load_contingency_results,
    save_results_to_csv,
    ensure_directory_exists,
    get_safe_filename,
    create_results_summary
)

from .plotting import (
    plot_contingency_results,
    plot_generator_optimization,
    plot_power_flow_summary,
    create_optimization_dashboard
)

__all__ = [
    # Environment functions
    'pf_enviroment',
    'initialize_powerfactory',
    'activate_project',
    'list_and_select_study_case',
    'list_and_activate_operation_scenario',
    
    # Generator functions
    'create_static_generator',
    'update_generator_power',
    'delete_generator',
    'calculate_power_limits',
    'cleanup_existing_generator',
    'cleanup_all_test_generators',
    
    # Contingency functions
    'run_contingency_analysis',
    'process_cargabilidad',
    'optimize_generators_for_substations',
    'show_iteration_details',
    
    # CSV parsing functions
    'parse_contingency_results',
    'get_max_loading_summary',
    
    # I/O functions
    'load_contingency_results',
    'save_results_to_csv',
    'ensure_directory_exists',
    'get_safe_filename',
    'create_results_summary',
    
    # Plotting functions
    'plot_contingency_results',
    'plot_generator_optimization',
    'plot_power_flow_summary',
    'create_optimization_dashboard'
]