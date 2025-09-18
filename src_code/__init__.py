"""
PowerFactory Contingency Analysis Module

This module provides tools for PowerFactory contingency analysis and generator optimization.
It's organized into several submodules for better maintainability and reusability.

Modules:
- pf_env: PowerFactory environment setup and project management
- contingency: Contingency analysis and optimization functions
- generators: Static generator creation and management
- io_utils: File I/O utilities for CSV and data handling
- plotting: Visualization functions (optional)
"""

# Import main functions for easy access
from .pf_env import (
    pf_enviroment,
    initialize_powerfactory,
    activate_project,
    list_and_select_study_case,
    list_and_activate_operation_scenario
)

from .contingency import (
    run_contingency_analysis,
    process_cargabilidad,
    optimize_generators_for_substations
)

from .generators import (
    create_static_generator,
    calculate_power_limits,
    delete_generator,
    update_generator_power,
    cleanup_existing_generator
)

from .io_utils import (
    load_contingency_results,
    save_results_to_csv
)

# Optional plotting functions
try:
    from .plotting import (
        plot_contingency_results,
        plot_generator_optimization
    )
except ImportError:
    # Plotting functions are optional
    pass

__version__ = "1.0.0"
__author__ = "PowerFactory Python Scripting"

# Make main functions available at package level
__all__ = [
    # PowerFactory environment
    'pf_enviroment',
    'initialize_powerfactory',
    'activate_project', 
    'list_and_select_study_case',
    'list_and_activate_operation_scenario',
    
    # Contingency analysis
    'run_contingency_analysis',
    'process_cargabilidad',
    'optimize_generators_for_substations',
    
    # Generator management
    'create_static_generator',
    'calculate_power_limits',
    'delete_generator',
    'update_generator_power',
    'cleanup_existing_generator',
    
    # I/O utilities
    'load_contingency_results',
    'save_results_to_csv',
]
