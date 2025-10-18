# PowerFactory Generator Optimization Tool

A comprehensive Python tool for optimizing static generators in PowerFactory by running N-1 contingency analysis and finding the maximum safe power that can be injected at each substation without exceeding loading limits.

## Features

- **Modular Design**: Clean, organized code structure with separate modules for different functionalities
- **N-1 Contingency Analysis**: Automated contingency analysis with configurable parameters
- **Generator Optimization**: Find maximum safe power injection at each substation
- **Convergence Detection**: Smart detection and handling of convergence issues
- **Results Visualization**: Comprehensive dashboards and plots
- **Export Capabilities**: Save results in multiple formats (CSV, text reports)
- **Error Handling**: Robust error handling and cleanup procedures

## Project Structure

```
PowerFactory-Generator-Optimization/
├── main.py                 # Main execution script
├── requirements.txt        # Python dependencies
├── README.md              # This file
├── src_code/              # Source code package
│   ├── __init__.py        # Package initialization
│   ├── pf_env.py          # PowerFactory environment setup
│   ├── generators.py      # Generator management functions
│   ├── contingency.py     # Contingency analysis functions
│   ├── io_utils.py        # Input/output utilities
│   └── plotting.py        # Visualization functions
├── output/                # Results output directory (created automatically)
├── plots/                 # Generated plots directory (created automatically)
└── notebooks/             # Jupyter notebooks (legacy code)
    ├── max_contingency_analysis.ipynb
    └── max_contingency_job.ipynb
```

## Prerequisites

1. **PowerFactory Software**: DIgSILENT PowerFactory 2021 SP2 or compatible version
2. **Python Environment**: Python 3.9 (as required by PowerFactory)
3. **PowerFactory License**: Valid PowerFactory license for running calculations

## Installation

1. **Clone or download** this repository to your local machine

2. **Install Python dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Verify PowerFactory installation**:
   - Ensure PowerFactory is installed and licensed
   - Note the PowerFactory Python installation path (usually `C:\Program Files\DIgSILENT\PowerFactory 2021 SP2\Python\3.9`)

## Usage

### Quick Start

1. **Open PowerFactory** and load your project
2. **Run the main script**:
   ```bash
   python main.py
   ```

### Configuration

Edit the configuration section in `main.py` to customize:

```python
# PowerFactory configuration
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

# Substations to analyze
SUBSTATIONS = [
    "bonao3",
    "canabacoa", 
    "guayubin"
    # Add more substations as needed
]
```

### Advanced Usage

#### Using Individual Modules

```python
from src_code import (
    initialize_powerfactory,
    create_static_generator,
    run_contingency_analysis,
    optimize_generators_for_substations
)

# Initialize PowerFactory
app = initialize_powerfactory()

# Run optimization for specific substations
results = optimize_generators_for_substations(
    app=app,
    substations=["substation1", "substation2"],
    network_data=network_data,
    hoja="Grid",
    initial_potencia=1,
    factor_potencia=0.95,
    max_cargabilidad=110
)
```

#### Custom Analysis

```python
from src_code.contingency import run_contingency_analysis, process_cargabilidad
from src_code.plotting import plot_contingency_results

# Run contingency analysis
df = run_contingency_analysis(app)

# Process results
line_load_df = process_cargabilidad(df)

# Create visualization
plot_contingency_results(line_load_df, top_n=20)
```

## Output Files

The tool generates several output files:

### Results Directory (`output/`)
- `optimization_results.csv`: Detailed results for each substation
- `summary_statistics.csv`: Statistical summary of results
- `optimization_report.txt`: Human-readable report

### Plots Directory (`plots/`)
- `optimization_dashboard.png`: Comprehensive dashboard with multiple plots
- Individual plot files for specific analyses

## Key Functions

### Environment Management
- `pf_enviroment()`: Initialize PowerFactory Python environment
- `initialize_powerfactory()`: Connect to PowerFactory application
- `activate_project()`: Activate PowerFactory project
- `list_and_select_study_case()`: Select study case
- `list_and_activate_operation_scenario()`: Select operation scenario

### Generator Management
- `create_static_generator()`: Create static generator at specified bus
- `update_generator_power()`: Update generator power settings
- `delete_generator()`: Delete generator and associated equipment
- `cleanup_all_test_generators()`: Clean up all test generators

### Contingency Analysis
- `run_contingency_analysis()`: Execute N-1 contingency analysis
- `process_cargabilidad()`: Process contingency results
- `optimize_generators_for_substations()`: Main optimization algorithm

### Visualization
- `plot_contingency_results()`: Plot line loading results
- `plot_generator_optimization()`: Plot optimization results
- `create_optimization_dashboard()`: Create comprehensive dashboard

## Troubleshooting

### Common Issues

1. **PowerFactory Connection Failed**
   - Ensure PowerFactory is running
   - Check that your project is loaded
   - Verify PowerFactory license is valid

2. **Import Errors**
   - Check PowerFactory Python path in `main.py`
   - Ensure all dependencies are installed: `pip install -r requirements.txt`

3. **Generator Creation Failed**
   - Verify bus names exist in the project
   - Check that the specified sheet/folder exists
   - Ensure you have write permissions in the project

4. **Convergence Issues**
   - Adjust `THRESHOLD_INCONVERGENCE` parameter
   - Check network topology and loading conditions
   - Verify generator parameters are realistic

### Debug Mode

Enable detailed logging by modifying the print statements in the source code or adding logging configuration.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly with PowerFactory
5. Submit a pull request

## License

This project is provided as-is for educational and research purposes. Please ensure compliance with DIgSILENT PowerFactory licensing terms.

## Support

For issues related to:
- **PowerFactory API**: Consult DIgSILENT documentation
- **Python dependencies**: Check package documentation
- **This tool**: Create an issue in the repository

## Version History

- **v1.0.0**: Initial release with modular design and comprehensive functionality

---

**Note**: This tool requires a valid PowerFactory license and should be used in accordance with DIgSILENT's terms of service.