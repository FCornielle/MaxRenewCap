# PowerFactory Generator Optimization Tool

<div align="center">

![PowerFactory](https://img.shields.io/badge/PowerFactory-2021%20SP2-blue)
![Python](https://img.shields.io/badge/Python-3.9-green)
![License](https://img.shields.io/badge/License-Educational%20Use-yellow)

**A professional-grade Python tool for optimizing static generators in PowerFactory through N-1 contingency analysis**

[Features](#features) • [Installation](#installation) • [Usage](#usage) • [Results](#results) • [Architecture](#architecture)

</div>

## 🚀 Overview

This tool automates the process of finding the **maximum safe power injection** at each substation in a PowerFactory network without exceeding loading limits. It performs comprehensive N-1 contingency analysis, handles reactive power limits, and provides detailed optimization results with professional visualizations.

### Key Capabilities

- **Automated N-1 Contingency Analysis** with configurable thresholds
- **Reactive Power Limit Enforcement** for realistic power flow calculations
- **Convergence Detection** and intelligent handling of numerical issues
- **Real-time Progress Monitoring** with detailed iteration logs
- **Professional Results Export** with timestamped CSV files
- **Interactive Visualizations** with bar charts and statistical summaries

## 📊 Sample Results

<div align="center">

![Optimization Results](https://via.placeholder.com/800x400/2E86AB/FFFFFF?text=Sample+Optimization+Results+Chart)

*Example: Maximum power injection capacity by substation*

</div>

### Typical Output Summary
```
📈 RESUMEN ESTADÍSTICO:
   • Total de subestaciones: 15
   • Potencia máxima total: 2,450 MW
   • Potencia promedio: 163.3 MW
   • Potencia máxima individual: 358 MW
   • Potencia mínima individual: 89 MW

🔍 LÍNEAS CRÍTICAS:
   • Line 03 - 04: 3 subestación(es)
   • Line 05 - 06: 2 subestación(es)
   • Line 16 - 17: 2 subestación(es)
```

## 🏗️ Architecture

### Project Structure
```
PowerFactory-Generator-Optimization/
├── 📁 src_code/                    # Core business logic
│   ├── __init__.py                # Package exports
│   ├── pf_env.py                  # PowerFactory environment setup
│   ├── generators.py              # Generator lifecycle management
│   ├── contingency.py             # N-1 analysis & optimization
│   └── csv_parser.py              # Results parsing & processing
├── 📁 notebooks/                  # Interactive workflow
│   ├── max_contingency_job.ipynb  # Main optimization notebook
│   └── Resultados.csv             # Generated results
├── 📁 reference_code/             # Reference implementations
├── 📁 my_env/                     # Python virtual environment
├── 📁 pf_project/                 # PowerFactory project files
├── requirements.txt               # Python dependencies
├── .gitignore                     # Git ignore rules
└── README.md                      # This documentation
```

### Core Modules

| Module | Purpose | Key Functions |
|--------|---------|---------------|
| `pf_env.py` | Environment setup | `pf_enviroment()`, `initialize_powerfactory()` |
| `generators.py` | Generator management | `create_static_generator()`, `update_generator_power()` |
| `contingency.py` | Optimization engine | `optimize_generators_for_substations()` |
| `csv_parser.py` | Data processing | `parse_contingency_results()` |

## ⚙️ Installation

### Prerequisites
- **PowerFactory 2021 SP2** (or compatible version)
- **Python 3.9** (PowerFactory requirement)
- **Valid PowerFactory License**

### Setup Steps

1. **Clone the repository**
   ```bash
   git clone <repository-url>
   cd PowerFactory-Generator-Optimization
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Verify PowerFactory installation**
   - Ensure PowerFactory is installed and licensed
   - Note the Python path: `C:\Program Files\DIgSILENT\PowerFactory 2021 SP2\Python\3.9`

## 🎯 Usage

### Quick Start (Notebook Workflow)

1. **Open Jupyter Notebook**
   ```bash
   jupyter notebook notebooks/max_contingency_job.ipynb
   ```

2. **Configure parameters** in the first cell:
   ```python
   # Substations to analyze
   substations = ["bonao3", "canabacoa", "guayubin"]
   
   # Optimization parameters
   initial_potencia = 1      # MW
   factor_potencia = 0.95    # Power factor
   max_cargabilidad = 110    # % loading limit
   threshold_inconvergence = 10  # % convergence threshold
   ```

3. **Run the optimization workflow** cell by cell

4. **View results** in the final visualization cell

### Advanced Configuration

#### PowerFactory Environment
```python
# Configure PowerFactory connection
dig_path = r'C:\Program Files\DIgSILENT\PowerFactory 2021 SP2\Python\3.9'
project_name = "Your Project Name"
study_case = "1. Power Flow"
operation_scenario = "Your Scenario"
```

#### Optimization Parameters
```python
# Fine-tune optimization behavior
INITIAL_POTENCIA = 1           # Starting power (MW)
FACTOR_POTENCIA = 0.95         # Power factor
MAX_CARGABILIDAD = 110         # Maximum line loading (%)
THRESHOLD_INCONVERGENCE = 10   # Convergence jump threshold (%)
```

## 📈 Results & Outputs

### Automatic CSV Export
The tool automatically generates timestamped results files:
```
optimization_results_20241215_143022.csv
```

### CSV Structure
| Column | Description | Example |
|--------|-------------|---------|
| `Subestacion` | Substation name | "bonao3" |
| `Potencia_Maxima` | Maximum safe power (MW) | 358 |
| `Linea_Critica` | Limiting line | "Line 03 - 04" |
| `Cargabilidad_Maxima` | Maximum loading (%) | 110.03 |

### Visualization Features
- **Horizontal bar charts** sorted by power capacity
- **Value labels** showing exact MW values
- **Statistical summaries** with key metrics
- **Critical line analysis** identifying bottlenecks

## 🔧 Technical Details

### Algorithm Flow
1. **Environment Setup**: Initialize PowerFactory connection
2. **Project Activation**: Load project, study case, and scenario
3. **Generator Creation**: Create static generator at each substation
4. **Iterative Optimization**:
   - Run power flow with reactive power limits
   - Execute N-1 contingency analysis
   - Check line loading constraints
   - Increment power until limit reached
5. **Results Export**: Save timestamped CSV with comprehensive data

### Key Technical Features
- **Reactive Power Limits**: `iopt_lim = 1` ensures realistic power flow
- **Convergence Detection**: Handles numerical instabilities intelligently
- **Clean Output**: Removed debug messages for professional presentation
- **Error Handling**: Robust error handling with graceful degradation

### Performance Considerations
- **Memory Efficient**: Processes one substation at a time
- **Cleanup**: Automatic generator removal after each test
- **Progress Tracking**: Real-time iteration monitoring

## 🛠️ Troubleshooting

### Common Issues

| Issue | Solution |
|-------|----------|
| **PowerFactory Connection Failed** | Verify PowerFactory is running and licensed |
| **Import Errors** | Check Python path and install dependencies |
| **Generator Creation Failed** | Verify bus names and folder permissions |
| **Convergence Issues** | Adjust `THRESHOLD_INCONVERGENCE` parameter |

### Debug Mode
Enable detailed logging by modifying print statements in source code.

## 📚 API Reference

### Core Functions

#### Environment Management
```python
from src_code import pf_enviroment, initialize_powerfactory

# Initialize PowerFactory environment
pf_enviroment(dig_path)
app = initialize_powerfactory()
```

#### Generator Operations
```python
from src_code import create_static_generator, update_generator_power

# Create and manage generators
generator = create_static_generator(app, network_data, hoja, substation, power, pf)
update_generator_power(generator, new_power, pf)
```

#### Optimization Engine
```python
from src_code import optimize_generators_for_substations

# Run complete optimization
results = optimize_generators_for_substations(
    app=app,
    substations=substations,
    network_data=network_data,
    hoja=hoja,
    initial_potencia=1,
    factor_potencia=0.95,
    max_cargabilidad=110,
    threshold_inconvergence=10
)
```

## 🤝 Contributing

1. **Fork** the repository
2. **Create** a feature branch (`git checkout -b feature/amazing-feature`)
3. **Commit** your changes (`git commit -m 'Add amazing feature'`)
4. **Push** to the branch (`git push origin feature/amazing-feature`)
5. **Open** a Pull Request

### Development Guidelines
- Follow PEP 8 style guidelines
- Add comprehensive docstrings
- Test with PowerFactory before submitting
- Update documentation for new features

## 📄 License

This project is provided for **educational and research purposes**. Please ensure compliance with DIgSILENT PowerFactory licensing terms.

## 🆘 Support

| Type | Resource |
|------|----------|
| **PowerFactory API** | [DIgSILENT Documentation](https://www.digsilent.de/) |
| **Python Issues** | [Python Documentation](https://docs.python.org/) |
| **Tool Issues** | Create an issue in this repository |

## 📋 Version History

- **v2.0.0** (Current)
  - ✅ Removed unused modules (`io_utils.py`, `plotting.py`)
  - ✅ Cleaned debug output for professional presentation
  - ✅ Added automatic CSV export with timestamps
  - ✅ Enhanced visualization with statistical summaries
  - ✅ Improved error handling and convergence detection
  - ✅ Streamlined project structure

- **v1.0.0** (Legacy)
  - Initial release with modular design

---

<div align="center">

**Built with ❤️ for PowerFactory optimization**

*Professional-grade tools for power system analysis*

</div>