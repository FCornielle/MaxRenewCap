# PowerFactory Contingency Analysis Optimizer

## Project Overview

This project implements an automated contingency analysis optimizer for power systems using DIgSILENT PowerFactory. The system creates static generators at various substations and determines the maximum safe power injection before any transmission line exceeds 110% loading capacity under N-1 contingency conditions.

## Key Features

- **Static Generator Management**: Creates and manages static generators with constant voltage (`constv`) control mode
- **N-1 Contingency Analysis**: Automatically runs contingency analysis for all line outages
- **Line Loading Optimization**: Finds maximum safe power before exceeding 110% line loading
- **Inconvergence Detection**: Handles cases where power increases cause system instability
- **Comprehensive Results**: Tracks maximum power, critical lines, and iterations for each substation

## Project Setup Considerations

### ⚠️ **IMPORTANT: System Modifications Required**

Before running this project, the following modifications must be made to the PowerFactory project:

#### 1. **Line Capacity Duplication**
- **Lines exceeding 100% loading must be duplicated** to create parallel circuits
- Each overloaded line should have a second identical line in parallel
- This effectively doubles the capacity of overloaded transmission corridors
- **Reason**: The original 39 Bus New England System has several lines that are already overloaded under normal load flow conditions (e.g., Line 23-24 at 158%, Line 21-22 at 157%, etc.). Duplicating these lines reduces their loading to manageable levels (typically 50-80%) and allows for proper testing of the contingency analysis.

**Example of Required Duplication:**
- Lines over 100% loading: Line 23-24, Line 21-22, Line 05-06, Line 06-07, Line 05-08, Line 16-21, Line 10-11, Line 22-23, Line 10-13, Line 06-11, Line 13-14, Line 16-24
- After duplication, these lines show as two parallel circuits (Par.no. 1 and Par.no. 2) with reduced loading percentages

#### 2. **Generator and Transformer Duplication**
- **The following components must be duplicated** to create parallel units:
  - **Static Generator 'G 02'** (Type Gen 02) - Currently operating at ~76-77 MW, 151-152 MVar
  - **Transformer 'Trf 06 - 31'** (Type 06 - 31 YNy0) - Currently operating at high loading levels
- Each original generator and transformer should have a second identical unit in parallel
- **Reason**: These components are operating close to their limits under normal conditions (as shown in the power flow results). Duplicating them provides additional capacity headroom for testing the contingency analysis and power injection without immediate overloads before adding static generators for maximum power injection testing.

#### 3. **Current System Status**
The system has been properly configured with duplicated lines. Current loading percentages show:
- **Duplicated lines (Par.no. 2)**: Loading reduced to 50-80% range
- **Original lines (Par.no. 1)**: Maintained for reference
- **Lines 09-39A and 09-39B**: Show empty loading (likely disconnected or special configuration)

#### 4. **Contingency Analysis Configuration**
- **Create fault cases for ALL lines** in the system (both original and duplicated)
- Each line should have a corresponding contingency case
- Configure contingency analysis with the following settings:
  - **Calculation Method**: AC Load Flow Calculation
  - **Load Flow**: Use the standard Load Flow Calculation
  - **Static Contingencies**: Include all line outages (typically 34+ contingencies)
  - **Dynamic Contingencies**: Not required for this analysis

#### 5. **Variable Selection Setup**
- **Only include "Line" objects** in the variable selection
- Remove all other object types (transformers, generators, loads, etc.)
- This ensures the analysis focuses only on transmission line loading
- The variable selection should contain only line loading variables

#### 6. **Static Generator Placement**
- **Static generators will be connected to duplicated lines** (Par.no. 2) when available
- This ensures the analysis tests the system under realistic loading conditions
- The duplicated lines provide the necessary capacity headroom for power injection testing

#### 7. **Contingency Analysis Parameters**
Configure the contingency analysis with these specific settings:
```
- iopt_Linear = 0 (AC Load Flow)
- loadmax = 50 (Maximum loading percentage to analyze)
- vlmin = 0.9 (Minimum voltage limit)
- vlmax = 1.1 (Maximum voltage limit)
- vmax_step = 5 (Voltage step for analysis)
```

## Project Structure

```
├── src_code/
│   ├── __init__.py
│   ├── pf_env.py              # PowerFactory environment setup
│   ├── generators.py          # Static generator management
│   ├── contingency.py         # Contingency analysis functions
│   ├── plotting.py            # Visualization functions
│   └── io_utils.py            # Input/output utilities
├── notebooks/
│   ├── max_contingency_job.ipynb      # Main analysis notebook
│   └── max_contingency_analysis.ipynb # Reference implementation
├── reference_code/
│   └── code.ipynb             # PowerFactory reference examples
├── pf_project/
│   └── 39 Bus New England System.pfd  # PowerFactory project file
├── Resultados.csv             # Contingency analysis results
├── requirements.txt           # Python dependencies
└── README.md                  # This file
```

## Installation

1. **Install Python Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **PowerFactory Setup**:
   - Install DIgSILENT PowerFactory 2021 SP2 or later
   - Ensure Python 3.9 is available in PowerFactory's Python environment
   - Update the PowerFactory path in the notebook if needed

3. **Project Modifications** (See Setup Considerations above):
   - Duplicate all transmission lines
   - Configure contingency analysis
   - Set up variable selection

## Usage

### Running the Analysis

1. **Open the main notebook**:
   ```bash
   jupyter notebook notebooks/max_contingency_job.ipynb
   ```

2. **Execute cells in order**:
   - Cell 1-8: Initialize PowerFactory and activate project
   - Cell 9-14: Set up study case and operation scenario
   - Cell 15-17: Configure network data
   - Cell 23: **Run the complete contingency analysis optimizer**

### Configuration Parameters

The analysis can be customized by modifying these parameters in the main cell:

```python
# Substations to analyze
substations = [
    "Bus 03", "Bus 04", "Bus 05", "Bus 06", 
    "Bus 07", "Bus 08", "Bus 09"
]

# Analysis parameters
initial_potencia = 1          # Starting power (MW)
factor_potencia = 0.95        # Power factor
max_cargabilidad = 110        # Maximum line loading (%)
threshold_inconvergence = 10  # Inconvergence threshold (%)
```

## How It Works

### 1. **Generator Creation**
- Creates static generators with `constv` (constant voltage) control mode
- Generators are placed in the correct Grid data folder
- Each generator starts at 1 MW with 0.95 power factor

### 2. **Contingency Analysis Loop**
For each substation:
- Creates a static generator
- Incrementally increases power by 1 MW
- Runs N-1 contingency analysis after each increase
- Checks all line loading percentages
- Stops when any line exceeds 110% loading

### 3. **Results Processing**
- Identifies the critical line that limits power injection
- Records maximum safe power for each substation
- Tracks loading percentages and iteration counts
- Provides comprehensive summary statistics

## Output Results

The analysis produces a DataFrame with the following columns:

| Column | Description |
|--------|-------------|
| `Subestacion` | Bus name where generator was placed |
| `Potencia_Maxima` | Maximum safe power (MW) |
| `Linea_Critica` | Critical line that limits power |
| `Cargabilidad_Maxima` | Loading percentage of critical line |
| `Iteraciones` | Number of iterations required |

## Troubleshooting

### Common Issues

1. **Generator Creation Fails**:
   - Ensure PowerFactory project is properly loaded
   - Check that Grid folder exists and is accessible
   - Verify bus names are correct

2. **Contingency Analysis Errors**:
   - Confirm contingency cases are properly configured
   - Check that variable selection includes only lines
   - Verify contingency analysis module is available

3. **Line Overloading**:
   - ✅ **RESOLVED**: Lines have been properly duplicated (Par.no. 2 shows 50-80% loading)
   - Verify that duplicated lines (Par.no. 2) are being used for generator connections
   - Check that load flow converges before running contingencies
   - Note: Lines 09-39A and 09-39B show empty loading - this is normal

### Debug Information

The code provides extensive debug output including:
- Folder structure navigation
- Generator creation attempts
- Power flow execution status
- Contingency analysis progress
- Line loading results

## Technical Details

### PowerFactory Integration

- Uses PowerFactory Python API for automation
- Implements proper object creation and deletion
- Handles PowerFactory-specific attribute names
- Manages study cases and operation scenarios

### Generator Control

- **Voltage Control**: `constv` mode for reactive power management
- **Power Limits**: Calculated based on power factor constraints
- **Reactive Limits**: Automatically calculated from apparent power

### Contingency Analysis

- **Method**: AC Load Flow Calculation
- **Scope**: All transmission line outages
- **Limits**: 110% maximum line loading
- **Convergence**: Handles inconvergence detection

## Contributing

When modifying this project:

1. **Test with small substation sets** first
2. **Verify line duplication** is complete
3. **Check contingency configuration** before running full analysis
4. **Monitor PowerFactory memory usage** during long runs

## License

This project is for educational and research purposes. Please ensure compliance with DIgSILENT PowerFactory licensing terms.

## Support

For issues related to:
- **PowerFactory**: Consult DIgSILENT documentation
- **Python Code**: Check debug output and error messages
- **System Setup**: Verify all setup considerations are met

---

**Note**: This project requires significant PowerFactory project modifications before use. Please ensure all setup considerations are properly implemented to avoid analysis failures.
