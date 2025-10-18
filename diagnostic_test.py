#!/usr/bin/env python3
"""
PowerFactory Diagnostic Test Script

This script performs comprehensive diagnostics on your PowerFactory Python setup
to identify and help fix any issues.
"""

import sys
import os
import traceback
from pathlib import Path

def test_environment():
    """Test Python environment and paths"""
    print("=" * 60)
    print("1. TESTING PYTHON ENVIRONMENT")
    print("=" * 60)
    
    print(f"Python version: {sys.version}")
    print(f"Python executable: {sys.executable}")
    print(f"Current working directory: {os.getcwd()}")
    print(f"Python path: {sys.path[:3]}...")  # Show first 3 entries
    
    return True

def test_dependencies():
    """Test required Python packages"""
    print("\n" + "=" * 60)
    print("2. TESTING PYTHON DEPENDENCIES")
    print("=" * 60)
    
    required_packages = ['pandas', 'numpy', 'matplotlib']
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
            print(f"✅ {package}: Available")
        except ImportError as e:
            print(f"❌ {package}: Missing - {e}")
            missing_packages.append(package)
    
    if missing_packages:
        print(f"\n⚠️  Missing packages: {missing_packages}")
        print("Install with: pip install " + " ".join(missing_packages))
        return False
    
    return True

def test_module_imports():
    """Test local module imports"""
    print("\n" + "=" * 60)
    print("3. TESTING LOCAL MODULE IMPORTS")
    print("=" * 60)
    
    try:
        from src_code import pf_enviroment, initialize_powerfactory
        print("✅ Main module imports: Success")
        
        from src_code import (
            activate_project, list_and_select_study_case, 
            list_and_activate_operation_scenario
        )
        print("✅ Environment functions: Success")
        
        from src_code import (
            create_static_generator, calculate_power_limits, 
            delete_generator, update_generator_power
        )
        print("✅ Generator functions: Success")
        
        from src_code import (
            run_contingency_analysis, process_cargabilidad, 
            optimize_generators_for_substations
        )
        print("✅ Contingency functions: Success")
        
        from src_code import create_results_summary, create_optimization_dashboard
        print("✅ I/O and plotting functions: Success")
        
        return True
        
    except ImportError as e:
        print(f"❌ Module import error: {e}")
        return False

def test_powerfactory_path():
    """Test PowerFactory installation path"""
    print("\n" + "=" * 60)
    print("4. TESTING POWERFACTORY INSTALLATION")
    print("=" * 60)
    
    # Common PowerFactory paths
    possible_paths = [
        r'C:\Program Files\DIgSILENT\PowerFactory 2021 SP2\Python\3.9',
        r'C:\Program Files\DIgSILENT\PowerFactory 2021 SP3\Python\3.9',
        r'C:\Program Files\DIgSILENT\PowerFactory 2022\Python\3.9',
        r'C:\Program Files\DIgSILENT\PowerFactory 2023\Python\3.9',
    ]
    
    valid_paths = []
    for path in possible_paths:
        if os.path.exists(path):
            print(f"✅ Found PowerFactory at: {path}")
            valid_paths.append(path)
        else:
            print(f"❌ Not found: {path}")
    
    if not valid_paths:
        print("\n⚠️  No PowerFactory installation found in common locations.")
        print("Please check your PowerFactory installation path.")
        return False, None
    
    # Test the first valid path
    test_path = valid_paths[0]
    print(f"\nTesting PowerFactory path: {test_path}")
    
    try:
        from src_code.pf_env import pf_enviroment
        pf_enviroment(test_path)
        print("✅ PowerFactory environment setup: Success")
        return True, test_path
    except Exception as e:
        print(f"❌ PowerFactory environment setup failed: {e}")
        return False, test_path

def test_powerfactory_connection(pf_path):
    """Test PowerFactory application connection"""
    print("\n" + "=" * 60)
    print("5. TESTING POWERFACTORY CONNECTION")
    print("=" * 60)
    
    if not pf_path:
        print("❌ No PowerFactory path available for testing")
        return False
    
    try:
        from src_code.pf_env import pf_enviroment, initialize_powerfactory
        
        # Setup environment
        pf_enviroment(pf_path)
        print("✅ Environment setup: Success")
        
        # Try to connect
        app = initialize_powerfactory()
        if app:
            print("✅ PowerFactory connection: Success")
            print(f"✅ Application object: {type(app)}")
            return True
        else:
            print("❌ PowerFactory connection: Failed")
            print("\n🔧 TROUBLESHOOTING STEPS:")
            print("1. Make sure PowerFactory is running")
            print("2. Load your project in PowerFactory")
            print("3. Check PowerFactory license")
            print("4. Try running this script from within PowerFactory's Python console")
            return False
            
    except Exception as e:
        print(f"❌ PowerFactory connection error: {e}")
        print(f"Error details: {traceback.format_exc()}")
        return False

def test_project_access(app):
    """Test project access and configuration"""
    print("\n" + "=" * 60)
    print("6. TESTING PROJECT ACCESS")
    print("=" * 60)
    
    if not app:
        print("❌ No PowerFactory application available")
        return False
    
    try:
        # Test project access
        project = app.GetActiveProject()
        if project:
            print(f"✅ Active project: {project.loc_name}")
        else:
            print("❌ No active project found")
            print("Please load a project in PowerFactory")
            return False
        
        # Test study case access
        study_cases = app.GetProjectFolder('study')
        if study_cases:
            print("✅ Study case folder: Accessible")
            cases = study_cases.GetContents('*.Intcase', 0)
            print(f"✅ Available study cases: {len(cases)}")
            for case in cases[:3]:  # Show first 3
                print(f"   - {case.loc_name}")
        else:
            print("❌ Study case folder: Not accessible")
            return False
        
        # Test operation scenarios
        scenarios = app.GetProjectFolder('scen')
        if scenarios:
            print("✅ Operation scenario folder: Accessible")
            scenario_list = scenarios.GetChildren(1)
            print(f"✅ Available scenarios: {len(scenario_list)}")
            for scenario in scenario_list[:3]:  # Show first 3
                print(f"   - {scenario.loc_name}")
        else:
            print("❌ Operation scenario folder: Not accessible")
            return False
        
        return True
        
    except Exception as e:
        print(f"❌ Project access error: {e}")
        return False

def test_network_data_access(app):
    """Test network data access"""
    print("\n" + "=" * 60)
    print("7. TESTING NETWORK DATA ACCESS")
    print("=" * 60)
    
    if not app:
        print("❌ No PowerFactory application available")
        return False
    
    try:
        project = app.GetActiveProject()
        if not project:
            print("❌ No active project")
            return False
        
        # Test network data access
        network_data = project.GetContents('Network Model.IntPrjfolder\\Network Data', 1)
        if network_data:
            print("✅ Network data folder: Accessible")
            print(f"✅ Network data type: {type(network_data[0])}")
            
            # List available folders
            folders = network_data[0].GetContents('*', 1)
            print(f"✅ Available network folders: {len(folders)}")
            for folder in folders[:5]:  # Show first 5
                print(f"   - {folder.loc_name} ({folder.GetClassName()})")
        else:
            print("❌ Network data folder: Not accessible")
            return False
        
        # Test bus access
        buses = app.GetCalcRelevantObjects('*.ElmTerm')
        print(f"✅ Available buses: {len(buses)}")
        for bus in buses[:5]:  # Show first 5
            print(f"   - {bus.loc_name}")
        
        return True
        
    except Exception as e:
        print(f"❌ Network data access error: {e}")
        return False

def test_generator_creation(app):
    """Test generator creation functionality"""
    print("\n" + "=" * 60)
    print("8. TESTING GENERATOR CREATION")
    print("=" * 60)
    
    if not app:
        print("❌ No PowerFactory application available")
        return False
    
    try:
        from src_code.generators import calculate_power_limits
        
        # Test power calculations
        s, q_max, q_min = calculate_power_limits(10, 0.95)
        print(f"✅ Power calculations: S={s:.2f} MVA, Q_max={q_max:.2f} MVar, Q_min={q_min:.2f} MVar")
        
        # Test bus access for generator creation
        buses = app.GetCalcRelevantObjects('*.ElmTerm')
        if buses:
            test_bus = buses[0]
            print(f"✅ Test bus available: {test_bus.loc_name}")
            print(f"✅ Bus voltage: {test_bus.GetAttribute('m:u'):.4f} pu")
        else:
            print("❌ No buses available for testing")
            return False
        
        return True
        
    except Exception as e:
        print(f"❌ Generator creation test error: {e}")
        return False

def main():
    """Run all diagnostic tests"""
    print("🔍 POWERFACTORY PYTHON DIAGNOSTIC TEST")
    print("=" * 60)
    
    results = {}
    
    # Run all tests
    results['environment'] = test_environment()
    results['dependencies'] = test_dependencies()
    results['imports'] = test_module_imports()
    results['pf_path'], pf_path = test_powerfactory_path()
    results['pf_connection'] = test_powerfactory_connection(pf_path)
    
    app = None
    if results['pf_connection']:
        from src_code.pf_env import initialize_powerfactory
        app = initialize_powerfactory()
    
    results['project_access'] = test_project_access(app)
    results['network_data'] = test_network_data_access(app)
    results['generator_creation'] = test_generator_creation(app)
    
    # Summary
    print("\n" + "=" * 60)
    print("DIAGNOSTIC SUMMARY")
    print("=" * 60)
    
    for test_name, passed in results.items():
        status = "✅ PASS" if passed else "❌ FAIL"
        print(f"{test_name.replace('_', ' ').title()}: {status}")
    
    # Recommendations
    print("\n" + "=" * 60)
    print("RECOMMENDATIONS")
    print("=" * 60)
    
    if not results['pf_connection']:
        print("🔧 CRITICAL: PowerFactory connection failed")
        print("   - Start PowerFactory application")
        print("   - Load your project")
        print("   - Check PowerFactory license")
        print("   - Try running from PowerFactory's Python console")
    
    if not results['project_access']:
        print("🔧 CRITICAL: Project access failed")
        print("   - Load a project in PowerFactory")
        print("   - Check project permissions")
    
    if not results['network_data']:
        print("🔧 WARNING: Network data access issues")
        print("   - Check project structure")
        print("   - Verify network data folder exists")
    
    if all(results.values()):
        print("🎉 ALL TESTS PASSED! Your setup is ready to use.")
        print("   - You can now run: python main.py")
    else:
        print("⚠️  Some tests failed. Please fix the issues above before running the main script.")
    
    return results

if __name__ == "__main__":
    main()
