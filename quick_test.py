#!/usr/bin/env python3
"""
Quick PowerFactory Connection Test
Run this after starting PowerFactory and loading your project
"""

def test_connection():
    try:
        from src_code.pf_env import pf_enviroment, initialize_powerfactory
        
        # Setup PowerFactory environment
        dig_path = r'C:\Program Files\DIgSILENT\PowerFactory 2021 SP2\Python\3.9'
        pf_enviroment(dig_path)
        
        # Try to connect
        app = initialize_powerfactory()
        
        if app:
            print("🎉 SUCCESS! PowerFactory connection established!")
            
            # Test project access
            project = app.GetActiveProject()
            if project:
                print(f"✅ Project loaded: {project.loc_name}")
                
                # Test study cases
                study_cases = app.GetProjectFolder('study')
                cases = study_cases.GetContents('*.Intcase', 0)
                print(f"✅ Study cases available: {len(cases)}")
                
                # Test buses
                buses = app.GetCalcRelevantObjects('*.ElmTerm')
                print(f"✅ Buses available: {len(buses)}")
                
                print("\n🚀 Your setup is ready! You can now run:")
                print("   python main.py")
                
                return True
            else:
                print("❌ No project loaded. Please load your project in PowerFactory.")
                return False
        else:
            print("❌ Failed to connect to PowerFactory.")
            print("Make sure PowerFactory is running and your project is loaded.")
            return False
            
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

if __name__ == "__main__":
    print("🔍 Testing PowerFactory Connection...")
    print("Make sure PowerFactory is running and your project is loaded!")
    print("-" * 50)
    
    success = test_connection()
    
    if not success:
        print("\n🔧 TROUBLESHOOTING:")
        print("1. Start PowerFactory application")
        print("2. Load your project (39 Bus New England System)")
        print("3. Check PowerFactory license")
        print("4. Try running this script again")
