# run_app.py
import os
import sys
import streamlit.web.cli as stcli

if __name__ == '__main__':
    # Locate the absolute system paths of your app.py file dynamically
    script_path = os.path.join(os.path.dirname(__file__), 'app.py')
    
    # Force the local python sub-thread to inject 'streamlit run app.py' parameters safely
    sys.argv = ["streamlit", "run", script_path, "--global.developmentMode=false"]
    
    # Initialize the engine gate handle
    sys.exit(stcli.main())
