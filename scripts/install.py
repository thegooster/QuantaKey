import os
import subprocess
import sys

def create_and_activate_venv():
    # Create a virtual environment named 'venv' in the current directory
    subprocess.call([sys.executable, "-m", "venv", "venv"])
    
    # Activate the virtual environment
    # For Windows
    if os.name == 'nt':
        activate_script = os.path.join('.', 'venv', 'Scripts', 'activate')
    # For Unix or MacOS
    else:
        activate_script = os.path.join('.', 'venv', 'bin', 'activate')
    
    # Execute the activate script
    activate_command = f"source {activate_script}"
    os.system(activate_command)

def install_requirements():
    # Install packages from requirements.txt
    subprocess.call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])

def run_main():
    # Run the main.py script
    subprocess.call([sys.executable, "main.py"])

if __name__ == "__main__":
    create_and_activate_venv()
    install_requirements()
    run_main()