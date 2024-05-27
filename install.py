import subprocess
import sys
import os

def create_venv():
    # Check if the virtual environment folder already exists
    if not os.path.exists('venv'):
        # Create the virtual environment
        subprocess.check_call([sys.executable, '-m', 'venv', 'venv'])
        print("Virtual environment created.")
    else:
        print("Virtual environment already exists.")

def install_requirements():
    # Activate the virtual environment
    activate_script = 'venv\\Scripts\\activate_this.py' if os.name == 'nt' else 'venv/bin/activate_this.py'
    exec(open(activate_script).read(), {'__file__': activate_script})

    # Install requirements
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', '-r', 'requirements.txt'])
    print("Requirements installed.")

if __name__ == "__main__":
    create_venv()
    install_requirements()