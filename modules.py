#Use this scrip to install the python modules

import sys
import subprocess

subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'requests'])
subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'openpyxl'])
subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'xlsxwriter'])
subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pandas'])

