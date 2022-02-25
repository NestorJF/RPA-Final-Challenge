import os 

TEMP_FOLDER = os.path.join(os.environ.get("ROBOT_ROOT", os.getcwd()), 'input')
OUTPUT_FOLDER = os.path.join(os.environ.get("ROBOT_ROOT", os.getcwd()), 'output')
tabs_dict = {}