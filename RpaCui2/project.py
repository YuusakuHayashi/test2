import os
import yaml

RPA_DIR = os.path.expanduser("~/rpa_project")
YML = RPA_DIR + '/setting.yaml'

def read_yaml():
    if os.path.exists(YML):
        with open(YML, RED, encoding=S_JIS) as ymf:
            data = yaml.safe_load(ymf)
            return data
    else:
        return False
