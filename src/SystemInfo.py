
import platform

# Platform Dependencies (Temporary)
DEFAULT_PATH_WINDOWS = "D:\\"
DEFAULT_PATH_WINDOWS_WSL = "/mnt/D/"
DEFAULT_PATH_LINUX = "/home/tosin/"
# End Platform

CURRENT_PATH = None
BUILD_MODE = 'Debug'

match platform.node():
    case 'cardinal-sys22' :
        CURRENT_PATH = DEFAULT_PATH_WINDOWS_WSL
    case 'workstation':
        CURRENT_PATH = DEFAULT_PATH_LINUX