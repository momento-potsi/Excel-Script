from SystemInfo import CURRENT_PATH, BUILD_MODE
from UnitTests import StylingUnitTest, DataEntryTest

import platform

# test = StylingUnitTest() # Test Complete
test = DataEntryTest()

def main():
    print("[Running Excel Script] <Build Mode: " + BUILD_MODE + "> ...")
    
    if BUILD_MODE == 'Debug':
        print("[Program Info { deviceName: `" + platform.node() + "`, path: `" + CURRENT_PATH + "` }")
    
    # test.colorStyling() # Test Complete
    # test.entryInstanceTest() # Test Complete
    # test.excelSheetEditTest() # Test Complete
    test.excelSheetUpdateTest()
    print("Done!")
main()
