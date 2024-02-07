
from UnitTests import StylingUnitTest, DataEntryTest

# test = StylingUnitTest() # Test Complete
test = DataEntryTest()

def main():
    print("[Testing]...")
    
    # test.colorStyling() # Test Complete
    # test.entryInstanceTest() # Test Complete
    # test.excelSheetEditTest() # Test Complete
    test.excelSheetUpdateTest()
    print("Done!")
main()
