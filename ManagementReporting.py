import time
import Helpers as helper
# Download excel files from MR
helper.DownloadExcels()
print("Ready to start copying expenses.")
#time.sleep(100)

# Update expenses in pivot sheet
#helper.MoveExpensesData()
#print("Ready to start copying timecards.")
#time.sleep(100)

# Update timesheets in pivot sheet
#helper.MoveTimesheetData()
#time.sleep(100)

# Update data sources on pivot sheets.

# Upload in Box
