# Spn2Pi
This is a script to read xlsm report from SPN server and parse needed data to be used in OSISoft PI database.

# Setup
1. Add config.txt, importShift.txt and log.txt file in directory.

2. In importShift.txt add lines to show mapping of the data being pulled:  
COLUMN_NAME \~ PITAG_NAME \~ COLUMN_NUMBER  
COLUMN_NAME \~ PITAG_NAME \~ COLUMN_NUMBER  
COLUMN_NAME \~ PITAG_NAME \~ COLUMN_NUMBER  


3. In config.txt file add these lines.  
[  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CopyFileSource: "REPORT FILE LOCATION ... MAKE SURE TO USE double slashes for each directory: example:   \\\\167.147.23.45\\D\\Test.xlsm"  
]  

4. Create the scheduled tasks for both Daily and Shift programs. ( Daily = SpnToPi.ps1 and Shift = SpnToPi_Shift.ps1 ).
