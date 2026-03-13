FSL Master Roster Builder

Purpose
Build one Excel workbook from many year-by-year roster files.

What the script does
- Searches every supported Excel file inside your input folder and all subfolders.
- Tries to find a header row automatically.
- Pulls these fields when found:
  Last Name, First Name, Banner ID, Email, Status, Semester Joined, Position, Chapter
- Adds:
  Year, Source File, Source Sheet
- Creates one output workbook.
- Splits each year into separate sheets of 1000 rows.
- Adds a Summary sheet with counts and any skipped-file notes.

Supported file types
- .xlsx
- .xlsm
- .xltx
- .xltm

Recommended folder layout
Rosters/
  2012/
    Chapter A.xlsx
    Chapter B.xlsx
  2013/
    yearly_master.xlsx
    another chapter.xlsx
  2014/
    ...

How to run
Open Command Prompt or PowerShell in the folder where the script is saved and run:

python fsl_master_roster_builder.py "C:\Path\To\Rosters" -o "C:\Path\To\Master_FSL_Roster.xlsx"

Example
python fsl_master_roster_builder.py "C:\Users\YourName\Desktop\Rosters" -o "C:\Users\YourName\Desktop\Master_FSL_Roster.xlsx"

Optional settings
--chunk-size 1000
  Number of rows per year sheet.

--keep-duplicates
  Keeps exact duplicate rows instead of removing them.

--verbose
  Prints each workbook as it is processed.

Example with options
python fsl_master_roster_builder.py "C:\Users\YourName\Desktop\Rosters" -o "C:\Users\YourName\Desktop\Master_FSL_Roster.xlsx" --chunk-size 1000 --verbose

Important notes
- The script identifies years from folder names like 2012, 2013, 2014.
- If a roster has no Chapter column, the script tries to infer chapter from the file or parent-folder name.
- Old .xls files are not supported by this version.
- If some files use unusual headers, those sheets may be skipped and listed on the Summary sheet.
- Banner IDs are kept exactly as text when read from the source workbook.

Standardized output columns
Year
Source File
Source Sheet
Chapter
Last Name
First Name
Banner ID
Email
Status
Semester Joined
Position
