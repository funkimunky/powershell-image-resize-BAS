# powershell image resize BAS
This is a powershell script for batch resizing images on a windows server. This is intended to be run as a scheduled task. The application will use a folder URI list from a csv.

To start run 
Import-Module -Name .\modules\Import-Excel\ImportExcel -Verbose 

from root script process_files.ps1

Then run process_files.ps1