# This script is for Excel Data Extraction. 
# If it works, It was written by Tyler
# If not, I don't know who wrote it.

# This code could use some additional optimization and updating however it is fully functional in it's current state.

# Some metrics: This script is capable of looping through several thousand excel files in a relatively short amount of time
# I was able to process 700 excel sheets in under an hour and while using less than 300MB of memory. (On a 2011 (3.2G) i7 CPU)

# ================
# Global Variables
# ================

$NewTable = New-Object System.Collections.ArrayList # Master list for the new table, had to use an arraylist since arrays can't hold arrays
$CurrentFile = "" # Current working file, prints this out while its processing

$FieldsOfInterest = Import-Csv "Terms.csv"  # Answer
$FirstRun = $true # Lets the script know if it is the first excel file or not
$ExtractSheet = "Sheet2"
# !!! The directory argument is at the bottom, you need it in two spots (for now)

# =========
# Functions
# =========

# Get data, this function calls to Access Database Engine and grabs the data from the excel file and returns the result of the entire file
function Get-ExcelData {

# I did not write this section, most was shamelessly borrowed from technet
# https://blogs.technet.microsoft.com/pstips/2014/06/02/get-excel-data-without-excel/
# which is why its not documented as well. Future comments to be made!!!

    [CmdletBinding(DefaultParameterSetName='Worksheet')]
    Param(
        [Parameter(Mandatory=$true, Position=0)]
        [String] $Path,

        [Parameter(Position=1, ParameterSetName='Worksheet')]
        [String] $WorksheetName = 'Sheet1',

        [Parameter(Position=1, ParameterSetName='Query')]
        [String] $Query = 'SELECT * FROM [Sheet1$]'
    )

    switch ($pscmdlet.ParameterSetName) {
        'Worksheet' {
            $Query = 'SELECT * FROM [{0}$]' -f $WorksheetName
            break
        }
        'Query' {
            # Make sure the query is in the correct syntax (e.g. 'SELECT * FROM [SheetName$]')
            $Pattern = '.*from\b\s*(?<Table>\w+).*'
            if($Query -match $Pattern) {
                $Query = $Query -replace $Matches.Table, ('[{0}$]' -f $Matches.Table)
            }
        }
    }

    # Create the scriptblock to run in a job
    $JobCode = {
        Param($Path, $Query)

        # Check if the file is XLS or XLSX 
        if ((Get-Item -Path $Path).Extension -eq 'xls') {
            $Provider = 'Microsoft.Jet.OLEDB.4.0'
            $ExtendedProperties = 'Excel 8.0;HDR=YES;IMEX=1'
        } else {
            $Provider = 'Microsoft.ACE.OLEDB.12.0'
            $ExtendedProperties = 'Excel 12.0;HDR=YES'
        }
        
        # Build the connection string and connection object
        $ConnectionString = 'Provider={0};Data Source={1};Extended Properties="{2}"' -f $Provider, $Path, $ExtendedProperties
        $Connection = New-Object System.Data.OleDb.OleDbConnection $ConnectionString

        try {
            # Open the connection to the file, and fill the datatable
            $Connection.Open()
            $Adapter = New-Object -TypeName System.Data.OleDb.OleDbDataAdapter $Query, $Connection
            $DataTable = New-Object System.Data.DataTable
            $Adapter.Fill($DataTable) | Out-Null
        }
        catch {
            # something went wrong ??
            Write-Error $_.Exception.Message
        }
        finally {
            # Close the connection
            if ($Connection.State -eq 'Open') {
                $Connection.Close()
            }
        }

        # Return the results as an array
        return ,$DataTable
    }

    # Run the code in a 32bit job, since the provider is 32bit only
    $job = Start-Job $JobCode -RunAs32 -ArgumentList $Path, $Query
    $job | Wait-Job | Receive-Job
    Remove-Job $job
}

# This function loops through the data returned by the Get-ExcelData function line by line
function Get-Fields() {

	# Loads in the temp data line by line so it's easier to work with below.
	Write-Host "Bringing them back in..."
	$WorkingTable = Get-Content temp.txt

	# Some variables to get us going
	$LinePosition = 0
	$NewRow = @()
	$Indexes = @()

	# This section Pre-Loads the DisplayName (The column labels) into the sheet on the first run
	if($FirstRun -eq $true)
	{
		foreach ($term in $FieldsOfInterest)
		{
			$NewRow += $term.DisplayName # Grabs every display name and adds it to the top of the sheet
		}

		$NewTable.Add($NewRow) # Adds the row to the table
		$NewRow = @()
	}

	$FirstRun = $false

	# The main ForEach loop, this does the majority of the work
	# It loops through, finds the indexes of all the references and then
	# dumps that to an array listed above
	Write-Host "Here we goooo..."

	foreach ($line in $WorkingTable) # Loops through grabbing each line from Working Table
	{
		$RowCap = @() # New array for our additions

		# We're just band-aiding this for now, the foreach loop below runs three times because we cannot afford
		# the time to do it properly (with a function etc. See Sibyl once complete).

		foreach ($term in $FieldsOfInterest) # Loops every search term against the lines above
		{			
			# compares the two, just would like to point out the -contains and -like were worthless and it took me a minute to figure that out
			if(($line.ToString() -match $term.Term.ToString()))
			{
				Write-Host "Match found!" # Notify User
				Write-Host $LinePosition # Dump line to console
				$Indexes += $LinePosition + $term.Index # Adds the index to the list of indexes to grab, + Term[1] is there cause we want the OFFSET line not the current one
			}
		}

		$LinePosition++
	}

	# This part grabs all the data and builds the row
	foreach($index in $Indexes)
	{
		$indexDelim = $WorkingTable[$index].IndexOf(":") # Finds the index of the deliminator on that line
		$substring = $WorkingTable[$index].Substring($indexDelim + 1) # Isolates the data from the rest of the line (Plus 1 gets rid of that space char)
		$NewRow += $substring # Adds the data to the new row being built
	}

	Write-Host $NewRow # Writes the row to the console for debugging

	$NewRow += $CurrentFile # Adds the file name to the data for troubleshooting

	$NewTable.Add($NewRow) # Adds the row to the table

	Write-Host $NewTable.count # Shows the number it has processed so far

	$FirstRun = $false
}

# This function facilitates getting the data from Get-ExcelData and gets it in a temporary format usable by the above
function Get-DataFromSheet($path){

	# Removes Old Data (And lets the user know)
	Write-Host "Removing old temp files..."
	If(Test-Path temp.txt)
	{
		Remove-Item temp.txt
	}

	# Dumps the result of Get-ExcelData to a text file. I'm sure someone more versed than I in powershell could get it into an array
	# but none of my attampts to get it into an Array worked. The array always ended up empty so this seemed an easier solution
	Write-Host "Dumping.... To temp file"
	Get-ExcelData -Path $path -WorksheetName $ExtractSheet >> temp.txt
}


# ===================================
# Main Code Area (Think void Main)
# ===================================

# Directory and introduction!
Write-Host "Welcome to the Excel Data Extraction Tool!"
Write-Host "To get started, please enter the directory below. `n`n"
$DirectoryToUse = Read-Host "Directory: > "

if(Test-Path $DirectoryToUse)
{
	Write-Host "All set, here we go! `n`n"
}
else
{
	Write-Host "Bad Path! `n Try something like C:\ExcelFiles\"
	Read-Host "Nothing to do, please press enter!"
	exit
}

$ListOfFiles = Get-ChildItem $DirectoryToUse -Filter *.xlsx

# Loops through the files, running the above on all of them
foreach ($file in $ListOfFiles)
{
	Write-Host $file # Writes the file being processed
	$dfsFile = $DirectoryToUse + "\" + $file # Appends the path to the file
	$CurrentFile = $dfsFile # Worthless, remove
	Get-DataFromSheet $dfsFile # Gets the data from the sheet (Removing the temp file in the process: defined above)
	Get-Fields # Gets the fields, takes nothing in since it gets it from the temp file
	$FirstRun = $false
}

# This part builds the CSV. It typically takes a minute
Write-Host "Building the CSV, could take a while depending on how many files we found"
# Found this gem on stack sxchange. Pipes the arraylist through a join then outfile
$NewTable | % { $_ -join '|'} |  out-file .\export.csv  # You can change the deliminator by changing the join statement
Write-Host "COMPLETE!!! `n`n"
$halp = Read-Host "Press Enter!!!" # Just keeps the console open after completion

# After hitting enter, it deletes the temp file again for a completely clean execution
Write-Host "Removing old temp files..."
If(Test-Path temp.txt)
{
	# Remove-Item temp.txt # Comment this line out for first run so you can make offsets.
}