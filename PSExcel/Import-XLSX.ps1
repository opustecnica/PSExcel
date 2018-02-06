function Import-XLSX {
    <#
    .SYNOPSIS
        Import data from Excel

    .DESCRIPTION
        Import data from Excel

    .PARAMETER Path
        Path to an xlsx file to import

    .PARAMETER Sheet
        Index or name of Worksheet to import

    .PARAMETER Header
        Replacement headers.  Must match order and count of your data's properties.
    
    .PARAMETER DateTimeFormat
        This format is used to detect datetime values within the excel sheet.
    
    .PARAMETER RowStart
        First row to start reading from, typically the header. Default is 1

    .PARAMETER ColumnStart
        First column to start reading from. Default is 1

    .PARAMETER FirstRowIsData
        Indicates that the first row is data, not headers.  Must be used with -Header.

    .PARAMETER Text
        Extract cell text, rather than value.

        For example, if you have a cell with value 5:
            If the Number Format is '0', the text would be 5
            If the Number Format is 0.00, the text would be 5.00 

    .EXAMPLE
        Import-XLSX -Path "C:\Excel.xlsx"

        # Import data from C:\Excel.xlsx

    .EXAMPLE
        Import-XLSX -Path "C:\Excel.xlsx" -ReadOnly

        # Import data from C:\Excel.xlsx even when the file has already been opened by a 
        # different application/user

    .EXAMPLE
        Import-XLSX -Path "C:\Excel.xlsx" -Header One, Two, Five

        # Import data from C:\Excel.xlsx
        # Replace headers with One, Two, Five

    .EXAMPLE
        Import-XLSX -Path "C:\Excel.xlsx" -Header One, Two, Five -FirstRowIsData -Sheet 2

        # Import data from C:\Excel.xlsx
        # Assume first row is data
        # Use headers One, Two, Five
        # Pull from sheet 2 (sheet 1 is default)

    .EXAMPLE
       #    A        B        C 
       # 1  Random text to mess with you!
       # 2  Header1  Header2  Header3
       # 3  data1    Data2    Data3

       # Your worksheet has data you don't care about in the first row or column
       # Use the ColumnStart or RowStart parameters to solve this.

       Import-XLSX -Path C:\RandomTextInRow1.xlsx -RowStart 2

    .NOTES
        Thanks to Doug Finke for his example:
            https://github.com/dfinke/ImportExcel/blob/master/ImportExcel.psm1

        Thanks to Philip Thompson for an expansive set of examples on working with EPPlus in PowerShell:
            https://excelpslib.codeplex.com/

        [OpusTecnica] - Added ReadOnly option to access files that are already open.
        [OpusTecnica] - Added native excel column headers when the first row is data.
                        This is limited to columns A to Z at the moment, but can be easily expanded.  
        [OpusTecnica] - Some reformatting for personal preference.
        [OpusTecnica] - Added trimming of the headers to account for superfluous white spaces.
        [OpusTecnica] - Added ColumnEnd Parameter.

    .LINK
        https://github.com/RamblingCookieMonster/PSExcel

    .FUNCTIONALITY
        Excel
    #>
    [cmdletbinding()]
    param(
        [parameter( Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [validatescript({Test-Path $_})]
        [string[]]$Path,

        $Sheet = 1,

        [string[]]$Header,

        [switch]$FirstRowIsData,

        [switch]$Text,
        
        [string]$DateTimeFormat = "M/d/yyy h:mm",

        [int]$RowStart = 1,

        [int]$ColumnStart = 1,
        
        [int]$ColumnEnd = $null,

        [switch]$ReadOnly
    )
    
    Begin {
        [string[]]$Alphabet = [char[]]([int][char]'A'..[int][char]'Z')
    }
    Process {
        foreach($File in $Path) {
            #Resolve relative paths... Thanks Oisin! http://stackoverflow.com/a/3040982/3067642
            $File = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($File)

            write-verbose "Target excel file $($File)"

            Try {
                if ($ReadOnly) {
                    Write-Verbose "Opening $($File) as Read-Only"
                    $ReadOnlyFile = [IO.File]::Open($File,'Open','Read','ReadWrite')
                    $ExcelFile = New-Object OfficeOpenXml.ExcelPackage
                    $ExcelFile.Load($ReadOnlyFile)
                    $workbook  = $ExcelFile.Workbook
                } else {
                    $ExcelFile = New-Object OfficeOpenXml.ExcelPackage $File
                    $Workbook  = $ExcelFile.Workbook
                }
            }

            Catch {
                Write-Error "Failed to open '$File':`n$_"
                continue
            }

            Try {
                if( @($Workbook.Worksheets).count -eq 0) {
                    Throw "No worksheets found"
                } else {
                    $Worksheet = $Workbook.Worksheets[$Sheet]
                    $Dimension = $Worksheet.Dimension
                    Write-Verbose "WORKSHEET: $Worksheet"
                    Write-Verbose "DIMENSIONS: $Dimension"

                    $Rows = $Dimension.Rows
                    $Columns = $Dimension.Columns
                    Write-Verbose "ROWS: $Rows"
                    Write-Verbose "COLUMNS: $Columns"

                    # [opustecnica] Add ColumnEnd User Input Parameter
                    if ( $ColumnEnd ) { $Columns = $ColumnEnd }                   
                    $ColumnEnd = $Columns + $ColumnStart - 1
                    $RowEnd = $Rows + $RowStart - 1
                    Write-Verbose "LAST COLUMN: $ColumnEnd"
                    Write-Verbose "LAST ROW: $RowEnd"
                } 
            }

            Catch {
                Write-Error "Failed to gather Worksheet '$Sheet' data for file '$File':"
                Write-Error $_
                continue
            }
  
            if($Header -and $Header.count -gt 0) {
                if($Header.count -ne $Columns) {
                    Write-Error "Found '$Columns' columns, provided $($Header.count) headers.  You must provide a header for every column."
                }
                Write-Verbose "User defined headers: $Header"
            } else {
                $Header = @( foreach ($Column in $ColumnStart..$ColumnEnd) {
                    if($Text) {
                        $PotentialHeader = $Worksheet.Cells.Item($RowStart,$Column).Text
                    } else {
                        # Handle Native Column Headers in absence of headers.
                        if ( $FirstRowIsData ) {
                            if ( $Column -le $Alphabet.Count ) {
                                $PotentialHeader = $Alphabet[$Column - 1]
                            } else {
                                $PotentialHeader = $Alphabet[[int]($Column / $Alphabet.Count) - 1] + $Alphabet[[int]($Column % $Alphabet.Count) - 1]
                            } # End if FirstRowIsData
                        } else {
                            $PotentialHeader = $Worksheet.Cells.Item($RowStart,$Column).Value
                        }
                    }

                    # [opustecnica] if( -Not $PotentialHeader -Or $PotentialHeader.Trim().Equals("") ) # Produces Error if data has no header.
                    if ( -Not $PotentialHeader -Or ($PotentialHeader -isnot [string]) -or $PotentialHeader.Trim().Equals("")) {
                        Write-Verbose "Header in column $Column is whitespace or empty, setting header to '<Column $Column>'"
                        $PotentialHeader = "<Column $Column>" # Use placeholder name
                    }
                    $PotentialHeader.Trim()
                }) # End Header =
            }

            [string[]]$SelectedHeaders = @( $Header | select -Unique )
            Write-Verbose "Found $Rows rows, $Columns columns, with headers:`n$($Header | Out-String)"

            if (-not $FirstRowIsData) { $RowStart++ }

            foreach ($Row in $RowStart..$RowEnd) {
                $RowData = @{}

                foreach ($Column in 0..($Columns - 1)) {
                    $Name  = $Header[$Column]
                    if ($Text) { 
                        $Value = $Worksheet.Cells.Item($Row, ($Column + $ColumnStart)).Text
                    } else {
                        $Value = $Worksheet.Cells.Item($Row, ($Column + $ColumnStart)).Value
                    }

                    Write-Debug "Row: $Row, Column: $Column, Name: $Name, Value = $Value"

                    # Handle dates, they're too common to overlook... Could use help, not sure if this is the best regex to use?
                    $Format = $Worksheet.Cells.Item($Row, ($Column + $ColumnStart)).style.numberformat.format
                    if ($Format -match '\w{1,4}/\w{1,2}/\w{1,4}( \w{1,2}:\w{1,2})?' -or $Format -match $DateTimeFormat) {
                        Try {
                            $Value = [datetime]::FromOADate($Value)
                        }
                        Catch {
                            Write-Verbose "Error converting '$Value' to datetime"
                        }
                    }

                    if ( $RowData.ContainsKey($Name) ) {
                        Write-Warning "Duplicate header for '$Name' found, with value '$Value', in row $Row"
                    } else {
                        $RowData.Add($Name, $Value)
                    }
                }
                New-Object -TypeName PSObject -Property $RowData | Select -Property $SelectedHeaders
            }

            $ExcelFile.Dispose()
            $ExcelFile = $null
        }
    }
}
