<#
    .SYNOPSIS
    Get all software for a computer

    .DESCRIPTION
    Get all software for a computer with multiple options to
        return objects or export to a file.

    .RETURNS
    Network adapter configuration objects from the remote machine or
    A DataTable of all configuration information

    .OUTPUTS
    An HTML of the DataTable (the default) (which you can manually change the 
        extension to .xls and open in MS Excel for manipulation)
    An XML file
    A CSV file

    .PARAMETER Computers
    Array:  List of computer names to log into via PS Remoting

    .PARAMETER OutputOriginalObjects
    Boolean: Output the network adapter configuration objects to the pipeline

    .PARAMETER OutputDataTable
    Boolean: Output a DataTable of the network adapter configuration objects to the pipeline
    
    .PARAMETER ExportHtml
    Boolean: Export all data to an HTML file

    .PARAMETER ExportXml
    Boolean: Export all data to an XML file (preserving objects for later import+use)

    .PARAMETER ExportCsv
    Boolean: Export all data to an CSV file

    .PARAMETER IncludeDetails
    Boolean: Include Install/Uninstall string information

    .PARAMETER OutputFileName
    String:  File name for output.  This can be short or full name format.
    Do NOT add an extension as the correct one will be added before export.

    .NOTES
    This is intended to be run over PS Remoting.
    Booleans were used instead of switches so the true/false can be set for defaults
    Multiple output file types can be specified simultaneously

    Author: Donald Hess
    Version History:
        2.0    2018-04-23    Added OutputFileName, ExportCsv, help file info
        1.1    2017-10-27    Added switches to provide different return options
        1.0    2017-08-08    Initial release
    
    .EXAMPLE
    Get-InstalledSoftware -Computers 'comp1','comp2','etc' -ExportHtml $true `
        -OutputFileName 'myfilename'
    Runs the scriptblock remotely and outputs a 'myfilename.html' file on the local 
    machine in the current directory

    .EXAMPLE
    Get-InstalledSoftware -Computers 'comp1','comp2','etc' -ExportHtml $true `
        -ExportCvs $true -OutputFileName 'C:\full\path\to\myfilename'
    Runs the scriptblock remotely and outputs 'C:\full\path\to\myfilename.html' 
    and 'C:\full\path\to\myfilename.csv' files on the local machine
#>

param ( [array] $Computers = @(),
        [bool] $OutputOriginalObjects = $false,
        [bool] $OutputDataTable = $false,
        [bool] $ExportHtml = $true,
        [bool] $ExportXml = $false,
        [bool] $ExportCsv = $false,
        [bool] $IncludeDetails = $false,  # Include additional software details that normally would not be needed
        [string] $OutputFileName = 'software_results'  # No Extension, can be short or full path
      )

Set-StrictMode -Version latest -Verbose
$ErrorActionPreference = 'Stop'
$PSDefaultParameterValues['*:ErrorAction']='Stop'

$sb1 = {
    param ( [bool] $IncludeDetails = $false )
    function funcGetInstalledVersions() { 
        $sCompName = $env:COMPUTERNAME
        Write-Host 'Working on:' $sCompName

        if ( $IncludeDetails ) {
            filter filtSelectDetails{
                $_ | Select-Object DisplayName,DisplayVersion,Publisher,InstallDate,PSChildName,InstallLocation,UninstallString
            }
        } else {
            filter filtSelectDetails{
                $_ | Select-Object DisplayName,DisplayVersion,Publisher,InstallDate,PSChildName
            }
        }
        # Check for 64 bit program on 64 bit OS, also covers 32 bit program on 32 bit OS.
        # This will check the Uninstall reg keys for programs that wildcard match the $sMatch variable.
        $sBitness = "native"
        $aNative = @( Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall | `
                        ForEach-Object { Get-ItemProperty $_.PSPath } | filtSelectDetails | `
                        # Add the program bitness before we return it. Note $_ here is the Get-ItemProperty $_.PSPath from above.
                        # Not sure why, but we must use Write-Output $_; in front of the Add-Member for this to work.
                        ForEach-Object -Process { Write-Output $_; 
                            Add-Member -InputObject $_ -MemberType NoteProperty -Name Bitness -Value $sBitness; 
                            Add-Member -InputObject $_ -MemberType NoteProperty -Name CompName -Value $sCompName;
                        }
                    )
        # Check for 32 bit program on 64 bit OS
        # This will check the Uninstall reg keys for programs that wildcard match the $sMatch variable.
        # Must check for the full path because some software creates Wow6432Node on 32 bit systems
        if ( Test-Path -Path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall' ) {
            $sBitness = "32on64"
            $a32on64 = @( Get-ChildItem 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall' | `
                            ForEach-Object { Get-ItemProperty $_.PSPath } | filtSelectDetails | `
                            # Add the program bitness before we return it. Note $_ here is the Get-ItemProperty $_.PSPath from above.
                            # Not sure why, but we must use Write-Output $_; in front of the Add-Member for this to work.
                            ForEach-Object -Process { Write-Output $_; 
                                Add-Member -InputObject $_ -MemberType NoteProperty -Name Bitness -Value $sBitness;
                                Add-Member -InputObject $_ -MemberType NoteProperty -Name CompName -Value $sCompName;
                            }
                        )
        } # End if Test-Path
        else {  
            # We are only on 32 bit.  Create an empty array so it matches $aNative
            $a32on64 = @()
        }
        # Prep to check if empty array, sort through the arrays and combine if needed.
        # Sort with newest version at the top
        if ( $aNative.Length -gt 1 ) {
            $aNative = ( $aNative | Sort-Object -Property DisplayVersion -Descending )
        }
        if ( $a32on64.Length -gt 1 ) {
            $a32on64 = ( $a32on64 | Sort-Object -Property DisplayVersion -Descending )
        }
        $aNative = $aNative += $a32on64
        $aNative = $aNative | Sort -Property DisplayName,PSChildName,Bitness
        # If we don't have anything by this point, return 0
        if ( $aNative.Length -eq 0 ) {
            $aNative = @()
        }
        Return ,$aNative # Make sure to wrap so we keep the array
    } # End of funcGetInstalledVersions
    funcGetInstalledVersions  # Returning each object, not array
} # End scriptblock

function funcConvertTo-DataTable {
    <#  .SYNOPSIS
            Convert regular PowerShell objects to a DataTable object.
        .DESCRIPTION
            Convert regular PowerShell objects to a DataTable object.
        .EXAMPLE
            $myDataTable = $myObject | ConvertTo-DataTable
        .NOTES
            Name: ConvertTo-DataTable
            Author: Oyvind Kallstad @okallstad
            Version: 1.1
    #>
    [CmdletBinding()]
    param (
        # The object to convert to a DataTable
        [Parameter(ValueFromPipeline = $true)]
        [PSObject[]] $InputObject,

        # Override the default type.
        [Parameter()]
        [string] $DefaultType = 'System.String'
    )
    begin {
        # Create an empty datatable
        try {
            $dataTable = New-Object -TypeName 'System.Data.DataTable'
            Write-Verbose -Message 'Empty DataTable created'
        } catch {
            Write-Warning -Message $_.Exception.Message
            break
        }
        # Define a boolean to keep track of the first datarow
        $first = $true
        # Define array of supported .NET types
        $types = @(
            'System.String',
            'System.Boolean',
            'System.Byte[]',
            'System.Byte',
            'System.Char',
            'System.DateTime',
            'System.Decimal',
            'System.Double',
            'System.Guid',
            'System.Int16',
            'System.Int32',
            'System.Int64',
            'System.Single',
            'System.UInt16',
            'System.UInt32',
            'System.UInt64'
        )
    }
    process {
        # Iterate through each input object
        foreach ($object in $InputObject) {
            try {
                # Create a new datarow
                $dataRow = $dataTable.NewRow()
                Write-Verbose -Message 'New DataRow created'
                # Iterate through each object property
                foreach ($property in $object.PSObject.get_properties()) {
                    # Check if we are dealing with the first row or not
                    if ($first) {
                        # handle data types
                        if ($types -contains $property.TypeNameOfValue) {
                            $dataType = $property.TypeNameOfValue
                            Write-Verbose -Message "$($property.Name): Supported datatype <$($dataType)>"
                        } else {
                            $dataType = $DefaultType
                            Write-Verbose -Message "$($property.Name): Unsupported datatype ($($property.TypeNameOfValue)), using default <$($DefaultType)>"
                        }
                        # Create a new datacolumn
                        $dataColumn = New-Object 'System.Data.DataColumn' $property.Name, $dataType
                        Write-Verbose -Message 'Created new DataColumn'

                        # Add column to DataTable
                        $dataTable.Columns.Add($dataColumn)
                        Write-Verbose -Message 'DataColumn added to DataTable'
                    }                  
                    # Add values to column
                    if ($property.Value -ne $null) {
                        # If array or collection, add as XML
                        if (($property.Value.GetType().IsArray) -or ($property.TypeNameOfValue -like '*collection*')) {
                            $dataRow.Item($property.Name) = $property.Value | ConvertTo-Xml -As 'String' -NoTypeInformation -Depth 1
                            Write-Verbose -Message 'Value added to row as XML'
                        } else {
                            $dataRow.Item($property.Name) = $property.Value -as $dataType
                            Write-Verbose -Message "Value ($($property.Value)) added to row as $($dataType)"
                        }
                    }
                }
                # Add DataRow to DataTable
                $dataTable.Rows.Add($dataRow)
                Write-Verbose -Message 'DataRow added to DataTable'
                $first = $false
            } catch {
                Write-Warning -Message $_.Exception.Message
            }
        }
    }
    end { Write-Output (,($dataTable)) }
} # End funcConvertTo-DataTable

if ( $Computers.count -ne 0 ) {
    # Something passed in
    $aComputers = $Computers
} else {
    # This will filter so we get only computer object that are backed by a real workstation
    $sDomainSnip = (Get-WmiObject Win32_ComputerSystem).Domain.Trim().Substring(0,4)
    $aComputers = Get-ADComputer -Filter * | Select-Object Name,DNSHostName | Where-Object { $_.Name -Like "$sDomainSnip*" -and $null -ne $_.DNSHostName } | ForEach-Object { $_.Name } | Sort
}

$aResults = @(Invoke-Command -ThrottleLimit 1 -ScriptBlock $sb1 -ComputerName $aComputers -ArgumentList $IncludeDetails -ErrorAction Continue)
if ( $OutputOriginalObjects ) {
    $aResults  # Returning each object, not array
}
if ( $OutputDataTable -or $ExportHtml -or $ExportCsv ) {
    $aDtResults = @()
    $aResults | ForEach-Object {
        $aDtResults += @(funcConvertTo-DataTable $_)
    }
}
if ( $OutputDataTable ) {
    $aDtResults
}
if ( $ExportXml ) {
    Export-Clixml -Depth 10 -InputObject $aResults -Path (@($OutputFileName,'.xml') -join '')
}
if ( $ExportHtml ) {
    #HTML output
    $aHtmlContent = @()
    $aHtmlContent += '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
    <html xmlns="http://www.w3.org/1999/xhtml"><head><title>HTML TABLE</title>
    <style type="text/css">
    table {
	    border: thin solid lightgray;
	    border-collapse: collapse;
    }
    td {
	    border: thin solid lightgray;
	    padding-left: 10px;
	    padding-right: 10px;
    }
    </style></head><body>'
    $aDtResults | ForEach-Object {
        $aHtmlContent += ($_ | Select * -ExcludeProperty RowError, RowState, HasErrors, Name, Table, ItemArray | ConvertTo-Html -Fragment) -join ''
    }
    $aHtmlContent += '</body></html>'
    $aHtmlContent -join '</br><hr></br>' > (@($OutputFileName,'.html') -join '')
}
if ( $ExportCsv ) {
    $aDtResults | ForEach-Object { 
        $oDataTable = $_
        $oDataTable | Export-Csv -NoTypeInformation -Append  -Path (@($OutputFileName,'.csv') -join '')
    }
}

