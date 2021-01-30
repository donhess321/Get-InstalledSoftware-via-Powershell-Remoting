# Get-InstalledSoftware-via-Powershell-Remoting

This will use PS Remoting to get a machine's installed software and output to multiple formats.  Input is an array of computer names.  If no computer name is specified, a list of all computers in the domain is created.  The software inventory is retrieved from the registry as this will return programs that do not show up through the usual MSI installation process.  The results for the installed software or a DataTable version of the results can be returned so the machine originating PS Remoting can further process them.  File output for all remote machines is to a single file for each filetype on the machine originating PS Remoting and can be to an HTML, XML, or CSV text file.  The HTML or CSV file can be opened in Excel for further processing.  The XML file allows for storing the results in a way to recreate the original objects at a later date.  You can also increase the throttle limit on Invoke-Command as it currently is set to sequencially connect to each machine.  
 
 This is a reposting from my Microsoft Technet Gallery.
 
