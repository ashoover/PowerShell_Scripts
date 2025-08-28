## Domain PC Renamer ##
## Requires a CSV file with the following columns ##
## CurrentPCName, NewPCName ##
## currentName.domain.com, newName.domain.com ##

$path = '.\rename_list.csv'
$csv = Import-Csv -path $path

ForEach ($line in $csv)
{
    #Write-Output $_.ColumnName
    Write-Output "Changing $($line.CurrentPCName) to $($line.NewPCName)"
    Rename-Computer ComputerName = $line.CurrentPCName -NewName $line.NewPCName -DomainCredential domain.com\DomainAdmin -Restart
    Write-Output "Now rebooting $($line.NewPCName)"
}