Write-Host "------------------------------------"
Write-Host "|   *******'s WONDERWARE SCRIPT    |"
Write-Host "------------------------------------"

#Install Windows Update module
Install-Module PSWindowsUpdate

#Allows PS to update non-Windows MS products (like Office)
Add-WUServiceManager -MicrosoftUpdate 

#function to run Selenium/Python to pull the list of updates.I should bring in pyAutoGui and create a popup for creds.
python C:\users\admin2\desktop\Wonderware_Selenium.py

#We now have Aveva data, which needs to be formatted.
$inbound = "c:\users\admin\desktop\SecurityCentralSupportedProducts.xlsx"
$fileTocsv = "c:\users\admin\desktop\wonderware.csv"
$avevaList = "c:\users\admin\desktop\lifeboat\avevalist.txt"

#convert Aveva inbound xlsx to a csv
$excel = New-Object -ComObject Excel.Application
$reportOut = $excel.Workbooks.Open($inbound)
$excel.DisplayAlerts = $false;
$reportOut.SaveAs($fileTocsv,[Microsoft.Office.Interop.Excel.XlFileFormat]::xlCSV)
$reportOut.Close
$excel.Quit()

#begin data manipulation on Aveva csv
$csv = Import-CSV -Path $fileTocsv

forEach ($row in $csv){
    $postedDate = Get-Date ($row.Posted)
    $today = get-Date -format "MM/dd/yyyy HH:mm:ss tt"
    $days = New-TimeSpan -Start $today -End $postedDate
    $kb = $row.'MS KB Number'
    if(($days.Days -ge '-60') -and ($row.Status -eq 'Supported')){
        Add-Content $avevaList "$kb"
        }
    else{
    }
    }

#Now Aveva updates are in a txt list

#Cleans up Aveva list
$file = get-content -raw $avevaList
$newlines = foreach ($line in $file){
    $line -replace ', ',"`r"
    }
Set-Content $avevaList $newlines
#Aveva list is now KB only and ready for use

#We need a list of available updates via the os. FORMAT UNVERIFIED
get-wulist | Format-List -property KB | out-file c:\users\admin2\desktop\availableUpdates.txt

#cleans available update output
$file2 = "C:\users\admin\desktop\availableupdates.txt"
$newlines = foreach ($line in $file2){
    $line -replace 'KB : ',""
    }
Set-Content $file2 $newlines

#Compare lists. 2/22/21 this works, leave alone.

$list2 = "C:\users\admin2\desktop\availableUpdates.txt" #diff obj, Get-WUList
$list3 = "C:\users\admin2\desktop\Final_List.txt" #output list

#Populate final list with matches from
Compare-Object -IncludeEqual -ExcludeDifferent -passthru -ReferenceObject (Get-Content -Path $avevaList) -DifferenceObject (Get-Content -Path $list2) | Out-File -FilePath $list3

#Install updates from final list
forEach ($i in $list3){
    Get-WindowsUpdate -KBArticleID $i -Install 
    Write-host $i
}

#Install specific updates
#Get-WindowsUpdate -KBArticleID "KB1111111","KB2222222","etc" -Install 