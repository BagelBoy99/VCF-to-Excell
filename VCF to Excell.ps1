#Uses file explorer window for picking the VCF file
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
$FileBrowser.Title = "Choose your .vcf"
$result = $FileBrowser.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))

$vcf = Get-Content $FileBrowser.FileName
#creates an array to be filled with with foreach loop
$contacts = @()
$contact = @{}


foreach ($line in $vcf) {
    if ($line -match "^BEGIN:VCARD") {
        $contact = [ordered]@{}
    } elseif ($line -match "^END:VCARD") {
        $contacts += [pscustomobject]$contact
    } elseif ($line -match "^(FN):(.*)") {
        $contact["Name"] = $matches[2]
    } elseif ($line -match "^TEL;TYPE=Work;(.*);(.*):(.*)") {
        $Contact["Work-Number"] = $matches[3]
    } elseif ($line -match "^TEL;TYPE=Cell;(.*);(.*):(.*)") {
        $Contact["Cell-Number"] = $matches[3]
    } elseif ($line -match "^TEL;TYPE=Fax;(.*);(.*):(.*)") {
        $Contact["FAX"] = $matches[3]
    } elseif ($line -match "^(EMAIL);(.*);(.*);(.*):(.*)") {
        $contact["Email"] = $matches[5]
    } elseif ($line -match "^(ADR);(.*):;;(.*);(.*);(.*);(.*)") {
        $contact["Address"] = "$($matches[3]), $($matches[4]) $($matches[5])"
    } elseif ($line -match "^(ORG):(.*);") {
        $contact["Organization"] = $matches[2]
    }
}

#Pulls up a second box for saving the file
$csv = New-Object System.Windows.Forms.SaveFileDialog
$csv.Title = "Save As"
$csv.DefaultExt = "csv"
$result = $csv.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true }))

$contacts | Select-Object -Property Name, Organization, Work-Number, Cell-Number, Email | Sort-Object -Property Name -Descending | Out-GridView -PassThru | Export-csv -Path $csv.FileName


#Sample of Raw contact info 
#VERSION:3.0
#PRODID:-//Apple Inc.//iPhone OS 18.3.1//EN
#N:Becker;Michael;;;
#FN:Michael Becker
#ORG:Royal Brothers Construction San Angelo ;
#TEL;type=CELL;type=VOICE;type=pref:(817) 690-0709
#END:VCARD