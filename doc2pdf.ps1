$message  = 'This script will convert any doc/docx files in this folder and subfolders to PDF and delete the existing doc/docx files.'
$question = 'Are you sure you want to proceed?'

$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))

$decision = $Host.UI.PromptForChoice($message, $question, $choices, 1)
if ($decision -eq 0) {
  Write-Host 'confirmed'
} else {
  Write-Host 'cancelled'
}

$confirmation = Read-Host "Are you REALLY Sure You Want To Proceed?"
if ($confirmation -eq 'y') {
 




$documents_path = Split-Path -parent $MyInvocation.MyCommand.Path
#If you want to enter your one path comment out the line above and use the line below.
#$documents_path = 'C:\'


$word_app = New-Object -ComObject Word.Application

# This filter will find .doc as well as .docx documents
Get-ChildItem -Path $documents_path -Filter *.doc? -Recurse | ForEach-Object {

    $document = $word_app.Documents.Open($_.FullName)

    $pdf_filename = "$($_.DirectoryName)\$($_.BaseName).pdf"

    $document.SaveAs([ref] $pdf_filename, [ref] 17)

    $document.Close()
}

get-childitem -Path $documents_path -include *.docx -recurse | foreach ($_) {remove-item $_.fullname}

$word_app.Quit()

}