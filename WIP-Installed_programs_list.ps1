# This script depends on 7zip software for one option. Make sure that 7zip is installed on the computer before use the third option. #

############## Variables ##############

$7zexePath = "C:\Program Files\7-Zip\7z.exe" # Your current 7z.exe installation path
$DocumentPath = [Environment]::GetFolderPath("MyDocuments") + "\" # Find your default Documents folder location

$StarsFR = "*********************************************************************" # Stars banner for FR GUI
$StarsEN = "****************************************************" # Stars banner for EN GUI

# Variables for FR language
$FilenameFR = "Programmes Installés.txt"
$Filename7zFR = "Programmes Installés.7z"
$FilePathFR = $DocumentPath + $FilenameFR
$FilePath7zFR = $DocumentPath + $Filename7zFR
#

# Variables for EN language
$FilenameEN = "Installed Programs.txt"
$Filename7zEN = "Installed Programs.7z"
$FilePathEN = $DocumentPath + $FilenameEN
$FilePath7zEN = $DocumentPath + $Filename7zEN
#

# TEMP files
$TMPFile0 = $DocumentPath + "TMP_File0.txt"
$TMPFile1 = $DocumentPath + "TMP_File1.txt"
#

$Language = Get-Culture | Select-Object Name # Find your current locale

#$Language = "en-US"

################### Functions ###################

function Choix2et3 {
    Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object  @{Name="Nom du programme";Expression={$_.DisplayName}}, @{Name="Version du programme";Expression={$_.DisplayVersion}} | Sort-Object DisplayName | Format-Table –AutoSize | Out-File -FilePath $FilePathFR # Redirects the results in a file
    Get-ItemProperty HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object  @{Name="Nom du programme";Expression={$_.DisplayName}}, @{Name="Version du programme";Expression={$_.DisplayVersion}} | Sort-Object DisplayName | Format-Table –AutoSize | Out-File -FilePath $TMPFile0
    Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object  @{Name="Nom du programme";Expression={$_.DisplayName}}, @{Name="Version du programme";Expression={$_.DisplayVersion}} | Sort-Object DisplayName | Format-Table –AutoSize | Out-File -FilePath $TMPFile1
    Add-Content -Path $FilePathFR -Value (Get-Content -Path $TMPFile0) # Add content to the main file
    Add-Content -Path $FilePathFR -Value (Get-Content -Path $TMPFile1)
    rm $TMPFile0, $TMPFile1 # Remove temporary files
}

function Choice2and3 {
    Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object  @{Name="Program name";Expression={$_.DisplayName}}, @{Name="Program version";Expression={$_.DisplayVersion}} | Sort-Object DisplayName | Format-Table –AutoSize | Out-File -FilePath $FilePathEN
    Get-ItemProperty HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object  @{Name="Program name";Expression={$_.DisplayName}}, @{Name="Program version";Expression={$_.DisplayVersion}} | Sort-Object DisplayName | Format-Table –AutoSize | Out-File -FilePath $TMPFile0
    Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object  @{Name="Program name";Expression={$_.DisplayName}}, @{Name="Program version";Expression={$_.DisplayVersion}} | Sort-Object DisplayName | Format-Table –AutoSize | Out-File -FilePath $TMPFile1
    Add-Content -Path $FilePathEN -Value (Get-Content -Path $TMPFile0)
    Add-Content -Path $FilePathEN -Value (Get-Content -Path $TMPFile1)
    rm $TMPFile0, $TMPFile1
}

################### GUI #######################

if ($Language.Name -like "fr-*") {
    # FR Version #
    Write-Host $StarsFR
    Write-Host "| Veuillez sélectionner une option                                  |"
    Write-Host "|                                                                   |"
    Write-Host "| 1 - Afficher les résultats dans la console                        |"
    Write-Host "| 2 - Rediriger les résultats dans un fichier                       |"
    Write-Host "| 3 - Rediriger les résultats dans une archive 7z                   |"
    Write-Host $StarsFR
    Write-Host ""
} else {
    # EN Version #
    Write-Host $StarsEN
    Write-Host "| Please select an option                          |"
    Write-Host "|                                                  |"
    Write-Host "| 1 - Display the results in the console           |"
    Write-Host "| 2 - Redirect the results in a file               |"
    Write-Host "| 3 - Redirect the results in an 7z archive        |"
    Write-Host $StarsEN
    Write-Host ""
}

#################### Process ######################

# Ask for user input their choice
if ($Language.Name -like "fr-*") {
    $Choice = Read-Host -Prompt 'Entrer le numéro de votre choix' # FR Version
} else {
    $Choice = Read-Host -Prompt 'Input the number of your choice' # EN Version
}

# Process the user choice
if ($Language.Name -like "fr-*") {
    switch ($Choice) {
        # 1st option
        {$Choice -eq 1} {
            Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object  @{Name="Nom du programme";Expression={$_.DisplayName}}, @{Name="Version du programme";Expression={$_.DisplayVersion}} | Sort-Object DisplayName | Format-Table –AutoSize #Displays the results in the console
            Get-ItemProperty HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object  @{Name="Nom du programme";Expression={$_.DisplayName}}, @{Name="Version du programme";Expression={$_.DisplayVersion}} | Sort-Object DisplayName | Format-Table –AutoSize
            Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object  @{Name="Nom du programme";Expression={$_.DisplayName}}, @{Name="Version du programme";Expression={$_.DisplayVersion}} | Sort-Object DisplayName | Format-Table –AutoSize
        }
        # 2nd option
        {$Choice -eq 2} {
            Choix2et3
            Write-Host "Votre fichier est disponible à l'emplacement suivant : $FilePathFR"
        }
        # 3rd option
        {$Choice -eq 3} {
            Choix2et3
            & $7zexePath a -sdel $FilePath7zFR $FilePathFR # Compress and archive the file thanks to 7zip
            Write-Host ""
            Write-Host "Votre archive est disponible à l'emplacement suivant : $FilePath7zFR"
        }
        # If user provide an invalid option
        default {
            Write-Host Option invalide
        }
    }
} else {
    switch ($Choice) {
        # 1st option
        {$Choice -eq 1} {
            Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object  @{Name="Program name";Expression={$_.DisplayName}}, @{Name="Program version";Expression={$_.DisplayVersion}} | Sort-Object DisplayName | Format-Table –AutoSize
            Get-ItemProperty HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object  @{Name="Program name";Expression={$_.DisplayName}}, @{Name="Program version";Expression={$_.DisplayVersion}} | Sort-Object DisplayName | Format-Table –AutoSize
            Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object  @{Name="Program name";Expression={$_.DisplayName}}, @{Name="Program version";Expression={$_.DisplayVersion}} | Sort-Object DisplayName | Format-Table –AutoSize
        }
        # 2nd option
        {$Choice -eq 2} {
            Choice2and3
            Write-Host "Your file is available in the following location : $FilePathEN"
        }
        # 3rd option
        {$Choice -eq 3} {
            Choice2and3
            & $7zexePath a -sdel $FilePath7zEN $FilePathEN
            Write-Host ""
            Write-Host "Your archive is available in the following location : $FilePath7zEN"
        }
        # If user provide an invalid option
        default {
            Write-Host Invalid option
        }
    }
}