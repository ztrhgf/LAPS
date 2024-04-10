# LAPS
powershell functions to make your life with LAPS easier



## Invoke-MSTS

Function to simplification of using LAPS password for RDP connection. 
It automatically fills the hostname, login, and LAPS password to the RDP connection (mstsc.exe). 
It uses cmdkey.exe to import credentials to a credential manager or if no credentials are available autofill connection details with the help of the great AutoItX module.

How to use:
- download the whole repository as a ZIP file and extract it to your hard drive
- open the PowerShell console under an account with the right to read the LAPS password from AD and navigate to the extracted folder
  - `Set-Location '<extractedFolderRoot>'`
- import modules AdmPwd.PS and AutoItX
  - `Import-Module AdmPwd.PS`
  - `Import-Module AutoItX` 
- dot source Invoke-MSTS.ps1 and Test-Connection2.ps1 files
  - `. .\Invoke-MSTS.ps1`
  - `. .\Test-Connection2.ps1` 
- run function Invoke-MSTS or its alias RDP
  - `Invoke-MSTS PC-01`


![rdp](https://user-images.githubusercontent.com/2930419/119770427-7c17dd80-bebc-11eb-86b8-f82e5d6c781f.gif)



## Send-LAPSPassword

Function to securely send LAPS password via email (SSL) to specified email address. 
It is for example helpful in situations, when your colleague is solving unexpected problem and need local admin privileges. Solution is to send him LAPS password.
Password will also be automatically reset after specified time.

PS: another nice solution is to create HTTPS listener, that is used to send these password, securely

How to use:
- download the whole repository as a ZIP file and extract it to your hard drive
- open the PowerShell console under an account with the right to read the LAPS password from AD and navigate to the extracted folder
  - `Set-Location '<extractedFolderRoot>'`
- import modules AdmPwd.PS
  - `Import-Module AdmPwd.PS` 
- dot source Send-LAPSPassword.ps1 file
  - `. .\Send-LAPSPassword.ps1` 
- run function Send-LAPSPassword
  - `Send-LAPSPassword -computerName pc1 -to admin@domain.com`
 
