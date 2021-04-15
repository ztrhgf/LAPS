# LAPS
powershell functions to work with LAPS



Invoke-MSTS

Function to simplification of using LAPS password for RDP connection. It automatically fill hostname, login and LAPS password to RDP connection (mstsc.exe). It uses great AutoItX module for this.

How to use:
- open powershell console under account with right to read LAPS password from AD
- import modules AdmPwd.PS and AutoItX
- dot source Invoke-MSTS.ps1 and Test-Connection2.ps1 files
- run function Invoke-MSTS



Send-LAPSPassword

Function to securely send LAPS password via email (SSL) to specified email address. 
It is for example helpful in situations, when your colleague is solving unexpected problem and need local admin privileges. Solution is to send him LAPS password.
Password will also be automatically reset after specified time.

PS: another nice solution is to create HTTPS listener, that is used to send these password, securely

How to use:
- open powershell console under account with right to read LAPS password from AD
- import modules AdmPwd.PS
- open Send-LAPSPassword.ps1 file
- run function Send-LAPSPassword (Send-LAPSPassword -computerName pc1 -to admin@domain.cz)
 