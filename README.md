# LAPS
powershell functions to work with LAPS



Invoke-RDPwithLAPS

Function to simplification of using LAPS password for RDP connection. It automatically fill hostname, login and LAPS password to RDP connection (mstsc.exe). It uses great AutoItX module for this.

How to use:
- open powershell console under account with right to read LAPS password from AD
- import modules AdmPwd.PS and AutoItX
- open Invoke-RDPwithLAPS.ps1 file
- run function Invoke-RDPwithLAPS (Invoke-RDPwithLAPS -computerName pc1)



Send-LAPSPassword

Function to securely send LAPS password via email (SSL) to specified email address. 
It is for example helpful in situations, when your colleague is solving unexpected problem and need local admin privileges. Solution is to send him LAPS password.
Password will also be automatically reset after specified time.

PS: another nice solution is to create HTTPS listener, that is used to send these password, securely

How to use:
- edit file Send-LAPSPassword.ps1
    - set from, to, smtpserver in sub function Send-Email to meet your environment
- open powershell console under account with right to read LAPS password from AD
- import modules AdmPwd.PS
- open Send-LAPSPassword.ps1 file
- run function Send-LAPSPassword (Send-LAPSPassword -computerName pc1 -to admin@domain.cz)
 