# LAPS
powershell functions to work with LAPS

Invoke-RDPwithLAPS
Function to simplification of using LAPS password for RDP connection. It automatically fill hostname, login and LAPS password to RDP connection (mstsc.exe). It uses great AutoItX module for this.

How to use:
- open powershell console under account with right to read LAPS password from AD
- import modules AdmPwd.PS and AutoItX
- open Invoke-RDPwithLAPS.ps1 file
- run function Invoke-RDPwithLAPS (Invoke-RDPwithLAPS -computerName pc1)
