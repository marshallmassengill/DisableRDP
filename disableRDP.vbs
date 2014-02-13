const HKEY_LOCAL_MACHINE = &H80000002
strComputer = "."
Set oReg=GetObject( _
    "winmgmts:{impersonationLevel=impersonate}!\\" &_ 
    strComputer & "\root\default:StdRegProv")
strKeyPath = "SYSTEM\CurrentControlSet\Control\Terminal Server"
strValueName = "fDenyTSConnections"
oReg.SetDWORDValue _ 
    HKEY_LOCAL_MACHINE,strKeyPath,strValueName,1
If Err = 1 Then
   oReg.GetDWORDValue _
       HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
   WScript.Echo _
       "Problem with setting this registry value"
End If