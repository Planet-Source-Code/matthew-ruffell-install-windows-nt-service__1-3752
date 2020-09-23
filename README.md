<div align="center">

## Install Windows NT Service


</div>

### Description

Install a Windows NT service on a local or remote server. Configure how the service is installed. Requires Windows NT and administrator rights.
 
### More Info
 
ServiceFileName [string] = binary service file path and name, ServiceName [string] = name of service, DisplayName [string] = unofficial name of service, Interactive [boolean] = communicates with desktop, AutoStart [boolean] = run when system starts, MachineName [optional string] = target server name

Requires Windows NT and administrator rights.

Returns True if serive was successfully installed.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matthew Ruffell](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matthew-ruffell.md)
**Level**          |Unknown
**User Rating**    |4.9 (44 globes from 9 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matthew-ruffell-install-windows-nt-service__1-3752/archive/master.zip)

### API Declarations

```
Private Type SERVICE_STATUS
 dwServiceType As Long
 dwCurrentState As Long
 dwControlsAccepted As Long
 dwWin32ExitCode As Long
 dwServiceSpecificExitCode As Long
 dwCheckPoint As Long
 dwWaitHint As Long
End Type
Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function CreateService Lib "advapi32.dll" Alias "CreateServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal lpDisplayName As String, ByVal dwDesiredAccess As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, lpdwTagId As Long, ByVal lpDependencies As String, ByVal lpServiceStartName As String, ByVal lpPassword As String) As Long
Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
```


### Source Code

```
Public Function Install_SVC(strServiceFileName As String, strServiceName As String, strDisplayName As String, bolInteractive As Boolean, bolAutoStart As Boolean, Optional strMachineName As Variant, Optional strAccount As Variant, Optional strAccountPassword As Variant) As Boolean
 Dim hSCM As Long
 Dim hSVC As Long
 Dim lngInteractive As Long
 Dim lngAutoStart As Long
 Dim pSTATUS As SERVICE_STATUS
 If bolInteractive = True Then lngInteractive = (&H100 Or &H10) Else lngInteractive = &H10
 If bolAutoStart = True Then lngAutoStart = &H2 Else lngAutoStart = &H3
 If IsMissing(strMachineName) = True Then strMachineName = vbNullString Else strMachineName = CStr(strMachineName)
 If IsMissing(strAccount) = True Then strAccount = vbNullString Else strAccount = CStr(strAccount)
 If IsMissing(strAccountPassword) = True Then strAccountPassword = vbNullString Else strAccountPassword = CStr(strAccountPassword)
 '// Open the service manager
 hSCM = OpenSCManager(strMachineName, vbNullString, &H2)
 If hSCM = 0 Then Exit Function '// error opening
 '// Install the service
 hSVC = CreateService(hSCM, _
 strServiceName, _
 strDisplayName, _
 983551, _
 lngInteractive, _
 lngAutoStart, _
 0, _
 strServiceFileName, _
 vbNull, _
 vbNull, _
 vbNullString, _
 strAccount, _
 strAccountPassword)
 If hSVC <> 0 Then Install_SVC = True
 Call CloseServiceHandle(hSVC)
 Call CloseServiceHandle(hSCM)
End Function
```

