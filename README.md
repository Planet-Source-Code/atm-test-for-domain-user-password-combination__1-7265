<div align="center">

## Test for domain/user/password combination\.


</div>

### Description

Allow You to test Domain/User/Password combination ...
 
### More Info
 
Create form1,text1 - domain,text2 - user,text3 - password,command1.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ATM](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/atm.md)
**Level**          |Beginner
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/atm-test-for-domain-user-password-combination__1-7265/archive/master.zip)

### API Declarations

```
Private Declare Function LogonUser Lib "advapi32" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Const LOGON32_LOGON_BATCH = 4
Private Const LOGON32_PROVIDER_DEFAULT = 0
```


### Source Code

```
Private Sub Command1_Click()
Dim szDomain As String
Dim szUser As String
Dim szPassword As String
Dim lToken As Long
Dim lResult As Long
szDomain = Text1.Text & Chr(0)
szUser = Text2.Text & Chr(0)
szPassword = Text3.Text & Chr(0)
lToken = 0&
lResult = LogonUser(szUser, _
       szDomain, _
       szPassword, _
       ByVal LOGON32_LOGON_BATCH, _
       ByVal LOGON32_PROVIDER_DEFAULT, _
       lToken)
If lResult = 0 Then
 MsgBox "Error: " & Err.LastDllError
Else
 If lToken = 0 Then
 MsgBox "Not Valid user, password or domain"
 Else
 MsgBox "Valid User"
 End If
End If
End Sub
```

