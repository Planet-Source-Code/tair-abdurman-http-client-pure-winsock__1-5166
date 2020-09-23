<div align="center">

## HTTP Client \- pure WinSock


</div>

### Description

Allow retrieve HTML page sources anywhere from web, directly or via proxy server, can access virtual domains. Pure winsock, no any other components used! Wanna know web transactions in deep?
 
### More Info
 
'based on HTTP 1.0 - RFC 1945

'see http://www.tair.freeservers.com for more info, details and downloads!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tair Abdurman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tair-abdurman.md)
**Level**          |Advanced
**User Rating**    |4.2 (50 globes from 12 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tair-abdurman-http-client-pure-winsock__1-5166/archive/master.zip)





### Source Code

```
'based on HTTP 1.0 - RFC 1945
'see http://www.tair.freeservers.com for more info, details and downloads!
Public JobURL As String
Public ResponseDocument As String
Public StepCount As Long
Public IsProxyUsed As Boolean
Public ServerHostIP As String
Public ServerPort As Long
'------------------------------------------------------------
Dim LocalStepCounter As Long
Dim RequestHeader As String
Dim RequestTemplate As String
'------------------------------------------------------------
Public Sub ActionStartup()
 If UCase(Left(JobURL, 7)) <> "HTTP://" Then
 MsgBox "Please enter url with http://", vbCritical + vbOK
 FrmActionWait.Hide
 Unload FrmActionWait
 Exit Sub
 End If
 LocalStepCounter = 0
 RequestHeader = ""
 RequestTemplate = "GET _$-$_$- HTTP/1.0" & Chr(13) & Chr(10) & _
  "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, application/x-comet, */*" & Chr(13) & Chr(10) & _
  "Accept-Language: en" & Chr(13) & Chr(10) & _
  "Accept-Encoding: gzip , deflate" & Chr(13) & Chr(10) & _
  "Cache-Control: no-cache" & Chr(13) & Chr(10) & _
  "Proxy-Connection: Keep-Alive" & Chr(13) & Chr(10) & _
  "User-Agent: SSM Agent 1.0" & Chr(13) & Chr(10) & _
  "Host: @$@@$@" & Chr(13) & Chr(10)
 pureURL = Right(JobURL, Len(JobURL) - 7)
 startPos = InStr(1, pureURL, "/")
 If startPos < 1 Then
 ServerAddress = pureURL
 documentURI = "/"
 Else
 ServerAddress = Left(pureURL, startPos - 1)
 documentURI = Right(pureURL, Len(pureURL) - startPos + 1)
 End If
 If ServerAddress = "" Or documentURI = "" Then
 MsgBox "Unable to detect target page!", vbCritical + vbOK
 FrmActionWait.Hide
 Unload FrmActionWait
 Exit Sub
 End If
 If IsProxyUsed Then
 If ServerHostIP = "" Then
  MsgBox "Unable to detect proxy address!", vbCritical + vbOK
  FrmActionWait.Hide
  Unload FrmActionWait
  Exit Sub
 End If
 RequestHeader = RequestTemplate
 RequestHeader = Replace(RequestHeader, "_$-$_$-", JobURL)
 Else
 ServerHostIP = ServerAddress
 ServerPort = 80
 RequestHeader = RequestTemplate
 RequestHeader = Replace(RequestHeader, "_$-$_$-", documentURI)
 End If
 Me.Show
 RequestHeader = Replace(RequestHeader, "@$@@$@", ServerAddress)
 RequestHeader = RequestHeader & Chr(13) & Chr(10)
 TxtStatus.Text = "Connecting to server ..."
 TxtStatus.Refresh
 WS_HTTP.Connect ServerHostIP, ServerPort
End Sub
Private Sub WS_HTTP_Close()
 WS_HTTP.Close
 TxtStatus.Text = "Transaction completed ..."
 TxtStatus.Refresh
 Me.Hide
 Unload Me
End Sub
Private Sub WS_HTTP_Connect()
 WS_HTTP.SendData RequestHeader
 TxtStatus.Text = "Connected, try to obtain page ..."
 TxtStatus.Refresh
 FrmMainWin.TxtResponse.Text = ""
 FrmMainWin.TxtResponse.Refresh
End Sub
Private Sub WS_HTTP_DataArrival(ByVal bytesTotal As Long)
 Dim tmpString As String
 WS_HTTP.GetData tmpString, vbString
 FrmMainWin.TxtResponse.Text = FrmMainWin.TxtResponse.Text & tmpString
 FrmMainWin.TxtResponse.Refresh
 TxtStatus.Text = "Data from server, continue ..."
 TxtStatus.Refresh
End Sub
Private Sub WS_HTTP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
 WS_HTTP.Close
 TxtStatus.Text = "Errors occured ..."
 TxtStatus.Refresh
 Me.Hide
 Unload Me
End Sub
```

