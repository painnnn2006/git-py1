Attribute VB_Name = "Module2"
Sub callApiDone()

Dim httpObject As Object
Set httpObject = CreateObject("MSXML2.XMLHTTP")

sURL = "http://localhost/api/v1/lots/done"

sRequest = sURL
httpObject.Open "GET", sRequest, False
httpObject.Send

End Sub

Sub test()
 On Error Resume Next

Application.DisplayAlerts = False
Application.ScreenUpdating = False

Application.Quit
End Sub
