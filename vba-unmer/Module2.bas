Attribute VB_Name = "Module2"
Sub callApiDone()


Dim httpObject As Object
Set httpObject = CreateObject("MSXML2.XMLHTTP")
Dim lot_id, path, type As String
    lot_id = ""
    path = ""
    type = ""

    sURL = "http://localhost/api/v1/lots/" & lot_id & "/done?path=" & path & "&type=" & type


    sRequest = sURL
httpObject.Open "GET", sRequest, False
httpObject.Send

End Sub


