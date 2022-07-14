VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Userform_trace 
   ClientHeight    =   3996
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   8832.001
   OleObjectBlob   =   "Userform_trace.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Userform_trace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    'Get list ds
   
            Dim oFSO As Object
            Dim oFolder As Object
            Dim oFile As Object
            Dim i As Integer
            
            Set oFSO = CreateObject("Scripting.FileSystemObject")
            
            Set oFolder = oFSO.GetFolder(Me.txt_path)
            
            For Each oFile In oFolder.Files
            
                  'Make new file
                    Dim createopt As StructCreateOptions
                        Set createopt = CreateStructCreateOptions
                        With createopt
                            .Name = "Untitled-1"
                            .Units = cdrPixel
                            .PageWidth = GetImageSize(oFile.Path)(0) * 0.0000254
                            .PageHeight = GetImageSize(oFile.Path)(1) * 0.0000254
                            .Resolution = 300#
                            .ColorContext = CreateColorContext2("sRGB IEC61966-2.1,U.S. Web Coated (SWOP) v2,Dot Gain 20%", BlendingColorModel:=clrColorModelCMYK)
                        End With
                        Dim doc1 As Document
                        Set doc1 = CreateDocumentEx(createopt)
     
               Debug.Print GetImageSize(oFile.Path)(0)
               Debug.Print GetImageSize(oFile.Path)(1)
                   
    'Import ds
                    ActiveLayer.Import (oFile.Path)
            
                   
    'Trace ds


                        Dim s As Shape, sr As ShapeRange
                        Dim t As TraceSettings
                        For Each s In ActivePage.FindShapes(, cdrBitmapShape)
                          Set t = s.Bitmap.trace(cdrTraceLineArt, 25, 100, cdrColorBlackAndWhite, cdrUniform, , True, True, True)
                          t.Finish
                        Next s
                        Set t = Nothing
                        Set s = Nothing
                   
                   
                   
    ' save file
                        Dim SaveOptions As StructSaveAsOptions
                        Set SaveOptions = CreateStructSaveAsOptions
                        With SaveOptions
                            .EmbedVBAProject = False
                            .Filter = cdrCDR
                            .IncludeCMXData = False
                            .Range = cdrAllPages
                            .EmbedICCProfile = False
                            .Version = cdrVersion14
                            .KeepAppearance = True
                        End With
                        doc1.SaveAs Me.txt_path & "\" & Split(oFile.Name, ".")(0) & ".cdr", SaveOptions
                        doc1.Close
                   
            Next oFile
            
            
 Me.txt_path = ""
'MsgBox "Done"

End Sub



Function GetImageSize(ImagePath As String) As Variant
    
   
    Dim imgSize(1)  As Integer
    Dim wia         As Object
    
    'Check  file exists
    If FileExists(ImagePath) = False Then Exit Function
    
    'Check image format
    If IsValidImageFormat(ImagePath) = False Then Exit Function
    
    On Error Resume Next
    Set wia = CreateObject("WIA.ImageFile")
    If wia Is Nothing Then Exit Function
    On Error GoTo 0
    
    
    wia.LoadFile ImagePath
    
    
    imgSize(0) = wia.Width
    imgSize(1) = wia.Height
    
    
    Set wia = Nothing
    
    
    GetImageSize = imgSize

End Function

Function FileExists(FilePath As String) As Boolean
   
   

    On Error Resume Next
    If Len(FilePath) > 0 Then
        If Not Dir(FilePath, vbDirectory) = vbNullString Then FileExists = True
    End If
    On Error GoTo 0
   
End Function

Function IsValidImageFormat(FilePath As String) As Boolean
   
    Dim imageFormats    As Variant
    Dim i               As Integer
                
   
    imageFormats = Array(".png")   '".bmp", ".jpg", ".gif", ".tif",
                    
  
    For i = LBound(imageFormats) To UBound(imageFormats)
        
        If InStr(1, UCase(FilePath), UCase(imageFormats(i)), vbTextCompare) > 0 Then
            IsValidImageFormat = True
            Exit Function
        End If
    Next i
   
End Function



Private Sub CommandButton2_Click()
Userform_trace.Hide
End Sub

Private Sub UserForm_Click()

End Sub
