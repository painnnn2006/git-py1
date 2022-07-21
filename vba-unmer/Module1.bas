Attribute VB_Name = "Module1"




Sub fix_lot()

Dim selectFile As Variant
Dim wbinput, rdb As Workbook
Dim i, ifile, isheetnum As Integer
Dim wb_name As String
Dim ws_cl  As Worksheet
                 
  On Error Resume Next


Set rdb = ThisWorkbook

Dim ws_data, ws_layout, ws_var, ws_m7, ws_pro_cn As Worksheet
Set ws_var = rdb.Sheets("name_varcode")
Set ws_layout = rdb.Sheets("Layout")
Set ws_m7 = rdb.Sheets("m7 product")
Set ws_pro_cn = rdb.Sheets("pro_cn")
Set ws_data = rdb.Sheets("data")


 If ws_data.Range("A4").Value = 0 Then
        MsgBox ("méo có file")
        Application.Quit
        

       End If


Application.DisplayAlerts = False
Application.ScreenUpdating = False


    'get file xlsx

    Dim path_folder_lot, path_lot_raw, path_lot, final_path, file_name_raw, lot_id, lot_type As String

    path_folder_lot = ws_data.Range("A2").Value
        path_lot_raw = path_folder_lot & "\" & Dir(path_folder_lot & "\*.xlsx")
        file_name_raw = Dir(path_folder_lot & "\*.xlsx")
    ' MissingOrder-URGENT-42084-ST-20220929-10_HAIAN_123
    'URGENT-42084-ST-20220929-10_HAIAN_123

    lot_type = Split(file_name_raw, "-")(0)
    If lot_type = "URGENT" Or "MissingOrder" Then
        If lot_type = "URGENT" Then
            file_name = replace(file_name_raw, "URGENT-", "")
        ElseIf lot_type = "MissingOrder" Then
            file_name = replace(file_name_raw, "MissingOrder-URGENT-", "")
        End If

        path_lot = ""
            'rename
        Else
            lot_type = ""
    End If

    Debug.Print path_lot
    'open file
        
       Set wbinput = Workbooks.Open(path_lot)
       
       
       Dim fac, pro_code As String

    fac = Split(wbinput.Name, "_")(1)
    pro_code = Split(wbinput.Name, "-")(1)
    lot_id = Split(Split(wbinput.Name, "_")(2), ".")(0)
    Debug.Print(fac & "___" & pro_code & "___" & lot_id)


    'check product


    isheetnum = wbinput.Worksheets.Count
      
      If isheetnum = 1 Then
      
      MsgBox ("thieu sheets bn ui!!")
      
      Dim del_path As String
      del_path = wbinput.Path & "\" & wbinput.Name
      Debug.Print ("del path   " & del_path)
      wbinput.Close
      
    Kill (del_path)
      Application.Quit
      End If
        
        
        
'make folder final
final_path = ws_data.Range("A3").Value & "hiweb-design-tool\storage\app\public\excel\processed\" & Split(Dir(path_folder_lot & "\*.xlsx"), ".")(0)

    If Dir(ws_data.Range("A3").Value & "hiweb-design-tool\storage\app\public\excel\processed", vbDirectory) = "" Then
    
        MkDir (ws_data.Range("A3").Value & "hiweb-design-tool\storage\app\public\excel\processed")
    End If
    
        
    If Dir(final_path, vbDirectory) = "" Then
    
    
        MkDir final_path
        
        
    End If
    
        
      

        
        


      
      
      'fix name file
      
      
   
    'dected product
            For i = 1 To isheetnum
                
                If wbinput.Sheets(i).Name = "SINGLE" Then
                        wbinput.Sheets("TOTAL").Range("B1").Value = wbinput.Sheets(i).Range("H2").Value
                                    wbinput.Sheets("TOTAL").Range("A1").Value = Excel.WorksheetFunction.Sum(wbinput.Sheets(i).Range("E:E"))
                            
                                        wbinput.Sheets(i).Range("C:C").HorizontalAlignment = xlLeft
                                        wbinput.Sheets(i).Range("C:C").VerticalAlignment = xlBottom
                                  If pro_code = "CR" Then
                                             For b = 2 To Excel.WorksheetFunction.CountA(wbinput.Sheets(i).Range("G:G"))
                                                       If wbinput.Sheets(i).Range("H" & b).Value <> "A01" Then
                                                           wbinput.Sheets(i).Range("G" & b).Value = wbinput.Sheets(i).Range("G" & b).Value & "-N"
                                                           
                                                       
                                                       End If
                                                       
                                                Next
                                  
                                        wbinput.Sheets(i).Columns("E:I").EntireColumn.AutoFit
                                        
                                        Else
                                        
                                        wbinput.Sheets(i).Columns("D:I").EntireColumn.AutoFit
                                    
                                    End If
                                               'le
                                               If pro_code = "LE" Then
                                                    
                                                    Sheets.Add after:=wbinput.Sheets(i)
                                                    
                                                    Set ws_cl = ActiveSheet
                                                    wbinput.Sheets(i).Range("F:G").Copy
                                                    ws_cl.Range("A1").PasteSpecial xlPasteValues
                                                    
                                                    ws_cl.Columns("A:A").TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
                                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                                Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
                                :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
                                'add N nonbrand- non A01
                                        For a = 2 To Excel.WorksheetFunction.CountA(ws_cl.Range("A:A"))
                                            If ws_cl.Range("D" & a).Value = "" And ws_cl.Range("B" & a).Value <> "A01" Then
                                                ws_cl.Range("A" & a).Value = ws_cl.Range("A" & a).Value & "-N"
                                                
                                            
                                            End If
                                            
                                        Next
                                        ws_cl.Range("A:A").Copy wbinput.Sheets(i).Range("F:F")
                                        ws_cl.Delete
                                        
                                               End If
                       
                        
                        
                        
                        
                        'remove cd js
                        
                         If fac = "LE2" Or fac = "LE1" Then
                                     wbinput.Sheets(i).Range("F:F").Replace What:="-CD", Replacement:="", LookAt:=xlPart, _
                                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                                    ReplaceFormat:=False
        
                                     wbinput.Sheets(i).Range("F:F").Replace What:="-JS", Replacement:="", LookAt:=xlPart, _
                                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                                    ReplaceFormat:=False
                                    
                                     wbinput.Sheets(i).Range("F:F").Replace What:="-CE", Replacement:="", LookAt:=xlPart, _
                                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                                    ReplaceFormat:=False
                         End If
                         
                       
                       
                                        wbinput.Sheets(i).Columns("E:I").EntireColumn.AutoFit
                                        
                         
                    
                End If
                
                
                
                
                
                
            
                If wbinput.Sheets(i).Name = "MULTI" Then
                
                wbinput.Sheets("TOTAL").Range("B1").Value = wbinput.Sheets(i).Range("J2").Value
                                wbinput.Sheets(i).Range("C:C").HorizontalAlignment = xlLeft
                                        wbinput.Sheets(i).Range("C:C").VerticalAlignment = xlBottom
                                        
                                          If pro_code = "LE" Then
                                                    
                                                    Sheets.Add after:=wbinput.Sheets(i)
                                                    
                                                    Set ws_cl = ActiveSheet
                                                    wbinput.Sheets(i).Range("H:I").Copy
                                                    ws_cl.Range("A1").PasteSpecial xlPasteValues
                                                    
                                                    ws_cl.Columns("A:A").TextToColumns Destination:=Range("C1"), DataType:=xlDelimited, _
                                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                                Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
                                :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
                                'add N nonbrand- non A01
                                        For a = 2 To Excel.WorksheetFunction.CountA(ws_cl.Range("A:A"))
                                            If ws_cl.Range("D" & a).Value = "" And ws_cl.Range("B" & a).Value <> "A01" Then
                                                ws_cl.Range("A" & a).Value = ws_cl.Range("A" & a).Value & "-N"
                                                
                                            
                                            End If
                                            
                                        Next
                                        ws_cl.Range("A:A").Copy wbinput.Sheets(i).Range("H:H")
                                        ws_cl.Delete
                                        
                                               End If
                                               
                                               
                                               
                                               
                                        
                                        
                                        If fac = "LE1" Or fac = "LE2" Then
                                     wbinput.Sheets(i).Range("H:H").Replace What:="-CD", Replacement:="", LookAt:=xlPart, _
                                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                                    ReplaceFormat:=False
        
                                     wbinput.Sheets(i).Range("H:H").Replace What:="-JS", Replacement:="", LookAt:=xlPart, _
                                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                                    ReplaceFormat:=False
                                         End If
                         
                         
                                        
                                
                                  
                                        wbinput.Sheets(i).Columns("E:J").EntireColumn.AutoFit
                                
                                wbinput.Sheets("TOTAL").Range("A2").Value = Excel.WorksheetFunction.CountA(wbinput.Sheets(i).Range("H:H")) - 1
                                
                                  
                    
                End If
                
                
                wbinput.Sheets("TOTAL").Range("A3").Value = wbinput.Sheets("TOTAL").Range("A2").Value + wbinput.Sheets("TOTAL").Range("A1").Value
                
                 
                        
                
                
            
            Next
            
                    If wbinput.Sheets("TOTAL").Range("A2") = "" Then
                    
                           
                           If Excel.WorksheetFunction.CountIf(ws_var.Range("A:A"), wbinput.Sheets("SINGLE").Range("H2").Value) > 0 Then
                                 wb_name = Split(wbinput.Name, "-")(0) & "-" & wbinput.Sheets("SINGLE").Range("H2").Value & "-" & Split(wbinput.Name, "-")(2) & "-" & Split(wbinput.Name, "-")(3)
                                     Else
                                     
                           wb_name = Split(wbinput.Name, "-")(0) & "-" & Split(wbinput.Name, "-")(1) & "-" & Split(wbinput.Name, "-")(2) & "-" & Split(wbinput.Name, "-")(3)
                             
                             End If
                        Else
                            If Excel.WorksheetFunction.CountIf(ws_var.Range("A:A"), wbinput.Sheets("MULTI").Range("J2").Value) > 0 Then
                                 wb_name = Split(wbinput.Name, "-")(0) & "-" & wbinput.Sheets("MULTI").Range("J2").Value & "-" & Split(wbinput.Name, "-")(2) & "-" & wbinput.Sheets("TOTAL").Range("A3").Value & "packs-" & Split(wbinput.Name, "-")(3) & "pcs"
                                     Else
                                     
                           wb_name = Split(wbinput.Name, "-")(0) & "-" & Split(wbinput.Name, "-")(1) & "-" & Split(wbinput.Name, "-")(2) & "-" & wbinput.Sheets("TOTAL").Range("A3").Value & "packs-" & Split(wbinput.Name, "-")(3) & "pcs"
                            End If
                            
                            
                        End If
                        
                          wbinput.Sheets("TOTAL").Range("A1:A3").Clear
            
                        wbinput.Sheets("TOTAL").Range("D1").Value = wb_name
                            
                         wbinput.Sheets("TOTAL").Columns("D:AA").EntireColumn.AutoFit
                         
                         
                
                
                       old_path = wbinput.Path & "/" & wbinput.Name
                
                                   
                
                
'detect xlds
Debug.Print (Excel.WorksheetFunction.CountIf(ws_pro_cn.Range("A:A"), pro_code))



If fac = "SG" Or (fac <> "SG" And Excel.WorksheetFunction.CountIf(ws_pro_cn.Range("A:A"), pro_code) > 0) Or wbinput.Sheets("TOTAL").Range("B1").Value = "ST-R8" Then      'xlds

 wbinput.SaveAs Filename:= _
                                                            final_path & "/" & wb_name, FileFormat:= _
                                                            xlOpenXMLWorkbook, CreateBackup:=False
                                                            wbinput.Close
                                                            
ws_data.Range("AA1").Value = final_path
UserForm1.Show

ws_data.Range("AA1").Value = ""

Debug.Print ("ds tool")


Else
'nomal
'Debug.Print final_path

 wbinput.SaveAs Filename:= _
                                                            final_path & "/" & wb_name & "__0__0", FileFormat:= _
                                                            xlOpenXMLWorkbook, CreateBackup:=False
                                                            wbinput.Close
                                                            
                                                            Debug.Print ("nomal")

                                                            

End If


            
            
    
    Kill (old_path)
    Call Module2.callApiDone



            
            
Application.ScreenUpdating = True
            

'detect fix
If Split(Split(ThisWorkbook.Name, "_")(1), ".")(0) = "1" Then
ThisWorkbook.Save
Application.Quit

End If
'Application.DisplayAlerts = True


End Sub




Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function



Function isSheetExist(shName As String) As Boolean
  isSheetExist = Not Evaluate("IsError(" & shName & "!1:1)")
End Function

Function CountFiles(strDir As String, Optional strType As String)
    Dim file As Variant, i As Integer
    If Left(strDir, 1) <> "\" Then strDir = strDir & "\"
    file = Dir(strDir & strType)
    While (file <> "")
        i = i + 1
        file = Dir
    Wend
    CountFiles = i
End Function





