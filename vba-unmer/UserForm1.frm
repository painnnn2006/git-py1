VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "                    *****UNMERGE FILE*****"
   ClientHeight    =   5190
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6360
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub UserForm_Activate()

Dim selectFile As Variant
Dim wbinput As Workbook
Dim i, ifile, isheetnum As Integer
Dim wb_name, wb_code, wb_old_path As String
Dim ws_db, ws_m7, ws_br As Worksheet
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.ErrorCheckingOptions.BackgroundChecking = False


  On Error Resume Next


 Dim Pshp_non, Pshp_non_mul, Pshp_br, Pshp_br_mul   As Shape
            Dim xRg_non, xRg_non_mul, xRg_br, xRg_br_mul As Range
            Dim xCol_non, xCol_non_mul, xCol_br, xCol_br_mul  As Long
            Dim lr_mk_br, lr_mk_br_mul, lr_mk_non_mul, lr_mk_non As Integer

Set ws_m7 = ThisWorkbook.Sheets("m7 product")
Set ws_db = ThisWorkbook.Sheets("data")
Dim ws_layout As Worksheet
Set ws_layout = ThisWorkbook.Sheets("Layout")


Dim ur_code, ur_code_br As String

    '
    '

    'get file xlsx

    Dim path_folder_lot, path_lot, final_path, fac As String

    path_folder_lot = ws_db.Range("AA1").Value
        'get folder fix path
                
        path_lot = path_folder_lot & "\" & Dir(path_folder_lot & "\*.xlsx")
        
        fac = Split(path_folder_lot, "_")(1)
    final_path = ws_db.Range("AA1").Value

    Debug.Print (final_path)
        Debug.Print (fac)
        
        
       Set wbinput = Workbooks.Open(path_lot)


   
        
      isheetnum = wbinput.Worksheets.Count
      
                 For i = 1 To isheetnum
                 
                 'MsgBox (wbinput.Sheets(i).Name)
                 
                                If i = 1 Then
                 
                  
                    
                        
                                     Dim wsc, arts As Worksheet
                                     wbinput.Sheets(isheetnum).Select
                                     
                                     
                                    Sheets.Add after:=ActiveSheet
                                    
                                    
                                    ActiveSheet.Name = "pro_code"
                                    Set wsc = wbinput.Sheets("pro_code")
                                    
                                    
                                    wsc.Range("A1").Value = Split(wbinput.Name, "-")(1)
                                    
'
                                        wb_code = wsc.Range("A1").Value
                                        
                                        wsc.Range("B1").Value = Split(wbinput.Name, ".")(0)
                                        wsc.Columns("B:B").TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
                                        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                                        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
                                        :="_", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
                                        
                                    
                                    
                                        Range("B:B").Delete shift:=xlToLeft
                                        wsc.Range("BC1").Value = "brand code"
                                        
                                        
                                
                                
                                'add layout, var code
                                Dim var_code As String
                                
                                
                                            If isSheetExist("SINGLE") = True Then
                                                    var_code = wbinput.Sheets("SINGLE").Range("H2").Value
                                                If Excel.WorksheetFunction.CountIf(ws_layout.Range("A:A"), wbinput.Sheets("SINGLE").Range("H2").Value) > 0 Then


                                                    
                                                    Sheets.Add after:=wsc
                                                    Set cws = ActiveSheet
                                                    cws.Name = "LAYOUT"
                                                     ws_db.Range("b1").Value = wbinput.Sheets("SINGLE").Range("H2").Value
                                                     
                                                    ws_db.Range("D1:H2").Copy
                                                    cws.Range("A1:E2").PasteSpecial xlPasteValues
                                             
                                                End If
                                            Else
                                                var_code = wbinput.Sheets("MULTI").Range("J2").Value
                                                If Excel.WorksheetFunction.CountIf(ws_layout.Range("A:A"), wbinput.Sheets("MULTI").Range("J2").Value) > 0 Then
                                                        
                                                    Sheets.Add after:=wsc
                                                    Set cws = ActiveSheet
                                                    cws.Name = "LAYOUT"
                                                        ws_db.Range("b1").Value = wbinput.Sheets("MULTI").Range("J2").Value
                                                        ws_db.Range("D1:H2").Copy
                                                        cws.Range("A1:E2").PasteSpecial xlPasteValues
                                                End If
                                            
                                            End If
                                            
                                                 


                                        
                                    
                                End If
                                        
                                        
                               If wbinput.Sheets(i).Name = "MULTI" Then
                               
                                 wsc.Range("BE3").Value = 1
                            'add sheet xly
                                        
                                        'clear count
                                        
                                        wbinput.Sheets(i).Range("A" & Excel.WorksheetFunction.CountA(wbinput.Sheets(i).Range("A:A")) + 1).Clear
                                        
                                          
                                    Sheets.Add after:=wbinput.Sheets(i)
                                    'Dim nws As Worksheet
                                    Set nws = ActiveSheet
                                    
                                    
                                    wbinput.Sheets(i).Range("F:F").Copy nws.Range("A:A")
                                    nws.Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
                                        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                                        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
                                        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
                                    nws.Columns("A:A").Delete
                                    
                                     nws.Range("A:A").Replace What:=" ", Replacement:="", LookAt:=xlPart, _
                                        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                                        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
                                        
                                        
                                    
                                    
                                    nws.Range("AA1").Value = wsc.Range("A1").Value
                                    
                                    
                                    
                                    'detect color - 2d product
                                    
                                    If wsc.Range("A1").Value = "TT1" Or wsc.Range("A1").Value = "TT2" Or wsc.Range("A1").Value = "HDT" Or wsc.Range("A1").Value = "HDS" Or wsc.Range("A1").Value = "SST" Or wsc.Range("A1").Value = "SSS" Or wsc.Range("A1").Value = "TAT" Or wsc.Range("A1").Value = "TAS" Or wsc.Range("A1").Value = "WTT" Or wsc.Range("A1").Value = "WTS" Or wsc.Range("A1").Value = "BG" Then
                                            MsgBox "2d product"
                                            
                                            
                                         wbinput.Sheets(i).Range("F:F").Copy nws.Range("C:C")
                                            
                                             If wsc.Range("A1").Value = "TT1" Or wsc.Range("A1").Value = "TT2" Then

                                                    nws.Range("D2").Value = "=MID(C2,LEN(""Tshirt 2D "")+1,1)"
                                                ElseIf wsc.Range("A1").Value = "HDT" Or wsc.Range("A1").Value = "HDS" Then
                                        
                                                    nws.Range("D2").Value = "=MID(C2,LEN(""Hooded Sweatshirt 2D T "")+1,1)"
                                                    ElseIf wsc.Range("A1").Value = "TAT" Or wsc.Range("A1").Value = "TAS" Then
                                        
                                                             nws.Range("D2").Value = "=MID(C2,LEN(""Tanktop 2D T "")+1,1)"
                                                        ElseIf wsc.Range("A1").Value = "SST" Or wsc.Range("A1").Value = "SSS" Then
                                        
                                                            nws.Range("D2").Value = "=MID(C2,LEN(""Sweatshirt 2D T "")+1,1)"
                                                             ElseIf wsc.Range("A1").Value = "WTT" Or wsc.Range("A1").Value = "WTS" Then
                                                            
                                                            nws.Range("D2").Value = "=MID(C2,LEN(""Women's tank top 2D "")+1,1)"
                                                            
                                                             
                                                End If
                                        
                                                
                                                
                                    
                                        nws.Range("B2").Value = "=$AA$1 & ""-"" & D2&A2"
                                   Else
                                    nws.Range("B2").Value = "=$AA$1 & ""-"" & A2"
                                   End If
                                   
                                   
                            
                 
                                    
                                    
                                    
                                    Dim lr_mul As Double
                                    lr_mul = Excel.WorksheetFunction.CountA(nws.Range("A:A"))
                                    
                                    If lr_mul > 1 Then
                                    
                                            nws.Range("D2:D" & lr_mul + 1).FillDown
                                            nws.Range("B2 : B" & lr_mul + 1).FillDown
                                    End If
                                                        
                                            wbinput.Sheets(i).Range("J:J").unmerge
                                            
                                            nws.Range("B2:B" & lr_mul + 1).Copy
                                            wbinput.Sheets(i).Range("J2").PasteSpecial Paste:=xlPasteValues
                                            'remove space
                                            wbinput.Sheets(i).Range("J2:J1000").Replace What:=" ", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                                       ReplaceFormat:=False
                                             
                                'count pcs mul
                                    If wbinput.Sheets(i).Range("E1").Value = "Quantity" Then
                                        wsc.Range("BB1").Value = Excel.WorksheetFunction.Sum(wbinput.Sheets(i).Range("E2:E" & lr_mul + 1))
                                    End If
                                
                                    
                                        
                                            
                                nws.Delete
                                            
                                'split brand
                                 If fac = "SG" And wb_code <> "ST" And wb_code <> "BO" Then
                                 wbinput.Sheets(i).Range("A:B", "G:I").unmerge
                                
                                For C = 3 To lr_mul + 1
                                    If wbinput.Sheets(i).Range("H" & C).Value = "" Then
                                        wbinput.Sheets(i).Range("H" & C).Value = wbinput.Sheets(i).Range("H" & C - 1).Value
                                        
                                    End If
                                    
                                Next
                                
                                wbinput.Sheets(i).Columns("H:H").TextToColumns Destination:=Range("L1"), DataType:=xlDelimited, _
                                        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                                        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
                                        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
                                wbinput.Sheets(i).Range("M2:M10000").Copy wbinput.Sheets(i).Range("I2:I100000")
                                wbinput.Sheets(i).Range("M2:M10000").Copy wsc.Range("BC2:BC10000")
                                wbinput.Sheets(i).Range("L:M").Clear
                                
                                wsc.Range("BC:BC").RemoveDuplicates Columns:=1, Header:=xlYes
                                wsc.Columns("BC:BC").SpecialCells(xlCellTypeBlanks).Delete shift:=xlUp
                           
                                
                                
                           
'
                                 End If 'end if split brand
                                 
                             
            
                        
                    End If
                    
                 
                
                
                'sgl
                
                If wbinput.Sheets(i).Name = "SINGLE" Then
                  wsc.Range("BE2").Value = 1
                
                wbinput.Sheets(i).Range("A" & Excel.WorksheetFunction.CountA(wbinput.Sheets(i).Range("A:A")) + 1).Clear
                
                    wsc.Range("BB2").Value = Excel.WorksheetFunction.Sum(wbinput.Sheets(i).Range("E:E"))
                    
                     'remove space
                                            wbinput.Sheets(i).Range("H2:H1000").Replace What:=" ", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                                       ReplaceFormat:=False
                    
                    
                    ' split sgl
                            If fac = "SG" And wb_code <> "ST" And wb_code <> "BO" Then
                                 
                                   ' Sheets.Add after:=wbinput.Sheets(i)
                                    
                                wbinput.Sheets(i).Range("F:F").unmerge
                                wbinput.Sheets(i).Range("G:G").unmerge
                                
                                For C = 3 To Excel.WorksheetFunction.CountA(wbinput.Sheets(i).Range("C:C"))
                                    If wbinput.Sheets(i).Range("F" & C).Value = "" Then
                                        wbinput.Sheets(i).Range("F" & C).Value = wbinput.Range("F" & C - 1).Value
                                        
                                    End If
                                    
                                Next
                                
                                wbinput.Sheets(i).Columns("F:F").TextToColumns Destination:=Range("J1"), DataType:=xlDelimited, _
                                        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                                        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
                                        :="-", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
                                wbinput.Sheets(i).Range("K2:K10000").Copy wbinput.Sheets(i).Range("G2:G10000")
                                
                                wbinput.Sheets(i).Range("K2:K10000").Copy wsc.Range("BD1:BD9999")
                                wbinput.Sheets(i).Range("J:K").Clear
                                wsc.Range("BD:BD").RemoveDuplicates Columns:=1, Header:=xlYes
                           
'                                 ' brand
'                        'remove blank
                           wsc.Range("BD:BD").SpecialCells(xlCellTypeBlanks).Delete shift:=xlUp
                                 End If
                    
                    
                    
                End If
                
                          
                    
                Next
                
                wb_old_path = wbinput.Path & "/" & wbinput.Name
                
                'detect worker
                
                        Dim pcs As Integer
                        
                        pcs = wsc.Range("BB1").Value + wsc.Range("BB2").Value
                        
                                            
                            'check type worker
                            
                                'group1 & 4
                                    If fac = "SG" And wb_code <> "ST" And wb_code <> "BO" Then
                                    
                                    
                                    
                                    'non brand
                                    Workbooks.Add
                                    Dim wb_non As Workbook
                                    
                                    Set wb_non = ActiveWorkbook
                                    Sheets.Add after:=wb_non.Sheets("Sheet1")
                                 ActiveSheet.Name = "SINGLE"
                                 
                                    
                                    'rp
                                    Sheets.Add after:=ActiveSheet
                                    ActiveSheet.Name = "Total"
                                    Dim ws_non_tt As Worksheet
                                    Set ws_non_tt = wb_non.Sheets("Total")
                                    ws_non_tt.Range("A1").Value = "1"
                                    ws_non_tt.Range("B1").Value = "Single"
                                       ws_non_tt.Range("C1").Value = "Multi"
                                    ws_non_tt.Range("D1").Value = "Total"
                                   
                                    
                                    
                                    
                                    
                                    If wsc.Range("BE2").Value = 1 Then
                                    
                                    
                                wbinput.Sheets("SINGLE").Range("G:G").AutoFilter
                                 wbinput.Sheets("SINGLE").Range("G:G").AutoFilter field:=1, Criteria1:="="
                                 
                                 wbinput.Sheets("SINGLE").Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible).Copy wb_non.Sheets("SINGLE").Range("A1")
                                 wbinput.Sheets("SINGLE").Range("G:G").AutoFilter
                                 
                                 If Excel.WorksheetFunction.CountA(wb_non.Sheets("SINGLE").Range("C:C")) > 1 Then
                                 
                                   wb_non.Sheets("Sheet1").Range("A1").Value = Excel.WorksheetFunction.Sum(wb_non.Sheets("SINGLE").Range("E:E"))
                                   wb_non.Sheets("Sheet1").Range("B1").Value = Excel.WorksheetFunction.CountA(wb_non.Sheets("SINGLE").Range("F:F")) - 1
                                   
                                   'gen mokup
                                   
                                   wb_non.Sheets("SINGLE").Columns("B:B").ColumnWidth = 20
                                    wb_non.Sheets("SINGLE").Rows("2:" & Excel.WorksheetFunction.CountA(wb_non.Sheets("SINGLE").Range("C:C"))).RowHeight = 120
                                                
                                                            lr_mk_non = Excel.WorksheetFunction.CountA(wb_non.Sheets("SINGLE").Range("C:C"))
                                                            If lr_mk_non > 1 Then
                                                            
                                                                             Set Rng_non = wb_non.Sheets("SINGLE").Range("B2:B" & lr_mk_non)
                                                                            For Each cell In Rng_non
                                                
                                                     filenam = cell
                                                
                                                          
                                     pic_sgl_non = wb_non.Sheets("SINGLE").Shapes.AddPicture(Filename:=filenam, linktofile:=msfalse, savewithdocument:=msoTrue, Left:=cell.Left + (cell.Width - 100) / 2, Top:=cell.Top + (cell.Height - 100) / 2, Width:=100, Height:=100)
                                        
'
                                                                    Next
                                End If
                                                     
                                            wb_non.Sheets("SINGLE").Columns("D:H").EntireColumn.AutoFit
    
                                    'copy var
                                    wb_non.Sheets("SINGLE").Range("H2:H10000").Copy ws_non_tt.Range("A" & Excel.WorksheetFunction.CountA(ws_non_tt.Range("A:A")) + 1 & ":" & "A" & Excel.WorksheetFunction.CountA(ws_non_tt.Range("A:A")) + 10001)
                                    
                                          End If

                                    End If
                                    If wsc.Range("BE3").Value = 1 Then
                                         wbinput.Sheets("MULTI").Range("I:I").AutoFilter
                                 wbinput.Sheets("MULTI").Range("I:I").AutoFilter field:=1, Criteria1:="="
                                 Sheets.Add after:=wb_non.Sheets("Sheet1")
                                 
                                 ActiveSheet.Name = "MULTI"
                                 wbinput.Sheets("MULTI").Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible).Copy wb_non.Sheets("MULTI").Range("A1")
                                 wbinput.Sheets("MULTI").Range("I:I").AutoFilter
                                 If Excel.WorksheetFunction.CountA(wb_non.Sheets("MULTI").Range("D:D")) > 1 Then
                                   wb_non.Sheets("Sheet1").Range("A2").Value = Excel.WorksheetFunction.Sum(wb_non.Sheets("MULTI").Range("E:E"))
                                   wb_non.Sheets("MULTI").Range("H:H").Copy wb_non.Sheets("Sheet1").Range("C:C")
                                   wb_non.Sheets("Sheet1").Range("C:C").RemoveDuplicates Columns:=1, Header:=xlYes
                                   wb_non.Sheets("Sheet1").Range("B2").Value = Excel.WorksheetFunction.CountA(wb_non.Sheets("Sheet1").Range("C:C")) - 1
                                   
                                   wb_non.Sheets("MULTI").Columns("C:C").ColumnWidth = 20
                                                    wb_non.Sheets("MULTI").Rows("2:" & Excel.WorksheetFunction.CountA(wb_non.Sheets("MULTI").Range("D:D"))).RowHeight = 120
                                   
                                     lr_mk_non_mul = Excel.WorksheetFunction.CountA(wb_non.Sheets("MULTI").Range("C:C"))
                                     If lr_mk_non_mul > 1 Then
                                     
                                                 Set Rng_br_mul = wb_non.Sheets("MULTI").Range("C2:C" & lr_mk_non_mul)
                                                For Each cell In Rng_br_mul
                                                filenam = cell
                                                
                                                  pic_mul_non = wb_non.Sheets("MULTI").Shapes.AddPicture(Filename:=filenam, linktofile:=msfalse, savewithdocument:=msoTrue, Left:=cell.Left + (cell.Width - 100) / 2, Top:=cell.Top + (cell.Height - 100) / 2, Width:=100, Height:=100)
'
'
                                                                            Next
                                                                            End If
                                                    wb_non.Sheets("MULTI").Columns("E:J").EntireColumn.AutoFit
    
'
'copy var
                                    wb_non.Sheets("MULTI").Range("J2:J10000").Copy ws_non_tt.Range("A" & Excel.WorksheetFunction.CountA(ws_non_tt.Range("A:A")) + 1 & ":" & "A" & Excel.WorksheetFunction.CountA(ws_non_tt.Range("A:A")) + 10001)
                                    
                                                                            
                                   End If
                                   
                                    End If
                                    
                                    
                                    'save wb_non
                                     Dim pcs_non, pack_non As Integer
                                     pcs_non = wb_non.Sheets("Sheet1").Range("A2").Value + wb_non.Sheets("Sheet1").Range("A1").Value
                                      pack_non = wb_non.Sheets("Sheet1").Range("B2").Value + wb_non.Sheets("Sheet1").Range("B1").Value
                                      
                                      ur_code = wsc.Range("B1").Value
                                      Dim wb_non_name As String
                                      If ur_code = "" Then
                                        wb_non_name = Split(Split(wbinput.Name, ".")(0), "-")(0) & "-" & Split(Split(wbinput.Name, ".")(0), "-")(1) & "-" & _
                                      Split(Split(wbinput.Name, ".")(0), "-")(2) & "-" & pcs_non & "pcs-(NONBRAND)__" & Me.g1.Value & "__" & pcs_non & ".xlsx"
                                      
                                      Else
                                        wb_non_name = Split(Split(wbinput.Name, ".")(0), "_")(1) & "-" & Split(Split(wbinput.Name, ".")(0), "-")(0) & "-" & Split(Split(wbinput.Name, ".")(0), "-")(1) & "-" & _
                                      Split(Split(wbinput.Name, ".")(0), "-")(2) & "-" & pack_non & "pack-" & pcs_non & "pcs-(NONBRAND)__" & Me.g1.Value & "__" & pcs_non & ".xlsx"
                                      
                                      End If
                                      
                                      
                                      If Excel.WorksheetFunction.Sum(wb_non.Sheets("Sheet1").Range("A:A")) > 0 Then
                                      
                                      ws_non_tt.Range("A1").Value = Split(wb_non_name, ".")(0)
                                       ws_non_tt.Columns("A:A").EntireColumn.AutoFit
                                       
                                       wbinput.Sheets("TOTAL").Range("D2").Copy
                                    ws_non_tt.Range("A1:D1").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                              SkipBlanks:=False, Transpose:=False
                                    ws_non_tt.Range("A1:D1").Font.Bold = True
                                                        
                                      
                                      'count quantity
                                      
                                      
                                      With ws_non_tt
                                        .Range("A:A").RemoveDuplicates Columns:=1, Header:=xlYes
                                        If wsc.Range("BE2").Value = 1 Then
                                            If Excel.WorksheetFunction.CountA(wb_non.Sheets("SINGLE").Range("C:C")) > 1 Then
                                        
                                            .Range("B2").Value = "=SUMIF(SINGLE!H:H,A2,SINGLE!E:E)"
                                            
                                            Else
                                            wb_non.Sheets("SINGLE").Delete
                                            
                                            
                                           End If
                                            
                                        End If
                                        If wsc.Range("BE3").Value = 1 Then
                                           If Excel.WorksheetFunction.CountA(wb_non.Sheets("MULTI").Range("D:D")) > 1 Then
                                    
                                            .Range("c2").Value = "=SUMIF(MULTI!J:J,A2,MULTI!E:E)"
                                            
                                            Else
                                            
                                            wb_non.Sheets("MULTI").Delete
                                            
                                           End If
                                        End If
                                        .Range("D2").Value = "=C2+B2"
                                        
                                        If Excel.WorksheetFunction.CountA(ws_non_tt.Range("A:A")) > 2 Then
                                            .Range("B2:D" & Excel.WorksheetFunction.CountA(ws_non_tt.Range("A:A"))).FillDown
                                            
                                        End If
                                        
                                        
                                        .Range("B2:D" & Excel.WorksheetFunction.CountA(ws_non_tt.Range("A:A"))).Copy
                                        .Range("B2:D" & Excel.WorksheetFunction.CountA(ws_non_tt.Range("A:A"))).PasteSpecial xlPasteValues
                                        
                                        End With
                                        
                                      
                                       wb_non.Sheets("Sheet1").Delete
 
                                      wb_non.SaveAs Filename:= _
                                                            final_path & "/" & wb_non_name, FileFormat:= _
                                                            xlOpenXMLWorkbook, CreateBackup:=False
                                                            
                                        
                                        End If
                                        wb_non.Close
                                        
                                    
                                    'brand
                                    If Excel.WorksheetFunction.CountA(wsc.Range("BC:BC")) = 0 Then
                                        wsc.Range("BD1:BD10000").Copy wsc.Range("BC2:bc10001")
                                        
                                    Else
                                    
                                    
                                     wsc.Range("BD1:BD10000").Copy wsc.Range("BC" & Excel.WorksheetFunction.CountA(wsc.Range("BC:BC")) + 1 & ":" & "BC" & Excel.WorksheetFunction.CountA(wsc.Range("BC:BC")) + 10001)
                           wsc.Range("BC:BC").RemoveDuplicates Columns:=1, Header:=xlNo
                                    End If
                                    
                                    'detect sheet
                                    
                                    
                                    
                                     a = 2

                                                While (wsc.Range("BC" & a) <> "")
                                                        Workbooks.Add
                                                        Dim wb_br As Workbook
                                                        Set wb_br = ActiveWorkbook
                                                         'rp
                                                    Sheets.Add after:=ActiveSheet
                                                    ActiveSheet.Name = "Total"
                                                    Dim ws_br_tt As Worksheet
                                                    Set ws_br_tt = wb_br.Sheets("Total")
                                                    ws_br_tt.Range("A1").Value = "1"
                                                    
                                                    ws_br_tt.Range("B1").Value = "Single"
                                                       ws_br_tt.Range("C1").Value = "Multi"
                                                    ws_br_tt.Range("D1").Value = "Total"
                                                     
                                                        
                                              
                                                'filter br
                                                
                                                If wsc.Range("BE2").Value = 1 Then
                                                    If Excel.WorksheetFunction.CountIf(wbinput.Sheets("SINGLE").Range("G:G"), wsc.Range("BC" & a).Value) > 0 Then
                                                    
                                    
                                                wbinput.Sheets("SINGLE").Range("G:G").AutoFilter field:=1, Criteria1:=wsc.Range("BC" & a)
                                                wbinput.Sheets("SINGLE").Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible).Copy
                                                Sheets.Add after:=wb_br.Sheets("Sheet1")
                                                ActiveSheet.Name = "SINGLE"
                                                
                                                wb_br.Sheets("SINGLE").Range("A1").PasteSpecial xlPasteValues
                                                
                                                wbinput.Sheets("SINGLE").Range("G:G").AutoFilter
                                                 wb_br.Sheets("Sheet1").Range("A1").Value = Excel.WorksheetFunction.Sum(wb_br.Sheets("SINGLE").Range("E:E"))
                                                 wb_br.Sheets("Sheet1").Range("B1").Value = Excel.WorksheetFunction.CountA(wb_br.Sheets("SINGLE").Range("F:F")) - 1
                                                 
                                                'gen mokup
                                                wb_br.Sheets("SINGLE").Columns("B:B").ColumnWidth = 20
                                                    wb_br.Sheets("SINGLE").Rows("2:" & Excel.WorksheetFunction.CountA(wb_br.Sheets("SINGLE").Range("C:C"))).RowHeight = 120
                                                            lr_mk_br = Excel.WorksheetFunction.CountA(wb_br.Sheets("SINGLE").Range("C:C"))
'
                                                            If lr_mk_br > 1 Then
                                                            
                                                                             Set Rng_br = wb_br.Sheets("SINGLE").Range("B2:B" & lr_mk_br)
                                                                            For Each cell_br In Rng_br
                                                                            
'
                                                     filenam_br = cell_br
                                                    
                                                
                                                          
                                     pic_sgl_br = wb_br.Sheets("SINGLE").Shapes.AddPicture(Filename:=filenam_br, linktofile:=msfalse, savewithdocument:=msoTrue, Left:=cell_br.Left + (cell_br.Width - 100) / 2, Top:=cell_br.Top + (cell_br.Height - 100) / 2, Width:=100, Height:=100)
                                      
                                                                            Next
                                                                            End If
                                                        wb_br.Sheets("SINGLE").Columns("D:H").EntireColumn.AutoFit
    
'

'copy var
                                    wb_br.Sheets("SINGLE").Range("H2:H10000").Copy ws_br_tt.Range("A" & Excel.WorksheetFunction.CountA(ws_br_tt.Range("A:A")) + 1 & ":" & "A" & Excel.WorksheetFunction.CountA(ws_br_tt.Range("A:A")) + 10001)
                                    
                                    'paste format
                                    
                                            wbinput.Sheets("SINGLE").Range("A1").Copy
                                            wb_br.Sheets("SINGLE").Range("A1:H1").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                                SkipBlanks:=False, Transpose:=False
                                                                            
                                    
                                                                            
                                                    End If
                                                    
                            
                                                 End If
                                                 
                                                 
                                                 If wsc.Range("BE3").Value = 1 Then
                                                    If Excel.WorksheetFunction.CountIf(wbinput.Sheets("MULTI").Range("I:I"), wsc.Range("BC" & a).Value) > 0 Then
                                                    
                                                 wbinput.Sheets("MULTI").Range("I:I").AutoFilter field:=1, Criteria1:=wsc.Range("BC" & a)
                                                wbinput.Sheets("MULTI").Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible).Copy
                                                Sheets.Add after:=wb_br.Sheets("Sheet1")
                                                ActiveSheet.Name = "MULTI"
                                                
                                                wb_br.Sheets("MULTI").Range("A1").PasteSpecial xlPasteValues
                                                
                                                wbinput.Sheets("MULTI").Range("I:I").AutoFilter
                                                 wb_br.Sheets("Sheet1").Range("A2").Value = Excel.WorksheetFunction.Sum(wb_br.Sheets("MULTI").Range("E:E"))
                                                 wb_br.Sheets("MULTI").Range("H:H").Copy wb_br.Sheets("Sheet1").Range("C:C")
                                                 wb_br.Sheets("Sheet1").Range("C:C").RemoveDuplicates Columns:=1, Header:=xlYes
                                               wb_br.Sheets("Sheet1").Range("B2").Value = Excel.WorksheetFunction.CountA(wb_br.Sheets("Sheet1").Range("C:C")) - 1
                                                 wb_br.Sheets("MULTI").Columns("C:C").ColumnWidth = 20
                                                    Rows("2:" & Excel.WorksheetFunction.CountA(wb_br.Sheets("MULTI").Range("D:D"))).RowHeight = 120
                                                
                                                            lr_mk_br_mul = Excel.WorksheetFunction.CountA(wb_br.Sheets("MULTI").Range("C:C"))
                                                            
                                                            If lr_mk_br_mul > 1 Then
                                                                             Set Rng_br_mul = wb_br.Sheets("MULTI").Range("C2:C" & lr_mk_br_mul)
                                                                            For Each cell In Rng_br_mul
                                                                            filenam = cell
                                                                            wb_br.Sheets("MULTI").Pictures.Insert(filenam).Select
                                                                            Set Pshp_br_mul = Selection.ShapeRange.Item(1)
                                                                            If Pshp_br_mul Is Nothing Then GoTo lab_br_mul
                                                                            xCol_br_mul = cell.Column
                                                                            Set xRg_br_mul = Cells(cell.Row, xCol_br_mul)
                                                                            With Pshp_br_mul
                                                                            .LockAspectRatio = msoFalse
                                                                            .Width = 100
                                                                            .Height = 100
                                                                            .Top = xRg_br_mul.Top + (xRg_br_mul.Height - .Height) / 2
                                                                            .Left = xRg_br_mul.Left + (xRg_br_mul.Width - .Width) / 2
                                                                            
                                                                            End With
lab_br_mul:
                                                                            Set Pshp_br_mul = Nothing
                                                                            
                                                                            Next
                                                                End If
                                                wb_br.Sheets("MULTI").Columns("E:J").EntireColumn.AutoFit
    



'copy var
                                    wb_br.Sheets("MULTI").Range("J2:J10000").Copy ws_br_tt.Range("A" & Excel.WorksheetFunction.CountA(ws_br_tt.Range("A:A")) + 1 & ":" & "A" & Excel.WorksheetFunction.CountA(ws_br_tt.Range("A:A")) + 10001)
                                  
                                             wbinput.Sheets("MULTI").Range("A1").Copy
                                            wb_br.Sheets("MULTI").Range("A1:J1").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                                SkipBlanks:=False, Transpose:=False
                                                                            

                                                 End If
                                                 
                                                 End If
                                                 Dim pcs_br, pack_br As Integer
                                                 pcs_br = wb_br.Sheets("Sheet1").Range("A2").Value + wb_br.Sheets("Sheet1").Range("A1").Value
                                                 pack_br = wb_br.Sheets("Sheet1").Range("B2").Value + wb_br.Sheets("Sheet1").Range("B1").Value
                                                 'save br
                                                 
                                                 ur_code_br = wsc.Range("B1").Value
                                      If ur_code_br = "" Then
                                      Dim wb_br_name As String
                                        wb_br_name = Split(Split(wbinput.Name, ".")(0), "-")(0) & "-" & _
                                                  Split(Split(wbinput.Name, ".")(0), "-")(1) & "-" & Split(Split(wbinput.Name, ".")(0), "-")(2) & "-" & _
                                                   pcs_br & "pcs-(BR-" & wsc.Range("BC" & a).Value & ")__" & Me.g1.Value & "__" & pcs_br & ".xlsx"
                                                 
                                      Else
                                                 wb_br_name = Split(Split(wbinput.Name, ".")(0), "_")(1) & "-" & Split(Split(wbinput.Name, ".")(0), "-")(0) & "-" & _
                                                  Split(Split(wbinput.Name, ".")(0), "-")(1) & "-" & Split(Split(wbinput.Name, ".")(0), "-")(2) & "-" & _
                                                   pcs_br & "pcs-(BR-" & wsc.Range("BC" & a).Value & ")__" & Me.g1.Value & "__" & pcs_br & ".xlsx"
                                                 
                                        End If
                                        
                                                
                                          
'count quantity br
                                     With ws_br_tt

                                        .Range("A:A").RemoveDuplicates Columns:=1, Header:=xlYes
                                        If wsc.Range("BE2").Value = 1 Then
                                                    If Excel.WorksheetFunction.CountIf(wbinput.Sheets("SINGLE").Range("G:G"), wsc.Range("BC" & a).Value) > 0 Then
                                        
                                            .Range("B2").Value = "=SUMIF(SINGLE!H:H,A2,SINGLE!E:E)"
                                            End If
                                        End If
                                         If wsc.Range("BE3").Value = 1 Then
                                                    If Excel.WorksheetFunction.CountIf(wbinput.Sheets("MULTI").Range("I:I"), wsc.Range("BC" & a).Value) > 0 Then
                                    
                                            .Range("c2").Value = "=SUMIF(MULTI!J:J,A2,MULTI!E:E)"
                                            End If
                                        End If
                                        .Range("D2").Value = "=C2+B2"
                                        
                                        If Excel.WorksheetFunction.CountA(ws_br_tt.Range("A:A")) > 2 Then
                                            .Range("B2:D" & Excel.WorksheetFunction.CountA(ws_br_tt.Range("A:A"))).FillDown
                                            
                                        End If
                                        
                                       
                                                .Range("A1").Value = Split(wb_br_name, ".")(0)
                                                .Columns("A:A").EntireColumn.AutoFit
                                                              .Range("A1:D1").Font.Bold = True
                                                              wbinput.Sheets("TOTAL").Range("D2").Copy
                                                    .Range("A1:D1").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                                                        SkipBlanks:=False, Transpose:=False
                                                .Range("A1:D1").Font.Bold = True
                                                
                                                .Range("B2:D" & Excel.WorksheetFunction.CountA(ws_br_tt.Range("A:A"))).Copy
                                                .Range("B2:D" & Excel.WorksheetFunction.CountA(ws_br_tt.Range("A:A"))).PasteSpecial xlPasteValues
                                          End With
                                                        wb_br.Sheets("Sheet1").Delete
                                                         wb_br.SaveAs Filename:= _
                                                            final_path & "/" & wb_br_name, FileFormat:= _
                                                            xlOpenXMLWorkbook, CreateBackup:=False
                                                
                                                
                                                wb_br.Close
                                                a = a + 1
                                                Wend
                                        
                                                   
                                        
                                        'sg + st bo : Group4
                                        
                                         ElseIf fac = "SG" And (wb_code = "ST" Or wb_code = "BO") Then
                                                 wb_name = Split(wbinput.Name, ".")(0) & "__" & Me.txt4.Value & "__" & psc & ".xlsx"
                                                wsc.Delete
                                          wbinput.SaveAs Filename:= _
                                                            final_path & "/" & wb_name, FileFormat:= _
                                                            xlOpenXMLWorkbook, CreateBackup:=False
                                                            
                                         
                                         
                                         
                                          ElseIf fac <> "SG" Then
                                                    Dim pcs_tq_total As Integer
                                                    'cal pcs sgl
                                                        If wsc.Range("BE2").Value = 1 Then
                                                            If wbinput.Sheets("SINGLE").Range("E1").Value = "Quantity" Then
                                                                wsc.Range("BE4").Value = Excel.WorksheetFunction.Sum(wbinput.Sheets("SINGLE").Range("E2:E10000"))
                                                            Else
                                                                If wbinput.Sheets("SINGLE").Range("F1").Value = "Quantity" Then
                                                                wsc.Range("BE4").Value = Excel.WorksheetFunction.Sum(wbinput.Sheets("SINGLE").Range("F2:F10000"))
                                                                End If
                                                            End If
                                                            
                                                        End If
                                                        
                                                        If wsc.Range("BE3").Value = 1 Then
                                                            
                                                            wsc.Range("BE5").Value = Excel.WorksheetFunction.Sum(wbinput.Sheets("MULTI").Range("E2:E10000"))
                                                        End If
                                                        
                                                    pcs_tq_total = wsc.Range("BE4").Value + wsc.Range("BE5").Value
                                                    Debug.Print ("check LC: " & wb_code)
                                                    
                                                    If wb_code = "CR" And wsc.Range("BE2").Value = 1 Then
                                                            wbinput.Sheets("SINGLE").Range("E:E").Copy wsc.Range("CC:CC")
                                                            wsc.Range("CC:CC").Replace What:=" ", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                                       ReplaceFormat:=False
                                                        wsc.Range("CC:CC").Replace What:="CROCS", Replacement:="CR", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                                       ReplaceFormat:=False
                                                        wsc.Range("CC:CC").Copy wbinput.Sheets("SINGLE").Range("I:I")
                                                      'add size to printcode
                                                      wbinput.Sheets("SINGLE").Range("G:G").Copy wsc.Range("CA:CA")
                                                        wsc.Range("CC:CC").Replace What:="CR-", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                                       ReplaceFormat:=False
                                                          p = 2
                                                          While wsc.Range("CA" & p).Value <> ""
                                                                wsc.Range("CA" & p).Value = wsc.Range("CC" & p).Value & "-" & wsc.Range("CA" & p).Value
                                                            p = p + 1
                                                          Wend
                                                          
                                                        wsc.Range("CA:CA").Replace What:="-N", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                                       ReplaceFormat:=False
                                                          
                                                    wsc.Range("CA:CA").Copy wbinput.Sheets("SINGLE").Range("G:G")
                                                          wbinput.Sheets("SINGLE").Range("I1").Value = "Variant Code"
                                                          
                                                          
                                                          
                                                          
                                                    
                                                    
                                                    End If
                                                    
                                                    
                                                    
          
                                                       ' MsgBox wb_code
                                                        wsc.Delete
                                                            If Excel.WorksheetFunction.CountIf(ws_m7.Range("A:A"), var_code) > 0 Then
                                                                 'group 3
                                                                
                                                                   wb_name = Split(wbinput.Name, ".")(0) & "__" & Me.g3.Value & "__" & pcs_tq_total & ".xlsx"
                                                                   
'                                                            ElseIf fac = "HAIAN" And (wb_code = "ST" Or wb_code = "BO") Then
'
'                                                            'group 5:
'                                                                 wb_name = Split(wbinput.Name, ".")(0) & "__" & Me.g5.Value & "__" & pcs_tq_total & ".xlsx"
                                                            Else
                                                            
                                                                'group 2
                                                               
                                                               
                                                                wb_name = Split(wbinput.Name, ".")(0) & "__" & Me.g2.Value & "__" & pcs_tq_total & ".xlsx"
                                                                
                                                        End If
                                                        wbinput.SaveAs Filename:= _
                                                            final_path & "/" & wb_name, FileFormat:= _
                                                            xlOpenXMLWorkbook, CreateBackup:=False
                                                            
                
                                                  End If
                                      
                            
                'del
                
                                       
                    
                wbinput.Close
                Kill (wb_old_path)
                    
            
                
Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.ErrorCheckingOptions.BackgroundChecking = True

          
                     UserForm1.Hide
                     

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


