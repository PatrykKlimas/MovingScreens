Attribute VB_Name = "Macros"
Option Explicit

Sub screen_moving()

    Dim template As Worksheet
    Dim col As Object
    Set col = CreateObject("Scripting.Dictionary")
    Dim shp As Shape
    Dim rng As Range
    Dim EGA As Workbook
    Dim i As Integer: i = 0
    
    On Error GoTo screen_moving
    Set template = ThisWorkbook.Sheets("Screen moving")
    On Error GoTo 0
    
    Set rng = template.Range("A2")
    
    On Error GoTo multi_shape
    For Each shp In template.Shapes
        col.Add shp.TopLeftCell.Address, shp
    Next
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    Do While rng.Offset(i, 0).Value <> ""
    
        If isWbkOpenXLSM(rng.Offset(i, 0).Value) Then
        
            Set EGA = Workbooks(rng.Offset(i, 0).Value & ".xlsm")
            
            If sheetexists(rng.Offset(i, 1).Value, EGA) Then
                
                If col.Exists(rng.Offset(i, 3).Address) Then
                    
                    col.Item(rng.Offset(i, 3).Address).Copy
                    EGA.Sheets(rng.Offset(i, 1).Value).Select
                    Range(rng.Offset(i, 2).Value).Select
                    ActiveSheet.Paste
                    
                    'Ustalic rozmiar
                    Selection.ShapeRange.Width = 500
                End If
            Else
            
                With rng.Offset(i, 1)
                    .AddComment
                    .Comment.Text Text:="Sheet does not exist!"
                    .Comment.Text
                    .Comment.Visible = True
                End With
                
            End If
            
        End If
        i = i + 1
        
    Loop
    
    template.Activate
    Exit Sub
screen_moving:
    MsgBox "'Screen moving' tab does not exists!", vbCritical, "Error!"
    
multi_shape:
    MsgBox "There are some cells with more then one shape! Comment is also treated as a shape.", vbCritical, "Error!"
End Sub


Sub raports_moving()
    Dim template As Worksheet
    Dim EGA As Workbook
    Dim EGA_wst As Worksheet
    Dim rng As Range
    Dim i As Integer: i = 0
    
    Application.ScreenUpdating = False
    
    On Error GoTo raport_moving
    Set template = ThisWorkbook.Sheets("Raport Moving")
    On Error GoTo 0
    
    Set rng = template.Range("A2")
    
    Do While rng.Offset(i, 0).Value <> ""
        If isWbkOpenXLSM(rng.Offset(i, 0).Value) Then
            Set EGA = Workbooks(rng.Offset(i, 0).Value & ".xlsm")
            
            If Not (sheetexists(rng.Offset(i, 3).Value, EGA)) Then
                EGA.Sheets.Add After:=EGA.Sheets(EGA.Sheets.Count)
                EGA.Sheets(EGA.Sheets.Count).Name = rng.Offset(i, 3).Value
            End If
            
            Set EGA_wst = EGA.Sheets(rng.Offset(i, 3).Value)
            EGA_wst.Cells.ClearContents
            
            If sheetexists(rng.Offset(i, 1).Value, ThisWorkbook) Then
                
                ThisWorkbook.Sheets(rng.Offset(i, 1).Value).Range(rng.Offset(i, 2).Value).Copy
                EGA_wst.Range(rng.Offset(i, 4).Value).PasteSpecial Paste:=xlPasteValues
            End If
        End If
        i = i + 1
    Loop
    
    Exit Sub
    
raport_moving:
    MsgBox "'Raport Moving' tab does not exists!", vbCritical, "Error!"
End Sub

Sub filtering()
    Dim template As Worksheet
    Dim EGA As Workbook
    Dim EGA_wst As Worksheet
    Dim wst_to_copy As Worksheet
    Dim rng As Range
    Dim i As Integer: i = 0
    Dim k As Integer
    Dim i_rows As Integer
    
    Application.ScreenUpdating = False
    On Error GoTo filter_raport
    Set template = ThisWorkbook.Sheets("Raport Filter")
    On Error GoTo 0
    
    Set rng = template.Range("A2")
    
    Do While rng.Offset(i, 0).Value <> ""
        If isWbkOpenXLSM(rng.Offset(i, 0).Value) Then
            Set EGA = Workbooks(rng.Offset(i, 0).Value & ".xlsm")
            If sheetexists(rng.Offset(i, 2).Value, EGA) Then
            
                k = 0
                
                Set EGA_wst = EGA.Sheets(rng.Offset(i, 2).Value)
                Set wst_to_copy = ThisWorkbook.Sheets(rng.Offset(i, 1).Value)
                
                wst_to_copy.Range("A:CX").AutoFilter
                
                Do While rng.Offset(i + 3, k).Value <> ""
                
                    wst_to_copy.Range("A:CX").AutoFilter _
                        Field:=CInt(rng.Offset(i + 4, k).Value), _
                        Criteria1:=rng.Offset(i + 3, k).Value
                        k = k + 1
                        
                Loop
                k = Application.WorksheetFunction.CountA(wst_to_copy.Range("A:A"))
                i_rows = Application.WorksheetFunction.CountA(wst_to_copy.Range("A:A").SpecialCells(xlCellTypeVisible))
                
                EGA_wst.Rows((rng.Offset(i, 3).Value + 1) & ":" & (rng.Offset(i, 3).Value + i_rows - 2)).Insert Shift:=xlUp
                
                wst_to_copy.Range(rng.Offset(i + 6, 0).Value & "2:" & rng.Offset(i + 7, 0).Value & k).Copy
                EGA_wst.Range(rng.Offset(i + 6, 1).Value & rng.Offset(i, 3).Value).PasteSpecial Paste:=xlPasteValues
                
            Else
                With rng.Offset(i, 2)
                    .AddComment
                    .Comment.Text Text:="Sheet does not exist!"
                    .Comment.Text
                    .Comment.Visible = True
                End With
            End If
        
        End If
        i = i + 10
    Loop
        
  Exit Sub
filter_raport:
    MsgBox "'Raport Filter' tab does not exists!", vbCritical, "Error!"
End Sub
