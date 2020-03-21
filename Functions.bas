Attribute VB_Name = "Functions"
Option Explicit

Function isWbkOpenXLSM(s_wbk As String) As Boolean

    isWbkOpenXLSM = True
    
    On Error GoTo Fail
    Workbooks(s_wbk & ".xlsm").Activate
    
    Exit Function
    
Fail:
isWbkOpenXLSM = False
End Function

Function sheetexists(s_wks As String, Optional wbk As Workbook) As Boolean
    
    If wbk Is Nothing Then Set wbk = ThisWorkbook
    sheetexists = False
    
    On Error Resume Next
    
    sheetexists = wbk.Sheets(s_wks).Index > 0
    
End Function

Function isWbkOpenXLSM_2(s_wbk As Range) As Boolean

    isWbkOpenXLSM_2 = True
    
    On Error GoTo Fail
    Workbooks(s_wbk.Value & ".xlsm").Activate
    
    Exit Function
    
Fail:
isWbkOpenXLSM_2 = False
End Function

Function CorrectRange(r As Range) As Boolean
    
    If r.Value = "" Then GoTo Fail
    
    Dim rng As Range
    CorrectRange = True
    
    On Error GoTo Fail
    
    Set rng = Range(r.Value)
    
    
    Exit Function
    
Fail:
    CorrectRange = False
End Function

