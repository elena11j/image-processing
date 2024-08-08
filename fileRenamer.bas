Attribute VB_Name = "Module1"
Option Explicit

Sub list_file()
    Dim wb As Workbook
    Dim wb_path As String
    Dim file As String
    Dim i As Integer
    Dim filled_rng As Range
    Dim ext As String
    
    
    Set wb = ThisWorkbook
    Set filled_rng = Range(Range("B3:C3"), Range("B3:C3").End(xlDown))
    wb_path = wb.Path & "\"
    
    filled_rng.ClearContents
    
    ext = Range("exten").Value
    file = Dir(wb_path)
    i = 3
    While file <> ""
        If ext = "" Then
            ext = "*"
        End If
        If file Like "*." & ext Then
            Cells(i, 2).Value = file
            i = i + 1
        End If
        file = Dir
    Wend
    
    Range(Range("B3:E3"), Range("B3:E3").End(xlDown)).Sort key1:=Range("E3"), order1:=xlAscending, Header:=xlNo
    
    'Range("C3").FormulaR1C1 = _
        "=IF(RC[-1]="""","""",IFERROR(ROW()-2&"" ""&MID(RC[-1],SEARCH("" "",RC[-1])+1,LEN(RC[-1])-SEARCH("" "",RC[-1])-LEN(RC[3])-1),""""))"
    'Range("C3").AutoFill Destination:=Range("C3:C201")

    
End Sub

Sub rename_files()
    Dim wb As Workbook
    Dim wb_path As String
    Dim files As Range
    Dim r As Range
    Dim new_name As String
    Dim i As Integer
    
    Set wb = ThisWorkbook
    Set files = Range(Range("B3"), Range("B3").End(xlDown))
    wb_path = wb.Path & "\"
    
    i = 0
    For Each r In files.Rows
        If Cells(r.Row, 3).Value <> "" Then
            new_name = Cells(r.Row, 4).Value
            Name wb_path & Cells(r.Row, 2).Value As wb_path & Cells(r.Row, 4).Value
            i = i + 1
        End If
    Next r
    
    Call list_file
    MsgBox i & " file(s) renamed!", vbInformation
    
End Sub
