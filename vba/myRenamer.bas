Attribute VB_Name = "Module11"
Option Explicit

Sub list_file()
    Dim wb As Workbook
    Dim wb_path As String
    Dim file As String
    Dim i As Integer
    Dim filled_rng As Range
    Dim ext As String
        
    Set wb = ThisWorkbook
    Set filled_rng = Range(Range("B2", "L2"), Range("B2", "L2").End(xlDown))
    wb_path = wb.Path & "\"

    file = Dir(wb_path)
    
End Sub

Sub rename_files()
    Dim wb As Workbook
    Dim wb_path As String
    Dim files As Range
    Dim r As Range
    Dim new_name As String
    Dim i As Integer
    
    Set wb = ThisWorkbook
    Set files = Range(Range("B2"), Range("B2").End(xlDown))
    wb_path = wb.Path & "\"
    
    i = 0
    For Each r In files.Rows
        Name wb_path & Cells(r.Row, 2).Value As wb_path & Cells(r.Row, 12).Value
        i = i + 1
    Next r
    
    Call list_file
    MsgBox i & " file(s) renamed!", vbInformation
    
End Sub
