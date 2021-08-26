Sub Hide_completed()
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    For Each c In Range(Cells(4, 19), Cells(lastrow, 19)).Cells  'Set Range (S4: S_lastRow)
        If c.Value = "Complete" Then
            c.EntireRow.Hidden = True
        End If
    Next c
End Sub
Sub show_all()
    Columns.EntireColumn.Hidden = False
    Rows.EntireRow.Hidden = False
    If (ActiveSheet.AutoFilterMode And ActiveSheet.FilterMode) Or ActiveSheet.FilterMode Then
        ActiveSheet.ShowAllData
    End If
    Call Scroll_last_row
End Sub
Sub HH_report_incomplete()
With ActiveSheet.Range("$A$1")
.AutoFilter Field:=5, Criteria1:="HH"
.AutoFilter Field:=8, Criteria1:=""
End With
End Sub
Sub GW_report_incomplete()
With ActiveSheet.Range("$A$1")
.AutoFilter Field:=5, Criteria1:="GW"
.AutoFilter Field:=8, Criteria1:=""
End With
End Sub
Sub Scroll_last_row()
Dim last_row As Integer
last_row = Cells(Rows.Count, 1).End(xlUp).Row
ActiveSheet.Cells(last_row, 1).Select
'With ActiveSheet
'    .Cells(.Rows.Count, ActiveCell.Column).End(xlUp).Select
'End With
End Sub

Function SubRe(Value As String, Pattern As String, ReplaceWith As String, Optional IgnoreCase As Boolean = False)
    Dim r As New VBScript_RegExp_55.RegExp
    r.Pattern = Pattern
    r.IgnoreCase = IgnoreCase
    r.Global = True
    SubRe = r.Replace(Value, ReplaceWith)
End Function


Sub import_json()
    Dim MyData As String
    Dim lineData() As String, strData() As String, myFile As String, ReplacedText As String
    Dim pat As String
    Dim i As Long, rng As Range

    ' lets make it a little bit easier for the user
    myFile = "C:\Users\H_Han\Desktop\Python_scripts\SmartDir1.1\dist\Index_latest.json"

    Open myFile For Binary As #1
    MyData = Space$(LOF(1))
    Get #1, , MyData
    Close #1
    ' Split into wholes line
    lineData() = Split(MyData, vbNewLine)
    Set rng = Sheets("dict").Range("A1")
    ' For each line
    ' Chr(34) is identical to quotemark "
    pat = "\s+" & Chr(34) & "([^""]+)" & Chr(34) & ": " & Chr(34) & "([^""]+)" & Chr(34) & ",?"
    For i = 0 To UBound(lineData)
        ' Split the line
        ReplacedText = SubRe(lineData(i), pat, "$1:$2", False)
        ReplacedText = Replace(ReplacedText, "\\", "\")
        ' Write to the sheet
        If Len(ReplacedText) > 1 Then
            strData = Split(ReplacedText, ":", 2)
            rng.Offset(i, 0).Resize(1, UBound(strData) + 1) = strData
        End If
    Next
    MsgBox "JSON Dictionary Successfully Reloaded!", vbInformation
End Sub

