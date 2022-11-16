Attribute VB_Name = "MiExcel"
Option Explicit

Public Sub GenerarExcel(rs As Recordset)
On Error Resume Next
    Dim Excel As Object
    Dim libro As Object
    Dim hoja As Object
    
    Set Excel = CreateObject("Excel.Application")
    Set libro = Excel.Workbooks.Add
    Excel.Visible = True
    
    
    Set hoja = libro.Worksheets(1)
    hoja.Cells(1, 1).CopyFromRecordset rs
'    hoja.Range("A1").Value = "que pedo puto"
'
'    libro.SaveAs "C:\mamalon.xls"
'    Excel.Quit
End Sub
