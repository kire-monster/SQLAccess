Attribute VB_Name = "SqlAccess"
Option Explicit

'Variables Globales
'Public ConDB As Connection
'Public RsADO As Recordset
'"PROVIDER=SQLOLEDB; DATA SOURCE=__IP__; UID=__USR__; PWD=__PWD__; DATABASE=__DB__"


Public Sub execSQL(db As Connection, rs As Recordset, SQL As String)
On Error Resume Next
    rs.Open SQL, db, adOpenStatic, adLockOptimistic, adAsyncConnect
    'Se añade ejecucion asincrona
    Do While rs.State = adStateExecuting
        If Err.Description <> "" Then GoTo Error_salir 'Exit Sub
        If rs.State = 0 Then GoTo Error_salir 'Exit Sub
        DoEvents
    Loop
    Exit Sub
Error_salir:
    MsgBox "Error => Estado: " & rs.State & " _ Error: " & Err.Description
    Exit Sub
End Sub



Public Sub openDB(db As Connection, connString As String)
On Error Resume Next
  db.CommandTimeout = 120
  
  db.Open connString, , , adAsyncConnect
  
  Do While db.State = adStateConnecting
    If Err.Description <> "" Then GoTo hError
    If db.State = 0 Then Exit Sub
    DoEvents
  Loop
  Exit Sub
hError:
    MsgBox "Error => Number: " & Err.Number & " Description: " & Err.Description
End Sub



Public Sub closeDB(db As Connection)
  Do While db.State > 1
    DoEvents
  Loop
  If db.State Then db.Close
  Set db = Nothing
End Sub
