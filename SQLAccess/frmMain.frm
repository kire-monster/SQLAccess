VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sql Connect"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtQuery 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   3015
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmMain.frx":0000
      Top             =   0
      Width           =   7815
   End
   Begin VB.CommandButton btnEjecutar 
      Caption         =   "Ejecutar"
      Height          =   615
      Left            =   7920
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnEjecutar_Click()
    Dim con As Connection
    Dim rs As Recordset
    Set con = New Connection
    Set rs = New Recordset
    
    'abrimos la conexion
    openDB con, "PROVIDER=SQLOLEDB; DATA SOURCE=192.168.100.96; UID=sa; PWD=789512357456abcdfG1;DATABASE=Pruebax;" ' DATABASE=__DB__
    If Err.Description <> "" Then
        MsgBox "Error al conectarse, Detalle: " + Err.Description
        Exit Sub
    End If
    
    'configuramos el objeto de recorrido
'    rs.CursorLocation = adUseClient
'    rs.CursorType = adOpenStatic
'    rs.LockType = adLockBatchOptimistic
    
    'ejecutamos la consulta
    execSQL con, rs, txtQuery.Text
    If Err.Description <> "" Then
        MsgBox "Error al ejecutar SQL, Detalle: " + Err.Description
        Exit Sub
    Else
'        rs.MoveFirst
        GenerarExcel rs
        If Err.Description <> "" Then
            MsgBox "Error al generar el excel, Detalle: " + Err.Description
            Exit Sub
        Else
            MsgBox "termino"
        End If
'        MsgBox "termino"
    End If
    
    
    closeDB con
    If Err.Description <> "" Then
        MsgBox "Error al desconectarse, Detalle: " + Err.Description
        Exit Sub
    End If
End Sub
