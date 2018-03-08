VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLoginManterDadosNuvem 
   Caption         =   "Login"
   ClientHeight    =   2160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4050
   OleObjectBlob   =   "frmLoginManterDadosNuvem.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLoginManterDadosNuvem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEntrar_Click()

On Error GoTo Erro

Dim cnn As New ADODB.Connection

    'cnn.ConnectionString = "Driver={ODBC Driver 13 for SQL Server};Server=tcp:contarcondb.cmxd2lqddzlw.sa-east-1.rds.amazonaws.com,1433;Database=fluxocaixa;Uid=" & Me.txtUsua.Text & ";Pwd={" & Me.txtSenha.Text & "};Encrypt=yes;TrustServerCertificate=no;Connection Timeout=300;"
    cnn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & Me.txtSenha.Text & ";Persist Security Info=True;User ID=" & Me.txtUsua.Text & ";Initial Catalog=fluxocaixa;Data Source=contarcondb.cmxd2lqddzlw.sa-east-1.rds.amazonaws.com,1433"
    cnn.Open
    
    MsgBox "Acesso permitido!", vbOKOnly, "Fluxo de Caixa"
    cnn.Close
    
    manterDadosAposLogin = True
    
    usuario = Me.txtUsua.Text
    senha = Me.txtSenha.Text
    
    Unload Me
    
    Exit Sub
    
Erro:

    MsgBox "Erro de acesso. Confira o usuário e a senha.", vbOKOnly, "Fluxo de Caixa"
    manterDadosAposLogin = False
    Exit Sub
    
End Sub

Private Sub cmdFechar_Click()

    Unload Me
    manterDadosAposLogin = False
    'ActiveWorkbook.Close SaveChanges:=False

End Sub


Private Sub UserForm_Click()

End Sub
