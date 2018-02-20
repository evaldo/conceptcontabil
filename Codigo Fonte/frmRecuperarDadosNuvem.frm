VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRecuperarDadosNuvem 
   Caption         =   "Recuperar Dados da Nuvem"
   ClientHeight    =   5460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4590
   OleObjectBlob   =   "frmRecuperarDadosNuvem.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRecuperarDadosNuvem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nomePlanilha As String

Private Sub cmdFechar_Click()

    Unload Me

End Sub

Private Sub cmdRecuperarDados_Click()

On Error GoTo Erro

    cnn.ConnectionString = "Driver={ODBC Driver 13 for SQL Server};Server=tcp:contarcondb.cmxd2lqddzlw.sa-east-1.rds.amazonaws.com,1433;Database=fluxocaixa;Uid=evaldo;Pwd={Gcas1302};Encrypt=yes;TrustServerCertificate=no;Connection Timeout=300;"
    cnn.Open
    
    Worksheets("Configurações Básicas").Activate
    ano = Range("E5").Value
    nomeClie = Range("E9").Value
    cnpjClie = Range("E8").Value
    
    frmProgresso.Visible = True
    
    Call barraProgresso("Processando a recuperação de dados...", 1)
    
    Application.ScreenUpdating = False
    
    '----------------------------------------------------------------------------------------------
    'Códigos SELECT/SQL para recuperação da base de dados, caso as opções estejam selecionadas.
    '----------------------------------------------------------------------------------------------
    'escrever aqui.
    '
    'Lembrete: não esquecer de armazenar na rotina de inserir dados, os dados da exportação de
    'plano de contas.
    '----------------------------------------------------------------------------------------------
    
    
    Application.ScreenUpdating = True
    frmProgresso.Visible = False
    
    Exit Sub
    
Erro:
    
    MsgBox "Erro ao recuperar os dados selecionados. Refaça a operação.", vbOKOnly, "Recuperar dados da nuvem"
    frmProgresso.Visible = False
    Exit Sub

End Sub

Private Sub UserForm_Initialize()
    
Dim todosMeses(1 To 12) As String
    
    nomePlanilha = ActiveSheet.Name
    
    todosMeses(1) = "Jan"
    todosMeses(2) = "Fev"
    todosMeses(3) = "Mar"
    todosMeses(4) = "Abr"
    todosMeses(5) = "Mai"
    todosMeses(6) = "Jun"
    todosMeses(7) = "Jul"
    todosMeses(8) = "Ago"
    todosMeses(9) = "Set"
    todosMeses(10) = "Out"
    todosMeses(11) = "Nov"
    todosMeses(12) = "Dez"
    
    Me.lstMeses.List = todosMeses
    
End Sub

Public Sub barraProgresso(mensagem As String, percentual As Integer)

    Me.lblDescricaoProgresso.Caption = mensagem + "... " + CStr(percentual) + " registros"
    DoEvents
    Me.lblProgresso.Width = ((percentual / 10000) * Me.lblDescricaoProgresso.Width)
    DoEvents
    
End Sub
