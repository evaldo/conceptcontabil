VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEscolhaLancamento 
   Caption         =   "Escolha da Planilha de Lançamento de Fluxo de Caixa"
   ClientHeight    =   1590
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4875
   OleObjectBlob   =   "frmEscolhaLancamento.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEscolhaLancamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFechar_Click()

    Unload Me

End Sub

Private Sub cmdPlanilhaEscolhida_Click()

On Error GoTo Erro

    Sheets(cmbListagemMes.Text).Activate
    Unload Me
    
    Exit Sub
    
Erro:

    MsgBox "Erro ao escolher o mês para lançamento.", vbOKOnly + vbInformation, "Erro ao escolher o mês para lançamento"
    Exit Sub

End Sub

Private Sub UserForm_Activate()

Dim mes(1 To 12) As String
    
    mes(1) = "Jan"
    mes(2) = "Fev"
    mes(3) = "Mar"
    mes(4) = "Abr"
    mes(5) = "Mai"
    mes(6) = "Jun"
    mes(7) = "Jul"
    mes(8) = "Ago"
    mes(9) = "Set"
    mes(10) = "Out"
    mes(11) = "Nov"
    mes(12) = "Dez"


    cmbListagemMes.List = mes

End Sub


