VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBarraProgressaoImportacao 
   Caption         =   "Processamento de Cenário e Importação de Dados"
   ClientHeight    =   1845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8325
   OleObjectBlob   =   "frmBarraProgressaoImportacao.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBarraProgressaoImportacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Activate()

    lblProgresso.Width = 0
    Call frmImportarPlanilhaComParametro.SalvarImportacao
    lblProgresso.Width = 0
    Call frmImportarPlanilhaComParametro.ProcessaImportacao
    
End Sub

Sub AtualizaBarra(percentual As Single, informacao As String)

 With frmBarraProgressaoImportacao
    .nomeQuadro.Caption = "Processando... " + informacao + "... " + CStr(Format(percentual, "00%"))
    .lblProgresso.Width = percentual * (.nomeQuadro.Width - 10)
 End With

 DoEvents
End Sub
