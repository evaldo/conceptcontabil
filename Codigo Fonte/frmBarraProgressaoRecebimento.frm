VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBarraProgressaoRecebimento 
   Caption         =   "Processamento de Recebimento de Caixa"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8280
   OleObjectBlob   =   "frmBarraProgressaoRecebimento.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBarraProgressaoRecebimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Activate()

    lblProgresso.Width = 0
    
    Call processa_recebimento_caixa
    
End Sub

Sub AtualizaBarra(percentual As Single, informacao As String)
    
    With frmBarraProgressaoRecebimento
       
       .nomeQuadro.Caption = "Processando... " + informacao + "... " + CStr(Format(percentual, "00%"))
       .lblProgresso.Width = percentual * (.nomeQuadro.Width - 10)
       
    End With
    
End Sub
