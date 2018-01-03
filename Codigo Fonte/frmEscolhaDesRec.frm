VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEscolhaDesRec 
   Caption         =   "Escolha do Tipo de Importação (Despesas ou Receitas)"
   ClientHeight    =   1665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5010
   OleObjectBlob   =   "frmEscolhaDesRec.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEscolhaDesRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bolClassificacaoReceita As Boolean
Public bolClassificacaoDespesa As Boolean

Private Sub cmdEscolhaTipoClassificacao_Click()
    
    bolClassificacaoReceita = False
    bolClassificacaoDespesa = False
    
    If optClassificacaoReceita.Value = True Then
        bolClassificacaoReceita = True
    Else
        bolClassificacaoDespesa = True
    End If
    
    frmImportarPlanilhaComParametro.Show
    
End Sub

Private Sub cmdFechar_Click()

    Unload Me

End Sub

Public Sub importar_Com_Parametro()
    
    resposta = MsgBox("Deseja realmente processar a importação com Parâmetros?", vbYesNo + vbExclamation, "Processamento de Recebimentos")
 
    If resposta = vbYes Then frmEscolhaDesRec.Show
    
End Sub


