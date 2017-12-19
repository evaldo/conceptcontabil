Attribute VB_Name = "Importacao"
Sub importar_Com_Parametro()
    
    resposta = MsgBox("Deseja realmente processar a importação com Parâmetros?", vbYesNo + vbExclamation, "Processamento de Recebimentos")
 
    If resposta = vbYes Then frmEscolhaDesRec.Show
    
End Sub
