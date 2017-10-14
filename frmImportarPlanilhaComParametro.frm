VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImportarPlanilhaComParametro 
   Caption         =   "Importar Dados de Planilhas"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9420
   OleObjectBlob   =   "frmImportarPlanilhaComParametro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmImportarPlanilhaComParametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim itemListaClassificacao As Integer
Dim mes_processamento  As String
Dim WB1 As Workbook

Private Sub btnAtualizaClassificacao_Click()

    lstClassificacao.List(itemListaClassificacao, 1) = cmbClassificacao.Text

End Sub

Private Sub btnCarregaDados_Click()
 
    Dim classificacao(1 To 1000, 1 To 2) As String
        
    Dim i As Integer, j As Integer
    Dim linha As Integer
            
    If txtCaminhoPlanilha.Text <> "" Then
        
        mes_processamento = ActiveSheet.Name
        
        classificacao(1, 1) = "Classificação Importada"
        classificacao(1, 2) = "Classificação Utilizada"
            
        Set WB1 = Workbooks.Open(txtCaminhoPlanilha.Text)
        
        cmbClassificacao.Clear
        lstClassificacao.Clear
                
        If txtLinhaInicial.Text <> "" Then
        
            If txtColunaClassificacao.Text <> "" Then
        
                Range(txtColunaClassificacao.Text + Trim(txtLinhaInicial.Text)).Select
                
                linha = CInt(txtLinhaInicial.Text)
                i = 2
                
                Do While Range(txtColunaClassificacao.Text + CStr(linha)).Value <> ""
                
                   classificacao(i, 1) = Range(txtColunaClassificacao.Text + CStr(linha)).Text
                   classificacao(i, 2) = ""
                   
                   linha = linha + 1
                   i = i + 1
                   
                Loop
                
                lstClassificacao.List = classificacao
                
                WB1.Close
                
                If optClassificacaoReceita.Value = True Then
                
                    Worksheets("PC Receitas").Activate
                
                Else
                
                    Worksheets("PC Despesas").Activate
                
                End If
                
                Range("D5").Select
                linha = 5
                    
                Do While (Range("D" + CStr(linha)).Value <> "" And Range("D" + CStr(linha)).Value <> "-")
                    
                   cmbClassificacao.AddItem Range("C" + CStr(linha)).Text
                   linha = linha + 1
                       
                Loop
                
                Worksheets(mes_processamento).Activate
                
            Else
            
                MsgBox "Insira a coluna que possui, na planilha de origem, os dados de classificação de receita ou despesa.", vbInformation, "Processamento de Recebimentos"
            
            End If
        
        Else
        
            MsgBox "Insira o número da linha que inicia a carga dos dados.", vbInformation, "Processamento de Recebimentos"
        
        End If
        
    Else
    
        MsgBox "Insira caminho no qual se encontra a planilha de origem.", vbInformation, "Processamento de Recebimentos"
            
    End If
    
End Sub

Private Sub btnFechar_Click()

    Unload Me
    
End Sub

Private Sub btnImportarDados_Click()

    Dim classificacao As String
    Dim dia As String
    Dim docref As String
    Dim instfin As String
    Dim valor As String
    Dim status As String
    
    Dim linha As Integer
    Dim contador As Integer
        
    Dim processamentoImportacao(1 To 10000, 1 To 6) As String
        
    mes_processamento = ActiveSheet.Name
     
    Set WB1 = Workbooks.Open(txtCaminhoPlanilha.Text)
     
    linha = CInt(txtLinhaInicial.Text)
    contador = 1
                
    Do While Range("A" + CStr(linha)).Value <> ""
              
        processamentoImportacao(contador, 2) = Range(txtDiaOrigem.Text + CStr(linha)).Value
        processamentoImportacao(contador, 1) = Range(txtColunaClassificacao.Text + CStr(linha)).Value
        processamentoImportacao(contador, 3) = Range(txtDocRefOrigem.Text + CStr(linha)).Value
        processamentoImportacao(contador, 4) = Range(txtInstFinOrigem.Text + CStr(linha)).Value
        processamentoImportacao(contador, 5) = Range(txtValorOrigem.Text + CStr(linha)).Value
        processamentoImportacao(contador, 6) = Range(txtStatusOrigem.Text + CStr(linha)).Value
        
        linha = linha + 1
        contador = contador + 1
                   
    Loop
    
    WB1.Close
    
    Worksheets(mes_processamento).Activate
    
    contador = 1
    linha = 5
    
    Do While processamentoImportacao(contador, 1) <> ""
              
        Range(txtColunaClassificacaoDestino.Text + CStr(linha)).Value = processamentoImportacao(contador, 1)
        Range(txtDiaDestino.Text + CStr(linha)).Value = CInt(Mid(processamentoImportacao(contador, 2), 1, 2))
        Range(txtDocRefDestino.Text + CStr(linha)).Value = processamentoImportacao(contador, 3)
        Range(txtInstFinDestino.Text + CStr(linha)).Value = processamentoImportacao(contador, 4)
        Range(txtValorDestino.Text + CStr(linha)).Value = CDbl(processamentoImportacao(contador, 5))
        Range(txtStatusDestino.Text + CStr(linha)).Value = processamentoImportacao(contador, 6)
        
        linha = linha + 1
        contador = contador + 1
                   
    Loop
        
    Range("C5").Select
    frmImportarPlanilhaComParametro.Hide
    
    MsgBox "Importação realizada com sucesso!", vbInformation, "Processamento de Recebimentos"
    
End Sub




Private Sub lstClassificacao_Click()

    itemListaClassificacao = lstClassificacao.ListIndex
    txtCodigoClassificacaoOrigem = lstClassificacao.List(itemListaClassificacao, 0)
    
End Sub


Private Sub UserForm_Click()

End Sub
