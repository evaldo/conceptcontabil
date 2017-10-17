VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImportarPlanilhaComParametro 
   Caption         =   "Importar Dados de Planilhas"
   ClientHeight    =   9870
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
Dim classificacao(1 To 1000, 1 To 4) As String
Dim descricaoClassificacao(1 To 20, 1 To 3) As String

Private Sub btnAtualizaClassificacao_Click()

    lstClassificacao.List(itemListaClassificacao, 1) = cmbClassificacao.Text
    
    mes_processamento = ActiveSheet.Name
    
    If optClassificacaoReceita.Value = True Then
                
        Worksheets("PC Receitas").Activate
    
    Else
    
        Worksheets("PC Despesas").Activate
    
    End If
    
    Range("D5").Select
    linha = 5
        
    Do While Range(descricaoClassificacao(cmbListaDescricaoClassificacao.ListIndex + 1, 2) + CStr(linha)).Value <> ""
        
        If cmbClassificacao.Text = Range(descricaoClassificacao(cmbListaDescricaoClassificacao.ListIndex + 1, 3) + CStr(linha)).Text Then
            lstClassificacao.List(itemListaClassificacao, 2) = Range(descricaoClassificacao(cmbListaDescricaoClassificacao.ListIndex + 1, 2) + CStr(linha)).Value
            classificacao(itemListaClassificacao, 3) = Range(descricaoClassificacao(cmbListaDescricaoClassificacao.ListIndex + 1, 2) + CStr(linha)).Value
            classificacao(itemListaClassificacao, 4) = cmbListaDescricaoClassificacao.Text
        End If
        
        linha = linha + 1
           
    Loop
    
    Worksheets(mes_processamento).Activate

End Sub

Private Sub btnCarregaDados_Click()

'On Error Resume Next

    Dim i As Integer, j As Integer
    Dim i_armazenada As Integer
    Dim linha As Integer
        
    Dim bol_ja_existe_classificacao As Boolean
            
    If txtCaminhoPlanilha.Text <> "" Then
        
        classificacao(1, 1) = "Classificação Importada"
        classificacao(1, 2) = "Classificação Utilizada"
        classificacao(1, 3) = "Descrição da Classificação"
            
        Set WB1 = Workbooks.Open(txtCaminhoPlanilha.Text)
        
        cmbClassificacao.Clear
        lstClassificacao.Clear
                
        If txtLinhaInicial.Text <> "" Then
        
            If txtColunaClassificacao.Text <> "" Then
        
                Range(txtColunaClassificacao.Text + Trim(txtLinhaInicial.Text)).Select
                
                linha = CInt(txtLinhaInicial.Text)
                i = 2
                
                Do While Range(txtColunaClassificacao.Text + CStr(linha)).Value <> ""
                    
                    bol_ja_existe_classificacao = False
                    
                    For i_armazenada = 2 To i
                        
                        If classificacao(i_armazenada, 1) = Range(txtColunaClassificacao.Text + CStr(linha)).Text Then
                            bol_ja_existe_classificacao = True
                        End If
                        
                    Next i_armazenada
                    
                    If bol_ja_existe_classificacao = False Then
                    
                        classificacao(i, 1) = Range(txtColunaClassificacao.Text + CStr(linha)).Text
                        classificacao(i, 2) = ""
                        classificacao(i, 3) = ""
                        
                        i = i + 1
                        
                    End If
                    
                    linha = linha + 1
                    
                Loop
                
                lstClassificacao.List = classificacao
                
                WB1.Close
                
                cmbListaDescricaoClassificacao.Clear
                
                For linha = 1 To 20
                    
                    descricaoClassificacao(linha, 2) = ""
                    descricaoClassificacao(linha, 1) = ""
                    descricaoClassificacao(linha, 3) = ""
                    
                Next linha
                
                If optClassificacaoReceita.Value = True Then
                
                    descricaoClassificacao(1, 2) = "D"
                    descricaoClassificacao(1, 1) = "RECEITAS COM PRODUTO"
                    descricaoClassificacao(1, 3) = "C"
                    
                    descricaoClassificacao(2, 2) = "E"
                    descricaoClassificacao(2, 1) = "RECEBIMENTOS REALIZADOS"
                    descricaoClassificacao(2, 3) = "C"
                    
                    descricaoClassificacao(3, 2) = "H"
                    descricaoClassificacao(3, 1) = "RECEITAS COM SERVIÇOS"
                    descricaoClassificacao(3, 3) = "G"
                    
                    descricaoClassificacao(4, 2) = "K"
                    descricaoClassificacao(4, 1) = "RECEITAS NÃO OPERACIONAIS"
                    descricaoClassificacao(4, 3) = "J"
                
                End If
                
                If optClassificacaoDespesa.Value = True Then
                
                    descricaoClassificacao(1, 2) = "D"
                    descricaoClassificacao(1, 1) = "DESPESAS COM PRODUTOS"
                    descricaoClassificacao(1, 3) = "C"
                    
                    descricaoClassificacao(2, 2) = "G"
                    descricaoClassificacao(2, 1) = "DESPESAS COM SERVIÇOS"
                    descricaoClassificacao(2, 3) = "F"
                    
                    descricaoClassificacao(3, 2) = "J"
                    descricaoClassificacao(3, 1) = "DESPESAS NÃO OPERACIONAIS"
                    descricaoClassificacao(3, 3) = "I"
                    
                    descricaoClassificacao(4, 2) = "M"
                    descricaoClassificacao(4, 1) = "DESPESAS COM RH"
                    descricaoClassificacao(4, 3) = "L"
                    
                    descricaoClassificacao(5, 2) = "P"
                    descricaoClassificacao(5, 1) = "DESPESAS OPERACIONAIS"
                    descricaoClassificacao(5, 3) = "O"
                    
                    descricaoClassificacao(6, 2) = "S"
                    descricaoClassificacao(6, 1) = "DESPESAS DE MARKETING"
                    descricaoClassificacao(6, 3) = "R"
                    
                    descricaoClassificacao(7, 2) = "V"
                    descricaoClassificacao(7, 1) = "IMPOSTOS"
                    descricaoClassificacao(7, 3) = "U"
                    
                    descricaoClassificacao(8, 2) = "Y"
                    descricaoClassificacao(8, 1) = "INVESTIMENTOS"
                    descricaoClassificacao(8, 3) = "X"
                
                End If
                
                
                cmbListaDescricaoClassificacao.List = descricaoClassificacao
                
                
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

    'Dim classificacao As String
    Dim dia As String
    Dim docref As String
    Dim instfin As String
    Dim valor As String
    Dim status As String
    
    Dim linha As Integer
    Dim contador As Integer
    Dim linha_classificacao As Integer
        
    Dim processamentoImportacao(1 To 10000, 1 To 7) As String
        
    mes_processamento = ActiveSheet.Name
     
    Set WB1 = Workbooks.Open(txtCaminhoPlanilha.Text)
     
    linha = CInt(txtLinhaInicial.Text)
    contador = 1
                
    Do While Range("A" + CStr(linha)).Value <> ""
              
        processamentoImportacao(contador, 2) = Range(txtDiaOrigem.Text + CStr(linha)).Value
        
        linha_classificacao = 1
        
        Do While linha_classificacao <= 10000
            
            If classificacao(linha_classificacao, 1) = Range(txtColunaClassificacao.Text + CStr(linha)).Value Then
                processamentoImportacao(contador, 1) = classificacao(linha_classificacao - 1, 3)
                processamentoImportacao(contador, 7) = classificacao(linha_classificacao - 1, 4)
                
                Exit Do
                
            End If
            
            linha_classificacao = linha_classificacao + 1
            
        Loop
        
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
    
    Do While processamentoImportacao(contador, 2) <> ""
              
        Range(txtColunaClassificacaoDestino.Text + CStr(linha)).Value = processamentoImportacao(contador, 1)
        Range(txtDiaDestino.Text + CStr(linha)).Value = CInt(Mid(processamentoImportacao(contador, 2), 1, 2))
        Range(txtDocRefDestino.Text + CStr(linha)).Value = processamentoImportacao(contador, 3)
        Range(txtInstFinDestino.Text + CStr(linha)).Value = processamentoImportacao(contador, 4)
        If processamentoImportacao(contador, 5) = "" Then
            Range(txtValorDestino.Text + CStr(linha)).Value = 0
        Else
            Range(txtValorDestino.Text + CStr(linha)).Value = CDbl(processamentoImportacao(contador, 5))
        End If
        Range(txtStatusDestino.Text + CStr(linha)).Value = processamentoImportacao(contador, 6)
        Range(txtColunaDescricaoClassificacaoDestino.Text + CStr(linha)).Value = processamentoImportacao(contador, 7)
        
        linha = linha + 1
        contador = contador + 1
                   
    Loop
        
    Range("C5").Select
    frmImportarPlanilhaComParametro.Hide
    
    MsgBox "Importação realizada com sucesso!", vbInformation, "Processamento de Recebimentos"
    
End Sub


Private Sub cmbListaDescricaoClassificacao_Click()

    Dim linha As Integer
    
    cmbClassificacao.Clear
    
    mes_processamento = ActiveSheet.Name
    linha = 5
    
    If optClassificacaoReceita.Value = True Then
                
        Worksheets("PC Receitas").Activate
    
    Else
    
        Worksheets("PC Despesas").Activate
    
    End If
    
    Range("D5").Select
    
    Do While (Range(descricaoClassificacao(cmbListaDescricaoClassificacao.ListIndex + 1, 2) + CStr(linha)).Value <> "" And Range(descricaoClassificacao(cmbListaDescricaoClassificacao.ListIndex + 1, 2) + CStr(linha)).Value <> "-")
                    
        cmbClassificacao.AddItem Range(descricaoClassificacao(cmbListaDescricaoClassificacao.ListIndex + 1, 3) + CStr(linha)).Text
        linha = linha + 1
                       
    Loop
    
    Worksheets(mes_processamento).Activate

End Sub

Private Sub lstClassificacao_Click()

    itemListaClassificacao = lstClassificacao.ListIndex
    txtCodigoClassificacaoOrigem = lstClassificacao.List(itemListaClassificacao, 0)
    
End Sub

Function ConverteParaLetra(iCol As Integer) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   
   If iAlpha > 0 Then
      ConverteParaLetra = Chr(iAlpha + 64)
   End If
   
   If iRemainder > 0 Then
      ConverteParaLetra = ConverteParaLetra & Chr(iRemainder + 64)
   End If
   
End Function

Private Sub optClassificacaoDespesa_Click()


    For linha = 1 To 20
                    
        descricaoClassificacao(linha, 2) = ""
        descricaoClassificacao(linha, 1) = ""
        descricaoClassificacao(linha, 3) = ""
        
    Next linha

    descricaoClassificacao(1, 2) = "D"
    descricaoClassificacao(1, 1) = "DESPESAS COM PRODUTOS"
    descricaoClassificacao(1, 3) = "C"
    
    descricaoClassificacao(2, 2) = "G"
    descricaoClassificacao(2, 1) = "DESPESAS COM SERVIÇOS"
    descricaoClassificacao(2, 3) = "F"
    
    descricaoClassificacao(3, 2) = "J"
    descricaoClassificacao(3, 1) = "DESPESAS NÃO OPERACIONAIS"
    descricaoClassificacao(3, 3) = "I"
    
    descricaoClassificacao(4, 2) = "M"
    descricaoClassificacao(4, 1) = "DESPESAS COM RH"
    descricaoClassificacao(4, 3) = "L"
    
    descricaoClassificacao(5, 2) = "P"
    descricaoClassificacao(5, 1) = "DESPESAS OPERACIONAIS"
    descricaoClassificacao(5, 3) = "O"
    
    descricaoClassificacao(6, 2) = "S"
    descricaoClassificacao(6, 1) = "DESPESAS DE MARKETING"
    descricaoClassificacao(6, 3) = "R"
    
    descricaoClassificacao(7, 2) = "V"
    descricaoClassificacao(7, 1) = "IMPOSTOS"
    descricaoClassificacao(7, 3) = "U"
    
    descricaoClassificacao(8, 2) = "Y"
    descricaoClassificacao(8, 1) = "INVESTIMENTOS"
    descricaoClassificacao(8, 3) = "X"
    
    cmbListaDescricaoClassificacao.Clear
    cmbListaDescricaoClassificacao.List = descricaoClassificacao

End Sub

Private Sub optClassificacaoReceita_Click()
 
    For linha = 1 To 20
                    
        descricaoClassificacao(linha, 2) = ""
        descricaoClassificacao(linha, 1) = ""
        descricaoClassificacao(linha, 3) = ""
        
    Next linha
    
    descricaoClassificacao(1, 2) = "D"
    descricaoClassificacao(1, 1) = "RECEITAS COM PRODUTO"
    descricaoClassificacao(1, 3) = "C"
    
    descricaoClassificacao(2, 2) = "E"
    descricaoClassificacao(2, 1) = "RECEBIMENTOS REALIZADOS"
    descricaoClassificacao(2, 3) = "C"
    
    descricaoClassificacao(3, 2) = "H"
    descricaoClassificacao(3, 1) = "RECEITAS COM SERVIÇOS"
    descricaoClassificacao(3, 3) = "G"
    
    descricaoClassificacao(4, 2) = "K"
    descricaoClassificacao(4, 1) = "RECEITAS NÃO OPERACIONAIS"
    descricaoClassificacao(4, 3) = "J"
    
    cmbListaDescricaoClassificacao.Clear
    cmbListaDescricaoClassificacao.List = descricaoClassificacao
   
End Sub
