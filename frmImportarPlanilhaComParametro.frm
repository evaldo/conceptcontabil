VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImportarPlanilhaComParametro 
   Caption         =   "Importação de Dados de Planilhas"
   ClientHeight    =   10200
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11175
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
Dim classificacao(0 To 1000, 1 To 5) As String
Dim descricaoClassificacao(1 To 20, 1 To 3) As String

Private Sub btnAtualizaClassificacao_Click()

On Error GoTo Erro

    Dim receitaDespesa As String

    lstClassificacao.List(itemListaClassificacao, 1) = cmbClassificacao.Text
    
    mes_processamento = ActiveSheet.Name
    
    If optClassificacaoReceita.Value = True Then
                
        Worksheets("PC Receitas").Activate
        receitaDespesa = "R"
    
    Else
    
        Worksheets("PC Despesas").Activate
        receitaDespesa = "D"
    
    End If
    
    Range("D5").Select
    linha = 5
        
    Do While Range(descricaoClassificacao(cmbListaDescricaoClassificacao.ListIndex + 1, 2) + CStr(linha)).Value <> ""
        
        If cmbClassificacao.Text = Range(descricaoClassificacao(cmbListaDescricaoClassificacao.ListIndex + 1, 3) + CStr(linha)).Text Then
        
            lstClassificacao.List(itemListaClassificacao, 2) = Range(descricaoClassificacao(cmbListaDescricaoClassificacao.ListIndex + 1, 2) + CStr(linha)).Value
            lstClassificacao.List(itemListaClassificacao, 4) = receitaDespesa
            
            txtDescricaoClassificacao = Range(descricaoClassificacao(cmbListaDescricaoClassificacao.ListIndex + 1, 2) + CStr(linha)).Value
            
            classificacao(itemListaClassificacao, 2) = cmbClassificacao.Text
            classificacao(itemListaClassificacao, 3) = Range(descricaoClassificacao(cmbListaDescricaoClassificacao.ListIndex + 1, 2) + CStr(linha)).Value
            classificacao(itemListaClassificacao, 4) = cmbListaDescricaoClassificacao.Text
            classificacao(itemListaClassificacao, 5) = receitaDespesa
            
            Exit Do
            
        End If
        
        linha = linha + 1
           
    Loop
    
    Worksheets(mes_processamento).Activate
    
    Exit Sub
    
Erro:

    MsgBox "Erro ao atualizar os dados.", vbOKOnly + vbInformation, "Erro ao Carregar Dados"

End Sub

Private Sub btnCarregaDados_Click()

On Error GoTo Erro

    Dim i As Integer, j As Integer
    Dim i_armazenada As Integer
    Dim linha As Integer
    Dim contadorPalavra As Integer
            
    Dim bol_ja_existe_classificacao As Boolean
    Dim bol_encontrou_palavra As Boolean
    
    Dim mes_processamento As String
    
    If MsgBox("Deseja refazer os parâmetros de dados para importação?", vbYesNo, "Carga de Dados para Importação") = vbNo Then
    
        mes_processamento = ActiveSheet.Name
        Worksheets("Configurações Básicas").Activate
        
        If Range("E6").Value = "Sim" Then
        
            Worksheets(mes_processamento).Activate
            Call fazLeituraDadosImportacao
                    
            Exit Sub
            
        End If
        
    End If
    
    mes_processamento = ActiveSheet.Name
    Worksheets(mes_processamento).Activate
    
    For i = 0 To 1000
        
        classificacao(i, 1) = ""
        classificacao(i, 2) = ""
        classificacao(i, 3) = ""
        classificacao(i, 4) = ""
        classificacao(i, 5) = ""
        
    Next i
            
    If txtCaminhoPlanilha.Text <> "" Then
        
        classificacao(0, 1) = "Classificação Importada"
        classificacao(0, 2) = "Classificação Utilizada"
        classificacao(0, 3) = "Descrição da Classificação"
            
        Set WB1 = Workbooks.Open(txtCaminhoPlanilha.Text)
        
        cmbClassificacao.Clear
        lstClassificacao.Clear
                
        If txtLinhaInicial.Text <> "" Then
        
            If txtColunaClassificacao.Text <> "" Then
        
                Range(txtColunaClassificacao.Text + Trim(txtLinhaInicial.Text)).Select
                
                linha = CInt(txtLinhaInicial.Text)
                i = 1
                
                bol_encontrou_palavra = False
                contadorPalavra = 0
                
                Do While (linha >= CInt(txtLinhaInicial.Text) And linha <= CInt(txtLinhaFinal.Text))
                
                    bol_encontrou_palavra = False
                
                    Do While contadorPalavra <= lstPalavraExistente.ListCount - 1
            
                        If Range(txtColunaContemPalavra + CStr(linha)).Value = lstPalavraExistente.List(contadorPalavra) Then
                        
                            bol_encontrou_palavra = True
                            Exit Do
                        
                        End If
                        
                        contadorPalavra = contadorPalavra + 1
                        
                    Loop
                    
                    contadorPalavra = 0
                    
                    If bol_encontrou_palavra = False Then
                        
                        For i_armazenada = 1 To 100
                            
                            If classificacao(i_armazenada, 1) = Range(txtColunaClassificacao.Text + CStr(linha)).Text Then
                                bol_ja_existe_classificacao = True
                            End If
                            
                        Next i_armazenada
                        
                        If bol_ja_existe_classificacao = False Then
                        
                            classificacao(i, 1) = Range(txtColunaClassificacao.Text + CStr(linha)).Text
                            classificacao(i, 2) = ""
                            classificacao(i, 3) = ""
                            classificacao(i, 4) = ""
                            classificacao(i, 5) = ""
                            
                            i = i + 1
                            
                        End If
                        
                        bol_ja_existe_classificacao = False
                        
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
    
    Exit Sub
    
Erro:

    MsgBox "Foi localizado um erro no processamento de dados. Favor observar os seguintes itens: " & Chr(13) & Chr(13) & _
    "-> Verifique se o nome do arquivo está correto." & Chr(13) & _
    "-> Verifique se a coluna de origem está correta para transferir os dados." & Chr(13) & _
    "-> Verifique se a coluna de destino está correta para receber os dados.", vbOKOnly + vbInformation, "Erro ao Carregar Dados"
    
End Sub

Private Sub btnCarregar_Click()

End Sub

Private Sub btnFechar_Click()

    Unload Me
    
End Sub

Private Sub btnImportarDados_Click()
On Error GoTo Erro

    Dim dia As String
    Dim docref As String
    Dim instfin As String
    Dim valor As String
    Dim status As String
    
    Dim linha As Integer
    Dim contador As Integer
    Dim linha_classificacao As Integer
    Dim contadorPalavra As Integer
        
    Dim processamentoImportacao(1 To 1000, 1 To 8) As String
    
    Dim bol_encontrou_palavra As Boolean
        
    mes_processamento = ActiveSheet.Name
    
    linha = 5
    
    Do While linha <= CInt(txtLinhaFinal.Text)
        
        Range(txtColunaClassificacaoDestino.Text + CStr(linha)).Value = ""
        Range(txtDiaDestino.Text + CStr(linha)).Value = ""
        Range(txtDocRefDestino.Text + CStr(linha)).Value = ""
        Range(txtInstFinDestino.Text + CStr(linha)).Value = ""
        Range(txtValorDestinoDespesa.Text + CStr(linha)).Value = CDbl(0)
        Range(txtValorDestinoReceita.Text + CStr(linha)).Value = CDbl(0)
        Range("L" + CStr(linha)).Value = ""
        Range(txtColunaDescricaoClassificacaoDestino.Text + CStr(linha)).Value = ""
        
        linha = linha + 1
        
    Loop
    
    Set WB1 = Workbooks.Open(txtCaminhoPlanilha.Text)
     
    linha = CInt(txtLinhaInicial.Text)
    contador = 1
                
    bol_encontrou_palavra = False
    contadorPalavra = 0
    
    Do While (linha >= CInt(txtLinhaInicial.Text) And linha <= CInt(txtLinhaFinal.Text))
        
        bol_encontrou_palavra = False
        
        Do While contadorPalavra <= lstPalavraExistente.ListCount - 1
            
            If Range(txtColunaContemPalavra + CStr(linha)).Value = lstPalavraExistente.List(contadorPalavra) Then
            
                bol_encontrou_palavra = True
                Exit Do
            
            End If
            
            contadorPalavra = contadorPalavra + 1
            
        Loop
        
        contadorPalavra = 0
        
        If bol_encontrou_palavra = False Then
        
            linha_classificacao = 2
            
            Do While linha_classificacao <= 1000
                
                If classificacao(linha_classificacao, 1) = Range(txtColunaClassificacao.Text + CStr(linha)).Value Then
                
                    If Range(txtDiaOrigem.Text + CStr(linha)).Value = "" Then
                        processamentoImportacao(contador, 1) = "1"
                    Else
                        processamentoImportacao(contador, 1) = Mid(Range(txtDiaOrigem.Text + CStr(linha)).Value, 1, 2)
                    End If
                
                    processamentoImportacao(contador, 2) = classificacao(linha_classificacao - 1, 4)
                    processamentoImportacao(contador, 3) = Range(txtDocRefOrigem.Text + CStr(linha)).Value
                    processamentoImportacao(contador, 4) = classificacao(linha_classificacao - 1, 3)
                    processamentoImportacao(contador, 5) = Range(txtInstFinOrigem.Text + CStr(linha)).Value
                    If Range(txtValorOrigem.Text + CStr(linha)).Value = "" Or Not IsNumeric(Range(txtValorOrigem.Text + CStr(linha)).Value) Then
                        processamentoImportacao(contador, 6) = CDbl(0)
                    Else
                        processamentoImportacao(contador, 6) = CDbl(Range(txtValorOrigem.Text + CStr(linha)).Value)
                    End If
                    
                    processamentoImportacao(contador, 7) = classificacao(linha_classificacao - 1, 5)
                    
                    contador = contador + 1
            
                    Exit Do
                    
                End If
                
                linha_classificacao = linha_classificacao + 1
                
            Loop
            
        End If
        
        linha = linha + 1
                   
    Loop
    
    WB1.Close
    
    Worksheets(mes_processamento).Activate
    
    contador = 2
    linha = 5
    
    Do While contador <= CInt(txtLinhaFinal.Text)
             
        Range(txtDiaDestino.Text + CStr(linha)).Value = processamentoImportacao(contador, 1)
        Range(txtColunaClassificacaoDestino.Text + CStr(linha)).Value = processamentoImportacao(contador, 4)
        Range(txtDocRefDestino.Text + CStr(linha)).Value = processamentoImportacao(contador, 3)
        Range(txtColunaDescricaoClassificacaoDestino.Text + CStr(linha)).Value = processamentoImportacao(contador, 2)
        Range(txtInstFinDestino.Text + CStr(linha)).Value = processamentoImportacao(contador, 5)
        
        If processamentoImportacao(contador, 7) = "D" Then
            Range(txtValorDestinoDespesa.Text + CStr(linha)).Value = processamentoImportacao(contador, 6)
            Range(txtValorDestinoReceita.Text + CStr(linha)).Value = ""
            Range("L" + CStr(linha)).Value = "Pago"
        Else
            Range(txtValorDestinoReceita.Text + CStr(linha)).Value = processamentoImportacao(contador, 6)
            Range(txtValorDestinoDespesa.Text + CStr(linha)).Value = ""
            Range("L" + CStr(linha)).Value = "Realizado"
        End If
        
        linha = linha + 1
        contador = contador + 1
                   
    Loop
        
    Range("C5").Select
    frmImportarPlanilhaComParametro.Hide
    
    MsgBox "Importação realizada com sucesso!", vbInformation, "Processamento de Recebimentos"
    
    Exit Sub
    
Erro:

    MsgBox "Foi localizado um erro no processamento de dados. Favor observar os seguintes itens: " & Chr(13) & Chr(13) & _
    "-> Verifique se o nome do arquivo está correto." & Chr(13) & _
    "-> Verifique se a coluna de origem está correta para transferir os dados." & Chr(13) & _
    "-> Verifique se a coluna de destino está correta para receber os dados.", vbOKOnly + vbInformation, "Erro ao Carregar Dados"
    
End Sub

Private Sub btnSalvarImportacao_Click()

On Error GoTo Erro

    Dim linha As Integer
    Dim contador As Integer
    Dim bol_encontrou_palavra As Boolean
    Dim contadorPalavra As Integer
    Dim salvarImportacao As Boolean
    
    If MsgBox("Deseja salvar importação?", vbYesNo, "Salvar Importação") = vbYes Then
        salvarImportacao = True
    Else
        salvarImportacao = False
    End If
    
    If txtCaminhoPlanilha.Text = "" Or txtLinhaInicial.Text = "" Or txtLinhaFinal.Text = "" Or txtColunaClassificacao.Text = "" _
        Or txtDiaOrigem.Text = "" Or txtDocRefOrigem.Text = "" Or txtInstFinOrigem.Text = "" Or _
        txtValorOrigem.Text = "" Then
        
        MsgBox "Os dados sobre o caminho do arquivo, valor de linha inicial, valor de linha final, coluna de origem da classificação, " & Chr(13) & _
                "coluna de origem do documento de referência e coluna de origem de valor, devem estar preenchidos.", vbInformation, "Gravação dos Dados de Importação"
        
        Exit Sub
        
    End If
    
    mes_processamento = ActiveSheet.Name
    
    Worksheets("Configurações Básicas").Activate
    
    Range("E6").Select
    Range("E6").Value = IIf(salvarImportacao = True, "Sim", "Não")
    
    If salvarImportacao = True Then
    
        Range("G5").Select
        linha = 5
        contador = 1
        
        Do While contador <= 1000
            
            Range("G" + CStr(linha)).Value = ""
            Range("H" + CStr(linha)).Value = ""
            Range("I" + CStr(linha)).Value = ""
            Range("J" + CStr(linha)).Value = ""
            Range("T" + CStr(linha)).Value = ""
            
            linha = linha + 1
            contador = contador + 1
            
        Loop
        
        linha = 5
        
        Do While linha <= 1000
            
            Range("O" + CStr(linha)).Value = ""
                        
            linha = linha + 1
            
        Loop
        
        Range("K5").Value = ""
        Range("L5").Value = ""
        Range("M5").Value = ""
        Range("N5").Value = ""
        Range("P5").Value = ""
        Range("Q5").Value = ""
        Range("R5").Value = ""
        Range("S5").Value = ""
        
        Range("G5").Select
        linha = 5
        contador = 1
        contadorPalavra = 0
        
        bol_encontrou_palavra = False
        
        Do While contador <= lstClassificacao.ListCount - 1
            
            If lstClassificacao.List(contador, 0) = "" Then Exit Do
                    
            bol_encontrou_palavra = False
            contadorPalavra = 0
        
            Do While contadorPalavra <= lstPalavraExistente.ListCount - 1
    
                If lstClassificacao.List(contador, 0) = lstPalavraExistente.List(contadorPalavra) Then
                
                    bol_encontrou_palavra = True
                    Exit Do
                
                End If
                
                contadorPalavra = contadorPalavra + 1
                
            Loop
            
            If bol_encontrou_palavra = False Then
                    
                Range("G" + CStr(linha)).Value = classificacao(contador, 1)
                Range("H" + CStr(linha)).Value = classificacao(contador, 4)
                Range("I" + CStr(linha)).Value = classificacao(contador, 2)
                Range("J" + CStr(linha)).Value = classificacao(contador, 3)
                Range("T" + CStr(linha)).Value = classificacao(contador, 5)
                                
                linha = linha + 1
                
            End If
            
            contador = contador + 1
            
        Loop
        
        Range("K5").Value = txtCaminhoPlanilha.Text
        Range("L5").Value = txtLinhaInicial.Text
        Range("M5").Value = txtLinhaFinal.Text
        Range("N5").Value = txtColunaClassificacao.Text
        Range("P5").Value = txtDiaOrigem.Text
        Range("Q5").Value = txtDocRefOrigem.Text
        Range("R5").Value = txtInstFinOrigem.Text
        Range("S5").Value = txtValorOrigem.Text
        
        linha = 5
        contador = 0
        
        Do While contador <= lstPalavraExistente.ListCount - 1
            
            Range("O" + CStr(linha)).Value = lstPalavraExistente.List(contador, 0)
                        
            linha = linha + 1
            contador = contador + 1
            
        Loop
        
        Range("D5").Select
        
        Worksheets(mes_processamento).Activate
        
        MsgBox "Gravação dos dados realizada com sucesso!", vbInformation, "Importação de Dados"
        
    Else
    
        Worksheets(mes_processamento).Activate
    
    End If
    
    Exit Sub

Erro:

    MsgBox "Erro salvar os dados.", vbOKOnly + vbInformation, "Erro ao Salvar os Dados de Importação"

End Sub

Private Sub cmbClassificacao_Click()
        
On Error GoTo Erro

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
                    
            txtDescricaoClassificacao = Range(descricaoClassificacao(cmbListaDescricaoClassificacao.ListIndex + 1, 2) + CStr(linha)).Value
            
            Exit Do
            
        End If
        
        linha = linha + 1
           
    Loop
    
    Worksheets(mes_processamento).Activate
    
    Exit Sub
    
Erro:

    MsgBox "Erro ao consultar a descrição do plano de contas selecionado.", vbOKOnly + vbInformation, "Erro escolher plano de contas"

    
End Sub

Private Sub cmbListaDescricaoClassificacao_Click()

    Dim linha As Integer
    
    cmbClassificacao.Clear
    txtDescricaoClassificacao.Text = ""
    
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

Private Sub cmdOkInserePalavraExistente_Click()

    lstPalavraExistente.AddItem txtPalavra.Text
    txtPalavra.Text = ""

End Sub

Private Sub cmdRetiraPalavraExistente_Click()

    Dim i As Long
    
    For i = 0 To lstPalavraExistente.ListCount - 1
        If lstPalavraExistente.Selected(i) Then
            txtPalavra.Text = lstPalavraExistente.Text
            lstPalavraExistente.RemoveItem (lstPalavraExistente.ListIndex)
        End If
    Next i

End Sub



Private Sub lstClassificacao_Click()

    itemListaClassificacao = lstClassificacao.ListIndex
    txtCodigoClassificacaoOrigem.Text = lstClassificacao.List(itemListaClassificacao, 0)
    txtPalavra.Text = lstClassificacao.List(itemListaClassificacao, 0)
    
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

Sub fazLeituraDadosImportacao()

    Dim linha As Integer
    Dim contador As Integer
    
    mes_processamento = ActiveSheet.Name
    
    lstPalavraExistente.Clear
    lstClassificacao.Clear
    
    Call optClassificacaoReceita_Click
    
    For i = 0 To 1000
        
        classificacao(i, 1) = ""
        classificacao(i, 2) = ""
        classificacao(i, 3) = ""
        classificacao(i, 4) = ""
        classificacao(i, 5) = ""
        
    Next i

    Worksheets("Configurações Básicas").Activate
    
    txtCaminhoPlanilha.Text = Range("K5").Value
    txtLinhaInicial.Text = Range("L5").Value
    txtLinhaFinal.Text = Range("M5").Value
    txtColunaClassificacao.Text = Range("N5").Value
    txtDiaOrigem.Text = Range("P5").Value
    txtDocRefOrigem.Text = Range("Q5").Value
    txtInstFinOrigem.Text = Range("R5").Value
    txtValorOrigem.Text = Range("S5").Value
    
    linha = 5
        
    Do While Range("O" + CStr(linha)).Value <> ""
            
        lstPalavraExistente.AddItem Range("O" + CStr(linha)).Value
        linha = linha + 1
            
    Loop
        
    Range("G5").Select
    contador = 1
    linha = 5
    
    classificacao(0, 1) = "Classificação Importada"
    classificacao(0, 2) = "Classificação Utilizada"
    classificacao(0, 3) = "Descrição da Classificação"
    
    Do While Range("G" + CStr(linha)).Value <> ""
    
        classificacao(contador, 1) = Range("G" + CStr(linha)).Value
        classificacao(contador, 2) = Range("I" + CStr(linha)).Value
        classificacao(contador, 3) = Range("J" + CStr(linha)).Value
        classificacao(contador, 4) = Range("H" + CStr(linha)).Value
        classificacao(contador, 5) = Range("T" + CStr(linha)).Value
    
        linha = linha + 1
        contador = contador + 1
    
    Loop
    
    lstClassificacao.List = classificacao
    
    Range("D5").Select
    Worksheets(mes_processamento).Activate
    
End Sub


