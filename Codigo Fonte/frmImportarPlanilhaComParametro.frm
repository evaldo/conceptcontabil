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
Dim WB1 As Workbook

Dim itemListaClassificacao As Integer

Dim mes_processamento  As String

Dim classificacao(0 To 10000, 1 To 5) As String
Dim descricaoClassificacao(1 To 100, 1 To 3) As String

Public erroAtualizaCenario As Boolean
Public bolSalvarImportacao As Boolean
Public bolExistemDados As Boolean
Public bolLimparDados As Boolean


Private Sub btnAtualizaClassificacao_Click()

On Error GoTo Erro

    Dim receitaDespesa As String
    Dim i As Long

    'lstClassificacao.List(itemListaClassificacao, 1) = cmbClassificacao.Text
    
    mes_processamento = ActiveSheet.Name
    
    If frmEscolhaDesRec.bolClassificacaoReceita = True Then
                
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
            
            For i = 0 To lstClassificacao.ListCount - 1
            
               If lstClassificacao.Selected(i) = True Then
                
                    lstClassificacao.List(i, 1) = cmbClassificacao.Text
                    lstClassificacao.List(i, 2) = Range(descricaoClassificacao(cmbListaDescricaoClassificacao.ListIndex + 1, 2) + CStr(linha)).Value
                    lstClassificacao.List(i, 4) = receitaDespesa
                    
                    txtDescricaoClassificacao.Text = Range(descricaoClassificacao(cmbListaDescricaoClassificacao.ListIndex + 1, 2) + CStr(linha)).Value
                    
                    classificacao(i, 2) = cmbClassificacao.Text
                    classificacao(i, 3) = Range(descricaoClassificacao(cmbListaDescricaoClassificacao.ListIndex + 1, 2) + CStr(linha)).Value
                    classificacao(i, 4) = cmbListaDescricaoClassificacao.Text
                    classificacao(i, 5) = receitaDespesa
                    
               End If
               
            Next i
            
            Exit Do
            
        End If
        
        linha = linha + 1
           
    Loop
    
    Worksheets(mes_processamento).Activate
    
    txtDescricaoClassificacao.Text = ""
    
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
    Dim contador_classifcacao As Integer
    Dim linha_classificacao As Integer
            
    Dim bol_ja_existe_classificacao As Boolean
    Dim bol_encontrou_palavra As Boolean
    Dim encontrou_classificacao As Boolean
    
    Dim classificacao_importada(1 To 10000) As String
    Dim mes_processamento As String
    
    If MsgBox("Deseja refazer os parâmetros de dados para importação?", vbYesNo, "Carga de Dados para Importação") = vbNo Then
    
        mes_processamento = ActiveSheet.Name
        Worksheets("Configurações Básicas").Activate
        
        If Range("E6").Value = "Sim" Then
        
            Worksheets(mes_processamento).Activate
            Call fazLeituraDadosImportacao
            Worksheets(mes_processamento).Activate
                    
            Exit Sub
            
        End If
        
    End If
    
    mes_processamento = ActiveSheet.Name
    Worksheets(mes_processamento).Activate
    
    For i = 0 To 10000
        
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
                        
                        For i_armazenada = 1 To 10000
                            
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
                
                'ReDim Preserve classificacao(i, 5)
                lstClassificacao.List = classificacao
                
                WB1.Save
                WB1.Close
                
                cmbListaDescricaoClassificacao.Clear
                
                For linha = 1 To 100
                    
                    descricaoClassificacao(linha, 2) = ""
                    descricaoClassificacao(linha, 1) = ""
                    descricaoClassificacao(linha, 3) = ""
                    
                Next linha
                
                If frmEscolhaDesRec.bolClassificacaoReceita = True Then
                
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
                
                If frmEscolhaDesRec.bolClassificacaoDespesa = True Then
                
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


Private Sub btnFechar_Click()

    Unload Me
    
End Sub

Private Sub btnImportarDados_Click()
    
    erroAtualizaCenario = False
    bolExistemDados = False
    bolLimparDados = False
    
    If ValidaPlanilhaProcessamento() = False Then
        MsgBox "Escolha um planilha para lançamento do Fluxo de Caixa entre Jan e Dez.", vbOKOnly + vbInformation, "Importação de Dados"
        Exit Sub
    End If
    
    If MsgBox("Os dados de cenário e da planilha serão carregados. Deseja atualizar os dados?", vbYesNo, "Atualização o Cenário Carga de Dados para Importação") = vbYes Then
        
        Range("C5").Select
        linha = 5
        
        Do While Range("C" + CStr(linha)).Value <> ""
            
            linha = linha + 1
            
            If Range("C" + CStr(linha)).Value <> "" Then
                bolExistemDados = True
                Exit Do
            End If
            
        Loop
        
        If bolExistemDados = True Then
            If MsgBox("Existem dados importados ou digitados na planilha. Deseja acrescentar os dados da planilha de origem?", vbYesNo, "Atualização o Cenário Carga de Dados para Importação") = vbYes Then
                bolLimparDados = False
            Else
                bolLimparDados = True
            End If
        End If
        
        frmBarraProgressaoImportacao.Show
        
    Else
    
        Exit Sub
        
    End If
    
End Sub

Private Sub cmbClassificacao_Click()
        
On Error GoTo Erro

    mes_processamento = ActiveSheet.Name
    
    If frmEscolhaDesRec.bolClassificacaoReceita = True Then
                
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
    
    If frmEscolhaDesRec.bolClassificacaoReceita = True Then
                
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

Private Sub cmdCaminho_Click()

    Dim intResult As Integer
    Dim strPath As String
    
    'the dialog is displayed to the user
    intResult = Application.FileDialog(msoFileDialogOpen).Show
    
    'checks if user has cancled the dialog
    If intResult <> 0 Then
        'dispaly message box
        txtCaminhoPlanilha.Text = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    End If

End Sub

Private Sub cmdOkInserePalavraExistente_Click()
    
    Dim i As Long
    
    For i = 0 To lstClassificacao.ListCount - 1
        If lstClassificacao.Selected(i) = True Then
           lstPalavraExistente.AddItem lstClassificacao.List(i, 0)
        End If
    Next i
    
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

Private Sub cmdSalvarCenario_Click()
    
    bolSalvarImportacao = True
    frmBarraProgressaoImportacao.Show
    bolSalvarImportacao = False

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
    Dim linhaFinal As Integer
    Dim contador As Integer
    Dim contador_comparacao As Integer
    Dim contador_classificacao As Integer
    Dim contadorPalavra As Integer
    
    Dim encontrou_classificacao As Boolean
    Dim bol_encontrou_palavra As Boolean
    
    mes_processamento = ActiveSheet.Name
    
    lstPalavraExistente.Clear
    lstClassificacao.Clear
    
    If frmEscolhaDesRec.bolClassificacaoReceita = True Then
        Call optClassificacaoReceita_Click
        Worksheets("Cenario Receitas").Activate
    Else
        Call optClassificacaoDespesa_Click
        Worksheets("Cenario Despesas").Activate
    End If
    
    For i = 0 To 10000
        
        classificacao(i, 1) = ""
        classificacao(i, 2) = ""
        classificacao(i, 3) = ""
        classificacao(i, 4) = ""
        classificacao(i, 5) = ""
        
    Next i

    If txtCaminhoPlanilha.Text = "" Then
        MsgBox "Favor selecionar o arquivo a ser importado!", vbOKOnly, "Carga de Dados para Importação"
        Exit Sub
    End If
    
    If Range("K5").Value <> txtCaminhoPlanilha.Text Then
        If MsgBox("O arquivo de origem é diferente do cenário atual. Deseja aceitar os dados do arquivo selecionado?", vbYesNo, "Carga de Dados para Importação") = vbYes Then
            txtCaminhoPlanilha.Text = txtCaminhoPlanilha.Text
        Else
            txtCaminhoPlanilha.Text = Range("K5").Value
        End If
    End If
    
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
    
    linhaFinal = CInt(Range("M5").Value)
    
    Set WB1 = Workbooks.Open(txtCaminhoPlanilha.Text)
    
    linha = CInt(txtLinhaInicial.Text)
    contador_classificacao = contador
    contador = 1
    contador_comparacao = 1
        
    encontrou_classificacao = True
                            
    Do While contador <= linhaFinal
        
        Do While contador_comparacao <= linhaFinal
            
            If classificacao(contador_comparacao, 1) = Range(txtColunaClassificacao.Text + CStr(linha)).Text Then
                encontrou_classificacao = True
                Exit Do
            End If
            
            contador_comparacao = contador_comparacao + 1
            
        Loop
        
        contador_comparacao = 1
        contadorPalavra = 0
        
        If encontrou_classificacao = False Then
        
            bol_encontrou_palavra = False
                
            Do While contadorPalavra <= lstPalavraExistente.ListCount - 1
    
                If Range(txtColunaClassificacao.Text + CStr(linha)).Text = lstPalavraExistente.List(contadorPalavra) Then
                
                    bol_encontrou_palavra = True
                    Exit Do
                
                End If
                
                contadorPalavra = contadorPalavra + 1
                
            Loop
            
            contadorPalavra = 1
            
            If bol_encontrou_palavra = False Then
            
                classificacao(contador_classificacao, 1) = Range(txtColunaClassificacao.Text + CStr(linha)).Text
                classificacao(contador_classificacao, 2) = ""
                classificacao(contador_classificacao, 3) = ""
                classificacao(contador_classificacao, 4) = ""
                classificacao(contador_classificacao, 5) = ""
                
                contador_classificacao = contador_classificacao + 1
                
            End If
                        
        End If
        
        encontrou_classificacao = False
        
        contador = contador + 1
        linha = linha + 1
        
    Loop
    
    WB1.Save
    WB1.Close
    
    lstClassificacao.List = classificacao
    
    Worksheets(mes_processamento).Activate
    Range("D5").Select
    
End Sub

Public Sub ProcessaImportacao()

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
    
    If bolLimparDados = True Then
    
        linha = 5
        
        'Etapa de limpeza dados da planilha do mês atual
        frmBarraProgressaoImportacao.AtualizaBarra (20 / 100), "Limpando os dados da planilha " + mes_processamento
        
        Do While linha <= 10000
            
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
        
    End If
    
    Set WB1 = Workbooks.Open(txtCaminhoPlanilha.Text)
    
    'Etapa de limpeza dados da planilha do mês atual
    frmBarraProgressaoImportacao.AtualizaBarra (40 / 100), "Abrindo a planilha " + txtCaminhoPlanilha.Text
     
    linha = CInt(txtLinhaInicial.Text)
    contador = 1
                
    bol_encontrou_palavra = False
    contadorPalavra = 0
    
    'Etapa de limpeza dados da planilha do mês atual
    frmBarraProgressaoImportacao.AtualizaBarra (50 / 100), "Lendo os dados da planilha " + txtCaminhoPlanilha.Text
    
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
        
        'Etapa de limpeza dados da planilha do mês atual
        frmBarraProgressaoImportacao.AtualizaBarra (70 / 100), "Armazenando os dados em memória "
        
        If bol_encontrou_palavra = False Then
        
            linha_classificacao = 1
            
            Do While linha_classificacao <= 1000
                
                If classificacao(linha_classificacao, 1) = Range(txtColunaClassificacao.Text + CStr(linha)).Value Then
                
                    If Range(txtDiaOrigem.Text + CStr(linha)).Value = "" Then
                        processamentoImportacao(contador, 1) = "1"
                    Else
                        processamentoImportacao(contador, 1) = Mid(Range(txtDiaOrigem.Text + CStr(linha)).Value, 1, 2)
                    End If
                
                    processamentoImportacao(contador, 2) = classificacao(linha_classificacao, 4)
                    processamentoImportacao(contador, 3) = Range(txtDocRefOrigem.Text + CStr(linha)).Value
                    processamentoImportacao(contador, 4) = classificacao(linha_classificacao, 3)
                    processamentoImportacao(contador, 5) = Range(txtInstFinOrigem.Text + CStr(linha)).Value
                    
                    If Range(txtValorOrigem.Text + CStr(linha)).Value = "" Or Not IsNumeric(Range(txtValorOrigem.Text + CStr(linha)).Value) Then
                        processamentoImportacao(contador, 6) = 0
                    Else
                        processamentoImportacao(contador, 6) = Range(txtValorOrigem.Text + CStr(linha)).Value
                    End If
                    
                    processamentoImportacao(contador, 7) = classificacao(linha_classificacao, 5)
                    
                    contador = contador + 1
            
                    Exit Do
                    
                End If
                
                linha_classificacao = linha_classificacao + 1
                
            Loop
            
        End If
        
        linha = linha + 1
                   
    Loop
    
    WB1.Save
    WB1.Close
    
    Worksheets(mes_processamento).Activate
    
    contador = 1
    
    If bolLimparDados = True Then
        
        linha = 5
        
    Else
       
        Range("C5").Select
        linha = 5
        
        Do While Range("C" + CStr(linha)).Value <> ""
            linha = linha + 1
            If Range("C" + CStr(linha)).Value = "" Then Exit Do
        Loop
        
    End If
    
    'Etapa de limpeza dados da planilha do mês atual
    frmBarraProgressaoImportacao.AtualizaBarra (80 / 100), "Gravando os dados na planilha atual "
    
    Do While contador <= CInt(txtLinhaFinal.Text)
          
        If processamentoImportacao(contador, 4) <> "" Then
          
            'contador = contador - 1
          
            Range(txtDiaDestino.Text + CStr(linha)).Value = processamentoImportacao(contador, 1)
            Range(txtColunaClassificacaoDestino.Text + CStr(linha)).Value = processamentoImportacao(contador, 4)
            Range(txtDocRefDestino.Text + CStr(linha)).Value = processamentoImportacao(contador, 3)
            Range(txtColunaDescricaoClassificacaoDestino.Text + CStr(linha)).Value = processamentoImportacao(contador, 2)
            Range(txtInstFinDestino.Text + CStr(linha)).Value = processamentoImportacao(contador, 5)
            
            If processamentoImportacao(contador, 7) = "D" Then
                Range(txtValorDestinoDespesa.Text + CStr(linha)).Value = CDbl(Trim(IIf(processamentoImportacao(contador, 6) = "", 0, processamentoImportacao(contador, 6))))
                Range(txtValorDestinoReceita.Text + CStr(linha)).Value = 0
                Range("L" + CStr(linha)).Value = "Pago"
            Else
                Range(txtValorDestinoReceita.Text + CStr(linha)).Value = CDbl(Trim(IIf(processamentoImportacao(contador, 6) = "", 0, processamentoImportacao(contador, 6))))
                Range(txtValorDestinoDespesa.Text + CStr(linha)).Value = 0
                Range("L" + CStr(linha)).Value = "Realizado"
            End If
            
            linha = linha + 1
            
        End If
            
        contador = contador + 1
                   
    Loop
        
    Range("C5").Select
    frmBarraProgressaoImportacao.Hide
    
    Range("C4:N10000").Select
    ActiveWorkbook.Worksheets(mes_processamento).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(mes_processamento).Sort.SortFields.Add Key:=Range("C5:C10000"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(mes_processamento).Sort
        .SetRange Range("C4:N10000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("C5").Select
    
    MsgBox "Importação realizada com sucesso!", vbInformation, "Processamento de Recebimentos"
    
    Exit Sub
    
Erro:

    MsgBox "Foi localizado um erro no processamento de dados. Favor observar os seguintes itens: " & Chr(13) & Chr(13) & _
    "-> Verifique se o nome do arquivo está correto." & Chr(13) & _
    "-> Verifique se a coluna de origem está correta para transferir os dados." & Chr(13) & _
    "-> Verifique se a coluna de destino está correta para receber os dados.", vbOKOnly + vbInformation, "Erro ao Carregar Dados"
    
    Worksheets(mes_processamento).Activate
    
    frmBarraProgressaoImportacao.Hide
    
End Sub
Public Sub SalvarImportacao()

On Error GoTo Erro

    Dim linha As Integer
    Dim contador As Integer
    Dim contadorPalavra As Integer
    
    Dim bol_encontrou_palavra As Boolean
    Dim SalvarImportacao As Boolean
        
    'If MsgBox("Deseja salvar importação?", vbYesNo, "Salvar Importação") = vbYes Then
    SalvarImportacao = True
    'Else
    '    salvarImportacao = False
    'End If
    
    If txtCaminhoPlanilha.Text = "" Or txtLinhaInicial.Text = "" Or txtLinhaFinal.Text = "" Or txtColunaClassificacao.Text = "" _
        Or txtDiaOrigem.Text = "" Or txtDocRefOrigem.Text = "" Or txtInstFinOrigem.Text = "" Or _
        txtValorOrigem.Text = "" Then
        
        MsgBox "Os dados sobre o caminho do arquivo, valor de linha inicial, valor de linha final, coluna de origem da classificação, " & Chr(13) & _
                "coluna de origem do documento de referência e coluna de origem de valor, devem estar preenchidos.", vbInformation, "Gravação dos Dados de Importação"
        
        frmBarraProgressaoImportacao.Hide
        
        erroAtualizaCenario = True
                
        Exit Sub
        
    End If
    
    mes_processamento = ActiveSheet.Name
    
    Worksheets("Configurações Básicas").Activate
    
    Range("E6").Select
    Range("E6").Value = IIf(SalvarImportacao = True, "Sim", "Não")
    
    If frmEscolhaDesRec.bolClassificacaoReceita = True Then
        Worksheets("Cenario Receitas").Activate
    Else
        Worksheets("Cenario Despesas").Activate
    End If
    
    If SalvarImportacao = True Then
    
        Range("G5").Select
        linha = 5
        contador = 1
        
        'Etapa de limpeza das configurações da aplicação
        frmBarraProgressaoImportacao.AtualizaBarra (20 / 100), "Limpando os dados para o novo cenário..."
        
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
        
        'Etapa de descarte de palavras
        frmBarraProgressaoImportacao.AtualizaBarra (60 / 100), "Descartando palavras..."
        
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
        
        'Etapa de término da gravação do cenário
        frmBarraProgressaoImportacao.AtualizaBarra (80 / 100), "Terminando a gravação de cenário..."
        
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
        
        'Etapa de finalização
        frmBarraProgressaoImportacao.AtualizaBarra (95 / 100), "Finalização da gravação de cenário..."
        
        Do While contador <= lstPalavraExistente.ListCount - 1
            
            Range("O" + CStr(linha)).Value = lstPalavraExistente.List(contador, 0)
                        
            linha = linha + 1
            contador = contador + 1
            
        Loop
        
        Range("D5").Select
        
        Worksheets(mes_processamento).Activate
        
        If frmImportarPlanilhaComParametro.bolSalvarImportacao = True Then
            frmBarraProgressaoImportacao.Hide
        End If
        
        'MsgBox "Gravação dos dados realizada com sucesso!", vbInformation, "Importação de Dados"
        
    Else
    
        Worksheets(mes_processamento).Activate
    
    End If
    
    Exit Sub

Erro:

    MsgBox "Erro salvar os dados.", vbOKOnly + vbInformation, "Erro ao Salvar os Dados de Importação"
    
    Worksheets(mes_processamento).Activate
    
    frmImportarPlanilhaComParametro.erroAtualizaCenario = True
    
    frmBarraProgressaoImportacao.Hide
    
End Sub

Private Sub UserForm_Activate()

    If frmEscolhaDesRec.bolClassificacaoDespesa = True Then
        Call optClassificacaoDespesa_Click
        optClassificacaoReceita.Enabled = False
    Else
        Call optClassificacaoReceita_Click
        optClassificacaoDespesa.Enabled = False
    End If
        
End Sub


