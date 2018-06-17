Attribute VB_Name = "ProcessarRecebimento"
    
Sub processa_recebimento_caixa()

On Error GoTo Erro

    Dim plano_conta As String
    Dim mes(1 To 12) As String
    Dim mes_processamento As String
    Dim doc_ref As String
    Dim instituicao_finaceira As String
    Dim plano_contas As String
    Dim classificacao As String
    Dim mes_baixa As String
    Dim classificacaoCaixa(1 To 2000, 1 To 4) As String
    Dim indicadorRecebimento As String
    Dim colunaClassificacao As String
    
    Dim valor_recebimento As Double
            
    Dim linha_planilha As Integer
    Dim linha_planilha_Receita As Integer
    Dim conta_mes As Integer
    Dim flag_mes_processamento As Integer
    Dim linha_planilha_mes_processmento As Integer
    Dim contador As Integer
    Dim linha_log_processamento As Integer
    
    Dim bol_processar_classificacao As Boolean
    Dim bol_processar_recebimento_planilha As Boolean
        
    Dim percentual As Single
                        
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
    
    Application.ScreenUpdating = False
    
    If ValidaPlanilhaProcessamento() = False Then
        MsgBox "Escolha um planilha para lançamento do Fluxo de Caixa entre Jan e Dez.", vbOKOnly + vbInformation, "Processamento dos Recebimentos"
        frmBarraProgressaoRecebimento.Hide
        Exit Sub
    End If
    
    mes_processamento = ActiveSheet.Name
    linha_mes_processamento = ActiveCell.Row - 1
    
    linha_planilha = 12
    linha_planilha_Receita = 6
    contador = 1
    
    '-----------------------------------------------------------------------------------------------------
    'Configurações dos planos de contas e suas posições na planilha
    '-----------------------------------------------------------------------------------------------------
    
    Worksheets("Configurações Básicas").Activate
    
    Do While (Range("E" + CStr(linha_planilha)).Value <> "" And Range("E" + CStr(linha_planilha)).Value <> "-")
     
        If Range("F" + CStr(linha_planilha)).Value = "R" Then
            
            colunaClassificacao = Range("H" + CStr(linha_planilha)).Value
            colunaReceitas = Range("H" + CStr(linha_planilha)).Value
            colunaReceitas = Chr((Range(colunaReceitas + CStr(linha_planilha_Receita)).Column + 1) + 64)
            classificacao = Range("E" + CStr(linha_planilha)).Value
            
            Sheets("PC Receitas").Select
            linha_planilha_Receita = 6
            
            Do While Range(colunaClassificacao + CStr(linha_planilha_Receita)).Value <> "-" And _
                     Range(colunaClassificacao + CStr(linha_planilha_Receita)).Value <> ""
                
                indicadorRecebimento = Range(colunaReceitas + CStr(linha_planilha_Receita)).Value
                
                If Not IsEmpty(indicadorRecebimento) And indicadorRecebimento <> "-" Then
                
                    classificacaoCaixa(contador, 1) = Range(colunaClassificacao + CStr(linha_planilha_Receita)).Value
                    classificacaoCaixa(contador, 2) = colunaReceitas
                    classificacaoCaixa(contador, 3) = "S"
                    classificacaoCaixa(contador, 4) = classificacao
                    
                    contador = contador + 1
                    
                End If
                
                linha_planilha_Receita = linha_planilha_Receita + 1
                
            Loop
            
        End If
        
        Worksheets("Configurações Básicas").Activate
        
        linha_planilha = linha_planilha + 1
    
    Loop
        
    Worksheets(mes_processamento).Activate
    
    linha_planilha = 5
    linha_planilha_mes_processmento = 5
    
     For conta_mes = 1 To 12
        If mes(conta_mes) = mes_processamento Then
            flag_mes_processamento = conta_mes
            Exit For
        End If
    Next conta_mes
    
    '-----------------------------------------------------------------------------------------------------
    'Processar recebimento dos meses anteriores, exceto atual
    '-----------------------------------------------------------------------------------------------------
    
    Do While Range("E" + CStr(linha_planilha_mes_processmento)).Value <> ""
    
        frmBarraProgressaoRecebimento.AtualizaBarra percentual, "Processando Recebimento dos meses"
    
        'doc_ref = Range("F" + CStr(linha_planilha_mes_processmento)).Value
        instituicao_finaceira = Range("H" + CStr(linha_planilha_mes_processmento)).Value
        classificacao = Range("E" + CStr(linha_planilha_mes_processmento)).Value
        plano_contas = Range("G" + CStr(linha_planilha_mes_processmento)).Value
        valor_recebimento = Range("J" + CStr(linha_planilha_mes_processmento)).Value
        mes_baixa = Range("I" + CStr(linha_planilha_mes_processmento)).Value
        
        bol_processar_classificacao = False
        contador = 1
        Do While contador <= 2000
            If classificacao = classificacaoCaixa(contador, 4) And plano_contas = classificacaoCaixa(contador, 1) Then
                If classificacaoCaixa(contador, 3) = "S" Then
                    bol_processar_classificacao = True
                    Exit Do
                End If
            End If
            contador = contador + 1
        Loop
        
        If bol_processar_classificacao = True Then
            For conta_mes = 1 To flag_mes_processamento - 1
                If mes(conta_mes) = mes_baixa Then
                        
                    Sheets(mes(conta_mes)).Select
                                   
                    Do While Range("E" + CStr(linha_planilha)).Value <> ""
                    
                        percentual = linha_planilha / 1000
                        
                        'And Range("F" + CStr(linha_planilha)).Value = doc_ref
                        'And Range("M" + CStr(linha_planilha_mes_processmento)).Value = ""

                        If Range("E" + CStr(linha_planilha)).Value = classificacao _
                           And Range("H" + CStr(linha_planilha)).Value = instituicao_finaceira _
                           And Range("G" + CStr(linha_planilha)).Value = plano_contas _
                           And Range("L" + CStr(linha_planilha)).Value = "Não Pago" _
                           And Range("I" + CStr(linha_planilha)).Value = "" _
                        Then
                                
                                Range("J" + CStr(linha_planilha)).Value = Range("J" + CStr(linha_planilha)).Value - valor_recebimento
                                If Range("J" + CStr(linha_planilha)).Value <= 0 Then
                                    Range("L" + CStr(linha_planilha)).Value = "Realizado"
                                Else
                                    Range("L" + CStr(linha_planilha)).Value = "Não Pago"
                                End If
                                Range("M" + CStr(linha_planilha)).Value = "Sim"
                                
                                Worksheets(mes_processamento).Activate
                                Range("I" + CStr(linha_planilha_mes_processmento)).Value = ""
                                Range("M" + CStr(linha_planilha_mes_processmento)).Value = mes(conta_mes)
                                Sheets(mes(conta_mes)).Select
                                
                                If Range("J" + CStr(linha_planilha)).Value < 0 Then
                                
                                    valor_recebimento = Range("J" + CStr(linha_planilha)).Value
                                    
                                    Sheets("Log de Proc Recebimentos").Select
                                    
                                    linha_log_processamento = 5
                                    
                                    Do While linha_log_processamento <= 1000
                                    
                                        If Range("D" + CStr(linha_log_processamento)).Value = "" Then
                                        
                                            Range("D" + CStr(linha_log_processamento)).Value = mes_baixa
                                            Range("E" + CStr(linha_log_processamento)).Value = plano_contas
                                            Range("F" + CStr(linha_log_processamento)).Value = mes_baixa
                                            Range("G" + CStr(linha_log_processamento)).Value = valor_recebimento
                                            Range("H" + CStr(linha_log_processamento)).Value = Date
                                            Range("I" + CStr(linha_log_processamento)).Value = Time
                                            Range("J" + CStr(linha_log_processamento)).Value = "Processamento realizado no mês " + mes_baixa + ". Com valor negativo: " + CStr(valor_recebimento)
                                            
                                            Exit Do
                                            
                                        End If
                                                                            
                                        linha_log_processamento = linha_log_processamento + 1
                                    
                                    Loop
                                    
                                End If
                                
                        End If
                    
                        linha_planilha = linha_planilha + 1
                        
                    Loop
                    
                End If
                
                linha_planilha = 5
                Range("E" + CStr(linha_planilha)).Select
                
            Next conta_mes
            
        End If
        
        Sheets(mes(flag_mes_processamento)).Select
        linha_planilha_mes_processmento = linha_planilha_mes_processmento + 1
        
       Loop
       '-----------------------------------------------------------------------------------------------------
       'Processar recebimento do mês atual
       'O código é o mesmo
       '-----------------------------------------------------------------------------------------------------
        
       valor_recebimento = 0
        
       linha_planilha = 5
       linha_planilha_mes_processmento = 5
        
       bol_processar_recebimento_planilha = False
        
       Do While Range("E" + CStr(linha_planilha)).Value <> ""
           linha_planilha = linha_planilha + 1
       Loop
        
       linha_planilha = linha_planilha - 1
        
       For conta_mes = 1 To 12
           If mes(conta_mes) = mes_processamento Then
               flag_mes_processamento = conta_mes
               Exit For
           End If
       Next conta_mes
        
       Sheets(mes(flag_mes_processamento)).Select
        
       Do While Range("E" + CStr(linha_planilha)).Value <> ""
                
           percentual = linha_planilha / 1000
            
           frmBarraProgressaoRecebimento.AtualizaBarra percentual, "Processando Recebimento do mês atual"
    
           If Range("I" + CStr(linha_planilha)).Value = "" And Range("M" + CStr(linha_planilha)).Value = "" Then
    
               'doc_ref = Range("F" + CStr(linha_planilha_mes_processmento)).Value
               instituicao_finaceira = Range("H" + CStr(linha_planilha)).Value
               classificacao = Range("E" + CStr(linha_planilha)).Value
               plano_contas = Range("G" + CStr(linha_planilha)).Value
               If valor_recebimento = 0 Then valor_recebimento = Range("J" + CStr(linha_planilha)).Value
               mes_baixa = Range("I" + CStr(linha_planilha)).Value
                
               bol_processar_classificacao = False
               contador = 1
               Do While contador <= 2000
                    If classificacao = classificacaoCaixa(contador, 4) And plano_contas = classificacaoCaixa(contador, 1) Then
                        If classificacaoCaixa(contador, 3) = "S" Then
                            bol_processar_classificacao = True
                            Exit Do
                        End If
                    End If
                    contador = contador + 1
                Loop
                
                linha_planilha_mes_processmento = linha_planilha
                
                Do While Range("E" + CStr(linha_planilha_mes_processmento)).Value <> ""
                
                    'And Range("F" + CStr(linha_planilha)).Value = doc_ref
                    
                    If Range("E" + CStr(linha_planilha_mes_processmento)).Value = classificacao _
                       And Range("H" + CStr(linha_planilha_mes_processmento)).Value = instituicao_finaceira _
                       And Range("G" + CStr(linha_planilha_mes_processmento)).Value = plano_contas _
                       And Range("L" + CStr(linha_planilha_mes_processmento)).Value = "Pago" _
                       And Range("M" + CStr(linha_planilha_mes_processmento)).Value = "" _
                       And bol_processar_classificacao = True _
                    Then
                        valor_recebimento = Range("J" + CStr(linha_planilha_mes_processmento)).Value
                        linha_planilha_Receita = linha_planilha_mes_processmento
                    End If
                    
                    'And Range("M" + CStr(linha_planilha_mes_processmento)).Value = ""
                    If Range("E" + CStr(linha_planilha_mes_processmento)).Value = classificacao _
                       And Range("H" + CStr(linha_planilha_mes_processmento)).Value = instituicao_finaceira _
                       And Range("G" + CStr(linha_planilha_mes_processmento)).Value = plano_contas _
                       And Range("I" + CStr(linha_planilha_mes_processmento)).Value = mes(flag_mes_processamento) _
                       And Range("L" + CStr(linha_planilha_mes_processmento)).Value = "Não Pago" _
                       And bol_processar_classificacao = True _
                       Then
                            
                        If valor_recebimento = 0 Then
                            valor_recebimento = Range("J" + CStr(linha_planilha_mes_processmento)).Value
                        Else
                            If valor_recebimento >= CDbl(Range("J" + CStr(linha_planilha_mes_processmento)).Value) Then
                                valor_recebimento = valor_recebimento - CDbl(Range("J" + CStr(linha_planilha_mes_processmento)).Value)
                                'Range("J" + CStr(linha_planilha_mes_processmento)).Value = 0
                                Range("I" + CStr(linha_planilha_mes_processmento)).Value = ""
                                Range("L" + CStr(linha_planilha_mes_processmento)).Value = "Realizado"
                                Range("M" + CStr(linha_planilha_mes_processmento)).Value = mes(conta_mes)
                            Else
                                Range("J" + CStr(linha_planilha_mes_processmento)).Value = CDbl(Range("J" + CStr(linha_planilha_mes_processmento)).Value) - valor_recebimento
                                Range("M" + CStr(linha_planilha_Receita)).Value = mes(conta_mes)
                                If Range("J" + CStr(linha_planilha_mes_processmento)).Value = 0 Then
                                    Range("I" + CStr(linha_planilha_mes_processmento)).Value = ""
                                    Range("L" + CStr(linha_planilha_mes_processmento)).Value = "Realizado"
                                    Range("M" + CStr(linha_planilha_mes_processmento)).Value = mes(conta_mes)
                                    Range("M" + CStr(linha_planilha_Receita)).Value = mes(conta_mes)
                                End If
                                bol_processar_recebimento_planilha = True
                                Exit Do
                            End If
                        End If
                        
                        Sheets("Log de Proc Recebimentos").Select
                            
                        linha_log_processamento = 5
                        
                        Do While linha_log_processamento <= 1000
                        
                            If Range("D" + CStr(linha_log_processamento)).Value = "" Then
                            
                                Range("D" + CStr(linha_log_processamento)).Value = mes(flag_mes_processamento)
                                Range("E" + CStr(linha_log_processamento)).Value = plano_contas
                                Range("F" + CStr(linha_log_processamento)).Value = mes(flag_mes_processamento)
                                Range("G" + CStr(linha_log_processamento)).Value = valor_recebimento
                                Range("H" + CStr(linha_log_processamento)).Value = Date
                                Range("I" + CStr(linha_log_processamento)).Value = Time
                                Range("J" + CStr(linha_log_processamento)).Value = "Processamento realizado no mês " + mes_baixa + ". Com valor negativo: " + CStr(Format(valor_recebimento, "Currency"))
                                
                                Exit Do
                                
                            End If
                                                                
                            linha_log_processamento = linha_log_processamento + 1
                        
                        Loop
                        
                        Sheets(mes(flag_mes_processamento)).Select
                            
                    End If
                
                    linha_planilha_mes_processmento = linha_planilha_mes_processmento - 1
                
                Loop
                
            End If
            
            linha_planilha = linha_planilha - 1
            
            If linha_planilha = 4 Then Exit Do
            If bol_processar_recebimento_planilha = True Then Exit Do
        Loop
        
    
    frmBarraProgressaoRecebimento.Hide
    
    MsgBox "Processamento Realizado com sucesso.", vbInformation, "Processamento de Recebimentos"
    
    Application.ScreenUpdating = True
        
    Exit Sub
    
Erro:

    MsgBox "Erro ao processar o recebimento.", vbOKOnly + vbInformation, "Erro ao Carregar Dados"
    Worksheets(mes_processamento).Activate
    
End Sub

Sub processar_recebimento_com_barra()
    
    resposta = MsgBox("Deseja realmente processar recebimentos?", vbYesNo + vbExclamation, "Processamento de Recebimentos")
 
    If resposta = vbYes Then frmBarraProgressaoRecebimento.Show
    
End Sub






