Attribute VB_Name = "ProcessarRecebimento"
Sub ExportarCSV()
    
On Error Resume Next
    
    Dim NomeDoArquivo As String
    Dim WB1 As Workbook
    Dim WB2 As Workbook
    Dim rng As Range
     
    Set WB1 = ActiveWorkbook
    Set rng = Application.InputBox("Selecione o intervalo de células para a exportação da planilha atual:", "Processamento de Recebimentos", Default:="Por exemplo -> C5:L35", Type:=8)
    
    If rng <> "" Then
    
       Application.ScreenUpdating = False
       rng.Copy
    
       Set WB2 = Application.Workbooks.Add(1)
       WB2.Sheets(1).Range("A1").PasteSpecial xlPasteValues
        
       NomeDoArquivo = "CSV_Export_" & Format(Date, "ddmmyyyy")
       FullPath = WB1.Path & "\" & NomeDoArquivo
        
       Application.DisplayAlerts = False
       
       If MsgBox("Dados copiados para " & WB1.Path & "\" & NomeDoArquivo & vbCrLf & _
       "Atenção: Arquivos no diretório com mesmo nome serão sobrescritos!!", vbQuestion + vbYesNo) <> vbYes Then
           Exit Sub
       End If
        
       If Not Right(NomeDoArquivo, 4) = ".csv" Then MyFileName = NomeDoArquivo & ".csv"
       
       With WB2
           .SaveAs Filename:=FullPath, FileFormat:=xlCSV, CreateBackup:=False
           .Close False
       End With
       
       Application.DisplayAlerts = True
       
    End If
End Sub

Sub processa_recebimento_caixa()

On Error GoTo Erro

    Dim plano_conta As String
    Dim mes(1 To 12) As String
    Dim celula_planilha_atual As String
    Dim celula_planilha_lida As String
    Dim mes_processamento As String
    Dim doc_ref As String
    Dim instituicao_finaceira As String
    Dim plano_contas As String
    Dim mes_processamento_anterior As String
    Dim plano_contas_anterior As String
    Dim mes_baixa As String
    
    Dim valor_recebimento As Double
            
    Dim linha_planilha As Integer
    Dim conta_mes As Integer
    Dim flag_mes_processamento As Integer
    Dim linha_planilha_mes_processmento As Integer
    Dim contador As Integer
    Dim linha_log_processamento As Integer
                
    Dim bol_processar_classificacao As Boolean
        
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
        
    linha_planilha = 5
    linha_planilha_mes_processmento = 5
    
    mes_processamento = ActiveSheet.Name
    
    For conta_mes = 1 To 12
        If mes(conta_mes) = mes_processamento Then
            flag_mes_processamento = conta_mes
            Exit For
        End If
    Next conta_mes
    
    Do While Range("E" + CStr(linha_planilha_mes_processmento)).Value <> ""
    
        doc_ref = Range("F" + CStr(linha_planilha_mes_processmento)).Value
        instituicao_finaceira = Range("H" + CStr(linha_planilha_mes_processmento)).Value
        classificacao = Range("E" + CStr(linha_planilha_mes_processmento)).Value
        plano_contas = Range("G" + CStr(linha_planilha_mes_processmento)).Value
        valor_recebimento = Range("J" + CStr(linha_planilha_mes_processmento)).Value
        mes_baixa = Range("I" + CStr(linha_planilha_mes_processmento)).Value
        
        bol_processar_classificacao = False
                
        If classificacao = "RECEITAS COM PRODUTO" Then
        
           Sheets("PC Receitas").Select
           linha_planilha_recebimento = 5
           
           Do While Range("D" + CStr(linha_planilha_recebimento)).Value <> ""
           
                If Range("D" + CStr(linha_planilha_recebimento)).Value = plano_contas And _
                  (Not IsEmpty(Range("E" + CStr(linha_planilha_recebimento)).Value) _
                  Or Range("E" + CStr(linha_planilha_recebimento)).Value = "-") Then
                    bol_processar_classificacao = True
                    Exit Do
                End If
                 
                linha_planilha_recebimento = linha_planilha_recebimento + 1
            
           Loop
           
        End If
        
        If bol_processar_classificacao = True Then
        
            For conta_mes = 1 To flag_mes_processamento - 1
               
                If mes(conta_mes) = mes_baixa Then
                        
                    Sheets(mes(conta_mes)).Select
                    
                    percentual = conta_mes / flag_mes_processamento
                    frmBarraProgressaoRecebimento.AtualizaBarra percentual, mes(conta_mes)
                                   
                    Do While Range("E" + CStr(linha_planilha)).Value <> ""
                        
                        If Range("E" + CStr(linha_planilha)).Value = classificacao _
                           And Range("H" + CStr(linha_planilha)).Value = instituicao_finaceira _
                           And Range("F" + CStr(linha_planilha)).Value = doc_ref _
                           And Range("G" + CStr(linha_planilha)).Value = plano_contas _
                           Then
                                
                                Range("J" + CStr(linha_planilha)).Value = Range("J" + CStr(linha_planilha)).Value - valor_recebimento
                                Range("L" + CStr(linha_planilha)).Value = "Não Pago"
                                Range("M" + CStr(linha_planilha)).Value = "Sim"
                                
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
    
    frmBarraProgressaoRecebimento.Hide
    
    MsgBox "Processamento Realizado com sucesso.", vbInformation, "Processamento de Recebimentos"
    
    Exit Sub
    
Erro:

    MsgBox "Erro ao processar o recebimento.", vbOKOnly + vbInformation, "Erro ao Carregar Dados"
    
End Sub

Sub processar_recebimento_com_barra()
    
    resposta = MsgBox("Deseja realmente processar recebimentos?", vbYesNo + vbExclamation, "Processamento de Recebimentos")
 
    If resposta = vbYes Then frmBarraProgressao.Show
    
End Sub






