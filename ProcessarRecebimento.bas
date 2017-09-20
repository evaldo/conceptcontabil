Attribute VB_Name = "ProcessarRecebimento"
Sub processa_recebimento_caixa()
    
    Dim plano_conta As String
    Dim mes(1 To 12) As String
    Dim celula_planilha_atual As String
    Dim celula_planilha_lida As String
    Dim mes_processamento As String
    Dim doc_ref As String
    Dim instituicao_finaceira As String
    Dim plano_contas As String
            
    Dim linha_planilha As Integer
    Dim conta_mes As Integer
    Dim flag_mes_processamento As Integer
    Dim linha_planilha_mes_processmento As Integer
    Dim linha_planilha_recebimento As Integer
    Dim contador As Integer
    
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
        
        bol_processar_classificacao = False
                
        If classificacao = "RECEITAS COM PRODUTO" Then
        
           Sheets("PC Receitas").Select
           linha_planilha_recebimento = 5
           
           Do While Range("D" + CStr(linha_planilha_recebimento)).Value <> ""
           
            If Range("D" + CStr(linha_planilha_recebimento)).Value = plano_contas And (Not IsEmpty(Range("E" + CStr(linha_planilha_recebimento)).Value) Or Range("E" + CStr(linha_planilha_recebimento)).Value = "-") Then
                bol_processar_classificacao = True
                Exit Do
            End If
            
            linha_planilha_recebimento = linha_planilha_recebimento + 1
            
           Loop
           
        End If
       
        If bol_processar_classificacao = True Then
       
            For conta_mes = 1 To 12
                
                If conta_mes > flag_mes_processamento Then Exit For
                        
                Sheets(mes(conta_mes)).Select
                percentual = conta_mes / flag_mes_processamento
                frmBarraProgressao.AtualizaBarra percentual, mes(conta_mes)
                               
                Do While Range("E" + CStr(linha_planilha)).Value <> ""
                            
                    If Range("E" + CStr(linha_planilha)).Value = classificacao _
                       And Range("H" + CStr(linha_planilha)).Value = instituicao_finaceira _
                       And Range("F" + CStr(linha_planilha)).Value = doc_ref _
                       And Range("G" + CStr(linha_planilha)).Value = plano_contas _
                       And Range("K" + CStr(linha_planilha)).Value = "Pago" _
                       Then
                        
                            Sheets(mes(flag_mes_processamento)).Select
                            Range("I" + CStr(linha_planilha_mes_processmento)).Value = 0
                            Range("L" + CStr(linha_planilha_mes_processmento)).Value = "Sim"
                            Range("K" + CStr(linha_planilha_mes_processmento)).Value = "Realizado"
                            
                            Exit Do
                        
                    End If
                
                    linha_planilha = linha_planilha + 1
                Loop
                
                linha_planilha = 5
                Range("E" + CStr(linha_planilha)).Select
                
            Next conta_mes
            
        End If
        
        Sheets(mes(flag_mes_processamento)).Select
        linha_planilha_mes_processmento = linha_planilha_mes_processmento + 1
        
    Loop
    
    frmBarraProgressao.Hide
    
    MsgBox "Processamento Realizado com sucesso.", vbInformation, "Processamento de Recebimentos"
    
End Sub

Sub processar_recebimento_com_barra()
    
    resposta = MsgBox("Deseja realmente processar recebimentos?", vbYesNo + vbExclamation, "Processamento de Recebimentos")
 
    If resposta = vbYes Then frmBarraProgressao.Show
    
End Sub

