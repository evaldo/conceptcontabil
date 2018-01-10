Attribute VB_Name = "SQL"
Dim strSQLCenario As String
Sub ExportardadosSQL()

On Error GoTo Erro

    Dim ano As String
    Dim mes(1 To 12) As String
    Dim planoClassificacaoPlanoConta(1 To 3000, 1 To 5) As String
    Dim planoPlanoConta(1 To 3000, 1 To 7) As String
    Dim numeroMes As Integer
    Dim mes_processamento As String
    Dim strSQL As String
    Dim ConnectionString As String
    Dim StrQuery As String
    Dim dataTransformada As String
    Dim nomeClie As String
    Dim cnpjClie As String
    Dim codigoClassificacaoPlano As String
    Dim descricaoClassificacaoPlano As String
    Dim codigoPlano As String
    Dim descricaoPlano As String
    Dim indicadorClassificacaoPlanoContas As String
    Dim DS_CMNH_ARQV_ORIG As String
    Dim NU_INIC_LTRA_ARQV_ORIG As String
    Dim NU_FIM_LTRA_ARQV_ORIG As String
    Dim CD_COL_CLSSF_PLANO_CONTA As String
    Dim CD_COL_DIA As String
    Dim CD_COL_DCTO_RFRC_FLUXO_CAIXA As String
    Dim CD_COL_INSTT_FNCR As String
    Dim CD_COL_VL_FLUXO_CAIXA As String

    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim rstTempo As New ADODB.Recordset
    
    Dim linha As Integer
    Dim linhaplanoConta As Integer
    Dim qtFluxo As Integer
    Dim qtRegistroCommit As Integer
    Dim indice As Integer
    Dim indicePlano As Integer
    
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
    
    mes_processamento = ActiveSheet.Name
    
    If ValidaPlanilhaProcessamento() = False Then
        MsgBox "Escolha um planilha para lan�amento do Fluxo de Caixa entre Jan e Dez.", vbOKOnly + vbInformation, "Salvar Dados"
        Exit Sub
    End If
    
    For numeroMes = 1 To 12
        If mes(numeroMes) = mes_processamento Then Exit For
    Next numeroMes
    
    cnn.ConnectionString = "Driver={ODBC Driver 13 for SQL Server};Server=tcp:contarcon.database.windows.net,1433;Database=fluxocaixa;Uid=evaldo@contarcon;Pwd={Gcas1302};Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"
    cnn.Open
    
    Worksheets("Configura��es B�sicas").Activate
    ano = Range("E5").Value
    nomeClie = Range("E9").Value
    cnpjClie = Range("E8").Value
    
    
    linha = 12
    indice = 1
    indicePlano = 1
    
    Do While Range("D" + CStr(linha)).Value <> ""
            
        'C�digo da classifica��o do plano de contas
        planoClassificacaoPlanoConta(indice, 1) = Range("D" + CStr(linha)).Value
        'Descri��o da classifica��o do plano de contas
        planoClassificacaoPlanoConta(indice, 2) = Range("E" + CStr(linha)).Value
        'Indica��o da classifica��o do plano de contas
        planoClassificacaoPlanoConta(indice, 3) = Range("F" + CStr(linha)).Value
        'Coluna do c�digo da classifica��o do plano de contas
        planoClassificacaoPlanoConta(indice, 4) = Range("G" + CStr(linha)).Value
        'Coluna da descri��o da classifica��o do plano de contas
        planoClassificacaoPlanoConta(indice, 5) = Range("H" + CStr(linha)).Value
                
        StrQuery = "SELECT COUNT(1) FROM T_CLSSF_PLANO_CONTA WHERE CD_CLSSF_PLANO_CONTA = '" & Range("D" + CStr(linha)).Value & "'"
        rst.Open (StrQuery), cnn
        
        If rst(0).Value = 0 Then
                
            strSQL = "INSERT INTO T_CLSSF_PLANO_CONTA ("
            strSQL = strSQL + "ID_CLSSF_PLANO_CONTA,"
            strSQL = strSQL + "CD_CLSSF_PLANO_CONTA, "
            strSQL = strSQL + "NU_CNPJ,"
            strSQL = strSQL + "IC_TIPO_TRANS_FLUXO_CAIXA, "
            strSQL = strSQL + "DS_CLSSF_PLANO_CONTA, "
            strSQL = strSQL + "CD_PLANO_CONTA, "
            strSQL = strSQL + "DS_PLANO_CONTA) "
            strSQL = strSQL + "VALUES("
            strSQL = strSQL + "NEXT VALUE FOR SQ_CLSSF_PLANO_CONTA,"
            strSQL = strSQL + "'" & Range("D" + CStr(linha)).Value & "',"
            strSQL = strSQL + "'" & cnpjClie & "',"
            strSQL = strSQL + "'" & Range("F" + CStr(linha)).Value & "',"
            strSQL = strSQL + "'" & Range("E" + CStr(linha)).Value & "',"
            strSQL = strSQL + "'" & Range("D" + CStr(linha)).Value & "',"
            strSQL = strSQL + "'" & Range("E" + CStr(linha)).Value & "');"
            
            cnn.Execute strSQL
                        
            'Indica��o da classifica��o do plano de contas
            If planoClassificacaoPlanoConta(indice, 3) = "R" Then
                Worksheets("PC Receitas").Activate
            Else
                Worksheets("PC Despesas").Activate
            End If
            
            linhaplanoConta = 6
            
            Do While Range(planoClassificacaoPlanoConta(indice, 4) + CStr(linhaplanoConta)).Value <> ""
            
                strSQL = "INSERT INTO T_CLSSF_PLANO_CONTA (ID_CLSSF_PLANO_CONTA, CD_CLSSF_PLANO_CONTA, NU_CNPJ,IC_TIPO_TRANS_FLUXO_CAIXA, DS_CLSSF_PLANO_CONTA, CD_PLANO_CONTA, DS_PLANO_CONTA) "
                strSQL = strSQL + "VALUES("
                strSQL = strSQL + "NEXT VALUE FOR SQ_CLSSF_PLANO_CONTA, "
                strSQL = strSQL + "'" & planoClassificacaoPlanoConta(indice, 1) & "', "
                strSQL = strSQL + "'" & cnpjClie & "', '" & planoClassificacaoPlanoConta(indice, 3) & "',"
                strSQL = strSQL + "'" & planoClassificacaoPlanoConta(indice, 2) & "', '" & Range(planoClassificacaoPlanoConta(indice, 4) + CStr(linhaplanoConta)).Value & "',"
                strSQL = strSQL + "'" & Range(planoClassificacaoPlanoConta(indice, 5) + CStr(linhaplanoConta)).Value & "');"
            
                cnn.Execute strSQL
                                
                'C�digo da classifica��o do plano de contas
                planoPlanoConta(indicePlano, 1) = planoClassificacaoPlanoConta(indice, 1)
                'Descri��o da classifica��o do plano de contas
                planoPlanoConta(indicePlano, 2) = planoClassificacaoPlanoConta(indice, 2)
                'Indica��o da classifica��o do plano de contas
                planoPlanoConta(indicePlano, 3) = planoClassificacaoPlanoConta(indice, 3)
                'Coluna do c�digo da classifica��o do plano de contas
                planoPlanoConta(indicePlano, 4) = planoClassificacaoPlanoConta(indice, 4)
                'Coluna da descri��o da classifica��o do plano de contas
                planoPlanoConta(indicePlano, 5) = planoClassificacaoPlanoConta(indice, 5)
                'Codigo do plano de contas
                planoPlanoConta(indicePlano, 6) = Range(planoClassificacaoPlanoConta(indice, 4) + CStr(linhaplanoConta)).Value
                'Descri��o do plano de contas
                planoPlanoConta(indicePlano, 7) = Range(planoClassificacaoPlanoConta(indice, 5) + CStr(linhaplanoConta)).Value
                
                indicePlano = indicePlano + 1
                
                linhaplanoConta = linhaplanoConta + 1
            
            Loop
            
            Worksheets("Configura��es B�sicas").Activate
            
        Else
                
            strSQL = "UPDATE T_CLSSF_PLANO_CONTA SET NU_CNPJ = '" & cnpjClie & "', IC_TIPO_TRANS_FLUXO_CAIXA = '" & Range("F" + CStr(linha)).Value & "', DS_CLSSF_PLANO_CONTA = '" & Range("E" + CStr(linha)).Value & "'"
            strSQL = strSQL + "WHERE CD_CLSSF_PLANO_CONTA = '" & Range("D" + CStr(linha)).Value & "';"
            
            cnn.Execute strSQL
        
            'Indica��o da classifica��o do plano de contas
            If planoClassificacaoPlanoConta(indice, 3) = "R" Then
                Worksheets("PC Receitas").Activate
            Else
                Worksheets("PC Despesas").Activate
            End If
            
            linhaplanoConta = 6
            
            Do While Range(planoClassificacaoPlanoConta(indice, 4) + CStr(linhaplanoConta)).Value <> ""
            
                strSQL = "UPDATE T_CLSSF_PLANO_CONTA SET NU_CNPJ = '" & cnpjClie & "'"
                strSQL = strSQL + ", IC_TIPO_TRANS_FLUXO_CAIXA = '" & planoClassificacaoPlanoConta(indice, 3) & "'"
                strSQL = strSQL + ", DS_CLSSF_PLANO_CONTA = '" & planoClassificacaoPlanoConta(indice, 2) & "'"
                strSQL = strSQL + ", CD_CLSSF_PLANO_CONTA = '" & planoClassificacaoPlanoConta(indice, 1) & "'"
                strSQL = strSQL + ", DS_PLANO_CONTA = '" & Range(planoClassificacaoPlanoConta(indice, 5) + CStr(linhaplanoConta)).Value & "'"
                strSQL = strSQL + "WHERE CD_PLANO_CONTA = '" & Range(planoClassificacaoPlanoConta(indice, 4) + CStr(linhaplanoConta)).Value & "';"
            
                cnn.Execute strSQL
                
                'C�digo da classifica��o do plano de contas
                planoPlanoConta(indicePlano, 1) = planoClassificacaoPlanoConta(indice, 1)
                'Descri��o da classifica��o do plano de contas
                planoPlanoConta(indicePlano, 2) = planoClassificacaoPlanoConta(indice, 2)
                'Indica��o da classifica��o do plano de contas
                planoPlanoConta(indicePlano, 3) = planoClassificacaoPlanoConta(indice, 3)
                'Coluna do c�digo da classifica��o do plano de contas
                planoPlanoConta(indicePlano, 4) = planoClassificacaoPlanoConta(indice, 4)
                'Coluna da descri��o da classifica��o do plano de contas
                planoPlanoConta(indicePlano, 5) = planoClassificacaoPlanoConta(indice, 5)
                'Codigo do plano de contas
                planoPlanoConta(indicePlano, 6) = Range(planoClassificacaoPlanoConta(indice, 4) + CStr(linhaplanoConta)).Value
                'Descri��o do plano de contas
                planoPlanoConta(indicePlano, 7) = Range(planoClassificacaoPlanoConta(indice, 5) + CStr(linhaplanoConta)).Value
                
                indicePlano = indicePlano + 1
                
                linhaplanoConta = linhaplanoConta + 1
            
            Loop
            
            Worksheets("Configura��es B�sicas").Activate
        
        End If
        
        indice = indice + 1
        linha = linha + 1
        
        rst.Close
        
    Loop
        
    Worksheets("Cenario Despesas").Activate
    linha = 5
    
    DS_CMNH_ARQV_ORIG = Range("K" + CStr(linha)).Value
    NU_INIC_LTRA_ARQV_ORIG = Range("L" + CStr(linha)).Value
    NU_FIM_LTRA_ARQV_ORIG = Range("M" + CStr(linha)).Value
    CD_COL_CLSSF_PLANO_CONTA = Range("N" + CStr(linha)).Value
    CD_COL_DIA = Range("P" + CStr(linha)).Value
    CD_COL_DCTO_RFRC_FLUXO_CAIXA = Range("Q" + CStr(linha)).Value
    CD_COL_INSTT_FNCR = Range("R" + CStr(linha)).Value
    CD_COL_VL_FLUXO_CAIXA = Range("S" + CStr(linha)).Value
    
    strSQL = "DELETE FROM T_CNRIO_IMPRT_ARQV "
    strSQL = strSQL + "WHERE NU_ANO_PLAN_ORIG_PROC = " & ano & " "
    strSQL = strSQL + "and NU_CNPJ = '" & cnpjClie & "';"
    
    cnn.Execute strSQL
    
    Do While Range("G" + CStr(linha)).Value <> ""
    
        Call insturcaoSQLCenario(cnpjClie, _
                        Range("G" + CStr(linha)).Value, _
                        Range("H" + CStr(linha)).Value, _
                        Range("I" + CStr(linha)).Value, _
                        Range("J" + CStr(linha)).Value, _
                        DS_CMNH_ARQV_ORIG, _
                        NU_INIC_LTRA_ARQV_ORIG, _
                        NU_FIM_LTRA_ARQV_ORIG, _
                        CD_COL_CLSSF_PLANO_CONTA, _
                        CD_COL_DIA, _
                        CD_COL_DCTO_RFRC_FLUXO_CAIXA, _
                        CD_COL_INSTT_FNCR, _
                        CD_COL_VL_FLUXO_CAIXA, _
                        "D", _
                        ano)
                        
        cnn.Execute strSQLCenario
        
        linha = linha + 1
        
    Loop
        
    linha = 5
    
    strSQL = "DELETE FROM T_LISTA_PLVR_EXCD "
    strSQL = strSQL + "WHERE NU_ANO_PLAN_ORIG_PROC = " & ano & " "
    strSQL = strSQL + "and NU_CNPJ = '" & cnpjClie & "';"
    
    cnn.Execute strSQL
    
    Do While Range("O" + CStr(linha)).Value <> ""
        
        strSQL = "INSERT INTO T_LISTA_PLVR_EXCD"
        strSQL = strSQL + "("
        strSQL = strSQL + "ID_LISTA_PLVR_EXCD"
        strSQL = strSQL + ",NU_CNPJ"
        strSQL = strSQL + ",NU_ANO_PLAN_ORIG_PROC"
        strSQL = strSQL + ",DS_PLVR_EXCD"
        strSQL = strSQL + ",IC_TIPO_TRANS_FLUXO_CAIXA"
        strSQL = strSQL + ")"
        strSQL = strSQL + "VALUES"
        strSQL = strSQL + "("
        strSQL = strSQL + "NEXT VALUE FOR SQ_LISTA_PLVR_EXCD"
        strSQL = strSQL + ",'" & cnpjClie & "'"
        strSQL = strSQL + ",'" & ano & "'"
        strSQL = strSQL + ",'" & Range("O" + CStr(linha)).Value & "'"
        strSQL = strSQL + ",'D'"
        strSQL = strSQL + ");"
        
        cnn.Execute strSQL
        
        linha = linha + 1
        
    Loop
        
    Worksheets("Cenario Receitas").Activate
    linha = 5
    
    Do While Range("G" + CStr(linha)).Value <> ""
    
         Call insturcaoSQLCenario(cnpjClie, _
                        Range("G" + CStr(linha)).Value, _
                        Range("H" + CStr(linha)).Value, _
                        Range("I" + CStr(linha)).Value, _
                        Range("J" + CStr(linha)).Value, _
                        DS_CMNH_ARQV_ORIG, _
                        NU_INIC_LTRA_ARQV_ORIG, _
                        NU_FIM_LTRA_ARQV_ORIG, _
                        CD_COL_CLSSF_PLANO_CONTA, _
                        CD_COL_DIA, _
                        CD_COL_DCTO_RFRC_FLUXO_CAIXA, _
                        CD_COL_INSTT_FNCR, _
                        CD_COL_VL_FLUXO_CAIXA, _
                        "R", _
                        ano)
                        
        cnn.Execute strSQLCenario
        
        linha = linha + 1
        
    Loop
    
    linha = 5
    
    Do While Range("O" + CStr(linha)).Value <> ""
        
        strSQL = "INSERT INTO T_LISTA_PLVR_EXCD"
        strSQL = strSQL + "("
        strSQL = strSQL + "ID_LISTA_PLVR_EXCD"
        strSQL = strSQL + ",NU_CNPJ"
        strSQL = strSQL + ",NU_ANO_PLAN_ORIG_PROC"
        strSQL = strSQL + ",DS_PLVR_EXCD"
        strSQL = strSQL + ",IC_TIPO_TRANS_FLUXO_CAIXA"
        strSQL = strSQL + ")"
        strSQL = strSQL + "VALUES"
        strSQL = strSQL + "("
        strSQL = strSQL + "NEXT VALUE FOR SQ_LISTA_PLVR_EXCD"
        strSQL = strSQL + ",'" & cnpjClie & "'"
        strSQL = strSQL + ",'" & ano & "'"
        strSQL = strSQL + ",'" & Range("O" + CStr(linha)).Value & "'"
        strSQL = strSQL + ",'R'"
        strSQL = strSQL + ");"
        
        cnn.Execute strSQL
        
        linha = linha + 1
        
    Loop
    
    Worksheets(mes_processamento).Activate
    
    cnn.BeginTrans
    
    'StrQuery = "SELECT COUNT(1), MAX(ID_FLUXO_CAIXA)+1 FROM T_FLUXO_CAIXA"
    'rst.Open (StrQuery), cnn
    
    'If rst(0).Value = 0 Then
    '    qtFluxo = 1
    'Else
    '    qtFluxo = rst(1).Value
    'End If
    
    'rst.Close
    
    linha = 5
    qtRegistroCommit = 0
    
    strSQL = "DELETE FROM T_FLUXO_CAIXA "
    strSQL = strSQL + "WHERE NU_ANO_PLAN_ORIG_PROC = " & ano & " "
    strSQL = strSQL + "and DS_PLAN_ORIG_PROC = '" & UCase(mes_processamento) & "' "
    strSQL = strSQL + "and NU_CNPJ = '" & cnpjClie & "';"
    
    cnn.Execute strSQL
    
    Do While Range("C" + CStr(linha)).Value <> ""
    
        If Not IsDate("" & numeroMes & "/" & Range("C" + CStr(linha)).Value & "/" & ano & "") Then
            StrQuery = "SELECT ID_DMSAO_TEMPO FROM T_DMSAO_TEMPO WHERE DT_DMSAO_TEMPO = CONVERT(VARCHAR(10), '" & UltimoDiaMes(CDate("1/" & numeroMes & "/" & ano)) & "', 103)"
            dataTransformada = UltimoDiaMes(CDate("1/" & numeroMes & "/" & ano))
        Else
            StrQuery = "SELECT ID_DMSAO_TEMPO FROM T_DMSAO_TEMPO WHERE DT_DMSAO_TEMPO = CONVERT(VARCHAR(10), '" & numeroMes & "/" & Range("C" + CStr(linha)).Value & "/" & ano & "', 103)"
            dataTransformada = "" & numeroMes & "/" & Range("C" + CStr(linha)).Value & "/" & ano & ""
        End If
        
        rstTempo.Open (StrQuery), cnn
        
        indicePlano = 1
        
        Do While indicePlano <= UBound(planoPlanoConta)
            
            If planoPlanoConta(indicePlano, 7) = Range("G" + CStr(linha)).Value Then
                    
                codigoClassificacaoPlano = planoPlanoConta(indicePlano, 1)
                descricaoClassificacaoPlano = planoPlanoConta(indicePlano, 2)
                indicadorClassificacaoPlanoContas = planoPlanoConta(indicePlano, 3)
                codigoPlano = planoPlanoConta(indicePlano, 6)
                descricaoPlano = planoPlanoConta(indicePlano, 7)
                
                Exit Do
                
            End If
            
            indicePlano = indicePlano + 1
            
        Loop
        
        strSQL = "INSERT INTO T_FLUXO_CAIXA ("
        strSQL = strSQL + "  ID_FLUXO_CAIXA"
        strSQL = strSQL + ", NU_CNPJ"
        strSQL = strSQL + ", SK_DMSAO_TEMPO"
        strSQL = strSQL + ", DT_MVMT_FLUXO_CAIXA"
        strSQL = strSQL + ", NM_CLIE_FLUXO_CAIXA"
        strSQL = strSQL + ", DS_CLSSF_PLANO_CONTA"
        strSQL = strSQL + ", CD_DCTO_RFRC_FLUXO_CAIXA"
        strSQL = strSQL + ", CD_PLANO_CONTA"
        strSQL = strSQL + ", DS_PLANO_CONTA"
        strSQL = strSQL + ", DS_INSTT_FNCR"
        strSQL = strSQL + ", VL_ENTR_FLUXO_CAIXA"
        strSQL = strSQL + ", VL_SAIDA_FLUXO_CAIXA"
        strSQL = strSQL + ", IC_STATUS_VALOR"
        strSQL = strSQL + ", NU_MATL_INCS"
        strSQL = strSQL + ", DT_INCS"
        strSQL = strSQL + ", IC_TIPO_TRANS_FLUXO_CAIXA"
        strSQL = strSQL + ", DS_PLAN_ORIG_PROC"
        strSQL = strSQL + ", CD_CLSSF_PLANO_CONTA"
        strSQL = strSQL + ", ID_CLSSF_PLANO_CONTA"
        strSQL = strSQL + ", NU_ANO_PLAN_ORIG_PROC"
        strSQL = strSQL + " ) "
        strSQL = strSQL + "VALUES("
        strSQL = strSQL + "NEXT VALUE FOR SQ_FLUXO_CAIXA,"
        strSQL = strSQL + "'" & cnpjClie & "',"
        strSQL = strSQL + "" & rstTempo(0).Value & ","
        strSQL = strSQL + "CONVERT(VARCHAR(10), '" & dataTransformada & "', 103),"
        strSQL = strSQL + "'" & nomeClie & "',"
        strSQL = strSQL + "'" & UCase(Range("E" + CStr(linha)).Value) & "',"
        strSQL = strSQL + "'" & Range("F" + CStr(linha)).Value & "',"
        strSQL = strSQL + "'" & codigoPlano & "',"
        strSQL = strSQL + "'" & UCase(Range("G" + CStr(linha)).Value) & "',"
        strSQL = strSQL + "'" & IIf(Range("H" + CStr(linha)).Value = "", "RECEITA", Range("H" + CStr(linha)).Value) & "',"
        strSQL = strSQL + "'" & Replace(Range("J" + CStr(linha)).Value, ",", ".") & "',"
        strSQL = strSQL + "'" & Replace(Range("K" + CStr(linha)).Value, ",", ".") & "',"
        strSQL = strSQL + "'" & Range("L" + CStr(linha)).Value & "',"
        strSQL = strSQL + "'" & cnpjClie & "',"
        strSQL = strSQL + "getdate(),"
        strSQL = strSQL + "'" & indicadorClassificacaoPlanoContas & "',"
        strSQL = strSQL + "'" & UCase(mes_processamento) & "',"
        strSQL = strSQL + "'" & codigoClassificacaoPlano & "',"
        strSQL = strSQL + "(SELECT ID_CLSSF_PLANO_CONTA FROM T_CLSSF_PLANO_CONTA WHERE CD_CLSSF_PLANO_CONTA = '" & codigoClassificacaoPlano & "' AND CD_PLANO_CONTA = '" & codigoPlano & "'),"
        strSQL = strSQL + "" & ano & ""
        strSQL = strSQL + ");"
        
        cnn.Execute strSQL
        
        If qtRegistroCommit = 10 Then
            cnn.CommitTrans
            qtRegistroCommit = 0
            cnn.BeginTrans
        End If
        
        linha = linha + 1
        'qtFluxo = qtFluxo + 1
        qtRegistroCommit = qtRegistroCommit + 1
        
        rstTempo.Close
        
    Loop
    
    cnn.CommitTrans
    cnn.Close
    
    Exit Sub
    
Erro:

    MsgBox "Erro ao processar o envio de dados para nuvem. " + Err.Description, vbOKOnly + vbInformation, "Envio de Dados para Nuvem"
    Worksheets(mes_processamento).Activate
    
    Exit Sub

End Sub

Function UltimoDiaMes(Data As Date) As String

    UltimoDiaMes = VBA.DateSerial(VBA.Year(Data), VBA.Month(Data) + 1, 0)
    UltimoDiaMes = CStr(Month(UltimoDiaMes) & "/" & Day(UltimoDiaMes) & "/" & Year(UltimoDiaMes))

End Function

Sub insturcaoSQLCenario(pNU_CNPJ As String, _
                        pDS_CONTA_CLIE As String, _
                        pDS_CLSSF_PLANO_CONTA As String, _
                        pCD_PLANO_CONTA As String, _
                        pDS_PLANO_CONTA As String, _
                        pDS_CMNH_ARQV_ORIG As String, _
                        pNU_INIC_LTRA_ARQV_ORIG As String, _
                        pNU_FIM_LTRA_ARQV_ORIG As String, _
                        pCD_COL_CLSSF_PLANO_CONTA As String, _
                        pCD_COL_DIA As String, _
                        pCD_COL_DCTO_RFRC_FLUXO_CAIXA As String, _
                        pCD_COL_INSTT_FNCR As String, _
                        pCD_COL_VL_FLUXO_CAIXA As String, _
                        pIC_TIPO_TRANS_FLUXO_CAIXA As String, _
                        pNU_ANO_PLAN_ORIG_PROC As String)

    strSQLCenario = "INSERT INTO T_CNRIO_IMPRT_ARQV"
    strSQLCenario = strSQLCenario + "("
    strSQLCenario = strSQLCenario + "ID_CNRIO_IMPRT_ARQV"
    strSQLCenario = strSQLCenario + ",NU_CNPJ"
    strSQLCenario = strSQLCenario + ",DS_CONTA_CLIE"
    strSQLCenario = strSQLCenario + ",DS_CLSSF_PLANO_CONTA"
    strSQLCenario = strSQLCenario + ",CD_PLANO_CONTA"
    strSQLCenario = strSQLCenario + ",DS_PLANO_CONTA"
    strSQLCenario = strSQLCenario + ",DS_CMNH_ARQV_ORIG"
    strSQLCenario = strSQLCenario + ",NU_INIC_LTRA_ARQV_ORIG"
    strSQLCenario = strSQLCenario + ",NU_FIM_LTRA_ARQV_ORIG"
    strSQLCenario = strSQLCenario + ",CD_COL_CLSSF_PLANO_CONTA"
    strSQLCenario = strSQLCenario + ",CD_COL_DIA"
    strSQLCenario = strSQLCenario + ",CD_COL_DCTO_RFRC_FLUXO_CAIXA"
    strSQLCenario = strSQLCenario + ",CD_COL_INSTT_FNCR"
    strSQLCenario = strSQLCenario + ",CD_COL_VL_FLUXO_CAIXA"
    strSQLCenario = strSQLCenario + ",IC_TIPO_TRANS_FLUXO_CAIXA"
    strSQLCenario = strSQLCenario + ",NU_ANO_PLAN_ORIG_PROC"
    strSQLCenario = strSQLCenario + ")"
    strSQLCenario = strSQLCenario + "VALUES"
    strSQLCenario = strSQLCenario + "("
    strSQLCenario = strSQLCenario + "NEXT VALUE FOR SQ_CNRIO_IMPRT_ARQV"
    strSQLCenario = strSQLCenario + ",'" & pNU_CNPJ & "'"
    strSQLCenario = strSQLCenario + ",'" & pDS_CONTA_CLIE & "'"
    strSQLCenario = strSQLCenario + ",'" & pDS_CLSSF_PLANO_CONTA & "'"
    strSQLCenario = strSQLCenario + ",'" & pCD_PLANO_CONTA & "'"
    strSQLCenario = strSQLCenario + ",'" & pDS_PLANO_CONTA & "'"
    strSQLCenario = strSQLCenario + ",'" & pDS_CMNH_ARQV_ORIG & "'"
    strSQLCenario = strSQLCenario + ",'" & pNU_INIC_LTRA_ARQV_ORIG & "'"
    strSQLCenario = strSQLCenario + ",'" & pNU_FIM_LTRA_ARQV_ORIG & "'"
    strSQLCenario = strSQLCenario + ",'" & pCD_COL_CLSSF_PLANO_CONTA & "'"
    strSQLCenario = strSQLCenario + ",'" & pCD_COL_DIA & "'"
    strSQLCenario = strSQLCenario + ",'" & pCD_COL_DCTO_RFRC_FLUXO_CAIXA & "'"
    strSQLCenario = strSQLCenario + ",'" & pCD_COL_INSTT_FNCR & "'"
    strSQLCenario = strSQLCenario + ",'" & pCD_COL_VL_FLUXO_CAIXA & "'"
    strSQLCenario = strSQLCenario + ",'" & pIC_TIPO_TRANS_FLUXO_CAIXA & "'"
    strSQLCenario = strSQLCenario + ",'" & pNU_ANO_PLAN_ORIG_PROC & "'"
    strSQLCenario = strSQLCenario + ");"

End Sub
