VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRecuperarDadosNuvem 
   Caption         =   "Recuperar Dados da Nuvem"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6525
   OleObjectBlob   =   "frmRecuperarDadosNuvem.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRecuperarDadosNuvem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nomePlanilha As String
Dim linha As Integer
Dim cnn As New ADODB.Connection
Dim rst As New ADODB.Recordset
Dim rstQtdePlanoConta As New ADODB.Recordset
Dim cnpjParam As String
Dim anoRecuperacao As Integer

Private Sub cmdFechar_Click()

    Unload Me

End Sub

Private Sub cmdRecuperarDados_Click()

On Error GoTo Erro

Dim linha As Integer
Dim lista As Integer

Dim colunaCodigoPlanoContaPlanilha As String
Dim colunaDescricaoPlanoContaPlanilha As String
Dim planoConta As String
Dim icTipoPlanoConta As String
Dim strSQLQtdePlanoConta As String

Dim bolExisteQtdePlanoConta As Boolean

    frmLoginManterDadosNuvem.Show

    If manterDadosAposLogin = False Then
        MsgBox "Login inválido não foi possível manter dados.", vbOKOnly + vbInformation, "Salvar Dados"
        Exit Sub
    End If
    
    manterDadosAposLogin = False

    frmProgresso.Visible = True

    'cnn.ConnectionString = "Driver={ODBC Driver 13 for SQL Server};Server=tcp:contarcondb.cmxd2lqddzlw.sa-east-1.rds.amazonaws.com,1433;Database=fluxocaixa;Uid=" & usuario & ";Pwd={" & senha & "};Encrypt=yes;TrustServerCertificate=no;Connection Timeout=300;"
    cnn.ConnectionString = "Provider=SQLOLEDB.1;Password=" & senha & ";Persist Security Info=True;User ID=" & usuario & ";Initial Catalog=fluxocaixa;Data Source=contarcondb.cmxd2lqddzlw.sa-east-1.rds.amazonaws.com,1433"
    cnn.Open
    
    nomePlanilha = ActiveSheet.Name
    
    Worksheets("Configurações Básicas").Activate
        
    ano = Range("E5").Value
    nomeClie = Range("E9").Value
    cnpjClie = Range("E8").Value
    
    cnpjParam = cnpjClie
    anoRecuperacao = ano
    
    Application.ScreenUpdating = False
    
    If Me.chkPlanoContas.Value = True Then
    
        Application.EnableEvents = False
    
        bolExisteQtdePlanoConta = False
        
        Call barraProgresso("Processando a recuperação de dados...", 1)
        
        '----------------------------------------------------------------------------------------------
        'Recuperação do Plano de Contas
        '----------------------------------------------------------------------------------------------
        Call limparDadosPlanoContas("D")
        Call limparDadosPlanoContas("R")
            
        strSQL = "SELECT CD_CLSSF_PLANO_CONTA"
        strSQL = strSQL + "  ,NU_CNPJ"
        strSQL = strSQL + "  ,IC_TIPO_TRANS_FLUXO_CAIXA"
        strSQL = strSQL + "  ,DS_CLSSF_PLANO_CONTA"
        strSQL = strSQL + "  ,CD_PLANO_CONTA"
        strSQL = strSQL + "  ,DS_PLANO_CONTA"
        strSQL = strSQL + "  ,CD_CLUN_CDGO_CLSSF_PLANO_CONTA"
        strSQL = strSQL + "  ,CD_CLUN_DSCR_PLANO_CONTA "
        strSQL = strSQL + " FROM T_CLSSF_PLANO_CONTA "
        strSQL = strSQL + " WHERE NU_CNPJ = '" & cnpjClie & "'"
        strSQL = strSQL + " ORDER BY CD_CLSSF_PLANO_CONTA"
        strSQL = strSQL + "  ,IC_TIPO_TRANS_FLUXO_CAIXA"
        strSQL = strSQL + "  ,CD_PLANO_CONTA"
        
         With rst
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .Open strSQL, cnn
        End With
        
        If rst.BOF = False Then
        
            If rst.EOF = False Then rst.MoveFirst
            
            colunaCodigoPlanoContaPlanilha = rst(6).Value
            colunaDescricaoPlanoContaPlanilha = rst(7).Value
            planoConta = rst(0).Value
            icTipoPlanoConta = rst(2).Value
            
            linha = 4
            
            If rst(2).Value = "D" Then
                Worksheets("PC Despesas").Activate
            Else
                Worksheets("PC Receitas").Activate
            End If
            
            Range(colunaCodigoPlanoContaPlanilha + CStr(linha)) = "'" + rst(0).Value
            Range(colunaDescricaoPlanoContaPlanilha + CStr(linha)) = rst(3).Value
            
            rst.MoveNext
            
            linha = linha + 1
            
            Do While rst.EOF = False
            
                 Call barraProgresso("Recuperando dados do Plano de Contas de " + IIf(rst(3).Value = "D", "Despesas", "Receitas") + "", linha)
                
                 If rst(2).Value <> icTipoPlanoConta Then
                    icTipoPlanoConta = rst(2).Value
                    If rst(2).Value = "D" Then
                        Worksheets("PC Despesas").Activate
                        linha = 4
                    Else
                        Worksheets("PC Receitas").Activate
                        linha = 4
                    End If
                 End If
                 
                 If planoConta <> rst(0).Value Then
                    
                    linha = 4
                    
                    colunaCodigoPlanoContaPlanilha = rst(6).Value
                    colunaDescricaoPlanoContaPlanilha = rst(7).Value
                    planoConta = rst(0).Value
                    
                    Range(colunaCodigoPlanoContaPlanilha + CStr(linha)) = "'" + rst(0).Value
                    Range(colunaDescricaoPlanoContaPlanilha + CStr(linha)) = rst(3).Value
                    
                    linha = linha + 1
                    
                    rst.MoveNext
                    
                 End If
                
                 If rst.EOF = False Then
                 
                    Range(colunaCodigoPlanoContaPlanilha + CStr(linha)) = rst(4).Value
                    Range(colunaDescricaoPlanoContaPlanilha + CStr(linha)) = rst(5).Value
                    
                    rst.MoveNext
                        
                 End If
                    
                 linha = linha + 1
                
                 bolExisteQtdePlanoConta = False
                
            Loop
            
        End If
        
        rst.Close
        
        Application.EnableEvents = True
        
    End If
    
    If Me.chkCenarioImportacao.Value = True Then
    
        Application.EnableEvents = False
        
        Call carregarDadosDeImportacao("R")
        Call carregarDadosDePalavrasImportacao("R")
        Call carregarDadosDeImportacao("D")
        Call carregarDadosDePalavrasImportacao("D")
        Call carregarDadosDeImportacao("N")
        Call carregarDadosDePalavrasImportacao("N")
        
        Application.EnableEvents = True
        
    End If
    
    If Me.chkCenarioExportacao.Value = True Then
        
        Application.EnableEvents = False
        
        Call carregarDadosDeExportacao("DOCREF")
        Call carregarDadosDeExportacao("INSTFIN")
        
        Application.EnableEvents = True
        
    End If
    
    If Me.chkMesAtual.Value = True Then
        
        Worksheets(nomePlanilha).Activate
        Call carregarDadosDoMes(nomePlanilha)
        
    End If
    
    For lista = 0 To Me.lstMeses.ListCount - 1
        
        If Me.lstMeses.Selected(lista) = True Then
            Worksheets(Me.lstMeses.List(lista)).Activate
            Call carregarDadosDoMes(Me.lstMeses.List(lista))
            Me.lstMeses.Selected(lista) = False
        End If
        
    Next lista
    
    cnn.Close
    
    frmProgresso.Visible = False
    Application.ScreenUpdating = True
    Worksheets(nomePlanilha).Activate
    
    Exit Sub
    
Erro:
    
    MsgBox "Erro ao recuperar os dados selecionados. Refaça a operação.", vbOKOnly, "Recuperar dados da nuvem"
    
    Set cnn = Nothing
    
    frmProgresso.Visible = False
    Application.ScreenUpdating = True
    Worksheets(nomePlanilha).Activate
    Application.EnableEvents = True
    
    Exit Sub

End Sub


Private Sub UserForm_Initialize()
    
Dim todosMeses(1 To 12) As String
    
    nomePlanilha = ActiveSheet.Name
    
    todosMeses(1) = "Jan"
    todosMeses(2) = "Fev"
    todosMeses(3) = "Mar"
    todosMeses(4) = "Abr"
    todosMeses(5) = "Mai"
    todosMeses(6) = "Jun"
    todosMeses(7) = "Jul"
    todosMeses(8) = "Ago"
    todosMeses(9) = "Set"
    todosMeses(10) = "Out"
    todosMeses(11) = "Nov"
    todosMeses(12) = "Dez"
    
    Me.lstMeses.List = todosMeses
    
End Sub

Public Sub barraProgresso(mensagem As String, percentual As Integer)

    Me.lblDescricaoProgresso.Caption = mensagem + "... " + CStr(percentual) + " registros"
    DoEvents
    Me.lblProgresso.Width = ((percentual / 10) * Me.lblDescricaoProgresso.Width)
    DoEvents
    
End Sub

Public Sub limparDadosPlanoContas(icTipoPlanoDados As String)
        
    If icTipoPlanoDados = "D" Then
        Worksheets("PC Despesas").Activate
        
        Range("C5:Y105").Select
        Selection.ClearContents
        Range("C5").Select
        
    Else
        Worksheets("PC Receitas").Activate
        
        Range("C5:M106").Select
        Selection.ClearContents
        Range("C5").Select
        
    End If
        
End Sub


Public Sub carregarDadosDeImportacao(icTipoPlanoDados As String)

Dim descricaoTipoProcessamento As String

Dim n As Integer
Dim qtdeRegistros As Integer

    strSQL = "SELECT DS_CONTA_CLIE"
    strSQL = strSQL + "  ,DS_CLSSF_PLANO_CONTA"
    strSQL = strSQL + "  ,CD_PLANO_CONTA"
    strSQL = strSQL + "  ,DS_PLANO_CONTA"
    strSQL = strSQL + "  ,DS_CMNH_ARQV_ORIG"
    strSQL = strSQL + "  ,NU_INIC_LTRA_ARQV_ORIG"
    strSQL = strSQL + "  ,NU_FIM_LTRA_ARQV_ORIG"
    strSQL = strSQL + "  ,CD_COL_CLSSF_PLANO_CONTA"
    strSQL = strSQL + "  ,CD_COL_DIA"
    strSQL = strSQL + "  ,CD_COL_DCTO_RFRC_FLUXO_CAIXA"
    strSQL = strSQL + "  ,CD_COL_INSTT_FNCR"
    strSQL = strSQL + "  ,CD_COL_VL_FLUXO_CAIXA"
    strSQL = strSQL + "  ,IC_TIPO_TRANS_FLUXO_CAIXA"
    strSQL = strSQL + "  FROM T_CNRIO_IMPRT_ARQV"
    
    If icTipoPlanoDados = "N" Then
        strSQL = strSQL + "  WHERE (IC_TIPO_TRANS_FLUXO_CAIXA <> 'R' AND IC_TIPO_TRANS_FLUXO_CAIXA <> 'D')"
    Else
        strSQL = strSQL + "  WHERE IC_TIPO_TRANS_FLUXO_CAIXA = '" & icTipoPlanoDados & " '"
    End If
    
    strSQL = strSQL + "    AND NU_ANO_PLAN_ORIG_PROC = " & anoRecuperacao & ""
    strSQL = strSQL + "    AND NU_CNPJ ='" & cnpjParam & "'"
    strSQL = strSQL + "  ORDER BY DS_CONTA_CLIE"
    
    With rst
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .Open strSQL, cnn
    End With
    
    If rst.BOF = False Then
    
        If icTipoPlanoDados = "R" Then
        
            Worksheets("Cenario Receitas").Activate
            
            descricaoTipoProcessamento = "receitas"
            
            Range("G5:S10000").Select
            Selection.ClearContents
            Range("G5").Select
            
        Else
            If icTipoPlanoDados = "D" Then
            
                Worksheets("Cenario Despesas").Activate
                
                descricaoTipoProcessamento = "despesas"
                
                Range("G5:S10000").Select
                Selection.ClearContents
                Range("G5").Select
                
            Else
            
                Worksheets("Cenario ReceitasDespesas").Activate
                
                descricaoTipoProcessamento = "receitas/despesas"
                
                Range("G5:S10000").Select
                Selection.ClearContents
                Range("G5").Select
                
            End If
        End If
        
        If rst.EOF = False Then
            
            rst.MoveFirst
            
            rst.MoveLast
            qtdeRegistros = rst.RecordCount
            
            rst.MoveFirst
            
            linha = 5
            
            For n = 0 To qtdeRegistros - 1
                    
                Call barraProgresso("Processando a recuperação de dados do cenário de importação de " + descricaoTipoProcessamento + "...", linha - 5)
                    
                Range("G" + CStr(linha)).Value = rst(0).Value
                Range("H" + CStr(linha)).Value = rst(1).Value
                Range("I" + CStr(linha)).Value = rst(2).Value
                Range("J" + CStr(linha)).Value = rst(3).Value
                
                If Range("K5").Value = "" Then
                    Range("K5").Value = rst(4).Value
                    Range("L5").Value = rst(5).Value
                    Range("M5").Value = rst(6).Value
                    Range("N5").Value = rst(7).Value
                    Range("P5").Value = rst(8).Value
                    Range("Q5").Value = rst(9).Value
                    Range("R5").Value = rst(10).Value
                    Range("S5").Value = rst(11).Value
                End If
                
                Range("T" + CStr(linha)).Value = rst(12).Value
                
                linha = linha + 1
                
                rst.MoveNext
                
            Next n
        
        End If
        
    End If
    
    rst.Close

End Sub

Public Sub carregarDadosDePalavrasImportacao(icTipoPlanoDados As String)

Dim descricaoTipoProcessamento As String

Dim n As Integer
Dim qtdeRegistros As Integer

    strSQL = "SELECT "
    strSQL = strSQL + "   DS_PLVR_EXCD"
    strSQL = strSQL + "  ,IC_TIPO_TRANS_FLUXO_CAIXA"
    strSQL = strSQL + "  FROM T_LISTA_PLVR_EXCD"
       
    If icTipoPlanoDados = "N" Then
        strSQL = strSQL + "  WHERE (IC_TIPO_TRANS_FLUXO_CAIXA <> 'R' AND IC_TIPO_TRANS_FLUXO_CAIXA <> 'D')"
    Else
        strSQL = strSQL + "  WHERE IC_TIPO_TRANS_FLUXO_CAIXA = '" & icTipoPlanoDados & " '"
    End If
    
    strSQL = strSQL + "    AND NU_ANO_PLAN_ORIG_PROC = " & anoRecuperacao & ""
    strSQL = strSQL + "    AND NU_CNPJ ='" & cnpjParam & "'"
    strSQL = strSQL + "  ORDER BY DS_PLVR_EXCD"
    
    With rst
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .Open strSQL, cnn
    End With
    
    If rst.BOF = False Then
    
        If icTipoPlanoDados = "R" Then
            Worksheets("Cenario Receitas").Activate
            descricaoTipoProcessamento = "receitas"
        Else
            If icTipoPlanoDados = "D" Then
                Worksheets("Cenario Despesas").Activate
                descricaoTipoProcessamento = "despesas"
            Else
                Worksheets("Cenario ReceitasDespesas").Activate
                descricaoTipoProcessamento = "receitas/despesas"
            End If
        End If
        
        If rst.EOF = False Then
            
            rst.MoveFirst
            
            rst.MoveLast
            qtdeRegistros = rst.RecordCount
            
            rst.MoveFirst
            
            linha = 5
            
            For n = 0 To qtdeRegistros - 1
            
                Range("O" + CStr(linha)).Value = ""
                
                Call barraProgresso("Limpando cenário de palavras importação de " + descricaoTipoProcessamento + "...", linha - 5)
                
                linha = linha + 1
                
            Next n
            
            linha = 5
            
            Do While rst.EOF = False
                    
                Call barraProgresso("Processando a recuperação de palavras do cenário de importação de " + descricaoTipoProcessamento + "...", linha - 5)
                    
                Range("O" + CStr(linha)).Value = rst(0).Value
                
                linha = linha + 1
                
                rst.MoveNext
                
            Loop
        
        End If
        
    End If
    
    rst.Close

End Sub


Public Sub carregarDadosDeExportacao(icTipoExportacao As String)

Dim descricaoTipoProcessamento As String

Dim n As Integer
Dim qtdeRegistros As Integer

    strSQL = "SELECT CD_INSTT_FNCR"
    strSQL = strSQL + " ,DS_INSTT_FNCR"
    strSQL = strSQL + " ,CD_DCTO_RFRC_FLUXO_CAIXA"
    strSQL = strSQL + " ,DS_DCTO_RFRC_FLUXO_CAIXA"
    strSQL = strSQL + " ,NU_ANO_PLAN_ORIG_PROC"
    strSQL = strSQL + " ,NU_CNPJ"
    strSQL = strSQL + " ,TP_CNRO_EXPRT"
    strSQL = strSQL + " FROM T_CNRO_EXPRT_ARQV"
    strSQL = strSQL + " WHERE NU_ANO_PLAN_ORIG_PROC = " & anoRecuperacao & ""
    strSQL = strSQL + "   AND NU_CNPJ ='" & cnpjParam & "'"
    
    If icTipoExportacao <> "DOCREF" Then
        strSQL = strSQL + "   AND CD_INSTT_FNCR IS NOT NULL"
        strSQL = strSQL + "  ORDER BY CD_DCTO_RFRC_FLUXO_CAIXA"
    Else
        strSQL = strSQL + "   AND CD_INSTT_FNCR IS NULL"
        strSQL = strSQL + "  ORDER BY CD_INSTT_FNCR"
        
        Worksheets("Cenario de Exportacao").Activate
        
        Range("G5:J10000").Select
        Selection.ClearContents
        Range("G5").Select
        
    End If
    
    With rst
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .Open strSQL, cnn
    End With
    
    If rst.BOF = False Then
    
        If rst.EOF = False Then
        
            rst.MoveFirst
            
            rst.MoveLast
            qtdeRegistros = rst.RecordCount
            
            rst.MoveFirst
            
            linha = 5
            
            For n = 0 To qtdeRegistros - 1
                    
                Call barraProgresso("Processando cenário de dados de exportação...", linha - 5)
                
                If icTipoExportacao <> "DOCREF" Then
                    Range("G" + CStr(linha)).Value = rst(0).Value
                    Range("H" + CStr(linha)).Value = rst(1).Value
                Else
                    Range("I" + CStr(linha)).Value = rst(2).Value
                    Range("J" + CStr(linha)).Value = rst(3).Value
                End If
                
                linha = linha + 1
                
                rst.MoveNext
                
            Next n
        
        End If
        
    End If
    
    rst.Close

End Sub

Public Sub carregarDadosDoMes(planilha As String)
    
Dim descricaoTipoProcessamento As String

Dim n As Integer
Dim qtdeRegistros As Integer

    Range("C5:M10000").Select
    Selection.ClearContents
    Range("C5").Select

    strSQL = "SELECT DT_MVMT_FLUXO_CAIXA "
    strSQL = strSQL + "  ,DS_CLSSF_PLANO_CONTA "
    strSQL = strSQL + "  ,CD_DCTO_RFRC_FLUXO_CAIXA "
    strSQL = strSQL + "  ,DS_PLANO_CONTA "
    strSQL = strSQL + "  ,DS_INSTT_FNCR "
    strSQL = strSQL + "  ,VL_ENTR_FLUXO_CAIXA "
    strSQL = strSQL + "  ,VL_SAIDA_FLUXO_CAIXA "
    strSQL = strSQL + "  ,IC_STATUS_VALOR "
    strSQL = strSQL + "  ,DS_MES_PROC_RECB "
    strSQL = strSQL + " FROM T_FLUXO_CAIXA"
    strSQL = strSQL + " WHERE DS_PLAN_ORIG_PROC = '" & planilha & "'"
    strSQL = strSQL + " AND NU_ANO_PLAN_ORIG_PROC = " & anoRecuperacao & ""
    strSQL = strSQL + " AND NU_CNPJ = '" & cnpjParam & "'"
    
    With rst
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .Open strSQL, cnn
    End With
    
    If rst.BOF = False Then
    
        If rst.EOF = False Then
        
            rst.MoveFirst
            
            rst.MoveLast
            qtdeRegistros = rst.RecordCount
            
            rst.MoveFirst
            
            linha = 5
            
            For n = 0 To qtdeRegistros - 1
                    
                Call barraProgresso("Processando os dados do mês " + planilha + "...", linha - 5)
                
                Range("C" + CStr(linha)).Value = Mid(rst(0).Value, 1, 2)
                Range("E" + CStr(linha)).Value = rst(1).Value
                Range("F" + CStr(linha)).Value = rst(2).Value
                Range("G" + CStr(linha)).Value = rst(3).Value
                Range("H" + CStr(linha)).Value = rst(4).Value
                Range("J" + CStr(linha)).Value = rst(5).Value
                Range("K" + CStr(linha)).Value = rst(6).Value
                Range("L" + CStr(linha)).Value = rst(7).Value
                Range("I" + CStr(linha)).Value = rst(8).Value
                
                linha = linha + 1
                
                rst.MoveNext
                
            Next n
        
        End If
        
    End If
    
    rst.Close
    
End Sub
