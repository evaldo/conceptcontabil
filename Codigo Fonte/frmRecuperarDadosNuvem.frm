VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRecuperarDadosNuvem 
   Caption         =   "Recuperar Dados da Nuvem"
   ClientHeight    =   5460
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6255
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

Private Sub cmdFechar_Click()

    Unload Me

End Sub

Private Sub cmdRecuperarDados_Click()

'On Error GoTo Erro

Dim linha As Integer

Dim colunaCodigoPlanoContaPlanilha As String
Dim colunaDescricaoPlanoContaPlanilha As String
Dim planoConta As String
Dim icTipoPlanoConta As String
Dim strSQLQtdePlanoConta As String

Dim bolExisteQtdePlanoConta As Boolean

    frmProgresso.Visible = True

    cnn.ConnectionString = "Driver={ODBC Driver 13 for SQL Server};Server=tcp:contarcondb.cmxd2lqddzlw.sa-east-1.rds.amazonaws.com,1433;Database=fluxocaixa;Uid=evaldo;Pwd={Gcas1302};Encrypt=yes;TrustServerCertificate=no;Connection Timeout=300;"
    cnn.Open
    
    nomePlanilha = ActiveSheet.Name
    
    bolExisteQtdePlanoConta = False
    
    Worksheets("Configurações Básicas").Activate
    
    ano = Range("E5").Value
    nomeClie = Range("E9").Value
    cnpjClie = Range("E8").Value
    
    Call barraProgresso("Processando a recuperação de dados...", 1)
    
    Application.ScreenUpdating = False
    
    '----------------------------------------------------------------------------------------------
    'Recuperação do Plano de Contas
    '----------------------------------------------------------------------------------------------
    cnpjParam = cnpjClie
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
    
    rst.Open (strSQL), cnn
    
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
                
                'strSQLQtdePlanoConta = "SELECT COUNT(1) "
                'strSQLQtdePlanoConta = strSQLQtdePlanoConta + "FROM T_CLSSF_PLANO_CONTA "
                'strSQLQtdePlanoConta = strSQLQtdePlanoConta + "WHERE CD_CLSSF_PLANO_CONTA = '" & rst(0).Value & "'"
                
                'rstQtdePlanoConta.Open (strSQLQtdePlanoConta), cnn
    
                'If rstQtdePlanoConta.EOF = False Then
                    
                '    If rstQtdePlanoConta(0).Value <= 1 Then
                '        rst.MoveNext
                '        bolExisteQtdePlanoConta = True
                '    End If
                    
                'End If
                
                'rstQtdePlanoConta.Close
                
             End If
            
             'If bolExisteQtdePlanoConta = False Then
             
             '   If planoConta = rst(0).Value Then
             If rst.EOF = False Then
                Range(colunaCodigoPlanoContaPlanilha + CStr(linha)) = rst(4).Value
                Range(colunaDescricaoPlanoContaPlanilha + CStr(linha)) = rst(5).Value
                
                rst.MoveNext
                    
             End If
                
             linha = linha + 1
                
             'End If
            
             bolExisteQtdePlanoConta = False
            
        Loop
        
    End If
    
    rst.Close
    
    Application.ScreenUpdating = True
    frmProgresso.Visible = False
    
    cnn.Close
    
    Worksheets(nomePlanilha).Activate
    
    MsgBox "Recuperação realizada com sucesso!!", vbOKOnly, "Recuperar dados da nuvem"
    Unload Me
    
    Exit Sub
    
'Erro:
    
'    MsgBox "Erro ao recuperar os dados selecionados. Refaça a operação.", vbOKOnly, "Recuperar dados da nuvem"
'    frmProgresso.Visible = False
    
'    Exit Sub

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

    strSQL = "SELECT CD_CLSSF_PLANO_CONTA"
    strSQL = strSQL + "  ,NU_CNPJ"
    strSQL = strSQL + "  ,IC_TIPO_TRANS_FLUXO_CAIXA"
    strSQL = strSQL + "  ,DS_CLSSF_PLANO_CONTA"
    strSQL = strSQL + "  ,CD_PLANO_CONTA"
    strSQL = strSQL + "  ,DS_PLANO_CONTA"
    strSQL = strSQL + "  ,CD_CLUN_CDGO_CLSSF_PLANO_CONTA"
    strSQL = strSQL + "  ,CD_CLUN_DSCR_PLANO_CONTA "
    strSQL = strSQL + "FROM T_CLSSF_PLANO_CONTA "
    strSQL = strSQL + "WHERE NU_CNPJ = '" & cnpjParam & "'"
    strSQL = strSQL + "  AND IC_TIPO_TRANS_FLUXO_CAIXA = '" & icTipoPlanoDados & "'"
    strSQL = strSQL + "  AND CD_CLUN_CDGO_CLSSF_PLANO_CONTA IS NOT NULL"
    strSQL = strSQL + "  ORDER BY 1"
    
    rst.Open (strSQL), cnn
     
    If rst.BOF = False Then
        
        If icTipoPlanoDados = "D" Then
            Worksheets("PC Despesas").Activate
        Else
            Worksheets("PC Receitas").Activate
        End If
        
        linha = 4
        
        If rst.EOF = False Then rst.MoveFirst
        Do While rst.EOF = False
            Do While Range(rst(6).Value + CStr(linha)) <> ""
            
                Call barraProgresso("Apagando dados do Plano de Contas de " + IIf(icTipoPlanoDados = "D", "Despesas", "Receitas") + "", linha)
                        
                Range(rst(6).Value + CStr(linha)) = ""
                Range(rst(7).Value + CStr(linha)) = ""
                
                linha = linha + 1
                
            Loop
            
            linha = 4
            rst.MoveNext
            
        Loop
        
    End If
    
    rst.Close

End Sub

