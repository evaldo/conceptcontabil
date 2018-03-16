VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmConsultaClassificacao 
   Caption         =   "Consulta Classificação"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7005
   OleObjectBlob   =   "frmConsultaClassificacao.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmConsultaClassificacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim descricaoClassificacao(1 To 100, 1 To 3) As String

Private Sub cmbListaDescricaoClassificacao_Click()

Dim linha As Integer
Dim linhaListaClassificacao As Integer

Dim arrClassificacaoPlanoContas(1 To 1000, 1 To 2) As String
    
    Me.lstClassificacao.Clear
   
    mes_processamento = ActiveSheet.Name
    
    linha = 5
    linhaListaClassificacao = 1
    
    arrClassificacaoPlanoContas(linhaListaClassificacao, 1) = "Código"
    arrClassificacaoPlanoContas(linhaListaClassificacao, 2) = "Descrição do Plano de Contas"
    
    linhaListaClassificacao = linhaListaClassificacao + 1
    
    If frmImportarPlanilhaComParametro.optClassificacaoReceita = True Then
                
        Worksheets("PC Receitas").Activate
    
    Else
    
        Worksheets("PC Despesas").Activate
    
    End If
    
    Range("D5").Select
    
    Do While (Range(descricaoClassificacao(cmbListaDescricaoClassificacao.ListIndex + 1, 2) + CStr(linha)).Value <> "" And Range(descricaoClassificacao(cmbListaDescricaoClassificacao.ListIndex + 1, 2) + CStr(linha)).Value <> "-")
                    
        arrClassificacaoPlanoContas(linhaListaClassificacao, 2) = Range(descricaoClassificacao(cmbListaDescricaoClassificacao.ListIndex + 1, 2) + CStr(linha)).Text
        arrClassificacaoPlanoContas(linhaListaClassificacao, 1) = Range(descricaoClassificacao(cmbListaDescricaoClassificacao.ListIndex + 1, 3) + CStr(linha)).Text
        
        linhaListaClassificacao = linhaListaClassificacao + 1
        
        linha = linha + 1
                       
    Loop
    
    Me.lstClassificacao.List = arrClassificacaoPlanoContas
    
    Worksheets(mes_processamento).Activate


End Sub

Private Sub cmdConfirma_Click()
    
Dim i As Integer

     For i = 0 To Me.lstClassificacao.ListCount - 1
        If lstClassificacao.Selected(i) = True Then
            codigoPlanoContas = lstClassificacao.List(i, 0)
            descricaoPlanoContas = lstClassificacao.List(i, 1)
        End If
    Next i
    
    frmImportarPlanilhaComParametro.cmbListaDescricaoClassificacao.Text = Me.cmbListaDescricaoClassificacao.Text
    frmImportarPlanilhaComParametro.cmbClassificacao.Text = codigoPlanoContas
    frmImportarPlanilhaComParametro.txtDescricaoClassificacao.Text = descricaoPlanoContas
    
    Unload Me
    
End Sub

Private Sub cmdFechar_Click()

    Unload Me

End Sub


Private Sub UserForm_Activate()

    If frmImportarPlanilhaComParametro.optClassificacaoDespesa.Value = True Then
        Call classificacaoDespesa
    Else
        Call classificacaoReceita
    End If

End Sub


Public Sub classificacaoDespesa()

Dim i As Integer
Dim linha As Integer
Dim contador_classifcacao As Integer

Dim planoContas(1 To 100, 1 To 5) As String
Dim mes_processamento As String

    For linha = 1 To 20
                    
        descricaoClassificacao(linha, 2) = ""
        descricaoClassificacao(linha, 1) = ""
        descricaoClassificacao(linha, 3) = ""
        
    Next linha
 
    mes_processamento = ActiveSheet.Name
    Worksheets("Configurações Básicas").Activate
    
    linha = 12
    i = 1
    
    Do While Range("D" + CStr(linha)).Value <> ""
        
        planoContas(i, 1) = Range("D" + CStr(linha)).Value 'codigo plano de contas
        planoContas(i, 2) = Range("E" + CStr(linha)).Value 'descrição do plano de contas
        planoContas(i, 3) = Range("F" + CStr(linha)).Value 'tipo de operação do plano de contas (R/D)
        planoContas(i, 4) = Range("G" + CStr(linha)).Value 'coluna do código do plano de contas
        planoContas(i, 5) = Range("H" + CStr(linha)).Value 'coluna da descrição do plano de contas
        
        linha = linha + 1
        i = i + 1
        
    Loop
    
    Worksheets(mes_processamento).Activate
    
    i = 1
    contador_classifcacao = 1
    
    For i = 1 To 100
        
        If planoContas(i, 3) = "D" Then
            
            descricaoClassificacao(contador_classifcacao, 2) = planoContas(i, 5)
            descricaoClassificacao(contador_classifcacao, 1) = planoContas(i, 2)
            descricaoClassificacao(contador_classifcacao, 3) = planoContas(i, 4)
            
            contador_classifcacao = contador_classifcacao + 1
            
        End If
        
    Next i
   
    cmbListaDescricaoClassificacao.Clear
    cmbListaDescricaoClassificacao.List = descricaoClassificacao

End Sub

Public Sub classificacaoReceita()
 
Dim i As Integer
Dim linha As Integer
Dim contador_classifcacao As Integer

Dim planoContas(1 To 100, 1 To 5) As String
Dim mes_processamento As String

    For linha = 1 To 20
                    
        descricaoClassificacao(linha, 2) = ""
        descricaoClassificacao(linha, 1) = ""
        descricaoClassificacao(linha, 3) = ""
        
    Next linha
 
    mes_processamento = ActiveSheet.Name
    Worksheets("Configurações Básicas").Activate
    
    linha = 12
    i = 1
    
    Do While Range("D" + CStr(linha)).Value <> ""
        
        planoContas(i, 1) = Range("D" + CStr(linha)).Value 'codigo plano de contas
        planoContas(i, 2) = Range("E" + CStr(linha)).Value 'descrição do plano de contas
        planoContas(i, 3) = Range("F" + CStr(linha)).Value 'tipo de operação do plano de contas (R/D)
        planoContas(i, 4) = Range("G" + CStr(linha)).Value 'coluna do código do plano de contas
        planoContas(i, 5) = Range("H" + CStr(linha)).Value 'coluna da descrição do plano de contas
        
        linha = linha + 1
        i = i + 1
        
    Loop
    
    Worksheets(mes_processamento).Activate
    
    i = 1
    contador_classifcacao = 1
    
    For i = 1 To 100
        
        If planoContas(i, 3) = "R" Then
            
            descricaoClassificacao(contador_classifcacao, 2) = planoContas(i, 5)
            descricaoClassificacao(contador_classifcacao, 1) = planoContas(i, 2)
            descricaoClassificacao(contador_classifcacao, 3) = planoContas(i, 4)
            
            contador_classifcacao = contador_classifcacao + 1
            
        End If
        
    Next i
   
    cmbListaDescricaoClassificacao.Clear
    cmbListaDescricaoClassificacao.List = descricaoClassificacao
   
End Sub
