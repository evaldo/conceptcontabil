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
    
    If frmEscolhaDesRec.bolClassificacaoReceita = True Then
                
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

    If frmEscolhaDesRec.bolClassificacaoDespesa = True Then
        Call classificacaoDespesa
    Else
        Call classificacaoReceita
    End If

End Sub


Public Sub classificacaoDespesa()

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

Public Sub classificacaoReceita()
 
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
