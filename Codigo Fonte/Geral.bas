Attribute VB_Name = "Geral"
Public anoAtual As String
Public nomeClieAtual As String
Public cnpjClieAtual As String
Public codigoPlanoContas As String
Public descricaoPlanoContas As String
Public codigoReduzidoPlanoContas As String
Public classificacaoContabil As String
Public nomePlanoConta As String

Public linhaInicialLeituraPlanilhaPlanoContas As Integer
Public manterDadosAposLogin As Boolean
Public loginAcesso As Boolean
Public usuario As String
Public senha As String

Public Sub Desfazer()
    
    Application.SendKeys ("^(z)")
    
End Sub

Public Sub Refazer()
    
    Application.SendKeys ("^(y)")
    
    
End Sub

Public Sub Salvar()
    
    ThisWorkbook.Save
    
End Sub

Public Sub FecharArquivo()

    ActiveWorkbook.Close
    ActiveWorkbook.Save

End Sub

Public Sub AbrirFormExportacao()

    frmEscolhaSistemaExportacao.Show (1)

End Sub

Public Sub ordernarPlanilhaLancamento(nomePlanilha As String)

    Range("C4:N10000").Select
    ActiveWorkbook.Worksheets(nomePlanilha).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(nomePlanilha).Sort.SortFields.Add Key:=Range("C5:C10000"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(nomePlanilha).Sort
        .SetRange Range("C4:N10000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("C5").Select

End Sub
