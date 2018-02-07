Attribute VB_Name = "Exportacao"
Public dataInicialExportacao As String
Public dataFinalExportacao As String

Public Sub ExportarCSVSemFormato()
    
On Error GoTo Erro
    
    Dim myCSVFileName As String
    Dim myWB As Workbook
    Dim rngToSave As Range
    Dim fNum As Integer
    Dim csvVal As String
    Dim strIntervalo As String

    Set myWB = ThisWorkbook
    
    myCSVFileName = myWB.path & "\" & "FluxoCaixaSemFormato_Exportado" & VBA.Format(VBA.Now, "dd-MM-yyyy hh-mm") & ".csv"
    
    csvVal = ""
    
    fNum = FreeFile
    
    strIntervalo = InputBox("O intervalo de c�lulas para a exporta��o da planilha atual ser� (coluna C = Dia e N igual Saldo di�rio):", "Exporta��o de Dados para .CSV", Default:="C5:N10000")
    
    Set rngToSave = Range(strIntervalo)

    Open myCSVFileName For Output As #fNum

    For i = 1 To rngToSave.Rows.Count
        For j = 1 To rngToSave.Columns.Count
            csvVal = csvVal & Chr(34) & rngToSave(i, j).Value & Chr(34) & ";"
        Next
        csvVal = csvVal & Chr(34)
        Print #fNum, Left(csvVal, Len(csvVal) - 2)
        csvVal = ""
    Next

    Close #fNum
    
    MsgBox "Exporta��o realizada com sucesso. Nome do arquivo exportado: " & "FluxoCaixa_Exportado" & VBA.Format(VBA.Now, "dd-MM-yyyy hh-mm") & ".csv" & Chr(13) & Chr(13) & _
    " no diret�rio: " & myWB.path, vbOKOnly + vbInformation, "Exporta��o de Dados para .CSV"
    
    Exit Sub
    
Erro:

    MsgBox "Erro ao processar a exporta��o para .CSV. " + Err.Description + ". Tente exportar novamente em instantes.", vbOKOnly + vbInformation, "Erro ao Exportar"

End Sub

Public Sub ExportarDominio()
    
On Error GoTo Erro
    
    Dim myCSVFileName As String
    Dim myWB As Workbook
    Dim rngToSave As Range
    Dim fNum As Integer
    Dim csvVal As String
    Dim strIntervalo As String
    Dim nomePlanilhaAtual As String
    Dim valorLancamento As String
    
    Dim numeroLancamento As Integer
    Dim i As Integer
    
    nomePlanilhaAtual = ActiveSheet.Name

    Worksheets("Configura��es B�sicas").Activate
    ano = Range("E5").Value
    nomeClie = Range("E9").Value
    cnpjClie = Replace(Replace(Replace(Range("E8").Value, ".", ""), "-", ""), "/", "")
    
    Worksheets(nomePlanilhaAtual).Activate

    dataInicialFinal (nomePlanilhaAtual)

    Set myWB = ThisWorkbook
    
    ordernarPlanilhaLancamento (ActiveSheet.Name)
    
    If frmEscolhaSistemaExportacao.txtCodigoEmpresaExportacao = "" Then
        MsgBox "Digite o c�digo da empresa para exporta��o de dados para o Sistema Dom�nio.", vbOKOnly + vbInformation, "Erro ao Exportar"
        Exit Sub
    End If
    
    If frmEscolhaSistemaExportacao.txtLoginUsuarioExportacao = "" Then
        MsgBox "Digite o usu�rio no Sistema Dom�nio para exporta��o de dados.", vbOKOnly + vbInformation, "Erro ao Exportar"
        Exit Sub
    End If
    
    myCSVFileName = myWB.path & "\" & "FluxoCaixaDominio_Exportado" & VBA.Format(VBA.Now, "dd-MM-yyyy hh-mm") & ".txt"
    
    csvVal = ""
    
    fNum = FreeFile
    
    Open myCSVFileName For Output As #fNum
    
    strIntervalo = "C5:N10000"
    Set rngToSave = Range(strIntervalo)
    
    '-------Cabe�alho-------------------------------
    '0100000 - id de cabe�alho
    '11  N�mero da Empresa no Sistema Dominio
    '07165722000145  CNPJ
    '02/10/201803/10/2018 Per�odo dos lan�amentos
    'N0500000117 (Padr�o)
    '-------Cabe�alho-------------------------------
    csvVal = "0100000" + frmEscolhaSistemaExportacao.txtCodigoEmpresaExportacao + cnpjClie + dataInicialExportacao + dataFinalExportacao + "N0500000117"
    Print #fNum, csvVal
    
    numeroLancamento = 1
    
    frmEscolhaSistemaExportacao.frameProgressoExportacao.Visible = True
    
    For i = 1 To rngToSave.Rows.Count
    
        If rngToSave(i, 2).Value <> "" Then
    
            '-------Para cada lan�amento, o registro abaixo � lan�ado para o respectivo usu�rio----
            '02000000 - id Fixo ao identificar a linha usuario
            '1 - numero do lancamento
            'X - ver no layout
            '02/10/2018 - data do lancamento
            'CRISTIABEL -usuario
            '-------Para cada lan�amento, o registro abaixo � lan�ado para o respectivo usu�rio----
            csvVal = "02" + Format(CStr(numeroLancamento), "0000000") + "x" + rngToSave(i, 2).Value + frmEscolhaSistemaExportacao.txtLoginUsuarioExportacao
            Print #fNum, csvVal
            
            numeroLancamento = numeroLancamento + 1
            
            '-------Para cada lan�amento, o registro abaixo � lan�ado----
            '03 - id Fixo ao identificar a linha de lancamento
            '00000002 - numero do lancamento
            '1120600 - conta devedora
            '1120200 - conta credora (institui��o financeira)
            '0000000 - ver no layout
            '116600 - valor (ver no layout)
            '0000001 - historico padrao
            'teste - historico variavel
            '-------Para cada lan�amento, o registro abaixo � lan�ado----
            
            If rngToSave(8, j).Value = "" Then
                valorLancamento = CStr(Replace(rngToSave(i, 9).Value, ",", ""))
            Else
                valorLancamento = CStr(Replace(rngToSave(i, 8).Value, ",", ""))
            End If
            
            csvVal = "03" + Format(CStr(numeroLancamento), "0000000") _
                    + Replace(Format(localizaContaDevedora(rngToSave(i, 5).Value), "00000"), ".", "") _
                    + Format(localizaContaCredora(rngToSave(i, 6).Value), "00000") _
                    + Format(valorLancamento, "0000000000000") _
                    + "0000001" _
                    + Trim(Format(rngToSave(i, 4).Value, "@@@@@@@@@@"))
            Print #fNum, csvVal
            
            numeroLancamento = numeroLancamento + 1
            
            Call frmEscolhaSistemaExportacao.barraProgresso("Exportando dados no formato do Sistema Dom�nio ", numeroLancamento)
            
         Else
            
            csvVal = "0000000"
            Print #fNum, csvVal
            
            csvVal = "9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999"
            Print #fNum, csvVal
            
            Exit For
        
        End If
        
    Next

    Close #fNum
    
    frmEscolhaSistemaExportacao.frameProgressoExportacao.Visible = False
    
    MsgBox "Exporta��o de dados para o Sistema Dom�nio realizada com sucesso. Nome do arquivo exportado: " & "FluxoCaixaDominio_Exportado" & VBA.Format(VBA.Now, "dd-MM-yyyy hh-mm") & ".csv" & Chr(13) & Chr(13) & _
    " no diret�rio: " & myWB.path, vbOKOnly + vbInformation, "Exporta��o de Dados"
    
    Exit Sub
    
Erro:

    MsgBox "Erro ao processar a exporta��o para .txt. " + Err.Description + ". Tente exportar novamente em instantes.", vbOKOnly + vbInformation, "Erro ao Exportar"

End Sub

Public Sub ExportarProsoft()
    
On Error GoTo Erro
    
    Dim myCSVFileName As String
    Dim myWB As Workbook
    Dim rngToSave As Range
    Dim fNum As Integer
    Dim csvVal As String
    Dim strIntervalo As String
    Dim nomePlanilhaAtual As String
    Dim valorLancamento As String
    
    Dim numeroLancamento As Integer
    
    nomePlanilhaAtual = ActiveSheet.Name

    Worksheets(nomePlanilhaAtual).Activate

    Set myWB = ThisWorkbook
    
    ordernarPlanilhaLancamento (ActiveSheet.Name)
    
    myCSVFileName = myWB.path & "\" & "FluxoCaixaProsoft_Exportado" & VBA.Format(VBA.Now, "dd-MM-yyyy hh-mm") & ".txt"
    
    csvVal = ""
    
    fNum = FreeFile
    
    Open myCSVFileName For Output As #fNum
    
    strIntervalo = "C5:N10000"
    Set rngToSave = Range(strIntervalo)
    
    numeroLancamento = 1
    
    frmEscolhaSistemaExportacao.frameProgressoExportacao.Visible = True
    
    For i = 1 To rngToSave.Rows.Count
    
        If rngToSave(i, 2).Value <> "" Then
     
            '-------Para cada lan�amento, o registro abaixo � lan�ado----
            'LC100001 - Numero do lan�amento
            '102102018 -  1+data do lan�amento no fluxo de caixa
            '11206 - conta devedora (do plano de contas)
            '11202 - conta credora (instituicao financeira)
            '0000000001166.00 - valor
            'Histotico padrao - (vazio)
            'teste - historico variavel (doc/ref)
            '-------Para cada lan�amento, o registro abaixo � lan�ado----
            
            If rngToSave(8, j).Value = "" Then
                valorLancamento = CStr(Replace(rngToSave(i, 9).Value, ",", "."))
            Else
                valorLancamento = CStr(Replace(rngToSave(i, 8).Value, ",", "."))
            End If
            
            csvVal = "LC" + Format(CStr(numeroLancamento), "000000") _
                    + "   " _
                    + "                                                " _
                    + "1" + Replace(rngToSave(i, 2).Value, "/", "") _
                    + Replace(Format(localizaContaDevedora(rngToSave(i, 5).Value, rngToSave(i, 3).Value), "00000"), ".", "") _
                    + "                   " _
                    + Format(localizaContaCredora(rngToSave(i, 6).Value), "00000") _
                    + "                   " _
                    + Format(valorLancamento, "0000000000000000") _
                    + "1  " _
                    + Trim(Format(rngToSave(i, 4).Value, "@@@@@@@@@@"))
            Print #fNum, csvVal
            
            numeroLancamento = numeroLancamento + 1
            
            Call frmEscolhaSistemaExportacao.barraProgresso("Exportando dados no formato do Sistema Prosoft ", numeroLancamento)
            
        Else
        
            Exit For
        
        End If
        
    Next

    Close #fNum
    
    MsgBox "Exporta��o de dados para o Sistema Prosoft realizada com sucesso. Nome do arquivo exportado: " & "FluxoCaixaProsoft_Exportado" & VBA.Format(VBA.Now, "dd-MM-yyyy hh-mm") & ".csv" & Chr(13) & Chr(13) & _
    " no diret�rio: " & myWB.path, vbOKOnly + vbInformation, "Exporta��o de Dados"
    
    Exit Sub
    
Erro:

    MsgBox "Erro ao processar a exporta��o para .txt. " + Err.Description + ". Tente exportar novamente em instantes.", vbOKOnly + vbInformation, "Erro ao Exportar"

End Sub


Public Sub dataInicialFinal(nomePlanilha As String)
    
Dim linha_panilha As Integer
    
    linha_planilha = 5
    
    dataInicialExportacao = Range("D" + CStr(linha_planilha)).Value
    
    Do While Range("C" + CStr(linha_planilha)).Value <> ""
        
        linha_planilha = linha_planilha + 1
        
    Loop
    
    dataFinalExportacao = Range("D" + CStr(linha_planilha - 1)).Value
    
End Sub

Public Function localizaContaDevedora(descricaoContaDevedora As String)

Dim linha_planilha As Integer
Dim contaDevedora As String

    nomePlanilha = ActiveSheet.Name
    
    Worksheets("Cenario de Exportacao").Activate
    
    linha_planilha = 5
    
    Do While Range("K" + CStr(linha_planilha)).Value <> ""
    
        If descricaoContaDevedora = Range("L" + CStr(linha_planilha)).Value Then
        
            contaDevedora = Range("K" + CStr(linha_planilha)).Value
            Exit Do
            
        End If
        
        linha_planilha = linha_planilha + 1
        
        Call frmEscolhaSistemaExportacao.barraProgresso("Localizando Conta Devedora ", linha_planilha)
        
    Loop
    
    localizaContaDevedora = contaDevedora
    
    Worksheets(nomePlanilha).Activate

'Dim linha_planilha As Integer
'Dim linhaplanoConta As Integer
'Dim nomePlanilha As String
'Dim colCodigoPlanoConta As String
'Dim colDescricaoPlanoConta As String
'Dim contaDevedoraLocalizada As String

'Dim achouContaDevedora As Boolean
 
'    nomePlanilha = ActiveSheet.Name
    
'    Worksheets("Configura��es B�sicas").Activate
    
'    linha_planilha = 12
'    achouContaDevedora = False
'
'    Do While Range("E" + CStr(linha_planilha)).Value <> ""
'
'        If descricaoClassificacaoPlanoConta = Range("E" + CStr(linha_planilha)).Value Then
'
'            colCodigoPlanoConta = Range("G" + CStr(linha_planilha)).Value
'            colDescricaoPlanoConta = Range("H" + CStr(linha_planilha)).Value
'
'            If Range("F" + CStr(linha_planilha)).Value = "R" Then
'                Worksheets("PC Receitas").Activate
'            Else
'                Worksheets("PC Despesas").Activate
'            End If
'
'            linhaplanoConta = 5
'
'            Do While Range(colCodigoPlanoConta + CStr(linhaplanoConta)).Value <> ""
'
'                If Range(colDescricaoPlanoConta + CStr(linhaplanoConta)).Value = descricaoContaDevedora Then
'
'                    contaDevedoraLocalizada = Range(colCodigoPlanoConta + CStr(linhaplanoConta)).Value
'                    achouContaDevedora = True
'                    Exit Do
'
'                End If
'
'                linhaplanoConta = linhaplanoConta + 1
'
'                Call frmEscolhaSistemaExportacao.barraProgresso("Localizando Conta Devedora ", linhaplanoConta)
'
'            Loop
'
'        End If
'
'        If achouContaDevedora = True Then Exit Do
'
'        Worksheets("Configura��es B�sicas").Activate
'
'        linha_planilha = linha_planilha + 1
'
'    Loop
    
'    localizaContaDevedora = contaDevedoraLocalizada
    
'    Worksheets(nomePlanilha).Activate

End Function

Public Function localizaContaCredora(descricaoContaCredora As String)

Dim linha_planilha As Integer
Dim contaCredora As String

    nomePlanilha = ActiveSheet.Name
    
    Worksheets("Cenario de Exportacao").Activate
    
    linha_planilha = 5
    
    achouContaDevedora = False
    
    Do While Range("H" + CStr(linha_planilha)).Value <> ""
    
        If descricaoContaCredora = Range("H" + CStr(linha_planilha)).Value Then
        
            contaCredora = Range("G" + CStr(linha_planilha)).Value
            Exit Do
            
        End If
        
        linha_planilha = linha_planilha + 1
        
        Call frmEscolhaSistemaExportacao.barraProgresso("Localizando Conta Credora ", linha_planilha)
        
    Loop
    
    localizaContaCredora = contaCredora

End Function


Public Sub ExportarAlterdata()
    
On Error GoTo Erro
    
    Dim myCSVFileName As String
    Dim myWB As Workbook
    Dim rngToSave As Range
    Dim fNum As Integer
    Dim csvVal As String
    Dim strIntervalo As String
    Dim nomePlanilhaAtual As String
    Dim valorLancamento As String
    
    Dim numeroLancamento As Integer
    
    nomePlanilhaAtual = ActiveSheet.Name

    Worksheets(nomePlanilhaAtual).Activate

    Set myWB = ThisWorkbook
    
    ordernarPlanilhaLancamento (ActiveSheet.Name)
    
    myCSVFileName = myWB.path & "\" & "FluxoCaixaAlterdata_Exportado" & VBA.Format(VBA.Now, "dd-MM-yyyy hh-mm") & ".txt"
    
    csvVal = ""
    
    fNum = FreeFile
    
    Open myCSVFileName For Output As #fNum
    
    strIntervalo = "C5:N10000"
    Set rngToSave = Range(strIntervalo)
    
    numeroLancamento = 1
    
    frmEscolhaSistemaExportacao.frameProgressoExportacao.Visible = True
    
    For i = 1 To rngToSave.Rows.Count
    
        If rngToSave(i, 2).Value <> "" Then
     
            '-------Para cada lan�amento, o registro abaixo � lan�ado----
            'Coluna 1
            'C�digo do Lan�amento Autom�tico
             
            'Coluna 2
            'Conta D�bito
             
            'Coluna 3
            'Conta Cr�dito
             
            'Coluna 4
            'Data
             
            'Coluna 5
            'valor
             
            'Coluna 6
            'C�digo do Hist�rico
             
            'Coluna 7
            'Complemento Hist�rico
             
            'Coluna 8
            'Centro de custo D�bito
             
            'Coluna 9
            'Centro de custo Cr�dito
             
            'Coluna 10
            'N�mero de Documento
            '-------Para cada lan�amento, o registro abaixo � lan�ado----
            
            If rngToSave(8, j).Value = "" Then
                valorLancamento = CStr(Replace(rngToSave(i, 9).Value, ",", "."))
            Else
                valorLancamento = CStr(Replace(rngToSave(i, 8).Value, ",", "."))
            End If
            
            csvVal = CStr(numeroLancamento) + ";" & _
                    localizaContaDevedora(rngToSave(i, 5).Value, rngToSave(i, 3).Value) + ";" & _
                    localizaContaCredora(rngToSave(i, 6).Value) + ";" & _
                    Format(rngToSave(i, 2).Value, "dd/mm/yyyy") + ";" & _
                    valorLancamento + ";" & _
                    "" + ";" & _
                    "" + ";" & _
                    "" + ";" & _
                    "" + ";" & _
                    CStr(rngToSave(i, 4).Value) + ";"
            Print #fNum, csvVal
            
            numeroLancamento = numeroLancamento + 1
            
            Call frmEscolhaSistemaExportacao.barraProgresso("Exportando dados no formato do Sistema Prosoft ", numeroLancamento)
            
        Else
        
            Exit For
        
        End If
        
    Next

    Close #fNum
    
    MsgBox "Exporta��o de dados para o Sistema Alterdata realizada com sucesso. Nome do arquivo exportado: " & "FluxoCaixaProsoft_Exportado" & VBA.Format(VBA.Now, "dd-MM-yyyy hh-mm") & ".csv" & Chr(13) & Chr(13) & _
    " no diret�rio: " & myWB.path, vbOKOnly + vbInformation, "Exporta��o de Dados"
    
    Exit Sub
    
Erro:

    MsgBox "Erro ao processar a exporta��o para .txt. " + Err.Description + ". Tente exportar novamente em instantes.", vbOKOnly + vbInformation, "Erro ao Exportar"
    Close #fNum

End Sub


