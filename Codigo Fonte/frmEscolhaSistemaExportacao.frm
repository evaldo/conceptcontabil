VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEscolhaSistemaExportacao 
   Caption         =   "Vinculação de Dados para Exportação dos Dados em Sistemas Contábeis"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7125
   OleObjectBlob   =   "frmEscolhaSistemaExportacao.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEscolhaSistemaExportacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arrayPlanoConta(1 To 5000, 1 To 3) As String
Dim caracterImportacaoArquivoTexto As String
Dim nomePlanilha As String

Public indiceArrayListaInstFinanc As Integer
Public bolAchouInstFinancLista As Boolean
Public indiceAtualizarInstFinanc As Integer
Public indicePlanoConta As Integer
Dim linha As Integer

Public indiceArrayListaDocRef As Integer
Public bolAchoudocRef As Boolean
Public indiceAtualizarDocRef As Integer

Public indiceArrayPlanoConta As Integer
Public bolAchouPlanoConta As Boolean
Public indiceAtualizarPlanoConta As Integer


Private Sub btnGerarExportacao_Click()

Dim indice As Integer
Dim linha_planilha As Integer

    If ValidaPlanilhaProcessamento() = False Then
        MsgBox "Escolha um planilha para lançamento do Fluxo de Caixa entre Jan e Dez.", vbOKOnly + vbInformation, "Salvar Dados"
        Exit Sub
    End If
    
    nomePlanilha = ActiveSheet.Name
    
    Worksheets("Cenario de Exportacao").Activate
    
    '-----------------------------------------------------------------------------------------------------------
    'Atualização do Cenário de Exportação
    '-----------------------------------------------------------------------------------------------------------
    
    Me.frameProgressoExportacao.Visible = True
    
    Call barraProgresso("Atualizando os dados de Intituição Financeira e Documento de Referência ", 1)
    
    Application.ScreenUpdating = False
    
    linha_planilha = 5
    indice = 1
    
    Do While Range("H" + CStr(linha_planilha)).Value <> ""
    
        Range("G" + CStr(linha_planilha)).Value = ""
        Range("H" + CStr(linha_planilha)).Value = ""
        Range("I" + CStr(linha_planilha)).Value = ""
        Range("H" + CStr(linha_planilha)).Value = ""
        
        indice = indice + 1
        
        linha_planilha = linha_planilha + 1
        
        Call barraProgresso("Eliminando dados de Intituição Financeira e Documento de Referência ", indice)
        
    Loop
    
    linha_planilha = 5
    
    For indice = 0 To Me.lstInstFinancCodigo.ListCount
        
        If Me.lstInstFinancCodigo.List(indice, 1) = "" Then
            Exit For
        End If
        
        Range("H" + CStr(linha_planilha)).Value = Me.lstInstFinancCodigo.List(indice, 0)
        Range("G" + CStr(linha_planilha)).Value = Me.lstInstFinancCodigo.List(indice, 1)
        
        linha_planilha = linha_planilha + 1
        
        Call barraProgresso("Gravando dados de Intituição Financeira ", linha_planilha)
        
    Next
    
    linha_planilha = 5
    
    For indice = 0 To Me.lstDocRefCodigo.ListCount
        
        If Me.lstDocRefCodigo.List(indice, 1) = "" Then Exit For
        
        Range("J" + CStr(linha_planilha)).Value = Me.lstDocRefCodigo.List(indice, 0)
        Range("I" + CStr(linha_planilha)).Value = Me.lstDocRefCodigo.List(indice, 1)
        
        linha_planilha = linha_planilha + 1
        
        Call barraProgresso("Gravando dados de Documento de Referência ", linha_planilha)
        
    Next
    
    linha_planilha = 5
    
    For indice = 0 To Me.lstPlanoContaCodigo.ListCount
        
        If Me.lstPlanoContaCodigo.List(indice, 1) = "" Then Exit For
        
        Range("L" + CStr(linha_planilha)).Value = Me.lstPlanoContaCodigo.List(indice, 0)
        Range("K" + CStr(linha_planilha)).Value = Me.lstPlanoContaCodigo.List(indice, 1)
        
        linha_planilha = linha_planilha + 1
        
        Call barraProgresso("Gravando dados de Plano de Contas ", linha_planilha)
        
    Next
    
    Range("H5").Select
    
    Worksheets(nomePlanilha).Activate
    
    Select Case cmbSistemaExportacao.Text
    
        Case "Sem Formato"
            Call ExportarCSVSemFormato
            
        Case "Dominio"
            Call ExportarDominio
            
        Case "Prosoft"
            Call ExportarProsoft
            
        Case "Alterdata"
            Call ExportarAlterdata
    
    End Select
    
    Me.frameProgressoExportacao.Visible = False
    Application.ScreenUpdating = True
    
    Worksheets(nomePlanilha).Activate

End Sub

Private Sub cmbDocRefCodigo_Change()

End Sub

Private Sub cmbDocRefCodigo_Click()
    
    Me.txtCodigoRef.Text = Me.cmbDocRefCodigo.Text
    
End Sub

Private Sub cmbPlanoContaImportado_Click()

indicePlanoConta = 1

    If cmbPlanoContaImportado.Text <> "" Then

        Do While indicePlanoConta <= UBound(arrayPlanoConta)
            
            If cmbPlanoContaImportado.Text = arrayPlanoConta(indicePlanoConta, 3) Then
                
                txtPlanoContaImportado.Text = arrayPlanoConta(indicePlanoConta, 1)
                Exit Sub
                
            End If
            
            indicePlanoConta = indicePlanoConta + 1
                
        Loop
        
    End If
    
End Sub

Private Sub cmbSistemaExportacao_Click()
    
    If cmbSistemaExportacao.Text = "Dominio" Then
    
        txtCodigoEmpresaExportacao.Enabled = True
        txtLoginUsuarioExportacao.Enabled = True
        
        txtCodigoEmpresaExportacao.BackColor = 16777215
        txtLoginUsuarioExportacao.BackColor = 16777215
        
        txtCodigoEmpresaExportacao.SetFocus
        
    Else
    
        txtCodigoEmpresaExportacao.Enabled = False
        txtLoginUsuarioExportacao.Enabled = False
        
        txtCodigoEmpresaExportacao.BackColor = 12632256
        txtLoginUsuarioExportacao.BackColor = 12632256
    
    End If
    
End Sub

Private Sub cmdAtualizarDocRefCodigo_Click()

Dim arrayDocRef(1 To 10000, 1 To 2) As String

    bolAchoudocRef = False

    indiceArrayListaDocRef = 0
    indiceAtualizarDocRef = 0

    If cmbDocRefCodigo.Text = "" Or txtCodigoRef.Text = "" Then
    
        MsgBox "Campos Documento de Referência e/ou Código do Documento de Referência não podem estar vazios.", vbOKOnly + vbInformation, "Exportação de Dados"
        Exit Sub
        
    End If

    For indiceArrayListaDocRef = 0 To Me.lstDocRefCodigo.ListCount - 1
    
        If Me.cmbDocRefCodigo.Text = lstDocRefCodigo.List(indiceArrayListaDocRef, 0) Then
            lstDocRefCodigo.List(indiceArrayListaDocRef, 1) = Me.txtCodigoRef.Text
            
            bolAchoudocRef = True
            
        End If
        
        If lstDocRefCodigo.List(indiceArrayListaDocRef, 0) <> "" Then
            indiceAtualizarDocRef = indiceAtualizarDocRef + 1
        End If
        
    Next
    
    
    If bolAchoudocRef = False Then
        
        lstDocRefCodigo.List(indiceAtualizarDocRef, 0) = cmbDocRefCodigo.Text
        lstDocRefCodigo.List(indiceAtualizarDocRef, 1) = txtCodigoRef.Text
    
    End If

End Sub

Private Sub cmdAtualizarInstFinanc_Click()

Dim arrayListaInstFinanc(1 To 10000, 1 To 2) As String

    bolAchouInstFinancLista = False

    indiceAtualizarInstFinanc = 0

    If cmbInstituicaoFinanc.Text = "" Or txtCodigoInstituicaoFinanc.Text = "" Then
    
        MsgBox "Campos Isntituição Financeira e/ou Código de Instituição Financeira não podem estar vazios.", vbOKOnly + vbInformation, "Exportação de Dados"
        Exit Sub
        
    End If

    For indiceArrayListaInstFinanc = 0 To lstInstFinancCodigo.ListCount - 1
    
        If cmbInstituicaoFinanc.Text = lstInstFinancCodigo.List(indiceArrayListaInstFinanc, 0) Then
            lstInstFinancCodigo.List(indiceArrayListaInstFinanc, 1) = txtCodigoInstituicaoFinanc.Text
            
            bolAchouInstFinancLista = True
            
        End If
        
        If lstInstFinancCodigo.List(indiceArrayListaInstFinanc, 0) <> "" Then
            indiceAtualizarInstFinanc = indiceAtualizarInstFinanc + 1
        End If
        
    Next
    
    
    If bolAchouInstFinancLista = False Then
        
        lstInstFinancCodigo.List(indiceAtualizarInstFinanc, 0) = cmbInstituicaoFinanc.Text
        lstInstFinancCodigo.List(indiceAtualizarInstFinanc, 1) = txtCodigoInstituicaoFinanc.Text
    
    End If
    
End Sub

Private Sub cmdAtualizarPlanoConta_Click()

Dim arrayPlanoConta(1 To 10000, 1 To 2) As String
Dim planoConta As String
Dim descricaoPlanoConta As String

    bolAchouPlanoConta = False

    indiceAtualizarPlanoConta = 0
    
    If Me.optNaoExportSistemaContabil.Value = True Then
        planoConta = Me.txtCodigoPlanoContas.Text
        descricaoPlanoConta = Me.cmbPlanoContaCodigo.Text
    Else
        planoConta = Me.txtPlanoContaImportado.Text
        descricaoPlanoConta = Me.cmbPlanoContaImportado.Text
    End If
    
    If planoConta = "" Then
    
        MsgBox "Campos Plano de Contas e/ou Código de Plano de Contas não podem estar vazios.", vbOKOnly + vbInformation, "Exportação de Dados"
        Exit Sub
        
    End If

    For indiceArrayPlanoConta = 0 To lstPlanoContaCodigo.ListCount - 1
        
        If descricaoPlanoConta = lstPlanoContaCodigo.List(indiceArrayPlanoConta, 0) Then
            lstPlanoContaCodigo.List(indiceArrayPlanoConta, 1) = planoConta
            
            bolAchouPlanoConta = True
            
        End If
        
        If lstPlanoContaCodigo.List(indiceArrayPlanoConta, 0) <> "" Then
            indiceAtualizarPlanoConta = indiceAtualizarPlanoConta + 1
        End If
        
    Next
    
    
    If bolAchouPlanoConta = False Then
        
        lstPlanoContaCodigo.List(indiceAtualizarPlanoConta, 0) = descricaoPlanoConta
        lstPlanoContaCodigo.List(indiceAtualizarPlanoConta, 1) = planoConta
    
    End If
    
End Sub

Private Sub cmdCaminho_Click()

Dim intResult As Integer
Dim strPath As String
    
    'the dialog is displayed to the user
    intResult = Application.FileDialog(msoFileDialogOpen).Show
    
    'checks if user has cancled the dialog
    If intResult <> 0 Then
        'dispaly message box
        txtArquivoPlanoContas.Text = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
    End If

End Sub

Private Sub cmdCarregarDados_Click()

Dim linha_panilha As Integer
Dim indice As Integer
Dim indiceAchou As Integer

Dim instFinanc(1 To 10000) As String
Dim docRef(1 To 10000) As String
Dim instFinancAchou(1 To 10000) As String
Dim docRefAchou(1 To 10000) As String
Dim arrayListaInstFinanc(1 To 10000, 1 To 2) As String
Dim arrayDocRef(1 To 10000, 1 To 2) As String
Dim arrayPlanoConta(1 To 10000, 1 To 2) As String
Dim colCodigoPlanoConta As String
Dim colDescricaoPlanoConta As String
Dim contaDevedoraLocalizada As String
Dim nomePlanilhaExport As String

Dim bolAchou As Boolean

    nomePlanilhaExport = ActiveSheet.Name
        
    
    lstInstFinancCodigo.Clear
    cmbInstituicaoFinanc.Clear
    
    Me.cmbDocRefCodigo.Clear
    Me.lstDocRefCodigo.Clear
    
    Me.cmbPlanoContaCodigo.Clear
    Me.lstPlanoContaCodigo.Clear
        
    linha_planilha = 5
    indice = 1
        
    Application.ScreenUpdating = False
    
    Me.frameProgressoExportacao.Visible = True
    
    Call barraProgresso("Processando dados na memória, Intituição Financeira e Documento de Referência ", 1)
    
    Do While Range("H" + CStr(linha_planilha)).Value <> ""
                
        instFinanc(indice) = Range("H" + CStr(linha_planilha)).Value
        docRef(indice) = Range("F" + CStr(linha_planilha)).Value
        
        indice = indice + 1
        linha_planilha = linha_planilha + 1
        
        Call barraProgresso("Processando dados na memória, Intituição Financeira e Documento de Referência ", indice)
        
    Loop
    
    Worksheets("Cenario de Exportacao").Activate
    
    '-----------------------------------------------------------------------------------------------------------
    'Instituição Financeira
    '-----------------------------------------------------------------------------------------------------------
    
    linha_planilha = 5
    indice = 1
    
    Do While Range("H" + CStr(linha_planilha)).Value <> ""
    
        cmbInstituicaoFinanc.AddItem Range("H" + CStr(linha_planilha)).Value
        
        arrayListaInstFinanc(indice, 1) = Range("H" + CStr(linha_planilha)).Value
        arrayListaInstFinanc(indice, 2) = Range("G" + CStr(linha_planilha)).Value
        
        indice = indice + 1
        
        linha_planilha = linha_planilha + 1
        
        Call barraProgresso("Listando Instituições Financeiras ", indice)
        
    Loop
    
    lstInstFinancCodigo.List = arrayListaInstFinanc
    
    linha_planilha = 5
    indiceAchou = 1
    bolVazio = False
    
    If Me.chkAtualizarCargaDados.Value = True Then
    
        For indice = 1 To UBound(instFinanc)
        
            Call barraProgresso("Eliminando termos repetidos ", indice)
            
            bolAchou = False
            
            Do While Range("H" + CStr(linha_planilha)).Value <> ""
        
                If Range("H" + CStr(linha_planilha)).Value = instFinanc(indice) Then
                    bolAchou = True
                    Exit Do
                End If
                
                linha_planilha = linha_planilha + 1
                
            Loop
            
            If bolAchou = False Then
            
                Range("H" + CStr(linha_planilha)).Value = instFinanc(indice)
                cmbInstituicaoFinanc.AddItem Range("H" + CStr(linha_planilha)).Value
                
            End If
            
            linha_planilha = 5
            
        Next
        
        bolAchou = False
        indiceNAchou = 1
        
        If linha_planilha = 5 Then
            
            For indice = 1 To UBound(instFinanc)
                
                Call barraProgresso("Armazenando Instituições Financeiras ", indice)
                
                bolAchou = False
                
                For indiceAchou = 1 To UBound(instFinancAchou)
                    If instFinancAchou(indiceAchou) = instFinanc(indice) Then
                        bolAchou = True
                        Exit For
                    End If
                Next
                
                
                If bolAchou = False Then
            
                    instFinancAchou(indiceNAchou) = instFinanc(indice)
                    
                    Range("H" + CStr(linha_planilha)).Value = instFinancAchou(indiceNAchou)
                    cmbInstituicaoFinanc.AddItem instFinancAchou(indiceNAchou)
                    
                    indiceNAchou = indiceNAchou + 1
                    linha_planilha = linha_planilha + 1
                
                End If
                
            Next
                
        End If
        
    End If
    
    '-----------------------------------------------------------------------------------------------------------
    'Documento de Referência
    '-----------------------------------------------------------------------------------------------------------
    
    linha_planilha = 5
    indice = 1
    
    Do While Range("J" + CStr(linha_planilha)).Value <> ""
    
        Me.cmbDocRefCodigo.AddItem Range("J" + CStr(linha_planilha)).Value
        
        arrayDocRef(indice, 1) = Range("J" + CStr(linha_planilha)).Value
        arrayDocRef(indice, 2) = Range("J" + CStr(linha_planilha)).Value
        
        indice = indice + 1
        
        linha_planilha = linha_planilha + 1
        
        Call barraProgresso("Listando Documentos de Referência ", indice)
        
    Loop
    
    Me.lstDocRefCodigo.List = arrayDocRef
    
    linha_planilha = 5
    indiceAchou = 1
    bolVazio = False
    
    If Me.chkAtualizarCargaDadosDocRef.Value = True Then
    
        For indice = 1 To UBound(docRef)
            
            bolAchou = False
            
            Call barraProgresso("Eliminando termos repetidos ", indice)
            
            Do While Range("J" + CStr(linha_planilha)).Value <> ""
        
                If Range("J" + CStr(linha_planilha)).Value = docRef(indice) Then
                    bolAchou = True
                    Exit Do
                End If
                
                linha_planilha = linha_planilha + 1
                
            Loop
            
            If bolAchou = False Then
            
                Range("J" + CStr(linha_planilha)).Value = docRef(indice)
                Me.cmbDocRefCodigo.AddItem Range("J" + CStr(linha_planilha)).Value
                
            End If
            
            linha_planilha = 5
            
        Next
        
        bolAchou = False
        indiceNAchou = 1
        
        If linha_planilha = 5 Then
            
            For indice = 1 To UBound(docRef)
            
                Call barraProgresso("Armazenando Documentos de Referência ", indice)
                
                bolAchou = False
                
                For indiceAchou = 1 To UBound(docRef)
                    If docRefAchou(indiceAchou) = docRef(indice) Then
                        bolAchou = True
                        Exit For
                    End If
                Next
                
                
                If bolAchou = False Then
            
                    docRefAchou(indiceNAchou) = docRef(indice)
                    
                    Range("J" + CStr(linha_planilha)).Value = docRefAchou(indiceNAchou)
                    cmbDocRefCodigo.AddItem docRefAchou(indiceNAchou)
                    
                    indiceNAchou = indiceNAchou + 1
                    linha_planilha = linha_planilha + 1
                
                End If
                
            Next
                
        End If
        
    End If
    
    '-----------------------------------------------------------------------------------------------------------
    'Plano de Contas
    '-----------------------------------------------------------------------------------------------------------
    
    linha_planilha = 5
    indice = 1
    
    Do While Range("L" + CStr(linha_planilha)).Value <> ""
    
        arrayPlanoConta(indice, 1) = Range("L" + CStr(linha_planilha)).Value
        arrayPlanoConta(indice, 2) = Range("K" + CStr(linha_planilha)).Value
        
        indice = indice + 1
        
        linha_planilha = linha_planilha + 1
        
        Call barraProgresso("Listando Planos de Conta ", indice)
        
    Loop
    
    Me.lstPlanoContaCodigo.List = arrayPlanoConta
    
    Worksheets("Configurações Básicas").Activate
    
    linha_planilha = 12
    achouContaDevedora = False
    
    Do While Range("E" + CStr(linha_planilha)).Value <> ""
    
        colCodigoPlanoConta = Range("G" + CStr(linha_planilha)).Value
        colDescricaoPlanoConta = Range("H" + CStr(linha_planilha)).Value

        If Range("F" + CStr(linha_planilha)).Value = "R" Then
            Worksheets("PC Receitas").Activate
        Else
            Worksheets("PC Despesas").Activate
        End If
    
        linhaplanoConta = 5
        
        If colCodigoPlanoConta <> "-" And colCodigoPlanoConta <> "" Then
        
            Do While Range(colCodigoPlanoConta + CStr(linhaplanoConta)).Value <> ""
                
                Me.cmbPlanoContaCodigo.AddItem Range(colDescricaoPlanoConta + CStr(linhaplanoConta)).Value
                linhaplanoConta = linhaplanoConta + 1
                
                Call barraProgresso("Armazenando Plano de Conta ", indice)
                
            Loop
            
        End If
        
        Worksheets("Configurações Básicas").Activate
        
        linha_planilha = linha_planilha + 1
        
    Loop
    
    Me.frameProgressoExportacao.Visible = False
    
    Application.ScreenUpdating = True
    
    Worksheets(nomePlanilhaExport).Activate
    
End Sub

Private Sub cmdFechar_Click()

    Unload Me

End Sub

Private Sub cmdImportarPlanoContas_Click()

On Error GoTo Erro

    If Me.cmbTipoArquivo.Text = "" Then
    
        MsgBox "Favor escolher o tipo de arquivo. " + Err.Description + ". Tente novamente ou avalie o formato do arquivo.", vbOKOnly + vbInformation, "Erro ao Importar Plano de Contas"
        Exit Sub
    
    End If

    If Me.cmbTipoArquivo.Text = "Texto" Then
        caracterImportacaoArquivoTexto = InputBox("Importação de Plano de Contas do Sistema de Origem para Vinculação de Contas do Fluxo de Caixa" + Chr(13) + "" + Chr(13) + "Digite o caracter separador do arquivo texto:", "Importação do Plano de Contas")
        Call Importar_Arquivo_Texto_PlanoContas
    Else
    
        frmArquivoPlanoContasExcel.Show
    
        If codigoReduzidoPlanoContas <> "" And classificacaoContabil <> "" And nomePlanoConta <> "" Then
            
            For indicePlanoConta = 1 To 5000
                arrayPlanoConta(indicePlanoConta, 1) = ""
                arrayPlanoConta(indicePlanoConta, 2) = ""
                arrayPlanoConta(indicePlanoConta, 3) = ""
            Next indicePlanoConta
            
            Set WB1 = Workbooks.Open(Me.txtArquivoPlanoContas.Text)
            
            linha = linhaInicialLeituraPlanilhaPlanoContas
            indicePlanoConta = 1
            
            Do While Range(codigoReduzidoPlanoContas + CStr(linha)).Value <> ""
                
                arrayPlanoConta(indicePlanoConta, 1) = Range(codigoReduzidoPlanoContas + CStr(linha)).Value
                arrayPlanoConta(indicePlanoConta, 2) = Range(classificacaoContabil + CStr(linha)).Value
                arrayPlanoConta(indicePlanoConta, 3) = Range(nomePlanoConta + CStr(linha)).Value
                
                indicePlanoConta = indicePlanoConta + 1
                
                linha = linha + 1
        
            Loop
            
            WB1.Close
                        
            Me.cmbPlanoContaImportado.Clear
    
            nomePlanilha = ActiveSheet.Name
             
            Worksheets("Cenario Importacao Plano Contas").Activate
            Range("G5").Select
            
            indicePlanoConta = 1
            
            Me.cmbPlanoContaImportado.Clear
            
            '--------------------------------------------------------------------------------------------
            'Limpar o plano de contas da planilha Cenario Importacao Plano Contas
            '--------------------------------------------------------------------------------------------
            linha = 5
            Do While Range("G" + CStr(linha)).Value <> ""
                
                Range("G" + CStr(linha)).Value = ""
                Range("H" + CStr(linha)).Value = ""
                Range("I" + CStr(linha)).Value = ""
                
                linha = linha + 1
                
            Loop
            '--------------------------------------------------------------------------------------------
            
            linha = 5
            indicePlanoConta = 1
            Do While arrayPlanoConta(indicePlanoConta, 1) <> ""
                 
                Me.cmbPlanoContaImportado.AddItem arrayPlanoConta(indicePlanoConta, 1)
                 
                Range("G" + CStr(linha)).Value = arrayPlanoConta(indicePlanoConta, 1)
                Range("H" + CStr(linha)).Value = arrayPlanoConta(indicePlanoConta, 2)
                Range("I" + CStr(linha)).Value = arrayPlanoConta(indicePlanoConta, 3)
                 
                indicePlanoConta = indicePlanoConta + 1
                 
                linha = linha + 1
                 
            Loop
             
            Range("G5").Select
             
            Worksheets(nomePlanilha).Activate
            
        Else
        
            MsgBox "Digite as três colunas de origem para a importação do plano de contas e a linha inicial. " + Err.Description + ". Tente novamente ou avalie o formato do arquivo.", vbOKOnly + vbInformation, "Erro ao Importar Plano de Contas"
            Exit Sub
        End If
            
    End If
    
    Exit Sub
    
Erro:

    MsgBox "Erro na importação do plano de contas. " + Err.Description + ". Tente novamente ou avalie o formato do arquivo.", vbOKOnly + vbInformation, "Erro ao Importar Plano de Contas"
    Exit Sub
    
End Sub



Private Sub optNaoExportSistemaContabil_Click()

    Me.txtArquivoPlanoContas.Enabled = False
    Me.cmbTipoArquivo.Enabled = False
    Me.cmbPlanoContaImportado.Visible = False
    Me.txtPlanoContaImportado.Visible = False
    Me.cmbPlanoContaImportado.Enabled = False
    Me.txtPlanoContaImportado.Enabled = False
    Me.lblPlanoContasImportado.Visible = False
    Me.lblPlanoContasDigitado.Visible = True
    Me.txtCodigoPlanoContas.Visible = True
    Me.cmdImportarPlanoContas.Enabled = False
    Me.cmdCaminho.Enabled = False
    
End Sub

Private Sub optSimExportSistemaContabil_Click()

    Me.txtArquivoPlanoContas.Enabled = True
    Me.cmbTipoArquivo.Enabled = True
    Me.cmbPlanoContaImportado.Visible = True
    Me.txtPlanoContaImportado.Visible = True
    Me.cmbPlanoContaImportado.Enabled = True
    Me.txtPlanoContaImportado.Enabled = False
    Me.lblPlanoContasImportado.Visible = True
    Me.lblPlanoContasDigitado.Visible = False
    Me.txtCodigoPlanoContas.Visible = False
    Me.cmdImportarPlanoContas.Enabled = True
    Me.cmdCaminho.Enabled = True
    
End Sub

Private Sub UserForm_Initialize()

Dim sistema(1 To 10) As String

    Call optSimExportSistemaContabil_Click
    Me.optSimExportSistemaContabil.Value = 1
    Me.optNaoExportSistemaContabil.Value = 0

    txtCodigoEmpresaExportacao.Enabled = False
    txtLoginUsuarioExportacao.Enabled = False
    
    txtCodigoEmpresaExportacao.BackColor = 12632256
    txtLoginUsuarioExportacao.BackColor = 12632256
    
    sistema(1) = "Sem Formato"
    sistema(2) = "Dominio"
    sistema(3) = "Prosoft"
    sistema(4) = "Alterdata"

    cmbSistemaExportacao.List = sistema
    
    cmbTipoArquivo.AddItem "Texto"
    cmbTipoArquivo.AddItem "Excel"
    
    '--------------------------------------------------------------------------------------------
    'Preencher o combo de plano de contas vindo da planilha Cenario Importacao Plano Contas
    '--------------------------------------------------------------------------------------------
        
    Call atualizaComboPlanoConta
    
    '--------------------------------------------------------------------------------------------
    
End Sub

Public Sub barraProgresso(mensagem As String, percentual As Integer)

    Me.lblDescricaoProgresso.Caption = mensagem + "... " + CStr(percentual) + " registros"
    DoEvents
    Me.lblProgresso.Width = ((percentual / 10000) * Me.lblDescricaoProgresso.Width)
    DoEvents
    
End Sub

Public Sub Importar_Arquivo_Texto_PlanoContas()

On Error GoTo Erro

    Dim Arquivo As String
    Dim X As Variant
    Dim S As String, N As Integer, C As Integer
    Dim rg As Range
    
    Arquivo = Me.txtArquivoPlanoContas.Text
    
    nomePlanilha = ActiveSheet.Name
    
    Worksheets("Cenario Importacao Plano Contas").Activate
    Range("G5").Select
    
    '--------------------------------------------------------------------------------------------
    'Limpar o plano de contas da planilha Cenario Importacao Plano Contas
    '--------------------------------------------------------------------------------------------
    linha = 5
    Do While Range("G" + CStr(linha)).Value <> ""
        
        Range("G" + CStr(linha)).Value = ""
        Range("H" + CStr(linha)).Value = ""
        Range("I" + CStr(linha)).Value = ""
        
        linha = linha + 1
        
    Loop
    '--------------------------------------------------------------------------------------------
    
    Set rg = Range("G5")
    
    Open Arquivo For Input As #1
        
    Do Until EOF(1)
        Line Input #1, S
        C = 0
        X = Split(S, caracterImportacaoArquivoTexto)
        For N = 0 To UBound(X)
            If X(N) <> "" Then
                rg.Offset(0, C) = X(N)
                C = C + 1
            End If
        Next N
        Set rg = rg.Offset(1, 0)
    Loop
    
    Close #1
    
    Worksheets(nomePlanilha).Activate
    
    Call atualizaComboPlanoConta
    
    Exit Sub
    
Erro:

    MsgBox "Erro ao processar a importação do arquivo. " + Err.Description + ". Tente novamente ou avalie o formato do arquivo.", vbOKOnly + vbInformation, "Erro ao Importar Plano de Contas"
    Close #1
    Worksheets(nomePlanilha).Activate
    
End Sub

Private Sub atualizaComboPlanoConta()

    Me.cmbPlanoContaCodigo.Clear
    
    nomePlanilha = ActiveSheet.Name
    
    Worksheets("Cenario Importacao Plano Contas").Activate
    Range("G5").Select
   
    For indicePlanoConta = 1 To 5000
        
        arrayPlanoConta(indicePlanoConta, 1) = ""
        arrayPlanoConta(indicePlanoConta, 2) = ""
        arrayPlanoConta(indicePlanoConta, 3) = ""
        
    Next indicePlanoConta
   
    indicePlanoConta = 1
   
    Me.cmbPlanoContaImportado.Clear
   
    linha = 5
    Do While Range("G" + CStr(linha)).Value <> ""
        
        If Range("I" + CStr(linha)).Value <> "" Then
        
            Me.cmbPlanoContaImportado.AddItem Range("I" + CStr(linha)).Value
            
            arrayPlanoConta(indicePlanoConta, 1) = Range("G" + CStr(linha)).Value
            arrayPlanoConta(indicePlanoConta, 2) = Range("H" + CStr(linha)).Value
            arrayPlanoConta(indicePlanoConta, 3) = Range("I" + CStr(linha)).Value
            
            indicePlanoConta = indicePlanoConta + 1
            
        End If
        
        linha = linha + 1
        
    Loop
    
    Range("G5").Select
    
    Worksheets(nomePlanilha).Activate

End Sub
