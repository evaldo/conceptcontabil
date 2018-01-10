VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEscolhaSistemaExportacao 
   Caption         =   "Escolha do Sistema Contábil para Exportação dos Dados"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7335
   OleObjectBlob   =   "frmEscolhaSistemaExportacao.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEscolhaSistemaExportacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public indiceArrayListaInstFinanc As Integer
Public bolAchouInstFinancLista As Boolean
Public indiceAtualizarInstFinanc As Integer

Public indiceArrayListaDocRef As Integer
Public bolAchoudocRef As Boolean
Public indiceAtualizarDocRef As Integer

Private Sub btnGerarExportacao_Click()

Dim indice As Integer
Dim linha_planilha As Integer
Dim nomePlanilha As String

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

Private Sub cmdCarregarDados_Click()

Dim linha_panilha As Integer
Dim indice As Integer
Dim indiceAchou As Integer

Dim nomePlanilha As String
Dim instFinanc(1 To 10000) As String
Dim docRef(1 To 10000) As String
Dim instFinancAchou(1 To 10000) As String
Dim docRefAchou(1 To 10000) As String
Dim arrayListaInstFinanc(1 To 10000, 1 To 2) As String
Dim arrayDocRef(1 To 10000, 1 To 2) As String

Dim bolAchou As Boolean

    nomePlanilha = ActiveSheet.Name
    
    lstInstFinancCodigo.Clear
    cmbInstituicaoFinanc.Clear
    
    Me.cmbDocRefCodigo.Clear
    Me.lstDocRefCodigo.Clear
        
    linha_planilha = 5
    indice = 1
    
    Me.frameProgressoExportacao.Visible = True
    
    Call barraProgresso("Processando dados na memória, Intituição Financeira e Documento de Referência ", 1)
    
    Application.ScreenUpdating = False
    
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
    
    Me.frameProgressoExportacao.Visible = False
    
    Application.ScreenUpdating = True
    
    Worksheets(nomePlanilha).Activate
    
End Sub

Private Sub cmdFechar_Click()

    Unload Me

End Sub

Private Sub UserForm_Initialize()

Dim sistema(1 To 10) As String

    txtCodigoEmpresaExportacao.Enabled = False
    txtLoginUsuarioExportacao.Enabled = False
    
    txtCodigoEmpresaExportacao.BackColor = 12632256
    txtLoginUsuarioExportacao.BackColor = 12632256
    
    sistema(1) = "Sem Formato"
    sistema(2) = "Dominio"
    sistema(3) = "Prosoft"
    sistema(4) = "Alterdata"

    cmbSistemaExportacao.List = sistema
    
End Sub

Public Sub barraProgresso(mensagem As String, percentual As Integer)

    Me.lblDescricaoProgresso.Caption = mensagem + "... " + CStr(percentual) + " registros"
    DoEvents
    Me.lblProgresso.Width = ((percentual / 10000) * Me.lblDescricaoProgresso.Width)
    DoEvents
    
End Sub


