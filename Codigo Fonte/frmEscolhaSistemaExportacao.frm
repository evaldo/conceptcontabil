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

Private Sub btnGerarExportacao_Click()

    If ValidaPlanilhaProcessamento() = False Then
        MsgBox "Escolha um planilha para lançamento do Fluxo de Caixa entre Jan e Dez.", vbOKOnly + vbInformation, "Salvar Dados"
        Exit Sub
    End If
    
    Select Case cmbSistemaExportacao.Text
    
        Case "Sem Formato"
            Call ExportarCSVSemFormato
            
        Case "Dominio"
            Call ExportarDominio
            
        Case "Prosoft"
            Call ExportarProsoft
    
    End Select

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
Dim arrayListaInstFinanc(1 To 10000, 1 To 1000) As String

Dim bolAchou As Boolean

    nomePlanilha = ActiveSheet.Name
    
    lstInstFinancCodigo.Clear
    cmbInstituicaoFinanc.Clear
    
    linha_planilha = 5
    indice = 1
    
    Do While Range("H" + CStr(linha_planilha)).Value <> ""
                
        instFinanc(indice) = Range("H" + CStr(linha_planilha)).Value
        docRef(indice) = Range("F" + CStr(linha_planilha)).Value
        indice = indice + 1
        linha_planilha = linha_planilha + 1
        
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
        
        linha_planilha = linha_planilha + 1
        
    Loop
    
    lstInstFinancCodigo.List = arrayListaInstFinanc
    
    linha_planilha = 5
    indiceAchou = 1
    bolVazio = False
    
    If Me.chkAtualizarCargaDados.Value = True Then
    
        For indice = 1 To UBound(instFinanc)
            
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

    cmbSistemaExportacao.List = sistema

End Sub


