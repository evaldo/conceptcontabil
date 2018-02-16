VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmArquivoPlanoContasExcel 
   Caption         =   "Definição das Colunas do Arquivo de Plano de Contas"
   ClientHeight    =   2700
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5070
   OleObjectBlob   =   "frmArquivoPlanoContasExcel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmArquivoPlanoContasExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceitarColunasPlanoContas_Click()

    codigoReduzidoPlanoContas = Me.txtColunaCodigoReduzido.Text
    classificacaoContabil = Me.txtColunaclassificacao.Text
    nomePlanoConta = Me.txtNomeConta.Text
    linhaInicialLeituraPlanilhaPlanoContas = Me.txtLinhaInicialLeitura.Text
    
    Me.Hide

End Sub

Private Sub cmdCancelar_Click()

    codigoReduzidoPlanoContas = ""
    classificacaoContabil = ""
    nomePlanoConta = ""
    linhaInicialLeituraPlanilhaPlanoContas = 0
    
    Unload Me

End Sub


Private Sub UserForm_Click()

End Sub
