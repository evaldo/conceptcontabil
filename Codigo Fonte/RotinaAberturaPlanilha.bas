Attribute VB_Name = "RotinaAberturaPlanilha"

Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
ByVal hwnd As Long, _
ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
ByVal hwnd As Long, _
ByVal nIndex As Long, _
ByVal dwNewLong As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

 
Private Enum ESetWindowPosStyles
    SWP_SHOWWINDOW = &H40
    SWP_HIDEWINDOW = &H80
    SWP_FRAMECHANGED = &H20
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_DRAWFRAME = SWP_FRAMECHANGED
    HWND_NOTOPMOST = -2
End Enum

 
Private Declare Function GetWindowRect Lib "user32" _
    (ByVal hwnd As Long, lpRect As RECT) As Long

 
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

 
Sub Title_Show()
    ShowTitleBar True
End Sub

 
Sub Title_Hide()
    ShowTitleBar False
End Sub

 
Sub ShowTitleBar(bShow As Boolean)
    Dim lStyle As Long
    Dim tRect As RECT
    Dim sWndTitle As String
    Dim xlhnd

    '// Untested should perhaps look for the class ?!
    sWndTitle = "Microsoft Excel - " & ActiveWindow.Caption
    xlhnd = FindWindow(vbNullString, sWndTitle)

    '// Get the window's position:
    GetWindowRect xlhnd, tRect

    '// Show the Title bar ?
    If Not bShow Then
        lStyle = GetWindowLong(xlhnd, GWL_STYLE)
        lStyle = lStyle And Not WS_SYSMENU
        lStyle = lStyle And Not WS_MAXIMIZEBOX
        lStyle = lStyle And Not WS_MINIMIZEBOX
        lStyle = lStyle And Not WS_CAPTION
    Else
        lStyle = GetWindowLong(xlhnd, GWL_STYLE)
        lStyle = lStyle Or WS_SYSMENU
        lStyle = lStyle Or WS_MAXIMIZEBOX
        lStyle = lStyle Or WS_MINIMIZEBOX
        lStyle = lStyle Or WS_CAPTION
    End If
    SetWindowLong xlhnd, GWL_STYLE, lStyle
    Application.DisplayFullScreen = Not bShow
    '// Ensure the style is set and makes the xlwindow the
    '// same size, regardless of the title bar.
    SetWindowPos xlhnd, 0, tRect.Left, tRect.Top, tRect.Right - tRect.Left, _
        tRect.Bottom - tRect.Top, SWP_NOREPOSITION Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub

Public Sub Ocultar()
    
Dim barras
    
    For Each barras In Application.CommandBars
        barras.Enabled = False
    Next
    
    ' *** File ***
    Application.OnKey "^N", "" 'Ctrl+N novo arquivo
    Application.OnKey "^O", "" 'Ctrl+O abrir arquivo
    Application.OnKey "{F12}", "" 'F12 salvar como
    Application.OnKey "{ESCAPE}", "" 'ESC
    '*** Edit ***
    Application.OnKey "^H", "" 'Ctrl+H replace
    Application.OnKey "{F5}", "" 'F5 Goto
    '*** Insert ***
    Application.OnKey "^+{+}", "" 'Ctrl+Shift+ + inserir dialog box
    Application.OnKey "+{F11}", "" 'Shift+F11 novo worksheet
    Application.OnKey "{F11}", "" 'F11 novo gráfico
    Application.OnKey "^{F11}", "" 'Ctrl+F11 macro do Excel 4.0
    Application.OnKey "+{F3}", "" 'Ctrl+F3 definir nome
    Application.OnKey "{F3}", "" 'F3 colar nomes
    Application.OnKey "^+{F3}", "" 'Ctrl+Shift+F3 criar nomes
    '*** Format ***
    Application.OnKey "^1", "" 'Ctrl+1 formatar células
    Application.OnKey "^9", "" 'Ctrl+9 esconder linhas
    Application.OnKey "^+{(}", "" 'Ctrl+Shift+( mostrar linhas
    Application.OnKey "^0", "" 'Ctrl+0 esconder colunas
    Application.OnKey "^+{)}", "" 'Ctrl+Shift+) mostrar colunas
    '*** Data ***
    Application.OnKey "%+{RIGHT}", "" 'Alt+Shift+RightArrow agrupa linhas/colunas
    Application.OnKey "%+{LEFT}", "" 'Alt+Shift+LeftArrow desagrupa linhas/colunas
    '*** Window ***
    Application.OnKey "{F6}", "" 'F6 próximo painel
    Application.OnKey "+{F6}", "" 'Shift+F6 painel anterior
    Application.OnKey "^{F6}", "" 'Ctrl+F6 próxima janela
    Application.OnKey "^+{F6}", "" 'Ctrl+Shift+F6 janela anterior
    '*** Outros ***
    Application.OnKey "^{PGUP}", "" 'Ctrl+PgUp sheet anterior
    Application.OnKey "^{PGDN}", "" 'Ctrl+PgDn sheet posterior
    Application.OnKey "+{F12}", "" 'Shift+F12 salvar
    
    Application.EnableCancelKey = xlDisabled
        
    ActiveWindow.DisplayHeadings = False
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayVerticalScrollBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.Caption = "Fluxo de Caixa"
    ActiveWindow.EnableResize = False
        
    Sheets("Início").Select

    Call Title_Hide
    
End Sub

Public Sub Reexibir()

Dim barras
    
    Application.EnableCancelKey = xlDisabled

On Error Resume Next
    For Each barras In Application.CommandBars
        barras.Enabled = True
    Next

    Application.DisplayFormulaBar = True
    
    ' *** File ***
    Application.OnKey "^N" 'Ctrl+N novo arquivo
    Application.OnKey "^O" 'Ctrl+O abrir arquivo
    Application.OnKey "^S" 'Ctrl+S salvar
    Application.OnKey "{F12}" 'F12 salvar como
    Application.OnKey "{ESCAPE}" 'ESC
    Application.OnKey "%{F4}" 'Alt+F4 sair do Excel
    ' *** Edit ***
    Application.OnKey "^H" 'Ctrl+H replace
    Application.OnKey "{F5}" 'F5 Goto
    ' *** Insert ***
    Application.OnKey "^+{+}" 'Ctrl+Shift+ + inserir dialog box
    Application.OnKey "+{F11}" 'Shift+F11 novo worksheet
    Application.OnKey "{F11}" 'F11 novo gráfico
    Application.OnKey "^{F11}" 'Ctrl+F11 macro do Excel 4.0
    Application.OnKey "+{F3}" 'Ctrl+F3 definir nome
    Application.OnKey "{F3}" 'F3 colar nomes
    Application.OnKey "^+{F3}" 'Ctrl+Shift+F3 criar nomes
    ' *** Format ***
    Application.OnKey "^1" 'Ctrl+1 formatar células
    Application.OnKey "^9" 'Ctrl+9 esconder linhas
    Application.OnKey "^+{(}" 'Ctrl+Shift+( mostrar linhas
    Application.OnKey "^0" 'Ctrl+0 esconder colunas
    Application.OnKey "^+{)}" 'Ctrl+Shift+) mostrar colunas
    ' *** Data ***
    Application.OnKey "%+{RIGHT}" 'Alt+Shift+RightArrow agrupa linhas/colunas
    Application.OnKey "%+{LEFT}" 'Alt+Shift+LeftArrow desagrupa linhas/colunas
    ' *** Window ***
    Application.OnKey "{F6}" 'F6 próximo painel
    Application.OnKey "+{F6}" 'Shift+F6 painel anterior
    Application.OnKey "^{F6}" 'Ctrl+F6 próxima janela
    Application.OnKey "^+{F6}" 'Ctrl+Shift+F6 janela anterior
    ' *** Outros ***
    Application.OnKey "^{PGUP}" 'Ctrl+PgUp sheet anterior
    Application.OnKey "^{PGDN}" 'Ctrl+PgDn sheet posterior
    Application.OnKey "+{F12}" 'Shift+F12 salvar
    Application.OnKey "^{F12}" 'Ctrl+F12 abrir
    Application.OnKey "^{TAB}" 'Ctrl+Tab próxima janela
    Application.OnKey "^+{TAB}" 'Ctrl+Shift+Tab janela anterior
    Application.OnKey "{TAB}" 'Tab
    Application.OnKey "^{-}" 'Ctrl+- exclui seleção
    Application.OnKey "^{;}" 'Ctrl+; insere data
    Application.OnKey "^{:}" 'Ctrl+: insere hora
    
    EnableControl 21, True ' Recortar
    EnableControl 19, True ' Copiar
    EnableControl 22, True ' Colar
    EnableControl 755, True 'ColarEspecial
    
    Application.DisplayFormulaBar = True
    Application.DisplayFullScreen = True
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    ActiveWindow.EnableResize = True
    
    Call Title_Show
   
End Sub


Public Sub ConfiguracoesBasicas()
'
' ConfiguracoesBasicas Macro
'

'
    Sheets("Configurações Básicas").Select
    
End Sub


Public Sub Inicio()
'
' AjudaeDica Macro
'

'
    Sheets("Início").Select
    
End Sub

Public Sub Imprimir()
'
' Imprimir Macro
'

'
    Sheets("Imprimir").Select
    
End Sub


Public Sub LogDeProcessamento()
'
' LogdeProcessamento Macro
'

'
    Sheets("Log de Proc Recebimentos").Select
    
End Sub


Public Sub Duvidas()
'
' Duvidas Macro
'

'
    Sheets("Dúvidas").Select
    
End Sub


Public Sub AjudaeDica()
'
' AjudaeDica Macro
'

'
    Sheets("Alertas").Select
    
End Sub

Public Sub Grafico()
'
' Grafico Macro
'

'
    Sheets("Gráficos").Select
    
End Sub

Public Sub ResultadoConsolidado()
'
' ResultadoConsolidado Macro
'

'
    Sheets("FC").Select
    
End Sub
Public Sub ManterPlanoContas()
Attribute ManterPlanoContas.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ManterPlanoContas Macro
'

'
    Sheets("PC Receitas").Select
    
End Sub


Public Sub ManterLancamento()
'
' ManterLancamento
'

'
    frmEscolhaLancamento.Show
    
End Sub

Public Function ValidaPlanilhaProcessamento() As Boolean

Dim mes(1 To 12) As String
Dim iMes As Integer
    
    mes(1) = "Jan"
    mes(2) = "Fev"
    mes(3) = "Mar"
    mes(4) = "Abr"
    mes(5) = "Mai"
    mes(6) = "Jun"
    mes(7) = "Jul"
    mes(8) = "Ago"
    mes(9) = "Set"
    mes(10) = "Out"
    mes(11) = "Nov"
    mes(12) = "Dez"
 
    ValidaPlanilhaProcessamento = False
    
    For iMes = 1 To 12
    
        If ActiveSheet.Name = mes(iMes) Then
            ValidaPlanilhaProcessamento = True
            Exit For
        End If
        
    Next iMes

End Function

Sub EnableControl(iId As Integer, blnState As Boolean)

    Dim ComBar As CommandBar
    Dim ComBarCtrl As CommandBarControl
    
On Error Resume Next
    
    For Each ComBar In Application.CommandBars
    
        Set ComBarCtrl = ComBar.FindControl(ID:=iId, recursive:=True)
    
        If Not ComBarCtrl Is Nothing Then ComBarCtrl.Enabled = blnState
    
    Next

End Sub



