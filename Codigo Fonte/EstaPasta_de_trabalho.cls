VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EstaPasta_de_trabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Activate()

End Sub

Private Sub Workbook_Open()
    
    Application.DisplayFullScreen = True
    Application.WindowState = xlMaximized
    
    'Call Ocultar
    
    Sheets("Início").Select
    
    '------------------------------------------------------------
    'Retirar quando colocar em produção
    '------------------------------------------------------------
    Call Reexibir
    
End Sub

Private Sub Workbook_WindowActivate(ByVal Wn As Window)
    
    'Call Ocultar
    
End Sub

Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)

    ActiveWorkbook.Save
    'Call Reexibir
    
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

Dim numeroMes As Integer
Dim planilhaMes As String

    'Application.EnableEvents = False
    
    If anoAtual = "" Then
    
        planilhaMes = ActiveSheet.Name
    
        Worksheets("Configurações Básicas").Activate
    
        anoAtual = Range("E5").Value
        nomeClieAtual = Range("E9").Value
        cnpjClieAtual = Range("E8").Value
       
        Worksheets(planilhaMes).Activate
        
    End If
    
    If Target.Column = 3 And _
        Target.Row >= 5 And _
        (Sh.Name = "Jan" Or _
         Sh.Name = "Fev" Or _
         Sh.Name = "Mar" Or _
         Sh.Name = "Abr" Or _
         Sh.Name = "Mai" Or _
         Sh.Name = "Jun" Or _
         Sh.Name = "Jul" Or _
         Sh.Name = "Ago" Or _
         Sh.Name = "Set" Or _
         Sh.Name = "Out" Or _
         Sh.Name = "Nov" Or _
         Sh.Name = "Dez" _
        ) Then

        Select Case Sh.Name
            Case "Jan"
                numeroMes = 1
            Case "Fev"
                numeroMes = 2
            Case "Mar"
                numeroMes = 3
            Case "Abr"
                numeroMes = 4
            Case "Mai"
                numeroMes = 5
            Case "Jun"
                numeroMes = 6
            Case "Jul"
                numeroMes = 7
            Case "Ago"
                numeroMes = 8
            Case "Set"
                numeroMes = 9
            Case "Out"
                numeroMes = 10
            Case "Nov"
                numeroMes = 11
            Case "Dez"
                numeroMes = 12
        End Select

        If Range("C" + CStr(Target.Row)).Value <> "" Then

            Range("D" + CStr(Target.Row)).Value = Format(Range("C" + CStr(Target.Row)).Value, "00") + "/" + Format(CStr(numeroMes), "00") + "/" + CStr(anoAtual)

            If Not IsDate(Range("D" + CStr(Target.Row)).Value) And Range("C" + CStr(Target.Row)).Value <> "" Then
               MsgBox "Data inválida. Favor repetir a digitação do dia de lançamento.", vbInformation, "Processamento de Recebimentos"

               Range("C" + CStr(Target.Row)).Activate

               Range("D" + CStr(Target.Row)).Value = ""

            Else
                Range("D" + CStr(Target.Row)).Value = CDate(Format(Range("C" + CStr(Target.Row)).Value, "00") + "/" + Format(CStr(numeroMes), "00") + "/" + CStr(anoAtual))
            End If

        Else

            Range("D" + CStr(Target.Row)).Value = ""

        End If
    
    End If

    'Application.EnableEvents = True
 
End Sub

Private Sub Workbook_WindowResize(ByVal Wn As Window)
    'Wn.WindowState = xlMaximized
    'Wn.EnableResize = False
    'Call Ocultar
End Sub

Function ConverteParaLetra(iCol As Integer) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   
   If iAlpha > 0 Then
      ConverteParaLetra = Chr(iAlpha + 64)
   End If
   
   If iRemainder > 0 Then
      ConverteParaLetra = ConverteParaLetra & Chr(iRemainder + 64)
   End If
   
End Function



