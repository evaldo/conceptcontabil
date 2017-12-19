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

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    
    Application.EnableEvents = False
    
    If Target.Column = 3 And Target.Row >= 5 Then
        If Not IsDate(Range("D" + CStr(Target.Row)).Value) And Range("C" + CStr(Target.Row)).Value <> "" Then
           MsgBox "Data inválida. Favor repetir a digitação do dia de lançamento.", vbInformation, "Processamento de Recebimentos"
           Range("C" + CStr(Target.Row)).Activate
        End If
        
    End If
    
    Application.EnableEvents = True
 
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

