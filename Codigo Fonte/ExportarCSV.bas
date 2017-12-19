Attribute VB_Name = "ExportarCSV"
Sub ExportarCSV()
    
On Error Resume Next
    
    Dim NomeDoArquivo As String
    Dim WB1 As Workbook
    Dim WB2 As Workbook
    Dim rng As Range
     
    Set WB1 = ActiveWorkbook
    Set rng = Application.InputBox("Selecione o intervalo de células para a exportação da planilha atual:", "Processamento de Recebimentos", Default:="Por exemplo -> C5:L35", Type:=8)
    
    If rng <> "" Then
    
       Application.ScreenUpdating = False
       rng.Copy
    
       Set WB2 = Application.Workbooks.Add(1)
       WB2.Sheets(1).Range("A1").PasteSpecial xlPasteValues
        
       NomeDoArquivo = "CSV_Export_" & Format(Date, "ddmmyyyy")
       FullPath = WB1.Path & "\" & NomeDoArquivo
        
       Application.DisplayAlerts = False
       
       If MsgBox("Dados copiados para " & WB1.Path & "\" & NomeDoArquivo & vbCrLf & _
       "Atenção: Arquivos no diretório com mesmo nome serão sobrescritos!!", vbQuestion + vbYesNo) <> vbYes Then
           Exit Sub
       End If
        
       If Not Right(NomeDoArquivo, 4) = ".csv" Then MyFileName = NomeDoArquivo & ".csv"
       
       With WB2
           .SaveAs Filename:=FullPath, FileFormat:=xlCSV, CreateBackup:=False
           .Close False
       End With
       
       Application.DisplayAlerts = True
       
    End If
End Sub
