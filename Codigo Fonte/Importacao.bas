Attribute VB_Name = "Importacao"
' Location of 'Adobe Acrobat Reader' (only used, if it is not the default PDF reader)
  Private Const AdobePDFReader As String = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.EXE"

' Public variable to test, if we were successful in copying from PDF document to Excel worksheet
  Public PDF2XL_Success As Boolean

' API Functions
 #If VBA7 = False Then
      Private Declare Function FindExecutable Lib "shell32" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
      Private Declare Function DownloadURLToFile Lib "URLMon.DLL" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
 #Else
      Private Declare PtrSafe Function FindExecutable Lib "shell32" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long
      Private Declare PtrSafe Function DownloadURLToFile Lib "URLMon.DLL" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
 #End If

Option Explicit

Sub importar_Com_Parametro()
    
    resposta = MsgBox("Deseja realmente processar a importação com Parâmetros?", vbYesNo + vbExclamation, "Processamento de Recebimentos")
 
    If resposta = vbYes Then frmEscolhaDesRec.Show
    
End Sub

Public Sub PDF2XL(Optional ByVal PDFFile As String = vbNullString, Optional ByVal DestinationWorksheet As Excel.Worksheet)

' This macro shows how you can 'import' a PDF document into Excel,
' by simply opening the PDF document in 'Adobe Acrobat Reader',
' and then just copy/paste the contents, using 'SendKeys' function.
'
' If you don't provide a PDF document, as argument, or, if the given PDF document cannot be found,
' the user is asked for at PDF document, using a standard file dialog.
'
' If you don't provide a destination worksheet to copy the PDF contents to, the active sheet will be used.
'
' The macro accepts PDF documents located on-line, on a 'http' URL, like "http://www.EXCELGAARD.dk/Files/PDFs/Extern Data.PDF"
'
' If your default PDF reader is not 'Adobe Acrobat Reader', there's a great risk, that this will not work!
' However, you can install 'Adobe Acrobat Reader', and still have your favorite PDF program use PDFs as default.
' To do so, you must provide the full path to the installed 'Adobe Acrobat Reader' in the constant in the declaration area at the top of the module.

' * ' Initialize
      On Error Resume Next


' * ' Define variables
      PDF2XL_Success = False
                  

      If TypeName(DestinationWorksheet) <> "Worksheet" Then Set DestinationWorksheet = ActiveSheet

      Dim TempFile As String
      If LCase$(Left$(PDFFile, 4)) = "http" Then                                    ' An online PDF document is given - try to download it
            TempFile = Environ("TMP")
            If Right$(TempFile, 1) <> Application.PathSeparator Then TempFile = TempFile & Application.PathSeparator
            TempFile = TempFile & "TempFile.PDF"

            Kill TempFile

            If DownloadURLToFile(0, PDFFile, TempFile, 16, 0) <> 0 Then GoTo ES:    ' Download of online PDF document failed
            PDFFile = TempFile
      End If

      If Len(PDFFile) < xlLess Or Len(Dir(PDFFile, vbHidden + vbSystem)) < 3 Then   ' If no PDF document is given, then ask for one
            PDFFile = Application.GetOpenFilename("PDF (*.PDF), *.PDF")
            If Len(PDFFile) < xlLess Then GoTo ES:                                  ' User clicked [Cancel]
      End If

      Dim FileAddressBuffer As String
      FileAddressBuffer = Space$(260)

      Dim FileHandle As Long
      FileHandle = FindExecutable(Mid$(PDFFile, InStrRev(PDFFile, Application.PathSeparator) + 1), Left$(PDFFile, InStrRev(PDFFile, Application.PathSeparator)), FileAddressBuffer)

      Dim PDFReader As String
      If FileHandle >= 32 Then                                                      ' System has a PDF application installed
            FileHandle = InStr(FileAddressBuffer, Chr$(0))
            PDFReader = Left$(FileAddressBuffer, FileHandle - 1)                    ' Default PDF application of system
      Else                                                                          ' System does not have a PDF application installed
            Select Case Application.LanguageSettings.LanguageID(2)                  ' Insert your own language below, if you want to
                  Case 1030, 1080:  MsgBox "Kunne ikke finde PDF Reader på computer.", vbOKOnly + vbCritical, " PDF Reader"
                  Case Else:        MsgBox "Could not locate PDF Reader on computer.", vbOKOnly + vbCritical, " PDF Reader"
            End Select
            GoTo ES:
      End If

      FileHandle = InStrRev(UCase$(PDFReader), "ADOBE")
      If FileHandle > 0 Then
            FileHandle = InStrRev(UCase$(PDFReader), "READER")
      End If
      If FileHandle < 1 Then                                                        ' The default PDF application is not 'Adobe PDF Reader'
            If Len(Dir(AdobePDFReader)) < 5 Then                                    ' The given PDF Reader in the constant in the declaration field can not be found
                  Select Case Application.LanguageSettings.LanguageID(2)            ' Insert your own language below, if you want to
                        Case 1030, 1080:  FileHandle = MsgBox("Den fundne PDF læser..." & vbNewLine & vbNewLine & PDFReader & vbNewLine & vbNewLine & "...ser ikke ud til at være 'Adobe Acrobat Reader'." & vbNewLine & vbNewLine & "Forsætte?", vbYesNo + vbExclamation, " PDF Reader")
                        Case Else:        FileHandle = MsgBox("The found PDF reader..." & vbNewLine & vbNewLine & PDFReader & vbNewLine & vbNewLine & "...doesn't seems to be 'Adobe Acrobat Reader'." & vbNewLine & vbNewLine & "Continue?", vbYesNo + vbExclamation, " PDF Reader")
                  End Select
                  If FileHandle = vbNo Then GoTo ES:
            Else
                  PDFReader = AdobePDFReader
            End If
      End If
      PDFReader = Chr(34) & PDFReader & Chr(34) & " " & Chr(34) & Replace(PDFFile, Chr(34), vbNullString) & Chr(34)


' * ' Prepare worksheet
      DestinationWorksheet.DisplayPageBreaks = False

      DestinationWorksheet.Unprotect
      If DestinationWorksheet.ProtectContents = True Then GoTo ES:

      DestinationWorksheet.Visible = xlSheetVisible
      If DestinationWorksheet.Visible <> xlSheetVisible Then GoTo ES:

      DestinationWorksheet.Select
      DestinationWorksheet.Cells.Delete

      Range("A1").Select


' * ' Transfer PDF contents to Excel
      Application.CutCopyMode = False                                               ' Clear/reset Cut/Copy mode

      Shell PDFReader, vbNormalFocus                                                ' Open PDF document

      Application.Wait Now + TimeValue("00:00:03")                                  ' Wait a little to give document time to fully open
      DoEvents

      SendKeys "^a"                                                                 ' Select all in PDF document
      SendKeys "^c"                                                                 ' Copy selected contents

      Application.Wait Now + TimeValue("00:00:02")                                  ' Wait a little to give clipboard time to copy (if huge contents)
      DoEvents

      SendKeys "^q"                                                                 ' Close PDF document

      Application.Wait Now + TimeValue("00:00:01")                                  ' Wait a little to give document time to close completely
      Application.Run "ActivateExcel", True                                         ' Re-activate Excel in case another application was activate when closing PDF Reader
      DoEvents

      Err.Clear
      With DestinationWorksheet
        .Range("A1").Select
        .PasteSpecial xlPasteFormats                                                  ' Paste PDF contents into worksheet
      End With
      
      If Err.Number = 0 Then PDF2XL_Success = True


ES: ' End of Sub
      Application.CutCopyMode = False                                               ' Clear/reset Cut/Copy mode

      Range("A1").Select

      Set DestinationWorksheet = Nothing
      If TempFile <> vbNullString Then Kill TempFile

End Sub

Public Sub PDF2XL_Test()

Dim NovoArquivoXLS As Workbook
Dim sht As Worksheet
Dim path As String
Dim dataAtual As String
Dim horaAtual As String

' * ' Initialize
      On Error Resume Next

      path = Application.ActiveWorkbook.path

      'Cria um novo arquivo excel
      Set NovoArquivoXLS = Application.Workbooks.Add
      
      'Salva o arquivo
       NovoArquivoXLS.Save
      
       With NovoArquivoXLS
          Set sht = .Sheets.Add(After:=.Sheets(.Sheets.Count))
          sht.Name = "PlanilhaPDF"
      End With
      
      NovoArquivoXLS.Activate
      sht.Select
      
' * ' Define variable
      Dim PDFFile As String
      'PDFFile = "http://www.EXCELGAARD.dk/Files/PDFs/Extern Data.PDF"
      PDFFile = "c:\teste.PDF"


' * ' Copy PDF contents to active Excel worksheet
      Call PDF2XL(PDFFile, sht)

' * ' Here you can adjust the copied contents
      If PDF2XL_Success = True Then Application.Run "PDF2XL_Adjust"
      
      dataAtual = Replace(CStr(Date), "/", "")
      horaAtual = Replace(CStr(Time), ":", "")
      
      NovoArquivoXLS.SaveAs path & "\" & "importacaoPDF_fluxocaixa" & dataAtual & "_" & horaAtual & ".xls", _
        FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
      
ES: ' End of Sub
      Range("A1").Select

End Sub
