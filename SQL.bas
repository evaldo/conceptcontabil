Attribute VB_Name = "SQL"
Sub ExportardadosSQL()

    Dim ano As String
    Dim mes(1 To 12) As String
    Dim numeroMes As Integer
    Dim mes_processamento As String
    Dim strSQL As String
    Dim ConnectionString As String
    Dim StrQuery As String
    Dim dataTransformada As String
    Dim nomeClie As String
    Dim cnpjClie As String
    
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim rstTempo As New ADODB.Recordset
    
    Dim linha As Integer
    Dim qtFluxo As Integer
    Dim qtRegistroCommit As Integer
    
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
    
    mes_processamento = ActiveSheet.Name
    
    For numeroMes = 1 To 12
        If mes(numeroMes) = mes_processamento Then Exit For
    Next numeroMes
    
    Worksheets("Configurações Básicas").Activate
    ano = Range("E5").Value
    nomeClie = Range("E9").Value
    cnpjClie = Range("E8").Value
    
    Worksheets(mes_processamento).Activate
    
    cnn.ConnectionString = "Driver={ODBC Driver 13 for SQL Server};Server=tcp:contarcon.database.windows.net,1433;Database=fluxocaixa;Uid=evaldo@contarcon;Pwd={Gcas1302};Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"
    cnn.Open
    
    cnn.BeginTrans
    
    StrQuery = "SELECT COUNT(1), MAX(ID_FLUXO_CAIXA)+1 FROM T_FLUXO_CAIXA"
    rst.Open (StrQuery), cnn
    
    If rst(0).Value = 0 Then
        qtFluxo = 1
    Else
        qtFluxo = rst(1).Value
    End If
    
    rst.Close
    
    linha = 5
    qtRegistroCommit = 0
    
    Do While Range("C" + CStr(linha)).Value <> ""
    
        If Not IsDate("" & numeroMes & "/" & Range("C" + CStr(linha)).Value & "/" & ano & "") Then
            StrQuery = "SELECT ID_DMSAO_TEMPO FROM T_DMSAO_TEMPO WHERE DT_DMSAO_TEMPO = CONVERT(VARCHAR(10), '" & UltimoDiaMes(CDate("1/" & numeroMes & "/" & ano)) & "', 103)"
            dataTransformada = UltimoDiaMes(CDate("1/" & numeroMes & "/" & ano))
        Else
            StrQuery = "SELECT ID_DMSAO_TEMPO FROM T_DMSAO_TEMPO WHERE DT_DMSAO_TEMPO = CONVERT(VARCHAR(10), '" & numeroMes & "/" & Range("C" + CStr(linha)).Value & "/" & ano & "', 103)"
            dataTransformada = "" & numeroMes & "/" & Range("C" + CStr(linha)).Value & "/" & ano & ""
        End If
        
        rstTempo.Open (StrQuery), cnn
        
        strSQL = "INSERT INTO T_FLUXO_CAIXA (ID_FLUXO_CAIXA, NU_CNPJ,SK_DMSAO_TEMPO,DT_MVMT_FLUXO_CAIXA, NM_CLIE_FLUXO_CAIXA,"
        strSQL = strSQL + "DS_CLSSF_PLANO_CONTA,CD_DCTO_RFRC_FLUXO_CAIXA,CD_PLANO_CONTA,DS_PLANO_CONTA,"
        strSQL = strSQL + "DS_INSTT_FNCR,VL_ENTR_FLUXO_CAIXA,VL_SAIDA_FLUXO_CAIXA, IC_STATUS_VALOR) VALUES("
        strSQL = strSQL + "" & qtFluxo & ","
        strSQL = strSQL + "'" & cnpjClie & "',"
        strSQL = strSQL + "" & rstTempo(0).Value & ","
        strSQL = strSQL + "CONVERT(VARCHAR(10), '" & dataTransformada & "', 103),"
        strSQL = strSQL + "'" & nomeClie & "',"
        strSQL = strSQL + "'" & Range("E" + CStr(linha)).Value & "',"
        strSQL = strSQL + "'" & Range("F" + CStr(linha)).Value & "',"
        strSQL = strSQL + "99999,"
        strSQL = strSQL + "'" & Range("G" + CStr(linha)).Value & "',"
        strSQL = strSQL + "'" & Range("H" + CStr(linha)).Value & "',"
        strSQL = strSQL + "'" & Replace(Range("J" + CStr(linha)).Value, ",", ".") & "',"
        strSQL = strSQL + "'" & Replace(Range("K" + CStr(linha)).Value, ",", ".") & "',"
        strSQL = strSQL + "'" & Range("L" + CStr(linha)).Value & "');"
        
        cnn.Execute strSQL
        
        If qtRegistroCommit = 10 Then
            cnn.CommitTrans
            qtRegistroCommit = 0
            cnn.BeginTrans
        End If
        
        linha = linha + 1
        qtFluxo = qtFluxo + 1
        qtRegistroCommit = qtRegistroCommit + 1
        
        rstTempo.Close
        
    Loop
    
    cnn.CommitTrans
    
    cnn.Close

End Sub

Function UltimoDiaMes(Data As Date) As String

    UltimoDiaMes = VBA.DateSerial(VBA.Year(Data), VBA.Month(Data) + 1, 0)
    UltimoDiaMes = CStr(Month(UltimoDiaMes) & "/" & Day(UltimoDiaMes) & "/" & Year(UltimoDiaMes))

End Function
