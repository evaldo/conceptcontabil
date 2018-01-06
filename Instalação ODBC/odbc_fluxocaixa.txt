-ECHO OFF 
odbcconf.exe /a {CONFIGDSN "ODBC Driver 13 for SQL Server" "DSN=fluxocaixa|Description=Dados Dimensionais do Caixa|SERVER=contarcon.database.windows.net|Trusted_Connection=Yes|Database=fluxocaixa"}
REM pause 
@CLS 
@EXIT