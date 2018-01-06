Attribute VB_Name = "Geral"
Public Sub Desfazer()
    
    Application.SendKeys ("^(z)")
    
End Sub

Public Sub Refazer()
    
    Application.SendKeys ("^(y)")
    
    
End Sub

Public Sub Salvar()
    
    ThisWorkbook.Save
    
End Sub

Public Sub FecharArquivo()

    ActiveWorkbook.Close
    ActiveWorkbook.Save

End Sub
