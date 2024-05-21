Sub BB_TXT()

	pasta = "C:\Users\usuario\Documents\" ' inserir caminho correto
		
	arquivo = Dir(pasta)
    
    'Faça 'até' arquivo ser vazio
    'Enquanto não for vazio ficará sendo executado
    Do Until arquivo = ""
    
    Set wb = Workbooks.Open(pasta & arquivo)

    ChDir pasta
    ActiveWorkbook.SaveAs Filename:=pasta & arquivo & ".xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        
    Application.DisplayAlerts = False
    wb.Close Savechanges:=1
    
' Esse comando permite encontrar o próximo arquivo
arquivo = Dir()

Loop

End Sub

