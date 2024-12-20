Sub AjustEmpenhoRestos()

'@Lang VBA

    Dim pasta As String
    Dim arquivo As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim cellValue As String
    Dim extractedValue As String
    Dim filePath As String
    
    pasta = "C:\Users\nome_do_usuario\Downloads\"
    arquivo = Dir(pasta)
    
    Do Until arquivo = ""
    
        If Right(arquivo, 5) = ".xlsx" Then
        
            Set wb = Workbooks.Open(pasta & arquivo)
            
            For Each ws In wb.Sheets
                cellValue = ws.Range("B4").Value
                
                ' Use uma express o regular para extrair as palavras "Restos" ou "Empenho"
                With CreateObject("VBScript.RegExp")
                    .Pattern = "Restos|Empenho"
                    If .Test(cellValue) Then
                        extractedValue = .Execute(cellValue)(0)
                        
                        ' Realizar as opera  es desejadas com os dados
                        Dim LR As Long
                        LR = ws.Cells(Rows.Count, 11).End(xlUp).Row
                        Selection.Copy
                        Range("A1").Select
                        ActiveSheet.Paste
                        Range("A1").NumberFormat = "#,##0.00"
                        Columns("A").AutoFit
                        
                        ' Defina o caminho e nome do arquivo
                        filePath = pasta & extractedValue & ".xlsx"
                        
                        ' Salve o arquivo com o novo nome
                        Application.DisplayAlerts = False
                        wb.SaveAs filePath, FileFormat:=xlOpenXMLWorkbook
                        wb.Close Savechanges:=True
                        Application.DisplayAlerts = True
                        
                        GoTo NextFile ' Ir para o pr ximo arquivo ap s salvar
                    End If
                End With
            Next ws
            
            ' Caso nenhum dos valores seja encontrado
            MsgBox "Nenhum dos valores 'Restos' ou 'Empenho' foi encontrado nas c lulas especificadas."
            wb.Close Savechanges:=False
            
        End If
        
NextFile:
        arquivo = Dir
        
    Loop
    
End Sub

