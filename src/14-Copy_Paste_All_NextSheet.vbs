Sub NextSheet()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim ultimaColuna As Long
    Dim destino As Range
    Dim dados As Worksheet
    
    ' Define a aba "DADOS"
    Set dados = Sheets("DADOS")
    
    ' Limpa a aba "DADOS" antes de colar os dados
    dados.Cells.Clear
    
    ' Define a célula de destino inicial na aba "DADOS"
    Set destino = dados.Range("A1")
    
    ' Percorre todas as abas
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "DADOS" Then
            ' Encontra a última linha e coluna preenchida na aba atual
            ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            ultimaColuna = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            
            ' Copia os dados da aba atual
            ws.Range(ws.Cells(1, 1), ws.Cells(ultimaLinha, ultimaColuna)).Copy
            
            ' Cola os dados na aba "DADOS"
            destino.PasteSpecial Paste:=xlPasteValues
            destino.PasteSpecial Paste:=xlPasteFormats
            
            ' Atualiza a célula de destino para a próxima colagem
            Set destino = dados.Cells(dados.Rows.Count, 1).End(xlUp).Offset(2, 0)
        End If
    Next ws
    
    ' Remove a seleção de cópia
    Application.CutCopyMode = False
End Sub

