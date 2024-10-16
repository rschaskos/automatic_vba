Sub Main()
'@Lang VBA

    CriarAba
    NextSheet
    Delete
    CopyDocumentoRows
    CopyCombustivelRows
    
End Sub

Sub CriarAba()
    Dim novaAba As Worksheet

    ' Cria uma nova aba chamada "DADOS"
    Set novaAba = ThisWorkbook.Worksheets.Add
    novaAba.Name = "DADOS"

    ' Move a aba "DADOS" para ser a primeira aba
    novaAba.Move Before:=ThisWorkbook.Worksheets(1)
End Sub

Sub NextSheet()

'@Lang VBA

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

Sub Delete()
    Dim ws As Worksheet
    Application.DisplayAlerts = False ' Desativa os alertas para evitar confirmações de exclusão

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "DADOS" Then
            ws.Delete
        End If
    Next ws

    Application.DisplayAlerts = True ' Reativa os alertas
End Sub

Sub CopyDocumentoRows()
    Dim ws As Worksheet
    Dim cell As Range
    Dim targetRow As Long
    Dim lastCol As Long
    
    Set ws = ThisWorkbook.Sheets("DADOS")
    
    targetRow = 1
    
    For Each cell In ws.Range("A:A")
        If InStr(1, cell.Value, "Documento:") > 0 Then
            lastCol = ws.Cells(cell.Row, ws.Columns.Count).End(xlToLeft).Column
            ws.Range(ws.Cells(cell.Row, 1), ws.Cells(cell.Row, lastCol)).Copy
            ws.Cells(targetRow, 15).PasteSpecial Paste:=xlPasteValues
            targetRow = targetRow + 1
        End If
    Next cell
    
    Application.CutCopyMode = False
End Sub

Sub CopyCombustivelRows()
    Dim ws As Worksheet
    Dim cell As Range
    Dim targetRow As Long
    Dim lastCol As Long
    
    Set ws = ThisWorkbook.Sheets("DADOS")
    
    targetRow = 1
    
    For Each cell In ws.Range("A:A")
        If InStr(1, cell.Value, "Combustível:") > 0 Then
            lastCol = ws.Cells(cell.Row, ws.Columns.Count).End(xlToLeft).Column
            ws.Range(ws.Cells(cell.Row, 1), ws.Cells(cell.Row, lastCol)).Copy
            ws.Cells(targetRow, 29).PasteSpecial Paste:=xlPasteValues
            targetRow = targetRow + 1
        End If
    Next cell
    
    Application.CutCopyMode = False
    ws.Range(ws.Columns(43), ws.Columns(55)).Delete
End Sub

