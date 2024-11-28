Dim wb As Workbook

Sub main()

    RELAMPAGO
    PASTEVALUES

End Sub
Sub RELAMPAGO()
    Dim SourceRange As Range
    Dim TargetRange As Range
    Dim SourceCell As Range
    Dim TargetCell As Range
    Dim Char As String
    Dim i As Integer
    Dim RangeToClear As Range
    
    Set wb = Workbooks.Open("C:\Users\seu_usuario\Documents\1-Contas_Livres\Superávit Financeiro.xlsx")
    
    Set RangeToClear = Range("E2:E16")
    RangeToClear.ClearContents
        
    ' Definir os intervalos de origem e destino na planilha ativa
    Set SourceRange = ActiveSheet.Range("D2:D16")
    Set TargetRange = ActiveSheet.Range("E2:E16")
    
    ' Iterar através das células de origem e destino correspondentes
    For Each SourceCell In SourceRange
        Set TargetCell = TargetRange.Cells(SourceCell.Row - SourceRange.Row + 1, 1)
        
        Result = ""
        
        For i = 1 To Len(SourceCell.Value)
            Char = Mid(SourceCell.Value, i, 1)
            If IsNumeric(Char) Or Char = "," Then
                Result = Result & Char
            End If
        Next i
        
        If Result = "" Then
            Result = "0"
        End If
        
        TargetCell.Value = Result
    Next SourceCell
End Sub

Sub PASTEVALUES()
    Dim SourceRange As Range
    Dim Cell As Range
    
    ' Definir o intervalo de origem
    Set SourceRange = Range("E2:E16")
    
    ' Copiar o intervalo de origem
    SourceRange.Copy
    
    ' Colar valores no mesmo intervalo
    SourceRange.PasteSpecial Paste:=xlPasteValues
    
    ' Converter valores em números, célula por célula
    For Each Cell In SourceRange
        If IsNumeric(Cell.Value) Then
            Cell.Value = Cell.Value * 1
        End If
    Next Cell
    
    ' Limpar a área de transferência
    Application.CutCopyMode = False
    
    ' Salvar o arquivo
    wb.Save
    'Fechar o arquivo
    wb.Close SaveChanges:=True
End Sub



