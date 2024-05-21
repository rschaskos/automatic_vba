Sub frota()

Dim i As Integer
Dim LR As Long
Dim qtde As Long

'renomear abas

Sheets("page 1").Name = "plan1"
Sheets.Add After:=ActiveSheet
Sheets("Planilha1").Select
Sheets("Planilha1").Name = "plan2"

'cabeçalho
    
Sheets("plan2").Select
Range("A1") = "DATA"
Range("B1") = "DOCUMENTO"
Range("C1") = "EMPENHO"
Range("D1") = "QUANTIDADE"
Columns("A:D").AutoFit


Sheets("plan1").Select

'limpa veículo

    'With Sheets("plan1")
        'qtde = WorksheetFunction.CountIf(.Range("A:BX"), "Veículo:")
        'MsgBox qtde
    'End With

For i = 1 To 90
            Cells.Find(What:="veículo:", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
   Row = ActiveCell.Row
   Rows(Row).EntireRow.Delete
Next i

            'Cells.Replace What:="Veículo", Replacement:="-", LookAt:=xlPart, _
        'SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        'ReplaceFormat:=False

    'With Sheets("plan1")
        'qtde = WorksheetFunction.CountIf(.Range("A:BX"), "Quantidade:")
        'MsgBox qtde
    'End With

Row = 10

For i = 1 To 400

Sheets("plan1").Select
Range("A" & Row).Select

    
'data
            Cells.Find(What:="Data:", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
    Row = ActiveCell.Row
    Col = ActiveCell.Column

    Cells(Row, Col + 2).Select
    Selection.Copy
    Sheets("plan2").Select
    LR = Cells(Rows.Count, 1).End(xlUp).Row
    Range("A" & LR + 1).PasteSpecial
    Range("A" & LR + 1).Select
    Selection.UnMerge
    Sheets("plan1").Select
    Cells(Row - 1, Col).Select
            
            Cells.Find(What:="Data:", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
            ActiveCell.Replace What:="Data:", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False


    
    
'documento
            Cells.Find(What:="Documento:", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
    Row = ActiveCell.Row
    Col = ActiveCell.Column
    Cells(Row, Col + 8).Select
    Selection.Copy
    Sheets("plan2").Select
    LR = Cells(Rows.Count, 2).End(xlUp).Row
    Range("B" & LR + 1).PasteSpecial
    Range("B" & LR + 1).Select
    Selection.UnMerge
    Sheets("plan1").Select
    Cells(Row - 1, Col).Select
    
            Cells.Find(What:="Documento:", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
            ActiveCell.Replace What:="Documento:", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    
'empenho
            Cells.Find(What:="Núm.Emp.:", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
    Row = ActiveCell.Row
    Col = ActiveCell.Column
    Cells(Row, Col + 11).Select
    Selection.Copy
    Sheets("plan2").Select
    LR = Cells(Rows.Count, 3).End(xlUp).Row
    Range("C" & LR + 1).PasteSpecial
    Range("C" & LR + 1).Select
    Selection.UnMerge
    Sheets("plan1").Select
    Cells(Row - 1, Col).Select
    
            Cells.Find(What:="Núm.Emp.:", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
            ActiveCell.Replace What:="Núm.Emp.:", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
'quantidade
            Cells.Find(What:="Quantidade:", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
    Row = ActiveCell.Row
    Col = ActiveCell.Column
    Cells(Row, Col + 13).Select
    Selection.Copy
    Sheets("plan2").Select
    LR = Cells(Rows.Count, 4).End(xlUp).Row
    Range("D" & LR + 1).PasteSpecial
    Range("D" & LR + 1).Select
    Selection.UnMerge
    Sheets("plan1").Select
    Cells(Row - 1, Col).Select
    
            Cells.Find(What:="Quantidade:", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
            ActiveCell.Replace What:="Quantidade:", Replacement:="-", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
Next i
    
End Sub
    


