Sub cadCnpj()

'@Lang VBA

	pasta = "C:\Users\usuario\Documents\" ' inserir caminho correto

	arquivo = Dir(pasta)

	Do Until arquivo = ""

        Set wb = Workbooks.Open(pasta & arquivo)

'PARTE 1
        Cells.Find(What:="NÚMERO DE", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
    Row = ActiveCell.Row
    Col = ActiveCell.Column
    Cells(Row + 1, Col).Select
    Selection.Copy
    Range("A1").Select
    ActiveSheet.Paste
    Cells(Row, Col).Delete
    
'PARTE 2
        Cells.Find(What:="NOME EMPRESARIAL", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
    Row = ActiveCell.Row
    Col = ActiveCell.Column
    Cells(Row + 1, Col).Select
    Selection.Copy
    Range("B1").Select
    ActiveSheet.Paste
    
'PARTE 3
        Cells.Find(What:="LOGRADOURO", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
    Row = ActiveCell.Row
    Col = ActiveCell.Column
    Cells(Row + 1, Col).Select
    Selection.Copy
    Range("C1").Select
    ActiveSheet.Paste
    
'PARTE 4

        Cells.Find(What:="NÚMERO", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
    Row = ActiveCell.Row
    Col = ActiveCell.Column
    Cells(Row + 1, Col).Select
    Selection.Copy
    Range("D1").Select
    ActiveSheet.Paste
    
'PARTE 5

        Cells.Find(What:="CEP", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
    Row = ActiveCell.Row
    Col = ActiveCell.Column
    Cells(Row + 1, Col).Select
    Selection.Copy
    Range("E1").Select
    ActiveSheet.Paste
    
'PARTE 6

        Cells.Find(What:="BAIRRO", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate

    Row = ActiveCell.Row
    Col = ActiveCell.Column
    Cells(Row + 1, Col).Select
    Selection.Copy
    Range("F1").Select
    ActiveSheet.Paste
    
'PARTE 7

        Cells.Find(What:="MUNICÍPIO", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
    Row = ActiveCell.Row
    Col = ActiveCell.Column
    Cells(Row + 1, Col).Select
    Selection.Copy
    Range("G1").Select
    ActiveSheet.Paste
    
'PARTE 8

        Cells.Find(What:="UF", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
        
   Row = ActiveCell.Row
   Col = ActiveCell.Column
   Cells(Row + 1, Col).Select
   Selection.Copy
   Range("H1").Select
   ActiveSheet.Paste
    
'PARTE 9

        Cells.Find(What:="ENDEREÇO ELETRÔNICO", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
    
  Row = ActiveCell.Row
  Col = ActiveCell.Column
  Cells(Row + 1, Col).Select
  Selection.Copy
  Range("I1").Select
  ActiveSheet.Paste
  
'PARTE 10

        Cells.Find(What:="TELEFONE", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
  
  Row = ActiveCell.Row
  Col = ActiveCell.Column
  Cells(Row + 1, Col).Select
  Selection.Copy
  Range("J1").Select
  ActiveSheet.Paste
  
    Application.DisplayAlerts = False
    wb.Close Savechanges:=1
  
arquivo = Dir()
        
Loop

End Sub


