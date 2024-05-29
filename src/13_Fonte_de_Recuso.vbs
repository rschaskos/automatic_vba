Option Explicit
Global C As Integer
Global R As Integer
Sub find_value()

    Do While IsEmpty(ActiveCell.Value)
        ActiveCell.Offset(0, 1).Select
    Loop
    
End Sub

Sub take_value()

    With ActiveCell
        C = .Column
        R = .Row
    End With

End Sub

Sub Main()

Dim ws As Worksheet
Dim pasta As String
Dim arquivo As String
Dim wb As Workbook
Dim Rng As Range


pasta = "C:\Extratos\Fontes\" 'nome caminho desejado

arquivo = Dir(pasta)

Do Until arquivo = ""

    Set wb = Workbooks.Open(pasta & arquivo)
    
    ' SALDO FIN. INICIAL AJUSTADO
    
            Cells.Find(What:="Saldo fin. inicial ajustado:", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        True, SearchFormat:=False).Activate
    take_value
    Set ws = wb.Sheets("Sheet1")
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("A1").Select
    ActiveSheet.Paste
    Set Rng = ws.Cells(R, C)
    Rng.Select
    find_value
    take_value
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("B1").Select
    ActiveSheet.Paste

' RECEITA ORCAMENTARIA
    
            Cells.Find(What:="Receita orçamentária:", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
    :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    True, SearchFormat:=False).Activate
    
    take_value
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("A2").Select
    ActiveSheet.Paste
    Set Rng = ws.Cells(R, C)
    Rng.Select
    find_value
    take_value
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("B2").Select
    ActiveSheet.Paste
        
' INSCRICAO EM CONSIGNACAO
    
            Cells.Find(What:="Inscrição de consignação:", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
    :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    True, SearchFormat:=False).Activate
    
    take_value
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("A3").Select
    ActiveSheet.Paste
    Set Rng = ws.Cells(R, C)
    Rng.Select
    find_value
    take_value
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("B3").Select
    ActiveSheet.Paste
    
' INGRESSO

            Cells.Find(What:="Ingresso:", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
    :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    True, SearchFormat:=False).Activate
    
    take_value
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("A4").Select
    ActiveSheet.Paste
    Set Rng = ws.Cells(R, C)
    Rng.Select
    find_value
    take_value
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("B4").Select
    ActiveSheet.Paste
    
' BAIXA DE CONSIGNACAO
    
            Cells.Find(What:="Baixa de consignação:", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
    :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    True, SearchFormat:=False).Activate
    
    take_value
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("A5").Select
    ActiveSheet.Paste
    Set Rng = ws.Cells(R, C)
    Rng.Select
    find_value
    take_value
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("B5").Select
    ActiveSheet.Paste
    
' BAIXA REALIZAVEL POR
    
            Cells.Find(What:="Baixa realizável por", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
    :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    True, SearchFormat:=False).Activate
    
    take_value
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("A5").Select
    ActiveSheet.Paste
    Set Rng = ws.Cells(R, C)
    Rng.Select
    find_value
    take_value
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("B5").Select
    ActiveSheet.Paste
    
' EGRESSO
    
            Cells.Find(What:="Egresso:", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
    :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    True, SearchFormat:=False).Activate
    
    take_value
    
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("A6").Select
    ActiveSheet.Paste
    Set Rng = ws.Cells(R, C)
    Rng.Select
    find_value
    take_value
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("B6").Select
    ActiveSheet.Paste

' DESPESA ORCAMENTARIA
    
            Cells.Find(What:="Despesa orçamentária:", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
    :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    True, SearchFormat:=False).Activate
    
    take_value
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("A7").Select
    ActiveSheet.Paste
    Set Rng = ws.Cells(R, C)
    Rng.Select
    find_value
    take_value
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("B7").Select
    ActiveSheet.Paste

' RESTOS A PAGAR INSCRITO
    
            Cells.Find(What:="Restos a pagar:", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
    :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    True, SearchFormat:=False).Activate
    
    take_value
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("A8").Select
    ActiveSheet.Paste
    Set Rng = ws.Cells(R, C)
    Rng.Select
    find_value
    take_value
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("B8").Select
    ActiveSheet.Paste
    
' RESULTADO DO AJUSTE FINAL
    
            Cells.Find(What:="Resultado do ajuste final:", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
    :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    True, SearchFormat:=False).Activate
    
    take_value
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("A9").Select
    ActiveSheet.Paste
    Set Rng = ws.Cells(R, C)
    Rng.Select
    find_value
    take_value
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("B9").Select
    ActiveSheet.Paste
    
' EMPENHOS A PAGAR
    
            Cells.Find(What:="Saldo a pagar:", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
    :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    True, SearchFormat:=False).Activate
    
    take_value
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("A10").Select
    ActiveSheet.Paste
    Set Rng = ws.Cells(R, C)
    Rng.Select
    find_value
    take_value
    Set Rng = ws.Cells(R, C)
    Rng.Select
    Selection.Cut
    Range("B10").Select
    ActiveSheet.Paste
      
      
' EMPENHOS A PAGAR LIQUIDO

    Range("A11").Value = "Restos a pagar Saldo"
    
     Cells.Find(What:="Inscritos", After:=ActiveCell, LookIn:=xlFormulas, LookAt _
    :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
    True, SearchFormat:=False).Activate

    take_value
    Cells(R, C).End(xlToLeft).Select
    ActiveCell.Offset(1, 0).Select
    find_value
    take_value
    Selection.Cut
    Range("B11").Select
    ActiveSheet.Paste
    Cells(R, C).Select
    
    ' parte 2
    find_value
    take_value
    Selection.Cut
    Range("C11").Select
    ActiveSheet.Paste
    Cells(R, C).Select
    
    ' parte 3
    find_value
    take_value
    Selection.Cut
    Range("D11").Select
    ActiveSheet.Paste
    Cells(R, C).Select
    
    'parte 3
    find_value
    take_value
    Selection.Cut
    Range("E11").Select
    ActiveSheet.Paste
    Cells(R, C).Select
    
    ' parte 4
    find_value
    take_value
    Selection.Cut
    Range("F11").Select
    ActiveSheet.Paste
    Cells(R, C).Select
    
    ' parte 5
    Dim counter As Integer
    counter = 1
    Do While IsEmpty(ActiveCell.Value)
        ActiveCell.Offset(0, 1).Select
        If counter = 5 Then
            Exit Do
        End If
        counter = counter + 1
    Loop
    take_value
    Selection.Cut
    Range("G11").Select
    ActiveSheet.Paste
    
    Columns("A:A").EntireColumn.AutoFit
    
    Application.DisplayAlerts = False
    wb.Close Savechanges:=1
    
arquivo = Dir()

Loop
End Sub


