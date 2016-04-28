Private Sub BtnSearch_Click()

    'Ejectua CopyDel_A hasta que Find_A = False
    Do
    CopyDel_A
    Loop Until Find_A

    MsgBox ("Proceso Finalizado")

End Sub

Private Function Find_A()

    'En la columna A, busca lo que hemos introducido en la celda F2
    Columns("A:A").Find (Range("F2"))

End Function

Private Function CopyDel_A()

    '1) En la columna A, busca lo que hemos introducido en la celda F2
    Columns("A:A").Select
    Selection.Find(What:=Range("F2"), After:=Range("A1"), _
        LookAt:=xlWhole, SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, MatchCase:=True, _
        SearchFormat:=False).Activate

    '2) Copia +1row y +3col desde celda activa
    Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(0, 3)).Copy
    Range("I2:L2").PasteSpecial

    '3) Añadimos fila nueva
    Range("I2:L2").Select
    Selection.Insert Shift:=xlShiftDown
    Selection.Insert Shift:=xlShiftDown

    '4) Repeat de 1)
    Columns("A:A").Select
        Selection.Find(What:=Range("F2"), After:=Range("A1"), _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=True, SearchFormat:=False).Activate

    '5) Elimina +1row y +3col desde celda activa
    Range(ActiveCell.Offset(1, 0), ActiveCell.Offset(0, 3)).Delete Shift:=xlShiftUp

End Function
