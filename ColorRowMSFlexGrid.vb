Dim index As Integer
If rsProduto!estoqueAtual <= 0 Then
   For index = 1 To .cols - 1
        .col = index
        .row = lRow
        .CellForeColor = vbWhite
        .CellBackColor = vbRed
   Next
End If
