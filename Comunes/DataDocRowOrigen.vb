Public Class DataDocRowOrigen
    Public Doc As DocumentCabLin
    Public RowOrigen As DataRow
    Public RowDestino As DataRow

    Public Sub New(ByVal Doc As DocumentCabLin, ByVal RowOrigen As DataRow, ByVal RowDestino As DataRow)
        Me.Doc = Doc
        Me.RowOrigen = RowOrigen
        Me.RowDestino = RowDestino
    End Sub

End Class
