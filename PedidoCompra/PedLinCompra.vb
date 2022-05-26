
Public MustInherit Class PedLinCompra

    Public MustOverride Function PrimaryKeyLinOrigen() As String

    Public IDLineaOrigen As Integer
    Public Cantidad As Double
    Public QConfirmada As Double

    Public Sub New(ByVal oRow As DataRow)
        If Length(PrimaryKeyLinOrigen) > 0 Then IDLineaOrigen = oRow(PrimaryKeyLinOrigen)
        Cantidad = Double.NaN
        QConfirmada = Double.NaN
    End Sub

End Class
