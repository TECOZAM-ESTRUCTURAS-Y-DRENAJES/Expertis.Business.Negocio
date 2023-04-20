Public MustInherit Class PedCabCompra
    Inherits CompraCab

    Public MustOverride Function PrimaryKeyCabOrigen() As String
    Public MustOverride Function FieldNOrigen() As String

    Public Origen As enumOrigenPedidoCompra
    Public IDOrigen As String  'Es String, por que algunos origenes son string
    Public NOrigen As String
    Public LineasOrigen(-1) As PedLinCompra

    Public ViewName As String
    'Public Agrupacion As Boolean

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(New DataRowPropertyAccessor(oRow))
        IDOrigen = oRow(PrimaryKeyCabOrigen)
        If Length(FieldNOrigen) > 0 Then NOrigen = oRow(FieldNOrigen) & String.Empty
    End Sub


    Public Sub Add(ByVal lin As PedLinCompra)
        ReDim Preserve LineasOrigen(LineasOrigen.Length)
        LineasOrigen(LineasOrigen.Length - 1) = lin
    End Sub

End Class
