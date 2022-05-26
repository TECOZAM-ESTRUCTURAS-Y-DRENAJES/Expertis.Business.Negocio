Public Class PedCabCompraCopia
    Inherits PedCabCompra

    Public IDPedido As Integer

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)
        IDPedido = oRow(PrimaryKeyCabOrigen)
        If Length(FieldNOrigen) > 0 Then NOrigen = oRow(FieldNOrigen) & String.Empty
    End Sub

    Public Overrides Function FieldNOrigen() As String
        Return "NPedido"
    End Function

    Public Overrides Function PrimaryKeyCabOrigen() As String
        Return "IDPedido"
    End Function

End Class