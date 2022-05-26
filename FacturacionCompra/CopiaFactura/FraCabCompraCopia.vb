Public Class FraCabCompraCopia
    Inherits FraCabCompra

    Public IDFactura As Integer

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(New DataRowPropertyAccessor(oRow))
        IDFactura = oRow("IDFactura")
    End Sub

End Class