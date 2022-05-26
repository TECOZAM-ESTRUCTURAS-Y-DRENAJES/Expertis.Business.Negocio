Public Class PedCabOfertaComercial
    Inherits PedCab

    Public IDOferta As Integer
    Public NOferta As String

    Public FechaEntrega As Date
    'Public IDLineaOfertaDetalle As Integer?
    Public IDDireccionCliente As Integer

    Public LineaOfertaDetalle(-1) As Integer

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)
        If Length(oRow("IDCliente")) > 0 Then Me.IDCliente = oRow("IDCliente")
        If Length(Me.IDCliente) = 0 AndAlso Length(oRow("IDEmpresa")) > 0 Then Me.IDCliente = oRow("IDEmpresa")

        IDOferta = oRow("IDOfertaComercial")
        NOferta = oRow("NumOferta")
        ' If Length(oRow("IDLineaOfertaDetalle")) > 0 Then IDLineaOfertaDetalle = oRow("IDLineaOfertaDetalle")

        Fecha = Today
        Me.Origen = enumOrigenPedido.Oferta
        Me.IDDireccionCliente = oRow("IDDireccionCliente")
        If Length(oRow("IDDireccionCliente")) > 0 Then IDDireccionCliente = oRow("IDDireccionCliente")
    End Sub

    Public Sub Add(ByVal oRow As DataRow)
        ReDim Preserve LineaOfertaDetalle(LineaOfertaDetalle.Length)
        LineaOfertaDetalle(LineaOfertaDetalle.Length - 1) = oRow("IDLineaOfertaDetalle")
    End Sub

End Class
