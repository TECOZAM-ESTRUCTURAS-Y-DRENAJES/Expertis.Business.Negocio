Public Class PedCabCompraProgramaCompra
    Inherits PedCabCompra

    Public IDAlmacen As String
    Public IDPedido As Integer? ' Vendrá relleno en algunas ocasiones
    Public IDOperario As String
    Public IDFormaEnvio As String
    Public IDCondicionEnvio As String

    Public Overrides Function PrimaryKeyCabOrigen() As String
        Return "IDPrograma"
    End Function
    Public Overrides Function FieldNOrigen() As String
        Return String.Empty
    End Function
    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)

        MyBase.ViewName = "vfrmConfirmacionProgramaCompra"
        MyBase.Origen = enumOrigenPedidoCompra.Programa
        IDAlmacen = oRow("IDAlmacen")
        If Length(oRow("IDPedido")) > 0 Then IDPedido = oRow("IDPedido") ' Vendrá relleno en algunas ocasiones
        If Length(oRow("IDOperario")) > 0 Then IDOperario = oRow("IDOperario")
        If Length(oRow("IDFormaEnvio")) > 0 Then IDFormaEnvio = oRow("IDFormaEnvio")
        If Length(oRow("IDCondicionEnvio")) > 0 Then IDCondicionEnvio = oRow("IDCondicionEnvio")
    End Sub

    Public Sub Add(ByVal lin As PedLinCompraProgramaCompra)
        MyBase.Add(lin)
    End Sub

End Class

