Public Class PedCabPrograma
    Inherits PedCab

    Public IDPrograma As String
    Public IDAlmacen As String
    ' Public PedidoCliente As String
    Public IDPedido As Integer ' Vendrá relleno en algunas ocasiones
    Public IDDireccionEnvio As Integer
    Public Texto As String

    Public Lineas(-1) As PedLinPrograma

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)

        IDPrograma = oRow("IDPrograma")
        IDAlmacen = oRow("IDAlmacen")
        If Length(oRow("ProgramaCliente")) > 0 Then PedidoCliente = oRow("ProgramaCliente")
        If Length(oRow("IDPedido")) > 0 Then IDPedido = oRow("IDPedido") ' Vendrá relleno en algunas ocasiones

        '    Fecha = oRow("FechaConfirmacionNew")
        Me.Origen = enumOrigenPedido.Programa
        Me.IDDireccionEnvio = oRow("IDDireccionEnvio")
        If oRow.Table.Columns.Contains("Texto") Then Me.Texto = oRow("Texto") & String.Empty
    End Sub

    Public Sub Add(ByVal lin As PedLinPrograma)
        ReDim Preserve Lineas(Lineas.Length)
        Lineas(Lineas.Length - 1) = lin
    End Sub

End Class
