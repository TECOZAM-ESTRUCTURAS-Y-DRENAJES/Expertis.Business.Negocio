<Serializable()> _
Public Class DataCalculoTarifaCompra
    Public IDArticulo As String
    Public IDProveedor As String
    Public Cantidad As Double
    Public Fecha As Date?
    Public IDMoneda As String           '//Moneda del contexto (dato de entrada). Se utiliza entre otras cosas, para devolver la Tarifa en esta Moneda
    Public IDUDMedida As String         '//Medida de la línea (dato de entrada)
    'Public IDOrdenRuta As Integer?
    Public UDValoracion As Integer?

    Public DatosTarifa As DataTarifaCompra

    Public Sub New()

    End Sub

    Public Sub New(ByVal IDArticulo As String, ByVal IDProveedor As String, ByVal Cantidad As Double)
        Me.IDArticulo = IDArticulo
        Me.IDProveedor = IDProveedor
        Me.Cantidad = Cantidad
    End Sub

End Class
