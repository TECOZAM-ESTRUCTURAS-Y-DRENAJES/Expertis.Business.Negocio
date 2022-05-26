<Serializable()> _
Public Class DataCalculoTarifaComercial
    Public IDTarifa As String           '//Desde el TPV se indicará la tarifa a tener en cuenta
    Public IDArticulo As String
    Public IDCliente As String
    Public Cantidad As Double
    Public CantidadAnterior As Double
    Public Fecha As Date?
    Public IDMoneda As String           '//Moneda del contexto (dato de entrada). Se utiliza entre otras cosas, para devolver la Tarifa en esta Moneda
    Public IDUDMedida As String         '//Medida de la línea (dato de entrada)
    'Public DtoComercialLinea As Double
    ''Friend IDOrdenRuta As Integer?
    Public IDAlmacen As String

    Public IDTipoIVA As String
    Public DebeSerPVP As Boolean
    Public EsRegalo As Boolean
    Public IDPromocion As String
    Public IDPromocionLinea As Integer
    Public UDValoracion As Double?

    Public DatosTarifa As DataTarifaComercial

    Public Sub New()
        Me.DatosTarifa = New DataTarifaComercial
    End Sub

    Public Sub New(ByVal IDArticulo As String, ByVal Cantidad As Double, ByVal Fecha As Date)
        Me.DatosTarifa = New DataTarifaComercial
        Me.IDArticulo = IDArticulo
        Me.Cantidad = Cantidad
        Me.Fecha = Fecha
    End Sub

    Public Sub New(ByVal IDArticulo As String, ByVal Cantidad As Double, ByVal Fecha As Date, ByVal IDAlmacen As String)
        Me.DatosTarifa = New DataTarifaComercial
        Me.IDArticulo = IDArticulo
        Me.Cantidad = Cantidad
        Me.Fecha = Fecha
        Me.IDAlmacen = IDAlmacen
    End Sub

    Public Sub New(ByVal IDArticulo As String, ByVal IDCliente As String, ByVal Cantidad As Double, ByVal Fecha As Date)
        Me.DatosTarifa = New DataTarifaComercial
        Me.IDCliente = IDCliente
        Me.IDArticulo = IDArticulo
        Me.Cantidad = Cantidad
        Me.Fecha = Fecha
    End Sub

End Class
