Public Class FraCabObra
    Inherits FraCab

    Public IDTrabajo As Integer
    Public DescTrabajo As String
    Public IDDiaPago As String
    'Public RetencionIRPF As Double
    Public IDContador As String
    Public ClienteGenerico As Boolean 'Alquiles
    Public TipoMnto As Integer = enumTipoObra.tpObra
    Public Retencion As Integer
    Public TipoRetencion As Integer
    Public FechaRetencion As Date
    Public Periodo As Integer
    Public TipoPeriodo As Integer
    Public NObra As String
    Public CambioA As Double
    Public CambioB As Double
    Public SeguroCambio As Boolean

    Public AgrupacionObra As enummcAgrupFacturaObra
    Public NumeroPedido As String   '//PedidoCliente  (Agrupaciones por ObraPedidoCliente)

    Public Lineas(-1) As FraLinObra

    Public Sub New(ByVal oRow As IPropertyAccessor)
        MyBase.New(oRow)
        IDTrabajo = oRow("IDTrabajo")
        DescTrabajo = Nz(oRow("DescTrabajo"), String.Empty)
        IDDiaPago = oRow("IDDiaPago") & String.Empty
        ' RetencionIRPF = oRow("RetencionIRPF")
        IDContador = IIf(Length(oRow("IDContadorCargo")) > 0, oRow("IDContadorCargo"), IDContador)
        If Length(oRow("ClienteGenerico")) > 0 Then ClienteGenerico = oRow("ClienteGenerico")
        Retencion = Nz(oRow("Retencion"), 0)
        TipoRetencion = Nz(oRow("TipoRetencion"), 0)
        FechaRetencion = Nz(oRow("FechaRetencion"), cnMinDate)
        Periodo = Nz(oRow("Periodo"), 0)
        TipoPeriodo = Nz(oRow("TipoPeriodo"), 0)
        AgrupacionObra = oRow("AgrupFacturaObra")
        If Length(oRow("NObra")) > 0 Then NObra = oRow("NObra")
        If Length(oRow("NumeroPedido")) > 0 Then NObra = oRow("NumeroPedido")
        Me.SeguroCambio = Nz(oRow("SeguroCambio"), False)
        Me.CambioA = Nz(oRow("CambioA"), 0)
        Me.CambioB = Nz(oRow("CambioB"), 0)
    End Sub

    Public Sub Add(ByVal lin As FraLinObra)
        ReDim Preserve Lineas(Lineas.Length)
        Lineas(Lineas.Length - 1) = lin
    End Sub

End Class