<Serializable()> _
Public Class DataTarifaCompra
    Inherits DataTarifa

    Public IDUDCompra As String

    Public Referencia As String
    Public DescReferencia As String

    Public IDContrato As String
    Public IDLineaContrato As Integer?
End Class