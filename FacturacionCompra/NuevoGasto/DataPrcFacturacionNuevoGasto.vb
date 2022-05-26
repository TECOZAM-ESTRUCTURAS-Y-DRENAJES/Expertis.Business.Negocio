<Serializable()> _
Public Class DataPrcFacturacionNuevoGasto
    Public FechaFactura As Date
    Public SuFechaFactura As Date
    Public SuFactura As String

    Public IDContador As String
    Public IDProveedor As String
    Public RazonSocial As String
    Public CIF As String
    Public IDMoneda As String
    Public IDDiaPago As String
    Public IDFormaPago As String
    Public IDCondicionPago As String
    Public IDBancoPropio As String
    Public IDTipoAsiento As Integer

    Public Lineas As New List(Of DataPrcFacturacionLineaNuevoGasto)
   
End Class


<Serializable()> _
Public Class DataPrcFacturacionLineaNuevoGasto
    Public IDArticulo As String
    Public DescArticulo As String
    Public RefProveedor As String
    Public DescRefProveedor As String
    Public IDCContable As String
    Public IDTipoIVA As String
    Public Importe As Double
    Public ImpIVA As Double
    Public ImporteTotal As Double
    Public IDDireccion As Integer
    Public Direccion As String
    Public CodPostal As String
    Public Poblacion As String
    Public Provincia As String
    Public IDPais As String
    Public IDObra As Integer?
    Public IDTrabajo As Integer?
    Public Analitica As DataTable
End Class
