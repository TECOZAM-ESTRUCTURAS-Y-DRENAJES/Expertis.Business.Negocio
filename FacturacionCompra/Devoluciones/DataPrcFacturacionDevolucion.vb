<Serializable()> _
   Public Class DataPrcFacturacionDevolucion
    '//Cabecera (en MonedaA)
    'Public IDFactura As Integer
    Public IDContador As String
    Public IDProveedor As String
    Public NFactura As String
    Public SuFactura As String
    Public FechaFactura As Date
    Public SuFechaFactura As Date?
    Public IDDevoluciones(-1) As Integer

    '//Linea de devolución por el importe de las comisiones
    Public IDArticulo As String
    Public CContable As String      '//CCCompra de Artículo o Cuenta introducida por el usuario
    Public IDTipoIVA As String      '//Tipo IVA del artículo o introducido por el usuario
    Public Precio As Double         '//Importe Comisiones
End Class
