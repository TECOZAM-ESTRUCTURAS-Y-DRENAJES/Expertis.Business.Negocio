'Relación de facturas resultado del proceso de facturación
<Serializable()> _
Public Class ResultFacturacion
    Public PropuestaFacturas As DataTable       '//Para la Propuesta
    Public Log As LogProcess                    '//Para hacer el Log de todo el proceso (Propuesta + CrearFactura)

    Public Sub New()
        Log = New LogProcess
    End Sub

    Public Sub New(ByVal n As Integer, ByVal id As Integer, ByVal LogProc As LogProcess)
        Log = LogProc
    End Sub

    Public Sub New(ByVal dtFacturas As DataTable)
        PropuestaFacturas = dtFacturas
        Log = New LogProcess
    End Sub

End Class
