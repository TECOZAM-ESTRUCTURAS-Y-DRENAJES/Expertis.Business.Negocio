'TODO informacion del proceso
Public Class ProcessInfoFra
    Inherits ProcessInfo

    Public IDTPV As String
    Public TipoLineaDef As String
    Public SuFactura As String
    Public SuFechaFactura As Date?
    Public ConPropuesta As Boolean = True

    Public Sub New(ByVal IDContador As String, _
                   ByVal TipoLineaDef As String, _
                   Optional ByVal SuFactura As String = Nothing, _
                   Optional ByVal IDTPV As String = Nothing, _
                   Optional ByVal SuFechaFactura As Date = cnMinDate, _
                   Optional ByVal ConPropuesta As Boolean = True)
        MyBase.New(IDContador)
        Me.TipoLineaDef = TipoLineaDef
        Me.SuFactura = SuFactura
        Me.IDTPV = IDTPV
        If SuFechaFactura <> cnMinDate Then Me.SuFechaFactura = SuFechaFactura
        Me.ConPropuesta = ConPropuesta
    End Sub

    Public Sub New()
        Me.TipoLineaDef = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, New ServiceProvider)
    End Sub
End Class
