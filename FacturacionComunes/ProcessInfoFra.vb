'TODO informacion del proceso
Public Class ProcessInfoFra
    Inherits ProcessInfo

    Public IDTPV As String
    Public TipoLineaDef As String
    Public SuFactura As String
    Public SuFechaFactura As Date?
    Public ConPropuesta As Boolean = True
    Public IDPisosPagos As String
    Public IDFacturaDocuwares As String

    Public Sub New(ByVal IDContador As String, _
                   ByVal TipoLineaDef As String, _
                   Optional ByVal SuFactura As String = Nothing, _
                   Optional ByVal IDTPV As String = Nothing, _
                   Optional ByVal SuFechaFactura As Date = cnMinDate, _
                   Optional ByVal ConPropuesta As Boolean = True, _
                   Optional ByVal IDPisosPagos As String = Nothing)
        MyBase.New(IDContador)
        Me.TipoLineaDef = TipoLineaDef
        Me.SuFactura = SuFactura
        Me.IDTPV = IDTPV
        If SuFechaFactura <> cnMinDate Then Me.SuFechaFactura = SuFechaFactura
        Me.ConPropuesta = ConPropuesta
        Try
            Me.IDPisosPagos = IDPisosPagos
        Catch ex As Exception
        End Try
        Try
            Me.IDFacturaDocuwares = IDFacturaDocuwares
        Catch ex As Exception
        End Try
    End Sub
    'Public Sub New(ByVal IDContador As String, _
    '           ByVal TipoLineaDef As String, _
    '           Optional ByVal SuFactura As String = Nothing, _
    '           Optional ByVal IDTPV As String = Nothing, _
    '           Optional ByVal SuFechaFactura As Date = cnMinDate, _
    '           Optional ByVal ConPropuesta As Boolean = True, _
    '           Optional ByVal IDPisoPago As String = Nothing)
    '    MyBase.New(IDContador)
    '    Me.TipoLineaDef = TipoLineaDef
    '    Me.SuFactura = SuFactura
    '    Me.IDTPV = IDTPV
    '    If SuFechaFactura <> cnMinDate Then Me.SuFechaFactura = SuFechaFactura
    '    Me.ConPropuesta = ConPropuesta
    '    Me.IDClaveRegimenEspecial2= IDPisoPago
    'End Sub

    Public Sub New()
        Me.TipoLineaDef = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, New ServiceProvider)
    End Sub
End Class
