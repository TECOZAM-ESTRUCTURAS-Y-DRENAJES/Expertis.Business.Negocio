'//Clase q proporciona datos de entrada al proceso
<Serializable()> _
Public Class DataPrcAlbaranar
    Public AlbVentaInfo() As CrearAlbaranVentaInfo
    Public IDContador As String
    Public FechaAlbaran As Date
    Public IDTipoAlbaran As String
    Public TipoExpedicion As enumTipoExpedicion
    Public NObra As String
    Public bandera As Boolean


    Public Sub New(ByVal AVInfo() As CrearAlbaranVentaInfo, Optional ByVal IDContador As String = Nothing, Optional ByVal FechaAlbaran As Date = cnMinDate, Optional ByVal IDTipoAlbaran As String = Nothing, Optional ByVal TipoExpedicion As enumTipoExpedicion = enumTipoExpedicion.tePedido, Optional ByVal NObra As String = Nothing, Optional ByVal bandera As Boolean = True)
        Me.AlbVentaInfo = AVInfo
        Me.IDContador = IDContador
        Me.FechaAlbaran = FechaAlbaran
        Me.IDTipoAlbaran = IDTipoAlbaran
        Me.TipoExpedicion = TipoExpedicion
        Me.NObra = NObra
        Me.bandera = bandera
    End Sub

End Class
