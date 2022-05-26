'//Clase q proporciona datos de entrada al proceso
'David Velasco
<Serializable()> _
Public Class DataPrcAlbaranarDeposito
    Public AlbVentaInfo() As CrearAlbaranVentaInfo
    Public IDContador As String
    Public FechaAlbaran As Date
    Public IDTipoAlbaran As String
    Public TipoExpedicion As enumTipoExpedicion
    Public NObra As String

    Public Sub New(ByVal AVInfo() As CrearAlbaranVentaInfo, Optional ByVal IDContador As String = Nothing, Optional ByVal FechaAlbaran As Date = cnMinDate, Optional ByVal IDTipoAlbaran As String = Nothing, Optional ByVal TipoExpedicion As enumTipoExpedicion = enumTipoExpedicion.tePedido, Optional ByVal NObra As String = Nothing)
        Me.AlbVentaInfo = AVInfo
        Me.IDContador = IDContador
        Me.FechaAlbaran = FechaAlbaran
        Me.IDTipoAlbaran = IDTipoAlbaran
        Me.TipoExpedicion = TipoExpedicion
        Me.NObra = NObra
    End Sub

End Class