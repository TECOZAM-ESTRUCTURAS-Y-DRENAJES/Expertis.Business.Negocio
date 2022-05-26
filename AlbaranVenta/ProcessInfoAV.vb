Public Class ProcessInfoAV
    Inherits ProcessInfo

    Public IDTipoAlbaran As String
    Public Tipo As Integer
    Public FromPresentation As Boolean
    Public FechaAlbaran As Date
    Public TipoExpedicion As enumTipoExpedicion
    'Public ResponsableExpedicion As String
   
    Public Sub New(ByVal IDContador As String, ByVal IDTipoAlbaran As String, ByVal FechaAlbaran As Date, ByVal TipoExpedicion As enumTipoExpedicion)
        MyBase.New(IDContador)
        Me.IDTipoAlbaran = IDTipoAlbaran
        Me.FechaAlbaran = FechaAlbaran
        Me.TipoExpedicion = TipoExpedicion

        Dim TipoAlbInfo As ProcesoAlbaranVenta.TipoAlbaranInfo = ProcessServer.ExecuteTask(Of String, ProcesoAlbaranVenta.TipoAlbaranInfo)(AddressOf ProcesoAlbaranVenta.TipoDeAlbaran, IDTipoAlbaran, New ServiceProvider)
        Me.Tipo = TipoAlbInfo.Tipo
    End Sub

End Class
