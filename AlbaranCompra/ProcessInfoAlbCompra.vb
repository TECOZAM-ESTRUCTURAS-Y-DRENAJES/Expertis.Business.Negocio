Public Class ProcessInfoAlbCompra
    Inherits ProcessInfo

    Public IDTipoCompra As String
    Public FechaAlbaran As Date

    Public Sub New(ByVal IDContador As String, ByVal IDTipoCompra As String, ByVal FechaAlbaran As Date)
        MyBase.New(IDContador)
        Me.IDTipoCompra = IDTipoCompra
        Me.FechaAlbaran = FechaAlbaran
    End Sub

End Class
