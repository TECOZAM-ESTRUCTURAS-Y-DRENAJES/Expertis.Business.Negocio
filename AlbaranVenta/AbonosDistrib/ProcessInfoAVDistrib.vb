Public Class ProcessInfoAVDistrib
    Inherits ProcessInfo

    Public FechaAlbaran As Date?

    Public Sub New()

    End Sub

    Public Sub New(ByVal IDContador As String, ByVal FechaAlbaran As Date?)
        MyBase.New(IDContador)

        Me.FechaAlbaran = FechaAlbaran
    End Sub
End Class
