Public Class ProcessInfoSubcontratacion
    Inherits ProcessInfo

    Public AgruparPorProveedor As Boolean

    Public Sub New(ByVal IDContador As String, ByVal AgruparPorProveedor As Boolean)
        MyBase.New(IDContador)
        Me.AgruparPorProveedor = AgruparPorProveedor
    End Sub
End Class
