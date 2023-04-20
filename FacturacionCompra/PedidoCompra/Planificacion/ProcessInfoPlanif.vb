Public Class ProcessInfoPlanif
    Inherits ProcessInfo

    Public AgruparPorProveedor As Boolean

    Public Sub New(ByVal IDContador As String, Optional ByVal AgruparPorProveedor As Boolean = False)
        MyBase.New(IDContador)
        Me.AgruparPorProveedor = AgruparPorProveedor
    End Sub

End Class
