Public Class ProcessInfoPV
    Inherits ProcessInfo

    Public Detalle As Boolean
    
    Public Sub New(ByVal IDContador As String, Optional ByVal Detalle As Boolean = False)
        MyBase.New(IDContador)
        Me.Detalle = Detalle
    End Sub

End Class
