Public Class ProcessInfoPC
    Inherits ProcessInfo

    Public Detalle As Boolean
    Public IDOperario As String
    Public FechaEntrega As Date

    Public Sub New(ByVal IDContador As String)
        MyBase.New(IDContador)
    End Sub

    Public Sub New(ByVal IDContador As String, ByVal Detalle As Boolean)
        MyBase.New(IDContador)
        Me.Detalle = Detalle
    End Sub

    Public Sub New(ByVal IDContador As String, ByVal IDOperario As String)
        MyBase.New(IDContador)
        Me.IDOperario = IDOperario
    End Sub

    Public Sub New(ByVal IDContador As String, ByVal IDOperario As String, ByVal FechaEntrega As Date)
        MyBase.New(IDContador)
        Me.IDOperario = IDOperario
        Me.FechaEntrega = FechaEntrega
    End Sub

End Class
