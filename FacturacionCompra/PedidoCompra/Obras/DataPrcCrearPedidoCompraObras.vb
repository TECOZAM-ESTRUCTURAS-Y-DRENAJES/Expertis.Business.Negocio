<Serializable()> _
Public Class DataPrcCrearPedidoCompraObras
    Public Obras() As DataOrigenPC
    Public IDContador As String
    Public IDOperario As String
    Public PorTrabajo As Boolean
    Public PorMaterial As Boolean
    Public FechaEntrega As Date?

    Public Sub New(ByVal Obras() As DataOrigenPC, ByVal IDContador As String, ByVal IDOperario As String, ByVal PorTrabajo As Boolean, ByVal PorMaterial As Boolean, ByVal FechaEntrega As Date?)
        Me.Obras = Obras
        Me.IDContador = IDContador
        Me.IDOperario = IDOperario
        Me.PorMaterial = PorMaterial
        Me.PorTrabajo = PorTrabajo
        If Not FechaEntrega Is Nothing Then Me.FechaEntrega = FechaEntrega
        If (Me.PorMaterial AndAlso Me.PorTrabajo) OrElse (Not Me.PorMaterial AndAlso Not Me.PorTrabajo) Then
            ApplicationService.GenerateError("Revise el Origen de datos, sólo se puede generar por Material o por Trabajo. Nunca por ambas y siempre por alguna de ellas.")
        End If
    End Sub
End Class
