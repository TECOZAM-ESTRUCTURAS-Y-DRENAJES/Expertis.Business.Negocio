<Serializable()> _
Public Class DataPrcCrearAlbaranVentaAbono
    Public IDContador As String
    Public FechaAlbaran As Date?
    Public IDAlbaranCliente(-1) As Integer

    Public Sub New(ByVal IDAlbaranCliente As Integer)
        ReDim Me.IDAlbaranCliente(Me.IDAlbaranCliente.Length)
        Me.IDAlbaranCliente(Me.IDAlbaranCliente.Length - 1) = IDAlbaranCliente
    End Sub

    Public Sub New(ByVal IDAlbaranCliente As Integer())
        Me.IDAlbaranCliente = IDAlbaranCliente
    End Sub

End Class
