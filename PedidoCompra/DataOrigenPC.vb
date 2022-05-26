<Serializable()> _
Public Class DataOrigenPC
    Public IDOrigen As Integer
    Public QPedir As Double

    Public Sub New(ByVal IDOrigen As Integer, ByVal QPedir As Double)
        Me.IDOrigen = IDOrigen
        Me.QPedir = QPedir
    End Sub
End Class
