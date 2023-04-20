<Serializable()> _
Public Class DataSubcontratacion

    Public IDOrdenRuta As Integer
    Public QPedida As Double
    Public QInterna As Double
    Public IDUDInterna As String
    Public IDUDProduccion As String
    Public FechaEntrega As Date

    Public Sub New(ByVal IDOrdenRuta As Integer, ByVal QPedida As Double, ByVal IDUDProduccion As String, ByVal QInterna As Double, ByVal IDUDInterna As String, ByVal FechaEntrega As Date)
        Me.IDOrdenRuta = IDOrdenRuta
        Me.QPedida = QPedida
        Me.QInterna = QInterna
        Me.IDUDInterna = IDUDInterna
        Me.IDUDProduccion = IDUDProduccion
        Me.FechaEntrega = FechaEntrega
    End Sub

End Class
