<Serializable()> _
Public Class DataPrograma

    Public IDPrograma As String
    Public IDLineaPrograma As Integer
    Public QConfirmar As Double
    Public QConfirmada As Double
    Public FechaConfirmacion As Date

    Public Sub New(ByVal IDPrograma As String, ByVal IDLineaPrograma As Integer, ByVal QConfirmar As Double, Optional ByVal FechaConfirmacion As Date = cnMinDate)
        Me.IDPrograma = IDPrograma
        Me.IDLineaPrograma = IDLineaPrograma
        Me.QConfirmar = QConfirmar
        If FechaConfirmacion = cnMinDate Then
            Me.FechaConfirmacion = Today
        Else
            Me.FechaConfirmacion = FechaConfirmacion
        End If
    End Sub

End Class
