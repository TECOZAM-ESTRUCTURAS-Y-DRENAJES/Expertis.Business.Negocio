Public Class FraLinObraCertificacion
    Inherits FraLinObra

    Public IDTrabajo As Integer
    Public QCertificada As Double
    Public IDTipoIva As String
    Public IDCertificacion As Integer


    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(oRow)

        IDTrabajo = oRow("IDTrabajo")
        IDTipoIva = oRow("IDTipoIva")
        QCertificada = oRow("QCertificada")
        IDCertificacion = oRow("IDCertificacion")
    End Sub

End Class