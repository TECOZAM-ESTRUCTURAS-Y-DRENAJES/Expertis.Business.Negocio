Public Class FraLinCompraLeasing

    Public IDPagoPeriodico As Integer

    Public IDPago As Integer
    Public QaFacturar As Double
    Public QIntAFacturar As Double

    Public Sub New(ByVal oRow As DataRow)
        If Length(oRow("IdPagoPeriodo")) > 0 Then IDPagoPeriodico = oRow("IdPagoPeriodo")

        IDPago = oRow("IDPago")

        QaFacturar = Double.NaN
        QIntAFacturar = Double.NaN
    End Sub

End Class
