Public Class FraLinMnto
    Public IDMntoOTControl As Integer
    Public QaFacturar As Double

    Public Sub New(ByVal oRow As DataRow)
        IDMntoOTControl = oRow("IDMntoOTControl")
        QaFacturar = Nz(oRow("QConsumida"), 0)
    End Sub

End Class
