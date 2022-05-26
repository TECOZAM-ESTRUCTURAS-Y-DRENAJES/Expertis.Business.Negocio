Public Class FraCabVencimiento
    Inherits FraCabObra

    Public Sub New(ByVal oRow As DataRow)
        MyBase.New(New DataRowPropertyAccessor(oRow))

        Me.Fecha = Nz(oRow("FechaVencimiento"), Today)
        Me.Dto = oRow("DtoComercial")
        Me.Edi = False

        'nwRw("RetencionIRPF") = oRow("RetencionIRPF")
        ' nwRw("IDCentroGestion") = Nz(oRow("IDCentroGestion"), mIDCentroGestion)?
        Me.TipoMnto = Nz(oRow("TipoMnto"), 0)
    End Sub

End Class
