Public Class FraCabObraCertificacion
    Inherits FraCabObra

    Public NRegistro As Integer
    Public Impuestos As Integer

    Public Lineas(-1) As FraLinObraCertificacion

    Public Sub New(ByVal oRow As IPropertyAccessor)
        MyBase.New(oRow)
        'NRegistro = Contador
        Impuestos = Nz(oRow("Impuestos"), TipoRetencionImpuestos.DespuesImpuestos)
        If Impuestos <> 1 Then Retencion = Nz(oRow("Retencion"), 0)
        Fecha = oRow("FechaVencimiento")
    End Sub

    Public Sub Add(ByVal lin As FraLinObraCertificacion)
        ReDim Preserve Lineas(Lineas.Length)
        Lineas(Lineas.Length - 1) = lin
    End Sub

End Class