Public Class FraLinCompraNuevoGasto
    Public IDArticulo As String
    Public DescArticulo As String
    Public RefProveedor As String
    Public DescRefProveedor As String
    Public CContable As String
    Public IDTipoIVA As String
    Public Importe As Double
    Public IDObra As Integer?
    Public IDTrabajo As Integer?
    Public Analitica As DataTable

    Public Sub New(ByVal data As DataPrcFacturacionLineaNuevoGasto)
        Me.IDArticulo = data.IDArticulo
        Me.DescArticulo = data.DescArticulo
        Me.RefProveedor = data.RefProveedor
        Me.DescRefProveedor = data.DescRefProveedor
        Me.CContable = data.IDCContable
        Me.IDTipoIVA = data.IDTipoIVA
        Me.Importe = data.Importe
        Me.IDObra = data.IDObra
        Me.IDTrabajo = data.IDTrabajo
        Me.Analitica = data.Analitica
    End Sub

End Class
