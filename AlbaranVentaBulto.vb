Public Class AlbaranVentaBulto
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbAlbaranVentaBulto"

#Region " Guardar Packing List "

    <Serializable()> _
    Public Class DataGuardarPackingList
        Public IDAlbaran As Integer
        Public PackingList As DataTable
        Public LineasAlbaran As DataTable

        Public Sub New(ByVal IDAlbaran As Integer, ByVal PackingList As DataTable, ByVal LineasAlbaran As DataTable)
            Me.IDAlbaran = IDAlbaran
            Me.PackingList = PackingList
            Me.LineasAlbaran = LineasAlbaran
        End Sub
    End Class

    <Task()> Public Shared Sub GuardarPackingList(ByVal data As DataGuardarPackingList, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of Object)(AddressOf Comunes.BeginTransaction, Nothing, services)
        ProcessServer.ExecuteTask(Of Integer)(AddressOf EliminarBultosAlbaran, data.IDAlbaran, services)
        ProcessServer.ExecuteTask(Of DataGuardarPackingList)(AddressOf CrearNuevosBultosAlbaran, data, services)
        ProcessServer.ExecuteTask(Of DataGuardarPackingList)(AddressOf ActualizarLineasAlbaran, data, services)
        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.CommitTransaction, True, services)
    End Sub

    <Task()> Public Shared Sub EliminarBultosAlbaran(ByVal IDAlbaran As Integer, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDAlbaran", IDAlbaran))
        Dim ClsAlbBulto As New AlbaranVentaBulto
        Dim dt As DataTable = ClsAlbBulto.Filter(f)
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            ClsAlbBulto.Delete(dt)
        End If
    End Sub

    <Task()> Public Shared Sub CrearNuevosBultosAlbaran(ByVal data As DataGuardarPackingList, ByVal services As ServiceProvider)
        If Not data.PackingList Is Nothing AndAlso data.PackingList.Rows.Count > 0 Then
            Dim dtBultos As DataTable = New AlbaranVentaBulto().AddNew
            For Each drPackingList As DataRow In data.PackingList.Rows
                Dim drBulto As DataRow = dtBultos.NewRow
                drBulto("IDALbaran") = data.IDAlbaran
                drBulto("NEmbalaje") = drPackingList("NEmbalaje")
                drBulto("IDLineaAlbaran") = drPackingList("IDLineaAlbaran")
                drBulto("NContenedor") = drPackingList("NContenedor")
                drBulto("Cantidad") = drPackingList("Cantidad")
                dtBultos.Rows.Add(drBulto.ItemArray)
            Next
            BusinessHelper.UpdateTable(dtBultos)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarLineasAlbaran(ByVal data As DataGuardarPackingList, ByVal services As ServiceProvider)
        If Not data.LineasAlbaran Is Nothing AndAlso data.LineasAlbaran.Rows.Count > 0 Then
            Dim DocAlbVta As New DocumentoAlbaranVenta(data.IDAlbaran)
            If Not DocAlbVta.dtLineas Is Nothing AndAlso DocAlbVta.dtLineas.Rows.Count > 0 Then
                For Each DrLinea As DataRow In DocAlbVta.dtLineas.Select
                    Dim DrSel() As DataRow = data.LineasAlbaran.Select("IDLineaAlbaran = " & DrLinea("IDLineaAlbaran"))
                    If DrSel.Length > 0 Then
                        DrLinea("QEtiContenedor") = DrSel(0)("QEtiContenedor")
                        DrLinea("QEtiEmbalaje") = DrSel(0)("QEtiEmbalaje")
                    End If
                Next
                Dim PckgAlb As New UpdatePackage
                PckgAlb.Add("AlbaranVentaCabecera", DocAlbVta.HeaderRow.Table)
                PckgAlb.Add("AlbaranVentaLinea", DocAlbVta.dtLineas)
                Dim ClsAlbCab As New AlbaranVentaCabecera
                ClsAlbCab.Update(PckgAlb)
            End If
        End If
    End Sub

#End Region

    Public Function SelOnIDAlbaran(ByVal IDAlbaran As Integer) As DataTable
        Return Filter(New NumberFilterItem("IDAlbaran", IDAlbaran))
    End Function

End Class