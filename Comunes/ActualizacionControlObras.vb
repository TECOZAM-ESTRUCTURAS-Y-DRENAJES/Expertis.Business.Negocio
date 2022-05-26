Public Class ActualizacionControlObras

    Public Enum enumOrigen
        Albaran
        Factura
    End Enum

    <Task()> Public Shared Function AlbaranGeneradoControl(ByVal data As DataRow, ByVal services As ServiceProvider) As Boolean
        If Length(data("IDLineaAlbaran")) > 0 Then
            Dim Albaranes As EntityInfoCache(Of AlbaranCompraLineaInfo) = services.GetService(Of EntityInfoCache(Of AlbaranCompraLineaInfo))()
            Dim Albaran As AlbaranCompraLineaInfo = Albaranes.GetEntity(data("IDLineaAlbaran"))

            Return Albaran.GeneradoControl
        End If
        Return False
    End Function

#Region " Delete Control Obras "

    Public Class dataDeleteControlObras
        Public drLinea As DataRow
        Public Origen As enumOrigen
        Public Control As dataControlObras

        Public Sub New(ByVal drLinea As DataRow, ByVal Origen As enumOrigen, Optional ByVal Control As dataControlObras = Nothing)
            Me.drLinea = drLinea
            Me.Origen = Origen
            Me.Control = Control
        End Sub
    End Class

    <Task()> Public Shared Sub DeleteObraMaterialControl(ByVal data As dataDeleteControlObras, ByVal services As ServiceProvider)
        Dim Control As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraMaterialControl")

        Dim dtMatControl As DataTable
        Dim f As New Filter
        If Length(data.drLinea("IDObra")) > 0 Then f.Add(New NumberFilterItem("IDObra", data.drLinea("IDObra")))
        If Not data.Control Is Nothing AndAlso data.Control.IDObra <> data.Control.IDObraAnterior AndAlso data.Control.IDObraAnterior > 0 Then
            f.Clear()
            f.Add(New NumberFilterItem("IDObra", data.Control.IDObraAnterior))
        End If
        If data.Origen = enumOrigen.Albaran Then
            f.Add(New NumberFilterItem("IDLineaAlbaran", data.drLinea("IDLineaAlbaran")))
            dtMatControl = Control.Filter(f)
        ElseIf data.Origen = enumOrigen.Factura Then
            f.Add(New NumberFilterItem("IDLinFactura", data.drLinea("IDLineaFactura")))
            dtMatControl = Control.Filter(f)
        Else
            dtMatControl = data.drLinea.Table.Clone
            dtMatControl.ImportRow(data.drLinea)
        End If
        If Not IsNothing(dtMatControl) AndAlso dtMatControl.Rows.Count > 0 Then
            For Each drMatControl As DataRow In dtMatControl.Select
                drMatControl("QReal") = 0 : drMatControl("ImpRealMatA") = 0
                Dim dataTrabajo As New dataActualizacion(drMatControl, enumfclTipoGastoObra.enumfclMaterial, enumfclTipoGastoObra.enumfclMaterial, data.Control)
                ProcessServer.ExecuteTask(Of dataActualizacion)(AddressOf ActualizarObraMaterial, dataTrabajo, services)
                ProcessServer.ExecuteTask(Of dataActualizacion)(AddressOf ActualizarObraTrabajo, dataTrabajo, services)
                ProcessServer.ExecuteTask(Of dataActualizacion)(AddressOf ActualizarObraCabecera, dataTrabajo, services)
            Next
            Control.Delete(dtMatControl)
        End If
    End Sub

    <Task()> Public Shared Sub DeleteObraGastoControl(ByVal data As dataDeleteControlObras, ByVal services As ServiceProvider)
        Dim Control As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraGastoControl")
        Dim f As New Filter
        If Length(data.drLinea("IDObra")) > 0 Then f.Add(New NumberFilterItem("IDObra", data.drLinea("IDObra")))
        If Not data.Control Is Nothing AndAlso data.Control.IDObra <> data.Control.IDObraAnterior AndAlso data.Control.IDObraAnterior > 0 Then
            f.Clear()
            f.Add(New NumberFilterItem("IDObra", data.Control.IDObraAnterior))
        End If
        If data.Origen = enumOrigen.Albaran Then
            f.Add(New NumberFilterItem("IDLineaAlbaran", data.drLinea("IDLineaAlbaran")))
        Else
            f.Add(New NumberFilterItem("IDLinFactura", data.drLinea("IDLineaFactura")))
        End If
        Dim dtGastoControl As DataTable = Control.Filter(f)
        If Not IsNothing(dtGastoControl) AndAlso dtGastoControl.Rows.Count > 0 Then
            For Each drGastoControl As DataRow In dtGastoControl.Select
                drGastoControl("ImpRealGastoA") = 0
                Dim dataTrabajo As New dataActualizacion(drGastoControl, enumfclTipoGastoObra.enumfclGastos)
                ProcessServer.ExecuteTask(Of dataActualizacion)(AddressOf ActualizarObraGasto, dataTrabajo, services)
                ProcessServer.ExecuteTask(Of dataActualizacion)(AddressOf ActualizarObraTrabajo, dataTrabajo, services)
                ProcessServer.ExecuteTask(Of dataActualizacion)(AddressOf ActualizarObraCabecera, dataTrabajo, services)
            Next
            Control.Delete(dtGastoControl)
        End If
    End Sub

    <Task()> Public Shared Sub DeleteObraVariosControl(ByVal data As dataDeleteControlObras, ByVal services As ServiceProvider)
        Dim Control As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraVariosControl")
        Dim f As New Filter
        If Length(data.drLinea("IDObra")) > 0 Then f.Add(New NumberFilterItem("IDObra", data.drLinea("IDObra")))
        If Not data.Control Is Nothing AndAlso data.Control.IDObra <> data.Control.IDObraAnterior AndAlso data.Control.IDObraAnterior > 0 Then
            f.Clear()
            f.Add(New NumberFilterItem("IDObra", data.Control.IDObraAnterior))
        End If
        If data.Origen = enumOrigen.Albaran Then
            f.Add(New NumberFilterItem("IDLineaAlbaran", data.drLinea("IDLineaAlbaran")))
        Else
            f.Add(New NumberFilterItem("IDLinFactura", data.drLinea("IDLineaFactura")))
        End If
        Dim dtVariosControl As DataTable = Control.Filter(f)
        If Not IsNothing(dtVariosControl) AndAlso dtVariosControl.Rows.Count > 0 Then
            For Each drVariosControl As DataRow In dtVariosControl.Select
                drVariosControl("ImpRealVariosA") = 0
                Dim dataTrabajo As New dataActualizacion(drVariosControl, enumfclTipoGastoObra.enumfclVarios)
                ProcessServer.ExecuteTask(Of dataActualizacion)(AddressOf ActualizarObraVarios, dataTrabajo, services)
                ProcessServer.ExecuteTask(Of dataActualizacion)(AddressOf ActualizarObraTrabajo, dataTrabajo, services)
                ProcessServer.ExecuteTask(Of dataActualizacion)(AddressOf ActualizarObraCabecera, dataTrabajo, services)
            Next
            Control.Delete(dtVariosControl)
        End If
    End Sub

#End Region

#Region " Actualizar Control Obras "

    Public Class dataControlObras
        Public drLinea As DataRow
        Public Fecha As Date
        Public Cantidad, CantidadAnterior, Precio, PrecioAnterior, Dto1, Dto1Anterior, Dto2, Dto2Anterior, Dto3, Dto3Anterior, ImporteA, ImporteAAnterior As Double
        Public TipoGasto As enumfclTipoGastoObra
        Public TipoGastoAnterior As enumfclTipoGastoObra = -1
        Public IDLineaPadre, IDLineaPadreAnterior As Integer
        Public IDObra, IDObraAnterior As Integer
        Public IDTrabajo, IDTrabajoAnterior As Integer
        Public Origen As enumOrigen

        Public Sub New(ByVal drLinea As DataRow, ByVal Fecha As Date, ByVal Origen As enumOrigen)
            Me.drLinea = drLinea
            Me.Fecha = Fecha
            Me.Origen = Origen
            Me.IDObra = Nz(drLinea("IDObra"), 0)
            Me.IDTrabajo = Nz(drLinea("IDTrabajo"), 0)
            Me.Cantidad = Nz(drLinea("QInterna"), 0)
            Me.Precio = Nz(drLinea("Precio"), 0)
            Me.Dto1 = Nz(drLinea("Dto1"), 0)
            Me.Dto2 = Nz(drLinea("Dto2"), 0)
            Me.Dto3 = Nz(drLinea("Dto3"), 0)
            Me.ImporteA = Nz(drLinea("ImporteA"), 0)
            Me.TipoGasto = Nz(drLinea("TipoGastoObra"), enumfclTipoGastoObra.enumfclMaterial)
            Me.IDLineaPadre = Nz(drLinea("IDLineaPadre"), -1)
            If drLinea.RowState = DataRowState.Modified Then
                Me.IDObraAnterior = Nz(drLinea("IDObra", DataRowVersion.Original), 0)
                Me.IDTrabajoAnterior = Nz(drLinea("IDTrabajo", DataRowVersion.Original), 0)
                Me.ImporteAAnterior = Nz(drLinea("ImporteA", DataRowVersion.Original), 0)
                'If drLinea("QInterna") <> Nz(drLinea("QInterna", DataRowVersion.Original), 0) Then
                CantidadAnterior = Nz(drLinea("QInterna", DataRowVersion.Original), 0)
                PrecioAnterior = Nz(drLinea("Precio", DataRowVersion.Original), 0)
                Dto1Anterior = Nz(drLinea("Dto1", DataRowVersion.Original), 0)
                Dto2Anterior = Nz(drLinea("Dto2", DataRowVersion.Original), 0)
                Dto3Anterior = Nz(drLinea("Dto3", DataRowVersion.Original), 0)
                'End If
                If Length(drLinea("TipoGastoObra", DataRowVersion.Original)) > 0 Then
                    Me.TipoGastoAnterior = drLinea("TipoGastoObra", DataRowVersion.Original)
                End If
                Me.IDLineaPadreAnterior = Nz(drLinea("IDLineaPadre", DataRowVersion.Original), -1)
            End If
        End Sub
    End Class

#Region " ObraMaterialControl "

    <Task()> Public Shared Sub ActualizarObraMaterialControl(ByVal data As dataControlObras, ByVal services As ServiceProvider)
        If Length(data.drLinea("IDObra")) > 0 Then
            Dim OC As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraCabecera")
            Dim dtObraCabecera As DataTable = OC.SelOnPrimaryKey(data.drLinea("IDObra"))
            If Not dtObraCabecera Is Nothing AndAlso dtObraCabecera.Rows.Count > 0 Then
                Dim blnActualizarMatControl As Boolean = True
                Dim context As New BusinessData(dtObraCabecera.Rows(0))
                ProcessServer.ExecuteTask(Of dataControlObras)(AddressOf BorrarControlAnterior, data, services)
                Dim drMatControl As DataRow
                Dim f As New Filter
                f.Add(New NumberFilterItem("IDObra", data.drLinea("IDObra")))
                If data.Origen = enumOrigen.Albaran Then
                    f.Add(New NumberFilterItem("IDLineaAlbaran", data.drLinea("IDLineaAlbaran")))
                Else
                    f.Add(New NumberFilterItem("IDLinFactura", data.drLinea("IDLineaFactura")))
                End If
                Dim datControl As dataControlObras
                Dim OMC As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraMaterialControl")
                Dim dtMatControl As DataTable = OMC.Filter(f)
                If dtMatControl Is Nothing OrElse dtMatControl.Rows.Count = 0 Then
                    drMatControl = ProcessServer.ExecuteTask(Of dataControlObras, DataRow)(AddressOf NuevaLineaObraMaterialControl, data, services)
                    If data.IDLineaPadre <> data.IDLineaPadreAnterior AndAlso data.IDLineaPadre > 0 Then
                        datControl = data
                    End If
                Else
                    drMatControl = dtMatControl.Rows(0)
                    datControl = data
                    If data.IDObra <> data.IDObraAnterior Then
                        drMatControl("IDObra") = data.IDObra
                    End If
                    If data.IDTrabajo <> data.IDTrabajoAnterior Then
                        If data.IDTrabajo > 0 Then
                            drMatControl("IDTrabajo") = data.IDTrabajo
                        Else : drMatControl("IDTrabajo") = System.DBNull.Value
                        End If
                    End If
                    If data.Fecha <> Nz(drMatControl("Fecha"), cnMinDate) Then
                        drMatControl("Fecha") = data.Fecha
                    End If
                    If data.Cantidad <> data.CantidadAnterior Then
                        'OMC.ApplyBusinessRule("QReal", drMatControl("QReal") + data.Cantidad - data.CantidadAnterior, drMatControl, context)
                        drMatControl("QReal") = drMatControl("QReal") + data.Cantidad - data.CantidadAnterior
                    End If
                    If data.Precio <> data.PrecioAnterior Then
                        'OMC.ApplyBusinessRule("PrecioRealMatA", data.Precio, drMatControl, context)
                        drMatControl("PrecioRealMatA") = data.Precio
                    End If

                    If data.Dto1 <> data.Dto1Anterior Then
                        drMatControl("Dto1") = data.Dto1
                    End If

                    If data.Dto2 <> data.Dto2Anterior Then
                        drMatControl("Dto2") = data.Dto2
                    End If

                    If data.Dto3 <> data.Dto3Anterior Then
                        drMatControl("Dto3") = data.Dto3
                    End If
                   
                    If data.ImporteA <> data.ImporteAAnterior Then
                        drMatControl("ImpRealMatA") = data.ImporteA
                    Else
                        If data.IDLineaPadre <> data.IDLineaPadreAnterior Then
                            data.ImporteA = data.ImporteAAnterior
                            drMatControl("ImpRealMatA") = data.ImporteA
                        End If
                    End If
                End If

                Dim dataTrabajo As New dataActualizacion(drMatControl, data.TipoGasto, data.TipoGastoAnterior, datControl)
                ProcessServer.ExecuteTask(Of dataActualizacion)(AddressOf ActualizarObraMaterial, dataTrabajo, services)
                ProcessServer.ExecuteTask(Of dataActualizacion)(AddressOf ActualizarObraTrabajo, dataTrabajo, services)
                ProcessServer.ExecuteTask(Of dataActualizacion)(AddressOf ActualizarObraCabecera, dataTrabajo, services)
                If blnActualizarMatControl Then BusinessHelper.UpdateTable(drMatControl.Table)
            End If

        ElseIf Length(data.drLinea("IDObra")) = 0 AndAlso Length(data.drLinea("IDObra", DataRowVersion.Original)) > 0 Then
            Dim dataDelete As New dataDeleteControlObras(data.drLinea, data.Origen, data)
            ProcessServer.ExecuteTask(Of dataDeleteControlObras)(AddressOf DeleteObraMaterialControl, dataDelete, services)
        End If
    End Sub

    <Task()> Public Shared Function NuevaLineaObraMaterialControl(ByVal data As dataControlObras, ByVal services As ServiceProvider) As DataRow
        Dim OMC As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraMaterialControl")
        Dim dtLineas As DataTable = OMC.AddNew
        Dim drMatControl As DataRow = dtLineas.NewRow
        Dim CodTrabajo As String

        Dim Obras As EntityInfoCache(Of ObraCabeceraInfo) = services.GetService(Of EntityInfoCache(Of ObraCabeceraInfo))()
        Dim Obra As ObraCabeceraInfo = Obras.GetEntity(data.drLinea("IDObra"))

        Dim context As New BusinessData
        context("IDCliente") = Obra.IDCliente


        drMatControl("IDLineaMaterialControl") = AdminData.GetAutoNumeric
        drMatControl("IDObra") = data.drLinea("IDObra")
        drMatControl("IDTrabajo") = data.drLinea("IDTrabajo")
        drMatControl("IDMaterial") = data.drLinea("IDArticulo")
        drMatControl("DescMaterial") = data.drLinea("DescArticulo")

        Dim objNegObraTrabajo As BusinessHelper
        objNegObraTrabajo = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("ObraTrabajo"))
        Dim objFilter As New Filter
        objFilter.Add(New StringFilterItem("IDTrabajo", data.drLinea("IDTrabajo")))
        Dim dtTrabajo As DataTable = objNegObraTrabajo.Filter(objFilter)
        If Not IsNothing(dtTrabajo) AndAlso dtTrabajo.Rows.Count > 0 Then
            CodTrabajo = dtTrabajo.Rows(0)("CodTrabajo")
        End If
        OMC.ApplyBusinessRule("CodTrabajo", CodTrabajo, drMatControl, context)
        OMC.ApplyBusinessRule("IDLineaMaterial", data.drLinea("IDLineaPadre"), drMatControl, context)
        OMC.ApplyBusinessRule("QReal", data.drLinea("QInterna"), drMatControl, context)

        If data.drLinea("Factor") = 0 Then
            drMatControl("PrecioRealMatA") = 0
        Else
            drMatControl("PrecioRealMatA") = data.drLinea("PrecioA") / data.drLinea("Factor")
        End If
        drMatControl("Dto1") = data.drLinea("Dto1")
        drMatControl("Dto2") = data.drLinea("Dto2")
        drMatControl("Dto3") = data.drLinea("Dto3")
        drMatControl("ImpRealMatA") = data.drLinea("ImporteA")
        drMatControl("Fecha") = data.Fecha
        If data.drLinea.Table.Columns.Contains("IDLineaAlbaran") Then drMatControl("IDLineaAlbaran") = data.drLinea("IDLineaAlbaran")
        If data.drLinea.Table.Columns.Contains("IDLineaFactura") Then drMatControl("IDLinFactura") = data.drLinea("IDLineaFactura")
        If data.Origen = enumOrigen.Factura Then
            drMatControl("IDAlmacen") = ProcessServer.ExecuteTask(Of DataRow, String)(AddressOf GetAlmacen, data.drLinea, services)
        Else
            drMatControl("IDAlmacen") = data.drLinea("IDAlmacen")
        End If
        drMatControl("UDValoracion") = data.drLinea("UDValoracion")
        If Length(data.drLinea("IDLineaAlbaran")) > 0 Then
            drMatControl("Actualizado") = CInt(enumomcActualizado.omcNoActualizado)
        Else
            drMatControl("Actualizado") = CInt(enumomcActualizado.omcSinGestion)
        End If
        drMatControl("TipoMaterial") = enumomcTipoMaterial.omcPedidoCompra

        dtLineas.Rows.Add(drMatControl.ItemArray)
        BusinessHelper.UpdateTable(dtLineas)

        Return drMatControl
    End Function

    <Task()> Public Shared Function GetAlmacen(ByVal data As DataRow, ByVal services As ServiceProvider) As String
        Dim IDAlmacen As String = String.Empty
        If Length(data("IDLineaPadre")) > 0 Then
            Dim dtOM As DataTable = New BE.DataEngine().Filter("tbObraMaterial", New NumberFilterItem("IDLineaMaterial", data("IDLineaPadre")), "IDAlmacen")
            If Not dtOM Is Nothing AndAlso dtOM.Rows.Count > 0 Then
                IDAlmacen = dtOM.Rows(0)("IDAlmacen") & String.Empty
            End If
        End If
        If Length(IDAlmacen) = 0 AndAlso Length(data("IDLineaAlbaran")) > 0 Then
            Dim Albaranes As EntityInfoCache(Of AlbaranCompraLineaInfo) = services.GetService(Of EntityInfoCache(Of AlbaranCompraLineaInfo))()
            Dim Albaran As AlbaranCompraLineaInfo = Albaranes.GetEntity(data("IDLineaAlbaran"))

            IDAlmacen = Albaran.IDAlmacen
        End If
        If Length(IDAlmacen) = 0 Then
            Dim StDatos As New DataArtAlm(data("IDArticulo"))
            IDAlmacen = ProcessServer.ExecuteTask(Of DataArtAlm, String)(AddressOf ArticuloAlmacen.AlmacenPredeterminadoArticulo, StDatos, services)
        End If
        Return IDAlmacen
    End Function

#End Region

#Region " ObraGastoControl "

    <Task()> Public Shared Sub ActualizarObraGastoControl(ByVal data As dataControlObras, ByVal services As ServiceProvider)
        If Length(data.drLinea("IDObra")) > 0 Then
            ProcessServer.ExecuteTask(Of dataControlObras)(AddressOf BorrarControlAnterior, data, services)
            Dim drGastoControl As DataRow
            Dim f As New Filter
            If data.Origen = enumOrigen.Albaran Then
                f.Add(New NumberFilterItem("IDLineaAlbaran", data.drLinea("IDLineaAlbaran")))
            Else
                f.Add(New NumberFilterItem("IDLinFactura", data.drLinea("IDLineaFactura")))
            End If
            Dim dtGastoControl As DataTable = BusinessHelper.CreateBusinessObject("ObraGastoControl").Filter(f)
            If dtGastoControl Is Nothing OrElse dtGastoControl.Rows.Count = 0 Then
                drGastoControl = ProcessServer.ExecuteTask(Of dataControlObras, DataRow)(AddressOf NuevaLineaObraGastoControl, data, services)
            Else
                drGastoControl = dtGastoControl.Rows(0)
                If data.Fecha <> Nz(drGastoControl("Fecha"), cnMinDate) Then
                    drGastoControl("Fecha") = data.Fecha
                End If
                If data.ImporteA <> data.ImporteAAnterior Then
                    drGastoControl("ImpRealGastoA") = drGastoControl("ImpRealGastoA") + data.ImporteA - data.ImporteAAnterior
                End If
            End If

            Dim dataTrabajo As New dataActualizacion(drGastoControl, data.TipoGasto, data.TipoGastoAnterior)
            ProcessServer.ExecuteTask(Of dataActualizacion)(AddressOf ActualizarObraGasto, dataTrabajo, services)
            ProcessServer.ExecuteTask(Of dataActualizacion)(AddressOf ActualizarObraTrabajo, dataTrabajo, services)
            ProcessServer.ExecuteTask(Of dataActualizacion)(AddressOf ActualizarObraCabecera, dataTrabajo, services)
            BusinessHelper.UpdateTable(drGastoControl.Table)
        ElseIf Length(data.drLinea("IDObra")) = 0 AndAlso Length(data.drLinea("IDObra", DataRowVersion.Original)) > 0 Then
            Dim dataDelete As New dataDeleteControlObras(data.drLinea, data.Origen)
            ProcessServer.ExecuteTask(Of dataDeleteControlObras)(AddressOf DeleteObraGastoControl, dataDelete, services)
        End If
    End Sub

    <Task()> Public Shared Function NuevaLineaObraGastoControl(ByVal data As dataControlObras, ByVal services As ServiceProvider) As DataRow
        Dim OGC As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraGastoControl")
        Dim dtLineas As DataTable = OGC.AddNew
        Dim drGastoControl As DataRow = dtLineas.NewRow
        Dim CodTrabajo As String
        Dim context As New BusinessData

        drGastoControl("IDLineaGastoControl") = AdminData.GetAutoNumeric
        drGastoControl("IDLineaGasto") = data.drLinea("IDLineaPadre")
        drGastoControl("IDObra") = data.drLinea("IDObra")
        drGastoControl("IDTrabajo") = data.drLinea("IDTrabajo")

        Dim objNegObraTrabajo As BusinessHelper
        objNegObraTrabajo = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("ObraTrabajo"))
        Dim objFilter As New Filter
        objFilter.Add(New StringFilterItem("IDTrabajo", data.drLinea("IDTrabajo")))
        Dim dtTrabajo As DataTable = objNegObraTrabajo.Filter(objFilter)
        If Not IsNothing(dtTrabajo) AndAlso dtTrabajo.Rows.Count > 0 Then
            CodTrabajo = dtTrabajo.Rows(0)("CodTrabajo")
        End If
        OGC.ApplyBusinessRule("CodTrabajo", CodTrabajo, drGastoControl, context)

        If Length(data.drLinea("IDLineaPadre")) > 0 Then
            Dim dtConcepto As DataTable = BusinessHelper.CreateBusinessObject("ObraGasto").Filter(New FilterItem("IDLineaGasto", data.drLinea("IDLineaPadre")))
            drGastoControl("IDGasto") = dtConcepto.Rows(0)("IDGasto")
            drGastoControl("DescGasto") = dtConcepto.Rows(0)("DescGasto")
        Else
            If Length(data.drLinea("IDConcepto")) > 0 Then
                drGastoControl("IDGasto") = data.drLinea("IDConcepto")
                Dim dtGasto As DataTable = New BE.DataEngine().Filter("tbMaestroGasto", New FilterItem("IDGasto", drGastoControl("IDGasto")))
                If Not dtGasto Is Nothing AndAlso dtGasto.Rows.Count > 0 Then
                    drGastoControl("DescGasto") = dtGasto.Rows(0)("DescGasto")
                End If
            Else
                Dim dtParam As DataTable = New Parametro().ConceptoGastosProyectos
                If Not IsNothing(dtParam) AndAlso dtParam.Rows.Count > 0 Then
                    drGastoControl("IDGasto") = dtParam.Rows(0)("IDGasto")
                    drGastoControl("DescGasto") = dtParam.Rows(0)("DescGasto")
                End If
            End If
        End If

        If data.drLinea.Table.Columns.Contains("IDLineaAlbaran") Then drGastoControl("IDLineaAlbaran") = data.drLinea("IDLineaAlbaran")
        If data.drLinea.Table.Columns.Contains("IDLineaFactura") Then drGastoControl("IDLinFactura") = data.drLinea("IDLineaFactura")
        drGastoControl("Fecha") = data.Fecha
        drGastoControl("ImpRealGastoA") = data.drLinea("ImporteA")

        dtLineas.Rows.Add(drGastoControl.ItemArray)
        BusinessHelper.UpdateTable(dtLineas)

        Return drGastoControl
    End Function

#End Region

#Region " ObraVariosControl "

    <Task()> Public Shared Sub ActualizarObraVariosControl(ByVal data As dataControlObras, ByVal services As ServiceProvider)
        If Length(data.drLinea("IDObra")) > 0 Then
            ProcessServer.ExecuteTask(Of dataControlObras)(AddressOf BorrarControlAnterior, data, services)
            Dim drVariosControl As DataRow
            Dim f As New Filter
            If data.Origen = enumOrigen.Albaran Then
                f.Add(New NumberFilterItem("IDLineaAlbaran", data.drLinea("IDLineaAlbaran")))
            Else
                f.Add(New NumberFilterItem("IDLinFactura", data.drLinea("IDLineaFactura")))
            End If
            Dim dtGastoControl As DataTable = BusinessHelper.CreateBusinessObject("ObraVariosControl").Filter(f)
            If dtGastoControl Is Nothing OrElse dtGastoControl.Rows.Count = 0 Then
                drVariosControl = ProcessServer.ExecuteTask(Of dataControlObras, DataRow)(AddressOf NuevaLineaObraVariosControl, data, services)
            Else
                drVariosControl = dtGastoControl.Rows(0)
                If data.Fecha <> Nz(drVariosControl("Fecha"), cnMinDate) Then
                    drVariosControl("Fecha") = data.Fecha
                End If
                If data.ImporteA <> data.ImporteAAnterior Then
                    drVariosControl("ImpRealVariosA") = drVariosControl("ImpRealVariosA") + data.ImporteA - data.ImporteAAnterior
                End If
            End If

            Dim dataTrabajo As New dataActualizacion(drVariosControl, data.TipoGasto, data.TipoGastoAnterior)
            ProcessServer.ExecuteTask(Of dataActualizacion)(AddressOf ActualizarObraVarios, dataTrabajo, services)
            ProcessServer.ExecuteTask(Of dataActualizacion)(AddressOf ActualizarObraTrabajo, dataTrabajo, services)
            ProcessServer.ExecuteTask(Of dataActualizacion)(AddressOf ActualizarObraCabecera, dataTrabajo, services)
            BusinessHelper.UpdateTable(drVariosControl.Table)
        ElseIf Length(data.drLinea("IDObra")) = 0 AndAlso Length(data.drLinea("IDObra", DataRowVersion.Original)) > 0 Then
            Dim dataDelete As New dataDeleteControlObras(data.drLinea, data.Origen)
            ProcessServer.ExecuteTask(Of dataDeleteControlObras)(AddressOf DeleteObraVariosControl, dataDelete, services)
        End If
    End Sub

    <Task()> Public Shared Function NuevaLineaObraVariosControl(ByVal data As dataControlObras, ByVal services As ServiceProvider) As DataRow
        Dim dtLineas As DataTable = BusinessHelper.CreateBusinessObject("ObraVariosControl").AddNew
        Dim drVariosControl As DataRow = dtLineas.NewRow

        drVariosControl("IDLineaVariosControl") = AdminData.GetAutoNumeric
        drVariosControl("IDLineaVarios") = data.drLinea("IDLineaPadre")
        drVariosControl("IDObra") = data.drLinea("IDObra")
        drVariosControl("IDTrabajo") = data.drLinea("IDTrabajo")

        If Length(data.drLinea("IDLineaPadre")) > 0 Then
            Dim dtConcepto As DataTable = BusinessHelper.CreateBusinessObject("ObraVarios").Filter(New FilterItem("IDLineaVarios", data.drLinea("IDLineaPadre")))
            drVariosControl("IDVarios") = dtConcepto.Rows(0)("IDVarios")
            drVariosControl("DescVarios") = dtConcepto.Rows(0)("DescVarios")
        Else
            If Length(data.drLinea("IDConcepto")) > 0 Then
                drVariosControl("IDVarios") = data.drLinea("IDConcepto")
                Dim dtVarios As DataTable = New BE.DataEngine().Filter("tbMaestroVarios", New FilterItem("IDVarios", drVariosControl("IDVarios")))
                If Not dtVarios Is Nothing AndAlso dtVarios.Rows.Count > 0 Then
                    drVariosControl("DescVarios") = dtVarios.Rows(0)("DescVarios")
                End If
            Else
                Dim dtParam As DataTable = New Parametro().ConceptoVariosProyectos
                If Not IsNothing(dtParam) AndAlso dtParam.Rows.Count > 0 Then
                    drVariosControl("IDVarios") = dtParam.Rows(0)("IDVarios")
                    drVariosControl("DescVarios") = dtParam.Rows(0)("DescVarios")
                End If
            End If
        End If

        If data.drLinea.Table.Columns.Contains("IDLineaAlbaran") Then drVariosControl("IDLineaAlbaran") = data.drLinea("IDLineaAlbaran")
        If data.drLinea.Table.Columns.Contains("IDLineaFactura") Then drVariosControl("IDLinFactura") = data.drLinea("IDLineaFactura")
        drVariosControl("Fecha") = data.Fecha
        drVariosControl("ImpRealVariosA") = data.drLinea("ImporteA")

        dtLineas.Rows.Add(drVariosControl.ItemArray)
        BusinessHelper.UpdateTable(dtLineas)

        Return drVariosControl
    End Function

#End Region

    <Task()> Public Shared Sub BorrarControlAnterior(ByVal data As dataControlObras, ByVal services As ServiceProvider)
        Dim dataDelete As New dataDeleteControlObras(data.drLinea, data.Origen, data)
        If data.TipoGastoAnterior <> -1 AndAlso (data.TipoGastoAnterior <> data.TipoGasto) Then
            If data.TipoGastoAnterior <> data.TipoGasto Then
                Select Case data.TipoGastoAnterior
                    Case enumfclTipoGastoObra.enumfclMaterial
                        ProcessServer.ExecuteTask(Of dataDeleteControlObras)(AddressOf DeleteObraMaterialControl, dataDelete, services)
                    Case enumfclTipoGastoObra.enumfclGastos
                        ProcessServer.ExecuteTask(Of dataDeleteControlObras)(AddressOf DeleteObraGastoControl, dataDelete, services)
                    Case enumfclTipoGastoObra.enumfclVarios
                        ProcessServer.ExecuteTask(Of dataDeleteControlObras)(AddressOf DeleteObraVariosControl, dataDelete, services)
                End Select
            End If
        ElseIf data.IDLineaPadre <> data.IDLineaPadreAnterior OrElse data.IDTrabajo <> data.IDTrabajoAnterior OrElse data.IDObra <> data.IDObraAnterior Then
            Select Case data.TipoGasto
                Case enumfclTipoGastoObra.enumfclMaterial
                    ProcessServer.ExecuteTask(Of dataDeleteControlObras)(AddressOf DeleteObraMaterialControl, dataDelete, services)
                Case enumfclTipoGastoObra.enumfclGastos
                    ProcessServer.ExecuteTask(Of dataDeleteControlObras)(AddressOf DeleteObraGastoControl, dataDelete, services)
                Case enumfclTipoGastoObra.enumfclVarios
                    ProcessServer.ExecuteTask(Of dataDeleteControlObras)(AddressOf DeleteObraVariosControl, dataDelete, services)
            End Select
        End If
    End Sub

#End Region

#Region " Actualizar Previstos Obras "

    Public Class dataActualizacion
        Public row As DataRow
        Public TipoGasto, TipoGastoAnterior As enumfclTipoGastoObra?
        Public ImporteA, ImporteAAnterior, QReal, QRealAnterior As Double
        Public Control As dataControlObras

        Public Sub New(ByVal row As DataRow, ByVal TipoGasto As enumfclTipoGastoObra, Optional ByVal TipoGastoAnterior As enumfclTipoGastoObra = -1, Optional ByVal Control As dataControlObras = Nothing)
            Me.row = row
            Me.TipoGasto = TipoGasto
            Me.TipoGastoAnterior = TipoGastoAnterior

            If row.RowState = DataRowState.Modified Then
                Select Case TipoGasto
                    Case enumfclTipoGastoObra.enumfclMaterial
                        Me.QReal = row("QReal")
                        If row("QReal") <> Nz(row("QReal", DataRowVersion.Original), 0) Then
                            Me.QRealAnterior = Nz(row("QReal", DataRowVersion.Original), 0)
                        Else
                            Me.QRealAnterior = row("QReal")
                        End If
                        Me.ImporteA = row("ImpRealMatA")
                        If row("ImpRealMatA") <> Nz(row("ImpRealMatA", DataRowVersion.Original), 0) Then
                            Me.ImporteAAnterior = Nz(row("ImpRealMatA", DataRowVersion.Original), 0)
                        Else
                            Me.ImporteAAnterior = row("ImpRealMatA")
                        End If
                    Case enumfclTipoGastoObra.enumfclGastos
                        Me.ImporteA = row("ImpRealGastoA")
                        If row("ImpRealGastoA") <> Nz(row("ImpRealGastoA", DataRowVersion.Original), 0) Then
                            Me.ImporteAAnterior = Nz(row("ImpRealGastoA", DataRowVersion.Original), 0)
                        Else
                            Me.ImporteAAnterior = row("ImpRealGastoA")
                        End If
                    Case enumfclTipoGastoObra.enumfclVarios
                        Me.ImporteA = row("ImpRealVariosA")
                        If row("ImpRealVariosA") <> Nz(row("ImpRealVariosA", DataRowVersion.Original), 0) Then
                            Me.ImporteAAnterior = Nz(row("ImpRealVariosA", DataRowVersion.Original), 0)
                        Else
                            Me.ImporteAAnterior = row("ImpRealVariosA")
                        End If
                End Select
            Else
                Select Case TipoGasto
                    Case enumfclTipoGastoObra.enumfclMaterial
                        Me.QReal = row("QReal")
                        Me.ImporteA = row("ImpRealMatA")
                    Case enumfclTipoGastoObra.enumfclGastos
                        Me.ImporteA = row("ImpRealGastoA")
                    Case enumfclTipoGastoObra.enumfclVarios
                        Me.ImporteA = row("ImpRealVariosA")
                End Select
            End If
            If Not Control Is Nothing Then Me.Control = Control
        End Sub
    End Class

    <Task()> Public Shared Sub ActualizarObraMaterial(ByVal data As dataActualizacion, ByVal services As ServiceProvider)
        If Length(data.row("IDLineaMaterial")) > 0 Then
            Dim dtMatPrev As DataTable = BusinessHelper.CreateBusinessObject("ObraMaterial").SelOnPrimaryKey(data.row("IDLineaMaterial"))
            If Not IsNothing(dtMatPrev) AndAlso dtMatPrev.Rows.Count > 0 Then
                dtMatPrev.Rows(0)("QReal") = dtMatPrev.Rows(0)("QReal") + data.QReal - data.QRealAnterior
                dtMatPrev.Rows(0)("ImpRealMatA") = dtMatPrev.Rows(0)("ImpRealMatA") + data.ImporteA - data.ImporteAAnterior

                BusinessHelper.UpdateTable(dtMatPrev)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarObraGasto(ByVal data As dataActualizacion, ByVal services As ServiceProvider)
        If Length(data.row("IDLineaGasto")) > 0 Then
            Dim dtGastoPrev As DataTable = BusinessHelper.CreateBusinessObject("ObraGasto").SelOnPrimaryKey(data.row("IDLineaGasto"))
            If Not IsNothing(dtGastoPrev) AndAlso dtGastoPrev.Rows.Count > 0 Then
                dtGastoPrev.Rows(0)("ImpRealGastoA") = dtGastoPrev.Rows(0)("ImpRealGastoA") + data.ImporteA - data.ImporteAAnterior

                BusinessHelper.UpdateTable(dtGastoPrev)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarObraVarios(ByVal data As dataActualizacion, ByVal services As ServiceProvider)
        If Length(data.row("IDLineaVarios")) > 0 Then
            Dim dtVariosPrev As DataTable = BusinessHelper.CreateBusinessObject("ObraVarios").SelOnPrimaryKey(data.row("IDLineaVarios"))
            If Not IsNothing(dtVariosPrev) AndAlso dtVariosPrev.Rows.Count > 0 Then
                dtVariosPrev.Rows(0)("ImpRealVariosA") = dtVariosPrev.Rows(0)("ImpRealVariosA") + data.ImporteA - data.ImporteAAnterior

                BusinessHelper.UpdateTable(dtVariosPrev)
            End If
        End If
    End Sub

#End Region

    <Task()> Public Shared Sub ActualizarObraTrabajo(ByVal data As dataActualizacion, ByVal services As ServiceProvider)
        '//Actualizar Trabajo Anterior
        Dim IDTrabajoAnterior As Integer
        If Not data.Control Is Nothing Then
            IDTrabajoAnterior = data.Control.IDTrabajoAnterior
            If (data.Control.IDTrabajo <> data.Control.IDTrabajoAnterior AndAlso data.Control.IDTrabajoAnterior > 0) Then
                Dim dtTrabajoAnterior As DataTable = BusinessHelper.CreateBusinessObject("ObraTrabajo").SelOnPrimaryKey(data.Control.IDTrabajoAnterior)
                If Not IsNothing(dtTrabajoAnterior) AndAlso dtTrabajoAnterior.Rows.Count > 0 Then
                    If data.TipoGastoAnterior <> -1 Then
                        Select Case data.TipoGastoAnterior
                            Case enumfclTipoGastoObra.enumfclMaterial
                                dtTrabajoAnterior.Rows(0)("ImpRealMatA") = dtTrabajoAnterior.Rows(0)("ImpRealMatA") - data.Control.ImporteAAnterior
                            Case enumfclTipoGastoObra.enumfclGastos
                                dtTrabajoAnterior.Rows(0)("ImpRealGastosA") = dtTrabajoAnterior.Rows(0)("ImpRealGastosA") - data.Control.ImporteAAnterior
                            Case enumfclTipoGastoObra.enumfclVarios
                                dtTrabajoAnterior.Rows(0)("ImpRealVariosA") = dtTrabajoAnterior.Rows(0)("ImpRealVariosA") - data.Control.ImporteAAnterior
                        End Select
                    End If
                    dtTrabajoAnterior.Rows(0)("ImpRealTrabajoA") = dtTrabajoAnterior.Rows(0)("ImpRealTrabajoA") - data.Control.ImporteAAnterior
                    BusinessHelper.UpdateTable(dtTrabajoAnterior)
                End If
            ElseIf data.Control.ImporteA <> data.Control.ImporteAAnterior Then
                Dim dtTrabajoAnterior As DataTable = BusinessHelper.CreateBusinessObject("ObraTrabajo").SelOnPrimaryKey(data.Control.IDTrabajoAnterior)
                If Not IsNothing(dtTrabajoAnterior) AndAlso dtTrabajoAnterior.Rows.Count > 0 Then
                    If data.TipoGastoAnterior <> -1 Then
                        Select Case data.TipoGastoAnterior
                            Case enumfclTipoGastoObra.enumfclMaterial
                                dtTrabajoAnterior.Rows(0)("ImpRealMatA") = dtTrabajoAnterior.Rows(0)("ImpRealMatA") + data.Control.ImporteA - data.Control.ImporteAAnterior
                            Case enumfclTipoGastoObra.enumfclGastos
                                dtTrabajoAnterior.Rows(0)("ImpRealGastosA") = dtTrabajoAnterior.Rows(0)("ImpRealGastosA") + data.Control.ImporteA - data.Control.ImporteAAnterior
                            Case enumfclTipoGastoObra.enumfclVarios
                                dtTrabajoAnterior.Rows(0)("ImpRealVariosA") = dtTrabajoAnterior.Rows(0)("ImpRealVariosA") + data.Control.ImporteA - data.Control.ImporteAAnterior
                        End Select
                    End If
                    dtTrabajoAnterior.Rows(0)("ImpRealTrabajoA") = dtTrabajoAnterior.Rows(0)("ImpRealTrabajoA") - data.Control.ImporteAAnterior
                    BusinessHelper.UpdateTable(dtTrabajoAnterior)
                End If
            End If
        End If

        If Length(data.row("IDTrabajo")) > 0 Then
            If IDTrabajoAnterior = 0 OrElse (data.row("IDTrabajo") <> IDTrabajoAnterior) Then
                '//Actualizar Trabajo Actual
                Dim dtTrabajo As DataTable = BusinessHelper.CreateBusinessObject("ObraTrabajo").SelOnPrimaryKey(data.row("IDTrabajo"))
                If Not IsNothing(dtTrabajo) AndAlso dtTrabajo.Rows.Count > 0 Then
                    Select Case data.TipoGasto
                        Case enumfclTipoGastoObra.enumfclMaterial
                            dtTrabajo.Rows(0)("ImpRealMatA") = dtTrabajo.Rows(0)("ImpRealMatA") + data.ImporteA - data.ImporteAAnterior
                        Case enumfclTipoGastoObra.enumfclGastos
                            dtTrabajo.Rows(0)("ImpRealGastosA") = dtTrabajo.Rows(0)("ImpRealGastosA") + data.ImporteA - data.ImporteAAnterior
                        Case enumfclTipoGastoObra.enumfclVarios
                            dtTrabajo.Rows(0)("ImpRealVariosA") = dtTrabajo.Rows(0)("ImpRealVariosA") + data.ImporteA - data.ImporteAAnterior
                    End Select

                    dtTrabajo.Rows(0)("ImpRealTrabajoA") = dtTrabajo.Rows(0)("ImpRealTrabajoA") + data.ImporteA - data.ImporteAAnterior

                    BusinessHelper.UpdateTable(dtTrabajo)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarObraCabecera(ByVal data As dataActualizacion, ByVal services As ServiceProvider)
        Dim IDObraAnterior As Integer
        If Not data.Control Is Nothing Then
            IDObraAnterior = data.Control.IDObraAnterior
            If data.Control.IDObra <> data.Control.IDObraAnterior AndAlso data.Control.IDObraAnterior > 0 Then
                Dim dtObraAnterior As DataTable = BusinessHelper.CreateBusinessObject("ObraCabecera").SelOnPrimaryKey(data.Control.IDObraAnterior)
                If Not IsNothing(dtObraAnterior) AndAlso dtObraAnterior.Rows.Count > 0 Then
                    dtObraAnterior.Rows(0)("ImpRealA") = dtObraAnterior.Rows(0)("ImpRealA") - data.Control.ImporteAAnterior
                    Dim infoCalculoMargen As New Comunes.DatosCalculoMargen(dtObraAnterior.Rows(0)("ImpFactA"), dtObraAnterior.Rows(0)("ImpRealA"))
                    dtObraAnterior.Rows(0)("MargenRealTrabajo") = ProcessServer.ExecuteTask(Of Comunes.DatosCalculoMargen, Double)(AddressOf Comunes.CalcularMargen, infoCalculoMargen, services)
                    BusinessHelper.UpdateTable(dtObraAnterior)
                End If
            ElseIf data.Control.ImporteA <> data.Control.ImporteAAnterior AndAlso data.Control.IDObraAnterior > 0 Then
                Dim dtObraAnterior As DataTable = BusinessHelper.CreateBusinessObject("ObraCabecera").SelOnPrimaryKey(data.Control.IDObraAnterior)
                If Not IsNothing(dtObraAnterior) AndAlso dtObraAnterior.Rows.Count > 0 Then
                    dtObraAnterior.Rows(0)("ImpRealA") = dtObraAnterior.Rows(0)("ImpRealA") + data.Control.ImporteA - data.Control.ImporteAAnterior
                    Dim infoCalculoMargen As New Comunes.DatosCalculoMargen(dtObraAnterior.Rows(0)("ImpFactA"), dtObraAnterior.Rows(0)("ImpRealA"))
                    dtObraAnterior.Rows(0)("MargenRealTrabajo") = ProcessServer.ExecuteTask(Of Comunes.DatosCalculoMargen, Double)(AddressOf Comunes.CalcularMargen, infoCalculoMargen, services)
                    BusinessHelper.UpdateTable(dtObraAnterior)
                End If
            End If
        End If

        If Length(data.row("IDObra")) > 0 Then
            If IDObraAnterior = 0 OrElse data.row("IDObra") <> IDObraAnterior Then
                Dim dtObra As DataTable = BusinessHelper.CreateBusinessObject("ObraCabecera").SelOnPrimaryKey(data.row("IDObra"))
                If Not IsNothing(dtObra) AndAlso dtObra.Rows.Count > 0 Then
                    dtObra.Rows(0)("ImpRealA") = dtObra.Rows(0)("ImpRealA") + data.ImporteA - data.ImporteAAnterior
                    Dim infoCalculoMargen As New Comunes.DatosCalculoMargen(dtObra.Rows(0)("ImpFactA"), dtObra.Rows(0)("ImpRealA"))
                    dtObra.Rows(0)("MargenRealTrabajo") = ProcessServer.ExecuteTask(Of Comunes.DatosCalculoMargen, Double)(AddressOf Comunes.CalcularMargen, infoCalculoMargen, services)
                    BusinessHelper.UpdateTable(dtObra)
                End If
            End If
        End If
    End Sub

End Class
