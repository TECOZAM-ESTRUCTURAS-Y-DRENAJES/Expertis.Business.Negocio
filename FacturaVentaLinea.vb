Public Class FacturaVentaLinea
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

#Region " Propiedades "

    Private mblnCopiaLineaAlbaran As Boolean

    Public WriteOnly Property CopiaLineaAlbaran()
        Set(ByVal Value)
            mblnCopiaLineaAlbaran = Value
        End Set
    End Property

#End Region

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbFacturaVentaLinea"

#Region " AddNewForm "

    'Public Overrides Function AddNewForm() As DataTable
    '    Return ProcessServer.ExecuteTask(Of DataTable, DataTable)(AddressOf FillDefaultValues, Nothing, Nothing)
    'End Function

    '<Task()> Public Shared Function FillDefaultValues(ByVal DataSchema As DataTable, ByVal services As ServiceProvider) As DataTable
    '    FillDefaultValues = MyBase.AddNewForm()
    '    FillDefaultValues.Rows(0)("IdLineaFactura") = AdminData.GetAutoNumeric
    '    FillDefaultValues.Rows(0)("Dto1") = 0
    '    FillDefaultValues.Rows(0)("Dto2") = 0
    '    FillDefaultValues.Rows(0)("Dto3") = 0
    '    FillDefaultValues.Rows(0)("Dto") = 0
    '    FillDefaultValues.Rows(0)("DtoProntoPago") = 0
    '    FillDefaultValues.Rows(0)("UdValoracion") = 1
    '    FillDefaultValues.Rows(0)("Factor") = 1
    '    FillDefaultValues.Rows(0)("IDTipoLinea") = New TipoLinea().TipoLineaPorDefecto()
    '    FillDefaultValues.Rows(0)("QInterna") = 0
    '    FillDefaultValues.Rows(0)("Cantidad") = 0
    '    FillDefaultValues.Rows(0)("Regalo") = False
    'End Function

#End Region

#Region " Delete "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf Business.General.Comunes.BeginTransaction)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarLineaAlbaranDelete)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarConceptosObrasDelete)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarVencimientosTrabajoFacturacion)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarVencimientosPromoLocal)
        'deleteProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.DeleteEntityRow)
        'deleteProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoEliminado)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarControlOTDelete)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf EliminarRestriccionesEntregasACuenta)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf Business.General.Comunes.DeleteEntityRow)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf Business.General.Comunes.MarcarComoEliminado)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarLineaPromocion)
        deleteProcess.AddIfThenTask(Of DataRow)(AddressOf NoHaSidoEliminada, AddressOf ActualizarCantidadPromocionada)
    End Sub

    <Task()> Public Shared Function NoHaSidoEliminada(ByVal data As DataRow, ByVal services As ServiceProvider) As Boolean
        Dim listaEliminados As LineasFacturaEliminadas = services.GetService(Of LineasFacturaEliminadas)()
        Dim haSidoEliminado As Boolean = listaEliminados.IDLineas.Contains(data("IDLineaFactura"))
        If haSidoEliminado Then services.GetService(Of DeleteProcessContext).Deleted = haSidoEliminado
        Return Not haSidoEliminado
    End Function

    <Task()> Public Shared Sub ActualizarLineaPromocion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDPromocionLinea")) > 0 AndAlso Length(data("IDLineaPedido")) = 0 AndAlso Length(data("IDLineaAlbaran")) = 0 AndAlso data("Regalo") = 0 Then
            Dim StDatos As New PromocionLinea.DatosActuaLinPromoDr(data.Table, True)
            ProcessServer.ExecuteTask(Of PromocionLinea.DatosActuaLinPromoDr)(AddressOf PromocionLinea.ActualizarLineaPromocion, StDatos, services)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarLineaAlbaranDelete(ByVal data As DataRow, ByVal services As ServiceProvider)
        'Se cambia el estado del albarán (tanto de la cabecera como de las líneas).
        If Length(data("IDLineaAlbaran")) > 0 AndAlso Length(data("IDVencimiento")) = 0 Then
            data("Cantidad") = 0
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoFacturacionVenta.ActualizarQFacturadaLineaAlbaran, data, services)
            ProcessServer.ExecuteTask(Of Object)(AddressOf ProcesoFacturacionVenta.GrabarAlbaranes, Nothing, services)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarConceptosObrasDelete(ByVal data As DataRow, ByVal services As ServiceProvider)
        'Si proviene de una Obra se dejará tal y como estaba antes de facturar.
        If Length(data("IDObra")) > 0 Or Length(data("IDTrabajo")) > 0 Or Length(data("IDCertificacion")) > 0 Then
            Dim datos As New ProcesoFacturacionObras.DataActualizarRowConceptosObras(data, True)
            ProcessServer.ExecuteTask(Of ProcesoFacturacionObras.DataActualizarRowConceptosObras)(AddressOf ProcesoFacturacionObras.ActualizarConceptosObrasPorLinea, datos, services)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarControlOTDelete(ByVal data As DataRow, ByVal services As ServiceProvider)
        'Si proviene de una OT se dejará tal y como estaba antes de facturar.
        If Length(data("IDMntoOTControl")) > 0 Then
            Dim datos As New ProcesoFacturacionVenta.DataActualizarRowControlOT(data, True)
            ProcessServer.ExecuteTask(Of ProcesoFacturacionVenta.DataActualizarRowControlOT)(AddressOf ProcesoFacturacionVenta.ActualizarControlOT, datos, services)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarVencimientosTrabajoFacturacion(ByVal data As DataRow, ByVal services As ServiceProvider)
        'Si proviene de un Vencimiento de Obra se dejará tal y como estaba antes de facturar.
        If Length(data("IDVencimiento")) > 0 Then
            Dim f As New Filter
            f.Add(New NumberFilterItem("IDVencimiento", data("IDVencimiento")))
            ProcessServer.ExecuteTask(Of DataActualizarVencimientos)(AddressOf ActualizarVencimientos, New DataActualizarVencimientos(f, "ObraTrabajoFacturacion"), services)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarVencimientosPromoLocal(ByVal data As DataRow, ByVal services As ServiceProvider)
        'Si proviene de un Vencimiento de ObraPromo se dejará tal y como estaba antes de facturar.
        If Length(data("IDLocalVencimiento")) > 0 Then
            Dim f As New Filter
            f.Add(New NumberFilterItem("IDLocalVencimiento", data("IDLocalVencimiento")))
            ProcessServer.ExecuteTask(Of DataActualizarVencimientos)(AddressOf ActualizarVencimientos, New DataActualizarVencimientos(f, "ObraPromoLocalVencimiento"), services)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarCantidadPromocionada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDPromocionLinea")) > 0 AndAlso data("Regalo") = 0 Then
            ProcessServer.ExecuteTask(Of Integer)(AddressOf PromocionLinea.ActualizarLineaPromocionQ, data("IDPromocionLinea"), services)
        End If
    End Sub

    <Serializable()> _
    Public Class DataActualizarVencimientos
        Public Filtro As Filter
        Public Entidad As String

        Public Sub New(ByVal Filtro As Filter, ByVal Entidad As String)
            Me.Filtro = Filtro
            Me.Entidad = Entidad
        End Sub
    End Class
    <Task()> Public Shared Sub ActualizarVencimientos(ByVal data As DataActualizarVencimientos, ByVal services As ServiceProvider)
        Dim Business As BusinessHelper
        Business = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo(data.Entidad))
        Dim dt As DataTable = Business.Filter(data.Filtro)
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            dt.Rows(0)("Facturado") = 0
            dt.Rows(0)("IDFactura") = System.DBNull.Value
            dt.Rows(0)("NFactura") = System.DBNull.Value

            BusinessHelper.UpdateTable(dt)
        End If
    End Sub

    <Task()> Public Shared Sub EliminarRestriccionesEntregasACuenta(ByVal data As DataRow, ByVal services As ServiceProvider)
        '//Si proviene de una Entrega modificamos los campos necesarios de la Entrega.
        If Length(data("IDEntrega")) > 0 AndAlso Length(data("IDFactura")) > 0 Then
            Dim objNegEC As New EntregasACuenta
            Dim StDatos As New EntregasACuenta.DatosElimRestricEntFn
            StDatos.IDEntrega = data("IDEntrega")
            StDatos.IDFactura = data("IDFactura")
            StDatos.Circuito = Circuito.Ventas
            Dim blnEliminarEntregaRetencion As Boolean = ProcessServer.ExecuteTask(Of EntregasACuenta.DatosElimRestricEntFn, Boolean)(AddressOf EntregasACuenta.EliminarRestriccionesDeleteEntregaCuentaFn, StDatos, services)
            If blnEliminarEntregaRetencion Then
                Dim dtEC As DataTable = objNegEC.Filter(New NumberFilterItem("IDEntrega", data("IDEntrega")))
                If Not IsNothing(dtEC) AndAlso dtEC.Rows.Count > 0 Then
                    '//Eliminamos el vínculo con las Entregas a Cuenta, para que se puede elimar la entrega
                    If Length(data("IDEntrega")) > 0 Then
                        data("IDEntrega") = System.DBNull.Value
                        BusinessHelper.UpdateTable(data.Table)
                    End If

                    '//Eliminamos las Entregas de TipoRetención que no tengan vinculada ninguna factura.
                    ProcessServer.ExecuteTask(Of DataTable)(AddressOf EntregasACuenta.EliminarEntregasRetencionSinFactura, dtEC, services)
                End If
            End If
        End If
    End Sub

#End Region

#Region " Validate "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarArticuloBloqueado)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarFacturaContabilizada)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoComercial.DetailCommonUpdateRules)  'Validaciones Generales Comercial 
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarCantidadLineaFactura)
    End Sub


    <Task()> Public Shared Sub ValidarFacturaContabilizada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDFactura")) <> 0 Then
            Dim Cabecera As DataTable = New FacturaVentaCabecera().SelOnPrimaryKey(data("IDFactura"))
            If Not Cabecera Is Nothing AndAlso Cabecera.Rows.Count > 0 Then
                If Cabecera.Rows(0)("Estado") = enumfvcEstado.fvcContabilizado Then
                    If New Parametro().Contabilidad Then
                        ApplicationService.GenerateError("La Factura está Contabilizada.")
                    Else : ApplicationService.GenerateError("La Factura está Bloqueda y generados los vencimientos (o efectos)")
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarArticuloBloqueado(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) <> 0 Then
            Dim Cabecera As DataTable = New FacturaVentaCabecera().SelOnPrimaryKey(data("IDFactura"))
            If Not Cabecera Is Nothing AndAlso Cabecera.Rows.Count > 0 Then
                Dim StDatos As New Cliente.DataBloqArtClie
                StDatos.IDArticulo = data("IDArticulo") : StDatos.IDCliente = Cabecera.Rows(0)("IDCliente")
                If ProcessServer.ExecuteTask(Of Cliente.DataBloqArtClie, Boolean)(AddressOf Cliente.ComprobarBloqueoArticuloCliente, StDatos, services) Then
                    ApplicationService.GenerateError("El Artículo está bloqueado para este Cliente.")
                End If
            End If
        End If
    End Sub

#End Region

#Region " BUSINESS RULES "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim services As New ServiceProvider
        Dim oBRL As New BusinessRules
        oBRL = ProcessServer.ExecuteTask(Of BusinessRules, BusinessRules)(AddressOf ProcesoComercial.DetailBusinessRulesLin, oBRL, services)
        Return oBRL
    End Function

#End Region

#Region " Exportación "

    Public Function UpdateExportacion(ByVal dttSource As System.Data.DataTable) As System.Data.DataTable
        'If Not IsNothing(dttSource) AndAlso dttSource.Rows.Count Then
        '    Dim blnAdd As Boolean
        '    Dim FV As New FacturaVentaCabecera
        '    Dim com As New ProcesoComercial
        '    Dim Analitica As DataTable
        '    Dim Representantes As DataTable
        '    Dim carrier(-1) As DataTable
        '    Dim OT As BusinessHelper
        '    Dim services As New ServiceProvider
        '    Dim AppParamsConta As ParametroContabilidadVenta = services.GetService(GetType(ParametroContabilidadVenta))
        '    Me.BeginTx()
        '    For Each dr As DataRow In dttSource.Rows
        '        If dr.RowState And (DataRowState.Added Or DataRowState.Modified) Then
        '            If ProcesoComercial.DetailCommonUpdateRules(dr, services) Then
        '                If Not IsNumeric(dr("Cantidad")) Then
        '                    ApplicationService.GenerateError("La cantidad no es válida.")
        '                ElseIf dr("Cantidad") = 0 Then
        '                    ApplicationService.GenerateError("La cantidad no puede ser cero.")
        '                End If
        '                'Es necesario recalcular los importes A y B porque en Obra se usan los importes 
        '                Dim dt As DataTable = FV.Filter("IDMoneda,FechaFactura,Estado,NFactura", "IDFactura=" & dr("IDFactura"))
        '                If Not IsNothing(dt) AndAlso dt.Rows.Count Then
        '                    'If dt.Rows(0)("Estado") = enumfvcEstado.fvcContabilizado Then
        '                    '    ApplicationService.GenerateError("La Factura está Contabilizada.")
        '                    'End If
        '                    MantenimientoValoresAyB(dr, dt.Rows(0)("IDMoneda") & String.Empty, dt.Rows(0)("CambioA"), dt.Rows(0)("CambioB"))
        '                End If

        '                If dr.RowState = DataRowState.Added Then
        '                    If Length(dr("IdLineaFactura")) = 0 Then
        '                        dr("IdLineaFactura") = AdminData.GetAutoNumeric
        '                        dr("NFactura") = dt.Rows(0)("NFactura")
        '                    End If
        '                    blnAdd = True

        '                    If AppParamsConta.Contabilidad AndAlso AppParamsConta.Analitica.AplicarAnalitica Then
        '                        Analitica = NuevaAnalitica(dr)
        '                    End If

        '                    Representantes = com.NuevoRepresentante(dr)
        '                ElseIf dr.RowState = DataRowState.Modified Then
        '                    'If Not IsDBNull(dr("IDAlbaran")) And Not IsDBNull(dr("IDLineaAlbaran")) And (dr("Precio", DataRowVersion.Original) <> dr("Precio") Or dr("Cantidad", DataRowVersion.Original) <> dr("Cantidad")) Then
        '                    '    MetodosFacturacion.ActualizarAlbaran(dr, services)
        '                    'End If

        '                    If AppParamsConta.Contabilidad AndAlso AppParamsConta.Analitica.AplicarAnalitica Then Analitica = ActualizarAnalitica(dr)
        '                    Representantes = com.ActualizarRepresentantes(dr)

        '                    'If AreDifferents(dr("ImporteA"), Nz(dr("ImporteA", DataRowVersion.Original), 0)) Then
        '                    '    FV.ActualizarConceptosObras(dr)
        '                    'End If
        '                    If Nz(dr("ImporteA")) <> Nz(dr("ImporteA", DataRowVersion.Original)) Or _
        '                        Nz(dr("IDObra")) <> Nz(dr("IDObra", DataRowVersion.Original)) Or _
        '                        Nz(dr("IDTrabajo")) <> Nz(dr("IDTrabajo", DataRowVersion.Original)) Then
        '                        If (Nz(dr("IDTrabajo", DataRowVersion.Original), 0) > 0 AndAlso Length(dr("IDTrabajo")) = 0) OrElse _
        '                            (Length(dr("IDTrabajo", DataRowVersion.Original)) > 0 AndAlso Length(dr("IDTrabajo")) > 0 AndAlso dr("IDTrabajo", DataRowVersion.Original) <> dr("IDTrabajo")) Then
        '                            OT = BusinessHelper.CreateBusinessObject("ObraTrabajo")
        '                            Dim dtOT As DataTable = OT.SelOnPrimaryKey(dr("IDTrabajo", DataRowVersion.Original))
        '                            If Not IsNothing(dtOT) AndAlso dtOT.Rows.Count > 0 Then
        '                                dtOT.Rows(0)("ImpFactTrabajoA") = Nz(dtOT.Rows(0)("ImpFactTrabajoA"), 0) - dr("ImporteA")
        '                                OT.Update(dtOT)
        '                            End If
        '                        End If
        '                        Dim oPF As New ProcesoFacturacionVenta
        '                        oPF.ActualizarConceptosObras(dr)
        '                    End If
        '                End If

        '                If Not IsNothing(Analitica) Then
        '                    ReDim Preserve carrier(UBound(carrier) + 1) : carrier(UBound(carrier)) = Analitica
        '                    If Not Analitica Is Nothing Then Analitica.Dispose()
        '                End If
        '                If Not IsNothing(Representantes) Then
        '                    ReDim Preserve carrier(UBound(carrier) + 1) : carrier(UBound(carrier)) = Representantes
        '                    If Not Representantes Is Nothing Then Representantes.Dispose()
        '                End If
        '            End If
        '        End If
        '    Next
        '    AdminData.SetData(dttSource)
        '    AdminData.SetData(carrier)

        '    If blnAdd Then
        '        Dim oPF As New ProcesoFacturacionVenta
        '        oPF.ActualizarConceptosObras(dttSource)
        '    End If
        'End If

        'Return dttSource
    End Function

#End Region

End Class