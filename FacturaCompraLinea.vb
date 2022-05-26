Public Class FacturaCompraLinea
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbFacturaCompraLinea"

#Region " Delete "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.BeginTransaction)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarLineaAlbaranDelete)
        deleteProcess.AddTask(Of DataRow)(AddressOf EliminarControlObras)
        deleteProcess.AddTask(Of DataRow)(AddressOf ProcesoFacturacionCompra.DeleteOTMaterialControl)
        deleteProcess.AddTask(Of DataRow)(AddressOf EliminarRestriccionesEntregasACuenta)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarBodegaEntradaVto)
    End Sub

    <Task()> Public Shared Sub ActualizarLineaAlbaranDelete(ByVal data As DataRow, ByVal services As ServiceProvider)
        'Se cambia el estado del albarán (tanto de la cabecera como de las líneas).
        If Length(data("IDLineaAlbaran")) > 0 Then
            data("Cantidad") = 0
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoFacturacionCompra.ActualizarQFacturadaLineaAlbaran, data, services)
            ProcessServer.ExecuteTask(Of Object)(AddressOf ProcesoFacturacionCompra.GrabarAlbaranes, Nothing, services)
        End If
    End Sub

    <Task()> Public Shared Sub EliminarRestriccionesEntregasACuenta(ByVal data As DataRow, ByVal services As ServiceProvider)
        '//Si proviene de una Entrega modificamos los campos necesarios de la Entrega.
        If Length(data("IDEntrega")) > 0 AndAlso Length(data("IDFactura")) > 0 Then
            Dim objNegEC As New EntregasACuenta
            Dim StDatos As New EntregasACuenta.DatosElimRestricEntFn
            StDatos.IDEntrega = data("IDEntrega")
            StDatos.IDFactura = data("IDFactura")
            StDatos.Circuito = Circuito.Compras
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

    <Task()> Public Shared Sub EliminarControlObras(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDObra")) > 0 Then
            Dim GeneradoControl As Boolean = ProcessServer.ExecuteTask(Of DataRow, Boolean)(AddressOf ActualizacionControlObras.AlbaranGeneradoControl, data, services)
            If Not GeneradoControl Then
                Dim dataDelete As New ActualizacionControlObras.dataDeleteControlObras(data, ActualizacionControlObras.enumOrigen.Factura)
                Select Case CType(Nz(data("TipoGastoObra"), 0), enumfclTipoGastoObra)
                    Case enumfclTipoGastoObra.enumfclMaterial
                        ProcessServer.ExecuteTask(Of ActualizacionControlObras.dataDeleteControlObras)(AddressOf ActualizacionControlObras.DeleteObraMaterialControl, dataDelete, services)
                    Case enumfclTipoGastoObra.enumfclGastos
                        ProcessServer.ExecuteTask(Of ActualizacionControlObras.dataDeleteControlObras)(AddressOf ActualizacionControlObras.DeleteObraGastoControl, dataDelete, services)
                    Case enumfclTipoGastoObra.enumfclVarios
                        ProcessServer.ExecuteTask(Of ActualizacionControlObras.dataDeleteControlObras)(AddressOf ActualizacionControlObras.DeleteObraVariosControl, dataDelete, services)
                End Select
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarBodegaEntradaVto(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim ClsBdg As BusinessHelper = CreateBusinessObject("BdgEntradaVto")

        Dim DtBdg As DataTable
        DtBdg = ClsBdg.Filter(New FilterItem("IDLineaFactura", FilterOperator.Equal, data("IDLineaFactura")))
        If Not DtBdg Is Nothing AndAlso DtBdg.Rows.Count > 0 Then
            For Each Dr As DataRow In DtBdg.Select
                Dr("IDLineaFactura") = DBNull.Value
            Next
            AdminData.SetData(DtBdg)
        End If

        DtBdg = ClsBdg.Filter(New FilterItem("IDLineaFacturaExc", FilterOperator.Equal, data("IDLineaFactura")))
        If Not DtBdg Is Nothing AndAlso DtBdg.Rows.Count > 0 Then
            For Each Dr As DataRow In DtBdg.Select
                Dr("IDLineaFacturaExc") = DBNull.Value
            Next
            AdminData.SetData(DtBdg)
        End If

        DtBdg = ClsBdg.Filter(New FilterItem("IDLineaFacturaO", FilterOperator.Equal, data("IDLineaFactura")))
        If Not DtBdg Is Nothing AndAlso DtBdg.Rows.Count > 0 Then
            For Each Dr As DataRow In DtBdg.Select
                Dr("IDLineaFacturaO") = DBNull.Value
            Next
            AdminData.SetData(DtBdg)
        End If

        DtBdg = ClsBdg.Filter(New FilterItem("IDLineaFacturaSO", FilterOperator.Equal, data("IDLineaFactura")))
        If Not DtBdg Is Nothing AndAlso DtBdg.Rows.Count > 0 Then
            For Each Dr As DataRow In DtBdg.Select
                Dr("IDLineaFacturaSO") = DBNull.Value
            Next
            AdminData.SetData(DtBdg)
        End If

    End Sub
#End Region

#Region "Validate"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarFacturaContabilizada)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoCompra.DetailCommonUpdateRules)  'Validaciones Generales Compras 
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarCantidadLineaFactura)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoFacturacionCompra.ValidarConceptosGastosObra)
    End Sub



    <Task()> Public Shared Sub ValidarFacturaContabilizada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDFactura")) <> 0 Then
            Dim Cabecera As DataTable = New FacturaCompraCabecera().SelOnPrimaryKey(data("IDFactura"))
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

#End Region

#Region " Business Rules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim services As New ServiceProvider

        Dim oBRL As New BusinessRules
        '//BusinessRules - Genéricas del circuito de compras
        oBRL = ProcessServer.ExecuteTask(Of BusinessRules, BusinessRules)(AddressOf ProcesoCompra.DetailBusinessRulesLin, oBRL, services)


        ''//BusinessRules - Específicas FCL  
        oBRL.Add("TipoGastoObra", AddressOf CambioTipoGastoObra)
        oBRL.Add("IdActivoAImputar", AddressOf CambioActivoAImputar)
        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioTipoGastoObra(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("TipoGastoObra")) = 0 Then data.Current("TipoGastoObra") = enumfclTipoGastoObra.enumfclMaterial
        data.Current("IDLineaPadre") = System.DBNull.Value
        If data.Current("TipoGastoObra") <> enumfclTipoGastoObra.enumfclMaterial Then
            data.Current("IDConcepto") = System.DBNull.Value
        Else
            data.Current("IDConcepto") = data.Current("IdArticulo")
        End If
    End Sub

    <Task()> Public Shared Sub CambioActivoAImputar(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDActivoAImputar")) > 0 Then
            Dim dt As DataTable = New Activo().SelOnPrimaryKey(data.Current("IDActivoAImputar"))
            If Not dt Is Nothing AndAlso dt.Rows.Count = 0 Then
                ApplicationService.GenerateError("El Activo no existe.")
            End If
        End If
    End Sub

#End Region

    <Task()> Public Shared Function ExistenLineasInmovilizadas(ByVal strIDProcess As String, ByVal services As ServiceProvider) As Boolean
        Dim dt As DataTable = New BE.DataEngine().Filter("FacturaCompraInmovilizada", New FilterItem("IDProcess", FilterOperator.Equal, strIDProcess))
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            ExistenLineasInmovilizadas = True
        Else
            ExistenLineasInmovilizadas = False
        End If
    End Function

#Region " Consultas interactivas (Estadísticas) "

    <Task()> Public Shared Function ObtenerDatosDeclaracionCompraLineaIntracomunitaria(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim strWhere As String = "NDeclaracionIVA IS NOT NULL OR AñoDeclaracionIVA IS NOT NULL"
        Dim FilWhere As New Filter
        FilWhere.Add(New IsNullFilterItem("NDeclaracionIVA", False))
        FilWhere.Add(New IsNullFilterItem("AñoDeclaracionIVA", False))
        Dim dt As DataTable = New BE.DataEngine().Filter("vCtlCIIVACompra", FilWhere, "NFacturaIva")
        Return dt
    End Function

#End Region

#Region " Exportación "
    'TODO
    'Public Function UpdateExportacion(ByVal dttSource As System.Data.DataTable) As System.Data.DataTable
    '    If Not IsNothing(dttSource) AndAlso dttSource.Rows.Count Then
    '        Dim blnAdd As Boolean
    '        'Dim blnGenerarAnaliticaGestion As Boolean
    '        Dim strFrom As String = String.Empty
    '        Dim strEntidad As String = String.Empty
    '        Dim dt As DataTable
    '        Dim FCC As New FacturaCompraCabecera
    '        Dim AP As New ArticuloProveedor
    '        Dim AU As New ArticuloUnidadAB
    '        Dim com As New Compra
    '        Dim CG As New CentroGestion
    '        Dim A As New Activo
    '        'Dim GenerarAnalitica As Boolean
    '        Dim Analitica As New DataTable
    '        Dim carrier(-1) As DataTable
    '        'Dim p As New Parametro
    '        'GenerarAnalitica = p.CAnaliticPredet
    '        'blnGenerarAnaliticaGestion = p.CAnaliticGestion
    '        Dim services As New ServiceProvider
    '        Dim AppParamsConta As ParametroContabilidadCompra = services.GetService(GetType(ParametroContabilidadCompra))
    '        Me.BeginTx()
    '        For Each dr As DataRow In dttSource.Rows
    '            If dr.RowState And (DataRowState.Added Or DataRowState.Modified) Then
    '                If com.DetailCommonUpdateRules(dr, services) Then
    '                    'Se comprueba la existencia de la factura que corresponde a la línea introducida

    '                    Dim dtFacturaCab As DataTable = FCC.SelOnPrimaryKey(dr("IDFactura"))
    '                    If IsNothing(dtFacturaCab) OrElse dtFacturaCab.Rows.Count = 0 Then
    '                        ApplicationService.GenerateError("La Factura de la línea introducida no existe.")
    '                    Else
    '                        'If dtFacturaCab.Rows(0)("Estado") = enumfccEstado.fccContabilizado Then
    '                        '    ApplicationService.GenerateError("La Factura está Contabilizada.")
    '                        'End If
    '                        MantenimientoValoresAyB(dr, dtFacturaCab.Rows(0)("IDMoneda"), dtFacturaCab.Rows(0)("CambioA"), dtFacturaCab.Rows(0)("CambioB"))
    '                    End If

    '                    If Length(dr("IDCentroGestion")) = 0 Then
    '                        ApplicationService.GenerateError("El Centro de Gestión es obligatorio.")
    '                    Else
    '                        dt = CG.SelOnPrimaryKey(dr("IdCentroGestion"))
    '                        If IsNothing(dt) AndAlso dt.Rows.Count = 0 Then
    '                            dt.Dispose()
    '                            ApplicationService.GenerateError("El Centro Gestión no existe.")
    '                        End If
    '                    End If
    '                    If Length(dr("IdActivoAImputar")) > 0 Then
    '                        dt = A.SelOnPrimaryKey(dr("IdActivoAImputar"))
    '                        If IsNothing(dt) AndAlso dt.Rows.Count = 0 Then
    '                            dt.Dispose()
    '                            ApplicationService.GenerateError("El Activo no existe.") ' 1282
    '                        End If
    '                    End If
    '                    If Length(dr("IDConcepto")) > 0 Then
    '                        Dim FilWhere As New Filter
    '                        If dr("TipoGastoObra") = enumfclTipoGastoObra.enumfclGastos Then
    '                            FilWhere.Add("IDGasto", FilterOperator.Equal, dr("IDConcepto"))
    '                            strFrom = "tbMaestroGasto"
    '                            strEntidad = "Gastos"
    '                        ElseIf dr("TipoGastoObra") = enumfclTipoGastoObra.enumfclVarios Then
    '                            FilWhere.Add("IDVarios", FilterOperator.Equal, dr("IDConcepto"))
    '                            strFrom = "tbMaestroVarios"
    '                            strEntidad = "Varios"
    '                        Else
    '                            strFrom = String.Empty
    '                        End If
    '                        If Len(strFrom) > 0 Then
    '                            dt = AdminData.GetData(strFrom, FilWhere)
    '                            If IsNothing(dt) OrElse dt.Rows.Count = 0 Then
    '                                ApplicationService.GenerateError("El Concepto introducido no existe en |.", strEntidad)
    '                            End If
    '                        End If
    '                    End If

    '                    If Length(dr("RefProveedor")) > 0 Then
    '                        Dim f As New Filter
    '                        f.Add(New StringFilterItem("IDArticulo", dr("IDArticulo")))
    '                        f.Add(New StringFilterItem("RefProveedor", dr("RefProveedor")))
    '                        Dim dtRefProv As DataTable = AP.Filter(f)
    '                        If Not dtRefProv Is Nothing AndAlso dtRefProv.Rows.Count > 0 Then
    '                            dr("DescRefProveedor") = dtRefProv.Rows(0)("DescRefProveedor")
    '                        End If
    '                    End If

    '                    '///Validacion del factor de conversion y mantenimiento de las dos cantidades
    '                    If IsEmptyValue(dr("Factor")) Or dr("Factor") = 0 Then
    '                        dr("Factor") = AU.FactorDeConversion(dr("IDArticulo") & String.Empty, dr("IDUDMedida") & String.Empty, dr("IDUDInterna"))
    '                    End If
    '                    '(Asegurar la coherencia entre las dos cantidades)
    '                    If IsEmptyValue(dr("Cantidad")) Then dr("Cantidad") = 0
    '                    dr("QInterna") = dr("Factor") * dr("Cantidad")

    '                    If Length(dr("TipoGastoObra")) = 0 Then dr("TipoGastoObra") = enumfclTipoGastoObra.enumfclMaterial

    '                    If dr.RowState = DataRowState.Added Then
    '                        If IsDBNull(dr("IdLineaFactura")) OrElse Nz(dr("IdLineaFactura"), 0) = 0 Then
    '                            dr("IdLineaFactura") = AdminData.GetAutoNumeric
    '                            dr("NFactura") = dtFacturaCab.Rows(0)("NFactura")
    '                        End If
    '                        blnAdd = True

    '                        If AppParamsConta.Contabilidad AndAlso AppParamsConta.Analitica.AplicarAnalitica Then
    '                            Analitica = NuevaAnalitica(dr)
    '                        ElseIf AppParamsConta.Analitica.AnaliticaCentroGestion Then
    '                            'CUIDADO CON ESTO: La analitica por C.Gestion tiene que tener activado el de Aplicar Analitica
    '                            '**************PENDIENTE DE CAMBIAR
    '                            'Analitica = PrepararFacturaAnaliticaGestion(dr)
    '                            '**************FIN PENDIENTE DE CAMBIAR
    '                        End If
    '                    ElseIf dr.RowState = DataRowState.Modified Then
    '                        If (AreDifferents(dr("ImporteA"), Nz(dr("ImporteA", DataRowVersion.Original), 0))) _
    '                         Or (AreDifferents(dr("TipoGastoObra"), Nz(dr("TipoGastoObra", DataRowVersion.Original), 0))) _
    '                         Or (AreDifferents(dr("IDTrabajo"), Nz(dr("IDTrabajo", DataRowVersion.Original), 0))) _
    '                         Or (AreDifferents(dr("IDLineaPadre"), Nz(dr("IDLineaPadre", DataRowVersion.Original), 0))) Then
    '                            CrearConceptosObraControl(dtFacturaCab.Rows(0), dr)
    '                        End If
    '                        If AppParamsConta.Contabilidad AndAlso AppParamsConta.Analitica.AplicarAnalitica Then
    '                            Analitica = ActualizarAnalitica(dr)
    '                        ElseIf AppParamsConta.Analitica.AnaliticaCentroGestion Then
    '                            'CUIDADO CON ESTO: La analitica por C.Gestion tiene que tener activado el de Aplicar Analitica

    '                            '**************PENDIENTE DE CAMBIAR
    '                            'Analitica = PrepararFacturaAnaliticaGestion(dr)
    '                            '**************FIN PENDIENTE DE CAMBIAR
    '                        End If
    '                    End If
    '                    If Not IsNothing(Analitica) Then
    '                        ReDim Preserve carrier(UBound(carrier) + 1) : carrier(UBound(carrier)) = Analitica
    '                        Analitica.Dispose()
    '                    End If

    '                    '**************PENDIENTE DE CAMBIAR
    '                    'If Length(dr("IDMntoOTPrev")) Then ActualizarControlMaterialesOT(rs)
    '                    '**************FIN PENDIENTE DE CAMBIAR


    '                    'If Not Compare(dr, "Precio") And Not Length(dr("IDAlbaran") = 0) And Not Length(dr("IDLineaAlbaran") = 0) Then
    '                    ' RecalcularAlbaran(dtFacturaCab, dr)
    '                    'End If
    '                End If
    '            End If
    '            If dr.RowState = DataRowState.Modified Then
    '                If Not IsDBNull(dr("IDAlbaran")) And Not IsDBNull(dr("IDLineaAlbaran")) And (dr("Precio", DataRowVersion.Original) <> dr("Precio") Or dr("Cantidad", DataRowVersion.Original) <> dr("Cantidad")) Then
    '                    ProcesoFacturacionCompra.ActualizarAlbaran(dr)
    '                End If
    '            End If
    '        Next
    '        AdminData.SetData(dttSource)
    '        AdminData.SetData(carrier)
    '        If blnAdd Then CrearConceptosObraControl(dttSource)
    '    End If
    '    Return dttSource
    'End Function

#End Region

End Class