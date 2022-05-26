Public Enum enumResultadoCambioPagos
    Ok = 0
    PagadoNoContabilizado = 1
End Enum

Public Class Pago
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbPago"

#Region " RegisterAddNewTasks "

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarIdentificador, data, services)
    End Sub

    <Task()> Public Shared Sub AsignarIdentificador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDPago")) = 0 Then data("IDPago") = AdminData.GetAutoNumeric
    End Sub

#End Region

#Region " Delete "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.BeginTransaction)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizaFechaVencimientoCabecera)
        deleteProcess.AddTask(Of DataRow)(AddressOf EliminarPagoDesdeCobro)
        deleteProcess.AddTask(Of DataRow)(AddressOf EliminarEntregasACuenta)
        deleteProcess.AddTask(Of DataRow)(AddressOf CambiarEstadoCobro)
    End Sub

    <Task()> Public Shared Sub CambiarEstadoCobro(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Nz(data("IDCobro"), -1) <> -1 Then
            '//si el Cobro está en estado GeneradoPago, lo volvemos a la situación anterior
            Dim c As New Cobro
            Dim f As New Filter
            f.Add(New NumberFilterItem("IDCobro", data("IDCobro")))
            f.Add(New NumberFilterItem("Situacion", enumCobroSituacion.GeneradoPago))
            Dim dtCobro As DataTable = c.SelOnPrimaryKey(data("IDCobro"))
            If dtCobro.Rows.Count > 0 Then
                dtCobro.Rows(0)("Situacion") = enumCobroSituacion.NoNegociado
            End If
            c.Update(dtCobro)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizaFechaVencimientoCabecera(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDFactura")) > 0 Then
            Dim fcc As New FacturaCompraCabecera
            Dim dtFCC As DataTable = fcc.SelOnPrimaryKey(data("IDFactura"))
            If Not dtFCC Is Nothing AndAlso dtFCC.Rows.Count > 0 Then
                If dtFCC.Rows(0)("VencimientosManuales") Then
                    Dim dtPago As DataTable = New Pago().Filter(New NumberFilterItem("IDFactura", data("IDFactura")), "FechaVencimiento")
                    If Not dtPago Is Nothing AndAlso dtPago.Rows.Count > 0 Then
                        dtFCC.Rows(0)("FechaVencimiento") = dtPago.Rows(0)("FechaVencimiento")
                    Else
                        dtFCC.Rows(0)("FechaVencimiento") = System.DBNull.Value
                    End If
                    BusinessHelper.UpdateTable(dtFCC)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub EliminarPagoDesdeCobro(ByVal data As DataRow, ByVal services As ServiceProvider)
        '//Si el pago a eliminar proviene de un Cobro. Debemos dejar el cobro origen en situación No Negociado
        If Length(data("IDCobro")) > 0 Then
            Dim AppParams As ParametroTesoreria = services.GetService(Of ParametroTesoreria)()
            If data("IDTipoPago") = AppParams.TipoPagoDesdeCobro Then
                Dim objNegCobro As New Cobro
                Dim dtCobro As DataTable = objNegCobro.SelOnPrimaryKey(data("IDCobro"))
                If Not IsNothing(dtCobro) AndAlso dtCobro.Rows.Count > 0 Then
                    dtCobro.Rows(0)("Situacion") = enumCobroSituacion.NoNegociado
                    dtCobro.Rows(0)("IDPago") = System.DBNull.Value
                    objNegCobro.Update(dtCobro)
                End If
            Else
                ApplicationService.GenerateError("No puede eliminarse el Pago. Debe anular primero el Cobro origen.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub EliminarEntregasACuenta(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim blnDelete As Boolean
        '//Si PROVIENE de una ENTREGA y está VINCULADO a una FACTURA, modificamos los campos necesarios de la Entrega.
        If Length(data("IDEntrega")) > 0 AndAlso Length(data("IDFactura")) > 0 Then
            '//Eliminamos los vínculos del pago de la Factura y la Entrega de la que proviene.
            Dim objNegEC As New EntregasACuenta
            Dim StDatos As New EntregasACuenta.DatosElimRestricEntFn
            StDatos.IDFactura = data("IDFactura")
            StDatos.IDEntrega = data("IDEntrega")
            StDatos.Circuito = Circuito.Compras
            ProcessServer.ExecuteTask(Of EntregasACuenta.DatosElimRestricEntFn, Boolean)(AddressOf EntregasACuenta.EliminarRestriccionesDeleteEntregaCuentaFn, StDatos, services)
            '//NO BORRAMOS EL PAGO, lo desvinculamos de la Factura. 
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf DesvincularPagoDeFactura, data, services)
            blnDelete = False
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoEliminado, data, services)
        ElseIf Length(data("IDEntrega")) > 0 AndAlso Length(data("IDFactura")) = 0 Then
            '//Si PROVIENE de una ENTREGA y NO está VINCULADO a una FACTURA.
            Dim StDatos As New EntregasACuenta.DatosElimRestricEntCobro
            StDatos.IDEntrega = data("IDEntrega")
            StDatos.IDCobroPago = data("IDPago")
            ProcessServer.ExecuteTask(Of EntregasACuenta.DatosElimRestricEntCobro)(AddressOf EntregasACuenta.EliminarRestriccionesDeleteEntregaCuentaCobroPago, StDatos, services)
        End If
        If blnDelete Then
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf Business.General.Comunes.DeleteEntityRow, data, services)
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoEliminado, data, services)
        End If
    End Sub

    <Task()> Public Shared Sub DesvincularPagoDeFactura(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("IDFactura") = System.DBNull.Value
        Dim dtModif As DataTable = data.Table.Clone
        dtModif.ImportRow(data)
        Dim p As New Pago
        p.Update(dtModif)
    End Sub

    <Task()> Public Shared Sub DeletePagoManual(ByVal IDPagos() As Object, ByVal services As ServiceProvider)
        If Not IDPagos Is Nothing AndAlso IDPagos.Length > 0 Then
            Dim fFilter As New Filter
            fFilter.Add("Contabilizado", enumPagoContabilizado.PagoNoContabilizado)
            fFilter.Add("Situacion", enumPagoSituacion.NoPagado)
            fFilter.Add(New InListFilterItem("IDPago", IDPagos, FilterType.Numeric))

            Dim p As New Pago
            Dim dtPagos As DataTable = p.Filter(fFilter)
            If dtPagos.Rows.Count > 0 Then
                Dim dtPagosDel As DataTable = dtPagos.Clone
                For Each Dr As DataRow In dtPagos.Select
                    If Length(Dr("IDFactura")) = 0 Then
                        Dim dtPagoAgrupado As DataTable = p.Filter(New NumberFilterItem("IDPagoAgrupado", Dr("IDPago")))
                        If dtPagoAgrupado Is Nothing OrElse dtPagoAgrupado.Rows.Count = 0 Then
                            dtPagosDel.ImportRow(Dr)
                        End If
                    End If
                Next
                p.Delete(dtPagosDel)
            End If
        End If
    End Sub

#End Region

#Region " Validate "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFormaPagoObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosPagoNormal)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFechaVencimientoObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificador)
    End Sub

    ''' Validación para Pagos que no vienen desde Cobros
    <Task()> Public Shared Sub ValidarDatosPagoNormal(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarTipoPagoPredeterminadoFC, data, services)
        Dim AppParams As ParametroTesoreria = services.GetService(Of ParametroTesoreria)()
        If data("IDTipoPago") <> AppParams.TipoPagoDesdeCobro Then
            If Length(data("IDPagoPeriodo")) = 0 Then ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.ValidarProveedorObligatorio, data, services)
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.ValidarCContableObligatoria, data, services)
        End If
    End Sub

    <Task()> Public Shared Sub AsignarTipoPagoPredeterminadoFC(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipoPago")) = 0 Then
            Dim AppParams As ParametroTesoreria = services.GetService(Of ParametroTesoreria)()
            data("IDTipoPago") = AppParams.TipoPagoFC
        End If
    End Sub

#End Region

#Region " Update "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.BeginTransaction)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizaVencimientosManuales)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarTipoPagoPredeterminadoFC)
        updateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.AsignarMonedaPredeterminada)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificador)
        updateProcess.AddTask(Of DataRow)(AddressOf CambioSituacionCobroAsociado)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarValoresAyB)
    End Sub

    <Task()> Public Shared Sub CambioSituacionCobroAsociado(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified AndAlso Length(data("IDCobro")) > 0 Then
            If data("Situacion") = enumPagoSituacion.Pagado AndAlso data("Situacion", DataRowVersion.Original) <> enumPagoSituacion.Pagado Then
                Dim dtCobro As DataTable = New Cobro().SelOnPrimaryKey(data("IDCobro"))
                Dim datosCambio As New Cobro.DataCambioSituacionManual
                datosCambio.Cobros = dtCobro
                datosCambio.NuevaSituacion = enumCobroSituacion.Cobrado
                ProcessServer.ExecuteTask(Of Cobro.DataCambioSituacionManual)(AddressOf Cobro.CambioSituacionManual, datosCambio, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarValoresAyB(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(data), data("IDMoneda"), data("CambioA"), data("CambioB"))
        ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
    End Sub

    <Serializable()> _
    Public Class DataAsociarPagoARemesa
        Public IDProcess As Guid
        Public IDRemesa As Integer
        Public IDBancoPropio As String
        Public Impreso As Boolean
        Public PagosRemesar As DataTable
        Public NuevaSituacion As enumPagoSituacion
    End Class

    <Task()> Public Shared Sub AsociarPagoARemesa(ByVal data As DataAsociarPagoARemesa, ByVal services As ServiceProvider)
        If Not data.PagosRemesar Is Nothing AndAlso data.PagosRemesar.Rows.Count > 0 Then
            For Each dr As DataRow In data.PagosRemesar.Rows
                dr("IdRemesa") = data.IDRemesa
                dr("IDBancoPropio") = IIf(Length(data.IDBancoPropio) = 0, DBNull.Value, data.IDBancoPropio)
                dr("Impreso") = data.Impreso
                If data.NuevaSituacion <> -1 Then dr("Situacion") = data.NuevaSituacion
            Next

            data.PagosRemesar.TableName = GetType(Pago).Name
            BusinessHelper.UpdateTable(data.PagosRemesar)
        End If
    End Sub

#End Region

#Region " BUSINESS RULES "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("IDProveedor", AddressOf CambioProveedor)
        oBRL.Add("CContable", AddressOf NegocioGeneral.FormatoCuentaContable)
        oBRL.Add("ImpInteresPeriodo", AddressOf CambioImporteInteresPeriodo)
        oBRL.Add("ImpVencimiento", AddressOf NegocioGeneral.CambioImporteVencimiento)
        oBRL.Add("ImpVencimientoA", AddressOf NegocioGeneral.CambioImporteVencimiento)
        oBRL.Add("ImpVencimientoB", AddressOf NegocioGeneral.CambioImporteVencimiento)
        oBRL.Add("IDMoneda", AddressOf CambioMonedaFechaVto)
        oBRL.Add("FechaVencimiento", AddressOf CambioMonedaFechaVto)
        oBRL.Add("CambioA", AddressOf NegocioGeneral.CambioEnCambiosMoneda)
        oBRL.Add("CambioB", AddressOf NegocioGeneral.CambioEnCambiosMoneda)
        oBRL.Add("IDProveedorBanco", AddressOf CambioProveedorBanco)
        oBRL.Add("ImpInteres", AddressOf CambioImporteInteresLeasing)
        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioProveedor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDProveedor")) > 0 Then
            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
            Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data.Current("IDProveedor"))
            If Not ProvInfo Is Nothing AndAlso Length(ProvInfo.IDProveedor) > 0 Then
                Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
                If AppParamsConta.Contabilidad Then data.Current("CContable") = ProvInfo.CCProveedor
                data.Current("IDFormaPago") = ProvInfo.IDFormaPago
                data.Current("IDMoneda") = ProvInfo.IDMoneda
                data.Current("Titulo") = ProvInfo.RazonSocial
                If ProvInfo.IDProveedorBanco <> 0 Then data.Current("IDProveedorBanco") = ProvInfo.IDProveedorBanco
                Dim stDatosDirec As New ProveedorDireccion.DataDirecEnvio
                stDatosDirec.IDProveedor = data.Current("IDProveedor")
                stDatosDirec.TipoDireccion = enumpdTipoDireccion.pdDireccionPago
                Dim dtDireccion As DataTable = ProcessServer.ExecuteTask(Of ProveedorDireccion.DataDirecEnvio, DataTable)(AddressOf ProveedorDireccion.ObtenerDireccionEnvio, stDatosDirec, services)
                If Not IsNothing(dtDireccion) AndAlso dtDireccion.Rows.Count > 0 Then
                    data.Current("IDDireccion") = dtDireccion.Rows(0)("IDDireccion")
                End If
            End If
        Else
            data.Current("CContable") = System.DBNull.Value
            data.Current("IDFormaPago") = System.DBNull.Value
            data.Current("IdMoneda") = System.DBNull.Value
            data.Current("Titulo") = System.DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub CambioImporteInteresPeriodo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim dblImpBaseImponible As Double = Nz(data.Current("ImpInteresPeriodo"), 0) + Nz(data.Current("ImpAmortizacionPeriodo"), 0)
        data.Current("Importe") = dblImpBaseImponible + (dblImpBaseImponible * (Nz(data.Current("Factor"), 0) / 100))
    End Sub

    <Task()> Public Shared Sub CambioMonedaFechaVto(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioFechaVtoLeasing, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioMonedaLeasing, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.CambioMonedaFechaVto, data, services)
    End Sub

    <Task()> Public Shared Sub CambioFechaVtoLeasing(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Not data.Context Is Nothing AndAlso data.Context.ContainsKey("ModificarLeasing") AndAlso Nz(data.Context("ModificarLeasing"), False) Then
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf NegocioGeneral.ValidarFechaVencimientoObligatoria, data.Current, services)
        End If
    End Sub

    <Task()> Public Shared Sub CambioMonedaLeasing(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.ColumnName = "IDMoneda" Then
            If Not data.Context Is Nothing AndAlso data.Context.ContainsKey("ModificarLeasing") AndAlso Nz(data.Context("ModificarLeasing"), False) Then
                If Not Nz(data.Current("Contabilizado"), False) Then
                    Dim dblCambioAOld As Double = data.Current("CambioA")
                    Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                    Dim MonInfoA As MonedaInfo = Monedas.MonedaA
                    If Not MonInfoA Is Nothing Then
                        Dim dblCambioANew As Double = MonInfoA.CambioA
                        data.Current("ImpIntereses") = xRound((data.Current("ImpIntereses") * dblCambioAOld) / MonInfoA.CambioA, MonInfoA.NDecimalesImporte)
                    End If
                Else
                    ApplicationService.GenerateError("No se puede modificar el Importe ni la Moneda, porque ya está contabilizado.")
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioProveedorBanco(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Not IsNothing(data.Context) AndAlso data.Context.ContainsKey("ModificarLeasing") AndAlso Nz(data.Context("ModificarLeasing"), False) Then
            If Length(data.Current("IDProveedorBanco")) = 0 OrElse Not IsNumeric(data.Current("IDProveedorBanco")) Then
                ApplicationService.GenerateError("El campo Proveedor Banco no puede ser vacío y debe ser numérico.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioImporteInteresLeasing(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Not IsNothing(data.Context) AndAlso data.Context.ContainsKey("ModificarLeasing") AndAlso Nz(data.Context("ModificarLeasing"), False) Then
            data.Current(data.ColumnName) = data.Value
            If data.Current("ImpIntereses") Is System.DBNull.Value Then ApplicationService.GenerateError("El campo Importe de Intereses es obligatorio.")
            If Not IsNumeric(data.Current("ImpIntereses")) Then ApplicationService.GenerateError("El valor del campo Importe de Intereses ha de ser numérico.")
            If Not Nz(data.Current("Contabilizado"), False) Then
                data.Current("ImpBaseImponible") = Nz(data.Current("ImpIntereses"), 0) + Nz(data.Current("ImpAmortizacion"), 0)

                Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                Dim MonInfo As MonedaInfo = Monedas.GetMoneda(data.Current("IDMoneda"))
                If Not MonInfo Is Nothing AndAlso Length(MonInfo.ID) > 0 Then
                    Dim TiposIVA As EntityInfoCache(Of TipoIvaInfo) = services.GetService(Of EntityInfoCache(Of TipoIvaInfo))()
                    Dim TIVAInfo As TipoIvaInfo = TiposIVA.GetEntity(data.Current("IDTipoIVA"), data.Current("FechaVencimiento"))
                    If Not TIVAInfo Is Nothing AndAlso Length(TIVAInfo.IDTipoIVA) > 0 Then
                        data.Current("ImpVencimiento") = xRound(data.Current("ImpBaseImponible") + data.Current("ImpBaseImponible") * TIVAInfo.Factor / 100, MonInfo.NDecimalesImporte)
                    End If

                    Dim ValAyB As New ValoresAyB(data.Current, data.Current("IDMoneda"), Nz(data.Current("CambioA"), 0), Nz(data.Current("CambioB"), 0))
                    ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
                End If
            Else
                ApplicationService.GenerateError("No se puede modificar el importe ni la moneda porque ya está contabilizado.")
            End If
        End If
    End Sub


#End Region

#Region " Desglosar Pagos "

    <Serializable()> _
    Public Class DataDesglosarPagos
        Public IDPagoDesglosar As Integer
        Public NuevosPagos As DataTable
    End Class

    <Task()> Public Shared Sub DesglosarPagos(ByVal data As DataDesglosarPagos, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
        Dim p As New Pago
        Dim dtPago As DataTable = p.SelOnPrimaryKey(data.IDPagoDesglosar)
        p.Delete(dtPago)
        p.Update(data.NuevosPagos)
        ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.CommitTransaction, Nothing, services)
    End Sub

#End Region

#Region " Actualiza Vencimientos Manuales "

    <Task()> Public Shared Sub ActualizaVencimientosManuales(ByVal dr As DataRow, ByVal services As ServiceProvider)
        If Length(dr("IDFactura")) > 0 Then
            Dim fcc As New FacturaCompraCabecera
            Dim dtFCC As DataTable = fcc.SelOnPrimaryKey(dr("IDFactura"))
            If Not dtFCC Is Nothing AndAlso dtFCC.Rows.Count > 0 Then
                If dtFCC.Rows(0)("VencimientosManuales") Then
                    Dim strWhere As String = "IDFactura= " & dr("IDFactura")
                    Dim dtPago As DataTable = New Pago().Filter("FechaVencimiento", strWhere, "FechaVencimiento")
                    If Not dtPago Is Nothing AndAlso dtPago.Rows.Count > 0 Then
                        If dtPago.Rows(0)("FechaVencimiento") < dr("FechaVencimiento") Then
                            If dr.RowState = DataRowState.Modified Then
                                If Length(dr("FechaVencimiento", DataRowVersion.Original)) = 0 OrElse dr("FechaVencimiento", DataRowVersion.Original) = dtPago.Rows(0)("FechaVencimiento") Then
                                    dtFCC.Rows(0)("FechaVencimiento") = dr("FechaVencimiento")
                                End If
                            Else
                                dtFCC.Rows(0)("FechaVencimiento") = dtPago.Rows(0)("FechaVencimiento")
                            End If
                        Else
                            dtFCC.Rows(0)("FechaVencimiento") = dr("FechaVencimiento")
                        End If
                    Else
                        dtFCC.Rows(0)("FechaVencimiento") = dr("FechaVencimiento")
                    End If
                    BusinessHelper.UpdateTable(dtFCC)
                    If Length(dr("FechaVencimientoFactura")) = 0 Then dr("FechaVencimientoFactura") = dr("FechaVencimiento")
                    If dtFCC.Rows(0)("Estado") = enumfccEstado.fccNoContabilizado Then
                        dr("IDObra") = dtFCC.Rows(0)("IDObra")
                        Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
                        If AppParamsConta.Contabilidad Then
                            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
                            Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(dr("IDProveedor"))
                            If Length(ProvInfo.IDProveedor) > 0 Then dr("CContable") = ProvInfo.CCProveedor
                        End If
                        If Length(dtFCC.Rows(0)("IDDireccion")) > 0 Then
                            Dim cd As New ProveedorDireccion
                            Dim StDatosDirec As New ProveedorDireccion.DataDirecDe
                            StDatosDirec.IDDireccion = dtFCC.Rows(0)("IDDireccion")
                            StDatosDirec.TipoDireccion = enumpdTipoDireccion.pdDireccionPago
                            If ProcessServer.ExecuteTask(Of ProveedorDireccion.DataDirecDe, Boolean)(AddressOf ProveedorDireccion.EsDireccionDe, StDatosDirec, services) = True Then
                                dr("IDDireccion") = dtFCC.Rows(0)("IDDireccion")
                            Else
                                Dim StDatosDirecEnv As New ProveedorDireccion.DataDirecEnvio
                                StDatosDirecEnv.IDProveedor = dtFCC.Rows(0)("IDProveedor")
                                StDatosDirecEnv.TipoDireccion = enumpdTipoDireccion.pdDireccionPago
                                Dim direc As DataTable = ProcessServer.ExecuteTask(Of ProveedorDireccion.DataDirecEnvio, DataTable)(AddressOf ProveedorDireccion.ObtenerDireccionEnvio, StDatosDirecEnv, services)
                                If Not IsNothing(direc) AndAlso direc.Rows.Count Then
                                    dr("IDDireccion") = direc.Rows(0)("IDDireccion")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub

#End Region

#Region " Añadir/Quitar pagos a Remesas "

    <Serializable()> _
    Public Class DataAñadirPagosARemesa
        Public IDRemesa As Integer
        Public PagosAAñadir As DataTable
    End Class

    <Task()> Public Shared Sub AñadirPagosARemesa(ByVal data As DataAñadirPagosARemesa, ByVal services As ServiceProvider)
        If Not data.PagosAAñadir Is Nothing AndAlso data.PagosAAñadir.Rows.Count > 0 Then
            AdminData.BeginTx()

            Dim fRemesaContabilizada As New Filter
            fRemesaContabilizada.Add(New NumberFilterItem("NDocumento", data.IDRemesa))
            fRemesaContabilizada.Add(New NumberFilterItem("IDTipoApunte", enumDiarioTipoApunte.RemesaPago))
            Dim dtDiario As DataTable = New BE.DataEngine().Filter("tbDiarioContable", fRemesaContabilizada, "TOP 1 NAsiento")
            If dtDiario.Rows.Count > 0 Then
                ApplicationService.GenerateError("La Remesa está contabilizada. No se permite incluir más Pagos.")
            End If

            Dim f As New Filter
            f.Add(New NumberFilterItem("IDRemesa", data.IDRemesa))
            Dim dtRemesa As DataTable = New RemesaPago().Filter(f)
            If Not dtRemesa Is Nothing AndAlso dtRemesa.Rows.Count > 0 Then
                For Each dr As DataRow In data.PagosAAñadir.Rows
                    dr("IdRemesa") = data.IDRemesa
                Next

                BusinessHelper.UpdateTable(data.PagosAAñadir)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub RetirarPagosDeRemesas(ByVal IDPagos() As String, ByVal services As ServiceProvider)
        If Not IDPagos Is Nothing AndAlso IDPagos.Length > 0 Then
            ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
            Dim f As New Filter
            f.Add(New InListFilterItem("IDPago", IDPagos, FilterType.Numeric))
            Dim dtPagos As DataTable = New Pago().Filter(f)
            If Not dtPagos Is Nothing AndAlso dtPagos.Rows.Count > 0 Then
                For Each dr As DataRow In dtPagos.Select("Contabilizado = " & enumPagoContabilizado.PagoNoContabilizado)
                    dr("IDRemesa") = DBNull.Value
                Next
                BusinessHelper.UpdateTable(dtPagos)
            End If
        End If
    End Sub

#End Region

#Region " Pagos Agrupados "

    <Task()> Public Shared Function PagosAgrupables(ByVal criterios As Filter, ByVal services As ServiceProvider) As DataTable
        Dim strSelect As String = "Contabilizado,MIN(IdProveedor) AS IdProveedor"
        strSelect = strSelect & ",MIN(Titulo) AS Titulo,COUNT(Pagos) AS Pagos"
        strSelect = strSelect & ",SUM(ImpVencimiento) AS ImpVencimiento,MIN(IDPago) AS IDPago"
        strSelect = strSelect & ",MIN(AbrvMoneda) AS AbrvMoneda,SUM(ImpVencimientoA) AS ImpVencimientoA"

        Dim strGroupBy As String = "Contabilizado"
        Dim p As New Parametro
        p.ConfiguracionAgrupacionPagos(strGroupBy)

        Dim dt As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf EstadoPago.EstadosPagosAgrupables, Nothing, services)
        Dim fSituacionPagos As New Filter(FilterUnionOperator.Or)
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                fSituacionPagos.Add(New NumberFilterItem("Situacion", FilterOperator.Equal, dr("IDEstado")))
            Next
        End If

        Dim fWhere As New Filter
        fWhere.Add(New BooleanFilterItem("Contabilizado", FilterOperator.Equal, False))
        fWhere.Add(criterios)
        fWhere.Add(fSituacionPagos)

        Dim strWhere As String = AdminData.ComposeFilter(fWhere)
        If Len(strWhere) > 0 Then strWhere = strWhere & " "
        strWhere = strWhere & "GROUP BY " & strGroupBy & " HAVING COUNT(Pagos)>1"

        Dim dtPagos As DataTable = AdminData.Filter("vNegPagosAgrupables", strSelect, strWhere)
        For Each oCol As DataColumn In dtPagos.Columns
            oCol.ReadOnly = False
        Next

        Return dtPagos
    End Function

    <Serializable()> _
    Public Class DataResultPagosAgrupados
        Public PagosAgrupables As DataTable
        Public PropuestaPagosAgrupados As DataTable
    End Class

    <Serializable()> _
   Public Class DataPagosAgrupados
        Public PagosAgrupables As DataTable
        Public Criterios As Filter
    End Class

    <Task()> Public Shared Function PagosAgrupados(ByVal data As DataPagosAgrupados, ByVal services As ServiceProvider) As DataResultPagosAgrupados
        Dim p As New Parametro
        Dim strGroupBy As String = "Especial,Contabilizado"
        Dim fAgrupaciones As Filter = p.ConfiguracionAgrupacionPagos(data.PagosAgrupables, strGroupBy)

        Dim dt As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf EstadoPago.EstadosPagosAgrupables, Nothing, services)
        Dim fSituacionPagos As New Filter(FilterUnionOperator.Or)
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                fSituacionPagos.Add(New NumberFilterItem("Situacion", FilterOperator.Equal, dr("IDEstado")))
            Next
        End If

        Dim fWhere As New Filter
        If Not data.Criterios Is Nothing Then fWhere.Add(data.Criterios)
        fWhere.Add(New BooleanFilterItem("Contabilizado", FilterOperator.Equal, False))
        fWhere.Add(fAgrupaciones)
        fWhere.Add(fSituacionPagos)
        Dim strWhere As String = AdminData.ComposeFilter(fWhere)
        Dim result As New DataResultPagosAgrupados
        result.PagosAgrupables = AdminData.Filter("vNegPagosAgrupados", "*", strWhere)

        Dim strSelect As String = "Especial,Contabilizado,MIN(IdProveedor) AS IdProveedor"
        strSelect = strSelect & ",MIN(Titulo) AS Titulo,COUNT(Pagos) AS Pagos,MIN(IdFormaPago) AS IdFormaPago, MIN(DescFormaPago) AS DescFormaPago"
        strSelect = strSelect & ",MIN(FechaVencimiento) AS FechaVencimiento,SUM(ImpVencimiento) AS ImpVencimiento"
        strSelect = strSelect & ",MIN(AbrvMoneda) AS AbrvMoneda,SUM(ImpVencimientoA) AS ImpVencimientoA"
        strSelect = strSelect & ",MIN(CambioA) AS CambioA,MIN(CambioB) AS CambioB,MIN(IdMoneda) AS IdMoneda"
        strSelect = strSelect & ",MIN(IDDireccion) AS IDDireccion, MIN(IDProveedorBanco) AS IDProveedorBanco"
        strSelect = strSelect & ",MIN(IdCondicionPago) AS IdCondicionPago,MIN(IDBancoPropio) AS IDBancoPropio"
        Dim strWhere2 As String = AdminData.ComposeFilter(fWhere)
        If Len(strWhere2) > 0 Then strWhere2 = strWhere2 & " "
        strWhere2 = strWhere2 & "GROUP BY " & strGroupBy & " HAVING COUNT(Pagos)>1"

        result.PropuestaPagosAgrupados = AdminData.Filter("vNegNuevosPagosAgrupados", strSelect, strWhere2)
        For Each oCol As DataColumn In result.PropuestaPagosAgrupados.Columns
            oCol.ReadOnly = False
        Next
        Return result
    End Function

    <Serializable()> _
    Public Class DataAddPagosAgrupados
        Public IDProcess As Guid
        Public NuevosPagos As DataTable
    End Class
    <Task()> Public Shared Function AddPagosAgrupados(ByVal data As DataAddPagosAgrupados, ByVal services As ServiceProvider) As ClassErrors()
        Dim Errores(-1) As ClassErrors
        Dim pa As New Pago
        ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
        Dim blnError As Boolean
        Dim dtPagosSelec As DataTable = New BE.DataEngine().Filter("vNegPagosAgrupados", New GuidFilterItem("IDProcess", data.IDProcess))
        If Not dtPagosSelec Is Nothing AndAlso dtPagosSelec.Rows.Count > 0 Then
            Dim p As New Parametro
            Dim strNFactura As String = p.NFacturaPagoAgupado
            Dim strIN As String
            Dim dtPago As DataTable = pa.AddNew
            Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
            Dim AppParamsFC As ParametroFacturaCompra = services.GetService(Of ParametroFacturaCompra)()
            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
            For Each drNewPago As DataRow In data.NuevosPagos.Rows
                Dim NewRow As DataRow = dtPago.NewRow

                NewRow("IdPago") = AdminData.GetAutoNumeric
                NewRow("IDProveedor") = drNewPago("IDProveedor")
                If Nz(drNewPago("Especial"), False) Then
                    NewRow("IDTipoPago") = AppParamsFC.TipoPagoFacturaCompraB
                End If
                Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(NewRow("IDProveedor"))
                If Not ProvInfo Is Nothing AndAlso Length(ProvInfo.IDProveedor) > 0 Then
                    If AppParamsConta.Contabilidad Then
                        If Length(ProvInfo.CCProveedor) Then
                            NewRow("CContable") = ProvInfo.CCProveedor
                        Else
                            ReDim Preserve Errores(Errores.Length)
                            Errores(Errores.Length - 1) = New ClassErrors
                            Errores(Errores.Length - 1).Elements = drNewPago("IDProveedor")
                            Errores(Errores.Length - 1).MessageError = Engine.ParseFormatString(AdminData.GetMessageText("La Cuenta Contable del Proveedor {0} es un dato obligatorio."), Quoted(NewRow("IDProveedor")))

                            blnError = True
                        End If
                    End If
                End If

                NewRow("Titulo") = drNewPago("Titulo")
                NewRow("IDFormaPago") = drNewPago("IDFormaPago")
                NewRow("FechaVencimiento") = drNewPago("FechaVencimiento")
                NewRow("ImpVencimiento") = drNewPago("ImpVencimiento")
                NewRow("CambioA") = drNewPago("CambioA")
                NewRow("CambioB") = drNewPago("CambioB")
                NewRow("IDMoneda") = drNewPago("IDMoneda")
                NewRow("NFactura") = strNFactura
                NewRow("IdProveedorBanco") = drNewPago("IdProveedorBanco")
                NewRow("IDDireccion") = drNewPago("IDDireccion")
                NewRow("IDBancoPropio") = drNewPago("IDBancoPropio")

                Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(NewRow), NewRow("IDMoneda"), NewRow("CambioA"), NewRow("CambioB"))
                ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)

                If Not blnError Then
                    'Actualización del Pago
                    Dim dv As DataView = dtPagosSelec.DefaultView
                    dv.RowFilter = "IDProveedor='" & NewRow("IDProveedor") & "'"
                    If Not dv Is Nothing Then
                        For Each dr As DataRowView In dv
                            dr.Row("IDPagoAgrupado") = NewRow("IDPago")
                        Next
                    End If
                    dv.RowFilter = ""

                    dtPago.Rows.Add(NewRow)
                Else
                    blnError = False
                End If
            Next
            pa.Update(dtPago)
            dtPagosSelec.TableName = GetType(Pago).Name
            BusinessHelper.UpdateTable(dtPagosSelec)
            Return Errores
        Else
            ApplicationService.GenerateError("No existe ningún Pago seleccionado para agrupar.")
        End If
    End Function

    <Serializable()> _
    Public Class DataDesagruparPagos
        Public PagosAgrupados As DataTable
        Public PagosDesagrupados As DataTable
    End Class
    <Task()> Public Shared Sub EliminarPagoAgrupado(ByVal data As DataDesagruparPagos, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
        If Not IsNothing(data.PagosDesagrupados) AndAlso data.PagosDesagrupados.Rows.Count > 0 Then
            Dim dtPagosEliminar As DataTable = data.PagosAgrupados.Clone
            'Dim PagoAgrupado As String = New Parametro().NFacturaPagoAgupado
            'If Length(PagoAgrupado) = 0 Then ApplicationService.GenerateError("Revise el parámetro de indica el texto de Pago Agrupado.")

            Dim IDPagosAgrupadores As List(Of Object) = (From c In data.PagosAgrupados Select c("IDPago") Distinct).ToList
            Dim dtPagosAgrupados As DataTable = New BE.DataEngine().Filter("tbPago", New InListFilterItem("IDPagoAgrupado", IDPagosAgrupadores.ToArray, FilterType.Numeric), "IDPagoAgrupado")
            Dim IDPagosAgrupadoresBBDD As List(Of Object) = (From c In dtPagosAgrupados Select c("IDPagoAgrupado") Distinct).ToList

            Dim dtEstPago As DataTable = New EstadoPago().Filter(New BooleanFilterItem("Desagrupable", True), , "IDEstado")

            For Each drPagoAgrupado As DataRow In data.PagosAgrupados.Rows
                If drPagoAgrupado("Contabilizado") = CBool(enumPagoContabilizado.PagoContabilizado) Then
                    ApplicationService.GenerateError("Algún Pago Agrupado está contabilizado. Debe descontabilizar el Pago antes de desagruparlo.")
                End If
                'If drPagoAgrupado("Situacion") <> enumPagoSituacion.NoPagado Then
                '    ApplicationService.GenerateError("Algún Pago Agrupado está en una Situación distinta de No Pagado. No se puede deshacer la agrupación.")
                'End If
                If Not dtEstPago Is Nothing AndAlso dtEstPago.Rows.Count > 0 Then
                    Dim dr() As DataRow = dtEstPago.Select("IDEstado = " & drPagoAgrupado("Situacion"))
                    If dr.Length = 0 Then
                        ApplicationService.GenerateError("Algún Pago Agrupado está en una Situación que no es desagrupable. No se puede deshacer la agrupación.")
                    End If
                End If

                'If drPagoAgrupado("NFactura") & String.Empty = PagoAgrupado Then
                If Not IDPagosAgrupadoresBBDD Is Nothing AndAlso IDPagosAgrupadoresBBDD.Contains(drPagoAgrupado("IDPago")) Then
                    dtPagosEliminar.ImportRow(drPagoAgrupado)
                End If
            Next

            If data.PagosDesagrupados.Rows.Count > 0 Then
                For Each dr As DataRow In data.PagosDesagrupados.Select
                    dr("IdPagoAgrupado") = System.DBNull.Value
                Next

                data.PagosDesagrupados.TableName = GetType(Pago).Name
                BusinessHelper.UpdateTable(data.PagosDesagrupados)
                Dim p As New Pago
                p.Delete(dtPagosEliminar)
            End If
        End If
    End Sub

#End Region

#Region " Cambio Situacion del Pago "

    <Serializable()> _
    Public Class DataCambioSituacionManual
        Public Pagos As DataTable
        Public NuevaSituacion As enumPagoSituacion?
        Public NuevaFechaPago As Date?

        Public Sub New()
        End Sub

        Public Sub New(ByVal Pagos As DataTable, Optional ByVal NuevaSituacion As enumPagoSituacion = -1, Optional ByVal NuevaFechaPago As Date = cnMinDate)
            Me.Pagos = Pagos
            If NuevaSituacion <> -1 Then Me.NuevaSituacion = NuevaSituacion
            If NuevaFechaPago <> cnMinDate Then Me.NuevaFechaPago = NuevaFechaPago
        End Sub
    End Class

    <Task()> Public Shared Function CambioSituacionManual(ByVal data As DataCambioSituacionManual, ByVal services As ServiceProvider) As ClassErrors()
        If Not data.Pagos Is Nothing Then
            Dim NuevaSituacion As enumPagoSituacion
            Dim Errores(-1) As ClassErrors
            For Each dr As DataRow In data.Pagos.Rows
                Dim ResultState As enumResultadoCambioPagos = enumResultadoCambioPagos.Ok
                '//Si no se ha introducido una única situación para todos los registros, ésta deberá venir indicada en el Pago.
                If data.NuevaSituacion Is Nothing Then
                    NuevaSituacion = dr("Situacion")
                Else
                    NuevaSituacion = data.NuevaSituacion
                    dr("Situacion") = NuevaSituacion
                End If

                Dim EstadosPago As EntityInfoCache(Of EstadoPagoInfo) = services.GetService(Of EntityInfoCache(Of EstadoPagoInfo))()
                Dim EstPagInfo As EstadoPagoInfo = EstadosPago.GetEntity(NuevaSituacion)
                If Not EstPagInfo Is Nothing AndAlso Length(EstPagInfo.IDEstado) > 0 Then
                    dr("IDAgrupacion") = EstPagInfo.IDAgrupacion
                End If

                If NuevaSituacion = enumPagoSituacion.Pagado Then
                    If dr("Contabilizado") = False Then
                        ResultState = enumResultadoCambioPagos.PagadoNoContabilizado
                    End If

                    '//Si no se ha introducido una única fecha pago para todos los registros, ésta deberá venir indicada en el Pago.
                    If Not data.NuevaFechaPago Is Nothing Then
                        dr("FechaPago") = data.NuevaFechaPago
                        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf Pago.ActualizarFechaParaDeclaracionFactura, New DataRowPropertyAccessor(dr), services)
                    End If
                End If

                If ResultState <> enumResultadoCambioPagos.Ok Then
                    ReDim Preserve Errores(Errores.Length)
                    Errores(Errores.Length - 1) = New ClassErrors
                    Errores(Errores.Length - 1).Elements = dr("IDPago")
                    Errores(Errores.Length - 1).MessageError = ResultState
                End If
            Next

            data.Pagos.TableName = GetType(Pago).Name
            BusinessHelper.UpdateTable(data.Pagos)
            Return Errores
        End If
    End Function

    <Task()> Public Shared Sub ActualizarFechaPagoDesdeCobro(ByVal drCobro As IPropertyAccessor, ByVal services As ServiceProvider)
        If Length(drCobro("IDPago")) > 0 Then
            Dim dtPago As DataTable = New Pago().SelOnPrimaryKey(drCobro("IDPago"))
            If dtPago.Rows.Count > 0 Then
                Dim Situacion As enumPagoSituacion = enumPagoSituacion.Pagado
                If Nz(drCobro("FechaCobro"), cnMinDate) = cnMinDate Then
                    Situacion = enumPagoSituacion.GeneradoCobro
                    dtPago.Rows(0)("FechaPago") = System.DBNull.Value
                End If
                Dim datosCambio As New Pago.DataCambioSituacionManual(dtPago, Situacion, Nz(drCobro("FechaCobro"), cnMinDate))
                ProcessServer.ExecuteTask(Of Pago.DataCambioSituacionManual)(AddressOf Pago.CambioSituacionManual, datosCambio, services)
            End If

        End If
    End Sub


#End Region

#Region " Cambio Banco Propio "

    <Serializable()> _
    Public Class DataCambioBancoPropio
        Public IDRemesa As Integer?
        Public Pagos As DataTable
        Public NuevoBancoPropio As String
    End Class

    <Task()> Public Shared Sub CambioBancoPropio(ByVal data As DataCambioBancoPropio, ByVal services As ServiceProvider)
        If Not data.IDRemesa Is Nothing AndAlso data.IDRemesa <> 0 AndAlso (data.Pagos Is Nothing OrElse data.Pagos.Rows.Count = 0) Then
            data.Pagos = New Pago().Filter(New NumberFilterItem("IDRemesa", data.IDRemesa))
        End If
        If Not data.Pagos Is Nothing AndAlso data.Pagos.Rows.Count > 0 Then
            Dim f As New Filter
            f.Add(New BooleanFilterItem("Contabilizado", False))
            Dim WhereNotContabilizado As String = f.Compose(New AdoFilterComposer)
            For Each dr As DataRow In data.Pagos.Select(WhereNotContabilizado)
                If Length(data.NuevoBancoPropio) > 0 Then
                    dr("IDBancoPropio") = data.NuevoBancoPropio
                End If
            Next
            data.Pagos.AcceptChanges()
            data.Pagos.TableName = GetType(Pago).Name
            BusinessHelper.UpdateTable(data.Pagos)
        End If
    End Sub

#End Region


#Region " Desglose de Pago "

    <Serializable()> _
    Public Class DataInsertarDesglosePago
        Public IDPagoEliminar As Integer
        Public NuevosPagos As DataTable
    End Class
    <Task()> Public Shared Function InsertarDesglosePago(ByVal data As DataInsertarDesglosePago, ByVal services As ServiceProvider) As Integer
        ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
        Dim p As New Pago
        Dim dtAux As DataTable = p.SelOnPrimaryKey(data.IDPagoEliminar)
        If Not dtAux Is Nothing AndAlso dtAux.Rows.Count > 0 Then p.Delete(dtAux.Rows(0))
        For Each Dr As DataRow In data.NuevosPagos.Select
            Dim dt As DataTable = p.AddNewForm()
            dt.Rows(0)("IDProveedor") = dtAux.Rows(0)("IDProveedor")
            dt.Rows(0)("Titulo") = dtAux.Rows(0)("Titulo")
            dt.Rows(0)("CContable") = dtAux.Rows(0)("CContable")
            dt.Rows(0)("IDFormaPago") = Dr("IDFormaPago")
            dt.Rows(0)("ImpVencimiento") = Dr("ImpVencimiento")
            dt.Rows(0)("FechaVencimiento") = Dr("FechaVencimiento")
            dt.Rows(0)("ImpVencimientoA") = Dr("ImpVencimientoA")
            dt.Rows(0)("ImpVencimientoB") = Dr("ImpVencimientoB")
            dt.Rows(0)("CambioA") = dtAux.Rows(0)("CambioA")
            dt.Rows(0)("CambioB") = dtAux.Rows(0)("CambioB")
            dt.Rows(0)("IDMoneda") = dtAux.Rows(0)("IDMoneda")
            dt.Rows(0)("IdProveedorBanco") = dtAux.Rows(0)("IdProveedorBanco")
            dt.Rows(0)("IDFactura") = dtAux.Rows(0)("IDFactura")
            dt.Rows(0)("NFactura") = dtAux.Rows(0)("NFactura")
            dt.Rows(0)("Contabilizado") = dtAux.Rows(0)("Contabilizado")
            dt.Rows(0)("Situacion") = dtAux.Rows(0)("Situacion")
            dt.Rows(0)("IDTipoPago") = dtAux.Rows(0)("IDTipoPago")
            p.Update(dt)
        Next
    End Function

#End Region

#Region " Pago Desde Cobro "

    <Task()> Public Shared Sub InsertarPagoDesdeCobro(ByVal IDCobro As Integer, ByVal services As ServiceProvider)
        If IDCobro > 0 Then
            Dim co As New Cobro : Dim pa As New Pago
            Dim dtCobro As DataTable = co.SelOnPrimaryKey(IDCobro)
            If Not dtCobro Is Nothing AndAlso dtCobro.Rows.Count > 0 Then
                ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
                '//Generamos el pago a partir del cobro
                Dim dtPago As DataTable = pa.AddNewForm()
                Dim AppParams As ParametroTesoreria = services.GetService(Of ParametroTesoreria)()
                dtPago.Rows(0)("IdTipoPago") = AppParams.TipoPagoDesdeCobro
                'dtPago.Rows(0)("Titulo") = dtCobro.Rows(0)("Titulo")
                If Length(dtCobro.Rows(0)("NFactura")) > 0 Then
                    dtPago.Rows(0)("NFactura") = dtCobro.Rows(0)("NFactura")
                End If
                dtPago.Rows(0)("IDFormaPago") = dtCobro.Rows(0)("IDFormaPago")
                If Length(dtCobro.Rows(0)("IDBancoPropio")) > 0 Then
                    dtPago.Rows(0)("IDBancoPropio") = dtCobro.Rows(0)("IDBancoPropio")
                End If
                dtPago.Rows(0)("CContable") = dtCobro.Rows(0)("CContable")
                dtPago.Rows(0)("IDMoneda") = dtCobro.Rows(0)("IDMoneda")
                dtPago.Rows(0)("Situacion") = enumPagoSituacion.NoPagado
                dtPago.Rows(0)("FechaVencimiento") = dtCobro.Rows(0)("FechaVencimiento")
                dtPago.Rows(0)("ImpVencimiento") = -dtCobro.Rows(0)("ImpVencimiento")
                dtPago.Rows(0)("ImpVencimientoA") = -dtCobro.Rows(0)("ImpVencimientoA")
                dtPago.Rows(0)("ImpVencimientoB") = -dtCobro.Rows(0)("ImpVencimientoB")
                dtPago.Rows(0)("CambioA") = dtCobro.Rows(0)("CambioA")
                dtPago.Rows(0)("CambioB") = dtCobro.Rows(0)("CambioB")
                dtPago.Rows(0)("IDCobro") = dtCobro.Rows(0)("IDCobro")

                If Length(dtCobro.Rows(0)("IDCliente")) > 0 Then
                    Dim dtClteProv As DataTable = New BE.DataEngine().Filter("vNegClteProveedorAsoc", New StringFilterItem("IDCliente", dtCobro.Rows(0)("IDCliente")))
                    If Not dtClteProv Is Nothing AndAlso dtClteProv.Rows.Count > 0 Then
                        If Length(dtClteProv.Rows(0)("IDProveedorAsociado")) > 0 Then dtPago.Rows(0)("IDProveedor") = dtClteProv.Rows(0)("IDProveedorAsociado")
                        'If Length(dtClteProv.Rows(0)("CCProveedor")) > 0 Then dtPago.Rows(0)("CContable") = dtClteProv.Rows(0)("CCProveedor")
                        If Length(dtClteProv.Rows(0)("RazonSocial")) > 0 Then dtPago.Rows(0)("Titulo") = dtClteProv.Rows(0)("RazonSocial")
                    Else : dtPago.Rows(0)("Titulo") = dtCobro.Rows(0)("Titulo")
                    End If
                End If

                Dim intPago As Integer = dtPago.Rows(0)("IDPago")
                pa.Update(dtPago)

                '//Asociar el pago generado al cobro origen
                dtCobro.Rows(0)("IDPago") = intPago
                dtCobro.Rows(0)("Situacion") = enumCobroSituacion.GeneradoPago
                co.Update(dtCobro)
                ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.CommitTransaction, True, services)
            End If
        End If

    End Sub

#End Region

#Region " Pago Periodico "

    <Serializable()> _
    Public Class DataAddPagoPeriodico
        Public PagosPeriodicos As DataTable
        Public FechaFinal As Date
        Public Simulacion As Boolean

        Public Sub New(ByVal PagosPeriodicos As DataTable, ByVal FechaFinal As Date, Optional ByVal Simulacion As Boolean = False)
            Me.PagosPeriodicos = PagosPeriodicos
            Me.FechaFinal = FechaFinal
            Me.Simulacion = Simulacion
        End Sub
    End Class

    <Task()> Public Shared Function AddPagoPeriodico(ByVal data As DataAddPagoPeriodico, ByVal services As ServiceProvider) As DataTable
        If Not data.PagosPeriodicos Is Nothing AndAlso data.PagosPeriodicos.Rows.Count > 0 Then
            Dim p As New Pago : Dim g As New NegocioGeneral
            Dim dtNewPago As DataTable = p.AddNew()
            For Each dr As DataRow In data.PagosPeriodicos.Select
                If Nz(dr("FechaUltimaActualizacion"), Date.MinValue) < dr("FechaFin") Then
                    Dim strUnidad As String = g.GetPeriodString(dr("Unidad"))
                    Dim dtFechaComienzo As Date
                    If Length(dr("FechaUltimaActualizacion")) = 0 Then
                        dtFechaComienzo = dr("FechaInicio")
                    Else
                        dtFechaComienzo = DateAdd(strUnidad, dr("Periodo"), dr("FechaUltimaActualizacion"))
                    End If

                    Dim dtFechaTope As Date = IIf(dr("FechaFin") < data.FechaFinal, dr("FechaFin"), data.FechaFinal)
                    Dim strAgrupacion As String = dr("IDAgrupacion") & String.Empty
                    Dim intPeriodo As Integer = 0

                    Do While DateAdd(strUnidad, intPeriodo * dr("Periodo"), dtFechaComienzo) <= dtFechaTope
                        Dim drNewPago As DataRow = dtNewPago.NewRow

                        If Not data.Simulacion Then drNewPago("IDPago") = AdminData.GetAutoNumeric
                        drNewPago("Titulo") = dr("DescPago")
                        If intPeriodo = 0 Then
                            drNewPago("FechaVencimiento") = dtFechaComienzo
                        Else
                            drNewPago("FechaVencimiento") = DateAdd(strUnidad, intPeriodo * dr("Periodo"), dtFechaComienzo)
                        End If

                        drNewPago("CContable") = dr("IDCContable")
                        drNewPago("IDProveedor") = dr("IDProveedor")
                        drNewPago("IdTipoPago") = dr("IdTipoPago")
                        drNewPago("IDFormaPago") = dr("IDFormaPago")
                        drNewPago("IDBancoPropio") = dr("IDBancoPropio")

                        '// Recuperamos el Banco Predereteminado del Proveedor
                        Dim fBancoProv As New Filter
                        fBancoProv.Add(New BooleanFilterItem("Predeterminado", True))
                        fBancoProv.Add(New StringFilterItem("IDProveedor", dr("IDProveedor")))
                        Dim dtBancoProv As DataTable = New ProveedorBanco().Filter(fBancoProv)
                        If dtBancoProv.Rows.Count > 0 Then
                            drNewPago("IDProveedorBanco") = dtBancoProv.Rows(0)("IDProveedorBanco") & String.Empty
                        End If

                        drNewPago("IDMoneda") = dr("IDMoneda")
                        drNewPago("CambioA") = dr("CambioA")
                        drNewPago("CambioB") = dr("CambioB")
                        drNewPago("Situacion") = enumPagoSituacion.NoPagado
                        drNewPago("Contabilizado") = enumPagoContabilizado.PagoNoContabilizado
                        drNewPago("ImpVencimiento") = dr("Importe")
                        drNewPago("ImpRecuperacionCoste") = dr("ImpRecuperacionCostePeriodo")
                        drNewPago("ImpIntereses") = dr("ImpInteresPeriodo")
                        drNewPago("ImpCuota") = drNewPago("ImpIntereses") + drNewPago("ImpRecuperacionCoste")
                        If Length(strAgrupacion) > 0 Then drNewPago("IDAgrupacion") = strAgrupacion
                        drNewPago("IdPagoPeriodo") = dr("ID")

                        Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(drNewPago), dr("IDMoneda"), dr("CambioA"), dr("CambioB"))
                        ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)

                        intPeriodo = intPeriodo + 1

                        'Actualización FechaUltimaActualizacion del Pago periódico.
                        dr("FechaUltimaActualizacion") = drNewPago("FechaVencimiento")

                        dtNewPago.Rows.Add(drNewPago)
                    Loop
                End If
            Next

            If Not data.Simulacion AndAlso Not IsNothing(dtNewPago) AndAlso dtNewPago.Rows.Count > 0 Then
                ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
                BusinessHelper.UpdateTable(dtNewPago)
                BusinessHelper.UpdateTable(data.PagosPeriodicos)
            End If
            Return dtNewPago
        End If
    End Function

#End Region

    <Serializable()> _
    Public Class DataInsertarVencimiento
        Public IDFactura As Integer
        Public ImpVencimiento As Double
        Public FechaVencimiento As Date
        Public IDFormaPago As String
        Public RecargoFinanciero As Double
    End Class
    <Task()> Public Shared Function InsertarVencimiento(ByVal data As DataInsertarVencimiento, ByVal services As ServiceProvider) As Integer
        Dim dtFCC As DataTable = New FacturaCompraCabecera().SelOnPrimaryKey(data.IDFactura)
        Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
        Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(dtFCC.Rows(0)("IDProveedor"))
        If Not ProvInfo Is Nothing AndAlso Length(ProvInfo.IDProveedor) > 0 Then
            Dim dtFCL As DataTable = New FacturaCompraLinea().Filter(New NumberFilterItem("IDFactura", data.IDFactura))
            Dim AppParams As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
            If dtFCC.Rows.Count > 0 Then
                Dim p As New Pago
                Dim dtPago As DataTable = p.AddNewForm
                InsertarVencimiento = dtPago.Rows(0)("IDPago")
                dtPago.Rows(0)("IDFactura") = dtFCC.Rows(0)("IDFactura")
                dtPago.Rows(0)("NFactura") = dtFCC.Rows(0)("NFactura")
                dtPago.Rows(0)("IDProveedor") = dtFCC.Rows(0)("IDProveedor")
                dtPago.Rows(0)("IdPRoveedorBanco") = dtFCC.Rows(0)("IdPRoveedorBanco")
                dtPago.Rows(0)("IDBancoPropio") = dtFCC.Rows(0)("IDBancoPropio")
                If Len(dtFCC.Rows(0)("RazonSocial")) > 0 Then
                    dtPago.Rows(0)("Titulo") = dtFCC.Rows(0)("RazonSocial")
                End If

                If AppParams.Contabilidad Then
                    If Length(ProvInfo.CCProveedor) > 0 Then
                        dtPago.Rows(0)("CContable") = ProvInfo.CCProveedor
                    Else
                        ApplicationService.GenerateError("La Cuenta Contable es un campo obligatorio.")
                    End If
                End If

                dtPago.Rows(0)("Impreso") = False
                dtPago.Rows(0)("IDMoneda") = dtFCC.Rows(0)("IDMoneda")
                dtPago.Rows(0)("CambioA") = dtFCC.Rows(0)("CambioA")
                dtPago.Rows(0)("CambioB") = dtFCC.Rows(0)("CambioB")
                Dim ValAyB As New ValoresAyB(data.ImpVencimiento, dtPago.Rows(0)("IDMoneda"), dtPago.Rows(0)("CambioA"), dtPago.Rows(0)("CambioB"))
                Dim fImp As fImporte = ProcessServer.ExecuteTask(Of ValoresAyB, fImporte)(AddressOf NegocioGeneral.MantenimientoValoresImporteAyB, ValAyB, services)
                dtPago.Rows(0)("ImpVencimiento") = fImp.Importe
                dtPago.Rows(0)("ImpVencimientoA") = fImp.ImporteA
                dtPago.Rows(0)("ImpVencimientoB") = fImp.ImporteB
                dtPago.Rows(0)("FechaVencimiento") = data.FechaVencimiento
                dtPago.Rows(0)("IDFormaPago") = data.IDFormaPago
                ValAyB = New ValoresAyB(data.RecargoFinanciero, dtPago.Rows(0)("IDMoneda"), dtPago.Rows(0)("CambioA"), dtPago.Rows(0)("CambioB"))
                fImp = ProcessServer.ExecuteTask(Of ValoresAyB, fImporte)(AddressOf NegocioGeneral.MantenimientoValoresImporteAyB, ValAyB, services)
                dtPago.Rows(0)("RecargoFinanciero") = fImp.Importe
                dtPago.Rows(0)("RecargoFinancieroA") = fImp.ImporteA
                dtPago.Rows(0)("RecargoFinancieroB") = fImp.ImporteB
                dtPago.Rows(0)("NOperacion") = 0
                Select Case dtFCC.Rows(0)("IdTipoAsiento")
                    Case enumTipoAsiento.taBancoSinPago
                        dtPago.Rows(0)("Contabilizado") = enumPagoContabilizado.PagoContabilizado
                        dtPago.Rows(0)("Situacion") = enumPagoSituacion.Pagado
                    Case enumTipoAsiento.taProveedorConPagoNPyNC
                        dtPago.Rows(0)("Contabilizado") = enumPagoContabilizado.PagoNoContabilizado
                        dtPago.Rows(0)("Situacion") = enumPagoSituacion.NoPagado
                    Case enumTipoAsiento.taProveedorConPagoPyNC
                        dtPago.Rows(0)("Contabilizado") = enumPagoContabilizado.PagoNoContabilizado
                        dtPago.Rows(0)("Situacion") = enumPagoSituacion.Pagado
                    Case enumTipoAsiento.taProveedorSinPago
                        dtPago.Rows(0)("Contabilizado") = enumPagoContabilizado.PagoContabilizado
                        dtPago.Rows(0)("Situacion") = enumPagoSituacion.Pagado
                End Select

                'Obtengo la PartidaEstadistica del 1er Articulo de la factura que la lleve rellena
                'para pasarsela a los Pagos
                Dim strPartidaEstadistica As String
                If dtFCL.Rows.Count <> 0 Then
                    For Each drFCl As DataRow In dtFCL.Rows
                        Dim dtArticulo = New Articulo().SelOnPrimaryKey(drFCl("IDArticulo"))
                        If dtArticulo.Rows.Count <> 0 Then
                            If Not IsDBNull(dtArticulo.rows(0)("IDPartidaEstadistica")) AndAlso _
                               Len(dtArticulo.rows(0)("IDPartidaEstadistica")) > 0 Then
                                strPartidaEstadistica = dtArticulo.rows(0)("IDPartidaEstadistica")
                                Exit For
                            End If
                        End If
                    Next
                End If
                dtPago.Rows(0)("IDPartidaEstadistica") = strPartidaEstadistica

                p.Update(dtPago)
            End If
        End If
    End Function

    <Serializable()> _
   Public Class DataActualizarPagosImpresos
        Public IDPagos() As Object
        Public Impreso As Boolean
        Public IDBancoPropio As String
    End Class
    <Task()> Public Shared Sub ActualizarCobrosImpresos(ByVal data As DataActualizarPagosImpresos, ByVal services As ServiceProvider)
        If data.IDPagos Is Nothing OrElse data.IDPagos.Length = 0 Then Exit Sub
        Dim dtPago As DataTable = New Pago().Filter(New InListFilterItem("IDPago", data.IDPagos, FilterType.Numeric))
        If Not dtPago Is Nothing AndAlso dtPago.Rows.Count > 0 Then
            For Each Dr As DataRow In dtPago.Select
                If Length(data.IDBancoPropio) > 0 Then Dr("IDBancoPropio") = data.IDBancoPropio
                Dr("Impreso") = data.Impreso
            Next
            BusinessHelper.UpdateTable(dtPago)
        End If
    End Sub

    <Serializable()> _
    Public Class DataAjustarPagos
        Public PagosModificados As DataTable
        Public PagosEliminados As DataTable
    End Class
    <Task()> Public Shared Sub AjustarPagos(ByVal data As DataAjustarPagos, ByVal services As ServiceProvider)
        Dim p As New Pago
        If Not data.PagosModificados Is Nothing AndAlso data.PagosModificados.Rows.Count > 0 Then
            Dim IDPagos(-1) As Object
            For Each Dr As DataRow In data.PagosModificados.Select
                ReDim Preserve IDPagos(IDPagos.Length)
                IDPagos(IDPagos.Length - 1) = Dr("IDPago")
            Next
            Dim dtPagos As DataTable = p.Filter(New InListFilterItem("IDPago", IDPagos, FilterType.Numeric))
            If Not dtPagos Is Nothing Then
                If dtPagos.Rows.Count > 0 Then
                    For Each drPago As DataRow In dtPagos.Rows
                        Dim DrUp() As DataRow = data.PagosModificados.Select("IDPago=" & drPago("IDPago"))
                        If Not DrUp Is Nothing AndAlso DrUp.Length > 0 Then
                            Dim ValAyB As New ValoresAyB(CDbl(Nz(DrUp(0)("ImporteNew"), 0)), drPago("IDMoneda"), drPago("CambioA"), drPago("CambioB"))
                            Dim fImp As fImporte = ProcessServer.ExecuteTask(Of ValoresAyB, fImporte)(AddressOf NegocioGeneral.MantenimientoValoresImporteAyB, ValAyB, services)
                            drPago("ImpVencimiento") = fImp.Importe
                            drPago("ImpVencimientoA") = fImp.ImporteA
                            drPago("ImpVencimientoB") = fImp.ImporteB
                        End If
                    Next
                    BusinessHelper.UpdateTable(dtPagos)
                End If
            End If
            If Not data.PagosEliminados Is Nothing AndAlso data.PagosEliminados.Rows.Count > 0 Then p.Delete(data.PagosEliminados)
        End If
    End Sub


#Region "Informes"
    Public Shared cn_DEFAULT_LANGUAGE As String = "ES"
    <Task()> Public Shared Function CrearDTCartasProveedores(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim dtCarta As New DataTable

        '//Crear los campos
        With dtCarta.Columns
            .Add("IdProveedor", GetType(String))
            .Add("Titulo", GetType(String))
            .Add("IdBancoPropio", GetType(String))
            .Add("DescBancoPropio", GetType(String))
            .Add("ImporteTotalA", GetType(Double))
            .Add("NCheque", GetType(String))
            .Add("Direccion", GetType(String))
            .Add("CodPostal", GetType(String))
            .Add("Poblacion", GetType(String))
            .Add("Provincia", GetType(String))
            .Add("DescEmpresa", GetType(String))
            .Add("DirEmpresa", GetType(String))
            .Add("PoblEmpresa", GetType(String))
            .Add("ProvEmpresa", GetType(String))
            .Add("CIFEmpresa", GetType(String))
            .Add("CPEmpresa", GetType(String))
            .Add("TelfEmpresa", GetType(String))
            .Add("FaxEmpresa", GetType(String))
            .Add("EMailEmpresa", GetType(String))
            .Add("IDBanco", GetType(String))
            .Add("Sucursal", GetType(String))
            .Add("DigitoControl", GetType(String))
            .Add("NCuenta", GetType(String))
            .Add("IDIdioma", GetType(String))
        End With

        Return dtCarta
    End Function

    <Task()> Public Shared Function CartasProveedores(ByVal IDPagos() As Object, ByVal services As ServiceProvider) As DataTable
        '//Creamos la estructura del DataTable.
        Dim dtCartaProveedor As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CrearDTCartasProveedores, Nothing, services)

        '//Recogemos la información de la Empresa
        Dim EmpInfo As DatosEmpresaInfo = ProcessServer.ExecuteTask(Of Object, DatosEmpresaInfo)(AddressOf DatosEmpresa.ObtenerDatosEmpresa, Nothing, services)
        If EmpInfo Is Nothing OrElse Length(EmpInfo.ID) > 0 Then Exit Function

        Dim objNegBancoPropio As New BancoPropio
        Dim objNegProveedor As New Proveedor
        Dim dtBancoPropio As DataTable
        Dim dtProveedor As DataTable
        Dim objFilter As New Filter
        Dim strIDProvAnt, strIDBPAnt As String

        '//Cargamos los campos
        If IDPagos Is Nothing OrElse IDPagos.Length = 0 Then Exit Function
        Dim dtPagos As DataTable = New Pago().Filter(New InListFilterItem("IDPago", IDPagos, FilterType.Numeric))
        For Each drRowPago As DataRow In dtPagos.Select(Nothing, "IDProveedor,IDBancoPropio")
            Dim dblImpTotal As Double : Dim drRowCarta As DataRow
            If strIDProvAnt <> drRowPago("IDProveedor") OrElse strIDBPAnt <> drRowPago("IDBancoPropio") Then
                drRowCarta = dtCartaProveedor.NewRow
                dblImpTotal = 0
            End If
            If Length(EmpInfo.DescEmpresa) > 0 Then drRowCarta("DescEmpresa") = EmpInfo.DescEmpresa
            If Length(EmpInfo.Direccion) > 0 Then drRowCarta("DirEmpresa") = EmpInfo.Direccion
            If Length(EmpInfo.Poblacion) > 0 Then drRowCarta("PoblEmpresa") = EmpInfo.Poblacion
            If Length(EmpInfo.Provincia) > 0 Then drRowCarta("ProvEmpresa") = EmpInfo.Provincia
            If Length(EmpInfo.CodPostal) > 0 Then drRowCarta("CPEmpresa") = EmpInfo.CodPostal
            If Length(EmpInfo.CIF) > 0 Then drRowCarta("CIFEmpresa") = EmpInfo.CIF
            If Length(EmpInfo.Telefono) > 0 Then drRowCarta("TelfEmpresa") = EmpInfo.Telefono
            If Length(EmpInfo.Fax) > 0 Then drRowCarta("FaxEmpresa") = EmpInfo.Fax
            If Length(EmpInfo.Email) > 0 Then drRowCarta("EMailEmpresa") = EmpInfo.Email

            If Length(drRowPago("Titulo")) > 0 Then drRowCarta("Titulo") = drRowPago("Titulo") & String.Empty

            '//Recogemos la información del B.Propio
            drRowCarta("IDBancoPropio") = drRowPago("IDBancoPropio")
            Dim BancosPropios As EntityInfoCache(Of BancoPropioInfo) = services.GetService(Of EntityInfoCache(Of BancoPropioInfo))()
            Dim BPInfo As BancoPropioInfo = BancosPropios.GetEntity(drRowCarta("IDBancoPropio"))
            If Not BPInfo Is Nothing AndAlso Length(BPInfo.IDBancoPropio) > 0 Then
                If Length(BPInfo.DescBancoPropio) > 0 Then drRowCarta("DescBancoPropio") = BPInfo.DescBancoPropio
                If Length(BPInfo.IDBanco) > 0 Then drRowCarta("IDBanco") = BPInfo.IDBanco
                If Length(BPInfo.Sucursal) > 0 Then drRowCarta("Sucursal") = BPInfo.Sucursal
                If Length(BPInfo.DigitoControl) > 0 Then drRowCarta("DigitoControl") = BPInfo.DigitoControl
                If Length(BPInfo.NCuenta) > 0 Then drRowCarta("NCuenta") = BPInfo.NCuenta
            End If

            '//Recogemos la información del Proveedor
            drRowCarta("IDProveedor") = drRowPago("IDProveedor")
            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
            Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(drRowCarta("IDProveedor"))
            If Not ProvInfo Is Nothing AndAlso Length(ProvInfo.IDProveedor) > 0 Then
                If Length(ProvInfo.Direccion) > 0 Then drRowCarta("Direccion") = ProvInfo.Direccion
                If Length(ProvInfo.CodPostal) > 0 Then drRowCarta("CodPostal") = ProvInfo.CodPostal
                If Length(ProvInfo.Poblacion) > 0 Then drRowCarta("Poblacion") = ProvInfo.Poblacion
                If Length(ProvInfo.Provincia) > 0 Then drRowCarta("Provincia") = ProvInfo.Provincia
                If Length(ProvInfo.IDIdioma) > 0 Then
                    drRowCarta("IDIdioma") = ProvInfo.IDIdioma
                Else
                    drRowCarta("IDIdioma") = cn_DEFAULT_LANGUAGE
                End If
            End If

            '//Calculamos el Importe Total de los Pagos de un mismo Proveedor para un mismo Banco.
            objFilter.Clear()
            objFilter.Add(New StringFilterItem("IDProveedor", drRowPago("IDProveedor")))
            objFilter.Add(New StringFilterItem("IDBancoPropio", drRowPago("IDBancoPropio")))
            Dim dvPagosPBP As DataView = New DataView(dtPagos, objFilter.Compose(New AdoFilterComposer), Nothing, DataViewRowState.CurrentRows)
            If Not IsNothing(dvPagosPBP) AndAlso dvPagosPBP.Count > 0 Then
                dblImpTotal = dblImpTotal + Nz(drRowPago("ImpVencimientoA"), 0)
            End If
            drRowCarta("ImporteTotalA") = dblImpTotal

            If strIDProvAnt <> drRowPago("IDProveedor") OrElse strIDBPAnt <> drRowPago("IDBancoPropio") Then
                dtCartaProveedor.Rows.Add(drRowCarta)
                strIDProvAnt = drRowPago("IDProveedor")
                strIDBPAnt = drRowPago("IDBancoPropio")
            End If
        Next drRowPago

        Return dtCartaProveedor
    End Function

    <Task()> Public Shared Function EsAgrupado(ByVal IDPago As Integer, ByVal services As ServiceProvider) As Boolean
        Dim dt As DataTable = New Pago().Filter(New NumberFilterItem("IDPagoAgrupado", IDPago), , "TOP 1 IDPagoAgrupado")
        Return (Not dt Is Nothing AndAlso dt.Rows.Count > 0)
    End Function

    <Task()> Public Shared Function EsDesagrupable(ByVal Situacion As enumPagoSituacion, ByVal services As ServiceProvider) As Boolean
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDEstado", FilterOperator.Equal, Situacion))
        Dim dt As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf EstadoPago.EstadosPagosAgrupables, Nothing, services)
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            Dim WhereEstado As String = f.Compose(New AdoFilterComposer)
            Dim dr() As DataRow = dt.Select(WhereEstado)
            If dr.Length > 0 Then
                Return True
            Else : Return False
            End If
        End If
    End Function

    <Task()> Public Shared Function NumeracionPagares(ByVal IdPagos() As Object, ByVal services As ServiceProvider) As Boolean
        Dim FilPagos As New Filter
        FilPagos.Add(New InListFilterItem("IDPago", IdPagos, FilterType.Numeric))
        Dim dtPago As DataTable = New Pago().Filter(FilPagos, "FechaVencimiento ASC")
        If Not IsNothing(dtPago) AndAlso dtPago.Rows.Count > 0 Then
            Dim AppParams As ParametroTesoreria = services.GetService(Of ParametroTesoreria)()
            Dim strIDContador As String = AppParams.ContadorPagare
            If Length(strIDContador) > 0 Then
                Dim c As New Contador
                For Each drPago As DataRow In dtPago.Select
                    If Length(drPago("NPagare")) = 0 Then
                        If Length(drPago("IDFormaPago")) > 0 Then
                            Dim FormasPago As EntityInfoCache(Of FormaPagoInfo) = services.GetService(Of EntityInfoCache(Of FormaPagoInfo))()
                            Dim FPInfo As FormaPagoInfo = FormasPago.GetEntity(drPago("IDFormaPago"))
                            If Not FPInfo Is Nothing AndAlso Length(FPInfo.IDFormaPago) > 0 Then
                                If FPInfo.CobroImprimible Then
                                    drPago("NPagare") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, strIDContador, services)
                                End If
                            End If
                        End If
                    End If
                Next
                BusinessHelper.UpdateTable(dtPago)
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    <Serializable()> _
    Public Class DataImpresionPagos
        Public IdPagos() As Object
        Public IDBancoPropio As String
        Public Pagare As Boolean
        Public NuevaSituacion As enumPagoSituacion

        Public Sub New(ByVal IdPagos() As Object, ByVal IDBancoPropio As String, Optional ByVal Pagare As Boolean = False, Optional ByVal NuevaSituacion As enumPagoSituacion = -1)
            Me.IdPagos = IdPagos
            Me.IDBancoPropio = IDBancoPropio
            Me.NuevaSituacion = NuevaSituacion
            Me.Pagare = Pagare
        End Sub
    End Class

    <Task()> Public Shared Function ImpresionPagos(ByVal data As DataImpresionPagos, ByVal services As ServiceProvider) As Boolean
        Dim FilPagos As New Filter
        FilPagos.Add(New InListFilterItem("IDPago", data.IdPagos, FilterType.Numeric))
        Dim dtPago As DataTable = New BE.DataEngine().Filter("vNegImpresionPagosPagare", FilPagos, , "FechaVencimiento ASC")
        dtPago.TableName = "Pago"
        If Not IsNothing(dtPago) AndAlso dtPago.Rows.Count > 0 Then
            Dim AppParams As ParametroTesoreria = services.GetService(Of ParametroTesoreria)()
            Dim strIDContador As String = AppParams.ContadorPagare
            If Length(strIDContador) > 0 Then
                Dim c As New Contador
                Dim IDProveedorAnt As String
                Dim IDFormaPagoAnt As String
                Dim NumeroPagare As String
                For Each drPago As DataRow In dtPago.Select(Nothing, "IDProveedor, IDFormaPago")
                    If data.Pagare Then
                        If IDProveedorAnt <> drPago("IDProveedor") OrElse IDFormaPagoAnt <> drPago("IDFormaPago") & String.Empty Then
                            IDProveedorAnt = drPago("IDProveedor")
                            IDFormaPagoAnt = drPago("IDFormaPago") & String.Empty

                            If Nz(drPago("CobroImprimible"), False) AndAlso Length(drPago("NPagare")) = 0 Then
                                NumeroPagare = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, strIDContador, services)
                            Else
                                NumeroPagare = String.Empty
                            End If
                        End If

                        If Length(drPago("NPagare")) = 0 Then
                            drPago("NPagare") = NumeroPagare
                        End If
                    End If

                    If Length(data.IDBancoPropio) > 0 Then drPago("IDBancoPropio") = data.IDBancoPropio
                    drPago("Impreso") = True
                    If data.NuevaSituacion <> -1 Then drPago("Situacion") = data.NuevaSituacion
                Next

                BusinessHelper.UpdateTable(dtPago)
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    <Task()> Public Shared Function CrearDTPagare(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim dtPagare As New DataTable
        '//Crear los campos
        With dtPagare.Columns
            .Add("IDPago", GetType(Integer))
            .Add("Titulo", GetType(String))
            .Add("ImpVencimientoA", GetType(Double))
            .Add("FechaVencimiento", GetType(Date))
            .Add("DescMoneda", GetType(String))
            .Add("NDecimalesImp", GetType(Integer))
            .Add("Abreviatura", GetType(String))
            .Add("Poblacion", GetType(String))
        End With

        Return dtPagare
    End Function
    <Task()> Public Shared Function DatosPagare(ByVal IDPagos() As Object, ByVal services As ServiceProvider) As DataTable
        If IDPagos Is Nothing OrElse IDPagos.Length = 0 Then Exit Function
        Dim dtPagare As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CrearDTPagare, Nothing, services)
        Dim dtPagos As DataTable = New Pago().Filter(New InListFilterItem("IDPago", IDPagos, FilterType.Numeric))
        If Not dtPagos Is Nothing AndAlso dtPagos.Rows.Count > 0 Then
            Dim strPoblacion As String = ProcessServer.ExecuteTask(Of Object, DatosEmpresaInfo)(AddressOf DatosEmpresa.ObtenerDatosEmpresa, Nothing, services).Poblacion
            For Each dr As DataRow In dtPagos.Rows
                Dim drPagare As DataRow = dtPagare.NewRow
                'DATOS DEL PAGO
                drPagare("IDPago") = dr("IDPago")
                drPagare("Titulo") = dr("Titulo")
                drPagare("ImpVencimientoA") = dr("ImpVencimientoA")
                drPagare("FechaVencimiento") = dr("FechaVencimiento")
                'DATOS DE LA MONEDA
                Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                Dim MonInfo As MonedaInfo = Monedas.GetMoneda(dr("IDMoneda"))
                drPagare("DescMoneda") = MonInfo.Texto
                drPagare("Abreviatura") = MonInfo.Abreviatura
                drPagare("NDecimalesImp") = MonInfo.NDecimalesImporte
                'OTROS DATOS
                drPagare("Poblacion") = strPoblacion
                dtPagare.Rows.Add(drPagare)
            Next
        End If
        Return dtPagare
    End Function
#End Region

#Region " Declaración IVA Caja "

    <Task()> Public Shared Sub ActualizarFechaParaDeclaracionFactura(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        If data.ContainsKey("IDProveedor") AndAlso Length(data("IDProveedor")) > 0 Then
            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
            Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data("IDProveedor"))
            Dim AppParamsGral As ParametroGeneral = services.GetService(Of ParametroGeneral)()

            If ProvInfo.IVACaja Then
                If data.ContainsKey("IDFactura") AndAlso Length(data("IDFactura")) > 0 Then
                    Dim fFraNoDeclarada As New Filter
                    fFraNoDeclarada.Add(New NumberFilterItem("IDFactura", data("IDFactura")))
                    fFraNoDeclarada.Add(New IsNullFilterItem("NDeclaracionIVA", True))
                    fFraNoDeclarada.Add(New BooleanFilterItem("FechaDeclaracionManual", False))
                    Dim dtFCC As DataTable = New FacturaCompraCabecera().Filter(fFraNoDeclarada)
                    If dtFCC.Rows.Count > 0 Then
                        dtFCC.Rows(0)("FechaParaDeclaracion") = Nz(data("FechaPago"), New Date(Year(dtFCC.Rows(0)("FechaFactura")) + 1, 12, 31)) 'NegocioGeneral.cnMAX_DATE)
                    End If
                    AdminData.SetData(dtFCC)
                ElseIf data.ContainsKey("IDFactura") AndAlso Length(data("IDFactura")) = 0 AndAlso Length(data("NFactura")) > 0 AndAlso data("NFactura") = AppParamsGral.NFacturaPagoAgrupado Then
                    Dim dtFrasPagoAgrupado As DataTable = New Pago().Filter(New NumberFilterItem("IDPagoAgrupado", data("IDPago")), Nothing, "IDFactura")
                    If dtFrasPagoAgrupado.Rows.Count > 0 Then
                        Dim IDFacturas() As Object = (From c In dtFrasPagoAgrupado Where Not c.IsNull("IDFactura") Select c("IDFactura") Distinct).ToArray
                        If Not IDFacturas Is Nothing AndAlso IDFacturas.Count > 0 Then
                            Dim fFraNoDeclarada As New Filter
                            fFraNoDeclarada.Add(New InListFilterItem("IDFactura", IDFacturas, FilterType.Numeric))
                            fFraNoDeclarada.Add(New IsNullFilterItem("NDeclaracionIVA", True))
                            fFraNoDeclarada.Add(New BooleanFilterItem("FechaDeclaracionManual", False))
                            Dim dtFCC As DataTable = New FacturaCompraCabecera().Filter(fFraNoDeclarada)
                            If dtFCC.Rows.Count > 0 Then
                                For Each drFra As DataRow In dtFCC.Rows
                                    drFra("FechaParaDeclaracion") = Nz(data("FechaPago"), New Date(Year(drFra("FechaFactura")) + 1, 12, 31)) 'NegocioGeneral.cnMAX_DATE)
                                Next
                            End If
                            AdminData.SetData(dtFCC)
                        End If
                    End If
                End If
            End If
        End If

    End Sub



#End Region

End Class

