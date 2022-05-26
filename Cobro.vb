Imports System.Collections.Generic 

Public Class Cobro
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbCobro"

    Public Enum enumResultadoCambioCobros
        Ok = 0
        CobradoNoContabilizado = 1
        CobradoNoContabilizadoRemesable = 2
    End Enum

#Region " Delete "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.BeginTransaction)
        deleteProcess.AddTask(Of DataRow)(AddressOf ValidarDelGastosAsociados)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizaFechaVencimientoCabecera)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarVencimientosPromotoras)
        deleteProcess.AddTask(Of DataRow)(AddressOf EliminarEntregasACuenta)
        deleteProcess.AddTask(Of DataRow)(AddressOf CambiarEstadoPago)
    End Sub

    <Task()> Public Shared Sub CambiarEstadoPago(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Nz(data("IDPago"), -1) <> -1 Then
            '//si el pago está en estado GeneradoCobro, lo volvemos a la situación anterior
            Dim p As New Pago
            Dim f As New Filter
            f.Add(New NumberFilterItem("IDPago", data("IDPago")))
            f.Add(New NumberFilterItem("Situacion", enumPagoSituacion.GeneradoCobro))
            Dim dtPago As DataTable = p.SelOnPrimaryKey(data("IDPago"))
            If dtPago.Rows.Count > 0 Then
                dtPago.Rows(0)("Situacion") = enumPagoSituacion.NoPagado
            End If
            p.Update(dtPago)
        End If
    End Sub
    <Task()> Public Shared Sub ValidarDelGastosAsociados(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dtGastosRemesa As DataTable = New CobroFacturaCompra().Filter(New NumberFilterItem("IDCobro", data("IDCobro")))
        If dtGastosRemesa.Rows.Count > 0 Then
            ApplicationService.GenerateError("No se puede eliminar el Cobro, tiene gastos asociados.")
        End If
    End Sub

    <Task()> Public Shared Sub ActualizaFechaVencimientoCabecera(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDFactura")) > 0 Then
            Dim fvc As New FacturaVentaCabecera
            Dim dtFVC As DataTable = fvc.SelOnPrimaryKey(data("IDFactura"))
            If Not dtFVC Is Nothing AndAlso dtFVC.Rows.Count > 0 Then
                If dtFVC.Rows(0)("VencimientosManuales") Then
                    Dim dtCobro As DataTable = New Cobro().Filter(New NumberFilterItem("IDFactura", data("IDFactura")), "FechaVencimiento")
                    If Not dtCobro Is Nothing AndAlso dtCobro.Rows.Count > 0 Then
                        dtFVC.Rows(0)("FechaVencimiento") = dtCobro.Rows(0)("FechaVencimiento")
                    Else
                        dtFVC.Rows(0)("FechaVencimiento") = System.DBNull.Value
                    End If
                    BusinessHelper.UpdateTable(dtFVC)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarVencimientosPromotoras(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If AppParams.GestionPromotoras Then
            ProcessServer.ExecuteTask(Of Integer)(AddressOf EliminarCobroVencimientosLocalPromociones, data("IDCobro"), services)
        End If
    End Sub

    <Task()> Public Shared Sub EliminarEntregasACuenta(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim blnDelete As Boolean
        '//Si PROVIENE de una ENTREGA y está VINCULADO a una FACTURA, modificamos los campos necesarios de la Entrega.
        If Length(data("IDEntrega")) > 0 AndAlso Length(data("IDFactura")) > 0 Then
            '//Eliminamos los vínculos del cobro de la Factura y la Entrega de la que proviene.
            Dim objNegEC As New EntregasACuenta
            Dim StDatos As New EntregasACuenta.DatosElimRestricEntFn
            StDatos.IDEntrega = data("IDEntrega")
            StDatos.IDFactura = data("IDFactura")
            StDatos.Circuito = Circuito.Ventas
            ProcessServer.ExecuteTask(Of EntregasACuenta.DatosElimRestricEntFn, Boolean)(AddressOf EntregasACuenta.EliminarRestriccionesDeleteEntregaCuentaFn, StDatos, services)
            '//NO BORRAMOS EL COBRO, lo desvinculamos de la Factura. 
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf DesvincularCobroDeFactura, data, services)
            blnDelete = False
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoEliminado, data, services)
        ElseIf Length(data("IDEntrega")) > 0 AndAlso Length(data("IDFactura")) = 0 Then
            '//Si PROVIENE de una ENTREGA y NO está VINCULADO a una FACTURA.
            Dim datosEntrega As New EntregasACuenta.DatosElimRestricEntCobro
            datosEntrega.IDEntrega = data("IDEntrega")
            datosEntrega.IDCobroPago = data("IDCobro")
            ProcessServer.ExecuteTask(Of EntregasACuenta.DatosElimRestricEntCobro)(AddressOf EntregasACuenta.EliminarRestriccionesDeleteEntregaCuentaCobroPago, datosEntrega, services)
        End If
        If blnDelete Then
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf Business.General.Comunes.DeleteEntityRow, data, services)
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoEliminado, data, services)
        End If
    End Sub

    <Task()> Public Shared Sub DesvincularCobroDeFactura(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("IDFactura") = System.DBNull.Value
        Dim dtModif As DataTable = data.Table.Clone
        dtModif.ImportRow(data)
        Dim c As New Cobro
        c.Update(dtModif)
    End Sub

#End Region

#Region " Validate "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFormaPagoObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosCobroNormal)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFechaVencimientoObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificador)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCambioBancoPropio)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCambioFormaPago)
    End Sub

    ''' Validación para Cobros que no vienen desde Pagos
    <Task()> Public Shared Sub ValidarDatosCobroNormal(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarTipoCobroPredeterminadoFV, data, services)
        Dim AppParams As ParametroTesoreria = services.GetService(Of ParametroTesoreria)()
        If data("IDTipoCobro") <> AppParams.TipoCobroDesdePago Then
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.ValidarClienteObligatorio, data, services)
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.ValidarCContableObligatoria, data, services)
        End If
    End Sub

    <Task()> Public Shared Sub ValidarCambioBancoPropio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            If Length(data("IDBancoPropio", DataRowVersion.Original)) > 0 AndAlso _
               data("IDBancoPropio") & String.Empty <> data("IDBancoPropio", DataRowVersion.Original) Then
                If Length(data("IDCheque")) > 0 Then
                    ApplicationService.GenerateError("No se puede modificar el Banco Propio. El Cobro está asociado a un Cheque.")
                End If
                If Length(data("IDTarjeta")) > 0 Then
                    ApplicationService.GenerateError("No se puede modificar el Banco Propio. El Cobro está asociado a una Tarjeta.")
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarCambioFormaPago(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            If Length(data("IDFormaPago", DataRowVersion.Original)) > 0 AndAlso _
               data("IDFormaPago") & String.Empty <> data("IDFormaPago", DataRowVersion.Original) Then
                If Length(data("IDCheque")) > 0 Then
                    ApplicationService.GenerateError("No se puede modificar la Forma de Pago. El Cobro está asociado a un Cheque.")
                End If
                If Length(data("IDTarjeta")) > 0 Then
                    ApplicationService.GenerateError("No se puede modificar la Forma de Pago. El Cobro está asociado a una Tarjeta.")
                End If
            End If
        End If
    End Sub

#End Region

#Region " Update "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.BeginTransaction)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizaVencimientosManuales)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarTipoCobroPredeterminadoFV)
        updateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.AsignarMonedaPredeterminada)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificador)
        updateProcess.AddTask(Of DataRow)(AddressOf CambioSituacionPagoAsociado)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarImporteRemesaAnticipo)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarValoresAyB)
        updateProcess.AddTask(Of DataRow)(AddressOf EliminarMandatoSiFormaPagoNoRemesable)
        updateProcess.AddTask(Of DataRow)(AddressOf CambioFechaParaDeclaracion)
    End Sub

    <Task()> Public Shared Sub EliminarMandatoSiFormaPagoNoRemesable(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Not data.Table.Columns.Contains("IDMandato") Then Exit Sub

        If Length(data("IDFormaPago")) > 0 AndAlso _
           (data.RowState = DataRowState.Added OrElse _
           (data.RowState = DataRowState.Modified AndAlso data("IDFormaPago") <> data("IDFormaPago", DataRowVersion.Original) & String.Empty) OrElse Nz(data("IDMandato"), 0) <> Nz(data("IDMandato", DataRowVersion.Original), 0)) Then

            Dim FormasPago As EntityInfoCache(Of FormaPagoInfo) = services.GetService(Of EntityInfoCache(Of FormaPagoInfo))()
            Dim FPInfo As FormaPagoInfo = FormasPago.GetEntity(data("IDFormaPago"))
            If Not FPInfo.CobroRemesable Then
                data("IDMandato") = System.DBNull.Value
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarTipoCobroPredeterminadoFV(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IdTipoCobro")) = 0 Then
            Dim AppParams As ParametroTesoreria = services.GetService(Of ParametroTesoreria)()
            data("IdTipoCobro") = AppParams.TipoCobroFV
        End If
    End Sub

    <Task()> Public Shared Sub AsignarIdentificador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IdCobro")) = 0 Then data("IdCobro") = AdminData.GetAutoNumeric
    End Sub

    <Task()> Public Shared Sub CambioSituacionPagoAsociado(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified AndAlso Length(data("IDPago")) > 0 Then
            If data("Situacion") = enumCobroSituacion.Cobrado AndAlso data("Situacion", DataRowVersion.Original) <> enumCobroSituacion.Cobrado Then
                Dim dtPago As DataTable = New Pago().SelOnPrimaryKey(data("IDPago"))
                Dim datosCambio As New Pago.DataCambioSituacionManual
                datosCambio.Pagos = dtPago
                datosCambio.NuevaSituacion = enumPagoSituacion.Pagado
                ProcessServer.ExecuteTask(Of Pago.DataCambioSituacionManual)(AddressOf Pago.CambioSituacionManual, datosCambio, services)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarImporteRemesaAnticipo(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            data("ImporteRemesaAnticipo") = Nz(data("ImpVencimiento"), 0)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarValoresAyB(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(data), data("IDMoneda"), data("CambioA"), data("CambioB"))
        ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
    End Sub

    <Task()> Public Shared Sub CambioFechaParaDeclaracion(ByVal data As DataRow, ByVal services As ServiceProvider)
        ' If Nz(data("FechaCobro"), cnMinDate) <> cnMinDate Then
        If data.RowState = DataRowState.Added OrElse (data.RowState = DataRowState.Modified AndAlso Nz(data("FechaCobro"), cnMinDate) <> Nz(data("FechaCobro", DataRowVersion.Original), cnMinDate)) Then
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf Cobro.ActualizarFechaParaDeclaracionFactura, New DataRowPropertyAccessor(data), services)
        End If
        ' End If
    End Sub
#End Region

#Region " RegisterAddNewTasks "

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarIdentificador, data, services)
    End Sub

#End Region

#Region " Business Rules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        'Return MyBase.GetBusinessRules()
        Dim oBRL As New BusinessRules
        oBRL.Add("IDTipoCobro", AddressOf CambioTipoCobro)
        oBRL.Add("IDCliente", AddressOf CambioCliente)
        oBRL.Add("IDClienteBanco", AddressOf CambioBancoCliente)
        oBRL.Add("IDFormaPago", AddressOf CambioFormaPago)
        oBRL.Add("IDBancoPropio", AddressOf CambioBancoPropio)
        oBRL.Add("Situacion", AddressOf CambioSituacionCobro)
        oBRL.Add("ImpVencimiento", AddressOf NegocioGeneral.CambioImporteVencimiento)
        oBRL.Add("ImpVencimientoA", AddressOf NegocioGeneral.CambioImporteVencimiento)
        oBRL.Add("ImpVencimientoB", AddressOf NegocioGeneral.CambioImporteVencimiento)
        oBRL.Add("ImporteRemesaAnticipo", AddressOf CambioImporteRemesaAnticipo)
        oBRL.Add("ImporteRemesaAnticipoA", AddressOf CambioImporteRemesaAnticipo)
        oBRL.Add("ImporteRemesaAnticipoB", AddressOf CambioImporteRemesaAnticipo)
        oBRL.Add("ARepercutir", AddressOf NegocioGeneral.CambioImporteRepercutir)
        oBRL.Add("IDMoneda", AddressOf NegocioGeneral.CambioMonedaFechaVto)
        oBRL.Add("FechaVencimiento", AddressOf NegocioGeneral.CambioMonedaFechaVto)
        oBRL.Add("CambioA", AddressOf NegocioGeneral.CambioEnCambiosMoneda)
        oBRL.Add("CambioB", AddressOf NegocioGeneral.CambioEnCambiosMoneda)
        oBRL.Add("NFactura", AddressOf CambioNFactura)
        oBRL.Add("IDAgrupacion", AddressOf CambioAgrupacion)
        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioImporteRemesaAnticipo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDMoneda")) > 0 Then
            Select Case data.ColumnName
                Case "ImporteRemesaAnticipoA"
                    If Nz(data.Current("CambioA"), 0) <> 0 Then
                        data.Current("ImporteRemesaAnticipo") = Nz(data.Current("ImporteRemesaAnticipoA"), 0) / data.Current("CambioA")
                    Else
                        data.Current("ImporteRemesaAnticipo") = 0
                    End If
                Case "ImporteRemesaAnticipoB"
                    If Nz(data.Current("CambioB"), 0) <> 0 Then
                        data.Current("ImporteRemesaAnticipo") = Nz(data.Current("ImporteRemesaAnticipoB"), 0) / data.Current("CambioB")
                    Else
                        data.Current("ImporteRemesaAnticipo") = 0
                    End If
            End Select

            Dim ValAyB As New ValoresAyB(data.Current, data.Current("IDMoneda"), Nz(data.Current("CambioA"), 0), Nz(data.Current("CambioB"), 0))
            ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
        End If
    End Sub

    <Task()> Public Shared Sub CambioTipoCobro(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDTipoCobro")) > 0 Then
            Dim dt As DataTable = New TipoCobro().SelOnPrimaryKey(data.Current("IDTipoCobro"))
            If dt.Rows.Count = 0 Then
                ApplicationService.GenerateError("No existe el Tipo de Cobro indicado.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioCliente(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDCliente")) > 0 Then
            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data.Current("IDCliente"))
            If Not ClteInfo Is Nothing AndAlso Length(ClteInfo.IDCliente) > 0 Then
                Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
                If AppParamsConta.Contabilidad Then data.Current("CContable") = ClteInfo.CCCliente
                data.Current("IDFormaPago") = ClteInfo.FormaPago
                data.Current("IDMoneda") = ClteInfo.Moneda
                data.Current("Titulo") = ClteInfo.RazonSocial
                data.Current("IDBancoPropio") = ClteInfo.IDBancoPropio

                Dim stDatosDirec As New ClienteDireccion.DataDirecEnvio(data.Current("IDCliente"), enumcdTipoDireccion.cdDireccionGiro)
                Dim dtClienteDireccion As DataTable = ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecEnvio, DataTable)(AddressOf ClienteDireccion.ObtenerDireccionEnvio, stDatosDirec, services)
                If Not IsNothing(dtClienteDireccion) AndAlso dtClienteDireccion.Rows.Count > 0 Then
                    data.Current("IDDireccion") = dtClienteDireccion.Rows(0)("IDDireccion")
                End If

                Dim IDClienteBanco As Integer = ProcessServer.ExecuteTask(Of String, Integer)(AddressOf ClienteBanco.GetBancoPredeterminado, data.Current("IDCliente"), services)
                If IDClienteBanco <> 0 Then
                    data.Current("IDClienteBanco") = IDClienteBanco
                    ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf GetMandatoSEPAPredeterminado, data.Current, services)
                End If
            End If
        Else
            data.Current("CContable") = DBNull.Value
            data.Current("IDFormaPago") = DBNull.Value
            data.Current("IDMoneda") = DBNull.Value
            data.Current("Titulo") = DBNull.Value
            data.Current("IDBancoPropio") = DBNull.Value
            data.Current("IDDireccion") = DBNull.Value
            data.Current("IDClienteBanco") = DBNull.Value
            data.Current("IDMandato") = DBNull.Value
            data.Current("NMandato") = DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub CambioBancoCliente(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf GetMandatoSEPAPredeterminado, data.Current, services)
    End Sub

    <Task()> Public Shared Sub CambioFormaPago(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDFormaPago")) > 0 Then
            Dim FormasPago As EntityInfoCache(Of FormaPagoInfo) = services.GetService(Of EntityInfoCache(Of FormaPagoInfo))()
            Dim FPInfo As FormaPagoInfo = FormasPago.GetEntity(data.Current("IDFormaPago"))
            If Not FPInfo.CobroRemesable Then
                data.Current("IDMandato") = System.DBNull.Value
                If data.Current.ContainsKey("NMandato") Then data.Current("NMandato") = System.DBNull.Value
            Else
                Dim FormasPagoSEPA As List(Of String) = New Parametro().FormaPagoMandatoSEPA
                If Not FormasPagoSEPA Is Nothing AndAlso FormasPagoSEPA.Count > 0 Then
                    If FormasPagoSEPA.Contains(UCase(data.Current("IDFormaPago"))) Then
                        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf GetMandatoSEPAPredeterminado, data.Current, services)
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub GetMandatoSEPAPredeterminado(ByVal Current As IPropertyAccessor, ByVal services As ServiceProvider)
        If Current.ContainsKey("IDClienteBanco") AndAlso Nz(Current("IDClienteBanco")) <> 0 Then
            Dim fMdto As New Filter
            fMdto.Add("IDClienteBanco", FilterOperator.Equal, Current("IDClienteBanco"))
            fMdto.Add("Caducado", FilterOperator.Equal, False)
            fMdto.Add("Estado", FilterOperator.Equal, 1) 'BusinessEnum.MandatoEstado.Aceptado
            fMdto.Add("Predeterminado", FilterOperator.Equal, True)
            Dim ClsMandato As BusinessHelper = BusinessHelper.CreateBusinessObject("Mandato")
            Dim dtMandato As DataTable = ClsMandato.Filter(fMdto)
            If dtMandato.Rows.Count > 0 Then
                Current("IDMandato") = dtMandato.Rows(0)("IDMandato")
                If Current.ContainsKey("NMandato") Then Current("NMandato") = dtMandato.Rows(0)("NMandato")
            Else
                Current("IDMandato") = System.DBNull.Value
                If Current.ContainsKey("NMandato") Then Current("NMandato") = System.DBNull.Value
            End If
        Else
            Current("IDMandato") = System.DBNull.Value
            If Current.ContainsKey("NMandato") Then Current("NMandato") = System.DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub CambioBancoPropio(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDBancoPropio")) > 0 Then
            Dim dt As DataTable = New BancoPropio().SelOnPrimaryKey(data.Current("IDBancoPropio"))
            If dt.Rows.Count = 0 Then
                ApplicationService.GenerateError("No existe el Banco Propio indicado.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioAgrupacion(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDAgrupacion")) > 0 Then
            Dim dt As DataTable = New Agrupacion().SelOnPrimaryKey(data.Current("IDAgrupacion"))
            If dt.Rows.Count = 0 Then
                ApplicationService.GenerateError("No existe la Agrupación indicada.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioSituacionCobro(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("Situacion")) > 0 Then
            If Not IsNumeric(data.Current("Situacion")) Then ApplicationService.GenerateError("El campo Situación debe ser numérico.")
            Dim dt As DataTable = New EstadoCobro().SelOnPrimaryKey(data.Current("Situacion"))
            If dt.Rows.Count = 0 Then
                ApplicationService.GenerateError("No existe la Situación indicada.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioNFactura(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDCobro")) > 0 Then
            If ProcessServer.ExecuteTask(Of Integer, Boolean)(AddressOf EsAgrupado, data.Current("IDCobro"), services) Then
                data.Current("IDFactura") = DBNull.Value
            Else
                'Si el NFactura no existe, borrar el IDFactura
                Dim dtFact As DataTable = New FacturaVentaCabecera().Filter(New StringFilterItem("NFactura", data.Current("NFactura")))
                If dtFact.Rows.Count = 0 Then
                    data.Current("IDFactura") = DBNull.Value
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Function EsAgrupado(ByVal IDCobro As Integer, ByVal services As ServiceProvider) As Boolean
        Dim dt As DataTable = New Cobro().Filter(New NumberFilterItem("IDCobroAgrupado", IDCobro), , "TOP 1 IDCobroAgrupado")
        Return (Not dt Is Nothing AndAlso dt.Rows.Count > 0)
    End Function

#End Region

#Region " Desglosar Cobros "

    <Serializable()> _
    Public Class DataDesglosarCobros
        Public IDCobroDesglosar As Integer
        Public NuevosCobros As DataTable
    End Class

    <Task()> Public Shared Sub DesglosarCobros(ByVal data As DataDesglosarCobros, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
        Dim c As New Cobro
        Dim dtCobro As DataTable = c.SelOnPrimaryKey(data.IDCobroDesglosar)
        c.Delete(dtCobro)
        c.Update(data.NuevosCobros)
        ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.CommitTransaction, Nothing, services)
    End Sub

#End Region

#Region " Actualiza Vencimientos Manuales "

    <Task()> Public Shared Sub ActualizaVencimientosManuales(ByVal dr As DataRow, ByVal services As ServiceProvider)
        If Length(dr("IDFactura")) > 0 Then
            Dim fvc As New FacturaVentaCabecera
            Dim dtFVC As DataTable = fvc.SelOnPrimaryKey(dr("IDFactura"))
            If Not dtFVC Is Nothing AndAlso dtFVC.Rows.Count > 0 Then
                If dtFVC.Rows(0)("VencimientosManuales") Then
                    Dim dtCobro As DataTable = New Cobro().Filter(New NumberFilterItem("IDFactura", dr("IDFactura")), "FechaVencimiento")
                    If Not dtCobro Is Nothing AndAlso dtCobro.Rows.Count > 0 Then
                        If dtCobro.Rows(0)("FechaVencimiento") < dr("FechaVencimiento") Then
                            If dr.RowState = DataRowState.Modified Then
                                If Length(dr("FechaVencimiento", DataRowVersion.Original)) = 0 OrElse dr("FechaVencimiento", DataRowVersion.Original) = dtCobro.Rows(0)("FechaVencimiento") Then
                                    dtFVC.Rows(0)("FechaVencimiento") = dr("FechaVencimiento")
                                End If
                            Else
                                dtFVC.Rows(0)("FechaVencimiento") = dtCobro.Rows(0)("FechaVencimiento")
                            End If
                        Else
                            dtFVC.Rows(0)("FechaVencimiento") = dr("FechaVencimiento")
                        End If
                    Else
                        dtFVC.Rows(0)("FechaVencimiento") = dr("FechaVencimiento")
                    End If
                    BusinessHelper.UpdateTable(dtFVC)
                    If Length(dr("FechaVencimientoFactura")) = 0 Then dr("FechaVencimientoFactura") = dr("FechaVencimiento")

                    If dtFVC.Rows(0)("Estado") = enumfvcEstado.fvcNoContabilizado Then
                        dr("IDObra") = dtFVC.Rows(0)("IDObra")
                        Dim AppParams As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
                        If AppParams.Contabilidad Then
                            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
                            Dim ClteInfo As ClienteInfo = Clientes.GetEntity(dtFVC.Rows(0)("IDCliente"))
                            If Length(ClteInfo.IDCliente) > 0 Then dr("CContable") = ClteInfo.CCCliente
                        End If
                        Dim cd As New ClienteDireccion
                        Dim StDatosDirec As New ClienteDireccion.DataDirecDe(dtFVC.Rows(0)("IDDireccion"), enumcdTipoDireccion.cdDireccionGiro)
                        If ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecDe, Boolean)(AddressOf ClienteDireccion.EsDireccionDe, StDatosDirec, services) = True Then
                            dr("IDDireccion") = dtFVC.Rows(0)("IDDireccion")
                        Else
                            Dim StDatosDirecEnv As New ClienteDireccion.DataDirecEnvio(dtFVC.Rows(0)("IDCliente"), enumcdTipoDireccion.cdDireccionGiro)
                            Dim direc As DataTable = ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecEnvio, DataTable)(AddressOf ClienteDireccion.ObtenerDireccionEnvio, StDatosDirecEnv, services)
                            If Not IsNothing(direc) AndAlso direc.Rows.Count Then
                                dr("IDDireccion") = direc.Rows(0)("IDDireccion")
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End Sub

#End Region

#Region " Cobros remesables "

    <Task()> Public Shared Function ComprobarCobrosRemesables(ByVal IDCobros() As Integer, ByVal services As ServiceProvider) As Boolean
        If IDCobros Is Nothing OrElse IDCobros.Length = 0 Then ApplicationService.GenerateError("Debe indicar al menos un cobro.")
        Dim blnRemesables As Boolean = True
        Dim f As New Filter : Dim BEDataEngine As New BE.DataEngine
        For Each IDCobro As Integer In IDCobros
            f.Clear()
            f.Add(New NumberFilterItem("IDCobro", IDCobro))
            Dim dtCobrosRemesables As DataTable = BEDataEngine.Filter("NegCobrosRemesables", f)
            If dtCobrosRemesables.Rows.Count = 0 Then
                blnRemesables = False
                Exit For
            End If
        Next

        Return blnRemesables
    End Function

#End Region

#Region " Añadir/Quitar cobros a Remesas "

    <Serializable()> _
    Public Class DataAñadirCobrosARemesa
        Public IDRemesa As Integer
        Public CobrosAAñadir As DataTable
    End Class

    <Task()> Public Shared Sub AñadirCobrosARemesa(ByVal data As DataAñadirCobrosARemesa, ByVal services As ServiceProvider)
        If Not data.CobrosAAñadir Is Nothing AndAlso data.CobrosAAñadir.Rows.Count > 0 Then
            AdminData.BeginTx()
            Dim intTipoAsientoRemesa As String = New Parametro().TipoAsientoRemesa
            Dim blActAutomaticaSituacionRemCobro As Boolean = New Parametro().ActAutomaticaSituacionRemCobro
            Dim f As New Filter
            f.Add(New NumberFilterItem("IDRemesa", data.IDRemesa))
            Dim dtRemesa As DataTable = New Remesa().Filter(f)
            If Not dtRemesa Is Nothing AndAlso dtRemesa.Rows.Count > 0 Then
                If dtRemesa.Rows(0)("IDTipoNegociacion") <> enumTipoRemesa.RemesaAnticipo Then
                    '//Comprobar si la remesa ya está contabilizada, en ese caso no dejar Añadir
                    Dim fRemConta As New Filter
                    fRemConta.Add(New NumberFilterItem("IDRemesa", data.IDRemesa))
                    fRemConta.Add(New NumberFilterItem("Contabilizado", FilterOperator.NotEqual, enumContabilizado.NoContabilizado))
                    Dim dtCobros As DataTable = New Cobro().Filter(fRemConta)
                    If dtCobros.Rows.Count > 0 Then
                        ApplicationService.GenerateError("La Remesa donde se está intentando introducir el Cobro está contabilizada. No se pueden añadir los Cobros.")
                    End If
                End If
                For Each dr As DataRow In data.CobrosAAñadir.Rows
                    dr("IDBancoPropio") = dtRemesa.Rows(0)("IDBancoPropio")

                    If dtRemesa.Rows(0)("IDTipoNegociacion") <> enumTipoRemesa.RemesaAnticipo Then
                        dr("IDRemesa") = data.IDRemesa
                    Else
                        dr("IDRemesaAnticipo") = data.IDRemesa
                    End If

                    If dtRemesa.Rows(0)("IDTipoNegociacion") <> enumTipoRemesa.RemesaAnticipo AndAlso _
                       intTipoAsientoRemesa = enumTipoAsientoRemesa.Banco_a_Cliente And blActAutomaticaSituacionRemCobro = True Then
                        'Si el TipoAsientoRemesa=0 y al generar la remesa se pasan a 'cobrado' automáticamente,
                        ' al añadir cobros tienen que pasarse a 'cobrado'.
                        dr("Situacion") = enumCobroSituacion.Cobrado
                    Else
                        If dtRemesa.Rows(0)("IDTipoNegociacion") = enumTipoRemesa.RemesaAlCobro Then 'Al Cobro
                            dr("Situacion") = enumCobroSituacion.Negociado
                        ElseIf dtRemesa.Rows(0)("IDTipoNegociacion") = enumTipoRemesa.RemesaAlDescuento Then 'Al Descuento
                            dr("Situacion") = enumCobroSituacion.Descontado
                        ElseIf dtRemesa.Rows(0)("IDTipoNegociacion") = enumTipoRemesa.RemesaAnticipo Then 'Anticipo
                            If dr("Situacion") = enumCobroSituacion.NoNegociado Then dr("Situacion") = enumCobroSituacion.Anticipado
                            dr("EstadoAnticipo") = enumEstadoAnticipo.PdteAbono
                        End If
                    End If
                Next

                BusinessHelper.UpdateTable(data.CobrosAAñadir)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub RetirarCobrosDeRemesas(ByVal IDCobros() As Object, ByVal services As ServiceProvider)
        If Not IDCobros Is Nothing AndAlso IDCobros.Length > 0 Then
            AdminData.BeginTx()
            Dim f As New Filter
            f.Add(New InListFilterItem("IDCobro", IDCobros, FilterType.Numeric))
            Dim dtCobros As DataTable = New Cobro().Filter(f)
            If Not dtCobros Is Nothing AndAlso dtCobros.Rows.Count > 0 Then
                For Each dr As DataRow In dtCobros.Select("Contabilizado = " & enumContabilizado.NoContabilizado)
                    If Length(dr("IDRemesaAnticipo")) = 0 Then dr("IDBancoPropio") = DBNull.Value
                    dr("IDRemesa") = DBNull.Value
                    'dr("Situacion") = enumCobroSituacion.NoNegociado
                    If Length(dr("IDRemesaAnticipo")) = 0 Then
                        dr("Situacion") = enumCobroSituacion.NoNegociado
                    Else
                        If dr("EstadoAnticipo") = enumEstadoAnticipo.Cancelado Then
                            dr("Situacion") = enumCobroSituacion.NoNegociado
                        Else
                            dr("Situacion") = enumCobroSituacion.Anticipado
                        End If
                    End If

                Next
                BusinessHelper.UpdateTable(dtCobros)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub RetirarCobrosDeAnticipos(ByVal IDCobros() As Object, ByVal services As ServiceProvider)
        If Not IDCobros Is Nothing AndAlso IDCobros.Length > 0 Then
            AdminData.BeginTx()
            Dim f As New Filter
            f.Add(New InListFilterItem("IDCobro", IDCobros, FilterType.Numeric))
            f.Add(New NumberFilterItem("Contabilizado", enumContabilizado.NoContabilizado))
            Dim dtCobros As DataTable = New BE.DataEngine().Filter("vNegRetirarCobroRemesaAnticipo", f)
            dtCobros.TableName = "Cobro"
            If Not dtCobros Is Nothing AndAlso dtCobros.Rows.Count > 0 Then
                For Each dr As DataRow In dtCobros.Rows
                    If Nz(dr("ContabilizadoAnticipo"), -1) = enumContabilizado.NoContabilizado Then
                        dr("IDRemesaAnticipo") = DBNull.Value
                        If dr("Situacion") = enumCobroSituacion.Anticipado Then
                            If Length(dr("IDRemesa")) = 0 Then
                                dr("Situacion") = enumCobroSituacion.NoNegociado
                            Else
                                f.Clear()
                                f.Add(New NumberFilterItem("IDRemesa", dr("IDRemesa")))
                                Dim dtRemesa As DataTable = New Remesa().Filter(f)
                                If Not dtRemesa Is Nothing AndAlso dtRemesa.Rows.Count > 0 Then
                                    If dtRemesa.Rows(0)("IDTipoNegociacion") = enumTipoRemesa.RemesaAlCobro Then 'Al Cobro
                                        dr("Situacion") = enumCobroSituacion.Negociado
                                    Else 'If dtRemesa.Rows(0)("IDTipoNegociacion") = enumTipoRemesa.RemesaAlDescuento Then 'Al Descuento
                                        dr("Situacion") = enumCobroSituacion.Descontado
                                    End If
                                End If
                            End If
                        End If
                        dr("EstadoAnticipo") = DBNull.Value
                        dr("ReferenciaCancelacionAnticipo") = DBNull.Value
                        dr("FechaCancelacionAnticipo") = DBNull.Value
                        dr("FechaCancelacionAnticipoPrev") = DBNull.Value
                    End If
                Next
                BusinessHelper.UpdateTable(dtCobros)
            End If
        End If
    End Sub


#End Region

#Region " Liquidación de Remesas sin Contabilidad "

    <Serializable()> _
    Public Class DataLiquidacionRemesaSinConta
        Public IDCobros() As Object
        Public NuevaSituacion As enumCobroSituacion
    End Class

    <Task()> Public Shared Sub LiquidacionRemesaSinConta(ByVal data As DataLiquidacionRemesaSinConta, ByVal services As ServiceProvider)
        Dim dtCobros As DataTable = New BE.DataEngine().Filter("vNegCobroRemesa", New InListFilterItem("IDCobro", data.IDCobros, FilterType.Numeric))
        If Not dtCobros Is Nothing Then
            For Each drCobro As DataRow In dtCobros.Rows
                If drCobro("IDTipoNegociacion") = enumTipoRemesa.RemesaAnticipo Then
                    drCobro("EstadoAnticipo") = enumEstadoAnticipo.Cancelado
                Else
                    drCobro("Liquidado") = enumContabilizado.Contabilizado
                End If
                drCobro("Situacion") = data.NuevaSituacion
            Next
        End If
        dtCobros.TableName = GetType(Cobro).Name
        BusinessHelper.UpdateTable(dtCobros)
    End Sub

    <Task()> Public Shared Sub EliminarLiquidacionRemesaSinConta(ByVal data As DataLiquidacionRemesaSinConta, ByVal services As ServiceProvider)
        Dim dtCobros As DataTable = New BE.DataEngine().Filter("vNegCobroRemesa", New InListFilterItem("IDCobro", data.IDCobros, FilterType.Numeric))
        For Each dr As DataRow In dtCobros.Rows
            If dr("IDTipoNegociacion") = enumTipoRemesa.RemesaAlCobro Then
                dr("Situacion") = enumCobroSituacion.Negociado
                dr("Liquidado") = enumContabilizado.NoContabilizado
            ElseIf dr("IDTipoNegociacion") = enumTipoRemesa.RemesaAnticipo Then
                dr("Situacion") = enumCobroSituacion.Anticipado
                dr("EstadoAnticipo") = enumEstadoAnticipo.Abonado
            Else
                dr("Situacion") = enumCobroSituacion.Descontado
                dr("Liquidado") = enumContabilizado.NoContabilizado
            End If
        Next
        dtCobros.TableName = GetType(Cobro).Name
        BusinessHelper.UpdateTable(dtCobros)
    End Sub
#End Region

#Region " Actualizar Vtos Promociones Obras"

    <Task()> Public Shared Sub EliminarCobroVencimientosLocalPromociones(ByVal IDCobro As Integer, ByVal services As ServiceProvider)
        'Dim Local As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraPromoLocalVencimiento")
        Dim dtLocal As DataTable = BusinessHelper.CreateBusinessObject("ObraPromoLocalVencimiento").Filter(New NumberFilterItem("IDCobro", IDCobro))
        If Not dtLocal Is Nothing AndAlso dtLocal.Rows.Count > 0 Then
            For Each drLocal As DataRow In dtLocal.Rows
                drLocal("CobroGenerado") = False
                drLocal("IDCobro") = System.DBNull.Value
            Next

            BusinessHelper.UpdateTable(dtLocal)
        End If
    End Sub

    <Serializable()> _
    Public Class dataAsociarCobroVencimientosLocalPromociones
        Public IDLocalVencimiento As Integer
        Public IDCobro As Integer

        Public Sub New(ByVal IDLocalVencimiento As Integer, ByVal IDCobro As Integer)
            Me.IDLocalVencimiento = IDLocalVencimiento
            Me.IDCobro = IDCobro
        End Sub
    End Class
    <Task()> Public Shared Sub AsociarCobroVencimientosLocalPromociones(ByVal data As dataAsociarCobroVencimientosLocalPromociones, ByVal services As ServiceProvider)
        'Dim Local As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraPromoLocalVencimiento")
        Dim dtLocal As DataTable = BusinessHelper.CreateBusinessObject("ObraPromoLocalVencimiento").SelOnPrimaryKey(data.IDLocalVencimiento)
        If Not dtLocal Is Nothing AndAlso dtLocal.Rows.Count > 0 Then
            dtLocal.Rows(0)("CobroGenerado") = True
            dtLocal.Rows(0)("IDCobro") = data.IDCobro

            BusinessHelper.UpdateTable(dtLocal)
        End If
    End Sub

#End Region

#Region " Cobros Agrupados "

    <Task()> Public Shared Function CobrosAgrupables(ByVal criterios As Filter, ByVal services As ServiceProvider) As DataTable
        Dim strSelect As String = "Especial,Contabilizado,MIN(IdCliente) AS IdCliente"
        strSelect = strSelect & ",MIN(Titulo) AS Titulo,COUNT(Cobros) AS Cobros"
        strSelect = strSelect & ",SUM(ImpVencimiento) AS ImpVencimiento,MIN(IDCobro) AS IDCobro"
        strSelect = strSelect & ",MIN(AbrvMoneda) AS AbrvMoneda,SUM(ImpVencimientoA) AS ImpVencimientoA"

        Dim strGroupBy As String = "Especial,Contabilizado"
        Dim p As New Parametro
        p.ConfiguracionAgrupacionCobros(strGroupBy)

        Dim fWhere As New Filter
        fWhere.Add(New BooleanFilterItem("Contabilizado", FilterOperator.Equal, False))
        If Not criterios Is Nothing Then fWhere.Add(criterios)
        fWhere.Add(New BooleanFilterItem("Desagrupable", True))
        Dim strWhere As String = AdminData.ComposeFilter(fWhere)
        If Len(strWhere) > 0 Then strWhere = strWhere & " "
        strWhere = strWhere & "GROUP BY " & strGroupBy & " HAVING COUNT(cobros)>1"

        Dim dtCobros As DataTable = AdminData.Filter("vNegCobrosAgrupables", strSelect, strWhere)
        For Each oCol As DataColumn In dtCobros.Columns
            oCol.ReadOnly = False
        Next

        Return dtCobros
    End Function

    <Serializable()> _
    Public Class DataResultCobrosAgrupados
        Public CobrosAgrupables As DataTable
        Public PropuestaCobrosAgrupados As DataTable
    End Class

    <Serializable()> _
   Public Class DataCobrosAgrupados
        Public CobrosAgrupables As DataTable
        Public Criterios As Filter
    End Class

    <Task()> Public Shared Function CobrosAgrupados(ByVal data As DataCobrosAgrupados, ByVal services As ServiceProvider) As DataResultCobrosAgrupados
        Dim p As New Parametro
        Dim strGroupBy As String = "Especial,Contabilizado"
        Dim fAgrupaciones As Filter = p.ConfiguracionAgrupacionCobros(data.CobrosAgrupables, strGroupBy)

        Dim fWhere As New Filter
        If Not data.Criterios Is Nothing Then fWhere.Add(data.Criterios)
        fWhere.Add(New BooleanFilterItem("Contabilizado", FilterOperator.Equal, False))
        fWhere.Add(fAgrupaciones)
        fWhere.Add(New BooleanFilterItem("Desagrupable", True))
        Dim strWhere As String = AdminData.ComposeFilter(fWhere)

        Dim result As New DataResultCobrosAgrupados
        result.CobrosAgrupables = AdminData.Filter("vNegCobrosAgrupados", "*", strWhere)

        Dim strSelect As String = "Especial,Contabilizado,MIN(IdCliente) AS IdCliente"
        strSelect = strSelect & ",MIN(Titulo) AS Titulo,COUNT(Cobros) AS Cobros,MIN(IdFormaPago) AS IdFormaPago, MIN(DescFormaPago) AS DescFormaPago"
        strSelect = strSelect & ",MIN(FechaVencimiento) AS FechaVencimiento,SUM(ImpVencimiento) AS ImpVencimiento"
        strSelect = strSelect & ",MIN(AbrvMoneda) AS AbrvMoneda,SUM(ImpVencimientoA) AS ImpVencimientoA"
        strSelect = strSelect & ",MIN(CambioA) AS CambioA,MIN(CambioB) AS CambioB,MIN(IdMoneda) AS IdMoneda"
        strSelect = strSelect & ",MIN(IdDireccion) AS IdDireccion,MIN(IDClienteBanco) AS IDClienteBanco"
        strSelect = strSelect & ",MIN(IdCondicionPago) AS IdCondicionPago,MIN(IDBancoPropio) AS IDBancoPropio"
        strSelect = strSelect & ",SUM(ARepercutir) AS ARepercutir, SUM(ARepercutirA) AS ARepercutirA"

        Dim strWhere2 As String = AdminData.ComposeFilter(fWhere)
        If Len(strWhere2) > 0 Then strWhere2 = strWhere2 & " "
        strWhere2 = strWhere2 & "GROUP BY " & strGroupBy & " HAVING COUNT(cobros)>1"

        result.PropuestaCobrosAgrupados = AdminData.Filter("vNegNuevosCobrosAgrupados", strSelect, strWhere2)
        For Each oCol As DataColumn In result.PropuestaCobrosAgrupados.Columns
            oCol.ReadOnly = False
        Next
        Return result
    End Function

    <Serializable()> _
    Public Class DataAddCobrosAgrupados
        Public IDProcess As Guid
        Public NuevosCobros As DataTable
    End Class
    <Task()> Public Shared Function AddCobrosAgrupados(ByVal data As DataAddCobrosAgrupados, ByVal services As ServiceProvider) As ClassErrors()
        Dim Errores(-1) As ClassErrors
        Dim co As New Cobro
        ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)

        Dim blnError As Boolean
        Dim dtCobrosSelec As DataTable = New BE.DataEngine().Filter("vNegCobrosAgrupados", New GuidFilterItem("IDProcess", data.IDProcess))
        If Not dtCobrosSelec Is Nothing AndAlso dtCobrosSelec.Rows.Count > 0 Then
            Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
            Dim AppParamsTes As ParametroTesoreria = services.GetService(Of ParametroTesoreria)()
            Dim AppParamsFV As ParametroFacturaVenta = services.GetService(Of ParametroFacturaVenta)()
            'Dim strIN As String
            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Dim dtCobro As DataTable = co.AddNew
            For Each drNewCobro As DataRow In data.NuevosCobros.Rows
                Dim NewRow As DataRow = dtCobro.NewRow
                NewRow("IdCobro") = AdminData.GetAutoNumeric
                NewRow("IDCliente") = drNewCobro("IDCliente")

                If Nz(drNewCobro("Especial"), False) Then
                    NewRow("IDTipoCobro") = AppParamsFV.TipoCobroFacturaVentaB
                End If
                Dim ClteInfo As ClienteInfo = Clientes.GetEntity(NewRow("IDCliente"))
                If Not ClteInfo Is Nothing AndAlso Length(ClteInfo.IDCliente) > 0 Then
                    If AppParamsConta.Contabilidad Then
                        If Length(ClteInfo.CCCliente) Then
                            NewRow("CContable") = ClteInfo.CCCliente
                        Else
                            ReDim Preserve Errores(Errores.Length)
                            Errores(Errores.Length - 1) = New ClassErrors
                            Errores(Errores.Length - 1).Elements = NewRow("IDCliente")
                            Errores(Errores.Length - 1).MessageError = Engine.ParseFormatString(AdminData.GetMessageText("La Cuenta Contable del Cliente {0} es un dato obligatorio."), Quoted(NewRow("IDCliente")))

                            blnError = True
                        End If
                    End If
                End If
                NewRow("Situacion") = dtCobrosSelec.Rows(0)("Situacion")    'Se coge la situación del primer cobro
                NewRow("Titulo") = drNewCobro("Titulo")
                NewRow("IDFormaPago") = drNewCobro("IDFormaPago")
                NewRow("FechaVencimiento") = drNewCobro("FechaVencimiento")
                NewRow("ImpVencimiento") = drNewCobro("ImpVencimiento")
                NewRow("ARepercutir") = drNewCobro("ARepercutir")
                NewRow("ARepercutirA") = drNewCobro("ARepercutirA")
                NewRow("CambioA") = drNewCobro("CambioA")
                NewRow("CambioB") = drNewCobro("CambioB")
                NewRow("IDMoneda") = drNewCobro("IDMoneda")
                NewRow("NFactura") = AppParamsTes.NFacturaCobroAgupado
                NewRow("IdClienteBanco") = drNewCobro("IdClienteBanco")
                NewRow("IdDireccion") = drNewCobro("IdDireccion")
                NewRow("IDBancoPropio") = drNewCobro("IDBancoPropio")

                If Not blnError AndAlso Length(NewRow("IDClienteBanco")) > 0 AndAlso dtCobrosSelec.Columns.Contains("IDMandato") Then
                    Dim ClienteBancoDistintos As List(Of Object) = (From c In dtCobrosSelec _
                                                               Where Not c.IsNull("IDMandato") AndAlso c("IDCliente") = NewRow("IDCliente") _
                                                               AndAlso (Not c.IsNull("IDClienteBanco") AndAlso c("IDClienteBanco") <> NewRow("IDClienteBanco") OrElse _
                                                                       c.IsNull("IDClienteBanco")) _
                                                               Select c("IDClienteBanco") Distinct).ToList
                    If ClienteBancoDistintos.Count > 0 Then
                        '//Registramos el mensaje de "error", pero no cancelamos la agrupación
                        ReDim Preserve Errores(Errores.Length)
                        Errores(Errores.Length - 1) = New ClassErrors
                        Errores(Errores.Length - 1).Elements = NewRow("IDCliente")
                        Errores(Errores.Length - 1).MessageError = Engine.ParseFormatString(AdminData.GetMessageText("Los Cobros del Cliente {0} tienen asociados Bancos diferentes. No se asignará el Mandato."), Quoted(NewRow("IDCliente")))
                    Else
                        Dim MandatosDistintos As List(Of Object) = (From c In dtCobrosSelec _
                                                                    Where Not c.IsNull("IDMandato") AndAlso c("IDCliente") = NewRow("IDCliente") _
                                                                    AndAlso Not c.IsNull("IDClienteBanco") AndAlso c("IDClienteBanco") = NewRow("IDClienteBanco") _
                                                                    Select c("IDMandato") Distinct).ToList
                        If MandatosDistintos.Count > 1 Then
                            '//Registramos el mensaje de "error", pero no cancelamos la agrupación
                            ReDim Preserve Errores(Errores.Length)
                            Errores(Errores.Length - 1) = New ClassErrors
                            Errores(Errores.Length - 1).Elements = NewRow("IDCliente")
                            Errores(Errores.Length - 1).MessageError = Engine.ParseFormatString(AdminData.GetMessageText("Los Cobros del Cliente {0} tienen asociados Mandatos diferentes.  No se asignará el Mandato."), Quoted(NewRow("IDCliente")))
                        ElseIf MandatosDistintos.Count = 1 Then
                            If Length(NewRow("IDFormaPago")) > 0 Then
                                Dim FormasPago As EntityInfoCache(Of FormaPagoInfo) = services.GetService(Of EntityInfoCache(Of FormaPagoInfo))()
                                Dim FPInfo As FormaPagoInfo = FormasPago.GetEntity(NewRow("IDFormaPago"))
                                If FPInfo.CobroRemesable Then
                                    NewRow("IDMandato") = MandatosDistintos(0)
                                End If
                            End If
                        End If
                    End If
                End If

                If Not blnError Then
                    'Actualización del Cobro
                    Dim dv As DataView = dtCobrosSelec.DefaultView
                    dv.RowFilter = "IDCliente='" & NewRow("IDCliente") & "'"
                    If Not dv Is Nothing Then
                        For Each dr As DataRowView In dv
                            dr.Row("IDCobroAgrupado") = NewRow("IDCobro")
                        Next
                    End If
                    dv.RowFilter = ""
                    dtCobro.Rows.Add(NewRow)
                Else
                    blnError = False
                End If
            Next

            'Recuperar registros de devolución
            Dim fDevoluciones As New Filter(FilterUnionOperator.Or)
            For Each drCobro As DataRow In dtCobrosSelec.Rows
                fDevoluciones.Add("IDCobro", FilterOperator.Equal, drCobro("IDCobro"), FilterType.Numeric)
            Next
            Dim ClsCoDev As New CobroDevolucion
            Dim dtDevoluciones As DataTable = ClsCoDev.Filter(fDevoluciones)
            Dim dtNewDevolucion As DataTable
            If Not dtDevoluciones Is Nothing AndAlso dtDevoluciones.Rows.Count > 0 Then
                dtNewDevolucion = ClsCoDev.AddNew
                For Each row As DataRow In dtDevoluciones.Rows
                    Dim drNew As DataRow = dtNewDevolucion.NewRow
                    drNew.ItemArray = row.ItemArray
                    drNew("IDDevolucion") = AdminData.GetAutoNumeric
                    drNew("IDCobro") = dtCobro.Rows(0)("IDCobro")
                    dtNewDevolucion.Rows.Add(drNew)
                Next
            End If

            co.Update(dtCobro)
            dtCobrosSelec.TableName = GetType(Cobro).Name
            BusinessHelper.UpdateTable(dtCobrosSelec)

            If Not dtNewDevolucion Is Nothing AndAlso dtNewDevolucion.Rows.Count > 0 Then BusinessHelper.UpdateTable(GetType(CobroDevolucion).Name, dtNewDevolucion)

            Return Errores
        Else
            ApplicationService.GenerateError("No existe ningún cobro seleccionado para agrupar.")
        End If
    End Function

    <Serializable()> _
    Public Class DataDesagruparCobros
        Public CobrosAgrupados As DataTable
        Public CobrosDesagrupados As DataTable
    End Class
    <Task()> Public Shared Sub EliminarCobroAgrupado(ByVal data As DataDesagruparCobros, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
        If Not IsNothing(data.CobrosAgrupados) AndAlso data.CobrosAgrupados.Rows.Count > 0 Then
            Dim dtCobrosEliminar As DataTable = data.CobrosAgrupados.Clone
            'Dim CobroAgrupado As String = New Parametro().NFacturaCobroAgupado
            'If Length(CobroAgrupado) = 0 Then ApplicationService.GenerateError("Revise el parámetro de indica el texto de Cobro Agrupado.")

            Dim IDCobrosAgrupadores As List(Of Object) = (From c In data.CobrosAgrupados Select c("IDCobro") Distinct).ToList
            Dim dtCobrosAgrupados As DataTable = New BE.DataEngine().Filter("tbCobro", New InListFilterItem("IDCobroAgrupado", IDCobrosAgrupadores.ToArray, FilterType.Numeric), "IDCobroAgrupado")
            Dim IDCobrosAgrupadoresBBDD As List(Of Object) = (From c In dtCobrosAgrupados Select c("IDCobroAgrupado") Distinct).ToList

            Dim dtEstCobro As DataTable = New EstadoCobro().Filter(New BooleanFilterItem("Desagrupable", True), , "IDEstado")

            For Each drCobroAgrupado As DataRow In data.CobrosAgrupados.Rows
                If drCobroAgrupado("Contabilizado") = CBool(enumCobroContabilizado.CobroContabilizado) Then
                    ApplicationService.GenerateError("Algún Cobro Agrupado está contabilizado. Debe descontabilizar el Cobro antes de desagruparlo.")
                End If
                'If drCobroAgrupado("Situacion") <> enumCobroSituacion.NoNegociado Then
                '    ApplicationService.GenerateError("Algún Cobro Agrupado está en una Situación distinta de No Negociado. No se puede deshacer la agrupación.")
                'End If
                If Not dtEstCobro Is Nothing AndAlso dtEstCobro.Rows.Count > 0 Then
                    Dim dr() As DataRow = dtEstCobro.Select("IDEstado = " & drCobroAgrupado("Situacion"))
                    If dr.Length = 0 Then
                        ApplicationService.GenerateError("Algún Cobro Agrupado está en una Situación que no es desagrupable. No se puede deshacer la agrupación.")
                    End If
                End If

                ' If drCobroAgrupado("NFactura") & String.Empty = CobroAgrupado Then
                If Not IDCobrosAgrupadoresBBDD Is Nothing AndAlso IDCobrosAgrupadoresBBDD.Contains(drCobroAgrupado("IDCobro")) Then
                    dtCobrosEliminar.ImportRow(drCobroAgrupado)
                End If
                'End If
            Next

            If Not IsNothing(data.CobrosDesagrupados) AndAlso data.CobrosDesagrupados.Rows.Count > 0 Then
                For Each dr As DataRow In data.CobrosDesagrupados.Select
                    dr("IDCobroAgrupado") = System.DBNull.Value
                Next

                data.CobrosDesagrupados.TableName = GetType(Cobro).Name
                BusinessHelper.UpdateTable(data.CobrosDesagrupados)

                Dim c As New Cobro
                c.Delete(dtCobrosEliminar)
            End If
        End If
    End Sub



#End Region

#Region " Cambio Situación "

    <Serializable()> _
   Public Class DataCambioSituacionManual
        Public Cobros As DataTable
        Public NuevaSituacion As enumCobroSituacion?
        Public NuevaFechaCobro As Date?

        Public Sub New()
        End Sub

        Public Sub New(ByVal Cobros As DataTable, Optional ByVal NuevaSituacion As enumCobroSituacion = -1, Optional ByVal NuevaFechaCobro As Date = cnMinDate)
            Me.Cobros = Cobros
            If NuevaSituacion <> 1 Then Me.NuevaSituacion = NuevaSituacion
            If NuevaFechaCobro <> cnMinDate Then Me.NuevaFechaCobro = NuevaFechaCobro
        End Sub
    End Class

    <Task()> Public Shared Function CambioSituacionManual(ByVal data As DataCambioSituacionManual, ByVal services As ServiceProvider) As ClassErrors()
        If Not data.Cobros Is Nothing Then
            Dim NuevaSituacion As enumCobroSituacion
            Dim Errores(-1) As ClassErrors
            Dim ResultState As enumResultadoCambioCobros
            For Each dr As DataRow In data.Cobros.Rows
                If dr("Situacion") <> dr("Situacion", DataRowVersion.Original) Then
                    If dr("Situacion") = enumCobroSituacion.Cobrado Then
                        ResultState = enumResultadoCambioCobros.CobradoNoContabilizado
                    Else
                        ResultState = enumResultadoCambioCobros.Ok
                    End If
                Else
                    ResultState = enumResultadoCambioCobros.Ok
                    '//Si no se ha introducido una única situación para todos los registros, ésta deberá venir indicada en el Cobro.
                    If data.NuevaSituacion Is Nothing Then
                        NuevaSituacion = dr("Situacion")
                    Else
                        NuevaSituacion = data.NuevaSituacion
                        dr("Situacion") = NuevaSituacion
                    End If

                    Dim EstadosCobro As EntityInfoCache(Of EstadoCobroInfo) = services.GetService(Of EntityInfoCache(Of EstadoCobroInfo))()
                    Dim EstCobInfo As EstadoCobroInfo = EstadosCobro.GetEntity(NuevaSituacion)
                    If Not EstCobInfo Is Nothing AndAlso Length(EstCobInfo.IDEstado) > 0 Then
                        dr("IDAgrupacion") = EstCobInfo.IDAgrupacion
                    End If

                    If NuevaSituacion = enumCobroSituacion.Cobrado Then
                        If dr("Contabilizado") = False Then
                            ResultState = enumResultadoCambioCobros.CobradoNoContabilizado

                            Dim FormasPago As EntityInfoCache(Of FormaPagoInfo) = services.GetService(Of EntityInfoCache(Of FormaPagoInfo))()
                            Dim FPagoInfo As FormaPagoInfo = FormasPago.GetEntity(dr("IDFormaPago"))
                            If Not FPagoInfo Is Nothing AndAlso Length(FPagoInfo.IDFormaPago) > 0 AndAlso FPagoInfo.CobroRemesable Then
                                ResultState = enumResultadoCambioCobros.CobradoNoContabilizadoRemesable
                            End If
                        End If

                        '//Si no se ha introducido una única fecha Cobro para todos los registros, ésta deberá venir indicada en el Cobro.
                        If Not data.NuevaFechaCobro Is Nothing Then
                            dr("FechaCobro") = data.NuevaFechaCobro
                            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf Cobro.ActualizarFechaParaDeclaracionFactura, New DataRowPropertyAccessor(dr), services)
                        End If
                    End If
                End If

                If ResultState <> enumResultadoCambioCobros.Ok Then
                    ReDim Preserve Errores(Errores.Length)
                    Errores(Errores.Length - 1) = New ClassErrors
                    Errores(Errores.Length - 1).Elements = dr("IDCobro")
                    Errores(Errores.Length - 1).MessageError = ResultState
                End If
            Next

            data.Cobros.TableName = GetType(Cobro).Name
            BusinessHelper.UpdateTable(data.Cobros)
            Return Errores
        End If
    End Function

    <Task()> Public Shared Sub ActualizarFechaCobroDesdePago(ByVal drPago As IPropertyAccessor, ByVal services As ServiceProvider)
        If Length(drPago("IDCobro")) > 0 Then
            Dim dtCobro As DataTable = New Cobro().SelOnPrimaryKey(drPago("IDCobro"))
            If dtCobro.Rows.Count > 0 Then
                Dim Situacion As enumCobroSituacion = enumCobroSituacion.Cobrado
                If Nz(drPago("FechaPago"), cnMinDate) = cnMinDate Then
                    Situacion = enumCobroSituacion.GeneradoPago
                    dtCobro.Rows(0)("FechaCobro") = System.DBNull.Value
                End If
                Dim datosCambio As New Cobro.DataCambioSituacionManual(dtCobro, Situacion, Nz(drPago("FechaPago"), cnMinDate))
                ProcessServer.ExecuteTask(Of Cobro.DataCambioSituacionManual)(AddressOf Cobro.CambioSituacionManual, datosCambio, services)
            End If
        End If
    End Sub

#End Region

#Region " Devolución de cobros "

    <Task()> Public Shared Function DevolucionCobros(ByVal dtDatosDevolucion As DataTable, ByVal services As ServiceProvider) As Integer()
        If dtDatosDevolucion Is Nothing OrElse dtDatosDevolucion.Rows.Count = 0 Then Exit Function

        Dim IDDevoluciones(-1) As Integer
        Dim c As New Cobro
        Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
        For Each drCobro As DataRow In dtDatosDevolucion.Rows
            Dim IDCobro As Integer = drCobro("IDCobro")
            Dim dtCobro As DataTable = c.SelOnPrimaryKey(IDCobro)
            If Not IsNothing(dtCobro) AndAlso dtCobro.Rows.Count > 0 Then
                Dim dtDevolucion As DataTable

                Dim dvDatosDevolucion As New DataView(dtDatosDevolucion)
                dvDatosDevolucion.RowFilter = "IDCobro = " & IDCobro
                Dim IDRemesa As Integer = Nz(dtCobro.Rows(0)("IdRemesa"), 0)

                If dvDatosDevolucion.Count > 0 Then
                    Dim blnInsertarDev As Boolean = (Not AppParamsConta.ContabilidadMultiple)
                    If AppParamsConta.ContabilidadMultiple Then blnInsertarDev = dtCobro.Rows(0)("Contabilizado") <> enumContabilizado.ContabilizadoNIIF AndAlso dtCobro.Rows(0)("Contabilizado") <> enumContabilizado.ContabilizadoTributario
                    If blnInsertarDev Then
                        dtDevolucion = ProcessServer.ExecuteTask(Of DataView, DataTable)(AddressOf InsertarDevolucion, dvDatosDevolucion, services)
                        dtCobro.Rows(0)("ARepercutir") = dtCobro.Rows(0)("ARepercutir") + dvDatosDevolucion(0).Row("ARepercutir")
                        dtCobro.Rows(0)("ARepercutirA") = dtCobro.Rows(0)("ARepercutirA") + dvDatosDevolucion(0).Row("ARepercutirA")
                        dtCobro.Rows(0)("ARepercutirB") = dtCobro.Rows(0)("ARepercutirB") + dvDatosDevolucion(0).Row("ARepercutirB")
                    End If

                End If
                'dtCobro.Rows(0)("Contabilizado") = enumContabilizado.NoContabilizado

                Dim datValEst As New Comunes.DataValidarEstado(dtCobro.Rows(0)("IDCobro"), enumDiarioTipoApunte.DevolucionRemesa)
                If AppParamsConta.Contabilidad Then
                    If AppParamsConta.ContabilidadMultiple Then
                        '//Si estaba contabilizado completo, lo descontabilizamos. En este caso no se borran los apuntes del diario, por lo que sabremos en cuantos ejercios está contabilizado.
                        Dim Contabilizado As enumContabilizado = ProcessServer.ExecuteTask(Of Comunes.DataValidarEstado, Integer)(AddressOf Comunes.ValidarEstado, datValEst, services)
                        If Contabilizado = enumContabilizado.Contabilizado Then
                            dtCobro.Rows(0)("Contabilizado") = enumContabilizado.NoContabilizado
                        Else
                            dtCobro.Rows(0)("Contabilizado") = Contabilizado
                        End If
                    Else
                        dtCobro.Rows(0)("Contabilizado") = enumContabilizado.NoContabilizado
                    End If
                Else
                    dtCobro.Rows(0)("Contabilizado") = enumContabilizado.NoContabilizado
                End If

                AdminData.BeginTx()

                If dtCobro.Rows(0)("Contabilizado") = enumContabilizado.NoContabilizado Then
                    dtCobro.Rows(0)("Situacion") = enumCobroSituacion.Devuelto
                    'dtCobro.Rows(0)("IDEjercicio") = System.DBNull.Value
                    dtCobro.Rows(0)("FechaContabilizacion") = System.DBNull.Value
                    dtCobro.Rows(0)("FechaCobro") = System.DBNull.Value
                    dtCobro.Rows(0)("IdRemesa") = System.DBNull.Value
                    dtCobro.Rows(0)("Liquidado") = enumContabilizado.NoContabilizado

                    ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf Cobro.ActualizarFechaParaDeclaracionFactura, New DataRowPropertyAccessor(dtCobro.Rows(0)), services)
                    ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf Pago.ActualizarFechaPagoDesdeCobro, New DataRowPropertyAccessor(dtCobro.Rows(0)), services)
                Else
                    IDRemesa = 0
                End If


                '//Cambiar el tipo del asiento de liquidación (si es que tenia) y ponerlo de tipo remesa. Esto se hace por si hay varios cobros en la remesa
                '//como este cobro va a quedar desvinculado de la remesa, para que no aparezcan en los asientos de liquidación de la remesa. Sólo debemos modificar el apunte
                '//del cobro (podemos tener IDCobro en el IDDocumento del apunte del banco, en este caso el apunte del banco no debemos cambiarlo, para podernos apoyar en él por que 
                '// en él tenemos en el NDocumento el NºRemesa)
                If AppParamsConta.Contabilidad AndAlso dtCobro.Rows(0)("Contabilizado") = enumContabilizado.NoContabilizado AndAlso IDRemesa <> 0 Then
                    Dim datLiquidacion As New DataTipoLiquidacionATipoRemesa(IDCobro, IDRemesa)
                    Dim dtDiario As DataTable = ProcessServer.ExecuteTask(Of DataTipoLiquidacionATipoRemesa, DataTable)(AddressOf TipoLiquidacionATipoRemesa, datLiquidacion, services)
                    BusinessHelper.UpdateTable(dtDiario)
                End If
                BusinessHelper.UpdateTable(dtCobro)
                BusinessHelper.UpdateTable(dtDevolucion)  'CobroDevolucion

                ReDim Preserve IDDevoluciones(IDDevoluciones.Length)
                IDDevoluciones(IDDevoluciones.Length - 1) = dtDevolucion.Rows(0)("IDDevolucion")

                '//Comprobamos si la remesa de este cobro tiene mas cobros
                'If IDRemesa <> 0 Then
                '    dtCobro = c.Filter(New NumberFilterItem("IDRemesa", IDRemesa))
                '    If IsNothing(dtCobro) OrElse dtCobro.Rows.Count = 0 Then
                '        'Si la remesa ya no tiene cobros, se elimina
                '        AdminData.Execute("DELETE FROM tbRemesa WHERE IdRemesa=" & IDRemesa)
                '    End If
                'End If
                AdminData.CommitTx(True)
            End If
        Next
        Return IDDevoluciones
    End Function

    <Task()> Public Shared Function InsertarDevolucion(ByVal DvCobroDevolucion As DataView, ByVal services As ServiceProvider) As DataTable
        Dim ClsCobroDev As New CobroDevolucion
        Dim DtNewCobroDev As DataTable = ClsCobroDev.AddNewForm
        DtNewCobroDev.Rows(0)("IdCobro") = DvCobroDevolucion(0)("IdCobro")
        DtNewCobroDev.Rows(0)("FechaDevolucion") = DvCobroDevolucion(0)("FechaDevolucion")
        DtNewCobroDev.Rows(0)("Gasto") = DvCobroDevolucion(0)("Gasto")
        DtNewCobroDev.Rows(0)("GastoA") = DvCobroDevolucion(0)("GastoA")
        DtNewCobroDev.Rows(0)("GastoB") = DvCobroDevolucion(0)("GastoB")
        DtNewCobroDev.Rows(0)("Comision") = DvCobroDevolucion(0)("Comision")
        DtNewCobroDev.Rows(0)("ComisionA") = DvCobroDevolucion(0)("ComisionA")
        DtNewCobroDev.Rows(0)("ComisionB") = DvCobroDevolucion(0)("ComisionB")
        DtNewCobroDev.Rows(0)("ARepercutir") = DvCobroDevolucion(0)("ARepercutir")
        DtNewCobroDev.Rows(0)("ARepercutirA") = DvCobroDevolucion(0)("ARepercutirA")
        DtNewCobroDev.Rows(0)("ARepercutirB") = DvCobroDevolucion(0)("ARepercutirB")

        If DvCobroDevolucion.Table.Columns.Contains("IDEjercicio") Then DtNewCobroDev.Rows(0)("IDEjercicio") = DvCobroDevolucion(0)("IDEjercicio")
        If DvCobroDevolucion.Table.Columns.Contains("NAsiento") Then DtNewCobroDev.Rows(0)("NAsiento") = DvCobroDevolucion(0)("NAsiento")
        If DvCobroDevolucion.Table.Columns.Contains("IDEjercicioTributario") Then DtNewCobroDev.Rows(0)("IDEjercicioTributario") = DvCobroDevolucion(0)("IDEjercicioTributario")
        If DvCobroDevolucion.Table.Columns.Contains("NAsientoTributario") Then DtNewCobroDev.Rows(0)("NAsientoTributario") = DvCobroDevolucion(0)("NAsientoTributario")
        If DvCobroDevolucion.Table.Columns.Contains("IDFacturaCompra") Then DtNewCobroDev.Rows(0)("IDFacturaCompra") = DvCobroDevolucion(0)("IDFacturaCompra")
        If DvCobroDevolucion.Table.Columns.Contains("FechaCobro") Then DtNewCobroDev.Rows(0)("FechaCobro") = DvCobroDevolucion(0)("FechaCobro")

        If Length(DvCobroDevolucion(0)("IDRemesa")) > 0 Then DtNewCobroDev.Rows(0)("IDRemesaAnterior") = DvCobroDevolucion(0)("IDRemesa")
        Return DtNewCobroDev
    End Function

    <Serializable()> _
    Public Class DataTipoLiquidacionATipoRemesa
        Public IDCobro As Integer
        Public IDRemesa As Integer
        Public Inversa As Boolean

        Public Sub New(ByVal IDCobro As Integer, ByVal IDRemesa As Integer)
            Me.IDCobro = IDCobro
            Me.IDRemesa = IDRemesa
        End Sub
    End Class

    <Task()> Public Shared Function TipoLiquidacionATipoRemesa(ByVal data As DataTipoLiquidacionATipoRemesa, ByVal services As ServiceProvider) As DataTable
        Dim DC As BusinessHelper = CreateBusinessObject("DiarioContable")

        '//El NºRemesa va en el apunte del banco no en el del cobro, pero el IDcobro puede ir tb en el apunte del Banco, sólo debemos cambiar el apunte del cobro.
        '//Buscamos el NAsiento, IDEjercicio, en el que tenemos que cambiar el tipo de asiento (por si tenemos varias devoluciones del un mismo cobro)
        Dim fBuscarAsiento As New Filter
        fBuscarAsiento.Add(New NumberFilterItem("IDDocumento", data.IDCobro))
        'If data.Inversa Then
        '    fBuscarAsiento.Add(New NumberFilterItem("IDTipoApunte", enumDiarioTipoApunte.Remesa))
        'Else
        fBuscarAsiento.Add(New NumberFilterItem("IDTipoApunte", enumDiarioTipoApunte.LiquidacionRemesa))
        'End If
        Dim dtAsientosLiquidacion As DataTable = DC.Filter(fBuscarAsiento)
        Dim fAsientos As New Filter(FilterUnionOperator.Or)
        Dim fAsiento As New Filter
        Dim AsientosLiquidacion As List(Of DataRow) = (From c In dtAsientosLiquidacion _
                                                       Where (Not c.isnull("NDocumento")) AndAlso (c("NDocumento") = CStr(data.IDRemesa))).ToList()
        For Each dr As DataRow In AsientosLiquidacion
            fAsiento = New Filter
            fAsiento.Add(New StringFilterItem("IDEjercicio", dr("IDEjercicio")))
            fAsiento.Add(New NumberFilterItem("NAsiento", dr("NAsiento")))

            fAsientos.Add(fAsiento)
        Next

        Dim f As New Filter
        f.Add(fAsientos)
        f.Add(New NumberFilterItem("IDDocumento", data.IDCobro))
        f.Add(New StringFilterItem("NDocumento", FilterOperator.NotEqual, data.IDRemesa)) '//Para evitar el apunte del banco, el NºRemesa va en el apunte del banco no en el del cobro
        If data.Inversa Then
            f.Add(New NumberFilterItem("IDTipoApunte", enumDiarioTipoApunte.Remesa))
        Else
            f.Add(New NumberFilterItem("IDTipoApunte", enumDiarioTipoApunte.LiquidacionRemesa))
        End If

        Dim dtDiario As DataTable = DC.Filter(f)
        For Each dr As DataRow In dtDiario.Rows
            If data.Inversa Then
                dr("IDTipoApunte") = enumDiarioTipoApunte.LiquidacionRemesa
            Else
                dr("IDTipoApunte") = enumDiarioTipoApunte.Remesa
            End If
        Next

        Return dtDiario
    End Function

#End Region

#Region "Insertar Cobros"

    <Serializable()> _
    Public Class dataInsertarCobros
        Public Cobros As DataTable
        Public CobroSinFactura As Boolean

        Public Sub New(ByVal Cobros As DataTable, ByVal CobroSinFactura As Boolean)
            Me.Cobros = Cobros
            Me.CobroSinFactura = CobroSinFactura
        End Sub
    End Class
    <Task()> Public Shared Sub InsertarCobros(ByVal data As dataInsertarCobros, ByVal services As ServiceProvider)
        If data.Cobros.Rows.Count > 0 Then
            ProcessServer.ExecuteTask(Of dataInsertarCobros)(AddressOf TratarCobrosSinFactura, data, services)
            ProcessServer.ExecuteTask(Of dataInsertarCobros)(AddressOf AddCobros, data, services)
        End If
    End Sub

    <Task()> Public Shared Sub TratarCobrosSinFactura(ByVal data As dataInsertarCobros, ByVal services As ServiceProvider)
        If data.CobroSinFactura Then
            Dim ParamsTesoreria As ParametroTesoreria = services.GetService(Of ParametroTesoreria)()
            Dim ParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
            Dim IDTipoCobro As Integer = ParamsTesoreria.TipoCobroFV
            For Each drCobro As DataRow In data.Cobros.Rows
                If ParamsConta.Contabilidad AndAlso Length(drCobro("CContable")) = 0 Then
                    ApplicationService.GenerateError("La C. Contable es un campo obligatorio. El Cliente no tiene asignada una C.Contable.")
                Else
                    drCobro("IDTipoCobro") = IDTipoCobro
                    drCobro("NFactura") = System.DBNull.Value
                    drCobro("Situacion") = enumCobroSituacion.NoNegociado
                    drCobro("FechaVencimientoFactura") = drCobro("FechaVencimiento")
                    If drCobro("CambioA") > 0 Then
                        drCobro("ImpVencimiento") = drCobro("ImpTotalA") / drCobro("CambioA")
                    Else
                        drCobro("ImpVencimiento") = 0
                    End If
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Sub AddCobros(ByVal data As dataInsertarCobros, ByVal services As ServiceProvider)
        For Each dr As DataRow In data.Cobros.Rows
            Dim C As New Cobro
            Dim dtCobro As DataTable = C.AddNew
            Dim drCobro As DataRow = dtCobro.NewRow

            drCobro("IDTipoCobro") = dr("IDTipoCobro")
            drCobro("IDCliente") = dr("IDCliente")
            drCobro("Titulo") = dr("Titulo")
            If (Length(dr("Titulo")) = 0 Or Length(dr("IDBancoPropio")) = 0) AndAlso Length(dr("IDCliente")) > 0 Then
                Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
                Dim Cliente As ClienteInfo = Clientes.GetEntity(dr("IDCliente"))
                If Not Cliente Is Nothing AndAlso Length(Cliente.IDCliente) > 0 Then
                    drCobro("Titulo") = Cliente.DescCliente
                    drCobro("IDBancoPropio") = Cliente.IDBancoPropio
                End If
            End If

            If Length(dr("IDClienteBanco")) = 0 Then
                Dim IDClienteBanco As Integer = ProcessServer.ExecuteTask(Of String, Integer)(AddressOf ClienteBanco.GetBancoPredeterminado, drCobro("IDCliente"), services)
                If IDClienteBanco > 0 Then drCobro("IDClienteBanco") = IDClienteBanco
            Else
                drCobro("IDClienteBanco") = dr("IDClienteBanco")
            End If

            Dim dataDireccionGiro As New ClienteDireccion.DataDirecEnvio(drCobro("IDCliente"), enumcdTipoDireccion.cdDireccionGiro)
            Dim dtDireccion As DataTable = ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecEnvio, DataTable)(AddressOf ClienteDireccion.ObtenerDireccionEnvio, dataDireccionGiro, services)
            If dtDireccion.Rows.Count > 0 Then drCobro("IDDireccion") = dtDireccion.Rows(0)("IDDireccion")

            If Length(dr("NFactura")) > 0 Then drCobro("NFactura") = dr("NFactura")
            If Length(dr("IDFactura")) > 0 Then drCobro("IDFactura") = dr("IDFactura")

            drCobro("CContable") = dr("CContable")
            drCobro("IDFormaPago") = dr("IDFormaPago")
            drCobro("IDBancoPropio") = dr("IDBancoPropio")
            drCobro("IDMoneda") = dr("IDMoneda")
            drCobro("Situacion") = dr("Situacion")
            drCobro("FechaVencimiento") = dr("FechaVencimiento")
            If Length(dr("FechaVencimientoFactura")) > 0 Then
                drCobro("FechaVencimientoFactura") = dr("FechaVencimientoFactura")
            Else
                drCobro("FechaVencimientoFactura") = dr("FechaVencimiento")
            End If
            drCobro("CambioA") = dr("CambioA")
            drCobro("CambioB") = dr("CambioB")

            drCobro("ImpVencimiento") = dr("ImpVencimiento")

            dtCobro.Rows.Add(drCobro)
            C.Update(dtCobro)

            If data.Cobros.Columns.Contains("IDLocalVencimiento") AndAlso dr("IDLocalVencimiento") Then
                Dim datosPromoVto As New Cobro.dataAsociarCobroVencimientosLocalPromociones(dr("IDLocalVencimiento"), dtCobro.Rows(0)("IDCobro"))
                ProcessServer.ExecuteTask(Of Cobro.dataAsociarCobroVencimientosLocalPromociones)(AddressOf AsociarCobroVencimientosLocalPromociones, datosPromoVto, services)
            End If
        Next

    End Sub


    '<Serializable()> _
    'Public Class DataInsertarCobros
    '    Public Cobros As DataTable
    '    Public CobroSinFactura As Boolean
    'End Class

    '<Task()> Public Shared Sub InsertarCobros(ByVal data As DataInsertarCobros, ByVal services As ServiceProvider)
    '    ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
    '    ProcessServer.ExecuteTask(Of DataInsertarCobros)(AddressOf TratarCobrosSinFactura, data, services)

    '    If Not data.Cobros Is Nothing AndAlso data.Cobros.Rows.Count > 0 Then
    '        Dim c As New Cobro
    '        Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
    '        For Each dr As DataRow In data.Cobros.Rows
    '            Dim dtCobro As DataTable = c.AddNew
    '            Dim drCobro As DataRow = dtCobro.NewRow

    '            Dim datos As New DataCopiaDataRow
    '            Dim DtOrigen As DataTable = dr.Table.Clone
    '            DtOrigen.ImportRow(dr)
    '            datos.DtOrigen = DtOrigen
    '            Dim DtDestino As DataTable = dtCobro.Clone
    '            DtDestino.ImportRow(drCobro)
    '            datos.DtDestino = dtCobro
    '            ProcessServer.ExecuteTask(Of DataCopiaDataRow)(AddressOf AsignarDatosDesdeCobroOrigen, datos, services)
    '            ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarDatosCliente, drCobro, services)
    '            ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarClienteBancoPredeterminado, drCobro, services)
    '            ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarClienteDireccion, drCobro, services)
    '            drCobro.ItemArray = datos.DtDestino.Rows(0).ItemArray
    '            dtCobro.Rows.Add(drCobro)
    '            dtCobro = c.Update(dtCobro)

    '            If data.Cobros.Columns.Contains("IDLocalVencimiento") AndAlso dr("IDLocalVencimiento") Then
    '                Dim datosPromoVto As Cobro.DataAsociarCobroVencimientosLocalPromociones
    '                datosPromoVto.IDLocalVencimiento = dr("IDLocalVencimiento")
    '                datosPromoVto.IDCobro = dtCobro.Rows(0)("IDCobro")
    '                ProcessServer.ExecuteTask(Of Cobro.DataAsociarCobroVencimientosLocalPromociones)(AddressOf AsociarCobroVencimientosLocalPromociones, datosPromoVto, services)
    '            End If
    '        Next
    '    End If
    '    ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.CommitTransaction, True, services)
    'End Sub

    '<Task()> Public Shared Sub TratarCobrosSinFactura(ByVal data As DataInsertarCobros, ByVal services As ServiceProvider)
    '    If data.CobroSinFactura Then
    '        Dim AppParams As ParametroTesoreria = services.GetService(Of ParametroTesoreria)()
    '        Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
    '        Dim intIDTipoCobro As Integer = AppParams.TipoCobroFV
    '        For Each dr As DataRow In data.Cobros.Rows
    '            If AppParamsConta.Contabilidad AndAlso Length(dr("CContable")) = 0 Then
    '                ApplicationService.GenerateError("La C. Contable es un campo obligatorio. El Cliente no tiene asignada una C.Contable.")
    '            Else
    '                dr("IDTipoCobro") = intIDTipoCobro
    '                dr("NFactura") = System.DBNull.Value
    '                dr("Situacion") = enumCobroSituacion.NoNegociado
    '                dr("FechaVencimientoFactura") = dr("FechaVencimiento")
    '                If dr("CambioA") > 0 Then
    '                    dr("ImpVencimiento") = dr("ImpTotalA") / dr("CambioA")
    '                Else
    '                    dr("ImpVencimiento") = 0
    '                End If
    '            End If
    '        Next
    '    End If
    'End Sub

    '<Task()> Public Shared Sub AsignarDatosCliente(ByVal data As DataRow, ByVal services As ServiceProvider)
    '    If (Length(data("Titulo")) = 0 Or Length(data("IDBancoPropio")) = 0) AndAlso Length(data("IdCliente")) > 0 Then
    '        Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
    '        Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data("IdCliente"))
    '        If Not ClteInfo Is Nothing AndAlso Length(ClteInfo.IDCliente) > 0 Then
    '            data("Titulo") = ClteInfo.DescCliente
    '            data("IDBancoPropio") = ClteInfo.IDBancoPropio
    '        End If
    '    End If
    'End Sub

    '<Task()> Public Shared Sub AsignarClienteBancoPredeterminado(ByVal data As DataRow, ByVal services As ServiceProvider)
    '    If Length(data("IDClienteBanco")) = 0 Then
    '        Dim intPredeterminado As Integer = ProcessServer.ExecuteTask(Of String, Integer)(AddressOf ClienteBanco.GetBancoPredeterminado, data("IdCliente"), services)
    '        If intPredeterminado > 0 Then data("IDClienteBanco") = intPredeterminado
    '    End If
    'End Sub

    '<Task()> Public Shared Sub AsignarClienteDireccion(ByVal data As DataRow, ByVal services As ServiceProvider)
    '    Dim StDatosDirec As New ClienteDireccion.DataDirecEnvio(data("IdCliente"), enumcdTipoDireccion.cdDireccionGiro)
    '    Dim dtCB As DataTable = ProcessServer.ExecuteTask(Of ClienteDireccion.DataDirecEnvio, DataTable)(AddressOf ClienteDireccion.ObtenerDireccionEnvio, StDatosDirec, services)
    '    If Not IsNothing(dtCB) > 0 AndAlso dtCB.Rows.Count > 0 Then
    '        data("IDDireccion") = dtCB.Rows(0)("IDDireccion")
    '    End If
    'End Sub

    '<Serializable()> _
    'Public Class DataCopiaDataRow
    '    Public DtOrigen As DataTable
    '    Public DtDestino As DataTable
    'End Class

    '<Task()> Public Shared Sub AsignarDatosDesdeCobroOrigen(ByVal data As DataCopiaDataRow, ByVal services As ServiceProvider)
    '    data.DtDestino.Rows(0)("IdTipoCobro") = data.DtOrigen.Rows(0)("IdTipoCobro")
    '    data.DtDestino.Rows(0)("IdCliente") = data.DtOrigen.Rows(0)("IdCliente")
    '    data.DtDestino.Rows(0)("Titulo") = data.DtOrigen.Rows(0)("Titulo")
    '    data.DtDestino.Rows(0)("IDBancoPropio") = data.DtOrigen.Rows(0)("IDBancoPropio")
    '    If Length(data.DtOrigen.Rows(0)("NFactura")) > 0 Then data.DtDestino.Rows(0)("NFactura") = data.DtOrigen.Rows(0)("NFactura")
    '    If Length(data.DtOrigen.Rows(0)("IDFactura")) > 0 Then data.DtDestino.Rows(0)("IDFactura") = data.DtOrigen.Rows(0)("IDFactura")

    '    data.DtDestino.Rows(0)("CContable") = data.DtOrigen.Rows(0)("CContable")
    '    data.DtDestino.Rows(0)("IDFormaPago") = data.DtOrigen.Rows(0)("IDFormaPago")
    '    data.DtDestino.Rows(0)("Situacion") = data.DtOrigen.Rows(0)("Situacion")
    '    data.DtDestino.Rows(0)("FechaVencimiento") = data.DtOrigen.Rows(0)("FechaVencimiento")
    '    If Length(data.DtOrigen.Rows(0)("FechaVencimientoFactura")) > 0 Then
    '        data.DtDestino.Rows(0)("FechaVencimientoFactura") = data.DtOrigen.Rows(0)("FechaVencimientoFactura")
    '    Else
    '        data.DtDestino.Rows(0)("FechaVencimientoFactura") = data.DtOrigen.Rows(0)("FechaVencimiento")
    '    End If

    '    data.DtDestino.Rows(0)("IDMoneda") = data.DtOrigen.Rows(0)("IDMoneda")
    '    data.DtDestino.Rows(0)("CambioA") = data.DtOrigen.Rows(0)("CambioA")
    '    data.DtDestino.Rows(0)("CambioB") = data.DtOrigen.Rows(0)("CambioB")

    '    data.DtDestino.Rows(0)("ImpVencimiento") = data.DtOrigen.Rows(0)("ImpVencimiento")
    '    data.DtDestino.Rows(0)("IDClienteBanco") = data.DtOrigen.Rows(0)("IDClienteBanco")
    'End Sub

    <Task()> Public Shared Sub InsertarCobroDesdePago(ByVal IDPago As Integer, ByVal services As ServiceProvider)
        If IDPago > 0 Then
            Dim pa As New Pago
            Dim dtPago As DataTable = pa.SelOnPrimaryKey(IDPago)
            If Not dtPago Is Nothing AndAlso dtPago.Rows.Count > 0 Then
                AdminData.BeginTx()
                '//Generamos el cobro a partir del pago
                Dim c As New Cobro
                Dim dtCobro As DataTable = c.AddNewForm()
                Dim IDCobro As Integer = dtCobro.Rows(0)("IDCobro")
                Dim AppParams As ParametroTesoreria = services.GetService(Of ParametroTesoreria)()
                dtCobro.Rows(0)("IdTipoCobro") = AppParams.TipoCobroDesdePago
                dtCobro.Rows(0)("Titulo") = dtPago.Rows(0)("Titulo")
                If Length(dtPago.Rows(0)("NFactura")) > 0 Then
                    dtCobro.Rows(0)("NFactura") = dtPago.Rows(0)("NFactura")
                End If
                dtCobro.Rows(0)("IDFormaPago") = dtPago.Rows(0)("IDFormaPago")
                If Length(dtPago.Rows(0)("IDBancoPropio")) > 0 Then
                    dtCobro.Rows(0)("IDBancoPropio") = dtPago.Rows(0)("IDBancoPropio")
                End If
                dtCobro.Rows(0)("CContable") = dtPago.Rows(0)("CContable")
                dtCobro.Rows(0)("IDMoneda") = dtPago.Rows(0)("IDMoneda")
                dtCobro.Rows(0)("Situacion") = enumCobroSituacion.NoNegociado
                dtCobro.Rows(0)("FechaVencimiento") = dtPago.Rows(0)("FechaVencimiento")
                dtCobro.Rows(0)("ImpVencimiento") = -dtPago.Rows(0)("ImpVencimiento")
                dtCobro.Rows(0)("ImpVencimientoA") = -dtPago.Rows(0)("ImpVencimientoA")
                dtCobro.Rows(0)("ImpVencimientoB") = -dtPago.Rows(0)("ImpVencimientoB")
                dtCobro.Rows(0)("CambioA") = dtPago.Rows(0)("CambioA")
                dtCobro.Rows(0)("CambioB") = dtPago.Rows(0)("CambioB")
                dtCobro.Rows(0)("IDPago") = dtPago.Rows(0)("IDPago")
                c.Update(dtCobro)

                '//Asociar el cobro generado al pago origen
                dtPago.Rows(0)("IDCobro") = IDCobro
                dtPago.Rows(0)("Situacion") = enumPagoSituacion.GeneradoCobro
                pa.Update(dtPago)
                AdminData.CommitTx(True)
            End If
        End If
    End Sub

#End Region

#Region " Remesas "

    <Serializable()> _
    Public Class DataUpdateRemesa
        Public IDRemesa As Integer?
        Public Cobros As DataTable
        Public IDBancoPropio As String
        Public TipoNegociacion As Integer?
        Public FechaEmision As Date?
        Public FechaCargo As Date?
        Public Situacion As Integer?
        Public MaquinaUsuario As String
        Public Ruta As String

        Public IDContador As String
        Public FechaAbonoAnticipo As Date?
        Public FechaCancelacionAnticipo As Date?
        Public FechaCancelacionAnticipoPrev As Date?
        Public ReferenciaAbonoAnticipo As String
        Public TipoInteresInicial As Double?

    End Class

    <Task()> Public Shared Function UpdateRemesa(ByVal data As DataUpdateRemesa, ByVal services As ServiceProvider) As DataTable
        Dim C As New Cobro
        Dim dtCobrosActualizar As DataTable
        Dim r As New Remesa
        Dim DtRemesa As DataTable
        If data.IDRemesa Is Nothing OrElse data.IDRemesa = 0 Then
            Dim AppParams As ParametroTesoreria = services.GetService(Of ParametroTesoreria)()
            Dim StrIDContador As String
            If data.TipoNegociacion = enumTipoRemesa.RemesaAnticipo Then
                StrIDContador = data.IDContador
                If Length(StrIDContador) = 0 Then
                    StrIDContador = AppParams.ContadorRemesaAnticipo
                End If
            Else
                StrIDContador = AppParams.ContadorRemesa
            End If
            Dim IntIDRemesa As Integer

            If Length(StrIDContador) Then IntIDRemesa = CInt(ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, StrIDContador, services))
            DtRemesa = r.AddNewForm
            If Length(IntIDRemesa) <> 0 AndAlso IntIDRemesa <> 0 Then DtRemesa.Rows(0)("IdRemesa") = IntIDRemesa
            If Length(data.IDBancoPropio) <> 0 Then DtRemesa.Rows(0)("IDBancoPropio") = data.IDBancoPropio
            If Not data.TipoNegociacion Is Nothing Then DtRemesa.Rows(0)("IDTipoNegociacion") = data.TipoNegociacion
            If Length(data.FechaEmision) <> 0 AndAlso data.FechaEmision <> cnMinDate Then DtRemesa.Rows(0)("FechaNegociacion") = data.FechaEmision
            If Length(data.Ruta) <> 0 Then DtRemesa.Rows(0)("Ruta") = data.Ruta
            If Length(data.FechaCargo) <> 0 AndAlso data.FechaCargo <> cnMinDate Then DtRemesa.Rows(0)("FechaCargo") = data.FechaCargo

            Dim ImporteInicialAnticipos As Double = 0
            For Each Dr As DataRow In data.Cobros.Select
                Dim Dt As DataTable = C.SelOnPrimaryKey(Dr("IdCobro"))
                If Length(data.IDBancoPropio) <> 0 Then Dt.Rows(0)("IDBancoPropio") = data.IDBancoPropio
                If data.TipoNegociacion = enumTipoRemesa.RemesaAnticipo Then
                    If Len(IntIDRemesa & String.Empty) <> 0 Then Dt.Rows(0)("IdRemesaAnticipo") = IntIDRemesa
                    If Nz(data.FechaCancelacionAnticipo, cnMinDate) <> cnMinDate Then Dt.Rows(0)("FechaCancelacionAnticipo") = data.FechaCancelacionAnticipo
                    If Nz(data.FechaCancelacionAnticipoPrev, cnMinDate) <> cnMinDate Then Dt.Rows(0)("FechaCancelacionAnticipoPrev") = data.FechaCancelacionAnticipoPrev
                    Dt.Rows(0)("EstadoAnticipo") = enumEstadoAnticipo.PdteAbono
                    ImporteInicialAnticipos += Nz(Dt.Rows(0)("ImporteRemesaAnticipoA"), 0)
                Else
                    If Len(IntIDRemesa & String.Empty) <> 0 Then Dt.Rows(0)("IdRemesa") = IntIDRemesa
                End If

                If Length(data.Situacion) > 0 AndAlso data.Situacion <> -1 Then
                    Dt.Rows(0)("Situacion") = data.Situacion
                    Dim EstadosCobro As EntityInfoCache(Of EstadoCobroInfo) = services.GetService(Of EntityInfoCache(Of EstadoCobroInfo))()
                    Dim ECInfo As EstadoCobroInfo = EstadosCobro.GetEntity(data.Situacion)
                    If Not ECInfo Is Nothing AndAlso Length(ECInfo.IDEstado) > 0 Then
                        Dt.Rows(0)("IDAgrupacion") = ECInfo.IDAgrupacion
                    End If
                End If
                If dtCobrosActualizar Is Nothing Then dtCobrosActualizar = Dt.Clone
                dtCobrosActualizar.ImportRow(Dt.Rows(0))
            Next
            If Nz(data.TipoInteresInicial, 0) <> 0 Then DtRemesa.Rows(0)("TipoInteresInicial") = data.TipoInteresInicial
            DtRemesa.Rows(0)("ImporteInicial") = ImporteInicialAnticipos
        Else
            Dim IntIDRemesa As Integer = data.IDRemesa
            DtRemesa = r.Filter(New FilterItem("IDRemesa", FilterOperator.Equal, IntIDRemesa))
            If Not DtRemesa Is Nothing AndAlso DtRemesa.Rows.Count > 0 Then
                If Length(data.IDBancoPropio) <> 0 Then DtRemesa.Rows(0)("IDBancoPropio") = data.IDBancoPropio
                If Not data.TipoNegociacion Is Nothing Then DtRemesa.Rows(0)("IDTipoNegociacion") = data.TipoNegociacion
                If Length(data.FechaEmision) <> 0 AndAlso data.FechaEmision <> cnMinDate Then DtRemesa.Rows(0)("FechaNegociacion") = data.FechaEmision
                If Length(data.Ruta) <> 0 Then DtRemesa.Rows(0)("Ruta") = data.Ruta
                If Length(data.FechaCargo) <> 0 AndAlso data.FechaCargo <> cnMinDate Then DtRemesa.Rows(0)("FechaCargo") = data.FechaCargo
            End If
        End If

        AdminData.BeginTx()
        DtRemesa = r.Update(DtRemesa)
        If Not dtCobrosActualizar Is Nothing AndAlso dtCobrosActualizar.Rows.Count > 0 Then C.Update(dtCobrosActualizar)

        Return DtRemesa
    End Function

    <Task()> Public Shared Function NuevoIdRemesa(ByVal data As Object, ByVal services As ServiceProvider) As Integer
        Dim IntIdRemesa As Integer = -1
        Dim AppParams As ParametroTesoreria = services.GetService(Of ParametroTesoreria)()
        Dim StrIDContador As String = AppParams.ContadorRemesa
        Dim DtContador As DataTable = New Contador().SelOnPrimaryKey(StrIDContador)
        If Not DtContador Is Nothing AndAlso DtContador.Rows.Count > 0 Then
            If Length(DtContador.Rows(0)("Numerico")) > 0 Then
                IntIdRemesa = DtContador.Rows(0)("Contador")
            End If
        End If
        Return IntIdRemesa
    End Function

    <Task()> Public Shared Function NuevoIdRemesaAnticipo(ByVal IDContador As String, ByVal services As ServiceProvider) As Integer
        Dim IntIdRemesa As Integer = -1
        Dim AppParams As ParametroTesoreria = services.GetService(Of ParametroTesoreria)()
        If Length(IDContador) = 0 Then IDContador = AppParams.ContadorRemesaAnticipo
        Dim DtContador As DataTable = New Contador().SelOnPrimaryKey(IDContador)
        If DtContador.Rows.Count > 0 Then
            If Length(DtContador.Rows(0)("Numerico")) > 0 Then
                IntIdRemesa = DtContador.Rows(0)("Contador")
            End If
        End If
        Return IntIdRemesa
    End Function

#Region " Remesas - Eliminación/Descontabilización "

    <Serializable()> _
    Public Class DataDeleteRemesa
        Public IDRemesa As Integer
        Public TodosAsientos As Boolean
        Public IDEjercicio As String
        Public IDEjercicioTributario As String
        Public NAsiento As Integer
        Public TipoNegociacion As enumTipoRemesa
    End Class

    <Task()> Public Shared Function DeleteRemesa(ByVal data As DataDeleteRemesa, ByVal services As ServiceProvider) As Boolean
        '//Sólo dejaremos borrar si no tenemos liquidados, no dejaremos los que esten parcialmente liquidados  (Contabilidad Multiple)
        AdminData.BeginTx()
        Dim blnContabilidadMultiple As Boolean = New Parametro().ContabilidadMultiple
        Dim blnEsEjercicioTributario As Boolean
        Dim c As New Cobro
        Dim DtCobros As DataTable
        Dim IDCobros(-1) As Object
        If data.TodosAsientos Then
            '//Cuando en la pantalla de remesas no se encuentra ningún asiento vinculado a la remesa, que cumpla las condiciones de borrado.
            If data.TipoNegociacion = enumTipoRemesa.RemesaAnticipo Then
                DtCobros = c.Filter(, "EstadoAnticipo <> " & enumEstadoAnticipo.Cancelado & " AND IdRemesaAnticipo=" & data.IDRemesa)
            Else
                DtCobros = c.Filter(, "Liquidado=0 AND IdRemesa=" & data.IDRemesa)
            End If
        Else
            '//Cuando en la pantalla de remesas se encuentra algún asiento vinculado a la remesa, que cumpla las condiciones de borrado. 
            '//Se ejecutará cada asiento de manera independiente.
            '//en el caso de las NIIFs, vendrá por aquí con cada ejercicio y asiento contable
            Dim IDEjercicioBusq As String = data.IDEjercicio
            If blnContabilidadMultiple Then
                Dim f As New Filter
                f.Add(New StringFilterItem("IDEjercicioTributario", data.IDEjercicio))
                Dim dtEjercicio As DataTable = AdminData.GetData("tbMaestroEjercicio", f)
                If dtEjercicio.Rows.Count > 0 Then
                    IDEjercicioBusq = dtEjercicio.Rows(0)("IDEjercicio")
                    blnEsEjercicioTributario = True
                End If
            End If
            If data.TipoNegociacion = enumTipoRemesa.RemesaAnticipo Then
                DtCobros = c.Filter(, "IdCobro IN (Select IdDocumento From tbDiarioContable Where NAsiento=" & data.NAsiento & " AND IdEjercicio='" & IDEjercicioBusq & "') AND EstadoAnticipo <> " & enumEstadoAnticipo.Cancelado & " AND IdRemesaAnticipo=" & data.IDRemesa)
            Else
                DtCobros = c.Filter(, "IdCobro IN (Select IdDocumento From tbDiarioContable Where NAsiento=" & data.NAsiento & " AND IdEjercicio='" & IDEjercicioBusq & "') AND Liquidado=0 AND IdRemesa=" & data.IDRemesa)
            End If
        End If

        If data.TodosAsientos Then
            For Each Dr As DataRow In DtCobros.Select
                ReDim Preserve IDCobros(IDCobros.Length)
                IDCobros(IDCobros.Length - 1) = Dr("IDCobro")
            Next

            If IDCobros.Length > 0 Then
                Dim fDelete As New Filter
                fDelete.Add(New InListFilterItem("IDDocumento", IDCobros, FilterType.Numeric))
                If data.TipoNegociacion = enumTipoRemesa.RemesaAnticipo Then
                    fDelete.Add(New NumberFilterItem("IDTipoApunte", enumDiarioTipoApunte.CancelacionAnticipo))
                Else
                    fDelete.Add(New NumberFilterItem("IDTipoApunte", enumDiarioTipoApunte.LiquidacionRemesa))
                End If

                Dim fAux As New Filter
                fAux.Add(New StringFilterItem("IDEjercicio", data.IDEjercicio))
                fAux.Add(New IsNullFilterItem("MesCierre", False))
                fAux.Add(fDelete)
                Dim dtApuntes As DataTable = New BE.DataEngine().Filter("tbDiarioContable", fAux, "TOP 1 IDApunte, MesCierre")
                If dtApuntes.Rows.Count > 0 Then
                    ApplicationService.GenerateError("No se puede eliminar la Remesa, ya que el Periodo {0} está cerrado.", Quoted(dtApuntes.Rows(0)("MesCierre")))
                End If
                NegocioGeneral.DeleteWhere(data.IDEjercicio, fDelete)
            End If
        Else
            Dim fDelete As New Filter
            fDelete.Add(New NumberFilterItem("NAsiento", data.NAsiento))

            Dim fAux As New Filter
            fAux.Add(New StringFilterItem("IDEjercicio", data.IDEjercicio))
            fAux.Add(New IsNullFilterItem("MesCierre", False))
            fAux.Add(fDelete)
            Dim dtApuntes As DataTable = New BE.DataEngine().Filter("tbDiarioContable", fAux, "TOP 1 IDApunte, MesCierre")
            If dtApuntes.Rows.Count > 0 Then
                ApplicationService.GenerateError("No se puede eliminar la Remesa, ya que el Periodo {0} está cerrado.", Quoted(dtApuntes.Rows(0)("MesCierre")))
            End If
            NegocioGeneral.DeleteWhere(data.IDEjercicio, fDelete)
        End If

        For Each Dr As DataRow In DtCobros.Select
            If data.TipoNegociacion = enumTipoRemesa.RemesaAnticipo Then
                Dr("IdRemesaAnticipo") = DBNull.Value
                Dr("EstadoAnticipo") = DBNull.Value
                Dr("ReferenciaCancelacionAnticipo") = DBNull.Value
                Dr("FechaCancelacionAnticipo") = DBNull.Value
                Dr("FechaCancelacionAnticipoPrev") = DBNull.Value
                If Dr("Situacion") = enumCobroSituacion.Anticipado Then
                    'Dr("Situacion") = enumCobroSituacion.NoNegociado
                    If Length(Dr("IDRemesa")) = 0 Then
                        Dr("Situacion") = enumCobroSituacion.NoNegociado
                    Else
                        Dim f As New Filter
                        f.Add(New NumberFilterItem("IDRemesa", Dr("IDRemesa")))
                        Dim dtRemesa As DataTable = New Remesa().Filter(f)
                        If Not dtRemesa Is Nothing AndAlso dtRemesa.Rows.Count > 0 Then
                            If dtRemesa.Rows(0)("IDTipoNegociacion") = enumTipoRemesa.RemesaAlCobro Then 'Al Cobro
                                Dr("Situacion") = enumCobroSituacion.Negociado
                            Else 'If dtRemesa.Rows(0)("IDTipoNegociacion") = enumTipoRemesa.RemesaAlDescuento Then 'Al Descuento
                                Dr("Situacion") = enumCobroSituacion.Descontado
                            End If
                        End If
                    End If
                End If
            Else
                Dr("IDBancoPropio") = DBNull.Value
                Dr("IdRemesa") = DBNull.Value
                Dr("FechaCobro") = DBNull.Value
                ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf Pago.ActualizarFechaPagoDesdeCobro, New DataRowPropertyAccessor(Dr), services)

                Dim datValEst As New Comunes.DataValidarEstado(Dr("IDCobro"), enumDiarioTipoApunte.LiquidacionRemesa)
                datValEst.Descontabilizar = True
                Dr("Liquidado") = ProcessServer.ExecuteTask(Of Comunes.DataValidarEstado, Integer)(AddressOf Comunes.ValidarEstado, datValEst, services)
                'Dr("Liquidado") = enumCobroContabilizado.CobroNoContabilizado

                Dim datValEstConta As New Comunes.DataValidarEstado(Dr("IDCobro"), enumDiarioTipoApunte.Remesa)
                datValEstConta.Descontabilizar = True
                Dr("Contabilizado") = ProcessServer.ExecuteTask(Of Comunes.DataValidarEstado, Integer)(AddressOf Comunes.ValidarEstado, datValEstConta, services)
                'Dr("Situacion") = enumCobroSituacion.NoNegociado
                If Length(Dr("IDRemesaAnticipo")) = 0 Then
                    Dr("Situacion") = enumCobroSituacion.NoNegociado
                Else
                    If Dr("EstadoAnticipo") = enumEstadoAnticipo.Cancelado Then
                        Dr("Situacion") = enumCobroSituacion.NoNegociado
                    Else
                        Dr("Situacion") = enumCobroSituacion.Anticipado
                    End If
                End If

            End If
        Next
        c.Update(DtCobros)

        Dim fRemesa As New Filter
        If Not blnEsEjercicioTributario AndAlso DtCobros.Rows.Count = 0 Then
            If data.TipoNegociacion = enumTipoRemesa.RemesaAnticipo Then
                fRemesa.Add(New NumberFilterItem("IDRemesaAnticipo", data.IDRemesa))
            Else
                fRemesa.Add(New NumberFilterItem("IDRemesa", data.IDRemesa))
            End If
            Dim dtCobrosRemesa As DataTable = New Cobro().Filter(fRemesa)
            If dtCobrosRemesa.Rows.Count > 0 Then
                ApplicationService.GenerateError("La Remesa {0} tiene cobros asociados. No se puede eliminar. Revise sus datos.", Quoted(data.IDRemesa))
            End If
        End If

        fRemesa.Clear()
        fRemesa.Add(New NumberFilterItem("IDRemesa", data.IDRemesa))
        Dim dtGastosRemesa As DataTable = New RemesaCobroFacturaCompra().Filter(fRemesa)
        If dtGastosRemesa.Rows.Count > 0 Then
            ApplicationService.GenerateError("La Remesa {0} tiene gastos asociados. No se puede eliminar. Revise sus datos.", Quoted(data.IDRemesa))
        End If

        If Not blnEsEjercicioTributario Then AdminData.Execute("Delete From tbRemesa Where IdRemesa = " & data.IDRemesa)
        DeleteRemesa = True
    End Function

    <Serializable()> _
    Public Class DataDescontabilizarRemesa
        Public IDRemesa As Integer
        Public TodosAsientos As Boolean
        Public IDEjercicio As String
        Public NAsiento As Integer
        Public TipoNegociacion As enumTipoRemesa
    End Class

    <Task()> Public Shared Function DescontabilizarRemesa(ByVal data As DataDescontabilizarRemesa, ByVal services As ServiceProvider) As Boolean
        AdminData.BeginTx()
        Dim c As New Cobro
        Dim DtCobros As DataTable
        Dim IDCobros(-1) As Object
        Dim dtRemesa As DataTable

        If data.TodosAsientos Then
            If data.TipoNegociacion = enumTipoRemesa.RemesaAnticipo Then
                DtCobros = c.Filter(, "EstadoAnticipo <> " & enumEstadoAnticipo.Cancelado & " AND IdRemesaAnticipo=" & data.IDRemesa)
                dtRemesa = New Remesa().SelOnPrimaryKey(data.IDRemesa)
            Else
                DtCobros = c.Filter(, "Liquidado=0 AND IdRemesa=" & data.IDRemesa)
            End If
        Else
            If data.TipoNegociacion = enumTipoRemesa.RemesaAnticipo Then
                DtCobros = c.Filter(, "IdCobro IN (Select IdDocumento From tbDiarioContable Where NAsiento=" & data.NAsiento & " AND IdEjercicio='" & data.IDEjercicio & "') AND EstadoAnticipo <> " & enumEstadoAnticipo.Cancelado & " AND IdRemesaAnticipo=" & data.IDRemesa)
                dtRemesa = New Remesa().SelOnPrimaryKey(data.IDRemesa)
            Else
                DtCobros = c.Filter(, "IdCobro IN (Select IdDocumento From tbDiarioContable Where NAsiento=" & data.NAsiento & " AND IdEjercicio='" & data.IDEjercicio & "') AND Liquidado=0 AND IdRemesa=" & data.IDRemesa)
            End If
        End If

        For Each Dr As DataRow In DtCobros.Select
            ReDim Preserve IDCobros(IDCobros.Length)
            IDCobros(IDCobros.Length - 1) = Dr("IDCobro")
        Next

        If data.TodosAsientos Then
            If IDCobros.Length > 0 Then
                Dim fDelete As New Filter
                fDelete.Add(New InListFilterItem("IDDocumento", IDCobros, FilterType.Numeric))
                If data.TipoNegociacion = enumTipoRemesa.RemesaAnticipo Then
                    fDelete.Add(New NumberFilterItem("IDTipoApunte", enumDiarioTipoApunte.CancelacionAnticipo))
                Else
                    fDelete.Add(New NumberFilterItem("IDTipoApunte", enumDiarioTipoApunte.LiquidacionRemesa))
                End If

                Dim fAux As New Filter
                fAux.Add(New StringFilterItem("IDEjercicio", data.IDEjercicio))
                fAux.Add(New IsNullFilterItem("MesCierre", False))
                fAux.Add(fDelete)
                Dim dtApuntes As DataTable = New BE.DataEngine().Filter("tbDiarioContable", fAux, "TOP 1 IDApunte, MesCierre")
                If dtApuntes.Rows.Count > 0 Then
                    ApplicationService.GenerateError("No se puede eliminar la Remesa, ya que el Periodo {0} está cerrado.", Quoted(dtApuntes.Rows(0)("MesCierre")))
                End If

                NegocioGeneral.DeleteWhere(data.IDEjercicio, fDelete)
            End If
        Else
            Dim fDelete As New Filter
            fDelete.Add(New NumberFilterItem("NAsiento", data.NAsiento))

            Dim fAux As New Filter
            fAux.Add(New StringFilterItem("IDEjercicio", data.IDEjercicio))
            fAux.Add(New IsNullFilterItem("MesCierre", False))
            fAux.Add(fDelete)
            Dim dtApuntes As DataTable = New BE.DataEngine().Filter("tbDiarioContable", fAux, "TOP 1 IDApunte, MesCierre")
            If dtApuntes.Rows.Count > 0 Then
                ApplicationService.GenerateError("No se puede eliminar la Remesa, ya que el Periodo {0} está cerrado.", Quoted(dtApuntes.Rows(0)("MesCierre")))
            End If

            NegocioGeneral.DeleteWhere(data.IDEjercicio, fDelete)
        End If

        For Each Dr As DataRow In DtCobros.Select
            If data.TipoNegociacion = enumTipoRemesa.RemesaAnticipo Then
                Dr("EstadoAnticipo") = enumEstadoAnticipo.PdteAbono
                Dr("FechaCancelacionAnticipoPrev") = DBNull.Value
            Else
                'Dr("Liquidado") = enumContabilizado.NoContabilizado
                Dim datValEstLiq As New Comunes.DataValidarEstado(Dr("IDCobro"), enumDiarioTipoApunte.LiquidacionRemesa)
                datValEstLiq.Descontabilizar = True
                Dr("Liquidado") = ProcessServer.ExecuteTask(Of Comunes.DataValidarEstado, Integer)(AddressOf Comunes.ValidarEstado, datValEstLiq, services)

                Dim datValEstConta As New Comunes.DataValidarEstado(Dr("IDCobro"), enumDiarioTipoApunte.Remesa)
                datValEstConta.Descontabilizar = True
                Dr("Contabilizado") = ProcessServer.ExecuteTask(Of Comunes.DataValidarEstado, Integer)(AddressOf Comunes.ValidarEstado, datValEstConta, services)


                If Nz(Dr("Contabilizado"), enumContabilizado.NoContabilizado) = enumContabilizado.NoContabilizado AndAlso _
                    Nz(Dr("RecibidoEfecto"), enumContabilizado.NoContabilizado) = enumContabilizado.NoContabilizado Then
                    Dr("FechaCobro") = System.DBNull.Value
                    ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf Cobro.ActualizarFechaParaDeclaracionFactura, New DataRowPropertyAccessor(Dr), services)
                    ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf Pago.ActualizarFechaPagoDesdeCobro, New DataRowPropertyAccessor(Dr), services)
                End If
            End If
        Next
        c.Update(DtCobros)

        'If DtCobros.Rows.Count = 0 Then
        '    Dim fRemesa As New Filter
        '    If data.TipoNegociacion = enumTipoRemesa.RemesaAnticipo Then
        '        fRemesa.Add(New NumberFilterItem("IDRemesaAnticipo", data.IDRemesa))
        '    Else
        '        fRemesa.Add(New NumberFilterItem("IDRemesa", data.IDRemesa))
        '    End If
        '    Dim dtCobrosRemesa As DataTable = New Cobro().Filter(fRemesa)
        '    If dtCobrosRemesa.Rows.Count > 0 Then
        '        ApplicationService.GenerateError("La Remesa {0} tiene cobros asociados. Revise sus datos.", Quoted(data.IDRemesa))
        '    End If
        'End If

        If Not dtRemesa Is Nothing AndAlso dtRemesa.Rows.Count > 0 Then
            '//Sólo tendremos dtRemesa, cuando venimos de anticipos.
            dtRemesa.Rows(0)("IDEjercicioAnticipo") = System.DBNull.Value
            dtRemesa.Rows(0)("ContabilizadoAnticipo") = enumContabilizado.NoContabilizado
            BusinessHelper.UpdateTable(dtRemesa)


            Dim fGastoAnt As New Filter
            fGastoAnt.Add(New NumberFilterItem("IDRemesa", dtRemesa.Rows(0)("IDRemesa")))
            fGastoAnt.Add(New BooleanFilterItem("IncluirEnAsientoAnticipo", True))
            Dim dtGastosEnAsientoAnticipo As DataTable = New RemesaCobroFacturaCompra().Filter(fGastoAnt)
            If dtGastosEnAsientoAnticipo.Rows.Count > 0 Then
                Dim dtFC As DataTable = New FacturaCompraCabecera().SelOnPrimaryKey(dtGastosEnAsientoAnticipo.Rows(0)("IDFacturaCompra"))
                If dtFC.Rows.Count > 0 Then
                    If Nz(dtFC.Rows(0)("AñoDeclaracionIVA"), 0) <> 0 OrElse Nz(dtFC.Rows(0)("NDeclaracionIVA"), 0) <> 0 Then
                        ApplicationService.GenerateError("La Factura {0} asociada al gasto está declarada, por lo que no se puede descontabilizar el Anticipo.", Quoted(dtFC.Rows(0)("NFactura")))
                    End If
                    dtFC.Rows(0)("Estado") = dtRemesa.Rows(0)("ContabilizadoAnticipo")
                    BusinessHelper.UpdateTable(dtFC)
                End If
            End If
        End If
        DescontabilizarRemesa = True
    End Function

    <Serializable()> _
    Public Class DataDeleteLiquidacionRemesa
        Public IDRemesa As Integer
        Public TodosAsientos As Boolean
        Public IDEjercicio As String
        Public NAsiento As Integer
        'Public TipoNegociacion As enumTipoRemesa
    End Class

    <Task()> Public Shared Function DeleteLiquidacionRemesa(ByVal data As DataDeleteLiquidacionRemesa, ByVal services As ServiceProvider) As Boolean
        AdminData.BeginTx()

        Dim TipoAsientoRemesa As Integer = New Parametro().TipoAsientoRemesa
        Dim IntSituacion As New enumCobroSituacion
        Dim c As New Cobro
        Dim DtRemesa As DataTable = New Remesa().SelOnPrimaryKey(data.IDRemesa)
        If DtRemesa.Rows(0)("IDTipoNegociacion") = CInt(enumTipoRemesa.RemesaAlCobro) Then
            IntSituacion = enumCobroSituacion.Negociado
        ElseIf DtRemesa.Rows(0)("IDTipoNegociacion") = CInt(enumTipoRemesa.RemesaAlDescuento) Then
            IntSituacion = enumCobroSituacion.Descontado
        End If
        Dim IDCobros(-1) As Object
        Dim DtCobros As DataTable
        If data.TodosAsientos Then
            Dim FilCobro As New Filter
            If DtRemesa.Rows(0)("IDTipoNegociacion") <> CInt(enumTipoRemesa.RemesaAnticipo) Then
                FilCobro.Add("Liquidado", FilterOperator.NotEqual, enumContabilizado.NoContabilizado)
                FilCobro.Add("IDRemesa", FilterOperator.Equal, data.IDRemesa)
            End If
            DtCobros = c.Filter(FilCobro)
        Else
            If DtRemesa.Rows(0)("IDTipoNegociacion") <> CInt(enumTipoRemesa.RemesaAnticipo) Then
                DtCobros = c.Filter(, "IdCobro IN (Select IdDocumento From tbDiarioContable Where NAsiento=" & data.NAsiento & " AND IdEjercicio='" & data.IDEjercicio & "') AND Liquidado<>0 AND IdRemesa=" & data.IDRemesa)
            End If
        End If

        For Each Dr As DataRow In DtCobros.Select
            ReDim Preserve IDCobros(IDCobros.Length)
            IDCobros(IDCobros.Length - 1) = Dr("IDCobro")
        Next

        If data.TodosAsientos Then
            If IDCobros.Length > 0 Then
                Dim fDelete As New Filter
                fDelete.Add(New InListFilterItem("IDDocumento", IDCobros, FilterType.Numeric))
                If DtRemesa.Rows(0)("IDTipoNegociacion") <> CInt(enumTipoRemesa.RemesaAnticipo) Then
                    fDelete.Add(New NumberFilterItem("IDTipoApunte", enumDiarioTipoApunte.LiquidacionRemesa))
                End If

                Dim fAux As New Filter
                fAux.Add(New StringFilterItem("IDEjercicio", data.IDEjercicio))
                fAux.Add(New IsNullFilterItem("MesCierre", False))
                fAux.Add(fDelete)
                Dim dtApuntes As DataTable = New BE.DataEngine().Filter("tbDiarioContable", fAux, "TOP 1 IDApunte, MesCierre")
                If dtApuntes.Rows.Count > 0 Then
                    ApplicationService.GenerateError("No se puede eliminar la Remesa, ya que el Periodo {0} está cerrado.", Quoted(dtApuntes.Rows(0)("MesCierre")))
                End If

                NegocioGeneral.DeleteWhere(data.IDEjercicio, fDelete)
            End If
        Else
            Dim fDelete As New Filter
            fDelete.Add(New NumberFilterItem("NAsiento", data.NAsiento))

            Dim fAux As New Filter
            fAux.Add(New StringFilterItem("IDEjercicio", data.IDEjercicio))
            fAux.Add(New IsNullFilterItem("MesCierre", False))
            fAux.Add(fDelete)
            Dim dtApuntes As DataTable = New BE.DataEngine().Filter("tbDiarioContable", fAux, "TOP 1 IDApunte, MesCierre")
            If dtApuntes.Rows.Count > 0 Then
                ApplicationService.GenerateError("No se puede eliminar la Remesa, ya que el Periodo {0} está cerrado.", Quoted(dtApuntes.Rows(0)("MesCierre")))
            End If

            NegocioGeneral.DeleteWhere(data.IDEjercicio, fDelete)
        End If

        For Each Dr As DataRow In DtCobros.Select
            'ReDim Preserve IDCobros(IDCobros.Length)
            'IDCobros(IDCobros.Length - 1) = Dr("IDCobro")

            If DtRemesa.Rows(0)("IDTipoNegociacion") <> CInt(enumTipoRemesa.RemesaAnticipo) Then
                'Dr("Liquidado") = False
                Dim datValEstLiq As New Comunes.DataValidarEstado(Dr("IDCobro"), enumDiarioTipoApunte.LiquidacionRemesa)
                datValEstLiq.Descontabilizar = True
                Dr("Liquidado") = ProcessServer.ExecuteTask(Of Comunes.DataValidarEstado, Integer)(AddressOf Comunes.ValidarEstado, datValEstLiq, services)

                If Dr("Liquidado") = enumContabilizado.NoContabilizado Then Dr("Situacion") = IntSituacion
                If TipoAsientoRemesa = enumTipoAsientoRemesa.Banco_a_EfectoDto AndAlso DtRemesa.Rows(0)("IDTipoNegociacion") = CInt(enumTipoRemesa.RemesaAlDescuento) Then
                    Dr("FechaCobro") = System.DBNull.Value
                    ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf Pago.ActualizarFechaPagoDesdeCobro, New DataRowPropertyAccessor(Dr), services)
                End If

            End If
        Next
        c.Update(DtCobros)

        DeleteLiquidacionRemesa = True
    End Function


    <Serializable()> _
    Public Class DataDeleteLiquidacionRemesaAnticipo
        Public IDEjercicio As String
        Public NAsiento As Integer
    End Class

    <Task()> Public Shared Function DeleteLiquidacionRemesaAnticipo(ByVal data As DataDeleteLiquidacionRemesaAnticipo, ByVal services As ServiceProvider) As Boolean
        AdminData.BeginTx()

        If data Is Nothing Then Exit Function
        If Length(data.IDEjercicio) = 0 OrElse Length(data.NAsiento) = 0 Then
            ApplicationService.GenerateError("Debe indicar un Ejercicio y un asiento.")
        End If

        Dim IDCobrosObj(-1) As Object
        Dim BEDataEngine As New BE.DataEngine
        'data.IDCobros.CopyTo(IDCobrosObj, 0)

        Dim c As New Cobro
        Dim DtCobros As DataTable = c.Filter(, "IdCobro IN (Select IdDocumento From tbDiarioContable Where NAsiento=" & data.NAsiento & " AND IdEjercicio='" & data.IDEjercicio & "') AND EstadoAnticipo = " & enumEstadoAnticipo.Cancelado)
   
        For Each Dr As DataRow In DtCobros.Select
            ReDim Preserve IDCobrosObj(IDCobrosObj.Length)
            IDCobrosObj(IDCobrosObj.Length - 1) = Dr("IDCobro")
        Next

        If IDCobrosObj.Length > 0 Then
            Dim fDelete As New Filter
            fDelete.Add(New InListFilterItem("IDDocumento", IDCobrosObj, FilterType.Numeric))
            fDelete.Add(New NumberFilterItem("IDTipoApunte", enumDiarioTipoApunte.CancelacionAnticipo))

            Dim fAux As New Filter
            fAux.Add(New StringFilterItem("IDEjercicio", data.IDEjercicio))
            fAux.Add(New NumberFilterItem("NAsiento", data.NAsiento))
            fAux.Add(New IsNullFilterItem("MesCierre", False))
            fAux.Add(fDelete)
            Dim dtApuntes As DataTable = BEDataEngine.Filter("tbDiarioContable", fAux, "TOP 1 IDApunte, MesCierre")
            If dtApuntes.Rows.Count > 0 Then
                ApplicationService.GenerateError("No se puede la Cancelación de los Anticipos, ya que el Periodo {0} está cerrado.", Quoted(dtApuntes.Rows(0)("MesCierre")))
            End If

            NegocioGeneral.DeleteWhere(data.IDEjercicio, fDelete)
        End If


        For Each Dr As DataRow In DtCobros.Select
            'ReDim Preserve IDCobros(IDCobros.Length)
            'IDCobros(IDCobros.Length - 1) = Dr("IDCobro")

            Dim datValEstCancel As New Comunes.DataValidarEstado(Dr("IDCobro"), enumDiarioTipoApunte.CancelacionAnticipo)
            datValEstCancel.Descontabilizar = True
            Dim Liquidado As enumContabilizado = ProcessServer.ExecuteTask(Of Comunes.DataValidarEstado, Integer)(AddressOf Comunes.ValidarEstado, datValEstCancel, services)
            If Liquidado = enumContabilizado.NoContabilizado Then
                Dr("EstadoAnticipo") = enumEstadoAnticipo.Abonado
                Dr("FechaCancelacionAnticipo") = System.DBNull.Value
                Dr("ReferenciaCancelacionAnticipo") = System.DBNull.Value
                Dr("Situacion") = enumCobroSituacion.Anticipado
            End If

            If Length(Dr("IDEjercicio")) > 0 Then
                Dr("Contabilizado") = enumContabilizado.NoContabilizado

                Dim fCobroContabilizado As New Filter
                fCobroContabilizado.Add(New NumberFilterItem("IDDocumento", Dr("IDCobro")))
                fCobroContabilizado.Add(New NumberFilterItem("IDTipoApunte", enumDiarioTipoApunte.Cobro))
                fCobroContabilizado.Add(New StringFilterItem("IDEjercicio", Dr("IDEjercicio")))
                Dim dtApuntes As DataTable = BEDataEngine.Filter("tbDiarioContable", fCobroContabilizado, "TOP 1 IDApunte")
                If dtApuntes.Rows.Count > 0 Then
                    Dr("Contabilizado") = enumCobroContabilizado.CobroContabilizado
                    Dr("Situacion") = enumCobroSituacion.Cobrado
                End If
            End If
        Next
        c.Update(DtCobros)

        DeleteLiquidacionRemesaAnticipo = True
    End Function

#End Region

#End Region

#Region " Cobro Periodico "

    <Serializable()> _
    Public Class DataAddCobroPeriodico
        Public CobrosPeriodicos As DataTable
        Public FechaFinal As Date
        Public Simulacion As Boolean

        Public Sub New(ByVal CobrosPeriodicos As DataTable, ByVal FechaFinal As Date, Optional ByVal Simulacion As Boolean = False)
            Me.CobrosPeriodicos = CobrosPeriodicos
            Me.FechaFinal = FechaFinal
            Me.Simulacion = Simulacion
        End Sub
    End Class

    <Task()> Public Shared Function AddCobroPeriodico(ByVal data As DataAddCobroPeriodico, ByVal services As ServiceProvider) As DataTable
        If Not IsNothing(data.CobrosPeriodicos) AndAlso data.CobrosPeriodicos.Rows.Count > 0 Then
            Dim strUnidad, strAgrupacion As String
            Dim dtFechaComienzo, dtFechaTope As Date
            Dim c As New Cobro : Dim g As New NegocioGeneral
            Dim dtNewCobro As DataTable = c.AddNew()
            For Each dr As DataRow In data.CobrosPeriodicos.Select
                If Nz(dr("FechaUltimaActualizacion"), Date.MinValue) < dr("FechaFin") Then
                    strUnidad = g.GetPeriodString(dr("Unidad"))
                    If Length(dr("FechaUltimaActualizacion")) = 0 Then
                        dtFechaComienzo = DateAdd(strUnidad, -dr("Periodo"), dr("FechaInicio"))
                    Else
                        dtFechaComienzo = dr("FechaUltimaActualizacion")
                    End If

                    dtFechaTope = IIf(dr("FechaFin") < data.FechaFinal, dr("FechaFin"), data.FechaFinal)
                    strAgrupacion = dr("IDAgrupacion") & String.Empty
                    Dim intPeriodo As Integer = 1

                    Do While DateAdd(strUnidad, intPeriodo * dr("Periodo"), dtFechaComienzo) <= dtFechaTope
                        Dim drNewCobro As DataRow = dtNewCobro.NewRow

                        If Not data.Simulacion Then drNewCobro("IDCobro") = AdminData.GetAutoNumeric
                        drNewCobro("Titulo") = dr("DescCobro")
                        drNewCobro("FechaVencimiento") = DateAdd(strUnidad, intPeriodo * dr("Periodo"), dtFechaComienzo)
                        drNewCobro("CContable") = dr("IDCContable")
                        drNewCobro("IDCliente") = dr("IDCliente")
                        drNewCobro("IDClienteBanco") = dr("IDClienteBanco")
                        drNewCobro("IdTipoCobro") = dr("IdTipoCobro")
                        drNewCobro("IDFormaPago") = dr("IDFormaPago")
                        drNewCobro("IDBancoPropio") = dr("IDBancoPropio")
                        drNewCobro("IDMoneda") = dr("IDMoneda")
                        drNewCobro("CambioA") = dr("CambioA")
                        drNewCobro("CambioB") = dr("CambioB")
                        drNewCobro("Situacion") = enumCobroSituacion.NoNegociado
                        drNewCobro("Contabilizado") = enumCobroContabilizado.CobroNoContabilizado
                        drNewCobro("ImpVencimiento") = dr("Importe")
                        If Length(strAgrupacion) > 0 Then drNewCobro("IDAgrupacion") = strAgrupacion

                        Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(drNewCobro), dr("IDMoneda") & String.Empty, Nz(dr("CambioA"), 0), Nz(dr("CambioB"), 0))
                        ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)

                        If dr.Table.Columns.Contains("IDMandato") Then drNewCobro("IDMandato") = dr("IDMandato")

                        intPeriodo = intPeriodo + 1

                        'Actualización FechaUltimaActualizacion del Cobro periódico.
                        dr("FechaUltimaActualizacion") = drNewCobro("FechaVencimiento")




                        dtNewCobro.Rows.Add(drNewCobro)
                    Loop
                End If
            Next

            If Not data.Simulacion AndAlso Not IsNothing(dtNewCobro) AndAlso dtNewCobro.Rows.Count > 0 Then
                ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
                BusinessHelper.UpdateTable(dtNewCobro)
                BusinessHelper.UpdateTable(data.CobrosPeriodicos)
            End If
            Return dtNewCobro
        End If

    End Function

    'Private Function FechaMenor(ByVal dtF1 As Date, ByRef dtF2 As Date) As Date
    '    If dtF1 < dtF2 Then
    '        FechaMenor = dtF1
    '    Else
    '        FechaMenor = dtF2
    '    End If
    'End Function
#End Region

#Region " CambioSituaciónAutomatico "

    <Task()> Public Shared Function CambioSituacionAutomatico(ByVal IDCobros() As Integer, ByVal services As ServiceProvider) As Boolean
        If IDCobros Is Nothing OrElse IDCobros.Length = 0 Then ApplicationService.GenerateError("Debe indicar al menos un cobro.")

        Dim strIDAgrupacion As String
        Dim blnCambioSituacion As Boolean

        Dim c As New Cobro
        Dim TipoAsientoRemesa As enumTipoAsientoRemesa = New Parametro().TipoAsientoRemesa()
        Dim EstadosCobro As EntityInfoCache(Of EstadoCobroInfo) = services.GetService(Of EntityInfoCache(Of EstadoCobroInfo))()
        For Each IdCobro As Integer In IDCobros
            Dim dtCobro As DataTable = c.SelOnPrimaryKey(IdCobro)
            If dtCobro.Rows.Count > 0 Then
                Select Case dtCobro.Rows(0)("Situacion")
                    Case enumCobroSituacion.NoNegociado
                        dtCobro.Rows(0)("Situacion") = CInt(enumCobroSituacion.Vencido)
                        Dim ECInfo As EstadoCobroInfo = EstadosCobro.GetEntity(enumCobroSituacion.Vencido)
                        strIDAgrupacion = ECInfo.IDAgrupacion
                        If Len(strIDAgrupacion) > 0 Then dtCobro.Rows(0)("IDAgrupacion") = strIDAgrupacion
                    Case enumCobroSituacion.Vencido
                        dtCobro.Rows(0)("Situacion") = enumCobroSituacion.Impagado
                        Dim ECInfo As EstadoCobroInfo = EstadosCobro.GetEntity(enumCobroSituacion.Impagado)
                        strIDAgrupacion = ECInfo.IDAgrupacion
                        If Len(strIDAgrupacion) > 0 Then dtCobro.Rows(0)("IDAgrupacion") = strIDAgrupacion
                    Case enumCobroSituacion.Negociado, enumCobroSituacion.Descontado
                        If Nz(dtCobro.Rows(0)("Contabilizado"), False) AndAlso Not Nz(dtCobro.Rows(0)("Liquidado"), False) Then
                            Select Case TipoAsientoRemesa
                                Case enumTipoAsientoRemesa.Banco_a_Cliente
                                    dtCobro.Rows(0)("Situacion") = enumCobroSituacion.Cobrado
                                    Dim ECInfo As EstadoCobroInfo = EstadosCobro.GetEntity(enumCobroSituacion.Cobrado)
                                    strIDAgrupacion = ECInfo.IDAgrupacion
                                    If Len(strIDAgrupacion) > 0 Then dtCobro.Rows(0)("IDAgrupacion") = strIDAgrupacion

                                    'dtCobro.Rows(0)("FechaCobro") = dtCobro.Rows(0)("FechaVencimiento")
                                    'ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf Pago.ActualizarFechaPagoDesdeCobro, New DataRowPropertyAccessor(dtCobro.Rows(0)), services)
                                Case enumTipoAsientoRemesa.Banco_a_EfectoDto
                                    If dtCobro.Rows(0)("Situacion") = enumCobroSituacion.Descontado AndAlso Length(dtCobro.Rows(0)("IDRemesa")) > 0 Then
                                        blnCambioSituacion = True

                                        dtCobro.Rows(0)("FechaCobro") = dtCobro.Rows(0)("FechaVencimiento")
                                        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf Cobro.ActualizarFechaParaDeclaracionFactura, New DataRowPropertyAccessor(dtCobro.Rows(0)), services)
                                        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf Pago.ActualizarFechaPagoDesdeCobro, New DataRowPropertyAccessor(dtCobro.Rows(0)), services)
                                    ElseIf dtCobro.Rows(0)("Situacion") = enumCobroSituacion.Negociado AndAlso Length(dtCobro.Rows(0)("IDRemesa")) > 0 Then
                                        dtCobro.Rows(0)("Situacion") = enumCobroSituacion.Cobrado
                                    End If

                                Case enumTipoAsientoRemesa.Banco_a_EfectoDto_EfectoClte
                                    If Length(dtCobro.Rows(0)("IDRemesa")) > 0 Then
                                        blnCambioSituacion = True
                                    End If
                            End Select
                        End If
                End Select
            End If

            c.Update(dtCobro)
        Next

        Return blnCambioSituacion
    End Function

#End Region

    <Task()> Public Shared Function UpdateCobro(ByVal dt As DataTable, ByVal services As ServiceProvider) As Integer
        '//Prepara el DataTable para actualizar los cobros que se han modificado manualmente
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
            Dim MonInfoA As MonedaInfo = Monedas.MonedaA
            Dim MonInfoB As MonedaInfo = Monedas.MonedaB
            ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
            Dim c As New Cobro
            For Each drRow As DataRow In dt.Rows
                Dim MonInfo As MonedaInfo = Monedas.GetMoneda(drRow("IDMoneda"))
                If drRow.RowState = DataRowState.Modified Then
                    Dim dtCobroUpdate As DataTable = c.SelOnPrimaryKey(drRow("IDCobro"))
                    For Each drRowUpdate As DataRow In dtCobroUpdate.Select
                        drRowUpdate("IDFormaPago") = drRow("IDFormaPago")
                        drRowUpdate("FechaVencimiento") = drRow("FechaVencimiento")
                        drRowUpdate("ImpVencimiento") = xRound(drRow("ImpVencimiento"), MonInfo.NDecimalesImporte)
                        drRowUpdate("ImpVencimientoA") = xRound(drRow("ImpVencimientoA"), MonInfoA.NDecimalesImporte)
                        drRowUpdate("ImpVencimientoB") = xRound(drRow("ImpVencimientoB"), MonInfoB.NDecimalesImporte)
                        drRowUpdate("IDMoneda") = drRow("IDMoneda")
                        drRowUpdate("IDBancoPropio") = drRow("IDBancoPropio")
                        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ValidarCambioBancoPropio, drRowUpdate, services)
                        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ValidarCambioFormaPago, drRowUpdate, services)
                        If Length(drRow("IdDireccion")) > 0 Then
                            drRowUpdate("IdDireccion") = drRow("IdDireccion")
                        End If
                        If Length(drRow("IdClienteBanco")) > 0 Then
                            drRowUpdate("IdClienteBanco") = drRow("IdClienteBanco")
                        End If
                        drRowUpdate("Texto") = drRow("Texto")
                        drRowUpdate("CambioA") = drRow("CambioA")
                        drRowUpdate("CambioB") = drRow("CambioB")
                        drRowUpdate("CContable") = drRow("CContable")
                        'drRowUpdate("GeneradoAsientoTalon") = drRow("GeneradoAsientoTalon")
                        drRowUpdate("Titulo") = drRow("Titulo")
                        'drRowUpdate("Impreso") = drRow("Impreso")
                    Next drRowUpdate
                    c.Update(dtCobroUpdate)
                End If
            Next drRow
        End If

    End Function

    <Serializable()> _
    Public Class DataActualizarCobrosImpresos
        Public IDCobros() As Object
        Public Impreso As Boolean
    End Class
    <Task()> Public Shared Sub ActualizarCobrosImpresos(ByVal data As DataActualizarCobrosImpresos, ByVal services As ServiceProvider)
        If data.IDCobros Is Nothing OrElse data.IDCobros.Length = 0 Then Exit Sub
        Dim dtCobro As DataTable = New Cobro().Filter(New InListFilterItem("IDCobro", data.IDCobros, FilterType.Numeric))
        If Not dtCobro Is Nothing AndAlso dtCobro.Rows.Count > 0 Then
            For Each Dr As DataRow In dtCobro.Select
                Dr("Impreso") = data.Impreso
            Next
            BusinessHelper.UpdateTable(dtCobro)
        End If
    End Sub

    <Task()> Public Shared Sub DeleteCobroManual(ByVal IDCobros() As Object, ByVal services As ServiceProvider)
        If Not IDCobros Is Nothing AndAlso IDCobros.Length > 0 Then
            Dim fFilter As New Filter
            fFilter.Add("Contabilizado", enumCobroContabilizado.CobroNoContabilizado)
            fFilter.Add(New InListFilterItem("IDCobro", IDCobros, FilterType.Numeric))

            Dim c As New Cobro
            Dim dtCobros As DataTable = c.Filter(fFilter)
            If dtCobros.Rows.Count > 0 Then
                Dim dtCobrosDel As DataTable = dtCobros.Clone
                For Each Dr As DataRow In dtCobros.Select
                    If Length(Dr("IDFactura")) = 0 Then
                        Dim dtCobroAgrupado As DataTable = c.Filter(New NumberFilterItem("IDCobroAgrupado", Dr("IDCobro")))
                        If dtCobroAgrupado Is Nothing OrElse dtCobroAgrupado.Rows.Count = 0 Then
                            dtCobrosDel.ImportRow(Dr)
                        End If
                    End If
                Next
                c.Delete(dtCobrosDel)
            End If
        End If
    End Sub
    <Task()> Public Shared Function EsDesagrupable(ByVal intSituacion As enumCobroSituacion, ByVal services As ServiceProvider) As Boolean
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDEstado", FilterOperator.Equal, intSituacion))
        Dim dt As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf EstadoCobro.EstadosCobrosAgrupables, Nothing, services)
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            Dim WhereEstado As String = f.Compose(New AdoFilterComposer)
            Dim dr() As DataRow = dt.Select(WhereEstado)
            If dr.Length > 0 Then
                Return True
            Else : Return False
            End If
        End If
    End Function

    <Serializable()> _
    Public Class DataUpdateFechaVto
        Public IDCobros() As Object
        Public FechaVencimiento As Date
    End Class

    <Task()> Public Shared Function UpdateFechaVto(ByVal data As DataUpdateFechaVto, ByVal services As ServiceProvider) As Boolean
        Dim c As New Cobro
        Dim dtCobros As DataTable = c.Filter(New InListFilterItem("IDCobro", data.IDCobros, FilterType.Numeric))
        For Each Dr As DataRow In dtCobros.Select
            Dr("FechaVencimiento") = data.FechaVencimiento
        Next
        c.Update(dtCobros)
        Return True
    End Function

    <Serializable()> _
    Public Class DataInsertarVencimiento
        Public IDFactura As Integer
        Public ImpVencimiento As Double
        Public FechaVencimiento As Date
        Public IDFormaPago As String
        Public RecargoFinanciero As Double
    End Class

    <Task()> Public Shared Function InsertarVencimiento(ByVal data As DataInsertarVencimiento, ByVal services As ServiceProvider) As Integer
        InsertarVencimiento = -1
        Dim dtFactCab As DataTable = New FacturaVentaCabecera().SelOnPrimaryKey(data.IDFactura)
        Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
        Dim ClteInfo As ClienteInfo = Clientes.GetEntity(dtFactCab.Rows(0)("IDCliente"))
        If Not ClteInfo Is Nothing AndAlso Length(ClteInfo.IDCliente) > 0 Then
            Dim c As New Cobro
            Dim dtCobro As DataTable = c.AddNewForm
            InsertarVencimiento = dtCobro.Rows(0)("IdCobro")
            dtCobro.Rows(0)("IDFactura") = dtFactCab.Rows(0)("IDFactura")
            dtCobro.Rows(0)("NFactura") = dtFactCab.Rows(0)("NFactura")
            dtCobro.Rows(0)("IdCliente") = dtFactCab.Rows(0)("IdCliente")
            dtCobro.Rows(0)("IdClienteBanco") = dtFactCab.Rows(0)("IdClienteBanco")
            dtCobro.Rows(0)("IDBancoPropio") = dtFactCab.Rows(0)("IDBancoPropio")

            If Length(dtFactCab.Rows(0)("RazonSocial")) > 0 Then
                dtCobro.Rows(0)("Titulo").Value = dtFactCab.Rows(0)("RazonSocial")
            End If
            Dim AppParams As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
            If AppParams.Contabilidad Then
                If Length(ClteInfo.CCCliente) > 0 Then
                    dtCobro.Rows(0)("CContable") = ClteInfo.CCCliente
                Else
                    ApplicationService.GenerateError("La Cuenta Contable es un campo obligatorio.")
                End If
            End If
            dtCobro.Rows(0)("IDMoneda") = dtFactCab.Rows(0)("IDMoneda")
            dtCobro.Rows(0)("CambioA") = dtFactCab.Rows(0)("CambioA")
            dtCobro.Rows(0)("CambioB") = dtFactCab.Rows(0)("CambioB")
            Dim ValAyB As New ValoresAyB(data.ImpVencimiento, dtCobro.Rows(0)("IDMoneda"), dtCobro.Rows(0)("CambioA"), dtCobro.Rows(0)("CambioB"))
            Dim fImp As fImporte = ProcessServer.ExecuteTask(Of ValoresAyB, fImporte)(AddressOf NegocioGeneral.MantenimientoValoresImporteAyB, ValAyB, services)
            dtCobro.Rows(0)("ImpVencimiento") = fImp.Importe
            dtCobro.Rows(0)("ImpVencimientoA") = fImp.ImporteA
            dtCobro.Rows(0)("ImpVencimientoB") = fImp.ImporteB
            dtCobro.Rows(0)("FechaVencimiento") = data.ImpVencimiento
            dtCobro.Rows(0)("IDFormaPago") = data.IDFormaPago
            dtCobro.Rows(0)("Situacion") = enumCobroSituacion.NoNegociado
            dtCobro.Rows(0)("Contabilizado") = enumCobroContabilizado.CobroNoContabilizado
            dtCobro.Rows(0)("Impreso") = False
            ValAyB = New ValoresAyB(data.RecargoFinanciero, dtCobro.Rows(0)("IDMoneda"), dtCobro.Rows(0)("CambioA"), dtCobro.Rows(0)("CambioB"))
            fImp = ProcessServer.ExecuteTask(Of ValoresAyB, fImporte)(AddressOf NegocioGeneral.MantenimientoValoresImporteAyB, ValAyB, services)
            dtCobro.Rows(0)("RecargoFinanciero") = fImp.Importe
            dtCobro.Rows(0)("RecargoFinancieroA") = fImp.ImporteA
            dtCobro.Rows(0)("RecargoFinancieroB") = fImp.ImporteB
            dtCobro.Rows(0)("ARepercutirA") = 0
            dtCobro.Rows(0)("ARepercutirB") = 0
            c.Update(dtCobro)
        End If
    End Function

    <Task()> Public Shared Sub DeleteCobro(ByVal IDCobro As Integer, ByVal services As ServiceProvider)
        Dim C As New Cobro
        Dim dt As DataTable = C.SelOnPrimaryKey(IDCobro)
        If dt.Rows.Count > 0 Then C.Delete(dt)
    End Sub

    <Serializable()> _
    Public Class DataAjustarCobros
        Public CobrosModificados As DataTable
        Public CobrosEliminados As DataTable
    End Class
    <Task()> Public Shared Sub AjustarCobros(ByVal data As DataAjustarCobros, ByVal services As ServiceProvider)
        Dim c As New Cobro
        If Not data.CobrosModificados Is Nothing AndAlso data.CobrosModificados.Rows.Count > 0 Then
            Dim IDCobros(-1) As Object
            For Each Dr As DataRow In data.CobrosModificados.Select
                ReDim Preserve IDCobros(IDCobros.Length)
                IDCobros(IDCobros.Length - 1) = Dr("IDCobro")
            Next
            Dim dtCobros As DataTable = c.Filter(New InListFilterItem("IDCobro", IDCobros, FilterType.Numeric))
            If Not dtCobros Is Nothing AndAlso dtCobros.Rows.Count > 0 Then
                For Each drC As DataRow In dtCobros.Select
                    Dim drUp() As DataRow = data.CobrosModificados.Select("IDCobro = " & drC("IDCobro"))
                    If Not drUp Is Nothing AndAlso drUp.Length > 0 Then
                        Dim ValAyB As New ValoresAyB(CDbl(Nz(drUp(0)("ImporteNew"), 0)), drC("IDMoneda"), drC("CambioA"), drC("CambioB"))
                        Dim fImp As fImporte = ProcessServer.ExecuteTask(Of ValoresAyB, fImporte)(AddressOf NegocioGeneral.MantenimientoValoresImporteAyB, ValAyB, services)
                        drC("ImpVencimiento") = fImp.Importe
                        drC("ImpVencimientoA") = fImp.ImporteA
                        drC("ImpVencimientoB") = fImp.ImporteB
                    End If
                Next
                BusinessHelper.UpdateTable(dtCobros)
            End If
        End If
        If Not data.CobrosEliminados Is Nothing AndAlso data.CobrosEliminados.Rows.Count > 0 Then c.Delete(data.CobrosEliminados)
    End Sub

    <Serializable()> _
    Public Class DataInsertarDesgloseCobro
        Public IDCobroEliminar As Integer
        Public NuevosCobros As DataTable
    End Class
    <Task()> Public Shared Function InsertarDesgloseCobro(ByVal data As DataInsertarDesgloseCobro, ByVal services As ServiceProvider) As Integer
        ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
        Dim c As New Cobro
        Dim DtAux As DataTable = c.SelOnPrimaryKey(data.IDCobroEliminar)
        Dim dt As DataTable = c.AddNew
        Dim IDCobro As Integer
        For Each Dr As DataRow In data.NuevosCobros.Select
            Dim drNew As DataRow = dt.NewRow
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarIdentificador, drNew, services)
            If IDCobro = 0 Then
                IDCobro = drNew("IDCobro")
            End If

            For Each col As DataColumn In dt.Columns
                If col.ColumnName <> "IDCobro" Then
                    If DtAux.Columns.Contains(col.ColumnName) Then
                        drNew(col.ColumnName) = DtAux.Rows(0)(col.ColumnName)
                    End If
                End If
            Next
           
            drNew("IDFormaPago") = Dr("IDFormaPago")
            drNew("ImpVencimiento") = Dr("ImpVencimiento")
            drNew("FechaVencimiento") = Dr("FechaVencimiento")
            drNew("ImpVencimientoA") = Dr("ImpVencimientoA")
            drNew("ImpVencimientoB") = Dr("ImpVencimientoB")
            drNew("ARepercutir") = Dr("ARepercutir")
            drNew("ARepercutirA") = Dr("ARepercutirA")
            drNew("ARepercutirB") = Dr("ARepercutirB")
            dt.Rows.Add(drNew)
        Next
        c.Update(dt)

        Dim ClsCobroDev As New CobroDevolucion
        Dim DtCobroDev As DataTable = ClsCobroDev.Filter(New NumberFilterItem("IDCobro", data.IDCobroEliminar))
        If Not DtCobroDev Is Nothing AndAlso DtCobroDev.Rows.Count > 0 AndAlso dt.Rows.Count > 0 Then
            For Each Dr As DataRow In DtCobroDev.Select
                Dr("IDCobro") = IDCobro
            Next
            ClsCobroDev.Update(DtCobroDev)
        End If

        If Not DtAux Is Nothing AndAlso DtAux.Rows.Count > 0 Then c.Delete(DtAux.Rows(0))
    End Function

#Region " FACTORING "

    <Serializable()> _
    Public Class DataActualizarFactoring
        Public IDProcess As Guid
        Public NFactoring As String
        Public NuevaSituacion As Integer
        Public IDBancoPropio As String
    End Class

    <Task()> Public Shared Sub ActualizarSituacionFactoring(ByVal data As DataActualizarFactoring, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New GuidFilterItem("IDProcess", FilterOperator.Equal, data.IDProcess))
        f.Add(New NumberFilterItem("ImpVencimiento", FilterOperator.GreaterThan, 0))
        f.Add(New BooleanFilterItem("Factoring", True))
        f.Add(New IsNullFilterItem("IDFactura", False))
        Dim c As New Cobro
        Dim dtActualizarSituacion As DataTable = New BE.DataEngine().Filter("NegFactoringActualizarCobros", f)
        If Not IsNothing(dtActualizarSituacion) AndAlso dtActualizarSituacion.Rows.Count > 0 Then
            For Each drRowSituacion As DataRow In dtActualizarSituacion.Select
                If Length(data.NFactoring) > 0 Then drRowSituacion("NFactoring") = data.NFactoring
                drRowSituacion("Situacion") = data.NuevaSituacion
                drRowSituacion("IDBancoPropio") = data.IDBancoPropio
            Next drRowSituacion

            c.Update(dtActualizarSituacion)
        End If

    End Sub

    <Task()> Public Shared Sub ActualizarContadorFactoring(ByVal data As DataActualizarFactoring, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New GuidFilterItem("IDProcess", FilterOperator.Equal, data.IDProcess))
        f.Add(New NumberFilterItem("ImpVencimiento", FilterOperator.GreaterThan, 0))
        f.Add(New BooleanFilterItem("Factoring", True))
        f.Add(New IsNullFilterItem("IDFactura", False))
        Dim c As New Cobro
        Dim dtActualizarSituacion As DataTable = New BE.DataEngine().Filter("NegFactoringActualizarCobros", f)
        If Not IsNothing(dtActualizarSituacion) AndAlso dtActualizarSituacion.Rows.Count > 0 Then
            For Each drRowSituacion As DataRow In dtActualizarSituacion.Select
                If Length(data.NFactoring) > 0 Then drRowSituacion("NFactoring") = data.NFactoring
            Next drRowSituacion

            c.Update(dtActualizarSituacion)
        End If
    End Sub

#End Region

    <Serializable()> _
    Public Class dataActualizarFormaPago
        Public IDCobro() As Object
        Public IDformaPago As String

        Public Sub New(ByVal IDCobro() As Object, ByVal IDformaPago As String)
            Me.IDCobro = IDCobro
            Me.IDformaPago = IDformaPago
        End Sub
    End Class
    <Task()> Public Shared Sub ActualizarFormaPago(ByVal data As dataActualizarFormaPago, ByVal services As ServiceProvider)
        Dim dtCobros As DataTable = New Cobro().Filter(New InListFilterItem("IDCobro", data.IDCobro, FilterType.Numeric))
        If dtCobros.Rows.Count > 0 Then
            For Each drCobro As DataRow In dtCobros.Rows
                drCobro("IDformaPago") = data.IDformaPago
            Next
            Cobro.UpdateTable(dtCobros)
        End If
    End Sub

    <Serializable()> _
    Public Class dataCrearDesgloseCobro
        Public IDCobro As Integer
        Public Importe As Double

        Public Sub New(ByVal IDCobro As Integer, ByVal Importe As Double)
            Me.IDCobro = IDCobro
            Me.Importe = Importe
        End Sub
    End Class
    <Task()> Public Shared Function CrearDesgloseCobro(ByVal data As dataCrearDesgloseCobro, ByVal services As ServiceProvider) As Integer
        Dim dtCobro As DataTable = New Cobro().SelOnPrimaryKey(data.IDCobro)
        If dtCobro.Rows.Count > 0 Then
            Dim FechaVencimiento As Date = dtCobro.Rows(0)("FechaVencimiento")

            Dim datMon As New Moneda.DatosObtenerMoneda
            datMon.IDMoneda = dtCobro.Rows(0)("IDMoneda")
            datMon.Fecha = FechaVencimiento

            Dim MonInfo As MonedaInfo = ProcessServer.ExecuteTask(Of Moneda.DatosObtenerMoneda, MonedaInfo)(AddressOf Moneda.ObtenerMoneda, datMon, services)
            Dim MonInfoA As MonedaInfo = ProcessServer.ExecuteTask(Of Date, MonedaInfo)(AddressOf Moneda.MonedaA, FechaVencimiento, services)
            Dim MonInfoB As MonedaInfo = ProcessServer.ExecuteTask(Of Date, MonedaInfo)(AddressOf Moneda.MonedaB, FechaVencimiento, services)

            dtCobro.Rows(0)("ImpVencimiento") = xRound(dtCobro.Rows(0)("ImpVencimiento") - data.Importe, MonInfo.NDecimalesImporte)
            dtCobro.Rows(0)("ImpVencimientoA") = xRound(dtCobro.Rows(0)("ImpVencimiento") * dtCobro.Rows(0)("CambioA"), MonInfoA.NDecimalesImporte)
            dtCobro.Rows(0)("ImpVencimientoB") = xRound(dtCobro.Rows(0)("ImpVencimiento") * dtCobro.Rows(0)("CambioB"), MonInfoB.NDecimalesImporte)

            Dim drCobro As DataRow = dtCobro.NewRow
            For Each col As DataColumn In dtCobro.Columns
                If col.ColumnName <> "IDCobro" AndAlso dtCobro.Columns.Contains(col.ColumnName) Then
                    drCobro(col.ColumnName) = dtCobro.Rows(0)(col.ColumnName)
                End If
            Next

            drCobro("IDCobro") = AdminData.GetAutoNumeric
            drCobro("ImpVencimiento") = xRound(data.Importe, MonInfo.NDecimalesImporte)
            drCobro("ImpVencimientoA") = xRound(drCobro("ImpVencimiento") * dtCobro.Rows(0)("CambioA"), MonInfoA.NDecimalesImporte)
            drCobro("ImpVencimientoB") = xRound(drCobro("ImpVencimiento") * dtCobro.Rows(0)("CambioB"), MonInfoB.NDecimalesImporte)
            dtCobro.Rows.Add(drCobro)

            Cobro.UpdateTable(dtCobro)
            Return drCobro("IDCobro")
        End If
    End Function


#Region " Riesgo Asegurado Plazo "

    <Serializable()> _
    Public Class DataGetRiesgoAseguradoPlazo
        Public FechaCalculo As Date
        Public Filtro As Filter
        Public IDBaseDatos As Guid
        Public Multiempresa As Boolean

        Public Sub New(ByVal FechaCalculo As Date, ByVal Filtro As Filter, ByVal Multiempresa As Boolean, ByVal IDBaseDatos As Guid)
            Me.FechaCalculo = FechaCalculo
            Me.Filtro = Filtro
            Me.Multiempresa = Multiempresa
            Me.IDBaseDatos = IDBaseDatos
        End Sub
    End Class

    <Task()> Public Shared Function GetRiesgoAseguradoPlazo(ByVal data As DataGetRiesgoAseguradoPlazo, ByVal services As ServiceProvider) As DataTable
        Dim Esquema As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf Comunes.GetEsquemaBD, Nothing, services)
        Dim Sql As String = "SELECT IDClienteProveedor AS IDCliente, Titulo, CCClienteProveedor AS CCCliente, " _
                                    & "IDFactura, NFactura, FechaVencimiento, ImpVencimiento, ImpVencimientoA, ImpVencimientoB, " _
                                    & "FechaContabilizacionFactura, Fecha, Especial, IDCobro, CLTE.RefAseguradora, " _
                                    & "Esquema.fRiesgoConcedidoFecha(CLTE.CifCliente, '" & Format(data.FechaCalculo, "yyyyMMdd") & "') AS RiesgoConcedido, " _
                                    & "CAST(DATEDIFF(d,FechaVencimiento,'" & Format(data.FechaCalculo, "yyyyMMdd") & "') AS numeric(23,0)) AS Plazo " _
                                & "FROM CtlCIComparacionSaldoCobroFechaDesglose " _
                                    & "RIGHT OUTER JOIN tbMaestroCliente CLTE ON CtlCIComparacionSaldoCobroFechaDesglose.IDClienteProveedor = CLTE.IDCliente "

        Dim fEmpresasOrigen As New Filter
        If Not CType(data.IDBaseDatos, Guid).Equals(Guid.Empty) Then fEmpresasOrigen.Add(New GuidFilterItem("IDBaseDatos", data.IDBaseDatos))

        Dim QueryMultipleDB As New BEGetQueryMultipleDB
        Dim dataQueryMultipleDB As New BEGetQueryMultipleDB.DataGetQueryMultipleDB(Sql, data.Filtro, fEmpresasOrigen, data.Multiempresa)
        Return QueryMultipleDB.Execute(dataQueryMultipleDB)
    End Function
#End Region

#Region " Declaración IVA Caja "
    <Task()> Public Shared Sub ActualizarFechaParaDeclaracionFactura(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        If data.ContainsKey("IDCliente") AndAlso Length(data("IDCliente")) > 0 Then
            Dim AppParams As ParametroVenta = services.GetService(Of ParametroVenta)()
            Dim AppParamsGral As ParametroGeneral = services.GetService(Of ParametroGeneral)()
            If AppParams.IvaCajaCircuitoVentas Then
                If data.ContainsKey("IDFactura") AndAlso Length(data("IDFactura")) > 0 Then
                    Dim fFraNoDeclarada As New Filter
                    fFraNoDeclarada.Add(New NumberFilterItem("IDFactura", data("IDFactura")))
                    fFraNoDeclarada.Add(New IsNullFilterItem("NDeclaracionIVA", True))
                    fFraNoDeclarada.Add(New BooleanFilterItem("FechaDeclaracionManual", False))
                    Dim dtFVC As DataTable = New FacturaVentaCabecera().Filter(fFraNoDeclarada)
                    If dtFVC.Rows.Count > 0 Then
                        dtFVC.Rows(0)("FechaParaDeclaracion") = Nz(data("FechaCobro"), New Date(Year(dtFVC.Rows(0)("FechaFactura")) + 1, 12, 31)) 'NegocioGeneral.cnMAX_DATE)
                    End If
                    AdminData.SetData(dtFVC)
                ElseIf data.ContainsKey("IDFactura") AndAlso Length(data("IDFactura")) = 0 AndAlso Length(data("NFactura")) > 0 AndAlso data("NFactura") = AppParamsGral.NFacturaCobroAgrupado Then
                    Dim dtFrasCobroAgrupado As DataTable = New Cobro().Filter(New NumberFilterItem("IDCobroAgrupado", data("IDCobro")), Nothing, "IDFactura")
                    If dtFrasCobroAgrupado.Rows.Count > 0 Then
                        Dim IDFacturas() As Object = (From c In dtFrasCobroAgrupado Where Not c.IsNull("IDFactura") Select c("IDFactura") Distinct).ToArray
                        If Not IDFacturas Is Nothing AndAlso IDFacturas.Count > 0 Then
                            Dim fFraNoDeclarada As New Filter
                            fFraNoDeclarada.Add(New InListFilterItem("IDFactura", IDFacturas, FilterType.Numeric))
                            fFraNoDeclarada.Add(New IsNullFilterItem("NDeclaracionIVA", True))
                            fFraNoDeclarada.Add(New BooleanFilterItem("FechaDeclaracionManual", False))
                            Dim dtFVC As DataTable = New FacturaVentaCabecera().Filter(fFraNoDeclarada)
                            If dtFVC.Rows.Count > 0 Then
                                For Each drFra As DataRow In dtFVC.Rows
                                    drFra("FechaParaDeclaracion") = Nz(data("FechaCobro"), New Date(Year(drFra("FechaFactura")) + 1, 12, 31)) 'NegocioGeneral.cnMAX_DATE)
                                Next
                            End If
                            AdminData.SetData(dtFVC)
                        End If
                    End If
                End If
            End If
        End If
    End Sub

#End Region

#Region " Seguimiento del Cobro "

    <Task()> Public Shared Function GetAsientosCobro(ByVal IDCobro As Integer, ByVal services As ServiceProvider) As DataTable
        Dim f As New Filter
        f.Add(New FilterItem("IDDocumento", IDCobro))
        Dim fTiposApunte As New Filter(FilterUnionOperator.Or)
        fTiposApunte.Add(New FilterItem("IDTipoApunte", enumDiarioTipoApunte.Cobro))
        fTiposApunte.Add(New FilterItem("IDTipoApunte", enumDiarioTipoApunte.Remesa))
        fTiposApunte.Add(New FilterItem("IDTipoApunte", enumDiarioTipoApunte.DevolucionRemesa))
        fTiposApunte.Add(New FilterItem("IDTipoApunte", enumDiarioTipoApunte.LiquidacionRemesa))
        fTiposApunte.Add(New FilterItem("IDTipoApunte", enumDiarioTipoApunte.ProtestoCheque))
        fTiposApunte.Add(New FilterItem("IDTipoApunte", enumDiarioTipoApunte.RecibidoEfectoClte))
        fTiposApunte.Add(New FilterItem("IDTipoApunte", enumDiarioTipoApunte.Anticipo))
        fTiposApunte.Add(New FilterItem("IDTipoApunte", enumDiarioTipoApunte.CancelacionAnticipo))
        f.Add(fTiposApunte)

        Dim fAsiento As New Filter '


        Dim BEDataEngine As New BE.DataEngine
        Dim dtApuntesCobro As DataTable = BEDataEngine.Filter("tbDiarioContable", f, "IDEjercicio,NAsiento")
        If dtApuntesCobro.Rows.Count > 0 Then
            f.Clear()
            f.UnionOperator = FilterUnionOperator.Or

            For Each dr As DataRow In dtApuntesCobro.Rows
                fAsiento = New Filter
                fAsiento.Add(New FilterItem("IDEjercicio", dr("IDEjercicio")))
                fAsiento.Add(New FilterItem("NAsiento",  dr("NAsiento")))
                f.Add(fAsiento)
            Next

            Dim dtAsientosCobro As DataTable = BEDataEngine.Filter("tbDiarioContable", f)
            Return dtAsientosCobro
        End If

    End Function

#End Region

End Class