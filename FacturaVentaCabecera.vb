Imports Solmicro.Expertis.Business.Negocio.NegocioGeneral

Public Class FacturaVentaCabecera
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbFacturaVentaCabecera"

#Region " RegisterAddNewTasks "

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoComunes.AsignarIdentificadorFactura, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.AsignarCentroGestion, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarContador, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarNumeroFacturaProvisional, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarDatosDeclaraciones, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoComunes.AsignarFechaFactura, data, services)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoComunes.AsignarEjercicioContableFactura, New DataRowPropertyAccessor(data), services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoComunes.AsignarFechaParaDeclaracion, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarTipoFactura, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarEstadoFactura, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarArqueoCaja, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.AsignarMarcaIVAManual, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.AsignarMarcaVtosManuales, data, services)
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim CE As New CentroEntidad
        CE.CentroGestion = data("IDCentroGestion") & String.Empty
        CE.ContadorEntidad = CentroGestion.ContadorEntidad.FacturaVenta
        data("IDContador") = ProcessServer.ExecuteTask(Of CentroEntidad, String)(AddressOf CentroGestion.GetContadorPredeterminado, CE, services)
    End Sub

    <Task()> Public Shared Sub AsignarNumeroFacturaProvisional(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDContador")) > 0 Then
            Dim dtContadores As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDt, GetType(FacturaVentaCabecera).Name, services)
            Dim adr As DataRow() = dtContadores.Select("IDContador = " & Quoted(data("IDContador")))
            If Not IsNothing(adr) AndAlso adr.Length > 0 Then
                data("NFactura") = adr(0)("ValorProvisional")
            Else
                'Si no está bien configurado el Contador de Facturas de Venta en el Centro de Gestión,
                'cogemos el Contador por defecto de la entidad Factura Venta Cabecera.
                Dim dtContadorPred As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDefault, GetType(FacturaVentaCabecera).Name, services)
                If Not dtContadorPred Is Nothing AndAlso dtContadorPred.Rows.Count > 0 Then
                    data("IDContador") = dtContadorPred.Rows(0)("IDContador")
                    adr = dtContadores.Select("IDContador = " & Quoted(data("IDContador")))
                    If Not IsNothing(adr) AndAlso adr.Length > 0 Then
                        data("NFactura") = adr(0)("ValorProvisional")
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDatosDeclaraciones(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("Enviar347") = False : data("Enviar349") = False
        data("Servicios349") = False
    End Sub

    <Task()> Public Shared Sub AsignarTipoFactura(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("TipoFactura") = enumTipoFactura.tfNormal
    End Sub

    <Task()> Public Shared Sub AsignarEstadoFactura(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("Estado") = CInt(enumfvcEstado.fvcNoContabilizado)
    End Sub
    <Task()> Public Shared Sub AsignarArqueoCaja(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("Arqueo") = False
    End Sub
#End Region

#Region " Delete "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)

        deleteProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.BeginTransaction)
        deleteProcess.AddTask(Of DataRow)(AddressOf ProcesoFacturacionVenta.ValidarFacturaContabilizada)
        deleteProcess.AddTask(Of DataRow)(AddressOf ProcesoFacturacionVenta.ValidarFacturaDeclarada)
        deleteProcess.AddTask(Of DataRow)(AddressOf ProcesoFacturacionVenta.ValidarFacturaArqueoCaja)
        deleteProcess.AddTask(Of DataRow)(AddressOf ValidarDelFVentaControlSII)
        deleteProcess.AddTask(Of DataRow)(AddressOf ProcesoFacturacionVenta.ActualizarEntregasACuenta)
        deleteProcess.AddTask(Of DataRow)(AddressOf ProcesoFacturacionVenta.ActualizarPromociones)
        deleteProcess.AddTask(Of DataRow)(AddressOf ProcesoFacturacionVenta.EliminarCobrosCompensacionesOServicio)
    End Sub

    <Task()> Public Shared Sub ValidarDelFVentaControlSII(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter(FilterUnionOperator.Or)
        f.Add(New NumberFilterItem("IDFacturaVenta", data("IDFactura")))
        f.Add(New NumberFilterItem("IDFacturaCobro", data("IDFactura")))
        f.Add(New NumberFilterItem("IDFacturaIntracomunitaria", data("IDFactura")))
        Dim dtSIIControl As DataTable
        Try
            dtSIIControl = New BE.DataEngine().Filter("tbSIIControlEnvioLinea", f)
        Catch ex As Exception
            '///Para que no de error en los que no tienen las tablas del SII
        End Try

        If Not dtSIIControl Is Nothing AndAlso dtSIIControl.Rows.Count > 0 Then
            Dim EstadoRegistroRechazado As Integer = 2  '//Incorrecto
            Dim AceptadasSII As List(Of DataRow) = (From c In dtSIIControl Where Not c.IsNull("EstadoRegistro") AndAlso c("EstadoRegistro") <> EstadoRegistroRechazado Select c).ToList
            If Not AceptadasSII Is Nothing AndAlso AceptadasSII.Count > 0 Then
                ApplicationService.GenerateError("No se puede eliminar la Factura. Ha sido enviada y Aceptada en la AEAT. (SII)")
            End If

            Dim RechazadasSII As List(Of DataRow) = (From c In dtSIIControl Where Not c.IsNull("EstadoRegistro") = EstadoRegistroRechazado Select c).ToList
            If dtSIIControl.Rows.Count = RechazadasSII.Count Then
                Dim fIDDetalle As New Filter(FilterUnionOperator.Or)
                For Each dr As DataRow In dtSIIControl.Rows
                    fIDDetalle.Add(New FilterItem("IDSIIControlEnvioLinea", dr("IDSIIControlEnvioLinea")))
                Next
                If fIDDetalle.Count > 0 Then
                    Dim strWhere As String = AdminData.ComposeFilter(fIDDetalle)
                    Dim SQLDetalle As String = "DELETE FROM tbSIIControlEnvioLineaDetalle WHERE " & strWhere
                    AdminData.Execute(SQLDetalle)
                    Dim SQLLinea As String = "DELETE FROM tbSIIControlEnvioLinea WHERE " & strWhere
                    AdminData.Execute(SQLLinea)
                End If

            End If
        End If
    End Sub

#End Region

#Region " Update "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of UpdatePackage, DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CrearDocumento)
        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarNumeroFactura)
        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarDatosFiscales)
        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComunes.AsignarCentroGestion)
        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComunes.ActualizarCambiosMoneda)
        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.TratarPromocionesLineas)
        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularImporteLineasFacturas)
        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf NegocioGeneral.CalcularImpuestos)
        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularRepresentantes)
        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularAnalitica)
        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularBasesImponibles)
        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularTotales)
        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularPuntoVerde)
        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularVencimientos)
        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.AsignarClaveOperacion)
'        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoComunes.ValidarIVASDocFV)
        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf Business.General.Comunes.BeginTransaction)
        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.ActualizarAlbaran)
        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf Business.General.Comunes.UpdateDocument)
        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf Business.General.Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionObras.ActualizarObras)
        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.ActualizarOTs)
        updateProcess.AddTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.ActualizarQLineasPromociones)
    End Sub

    <Task()> Public Shared Sub CambiarEstadoFactura(ByVal data As DataTable, ByVal services As ServiceProvider)
        If data Is Nothing OrElse data.Rows.Count = 0 Then Exit Sub
        If data.Rows(0)("Estado") = enumfvcEstado.fvcContabilizado Then
            data.Rows(0)("Estado") = enumfvcEstado.fvcNoContabilizado
        Else
            data.Rows(0)("Estado") = enumfvcEstado.fvcContabilizado
        End If
        data.TableName = GetType(FacturaVentaCabecera).Name
        BusinessHelper.UpdateTable(data)
    End Sub


#End Region

#Region " Validate "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarClienteObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFacturaObligatoria)
        'validateProcess.AddTask(Of DataRow)(AddressOf ProcesoFacturacionVenta.ValidarNumeroFactura)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarCIFObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarClienteBloqueado)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoFacturacionVenta.ValidarFacturaContabilizada)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFechaFacturaObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarEjercicioContableFactura)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarMonedaObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFormaPagoObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarCondicionPagoObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarClaveOperacion)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarDtoProntoPagoRecFinan)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarFechaFacturaAnterior)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFechaDeclaracion)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarEnvio347)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarEnvio349)
    End Sub

    <Task()> Public Shared Sub ValidarFechaFacturaAnterior(ByVal data As DataRow, ByVal services As ServiceProvider)
        If New Parametro().ValidarCambioFechaFacturas Then
            Dim FilFacturas As New Filter
            FilFacturas.Add("IDFactura", FilterOperator.NotEqual, data("IDFactura"))
            FilFacturas.Add("IDContador", FilterOperator.Equal, data("IDContador"))
            FilFacturas.Add("IDEjercicio", FilterOperator.Equal, data("IDEjercicio"))
            FilFacturas.Add("FechaFactura", FilterOperator.GreaterThan, data("FechaFactura"))
            Dim DtFacturas As DataTable = New FacturaVentaCabecera().Filter(FilFacturas)
            If Not DtFacturas Is Nothing AndAlso DtFacturas.Rows.Count > 0 Then
                If data.RowState = DataRowState.Added Then
                    ApplicationService.GenerateError("No se puede generar la factura con la fecha introducida. Existen facturas generadas posteriores a la fecha.")
                ElseIf data.RowState = DataRowState.Modified AndAlso Nz(data("FechaFactura")) <> Nz(data("FechaFactura", DataRowVersion.Original)) Then
                    ApplicationService.GenerateError("No se puede modificar la fecha de la factura con la fecha introducida. Existen facturas generadas posteriores a la fecha.")
                End If
            End If
        End If
    End Sub

#End Region

#Region " BusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        'TODO: ESto es necesario?
        '    Dim services As New ServiceProvider
        '    Dim TipoLineaDef As String = New TipoLinea().TipoLineaPorDefecto
        '    Dim strContador As String
        '    If ColumnName = "IDContador" Then
        '        strContador = Value
        '    Else
        '        strContador = current("IDContador")
        '    End If
        '    services.RegisterService(New ProcessInfo(strContador, TipoLineaDef))

        Dim oBRL As New SynonymousBusinessRules
        oBRL.AddSynonymous("FechaFactura", "Fecha")

        Dim services As New ServiceProvider
        oBRL = ProcessServer.ExecuteTask(Of BusinessRules, BusinessRules)(AddressOf ProcesoComercial.DetailBusinessRulesCab, oBRL, services)

        oBRL("IDCliente") = AddressOf CambioCliente

        oBRL.Add("IDContador", AddressOf CambioContador)
        oBRL.Add("Fecha", AddressOf CambioFechaFactura) 'El nuevo nombre indicado en el Synonimous
        oBRL.Add("FechaDeclaracionManual", AddressOf CambioDeclaracionManual)
        oBRL.Add("IDFacturaRectificada", AddressOf CambioFacturaRectificada)
        oBRL.Add("BaseRetencion", AddressOf CambioBaseRetencion)
        oBRL("IDCondicionPago") = AddressOf CambioCondicionPago
        oBRL.Add("IDFormaPago", AddressOf CambioFormaPago)
        oBRL.Add("Enviar347", AddressOf CambioEnviar347)
        oBRL.Add("Enviar349", AddressOf CambioEnviar349)
        oBRL.Add("RetencionIRPF", AddressOf CambioRetencionIRPF)
        oBRL.Add("CIFCliente", AddressOf CambioCifCliente)
        oBRL.Add("IDPais", AddressOf CambioIDPais)
        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioCifCliente(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Nz(data.Current("Enviar349"), False) Then
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf NegocioGeneral.ValidarEnvio349IPropAcc, data.Current, services)
        End If
'        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoComunes.ValidarDocumentoIdentificativoFV, data.Current, services)
    End Sub
    <Task()> Public Shared Sub CambioIDPais(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.AsignarEnvio347, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.AsignarEnvio349, data, services)
        If Nz(data.Current("Enviar349"), False) Then
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf NegocioGeneral.ValidarEnvio349IPropAcc, data.Current, services)
        End If
'       ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoComunes.ValidarDocumentoIdentificativoFV, data.Current, services)
    End Sub
    <Task()> Public Shared Sub CambioEnviar347(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Nz(data.Current("Enviar347"), False) Then
            data.Current("SinMensaje") = True
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf NegocioGeneral.ValidarEnvio347IPropAcc, data.Current, services)
            data.Current("Enviar349") = False
        End If
    End Sub

    <Task()> Public Shared Sub CambioEnviar349(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Nz(data.Current("Enviar349"), False) Then
            data.Current("SinMensaje") = True
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf NegocioGeneral.ValidarEnvio349IPropAcc, data.Current, services)
            data.Current("Enviar347") = False
        End If
    End Sub

    <Task()> Public Shared Sub CambioRetencionIRPF(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value

        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.AsignarEnvio347, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.AsignarEnvio349, data, services)
    End Sub

    <Task()> Public Shared Sub CambioCondicionPago(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.CambioCondicionPago, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf AsignarMotivoNoAsegurado, data, services)
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
                If Not FormasPagoSEPA Is Nothing AndAlso FormasPagoSEPA.Count > 0 AndAlso FormasPagoSEPA.Contains(UCase(data.Current("IDFormaPago"))) Then
                    ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf Cobro.GetMandatoSEPAPredeterminado, data.Current, services)
                End If
            End If
        End If
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf AsignarMotivoNoAsegurado, data, services)
    End Sub

    <Task()> Public Shared Sub CambioContador(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.ColumnName = "IDContador" Then data.Current(data.ColumnName) = data.Value

        If IsDate(data.Current("FechaFactura")) AndAlso Not IsDBNull(data.Current("IDContador")) Then
            Dim AppParamsConta As ParametroContabilidadVenta = services.GetService(Of ParametroContabilidadVenta)()
            If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf NegocioGeneral.ContadorB, data.Current("IDContador"), services) Then
                If AppParamsConta.Contabilidad Then data.Current("IDEjercicio") = ProcessServer.ExecuteTask(Of Date, String)(AddressOf NegocioGeneral.EjercicioPredeterminadoB, data.Current("FechaFactura"), services)
                'data.Current("Enviar347") = False
            Else
                If AppParamsConta.Contabilidad Then data.Current("IDEjercicio") = ProcessServer.ExecuteTask(Of Date, String)(AddressOf NegocioGeneral.EjercicioPredeterminado, data.Current("FechaFactura"), services)
                'data.Current("Enviar347") = True
            End If
        End If

        If Not IsDBNull(data.Current("IDContador")) Then
            Dim Contadores As EntityInfoCache(Of ContadorInfo) = services.GetService(Of EntityInfoCache(Of ContadorInfo))()
            Dim ContInfo As ContadorInfo = Contadores.GetEntity(data.Current("IDContador"))
            If Length(ContInfo.IDTipoComprobante) > 0 AndAlso Length(ContInfo.ClaveOperacion) > 0 Then
                '//Le asignamos la clave de operación del Tipo de Comprobante asociado al contador.
                data.Current("ClaveOperacion") = ContInfo.ClaveOperacion
            Else
                data.Current("ClaveOperacion") = System.DBNull.Value
            End If
        End If

        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.AsignarEnvio347, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.AsignarEnvio349, data, services)
    End Sub

    <Task()> Public Shared Sub CambioCliente(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComercial.CambioCliente, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComunes.CambioMoneda, data, services)

        Dim dir As New DataDireccionClte(enumpdTipoDireccion.pdDireccionFactura, "IDDireccion", data.Current)
        ProcessServer.ExecuteTask(Of DataDireccionClte)(AddressOf ProcesoComercial.AsignarDireccionCliente, dir, services)

        Dim Obs As New DataObservaciones(GetType(FacturaVentaCabecera).Name, "Texto", data.Current)
        ProcessServer.ExecuteTask(Of DataObservaciones)(AddressOf ProcesoComercial.AsignarObservacionesCliente, Obs, services)

        If Length(data.Current("IDCliente")) > 0 Then
            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data.Current("IDCliente"))

            If ClteInfo.Bloqueado Then ApplicationService.GenerateError("El Cliente está bloqueado.")
            data.Current("IDPais") = ClteInfo.Pais
            data.Current("Telefono") = ClteInfo.Telefono
            data.Current("Fax") = ClteInfo.Fax
            data.Current("IDModoTransporte") = ClteInfo.IDModoTransporte
            data.Current("IdBancoPropio") = ClteInfo.IDBancoPropio
            data.Current("DtoFactura") = ClteInfo.DtoComercial
            data.Current("IDTipoAsiento") = ClteInfo.IDTipoAsiento
            data.Current("RetencionIRPF") = ClteInfo.RetencionIRPF
            data.Current("EDI") = (Length(ClteInfo.IDConsignatario) > 0 And Length(ClteInfo.IDEDIFormato) > 0)
            data.Current("IDProveedor") = ClteInfo.IDProveedor
            data.Current("IDOperario") = ClteInfo.IDOperario
            data.Current("IDClienteInicial") = data.Current("IDCliente")

            Dim IDBanco As Integer = ProcessServer.ExecuteTask(Of String, Integer)(AddressOf ClienteBanco.GetBancoPredeterminado, data.Current("IDCliente"), services)
            If IDBanco > 0 Then
                data.Current("IDClienteBanco") = IDBanco
            Else
                data.Current("IDClienteBanco") = System.DBNull.Value
            End If
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf GetMandatoPredeterminado, data, services)
        Else
            data.Current("IDPais") = System.DBNull.Value
            data.Current("Telefono") = System.DBNull.Value
            data.Current("Fax") = System.DBNull.Value
            data.Current("IDModoTransporte") = System.DBNull.Value
            data.Current("IdBancoPropio") = System.DBNull.Value
            data.Current("DtoFactura") = 0
            data.Current("IDTipoAsiento") = System.DBNull.Value
            data.Current("RetencionIRPF") = 0
            data.Current("IDClienteBanco") = System.DBNull.Value
            data.Current("IDMandato") = System.DBNull.Value
            data.Current("EDI") = False
        End If
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.AsignarEnvio347, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.AsignarEnvio349, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf AsignarMotivoNoAsegurado, data, services)
'        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoComunes.ValidarDocumentoIdentificativoFV, data.Current, services)
    End Sub

    <Task()> Public Shared Sub AsignarMotivoNoAsegurado(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf AsignarMotivoNoAseguradoIProp, data.Current, services)
    End Sub

    <Task()> Public Shared Sub AsignarMotivoNoAseguradoIProp(ByVal data As IPropertyAccessor, ByVal services As ServiceProvider)
        Dim IDMotivoNoAsegurado As String
        Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
        Dim ClteInfo As ClienteInfo
        If Length(data("IDCliente")) > 0 Then
            ClteInfo = Clientes.GetEntity(data("IDCliente"))
            IDMotivoNoAsegurado = ClteInfo.IDMotivoNoAsegurado
        End If

        If Length(data("IDCondicionPago")) > 0 AndAlso (ClteInfo Is Nothing OrElse ClteInfo.CondicionPago <> data("IDCondicionPago")) Then
            Dim CondicionesPago As EntityInfoCache(Of CondicionPagoInfo) = services.GetService(Of EntityInfoCache(Of CondicionPagoInfo))()
            Dim CPInfo As CondicionPagoInfo = CondicionesPago.GetEntity(data("IDCondicionPago"))
            If Length(CPInfo.IDMotivoNoAsegurado) > 0 Then IDMotivoNoAsegurado = CPInfo.IDMotivoNoAsegurado
        End If

        If Length(data("IDFormaPago")) > 0 AndAlso (ClteInfo Is Nothing OrElse ClteInfo.FormaPago <> data("IDFormaPago")) Then
            Dim FormasPago As EntityInfoCache(Of FormaPagoInfo) = services.GetService(Of EntityInfoCache(Of FormaPagoInfo))()
            Dim FPInfo As FormaPagoInfo = FormasPago.GetEntity(data("IDFormaPago"))
            If Length(FPInfo.IDMotivoNoAsegurado) > 0 Then IDMotivoNoAsegurado = FPInfo.IDMotivoNoAsegurado
        End If
        data("IDMotivoNoAsegurado") = IDMotivoNoAsegurado
    End Sub

    <Task()> Public Shared Sub GetMandatoPredeterminado(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.Current.ContainsKey("IDMandato") Then
            If Nz(data.Current("IDClienteBanco"), 0) <> 0 Then
                Dim fMandatoPred As New Filter
                fMandatoPred.Add(New BooleanFilterItem("Predeterminado", True))
                fMandatoPred.Add(New NumberFilterItem("IDClienteBanco", data.Current("IDClienteBanco")))
                Dim dtMandato As DataTable = AdminData.GetData("tbMaestroMandato", fMandatoPred)
                If dtMandato.Rows.Count > 0 Then
                    data.Current("IDMandato") = dtMandato.Rows(0)("IDMandato")
                Else
                    data.Current("IDMandato") = DBNull.Value
                End If
            Else
                data.Current("IDMandato") = DBNull.Value
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioFechaFactura(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.ColumnName = "Fecha" Then
            '//Hay que ponerlo en los dos campos indicados en el Synonimous.
            data.Current(data.ColumnName) = data.Value
            data.Current("FechaFactura") = data.Value
        End If
        If Length(data.Current("SuFechaFactura")) = 0 Then data.Current("SuFechaFactura") = data.Value

        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComunes.CambioMoneda, data, services)

        '//Se le da un id. ejercicio a la factura, en función del contador
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioContador, data, services)

        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioDeclaracionManual, data, services)
    End Sub

    <Task()> Public Shared Sub CambioDeclaracionManual(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        '// Comprobar si la Fecha Para declaración es manual. Si no lo es, la fecha
        '// pra declaración será la fecha de factura.
        If data.ColumnName = "FechaDeclaracionManual" Then data.Current(data.ColumnName) = Nz(data.Value, False)
        If Not data.Current("FechaDeclaracionManual") Then
            data.Current("FechaParaDeclaracion") = data.Current("FechaFactura")
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoFacturacionVenta.FechaParaDeclaracionComoProveedor, data.Current, services)
        End If
    End Sub

    <Task()> Public Shared Sub CambioFacturaRectificada(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.ColumnName = "IDFacturaRectificada" Then data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDFacturaRectificada")) > 0 Then data.Current("ClaveOperacion") = ClaveOperacion.FacturaRectificativa
    End Sub

    <Task()> Public Shared Sub CambioBaseRetencion(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDMoneda")) > 0 AndAlso Length(data.Current("CambioA")) > 0 AndAlso Length(data.Current("CambioB")) > 0 Then
            Dim ValAyB As New ValoresAyB(data.Current, data.Current("IDMoneda"), data.Current("CambioA"), data.Current("CambioB"))
            ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf MantenimientoValoresAyB, ValAyB, services)
        End If
    End Sub

#End Region

#Region " Declaraciones "

    <Task()> Public Shared Sub DeclararIVAVenta(ByVal data As DataDeclaraciones, ByVal services As ServiceProvider)
        '//Asigna el NDeclaracionIVA y el AñoDeclaracionIVA a las Facturas que se le indica mediante el Filtro
        If data.Filtro Is Nothing OrElse data.Filtro.Count = 0 Then
            ApplicationService.GenerateError("Debe indicar un filtro para seleccionar las Facturas a Declarar.")
        End If

        AdminData.Execute("sp_DeclaracionIVAVenta", False, data.NDeclaracion, data.AnioDeclaracion, AdminData.ComposeFilter(data.Filtro))
    End Sub


#End Region

#Region " Intrastat "

    Public Function ObtenerListasInstrastatVenta() As DataTable
        Return New BE.DataEngine().Filter("vCtlCIIntrastatVentaListas", "*", "")
    End Function

    Public Function ObtenerDatosRptIntrastatVenta(ByVal f As Filter) As DataTable
        Return New BE.DataEngine().Filter("vRptIntrastatVentas", f)
    End Function

    Public Function GrabarDeclaracionIntrastat(ByVal FilterFacturas As Filter, ByVal strNDeclaracion As String, ByVal strAñoDeclaracion As String) As Integer
        Dim dtFacturas As DataTable = Filter(FilterFacturas)
        For Each oRw As DataRow In dtFacturas.Rows
            If strNDeclaracion = "0000" Then
                oRw("NDeclaracionIntrastat") = System.DBNull.Value
            Else
                oRw("NDeclaracionIntrastat") = strNDeclaracion
            End If
            If strAñoDeclaracion = "0000" Then
                oRw("AñoDeclaracionIntrastat") = System.DBNull.Value
            Else
                oRw("AñoDeclaracionIntrastat") = strAñoDeclaracion
            End If
        Next

        BusinessHelper.UpdateTable(dtFacturas)
    End Function

#End Region

#Region " Precio Optimo "

    <Serializable()> _
    Public Class DataPrecioOptimo
        Public IDFactura As Integer
        Public FechaCalculo As Date
        Public DocFra As DocumentoFacturaVenta

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDFactura As Integer, ByVal FechaCalculo As Date)
            Me.IDFactura = IDFactura
            Me.FechaCalculo = FechaCalculo
        End Sub
    End Class

    <Serializable()> _
    Public Class DataCalcPrecioOpt
        Public FechaCalculo As Date
        Public DocFra As DocumentoFacturaVenta

        Public Sub New()
        End Sub
        Public Sub New(ByVal FechaCalculo As Date, ByVal DocFra As DocumentoFacturaVenta)
            Me.FechaCalculo = FechaCalculo
            Me.DocFra = DocFra
        End Sub
    End Class

    <Task()> Public Shared Sub PrecioOptimo(ByVal data As DataPrecioOptimo, ByVal services As ServiceProvider)
        Dim DocFra As DocumentoFacturaVenta = ProcessServer.ExecuteTask(Of Integer, DocumentoFacturaVenta)(AddressOf CrearDocumento, data.IDFactura, services)
        Dim StData As New DataCalcPrecioOpt(data.FechaCalculo, DocFra)
        ProcessServer.ExecuteTask(Of DataCalcPrecioOpt)(AddressOf CalculoPrecioOptimo, StData, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularRepresentantes, DocFra, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularAnalitica, DocFra, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularBasesImponibles, DocFra, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularTotales, DocFra, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.CalcularVencimientos, DocFra, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf ProcesoFacturacionVenta.GrabarDocumento, DocFra, services)
    End Sub

    <Task()> Public Shared Function CrearDocumento(ByVal IDFactura As Integer, ByVal services As ServiceProvider) As DocumentoFacturaVenta
        Return New DocumentoFacturaVenta(IDFactura)
    End Function

    <Task()> Public Shared Sub CalculoPrecioOptimo(ByVal data As DataCalcPrecioOpt, ByVal services As ServiceProvider)
        If data.DocFra Is Nothing OrElse data.DocFra.dtLineas Is Nothing OrElse data.DocFra.dtLineas.Rows.Count = 0 Then Exit Sub

        '//Recogemos los articulos que esten relacionados con esa Factura.
        Dim dtArticulosFactura As DataTable = New BE.DataEngine().Filter("vNegFacturaVentaLineaArticulos", New StringFilterItem("IDFactura", data.DocFra.HeaderRow("IDFactura")))
        Dim f As New Filter
        For Each drArticuloFactura As DataRow In dtArticulosFactura.Select
            f.Clear()
            f.Add("IDArticulo", drArticuloFactura("IDArticulo"))

            '//Recogemos las lineas de la factura que tengan el articulo de este momento
            Dim Cantidad As Double = Nz(data.DocFra.dtLineas.Compute("SUM(Cantidad)", f.Compose(New AdoFilterComposer)), 0)
            Dim dataTarifa As New DataCalculoTarifaComercial
            dataTarifa.IDArticulo = drArticuloFactura("IDArticulo")
            dataTarifa.IDCliente = data.DocFra.IDCliente
            dataTarifa.Cantidad = Cantidad
            dataTarifa.Fecha = data.FechaCalculo

            ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial, DataTarifaComercial)(AddressOf ProcesoComercial.TarifaComercial, dataTarifa, services)
            If Not dataTarifa.DatosTarifa Is Nothing AndAlso dataTarifa.DatosTarifa.Precio <> 0 Then
                Dim FVL As New FacturaVentaLinea
                Dim context As New BusinessData(data.DocFra.HeaderRow)
                Dim WhereArticulo As String = f.Compose(New AdoFilterComposer)
                For Each drAlbaranLineaArticulo As DataRow In data.DocFra.dtLineas.Select(WhereArticulo)
                    FVL.ApplyBusinessRule("Precio", dataTarifa.DatosTarifa.Precio, drAlbaranLineaArticulo, context)
                    FVL.ApplyBusinessRule("Dto1", dataTarifa.DatosTarifa.Dto1, drAlbaranLineaArticulo, context)
                    FVL.ApplyBusinessRule("Dto2", dataTarifa.DatosTarifa.Dto2, drAlbaranLineaArticulo, context)
                    FVL.ApplyBusinessRule("Dto3", dataTarifa.DatosTarifa.Dto3, drAlbaranLineaArticulo, context)
                    FVL.ApplyBusinessRule("UDValoracion", dataTarifa.DatosTarifa.UDValoracion, drAlbaranLineaArticulo, context)

                    drAlbaranLineaArticulo("SeguimientoTarifa") = "Fecha Asignación de Precio: " & data.FechaCalculo & " - " & dataTarifa.DatosTarifa.SeguimientoTarifa
                    'If Length(dataTarifa.DatosTarifa.SeguimientoTarifa) > 0 Then
                    '    drAlbaranLineaArticulo("SeguimientoTarifa") = Left(dataTarifa.DatosTarifa.SeguimientoTarifa, Length(drAlbaranLineaArticulo("SeguimientoTarifa")))
                    'End If
                    'If Length(dtTarifa.Rows(0)("SeguimientoDtos")) > 0 Then
                    '    If Length(drAlbaranLineaArticulo("SeguimientoTarifa")) > 0 Then
                    '        drAlbaranLineaArticulo("SeguimientoTarifa") = Left(drAlbaranLineaArticulo("SeguimientoTarifa") & dtTarifa.Rows(0)("SeguimientoDtos"), Length(drAlbaranLineaArticulo("SeguimientoTarifa")))
                    '    Else
                    '        drAlbaranLineaArticulo("SeguimientoTarifa") = Left(dtTarifa.Rows(0)("SeguimientoDtos"), Length(dtTarifa.Rows(0)("SeguimientoTarifa")))
                    '    End If
                    'End If
                Next
            End If
            Cantidad = 0
        Next
    End Sub

#End Region

#Region " Contabilización y declaración "

    <Serializable()> _
   Public Class DataDeclaracion
        Public IdProcess As Guid
        Public dt As DataTable
    End Class
    <Serializable()> _
   Public Class DataContabilizacion
        Public FilFact As Filter
        Public Estado As Integer
    End Class
    <Task()> Public Shared Function AccionDeclaracion(ByVal DatosDeclaracion As DataDeclaracion, ByVal services As ServiceProvider) As String

        If Not DatosDeclaracion.dt Is Nothing AndAlso DatosDeclaracion.dt.Rows.Count = 0 Then Exit Function

        Dim CurrentBD As Guid = AdminData.GetConnectionInfo.IDDataBase

        Dim BBDDs As List(Of Object) = (From c In DatosDeclaracion.dt Select c("IDBaseDatos") Distinct).ToList

        Dim dtCambios As DataTable
        Dim IDsRiesgoSuperado(-1) As String
        For Each strIDBaseDatos As String In BBDDs
            Dim IDBaseDatos As New Guid(strIDBaseDatos)
            dtCambios = DatosDeclaracion.dt.Clone
            Try
                If CurrentBD <> IDBaseDatos Then
                    AdminData.SetCurrentConnection(IDBaseDatos)
                    AdminData.CommitTx(True)
                End If

                For Each dr As DataRow In DatosDeclaracion.dt.Select("IDBaseDatos = " & Quoted(IDBaseDatos.ToString))
                    dr("EnviadaEntidadAseguradora") = True
                    dtCambios.ImportRow(dr)
                Next

                If Not dtCambios Is Nothing AndAlso dtCambios.Rows.Count > 0 Then
                    BusinessHelper.UpdateTable(dtCambios)
                    dtCambios.Rows.Clear()
                End If
            Catch ex As Exception
                AdminData.RollBackTx()
            Finally
                If CurrentBD <> IDBaseDatos Then
                    AdminData.CommitTx(True)
                    AdminData.SetCurrentConnection(CurrentBD)
                    AdminData.CommitTx(True)
                End If
            End Try
        Next

    End Function

    <Task()> Public Shared Function CambiarEstadoContabFacturas(ByVal DatosContabilizacion As DataContabilizacion, ByVal services As ServiceProvider) As Boolean
        If Not DatosContabilizacion.FilFact Is Nothing AndAlso DatosContabilizacion.FilFact.Count > 0 Then
            Dim StrSql As String = "UPDATE tbFacturaVentaCabecera "
            StrSql &= "SET Estado = " & DatosContabilizacion.Estado & " "
            StrSql &= "WHERE " & AdminData.ComposeFilter(DatosContabilizacion.FilFact) & " "
            AdminData.Execute(StrSql)
            Return True
        End If
    End Function
#End Region

    'TODO: Pendiente Promotoras
#Region " Facturacion - Promotoras "

    <Serializable()> _
    Public Class DataNuevaLineaFacturaPromo
        Public IDFactura As Integer
        Public NFactura As String
        Public FechaFactura As Date
    End Class

    Private Function NuevaLineaFacturaObraPromo(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If Doc.HeaderRow("TipoFactura") = enumfvcTipoFactura.fvcFinal Then
            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf NuevaLineaFacturaObraPromoFinal, Doc, services)
        Else
            ProcessServer.ExecuteTask(Of DocumentoFacturaVenta)(AddressOf NuevaLineaFacturaObraPromoAnticipo, Doc, services)
        End If
    End Function

    <Task()> Public Shared Sub NuevaLineaFacturaObraPromoAnticipo(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)
        If Not Doc Is Nothing AndAlso Not Doc.dtLineas Is Nothing Then
            Dim fvl As New FacturaVentaLinea

            Dim ids(-1) As Object
            For Each dr As DataRow In Doc.dtLineas.Rows
                ReDim Preserve ids(ids.Length)
                ids(ids.Length - 1) = dr("IDLocalVencimiento")
            Next

            Dim oFltr As New Filter
            oFltr.Add(New InListFilterItem("IDLocalVencimiento", ids, FilterType.Numeric))
            oFltr.Add(New BooleanFilterItem("Facturado", False))

            Dim ObraPLV As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraPromoLocalVencimiento")
            Dim dtVencimiento As DataTable = ObraPLV.Filter(oFltr)

            Dim AppParamsConta As ParametroContabilidadVenta = services.GetService(Of ParametroContabilidadVenta)()
            Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
            Dim MonInfoA As MonedaInfo = Monedas.MonedaA

            If Not dtVencimiento Is Nothing AndAlso dtVencimiento.Rows.Count > 0 Then
                Dim intIDOrdenLinea As Integer = 0
                Dim Obra As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraCabecera")
                Dim TipoLineaPredet As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, services)
                ' Dim oArt As New Articulo
                Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                For Each drVencimiento As DataRow In dtVencimiento.Rows
                    Doc.dtLineas.DefaultView.RowFilter = "IdLocalVencimiento = " & drVencimiento("IdLocalVencimiento")
                    If Doc.dtLineas.Rows.Count > 0 Then
                        Dim drObra As DataRow = Obra.GetItemRow(drVencimiento("IDObra"))

                        Dim drlinea As DataRow = fvl.AddNewForm.Rows(0)

                        drlinea("IDFactura") = Doc.HeaderRow("IDFactura")
                        drlinea("NFactura") = Doc.HeaderRow("NFactura")
                        drlinea("IdLocalVencimiento") = drVencimiento("IdLocalVencimiento")
                        intIDOrdenLinea = intIDOrdenLinea + 1
                        drlinea("IDOrdenLinea") = intIDOrdenLinea
                        drlinea("IDArticulo") = drVencimiento("IDArticulo")

                        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(drVencimiento("IDArticulo"))
                        Dim strTextoLineaFactura As String = Doc.dtLineas.Rows(0)("Descripcion2")
                        If Length(strTextoLineaFactura) > 0 Then
                            Dim strTextoLinea As String = Doc.dtLineas.Rows(0)("Descripcion3") & " " & strTextoLineaFactura & " en " & Doc.dtLineas.Rows(0)("DireccionObra")
                            drlinea("DescArticulo") = strTextoLinea & " correspondiente al vencimiento " & drVencimiento("FechaVencimiento")
                        ElseIf Length(drVencimiento("DescVencimiento")) = 0 Then
                            drlinea("DescArticulo") = ArtInfo.DescArticulo
                        Else
                            drlinea("DescArticulo") = drVencimiento("DescVencimiento")
                        End If
                        If AppParamsConta.Contabilidad Then
                            If Length(Doc.HeaderRow("CCAnticipo")) = 0 Then
                                Dim Nacional As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf Cliente.Nacional, drVencimiento("IdCliente"), services)
                                If Nacional Then
                                    drlinea("CContable") = ArtInfo.CCVenta
                                Else
                                    drlinea("CContable") = ArtInfo.CCExport
                                End If
                            Else
                                drlinea("CContable") = Doc.HeaderRow("CCAnticipo")
                            End If
                        End If

                        drlinea("IDTipoIva") = drVencimiento("IDTipoIva")
                        drlinea("IDCentroGestion") = drObra("IDCentroGestion")

                        drlinea("Cantidad") = 1
                        drlinea("UdValoracion") = 1
                        drlinea("IDUDMedida") = ArtInfo.IDUDVenta
                        drlinea("IDUDInterna") = ArtInfo.IDUDInterna
                        Dim StDatos As New ArticuloUnidadAB.DatosFactorConversion
                        StDatos.IDArticulo = drVencimiento("IDArticulo")
                        StDatos.IDUdMedidaA = drlinea("IDUDMedida") & String.Empty
                        StDatos.IDUdMedidaB = drlinea("IDUDInterna")
                        StDatos.UnoSiNoExiste = True
                        drlinea("Factor") = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, StDatos, services)
                        drlinea("QInterna") = drlinea("Factor") * drlinea("Cantidad")

                        drlinea("IDTipoLinea") = TipoLineaPredet
                        drlinea("IDObra") = drVencimiento("IDObra")

                        If Nz(drObra("CambioA"), 0) > 0 Then
                            drlinea("Precio") = Nz(drVencimiento("ImpVencimientoA"), 0) / drObra("CambioA")
                        End If
                        If MonInfoA.ID <> Doc.IDMoneda Then
                            Dim datos As New DataCambioMoneda(New DataRowPropertyAccessor(drlinea), MonInfoA.ID, Doc.IDMoneda, Doc.Fecha)
                            ProcessServer.ExecuteTask(Of DataCambioMoneda)(AddressOf NegocioGeneral.CambioMoneda, datos, services)
                        End If

                        drlinea("Importe") = drlinea("Precio")
                        Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(drlinea), drObra("IDMoneda") & String.Empty, Nz(drObra("CambioA"), 0), Nz(drObra("CambioB"), 0))
                        ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf MantenimientoValoresAyB, ValAyB, services)

                        Doc.dtLineas.Rows.Add(drlinea.ItemArray)
                    End If
                Next
            End If
            Doc.dtLineas.DefaultView.RowFilter = String.Empty
        End If
    End Sub

    <Task()> Public Shared Sub NuevaLineaFacturaObraPromoFinal(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)

        If Not Doc Is Nothing AndAlso Not Doc.dtLineas Is Nothing Then
            Dim fvl As New FacturaVentaLinea

            Dim strCContable As String
            Dim strDescArticulo As String
            Dim strIDTipoIva As String
            Dim dblPrecioA As Double

            Dim ids(-1) As Object
            For Each dr As DataRow In Doc.dtLineas.Rows
                ReDim Preserve ids(ids.Length)
                ids(ids.Length - 1) = dr("IDLocal")
            Next

            Dim ObraPL As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraPromoLocal")
            Dim dtOPL As DataTable = ObraPL.Filter(New InListFilterItem("IDLocal", ids, FilterType.Numeric))

            Dim AppParamsConta As ParametroContabilidadVenta = services.GetService(Of ParametroContabilidadVenta)()
            If dtOPL.Rows.Count > 0 Then
                Dim Obra As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraCabecera")

                Dim TipoLineaPredet As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, services)
                Dim oArt As New Articulo
                Dim oCliente As New Cliente

                For Each drOPL As DataRow In dtOPL.Rows
                    Dim drObra As DataRow = Obra.GetItemRow(drOPL("IDObra"))

                    Doc.dtLineas.DefaultView.RowFilter = "IDLocal= " & drOPL("IdLocal")

                    Dim j As Integer = 0
                    Dim intTotalAnticipos As Integer = 0
                    Dim dtAnticipo As DataTable = New BE.DataEngine().Filter("vNegAnticiposPorLocal", "*", "IDLocal= " & drOPL("IdLocal"))
                    If dtAnticipo.Rows.Count > 0 Then
                        intTotalAnticipos = dtAnticipo.Rows.Count
                    End If
                    Dim blnTieneVivienda As Boolean

                    For NLin As Integer = 1 To 2 + intTotalAnticipos
                        If Doc.dtLineas.Rows.Count > 0 Then
                            Dim drlinea As DataRow = fvl.AddNewForm.Rows(0)
                            dblPrecioA = 0
                            drlinea("IDFactura") = Doc.HeaderRow("IDFactura")
                            drlinea("NFactura") = Doc.HeaderRow("NFactura")
                            drlinea("IDOrdenLinea") = NLin
                            drlinea("IDArticulo") = Doc.dtLineas.Rows(0)("IDArticulo")
                            strIDTipoIva = Doc.dtLineas.Rows(0)("IDTipoIva") & String.Empty

                            Dim drArt As DataRow = oArt.GetItemRow(drlinea("IDArticulo"))
                            Dim Nacional As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf Cliente.Nacional, drOPL("IdCliente"), services)
                            Select Case NLin
                                Case 1 'Importe total de la Venta
                                    strDescArticulo = Doc.dtLineas.Rows(0)("Descripcion4") & " " & Doc.dtLineas.Rows(0)("Descripcion2") & " en " & Doc.dtLineas.Rows(0)("DireccionObra")
                                    If AppParamsConta.Contabilidad Then strCContable = IIf(Nacional, drArt("CCVenta"), drArt("CCExport")) & String.Empty
                                    If drOPL("PrecioGarajeA") = 0 Then
                                        dblPrecioA = drOPL("ImpVentaA")
                                    Else
                                        dblPrecioA = drOPL("ImpVentaA") - drOPL("PrecioGarajeA")
                                    End If
                                    drlinea("IDLocalVencimiento") = Doc.dtLineas.Rows(0)("IDLocalVencimiento")
                                    If dblPrecioA > 0 Then blnTieneVivienda = True
                                Case 2 'Importe Garaje
                                    If drOPL("PrecioGarajeA") > 0 Then
                                        Dim strTextoLineaFacturaGaraje As String = "Garaje nº " & Doc.dtLineas.Rows(0)("NumeroGaraje")
                                        If Length(Doc.dtLineas.Rows(0)("Edificio")) > 0 Then
                                            strTextoLineaFacturaGaraje = strTextoLineaFacturaGaraje & " del Edificio " & Doc.dtLineas.Rows(0)("Edificio")
                                        End If
                                        If Length(Doc.dtLineas.Rows(0)("DireccionObra")) > 0 Then
                                            strTextoLineaFacturaGaraje = strTextoLineaFacturaGaraje & " en " & Doc.dtLineas.Rows(0)("DireccionObra")
                                        End If
                                        strDescArticulo = Doc.dtLineas.Rows(0)("Descripcion4") & " " & strTextoLineaFacturaGaraje

                                        If AppParamsConta.Contabilidad Then strCContable = IIf(Nacional, drArt("CCVenta"), drArt("CCExport")) & String.Empty
                                        dblPrecioA = drOPL("PrecioGarajeA")
                                        If Not blnTieneVivienda Then
                                            drlinea("IDLocalVencimiento") = Doc.dtLineas.Rows(0)("IDLocalVencimiento")
                                        End If
                                    End If
                                Case Else 'Anticipos
                                    strDescArticulo = "Importe entregado a cuenta"
                                    dblPrecioA = Nz(dtAnticipo.Rows(j)("ImpAnticipoA"), 0) * -1
                                    If AppParamsConta.Contabilidad Then
                                        strCContable = dtAnticipo.Rows(j)("CCAnticipo") & String.Empty
                                        If Len(strCContable) = 0 Then strCContable = Doc.HeaderRow("CCAnticipo") & String.Empty
                                    End If
                                    j = j + 1
                            End Select

                            If dblPrecioA <> 0 Then
                                drlinea("DescArticulo") = strDescArticulo
                                If AppParamsConta.Contabilidad Then drlinea("CContable") = strCContable
                                drlinea("IDTipoIva") = strIDTipoIva
                                drlinea("IDCentroGestion") = drObra("IDCentroGestion") & String.Empty
                                If Nz(drObra("CambioA"), 0) > 0 Then
                                    drlinea("Precio") = dblPrecioA / drObra("CambioA")
                                End If
                                drlinea("IDUDMedida") = drArt("IDUdVenta")
                                drlinea("IDUDInterna") = drArt("IDUDInterna")
                                drlinea("UdValoracion") = 1

                                drlinea("IDTipoLinea") = TipoLineaPredet
                                drlinea("IDObra") = drOPL("IDObra")
                                drlinea("Cantidad") = 1
                                drlinea("Importe") = drlinea("Precio") * drlinea("cantidad")


                                Dim StDatos As New ArticuloUnidadAB.DatosFactorConversion
                                StDatos.IDArticulo = drlinea("IDArticulo")
                                StDatos.IDUdMedidaA = drlinea("IDUDMedida") & String.Empty
                                StDatos.IDUdMedidaB = drlinea("IDUDInterna")
                                StDatos.UnoSiNoExiste = True
                                drlinea("Factor") = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, StDatos, services)
                                drlinea("QInterna") = drlinea("Factor") * drlinea("Cantidad")

                                Doc.dtLineas.Rows.Add(drlinea.ItemArray)
                            End If
                        End If
                    Next

                Next
            End If

            Doc.dtLineas.DefaultView.RowFilter = String.Empty
        End If
    End Sub

    <Task()> Private Shared Function NuevaLineaFacturaObraCertificacion(ByVal Doc As DocumentoFacturaVenta, ByVal services As ServiceProvider)

        Dim fvl As New FacturaVentaLinea

        Dim Lineas As DataTable = fvl.AddNew
        Dim LineasAux As DataTable = fvl.AddNew
        Dim dtObraAux As DataTable

        If Not Doc Is Nothing AndAlso Not Doc.dtLineas Is Nothing Then
            Dim ids(-1) As Object
            For Each dr As DataRow In Doc.dtLineas.Rows
                ReDim Preserve ids(ids.Length)
                ids(ids.Length - 1) = dr("IDTrabajo")
            Next

            Dim ObraTbjo As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraTrabajo")
            Dim dtVencimiento As DataTable = ObraTbjo.Filter(New InListFilterItem("IDTrabajo", ids, FilterType.Numeric), "IDTrabajo")
            Dim dblTotalLineas As Double

            Dim strTipoLinea As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, services)

            If dtVencimiento.Rows.Count > 0 Then
                Dim intIDOrdenLinea As Integer = 0
                Dim Obra As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraCabecera")
                Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                'Dim oArt As New Articulo
                dtObraAux = Obra.AddNew
                dtObraAux.DefaultView.Sort = "IDObra"
                Dim AppParamsConta As ParametroContabilidadVenta = services.GetService(Of ParametroContabilidadVenta)()
                For Each drVencimiento As DataRow In dtVencimiento.Rows
                    Doc.dtLineas.DefaultView.RowFilter = "IDTrabajo = " & drVencimiento("IDTrabajo")
                    If Doc.dtLineas.Rows.Count > 0 Then
                        Dim drlinea As DataRow = fvl.AddNewForm.Rows(0)
                        Dim drObra As DataRow = Obra.GetItemRow(drVencimiento("IDObra"))

                        drlinea("IDFactura") = Doc.HeaderRow("IDFactura")
                        drlinea("NFactura") = Doc.HeaderRow("NFactura")
                        intIDOrdenLinea = intIDOrdenLinea + 1
                        drlinea("IDOrdenLinea") = intIDOrdenLinea
                        drlinea("IDArticulo") = drVencimiento("IDArticulo")
                        drlinea("PedidoCliente") = drObra("NumeroPedido")

                        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(drVencimiento("IDArticulo"))
                        If Length(drVencimiento("DescTrabajo")) = 0 Then
                            drlinea("DescArticulo") = ArtInfo.DescArticulo
                        Else
                            drlinea("DescArticulo") = drVencimiento("DescTrabajo")
                        End If
                        If AppParamsConta.Contabilidad Then
                            Dim Nacional As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf Cliente.Nacional, drObra("IdCliente"), services)

                            If Nacional Then
                                drlinea("CContable") = ArtInfo.CCVenta
                            Else
                                drlinea("CContable") = ArtInfo.CCExport
                            End If
                        End If
                        drlinea("IDTipoIva") = Doc.dtLineas.Rows(0)("IDTipoIva")
                        drlinea("IDCentroGestion") = Doc.dtLineas.Rows(0)("IDCentroGestion")
                        drlinea("IDUDMedida") = ArtInfo.IDUDVenta
                        drlinea("IDUDInterna") = ArtInfo.IDUDInterna
                        drlinea("UDValoracion") = 1
                        drlinea("TipoFactAlquiler") = drVencimiento("TipoFactAlquiler")
                        drlinea("IDObra") = drVencimiento("IDObra")
                        drlinea("IDTrabajo") = drVencimiento("IDTrabajo")
                        drlinea("IDCertificacion") = Doc.dtLineas.Rows(0)("IDCertificacion")
                        drlinea("Cantidad") = Nz(Doc.dtLineas.Rows(0)("QCertificada"), 0)
                        drlinea("Precio") = Nz(drVencimiento("ImpPrevTrabajoVentaA"), 0) / Nz(drObra("CambioA"), 1)
                        drlinea("Importe") = drlinea("Cantidad") * drlinea("Precio")
                        drlinea("PrecioCosteA") = 0
                        drlinea("PrecioCosteB") = 0
                        dblTotalLineas = dblTotalLineas + drlinea("Importe")

                        drlinea("Factor") = 1
                        drlinea("QInterna") = drlinea("Factor") * drlinea("Cantidad")
                        drlinea("IDTipoLinea") = strTipoLinea

                        Lineas.Rows.Add(drlinea.ItemArray)

                        ''''''''''''Generación de líneas adicionales''''''''''''
                        If drObra("GastosGenerales") > 0 OrElse drObra("BeneficioIndustrial") > 0 OrElse drObra("CoefBaja") > 0 Then
                            LineasAux.Rows.Add(drlinea.ItemArray)
                            If dtObraAux.Rows.Count > 0 Then
                                If dtObraAux.DefaultView.Find(drVencimiento("IDObra")) < 0 Then
                                    dtObraAux.Rows.Add(drObra.ItemArray)
                                End If
                            Else
                                dtObraAux.Rows.Add(drObra.ItemArray)
                            End If
                        End If
                    End If
                Next
            End If

            If Not IsNothing(LineasAux) AndAlso LineasAux.Rows.Count > 0 Then
                Dim datNewLinCertif As New DataNuevaLineaCertificacion(Doc, dtObraAux, dblTotalLineas)
                ProcessServer.ExecuteTask(Of DataNuevaLineaCertificacion)(AddressOf NuevaLineaFacturaObraCertificacionDatosAdicionales, datNewLinCertif, services)
            End If

            Doc.dtLineas.DefaultView.RowFilter = String.Empty
            Return Lineas
        End If
    End Function

    <Serializable()> _
    Public Class DataNuevaLineaCertificacion
        Public Doc As DocumentoFacturaVenta
        Public Obras As DataTable
        Public TotalLineas As Double

        Public Sub New(ByVal Doc As DocumentoFacturaVenta, ByVal Obras As DataTable, ByVal TotalLineas As Double)
            Me.Doc = Doc
            Me.Obras = Obras
            Me.TotalLineas = TotalLineas
        End Sub
    End Class
    <Task()> Public Shared Sub NuevaLineaFacturaObraCertificacionDatosAdicionales(ByVal data As DataNuevaLineaCertificacion, ByVal services As ServiceProvider)
        If Not data.Doc.dtLineas Is Nothing AndAlso data.Doc.dtLineas.Rows.Count > 0 Then
            Dim dvLineasFc As New DataView(data.Doc.dtLineas)
            Dim fvl As New FacturaVentaLinea

            Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
            Dim strIDArticulo As String = AppParams.ArticuloFacturacionProyectos
            Dim strTipoLinea As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, services)

            For Each drObra As DataRow In data.Obras.Rows
                Dim drLinea As DataRow = data.Doc.dtLineas.NewRow
                Dim intIDOrdenLinea = data.Doc.dtLineas.Rows.Count + 1
                For i As Integer = 0 To 2
                    'Se tienen que generar 3 líneas como máximo (Si GastosGenerales,  
                    'BeneficioIndustrial o CoefBaja es cero no se generará línea)
                    drLinea("IDLineaFactura") = AdminData.GetAutoNumeric
                    drLinea("IDFactura") = data.Doc.HeaderRow("IDFactura")
                    drLinea("NFactura") = data.Doc.HeaderRow("NFactura")
                    drLinea("IDOrdenLinea") = intIDOrdenLinea
                    drLinea("IDArticulo") = strIDArticulo
                    drLinea("PedidoCliente") = drObra("NumeroPedido")
                    drLinea("UDValoracion") = 1
                    drLinea("Cantidad") = 1
                    drLinea("QInterna") = 1
                    drLinea("Factor") = 1
                    drLinea("IDTipoLinea") = strTipoLinea
                    drLinea("Regalo") = False
                    drLinea("IDCentroGestion") = data.Doc.HeaderRow("IDCentroGestion")

                    Dim Context As New BusinessData(data.Doc.HeaderRow)

                    drLinea = fvl.ApplyBusinessRule("IDArticulo", drLinea("IDArticulo"), drLinea, Context)

                    Dim blnCancel As Double = False
                    Select Case i
                        Case 0
                            If drObra("GastosGenerales") > 0 Then
                                drLinea("Precio") = data.TotalLineas * drObra("GastosGenerales") / 100
                                drLinea("DescArticulo") = "Gastos Generales"
                            Else
                                blnCancel = True
                            End If
                        Case 1
                            If drObra("BeneficioIndustrial") > 0 Then
                                drLinea("Precio") = data.TotalLineas * drObra("BeneficioIndustrial") / 100
                                drLinea("DescArticulo") = "Beneficio Industrial"
                            Else
                                blnCancel = True
                            End If
                        Case 2
                            If drObra("CoefBaja") > 0 Then
                                drLinea("Precio") = data.TotalLineas * drObra("CoefBaja") / 100
                                drLinea("DescArticulo") = "Coeficiente Baja"
                            Else
                                blnCancel = True
                            End If
                    End Select
                    drLinea("Importe") = drLinea("Precio")

                    If Not blnCancel Then
                        intIDOrdenLinea = intIDOrdenLinea + 1
                        data.Doc.dtLineas.Rows.Add(drLinea.ItemArray)
                    End If
                Next
            Next
        End If
    End Sub

#End Region

    'TODO: Pendiente Exportación facturas
#Region " Exportación de Facturas "

    Public Sub ExportarFacturas(ByVal dtFacturas As DataTable, ByVal strEjercicio As String, ByVal strBBDDDestino As String)
        If Not dtFacturas Is Nothing AndAlso dtFacturas.Rows.Count > 0 Then
            Dim g As New NegocioGeneral
            Dim strLstIDs As String

            For Each drFCC As DataRow In dtFacturas.Rows
                If Length(strLstIDs) > 0 Then strLstIDs = strLstIDs & ","
                strLstIDs = strLstIDs & drFCC("IDFactura")
            Next

            Dim dtFVCO As DataTable = Filter(, "IDFactura IN(" & strLstIDs & ")")
            Dim fvl As New FacturaVentaLinea
            Dim dtFVLO As DataTable = fvl.Filter(, "IDFactura IN(" & strLstIDs & ")")

            '//Preparar un DataTable con los pares IDEjercicio,IDDContable a exportar.
            '//Con las C.Contables de las Lineas de Factura y la C.Contable del Proveedor.
            Dim dtCuentasExportar As DataTable = PrepararCtasExportacion(dtFVCO, dtFVLO)

            '// Preparamos un DataTable con los IDObra, NObra, IDTrabajo, CodTrabajo a comprobar en la BD Destino.
            Dim dtObrasTrabajosExp As DataTable = g.PrepararObrasTrabajosExportacion(dtFVLO)

            Me.BeginTx()
            Dim strBBDDInic As String = AdminData.GetSessionInfo.DataBase.DataBaseName

            For Each origenDestino As DataRow In dtFVCO.Select
                origenDestino("Exportado") = 1
            Next

            '//Recuperamos el PlanContable en la BD Origen
            '//Contruir el filtro de los para tener el plan contable de los ejercicios seleccionados.
            Dim objFilterOR As New Filter(FilterUnionOperator.Or)
            Dim strIN As String
            For Each drFVCO As DataRow In dtFVCO.Rows
                If InStr(strIN, drFVCO("IDEjercicio"), CompareMethod.Text) = 0 Then
                    If Len(strIN) > 0 Then strIN = strIN & ","
                    strIN = strIN & drFVCO("IDEjercicio")

                    objFilterOR.Add(New StringFilterItem("IDEjercicio", drFVCO("IDEjercicio")))
                End If
            Next drFVCO
            Dim objNegPC As BusinessHelper = BusinessHelper.CreateBusinessObject("PlanContable")
            Dim dtPlanContOrigen As DataTable = objNegPC.Filter(objFilterOR)

            AdminData.SetSessionDataBase(strBBDDDestino)

            Dim dtFVCD As DataTable = Filter(New NoRowsFilterItem)
            Dim fvlD As New FacturaVentaLinea
            Dim dtFVLD As DataTable = fvlD.Filter(New NoRowsFilterItem)

            For Each drOrigenCabecera As DataRow In dtFVCO.Rows
                Dim drDestinoCabecera As DataRow = dtFVCD.NewRow
                For Each oCol As DataColumn In dtFVCO.Columns
                    drDestinoCabecera(oCol.ColumnName) = drOrigenCabecera(oCol)
                Next
                drDestinoCabecera("IDEjercicio") = strEjercicio
                dtFVCD.Rows.Add(drDestinoCabecera)
            Next

            For Each drOrigenLinea As DataRow In dtFVLO.Rows
                Dim drDestinoLinea As DataRow = dtFVLD.NewRow
                For Each oCol As DataColumn In dtFVLO.Columns
                    drDestinoLinea(oCol.ColumnName) = drOrigenLinea(oCol)
                Next
                dtFVLD.Rows.Add(drDestinoLinea)
            Next

            '//Exportamos las C.Contables del Plan Contable, antes de actualizar las Facturas.
            g.ExportarPlanContableFactura(strEjercicio, dtPlanContOrigen, dtCuentasExportar)

            BusinessHelper.UpdateTable(dtFVCD)

            '//Exportamos las Obras/ObrasTrabajos, antes de actualizar las Líneas de la Factura.
            dtFVLD = g.ExportarObraTrabajoFactura(dtObrasTrabajosExp, dtFVLD)
            '//NOTA:Hay que hacer un Update especial para la exportación, para poder exportar las facturas contabilizadas
            '//     y que se recalculen las Obras.
            fvl.UpdateExportacion(dtFVLD)
            Updated(dtFVCD) 'Recalculo de la factura

            '//Habilitamos la BD Origen. Es decir, la BD activa a partir de este momento es la de Origen.
            AdminData.SetSessionDataBase(strBBDDInic)
            BusinessHelper.UpdateTable(dtFVCO)

            Me.CommitTx()
        End If
    End Sub

    Private Function PrepararCtasExportacion(ByVal dtCabecera As DataTable, ByVal dtLineas As DataTable) As DataTable
        Dim dtCuentasExportar As DataTable = New NegocioGeneral().CrearDTExportacionCuentas()
        Dim drCuentasExportar As DataRow
        Dim objNegCliente As New Cliente
        Dim dtProveedor As DataTable

        For Each drFCCO As DataRow In dtCabecera.Rows
            dtProveedor = objNegCliente.SelOnPrimaryKey(drFCCO("IDCliente"))
            If Not IsNothing(dtProveedor) AndAlso dtProveedor.Rows.Count > 0 Then
                If Length(dtProveedor.Rows(0)("CCCliente") & String.Empty) > 0 Then
                    drCuentasExportar = dtCuentasExportar.NewRow
                    drCuentasExportar("IDEjercicio") = drFCCO("IDEjercicio") & String.Empty
                    drCuentasExportar("IDCContable") = dtProveedor.Rows(0)("CCCliente") & String.Empty
                    dtCuentasExportar.Rows.Add(drCuentasExportar)
                End If
            End If

            For Each drFCLO As DataRow In dtLineas.Rows
                If drFCCO("IDFactura") = drFCLO("IDFactura") Then
                    If Length(drFCLO("CContable") & String.Empty) > 0 Then
                        drCuentasExportar = dtCuentasExportar.NewRow
                        drCuentasExportar("IDEjercicio") = drFCCO("IDEjercicio") & String.Empty
                        drCuentasExportar("IDCContable") = drFCLO("CContable") & String.Empty
                        dtCuentasExportar.Rows.Add(drCuentasExportar)
                    End If
                End If
            Next drFCLO
        Next drFCCO
        objNegCliente = Nothing

        Return dtCuentasExportar
    End Function
#End Region

#Region " Consultas interactivas (Estadísticas) "

#Region " ObtenerEstadisticaCantidadesMeses "

    <Serializable()> _
Public Class DataEstadisticaCantidadesMeses
        Public CamposSelect As String
        Public CampoATotalizar As String
        Public CamposOrden As String
        Public GroupBy As String

        Public IDTipo As String
        Public IDFamilia As String
        Public IDSubFamilia As String
        Public IDArticulo As String
        Public IDCliente As String
        Public IDGrupoCliente As String
        Public IDMercado As String
        Public Provincia As String
        Public IDZona As String
        Public IDCentroGestion As String
        Public IDPais As String
        Public CEE As enumBoolean
        Public Extranjero As enumBoolean
        Public Año, Año2 As Integer
        Public EmpresaGrupo As enumBoolean

        Public Sub New(ByVal CamposSelect As String, ByVal CampoATotalizar As String, ByVal IDTipo As String, ByVal IDFamilia As String, ByVal IDSubFamilia As String, _
                       ByVal IDArticulo As String, ByVal IDCliente As String, ByVal IDGrupoCliente As String, ByVal IDMercado As String, ByVal Provincia As String, ByVal IDZona As String, _
                       ByVal IDCentroGestion As String, ByVal IDPais As String, ByVal CEE As enumBoolean, ByVal Extranjero As enumBoolean, _
                       ByVal Año As Integer, ByVal EmpresaGrupo As enumBoolean, ByVal GroupBy As String, ByVal CamposOrden As String)

            Me.CamposSelect = CamposSelect
            Me.CampoATotalizar = CampoATotalizar
            Me.CamposOrden = CamposOrden
            Me.GroupBy = GroupBy
            Me.IDTipo = IDTipo
            Me.IDFamilia = IDFamilia
            Me.IDSubFamilia = IDSubFamilia
            Me.IDArticulo = IDArticulo
            Me.IDCliente = IDCliente
            Me.IDGrupoCliente = IDGrupoCliente
            Me.IDMercado = IDMercado
            Me.Provincia = Provincia
            Me.IDZona = IDZona
            Me.IDCentroGestion = IDCentroGestion
            Me.IDPais = IDPais
            Me.CEE = CEE
            Me.Extranjero = Extranjero
            Me.Año = Año
            Me.EmpresaGrupo = EmpresaGrupo
        End Sub
        Public Sub New(ByVal CamposSelect As String, ByVal CampoATotalizar As String, ByVal IDTipo As String, ByVal IDFamilia As String, ByVal IDSubFamilia As String, _
                       ByVal IDArticulo As String, ByVal IDCliente As String, ByVal IDGrupoCliente As String, ByVal IDMercado As String, ByVal Provincia As String, ByVal IDZona As String, _
                       ByVal IDCentroGestion As String, ByVal IDPais As String, ByVal CEE As enumBoolean, ByVal Extranjero As enumBoolean, _
                       ByVal Año As Integer, ByVal Año2 As Integer, ByVal EmpresaGrupo As enumBoolean, ByVal GroupBy As String, ByVal CamposOrden As String)

            Me.CamposSelect = CamposSelect
            Me.CampoATotalizar = CampoATotalizar
            Me.CamposOrden = CamposOrden
            Me.GroupBy = GroupBy
            Me.IDTipo = IDTipo
            Me.IDFamilia = IDFamilia
            Me.IDSubFamilia = IDSubFamilia
            Me.IDArticulo = IDArticulo
            Me.IDCliente = IDCliente
            Me.IDGrupoCliente = IDGrupoCliente
            Me.IDMercado = IDMercado
            Me.Provincia = Provincia
            Me.IDZona = IDZona
            Me.IDCentroGestion = IDCentroGestion
            Me.IDPais = IDPais
            Me.CEE = CEE
            Me.Extranjero = Extranjero
            Me.Año = Año
            Me.Año2 = Año2
            Me.EmpresaGrupo = EmpresaGrupo
        End Sub

    End Class
    <Task()> Public Shared Function ObtenerEstadisticaCantidadesMeses(ByVal data As DataEstadisticaCantidadesMeses, ByVal services As ServiceProvider) As DataTable
        If data.Año2 = 0 Then
            Return ProcessServer.ExecuteTask(Of DataEstadisticaCantidadesMeses, DataTable)(AddressOf ObtenerEstadisticaCantidadesMesesAño, data, services)
        Else
            Dim dt1 As DataTable = ProcessServer.ExecuteTask(Of DataEstadisticaCantidadesMeses, DataTable)(AddressOf ObtenerEstadisticaCantidadesMesesAño, data, services)
            Dim Año As Integer = data.Año
            data.Año = data.Año2
            Dim dt2 As DataTable = ProcessServer.ExecuteTask(Of DataEstadisticaCantidadesMeses, DataTable)(AddressOf ObtenerEstadisticaCantidadesMesesAño, data, services)

            dt1 = ProcessServer.ExecuteTask(Of DataTable, DataTable)(AddressOf ADDColumnsAño2, dt1, services)
            For Each dr2 As DataRow In dt2.Select()
                Dim Campo As String = dt2.Columns(0).ColumnName
                Dim dr1() As DataRow = dt1.Select(New FilterItem(Campo, dr2(Campo)).Compose(New AdoFilterComposer))
                If dr1.Length = 0 Then
                    Dim dr As DataRow = dt1.NewRow
                    dr(Campo) = dr2(Campo)
                    dr(dt2.Columns(1).ColumnName) = dr2(dt2.Columns(1).ColumnName)
                    dr("Año2") = dr2("Año") : dr("Año") = Año
                    dr("SEnero2") = dr2("SEnero") : dr("SEnero") = 0
                    dr("SFebrero2") = dr2("SFebrero") : dr("SFebrero") = 0
                    dr("SMarzo2") = dr2("SMarzo") : dr("SMarzo") = 0
                    dr("SAbril2") = dr2("SAbril") : dr("SAbril") = 0
                    dr("SMayo2") = dr2("SMayo") : dr("SMayo") = 0
                    dr("SJunio2") = dr2("SJunio") : dr("SJunio") = 0
                    dr("SJulio2") = dr2("SJulio") : dr("SJulio") = 0
                    dr("SAgosto2") = dr2("SAgosto") : dr("SAgosto") = 0
                    dr("SSeptiembre2") = dr2("SSeptiembre") : dr("SSeptiembre") = 0
                    dr("SOctubre2") = dr2("SOctubre") : dr("SOctubre") = 0
                    dr("SNoviembre2") = dr2("SNoviembre") : dr("SNoviembre") = 0
                    dr("SDiciembre2") = dr2("SDiciembre") : dr("SDiciembre") = 0
                    dr("STotalLinea2") = dr2("STotalLinea") : dr("STotalLinea") = 0
                    dt1.Rows.Add(dr.ItemArray)
                Else
                    dr1(0)("Año2") = dr2("Año")
                    dr1(0)("SEnero2") = Nz(dr2("SEnero"), 0)
                    dr1(0)("SFebrero2") = Nz(dr2("SFebrero"), 0)
                    dr1(0)("SMarzo2") = Nz(dr2("SMarzo"), 0)
                    dr1(0)("SAbril2") = Nz(dr2("SAbril"), 0)
                    dr1(0)("SMayo2") = Nz(dr2("SMayo"), 0)
                    dr1(0)("SJunio2") = Nz(dr2("SJunio"), 0)
                    dr1(0)("SJulio2") = Nz(dr2("SJulio"), 0)
                    dr1(0)("SAgosto2") = Nz(dr2("SAgosto"), 0)
                    dr1(0)("SSeptiembre2") = Nz(dr2("SSeptiembre"), 0)
                    dr1(0)("SOctubre2") = Nz(dr2("SOctubre"), 0)
                    dr1(0)("SNoviembre2") = Nz(dr2("SNoviembre"), 0)
                    dr1(0)("SDiciembre2") = Nz(dr2("SDiciembre"), 0)
                    dr1(0)("STotalLinea2") = Nz(dr2("STotalLinea"), 0)
                End If
            Next
            For Each dr1 As DataRow In dt1.Select()
                Dim Campo As String = dt1.Columns(0).ColumnName
                Dim dr2() As DataRow = dt2.Select(New FilterItem(Campo, dr1(Campo)).Compose(New AdoFilterComposer))
                If dr2.Length = 0 Then
                    dr1("Año2") = data.Año2
                    dr1("SEnero2") = 0
                    dr1("SFebrero2") = 0
                    dr1("SMarzo2") = 0
                    dr1("SAbril2") = 0
                    dr1("SMayo2") = 0
                    dr1("SJunio2") = 0
                    dr1("SJulio2") = 0
                    dr1("SAgosto2") = 0
                    dr1("SSeptiembre2") = 0
                    dr1("SOctubre2") = 0
                    dr1("SNoviembre2") = 0
                    dr1("SDiciembre2") = 0
                    dr1("STotalLinea2") = 0
                End If
            Next

            Return dt1
        End If
    End Function

    <Task()> Public Shared Function ObtenerEstadisticaCantidadesMesesAño(ByVal data As DataEstadisticaCantidadesMeses, ByVal services As ServiceProvider) As DataTable
        Dim selectSQL As New System.Text.StringBuilder
        selectSQL.Append(String.Format( _
            "SELECT {0}, year([FechaFactura]) as Año, " & _
            "SUM(CASE MONTH([FechaFactura]) WHEN 1 THEN {1} ELSE 0 END) AS SEnero," & _
            "SUM(CASE MONTH([FechaFactura]) WHEN 2 THEN {1} ELSE 0 END) AS SFebrero," & _
            "SUM(CASE MONTH([FechaFactura]) WHEN 3 THEN {1} ELSE 0 END) AS SMarzo," & _
            "SUM(CASE MONTH([FechaFactura]) WHEN 4 THEN {1} ELSE 0 END) AS SAbril," & _
            "SUM(CASE MONTH([FechaFactura]) WHEN 5 THEN {1} ELSE 0 END) AS SMayo," & _
            "SUM(CASE MONTH([FechaFactura]) WHEN 6 THEN {1} ELSE 0 END) AS SJunio," & _
            "SUM(CASE MONTH([FechaFactura]) WHEN 7 THEN {1} ELSE 0 END) AS SJulio," & _
            "SUM(CASE MONTH([FechaFactura]) WHEN 8 THEN {1} ELSE 0 END) AS SAgosto," & _
            "SUM(CASE MONTH([FechaFactura]) WHEN 9 THEN {1} ELSE 0 END) AS SSeptiembre," & _
            "SUM(CASE MONTH([FechaFactura]) WHEN 10 THEN {1} ELSE 0 END) AS SOctubre," & _
            "SUM(CASE MONTH([FechaFactura]) WHEN 11 THEN {1} ELSE 0 END) AS SNoviembre," & _
            "SUM(CASE MONTH([FechaFactura]) WHEN 12 THEN {1} ELSE 0 END) AS SDiciembre," & _
            "SUM({1}) As STotalLinea", data.CamposSelect, data.CampoATotalizar))

        selectSQL.Append(" FROM tbMaestroCentroGestion RIGHT OUTER JOIN " & _
        "tbFacturaVentaLinea INNER JOIN " & _
        "vFacturaVentaCabecera ON tbFacturaVentaLinea.IDFactura = vFacturaVentaCabecera.IDFactura INNER JOIN " & _
        "tbMaestroCliente ON vFacturaVentaCabecera.IDCliente = tbMaestroCliente.IDCliente INNER JOIN " & _
        "tbMaestroArticulo ON tbFacturaVentaLinea.IDArticulo = tbMaestroArticulo.IDArticulo INNER JOIN " & _
        "tbMaestroPais ON tbMaestroCliente.IDPais = tbMaestroPais.IDPais INNER JOIN " & _
        "tbMaestroFamilia ON tbMaestroArticulo.IDFamilia = tbMaestroFamilia.IDFamilia AND " & _
        "tbMaestroArticulo.IDTipo = tbMaestroFamilia.IDTipo INNER JOIN " & _
        "tbMaestroTipoArticulo ON tbMaestroArticulo.IDTipo = tbMaestroTipoArticulo.IDTipo ON " & _
        "tbMaestroCentroGestion.IDCentroGestion = vFacturaVentaCabecera.IDCentroGestion LEFT OUTER JOIN " & _
        "tbObraCabecera ON tbFacturaVentaLinea.IDObra = tbObraCabecera.IDObra LEFT OUTER JOIN " & _
        "tbMaestroConcepto ON tbFacturaVentaLinea.IDConcepto = tbMaestroConcepto.IDConcepto LEFT OUTER JOIN " & _
        "tbMaestroActivo ON tbFacturaVentaLinea.Lote = tbMaestroActivo.IDActivo LEFT OUTER JOIN " & _
        "tbMaestroSubfamilia ON tbMaestroArticulo.IDTipo = tbMaestroSubfamilia.IDTipo AND " & _
        "tbMaestroArticulo.IDFamilia = tbMaestroSubfamilia.IDFamilia AND " & _
        "tbMaestroArticulo.IDSubfamilia = tbMaestroSubfamilia.IDSubfamilia LEFT OUTER JOIN " & _
        "tbMaestroZona ON tbMaestroCliente.IDZona = tbMaestroZona.IDZona LEFT OUTER JOIN " & _
        "tbMaestroMercado ON tbMaestroCliente.IDMercado = tbMaestroMercado.IDMercado LEFT OUTER JOIN " & _
        "tbMaestroCliente AS tbMaestroCliente_1 ON vFacturaVentaCabecera.IDGrupoCliente = tbMaestroCliente_1.IDCliente")

        Dim whereSQL As New Text.StringBuilder
        If data.Año.ToString.Length > 0 Then
            whereSQL.Append("YEAR(vFacturaVentaCabecera.FechaFactura) = " & data.Año & " AND ")
        End If
        If data.IDTipo.Length > 0 Then
            whereSQL.Append("tbMaestroArticulo.IDTipo = '" & data.IDTipo & "' AND ")
        End If
        If data.IDFamilia.Length > 0 Then
            whereSQL.Append("tbMaestroArticulo.IDFamilia = '" & data.IDFamilia & "' AND ")
        End If
        If data.IDSubFamilia.Length > 0 Then
            whereSQL.Append("tbMaestroArticulo.IDSubFamilia = '" & data.IDSubFamilia & "' AND ")
        End If
        If data.IDArticulo.Length > 0 Then
            whereSQL.Append("tbMaestroArticulo.IDArticulo = '" & data.IDArticulo & "' AND ")
        End If
        If data.IDCliente.Length > 0 Then
            whereSQL.Append("COALESCE(vFacturaVentaCabecera.IDClienteInicial,tbMaestroCliente.IDCliente) = '" & data.IDCliente & "' AND ")
        End If
        If data.IDGrupoCliente.Length > 0 Then
            whereSQL.Append("dbo.fGrupoCliente(vFacturaVentaCabecera.IDCliente, tbMaestroCliente.IDGrupoCliente, vFacturaVentaCabecera.IDClienteInicial, tbMaestroCliente_1.IDGrupoCliente) = '" & data.IDGrupoCliente & "' AND ")
        End If
        If data.Provincia.Length > 0 Then
            whereSQL.Append("tbMaestroCliente.Provincia = '" & data.Provincia & "' AND ")
        End If
        If data.IDZona.Length > 0 Then
            whereSQL.Append("tbMaestroCliente.IDZona = '" & data.IDZona & "' AND ")
        End If
        If data.IDMercado.Length > 0 Then
            whereSQL.Append("tbMaestroCliente.IDMercado = '" & data.IDMercado & "' AND ")
        End If
        If data.IDCentroGestion.Length > 0 Then
            whereSQL.Append("tbMaestroCentroGestion.IDCentroGestion = '" & data.IDCentroGestion & "' AND ")
        End If
        If data.IDPais.Length > 0 Then
            whereSQL.Append("tbMaestroPais.IDPais = '" & data.IDPais & "' AND ")
        End If

        Select Case data.CEE
            Case enumBoolean.Si
                whereSQL.Append("tbMaestroPais.CEE = 1 AND ")
            Case enumBoolean.No
                whereSQL.Append("tbMaestroPais.CEE = 0 AND ")
        End Select

        Select Case data.Extranjero
            Case enumBoolean.Si
                whereSQL.Append("tbMaestroPais.Extranjero = 1 AND ")
            Case enumBoolean.No
                whereSQL.Append("tbMaestroPais.Extranjero = 0 AND ")
        End Select

        Select Case data.EmpresaGrupo
            Case enumBoolean.Si
                whereSQL.Append("tbMaestroCliente.EmpresaGrupo = 1 AND ")
            Case enumBoolean.No
                whereSQL.Append("tbMaestroCliente.EmpresaGrupo = 0 AND ")
        End Select

        If whereSQL.Length > 0 Then
            selectSQL.Append(" WHERE ")
            selectSQL.Append(whereSQL.ToString.Substring(0, whereSQL.Length - 4))
        End If

        selectSQL.Append(" GROUP BY year([FechaFactura])")
        If Length(data.GroupBy) > 0 Then selectSQL.Append(", " & data.GroupBy)
        selectSQL.Append(" ORDER BY ")
        selectSQL.Append(data.CamposOrden)

        Dim cmdEstadisticas As Common.DbCommand = AdminData.GetCommand
        cmdEstadisticas.CommandType = CommandType.Text
        cmdEstadisticas.CommandText = selectSQL.ToString()
        Return AdminData.Execute(cmdEstadisticas, ExecuteCommand.ExecuteReader)
    End Function

    <Task()> Public Shared Function ADDColumnsAño2(ByVal data As DataTable, ByVal services As ServiceProvider) As DataTable
        data.Columns.Add("Año2", GetType(Integer))
        data.Columns.Add("SEnero2", GetType(Double))
        data.Columns.Add("SFebrero2", GetType(Double))
        data.Columns.Add("SMarzo2", GetType(Double))
        data.Columns.Add("SAbril2", GetType(Double))
        data.Columns.Add("SMayo2", GetType(Double))
        data.Columns.Add("SJunio2", GetType(Double))
        data.Columns.Add("SJulio2", GetType(Double))
        data.Columns.Add("SAgosto2", GetType(Double))
        data.Columns.Add("SSeptiembre2", GetType(Double))
        data.Columns.Add("SOctubre2", GetType(Double))
        data.Columns.Add("SNoviembre2", GetType(Double))
        data.Columns.Add("SDiciembre2", GetType(Double))
        data.Columns.Add("STotalLinea2", GetType(Double))

        Return data
    End Function

#End Region

    ' Método para IVA Venta 
    Public Function ObtenerDatosIVAVenta() As DataTable
        Return AdminData.Execute("SELECT AñoDeclaracionIVA, NDeclaracionIVA " & _
                                "FROM vCtlCIIVAVenta " & _
                                "WHERE NDeclaracionIVA IS NOT NULL OR AñoDeclaracionIVA IS NOT NULL " & _
                                "GROUP BY AñoDeclaracionIVA, NDeclaracionIVA " & _
                                "ORDER BY AñoDeclaracionIVA, NDeclaracionIVA", ExecuteCommand.ExecuteReader)
    End Function

    ' Registros marcados IVA Venta
    Public Function ObtenerRegistrosMarcadosIVAVenta(ByVal guidProceso As System.Guid) As DataTable
        Return New BE.DataEngine().Filter("vCtlCIIvaVenta", New GuidFilterItem("IDProcess", FilterOperator.Equal, guidProceso), , "NFactura")
    End Function
#End Region

    '#Region " Copiar Factura "

    '    <Task()> Public Shared Function CopiarFacturaVenta(ByVal intIDFactura As Integer, ByVal services As ServiceProvider) As DataTable
    '        If intIDFactura > 0 Then
    '            Dim FVC As New FacturaVentaCabecera
    '            Dim dtCabeceraOrigen As DataTable = FVC.SelOnPrimaryKey(intIDFactura)
    '            Dim f As New Filter
    '            f.Add(New NumberFilterItem("IDFactura", intIDFactura))
    '            Dim fvl As New FacturaVentaLinea
    '            Dim dtLineasOrigen As DataTable = fvl.Filter(f)

    '            Dim dtFVCD As DataTable = FVC.Filter(New NoRowsFilterItem)
    '            Dim dtFVLD As DataTable = fvl.Filter(New NoRowsFilterItem)

    '            Dim AppParamsConta As ParametroContabilidadVenta = services.GetService(Of ParametroContabilidadVenta)()

    '            'Copia Cabecera
    '            For Each drOrigenCabecera As DataRow In dtCabeceraOrigen.Rows
    '                Dim drDestinoCabecera As DataRow = dtFVCD.NewRow
    '                For Each dc As DataColumn In dtCabeceraOrigen.Columns
    '                    If dc.ColumnName <> "IDFactura" And dc.ColumnName <> "NFactura" And _
    '                       dc.ColumnName <> "FechaContabilizacion" And dc.ColumnName <> "IDObra" And _
    '                       dc.ColumnName <> "IDFacturaCompra" And dc.ColumnName <> "NDeclaracionIVA" And _
    '                       dc.ColumnName <> "AñoDeclaracionIva" And dc.ColumnName <> "NDeclaracionIntrastat" And _
    '                       dc.ColumnName <> "AñoDeclaracionIntrastat" And dc.ColumnName <> "DirecFacturaPDF" And _
    '                       dc.ColumnName <> "DirecFacturaXML" Then
    '                        drDestinoCabecera(dc.ColumnName) = drOrigenCabecera(dc)
    '                    End If
    '                Next

    '                drDestinoCabecera("IDFactura") = AdminData.GetAutoNumeric
    '                If Length(drDestinoCabecera("IDContador")) > 0 Then
    '                    Dim StDatos As New Contador.DatosCounterValue
    '                    StDatos.IDCounter = drDestinoCabecera("IDContador")
    '                    StDatos.TargetClass = FVC
    '                    StDatos.TargetField = "NFactura"
    '                    StDatos.DateField = "FechaFactura"
    '                    StDatos.DateValue = drDestinoCabecera("FechaFactura")
    '                    StDatos.IDEjercicio = drDestinoCabecera("IDEjercicio") & String.Empty
    '                    drDestinoCabecera("NFactura") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, StDatos, services)
    '                End If
    '                drDestinoCabecera("FechaFactura") = Date.Today
    '                drDestinoCabecera("FechaParaDeclaracion") = Date.Today

    '                If AppParamsConta.Contabilidad Then
    '                    Dim DataEjer As New DataEjercicio(New DataRowPropertyAccessor(drDestinoCabecera), Today.Date)
    '                    ProcessServer.ExecuteTask(Of DataEjercicio)(AddressOf NegocioGeneral.AsignarEjercicioContable, DataEjer, services)
    '                End If

    '                drDestinoCabecera("Estado") = enumfvcEstado.fvcNoContabilizado
    '                drDestinoCabecera("IVAManual") = 0
    '                drDestinoCabecera("VencimientosManuales") = 0
    '                drDestinoCabecera("GeneradoFichero") = 0
    '                drDestinoCabecera("EnviadaEntidadAseguradora") = 0
    '                drDestinoCabecera("Exportado") = 0
    '                drDestinoCabecera("Exportar") = 1

    '                dtFVCD.Rows.Add(drDestinoCabecera)
    '            Next

    '            'Copia Líneas
    '            For Each drOrigenLinea As DataRow In dtLineasOrigen.Rows
    '                Dim drDestinoLinea As DataRow = dtFVLD.NewRow
    '                For Each dc As DataColumn In dtLineasOrigen.Columns
    '                    If dc.ColumnName <> "IDLineaFactura" And dc.ColumnName <> "IDFactura" And _
    '                       dc.ColumnName <> "IDPedido" And dc.ColumnName <> "IDLineaPedido" And _
    '                       dc.ColumnName <> "IDAlbaran" And dc.ColumnName <> "IDLineaAlbaran" And _
    '                       dc.ColumnName <> "IDVencimiento" And dc.ColumnName <> "IDLineaVencimiento" And _
    '                       dc.ColumnName <> "IDObra" And dc.ColumnName <> "IDTrabajo" And _
    '                       dc.ColumnName <> "IDLineaMaterial" And dc.ColumnName <> "IDLineaMOD" And _
    '                       dc.ColumnName <> "IDLineaCentro" And dc.ColumnName <> "IDLineaGasto" And _
    '                       dc.ColumnName <> "IDLineaVarios" And dc.ColumnName <> "IDPromocionLinea" And _
    '                       dc.ColumnName <> "IDAlbaranRetorno" And dc.ColumnName <> "IDLineaAlbaranRetorno" And _
    '                       dc.ColumnName <> "IDLineaOfertaDetalle" And dc.ColumnName <> "IDCertificacion" Then
    '                        drDestinoLinea(dc.ColumnName) = drOrigenLinea(dc)
    '                    End If
    '                Next
    '                drDestinoLinea("IDFactura") = dtFVCD.Rows(0)("IDFactura")
    '                drDestinoLinea("IDLineaFactura") = AdminData.GetAutoNumeric
    '                dtFVLD.Rows.Add(drDestinoLinea)
    '            Next

    '            'FVC.BeginTx()
    '            FVC.Update(dtFVCD)
    '            fvl.Update(dtFVLD)
    '            'Updated(dtFVCD) 'Recalculo de la factura


    '            Return dtFVCD
    '        End If
    '    End Function


    '#End Region

#Region " CargarMedidasAB "

    <Serializable()> _
    Public Class DataExisteLinea
        Public Fact() As String
        Public IDLineaFactura As Integer
    End Class

    <Task()> Public Shared Function CargarMedidasAB(ByVal FilForm As Filter, ByVal services As ServiceProvider) As DataTable
        'Primero cargamos los datos de la vista cuyos articulos tengan configurado Articulos AB
        Dim DtArtAB As DataTable = New BE.DataEngine().Filter("vFrmCIFactVentaUD", FilForm)
        If Not DtArtAB Is Nothing AndAlso DtArtAB.Rows.Count > 0 Then
            Dim StrFact() As String
            Dim i As Integer = 0
            For Each Dr As DataRow In DtArtAB.Select("", "IDLineaFactura")
                Dim ExisteLinea As New DataExisteLinea
                ExisteLinea.Fact = StrFact
                ExisteLinea.IDLineaFactura = Dr("IDLineaFactura")
                Dim Existe As Boolean = ProcessServer.ExecuteTask(Of DataExisteLinea, Boolean)(AddressOf ExisteLineaFactura, ExisteLinea, services)
                If Not Existe Then
                    ReDim Preserve StrFact(i)
                    StrFact(i) = Dr("IDLineaFactura")
                    i += 1
                End If
            Next
            FilForm.Add(New InListFilterItem("IDLineaFactura", StrFact, FilterType.String, False))
            Dim DtArtConvAB As DataTable = New BE.DataEngine().Filter("vFrmCIFactVentaConverUD", FilForm)
            If Not DtArtConvAB Is Nothing AndAlso DtArtConvAB.Rows.Count > 0 Then
                For Each DrConv As DataRow In DtArtConvAB.Select("", "IDLineaFactura")
                    DtArtAB.Rows.Add(DrConv.ItemArray)
                Next
                DtArtAB.AcceptChanges()
            End If
            Return DtArtAB
        Else
            Dim DtArtConvAB As DataTable = New BE.DataEngine().Filter("vFrmCIFactVentaConverUD", FilForm)
            If Not DtArtConvAB Is Nothing AndAlso DtArtConvAB.Rows.Count > 0 Then
                Return DtArtConvAB
            End If
        End If
    End Function

    <Task()> Public Shared Function ExisteLineaFactura(ByVal ExisteLinea As DataExisteLinea, ByVal services As ServiceProvider) As Boolean
        If Not ExisteLinea.Fact Is Nothing AndAlso ExisteLinea.Fact.Length > 0 Then
            For i As Integer = 0 To ExisteLinea.Fact.Length - 1
                If ExisteLinea.Fact(i) = ExisteLinea.IDLineaFactura Then Return True
            Next
            Return False
        Else
            Return False
        End If
    End Function

#End Region

#Region " TPV: Factura para TPV y Ticket "

    <Serializable()> _
   Public Class DataResultGenerarFacturaTPV
        Public FacturaTPV As DataTable
        Public FormasCobro As DataTable
    End Class

    <Task()> Public Shared Function GenerarFacturaTPV(ByVal IDAlbaranPasado As Integer, ByVal services As ServiceProvider) As DataResultGenerarFacturaTPV
        Dim dataResult As New DataResultGenerarFacturaTPV
        'vista sqlServer: vRptFacturaTPV
        dataResult.FacturaTPV = New BE.DataEngine().Filter("vRptFacturaTPV", New FilterItem("IDAlbaran", FilterOperator.Equal, IDAlbaranPasado))

        'vista sqlServer: vRptFacturaTPVFormaPago

        dataResult.FormasCobro = New BE.DataEngine().Filter("vRptFacturaTPVFormaPago", New FilterItem("IDAlbaran", FilterOperator.Equal, IDAlbaranPasado))
        If Not dataResult.FacturaTPV Is Nothing AndAlso dataResult.FacturaTPV.Rows.Count > 0 Then
            'comprobamos si datos nulos
            If Length(dataResult.FacturaTPV.Rows(0)("DescEmpresa")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("DescEmpresa") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("Direccion")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("Direccion") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("CodPostal")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("CodPostal") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("Poblacion")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("Poblacion") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("Provincia")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("Provincia") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("Cif")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("Cif") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("RazonSocialClte")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("RazonSocialClte") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("DireccionClte")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("DireccionClte") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("CodPostalClte")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("CodPostalClte") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("PoblacionClte")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("PoblacionClte") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("ProvinciaClte")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("ProvinciaClte") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("CifCliente")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("CifCliente") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("NFactura")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("NFactura") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("ImpLineas")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("ImpLineas") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("ImpIva")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("ImpIva") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("DescArticulo")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("DescArticulo") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("Cantidad")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("Cantidad") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("Precio")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("Precio") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("Importe")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("Importe") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("DescTipoIva")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("DescTipoIva") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("ImpTotal")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("ImpTotal") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("DescFormaPago")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("DescFormaPago") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("DescBancoPropio")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("DescBancoPropio") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("IDCliente")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("IDCliente") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("IDArticulo")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("IDArticulo") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("Dto1")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("Dto1") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("Dto2")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("Dto2") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("Dto3")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("Dto3") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("Dto")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("Dto") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("IDAlbaran")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("IDAlbaran") = String.Empty
            End If
            If Length(dataResult.FacturaTPV.Rows(0)("IDFactura")) = 0 Then
                dataResult.FacturaTPV.Rows(0)("IDFactura") = String.Empty
            End If
            If dataResult.FormasCobro.Rows.Count > 0 Then
                If Length(dataResult.FormasCobro.Rows(0)("ImpVencimiento")) = 0 Then
                    dataResult.FormasCobro.Rows(0)("ImpVencimiento") = String.Empty
                End If
                If Length(dataResult.FormasCobro.Rows(0)("IDFormaPago")) = 0 Then
                    dataResult.FormasCobro.Rows(0)("IDFormaPago") = String.Empty
                End If
                If Length(dataResult.FormasCobro.Rows(0)("DescFormaPago")) = 0 Then
                    dataResult.FormasCobro.Rows(0)("DescFormaPago") = String.Empty
                End If
                If Length(dataResult.FormasCobro.Rows(0)("IDFactura")) = 0 Then
                    dataResult.FormasCobro.Rows(0)("IDFactura") = String.Empty
                End If
                If Length(dataResult.FormasCobro.Rows(0)("NFactura")) = 0 Then
                    dataResult.FormasCobro.Rows(0)("NFactura") = String.Empty
                End If
                If Length(dataResult.FormasCobro.Rows(0)("FechaVencimiento")) = 0 Then
                    dataResult.FormasCobro.Rows(0)("FechaVencimiento") = String.Empty
                End If
                If Length(dataResult.FormasCobro.Rows(0)("IDBancoPropio")) = 0 Then
                    dataResult.FormasCobro.Rows(0)("IDBancoPropio") = String.Empty
                End If
                If Length(dataResult.FormasCobro.Rows(0)("DescBancoPropio")) = 0 Then
                    dataResult.FormasCobro.Rows(0)("DescBancoPropio") = String.Empty
                End If
            End If
        End If
        Return dataResult
    End Function

    <Serializable()> _
    Public Class DataResultGenerarTicketTPV
        Public TickectTPV As DataTable
        Public FormasCobro As DataTable
        Public TipoIVA As DataTable
        Public PuntosVenta As Integer
        Public PuntosTotales As Integer
    End Class

    <Serializable()> _
    Public Class DataGeneraTicketTPV
        Public IDAlbaranPasado As Integer
        Public IDTarjetaFidelizacion As String
    End Class

    <Task()> Public Shared Function GenerarTicketTPV(ByVal data As DataGeneraTicketTPV, ByVal services As ServiceProvider) As DataResultGenerarTicketTPV
        Dim dataResult As New DataResultGenerarTicketTPV
        'vista sqlServer: vRptAlbaranVenta
        Dim FilVta As New Filter
        FilVta.Add("IDAlbaran", FilterOperator.Equal, data.IDAlbaranPasado)
        FilVta.Add(New IsNullFilterItem("IDLineaPadre"))
        dataResult.TickectTPV = New BE.DataEngine().Filter("vRptAlbaranVta", FilVta)

        'vista sqlServer: vRptAlbaranVtaFormaPago
        dataResult.FormasCobro = New BE.DataEngine().Filter("vRptAlbaranVtaFormaPago", New FilterItem("IDAlbaran", FilterOperator.Equal, data.IDAlbaranPasado))

        'vista sqlServer: vRptAlbaranTipoIVA
        dataResult.TipoIVA = New BE.DataEngine().Filter("vRptAlbaranTipoIVA", New FilterItem("IDAlbaran", FilterOperator.Equal, data.IDAlbaranPasado))

        If Not dataResult.TickectTPV Is Nothing AndAlso dataResult.TickectTPV.Rows.Count > 0 Then
            'comprobamos si datos nulos
            If Length(dataResult.TickectTPV.Rows(0)("DescEmpresa")) = 0 Then
                dataResult.TickectTPV.Rows(0)("DescEmpresa") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("Direccion")) = 0 Then
                dataResult.TickectTPV.Rows(0)("Direccion") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("Telefono")) = 0 Then
                dataResult.TickectTPV.Rows(0)("Telefono") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("Cif")) = 0 Then
                dataResult.TickectTPV.Rows(0)("PVP") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("DescCentroGestion")) = 0 Then
                dataResult.TickectTPV.Rows(0)("DescCentroGestion") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("DescMoneda")) = 0 Then
                dataResult.TickectTPV.Rows(0)("DescMoneda") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("DescVendedor")) = 0 Then
                dataResult.TickectTPV.Rows(0)("DescVendedor") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("DescArticulo")) = 0 Then
                dataResult.TickectTPV.Rows(0)("DescArticulo") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("QServida")) = 0 Then
                dataResult.TickectTPV.Rows(0)("QServida") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("ImportePVP")) = 0 Then
                dataResult.TickectTPV.Rows(0)("ImportePVP") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("Importe")) = 0 Then
                dataResult.TickectTPV.Rows(0)("Importe") = 0
            End If
            If Length(dataResult.TickectTPV.Rows(0)("ImpIva")) = 0 Then
                dataResult.TickectTPV.Rows(0)("ImpIva") = 0
            End If
            If Length(dataResult.TickectTPV.Rows(0)("ImpTotal")) = 0 Then
                dataResult.TickectTPV.Rows(0)("ImpTotal") = 0
            End If
            If Length(dataResult.TickectTPV.Rows(0)("NAlbaran")) = 0 Then
                dataResult.TickectTPV.Rows(0)("NAlbaran") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("ImportePVPA")) = 0 Then
                dataResult.TickectTPV.Rows(0)("ImportePVPA") = 0
            End If

            If Length(dataResult.TickectTPV.Rows(0)("ImpAlbaran")) = 0 Then
                dataResult.TickectTPV.Rows(0)("ImpAlbaran") = 0
            End If
            If Length(dataResult.TickectTPV.Rows(0)("DescTipoIva")) = 0 Then
                dataResult.TickectTPV.Rows(0)("DescTipoIva") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("IDAlbaran")) = 0 Then
                dataResult.TickectTPV.Rows(0)("IDAlbaran") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("Precio")) = 0 Then
                dataResult.TickectTPV.Rows(0)("Precio") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("ImpLinea")) = 0 Then
                dataResult.TickectTPV.Rows(0)("ImpLinea") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("PVP")) = 0 Then
                dataResult.TickectTPV.Rows(0)("PVP") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("IDFormaPago")) = 0 Then
                dataResult.TickectTPV.Rows(0)("IDFormaPago") = String.Empty
            End If
            If dataResult.TipoIVA.Rows.Count > 0 Then
                If Length(dataResult.TipoIVA.Rows(0)("IDAlbaran")) = 0 Then
                    dataResult.TipoIVA.Rows(0)("IDAlbaran") = String.Empty
                End If
                If Length(dataResult.TipoIVA.Rows(0)("NAlbaran")) = 0 Then
                    dataResult.TipoIVA.Rows(0)("NAlbaran") = String.Empty
                End If

                If Length(dataResult.TipoIVA.Rows(0)("Factor")) = 0 Then
                    dataResult.TipoIVA.Rows(0)("Factor") = String.Empty
                End If
                If Length(dataResult.TipoIVA.Rows(0)("Importe")) = 0 Then
                    dataResult.TipoIVA.Rows(0)("Importe") = String.Empty
                End If
            End If
            If dataResult.FormasCobro.Rows.Count > 0 Then
                If Length(dataResult.FormasCobro.Rows(0)("IDAlbaran")) = 0 Then
                    dataResult.FormasCobro.Rows(0)("IDAlbaran") = String.Empty
                End If
                If Length(dataResult.FormasCobro.Rows(0)("NAlbaran")) = 0 Then
                    dataResult.FormasCobro.Rows(0)("NAlbaran") = String.Empty
                End If

                If Length(dataResult.FormasCobro.Rows(0)("IDFormaPago")) = 0 Then
                    dataResult.FormasCobro.Rows(0)("IDFormaPago") = String.Empty
                End If
                If Length(dataResult.FormasCobro.Rows(0)("DescFormaPago")) = 0 Then
                    dataResult.FormasCobro.Rows(0)("DescFormaPago") = String.Empty
                End If
                If Length(dataResult.FormasCobro.Rows(0)("Importe")) = 0 Then
                    dataResult.FormasCobro.Rows(0)("Importe") = String.Empty
                End If
            End If

            If Length(data.IDTarjetaFidelizacion) > 0 Then
                Dim DtPuntosVenta As DataTable = New BE.DataEngine().Filter("vFrmMntoAlbaranVentaPuntos", New FilterItem("IDAlbaran", FilterOperator.Equal, data.IDAlbaranPasado))
                If Not DtPuntosVenta Is Nothing AndAlso DtPuntosVenta.Rows.Count > 0 Then
                    dataResult.PuntosVenta = DtPuntosVenta.Compute("SUM(PuntosMarketing)", Nothing)
                End If
                Dim DtPuntosTotales As DataTable = New BE.DataEngine().Filter("vFrmMntoAlbaranVentaPuntos", New FilterItem("IDTarjetaFidelizacion", FilterOperator.Equal, data.IDTarjetaFidelizacion))
                If Not DtPuntosTotales Is Nothing AndAlso DtPuntosTotales.Rows.Count > 0 Then
                    dataResult.PuntosTotales = DtPuntosTotales.Compute("SUM(PuntosMarketing)", Nothing) - DtPuntosTotales.Compute("SUM(PuntosUtilizados)", Nothing)
                End If
            Else : dataResult.PuntosTotales = 0 : dataResult.PuntosVenta = 0
            End If
        End If
        Return dataResult
    End Function

    <Serializable()> _
    Public Class DataResultGenerarTicketTPVVale
        Public TickectTPV As DataTable
        Public DatosVale As DataTable
    End Class

    <Task()> Public Shared Function GenerarTicketVale(ByVal data As DataGeneraTicketTPV, ByVal services As ServiceProvider) As DataResultGenerarTicketTPVVale
        Dim dataResult As New DataResultGenerarTicketTPVVale
        'vista sqlServer: vRptAlbaranVenta
        dataResult.TickectTPV = New BE.DataEngine().Filter("vRptAlbaranVta", New FilterItem("IDAlbaran", FilterOperator.Equal, data.IDAlbaranPasado))

        dataResult.DatosVale = New BE.DataEngine().Filter("vRptTicketVale", New FilterItem("IDAlbaran", FilterOperator.Equal, data.IDAlbaranPasado))

        If Not dataResult.TickectTPV Is Nothing AndAlso dataResult.TickectTPV.Rows.Count > 0 Then
            'comprobamos si datos nulos
            If Length(dataResult.TickectTPV.Rows(0)("DescEmpresa")) = 0 Then
                dataResult.TickectTPV.Rows(0)("DescEmpresa") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("Direccion")) = 0 Then
                dataResult.TickectTPV.Rows(0)("Direccion") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("Telefono")) = 0 Then
                dataResult.TickectTPV.Rows(0)("Telefono") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("Cif")) = 0 Then
                dataResult.TickectTPV.Rows(0)("PVP") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("DescCentroGestion")) = 0 Then
                dataResult.TickectTPV.Rows(0)("DescCentroGestion") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("DescMoneda")) = 0 Then
                dataResult.TickectTPV.Rows(0)("DescMoneda") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("DescVendedor")) = 0 Then
                dataResult.TickectTPV.Rows(0)("DescVendedor") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("DescArticulo")) = 0 Then
                dataResult.TickectTPV.Rows(0)("DescArticulo") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("QServida")) = 0 Then
                dataResult.TickectTPV.Rows(0)("QServida") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("ImportePVP")) = 0 Then
                dataResult.TickectTPV.Rows(0)("ImportePVP") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("Importe")) = 0 Then
                dataResult.TickectTPV.Rows(0)("Importe") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("ImpIva")) = 0 Then
                dataResult.TickectTPV.Rows(0)("ImpIva") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("ImpTotal")) = 0 Then
                dataResult.TickectTPV.Rows(0)("ImpTotal") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("NAlbaran")) = 0 Then
                dataResult.TickectTPV.Rows(0)("NAlbaran") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("ImportePVPA")) = 0 Then
                dataResult.TickectTPV.Rows(0)("ImportePVPA") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("ImpAlbaran")) = 0 Then
                dataResult.TickectTPV.Rows(0)("ImpAlbaran") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("DescTipoIva")) = 0 Then
                dataResult.TickectTPV.Rows(0)("DescTipoIva") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("IDAlbaran")) = 0 Then
                dataResult.TickectTPV.Rows(0)("IDAlbaran") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("Precio")) = 0 Then
                dataResult.TickectTPV.Rows(0)("Precio") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("ImpLinea")) = 0 Then
                dataResult.TickectTPV.Rows(0)("ImpLinea") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("PVP")) = 0 Then
                dataResult.TickectTPV.Rows(0)("PVP") = String.Empty
            End If
            If Length(dataResult.TickectTPV.Rows(0)("IDFormaPago")) = 0 Then
                dataResult.TickectTPV.Rows(0)("IDFormaPago") = String.Empty
            End If
        End If
        Return dataResult
    End Function

#End Region

#Region " Facturación Electrónica"
    <Serializable()> _
    Public Class DirectFactElecInfo
        Public StrCampoDirec As String
        Public DtFact As DataTable
        Public StrRuta As String
        Public StrCampoCorreo As String
        Public StrCorreo As String
    End Class

    <Task()> Public Shared Sub ActualizarDirecFactElec(ByVal data As DirectFactElecInfo, ByVal services As ServiceProvider)
        For Each Dr As DataRow In data.DtFact.Select
            Dr(data.StrCampoDirec) = data.StrRuta
            Dr(data.StrCampoCorreo) = data.StrCorreo
        Next
        BusinessHelper.UpdateTable(data.DtFact)
    End Sub

#End Region

    <Task()> Public Shared Sub ActualizarFicheroGeneradoEDI(ByVal IDFactura() As Object, ByVal services As ServiceProvider)
        If IDFactura.Length > 0 Then
            Dim dt As DataTable = New FacturaVentaCabecera().Filter(New InListFilterItem("IDFactura", IDFactura, FilterType.Numeric, True))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    dr("GeneradoFichero") = True
                Next
                BusinessHelper.UpdateTable(dt)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub QuitarFacturaElectronica(ByVal IDFactura As Integer, ByVal services As ServiceProvider)
        Dim DtFact As DataTable = New FacturaVentaCabecera().SelOnPrimaryKey(IDFactura)
        DtFact.Rows(0)("DirecFacturaPDF") = DBNull.Value
        BusinessHelper.UpdateTable(DtFact)
    End Sub

    <Task()> Public Shared Sub QuitarFacturaElectronicaXML(ByVal IDFactura As Integer, ByVal services As ServiceProvider)
        Dim DtFact As DataTable = New FacturaVentaCabecera().SelOnPrimaryKey(IDFactura)
        DtFact.Rows(0)("DirecFacturaXML") = DBNull.Value
        BusinessHelper.UpdateTable(DtFact)
    End Sub

    <Task()> Public Shared Function GetParamsFacturaVenta(ByVal data As Object, ByVal services As ServiceProvider) As DataParamFacturaVenta
        Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
        Dim AppParamsGeneral As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim AppParamsStock As ParametroStocks = services.GetService(Of ParametroStocks)()

        Dim datParams As New DataParamFacturaVenta
        datParams.GestionAnalitica = AppParamsConta.Analitica.AplicarAnalitica
        datParams.GestionAlquiler = AppParamsGeneral.AplicacionGestionAlquiler
        datParams.ExpertisSAAS = AppParamsGeneral.SAAS
        datParams.Contabilidad = AppParamsConta.Contabilidad
        datParams.GestionDobleUnidad = AppParamsStock.GestionDobleUnidad
        datParams.GAIANetExchange = New Parametro().GAIANetExchange

        datParams.TipoLineaPredeterminado = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, services)

        datParams.MonInfoA = Monedas.MonedaA
        datParams.MonInfoB = Monedas.MonedaB

        Return datParams
    End Function

#Region " Gestor Cobro "

    <Serializable()> _
    Public Class dataAddGestorCobro
        Public IDFactura() As Object
        Public FechaComunicacion As Date

        Friend IDProveedor As String
        Friend IDOperario As String

        Public Sub New(ByVal IDFactura() As Object, ByVal IDGestor As String, ByVal GestorInterno As Boolean, ByVal FechaComunicacion As Date)
            Me.IDFactura = IDFactura
            Me.FechaComunicacion = FechaComunicacion
            If GestorInterno Then
                Me.IDOperario = IDGestor
            Else
                Me.IDProveedor = IDGestor
            End If
        End Sub
    End Class
    <Task()> Public Shared Sub AddGestorCobro(ByVal data As dataAddGestorCobro, ByVal services As ServiceProvider)
        If data.IDFactura.Length > 0 Then
            Dim dtFactura As DataTable = New FacturaVentaCabecera().Filter(New InListFilterItem("IDFactura", data.IDFactura, FilterType.Numeric))
            If dtFactura.Rows.Count > 0 Then
                For Each drFactura As DataRow In dtFactura.Rows
                    drFactura("IDProveedor") = data.IDProveedor
                    drFactura("IDOperario") = data.IDOperario
                    drFactura("FechaComunicacionGestorCobro") = data.FechaComunicacion
                    drFactura("ComunicadoGestorCobro") = True
                Next
            End If
            AdminData.BeginTx()
            FacturaVentaCabecera.UpdateTable(dtFactura)
            ProcessServer.ExecuteTask(Of dataAddGestorCobro)(AddressOf ActualizarCobro, data, services)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarCobro(ByVal data As dataAddGestorCobro, ByVal services As ServiceProvider)
        If data.IDFactura.Length > 0 Then
            Dim dtCobro As DataTable = New Cobro().Filter(New InListFilterItem("IDFactura", data.IDFactura, FilterType.Numeric))
            If dtCobro.Rows.Count > 0 Then
                For Each drCobro As DataRow In dtCobro.Rows
                    drCobro("IDProveedor") = data.IDProveedor
                    drCobro("IDOperario") = data.IDOperario
                    drCobro("FechaComunicacionGestorCobro") = data.FechaComunicacion
                    drCobro("ComunicadoGestorCobro") = True
                Next
                Cobro.UpdateTable(dtCobro)
            End If
        End If
    End Sub

#End Region

#Region " Actualización masiva Facturas Venta "

    <Task()> Public Shared Sub ActualizarMarcasEnviar349(ByVal data As DataTable, ByVal services As ServiceProvider)
        For Each drMarcados As DataRow In data.Rows
            If drMarcados("Enviar349") Then
                drMarcados("Enviar349") = False
            Else
                drMarcados("Enviar349") = True
            End If
        Next
        BusinessHelper.UpdateTable(data)
    End Sub

    <Task()> Public Shared Sub ActualizarMarcasServicios349(ByVal data As DataTable, ByVal services As ServiceProvider)
        For Each drMarcados As DataRow In data.Rows
            If drMarcados("Servicios349") Then
                drMarcados("Servicios349") = False
            Else
                drMarcados("Servicios349") = True
            End If
        Next
        BusinessHelper.UpdateTable(data)
    End Sub

    <Task()> Public Shared Sub ActualizarMarcasEnviar347(ByVal data As DataTable, ByVal services As ServiceProvider)
        For Each drMarcados As DataRow In data.Rows
            If drMarcados("Enviar347") Then
                drMarcados("Enviar347") = False
            Else
                drMarcados("Enviar347") = True
            End If
        Next
        BusinessHelper.UpdateTable(data)
    End Sub

    <Task()> Public Shared Sub ActualizarMarcasOpeTriangular(ByVal data As DataTable, ByVal services As ServiceProvider)
        For Each drMarcados As DataRow In data.Rows
            If drMarcados("OpeTriangular") Then
                drMarcados("OpeTriangular") = False
            Else
                drMarcados("OpeTriangular") = True
            End If
        Next
        BusinessHelper.UpdateTable(data)
    End Sub

#End Region

End Class


