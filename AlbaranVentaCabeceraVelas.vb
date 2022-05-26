﻿Public Class AlbaranVentaCabeceraVelas
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Private Const cnEntidad As String = "tbAlbaranVentaCabecera"

    Private _AVC As _AlbaranVentaCabecera
    Private _AVL As _AlbaranVentaLinea
    Private _PVL As _PedidoVentaLinea
    Private _PVC As _PedidoVentaCabecera
    Private _AA As _ArticuloAlmacen
    Private _AAL As _ArticuloAlmacenLote
    Private _AVLT As _AlbaranVentaLote

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

#Region " RegisterAddNewTasks "

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoComunes.AsignarIdentificadorAlbaran, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.AsignarCentroGestion, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarContador, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarNumeroAlbaranProvisional, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.AsignarAlmacen, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoComunes.AsignarFechaAlbaran, data, services)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoComunes.AsignarEjercicioContableAlbaran, New DataRowPropertyAccessor(data), services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.AsignarTipoAlbaran, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarEstado, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarAparcado, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarArqueoCaja, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarResponsableExpedicion, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarAutomatico, data, services)
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim CE As New CentroEntidad
        CE.CentroGestion = data("IDCentroGestion").ToString & String.Empty
        CE.ContadorEntidad = CentroGestion.ContadorEntidad.AlbaranVenta
        data("IDContador") = ProcessServer.ExecuteTask(Of CentroEntidad, String)(AddressOf CentroGestion.GetContadorPredeterminado, CE, services)
    End Sub

    <Task()> Public Shared Sub AsignarNumeroAlbaranProvisional(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDContador")) > 0 Then
            Dim dtContadores As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDt, "AlbaranVentaCabecera", services)
            Dim adr As DataRow() = dtContadores.Select("IDContador = " & Quoted(data("IDContador")))
            If Not IsNothing(adr) AndAlso adr.Length > 0 Then
                data("NAlbaran") = adr(0)("ValorProvisional")
            Else
                'Si no está bien configurado el Contador de Pedidos de Venta en el Centro de Gestión,
                'cogemos el Contador por defecto de la entidad Albaran Venta Cabecera.
                Dim dtContadorPred As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDefault, GetType(AlbaranVentaCabecera).Name, services)
                If Not dtContadorPred Is Nothing AndAlso dtContadorPred.Rows.Count > 0 Then
                    data("IDContador") = dtContadorPred.Rows(0)("IDContador")
                    adr = dtContadores.Select("IDContador = " & Quoted(data("IDContador")))
                    If Not IsNothing(adr) AndAlso adr.Length > 0 Then
                        data("NAlbaran") = adr(0)("ValorProvisional")
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarEstado(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.IsNull("Estado") Then data("Estado") = enumavcEstadoFactura.avcNoFacturado
    End Sub
    <Task()> Public Shared Sub AsignarAparcado(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("Aparcado") = False
    End Sub
    <Task()> Public Shared Sub AsignarArqueoCaja(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("Arqueo") = False
    End Sub
    <Task()> Public Shared Sub AsignarResponsableExpedicion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("ResponsableExpedicion")) = 0 Then
            Dim strIDOper As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf Operario.ObtenerIDOperarioUsuario, Nothing, services)
            If Len(strIDOper) > 0 Then data("ResponsableExpedicion") = strIDOper
        End If
    End Sub
    <Task()> Public Shared Sub AsignarAutomatico(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("Automatico") = False
    End Sub

#End Region
#Region " Delete "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ValidarDelAlbaranFacturado)
        deleteProcess.AddTask(Of DataRow)(AddressOf ValidarDelConDAA)
        deleteProcess.AddTask(Of DataRow)(AddressOf ValidarDelAlbaranAbonado)
        deleteProcess.AddTask(Of DataRow)(AddressOf ValidarDelAlbaranRetorno)
        deleteProcess.AddTask(Of DataRow)(AddressOf ValidarDelAlbaranConContador)
        deleteProcess.AddTask(Of DataRow)(AddressOf ProcesoAlbaranVenta.ActualizarAlbaranesMultiEmpresa)
        deleteProcess.AddTask(Of DataRow)(AddressOf ProcesoAlbaranVenta.ValidarAlbaranArqueoCaja)
        deleteProcess.AddTask(Of DataRow)(AddressOf ProcesoAlbaranVenta.ValidarAlbaranContado)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarAlbaranExpedicionDistribuidor)
    End Sub

    <Task()> Public Shared Sub ValidarDelConDAA(ByVal DocHeaderRow As DataRow, ByVal services As ServiceProvider)
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If Not AppParams.GestionBodegas Then Exit Sub
        If Not IsDBNull(DocHeaderRow("IDDAA")) AndAlso Not CType(DocHeaderRow("IDDAA"), Guid).Equals(Guid.Empty) Then
            ApplicationService.GenerateError("No se puede borrar el Albarán. Está asociado a un DAA.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDelAlbaranFacturado(ByVal DocHeaderRow As DataRow, ByVal services As ServiceProvider)
        If DocHeaderRow("Estado") = enumavcEstadoFactura.avcParcFacturado OrElse _
           DocHeaderRow("Estado") = enumavcEstadoFactura.avcFacturado Then
            ApplicationService.GenerateError("No se puede borrar el Albarán. Está Facturado o Parcialmente Facturado.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDelAlbaranAbonado(ByVal DocHeaderRow As DataRow, ByVal services As ServiceProvider)
        If Length(DocHeaderRow("IDTipoAlbaran")) > 0 Then
            Dim TipoAlbInfo As ProcesoAlbaranVenta.TipoAlbaranInfo = ProcessServer.ExecuteTask(Of String, ProcesoAlbaranVenta.TipoAlbaranInfo)(AddressOf ProcesoAlbaranVenta.TipoDeAlbaran, DocHeaderRow("IDTipoAlbaran"), services)
            If TipoAlbInfo.Tipo = enumTipoAlbaran.ExpedDistribuidor AndAlso Nz(DocHeaderRow("IDAlbaranAbono"), 0) <> 0 Then
                ApplicationService.GenerateError("No se puede borrar el Albarán. Tiene un Albarán de Abono vinculado.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDelAlbaranRetorno(ByVal DocHeaderRow As DataRow, ByVal services As ServiceProvider)
        Dim AppParamsVenta As ParametroVenta = services.GetService(Of ParametroVenta)()
        If AppParamsVenta.General.AplicacionGestionAlquiler Then
            Dim dtAlbRetorno As DataTable = AdminData.GetData("vAlbaranConRetorno", New NumberFilterItem("IDAlbaran", DocHeaderRow("IDAlbaran")))
            If Not dtAlbRetorno Is Nothing AndAlso dtAlbRetorno.Rows.Count > 0 Then
                ApplicationService.GenerateError("No se puede borrar el Albarán de Depósito, tiene un Albarán de Retorno asociado.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDelAlbaranConContador(ByVal DocHeaderRow As DataRow, ByVal services As ServiceProvider)
        Dim AppParamsVenta As ParametroVenta = services.GetService(Of ParametroVenta)()
        If AppParamsVenta.General.AplicacionGestionAlquiler Then
            Dim AppParamsAlb As ParametroAlbaranVenta = services.GetService(Of ParametroAlbaranVenta)()
            Dim strVista As String
            If AppParamsAlb.TipoAlbaranRetornoAlquiler = DocHeaderRow("IDTipoAlbaran") Then
                strVista = "VnegAlbaranesConContadorRetornos"
            Else
                strVista = "VnegAlbaranesConContadorEnvios"
            End If

            Dim dtPrevCont As DataTable = AdminData.GetData(strVista, New NumberFilterItem("IDAlbaran", DocHeaderRow("IDAlbaran")))
            If Not dtPrevCont Is Nothing AndAlso dtPrevCont.Rows.Count > 0 Then
                ApplicationService.GenerateError("No se puede borrar este Registro. La Línea está asociada a un Contador.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarAlbaranExpedicionDistribuidor(ByVal DocHeaderRow As DataRow, ByVal services As ServiceProvider)
        If Length(DocHeaderRow("IDTipoAlbaran")) > 0 Then
            Dim TipoAlbInfo As ProcesoAlbaranVenta.TipoAlbaranInfo = ProcessServer.ExecuteTask(Of String, ProcesoAlbaranVenta.TipoAlbaranInfo)(AddressOf ProcesoAlbaranVenta.TipoDeAlbaran, DocHeaderRow("IDTipoAlbaran"), services)
            If TipoAlbInfo.Tipo = enumTipoAlbaran.AbonoDistribuidor Then
                Dim fAVExped As New Filter
                fAVExped.Add(New NumberFilterItem("IDAlbaranAbono", DocHeaderRow("IDAlbaran")))
                Dim dtAVExped As DataTable = New AlbaranVentaCabecera().Filter(fAVExped)
                If dtAVExped.Rows.Count > 0 Then
                    For Each dr As DataRow In dtAVExped.Rows
                        dr("IDAlbaranAbono") = System.DBNull.Value
                    Next
                End If
                BusinessHelper.UpdateTable(dtAVExped)
            End If
        End If
    End Sub

#End Region
#Region " Validate "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoAlbaranVenta.ValidarAlbaranFacturado)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarClienteObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarClienteBloqueado)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFechaAlbaranObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarMonedaObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarAlbaranObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarTipoAlbaranObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoAlbaranVenta.ValidarFacturasPorCondicionEnvio)
        ' validateProcess.AddTask(Of DataRow)(AddressOf ProcesoAlbaranVenta.ValidarNumeroAlbaran)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoAlbaranVenta.ValidarCondicionesEconomicas)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoAlbaranVenta.ValidacionesContabilidad)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarFechaAlbaranAnterior)
    End Sub

    <Task()> Public Shared Sub ValidarFechaAlbaranAnterior(ByVal data As DataRow, ByVal services As ServiceProvider)
        If New Parametro().ValidarCambioFechaFacturas Then
            Dim FilFacturas As New Filter
            FilFacturas.Add("IDAlbaran", FilterOperator.NotEqual, data("IDAlbaran"))
            FilFacturas.Add("IDContador", FilterOperator.Equal, data("IDContador"))
            FilFacturas.Add("IDEjercicio", FilterOperator.Equal, data("IDEjercicio"))
            FilFacturas.Add("FechaAlbaran", FilterOperator.GreaterThan, data("FechaAlbaran"))
            Dim DtAlbaranes As DataTable = New AlbaranVentaCabecera().Filter(FilFacturas)
            If Not DtAlbaranes Is Nothing AndAlso DtAlbaranes.Rows.Count > 0 Then
                If data.RowState = DataRowState.Added Then
                    ApplicationService.GenerateError("No se puede generar el Albarán con la fecha introducida. Existen albaranes generados posteriores a la fecha.")
                ElseIf data.RowState = DataRowState.Modified AndAlso Nz(data("FechaAlbaran")) <> Nz(data("FechaAlbaran", DataRowVersion.Original)) Then
                    ApplicationService.GenerateError("No se puede modificar la fecha de albarán con la fecha introducida. Existen albaranes generados posteriores a la fecha.")
                End If
            End If
        End If
    End Sub

#End Region
#Region " Update "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)

        updateProcess.AddTask(Of UpdatePackage, DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.CrearDocumento)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.AsignarCentroGestion)
        'updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.ValidarDocumento)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.TratarTipoAlbaran)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.AsignarNumeroAlbaran)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.ResponsableExpedicion)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.ActualizarCambiosMoneda)

        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.AsignarDireccionFacturaEnLineas)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.AsignarClienteBancoEnLineas)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.AsignarEstadoFacturaEnLineas)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.AsignarCondicionesEnLineas)

        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.GestionArticulosKit)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.GestionArticulosFantasma)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.ActualizarEstadoLineas)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.ActualizarEstadoAlbaran)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.TratarPromocionesLineas)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.CalcularImporteLineasAlbaran)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaPedidos.CalcularRepresentantes)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaPedidos.CalcularAnalitica)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.CalcularBasesImponibles)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.TotalDocumento)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.TotalPesos)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf Business.General.Comunes.BeginTransaction)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.CorregirMovimientos)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaPedidos.ActualizarPedidoDesdeAlbaran)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaObras.ActualizarObrasDesdeAlbaran)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.ActualizarPuntosMarketing)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf Business.General.Comunes.UpdateDocument)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf Business.General.Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.ActualizarQLineasPromociones)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.ActualizacionAutomaticaStock)
        updateProcess.AddTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.DetalleActualizacionStocks)
    End Sub

#End Region
#Region " Business Rules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New SynonymousBusinessRules
        oBRL.AddSynonymous("FechaAlbaran", "Fecha")

        Dim services As New ServiceProvider
        oBRL = ProcessServer.ExecuteTask(Of BusinessRules, BusinessRules)(AddressOf ProcesoComercial.DetailBusinessRulesCab, oBRL, services)

        oBRL("IDCliente") = AddressOf CambioCliente
        oBRL.Add("Fecha", AddressOf ProcesoComunes.CambioFechaAlbaran)
        oBRL.Add("IDAlmacen", AddressOf ProcesoComunes.CambioAlmacen)
        oBRL.Add("IDCentroGestion", AddressOf ProcesoComunes.CambioCentroGestion)
        oBRL.Add("Ticket", AddressOf CambioTicket)

        oBRL.Add("IDFormaEnvio", AddressOf CambioIDFormaEnvio)
        oBRL.Add("IDProveedor", AddressOf CambioIDProveedor)
        oBRL.Add("IDDireccion", AddressOf ProcesoComercial.CambioDireccion)
        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioCliente(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComercial.CambioCliente, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComunes.CambioMoneda, data, services)
        Dim dir As New DataDireccionClte(enumcdTipoDireccion.cdDireccionEnvio, "IDDireccion", data.Current)
        ProcessServer.ExecuteTask(Of DataDireccionClte)(AddressOf ProcesoComercial.AsignarDireccionCliente, dir, services)
        Dim obs As New DataObservaciones(GetType(AlbaranVentaCabecera).Name, "Texto", data.Current)
        ProcessServer.ExecuteTask(Of DataObservaciones)(AddressOf ProcesoComercial.AsignarObservacionesCliente, obs, services)

        If Length(data.Current("IDCliente")) > 0 Then
            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data.Current("IDCliente"))

            If ClteInfo.Bloqueado Then ApplicationService.GenerateError("El Cliente está bloqueado.")
            data.Current("IDPais") = ClteInfo.Pais
            data.Current("Telefono") = ClteInfo.Telefono
            data.Current("Fax") = ClteInfo.Fax
            data.Current("IDModoTransporte") = ClteInfo.IDModoTransporte
            data.Current("IdBancoPropio") = ClteInfo.IDBancoPropio
            data.Current("DtoAlbaran") = ClteInfo.DtoComercial
            data.Current = New AlbaranVentaCabecera().ApplyBusinessRule("IDFormaEnvio", data.Current("IDFormaEnvio"), data.Current, data.Context)
            data.Current("TieneRE") = ClteInfo.TieneRE
        Else
            data.Current("IDPais") = System.DBNull.Value
            data.Current("Telefono") = System.DBNull.Value
            data.Current("Fax") = System.DBNull.Value
            data.Current("IDModoTransporte") = System.DBNull.Value
            data.Current("IdBancoPropio") = System.DBNull.Value
            data.Current("DtoAlbaran") = 0
            data.Current("TieneRE") = System.DBNull.Value
        End If

        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf GetTarifaCliente, data, services)
    End Sub

    <Task()> Public Shared Sub GetTarifaCliente(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Current("IDCliente")) > 0 Then
            Dim f As New Filter
            Dim dt As DataTable = New ClienteTarifa().Filter(New StringFilterItem("IDCliente", data.Current("IDCliente")), "Orden")
            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf GetTarifaCentroGestion, data, services)
            Else
                data.Current("IDTarifa") = dt.Rows(0)("IDTarifa")
            End If
        Else
            data.Current("IDTarifa") = System.DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub GetTarifaCentroGestion(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim IDCentroGestion As String = data.Current("IDCentroGestion") & String.Empty
        If Length(IDCentroGestion) = 0 Then
            Dim cgu As New UsuarioCentroGestion.UsuarioCentroGestionInfo
            cgu = ProcessServer.ExecuteTask(Of UsuarioCentroGestion.UsuarioCentroGestionInfo, UsuarioCentroGestion.UsuarioCentroGestionInfo)(AddressOf UsuarioCentroGestion.ObtenerUsuarioCentroGestion, cgu, services)
            IDCentroGestion = cgu.IDCentroGestion
        End If
        Dim dt As DataTable = New CentroGestion().SelOnPrimaryKey(IDCentroGestion)
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            data.Current("IDTarifa") = dt.Rows(0)("IDTarifa") & String.Empty
        End If
    End Sub

    <Task()> Public Shared Sub CambioTicket(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value

        Dim AppParams As ParametroVenta = services.GetService(Of ParametroVenta)()
        If Length(AppParams.ContadorProvisionalTPV) > 0 AndAlso AppParams.ContadorProvisionalTPV = data.Current("IDContador") Then
            Exit Sub
        End If
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf AsignarContadorTPV, data, services)
    End Sub

    <Task()> Public Shared Sub AsignarContadorTPV(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Current("IDCentroGestion")) = 0 Then Exit Sub
        Dim strContador As String
        Dim BlnApplyChange As Boolean = True
        If data.Current.ContainsKey("Ticket") AndAlso Nz(data.Current("Ticket"), False) Then
            Dim ClsVendTPV As New Operario
            Dim DtVend As DataTable = ClsVendTPV.SelOnPrimaryKey(data.Current("IDVendedor"))

            If Length(DtVend.Rows(0)("IDContadorTicket")) > 0 Then strContador = DtVend.Rows(0)("IDContadorTicket")

            If Length(strContador) = 0 Then
                Dim ClsPCCentro As BusinessHelper = CreateBusinessObject("PCCentroGestion")
                Dim FilPCCentro As New Filter
                FilPCCentro.Add("IDCentrogestion", FilterOperator.Equal, data.Current("IDCentroGestion"))
                FilPCCentro.Add("IDTPV", FilterOperator.Equal, data.Current("IDTPV"))
                Dim DtPCCentro As DataTable = ClsPCCentro.Filter(FilPCCentro)
                If Not DtPCCentro Is Nothing AndAlso DtPCCentro.Rows.Count > 0 Then
                    If Length(DtPCCentro.Rows(0)("IDContadorTicket")) > 0 Then
                        strContador = DtPCCentro.Rows(0)("IDContadorTicket")
                    End If
                End If
            End If
            If Length(strContador) = 0 Then
                Dim dt As DataTable = New CentroGestion().SelOnPrimaryKey(data.Current("IDCentroGestion"))
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    If data.Current.ContainsKey("Ticket") AndAlso Nz(data.Current("Ticket"), False) Then
                        If Length(dt.Rows(0)("IDContadorAlbaranVentaTPV")) > 0 Then
                            strContador = dt.Rows(0)("IDContadorAlbaranVentaTPV")
                        Else
                            If Length(dt.Rows(0)("IDContadorAlbaranVenta")) > 0 Then
                                strContador = dt.Rows(0)("IDContadorAlbaranVenta")
                            Else : ApplicationService.GenerateError("No se ha configurado un contador predeterminado de Albaranes de Venta de TPV, ni de Albaranes de Venta.|Revise la configuración del centro gestión.", vbNewLine)
                            End If
                        End If
                    Else
                        If Length(dt.Rows(0)("IDContadorAlbaranVenta")) > 0 Then
                            strContador = dt.Rows(0)("IDContadorAlbaranVenta")
                        Else : ApplicationService.GenerateError("No se ha configurado un contador predeterminado de Albaranes de Venta de Venta.|Revise la configuración del centro gestión.", vbNewLine)
                        End If
                    End If
                End If
            End If
        Else
            Dim dt As DataTable = New CentroGestion().SelOnPrimaryKey(data.Current("IDCentroGestion"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                If Length(dt.Rows(0)("IDContadorAlbaranVenta")) > 0 Then
                    If data.Current("IDContador") <> dt.Rows(0)("IDContadorAlbaranVenta") Then
                        strContador = dt.Rows(0)("IDContadorAlbaranVenta")
                    Else : BlnApplyChange = False
                    End If
                Else : ApplicationService.GenerateError("No se ha configurado un contador predeterminado de Albaranes de Venta de Venta.|Revise la configuración del centro gestión.", vbNewLine)
                End If
            End If
        End If
        If Length(strContador) > 0 AndAlso BlnApplyChange Then
            'Dim DtCont As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDt, GetType(AlbaranVentaCabecera).Name, services)
            'Dim f As New Filter
            'f.Add(New StringFilterItem("IDContador", strContador))
            'Dim WhereContador As String = f.Compose(New AdoFilterComposer)
            'Dim adr() As DataRow = DtCont.Select(WhereContador)
            'If Not IsNothing(adr) AndAlso adr.Length > 0 Then
            '    data.Current("NAlbaran") = adr(0)("ValorProvisional")
            '    data.Current("IDContador") = strContador
            'End If

            Dim StCont As Contador.CounterTx = ProcessServer.ExecuteTask(Of String, Contador.CounterTx)(AddressOf Contador.CounterValueTx, strContador, services)
            data.Current("NAlbaran") = StCont.strCounterValue
            data.Current("IDContador") = strContador
        End If
    End Sub


    <Task()> Public Shared Sub CambioIDFormaEnvio(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim AVC As New AlbaranVentaCabecera
            Dim filtro As New StringFilterItem("IDFormaEnvio", FilterOperator.Equal, data.Value)
            Dim dtEnvio As DataTable = New FormaEnvio().Filter(filtro)
            If (Not dtEnvio Is Nothing AndAlso dtEnvio.Rows.Count > 0) Then
                data.Current("IdFormaEnvio") = dtEnvio.Rows(0)("IDFormaEnvio")
                data.Current("EmpresaTransp") = dtEnvio.Rows(0)("DescFormaEnvio")
                data.Current("IDProveedor") = dtEnvio.Rows(0)("IDProveedor")
                data.Current = AVC.ApplyBusinessRule("IDProveedor", data.Current("IDProveedor"), data.Current)
                Dim filtroDetalle As New Filter
                filtroDetalle.Add("IDFormaEnvio", FilterOperator.Equal, data.Current("IdFormaEnvio"))
                filtroDetalle.Add("Predeterminado", FilterOperator.Equal, True)
                Dim dtEnvioD As DataTable = New FormaEnvioDetalle().Filter(filtroDetalle)
                If (Not dtEnvioD Is Nothing AndAlso dtEnvioD.Rows.Count > 0) Then
                    data.Current("CONDUCTOR") = dtEnvioD.Rows(0)("Conductor")
                    data.Current("DNICONDUCTOR") = dtEnvioD.Rows(0)("DNIConductor")
                    data.Current("MATRICULA") = dtEnvioD.Rows(0)("Matricula")
                    data.Current("Remolque") = dtEnvioD.Rows(0)("Remolque")
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioIDProveedor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Current("IDProveedor")) > 0 Then
            If Length(data.Value) > 0 Then
                Dim dr As DataRow = New Proveedor().GetItemRow(data.Value)
                data.Current("CifTransportista") = dr("CifProveedor")
            End If
        Else
            data.Current("CifTransportista") = DBNull.Value
        End If
    End Sub
#End Region
#Region " Actualizacion de stocks "

    <Task()> Public Shared Function ActualizarStockAlbaranes(ByVal IDAlbaran() As Integer, ByVal services As ServiceProvider) As StockUpdateData()
        Dim updateDataArray(-1) As StockUpdateData
        For Each id As Integer In IDAlbaran
            Dim updateData() As StockUpdateData = ProcessServer.ExecuteTask(Of Integer, StockUpdateData())(AddressOf ActualizarStockAlbaran, id, services)
            ArrayManager.Copy(updateData, updateDataArray)
        Next
        Return updateDataArray
    End Function

    <Task()> Public Shared Function ActualizarStockAlbaran(ByVal IDAlbaran As Integer, ByVal services As ServiceProvider) As StockUpdateData()
        Dim updateDataArray(-1) As StockUpdateData
        If IDAlbaran <> 0 Then
            Dim Doc As New DocumentoAlbaranVenta(IDAlbaran)
            Dim actLin As New ProcesoStocks.DataActualizarStockLineas(Doc)
            Dim updateData() As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarStockLineas, StockUpdateData())(AddressOf ProcesoAlbaranVenta.ActualizarStockLineas, actLin, services)
            If Not IsNothing(updateData) AndAlso Length(updateData) > 0 Then
                ArrayManager.Copy(updateData, updateDataArray)
            End If
        End If
        Return updateDataArray
    End Function

    <Task()> Public Shared Function ActualizarStockLineasAlbaran(ByVal IDProcess As Guid, ByVal services As ServiceProvider) As StockUpdateData()
        Dim updateDataArray(-1) As StockUpdateData
        Dim LineasAlbaran As DataTable = New BE.DataEngine().Filter("vFrmActualizacionAlbaranVentaLinea", New GuidFilterItem("IDProcess", IDProcess))
        If LineasAlbaran.Rows.Count > 0 Then
            Dim IDAlbaranes(-1) As Integer
            Dim IDLineasAlbaran(-1) As Integer
            For Each linea As DataRow In LineasAlbaran.Select(Nothing, "IDAlbaran")
                If Array.IndexOf(IDAlbaranes, linea("IDAlbaran")) < 0 Then
                    ReDim Preserve IDAlbaranes(IDAlbaranes.Length)
                    IDAlbaranes(IDAlbaranes.Length - 1) = linea("IDAlbaran")
                End If
            Next

            For Each IDAlbaran As Integer In IDAlbaranes
                Dim Doc As New DocumentoAlbaranVenta(IDAlbaran)
                ReDim IDLineasAlbaran(-1)
                For Each LineaAlbaran As DataRow In LineasAlbaran.Select("IDAlbaran=" & IDAlbaran)
                    ReDim Preserve IDLineasAlbaran(IDLineasAlbaran.Length)
                    IDLineasAlbaran(IDLineasAlbaran.Length - 1) = LineaAlbaran("IDLineaAlbaran")
                Next
                Dim actLin As New ProcesoStocks.DataActualizarStockLineas(Doc, IDLineasAlbaran)
                Dim updateData() As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarStockLineas, StockUpdateData())(AddressOf ProcesoAlbaranVenta.ActualizarStockLineas, actLin, services)
                If updateData.Length > 0 Then
                    ArrayManager.Copy(updateData, updateDataArray)
                End If
            Next
        End If
        Return updateDataArray
    End Function

#End Region
#Region " Precio Optimo "

    <Serializable()> _
    Public Class DataPrecioOptimo
        Public IDAlbaran As Integer
        Public FechaCalculo As Date

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDAlbaran As Integer, ByVal FechaCalculo As Date)
            Me.IDAlbaran = IDAlbaran
            Me.FechaCalculo = FechaCalculo
        End Sub
    End Class

    <Serializable()> _
    Public Class DataCalcPrecioOpt
        Public FechaCalculo As Date
        Public DocAlb As DocumentoAlbaranVenta

        Public Sub New()
        End Sub
        Public Sub New(ByVal FechaCalculo As Date, ByVal DocAlb As DocumentoAlbaranVenta)
            Me.FechaCalculo = FechaCalculo
            Me.DocAlb = DocAlb
        End Sub
    End Class

    <Task()> Public Shared Sub PrecioOptimo(ByVal data As DataPrecioOptimo, ByVal services As ServiceProvider)
        Dim DocAlb As DocumentoAlbaranVenta = ProcessServer.ExecuteTask(Of Integer, DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.GetDocumento, data.IDAlbaran, services)
        Dim StData As New DataCalcPrecioOpt(data.FechaCalculo, DocAlb)
        ProcessServer.ExecuteTask(Of DataCalcPrecioOpt)(AddressOf CalculoPrecioOptimo, StData, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaPedidos.CalcularRepresentantes, DocAlb, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaPedidos.CalcularAnalitica, DocAlb, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.CalcularBasesImponibles, DocAlb, services)
        'ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.CalcularImportesAlbaran, DocAlb, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.TotalDocumento, DocAlb, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.GrabarDocumento, DocAlb, services)
    End Sub

    <Task()> Public Shared Sub CalculoPrecioOptimo(ByVal data As DataCalcPrecioOpt, ByVal services As ServiceProvider)
        If data.DocAlb Is Nothing OrElse data.DocAlb.dtLineas Is Nothing OrElse data.DocAlb.dtLineas.Rows.Count = 0 Then Exit Sub

        '//Recogemos los articulos que esten relacionados con esa Albaran.
        Dim dtArticulosAlbaran As DataTable = AdminData.GetData("vNegAlbaranVentaLineaArticulos", New StringFilterItem("IDAlbaran", data.DocAlb.HeaderRow("IDAlbaran")))
        Dim f As New Filter
        For Each drArticuloAlbaran As DataRow In dtArticulosAlbaran.Select
            f.Clear()
            f.Add("IDArticulo", drArticuloAlbaran("IDArticulo"))

            '//Recogemos las lineas del albarán que tengan el articulo de este momento
            Dim QServida As Double = Nz(data.DocAlb.dtLineas.Compute("SUM(QServida)", f.Compose(New AdoFilterComposer)), 0)

            Dim dataTarifa As New DataCalculoTarifaComercial
            dataTarifa.IDArticulo = drArticuloAlbaran("IDArticulo")
            dataTarifa.IDCliente = data.DocAlb.IDCliente
            dataTarifa.Cantidad = QServida
            dataTarifa.Fecha = data.FechaCalculo

            ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial, DataTarifaComercial)(AddressOf ProcesoComercial.TarifaComercial, dataTarifa, services)
            If Not dataTarifa.DatosTarifa Is Nothing AndAlso dataTarifa.DatosTarifa.Precio <> 0 Then
                Dim AVL As New AlbaranVentaLinea
                Dim context As New BusinessData(data.DocAlb.HeaderRow)
                Dim WhereArticulo As String = f.Compose(New AdoFilterComposer)
                For Each drAlbaranLineaArticulo As DataRow In data.DocAlb.dtLineas.Select(WhereArticulo)
                    AVL.ApplyBusinessRule("Precio", dataTarifa.DatosTarifa.Precio, drAlbaranLineaArticulo, context)
                    AVL.ApplyBusinessRule("Dto1", dataTarifa.DatosTarifa.Dto1, drAlbaranLineaArticulo, context)
                    AVL.ApplyBusinessRule("Dto2", dataTarifa.DatosTarifa.Dto2, drAlbaranLineaArticulo, context)
                    AVL.ApplyBusinessRule("Dto3", dataTarifa.DatosTarifa.Dto3, drAlbaranLineaArticulo, context)
                    AVL.ApplyBusinessRule("UDValoracion", dataTarifa.DatosTarifa.UDValoracion, drAlbaranLineaArticulo, context)

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
            QServida = 0
        Next
    End Sub

#End Region
#Region " Transporte Propio "

    <Task()> Public Shared Sub AñadirTransportePropio(ByVal IDAlbaran As Integer, ByVal services As ServiceProvider)
        Dim Doc As DocumentoAlbaranVenta = ProcessServer.ExecuteTask(Of Integer, DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.GetDocumento, IDAlbaran, services)

        Dim AppParams As ParametroAlbaranVenta = services.GetService(Of ParametroAlbaranVenta)()
        Dim IDArtTrans As String = AppParams.ArticuloTransportePropio
        If Len(IDArtTrans) = 0 Then ApplicationService.GenerateError("No se ha definido en los parámetros de la aplicación un artículo para identificar los transportes propios.")

        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(IDArtTrans)
        Dim IDUdMedidaT As String = ArtInfo.IDUDVenta
        If Len(IDUdMedidaT) = 0 Then IDUdMedidaT = ArtInfo.IDUDInterna
        If Len(IDUdMedidaT) = 0 Then ApplicationService.GenerateError("El artículo que define los transportes propios necesita una unidad de medida.")

        Dim Q As Double
        Dim rwAvl As DataRow
        For Each oRw As DataRow In Doc.dtLineas.Rows
            If oRw("IDArticulo") = IDArtTrans Then
                rwAvl = oRw
            Else
                Dim UDMedida As String = oRw("IDUdMedida") & String.Empty
                If Len(UDMedida) Then
                    Dim StDatos As New ArticuloUnidadAB.DatosFactorConversion
                    StDatos.IDArticulo = oRw("IDArticulo")
                    StDatos.IDUdMedidaA = UDMedida
                    StDatos.IDUdMedidaB = IDUdMedidaT
                    StDatos.UnoSiNoExiste = True
                    Dim Factor As Double = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, StDatos, services)
                    If Factor = 0 Then Factor = 1
                    Q += oRw("QServida") * Factor
                End If
            End If
        Next

        If Q > 0 Then
            Dim oAvl As New AlbaranVentaLinea
            Dim NuevaLinea As Boolean = False
            If rwAvl Is Nothing Then
                rwAvl = oAvl.AddNewForm.Rows(0)
                rwAvl("IDLineaAlbaran") = AdminData.GetAutoNumeric
                rwAvl("IDAlbaran") = IDAlbaran
                rwAvl("QServida") = 0
                rwAvl("Precio") = 0
                rwAvl("Dto1") = 0
                rwAvl("Dto2") = 0
                rwAvl("Dto3") = 0
                rwAvl("IDTipoLinea") = ProcessServer.ExecuteTask(Of Object, String)(AddressOf TipoLinea.TipoLineaPorDefecto, Nothing, services)
                rwAvl("TipoLineaAlbaran") = enumavlTipoLineaAlbaran.avlNormal
                rwAvl("IDAlmacen") = Doc.HeaderRow("IDAlmacen")
                rwAvl("IDFormaPago") = Doc.HeaderRow("IDFormaPago")
                rwAvl("IDCondicionPago") = Doc.HeaderRow("IDCondicionPago")
                rwAvl("IDDireccionFra") = Doc.HeaderRow("IDDireccionFra")
                rwAvl("EstadoFactura") = BusinessEnum.enumavlEstadoFactura.avlNoFacturado

                NuevaLinea = True
            End If

            Dim context As New BusinessData(Doc.HeaderRow)
            oAvl.ApplyBusinessRule("IDArticulo", IDArtTrans, rwAvl, context)
            oAvl.ApplyBusinessRule("QServida", Q, rwAvl, context)

            If NuevaLinea Then Doc.dtLineas.Rows.Add(rwAvl.ItemArray)

            ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.CalcularImporteLineasAlbaran, Doc, services)
            ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaPedidos.CalcularRepresentantes, Doc, services)
            ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVentaPedidos.CalcularAnalitica, Doc, services)
            ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoAlbaranVenta.CalcularBasesImponibles, Doc, services)
            ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf ProcesoComunes.TotalDocumento, Doc, services)
            ProcessServer.ExecuteTask(Of DocumentoAlbaranVenta)(AddressOf Business.General.Comunes.UpdateDocument, Doc, services)
        End If
    End Sub

#End Region
#Region " Estadisticas "

    <Task()> Public Shared Function ObtenerEstadisticaAVTipos(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Return New DataEngine().Filter("tbEstadisticaVentaAnual", String.Empty, String.Empty)
    End Function

    <Serializable()> _
    Public Class DataEstadisticaCantidadesMeses
        Public CamposSelect As String
        Public CampoATotalizar As String
        Public CamposOrden As String
        Public GroupBy As String

        Public IDTipo As String
        Public IDFamilia As String
        Public IDArticulo As String
        Public IDCliente As String
        Public IDGrupoCliente As String
        Public IDTipoAlbaran As String
        Public IDMercado As String
        Public IDProvincia As String
        Public IDZona As String
        Public IDPais As String
        Public Facturable As enumBoolean
        Public CEE As enumBoolean
        Public Extranjero As enumBoolean
        Public Año As Integer
        Public EmpresaGrupo As enumBoolean

        Public Sub New(ByVal CamposSelect As String, ByVal CampoATotalizar As String, _
                       ByVal IDTipo As String, ByVal IDFamilia As String, _
                       ByVal IDArticulo As String, ByVal IDCliente As String, _
                       ByVal IDGrupoCliente As String, _
                       ByVal IDTipoAlbaran As String, ByVal IDMercado As String, _
                       ByVal IDProvincia As String, ByVal IDZona As String, ByVal IDPais As String, _
                       ByVal Facturable As enumBoolean, ByVal CEE As enumBoolean, _
                       ByVal Extranjero As enumBoolean, ByVal Año As Integer, _
                       ByVal EmpresaGrupo As enumBoolean, ByVal GroupBy As String, _
                       ByVal CamposOrden As String)

            Me.CamposSelect = CamposSelect
            Me.CampoATotalizar = CampoATotalizar
            Me.CamposOrden = CamposOrden
            Me.GroupBy = GroupBy
            Me.IDTipo = IDTipo
            Me.IDFamilia = IDFamilia
            Me.IDArticulo = IDArticulo
            Me.IDCliente = IDCliente
            Me.IDGrupoCliente = IDGrupoCliente
            Me.IDTipoAlbaran = IDTipoAlbaran
            Me.IDMercado = IDMercado
            Me.IDProvincia = IDProvincia
            Me.IDZona = IDZona
            Me.IDPais = IDPais
            Me.Facturable = Facturable
            Me.CEE = CEE
            Me.Extranjero = Extranjero
            Me.Año = Año
            Me.EmpresaGrupo = EmpresaGrupo
        End Sub
    End Class
    <Task()> Public Shared Function ObtenerEstadisticaCantidadesMeses(ByVal data As DataEstadisticaCantidadesMeses, ByVal services As ServiceProvider) As DataTable
        Dim selectSQL As New System.Text.StringBuilder
        selectSQL.Append(String.Format( _
            "SELECT {0}, " & _
            "SUM(CASE MONTH([FechaAlbaran]) WHEN 1 THEN {1} ELSE 0 END) AS SEnero," & _
            "SUM(CASE MONTH([FechaAlbaran]) WHEN 2 THEN {1} ELSE 0 END) AS SFebrero," & _
            "SUM(CASE MONTH([FechaAlbaran]) WHEN 3 THEN {1} ELSE 0 END) AS SMarzo," & _
            "SUM(CASE MONTH([FechaAlbaran]) WHEN 4 THEN {1} ELSE 0 END) AS SAbril," & _
            "SUM(CASE MONTH([FechaAlbaran]) WHEN 5 THEN {1} ELSE 0 END) AS SMayo," & _
            "SUM(CASE MONTH([FechaAlbaran]) WHEN 6 THEN {1} ELSE 0 END) AS SJunio," & _
            "SUM(CASE MONTH([FechaAlbaran]) WHEN 7 THEN {1} ELSE 0 END) AS SJulio," & _
            "SUM(CASE MONTH([FechaAlbaran]) WHEN 8 THEN {1} ELSE 0 END) AS SAgosto," & _
            "SUM(CASE MONTH([FechaAlbaran]) WHEN 9 THEN {1} ELSE 0 END) AS SSeptiembre," & _
            "SUM(CASE MONTH([FechaAlbaran]) WHEN 10 THEN {1} ELSE 0 END) AS SOctubre," & _
            "SUM(CASE MONTH([FechaAlbaran]) WHEN 11 THEN {1} ELSE 0 END) AS SNoviembre," & _
            "SUM(CASE MONTH([FechaAlbaran]) WHEN 12 THEN {1} ELSE 0 END)  AS SDiciembre," & _
            "SUM({1}) As STotalLinea", data.CamposSelect, data.CampoATotalizar))

        selectSQL.Append(" FROM tbAlbaranVentaLinea INNER JOIN " & _
        "vAlbaranVentaCabecera ON tbAlbaranVentaLinea.IDAlbaran = vAlbaranVentaCabecera.IDAlbaran INNER JOIN " & _
        "tbMaestroCliente ON vAlbaranVentaCabecera.IDCliente = tbMaestroCliente.IDCliente INNER JOIN " & _
        "tbMaestroArticulo ON tbAlbaranVentaLinea.IDArticulo = tbMaestroArticulo.IDArticulo  LEFT OUTER JOIN " & _
        "tbMaestroPais ON tbMaestroCliente.IDPais = tbMaestroPais.IDPais INNER JOIN " & _
        "tbMaestroFamilia ON tbMaestroArticulo.IDFamilia = tbMaestroFamilia.IDFamilia AND " & _
        "tbMaestroArticulo.IDTipo = tbMaestroFamilia.IDTipo INNER JOIN " & _
        "tbMaestroTipoAlbaran ON vAlbaranVentaCabecera.IDTipoAlbaran = tbMaestroTipoAlbaran.IDTipoAlbaran  LEFT OUTER JOIN " & _
        "tbMaestroZona ON tbMaestroCliente.IDZona = tbMaestroZona.IDZona  LEFT OUTER JOIN " & _
        "tbMaestroMercado ON tbMaestroCliente.IDMercado = tbMaestroMercado.IDMercado INNER JOIN " & _
        "tbMaestroTipoIva ON tbAlbaranVentaLinea.IDTipoIva = tbMaestroTipoIva.IDTipoIva LEFT OUTER JOIN " & _
        "tbMaestroCliente AS tbMaestroCliente_1 ON vAlbaranVentaCabecera.IDCliente = tbMaestroCliente_1.IDCliente")

        Dim whereSQL As New Text.StringBuilder
        If data.Año.ToString.Length > 0 Then
            whereSQL.Append("YEAR(vAlbaranVentaCabecera.FechaAlbaran) = " & data.Año & " AND ")
        End If
        If data.IDTipo.Length > 0 Then
            whereSQL.Append("tbMaestroArticulo.IDTipo = '" & data.IDTipo & "' AND ")
        End If
        If data.IDFamilia.Length > 0 Then
            whereSQL.Append("tbMaestroArticulo.IDFamilia = '" & data.IDFamilia & "' AND ")
        End If
        If data.IDArticulo.Length > 0 Then
            whereSQL.Append("tbMaestroArticulo.IDArticulo = '" & data.IDArticulo & "' AND ")
        End If

        If data.IDCliente.Length > 0 Then
            whereSQL.Append("tbMaestroCliente.IDCliente = '" & data.IDCliente & "' AND ")
        End If
        If data.IDGrupoCliente.Length > 0 Then
            whereSQL.Append("dbo.fGrupoCliente(vAlbaranVentaCabecera.IDCliente, tbMaestroCliente.IDGrupoCliente, tbMaestroCliente_1.IDCliente, tbMaestroCliente_1.IDGrupoCliente) = '" & data.IDGrupoCliente & "' AND ")
        End If
        If data.IDTipoAlbaran.Length > 0 Then
            whereSQL.Append("tbMaestroTipoAlbaran.IDTipoAlbaran = '" & data.IDTipoAlbaran & "' AND ")
        End If
        If data.IDProvincia.Length > 0 Then
            whereSQL.Append("tbMaestroCliente.Provincia = '" & data.IDProvincia & "' AND ")
        End If
        If data.IDZona.Length > 0 Then
            whereSQL.Append("tbMaestroCliente.IDZona = '" & data.IDZona & "' AND ")
        End If
        If data.IDMercado.Length > 0 Then
            whereSQL.Append("tbMaestroCliente.IDMercado = '" & data.IDMercado & "' AND ")
        End If
        If data.IDPais.Length > 0 Then
            whereSQL.Append("tbMaestroPais.IDPais = '" & data.IDPais & "' AND ")
        End If

        Select Case data.Facturable
            Case enumBoolean.Si
                whereSQL.Append("tbMaestroTipoAlbaran.Facturable = 1 AND ")
            Case enumBoolean.No
                whereSQL.Append("tbMaestroTipoAlbaran.Facturable = 0 AND ")
        End Select

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

        selectSQL.Append(" GROUP BY ")
        selectSQL.Append(data.GroupBy)
        selectSQL.Append(" ORDER BY ")
        selectSQL.Append(data.CamposOrden)

        Dim cmdEstadisticas As Common.DbCommand = AdminData.GetCommand
        cmdEstadisticas.CommandType = CommandType.Text
        cmdEstadisticas.CommandText = selectSQL.ToString()
        Return AdminData.Execute(cmdEstadisticas, ExecuteCommand.ExecuteReader)

    End Function

#End Region
#Region " Numeros Serie "

    <Task()> Public Shared Function ComprobarNumerosSerieAlbaran(ByVal IDAlbaran As Integer, ByVal services As ServiceProvider) As Boolean
        Dim dtLineas As DataTable = New AlbaranVentaLinea().Filter(New FilterItem("IDAlbaran", IDAlbaran))
        If Not dtLineas Is Nothing AndAlso dtLineas.Rows.Count > 0 Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            For Each Dr As DataRow In dtLineas.Select
                Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(Dr("IDArticulo"))
                If ArtInfo.NSerieObligatorio AndAlso Length(Dr("Lote")) = 0 Then
                    Return False
                End If
            Next
        End If
        Return True
    End Function

#End Region

#Region " Arqueo TPV "
    <Serializable()> _
    Public Class DataArqueoTPV
        Public IDCentroGestion As String
        Public IDTPV As String
        Public FechaDesde As Date
        Public FechaHasta As Date
    End Class
    '<Task()> Public Shared Function Arqueo(ByVal ArqueoTPV As DataArqueoTPV, ByVal services As ServiceProvider)

    '    Dim f As New Filter

    '    If Length(ArqueoTPV.IDCentroGestion) > 0 Then f.Add("IDCentroGestion", FilterOperator.Equal, ArqueoTPV.IDCentroGestion)
    '    If Length(ArqueoTPV.IDTPV) > 0 Then f.Add("IDTPV", FilterOperator.Equal, ArqueoTPV.IDTPV)
    '    If Length(ArqueoTPV.FechaDesde) > 0 Then f.Add(New DateFilterItem("FechaAlbaran", FilterOperator.GreaterThanOrEqual, ArqueoTPV.FechaDesde))
    '    If Length(ArqueoTPV.FechaHasta) > 0 Then f.Add(New DateFilterItem("FechaAlbaran", FilterOperator.LessThanOrEqual, ArqueoTPV.FechaHasta))
    '    f.Add(New NumberFilterItem("Estado", FilterOperator.Equal, enumaccEstado.accNoFacturado))

    '    Dim dtAlbaranes As DataTable = AdminData.GetData("vFrmTPVVentasAlbArqueo", f)
    '    For Each drAlb As DataRow In dtAlbaranes.Rows
    '        Dim dtAlbModif As DataTable = New AlbaranVentaCabecera().SelOnPrimaryKey(drAlb("IDAlbaran"))
    '        If Not dtAlbModif Is Nothing AndAlso dtAlbModif.Rows.Count > 0 Then
    '            dtAlbModif.Rows(0)("Arqueo") = True
    '            BusinessHelper.UpdateTable(dtAlbModif)
    '        End If
    '    Next


    '    f.Clear()
    '    If Length(ArqueoTPV.IDCentroGestion) > 0 Then f.Add("IDCentroGestion", FilterOperator.Equal, ArqueoTPV.IDCentroGestion)
    '    If Length(ArqueoTPV.IDTPV) > 0 Then f.Add("IDTPV", FilterOperator.Equal, ArqueoTPV.IDTPV)
    '    If Length(ArqueoTPV.FechaDesde) > 0 Then f.Add(New DateFilterItem("FechaFactura", FilterOperator.GreaterThanOrEqual, ArqueoTPV.FechaDesde))
    '    If Length(ArqueoTPV.FechaHasta) > 0 Then f.Add(New DateFilterItem("FechaFactura", FilterOperator.LessThanOrEqual, ArqueoTPV.FechaHasta))
    '    Dim dtFacturas As DataTable = New BE.DataEngine().Filter("vFrmTPVVentasFactArqueo", f)
    '    Dim objFactura As New FacturaVentaCabecera
    '    For Each drFact As DataRow In dtFacturas.Rows
    '        Dim dtFactModif As DataTable = objFactura.SelOnPrimaryKey(drFact("IDFactura"))
    '        If Not dtFactModif Is Nothing AndAlso dtFactModif.Rows.Count > 0 Then
    '            dtFactModif.Rows(0)("Arqueo") = True
    '            BusinessHelper.UpdateTable(dtFactModif)
    '        End If
    '    Next


    'End Function

    'Public Function DesArqueo(ByVal ArqueoTPV As DataArqueoTPV, ByVal services As ServiceProvider)

    '    Dim f As New Filter

    '    If Length(ArqueoTPV.IDCentroGestion) > 0 Then f.Add("IDCentroGestion", FilterOperator.Equal, ArqueoTPV.IDCentroGestion)
    '    If Length(ArqueoTPV.IDTPV) > 0 Then f.Add("IDTPV", FilterOperator.Equal, ArqueoTPV.IDTPV)
    '    If Length(ArqueoTPV.FechaDesde) > 0 Then f.Add(New DateFilterItem("FechaAlbaran", FilterOperator.GreaterThanOrEqual, ArqueoTPV.FechaDesde))
    '    If Length(ArqueoTPV.FechaHasta) > 0 Then f.Add(New DateFilterItem("FechaAlbaran", FilterOperator.LessThanOrEqual, ArqueoTPV.FechaHasta))
    '    f.Add(New BooleanFilterItem("Arqueo", FilterOperator.Equal, True))

    '    Dim dtAlbaranes As DataTable = AdminData.GetData("vFrmTPVVentasAlbArqueo", f)
    '    For Each drAlb As DataRow In dtAlbaranes.Rows
    '        Dim dtAlbModif As DataTable = Me.SelOnPrimaryKey(drAlb("IDAlbaran"))
    '        If Not dtAlbModif Is Nothing AndAlso dtAlbModif.Rows.Count > 0 Then
    '            dtAlbModif.Rows(0)("Arqueo") = False
    '            BusinessHelper.UpdateTable(dtAlbModif)
    '        End If
    '    Next

    '    f.Clear()
    '    If Length(ArqueoTPV.IDCentroGestion) > 0 Then f.Add("IDCentroGestion", FilterOperator.Equal, ArqueoTPV.IDCentroGestion)
    '    If Length(ArqueoTPV.IDTPV) > 0 Then f.Add("IDTPV", FilterOperator.Equal, ArqueoTPV.IDTPV)
    '    If Length(ArqueoTPV.FechaDesde) > 0 Then f.Add(New DateFilterItem("FechaFactura", FilterOperator.GreaterThanOrEqual, ArqueoTPV.FechaDesde))
    '    If Length(ArqueoTPV.FechaHasta) > 0 Then f.Add(New DateFilterItem("FechaFactura", FilterOperator.LessThanOrEqual, ArqueoTPV.FechaHasta))
    '    Dim dtFacturas As DataTable = AdminData.GetData("vFrmTPVVentasFactArqueo", f)
    '    Dim objFactura As New FacturaVentaCabecera
    '    For Each drFact As DataRow In dtFacturas.Rows
    '        Dim dtFactModif As DataTable = objFactura.SelOnPrimaryKey(drFact("IDFactura"))
    '        If Not dtFactModif Is Nothing AndAlso dtFactModif.Rows.Count > 0 Then
    '            dtFactModif.Rows(0)("Arqueo") = False
    '            BusinessHelper.UpdateTable(dtFactModif)
    '        End If
    '    Next

    'End Function
#End Region

#Region "Nota Transportista"

    <Task()> Public Shared Function AsignarNotasTransportistaMultiEmpresa(ByVal data As DataTable, ByVal services As ServiceProvider) As String
        AdminData.BeginTx()
        Dim DescBBDD As String = AdminData.GetSessionInfo.DataBase.DataBaseDescription
        Dim DescBBDDOriginal As String = DescBBDD
        Dim IDBaseDatosOriginal As Guid = AdminData.GetConnectionInfo.IDDataBase

        Dim StrContador As String = String.Empty
        Dim StrIDContador As String = String.Empty
        Dim DtParam As DataTable = New Parametro().SelOnPrimaryKey("CNTRANS")
        If Not DtParam Is Nothing AndAlso DtParam.Rows.Count > 0 Then
            If Length(DtParam.Rows(0)("Valor")) > 0 Then
                Dim DtCont As DataTable = New Contador().SelOnPrimaryKey(DtParam.Rows(0)("Valor"))
                If Not DtCont Is Nothing AndAlso DtCont.Rows.Count > 0 Then
                    StrContador = DtCont.Rows(0)("Contador")
                    StrIDContador = DtCont.Rows(0)("IDContador")
                Else : ApplicationService.GenerateError("El contador | no existe o no está correctamente configurado", DtCont.Rows(0)("IDContador"))
                End If
            Else : ApplicationService.GenerateError("El parámetro CNTRANS no tiene valor establecido.")
            End If
        Else : ApplicationService.GenerateError("El parámetro CNTRANS no existe.")
        End If

        For Each Dr As DataRow In data.Select("NotaTransportista IS NULL AND Empresa = '" & DescBBDDOriginal & "'")
            Dim DtAlbCab As DataTable = New AlbaranVentaCabecera().SelOnPrimaryKey(Dr("IDAlbaran"))
            DtAlbCab.Rows(0)("NotaTransportistaBaseDatos") = IDBaseDatosOriginal
            DtAlbCab.Rows(0)("NotaTransportista") = StrContador
            BusinessHelper.UpdateTable(DtAlbCab)
        Next

        Dim dtBBDD As DataTable = AdminData.GetUserDataBases
        For Each drBBDD As DataRow In dtBBDD.Rows
            AdminData.SetCurrentConnection(drBBDD("IDBaseDatos"))
            DescBBDD = AdminData.GetSessionConnection.Connection.Database
            If DescBBDDOriginal <> DescBBDD Then
                For Each Dr As DataRow In data.Select("NotaTransportista IS NULL AND Empresa = '" & DescBBDD & "'")
                    Dim DtAlbCab As DataTable = New AlbaranVentaCabecera().SelOnPrimaryKey(Dr("IDAlbaran"))
                    DtAlbCab.Rows(0)("NotaTransportistaBaseDatos") = IDBaseDatosOriginal
                    DtAlbCab.Rows(0)("NotaTransportista") = StrContador
                    BusinessHelper.UpdateTable(DtAlbCab)
                Next
            End If
        Next
        AdminData.SetCurrentConnection(IDBaseDatosOriginal)

        ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, StrIDContador, services)
        Return StrContador
    End Function

    <Task()> Public Shared Sub DesAsignarNotasTransportistaMultiEmpresa(ByVal data As DataTable, ByVal services As ServiceProvider)
        AdminData.BeginTx()
        Dim DescBBDD As String = AdminData.GetSessionInfo.DataBase.DataBaseDescription
        Dim DescBBDDOriginal As String = DescBBDD
        Dim IDBaseDatosOriginal As Guid = AdminData.GetConnectionInfo.IDDataBase

        For Each Dr As DataRow In data.Select("NotaTransportista IS NOT NULL AND Empresa = '" & DescBBDDOriginal & "'")
            Dim DtAlbCab As DataTable = New AlbaranVentaCabecera().SelOnPrimaryKey(Dr("IDAlbaran"))
            DtAlbCab.Rows(0)("NotaTransportista") = DBNull.Value
            DtAlbCab.Rows(0)("NotaTransportistaBaseDatos") = DBNull.Value
            BusinessHelper.UpdateTable(DtAlbCab)
        Next

        Dim dtBBDD As DataTable = AdminData.GetUserDataBases
        For Each drBBDD As DataRow In dtBBDD.Rows
            AdminData.SetCurrentConnection(drBBDD("IDBaseDatos"))
            DescBBDD = AdminData.GetSessionConnection.Connection.Database
            If DescBBDDOriginal <> DescBBDD Then
                For Each Dr As DataRow In data.Select("NotaTransportista IS NOT NULL AND Empresa = '" & DescBBDD & "'")
                    Dim DtAlbCab As DataTable = New AlbaranVentaCabecera().SelOnPrimaryKey(Dr("IDAlbaran"))
                    DtAlbCab.Rows(0)("NotaTransportista") = DBNull.Value
                    DtAlbCab.Rows(0)("NotaTransportistaBaseDatos") = DBNull.Value
                    BusinessHelper.UpdateTable(DtAlbCab)
                Next
            End If
        Next
        AdminData.SetCurrentConnection(IDBaseDatosOriginal)
    End Sub
#End Region
    <Task()> Public Shared Function ObtenerDatosInformes(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Return New DataEngine().Filter("xReport", Nothing, , , , True)
    End Function

    <Task()> Public Shared Sub ActualizarFicheroGeneradoEDI(ByVal IDAlbaran() As Object, ByVal services As ServiceProvider)
        If IDAlbaran.Length > 0 Then
            Dim dt As DataTable = New AlbaranVentaCabecera().Filter(New InListFilterItem("IDAlbaran", IDAlbaran, FilterType.Numeric, True))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                For Each dr As DataRow In dt.Rows
                    dr("GeneradoFichero") = True
                Next
                BusinessHelper.UpdateTable(dt)
            End If
        End If
    End Sub

End Class
