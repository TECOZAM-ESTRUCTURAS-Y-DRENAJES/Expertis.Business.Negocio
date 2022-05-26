Imports Solmicro.Expertis.Business.Negocio.NegocioGeneral

Public Class FacturaCompraCabeceraInfo
    Inherits ClassEntityInfo

    Public IDFactura As Integer
    Public NFactura As String
    Public FechaFactura As Date

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dt As DataTable
        If Not IsNothing(PrimaryKey) AndAlso PrimaryKey.Length > 0 AndAlso Length(PrimaryKey(0)) > 0 Then
            dt = New FacturaCompraCabecera().SelOnPrimaryKey(PrimaryKey(0))
        End If

        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            Me.Fill(dt.Rows(0))
        End If
    End Sub

End Class

Public Class FacturaCompraCabecera
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Private Const cnEntidad As String = "tbFacturaCompraCabecera"

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

#Region "Eventos RegisterAddNewTasks "

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
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarSuFechaFactura, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoComunes.AsignarFechaParaDeclaracion, data, services)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoComunes.AsignarEjercicioContableFactura, New DataRowPropertyAccessor(data), services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarTipoFactura, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarEstadoFactura, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.AsignarMarcaIVAManual, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.AsignarMarcaVtosManuales, data, services)
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim CE As New CentroEntidad
        CE.CentroGestion = data("IDCentroGestion") & String.Empty
        CE.ContadorEntidad = CentroGestion.ContadorEntidad.FacturaCompra
        data("IDContador") = ProcessServer.ExecuteTask(Of CentroEntidad, String)(AddressOf New CentroGestion().GetContadorPredeterminado, CE, services)
    End Sub

    <Task()> Public Shared Sub AsignarNumeroFacturaProvisional(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDContador")) > 0 Then
            Dim dtContadores As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDt, GetType(FacturaCompraCabecera).Name, services)
            Dim adr As DataRow() = dtContadores.Select("IDContador = " & Quoted(data("IDContador")))
            If Not IsNothing(adr) AndAlso adr.Length > 0 Then
                data("NFactura") = adr(0)("ValorProvisional")
            Else
                'Si no está bien configurado el Contador de Facturas de Compra en el Centro de Gestión,
                'cogemos el Contador por defecto de la entidad Factura Compra Cabecera.
                Dim dtContadorPred As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDefault, GetType(FacturaCompraCabecera).Name, services)
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

    <Task()> Public Shared Sub AsignarSuFechaFactura(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.IsNull("SuFechaFactura") Then data("SuFechaFactura") = Date.Today
        Dim AppParams As ParametroFacturaCompra = services.GetService(Of ParametroFacturaCompra)()
        If AppParams.ControlarFechaFCProveedor Then
            data("FechaFactura") = data("SuFechaFactura")
            data("FechaParaDeclaracion") = data("SuFechaFactura")
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoFacturacionCompra.FechaParaDeclaracionPorProveedor, New DataRowPropertyAccessor(data), services)
        End If
    End Sub

    <Task()> Public Shared Sub AsignarTipoFactura(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("TipoFactura") = enumfccTipoFactura.fccNormal
    End Sub

    <Task()> Public Shared Sub AsignarEstadoFactura(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("Estado") = CInt(enumfccEstado.fccNoContabilizado)
    End Sub

#End Region

#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.BeginTransaction)
        deleteProcess.AddTask(Of DataRow)(AddressOf ValidarDelExistenGastosAsociados)
        deleteProcess.AddTask(Of DataRow)(AddressOf ProcesoFacturacionCompra.ValidarFacturaContabilizada)
        deleteProcess.AddTask(Of DataRow)(AddressOf ProcesoFacturacionCompra.ValidarFacturaDeclarada)
        deleteProcess.AddTask(Of DataRow)(AddressOf ValidarDelFCompraControlSII)
        deleteProcess.AddTask(Of DataRow)(AddressOf ProcesoFacturacionCompra.ActualizarEntregasACuenta)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarBodegaVto)
    End Sub

    <Task()> Public Shared Sub ValidarDelExistenGastosAsociados(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDFactura", data("IDFactura")))
        Dim dtGastos As DataTable = New CobroFacturaCompra().Filter(f)
        If dtGastos.Rows.Count > 0 Then
            ApplicationService.GenerateError("No se puede eliminar la Factura. Contiene gastos asociados a algún Cobro.")
        End If

        f.Clear()
        f.Add(New NumberFilterItem("IDFacturaCompra", data("IDFactura")))
        Dim RCFC As New RemesaCobroFacturaCompra
        dtGastos = RCFC.Filter(f)
        If dtGastos.Rows.Count > 0 Then
            RCFC.Delete(dtGastos)
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDelFCompraControlSII(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter(FilterUnionOperator.Or)
        f.Add(New NumberFilterItem("IDFacturaCompra", data("IDFactura")))
        f.Add(New NumberFilterItem("IDFacturaPago", data("IDFactura")))
        f.Add(New NumberFilterItem("IDFacturaInversion", data("IDFactura")))
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

    <Task()> Public Shared Sub ComprobarBodegaVto(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim ClsBdg As BusinessHelper = CreateBusinessObject("BdgProveedorVto")
        Dim DtBdg As DataTable = ClsBdg.Filter(New FilterItem("IDFactura", FilterOperator.Equal, data("IDFactura")))
        If Not DtBdg Is Nothing AndAlso DtBdg.Rows.Count > 0 Then
            For Each Dr As DataRow In DtBdg.Select
                Dr("IDFactura") = DBNull.Value
            Next
            AdminData.SetData(DtBdg)
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of UpdatePackage, DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CrearDocumento)
        updateProcess.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarNumeroFactura)
        updateProcess.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarSuFactura)
        updateProcess.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoComunes.AsignarCentroGestion)
        updateProcess.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoComunes.ActualizarCambiosMoneda)
        updateProcess.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularImporteLineasFacturas)
        updateProcess.AddTask(Of DocumentoFacturaCompra)(AddressOf NegocioGeneral.CalcularImpuestos)
        updateProcess.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularAnaliticaFacturas)
        updateProcess.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularBasesImponibles)
        updateProcess.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularTotales)
        updateProcess.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularVencimientos)
        updateProcess.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.AsignarClaveOperacion)
'        updateProcess.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoComunes.ValidarIVASDocFC)
        updateProcess.AddTask(Of DocumentoFacturaCompra)(AddressOf Business.General.Comunes.BeginTransaction)
        updateProcess.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.ActualizarAlbaran)
        updateProcess.AddTask(Of DocumentoFacturaCompra)(AddressOf Business.General.Comunes.UpdateDocument)
        updateProcess.AddTask(Of DocumentoFacturaCompra)(AddressOf Business.General.Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.ActualizarObras)
        updateProcess.AddTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.ActualizarOTs)
    End Sub

    <Task()> Public Shared Sub CambiarEstadoFactura(ByVal data As DataTable, ByVal services As ServiceProvider)
        If data Is Nothing OrElse data.Rows.Count = 0 Then Exit Sub
        If data.Rows(0)("Estado") = enumfccEstado.fccContabilizado Then
            data.Rows(0)("Estado") = enumfccEstado.fccNoContabilizado
        Else
            data.Rows(0)("Estado") = enumfccEstado.fccContabilizado
        End If
        data.TableName = GetType(FacturaCompraCabecera).Name
        BusinessHelper.UpdateTable(data)
    End Sub


#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarProveedorObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFacturaObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarCIFObligatorio)
        'validateProcess.AddTask(Of DataRow)(AddressOf ProcesoFacturacionCompra.ValidarNumeroFactura)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoFacturacionCompra.ValidarSuFactura)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoFacturacionCompra.ValidarFacturaContabilizada)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFechaFacturaObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarEjercicioContableFactura)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarMonedaObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFormaPagoObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarCondicionPagoObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarClaveOperacion)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarDtoProntoPagoRecFinan)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoFacturacionCompra.ValidarFechaRetencion)
        'validateProcess.AddTask(Of DataRow)(AddressOf ValidarFechaFacturaAnterior)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFechaDeclaracion)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarEnvio347)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarEnvio349)
    End Sub

    '<Task()> Public Shared Sub ValidarFechaFacturaAnterior(ByVal data As DataRow, ByVal services As ServiceProvider)
    '    If New Parametro().ValidarCambioFechaFacturas Then
    '        Dim FilFacturas As New Filter
    '        FilFacturas.Add("IDContador", FilterOperator.Equal, data("IDContador"))
    '        FilFacturas.Add("IDEjercicio", FilterOperator.Equal, data("IDEjercicio"))
    '        FilFacturas.Add("FechaFactura", FilterOperator.GreaterThan, data("FechaFactura"))
    '        Dim DtFacturas As DataTable = New FacturaCompraCabecera().Filter(FilFacturas)
    '        If Not DtFacturas Is Nothing AndAlso DtFacturas.Rows.Count > 0 Then
    '            If data.RowState = DataRowState.Added Then
    '                ApplicationService.GenerateError("No se puede generar la factura con la fecha introducida. Existen facturas generadas posteriores a la fecha.")
    '            ElseIf data.RowState = DataRowState.Modified AndAlso Nz(data("FechaFactura")) <> Nz(data("FechaFactura", DataRowVersion.Original)) Then
    '                ApplicationService.GenerateError("No se puede modificar la fecha de la factura con la fecha introducida. Existen facturas generadas posteriores a la fecha.")
    '            End If
    '        End If
    '    End If
    'End Sub

#End Region

#Region " BUSINESSRULES "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim services As New ServiceProvider
        Dim oBRL As New SynonymousBusinessRules
        oBRL.AddSynonymous("FechaFactura", "Fecha")

        oBRL = ProcessServer.ExecuteTask(Of BusinessRules, BusinessRules)(AddressOf ProcesoCompra.DetailBusinessRulesCab, oBRL, services)

        oBRL("IDProveedor") = AddressOf CambioProveedor

        oBRL.Add("IDContador", AddressOf CambioContador)
        oBRL.Add("Fecha", AddressOf CambioFechaFactura) 'El nuevo nombre indicado en el Synonimous
        oBRL.Add("FechaDeclaracionManual", AddressOf CambioDeclaracionManual)
        oBRL.Add("IDFacturaRectificada", AddressOf CambioFacturaRectificada)
        oBRL.Add("BaseRetencion", AddressOf CambioBaseRetencion)
        oBRL.Add("SuFechaFactura", AddressOf CambioSuFechaFactura)
        oBRL.Add("IDDireccion", AddressOf CambioIdDireccion)
        oBRL.Add("Enviar347", AddressOf CambioEnviar347)
        oBRL.Add("Enviar349", AddressOf CambioEnviar349)
        oBRL.Add("RetencionIRPF", AddressOf CambioRetencionIRPF)
        oBRL.Add("CIFProveedor", AddressOf CambioCifProveedor)
        oBRL.Add("IDPais", AddressOf CambioIDPais)
        Return oBRL
    End Function


    <Task()> Public Shared Sub CambioCifProveedor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Nz(data.Current("Enviar349"), False) Then
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf NegocioGeneral.ValidarEnvio349IPropAcc, data.Current, services)
        End If
'        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoComunes.ValidarDocumentoIdentificativoFC, data.Current, services)
    End Sub
    <Task()> Public Shared Sub CambioIDPais(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.AsignarEnvio347, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.AsignarEnvio349, data, services)
        If Nz(data.Current("Enviar349"), False) Then
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf NegocioGeneral.ValidarEnvio349IPropAcc, data.Current, services)
        End If
'        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoComunes.ValidarDocumentoIdentificativoFC, data.Current, services)
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


    <Task()> Public Shared Sub CambioSuFechaFactura(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim AppParams As ParametroFacturaCompra = services.GetService(Of ParametroFacturaCompra)()
        If AppParams.ControlarFechaFCProveedor Then
            data.Current(data.ColumnName) = data.Value
            If Nz(data.Current("FechaFactura"), cnMinDate) <> Nz(data.Current("SuFechaFactura"), cnMinDate) Then
                data.Current("FechaFactura") = data.Current("SuFechaFactura")
                If Length(data.Current("SuFechaFactura")) > 0 AndAlso IsDate(data.Current("SuFechaFactura")) Then
                    Dim FCC As New FacturaCompraCabecera
                    FCC.ApplyBusinessRule("FechaFactura", data.Current("SuFechaFactura"), data.Current)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioIdDireccion(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Nz(data.Current("IDDireccion"), 0) <> 0 Then
            Dim PD As New ProveedorDireccion
            Dim dtDirProveedor As DataTable = PD.SelOnPrimaryKey(data.Current("IDDireccion"))
            If Not IsNothing(dtDirProveedor) AndAlso dtDirProveedor.Rows.Count > 0 Then
                data.Current("DirFacturacion") = dtDirProveedor.Rows(0)("RazonSocial") & vbNewLine & _
                                                 dtDirProveedor.Rows(0)("Direccion") & vbNewLine & _
                                                 dtDirProveedor.Rows(0)("Provincia") & " " & _
                                                 dtDirProveedor.Rows(0)("CodPostal") & " " & _
                                                 dtDirProveedor.Rows(0)("Poblacion")
            End If
        Else
            data.Current("DirFacturacion") = String.Empty
        End If
    End Sub

    <Task()> Public Shared Sub CambioContador(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.ColumnName = "IDContador" Then data.Current(data.ColumnName) = data.Value

        If IsDate(data.Current("FechaFactura")) AndAlso Not IsDBNull(data.Current("IDContador")) Then
            Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
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

    <Task()> Public Shared Sub CambioProveedor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)

        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoCompra.CambioProveedor, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComunes.CambioMoneda, data, services)

        Dim dir As New DataDireccionProv(enumpdTipoDireccion.pdDireccionFactura, "IDDireccion", data.Current)
        ProcessServer.ExecuteTask(Of DataDireccionProv)(AddressOf ProcesoCompra.AsignarDireccionProveedor, dir, services)

        Dim Obs As New DataObservaciones(GetType(FacturaCompraCabecera).Name, "Texto", data.Current)
        ProcessServer.ExecuteTask(Of DataObservaciones)(AddressOf ProcesoCompra.AsignarObservacionesProveedor, Obs, services)

        If Length(data.Current("IDProveedor")) > 0 Then
            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
            Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data.Current("IDProveedor"))
            data.Current("IDPais") = ProvInfo.IDPais
            data.Current("Telefono") = ProvInfo.Telefono
            data.Current("Fax") = ProvInfo.Fax
            data.Current("IDModoTransporte") = ProvInfo.IDModoTransporte
            data.Current("IdBancoPropio") = ProvInfo.IDBancoPropio
            data.Current("DtoFactura") = ProvInfo.DtoComercial
            data.Current("IDTipoAsiento") = ProvInfo.IDTipoAsiento
            data.Current("RetencionIRPF") = ProvInfo.RetencionIRPF
            data.Current("RegimenEspecial") = ProvInfo.RegimenEspecial
            data.Current("TipoRetencionIRPF") = ProvInfo.TipoRetencionIRPF
            If Length(ProvInfo.IDTipoClasif) > 0 Then
                Dim DtClasif As DataTable = New TipoClasif().SelOnPrimaryKey(ProvInfo.IDTipoClasif)
                If Not DtClasif Is Nothing AndAlso DtClasif.Rows.Count > 0 Then
                    data.Current("TipoFactura") = DtClasif.Rows(0)("IDTipoFactura")
                Else : data.Current("TipoFactura") = enumfccTipoFactura.fccNormal
                End If
            Else : data.Current("TipoFactura") = enumfccTipoFactura.fccNormal
            End If
            Dim IDBanco As Integer = ProcessServer.ExecuteTask(Of String, Integer)(AddressOf New ProveedorBanco().GetBancoPredeterminado, data.Current("IDProveedor"), services)
            If IDBanco > 0 Then
                data.Current("IDProveedorBanco") = IDBanco
            Else
                data.Current("IDProveedorBanco") = System.DBNull.Value
            End If

        Else
            data.Current("IDPais") = System.DBNull.Value
            data.Current("Telefono") = System.DBNull.Value
            data.Current("Fax") = System.DBNull.Value
            data.Current("IDModoTransporte") = System.DBNull.Value
            data.Current("IdBancoPropio") = System.DBNull.Value
            data.Current("DtoFactura") = 0
            data.Current("IDTipoAsiento") = System.DBNull.Value
            data.Current("RetencionIRPF") = 0
            data.Current("TipoRetencionIRPF") = System.DBNull.Value
            data.Current("TipoFactura") = enumfccTipoFactura.fccNormal
            data.Current("IDProveedorBanco") = System.DBNull.Value
        End If

        'If Not data.Current("FechaDeclaracionManual") AndAlso data.Current.Contains("IDProveedor") AndAlso Length(data.Current("IDProveedor")) > 0 Then
        '    Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
        '    Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data.Current("IDProveedor"))
        '    If ProvInfo.IVACaja Then
        '        data.Current("FechaParaDeclaracion") = cnMAX_DATE
        '    End If
        'End If
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoFacturacionCompra.FechaParaDeclaracionPorProveedor, data.Current, services)

        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.AsignarEnvio347, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.AsignarEnvio349, data, services)
'        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoComunes.ValidarDocumentoIdentificativoFC, data.Current, services)
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
        If data.ColumnName = "FechaDeclaracionManual" Then data.Current(data.ColumnName) = data.Value
        If Not data.Current("FechaDeclaracionManual") Then
            data.Current("FechaParaDeclaracion") = data.Current("FechaFactura")
            ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoFacturacionCompra.FechaParaDeclaracionPorProveedor, data.Current, services)
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
    

    'Private Function ExistenLineasFactura(ByVal intIDFactura As Integer) As Boolean
    '    Dim dtFCL As DataTable = New FacturaCompraLinea().Filter(New NumberFilterItem("IDFactura", intIDFactura))
    '    Return (Not IsNothing(dtFCL) AndAlso dtFCL.Rows.Count > 0)
    'End Function

#End Region

#Region " Declaraciones "

    <Task()> Public Shared Sub DeclararIVACompra(ByVal data As DataDeclaraciones, ByVal services As ServiceProvider)
        '//Asigna el NDeclaracionIVA y el AñoDeclaracionIVA a las Facturas que se le indica mediante el Filtro
        If data.Filtro Is Nothing OrElse data.Filtro.Count = 0 Then
            ApplicationService.GenerateError("Debe indicar un filtro para seleccionar las Facturas a Declarar.")
        End If

        AdminData.Execute("sp_DeclaracionIVACompra", False, data.NDeclaracion, data.AnioDeclaracion, AdminData.ComposeFilter(data.Filtro))
    End Sub

    Public Sub RenumerarNFacturaIVA(ByVal intAñoDecl As Integer, ByVal intNDecl As Integer, _
                                    ByVal BlnOrdenFecha As Boolean, ByVal intPrimerNum As Integer, _
                                    ByVal StrIDContador As String)

        '//Construimos la clausula Where por la que se van a filtrar las facturas a renumerar
        Dim objFilter As New Filter
        objFilter.Add(New NumberFilterItem("AñoDeclaracionIVA", intAñoDecl))
        objFilter.Add(New NumberFilterItem("NDeclaracionIVA", intNDecl))

        '//Si se ha rellenado el campo contador, se escogen los registros con ese contador
        If Length(StrIDContador) > 0 Then objFilter.Add(New StringFilterItem("IDContador", StrIDContador))

        '//Obtenemos la clausula ORDER BY
        Dim StrOrder As String = IIf(BlnOrdenFecha, "IDContador, FechaParaDeclaracion, NFactura", "IDContador, NFactura")


        Dim intNFacturaIVA As Integer
        '//Cogemos de la tabla las Facturas que cumplen la condicion, en el orden indicado
        Dim dtFact As DataTable = New BE.DataEngine().Filter("vNegFacturasCompraAIVA", objFilter, , StrOrder)
        dtFact.TableName = "FacturaCompraCabecera"
        If Not dtFact Is Nothing AndAlso dtFact.Rows.Count > 0 Then
            If Length(StrIDContador) > 0 Then
                '//Asignamos el NFacturaIVA a las facturas seleccionadas(renumeramos sin más)
                intNFacturaIVA = intPrimerNum
                For Each Dr As DataRow In dtFact.Select(Nothing, StrOrder)
                    Dr("NFacturaIVA") = intNFacturaIVA
                    intNFacturaIVA += 1
                Next
            Else
                Dim IDContador As String
                Dim objFilterFactRenumerar As New Filter
                Dim dtFactAux As DataTable = dtFact.Copy
                dtFact.DefaultView.Sort = StrOrder
                '//Se renumeran por separado según el contador
                For Each drFactura As DataRow In dtFactAux.Select(Nothing, StrOrder)
                    If Length(IDContador) = 0 OrElse IDContador <> drFactura("IDContador") & String.Empty Then
                        IDContador = drFactura("IDContador") & String.Empty
                        intNFacturaIVA = intPrimerNum
                        objFilterFactRenumerar.Clear()
                        If Length(IDContador) > 0 Then
                            objFilterFactRenumerar.Add(New StringFilterItem("IDContador", IDContador))
                        Else
                            objFilterFactRenumerar.Add(New IsNullFilterItem("IDContador"))
                        End If
                        dtFact.DefaultView.RowFilter = objFilterFactRenumerar.Compose(New AdoFilterComposer)
                        For Each dr As DataRowView In dtFact.DefaultView
                            dr.Row("NFacturaIVA") = intNFacturaIVA
                            intNFacturaIVA += 1
                        Next
                        dtFact.DefaultView.RowFilter = Nothing
                    End If
                Next
            End If
            BusinessHelper.UpdateTable(dtFact)
        Else
            ApplicationService.GenerateError("No existe ninguna Factura que cumpla los criterios.")
        End If
    End Sub

    Public Function ContadorAutofactura() As String
        Dim StrContador As String = New Parametro().ContadorAutofactura
        If Length(StrContador) > 0 Then
            ' Coger valor del contador y el nº de factura predeterminado.
            ContadorAutofactura = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, StrContador, New ServiceProvider)
        End If
    End Function

    Public Sub ActualizarRegimenEspecial(ByVal drProveedor As DataRow)
        Dim dt As DataTable = drProveedor.Table
        If dt.Columns.Contains("IdProveedor") And dt.Columns.Contains("RegimenEspecial") Then
            Dim dtFCC As DataTable = Filter(New StringFilterItem("IDProveedor", drProveedor("IdProveedor")))
            If Not dtFCC Is Nothing AndAlso dtFCC.Rows.Count > 0 Then
                For Each drFCC As DataRow In dtFCC.Rows
                    If drFCC("estado") = enumfccEstado.fccNoContabilizado Then
                        drFCC("RegimenEspecial") = drProveedor("RegimenEspecial")
                    End If
                Next
                Update(dtFCC)
            End If
        End If
    End Sub
#End Region

#Region " Intrastat "

    Public Function ObtenerListasInstrastatCompra() As DataTable
        Return New BE.DataEngine().Filter("vCtlCIIntrastatCompraListas", "*", "")
    End Function

    Public Function ObtenerDatosRptIntrastatCompra(ByVal f As Filter) As DataTable
        Return New BE.DataEngine().Filter("vRptIntrastatCompra", f)
    End Function

    Public Function GrabarDeclaracionIntrastat(ByVal StrIDFacturas As String, ByVal StrNDeclaracion As String, _
                                           ByVal StrAñoDeclaracion As String) As Integer
        Dim DtFacturas As DataTable = Me.Filter("*", StrIDFacturas)
        If Not DtFacturas Is Nothing AndAlso DtFacturas.Rows.Count <> 0 Then
            For Each Dr As DataRow In DtFacturas.Select
                If StrNDeclaracion = "0000" Then
                    Dr("NDeclaracionIntrastat") = System.DBNull.Value
                Else
                    Dr("NDeclaracionIntrastat") = StrNDeclaracion
                End If
                If StrAñoDeclaracion = "0000" Then
                    Dr("AñoDeclaracionIntrastat") = System.DBNull.Value
                Else
                    Dr("AñoDeclaracionIntrastat") = StrAñoDeclaracion
                End If
            Next
            Me.Update(DtFacturas)
        End If
    End Function

#End Region

#Region " Facturacion "

#Region " Propuesta de Facturas "

    Public Function PropuestaFacturaCompraRealquiler(ByVal Dt As DataTable, ByVal strIDContador As String, _
                                                    ByVal dFechaFacturacion As Date) As DataSet
        'TODO: Cuando se haga con Documentos esto se quitará
        'Return PropuestaFacturaCompra(Dt, strIDContador, enumTipoFactCompra.tfcNormal, dFechaFacturacion)
    End Function

    'Public Function PropuestaFacturaCompra(ByVal Dt As DataTable, ByVal strIDContador As String, _
    '                                        Optional ByVal IntTipoFact As enumTipoFactCompra = enumTipoFactCompra.tfcNormal, _
    '                                        Optional ByVal dFechaFacturacion As Date = cnMinDate) As DataSet

    '    Dim com As New Compra
    '    Dim dtCabecera As DataTable
    '    Dim dtLineas As DataTable
    '    Dim dtEntregas As DataTable
    '    Dim fcl As New FacturaCompraLinea
    '    Dim Facturas As DataTable = Me.AddNew
    '    Dim Lineas As DataTable = fcl.AddNew
    '    Dim DtLineasF As DataTable = Lineas.Clone
    '    Dim Albaranes As DataTable
    '    Dim dtOrigen As DataTable
    '    Dim aca As New AlbaranCompraAnalitica
    '    Dim Analitica As DataTable
    '    Dim fca As New FacturaCompraAnalitica
    '    Dim dtPagoLeasing As DataTable

    '    Dim services As New ServiceProvider

    '    If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
    '        Dim dteSuFechaFactura As Date
    '        Dim strSuFactura As String
    '        If Not Dt.Rows(0)("SuFechaFactura") Is Nothing Then
    '            If Length(Dt.Rows(0)("SuFechaFactura")) > 0 Then
    '                dteSuFechaFactura = Dt.Rows(0)("SuFechaFactura")
    '            Else
    '                dteSuFechaFactura = cnMinDate
    '            End If
    '        End If
    '        If Not Dt.Rows(0)("SuFactura") Is Nothing Then
    '            If Length(Dt.Rows(0)("SuFactura")) > 0 Then
    '                strSuFactura = Dt.Rows(0)("SuFactura")
    '            Else
    '                strSuFactura = String.Empty
    '            End If
    '        End If

    '        Select Case IntTipoFact
    '            Case enumTipoFactCompra.tfcNormal
    '                dtLineas = AgruparAlbaranes(Dt, dtCabecera, strIDContador, dFechaFacturacion)
    '            Case enumTipoFactCompra.tfcConcesion
    '                dtLineas = AgruparConcesion(Dt, dtCabecera, strIDContador)
    '            Case enumTipoFactCompra.tfcLeasing
    '                dtLineas = AgruparPagos(Dt, dtCabecera, strIDContador)
    '                strSuFactura = dtCabecera.Rows(0)("SuFactura")
    '                If IsNothing(dtPagoLeasing) Then dtPagoLeasing = New DataTable
    '                dtPagoLeasing.Columns.Add("IDPago", GetType(Integer))
    '                dtPagoLeasing.Columns.Add("IDFactura", GetType(Integer))
    '                dtPagoLeasing.Columns.Add("NFactura", GetType(String))
    '        End Select

    '        If Not dtLineas Is Nothing AndAlso dtLineas.Rows.Count > 0 Then
    '            Dim ds As New DataSet
    '            If Not dtCabecera Is Nothing AndAlso dtCabecera.Rows.Count > 0 Then
    '                Dim TiposIva As New EntityInfoCache(Of TipoIvaInfo)
    '                Dim Proveedores As New EntityInfoCache(Of ProveedorInfo)
    '                Dim Monedas As New MonedaCache
    '                Dim AppParams As New ParametroCache
    '                Dim Factura As DataRow
    '                Dim oCont As New FCCounter
    '                For Each drCabecera As DataRow In dtCabecera.Rows
    '                    Select Case IntTipoFact
    '                        Case enumTipoFactCompra.tfcNormal
    '                            Factura = NuevaCabeceraFactura(drCabecera, oCont, , , strSuFactura, dteSuFechaFactura, services, AppParams)
    '                        Case enumTipoFactCompra.tfcConcesion
    '                            Factura = NuevaCabeceraFactura(drCabecera, oCont, 1, True, , , services, AppParams)
    '                        Case enumTipoFactCompra.tfcLeasing
    '                            Factura = NuevaCabeceraFactura(drCabecera, oCont, 1, True, , , services, AppParams)
    '                    End Select
    '                    If Not Factura Is Nothing Then
    '                        Dim oFilter As New Filter
    '                        oFilter.Add(New NumberFilterItem("NRegistro", FilterOperator.Equal, drCabecera("NRegistro")))
    '                        Dim dvlineas As DataView = dtLineas.DefaultView
    '                        dvlineas.RowFilter = oFilter.Compose(New AdoFilterComposer)
    '                        Lineas = Lineas.Clone
    '                        Select Case IntTipoFact
    '                            Case enumTipoFactCompra.tfcNormal
    '                                Lineas = NuevaLineaFactura(Factura("IDFactura"), Factura("NFactura"), _
    '                                                           dvlineas, Factura("FechaFactura"), Factura("IDProveedor"), Lineas, dFechaFacturacion, services)

    '                            Case enumTipoFactCompra.tfcConcesion
    '                                Lineas = NuevaLineaFacturaConcesion(Factura("IDFactura"), Factura("NFactura"), _
    '                                                  dvlineas, Factura("FechaFactura"), Factura("IDProveedor"), Lineas, Factura("IDCentroGestion"), Factura("IDMoneda"), Monedas)
    '                            Case enumTipoFactCompra.tfcLeasing
    '                                Lineas = NuevaLineaFacturaLeasing(Factura("IDFactura"), Factura("NFactura"), _
    '                                                  dvlineas, Factura("FechaFactura"), Factura("IDProveedor"), _
    '                                                  Lineas, Factura("IDCentroGestion"), Factura("IDMoneda"), dtPagoLeasing, Monedas)
    '                        End Select

    '                        If Not Lineas Is Nothing AndAlso Lineas.Rows.Count > 0 Then
    '                            For Each linea As DataRow In Lineas.Rows
    '                                '///Recalcular importes albaran
    '                                CalcularPrecioImporte(linea, TiposIva)
    '                                MantenimientoValoresAyB(linea, Factura("IDMoneda"), Factura("CambioA"), Factura("CambioB"), Monedas)
    '                            Next
    '                        End If
    '                    End If
    '                    Factura = ImportePropuesta(Factura, Lineas, services)
    '                    Facturas.Rows.Add(Factura.ItemArray)
    '                    If IntTipoFact = enumTipoFactCompra.tfcNormal Then
    '                        NuevaFacturaAnalitica(Lineas, Factura("CambioA"), Factura("CambioB"), Factura("IDMoneda"), Analitica, Monedas)
    '                        PropuestaEntregasFacturaCompra(Factura, Lineas, dtEntregas)
    '                    End If
    '                    For Each Dr As DataRow In Lineas.Select
    '                        DtLineasF.Rows.Add(Dr.ItemArray)
    '                    Next
    '                Next
    '            End If

    '            ds.Tables.Add(Facturas)
    '            If IntTipoFact = enumTipoFactCompra.tfcNormal Then
    '                ds.Tables.Add(dtEntregas)
    '            End If
    '            ds.Tables.Add(DtLineasF)
    '            If IntTipoFact = enumTipoFactCompra.tfcNormal AndAlso Not IsNothing(Analitica) Then
    '                ds.Tables.Add(Analitica)
    '            End If
    '            If IntTipoFact = enumTipoFactCompra.tfcLeasing Then
    '                Dim dtPago As DataTable = ActualizarPago(dtPagoLeasing)
    '                ds.Tables.Add(dtPago)
    '            End If
    '            Return ds
    '        End If
    '    End If
    'End Function

    'Private Sub PropuestaEntregasFacturaCompra(ByVal drCabeceraFact As DataRow, ByVal dtLineasFact As DataTable, ByRef dtEntregas As DataTable)
    '    '//Creamos una línea en Entregas a Cuenta, de tipo Retención.
    '    Dim objNegEC As New EntregasACuenta
    '    If IsNothing(dtEntregas) Then dtEntregas = objNegEC.AddNew
    '    Dim drNuevaEntrega As DataRow = objNegEC.NuevaEntregaTipoRetencionFacturaObra(drCabeceraFact, dtLineasFact, dtEntregas, Circuito.Compras)

    '    If Not IsNothing(drNuevaEntrega) Then
    '        '//Creamos una línea de factura con la entrega
    '        Dim dtNuevaEntregaAux As DataTable = dtEntregas.Clone
    '        dtNuevaEntregaAux.Rows.Add(drNuevaEntrega.ItemArray)

    '        '//Creamos una nueva línea de factura de tipo Retención
    '        dtLineasFact = objNegEC.NuevaLineaFacturaEntregaCuenta(drCabeceraFact, dtLineasFact, dtNuevaEntregaAux, Circuito.Compras)
    '    End If
    'End Sub

#Region " Agruoaciones "

    Private Function AgruparPagos(ByVal Dt As DataTable, ByRef DtCondiciones As DataTable, _
                                  ByVal StrContador As String) As DataTable
        Dim ClsPago As New Pago
        Dim ClsPagoPer As New PagoPeriodico
        Dim DtLineas As New DataTable
        Dim DtFacturaLin As New DataTable
        Dim StrIN, StrWhere, StrOrder As String
        Dim StrIDProveedor, StrIDMoneda, StrIDFormaPago, StrIDCondicionPago, StrContadorCargo, StrIDInmov As String
        Dim IntIdPago, IntContador, IntContadorAux, IntNRegistro As Integer
        Dim DblImporteLinea, DblPagoPeriodo As Double

        'se seleccionan todas las lineas de albaran no facturadas
        For Each Dr As DataRow In Dt.Select
            If Length(StrIN) Then StrIN &= ","
            StrIN &= Dr("IDPago")
        Next

        StrWhere = "IDPago in (" & StrIN & ") "
        StrOrder = "IDProveedor, IDMoneda, IDFormaPago, IDCondicionPago, IDPago,FechaVencimiento"
        DtLineas = New BE.DataEngine().Filter("vNegCompraCrearFacturaLeasing", "*", StrWhere)

        If Not DtLineas Is Nothing AndAlso DtLineas.Rows.Count Then
            If IsNothing(DtCondiciones) Then DtCondiciones = New DataTable
            With DtCondiciones
                .Columns.Add("IDProveedor", GetType(String))
                .Columns.Add("IDFormaPago", GetType(String))
                .Columns.Add("IDCondicionPago", GetType(String))
                .Columns.Add("IdMoneda", GetType(String))
                .Columns.Add("IdFactura", GetType(Integer))
                .Columns.Add("NFactura", GetType(String))
                .Columns.Add("IDCentroGestion", GetType(String))
                .Columns.Add("NRegistro", GetType(Integer))
                .Columns.Add("IDPago", GetType(Integer))
                .Columns.Add("FechaFactura", GetType(Date))
                .Columns.Add("IDContador", GetType(String))
                .Columns.Add("Dto", GetType(Double))
                .Columns.Add("CifProveedor", GetType(String))
                .Columns.Add("RazonSocial", GetType(String))
                .Columns.Add("IDDiaPago", GetType(String))
                .Columns.Add("IDBancoPropio", GetType(String))
                .Columns.Add("FacturaPagoPeriodicoSN", GetType(Boolean))
                .Columns.Add("IDProveedorBanco", GetType(Integer))
                .Columns.Add("SuFactura", GetType(String))
                .Columns.Add("IDDireccion", GetType(Integer))
                .Columns.Add("IDObra", GetType(Integer))
                '.Columns.Add("Retencion", GetType(Integer))
                '.Columns.Add("TipoRetencion", GetType(Integer))
                '.Columns.Add("FechaRetencion", GetType(Date))
                '.Columns.Add("Periodo", GetType(Integer))
                '.Columns.Add("TipoPeriodo", GetType(Integer))
            End With
            With DtFacturaLin
                .Columns.Add("NRegistro", GetType(Integer))
                .Columns.Add("IDPago", GetType(Integer))
            End With
        End If
        For Each Dr As DataRow In DtLineas.Select
            If StrIDProveedor <> Dr("IDProveedor") OrElse IntIdPago <> Dr("IDPago") _
            Or StrIDMoneda <> Dr("IDMoneda") OrElse StrIDFormaPago <> Dr("IDFormaPago") _
            Or StrIDCondicionPago <> Dr("IDCondicionPago") Then
                If DblPagoPeriodo <> Nz(Dr("IdPagoPeriodo"), 0) Then
                    IntContadorAux = 1
                    Dim FilPeriodo As New Filter
                    FilPeriodo.Add("IDPagoPeriodo", FilterOperator.Equal, Dr("IDPagoPeriodo"), FilterType.Numeric)
                    FilPeriodo.Add(New BooleanFilterItem("Contabilizado", True))
                    Dim DtPago As DataTable = ClsPago.Filter(FilPeriodo)
                    If Not DtPago Is Nothing AndAlso DtPago.Rows.Count > 0 Then
                        IntContadorAux += DtPago.Rows.Count
                    End If
                    Dim DtPagoPer As DataTable = ClsPagoPer.Filter(New FilterItem("ID", FilterOperator.Equal, Dr("IDPagoPeriodo"), FilterType.Numeric))
                    If Not DtPagoPer Is Nothing AndAlso DtPagoPer.Rows.Count > 0 Then
                        StrIDInmov = DtPagoPer.Rows(0)("IDInmovilizado")
                    End If
                Else
                    IntContadorAux += 1
                End If

                DblPagoPeriodo = Nz(Dr("IDpagoPeriodo"), 0)

                Dim StrSuFactura As String = StrIDInmov & "/" & CStr(IntContadorAux)

                StrIDProveedor = Dr("IDProveedor")
                IntIdPago = Dr("IDPago")
                StrIDMoneda = Dr("IDMoneda")
                StrIDFormaPago = Dr("IDFormaPago")
                StrIDCondicionPago = Dr("IDCondicionPago")
                DblImporteLinea = Dr("ImpCuota")
                StrContadorCargo = Dr("IDContadorCargo") & String.Empty

                IntContador += 1
                Dim DrNew As DataRow = DtCondiciones.NewRow()
                DrNew("SuFactura") = StrSuFactura
                DrNew("IDProveedor") = StrIDProveedor
                DrNew("IDPago") = IntIdPago
                DrNew("IDMoneda") = StrIDMoneda
                DrNew("IDFormaPago") = StrIDFormaPago
                DrNew("RazonSocial") = Dr("RazonSocial") & String.Empty
                DrNew("CifProveedor") = Dr("CifProveedor") & String.Empty
                DrNew("IDDiaPago") = Dr("IDDiaPago") & String.Empty
                DrNew("IDBancoPropio") = Dr("IDBancoPropio") & String.Empty
                DrNew("FacturaPagoPeriodicoSN") = 1
                DrNew("IDCondicionPago") = StrIDCondicionPago
                DrNew("NRegistro") = IntContador
                DrNew("FechaFactura") = Dr("FechaVencimiento")
                DrNew("IDContador") = IIf(Length(StrContadorCargo) > 0, StrContadorCargo, StrContador)
                DrNew("Dto") = 0
                Dim ClsBanco As New ProveedorBanco
                Dim FilProv As New Filter
                FilProv.Add("Predeterminado", FilterOperator.Equal, 1, FilterType.Boolean)
                FilProv.Add("IdProveedor", FilterOperator.Equal, Dr("IDProveedor"), FilterType.String)
                Dim DtBanco As DataTable = ClsBanco.Filter(FilProv)
                If Not DtBanco Is Nothing AndAlso DtBanco.Rows.Count > 0 Then
                    DrNew("IDProveedorBanco") = DtBanco.Rows(0)("IDProveedorBanco") & DBNull.Value
                End If

                DtCondiciones.Rows.Add(DrNew)
                IntNRegistro = IntContador
            End If
            Dim DrNewfact As DataRow = DtFacturaLin.NewRow()
            DrNewfact("NRegistro") = IntNRegistro
            DrNewfact("IDPago") = Dr("IDPago")
            DtFacturaLin.Rows.Add(DrNewfact)
        Next
        Return DtFacturaLin
    End Function

    Private Function AgruparConcesion(ByVal Dt As DataTable, ByRef DtCondiciones As DataTable, _
                                      ByVal StrContador As String) As DataTable
        Dim DtLineas As New DataTable
        Dim DtFacturaLin As New DataTable
        Dim StrIN, StrWhere, StrOrder As String
        Dim StrIDProveedor, StrIDMoneda, StrIDFormaPago, StrIDCondicionPago, StrContadorCargo As String
        Dim IntIdPagoPer, IntContador, IntNRegistro As Integer
        Dim DblImporteLinea As Double

        'se seleccionan todas las lineas de albaran no facturadas
        For Each Dr As DataRow In Dt.Select
            If Length(StrIN) > 0 Then StrIN &= ","
            StrIN &= Dr("IDPagoPer")
        Next

        StrWhere = "ID in (" & StrIN & ") "
        StrOrder = "IDProveedor, IDMoneda, IDFormaPago, IDCondicionPago, ID,FechaVencimiento"
        DtLineas = New BE.DataEngine().Filter("vNegCompraCrearFacturaConcesion", "*", StrWhere)

        If Not DtLineas Is Nothing AndAlso DtLineas.Rows.Count > 0 Then
            If IsNothing(DtCondiciones) Then DtCondiciones = New DataTable
            With DtCondiciones
                .Columns.Add("IDProveedor", GetType(String))
                .Columns.Add("IDFormaPago", GetType(String))
                .Columns.Add("IDCondicionPago", GetType(String))
                .Columns.Add("IdMoneda", GetType(String))
                .Columns.Add("IdFactura", GetType(Integer))
                .Columns.Add("NFactura", GetType(String))
                .Columns.Add("IDCentroGestion", GetType(String))
                .Columns.Add("NRegistro", GetType(Integer))
                .Columns.Add("IDPagoPer", GetType(Integer))
                .Columns.Add("FechaFactura", GetType(Date))
                .Columns.Add("IDContador", GetType(String))
                .Columns.Add("Dto", GetType(Double))
                .Columns.Add("CifProveedor", GetType(String))
                .Columns.Add("RazonSocial", GetType(String))
                .Columns.Add("IDDiaPago", GetType(String))
                .Columns.Add("IDBancoPropio", GetType(String))
                .Columns.Add("FacturaPagoPeriodicoSN", GetType(Boolean))
                .Columns.Add("IDProveedorBanco", GetType(Integer))
                .Columns.Add("IDDireccion", GetType(Integer))
                .Columns.Add("IDObra", GetType(Integer))
                '.Columns.Add("TipoRetencion", GetType(Integer))
                '.Columns.Add("Retencion", GetType(Integer))
                '.Columns.Add("FechaRetencion", GetType(Date))
                '.Columns.Add("Periodo", GetType(Integer))
                '.Columns.Add("TipoPeriodo", GetType(Integer))
            End With
            With DtFacturaLin
                .Columns.Add("NRegistro", GetType(Integer))
                .Columns.Add("IDPagoPer", GetType(Integer))
            End With
        End If
        For Each Dr As DataRow In DtLineas.Select
            If Not AreEquals(StrIDProveedor, Dr("IDProveedor") & String.Empty) _
            OrElse Not AreEquals(IntIdPagoPer, Dr("ID")) _
            OrElse Not AreEquals(StrIDMoneda, Dr("IDMoneda") & String.Empty) _
            OrElse Not AreEquals(StrIDFormaPago, Dr("IDFormaPago") & String.Empty) _
            OrElse Not AreEquals(StrIDCondicionPago, Dr("IDCondicionPago") & String.Empty) Then
                StrIDProveedor = Dr("IDProveedor") & String.Empty
                IntIdPagoPer = Dr("ID")
                StrIDMoneda = Dr("IDMoneda") & String.Empty
                StrIDFormaPago = Dr("IDFormaPago") & String.Empty
                StrIDCondicionPago = Dr("IDCondicionPago") & String.Empty
                DblImporteLinea = Nz(Dr("ImpRecuperacionCoste"), 0)
                StrContadorCargo = Dr("IDContadorCargo") & String.Empty
                IntContador += 1
                Dim DrNew As DataRow = DtCondiciones.NewRow()
                DrNew("IDProveedor") = StrIDProveedor
                DrNew("IDPagoPer") = IntIdPagoPer
                DrNew("IDMoneda") = StrIDMoneda
                DrNew("IDFormaPago") = IntIdPagoPer
                DrNew("RazonSocial") = Dr("RazonSocial") & String.Empty
                DrNew("CifProveedor") = Dr("CifProveedor") & String.Empty
                DrNew("IDDiaPago") = Dr("IDDiaPago") & String.Empty
                DrNew("IDBancoPropio") = Dr("IDBancoPropio") & String.Empty
                DrNew("FacturaPagoPeriodicoSN") = 1
                DrNew("IDCondicionPago") = StrIDCondicionPago
                DrNew("NRegistro") = IntContador
                DrNew("FechaFactura") = Today.Date
                DrNew("IDContador") = IIf(Length(StrContadorCargo) > 0, StrContadorCargo, StrContador)
                DrNew("Dto") = 0
                Dim ClsBanco As New ProveedorBanco
                Dim FilProv As New Filter
                FilProv.Add("IdProveedor", FilterOperator.Equal, Dr("IDProveedor"), FilterType.String)
                FilProv.Add("Predeterminado", FilterOperator.Equal, 1, FilterType.Boolean)
                Dim DtBanco As DataTable = ClsBanco.Filter(FilProv)
                If Not DtBanco Is Nothing AndAlso DtBanco.Rows.Count > 0 Then
                    DrNew("IDProveedorBanco") = DtBanco.Rows(0)("IDProveedorBanco") & DBNull.Value
                End If

                DtCondiciones.Rows.Add(DrNew)
                IntNRegistro = IntContador
            End If
            Dim DrNewFact As DataRow = DtFacturaLin.NewRow()
            DrNewFact("NRegistro") = IntNRegistro
            DrNewFact("IDPagoPer") = Dr("ID")
            DtFacturaLin.Rows.Add(DrNewFact)
        Next
        Return DtFacturaLin
    End Function

#End Region

#End Region

    Public Function CrearFacturaCompraRealquiler(ByVal dsPropuesta As DataSet, ByVal dtLineasAlbaran As DataTable, ByVal dFechaFacturacion As Date) As DataTable
        ''TODO: Cuando esté todo con Documentos quitar este método
        'Me.BeginTx()

        'For Each drFactura As DataRow In dsPropuesta.Tables("FacturaCompraCabecera").Select
        '    drFactura("FechaParaDeclaracion") = drFactura("FechaFactura")
        'Next

        'CrearFacturaCompra(dsPropuesta)

        'Dim dsAlbaranes As DataSet = PrepararAlbaranRealquiler(dsPropuesta, dtLineasAlbaran, dFechaFacturacion)

        ''AdminData.SetData(dtCabecerasFactura)
        'For Each dt As DataTable In dsAlbaranes.Tables  '//Actualizaciones sobre los albaranes
        '    AdminData.SetData(dt)
        'Next

        'Dim dtCabecerasFactura As DataTable = dsPropuesta.Tables("FacturaCompraCabecera")
        'ActualizarImporte(dtCabecerasFactura)           '//Recalcular importes y generar los Pagos
        'Me.CommitTx()
        'Return dtCabecerasFactura
    End Function

    Private Function PrepararAlbaranRealquiler(ByVal dsPropuesta As DataSet, ByVal dtLineasAlbaran As DataTable, ByVal dFechaFacturacion As Date) As DataSet
        Dim dtACC As DataTable = Nothing
        Dim dtACL As DataTable = Nothing
        Dim dtACD As DataTable = Nothing

        Dim dtLineasFactura As DataTable = dsPropuesta.Tables("FacturaCompraLinea")

        If Not IsNothing(dtLineasAlbaran) AndAlso dtLineasAlbaran.Rows.Count > 0 Then
            '//Construimos los filtros
            Dim objFilterIDAlb As New Filter(FilterUnionOperator.Or)
            Dim objFilterIDLinDevol As New Filter(FilterUnionOperator.Or)
            Dim objFilterIDLinAlb As New Filter(FilterUnionOperator.Or)

            For Each drLineaAlb As DataRow In dtLineasAlbaran.Rows
                objFilterIDAlb.Add(New NumberFilterItem("IDAlbaran", drLineaAlb("IDAlbaran")))
            Next

            For Each drLineaFact As DataRow In dtLineasFactura.Rows
                If Length(drLineaFact("IDLineaDevolucion")) > 0 Then objFilterIDLinDevol.Add(New NumberFilterItem("IDLineaDevolucion", drLineaFact("IDLineaDevolucion")))
                objFilterIDLinAlb.Add(New NumberFilterItem("IDLineaAlbaran", drLineaFact("IDLineaAlbaran")))
            Next

            '//Recuperamos los datos a modificar
            If Not IsNothing(objFilterIDAlb) AndAlso objFilterIDAlb.Count > 0 Then
                dtACC = New BE.DataEngine().Filter("tbAlbaranCompraCabecera", objFilterIDAlb)
            End If
            If Not IsNothing(objFilterIDAlb) AndAlso objFilterIDAlb.Count > 0 Then
                dtACL = New BE.DataEngine().Filter("tbAlbaranCompraLinea", objFilterIDLinAlb)
            End If
            If Not IsNothing(objFilterIDAlb) AndAlso objFilterIDAlb.Count > 0 Then
                dtACD = New BE.DataEngine().Filter("tbAlbaranCompraDevolucion", objFilterIDLinDevol)
            End If


            '//LINEAS DE ALBARAN: - Si es un Línea normal, se marcará como facturada. 
            '//                   - Si es de Realquiler, dependiendo de si tiene Cantidad pendiente 
            '//                     de devolver se marcará como Facturada o Parc. Facturada.
            If Not IsNothing(dtACL) AndAlso dtACL.Rows.Count > 0 Then
                Dim objFilterAlbDev As Filter
                Dim objNegACD As BusinessHelper = BusinessHelper.CreateBusinessObject(AdminData.GetEntityInfo("AlbaranCompraDevolucion"))
                For Each drACL As DataRow In dtACL.Select
                    Select Case CType(drACL("TipoLineaAlbaran"), enumaclTipoLineaAlbaran)
                        Case enumaclTipoLineaAlbaran.aclRealquiler
                            objFilterAlbDev = New Filter
                            objFilterAlbDev.Add(New NumberFilterItem("IDLineaAlbaran", drACL("IDLineaAlbaran")))
                            objFilterAlbDev.Add(New DateFilterItem("FechaDevolucion", FilterOperator.LessThanOrEqual, dFechaFacturacion))
                            Dim dblQDevuelta As Double = 0
                            If Not IsNothing(dtACD) AndAlso dtACD.Rows.Count > 0 Then
                                Dim WhereAlbaranDevol As String = objFilterAlbDev.Compose(New AdoFilterComposer)
                                For Each drACD As DataRow In dtACD.Select(WhereAlbaranDevol)
                                    dblQDevuelta = dblQDevuelta + drACD("QDevuelta")
                                Next
                            End If

                            drACL("QFacturada") = dblQDevuelta
                            If drACL("QServida") <= dblQDevuelta Then
                                drACL("EstadoFactura") = enumaclEstadoFactura.aclFacturado
                            Else
                                drACL("EstadoFactura") = enumaclEstadoFactura.aclParcFacturado
                            End If
                        Case Else
                            drACL("QFacturada") = drACL("QServida")
                            drACL("EstadoFactura") = enumaclEstadoFactura.aclFacturado
                    End Select
                Next
            End If

            '//CABECERAS DE ALBARAN: Se marcarán como Parc.Facturadas mientras no tengan todas las líneas marcadas como facturadas.
            If Not IsNothing(dtACC) AndAlso dtACC.Rows.Count > 0 Then
                Dim objFilterAlb As Filter
                For Each drACC As DataRow In dtACC.Select
                    objFilterAlb = New Filter
                    objFilterAlb.Add(New NumberFilterItem("IDAlbaran", drACC("IDAlbaran")))
                    Dim dvACL As DataView = New DataView(dtACL, objFilterAlb.Compose(New AdoFilterComposer), Nothing, DataViewRowState.CurrentRows)
                    objFilterAlb.Add(New NumberFilterItem("EstadoFactura", enumaclEstadoFactura.aclFacturado))
                    Dim dvACLFact As DataView = New DataView(dtACL, objFilterAlb.Compose(New AdoFilterComposer), Nothing, DataViewRowState.CurrentRows)
                    If Not IsNothing(dvACL) AndAlso Not IsNothing(dvACLFact) Then
                        If dvACL.Count > 0 Then
                            If dvACL.Count = dvACLFact.Count Then
                                drACC("Estado") = enumaccEstado.accFacturado
                            Else
                                '//Si hay alguna línea en estado Parc.Facturado, el albarán pasará  a Prac.Facturado.
                                Dim objFilterAlbPF As New Filter
                                objFilterAlbPF.Add(New NumberFilterItem("IDAlbaran", drACC("IDAlbaran")))
                                objFilterAlbPF.Add(New NumberFilterItem("EstadoFactura", enumaclEstadoFactura.aclParcFacturado))
                                Dim dvACLParcFact As DataView = New DataView(dtACL, objFilterAlbPF.Compose(New AdoFilterComposer), Nothing, DataViewRowState.CurrentRows)
                                If dvACLParcFact.Count > 0 Then
                                    drACC("Estado") = enumaccEstado.accParcFacturado
                                Else
                                    drACC("Estado") = enumaccEstado.accNoFacturado
                                End If

                            End If
                        End If
                    End If
                Next
            End If

            'LINEAS DE DEVOLUCION: Se marcan como facturadas
            If Not IsNothing(dtACD) AndAlso dtACD.Rows.Count > 0 Then
                For Each drACD As DataRow In dtACD.Select
                    drACD("EstadoFactura") = enumaclEstadoFactura.aclFacturado
                Next
            End If
        End If

        Dim ds As New DataSet
        If Not IsNothing(dtACC) Then ds.Tables.Add(dtACC) : dtACC.TableName = "AlbaranCompraCabecera" '//Cabecera
        If Not IsNothing(dtACL) Then ds.Tables.Add(dtACL) : dtACL.TableName = "AlbaranCompraLinea" '//Lineas
        If Not IsNothing(dtACD) Then ds.Tables.Add(dtACD) : dtACD.TableName = "AlbaranCompraDevolucion" '//Devoluciones

        Return ds
    End Function

#End Region

#Region " Copiar Factura "

    <Task()> Public Shared Function CopiarFacturaCompra(ByVal intIDFactura As Integer, ByVal services As ServiceProvider) As DataTable
        If intIDFactura > 0 Then
            Dim FCC As New FacturaCompraCabecera
            Dim dtCabeceraOrigen As DataTable = FCC.SelOnPrimaryKey(intIDFactura)
            Dim f As New Filter
            f.Add(New NumberFilterItem("IDFactura", intIDFactura))
            Dim fcl As New FacturaCompraLinea
            Dim dtLineasOrigen As DataTable = fcl.Filter(f)

            Dim dtFCCD As DataTable = FCC.Filter(New NoRowsFilterItem)
            Dim dtFCLD As DataTable = fcl.Filter(New NoRowsFilterItem)
            Dim AppParamsConta As ParametroContabilidadCompra = services.GetService(GetType(ParametroContabilidadCompra))
            'Copia Cabecera
            For Each drOrigenCabecera As DataRow In dtCabeceraOrigen.Rows
                Dim drDestinoCabecera As DataRow = dtFCCD.NewRow
                For Each dc As DataColumn In dtCabeceraOrigen.Columns
                    If dc.ColumnName <> "IDFactura" And dc.ColumnName <> "NFactura" And _
                       dc.ColumnName <> "FechaContabilizacion" And dc.ColumnName <> "IDObra" And _
                       dc.ColumnName <> "FechaIntrastat" And dc.ColumnName <> "NDeclaracionIVA" And _
                       dc.ColumnName <> "AñoDeclaracionIva" And dc.ColumnName <> "NFacturaIva" And _
                       dc.ColumnName <> "NDeclaracionIntrastat" And dc.ColumnName <> "AñoDeclaracionIntrastat" And _
                       dc.ColumnName <> "IDFacturaVenta" And dc.ColumnName <> "NFacturaAutofactura" Then
                        drDestinoCabecera(dc.ColumnName) = drOrigenCabecera(dc)
                    End If
                Next

                drDestinoCabecera("IDFactura") = AdminData.GetAutoNumeric
                If Length(drDestinoCabecera("IDContador")) > 0 Then
                    Dim StDatos As New Contador.DatosCounterValue
                    StDatos.IDCounter = drDestinoCabecera("IDContador")
                    StDatos.TargetClass = FCC
                    StDatos.TargetField = "NFactura"
                    StDatos.DateField = "FechaFactura"
                    StDatos.IDEjercicio = drDestinoCabecera("IDEjercicio") & String.Empty
                    drDestinoCabecera("NFactura") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, StDatos, services)
                End If

                drDestinoCabecera("SuFactura") = drDestinoCabecera("NFactura")
                drDestinoCabecera("FechaFactura") = Date.Today
                drDestinoCabecera("SuFechaFactura") = Date.Today
                drDestinoCabecera("FechaDeclaracionManual") = False
                drDestinoCabecera("FechaParaDeclaracion") = Date.Today
                ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoFacturacionCompra.FechaParaDeclaracionPorProveedor, New DataRowPropertyAccessor(drDestinoCabecera), services)
                If AppParamsConta.Contabilidad Then
                    drDestinoCabecera("IDEjercicio") = ProcessServer.ExecuteTask(Of Date, String)(AddressOf NegocioGeneral.EjercicioPredeterminado, Today, services)
                End If

                drDestinoCabecera("Estado") = enumfccEstado.fccNoContabilizado

                drDestinoCabecera("IVAManual") = 0
                drDestinoCabecera("VencimientosManuales") = 0
                drDestinoCabecera("IntrastatProcesado") = 0

                'drDestinoCabecera("Enviar347") = False
                drDestinoCabecera("Exportado") = 0
                drDestinoCabecera("Exportar") = 1
                drDestinoCabecera("FacturaPagoPeriodicoSN") = 0
                drDestinoCabecera("NoDescontabilizar") = 0
                drDestinoCabecera("RetencionManual") = 0

                dtFCCD.Rows.Add(drDestinoCabecera)
            Next

            'Copia Líneas
            For Each drOrigenLinea As DataRow In dtLineasOrigen.Rows
                Dim drDestinoLinea As DataRow = dtFCLD.NewRow
                For Each dc As DataColumn In dtLineasOrigen.Columns
                    If dc.ColumnName <> "IDLineaFactura" And dc.ColumnName <> "IDFactura" And _
                       dc.ColumnName <> "IDPedido" And dc.ColumnName <> "IDLineaPedido" And _
                       dc.ColumnName <> "IDAlbaran" And dc.ColumnName <> "IDLineaAlbaran" And _
                       dc.ColumnName <> "IDObra" And dc.ColumnName <> "IDTrabajo" And _
                       dc.ColumnName <> "IDLineaPadre" And dc.ColumnName <> "IDMntoOPrev" And _
                       dc.ColumnName <> "IDLineaOfertaDetalle" And dc.ColumnName <> "IDActivoAImputar" Then
                        drDestinoLinea(dc.ColumnName) = drOrigenLinea(dc)
                    End If
                Next
                drDestinoLinea("IDLineaFactura") = AdminData.GetAutoNumeric
                drDestinoLinea("IDFactura") = dtFCCD.Rows(0)("IDFactura")


                dtFCLD.Rows.Add(drDestinoLinea)
            Next

            '  Me.BeginTx()
            FCC.Update(dtFCCD)
            fcl.Update(dtFCLD)
            'Updated(dtFCCD) 'Recalculo de la factura

            Return dtFCCD
        End If
    End Function

#End Region

#Region " Precio Optimo "
    <Task()> Public Shared Sub PrecioOptimo(ByVal IDFactura As Integer, ByVal services As ServiceProvider)
        Dim DocFra As DocumentoFacturaCompra = ProcessServer.ExecuteTask(Of Integer, DocumentoFacturaCompra)(AddressOf CrearDocumento, IDFactura, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf CalculoPrecioOptimo, DocFra, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularAnaliticaFacturas, DocFra, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularBasesImponibles, DocFra, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularTotales, DocFra, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.CalcularVencimientos, DocFra, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.GrabarDocumento, DocFra, services)
        ProcessServer.ExecuteTask(Of DocumentoFacturaCompra)(AddressOf ProcesoFacturacionCompra.ActualizarAlbaran, DocFra, services)
    End Sub

    <Task()> Public Shared Function CrearDocumento(ByVal IDFactura As Integer, ByVal services As ServiceProvider) As DocumentoFacturaCompra
        Return New DocumentoFacturaCompra(IDFactura)
    End Function

    <Task()> Public Shared Sub CalculoPrecioOptimo(ByVal Doc As DocumentoFacturaCompra, ByVal services As ServiceProvider)
        If Doc Is Nothing OrElse Doc.dtLineas Is Nothing OrElse Doc.dtLineas.Rows.Count = 0 Then Exit Sub

        '//Recogemos los articulos que esten relacionados con esa Albaran.
        Dim dtArticulosFactura As DataTable = New BE.DataEngine().Filter("vNegFacturaCompraLineaArticulos", New StringFilterItem("IDFactura", Doc.HeaderRow("IDFactura")))
        Dim f As New Filter
        For Each drArticuloFactura As DataRow In dtArticulosFactura.Select
            f.Clear()
            f.Add("IDArticulo", drArticuloFactura("IDArticulo"))

            '//Recogemos las lineas de la factura que tengan el articulo de este momento
            Dim QFacturar As Double = Nz(Doc.dtLineas.Compute("SUM(cantidad)", f.Compose(New AdoFilterComposer)), 0)

            Dim dataTarifa As New DataCalculoTarifaCompra
            dataTarifa.IDArticulo = drArticuloFactura("IDArticulo")
            dataTarifa.IDProveedor = Doc.IDProveedor
            dataTarifa.Cantidad = QFacturar
            dataTarifa.Fecha = Doc.Fecha
            dataTarifa.IDMoneda = Doc.IDMoneda
            ProcessServer.ExecuteTask(Of DataCalculoTarifaCompra)(AddressOf ProcesoCompra.TarifaCompra, dataTarifa, services)
            If Not dataTarifa.DatosTarifa Is Nothing Then
                Dim FCL As New FacturaCompraLinea
                Dim context As New BusinessData(Doc.HeaderRow)
                Dim WhereArticulo As String = f.Compose(New AdoFilterComposer)
                For Each drFacturaLineaArticulo As DataRow In Doc.dtLineas.Select(WhereArticulo)
                    FCL.ApplyBusinessRule("Precio", dataTarifa.DatosTarifa.Precio, drFacturaLineaArticulo, context)
                    FCL.ApplyBusinessRule("Dto1", dataTarifa.DatosTarifa.Dto1, drFacturaLineaArticulo, context)
                    FCL.ApplyBusinessRule("Dto2", dataTarifa.DatosTarifa.Dto2, drFacturaLineaArticulo, context)
                    FCL.ApplyBusinessRule("Dto3", dataTarifa.DatosTarifa.Dto3, drFacturaLineaArticulo, context)
                    FCL.ApplyBusinessRule("UDValoracion", dataTarifa.DatosTarifa.UDValoracion, drFacturaLineaArticulo, context)
                    If Length(dataTarifa.DatosTarifa.SeguimientoTarifa) > 0 Then
                        drFacturaLineaArticulo("SeguimientoTarifa") = dataTarifa.DatosTarifa.SeguimientoTarifa
                    End If
                Next
            End If
            QFacturar = 0
        Next
    End Sub
    
#End Region

#Region " Facturas Leasing "

    Private Function NuevaLineaFacturaConcesion(ByVal lngIDFactura As Integer, ByVal strNFactura As String, _
                                              ByRef dvData As DataView, ByVal dtmFechaFactura As Date, _
                                              ByVal strIDProveedor As String, ByVal Lineas As DataTable, _
                                              ByVal strIDCentroGestion As String, ByVal strIDMoneda As String, _
                                              ByVal services As ServiceProvider) As DataTable

        If Not dvData Is Nothing Then
            Dim fcl As New FacturaCompraLinea
            Dim strIN As String
            Dim ClsArtProv As New ArticuloProveedor

            For Each dr As DataRowView In dvData
                If Len(strIN) Then strIN = strIN & ","
                strIN = strIN & dr.Row("IdPagoPer")
            Next

            Dim objFilter As New Filter
            objFilter.Add(New InListFilterItem("IDPagoPer", strIN))
            Dim dtPago As DataTable = New BE.DataEngine().Filter("vCtlCIPagoPerGeneraFactura", objFilter)
            If Not dtPago Is Nothing AndAlso dtPago.Rows.Count > 0 Then
                Dim AppParamsConta As ParametroContabilidadCompra = services.GetService(GetType(ParametroContabilidadCompra))
                Dim objFilterPago As New Filter
                For Each pago As DataRow In dtPago.Rows
                    objFilterPago.Clear()
                    objFilterPago.Add(New NumberFilterItem("IDPago", pago("IDPago")))
                    dvData.RowFilter = objFilterPago.Compose(New AdoFilterComposer)
                    If dvData.Count > 0 Then
                        Dim linea As DataRow = fcl.AddNewForm.Rows(0)
                        linea("IDFactura") = lngIDFactura
                        linea("NFactura") = strNFactura
                        linea("IDCentroGestion") = strIDCentroGestion
                        linea("Cantidad") = 1
                        linea("Factor") = 1
                        linea("QInterna") = 1
                        linea("UdValoracion") = 1
                        linea("Precio") = pago("ImpNetoNominal")
                        linea("Importe") = pago("ImpNetoNominal")
                        If AppParamsConta.Contabilidad Then linea("CContable") = pago("CCNominal")
                        linea("Dto1") = 0
                        linea("Dto2") = 0
                        linea("Dto3") = 0
                        linea("TipoGastoObra") = enumfclTipoGastoObra.enumfclMaterial
                        linea("IDUDInterna") = pago("IDUDInterna")

                        objFilter.Clear()
                        objFilter.Add(New StringFilterItem("IDArticulo", pago("IDArticulo")))
                        objFilter.Add(New BooleanFilterItem("Compra", True))
                        objFilter.Add(New BooleanFilterItem("Activo", True))
                        Dim DtArt As DataTable = New BE.DataEngine().Filter("vNegCaractArticulo", objFilter)
                        If Not DtArt Is Nothing AndAlso DtArt.Rows.Count > 0 Then
                            linea("IDArticulo") = pago("IDArticulo")
                            Dim DtArtProv As DataTable = ClsArtProv.SelOnPrimaryKey(pago("IDProveedor"), pago("IDArticulo"))
                            If Not DtArtProv Is Nothing AndAlso DtArtProv.Rows.Count > 0 Then
                                linea("RefProveedor") = DtArtProv.Rows(0)("RefProveedor") & String.Empty
                                linea("DescArticulo") = DtArtProv.Rows(0)("DescRefProveedor") & String.Empty
                                linea("IDUDMedida") = DtArtProv.Rows(0)("IdUdCompra") & String.Empty
                                linea("UdValoracion") = DtArtProv.Rows(0)("UdValoracion")
                            End If
                            If Length(linea("DescArticulo")) = 0 Then linea("DescArticulo") = DtArt.Rows(0)("DescArticulo")
                            'TODO: REvisar ObtenerIVA
                            'linea("IDTipoIva") = ClsCompra.ObtenerIVA(pago("IdProveedor"), pago("IDArticulo"))
                            If linea("UdValoracion") = 0 Then linea("UdValoracion") = IIf(DtArt.Rows(0)("UdValoracion") > 0, DtArt.Rows(0)("IDArticulo"), 1)
                            If Length(linea("IDUDMedida")) = 0 Then linea("IDUDMedida") = DtArt.Rows(0)("IDUDInterna")
                        End If
                        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                        Dim MonInfo As MonedaInfo = Monedas.GetMoneda(strIDMoneda)
                        Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(linea), MonInfo.ID, MonInfo.CambioA, MonInfo.CambioB)
                        ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf MantenimientoValoresAyB, ValAyB, services)
                        Lineas.Rows.Add(linea.ItemArray)
                    End If
                Next
            End If
            dvData.RowFilter = String.Empty
            Return Lineas
        End If
    End Function

    'Private Function NuevaLineaFacturaLeasing(ByVal lngIDFactura As Integer, ByVal strNFactura As String, _
    '                                         ByRef dvData As DataView, ByVal dtmFechaFactura As Date, ByVal strIDProveedor As String, _
    '                                         ByVal Lineas As DataTable, ByVal strIDCentroGestion As String, _
    '                                         ByVal strIDMoneda As String, ByRef DtPagoLeasing As DataTable, _
    '                                         ByVal services As ServiceProvider) As DataTable

    '    If Not dvData Is Nothing Then
    '        Dim fcl As New FacturaCompraLinea
    '        Dim ClsArtProv As New ArticuloProveedor
    '        Dim ClsProv As New Proveedor
    '        Dim strIN As String
    '        For Each dr As DataRowView In dvData
    '            If Len(strIN) Then strIN = strIN & ","
    '            strIN = strIN & dr.Row("IdPago")
    '        Next

    '        Dim objFilter As New Filter
    '        objFilter.Add(New InListFilterItem("IDPago", strIN))
    '        Dim DtPago As DataTable = New BE.DataEngine().Filter("vCtlCIPagoContGeneraFactura", objFilter)
    '        If Not DtPago Is Nothing AndAlso DtPago.Rows.Count > 0 Then
    '            Dim AppParamsConta As ParametroContabilidadCompra = services.GetService(GetType(ParametroContabilidadCompra))
    '            Dim objFilterPago As New Filter
    '            For Each pago As DataRow In DtPago.Rows
    '                objFilterPago.Clear()
    '                objFilterPago.Add(New NumberFilterItem("IDPago", pago("IDPago")))
    '                dvData.RowFilter = objFilterPago.Compose(New AdoFilterComposer)
    '                If dvData.Count > 0 Then
    '                    Dim linea As DataRow = fcl.AddNewForm.Rows(0)
    '                    linea("IDFactura") = lngIDFactura

    '                    Dim DrNew As DataRow = DtPagoLeasing.NewRow()
    '                    DrNew("IDFactura") = lngIDFactura
    '                    DrNew("IDPago") = pago("IDPago")
    '                    DrNew("NFactura") = strNFactura
    '                    DtPagoLeasing.Rows.Add(DrNew)


    '                    linea("NFactura") = strNFactura
    '                    linea("IDCentroGestion") = strIDCentroGestion
    '                    linea("Cantidad") = 1
    '                    linea("Factor") = 1
    '                    linea("QInterna") = 1
    '                    linea("UdValoracion") = 1
    '                    linea("Precio") = pago("Importe")
    '                    linea("Importe") = pago("Importe")
    '                    If AppParamsConta.Contabilidad Then
    '                        If Length(pago("IDProveedor")) > 0 Then
    '                            Dim DtProv As DataTable = ClsProv.SelOnPrimaryKey(pago("IDProveedor"))
    '                            If Not DtProv Is Nothing AndAlso DtProv.Rows.Count > 0 Then
    '                                linea("CContable") = DtProv.Rows(0)("CCInMovilizadoCortoPlazo")
    '                            End If
    '                        End If
    '                    End If

    '                    linea("Dto1") = 0
    '                    linea("Dto2") = 0
    '                    linea("Dto3") = 0
    '                    linea("Dto") = 0
    '                    linea("DtoProntoPago") = 0
    '                    linea("TipoGastoObra") = enumfclTipoGastoObra.enumfclMaterial
    '                    linea("IDUDInterna") = pago("IDUDInterna")

    '                    objFilter.Clear()
    '                    objFilter.Add(New StringFilterItem("IDArticulo", pago("IDArticulo")))
    '                    objFilter.Add(New BooleanFilterItem("Compra", True))
    '                    objFilter.Add(New BooleanFilterItem("Activo", True))

    '                    Dim DtArt As DataTable = New BE.DataEngine().Filter("vNegCaractArticulo", objFilter)
    '                    If Not DtArt Is Nothing AndAlso DtArt.Rows.Count > 0 Then
    '                        linea("IDArticulo") = DtArt.Rows(0)("IDArticulo")
    '                        Dim DtRef As DataTable = ClsArtProv.SelOnPrimaryKey(pago("IDProveedor"), pago("IDArticulo"))
    '                        If Not DtRef Is Nothing AndAlso DtRef.Rows.Count > 0 Then
    '                            linea("RefProveedor") = DtRef.Rows(0)("RefProveedor") & String.Empty
    '                            linea("DescArticulo") = DtRef.Rows(0)("DescRefProveedor") & String.Empty
    '                            linea("IDUDMedida") = DtRef.Rows(0)("IdUdCompra") & String.Empty
    '                            linea("UdValoracion") = DtRef.Rows(0)("UdValoracion")
    '                        End If
    '                        If Length(linea("DescArticulo")) = 0 Then linea("DescArticulo") = DtArt.Rows(0)("DescArticulo")
    '                        'TODO: ObtenerIVA, poner esto correctamente
    '                        'linea("IDTipoIva") = ClsCompra.ObtenerIVA(pago("IDProveedor"), pago("IDArticulo"))
    '                        If linea("UdValoracion") = 0 Then linea("UdValoracion") = IIf(DtArt.Rows(0)("UdValoracion") > 0, DtArt.Rows(0)("UdValoracion"), 1)
    '                        If Length(linea("IDUDMedida")) = 0 Then linea("IDUDMedida") = DtArt.Rows(0)("IDUDInterna")
    '                    End If
    '                    Dim Monedas As MonedaCache = services.GetService(GetType(MonedaCache))
    '                    Dim MonInfo As MonedaInfo = Monedas.GetMoneda(strIDMoneda)
    '                    Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(linea), MonInfo.ID, MonInfo.CambioA, MonInfo.CambioB)
    '                    ProcessServer.ExecuteTask(Of ValoresAyB)(AddressOf MantenimientoValoresAyB, ValAyB, services)
    '                    Lineas.Rows.Add(linea.ItemArray)
    '                End If
    '            Next
    '        End If
    '        dvData.RowFilter = String.Empty
    '        Return Lineas
    '    End If
    'End Function

    'Public Function CrearFacturaCompra(ByVal ds As DataSet, Optional ByVal strTipo As String = vbNullString) As DataTable
    '    Dim dtCabecerasFactura As DataTable = ds.Tables("FacturaCompraCabecera")
    '    Dim dtLineasFactura As DataTable = ds.Tables("FacturaCompraLinea")

    '    Me.BeginTx()
    '    Me.Update(ds.Tables(0)) ''La cabecera la se actualiza con el Update para actualizar el contador.
    '    Dim i As Integer
    '    For i = 1 To ds.Tables.Count - 1
    '        AdminData.SetData(ds.Tables(i))
    '    Next
    '    'Dim Monedas As New MonedaCache
    '    'Dim AppParams As New ParametroCache
    '    'Dim TiposIVA As New EntityInfoCache(Of TipoIvaInfo)
    '    ActualizarImporte(dtCabecerasFactura, strTipo, New ServiceProvider) '///Recalcular importes albaran
    '    Select Case strTipo
    '        Case "Leasing"
    '            '    ActualizarPago(dtCabecerasFactura)
    '        Case "Concesion"
    '            ActualizarPagosPer(dtLineasFactura)
    '        Case Else
    '            ActualizarAlbaran(dtLineasFactura)
    '    End Select
    '    Dim FCL As New FacturaCompraLinea
    '    FCL.CrearConceptosObraControl(dtLineasFactura)
    '    Return dtCabecerasFactura
    'End Function

    'Private Function ActualizarPago(ByVal dtPagoLeasing As DataTable) As DataTable
    '    Dim strIN As String
    '    Dim strINPago As String
    '    Dim dtPago As DataTable
    '    Dim P As New Pago
    '    Dim dtPagoAct As DataTable = P.AddNew

    '    For Each dr As DataRow In dtPagoLeasing.Rows
    '        If InStr(strIN, CStr(dr("IDPago")), CompareMethod.Text) = 0 Then
    '            If Len(strIN) > 0 Then strIN = strIN & ","
    '            strIN = strIN & dr("IDPago")

    '            dtPago = P.SelOnPrimaryKey(dr("IDPago"))
    '            If Not dtPago Is Nothing AndAlso dtPago.Rows.Count > 0 Then
    '                dtPago.Rows(0)("IDFactura") = dr("IDFactura")
    '                dtPago.Rows(0)("NFactura") = dr("NFactura")

    '                dtPagoAct.ImportRow(dtPago.Rows(0))
    '            End If

    '        End If
    '    Next
    '    Return dtPagoAct
    'End Function

    'Private Function ActualizarPagosPer(ByVal DtLineas As DataTable) As DataTable
    '    Dim StrIN, StrWhere As String
    '    Dim LngNRegAlbaran As Integer

    '    If Not DtLineas Is Nothing AndAlso DtLineas.Rows.Count > 0 Then
    '        For Each Dr As DataRow In DtLineas.Select
    '            If Length(StrIN) Then StrIN &= ","
    '            StrIN &= Dr("IDPagoPer")
    '        Next
    '        StrWhere = "ID IN (" & StrIN & ")"
    '        Dim DtPagoPer As DataTable = AdminData.GetData("tbPagoPeriodico", , StrWhere)
    '        If Not DtPagoPer Is Nothing AndAlso DtPagoPer.Rows.Count > 0 Then
    '            For Each DrPago As DataRow In DtPagoPer.Select
    '                For Each DrLinea As DataRow In DtLineas.Select
    '                    If DrPago("ID") = DrLinea("IDPagoPer") Then
    '                        DrPago("IDFactura") = DrLinea("IDFactura")
    '                        Exit For
    '                    End If
    '                Next
    '            Next
    '            Return DtPagoPer
    '        End If
    '    End If
    'End Function


    'Public Sub BorrarFacturasLeasing(ByVal StrIDProcess As String, _
    '                                 Optional ByVal LngIDFactura As Integer = -1, _
    '                                 Optional ByVal DblDescontabilizaLeasing As Boolean = True, _
    '                                 Optional ByVal LngNuevaSituacion As Integer = 0)
    '    Dim ClsFCL As New FacturaCompraLinea
    '    Dim ClsFCBI As New FacturaCompraBaseImponible
    '    Dim ClsPagoPer As New PagoPeriodico
    '    Dim Dt, DtFact, DtPago As DataTable
    '    Dim BlmDescontabilizar As Boolean
    '    Dim StrWhere As String

    '    If LngIDFactura <= 0 Then
    '        'StrWhere = "IdPrograma='" & StrIDPrograma & "' AND IdUsuario='" & StrMaquinaUsuario & "'"
    '        StrWhere = "IdProcess='" & StrIDProcess & "'"
    '        If DblDescontabilizaLeasing Then
    '            Dt = New BE.DataEngine().Filter("vCtlCIPagoCont", "IDFactura", StrWhere, , False)
    '        Else
    '            Dt = New BE.DataEngine().Filter("vNegSimulacionContablePago", "IDFactura", StrWhere, , False)
    '        End If
    '        If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
    '            StrWhere = String.Empty
    '            For Each Dr As DataRow In Dt.Select
    '                If Length(StrWhere) > 0 Then StrWhere &= " OR "
    '                StrWhere &= "IDFactura= " & Dr("IDFactura")
    '            Next
    '        Else : StrWhere = "1=2"
    '        End If
    '    Else
    '        StrWhere = "IDFactura= " & LngIDFactura
    '        DtPago = ClsPagoPer.Filter(, StrWhere)
    '        If Not DtPago Is Nothing AndAlso DtPago.Rows.Count > 0 Then
    '            ClsPagoPer.Delete(DtPago)
    '        End If
    '    End If
    '    Dt = Me.Filter("*", StrWhere)
    '    If Not Dt Is Nothing AndAlso Dt.Rows.Count > 0 Then
    '        For Each Dr As DataRow In Dt.Select
    '            Dr("estado") = enumfccEstado.fccNoContabilizado
    '        Next
    '        BusinessHelper.UpdateTable(Dt)
    '        DtFact = ClsFCBI.Filter("IDBaseImponible", StrWhere)
    '        If Not DtFact Is Nothing AndAlso DtFact.Rows.Count > 0 Then
    '            ClsFCBI.Delete(DtFact)
    '        End If
    '        DtFact = ClsFCL.Filter("IDLineaFactura", StrWhere)
    '        If Not DtFact Is Nothing AndAlso DtFact.Rows.Count > 0 Then
    '            ClsFCL.Delete(DtFact)
    '        End If
    '        Dim pa As New Pago
    '        DtPago = pa.Filter("IDFactura,NFactura,IDPago,Contabilizado,IDEjercicio,FechaContabilizacion,Situacion", StrWhere)
    '        If Not DtPago Is Nothing AndAlso DtPago.Rows.Count > 0 Then
    '            If Length(StrIDProcess) > 0 Then BlmDescontabilizar = True
    '            For Each Dr As DataRow In DtPago.Select
    '                If BlmDescontabilizar Then DescontabilizarPagosLeasing(StrIDProcess, Dr("IDFactura"), LngNuevaSituacion)
    '                Dr("Contabilizado") = enumPagoContabilizado.PagoNoContabilizado
    '                Dr("IDEjercicio") = DBNull.Value
    '                Dr("FechaContabilizacion") = DBNull.Value
    '                Dr("IDFactura") = DBNull.Value
    '                Dr("NFactura") = DBNull.Value
    '                If LngNuevaSituacion >= 0 Then Dr("Situacion") = LngNuevaSituacion
    '            Next
    '            BusinessHelper.UpdateTable(DtPago)
    '        End If
    '        Me.Delete(Dt)
    '    End If
    'End Sub

#End Region

#Region " Consultas interactivas (Estadísticas) "

    <Task()> Public Shared Function ObtenerEstadisticaFCTipos(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Return New DataEngine().Filter("tbEstadisticaCompraAnual", String.Empty, String.Empty)
    End Function

#Region " ObtenerEstadisticaCantidadesMeses "

    <Serializable()> _
    Public Class DataEstadisticaCantidadesMeses
        Public CamposSelect As String
        Public CampoATotalizar As String
        Public CamposOrden As String
        Public GroupBy As String

        Public IDTipo As String
        Public IDFamilia As String
        Public IDArticulo As String
        Public IDProveedor As String
        Public IDMercado As String
        Public Provincia As String
        Public IDZona As String
        Public IDPais As String
        Public CEE As enumBoolean
        Public Extranjero As enumBoolean
        Public Año, Año2 As Integer
        Public EmpresaGrupo As enumBoolean

        Public Sub New(ByVal CamposSelect As String, ByVal CampoATotalizar As String, ByVal IDTipo As String, ByVal IDFamilia As String, _
                       ByVal IDArticulo As String, ByVal IDProveedor As String, ByVal IDMercado As String, ByVal Provincia As String, ByVal IDZona As String, _
                       ByVal IDPais As String, ByVal CEE As enumBoolean, ByVal Extranjero As enumBoolean, _
                       ByVal Año As Integer, ByVal EmpresaGrupo As enumBoolean, ByVal GroupBy As String, ByVal CamposOrden As String)

            Me.CamposSelect = CamposSelect
            Me.CampoATotalizar = CampoATotalizar
            Me.CamposOrden = CamposOrden
            Me.GroupBy = GroupBy
            Me.IDTipo = IDTipo
            Me.IDFamilia = IDFamilia
            Me.IDArticulo = IDArticulo
            Me.IDProveedor = IDProveedor
            Me.IDMercado = IDMercado
            Me.Provincia = Provincia
            Me.IDZona = IDZona
            Me.IDPais = IDPais
            Me.CEE = CEE
            Me.Extranjero = Extranjero
            Me.Año = Año
            Me.EmpresaGrupo = EmpresaGrupo
        End Sub
        Public Sub New(ByVal CamposSelect As String, ByVal CampoATotalizar As String, ByVal IDTipo As String, ByVal IDFamilia As String, _
                       ByVal IDArticulo As String, ByVal IDProveedor As String, ByVal IDMercado As String, ByVal Provincia As String, ByVal IDZona As String, _
                       ByVal IDPais As String, ByVal CEE As enumBoolean, ByVal Extranjero As enumBoolean, _
                       ByVal Año As Integer, ByVal Año2 As Integer, ByVal EmpresaGrupo As enumBoolean, ByVal GroupBy As String, ByVal CamposOrden As String)

            Me.CamposSelect = CamposSelect
            Me.CampoATotalizar = CampoATotalizar
            Me.CamposOrden = CamposOrden
            Me.GroupBy = GroupBy
            Me.IDTipo = IDTipo
            Me.IDFamilia = IDFamilia
            Me.IDArticulo = IDArticulo
            Me.IDProveedor = IDProveedor
            Me.IDMercado = IDMercado
            Me.Provincia = Provincia
            Me.IDZona = IDZona
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
            "SUM(CASE MONTH([FechaFactura]) WHEN 12 THEN {1} ELSE 0 END)  AS SDiciembre," & _
            "SUM({1}) As STotalLinea", data.CamposSelect, data.CampoATotalizar))

        selectSQL.Append(" FROM tbMaestroMercado RIGHT OUTER JOIN" & _
            " tbFacturaCompraLinea INNER JOIN" & _
            " tbFacturaCompraCabecera ON tbFacturaCompraLinea.IDFactura = tbFacturaCompraCabecera.IDFactura INNER JOIN" & _
            " tbMaestroProveedor ON tbFacturaCompraCabecera.IDProveedor = tbMaestroProveedor.IDProveedor INNER JOIN" & _
            " tbMaestroArticulo ON tbFacturaCompraLinea.IDArticulo = tbMaestroArticulo.IDArticulo INNER JOIN" & _
            " tbMaestroPais ON tbMaestroProveedor.IDPais = tbMaestroPais.IDPais LEFT OUTER JOIN" & _
            " tbMaestroSubFamilia ON tbMaestroArticulo.IDSubFamilia = tbMaestroSubFamilia.IDSubFamilia and" & _
            " tbMaestroArticulo.IDFamilia = tbMaestroSubFamilia.IDFamilia AND" & _
            " tbMaestroArticulo.IDTipo = tbMaestroSubFamilia.IDTipo LEFT OUTER JOIN" & _
            " tbMaestroFamilia ON tbMaestroArticulo.IDFamilia = tbMaestroFamilia.IDFamilia AND" & _
            " tbMaestroArticulo.IDTipo = tbMaestroFamilia.IDTipo LEFT OUTER JOIN" & _
            " tbMaestroZona ON tbMaestroProveedor.IDZona = tbMaestroZona.IDZona ON" & _
            " tbMaestroMercado.IDMercado = tbMaestroProveedor.IDMercado")

        Dim whereSQL As New Text.StringBuilder
        If data.Año.ToString.Length > 0 Then
            whereSQL.Append("YEAR(tbFacturaCompraCabecera.FechaFactura) = " & data.Año & " AND ")
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
        If data.IDProveedor.Length > 0 Then
            whereSQL.Append("tbMaestroProveedor.IDProveedor = '" & data.IDProveedor & "' AND ")
        End If
        If data.Provincia.Length > 0 Then
            whereSQL.Append("tbMaestroProveedor.Provincia = '" & data.Provincia & "' AND ")
        End If
        If data.IDZona.Length > 0 Then
            whereSQL.Append("tbMaestroProveedor.IDZona = '" & data.IDZona & "' AND ")
        End If
        If data.IDMercado.Length > 0 Then
            whereSQL.Append("tbMaestroProveedor.IDMercado = '" & data.IDMercado & "' AND ")
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
                whereSQL.Append("tbMaestroProveedor.EmpresaGrupo = 1 AND ")
            Case enumBoolean.No
                whereSQL.Append("tbMaestroProveedor.EmpresaGrupo = 0 AND ")
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

    Public Function EsOrigenCompraConsultaDiferencias(ByVal guidFormulario As System.Guid) As Boolean
        Dim Dt As DataTable = New BE.DataEngine().Filter("xProgram", New GuidFilterItem("IDPrograma", FilterOperator.Equal, guidFormulario), , , , True)
        If Dt.Rows(0)("Alias") = "IVACCONDIF" Then
            Return True
        Else : Return False
        End If
    End Function

    Public Function EsOrigenCompraCompraIntracom(ByVal guidFormulario As System.Guid) As Boolean
        Dim Dt As DataTable = New BE.DataEngine().Filter("xProgram", New GuidFilterItem("IDPrograma", FilterOperator.Equal, guidFormulario), , , , True)
        If Dt.Rows(0)("Alias") = "CIIVACINT" Then
            Return True
        Else : Return False
        End If
    End Function

    Public Function ObtenerRegistrosMarcadosIVASCompras(ByVal guidProceso As System.Guid, ByVal esCompra As Boolean) As DataTable
        If esCompra Then
            Return New BE.DataEngine().Filter("vCtlCIIvaCompra", New GuidFilterItem("IDProcess", FilterOperator.Equal, guidProceso), , "NFacturaIva")
        Else
            Return New BE.DataEngine().Filter("vCtlCIIvaVenta", New GuidFilterItem("IDProcess", FilterOperator.Equal, guidProceso), , "NFactura")
        End If
    End Function

    Public Function ObtenerDatosDeclaracionCompraIntracom(ByVal esCompra As Boolean) As DataTable
        If esCompra Then
            Return AdminData.Execute("SELECT NDeclaracionIVA, AñoDeclaracionIVA, NFacturaIVA " & _
                "FROM vCtlCIIVACompra " & _
                "WHERE NDeclaracionIVA IS NOT NULL OR AñoDeclaracionIVA IS NOT NULL ORDER BY NFacturaIva", ExecuteCommand.ExecuteReader)
        Else
            Return AdminData.Execute("SELECT NDeclaracionIVA, AñoDeclaracionIVA " & _
                "FROM vCtlCIIVAVenta " & _
                "WHERE NDeclaracionIVA IS NOT NULL OR AñoDeclaracionIVA IS NOT NULL ORDER BY NFactura", ExecuteCommand.ExecuteReader)
        End If
    End Function

    Public Function ObtenerDatosDeclaracionIVACompra() As DataTable
        Return AdminData.Execute("SELECT AñoDeclaracionIVA, NDeclaracionIVA " & _
                                "FROM vCtlCIIVACompra " & _
                                "WHERE NDeclaracionIVA IS NOT NULL OR AñoDeclaracionIVA IS NOT NULL " & _
                                "GROUP BY AñoDeclaracionIVA, NDeclaracionIVA " & _
                                "ORDER BY AñoDeclaracionIVA, NDeclaracionIVA", ExecuteCommand.ExecuteReader)
    End Function

#End Region

    <Serializable()> _
    Public Class DataCambiarEstadoFactura
        Public Filtro As Filter
        Public Estado As Integer

        Public Sub New(ByVal Filtro As Filter, ByVal Estado As Integer)
            Me.Filtro = Filtro
            Me.Estado = Estado
        End Sub
    End Class

    <Task()> Public Shared Function CambiarEstadoContabFacturas(ByVal data As DataCambiarEstadoFactura, ByVal services As ServiceProvider) As Boolean
        If Not data.Filtro Is Nothing AndAlso data.Filtro.Count > 0 Then
            Dim StrSql As String = "UPDATE tbFacturaCompraCabecera "
            StrSql &= "SET Estado = " & data.Estado & " "
            StrSql &= "WHERE " & AdminData.ComposeFilter(data.Filtro) & " "
            AdminData.Execute(StrSql)
            Return True
        End If
    End Function

    <Task()> Public Shared Function GetParamsFacturaCompra(ByVal data As Object, ByVal services As ServiceProvider) As DataParamFacturaCompra
        Dim AppParamsConta As ParametroContabilidad = services.GetService(Of ParametroContabilidad)()
        Dim AppParamsGeneral As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim AppParamsStock As ParametroStocks = services.GetService(Of ParametroStocks)()

        Dim datParams As New DataParamFacturaCompra
        datParams.GestionAnalitica = AppParamsConta.Analitica.AplicarAnalitica
        datParams.ExpertisSAAS = AppParamsGeneral.SAAS
        datParams.Contabilidad = AppParamsConta.Contabilidad
        datParams.GestionDobleUnidad = AppParamsStock.GestionDobleUnidad

        datParams.MonInfoA = Monedas.MonedaA
        datParams.MonInfoB = Monedas.MonedaB

        Return datParams
    End Function

#Region " Actualización masiva Facturas Compra "

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

