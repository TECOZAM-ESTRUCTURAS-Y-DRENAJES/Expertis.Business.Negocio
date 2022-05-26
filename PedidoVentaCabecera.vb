Imports Solmicro.Expertis.Business.Negocio.NegocioGeneral

Public Class _PedidoVentaCabecera
    Public Const IDPedido As String = "IDPedido"
    Public Const NPedido As String = "NPedido"
    Public Const FechaPedido As String = "FechaPedido"
    Public Const IDContador As String = "IDContador"
    Public Const IDCliente As String = "IDCliente"
    Public Const IDCentroGestion As String = "IDCentroGestion"
    Public Const IDFormaEnvio As String = "IDFormaEnvio"
    Public Const IDEstadoPedido As String = "IDEstadoPedido"
    Public Const IDAlmacen As String = "IDAlmacen"
    Public Const IDCondicionEnvio As String = "IDCondicionEnvio"
    Public Const IDMoneda As String = "IDMoneda"
    Public Const IDDireccionEnvio As String = "IDDireccionEnvio"
    Public Const IDFormaPago As String = "IDFormaPago"
    Public Const IDCondicionPago As String = "IDCondicionPago"
    Public Const PedidoCliente As String = "PedidoCliente"
    Public Const FechaEntrega As String = "FechaEntrega"
    Public Const GastosEnvio As String = "GastosEnvio"
    Public Const ImpPedido As String = "ImpPedido"
    Public Const ImpPedidoA As String = "ImpPedidoA"
    Public Const ImpPedidoB As String = "ImpPedidoB"
    Public Const Estado As String = "Estado"
    Public Const CambioA As String = "CambioA"
    Public Const CambioB As String = "CambioB"
    Public Const Texto As String = "Texto"
    Public Const Prioridad As String = "Prioridad"
    Public Const IdCentroSolicitante As String = "IdCentroSolicitante"
    Public Const PedidoInterno As String = "PedidoInterno"
    Public Const ImpTotal As String = "ImpTotal"
    Public Const ImpTotalA As String = "ImpTotalA"
    Public Const ImpTotalB As String = "ImpTotalB"
    Public Const ImpIva As String = "ImpIva"
    Public Const ImpIvaA As String = "ImpIvaA"
    Public Const ImpIvaB As String = "ImpIvaB"
    Public Const ImpRE As String = "ImpRE"
    Public Const ImpREA As String = "ImpREA"
    Public Const ImpREB As String = "ImpREB"
    Public Const DtoPedido As String = "DtoPedido"
    Public Const ImpDto As String = "ImpDto"
    Public Const ImpDtoA As String = "ImpDtoA"
    Public Const ImpDtoB As String = "ImpDtoB"
    Public Const EDI As String = "EDI"
    Public Const FechaCreacionAudi As String = "FechaCreacionAudi"
    Public Const FechaModificacionAudi As String = "FechaModificacionAudi"
    Public Const UsuarioAudi As String = "UsuarioAudi"
    Public Const SyncDB As String = "SyncDB"
    Public Const PedidoClienteDestino As String = "PedidoClienteDestino"
    Public Const FechaAviso As String = "FechaAviso"
    Public Const FechaPreparacion As String = "FechaPreparacion"
    Public Const IDEjercicio As String = "IDEjercicio"
    Public Const ImpDpp As String = "ImpDpp"
    Public Const ImpDppA As String = "ImpDppA"
    Public Const ImpDppB As String = "ImpDppB"
    Public Const DtoProntoPago As String = "DtoProntoPago"
    Public Const RecFinan As String = "RecFinan"
    Public Const ImpRecFinan As String = "ImpRecFinan"
    Public Const ImpRecFinanA As String = "ImpRecFinanA"
    Public Const ImpRecFinanB As String = "ImpRecFinanB"
    Public Const BaseImponible As String = "BaseImponible"
    Public Const BaseImponibleA As String = "BaseImponibleA"
    Public Const BaseImponibleB As String = "BaseImponibleB"
    Public Const IDDireccionFra As String = "IDDireccionFra"
    Public Const IDClienteBanco As String = "IDClienteBanco"
End Class

<Serializable()> _
Public Class PedidoVentaInfo
    Public Fields As New BusinessData
    Public Lineas(-1) As PedidoVentaLineaInfo
End Class

<Serializable()> _
Public Class PedidoVentaLineaInfo
    Public Fields As New BusinessData
    Public Analitica As DataTable
    Public Representantes As DataTable
End Class

<Serializable()> _
Public Class PedidoVentaUpdateData
    Public Pedidos As DataTable
    Public ProgramasError(-1) As String
    Public MensajeError As String
End Class

Public Class PedidoVentaCabeceraInfo
    Inherits ClassEntityInfo

    Public IDPedido As Integer
    Public NPedido As String
    Public IDCliente As String
    Public PedidoCliente As String
    Public IDCondicionPago As String
    Public IDFormaPago As String
    Public DtoPedido As Double
    Public IDMoneda As String
    Public IDDireccionFra As Integer
    Public IDClienteBanco As Integer

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dt As DataTable = New PedidoVentaCabecera().SelOnPrimaryKey(PrimaryKey(0))
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            Me.Fill(dt.Rows(0))
        End If
    End Sub

End Class


Public Class PedidoVentaCabecera
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbPedidoVentaCabecera"

#Region " RegisterAddNewTasks "

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoComunes.AsignarIdentificadorPedido, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.AsignarCentroGestion, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarContador, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarNumeroPedidoProvisional, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.AsignarAlmacen, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoComunes.AsignarFechaPedido, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarResponsable, data, services)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoComunes.AsignarEjercicioContablePedido, New DataRowPropertyAccessor(data), services)
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim CE As New CentroEntidad
        CE.CentroGestion = data("IDCentroGestion") & String.Empty
        CE.ContadorEntidad = CentroGestion.ContadorEntidad.PedidoVenta
        data("IDContador") = ProcessServer.ExecuteTask(Of CentroEntidad, String)(AddressOf New CentroGestion().GetContadorPredeterminado, CE, services)
    End Sub

    <Task()> Public Shared Sub AsignarNumeroPedidoProvisional(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDContador")) > 0 Then
            Dim dtContadores As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDt, GetType(PedidoVentaCabecera).Name, services)
            Dim adr As DataRow() = dtContadores.Select("IDContador = " & Quoted(data("IDContador")))
            If Not IsNothing(adr) AndAlso adr.Length > 0 Then
                data("NPedido") = adr(0)("ValorProvisional")
            Else
                'Si no está bien configurado el Contador de Pedidos de Venta en el Centro de Gestión,
                'cogemos el Contador por defecto de la entidad Pedido Venta Cabecera.
                Dim dtContadorPred As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDefault, GetType(PedidoVentaCabecera).Name, services)
                If Not dtContadorPred Is Nothing AndAlso dtContadorPred.Rows.Count > 0 Then
                    data("IDContador") = dtContadorPred.Rows(0)("IDContador")
                    adr = dtContadores.Select("IDContador = " & Quoted(data("IDContador")))
                    If Not IsNothing(adr) AndAlso adr.Length > 0 Then
                        data("NPedido") = adr(0)("ValorProvisional")
                    End If
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarResponsable(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim strIDOper As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf Operario.ObtenerIDOperarioUsuario, Nothing, services)
        If Len(strIDOper) > 0 Then data("Responsable") = strIDOper
    End Sub

#End Region
#Region " Validate "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarClienteObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarClienteBloqueado)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFechaPedidoObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarMonedaObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFormaPagoObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarCondicionPagoObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoPedidoVenta.ValidacionesContabilidad)
    End Sub


#End Region
#Region " Update "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of UpdatePackage, DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.CrearDocumento)
        updateProcess.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.AsignarValoresPredeterminadosGenerales)
        'updateProcess.AddTask(Of DocumentoPedidoVenta)(AddressOf MetodosPedidos.AsignarDatosCliente)
        updateProcess.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.AsignarNumeroPedido)
        updateProcess.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.AsignarFechaAviso)
        updateProcess.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.AsignarEstadoPedido)
        updateProcess.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.AsignarCentroGestionPedido)
        updateProcess.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.ActualizarCambiosMoneda)
        updateProcess.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.AsignarFechaPreparacion)
        updateProcess.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.TratarPromocionesLineas)
        updateProcess.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.AsignarConfirmacionLineas)
        updateProcess.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.CalcularImporteLineasPedido)
        updateProcess.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.CalcularRepresentantes)
        updateProcess.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.CalcularAnalitica)
        updateProcess.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.CalcularBasesImponibles)
        updateProcess.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoComunes.TotalDocumento)
        updateProcess.AddTask(Of DocumentoPedidoVenta)(AddressOf Business.General.Comunes.BeginTransaction)
        updateProcess.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.ActualizarPrograma)
        updateProcess.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.Actualizarofertacomercial)
        updateProcess.AddTask(Of DocumentoPedidoVenta)(AddressOf Business.General.Comunes.UpdateDocument)
        updateProcess.AddTask(Of DocumentoPedidoVenta)(AddressOf Business.General.Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.ActualizarQLineasPromociones)
    End Sub

#End Region
#Region " Business Rules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New SynonymousBusinessRules
        oBRL.AddSynonymous("FechaPedido", "Fecha")

        Dim services As New ServiceProvider
        oBRL = ProcessServer.ExecuteTask(Of BusinessRules, BusinessRules)(AddressOf ProcesoComercial.DetailBusinessRulesCab, oBRL, services)

        oBRL("IDCliente") = AddressOf CambioClientePedido
        oBRL.Add("Fecha", AddressOf CambioFechaPedido)
        oBRL.Add("FechaPreparacion", AddressOf CambioFechasPreparacionAviso)
        oBRL.Add("FechaAviso", AddressOf CambioFechasPreparacionAviso)
        oBRL.Add("IDCentroGestion", AddressOf ProcesoComunes.CambioCentroGestion)
        oBRL.Add("IDDireccionEnvio", AddressOf ProcesoComercial.CambioDireccion)

        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioFechaPedido(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.ColumnName = "Fecha" Then
            '//Hay que ponerlo en los dos campos indicados en el Synonimous.
            data.Current(data.ColumnName) = data.Value
            data.Current("FechaPedido") = data.Value
        End If
        'ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoComunes.AsignarFechaEntrega, data.Current, services)
        ProcessServer.ExecuteTask(Of IPropertyAccessor)(AddressOf ProcesoComunes.AsignarEjercicioContablePedido, data.Current, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComunes.CambioMoneda, data, services)
    End Sub

    <Task()> Public Shared Sub CambioClientePedido(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComercial.CambioCliente, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComunes.CambioMoneda, data, services)
        Dim dir As New DataDireccionClte(enumpdTipoDireccion.pdDireccionPedido, "IDDireccionEnvio", data.Current)
        ProcessServer.ExecuteTask(Of DataDireccionClte)(AddressOf ProcesoComercial.AsignarDireccionCliente, dir, services)
        Dim Obs As New DataObservaciones(GetType(PedidoVentaCabecera).Name, "TextoComercial", data.Current)
        ProcessServer.ExecuteTask(Of DataObservaciones)(AddressOf ProcesoComercial.AsignarObservacionesCliente, Obs, services)

        If Length(data.Current("IDCliente")) > 0 Then
            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            data.Current(data.ColumnName) = data.Value
            Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data.Current("IDCliente"))

            If ClteInfo.Bloqueado Then ApplicationService.GenerateError("El Cliente está bloqueado.")
            If ClteInfo.Riesgo Then
                Dim StRiesgoClie As New Cliente.DataRiesgoCliente
                StRiesgoClie.IDCliente = ClteInfo.IDCliente
                data.Current("RiesgoCliente") = ProcessServer.ExecuteTask(Of Cliente.DataRiesgoCliente, RiesgoCliente)(AddressOf Cliente.ObtenerRiesgoCliente, StRiesgoClie, services)
            End If
            data.Current("DtoPedido") = ClteInfo.DtoComercial
            data.Current("Prioridad") = ClteInfo.Prioridad
            data.Current("IDModoTransporte") = ClteInfo.IDModoTransporte
        Else
            data.Current("DtoPedido") = 0
            data.Current("IDModoTransporte") = System.DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub CambioFechasPreparacionAviso(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If data.ColumnName = "FechaPreparacion" Then
            If Length(data.Current("FechaPreparacion")) > 0 AndAlso Nz(data.Current("FechaAviso"), Date.MinValue) = Date.MinValue Then
                ApplicationService.GenerateError("La fecha de aviso está vacía.")
            End If
        End If
        If IsDate(data.Current("FechaPreparacion")) Or IsDate(data.Current("FechaAviso")) Then
            If Nz(data.Current("FechaPreparacion"), Date.MinValue) < Nz(data.Current("FechaAviso"), Date.MinValue) Then
                ApplicationService.GenerateError("La fecha de preparación es menor que la de aviso.")
            End If
        End If
    End Sub

#End Region
#Region " Delete "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarProgramaDelete)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarPedidosMultiEmpresaDelete)
    End Sub

    <Task()> Public Shared Sub ActualizarProgramaDelete(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dt As DataTable = New ProgramaCabecera().Filter(New NumberFilterItem("IDPedido", data("IDPedido")))
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                dt.Rows(0)("IDPedido") = System.DBNull.Value
            Next
            BusinessHelper.UpdateTable(dt)
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarPedidosMultiEmpresaDelete(ByVal data As DataRow, ByVal services As ServiceProvider)
        '//control pedidos multiempresa
        Dim NPedidoCompra, NPedidoVenta, DescBaseDatos As String

        Dim grp As New GRPPedidoVentaCompraLinea
        Dim control As DataTable = grp.TrazaPVPrincipal(data("IDPedido"))
        If Not control Is Nothing AndAlso control.Rows.Count > 0 Then
            If Length(control.Rows(0)("IDPVSecundaria")) > 0 Then
                NPedidoVenta = control.Rows(0)("NPVSecundaria")
                DescBaseDatos = New NegocioGeneral().GetDataBaseDescription(CType(control.Rows(0)("IDBDSecundaria"), Guid))
                Throw New Exception("No se puede eliminar el pedido de venta. Este pedido ha generado el pedido de venta Nº " & NPedidoVenta & " en la empresa " & Quoted(DescBaseDatos) & ".")
            ElseIf Length(control.Rows(0)("IDPCPrincipal")) > 0 Then
                NPedidoCompra = control.Rows(0)("NPCPrincipal")
                Throw New Exception("No se puede eliminar el pedido de venta. Este pedido ha generado el pedido de compra Nº " & NPedidoCompra & ".")
            Else
                For Each dr As DataRow In control.Rows
                    grp.Delete(dr("IDPVLinea"))
                Next
            End If
        Else
            '//Este control tiene sentido en la bbdd secundaria. 
            '//En la bbdd principal nunca devolvera registros y por elllo es una comprobacion que no tiene ningun efecto.
            control = grp.TrazaPVSecundaria(data("IDPedido"))
            If Not control Is Nothing AndAlso control.Rows.Count > 0 Then
                For Each dr As DataRow In control.Rows
                    If Length(dr("IDPVSecundaria")) > 0 Then
                        dr("IDPVSecundaria") = DBNull.Value
                        dr("NPVSecundaria") = DBNull.Value
                        dr("IDLineaPVSecundaria") = DBNull.Value
                        dr("IDBDSecundaria") = DBNull.Value
                    End If
                Next
                BusinessHelper.UpdateTable(control)
            End If
        End If
    End Sub

#End Region
#Region " Precio Optimo "

    <Serializable()> _
    Public Class DataPrecioOptimo
        Public IDPedido As Integer
        Public FechaCalculo As Date
        Public DocPed As DocumentoPedidoVenta

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDPedido As Integer, ByVal FechaCalculo As Date)
            Me.IDPedido = IDPedido
            Me.FechaCalculo = FechaCalculo
        End Sub
    End Class

    <Serializable()> _
    Public Class DataCalcPrecioOpt
        Public FechaCalculo As Date
        Public DocPed As DocumentoPedidoVenta

        Public Sub New()
        End Sub
        Public Sub New(ByVal FechaCalculo As Date, ByVal DocPed As DocumentoPedidoVenta)
            Me.FechaCalculo = FechaCalculo
            Me.DocPed = DocPed
        End Sub
    End Class

    <Task()> Public Shared Sub PrecioOptimo(ByVal data As DataPrecioOptimo, ByVal services As ServiceProvider)
        Dim DocPed As DocumentoPedidoVenta = ProcessServer.ExecuteTask(Of Integer, DocumentoPedidoVenta)(AddressOf CrearDocumento, data.IDPedido, services)
        Dim StData As New DataCalcPrecioOpt(data.FechaCalculo, DocPed)
        ProcessServer.ExecuteTask(Of DataCalcPrecioOpt)(AddressOf CalculoPrecioOptimo, StData, services)
        ProcessServer.ExecuteTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.CalcularRepresentantes, DocPed, services)
        ProcessServer.ExecuteTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.CalcularAnalitica, DocPed, services)
        ProcessServer.ExecuteTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.CalcularBasesImponibles, DocPed, services)
        ProcessServer.ExecuteTask(Of DocumentoPedidoVenta)(AddressOf ProcesoComunes.TotalDocumento, DocPed, services)
        ProcessServer.ExecuteTask(Of DocumentoPedidoVenta)(AddressOf ProcesoPedidoVenta.GrabarDocumento, DocPed, services)
    End Sub

    <Task()> Public Shared Function CrearDocumento(ByVal IDPedido As Integer, ByVal services As ServiceProvider) As DocumentoPedidoVenta
        Return New DocumentoPedidoVenta(IDPedido)
    End Function

    <Task()> Public Shared Sub CalculoPrecioOptimo(ByVal data As DataCalcPrecioOpt, ByVal services As ServiceProvider)
        If data.DocPed Is Nothing OrElse data.DocPed.dtLineas Is Nothing OrElse data.DocPed.dtLineas.Rows.Count = 0 Then Exit Sub

        'Recogemos los articulos que esten relacionados con ese Pedido.
        Dim dtArticulosPedido As DataTable = New BE.DataEngine().Filter("vNegPedidoVentaLineaArticulos", New StringFilterItem("IDPedido", data.DocPed.HeaderRow("IDPedido")))
        Dim f As New Filter
        For Each drArticuloPedido As DataRow In dtArticulosPedido.Select
            f.Clear()
            f.Add("IDArticulo", drArticuloPedido("IDArticulo"))
            f.Add("Estado", FilterOperator.NotEqual, enumpvlEstado.pvlServido)
            'Recogemos las lineas del pedido que tengan el articulo de este momento
            Dim QPedida As Double = Nz(data.DocPed.dtLineas.Compute("SUM(QPedida)", f.Compose(New AdoFilterComposer)), 0)

            Dim dataTarifa As New DataCalculoTarifaComercial
            dataTarifa.IDArticulo = drArticuloPedido("IDArticulo")
            dataTarifa.IDCliente = data.DocPed.IDCliente
            dataTarifa.Cantidad = QPedida
            dataTarifa.Fecha = data.FechaCalculo

            ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial, DataTarifaComercial)(AddressOf ProcesoComercial.TarifaComercial, dataTarifa, services)
            If Not dataTarifa.DatosTarifa Is Nothing AndAlso dataTarifa.DatosTarifa.Precio <> 0 Then
                Dim PVL As New PedidoVentaLinea
                Dim context As New BusinessData(data.DocPed.HeaderRow)
                Dim WhereArticulosNoServidos As String = f.Compose(New AdoFilterComposer)
                For Each drAlbaranLineaArticulo As DataRow In data.DocPed.dtLineas.Select(WhereArticulosNoServidos)
                    PVL.ApplyBusinessRule("Precio", dataTarifa.DatosTarifa.Precio, drAlbaranLineaArticulo, context)
                    PVL.ApplyBusinessRule("Dto1", dataTarifa.DatosTarifa.Dto1, drAlbaranLineaArticulo, context)
                    PVL.ApplyBusinessRule("Dto2", dataTarifa.DatosTarifa.Dto2, drAlbaranLineaArticulo, context)
                    PVL.ApplyBusinessRule("Dto3", dataTarifa.DatosTarifa.Dto3, drAlbaranLineaArticulo, context)
                    PVL.ApplyBusinessRule("UDValoracion", dataTarifa.DatosTarifa.UDValoracion, drAlbaranLineaArticulo, context)

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
            QPedida = 0
        Next
    End Sub

#End Region

#Region " Seguimiento "
    Public Function SeguimientoPedidoCompra(ByVal IDPedido As Integer) As DataTable
        Return New BE.DataEngine().Filter("vConsultaPedidosCompraVenta", New NumberFilterItem("IDPVPrincipal", IDPedido))
    End Function
#End Region

End Class