Imports Solmicro.Expertis.Business.Negocio.NegocioGeneral

Public Class _AlbaranCompraCabecera
    Public Const IDAlbaran As String = "IDAlbaran"
    Public Const NAlbaran As String = "NAlbaran"
    Public Const IDContador As String = "IDContador"
    Public Const IDProveedor As String = "IDProveedor"
    Public Const FechaAlbaran As String = "FechaAlbaran"
    Public Const IDCentroGestion As String = "IDCentroGestion"
    Public Const SuAlbaran As String = "SuAlbaran"
    Public Const SuFecha As String = "SuFecha"
    Public Const Importe As String = "Importe"
    Public Const ImporteA As String = "ImporteA"
    Public Const ImporteB As String = "ImporteB"
    Public Const IDFormaEnvio As String = "IDFormaEnvio"
    Public Const IDDireccion As String = "IDDireccion"
    Public Const Texto As String = "Texto"
    Public Const IDAlmacen As String = "IDAlmacen"
    Public Const Estado As String = "Estado"
    Public Const IDTipoCompra As String = "IDTipoCompra"
    Public Const Dto As String = "Dto"
    Public Const IDCondicionPago As String = "IDCondicionPago"
    Public Const IDFormaPago As String = "IDFormaPago"
    Public Const IDMoneda As String = "IDMoneda"
    Public Const CambioA As String = "CambioA"
    Public Const CambioB As String = "CambioB"
    Public Const IDCondicionEnvio As String = "IDCondicionEnvio"
    Public Const IDModoTransporte As String = "IDModoTransporte"
    Public Const NMovimiento As String = "NMovimiento"
    Public Const Marca As String = "Marca"
    Public Const Automatico As String = "Automatico"
    Public Const FechaCreacionAudi As String = "FechaCreacionAudi"
    Public Const FechaModificacionAudi As String = "FechaModificacionAudi"
    Public Const UsuarioAudi As String = "UsuarioAudi"
    Public Const SyncDB As String = "SyncDB"
    Public Const Matricula As String = "Matricula"
    Public Const Remolque As String = "Remolque"
    Public Const Conductor As String = "Conductor"
    Public Const DNIConductor As String = "DNIConductor"
End Class

<Serializable()> _
Public Class AlbaranCompraUpdateData
    Public IDAlbaran() As String
    Public NAlbaran() As String
    Public StockUpdateData() As StockUpdateData
    Public PedidosError() As String
    Public ProveedoresError() As String
    Public MensajeError As String
End Class

<Serializable()> _
Public Class CrearAlbaranCompraInfo
    Implements IComparable, IComparer

    Public IDLinea As Integer
    Public Cantidad As Double
    Public CantidadUD As Double
    Public Cantidad2 As Double?
    Public IDPedido As Integer
    Public IDProveedor As String
    Public IDMoneda As String
    Public FechaEntregaModificado As Date
    Public SuAlbaran As String
    Public IDOrdenRuta As Integer
    Public Lotes As DataTable
    Public Series As DataTable
    Public IDTipoClasif As String
    Public IDTipoCompra As String
    Public TratarTipoClasif As Boolean = False
    Public TratarTipoCompra As Boolean = False
    Public SuFecha As Date

    Public Sub New()
    End Sub

    Public Sub New(ByVal IDLinea As Integer)
        Me.IDLinea = IDLinea
    End Sub

    Public Function CompareTo(ByVal obj As Object) As Integer Implements System.IComparable.CompareTo
        If TypeOf obj Is CrearAlbaranCompraInfo Then
            Dim p As CrearAlbaranCompraInfo = CType(obj, CrearAlbaranCompraInfo)
            Return Me.IDLinea.CompareTo(p.IDLinea)
        Else
            Throw New ArgumentException("El objeto no es del tipo CrearAlbaranCompraInfo.")
        End If
    End Function

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
        If TypeOf x Is CrearAlbaranCompraInfo And TypeOf y Is CrearAlbaranCompraInfo Then
            Return CType(x, CrearAlbaranCompraInfo).CompareTo(y)
        Else
            Throw New ArgumentException("El objeto no es del tipo CrearAlbaranCompraInfo.")
        End If
    End Function

    Public Shared Function Find(ByVal a As Array, ByVal IDLinea As Integer) As CrearAlbaranCompraInfo
        Dim i As Integer
        i = Array.BinarySearch(a, New CrearAlbaranCompraInfo(IDLinea))
        If i >= 0 Then
            Return a(i)
        End If
    End Function
End Class

Public Class AlbaranCompraCabeceraInfo
    Inherits ClassEntityInfo

    Public IDAlbaran As Integer
    Public NAlbaran As String
    Public FechaAlbaran As Date
    Public IDCentroGestion As String
    Public IDMoneda As String

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dt As DataTable
        If Not IsNothing(PrimaryKey) AndAlso PrimaryKey.Length > 0 AndAlso Length(PrimaryKey(0)) > 0 Then
            dt = New AlbaranCompraCabecera().SelOnPrimaryKey(PrimaryKey(0))
        End If

        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            Me.Fill(dt.Rows(0))
        End If
    End Sub

End Class

Public Class AlbaranCompraCabecera
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbAlbaranCompraCabecera"

    Private _ACC As _AlbaranCompraCabecera
    Private _ACL As _AlbaranCompraLinea
    Private _PCL As _PedidoCompraLinea
    Private _PCC As _PedidoCompraCabecera
    Private _AAL As _ArticuloAlmacenLote
    Private _ACLT As _AlbaranCompraLote

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
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarEstado, data, services)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf ProcesoCompra.AsignarTipoCompra, data, services)
    End Sub

    <Task()> Public Shared Sub AsignarContador(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim CE As New CentroEntidad
        CE.CentroGestion = data("IDCentroGestion") & String.Empty
        CE.ContadorEntidad = CentroGestion.ContadorEntidad.AlbaranCompra
        data("IDContador") = ProcessServer.ExecuteTask(Of CentroEntidad, String)(AddressOf CentroGestion.GetContadorPredeterminado, CE, services)
    End Sub

    <Task()> Public Shared Sub AsignarNumeroAlbaranProvisional(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDContador")) > 0 Then
            Dim dtContadores As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDt, GetType(AlbaranCompraCabecera).Name, services)
            Dim adr As DataRow() = dtContadores.Select("IDContador = " & Quoted(data("IDContador")))
            If Not IsNothing(adr) AndAlso adr.Length > 0 Then
                data("NAlbaran") = adr(0)("ValorProvisional")
            Else
                'Si no está bien configurado el Contador de Albaranes de Compra en el Centro de Gestión,
                'cogemos el Contador por defecto de la entidad Albaran Compra Cabecera.
                Dim dtContadorPred As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf Contador.CounterDefault, GetType(AlbaranCompraCabecera).Name, services)
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
        If data.IsNull("Estado") Then data("Estado") = enumaccEstado.accNoFacturado
    End Sub

#End Region

#Region " Delete "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ValidarDelAlbaranFacturado)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarAlbaranesMultiEmpresa)
    End Sub
    <Task()> Public Shared Sub ValidarDelAlbaranFacturado(ByVal HeaderRow As DataRow, ByVal services As ServiceProvider)
        If HeaderRow("Estado") = enumaccEstado.accParcFacturado OrElse _
           HeaderRow("Estado") = enumaccEstado.accFacturado Then
            ApplicationService.GenerateError("No se puede borrar el Albarán. Está Facturado o Parcialmente Facturado.")
        End If
    End Sub
    <Task()> Public Shared Sub ActualizarAlbaranesMultiEmpresa(ByVal HeaderRow As DataRow, ByVal services As ServiceProvider)
        Dim control As DataTable = New GRPAlbaranVentaCompraLinea().TrazaACPrincipal(HeaderRow("IDAlbaran"))
        If Not control Is Nothing AndAlso control.Rows.Count > 0 Then
            For Each dr As DataRow In control.Rows
                dr("IDACPrincipal") = DBNull.Value
                dr("NACPrincipal") = DBNull.Value
                dr("IDLineaACPrincipal") = DBNull.Value
            Next
            BusinessHelper.UpdateTable(control)
        End If
    End Sub

#End Region

#Region " Validate "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoAlbaranCompra.ValidarAlbaranFacturado)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFechaAlbaranObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarAlbaranObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarProveedorObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarMonedaObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoAlbaranCompra.ValidarCondicionesEconomicas)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoAlbaranCompra.ValidacionesContabilidad)
        validateProcess.AddTask(Of DataRow)(AddressOf ProcesoAlbaranCompra.ValidarSuNumAlbaranFecha)
        'validateProcess.AddTask(Of DataRow)(AddressOf ProcesoAlbaranCompra.ValidarNumeroAlbaran)
    End Sub

#End Region

#Region " Update "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)

        updateProcess.AddTask(Of UpdatePackage, DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.CrearDocumento)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.AsignarDireccion)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoComunes.AsignarCentroGestion)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoComunes.AsignarAlmacen)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoComunes.AsignarNumeroAlbaran)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoComunes.AsignarCambiosMoneda)
        'updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.ValidarDocumento)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.ActualizarEstadoLineas)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.ActualizarEstadoAlbaran)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.GestionArticulosKit)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.GestionCalidadArticulo)        'PENDIENTE  (Ver Update de las líneas)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoComunes.CalcularImporteLineasAlbaran)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.ValoracionSuministro)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.CalcularAnalitica)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf Business.General.Comunes.BeginTransaction)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.CalcularAlbaranCompraGastos)
        'updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.AsignarAlbaranCompraLotes)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.TotalPesos)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.CalcularBasesImponibles)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoComunes.TotalDocumento)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.PrepararArticulosUltimaCompra)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.CorregirMovimientos)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.ActualizarPedidoDesdeAlbaran)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.ActualizarProgramaDesdeAlbaran)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.ActualizarDAAARCBodegas)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf Business.General.Comunes.UpdateDocument)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf Business.General.Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.ActualizarObras)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.ActualizacionAutomaticaStock)
        updateProcess.AddTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.DetalleActualizacionStocks)
    End Sub

#End Region

#Region " BusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim services As New ServiceProvider
        Dim oBRL As New SynonymousBusinessRules
        oBRL.AddSynonymous("FechaAlbaran", "Fecha")

        oBRL = ProcessServer.ExecuteTask(Of BusinessRules, BusinessRules)(AddressOf ProcesoCompra.DetailBusinessRulesCab, oBRL, services)

        oBRL("IDProveedor") = AddressOf CambioProveedor
        oBRL.Add("Fecha", AddressOf ProcesoComunes.CambioFechaAlbaran)
        oBRL.Add("IDAlmacen", AddressOf ProcesoComunes.CambioAlmacen)
        oBRL.Add("IDTipoCompra", AddressOf CambioTipoCompra)
        oBRL.Add("IDCentroGestion", AddressOf ProcesoComunes.CambioCentroGestion)

        oBRL.Add("IDFormaEnvio", AddressOf CambioIDFormaEnvio)
        oBRL.Add("IDTransportista", AddressOf CambioIDTransportista)
        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioProveedor(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoCompra.CambioProveedor, data, services)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf ProcesoComunes.CambioMoneda, data, services)
        Dim dir As New DataDireccionProv(enumpdTipoDireccion.pdDireccionPedido, "IDDireccion", data.Current)
        ProcessServer.ExecuteTask(Of DataDireccionProv)(AddressOf ProcesoCompra.AsignarDireccionProveedor, dir, services)
        Dim obs As New DataObservaciones(GetType(AlbaranCompraCabecera).Name, "Texto", data.Current)
        ProcessServer.ExecuteTask(Of DataObservaciones)(AddressOf ProcesoCompra.AsignarObservacionesProveedor, obs, services)

        If Length(data.Current("IDProveedor")) > 0 Then
            Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()
            Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(data.Current("IDProveedor"))
            data.Current("IDModoTransporte") = ProvInfo.IDModoTransporte
            data.Current("Dto") = ProvInfo.DtoComercial
            'data.Current("IDPais") = ProvInfo.IDPais
            'data.Current("Telefono") = ProvInfo.Telefono
            'data.Current("Fax") = ProvInfo.Fax
            'data.Current("IdBancoPropio") = ProvInfo.IDBancoPropio
            data.Current = New AlbaranCompraCabecera().ApplyBusinessRule("IDFormaEnvio", data.Current("IDFormaEnvio"), data.Current, data.Context)
        Else
            data.Current("IDModoTransporte") = System.DBNull.Value
            data.Current("Dto") = 0
            'data.Current("IDPais") = System.DBNull.Value
            'data.Current("Telefono") = System.DBNull.Value
            'data.Current("Fax") = System.DBNull.Value
            'data.Current("IdBancoPropio") = System.DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub CambioTipoCompra(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Current("IDAlbaran")) > 0 Then
            Dim dtACL As DataTable = New AlbaranCompraLinea().Filter(New NumberFilterItem("IDAlbaran", data.Current("IDAlbaran")))
            If Not IsNothing(dtACL) AndAlso dtACL.Rows.Count > 0 Then
                ApplicationService.GenerateError("No es posible modificar el Tipo Compra, existen líneas de albarán.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CCInmovilizado(ByVal current As IPropertyAccessor, ByVal services As ServiceProvider)
        If Length(current("IDEjercicio")) > 0 AndAlso Length(current("IDAlbaran")) > 0 Then
            '// Comprobar si para las líneas del Albaran que existan, las C.Contables son de inmovilizado o 
            '// no según corresponda. 
            Dim dtACL As DataTable = New AlbaranCompraLinea().Filter(New NumberFilterItem("IDAlbaran", current("IDAlbaran")))
            If dtACL.Rows.Count > 0 Then
                If dtACL.Columns.Contains("Inmovilizado") AndAlso dtACL.Columns.Contains("CContable") Then
                    Dim data As New ProcesoCompra.DataCContableInmovilizado
                    For Each drRow As DataRow In dtACL.Rows
                        data.IDEjercicio = current("IDEjercicio") & String.Empty
                        data.CContable = drRow("CContable") & String.Empty
                        data.Inmovilizado = Nz(drRow("Inmovilizado"), False)
                        ProcessServer.ExecuteTask(Of ProcesoCompra.DataCContableInmovilizado)(AddressOf ProcesoCompra.ValidarCuentaInmovilizado, data, services)
                    Next
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioIDFormaEnvio(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim ACC As New AlbaranCompraCabecera
            Dim fe As New FormaEnvio
            Dim filtro As New StringFilterItem("IDFormaEnvio", FilterOperator.Equal, data.Value)
            Dim dtEnvio As DataTable = New FormaEnvio().Filter(filtro)
            If (Not dtEnvio Is Nothing AndAlso dtEnvio.Rows.Count > 0) Then
                data.Current("IdFormaEnvio") = dtEnvio.Rows(0)("IDFormaEnvio")
                data.Current("EmpresaTransp") = dtEnvio.Rows(0)("DescFormaEnvio")
                Dim dt As DataTable = fe.SelOnPrimaryKey(dtEnvio.Rows(0)("IDFormaEnvio"))
                data.Current("IDTransportista") = dt.Rows(0)("IDProveedor")
                data.Current = ACC.ApplyBusinessRule("IDTransportista", dt.Rows(0)("IDProveedor"), data.Current)
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

    <Task()> Public Shared Sub CambioIDTransportista(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Current("IDTransportista")) > 0 Then
            If Length(data.Value) > 0 Then
                Dim dr As DataRow = New Proveedor().GetItemRow(data.Value)
                data.Current("CifTransportista") = dr("CifProveedor")
            End If
        Else
            data.Current("CifTransportista") = DBNull.Value
        End If
    End Sub

#End Region

#Region " Precio Optimo "

    <Task()> Public Shared Sub PrecioOptimo(ByVal IDAlbaran As Integer, ByVal services As ServiceProvider)
        Dim DocAlb As DocumentoAlbaranCompra = ProcessServer.ExecuteTask(Of Integer, DocumentoAlbaranCompra)(AddressOf CrearDocumento, IDAlbaran, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra)(AddressOf CalculoPrecioOptimo, DocAlb, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.CalcularAnalitica, DocAlb, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.CalcularBasesImponibles, DocAlb, services)
        'ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranVenta.CalcularImportesAlbaran, DocAlb, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoComunes.TotalDocumento, DocAlb, services)
        ProcessServer.ExecuteTask(Of DocumentoAlbaranCompra)(AddressOf ProcesoAlbaranCompra.GrabarDocumento, DocAlb, services)
    End Sub

    <Task()> Public Shared Function CrearDocumento(ByVal IDAlbaran As Integer, ByVal services As ServiceProvider) As DocumentoAlbaranCompra
        Return New DocumentoAlbaranCompra(IDAlbaran)
    End Function

    <Task()> Public Shared Sub CalculoPrecioOptimo(ByVal Doc As DocumentoAlbaranCompra, ByVal services As ServiceProvider)
        If Doc Is Nothing OrElse Doc.dtLineas Is Nothing OrElse Doc.dtLineas.Rows.Count = 0 Then Exit Sub

        '//Recogemos los articulos que esten relacionados con esa Albaran.
        Dim dtArticulosAlbaran As DataTable = New BE.DataEngine().Filter("vNegAlbaranCompraLineaArticulos", New StringFilterItem("IDAlbaran", Doc.HeaderRow("IDAlbaran")))
        Dim f As New Filter
        For Each drArticuloAlbaran As DataRow In dtArticulosAlbaran.Select
            f.Clear()
            f.Add("IDArticulo", drArticuloAlbaran("IDArticulo"))

            '//Recogemos las lineas del albarán que tengan el articulo de este momento
            Dim QServida As Double = Nz(Doc.dtLineas.Compute("SUM(QServida)", f.Compose(New AdoFilterComposer)), 0)

            Dim dataTarifa As New DataCalculoTarifaCompra
            dataTarifa.IDArticulo = drArticuloAlbaran("IDArticulo")
            dataTarifa.IDProveedor = Doc.IDProveedor
            dataTarifa.Cantidad = QServida
            dataTarifa.Fecha = Doc.Fecha
            dataTarifa.IDMoneda = Doc.IDMoneda
            ProcessServer.ExecuteTask(Of DataCalculoTarifaCompra)(AddressOf ProcesoCompra.TarifaCompra, dataTarifa, services)

            'Dim dtTarifa As DataTable = ProcesoCompra.TarifaCompra(drArticuloAlbaran("IDArticulo"), Doc.IDProveedor, QServida, Doc.Fecha)
            'If Not dataTarifa Is Nothing AndAlso dtTarifa.Rows.Count > 0 Then
            'If Doc.IDMoneda <> dataTarifa.DatosTarifa.IDMoneda Then
            '    dtTarifa = General.CambioMoneda(dtTarifa, dtTarifa.Rows(0)("IDMoneda"), Doc.IDMoneda, Doc.Fecha)
            'End If
            'If Not dtTarifa Is Nothing AndAlso dtTarifa.Rows.Count > 0 Then
            If Not dataTarifa.DatosTarifa Is Nothing Then

                Dim ACL As New AlbaranCompraLinea
                Dim context As New BusinessData(Doc.HeaderRow)
                Dim WhereArticulo As String = f.Compose(New AdoFilterComposer)
                For Each drAlbaranLineaArticulo As DataRow In Doc.dtLineas.Select(WhereArticulo)
                    If AreDifferents(drAlbaranLineaArticulo("EstadoFactura"), enumaclEstadoFactura.aclFacturado) Then
                        If AreDifferents(drAlbaranLineaArticulo("TipoLineaAlbaran"), enumaclTipoLineaAlbaran.aclComponente) Then
                            ACL.ApplyBusinessRule("Precio", dataTarifa.DatosTarifa.Precio, drAlbaranLineaArticulo, context)
                            ACL.ApplyBusinessRule("Dto1", dataTarifa.DatosTarifa.Dto1, drAlbaranLineaArticulo, context)
                            ACL.ApplyBusinessRule("Dto2", dataTarifa.DatosTarifa.Dto2, drAlbaranLineaArticulo, context)
                            ACL.ApplyBusinessRule("Dto3", dataTarifa.DatosTarifa.Dto3, drAlbaranLineaArticulo, context)
                            ACL.ApplyBusinessRule("UDValoracion", dataTarifa.DatosTarifa.UDValoracion, drAlbaranLineaArticulo, context)
                            If Length(dataTarifa.DatosTarifa.SeguimientoTarifa) > 0 Then
                                drAlbaranLineaArticulo("SeguimientoTarifa") = dataTarifa.DatosTarifa.SeguimientoTarifa
                            End If
                        End If
                    End If
                Next
            End If
            'End If
            QServida = 0
        Next
    End Sub

#End Region

#Region " PrepararOrdenRuta "

    <Task()> Public Shared Function PrepararOrdenRuta(ByVal Lineas As DataTable, ByVal services As ServiceProvider) As DataTable()
        Dim dt As DataTable
        Dim LineasRuta(-1) As DataTable

        Dim DtFinal As DataTable = Lineas.Clone
        Dim IntID As Integer
        For Each DrN As DataRow In Lineas.Select("", "IDOrdenRuta")
            If IntID <> DrN("IDOrdenRuta") Then
                IntID = DrN("IDOrdenRuta")
                DtFinal.Rows.Add(DrN.ItemArray)
            Else : DtFinal.Rows(DtFinal.Rows.Count - 1)("QServida") += DrN("QServida")
            End If
        Next
        For Each dr As DataRow In DtFinal.Select()
            dt = ProcessServer.ExecuteTask(Of DataRow, DataTable)(AddressOf PrepararOrdenRutaLinea, dr, services)
            ReDim Preserve LineasRuta(UBound(LineasRuta) + 1) : LineasRuta(UBound(LineasRuta)) = dt
        Next
        Return LineasRuta
    End Function

    <Task()> Public Shared Function PrepararOrdenRutaLinea(ByVal lineas As DataRow, ByVal services As ServiceProvider) As DataTable
        Dim blnEsConsumo As Boolean
        Dim dblQServidaNew As Double
        Dim dblQServidaOld As Double
        Dim strIN As String
        Dim dblIncQPedida, dblIncQServida As Double

        If Length(lineas("IDLineaPedido")) > 0 And Length(lineas("IDOrdenRuta")) > 0 And lineas("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclSubcontratacion Then
            Dim ORuta As BusinessHelper = BusinessHelper.CreateBusinessObject("OrdenRuta")
            Dim dtOrdenRuta As DataTable = ORuta.Filter(New NumberFilterItem("IDOrdenRuta", lineas("IDOrdenRuta")))

            If Not dtOrdenRuta Is Nothing AndAlso dtOrdenRuta.Rows.Count > 0 Then
                dblQServidaNew = lineas("QServida")
                If lineas.RowState = DataRowState.Modified Then
                    dblQServidaOld = lineas("QServida", DataRowVersion.Original)
                End If
                For Each OrdenRuta As DataRow In dtOrdenRuta.Rows
                    '  dblIncQPedida = lineas("QPedida") - lineas("QPedida").OriginalValue
                    dblIncQServida = lineas("QServida") - dblQServidaOld
                    OrdenRuta("QEnviada") = OrdenRuta("QEnviada") + (lineas("Factor") * (dblIncQPedida - dblIncQServida))
                    If OrdenRuta("QEnviada") < 0 Then OrdenRuta("QEnviada") = 0
                Next
                Return dtOrdenRuta
            End If
        End If

    End Function

#End Region

#Region " Consultas interactivas (Estadísticas) "

    <Task()> Public Shared Function ObtenerEstadisticaACTipos(ByVal obj As Object, ByVal services As ServiceProvider) As DataTable
        Return New BE.DataEngine().Filter("tbEstadisticaCompraAnual", Nothing)
    End Function

    <Serializable()> _
    Public Class DataEstadisticaCantidadMeses
        Public CamposSelect As String
        Public CampoAtotalizar As String
        Public IDTipo As String
        Public IDFamilia As String
        Public IDArticulo As String
        Public IDProveedor As String
        Public Provincia As String
        Public IDZona As String
        Public IDMercado As String
        Public IDPais As String
        Public CEE As enumBoolean
        Public Extranjero As enumBoolean
        Public Año As Integer
        Public EmpresasGrupo As Integer
        Public GroupBy As String
        Public CamposOrden As String

        Public Sub New(ByVal CamposSelect As String, ByVal CampoATotalizar As String, ByVal IDTipo As String, ByVal IDFamilia As String, _
                       ByVal IDArticulo As String, ByVal IDProveedor As String, ByVal Provincia As String, ByVal IDZona As String, ByVal IDMercado As String, _
                       ByVal IDPais As String, ByVal CEE As enumBoolean, ByVal Extranjero As enumBoolean, ByVal Año As Integer, _
                       ByVal EmpresasGrupo As Integer, ByVal GroupBy As String, ByVal CamposOrden As String)
            Me.CamposSelect = CamposSelect
            Me.CampoAtotalizar = CampoATotalizar
            Me.IDTipo = IDTipo
            Me.IDFamilia = IDFamilia
            Me.IDArticulo = IDArticulo
            Me.IDProveedor = IDProveedor
            Me.Provincia = Provincia
            Me.IDZona = IDZona
            Me.IDMercado = IDMercado
            Me.IDPais = IDPais
            Me.CEE = CEE
            Me.Extranjero = Extranjero
            Me.Año = Año
            Me.EmpresasGrupo = EmpresasGrupo
            Me.GroupBy = GroupBy
            Me.CamposOrden = CamposOrden
        End Sub
    End Class

    <Task()> Public Shared Function ObtenerEstadisticaCantidadesMeses(ByVal data As DataEstadisticaCantidadMeses, ByVal services As ServiceProvider) As DataTable
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
            "SUM({1}) As STotalLinea", data.CamposSelect, data.CampoAtotalizar))

        selectSQL.Append(" FROM tbMaestroMercado RIGHT OUTER JOIN" & _
            " tbAlbaranCompraLinea INNER JOIN" & _
            " tbAlbaranCompraCabecera ON tbAlbaranCompraLinea.IDAlbaran = tbAlbaranCompraCabecera.IDAlbaran INNER JOIN" & _
            " tbMaestroProveedor ON tbAlbaranCompraCabecera.IDProveedor = tbMaestroProveedor.IDProveedor INNER JOIN" & _
            " tbMaestroArticulo ON tbAlbaranCompraLinea.IDArticulo = tbMaestroArticulo.IDArticulo INNER JOIN" & _
            " tbMaestroPais ON tbMaestroProveedor.IDPais = tbMaestroPais.IDPais LEFT OUTER JOIN" & _
            " tbMaestroSubFamilia ON tbMaestroArticulo.IDSubFamilia = tbMaestroSubFamilia.IDSubFamilia AND" & _
            " tbMaestroArticulo.IDFamilia = tbMaestroSubFamilia.IDFamilia AND" & _
            " tbMaestroArticulo.IDTipo = tbMaestroSubFamilia.IDTipo LEFT OUTER JOIN" & _
            " tbMaestroFamilia ON tbMaestroArticulo.IDFamilia = tbMaestroFamilia.IDFamilia AND" & _
            " tbMaestroArticulo.IDTipo = tbMaestroFamilia.IDTipo LEFT OUTER JOIN" & _
            " tbMaestroZona ON tbMaestroProveedor.IDZona = tbMaestroZona.IDZona ON" & _
            " tbMaestroMercado.IDMercado = tbMaestroProveedor.IDMercado")

        Dim whereSQL As New Text.StringBuilder
        'para que no entren en la estadística las líneas de componentes de albaranes de subcontratación.
        whereSQL.Append("tbAlbaranCompraLinea.TipoLineaAlbaran <> 2 AND ")

        If Length(data.Año) > 0 Then
            whereSQL.Append("YEAR(tbAlbaranCompraCabecera.FechaAlbaran) = " & data.Año & " AND ")
        End If
        If Length(data.IDTipo) > 0 Then
            whereSQL.Append("tbMaestroArticulo.IDTipo = '" & data.IDTipo & "' AND ")
        End If
        If Length(data.IDFamilia) > 0 Then
            whereSQL.Append("tbMaestroArticulo.IDFamilia = '" & data.IDFamilia & "' AND ")
        End If
        If Length(data.IDArticulo) > 0 Then
            whereSQL.Append("tbMaestroArticulo.IDArticulo = '" & data.IDArticulo & "' AND ")
        End If
        If Length(data.IDProveedor) > 0 Then
            whereSQL.Append("tbMaestroProveedor.IDProveedor = '" & data.IDProveedor & "' AND ")
        End If
        If Length(data.Provincia) > 0 Then
            whereSQL.Append("tbMaestroProveedor.Provincia = '" & data.Provincia & "' AND ")
        End If
        If Length(data.IDZona) > 0 Then
            whereSQL.Append("tbMaestroProveedor.IDZona = '" & data.IDZona & "' AND ")
        End If
        If Length(data.IDMercado) > 0 Then
            whereSQL.Append("tbMaestroProveedor.IDMercado = '" & data.IDMercado & "' AND ")
        End If
        If Length(data.IDPais) > 0 Then
            whereSQL.Append("tbMaestroPais.IDPais = '" & data.IDPais & "' AND ")
        End If

        Select Case data.CEE
            Case enumBoolean.Si
                whereSQL.Append("tbMaestroPais.CEE = 1 AND ")
            Case enumBoolean.No
                whereSQL.Append("tbMaestroPais.CEE = 0 AND ")
        End Select

        Select Case data.EmpresasGrupo
            Case enumBoolean.Si
                whereSQL.Append("tbMaestroProveedor.EmpresaGrupo = 1 AND ")
            Case enumBoolean.No
                whereSQL.Append("tbMaestroProveedor.EmpresaGrupo = 0 AND ")
        End Select

        Select Case data.Extranjero
            Case enumBoolean.Si
                whereSQL.Append("tbMaestroPais.Extranjero = 1 AND ")
            Case enumBoolean.No
                whereSQL.Append("tbMaestroPais.Extranjero = 0 AND ")
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

#Region " Actualización de stocks "

    <Task()> Public Shared Function ActualizarStockAlbaranes(ByVal IDAlbaran() As Integer, ByVal services As ServiceProvider) As StockUpdateData()
        Dim updateDataArray(-1) As StockUpdateData
        For Each id As Integer In IDAlbaran
            Dim updateData() As StockUpdateData = ProcessServer.ExecuteTask(Of Integer, StockUpdateData())(AddressOf ActualizarStockAlbaran, id, services)
            If updateData.Length > 0 Then
                ArrayManager.Copy(updateData, updateDataArray)
            End If
        Next
        Return updateDataArray
    End Function

    <Task()> Public Shared Function ActualizarStockAlbaran(ByVal IDAlbaran As Integer, ByVal services As ServiceProvider) As StockUpdateData()
        Dim updateDataArray(-1) As StockUpdateData
        If IDAlbaran <> 0 Then
            Dim Doc As New DocumentoAlbaranCompra(IDAlbaran)
            Dim actLin As New ProcesoStocks.DataActualizarStockLineas(Doc)
            Dim updateData() As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarStockLineas, StockUpdateData())(AddressOf ProcesoAlbaranCompra.ActualizarStockLineas, actLin, services)
            If updateData.Length > 0 Then
                ArrayManager.Copy(updateData, updateDataArray)
            End If
        End If
        Return updateDataArray
    End Function

    <Task()> Public Shared Function ActualizarStockLineasAlbaran(ByVal IDProcess As Guid, ByVal services As ServiceProvider) As StockUpdateData()
        Dim updateDataArray(-1) As StockUpdateData
        Dim LineasAlbaran As DataTable = New BE.DataEngine().Filter("vFrmActualizacionAlbaranCompraLinea", New GuidFilterItem("IDProcess", IDProcess))
        If LineasAlbaran.Rows.Count > 0 Then
            Dim IDAlbaranes(-1) As Integer
            Dim IDLineasAlbaran(-1) As Integer
            'Dim Doc As DocumentoAlbaranCompra
            For Each linea As DataRow In LineasAlbaran.Select(Nothing, "IDAlbaran")
                If Array.IndexOf(IDAlbaranes, linea("IDAlbaran")) < 0 Then
                    ReDim Preserve IDAlbaranes(IDAlbaranes.Length)
                    IDAlbaranes(IDAlbaranes.Length - 1) = linea("IDAlbaran")
                End If
            Next

            For Each IDAlbaran As Integer In IDAlbaranes
                Dim Doc As New DocumentoAlbaranCompra(IDAlbaran)
                ReDim IDLineasAlbaran(-1)
                For Each LineaAlbaran As DataRow In LineasAlbaran.Select("IDAlbaran=" & IDAlbaran)
                    ReDim Preserve IDLineasAlbaran(IDLineasAlbaran.Length)
                    IDLineasAlbaran(IDLineasAlbaran.Length - 1) = LineaAlbaran("IDLineaAlbaran")
                Next
                Dim actLin As New ProcesoStocks.DataActualizarStockLineas(Doc, IDLineasAlbaran)
                Dim updateData() As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarStockLineas, StockUpdateData())(AddressOf ProcesoAlbaranCompra.ActualizarStockLineas, actLin, services)
                If updateData.Length > 0 Then
                    ArrayManager.Copy(updateData, updateDataArray)
                End If
            Next
        End If
        Return updateDataArray
    End Function

#End Region

#Region " Números de Serie "

    <Task()> Public Shared Function ComprobarNumerosSerieAlbaranes(ByVal IDAlbaranes() As Integer, ByVal services As ServiceProvider) As Boolean
        If Not IDAlbaranes Is Nothing AndAlso IDAlbaranes.Length > 0 Then
            For Each IDAlbaran As Integer In IDAlbaranes
                If Not ProcessServer.ExecuteTask(Of Integer, Boolean)(AddressOf ComprobarNumerosSerieAlbaran, IDAlbaran, services) Then
                    Return False
                End If
            Next
        End If
        Return True
    End Function

    <Task()> Public Shared Function ComprobarNumerosSerieAlbaran(ByVal IntAlbaran As Integer, ByVal services As ServiceProvider) As Boolean
        Dim DtLineas As DataTable = New AlbaranCompraLinea().Filter(New NumberFilterItem("IDAlbaran", IntAlbaran))
        If Not DtLineas Is Nothing AndAlso DtLineas.Rows.Count > 0 Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            For Each Dr As DataRow In DtLineas.Select
                Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(Dr("IDArticulo"))
                If Not ArtInfo Is Nothing AndAlso Length(ArtInfo.IDArticulo) > 0 Then
                    If ArtInfo.NSerieObligatorio AndAlso Length(Dr("Lote")) = 0 Then
                        Return False
                    End If
                End If
            Next
        End If
        Return True
    End Function

#End Region

End Class