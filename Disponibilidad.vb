Imports System.Math

<Serializable()> _
Public Class DisponibilidadInfo
    Public IDArticulo As String
    Public IDAlmacen As String
    Public IDTipo As String
    Public IDFamilia As String
    Public IDSubfamilia As String
    Public Pendiente As String
    Public Disponible As String
    Public FechaHasta As Date
    <Serializable()> _
    Public Class GeneracionPedidosCompraInfo
        Public IDPedidoVenta1 As Integer        '//IDPedido Venta Origen
        Public NPedidoVenta1 As String          '//NPedido Venta Origen
        Public IDPedidoCompra As Integer        '//IDPedido Compra creado
        Public NPedidoCompra As String          '//NPedido Compra creado
        Public Proveedor As String              '//ID - Desc. Proveedor
        Public EmpresaGrupo As Boolean          '//Dato de Origen (el proveedor es de Empresa de Grupo)
        Public EntregaProveedor As Boolean      '//Dato de Origen (el proveedor entrega la mercancía)
        Public IDPedidoVenta2 As Integer        '//IDPedido Venta en BBDD Multiempresa
        Public NPedidoVenta2 As String          '//NPedido Venta en BBDD Multiempresa
        Public Cliente As String
        Public BaseDatos1 As String             '//BBDD Origen
        Public BaseDatos2 As String             '//BBDD Destino
    End Class

    Public Sub New()
    End Sub

    Public Sub New(ByVal articulo As String, ByVal almacen As String, ByVal fechahasta As Date)
        Me.IDArticulo = articulo
        Me.IDAlmacen = almacen
        Me.FechaHasta = fechahasta
    End Sub
End Class

'inicio: Se añaden los nuevos campos del Filtrado
<Serializable()> _
Public Class DisponibilidadporArticuloAlmacen
    Inherits DisponibilidadInfo

    Public DifPtoPedido As String
    Public DifStockSeg As String
End Class
'Fin

<Serializable()> _
Public Class DatosDisponibilidadOfs
    Public EntradasOf As DataTable
    Public ConsumosOf As DataTable
End Class

<Serializable()> _
Public Class DatosDisponibilidadAlbaranes
    Public AlbaranesCompraPendientes As DataTable
    Public AlbaranesVentaDeposito As DataTable
    Public AlbaranesPendientesExpedir As DataTable
End Class

<Serializable()> _
Public Class DisponibilidadVentaInfo
    Inherits DisponibilidadInfo

    Public IDCliente As String
    Public Estado As String
    Public Prioridad As String
    Public NPedidoDesde As String
    Public NPedidoHasta As String
End Class

<Serializable()> _
Public Class DisponibilidadCompraInfo
    Inherits DisponibilidadInfo

    Public IDProveedor As String
    Public Estado As String
    Public NPedidoDesde As String
    Public NPedidoHasta As String
End Class

<Serializable()> _
Public Class DisponibilidadExpObraInfo
    Inherits DisponibilidadInfo

    Public IDCliente As String
    Public NObra As String
End Class

<Serializable()> _
Public Class GeneracionPedidosCompraInfo
    Public IDPedidoVenta1 As Integer
    Public NPedidoVenta1 As String
    Public Proveedor As String
    Public EmpresaGrupo As Boolean
    Public EntregaProveedor As Boolean
    Public IDPedidoCompra As Integer
    Public NPedidoCompra As String
    Public IDPedidoVenta2 As Integer
    Public NPedidoVenta2 As String
    Public Cliente As String
    Public BaseDatos1 As String
    Public BaseDatos2 As String
    Public StrError As String
End Class

<Serializable()> _
Public Class EvolucionDisponibilidadInfo
    Public IDArticulo As String
    Public IDAlmacen As String
    Public Fecha As Date
    Public StockFisico As Double
    Public Datos As DataTable
End Class

<Transactional()> _
Public Class Disponibilidad
    Inherits ContextBoundObject

    'pend las tablas que mas registros pueden tener son las de pedidos, albaranes, 
    'revisar si es necesario indexar ciertos campos (Estado de las lineas, etc)
    'para agilizar el proceso

    Private Const viewArticuloAlmacen As String = "vDisponibilidadArticuloAlmacen"
    Private Const viewConsumoOF As String = "vDisponibilidadConsumoOF"
    Private Const viewEntradaOF As String = "vDisponibilidadEntradaOF"
    Private Const viewPendienteAlbaranCompra As String = "vDisponibilidadPendienteAlbaranCompra"
    Private Const viewPendienteAlbaranVentaDeposito As String = "vDisponibilidadPendienteAlbaranVentaDeposito"
    Private Const viewPendienteExpediciones As String = "vDisponibilidadPendienteExpediciones"
    Private Const viewPendienteObras As String = "vDisponibilidadPendienteObras"
    Private Const viewPendientePedidoVenta As String = "vDisponibilidadPendientePedidoVenta"
    Private Const viewPendienteRecibir As String = "vDisponibilidadPendienteRecibir"
    Private Const viewPendienteRecibirDeposito As String = "vDisponibilidadPendienteRecibirDeposito"

    Private Const viewPedidoVenta As String = "vDisponibilidadPedidoVenta"
    Private Const viewExpedicionObra As String = "vDisponibilidadExpedicionObra"
    Private Const viewPedidoCompra As String = "vDisponibilidadPedidoCompra"

    Private Const viewPendienteEnviarSolicTransferencia As String = "vDisponibilidadPendienteEnviarSolicTransferencia"
    Private Const viewPendienteRecibirSolicTransferencia As String = "vDisponibilidadPendienteRecibirSolicTransferencia"

    Private mValores As Hashtable
    Private mDataBases As Hashtable
    Private mPedidosGeneradosInfo As Hashtable

    <Serializable()> _
    Public Class DatosAnalisis
        Public PedidosPendientes As DataTable
        Public PreparadoEnTransporte As DataTable
        Public MaterialesPendientes As DataTable
        Public EntradasOF As DataTable
        Public ConsumosOF As DataTable
        Public AlbaranesCompraPendientes As DataTable
        Public AlbaranesVentaDeposito As DataTable
        Public AlbaranesPendientesExpedir As DataTable
        Public PedidosCompraPendientes As DataTable
        Public PedidosVentaDeposito As DataTable
        Public SolicTransferenciaPendienteEnviar As DataTable
        Public SolicTransferenciaPendienteRecibir As DataTable
    End Class

    <Serializable()> _
    Public Class DatosPedidoEmpresa
        Public BaseDatos As Guid
        Public IDDireccion As Integer
        Public PedidoCliente As String

        Public Sub New(ByVal BaseDatos As Guid)
            Me.BaseDatos = BaseDatos
        End Sub
        Public Sub New(ByVal BaseDatos As Guid, ByVal PedidoCliente As String)
            Me.BaseDatos = BaseDatos
            Me.PedidoCliente = PedidoCliente
        End Sub
        Public Sub New(ByVal BaseDatos As Guid, ByVal IDDireccion As Integer)
            Me.BaseDatos = BaseDatos
            Me.IDDireccion = IDDireccion
        End Sub
        Public Sub New(ByVal BaseDatos As Guid, ByVal IDDireccion As Integer, ByVal PedidoCliente As String)
            Me.BaseDatos = BaseDatos
            Me.IDDireccion = IDDireccion
            Me.PedidoCliente = PedidoCliente
        End Sub
    End Class

    <Task()> Public Shared Function NuevoDataTableAnalisis(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("IDArticulo", GetType(String))
        dt.Columns.Add("DescArticulo", GetType(String))
        dt.Columns.Add("StockFisico", GetType(Double)).DefaultValue = 0
        dt.Columns.Add("QPendiente", GetType(Double)).DefaultValue = 0
        dt.Columns.Add("QObras", GetType(Double)).DefaultValue = 0
        dt.Columns.Add("QFabricacion", GetType(Double)).DefaultValue = 0
        dt.Columns.Add("QAlbaranPdteActualizar", GetType(Double)).DefaultValue = 0
        dt.Columns.Add("QPendienteRecibir", GetType(Double)).DefaultValue = 0
        dt.Columns.Add("QPendienteEnviarSolicTransferencia", GetType(Double)).DefaultValue = 0
        dt.Columns.Add("QPendienteRecibirSolicTransferencia", GetType(Double)).DefaultValue = 0

        '//Campos calculados (TODAS LAS CANTIDADES LLEVAN SU SIGNO!!!) 
        Dim dc As DataColumn
        dc = dt.Columns.Add("QDisponible", GetType(Double))
        dc.DefaultValue = 0
        dc.Expression = "StockFisico+QPendiente+QObras+QFabricacion+QAlbaranPdteActualizar+QPendienteRecibir+QPendienteEnviarSolicTransferencia+QPendienteRecibirSolicTransferencia"
        Return dt
        '//
    End Function

    <Task()> Public Shared Function NuevoDataTableEvolucion(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim dt As New DataTable
        dt.Columns.Add("Fecha", GetType(Date))
        dt.Columns.Add("QPedidoVentaPendiente", GetType(Double)).DefaultValue = 0
        dt.Columns.Add("QObras", GetType(Double)).DefaultValue = 0
        dt.Columns.Add("QEntradaOf", GetType(Double)).DefaultValue = 0
        dt.Columns.Add("QConsumoOf", GetType(Double)).DefaultValue = 0
        dt.Columns.Add("QAlbaranCompraPendiente", GetType(Double)).DefaultValue = 0
        dt.Columns.Add("QAlbaranVentaDeposito", GetType(Double)).DefaultValue = 0
        dt.Columns.Add("QAlbaranVentaPendiente", GetType(Double)).DefaultValue = 0
        dt.Columns.Add("QPedidoCompraPendiente", GetType(Double)).DefaultValue = 0
        dt.Columns.Add("QPedidoVentaDeposito", GetType(Double)).DefaultValue = 0
        dt.Columns.Add("QDisponible", GetType(Double)).DefaultValue = 0
        dt.Columns.Add("QPendienteEnviarSolicTransferencia", GetType(Double)).DefaultValue = 0
        dt.Columns.Add("QPendienteRecibirSolicTransferencia", GetType(Double)).DefaultValue = 0
        Return dt
    End Function

    <Task()> Public Shared Function AnalisisPorArticulo(ByVal dispInfo As DisponibilidadInfo, ByVal services As ServiceProvider) As DataTable
        Dim dt As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf NuevoDataTableAnalisis, Nothing, services)

        Dim fecha As Date = dispInfo.FechaHasta
        If fecha = cnMinDate Then
            fecha = Today
        End If

        Dim f0 As New InnerFilter
        f0.Add("IDArticulo", FilterOperator.Equal, dispInfo.IDArticulo, FilterType.String)
        f0.Add("IDTipo", FilterOperator.Equal, dispInfo.IDTipo, FilterType.String)
        f0.Add("IDFamilia", FilterOperator.Equal, dispInfo.IDFamilia, FilterType.String)
        f0.Add("IDSubfamilia", FilterOperator.Equal, dispInfo.IDSubfamilia, FilterType.String)

        Dim orderby As String = "IDArticulo"
        Dim stocks As DataTable = New BE.DataEngine().Filter("vDisponibilidadArticuloAlmacen", f0, , orderby)
        If stocks.Rows.Count > 0 Then
            Dim IDArticulo As String
            Dim newrow As DataRow
            '//Stock de los almacenes
            For Each stock As DataRow In stocks.Rows
                If AreDifferents(IDArticulo, stock("IDArticulo")) Then
                    IDArticulo = stock("IDArticulo")
                    newrow = dt.NewRow
                    newrow("IDArticulo") = stock("IDArticulo")
                    newrow("DescArticulo") = stock("DescArticulo")
                    dt.Rows.Add(newrow)
                End If
                newrow("StockFisico") += stock("StockFisico")
            Next

            Dim StDataDatos As New DataDatos(f0, fecha, orderby)
            Dim datos As DatosAnalisis = ProcessServer.ExecuteTask(Of DataDatos, DatosAnalisis)(AddressOf Disponibilidad.Datos, StDataDatos, services)

            '//contabilizar
            '//Este dataview va a servir para ir completando el calculo
            Dim dv As New DataView(dt)
            dv.Sort = "IDArticulo"
            Dim StDataCont As New DataContArt("QPendiente", dv, "QPendiente", datos.PedidosPendientes, -1)
            ProcessServer.ExecuteTask(Of DataContArt)(AddressOf ContabilizarPorArticulo, StDataCont, services)
            StDataCont.key1 = "QObras" : StDataCont.data = datos.MaterialesPendientes
            ProcessServer.ExecuteTask(Of DataContArt)(AddressOf ContabilizarPorArticulo, StDataCont, services)
            StDataCont.key1 = "QFabricacion" : StDataCont.key2 = "QFabricar" : StDataCont.data = datos.EntradasOF : StDataCont.signo = 1
            ProcessServer.ExecuteTask(Of DataContArt)(AddressOf ContabilizarPorArticulo, StDataCont, services)
            StDataCont.key1 = "QFabricacion" : StDataCont.key2 = "QConsumida" : StDataCont.data = datos.ConsumosOF : StDataCont.signo = -1
            ProcessServer.ExecuteTask(Of DataContArt)(AddressOf ContabilizarPorArticulo, StDataCont, services)
            StDataCont.key1 = "QAlbaranPdteActualizar" : StDataCont.key2 = "QInterna" : StDataCont.data = datos.AlbaranesCompraPendientes : StDataCont.signo = 1
            ProcessServer.ExecuteTask(Of DataContArt)(AddressOf ContabilizarPorArticulo, StDataCont, services)
            StDataCont.data = datos.AlbaranesVentaDeposito : StDataCont.signo = 1
            ProcessServer.ExecuteTask(Of DataContArt)(AddressOf ContabilizarPorArticulo, StDataCont, services)
            StDataCont.data = datos.AlbaranesPendientesExpedir : StDataCont.signo = -1
            ProcessServer.ExecuteTask(Of DataContArt)(AddressOf ContabilizarPorArticulo, StDataCont, services)
            StDataCont.key1 = "QPendienteRecibir" : StDataCont.key2 = "QPendiente" : StDataCont.data = datos.PedidosCompraPendientes : StDataCont.signo = 1
            ProcessServer.ExecuteTask(Of DataContArt)(AddressOf ContabilizarPorArticulo, StDataCont, services)
            StDataCont.data = datos.PedidosVentaDeposito : StDataCont.signo = 1
            ProcessServer.ExecuteTask(Of DataContArt)(AddressOf ContabilizarPorArticulo, StDataCont, services)
            StDataCont.key1 = "QPendienteEnviarSolicTransferencia" : StDataCont.key2 = "CantidadSolicitada" : StDataCont.data = datos.SolicTransferenciaPendienteEnviar : StDataCont.signo = -1
            ProcessServer.ExecuteTask(Of DataContArt)(AddressOf ContabilizarPorArticulo, StDataCont, services)
            StDataCont.data = datos.SolicTransferenciaPendienteRecibir : StDataCont.signo = 1
            ProcessServer.ExecuteTask(Of DataContArt)(AddressOf ContabilizarPorArticulo, StDataCont, services)
        End If

        '//las condiciones sobre QDisponible y QPendiente se aplican al final
        If dt.Rows.Count > 0 AndAlso (IsNumeric(dispInfo.Disponible) Or IsNumeric(dispInfo.Pendiente)) Then
            Dim StData As New DataAplicarFiltro(dt, dispInfo)
            Return ProcessServer.ExecuteTask(Of DataAplicarFiltro, DataTable)(AddressOf AplicarFiltrosPendienteDisponible, StData, services)
        Else
            Return dt
        End If
    End Function

    <Task()> Public Shared Function AnalisisPorArticuloAlmacen(ByVal dispInfo As DisponibilidadporArticuloAlmacen, ByVal services As ServiceProvider) As DataTable
        Dim dt As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf NuevoDataTableAnalisis, Nothing, services)
        dt.Columns.Add("IDAlmacen", GetType(String))
        dt.Columns.Add("DescAlmacen", GetType(String))
        'Inicio
        dt.Columns.Add("PuntoPedido", GetType(Double))
        dt.Columns.Add("LoteMinimo", GetType(Double))
        dt.Columns.Add("StockSeguridad", GetType(Double))
        dt.Columns.Add("DifPtoPedido", GetType(Double))
        dt.Columns.Add("DifStockSeg", GetType(Double))
        'Fin


        Dim fecha As Date = dispInfo.FechaHasta
        If fecha = cnMinDate Then
            fecha = Today
        End If

        Dim f0 As New InnerFilter
        f0.Add("IDArticulo", FilterOperator.Equal, dispInfo.IDArticulo, FilterType.String)
        f0.Add("IDAlmacen", FilterOperator.Equal, dispInfo.IDAlmacen, FilterType.String)
        f0.Add("IDTipo", FilterOperator.Equal, dispInfo.IDTipo, FilterType.String)
        f0.Add("IDFamilia", FilterOperator.Equal, dispInfo.IDFamilia, FilterType.String)
        f0.Add("IDSubfamilia", FilterOperator.Equal, dispInfo.IDSubfamilia, FilterType.String)

        Dim orderby As String = "IDArticulo,IDAlmacen"
        Dim stocks As DataTable = New BE.DataEngine().Filter("vDisponibilidadArticuloAlmacen", f0, , orderby)
        If stocks.Rows.Count > 0 Then
            Dim IDArticulo, IDAlmacen As String
            Dim newrow As DataRow
            '//Stock de los almacenes
            For Each stock As DataRow In stocks.Rows
                If AreDifferents(IDArticulo, stock("IDArticulo")) Or AreDifferents(IDAlmacen, stock("IDAlmacen")) Then
                    IDArticulo = stock("IDArticulo")
                    IDAlmacen = stock("IDAlmacen")
                    newrow = dt.NewRow
                    newrow("IDArticulo") = stock("IDArticulo")
                    newrow("DescArticulo") = stock("DescArticulo")
                    newrow("IDAlmacen") = stock("IDAlmacen")
                    newrow("DescAlmacen") = stock("DescAlmacen")
                    'Inicio: Carga los datos cuando no son iguales el articulo actual y el siguiente de la DT
                    newrow("PuntoPedido") = stock("PuntoPedido")
                    newrow("LoteMinimo") = stock("LoteMinimo")
                    newrow("StockSeguridad") = stock("StockSeguridad")
                    'Fin
                    dt.Rows.Add(newrow)
                End If
                newrow("StockFisico") += stock("StockFisico")
            Next
            Dim StDataDatos As New DataDatos(f0, fecha, orderby)
            Dim datos As DatosAnalisis = ProcessServer.ExecuteTask(Of DataDatos, DatosAnalisis)(AddressOf Disponibilidad.Datos, StDataDatos, services)
            '//contabilizar
            '//Este dataview va a servir para ir completando el calculo
            Dim dv As New DataView(dt)
            dv.Sort = "IDArticulo,IDAlmacen"
            Dim StDataCont As New DataContArtAlm("QPendiente", dv, "QPendiente", datos.PedidosPendientes.DefaultView, -1)
            ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StDataCont, services)
            StDataCont.key1 = "QObras" : StDataCont.data = datos.MaterialesPendientes.DefaultView : StDataCont.signo = -1
            ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StDataCont, services)
            StDataCont.key1 = "QFabricacion" : StDataCont.key2 = "QFabricar" : StDataCont.data = datos.EntradasOF.DefaultView : StDataCont.signo = 1
            ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StDataCont, services)
            StDataCont.key2 = "QConsumida" : StDataCont.data = datos.ConsumosOF.DefaultView : StDataCont.signo = -1
            ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StDataCont, services)
            StDataCont.key1 = "QAlbaranPdteActualizar" : StDataCont.key2 = "QInterna" : StDataCont.data = datos.AlbaranesCompraPendientes.DefaultView : StDataCont.signo = 1
            ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StDataCont, services)
            StDataCont.data = datos.AlbaranesVentaDeposito.DefaultView
            ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StDataCont, services)
            StDataCont.data = datos.AlbaranesPendientesExpedir.DefaultView : StDataCont.signo = -1
            ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StDataCont, services)
            Dim filPedCompPen As New Filter(FilterUnionOperator.Or)
            Dim filSub As New Filter
            filSub.Add("TipoLineaCompra", FilterOperator.Equal, enumaclTipoLineaAlbaran.aclSubcontratacion)
            filSub.Add("IdOrdenRuta", FilterOperator.Equal, DBNull.Value)
            filPedCompPen.Add(filSub)
            filPedCompPen.Add("TipoLineaCompra", FilterOperator.NotEqual, enumaclTipoLineaAlbaran.aclSubcontratacion)
            datos.PedidosCompraPendientes.DefaultView.RowFilter = filPedCompPen.Compose(New AdoFilterComposer)
            StDataCont.key1 = "QPendienteRecibir" : StDataCont.key2 = "QPendiente" : StDataCont.data = datos.PedidosCompraPendientes.DefaultView : StDataCont.signo = 1
            ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StDataCont, services)
            StDataCont.data = datos.PedidosVentaDeposito.DefaultView : StDataCont.signo = 1
            ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StDataCont, services)
            StDataCont.key1 = "QPendienteEnviarSolicTransferencia" : StDataCont.key2 = "CantidadSolicitada" : StDataCont.data = datos.SolicTransferenciaPendienteEnviar.DefaultView : StDataCont.signo = -1
            ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StDataCont, services)
            StDataCont.key2 = "QPendienteRecibirSolicTransferencia" : StDataCont.data = datos.SolicTransferenciaPendienteRecibir.DefaultView : StDataCont.signo = 1
            ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StDataCont, services)
            'Inicio
            ProcessServer.ExecuteTask(Of DataView)(AddressOf CalcularDiferecias, dv, services)
            'Fin
        End If

        '//las condiciones sobre QDisponible y QPendiente se aplican al final
        If dt.Rows.Count > 0 AndAlso (IsNumeric(dispInfo.Disponible) Or IsNumeric(dispInfo.Pendiente)) Then
            Dim StDataDif As New DataAplicarFiltro(dt, dispInfo)
            dt = ProcessServer.ExecuteTask(Of DataAplicarFiltro, DataTable)(AddressOf AplicarFiltrosPendienteDisponible, StDataDif, services)
        End If
        '//las condiciones de diferencia se calculan al final
        If dt.Rows.Count > 0 AndAlso (IsNumeric(dispInfo.DifPtoPedido) Or IsNumeric(dispInfo.DifStockSeg)) Then
            Dim StDataDif As New DataAplicarFiltro(dt, dispInfo)
            dt = ProcessServer.ExecuteTask(Of DataAplicarFiltro, DataTable)(AddressOf AplicarFiltrosDiferencias, StDataDif, services)
        End If
        Return dt
    End Function

    <Serializable()> _
    Public Class DataDatos
        Public f As Filter
        Public fecha As Date
        Public orderby As String

        Public Sub New()
        End Sub
        Public Sub New(ByVal f As Filter, ByVal fecha As Date, ByVal orderby As String)
            Me.f = f
            Me.fecha = fecha
            Me.orderby = orderby
        End Sub
    End Class

    <Task()> Public Shared Function Datos(ByVal data As DataDatos, ByVal services As ServiceProvider) As DatosAnalisis
        Dim d As New DatosAnalisis
        '//QInterna de pedidos de venta pendientes o parc. servidos
        Dim StDataPend As New DataPendiente(data.f, data.fecha, data.orderby)
        d.PedidosPendientes = ProcessServer.ExecuteTask(Of DataPendiente, DataTable)(AddressOf PedidosPendientesServir, StDataPend, services)
        '//cantidades pendientes de servir desde ObraMaterial
        d.MaterialesPendientes = ProcessServer.ExecuteTask(Of DataPendiente, DataTable)(AddressOf PendienteExpedirObras, StDataPend, services)
        '//Balance de lo que va a entrar por fabricacion menos lo que se va a consumir en las ofs en curso
        d.EntradasOF = ProcessServer.ExecuteTask(Of DataPendiente, DataTable)(AddressOf PendienteFabricarOf, StDataPend, services)
        d.ConsumosOF = ProcessServer.ExecuteTask(Of DataPendiente, DataTable)(AddressOf PendienteConsumirOf, StDataPend, services)
        '//Balance de albaranes pendientes de actualizar + albaranes de deposito pendientes - albaranes de venta pendientes
        d.AlbaranesCompraPendientes = ProcessServer.ExecuteTask(Of DataPendiente, DataTable)(AddressOf AlbaranesCompraPendientes, StDataPend, services)
        d.AlbaranesVentaDeposito = ProcessServer.ExecuteTask(Of DataPendiente, DataTable)(AddressOf AlbaranesVentaDeposito, StDataPend, services)
        d.AlbaranesPendientesExpedir = ProcessServer.ExecuteTask(Of DataPendiente, DataTable)(AddressOf AlbaranesVentaPendientesExpedir, StDataPend, services)
        '//suma de las cantidades pendientes de recibir desde pedidos de compra y pendientes de recibir de pedidos de venta en almacenes de deposito de la empresa
        d.PedidosCompraPendientes = ProcessServer.ExecuteTask(Of DataPendiente, DataTable)(AddressOf PedidosCompraPendientesRecibir, StDataPend, services)
        d.PedidosVentaDeposito = ProcessServer.ExecuteTask(Of DataPendiente, DataTable)(AddressOf PedidosVentaDeposito, StDataPend, services)

        '//Solicitudes de Transferencia por Almacén Origen (desde el que sale la mercancía)
        Dim StDataSolic As New DatafSolicTrans(data.f, data.fecha, data.orderby)
        d.SolicTransferenciaPendienteEnviar = ProcessServer.ExecuteTask(Of DatafSolicTrans, DataTable)(AddressOf fSolicTransferenciaPendienteEnviar, StDataSolic, services)

        '//Solicitudes de Transferencia por Almacén Destino (el que recibe la mercancía)
        d.SolicTransferenciaPendienteRecibir = ProcessServer.ExecuteTask(Of DatafSolicTrans, DataTable)(AddressOf fSolicTransferenciaPendienteRecibir, StDataSolic, services)
        Return d
    End Function

    <Serializable()> _
    Public Class DatafSolicTrans
        Public f As Filter
        Public fecha As Date
        Public orderby As String

        Public Sub New()
        End Sub
        Public Sub New(ByVal f As Filter, ByVal fecha As Date, ByVal orderby As String)
            Me.f = f
            Me.fecha = fecha
            Me.orderby = orderby
        End Sub
    End Class

    <Task()> Public Shared Function fSolicTransferenciaPendienteEnviar(ByVal data As DatafSolicTrans, ByVal services As ServiceProvider) As DataTable
        '//Solicitudes de Transferencia por Almacén Origen (desde el que sale la mercancía)
        data.orderby = "IDArticulo, IDAlmacen, FechaPrevistaNecesidad, NSolicitud, IdSolicitudLinea"
        Dim f1 As New Filter
        f1.Add(data.f)
        f1.Add(New DateFilterItem("FechaPrevistaNecesidad", FilterOperator.LessThanOrEqual, data.fecha))
        Return New BE.DataEngine().Filter("vDisponibilidadPendienteEnviarSolicTransferencia", f1, , data.orderby)
    End Function

    <Task()> Public Shared Function fSolicTransferenciaPendienteRecibir(ByVal data As DatafSolicTrans, ByVal services As ServiceProvider) As DataTable
        '//Solicitudes de Transferencia por Almacén Destino (el que recibe la mercancía)
        data.orderby = "IDArticulo, IDAlmacen, FechaPrevistaNecesidad, NSolicitud, IdSolicitudLinea"
        Dim f1 As New Filter
        f1.Add(data.f)
        f1.Add(New DateFilterItem("FechaPrevistaNecesidad", FilterOperator.LessThanOrEqual, data.fecha))
        Return New BE.DataEngine().Filter("vDisponibilidadPendienteRecibirSolicTransferencia", f1, , data.orderby)
    End Function

    <Serializable()> _
    Public Class DataPendiente
        Public f As Filter
        Public fecha As Date
        Public orderby As String

        Public Sub New()
        End Sub
        Public Sub New(ByVal f As Filter, ByVal fecha As Date, ByVal orderby As String)
            Me.f = f
            Me.fecha = fecha
            Me.orderby = orderby
        End Sub
    End Class

    <Task()> Public Shared Function PedidosPendientesServir(ByVal data As DataPendiente, ByVal services As ServiceProvider) As DataTable
        '//QInterna de pedidos de venta pendientes o parc. servidos
        data.orderby = "IDArticulo, IDAlmacen, FechaEntrega, Prioridad, IDPedido, IDLineaPedido"
        Dim f1 As New Filter
        f1.Add(data.f)
        f1.Add(New DateFilterItem("FechaEntrega", FilterOperator.LessThanOrEqual, data.fecha))
        Return New BE.DataEngine().Filter("vDisponibilidadPendientePedidoVenta", f1, , data.orderby)
    End Function

    <Task()> Public Shared Function PendienteExpedirObras(ByVal data As DataPendiente, ByVal services As ServiceProvider) As DataTable
        '//cantidades pendientes de servir desde ObraMaterial
        data.orderby = "IDArticulo, IDAlmacen, FechaEntrega"
        Dim f1 As New Filter
        f1.Add(data.f)
        f1.Add(New DateFilterItem("FechaEntrega", FilterOperator.LessThanOrEqual, data.fecha))
        Return New BE.DataEngine().Filter("vDisponibilidadPendienteObras", f1, , data.orderby)
    End Function

    <Task()> Public Shared Function PendienteFabricarOf(ByVal data As DataPendiente, ByVal services As ServiceProvider) As DataTable
        '//Pendiente de fabricacion
        data.orderby = "IDArticulo, IDAlmacen, FechaFin, IDOrden"
        Dim f1 As New Filter
        f1.Add(data.f)
        f1.Add(New DateFilterItem("FechaFin", FilterOperator.LessThanOrEqual, data.fecha))
        Return New BE.DataEngine().Filter("vDisponibilidadEntradaOF", f1, , data.orderby)
    End Function

    <Task()> Public Shared Function PendienteConsumirOf(ByVal data As DataPendiente, ByVal services As ServiceProvider) As DataTable
        '//Pendiente de consumir
        data.orderby = "IDArticulo, IDAlmacen, FechaFin, IDOrden"
        Dim f1 As New Filter
        f1.Add(data.f)
        f1.Add(New DateFilterItem("FechaInicio", FilterOperator.LessThanOrEqual, data.fecha))
        Return New BE.DataEngine().Filter("vDisponibilidadConsumoOF", f1, , data.orderby)
    End Function

    <Task()> Public Shared Function AlbaranesCompraPendientes(ByVal data As DataPendiente, ByVal services As ServiceProvider) As DataTable
        '//Albaranes de compra pendientes de actualizar
        data.orderby = "IDArticulo, IDAlmacen, FechaAlbaran"
        Dim f1 As New Filter
        f1.Add(data.f)
        f1.Add(New DateFilterItem("FechaAlbaran", FilterOperator.LessThanOrEqual, data.fecha))
        Return New BE.DataEngine().Filter("vDisponibilidadPendienteAlbaranCompra", f1, , data.orderby)
    End Function

    <Task()> Public Shared Function AlbaranesVentaDeposito(ByVal data As DataPendiente, ByVal services As ServiceProvider) As DataTable
        '//Albaranes de deposito pendientes
        data.orderby = "IDArticulo, IDAlmacen, FechaAlbaran"
        Dim f1 As New Filter
        f1.Add(data.f)
        f1.Add(New DateFilterItem("FechaAlbaran", FilterOperator.LessThanOrEqual, data.fecha))
        Return New BE.DataEngine().Filter("vDisponibilidadPendienteAlbaranVentaDeposito", f1, , data.orderby)
    End Function

    <Task()> Public Shared Function AlbaranesVentaPendientesExpedir(ByVal data As DataPendiente, ByVal services As ServiceProvider) As DataTable
        '//Albaranes de venta pendientes de expedir
        data.orderby = "IDArticulo, IDAlmacen, FechaAlbaran"
        Dim f1 As New Filter
        f1.Add(data.f)
        f1.Add(New DateFilterItem("FechaAlbaran", FilterOperator.LessThanOrEqual, data.fecha))
        Return New BE.DataEngine().Filter("vDisponibilidadPendienteExpediciones", f1, , data.orderby)
    End Function

    <Task()> Public Shared Function PedidosCompraPendientesRecibir(ByVal data As DataPendiente, ByVal services As ServiceProvider) As DataTable
        '//Pendientes de recibir desde pedidos de compra
        data.orderby = "IDArticulo, FechaEntrega, IDAlmacen"
        Dim f1 As New Filter
        f1.Add(data.f)
        f1.Add(New DateFilterItem("FechaEntrega", FilterOperator.LessThanOrEqual, data.fecha))
        Return New BE.DataEngine().Filter("vDisponibilidadPendienteRecibir", f1, , data.orderby)
    End Function

    <Task()> Public Shared Function PedidosVentaDeposito(ByVal data As DataPendiente, ByVal services As ServiceProvider) As DataTable
        '//Pendientes de recibir de pedidos de venta en almacenes de deposito de la empresa
        Dim f1 As New Filter
        f1.Add(data.f)
        f1.Add(New DateFilterItem("FechaEntrega", FilterOperator.LessThanOrEqual, data.fecha))
        Return New BE.DataEngine().Filter("vDisponibilidadPendienteRecibirDeposito", f1, , data.orderby)
    End Function

    <Task()> Public Shared Function DatosFabricacion(ByVal data As DataPendiente, ByVal services As ServiceProvider) As DatosDisponibilidadOfs
        Dim dataDisp As New DatosDisponibilidadOfs
        dataDisp.EntradasOf = ProcessServer.ExecuteTask(Of DataPendiente, DataTable)(AddressOf PendienteFabricarOf, data, services)
        dataDisp.ConsumosOf = ProcessServer.ExecuteTask(Of DataPendiente, DataTable)(AddressOf PendienteConsumirOf, data, services)
        Return dataDisp
    End Function

    <Serializable()> _
    Public Class DataDispArtPedVta
        Public IDArticulo As String
        Public IDAlmacen As String
        Public Fecha As Date

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDArticulo As String, ByVal Fecha As Date, Optional ByVal IDAlmacen As String = Nothing)
            Me.IDArticulo = IDArticulo
            Me.Fecha = Fecha
            If Not IDAlmacen Is Nothing Then Me.IDAlmacen = IDAlmacen
        End Sub
    End Class

    <Task()> Public Shared Function DatosDisponibilidadArticuloPedidoVenta(ByVal data As DataDispArtPedVta, ByVal services As ServiceProvider) As DataSet
        '//Se utiliza desde la pantalla de disponibilidad de articulos, solapa de pedidos pendientes de servir.
        '//Obtiene toda la informacion para el grid de pedidos de venta:
        '//lineas de pedido pendientes + lineas de transporte
        Dim f As New Filter
        f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
        If Not data.IDAlmacen Is Nothing Then
            f.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
        End If
        f.Add(New DateFilterItem("FechaEntrega", FilterOperator.LessThanOrEqual, data.Fecha))
        Dim ds As New DataSet
        Dim pedidos As DataTable = New BE.DataEngine().Filter("vFrmMntoDispPedidoVentaPendiente", f, , "FechaPedido, FechaEntrega, IDLineaPedido, IDPedido")
        pedidos.TableName = "PedidoVentaLinea"
        ds.Tables.Add(pedidos)
        Return ds
    End Function

    <Task()> Public Shared Function DatosAlbaranes(ByVal data As DataPendiente, ByVal services As ServiceProvider) As DatosDisponibilidadAlbaranes
        Dim dataDisp As New DatosDisponibilidadAlbaranes
        dataDisp.AlbaranesCompraPendientes = ProcessServer.ExecuteTask(Of DataPendiente, DataTable)(AddressOf AlbaranesCompraPendientes, data, services)
        dataDisp.AlbaranesVentaDeposito = ProcessServer.ExecuteTask(Of DataPendiente, DataTable)(AddressOf AlbaranesVentaDeposito, data, services)
        dataDisp.AlbaranesPendientesExpedir = ProcessServer.ExecuteTask(Of DataPendiente, DataTable)(AddressOf AlbaranesVentaPendientesExpedir, data, services)
        Return dataDisp
    End Function

    <Serializable()> _
    Public Class DataContArt
        Public key1 As String
        Public dv As DataView
        Public key2 As String
        Public data As DataTable
        Public signo As Integer

        Public Sub New()
        End Sub
        Public Sub New(ByVal key1 As String, ByVal dv As DataView, ByVal key2 As String, ByVal data As DataTable, ByVal signo As Integer)
            Me.key1 = key1
            Me.dv = dv
            Me.key2 = key2
            Me.data = data
            Me.signo = signo
        End Sub
    End Class

    <Task()> Public Shared Sub ContabilizarPorArticulo(ByVal data As DataContArt, ByVal services As ServiceProvider)
        Dim index As Integer
        Dim IDArticulo As String
        Dim dr1 As DataRow
        For Each dr2 As DataRow In data.data.Rows
            If AreDifferents(IDArticulo, dr2("IDArticulo")) Then
                dr1 = Nothing
                IDArticulo = dr2("IDArticulo")
                index = data.dv.Find(IDArticulo)
                If index >= 0 Then
                    dr1 = data.dv(index).Row
                End If
            End If
            If Not dr1 Is Nothing Then
                dr1(data.key1) += Sign(data.signo) * dr2(data.key2)
            End If
        Next
    End Sub

    <Serializable()> _
    Public Class DataContArtAlm
        Public key1 As String
        Public dv As DataView
        Public key2 As String
        Public data As DataView
        Public signo As Integer

        Public Sub New()
        End Sub
        Public Sub New(ByVal key1 As String, ByVal dv As DataView, ByVal key2 As String, ByVal data As DataView, ByVal signo As Integer)
            Me.key1 = key1
            Me.dv = dv
            Me.key2 = key2
            Me.data = data
            Me.signo = signo
        End Sub
    End Class

    <Task()> Public Shared Sub ContabilizarPorArticuloAlmacen(ByVal data As DataContArtAlm, ByVal services As ServiceProvider)
        Dim index As Integer
        Dim IDArticulo, IDAlmacen As String
        Dim dr1 As DataRow
        For Each drv2 As DataRowView In data.data
            If AreDifferents(IDArticulo, drv2("IDArticulo")) Or AreDifferents(IDAlmacen, drv2("IDAlmacen")) Then
                IDArticulo = drv2("IDArticulo")
                IDAlmacen = drv2("IDAlmacen")
                index = data.dv.Find(New String() {IDArticulo, IDAlmacen})
                If index >= 0 Then
                    dr1 = data.dv(index).Row
                End If
            End If
            If Not dr1 Is Nothing Then
                dr1(data.key1) += Sign(data.signo) * drv2(data.key2)
            End If
        Next
    End Sub

    'Inicio: Cacula DifPtoPedido= Disponible - PtoPedido, DifStockSeg = QDisponible - StockSeguridad

    <Task()> Public Shared Sub CalcularDiferecias(ByVal dv As DataView, ByVal services As ServiceProvider)
        For Each drv As DataRowView In dv
            drv("DifPtoPedido") = Nz(drv("QDisponible"), 0) - Nz(drv("PuntoPedido"), 0)
            drv("DifStockSeg") = Nz(drv("QDisponible"), 0) - Nz(drv("StockSeguridad"), 0)
        Next
    End Sub
    'Fin
    <Serializable()> _
    Public Class DataAplicarFiltro
        Public Dt As DataTable
        Public DispInfo As DisponibilidadporArticuloAlmacen

        Public Sub New()
        End Sub
        Public Sub New(ByVal Dt As DataTable, ByVal DispInfo As DisponibilidadporArticuloAlmacen)
            Me.Dt = Dt
            Me.DispInfo = DispInfo
        End Sub
    End Class

    <Task()> Public Shared Function AplicarFiltrosPendienteDisponible(ByVal data As DataAplicarFiltro, ByVal services As ServiceProvider) As DataTable
        '//NOTA:QPendiente viene como cantidad positiva y en las tablas calculadas esta como negativa.
        Dim f As New Filter
        If IsNumeric(data.DispInfo.Disponible) And IsNumeric(data.DispInfo.Pendiente) Then
            f.Add(New NumberFilterItem("QDisponible", FilterOperator.LessThan, CDbl(data.DispInfo.Disponible)))
        ElseIf IsNumeric(data.DispInfo.Disponible) And Not IsNumeric(data.DispInfo.Pendiente) Then
            f.Add(New NumberFilterItem("QDisponible", FilterOperator.LessThan, CDbl(data.DispInfo.Disponible)))
        End If
        Dim auxdv As DataView = New DataView(data.Dt)
        auxdv.RowFilter = f.Compose(New AdoFilterComposer)

        Dim auxdt As DataTable = data.Dt.Clone
        For Each drv As DataRowView In auxdv
            auxdt.ImportRow(drv.Row)
        Next
        Return auxdt
    End Function

    'Inicio: Filtra los datos apartir de DifPtoPedido y DifStockSeg
    <Task()> Public Shared Function AplicarFiltrosDiferencias(ByVal data As DataAplicarFiltro, ByVal services As ServiceProvider) As DataTable
        '//NOTA:QPendiente viene como cantidad positiva y en las tablas calculadas esta como negativa.
        Dim f As New Filter
        If IsNumeric(data.DispInfo.DifPtoPedido) Then
            f.Add(New NumberFilterItem("DifPtoPedido", FilterOperator.LessThan, CDbl(data.DispInfo.DifPtoPedido)))
        End If

        If IsNumeric(data.DispInfo.DifStockSeg) Then
            f.Add(New NumberFilterItem("DifStockSeg", FilterOperator.LessThan, CDbl(data.DispInfo.DifStockSeg)))
        End If

        Dim auxdv As DataView = New DataView(data.Dt)
        auxdv.RowFilter = f.Compose(New AdoFilterComposer)

        Dim auxdt As DataTable = data.Dt.Clone
        For Each drv As DataRowView In auxdv
            auxdt.ImportRow(drv.Row)
        Next
        Return auxdt
    End Function

    <Task()> Public Shared Function AnalisisDesdePedidosDeVenta(ByVal dispInfo As DisponibilidadVentaInfo, ByVal services As ServiceProvider) As DataTable
        Dim fecha As Date = dispInfo.FechaHasta
        If fecha = cnMinDate Then
            fecha = Today
        End If

        Dim f As New Filter
        Dim f01 As New InnerFilter
        f01.Add("IDArticulo", FilterOperator.Equal, dispInfo.IDArticulo, FilterType.String)
        f01.Add("IDAlmacen", FilterOperator.Equal, dispInfo.IDAlmacen, FilterType.String)
        f01.Add("IDTipo", FilterOperator.Equal, dispInfo.IDTipo, FilterType.String)
        f01.Add("IDFamilia", FilterOperator.Equal, dispInfo.IDFamilia, FilterType.String)
        f01.Add("IDSubfamilia", FilterOperator.Equal, dispInfo.IDSubfamilia, FilterType.String)
        f01.Add("IDCliente", FilterOperator.Equal, dispInfo.IDCliente, FilterType.String)
        f01.Add("Estado", FilterOperator.Equal, dispInfo.Estado, FilterType.Numeric)
        f01.Add("Prioridad", FilterOperator.Equal, dispInfo.Prioridad, FilterType.Numeric)
        f01.Add("NPedido", FilterOperator.GreaterThanOrEqual, dispInfo.NPedidoDesde, FilterType.String)
        f01.Add("NPedido", FilterOperator.LessThanOrEqual, dispInfo.NPedidoHasta, FilterType.String)
        Dim f02 As New InnerFilter
        f02.Add("FechaEntrega", FilterOperator.LessThanOrEqual, fecha, FilterType.DateTime)
        f.Add(f01)
        f.Add(f02)

        Dim data As DataTable = New BE.DataEngine().Filter("vDisponibilidadPedidoVenta", f, , "IDArticulo, IDAlmacen, FechaEntrega, Prioridad, IDLineaPedido, IDPedido")
        Dim StDataCalcDisp As New DataCalcDisp(data, 1)
        data = ProcessServer.ExecuteTask(Of DataCalcDisp, DataTable)(AddressOf CalcularDisponibilidad, StDataCalcDisp, services)

        If IsNumeric(dispInfo.Disponible) Then
            Dim StDataFil As New DataFilQDisp(data, CDbl(dispInfo.Disponible))
            data = ProcessServer.ExecuteTask(Of DataFilQDisp, DataTable)(AddressOf FiltroQDisponible, StDataFil, services)
        End If
        Return data
    End Function

    <Task()> Public Shared Function AnalisisDesdeExpedicionesDeObra(ByVal dispInfo As DisponibilidadExpObraInfo, ByVal services As ServiceProvider) As DataTable
        Dim fecha As Date = dispInfo.FechaHasta
        If fecha = cnMinDate Then
            fecha = Today
        End If

        Dim f0 As New InnerFilter
        f0.Add("IDArticulo", FilterOperator.Equal, dispInfo.IDArticulo, FilterType.String)
        f0.Add("IDAlmacen", FilterOperator.Equal, dispInfo.IDAlmacen, FilterType.String)
        f0.Add("IDTipo", FilterOperator.Equal, dispInfo.IDTipo, FilterType.String)
        f0.Add("IDFamilia", FilterOperator.Equal, dispInfo.IDFamilia, FilterType.String)
        f0.Add("IDSubfamilia", FilterOperator.Equal, dispInfo.IDSubfamilia, FilterType.String)
        f0.Add("IDCliente", FilterOperator.Equal, dispInfo.IDCliente, FilterType.String)
        f0.Add("NObra", FilterOperator.Equal, dispInfo.NObra, FilterType.String)
        f0.Add("FechaEntrega", FilterOperator.LessThanOrEqual, fecha, FilterType.DateTime)

        Dim data As DataTable = New BE.DataEngine().Filter("vDisponibilidadExpedicionObra", f0, , "IDArticulo, IDAlmacen, FechaEntrega, IDObra")
        Dim StDataCalcDisp As New DataCalcDisp(data, 1)
        data = ProcessServer.ExecuteTask(Of DataCalcDisp, DataTable)(AddressOf CalcularDisponibilidad, StDataCalcDisp, services)
        If IsNumeric(dispInfo.Disponible) Then
            Dim StDataFilQDisp As New DataFilQDisp(data, CDbl(dispInfo.Disponible))
            data = ProcessServer.ExecuteTask(Of DataFilQDisp, DataTable)(AddressOf FiltroQDisponible, StDataFilQDisp, services)
        End If
        Return data
    End Function

    <Task()> Public Shared Function AnalisisDesdePedidosDeCompra(ByVal dispInfo As DisponibilidadCompraInfo, ByVal services As ServiceProvider) As DataTable
        Dim fecha As Date = dispInfo.FechaHasta
        If fecha = cnMinDate Then
            fecha = Today
        End If

        Dim f0 As New InnerFilter
        f0.Add("IDArticulo", FilterOperator.Equal, dispInfo.IDArticulo, FilterType.String)
        f0.Add("IDAlmacen", FilterOperator.Equal, dispInfo.IDAlmacen, FilterType.String)
        f0.Add("IDTipo", FilterOperator.Equal, dispInfo.IDTipo, FilterType.String)
        f0.Add("IDFamilia", FilterOperator.Equal, dispInfo.IDFamilia, FilterType.String)
        f0.Add("IDSubfamilia", FilterOperator.Equal, dispInfo.IDSubfamilia, FilterType.String)
        f0.Add("IDProveedor", FilterOperator.Equal, dispInfo.IDProveedor, FilterType.String)
        f0.Add("Estado", FilterOperator.Equal, dispInfo.Estado, FilterType.Numeric)
        f0.Add("FechaEntrega", FilterOperator.LessThanOrEqual, fecha, FilterType.DateTime)
        f0.Add("NPedido", FilterOperator.GreaterThanOrEqual, dispInfo.NPedidoDesde, FilterType.String)
        f0.Add("NPedido", FilterOperator.LessThanOrEqual, dispInfo.NPedidoHasta, FilterType.String)

        Dim data As DataTable = New BE.DataEngine().Filter("vDisponibilidadPedidoCompra", f0, , "IDArticulo, IDAlmacen, FechaEntrega, IDLineaPedido, IDPedido")
        Dim StDataCalcDisp As New DataCalcDisp(data, -1)
        data = ProcessServer.ExecuteTask(Of DataCalcDisp, DataTable)(AddressOf CalcularDisponibilidad, StDataCalcDisp, services)
        If IsNumeric(dispInfo.Disponible) Then
            Dim StDataFilQDisp As New DataFilQDisp(data, CDbl(dispInfo.Disponible))
            data = ProcessServer.ExecuteTask(Of DataFilQDisp, DataTable)(AddressOf FiltroQDisponible, StDataFilQDisp, services)
        End If
        Return data
    End Function

    <Serializable()> _
    Public Class DataCalcDisp
        Public Dt As DataTable
        Public Signo As Integer

        Public Sub New()
        End Sub
        Public Sub New(ByVal Dt As DataTable, ByVal signo As Integer)
            Me.Dt = Dt
            Me.Signo = signo
        End Sub
    End Class

    <Task()> Public Shared Function CalcularDisponibilidad(ByVal data As DataCalcDisp, ByVal services As ServiceProvider) As DataTable
        '//Esta funcion se la llama desde
        '//1.analisis desde pedidos de venta
        '//2.analisis desde pedidos de compra
        '//3.analisis desde obras pendientes de expedir
        '//Agrega y calcula los campos 'QDisponibleInicio' y 'QDisponibleFinal'.
        Dim QDisponible, QPendienteAcumulado As Double
        Dim IDArticulo, IDAlmacen As String
        Dim fecha As Date
        data.Dt.Columns.Add("QDisponibleInicio", GetType(Double))
        data.Dt.Columns.Add("QDisponibleFinal", GetType(Double))

        Dim dv1 As New DataView(data.Dt)
        If data.Dt.Rows.Count > 0 Then
            Dim analisis As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf NuevoDataTableAnalisis, Nothing, services)
            analisis.Columns.Add("IDAlmacen", GetType(String))
            Dim datos As DatosAnalisis
            Dim stocks As DataTable
            Dim f As New Filter
            For Each dr As DataRow In data.Dt.Rows
                If AreDifferents(IDArticulo, dr("IDArticulo")) Or AreDifferents(IDAlmacen, dr("IDAlmacen")) Then
                    datos = Nothing
                    QDisponible = 0
                    QPendienteAcumulado = 0
                    fecha = Date.MaxValue

                    IDArticulo = dr("IDArticulo")
                    IDAlmacen = dr("IDAlmacen")

                    f.Clear()
                    f.Add(New StringFilterItem("IDArticulo", IDArticulo))
                    f.Add(New StringFilterItem("IDAlmacen", IDAlmacen))

                    stocks = New ArticuloAlmacen().Filter(f)
                    If stocks.Rows.Count > 0 Then
                        Dim fechaEntregaMayor As Date
                        dv1.RowFilter = f.Compose(New AdoFilterComposer)
                        If dv1.Count > 0 Then
                            dv1.Sort = "FechaEntrega DESC"
                            fechaEntregaMayor = dv1(0)("FechaEntrega")
                            Dim StDataDatos As New DataDatos(f, fechaEntregaMayor, "IDArticulo,IDAlmacen")
                            datos = ProcessServer.ExecuteTask(Of DataDatos, DatosAnalisis)(AddressOf Disponibilidad.Datos, StDataDatos, services)
                        End If
                    End If
                End If

                If stocks.Rows.Count > 0 And Not IsNothing(datos) Then
                    If AreDifferents(fecha, dr("FechaEntrega")) Then
                        fecha = dr("FechaEntrega")

                        analisis.Clear()
                        Dim newrow As DataRow = analisis.NewRow
                        newrow("IDArticulo") = IDArticulo
                        newrow("IDAlmacen") = IDAlmacen
                        newrow("StockFisico") = stocks.Rows(0)("StockFisico")
                        analisis.Rows.Add(newrow)

                        Dim dv2 As New DataView(analisis)
                        dv2.Sort = "IDArticulo,IDAlmacen"

                        datos.PedidosPendientes.DefaultView.RowFilter = New DateFilterItem("FechaEntrega", FilterOperator.LessThanOrEqual, fecha).Compose(New AdoFilterComposer)
                        datos.MaterialesPendientes.DefaultView.RowFilter = New DateFilterItem("FechaEntrega", FilterOperator.LessThanOrEqual, fecha).Compose(New AdoFilterComposer)
                        datos.EntradasOF.DefaultView.RowFilter = New DateFilterItem("FechaInicio", FilterOperator.LessThanOrEqual, fecha).Compose(New AdoFilterComposer)
                        datos.ConsumosOF.DefaultView.RowFilter = New DateFilterItem("FechaInicio", FilterOperator.LessThanOrEqual, fecha).Compose(New AdoFilterComposer)
                        datos.AlbaranesCompraPendientes.DefaultView.RowFilter = New DateFilterItem("FechaAlbaran", FilterOperator.LessThanOrEqual, fecha).Compose(New AdoFilterComposer)
                        datos.AlbaranesVentaDeposito.DefaultView.RowFilter = New DateFilterItem("FechaAlbaran", FilterOperator.LessThanOrEqual, fecha).Compose(New AdoFilterComposer)
                        datos.AlbaranesPendientesExpedir.DefaultView.RowFilter = New DateFilterItem("FechaAlbaran", FilterOperator.LessThanOrEqual, fecha).Compose(New AdoFilterComposer)
                        datos.PedidosCompraPendientes.DefaultView.RowFilter = New DateFilterItem("FechaEntrega", FilterOperator.LessThanOrEqual, fecha).Compose(New AdoFilterComposer)
                        datos.PedidosVentaDeposito.DefaultView.RowFilter = New DateFilterItem("FechaEntrega", FilterOperator.LessThanOrEqual, fecha).Compose(New AdoFilterComposer)
                        datos.SolicTransferenciaPendienteEnviar.DefaultView.RowFilter = New DateFilterItem("FechaPrevistaNecesidad", FilterOperator.LessThanOrEqual, fecha).Compose(New AdoFilterComposer)
                        datos.SolicTransferenciaPendienteRecibir.DefaultView.RowFilter = New DateFilterItem("FechaPrevistaNecesidad", FilterOperator.LessThanOrEqual, fecha).Compose(New AdoFilterComposer)

                        Dim StContab As New DataContArtAlm("QPendiente", dv2, "QPendiente", datos.PedidosPendientes.DefaultView, -1)
                        ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StContab, services)
                        StContab = New DataContArtAlm("QObras", dv2, "QPendiente", datos.MaterialesPendientes.DefaultView, -1)
                        ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StContab, services)
                        StContab = New DataContArtAlm("QFabricacion", dv2, "QFabricar", datos.EntradasOF.DefaultView, 1)
                        ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StContab, services)
                        StContab = New DataContArtAlm("QFabricacion", dv2, "QConsumida", datos.ConsumosOF.DefaultView, -1)
                        ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StContab, services)
                        StContab = New DataContArtAlm("QAlbaranPdteActualizar", dv2, "QInterna", datos.AlbaranesCompraPendientes.DefaultView, 1)
                        ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StContab, services)
                        StContab = New DataContArtAlm("QAlbaranPdteActualizar", dv2, "QInterna", datos.AlbaranesVentaDeposito.DefaultView, 1)
                        ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StContab, services)
                        StContab = New DataContArtAlm("QAlbaranPdteActualizar", dv2, "QInterna", datos.AlbaranesPendientesExpedir.DefaultView, -1)
                        ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StContab, services)
                        StContab = New DataContArtAlm("QPendienteRecibir", dv2, "QPendiente", datos.PedidosCompraPendientes.DefaultView, 1)
                        ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StContab, services)
                        StContab = New DataContArtAlm("QPendienteRecibir", dv2, "QPendiente", datos.PedidosVentaDeposito.DefaultView, 1)
                        ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StContab, services)
                        StContab = New DataContArtAlm("QPendienteEnviarSolicTransferencia", dv2, "CantidadSolicitada", datos.SolicTransferenciaPendienteEnviar.DefaultView, -1)
                        ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StContab, services)
                        StContab = New DataContArtAlm("QPendienteRecibirSolicTransferencia", dv2, "CantidadSolicitada", datos.SolicTransferenciaPendienteRecibir.DefaultView, +1)
                        ProcessServer.ExecuteTask(Of DataContArtAlm)(AddressOf ContabilizarPorArticuloAlmacen, StContab, services)

                        '//La QDisponible calculada aqui es en realidad la QDisponible Final.
                        '//Es decir que tiene en cuenta todos los pedidos hasta la fecha de este pedido, incluido el propio pedido.
                        '//(Ademas tener en cuenta que puede haber mas de una linea de pedido en el mismo articulo y almacen para la misma fecha)
                        QDisponible = analisis.Rows(0)("QDisponible") + (Sign(data.Signo) * dr("QPendiente")) - QPendienteAcumulado

                    End If

                End If

                QPendienteAcumulado += dr("QPendiente")

                dr("QDisponibleInicio") = QDisponible
                dr("QDisponibleFinal") = dr("QDisponibleInicio") - (Sign(data.Signo) * dr("QPendiente"))
                QDisponible = dr("QDisponibleFinal")

            Next
        End If

        Return data.Dt

    End Function

    <Serializable()> _
    Public Class DataFilQDisp
        Public Dt As DataTable
        Public QDisponible As Double

        Public Sub New()
        End Sub
        Public Sub New(ByVal Dt As DataTable, ByVal QDisponible As Double)
            Me.Dt = Dt
            Me.QDisponible = QDisponible
        End Sub
    End Class

    <Task()> Public Shared Function FiltroQDisponible(ByVal data As DataFilQDisp, ByVal services As ServiceProvider) As DataTable
        '//Se aplica el filtro sobre el campo 'QDisponibleFinal'.
        If data.Dt.Rows.Count > 0 Then
            Dim f As New Filter
            f.Add(New NumberFilterItem("QDisponibleFinal", FilterOperator.LessThan, data.QDisponible))

            Dim auxdv As DataView = New DataView(data.Dt)
            auxdv.RowFilter = f.Compose(New AdoFilterComposer)

            Dim auxdt As DataTable = data.Dt.Clone
            For Each drv As DataRowView In auxdv
                auxdt.ImportRow(drv.Row)
            Next
            data.Dt = auxdt
        End If
        Return data.Dt
    End Function

    <Serializable()> _
    Public Class DataEvolDisp
        Public IDArticulo As String
        Public IDAlmacen As String
        Public Fecha As Date

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDArticulo As String, ByVal Fecha As Date, Optional ByVal IDAlmacen As String = Nothing)
            Me.IDArticulo = IDArticulo
            Me.Fecha = Fecha
            If Not IDAlmacen Is Nothing Then Me.IDAlmacen = IDAlmacen
        End Sub
    End Class

    <Task()> Public Shared Function EvolucionDisponibilidadPedidoVenta(ByVal data As DataEvolDisp, ByVal services As ServiceProvider) As EvolucionDisponibilidadInfo
        '//Esta funcion toma como base del desglose las fechas de entrega de los pedidos de venta       
        If Len(data.IDArticulo) = 0 Then
            ApplicationService.GenerateError("El Artículo es obligatorio")
        ElseIf data.Fecha = Date.MinValue Then
            ApplicationService.GenerateError("La fecha no es válida.")
        Else
            Dim resultado As New EvolucionDisponibilidadInfo
            resultado.IDArticulo = data.IDArticulo
            resultado.Fecha = data.Fecha
            If Not data.IDAlmacen Is Nothing Then
                resultado.IDAlmacen = data.IDAlmacen
            End If

            Dim dt As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf NuevoDataTableEvolucion, Nothing, services)

            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
            If Len(data.IDAlmacen) > 0 Then
                f.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
            End If

            Dim stocks As DataTable = New BE.DataEngine().Filter("vDisponibilidadArticuloAlmacen", f)
            If stocks.Rows.Count > 0 Then
                Dim stockFisico As Double
                '//Suma del stock de todos los almacenes
                For Each stock As DataRow In stocks.Rows
                    stockFisico += stock("StockFisico")
                Next

                resultado.StockFisico = stockFisico

                Dim StDataDatos As New DataDatos(f, data.Fecha, "IDArticulo")

                Dim datos As DatosAnalisis = ProcessServer.ExecuteTask(Of DataDatos, DatosAnalisis)(AddressOf Disponibilidad.Datos, StDataDatos, services)
                If datos.PedidosPendientes.Rows.Count > 0 Then
                    Dim f1, f2 As Date
                    Dim fechaEntrega As Date
                    Dim nr As DataRow

                    For Each pedidoVenta As DataRow In datos.PedidosPendientes.Rows
                        If AreDifferents(fechaEntrega, pedidoVenta("FechaEntrega")) Then
                            fechaEntrega = pedidoVenta("FechaEntrega")
                            If dt.Rows.Count > 0 Then
                                f2 = f1
                            End If
                            f1 = fechaEntrega

                            nr = dt.NewRow
                            'nr("IDArticulo") = stocks.Rows(0)("IDArticulo")
                            'nr("DescArticulo") = stocks.Rows(0)("DescArticulo")
                            nr("Fecha") = fechaEntrega
                            Dim StDataSum As New DataSumFechas("QPendiente", datos.MaterialesPendientes, "FechaEntrega", f1, f2)
                            nr("QObras") = -ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)
                            StDataSum = New DataSumFechas("QFabricar", datos.EntradasOF, "FechaInicio", f1, f2)
                            nr("QEntradaOf") = ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)
                            StDataSum = New DataSumFechas("QConsumida", datos.ConsumosOF, "FechaInicio", f1, f2)
                            nr("QConsumoOf") = -ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)
                            StDataSum = New DataSumFechas("QInterna", datos.AlbaranesCompraPendientes, "FechaAlbaran", f1, f2)
                            nr("QAlbaranCompraPendiente") = ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)
                            StDataSum = New DataSumFechas("QInterna", datos.AlbaranesVentaDeposito, "FechaAlbaran", f1, f2)
                            nr("QAlbaranVentaDeposito") = ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)
                            StDataSum = New DataSumFechas("QInterna", datos.AlbaranesPendientesExpedir, "FechaAlbaran", f1, f2)
                            nr("QAlbaranVentaPendiente") = -ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)
                            StDataSum = New DataSumFechas("QPendiente", datos.PedidosCompraPendientes, "FechaEntrega", f1, f2)
                            nr("QPedidoCompraPendiente") = ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)
                            StDataSum = New DataSumFechas("QPendiente", datos.PedidosVentaDeposito, "FechaEntrega", f1, f2)
                            nr("QPedidoVentaDeposito") = ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)

                            stockFisico += (nr("QObras") + nr("QEntradaOf") + nr("QConsumoOf") + nr("QAlbaranCompraPendiente") + nr("QAlbaranVentaDeposito") + nr("QAlbaranVentaPendiente") + nr("QPedidoCompraPendiente") + nr("QPedidoVentaDeposito"))

                            dt.Rows.Add(nr)
                        End If
                        nr("QPedidoVentaPendiente") -= pedidoVenta("QPendiente")
                        stockFisico -= pedidoVenta("QPendiente")
                        nr("QDisponible") = stockFisico
                    Next
                End If
            End If

            resultado.Datos = dt
            Return resultado
        End If
    End Function

    <Task()> Public Shared Function EvolucionDisponibilidadPedidoCompra(ByVal data As DataEvolDisp, ByVal services As ServiceProvider) As EvolucionDisponibilidadInfo
        '//Esta funcion toma como base del desglose las fechas de entrega de los pedidos de compra
        If Len(data.IDArticulo) = 0 Then
            ApplicationService.GenerateError("El Artículo es obligatorio")
        ElseIf data.Fecha = Date.MinValue Then
            ApplicationService.GenerateError("La fecha no es válida.")
        Else
            Dim resultado As New EvolucionDisponibilidadInfo
            resultado.IDArticulo = data.IDArticulo
            resultado.Fecha = data.Fecha
            If Not data.IDAlmacen Is Nothing Then
                resultado.IDAlmacen = data.IDAlmacen
            End If

            Dim dt As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf NuevoDataTableEvolucion, Nothing, services)

            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
            If Not data.IDAlmacen Is Nothing Then
                f.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
            End If

            Dim stocks As DataTable = New BE.DataEngine().Filter("vDisponibilidadArticuloAlmacen", f)
            If stocks.Rows.Count > 0 Then
                Dim stockFisico As Double
                '//Suma del stock de todos los almacenes
                For Each stock As DataRow In stocks.Rows
                    stockFisico += stock("StockFisico")
                Next

                resultado.StockFisico = stockFisico
                Dim StDataDatos As New DataDatos(f, data.Fecha, "IDArticulo")
                Dim datos As DatosAnalisis = ProcessServer.ExecuteTask(Of DataDatos, DatosAnalisis)(AddressOf Disponibilidad.Datos, StDataDatos, services)
                If datos.PedidosCompraPendientes.Rows.Count > 0 Then
                    Dim f1, f2 As Date
                    Dim fechaEntrega As Date
                    Dim nr As DataRow

                    For Each pedidoCompra As DataRow In datos.PedidosCompraPendientes.Rows
                        If AreDifferents(fechaEntrega, pedidoCompra("FechaEntrega")) Then
                            fechaEntrega = pedidoCompra("FechaEntrega")
                            If dt.Rows.Count > 0 Then
                                f2 = f1
                            End If
                            f1 = fechaEntrega

                            nr = dt.NewRow
                            'nr("IDArticulo") = stocks.Rows(0)("IDArticulo")
                            'nr("DescArticulo") = stocks.Rows(0)("DescArticulo")
                            nr("Fecha") = fechaEntrega

                            Dim StDataSum As New DataSumFechas("QPendiente", datos.PedidosPendientes, "FechaEntrega", f1, f2)
                            nr("QPedidoVentaPendiente") = -ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)
                            StDataSum = New DataSumFechas("QPendiente", datos.MaterialesPendientes, "FechaEntrega", f1, f2)
                            nr("QObras") = -ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)
                            StDataSum = New DataSumFechas("QFabricar", datos.EntradasOF, "FechaInicio", f1, f2)
                            nr("QEntradaOf") = ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)
                            StDataSum = New DataSumFechas("QConsumida", datos.ConsumosOF, "FechaInicio", f1, f2)
                            nr("QConsumoOf") = -ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)
                            StDataSum = New DataSumFechas("QInterna", datos.AlbaranesCompraPendientes, "FechaAlbaran", f1, f2)
                            nr("QAlbaranCompraPendiente") = ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)
                            StDataSum = New DataSumFechas("QInterna", datos.AlbaranesVentaDeposito, "FechaAlbaran", f1, f2)
                            nr("QAlbaranVentaDeposito") = ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)
                            StDataSum = New DataSumFechas("QInterna", datos.AlbaranesPendientesExpedir, "FechaAlbaran", f1, f2)
                            nr("QAlbaranVentaPendiente") = -ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)
                            StDataSum = New DataSumFechas("QPendiente", datos.PedidosVentaDeposito, "FechaEntrega", f1, f2)
                            nr("QPedidoVentaDeposito") = ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)

                            stockFisico += (nr("QPedidoVentaPendiente") + nr("QObras") + nr("QEntradaOf") + nr("QConsumoOf") + nr("QAlbaranCompraPendiente") + nr("QAlbaranVentaDeposito") + nr("QAlbaranVentaPendiente") + nr("QPedidoVentaDeposito"))

                            dt.Rows.Add(nr)
                        End If
                        nr("QPedidoCompraPendiente") += pedidoCompra("QPendiente")
                        stockFisico += pedidoCompra("QPendiente")
                        nr("QDisponible") = stockFisico
                    Next
                End If
            End If

            resultado.Datos = dt
            Return resultado
        End If
    End Function

    <Task()> Public Shared Function EvolucionDisponibilidadObras(ByVal data As DataEvolDisp, ByVal services As ServiceProvider) As EvolucionDisponibilidadInfo
        '//Esta funcion toma como base del desglose las fechas de entrega de los materiales de obra pendientes de expedir
        If Len(data.IDArticulo) = 0 Then
            ApplicationService.GenerateError("El Artículo es obligatorio")
        ElseIf data.Fecha = Date.MinValue Then
            ApplicationService.GenerateError("La fecha no es válida.")
        Else
            Dim resultado As New EvolucionDisponibilidadInfo
            resultado.IDArticulo = data.IDArticulo
            resultado.Fecha = data.Fecha
            If Not data.IDAlmacen Is Nothing Then
                resultado.IDAlmacen = data.IDAlmacen
            End If

            Dim dt As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf NuevoDataTableEvolucion, Nothing, services)

            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
            If Not data.IDAlmacen Is Nothing Then
                f.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
            End If

            Dim stocks As DataTable = New BE.DataEngine().Filter("vDisponibilidadArticuloAlmacen", f)
            If stocks.Rows.Count > 0 Then
                Dim stockFisico As Double
                '//Suma del stock de todos los almacenes
                For Each stock As DataRow In stocks.Rows
                    stockFisico += stock("StockFisico")
                Next

                resultado.StockFisico = stockFisico
                Dim StDataDatos As New DataDatos(f, data.Fecha, "IDArticulo")
                Dim datos As DatosAnalisis = ProcessServer.ExecuteTask(Of DataDatos, DatosAnalisis)(AddressOf Disponibilidad.Datos, StDataDatos, services)
                If datos.MaterialesPendientes.Rows.Count > 0 Then
                    Dim f1, f2 As Date
                    Dim fechaEntrega As Date
                    Dim nr As DataRow

                    For Each material As DataRow In datos.MaterialesPendientes.Rows
                        If AreDifferents(fechaEntrega, material("FechaEntrega")) Then
                            fechaEntrega = material("FechaEntrega")
                            If dt.Rows.Count > 0 Then
                                f2 = f1
                            End If
                            f1 = fechaEntrega

                            nr = dt.NewRow
                            'nr("IDArticulo") = stocks.Rows(0)("IDArticulo")
                            'nr("DescArticulo") = stocks.Rows(0)("DescArticulo")
                            nr("Fecha") = fechaEntrega

                            Dim StDataSum As New DataSumFechas("QPendiente", datos.PedidosPendientes, "FechaEntrega", f1, f2)
                            nr("QPedidoVentaPendiente") = -ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)
                            StDataSum = New DataSumFechas("QFabricar", datos.EntradasOF, "FechaInicio", f1, f2)
                            nr("QEntradaOf") = ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)
                            StDataSum = New DataSumFechas("QConsumida", datos.ConsumosOF, "FechaInicio", f1, f2)
                            nr("QConsumoOf") = -ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)
                            StDataSum = New DataSumFechas("QInterna", datos.AlbaranesCompraPendientes, "FechaAlbaran", f1, f2)
                            nr("QAlbaranCompraPendiente") = ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)
                            StDataSum = New DataSumFechas("QInterna", datos.AlbaranesVentaDeposito, "FechaAlbaran", f1, f2)
                            nr("QAlbaranVentaDeposito") = ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)
                            StDataSum = New DataSumFechas("QInterna", datos.AlbaranesPendientesExpedir, "FechaAlbaran", f1, f2)
                            nr("QAlbaranVentaPendiente") = -ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)
                            StDataSum = New DataSumFechas("QPendiente", datos.PedidosCompraPendientes, "FechaEntrega", f1, f2)
                            nr("QPedidoCompraPendiente") = ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)
                            StDataSum = New DataSumFechas("QPendiente", datos.PedidosVentaDeposito, "FechaEntrega", f1, f2)
                            nr("QPedidoVentaDeposito") = ProcessServer.ExecuteTask(Of DataSumFechas, Double)(AddressOf SumFechas, StDataSum, services)

                            stockFisico += (nr("QPedidoVentaPendiente") + nr("QEntradaOf") + nr("QConsumoOf") + nr("QAlbaranCompraPendiente") + nr("QAlbaranVentaDeposito") + nr("QAlbaranVentaPendiente") + nr("QPedidoCompraPendiente") + nr("QPedidoVentaDeposito"))

                            dt.Rows.Add(nr)
                        End If
                        nr("QObras") -= material("QPendiente")
                        stockFisico -= material("QPendiente")
                        nr("QDisponible") = stockFisico
                    Next
                End If
            End If

            resultado.Datos = dt
            Return resultado
        End If
    End Function

    <Task()> Public Shared Function EvolucionDisponibilidad(ByVal data As DataEvolDisp, ByVal services As ServiceProvider) As EvolucionDisponibilidadInfo
        If Len(data.IDArticulo) = 0 Then
            ApplicationService.GenerateError("El Artículo es obligatorio")
        ElseIf data.Fecha = Date.MinValue Then
            ApplicationService.GenerateError("La fecha no es válida.")
        Else
            Dim resultado As New EvolucionDisponibilidadInfo
            resultado.IDArticulo = data.IDArticulo
            resultado.Fecha = data.Fecha

            Dim dt As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf NuevoDataTableEvolucion, Nothing, services)
            dt.Columns.Add("QPendiente", GetType(Double)).DefaultValue = 0
            dt.Columns("QPendiente").Expression = "QPedidoVentaPendiente+QObras+QEntradaOf+QConsumoOf+QAlbaranCompraPendiente+QAlbaranVentaDeposito+QAlbaranVentaPendiente+QPedidoCompraPendiente+QPedidoVentaDeposito+QPendienteEnviarSolicTransferencia+QPendienteRecibirSolicTransferencia"

            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
            If Not data.IDAlmacen Is Nothing Then
                f.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
            End If

            Dim stocks As DataTable = New BE.DataEngine().Filter("vDisponibilidadArticuloAlmacen", f)
            If stocks.Rows.Count > 0 Then
                Dim stockFisico As Double
                '//Suma del stock de todos los almacenes
                For Each stock As DataRow In stocks.Rows
                    stockFisico += stock("StockFisico")
                Next

                resultado.StockFisico = stockFisico
                Dim StDataDatos As New DataDatos(f, data.Fecha, "IDArticulo")
                Dim datos As DatosAnalisis = ProcessServer.ExecuteTask(Of DataDatos, DatosAnalisis)(AddressOf Disponibilidad.Datos, StDataDatos, services)
                Dim StDatosSum As New DataSum(dt, "QPedidoVentaPendiente", datos.PedidosPendientes, "QPendiente", "FechaEntrega", -1)
                ProcessServer.ExecuteTask(Of DataSum)(AddressOf Sum, StDatosSum, services)
                StDatosSum = New DataSum(dt, "QObras", datos.MaterialesPendientes, "QPendiente", "FechaEntrega", -1)
                ProcessServer.ExecuteTask(Of DataSum)(AddressOf Sum, StDatosSum, services)
                StDatosSum = New DataSum(dt, "QEntradaOf", datos.EntradasOF, "QFabricar", "FechaInicio", +1)
                ProcessServer.ExecuteTask(Of DataSum)(AddressOf Sum, StDatosSum, services)
                StDatosSum = New DataSum(dt, "QConsumoOf", datos.ConsumosOF, "QConsumida", "FechaInicio", -1)
                ProcessServer.ExecuteTask(Of DataSum)(AddressOf Sum, StDatosSum, services)
                StDatosSum = New DataSum(dt, "QAlbaranCompraPendiente", datos.AlbaranesCompraPendientes, "QInterna", "FechaAlbaran", +1)
                ProcessServer.ExecuteTask(Of DataSum)(AddressOf Sum, StDatosSum, services)
                StDatosSum = New DataSum(dt, "QAlbaranVentaDeposito", datos.AlbaranesVentaDeposito, "QInterna", "FechaAlbaran", +1)
                ProcessServer.ExecuteTask(Of DataSum)(AddressOf Sum, StDatosSum, services)
                StDatosSum = New DataSum(dt, "QAlbaranVentaPendiente", datos.AlbaranesPendientesExpedir, "QInterna", "FechaAlbaran", -1)
                ProcessServer.ExecuteTask(Of DataSum)(AddressOf Sum, StDatosSum, services)
                StDatosSum = New DataSum(dt, "QPedidoCompraPendiente", datos.PedidosCompraPendientes, "QPendiente", "FechaEntrega", +1)
                ProcessServer.ExecuteTask(Of DataSum)(AddressOf Sum, StDatosSum, services)
                StDatosSum = New DataSum(dt, "QPedidoVentaDeposito", datos.PedidosVentaDeposito, "QPendiente", "FechaEntrega", +1)
                ProcessServer.ExecuteTask(Of DataSum)(AddressOf Sum, StDatosSum, services)

                StDatosSum = New DataSum(dt, "QPendienteEnviarSolicTransferencia", datos.SolicTransferenciaPendienteEnviar, "CantidadSolicitada", "FechaPrevistaNecesidad", -1)
                ProcessServer.ExecuteTask(Of DataSum)(AddressOf Sum, StDatosSum, services)
                StDatosSum = New DataSum(dt, "QPendienteRecibirSolicTransferencia", datos.SolicTransferenciaPendienteRecibir, "CantidadSolicitada", "FechaPrevistaNecesidad", +1)
                ProcessServer.ExecuteTask(Of DataSum)(AddressOf Sum, StDatosSum, services)

                For Each dr As DataRow In dt.Rows
                    stockFisico += dr("QPendiente")
                    dr("QDisponible") = stockFisico
                Next
            End If

            resultado.Datos = dt
            Return resultado
        End If
    End Function

    <Serializable()> _
    Public Class DataSum
        Public Resultado As DataTable
        Public CampoResultado As String
        Public Datos As DataTable
        Public CampoControl As String
        Public CampoFecha As String
        Public Signo As Integer

        Public Sub New()
        End Sub
        Public Sub New(ByVal Resultado As DataTable, ByVal CampoResultado As String, ByVal Datos As DataTable, ByVal CampoControl As String, ByVal CampoFecha As String, ByVal Signo As Integer)
            Me.Resultado = Resultado
            Me.CampoResultado = CampoResultado
            Me.Datos = Datos
            Me.CampoControl = CampoControl
            Me.CampoFecha = CampoFecha
            Me.Signo = Signo
        End Sub
    End Class

    <Task()> Public Shared Sub Sum(ByVal data As DataSum, ByVal services As ServiceProvider)
        If data.Datos.Rows.Count > 0 Then
            Dim match As Boolean
            Dim pos As Integer
            Dim row As DataRow
            For Each dr As DataRow In data.Datos.Select(Nothing, data.CampoFecha)
                pos = 0
                match = False
                For Each drv As DataRowView In data.Resultado.DefaultView
                    If drv("Fecha") = dr(data.CampoFecha) Then
                        match = True
                        row = drv.Row
                        Exit For
                    Else
                        If drv("Fecha") < dr(data.CampoFecha) Then
                            pos += 1
                        End If
                    End If
                Next
                If Not match Then
                    row = data.Resultado.NewRow
                    row("Fecha") = dr(data.CampoFecha)
                    data.Resultado.Rows.InsertAt(row, pos)
                End If
                row(data.CampoResultado) += Sign(data.Signo) * dr(data.CampoControl)
            Next
        End If
    End Sub

    <Serializable()> _
    Public Class DataSumFechas
        Public AcumulateField As String
        Public Data As DataTable
        Public DateField As String
        Public LessThanOrEqualDate As Date
        Public GreaterThanDate As Date

        Public Sub New()
        End Sub
        Public Sub New(ByVal AcumulateField As String, ByVal Data As DataTable, ByVal DateField As String, ByVal LessThanOrEqualDate As Date, ByVal GreaterThanDate As Date)
            Me.AcumulateField = AcumulateField
            Me.Data = Data
            Me.DateField = DateField
            Me.LessThanOrEqualDate = LessThanOrEqualDate
            Me.GreaterThanDate = GreaterThanDate
        End Sub
    End Class

    <Task()> Public Shared Function SumFechas(ByVal data As DataSumFechas, ByVal services As ServiceProvider) As Double
        If data.LessThanOrEqualDate > Date.MinValue Then
            Dim f As New Filter
            f.Add(New DateFilterItem(data.DateField, FilterOperator.LessThanOrEqual, data.LessThanOrEqualDate))
            If data.GreaterThanDate > Date.MinValue Then
                f.Add(New DateFilterItem(data.DateField, FilterOperator.GreaterThan, data.GreaterThanDate))
            End If
            data.Data.DefaultView.RowFilter = f.Compose(New AdoFilterComposer)
            For Each drv As DataRowView In data.Data.DefaultView
                SumFechas += drv(data.AcumulateField)
            Next
        End If
    End Function

    Public Function Agrupar(ByVal detalle As DataRowView, ByVal sort As String) As Boolean
        Dim dif As Boolean
        Dim keys() As String = Split(sort, ",")
        For Each key As String In keys
            If AreDifferents(mValores(key), detalle(key)) Then
                dif = True
                mValores(key) = detalle(key)
            End If
        Next
        Return dif
    End Function

    Public Function GetDataBase() As DataTable
        Dim dt As DataTable = New BE.DataEngine().Filter("xDataBase", "*", "", , , True)
        dt.DefaultView.Sort = "IDBaseDatos"
        Return dt
    End Function

    Public Function GetDataBase(ByVal IDBaseDatos As Guid) As DataRow
        Dim dt As DataTable = New BE.DataEngine().Filter("xDataBase", New GuidFilterItem("IdBaseDatos", IDBaseDatos), , , , True)
        If dt.Rows.Count > 0 Then
            Return dt.Rows(0)
        End If
    End Function

    Public Function GetDataBase(ByVal IDBaseDatos As Guid, ByVal sortedDatabases As DataView) As DataRow
        Dim i As Integer
        i = sortedDatabases.Find(IDBaseDatos)
        If i >= 0 Then
            Return sortedDatabases(i).Row
        End If
    End Function

    Public Function GetDataBaseDescription(ByVal IDBaseDatos As Guid, ByVal sortedDatabases As DataView) As String
        Dim dataBase As DataRow = GetDataBase(IDBaseDatos, sortedDatabases)
        If Not dataBase Is Nothing Then
            Return Nz(dataBase("DescBaseDatos"))
        End If
    End Function

    <Serializable()> _
    Public Class DataCalcNecMat
        Public IDArticulo As String
        Public IDAlmacen As String
        Public Fecha As Date

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDArticulo As String, ByVal Fecha As Date, Optional ByVal IDAlmacen As String = Nothing)
            Me.IDArticulo = IDArticulo
            Me.Fecha = Fecha
            If Not IDAlmacen Is Nothing Then Me.IDAlmacen = IDAlmacen
        End Sub
    End Class

    <Task()> Public Shared Function CalcularNecesidadesMateriasPrimasDesdePedidosVenta(ByVal data As DataCalcNecMat, ByVal Services As ServiceProvider) As DataTable
        Dim resultado As New DataTable
        resultado.Columns.Add("IDArticulo", GetType(String))
        resultado.Columns.Add("DescArticulo", GetType(String))
        resultado.Columns.Add("IDAlmacen", GetType(String))
        resultado.Columns.Add("DescAlmacen", GetType(String))
        resultado.Columns.Add("QStocks", GetType(Double)).DefaultValue = 0
        resultado.Columns.Add("QPedidos", GetType(Double)).DefaultValue = 0
        resultado.Columns.Add("QNecesidad", GetType(Double)).DefaultValue = 0
        resultado.Columns.Add("QDisponible", GetType(Double), "QPedidos+QStocks-QNecesidad").DefaultValue = 0

        Dim IDArticulo2, IDAlmacen2 As String
        Dim componentes As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf ArticuloEstructura.CalcularEstructuraExplosion, data.IDArticulo, Services)
        If Not componentes Is Nothing AndAlso componentes.Rows.Count > 0 Then
            Dim f As New Filter
            'Dim nr As DataRow
            For Each componente As DataRow In componentes.Select(Nothing, "IDComponente")
                If componente("Compra") Then
                    If Not AreEquals(IDArticulo2, componente("IDComponente")) Then
                        IDArticulo2 = componente("IDComponente")
                        IDAlmacen2 = Nothing

                        f.Clear()
                        f.Add(New StringFilterItem("IDArticulo", componente("IDComponente")))
                        If Not data.IDAlmacen Is Nothing Then
                            f.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
                        End If
                        'f.Add(New DateFilterItem("FechaEntrega", FilterOperator.LessThanOrEqual, Fecha))

                        Dim pedidos As DataTable = AdminData.GetData("vNegPedidosParaFabrica", f)
                        If Not pedidos Is Nothing AndAlso pedidos.Rows.Count > 0 Then
                            For Each pedido As DataRow In pedidos.Rows
                                Dim DrFind() As DataRow = resultado.Select("IDArticulo = '" & componente("IDComponente") & "' AND IDAlmacen = '" & pedido("IDAlmacen") & "'")
                                If DrFind.Length = 0 Then
                                    If Not AreEquals(IDAlmacen2, pedido("IDAlmacen")) Then
                                        IDAlmacen2 = pedido("IDAlmacen")

                                        Dim nr As DataRow = resultado.NewRow()
                                        nr("IDArticulo") = pedido("IDArticulo")
                                        nr("DescArticulo") = pedido("DescArticulo")
                                        nr("IDAlmacen") = pedido("IDAlmacen")
                                        nr("DescAlmacen") = pedido("DescAlmacen")
                                        nr("QStocks") = pedido("StockFisico")
                                        resultado.Rows.Add(nr)
                                    End If
                                    'resultado.Rows(resultado.Rows.Count - 1)("QPedidos") += pedido("Cantidad")
                                    'resultado.Rows(resultado.Rows.Count - 1)("QNecesidad") += pedido("Qnecesaria")

                                    'Pedidos de compra pendientes de recibir (aportan stock)
                                    Dim fPedidos As New Filter
                                    fPedidos.Add(New StringFilterItem("IDArticulo", pedido("IDArticulo")))
                                    fPedidos.Add(New StringFilterItem("IDAlmacen", pedido("IDAlmacen")))
                                    fPedidos.Add(New DateFilterItem("FechaEntrega", FilterOperator.LessThanOrEqual, data.Fecha))
                                    Dim PedPdtes As DataTable = AdminData.GetData("vNegPedidosCompraPendientes", fPedidos, "SUM(Cantidad) AS Cantidad")
                                    If Not PedPdtes Is Nothing AndAlso PedPdtes.Rows.Count > 0 Then
                                        resultado.Rows(resultado.Rows.Count - 1)("QPedidos") += Nz(PedPdtes.Rows(0)("Cantidad"), 0)
                                    End If

                                    'Ordenes de fabricación pendientes (consumen stock)
                                    Dim fOrdenes As New Filter
                                    fOrdenes.Add(New StringFilterItem("IDComponente", pedido("IDArticulo")))
                                    fOrdenes.Add(New StringFilterItem("IDAlmacen", pedido("IDAlmacen")))
                                    fOrdenes.Add(New DateFilterItem("FechaInicio", FilterOperator.LessThanOrEqual, data.Fecha))
                                    Dim OfsPdtes As DataTable = AdminData.GetData("vNegPedidosParaFabricaOFs", fOrdenes, "SUM(Qnecesaria) AS Qnecesaria")
                                    If Not OfsPdtes Is Nothing AndAlso OfsPdtes.Rows.Count > 0 Then
                                        resultado.Rows(resultado.Rows.Count - 1)("QNecesidad") += Nz(OfsPdtes.Rows(0)("Qnecesaria"), 0)
                                    End If

                                End If
                            Next
                        End If
                    End If
                End If
            Next
            Return resultado
        End If
    End Function

    <Serializable()> _
    Public Class DataCalcNecKit
        Public IDArticulo As String
        Public IDAlmacen As String
        Public Fecha As Date

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDArticulo As String, ByVal Fecha As Date, Optional ByVal IDAlmacen As String = Nothing)
            Me.IDArticulo = IDArticulo
            Me.Fecha = Fecha
            If Not IDAlmacen Is Nothing Then Me.IDAlmacen = IDAlmacen
        End Sub
    End Class

    <Task()> Public Shared Function CalcularDisponibilidadKitMP(ByVal data As DataCalcNecKit, ByVal services As ServiceProvider) As DataTable
        Dim resultado As New DataTable
        resultado.Columns.Add("IDComponente", GetType(String))
        resultado.Columns.Add("DescComponente", GetType(String))
        resultado.Columns.Add("IDAlmacen", GetType(String))
        resultado.Columns.Add("DescAlmacen", GetType(String))
        resultado.Columns.Add("StockFisico", GetType(Double)).DefaultValue = 0
        resultado.Columns.Add("QPdte", GetType(Double)).DefaultValue = 0
        resultado.Columns.Add("CantidadPorKit", GetType(Double)).DefaultValue = 0
        resultado.Columns.Add("StockKit", GetType(Double)).DefaultValue = 0

        Dim IDArticulo2 As String
        Dim f As New Filter
        f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
        If Not data.IDAlmacen Is Nothing Then
            f.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
        End If

        Dim componentes As DataTable = AdminData.GetData("vFrmDisponibilidadKitVenta", f)
        If Not componentes Is Nothing AndAlso componentes.Rows.Count > 0 Then

            For Each componente As DataRow In componentes.Select(Nothing, "IDComponente")

                Dim nr As DataRow = resultado.NewRow()
                nr("IDComponente") = componente("IDComponente")
                nr("DescComponente") = componente("DescComponente")
                nr("IDAlmacen") = componente("IDAlmacen")
                nr("DescAlmacen") = componente("DescAlmacen")
                nr("StockFisico") = componente("StockFisico")
                nr("CantidadPorKit") = componente("CantidadPorKit")
                nr("StockKit") = componente("StockKit")
                resultado.Rows.Add(nr)

                If Not AreEquals(IDArticulo2, componente("IDComponente")) Then
                    IDArticulo2 = componente("IDComponente")

                    f.Clear()
                    f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
                    f.Add(New StringFilterItem("IDComponente", componente("IDComponente")))
                    If Not data.IDAlmacen Is Nothing Then
                        f.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
                    End If
                    f.Add(New DateFilterItem("FechaEntrega", FilterOperator.LessThanOrEqual, data.Fecha))

                    Dim pedidos As DataTable = AdminData.GetData("vDisponibilidadPendienteRecibirKitMP", f, "SUM(Qpdte) AS Qpdte")
                    If Not pedidos Is Nothing AndAlso pedidos.Rows.Count > 0 Then
                        resultado.Rows(resultado.Rows.Count - 1)("Qpdte") += Nz(pedidos.Rows(0)("Qpdte"), 0)
                        resultado.Rows(resultado.Rows.Count - 1)("StockKit") += Nz(pedidos.Rows(0)("Qpdte"), 0) / resultado.Rows(resultado.Rows.Count - 1)("CantidadPorKit")
                    End If
                End If
            Next
            Return resultado
        End If
    End Function

#Region " Multiempresa "

    <Task()> Public Shared Function CrearPedidoVentaEnBDSecundaria(ByVal IDPedidoCompra() As Integer, ByVal services As ServiceProvider) As DataResultadoMultiempresaPC
        '//Utilizada desde: 
        '//1. El Mantenimiento de pedidos de compra 
        '//2. Disponibilidad de Pedidos de Compra
        Dim PedidosCompra As DataResultadoMultiempresaPC
        Dim Proveedores As EntityInfoCache(Of ProveedorInfo) = services.GetService(Of EntityInfoCache(Of ProveedorInfo))()

        For i As Integer = 0 To IDPedidoCompra.Length - 1
            ProcessServer.ExecuteTask(Of Object)(AddressOf Comunes.BeginTransaction, Nothing, services)
            Dim Doc As New DocumentoPedidoCompra(IDPedidoCompra(i))
            PedidosCompra = services.GetService(Of DataResultadoMultiempresaPC)()

            Dim ProvInfo As ProveedorInfo = Proveedores.GetEntity(Doc.HeaderRow("IDProveedor"))

            Dim PCInfoNew As New GeneracionPedidosCompraInfo
            PCInfoNew.Proveedor = String.Concat(ProvInfo.IDProveedor, " - ", ProvInfo.DescProveedor)
            PCInfoNew.EmpresaGrupo = ProvInfo.EmpresaGrupo
            PCInfoNew.EntregaProveedor = Doc.HeaderRow("EntregaProveedor")
            PCInfoNew.IDPedidoCompra = Doc.HeaderRow("IDPedido")
            PCInfoNew.NPedidoCompra = Doc.HeaderRow("NPedido")
            PCInfoNew.BaseDatos1 = AdminData.GetSessionInfo.DataBase.DataBaseDescription
            PedidosCompra.Add(PCInfoNew)

            Dim dat As New DataPrcCrearPedidoVentaEnBDSecundaria(AdminData.GetSessionInfo.DataBase.DataBaseID, Doc)
            ProcessServer.RunProcess(GetType(PrcCrearPedidoVentaEnBDSecundaria), dat, services)

            PedidosCompra = services.GetService(Of DataResultadoMultiempresaPC)()
            If Length(PedidosCompra.Item(Doc.HeaderRow("IDPedido")).StrError) = 0 Then
                ProcessServer.ExecuteTask(Of DocumentoPedidoCompra)(AddressOf Comunes.UpdateDocument, Doc, services)
            End If
            ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.CommitTransaction, True, services)
        Next
        Return PedidosCompra
    End Function

#End Region

End Class

Friend Class InnerFilter
    Inherits Expertis.Engine.Filter

    Public Sub New(Optional ByVal unionOperator As FilterUnionOperator = FilterUnionOperator.And)
        MyBase.New(unionOperator)
    End Sub

    Public Overloads Overrides Sub Add(ByVal strFieldName As String, ByVal intOperator As Engine.FilterOperator, ByVal oValue As Object)
        If IsValidValue(oValue) Then
            MyBase.Add(strFieldName, intOperator, oValue)
        End If
    End Sub

    Public Overloads Overrides Sub Add(ByVal strFieldName As String, ByVal intOperator As Engine.FilterOperator, ByVal oValue As Object, ByVal intDataType As Engine.FilterType)
        If IsValidValue(oValue) Then
            MyBase.Add(strFieldName, intOperator, oValue, intDataType)
        End If
    End Sub

    Public Overloads Overrides Sub Add(ByVal strFieldName As String, ByVal intOperator As Engine.FilterOperator, ByVal oValue As Object, ByVal SystemType As System.Type)
        If IsValidValue(oValue) Then
            MyBase.Add(strFieldName, intOperator, oValue, SystemType)
        End If
    End Sub

    Public Overloads Overrides Sub Add(ByVal strFieldName As String, ByVal oValue As Object)
        If IsValidValue(oValue) Then
            MyBase.Add(strFieldName, oValue)
        End If
    End Sub

    Private Function IsValidValue(ByVal Value As Object) As Boolean
        If TypeOf Value Is String Then
            Return Value <> String.Empty
        Else
            Return Not Value Is Nothing AndAlso Not Value Is System.DBNull.Value
        End If
    End Function

End Class