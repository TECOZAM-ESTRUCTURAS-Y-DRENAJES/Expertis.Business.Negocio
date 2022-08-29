Imports System.Collections.Generic


'Public Class PRUEBAS_STK
'    Public Const mblnMejoraTiempos As Boolean = True
'End Class

Public Class _HistoricoMovimiento
    Public Const IDLineaMovimiento As String = "IDLineaMovimiento"
    Public Const IDMovimiento As String = "IDMovimiento"
    Public Const IDArticulo As String = "IDArticulo"
    Public Const IDTipoMovimiento As String = "IDTipoMovimiento"
    Public Const IDAlmacen As String = "IDAlmacen"
    Public Const Cantidad As String = "Cantidad"
    Public Const PrecioA As String = "PrecioA"
    Public Const PrecioB As String = "PrecioB"
    Public Const Texto As String = "Texto"
    Public Const Lote As String = "Lote"
    Public Const Ubicacion As String = "Ubicacion"
    Public Const Acumulado As String = "Acumulado"
    Public Const FechaMovimiento As String = "FechaMovimiento"
    Public Const Documento As String = "Documento"
    Public Const FechaDocumento As String = "FechaDocumento"
    Public Const IDObra As String = "IDObra"
    Public Const FechaCreacionAudi As String = "FechaCreacionAudi"
    Public Const FechaModificacionAudi As String = "FechaModificacionAudi"
    Public Const UsuarioAudi As String = "UsuarioAudi"
    Public Const IDLineaMaterial As String = "IDLineaMaterial"
    Public Const IDActivo As String = "IDActivo"
    Public Const Traza As String = "Traza"
    Public Const IDDocumento As String = "IDDocumento"
    Public Const IDEstadoActivo As String = "IDEstadoActivo"
    Public Const ClaseMovimiento As String = "ClaseMovimiento"
End Class

#Region " Clases de Articulo - Almacén "

<Serializable()> _
Public Class DataArticuloAlmacen
    Private mIDArticulo As String
    Private mIDAlmacen As String

    Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String)
        Me.mIDArticulo = IDArticulo
        Me.mIDAlmacen = IDAlmacen
    End Sub

    Public Property IDArticulo() As String
        Get
            Return mIDArticulo
        End Get
        Set(ByVal value As String)
            mIDArticulo = value
        End Set
    End Property

    Public Property IDAlmacen() As String
        Get
            Return mIDAlmacen
        End Get
        Set(ByVal value As String)
            mIDAlmacen = value
        End Set
    End Property
End Class

<Serializable()> _
Public Class DataArticuloAlmacenFecha
    Inherits DataArticuloAlmacen

    Private mIDArticulo As String
    Private mIDAlmacen As String
    Private mFecha As Date

    Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal Fecha As Date)
        MyBase.New(IDArticulo, IDAlmacen)
        Me.mFecha = Fecha
    End Sub

    Public Property Fecha() As Date
        Get
            Return mFecha
        End Get
        Set(ByVal value As Date)
            mFecha = value
        End Set
    End Property

End Class

#End Region

#Region "Control de la produccion"

Public Interface IControlProduccion
    Function ControlProduccion(ByVal parteTrabajo As DataRow) As ControlProduccionUpdateData
    Function ControlProduccion(ByVal parteTrabajo As DataTable) As ControlProduccionUpdateData
    Sub EliminarOFControl(ByVal IDOFControl As Integer)
End Interface

'//Esta clases se utilizan desde Business.Negocio y Business.Produccion
<Serializable()> _
Public Class ControlProduccionUpdateData
    Public OrdenFabricacion As DataTable
    Public OrdenesFinales As DataTable
    Public OFControl As DataTable
    Public OFControlEstructura As DataTable
    Public Entradas() As StockUpdateData
    Public Salidas() As StockUpdateData
End Class
#End Region

#Region " Clases genéricas utilizadas en Stocks "

Public Enum EstadoStock
    NoActualizado = 0
    Actualizado = 1
    SinGestion = 2
End Enum

<Serializable()> _
Public Class ValoracionPreciosInfo
    Public IDArticulo As String
    Public IDAlmacen As String
    Public FechaCalculo As Date
    Public CriterioValoracion As BusinessEnum.enumtaValoracion
    Public PrecioA As Double
    Public PrecioB As Double
End Class

<Serializable()> _
Public Class StockAFechaInfo
    Public IDArticulo As String
    Public IDAlmacen As String
    Public FechaCalculo As Date
    Public StockAFecha As Double
    Public StockAFecha2 As Double
    Public StockAFechaUdValoracion As Double
    Public PrecioMedio As Double
    Public PrecioEstandar As Double
    Public PrecioUltimaCompra As Double
    Public FifoF As Double
    Public FifoFD As Double
End Class

<Serializable()> _
Public Class ValoracionInfo
    Public IDArticulo As String
    Public IDAlmacen As String
    Public FechaCalculo As Date
    Public Stock As New StockAFechaInfo
    Public Precios As New ValoracionPreciosInfo
End Class

<Serializable()> _
  Public Class DataTratarStocks
    Public Items() As StockData
    Public Sinc As Boolean

    Public Sub New(ByVal Items() As StockData, Optional ByVal Sinc As Boolean = True)
        Me.Items = Items
        Me.Sinc = Sinc
    End Sub
End Class

Public Class DataMovimiento
    Public Movimiento As DataRow
    Public stkData As StockData

    Public Sub New(ByVal Movimiento As DataRow, ByVal stkData As StockData)
        If Not Movimiento Is Nothing Then Me.Movimiento = Movimiento
        Me.stkData = stkData
    End Sub
End Class

Public Class DataMovimientoSinc
    Inherits DataMovimiento

    Public Sinc As Boolean

    Public Sub New(ByVal Movimiento As DataRow, ByVal stkData As StockData, Optional ByVal sinc As Boolean = True)
        MyBase.New(Movimiento, stkData)
        Me.Sinc = sinc
    End Sub
End Class

<Serializable()> _
Public Class DataNumeroMovimiento
    Public NumeroMovimiento As Integer
    Public stkData As StockData

    Public Sub New(ByVal NumeroMovimiento As Integer, ByVal stkData As StockData)
        Me.NumeroMovimiento = NumeroMovimiento
        Me.stkData = stkData
    End Sub
End Class

<Serializable()> _
Public Class DataNumeroMovimientoSinc
    Inherits DataNumeroMovimiento

    Public Sinc As Boolean

    Public Sub New(ByVal NumeroMovimiento As Integer, ByVal stkData As StockData, Optional ByVal sinc As Boolean = True)
        MyBase.New(NumeroMovimiento, stkData)
        Me.Sinc = sinc
    End Sub
End Class

#Region "Contexto del proceso de actualizacion de stock"

<Serializable()> _
Public Class StockData
    Public Articulo As String
    Public Almacen As String
    Public Cantidad As Double
    Public Cantidad2 As Double?
    Public PrecioA As Double
    Public PrecioB As Double
    Public Lote As String
    Public Ubicacion As String
    Public FechaCaducidad As Date
    Public FechaDocumento As Date
    Public Documento As String
    Public Texto As String
    Public TipoMovimiento As enumTipoMovimiento
    Public Obra As Integer
    Public NSerie As String
    Public Activo As String
    Public EstadoNSerie As String
    Public EstadoNSerieAnterior As String
    Public Operario As String
    Public Traza As Guid
    Public Enlace As Integer
    Public IDDocumento As Integer
    Public Context As StockContext
    Public ContextCorrect As StockCorrectContext

    'PRECINTAS
    Public PrecintaNSerie As String
    Public PrecintaDesde As Integer
    Public PrecintaUtilizadaDesde As Integer
    Public PrecintaHasta As Integer
    Public PrecintaUtilizadaHasta As Integer

    Public IDOperacionPlan As String

    Public Sub New()
        Me.Context = New StockContext()
        Me.ContextCorrect = New StockCorrectContext(False, False, False, False, False)
    End Sub
    Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal Cantidad As Double, ByVal PrecioA As Double, ByVal PrecioB As Double, ByVal FechaDocumento As Date, ByVal TipoMovimiento As enumTipoMovimiento)
        Me.Articulo = IDArticulo
        Me.Almacen = IDAlmacen
        Me.Cantidad = Cantidad
        Me.PrecioA = PrecioA
        Me.PrecioB = PrecioB
        Me.FechaDocumento = FechaDocumento
        Me.TipoMovimiento = TipoMovimiento
        Me.Documento = Documento
        Me.Context = New StockContext(IDArticulo, IDAlmacen, FechaDocumento)
        Me.ContextCorrect = New StockCorrectContext(False, False, False, False, False)
    End Sub

    Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal Cantidad As Double, ByVal PrecioA As Double, ByVal PrecioB As Double, ByVal FechaDocumento As Date, ByVal TipoMovimiento As enumTipoMovimiento, ByVal Documento As String, Optional ByVal IDDocumento As Integer = -1)
        Me.New(IDArticulo, IDAlmacen, Cantidad, PrecioA, PrecioB, FechaDocumento, TipoMovimiento)
        Me.Documento = Documento
        If IDDocumento <> -1 Then Me.IDDocumento = IDDocumento
    End Sub

    '//para la gestion de lotes
    Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal Lote As String, ByVal Ubicacion As String, ByVal Cantidad As Double, ByVal PrecioA As Double, ByVal PrecioB As Double, ByVal FechaDocumento As Date, ByVal TipoMovimiento As enumTipoMovimiento)
        Me.New(IDArticulo, IDAlmacen, Cantidad, PrecioA, PrecioB, FechaDocumento, TipoMovimiento)
        Me.Lote = Lote
        Me.Ubicacion = Ubicacion
    End Sub

    Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal Lote As String, ByVal Ubicacion As String, ByVal Cantidad As Double, ByVal PrecioA As Double, ByVal PrecioB As Double, ByVal FechaDocumento As Date, ByVal TipoMovimiento As enumTipoMovimiento, ByVal Documento As String, Optional ByVal IDDocumento As Integer = -1)
        Me.New(IDArticulo, IDAlmacen, Lote, Ubicacion, Cantidad, PrecioA, PrecioB, FechaDocumento, TipoMovimiento)
        Me.Documento = Documento
        If IDDocumento <> -1 Then Me.IDDocumento = IDDocumento
    End Sub

    '//para la gestion de numeros de serie
    Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal NSerie As String, ByVal Activo As String, ByVal Estado As String, ByVal Operario As String, ByVal FechaDocumento As Date, ByVal TipoMovimiento As enumTipoMovimiento, optional ByVal PrecioA As Double=0, optional ByVal PrecioB As Double= 0)
        Me.New(IDArticulo, IDAlmacen, IIf(TipoMovimiento = enumTipoMovimiento.tmSalAjuste, -1, 1), PrecioA, PrecioB, FechaDocumento, TipoMovimiento)
        Me.NSerie = NSerie
        Me.Activo = Activo
        Me.EstadoNSerie = Estado
        Me.Operario = Operario
    End Sub

    Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal NSerie As String, ByVal Activo As String, ByVal Estado As String, ByVal Operario As String, ByVal FechaDocumento As Date, ByVal TipoMovimiento As enumTipoMovimiento, ByVal Documento As String, Optional ByVal IDDocumento As Integer = -1, Optional ByVal PrecioA As Double = 0, Optional ByVal PrecioB As Double = 0)
        Me.New(IDArticulo, IDAlmacen, NSerie, Activo, Estado, Operario, FechaDocumento, TipoMovimiento, PrecioA, PrecioB)
        Me.Documento = Documento
        If IDDocumento <> -1 Then Me.IDDocumento = IDDocumento
    End Sub
End Class

<Serializable()> _
Public Class StockContext   '// Es Serializable por que utiliza en la clase StockData
    Public Articulo As String
    Public Almacen As String
    Public FechaDocumento As Date

    Public StockFisico As Double
    Public StockFisico2 As Double?
    Public FechaUltimoInventario As Date
    Public FechaUltimoMovimiento As Date
    Public ArticuloGenerico As String

    Public Lote As String
    Public Ubicacion As String
    Public StockFisicoLote As Double
    Public StockFisicoLote2 As Double?
    Public LoteBloqueado As Boolean
    Public Traza As Guid

    Public NSerie As String
    Public Activo As String
    Public EstadoNSerie As String
    Public IDEstadoActivoAnterior As String            '// Entrada, Salida
    Public Operario As String
    Public PropiedadesEstado As New PropiedadesEstadoNSerie
    Public PropiedadesEstadoBBDD As New PropiedadesEstadoNSerie

    Public NumeroMovimiento As Integer
    Public Documento As String
    Public DocumentoOriginal As String
    Public IDDocumento As Integer
    Public TipoMovimiento As enumTipoMovimiento         '// EntradaAjuste, EntradaAlbaran,.....
    Public ClaseMovimiento As enumtpmTipoMovimiento     '// Entrada, Salida, Inventario
    Public FechaUltimoCierre As Date
    Public CantidadConSigno As Double
    Public CantidadConSigno2 As Double?
    Public PrecioA As Double
    Public PrecioB As Double

    Public ArticuloAlmacen As DataTable
    Public LoteBBDD As DataTable
    Public SerieBBDD As DataTable
    Public ActivoBBDD As DataTable
    Public HistoricoEstadoActivo As DataTable

    Public Movimientos As DataTable
    Public Acumulados As DataTable
    Public QIntermedia As Double                        '// Inventarios
    Public QIntermedia2 As Double                       '// Inventarios
    Public QAjuste As Double                            '// Ajustes
    Public QAjuste2 As Double                           '// Ajustes
    Public QIntermediaRestoLotes As Double              '// Inventarios - Lotes
    Public QIntermediaRestoLotes2 As Double             '// Inventarios - Lotes

    Public EsCorreccion As Boolean
    Public GenerarMovimiento As Boolean

    Public Cancel As Boolean
    Public Estado As EstadoStock
    Public IDLineaMovimiento As Integer
    Public Log As String
    Public mDetalle As String
    Public Property Detalle() As String
        Get
            Return mDetalle
        End Get
        Set(ByVal Value As String)
            mDetalle = Value
            If Me.Cancel Then
                Dim services As New ServiceProvider
                Dim msg As String = ProcessServer.ExecuteTask(Of ProcesoStocks.DataMessage, String)(AddressOf ProcesoStocks.Message, New ProcesoStocks.DataMessage(0, Me.Articulo, Me.Almacen), services)
                Me.Log = String.Concat(msg, ControlChars.NewLine, mDetalle)
            Else
                Me.Log = mDetalle
            End If
        End Set
    End Property

    Public Obra As String

    Public CantidadOriginal As Double
    Public CantidadOriginal2 As Double?
    Public PrecioAOriginal As Double
    Public PrecioBOriginal As Double
    Public FechaDocumentoOriginal As Date
    Public CorreccionEnDocumento As Boolean

    Public Sub New()

    End Sub
    Public Sub New(ByVal Articulo As String, ByVal Almacen As String, ByVal FechaDocumento As Date, Optional ByVal NumeroMovimiento As Integer = 0)
        Me.Articulo = Articulo
        Me.Almacen = Almacen
        Me.FechaDocumento = FechaDocumento
        Me.Estado = EstadoStock.NoActualizado
        If NumeroMovimiento > 0 Then Me.NumeroMovimiento = NumeroMovimiento
        Me.GenerarMovimiento = True
    End Sub
End Class

<Serializable()> _
Public Class StockCorrectContext                    '// Es Serializable por que utiliza en la clase StockData
    Public CorreccionEnCantidad As Boolean          '// Correcciones
    Public CorreccionEnCantidad2 As Boolean?        '// Correcciones
    Public CorreccionEnPrecio As Boolean            '// Correcciones
    Public CorreccionEnFecha As Boolean             '// Correcciones
    Public CorreccionEnDocumento As Boolean         '// Correcciones
    Public EsBorrado As Boolean                     '// Correcciones

    Public Sub New()

    End Sub

    Public Sub New(ByVal CorreccionEnCantidad As Boolean, ByVal CorreccionEnPrecio As Boolean, ByVal CorreccionEnFecha As Boolean, ByVal CorreccionEnDocumento As Boolean, ByVal EsBorrado As Boolean)
        Me.CorreccionEnCantidad = CorreccionEnCantidad
        Me.CorreccionEnPrecio = CorreccionEnPrecio
        Me.CorreccionEnFecha = CorreccionEnFecha
        Me.CorreccionEnDocumento = CorreccionEnDocumento
        Me.EsBorrado = EsBorrado
    End Sub
End Class

<Serializable()> _
Public Class StockUpdateData
    Public StockData As StockData
    Public Estado As EstadoStock
    Public IDLineaMovimiento As Integer
    Public NumeroMovimiento As Integer
    Public CantidadMovimiento As Double
    Public CantidadMovimiento2 As Double?
    Public Log As String
    Public Detalle As String

    Public ArticuloAlmacen As DataTable
    Public Movimientos As DataTable
    Public Acumulados As DataTable
    Public Lote As DataTable
    Public Serie As DataTable
    Public Activo As DataTable
    Public HistoricoEstadoActivo As DataTable
End Class

<Serializable()> _
Public Class PropiedadesEstadoNSerie                '// Es Serializable por que utiliza en la clase StockContext
    Public Disponible As Boolean
    Public EnCurso As Boolean
    Public Baja As Boolean
    Public Sistema As Boolean
End Class

#End Region

#End Region

Public Interface IStock
    Function SincronizarSalida(ByVal NumeroMovimiento As Integer, ByVal data As StockData) As ProcesoStocks.DataVinoQ
    Function SincronizarEntrada(ByVal NumeroMovimiento As Integer, ByVal data As StockData) As ProcesoStocks.DataVinoQ
    Function SincronizarEntradaTransferencia(ByVal NumeroMovimiento As Integer, ByVal dataentrada As Negocio.StockData, ByVal datasalida As Negocio.StockData, _
                                             ByVal updateEntrada As StockUpdateData, ByVal updateSalida As StockUpdateData) As ProcesoStocks.DataVinoQ
    Function SincronizarInventario(ByVal NumeroMovimiento As Integer, ByVal data As StockData) As ProcesoStocks.DataVinoQ
    Function SincronizarAjuste(ByVal NumeroMovimiento As Integer, ByVal data As StockData) As ProcesoStocks.DataVinoQ
    Function SincronizarEliminarMovimiento(ByVal NumeroMovimiento As Integer, ByVal IDLineaMovimiento As Integer, ByVal dataOriginal As StockData) As ProcesoStocks.DataVinoQ
    Function SincronizarCorreccionMovimiento(ByVal NumeroMovimiento As Integer, ByVal data As StockData, ByVal dataOriginal As StockData) As ProcesoStocks.DataVinoQ

    Function AltaEntradaVino(ByVal data As DataAltEntVino) As Integer
    Function CrearOperacionArticulosCompatibles(ByVal data As DataArtCompatiblesExp) As String
    Sub ActualizarPrecioEntradaVino(ByVal data As DataPrecioEntVino)
    Sub ActualizarFechaEntradaVino(ByVal data As DataFechaEntVino)
    Sub ActualizarDAAARCEntradaVino(ByVal data As DataActDAAARCEntVino)
End Interface

<Serializable()> _
Public Class DataAltEntVino
    Public IDArticulo As String
    Public IDProveedor As String
    Public Precio As Double
    Public Cantidad As Double
    Public Lote As String
    Public Fecha As DateTime
    Public NDaa As String
    Public ARC As String

    Public Sub New()
    End Sub
    Public Sub New(ByVal IDArticulo As String, ByVal IDProveedor As String, ByVal Precio As Double, ByVal Cantidad As Double, ByVal Lote As String, ByVal Fecha As DateTime, _
                   ByVal NDaa As String, ByVal ARC As String)
        Me.IDArticulo = IDArticulo
        Me.IDProveedor = IDProveedor
        Me.Precio = Precio
        Me.Cantidad = Cantidad
        Me.Lote = Lote
        Me.Fecha = Fecha
        Me.NDaa = NDaa
        Me.ARC = ARC
    End Sub
End Class

<serializable()> _
Public Class DataPrecioEntVino
    Public NEntrada As Integer
    Public Precio As Double

    Public Sub New()
    End Sub
    Public Sub New(ByVal NEntrada As Integer, ByVal Precio As Double)
        Me.NEntrada = NEntrada
        Me.Precio = Precio
    End Sub
End Class

<serializable()> _
Public Class DataFechaEntVino
    Public NEntrada As Integer
    Public Fecha As Date

    Public Sub New()
    End Sub
    Public Sub New(ByVal NEntrada As Integer, ByVal Fecha As Date)
        Me.NEntrada = NEntrada
        Me.Fecha = Fecha
    End Sub
End Class

<Serializable()> _
Public Class DataActDAAARCEntVino
    Public NEntrada As Integer
    Public NDaa As String
    Public ARC As String

    Public Sub New()
    End Sub
    Public Sub New(ByVal NEntrada As String, ByVal NDaa As String, ByVal ARC As String)
        Me.NEntrada = NEntrada
        Me.NDaa = NDaa
        Me.ARC = ARC
    End Sub
End Class

Public Module InventariosPermanentes
    Public ENSAMBLADO_INV_PERMANENTE_STOCKS As String = "Expertis.Business.Financiero.dll"
    Public CLASE_INV_PERMANENTE_STOCKS As String = "Solmicro.Expertis.Business.Financiero.InvPermanenteStock"
End Module

Public Interface IStockInventarioPermanente

    Sub ValidarPeriodoCerrado(ByVal Fecha As Date, ByVal services As ServiceProvider)

    Function SincronizarContaAlbaranCompra(ByVal IDLineaAlbaran As Integer, ByVal Contabilizado As enumContabilizado, ByVal services As ServiceProvider) As Integer

    Function SincronizarContaAlbaranVenta(ByVal IDLineaAlbaran As Integer, ByVal Contabilizado As enumContabilizado, ByVal services As ServiceProvider) As Integer

    Sub SincronizarDescontaAlbaranVenta(ByVal dtApuntesDescontabilizar As DataTable, ByVal services As ServiceProvider)

    Sub SincronizarContaMovimientos(ByVal UpdateData() As StockUpdateData, ByVal services As ServiceProvider)

    Sub SincronizarContaOFControl(ByVal data As Hashtable, ByVal services As ServiceProvider)
    Sub SincronizarContaSalidaMP(ByVal UpdateData() As StockUpdateData, ByVal services As ServiceProvider)
    Sub SincronizarContaEntradaPT(ByVal UpdateData() As StockUpdateData, ByVal services As ServiceProvider)

End Interface

'<Serializable()> _
'Public Class DataAjuste
'    Public Fecha As Date
'    Public Importe As Double
'    ' Public Contabilizado As enumContabilizado?

'    Public Sub New(ByVal Fecha As Date, ByVal Importe As Double)
'        Me.Fecha = Fecha
'        Me.Importe = Importe
'    End Sub
'End Class

<Transactional()> _
Public Class ProcesoStocks
    Inherits ContextBoundObject



    'Private mAcumuladoInfo() As AcumuladoInfo
    <Serializable()> _
    Public Class AcumuladoInfo
        Public IDArticulo As String
        Public IDAlmacen As String
        Public Acumulado As Double
        Public Acumulado2 As Double?

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal Acumulado As Double)
            Me.IDArticulo = IDArticulo
            Me.IDAlmacen = IDAlmacen
            Me.Acumulado = Acumulado
        End Sub
    End Class

    Private Const cnEntidad As String = "tbHistoricoMovimiento"
    Private Const cnMyClass As String = "Stock"
    'Private Const IDTipoMovimientoAV As Integer = 2
    'Private Const IDTipoMovimientoAC As Integer = 1


    <Serializable()> _
    Public Class DataLote
        Public Lote As String
        Public Ubicacion As String
        Public Cantidad As Double
        Public Cantidad2 As Double?
        Public FechaCaducidad As Date?
        Public IDMovimientoEntrada As Integer?
        Public IDMovimientoSalida As Integer?

        Public PrecintaNSerie As String
        Public PrecintaDesde As Integer
        Public PrecintaHasta As Integer

        Public Sub New(ByVal Lote As String, ByVal Ubicacion As String, ByVal Cantidad As Double, Optional ByVal FechaCaducidad As Date = cnMinDate, Optional ByVal Cantidad2 As Double = Double.NaN)
            Me.Lote = Lote
            Me.Ubicacion = Ubicacion
            Me.Cantidad = Cantidad
            If FechaCaducidad <> cnMinDate Then Me.FechaCaducidad = FechaCaducidad
            If Cantidad2 <> Double.NaN Then Me.Cantidad2 = Cantidad2
        End Sub

        Public Sub New(ByVal Lote As String, ByVal Ubicacion As String, ByVal Cantidad As Double, ByVal PrecintaSerie As String, ByVal PrecintaDesde As Integer, ByVal PrecintaHasta As Integer, Optional ByVal FechaCaducidad As Date = cnMinDate, Optional ByVal Cantidad2 As Double = Double.NaN)
            Me.Lote = Lote
            Me.Ubicacion = Ubicacion
            Me.Cantidad = Cantidad
            If FechaCaducidad <> cnMinDate Then Me.FechaCaducidad = FechaCaducidad
            If Cantidad2 <> Double.NaN Then Me.Cantidad2 = Cantidad2
            Me.PrecintaDesde = PrecintaDesde
            Me.PrecintaHasta = PrecintaHasta
            Me.PrecintaNSerie = PrecintaSerie
        End Sub
    End Class

    <Serializable()> _
    Public Class DataLogActualizarStock
        Public Log As String
        Public Articulo As String
        Public Almacen As String
        Public Lote As String
        Public Ubicacion As String
        Public Sub New(ByVal Log As String)
            Me.Log = Log
        End Sub
        Public Sub New(ByVal Log As String, ByVal Articulo As String)
            Me.Log = Log
            Me.Articulo = Articulo
        End Sub
        Public Sub New(ByVal Log As String, ByVal Articulo As String, ByVal Almacen As String)
            Me.Log = Log
            Me.Articulo = Articulo
            Me.Almacen = Almacen
        End Sub
        Public Sub New(ByVal Log As String, ByVal Articulo As String, ByVal Almacen As String, ByVal Lote As String, ByVal Ubicacion As String)
            Me.Log = Log
            Me.Articulo = Articulo
            Me.Almacen = Almacen
            Me.Lote = Lote
            Me.Ubicacion = Ubicacion
        End Sub
    End Class
    <Task()> Public Shared Function LogActualizarStock(ByVal data As DataLogActualizarStock, ByVal services As ServiceProvider) As StockUpdateData
        Dim auxData As New StockData
        auxData.Articulo = data.Articulo
        auxData.Almacen = data.Almacen
        auxData.Lote = data.Lote
        auxData.Ubicacion = data.Ubicacion
        Dim auxUpdateData As New StockUpdateData
        auxUpdateData.StockData = auxData
        auxUpdateData.Detalle = data.Log
        auxUpdateData.Log = data.Log
        Return auxUpdateData
    End Function

#Region "Detalles"

    <Serializable()> _
    Public Class DataDetalleMovimientos
        Public IDArticulo As String
        Public IDAlmacen As String
        Public FechaInicio As Date?
        Public FechaFin As Date?
        Public Stock As Double
        Public Orden As enumstkValoracionFIFO?

        Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal FechaInicio As Date)
            Me.IDArticulo = IDArticulo
            Me.IDAlmacen = IDAlmacen
            Me.FechaInicio = FechaInicio
        End Sub

        Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal Stock As Double, ByVal FechaFin As Date, ByVal Orden As enumstkValoracionFIFO)
            Me.IDArticulo = IDArticulo
            Me.IDAlmacen = IDAlmacen
            Me.FechaInicio = FechaInicio
        End Sub

        Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal FechaInicio As Date, ByVal FechaFin As Date)
            Me.IDArticulo = IDArticulo
            Me.IDAlmacen = IDAlmacen
            Me.FechaInicio = FechaInicio
            Me.FechaFin = FechaFin
        End Sub

    End Class

    <Task()> Public Shared Function DetalleMovimientosStockAFecha(ByVal data As DataDetalleMovimientos, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDArticulo) > 0 AndAlso Length(data.IDAlmacen) > 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
            f.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
            If Not data.FechaInicio Is Nothing AndAlso data.FechaInicio <> cnMinDate Then
                f.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThanOrEqual, data.FechaInicio))
            End If
            If Not data.FechaFin Is Nothing AndAlso data.FechaFin <> cnMinDate Then
                f.Add(New DateFilterItem("FechaDocumento", FilterOperator.LessThanOrEqual, data.FechaFin))
            End If

            Return New DataEngine().Filter("vNegDetalleMovimientosStockAFecha", f, , "FechaDocumento DESC")
        End If
        Exit Function
    End Function

    <Task()> Public Shared Function DetalleMovimientosFIFO(ByVal data As DataDetalleMovimientos, ByVal services As ServiceProvider) As DataTable
        If Length(data.IDArticulo) > 0 AndAlso Length(data.IDAlmacen) > 0 AndAlso data.Stock > 0 Then
            Dim dtMovimientos As DataTable = AdminData.Execute("sp_DetalleMovimientosFIFO", False, data.IDAlmacen, data.Stock, data.Orden)
            If data.FechaFin <> cnMinDate Then
                If Not dtMovimientos Is Nothing AndAlso dtMovimientos.Rows.Count > 0 Then
                    Dim FilDt As New Filter
                    If Not data.FechaInicio Is Nothing AndAlso data.FechaInicio <> cnMinDate Then
                        FilDt.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThanOrEqual, data.FechaInicio))
                    End If
                    FilDt.Add("FechaDocumento", FilterOperator.LessThan, data.FechaFin, FilterType.DateTime)
                    dtMovimientos.DefaultView.RowFilter = FilDt.Compose(New AdoFilterComposer)
                End If
            End If
            Return dtMovimientos
        End If
    End Function

    <Task()> Public Shared Function DetalleMovimientosPrecioMedio(ByVal data As DataDetalleMovimientos, ByVal services As ServiceProvider) As DataTable
        If data.FechaFin Is Nothing OrElse data.FechaFin = cnMinDate Then data.FechaFin = Today.Date
        If Length(data.IDArticulo) > 0 AndAlso Length(data.IDAlmacen) > 0 AndAlso (data.FechaInicio < data.FechaFin) Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
            f.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
            If Not data.FechaInicio Is Nothing AndAlso data.FechaInicio <> cnMinDate Then
                f.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThanOrEqual, data.FechaInicio))
            End If
            f.Add(New DateFilterItem("FechaDocumento", FilterOperator.LessThanOrEqual, data.FechaFin))
            f.Add(New NumberFilterItem("Cantidad", FilterOperator.GreaterThan, 0))

            Return New BE.DataEngine().Filter("vNegDetalleMovimientosPrecioMedio", f, , "FechaDocumento DESC, IDLineaMovimiento DESC")
        End If
    End Function

#End Region

#Region " Stock desde Albaranes "

    Public Class DataCrearStockDataAlbaran
        Public IDAlbaran As Integer
        Public NAlbaran As String
        Public FechaAlbaran As Date
        Public LineaAlbaran As DataRow
        Public LineaAlbaranLote As DataLote
        'Public IDCircuito As Circuito
        Public ImporteExtraA As Double
        Public ImporteExtraB As Double

        Public Sub New(ByVal IDAlbaran As Integer, ByVal NAlbaran As String, ByVal FechaAlbaran As Date, ByVal LineaAlbaran As DataRow, ByVal LineaAlbaranLote As DataLote, ByVal ImporteExtraA As Double, ByVal ImporteExtraB As Double)
            Me.IDAlbaran = IDAlbaran
            Me.NAlbaran = NAlbaran
            Me.FechaAlbaran = FechaAlbaran
            Me.LineaAlbaran = LineaAlbaran
            Me.LineaAlbaranLote = LineaAlbaranLote
            Me.ImporteExtraA = ImporteExtraA
            Me.ImporteExtraB = ImporteExtraB
        End Sub

    End Class
    <Task()> Public Shared Function CrearStockDataAlbaran(ByVal data As DataCrearStockDataAlbaran, ByVal services As ServiceProvider) As StockData
        '//Volcar los datos en un objeto stockData sin importar si la gestion es normal, por lotes o numeros de serie.
        '//Este control ya se hace en las funciones de actualizacion el stock.
        Dim stkData As New StockData
        stkData.Articulo = data.LineaAlbaran("IDArticulo")
        stkData.Almacen = data.LineaAlbaran("IDAlmacen")

        Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, stkData.Articulo, services)
        If Length(data.LineaAlbaran("Lote")) > 0 Then
            stkData.Lote = data.LineaAlbaran("Lote")
            stkData.NSerie = data.LineaAlbaran("Lote")
        End If
        If Length(data.LineaAlbaran("Ubicacion")) > 0 Then
            stkData.Ubicacion = data.LineaAlbaran("Ubicacion")
        End If
        stkData.IDDocumento = data.IDAlbaran
        stkData.Documento = data.NAlbaran
        ' stkData.Texto = data.LineaAlbaran("Texto") & String.Empty
        stkData.FechaDocumento = data.FechaAlbaran
        If Not data.LineaAlbaranLote Is Nothing Then
            stkData.Lote = data.LineaAlbaranLote.Lote
            stkData.Ubicacion = data.LineaAlbaranLote.Ubicacion
            stkData.Cantidad = data.LineaAlbaranLote.Cantidad
            If Not String.IsNullOrEmpty(data.LineaAlbaranLote.PrecintaNSerie) Then
                stkData.PrecintaNSerie = data.LineaAlbaranLote.PrecintaNSerie
                stkData.PrecintaDesde = data.LineaAlbaranLote.PrecintaDesde
                stkData.PrecintaHasta = data.LineaAlbaranLote.PrecintaHasta
            End If
            If SegundaUnidad Then stkData.Cantidad2 = data.LineaAlbaranLote.Cantidad2
        Else
            stkData.Cantidad = data.LineaAlbaran("QInterna")
            If SegundaUnidad Then
                If Length(data.LineaAlbaran("QInterna2")) > 0 Then
                    stkData.Cantidad2 = CDbl(data.LineaAlbaran("QInterna2"))
                Else
                    ApplicationService.GenerateError("El Artículo {0} se gestiona con Doble Unidad. Debe indicar la misma.", Quoted(stkData.Articulo))
                End If
            End If
        End If
        If IsNumeric(data.LineaAlbaran("IDObra")) AndAlso data.LineaAlbaran("IDObra") > 0 Then
            stkData.Obra = data.LineaAlbaran("IDObra")
        End If
        If Length(data.LineaAlbaran("IDEstadoActivo")) > 0 Then
            stkData.EstadoNSerie = data.LineaAlbaran("IDEstadoActivo")
        End If
        If Length(data.LineaAlbaran("IDOperario")) > 0 Then
            stkData.Operario = data.LineaAlbaran("IDOperario")
        Else
            Dim StrOperario As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf Operario.ObtenerIDOperarioUsuario, Nothing, services)
            If Len(StrOperario) > 0 Then stkData.Operario = StrOperario
        End If

        If data.LineaAlbaran("Factor") <> 0 AndAlso data.LineaAlbaran("UdValoracion") <> 0 AndAlso stkData.Cantidad <> 0 Then
            Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
            Dim monedaA As MonedaInfo = Monedas.MonedaA
            Dim monedaB As MonedaInfo = Monedas.MonedaB

            stkData.PrecioA = xRound((data.ImporteExtraA / stkData.Cantidad) + (data.LineaAlbaran("PrecioA") / data.LineaAlbaran("Factor") / data.LineaAlbaran("UdValoracion") * (1 - data.LineaAlbaran("Dto1") / 100) * (1 - data.LineaAlbaran("Dto2") / 100) * (1 - data.LineaAlbaran("Dto3") / 100) * (1 - data.LineaAlbaran("Dto") / 100) * (1 - data.LineaAlbaran("DtoProntoPago") / 100)), monedaA.NDecimalesPrecio)
            stkData.PrecioB = xRound((data.ImporteExtraA / stkData.Cantidad) + (data.LineaAlbaran("PrecioB") / data.LineaAlbaran("Factor") / data.LineaAlbaran("UdValoracion") * (1 - data.LineaAlbaran("Dto1") / 100) * (1 - data.LineaAlbaran("Dto2") / 100) * (1 - data.LineaAlbaran("Dto3") / 100) * (1 - data.LineaAlbaran("Dto") / 100) * (1 - data.LineaAlbaran("DtoProntoPago") / 100)), monedaB.NDecimalesPrecio)
        End If

        Return stkData
    End Function

    <Serializable()> _
    Public Class DataCrearMovimiento
        Public NumeroMovimiento As Integer
        Public StkData As StockData

        Public Sub New(ByVal NumeroMovimiento As Integer, ByVal StkData As StockData)
            Me.NumeroMovimiento = NumeroMovimiento
            Me.StkData = StkData
        End Sub
    End Class
    <Task()> Public Shared Function CrearMovimiento(ByVal data As DataCrearMovimiento, ByVal services As ServiceProvider) As StockUpdateData
        If data.NumeroMovimiento = 0 Then
            data.NumeroMovimiento = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf NuevoNumeroMovimiento, Nothing, services)
        End If
        Dim stkUpdateData As StockUpdateData
        If data.StkData.TipoMovimiento = enumTipoMovimiento.tmSalSubcontratacion OrElse data.StkData.TipoMovimiento = enumTipoMovimiento.tmSalAlbaranVenta Then
            Dim dataSal As New DataNumeroMovimientoSinc(data.NumeroMovimiento, data.StkData)
            stkUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf Salida, dataSal, services)
        Else
            Dim datMovto As New DataNumeroMovimientoSinc(data.NumeroMovimiento, data.StkData)
            stkUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf Entrada, datMovto, services)
        End If
        'If Not stkUpdateData Is Nothing Then
        '    If stkUpdateData.Estado = EstadoStock.NoActualizado Then
        '        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.RollbackTransaction, True, services)
        '    End If
        'End If
        Return stkUpdateData
    End Function

    <Serializable()> _
    Public Class DataActualizarStockAlbaranTx
        Public IDCliente As String
        Public IDAlbaran As Integer
        Public NAlbaran As String
        Public FechaAlbaran As Date
        Public NumeroMovimiento As Integer?
        Public IDTipoAlbaran As String
        Public IDAlmacenDeposito As String
        Public LineaAlbaran As DataRow
        ' Public LotesLineaAlbaran() As DataLote
        Public LotesLineaAlbaran As DataTable
        Public ImporteExtraA As Double
        Public ImporteExtraB As Double

        Public LineasAlbaran As DataTable
        Public Circuito As Circuito?
    End Class

    <Task()> Public Shared Function ActualizarStockAlbaranTx(ByVal data As DataActualizarStockAlbaranTx, ByVal services As ServiceProvider) As StockUpdateData()
        If data.Circuito Is Nothing Then Exit Function
        Dim aStkUD(-1) As StockUpdateData
        Dim valorTipoAlbaran As BusinessEnum.enumTipoAlbaran
        If data.Circuito = Circuito.Ventas Then
            Dim dva As New DataValidarAlmacenes
            dva.IDTipoAlbaran = data.IDTipoAlbaran
            dva.IDAlmacenDeposito = data.IDAlmacenDeposito
            dva.LineaAlbaran = data.LineaAlbaran

            valorTipoAlbaran = services.GetService(Of BusinessEnum.enumTipoAlbaran)()
            If valorTipoAlbaran = enumTipoAlbaran.Desconocido Then
                valorTipoAlbaran = ProcessServer.ExecuteTask(Of String, enumTipoAlbaran)(AddressOf ProcesoAlbaranVenta.ValidarTipoAlbaran, data.IDTipoAlbaran, services)
            End If

            Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataValidarAlmacenes, StockUpdateData)(AddressOf ValidarAlmacenes, dva, services)
            If Not updateData Is Nothing Then
                ReDim Preserve aStkUD(aStkUD.Length)
                aStkUD(aStkUD.Length - 1) = updateData
                Return aStkUD
            End If
        End If


        Select Case data.Circuito
            Case Circuito.Ventas
                Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.LineaAlbaran("IDArticulo"))
                If ArtInfo.GestionStockPorLotes Then
                    Dim QLotesAsignado As Decimal = CDec(data.LotesLineaAlbaran.Compute("SUM(QInterna)", Nothing))
                    If CDec(data.LineaAlbaran("QInterna")) <> QLotesAsignado Then
                        ApplicationService.GenerateError([Global].ParseFormatString("La Cantidad de la línea del Albarán debe coincidir con la Cantidad repartida en los Lotes. {0} Artículo {1}", vbNewLine, Quoted(data.LineaAlbaran("IDArticulo"))))
                    Else
                        Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.LineaAlbaran("IDArticulo"), services)
                        If SegundaUnidad Then
                            Dim QLotesAsignado2 As Decimal = CDec(data.LotesLineaAlbaran.Compute("SUM(QInterna2)", Nothing))
                            If CDec(data.LineaAlbaran("QInterna2")) <> QLotesAsignado2 Then
                                ApplicationService.GenerateError([Global].ParseFormatString("La Cantidad en Segunda Unidad de la línea del Albarán debe coincidir con la Cantidad repartida en los Lotes. {0} Artículo {1}", vbNewLine, Quoted(data.LineaAlbaran("IDArticulo"))))
                            End If
                        End If
                    End If
                End If
        End Select

        Dim stkData As StockData
        Dim csd As New DataCrearStockDataAlbaran(data.IDAlbaran, data.NAlbaran, data.FechaAlbaran, data.LineaAlbaran, Nothing, data.ImporteExtraA, data.ImporteExtraB)
        If data.LotesLineaAlbaran Is Nothing OrElse data.LotesLineaAlbaran.Rows.Count = 0 Then
            '//Linea de Albarán sin lotes
            stkData = ProcessServer.ExecuteTask(Of DataCrearStockDataAlbaran, StockData)(AddressOf CrearStockDataAlbaran, csd, services)
            Dim EstTipoMovto As New DataEstablecerTipoMovimiento(stkData, csd.LineaAlbaran, data.LineasAlbaran)
            Select Case data.Circuito
                Case Circuito.Compras
                    stkData = ProcessServer.ExecuteTask(Of DataEstablecerTipoMovimiento, StockData)(AddressOf EstablecerTipoMovimientoAC, EstTipoMovto, services)
                Case Circuito.Ventas
                    stkData = ProcessServer.ExecuteTask(Of DataEstablecerTipoMovimiento, StockData)(AddressOf EstablecerTipoMovimientoAV, EstTipoMovto, services)
            End Select
            Dim dcm As New DataCrearMovimiento(data.NumeroMovimiento, stkData)
            Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataCrearMovimiento, StockUpdateData)(AddressOf CrearMovimiento, dcm, services)
            If Not updateData Is Nothing Then
                Select Case data.Circuito
                    Case Circuito.Compras
                        If updateData.Estado <> EstadoStock.NoActualizado Then
                            If updateData.Estado = EstadoStock.Actualizado AndAlso data.NumeroMovimiento <> updateData.NumeroMovimiento Then data.NumeroMovimiento = updateData.NumeroMovimiento
                            ProcessServer.ExecuteTask(Of DataRow)(AddressOf PrepararActivoUltimaCompra, data.LineaAlbaran, services)
                        Else
                            'ProcessServer.ExecuteTask(Of DataActualizarLineas)(AddressOf MovimientoCorreccion, act, services)
                            If updateData.Estado = EstadoStock.NoActualizado AndAlso Length(data.LineaAlbaran("IDMovimiento")) > 0 Then
                                Dim actMovto As New DataActualizarMovimiento(enumTipoActualizacion.Eliminar, data.LineaAlbaran("IDMovimiento"))
                                ProcessServer.ExecuteTask(Of DataActualizarMovimiento)(AddressOf ActualizarMovimiento, actMovto, services)
                            End If
                        End If

                        Dim act As New DataActualizarLineas(updateData, data.LineaAlbaran)
                        ProcessServer.ExecuteTask(Of DataActualizarLineas)(AddressOf ActualizarLineasAC, act, services)
                    Case Circuito.Ventas
                        'ProcessServer.ExecuteTask(Of DataActualizarLineas)(AddressOf MovimientoCorreccion, act, services)

                        If updateData.Estado = EstadoStock.NoActualizado AndAlso Length(data.LineaAlbaran("IDMovimiento")) > 0 Then
                            Dim actMovto As New DataActualizarMovimiento(enumTipoActualizacion.Eliminar, data.LineaAlbaran("IDMovimiento"))
                            ProcessServer.ExecuteTask(Of DataActualizarMovimiento)(AddressOf ActualizarMovimiento, actMovto, services)
                        End If
                        Dim act As New DataActualizarLineas(updateData, data.LineaAlbaran)
                        ProcessServer.ExecuteTask(Of DataActualizarLineas)(AddressOf ActualizarLineasAV, act, services)

                        If updateData.Estado = EstadoStock.Actualizado Then
                            Dim desda As New DataEntradaStockEnDepositoOAlquiler
                            desda.IDCliente = data.IDCliente
                            desda.lineaAlbaran = data.LineaAlbaran
                            desda.NumeroMovimiento = updateData.NumeroMovimiento
                            desda.Salida = stkData
                            desda.UpdateSalida = updateData
                            If valorTipoAlbaran = enumTipoAlbaran.Intercambio Then
                                ProcessServer.ExecuteTask(Of DataActualizarStockAlbaranTx)(AddressOf ActualizarLineaPedidoDeIntercambio, data, services)
                                ProcessServer.ExecuteTask(Of DataEntradaStockEnDepositoOAlquiler)(AddressOf EntradaStockDeIntercambio, desda, services)
                            Else
                                Dim updateEntrada As StockUpdateData = ProcessServer.ExecuteTask(Of DataEntradaStockEnDepositoOAlquiler, StockUpdateData)(AddressOf EntradaStockEnDepositoOAlquiler, desda, services)
                                If Not updateEntrada Is Nothing Then
                                    ReDim Preserve aStkUD(aStkUD.Length)
                                    aStkUD(aStkUD.Length - 1) = updateEntrada
                                End If
                            End If
                        End If
                End Select
            End If
            If Not updateData Is Nothing Then
                ReDim Preserve aStkUD(aStkUD.Length)
                aStkUD(aStkUD.Length - 1) = updateData
            End If
        Else

            '//DeshacerActualizados. Si tenemos varios lotes, se actualizan todos menos el último, por jemplo, los que se han actualizado se quedan como tal y el último 
            '//se queda sin actualizar. Para evitar esto, utilizamos ésta varible.
            Dim DeshacerActualizados As Boolean = False
            ProcessServer.ExecuteTask(Of Object)(AddressOf Comunes.BeginTransaction, Nothing, services)

            '//Linea de Albarán con lotes
            Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.LineaAlbaran("IDArticulo"), services)
            For Each lineaLote As DataRow In data.LotesLineaAlbaran.Rows
                If lineaLote.Table.Columns.Contains("SeriePrecinta") AndAlso Length(lineaLote("SeriePrecinta")) > 0 Then
                    csd.LineaAlbaranLote = New DataLote(lineaLote("Lote"), lineaLote("Ubicacion"), lineaLote("QInterna"), lineaLote("SeriePrecinta"), lineaLote("NDesdePrecinta"), lineaLote("NHastaPrecinta")) 'lote
                Else
                    csd.LineaAlbaranLote = New DataLote(lineaLote("Lote"), lineaLote("Ubicacion"), lineaLote("QInterna")) 'lote
                End If
                If SegundaUnidad Then csd.LineaAlbaranLote.Cantidad2 = CDbl(Nz(lineaLote("QInterna2"), 0))

                stkData = ProcessServer.ExecuteTask(Of DataCrearStockDataAlbaran, StockData)(AddressOf CrearStockDataAlbaran, csd, services)
                Dim EstTipoMovto As New DataEstablecerTipoMovimiento(stkData, csd.LineaAlbaran, data.LineasAlbaran)
                Select Case data.Circuito
                    Case Circuito.Compras
                        stkData = ProcessServer.ExecuteTask(Of DataEstablecerTipoMovimiento, StockData)(AddressOf EstablecerTipoMovimientoAC, EstTipoMovto, services)
                    Case Circuito.Ventas
                        stkData = ProcessServer.ExecuteTask(Of DataEstablecerTipoMovimiento, StockData)(AddressOf EstablecerTipoMovimientoAV, EstTipoMovto, services)
                End Select
                Dim dcm As New DataCrearMovimiento(data.NumeroMovimiento, stkData)
                Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataCrearMovimiento, StockUpdateData)(AddressOf CrearMovimiento, dcm, services)
                If Not updateData Is Nothing Then
                    Select Case data.Circuito
                        Case Circuito.Compras
                            If updateData.Estado <> EstadoStock.NoActualizado Then
                                If updateData.Estado = EstadoStock.Actualizado AndAlso data.NumeroMovimiento <> updateData.NumeroMovimiento Then data.NumeroMovimiento = updateData.NumeroMovimiento
                                ProcessServer.ExecuteTask(Of DataRow)(AddressOf PrepararActivoUltimaCompra, data.LineaAlbaran, services)
                            Else
                                If Length(lineaLote("IDMovimientoEntrada")) > 0 Then
                                    Dim actMovto As New DataActualizarMovimiento(enumTipoActualizacion.Eliminar, lineaLote("IDMovimientoEntrada"))
                                    ProcessServer.ExecuteTask(Of DataActualizarMovimiento)(AddressOf ActualizarMovimiento, actMovto, services)
                                End If
                            End If
                            Dim act As New DataActualizarLineas(updateData, data.LineaAlbaran, lineaLote)
                            ProcessServer.ExecuteTask(Of DataActualizarLineas)(AddressOf ActualizarLineasAC, act, services)
                        Case Circuito.Ventas
                            If updateData.Estado = EstadoStock.NoActualizado AndAlso Length(lineaLote("IDMovimientoSalida")) > 0 Then
                                Dim actMovto As New DataActualizarMovimiento(enumTipoActualizacion.Eliminar, lineaLote("IDMovimientoSalida"))
                                ProcessServer.ExecuteTask(Of DataActualizarMovimiento)(AddressOf ActualizarMovimiento, actMovto, services)
                            End If
                            Dim act As New DataActualizarLineas(updateData, data.LineaAlbaran, lineaLote)
                            ProcessServer.ExecuteTask(Of DataActualizarLineas)(AddressOf ActualizarLineasAV, act, services)
                            If updateData.Estado = EstadoStock.Actualizado Then
                                Dim desda As New DataEntradaStockEnDepositoOAlquiler
                                desda.IDCliente = data.IDCliente
                                desda.lineaAlbaran = data.LineaAlbaran
                                desda.NumeroMovimiento = updateData.NumeroMovimiento
                                desda.Salida = stkData
                                desda.UpdateSalida = updateData
                                desda.lineaLote = lineaLote
                                If valorTipoAlbaran = enumTipoAlbaran.Intercambio Then
                                    ProcessServer.ExecuteTask(Of DataActualizarStockAlbaranTx)(AddressOf ActualizarLineaPedidoDeIntercambio, data, services)
                                    ProcessServer.ExecuteTask(Of DataEntradaStockEnDepositoOAlquiler)(AddressOf EntradaStockDeIntercambio, desda, services)
                                Else
                                    Dim updateEntrada As StockUpdateData = ProcessServer.ExecuteTask(Of DataEntradaStockEnDepositoOAlquiler, StockUpdateData)(AddressOf EntradaStockEnDepositoOAlquiler, desda, services)
                                    If Not updateEntrada Is Nothing Then
                                        ReDim Preserve aStkUD(aStkUD.Length)
                                        aStkUD(aStkUD.Length - 1) = updateEntrada
                                    End If
                                End If
                            End If
                    End Select
                    If updateData.Estado = EstadoStock.NoActualizado Then
                        DeshacerActualizados = True
                    End If
                    ReDim Preserve aStkUD(aStkUD.Length)
                    aStkUD(aStkUD.Length - 1) = updateData
                End If
            Next
            If DeshacerActualizados Then

                For i As Integer = 0 To aStkUD.Length - 1
                    If aStkUD(i).Estado = EstadoStock.Actualizado Then
                        aStkUD(i).Estado = EstadoStock.NoActualizado
                        Dim datMsg As New DataMessage(48)
                        aStkUD(i).Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, datMsg, services)
                    End If
                Next

                Select Case data.Circuito
                    Case Circuito.Compras
                        For Each lineaLote As DataRow In data.LotesLineaAlbaran.Rows
                            lineaLote("IDMovimientoEntrada") = System.DBNull.Value
                        Next
                    Case Circuito.Ventas
                        For Each lineaLote As DataRow In data.LotesLineaAlbaran.Rows
                            lineaLote("IDMovimientoSalida") = System.DBNull.Value
                        Next

                End Select
                data.LineaAlbaran("EstadoStock") = enumavlEstadoStock.avlNoActualizado

                ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.RollbackTransaction, True, services)
            Else
                ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.CommitTransaction, True, services)
            End If
        End If
        Return aStkUD
    End Function
    'David Velasco 10/8/22
    <Task()> Public Shared Function ActualizarStockAlbaranTx2(ByVal data As DataActualizarStockAlbaranTx, ByVal services As ServiceProvider) As StockUpdateData()
        If data.Circuito Is Nothing Then Exit Function
        Dim aStkUD(-1) As StockUpdateData
        Dim valorTipoAlbaran As BusinessEnum.enumTipoAlbaran
        If data.Circuito = Circuito.Ventas Then
            Dim dva As New DataValidarAlmacenes
            dva.IDTipoAlbaran = data.IDTipoAlbaran
            dva.IDAlmacenDeposito = data.IDAlmacenDeposito
            dva.LineaAlbaran = data.LineaAlbaran

            valorTipoAlbaran = services.GetService(Of BusinessEnum.enumTipoAlbaran)()
            If valorTipoAlbaran = enumTipoAlbaran.Desconocido Then
                valorTipoAlbaran = ProcessServer.ExecuteTask(Of String, enumTipoAlbaran)(AddressOf ProcesoAlbaranVenta.ValidarTipoAlbaran, data.IDTipoAlbaran, services)
            End If

            Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataValidarAlmacenes, StockUpdateData)(AddressOf ValidarAlmacenes, dva, services)
            If Not updateData Is Nothing Then
                ReDim Preserve aStkUD(aStkUD.Length)
                aStkUD(aStkUD.Length - 1) = updateData
                Return aStkUD
            End If
        End If


        Select Case data.Circuito
            Case Circuito.Ventas
                Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.LineaAlbaran("IDArticulo"))
                If ArtInfo.GestionStockPorLotes Then
                    Dim QLotesAsignado As Decimal = CDec(data.LotesLineaAlbaran.Compute("SUM(QInterna)", Nothing))
                    If CDec(data.LineaAlbaran("QInterna")) <> QLotesAsignado Then
                        ApplicationService.GenerateError([Global].ParseFormatString("La Cantidad de la línea del Albarán debe coincidir con la Cantidad repartida en los Lotes. {0} Artículo {1}", vbNewLine, Quoted(data.LineaAlbaran("IDArticulo"))))
                    Else
                        Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.LineaAlbaran("IDArticulo"), services)
                        If SegundaUnidad Then
                            Dim QLotesAsignado2 As Decimal = CDec(data.LotesLineaAlbaran.Compute("SUM(QInterna2)", Nothing))
                            If CDec(data.LineaAlbaran("QInterna2")) <> QLotesAsignado2 Then
                                ApplicationService.GenerateError([Global].ParseFormatString("La Cantidad en Segunda Unidad de la línea del Albarán debe coincidir con la Cantidad repartida en los Lotes. {0} Artículo {1}", vbNewLine, Quoted(data.LineaAlbaran("IDArticulo"))))
                            End If
                        End If
                    End If
                End If
        End Select

        Dim stkData As StockData
        Dim csd As New DataCrearStockDataAlbaran(data.IDAlbaran, data.NAlbaran, data.FechaAlbaran, data.LineaAlbaran, Nothing, data.ImporteExtraA, data.ImporteExtraB)
        If data.LotesLineaAlbaran Is Nothing OrElse data.LotesLineaAlbaran.Rows.Count = 0 Then
            '//Linea de Albarán sin lotes
            stkData = ProcessServer.ExecuteTask(Of DataCrearStockDataAlbaran, StockData)(AddressOf CrearStockDataAlbaran, csd, services)
            Dim EstTipoMovto As New DataEstablecerTipoMovimiento(stkData, csd.LineaAlbaran, data.LineasAlbaran)
            Select Case data.Circuito
                Case Circuito.Compras
                    stkData = ProcessServer.ExecuteTask(Of DataEstablecerTipoMovimiento, StockData)(AddressOf EstablecerTipoMovimientoAC, EstTipoMovto, services)
                Case Circuito.Ventas
                    stkData = ProcessServer.ExecuteTask(Of DataEstablecerTipoMovimiento, StockData)(AddressOf EstablecerTipoMovimientoAV, EstTipoMovto, services)
            End Select
            Dim dcm As New DataCrearMovimiento(data.NumeroMovimiento, stkData)
            Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataCrearMovimiento, StockUpdateData)(AddressOf CrearMovimiento, dcm, services)
            If Not updateData Is Nothing Then
                Select Case data.Circuito
                    Case Circuito.Compras
                        If updateData.Estado <> EstadoStock.NoActualizado Then
                            If updateData.Estado = EstadoStock.Actualizado AndAlso data.NumeroMovimiento <> updateData.NumeroMovimiento Then data.NumeroMovimiento = updateData.NumeroMovimiento
                            ProcessServer.ExecuteTask(Of DataRow)(AddressOf PrepararActivoUltimaCompra, data.LineaAlbaran, services)
                        Else
                            'ProcessServer.ExecuteTask(Of DataActualizarLineas)(AddressOf MovimientoCorreccion, act, services)
                            If updateData.Estado = EstadoStock.NoActualizado AndAlso Length(data.LineaAlbaran("IDMovimiento")) > 0 Then
                                Dim actMovto As New DataActualizarMovimiento(enumTipoActualizacion.Eliminar, data.LineaAlbaran("IDMovimiento"))
                                ProcessServer.ExecuteTask(Of DataActualizarMovimiento)(AddressOf ActualizarMovimiento, actMovto, services)
                            End If
                        End If

                        Dim act As New DataActualizarLineas(updateData, data.LineaAlbaran)
                        ProcessServer.ExecuteTask(Of DataActualizarLineas)(AddressOf ActualizarLineasAC, act, services)
                    Case Circuito.Ventas
                        'ProcessServer.ExecuteTask(Of DataActualizarLineas)(AddressOf MovimientoCorreccion, act, services)

                        If updateData.Estado = EstadoStock.NoActualizado AndAlso Length(data.LineaAlbaran("IDMovimiento")) > 0 Then
                            Dim actMovto As New DataActualizarMovimiento(enumTipoActualizacion.Eliminar, data.LineaAlbaran("IDMovimiento"))
                            ProcessServer.ExecuteTask(Of DataActualizarMovimiento)(AddressOf ActualizarMovimiento, actMovto, services)
                        End If
                        Dim act As New DataActualizarLineas(updateData, data.LineaAlbaran)
                        ProcessServer.ExecuteTask(Of DataActualizarLineas)(AddressOf ActualizarLineasAV, act, services)

                        'If updateData.Estado = EstadoStock.Actualizado Then
                        '    Dim desda As New DataEntradaStockEnDepositoOAlquiler
                        '    desda.IDCliente = data.IDCliente
                        '    desda.lineaAlbaran = data.LineaAlbaran
                        '    desda.NumeroMovimiento = updateData.NumeroMovimiento
                        '    desda.Salida = stkData
                        '    desda.UpdateSalida = updateData
                        '    If valorTipoAlbaran = enumTipoAlbaran.Intercambio Then
                        '        ProcessServer.ExecuteTask(Of DataActualizarStockAlbaranTx)(AddressOf ActualizarLineaPedidoDeIntercambio, data, services)
                        '        ProcessServer.ExecuteTask(Of DataEntradaStockEnDepositoOAlquiler)(AddressOf EntradaStockDeIntercambio, desda, services)
                        '    Else
                        '        Dim updateEntrada As StockUpdateData = ProcessServer.ExecuteTask(Of DataEntradaStockEnDepositoOAlquiler, StockUpdateData)(AddressOf EntradaStockEnDepositoOAlquiler, desda, services)
                        '        If Not updateEntrada Is Nothing Then
                        '            ReDim Preserve aStkUD(aStkUD.Length)
                        '            aStkUD(aStkUD.Length - 1) = updateEntrada
                        '        End If
                        '    End If
                        'End If
                End Select
            End If
            If Not updateData Is Nothing Then
                ReDim Preserve aStkUD(aStkUD.Length)
                aStkUD(aStkUD.Length - 1) = updateData
            End If
        Else

            '//DeshacerActualizados. Si tenemos varios lotes, se actualizan todos menos el último, por jemplo, los que se han actualizado se quedan como tal y el último 
            '//se queda sin actualizar. Para evitar esto, utilizamos ésta varible.
            Dim DeshacerActualizados As Boolean = False
            ProcessServer.ExecuteTask(Of Object)(AddressOf Comunes.BeginTransaction, Nothing, services)

            '//Linea de Albarán con lotes
            Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.LineaAlbaran("IDArticulo"), services)
            For Each lineaLote As DataRow In data.LotesLineaAlbaran.Rows
                If lineaLote.Table.Columns.Contains("SeriePrecinta") AndAlso Length(lineaLote("SeriePrecinta")) > 0 Then
                    csd.LineaAlbaranLote = New DataLote(lineaLote("Lote"), lineaLote("Ubicacion"), lineaLote("QInterna"), lineaLote("SeriePrecinta"), lineaLote("NDesdePrecinta"), lineaLote("NHastaPrecinta")) 'lote
                Else
                    csd.LineaAlbaranLote = New DataLote(lineaLote("Lote"), lineaLote("Ubicacion"), lineaLote("QInterna")) 'lote
                End If
                If SegundaUnidad Then csd.LineaAlbaranLote.Cantidad2 = CDbl(Nz(lineaLote("QInterna2"), 0))

                stkData = ProcessServer.ExecuteTask(Of DataCrearStockDataAlbaran, StockData)(AddressOf CrearStockDataAlbaran, csd, services)
                Dim EstTipoMovto As New DataEstablecerTipoMovimiento(stkData, csd.LineaAlbaran, data.LineasAlbaran)
                Select Case data.Circuito
                    Case Circuito.Compras
                        stkData = ProcessServer.ExecuteTask(Of DataEstablecerTipoMovimiento, StockData)(AddressOf EstablecerTipoMovimientoAC, EstTipoMovto, services)
                    Case Circuito.Ventas
                        stkData = ProcessServer.ExecuteTask(Of DataEstablecerTipoMovimiento, StockData)(AddressOf EstablecerTipoMovimientoAV, EstTipoMovto, services)
                End Select
                Dim dcm As New DataCrearMovimiento(data.NumeroMovimiento, stkData)
                Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataCrearMovimiento, StockUpdateData)(AddressOf CrearMovimiento, dcm, services)
                If Not updateData Is Nothing Then
                    Select Case data.Circuito
                        Case Circuito.Compras
                            If updateData.Estado <> EstadoStock.NoActualizado Then
                                If updateData.Estado = EstadoStock.Actualizado AndAlso data.NumeroMovimiento <> updateData.NumeroMovimiento Then data.NumeroMovimiento = updateData.NumeroMovimiento
                                ProcessServer.ExecuteTask(Of DataRow)(AddressOf PrepararActivoUltimaCompra, data.LineaAlbaran, services)
                            Else
                                If Length(lineaLote("IDMovimientoEntrada")) > 0 Then
                                    Dim actMovto As New DataActualizarMovimiento(enumTipoActualizacion.Eliminar, lineaLote("IDMovimientoEntrada"))
                                    ProcessServer.ExecuteTask(Of DataActualizarMovimiento)(AddressOf ActualizarMovimiento, actMovto, services)
                                End If
                            End If
                            Dim act As New DataActualizarLineas(updateData, data.LineaAlbaran, lineaLote)
                            ProcessServer.ExecuteTask(Of DataActualizarLineas)(AddressOf ActualizarLineasAC, act, services)
                        Case Circuito.Ventas
                            If updateData.Estado = EstadoStock.NoActualizado AndAlso Length(lineaLote("IDMovimientoSalida")) > 0 Then
                                Dim actMovto As New DataActualizarMovimiento(enumTipoActualizacion.Eliminar, lineaLote("IDMovimientoSalida"))
                                ProcessServer.ExecuteTask(Of DataActualizarMovimiento)(AddressOf ActualizarMovimiento, actMovto, services)
                            End If
                            Dim act As New DataActualizarLineas(updateData, data.LineaAlbaran, lineaLote)
                            ProcessServer.ExecuteTask(Of DataActualizarLineas)(AddressOf ActualizarLineasAV, act, services)
                            If updateData.Estado = EstadoStock.Actualizado Then
                                Dim desda As New DataEntradaStockEnDepositoOAlquiler
                                desda.IDCliente = data.IDCliente
                                desda.lineaAlbaran = data.LineaAlbaran
                                desda.NumeroMovimiento = updateData.NumeroMovimiento
                                desda.Salida = stkData
                                desda.UpdateSalida = updateData
                                desda.lineaLote = lineaLote
                                If valorTipoAlbaran = enumTipoAlbaran.Intercambio Then
                                    ProcessServer.ExecuteTask(Of DataActualizarStockAlbaranTx)(AddressOf ActualizarLineaPedidoDeIntercambio, data, services)
                                    ProcessServer.ExecuteTask(Of DataEntradaStockEnDepositoOAlquiler)(AddressOf EntradaStockDeIntercambio, desda, services)
                                Else
                                    Dim updateEntrada As StockUpdateData = ProcessServer.ExecuteTask(Of DataEntradaStockEnDepositoOAlquiler, StockUpdateData)(AddressOf EntradaStockEnDepositoOAlquiler, desda, services)
                                    If Not updateEntrada Is Nothing Then
                                        ReDim Preserve aStkUD(aStkUD.Length)
                                        aStkUD(aStkUD.Length - 1) = updateEntrada
                                    End If
                                End If
                            End If
                    End Select
                    If updateData.Estado = EstadoStock.NoActualizado Then
                        DeshacerActualizados = True
                    End If
                    ReDim Preserve aStkUD(aStkUD.Length)
                    aStkUD(aStkUD.Length - 1) = updateData
                End If
            Next
            If DeshacerActualizados Then

                For i As Integer = 0 To aStkUD.Length - 1
                    If aStkUD(i).Estado = EstadoStock.Actualizado Then
                        aStkUD(i).Estado = EstadoStock.NoActualizado
                        Dim datMsg As New DataMessage(48)
                        aStkUD(i).Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, datMsg, services)
                    End If
                Next

                Select Case data.Circuito
                    Case Circuito.Compras
                        For Each lineaLote As DataRow In data.LotesLineaAlbaran.Rows
                            lineaLote("IDMovimientoEntrada") = System.DBNull.Value
                        Next
                    Case Circuito.Ventas
                        For Each lineaLote As DataRow In data.LotesLineaAlbaran.Rows
                            lineaLote("IDMovimientoSalida") = System.DBNull.Value
                        Next

                End Select
                data.LineaAlbaran("EstadoStock") = enumavlEstadoStock.avlNoActualizado

                ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.RollbackTransaction, True, services)
            Else
                ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.CommitTransaction, True, services)
            End If
        End If
        Return aStkUD
    End Function

    Public Class DataActualizarStockLineas
        Public DocumentoAlbaran As DocumentCabLin
        Public IDLineasAlbaran() As Integer
        Public Sub New(ByVal DocumentoAlbaran As DocumentCabLin, ByVal IDLineasAlbaran() As Integer)
            Me.DocumentoAlbaran = DocumentoAlbaran
            Me.IDLineasAlbaran = IDLineasAlbaran
        End Sub
        Public Sub New(ByVal DocumentoAlbaran As DocumentCabLin)
            Me.DocumentoAlbaran = DocumentoAlbaran
        End Sub
    End Class

    Public Class DataActualizarLineas
        Public stkUpdateData As StockUpdateData
        Public LineaAlbaran As DataRow
        Public LineaLote As DataRow

        Public Sub New(ByVal stkUpdateData As StockUpdateData, ByVal LineaAlbaran As DataRow)
            Me.stkUpdateData = stkUpdateData
            Me.LineaAlbaran = LineaAlbaran
        End Sub
        Public Sub New(ByVal stkUpdateData As StockUpdateData, ByVal LineaAlbaran As DataRow, ByVal LineaLote As DataRow)
            Me.stkUpdateData = stkUpdateData
            Me.LineaAlbaran = LineaAlbaran
            Me.LineaLote = LineaLote
        End Sub
    End Class


    Public Class DataEstablecerTipoMovimiento
        Public stkData As StockData
        Public LineaAlbaran As DataRow
        Public LineasAlbaran As DataTable

        Public Sub New(ByVal stkData As StockData, ByVal LineaAlbaran As DataRow, ByVal LineasAlbaran As DataTable)
            Me.stkData = stkData
            Me.LineaAlbaran = LineaAlbaran
            Me.LineasAlbaran = LineasAlbaran
        End Sub
    End Class

#Region " Compras "

    <Task()> Public Shared Sub ActualizarLineasAC(ByVal data As DataActualizarLineas, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataActualizarLineas)(AddressOf ActualizarEstadoStockAC, data, services)
        ProcessServer.ExecuteTask(Of DataActualizarLineas)(AddressOf ActualizarMovimientoStockAC, data, services)
    End Sub
    <Task()> Public Shared Sub ActualizarEstadoStockAC(ByVal data As DataActualizarLineas, ByVal services As ServiceProvider)
        Select Case data.stkUpdateData.Estado
            Case EstadoStock.Actualizado
                data.LineaAlbaran("EstadoStock") = enumaclEstadoStock.aclActualizado
            Case EstadoStock.NoActualizado
                data.LineaAlbaran("EstadoStock") = enumaclEstadoStock.aclNoActualizado
            Case EstadoStock.SinGestion
                data.LineaAlbaran("EstadoStock") = enumaclEstadoStock.aclSinGestion
        End Select
    End Sub
    <Task()> Public Shared Sub ActualizarMovimientoStockAC(ByVal data As DataActualizarLineas, ByVal services As ServiceProvider)
        Select Case data.stkUpdateData.Estado
            Case EstadoStock.Actualizado
                If data.LineaLote Is Nothing Then
                    data.LineaAlbaran("IDMovimiento") = data.stkUpdateData.IDLineaMovimiento
                Else
                    data.LineaLote("IDMovimientoEntrada") = data.stkUpdateData.IDLineaMovimiento
                End If
            Case Else
                If data.LineaLote Is Nothing Then
                    data.LineaAlbaran("IDMovimiento") = System.DBNull.Value
                Else
                    data.LineaLote("IDMovimientoEntrada") = System.DBNull.Value
                End If
        End Select
    End Sub

    <Task()> Public Shared Function EstablecerTipoMovimientoAC(ByVal data As DataEstablecerTipoMovimiento, ByVal services As ServiceProvider) As StockData
        If data.LineaAlbaran Is Nothing OrElse data.stkData Is Nothing Then Exit Function
        '//Determinar el tipo de movimiento que por defecto es entrada de albarán
        If data.LineaAlbaran("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclNormal OrElse data.LineaAlbaran("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclKit Then
            data.stkData.TipoMovimiento = enumTipoMovimiento.tmEntAlbaranCompra
        ElseIf data.LineaAlbaran("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclSubcontratacion Then
            data.stkData.TipoMovimiento = enumTipoMovimiento.tmEntSubcontratacion
        ElseIf data.LineaAlbaran("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclComponente Then
            '//si la linea es componente, hay que determinar si pertenece a un kit o es un componente de subcontratación manual
            Dim lineaPadre() As DataRow = data.LineasAlbaran.Select("IDLineaAlbaran = " & data.LineaAlbaran("IDLineaPadre"))
            If Not lineaPadre Is Nothing AndAlso lineaPadre.Length > 0 Then
                If lineaPadre(0)("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclKit Then
                    data.stkData.TipoMovimiento = enumTipoMovimiento.tmEntAlbaranCompra
                ElseIf lineaPadre(0)("TipoLineaAlbaran") = enumaclTipoLineaAlbaran.aclSubcontratacion Then
                    data.stkData.Cantidad = data.stkData.Cantidad
                    data.stkData.TipoMovimiento = enumTipoMovimiento.tmSalSubcontratacion
                End If
            End If
        End If
        Return data.stkData
    End Function

    <Task()> Public Shared Sub PrepararActivoUltimaCompra(ByVal dr As DataRow, ByVal services As ServiceProvider)
        If Not dr Is Nothing Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(dr("IDArticulo"))
            If Not ArtInfo Is Nothing AndAlso Length(ArtInfo.IDArticulo) > 0 AndAlso ArtInfo.NSerieObligatorio AndAlso Length(dr("Lote")) > 0 Then
                Dim f As New Filter
                f.Add(New StringFilterItem("IDArticulo", dr("IDArticulo")))
                f.Add(New StringFilterItem("NSerie", dr("Lote")))

                Dim dtArtNSerie As DataTable = New BE.DataEngine().Filter("vFrmArticuloNSerie", f)
                If Not IsNothing(dtArtNSerie) AndAlso dtArtNSerie.Rows.Count > 0 AndAlso Length(dtArtNSerie.Rows(0)("IDActivo")) > 0 Then
                    Dim dtActivo As DataTable = New Activo().SelOnPrimaryKey(dtArtNSerie.Rows(0)("IDActivo"))
                    If Not IsNothing(dtActivo) AndAlso dtActivo.Rows.Count > 0 Then
                        dtActivo.Rows(0)("IDProveedor") = ArtInfo.IDProveedorUltimaCompra
                        dtActivo.Rows(0)("PrecioUltimaCompra") = ArtInfo.PrecioUltimaCompraA
                        BusinessHelper.UpdateTable(dtActivo)
                    End If
                End If
            End If
        End If
    End Sub

#End Region

#Region " Ventas "
    <Task()> Public Shared Sub ActualizarLineasAV(ByVal data As DataActualizarLineas, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataActualizarLineas)(AddressOf ActualizarEstadoStockAV, data, services)
        ProcessServer.ExecuteTask(Of DataActualizarLineas)(AddressOf ActualizarMovimientoStockAV, data, services)
        ProcessServer.ExecuteTask(Of DataActualizarLineas)(AddressOf ActualizarPrecioCosteAV, data, services)
    End Sub
    <Task()> Public Shared Sub ActualizarEstadoStockAV(ByVal data As DataActualizarLineas, ByVal services As ServiceProvider)
        Select Case data.stkUpdateData.Estado
            Case EstadoStock.Actualizado
                data.LineaAlbaran("EstadoStock") = enumavlEstadoStock.avlActualizado
            Case EstadoStock.NoActualizado
                data.LineaAlbaran("EstadoStock") = enumavlEstadoStock.avlNoActualizado
            Case EstadoStock.SinGestion
                data.LineaAlbaran("EstadoStock") = enumavlEstadoStock.avlSinGestion
        End Select
    End Sub
    <Task()> Public Shared Sub ActualizarMovimientoStockAV(ByVal data As DataActualizarLineas, ByVal services As ServiceProvider)
        Select Case data.stkUpdateData.Estado
            Case EstadoStock.Actualizado
                If data.LineaLote Is Nothing Then
                    data.LineaAlbaran("IDMovimiento") = data.stkUpdateData.IDLineaMovimiento
                Else
                    data.LineaLote("IDMovimientoSalida") = data.stkUpdateData.IDLineaMovimiento
                End If
            Case Else
                If data.LineaLote Is Nothing Then
                    data.LineaAlbaran("IDMovimiento") = System.DBNull.Value
                Else
                    data.LineaLote("IDMovimientoSalida") = System.DBNull.Value
                End If
        End Select
    End Sub

    <Task()> Public Shared Sub ActualizarPrecioCosteAV(ByVal data As DataActualizarLineas, ByVal services As ServiceProvider)
        If Not data.stkUpdateData.Movimientos Is Nothing AndAlso data.stkUpdateData.Movimientos.Rows.Count > 0 Then
            Select Case data.stkUpdateData.Estado
                Case EstadoStock.Actualizado
                    Dim dataTarifa As New DataCalculoTarifaComercial(data.LineaAlbaran("IDArticulo"), data.LineaAlbaran("QInterna"), data.stkUpdateData.Movimientos.Rows(0)("FechaDocumento"))
                    dataTarifa.IDAlmacen = data.LineaAlbaran("IDAlmacen")
                    dataTarifa.UDValoracion = CDbl(Nz(data.LineaAlbaran("UDValoracion"), 1))
                    ProcessServer.ExecuteTask(Of DataCalculoTarifaComercial)(AddressOf ProcesoComercial.TarifaCosteArticulo, dataTarifa, services)
                    If Not dataTarifa.DatosTarifa Is Nothing Then
                        data.LineaAlbaran("PrecioCosteA") = dataTarifa.DatosTarifa.PrecioCosteA
                        data.LineaAlbaran("PrecioCosteB") = dataTarifa.DatosTarifa.PrecioCosteB
                    End If
            End Select
        End If
    End Sub

    <Task()> Public Shared Function EstablecerTipoMovimientoAV(ByVal data As DataEstablecerTipoMovimiento, ByVal services As ServiceProvider) As StockData
        If data.LineaAlbaran Is Nothing OrElse data.stkData Is Nothing Then Exit Function
        data.stkData.TipoMovimiento = enumTipoMovimiento.tmSalAlbaranVenta
        Return data.stkData
    End Function

    <Serializable()> _
    Public Class DataAlmacenesDeposito
        Public IDAlmacenDeposito As String
        Public IDUbicacionDeposito As String
        Public IDTipoAlbaran As String
    End Class
    Public Class DataValidarAlmacenes
        Public IDTipoAlbaran As String
        Public IDAlmacenDeposito As String
        Public LineaAlbaran As DataRow
    End Class
    <Task()> Public Shared Function ValidarAlmacenes(ByVal data As DataValidarAlmacenes, ByVal services As ServiceProvider) As StockUpdateData
        Dim valorTipoAlbaran As enumTipoAlbaran = services.GetService(Of enumTipoAlbaran)()
        If valorTipoAlbaran = enumTipoAlbaran.Desconocido Then
            valorTipoAlbaran = ProcessServer.ExecuteTask(Of String, enumTipoAlbaran)(AddressOf ProcesoAlbaranVenta.ValidarTipoAlbaran, data.IDTipoAlbaran, services)
        End If

        'Establecer los almacenes de salida/entrada deacuerdo con el tipo de albaran y el tipo de articulo
        '(por defecto)
        Dim IDAlmacen As String = data.LineaAlbaran("IDAlmacen") & String.Empty
        Dim IDAlmacenDeposito As String = data.IDAlmacenDeposito
        Dim IDUbicacionDeposito As String = String.Empty

        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.LineaAlbaran("IDArticulo"))
        Dim ArticuloDePortes As Boolean = ArtInfo.ArticuloDePortes
        If ArticuloDePortes Then
            Dim AppParams As ParametroTesoreria = services.GetService(Of ParametroTesoreria)()
            IDAlmacenDeposito = AppParams.AlmacenPortes()
        End If

        If (valorTipoAlbaran = enumTipoAlbaran.Deposito OrElse valorTipoAlbaran = enumTipoAlbaran.RetornoAlquiler) Then
            If Length(IDAlmacenDeposito) = 0 Then
                Dim das As New DataLogActualizarStock("El almacén de depósito no existe o está vacío.", data.LineaAlbaran("IDArticulo"))
                Return ProcessServer.ExecuteTask(Of DataLogActualizarStock, StockUpdateData)(AddressOf LogActualizarStock, das, services)
            End If
        End If

        If Length(IDAlmacenDeposito) > 0 Then
            Dim FilAlmDep As New Filter
            FilAlmDep.Add(New StringFilterItem("IDAlmacen", IDAlmacenDeposito))
            FilAlmDep.Add(New BooleanFilterItem("Predeterminada", True))
            Dim dtAlmDep As DataTable = New AlmacenUbicacion().Filter(FilAlmDep)
            If Not dtAlmDep Is Nothing AndAlso dtAlmDep.Rows.Count > 0 Then
                IDUbicacionDeposito = dtAlmDep.Rows(0)("IDUbicacion") & String.Empty
            End If
        End If

        If valorTipoAlbaran = enumTipoAlbaran.Deposito Then
            If IDAlmacen = IDAlmacenDeposito Then
                Dim das As New DataLogActualizarStock("El almacén de salida coincide con el almacén de depósito.", data.LineaAlbaran("IDArticulo"), IDAlmacen)
                Return ProcessServer.ExecuteTask(Of DataLogActualizarStock, StockUpdateData)(AddressOf LogActualizarStock, das, services)
            End If
        End If

        Dim Alm As DataAlmacenesDeposito = services.GetService(Of DataAlmacenesDeposito)()
        Alm.IDAlmacenDeposito = IDAlmacenDeposito
        Alm.IDUbicacionDeposito = IDUbicacionDeposito
        Alm.IDTipoAlbaran = valorTipoAlbaran
    End Function

    Public Class DataEntradaStockEnDepositoOAlquiler
        Public IDCliente As String
        Public Salida As StockData
        Public NumeroMovimiento As Integer
        Public lineaAlbaran As DataRow
        Public UpdateSalida As StockUpdateData
        Public lineaLote As DataRow

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDCliente As String, ByVal NumeroMovimiento As Integer, ByVal Salida As StockData, ByVal lineaAlbaran As DataRow, ByVal UpdateSalida As StockUpdateData)
            Me.IDCliente = IDCliente
            Me.NumeroMovimiento = NumeroMovimiento
            Me.Salida = Salida
            Me.lineaAlbaran = lineaAlbaran
            Me.UpdateSalida = UpdateSalida
        End Sub
    End Class

    <Task()> Public Shared Function EntradaStockEnDepositoOAlquiler(ByVal data As DataEntradaStockEnDepositoOAlquiler, ByVal services As ServiceProvider) As StockUpdateData
        Dim Alm As DataAlmacenesDeposito = services.GetService(Of DataAlmacenesDeposito)()
        If Alm.IDTipoAlbaran = enumTipoAlbaran.Deposito OrElse Alm.IDTipoAlbaran = enumTipoAlbaran.RetornoAlquiler Then
            If Not (Length(data.lineaAlbaran("IDObra")) > 0 AndAlso data.lineaAlbaran("TipoFactAlquiler") = enumTipoFacturacionAlquiler.enumTFASinAlquiler) Then
                '//Movimientos de entrada
                Dim stkEntrada As StockData
                Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Salida.Articulo)

                If Not ArtInfo Is Nothing AndAlso Length(ArtInfo.IDArticulo) > 0 Then
                    If ArtInfo.RecalcularValoracion = enumtaValoracionSalidas.taMantenerPrecio Then
                        ' El precio de la entrada en este caso será según criterio de valoración del artículo en el almacén de donde sale.
                        Dim PrecioEntradaA As Double
                        Dim PrecioEntradaB As Double
                        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                        Dim MonInfoB As MonedaInfo = Monedas.MonedaB
                        Select Case ArtInfo.CriterioValoracion
                            Case enumtaValoracion.taPrecioEstandar
                                PrecioEntradaA = data.UpdateSalida.Movimientos.Rows(0)("PrecioEstandar")
                                PrecioEntradaB = xRound(data.UpdateSalida.Movimientos.Rows(0)("PrecioEstandar") * Monedas.MonedaB.CambioB, MonInfoB.NDecimalesPrecio)
                            Case enumtaValoracion.taPrecioFIFOFecha
                                PrecioEntradaA = data.UpdateSalida.Movimientos.Rows(0)("FifoFD")
                                PrecioEntradaB = xRound(data.UpdateSalida.Movimientos.Rows(0)("FifoFD") * Monedas.MonedaB.CambioB, MonInfoB.NDecimalesPrecio)
                            Case enumtaValoracion.taPrecioFIFOMvto
                                PrecioEntradaA = data.UpdateSalida.Movimientos.Rows(0)("FifoF")
                                PrecioEntradaB = xRound(data.UpdateSalida.Movimientos.Rows(0)("FifoF") * Monedas.MonedaB.CambioB, MonInfoB.NDecimalesPrecio)
                            Case enumtaValoracion.taPrecioMedio
                                PrecioEntradaA = data.UpdateSalida.Movimientos.Rows(0)("PrecioMedio")
                                PrecioEntradaB = xRound(data.UpdateSalida.Movimientos.Rows(0)("PrecioMedio") * Monedas.MonedaB.CambioB, MonInfoB.NDecimalesPrecio)
                            Case enumtaValoracion.taPrecioUltCompra
                                PrecioEntradaA = data.UpdateSalida.Movimientos.Rows(0)("PrecioUltimaCompra")
                                PrecioEntradaB = xRound(data.UpdateSalida.Movimientos.Rows(0)("PrecioUltimaCompra") * Monedas.MonedaB.CambioB, MonInfoB.NDecimalesPrecio)
                        End Select

                        stkEntrada = New StockData(data.Salida.Articulo, Alm.IDAlmacenDeposito, data.Salida.Cantidad, PrecioEntradaA, PrecioEntradaB, _
                                                        data.Salida.FechaDocumento, enumTipoMovimiento.tmEntTransferencia, data.Salida.Documento, data.Salida.IDDocumento)
                    Else
                        stkEntrada = New StockData(data.Salida.Articulo, Alm.IDAlmacenDeposito, data.Salida.Cantidad, data.UpdateSalida.Movimientos.Rows(0)("precioA"), data.UpdateSalida.Movimientos.Rows(0)("precioB"), _
                                    data.Salida.FechaDocumento, enumTipoMovimiento.tmEntTransferencia, data.Salida.Documento, data.Salida.IDDocumento)
                    End If
                End If
                stkEntrada.Texto = data.Salida.Texto
                stkEntrada.Lote = data.Salida.Lote
                stkEntrada.Ubicacion = Alm.IDUbicacionDeposito
                stkEntrada.NSerie = data.Salida.NSerie
                stkEntrada.EstadoNSerie = data.Salida.EstadoNSerie
                stkEntrada.EstadoNSerieAnterior = data.Salida.EstadoNSerieAnterior
                stkEntrada.Operario = data.Salida.Operario
                stkEntrada.Obra = data.Salida.Obra

                Dim datMovto As New StEntradaTransfer(data.NumeroMovimiento, stkEntrada, data.Salida, data.UpdateSalida)
                Dim updateEntrada As StockUpdateData = ProcessServer.ExecuteTask(Of StEntradaTransfer, StockUpdateData)(AddressOf EntradaTransferencia, datMovto, services)
                If updateEntrada.Estado = EstadoStock.Actualizado Then
                    If Not data.lineaLote Is Nothing Then
                        data.lineaLote("IDMovimientoEntrada") = updateEntrada.IDLineaMovimiento
                    Else
                        data.lineaAlbaran("IDMovimientoEntrada") = updateEntrada.IDLineaMovimiento
                    End If
                    '//Mantenimiento correctivo-preventivo
                    If Len(stkEntrada.EstadoNSerie) > 0 Then
                        Dim addMnto As New ProcesoAlbaranVenta.DataAddMntoOT(data.Salida.EstadoNSerie, data.IDCliente, data.lineaAlbaran)
                        ProcessServer.ExecuteTask(Of ProcesoAlbaranVenta.DataAddMntoOT)(AddressOf ProcesoAlbaranVenta.AddMntoOT, addMnto, services)
                    End If
                End If
                '//Fin movimientos de entrada
                Return updateEntrada
            Else
                'David Velasco Herrero 15/06/22 Para crear movimiento de ferreteria
                '//Movimientos de entrada
                Dim stkEntrada As StockData
                Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Salida.Articulo)

                If Not ArtInfo Is Nothing AndAlso Length(ArtInfo.IDArticulo) > 0 Then
                    If ArtInfo.RecalcularValoracion = enumtaValoracionSalidas.taMantenerPrecio Then
                        ' El precio de la entrada en este caso será según criterio de valoración del artículo en el almacén de donde sale.
                        Dim PrecioEntradaA As Double
                        Dim PrecioEntradaB As Double
                        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                        Dim MonInfoB As MonedaInfo = Monedas.MonedaB
                        Select Case ArtInfo.CriterioValoracion
                            Case enumtaValoracion.taPrecioEstandar
                                PrecioEntradaA = data.UpdateSalida.Movimientos.Rows(0)("PrecioEstandar")
                                PrecioEntradaB = xRound(data.UpdateSalida.Movimientos.Rows(0)("PrecioEstandar") * Monedas.MonedaB.CambioB, MonInfoB.NDecimalesPrecio)
                            Case enumtaValoracion.taPrecioFIFOFecha
                                PrecioEntradaA = data.UpdateSalida.Movimientos.Rows(0)("FifoFD")
                                PrecioEntradaB = xRound(data.UpdateSalida.Movimientos.Rows(0)("FifoFD") * Monedas.MonedaB.CambioB, MonInfoB.NDecimalesPrecio)
                            Case enumtaValoracion.taPrecioFIFOMvto
                                PrecioEntradaA = data.UpdateSalida.Movimientos.Rows(0)("FifoF")
                                PrecioEntradaB = xRound(data.UpdateSalida.Movimientos.Rows(0)("FifoF") * Monedas.MonedaB.CambioB, MonInfoB.NDecimalesPrecio)
                            Case enumtaValoracion.taPrecioMedio
                                PrecioEntradaA = data.UpdateSalida.Movimientos.Rows(0)("PrecioMedio")
                                PrecioEntradaB = xRound(data.UpdateSalida.Movimientos.Rows(0)("PrecioMedio") * Monedas.MonedaB.CambioB, MonInfoB.NDecimalesPrecio)
                            Case enumtaValoracion.taPrecioUltCompra
                                PrecioEntradaA = data.UpdateSalida.Movimientos.Rows(0)("PrecioUltimaCompra")
                                PrecioEntradaB = xRound(data.UpdateSalida.Movimientos.Rows(0)("PrecioUltimaCompra") * Monedas.MonedaB.CambioB, MonInfoB.NDecimalesPrecio)
                        End Select

                        stkEntrada = New StockData(data.Salida.Articulo, Alm.IDAlmacenDeposito, data.Salida.Cantidad, PrecioEntradaA, PrecioEntradaB, _
                                                        data.Salida.FechaDocumento, enumTipoMovimiento.tmEntTransferencia, data.Salida.Documento, data.Salida.IDDocumento)
                    Else
                        stkEntrada = New StockData(data.Salida.Articulo, Alm.IDAlmacenDeposito, data.Salida.Cantidad, data.UpdateSalida.Movimientos.Rows(0)("precioA"), data.UpdateSalida.Movimientos.Rows(0)("precioB"), _
                                    data.Salida.FechaDocumento, enumTipoMovimiento.tmEntTransferencia, data.Salida.Documento, data.Salida.IDDocumento)
                    End If
                End If
                stkEntrada.Texto = data.Salida.Texto
                stkEntrada.Lote = data.Salida.Lote
                stkEntrada.Ubicacion = Alm.IDUbicacionDeposito
                stkEntrada.NSerie = data.Salida.NSerie
                stkEntrada.EstadoNSerie = data.Salida.EstadoNSerie
                stkEntrada.EstadoNSerieAnterior = data.Salida.EstadoNSerieAnterior
                stkEntrada.Operario = data.Salida.Operario
                stkEntrada.Obra = data.Salida.Obra

                Dim datMovto As New StEntradaTransfer(data.NumeroMovimiento, stkEntrada, data.Salida, data.UpdateSalida)
                Dim updateEntrada As StockUpdateData = ProcessServer.ExecuteTask(Of StEntradaTransfer, StockUpdateData)(AddressOf EntradaTransferencia, datMovto, services)
                If updateEntrada.Estado = EstadoStock.Actualizado Then
                    If Not data.lineaLote Is Nothing Then
                        data.lineaLote("IDMovimientoEntrada") = updateEntrada.IDLineaMovimiento
                    Else
                        data.lineaAlbaran("IDMovimientoEntrada") = updateEntrada.IDLineaMovimiento
                    End If
                    '//Mantenimiento correctivo-preventivo
                    If Len(stkEntrada.EstadoNSerie) > 0 Then
                        Dim addMnto As New ProcesoAlbaranVenta.DataAddMntoOT(data.Salida.EstadoNSerie, data.IDCliente, data.lineaAlbaran)
                        ProcessServer.ExecuteTask(Of ProcesoAlbaranVenta.DataAddMntoOT)(AddressOf ProcesoAlbaranVenta.AddMntoOT, addMnto, services)
                    End If
                End If
                '//Fin movimientos de entrada
                Return updateEntrada
            End If
        End If
    End Function

    <Task()> Public Shared Sub EntradaStockDeIntercambio(ByVal data As DataEntradaStockEnDepositoOAlquiler, ByVal services As ServiceProvider)
        Dim almacenDeposito As DataAlmacenesDeposito = services.GetService(Of DataAlmacenesDeposito)()
        If almacenDeposito.IDTipoAlbaran = enumTipoAlbaran.Intercambio Then
            '//Movimientos de entrada
            Dim stkEntrada As New StockData(data.Salida.Articulo, almacenDeposito.IDAlmacenDeposito, data.Salida.Cantidad, data.Salida.PrecioA, data.Salida.PrecioB, _
                                            data.Salida.FechaDocumento, enumTipoMovimiento.tmEntTransferencia, data.Salida.Documento)
            stkEntrada.Texto = data.Salida.Texto
            stkEntrada.Lote = data.Salida.Lote
            stkEntrada.Ubicacion = almacenDeposito.IDUbicacionDeposito
            stkEntrada.NSerie = data.Salida.NSerie
            stkEntrada.EstadoNSerie = data.Salida.EstadoNSerie
            stkEntrada.EstadoNSerieAnterior = data.Salida.EstadoNSerieAnterior
            stkEntrada.Operario = data.Salida.Operario
            stkEntrada.Obra = data.Salida.Obra

            Dim datMovto As New DataNumeroMovimientoSinc(data.NumeroMovimiento, stkEntrada)
            Dim updateEntrada As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf Entrada, datMovto, services)
            If updateEntrada.Estado = EstadoStock.Actualizado Then
                data.lineaAlbaran("IDMovimientoEntrada") = updateEntrada.IDLineaMovimiento
            End If
            '//Fin movimientos de entrada
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarLineaPedidoDeIntercambio(ByVal data As DataActualizarStockAlbaranTx, ByVal services As ServiceProvider)
        Dim Pedidos As DocumentInfoCache(Of DocumentoPedidoVenta) = services.GetService(Of DocumentInfoCache(Of DocumentoPedidoVenta))()
        Dim DocPed As DocumentoPedidoVenta = Pedidos.GetDocument(data.LineaAlbaran("IDPedido"))
        DocPed.SetQAlbaran(data.LineaAlbaran("IDLineaPedido"), 0, services)
    End Sub

    Public Class DataActualizarStockContenedores
        Public CabeceraAlbaran As DataRow
        Public LineaAlbaran As DataRow

        Public Sub New(ByVal CabeceraAlbaran As DataRow, ByVal LineaAlbaran As DataRow)
            Me.CabeceraAlbaran = CabeceraAlbaran
            Me.LineaAlbaran = LineaAlbaran
        End Sub
    End Class
    <Task()> Public Shared Function GestionContenedores(ByVal data As DataActualizarStockContenedores, ByVal services As ServiceProvider) As StockUpdateData()
        If Length(data.LineaAlbaran("IDArticuloContenedor")) > 0 Then
            Dim updateDataArray(-1) As StockUpdateData
            Dim updateContenedores() As StockUpdateData = ProcessServer.ExecuteTask(Of DataActualizarStockContenedores, StockUpdateData())(AddressOf ActualizarStockContenedores, data, services)
            If Not updateContenedores Is Nothing AndAlso updateContenedores.Length > 0 Then
                'Primer elemento corresponde a la salida
                'Segundo elemento corresponde a la entrada
                If Not updateContenedores(0) Is Nothing Then
                    ReDim Preserve updateDataArray(updateDataArray.Length)
                    updateDataArray(updateDataArray.Length - 1) = updateContenedores(0)
                End If
                If Not updateContenedores(1) Is Nothing Then
                    ReDim Preserve updateDataArray(updateDataArray.Length)
                    updateDataArray(updateDataArray.Length - 1) = updateContenedores(1)
                End If
            End If
            Return updateDataArray
        End If
    End Function
    <Task()> Public Shared Function ActualizarStockContenedores(ByVal data As DataActualizarStockContenedores, ByVal services As ServiceProvider) As StockUpdateData()
        Dim updateDataArray(1) As StockUpdateData

        Dim updateSalidaContenedor As StockUpdateData
        Dim updateEntradaContenedor As StockUpdateData

        If Length(data.LineaAlbaran("IDArticuloContenedor")) > 0 Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.LineaAlbaran("IDArticuloContenedor"))
            If Not ArtInfo Is Nothing AndAlso Length(ArtInfo.IDArticulo) > 0 AndAlso ArtInfo.GestionStock Then
                Dim NumeroMovimiento As Integer
                If IsNumeric(data.CabeceraAlbaran("NMovimiento")) AndAlso Not data.CabeceraAlbaran("NMovimiento") = 0 Then
                    NumeroMovimiento = data.CabeceraAlbaran("NMovimiento")
                Else
                    NumeroMovimiento = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf NuevoNumeroMovimiento, Nothing, services)
                    data.CabeceraAlbaran("NMovimiento") = NumeroMovimiento
                End If
                Dim Fecha As Date = data.CabeceraAlbaran("FechaAlbaran")

                Dim ArticuloContenedor As String = data.LineaAlbaran("IDArticuloContenedor")
                Dim AlmacenContenedor As String
                Dim AlmacenPredeterminado As String

                Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
                Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data.CabeceraAlbaran("IDCliente"))
                If Not ClteInfo Is Nothing AndAlso Length(ClteInfo.IDCliente) > 0 Then
                    If Length(ClteInfo.IDAlmacenContenedor) > 0 Then
                        AlmacenContenedor = ClteInfo.IDAlmacenContenedor
                    End If
                End If
                If Len(AlmacenContenedor) = 0 Then
                    Dim das As New DataLogActualizarStock("El cliente " & Quoted(data.CabeceraAlbaran("IDCliente")) & " no tiene asignado ningún almacén contenedor.", ArticuloContenedor)
                    updateDataArray(0) = ProcessServer.ExecuteTask(Of DataLogActualizarStock, StockUpdateData)(AddressOf LogActualizarStock, das, services)
                    Return updateDataArray
                Else
                    Dim QContenedor As Double = Nz(data.LineaAlbaran("QEtiContenedor"), 0)
                    Dim f As New Filter
                    f.Add(New StringFilterItem("IDArticulo", FilterOperator.Equal, ArticuloContenedor))
                    f.Add(New BooleanFilterItem("Predeterminado", FilterOperator.Equal, True))
                    Dim almacen As DataTable = New Negocio.ArticuloAlmacen().Filter(f)
                    If Not almacen Is Nothing AndAlso almacen.Rows.Count > 0 Then
                        AlmacenPredeterminado = almacen.Rows(0)("IDAlmacen")
                    End If
                    If Len(AlmacenPredeterminado) = 0 Then
                        Dim das As New DataLogActualizarStock("El cliente " & Quoted(data.CabeceraAlbaran("IDCliente")) & " no tiene asignado ningún almacén predeterminado.", ArticuloContenedor)
                        updateDataArray(0) = ProcessServer.ExecuteTask(Of DataLogActualizarStock, StockUpdateData)(AddressOf LogActualizarStock, das, services)
                        Return updateDataArray
                    Else
                        ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)

                        '//Salida de contenedor
                        If Length(data.LineaAlbaran("IDSalidaContenedor")) = 0 Then
                            Dim salidaContenedor As New StockData(ArticuloContenedor, AlmacenPredeterminado, QContenedor, 0, 0, Fecha, enumTipoMovimiento.tmSalTransferencia, data.CabeceraAlbaran("NAlbaran"))
                            salidaContenedor.Texto = data.LineaAlbaran("Texto") & String.Empty
                            If IsNumeric(data.LineaAlbaran("IDObra")) Then
                                salidaContenedor.Obra = data.LineaAlbaran("IDObra")
                            End If

                            Dim dataSal As New DataNumeroMovimiento(NumeroMovimiento, salidaContenedor)
                            updateSalidaContenedor = ProcessServer.ExecuteTask(Of DataNumeroMovimiento, StockUpdateData)(AddressOf Salida, dataSal, services)

                            If updateSalidaContenedor.Estado = EstadoStock.Actualizado Then
                                '//Entrada en el almacen del cliente
                                Dim entradaContenedor As New StockData(ArticuloContenedor, AlmacenContenedor, QContenedor, 0, 0, Fecha, enumTipoMovimiento.tmEntTransferencia, data.CabeceraAlbaran("NAlbaran"))
                                entradaContenedor.Texto = data.LineaAlbaran("Texto") & String.Empty
                                If Length(data.LineaAlbaran("IDObra")) > 0 Then
                                    entradaContenedor.Obra = data.LineaAlbaran("IDObra")
                                End If

                                Dim datMovto As New DataNumeroMovimiento(NumeroMovimiento, entradaContenedor)
                                updateEntradaContenedor = ProcessServer.ExecuteTask(Of DataNumeroMovimiento, StockUpdateData)(AddressOf Entrada, datMovto, services)
                                If updateEntradaContenedor.Estado = EstadoStock.Actualizado Then
                                    data.LineaAlbaran("IDEntradaContenedor") = updateEntradaContenedor.IDLineaMovimiento
                                    data.LineaAlbaran("IDSalidaContenedor") = updateSalidaContenedor.IDLineaMovimiento
                                Else
                                    ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.RollbackTransaction, True, services)
                                    data.LineaAlbaran("EstadoStock") = enumaclEstadoStock.aclNoActualizado
                                End If
                                '///Fin entrada en el almacen del cliente
                            Else
                                data.LineaAlbaran("EstadoStock") = enumaclEstadoStock.aclNoActualizado
                            End If
                            '//Fin Salida de contenedor

                            ReDim Preserve updateDataArray(updateDataArray.Length)
                            updateDataArray(updateDataArray.Length - 1) = updateSalidaContenedor
                            ReDim Preserve updateDataArray(updateDataArray.Length)
                            updateDataArray(updateDataArray.Length - 1) = updateEntradaContenedor
                        End If
                    End If
                End If
            End If
        End If

        Return updateDataArray
    End Function

#End Region

#End Region

#Region " Valoración de Almacenes "

    <Serializable()> _
    Public Class ValoracionResultInfo
        Public ArticuloAlmacen As DataTable
        Public Valoraciones(-1) As ValoracionInfo
    End Class

#Region "Precio Movimientos"

    <Serializable()> _
    Public Class DataPrecioMovimiento
        Public IDArticulo As String
        Public IDAlmacen As String
        Public FechaDocumento As Date
        Public Cantidad As Double
        Public ClaseMovimiento As enumtpmTipoMovimiento
        Public PrecioA As Double
        Public Movimiento As DataTable

        Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal ClaseMovimiento As enumtpmTipoMovimiento)
            Me.IDArticulo = IDArticulo
            Me.IDAlmacen = IDAlmacen
            Me.FechaDocumento = cnMinDate
            Me.Cantidad = 0
            Me.ClaseMovimiento = ClaseMovimiento
        End Sub

        Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal FechaDocumento As Date, ByVal Cantidad As Double, ByVal ClaseMovimiento As enumtpmTipoMovimiento, Optional ByVal PrecioA As Double = 0)
            Me.IDArticulo = IDArticulo
            Me.IDAlmacen = IDAlmacen
            Me.FechaDocumento = FechaDocumento
            Me.Cantidad = Cantidad
            Me.ClaseMovimiento = ClaseMovimiento
            Me.PrecioA = PrecioA
        End Sub
    End Class
    <Task()> Public Shared Function PrecioMovimiento(ByVal data As DataPrecioMovimiento, ByVal services As ServiceProvider) As Hashtable
        Dim precios As New Hashtable
        precios("PrecioA") = 0
        precios("PrecioB") = 0

        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.MonedaA
        Dim MonInfoB As MonedaInfo = Monedas.MonedaB

        If Len(data.IDArticulo) > 0 AndAlso Len(data.IDAlmacen) > 0 Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.IDArticulo)
            If Not ArtInfo Is Nothing AndAlso Length(ArtInfo.IDArticulo) > 0 Then
                If data.ClaseMovimiento = enumtpmTipoMovimiento.tpmOutput AndAlso ArtInfo.RecalcularValoracion = enumtaValoracionSalidas.taMantenerPrecio AndAlso data.PrecioA <> 0 Then
                    Return Nothing
                Else
                    Dim f As New Filter
                    f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
                    f.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))

                    Dim dtArticuloAlmacen As DataTable = New ArticuloAlmacen().Filter(f)
                    If Not dtArticuloAlmacen Is Nothing AndAlso dtArticuloAlmacen.Rows.Count > 0 Then
                        Select Case ArtInfo.CriterioValoracion
                            Case enumtaValoracion.taPrecioEstandar
                                precios("PrecioA") = ArtInfo.PrecioEstandarA / Nz(ArtInfo.UDValoracion, 1)
                                precios("PrecioB") = ArtInfo.PrecioEstandarB / Nz(ArtInfo.UDValoracion, 1)
                            Case enumtaValoracion.taPrecioFIFOFecha
                                If data.FechaDocumento = cnMinDate Then data.FechaDocumento = Date.Today
                                If data.ClaseMovimiento = enumtpmTipoMovimiento.tpmOutput AndAlso ArtInfo.RecalcularValoracion = enumtaValoracionSalidas.taNoRecalcular Then
                                    precios("PrecioA") = dtArticuloAlmacen.Rows(0)("PrecioFIFOFechaA")
                                    precios("PrecioB") = dtArticuloAlmacen.Rows(0)("PrecioFIFOFechaB")
                                ElseIf data.ClaseMovimiento <> enumtpmTipoMovimiento.tpmOutput OrElse ArtInfo.RecalcularValoracion = enumtaValoracionSalidas.taRecalcular OrElse ArtInfo.RecalcularValoracion = enumtaValoracionSalidas.taMantenerPrecio Then
                                    Dim datosPrecio As New DataValoracionFIFO(data.IDArticulo, data.IDAlmacen, dtArticuloAlmacen.Rows(0)("StockFisico"), data.Cantidad, data.FechaDocumento, enumstkValoracionFIFO.stkVFOrdenarPorFecha)
                                    Dim valoracion As ValoracionPreciosInfo = ProcessServer.ExecuteTask(Of DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ValoracionFIFO, datosPrecio, services)
                                    precios("PrecioA") = valoracion.PrecioA
                                    precios("PrecioB") = valoracion.PrecioB
                                End If
                            Case enumtaValoracion.taPrecioFIFOMvto
                                If data.FechaDocumento = cnMinDate Then data.FechaDocumento = Date.Today
                                If data.ClaseMovimiento = enumtpmTipoMovimiento.tpmOutput AndAlso ArtInfo.RecalcularValoracion = enumtaValoracionSalidas.taNoRecalcular Then
                                    precios("PrecioA") = dtArticuloAlmacen.Rows(0)("PrecioFIFOMvtoA")
                                    precios("PrecioB") = dtArticuloAlmacen.Rows(0)("PrecioFIFOMvtoB")
                                ElseIf data.ClaseMovimiento <> enumtpmTipoMovimiento.tpmOutput OrElse ArtInfo.RecalcularValoracion = enumtaValoracionSalidas.taRecalcular OrElse ArtInfo.RecalcularValoracion = enumtaValoracionSalidas.taMantenerPrecio Then
                                    Dim datosPrecio As New DataValoracionFIFO(data.IDArticulo, data.IDAlmacen, dtArticuloAlmacen.Rows(0)("StockFisico"), data.Cantidad, data.FechaDocumento, enumstkValoracionFIFO.stkVFOrdenarPorMvto)
                                    Dim valoracion As ValoracionPreciosInfo = ProcessServer.ExecuteTask(Of DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ValoracionFIFO, datosPrecio, services)
                                    precios("PrecioA") = valoracion.PrecioA
                                    precios("PrecioB") = valoracion.PrecioB
                                End If
                            Case enumtaValoracion.taPrecioMedio
                                If data.FechaDocumento = cnMinDate Then data.FechaDocumento = Date.Today
                                Dim valoracion As ValoracionPreciosInfo
                                Select Case data.ClaseMovimiento
                                    Case enumtpmTipoMovimiento.tpmInput
                                        If Not data.Movimiento Is Nothing AndAlso data.Movimiento.Rows.Count > 0 Then
                                            Dim DataCalc As New DataCalcValPrecioMedio(data.Movimiento.Rows(0), data.ClaseMovimiento)
                                            Dim PrecioMedio As Double = xRound(Nz(ProcessServer.ExecuteTask(Of DataCalcValPrecioMedio, Double)(AddressOf CalculoValAlmPrecioMedio, DataCalc, services), 0), MonInfoA.NDecimalesPrecio)
                                            precios("PrecioA") = xRound(PrecioMedio, MonInfoA.NDecimalesPrecio)
                                            precios("PrecioB") = xRound(PrecioMedio * Monedas.MonedaB.CambioB, MonInfoB.NDecimalesPrecio)
                                        Else
                                            Dim datosValPMF As New DataArticuloAlmacenFecha(data.IDArticulo, data.IDAlmacen, data.FechaDocumento)
                                            valoracion = ProcessServer.ExecuteTask(Of DataArticuloAlmacenFecha, ValoracionPreciosInfo)(AddressOf ValoracionPrecioMedioAFecha, datosValPMF, services)
                                            precios("PrecioA") = valoracion.PrecioA
                                            precios("PrecioB") = valoracion.PrecioB
                                        End If
                                    Case enumtpmTipoMovimiento.tpmInventario
                                        Dim datosValPMInv As New DataArticuloAlmacenFecha(data.IDArticulo, data.IDAlmacen, data.FechaDocumento)
                                        valoracion = ProcessServer.ExecuteTask(Of DataArticuloAlmacenFecha, ValoracionPreciosInfo)(AddressOf PrecioMedioInventario, datosValPMInv, services)
                                        precios("PrecioA") = valoracion.PrecioA
                                        precios("PrecioB") = valoracion.PrecioB
                                    Case enumtpmTipoMovimiento.tpmOutput
                                        If ArtInfo.RecalcularValoracion = enumtaValoracionSalidas.taNoRecalcular Then
                                            precios("PrecioA") = dtArticuloAlmacen.Rows(0)("PrecioMedioA")
                                            precios("PrecioB") = dtArticuloAlmacen.Rows(0)("PrecioMedioB")
                                        Else
                                            Dim datosValPMInv As New DataArticuloAlmacenFecha(data.IDArticulo, data.IDAlmacen, data.FechaDocumento)
                                            valoracion = ProcessServer.ExecuteTask(Of DataArticuloAlmacenFecha, ValoracionPreciosInfo)(AddressOf ValoracionPrecioMedioAFecha, datosValPMInv, services)
                                            precios("PrecioA") = valoracion.PrecioA
                                            precios("PrecioB") = valoracion.PrecioB
                                        End If
                                End Select

                            Case enumtaValoracion.taPrecioUltCompra
                                precios("PrecioA") = ArtInfo.PrecioUltimaCompraA
                                precios("PrecioB") = ArtInfo.PrecioUltimaCompraB
                        End Select
                    End If

                    If data.ClaseMovimiento <> enumtpmTipoMovimiento.tpmOutput AndAlso (precios("PrecioA") = 0 OrElse precios("PrecioB") = 0) Then
                        Dim AppParamsStocks As ParametroStocks = services.GetService(Of ParametroStocks)()
                        If Not AppParamsStocks.PrecioMovimientoCero Then
                            precios("PrecioA") = ArtInfo.PrecioEstandarA / Nz(ArtInfo.UDValoracion, 1)
                            precios("PrecioB") = ArtInfo.PrecioEstandarB / Nz(ArtInfo.UDValoracion, 1)
                        End If
                    End If
                End If
            End If


            precios("PrecioA") = xRound(precios("PrecioA"), MonInfoA.NDecimalesPrecio)
            precios("PrecioB") = xRound(precios("PrecioB"), MonInfoB.NDecimalesPrecio)
        End If

        Return precios
    End Function

#End Region

#Region " Stock a Fecha y Valoración Stock a Fecha "

    <Task()> Public Shared Function GetStockAcumuladoAFecha(ByVal data As DataArticuloAlmacenFecha, ByVal services As ServiceProvider) As StockAFechaInfo
        Dim valorStock As New StockAFechaInfo
        valorStock.IDArticulo = data.IDArticulo
        valorStock.IDAlmacen = data.IDAlmacen
        valorStock.StockAFecha = 0
        valorStock.StockAFecha2 = 0
        valorStock.FechaCalculo = data.Fecha
        If Len(data.IDArticulo) > 0 And Len(data.IDAlmacen) > 0 Then
            Dim ultimoMovimiento As DataRow = ProcessServer.ExecuteTask(Of DataArticuloAlmacenFecha, DataRow)(AddressOf ObtenerUltimoMovimientoAFecha, data, services)
            If Not ultimoMovimiento Is Nothing Then
                valorStock.StockAFecha = ultimoMovimiento("Acumulado")
                If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.IDArticulo, services) Then
                    valorStock.StockAFecha2 = Nz(ultimoMovimiento("Acumulado2"), 0)
                End If
                valorStock.PrecioMedio = Nz(ultimoMovimiento("PrecioMedio"), 0)
                valorStock.PrecioUltimaCompra = Nz(ultimoMovimiento("PrecioUltimaCompra"), 0)
                valorStock.PrecioEstandar = Nz(ultimoMovimiento("PrecioEstandar"), 0)
                valorStock.FifoF = Nz(ultimoMovimiento("FifoF"), 0)
                valorStock.FifoFD = Nz(ultimoMovimiento("FifoFD"), 0)
            End If
        End If

        Return valorStock
    End Function

    <Serializable()> _
    Public Class DataValoracionStockAFecha
        'Friend DatosAValorar As DataTable
        Public FechaCalculo As Date
        Public UDValoracion As String
        Public ArticuloAlmacen As DataTable
        'Friend Valoraciones As DataTable
        Public FechaInicioCalculo As Date
        Public CriterioValoracion As enumtaValoracion = -1

        Public Sub New(ByVal FechaCalculo As Date, ByVal ArticuloAlmacen As DataTable, Optional ByVal FechaInicioCalculo As Date = cnMinDate, Optional ByVal CriterioValoracion As enumtaValoracion = -1)
            Me.FechaCalculo = FechaCalculo
            Me.ArticuloAlmacen = ArticuloAlmacen
            If FechaInicioCalculo <> cnMinDate Then Me.FechaInicioCalculo = FechaInicioCalculo
            If CriterioValoracion <> -1 Then Me.CriterioValoracion = CriterioValoracion
        End Sub
    End Class

    <Task()> Public Shared Function GetValoracionStockAFecha(ByVal data As DataValoracionStockAFecha, ByVal services As ServiceProvider) As ValoracionInfo()
        Dim arrayValores(-1) As ValoracionInfo
        If Not data.ArticuloAlmacen Is Nothing AndAlso data.ArticuloAlmacen.Rows.Count > 0 Then
            If data.FechaCalculo <> cnMinDate Then
                Dim CriterioValArticulo As enumtaValoracion
                Dim f As New Filter
                Dim lstFactoresConversion As New Dictionary(Of String, Double)
                For Each dr As DataRow In data.ArticuloAlmacen.Rows
                    Dim StockAcum As New DataArticuloAlmacenFecha(dr("IDArticulo"), dr("IDAlmacen"), data.FechaCalculo)
                    Dim itemInfo As StockAFechaInfo = ProcessServer.ExecuteTask(Of DataArticuloAlmacenFecha, StockAFechaInfo)(AddressOf GetStockAcumuladoAFecha, StockAcum, services)
                    Dim val As New ValoracionInfo
                    val.IDArticulo = dr("IDArticulo")
                    val.IDAlmacen = dr("IDAlmacen")
                    val.FechaCalculo = data.FechaCalculo
                    val.Stock = itemInfo
                    Dim Factor As Double = 0
                    If Length(data.UDValoracion) > 0 Then
                        Dim Key As String = dr("IDArticulo") & ";" & dr("IDUDInterna") & String.Empty & ";" & data.UDValoracion
                        If lstFactoresConversion.ContainsKey(Key) Then
                            Factor = lstFactoresConversion(Key)
                        Else
                            Dim datFactor As New ArticuloUnidadAB.DatosFactorConversion(dr("IDArticulo"), dr("IDUDInterna") & String.Empty, data.UDValoracion, False)
                            Factor = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, datFactor, services)

                            lstFactoresConversion(Key) = Factor
                        End If
                    End If
                    val.Stock.StockAFechaUdValoracion = val.Stock.StockAFecha * Factor

                    Dim precios As New ValoracionPreciosInfo
                    'f.Clear()
                    'f.Add(New StringFilterItem("IDArticulo", dr("IDArticulo")))
                    'f.Add(New StringFilterItem("IDAlmacen", dr("IDAlmacen")))
                    'data.Valoraciones.DefaultView.RowFilter = f.Compose(New AdoFilterComposer)
                    'If data.Valoraciones.DefaultView.Count > 0 Then
                    If data.CriterioValoracion = -1 Then
                        CriterioValArticulo = dr("CriterioValoracion")
                    Else
                        CriterioValArticulo = data.CriterioValoracion
                    End If
                    Select Case CriterioValArticulo
                        Case enumtaValoracion.taPrecioEstandar
                            precios.IDArticulo = dr("IDArticulo")
                            precios.IDAlmacen = dr("IDAlmacen")
                            precios.FechaCalculo = data.FechaCalculo
                            precios.CriterioValoracion = enumtaValoracion.taPrecioEstandar
                            precios.PrecioA = dr("PrecioEstandarA")
                            precios.PrecioB = dr("PrecioEstandarB")
                        Case enumtaValoracion.taPrecioFIFOFecha
                            Dim datosPrecio As New DataValoracionFIFO(dr("IDArticulo"), dr("IDAlmacen"), itemInfo.StockAFecha, itemInfo.StockAFecha, data.FechaCalculo, enumstkValoracionFIFO.stkVFOrdenarPorFecha)
                            precios = ProcessServer.ExecuteTask(Of DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ValoracionFIFO, datosPrecio, services)
                        Case enumtaValoracion.taPrecioFIFOMvto
                            Dim datosPrecio As New DataValoracionFIFO(dr("IDArticulo"), dr("IDAlmacen"), itemInfo.StockAFecha, itemInfo.StockAFecha, data.FechaCalculo, enumstkValoracionFIFO.stkVFOrdenarPorMvto)
                            precios = ProcessServer.ExecuteTask(Of DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ValoracionFIFO, datosPrecio, services)
                        Case enumtaValoracion.taPrecioMedio
                            Dim datosValPM As New DataValoracionPrecioMedio(dr("IDArticulo"), dr("IDAlmacen"), data.FechaCalculo, data.FechaInicioCalculo)
                            precios = ProcessServer.ExecuteTask(Of DataValoracionPrecioMedio, ValoracionPreciosInfo)(AddressOf ValoracionPrecioMedio, datosValPM, services)
                        Case enumtaValoracion.taPrecioUltCompra
                            precios.IDArticulo = dr("IDArticulo")
                            precios.IDAlmacen = dr("IDAlmacen")
                            precios.FechaCalculo = data.FechaCalculo
                            precios.CriterioValoracion = enumtaValoracion.taPrecioUltCompra
                            precios.PrecioA = dr("PrecioUltimaCompraA")
                            precios.PrecioB = dr("PrecioUltimaCompraB")
                    End Select
                    'End If

                    val.Precios = precios

                    ''/////
                    dr("StockFisico") = itemInfo.StockAFecha ' val.Stock.StockAFecha
                    If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, dr("IDArticulo"), services) Then
                        dr("StockFisico2") = itemInfo.StockAFecha2
                    End If

                    dr("PrecioA") = val.Precios.PrecioA
                    dr("PrecioB") = val.Precios.PrecioB
                    ' If val.Stock.StockAFecha >= 0 Then
                    dr("ImporteA") = val.Stock.StockAFecha * val.Precios.PrecioA
                    dr("ImporteB") = val.Stock.StockAFecha * val.Precios.PrecioB
                    If dr.Table.Columns.Contains("StockAFechaUdValoracion") Then
                        dr("StockAFechaUdValoracion") = val.Stock.StockAFechaUdValoracion
                    End If

                    '' End If
                    '////

                    ReDim Preserve arrayValores(UBound(arrayValores) + 1)
                    arrayValores(UBound(arrayValores)) = val
                Next
            End If
        End If
        data.ArticuloAlmacen.AcceptChanges()
        Return arrayValores
    End Function

#End Region

#Region " Stock a fecha lote "

    <Serializable()> _
    Public Class DataGetStockAFechaLote
        Public Fecha As Date
        Public Stock As Double
        Public Stock2 As Double?
        Public Criteria As Filter

        Public Sub New(ByVal Fecha As Date, ByVal Stock As Double, Optional ByVal Stock2 As Double = Double.NaN)
            Me.Fecha = Fecha
            Me.Stock = Stock
            If Stock2 <> Double.NaN Then Me.Stock2 = Stock2
        End Sub
        Public Sub New(ByVal Fecha As Date, ByVal Stock As Double, ByVal Criteria As Filter, Optional ByVal Stock2 As Double = Double.NaN)
            Me.Fecha = Fecha
            Me.Stock = Stock
            Me.Criteria = Criteria
            If Stock2 <> Double.NaN Then Me.Stock2 = Stock2
        End Sub
    End Class
    <Task()> Public Shared Function GetStockAFechaLote(ByVal data As DataGetStockAFechaLote, ByVal services As ServiceProvider) As DataTable
        Dim AppParamsStocks As ParametroStocks = services.GetService(Of ParametroStocks)()
        Dim dttStock As DataTable = New BE.DataEngine().Filter("viewArticuloAlmacenLote", data.Criteria)
        Dim fMov As New Filter
        fMov = data.Criteria
        fMov.Add("FechaDocumento", FilterOperator.GreaterThanOrEqual, data.Fecha, FilterType.DateTime)
        Dim dttStockMov As DataTable = New BE.DataEngine().Filter("viewArticuloAlmacenLoteMovimientos", fMov)
        For Each row As DataRow In dttStockMov.Rows

            Dim FLote As New Filter
            FLote.Add("IdARticulo", FilterOperator.Equal, row("IDArticulo"), FilterType.String)
            FLote.Add("IDAlmacen", FilterOperator.Equal, row("IDAlmacen"), FilterType.String)
            FLote.Add("Lote", FilterOperator.Equal, row("Lote"), FilterType.String)
            FLote.Add("Ubicacion", FilterOperator.Equal, row("Ubicacion"), FilterType.String)
            Dim Filtro As String = FLote.Compose(New AdoFilterComposer)
            Dim LineaArticuloAlmacenLote() As DataRow = dttStock.Select(Filtro)

            Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, row("IDArticulo"), services)

            Dim F As New Filter
            F.Add("IdARticulo", FilterOperator.Equal, row("IDArticulo"), FilterType.String)
            F.Add("IDAlmacen", FilterOperator.Equal, row("IDAlmacen"), FilterType.String)
            F.Add("Lote", FilterOperator.Equal, row("Lote"), FilterType.String)
            F.Add("IdTipoMovimiento", FilterOperator.Equal, enumTipoMovimiento.tmInventario, FilterType.Numeric)
            F.Add("FechaDocumento", FilterOperator.GreaterThan, data.Fecha, FilterType.DateTime)
            Dim dttMov As DataTable = AdminData.GetEntityData(cnMyClass, F, "FechaDocumento ASC")
            If dttMov.Rows.Count > 0 Then

                LineaArticuloAlmacenLote(0)("StockFisico") = dttMov.Rows(0)("Cantidad")
                If SegundaUnidad Then row("StockFisico2") = dttMov.Rows(0)("Cantidad2")
                Select Case AppParamsStocks.TipoInventario
                    Case TipoInventario.UltimoMovimiento
                        F = New Filter
                        F.Add("IdARticulo", FilterOperator.Equal, row("IDArticulo"), FilterType.String)
                        F.Add("IDAlmacen", FilterOperator.Equal, row("IDAlmacen"), FilterType.String)
                        F.Add("Lote", FilterOperator.Equal, row("Lote"), FilterType.String)
                        F.Add("FechaDocumento", FilterOperator.Equal, dttMov.Rows(0)("FechaDocumento"), FilterType.DateTime)
                        F.Add("IdTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmInventario, FilterType.Numeric)
                        F.Add("IdTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion, FilterType.Numeric)
                        Dim dttMovDia As DataTable = AdminData.GetEntityData(cnMyClass, F, "FechaDocumento DESC,IDLineaMovimiento DESC")
                        For Each row1 As DataRow In dttMovDia.Rows
                            LineaArticuloAlmacenLote(0)("StockFisico") = Nz(LineaArticuloAlmacenLote(0)("StockFisico"), 0) - Nz(row1("Cantidad"), 0)
                            If SegundaUnidad Then LineaArticuloAlmacenLote(0)("StockFisico2") = Nz(LineaArticuloAlmacenLote(0)("StockFisico2"), 0) - Nz(row1("Cantidad2"), 0)
                        Next
                    Case TipoInventario.PrimerMovimiento
                End Select
                F = New Filter
                F.Add("IdARticulo", FilterOperator.Equal, row("IDArticulo"), FilterType.String)
                F.Add("IDAlmacen", FilterOperator.Equal, row("IDAlmacen"), FilterType.String)
                F.Add("Ubicacion", FilterOperator.Equal, row("Ubicacion"), FilterType.String)
                F.Add("Lote", FilterOperator.Equal, row("Lote"), FilterType.String)
                F.Add("FechaDocumento", FilterOperator.GreaterThan, data.Fecha, FilterType.DateTime)
                F.Add("FechaDocumento", FilterOperator.LessThan, dttMov.Rows(0)("FechaDocumento"), FilterType.DateTime)
                F.Add("IdTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion, FilterType.Numeric)
            Else
                F = New Filter
                F.Add("IdARticulo", FilterOperator.Equal, row("IDArticulo"), FilterType.String)
                F.Add("IDAlmacen", FilterOperator.Equal, row("IDAlmacen"), FilterType.String)
                F.Add("Ubicacion", FilterOperator.Equal, row("Ubicacion"), FilterType.String)
                F.Add("Lote", FilterOperator.Equal, row("Lote"), FilterType.String)
                F.Add("FechaDocumento", FilterOperator.GreaterThan, data.Fecha, FilterType.DateTime)
                F.Add("IdTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion, FilterType.Numeric)
            End If
            dttMov = AdminData.GetEntityData(cnMyClass, F, "FechaDocumento DESC,IDLineaMovimiento DESC")
            For Each row1 As DataRow In dttMov.Rows
                If row1("IdTipoMovimiento") <> enumTipoMovimiento.tmInventario Then
                    LineaArticuloAlmacenLote(0)("StockFisico") = Nz(LineaArticuloAlmacenLote(0)("StockFisico")) - row1("Cantidad")
                    If SegundaUnidad Then LineaArticuloAlmacenLote(0)("StockFisico2") = Nz(LineaArticuloAlmacenLote(0)("StockFisico2"), 0) - Nz(row1("Cantidad2"), 0)
                Else
                    LineaArticuloAlmacenLote(0)("StockFisico") = row1("Cantidad")
                    If SegundaUnidad Then LineaArticuloAlmacenLote(0)("StockFisico2") = row1("Cantidad2")
                End If
            Next
        Next

        For Each Dr As DataRow In dttStock.Select '("StockFisico <= " & data.Stock)
            Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, Dr("IDArticulo"), services)
            If Dr("StockFisico") <= data.Stock OrElse (SegundaUnidad AndAlso Not data.Stock2 Is Nothing AndAlso Nz(Dr("StockFisico2"), 0) <= data.Stock2) Then
                Dr.Delete()
            End If
        Next

        'For Each Dr As DataRow In dttStock.Select("StockFisico <= " & data.Stock)
        '    Dr.Delete()
        'Next
        dttStock.AcceptChanges()
        Return dttStock


    End Function

#End Region

#Region " Precio FIFO "

    <Serializable()> _
    Public Class DataValoracionAlmacenesFIFO
        Public ArticuloAlmacen As DataTable
        Public Orden As enumstkValoracionFIFO

        Public Sub New(ByVal ArticuloAlmacen As DataTable, Optional ByVal Orden As enumstkValoracionFIFO = enumstkValoracionFIFO.stkVFOrdenarPorFecha)
            Me.ArticuloAlmacen = ArticuloAlmacen
            Me.Orden = Orden
        End Sub
    End Class
    <Task()> Public Shared Function ValoracionAlmacenesFIFO(ByVal data As DataValoracionAlmacenesFIFO, ByVal services As ServiceProvider) As DataTable
        For Each dr As DataRow In data.ArticuloAlmacen.Select(Nothing, "IDArticulo,IDAlmacen")
            Dim datos As New DataValoracionFIFO(dr("IDArticulo"), dr("IDAlmacen"), dr("StockFisico"), dr("StockFisico"), Today, data.Orden)
            Dim valoracion As ValoracionPreciosInfo = ProcessServer.ExecuteTask(Of DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ValoracionFIFO, datos, services)
            If Not valoracion Is Nothing Then
                dr("PrecioFIFOA") = valoracion.PrecioA
                dr("PrecioFIFOB") = valoracion.PrecioB
                dr("ValorFIFOA") = valoracion.PrecioA * dr("StockFisico")
                dr("ValorFIFOB") = valoracion.PrecioB * dr("StockFisico")
            End If
        Next
        Return data.ArticuloAlmacen
    End Function

    <Serializable()> _
    Public Class DataValoracionFIFO
        Public IDArticulo As String
        Public IDAlmacen As String
        Public Stock As Double
        Public Cantidad As Double
        Public Fecha As Date
        Public Orden As enumstkValoracionFIFO

        Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal Stock As Double, ByVal Cantidad As Double, ByVal Fecha As Date)
            Me.IDArticulo = IDArticulo
            Me.IDAlmacen = IDAlmacen
            Me.Stock = Stock
            Me.Cantidad = Cantidad
            Me.Fecha = Fecha
            Me.Orden = enumstkValoracionFIFO.stkVFOrdenarPorFecha
        End Sub

        Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal Stock As Double, ByVal Cantidad As Double, ByVal Fecha As Date, ByVal Orden As enumstkValoracionFIFO)
            Me.IDArticulo = IDArticulo
            Me.IDAlmacen = IDAlmacen
            Me.Stock = Stock
            Me.Cantidad = Cantidad
            Me.Fecha = Fecha
            Me.Orden = Orden
        End Sub

        Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal Cantidad As Double, ByVal Fecha As Date, ByVal Orden As enumstkValoracionFIFO)
            Me.IDArticulo = IDArticulo
            Me.IDAlmacen = IDAlmacen
            Me.Cantidad = Cantidad
            Me.Fecha = Fecha
            Me.Orden = Orden
        End Sub

        Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal Cantidad As Double, ByVal Orden As enumstkValoracionFIFO)
            Me.IDArticulo = IDArticulo
            Me.IDAlmacen = IDAlmacen
            Me.Cantidad = Cantidad
            Me.Stock = Cantidad
            Me.Orden = Orden
        End Sub
    End Class
    <Task()> Public Shared Function ValoracionFIFO(ByVal data As DataValoracionFIFO, ByVal services As ServiceProvider) As ValoracionPreciosInfo
        Dim valoracion As New ValoracionPreciosInfo
        valoracion.IDArticulo = data.IDArticulo
        valoracion.IDAlmacen = data.IDAlmacen
        valoracion.FechaCalculo = data.Fecha
        Select Case data.Orden
            Case enumstkValoracionFIFO.stkVFOrdenarPorFecha
                valoracion.CriterioValoracion = enumtaValoracion.taPrecioFIFOFecha
            Case enumstkValoracionFIFO.stkVFOrdenarPorMvto
                valoracion.CriterioValoracion = enumtaValoracion.taPrecioFIFOMvto
        End Select

        If Len(data.IDArticulo) > 0 AndAlso Len(data.IDAlmacen) > 0 AndAlso data.Cantidad >= 0 AndAlso data.Fecha <> cnMinDate Then
            Dim movimientos As DataTable = ProcessServer.ExecuteTask(Of DataValoracionFIFO, DataTable)(AddressOf GetMovimientosFIFO, data, services)
            If Not movimientos Is Nothing AndAlso movimientos.Rows.Count > 0 Then
                Dim dCalcPrecio As New DataCalculoPrecioFIFO(movimientos, data.Stock, data.Cantidad)
                Dim precios As ValoracionPreciosInfo = ProcessServer.ExecuteTask(Of DataCalculoPrecioFIFO, ValoracionPreciosInfo)(AddressOf CalculoPrecioFIFO, dCalcPrecio, services)
                valoracion.PrecioA = precios.PrecioA
                valoracion.PrecioB = precios.PrecioB
            End If
        End If

        Return valoracion
    End Function

    <Task()> Public Shared Function GetMovimientosFIFO(ByVal data As DataValoracionFIFO, ByVal services As ServiceProvider) As DataTable
        Dim movimientos As DataTable
        movimientos = AdminData.Execute("sp_MovimientosFIFO", False, data.IDArticulo, data.IDAlmacen, data.Stock, data.Fecha, data.Orden)
        'Select Case data.Orden
        '    Case enumstkValoracionFIFO.stkVFOrdenarPorFecha
        '        'valoracion.CriterioValoracion = enumtaValoracion.taPrecioFIFOFecha
        '    Case enumstkValoracionFIFO.stkVFOrdenarPorMvto
        '        'valoracion.CriterioValoracion = enumtaValoracion.taPrecioFIFOMvto
        'End Select


        Return movimientos
    End Function


    <Serializable()> _
    Public Class DataCalculoPrecioFIFO
        Public Movimientos As DataTable
        Public Stock As Double
        Public Cantidad As Double

        Public Sub New(ByVal Movimientos As DataTable, ByVal Stock As Double, ByVal Cantidad As Double)
            Me.Movimientos = Movimientos
            Me.Stock = Stock
            Me.Cantidad = Cantidad
        End Sub
    End Class
    <Task()> Public Shared Function CalculoPrecioFIFO(ByVal data As DataCalculoPrecioFIFO, ByVal services As ServiceProvider) As ValoracionPreciosInfo
        Dim precios As New ValoracionPreciosInfo
        If Not data.Movimientos Is Nothing AndAlso data.Movimientos.Rows.Count > 0 Then
            Dim QRestante As Double = data.Cantidad
            Dim QAcumulado As Double
            Dim ValorFifoA, ValorFifoB As Double
            Dim arrayMovimientos As List(Of DataRow) = (From c In data.Movimientos Select c).ToList
            If data.Cantidad = data.Stock Then
                arrayMovimientos.Reverse()
            End If
            For Each movimiento As DataRow In arrayMovimientos
                QAcumulado += movimiento("Cantidad")

                If QAcumulado >= data.Cantidad Then
                    ValorFifoA += (QRestante * movimiento("PrecioA"))
                    ValorFifoB += (QRestante * movimiento("PrecioB"))
                    Exit For
                Else
                    QRestante -= movimiento("Cantidad")
                    ValorFifoA += (movimiento("Cantidad") * movimiento("PrecioA"))
                    ValorFifoB += (movimiento("Cantidad") * movimiento("PrecioB"))
                End If
            Next

            If data.Cantidad > QAcumulado Then
                '//no hay movimientos suficientes. Se calcula el fifo para la cantidad que hay.
                data.Cantidad = QAcumulado
            End If

            Dim PrecioA, PrecioB As Double
            If data.Cantidad > 0 Then
                PrecioA = ValorFifoA / data.Cantidad
                PrecioB = ValorFifoB / data.Cantidad
            End If

            precios.PrecioA = PrecioA
            precios.PrecioB = PrecioB
        End If

        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.MonedaA
        Dim MonInfoB As MonedaInfo = Monedas.MonedaB

        precios.PrecioA = xRound(precios.PrecioA, MonInfoA.NDecimalesPrecio)
        precios.PrecioB = xRound(precios.PrecioB, MonInfoB.NDecimalesPrecio)

        Return precios
    End Function

#End Region

#Region " Precio medio "

    <Task()> Public Shared Function ValoracionPrecioMedioAFecha(ByVal data As DataArticuloAlmacenFecha, ByVal services As ServiceProvider) As ValoracionPreciosInfo
        Dim ValIniPM As ValoresInicialesPrecioMedio = ProcessServer.ExecuteTask(Of DataArticuloAlmacenFecha, ValoresInicialesPrecioMedio)(AddressOf PrecioMedioValoresIniciales, data, services)
        Dim datosCalPM As New DataCalculoPrecioMedio(data, ValIniPM, enumtpmTipoMovimiento.tpmInput)
        Return ProcessServer.ExecuteTask(Of DataCalculoPrecioMedio, ValoracionPreciosInfo)(AddressOf CalculoPrecioMedio, datosCalPM, services)
    End Function

    <Serializable()> _
    Public Class DataValoracionPrecioMedio
        Public IDArticulo As String
        Public IDAlmacen As String
        Public FechaCalculo As Date
        Public FechaInicioCalculo As Date?

        Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal FechaCalculo As Date, Optional ByVal FechaInicioCalculo As Date = cnMinDate)
            Me.IDArticulo = IDArticulo
            Me.IDAlmacen = IDAlmacen
            Me.FechaCalculo = FechaCalculo
            If FechaInicioCalculo <> cnMinDate Then Me.FechaInicioCalculo = FechaInicioCalculo
        End Sub
    End Class
    <Task()> Public Shared Function ValoracionPrecioMedio(ByVal data As DataValoracionPrecioMedio, ByVal services As ServiceProvider) As ValoracionPreciosInfo
        If Not data.FechaInicioCalculo Is Nothing AndAlso data.FechaInicioCalculo <> cnMinDate AndAlso data.FechaInicioCalculo <= data.FechaCalculo Then
            '//Si se ha introducido una fecha de inicio y tenemos un periodo de cálculo.
            Dim obtMovto As New DataObtenerPrimerMovto(data.IDArticulo, data.IDAlmacen, data.FechaInicioCalculo, data.FechaCalculo)
            Dim ValIniPM As New ValoresInicialesPrecioMedio
            ValIniPM = ProcessServer.ExecuteTask(Of DataObtenerPrimerMovto, ValoresInicialesPrecioMedio)(AddressOf ObtenerPrimerMovimientoInventario, obtMovto, services)
            If ValIniPM.IsEmpty Then '//Si no hay inventarios en el período de cálculo.
                ValIniPM = ProcessServer.ExecuteTask(Of DataObtenerPrimerMovto, ValoresInicialesPrecioMedio)(AddressOf ObtenerPrimerMovimientoEntrada, obtMovto, services)
            End If
            If ValIniPM.IsEmpty Then
                '//Si no hay entradas en el período de cálculo, calculamos el Precio Medio a Fecha de Inicio
                Dim ArtAlmFechaInicio As New DataArticuloAlmacenFecha(data.IDArticulo, data.IDAlmacen, data.FechaInicioCalculo)
                Dim preciosIniciales As ValoracionPreciosInfo = ProcessServer.ExecuteTask(Of DataArticuloAlmacenFecha, ValoracionPreciosInfo)(AddressOf ValoracionPrecioMedioAFecha, ArtAlmFechaInicio, services)
                Dim stock As StockAFechaInfo = ProcessServer.ExecuteTask(Of DataArticuloAlmacenFecha, StockAFechaInfo)(AddressOf GetStockAcumuladoAFecha, ArtAlmFechaInicio, services)
                ValIniPM.Stock = stock.StockAFecha

                ValIniPM.Fecha = data.FechaInicioCalculo
                ValIniPM.PrecioA = preciosIniciales.PrecioA
                ValIniPM.PrecioB = preciosIniciales.PrecioB
            End If
            '// Realizamos el cálculo en el período indicado. Desde el punto de partida encontrado.
            Dim datosCalPM As New DataCalculoPrecioMedio(data.IDArticulo, data.IDAlmacen, data.FechaCalculo, ValIniPM, enumtpmTipoMovimiento.tpmInput)
            Return ProcessServer.ExecuteTask(Of DataCalculoPrecioMedio, ValoracionPreciosInfo)(AddressOf CalculoPrecioMedio, datosCalPM, services)
        ElseIf data.FechaInicioCalculo Is Nothing OrElse data.FechaInicioCalculo = cnMinDate Then
            '//Si no se ha introducido una Fecha de Inicio, calculamos el Precio Medio a Fecha de Cálculo
            Dim ArtAlmFechaFin As New DataArticuloAlmacenFecha(data.IDArticulo, data.IDAlmacen, data.FechaCalculo)
            Return ProcessServer.ExecuteTask(Of DataArticuloAlmacenFecha, ValoracionPreciosInfo)(AddressOf ValoracionPrecioMedioAFecha, ArtAlmFechaFin, services)
        End If
    End Function

    <Task()> Public Shared Function PrecioMedioInventario(ByVal data As DataArticuloAlmacenFecha, ByVal services As ServiceProvider) As ValoracionPreciosInfo
        Dim ValIniPM As ValoresInicialesPrecioMedio = ProcessServer.ExecuteTask(Of DataArticuloAlmacenFecha, ValoresInicialesPrecioMedio)(AddressOf PrecioMedioValoresIniciales, data, services)
        Dim datosCalPM As New DataCalculoPrecioMedio(data.IDArticulo, data.IDAlmacen, data.Fecha, ValIniPM, enumtpmTipoMovimiento.tpmInventario)
        Return ProcessServer.ExecuteTask(Of DataCalculoPrecioMedio, ValoracionPreciosInfo)(AddressOf CalculoPrecioMedio, datosCalPM, services)
    End Function

    <Serializable()> _
    Public Class DataCalculoPrecioMedio
        Inherits DataArticuloAlmacenFecha

        Public ValIniPM As New ValoresInicialesPrecioMedio
        Public ClaseMovimiento As enumtpmTipoMovimiento

        Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal FechaCalculo As Date, ByVal ValIniPM As ValoresInicialesPrecioMedio, ByVal ClaseMovimiento As enumtpmTipoMovimiento)
            MyBase.New(IDArticulo, IDAlmacen, FechaCalculo)
            Me.ValIniPM = ValIniPM
            Me.ClaseMovimiento = ClaseMovimiento
        End Sub

        Public Sub New(ByVal ArtAlmFec As DataArticuloAlmacenFecha, ByVal ValIniPM As ValoresInicialesPrecioMedio, ByVal ClaseMovimiento As enumtpmTipoMovimiento)
            MyBase.New(ArtAlmFec.IDArticulo, ArtAlmFec.IDAlmacen, ArtAlmFec.Fecha)
            Me.ValIniPM = ValIniPM
            Me.ClaseMovimiento = ClaseMovimiento
        End Sub
    End Class

    <Task()> Public Shared Function CalculoPrecioMedio(ByVal data As DataCalculoPrecioMedio, ByVal services As ServiceProvider) As ValoracionPreciosInfo
        Dim Stock As Double = data.ValIniPM.Stock
        Dim PrecioAIniPM As Double = data.ValIniPM.PrecioA
        Dim PrecioBIniPM As Double = data.ValIniPM.PrecioB
        Dim PrecioMedioA As Double = PrecioAIniPM
        Dim PrecioMedioB As Double = PrecioBIniPM
        Dim FechaInicio As Date = data.ValIniPM.Fecha

        Dim movimientos As SqlClient.SqlDataReader
        If Len(data.IDArticulo) > 0 And Len(data.IDAlmacen) > 0 And data.Fecha >= FechaInicio Then
            Dim obtMovto As New DataObtenerPrimerMovto(data.IDArticulo, data.IDAlmacen, FechaInicio, data.Fecha)
            Dim vi As ValoresInicialesPrecioMedio = ProcessServer.ExecuteTask(Of DataObtenerPrimerMovto, ValoresInicialesPrecioMedio)(AddressOf ObtenerPrimerMovimientoInventario, obtMovto, services)
            If Not vi.IsEmpty Then
                Dim AppParamsStocks As ParametroStocks = services.GetService(Of ParametroStocks)()
                '//parametro TipoInventario
                Select Case AppParamsStocks.TipoInventario
                    Case TipoInventario.PrimerMovimiento
                        If data.ClaseMovimiento = enumtpmTipoMovimiento.tpmInventario Then
                            '//si se esta valorando un movimiento de inventario, dicho movimiento se va a situar
                            '//como primer movimiento del dia a la Fecha de Calculo, entonces no hay que incluir los
                            '//movimientos de entrada que pudiese haber esa fecha
                            data.Fecha = data.Fecha.AddDays(-1)
                        End If
                    Case TipoInventario.UltimoMovimiento
                        If Not (vi.Fecha = Date.MinValue) Then
                            vi.Fecha = CDate(vi.Fecha).AddDays(1)
                        End If
                End Select

                Stock = vi.Stock
                PrecioAIniPM = vi.PrecioA
                PrecioBIniPM = vi.PrecioB
                PrecioMedioA = PrecioAIniPM
                PrecioMedioB = PrecioBIniPM
                FechaInicio = vi.Fecha
            End If

            If FechaInicio <= data.Fecha Then
                Dim f As New Filter
                f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
                f.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
                f.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThanOrEqual, FechaInicio))
                f.Add(New DateFilterItem("FechaDocumento", FilterOperator.LessThanOrEqual, data.Fecha))
                f.Add(New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion))
                f.Add(New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmInventario))

                Dim strSELECT As String = "IDLineaMovimiento, IDArticulo, IDAlmacen, IDTipoMovimiento, Cantidad, PrecioA, PrecioB, FechaDocumento, ClaseMovimiento, Fifo, Acumulado"
                Dim cnn As Common.DbConnection = AdminData.GetSessionConnection.Connection
                Using cmd As Common.DbCommand = cnn.CreateCommand()
                    cmd.CommandText = "SELECT * FROM vNegMovimientosPrecioMedio WHERE " + AdminData.ComposeFilter(f) + " ORDER BY FechaDocumento,IDLineaMovimiento"
                    cmd.Transaction = AdminData.GetSessionConnection.Transaction
                    movimientos = cmd.ExecuteReader()
                End Using
            End If

        End If

        Dim valoracion As New ValoracionPreciosInfo
        valoracion.IDArticulo = data.IDArticulo
        valoracion.IDAlmacen = data.IDAlmacen
        valoracion.CriterioValoracion = enumtaValoracion.taPrecioMedio
        valoracion.FechaCalculo = data.Fecha
        valoracion.PrecioA = PrecioMedioA
        valoracion.PrecioB = PrecioMedioB

        If (Not movimientos Is Nothing) Then
            Try
                Dim StockAnt As Double = Stock
                Do While movimientos.Read
                    Stock = xRound((Stock + movimientos("Cantidad")), 6)
                    If (Stock > 0) And (movimientos("ClaseMovimiento") = enumtpmTipoMovimiento.tpmInput) Then
                        If StockAnt > 0 Then
                            PrecioMedioA = ((PrecioAIniPM * StockAnt) + (movimientos("Cantidad") * movimientos("PrecioA"))) / Stock
                            PrecioMedioB = ((PrecioBIniPM * StockAnt) + (movimientos("Cantidad") * movimientos("PrecioB"))) / Stock
                        Else
                            PrecioMedioA = movimientos("PrecioA")
                            PrecioMedioB = movimientos("PrecioB")
                        End If
                        PrecioAIniPM = PrecioMedioA
                        PrecioBIniPM = PrecioMedioB
                    End If
                    StockAnt = Stock
                Loop
                valoracion.PrecioA = PrecioMedioA
                valoracion.PrecioB = PrecioMedioB
            Finally
                movimientos.Close()
            End Try
        End If

        Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
        Dim MonInfoA As MonedaInfo = Monedas.MonedaA
        Dim MonInfoB As MonedaInfo = Monedas.MonedaB

        valoracion.PrecioA = xRound(valoracion.PrecioA, MonInfoA.NDecimalesPrecio)
        valoracion.PrecioB = xRound(valoracion.PrecioB, MonInfoB.NDecimalesPrecio)

        Return valoracion
    End Function

    <Serializable()> _
    Public Class ValoresInicialesPrecioMedio
        Public Stock As Double
        Public Fecha As Date
        Public PrecioA As Double
        Public PrecioB As Double
        Public ReadOnly Property IsEmpty() As Boolean
            Get
                Return ((Me.Fecha = Date.MinValue) AndAlso (Me.PrecioA = 0) AndAlso (Me.PrecioB = 0) AndAlso (Me.Stock = 0))
                'Return ((Me.Fecha Is Nothing) AndAlso (Me.PrecioA Is Nothing) AndAlso (Me.PrecioB Is Nothing) AndAlso (Me.Stock Is Nothing))
            End Get
        End Property
    End Class

    <Serializable()> _
    Public Class DatosUltimoCierre
        Public FechaCalculo As Date
        Public IDEjercicio As String
        Public IDMesCierre As Integer
        Public FechaUltimoCierre As Date

        Public Sub New()
            Me.FechaCalculo = cnMinDate
        End Sub
        Public Sub New(ByVal FechaCalculo As Date, ByVal IDEjercicio As String, ByVal IDMesCierre As Integer, ByVal FechaUltimoCierre As Date)
            Me.FechaCalculo = FechaCalculo
            Me.IDEjercicio = IDEjercicio
            Me.IDMesCierre = IDMesCierre
            Me.FechaUltimoCierre = FechaUltimoCierre
        End Sub
    End Class

    <Task()> Public Shared Function PrecioMedioValoresIniciales(ByVal data As DataArticuloAlmacenFecha, ByVal services As ServiceProvider) As ValoresInicialesPrecioMedio
        Dim ValIniPM As New ValoresInicialesPrecioMedio
        If Len(data.IDArticulo) > 0 And Len(data.IDAlmacen) > 0 Then
            ValIniPM = ProcessServer.ExecuteTask(Of DataArticuloAlmacenFecha, ValoresInicialesPrecioMedio)(AddressOf GetValoresInicialesArticuloAlmacen, data, services)
            If ValIniPM.IsEmpty Then  '//No tenemos valores Iniciales en ArticuloAlmacen
                ValIniPM = ProcessServer.ExecuteTask(Of DataArticuloAlmacenFecha, ValoresInicialesPrecioMedio)(AddressOf GetValoresInicialesUltimoCierreInv, data, services)
            End If
        End If

        Return ValIniPM
    End Function
    <Task()> Public Shared Function GetValoresInicialesArticuloAlmacen(ByVal data As DataArticuloAlmacenFecha, ByVal services As ServiceProvider) As ValoresInicialesPrecioMedio
        Dim ValIniPM As New ValoresInicialesPrecioMedio
        Dim f As New Filter
        f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
        f.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
        Dim valores As DataTable = New BE.DataEngine().Filter("tbMaestroArticuloAlmacen", f, "TOP 1 PrecioMedioA,PrecioMedioB,FechaCalculo,StockFechaCalculo")
        If Not valores Is Nothing AndAlso valores.Rows.Count > 0 Then
            If IsDate(valores.Rows(0)("FechaCalculo")) Then
                If valores.Rows(0)("FechaCalculo") <= data.Fecha Then  '//valores1.Rows(0)("FechaCalculo") <= FechaCalculo
                    ValIniPM.Fecha = DateAdd(DateInterval.Day, 1, valores.Rows(0)("FechaCalculo"))
                    ValIniPM.PrecioA = valores.Rows(0)("PrecioMedioA")
                    ValIniPM.PrecioB = valores.Rows(0)("PrecioMedioB")
                    ValIniPM.Stock = valores.Rows(0)("StockFechaCalculo")
                End If
            End If
        End If
        Return ValIniPM
    End Function
    <Task()> Public Shared Function GetValoresInicialesUltimoCierreInv(ByVal data As DataArticuloAlmacenFecha, ByVal services As ServiceProvider) As ValoresInicialesPrecioMedio
        Dim ValIniPM As New ValoresInicialesPrecioMedio
        Dim ValUltCierre As DatosUltimoCierre = services.GetService(Of DatosUltimoCierre)()
        If ValUltCierre Is Nothing OrElse ValUltCierre.FechaCalculo = cnMinDate Then
            Dim valores As DataTable = New BE.DataEngine().Filter("vNegCierreInventarioUltimo", New DateFilterItem("FechaHasta", FilterOperator.LessThanOrEqual, data.Fecha), "TOP 1 IDEjercicio,IDMesCierre,FechaHasta", "FechaCierre DESC, FechaHasta DESC")
            If Not valores Is Nothing AndAlso valores.Rows.Count > 0 Then
                'Dim fechaUltimoCierre As Date
                If IsDate(valores.Rows(0)("FechaHasta")) Then
                    'fechaUltimoCierre = valores2.Rows(0)("FechaHasta")
                    ValUltCierre = New DatosUltimoCierre(data.Fecha, valores.Rows(0)("IDEjercicio"), valores.Rows(0)("IDMesCierre"), valores.Rows(0)("FechaHasta"))
                    services.RegisterService(ValUltCierre)
                End If
            End If
        End If
        If ValUltCierre.FechaUltimoCierre <= data.Fecha AndAlso ValUltCierre.FechaUltimoCierre > cnMinDate Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDEjercicio", ValUltCierre.IDEjercicio))
            f.Add(New NumberFilterItem("IDMesCierre", ValUltCierre.IDMesCierre))
            f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
            f.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
            Dim detalle As DataTable = New BE.DataEngine().Filter("tbCierreInventarioDetalle", f, "TOP 1 IDArticulo,IDAlmacen,PrecioMedioA,PrecioMedioB,StockFisico")
            If Not detalle Is Nothing AndAlso detalle.Rows.Count > 0 Then
                ValIniPM.Fecha = ValUltCierre.FechaUltimoCierre
                ValIniPM.PrecioA = detalle.Rows(0)("PrecioMedioA")
                ValIniPM.PrecioB = detalle.Rows(0)("PrecioMedioB")
                ValIniPM.Stock = detalle.Rows(0)("StockFisico")
            End If
        End If
        Return ValIniPM
    End Function

#End Region

#Region " Valoración en fecha "

    <Serializable()> _
    Public Class DataValoracionEnFecha
        Public FechaCalculo As Date         '// Fecha en la que se realiza el cálculo
        Public UDValoracion As String       '// Unidad de Medida para valoración extra en otra unidad
        Public SalvarPrecios As Boolean     '// Indica si se guardan los precios calculados en ArticuloAlmacen (Excepto Precio STD y Precio Ult. Compra)

        Public ArticuloAlmacen As DataTable '// registros a intervenir en el proceso

        Public Sub New(ByVal FechaCalculo As Date, ByVal ArticuloAlmacen As DataTable, Optional ByVal SalvarPrecios As Boolean = False)
            Me.FechaCalculo = FechaCalculo.Date  '//.Date para no considerar las horas de la fecha (en algunos casos llegan las fechas pueden legar con horas)
            Me.ArticuloAlmacen = ArticuloAlmacen
            Me.SalvarPrecios = SalvarPrecios
        End Sub
    End Class

    <Serializable()> _
    Public Class DataValoracionEnFechaPrecioMedio
        Inherits DataValoracionEnFecha

        '//Esta fecha podrá ser:
        '// 1.- Una Fecha indicada por el usuario
        '// 2.- La Fecha del último Inventario
        '// 3.- La Fecha del último Cierre
        '// 4.- La Fecha del último Cálculo guardado en tbArticuloAlmacen
        Public FechaInicioCalculo As Date         '// Fecha desde la que se empiezan a tener en cuenta los movimientos. 

        Public Sub New(ByVal FechaCalculo As Date, ByVal ArticuloAlmacen As DataTable, Optional ByVal SalvarPrecios As Boolean = False, Optional ByVal FechaInicioCalculo As Date = cnMinDate)
            MyBase.New(FechaCalculo, ArticuloAlmacen, SalvarPrecios)
            If FechaInicioCalculo <> cnMinDate Then Me.FechaInicioCalculo = FechaInicioCalculo.Date '//.Date para no considerar las horas de la fecha (en algunos casos llegan las fechas pueden legar con horas)
        End Sub
    End Class

    <Serializable()> _
    Public Class DataValoracionEnFechaFIFO
        Inherits DataValoracionEnFecha

        Public Orden As enumstkValoracionFIFO

        Public Sub New(ByVal FechaCalculo As Date, ByVal ArticuloAlmacen As DataTable, ByVal Orden As enumstkValoracionFIFO, Optional ByVal SalvarPrecios As Boolean = False)
            MyBase.New(FechaCalculo, ArticuloAlmacen, SalvarPrecios)
            Me.Orden = Orden
        End Sub
    End Class

    '<Task()> Public Shared Function ObtenerDatosProcesoValoracion(ByVal f As Filter, ByVal services As ServiceProvider) As DataTable
    '    '//Recuperamos los pares Articulo-Almacen a entrar en el proceso de cálculo
    '    Return New BE.DataEngine().Filter("vFrmCIValoracionAlmacenEnFecha", f)
    'End Function

    <Task()> Public Shared Function ValoracionEnFechaPrecioEstandar(ByVal data As DataValoracionEnFecha, ByVal services As ServiceProvider) As ValoracionResultInfo
        Dim dtArticuloAlmacen As DataTable = data.ArticuloAlmacen 'ProcessServer.ExecuteTask(Of Filter, DataTable)(AddressOf ObtenerDatosProcesoValoracion, data.fArticuloAlmacen, services)

        Dim result As ValoracionResultInfo
        If data.FechaCalculo > cnMinDate And data.FechaCalculo <= Today Then
            Dim ValStock As New DataValoracionStockAFecha(data.FechaCalculo, dtArticuloAlmacen, , enumtaValoracion.taPrecioEstandar)
            ValStock.UDValoracion = data.UDValoracion
            Dim Valoraciones() As ValoracionInfo = ProcessServer.ExecuteTask(Of DataValoracionStockAFecha, ValoracionInfo())(AddressOf GetValoracionStockAFecha, ValStock, services)
            result = New ValoracionResultInfo
            result.Valoraciones = Valoraciones
            If Not IsNothing(dtArticuloAlmacen) Then
                result.ArticuloAlmacen = dtArticuloAlmacen
                result.ArticuloAlmacen.RemotingFormat = SerializationFormat.Binary
            End If
        Else
            Throw New Exception(ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(25), services))
        End If

        Return result
    End Function

    <Task()> Public Shared Function ValoracionEnFechaPrecioUltimaCompra(ByVal data As DataValoracionEnFecha, ByVal services As ServiceProvider) As ValoracionResultInfo
        Dim dtArticuloAlmacen As DataTable = data.ArticuloAlmacen ' ProcessServer.ExecuteTask(Of Filter, DataTable)(AddressOf ObtenerDatosProcesoValoracion, data.fArticuloAlmacen, services)

        Dim result As ValoracionResultInfo
        If data.FechaCalculo > cnMinDate And data.FechaCalculo <= Today Then
            Dim nArticulo As New Negocio.Articulo
            Dim articulos As DataTable = nArticulo.Filter()
            Dim ValStock As New DataValoracionStockAFecha(data.FechaCalculo, dtArticuloAlmacen, , enumtaValoracion.taPrecioUltCompra)
            ValStock.UDValoracion = data.UDValoracion
            Dim Valoraciones() As ValoracionInfo = ProcessServer.ExecuteTask(Of DataValoracionStockAFecha, ValoracionInfo())(AddressOf GetValoracionStockAFecha, ValStock, services)
            result = New ValoracionResultInfo
            result.Valoraciones = Valoraciones
            If Not IsNothing(dtArticuloAlmacen) Then
                result.ArticuloAlmacen = dtArticuloAlmacen
                result.ArticuloAlmacen.RemotingFormat = SerializationFormat.Binary
            End If
        Else
            Throw New Exception(ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(25), services))
        End If

        Return result
    End Function

    <Task()> Public Shared Function ValoracionEnFechaPrecioMedio(ByVal data As DataValoracionEnFechaPrecioMedio, ByVal services As ServiceProvider) As ValoracionResultInfo
        Dim dtArticuloAlmacen As DataTable = data.ArticuloAlmacen 'ProcessServer.ExecuteTask(Of Filter, DataTable)(AddressOf ObtenerDatosProcesoValoracion, data.fArticuloAlmacen, services)
        Dim result As ValoracionResultInfo
        If data.FechaCalculo > cnMinDate And data.FechaCalculo <= Today Then
            Dim ValStock As New DataValoracionStockAFecha(data.FechaCalculo, dtArticuloAlmacen, data.FechaInicioCalculo, enumtaValoracion.taPrecioMedio)
            ValStock.UDValoracion = data.UDValoracion
            Dim Valoraciones() As ValoracionInfo = ProcessServer.ExecuteTask(Of DataValoracionStockAFecha, ValoracionInfo())(AddressOf GetValoracionStockAFecha, ValStock, services)
            result = New ValoracionResultInfo
            result.Valoraciones = Valoraciones
            If Not IsNothing(dtArticuloAlmacen) Then
                result.ArticuloAlmacen = dtArticuloAlmacen
                result.ArticuloAlmacen.RemotingFormat = SerializationFormat.Binary
            End If

            If data.SalvarPrecios Then
                Dim datosSalvar As New DataSalvarPrecios(data.FechaCalculo, Valoraciones)
                ProcessServer.ExecuteTask(Of DataSalvarPrecios)(AddressOf SalvarPrecios, datosSalvar, services)
            End If
        Else
            Throw New Exception(ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(25), services))
        End If

        Return result
    End Function

    <Task()> Public Shared Function ValoracionEnFechaPrecioFIFO(ByVal data As DataValoracionEnFechaFIFO, ByVal services As ServiceProvider) As ValoracionResultInfo
        Dim dtArticuloAlmacen As DataTable = data.ArticuloAlmacen 'ProcessServer.ExecuteTask(Of Filter, DataTable)(AddressOf ObtenerDatosProcesoValoracion, data.fArticuloAlmacen, services)

        Dim result As ValoracionResultInfo
        If data.FechaCalculo > cnMinDate And data.FechaCalculo <= Today Then
            Dim Valoraciones() As ValoracionInfo
            Select Case data.Orden
                Case enumstkValoracionFIFO.stkVFOrdenarPorFecha
                    Dim ValStock As New DataValoracionStockAFecha(data.FechaCalculo, dtArticuloAlmacen, , enumtaValoracion.taPrecioFIFOFecha)
                    ValStock.UDValoracion = data.UDValoracion
                    Valoraciones = ProcessServer.ExecuteTask(Of DataValoracionStockAFecha, ValoracionInfo())(AddressOf GetValoracionStockAFecha, ValStock, services)
                Case enumstkValoracionFIFO.stkVFOrdenarPorMvto
                    Dim ValStock As New DataValoracionStockAFecha(data.FechaCalculo, dtArticuloAlmacen, , enumtaValoracion.taPrecioFIFOMvto)
                    ValStock.UDValoracion = data.UDValoracion
                    Valoraciones = ProcessServer.ExecuteTask(Of DataValoracionStockAFecha, ValoracionInfo())(AddressOf GetValoracionStockAFecha, ValStock, services)
            End Select
            result = New ValoracionResultInfo
            result.Valoraciones = Valoraciones
            If Not IsNothing(dtArticuloAlmacen) Then
                result.ArticuloAlmacen = dtArticuloAlmacen
                result.ArticuloAlmacen.RemotingFormat = SerializationFormat.Binary
            End If
            If data.SalvarPrecios Then
                Dim datosSalvar As New DataSalvarPrecios(data.FechaCalculo, Valoraciones)
                ProcessServer.ExecuteTask(Of DataSalvarPrecios)(AddressOf SalvarPrecios, datosSalvar, services)
            End If
        Else
            Throw New Exception(ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(25), services))
        End If

        Return result
    End Function

    <Task()> Public Shared Function ValoracionEnFecha(ByVal data As DataValoracionEnFechaPrecioMedio, ByVal services As ServiceProvider) As ValoracionResultInfo
        Dim dtArticuloAlmacen As DataTable = data.ArticuloAlmacen 'ProcessServer.ExecuteTask(Of Filter, DataTable)(AddressOf ObtenerDatosProcesoValoracion, data.fArticuloAlmacen, services)

        Dim result As New ValoracionResultInfo

        '//Aplicar el criterio de valoracion del articulo
        Dim Valoraciones(-1) As ValoracionInfo
        If data.FechaCalculo > cnMinDate And data.FechaCalculo <= Today Then
            Dim ValStock As New DataValoracionStockAFecha(data.FechaCalculo, dtArticuloAlmacen, data.FechaInicioCalculo)
            ValStock.UDValoracion = data.UDValoracion
            Valoraciones = ProcessServer.ExecuteTask(Of DataValoracionStockAFecha, ValoracionInfo())(AddressOf GetValoracionStockAFecha, ValStock, services)
            If Not IsNothing(Valoraciones) Then
                If data.SalvarPrecios Then
                    Dim datosSalvar As New DataSalvarPrecios(data.FechaCalculo, Valoraciones)
                    ProcessServer.ExecuteTask(Of DataSalvarPrecios)(AddressOf SalvarPrecios, datosSalvar, services)
                End If
            End If
        Else
            Throw New Exception(ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(25), services))
        End If
        result.Valoraciones = Valoraciones
        If Not IsNothing(dtArticuloAlmacen) Then
            result.ArticuloAlmacen = dtArticuloAlmacen
            result.ArticuloAlmacen.RemotingFormat = SerializationFormat.Binary
        End If
        Return result
    End Function


    <Serializable()> _
    Public Class DataCalculoValoracionEnFechaAlmacen
        Public IDUDValoracion As String
        Public FechaCalculo As Date
        Public Filtro As Filter

        Public Sub New(ByVal FechaCalculo As Date, Optional ByVal IDUDValoracion As String = Nothing, Optional ByVal Filtro As Filter = Nothing)
            Me.FechaCalculo = FechaCalculo
            Me.IDUDValoracion = IDUDValoracion
            Me.Filtro = Filtro
        End Sub
    End Class

    <Task()> Public Shared Function CalculoValoracionEnFechaAlmacen(ByVal data As DataCalculoValoracionEnFechaAlmacen, ByVal services As ServiceProvider) As String
        Dim Esquema As String = ProcessServer.ExecuteTask(Of Object, String)(AddressOf Comunes.GetEsquemaBD, Nothing, services)

        Dim vStr As String = "SELECT v.*"
        If Length(data.IDUDValoracion) > 0 Then
            vStr &= ", " & Esquema & ".fFactorArticuloAB(v.IDArticulo, v.IDUdInterna, " & Quoted(data.IDUDValoracion & String.Empty) & ", 0) AS FactorUdValoracion, " & _
                             "v.Acumulado * " & Esquema & ".fFactorArticuloAB(v.IDArticulo, v.IDUdInterna, " & Quoted(data.IDUDValoracion & String.Empty) & ", 0) AS StockAFechaUdValoracion"
        End If
        vStr &= " FROM vFrmCIValoracionAlmacenFecha v RIGHT JOIN " & _
             "(SELECT   tbMaestroArticuloAlmacen.IDArticulo AS MIDArticulo, tbMaestroArticuloAlmacen.IDAlmacen AS MIDAlmacen," & Esquema & ".fMovimientoValArticulo('" & Format(CDate(data.FechaCalculo), "yyyyMMdd") & "', IDArticulo, IDAlmacen) AS MIDLineaMovimiento " & _
             "  FROM            tbMaestroArticuloAlmacen WHERE " & Esquema & ".fMovimientoValArticulo('" & Format(CDate(data.FechaCalculo), "yyyyMMdd") & "', IDArticulo, IDAlmacen)<>0 ) AS MovimientoArticuloAlmacen ON " & _
             " v.IDLineaMovimiento = MovimientoArticuloAlmacen.MIDLineaMovimiento " & _
             " AND v.IDArticulo = MovimientoArticuloAlmacen.MIDArticulo " & _
             " AND v.IDAlmacen = MovimientoArticuloAlmacen.MIDAlmacen " & _
             " WHERE v.IDLineaMovimiento IS NOT NULL"
        If Not data.Filtro Is Nothing AndAlso data.Filtro.Count > 0 Then
            vStr &= " AND " & data.Filtro.Compose(New AdoFilterComposer)
        End If

        Return vStr
    End Function

    <Serializable()> _
    Public Class DataSalvarPrecios
        Public FechaCalculo As Date
        Public Valoraciones() As ValoracionInfo
        'Friend Filtro As Filter

        Public Sub New(ByVal FechaCalculo As Date, ByVal Valoraciones() As ValoracionInfo) ', ByVal Filtro As Filter)
            Me.FechaCalculo = FechaCalculo
            Me.Valoraciones = Valoraciones
            'Me.Filtro = Filtro
        End Sub
    End Class
    <Task()> Public Shared Sub SalvarPrecios(ByVal data As DataSalvarPrecios, ByVal services As ServiceProvider)
        Dim f As New Filter : Dim aa As New ArticuloAlmacen
        For Each valoracion As ValoracionInfo In data.Valoraciones
            Dim FilCierre As New Filter
            FilCierre.Add("IDArticulo", FilterOperator.Equal, valoracion.IDArticulo)
            FilCierre.Add("IDAlmacen", FilterOperator.Equal, valoracion.IDAlmacen)
            FilCierre.Add("FechaCalculo", FilterOperator.GreaterThanOrEqual, data.FechaCalculo)
            Dim DtCierre As DataTable = New CierreInventarioDetalle().Filter(FilCierre)
            If DtCierre Is Nothing OrElse DtCierre.Rows.Count = 0 Then
                AdminData.BeginTx()
                f.Clear()
                f.Add(New StringFilterItem("IDArticulo", valoracion.IDArticulo))
                f.Add(New StringFilterItem("IDAlmacen", valoracion.IDAlmacen))
                Dim dt As DataTable = aa.Filter(f)
                If dt.Rows.Count > 0 Then
                    dt.Rows(0)("FechaCalculo") = data.FechaCalculo
                    dt.Rows(0)("StockFechaCalculo") = valoracion.Stock.StockAFecha
                    dt.Rows(0)("PrecioMedioA") = 0
                    dt.Rows(0)("PrecioMedioB") = 0
                    dt.Rows(0)("PrecioFIFOFechaA") = 0
                    dt.Rows(0)("PrecioFIFOFechaB") = 0
                    dt.Rows(0)("PrecioFIFOMvtoA") = 0
                    dt.Rows(0)("PrecioFIFOMvtoB") = 0
                    Select Case valoracion.Precios.CriterioValoracion
                        Case enumtaValoracion.taPrecioFIFOFecha
                            dt.Rows(0)("PrecioFIFOFechaA") = valoracion.Precios.PrecioA
                            dt.Rows(0)("PrecioFIFOFechaB") = valoracion.Precios.PrecioB
                        Case enumtaValoracion.taPrecioFIFOMvto
                            dt.Rows(0)("PrecioFIFOMvtoA") = valoracion.Precios.PrecioA
                            dt.Rows(0)("PrecioFIFOMvtoB") = valoracion.Precios.PrecioB
                        Case enumtaValoracion.taPrecioMedio
                            dt.Rows(0)("PrecioMedioA") = valoracion.Precios.PrecioA
                            dt.Rows(0)("PrecioMedioB") = valoracion.Precios.PrecioB
                    End Select
                    BusinessHelper.UpdateTable(dt)
                End If
                AdminData.CommitTx(True)
            End If
        Next
    End Sub

    Public Enum CriterioValoracion
        Articulo
        PrecioEstandar
        PrecioFifoFecha
        PrecioFifoMvto
        PrecioMedio
        PrecioUltimaCompra
    End Enum

    <Serializable()> _
    Public Class ParametrosValoracion
        Public FechaCalculo As Date
        Public CriterioValoracion As CriterioValoracion
        Public FechaInicioPrecioMedio As Date
        Public Salvar As Boolean
        Public ConRegistros As Boolean = True
        Public ConTotales As Boolean = True
        Public UDValoracion As String

        Public Sub New()
            Me.FechaCalculo = Today
            Me.CriterioValoracion = CriterioValoracion.Articulo
        End Sub

        Public Sub New(ByVal FechaCalculo As Date, ByVal criterio As CriterioValoracion, _
                       ByVal FechaInicioPrecioMedio As Date, ByVal salvar As Boolean, _
                       ByVal ConRegistros As Boolean, ByVal ConTotales As Boolean)
            Me.FechaCalculo = FechaCalculo
            Me.CriterioValoracion = criterio
            Me.FechaInicioPrecioMedio = FechaInicioPrecioMedio
            Me.Salvar = salvar
            Me.ConRegistros = ConRegistros
            Me.ConTotales = ConTotales
        End Sub
    End Class

    <Serializable()> _
    Public Class ParamsValoracionTotales
        Public Data As DataTable
        Public Params As ParametrosValoracion
        Public FilData As New Filter
        Public Sub New()
        End Sub
        Public Sub New(ByVal Data As DataTable, ByVal Params As ParametrosValoracion, Optional ByVal FilData As Filter = Nothing)
            Me.Data = Data
            Me.Params = Params
            Me.FilData = FilData
        End Sub
    End Class

    <Task()> Public Shared Function ValoracionEnFechaTotales(ByVal data As ParamsValoracionTotales, ByVal services As ServiceProvider) As Double
        Dim result As ProcesoStocks.ValoracionResultInfo
        Dim DtDatos As DataTable = New BE.DataEngine().Filter("vFrmCIValoracionAlmacenEnFecha", data.FilData)
        DtDatos.Columns.Add("Stock", GetType(Double))
        DtDatos.Columns.Add("PrecioA", GetType(Double))
        DtDatos.Columns.Add("PrecioB", GetType(Double))
        DtDatos.Columns.Add("ImporteA", GetType(Double))
        DtDatos.Columns.Add("ImporteB", GetType(Double))
        Select Case data.Params.CriterioValoracion
            Case ProcesoStocks.CriterioValoracion.Articulo
                Dim datosVal As New ProcesoStocks.DataValoracionEnFechaPrecioMedio(data.Params.FechaCalculo, DtDatos, data.Params.Salvar, data.Params.FechaInicioPrecioMedio)
                result = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionEnFechaPrecioMedio, ProcesoStocks.ValoracionResultInfo)(AddressOf ProcesoStocks.ValoracionEnFecha, datosVal, services)
            Case ProcesoStocks.CriterioValoracion.PrecioEstandar
                Dim datosValPStd As New ProcesoStocks.DataValoracionEnFecha(data.Params.FechaCalculo, DtDatos)
                result = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionEnFecha, ProcesoStocks.ValoracionResultInfo)(AddressOf ProcesoStocks.ValoracionEnFechaPrecioEstandar, datosValPStd, services)
            Case ProcesoStocks.CriterioValoracion.PrecioFifoFecha
                Dim datosValPFIFO As New ProcesoStocks.DataValoracionEnFechaFIFO(data.Params.FechaCalculo, DtDatos, enumstkValoracionFIFO.stkVFOrdenarPorFecha, data.Params.Salvar)
                result = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionEnFechaFIFO, ProcesoStocks.ValoracionResultInfo)(AddressOf ProcesoStocks.ValoracionEnFechaPrecioFIFO, datosValPFIFO, services)
            Case ProcesoStocks.CriterioValoracion.PrecioFifoMvto
                Dim datosValPFIFO As New ProcesoStocks.DataValoracionEnFechaFIFO(data.Params.FechaCalculo, DtDatos, enumstkValoracionFIFO.stkVFOrdenarPorMvto, data.Params.Salvar)
                result = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionEnFechaFIFO, ProcesoStocks.ValoracionResultInfo)(AddressOf ProcesoStocks.ValoracionEnFechaPrecioFIFO, datosValPFIFO, services)
            Case ProcesoStocks.CriterioValoracion.PrecioMedio
                Dim datosValPM As New ProcesoStocks.DataValoracionEnFechaPrecioMedio(data.Params.FechaCalculo, DtDatos, data.Params.Salvar, data.Params.FechaInicioPrecioMedio)
                result = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionEnFechaPrecioMedio, ProcesoStocks.ValoracionResultInfo)(AddressOf ProcesoStocks.ValoracionEnFechaPrecioMedio, datosValPM, services)
            Case ProcesoStocks.CriterioValoracion.PrecioUltimaCompra
                Dim datosValPUltC As New ProcesoStocks.DataValoracionEnFecha(data.Params.FechaCalculo, DtDatos)
                result = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionEnFecha, ProcesoStocks.ValoracionResultInfo)(AddressOf ProcesoStocks.ValoracionEnFechaPrecioUltimaCompra, datosValPUltC, services)
        End Select

        Dim lstFactoresConversion As New Dictionary(Of String, Double)

        If Not result.Valoraciones Is Nothing AndAlso result.Valoraciones.Length > 0 Then
            Dim dv As DataView = New DataView(DtDatos)
            dv.Sort = "IDArticulo,IDAlmacen"
            Dim index As Integer
            For Each valor As ValoracionInfo In result.Valoraciones
                index = dv.Find(New String(1) {valor.IDArticulo, valor.IDAlmacen})
                If index >= 0 Then
                    dv(index)("Stock") = valor.Stock.StockAFecha
                    dv(index)("StockFisico2") = valor.Stock.StockAFecha2
                    dv(index)("PrecioA") = valor.Precios.PrecioA
                    dv(index)("PrecioB") = valor.Precios.PrecioB
                    dv(index)("ImporteA") = valor.Stock.StockAFecha * valor.Precios.PrecioA
                    dv(index)("ImporteB") = valor.Stock.StockAFecha * valor.Precios.PrecioB

                    Dim Factor As Double = 0
                    Dim Key As String = dv(index)("IDArticulo") & ";" & dv(index)("IDUDInterna") & String.Empty & ";" & data.Params.UDValoracion
                    If lstFactoresConversion.ContainsKey(Key) Then
                        Factor = lstFactoresConversion(Key)
                    Else
                        Dim datFactor As New ArticuloUnidadAB.DatosFactorConversion(dv(index)("IDArticulo"), dv(index)("IDUDInterna") & String.Empty, data.Params.UDValoracion, False)
                        Factor = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, datFactor, services)

                        lstFactoresConversion(Key) = Factor
                    End If

                    dv(index)("FactorUdValoracion") = valor.Stock.StockAFecha * Factor

                    If DtDatos.Columns.Contains("StockAFechaUdValoracion") Then
                        dv(index)("StockAFechaUdValoracion") = dv(index)("Stock") * dv(index)("FactorUdValoracion")
                    End If
                End If
            Next
            Dim DblTotal As Double = dv.Table.Compute("SUM(ImporteA)", Nothing)
            Return DblTotal
        End If
    End Function

#End Region

#End Region

#Region " Mensajes de Stocks "

    Public Class DataMessage
        Public SegundaUnidad As Boolean
        Public Codigo As Integer
        Public Parametros() As String

        Public Sub New(ByVal codigo As Integer, ByVal ParamArray Parametros() As String)
            Me.Codigo = codigo
            Me.Parametros = Parametros
        End Sub
    End Class

    <Task()> Public Shared Function Message(ByVal data As DataMessage, ByVal services As ServiceProvider) As String
        Dim mensaje As String
        Select Case data.Codigo
            Case 0
                mensaje = AdminData.GetMessageText("Ha ocurrido el siguiente error en la actualización del artículo | en el almacén | :")
            Case 1
                mensaje = AdminData.GetMessageText("El número de movimiento no es válido.")
            Case 2
                mensaje = AdminData.GetMessageText("El artículo {0} no tiene gestión de stocks.")
            Case 3
                mensaje = AdminData.GetMessageText("La gestión de stock por lotes es incompatible con la gestión de números de serie para el artículo.")
            Case 4
                mensaje = AdminData.GetMessageText("El período está cerrado. No se permiten movimientos nuevos con fecha documento anterior a la fecha del ultimo período cerrado.")
            Case 5
                mensaje = AdminData.GetMessageText("La fecha de documento es anterior a la fecha del último inventario.")
            Case 6
                mensaje = AdminData.GetMessageText("No hay suficiente stock. El Artículo {0} no está dado de alta en el Almacén.")
            Case 7
                mensaje = AdminData.GetMessageText("No hay suficiente stock. El artículo {0} no permite stocks negativos.")
            Case 8
                mensaje = AdminData.GetMessageText("La cantidad interna no es válida.")
            Case 9
                mensaje = AdminData.GetMessageText("Se ha producido un error en la generación del número de movimiento.")
            Case 10
                mensaje = AdminData.GetMessageText("Actualizado.")
            Case 11
                mensaje = AdminData.GetMessageText("El Lote es obligatorio.")
            Case 12
                mensaje = AdminData.GetMessageText("La cantidad asignada no coincide con la cantidad a asignar.")
            Case 13
                mensaje = AdminData.GetMessageText("El artículo de origen no coincide con el artículo de destino.")
            Case 14
                mensaje = AdminData.GetMessageText("El almacén de origen coincide con el almacén de destino.")
            Case 15
                mensaje = AdminData.GetMessageText("El almacén está bloqueado.")
            Case 16
                mensaje = AdminData.GetMessageText("La cantidad interna es cero.")
            Case 17
                mensaje = AdminData.GetMessageText("El artículo {0} lleva gestión de stock por lotes. No se permiten stocks negativos.")
            Case 18
                mensaje = AdminData.GetMessageText("El lote está bloqueado.")
            Case 19
                mensaje = AdminData.GetMessageText("La ubicación es obligatoria.")
            Case 20
                mensaje = AdminData.GetMessageText("La relación Lote-Ubicación no existe en el almacén.")
            Case 21
                mensaje = AdminData.GetMessageText("La cantidad del ajuste es cero. No se ha generado ningún movimiento.")
            Case 22
                mensaje = AdminData.GetMessageText("El artículo es obligatorio.")
            Case 23
                mensaje = AdminData.GetMessageText("El almacén es obligatorio.")
            Case 24
                mensaje = AdminData.GetMessageText("El movimiento se ha eliminado o no existe.")
            Case 25
                mensaje = AdminData.GetMessageText("La fecha no es válida.")
            Case 26
                mensaje = AdminData.GetMessageText("La fecha de documento no es válida.")
            Case 27
                mensaje = AdminData.GetMessageText("Se ha producido un error en la creación del registro de la entidad artículo-almacén.")
            Case 28
                mensaje = AdminData.GetMessageText("El almacén no está activo.")
            Case 29
                mensaje = AdminData.GetMessageText("No hay suficiente stock para el artículo {0}. No se permiten stocks negativos.")
            Case 30
                mensaje = AdminData.GetMessageText("El tipo de movimiento no coincide con el tipo de movimiento original.")
            Case 31
                mensaje = AdminData.GetMessageText("El número de serie es obligatorio.")
            Case 32
                mensaje = AdminData.GetMessageText("El estado es obligatorio.")
            Case 33
                mensaje = AdminData.GetMessageText("El operario es obligatorio.")
            Case 34
                mensaje = AdminData.GetMessageText("El número de serie ya existe.")
            Case 35
                mensaje = AdminData.GetMessageText("Número de serie actualizado.")
            Case 36
                mensaje = AdminData.GetMessageText("El número de serie se ha eliminado o no existe.")
            Case 37
                mensaje = AdminData.GetMessageText("El estado asignado al número de serie no es compatible con el tipo de movimiento.")
            Case 38
                mensaje = AdminData.GetMessageText("La gestión de portes sólo se puede realizar desde el módulo de ventas.")
            Case 39
                mensaje = AdminData.GetMessageText("No se permite modificar la cantidad de un movimiento correspondiente a un número de serie.")
            Case 40
                mensaje = AdminData.GetMessageText("No actualizado. Se ha producido un error en el movimiento de entrada asociado.")
            Case 41
                mensaje = AdminData.GetMessageText("No actualizado. Se ha producido un error en el movimiento de salida asociado.")
            Case 42
                mensaje = AdminData.GetMessageText("El número de serie {0} ya está dado de baja.")
            Case 43
                mensaje = AdminData.GetMessageText("No se permite realizar el inventario con lotes a fecha | debido a que ya existen movimientos intermedios sin lotes para el artículo | en el almacén |.")
            Case 44
                mensaje = AdminData.GetMessageText("No se permite la corrección en fecha porque entre las fechas inicial y final hay un movimiento de inventario.")
            Case 45
                mensaje = AdminData.GetMessageText("No se han hecho cambios ni en la cantidad, ni en los precios, ni en la fecha de documento del movimiento.")
            Case 46
                mensaje = AdminData.GetMessageText("No se puede realizar un ajuste a fecha de hoy. Ya existe un movimiento de inventario.")
            Case 47
                mensaje = AdminData.GetMessageText("El artículo {0} lleva gestión de números de Serie. No se permiten stocks negativos.")
            Case 48
                mensaje = AdminData.GetMessageText("No actualizado. Se ha producido un error en los Lotes del Artículo.")
            Case 49
                mensaje = AdminData.GetMessageText("Artículo de Bodega. Tiene Movimiento de Inventario posterior |.")
            Case 50
                mensaje = AdminData.GetMessageText("Artículo de Bodega |. Tiene Operaciones posteriores: |.")
            Case Else
                mensaje = String.Empty
        End Select

        If mensaje.Length > 0 Then
            If data.SegundaUnidad Then
                mensaje = mensaje & " Revise los datos de la Segunda Unidad."
            End If
            For Each p As String In data.Parametros
                mensaje = Replace(mensaje, "|", Quoted(p), 1, 1)
            Next
        End If

        Return Engine.ParseFormatString(AdminData.GetMessageText(mensaje), data.Parametros)
    End Function

#End Region

#Region " Obtener Movimientos "

    <Serializable()> _
    Public Class DataObtenerUltimoMovimientoVigente
        Public IDArticulo As String
        Public IDAlmacen As String
        Public IDLineaMovimiento As Integer

        Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal IDLineaMovimiento As Integer)
            Me.IDArticulo = IDArticulo
            Me.IDAlmacen = IDAlmacen
            Me.IDLineaMovimiento = IDLineaMovimiento
        End Sub
    End Class
    ''' <summary>
    ''' Esta función devuelve el último movimiento que no sea de tipo Corrección y diferente del movimiento que tratamos, de acuerdo con el criterio de ordenación establecido:
    ''' FechaDocumento DESC,IDLineaMovimiento DESC
    ''' </summary>
    ''' <param name="data">Objeto con el artículo, el almacén y el IDLineaMovimiento</param>
    ''' <param name="services">Objeto para cacheo de información</param>
    ''' <returns></returns>
    ''' <remarks>Se utiliza en la corrección de movimientos.</remarks>
    <Task()> Public Shared Function ObtenerUltimoMovimientoVigente(ByVal data As DataObtenerUltimoMovimientoVigente, ByVal services As ServiceProvider) As DataRow
        '//Esta funcion devuelve el ultimo movimiento tal cual, de acuerdo con el criterio de ordenacion establecido:
        '//FechaDocumento DESC,IDLineaMovimiento DESC

        '//Se utiliza en la correccion de movimientos. Para el calculo de stock a fecha hay que utilizar la otra version.

        Dim f As New Filter
        f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
        f.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
        f.Add(New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion))
        If data.IDLineaMovimiento <> 0 Then
            f.Add(New NumberFilterItem("IDLineaMovimiento", FilterOperator.NotEqual, data.IDLineaMovimiento))
        End If

        Dim dt As DataTable = New BE.DataEngine().Filter(cnEntidad, f, "top 1 *", "FechaDocumento DESC,IDLineaMovimiento DESC")
        If dt.Rows.Count > 0 Then
            Return dt.Rows(0)
        End If
    End Function


    ''' <summary>
    '''  Obtiene el último movimiento en la fecha especificada, de acuerdo con el criterio de ordenación establecido: 
    '''   FechaDocumento DESC,IDLineaMovimiento DESC
    ''' </summary>
    ''' <param name="data">Objeto con el artículo, el almacén y la Fecha Hasta</param>
    ''' <param name="services">Objeto para cacheo de información</param>
    ''' <returns></returns>
    ''' <remarks>Se utiliza en el calculo del stock a fecha.</remarks>
    <Task()> Public Shared Function ObtenerUltimoMovimientoAFecha(ByVal data As DataArticuloAlmacenFecha, ByVal services As ServiceProvider) As DataRow
        '//Se utiliza en el calculo del stock a fecha. Obtiene el ultimo movimiento en la fecha especificada, de acuerdo
        '//con el criterio de ordenacion establecido: FechaDocumento DESC,IDLineaMovimiento DESC

        Dim AppParamsStocks As ParametroStocks = services.GetService(Of ParametroStocks)()
        Dim ultimoMovimiento As DataRow

        Dim fArticulo As New StringFilterItem("IDArticulo", data.IDArticulo)
        Dim fAlmacen As New StringFilterItem("IDAlmacen", data.IDAlmacen)
        Dim fNoCorreccion As New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion)
        Dim fInventario As New NumberFilterItem("IDTipoMovimiento", enumTipoMovimiento.tmInventario)

        Dim f As New Filter
        f.Add(fArticulo)
        f.Add(fAlmacen)
        f.Add(fNoCorreccion)
        f.Add(New DateFilterItem("FechaDocumento", FilterOperator.LessThanOrEqual, data.Fecha))
        Dim UltimaFecha As DataTable = New BE.DataEngine().Filter(cnEntidad, f, "TOP 1 FechaDocumento", "FechaDocumento DESC,IDLineaMovimiento DESC")
        If Not UltimaFecha Is Nothing AndAlso UltimaFecha.Rows.Count > 0 Then
            '//seleccionar todos los movimientos de ese dia
            f.Clear()
            f.Add(fArticulo)
            f.Add(fAlmacen)
            f.Add(fNoCorreccion)
            f.Add(New DateFilterItem("FechaDocumento", UltimaFecha.Rows(0)("FechaDocumento")))
            '//Se inluyen en el SELECT todos los campos, por temas de eficiencia. Para evitar el SELECT *
            Dim strSELECT As String = "IDLineaMovimiento, IDMovimiento, IDArticulo, IDTipoMovimiento, IDAlmacen, Cantidad, PrecioA, PrecioB, Texto, Lote, Ubicacion, Acumulado, FechaMovimiento, Documento, " & _
                                      "FechaDocumento, IDObra, IDLineaMaterial, IDActivo, Traza, IDDocumento, PrecioMedio, PrecioUltimaCompra, PrecioEstandar, FifoF, FifoFD, Contabilizado, SeriePrecinta, NDesdePrecinta, NHastaPrecinta, NDesdePrecintaUtilizada, NHastaPrecintaUtilizada"
            If AppParamsStocks.GestionDobleUnidad Then
                strSELECT &= ", Cantidad2, Acumulado2"
            End If
            Dim movimientos As DataTable = New BE.DataEngine().Filter(cnEntidad, f, strSELECT, "FechaDocumento DESC,IDLineaMovimiento DESC")
            If movimientos.Rows.Count = 1 Then
                ultimoMovimiento = movimientos.Rows(0)
            ElseIf movimientos.Rows.Count > 1 Then
                Select Case AppParamsStocks.TipoInventario
                    Case TipoInventario.PrimerMovimiento
                        If movimientos.Rows(0)("IDTipoMovimiento") = enumTipoMovimiento.tmInventario Then
                            ultimoMovimiento = movimientos.Rows(1)
                        Else
                            ultimoMovimiento = movimientos.Rows(0)
                        End If
                    Case TipoInventario.UltimoMovimiento
                        Dim MovtosInventario As List(Of DataRow) = (From c In movimientos _
                                                                       Where Not c.IsNull("IDTipoMovimiento") AndAlso _
                                                                             c("IDTipoMovimiento") = enumTipoMovimiento.tmInventario).ToList()

                        If Not MovtosInventario Is Nothing AndAlso MovtosInventario.Count > 0 Then
                            ultimoMovimiento = MovtosInventario(0)
                        Else
                            Dim MovtosOrdenados As List(Of DataRow) = (From c In movimientos _
                                                                       Order By c("FechaDocumento") Descending, c("IDLineaMovimiento") Descending).ToList()
                            If MovtosOrdenados.Count > 0 Then
                                ultimoMovimiento = MovtosOrdenados(0)
                            End If
                        End If
                End Select
            End If
        End If
        Return ultimoMovimiento
    End Function

    <Serializable()> _
    Public Class DataObtenerPrimerMovto
        Public IDArticulo As String
        Public IDAlmacen As String
        Public FechaInicioCalculo As Date
        Public FechaCalculo As Date

        Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal FechaInicioCalculo As Date, ByVal FechaCalculo As Date)
            Me.IDArticulo = IDArticulo
            Me.IDAlmacen = IDAlmacen
            Me.FechaInicioCalculo = FechaInicioCalculo
            Me.FechaCalculo = FechaCalculo
        End Sub
    End Class
    ''' <summary>
    ''' Método que devuelve el Primer Movimiento de Inventario hacia atras en el período indicado. Es decir, el último inventario realizado dentro del período de cálculo.
    ''' </summary>
    ''' <param name="data">Objeto que nos proporciona el Artículo, Almacén y las Fechas Límite del cálculo, para los que se realizará la operación.</param>
    ''' <param name="services">Objeto para información de cacheo</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Task()> Public Shared Function ObtenerPrimerMovimientoInventario(ByVal data As DataObtenerPrimerMovto, ByVal services As ServiceProvider) As ValoresInicialesPrecioMedio
        Dim ValIniPM As New ValoresInicialesPrecioMedio
        Dim f As New Filter
        f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
        f.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
        f.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThanOrEqual, data.FechaInicioCalculo))
        f.Add(New DateFilterItem("FechaDocumento", FilterOperator.LessThanOrEqual, data.FechaCalculo))
        f.Add(New NumberFilterItem("IDTipoMovimiento", enumTipoMovimiento.tmInventario))
        Dim inventario As DataTable = New BE.DataEngine().Filter(cnEntidad, f, "top 1 FechaDocumento,Cantidad,Acumulado,PrecioA,PrecioB", "FechaDocumento DESC")
        If inventario.Rows.Count > 0 Then
            ValIniPM.Stock = inventario.Rows(0)("Acumulado")
            ValIniPM.PrecioA = inventario.Rows(0)("PrecioA")
            ValIniPM.PrecioB = inventario.Rows(0)("PrecioB")
            ValIniPM.Fecha = inventario.Rows(0)("FechaDocumento")
        End If
        Return ValIniPM
    End Function

    ''' <summary>
    ''' Método que devuelve el Primer Movimiento de Entrada (excluyendo correcciones) en el período indicado. Independientemente del origen de la entrada.
    ''' (Ajustes, Albaranes de compra,....)
    ''' </summary>
    ''' <param name="data">Objeto que nos proporciona el Artículo, Almacén y las Fechas Límite del cálculo, para los que se realizará la operación.</param>
    ''' <param name="services">Objeto para información de cacheo</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <Task()> Public Shared Function ObtenerPrimerMovimientoEntrada(ByVal data As DataObtenerPrimerMovto, ByVal services As ServiceProvider) As ValoresInicialesPrecioMedio
        Dim v0 As New ValoresInicialesPrecioMedio
        Dim f As New Filter
        f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
        f.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
        f.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThanOrEqual, data.FechaInicioCalculo))
        f.Add(New DateFilterItem("FechaDocumento", FilterOperator.LessThanOrEqual, data.FechaCalculo))
        f.Add(New NumberFilterItem("ClaseMovimiento", FilterOperator.Equal, enumtpmTipoMovimiento.tpmInput))
        f.Add(New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion))
        Dim entrada As DataTable = New BE.DataEngine().Filter("vNegMovimientosPrecioMedio", f, "top 1 FechaDocumento,Cantidad,PrecioA,PrecioB", "FechaDocumento,IDLineaMovimiento")
        If entrada.Rows.Count > 0 Then
            'v0.Stock = entrada.Rows(0)("Cantidad)
            'v0.PrecioA = entrada.Rows(0)("PrecioA)
            'v0.PrecioB = entrada.Rows(0)("PrecioB)
            v0.Fecha = entrada.Rows(0)("FechaDocumento")
        End If
        Return v0
    End Function

    <Serializable()> _
    Public Class DataObtenerMovtoAnterior
        Public IDArticulo As String
        Public IDAlmacen As String
        Public FechaDocumento As Date
        Public IDLineaActual As Integer

        Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal FechaDocumento As Date, ByVal IDLineaActual As Integer)
            Me.IDArticulo = IDArticulo
            Me.IDAlmacen = IDAlmacen
            Me.FechaDocumento = FechaDocumento
            Me.IDLineaActual = IDLineaActual
        End Sub
    End Class

    <Task()> Public Shared Function ObtenerMovimientoAnterior(ByVal data As DataObtenerMovtoAnterior, ByVal services As ServiceProvider) As DataRow
        Dim AppParamsStocks As ParametroStocks = services.GetService(Of ParametroStocks)()
        Dim tipoInv As TipoInventario = AppParamsStocks.TipoInventario
        Dim UltimoMovimiento As DataRow
        Select Case tipoInv
            Case TipoInventario.PrimerMovimiento
                UltimoMovimiento = ProcessServer.ExecuteTask(Of DataObtenerMovtoAnterior, DataRow)(AddressOf ObtenerMovimientoAnteriorInvPrimerMovimiento, data, services)
            Case TipoInventario.UltimoMovimiento
                UltimoMovimiento = ProcessServer.ExecuteTask(Of DataObtenerMovtoAnterior, DataRow)(AddressOf ObtenerMovimientoAnteriorInvUltimoMovimiento, data, services)
        End Select
        Return UltimoMovimiento
    End Function

    <Task()> Public Shared Function ObtenerMovimientoAnteriorInvPrimerMovimiento(ByVal data As DataObtenerMovtoAnterior, ByVal services As ServiceProvider) As DataRow

        Dim BEDataEngine As New BE.DataEngine
        Dim UltimoMovimiento As DataTable

        Dim f1 As New Filter(FilterUnionOperator.Or)
        Dim f2 As New Filter
        Dim f3 As New Filter

        f2.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
        f2.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
        f2.Add(New NumberFilterItem("IDTipoMovimiento", enumTipoMovimiento.tmInventario))
        f2.Add(New DateFilterItem("FechaDocumento", data.FechaDocumento))
        Dim inventarioPrimerMovimiento As DataTable = BEDataEngine.Filter(cnEntidad, f2, , "FechaDocumento DESC,IDLineaMovimiento DESC")
        If (Not inventarioPrimerMovimiento Is Nothing AndAlso inventarioPrimerMovimiento.Rows.Count > 0) Then
            f2.Clear()
            f2.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
            f2.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
            f2.Add(New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion))
            f2.Add(New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmInventario))
            f2.Add(New DateFilterItem("FechaDocumento", data.FechaDocumento))
            f2.Add(New NumberFilterItem("IDLineaMovimiento", FilterOperator.LessThan, data.IDLineaActual))

            Dim movimientos As DataTable = BEDataEngine.Filter(cnEntidad, f2, "top 1 *", "FechaDocumento DESC,IDLineaMovimiento DESC")
            UltimoMovimiento = movimientos.Clone
            If movimientos.Rows.Count = 0 Then
                UltimoMovimiento.Rows.Add(inventarioPrimerMovimiento(0).ItemArray)
            Else
                UltimoMovimiento.Rows.Add(movimientos(0).ItemArray)
            End If
        Else
            f2.Clear()

            f2.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
            f2.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
            f2.Add(New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion))
            f2.Add(New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmInventario))
            f2.Add(New DateFilterItem("FechaDocumento", data.FechaDocumento))
            f2.Add(New NumberFilterItem("IDLineaMovimiento", FilterOperator.LessThan, data.IDLineaActual))

            f3.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
            f3.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
            f3.Add(New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion))
            f3.Add(New DateFilterItem("FechaDocumento", FilterOperator.LessThan, data.FechaDocumento))

            f1.Add(f2)
            f1.Add(f3)

            Dim movimientos As DataTable = BEDataEngine.Filter(cnEntidad, f1, "top 1 *", "FechaDocumento DESC,IDLineaMovimiento DESC")
            UltimoMovimiento = movimientos.Clone
            If movimientos.Rows.Count > 0 Then
                Dim fechaUltimoMovimiento As Date = movimientos.Rows(0)("FechaDocumento")

                f1.Clear()
                f2.Clear()
                f3.Clear()
                '//Es necesario volver a filtrar por la posibilidad de que en la fecha donde esta el ultimo
                '//movimiento exista a su vez un movimiento de inventario
                f2.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
                f2.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
                f2.Add(New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion))
                f2.Add(New DateFilterItem("FechaDocumento", fechaUltimoMovimiento))
                f2.Add(New NumberFilterItem("IDLineaMovimiento", FilterOperator.NotEqual, data.IDLineaActual))

                movimientos = New BE.DataEngine().Filter(cnEntidad, f2, , "FechaDocumento DESC,IDLineaMovimiento DESC")
                If movimientos.Rows.Count > 0 Then
                    Dim MovtosInventario As List(Of DataRow) = (From c In movimientos _
                                                                Where Not c.IsNull("IDTipoMovimiento") AndAlso _
                                                                            c("IDTipoMovimiento") = enumTipoMovimiento.tmInventario _
                                                                Order By c("FechaDocumento") Descending, c("IDLineaMovimiento") Descending).ToList()
                    If Not MovtosInventario Is Nothing AndAlso MovtosInventario.Count > 0 Then
                        If movimientos.Rows.Count = MovtosInventario.Count Then
                            '//Si hay lotes, los movimientos de inventario para los distintos lotes
                            '//deben tener el mismo acumulado, en ese caso se toma el primer inventario
                            UltimoMovimiento.Rows.Add(movimientos.Rows(0).ItemArray)
                        Else
                            Dim MovtosNoInventario As List(Of DataRow) = (From c In movimientos _
                                                                          Where Not c.IsNull("IDTipoMovimiento") AndAlso _
                                                                            c("IDTipoMovimiento") <> enumTipoMovimiento.tmInventario _
                                                                          Order By c("FechaDocumento") Descending, c("IDLineaMovimiento") Descending).ToList()

                            If Not MovtosNoInventario Is Nothing AndAlso MovtosNoInventario.Count > 0 Then
                                '//Este filtro tiene que dar un registro al menos.
                                '//Los registros ya vienen ordenados.
                                UltimoMovimiento.Rows.Add(MovtosNoInventario(0).ItemArray)
                            End If
                        End If
                    Else
                        If data.FechaDocumento = fechaUltimoMovimiento Then
                            Dim MovtosAnteriores As List(Of DataRow) = (From c In movimientos _
                                                                          Where Not c.IsNull("IDLineaMovimiento") AndAlso _
                                                                                c("IDLineaMovimiento") < data.IDLineaActual _
                                                                          Order By c("FechaDocumento") Descending, c("IDLineaMovimiento") Descending).ToList()
                            If Not MovtosAnteriores Is Nothing AndAlso MovtosAnteriores.Count > 0 Then
                                UltimoMovimiento.Rows.Add(MovtosAnteriores(0).ItemArray)
                            End If
                        Else
                            UltimoMovimiento.Rows.Add(movimientos.Rows(0).ItemArray)
                        End If
                    End If
                End If
            End If
        End If

        If Not UltimoMovimiento Is Nothing AndAlso UltimoMovimiento.Rows.Count > 0 Then
            Return UltimoMovimiento.Rows(0)
        End If
    End Function
    <Task()> Public Shared Function ObtenerMovimientoAnteriorInvUltimoMovimiento(ByVal data As DataObtenerMovtoAnterior, ByVal services As ServiceProvider) As DataRow
        Dim f1 As New Filter(FilterUnionOperator.Or)
        Dim f2 As New Filter
        Dim f3 As New Filter

        f2.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
        f2.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
        f2.Add(New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion))
        f2.Add(New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmInventario))
        f2.Add(New DateFilterItem("FechaDocumento", data.FechaDocumento))
        f2.Add(New NumberFilterItem("IDLineaMovimiento", FilterOperator.LessThan, data.IDLineaActual))

        f3.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
        f3.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
        f3.Add(New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion))
        f3.Add(New DateFilterItem("FechaDocumento", FilterOperator.LessThan, data.FechaDocumento))

        f1.Add(f2)
        f1.Add(f3)

        Dim BEDataEngine As New BE.DataEngine
        Dim movimientos As DataTable = BEDataEngine.Filter(cnEntidad, f1, "top 1 *", "FechaDocumento DESC,IDLineaMovimiento DESC")
        Dim UltimoMovimiento As DataTable = movimientos.Clone
        If movimientos.Rows.Count > 0 Then
            Dim fechaUltimoMovimiento As Date = movimientos.Rows(0)("FechaDocumento")
            If data.FechaDocumento = fechaUltimoMovimiento Then
                UltimoMovimiento.Rows.Add(movimientos.Rows(0).ItemArray)
            Else
                f1.Clear()
                f2.Clear()
                f3.Clear()
                '//Es necesario volver a filtrar por la posibilidad de que en la fecha donde esta el ultimo
                '//movimiento exista a su vez un movimiento de inventario
                f2.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
                f2.Add(New StringFilterItem("IDAlmacen", data.IDAlmacen))
                f2.Add(New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion))
                f2.Add(New DateFilterItem("FechaDocumento", fechaUltimoMovimiento))
                f2.Add(New NumberFilterItem("IDLineaMovimiento", FilterOperator.NotEqual, data.IDLineaActual))

                movimientos = New BE.DataEngine().Filter(cnEntidad, f2)
                If movimientos.Rows.Count > 0 Then
                    Dim MovtosInventario As List(Of DataRow) = (From c In movimientos _
                                                                Where Not c.IsNull("IDTipoMovimiento") AndAlso _
                                                                      c("IDTipoMovimiento") = enumTipoMovimiento.tmInventario _
                                                                Order By c("FechaDocumento") Descending, c("IDLineaMovimiento") Descending).ToList()

                    If Not MovtosInventario Is Nothing AndAlso MovtosInventario.Count > 0 Then
                        UltimoMovimiento.Rows.Add(MovtosInventario(0).ItemArray)
                    Else
                        '//Movimientos de una Fecha Anterior a la FechaDocumento del movimiento actual
                        Dim MovtosOrdenados As List(Of DataRow) = (From c In movimientos _
                                                                   Order By c("FechaDocumento") Descending, c("IDLineaMovimiento") Descending).ToList()
                        If MovtosOrdenados.Count > 0 Then
                            UltimoMovimiento.Rows.Add(MovtosOrdenados(0).ItemArray)
                        End If
                    End If
                End If
            End If
        End If

        If UltimoMovimiento.Rows.Count > 0 Then
            Return UltimoMovimiento.Rows(0)
        End If
    End Function



#End Region

#Region " Entradas "

    <Task()> Public Shared Function Entrada(ByVal data As DataNumeroMovimientoSinc, ByVal services As ServiceProvider) As StockUpdateData
        If Not data Is Nothing Then
            Dim dataEnt As New DataNumeroMovimiento(data.NumeroMovimiento, data.stkData)
            Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimiento, StockUpdateData)(AddressOf EntradaTx, dataEnt, services)
            If Not updateData Is Nothing Then
                If updateData.Estado = EstadoStock.Actualizado Then
                    AdminData.BeginTx()
                    Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
                    Dim AppParamsStk As ParametroStocks = services.GetService(Of ParametroStocks)()
                    If data.Sinc AndAlso AppParams.GestionBodegas AndAlso data.stkData.TipoMovimiento <> AppParamsStk.TipoMovimientoCantidad0 Then
                        Dim datSinc As New DataIntegracionConBodega(data, enumTipoSincronizacion.Entrada, , , , updateData)
                        Dim updateDataAux As StockUpdateData = ProcessServer.ExecuteTask(Of DataIntegracionConBodega, StockUpdateData)(AddressOf IntegracionConBodega, datSinc, services)
                        If Not updateDataAux Is Nothing Then
                            updateData = updateDataAux
                        End If

                    End If
                    ProcessServer.ExecuteTask(Of StockUpdateData)(AddressOf Actualizar, updateData, services)
                End If
                Return updateData
            End If
        End If
    End Function

    <Task()> Public Shared Function ValidarFechaUltimoCierre(ByVal data As StockData, ByVal services As ServiceProvider) As Boolean
        ValidarFechaUltimoCierre = True

        If Not data.ContextCorrect Is Nothing Then
            Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
            If Nz(data.ContextCorrect.CorreccionEnPrecio, False) AndAlso _
                Not Nz(data.ContextCorrect.CorreccionEnFecha, False) AndAlso _
                Not Nz(data.ContextCorrect.CorreccionEnCantidad, False) AndAlso _
                Not Nz(data.ContextCorrect.CorreccionEnCantidad2, False) Then
                If AppParams.ActualizarPrecioAlbaranPeriodoCerrado Then
                    ValidarFechaUltimoCierre = False
                Else
                    ValidarFechaUltimoCierre = True
                End If
            End If
        End If
    End Function
    <Task()> Public Shared Function EntradaTx(ByVal data As DataNumeroMovimiento, ByVal services As ServiceProvider) As StockUpdateData
        '//1.- CREAR CONTEXTO(data.Context). Variable que a lo largo de la ejecución tendrá los datos necesarios en cada momento.
        Dim dataCtx As New DataCrearContexto(data.NumeroMovimiento, data.stkData)
        ProcessServer.ExecuteTask(Of DataCrearContexto)(AddressOf CrearContexto, dataCtx, services)
        If Not data.stkData.Context.Cancel Then
            '//2.- VALIDAR CONTEXTO(data.Context). Realizamos ciertas validaciones relacionadas con el movimiento que estamos realizando.
            Dim stk As New ProcesoStocks
            Dim ProcInfo As ArticuloCosteEstandar.ProcInfoActualizarPrecioEstandar = services.GetService(Of ArticuloCosteEstandar.ProcInfoActualizarPrecioEstandar)()
            Dim Rules() As Regla = stk.EstablecerReglas(Regla.NumeroMovimiento, Regla.LoteBloqueado, Regla.AlmacenActivo, Regla.ArticuloDePortes)
            If ProcInfo Is Nothing OrElse Not ProcInfo.PermitirMovtoCantidad0 Then
                ReDim Preserve Rules(Rules.Length)
                Rules(Rules.Length - 1) = Regla.CantidadCero
            End If
            If ProcessServer.ExecuteTask(Of StockData, Boolean)(AddressOf ValidarFechaUltimoCierre, data.stkData, services) Then
                ReDim Preserve Rules(Rules.Length)
                Rules(Rules.Length - 1) = Regla.FechaUltimoCierre
            End If
            Dim datValCtx As New DataValidarContexto(data.stkData.Articulo, data.stkData.Almacen, data.stkData.Context, Rules)
            ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)
            If data.stkData.Context.Cancel Then
                '// Si algo ha ido mal en la validación, retornamos lo que ha ocurrido.
                Return ProcessServer.ExecuteTask(Of StockData, StockUpdateData)(AddressOf GetStockUpdateData, data.stkData, services)
            End If

            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.stkData.Articulo)

            Dim AppParamsStock As ParametroStocks = services.GetService(Of ParametroStocks)()
            '//Actualizacion del stock
            If (AppParamsStock.TipoInventario = TipoInventario.PrimerMovimiento And data.stkData.FechaDocumento >= data.stkData.Context.FechaUltimoInventario) _
            Or (AppParamsStock.TipoInventario = TipoInventario.UltimoMovimiento And data.stkData.FechaDocumento > data.stkData.Context.FechaUltimoInventario) Then
                '//ACTUALIZAR CONTEXTO. Tratamos el artículo según sus características (Lotes, NSerie, ninguna de ellas)
                If ArtInfo.GestionStockPorLotes Then
                    ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxArticuloAlmacenLoteEntrada, data.stkData, services)
                ElseIf ArtInfo.GestionPorNumeroSerie Then
                    ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxArticuloAlmacenNSerieEntrada, data.stkData, services)
                Else
                    ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxStockFisicoAlmacenEntrada, data.stkData, services)
                End If

                If data.stkData.Context.Cancel Then
                    '// Si algo ha ido mal en la validación, retornamos lo que ha ocurrido.
                    Return ProcessServer.ExecuteTask(Of StockData, StockUpdateData)(AddressOf GetStockUpdateData, data.stkData, services)
                End If
            End If

            '//Generacion del movimiento
            ProcessServer.ExecuteTask(Of StockData)(AddressOf NuevoMovimiento, data.stkData, services)
            ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxArticuloAlmacen, data.stkData, services)
        End If

        '// Retornamos lo que ha ocurrido en el proceso de entrada, para el artículo y almacén indicados.
        Return ProcessServer.ExecuteTask(Of StockData, StockUpdateData)(AddressOf GetStockUpdateData, data.stkData, services)
    End Function

    <Task()> Public Shared Sub SetCtxArticuloAlmacenLoteEntrada(ByVal data As StockData, ByVal services As ServiceProvider)
        Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Articulo, services)
        If data.Context.LoteBBDD.Rows.Count = 0 Then
            Dim newrow As DataRow = data.Context.LoteBBDD.NewRow
            newrow("IDArticulo") = data.Articulo
            newrow("IDAlmacen") = data.Almacen
            newrow("Lote") = data.Lote
            newrow("Ubicacion") = data.Ubicacion
            newrow("StockFisico") = data.Cantidad
            If SegundaUnidad Then newrow("StockFisico2") = data.Cantidad2
            newrow("Bloqueado") = False
            newrow("FechaUltEntrada") = Today
            newrow("Traza") = IIf(data.Traza.Equals(Guid.Empty), DBNull.Value, data.Traza)
            If Not String.IsNullOrEmpty(data.PrecintaNSerie) Then
                newrow("SeriePrecinta") = data.PrecintaNSerie
                newrow("NDesdePrecinta") = data.PrecintaDesde
                newrow("NHastaPrecinta") = data.PrecintaHasta
            End If
            If Length(data.FechaCaducidad) > 0 AndAlso data.FechaCaducidad <> cnMinDate Then newrow("FechaCaducidad") = data.FechaCaducidad
            data.Context.LoteBBDD.Rows.Add(newrow)
            data.Context.StockFisicoLote = newrow("StockFisico")
            If SegundaUnidad Then data.Context.StockFisicoLote2 = CDbl(Nz(newrow("StockFisico2"), 0))
        Else
            data.Context.LoteBBDD.Rows(0)("StockFisico") = data.Context.LoteBBDD.Rows(0)("StockFisico") + data.Cantidad
            If SegundaUnidad Then data.Context.LoteBBDD.Rows(0)("StockFisico2") = Nz(data.Context.LoteBBDD.Rows(0)("StockFisico2"), 0) + Nz(data.Cantidad2, 0)
            If Length(data.FechaCaducidad) > 0 AndAlso data.FechaCaducidad <> cnMinDate Then data.Context.LoteBBDD.Rows(0)("FechaCaducidad") = data.FechaCaducidad
            data.Context.StockFisicoLote = data.Context.LoteBBDD.Rows(0)("StockFisico")
            If SegundaUnidad Then data.Context.StockFisicoLote2 = CDbl(Nz(data.Context.LoteBBDD.Rows(0)("StockFisico2"), 0))
            If data.Cantidad > 0 Then
                data.Context.LoteBBDD.Rows(0)("FechaUltEntrada") = Today
            End If
        End If
        data.Context.StockFisico = data.Context.StockFisico + data.Cantidad
        If SegundaUnidad Then data.Context.StockFisico2 = data.Context.StockFisico2 + data.Cantidad2

        Dim Rules() As Regla = New ProcesoStocks().EstablecerReglas(Regla.StockLoteNegativo)
        Dim datValCtx As New DataValidarContexto(data.Articulo, data.Almacen, data.Context, Rules)
        ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)
    End Sub

    <Task()> Public Shared Sub SetCtxArticuloAlmacenNSerieEntrada(ByVal data As StockData, ByVal services As ServiceProvider)
        Dim stk As New ProcesoStocks
        Dim Rules() As Regla : Dim datValCtx As New DataValidarContexto
        Dim Serie As DataTable = data.Context.SerieBBDD
        If Serie.Rows.Count = 0 Then
            Dim newrow As DataRow = Serie.NewRow
            newrow("IDArticulo") = data.Articulo
            newrow("NSerie") = data.NSerie
            newrow("IDAlmacen") = data.Almacen
            newrow("IDEstadoActivo") = data.EstadoNSerie
            newrow("IDOperario") = data.Operario
            If newrow.Table.Columns.Contains("Ubicacion") Then newrow("Ubicacion") = data.Ubicacion
            newrow("MarcaAuto") = AdminData.GetAutoNumeric()
            Serie.Rows.Add(newrow)

            Rules = stk.EstablecerReglas(Regla.SerieUnica)
            datValCtx = New DataValidarContexto(data.Articulo, data.Almacen, data.Context, Rules)
            ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)

            data.Context.StockFisico += 1
            Dim AppParamsStock As ParametroStocks = services.GetService(Of ParametroStocks)()
            If AppParamsStock.GestionNumeroSerieConActivos Then
                Dim dataAct As New DataActualizarActivo(data, newrow)
                ProcessServer.ExecuteTask(Of DataActualizarActivo)(AddressOf ActualizarActivo, dataAct, services)
            End If
        Else
            Select Case data.Context.TipoMovimiento
                Case enumTipoMovimiento.tmEntAlbaranCompra, enumTipoMovimiento.tmEntSubcontratacion, enumTipoMovimiento.tmEntFabrica
                    If data.Context.CantidadConSigno > 0 Then
                        'Venimos de Albaranes de compra y con un numero de serie que está de baja y lo volvemos a poner en disponible
                        Serie.Rows(0)("IDAlmacen") = data.Almacen
                        data.Context.IDEstadoActivoAnterior = data.EstadoNSerieAnterior
                        Serie.Rows(0)("IDEstadoActivo") = data.EstadoNSerie
                        Serie.Rows(0)("IDOperario") = data.Operario
                        If Serie.Columns.Contains("Ubicacion") Then Serie.Rows(0)("Ubicacion") = data.Ubicacion
                        data.Context.StockFisico += 1
                    ElseIf data.Context.CantidadConSigno < 0 Then
                        'Venimos de Albaranes de compra y con un numero de serie en negativo para pasar de disponible a baja.
                        Serie.Rows(0)("IDAlmacen") = DBNull.Value
                        If Serie.Columns.Contains("Ubicacion") Then Serie.Rows(0)("Ubicacion") = DBNull.Value
                        data.Context.IDEstadoActivoAnterior = data.EstadoNSerieAnterior
                        Serie.Rows(0)("IDEstadoActivo") = data.EstadoNSerie
                        Serie.Rows(0)("IDOperario") = data.Operario
                        data.Context.StockFisico -= 1
                    End If
                Case enumTipoMovimiento.tmEntTransferencia
                    If data.Context.CantidadConSigno > 0 Then
                        Serie.Rows(0)("IDAlmacen") = data.Almacen
                        If Length(data.EstadoNSerieAnterior) <> 0 Then data.Context.IDEstadoActivoAnterior = data.EstadoNSerieAnterior
                        Serie.Rows(0)("IDEstadoActivo") = data.EstadoNSerie
                        Serie.Rows(0)("IDOperario") = data.Operario
                        If Serie.Columns.Contains("Ubicacion") Then Serie.Rows(0)("Ubicacion") = data.Ubicacion
                        data.Context.StockFisico += 1
                    ElseIf data.Context.CantidadConSigno < 0 Then
                        'Serie.Rows(0)("IDAlmacen") = data.Almacen
                        data.Context.IDEstadoActivoAnterior = data.EstadoNSerieAnterior
                        Serie.Rows(0)("IDEstadoActivo") = data.EstadoNSerie
                        Serie.Rows(0)("IDOperario") = data.Operario
                        If Serie.Columns.Contains("Ubicacion") Then Serie.Rows(0)("Ubicacion") = data.Ubicacion
                        data.Context.StockFisico -= 1
                    End If
                Case Else
                    If data.Context.CantidadConSigno <> 0 Then
                        data.Context.Cancel = True
                        data.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(34), services)
                    End If
            End Select

            Dim AppParamsStock As ParametroStocks = services.GetService(Of ParametroStocks)()
            If AppParamsStock.GestionNumeroSerieConActivos Then
                Dim dataAct As New DataActualizarActivo(data, Serie.Rows(0))
                ProcessServer.ExecuteTask(Of DataActualizarActivo)(AddressOf ActualizarActivo, dataAct, services)
                data.Context.Estado = EstadoStock.Actualizado
            End If
        End If

        Rules = stk.EstablecerReglas(Regla.StockSerieNegativo)
        datValCtx = New DataValidarContexto(data.Articulo, data.Almacen, data.Context, Rules)
        ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)

    End Sub

    <Task()> Public Shared Function SetCtxStockFisicoAlmacenEntrada(ByVal data As StockData, ByVal services As ServiceProvider) As StockUpdateData
        data.Context.StockFisico = data.Context.StockFisico + data.Cantidad
        If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Articulo, services) Then
            data.Context.StockFisico2 = data.Context.StockFisico2 + data.Cantidad2
        End If
        Dim Rules() As Regla = New ProcesoStocks().EstablecerReglas(Regla.StockNegativo)
        Dim datValCtx As New DataValidarContexto(data.Articulo, data.Almacen, data.Context, Rules)
        ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)
    End Function

    <Serializable()> _
    Public Class StEntradaTransfer
        Public NumeroMovimiento As Integer
        Public Data As StockData
        Public DataSalida As StockData
        Public UpdateSalida As StockUpdateData
        Public Sinc As Boolean

        Public Sub New()
        End Sub
        Public Sub New(ByVal NumeroMovimiento As Integer, ByVal Data As StockData, ByVal DataSalida As StockData, _
                       ByVal UpdateSalida As StockUpdateData, Optional ByVal Sinc As Boolean = True)
            Me.NumeroMovimiento = NumeroMovimiento
            Me.Data = Data
            Me.DataSalida = DataSalida
            Me.UpdateSalida = UpdateSalida
            Me.Sinc = Sinc
        End Sub
    End Class

    <Task()> Public Shared Function EntradaTransferencia(ByVal data As StEntradaTransfer, ByVal services As ServiceProvider) As StockUpdateData
        If Not data.Data Is Nothing Then
            Dim StNumData As New DataNumeroMovimiento(data.NumeroMovimiento, data.Data)
            Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimiento, StockUpdateData)(AddressOf EntradaTx, StNumData, services)
            If Not updateData Is Nothing Then
                If updateData.Estado = EstadoStock.Actualizado Then
                    If data.Sinc Then
                        Dim StNum As New DataNumeroMovimientoSinc(data.NumeroMovimiento, StNumData.stkData, data.Sinc)
                        Dim StInt As New DataIntegracionConBodega(StNum, enumTipoSincronizacion.EntradaTransferencia, , data.DataSalida, data.UpdateSalida, updateData)
                        Dim updateDataAux As StockUpdateData = ProcessServer.ExecuteTask(Of DataIntegracionConBodega, StockUpdateData)(AddressOf IntegracionConBodega, StInt, services)
                        If Not updateDataAux Is Nothing Then
                            updateData = updateDataAux
                        End If
                    End If
                    ProcessServer.ExecuteTask(Of StockUpdateData)(AddressOf Actualizar, updateData, services)
                End If
                Return updateData
            End If
        End If
    End Function

#End Region

#Region " Salidas "

    <Task()> Public Shared Function Salida(ByVal data As DataNumeroMovimientoSinc, ByVal services As ServiceProvider) As StockUpdateData
        If Not data Is Nothing Then
            Dim dataSal As New DataNumeroMovimiento(data.NumeroMovimiento, data.stkData)
            Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimiento, StockUpdateData)(AddressOf SalidaTx, dataSal, services)
            If Not updateData Is Nothing Then
                If updateData.Estado = EstadoStock.Actualizado Then
                    AdminData.BeginTx()
                    Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
                    Dim AppParamsStk As ParametroStocks = services.GetService(Of ParametroStocks)()
                    If data.Sinc AndAlso AppParams.GestionBodegas AndAlso data.stkData.TipoMovimiento <> AppParamsStk.TipoMovimientoCantidad0 Then
                        Dim datSinc As New DataIntegracionConBodega(data, enumTipoSincronizacion.Salida, , , updateData)
                        Dim updateDataAux As StockUpdateData = ProcessServer.ExecuteTask(Of DataIntegracionConBodega, StockUpdateData)(AddressOf IntegracionConBodega, datSinc, services)
                        If Not updateDataAux Is Nothing Then
                            updateData = updateDataAux
                        End If
                    End If
                    ProcessServer.ExecuteTask(Of StockUpdateData)(AddressOf Actualizar, updateData, services)
                End If
            End If
            Return updateData
        End If
    End Function

    <Task()> Public Shared Function SalidaTx(ByVal data As DataNumeroMovimiento, ByVal services As ServiceProvider) As StockUpdateData
        '//1.- CREAR CONTEXTO(data.Context). Variable que a lo largo de la ejecución tendrá los datos necesarios en cada momento.
        Dim dataCtx As New DataCrearContexto(data.NumeroMovimiento, data.stkData)
        ProcessServer.ExecuteTask(Of DataCrearContexto)(AddressOf CrearContexto, dataCtx, services)
        If Not data.stkData.Context.Cancel Then
            '//2.- VALIDAR CONTEXTO(data.Context). Realizamos ciertas validaciones relacionadas con el movimiento que estamos realizando.
            Dim stk As New ProcesoStocks
            Dim Rules() As Regla = stk.EstablecerReglas(Regla.NumeroMovimiento, Regla.CantidadCero, Regla.LoteBloqueado, Regla.AlbaranVenta, Regla.AlmacenActivo)
            If ProcessServer.ExecuteTask(Of StockData, Boolean)(AddressOf ValidarFechaUltimoCierre, data.stkData, services) Then
                ReDim Preserve Rules(Rules.Length)
                Rules(Rules.Length - 1) = Regla.FechaUltimoCierre
            End If
            Dim datValCtx As New DataValidarContexto(data.stkData.Articulo, data.stkData.Almacen, data.stkData.Context, Rules)
            ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)
            If data.stkData.Context.Cancel Then
                '// Si algo ha ido mal en la validación, retornamos lo que ha ocurrido.
                Return ProcessServer.ExecuteTask(Of StockData, StockUpdateData)(AddressOf GetStockUpdateData, data.stkData, services)
            End If

            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.stkData.Articulo)

            Dim AppParamsStock As ParametroStocks = services.GetService(Of ParametroStocks)()
            '//Actualizacion del stock
            If (AppParamsStock.TipoInventario = TipoInventario.PrimerMovimiento And data.stkData.FechaDocumento >= data.stkData.Context.FechaUltimoInventario) _
            Or (AppParamsStock.TipoInventario = TipoInventario.UltimoMovimiento And data.stkData.FechaDocumento > data.stkData.Context.FechaUltimoInventario) Then
                '//ACTUALIZAR CONTEXTO. Tratamos el artículo según sus características (Lotes, NSerie, ninguna de ellas)
                If ArtInfo.GestionStockPorLotes Then
                    ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxArticuloAlmacenLoteSalida, data.stkData, services)
                ElseIf ArtInfo.GestionPorNumeroSerie Then
                    ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxArticuloAlmacenNSerieSalida, data.stkData, services)
                Else
                    ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxStockFisicoAlmacenSalida, data.stkData, services)
                End If

                If data.stkData.Context.Cancel Then
                    '// Si algo ha ido mal en la validación, retornamos lo que ha ocurrido.
                    Return ProcessServer.ExecuteTask(Of StockData, StockUpdateData)(AddressOf GetStockUpdateData, data.stkData, services)
                End If
            End If

            '//Generacion del movimiento
            ProcessServer.ExecuteTask(Of StockData)(AddressOf NuevoMovimiento, data.stkData, services)
            ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxArticuloAlmacen, data.stkData, services)
        End If

        '// Retornamos lo que ha ocurrido en el proceso de entrada, para el artículo y almacén indicados.
        Return ProcessServer.ExecuteTask(Of StockData, StockUpdateData)(AddressOf GetStockUpdateData, data.stkData, services)
    End Function

    <Task()> Public Shared Sub SetCtxArticuloAlmacenLoteSalida(ByVal data As StockData, ByVal services As ServiceProvider)
        Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Articulo, services)

        If data.Context.LoteBBDD.Rows.Count = 0 Then
            data.Context.Cancel = True
            data.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(20), services)
        Else
            data.Context.LoteBBDD.Rows(0)("StockFisico") = data.Context.LoteBBDD.Rows(0)("StockFisico") - data.Cantidad
            data.Context.StockFisicoLote = data.Context.LoteBBDD.Rows(0)("StockFisico")

            If SegundaUnidad Then
                data.Context.LoteBBDD.Rows(0)("StockFisico2") = Nz(data.Context.LoteBBDD.Rows(0)("StockFisico2"), 0) - Nz(data.Cantidad2, 0)
                data.Context.StockFisicoLote2 = CDbl(data.Context.LoteBBDD.Rows(0)("StockFisico2"))
            End If
        End If
        data.Context.StockFisico = data.Context.StockFisico - data.Cantidad
        If SegundaUnidad Then data.Context.StockFisico2 = data.Context.StockFisico2 - data.Cantidad2

        Dim Rules() As Regla = New ProcesoStocks().EstablecerReglas(Regla.StockLoteNegativo)
        Dim datValCtx As New DataValidarContexto(data.Articulo, data.Almacen, data.Context, Rules)
        ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)
    End Sub

    <Task()> Public Shared Sub SetCtxArticuloAlmacenNSerieSalida(ByVal data As StockData, ByVal services As ServiceProvider)
        Dim Serie As DataTable = data.Context.SerieBBDD
        If Serie.Rows.Count = 0 Then
            data.Context.Cancel = True
            data.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(36), services)
        Else
            If data.Context.TipoMovimiento = enumTipoMovimiento.tmSalTransferencia Then
                Serie.Rows(0)("IDEstadoActivo") = data.Context.EstadoNSerie
                Serie.Rows(0)("IDOperario") = data.Operario
                Serie.Rows(0)("IDAlmacen") = DBNull.Value
                If Serie.Columns.Contains("Ubicacion") Then Serie.Rows(0)("Ubicacion") = DBNull.Value
                data.Context.StockFisico -= 1
            Else
                If data.Context.PropiedadesEstadoBBDD.Baja And Not data.Context.PropiedadesEstado.Disponible Then
                    'Se deja porque bajaUpdateData.Detalle ya viene traduccido
                    Throw New Exception(ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(42, Quoted(data.Context.NSerie)), services))
                Else
                    If Length(data.EstadoNSerieAnterior) <> 0 Then data.Context.IDEstadoActivoAnterior = data.EstadoNSerieAnterior
                    Serie.Rows(0)("IDEstadoActivo") = data.EstadoNSerie
                    Serie.Rows(0)("IDOperario") = data.Operario
                    If data.Context.PropiedadesEstado.Disponible Or data.Context.PropiedadesEstado.EnCurso Then
                        Serie.Rows(0)("IDAlmacen") = data.Almacen
                        If Serie.Columns.Contains("Ubicacion") Then Serie.Rows(0)("Ubicacion") = data.Ubicacion
                    ElseIf data.Context.PropiedadesEstado.Baja Then
                        Serie.Rows(0)("IDAlmacen") = DBNull.Value
                        If Serie.Columns.Contains("Ubicacion") Then Serie.Rows(0)("Ubicacion") = DBNull.Value
                    End If
                    data.Context.StockFisico -= data.Cantidad
                End If
            End If

            Dim AppParamsStock As ParametroStocks = services.GetService(Of ParametroStocks)()
            If AppParamsStock.GestionNumeroSerieConActivos Then
                Dim dataAct As New DataActualizarActivo(data, Serie.Rows(0))
                ProcessServer.ExecuteTask(Of DataActualizarActivo)(AddressOf ActualizarActivo, dataAct, services)
                data.Context.Estado = EstadoStock.Actualizado
            End If

            Dim Rules() As Regla = New ProcesoStocks().EstablecerReglas(Regla.StockSerieNegativo)
            Dim datValCtx As New DataValidarContexto(data.Articulo, data.Almacen, data.Context, Rules)
            ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)
        End If
    End Sub

    <Task()> Public Shared Function SetCtxStockFisicoAlmacenSalida(ByVal data As StockData, ByVal services As ServiceProvider) As StockUpdateData
        data.Context.StockFisico = data.Context.StockFisico - data.Cantidad

        If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Articulo, services) Then
            data.Context.StockFisico2 = data.Context.StockFisico2 - data.Cantidad2
        End If

        Dim Rules() As Regla = New ProcesoStocks().EstablecerReglas(Regla.StockNegativo)
        Dim datValCtx As New DataValidarContexto(data.Articulo, data.Almacen, data.Context, Rules)
        ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)
    End Function
#End Region

#Region " Ajustes "

    <Task()> Public Shared Function AjusteMasivo(ByVal data As DataTratarStocks, ByVal services As ServiceProvider) As StockUpdateData()
        Dim UpdateData(-1) As StockUpdateData
        If Not data Is Nothing AndAlso Not data.Items Is Nothing AndAlso data.Items.Length > 0 Then
            Dim N As Integer = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf NuevoNumeroMovimiento, Nothing, services)

            Dim DeshacerActualizacion As Boolean
            Dim IDArticuloAnt As String
            Dim IDAlmacenAnt As String
            Dim NumMovtoAnt As Integer

            For Each item As StockData In data.Items
                If IDArticuloAnt <> item.Articulo OrElse IDAlmacenAnt <> item.Almacen Then
                    '//Si existe una transacción abierta, se cierra para tratar de controlar
                    '//las transacciones de los Lotes de un mismo Artículo-Almacen
                    If DeshacerActualizacion Then
                        If AdminData.ExistsTX Then AdminData.RollBackTx()
                    Else
                        If AdminData.ExistsTX Then AdminData.CommitTx(True)
                    End If
                    AdminData.BeginTx()

                    IDArticuloAnt = item.Articulo : IDAlmacenAnt = item.Almacen
                Else
                    DeshacerActualizacion = False
                End If
                If Not AdminData.ExistsTX Then AdminData.BeginTx()

                Dim datAj As New DataNumeroMovimientoSinc(N, item, data.Sinc)
                Dim aux As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf Ajuste, datAj, services)
                ' ArrayManager.Copy(aux, UpdateData)

                If (Not aux Is Nothing AndAlso aux.Estado = EstadoStock.NoActualizado) OrElse DeshacerActualizacion Then
                    If (Not aux Is Nothing AndAlso aux.Estado = EstadoStock.NoActualizado) AndAlso _
                      Length(item.Lote) > 0 AndAlso aux.StockData.Cantidad <> aux.StockData.Context.StockFisicoLote Then
                        DeshacerActualizacion = True
                    End If

                    If DeshacerActualizacion Then
                        For Each ud As StockUpdateData In UpdateData
                            If ud.StockData.Articulo = aux.StockData.Articulo AndAlso _
                               ud.StockData.Almacen = aux.StockData.Almacen AndAlso _
                               ud.NumeroMovimiento = aux.NumeroMovimiento AndAlso _
                               ud.Estado = EstadoStock.Actualizado Then

                                '//"Deshacemos" la actualización del resto de lotes del artículo, que no sean el que ha provocado la desactualización
                                ud.Estado = EstadoStock.NoActualizado

                                Dim datMsg As New DataMessage(48)
                                ud.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, datMsg, services)
                            End If
                        Next

                        If Not aux Is Nothing AndAlso aux.Estado = EstadoStock.Actualizado AndAlso DeshacerActualizacion Then
                            '//"Deshacemos" la actualización del resto de lotes del artículo
                            aux.Estado = EstadoStock.NoActualizado

                            Dim datMsg As New DataMessage(48)
                            aux.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, datMsg, services)
                        End If
                    End If
                End If
                ArrayManager.Copy(aux, UpdateData)

                '//Controlamos el Commit/Rollback de la última vuelta del bucle
                If DeshacerActualizacion Then
                    If AdminData.ExistsTX Then AdminData.RollBackTx()
                Else
                    If AdminData.ExistsTX Then AdminData.CommitTx(True)
                End If
            Next
        End If


        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If AppParams.GestionInventarioPermanente Then
            Dim datIStock As New ProcesoStocks.DataCreateIStockClass(InventariosPermanentes.ENSAMBLADO_INV_PERMANENTE_STOCKS, InventariosPermanentes.CLASE_INV_PERMANENTE_STOCKS)
            Dim IStockClass As IStockInventarioPermanente = ProcessServer.ExecuteTask(Of ProcesoStocks.DataCreateIStockClass, IStockInventarioPermanente)(AddressOf ProcesoStocks.CreateIStockInventarios, datIStock, services)
            If Not IStockClass Is Nothing Then
                IStockClass.SincronizarContaMovimientos(UpdateData, services)
            End If
        End If

        Return UpdateData
    End Function

    <Task()> Public Shared Function Ajuste(ByVal data As DataNumeroMovimientoSinc, ByVal services As ServiceProvider) As StockUpdateData
        If Not data Is Nothing Then
            Dim dataAjte As New DataNumeroMovimiento(data.NumeroMovimiento, data.stkData)
            Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimiento, StockUpdateData)(AddressOf AjusteTx, dataAjte, services)
            If Not updateData Is Nothing Then
                If updateData.Estado = EstadoStock.Actualizado Then
                    AdminData.BeginTx()
                    Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
                    If data.Sinc AndAlso AppParams.GestionBodegas Then
                        Dim datSinc As New DataIntegracionConBodega(data, enumTipoSincronizacion.Ajuste, , , updateData)
                        Dim updateDataAux As StockUpdateData = ProcessServer.ExecuteTask(Of DataIntegracionConBodega, StockUpdateData)(AddressOf IntegracionConBodega, datSinc, services)
                        If Not updateDataAux Is Nothing Then
                            updateData = updateDataAux
                        End If
                    End If
                    ProcessServer.ExecuteTask(Of StockUpdateData)(AddressOf Actualizar, updateData, services)
                End If
                Return updateData
            End If
        End If
    End Function

    <Task()> Public Shared Function AjusteTx(ByVal data As DataNumeroMovimiento, ByVal services As ServiceProvider) As StockUpdateData
        data.stkData.FechaDocumento = Today
        Dim stk As New ProcesoStocks
        '//1.- CREAR CONTEXTO(data.Context). Variable que a lo largo de la ejecución tendrá los datos necesarios en cada momento.
        Dim dataCtx As New DataCrearContexto(data.NumeroMovimiento, data.stkData)
        ProcessServer.ExecuteTask(Of DataCrearContexto)(AddressOf CrearContexto, dataCtx, services)
        If Not data.stkData.Context.Cancel Then
            Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.stkData.Articulo, services)

            '//2.- VALIDAR CONTEXTO(data.Context). Realizamos ciertas validaciones relacionadas con el movimiento que estamos realizando.
            Dim Rules() As Regla = stk.EstablecerReglas(Regla.NumeroMovimiento, Regla.FechaUltimoCierre, Regla.CantidadPositiva, Regla.ArticuloDePortes, Regla.FechaAjuste)
            Dim datValCtx As New DataValidarContexto(data.stkData.Articulo, data.stkData.Almacen, data.stkData.Context, Rules)
            ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)
            If data.stkData.Context.Cancel Then
                '// Si algo ha ido mal en la validación, retornamos lo que ha ocurrido.
                Return ProcessServer.ExecuteTask(Of StockData, StockUpdateData)(AddressOf GetStockUpdateData, data.stkData, services)
            End If

            '//ACTUALIZAR CONTEXTO: Tratamos el artículo según sus características (Lotes, NSerie, ninguna de ellas)
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.stkData.Articulo)
            If ArtInfo.GestionStockPorLotes Then
                ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxArticuloAlmacenLoteAjuste, data.stkData, services)
            ElseIf ArtInfo.GestionPorNumeroSerie Then
                ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxArticuloAlmacenNSerieAjuste, data.stkData, services)
                data.stkData.Lote = data.stkData.NSerie
            Else
                ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxStockFisicoAlmacenAjuste, data.stkData, services)
            End If

            If data.stkData.Context.Cancel Then
                '// Si algo ha ido mal en la validación, retornamos lo que ha ocurrido.
                Return ProcessServer.ExecuteTask(Of StockData, StockUpdateData)(AddressOf GetStockUpdateData, data.stkData, services)
            End If

            If data.stkData.Context.QAjuste = 0 OrElse (SegundaUnidad AndAlso data.stkData.Context.QAjuste = 0 AndAlso data.stkData.Context.QAjuste2 = 0) Then
                '//La cantidad del ajuste es cero. No se ha generado ningún movimiento.
                Dim datMsg As New DataMessage(21)
                datMsg.SegundaUnidad = SegundaUnidad
                data.stkData.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, datMsg, services)
                Return ProcessServer.ExecuteTask(Of StockData, StockUpdateData)(AddressOf GetStockUpdateData, data.stkData, services)
            End If
            data.stkData.Context.CantidadConSigno = data.stkData.Context.QAjuste
            If SegundaUnidad Then data.stkData.Context.CantidadConSigno2 = data.stkData.Context.QAjuste2

            '//Generacion del movimiento
            ProcessServer.ExecuteTask(Of StockData)(AddressOf NuevoMovimiento, data.stkData, services)

            ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxArticuloAlmacen, data.stkData, services)
        End If

        '// Retornamos lo que ha ocurrido en el proceso de ajuste, para el artículo y almacén indicados.
        Return ProcessServer.ExecuteTask(Of StockData, StockUpdateData)(AddressOf GetStockUpdateData, data.stkData, services)
    End Function

    <Task()> Public Shared Sub SetCtxArticuloAlmacenLoteAjuste(ByVal data As StockData, ByVal services As ServiceProvider)
        Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Articulo, services)

        If data.Context.LoteBBDD.Rows.Count = 0 Then
            Dim newrow As DataRow = data.Context.LoteBBDD.NewRow
            newrow("IDArticulo") = data.Articulo
            newrow("IDAlmacen") = data.Almacen
            newrow("Lote") = data.Lote
            newrow("Ubicacion") = data.Ubicacion
            newrow("StockFisico") = data.Cantidad
            If SegundaUnidad Then newrow("StockFisico2") = data.Cantidad2
            newrow("Bloqueado") = False
            newrow("Traza") = IIf(data.Traza.Equals(Guid.Empty), DBNull.Value, data.Traza)
            data.Context.LoteBBDD.Rows.Add(newrow)
        Else
            data.Context.LoteBBDD.Rows(0)("StockFisico") = data.Cantidad
            If SegundaUnidad Then data.Context.LoteBBDD.Rows(0)("StockFisico2") = data.Cantidad2
        End If

        ''//Si tenemos doble unidad, haremos que si el stockfisico o el stockfisico2 se queda a 0, el otro tb se quedará a 0.
        'If data.Context.StockFisico = 0 OrElse (SegundaUnidad AndAlso data.Context.StockFisico2 = 0) Then
        '    If SegundaUnidad Then
        '        data.Context.StockFisico = 0
        '        data.Context.StockFisico2 = 0
        '    End If
        'End If

        Dim StockLoteInicial As Double = data.Context.StockFisicoLote
        data.Context.StockFisico = data.Context.StockFisico + data.Cantidad - data.Context.StockFisicoLote
        data.Context.StockFisicoLote = data.Cantidad
        data.Context.QAjuste = data.Context.StockFisicoLote - StockLoteInicial

        If SegundaUnidad Then
            Dim StockLoteInicial2 As Double = data.Context.StockFisicoLote2
            data.Context.StockFisico2 = data.Context.StockFisico2 + data.Cantidad2 - data.Context.StockFisicoLote2
            data.Context.StockFisicoLote2 = data.Cantidad2
            data.Context.QAjuste2 = data.Context.StockFisicoLote2 - StockLoteInicial2
        End If

        If data.Context.QAjuste > 0 Then
            data.Context.TipoMovimiento = enumTipoMovimiento.tmEntAjuste
        ElseIf data.Context.QAjuste < 0 Then
            data.Context.TipoMovimiento = enumTipoMovimiento.tmSalAjuste
        End If

        Dim Rules() As Regla = New ProcesoStocks().EstablecerReglas(Regla.StockLoteNegativo)
        Dim datValCtx As New DataValidarContexto(data.Articulo, data.Almacen, data.Context, Rules)
        ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)
    End Sub

    <Task()> Public Shared Sub SetCtxArticuloAlmacenNSerieAjuste(ByVal data As StockData, ByVal services As ServiceProvider)
        Dim stk As New ProcesoStocks
        Dim Rules() As Regla : Dim datValCtx As New DataValidarContexto
        Dim Serie As DataTable = data.Context.SerieBBDD
        If Serie.Rows.Count = 0 Then
            Dim newrow As DataRow = Serie.NewRow
            newrow("IDArticulo") = data.Articulo
            newrow("NSerie") = data.NSerie
            newrow("IDAlmacen") = data.Almacen
            newrow("IDEstadoActivo") = data.EstadoNSerie
            newrow("IDOperario") = data.Operario
            If newrow.Table.Columns.Contains("Ubicacion") Then newrow("Ubicacion") = data.Ubicacion
            newrow("MarcaAuto") = AdminData.GetAutoNumeric()
            Serie.Rows.Add(newrow)

            Rules = stk.EstablecerReglas(Regla.SerieUnica)
            datValCtx = New DataValidarContexto(data.Articulo, data.Almacen, data.Context, Rules)
            ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)

            data.Context.StockFisico += 1

            data.Context.TipoMovimiento = enumTipoMovimiento.tmEntAjuste
            data.Context.ClaseMovimiento = enumtpmTipoMovimiento.tpmInput
            data.Context.QAjuste = 1

            Dim AppParamsStock As ParametroStocks = services.GetService(Of ParametroStocks)()
            If AppParamsStock.GestionNumeroSerieConActivos Then
                Dim dataAct As New DataActualizarActivo(data, newrow)
                ProcessServer.ExecuteTask(Of DataActualizarActivo)(AddressOf ActualizarActivo, dataAct, services)
            End If
        Else
            If Serie.Rows(0)("IDEstadoActivo") <> data.EstadoNSerie Then
                Serie.Rows(0)("IDEstadoActivo") = data.EstadoNSerie
                Serie.Rows(0)("IDAlmacen") = data.Almacen
                If Serie.Columns.Contains("Ubicacion") Then Serie.Rows(0)("Ubicacion") = data.Ubicacion
                If data.Context.PropiedadesEstado.Baja AndAlso data.Context.PropiedadesEstadoBBDD.Disponible Then
                    Serie.Rows(0)("IDAlmacen") = DBNull.Value
                    If Serie.Columns.Contains("Ubicacion") Then Serie.Rows(0)("Ubicacion") = DBNull.Value
                    data.Context.StockFisico -= 1
                    data.Context.TipoMovimiento = enumTipoMovimiento.tmSalAjuste
                    data.Context.ClaseMovimiento = enumtpmTipoMovimiento.tpmOutput
                    data.Context.QAjuste = -1
                ElseIf data.Context.PropiedadesEstado.Disponible AndAlso data.Context.PropiedadesEstadoBBDD.Baja Then
                    Serie.Rows(0)("IDAlmacen") = data.Almacen
                    If Serie.Columns.Contains("Ubicacion") Then Serie.Rows(0)("Ubicacion") = data.Ubicacion
                    data.Context.StockFisico += 1
                    data.Context.TipoMovimiento = enumTipoMovimiento.tmEntAjuste
                    data.Context.ClaseMovimiento = enumtpmTipoMovimiento.tpmInput
                    data.Context.QAjuste = 1
                End If
                Dim AppParamsStock As ParametroStocks = services.GetService(Of ParametroStocks)()
                If AppParamsStock.GestionNumeroSerieConActivos Then
                    Dim dataAct As New DataActualizarActivo(data, Serie.Rows(0))
                    ProcessServer.ExecuteTask(Of DataActualizarActivo)(AddressOf ActualizarActivo, dataAct, services)
                    data.Context.Estado = EstadoStock.Actualizado
                End If
            End If
            If Nz(Serie.Rows(0)("IDOperario"), -1) <> data.Operario Then
                Serie.Rows(0)("IDOperario") = data.Operario
            End If
        End If

        Rules = stk.EstablecerReglas(Regla.StockSerieNegativo)
        datValCtx = New DataValidarContexto(data.Articulo, data.Almacen, data.Context, Rules)
        ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)
    End Sub

    <Task()> Public Shared Sub SetCtxStockFisicoAlmacenAjuste(ByVal data As StockData, ByVal services As ServiceProvider)
        Dim StockFisicoInicial As Double = data.Context.StockFisico
        data.Context.StockFisico = data.Cantidad
        data.Context.QAjuste = data.Context.StockFisico - StockFisicoInicial

        If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Articulo, services) Then
            Dim StockFisicoInicial2 As Double = data.Context.StockFisico2
            data.Context.StockFisico2 = data.Cantidad2
            data.Context.QAjuste2 = data.Context.StockFisico2 - StockFisicoInicial2
        End If

        If data.Context.QAjuste > 0 Then
            data.Context.TipoMovimiento = enumTipoMovimiento.tmEntAjuste
            data.Context.ClaseMovimiento = enumtpmTipoMovimiento.tpmInput
        ElseIf data.Context.QAjuste < 0 Then
            data.Context.TipoMovimiento = enumTipoMovimiento.tmSalAjuste
            data.Context.ClaseMovimiento = enumtpmTipoMovimiento.tpmOutput
        End If

        Dim Rules() As Regla = New ProcesoStocks().EstablecerReglas(Regla.StockNegativo)
        Dim datValCtx As New DataValidarContexto(data.Articulo, data.Almacen, data.Context, Rules)
        ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)
    End Sub

#End Region

#Region " Inventarios "

    <Task()> Public Shared Function InventarioMasivo(ByVal data As DataTratarStocks, ByVal services As ServiceProvider) As StockUpdateData()
        Dim UpdateData(-1) As StockUpdateData
        If Not IsNothing(data) AndAlso data.Items.Length > 0 Then
            Dim NumeroMovimiento As Integer = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf NuevoNumeroMovimiento, Nothing, services)
            ProcessServer.ExecuteTask(Of StockData())(AddressOf CalculoAcumuladoArticuloAlmacen, data.Items, services)

            Dim DeshacerActualizacion As Boolean
            Dim IDArticuloAnt As String
            Dim IDAlmacenAnt As String
            Dim NumMovtoAnt As Integer

            For Each item As StockData In data.Items
                If IDArticuloAnt <> item.Articulo OrElse IDAlmacenAnt <> item.Almacen Then
                    '//Si tengo una transacción abierta, la cierro para tratar de controlar 
                    '//las transacciones de los Lotes de un mismo Artículo-Almacén
                    If DeshacerActualizacion Then
                        If AdminData.ExistsTX Then ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.RollbackTransaction, True, services)
                    Else
                        If AdminData.ExistsTX Then ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.CommitTransaction, True, services)
                    End If
                    ProcessServer.ExecuteTask(Of Object)(AddressOf Comunes.BeginTransaction, Nothing, services)
                    IDArticuloAnt = item.Articulo : IDAlmacenAnt = item.Almacen
                    DeshacerActualizacion = False
                End If

                Dim datInv As New DataNumeroMovimientoSinc(NumeroMovimiento, item, data.Sinc)
                Dim aux As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf Inventario, datInv, services)
                If (Not aux Is Nothing AndAlso aux.Estado = EstadoStock.NoActualizado) OrElse DeshacerActualizacion Then
                    For Each ud As StockUpdateData In UpdateData
                        If ud.StockData.Articulo = aux.StockData.Articulo AndAlso _
                           ud.StockData.Almacen = aux.StockData.Almacen AndAlso _
                           ud.NumeroMovimiento = aux.NumeroMovimiento AndAlso _
                           ud.Estado = EstadoStock.Actualizado Then
                            ud.Estado = EstadoStock.NoActualizado
                            Dim datMsg As New DataMessage(48)
                            ud.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, datMsg, services)
                        End If
                    Next

                    If Not aux Is Nothing AndAlso aux.Estado = EstadoStock.Actualizado AndAlso DeshacerActualizacion Then
                        aux.Estado = EstadoStock.NoActualizado
                        Dim datMsg As New DataMessage(48)
                        aux.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, datMsg, services)
                    End If

                    DeshacerActualizacion = True
                End If
                ArrayManager.Copy(aux, UpdateData)
            Next

            '//Controlamos el Commit/Rollback de la última vuelta del bucle
            If DeshacerActualizacion Then
                If AdminData.ExistsTX Then ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.RollbackTransaction, True, services)
            Else
                If AdminData.ExistsTX Then ProcessServer.ExecuteTask(Of Boolean)(AddressOf Comunes.CommitTransaction, True, services)
            End If
        End If

        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If AppParams.GestionInventarioPermanente Then
            Dim datIStock As New ProcesoStocks.DataCreateIStockClass(InventariosPermanentes.ENSAMBLADO_INV_PERMANENTE_STOCKS, InventariosPermanentes.CLASE_INV_PERMANENTE_STOCKS)
            Dim IStockClass As IStockInventarioPermanente = ProcessServer.ExecuteTask(Of ProcesoStocks.DataCreateIStockClass, IStockInventarioPermanente)(AddressOf ProcesoStocks.CreateIStockInventarios, datIStock, services)
            If Not IStockClass Is Nothing Then
                IStockClass.SincronizarContaMovimientos(UpdateData, services)
            End If
        End If

        Return UpdateData
    End Function



    <Task()> Public Shared Sub CalculoAcumuladoArticuloAlmacen(ByVal data() As StockData, ByVal services As ServiceProvider)
        If Not IsNothing(data) AndAlso data.Length > 0 Then
            '//Se calcula el acumulado agrupando por Articulo-Almacen.
            '//Este acumulado solo se leera en la funcion 'Acumulado' para los articulos que llevan gestion por lotes.
            '//El Inventario de un articulo por lotes deberia hacerse desde esta funcion para un calculo correcto 
            '//del acumulado.
            Dim mAcumuladoInfo(-1) As AcumuladoInfo
            For Each item As StockData In data
                Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, item.Articulo, services)
                Dim infoItem As AcumuladoInfo
                Dim encontrado As Boolean = False
                For Each a As AcumuladoInfo In mAcumuladoInfo
                    If a.IDArticulo = item.Articulo And a.IDAlmacen = item.Almacen Then
                        infoItem = a
                        encontrado = True
                        Exit For
                    End If
                Next
                If Not encontrado Then
                    infoItem = New AcumuladoInfo(item.Articulo, item.Almacen, 0)
                    If SegundaUnidad Then infoItem.Acumulado2 = CDbl(0)
                    ReDim Preserve mAcumuladoInfo(mAcumuladoInfo.Length)
                    mAcumuladoInfo(mAcumuladoInfo.Length - 1) = infoItem
                End If
                If Length(item.NSerie) > 0 Then
                    If item.EstadoNSerie <> item.EstadoNSerieAnterior Then
                        Dim estado As DataTable = New BE.DataEngine().Filter("tbMntoEstadoActivo", New StringFilterItem("IDEstadoActivo", item.EstadoNSerie))
                        If estado.Rows.Count > 0 Then
                            If estado.Rows(0)("Disponible") Then
                                infoItem.Acumulado = infoItem.Acumulado + item.Cantidad
                            End If
                        End If
                    End If
                Else
                    infoItem.Acumulado = infoItem.Acumulado + item.Cantidad
                    If SegundaUnidad Then infoItem.Acumulado2 = infoItem.Acumulado2 + item.Cantidad2
                End If
            Next
            services.RegisterService(mAcumuladoInfo)
        End If

    End Sub

    <Task()> Public Shared Function Inventario(ByVal data As DataNumeroMovimientoSinc, ByVal services As ServiceProvider) As StockUpdateData
        If Not data Is Nothing Then
            Dim datInv As New DataNumeroMovimiento(data.NumeroMovimiento, data.stkData)
            Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimiento, StockUpdateData)(AddressOf InventarioTx, datInv, services)
            If Not updateData Is Nothing Then
                If updateData.Estado = EstadoStock.Actualizado Then
                    AdminData.BeginTx()
                    Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
                    If data.Sinc AndAlso AppParams.GestionBodegas Then
                        Dim datSinc As New DataIntegracionConBodega(data, enumTipoSincronizacion.Inventario, , , updateData)
                        Dim updateDataAux As StockUpdateData = ProcessServer.ExecuteTask(Of DataIntegracionConBodega, StockUpdateData)(AddressOf IntegracionConBodega, datSinc, services)
                        If Not updateDataAux Is Nothing Then
                            updateData = updateDataAux
                        End If
                    End If
                    ProcessServer.ExecuteTask(Of StockUpdateData)(AddressOf Actualizar, updateData, services)
                End If
                Return updateData
            End If
        End If
    End Function

    <Task()> Public Shared Function InventarioTx(ByVal data As DataNumeroMovimiento, ByVal services As ServiceProvider) As StockUpdateData
        '//1.- CREAR CONTEXTO(data.Context). Variable que a lo largo de la ejecución tendrá los datos necesarios en cada momento.
        Dim stk As New ProcesoStocks
        Dim dataCtx As New DataCrearContexto(data.NumeroMovimiento, data.stkData)
        ProcessServer.ExecuteTask(Of DataCrearContexto)(AddressOf CrearContexto, dataCtx, services)
        If Not data.stkData.Context.Cancel Then
            '//2.- VALIDAR CONTEXTO(data.Context). Realizamos ciertas validaciones relacionadas con el movimiento que estamos realizando.
            Dim Rules() As Regla = stk.EstablecerReglas(Regla.NumeroMovimiento, Regla.FechaUltimoCierre, Regla.FechaUltimoInventario, Regla.CantidadPositiva, Regla.ArticuloDePortes)
            Dim datValCtx As New DataValidarContexto(data.stkData.Articulo, data.stkData.Almacen, data.stkData.Context, Rules)
            ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)
            If data.stkData.Context.Cancel Then
                '// Si algo ha ido mal en la validación, retornamos lo que ha ocurrido.
                Return ProcessServer.ExecuteTask(Of StockData, StockUpdateData)(AddressOf GetStockUpdateData, data.stkData, services)
            End If

            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.stkData.Articulo)

            '//ACTUALIZAR CONTEXTO. Tratamos el artículo según sus características (Lotes, NSerie, ninguna de ellas)
            If ArtInfo.GestionStockPorLotes Then
                ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxArticuloAlmacenLoteInventario, data.stkData, services)
            ElseIf ArtInfo.GestionPorNumeroSerie Then
                ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxArticuloAlmacenNSerieInventario, data.stkData, services)
            Else
                ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxStockFisicoAlmacenInventario, data.stkData, services)
            End If

            If data.stkData.Context.Cancel Then
                '//Si algo ha ido mal en la validación, retornamos lo que ha ocurrido.
                Return ProcessServer.ExecuteTask(Of StockData, StockUpdateData)(AddressOf GetStockUpdateData, data.stkData, services)
            End If

            '// CORRECCION INVENTARIO
            If data.stkData.FechaDocumento = data.stkData.Context.FechaUltimoInventario Then
                ProcessServer.ExecuteTask(Of Object)(AddressOf CorreccionInventario, data.stkData, services)
            End If

            '// Generación del movimiento
            ProcessServer.ExecuteTask(Of StockData)(AddressOf NuevoMovimiento, data.stkData, services)

            ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxArticuloAlmacen, data.stkData, services)

            '//Actualizar la cantidad a pasar en los casos de integración
            Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
            If AppParams.GestionBodegas Then ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCandidadIntegracionBodega, data.stkData, services)
        End If

        '// Retornamos lo que ha ocurrido en el proceso de inventariado, para el artículo y almacén indicados.
        Return ProcessServer.ExecuteTask(Of StockData, StockUpdateData)(AddressOf GetStockUpdateData, data.stkData, services)
    End Function

    <Task()> Public Shared Sub SetCtxArticuloAlmacenLoteInventario(ByVal data As StockData, ByVal services As ServiceProvider)
        Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Articulo, services)
        'data.Context.QIntermedia = ProcessServer.ExecuteTask(Of StockData, Double)(AddressOf ContabilizarMovimientosIntermedios, data, services)
        ProcessServer.ExecuteTask(Of StockData)(AddressOf ContabilizarMovimientosIntermedios, data, services)
        If data.Context.LoteBBDD.Rows.Count = 0 Then
            Dim newrow As DataRow = data.Context.LoteBBDD.NewRow
            newrow("IDArticulo") = data.Articulo
            newrow("IDAlmacen") = data.Almacen
            newrow("Lote") = data.Lote
            newrow("Ubicacion") = data.Ubicacion
            newrow("StockFisico") = data.Cantidad + data.Context.QIntermedia
            If SegundaUnidad Then newrow("StockFisico2") = data.Cantidad2 + data.Context.QIntermedia2
            newrow("Bloqueado") = False
            newrow("Traza") = IIf(data.Traza.Equals(Guid.Empty), DBNull.Value, data.Traza)
            If data.FechaCaducidad <> cnMinDate Then newrow("FechaCaducidad") = data.FechaCaducidad
            newrow("FechaUltEntrada") = Today
            data.Context.LoteBBDD.Rows.Add(newrow)
        Else
            data.Context.LoteBBDD.Rows(0)("StockFisico") = data.Cantidad + data.Context.QIntermedia
            If data.FechaCaducidad <> cnMinDate Then data.Context.LoteBBDD.Rows(0)("FechaCaducidad") = data.FechaCaducidad
            If SegundaUnidad Then data.Context.LoteBBDD.Rows(0)("StockFisico2") = data.Cantidad2 + data.Context.QIntermedia2
        End If

        '//No se generan movimientos para los lotes que ya tienen stock fisico cero
        '//y que se vienen inventariados a cero tambien
        '//(ctx.StockFisicoLote es una variable que todavia aqui guarda el valor inicial 
        '//del stock fisico del lote)
        If (Not SegundaUnidad AndAlso data.Context.StockFisicoLote = 0 And data.Cantidad = 0) OrElse _
               (SegundaUnidad AndAlso (data.Context.StockFisicoLote = 0 And data.Cantidad = 0) AndAlso (data.Context.StockFisicoLote2 = 0 AndAlso data.Cantidad2 = 0)) Then
            data.Context.GenerarMovimiento = False
        End If

        data.Context.StockFisico += ((data.Cantidad + data.Context.QIntermedia) - data.Context.StockFisicoLote)
        data.Context.StockFisicoLote = data.Cantidad + data.Context.QIntermedia
        If SegundaUnidad Then
            data.Context.StockFisico2 += ((data.Cantidad2 + data.Context.QIntermedia2) - data.Context.StockFisicoLote2)
            data.Context.StockFisicoLote2 = data.Cantidad2 + data.Context.QIntermedia2
        End If

        Dim Rules() As Regla = New ProcesoStocks().EstablecerReglas(Regla.StockLoteNegativo)
        Dim datValCtx As New DataValidarContexto(data.Articulo, data.Almacen, data.Context, Rules)
        ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)
    End Sub

    <Task()> Public Shared Sub SetCtxArticuloAlmacenNSerieInventario(ByVal data As StockData, ByVal services As ServiceProvider)
        Dim stk As New ProcesoStocks
        Dim Rules() As Regla : Dim datValCtx As New DataValidarContexto
        Dim Serie As DataTable = data.Context.SerieBBDD
        If Serie.Rows.Count = 0 Then
            Dim newrow As DataRow = Serie.NewRow
            newrow("IDArticulo") = data.Articulo
            newrow("NSerie") = data.NSerie
            newrow("IDAlmacen") = data.Almacen
            newrow("IDEstadoActivo") = data.EstadoNSerie
            newrow("IDOperario") = data.Operario
            If newrow.Table.Columns.Contains("Ubicacion") Then newrow("Ubicacion") = data.Ubicacion
            newrow("MarcaAuto") = AdminData.GetAutoNumeric()
            Serie.Rows.Add(newrow)

            Rules = stk.EstablecerReglas(Regla.SerieUnica)
            datValCtx = New DataValidarContexto(data.Articulo, data.Almacen, data.Context, Rules)
            ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)

            data.Context.StockFisico += 1
            Dim AppParamsStock As ParametroStocks = services.GetService(Of ParametroStocks)()
            If AppParamsStock.GestionNumeroSerieConActivos Then
                Dim dataAct As New DataActualizarActivo(data, newrow)
                ProcessServer.ExecuteTask(Of DataActualizarActivo)(AddressOf ActualizarActivo, dataAct, services)
            End If
        Else
            If Serie.Rows(0)("IDEstadoActivo") <> data.EstadoNSerie Then
                Serie.Rows(0)("IDEstadoActivo") = data.EstadoNSerie
                If data.Context.PropiedadesEstado.Baja AndAlso data.Context.PropiedadesEstadoBBDD.Disponible Then
                    Serie.Rows(0)("IDAlmacen") = DBNull.Value
                    If Serie.Columns.Contains("Ubicacion") Then Serie.Rows(0)("Ubicacion") = DBNull.Value
                    data.Context.StockFisico -= 1
                ElseIf data.Context.PropiedadesEstado.Disponible AndAlso data.Context.PropiedadesEstadoBBDD.Baja Then
                    Serie.Rows(0)("IDAlmacen") = data.Almacen
                    If Serie.Columns.Contains("Ubicacion") Then Serie.Rows(0)("Ubicacion") = data.Ubicacion
                    data.Context.StockFisico += 1
                ElseIf data.Context.PropiedadesEstado.Disponible AndAlso data.Context.PropiedadesEstadoBBDD.Disponible Then
                    data.Context.StockFisico += 1
                    data.Context.GenerarMovimiento = False
                ElseIf data.Context.PropiedadesEstado.Baja AndAlso data.Context.PropiedadesEstadoBBDD.Baja Then
                    data.Context.StockFisico -= 1
                    data.Context.GenerarMovimiento = False
                End If
                Dim AppParamsStock As ParametroStocks = services.GetService(Of ParametroStocks)()
                If AppParamsStock.GestionNumeroSerieConActivos Then
                    Dim dataAct As New DataActualizarActivo(data, Serie.Rows(0))
                    ProcessServer.ExecuteTask(Of DataActualizarActivo)(AddressOf ActualizarActivo, dataAct, services)
                    data.Context.Estado = EstadoStock.Actualizado
                End If
            End If
            If Serie.Rows(0)("IDOperario") & String.Empty <> data.Operario Then
                Serie.Rows(0)("IDOperario") = data.Operario
            End If

            data.Lote = data.NSerie
        End If

        Rules = stk.EstablecerReglas(Regla.StockSerieNegativo)
        datValCtx = New DataValidarContexto(data.Articulo, data.Almacen, data.Context, Rules)
        ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)
    End Sub

    <Task()> Public Shared Sub SetCtxStockFisicoAlmacenInventario(ByVal data As StockData, ByVal services As ServiceProvider)
        'Dim QIntermedia As Double = ProcessServer.ExecuteTask(Of StockData, Double)(AddressOf ContabilizarMovimientosIntermedios, data, services)
        ProcessServer.ExecuteTask(Of StockData)(AddressOf ContabilizarMovimientosIntermedios, data, services)
        data.Context.StockFisico = data.Cantidad + data.Context.QIntermedia
        If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Articulo, services) Then
            data.Context.StockFisico2 = data.Cantidad2 + data.Context.QIntermedia2
        End If

        Dim Rules() As Regla = New ProcesoStocks().EstablecerReglas(Regla.StockNegativo)
        Dim datValCtx As New DataValidarContexto(data.Articulo, data.Almacen, data.Context, Rules)
        ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)
    End Sub

    <Task()> Public Shared Sub CorreccionInventario(ByVal data As StockData, ByVal services As ServiceProvider)
        '//Correccion de inventario
        Dim f As New Filter
        f.Add(New StringFilterItem("IDArticulo", data.Articulo))
        f.Add(New StringFilterItem("IDAlmacen", data.Almacen))
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Articulo)
        If ArtInfo.GestionStockPorLotes Then
            f.Add(New StringFilterItem("Lote", data.Lote))
            f.Add(New StringFilterItem("Ubicacion", data.Ubicacion))
        ElseIf ArtInfo.GestionPorNumeroSerie Then
            f.Add(New StringFilterItem("Lote", data.NSerie))
            f.Add(New IsNullFilterItem("Ubicacion"))
        End If
        f.Add(New NumberFilterItem("IDTipoMovimiento", enumTipoMovimiento.tmInventario))
        f.Add(New DateFilterItem("FechaDocumento", data.FechaDocumento))
        Dim ultimoInventario As DataTable = New BE.DataEngine().Filter(cnEntidad, f)
        If ultimoInventario.Rows.Count > 0 Then

            data.Context.NumeroMovimiento = ultimoInventario.Rows(0)("IDMovimiento")
            '//Hacer un copia del movimiento original y guardarla como movimiento de correccion
            Dim correccion As DataRow = data.Context.Movimientos.NewRow
            correccion.ItemArray = ultimoInventario.Rows(0).ItemArray
            correccion("IDLineaMovimiento") = AdminData.GetAutoNumeric
            correccion("IDTipoMovimiento") = enumTipoMovimiento.tmCorreccion
            correccion("FechaMovimiento") = Today
            correccion("Texto") = DataRowState.Modified.ToString
            data.Context.Movimientos.Rows.Add(correccion)

            '//Volcar los nuevos valores sobre el movimiento original
            If ArtInfo.GestionPorNumeroSerie Then
                If data.Context.PropiedadesEstado.Disponible Then
                    ultimoInventario.Rows(0)("Cantidad") = data.Cantidad
                Else
                    ultimoInventario.Rows(0)("Cantidad") = 0
                End If
            Else
                ultimoInventario.Rows(0)("Cantidad") = data.Cantidad
            End If

            Dim SegundaUnidad As Boolean = (ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Articulo, services) AndAlso Length(data.Cantidad2) > 0)
            '//Calculamos el Acumulado y el Acumulado2, para todos los movimientos, incluido el que estamos tratatando
            Dim datAcumInv As New DataAcumulado(data, ultimoInventario.Rows(0)("IDLineaMovimiento"), SegundaUnidad)
            Dim datValorAcum As DataValorAcumulado = ProcessServer.ExecuteTask(Of DataAcumulado, DataValorAcumulado)(AddressOf ProcesoStocks.AcumuladoInventario, datAcumInv, services)
            ultimoInventario.Rows(0)("Acumulado") = datValorAcum.Valor
            If SegundaUnidad Then
                ultimoInventario.Rows(0)("Cantidad2") = data.Cantidad2
                If Not datValorAcum.Valor2 Is Nothing Then ultimoInventario.Rows(0)("Acumulado2") = datValorAcum.Valor2
            End If

            ultimoInventario.Rows(0)("FechaDocumento") = data.FechaDocumento
            ultimoInventario.Rows(0)("PrecioA") = data.PrecioA
            ultimoInventario.Rows(0)("PrecioB") = data.PrecioB
            If data.PrecioA = 0 Or data.PrecioB = 0 Then
                Dim dataPrecioEnt As New DataPrecioMovimiento(data.Articulo, data.Almacen, data.FechaDocumento, data.Cantidad, enumtpmTipoMovimiento.tpmInventario)
                Dim precios As Hashtable = ProcessServer.ExecuteTask(Of DataPrecioMovimiento, Hashtable)(AddressOf PrecioMovimiento, dataPrecioEnt, services)
                If Not precios Is Nothing Then
                    ultimoInventario.Rows(0)("PrecioA") = precios("PrecioA")
                    ultimoInventario.Rows(0)("PrecioB") = precios("PrecioB")
                End If
            End If
            data.Context.Movimientos.ImportRow(ultimoInventario.Rows(0))
            data.Context.EsCorreccion = True
            ProcessServer.ExecuteTask(Of DataMovimiento)(AddressOf SetMessageMovimientoActualizado, New DataMovimiento(ultimoInventario.Rows(0), data), services)
        End If
    End Sub

    ' <Task()> Public Shared Function ContabilizarMovimientosIntermedios(ByVal data As StockData, ByVal services As ServiceProvider) As Double
    <Task()> Public Shared Sub ContabilizarMovimientosIntermedios(ByVal data As StockData, ByVal services As ServiceProvider)
        Dim q, q2 As Double
        Dim qresto, qresto2 As Double

        Dim f As New Filter
        f.Add(New StringFilterItem("IDArticulo", data.Articulo))
        f.Add(New StringFilterItem("IDAlmacen", data.Almacen))
        Dim AppParamsStock As ParametroStocks = services.GetService(Of ParametroStocks)()

        Select Case AppParamsStock.TipoInventario
            Case TipoInventario.PrimerMovimiento
                f.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThanOrEqual, data.FechaDocumento))
            Case TipoInventario.UltimoMovimiento
                f.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThan, data.FechaDocumento))
        End Select

        Dim movimientos As DataTable = New BE.DataEngine().Filter(cnEntidad, f, , "FechaDocumento")
        If Not IsNothing(movimientos) AndAlso movimientos.Rows.Count Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Articulo)

            For Each m As DataRow In movimientos.Rows
                If m("IDTipoMovimiento") <> enumTipoMovimiento.tmInventario And m("IDTipoMovimiento") <> enumTipoMovimiento.tmCorreccion Then
                    If ArtInfo.GestionStockPorLotes Then
                        If Length(m("Lote")) > 0 And Length(m("Ubicacion")) > 0 Then
                            If m("Lote") = data.Lote And m("Ubicacion") = data.Ubicacion Then
                                q = q + m("Cantidad")
                                q2 = q2 + Nz(m("Cantidad2"), 0)
                            Else
                                qresto = qresto + m("Cantidad")
                                qresto2 = qresto2 + Nz(m("Cantidad2"), 0)
                            End If
                        Else
                            Throw New Exception(ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(43, data.FechaDocumento, Quoted(data.Articulo), Quoted(data.Almacen)), services))
                        End If
                    ElseIf ArtInfo.GestionPorNumeroSerie Then
                        'pend explicar
                    Else
                        q = q + m("Cantidad")
                        q2 = q2 + Nz(m("Cantidad2"), 0)
                    End If
                End If
            Next
        End If

        data.Context.QIntermedia = q
        data.Context.QIntermedia2 = q2
        data.Context.QIntermediaRestoLotes = qresto
        data.Context.QIntermediaRestoLotes2 = qresto2

        ' Return q
    End Sub

    <Task()> Public Shared Sub InicializarInventariados(ByVal data As Object, ByVal services As ServiceProvider)
        Dim dt As DataTable = New ArticuloAlmacen().Filter()
        If Not IsNothing(dt) AndAlso dt.Rows.Count Then
            For Each dr As DataRow In dt.Rows
                dr("Inventariado") = False
            Next
            BusinessHelper.UpdateTable(dt)
        End If
    End Sub

#End Region

#Region " Correcciones "

    <Serializable()> _
   Public Class DataActualizarMovimiento
        Public IDLineaMovimiento As Integer
        Public Sinc As Boolean
        Public Cantidad As Double?
        Public Cantidad2 As Double?
        Public FechaDocumento As Date?
        Public PrecioA As Double?
        Public PrecioB As Double?
        Public TipoActualizacion As enumTipoActualizacion
        Public Documento As String

        Public CorrectContext As StockCorrectContext

        Public Sub New(ByVal TipoActualizacion As enumTipoActualizacion, ByVal IDLineaMovimiento As Integer, Optional ByVal Sinc As Boolean = True)
            Me.IDLineaMovimiento = IDLineaMovimiento
            Me.Sinc = Sinc
            Me.TipoActualizacion = TipoActualizacion
            If Me.TipoActualizacion = enumTipoActualizacion.Eliminar Then
                Me.Cantidad = 0
                Me.CorrectContext = New StockCorrectContext(False, False, False, False, True)
            End If
        End Sub

        Public Sub New(ByVal TipoActualizacion As enumTipoActualizacion, ByVal IDLineaMovimiento As Integer, ByVal Cantidad As Double, Optional ByVal Sinc As Boolean = True)
            Me.New(TipoActualizacion, IDLineaMovimiento, Sinc)
            Me.Cantidad = Cantidad
            Me.CorrectContext = New StockCorrectContext(True, False, False, False, False)
        End Sub

        Public Sub New(ByVal TipoActualizacion As enumTipoActualizacion, ByVal IDLineaMovimiento As Integer, ByVal Cantidad As Double, ByVal NDocumento As String, Optional ByVal Sinc As Boolean = True)
            Me.New(TipoActualizacion, IDLineaMovimiento, Cantidad, Sinc)
            Me.Documento = NDocumento
            Me.CorrectContext = New StockCorrectContext(True, False, False, True, False)
        End Sub

        Public Sub New(ByVal TipoActualizacion As enumTipoActualizacion, ByVal IDLineaMovimiento As Integer, ByVal FechaDocumento As Date, Optional ByVal Sinc As Boolean = True)
            Me.New(TipoActualizacion, IDLineaMovimiento, Sinc)
            Me.FechaDocumento = FechaDocumento
            Me.CorrectContext = New StockCorrectContext(False, False, True, False, False)
        End Sub

        Public Sub New(ByVal TipoActualizacion As enumTipoActualizacion, ByVal IDLineaMovimiento As Integer, ByVal PrecioA As Double, ByVal PrecioB As Double, Optional ByVal Sinc As Boolean = True)
            Me.New(TipoActualizacion, IDLineaMovimiento, Sinc)
            Me.PrecioA = PrecioA
            Me.PrecioB = PrecioB
            Me.CorrectContext = New StockCorrectContext(False, True, False, False, False)
        End Sub

        Public Sub New(ByVal TipoActualizacion As enumTipoActualizacion, ByVal IDLineaMovimiento As Integer, ByVal PrecioA As Double, ByVal PrecioB As Double, ByVal NDocumento As String, Optional ByVal Sinc As Boolean = True)
            Me.New(TipoActualizacion, IDLineaMovimiento, PrecioA, PrecioB, Sinc)
            Me.Documento = NDocumento
            Me.CorrectContext = New StockCorrectContext(False, True, False, True, False)
        End Sub

        Public Sub New(ByVal TipoActualizacion As enumTipoActualizacion, ByVal IDLineaMovimiento As Integer, ByVal Cantidad As Double, ByVal PrecioA As Double, ByVal PrecioB As Double, Optional ByVal Sinc As Boolean = True)
            Me.New(TipoActualizacion, IDLineaMovimiento, PrecioA, PrecioB, Sinc)
            Me.Cantidad = Cantidad
            Me.CorrectContext = New StockCorrectContext(True, True, False, False, False)
        End Sub

        Public Sub New(ByVal TipoActualizacion As enumTipoActualizacion, ByVal IDLineaMovimiento As Integer, ByVal Cantidad As Double, ByVal PrecioA As Double, ByVal PrecioB As Double, ByVal NDocumento As String, Optional ByVal Sinc As Boolean = True)
            Me.New(TipoActualizacion, IDLineaMovimiento, Cantidad, PrecioA, PrecioB, Sinc)
            Me.Documento = NDocumento
            Me.CorrectContext = New StockCorrectContext(True, True, False, True, False)
        End Sub

        Public Sub New(ByVal TipoActualizacion As enumTipoActualizacion, ByVal IDLineaMovimiento As Integer, ByVal Cantidad As Double, ByVal FechaDocumento As Date, ByVal PrecioA As Double, ByVal PrecioB As Double, Optional ByVal Sinc As Boolean = True)
            Me.New(TipoActualizacion, IDLineaMovimiento, Cantidad, PrecioA, PrecioB, Sinc)
            Me.FechaDocumento = FechaDocumento
            Me.CorrectContext = New StockCorrectContext(True, True, True, False, False)
        End Sub

        Public Sub New(ByVal TipoActualizacion As enumTipoActualizacion, ByVal IDLineaMovimiento As Integer, ByVal FechaDocumento As Date, ByVal NDocumento As String, Optional ByVal Sinc As Boolean = True)
            Me.New(TipoActualizacion, IDLineaMovimiento, FechaDocumento, Sinc)
            Me.FechaDocumento = FechaDocumento
            Me.CorrectContext = New StockCorrectContext(False, False, True, True, False)
        End Sub
    End Class

    <Task()> Public Shared Function ActualizarMovimiento(ByVal data As DataActualizarMovimiento, ByVal services As ServiceProvider) As StockUpdateData
        Dim movimiento As DataRow = ProcessServer.ExecuteTask(Of Integer, DataRow)(AddressOf GetMovimiento, data.IDLineaMovimiento, services)
        Dim returnData As StockUpdateData
        If movimiento Is Nothing Then
            returnData = New StockUpdateData
            returnData.Estado = EstadoStock.NoActualizado
            returnData.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(24), services)
        Else
            Dim stkData As StockData = ProcessServer.ExecuteTask(Of DataRow, StockData)(AddressOf ConvertirMovimientoAStockData, movimiento, services)
            Select Case data.TipoActualizacion
                Case enumTipoActualizacion.Eliminar
                    '//Al eliminar un movimiento, no se borra, pasa a ser un movimiento de corrección con cantidad 0.
                    stkData.Cantidad = 0
                    stkData.Cantidad2 = 0
                    stkData.ContextCorrect = data.CorrectContext
                Case enumTipoActualizacion.Corregir
                    If Not data.Cantidad Is Nothing Then stkData.Cantidad = data.Cantidad
                    If Not data.Cantidad2 Is Nothing Then stkData.Cantidad2 = data.Cantidad2
                    If Not data.PrecioA Is Nothing Then
                        stkData.PrecioA = data.PrecioA
                        stkData.PrecioB = data.PrecioB
                    End If
                    If Not data.FechaDocumento Is Nothing Then stkData.FechaDocumento = data.FechaDocumento
                    stkData.ContextCorrect = data.CorrectContext
            End Select
            Dim datCorrectMovto As New DataMovimientoSinc(movimiento, stkData, data.Sinc)
            returnData = ProcessServer.ExecuteTask(Of DataMovimientoSinc, StockUpdateData)(AddressOf CorregirMovimiento, datCorrectMovto, services)
        End If
        Return returnData
    End Function

    <Task()> Public Shared Function CorregirMovimiento(ByVal data As DataMovimientoSinc, ByVal services As ServiceProvider) As StockUpdateData
        If Not data Is Nothing Then
            Dim stkDataOriginal As StockData = ProcessServer.ExecuteTask(Of DataRow, StockData)(AddressOf ConvertirMovimientoAStockData, data.Movimiento, services)
            Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataMovimiento, StockUpdateData)(AddressOf CorregirMovimientoTx, data, services)
            If Not updateData Is Nothing Then
                If updateData.Estado = EstadoStock.Actualizado Then
                    AdminData.BeginTx()
                    Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
                    If data.Sinc AndAlso AppParams.GestionBodegas Then
                        If data.stkData.ContextCorrect.EsBorrado Then
                            Dim datSinc As New DataIntegracionConBodega(New DataNumeroMovimientoSinc(data.Movimiento("IDMovimiento"), data.stkData, data.Sinc), enumTipoSincronizacion.EliminarMovimiento, stkDataOriginal, , updateData, , data.Movimiento("IDLineaMovimiento"))
                            Dim updateDataAux As StockUpdateData = ProcessServer.ExecuteTask(Of DataIntegracionConBodega, StockUpdateData)(AddressOf IntegracionConBodega, datSinc, services)
                            If Not updateDataAux Is Nothing Then
                                updateData = updateDataAux
                            End If
                        Else
                            Dim datSinc As New DataIntegracionConBodega(New DataNumeroMovimientoSinc(data.Movimiento("IDMovimiento"), data.stkData, data.Sinc), enumTipoSincronizacion.Correccion, stkDataOriginal, , updateData)
                            Dim updateDataAux As StockUpdateData = ProcessServer.ExecuteTask(Of DataIntegracionConBodega, StockUpdateData)(AddressOf IntegracionConBodega, datSinc, services)
                            If Not updateDataAux Is Nothing Then
                                updateData = updateDataAux
                            End If
                        End If
                    End If
                    ProcessServer.ExecuteTask(Of StockUpdateData)(AddressOf Actualizar, updateData, services)
                    If updateData.Estado = EstadoStock.NoActualizado Then
                        ApplicationService.GenerateError(updateData.Detalle)
                    End If
                Else
                    If Nz(updateData.StockData.Context.FechaUltimoCierre, cnMinDate) <> cnMinDate Then
                        If updateData.StockData.Context.FechaDocumento <= updateData.StockData.Context.FechaUltimoCierre Then
                            ApplicationService.GenerateError(updateData.Detalle)
                        End If
                    End If
                End If
                Return updateData
            End If
        End If
    End Function

    <Task()> Public Shared Function CorregirMovimientoTx(ByVal data As DataMovimiento, ByVal services As ServiceProvider) As StockUpdateData
        ProcessServer.ExecuteTask(Of DataMovimiento)(AddressOf CrearContextoCorreccion, data, services)
        If Not data.stkData.Context.Cancel Then
            Dim stk As New ProcesoStocks
            Dim Rules() As Regla = stk.EstablecerReglas(Regla.LoteBloqueado)
            If ProcessServer.ExecuteTask(Of StockData, Boolean)(AddressOf ValidarFechaUltimoCierre, data.stkData, services) Then
                ReDim Preserve Rules(Rules.Length)
                Rules(Rules.Length - 1) = Regla.FechaUltimoCierre
            End If
            Dim datValCtx As New DataValidarContexto(data.stkData.Articulo, data.stkData.Almacen, data.stkData.Context, Rules)
            ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)

            If data.stkData.Context.Cancel Then
                Return ProcessServer.ExecuteTask(Of StockData, StockUpdateData)(AddressOf GetStockUpdateData, data.stkData, services)
            End If

            '//Los movimientos de ajuste no se pueden corregir.
            '//Los movimientos de inventario se corrigen en la propia funcion de inventarios.
            If data.Movimiento("IDTipoMovimiento") <> enumTipoMovimiento.tmCorreccion AndAlso _
               data.Movimiento("IDTipoMovimiento") <> enumTipoMovimiento.tmEntAjuste AndAlso _
               data.Movimiento("IDTipoMovimiento") <> enumTipoMovimiento.tmSalAjuste AndAlso _
               data.Movimiento("IDTipoMovimiento") <> enumTipoMovimiento.tmInventario Then

                Dim AppParamsStock As ParametroStocks = services.GetService(Of ParametroStocks)()
                If (AppParamsStock.TipoInventario = TipoInventario.PrimerMovimiento And data.stkData.FechaDocumento >= data.stkData.Context.FechaUltimoInventario) _
                Or (AppParamsStock.TipoInventario = TipoInventario.UltimoMovimiento And data.stkData.FechaDocumento > data.stkData.Context.FechaUltimoInventario) Then
                    Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                    Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.stkData.Articulo)
                    If ArtInfo.GestionStockPorLotes Then
                        ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxArticuloAlmacenLoteCorreccion, data.stkData, services)
                    ElseIf ArtInfo.GestionPorNumeroSerie Then
                        ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxArticuloAlmacenNSerieCorreccion, data.stkData, services)
                    Else
                        ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxStockFisicoAlmacenCorreccion, data.stkData, services)
                    End If

                    If data.stkData.Context.Cancel Then
                        Return ProcessServer.ExecuteTask(Of StockData, StockUpdateData)(AddressOf GetStockUpdateData, data.stkData, services)
                    End If
                End If

                If Not data.stkData.ContextCorrect.EsBorrado Then
                    '//Hacer un copia del movimiento original y guardarla como movimiento de correccion
                    ProcessServer.ExecuteTask(Of DataMovimiento)(AddressOf ConvertirMovimientoEnCorreccion, data, services)

                    '//Volcar los nuevos valores sobre el movimiento original
                    ProcessServer.ExecuteTask(Of DataMovimiento)(AddressOf NuevoMovimientoTrasCorreccion, data, services)
                Else
                    ProcessServer.ExecuteTask(Of DataMovimiento)(AddressOf ConvertirMovimientoEnCorreccionPorBorrado, data, services)
                End If

                ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxArticuloAlmacen, data.stkData, services)
                '''''Dim ArticuloAlmacen As DataTable = data.stkData.Context.ArticuloAlmacen
                '''''ArticuloAlmacen.Rows(0)("StockFisico") = data.stkData.Context.StockFisico
                '''''If data.stkData.Context.FechaUltimoMovimiento = Date.MinValue Then
                '''''    ArticuloAlmacen.Rows(0)("FechaUltimoMovimiento") = DBNull.Value
                '''''Else
                '''''    ArticuloAlmacen.Rows(0)("FechaUltimoMovimiento") = data.stkData.Context.FechaUltimoMovimiento
                '''''End If

                ProcessServer.ExecuteTask(Of DataMovimiento)(AddressOf SetMessageMovimientoActualizado, data, services)
            End If
        End If

        '// Retornamos lo que ha ocurrido en el proceso de corrección, para el artículo y almacén indicados.
        Return ProcessServer.ExecuteTask(Of StockData, StockUpdateData)(AddressOf GetStockUpdateData, data.stkData, services)
    End Function

    <Task()> Public Shared Sub CrearContextoCorreccion(ByVal data As DataMovimiento, ByVal services As ServiceProvider)
        Dim dataCtx As New DataCrearContexto(data.Movimiento("IDMovimiento"), data.stkData, True)
        ProcessServer.ExecuteTask(Of DataCrearContexto)(AddressOf CrearContexto, dataCtx, services)

        If Not data.stkData.Context.Cancel Then
            If data.stkData.Context.TipoMovimiento <> data.Movimiento("IDTipoMovimiento") Then
                data.stkData.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(30), services)
                data.stkData.Context.Cancel = True
            Else
                Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.stkData.Articulo, services)
                data.stkData.Context.CantidadOriginal = data.Movimiento("Cantidad")
                If SegundaUnidad AndAlso Length(data.Movimiento("Cantidad2")) > 0 Then
                    data.stkData.Context.CantidadOriginal2 = CDbl(data.Movimiento("Cantidad2"))
                End If
                data.stkData.Context.PrecioAOriginal = data.Movimiento("PrecioA")
                data.stkData.Context.PrecioBOriginal = data.Movimiento("PrecioB")
                data.stkData.Context.FechaDocumentoOriginal = data.Movimiento("FechaDocumento")
                data.stkData.Context.DocumentoOriginal = data.Movimiento("Documento") & String.Empty

                If Not data.stkData.ContextCorrect.EsBorrado Then
                    '//correccion introducida por los movimientos de devolucion
                    If Not data.stkData.ContextCorrect.CorreccionEnCantidad Then
                        data.stkData.Context.CantidadConSigno = data.stkData.Cantidad
                        If SegundaUnidad Then
                            data.stkData.Context.CantidadConSigno2 = data.stkData.Cantidad2
                        End If
                    End If
                    '//

                    If data.stkData.Context.CantidadConSigno <> data.stkData.Context.CantidadOriginal Then
                        data.stkData.ContextCorrect.CorreccionEnCantidad = True
                    End If
                    If SegundaUnidad AndAlso data.stkData.Context.CantidadConSigno2 <> data.stkData.Context.CantidadOriginal2 Then
                        data.stkData.ContextCorrect.CorreccionEnCantidad2 = True
                    End If
                    If data.stkData.Context.PrecioA <> data.stkData.Context.PrecioAOriginal Then
                        data.stkData.ContextCorrect.CorreccionEnPrecio = True
                    End If
                    If data.stkData.FechaDocumento <> data.stkData.Context.FechaDocumentoOriginal Then
                        data.stkData.ContextCorrect.CorreccionEnFecha = True
                    End If
                    If data.stkData.Documento <> data.stkData.Context.DocumentoOriginal Then
                        data.stkData.ContextCorrect.CorreccionEnDocumento = True
                    End If

                    If Not data.stkData.ContextCorrect.CorreccionEnCantidad And Not data.stkData.ContextCorrect.CorreccionEnPrecio And Not data.stkData.ContextCorrect.CorreccionEnFecha And Not data.stkData.ContextCorrect.CorreccionEnDocumento AndAlso Not data.stkData.ContextCorrect.CorreccionEnCantidad2 Then
                        data.stkData.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(45), services)
                        data.stkData.Context.Cancel = True
                    End If
                End If

                If Not data.stkData.Context.Cancel And data.stkData.ContextCorrect.CorreccionEnFecha Then
                    '//Todas las funciones de corregir movimientos pasan por aqui.
                    ProcessServer.ExecuteTask(Of StockData)(AddressOf ValidarCorreccionEnFecha, data.stkData, services)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarCorreccionEnFecha(ByVal data As StockData, ByVal services As ServiceProvider)
        '//Respecto a un inventario, la FechaDocumento FINAL del movimiento tiene que estar al "mismo lado" 
        '//del inventario que la FechaDocumento INICIAL, es decir, un movimiento no se puede mover en fecha 
        '//pasando por encima de un movimiento de inventario y todo esto teniendo en cuenta el parametro 
        '//que establece si el inventario es el primer o ultimo movimiento del dia.

        Dim fecha1, fecha2 As Date
        If data.Context.FechaDocumentoOriginal < data.FechaDocumento Then
            fecha1 = data.Context.FechaDocumentoOriginal
            fecha2 = data.FechaDocumento
        ElseIf data.Context.FechaDocumentoOriginal > data.FechaDocumento Then
            fecha1 = data.FechaDocumento
            fecha2 = data.Context.FechaDocumentoOriginal
        End If

        Dim f As New Filter
        f.Add(New StringFilterItem("IDArticulo", data.Articulo))
        f.Add(New StringFilterItem("IDAlmacen", data.Almacen))
        f.Add(New NumberFilterItem("IDTipoMovimiento", enumTipoMovimiento.tmInventario))
        f.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThanOrEqual, fecha1))
        f.Add(New DateFilterItem("FechaDocumento", FilterOperator.LessThanOrEqual, fecha2))
        Dim inventario As DataTable = New BE.DataEngine().Filter(cnEntidad, f)
        If inventario.Rows.Count > 1 Then
            data.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(44), services)
            data.Context.Cancel = True
        ElseIf inventario.Rows.Count = 1 Then
            Dim fechaInventario As Date = inventario.Rows(0)("FechaDocumento")
            Dim AppParamsStock As ParametroStocks = services.GetService(Of ParametroStocks)()
            Select Case AppParamsStock.TipoInventario
                Case TipoInventario.PrimerMovimiento
                    If data.FechaDocumento < data.Context.FechaDocumentoOriginal Then
                        '//Correcciones en fecha "hacia atras", cuando hay un inventario 
                        '//y es el primero movimiento del dia, el unico caso permitido es
                        '//que el movimiento caiga en el propio dia del inventario.
                        If data.FechaDocumento <> fechaInventario Then
                            data.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(44), services)
                            data.Context.Cancel = True
                        End If
                    ElseIf data.Context.FechaDocumentoOriginal < data.FechaDocumento Then
                        '//Correcciones en fecha "hacia delante", cuando hay un inventario 
                        '//y es el primero movimiento del dia, el unico caso permitido es
                        '//que el movimiento original este en el propio dia del inventario.
                        If data.Context.FechaDocumentoOriginal <> fechaInventario Then
                            data.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(44), services)
                            data.Context.Cancel = True
                        End If
                    End If
                Case TipoInventario.UltimoMovimiento
                    If data.FechaDocumento < data.Context.FechaDocumentoOriginal Then
                        '//Correcciones en fecha "hacia atras", cuando hay un inventario 
                        '//y es el ultimo movimiento del dia, el unico caso permitido es
                        '//que el movimiento original este en el propio dia del inventario.
                        If data.Context.FechaDocumentoOriginal <> fechaInventario Then
                            data.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(44), services)
                            data.Context.Cancel = True
                        End If
                    ElseIf data.Context.FechaDocumentoOriginal < data.FechaDocumento Then
                        '//Correcciones en fecha "hacia delante", cuando hay un inventario 
                        '//y es el ultimo movimiento del dia, el unico caso permitido es
                        '//que el movimiento caiga en el propio dia del inventario.
                        If data.FechaDocumento <> fechaInventario Then
                            data.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(44), services)
                            data.Context.Cancel = True
                        End If
                    End If
            End Select
        End If
    End Sub

    <Task()> Public Shared Sub SetCtxArticuloAlmacenLoteCorreccion(ByVal data As StockData, ByVal services As ServiceProvider)
        Dim CantidadOriginal As Double = data.Context.CantidadOriginal
        Dim CantidadFinal As Double = data.Context.CantidadConSigno

        Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Articulo, services)
        Dim CantidadOriginal2, CantidadFinal2 As Double
        If SegundaUnidad Then
            CantidadOriginal2 = data.Context.CantidadOriginal2
            CantidadFinal2 = data.Context.CantidadConSigno2
        End If

        If data.Context.LoteBBDD.Rows.Count = 0 Then
            data.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(20), services)
            data.Context.Cancel = True
        Else
            data.Context.LoteBBDD.Rows(0)("StockFisico") = data.Context.LoteBBDD.Rows(0)("StockFisico") - CantidadOriginal + CantidadFinal
            data.Context.StockFisicoLote = data.Context.LoteBBDD.Rows(0)("StockFisico")
            If SegundaUnidad Then
                data.Context.LoteBBDD.Rows(0)("StockFisico2") = data.Context.LoteBBDD.Rows(0)("StockFisico2") - CantidadOriginal2 + CantidadFinal2
                data.Context.StockFisicoLote2 = CDbl(Nz(data.Context.LoteBBDD.Rows(0)("StockFisico2"), 0))
            End If
        End If
        If Not data.Context.Cancel Then
            data.Context.StockFisico = data.Context.StockFisico - CantidadOriginal + CantidadFinal
            If SegundaUnidad Then
                data.Context.StockFisico2 = data.Context.StockFisico2 - CantidadOriginal2 + CantidadFinal2
            End If
            Dim Rules() As Regla = New ProcesoStocks().EstablecerReglas(Regla.StockLoteNegativo)
            Dim datValCtx As New DataValidarContexto(data.Articulo, data.Almacen, data.Context, Rules)
            ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)
        End If
    End Sub

    <Task()> Public Shared Sub SetCtxArticuloAlmacenNSerieCorreccion(ByVal data As StockData, ByVal services As ServiceProvider)
        Dim Serie As DataTable = data.Context.SerieBBDD
        Dim CantidadOriginal As Double = data.Context.CantidadOriginal
        Dim CantidadFinal As Double = data.Context.CantidadConSigno

        If Not data.ContextCorrect.EsBorrado Then
            If data.ContextCorrect.CorreccionEnCantidad Then
                data.Context.Cancel = True
                data.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(39), services)
            End If
        Else
            If Serie.Rows.Count > 0 Then
                Serie.Rows(0)("IDAlmacen") = data.Almacen
                If Serie.Columns.Contains("Ubicacion") Then Serie.Rows(0)("Ubicacion") = data.Ubicacion
                If Length(data.EstadoNSerieAnterior) = 0 Then
                    If data.Context.ClaseMovimiento = enumtpmTipoMovimiento.tpmInput Then
                        If data.Context.CantidadOriginal < 0 Then
                            Serie.Rows(0)("IDEstadoActivo") = NegocioGeneral.ESTADOACTIVO_DISPONIBLE
                        Else
                            Serie.Rows(0)("IDEstadoActivo") = NegocioGeneral.ESTADOACTIVO_BAJA
                            Serie.Rows(0)("IDAlmacen") = DBNull.Value
                            If Serie.Columns.Contains("Ubicacion") Then Serie.Rows(0)("Ubicacion") = DBNull.Value
                        End If
                    End If
                    If data.Context.ClaseMovimiento = enumtpmTipoMovimiento.tpmOutput Then
                        If data.Context.CantidadOriginal > 0 Then
                            Serie.Rows(0)("IDEstadoActivo") = NegocioGeneral.ESTADOACTIVO_BAJA
                            Serie.Rows(0)("IDAlmacen") = DBNull.Value
                            If Serie.Columns.Contains("Ubicacion") Then Serie.Rows(0)("Ubicacion") = DBNull.Value
                        Else
                            Serie.Rows(0)("IDEstadoActivo") = NegocioGeneral.ESTADOACTIVO_DISPONIBLE
                        End If
                    End If
                Else
                    Serie.Rows(0)("IDEstadoActivo") = data.EstadoNSerieAnterior
                End If
                data.Context.StockFisico += -CantidadOriginal

                Dim AppParamsStock As ParametroStocks = services.GetService(Of ParametroStocks)()
                If AppParamsStock.GestionNumeroSerieConActivos Then
                    Dim dataAct As New DataActualizarActivo(data, Serie.Rows(0))
                    ProcessServer.ExecuteTask(Of DataActualizarActivo)(AddressOf ActualizarActivo, dataAct, services)
                    data.Context.Estado = EstadoStock.Actualizado
                End If

                Dim Rules() As Regla = New ProcesoStocks().EstablecerReglas(Regla.StockSerieNegativo)
                Dim datValCtx As New DataValidarContexto(data.Articulo, data.Almacen, data.Context, Rules)
                ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)
            Else
                data.Context.Cancel = True
                data.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(36), services)
            End If
        End If

    End Sub

    <Task()> Public Shared Sub SetCtxStockFisicoAlmacenCorreccion(ByVal data As StockData, ByVal services As ServiceProvider)
        Dim CantidadOriginal As Double = data.Context.CantidadOriginal
        Dim CantidadFinal As Double = data.Context.CantidadConSigno

        data.Context.StockFisico = data.Context.StockFisico - CantidadOriginal + CantidadFinal

        If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Articulo, services) Then
            Dim CantidadOriginal2 As Double? = data.Context.CantidadOriginal2
            Dim CantidadFinal2 As Double? = data.Context.CantidadConSigno2

            data.Context.StockFisico2 = data.Context.StockFisico2 - CantidadOriginal2 + CantidadFinal2
        End If


        Dim Rules() As Regla = New ProcesoStocks().EstablecerReglas(Regla.StockNegativo)
        Dim datValCtx As New DataValidarContexto(data.Articulo, data.Almacen, data.Context, Rules)
        ProcessServer.ExecuteTask(Of DataValidarContexto)(AddressOf ValidarContexto, datValCtx, services)
    End Sub

    <Task()> Public Shared Sub ConvertirMovimientoEnCorreccion(ByVal data As DataMovimiento, ByVal services As ServiceProvider)
        '//Hacer un copia del movimiento original y guardarla como movimiento de correccion
        Dim newrow As DataRow = data.stkData.Context.Movimientos.NewRow
        newrow.ItemArray = data.Movimiento.ItemArray
        newrow("IDLineaMovimiento") = AdminData.GetAutoNumeric
        newrow("IDTipoMovimiento") = enumTipoMovimiento.tmCorreccion
        newrow("FechaMovimiento") = Today
        newrow("Texto") = DataRowState.Modified.ToString
        data.stkData.Context.Movimientos.Rows.Add(newrow)
    End Sub

    <Task()> Public Shared Sub ConvertirMovimientoEnCorreccionPorBorrado(ByVal data As DataMovimiento, ByVal services As ServiceProvider)
        '//Cuando eliminamos 
        data.Movimiento("IDTipoMovimiento") = enumTipoMovimiento.tmCorreccion
        data.Movimiento("Texto") = DataRowState.Deleted.ToString

        Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.stkData.Articulo, services)
        Dim datAcum As New DataAcumulado(data.stkData, data.Movimiento("IDLineaMovimiento"), SegundaUnidad)
        Dim datValorAcum As DataValorAcumulado = ProcessServer.ExecuteTask(Of DataAcumulado, DataValorAcumulado)(AddressOf Acumulado, datAcum, services)
        data.Movimiento("Acumulado") = datValorAcum.Valor
        If SegundaUnidad AndAlso Not datValorAcum.Valor2 Is Nothing Then
            data.Movimiento("Acumulado2") = datValorAcum.Valor2
        End If

        If data.Movimiento("FechaDocumento") = data.stkData.Context.FechaUltimoMovimiento Then
            '//si es el ultimo movimiento de ese dia, hay que modificar la fecha de ultimo movimiento en Articulo-Almacen
            Dim datosUltMovto As New DataObtenerUltimoMovimientoVigente(data.stkData.Articulo, data.stkData.Almacen, CInt(data.Movimiento("IDLineaMovimiento")))
            Dim ultimoMovimiento As DataRow = ProcessServer.ExecuteTask(Of DataObtenerUltimoMovimientoVigente, DataRow)(AddressOf ObtenerUltimoMovimientoVigente, datosUltMovto, services)
            If Not ultimoMovimiento Is Nothing Then
                If data.stkData.Context.FechaUltimoMovimiento <> ultimoMovimiento("FechaDocumento") Then
                    data.stkData.Context.FechaUltimoMovimiento = ultimoMovimiento("FechaDocumento")
                End If
            Else
                data.stkData.Context.FechaUltimoMovimiento = Date.MinValue
            End If
        End If

        data.stkData.Context.Movimientos.ImportRow(data.Movimiento)
    End Sub


    Public Class DataGetPrecioMovimiento
        Public stk As StockData

        Public Movimiento As DataRow

        Public PrecioOutputA As Double

        Public PrecioOutputB As Double

        Public Sub New(ByVal stk As StockData, ByVal Movimiento As DataRow)
            Me.stk = stk
            Me.Movimiento = Movimiento
        End Sub
    End Class

    <Task()> Public Shared Function GetPrecioMovimiento(ByVal data As DataGetPrecioMovimiento, ByVal services As ServiceProvider) As DataGetPrecioMovimiento
        '////////////////////////////////////////
        '//Calculamos los precios A y B del movimiento, xq necesitamos saber el precio real del movimiento para calcular el precio de los movimientos posteriores que se calcula en la función del acumulado (precio medio).
        Dim PrecioMovimientoA As Double = data.stk.PrecioA
        Dim PrecioMovimientoB As Double = data.stk.PrecioB

        Select Case data.stk.Context.ClaseMovimiento
            Case enumtpmTipoMovimiento.tpmOutput
                '//Recalcular precio salidas
                Dim datosPrecioSal As New DataPrecioMovimiento(data.Movimiento("IDArticulo"), data.Movimiento("IDAlmacen"), data.Movimiento("FechaDocumento"), Math.Abs(data.stk.Cantidad), data.stk.Context.ClaseMovimiento, data.stk.PrecioA)
                datosPrecioSal.Movimiento = data.Movimiento.Table
                Dim precios As Hashtable = ProcessServer.ExecuteTask(Of DataPrecioMovimiento, Hashtable)(AddressOf PrecioMovimiento, datosPrecioSal, services)
                If Not precios Is Nothing Then
                    PrecioMovimientoA = precios("PrecioA")
                    PrecioMovimientoB = precios("PrecioB")
                End If
            Case Else
                If data.stk.PrecioA = 0 Or data.stk.PrecioB = 0 Then
                    Dim AppStkParams As ParametroStocks = services.GetService(Of ParametroStocks)()

                    'Los albaranes de compra con precio 0 se tienen que generar con PrecioA = 0
                    If data.Movimiento("IDTipoMovimiento") <> enumTipoMovimiento.tmEntAlbaranCompra Then
                        Dim PermitirPrecioCero As Boolean = AppStkParams.PrecioMovimientoCero()
                        If Not PermitirPrecioCero OrElse data.Movimiento("IDTipoMovimiento") = enumTipoMovimiento.tmInventario _
                                    OrElse data.Movimiento("IDTipoMovimiento") = enumTipoMovimiento.tmEntAjuste _
                                    OrElse data.Movimiento("IDTipoMovimiento") = enumTipoMovimiento.tmEntTransferencia _
                                    OrElse data.Movimiento("IDTipoMovimiento") = enumTipoMovimiento.tmEntFabrica Then

                            Dim dataPrecioEnt As New DataPrecioMovimiento(data.Movimiento("IDArticulo"), data.Movimiento("IDAlmacen"), data.Movimiento("FechaDocumento"), data.stk.Cantidad, data.stk.Context.ClaseMovimiento, data.stk.PrecioA)
                            dataPrecioEnt.Movimiento = data.Movimiento.Table
                            Dim precios As Hashtable = ProcessServer.ExecuteTask(Of DataPrecioMovimiento, Hashtable)(AddressOf PrecioMovimiento, dataPrecioEnt, services)
                            If Not precios Is Nothing Then
                                PrecioMovimientoA = precios("PrecioA")
                                PrecioMovimientoB = precios("PrecioB")
                            End If
                        End If
                    End If
                End If
        End Select

        data.PrecioOutputA = PrecioMovimientoA
        data.PrecioOutputB = PrecioMovimientoB

        '///////////////////////////
        Return data
    End Function

    <Task()> Public Shared Sub NuevoMovimientoTrasCorreccion(ByVal data As DataMovimiento, ByVal services As ServiceProvider)
        '//Volcar los nuevos valores sobre el movimiento original

        '//////////////////////////////////////////
        Dim datPrecio As New DataGetPrecioMovimiento(data.stkData, data.Movimiento)
        datPrecio = ProcessServer.ExecuteTask(Of DataGetPrecioMovimiento, DataGetPrecioMovimiento)(AddressOf GetPrecioMovimiento, datPrecio, services)

        data.stkData.PrecioA = datPrecio.PrecioOutputA
        data.stkData.PrecioB = datPrecio.PrecioOutputB
        '//////////////////////////////////////////

        Dim CantidadFinal As Double = data.stkData.Context.CantidadConSigno
        data.Movimiento("Cantidad") = CantidadFinal
        Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.stkData.Articulo, services)
        If SegundaUnidad AndAlso Length(data.stkData.Context.CantidadConSigno2) > 0 Then data.Movimiento("Cantidad2") = data.stkData.Context.CantidadConSigno2
        If data.stkData.ContextCorrect.CorreccionEnPrecio OrElse data.stkData.ContextCorrect.CorreccionEnCantidad OrElse data.stkData.ContextCorrect.CorreccionEnCantidad2 OrElse data.stkData.ContextCorrect.CorreccionEnFecha Then
            If data.stkData.ContextCorrect.CorreccionEnFecha Then
                Dim dataAcum As New DataAcumulado(data.stkData, data.Movimiento("IDLineaMovimiento"), SegundaUnidad)
                Dim datValorAcum As DataValorAcumulado = ProcessServer.ExecuteTask(Of DataAcumulado, DataValorAcumulado)(AddressOf AcumuladoCorreccionEnFecha, dataAcum, services)
                data.Movimiento("Acumulado") = datValorAcum.Valor
                If SegundaUnidad AndAlso Not datValorAcum.Valor2 Is Nothing Then
                    data.Movimiento("Acumulado2") = datValorAcum.Valor2
                End If
                If data.stkData.FechaDocumento > data.stkData.Context.FechaUltimoMovimiento Then
                    data.stkData.Context.FechaUltimoMovimiento = data.stkData.FechaDocumento
                ElseIf data.stkData.Context.FechaDocumentoOriginal = data.stkData.Context.FechaUltimoMovimiento Then
                    Dim datosUltMovto As New DataObtenerUltimoMovimientoVigente(data.stkData.Articulo, data.stkData.Almacen, CInt(data.Movimiento("IDLineaMovimiento")))
                    Dim ultimoMovimiento As DataRow = ProcessServer.ExecuteTask(Of DataObtenerUltimoMovimientoVigente, DataRow)(AddressOf ObtenerUltimoMovimientoVigente, datosUltMovto, services)
                    If Not ultimoMovimiento Is Nothing Then
                        If data.stkData.Context.FechaUltimoMovimiento <> ultimoMovimiento("FechaDocumento") Then
                            data.stkData.Context.FechaUltimoMovimiento = ultimoMovimiento("FechaDocumento")
                        End If
                    Else
                        data.stkData.Context.FechaUltimoMovimiento = data.stkData.FechaDocumento
                    End If
                End If
            Else
                Dim datAcum As New DataAcumulado(data.stkData, data.Movimiento("IDLineaMovimiento"), SegundaUnidad)
                Dim datValorAcum As DataValorAcumulado = ProcessServer.ExecuteTask(Of DataAcumulado, DataValorAcumulado)(AddressOf Acumulado, datAcum, services)
                data.Movimiento("Acumulado") = datValorAcum.Valor
                If SegundaUnidad AndAlso Not datValorAcum.Valor2 Is Nothing Then
                    data.Movimiento("Acumulado2") = datValorAcum.Valor2
                End If
            End If
        End If
        data.Movimiento("FechaDocumento") = data.stkData.FechaDocumento
        data.Movimiento("Documento") = data.stkData.Documento
        data.Movimiento("PrecioA") = data.stkData.PrecioA
        data.Movimiento("PrecioB") = data.stkData.PrecioB

        Dim dtArticulo As DataTable = New Articulo().SelOnPrimaryKey(data.Movimiento("IDArticulo"))
        Dim DataCalc As New DataCalcValPrecioMedio(data.Movimiento, data.stkData.Context.ClaseMovimiento)
        data.Movimiento("PrecioMedio") = ProcessServer.ExecuteTask(Of DataCalcValPrecioMedio, Double)(AddressOf CalculoValAlmPrecioMedio, DataCalc, services)
        data.Movimiento("PrecioUltimaCompra") = Nz(dtArticulo.Rows(0)("PrecioUltimaCompraA"), 0)
        data.Movimiento("PrecioEstandar") = Nz(dtArticulo.Rows(0)("PrecioEstandarA"), 0) / Nz(dtArticulo.Rows(0)("UDValoracion"), 1)

        'Dim f As New Filter
        'f.Add(New StringFilterItem("IDArticulo", data.Movimiento("IDArticulo")))
        'f.Add(New StringFilterItem("IDAlmacen", data.Movimiento("IDAlmacen")))

        'Dim dtArticuloAlmacen As DataTable = New ArticuloAlmacen().Filter(f)

        'El FIFO SE CALCULA DESPUES DE ACTUALIZAR (COMENTADO)
        ''Dim datosValFIFO As New ProcesoStocks.DataValoracionFIFO(data.Movimiento("IDArticulo"), data.Movimiento("IDAlmacen"), data.Movimiento("Acumulado"), data.Movimiento("Acumulado"), data.Movimiento("FechaDocumento"), enumstkValoracionFIFO.stkVFOrdenarPorFecha)
        ''Dim valoracion As ValoracionPreciosInfo
        ''valoracion = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ProcesoStocks.ValoracionFIFO, datosValFIFO, services)
        ''If Not valoracion Is Nothing Then
        ''    data.Movimiento("FifoFD") = valoracion.PrecioA
        ''End If
        ''datosValFIFO.Orden = enumstkValoracionFIFO.stkVFOrdenarPorMvto
        ''valoracion = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ProcesoStocks.ValoracionFIFO, datosValFIFO, services)
        ''If Not valoracion Is Nothing Then
        ''    data.Movimiento("FifoF") = valoracion.PrecioA
        ''End If
        data.Movimiento("FifoFD") = 0
        data.Movimiento("FifoF") = 0
        data.stkData.Context.Movimientos.ImportRow(data.Movimiento)
    End Sub

    <Serializable()> _
    Public Class DataCalcValPrecioMedio
        Public DrMov As DataRow
        Public TipoMov As enumtpmTipoMovimiento

        Public Sub New()
        End Sub
        Public Sub New(ByVal DrMov As DataRow, ByVal TipoMov As enumtpmTipoMovimiento)
            Me.DrMov = DrMov
            Me.TipoMov = TipoMov
        End Sub
    End Class

    <Task()> Public Shared Function CalculoValAlmPrecioMedio(ByVal data As DataCalcValPrecioMedio, ByVal services As ServiceProvider) As Double
        Dim DblPrecioReturn As Double = 0
        Dim DblPrecioMedio As Double = 0
        Dim DblCantidad As Double = 0
        Dim FilHist As New Filter
        Dim FilHistOr As New Filter(FilterUnionOperator.Or)
        Dim FilHistAnd As New Filter

        FilHist.Add("IDArticulo", FilterOperator.Equal, data.DrMov("IDArticulo"))
        FilHist.Add("IDAlmacen", FilterOperator.Equal, data.DrMov("IDAlmacen"))
        FilHist.Add("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion)
        FilHist.Add("IDLineaMovimiento", FilterOperator.NotEqual, data.DrMov("IDLineaMovimiento"))

        FilHistOr.Add("FechaDocumento", FilterOperator.LessThan, data.DrMov("FechaDocumento"))

        FilHistAnd.Add("FechaDocumento", FilterOperator.Equal, data.DrMov("FechaDocumento"))
        FilHistAnd.Add("IDLineaMovimiento", FilterOperator.LessThan, data.DrMov("IDLineaMovimiento"))
        FilHistOr.Add(FilHistAnd)
        FilHist.Add(FilHistOr)

        Dim DtHistFind As DataTable = New BE.DataEngine().Filter("tbHistoricoMovimiento", FilHist, " top 2 PrecioMedio, Acumulado, IDTipoMovimiento, FechaDocumento", "FechaDocumento Desc, IDLineaMovimiento DESC")
        If Not DtHistFind Is Nothing AndAlso DtHistFind.Rows.Count > 0 Then

            Dim AppParamsStocks As ParametroStocks = services.GetService(Of ParametroStocks)()
            If DtHistFind.Rows(0)("IDTipoMovimiento") = enumTipoMovimiento.tmInventario AndAlso AppParamsStocks.TipoInventario = TipoInventario.PrimerMovimiento AndAlso DtHistFind.Rows.Count > 1 Then
                Dim FechaDocumentoInventario As Date = DtHistFind.Rows(0)("FechaDocumento")
                If DtHistFind.Rows(1)("FechaDocumento") = FechaDocumentoInventario Then
                    DblPrecioMedio = Nz(DtHistFind.Rows(1)("PrecioMedio"), 0)
                    DblCantidad = DtHistFind.Rows(1)("Acumulado")
                Else
                    DblPrecioMedio = Nz(DtHistFind.Rows(0)("PrecioMedio"), 0)
                    DblCantidad = DtHistFind.Rows(0)("Acumulado")
                End If
            Else
                DblPrecioMedio = Nz(DtHistFind.Rows(0)("PrecioMedio"), 0)
                DblCantidad = DtHistFind.Rows(0)("Acumulado")
            End If

            Select Case data.TipoMov
                Case enumtpmTipoMovimiento.tpmInput
                    If DblCantidad + Nz(data.DrMov("Cantidad"), 0) > 0 Then
                        If DblCantidad > 0 Then


                            DblPrecioMedio = (((Nz(data.DrMov("Cantidad"), 0) * Nz(data.DrMov("PrecioA"), 0)) + (DblCantidad * DblPrecioMedio)) / (DblCantidad + Nz(data.DrMov("Cantidad"), 0)))
                        Else : DblPrecioMedio = data.DrMov("PrecioA")
                        End If
                    End If
                Case enumtpmTipoMovimiento.tpmOutput
                    DblPrecioMedio = DblPrecioMedio
            End Select
        Else : DblPrecioMedio = Nz(data.DrMov("PrecioA"), 0)
        End If
        Return DblPrecioMedio
    End Function

    <Task()> Public Shared Function CalculoValAlmPrecioMedioInventario(ByVal data As DataCalcValPrecioMedio, ByVal services As ServiceProvider) As Double

        Dim DblPrecioReturn As Double = 0
        Dim DblPrecioMedio As Double = 0
        Dim DblCantidad As Double = 0
        Dim FilHist As New Filter

        Dim AppParamsStock As ParametroStocks = services.GetService(Of ParametroStocks)()

        FilHist.Add("IDArticulo", FilterOperator.Equal, data.DrMov("IDArticulo"))
        FilHist.Add("IDAlmacen", FilterOperator.Equal, data.DrMov("IDAlmacen"))
        If AppParamsStock.TipoInventario = TipoInventario.PrimerMovimiento Then
            FilHist.Add("FechaDocumento", FilterOperator.LessThanOrEqual, DateAdd(DateInterval.Day, -1, data.DrMov("FechaDocumento")))
        Else
            FilHist.Add("FechaDocumento", FilterOperator.LessThanOrEqual, data.DrMov("FechaDocumento"))
        End If
        FilHist.Add("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion)
        FilHist.Add("IDLineaMovimiento", FilterOperator.NotEqual, data.DrMov("IDLineaMovimiento"))


        Dim DtHistFind As DataTable = New BE.DataEngine().Filter("tbHistoricoMovimiento", FilHist, " top 1 PrecioMedio, Acumulado", "FechaDocumento Desc, IDLineaMovimiento DESC")
        If Not DtHistFind Is Nothing AndAlso DtHistFind.Rows.Count > 0 Then
            DblPrecioMedio = Nz(DtHistFind.Rows(0)("PrecioMedio"), 0)
        Else : DblPrecioMedio = data.DrMov("PrecioA")
        End If
        Return DblPrecioMedio



    End Function
#End Region

#Region "Transferencias"

    <Serializable()> _
    Public Class DataTransferenciaOrdenServicio
        Public IDArticulo As String
        Public NSerie As String
        Public IDAlmacenOrigen As String
        Public IDAlmacenDestino As String

        Public Sub New(ByVal IDArticulo As String, ByVal NSerie As String, ByVal IDAlmacenOrigen As String, ByVal IDAlmacenDestino As String)
            Me.IDArticulo = IDArticulo
            Me.NSerie = NSerie
            Me.IDAlmacenOrigen = IDAlmacenOrigen
            Me.IDAlmacenDestino = IDAlmacenDestino
        End Sub
    End Class

    <Task()> Public Shared Function TransferenciaOrdenServicio(ByVal data As DataTransferenciaOrdenServicio, ByVal services As ServiceProvider) As StockUpdateData()
        '//Realizamos el movimiento de un Almacen a otro
        Dim objFilter As New Filter
        objFilter.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
        objFilter.Add(New StringFilterItem("NSerie", data.NSerie))
        objFilter.Add(New StringFilterItem("IDAlmacen", data.IDAlmacenOrigen))

        Dim dtLote1 As DataTable = New BE.DataEngine().Filter("vFrmDesgloseNumeroDeSerie", objFilter)

        If Not IsNothing(dtLote1) AndAlso dtLote1.Rows.Count = 1 Then
            Dim origen As New StockData(data.IDArticulo, data.IDAlmacenOrigen, 1, 0, 0, Today, enumTipoMovimiento.tmSalTransferencia)
            origen.NSerie = dtLote1.Rows(0)("NSerie")
            origen.Activo = dtLote1.Rows(0)("IDActivo")
            origen.EstadoNSerie = dtLote1.Rows(0)("IDEstadoActivo")
            Dim strIDOperario As String = dtLote1.Rows(0)("IDOperario") & String.Empty
            If Length(strIDOperario) = 0 Then
                strIDOperario = ProcessServer.ExecuteTask(Of Object, String)(AddressOf Operario.ObtenerIDOperarioUsuario, Nothing, services)
            End If
            origen.Operario = strIDOperario

            Dim destino As New StockData(data.IDArticulo, data.IDAlmacenDestino, 1, 0, 0, Today, enumTipoMovimiento.tmEntTransferencia)
            destino.NSerie = dtLote1.Rows(0)("NSerie")
            destino.Activo = dtLote1.Rows(0)("IDActivo")
            destino.EstadoNSerie = dtLote1.Rows(0)("IDEstadoActivo")
            destino.Operario = strIDOperario

            Dim NumeroMovimiento As Integer = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf ProcesoStocks.NuevoNumeroMovimiento, Nothing, services)
            Return ProcessServer.ExecuteTask(Of DataTransferencia, StockUpdateData())(AddressOf Transferencia, New DataTransferencia(NumeroMovimiento, origen, destino), services)
        End If
    End Function

    <Serializable()> _
    Public Class DataTransferencia
        Public NumeroMovimiento As Integer
        Public stkDataOrigen As StockData
        Public stkDataDestino As StockData

        Public Sub New(ByVal NumeroMovimiento As Integer, ByVal stkDataOrigen As StockData, ByVal stkDataDestino As StockData)
            Me.NumeroMovimiento = NumeroMovimiento
            Me.stkDataOrigen = stkDataOrigen
            Me.stkDataDestino = stkDataDestino
        End Sub
    End Class

    ''' <summary>
    ''' Método para hacer la Transferencia entre almacenes cuando el artículo NO lleva gestión por Lotes o NSerie.
    ''' </summary>
    ''' <param name="data">Objeto con el Número de Movimiento, y los datos de Origen y Destino referentes a los distintos Artículos y Almacenes.</param>
    ''' <param name="services">Objeto para el cacheo de información a lo largo del proceso.</param>
    ''' <returns>Array con una serie de objetos indicando el estado en el que ha quedado cada uno de los elementos de la transferencia. </returns>
    ''' <remarks>Método para hacer la Transferencia entre almacenes cuando el artículo NO lleva gestión por Lotes o NSerie.</remarks>
    <Task()> Public Shared Function Transferencia(ByVal data As DataTransferencia, ByVal services As ServiceProvider) As StockUpdateData()
        '//Metodo para transferencias de articulo sin gestion por lotes ni gestion por numeros de serie
        If Not data.stkDataOrigen Is Nothing And Not data.stkDataDestino Is Nothing Then
            Dim updateData(1) As StockUpdateData
            updateData(0) = New StockUpdateData
            updateData(1) = New StockUpdateData

            updateData(0).StockData = data.stkDataOrigen
            updateData(1).StockData = data.stkDataDestino

            If data.stkDataOrigen.Articulo <> data.stkDataDestino.Articulo Then
                updateData(0).Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(13), services)
                Return updateData
            ElseIf data.stkDataOrigen.Almacen = data.stkDataDestino.Almacen Then
                updateData(1).Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(14), services)
                Return updateData
            Else
                Dim Articulo As String = data.stkDataOrigen.Articulo
                Dim AlmacenOrigen As String = data.stkDataOrigen.Almacen
                Dim AlmacenDestino As String = data.stkDataDestino.Almacen
                Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, Articulo, services)

                If Len(Articulo) = 0 Or Len(AlmacenOrigen) = 0 Or Len(AlmacenDestino) = 0 Or data.stkDataOrigen.Cantidad <= 0 OrElse (SegundaUnidad AndAlso data.stkDataOrigen.Cantidad2 <= 0) Then
                    If data.stkDataOrigen.Cantidad <= 0 OrElse (SegundaUnidad AndAlso data.stkDataOrigen.Cantidad2 <= 0) Then
                        Dim datMsg As New DataMessage(8)
                        datMsg.SegundaUnidad = SegundaUnidad
                        updateData(0).Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, datMsg, services)
                        Return updateData
                    ElseIf Len(Articulo) = 0 Then
                        updateData(0).Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(22), services)
                        Return updateData
                    ElseIf Len(AlmacenOrigen) = 0 Then
                        updateData(0).Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(23), services)
                        Return updateData
                    ElseIf Len(AlmacenDestino) = 0 Then
                        updateData(1).Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(23), services)
                        Return updateData
                    End If
                Else
                    AdminData.BeginTx()
                    Dim dataSal As New DataNumeroMovimientoSinc(data.NumeroMovimiento, data.stkDataOrigen)
                    Dim updateDataSalida As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf Salida, dataSal, services)
                    If Not updateDataSalida Is Nothing Then
                        updateData(0) = updateDataSalida
                        If Not (updateDataSalida.Estado = EstadoStock.Actualizado) Then
                            Return updateData
                        Else
                            data.stkDataDestino.PrecioA = data.stkDataOrigen.PrecioA
                            data.stkDataDestino.PrecioB = data.stkDataOrigen.PrecioB

                            Dim datMovto As New DataNumeroMovimientoSinc(data.NumeroMovimiento, data.stkDataDestino)
                            Dim updateDataEntrada As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf Entrada, datMovto, services)
                            If Not updateDataEntrada Is Nothing Then
                                updateData(1) = updateDataEntrada
                                If Not (updateDataEntrada.Estado = EstadoStock.Actualizado) Then
                                    updateDataSalida.Estado = updateDataEntrada.Estado
                                    updateDataSalida.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(40), services)
                                    AdminData.RollBackTx()
                                    Return updateData
                                End If
                            End If
                        End If
                    End If
                    AdminData.CommitTx()
                End If
            End If

            Return updateData
        End If
    End Function

    <Serializable()> _
   Public Class DataTransferenciaDesgloseLote
        Public NumeroMovimiento As Integer
        Public Origen() As StockData
        Public Destino() As StockData

        Public Sub New(ByVal NumeroMovimiento As Integer, ByVal Origen() As StockData, ByVal Destino() As StockData)
            Me.NumeroMovimiento = NumeroMovimiento
            Me.Origen = Origen
            Me.Destino = Destino
        End Sub
    End Class
    ''' <summary>
    ''' Método para hacer la Transferencia entre almacenes cuando el artículo lleva gestión por Lotes o NSerie.
    ''' </summary>
    ''' <param name="data">Objeto con el Número de Movimiento, y los datos de Origen y Destino referentes a los distintos Lotes/NSerie</param>
    ''' <param name="services">Objeto para el cacheo de información a lo largo del proceso.</param>
    ''' <returns>Array con una serie de objetos indicando el estado en el que ha quedado cada uno de los elementos de la transferencia. </returns>
    ''' <remarks>Método para hacer la Transferencia entre almacenes cuando el artículo lleva gestión por Lotes o NSerie.</remarks>
    <Task()> Public Shared Function TransferenciaDesgloseLote(ByVal data As DataTransferenciaDesgloseLote, ByVal services As ServiceProvider) As StockUpdateData()
        '//Metodo para transferencias de articulo con gestion por lotes o por gestion por numeros de serie
        Dim updateData(-1) As StockUpdateData

        If Not data.Origen Is Nothing AndAlso data.Origen.Length > 0 And Not data.Destino Is Nothing AndAlso data.Destino.Length > 0 Then
            AdminData.BeginTx()
            For Each o As StockData In data.Origen
                Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, o.Articulo, services)
                If (Not SegundaUnidad AndAlso o.Cantidad > 0) OrElse (SegundaUnidad AndAlso o.Cantidad > 0 AndAlso o.Cantidad2 > 0) Then
                    Dim dataSal As New DataNumeroMovimientoSinc(data.NumeroMovimiento, o)
                    Dim updateDataSalida As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf Salida, dataSal, services)
                    If Not IsNothing(updateDataSalida) Then
                        If updateDataSalida.Estado = EstadoStock.Actualizado Then
                            ArrayManager.Copy(updateDataSalida, updateData)
                        Else
                            ArrayManager.Copy(updateDataSalida, updateData)
                            AdminData.RollBackTx()
                            Return updateData
                        End If
                    End If
                Else
                    ApplicationService.GenerateError("La cantidad interna para el lote {0} del artículo {1} no es válida.", Quoted(o.Lote), Quoted(o.Articulo))
                End If
            Next

            For Each d As StockData In data.Destino
                Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, d.Articulo, services)
                If (Not SegundaUnidad AndAlso d.Cantidad > 0) OrElse (SegundaUnidad AndAlso d.Cantidad > 0 AndAlso d.Cantidad2 > 0) Then
                    Dim datMovto As New DataNumeroMovimientoSinc(data.NumeroMovimiento, d)
                    Dim updateDataEntrada As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf Entrada, datMovto, services)
                    If Not IsNothing(updateDataEntrada) Then
                        If updateDataEntrada.Estado = EstadoStock.Actualizado Then
                            ArrayManager.Copy(updateDataEntrada, updateData)
                        Else
                            ArrayManager.Copy(updateDataEntrada, updateData)
                            AdminData.RollBackTx()
                            Return updateData
                        End If
                    End If
                Else
                    ApplicationService.GenerateError("La cantidad interna para el lote {0} del artículo {1} no es válida.", Quoted(d.Lote), Quoted(d.Articulo))
                End If
            Next
            AdminData.CommitTx()
        End If
        Return updateData
    End Function
#End Region

#Region "Sobrecarga de funciones para movimientos genericos de E/S"

    '<Task()> Public Shared Function MovimientosGenericosES(ByVal data As StockData(), Optional ByVal sinc As Boolean = True) As StockUpdateData()
    '    Dim services As New ServiceProvider

    '    Dim UpdateData(-1) As StockUpdateData
    '    If Not IsNothing(data) AndAlso data.Length > 0 Then
    '        Dim N As Integer = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf ProcesoStocks.NuevoNumeroMovimiento, Nothing, services)
    '        For Each item As StockData In data
    '            Dim aux As StockUpdateData = ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf MovimientosGenericosESTx, New DataNumeroMovimientoSinc(N, item, sinc), services)
    '            ArrayManager.Copy(aux, UpdateData)
    '        Next
    '    End If
    '    Return UpdateData
    'End Function

    <Serializable()> _
    Public Class DataMovimientosGenericosES
        Public NumeroMovimiento As Integer
        Public stksData() As StockData
        Public Sinc As Boolean

        Public Sub New(ByVal NumeroMovimiento As Integer, ByVal stksData() As StockData, Optional ByVal Sinc As Boolean = True)
            Me.NumeroMovimiento = NumeroMovimiento
            Me.stksData = stksData
            Me.Sinc = Sinc
        End Sub
    End Class

    <Task()> Public Shared Function MovimientosGenericosES(ByVal data As DataMovimientosGenericosES, ByVal services As ServiceProvider) As StockUpdateData()
        If Not data Is Nothing AndAlso Not data.stksData Is Nothing AndAlso data.stksData.Length > 0 Then
            Dim updateData(-1) As StockUpdateData
            For Each stkData As StockData In data.stksData
                Dim TiposMovimiento As EntityInfoCache(Of TipoMovimientoInfo) = services.GetService(Of EntityInfoCache(Of TipoMovimientoInfo))()
                Dim TMovtoInfo As TipoMovimientoInfo = TiposMovimiento.GetEntity(stkData.TipoMovimiento)

                Dim datOperaci As New DataNumeroMovimientoSinc(data.NumeroMovimiento, stkData, data.Sinc)

                Dim updtData As StockUpdateData
                Select Case TMovtoInfo.ClaseMovimiento
                    Case enumtpmTipoMovimiento.tpmInput
                        updtData = ProcessServer.ExecuteTask(Of DataNumeroMovimiento, StockUpdateData)(AddressOf EntradaTx, datOperaci, services)
                    Case enumtpmTipoMovimiento.tpmOutput
                        updtData = ProcessServer.ExecuteTask(Of DataNumeroMovimiento, StockUpdateData)(AddressOf SalidaTx, datOperaci, services)
                End Select

                If Not updtData Is Nothing Then
                    If updtData.Estado = EstadoStock.Actualizado Then
                        AdminData.BeginTx()
                        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
                        Dim AppParamsStk As ParametroStocks = services.GetService(Of ParametroStocks)()
                        If data.Sinc AndAlso AppParams.GestionBodegas AndAlso stkData.TipoMovimiento <> AppParamsStk.TipoMovimientoCantidad0 Then
                            Dim datSinc As DataIntegracionConBodega
                            Select Case TMovtoInfo.ClaseMovimiento
                                Case enumtpmTipoMovimiento.tpmInput
                                    datSinc = New DataIntegracionConBodega(datOperaci, enumTipoSincronizacion.Entrada, , , , updtData)
                                Case enumtpmTipoMovimiento.tpmOutput
                                    datSinc = New DataIntegracionConBodega(datOperaci, enumTipoSincronizacion.Salida, , , updtData)
                            End Select
                            Dim updateDataAux As StockUpdateData = ProcessServer.ExecuteTask(Of DataIntegracionConBodega, StockUpdateData)(AddressOf IntegracionConBodega, datSinc, services)
                            If Not updateDataAux Is Nothing Then
                                updtData = updateDataAux
                            End If
                        End If

                        ProcessServer.ExecuteTask(Of StockUpdateData)(AddressOf Actualizar, updtData, services)
                    End If

                    ReDim Preserve updateData(updateData.Length)
                    updateData(updateData.Length - 1) = updtData
                End If
            Next
            Return updateData
        End If
    End Function

#End Region

#Region " Métodos y clases generales para la gestión de los movimientos de Stock "

    <Task()> Public Shared Function Update(ByVal dt As DataTable, ByVal services As ServiceProvider) As DataTable
        BusinessHelper.UpdateTable(dt)
        Return dt
    End Function

#Region " Crear Contexto "

    Public Class DataCrearContexto
        Inherits DataNumeroMovimiento

        Public EsCorreccion As Boolean

        Public Sub New(ByVal NumeroMovimiento As Integer, ByVal stkData As StockData, Optional ByVal EsCorreccion As Boolean = False)
            MyBase.New(NumeroMovimiento, stkData)
            Me.EsCorreccion = EsCorreccion
        End Sub
    End Class
    <Task()> Public Shared Sub CrearContexto(ByVal data As DataCrearContexto, ByVal services As ServiceProvider)
        '//Crea un contexto para este proceso de actualizacion de stocks. 

        '//El contexto no es mas que una variable privada de la clase de stock que debe contener la informacion necesaria 
        '   para la ejecucion y validacion del proceso de stock actual. 
        '//Las reglas de validacion comunes a todos los procesos de stock se van a aplicar unicamente sobre el contexto.
        '1.Permite aplicar las reglas de actualizacion de stock con menos codigo.
        '2.Permite utilizar muchas menos variables locales en las funciones individuales.
        '3.Todo esto en conjunto permite descomponer los distintos procesos de actualizacion de stock y que todos
        '   utilicen el mismo contexto.
        If Not data.stkData Is Nothing Then
            ProcessServer.ExecuteTask(Of StockData)(AddressOf ValidarDatosObligatorios, data.stkData, services)

            data.stkData.Context = New StockContext(data.stkData.Articulo, data.stkData.Almacen, data.stkData.FechaDocumento, data.NumeroMovimiento)
            ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxFechaUltimoCierre, data.stkData, services)

            ProcessServer.ExecuteTask(Of StockData)(AddressOf ValidarTipoGestionArticulo, data.stkData, services)
            If Not data.stkData.Context.Cancel Then
                '//1.Obtener datos que no se encuentran ni en Articulo-Almacen, ni Articulo-Almacen-Lote, ni Articulo-NSerie
                Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.stkData.Articulo)

                ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxTipoMovimiento, data.stkData, services)
                ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxDatosPrincipales, data.stkData, services)
                If ArtInfo.GestionStockPorLotes Then
                    ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxDatosLote, data.stkData, services)
                ElseIf ArtInfo.GestionPorNumeroSerie Then
                    ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxDatosNSerie, data.stkData, services)
                End If

                '//2.Articulo-Almacen
                ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxDatosArticuloAlmacen, data.stkData, services)
                If Not data.stkData.Context.Cancel Then
                    If ArtInfo.GestionStockPorLotes Then
                        '//3.Articulo-Almacen-Lote
                        ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxDatosArticuloAlmacenLote, data.stkData, services)
                    ElseIf ArtInfo.GestionPorNumeroSerie Then
                        '//4.Articulo-NSerie 
                        ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxDatosArticuloAlmacenNSerie, data.stkData, services)
                        If Not data.stkData.Context.Cancel Then
                            If Not data.EsCorreccion Then
                                ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxPropiedadesEstadoActivo, data.stkData, services) ' ?
                            End If
                        End If
                    End If
                End If
            End If

            '//Movimientos - tabla vacia disponible para inserciones
            ProcessServer.ExecuteTask(Of StockData)(AddressOf SetCtxMovimientos, data.stkData, services)
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As StockData, ByVal services As ServiceProvider)
        If Len(data.Articulo) = 0 Then
            ApplicationService.GenerateError("El Artículo es un dato obligatorio en la gestión de stocks.")
        ElseIf Len(data.Almacen) = 0 Then
            ApplicationService.GenerateError("El Almacén es un dato obligatorio en la gestión de stocks.")
        ElseIf data.FechaDocumento = cnMinDate Then
            ApplicationService.GenerateError("La FechaDocumento es un dato obligatorio en la gestión de stocks.")
        End If
    End Sub
    <Task()> Public Shared Sub ValidarTipoGestionArticulo(ByVal data As StockData, ByVal services As ServiceProvider)
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Articulo)
        If Not ArtInfo.GestionStock Then
            data.Context.Cancel = True
            data.Context.Estado = EstadoStock.SinGestion
            data.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(2, data.Articulo), services)
        ElseIf ArtInfo.GestionStockPorLotes AndAlso ArtInfo.GestionPorNumeroSerie Then
            data.Context.Cancel = True
            data.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(3), services)
        End If
    End Sub
    <Task()> Public Shared Sub SetCtxTipoMovimiento(ByVal data As StockData, ByVal services As ServiceProvider)
        data.Context.TipoMovimiento = data.TipoMovimiento
        Dim TiposMovto As EntityInfoCache(Of TipoMovimientoInfo) = services.GetService(Of EntityInfoCache(Of TipoMovimientoInfo))()
        Dim TMvtoInfo As TipoMovimientoInfo = TiposMovto.GetEntity(data.TipoMovimiento)
        data.Context.ClaseMovimiento = TMvtoInfo.ClaseMovimiento
        ProcessServer.ExecuteTask(Of StockData)(AddressOf SignoCantidad, data, services)
    End Sub
    <Task()> Public Shared Sub SignoCantidad(ByVal data As StockData, ByVal services As ServiceProvider)
        '//Para los movimientos normales de entrada o salida, la cantidad se pone siempre en positivo en presentacion,
        '//y en negocio se pone el signo, de acuerdo con el IDTipoMovimiento.
        '//Si son devoluciones la cantidad se pone siempre en negativo en presentacion, tanto si son devoluciones de entradas 
        '//o devoluciones de salida.

        '//La variable 'CantidadConSigno' es la cantidad de presentacion pero que ya tiene en cuenta el tipo de movimiento.
        '//Para los movimientos de correccion 'CantidadConSigno' "se corrige" en el metodo CrearContextoCorreccion.

        If data.Context.ClaseMovimiento = enumtpmTipoMovimiento.tpmOutput Then
            data.Context.CantidadConSigno = -data.Cantidad
        Else
            data.Context.CantidadConSigno = data.Cantidad
        End If

        If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Articulo, services) Then
            If data.Context.ClaseMovimiento = enumtpmTipoMovimiento.tpmOutput Then
                data.Context.CantidadConSigno2 = -data.Cantidad2
            Else
                data.Context.CantidadConSigno2 = data.Cantidad2
            End If
        End If
    End Sub
    <Task()> Public Shared Sub SetCtxDatosPrincipales(ByVal data As StockData, ByVal services As ServiceProvider)
        data.Context.PrecioA = data.PrecioA
        data.Context.PrecioB = data.PrecioB
        data.Context.Documento = data.Documento
        data.Context.IDDocumento = data.IDDocumento
        data.Context.Obra = data.Obra
        data.Context.Operario = data.Operario
    End Sub
    <Task()> Public Shared Sub SetCtxDatosLote(ByVal data As StockData, ByVal services As ServiceProvider)
        data.Context.Lote = data.Lote
        data.Context.Ubicacion = data.Ubicacion
    End Sub
    <Task()> Public Shared Sub SetCtxDatosNSerie(ByVal data As StockData, ByVal services As ServiceProvider)
        data.Context.NSerie = data.NSerie
        data.Context.Activo = data.NSerie
        data.Context.EstadoNSerie = data.EstadoNSerie
        data.Context.Operario = data.Operario
        data.Context.IDEstadoActivoAnterior = data.EstadoNSerieAnterior
    End Sub
    <Task()> Public Shared Sub SetCtxDatosArticuloAlmacen(ByVal data As StockData, ByVal services As ServiceProvider)
        Dim ArtAlm As New DataArticuloAlmacen(data.Articulo, data.Almacen)
        data.Context.ArticuloAlmacen = ProcessServer.ExecuteTask(Of DataArticuloAlmacen, DataTable)(AddressOf GetRegistroArticuloAlmacen, ArtAlm, services)
        If Not IsNothing(data.Context.ArticuloAlmacen) AndAlso data.Context.ArticuloAlmacen.Rows.Count > 0 Then
            'Si el articulo tiene un articulo generico para este almacen, a partir de aqui se utilizara el articulo generico
            If data.Articulo <> data.Context.ArticuloAlmacen.Rows(0)("IDArticulo") Then
                data.Context.ArticuloGenerico = data.Context.ArticuloAlmacen.Rows(0)("IDArticulo")
                'pend stocks: no esta probado
                data.Articulo = data.Context.ArticuloGenerico '?
                data.Context.Articulo = data.Articulo '?
            End If
            If IsDate(data.Context.ArticuloAlmacen.Rows(0)("FechaUltimoInventario")) Then
                data.Context.FechaUltimoInventario = data.Context.ArticuloAlmacen.Rows(0)("FechaUltimoInventario")
            End If
            If IsDate(data.Context.ArticuloAlmacen.Rows(0)("FechaUltimoMovimiento")) Then
                data.Context.FechaUltimoMovimiento = data.Context.ArticuloAlmacen.Rows(0)("FechaUltimoMovimiento")
            End If
            data.Context.StockFisico = data.Context.ArticuloAlmacen.Rows(0)("StockFisico")
            If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Articulo, services) Then
                data.Context.StockFisico2 = CDbl(Nz(data.Context.ArticuloAlmacen.Rows(0)("StockFisico2"), 0))
            End If
        Else
            data.Context.Cancel = True
            data.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(26), services)
        End If
    End Sub
    <Task()> Public Shared Sub SetCtxDatosArticuloAlmacenLote(ByVal data As StockData, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of StockData)(AddressOf ValidarDatosLote, data, services)
        If Not data.Context.Cancel Then
            Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Articulo, services)
            '//Recuperamos los datos del Lote
            Dim AAL As New ArticuloAlmacenLote
            data.Context.LoteBBDD = AAL.SelOnPrimaryKey(data.Articulo, data.Almacen, data.Context.Lote, data.Context.Ubicacion)
            If data.Context.LoteBBDD.Rows.Count > 0 Then
                data.Context.StockFisicoLote = data.Context.LoteBBDD.Rows(0)("StockFisico")
                If SegundaUnidad Then data.Context.StockFisicoLote2 = CDbl(Nz(data.Context.LoteBBDD.Rows(0)("StockFisico2"), 0))
                data.Context.LoteBloqueado = data.Context.LoteBBDD.Rows(0)("Bloqueado")
            Else
                If SegundaUnidad Then data.Context.StockFisicoLote2 = 0
            End If

            Dim sqlText As String
            sqlText = String.Concat("SELECT SUM(StockFisico) ", IIf(SegundaUnidad, ",SUM(StockFisico2) ", String.Empty), "FROM ", AAL.Table, " GROUP BY IDArticulo,IDAlmacen")
            sqlText = String.Concat(sqlText, " HAVING IDArticulo=", Quoted(data.Articulo), " AND IDAlmacen=", Quoted(data.Almacen))
            Dim dt As DataTable = AdminData.Execute(sqlText, ExecuteCommand.ExecuteReader)
            If dt.Rows.Count > 0 Then
                data.Context.StockFisico = dt.Rows(0)(0)
                If SegundaUnidad Then data.Context.StockFisico2 = CDbl(dt.Rows(0)(1))
            Else
                data.Context.StockFisico = 0
                If SegundaUnidad Then data.Context.StockFisico2 = 0
            End If
        End If

    End Sub
    <Task()> Public Shared Sub ValidarDatosLote(ByVal data As StockData, ByVal services As ServiceProvider)
        Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Articulo, services)
        If Len(data.Context.Lote) = 0 AndAlso ((data.Context.CantidadConSigno <> 0 OrElse (SegundaUnidad AndAlso data.Context.CantidadConSigno2 <> 0))) Then
            data.Context.Cancel = True
            data.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(11), services)
        ElseIf Len(data.Context.Ubicacion) = 0 AndAlso ((data.Context.CantidadConSigno <> 0 OrElse (SegundaUnidad AndAlso data.Context.CantidadConSigno2 <> 0))) Then
            data.Context.Cancel = True
            data.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(19), services)
        End If
    End Sub
    <Task()> Public Shared Sub ValidarDatosNSerie(ByVal data As StockData, ByVal services As ServiceProvider)
        If Len(data.Context.NSerie) = 0 Then
            data.Context.Cancel = True
            data.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(31), services)
        End If
    End Sub
    <Task()> Public Shared Sub SetCtxDatosArticuloAlmacenNSerie(ByVal data As StockData, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of StockData)(AddressOf ValidarDatosNSerie, data, services)
        If Not data.Context.Cancel Then
            data.Context.SerieBBDD = New ArticuloNSerie().SelOnPrimaryKey(data.Context.Articulo, data.Context.NSerie)
            If data.Context.SerieBBDD.Rows.Count > 0 Then
                If Length(data.Context.IDEstadoActivoAnterior) = 0 Then data.Context.IDEstadoActivoAnterior = data.Context.SerieBBDD.Rows(0)("IDEstadoActivo")
                'data.EstadoNSerieAnterior = data.Context.IDEstadoActivoAnterior
                Dim estado As DataTable = New BE.DataEngine().Filter("tbMntoEstadoActivo", New StringFilterItem("IDEstadoActivo", data.Context.SerieBBDD.Rows(0)("IDEstadoActivo")))
                If estado.Rows.Count > 0 Then
                    data.Context.PropiedadesEstadoBBDD.Disponible = estado.Rows(0)("Disponible")
                    data.Context.PropiedadesEstadoBBDD.EnCurso = estado.Rows(0)("EnCurso")
                    data.Context.PropiedadesEstadoBBDD.Baja = estado.Rows(0)("Baja")
                    data.Context.PropiedadesEstadoBBDD.Sistema = estado.Rows(0)("Sistema")
                End If
            End If
        End If
    End Sub
    <Task()> Public Shared Sub SetCtxPropiedadesEstadoActivo(ByVal data As StockData, ByVal services As ServiceProvider)
        If Len(data.Context.EstadoNSerie) = 0 AndAlso (data.Cantidad <> 0 OrElse Nz(data.Cantidad2, 0) <> 0) Then
            data.Context.Cancel = True
            data.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(32), services)
        ElseIf Len(data.Context.Operario) = 0 AndAlso (data.Cantidad <> 0 OrElse Nz(data.Cantidad2, 0) <> 0) Then
            data.Context.Cancel = True
            data.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(33), services)
        Else
            If Len(data.Context.EstadoNSerie) > 0 Then
                Dim estado As DataTable = New BE.DataEngine().Filter("tbMntoEstadoActivo", New StringFilterItem("IDEstadoActivo", data.Context.EstadoNSerie))
                If estado.Rows.Count > 0 Then
                    data.Context.PropiedadesEstado.Disponible = estado.Rows(0)("Disponible")
                    data.Context.PropiedadesEstado.EnCurso = estado.Rows(0)("EnCurso")
                    data.Context.PropiedadesEstado.Baja = estado.Rows(0)("Baja")
                    data.Context.PropiedadesEstado.Sistema = estado.Rows(0)("Sistema")
                End If
            End If
        End If
    End Sub
    <Task()> Public Shared Sub SetCtxMovimientos(ByVal data As StockData, ByVal services As ServiceProvider)
        '//Movimientos - tabla vacia disponible para inserciones
        data.Context.Movimientos = AdminData.GetEntityData(cnMyClass, "", , , True)
    End Sub

#End Region

#Region " Validar Contexto "

    Public Enum Regla
        FechaUltimoCierre
        ArticuloDePortes
        AlmacenBloqueado
        AlmacenActivo
        NumeroMovimiento
        StockNegativo
        FechaUltimoInventario
        FechaAjuste
        CantidadPositiva
        LoteObligatorio
        LoteBloqueado
        StockLoteNegativo
        SerieObligatoria
        SerieUnica
        AlbaranVenta
        CantidadCero
        StockSerieNegativo
    End Enum

    Public Class DataValidarContexto
        Public IDArticulo As String
        Public IDAlmacen As String
        Public stkContext As StockContext
        Public Rules() As Regla

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal stkContext As StockContext, ByVal Rules() As Regla)
            Me.IDArticulo = IDArticulo
            Me.IDAlmacen = IDAlmacen
            Me.stkContext = stkContext
            Me.Rules = Rules
        End Sub
    End Class
    <Task()> Public Shared Sub ValidarContexto(ByVal data As DataValidarContexto, ByVal services As ServiceProvider)
        '//Reglas de validacion de stock, más o menos comunes en las operaciones de actualización de stock
        '//(es decir, todas aquellas que dependen más o menos del tipo de movimiento o que no se aplican siempre)
        If Not data.stkContext Is Nothing Then
            For Each r As Regla In data.Rules
                If Not data.stkContext.Cancel Then
                    Select Case r
                        Case Regla.NumeroMovimiento
                            ProcessServer.ExecuteTask(Of StockContext)(AddressOf ValidarCtxNumeroMovimiento, data.stkContext, services)
                        Case Regla.CantidadCero
                            ProcessServer.ExecuteTask(Of StockContext)(AddressOf ValidarCtxCantidadCero, data.stkContext, services)
                        Case Regla.FechaUltimoCierre
                            ProcessServer.ExecuteTask(Of StockContext)(AddressOf ValidarCtxFechaUltimoCierre, data.stkContext, services)
                        Case Regla.StockNegativo
                            ProcessServer.ExecuteTask(Of StockContext)(AddressOf ValidarCtxStockNegativo, data.stkContext, services)
                        Case Regla.FechaUltimoInventario
                            ProcessServer.ExecuteTask(Of StockContext)(AddressOf ValidarCtxFechaUltimoInventario, data.stkContext, services)
                        Case Regla.CantidadPositiva
                            ProcessServer.ExecuteTask(Of StockContext)(AddressOf ValidarCtxCantidadPositiva, data.stkContext, services)
                        Case Regla.LoteObligatorio
                            ProcessServer.ExecuteTask(Of StockContext)(AddressOf ValidarCtxLoteUbicacionObligatorios, data.stkContext, services)                            'ProcessServer.ExecuteTask(Of StockContext)(AddressOf ValidarCtxLote, data.stkContext, services)
                        Case Regla.StockLoteNegativo
                            ProcessServer.ExecuteTask(Of StockContext)(AddressOf ValidarCtxLoteUbicacionObligatorios, data.stkContext, services)
                            If Not data.stkContext.Cancel Then
                                ProcessServer.ExecuteTask(Of StockContext)(AddressOf ValidarCtxStockLoteNegativo, data.stkContext, services)
                            End If
                        Case Regla.LoteBloqueado
                            ProcessServer.ExecuteTask(Of StockContext)(AddressOf ValidarCtxLoteUbicacionObligatorios, data.stkContext, services)
                            If Not data.stkContext.Cancel Then
                                ProcessServer.ExecuteTask(Of StockContext)(AddressOf ValidarCtxLoteBloqueado, data.stkContext, services)
                            End If
                        Case Regla.SerieObligatoria
                            ProcessServer.ExecuteTask(Of StockContext)(AddressOf ValidarCtxSerieObligatoria, data.stkContext, services)
                        Case Regla.SerieUnica
                            ProcessServer.ExecuteTask(Of StockContext)(AddressOf ValidarCtxSerieUnica, data.stkContext, services)
                        Case Regla.AlbaranVenta
                            ProcessServer.ExecuteTask(Of StockContext)(AddressOf ValidarCtxSalidaAlbaranVenta, data.stkContext, services)
                        Case Regla.AlmacenActivo
                            ProcessServer.ExecuteTask(Of StockContext)(AddressOf ValidarCtxAlmacenActivo, data.stkContext, services)
                        Case Regla.ArticuloDePortes
                            ProcessServer.ExecuteTask(Of StockContext)(AddressOf ValidarCtxMovtoArticuloPortes, data.stkContext, services)
                        Case Regla.FechaAjuste
                            ProcessServer.ExecuteTask(Of StockContext)(AddressOf ValidarCtxFechaAjuste, data.stkContext, services)
                        Case Regla.StockSerieNegativo
                            ProcessServer.ExecuteTask(Of StockContext)(AddressOf ValidarCtxStockSerieNegativo, data.stkContext, services)
                    End Select
                End If
            Next
        End If
    End Sub

    <Task()> Public Shared Sub ValidarCtxNumeroMovimiento(ByVal ctx As StockContext, ByVal services As ServiceProvider)
        If ctx Is Nothing Then Exit Sub
        If ctx.NumeroMovimiento < 0 Then
            ctx.Cancel = True
            ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(1), services)
        End If
    End Sub
    <Task()> Public Shared Sub ValidarCtxCantidadCero(ByVal ctx As StockContext, ByVal services As ServiceProvider)
        If ctx Is Nothing Then Exit Sub
        Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, ctx.Articulo, services)
        If ctx.CantidadConSigno = 0 OrElse (SegundaUnidad AndAlso Nz(ctx.CantidadConSigno2, 0) = 0) Then
            ctx.Cancel = True
            Dim datMsg As New DataMessage(16)
            datMsg.SegundaUnidad = (SegundaUnidad AndAlso Nz(ctx.CantidadConSigno2, 0) = 0)
            ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, datMsg, services)
        End If

    End Sub
    <Task()> Public Shared Sub ValidarCtxFechaUltimoCierre(ByVal ctx As StockContext, ByVal services As ServiceProvider)
        If ctx Is Nothing Then Exit Sub
        If ctx.FechaDocumento <= ctx.FechaUltimoCierre Then
            ctx.Cancel = True
            ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(4), services)
        End If
    End Sub
    <Task()> Public Shared Sub ValidarCtxStockNegativo(ByVal ctx As StockContext, ByVal services As ServiceProvider)
        '//Comprobar la gestion de stocks negativos, gestión de stocks negativos por artículo//
        'SIGNIFICADO DE LOS PARAMETROS:
        '=0 NO SE PERMITEN STOCKS NEGATIVOS
        '=1 SI SE PERMITEN
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(ctx.Articulo)
        Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, ctx.Articulo, services)
        If ArtInfo.GestionStockPorLotes Then
            If ctx.StockFisico < 0 OrElse ctx.StockFisicoLote < 0 OrElse (SegundaUnidad AndAlso ctx.StockFisicoLote2 < 0) Then
                ctx.Cancel = True
                ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(17, ctx.Articulo), services)
            End If
        Else
            If (ctx.StockFisico < 0 OrElse (SegundaUnidad AndAlso Nz(ctx.StockFisico2, 0) < 0)) And _
               ((ctx.ClaseMovimiento = enumtpmTipoMovimiento.tpmOutput) Or _
                (ctx.ClaseMovimiento = enumtpmTipoMovimiento.tpmInput And ctx.CantidadConSigno <= 0) Or _
                 ctx.TipoMovimiento = enumTipoMovimiento.tmInventario Or ctx.TipoMovimiento = enumTipoMovimiento.tmEntAjuste Or ctx.TipoMovimiento = enumTipoMovimiento.tmSalAjuste) Then
                Dim AppParamsStocks As ParametroStocks = services.GetService(Of ParametroStocks)()
                If Not AppParamsStocks.GestionStockNegativo Then
                    If Not AppParamsStocks.GestionStockNegativoPorArticulo Then
                        ctx.Cancel = True
                        Dim datMsg As New DataMessage(29, ctx.Articulo)
                        datMsg.SegundaUnidad = (SegundaUnidad AndAlso Nz(ctx.StockFisico2, 0) < 0)
                        ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, datMsg, services)
                    Else
                        If Not ArtInfo.StockNegativo Then
                            ctx.Cancel = True
                            Dim datMsg As New DataMessage(7, ctx.Articulo)
                            datMsg.SegundaUnidad = (SegundaUnidad AndAlso Nz(ctx.StockFisico2, 0) < 0)
                            ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, datMsg, services)
                        End If
                    End If
                End If
            End If
        End If
    End Sub
    <Task()> Public Shared Sub ValidarCtxFechaUltimoInventario(ByVal ctx As StockContext, ByVal services As ServiceProvider)
        If ctx.FechaDocumento < ctx.FechaUltimoInventario Then
            ctx.Cancel = True
            ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(5), services)
        End If
    End Sub
    <Task()> Public Shared Sub ValidarCtxCantidadPositiva(ByVal ctx As StockContext, ByVal services As ServiceProvider)
        Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, ctx.Articulo, services)
        If ctx.CantidadConSigno < 0 OrElse (SegundaUnidad AndAlso Nz(ctx.CantidadConSigno2, 0) < 0) Then
            ctx.Cancel = True
            Dim datMsg As New DataMessage(8)
            datMsg.SegundaUnidad = (SegundaUnidad AndAlso Nz(ctx.CantidadConSigno2, 0) < 0)
            ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, datMsg, services)
        End If
    End Sub
    <Task()> Public Shared Sub ValidarCtxLoteUbicacionObligatorios(ByVal ctx As StockContext, ByVal services As ServiceProvider)
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(ctx.Articulo)
        If ArtInfo.GestionStock AndAlso ArtInfo.GestionStockPorLotes Then
            If Len(ctx.Lote) = 0 And ctx.CantidadConSigno <> 0 Then
                ctx.Cancel = True
                ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(11), services)
            ElseIf Len(ctx.Ubicacion) = 0 And ctx.CantidadConSigno <> 0 Then
                ctx.Cancel = True
                ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(19), services)
            End If
        End If
    End Sub
    <Task()> Public Shared Sub ValidarCtxStockLoteNegativo(ByVal ctx As StockContext, ByVal services As ServiceProvider)
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(ctx.Articulo)
        Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, ctx.Articulo, services)
        If ArtInfo.GestionStock AndAlso ArtInfo.GestionStockPorLotes Then
            If ctx.StockFisico < 0 OrElse ctx.StockFisicoLote < 0 OrElse (SegundaUnidad AndAlso ctx.StockFisicoLote2 < 0) Then
                ctx.Cancel = True
                ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(17, ctx.Articulo), services)
            End If
        End If
    End Sub
    <Task()> Public Shared Sub ValidarCtxLoteBloqueado(ByVal ctx As StockContext, ByVal services As ServiceProvider)
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(ctx.Articulo)
        If ArtInfo.GestionStock AndAlso ArtInfo.GestionStockPorLotes Then
            If ctx.CantidadConSigno < 0 And ctx.LoteBloqueado Then
                ctx.Cancel = True
                ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(18), services)
            End If
        End If
    End Sub
    <Task()> Public Shared Sub ValidarCtxSerieObligatoria(ByVal ctx As StockContext, ByVal services As ServiceProvider)
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(ctx.Articulo)
        If ArtInfo.GestionPorNumeroSerie Then
            If Len(ctx.NSerie) = 0 Then
                ctx.Cancel = True
                ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(31), services)
            ElseIf Len(ctx.EstadoNSerie) = 0 Then
                ctx.Cancel = True
                ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(32), services)
            ElseIf Len(ctx.Operario) = 0 Then
                ctx.Cancel = True
                ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(33), services)
            End If
        End If
    End Sub
    <Task()> Public Shared Sub ValidarCtxSerieUnica(ByVal ctx As StockContext, ByVal services As ServiceProvider)
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(ctx.Articulo)
        If ArtInfo.GestionPorNumeroSerie Then
            If Len(ctx.NSerie) > 0 Then
                Dim serie As DataTable = New ArticuloNSerie().SelOnPrimaryKey(ctx.NSerie)
                If serie.Rows.Count > 0 Then
                    ctx.Cancel = True
                    ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(34), services)
                End If
            End If
        End If
    End Sub
    <Task()> Public Shared Sub ValidarCtxSalidaAlbaranVenta(ByVal ctx As StockContext, ByVal services As ServiceProvider)
        If ctx.TipoMovimiento = enumTipoMovimiento.tmSalAlbaranVenta Then
            ProcessServer.ExecuteTask(Of StockContext)(AddressOf ValidarCtxAlmacenBloqueado, ctx, services)
        End If
    End Sub
    <Task()> Public Shared Sub ValidarCtxAlmacenBloqueado(ByVal ctx As StockContext, ByVal services As ServiceProvider)
        Dim Almacenes As EntityInfoCache(Of AlmacenInfo) = services.GetService(Of EntityInfoCache(Of AlmacenInfo))()
        Dim AlmInfo As AlmacenInfo = Almacenes.GetEntity(ctx.Almacen)
        If AlmInfo.Bloqueado Then
            ctx.Cancel = True
            ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(15), services)
        End If
    End Sub
    <Task()> Public Shared Sub ValidarCtxAlmacenActivo(ByVal ctx As StockContext, ByVal services As ServiceProvider)
        Dim Almacenes As EntityInfoCache(Of AlmacenInfo) = services.GetService(Of EntityInfoCache(Of AlmacenInfo))()
        Dim AlmInfo As AlmacenInfo = Almacenes.GetEntity(ctx.Almacen)
        If Not AlmInfo.Activo Then
            ctx.Cancel = True
            ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(28), services)
        End If
    End Sub
    <Task()> Public Shared Sub ValidarCtxMovtoArticuloPortes(ByVal ctx As StockContext, ByVal services As ServiceProvider)
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(ctx.Articulo)
        If ArtInfo.ArticuloDePortes And ctx.TipoMovimiento <> enumTipoMovimiento.tmSalAlbaranVenta Then
            ctx.Cancel = True
            ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(38), services)
        End If
    End Sub
    <Task()> Public Shared Sub ValidarCtxFechaAjuste(ByVal ctx As StockContext, ByVal services As ServiceProvider)
        If ctx.TipoMovimiento = enumTipoMovimiento.tmEntAjuste Or ctx.TipoMovimiento = enumTipoMovimiento.tmSalAjuste Then
            Dim AppParamsStock As ParametroStocks = services.GetService(Of ParametroStocks)()
            If AppParamsStock.TipoInventario = TipoInventario.UltimoMovimiento Then
                If ctx.FechaDocumento = ctx.FechaUltimoInventario Then
                    ctx.Cancel = True
                    ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(46), services)
                End If
            End If
        End If
    End Sub
    <Task()> Public Shared Sub ValidarCtxStockSerieNegativo(ByVal ctx As StockContext, ByVal services As ServiceProvider)
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(ctx.Articulo)
        If ArtInfo.GestionPorNumeroSerie Then
            If Len(ctx.NSerie) = 0 And ctx.CantidadConSigno <> 0 Then
                ctx.Cancel = True
                ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(31), services)
            Else
                If ctx.StockFisico < 0 Then
                    ctx.Estado = EstadoStock.NoActualizado
                    ctx.Cancel = True
                    ctx.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(47, ctx.Articulo), services)
                End If
            End If
        End If
    End Sub

#End Region

    <Task()> Public Shared Function GetMovimiento(ByVal IDLineaMovimiento As Integer, ByVal services As ServiceProvider) As DataRow
        If IDLineaMovimiento <> 0 Then
            Dim aux As DataTable = New BE.DataEngine().Filter(cnEntidad, New NumberFilterItem("IDLineaMovimiento", IDLineaMovimiento))
            If Not aux Is Nothing AndAlso aux.Rows.Count > 0 Then
                Return aux.Rows(0)
            End If
        End If
    End Function

    <Task()> Public Shared Sub SetMessageMovimientoActualizado(ByVal data As DataMovimiento, ByVal services As ServiceProvider)
        data.stkData.Context.Estado = EstadoStock.Actualizado
        If Not data.Movimiento Is Nothing AndAlso Length(data.Movimiento("IDLineaMovimiento")) > 0 Then data.stkData.Context.IDLineaMovimiento = data.Movimiento("IDLineaMovimiento")
        data.stkData.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(10), services)
    End Sub

    <Task()> Public Shared Function NuevoNumeroMovimiento(ByVal data As Object, ByVal services As ServiceProvider) As Integer
        Dim N As Integer = -1
        Dim AppParamsSTK As ParametroStocks = services.GetService(Of ParametroStocks)()
        Dim IDContador As String = AppParamsSTK.ContadorHistMovimientoPredeterminado
        If Len(IDContador) = 0 Then
            ApplicationService.GenerateError("El contador de movimientos no existe o no está correctamente configurado.")
        Else
            Dim Contadores As EntityInfoCache(Of ContadorInfo) = services.GetService(Of EntityInfoCache(Of ContadorInfo))()
            Dim ContInfo As ContadorInfo = Contadores.GetEntity(IDContador)
            If Not ContInfo Is Nothing AndAlso Length(ContInfo.IDContador) > 0 Then
                If Not CBool(ContInfo.Numerico) Then
                    ApplicationService.GenerateError("El contador de movimientos debe ser configurado como contador numérico.")
                Else
                    N = CInt(ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, IDContador, services))
                End If
            Else
                ApplicationService.GenerateError("El contador de movimientos no existe o no está correctamente configurado.")
            End If
        End If

        If N > 0 Then
            Return N
        Else
            'Se deja porque bajaUpdateData.Detalle ya viene traduccido
            Throw New Exception(ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(42, Quoted(data.Context.NSerie)), services))
            Throw New Exception(ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(9), services))
        End If
    End Function

    <Task()> Public Shared Sub NuevoMovimiento(ByVal data As StockData, ByVal services As ServiceProvider)
        '//Generacion del movimiento
        If Not data.Context.EsCorreccion And data.Context.GenerarMovimiento Then
            Dim movimiento As DataRow = data.Context.Movimientos.NewRow
            movimiento("IDLineaMovimiento") = AdminData.GetAutoNumeric
            movimiento("IDMovimiento") = data.Context.NumeroMovimiento
            movimiento("IDArticulo") = data.Articulo
            movimiento("IDAlmacen") = data.Almacen

            Dim dtArticulo As DataTable = New Articulo().SelOnPrimaryKey(data.Articulo)
            If dtArticulo.Rows(0)("GestionStockPorLotes") Then
                movimiento("Lote") = data.Lote
                movimiento("Ubicacion") = data.Ubicacion
            ElseIf dtArticulo.Rows(0)("NSerieObligatorio") Then
                movimiento("Lote") = data.NSerie
                movimiento("Ubicacion") = data.Ubicacion
                If Length(data.Context.IDEstadoActivoAnterior) > 0 Then movimiento("IDEstadoActivo") = data.Context.IDEstadoActivoAnterior
            End If
            movimiento("IDTipoMovimiento") = data.Context.TipoMovimiento
            movimiento("FechaDocumento") = data.FechaDocumento
            movimiento("FechaMovimiento") = Today

            '//////////////////////////////////////////
            Dim datPrecio As New DataGetPrecioMovimiento(data, movimiento)
            datPrecio = ProcessServer.ExecuteTask(Of DataGetPrecioMovimiento, DataGetPrecioMovimiento)(AddressOf GetPrecioMovimiento, datPrecio, services)

            data.PrecioA = datPrecio.PrecioOutputA
            data.PrecioB = datPrecio.PrecioOutputB
            '//////////////////////////////////////////


            Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Articulo, services)
            Dim dataCant As New DataSetValoresMovimiento(movimiento, data)
            ProcessServer.ExecuteTask(Of DataSetValoresMovimiento)(AddressOf SetCantidadMovimiento, dataCant, services)
            Dim datAcum As New DataAcumulado(data, movimiento("IDLineaMovimiento"), SegundaUnidad)
            Select Case data.Context.TipoMovimiento
                Case enumTipoMovimiento.tmInventario
                    Dim datValorAcum As DataValorAcumulado = ProcessServer.ExecuteTask(Of DataAcumulado, DataValorAcumulado)(AddressOf ProcesoStocks.AcumuladoInventario, datAcum, services)
                    movimiento("Acumulado") = datValorAcum.Valor
                    If SegundaUnidad Then
                        If Not datValorAcum.Valor2 Is Nothing Then movimiento("Acumulado2") = datValorAcum.Valor2
                    End If
                Case Else
                    Dim datValorAcum As DataValorAcumulado = ProcessServer.ExecuteTask(Of DataAcumulado, DataValorAcumulado)(AddressOf ProcesoStocks.Acumulado, datAcum, services)
                    movimiento("Acumulado") = datValorAcum.Valor
                    If SegundaUnidad Then
                        If Not datValorAcum.Valor2 Is Nothing Then movimiento("Acumulado2") = datValorAcum.Valor2
                    End If
            End Select

            movimiento("PrecioA") = data.PrecioA
            movimiento("PrecioB") = data.PrecioB

            movimiento("Documento") = IIf(Len(data.Documento) > 0, data.Documento, DBNull.Value)
            movimiento("Texto") = IIf(Len(data.Texto) > 0, data.Texto, DBNull.Value)
            movimiento("IDOperario") = IIf(Length(data.Operario) > 0, data.Operario, DBNull.Value)
            movimiento("IDObra") = IIf(Length(data.Obra) > 0, data.Obra, DBNull.Value)
            movimiento("Traza") = IIf(data.Traza.Equals(Guid.Empty), DBNull.Value, data.Traza)
            movimiento("IDDocumento") = IIf(data.IDDocumento > 0, data.IDDocumento, DBNull.Value)

            Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
            Dim MonInfoA As MonedaInfo = Monedas.MonedaA
            Dim MonInfoB As MonedaInfo = Monedas.MonedaB

            Dim DataCalc As New DataCalcValPrecioMedio(movimiento, data.Context.ClaseMovimiento)
            If data.Context.TipoMovimiento = enumTipoMovimiento.tmInventario Then
                movimiento("PrecioMedio") = xRound(Nz(ProcessServer.ExecuteTask(Of DataCalcValPrecioMedio, Double)(AddressOf CalculoValAlmPrecioMedioInventario, DataCalc, services), 0), MonInfoA.NDecimalesPrecio)
                '//No subimos estos Precios A y B a donde se calculan, xq despues del inventario, no vamos a tener movimientos posteriores.
                If dtArticulo.Rows(0)("CriterioValoracion") = enumtaValoracion.taPrecioMedio Then
                    movimiento("PrecioA") = xRound(movimiento("PrecioMedio"), MonInfoA.NDecimalesPrecio)
                    movimiento("PrecioB") = xRound(movimiento("PrecioMedio") * Monedas.MonedaB.CambioB, MonInfoB.NDecimalesPrecio)
                End If
            Else
                movimiento("PrecioMedio") = xRound(Nz(ProcessServer.ExecuteTask(Of DataCalcValPrecioMedio, Double)(AddressOf CalculoValAlmPrecioMedio, DataCalc, services), 0), MonInfoA.NDecimalesPrecio)
            End If

            movimiento("PrecioUltimaCompra") = Nz(dtArticulo.Rows(0)("PrecioUltimaCompraA"), 0)
            movimiento("PrecioEstandar") = Nz(dtArticulo.Rows(0)("PrecioEstandarA"), 0) / Nz(dtArticulo.Rows(0)("UDValoracion"), 1)

            'El FIFO SE CALCULA DESPUES DE ACTUALIZAR (COMENTADO)
            ''Dim datosValFIFO As New ProcesoStocks.DataValoracionFIFO(data.Articulo, data.Almacen, movimiento("Acumulado"), movimiento("Acumulado"), data.FechaDocumento, enumstkValoracionFIFO.stkVFOrdenarPorFecha) 'por fecha documento
            ''Dim valoracion As ValoracionPreciosInfo
            ''valoracion = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ProcesoStocks.ValoracionFIFO, datosValFIFO, services)
            ''If Not valoracion Is Nothing Then
            ''    movimiento("FifoFD") = valoracion.PrecioA
            ''End If
            ''datosValFIFO.Orden = enumstkValoracionFIFO.stkVFOrdenarPorMvto 'por fecha idLineamovimiento
            ''valoracion = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ProcesoStocks.ValoracionFIFO, datosValFIFO, services)
            ''If Not valoracion Is Nothing Then
            ''    movimiento("FifoF") = valoracion.PrecioA
            ''End If

            movimiento("FifoFD") = 0
            movimiento("FifoF") = 0
            If (Length(data.PrecintaNSerie) > 0) Then
                movimiento("SeriePrecinta") = data.PrecintaNSerie
                movimiento("NDesdePrecinta") = data.PrecintaDesde
                movimiento("NHastaPrecinta") = data.PrecintaHasta
                movimiento("NDesdePrecintaUtilizada") = IIf(Nz(data.PrecintaUtilizadaDesde, 0) = 0, data.PrecintaDesde, data.PrecintaUtilizadaDesde)
                movimiento("NHastaPrecintaUtilizada") = IIf(Nz(data.PrecintaUtilizadaHasta, 0) = 0, data.PrecintaHasta, data.PrecintaUtilizadaHasta)
            End If

            data.Context.Movimientos.Rows.Add(movimiento)

            ProcessServer.ExecuteTask(Of DataMovimiento)(AddressOf SetMessageMovimientoActualizado, New DataMovimiento(movimiento, data), services)
        ElseIf Not data.Context.GenerarMovimiento AndAlso data.Context.TipoMovimiento = enumTipoMovimiento.tmInventario Then
            ProcessServer.ExecuteTask(Of DataMovimiento)(AddressOf SetMessageMovimientoActualizado, New DataMovimiento(Nothing, data), services)
        End If
    End Sub

    Public Class DataSetValoresMovimiento
        Public Movimiento As DataRow
        Public stkData As StockData
        Public Sub New(ByVal movimiento As DataRow, ByVal stkData As StockData)
            Me.Movimiento = movimiento
            Me.stkData = stkData
        End Sub
    End Class
    <Task()> Public Shared Sub SetCantidadMovimiento(ByVal data As DataSetValoresMovimiento, ByVal services As ServiceProvider)
        Select Case data.stkData.Context.TipoMovimiento
            Case enumTipoMovimiento.tmEntAjuste, enumTipoMovimiento.tmSalAjuste     '//Ajuste
                data.Movimiento("Cantidad") = data.stkData.Context.CantidadConSigno
            Case enumTipoMovimiento.tmInventario                                    '//Inventario
                Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.stkData.Context.Articulo)
                If ArtInfo.GestionPorNumeroSerie Then
                    If data.stkData.Context.PropiedadesEstado.Disponible Then
                        data.Movimiento("Cantidad") = data.stkData.Cantidad
                    Else
                        data.Movimiento("Cantidad") = 0
                    End If
                Else
                    data.Movimiento("Cantidad") = data.stkData.Cantidad
                End If
            Case enumTipoMovimiento.tmEntAlbaranCompra, enumTipoMovimiento.tmEntFabrica, enumTipoMovimiento.tmEntSubcontratacion, enumTipoMovimiento.tmEntTransferencia
                data.Movimiento("Cantidad") = data.stkData.Cantidad
            Case enumTipoMovimiento.tmSalAlbaranVenta, enumTipoMovimiento.tmSalFabrica, enumTipoMovimiento.tmSalSubcontratacion, enumTipoMovimiento.tmSalTransferencia, enumTipoMovimiento.tmSalRealquiler
                data.Movimiento("Cantidad") = -1 * data.stkData.Cantidad
            Case Else
                If data.stkData.Context.ClaseMovimiento = enumtpmTipoMovimiento.tpmInput Then
                    data.Movimiento("Cantidad") = data.stkData.Cantidad
                ElseIf data.stkData.Context.ClaseMovimiento = enumtpmTipoMovimiento.tpmOutput Then
                    data.Movimiento("Cantidad") = -1 * data.stkData.Cantidad
                End If
        End Select

        If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.stkData.Articulo, services) Then
            Select Case data.stkData.Context.TipoMovimiento
                Case enumTipoMovimiento.tmEntAjuste, enumTipoMovimiento.tmSalAjuste     '//Ajuste
                    data.Movimiento("Cantidad2") = data.stkData.Context.CantidadConSigno2
                Case enumTipoMovimiento.tmInventario                                    '//Inventario
                    data.Movimiento("Cantidad2") = data.stkData.Cantidad2
                Case enumTipoMovimiento.tmEntAlbaranCompra, enumTipoMovimiento.tmEntFabrica, enumTipoMovimiento.tmEntSubcontratacion, enumTipoMovimiento.tmEntTransferencia
                    data.Movimiento("Cantidad2") = data.stkData.Cantidad2
                Case enumTipoMovimiento.tmSalAlbaranVenta, enumTipoMovimiento.tmSalFabrica, enumTipoMovimiento.tmSalSubcontratacion, enumTipoMovimiento.tmSalTransferencia, enumTipoMovimiento.tmSalRealquiler
                    data.Movimiento("Cantidad2") = -1 * data.stkData.Cantidad2
                Case Else
                    If data.stkData.Context.ClaseMovimiento = enumtpmTipoMovimiento.tpmInput Then
                        data.Movimiento("Cantidad2") = data.stkData.Cantidad2
                    ElseIf data.stkData.Context.ClaseMovimiento = enumtpmTipoMovimiento.tpmOutput Then
                        data.Movimiento("Cantidad2") = -1 * data.stkData.Cantidad2
                    End If
            End Select
        End If
    End Sub
    '<Task()> Public Shared Sub SetPrecioMovimiento(ByVal data As DataSetValoresMovimiento, ByVal services As ServiceProvider)
    '    Select Case data.stkData.Context.ClaseMovimiento
    '        Case enumtpmTipoMovimiento.tpmOutput
    '            If data.stkData.PrecioA = 0 Or data.stkData.PrecioB = 0 Then
    '                Dim dataPrecioEnt As New DataPrecioMovimiento(data.stkData.Articulo, data.stkData.Almacen, data.stkData.FechaDocumento, Math.Abs(data.stkData.Cantidad), data.stkData.Context.ClaseMovimiento)
    '                Dim precios As Hashtable = ProcessServer.ExecuteTask(Of DataPrecioMovimiento, Hashtable)(AddressOf PrecioMovimiento, dataPrecioEnt, services)
    '                If Not precios Is Nothing Then
    '                    data.Movimiento("PrecioA") = precios("PrecioA")
    '                    data.Movimiento("PrecioB") = precios("PrecioB")
    '                End If
    '            End If
    '        Case enumtpmTipoMovimiento.tpmInput
    '            If data.stkData.PrecioA = 0 Or data.stkData.PrecioB = 0 Then
    '                Dim AppStkParams As ParametroStocks = services.GetService(Of ParametroStocks)()
    '                Dim PermitirPrecioCero As Boolean = AppStkParams.PrecioMovimientoCero()
    '                If Not PermitirPrecioCero Then
    '                    Dim dataPrecioEnt As New DataPrecioMovimiento(data.stkData.Articulo, data.stkData.Almacen, data.stkData.FechaDocumento, data.stkData.Cantidad, data.stkData.Context.ClaseMovimiento)
    '                    Dim precios As Hashtable = ProcessServer.ExecuteTask(Of DataPrecioMovimiento, Hashtable)(AddressOf PrecioMovimiento, dataPrecioEnt, services)
    '                    If Not precios Is Nothing Then
    '                        data.Movimiento("PrecioA") = precios("PrecioA")
    '                        data.Movimiento("PrecioB") = precios("PrecioB")
    '                    End If
    '                End If
    '            End If
    '        Case enumtpmTipoMovimiento.tpmInventario
    '            If data.stkData.PrecioA = 0 Or data.stkData.PrecioB = 0 Then
    '                Dim dataPrecioEnt As New DataPrecioMovimiento(data.stkData.Articulo, data.stkData.Almacen, data.stkData.FechaDocumento, data.stkData.Cantidad, data.stkData.Context.ClaseMovimiento)
    '                Dim precios As Hashtable = ProcessServer.ExecuteTask(Of DataPrecioMovimiento, Hashtable)(AddressOf PrecioMovimiento, dataPrecioEnt, services)
    '                If Not precios Is Nothing Then
    '                    data.Movimiento("PrecioA") = precios("PrecioA")
    '                    data.Movimiento("PrecioB") = precios("PrecioB")
    '                End If
    '            End If
    '    End Select
    'End Sub

    <Task()> Public Shared Function GetRegistroArticuloAlmacen(ByVal data As DataArticuloAlmacen, ByVal services As ServiceProvider) As DataTable
        If Len(data.IDArticulo) > 0 And Len(data.IDAlmacen) > 0 Then
            Dim insert As Boolean
            Dim artalm As New ArticuloAlmacen
            Dim dt As DataTable = artalm.SelOnPrimaryKey(data.IDArticulo, data.IDAlmacen)
            If Not dt Is Nothing Then
                If dt.Rows.Count = 0 Then
                    insert = True
                Else
                    If Length(dt.Rows(0)("IDArticuloGenerico")) > 0 Then
                        data.IDArticulo = dt.Rows(0)("IDArticuloGenerico")
                        dt = artalm.SelOnPrimaryKey(data.IDArticulo, data.IDAlmacen)
                        If Not dt Is Nothing Then
                            If dt.Rows.Count = 0 Then
                                insert = True
                            End If
                        End If
                    End If
                End If

                If insert Then
                    dt = artalm.AddNew
                    Dim newrow As DataRow = dt.NewRow
                    newrow("MarcaAuto") = AdminData.GetAutoNumeric()
                    newrow("IDArticulo") = data.IDArticulo
                    newrow("IDAlmacen") = data.IDAlmacen
                    newrow("StockFisico") = 0
                    If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.IDArticulo, services) Then
                        newrow("StockFisico2") = 0
                    End If
                    newrow("StockMedio") = 0
                    newrow("StockSeguridad") = 0
                    newrow("PuntoPedido") = 0
                    newrow("LoteMinimo") = 0
                    newrow("PrecioMedioA") = 0
                    newrow("PrecioMedioB") = 0
                    newrow("PrecioFIFOFechaA") = 0
                    newrow("PrecioFIFOFechaB") = 0
                    newrow("PrecioFIFOMvtoA") = 0
                    newrow("PrecioFIFOMvtoB") = 0
                    newrow("Rotacion") = 0
                    newrow("Inventariado") = 0
                    dt.Rows.Add(newrow)
                End If
            End If
            Return dt
        End If
    End Function

    <Task()> Public Shared Sub Actualizar(ByVal UpdateData As StockUpdateData, ByVal services As ServiceProvider)
        Try
            If UpdateData.Estado = EstadoStock.Actualizado Then
                Dim dts(6) As DataTable
                dts(0) = UpdateData.ArticuloAlmacen
                dts(1) = UpdateData.Movimientos
                If Not UpdateData.Acumulados Is Nothing Then
                    dts(2) = UpdateData.Acumulados
                End If
                If Not UpdateData.Lote Is Nothing Then
                    dts(3) = UpdateData.Lote
                End If
                If Not UpdateData.Activo Is Nothing Then
                    dts(4) = UpdateData.Activo
                End If
                If Not UpdateData.Serie Is Nothing Then
                    dts(5) = UpdateData.Serie
                End If
                If Not UpdateData.HistoricoEstadoActivo Is Nothing Then
                    dts(6) = UpdateData.HistoricoEstadoActivo
                End If
                For Each Dt As DataTable In dts
                    If Not Dt Is Nothing Then BusinessHelper.UpdateTable(Dt)
                Next
                Dim FifosCalculados As New List(Of Integer)
                If Not UpdateData.Movimientos Is Nothing AndAlso UpdateData.Movimientos.Rows.Count > 0 Then
                    UpdateData.Movimientos.AcceptChanges()
                    For Each Dr As DataRow In UpdateData.Movimientos.Select
                        Dim datosValFIFO As New ProcesoStocks.DataValoracionFIFO(Dr("IDArticulo"), Dr("IDAlmacen"), Dr("Acumulado"), Dr("Acumulado"), Dr("FechaDocumento"), enumstkValoracionFIFO.stkVFOrdenarPorFecha) 'por fecha documento
                        Dim valoracion As ValoracionPreciosInfo
                        valoracion = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ProcesoStocks.ValoracionFIFO, datosValFIFO, services)
                        If Not valoracion Is Nothing Then
                            Dr("FifoFD") = valoracion.PrecioA
                        End If
                        datosValFIFO.Orden = enumstkValoracionFIFO.stkVFOrdenarPorMvto 'por fecha idLineamovimiento
                        valoracion = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ProcesoStocks.ValoracionFIFO, datosValFIFO, services)
                        If Not valoracion Is Nothing Then
                            Dr("FifoF") = valoracion.PrecioA
                        End If
                        FifosCalculados.Add(Dr("IDLineaMovimiento"))
                    Next
                    BusinessHelper.UpdateTable(UpdateData.Movimientos)
                End If
                If Not UpdateData.Acumulados Is Nothing AndAlso UpdateData.Acumulados.Rows.Count > 0 Then
                    UpdateData.Acumulados.AcceptChanges()
                    For Each Dr As DataRow In UpdateData.Acumulados.Select
                        If Not FifosCalculados.Contains(Dr("IDLineaMovimiento")) Then
                            Dim datosValFIFO As New ProcesoStocks.DataValoracionFIFO(Dr("IDArticulo"), Dr("IDAlmacen"), Dr("Acumulado"), Dr("Acumulado"), Dr("FechaDocumento"), enumstkValoracionFIFO.stkVFOrdenarPorFecha) 'por fecha documento
                            Dim valoracion As ValoracionPreciosInfo
                            valoracion = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ProcesoStocks.ValoracionFIFO, datosValFIFO, services)
                            If Not valoracion Is Nothing Then
                                Dr("FifoFD") = valoracion.PrecioA
                            End If
                            datosValFIFO.Orden = enumstkValoracionFIFO.stkVFOrdenarPorMvto 'por fecha idLineamovimiento
                            valoracion = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ProcesoStocks.ValoracionFIFO, datosValFIFO, services)
                            If Not valoracion Is Nothing Then
                                Dr("FifoF") = valoracion.PrecioA
                            End If
                            FifosCalculados.Add(Dr("IDLineaMovimiento"))
                        End If
                    Next
                    BusinessHelper.UpdateTable(UpdateData.Acumulados)
                End If
            End If
        Catch ex As Exception
            UpdateData.Detalle = "No se ha podido actualizar el Stock. Debe actualizarlo Manualmente."
            UpdateData.Detalle &= vbNewLine & ex.Message
            UpdateData.Estado = EstadoStock.NoActualizado
            UpdateData.Log = "Error en Actualización de Stock"
        End Try
    End Sub


    ''' <summary>
    ''' Método que prepara la información de salida.
    ''' </summary>
    ''' <param name="data">Objeto StockData con la información necesaria para los procesos de Stock.</param>
    ''' <param name="services">Objeto con información comprtida a lo largo de los procesos.</param>
    ''' <returns>Objeto con la información recogida a lo largo del proceso de la gestión de stocks correspondiente.</returns>
    ''' <remarks>Método que a partir de la información de entrada, junto con la información de contexto reunidad a lo largo del proceso, retorna un objeto con toda la información requerida.</remarks>
    <Task()> Public Shared Function GetStockUpdateData(ByVal data As StockData, ByVal services As ServiceProvider) As StockUpdateData
        Dim updateData As New StockUpdateData
        updateData.StockData = data
        updateData.Estado = EstadoStock.NoActualizado

        '//Coger la informacion del contexto de actualizacion si existe
        If Not data.Context Is Nothing Then
            updateData.Estado = data.Context.Estado
            updateData.IDLineaMovimiento = data.Context.IDLineaMovimiento
            updateData.NumeroMovimiento = data.Context.NumeroMovimiento
            updateData.CantidadMovimiento = data.Context.CantidadConSigno
            updateData.CantidadMovimiento2 = data.Context.CantidadConSigno2
            updateData.Detalle = data.Context.Detalle
            updateData.Log = data.Context.Log
            If updateData.Estado = EstadoStock.Actualizado Then
                updateData.ArticuloAlmacen = data.Context.ArticuloAlmacen
                updateData.Movimientos = data.Context.Movimientos
                If Not data.Context.Acumulados Is Nothing Then
                    updateData.Acumulados = data.Context.Acumulados
                End If
                If Not data.Context.LoteBBDD Is Nothing Then
                    updateData.Lote = data.Context.LoteBBDD
                End If
                If Not data.Context.SerieBBDD Is Nothing Then
                    updateData.Serie = data.Context.SerieBBDD
                End If
                If Not data.Context.ActivoBBDD Is Nothing Then
                    updateData.Activo = data.Context.ActivoBBDD
                End If
                If Not data.Context.HistoricoEstadoActivo Is Nothing Then
                    updateData.HistoricoEstadoActivo = data.Context.HistoricoEstadoActivo
                End If
            End If
        End If
        Return updateData
    End Function

    Public Function EstablecerReglas(ByVal ParamArray Rules() As Regla) As Regla()
        Dim aRules(-1) As Regla

        For Each r As Regla In Rules
            ReDim Preserve aRules(aRules.Length)
            aRules(aRules.Length - 1) = r
        Next
        Return aRules
    End Function

    <Task()> Public Shared Function ConvertirMovimientoAStockData(ByVal movimiento As DataRow, ByVal services As ServiceProvider) As StockData
        Dim data As New StockData(movimiento("IDArticulo"), movimiento("IDAlmacen"), movimiento("Cantidad"), movimiento("PrecioA"), movimiento("PrecioB"), _
                                  movimiento("FechaDocumento"), movimiento("IDTipoMovimiento"))
        If Length(movimiento("Lote")) > 0 Then
            data.Lote = movimiento("Lote")
            data.NSerie = movimiento("Lote")
            data.EstadoNSerieAnterior = movimiento("IDEstadoActivo") & String.Empty
        End If
        If Length(movimiento("Ubicacion")) > 0 Then
            data.Ubicacion = movimiento("Ubicacion")
        End If
        If Length(movimiento("Documento")) > 0 Then
            data.Documento = movimiento("Documento")
        End If
        If Not IsDBNull(movimiento("Traza")) Then
            data.Traza = movimiento("Traza")
        End If
        If movimiento.Table.Columns.Contains("Cantidad2") AndAlso Length(movimiento("Cantidad2")) > 0 Then
            data.Cantidad2 = CDbl(movimiento("Cantidad2"))
        End If
        Return data
    End Function

    <Task()> Public Shared Sub SetCtxFechaUltimoCierre(ByVal data As StockData, ByVal services As ServiceProvider)
        Dim dt As DataTable = AdminData.GetData("vNegCierreInventarioFechaUltimoCierre", , "TOP 1 FechaHasta", "FechaHasta DESC, FechaDesde DESC")
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            data.Context.FechaUltimoCierre = dt.Rows(0)("FechaHasta")
        End If
    End Sub

    <Task()> Public Shared Sub SetCtxArticuloAlmacen(ByVal data As StockData, ByVal services As ServiceProvider)
        ''//Si tenemos doble unidad, haremos que si el stockfisico o el stockfisico2 se queda a 0, el otro tb se quedará a 0.
        Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Articulo, services)
        'If data.Context.StockFisico = 0 OrElse (SegundaUnidad AndAlso data.Context.StockFisico2 = 0) Then
        '    If SegundaUnidad Then
        '        data.Context.StockFisico = 0
        '        data.Context.StockFisico2 = 0
        '    End If
        'End If

        Dim ArticuloAlmacen As DataTable = data.Context.ArticuloAlmacen
        ArticuloAlmacen.Rows(0)("StockFisico") = data.Context.StockFisico
        If SegundaUnidad Then
            ArticuloAlmacen.Rows(0)("StockFisico2") = data.Context.StockFisico2
        End If

        Select Case data.Context.TipoMovimiento
            Case enumTipoMovimiento.tmEntAjuste, enumTipoMovimiento.tmSalAjuste
                ArticuloAlmacen.Rows(0)("FechaUltimoAjuste") = data.FechaDocumento
            Case enumTipoMovimiento.tmInventario
                ArticuloAlmacen.Rows(0)("FechaUltimoInventario") = data.FechaDocumento
                ArticuloAlmacen.Rows(0)("Inventariado") = True
        End Select

        If data.Context.FechaUltimoMovimiento <= data.FechaDocumento Then
            ArticuloAlmacen.Rows(0)("FechaUltimoMovimiento") = data.FechaDocumento
        End If

    End Sub

    <Task()> Public Shared Sub SetCandidadIntegracionBodega(ByVal data As StockData, ByVal services As ServiceProvider)
        '//Actualizar la cantidad a pasar en los casos de integración
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.Articulo)
        If Len(ArtInfo.EnsambladoStock) > 0 And Len(ArtInfo.ClaseStock) > 0 Then
            data.Cantidad += data.Context.QIntermedia
            If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.Articulo, services) Then
                data.Cantidad2 += data.Context.QIntermedia2
            End If
        End If
    End Sub

#End Region

#Region " Acumulados "

    <Serializable()> _
    Public Class DataAcumulado
        Public IDLineaActual As Integer
        Public stkData As StockData
        Public SegundaUnidad As Boolean

        Public Sub New(ByVal stkData As StockData, ByVal IDLineaActual As Integer, Optional ByVal SegundaUnidad As Boolean = False)
            Me.stkData = stkData
            Me.IDLineaActual = IDLineaActual
            Me.SegundaUnidad = SegundaUnidad
        End Sub
    End Class

    <Serializable()> _
   Public Class DataValorAcumulado
        Public Valor As Double
        Public Valor2 As Double?
    End Class

    <Task()> Public Shared Function Acumulado(ByVal data As DataAcumulado, ByVal services As ServiceProvider) As DataValorAcumulado
        '//1.Asigna el acumulado al movimiento actual
        '//2.Recalcula el acumulado de los movimientos posteriores si es necesario

        '//'valor' es el valor del acumulado para el movimiento que se esta tratando en este momento
        Dim valor As Double = data.stkData.Context.StockFisico
        Dim valorprecio As Double = data.stkData.Context.PrecioA
        Dim Valor2 As Double?
        If data.SegundaUnidad Then
            Valor2 = CDbl(data.stkData.Context.StockFisico2)
        End If
        If data.stkData.FechaDocumento <= data.stkData.Context.FechaUltimoMovimiento OrElse _
         (data.stkData.ContextCorrect.CorreccionEnCantidad OrElse data.stkData.ContextCorrect.CorreccionEnCantidad2) OrElse _
          data.stkData.ContextCorrect.CorreccionEnPrecio Then
            '//los valores que se calculan ahora son respecto a la FechaDocumento del movimiento en cuestion
            '//(estos valores pueden no ser los mismos que los del contexto - ctx)

            '//TODO ESTE CALCULO ES DEBIDO A QUE EL MOVIMIENTO DE INVENTARIO, EN LA FECHADOCUMENTO CORRESPONDIENTE
            '//A DICHO INVENTARIO, NO SE SITUA SEGUN SU IDLINEAMOVIMIENTO, SI NO QUE POR CONVENIO DE LA APLICACION
            '//"LO SITUAMOS COMO" PRIMER O ULTIMO MOVIMIENTO DE ESE DIA (SEGUN EL PARAMETRO 'STTIPOINV')

            '//obtencion del acumulado del ultimo movimiento
            Dim ultimoAcumulado As Double : Dim ultimoAcumulado2 As Double?
            Dim ultimoPrecioMedio As Double
            Dim fechaInventario As Date
            Dim datMovtoAnt As New DataObtenerMovtoAnterior(data.stkData.Articulo, data.stkData.Almacen, data.stkData.Context.FechaDocumento, data.IDLineaActual)
            Dim movimientoAnterior As DataRow = ProcessServer.ExecuteTask(Of DataObtenerMovtoAnterior, DataRow)(AddressOf ObtenerMovimientoAnterior, datMovtoAnt, services)
            If Not movimientoAnterior Is Nothing Then
                ultimoAcumulado = movimientoAnterior("Acumulado")
                ultimoPrecioMedio = movimientoAnterior("PrecioMedio")
                If data.SegundaUnidad Then ultimoAcumulado2 = CDbl(Nz(movimientoAnterior("Acumulado2"), 0))
            End If

            '//Respecto a la fecha de documento de este movimiento puede haber uno o varios movimientos de inventario posteriores.
            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", data.stkData.Articulo))
            f.Add(New StringFilterItem("IDAlmacen", data.stkData.Almacen))
            Dim AppParamsStock As ParametroStocks = services.GetService(Of ParametroStocks)()
            Select Case AppParamsStock.TipoInventario
                Case TipoInventario.PrimerMovimiento
                    f.Add(New NumberFilterItem("IDTipoMovimiento", enumTipoMovimiento.tmInventario))
                    f.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThan, data.stkData.Context.FechaDocumento))
                Case TipoInventario.UltimoMovimiento
                    f.Add(New NumberFilterItem("IDTipoMovimiento", enumTipoMovimiento.tmInventario))
                    f.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThanOrEqual, data.stkData.Context.FechaDocumento))
            End Select

            Dim inventario As DataTable = New BE.DataEngine().Filter(cnEntidad, f, "TOP 1 *", "FechaDocumento")

            Dim f1 As New Filter(FilterUnionOperator.Or)
            Dim f2 As New Filter
            Dim f3 As New Filter
            Dim f4 As New Filter
            f2.Add(New StringFilterItem("IDArticulo", data.stkData.Articulo))
            f2.Add(New StringFilterItem("IDAlmacen", data.stkData.Almacen))
            f3.Add(New StringFilterItem("IDArticulo", data.stkData.Articulo))
            f3.Add(New StringFilterItem("IDAlmacen", data.stkData.Almacen))
            f4.Add(New StringFilterItem("IDArticulo", data.stkData.Articulo))
            f4.Add(New StringFilterItem("IDAlmacen", data.stkData.Almacen))
            If inventario.Rows.Count > 0 Then
                fechaInventario = inventario.Rows(0)("FechaDocumento")
                Select Case AppParamsStock.TipoInventario
                    Case TipoInventario.PrimerMovimiento
                        f2.Add(New DateFilterItem("FechaDocumento", FilterOperator.Equal, data.stkData.Context.FechaDocumento))
                        f2.Add(New NumberFilterItem("IDLineaMovimiento", FilterOperator.GreaterThan, data.IDLineaActual))

                        f3.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThan, data.stkData.Context.FechaDocumento))
                        f3.Add(New DateFilterItem("FechaDocumento", FilterOperator.LessThan, fechaInventario))
                    Case TipoInventario.UltimoMovimiento
                        If data.stkData.Context.FechaDocumento = fechaInventario Then
                            f2.Add(New DateFilterItem("FechaDocumento", FilterOperator.Equal, data.stkData.Context.FechaDocumento))
                            f2.Add(New NumberFilterItem("IDLineaMovimiento", FilterOperator.GreaterThan, data.IDLineaActual))
                        Else
                            f2.Add(New DateFilterItem("FechaDocumento", FilterOperator.Equal, data.stkData.Context.FechaDocumento))
                            f2.Add(New NumberFilterItem("IDLineaMovimiento", FilterOperator.GreaterThan, data.IDLineaActual))

                            f3.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThan, data.stkData.Context.FechaDocumento))
                            f3.Add(New DateFilterItem("FechaDocumento", FilterOperator.LessThan, fechaInventario))

                            f4.Add(New DateFilterItem("FechaDocumento", FilterOperator.Equal, fechaInventario))
                        End If
                End Select
            Else
                '//Si no hay movimientos de inventario por delante
                f2.Add(New DateFilterItem("FechaDocumento", FilterOperator.Equal, data.stkData.Context.FechaDocumento))
                f2.Add(New NumberFilterItem("IDLineaMovimiento", FilterOperator.GreaterThan, data.IDLineaActual))
                f3.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThan, data.stkData.Context.FechaDocumento))
            End If

            If f2.Count > 2 Then
                f1.Add(f2)
            End If
            If f3.Count > 2 Then
                f1.Add(f3)
            End If
            If f4.Count > 2 Then
                f1.Add(f4)
            End If

            valor = ultimoAcumulado + data.stkData.Context.CantidadConSigno
            If data.stkData.Context.ClaseMovimiento = enumtpmTipoMovimiento.tpmInput Then
                If valor > 0 Then
                    If ultimoAcumulado > 0 Then
                        valorprecio = ((ultimoPrecioMedio * ultimoAcumulado) + (data.stkData.Context.CantidadConSigno * data.stkData.Context.PrecioA)) / valor
                    Else
                        valorprecio = data.stkData.Context.PrecioA
                    End If
                Else
                    valorprecio = 0
                End If
            ElseIf data.stkData.Context.ClaseMovimiento = enumtpmTipoMovimiento.tpmOutput Then
                valorprecio = ultimoPrecioMedio
            End If
            If data.SegundaUnidad Then Valor2 = Nz(ultimoAcumulado2, 0) + data.stkData.Context.CantidadConSigno2
            Dim anterior As Double = valor : Dim anterior2 As Double = Nz(Valor2, 0)
            Dim anteriorprecio As Double = valorprecio
            Dim Acumulados As DataTable = AdminData.GetEntityData(cnMyClass, f1)

            Dim TipoMov As EntityInfoCache(Of TipoMovimientoInfo) = services.GetService(Of EntityInfoCache(Of TipoMovimientoInfo))()
            If Not Acumulados Is Nothing AndAlso Acumulados.Rows.Count > 0 Then
                Dim ProcInfo As ArticuloCosteEstandar.ProcInfoActualizarPrecioEstandar = services.GetService(Of ArticuloCosteEstandar.ProcInfoActualizarPrecioEstandar)()
                Dim blnRecalcularMovtosPosteriores As Boolean = ProcInfo.RecalcularPrecioStdPosteriores
                Dim MovtosRecalcular As List(Of DataRow) = (From c In Acumulados _
                                                                Where c("IDTipoMovimiento") <> enumTipoMovimiento.tmCorreccion AndAlso _
                                                                      c("IDTipoMovimiento") <> enumTipoMovimiento.tmInventario AndAlso _
                                                                       c("IDLineaMovimiento") <> data.IDLineaActual _
                                                                Order By c("FechaDocumento"), c("IDLineaMovimiento") _
                                                                Select c).ToList
                ' Acumulados.DefaultView.Sort = "FechaDocumento,IDLineaMovimiento"
                If Not MovtosRecalcular Is Nothing AndAlso MovtosRecalcular.Count > 0 Then
                    For Each movimiento As DataRow In MovtosRecalcular
                        'If movimiento("IDTipoMovimiento") <> enumTipoMovimiento.tmCorreccion _
                        'And movimiento("IDTipoMovimiento") <> enumTipoMovimiento.tmInventario Then
                        ' If movimiento("IDLineaMovimiento") <> data.IDLineaActual Then
                        If (data.stkData.FechaDocumento <= data.stkData.Context.FechaUltimoMovimiento) OrElse data.stkData.ContextCorrect.CorreccionEnCantidad Then
                            movimiento("Acumulado") = anterior + movimiento("Cantidad")
                        End If
                        Dim TipoMovInfo As TipoMovimientoInfo = TipoMov.GetEntity(movimiento("IDTipoMovimiento"))
                        If TipoMovInfo.ClaseMovimiento = enumtpmTipoMovimiento.tpmInput Then
                            If movimiento("Acumulado") > 0 Then
                                movimiento("PrecioMedio") = ((anteriorprecio * anterior) + (movimiento("Cantidad") * movimiento("PrecioA"))) / movimiento("Acumulado")
                            Else
                                movimiento("PrecioMedio") = 0
                            End If
                        ElseIf TipoMovInfo.ClaseMovimiento = enumtpmTipoMovimiento.tpmOutput Then
                            movimiento("PrecioMedio") = anteriorprecio
                        End If

                        Dim datMovto As New DataRecalcularPrecioEstandarMovto(data.stkData, movimiento)
                        datMovto.RecalcularPrecioStdPosteriores = blnRecalcularMovtosPosteriores
                        If blnRecalcularMovtosPosteriores Then
                            datMovto.PrecioEstandar = movimiento("PrecioEstandar")
                            datMovto = ProcessServer.ExecuteTask(Of DataRecalcularPrecioEstandarMovto, DataRecalcularPrecioEstandarMovto)(AddressOf RecalcularPrecioEstandarMovto, datMovto, services)
                            movimiento("PrecioEstandar") = datMovto.PrecioEstandar
                            blnRecalcularMovtosPosteriores = datMovto.RecalcularPrecioStdPosteriores
                        End If

                        'El FIFO SE CALCULA DESPUES DE ACTUALIZAR (COMENTADO)
                        'Dim FilArtAlm As New Filter
                        'FilArtAlm.Add(New StringFilterItem("IDArticulo", movimiento("IDArticulo")))
                        'FilArtAlm.Add(New StringFilterItem("IDAlmacen", movimiento("IDAlmacen")))

                        'Dim dtArticuloAlmacen As DataTable = New ArticuloAlmacen().Filter(FilArtAlm)
                        'Dim datosValFIFO As New ProcesoStocks.DataValoracionFIFO(movimiento("IDArticulo"), movimiento("IDAlmacen"), dtArticuloAlmacen.Rows(0)("StockFisico"), dtArticuloAlmacen.Rows(0)("StockFisico"), movimiento("FechaDocumento"), enumstkValoracionFIFO.stkVFOrdenarPorFecha)
                        'Dim valoracion As ValoracionPreciosInfo
                        'valoracion = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ProcesoStocks.ValoracionFIFO, datosValFIFO, services)
                        'If Not valoracion Is Nothing Then
                        '    movimiento("FifoF") = valoracion.PrecioA
                        'End If
                        'datosValFIFO.Orden = enumstkValoracionFIFO.stkVFOrdenarPorMvto
                        'valoracion = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ProcesoStocks.ValoracionFIFO, datosValFIFO, services)
                        'If Not valoracion Is Nothing Then
                        '    movimiento("FifoFD") = valoracion.PrecioA
                        'End If


                        '  ultimoAcumulado = movimiento("Acumulado")
                        anterior = movimiento("Acumulado")
                        anteriorprecio = movimiento("PrecioMedio")
                        If data.SegundaUnidad Then
                            If (data.stkData.FechaDocumento <= data.stkData.Context.FechaUltimoMovimiento) OrElse data.stkData.ContextCorrect.CorreccionEnCantidad2 Then
                                movimiento("Acumulado2") = anterior2 + Nz(movimiento("Cantidad2"), 0)
                            End If
                            anterior2 = Nz(movimiento("Acumulado2"), 0)
                        End If
                        ' End If
                        'End If
                    Next
                End If

                data.stkData.Context.Acumulados = Acumulados
            End If
        End If

        Dim result As New DataValorAcumulado
        result.Valor = valor
        If data.SegundaUnidad AndAlso Not Valor2 Is Nothing Then result.Valor2 = CDbl(Valor2)
        Return result
    End Function


    Public Class DataRecalcularPrecioEstandarMovto
        Public movimiento As DataRow
        Public stkData As StockData

        Public RecalcularPrecioStdPosteriores As Boolean
        Public PrecioEstandar As Double
        Public Sub New(ByVal stkData As StockData, ByVal movimiento As DataRow)
            Me.stkData = stkData
            Me.movimiento = movimiento
        End Sub
    End Class
    <Task()> Public Shared Function RecalcularPrecioEstandarMovto(ByVal data As DataRecalcularPrecioEstandarMovto, ByVal services As ServiceProvider) As DataRecalcularPrecioEstandarMovto
        data.PrecioEstandar = data.movimiento("PrecioEstandar")
        If data.RecalcularPrecioStdPosteriores Then
            Dim AppParamsStock As ParametroStocks = services.GetService(Of ParametroStocks)()
            If AppParamsStock.TipoMovimientoCantidad0 <> 0 AndAlso AppParamsStock.TipoMovimientoCantidad0 > enumTipoMovimiento.tmSalRealquiler Then
                '//"Recalculamos" el precionEstadar entre movimientos del Tipo TipoMovimientoCantidad0
                If data.movimiento("IDTipoMovimiento") = AppParamsStock.TipoMovimientoCantidad0 Then
                    data.RecalcularPrecioStdPosteriores = False
                End If

                If data.RecalcularPrecioStdPosteriores AndAlso data.stkData.TipoMovimiento = AppParamsStock.TipoMovimientoCantidad0 Then
                    '//Le asignamos el precio del movimiento que estamos introduciendo de tipo TipoMovimientoCantidad0
                    data.PrecioEstandar = data.stkData.PrecioA
                End If
            End If
        End If

        Return data
    End Function

    <Task()> Public Shared Function AcumuladoInventario(ByVal data As DataAcumulado, ByVal services As ServiceProvider) As DataValorAcumulado
        '//1.Asigna el acumulado al movimiento actual de INVENTARIO
        '//2.Recalcula el acumulado de los movimientos posteriores si es necesario
        Dim mAcumuladoInfo() As AcumuladoInfo = services.GetService(Of AcumuladoInfo())()

        '//'valor' es el valor del acumulado para el movimiento que se esta tratando en este momento
        Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
        Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.stkData.Articulo)
        Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.stkData.Articulo, services)
        Dim valor As Double : Dim Valor2 As Double?
        Dim valorprecio As Double = data.stkData.Context.PrecioA
        If ArtInfo.GestionStockPorLotes Then
            '//por defecto
            Valor2 = Nothing
            valor = data.stkData.Cantidad + (data.stkData.Context.StockFisico - data.stkData.Context.StockFisicoLote - data.stkData.Context.QIntermediaRestoLotes)
            If SegundaUnidad Then Valor2 = CDbl(Nz(data.stkData.Cantidad2, 0) + (Nz(data.stkData.Context.StockFisico2, 0) - Nz(data.stkData.Context.StockFisicoLote2, 0) - data.stkData.Context.QIntermediaRestoLotes2))
            '//Este array solo se rellena en movimientos de inventario
            If Not mAcumuladoInfo Is Nothing AndAlso mAcumuladoInfo.Length > 0 Then
                For Each a As AcumuladoInfo In mAcumuladoInfo
                    If a.IDArticulo = data.stkData.Articulo And a.IDAlmacen = data.stkData.Almacen Then
                        valor = a.Acumulado
                        If SegundaUnidad AndAlso Not a.Acumulado2 Is Nothing Then Valor2 = CDbl(a.Acumulado2)
                    End If
                Next
            End If
        ElseIf ArtInfo.GestionPorNumeroSerie Then
            If Not mAcumuladoInfo Is Nothing AndAlso mAcumuladoInfo.Length > 0 Then
                For Each a As AcumuladoInfo In mAcumuladoInfo
                    If a.IDArticulo = data.stkData.Articulo And a.IDAlmacen = data.stkData.Almacen Then
                        valor = a.Acumulado
                    End If
                Next
            End If
        Else
            valor = data.stkData.Cantidad
            If data.SegundaUnidad Then
                Valor2 = CDbl(data.stkData.Cantidad2)
            End If
        End If

        '//obtencion del precio medio del mov anterior, para recalcular el precio medio del los movimientos siguientes al inventario.
        Dim datMovtoAnt As New DataObtenerMovtoAnterior(data.stkData.Articulo, data.stkData.Almacen, data.stkData.Context.FechaDocumento, data.IDLineaActual)
        Dim movimientoAnterior As DataRow = ProcessServer.ExecuteTask(Of DataObtenerMovtoAnterior, DataRow)(AddressOf ObtenerMovimientoAnterior, datMovtoAnt, services)
        If Not movimientoAnterior Is Nothing Then
            valorprecio = movimientoAnterior("PrecioMedio")
        End If

        If (data.stkData.FechaDocumento <= data.stkData.Context.FechaUltimoMovimiento) Then
            Dim AppParamsStocks As ParametroStocks = services.GetService(Of ParametroStocks)()
            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", data.stkData.Articulo))
            f.Add(New StringFilterItem("IDAlmacen", data.stkData.Almacen))

            'Para el calculo del precioFIFO
            Dim dtArticuloAlmacen As DataTable = New ArticuloAlmacen().Filter(f)

            Select Case AppParamsStocks.TipoInventario
                Case TipoInventario.PrimerMovimiento
                    f.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThanOrEqual, data.stkData.Context.FechaDocumento))
                Case TipoInventario.UltimoMovimiento
                    f.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThan, data.stkData.Context.FechaDocumento))
            End Select

            Dim anterior As Double = valor : Dim anterior2 As Double = Nz(Valor2, 0)
            Dim anteriorprecio As Double = valorprecio
            Dim TipoMov As EntityInfoCache(Of TipoMovimientoInfo) = services.GetService(Of EntityInfoCache(Of TipoMovimientoInfo))()
            Dim Acumulados As DataTable = AdminData.GetEntityData(cnMyClass, f)
            If Not Acumulados Is Nothing AndAlso Acumulados.Rows.Count > 0 Then
                'Acumulados.DefaultView.Sort = "FechaDocumento,IDLineaMovimiento"
                Dim MovtosRecalcular As List(Of DataRow) = (From c In Acumulados _
                                                               Where c("IDTipoMovimiento") <> enumTipoMovimiento.tmCorreccion AndAlso _
                                                                     c("IDTipoMovimiento") <> enumTipoMovimiento.tmInventario AndAlso _
                                                                      c("IDLineaMovimiento") <> data.IDLineaActual _
                                                               Order By c("FechaDocumento"), c("IDLineaMovimiento") _
                                                               Select c).ToList
                ' Acumulados.DefaultView.Sort = "FechaDocumento,IDLineaMovimiento"
                If Not MovtosRecalcular Is Nothing AndAlso MovtosRecalcular.Count > 0 Then
                    For Each movimiento As DataRow In MovtosRecalcular
                        'If movimiento("IDTipoMovimiento") <> enumTipoMovimiento.tmCorreccion _
                        'And movimiento("IDTipoMovimiento") <> enumTipoMovimiento.tmInventario Then
                        '    If movimiento("IDLineaMovimiento") <> data.IDLineaActual Then
                        movimiento("Acumulado") = anterior + movimiento("Cantidad")
                        Dim TipoMovInfo As TipoMovimientoInfo = TipoMov.GetEntity(movimiento("IDTipoMovimiento"))
                        If TipoMovInfo.ClaseMovimiento = enumtpmTipoMovimiento.tpmInput Then
                            If movimiento("Acumulado") > 0 Then
                                movimiento("PrecioMedio") = ((anteriorprecio * anterior) + (movimiento("Cantidad") * movimiento("PrecioA"))) / movimiento("Acumulado")
                            Else
                                movimiento("PrecioMedio") = 0
                            End If
                        ElseIf TipoMovInfo.ClaseMovimiento = enumtpmTipoMovimiento.tpmOutput Then
                            movimiento("PrecioMedio") = anteriorprecio
                        End If

                        Dim datosValFIFO As New ProcesoStocks.DataValoracionFIFO(movimiento("IDArticulo"), movimiento("IDAlmacen"), dtArticuloAlmacen.Rows(0)("StockFisico"), dtArticuloAlmacen.Rows(0)("StockFisico"), movimiento("FechaDocumento"), enumstkValoracionFIFO.stkVFOrdenarPorFecha)
                        Dim valoracion As ValoracionPreciosInfo
                        valoracion = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ProcesoStocks.ValoracionFIFO, datosValFIFO, services)
                        If Not valoracion Is Nothing Then
                            movimiento("FifoF") = valoracion.PrecioA
                        End If
                        datosValFIFO.Orden = enumstkValoracionFIFO.stkVFOrdenarPorMvto
                        valoracion = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ProcesoStocks.ValoracionFIFO, datosValFIFO, services)
                        If Not valoracion Is Nothing Then
                            movimiento("FifoFD") = valoracion.PrecioA
                        End If

                        anterior = movimiento("Acumulado")
                        anteriorprecio = movimiento("PrecioMedio")
                        If data.SegundaUnidad Then
                            movimiento("Acumulado2") = anterior2 + Nz(movimiento("Cantidad2"), 0)
                            anterior2 = Nz(movimiento("Acumulado2"), 0)
                        End If
                        '                End If
                        'End If
                    Next
                End If

                data.stkData.Context.Acumulados = Acumulados
            End If
        End If

        Dim result As New DataValorAcumulado
        result.Valor = valor
        If data.SegundaUnidad AndAlso Not Valor2 Is Nothing Then result.Valor2 = CDbl(Valor2)
        Return result
    End Function

    <Task()> Public Shared Function AcumuladoCorreccionEnFecha(ByVal data As DataAcumulado, ByVal services As ServiceProvider) As DataValorAcumulado
        Dim valor As Double = data.stkData.Context.StockFisico
        Dim valorPrecio As Double = data.stkData.Context.PrecioA
        Dim ultimoPrecioMedio As Double

        Dim SegundaUnidad As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, data.stkData.Articulo, services)
        Dim valor2 As Double?
        If SegundaUnidad Then valor2 = data.stkData.Context.StockFisico2

        If data.stkData.ContextCorrect.CorreccionEnFecha Then
            Dim movimientoAnterior As DataRow
            Dim ultimoAcumulado As Double : Dim ultimoAcumulado2 As Double?
            Dim IDLineaUltimoMovimiento As Integer
            Dim FechaUltimoMovimiento As Date


            If data.stkData.FechaDocumento > data.stkData.Context.FechaDocumentoOriginal Then
                Dim datMovtoAnt As New DataObtenerMovtoAnterior(data.stkData.Articulo, data.stkData.Almacen, data.stkData.Context.FechaDocumentoOriginal, data.IDLineaActual)
                movimientoAnterior = ProcessServer.ExecuteTask(Of DataObtenerMovtoAnterior, DataRow)(AddressOf ObtenerMovimientoAnterior, datMovtoAnt, services)
            ElseIf data.stkData.FechaDocumento < data.stkData.Context.FechaDocumentoOriginal Then
                Dim datMovtoAnt As New DataObtenerMovtoAnterior(data.stkData.Articulo, data.stkData.Almacen, data.stkData.FechaDocumento, data.IDLineaActual)
                movimientoAnterior = ProcessServer.ExecuteTask(Of DataObtenerMovtoAnterior, DataRow)(AddressOf ObtenerMovimientoAnterior, datMovtoAnt, services)
            End If

            If Not movimientoAnterior Is Nothing Then
                ultimoAcumulado = movimientoAnterior("Acumulado")
                ultimoPrecioMedio = movimientoAnterior("PrecioMedio")
                If SegundaUnidad Then ultimoAcumulado2 = CDbl(Nz(movimientoAnterior("Acumulado2"), 0))
                FechaUltimoMovimiento = movimientoAnterior("FechaDocumento")
            End If

            valor = ultimoAcumulado + data.stkData.Context.CantidadConSigno
            If SegundaUnidad Then
                valor2 = ultimoAcumulado2 + data.stkData.Context.CantidadConSigno2
            End If
            '//Construir el criterio de seleccion para obtener los movimientos que hay que recalcular su acumulado            
            Dim fechaInicial As Date
            Dim fechaFinal As Date
            If data.stkData.FechaDocumento > data.stkData.Context.FechaDocumentoOriginal Then
                fechaInicial = data.stkData.Context.FechaDocumentoOriginal
                fechaFinal = data.stkData.FechaDocumento
            ElseIf data.stkData.FechaDocumento < data.stkData.Context.FechaDocumentoOriginal Then
                fechaInicial = data.stkData.FechaDocumento
                fechaFinal = data.stkData.Context.FechaDocumentoOriginal
            End If

            Dim Acumulados As DataTable
            Dim f1 As New Filter(FilterUnionOperator.Or)
            Dim f2 As New Filter
            Dim f3 As New Filter
            Dim f4 As New Filter
            If Not data.stkData.ContextCorrect.CorreccionEnCantidad Then
                '//Si solo hay correccion en fecha este filtro es suficiente (incluir el movimiento que se esta corrigiendo)
                f2.Add(New StringFilterItem("IDArticulo", data.stkData.Articulo))
                f2.Add(New StringFilterItem("IDAlmacen", data.stkData.Almacen))
                f2.Add(New DateFilterItem("FechaDocumento", fechaInicial))
                f2.Add(New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion))
                f2.Add(New NumberFilterItem("IDLineaMovimiento", FilterOperator.GreaterThanOrEqual, data.IDLineaActual))

                f3.Add(New StringFilterItem("IDArticulo", data.stkData.Articulo))
                f3.Add(New StringFilterItem("IDAlmacen", data.stkData.Almacen))
                f3.Add(New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion))
                f3.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThan, fechaInicial))
                f3.Add(New DateFilterItem("FechaDocumento", FilterOperator.LessThan, fechaFinal))

                f4.Add(New StringFilterItem("IDArticulo", data.stkData.Articulo))
                f4.Add(New StringFilterItem("IDAlmacen", data.stkData.Almacen))
                f4.Add(New DateFilterItem("FechaDocumento", fechaFinal))
                f4.Add(New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion))
                f4.Add(New NumberFilterItem("IDLineaMovimiento", FilterOperator.LessThanOrEqual, data.IDLineaActual))

                f1.Add(f2)
                f1.Add(f3)
                f1.Add(f4)
            Else
                '//Hay que seleccionar un conjunto de registros similar al del calculo del acumulado normal,
                '//la diferencia principal es que hay que incluir el movimiento que se esta corrigiendo
                Dim f As New Filter
                f.Add(New StringFilterItem("IDArticulo", data.stkData.Articulo))
                f.Add(New StringFilterItem("IDAlmacen", data.stkData.Almacen))
                Dim AppParamsStock As ParametroStocks = services.GetService(Of ParametroStocks)()
                Select Case AppParamsStock.TipoInventario
                    Case TipoInventario.PrimerMovimiento
                        f.Add(New NumberFilterItem("IDTipoMovimiento", enumTipoMovimiento.tmInventario))
                        f.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThan, fechaInicial))
                    Case TipoInventario.UltimoMovimiento
                        f.Add(New NumberFilterItem("IDTipoMovimiento", enumTipoMovimiento.tmInventario))
                        f.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThanOrEqual, fechaInicial))
                End Select
                Dim inventario As DataTable = New BE.DataEngine().Filter(cnEntidad, f, "top 1 *", "FechaDocumento")

                f2.Add(New StringFilterItem("IDArticulo", data.stkData.Articulo))
                f2.Add(New StringFilterItem("IDAlmacen", data.stkData.Almacen))
                f2.Add(New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion))

                f3.Add(New StringFilterItem("IDArticulo", data.stkData.Articulo))
                f3.Add(New StringFilterItem("IDAlmacen", data.stkData.Almacen))
                f3.Add(New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion))

                f4.Add(New StringFilterItem("IDArticulo", data.stkData.Articulo))
                f4.Add(New StringFilterItem("IDAlmacen", data.stkData.Almacen))
                f4.Add(New NumberFilterItem("IDTipoMovimiento", FilterOperator.NotEqual, enumTipoMovimiento.tmCorreccion))
                If inventario.Rows.Count > 0 Then
                    Dim fechaInventario As Date = inventario.Rows(0)("FechaDocumento")
                    Select Case AppParamsStock.TipoInventario
                        Case TipoInventario.PrimerMovimiento
                            f2.Add(New DateFilterItem("FechaDocumento", FilterOperator.Equal, fechaInicial))
                            f2.Add(New NumberFilterItem("IDLineaMovimiento", FilterOperator.GreaterThanOrEqual, data.IDLineaActual))

                            f3.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThan, fechaInicial))
                            f3.Add(New DateFilterItem("FechaDocumento", FilterOperator.LessThan, fechaInventario))
                        Case TipoInventario.UltimoMovimiento
                            f2.Add(New DateFilterItem("FechaDocumento", FilterOperator.Equal, fechaInicial))
                            f2.Add(New NumberFilterItem("IDLineaMovimiento", FilterOperator.GreaterThanOrEqual, data.IDLineaActual))

                            f3.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThan, fechaInicial))
                            f3.Add(New DateFilterItem("FechaDocumento", FilterOperator.LessThan, fechaInventario))

                            f4.Add(New DateFilterItem("FechaDocumento", FilterOperator.Equal, fechaInventario))
                    End Select
                Else
                    '//Si no hay movimientos de inventario por delante
                    f2.Add(New DateFilterItem("FechaDocumento", FilterOperator.Equal, fechaInicial))
                    f2.Add(New NumberFilterItem("IDLineaMovimiento", FilterOperator.GreaterThanOrEqual, data.IDLineaActual))
                    f3.Add(New DateFilterItem("FechaDocumento", FilterOperator.GreaterThan, fechaInicial))
                End If

                If f2.Count > 2 Then
                    f1.Add(f2)
                End If
                If f3.Count > 2 Then
                    f1.Add(f3)
                End If
                If f4.Count > 2 Then
                    f1.Add(f4)
                End If
            End If


            If data.stkData.Context.ClaseMovimiento = enumtpmTipoMovimiento.tpmInput Then
                If valor > 0 Then
                    If ultimoAcumulado > 0 Then
                        valorPrecio = ultimoPrecioMedio
                    Else
                        valorPrecio = data.stkData.Context.PrecioA
                    End If
                Else
                    valorPrecio = 0
                End If
            ElseIf data.stkData.Context.ClaseMovimiento = enumtpmTipoMovimiento.tpmOutput Then
                valorPrecio = ultimoPrecioMedio
            End If

            Dim anterior As Double = ultimoAcumulado
            Dim anterior2 As Double?
            If SegundaUnidad Then anterior2 = ultimoAcumulado2
            Dim anteriorprecio As Double = valorPrecio
            Acumulados = AdminData.GetEntityData(cnMyClass, f1)
            If Not Acumulados Is Nothing AndAlso Acumulados.Rows.Count > 0 Then
                'Acumulados.DefaultView.Sort = "IDLineaMovimiento"
                'Dim i As Integer = Acumulados.DefaultView.Find(data.IDLineaActual)
                'If i >= 0 Then
                '    Acumulados.DefaultView(i)("FechaDocumento") = data.stkData.FechaDocumento
                'End If

                Dim MovtoActual As List(Of DataRow) = (From c In Acumulados _
                                                                Where c("IDLineaMovimiento") = data.IDLineaActual _
                                                                Order By c("FechaDocumento"), c("IDLineaMovimiento") _
                                                                Select c).ToList
                If Not MovtoActual Is Nothing AndAlso MovtoActual.Count > 0 Then
                    '//Le cambiamos la fecha momentaneamente, para que se tengan en cuenta sus valores
                    MovtoActual(0)("FechaDocumento") = data.stkData.FechaDocumento
                End If

                ' Acumulados.DefaultView.Sort = "FechaDocumento,IDLineaMovimiento"
                Dim MovtosRecalcularAcum As List(Of DataRow) = (From c In Acumulados _
                                                                Where c("IDTipoMovimiento") <> enumTipoMovimiento.tmCorreccion AndAlso _
                                                                    c("IDTipoMovimiento") <> enumTipoMovimiento.tmInventario _
                                                                Order By c("FechaDocumento"), c("IDLineaMovimiento") _
                                                                Select c).ToList
                If Not MovtosRecalcularAcum Is Nothing AndAlso MovtosRecalcularAcum.Count > 0 Then
                    For Each movimiento As DataRow In MovtosRecalcularAcum
                        'If movimiento("IDTipoMovimiento") <> enumTipoMovimiento.tmCorreccion _
                        'And movimiento("IDTipoMovimiento") <> enumTipoMovimiento.tmInventario Then
                        If movimiento("IDLineaMovimiento") = data.IDLineaActual Then

                            valor = anterior + data.stkData.Context.CantidadConSigno

                            If SegundaUnidad Then
                                valor2 = anterior2 + data.stkData.Context.CantidadConSigno2
                                anterior2 = valor2
                            End If

                            Dim TipoMov As EntityInfoCache(Of TipoMovimientoInfo) = services.GetService(Of EntityInfoCache(Of TipoMovimientoInfo))()
                            Dim TipoMovInfo As TipoMovimientoInfo = TipoMov.GetEntity(movimiento("IDTipoMovimiento"))
                            If TipoMovInfo.ClaseMovimiento = enumtpmTipoMovimiento.tpmInput Then
                                If valor > 0 Then
                                    anteriorprecio = ((anteriorprecio * anterior) + (data.stkData.Context.CantidadConSigno * data.stkData.Context.PrecioA)) / valor
                                End If
                            End If

                            anterior = valor

                            '//Cancelar las modificaciones para evitar un error de concurrencia
                            movimiento.RejectChanges()
                            ' movimiento("PrecioMedio") = anteriorprecio

                        Else
                            movimiento("Acumulado") = anterior + movimiento("Cantidad")

                            If SegundaUnidad Then
                                movimiento("Acumulado2") = anterior2 + Nz(movimiento("Cantidad2"), 0)
                                anterior2 = CDbl(Nz(movimiento("Acumulado2"), 0))
                            End If

                            Dim TipoMov As EntityInfoCache(Of TipoMovimientoInfo) = services.GetService(Of EntityInfoCache(Of TipoMovimientoInfo))()
                            Dim TipoMovInfo As TipoMovimientoInfo = TipoMov.GetEntity(movimiento("IDTipoMovimiento"))
                            If TipoMovInfo.ClaseMovimiento = enumtpmTipoMovimiento.tpmInput Then
                                If movimiento("Acumulado") > 0 Then
                                    If anterior > 0 Then
                                        movimiento("PrecioMedio") = ((anteriorprecio * anterior) + (movimiento("Cantidad") * movimiento("PrecioA"))) / movimiento("Acumulado")
                                    Else
                                        movimiento("PrecioMedio") = movimiento("PrecioA")
                                    End If

                                Else
                                    movimiento("PrecioMedio") = 0
                                End If
                            ElseIf TipoMovInfo.ClaseMovimiento = enumtpmTipoMovimiento.tpmOutput Then
                                movimiento("PrecioMedio") = anteriorprecio
                            End If

                            anterior = movimiento("Acumulado")
                            anteriorprecio = movimiento("PrecioMedio")
                        End If
                        'End If
                    Next

                    data.stkData.Context.Acumulados = Acumulados
                End If
            End If

        End If

        Dim result As New DataValorAcumulado
        result.Valor = valor
        If data.SegundaUnidad AndAlso Not valor2 Is Nothing Then result.Valor2 = CDbl(valor2)
        Return result
    End Function

#End Region

#Region " Integración con Bodega "

    <Serializable()> _
    Public Class DataVinoQ
        Public IDArticulo As String
        Public IDDeposito As String
        Public Lote As String
        Public IDAlmacen As String

        Public IDVino As Guid
        Public Cantidad As Double

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDArticulo As String, ByVal IDDeposito As String, ByVal Lote As String, ByVal IDAlmacen As String)
            Me.IDArticulo = IDArticulo
            Me.IDDeposito = IDDeposito
            Me.Lote = Lote
            Me.IDAlmacen = IDAlmacen
        End Sub

    End Class

    Public Enum enumTipoSincronizacion
        Entrada
        Salida
        Ajuste
        Inventario
        EliminarMovimiento
        Correccion
        EntradaTransferencia
    End Enum

    <Serializable()> _
    Public Class DataIntegracionConBodega
        Public DatosSincronizacion As DataNumeroMovimientoSinc
        Public TipoSincronizacion As enumTipoSincronizacion
        Public stkDataOriginal As StockData
        Public stkDataSalida As StockData
        Public stkDataUpdateSalida As StockUpdateData
        Public stkDataUpdateEntrada As StockUpdateData
        Public NumLinMovimientoBorrado As Integer

        Public Sub New(ByVal DatosSincronizacion As DataNumeroMovimientoSinc, ByVal TipoSincronizacion As enumTipoSincronizacion, _
                       Optional ByVal stkDataOriginal As StockData = Nothing, Optional ByVal stkDataSalida As StockData = Nothing, _
                       Optional ByVal stkDataUpdateSalida As StockUpdateData = Nothing, Optional ByVal stkDataUpdateEntrada As StockUpdateData = Nothing, _
                       Optional ByVal NumLinMovimientoBorrado As Integer = 0)
            Me.DatosSincronizacion = DatosSincronizacion
            Me.TipoSincronizacion = TipoSincronizacion
            Me.stkDataOriginal = stkDataOriginal
            Me.stkDataSalida = stkDataSalida
            Me.stkDataUpdateSalida = stkDataUpdateSalida
            Me.stkDataUpdateEntrada = stkDataUpdateEntrada
            Me.NumLinMovimientoBorrado = NumLinMovimientoBorrado
        End Sub
    End Class

    <Task()> Public Shared Function IntegracionConBodega(ByVal data As DataIntegracionConBodega, ByVal services As ServiceProvider) As StockUpdateData
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If Not AppParams.GestionBodegas Then Exit Function
        Dim ud As StockUpdateData
        If data.DatosSincronizacion.Sinc Then
            Try
                Dim AlmacenBodega As Boolean = ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf EsAlmacenDeBodega, data.DatosSincronizacion.stkData.Almacen, services)
                Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
                Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data.DatosSincronizacion.stkData.Articulo)
                If Len(ArtInfo.EnsambladoStock) > 0 AndAlso Len(ArtInfo.ClaseStock) > 0 AndAlso AlmacenBodega Then
                    Dim datIStock As New DataCreateIStockClass(ArtInfo.EnsambladoStock, ArtInfo.ClaseStock)
                    Dim IStockClass As IStock = ProcessServer.ExecuteTask(Of DataCreateIStockClass, IStock)(AddressOf CreateIStockClass, datIStock, services)
                    If Not IStockClass Is Nothing Then
                        Select Case data.TipoSincronizacion
                            Case enumTipoSincronizacion.Inventario, enumTipoSincronizacion.Ajuste, enumTipoSincronizacion.Salida, _
                                    enumTipoSincronizacion.Correccion, enumTipoSincronizacion.EliminarMovimiento
                                ud = data.stkDataUpdateSalida
                            Case enumTipoSincronizacion.Entrada, enumTipoSincronizacion.EntradaTransferencia
                                ud = data.stkDataUpdateEntrada
                        End Select

                        'Comprobar que no existan movimientos de inventario posteriores a la fecha del movimiento que se pretende hacer,
                        'ya que en ese caso se actualiza la cantidad de Bodega, se hace el movimiento en almacén pero no se actualiza el Stock Físico descuadrándose los datos.
                        'En los propios Inventarios no se puede controlar ya que con el primer lote de un Artículo-Almacén ya cambia la Fecha de Último Inventario. Tienen
                        'su propio control en la pantalla de Inventarios.
                        If data.TipoSincronizacion <> enumTipoSincronizacion.Inventario Then
                            ProcessServer.ExecuteTask(Of StockData)(AddressOf ValidarInventarioPosterior, ud.StockData, services)
                        End If

                        'Comprobar que no existan Operaciones de Bodega posteriores a la fecha del movimiento que se pretende hacer,
                        'ya que si se intentara borrar con posterioridad esas Operaciones se descuadra el Stock.
                        If Not ud.StockData.Context.Cancel Then
                            ProcessServer.ExecuteTask(Of StockData)(AddressOf ValidarOperacionPosterior, ud.StockData, services)
                        End If

                        If Not ud.StockData.Context.Cancel Then
                            Select Case data.TipoSincronizacion
                                Case enumTipoSincronizacion.Inventario
                                    Dim vinoQ As DataVinoQ = IStockClass.SincronizarInventario(data.DatosSincronizacion.NumeroMovimiento, data.DatosSincronizacion.stkData)
                                    Return ActualizarTrazaUpdateData(vinoQ, data.stkDataUpdateSalida)
                                Case enumTipoSincronizacion.Ajuste
                                    Dim vinoQ As DataVinoQ = IStockClass.SincronizarAjuste(data.DatosSincronizacion.NumeroMovimiento, data.DatosSincronizacion.stkData)
                                    Return ActualizarTrazaUpdateData(vinoQ, data.stkDataUpdateSalida)
                                Case enumTipoSincronizacion.Entrada
                                    Dim vinoQ As DataVinoQ = IStockClass.SincronizarEntrada(data.DatosSincronizacion.NumeroMovimiento, data.DatosSincronizacion.stkData)
                                    Return ActualizarTrazaUpdateData(vinoQ, data.stkDataUpdateEntrada)
                                Case enumTipoSincronizacion.Salida
                                    Dim vinoQ As DataVinoQ = IStockClass.SincronizarSalida(data.DatosSincronizacion.NumeroMovimiento, data.DatosSincronizacion.stkData)
                                    Return ActualizarTrazaUpdateData(vinoQ, data.stkDataUpdateSalida)
                                Case enumTipoSincronizacion.Correccion
                                    Dim vinoQ As DataVinoQ = IStockClass.SincronizarCorreccionMovimiento(data.DatosSincronizacion.NumeroMovimiento, data.DatosSincronizacion.stkData, data.stkDataOriginal)
                                    Return ActualizarTrazaUpdateData(vinoQ, data.stkDataUpdateSalida)
                                Case enumTipoSincronizacion.EliminarMovimiento
                                    Dim vinoQ As DataVinoQ = IStockClass.SincronizarEliminarMovimiento(data.DatosSincronizacion.NumeroMovimiento, data.NumLinMovimientoBorrado, data.stkDataOriginal)
                                    Return ActualizarTrazaUpdateData(vinoQ, data.stkDataUpdateSalida)
                                Case enumTipoSincronizacion.EntradaTransferencia
                                    Dim vinoQ As DataVinoQ = IStockClass.SincronizarEntradaTransferencia(data.DatosSincronizacion.NumeroMovimiento, data.DatosSincronizacion.stkData, data.stkDataSalida, data.stkDataUpdateEntrada, data.stkDataUpdateSalida)
                                    Return ActualizarTrazaUpdateData(vinoQ, data.stkDataUpdateEntrada)
                            End Select
                        Else
                            If Not ud Is Nothing Then
                                ud.Detalle = ud.StockData.Context.Detalle
                                ud.Estado = EstadoStock.NoActualizado
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                If Not ud Is Nothing Then
                    ud.Detalle = ex.Message
                    ud.Estado = EstadoStock.NoActualizado
                Else
                    Throw New Exception(ex.Message)
                End If
            End Try
        End If
        Return ud
    End Function

    Private Shared Function ActualizarTrazaUpdateData(ByVal vinoQ As DataVinoQ, ByVal stkUpdateData As StockUpdateData) As StockUpdateData
        If Not vinoQ Is Nothing AndAlso Not stkUpdateData Is Nothing AndAlso Not stkUpdateData.Movimientos Is Nothing Then
            Dim MovimientoVino As List(Of DataRow) = (From c In CType(stkUpdateData.Movimientos, DataTable) _
                                                            Where Not c.IsNull("Lote") AndAlso _
                                                                  Not c.IsNull("Ubicacion") AndAlso _
                                                                    c("IDArticulo") = vinoQ.IDArticulo AndAlso _
                                                                    c("IDAlmacen") = vinoQ.IDAlmacen AndAlso _
                                                                    c("Lote") = vinoQ.Lote AndAlso _
                                                                    c("Ubicacion") = vinoQ.IDDeposito).ToList()
            If Not MovimientoVino Is Nothing AndAlso MovimientoVino.Count > 0 Then
                For Each dr As DataRow In MovimientoVino
                    dr("Traza") = vinoQ.IDVino
                Next
            End If
            stkUpdateData.StockData.Traza = vinoQ.IDVino
        End If
        Return stkUpdateData
    End Function

    <Task()> Public Shared Function EsAlmacenDeBodega(ByVal IDAlmacen As String, ByVal services As ServiceProvider) As Boolean
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If Not AppParams.GestionBodegas Then Exit Function
        '// Si hacemos un Movimiento de Stock de un artículo, se comprueba que la tipología del articulo tiene los ensamblados de bodega.
        '// Esto  hace que, además de actualizar la cantidad o movimiento del Almacén, también se hace la actualización de los depósitos 
        '// de bodega que tienen este articulo Lote. 
        '// Cuando el almacén no pertenece a la bodega, el sistema intenta acceder a la actualización de bodega no encontrando dicho artículo.
        '// Se debe de Validar que el Artículo además de tener los ensamblados de bodega, deben tener que el almacén no sea uno de los que 
        '// tenemos en tbBdgNave o en el parámetro ALM_PREDET, que son los almacenes que se guardan en Bodega.
        Dim OC As BusinessHelper
        OC = BusinessHelper.CreateBusinessObject("BdgNave")
        Dim dt As DataTable = OC.Filter(New StringFilterItem("IDAlmacen", IDAlmacen))
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 OrElse IDAlmacen = AppParams.Almacen Then
            Return True
        Else
            Return False
        End If
    End Function

    <Task()> Public Shared Sub ValidarInventarioPosterior(ByVal data As StockData, ByVal services As ServiceProvider)
        Dim AppParamsStock As ParametroStocks = services.GetService(Of ParametroStocks)()
        If (AppParamsStock.TipoInventario = TipoInventario.PrimerMovimiento And data.FechaDocumento < data.Context.FechaUltimoInventario) _
        Or (AppParamsStock.TipoInventario = TipoInventario.UltimoMovimiento And data.FechaDocumento <= data.Context.FechaUltimoInventario) Then
            data.Context.Cancel = True
            data.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(49, data.Context.FechaUltimoInventario), services)
        End If
    End Sub

    <Task()> Public Shared Sub ValidarOperacionPosterior(ByVal data As StockData, ByVal services As ServiceProvider)

        Dim strOperaciones As String = String.Empty

        Dim f As New Filter
        f.Add(New StringFilterItem("IDAlmacen", data.Almacen))
        f.Add(New StringFilterItem("IDArticulo", data.Articulo))
        f.Add(New StringFilterItem("Lote", data.Lote))
        f.Add(New StringFilterItem("IDDeposito", data.Ubicacion))
        f.Add(New DateFilterItem("Fecha", FilterOperator.GreaterThan, data.FechaDocumento))
        Dim dtOperacionPosterior As DataTable = New BE.DataEngine().Filter("NegBdgOperacionPosterior", f)
        If dtOperacionPosterior.Rows.Count > 0 Then
            Dim lstOperaciones As List(Of String) = (From c In dtOperacionPosterior Where Not c.IsNull("NOperacion") Select CStr(c("NOperacion")) Distinct).ToList
            If Not lstOperaciones Is Nothing AndAlso lstOperaciones.Count > 0 Then
                strOperaciones = Strings.Join(lstOperaciones.ToArray, ",")
            End If
        End If

        If Length(strOperaciones) > 0 Then
            data.Context.Cancel = True
            data.Context.Detalle = ProcessServer.ExecuteTask(Of DataMessage, String)(AddressOf Message, New DataMessage(50, data.Articulo, strOperaciones), services)
        End If
    End Sub

    Public Class DataCreateIStockClass
        Public assemblyFile As String
        Public typeName As String

        Public Sub New(ByVal assemblyFile As String, ByVal typeName As String)
            Me.assemblyFile = assemblyFile
            Me.typeName = typeName
        End Sub
    End Class
    <Task()> Public Shared Function CreateIStockClass(ByVal data As DataCreateIStockClass, ByVal services As ServiceProvider) As IStock
        If Len(data.assemblyFile) > 0 And Len(data.typeName) > 0 Then
            Dim assemblyObject As System.Reflection.Assembly

            If System.IO.File.Exists(data.assemblyFile) Then
                assemblyObject = System.Reflection.Assembly.LoadFrom(data.assemblyFile)
            Else
                assemblyObject = System.Reflection.Assembly.Load(IO.Path.GetFileNameWithoutExtension(IO.Path.GetFileName(data.assemblyFile)))
            End If
            Return assemblyObject.CreateInstance(data.typeName, True)
        End If
    End Function

#End Region

#Region " Funciones relacionadas con la gestión de Activos asociados a Números de Serie "

    Public Class DataActualizarActivo
        Public stkData As StockData
        Public serie As DataRow

        Public Sub New(ByVal stkData As StockData, ByVal serie As DataRow)
            Me.stkData = stkData
            Me.serie = serie
        End Sub
    End Class
    <Task()> Public Shared Sub ActualizarActivo(ByVal data As DataActualizarActivo, ByVal services As ServiceProvider)
        If Not data.serie Is Nothing Then
            Dim a As New Activo
            Dim dt As DataTable = a.SelOnPrimaryKey(data.serie("NSerie"))
            Select Case data.serie.RowState
                Case DataRowState.Added
                    '//Si hay una gestion paralela de numeros de serie con activos,
                    '//si se da de alta un numero de serie automaticamente se da
                    '//de alta un activo
                    If dt.Rows.Count = 0 Then
                        dt = a.AddNewForm()
                    End If
                    Dim activo As DataRow = dt.Rows(0)
                    activo("IDEstadoActivo") = data.serie("IDEstadoActivo")
                    activo("FechaBaja") = DBNull.Value
                    activo("IDActivo") = data.serie("NSerie")
                    activo("NSerie") = activo("IDActivo")
                    If Length(activo("DescActivo")) = 0 Then activo("DescActivo") = "Activo Dado de Alta Automáticamente (Proceso de Floteo de Maquinaria)"
                    activo("IDArticulo") = data.serie("IDArticulo")
                    If Length(data.serie("IDOperario")) > 0 Then activo("IDOperario") = data.serie("IDOperario")

                    data.serie("IDActivo") = activo("IDActivo")

                    ProcessServer.ExecuteTask(Of DataActualizarActivo)(AddressOf ActualizarHistoricoEstadoActivo, data, services)
                Case DataRowState.Modified, DataRowState.Unchanged
                    If dt.Rows.Count > 0 Then
                        Dim activo As DataRow = dt.Rows(0)
                        If activo("IDEstadoActivo") <> data.serie("IDEstadoActivo") Then
                            activo("IDEstadoActivoAnterior") = activo("IDEstadoActivo")
                            If Length(activo("IDEstadoActivoAnterior")) = 0 Then
                                activo("IDEstadoActivoAnterior") = data.serie("IDEstadoActivo")
                            End If
                            activo("IDOperarioAnterior") = activo("IDOperario")
                            If Length(activo("IDOperarioAnterior")) = 0 Then
                                activo("IDOperarioAnterior") = data.serie("IDOperario")
                            End If
                            activo("FechaEstadoAnterior") = activo("FechaEstado")
                            If Not IsDate(activo("FechaEstadoAnterior")) Then
                                activo("FechaEstadoAnterior") = data.stkData.Context.FechaDocumento
                            End If
                            activo("IDArticulo") = data.serie("IDArticulo")
                            activo("IDEstadoActivo") = data.serie("IDEstadoActivo")
                            activo("IDOperario") = data.serie("IDOperario")
                            activo("FechaEstado") = data.stkData.Context.FechaDocumento
                            activo("NSerie") = data.serie("NSerie")
                            activo("FechaBaja") = DBNull.Value
                            If data.stkData.Context.PropiedadesEstado.Baja Then
                                activo("FechaBaja") = data.stkData.Context.FechaDocumento
                            End If
                            ProcessServer.ExecuteTask(Of DataActualizarActivo)(AddressOf ActualizarHistoricoEstadoActivo, data, services)
                        End If
                    End If
            End Select

            data.stkData.Context.ActivoBBDD = dt
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarHistoricoEstadoActivo(ByVal data As DataActualizarActivo, ByVal services As ServiceProvider)
        If Not data.serie Is Nothing AndAlso Length(data.serie("IDActivo")) > 0 Then
            Dim tipoMovimiento As DataRow = New TipoMovimiento().GetItemRow(data.stkData.Context.TipoMovimiento)
            Dim texto As String = Nz(tipoMovimiento("DescTipoMovimiento"))
            texto = String.Concat(texto, " ", data.stkData.Context.NumeroMovimiento)
            texto = String.Concat(texto, " ", Format(data.stkData.Context.FechaDocumento, "dd/MM/yy"))
            If Length(data.stkData.Context.Obra) > 0 Then
                Dim oc As BusinessHelper
                oc = BusinessHelper.CreateBusinessObject("ObraCabecera")
                Dim dt As DataTable = oc.SelOnPrimaryKey(data.stkData.Context.Obra)
                If dt.Rows.Count > 0 Then
                    texto = String.Concat(texto, " ", dt.Rows(0)("NObra"))
                End If
            End If

            Dim historico As DataTable = New HistoricoEstadoActivo().AddNewForm()
            If historico.Rows.Count > 0 Then
                historico.Rows(0)("IDActivo") = data.serie("IDActivo")
                historico.Rows(0)("IDEstadoActivo") = data.serie("IDEstadoActivo")
                historico.Rows(0)("FechaEstado") = data.stkData.Context.FechaDocumento
                historico.Rows(0)("IDOperario") = data.serie("IDOperario")
                If Len(texto) > 0 Then
                    'If Len(texto) > historico.Columns("Texto").MaxLength Then
                    '    historico.Rows(0)("Texto") = texto.Substring(0, historico.Columns("Texto").MaxLength).Trim()
                    'Else
                    historico.Rows(0)("Texto") = texto.Trim()
                    ' End If
                End If
            End If
            data.stkData.Context.HistoricoEstadoActivo = historico
        End If
    End Sub

#End Region

    <Task()> Public Shared Function ObtenerValoracionActualFIFO(ByVal data As Filter, ByVal services As ServiceProvider) As DataTable
        Return AdminData.Execute("sp_ObtenerValoracionActualFIFO", False, AdminData.ComposeFilter(data))
    End Function

#Region " Inventarios Permanentes "

    Public Enum enumTipoSincronizacionMovto
        AlbaranCompra
        AlbaranVenta
    End Enum

    <Task()> Public Shared Function CreateIStockInventarios(ByVal data As DataCreateIStockClass, ByVal services As ServiceProvider) As IStockInventarioPermanente
        If Len(data.assemblyFile) > 0 And Len(data.typeName) > 0 Then
            Dim assemblyObject As System.Reflection.Assembly

            If System.IO.File.Exists(data.assemblyFile) Then
                assemblyObject = System.Reflection.Assembly.LoadFrom(data.assemblyFile)
            Else
                assemblyObject = System.Reflection.Assembly.Load(IO.Path.GetFileNameWithoutExtension(IO.Path.GetFileName(data.assemblyFile)))
            End If
            Return assemblyObject.CreateInstance(data.typeName, True)
        End If
    End Function



#Region " CONTABILIDAD "

    <Task()> Public Shared Function GetLineasDescontabilizar(ByVal IDLineasMovimiento() As Object, ByVal services As ServiceProvider) As DataTable
        Dim f As New Filter

        Dim fLineasAlbaran As New Filter
        fLineasAlbaran.Add(New InListFilterItem("IDLineaMovimiento", IDLineasMovimiento, FilterType.Numeric))
        f.Add(fLineasAlbaran)

        Dim fTipoApunte As New Filter(FilterUnionOperator.Or)
        fTipoApunte.Add(New NumberFilterItem("IDTipoApunte", CInt(enumDiarioTipoApunte.MovimientoInput)))
        fTipoApunte.Add(New NumberFilterItem("IDTipoApunte", CInt(enumDiarioTipoApunte.MovimientoOutput)))
        f.Add(fTipoApunte)

        f.Add(New NumberFilterItem("Contabilizado", FilterOperator.NotEqual, CInt(enumContabilizado.NoContabilizado)))
        Dim dtLineasDesconta As DataTable = New BE.DataEngine().Filter("NegDescontabilizarMovtoStocks", f)
        Return dtLineasDesconta
    End Function

#End Region

#End Region

    '<Serializable()> _
    'Public Class dataValoracionEnFechaAlmacen
    '    Public Fecha As Date
    '    Public Filtros As Filter

    '    Public Sub New(ByVal Fecha As Date, ByVal Filtros As Filter)
    '        Me.Fecha = Fecha
    '        Me.Filtros = Filtros
    '    End Sub
    'End Class
    '<Task()> Public Shared Function ValoracionEnFechaAlmacen(ByVal data As dataValoracionEnFechaAlmacen, ByVal services As ServiceProvider) As DataTable
    '    Dim where As String = AdminData.ComposeFilter(data.Filtros)
    '    If where Is Nothing Then where = String.Empty
    '    Return AdminData.Execute("spCIValoracionAlmacenFecha", False, Format(data.Fecha, "yyyyMMdd"), where)
    'End Function

End Class