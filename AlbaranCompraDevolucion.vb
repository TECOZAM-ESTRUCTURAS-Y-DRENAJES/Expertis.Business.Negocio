Public Class AlbaranCompraDevolucion
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Private Const cnEntidad As String = "tbAlbaranCompraDevolucion"
    Private Const ESTADOACTIVO_DEVUELTOPROVEEDOR As String = "15"

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

#Region " RegisterUpdateTasks "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIDLineaDevolucion)
    End Sub

    <Task()> Public Shared Sub AsignarIDLineaDevolucion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDLineaDevolucion")) = 0 Then data("IDLineaDevolucion") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

#Region " Devoluciones "

    <Serializable()> _
    Public Class dataDevolucion
        Public IDLineaAlbaran As Integer
        Public Cantidad As Double
        Public Fecha As Date
        Public IDAlmacen As String
        Public IDLineaDevolucion As Integer

        Public Sub New(ByVal IDLineaAlbaran As Integer, ByVal Cantidad As Double, ByVal Fecha As Date, ByVal IDAlmacen As String, ByVal IDLineaDevolucion As Integer)
            Me.IDLineaAlbaran = IDLineaAlbaran
            Me.Cantidad = Cantidad
            Me.Fecha = Fecha
            Me.IDAlmacen = IDAlmacen
            Me.IDLineaDevolucion = IDLineaDevolucion
        End Sub
    End Class

#Region " Crear Devolucion "

    <Serializable()> _
    Public Class dataDevolucionMaterial
        Public datosDevolucion() As dataDevolucion
        Public IDOperario As String

        Public Sub New(ByVal datosDevolucion() As dataDevolucion, ByVal IDOperario As String)
            Me.datosDevolucion = datosDevolucion
            Me.IDOperario = IDOperario
        End Sub
    End Class
    <Task()> Public Shared Function DevolucionMaterial(ByVal data As dataDevolucionMaterial, ByVal services As ServiceProvider) As Integer
        Dim ACL As New AlbaranCompraLinea
        Dim ACD As New AlbaranCompraDevolucion
        Dim dtACD As DataTable = ACD.AddNew()
        For Each info As dataDevolucion In data.datosDevolucion
            Dim drACL As DataRow = ACL.GetItemRow(info.IDLineaAlbaran)
            drACL("QDevuelta") += info.Cantidad
            If drACL("QServida") <= drACL("QDevuelta") Then
                drACL("EstadoDevolucion") = enumacDevolucionRealquiler.acdDevuelto
            ElseIf drACL("QDevuelta") = 0 Then
                drACL("EstadoDevolucion") = enumacDevolucionRealquiler.acdPendienteDevolucion
            Else
                drACL("EstadoDevolucion") = enumacDevolucionRealquiler.acdParcialmenteDevuelto
            End If

            Dim datosStock As New dataActualizarStock(drACL("IDAlbaran"), drACL("FechaDevolucion"), drACL("IDArticulo"), drACL("IDAlmacen"), drACL("Texto"), drACL("Lote"), _
                                                      drACL("Ubicacion"), drACL("QDevuelta"), drACL("ImporteA"), drACL("ImporteB"), drACL("IDObra"), drACL("IDOperario"))

            Dim ResulStock As StockUpdateData = ProcessServer.ExecuteTask(Of dataActualizarStock, StockUpdateData)(AddressOf ActualizarStock, datosStock, services)
            If ResulStock.Estado = EstadoStock.Actualizado Then
                Dim drACD As DataRow = dtACD.NewRow
                drACD("IDLineaAlbaran") = info.IDLineaAlbaran
                drACD("QDevuelta") = info.Cantidad
                drACD("FechaDevolucion") = info.Fecha
                drACD("IDAlmacen") = info.IDAlmacen
                If Length(ResulStock.IDLineaMovimiento) > 0 Then
                    drACD("IDMovimientoDevolucion") = ResulStock.IDLineaMovimiento
                Else
                    drACD("IDMovimientoDevolucion") = DBNull.Value
                End If
                ACD.Update(drACD.Table)
                ACL.Update(drACL.Table)
            End If
        Next
    End Function

    <Serializable()> _
     Public Class dataActualizarStock
        Friend datosStock As StockData

        Public Sub New(ByVal IDAlbaran As Integer, ByVal FechaDevolucion As Date, ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal Texto As String, ByVal Lote As String, _
                       ByVal Ubicacion As String, ByVal QDevuelta As Double, ByVal ImporteA As Double, ByVal ImporteB As Double, ByVal IDObra As Integer, ByVal IDOperario As String)

            datosStock = New StockData
            datosStock.IDDocumento = IDAlbaran
            datosStock.FechaDocumento = FechaDevolucion
            datosStock.Articulo = IDArticulo
            datosStock.Almacen = IDAlmacen
            datosStock.TipoMovimiento = enumTipoMovimiento.tmSalRealquiler
            datosStock.Texto = Texto
            datosStock.Lote = Lote
            datosStock.NSerie = Lote
            datosStock.Ubicacion = Ubicacion
            datosStock.EstadoNSerie = ESTADOACTIVO_DEVUELTOPROVEEDOR
            datosStock.Cantidad = QDevuelta
            If datosStock.Cantidad <> 0 Then
                datosStock.PrecioA = ImporteA / datosStock.Cantidad
                datosStock.PrecioB = ImporteB / datosStock.Cantidad
            End If
            If IDObra > 0 Then datosStock.Obra = IDObra
            datosStock.Operario = IDOperario
        End Sub
    End Class
    <Task()> Public Shared Function ActualizarStock(ByVal data As dataActualizarStock, ByVal services As ServiceProvider) As StockUpdateData
        Dim NumMovimiento As Integer = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf ProcesoStocks.NuevoNumeroMovimiento, Nothing, services)

        Dim dataSalida As New DataNumeroMovimientoSinc(NumMovimiento, data.datosStock)
        Return ProcessServer.ExecuteTask(Of DataNumeroMovimientoSinc, StockUpdateData)(AddressOf ProcesoStocks.Salida, dataSalida, services)
    End Function

#End Region

#Region " Borrar Devolucion "

    <Serializable()> _
    Public Class dataBorrarDevolucionMaterial
        Public datosDevolucion() As dataDevolucion

        Public Sub New(ByVal datosDevolucion() As dataDevolucion)
            Me.datosDevolucion = datosDevolucion
        End Sub
    End Class
    <Task()> Public Shared Sub BorrarDevolucionMaterial(ByVal data As dataBorrarDevolucionMaterial, ByVal services As ServiceProvider)
        Dim ACL As New AlbaranCompraLinea
        Dim ACD As New AlbaranCompraDevolucion
        For Each info As dataDevolucion In data.datosDevolucion
            Dim dtACD As DataTable = ACD.SelOnPrimaryKey(info.IDLineaDevolucion)
            If Not dtACD Is Nothing AndAlso dtACD.Rows.Count > 0 Then ACD.Delete(dtACD)
            Dim dtACL As DataTable = ACL.SelOnPrimaryKey(info.IDLineaAlbaran)
            If Not dtACL Is Nothing AndAlso dtACL.Rows.Count > 0 Then
                dtACL.Rows(0)("QDevuelta") -= info.Cantidad
                If dtACL.Rows(0)("QDevuelta") = 0 Then
                    dtACL.Rows(0)("EstadoDevolucion") = enumacDevolucionRealquiler.acdPendienteDevolucion
                ElseIf dtACL.Rows(0)("QServida") <= dtACL.Rows(0)("QDevuelta") Then
                    dtACL.Rows(0)("EstadoDevolucion") = enumacDevolucionRealquiler.acdDevuelto
                Else
                    dtACL.Rows(0)("EstadoDevolucion") = enumacDevolucionRealquiler.acdParcialmenteDevuelto
                End If
            End If

            ACL.Update(dtACL)
        Next
    End Sub

#End Region

#End Region

End Class