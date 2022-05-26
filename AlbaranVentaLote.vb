Public Class AlbaranVentaLote
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Private Const cnEntidad As String = "tbAlbaranVentaLote"
    Private Shared _L As _AlbaranVentaLote
    Private Shared _AVL As _AlbaranVentaLinea

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub


#Region " RegisterAddNewTasks "

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        data(_L.IDLineaLote) = AdminData.GetAutoNumeric
    End Sub

#End Region

#Region " Delete "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf EliminarMovimientoSalida)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.DeleteEntityRow)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.MarcarComoEliminado)
        deleteProcess.AddTask(Of DataRow)(AddressOf EliminarOperacionBodega)
    End Sub

    <Task()> Public Shared Sub EliminarOperacionBodega(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.Table.Columns.Contains("IDOperacion") AndAlso Length(data("IDOperacion")) > 0 Then
            Dim ClsBEOper As BusinessHelper = BusinessHelper.CreateBusinessObject("BdgOperacion")
            Dim DtOper As DataTable = ClsBEOper.Filter(New FilterItem("NOperacion", FilterOperator.Equal, data("IDOperacion")))
            If Not DtOper Is Nothing AndAlso DtOper.Rows.Count > 0 Then
                ClsBEOper.Delete(DtOper)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub EliminarMovimientoSalida(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Not data.IsNull(_L.IDMovimientoSalida) Then
            Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataRow, StockUpdateData)(AddressOf AlbaranVentaLote.EliminarMovimientoLineaLote, data, services)
            If Not updateData Is Nothing Then
                If updateData.Estado = EstadoStock.Actualizado Then
                    ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarEstadoLineaAlbaran, data, services)
                Else
                    'Se deja porque bajaUpdateData.Detalle ya viene traduccido
                    If Length(updateData.Log) > 0 Then Throw New Exception(updateData.Log)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarEstadoLineaAlbaran(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim albaran As DataTable = New AlbaranVentaLinea().Filter(New NumberFilterItem(_AVL.IDLineaAlbaran, data(_L.IDLineaAlbaran)))
        If Not albaran Is Nothing AndAlso albaran.Rows.Count > 0 Then
            albaran.Rows(0)(_AVL.EstadoStock) = enumavlEstadoStock.avlNoActualizado
            BusinessHelper.UpdateTable(albaran)
        End If
    End Sub

#End Region

#Region " Corregir Movimiento "

    <Task()> Public Shared Function CorregirMovimiento(ByVal Lotes As DataTable, ByVal services As ServiceProvider) As StockUpdateData()
        Dim updateData(-1) As StockUpdateData
        ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)

        For Each lote As DataRow In Lotes.Rows
            Dim data As StockUpdateData = ProcessServer.ExecuteTask(Of DataRow, StockUpdateData)(AddressOf CorregirMovimientoLineaLote, lote, services)
            If Not data Is Nothing Then
                ReDim Preserve updateData(UBound(updateData) + 1)
                updateData(UBound(updateData)) = data
            End If
        Next
        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.CommitTransaction, True, services)

        Return updateData
    End Function

    <Task()> Public Shared Function CorregirMovimientoLineaLote(ByVal lineaLote As DataRow, ByVal services As ServiceProvider) As StockUpdateData
        Dim updateData As StockUpdateData
        Dim Cantidad As Double = lineaLote(_L.QInterna)
        ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
        If Not lineaLote.IsNull(_L.IDMovimientoEntrada) Then
            'Correccion movimiento de entrada
            Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, lineaLote("IDMovimientoEntrada"), Cantidad)
            updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
            If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.Actualizado Then
                lineaLote(_L.IDMovimientoEntrada) = updateData.IDLineaMovimiento
            Else
                ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.RollbackTransaction, True, services)
                Return updateData
            End If
        End If

        If Not lineaLote.IsNull(_L.IDMovimientoSalida) Then
            'Correccion movimiento de salida
            Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, lineaLote("IDMovimientoSalida"), Cantidad)
            updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
            If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.Actualizado Then
                lineaLote(_L.IDMovimientoSalida) = updateData.IDLineaMovimiento
            Else
                ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.RollbackTransaction, True, services)
                Return updateData
            End If
        End If
        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.CommitTransaction, True, services)

        Return updateData
    End Function

#End Region

#Region " Eliminar Movimiento "

    <Task()> Public Shared Function EliminarMovimiento(ByVal Lotes As DataTable, ByVal services As ServiceProvider) As StockUpdateData()
        Dim updateData(-1) As StockUpdateData
        ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
        For Each lote As DataRow In Lotes.Rows
            Dim data As StockUpdateData = ProcessServer.ExecuteTask(Of DataRow, StockUpdateData)(AddressOf EliminarMovimientoLineaLote, lote, services)
            If Not data Is Nothing Then
                ReDim Preserve updateData(UBound(updateData) + 1)
                updateData(UBound(updateData)) = data
            End If
        Next
        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.CommitTransaction, True, services)

        Return updateData
    End Function

    <Task()> Public Shared Function EliminarMovimientoLineaLote(ByVal lineaLote As DataRow, ByVal services As ServiceProvider) As StockUpdateData
        Dim updateData As StockUpdateData
        ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
        If Not lineaLote.IsNull(_L.IDMovimientoEntrada) Then
            'Correccion movimiento de entrada
            Dim datElimMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Eliminar, lineaLote(_L.IDMovimientoEntrada))
            updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datElimMovto, services)

            If updateData Is Nothing OrElse updateData.Estado <> EstadoStock.Actualizado Then
                ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.RollbackTransaction, True, services)
                Return updateData
            End If
        End If

        If Not lineaLote.IsNull(_L.IDMovimientoSalida) Then
            'Correccion movimiento de salida
            Dim datElimMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Eliminar, lineaLote(_L.IDMovimientoSalida))
            updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datElimMovto, services)

            If updateData Is Nothing OrElse updateData.Estado <> EstadoStock.Actualizado Then
                ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.RollbackTransaction, True, services)
                Return updateData
            End If
        End If

        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.CommitTransaction, True, services)

        Return updateData
    End Function

#End Region

#Region "Tareas Públicas"

    <Serializable()> _
    Public Class DataInsertLoteLin
        Public Lotes() As StockData
        Public IDLineaAlbaran As Integer

        Public Sub New()
        End Sub
        Public Sub New(ByVal Lotes() As StockData, ByVal IDLineaAlbaran As Integer)
            Me.Lotes = Lotes
            Me.IDLineaAlbaran = IDLineaAlbaran
        End Sub
    End Class

    <Task()> Public Shared Sub InsertarDesgloseLoteLinea(ByVal data As DataInsertLoteLin, ByVal services As ServiceProvider)
        Dim DtAlbLin As DataTable = New AlbaranVentaLinea().SelOnPrimaryKey(data.IDLineaAlbaran)
        If Not DtAlbLin Is Nothing AndAlso DtAlbLin.Rows.Count > 0 Then
            Dim DocAlb As New DocumentoAlbaranVenta(DtAlbLin.Rows(0)("IDAlbaran"))
            For Each StData As StockData In data.Lotes
                If StData.Cantidad <> 0 Then
                    Dim newrow As DataRow = DocAlb.dtLote.NewRow
                    newrow("IDLineaLote") = AdminData.GetAutoNumeric
                    newrow("IDLineaAlbaran") = data.IDLineaAlbaran
                    newrow("Lote") = StData.Lote
                    newrow("Ubicacion") = StData.Ubicacion
                    newrow("QInterna") = StData.Cantidad
                    DocAlb.dtLote.Rows.Add(newrow)
                End If
            Next
            DocAlb.SetData()
        End If
    End Sub

#End Region

End Class

Public Class _AlbaranVentaLote
    Public Const IDLineaLote As String = "IDLineaLote"
    Public Const IDLineaAlbaran As String = "IDLineaAlbaran"
    Public Const Lote As String = "Lote"
    Public Const Ubicacion As String = "Ubicacion"
    Public Const QInterna As String = "QInterna"
    Public Const IDMovimientoSalida As String = "IDMovimientoSalida"
    Public Const IDMovimientoEntrada As String = "IDMovimientoEntrada"
    Public Const FechaCreacionAudi As String = "FechaCreacionAudi"
    Public Const FechaModificacionAudi As String = "FechaModificacionAudi"
    Public Const UsuarioAudi As String = "UsuarioAudi"
End Class