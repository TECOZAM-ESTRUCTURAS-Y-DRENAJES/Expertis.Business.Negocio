Public Class AlbaranTransferenciaLinea

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbAlbaranTransferenciaLinea"

#End Region

    Dim DtUpdate As New DataTable

#Region "Eventos AlbaranTransferenciaLinea"

    Public Overrides Function AddNewForm() As DataTable
        Dim dt As DataTable = MyBase.AddNewForm
        Dim StDatos As New Contador.DatosDefaultCounterValue
        StDatos.Row = dt.Rows(0)
        StDatos.EntityName = Me.GetType.Name
        StDatos.FieldName = "NAlbaran"
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StDatos, New ServiceProvider)
        dt.Rows(0)("IdAlbaranTransferenciaLinea") = AdminData.GetAutoNumeric
        Return dt
    End Function

    Protected Overloads Overrides Sub Delete(ByVal dtrSource As System.Data.DataRow)
        Dim services As New ServiceProvider
        If Length(dtrSource("IdSolicitudLinea")) > 0 Then
            ActualizarSolicitud(dtrSource, True)
        End If

        Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of DataRow, StockUpdateData)(AddressOf EliminarMovimiento, dtrSource, services)
        If Not updateData Is Nothing Then
            If updateData.Estado <> EstadoStock.Actualizado Then
                Throw New Exception(updateData.Log)
            End If
        End If

        MyBase.Delete(dtrSource)
    End Sub

    Public Overloads Overrides Function Update(ByVal dttSource As System.Data.DataTable) As System.Data.DataTable
        Dim services As New ServiceProvider
        DtUpdate = dttSource.Clone
        For Each dr As DataRow In dttSource.Rows
            If dr.RowState = DataRowState.Added Then
                If Length(dr("IdAlbaranTransferenciaLinea")) = 0 Then
                    dr("IdAlbaranTransferenciaLinea") = AdminData.GetAutoNumeric
                End If
            ElseIf dr.RowState = DataRowState.Modified Then
                If Length(dr("IdSolicitudLinea")) > 0 And (Nz(dr("CantidadTransferida", DataRowVersion.Original)) <> Nz(dr("CantidadTransferida"))) Then
                    DtUpdate.Rows.Add(dr.ItemArray)
                ElseIf Nz(dr("IdSolicitudLinea", DataRowVersion.Original)) <> Nz(dr("IdSolicitudLinea")) Then
                    DtUpdate.Rows.Add(dr.ItemArray)
                ElseIf Length(dr("IdSolicitudLinea", DataRowVersion.Original)) > 0 AndAlso Length(dr("IdSolicitudLinea")) = 0 Then
                    DtUpdate.Rows.Add(dr.ItemArray)
                End If
                If dr("EstadoStock") = enumavlEstadoStock.avlActualizado AndAlso Nz(dr("CantidadTransferida", DataRowVersion.Original)) <> Nz(dr("CantidadTransferida")) Then
                    ProcessServer.ExecuteTask(Of DataRow)(AddressOf CorregirMovimiento, dr, services)
                End If
            End If
        Next

        AdminData.SetData(dttSource)
        Me.Updated(DtUpdate)
    End Function

    Public Overrides Sub Updated(ByVal data As System.Data.DataTable)
        For Each dr As DataRow In data.Rows
            If dr.RowState = DataRowState.Added Then
                ActualizarSolicitud(dr)
            ElseIf dr.RowState = DataRowState.Modified Then
                If Nz(dr("IdSolicitudLinea", DataRowVersion.Original)) <> Nz(dr("IdSolicitudLinea")) Then
                    ActualizarSolicitud(dr, True)
                Else
                    ActualizarSolicitud(dr)
                End If
            End If
        Next
    End Sub

#End Region

#Region "Funciones Públicas"

    Public Sub ActualizarSolicitud(ByRef dtLineasTransferencia As DataTable, Optional ByVal blnBorrado As Boolean = False)
        For Each drLineasTransferencia As DataRow In dtLineasTransferencia.Rows
            ActualizarSolicitud(drLineasTransferencia, blnBorrado)
        Next        
    End Sub

    Public Sub ActualizarSolicitud(ByRef drLineasTransferencia As DataRow, Optional ByVal blnBorrado As Boolean = False)
        Dim STL As New SolicitudTransferenciaLinea
        Dim dtAVC As DataTable = STL.SelOnPrimaryKey(drLineasTransferencia("IdSolicitudLinea"))
        Dim SolicitudLinea As DataRow = dtAVC.Rows(0)

        If blnBorrado Then
            SolicitudLinea("CantidadRecibida") = SolicitudLinea("CantidadRecibida") - drLineasTransferencia("CantidadTransferida")
        Else
            Dim TransferidaOriginal As Integer = 0
            '   If SolicitudLinea.RowState = DataRowState.Modified Then

            If drLineasTransferencia.RowState <> DataRowState.Added AndAlso Nz(drLineasTransferencia("CantidadTransferida", DataRowVersion.Original)) <> Nz(drLineasTransferencia("CantidadTransferida")) Then
                TransferidaOriginal = Nz(drLineasTransferencia("CantidadTransferida", DataRowVersion.Original))
            End If
            ' End If
            'SolicitudLinea("CantidadRecibida") = SolicitudLinea("CantidadRecibida") + (drLineasTransferencia("CantidadTransferida") - TransferidaOriginal)
            SolicitudLinea("CantidadRecibida") = (drLineasTransferencia("CantidadTransferida") - TransferidaOriginal)

            If Nz(SolicitudLinea("CantidadRecibida"), 0) <> 0 Then
                If System.Math.Abs(SolicitudLinea("CantidadRecibida")) >= System.Math.Abs(SolicitudLinea("CantidadSolicitada")) Then
                    SolicitudLinea("EstadoLinea") = enumSTLEstadoLinea.STLCerrada
                ElseIf System.Math.Abs(SolicitudLinea("CantidadRecibida")) < System.Math.Abs(SolicitudLinea("CantidadSolicitada")) Then
                    SolicitudLinea("EstadoLinea") = enumSTLEstadoLinea.STLRecibida
                End If
            Else
                SolicitudLinea("EstadoLinea") = enumSTLEstadoLinea.STLCerrada
            End If
        End If
        STL.Update(dtAVC)
    End Sub

    Friend Function ActualizarStock(ByVal cabeceraAlbaran As DataRow, ByVal lineaAlbaran As DataRow) As StockUpdateData()
        Dim updateDataArray(-1) As StockUpdateData
        Dim Cancel As Boolean

        '//Movimientos de la linea de albaran
        If Not Cancel Then
            Dim updateData() As StockUpdateData
            updateData = ActualizarStockTx(cabeceraAlbaran, lineaAlbaran)
            If updateData.Length > 0 Then
                'Primer elemento corresponde a la salida
                If Not updateData(0) Is Nothing Then
                    If updateData(0).Estado = EstadoStock.Actualizado Then
                        lineaAlbaran("IDMovimientoSalida") = updateData(0).IDLineaMovimiento
                        If Not updateData(1) Is Nothing Then
                            lineaAlbaran("IDMovimientoEntrada") = updateData(1).IDLineaMovimiento
                        End If
                        lineaAlbaran("EstadoStock") = enumavlEstadoStock.avlActualizado
                    Else
                        lineaAlbaran("IDMovimientoSalida") = DBNull.Value
                        lineaAlbaran("IDMovimientoEntrada") = DBNull.Value
                        If updateData(0).Estado = EstadoStock.NoActualizado Then
                            lineaAlbaran("EstadoStock") = enumavlEstadoStock.avlNoActualizado
                        ElseIf updateData(0).Estado = EstadoStock.SinGestion Then
                            lineaAlbaran("EstadoStock") = enumavlEstadoStock.avlSinGestion
                        End If
                    End If
                    ArrayManager.Copy(updateData(0), updateDataArray)
                End If
            End If
        End If

        Return updateDataArray
    End Function

    Private Function ActualizarStockTx(ByVal cabeceraAlbaran As DataRow, ByVal lineaAlbaran As DataRow, Optional ByVal lineaAlbaranLote As DataRow = Nothing) As StockUpdateData()
        Dim services As New ServiceProvider
        Dim updateDataArray(1) As StockUpdateData
        Dim updateSalida As StockUpdateData
        Dim updateEntrada As StockUpdateData

        Dim NumeroMovimiento As Long
        If IsNumeric(cabeceraAlbaran("NMovimiento")) AndAlso Not cabeceraAlbaran("NMovimiento") = 0 Then
            NumeroMovimiento = cabeceraAlbaran("NMovimiento")
        Else
            NumeroMovimiento = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf ProcesoStocks.NuevoNumeroMovimiento, Nothing, services)
            cabeceraAlbaran("NMovimiento") = NumeroMovimiento
        End If

        '//Movimientos de salida
        Dim salida As New StockData(lineaAlbaran("IDArticulo"), lineaAlbaran("IDAlmacenOrigen"), lineaAlbaran("CantidadTransferida"), 0, 0, cabeceraAlbaran("FechaAlbaran"), enumTipoMovimiento.tmSalTransferencia, cabeceraAlbaran("NAlbaranTransferencia"))

        Dim dataSal As New DataNumeroMovimiento(NumeroMovimiento, salida)
        updateSalida = ProcessServer.ExecuteTask(Of DataNumeroMovimiento, StockUpdateData)(AddressOf ProcesoStocks.Salida, dataSal, services)

        updateDataArray(0) = updateSalida

        '//Movimientos de salida
        Dim entrada As New StockData(lineaAlbaran("IDArticulo"), lineaAlbaran("IDAlmacenDestino"), lineaAlbaran("CantidadTransferida"), 0, 0, cabeceraAlbaran("FechaAlbaran"), enumTipoMovimiento.tmEntTransferencia, cabeceraAlbaran("NAlbaranTransferencia"))
       

        Dim datMovto As New DataNumeroMovimiento(NumeroMovimiento, entrada)
        updateEntrada = ProcessServer.ExecuteTask(Of DataNumeroMovimiento, StockUpdateData)(AddressOf ProcesoStocks.Entrada, datMovto, services)

        updateDataArray(1) = updateEntrada

        Return updateDataArray
    End Function

    <Task()> Public Shared Function CorregirMovimiento(ByVal lineaAlbaran As DataRow, ByVal services As ServiceProvider) As StockUpdateData 'Aqui hay q distinguir los dos movimientos (Entrada/Salida).
        Dim updateData As StockUpdateData
        ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
        'Salida
        Dim Cantidad As Double = lineaAlbaran("CantidadTransferida")
        If Not lineaAlbaran.IsNull("IDMovimientoSalida") Then
            '//Correccion movimiento de salida
            Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(lineaAlbaran("IDMovimientoSalida"), False, True, Cantidad, 0, 0)
            updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)

            If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.Actualizado Then
                lineaAlbaran("IDMovimientoSalida") = updateData.IDLineaMovimiento
            Else
                ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.RollbackTransaction, True, services)
                Return updateData
            End If
        End If
        'Entrada
        Cantidad = lineaAlbaran("CantidadTransferida")
        If Not lineaAlbaran.IsNull("IDMovimientoEntrada") Then
            '//Correccion movimiento de salida
            Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(lineaAlbaran("IDMovimientoEntrada"), False, True, Cantidad, 0, 0)
            updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)

            If Not updateData Is Nothing AndAlso updateData.Estado = EstadoStock.Actualizado Then
                lineaAlbaran("IDMovimientoEntrada") = updateData.IDLineaMovimiento
            Else
                ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.RollbackTransaction, True, services)
                Return updateData
            End If
        End If
        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.CommitTransaction, True, services)
        Return updateData
    End Function

    <Task()> Public Shared Function EliminarMovimiento(ByVal lineaAlbaran As DataRow, ByVal services As ServiceProvider) As StockUpdateData
        If IsNumeric(lineaAlbaran("IDMovimiento")) Then
            'Correccion movimiento de salida
            ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)

            Dim datElimMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Eliminar, lineaAlbaran("IDMovimiento"))
            Dim updateData As StockUpdateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datElimMovto, services)

            If updateData Is Nothing OrElse updateData.Estado <> EstadoStock.Actualizado Then
                ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.RollbackTransaction, True, services)
                Return updateData
            End If
            ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.CommitTransaction, True, services)

            'pend Numeros de serie
            Return updateData
        End If
    End Function

#End Region

End Class