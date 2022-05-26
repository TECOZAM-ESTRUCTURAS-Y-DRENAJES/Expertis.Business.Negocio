<Serializable()> _
Public Class AlbaranTransferenciaUpdateData
    Public IDAlbaranTransferencia() As String
    Public NAlbaran() As String
    Public StockUpdateData() As StockUpdateData
End Class

Public Class AlbaranTransferenciaCabecera

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbAlbaranTransferenciaCabecera"

#End Region

#Region "Eventos AlbaranTransferenciaCabecera"

    Public Overloads Overrides Function Update(ByVal dttSource As System.Data.DataTable) As System.Data.DataTable
        Dim services As ServiceProvider
        If Not dttSource Is Nothing AndAlso dttSource.Rows.Count > 0 Then
            For Each dr As DataRow In dttSource.Rows
                If dr.RowState = DataRowState.Added Then
                    If Length(dr("IdAlbaranTransferencia")) = 0 Then
                        dr("IdAlbaranTransferencia") = AdminData.GetAutoNumeric
                    End If
                ElseIf dr.RowState = DataRowState.Modified Then
                    If Length(dr("NMovimiento")) > 0 AndAlso dr("FechaAlbaran") & String.Empty <> dr("FechaAlbaran", DataRowVersion.Original) & String.Empty Then
                        ProcessServer.ExecuteTask(Of DataRow)(AddressOf CorregirMovimientos, dr, services)
                    End If
                End If
                'Incrementar el contador
                If Not IsDBNull(dr("IDContador")) Then
                    dr("NAlbaranTransferencia") = ProcessServer.ExecuteTask(Of String, String)(AddressOf Contador.CounterValueID, dr("IDContador"), services)
                End If
            Next

            AdminData.SetData(dttSource)
        End If
    End Function

    Public Overrides Function AddNewForm() As DataTable
        Dim DtNew As DataTable = MyBase.AddNewForm
        DtNew.Rows(0)("IdAlbaranTransferencia") = AdminData.GetAutoNumeric
        Dim StDatos As New Contador.DatosDefaultCounterValue
        StDatos.Row = DtNew.Rows(0)
        StDatos.EntityName = "AlbaranTransferenciaCabecera"
        StDatos.FieldName = "NAlbaranTransferencia"
        ProcessServer.ExecuteTask(Of Contador.DatosDefaultCounterValue)(AddressOf Contador.LoadDefaultCounterValue, StDatos, New ServiceProvider)
        DtNew.Rows(0)("FechaAlbaran") = Date.Today
        Return DtNew
    End Function

#End Region

#Region " CrearAlbaranTransferencia "

    Public Overridable Function CrearAlbaranTransferencia(ByVal dtSoliditudes As DataTable, ByVal IDContador As String, Optional ByVal FechaAlbaran As Date = cnMinDate) As AlbaranTransferenciaUpdateData
        Dim updateData As New AlbaranTransferenciaUpdateData
        ReDim updateData.IDAlbaranTransferencia(-1)
        ReDim updateData.NAlbaran(-1)
        ReDim updateData.StockUpdateData(-1)
        Dim ff As New Filter

        If Not dtSoliditudes Is Nothing AndAlso dtSoliditudes.Rows.Count > 0 Then
            Dim dtAlbaranTransferencia As DataTable = Me.AddNew
            Dim dtLineasTransferencia As DataTable = New AlbaranTransferenciaLinea().AddNew
            dtSoliditudes.DefaultView.Sort = "IDSolicitud"

            Dim drCabecera As DataRow
            Dim intIDSoliditudANT As Integer

            If Len(IDContador) > 0 Then
                Dim cont As New EntidadContador
                Dim f As New Filter
                f.Add(New StringFilterItem("Entidad", FilterOperator.Equal, "AlbaranTransferenciaCabecera"))
                f.Add(New StringFilterItem("IDContador", FilterOperator.Equal, IDContador))
                Dim dt As DataTable = cont.Filter(f)
                If IsNothing(dt) OrElse dt.Rows.Count = 0 Then
                    ApplicationService.GenerateError("El contador | no está asignado a los albaranes de venta ", Quoted(IDContador))
                End If
            End If

            For Each drSolicitud As DataRow In dtSoliditudes.Select(Nothing, "IDSolicitud")
                If drSolicitud("IDSolicitud") <> intIDSoliditudANT Then
                    drCabecera = NuevaCabeceraTransferencia(drSolicitud, IDContador, FechaAlbaran)
                    intIDSoliditudANT = drSolicitud("IDSolicitud")
                    dtAlbaranTransferencia.Rows.Add(drCabecera.ItemArray)
                    ActualizarEstadoSolicitud(drSolicitud("IDSolicitud"))
                End If

                If Not drCabecera Is Nothing Then
                    Dim drLineas As DataRow = NuevaLineaTransferencia(drCabecera("IDAlbaranTransferencia"), _
                                                                     drSolicitud, FechaAlbaran)
                    'Dim drLineas As DataRow = NuevaLineaTransferencia(drSolicitud)
                    dtLineasTransferencia.Rows.Add(drLineas.ItemArray)
                End If
            Next

            Me.BeginTx()
            AdminData.SetData(dtAlbaranTransferencia)
            AdminData.SetData(dtLineasTransferencia)

            'Actualizar Línea Solicitud
            Dim ATL As New AlbaranTransferenciaLinea
            ATL.ActualizarSolicitud(dtLineasTransferencia, False)

            For Each albaranActualizado As DataRow In dtAlbaranTransferencia.Rows
                ReDim Preserve updateData.IDAlbaranTransferencia(UBound(updateData.IDAlbaranTransferencia) + 1)
                updateData.IDAlbaranTransferencia(UBound(updateData.IDAlbaranTransferencia)) = albaranActualizado("IDAlbaranTransferencia")
                ReDim Preserve updateData.NAlbaran(UBound(updateData.NAlbaran) + 1)
                updateData.NAlbaran(UBound(updateData.NAlbaran)) = albaranActualizado("NAlbaranTransferencia")
            Next
            updateData.StockUpdateData = Me.ActualizarStock(dtAlbaranTransferencia)
            Return updateData
        End If
    End Function

    Private Sub ActualizarEstadoSolicitud(ByVal IDSolicitud As Integer)
        Dim clsSTC As New SolicitudTransferenciaCabecera
        Dim dt As DataTable = clsSTC.SelOnPrimaryKey(IDSolicitud)
        If dt.Rows.Count > 0 Then
            dt.Rows(0)("EstadoCabecera") = enumSTLEstadoLinea.STLCerrada
            AdminData.SetData(dt)
        End If
    End Sub

    Private Function NuevaCabeceraTransferencia(ByVal Cabecera As DataRow, ByVal IDContador As String, ByVal FechaAlbaran As Date) As DataRow
        Dim drCabeceraTransf As DataRow = Me.AddNewForm.Rows(0)

        If Len(IDContador) = 0 Then
            If Length(Cabecera("IDContador")) > 0 Then
                IDContador = Cabecera("IDContador")
            Else
                IDContador = Nz(drCabeceraTransf("IDContador"))
            End If
        End If

        If Length(IDContador) > 0 Then
            If Not Cabecera Is Nothing Then
                drCabeceraTransf("IDContador") = IDContador
                Dim StDatos As New Contador.DatosCounterValue
                StDatos.IDCounter = IDContador
                StDatos.TargetClass = Me
                StDatos.TargetField = "NAlbaranTransferencia"
                StDatos.DateField = "FechaAlbaran"
                StDatos.DateValue = FechaAlbaran
                drCabeceraTransf("NAlbaranTransferencia") = ProcessServer.ExecuteTask(Of Contador.DatosCounterValue, String)(AddressOf Contador.CounterValue, StDatos, New ServiceProvider)
                drCabeceraTransf("IDCentroGestionSolicitante") = Cabecera("IDCentroGestionSolicitante")
                drCabeceraTransf("IDAlmacenDestino") = Cabecera("IDAlmacenDestino")
                drCabeceraTransf("IDCentroGestionSolicitado") = Cabecera("IDCentroGestionSolicitado")
                drCabeceraTransf("IDAlmacenOrigen") = Cabecera("IDAlmacenOrigen")
                drCabeceraTransf("IDOperario") = Cabecera("IDOperario")
                drCabeceraTransf("FechaAlbaran") = FechaAlbaran
                drCabeceraTransf("IDSolicitud") = Cabecera("IDSolicitud")

            End If
        Else
            ApplicationService.GenerateError("No hay contadores definidos para esta entidad |")
        End If
        Return drCabeceraTransf
    End Function

    Private Function NuevaLineaTransferencia(ByVal intIDAlbaranTransferencia As Integer, ByVal drData As DataRow, ByVal FechaAlbaran As Date) As DataRow
        Dim fvl As New AlbaranTransferenciaLinea
        Dim dtLineas As DataTable = fvl.AddNew
        Dim Lineas As DataRow = dtLineas.NewRow

        If Not drData Is Nothing Then
            Lineas("IDAlbaranTransferenciaLinea") = AdminData.GetAutoNumeric
            Lineas("IDAlbaranTransferencia") = intIDAlbaranTransferencia
            Lineas("IdSolicitudLinea") = drData("IdSolicitudLinea")
            Lineas("IDArticulo") = drData("IDArticulo")
            Lineas("DescArticulo") = drData("DescArticulo")
            Lineas("CantidadTransferida") = drData("CantidadATransferir")

            Lineas("IdCentroGestionOrigen") = drData("IdCentroGestionOrigen")
            Lineas("IDAlmacenOrigen") = drData("IdAlmacenOrigen")
            Lineas("IDCentroGestionDestino") = drData("IdCentroGestionDestino")
            Lineas("IDAlmacenDestino") = drData("IdAlmacenDestino")
            Lineas("FechaTransferenciaReal") = FechaAlbaran
            'Lineas("IdAlbaranTransferencia") = drData("IdAlbaranTransferencia")

            Return Lineas
        End If
    End Function

#End Region

#Region "Actualizacion de stocks"

    Public Function ActualizarStock(ByVal Albaranes As DataTable) As StockUpdateData()
        Dim IDAlbaran(-1) As Integer
        For Each albaran As DataRow In Albaranes.Rows
            ArrayManager.Copy(CInt(albaran("IDAlbaranTransferencia")), IDAlbaran)
        Next
        Return Me.ActualizarStock(IDAlbaran)
    End Function

    Public Function ActualizarStock(ByVal IDAlbaran() As Integer) As StockUpdateData()
        Dim updateDataArray(-1) As StockUpdateData

        For Each id As Integer In IDAlbaran
            Dim updateData() As StockUpdateData
            updateData = Me.ActualizarStock(id)
            ArrayManager.Copy(updateData, updateDataArray)
        Next

        Return updateDataArray
    End Function

    Public Function ActualizarStock(ByVal IDAlbaran As Integer) As StockUpdateData()
        Dim updateDataArray(-1) As StockUpdateData

        If IDAlbaran <> 0 Then
            Dim Cabecera As DataTable = SelOnPrimaryKey(IDAlbaran)
            If Not Cabecera Is Nothing AndAlso Cabecera.Rows.Count > 0 Then
                Dim AVL As New AlbaranTransferenciaLinea
                Dim f As New Filter
                f.Add(New NumberFilterItem("IDAlbaranTransferencia", IDAlbaran))
                Dim Lineas As DataTable = AVL.Filter(f)
                If Not Lineas Is Nothing AndAlso Lineas.Rows.Count > 0 Then
                    Dim commit As Boolean
                    For Each linea As DataRow In Lineas.Rows
                        If linea("EstadoStock") = enumavlEstadoStock.avlNoActualizado Then
                            Me.BeginTx()
                            commit = True
                            Dim updateData() As StockUpdateData
                            updateData = AVL.ActualizarStock(Cabecera.Rows(0), linea)

                            If Not updateData Is Nothing Then
                                ArrayManager.Copy(updateData, updateDataArray)

                                For Each updateItem As StockUpdateData In updateData
                                    If Not updateItem Is Nothing Then
                                        If updateItem.Estado = EstadoStock.NoActualizado Then
                                            commit = False
                                            Me.RollbackTx()
                                            Exit For
                                        End If
                                    End If
                                Next
                                If commit Then Me.CommitTx()
                            End If
                        End If
                    Next
                    Me.BeginTx()
                    AdminData.SetData(Cabecera)
                    AdminData.SetData(Lineas)
                    Me.CommitTx()
                End If
            End If
        End If

        Return updateDataArray
    End Function

    <Task()> Private Shared Function CorregirMovimientos(ByVal cabeceraAlbaran As DataRow, ByVal services As ServiceProvider) As StockUpdateData()
        Dim returnData(-1) As StockUpdateData

        If cabeceraAlbaran("FechaAlbaran") <> DateTime.MinValue Then
            Dim FechaDocumento As Date = cabeceraAlbaran("FechaAlbaran")
            Dim ATL As New AlbaranTransferenciaLinea
            Dim f As New Filter
            f.Add(New NumberFilterItem("IDAlbaranTransferencia", cabeceraAlbaran("IDAlbaranTransferencia")))
            f.Add(New NumberFilterItem("EstadoStock", enumavlEstadoStock.avlActualizado))
            Dim lineasAlbaran As DataTable = ATL.Filter(f)
            If Not lineasAlbaran Is Nothing AndAlso lineasAlbaran.Rows.Count > 0 Then

                For Each lineaAlbaran As DataRow In lineasAlbaran.Rows
                    Dim updateData As StockUpdateData
                    '//Movimiento de salida
                    If Length(lineaAlbaran("IDMovimiento")) > 0 Then
                        Dim datCorrMovto As New ProcesoStocks.DataActualizarMovimiento(enumTipoActualizacion.Corregir, lineaAlbaran("IDMovimiento"), FechaDocumento)
                        updateData = ProcessServer.ExecuteTask(Of ProcesoStocks.DataActualizarMovimiento, StockUpdateData)(AddressOf ProcesoStocks.ActualizarMovimiento, datCorrMovto, services)
                        ArrayManager.Copy(updateData, returnData)
                    End If
                Next

                AdminData.SetData(lineasAlbaran)
            End If
        End If
        Return returnData
    End Function
#End Region

End Class

