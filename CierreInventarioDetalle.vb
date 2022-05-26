Public Class CierreInventarioDetalle

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbCierreInventarioDetalle"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim DtCabecera As DataTable = New CierreInventario().SelOnPrimaryKey(data("IDEjercicio"), data("IDMesCierre"))
        If DtCabecera.Rows(0)("Contabilizado") Then ApplicationService.GenerateError("No se permite actualizar un Cierre de Inventario Contabilizado.")
        If DtCabecera.Rows(0)("Cerrado") Then ApplicationService.GenerateError("No se permite actualizar un Cierre de Inventario Cerrado.")
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarFechaCalculo)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDDetalle")) = 0 Then data("IDDetalle") = AdminData.GetAutoNumeric
        End If
    End Sub

    <Task()> Public Shared Sub AsignarFechaCalculo(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaCalculo")) = 0 Then data("FechaCalculo") = Today.Date

    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class DatosRecalcStockAFecha
        Public FechaCierre As Date
        Public DtDatos As DataTable
    End Class

    <Serializable()> _
    Public Class DatosRecalcPrecio
        Public IDEjercicio As String
        Public Periodo As Integer
        Public DtDatos As DataTable
    End Class

    <Serializable()> _
    Public Class DatosNuevoCritVal
        Public NuevoCriterio As enumtaValoracion
        Public DtDatos As DataTable
    End Class

    <Task()> Public Shared Sub RecalcularStockAFecha(ByVal data As DatosRecalcStockAFecha, ByVal services As ServiceProvider)
        If Not data.DtDatos Is Nothing AndAlso data.DtDatos.Rows.Count > 0 Then

            For Each dr As DataRow In data.DtDatos.Rows
                Dim datosStock As New DataArticuloAlmacenFecha(dr("IDArticulo"), dr("IDAlmacen"), data.FechaCierre)
                Dim stk As StockAFechaInfo = ProcessServer.ExecuteTask(Of DataArticuloAlmacenFecha, StockAFechaInfo)(AddressOf ProcesoStocks.GetStockAcumuladoAFecha, datosStock, services)
                dr("StockFisico") = stk.StockAFecha
                If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, dr("IDArticulo"), services) Then
                    dr("StockFisico2") = stk.StockAFecha2
                End If
                dr("FechaCalculo") = Today
            Next
            Dim ClsInvDetalle As New CierreInventarioDetalle
            ClsInvDetalle.Update(data.DtDatos)
        End If
    End Sub

    <Task()> Public Shared Sub RecalcularPrecio(ByVal data As DatosRecalcPrecio, ByVal services As ServiceProvider)
        If Length(data.IDEjercicio) > 0 AndAlso Not data.DtDatos Is Nothing AndAlso data.DtDatos.Rows.Count > 0 Then
            Dim intUDValoracion As Integer
            Dim DblPrecioAlmacenA, DblPrecioAlmacenB, DblValorA, DblValorB, DblPrecioEstandarA, _
                DblPrecioEstandarB, DblFIFOFechaA, DblFIFOFechaB, DblFIFOMvtoA, DblFIFOMvtoB, _
                DblPrecioMedioA, DblPrecioMedioB, DblPrecioUltimoA, DblPrecioUltimoB, DblPMInicialA, _
                DblStockInicial, DblPMInicialB As Double
            Dim DteFechaDesde, DteFechaHasta, DteFechaUltimoCierre As Date

            'If Not DtDatos Is Nothing AndAlso DtDatos.Rows.Count > 0 Then
            Dim DrDatos() As DataRow = data.DtDatos.Select("StockFisico<0 AND Activo=1")
            If DrDatos.Length > 0 Then
                ApplicationService.GenerateError("Uno o varios de los Artículos tienen un stock en fecha negativo.")
            Else
                Dim ClsCierre As Object = BusinessHelper.CreateBusinessObject("Cierre")
                Dim DtCierre As DataTable = ClsCierre.SelOnPrimaryKey(data.IDEjercicio, data.Periodo)
                If Not IsNothing(DtCierre) AndAlso DtCierre.Rows.Count > 0 Then
                    DteFechaDesde = DtCierre.Rows(0)("FechaDesde")
                    DteFechaHasta = DtCierre.Rows(0)("FechaHasta")
                End If

                For Each dr As DataRow In data.DtDatos.Select
                    DblPrecioAlmacenA = 0 : DblPrecioAlmacenB = 0
                    DblValorA = 0 : DblValorB = 0
                    DblPrecioEstandarA = 0 : DblPrecioEstandarB = 0
                    DblFIFOFechaA = 0 : DblFIFOFechaB = 0
                    DblFIFOMvtoA = 0 : DblFIFOMvtoB = 0
                    DblPrecioMedioA = 0 : DblPrecioMedioB = 0
                    DblPrecioUltimoA = 0 : DblPrecioUltimoB = 0

                    If dr("Activo") Then
                        intUDValoracion = IIf(dr("UDValoracion") > 0, dr("UDValoracion"), 1)
                        DblPrecioEstandarA = dr("PrecioEstandarA") / intUDValoracion
                        DblPrecioEstandarB = dr("PrecioEstandarB") / intUDValoracion

                        '///FIFO ordenado por fecha de mvto
                        Dim info As ValoracionPreciosInfo
                        Dim datosPrecio As New ProcesoStocks.DataValoracionFIFO(dr("IDArticulo"), dr("IDAlmacen"), dr("StockFisico"), DteFechaHasta, enumstkValoracionFIFO.stkVFOrdenarPorFecha)
                        info = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ProcesoStocks.ValoracionFIFO, datosPrecio, services)
                        If Not info Is Nothing Then
                            DblFIFOFechaA = info.PrecioA : DblFIFOFechaB = info.PrecioB
                        End If
                        '///FIFO ordenado por IdLineaMovimiento
                        datosPrecio = New ProcesoStocks.DataValoracionFIFO(dr("IDArticulo"), dr("IDAlmacen"), dr("StockFisico"), DteFechaHasta, enumstkValoracionFIFO.stkVFOrdenarPorMvto)
                        info = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ProcesoStocks.ValoracionFIFO, datosPrecio, services)
                        If Not info Is Nothing Then
                            DblFIFOMvtoA = info.PrecioA : DblFIFOMvtoB = info.PrecioB
                        End If
                        '///Precio Medio
                        Dim datosValPM As New ProcesoStocks.DataValoracionPrecioMedio(dr("IDArticulo"), dr("IDAlmacen"), DteFechaHasta, DteFechaDesde)
                        info = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionPrecioMedio, ValoracionPreciosInfo)(AddressOf ProcesoStocks.ValoracionPrecioMedio, datosValPM, services)
                        If Not info Is Nothing Then
                            DblPrecioMedioA = info.PrecioA : DblPrecioMedioB = info.PrecioB
                        End If
                        '///Precio Ultimo de Compra
                        DblPrecioUltimoA = dr("PrecioUltimaCompraA")
                        DblPrecioUltimoB = dr("PrecioUltimaCompraB")

                        '///Asignar el precio de almacen de acuerdo al criterio de valoracion del articulo
                        Select Case dr("CriterioValoracion")
                            Case enumtaValoracion.taPrecioEstandar
                                DblPrecioAlmacenA = DblPrecioEstandarA
                                DblPrecioAlmacenB = DblPrecioEstandarB
                            Case enumtaValoracion.taPrecioFIFOFecha
                                DblPrecioAlmacenA = DblFIFOFechaA
                                DblPrecioAlmacenB = DblFIFOFechaB
                            Case enumtaValoracion.taPrecioFIFOMvto
                                DblPrecioAlmacenA = DblFIFOMvtoA
                                DblPrecioAlmacenB = DblFIFOMvtoB
                            Case enumtaValoracion.taPrecioMedio
                                DblPrecioAlmacenA = DblPrecioMedioA
                                DblPrecioAlmacenB = DblPrecioMedioB
                            Case enumtaValoracion.taPrecioUltCompra
                                DblPrecioAlmacenA = DblPrecioUltimoA
                                DblPrecioAlmacenB = DblPrecioUltimoB
                        End Select

                        If dr("StockFisico") > 0 Then
                            DblValorA = DblPrecioAlmacenA * dr("StockFisico")
                            DblValorB = DblPrecioAlmacenB * dr("StockFisico")
                        Else
                            DblValorA = 0 : DblValorB = 0
                        End If
                    End If

                    dr("PrecioEstandarA") = DblPrecioEstandarA
                    dr("PrecioEstandarB") = DblPrecioEstandarB
                    dr("PrecioFIFOFechaA") = DblFIFOFechaA
                    dr("PrecioFIFOFechaB") = DblFIFOFechaB
                    dr("PrecioFIFOMvtoA") = DblFIFOMvtoA
                    dr("PrecioFIFOMvtoB") = DblFIFOMvtoB
                    dr("PrecioMedioA") = DblPrecioMedioA
                    dr("PrecioMedioB") = DblPrecioMedioB
                    dr("PrecioUltimoA") = DblPrecioUltimoA
                    dr("PrecioUltimoB") = DblPrecioUltimoB
                    dr("PrecioAlmacenA") = DblPrecioAlmacenA
                    dr("PrecioAlmacenB") = DblPrecioAlmacenB
                    dr("ValorA") = DblValorA
                    dr("ValorB") = DblValorB
                    dr("FechaCalculo") = DteFechaHasta
                Next
                Dim ClsInvDetalle As New CierreInventarioDetalle
                ClsInvDetalle.Update(data.DtDatos)
            End If
            'End If
        End If
    End Sub

    <Task()> Public Shared Sub NuevoCriterioValoracion(ByVal data As DatosNuevoCritVal, ByVal services As ServiceProvider)
        If Not data.DtDatos Is Nothing AndAlso data.DtDatos.Rows.Count > 0 Then
            Dim strIN As String
            For Each Dr As DataRow In data.DtDatos.Select
                If InStr(strIN, "'" & Dr("IDArticulo") & "'", CompareMethod.Text) = 0 Then
                    If Length(strIN) > 0 Then strIN = strIN & ","
                    strIN &= "'" & Dr("IDArticulo") & "'"
                End If
            Next
            Dim StrWhere As String = "IDArticulo IN (" & strIN & ")"
            Dim dtArticulo As DataTable = New Articulo().Filter(, StrWhere)
            If Not dtArticulo Is Nothing AndAlso dtArticulo.Rows.Count > 0 Then
                For Each drArticulo As DataRow In dtArticulo.Select
                    If Nz(drArticulo("CriterioValoracion")) <> CInt(data.NuevoCriterio) Then
                        drArticulo("CriterioValoracion") = CInt(data.NuevoCriterio)
                    End If
                Next

                For Each dr As DataRow In data.DtDatos.Select
                    Select Case data.NuevoCriterio
                        Case enumtaValoracion.taPrecioEstandar
                            dr("PrecioAlmacenA") = dr("PrecioEstandarA")
                            dr("PrecioAlmacenB") = dr("PrecioEstandarB")
                        Case enumtaValoracion.taPrecioFIFOFecha
                            dr("PrecioAlmacenA") = dr("PrecioFIFOFechaA")
                            dr("PrecioAlmacenB") = dr("PrecioFIFOFechaB")
                        Case enumtaValoracion.taPrecioFIFOMvto
                            dr("PrecioAlmacenA") = dr("PrecioFIFOMvtoA")
                            dr("PrecioAlmacenB") = dr("PrecioFIFOMvtoB")
                        Case enumtaValoracion.taPrecioMedio
                            dr("PrecioAlmacenA") = dr("PrecioMedioA")
                            dr("PrecioAlmacenB") = dr("PrecioMedioB")
                        Case enumtaValoracion.taPrecioUltCompra
                            dr("PrecioAlmacenA") = dr("PrecioUltimoA")
                            dr("PrecioAlmacenB") = dr("PrecioUltimoB")
                    End Select
                    dr("ValorA") = dr("StockFisico") * dr("PrecioAlmacenA")
                    dr("ValorB") = dr("StockFisico") * dr("PrecioAlmacenB")
                Next
                Dim ClsInvDetalle As New CierreInventarioDetalle
                ClsInvDetalle.Update(data.DtDatos)
                BusinessHelper.UpdateTable(dtArticulo)
            End If
        End If
    End Sub

#End Region

End Class
