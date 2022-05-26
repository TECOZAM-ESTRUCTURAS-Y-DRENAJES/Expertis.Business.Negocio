Public Class ProgramaLinea

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbProgramaLinea"

#End Region

#Region "Eventos RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("IdLineaPrograma") = AdminData.GetAutoNumeric
    End Sub

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim services As New ServiceProvider
        Dim oBRL As New SynonymousBusinessRules
        oBRL.AddSynonymous("QPrevista", "Cantidad")
        oBRL = ProcessServer.ExecuteTask(Of BusinessRules, BusinessRules)(AddressOf ProcesoComercial.DetailBusinessRulesLin, oBRL, services)
        Return oBRL
    End Function

#End Region

#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarConfirmacion)
    End Sub

    <Task()> Public Shared Sub ComprobarConfirmacion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Nz(data("Confirmada"), False) Then
            ApplicationService.GenerateError("No se puede borrar líneas confirmadas.")
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarPrograma)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarArticulo)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarUDValoracion)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarTipoIva)
    End Sub

    <Task()> Public Shared Sub ComprobarPrograma(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDPrograma")) = 0 Then ApplicationService.GenerateError("El identificador del Programa de Entrega no es válido.")
    End Sub

    <Task()> Public Shared Sub ComprobarArticulo(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ComprobarUDValoracion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("UdValoracion")) > 0 Then
            If data("UdValoracion") <= 0 Then ApplicationService.GenerateError("La Unidad de Valoración ha de ser mayor que 0.")
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarTipoIva(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipoIva")) = 0 Then ApplicationService.GenerateError("El Tipo IVA es un dato obligatorio.")
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarUdMedida)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarValoresPredeterminados)
        updateProcess.AddTask(Of DataRow)(AddressOf RecalcularImporteLinea)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.UpdateEntityRow)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DataRow)(AddressOf GeneraLineasHistorico)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDLineaPrograma")) = 0 Then
                data("IDLineaPrograma") = AdminData.GetAutoNumeric
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarUdMedida(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) > 0 AndAlso Length(data("IDUDMedida")) = 0 Then
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(data("IDArticulo"))
            If Length(ArtInfo.IDUDVenta) > 0 Then
                data("IDUDMedida") = ArtInfo.IDUDVenta
            Else : ApplicationService.GenerateError("La Unidad de Medida  no es válida. Es posible que el Artículo no tenga establecida su Unidad de Venta.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarValoresPredeterminados(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            data("Confirmada") = CBool(enumplEstadoLinea.plNoConfirmada)
            data("QConfirmada") = 0
            data("FechaConfirmacion") = System.DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub RecalcularImporteLinea(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim BlnRecalculo As Boolean = False
        If data.RowState = DataRowState.Added Then BlnRecalculo = True
        If data.RowState = DataRowState.Modified Then
            BlnRecalculo = ((data("Precio", DataRowVersion.Default) <> data("Precio", DataRowVersion.Original)) OrElse _
                            (data("QPrevista", DataRowVersion.Default) <> data("QPrevista", DataRowVersion.Original)) OrElse _
                            (data("UdValoracion", DataRowVersion.Default) <> data("UdValoracion", DataRowVersion.Original)) OrElse _
                            (data("Dto1", DataRowVersion.Default) <> data("Dto1", DataRowVersion.Original)) OrElse _
                            (data("Dto2", DataRowVersion.Default) <> data("Dto2", DataRowVersion.Original)) OrElse _
                            (data("Dto3", DataRowVersion.Default) <> data("Dto3", DataRowVersion.Original)))
        End If

        If BlnRecalculo Then
            Dim Programas As EntityInfoCache(Of ProgramaCabeceraInfo) = services.GetService(Of EntityInfoCache(Of ProgramaCabeceraInfo))()
            Dim ProgInfo As ProgramaCabeceraInfo = Programas.GetEntity(data("IDPrograma"))
            Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
            Dim MonInfo As MonedaInfo = Monedas.GetMoneda(ProgInfo.IDMoneda, ProgInfo.FechaPrograma)
            data("Precio") = xRound(Nz(data("Precio"), 0), MonInfo.NDecimalesPrecio)
            If Nz(data("UdValoracion"), 0) > 0 Then
                data("Importe") = (data("Precio") * (data("QPrevista") / data("UdValoracion"))) * (1 - (data("Dto1") / 100)) * (1 - (data("Dto2") / 100)) * (1 - (data("Dto3") / 100))
            End If
            Dim ValAyB As New ValoresAyB(New DataRowPropertyAccessor(data), MonInfo.ID, MonInfo.CambioA, MonInfo.CambioB)
            ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
        End If
    End Sub

    <Task()> Public Shared Sub GeneraLineasHistorico(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dtHistorico As DataTable = New HistPrograma().Filter(New NoRowsFilterItem())
        Dim drRowHistorico As DataRow
        Dim drRowLineaPrograma As DataRow = data
        Select Case drRowLineaPrograma.RowState
            Case DataRowState.Added
                drRowHistorico = dtHistorico.NewRow
                drRowHistorico("IDHistProg") = AdminData.GetAutoNumeric
                drRowHistorico("IDLineaPrograma") = drRowLineaPrograma("IDLineaPrograma")
                drRowHistorico("IDPrograma") = drRowLineaPrograma("IDPrograma")
                drRowHistorico("FechaEntregaAntigua") = System.DBNull.Value
                drRowHistorico("FechaEntregaNueva") = drRowLineaPrograma("FechaEntrega")
                drRowHistorico("CantidadAntigua") = System.DBNull.Value
                drRowHistorico("CantidadNueva") = drRowLineaPrograma("QPrevista")
                drRowHistorico("ModificacionCantidades") = 0
                drRowHistorico("AlteracionFecha") = 0
                drRowHistorico("FechaCreacionAudi") = Today
                drRowHistorico("UsuarioAudi") = AdminData.GetSessionInfo.UserName
            Case DataRowState.Modified
                '//Sólo insertaremos línea en el histórico si ha variado la QPrevista o la Fecha de Entrega.
                If AreDifferents(drRowLineaPrograma("QPrevista"), drRowLineaPrograma("QPrevista", DataRowVersion.Original)) OrElse _
                   AreDifferents(drRowLineaPrograma("FechaEntrega"), drRowLineaPrograma("FechaEntrega", DataRowVersion.Original)) Then
                    drRowHistorico = dtHistorico.NewRow
                    drRowHistorico("IDHistProg") = AdminData.GetAutoNumeric
                    drRowHistorico("IDLineaPrograma") = drRowLineaPrograma("IDLineaPrograma")
                    drRowHistorico("IDPrograma") = drRowLineaPrograma("IDPrograma")
                    drRowHistorico("FechaEntregaAntigua") = drRowLineaPrograma("FechaEntrega", DataRowVersion.Original)
                    drRowHistorico("FechaEntregaNueva") = drRowLineaPrograma("FechaEntrega", DataRowVersion.Default)
                    drRowHistorico("CantidadAntigua") = drRowLineaPrograma("QPrevista", DataRowVersion.Original)
                    drRowHistorico("CantidadNueva") = drRowLineaPrograma("QPrevista", DataRowVersion.Default)
                    drRowHistorico("ModificacionCantidades") = drRowLineaPrograma("QPrevista", DataRowVersion.Default) - drRowLineaPrograma("QPrevista", DataRowVersion.Original)
                    If Not IsDBNull(drRowLineaPrograma("FechaEntrega", DataRowVersion.Original)) AndAlso Not drRowLineaPrograma.IsNull("FechaEntrega") Then
                        drRowHistorico("AlteracionFecha") = CDate(drRowLineaPrograma("FechaEntrega")).Subtract(CDate(drRowLineaPrograma("FechaEntrega", DataRowVersion.Original))).TotalDays
                    End If
                    drRowHistorico("FechaModificacionAudi") = Today
                    drRowHistorico("UsuarioAudi") = AdminData.GetSessionInfo.UserName
                End If
        End Select
        If Not IsNothing(drRowHistorico) Then dtHistorico.Rows.Add(drRowHistorico)
        BusinessHelper.UpdateTable(dtHistorico)
    End Sub

#End Region

End Class