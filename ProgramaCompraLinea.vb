Public Class ProgramaCompraLinea

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbProgramaCompraLinea"

#End Region

#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarConfirmacion)
    End Sub

    <Task()> Public Shared Sub ComprobarConfirmacion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Nz(data("Confirmada"), False) Then ApplicationService.GenerateError("No se puede borrar líneas confirmadas.")
    End Sub

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim services As New ServiceProvider
        Dim oBRL As New SynonymousBusinessRules
        oBRL.AddSynonymous("QPrevista", "Cantidad")
        oBRL = ProcessServer.ExecuteTask(Of BusinessRules, BusinessRules)(AddressOf ProcesoCompra.DetailBusinessRulesLin, oBRL, services)
        Return oBRL
    End Function

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarFactor)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarCantidades)
        updateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.AplicarDecimales)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.UpdateEntityRow)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DataRow, DataTable)(AddressOf GeneraLineasHistorico)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es un dato obligatorio.")
        If Length(data("DescArticulo")) = 0 Then ApplicationService.GenerateError("La descripción del Artículo es obligatoria")
        If Nz(data("QPrevista")) = 0 Then ApplicationService.GenerateError("Cantidad no válida.")
        If Length(data("CContable")) = 0 Then ApplicationService.GenerateError("La Cuenta Contable es un dato obligatorio.")
        If Length(data("FechaEntrega")) = 0 Then ApplicationService.GenerateError("La Fecha de Entrega es un campo obligatorio.")
        If Length(data("IDAlmacen")) = 0 Then ApplicationService.GenerateError("El Almacén es un dato obligatorio.")
        If Length(data("IDTipoIva")) = 0 Then ApplicationService.GenerateError("El Tipo IVA es un dato obligatorio.")
        If Length(data("IDUdMedida")) = 0 Then ApplicationService.GenerateError("La unidad de medida es obligatoria")
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDLineaPrograma")) = 0 Then data("IDLineaPrograma") = AdminData.GetAutoNumeric
        End If
    End Sub

    <Task()> Public Shared Sub AsignarFactor(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("Factor")) = 0 Then
            Dim StDatos As New ArticuloUnidadAB.DatosFactorConversion
            StDatos.IDArticulo = data("IDArticulo")
            StDatos.IDUdMedidaA = data("IDUDMedida")
            StDatos.IDUdMedidaB = data("IDUDInterna")
            data("Factor") = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, StDatos, services)
        End If
    End Sub

    <Task()> Public Shared Sub AsignarCantidades(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("QPrevista")) = 0 Then data("QPrevista") = 0
        data("QInterna") = xRound(data("Factor") * data("QPrevista"), 2)
    End Sub

    <Task()> Public Shared Function GeneraLineasHistorico(ByVal data As DataRow, ByVal services As ServiceProvider) As DataTable
        Dim dtHistorico As DataTable = New HistProgramaCompra().Filter(New NoRowsFilterItem())
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
                If drRowLineaPrograma("QPrevista", DataRowVersion.Default) <> drRowLineaPrograma("QPrevista", DataRowVersion.Original) OrElse _
                   drRowLineaPrograma("FechaEntrega", DataRowVersion.Default) <> drRowLineaPrograma("FechaEntrega", DataRowVersion.Original) Then
                    drRowHistorico = dtHistorico.NewRow
                    drRowHistorico("IDHistProg") = AdminData.GetAutoNumeric
                    drRowHistorico("IDLineaPrograma") = drRowLineaPrograma("IDLineaPrograma")
                    drRowHistorico("IDPrograma") = drRowLineaPrograma("IDPrograma")
                    drRowHistorico("FechaEntregaAntigua") = drRowLineaPrograma("FechaEntrega", DataRowVersion.Original)
                    drRowHistorico("FechaEntregaNueva") = drRowLineaPrograma("FechaEntrega", DataRowVersion.Default)
                    drRowHistorico("CantidadAntigua") = drRowLineaPrograma("QPrevista", DataRowVersion.Original)
                    drRowHistorico("CantidadNueva") = drRowLineaPrograma("QPrevista", DataRowVersion.Default)
                    drRowHistorico("ModificacionCantidades") = drRowLineaPrograma("QPrevista", DataRowVersion.Default) - drRowLineaPrograma("QPrevista", DataRowVersion.Original)
                    drRowHistorico("AlteracionFecha") = CInt(DateDiff(DateInterval.Day, drRowLineaPrograma("FechaEntrega", DataRowVersion.Original), drRowLineaPrograma("FechaEntrega", DataRowVersion.Default)))
                    drRowHistorico("FechaModificacionAudi") = Today
                    drRowHistorico("UsuarioAudi") = AdminData.GetSessionInfo.UserName
                End If
        End Select
        If Not IsNothing(drRowHistorico) Then dtHistorico.Rows.Add(drRowHistorico)
        BusinessHelper.UpdateTable(dtHistorico)
    End Function

#End Region

End Class