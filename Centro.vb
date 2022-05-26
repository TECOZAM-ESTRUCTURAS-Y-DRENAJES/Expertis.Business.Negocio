Public Class CentroInfo
    Inherits ClassEntityInfo

    Public IDCentro As String
    Public DescCentro As String
    Public TipoCentro As enumcentroTipoCentro
    Public IDSeccion As String
    Public FactorHombre As Double
    Public Capacidad As Double
    Public TasaPreparacionA As Double
    Public TasaEjecucionA As Double
    Public TasaManoObraA As Double
    Public TasaPreparacionB As Double
    Public TasaEjecucionB As Double
    Public TasaManoObraB As Double
    Public Bloqueado As Boolean
    Public Critico As Boolean
    Public Programable As Boolean
    Public FactorAfectaTP As Boolean
    Public Rendimiento As Double
    Public IDCaptura As String
    Public TiempoParada As Double
    Public IDUdMedida As String
    Public UdTiempo As Integer
    Public ProgVisible As Boolean
    Public ProgOrden As Integer
    Public NumeroIncripcionROMA As String
    Public FechaAdquisicion As Date
    Public FechaUltimaInspeccion As Date


    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Sub New(ByVal IDCentro As String)
        MyBase.New()
        Me.Fill(IDCentro)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        If Length(PrimaryKey(0)) = 0 Then Exit Sub
        Dim dtCentroInfo As DataTable = New Centro().SelOnPrimaryKey(PrimaryKey(0))
        If dtCentroInfo.Rows.Count > 0 Then
            Me.Fill(dtCentroInfo.Rows(0))
            If Length(dtCentroInfo.Rows(0)("UdTiempo")) = 0 Then
                Me.UdTiempo = enumstdUdTiempo.Horas
            End If
        Else
            ApplicationService.GenerateError("El Centro | no existe.", Quoted(PrimaryKey(0)))
        End If
    End Sub

End Class

Public Class Centro

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroCentro"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("TipoCentro")) = 0 Then ApplicationService.GenerateError("El Tipo de Centro es obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim DtCentro As DataTable = New Centro().SelOnPrimaryKey(data("IDCentro"))
            If Not DtCentro Is Nothing AndAlso DtCentro.Rows.Count > 0 Then
                ApplicationService.GenerateError("El centro introducido ya existe.")
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        'updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.UpdateEntityRow)
        'updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoActualizado)
        'updateProcess.AddTask(Of DataRow)(AddressOf CentroActualizarTasaUp)
    End Sub

    '<Task()> Public Shared Sub CentroActualizarTasaUp(ByVal data As DataRow, ByVal services As ServiceProvider)
    '    ProcessServer.ExecuteTask(Of DataTable)(AddressOf CentroActualizarTasa, data.Table, services)
    'End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub CentroActualizarTasa(ByVal data As DataTable, ByVal services As ServiceProvider)
        Dim VIEW_NAME As String = "vNegCentroTasaTotal"
        Dim dtTasa As DataTable
        Dim _filter As New Filter

        If Not data Is Nothing Then
            data.AcceptChanges()
            For Each dr As DataRow In data.Select
                _filter.Clear()
                _filter.Add("IDCentro", dr("IDCentro"))
                dtTasa = AdminData.Filter(VIEW_NAME, , _filter.Compose(New AdoFilterComposer))
                If Not dtTasa Is Nothing And dtTasa.Rows.Count > 0 Then
                    dr("TasaEjecucionA") = dtTasa.Rows(0)("TasaEjecucionA")
                    dr("TasaPreparacionA") = dtTasa.Rows(0)("TasaPreparacionA")
                    dr("TasaManoObraA") = dtTasa.Rows(0)("TasaManoObraA")
                    dr("TasaEjecucionB") = dtTasa.Rows(0)("TasaEjecucionB")
                    dr("TasaPreparacionB") = dtTasa.Rows(0)("TasaPreparacionB")
                    dr("TasaManoObraB") = dtTasa.Rows(0)("TasaManoObraB")
                Else
                    dr("TasaEjecucionA") = 0
                    dr("TasaPreparacionA") = 0
                    dr("TasaManoObraA") = 0
                    dr("TasaEjecucionB") = 0
                    dr("TasaPreparacionB") = 0
                    dr("TasaManoObraB") = 0
                End If
            Next
            BusinessHelper.UpdateTable(data)
        End If
    End Sub

    <Task()> Public Shared Function ValidaCentro(ByVal data As String, ByVal services As ServiceProvider) As DataTable
        Dim dt As DataTable = New Centro().SelOnPrimaryKey(data)
        If IsNothing(dt) OrElse dt.Rows.Count = 0 Then ApplicationService.GenerateError("El Centro | no existe.", data)
        Return dt
    End Function

    <Serializable()> _
    Public Class DatosCentroTasa
        Public IDCentro As String
        Public Fecha As DateTime
        Public TasaEjecucion As Double
        Public TasaPreparacion As Double
        Public TipoCosteFV As enumtcfvTipoCoste
        Public TipoCosteDI As enumtcdiTipoCoste
        Public Fiscal As Boolean

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDCentro As String, ByVal Fecha As DateTime)
            Me.IDCentro = IDCentro
            Me.Fecha = Fecha
        End Sub
    End Class

    <Task()> Public Shared Function ObtenerTasaCentro(ByVal data As DatosCentroTasa, ByVal services As ServiceProvider) As DatosCentroTasa
        Dim FilCentroTasa As New Filter
        FilCentroTasa.Add("IDCentro", FilterOperator.Equal, data.IDCentro)
        FilCentroTasa.Add("FechaDesde", FilterOperator.LessThanOrEqual, data.Fecha)
        FilCentroTasa.Add("FechaHasta", FilterOperator.GreaterThanOrEqual, data.Fecha)
        Dim DtCentroTasa As DataTable = New CentroTasa().Filter(FilCentroTasa)
        If Not DtCentroTasa Is Nothing AndAlso DtCentroTasa.Rows.Count > 0 Then
            data.TasaEjecucion = DtCentroTasa.Rows(0)("EjecucionValorA")
            data.TasaPreparacion = DtCentroTasa.Rows(0)("PreparacionValorA")
            data.TipoCosteFV = DtCentroTasa.Rows(0)("TipoCosteFV")
            data.TipoCosteDI = DtCentroTasa.Rows(0)("TipoCosteDI")
            data.Fiscal = DtCentroTasa.Rows(0)("Fiscal")
        End If
        Return data
    End Function

#End Region

End Class