Public Class Tasa

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Private Const cnEntidad As String = "tbMaestroTasa"

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    ''' <summary>
    ''' Relación de tareas asociadas a la validación 
    ''' </summary>
    ''' <param name="validateProcess">Proceso en el que se registran las tareas de edición</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClavePrimaria)
    End Sub

    ''' <summary>
    ''' Comprobar que el código y la descripción no sean nulos
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTasa")) = 0 Then ApplicationService.GenerateError("Tipo de tasa es un dato obligatorio.")
        If Length(data("ValorTasaA")) = 0 Then ApplicationService.GenerateError("El valor de la Tasa A es obligatoria.")
    End Sub

    ''' <summary>
    ''' Comprobar que no exista la clave
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New Tasa().SelOnPrimaryKey(data("IDTasa"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Código introducido ya existe.")
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarTipos)
        updateProcess.AddTask(Of DataRow)(AddressOf TratarMonedaB)
    End Sub
    <Task()> Public Shared Sub AsignarTipos(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("TipoCosteFV")) = 0 Then data("TipoCosteFV") = 0
        If Length(data("TipoCosteDI")) = 0 Then data("TipoCosteDI") = 0
    End Sub
    <Task()> Public Shared Sub TratarMonedaB(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim ClsMoneda As New Moneda
        Dim DtMonedaA As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf Moneda.ObtenerMonedaA, Nothing, services)
        Dim DtMonedaB As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf Moneda.ObtenerMonedaB, Nothing, services)
        Dim Cambios As MonedaInfo
        Dim DtData As DataTable
        Dim DblCambioB As Double
        Dim IntNumDecB As Integer
        If Not DtMonedaA Is Nothing AndAlso DtMonedaA.Rows.Count > 0 Then
            Dim StDatos As New Moneda.DatosObtenerMoneda
            StDatos.IDMoneda = DtMonedaA.Rows(0)("IDMoneda")
            Cambios = ProcessServer.ExecuteTask(Of Moneda.DatosObtenerMoneda, MonedaInfo)(AddressOf Moneda.ObtenerMoneda, StDatos, services)
            DblCambioB = Cambios.CambioB
        End If
        If Not DtMonedaB Is Nothing AndAlso DtMonedaB.Rows.Count > 0 Then
            IntNumDecB = DtMonedaB.Rows(0)("NDecimalesImp")
        End If
        data("ValorTasaB") = xRound(data("ValorTasaA") * DblCambioB, IntNumDecB)
    End Sub

#End Region

End Class