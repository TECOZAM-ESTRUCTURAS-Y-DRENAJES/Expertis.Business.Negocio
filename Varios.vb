Public Class Varios

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroVarios"

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
        If Length(data("IDVarios")) = 0 Then ApplicationService.GenerateError("Varios es un dato obligatorio.")
        If Length(data("DescVarios")) = 0 Then ApplicationService.GenerateError("La Descripción es un dato obligatorio.")
    End Sub

    ''' <summary>
    ''' Comprobar que no exista la clave
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New Varios().SelOnPrimaryKey(data("IDVarios"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Código introducido ya existe.")
            End If
        End If
    End Sub

    <Task()> Public Shared Function ValidaVarios(ByVal strIDVarios As String, ByVal services As ServiceProvider) As DataTable
        If Length(strIDVarios) Then
            Dim dt As DataTable = New Varios().SelOnPrimaryKey(strIDVarios)
            If IsNothing(dt) OrElse dt.Rows.Count = 0 Then
                ApplicationService.GenerateError("El Varios | no existe.", strIDVarios)
            End If

            Return dt
        End If
    End Function

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarTipos)
    End Sub
    <Task()> Public Shared Sub AsignarTipos(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("TipoCosteFV")) = 0 Then data("TipoCosteFV") = 0
        If Length(data("TipoCosteDI")) = 0 Then data("TipoCosteDI") = 0
    End Sub
 
#End Region

End Class