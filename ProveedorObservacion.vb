Public Class ProveedorObservacion

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Private Const cnEntidad As String = "tbProveedorObservacion"

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
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIDProveedor)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIDObservacion)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarObservacion)
    End Sub

    ''' <summary>
    ''' Comprobar que el código y la descripción no sean nulos
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarIDProveedor(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDProveedor")) = 0 Then ApplicationService.GenerateError("El Proveedor es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarIDObservacion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDObservacion")) = 0 Then ApplicationService.GenerateError("Observación es un dato obligatorio.")
    End Sub

    ''' <summary>
    ''' Comprobar que no exista la clave
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarObservacion(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim ofilter As New Filter
        If data.RowState = DataRowState.Modified Then
            ofilter.Add(New NumberFilterItem("IDProveedorObservacion", FilterOperator.NotEqual, data("IDProveedorObservacion")))
        End If
        If data.RowState = DataRowState.Modified Or data.RowState = DataRowState.Added Then
            ofilter.Add(New StringFilterItem("IDProveedor", data("IDProveedor")))
            ofilter.Add(New StringFilterItem("IDObservacion", data("IDObservacion")))
        End If

        Dim dt As DataTable = New ProveedorObservacion().Filter(ofilter)
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            ApplicationService.GenerateError("Ya existe esta observación para el proveedor actual.")
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    ''' <summary>
    ''' Relación de tareas asociadas a la modificación 
    ''' </summary>
    ''' <param name="updateProcess">Proceso en el que se registran las tareas de edición</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarPrimaryKey)
    End Sub

    ''' <summary>
    ''' Asignar la información por defecto
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub AsignarPrimaryKey(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDProveedorObservacion")) = 0 Then data("IDProveedorObservacion") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function ObtenerEntidades(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Return New BE.DataEngine().Filter("frmEntidadObservacion", "*", "", "Entidad")
    End Function

#End Region

End Class