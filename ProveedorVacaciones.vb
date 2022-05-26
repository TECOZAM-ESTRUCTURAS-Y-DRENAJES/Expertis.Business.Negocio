Public Class ProveedorVacaciones

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbProveedorVacaciones"

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
    End Sub

    ''' <summary>
    ''' Comprobar que las fechas y el proveedor estén rellenas
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("FechaDesde")) = 0 Or Length(data("FechaHasta")) = 0 Then
            ApplicationService.GenerateError("La Fecha Desde y Fecha Hasta son obligatorias.")
        End If
        If Length(data("FechaDivision")) = 0 AndAlso Length(data("FechaAlternativa")) = 0 Then
            ApplicationService.GenerateError("Debe indicar Fecha Alternativa o Fecha División.")
        End If
        If Length(data("IDProveedor")) = 0 Then ApplicationService.GenerateError("El Proveedor es un dato obligatorio.")

    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarPrimaryKey)
    End Sub

    <Task()> Public Shared Sub AsignarPrimaryKey(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            data("IDVacacion") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

#Region "Eventos GetBusinessRules"

    ''' <summary>
    ''' Reglas de negocio. Lista de tareas asociadas a cambios.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Solo se enstablece la lista en este punto no se ejecutan</remarks>
    Public Overrides Function GetBusinessRules() As Solmicro.Expertis.Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("FechaDesde", AddressOf CambioFechasVacaciones)
        oBRL.Add("FechaHasta", AddressOf CambioFechasVacaciones)
        Return oBRL
    End Function

    ''' <summary>
    ''' Validar que la fecha desde es menor que la fecha hasta
    ''' </summary>
    ''' <param name="data">Estructura con la información necesaria</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub CambioFechasVacaciones(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Select Case data.ColumnName
            Case "FechaDesde", "FechaHasta"
                If IsDate(data.Value) Then
                    data.Current(data.ColumnName) = data.Value
                    If Length(data.Current("FechaDesde")) > 0 And Length(data.Current("FechaHasta")) > 0 Then
                        If data.Current("FechaHasta") < data.Current("FechaDesde") Then
                            ApplicationService.GenerateError("La Fecha Hasta debe ser mayor que la Fecha Desde.")
                        End If
                    End If
                End If
        End Select
    End Sub

#End Region

End Class