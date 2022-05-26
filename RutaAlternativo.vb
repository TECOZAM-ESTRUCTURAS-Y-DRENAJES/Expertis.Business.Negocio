Public Class RutaAlternativo

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbRutaAlternativo"

#End Region

#Region "Eventos RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf AsignarValoresPredeterminados)
    End Sub
    
    ''' <summary>
    ''' Asignar la información por defecto
    ''' </summary>
    ''' <param name="data">Registro Nuevo</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub AsignarValoresPredeterminados(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim ClsParam As New Parametro
        Dim Param As Integer = ClsParam.UdTiempoPred
        data("TipoOperacion") = enumtrTipoOperacion.trInterna '0
        data("UdTiempoPrep") = Param
        data("UdTiempoEjec") = Param
        data("UdTiempo") = Param
        data("FactorProduccion") = 1
        data("Factorhombre") = 1
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
    End Sub

    ''' <summary>
    ''' Comprobar que el código y la descripción no sean nulos
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCentro")) = 0 Then ApplicationService.GenerateError("El centro es un dato obligatorio.")
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarPrimaryKey)
    End Sub

    <Task()> Public Shared Sub AsignarPrimaryKey(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            data("ID") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

End Class