Public Class TarifaEstado

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroTarifaEstado"

#End Region

#Region "Eventos RegisterValidateTasks"

    ''' <summary>
    ''' Relaci�n de tareas asociadas a la validaci�n 
    ''' </summary>
    ''' <param name="validateProcess">Proceso en el que se registran las tareas de edici�n</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClavePrimaria)
    End Sub

    ''' <summary>
    ''' Comprobar que el c�digo y la descripci�n no sean nulos
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Informaci�n compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDEstado")) = 0 Then ApplicationService.GenerateError("El c�digo de estado es obligatorio.")
        If Length(data("DescEstado")) = 0 Then ApplicationService.GenerateError("La descripci�n es obligatoria.")
    End Sub

    ''' <summary>
    ''' Comprobar que no exista la clave
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Informaci�n compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New TarifaEstado().SelOnPrimaryKey(data("IDEstado"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El C�digo introducido ya existe.")
            End If
        End If
    End Sub

#End Region

End Class