Public Class RutaUtillaje

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbRutaUtillaje"

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
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarUtillaje)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarRuta)
    End Sub

    ''' <summary>
    ''' Comprobar que el código y la descripción no sean nulos
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("IDRutaOp").ToString.Trim.Length = 0 Or data("IDUtillaje").ToString.Trim.Length = 0 Then
            ApplicationService.GenerateError("Introduzca la ruta y el utillaje.")
        End If
    End Sub

    ''' <summary>
    ''' Comprobar que no exista la clave
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If data("IDRutaOp").ToString.Trim.Length > 0 And _
                  data("IDUtillaje").ToString.Trim.Length > 0 Then
                Dim dt As DataTable = New RutaUtillaje().SelOnPrimaryKey(data("IDRutaOp"), data("IDUtillaje"))
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    ApplicationService.GenerateError("Ya existe este utillaje para la ruta actual.")
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Comprobar que el utillaje es válido 
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarUtillaje(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dtArticulo As DataTable = New Articulo().SelOnPrimaryKey(data("IDUtillaje"))
        If dtArticulo Is Nothing OrElse dtArticulo.Rows.Count = 0 Then
            ApplicationService.GenerateError("El utillaje no existe.")
        End If
    End Sub

    ''' <summary>
    ''' Comprobar que la ruta existe
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarRuta(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dtRuta As DataTable = New Ruta().SelOnPrimaryKey(data("IDRutaOp"))
        If dtRuta Is Nothing OrElse dtRuta.Rows.Count = 0 Then
            ApplicationService.GenerateError("La ruta no existe.")
        End If
    End Sub

#End Region

End Class