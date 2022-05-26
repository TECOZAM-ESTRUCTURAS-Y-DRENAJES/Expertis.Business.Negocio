Public Class TipoArticuloFoto

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbTipoArticuloFoto"

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
        If Length(data("IDTipoArticulo")) = 0 Then ApplicationService.GenerateError("Tipo de artículo es un dato obligatorio.")
        If IsDBNull(data("Foto")) OrElse data("Foto") Is Nothing OrElse CType(data("Foto"), Array).Length = 0 Then
            ApplicationService.GenerateError("Debe incluir una foto.")
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
            Dim dt As DataTable = New TipoArticuloFoto().SelOnPrimaryKey(data("IDTipoArticulo"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("Ya existe una foto para el Tipo de Artículo {0}.", Quoted(data("IDTipoArticulo")))
            End If
        End If
    End Sub

#End Region

End Class