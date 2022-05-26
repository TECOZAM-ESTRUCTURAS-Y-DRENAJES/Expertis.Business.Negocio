Public Class TipoArticulo

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroTipoArticulo"

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
        data("CriterioValoracion") = enumtaValoracion.taPrecioEstandar
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
        If Length(data("IDTipo")) = 0 Then ApplicationService.GenerateError("Tipo es un dato obligatorio.")
        If Length(data("DescTipo")) = 0 Then ApplicationService.GenerateError("La Descripción es un dato obligatorio.")
        If Nz((data("Fantasma")), False) And Nz((data("GestionStock")), False) Then
            ApplicationService.GenerateError("Las propiedades asignadas al tipo no son compatibles")
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
            Dim dt As DataTable = New TipoArticulo().SelOnPrimaryKey(data("IDTipo"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Código introducido ya existe.")
            End If
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
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarValoresVacios)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarCriterioValoracion)
    End Sub

    ''' <summary>
    ''' Asignar la información por defecto
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub AsignarValoresVacios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Or data.RowState = DataRowState.Modified Then
            If Length(data("CriterioValoracion")) = 0 Then data("CriterioValoracion") = enumtaValoracion.taPrecioEstandar
            If Nz(data("GestionStock"), False) Then
                If Length(data("RecalcularValoracion")) = 0 Then data("RecalcularValoracion") = New Parametro().RecalcularValoracion()
            End If
        End If
    End Sub

    ''' <summary>
    '''Modificar el criterio de valoración en los artículos del tipo modificado
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ActualizarCriterioValoracion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            If data("RecalcularValoracion") <> data("RecalcularValoracion", DataRowVersion.Original) Or _
                         data("CriterioValoracion") <> data("CriterioValoracion", DataRowVersion.Original) Then
                Dim f As New Filter
                f.Add(New StringFilterItem("IDTipo", data("IDTipo")))

                Dim a As New Articulo
                Dim dtArticulo As DataTable = a.Filter(f)

                If Not IsNothing(dtArticulo) AndAlso dtArticulo.Rows.Count > 0 Then
                    For Each dr As DataRow In dtArticulo.Rows
                        dr("RecalcularValoracion") = data("RecalcularValoracion")
                        dr("CriterioValoracion") = data("CriterioValoracion")
                    Next
                End If
                BusinessHelper.UpdateTable(dtArticulo)
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function ValidaTipoArticulo(ByVal strIDTipo As String, ByVal SERVICES As ServiceProvider) As DataTable
        Dim dt As DataTable = New TipoArticulo().SelOnPrimaryKey(strIDTipo)
        If IsNothing(dt) OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("El Tipo '|' no existe.", strIDTipo)
        End If
        Return dt
    End Function

#End Region

End Class