Public Class TipoLinea

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroTipoLinea"

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
        If Length(data("IDTipoLinea")) = 0 Then ApplicationService.GenerateError("Tipo Línea es un dato obligatorio.")
        If Length(data("DescTipoLinea")) = 0 Then ApplicationService.GenerateError("La Descripción es un dato obligatorio.")
    End Sub

    ''' <summary>
    ''' Comprobar que no exista la clave
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New TipoLinea().SelOnPrimaryKey(data("IDTipoLinea"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Código introducido ya existe.")
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf TratarPredeterminado)
    End Sub

    <Task()> Public Shared Sub TratarPredeterminado(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter(FilterUnionOperator.And)
        f.Add(New StringFilterItem("IDTipoLinea", FilterOperator.NotEqual, data("IDTipoLinea")))
        f.Add(New BooleanFilterItem("Predeterminada", FilterOperator.Equal, True))

        Dim dtTL As DataTable = New TipoLinea().Filter(f)

        If IsNothing(dtTL) OrElse dtTL.Rows.Count = 0 Then
            data("Predeterminada") = True
        Else
            If IsDBNull(data("Predeterminada")) Then data("Predeterminada") = False
            If data("Predeterminada") Then
                dtTL.Rows(0)("Predeterminada") = False
                BusinessHelper.UpdateTable(dtTL)
            ElseIf data.RowState = DataRowState.Modified AndAlso data("Predeterminada") <> data("Predeterminada", DataRowVersion.Original) AndAlso dtTL.Rows.Count = 1 Then
                data("Predeterminada") = True
            End If
        End If

    End Sub

#End Region

#Region "Eventos RegisterDeleteTasks"

    ''' <summary>
    ''' Relación de tareas asociadas al proceso de borrado
    ''' </summary>
    ''' <param name="deleteProcess">Proceso en el que se registran las tareas de borrado</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ValidarDelete)
    End Sub

    ''' <summary>
    ''' Borrado de tipos de línea
    ''' </summary>
    ''' <param name="data">Registro del tipo a borrar</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarDelete(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Not IsNothing(data) Then
            If data("Sistema") Then ApplicationService.GenerateError("No se puede borrar un Tipo de Línea de Sistema.")
            Dim blnPredeterminadoBorrado As Boolean = data("Predeterminada")
            '
            '  MyBase.Delete(data)
            If blnPredeterminadoBorrado Then
                Dim dtTL As DataTable = New TipoLinea().Filter()
                If Not IsNothing(dtTL) AndAlso dtTL.Rows.Count > 0 Then
                    dtTL.Rows(0)("Predeterminada") = True
                    BusinessHelper.UpdateTable(dtTL)
                Else
                    ApplicationService.GenerateError("No se puede borrar el único Tipo de Línea que existe.")
                End If
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    ''' <summary>
    ''' Identificar tipo de línea por defecto
    ''' </summary>
    ''' <param name="data">objeto</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Function TipoLineaPorDefecto(ByVal obj As Object, ByVal services As ServiceProvider) As String
        Dim strTipoLinea As String

        Dim dtTipoLinea As DataTable = New TipoLinea().Filter("IDTipoLinea", "Predeterminada <> 0")
        If Not dtTipoLinea Is Nothing AndAlso dtTipoLinea.Rows.Count > 0 Then
            strTipoLinea = dtTipoLinea.Rows(0)("IDTipoLinea")
        End If

        Return strTipoLinea
    End Function

    ''' <summary>
    ''' Identificar tipo de línea de regalo  por defecto
    ''' </summary>
    ''' <param name="data">objeto</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Function TipoLineaRegalo(ByVal obj As Object, ByVal services As ServiceProvider) As String
        Dim strTipoLinea As String

        Dim dtTipoLinea As DataTable = New TipoLinea().Filter("IDTipoLinea", "Regalo <> 0")
        If Not dtTipoLinea Is Nothing AndAlso dtTipoLinea.Rows.Count > 0 Then
            strTipoLinea = dtTipoLinea.Rows(0)("IDTipoLinea")
        End If

        Return strTipoLinea
    End Function

#End Region

End Class