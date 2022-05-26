Public Class TipoEtiqueta

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroTipoEtiqueta"

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
        If Length(data("IDTipoEtiqueta")) = 0 Then ApplicationService.GenerateError("Tipo etiqueta es un dato obligatorio.")
        If Length(data("DescTipoEtiqueta")) = 0 Then ApplicationService.GenerateError("La Descripción es un dato obligatorio.")
        If Length(data("Informe")) = 0 Then
            ApplicationService.GenerateError("El Informe es Obligatorio.")
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
            Dim dt As DataTable = New TipoEtiqueta().SelOnPrimaryKey(data("IDTipoEtiqueta"))
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
        If Length(data("Filas")) = 0 Then data("Filas") = 1
        If Length(data("Columnas")) = 0 Then data("Columnas") = 1
    End Sub

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("PredeterminadaContenedor", AddressOf CambiarPredeterminados)
        oBrl.Add("PredeterminadaCaja", AddressOf CambiarPredeterminados)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambiarPredeterminados(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If data.ColumnName = "PredeterminadaContenedor" Then data.Current("PredeterminadaCaja") = False
        If data.ColumnName = "PredeterminadaCaja" Then data.Current("PredeterminadaContenedor") = False
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function TiposDeEtiqueta(ByVal data As String, ByVal services As ServiceProvider) As DataTable
        Dim Fil As New Filter
        Dim FilOR As New Filter(FilterUnionOperator.Or)
        Dim Fil2 As New Filter
        If Length(data) > 0 Then
            Dim dtCliente As DataTable = New Cliente().SelOnPrimaryKey(data)
            If Not dtCliente Is Nothing AndAlso dtCliente.Rows.Count > 0 Then
                If Length(dtCliente.Rows(0)("IDTipoEtiquetaContenedor")) > 0 Then
                    Fil.Add("IDTipoEtiqueta", FilterOperator.Equal, dtCliente.Rows(0)("IDTipoEtiquetaContenedor"), FilterType.String)
                Else : Fil.Add("PredeterminadaContenedor", FilterOperator.Equal, 1)
                End If
                If Length(dtCliente.Rows(0)("IDTipoEtiquetaCaja")) > 0 Then
                    Fil2.Add("IDTipoEtiqueta", FilterOperator.Equal, dtCliente.Rows(0)("IDTipoEtiquetaCaja"), FilterType.String)
                Else : Fil2.Add("PredeterminadaCaja", FilterOperator.Equal, 1)
                End If
                FilOR.Add(Fil)
                FilOR.Add(Fil2)

            End If
        End If
        If FilOR.Count = 0 Then
            FilOR.Add("PredeterminadaContenedor", FilterOperator.Equal, 1)
            FilOR.Add("PredeterminadaCaja", FilterOperator.Equal, 1)
        End If
        Return New TipoEtiqueta().Filter(FilOR)
    End Function

#End Region

End Class