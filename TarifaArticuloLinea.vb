Public Class TarifaArticuloLinea

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbTarifaArticuloLinea"

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
        If Length(data("IDTarifa")) = 0 Then ApplicationService.GenerateError("El código de Tarifa es obligatorio.")
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El campo Artículo es obligatorio.")
    End Sub

    ''' <summary>
    ''' Comprobar que no exista la clave
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New TarifaArticuloLinea().SelOnPrimaryKey(data("IdTarifa"), data("IDArticulo"), data("QDesde"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("La cantidad introducida ya tiene precio.")
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf TratarCantidad)
        updateProcess.AddTask(Of DataRow)(AddressOf TarifaArticulo.AplicarDecimales)
    End Sub

    <Task()> Public Shared Sub TratarCantidad(ByVal data As DataRow, ByVal services As ServiceProvider)

        If Length(data("QDesde")) = 0 Then data("QDesde") = 0
        If data("QDesde") = 0 Then
            Dim dtTA As DataTable = New TarifaArticulo().SelOnPrimaryKey(data("IdTarifa"), data("IDArticulo"))
            If Not dtTA Is Nothing AndAlso dtTA.Rows.Count > 0 Then
                dtTA.Rows(0)("Precio") = data("Precio")
                dtTA.Rows(0)("PVP") = data("PVP")
                dtTA.Rows(0)("Dto1") = data("Dto1")
                dtTA.Rows(0)("Dto2") = data("Dto2")
                dtTA.Rows(0)("Dto3") = data("Dto3")
            End If
        End If
    End Sub


#End Region

#Region "Eventos RegisterBusinessRules"

    ''' <summary>
    ''' Reglas de negocio. Lista de tareas asociadas a cambios.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks>Solo se enstablece la lista en este punto no se ejecutan</remarks>
    Public Overrides Function GetBusinessRules() As Solmicro.Expertis.Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("PVP", AddressOf TarifaArticulo.CambioPVPPrecio)
        oBRL.Add("Precio", AddressOf TarifaArticulo.CambioPVPPrecio)
        Return oBRL
    End Function

#End Region

End Class