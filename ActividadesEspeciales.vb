Public Class ActividadesEspeciales

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroCAE"

#End Region

#Region "Eventos RegiserValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCamposObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarRegistroExistente)
    End Sub

    ''' <summary>
    ''' Método que valida los datos mínimos que debemos tener
    ''' </summary>
    ''' <param name="data">DataRow con el registro referente a la Actividad Especial</param>
    ''' <param name="services">Objeto para compartir información</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarCamposObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCAE")) = 0 Then ApplicationService.GenerateError("El CAE es obligatorio.")
        If Length(data("DescCAE")) = 0 Then ApplicationService.GenerateError("La Descripción es obligatoria.")
    End Sub

    ''' <summary>
    ''' Método que valida si el registro está dado de alta en el sistema
    ''' </summary>
    ''' <param name="data">DataRow con el registro referente a la Actividad Especial</param>
    ''' <param name="services">Objeto para compartir información</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarRegistroExistente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dtA As DataTable = New ActividadesEspeciales().SelOnPrimaryKey(data("IDCAE"))
            If Not IsNothing(dtA) AndAlso dtA.Rows.Count > 0 Then
                ApplicationService.GenerateError("El CAE ya existe.")
            End If
        End If
    End Sub

#End Region

End Class