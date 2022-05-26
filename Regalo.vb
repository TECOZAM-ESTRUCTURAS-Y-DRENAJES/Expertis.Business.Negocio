Public Class Regalo

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbRegalo"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarClave)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarDuplicados)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarClaveArticulo)
    End Sub

    <Task()> Public Shared Sub ComprobarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("Introduzca el código del regalo.")
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarDuplicados(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim DtDatos As DataTable = New Regalo().SelOnPrimaryKey(data("IDArticulo"))
        If Not DtDatos Is Nothing AndAlso DtDatos.Rows.Count > 0 Then ApplicationService.GenerateError("Ya existe un regalo con esa clave.")
    End Sub

    <Task()> Public Shared Sub ComprobarClaveArticulo(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim DtDatos As DataTable = New Articulo().SelOnPrimaryKey(data("IDArticulo"))
        If DtDatos Is Nothing OrElse DtDatos.Rows.Count = 0 Then ApplicationService.GenerateError("El Regalo no está dado de alta como Artículo.")
    End Sub

#End Region

End Class