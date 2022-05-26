Public Class TipoCobro

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroTipoCobro"

#End Region

#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarSistema)
    End Sub

    <Task()> Public Shared Sub ComprobarSistema(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("Sistema") Then ApplicationService.GenerateError("No se puede realizar esta operación sobre un Estado del Sistema.")
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarTipoCobro)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarDescripcion)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarDuplicados)
    End Sub

    <Task()> Public Shared Sub ComprobarTipoCobro(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDTipoCobro")) = 0 Then ApplicationService.GenerateError("El Tipo Cobro es un dato obligatorio.")
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarDescripcion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("DescTipo")) = 0 Then ApplicationService.GenerateError("La Descripción del Tipo Cobro es un dato obligatorio.")
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarDuplicados(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim DtDatos As DataTable = New TipoCobro().SelOnPrimaryKey(data("IDTipoCobro"))
            If Not DtDatos Is Nothing AndAlso DtDatos.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Tipo de Cobro ya existe.")
            End If
        End If
    End Sub

#End Region

End Class