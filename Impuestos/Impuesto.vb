Public Class Impuesto
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

#Region " Constructor "

    Private Const cnEntidad As String = "tbMaestroImpuesto"

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

#End Region

#Region " RegisterValidateTaks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClavePrimaria)
        '   validateProcess.AddTask(Of DataRow)(AddressOf General.ValidarRangoPorcentaje)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDImpuesto")) = 0 Then ApplicationService.GenerateError("El identificador del Impuesto es un dato obligatorio.")
        If Length(data("DescImpuesto")) = 0 Then ApplicationService.GenerateError("La Descripción es un dato obligatorio.")
        '       If Length(data("BaseImpuesto")) = 0 Then ApplicationService.GenerateError("La Base sobre la que aplicar el Impuesto es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New Impuesto().SelOnPrimaryKey(data("IDImpuesto"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Código introducido ya existe.")
            End If
        End If
    End Sub

#End Region


End Class
