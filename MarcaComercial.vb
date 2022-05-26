Public Class MarcaComercial

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroMarcaComercial"

#End Region

#Region "Eventos ValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosDuplicados)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDMarcaComercial")) = 0 Then ApplicationService.GenerateError("La marca comercial es un dato obligatorio.")
        If Length(data("DescMarcaComercial")) = 0 Then ApplicationService.GenerateError("La descripción es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarDatosDuplicados(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim DtMarca As DataTable = New MarcaComercial().SelOnPrimaryKey(data("IDMarcaComercial"))
            If Not DtMarca Is Nothing AndAlso DtMarca.Rows.Count > 0 Then
                ApplicationService.GenerateError("La marca comercial | ya existe en la base de datos.", data("IDMarcaComercial"))
            End If
        End If
    End Sub

#End Region

End Class