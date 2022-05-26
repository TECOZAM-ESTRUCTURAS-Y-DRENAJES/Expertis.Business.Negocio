Public Class CobroFacturaCompra
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbCobroFacturaCompra"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
    End Sub

#Region " Validate "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCamposObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarRegistroExistente)
    End Sub

    <Task()> Public Shared Sub ValidarCamposObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Nz(data("IDCobro"), 0) = 0 Then
            ApplicationService.GenerateError("Debe indicar el Cobro.")
        End If
        If Nz(data("IDFactura"), 0) = 0 Then
            ApplicationService.GenerateError("Debe indicar la Factura.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarRegistroExistente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim CFC As New CobroFacturaCompra
            Dim dt As DataTable = CFC.SelOnPrimaryKey(data("IDCobro"), data("IDFactura"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El registro intoducido ya existe.")
            End If
        End If
    End Sub

#End Region

    <Serializable()> _
    Public Class DataCobroFC
        Public IDCobro As Integer
        Public IDFactura As Integer
    End Class

    <Task()> Public Shared Sub AddCobroFacturaCompra(ByVal data() As DataCobroFC, ByVal services As ServiceProvider)
        Dim CFC As New CobroFacturaCompra
        Dim dtRegistros As DataTable = CFC.AddNew
        For Each gasto As DataCobroFC In data
            Dim drNew As DataRow = dtRegistros.NewRow
            drNew("IDCobro") = gasto.IDCobro
            drNew("IDFactura") = gasto.IDFactura
            dtRegistros.Rows.Add(drNew)
        Next
        CFC.Validate(dtRegistros)
        BusinessHelper.UpdateTable(dtRegistros)
    End Sub

End Class
