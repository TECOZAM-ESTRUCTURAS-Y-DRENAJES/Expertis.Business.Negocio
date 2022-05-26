Public Class RemesaCobroFacturaCompra
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbRemesaCobroFacturaCompra"

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
        If Nz(data("IDRemesa"), 0) = 0 Then
            ApplicationService.GenerateError("Debe indicar la Remesa.")
        End If
        If Nz(data("IDFacturaCompra"), 0) = 0 Then
            ApplicationService.GenerateError("Debe indicar la Factura.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarRegistroExistente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim RCFC As New RemesaCobroFacturaCompra
            Dim dt As DataTable = RCFC.SelOnPrimaryKey(data("IDRemesa"), data("IDFacturaCompra"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El registro intoducido ya existe.")
            End If
        End If
    End Sub

#End Region

    <Serializable()> _
   Public Class DataRemesaFC
        Public IDRemesa As Integer
        Public IDFactura As Integer
    End Class

    <Task()> Public Shared Sub AddRemesaFacturaCompra(ByVal data() As DataRemesaFC, ByVal services As ServiceProvider)
        Dim RCFC As New RemesaCobroFacturaCompra
        Dim dtRegistros As DataTable = RCFC.AddNew
        For Each gasto As DataRemesaFC In data
            Dim drNew As DataRow = dtRegistros.NewRow
            drNew("IDRemesa") = gasto.IDRemesa
            drNew("IDFacturaCompra") = gasto.IDFactura
            dtRegistros.Rows.Add(drNew)
        Next
        RCFC.Validate(dtRegistros)
        BusinessHelper.UpdateTable(dtRegistros)
    End Sub

End Class
