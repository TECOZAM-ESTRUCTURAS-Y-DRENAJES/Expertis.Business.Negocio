Public Class HistoricoPreciosProveedor

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbHistoricoPreciosProveedor"

#End Region

#Region "Eventos Entidad"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClave)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarPrecio)
    End Sub

    <Task()> Public Shared Sub ValidarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDProveedor")) = 0 Then ApplicationService.GenerateError("El Proveedor es un dato obligatorio.")
        If Length(data("FechaDesde")) = 0 Then ApplicationService.GenerateError("La Fecha Desde es un dato obligatorio.")
        If Length(data("FechaHasta")) = 0 Then ApplicationService.GenerateError("La Fecha Hasta es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim FilHist As New Filter
            FilHist.Add("IDProveedor", FilterOperator.Equal, data("IDProveedor"))
            FilHist.Add("IDArticulo", FilterOperator.Equal, data("IDArticulo"))
            FilHist.Add("FechaDesde", FilterOperator.Equal, data("FechaDesde"))
            FilHist.Add("FechaHasta", FilterOperator.Equal, data("FechaHasta"))

            Dim DtHist As DataTable = New HistoricoPreciosProveedor().Filter(FilHist)
            If Not DtHist Is Nothing AndAlso DtHist.Rows.Count > 0 Then
                ApplicationService.GenerateError("El histórico introducido para las fechas introducidas ya existe en el histórico.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarPrecio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("Precio")) = 0 OrElse data("Precio") = 0 Then ApplicationService.GenerateError("El Precio es un dato obligatorio.")
    End Sub

#End Region

End Class