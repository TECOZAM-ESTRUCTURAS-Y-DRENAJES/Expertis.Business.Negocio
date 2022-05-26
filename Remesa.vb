Public Class Remesa

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbRemesa"

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function FechaValor(ByVal data As Integer, ByVal services As ServiceProvider) As Date
        If Length(data) > 0 Then
            Dim dt As DataTable = AdminData.Filter("vRemesaFechaValor", , "IDRemesa=" & data, "IDRemesa")
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                Return dt.Rows(0)("FechaValor")
            End If
        End If
    End Function

    <Task()> Public Shared Function FechaValorProv(ByVal data As String, ByVal services As ServiceProvider) As Date
        If Length(data) > 0 Then
            Dim dt As DataTable = AdminData.Filter("vRemesaProvFechaValor", , "IDProcess='" & data & "'")
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                Return dt.Rows(0)("FechaValor")
            End If
        End If
    End Function

#End Region

#Region " Delete "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ValidarDelGastosAsociados)
    End Sub

    <Task()> Public Shared Sub ValidarDelGastosAsociados(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim dtGastosRemesa As DataTable = New RemesaCobroFacturaCompra().Filter(New NumberFilterItem("IDRemesa", data("IDRemesa")))
        If dtGastosRemesa.Rows.Count > 0 Then
            ApplicationService.GenerateError("No se puede eliminar la Remesa {0}, tiene gastos asociados.", Quoted(data("IDRemesa")))
        End If
    End Sub

#End Region

End Class