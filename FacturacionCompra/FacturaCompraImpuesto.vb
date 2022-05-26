Public Class FacturaCompraImpuesto

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

#Region " Constructor "

    Private Const cnEntidad As String = "tbFacturaCompraImpuesto"

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

#End Region

#Region " BusinessRules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("IDImpuesto", AddressOf NegocioGeneral.CambioImpuesto)
        oBRL.Add("Importe", AddressOf NegocioGeneral.CambioImporte)
        oBRL.Add("Porcentaje", AddressOf NegocioGeneral.CambioPorcentaje)
        oBRL.Add("Valor", AddressOf NegocioGeneral.CambioValor)
        Return oBRL
    End Function

#End Region

#Region " RegisterAddNewTasks "

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificador)
    End Sub

    <Task()> Public Shared Sub AsignarIdentificador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDLineaImpuesto")) = 0 Then data("IDLineaImpuesto") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

#Region " RegisterValidateTaks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarRegistroExistente)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDFactura")) = 0 Then ApplicationService.GenerateError("No se ha indicado el Identificador de la Factura.")
        If Length(data("IDLineaFactura")) = 0 Then ApplicationService.GenerateError("No se ha indicado el Identificador de la línea de la Factura.")
        If Length(data("IDImpuesto")) = 0 Then ApplicationService.GenerateError("El Identificador del impuesto es un dato obligatorio.")
        If Nz(data("Valor"), 0) = 0 Then ApplicationService.GenerateError("Debe indicar un Valor.")
    End Sub

    <Task()> Public Shared Sub ValidarRegistroExistente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim f As New Filter
            f.Add(New NumberFilterItem("IDLineaFactura", data("IDLineaFactura")))
            f.Add(New StringFilterItem("IDImpuesto", data("IDImpuesto")))
            Dim dt As DataTable = New FacturaCompraImpuesto().Filter(f)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El registro introducido ya existe.")
            End If
        End If
    End Sub

#End Region




End Class
