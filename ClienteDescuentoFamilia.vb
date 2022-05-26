Public Class ClienteDescuentoFamilia

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub
    Private Const cnEntidad As String = "tbClienteDescuentoFamilia"

#End Region

#Region "Eventos RegisterBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim Obrl As New BusinessRules
        Obrl.Add("IDTipo", AddressOf CambiarTipo)
        Obrl.Add("Dto1", AddressOf CambiarDescuentos)
        Obrl.Add("Dto2", AddressOf CambiarDescuentos)
        Obrl.Add("Dto3", AddressOf CambiarDescuentos)
        Return Obrl
    End Function

    <Task()> Public Shared Sub CambiarTipo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current("IDFamilia") = DBNull.Value
        data.Current("DescFamilia") = DBNull.Value
    End Sub

    <Task()> Public Shared Sub CambiarDescuentos(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Not IsNumeric(data.Value) Then
            ApplicationService.GenerateError("Campo no numérico.")
        End If
    End Sub

#End Region

#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarOfertaDetalle)
    End Sub

    <Task()> Public Shared Sub ActualizarOfertaDetalle(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDLineaOfertaDetalle")) > 0 Then
            Dim ClsObj As BusinessHelper = BusinessHelper.CreateBusinessObject("OfertaComercialDetalle")
            Dim dtO As DataTable = ClsObj.Filter("IDLineaOfertaDetalle,EstadoCliente", "IDLineaOfertaDetalle=" & data("IDLineaOfertaDetalle"))
            If Not dtO Is Nothing AndAlso dtO.Rows.Count > 0 Then
                dtO.Rows(0)("EstadoCliente") = False
                BusinessHelper.UpdateTable(dtO)
            End If
        End If

    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIDCliente)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIDTipo)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarFamilia)
    End Sub

    <Task()> Public Shared Sub ValidarIDCliente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDCliente")) = 0 Then ApplicationService.GenerateError("El Cliente es un dato obligatorio.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarIDTipo(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipo")) = 0 Then ApplicationService.GenerateError("El Tipo es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarFamilia(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            If AreDifferents(data("IDFamilia"), data("IDFamilia", DataRowVersion.Original)) OrElse AreDifferents(data("IDTipo"), data("IDTipo", DataRowVersion.Original)) Then
                Dim StrMessage As String
                Dim FilFam As New Filter
                FilFam.Add("IDCliente", FilterOperator.Equal, data("IDCliente"))
                FilFam.Add("IDTipo", FilterOperator.Equal, data("IDTipo"))
                If Length(data("IDFamilia")) > 0 Then
                    FilFam.Add("IDFamilia", FilterOperator.Equal, data("IDFamilia"))
                    StrMessage = "Ya existe un registro con ese Tipo y Familia."
                Else
                    FilFam.Add(New IsNullFilterItem("IDFamilia"))
                    StrMessage = "Ya existe un registro con ese Tipo y sin Familia."
                End If
                Dim dt As DataTable = New ClienteDescuentoFamilia().Filter(FilFam)
                If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then ApplicationService.GenerateError(StrMessage)
            End If
        Else
            If data.RowState = DataRowState.Added Then
                Dim StrMessage As String
                Dim FilFam As New Filter
                FilFam.Add("IDCliente", FilterOperator.Equal, data("IDCliente"))
                FilFam.Add("IDTipo", FilterOperator.Equal, data("IDTipo"))
                If Length(data("IDFamilia")) > 0 Then
                    FilFam.Add("IDFamilia", FilterOperator.Equal, data("IDFamilia"))
                    StrMessage = "Ya existe un registro con ese Tipo y Familia."
                Else
                    FilFam.Add(New IsNullFilterItem("IDFamilia"))
                    StrMessage = "Ya existe un registro con ese Tipo y sin Familia."
                End If
                Dim dt As DataTable = New ClienteDescuentoFamilia().Filter(FilFam)
                If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then ApplicationService.GenerateError(StrMessage)
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarPrimaryKey)
    End Sub

    <Task()> Public Shared Sub AsignarPrimaryKey(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If IsDBNull(data("IDClienteFamilia")) Then data("IDClienteFamilia") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

End Class