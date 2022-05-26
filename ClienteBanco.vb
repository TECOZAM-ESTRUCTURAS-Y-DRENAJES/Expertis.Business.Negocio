Public Class ClienteBanco

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub
    Private Const cnEntidad As String = "tbClienteBanco"

#End Region

#Region "Eventos BusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Return MyBase.GetBusinessRules()
        Dim OBRL As New BusinessRules
        OBRL.Add("IDBanco", AddressOf CambioDiaPagoSucursal)
        OBRL.Add("Sucursal", AddressOf CambioDiaPagoSucursal)
        OBRL.Add("DigitoControl", AddressOf CambioDigitoControl)
        OBRL.Add("NCuenta", AddressOf CambioNCuenta)
    End Function

    <Task()> Public Shared Sub CambioDiaPagoSucursal(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) < 4 Then ApplicationService.GenerateError("La Entidad y la Sucursal han de ser de 4 dígitos.")
    End Sub

    <Task()> Public Shared Sub CambioDigitoControl(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) < 2 Then ApplicationService.GenerateError("EL Dígito de control ha de ser de 2 dígitos.")
    End Sub

    <Task()> Public Shared Sub CambioNCuenta(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) < 10 Then ApplicationService.GenerateError("La Cuenta ha de ser de | dígitos.")
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarIDBanco)
        validateProcess.AddTask(Of DataRow)(AddressOf General.Comunes.ValidarCodigoIBAN)
    End Sub

    <Task()> Public Shared Sub ValidarIDBanco(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDBanco")) = 0 Then ApplicationService.GenerateError("El Banco es un dato obligatorio.")
    End Sub

#End Region

#Region "Eventos RegisterUpdateTask"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIDClienteBanco)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarDC)
        updateProcess.AddTask(Of DataRow)(AddressOf TratarPredeterminado)
    End Sub

    <Task()> Public Shared Sub AsignarIDClienteBanco(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDCliente")) = 0 Then ApplicationService.GenerateError("El Cliente es un dato obligatorio.")
            If IsDBNull(data("IdClienteBanco")) Then
                data("IdClienteBanco") = AdminData.GetAutoNumeric
            End If
        End If
    End Sub

    <Task()> Public Shared Sub AsignarDC(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDBanco")) > 0 AndAlso Length(data("Sucursal")) > 0 AndAlso Length(data("NCuenta")) > 0 Then
            Dim dataDC As New NegocioGeneral.dataCalculoDigitosControl(data("IDBanco"), data("Sucursal"), data("NCuenta"))
            data("DigitoControl") = ProcessServer.ExecuteTask(Of NegocioGeneral.dataCalculoDigitosControl, String)(AddressOf NegocioGeneral.CalculoDigitosControl, dataDC, services)
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Sub TratarPredeterminado(ByVal dr As DataRow, ByVal services As ServiceProvider)
        Dim f As New Filter
        f.Add(New StringFilterItem("IDCliente", dr("IDCliente")))
        f.Add(New BooleanFilterItem("Predeterminado", True))
        Dim dtCB As DataTable = New ClienteBanco().Filter(f)

        If IsNothing(dtCB) OrElse dtCB.Rows.Count = 0 Then
            ' No hay más bancos de ese tipo dentro del Cliente actual con lo cual será el predeterminado.
            dr("Predeterminado") = True
        Else
            If IsDBNull(dr("Predeterminado")) Then dr("Predeterminado") = False
            ' Si el banco ha sido marcado como predeterminado
            If dr("Predeterminado") Then
                If dr("IdClienteBanco") <> dtCB.Rows(0)("IdClienteBanco") Then
                    dtCB.Rows(0)("Predeterminado") = False
                    BusinessHelper.UpdateTable(dtCB)
                End If
            ElseIf dr.RowState = DataRowState.Modified AndAlso dr("Predeterminado") <> dr("Predeterminado", DataRowVersion.Original) AndAlso dtCB.Rows.Count = 1 Then
                'dr("Predeterminado") = True
            End If
        End If
    End Sub

    <Task()> Public Shared Function GetBancoPredeterminado(ByVal strIDCliente As String, ByVal services As ServiceProvider) As Integer
        Dim intIDClienteBanco As Integer
        If Length(strIDCliente) > 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDCliente", strIDCliente))
            f.Add(New BooleanFilterItem("Predeterminado", True))
            Dim dtCB As DataTable = New ClienteBanco().Filter(f)
            If Not dtCB Is Nothing AndAlso dtCB.Rows.Count > 0 Then
                intIDClienteBanco = dtCB.Rows(0)("IdClienteBanco")
            End If
        End If
        Return intIDClienteBanco
    End Function

#End Region

End Class