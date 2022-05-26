Public Class ProveedorBanco

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbProveedorBanco"

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
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIDProveedorBanco)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarDC)
        updateProcess.AddTask(Of DataRow)(AddressOf TratarPredeterminado)
    End Sub

    <Task()> Public Shared Sub AsignarIDProveedorBanco(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDProveedor")) = 0 Then ApplicationService.GenerateError("El Proveedor es un dato obligatorio.")
            If IsDBNull(data("IdProveedorBanco")) Then
                data("IdProveedorBanco") = AdminData.GetAutoNumeric
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
        f.Add(New StringFilterItem("IDProveedor", dr("IDProveedor")))
        f.Add(New BooleanFilterItem("Predeterminado", True))
        Dim dtPB As DataTable = New ProveedorBanco().Filter(f)

        If IsNothing(dtPB) OrElse dtPB.Rows.Count = 0 Then
            ' No hay más bancos de ese tipo dentro del Cliente actual con lo cual será el predeterminado.
            dr("Predeterminado") = True
        Else
            If IsDBNull(dr("Predeterminado")) Then dr("Predeterminado") = False
            ' Si el banco ha sido marcado como predeterminado
            If dr("Predeterminado") Then
                If dr("IdProveedorBanco") <> dtPB.Rows(0)("IdProveedorBanco") Then
                    dtPB.Rows(0)("Predeterminado") = False
                    BusinessHelper.UpdateTable(dtPB)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Function GetBancoPredeterminado(ByVal strIDProveedor As String, ByVal services As ServiceProvider) As Integer
        Dim intIDProveedorBanco As Integer
        If Length(strIDProveedor) > 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDProveedor", strIDProveedor))
            f.Add(New BooleanFilterItem("Predeterminado", True))
            Dim dtPB As DataTable = New ProveedorBanco().Filter(f)
            If Not dtPB Is Nothing AndAlso dtPB.Rows.Count > 0 Then
                intIDProveedorBanco = dtPB.Rows(0)("IdProveedorBanco")
            End If
        End If
        Return intIDProveedorBanco
    End Function

#End Region

#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarPredeterminado)
    End Sub

    <Task()> Public Shared Sub ComprobarPredeterminado(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("Predeterminado") Then
            Dim dt As DataTable = New ProveedorBanco().Filter(New FilterItem("IDProveedor", FilterOperator.Equal, data("IDProveedor")))
            If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                dt.Rows(0)("Predeterminado") = True
                BusinessHelper.UpdateTable(dt)
            End If
        End If
    End Sub

#End Region

End Class