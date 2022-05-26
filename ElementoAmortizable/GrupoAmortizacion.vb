Public Class GrupoAmortizacion

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroGrupoAmortizacion"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("CCFondoAmortiz")) > 0 AndAlso Not ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ValidarCContable, data("CCFondoAmortiz"), services) Then
            ApplicationService.GenerateError("La cuenta contable | no existe", data("CCFondoAmortiz"))
        End If
        If Length(data("CCGastoAmortiz")) > 0 AndAlso Not ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ValidarCContable, data("CCGastoAmortiz"), services) Then
            ApplicationService.GenerateError("La cuenta contable | no existe", data("CCGastoAmortiz"))
        End If
        If Length(data("CCReservaRevaloriz")) > 0 AndAlso Not ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ValidarCContable, data("CCReservaRevaloriz"), services) Then
            ApplicationService.GenerateError("La cuenta contable | no existe", data("CCReservaRevaloriz"))
        End If
        If Length(data("CCFondoAmortizPlusvalia")) > 0 AndAlso Not ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ValidarCContable, data("CCFondoAmortizPlusvalia"), services) Then
            ApplicationService.GenerateError("La cuenta contable | no existe", data("CCFondoAmortizPlusvalia"))
        End If
        If Length(data("CCGastoAmortizPlusvalia")) > 0 AndAlso Not ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ValidarCContable, data("CCGastoAmortizPlusvalia"), services) Then
            ApplicationService.GenerateError("La cuenta contable | no existe", data("CCGastoAmortizPlusvalia"))
        End If
        If Length(data("IDGrupoAmortiz")) = 0 Then ApplicationService.GenerateError("Introduzca el identificativo del grupo de amortización.")
        If Length(data("IDTipoAmortiz")) = 0 Then ApplicationService.GenerateError("El Tipo de Amortización es obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim DtGrupos As DataTable = New GrupoAmortizacion().SelOnPrimaryKey(data("IDGrupoAmortiz"))
            If Not DtGrupos Is Nothing AndAlso DtGrupos.Rows.Count > 0 Then ApplicationService.GenerateError("La clave de Grupo de Amortización introducida ya existe en la base de datos.")
        End If
    End Sub

    <Task()> Public Shared Function ValidarCContable(ByVal data As String, ByVal services As ServiceProvider) As Boolean
        If data.Length > 0 Then
            Dim PlanCont As BusinessHelper = BusinessHelper.CreateBusinessObject("PlanContable")
            Dim dtPlanContable As DataTable = PlanCont.Filter(New FilterItem("IDCContable", FilterOperator.Equal, data))
            If dtPlanContable Is Nothing OrElse dtPlanContable.Rows.Count = 0 Then
                Return False
            Else : Return True
            End If
        Else : Return True
        End If
    End Function

#End Region

End Class