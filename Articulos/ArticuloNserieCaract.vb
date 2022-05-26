Public Class ArticuloNserieCaract

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbArticuloNserieCaracteristica"

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IDCaracteristica", AddressOf CambioCaracteristica)
        Return oBrl
    End Function

    <Task()> Public Shared Sub CambioCaracteristica(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Current("IdAgrupacion")) > 0 Then
            Dim DtAgrup As DataTable = New CaracteristicaAgrupacion().SelOnPrimaryKey(data.Current("IdAgrupacion"))
            If Not DtAgrup Is Nothing AndAlso DtAgrup.Rows.Count > 0 Then
                data.Current("DescAgrupacion") = DtAgrup.Rows(0)("DescAgrupacion")
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarCaracteristica)
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCaracteristica")) > 0 Then
            Dim DtTemp As DataTable = New ArticuloNserieCaract().SelOnPrimaryKey(data("IDArticulo"), data("IDCaracteristica"))
            If Not DtTemp Is Nothing AndAlso DtTemp.Rows.Count > 0 Then
                ApplicationService.GenerateError("Este artículo ya tiene esta característica.")
            End If
        Else : ApplicationService.GenerateError("El código de la Característica es obligatorio.")
        End If
    End Sub

    <Task()> Public Shared Sub ValidarCaracteristica(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCaracteristica")) > 0 Then
            Dim DtExternos As DataTable = New Caracteristica().SelOnPrimaryKey(data("IDCaracteristica"))
            If DtExternos Is Nothing OrElse DtExternos.Rows.Count = 0 Then
                ApplicationService.GenerateError("El código de Característica no existe.")
            End If
        End If
    End Sub

#End Region

End Class