Public Class CobroPeriodico
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbCobroPeriodico"

#Region " Validate "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)

        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFechaInicioObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarFechaFinObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarUnidadObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarPeriodoObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarIDCContableObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarTipoCobroObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarMonedaObligatoria)
        validateProcess.AddTask(Of DataRow)(AddressOf NegocioGeneral.ValidarImporteObligatorio)
    End Sub

#End Region

#Region " Update "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificador)
    End Sub

    <Task()> Public Shared Sub AsignarIdentificador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("ID")) = 0 Then data("ID") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region


#Region " Business Rules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("CContable", AddressOf NegocioGeneral.FormatoCuentaContable)
        oBRL.Add("IdMoneda", AddressOf ProcesoComunes.CambioMoneda)
        oBRL.Add("IDCliente", AddressOf CambioCliente)
        oBRL.Add("IDFormaPago", AddressOf CambioFormaPago)
        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioCliente(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDCliente")) > 0 Then
            Dim Clientes As EntityInfoCache(Of ClienteInfo) = services.GetService(Of EntityInfoCache(Of ClienteInfo))()
            Dim ClteInfo As ClienteInfo = Clientes.GetEntity(data.Current("IDCliente"))
            data.Current("IDCContable") = ClteInfo.CCCliente
        End If
        data.Current("IDMandato") = System.DBNull.Value
        data.Current("NMandato") = System.DBNull.Value
    End Sub


    <Task()> Public Shared Sub CambioFormaPago(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("IDFormaPago")) > 0 Then
            Dim FormasPago As EntityInfoCache(Of FormaPagoInfo) = services.GetService(Of EntityInfoCache(Of FormaPagoInfo))()
            Dim FPInfo As FormaPagoInfo = FormasPago.GetEntity(data.Current("IDFormaPago"))
            If Not FPInfo.CobroRemesable Then
                data.Current("IDMandato") = System.DBNull.Value
                data.Current("NMandato") = System.DBNull.Value
            End If
        Else
            data.Current("IDMandato") = System.DBNull.Value
            data.Current("NMandato") = System.DBNull.Value
        End If
    End Sub

#End Region

End Class