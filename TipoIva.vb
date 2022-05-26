Public Class TipoIvaInfo
    Inherits ClassEntityInfo

    Public IDTipoIVA As String
    Public DescTipoIVA As String
    Public Factor As Double
    Public IVARE As Double
    Public IVAIntrastat As Double
    Public SinRepercutir As Boolean
    Public IVASinRepercutir As Double
    Public NoDeclarar As Boolean
    Public CCSoportado As String
    Public CCRepercutido As String
    Public Importe As Double

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)


        Dim dt As DataTable = New TipoIva().SelOnPrimaryKey(PrimaryKey(0))
        Dim dtHistorico As DataTable
        If PrimaryKey.Length > 1 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDTipoIVA", PrimaryKey(0)))
            f.Add(New DateFilterItem("FechaDesde", FilterOperator.LessThanOrEqual, PrimaryKey(1)))
            f.Add(New DateFilterItem("FechaHasta", FilterOperator.GreaterThanOrEqual, PrimaryKey(1)))
            dtHistorico = AdminData.GetData("tbHistoricoTipoIva", f, "TOP 1  Factor, IvaRE, IvaIntrastat, IVASinRepercutir")
            If dtHistorico.Rows.Count > 0 Then
                dt.Rows(0)("Factor") = dtHistorico.Rows(0)("Factor")
                dt.Rows(0)("IvaRE") = dtHistorico.Rows(0)("IvaRE")
                dt.Rows(0)("IvaIntrastat") = dtHistorico.Rows(0)("IvaIntrastat")
                dt.Rows(0)("IVASinRepercutir") = dtHistorico.Rows(0)("IVASinRepercutir")
            End If
        End If

        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            Me.Fill(dt.Rows(0))
        End If
    End Sub

End Class

Public Class TipoIva

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroTipoIva"

#End Region

#Region "Eventos RegisterValidateTasks"

    ''' <summary>
    ''' Relación de tareas asociadas a la validación 
    ''' </summary>
    ''' <param name="validateProcess">Proceso en el que se registran las tareas de edición</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClavePrimaria)
    End Sub

    ''' <summary>
    ''' Comprobar que el código y la descripción no sean nulos
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipoIva")) = 0 Then ApplicationService.GenerateError("El Tipo IVA es obligatorio.")
        If Length(data("DescTipoIva")) = 0 Then ApplicationService.GenerateError("La Descripción es un dato obligatorio.")
    End Sub

    ''' <summary>
    ''' Comprobar que no exista la clave
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New TipoIva().SelOnPrimaryKey(data("IDTipoIva"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Código introducido ya existe.")
            End If
        End If
    End Sub

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("CCRepercutido", AddressOf TratarCContable)
        oBRL.Add("CCSoportado", AddressOf TratarCContable)
        Return oBRL
    End Function

    <Task()> Public Shared Sub TratarCContable(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current("CContable") = data.Value
        Dim DataT As New BusinessRuleData("CContable", data.Value, data.Current, data.Context)
        ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf NegocioGeneral.FormatoCuentaContable, DataT, services)
        data.Current(data.ColumnName) = data.Current("CContable")
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function GetFactorIva(ByVal strIDTipoIva As String, ByVal services As ServiceProvider) As Double
        Dim dblFactor As Double

        If Length(strIDTipoIva) > 0 Then
            Dim dtTI As DataTable = New TipoIva().SelOnPrimaryKey(strIDTipoIva)
            If Not dtTI Is Nothing AndAlso dtTI.Rows.Count > 0 Then
                dblFactor = dtTI.Rows(0)("Factor") / 100
            Else
                dblFactor = 0
            End If
        End If

        Return dblFactor
    End Function

    <Task()> Public Shared Function GetIvaFactorCero(ByVal obj As Object, ByVal services As ServiceProvider) As String
        Dim strIDTipoIva As String = String.Empty
        Dim dtTipoIva As DataTable = New TipoIva().Filter(New FilterItem("NoDeclarar", FilterOperator.NotEqual, 0))
        If Not dtTipoIva Is Nothing AndAlso dtTipoIva.Rows.Count > 0 Then
            strIDTipoIva = dtTipoIva.Rows(0)("IDTipoIva")
        End If
        Return strIDTipoIva
    End Function

    <Serializable()> _
    Public Class DataCalcularImporteIVA
        Public IDTipoIVA As String
        Public Importe As Double

        Public Sub New(ByVal IDTipoIVA As String, ByVal Importe As Double)
            Me.IDTipoIVA = IDTipoIVA
            Me.Importe = Importe
        End Sub
    End Class
    <Task()> Public Shared Function CalcularImporteIVA(ByVal data As DataCalcularImporteIVA, ByVal services As ServiceProvider) As Double
        Dim dblImporteIVA As Double
        If Not IsNothing(data) AndAlso Length(data.IDTipoIVA) > 0 Then
            Dim TiposIva As EntityInfoCache(Of TipoIvaInfo) = services.GetService(Of EntityInfoCache(Of TipoIvaInfo))()
            Dim TIvaInfo As TipoIvaInfo = TiposIva.GetEntity(data.IDTipoIVA)
            Dim dblFactor As Double = TIvaInfo.Factor
            If TIvaInfo.SinRepercutir Then
                '//Cogemos el factor del Iva sin repercutir.
                dblFactor = TIvaInfo.IVASinRepercutir
            End If
            dblImporteIVA = data.Importe * dblFactor / 100
        End If
        Return dblImporteIVA
    End Function

#End Region

End Class