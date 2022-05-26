Public Class CondicionEnvioInfo
    Inherits ClassEntityInfo

    Public IDCondicionEnvio As String
    Public DescCondicionEnvio As String
    Public FactorValorEstadistico As Double
    Public DeclararIntrastat As Boolean
    Public FacturaPortes As Boolean
    Public FacturaDespacho As Boolean
    Public FacturaOtros As Boolean

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Sub New(ByVal IDCondicionEnvio As String)
        MyBase.New()
        Me.Fill(IDCondicionEnvio)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dtCondEnv As DataTable = New CondicionEnvio().Filter(New StringFilterItem("IDCondicionEnvio", PrimaryKey(0)))
        If dtCondEnv.Rows.Count > 0 Then
            Me.Fill(dtCondEnv.Rows(0))
        Else
            ApplicationService.GenerateError("La Condición de Envío | no existe.", Quoted(PrimaryKey(0)))
        End If
    End Sub

End Class

Public Class CondicionEnvio

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroCondicionEnvio"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        updateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescCondicionEnvio")) = 0 Then ApplicationService.GenerateError("Introduzca la descripción de la condición de envío")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDCondicionEnvio")) > 0 Then
                Dim DtTemp As DataTable = New CondicionEnvio().SelOnPrimaryKey(data("IDCondicionEnvio"))
                If Not DtTemp Is Nothing AndAlso DtTemp.Rows.Count > 0 Then
                    ApplicationService.GenerateError("Ya existe una condición de envío con esa clave.")
                End If
            Else : ApplicationService.GenerateError("Introduzca el código de la condición de envío.")
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function EstadoFacturas(ByVal data As String, ByVal services As ServiceProvider) As DataTable
        Return New CondicionEnvio().SelOnPrimaryKey(data)
    End Function

#End Region

End Class