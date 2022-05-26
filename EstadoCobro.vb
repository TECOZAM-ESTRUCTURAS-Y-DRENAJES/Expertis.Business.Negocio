Public Class EstadoCobroInfo
    Inherits ClassEntityInfo

    Public IDEstado As Integer
    Public DescEstado As String
    Public Abreviatura As String
    Public IDAgrupacion As String
    Public Riesgo As Boolean
    Public RiesgoFactoring As Integer
    Public Desagrupable As Boolean
    Public GeneraRemesa As Boolean
    Public SimulacionTesoreria As Boolean
    Public Contabilidad As Boolean

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dt As DataTable = New EstadoCobro().SelOnPrimaryKey(PrimaryKey(0))
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            Me.Fill(dt.Rows(0))
        Else
            ApplicationService.GenerateError("El Estado Cobro {0} no existe.", Quoted(PrimaryKey(0)))
        End If
    End Sub

End Class

Public Class EstadoCobro

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroEstadoCobro"

#End Region

#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ComprobarSistema)
    End Sub

    <Task()> Public Shared Sub ComprobarSistema(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("Sistema") Then ApplicationService.GenerateError("No se puede realizar esa operacion sobre un Estado del Sistema.")
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDEstado")) = 0 Then ApplicationService.GenerateError("El Estado es un dato obligatorio.")
        If Length(data("DescEstado")) = 0 Then ApplicationService.GenerateError("La descripción del Estado es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dtObjeto As DataTable = New EstadoCobro().SelOnPrimaryKey(data("IDEstado"))
            If Not dtObjeto Is Nothing AndAlso dtObjeto.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Estado ya existe.")
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function EstadosCobrosAgrupables(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Return New EstadoCobro().Filter(New FilterItem("Desagrupable", FilterOperator.Equal, True))
    End Function

#End Region

End Class