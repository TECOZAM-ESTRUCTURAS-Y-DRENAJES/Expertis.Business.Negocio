Public Class CondicionPagoInfo
    Inherits ClassEntityInfo

    Public IDCondicionPago As String
    Public DescCondicionPago As String
    Public DtoProntoPago As Double
    Public RecFinan As Double
    Public IDMotivoNoAsegurado As String

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Sub New(ByVal IDCondicionPago As String)
        MyBase.New()
        Me.Fill(IDCondicionPago)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dtCondPago As DataTable = New CondicionPago().Filter(New StringFilterItem("IDCondicionPago", PrimaryKey(0)))
        If dtCondPago.Rows.Count > 0 Then
            Me.Fill(dtCondPago.Rows(0))
        Else
            ApplicationService.GenerateError("La Condición de Pago | no existe.", Quoted(PrimaryKey(0)))
        End If
    End Sub

End Class

Public Class CondicionPago

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroCondicionPago"

#End Region

#Region "Eventos RegisterValidarTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        ' Que haya introducido la descripción de la condición de pago.
        If Length(data("DescCondicionPago")) = 0 Then ApplicationService.GenerateError("Introduzca la descripción de la condición de pago")
        If Length(data("DtoProntoPago")) = 0 Then data("DtoProntoPago") = 0
        If Length(data("RecFinan")) = 0 Then data("RecFinan") = 0
        If data("DtoProntoPago") <> 0 And data("RecFinan") <> 0 Then ApplicationService.GenerateError("No se puede tener descuento pronto pago y recargo financiero a la vez")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDCondicionPago")) > 0 Then
                Dim DtTemp As DataTable = New CondicionPago().SelOnPrimaryKey(data("IDCondicionPago"))
                If Not DtTemp Is Nothing AndAlso DtTemp.Rows.Count > 0 Then
                    ApplicationService.GenerateError("Ya existe una condición de pago con esa clave")
                End If
            Else : ApplicationService.GenerateError("Introduzca el código de la condición de pago.")
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function ValidaCondicionPago(ByVal data As String, ByVal services As ServiceProvider) As DataTable
        Dim dt As DataTable = New CondicionPago().SelOnPrimaryKey(data)
        If IsNothing(dt) OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("La Condición de Pago | no existe.", data)
        End If
        Return dt
    End Function

#End Region

End Class