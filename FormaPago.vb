Public Class FormaPagoInfo
    Inherits ClassEntityInfo

    Public IDFormaPago As String
    Public DescFormaPago As String
    Public CobroRemesable As Boolean
    Public CobroImprimible As Boolean
    Public CobroALaVista As Boolean
    Public ChequeTalon As Boolean
    Public Trasferencia As Boolean
    Public ContabilidadEnVto As Boolean
    Public DiasMargenRiesgo As Integer
    Public Factoring As Integer
    Public CondicionVentaFactoring As String
    Public CodigoFacturae As Integer
    Public Tarjeta As Boolean
    Public Efectivo As Boolean
    Public Vale As Boolean
    Public IDMotivoNoAsegurado As String

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dt As DataTable = New FormaPago().SelOnPrimaryKey(PrimaryKey(0))
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            Me.Fill(dt.Rows(0))
        Else
            ApplicationService.GenerateError("La Forma de Pago {0} no existe.", Quoted(PrimaryKey(0)))
        End If
    End Sub

End Class

Public Class FormaPago

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroFormaPago"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("DescFormaPago")) = 0 Then ApplicationService.GenerateError("Introduzca la descripción de la forma de pago")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDFormaPago")) > 0 Then
                Dim dtTemp As DataTable = New FormaPago().SelOnPrimaryKey(data("IDFormaPago"))
                If Not dtTemp Is Nothing AndAlso dtTemp.Rows.Count > 0 Then
                    ApplicationService.GenerateError("Ya existe una forma de pago con la misma clave.")
                End If
            Else : ApplicationService.GenerateError("Introduzca el código de la forma de pago.")
            End If
        End If
    End Sub

#End Region

#Region " Tareas públicas "

    <Task()> Public Shared Function EsTarjeta(ByVal IDFormaPago As String, ByVal services As ServiceProvider) As Boolean
        If Length(IDFormaPago) > 0 Then
            Dim dtFP As DataTable = New FormaPago().Filter(New StringFilterItem("IDFormaPago", IDFormaPago), , "IDFormaPago, Tarjeta")
            If Not dtFP Is Nothing AndAlso dtFP.Rows.Count > 0 Then
                Return dtFP.Rows(0)("Tarjeta")
            End If
        End If
        Return False
    End Function

#End Region

End Class