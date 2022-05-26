Public Class TipoFacturaInfo
    Inherits ClassEntityInfo

    Public IDTipoFactura As Integer
    Public DescTipoFactura As String
    Public IDAgrupacion As String

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Sub New(ByVal IDTipoFactura As String)
        MyBase.New()
        Me.Fill(IDTipoFactura)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dt As DataTable = New TipoFactura().SelOnPrimaryKey(PrimaryKey(0))
        If dt.Rows.Count > 0 Then
            Me.Fill(dt.Rows(0))
        Else
            ApplicationService.GenerateError("El Tipo Factura {0} no existe.", Quoted(PrimaryKey(0)))
        End If
    End Sub

End Class



Public Class TipoFactura

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroTipoFactura"

#End Region

#Region "Eventos RegisterValidateTasks"

    ''' <summary>
    ''' Relaci�n de tareas asociadas a la validaci�n 
    ''' </summary>
    ''' <param name="validateProcess">Proceso en el que se registran las tareas de edici�n</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClavePrimaria)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarAgrupacion)
    End Sub

    ''' <summary>
    ''' Comprobar que el c�digo y la descripci�n no sean nulos
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Informaci�n compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipoFactura")) = 0 Then ApplicationService.GenerateError("Tipo factura es un dato obligatorio.")
        If Length(data("DescTipoFactura")) = 0 Then ApplicationService.GenerateError("La Descripci�n es un dato obligatorio.")
    End Sub

    ''' <summary>
    ''' Comprobar que no exista la clave
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Informaci�n compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New TipoFactura().SelOnPrimaryKey(data("IDTipoFactura"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El C�digo introducido ya existe.")
            End If
        End If
    End Sub

    ''' <summary>
    ''' Comprobar que el c�digo de agrupaci�n exista
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Informaci�n compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarAgrupacion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data("IDAgrupacion").ToString.Length > 0 Then
            Dim dt As DataTable = New Agrupacion().SelOnPrimaryKey(data("IDAgrupacion"))
            If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                ApplicationService.GenerateError("El c�digo de agrupaci�n no existe.")
            End If
        End If
    End Sub

#End Region

End Class