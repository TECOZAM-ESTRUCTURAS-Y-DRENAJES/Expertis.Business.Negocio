Public Class AgrupacionInfo
    Inherits ClassEntityInfo

    Public IDAgrupacion As String
    Public DescAgrupacion As String
    Public DescGrupo As String

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dt As DataTable
        If Not IsNothing(PrimaryKey) AndAlso PrimaryKey.Length > 0 AndAlso Length(PrimaryKey(0)) > 0 Then
            dt = New Agrupacion().SelOnPrimaryKey(PrimaryKey(0))
        End If

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            ApplicationService.GenerateError("La agrupación | no existe.", Quoted(PrimaryKey(0)))
        Else
            Me.Fill(dt.Rows(0))
        End If
    End Sub

End Class


Public Class Agrupacion

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroAgrupacion"

#End Region

#Region "Eventos RegisterValidateTasks"

    ''' <summary>
    ''' Relación de tareas asociadas a la validación 
    ''' </summary>
    ''' <param name="validateProcess">Proceso en el que se registran las tareas de edición</param>
    ''' <remarks></remarks>
    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Solmicro.Expertis.Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarAgrupacionObligatoria)
    End Sub

    ''' <summary>
    ''' Comprobar que la agrupación tenga valor y que no exista
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarAgrupacionObligatoria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDAgrupacion")) = 0 Then ApplicationService.GenerateError("La agrupacion es un dato obligatorio.")
        If Length(data("DescAgrupacion")) = 0 Then ApplicationService.GenerateError("La descripcion de la agrupacion es un dato obligatorio.")
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New Agrupacion().SelOnPrimaryKey(data("IDAgrupacion"))
            If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("Código de Agrupación duplicado.")
            End If
        End If
    End Sub

#End Region

End Class