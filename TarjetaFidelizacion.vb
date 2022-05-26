Public Class TarjetaFidelizacion

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroTarjetaFidelizacion"

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
    End Sub

    ''' <summary>
    ''' Comprobar que el c�digo y la descripci�n no sean nulos
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Informaci�n compartida</param>
    ''' <remarks></remarks>
    <Task()> Private Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTarjetafidelizacion")) = 0 Then ApplicationService.GenerateError("El N�mero de Tarjeta no es v�lido.")
    End Sub

    ''' <summary>
    ''' Comprobar que no exista la clave
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Informaci�n compartida</param>
    ''' <remarks></remarks>
    <Task()> Private Sub ValidarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = SelOnPrimaryKey(data("IDTarjetaFidelizacion"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El C�digo introducido ya existe.")
            End If
        End If
    End Sub

#End Region

#Region "Funciones Publicas"

    ''' <summary>
    ''' Asignaci�n de puntos, devuleve los puntos asignados y utilizados
    ''' </summary>
    ''' <param name="data">Tarjeta de fidelizaci�n</param>
    ''' <param name="services">Informaci�n compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Function Puntos(ByVal StrIDTarjetaFidel As String, ByVal services As ServiceProvider) As Hashtable
        Dim HT As New Hashtable
        Dim DblPuntosAsig As Double = 0
        Dim DblPuntosUtil As Double = 0
        Dim DtPuntos As DataTable = New BE.DataEngine().Filter("vFrmMntoTFGridPuntos", New FilterItem("IDTarjetaFidelizacion", FilterOperator.Equal, StrIDTarjetaFidel, FilterType.String))
        If Not DtPuntos Is Nothing AndAlso DtPuntos.Rows.Count > 0 Then
            For Each Dr As DataRow In DtPuntos.Select
                DblPuntosAsig += Dr("PuntosAsignados")
                DblPuntosUtil += Dr("PuntosUtilizados")
            Next
        End If
        HT.Add("PuntosAsignados", DblPuntosAsig)
        HT.Add("PuntosUtilizados", DblPuntosUtil)
        Return HT
    End Function

#End Region

End Class