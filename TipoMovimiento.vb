Public Class _TipoMovimiento
    Public Const IDTipoMovimiento As String = "IDTipoMovimiento"
    Public Const DescTipoMovimiento As String = "DescTipoMovimiento"
    Public Const ClaseMovimiento As String = "ClaseMovimiento"
    Public Const CodTipoMovimiento As String = "CodTipoMovimiento"
    Public Const Fifo As String = "Fifo"
    Public Const Consumo As String = "Consumo"
    Public Const Manual As String = "Manual"
    Public Const Sistema As String = "Sistema"
    Public Const FechaCreacionAudi As String = "FechaCreacionAudi"
    Public Const FechaModificacionAudi As String = "FechaModificacionAudi"
    Public Const UsuarioAudi As String = "UsuarioAudi"
End Class

Public Class TipoMovimientoInfo
    Inherits ClassEntityInfo

    Public IDTipoMovimiento As Integer
    Public DescTipoMovimiento As String
    Public ClaseMovimiento As enumtpmTipoMovimiento
    Public CodTipoMovimiento As String
    Public Fifo As Boolean
    Public Consumo As Boolean
    Public Manual As Boolean
    Public Sistema As Boolean


    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Sub New(ByVal IDTipoMovimiento As Integer)
        MyBase.New()
        Me.Fill(IDTipoMovimiento)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dtTMovtoInfo As DataTable = New TipoMovimiento().SelOnPrimaryKey(PrimaryKey(0))
        If dtTMovtoInfo.Rows.Count > 0 Then
            Me.Fill(dtTMovtoInfo.Rows(0))
        Else
            ApplicationService.GenerateError("El Tipo de Movimiento | no existe.", Quoted(PrimaryKey(0)))
        End If
    End Sub

End Class

Public Class TipoMovimiento

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroTipoMovimiento"


    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        deleteProcess.AddTask(Of DataRow)(AddressOf ProcesoComunes.ValidarDelRegistroSistema)
    End Sub

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
        If Length(data("IDTipoMovimiento")) = 0 Then ApplicationService.GenerateError("Tipo Movimiento es un dato obligatorio.")
        If Length(data("DescTipoMovimiento")) = 0 Then ApplicationService.GenerateError("La Descripción es un dato obligatorio.")
        If Length(data("ClaseMovimiento")) = 0 Then ApplicationService.GenerateError("La Clase Movimiento es obligatoria. es un dato obligatorio.")
        If Length(data("CodTipoMovimiento")) = 0 Then ApplicationService.GenerateError("El Cod. Tipo Movimiento es un dato obligatorio.")
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If AppParams.GestionInventarioPermanente Then
            If Nz(data("Contabilizar"), False) Then
                If data("IDTipoMovimiento") <> enumTipoMovimiento.tmEntAjuste AndAlso _
                   data("IDTipoMovimiento") <> enumTipoMovimiento.tmSalAjuste AndAlso _
                   data("IDTipoMovimiento") <> enumTipoMovimiento.tmInventario AndAlso _
                   Length(data("IDCContable")) = 0 Then
                    ApplicationService.GenerateError("La Cuenta Contable es un dato obligatorio.")
                End If

                If (data("IDTipoMovimiento") = enumTipoMovimiento.tmEntAjuste OrElse _
                 data("IDTipoMovimiento") = enumTipoMovimiento.tmSalAjuste OrElse _
                 data("IDTipoMovimiento") = enumTipoMovimiento.tmInventario) AndAlso _
                (Length(data("IDCContableI")) = 0 OrElse Length(data("IDCContableG")) = 0) Then
                    ApplicationService.GenerateError("Las Cuentas de Ingreso y Gasto son un dato obligatorio.")
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Comprobar que no exista la clave
    ''' </summary>
    ''' <param name="data">Registro modificado</param>
    ''' <param name="services">Información compartida</param>
    ''' <remarks></remarks>
    <Task()> Public Shared Sub ValidarClavePrimaria(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dt As DataTable = New TipoMovimiento().SelOnPrimaryKey(data("IDTipoMovimiento"))
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Código introducido ya existe.")
            End If
        End If
    End Sub

#End Region

#Region " Business Rules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("Contabilizar", AddressOf CambioContabilizar)
        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioContabilizar(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If AppParams.GestionInventarioPermanente Then
            If Not Nz(data.Current("Contabilizar")) Then
                data.Current("IDCContable") = System.DBNull.Value
                data.Current("IDCContableI") = System.DBNull.Value
                data.Current("IDCContableG") = System.DBNull.Value
            End If
        End If
    End Sub

#End Region

End Class