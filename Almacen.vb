Public Class AlmacenInfo
    Inherits ClassEntityInfo

    Public IDAlmacen As String
    Public DescAlmacen As String
    Public IDCentroGestion As String
    Public Deposito As Boolean
    Public Principal As Boolean
    Public Empresa As Boolean
    Public Bloqueado As Boolean
    Public Activo As Boolean

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Sub New(ByVal IDAlmacen As String)
        MyBase.New()
        Me.Fill(IDAlmacen)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dtAlmInfo As DataTable = New Almacen().SelOnPrimaryKey(PrimaryKey(0))
        If dtAlmInfo.Rows.Count > 0 Then
            Me.Fill(dtAlmInfo.Rows(0))
        Else
            ApplicationService.GenerateError("El Almacen | no existe.", Quoted(PrimaryKey(0)))
        End If
    End Sub

End Class

Public Class Almacen

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroAlmacen"

#End Region

#Region " Delete "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.BeginTransaction)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.DeleteEntityRow)
        deleteProcess.AddTask(Of DataRow)(AddressOf Comunes.MarcarComoEliminado)
        deleteProcess.AddTask(Of DataRow)(AddressOf EliminarCorreccionesHistoricoMovimiento)
    End Sub

    <Task()> Public Shared Sub EliminarCorreccionesHistoricoMovimiento(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim strSql As String = "UPDATE tbHistoricoMovimiento SET IDTipoMovimiento=11 WHERE (IDAlmacen='" & data("IDAlmacen") & "')"
        AdminData.Execute(strSql)
    End Sub

#End Region

#Region " Validate "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarAlmacenObligatorio)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarAlmacenExistente)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarAlmacenObligatorio(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.ValidarAlmacenObligatorio, data, services)
    End Sub

    <Task()> Public Shared Sub ValidarAlmacenExistente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim dtA As DataTable = New Almacen().SelOnPrimaryKey(data("IDAlmacen"))
            If Not IsNothing(dtA) AndAlso dtA.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Almacén ya existe.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDCentroGestion")) = 0 Then ApplicationService.GenerateError("El Centro de Gestión es obligatorio.")
    End Sub

#End Region

#Region " Update "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.BeginTransaction)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.UpdateEntityRow)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DataRow)(AddressOf NuevaUbicacionAlmacen)
    End Sub

    <Task()> Public Shared Sub NuevaUbicacionAlmacen(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            ProcessServer.ExecuteTask(Of String)(AddressOf AlmacenUbicacion.NuevaUbicacion, data("IDAlmacen"), services)
        End If
    End Sub

#End Region

#Region " Alta Almacen "

    <Serializable()> _
    Public Class DataAltaAlmacen
        Public IDAlmacen, IDCentroGestion As String
        Public DescAlmacen As String
        Public Deposito As Boolean
        Public Empresa As Boolean
        Public Sub New(ByVal IDAlmacen As String, ByVal DescAlmacen As String, ByVal Deposito As Boolean, ByVal Empresa As Boolean, Optional ByVal IDCentroGestion As String = "")
            Me.IDAlmacen = IDAlmacen
            Me.DescAlmacen = DescAlmacen
            Me.Deposito = Deposito
            Me.Empresa = Empresa
            Me.IDCentroGestion = IDCentroGestion
        End Sub
    End Class
    <Task()> Public Shared Sub AltaAlmacen(ByVal data As DataAltaAlmacen, ByVal services As ServiceProvider)
        If Length(data.IDAlmacen) > 0 Then
            Dim a As New Almacen
            Dim dtA As DataTable = a.SelOnPrimaryKey(data.IDAlmacen)
            If Not dtA Is Nothing AndAlso dtA.Rows.Count = 0 Then
                Dim dt As DataTable = a.AddNewForm
                dt.Rows(0)("IDAlmacen") = data.IDAlmacen
                dt.Rows(0)("DescAlmacen") = data.DescAlmacen
                If Length(data.IDCentroGestion) > 0 Then dt.Rows(0)("IDCentroGestion") = data.IDCentroGestion
                dt.Rows(0)("Deposito") = data.Deposito
                dt.Rows(0)("Principal") = False
                dt.Rows(0)("Empresa") = data.Empresa
                a.Update(dt)
            End If
        End If
    End Sub

#End Region

#Region " Recuperar Almacen Alquiler "

    <Serializable()> _
    Public Class DataRecuperarAlmacenAlquiler
        Public IDArticulo As String
        Public IDCentroGestion As String
        Public Sub New(ByVal IDArticulo As String, ByVal IDCentroGestion As String)
            Me.IDArticulo = IDArticulo
            Me.IDCentroGestion = IDCentroGestion
        End Sub
    End Class

    <Task()> Public Shared Function RecuperaAlmacenAlquiler(ByVal data As DataRecuperarAlmacenAlquiler, ByVal services As ServiceProvider) As String
        Dim intTipoAlmAlq As Integer = New Parametro().RecuperarAlmacenAlq()

        Dim dt As DataTable = Nothing
        Select Case intTipoAlmAlq
            Case RecupAlmAlquiler.reAlmacenArticulo
                Dim f As New Filter
                f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
                f.Add(New BooleanFilterItem("Predeterminado", True))
                dt = New ArticuloAlmacen().Filter(f)
            Case RecupAlmAlquiler.reAlmacenCentroGestion
                dt = New CentroGestAlmSuminist().Filter(New StringFilterItem("IDCentroGestion", data.IDCentroGestion))
        End Select

        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            Return dt.Rows(0)("IDAlmacen") & String.Empty
        End If

        Return String.Empty
    End Function

#End Region

    <Task()> Public Shared Function GetAlmacenCentroGestion(ByVal IDCentroGestion As String, ByVal services As ServiceProvider) As String
        Dim AppParams As ParametroGeneral = services.GetService(Of ParametroGeneral)()
        If AppParams.AlmacenCentroGestionActivo Then
            Dim strAlmacen As String
            Dim f As New Filter
            f.Add(New StringFilterItem("IDCentroGestion", IDCentroGestion))
            f.Add(New BooleanFilterItem("Principal", True))

            Dim dtAlmacen As DataTable = New Almacen().Filter(f)
            If Not dtAlmacen Is Nothing AndAlso dtAlmacen.Rows.Count > 0 Then
                strAlmacen = dtAlmacen.Rows(0)("IDAlmacen") & String.Empty
            End If

            Return strAlmacen
        End If
    End Function

    <Task()> Public Shared Function ComprobarBloqueoAlmacen(ByVal IDAlmacen As String, ByVal services As ServiceProvider) As Boolean
        Dim Almacenes As EntityInfoCache(Of AlmacenInfo) = services.GetService(Of EntityInfoCache(Of AlmacenInfo))()
        Dim AlmInfo As AlmacenInfo = Almacenes.GetEntity(IDAlmacen)
        If Not AlmInfo Is Nothing AndAlso Length(AlmInfo.IDAlmacen) > 0 Then
            Return AlmInfo.Bloqueado
        End If
    End Function

    <Task()> Public Shared Function ActualizarGridCentroGestion3(ByVal dt As DataTable, ByVal services As ServiceProvider) As Boolean
        ' Se llama a esta función cuando se da de alta un nuevo registro, se modifica o se borra
        ' en el formulario de centros de gestión. En los 3 casos será finalmente, realmente, 
        ' una actualización sobre la tabla de almacenes
        Dim DtAlmacen As DataTable
        dt.TableName = "Almacen"
        Dim a As New Almacen
        ' Filas borradas
        If dt.Rows(0).RowState = DataRowState.Deleted Then
            dt.Rows(0).RejectChanges()
            ' Si no existe el almacén, "pasamos" de él
            DtAlmacen = a.SelOnPrimaryKey(dt.Rows(0)("IDAlmacen"))
            If Not DtAlmacen Is Nothing AndAlso DtAlmacen.Rows.Count > 0 Then
                DtAlmacen.Rows(0)("IDCentroGestion") = System.DBNull.Value
            Else : Return True
            End If
        Else
            ' Filas nuevas y modificadas
            Dim oAlmacenSuminist As New CentroGestAlmSuminist
            Dim idAlm, idCentroG As String

            ' Validación de datos
            ' IDAlmacen
            If dt.Rows(0)("IDAlmacen").ToString.Trim.Length = 0 Then
                ApplicationService.GenerateError("Introduzca el código del almacén")
            End If
            ' Descripción Almacén
            If dt.Rows(0)("DescAlmacen").ToString.Trim.Length = 0 Then
                ApplicationService.GenerateError("Introduzca la descripción del almacén")
            End If
            ' Almacén existe
            idAlm = dt.Rows(0)("IDAlmacen")
            DtAlmacen = a.SelOnPrimaryKey(idAlm)
            If DtAlmacen Is Nothing OrElse DtAlmacen.Rows.Count = 0 Then
                ApplicationService.GenerateError("El almacén introducido no existe")
            End If
            ' Miramos si es un almacén suministrador de este centro
            idCentroG = dt.Rows(0)("IDCentroGestion")
            Dim DtAux As DataTable = oAlmacenSuminist.SelOnPrimaryKey(idCentroG, idAlm)

            If Not DtAux Is Nothing AndAlso DtAux.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Almacén no puede ser gestionado por el Centro de Gestión porque es un Almacén Proveedor para este Centro de Gestión.")
            Else
                If dt.Rows(0).RowState = DataRowState.Added Then
                    ' Miramos si el almacén está asignado a otro centro de gestión. Sólo para el caso
                    ' de "nuevos" registros, puesto que los actualizados no pueden variar el IDAlmacen.
                    DtAux = a.Filter("IDAlmacen", "IDAlmacen='" & idAlm & "' " & _
                        "AND IDCentroGestion='" & idCentroG & "'")
                    If Not DtAux Is Nothing AndAlso DtAux.Rows.Count > 0 Then
                        ApplicationService.GenerateError("El almacén ya tiene asignado ese centro de gestión")
                    End If
                End If
                DtAlmacen.Rows(0)("IDCentroGestion") = idCentroG
                If dt.Rows(0)("Principal") Is System.DBNull.Value Then
                    DtAlmacen.Rows(0)("Principal") = False
                Else : DtAlmacen.Rows(0)("Principal") = dt.Rows(0)("Principal")
                End If
            End If
        End If
        BusinessHelper.UpdateTable(DtAlmacen)
        Return True
    End Function

End Class