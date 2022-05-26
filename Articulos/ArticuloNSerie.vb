Public Class ArticuloNSerieInfo
    Inherits ClassEntityInfo

    Public IDArticulo As String
    Public NSerie As String
    Public IDActivo As String
    Public IDEstadoActivo As String
    Public IDAlmacen As String
    Public IDOperario As String

    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub New(ByVal data As DataRow)
        MyBase.New(data)
    End Sub

    Public Sub New(ByVal IDArticulo As String, ByVal NSerie As String)
        MyBase.New()
        Me.Fill(IDArticulo, NSerie)
    End Sub

    Public Overloads Overrides Sub Fill(ByVal ParamArray PrimaryKey() As Object)
        Dim dtArtNSerie As DataTable = New ArticuloNSerie().SelOnPrimaryKey(PrimaryKey(0), PrimaryKey(1))
        If dtArtNSerie.Rows.Count > 0 Then
            Me.Fill(dtArtNSerie.Rows(0))
        End If
    End Sub

End Class


Public Class ArticuloNSerie
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Private Const COLOR_DEFAULTVALUE As Integer = 16777215

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbArticuloNSerie"
    Private Const cnEntidad2 As String = "vArticuloNSerie"

#Region " RegisterValidateTasks "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarRegistroExistente)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es obligatorio.")
        If Length(data("NSerie")) = 0 Then ApplicationService.GenerateError("El Nº de Serie es obligatorio.")
        If Length(data("IDAlmacen")) = 0 Then ApplicationService.GenerateError("El Almacén es un dato obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarRegistroExistente(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim aux As DataTable = New ArticuloNSerie().SelOnPrimaryKey(data("IDArticulo"), data("NSerie"))
            If aux.Rows.Count > 0 Then
                ApplicationService.GenerateError("El artículo {0} ya tiene asociado el Nº de serie {1}.", Quoted(data("IDArticulo")), Quoted(data("NSerie")))
            End If
        End If
    End Sub

#End Region


#Region " RegisterUpdateTasks "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificador)
    End Sub

    <Task()> Public Shared Sub AsignarIdentificador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added OrElse Nz(Length("MarcaAuto"), 0) = 0 Then
            data("MarcaAuto") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

#Region " AddNew "

    Public Overrides Function AddNew() As DataTable
        Dim serie As DataTable = MyBase.AddNew()
        AddAuxiliarColumns(serie)
        Return serie
    End Function

    Private Sub AddAuxiliarColumns(ByVal serie As DataTable)
        If Not serie Is Nothing Then
            Dim column As DataColumn
            column = serie.Columns.Add("Disponible", GetType(Boolean))
            column.DefaultValue = True
            column = serie.Columns.Add("EnCurso", GetType(Boolean))
            column.DefaultValue = False
            column = serie.Columns.Add("Baja", GetType(Boolean))
            column.DefaultValue = False
            column = serie.Columns.Add("Color", GetType(Integer))
            column.DefaultValue = COLOR_DEFAULTVALUE

            For Each dr As DataRow In serie.Rows
                dr("Disponible") = True
                dr("EnCurso") = False
                dr("Baja") = False
                If serie.Columns.Contains("CodProveedorAlquiler") Then dr("CodProveedorAlquiler") = Nothing
                If serie.Columns.Contains("MaquinaRealquilada") Then dr("MaquinaRealquilada") = False
                If serie.Columns.Contains("MaquinaRealquiladaGrupo") Then dr("MaquinaRealquiladaGrupo") = False
                dr("Color") = COLOR_DEFAULTVALUE
            Next
        End If
    End Sub

#End Region

#Region " Filter "

    Public Overloads Overrides Function Filter(ByVal oFilter As Engine.IFilter, Optional ByVal strOrderBy As String = Nothing, Optional ByVal strSelect As String = Nothing) As DataTable
        Dim serie As DataTable = New BE.DataEngine().Filter(cnEntidad2, oFilter, , strOrderBy)
        serie.TableName = Me.GetType.Name
        Return serie
    End Function

    Public Overloads Overrides Function Filter(Optional ByVal strSelect As String = Nothing, Optional ByVal strWhere As String = Nothing, Optional ByVal strOrderBy As String = Nothing) As DataTable
        Dim serie As DataTable = New BE.DataEngine().Filter(cnEntidad2, strSelect, strWhere, strOrderBy)
        serie.TableName = Me.GetType.Name
        Return serie
    End Function

#End Region

    Public Sub ActualizaArticuloNSerie(ByVal IDActivo As String, ByVal NSerie As String)
        Dim strSQL As String
        strSQL = "UPDATE tbArticuloNSerie"
        strSQL &= " SET IDActivo='" & IDActivo & "'"
        strSQL &= " WHERE Nserie='" & NSerie & "'"
        Try
            AdminData.Execute(strSQL)
        Catch ex As Exception

        End Try
    End Sub
    Public Sub ActualizaEstadoNSerie(ByVal NSerie As String, ByVal Estado As String)
        Dim strSQL As String
        strSQL = "UPDATE tbArticuloNSerie"
        strSQL &= " SET IDEstadoActivo='" & Estado & "'"
        strSQL &= " WHERE Nserie='" & NSerie & "'"
        Try
            AdminData.Execute(strSQL)
        Catch ex As Exception

        End Try
    End Sub


    Public Function ADDNSerieAutomaticamente(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal Cantidad As Double, _
                                             ByVal IDEstadoActivo As String, ByVal IDOperario As String, _
                                             ByVal codigo As String, ByVal numeracion As Integer, ByVal DtGrid As DataTable) As DataTable

        Dim Ubicacion As String = String.Empty
        If Length(IDAlmacen) > 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDAlmacen", IDAlmacen))
            f.Add(New BooleanFilterItem("Predeterminada", True))
            Dim dtUbicacion As DataTable = New AlmacenUbicacion().Filter(f)
            If dtUbicacion.Rows.Count > 0 Then Ubicacion = dtUbicacion.Rows(0)("IDUbicacion") & String.Empty
        End If

        Dim cnLONGITUD_MAX_NSERIE As Integer = 50
        Dim dtNew As DataTable = DtGrid.Clone

        For i As Integer = 1 To Cantidad
            If Length(codigo & numeracion) <= cnLONGITUD_MAX_NSERIE Then
                Dim drNew As DataRow = dtNew.NewRow

                drNew("IDArticulo") = IDArticulo
                drNew("IDAlmacen") = IDAlmacen
                drNew("NSerie") = codigo & numeracion
                drNew("IDEstadoActivo") = IDEstadoActivo
                drNew("IDOperario") = IDOperario
                drNew("Disponible") = True
                drNew("EnCurso") = False
                drNew("Baja") = False
                If drNew.Table.Columns.Contains("Ubicacion") Then drNew("Ubicacion") = Ubicacion
                dtNew.Rows.Add(drNew)

                numeracion = numeracion + 1
            Else
                Exit For
            End If
        Next

        Return dtNew
    End Function

End Class