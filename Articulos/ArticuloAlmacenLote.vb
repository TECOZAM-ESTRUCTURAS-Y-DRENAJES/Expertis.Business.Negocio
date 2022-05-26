Public Class _ArticuloAlmacenLote
    Public Const IDArticulo As String = "IDArticulo"
    Public Const IDAlmacen As String = "IDAlmacen"
    Public Const Lote As String = "Lote"
    Public Const Ubicacion As String = "Ubicacion"
    Public Const StockFisico As String = "StockFisico"
    Public Const FechaUltEntrada As String = "FechaUltEntrada"
    Public Const FechaCaducidad As String = "FechaCaducidad"
    Public Const FechaCreacionAudi As String = "FechaCreacionAudi"
    Public Const FechaModificacionAudi As String = "FechaModificacionAudi"
    Public Const UsuarioAudi As String = "UsuarioAudi"
    Public Const Bloqueado As String = "Bloqueado"
    Public Const Observaciones As String = "Observaciones"
    Public Const Traza As String = "Traza"

    Public Const SeriePrecinta As String = "SeriePrecinta"
    Public Const NDesdePrecinta As String = "NDesdePrecinta"
    Public Const NHastaPrecinta As String = "NHastaPrecinta"

    'Columna Auxiliar. En realidad no pertenece a la tabla. La actualizacion de stock utiliza esta columna.
    Public Const Cantidad As String = "Cantidad"
    Public Const Cantidad2 As String = "Cantidad2"
End Class

Public Class ArticuloAlmacenLote
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbArticuloAlmacenLote"

    Private Sub AddAuxiliarColumns(ByVal lote As DataTable)
        If Not lote Is Nothing Then
            Dim column As DataColumn = lote.Columns.Add(_ArticuloAlmacenLote.Cantidad, GetType(Double))
            column.DefaultValue = 0
            For Each dr As DataRow In lote.Rows
                dr(column.ColumnName) = 0
            Next

            Dim column2 As DataColumn = lote.Columns.Add(_ArticuloAlmacenLote.Cantidad2, GetType(Double))
            column2.DefaultValue = 0
            For Each dr As DataRow In lote.Rows
                dr(column2.ColumnName) = 0
            Next
        End If
    End Sub

    Public Overrides Function AddNew() As datatable
        Dim lote As DataTable = MyBase.AddNew()
        AddAuxiliarColumns(lote)
        Return lote
    End Function

    Public Overloads Overrides Function Filter(ByVal oFilter As Engine.IFilter, Optional ByVal strOrderBy As String = Nothing, Optional ByVal strSelect As String = Nothing) As DataTable
        Dim lote As DataTable = MyBase.Filter(oFilter)
        AddAuxiliarColumns(lote)
        Return lote
    End Function

    Public Overloads Overrides Function Filter(Optional ByVal strSelect As String = Nothing, Optional ByVal strWhere As String = Nothing, Optional ByVal strOrderBy As String = Nothing) As datatable
        Dim lote As DataTable = MyBase.Filter(strSelect, strWhere, strOrderBy)
        AddAuxiliarColumns(lote)
        Return lote
    End Function

    <Serializable()> _
    Public Class DataADDLoteAutomaticamente
        Public IDArticulo As String
        Public IDAlmacen As String
        Public Codigo As String
        Public Numeracion As Integer
        Public Lotes As DataTable
        Public CampoUnitario As String
        Public NumLotesCrear As Integer

        Public Sub New(ByVal IDArticulo As String, ByVal IDAlmacen As String, ByVal NumLotesCrear As Integer, ByVal Codigo As String, ByVal Numeracion As Integer, ByVal Lotes As DataTable, ByVal CampoUnitario As String)
            Me.IDArticulo = IDArticulo
            Me.IDAlmacen = IDAlmacen
            Me.Codigo = Codigo
            Me.Numeracion = Numeracion
            Me.Lotes = Lotes
            Me.CampoUnitario = CampoUnitario
            Me.NumLotesCrear = NumLotesCrear
        End Sub
    End Class

    <Task()> Public Shared Function ADDLoteAutomaticamente(ByVal data As DataADDLoteAutomaticamente, ByVal services As ServiceProvider) As DataTable
        Dim cnLONGITUD_MAX_LOTE As Integer = 25

        Dim dtNew As DataTable = data.Lotes.Clone

        Dim Ubicacion As String
        Dim f As New Filter
        f.Add("IDAlmacen", FilterOperator.Equal, data.IDAlmacen, FilterType.String)
        f.Add("Predeterminada", FilterOperator.Equal, True, FilterType.Boolean)
        Dim dtAU As DataTable = New AlmacenUbicacion().Filter(f, , "IDUbicacion")
        If Not dtAU Is Nothing AndAlso dtAU.Rows.Count > 0 Then
            Ubicacion = dtAU.Rows(0)("IDUbicacion")
        Else
            Ubicacion = New Parametro().UbicacionNoDefinida.IDUbicacion
        End If

        For i As Integer = 1 To data.NumLotesCrear
            If Length(data.Codigo & data.Numeracion) <= cnLONGITUD_MAX_LOTE Then
                Dim drNew As DataRow = dtNew.NewRow
                drNew("IDArticulo") = data.IDArticulo
                drNew("IDAlmacen") = data.IDAlmacen
                drNew("Lote") = data.Codigo & data.Numeracion
                drNew("Ubicacion") = Ubicacion
                drNew("Bloqueado") = False
                If Length(data.CampoUnitario) > 0 Then
                    drNew(data.CampoUnitario) = 1
                End If
                dtNew.Rows.Add(drNew)
                data.Numeracion = data.Numeracion + 1
            Else
                Exit For
            End If
        Next

        Return dtNew
    End Function

End Class