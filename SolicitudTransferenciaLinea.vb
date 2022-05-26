Public Class SolicitudTransferenciaLinea

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Private Const cnEntidad As String = "tbSolicitudTransferenciaLinea"

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Public Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificador)
    End Sub

    <Task()> Public Shared Sub AsignarIdentificador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IdSolicitudLinea")) = 0 Then data("IdSolicitudLinea") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("Articulo", AddressOf ExisteArticuloEnAlmacenes)
        oBRL.Add("AlmacenOrigen", AddressOf ExisteArticuloEnAlmacenes)
        oBRL.Add("AlmacenDestino", AddressOf ExisteArticuloEnAlmacenes)
        Return oBRL
    End Function

    <Task()> Public Shared Sub ExisteArticuloEnAlmacenes(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumName) = data.Value
        If Length(data.Current("Articulo")) > 0 AndAlso Length(data.Current("AlmacenOrigen")) > 0 AndAlso Length(data.Current("AlmacenDestino")) > 0 Then
            'Comprobar si el articulo existe en el almacen origen o destino
            Dim dataBuscar As DataBuscarArticuloAlm
            dataBuscar.IDArticulo = data.Current("Articulo")
            dataBuscar.IDAlmacenOrigen = data.Current("AlmacenOrigen")
            dataBuscar.IDAlmacenDestino = data.Current("AlmacenDestino")
            If Not ProcessServer.ExecuteTask(Of DataBuscarArticuloAlm, Boolean)(AddressOf BuscarArchivoAlmacenes, dataBuscar, services) Then
                ApplicationService.GenerateError("El Articulo no existe en el Almacen Origen ni en el Almacen Destino.")
            End If
        End If
    End Sub

    Private Structure DataBuscarArticuloAlm
        Friend IDArticulo As String
        Friend IDAlmacenOrigen As String
        Friend IDAlmacenDestino As String
    End Structure

    <Task()> Private Shared Function BuscarArchivoAlmacenes(ByVal data As DataBuscarArticuloAlm, ByVal services As ServiceProvider) As Boolean
        'Dim sql As String = "SELECT DISTINCT IDArticulo, IDAlmacen FROM dbo.tbMaestroArticuloAlmacen " & _
        '                    "WHERE (IDArticulo = '" & Articulo & "') AND ((IDAlmacen = '" & AlmacenDest & "' OR IDAlmacen = '" & AlmacenOrig & "'))"
        'Dim DtResults As DataTable = AdminData.Execute(sql, ExecuteCommand.ExecuteReader)
        'Return IIf(Not DtResults Is Nothing AndAlso DtResults.Rows.Count > 0, True, False)

        Dim fAlm As New Filter(FilterUnionOperator.Or)
        fAlm.Add(New StringFilterItem("IDAlmacen", data.IDAlmacenDestino))
        fAlm.Add(New StringFilterItem("IDAlmacen", data.IDAlmacenOrigen))
        Dim f As New Filter
        f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
        Dim DtResults As DataTable = AdminData.GetData("tbMaestroArticuloAlmacen", f, "DICTINT IDArticulo, IDAlmacen")
        Return IIf(Not DtResults Is Nothing AndAlso DtResults.Rows.Count > 0, True, False)
    End Function

#End Region

End Class