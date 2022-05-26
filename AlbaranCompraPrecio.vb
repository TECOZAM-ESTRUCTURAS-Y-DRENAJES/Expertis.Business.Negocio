Public Class _AlbaranCompraPrecio
    Public Const IDLineaAlbaranPrecio As String = "IDLineaAlbaranPrecio"
    Public Const IDLineaAlbaran As String = "IDLineaAlbaran"
    Public Const IDLineaAlbaranHija As String = "IDLineaAlbaranHija"
    Public Const IDArticulo As String = "IDArticulo"
    Public Const DescArticulo As String = "DescArticulo"
    Public Const Porcentaje As String = "Porcentaje"
    Public Const Importe As String = "Importe"
    Public Const ImporteA As String = "ImporteA"
    Public Const ImporteB As String = "ImporteB"
    Public Const FechaCreacionAudi As String = "FechaCreacionAudi"
    Public Const FechaModificacionAudi As String = "FechaModificacionAudi"
    Public Const UsuarioAudi As String = "UsuarioAudi"
End Class

Public Class AlbaranCompraPrecio
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbAlbaranCompraPrecio"

#Region " Delete "

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.DeleteEntityRow)
        deleteProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoEliminado)
        deleteProcess.AddTask(Of DataRow)(AddressOf CorregirMovimientoGastos)
    End Sub

    <Task()> Public Shared Sub CorregirMovimientoGastos(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim ACL As New AlbaranCompraLinea
        Dim drACL As DataRow = ACL.GetItemRow(data("IDLineaAlbaran", DataRowVersion.Current))
        Dim Doc As New DocumentoAlbaranCompra(drACL("IDAlbaran"))
        Dim ctx As New DataDocRow(Doc, drACL)
        ProcessServer.ExecuteTask(Of DataDocRow)(AddressOf ProcesoAlbaranCompra.CorregirMovimiento, ctx, services)
    End Sub

#End Region

#Region " RegisterAddNewTasks"
    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of DataRow)(AddressOf AsignarIdentificador, data, services)
    End Sub

    <Task()> Public Shared Sub AsignarIdentificador(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDLineaAlbaranPrecio")) = 0 Then data("IDLineaAlbaranPrecio") = AdminData.GetAutoNumeric
    End Sub


#End Region

#Region " Update "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarIdentificador)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.UpdateEntityRow)
        updateProcess.AddTask(Of DataRow)(AddressOf Business.General.Comunes.MarcarComoActualizado)
        updateProcess.AddTask(Of DataRow)(AddressOf CorregirMovimientoGastos)
    End Sub

#End Region

#Region " Business Rules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("IDArticulo", AddressOf CambioArticulo)
        oBRL.Add("IDLineaAlbaran", AddressOf CambioLineaAlbaran)
        oBRL.Add("IDLineaAlbaranHija", AddressOf CambioLineaAlbaranHija)
        oBRL.Add("Importe", AddressOf CambioImporte)
        oBRL.Add("Porcentaje", AddressOf CambioPorcentaje)
        oBRL.Add("PorPorcentaje", AddressOf CambioPorPorcentaje)
        Return oBRL
    End Function

    <Task()> Public Shared Sub CambioArticulo(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If Length(data.Value) > 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", data.Value))
            f.Add(New BooleanFilterItem("Especial", True))
            Dim dt As DataTable = New BE.DataEngine().Filter("vfrmArticuloEspecial", f)
            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                data.Current("DescArticulo") = dt.Rows(0)("DescArticulo")
            Else
                ApplicationService.GenerateError("Articulo no existe o no Especial.")
            End If
        Else
            data.Current("DescArticulo") = System.DBNull.Value
        End If
    End Sub

    <Task()> Public Shared Sub CambioLineaAlbaran(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.Context.ContainsKey("dtArtEspecial") Then
            Dim dtArtEspecial As DataTable = data.Context("dtArtEspecial")
            Dim dv As New DataView(dtArtEspecial)
            dv.RowFilter = "IDLineaAlbaran = " & data.Value
            If dv.Count > 0 Then
                data.Current("IDArticuloHijo") = dv.Item(0).Row("IDArticulo")
                data.Current("DescArticuloHijo") = dv.Item(0).Row("DescArticulo")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioLineaAlbaranHija(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        If data.Context.ContainsKey("dtArtEspecial") Then
            Dim dtArtEspecial As DataTable = data.Context("dtArtEspecial")
            Dim dv As New DataView(dtArtEspecial)
            dv.RowFilter = "IDLineaAlbaran = " & data.Value
            If dv.Count > 0 Then
                data.Current("IDArticulo") = dv.Item(0).Row("IDArticulo")
                data.Current("DescArticulo") = dv.Item(0).Row("DescArticulo")
                data.Current("Precio") = dv.Item(0).Row("Precio")
                data.Current("ImporteAHija") = dv.Item(0).Row("ImporteA")
                data.Current("ImporteBHija") = dv.Item(0).Row("ImporteB")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioImporte(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("Importe")) > 0 Then
            If IsNumeric(data.Current("Importe")) Then
                If Nz(data.Current("PorPorcentaje"), False) AndAlso Nz(data.Current("Precio"), 0) <> 0 Then
                    data.Current("Porcentaje") = IIf(data.Current("Importe") <> 0, (data.Current("Importe") * 100), 1) / data.Current("Precio")
                End If

                Dim ValAyB As New ValoresAyB(data.Current, data.Context("IDMoneda"), data.Context("CambioA"), data.Context("CambioB"))
                ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
            Else
                ApplicationService.GenerateError("Campo no numérico.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioPorcentaje(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Current("Porcentaje")) > 0 Then
            If IsNumeric(data.Current("Porcentaje")) Then
                If Length(data.Current("IDLineaAlbaranHija")) = 0 Then
                    If Nz(data.Current("PorPorcentaje"), False) Then
                        data.Current("Importe") = Nz(data.Current("Precio")) * IIf(data.Current("Porcentaje") <> 0, (data.Current("Porcentaje") / 100), 1)
                    End If
                    Dim ValAyB As New ValoresAyB(data.Current, data.Context("IDMoneda"), data.Context("CambioA"), data.Context("CambioB"))
                    ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
                Else
                    data.Current(data.ColumnName) = data.Value
                    If data.ColumnName = "Porcentaje" Then
                        data.Current("Importe") = data.Current("Precio") * IIf(data.Current("Porcentaje") <> 0, (data.Current("Porcentaje") / 100), 1)
                        Dim ValAyB As New ValoresAyB(data.Current, data.Context("IDMoneda"), data.Context("CambioA"), data.Context("CambioB"))
                        ProcessServer.ExecuteTask(Of ValoresAyB, IPropertyAccessor)(AddressOf NegocioGeneral.MantenimientoValoresAyB, ValAyB, services)
                    End If
                    If Length(data.Current("ImporteAHija")) > 0 Then
                        data.Current("ImporteA") = Nz(data.Current("ImporteAHija"), 0) * IIf(data.Current("Porcentaje") <> 0, (data.Current("Porcentaje") / 100), 1)
                    End If
                    If Length(data.Current("ImporteBHija")) > 0 Then
                        data.Current("ImporteB") = Nz(data.Current("ImporteBHija"), 0) * IIf(data.Current("Porcentaje") <> 0, (data.Current("Porcentaje") / 100), 1)
                    End If
                End If
            Else : ApplicationService.GenerateError("Campo no numérico.")
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CambioPorPorcentaje(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Nz(data.Current("PorPorcentaje"), False) Then
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioPorcentaje, data, services)
        Else
            ProcessServer.ExecuteTask(Of BusinessRuleData)(AddressOf CambioImporte, data, services)
        End If
    End Sub

#End Region

End Class