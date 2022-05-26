Public Class PedidoCompraPrecio

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbPedidoCompraPrecio"

#End Region
#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then data("IDLineaPedidoPrecio") = AdminData.GetAutoNumeric()
    End Sub

#End Region




#Region "Eventos GetBusinessRules"

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBrl As New BusinessRules
        oBrl.Add("IDArticulo", AddressOf CambioArticulo)
        oBrl.Add("IDLineaPedidoHija", AddressOf CambioLineaHija)
        oBrl.Add("IDLineaPedido", AddressOf CambioLineaPedido)
        oBrl.Add("Importe", AddressOf CambioImporte)
        oBrl.Add("Porcentaje", AddressOf CambioImporte)
        Return oBrl
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

    <Task()> Public Shared Sub CambioLineaHija(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim dtArtEspecial As DataTable
        If data.Context.ContainsKey("dtArtEspecial") Then
            dtArtEspecial = data.Context("dtArtEspecial")
            Dim dv As New DataView(dtArtEspecial)
            dv.RowFilter = "IDLineaPedido = " & data.Value
            If dv.Count > 0 Then
                data.Current("IDArticulo") = dv.Item(0).Row("IDArticulo")
                data.Current("DescArticulo") = dv.Item(0).Row("DescArticulo")
                data.Current("Precio") = dv.Item(0).Row("Precio")
            End If
        End If
    End Sub
    <Task()> Public Shared Sub CambioLineaPedido(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        Dim dtArtEspecial As DataTable
        If data.Context.ContainsKey("dtArtEspecial") Then
            dtArtEspecial = data.Context("dtArtEspecial")
            Dim dv As New DataView(dtArtEspecial)
            dv.RowFilter = "IDLineaPedido = " & data.Value
            If dv.Count > 0 Then
                data.Current("IDArticuloHijo") = dv.Item(0).Row("IDArticulo")
                data.Current("DescArticuloHijo") = dv.Item(0).Row("DescArticulo")
            End If
        End If
    End Sub
    <Task()> Public Shared Sub CambioImporte(ByVal data As BusinessRuleData, ByVal services As ServiceProvider)
        data.Current(data.ColumnName) = data.Value
        If Length(data.Value) > 0 Then
            If IsNumeric(data.Value) Then
                If data.ColumnName = "Porcentaje" Then
                    data.Current("Importe") = IIf(IsDBNull(data.Current("Precio")), 0, data.Current("Precio")) * IIf(data.Current("Porcentaje") <> 0, (data.Current("Porcentaje") / 100), 1)
                End If
                Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                Dim MonedaA As MonedaInfo = Monedas.MonedaA
                Dim MonedaB As MonedaInfo = Monedas.MonedaB
                data.Current("ImporteA") = xRound(data.Current("Importe") * data.Context("CambioA"), MonedaA.NDecimalesImporte)
                data.Current("ImporteB") = xRound(data.Current("Importe") * data.Context("CambioB"), MonedaB.NDecimalesImporte)
            Else
                ApplicationService.GenerateError("Campo no numérico.")
            End If
        End If
    End Sub
#End Region


   

End Class