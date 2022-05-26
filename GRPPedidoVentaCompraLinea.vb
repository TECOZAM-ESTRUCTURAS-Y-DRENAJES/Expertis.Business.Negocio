Public Class GRPPedidoVentaCompraLinea
    Inherits BusinessHelper

    Private Const cnEntidad As String = "tbGRPPedidoVentaCompraLinea"

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Public Function TrazaPVPrincipal(ByVal IDPedido As Integer) As DataTable
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDPVPrincipal", IDPedido))
        Return Me.Filter(f)
    End Function

    Public Function TrazaPVLPrincipal(ByVal IDLineaPedido As Integer) As DataTable
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDLineaPVPrincipal", IDLineaPedido))
        Return Me.Filter(f)
    End Function

    Public Function TrazaPCPrincipal(ByVal IDPedido As Integer) As DataTable
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDPCPrincipal", IDPedido))
        Return Me.Filter(f)
    End Function

    Public Function TrazaPCLPrincipal(ByVal IDLineaPedido As Integer) As DataTable
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDLineaPCPrincipal", IDLineaPedido))
        Return Me.Filter(f)
    End Function

    Public Function TrazaPVSecundaria(ByVal IDPedido As Integer) As DataTable
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDPVSecundaria", IDPedido))
        Return Me.Filter(f)
    End Function

    Public Function TrazaPVLSecundaria(ByVal IDLineaPedido As Integer) As DataTable
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDLineaPVSecundaria", IDLineaPedido))
        Return Me.Filter(f)
    End Function

    Public Function TrazaAVPrincipal(ByVal IDAlbaran As Integer) As DataTable
        Dim dt As DataTable = New BE.DataEngine().Filter("vConsultaAlbaranesVentaMultiempresa", New NumberFilterItem("IDAlbaran", IDAlbaran))
        If dt.Rows.Count > 0 Then
            Return Me.TrazaPVPrincipal(dt.Rows(0)("IDPedido"))
        End If
        Return Nothing
    End Function

    Public Function TrazaAVLPrincipal(ByVal IDLineaAlbaran As Integer) As DataTable
        Dim dr As DataRow = New AlbaranVentaLinea().GetItemRow(IDLineaAlbaran)
        If IsNumeric(dr("IDLineaPedido")) Then
            Return Me.TrazaPVLPrincipal(dr("IDLineaPedido"))
        End If
        Return Nothing
    End Function

    Public Function TrazaAVSecundaria(ByVal IDAlbaran As Integer) As DataTable
        Dim dt As DataTable = New BE.DataEngine().Filter("vConsultaAlbaranesVentaMultiempresa", New NumberFilterItem("IDAlbaran", IDAlbaran))
        If dt.Rows.Count > 0 Then
            Return Me.TrazaPVSecundaria(dt.Rows(0)("IDPedido"))
        End If
        Return Nothing
    End Function

    Public Function TrazaAVLSecundaria(ByVal IDLineaAlbaran As Integer) As DataTable
        Dim dr As DataRow = New AlbaranVentaLinea().GetItemRow(IDLineaAlbaran)
        If IsNumeric(dr("IDLineaPedido")) Then
            Return Me.TrazaPVLSecundaria(dr("IDLineaPedido"))
        End If
        Return Nothing
    End Function

    Public Function TrazaACPrincipal(ByVal IDAlbaran As Integer) As DataTable
        Dim dt As DataTable = New BE.DataEngine().Filter("vConsultaAlbaranesCompraMultiempresa", New NumberFilterItem("IDAlbaran", IDAlbaran))
        If dt.Rows.Count > 0 Then
            Return Me.TrazaPCPrincipal(dt.Rows(0)("IDPedido"))
        End If
        Return Nothing
    End Function

    Public Function TrazaACLPrincipal(ByVal IDLineaAlbaran As Integer) As DataTable
        Dim dr As DataRow = New AlbaranCompraLinea().GetItemRow(IDLineaAlbaran)
        If IsNumeric(dr("IDLineaPedido")) Then
            Return Me.TrazaPCLPrincipal(dr("IDLineaPedido"))
        End If
        Return Nothing
    End Function

    Public Function TrazaFVSecundaria(ByVal IDFactura As Integer) As DataTable
        Dim dt As DataTable = New BE.DataEngine().Filter("vConsultaFacturasVentaMultiempresa", New NumberFilterItem("IDFactura", IDFactura))
        If dt.Rows.Count > 0 Then
            Return Me.TrazaAVSecundaria(dt.Rows(0)("IDAlbaran"))
        End If
        Return Nothing
    End Function

    Public Function TrazaFVLSecundaria(ByVal IDLineaFactura As Integer) As DataTable
        Dim dr As DataRow = New FacturaVentaLinea().GetItemRow(IDLineaFactura)
        If IsNumeric(dr("IDLineaAlbaran")) Then
            Return Me.TrazaAVLSecundaria(dr("IDLineaAlbaran"))
        End If
        Return Nothing
    End Function

    Public Function TrazaFCPrincipal(ByVal IDFactura As Integer) As DataTable
        Dim dt As DataTable = New BE.DataEngine().Filter("vConsultaFacturasCompraMultiempresa", New NumberFilterItem("IDFactura", IDFactura))
        If dt.Rows.Count > 0 Then
            Return Me.TrazaACPrincipal(dt.Rows(0)("IDAlbaran"))
        End If
        Return Nothing
    End Function

    Public Function TrazaFCLPrincipal(ByVal IDLineaFactura As Integer) As DataTable
        Dim dr As DataRow = New FacturaCompraLinea().GetItemRow(IDLineaFactura)
        If IsNumeric(dr("IDLineaAlbaran")) Then
            Return Me.TrazaACLPrincipal(dr("IDLineaAlbaran"))
        End If
        Return Nothing
    End Function
End Class

