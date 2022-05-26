<Transactional()> _
Public Class GRPAlbaranVentaCompraLinea
    Inherits ContextBoundObject

    Private Const cnEntidad As String = "tbGRPAlbaranVentaCompraLinea"

    Public Function AddNew() As DataTable
        Return AdminData.GetEntityData(Me.GetType.Name, , , True)
    End Function

    Public Function Filter(ByVal f As IFilter) As DataTable
        Dim dt As DataTable = New BE.DataEngine().Filter(cnEntidad, f)
        dt.TableName = "GRPAlbaranVentaCompraLinea"
        Return dt
    End Function

    Public Sub Delete(ByVal IDPVLinea As Integer)
        If IDPVLinea > 0 Then
            AdminData.DeleteData(Me.GetType.Name, IDPVLinea)
        End If
    End Sub

    Public Function TrazaAVPrincipal(ByVal IDAlbaran As Integer) As DataTable
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDAVPrincipal", IDAlbaran))
        Return Me.Filter(f)
    End Function

    Public Function TrazaAVLPrincipal(ByVal IDLineaAlbaran As Integer) As DataTable
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDLineaAVPrincipal", IDLineaAlbaran))
        Return Me.Filter(f)
    End Function

    Public Function TrazaACPrincipal(ByVal IDAlbaran As Integer) As DataTable
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDACPrincipal", IDAlbaran))
        Return Me.Filter(f)
    End Function

    Public Function TrazaACLPrincipal(ByVal IDLineaAlbaran As Integer) As DataTable
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDLineaACPrincipal", IDLineaAlbaran))
        Return Me.Filter(f)
    End Function

    Public Function TrazaAVSecundaria(ByVal IDAlbaran As Integer) As DataTable
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDAVSecundaria", IDAlbaran))
        Return Me.Filter(f)
    End Function

    Public Function TrazaAVLSecundaria(ByVal IDLineaAlbaran As Integer) As DataTable
        Dim f As New Filter
        f.Add(New NumberFilterItem("IDLineaAVSecundaria", IDLineaAlbaran))
        Return Me.Filter(f)
    End Function

End Class
