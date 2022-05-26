Public Class Liquidacion

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbLiquidacion"

#End Region

#Region "Eventos RegisterAddNewTasks"

    Protected Overrides Sub RegisterAddnewTasks(ByVal addnewProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterAddnewTasks(addnewProcess)
        addnewProcess.AddTask(Of DataRow)(AddressOf FillDefaultValues)
    End Sub

    <Task()> Public Shared Sub FillDefaultValues(ByVal data As DataRow, ByVal services As ServiceProvider)
        data("IDLiquidacion") = AdminData.GetAutoNumeric
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDLiquidacion")) = 0 Then data("IDLiquidacion") = AdminData.GetAutoNumeric
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function LiquidarRepresentante(ByVal data As DataTable, ByVal services As ServiceProvider) As DataTable
        If Not data Is Nothing AndAlso data.Rows.Count > 0 Then
            data.DefaultView.Sort = "IDRepresentante"
            Dim strIDRepresentanteANT As String
            Dim ClsLiq As New Liquidacion
            Dim ClsFVR As New FacturaVentaRepresentante
            Dim ClsLiqDet As New LiquidacionDetalle
            Dim dtLiquidacion As DataTable = ClsLiq.AddNew
            Dim dtLiquidacionDetalle As DataTable = ClsLiqDet.AddNew
            Dim dtFactVRep As DataTable = ClsFVR.AddNew
            For Each dr As DataRow In data.Select("", "IdRepresentante")
                If strIDRepresentanteANT <> dr("IDRepresentante") Then
                    strIDRepresentanteANT = dr("IDRepresentante")
                    'Añadimos una nueva liquidación
                    Dim drLiquidacion As DataRow = dtLiquidacion.NewRow
                    drLiquidacion("IDLiquidacion") = AdminData.GetAutoNumeric
                    drLiquidacion("Fecha") = Date.Today
                    dtLiquidacion.Rows.Add(drLiquidacion)

                    'Añadimos las lineas de detalle de la liquidación
                    Dim DrFind() As DataRow = data.Select("IDRepresentante = '" & dr("IDRepresentante") & "'")
                    If DrFind.Length > 0 Then
                        For Each DrDet As DataRow In DrFind
                            Dim drLiquidacionDetalle As DataRow = dtLiquidacionDetalle.NewRow
                            drLiquidacionDetalle("IDLiquidacionDetalle") = AdminData.GetAutoNumeric
                            drLiquidacionDetalle("IDLiquidacion") = drLiquidacion("IDLiquidacion")
                            drLiquidacionDetalle("MarcaRepresentante") = DrDet("MarcaRepresentante")
                            drLiquidacionDetalle("ImpLiquidado") = DrDet("ImpALiquidar")
                            dtLiquidacionDetalle.Rows.Add(drLiquidacionDetalle)
                            DrDet("IDLiquidacion") = drLiquidacion("IDLiquidacion")

                            'Modificamos los siguientes campos en FacturaVentaRepresentante por linea de factura y representante
                            Dim filtro As New Filter
                            filtro.Add("IDLineaFactura", DrDet("IDLineaFactura"))
                            filtro.Add("IDRepresentante", DrDet("IDRepresentante"))
                            Dim dtFactVRepTemp As DataTable = ClsFVR.Filter(filtro)
                            dtFactVRepTemp.Rows(0)("IDLiquidacion") = drLiquidacion("IDLiquidacion")
                            dtFactVRepTemp.Rows(0)("ImporteLiquidado") += DrDet("ImpALiquidar")
                            dtFactVRep.ImportRow(dtFactVRepTemp.Rows(0))
                        Next
                    End If
                End If
            Next
            ClsLiq.Update(dtLiquidacion)
            ClsFVR.Update(dtFactVRep)
            ClsLiqDet.Update(dtLiquidacionDetalle)
            Return data
        End If
    End Function

    <Task()> Public Shared Sub AnularLiquidacion(ByVal data As DataTable, ByVal services As ServiceProvider)
        If Not data Is Nothing AndAlso data.Rows.Count > 0 Then
            Dim filtro As New Filter
            Dim ClsLD As New LiquidacionDetalle
            Dim ClsFVP As New FacturaVentaRepresentante
            Dim dt As New DataTable
            Dim dtFVR As DataTable
            For Each dr As DataRow In data.Rows
                filtro.Clear()
                filtro.Add("IDLiquidacion", dr("IDLiquidacion"))
                filtro.Add("MarcaRepresentante", dr("MarcaRepresentante"))
                dt = ClsLD.Filter(filtro)
                If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                    dr("ImporteLiquidado") -= dt.Rows(0)("ImpLiquidado")
                Else
                    dr("ImporteLiquidado") = 0
                End If
                ClsLD.Delete(dt)
                filtro.Clear()
                filtro.Add("IDLiquidacion", dr("IDLiquidacion"))
                filtro.Add("MarcaRepresentante", dr("MarcaRepresentante"))
                dtFVR = ClsFVP.Filter(filtro)
                If Not dtFVR Is Nothing AndAlso dtFVR.Rows.Count > 0 Then
                    dtFVR.Rows(0)("ImporteLiquidado") = dr("ImporteLiquidado")
                    If dtFVR.Rows(0)("ImporteLiquidado") = 0 Then
                        dtFVR.Rows(0)("IDLiquidacion") = System.DBNull.Value
                    Else
                        filtro.Clear()
                        filtro.Add("MarcaRepresentante", dr("MarcaRepresentante"))
                        dt = ClsLD.Filter(filtro, "IDLiquidacion desc")
                        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
                            dtFVR.Rows(0)("IDLiquidacion") = dt.Rows(0)("IDLiquidacion")
                        Else
                            dtFVR.Rows(0)("IDLiquidacion") = System.DBNull.Value
                        End If
                    End If
                    ClsFVP.Update(dtFVR)
                End If
            Next

        End If
    End Sub

#End Region

End Class