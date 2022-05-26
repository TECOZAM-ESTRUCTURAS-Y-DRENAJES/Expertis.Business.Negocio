Public Class TipoAmortizacionLinea

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbTipoAmortizacionLinea"

#End Region

#Region "Eventos RegisterDeleteTasks"

    Protected Overrides Sub RegisterDeleteTasks(ByVal deleteProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterDeleteTasks(deleteProcess)
        deleteProcess.AddTask(Of DataRow)(AddressOf ActualizarVidaUtil)
    End Sub

    <Task()> Public Shared Sub ActualizarVidaUtil(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDTipoAmortizacion")) > 0 Then
            'HALLA EL CÁLCULO DE LA VIDA UTIL.
            'Fijarse en aquel porcentaje que supone que el porcentaje acumulado
            'sobrepase el 100%. Los porcentajes anteriores van sobre 12 meses
            Dim dblPorcAcum As Double
            Dim intVidaUtil As Short
            Dim dblPorcAmort As Double
            Dim dblMesesAmort As Double


            'Obtenemos todas las líneas del tipo de amortización actual ordenadas por año.
            Dim dtTALin As DataTable = New TipoAmortizacionLinea().Filter(New StringFilterItem("IdTipoAmortizacion", data("IDTipoAmortizacion")), "NAño")
            If Not dtTALin Is Nothing AndAlso dtTALin.Rows.Count > 0 Then
                For Each dr As DataRow In dtTALin.Rows
                    dblPorcAcum += dr("PorcentajeAmortizar")
                    If dblPorcAcum > 100 Then
                        'dblPorcAmort: Porcentaje que quedaba por amortizar antes de tener en cuenta
                        'este último registro
                        dblPorcAcum -= dr("PorcentajeAmortizar")
                        dblPorcAmort = 100 - dblPorcAcum

                        dblMesesAmort = (12 * dblPorcAmort) / dr("PorcentajeAmortizar")
                        'Transformamos los meses en un nº entero
                        If dblMesesAmort - Math.Floor(dblMesesAmort) > 0 Then
                            dblMesesAmort = Math.Floor(dblMesesAmort) + 1
                        End If
                        intVidaUtil += dblMesesAmort
                        Exit For
                    Else : intVidaUtil += 12
                    End If
                Next
            End If

            'Actualización del dato en TipoAmortizacionCabecera.
            Dim dtTACab As DataTable = New TipoAmortizacionCabecera().SelOnPrimaryKey(data("IDTipoAmortizacion"))
            If Not dtTACab Is Nothing AndAlso dtTACab.Rows.Count > 0 Then
                dtTACab.Rows(0)("VidaUtil") = intVidaUtil
                BusinessHelper.UpdateTable(dtTACab)
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarPorcentajeAmortizacion)
        validateProcess.AddTask(Of DataRow)(AddressOf ComprobarActualizaciones)
    End Sub

    <Task()> Public Shared Sub ComprobarPorcentajeAmortizacion(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Not IsDBNull(data("PorcentajeAmortizar")) AndAlso data("PorcentajeAmortizar") > 100 Then
            ApplicationService.GenerateError("El Porcentaje a Amortizar de un año no puede ser superior a 100")
        End If
    End Sub

    <Task()> Public Shared Sub ComprobarActualizaciones(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Modified Then
            If CInt(data("NAño", DataRowVersion.Original)) <> CInt(data("NAño")) OrElse _
                CDbl(data("PorcentajeAmortizar", DataRowVersion.Original)) <> CDbl(data("PorcentajeAmortizar")) Then
                ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarVidaUtil, data, services)
            End If
        ElseIf data.RowState = DataRowState.Added And data.IsNull("IDLineaTipoAmortizacion") Then
            data("IDLineaTipoAmortizacion") = AdminData.GetAutoNumeric
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ActualizarVidaUtil, data, services)
        End If
    End Sub

#End Region

End Class