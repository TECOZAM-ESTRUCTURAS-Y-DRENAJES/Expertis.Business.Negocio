Public Class OperacionDetalle
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

#Region "Constructor"

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    ' ===============================================================
    ' Creado por: DnaGenerator 1.0 , SOLMICRO"
    ' Fecha : 13/06/2002 10:54:53
    '
    ' Descripción :
    ' Clase de objeto creada a partir de la clase de datos tbOperacionDetalle
    '
    ' ===============================================================

    Private Const cnEntidad As String = "tbOperacionDetalle"

#End Region

#Region "Eventos OperacionDetalle"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarClave)
        updateProcess.AddTask(Of DataRow)(AddressOf ActualizarSecuencias)
    End Sub

    <Task()> Public Shared Sub AsignarClave(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            If Length(data("IDParametro")) = 0 Then data("IDParametro") = AdminData.GetAutoNumeric
        End If
    End Sub

    <Task()> Public Shared Sub ActualizarSecuencias(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim blnModificado As Boolean
        Dim intSecuencia As Short
        Dim lngSkip As Integer

        Dim DtOp As DataTable = New OperacionDetalle().Filter(New FilterItem("IdOperacion", FilterOperator.Equal, data("IDOperacion")), "Secuencia")
        'Si no entra en el If mas interno, la funcion devuelve nothing por defecto
        If Not DtOp Is Nothing AndAlso DtOp.Rows.Count > 0 Then
            If data("Secuencia") <= DtOp.Rows(DtOp.Rows.Count - 1)("Secuencia") Then
                intSecuencia = data("Secuencia")
                lngSkip = 0
                For Each Dr As DataRow In DtOp.Select
                    Dim DrDatos() As DataRow = DtOp.Select("Secuencia=" & intSecuencia)
                    If DrDatos.Length > 0 Then
                        lngSkip = CInt(DrDatos.Length)
                        DrDatos(0)("Secuencia") += 10
                        intSecuencia = DrDatos(0)("Secuencia")
                        blnModificado = True
                    Else : Exit For
                    End If
                Next
                'rcsRuta.MoveFirst()
                'Si entra en el if ms interno, pero no hay modificaciones, la funcion
                'devuelve nothing por defecto
                If blnModificado Then BusinessHelper.UpdateTable(DtOp)

            End If
        End If
    End Sub

#End Region

End Class