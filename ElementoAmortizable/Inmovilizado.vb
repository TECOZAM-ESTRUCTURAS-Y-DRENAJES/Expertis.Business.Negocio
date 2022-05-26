Public Class Inmovilizado

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroInmovilizado"

#End Region

#Region "Eventos RegisterValidateTask"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDInmovilizado")) = 0 Then ApplicationService.GenerateError("El identificador de Inmovilizado esta vacío.")
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf AsignarFechaContrato)
    End Sub

    <Task()> Public Shared Sub AsignarFechaContrato(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.HasVersion(DataRowVersion.Original) AndAlso data("FechaInmovilizado").ToString <> data("FechaInmovilizado", DataRowVersion.Original).ToString() Then
            Dim dtPP As DataTable = New PagoPeriodico().Filter(New FilterItem("IDInmovilizado", FilterOperator.Equal, data("IDInmovilizado")))
            If Not dtPP Is Nothing AndAlso dtPP.Rows.Count > 0 Then
                For Each drPP As DataRow In dtPP.Select
                    drPP("FechaContrato") = data("FechaInmovilizado")
                Next
                BusinessHelper.UpdateTable(dtPP)
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Serializable()> _
    Public Class DatosCambioCondiciones
        Public IDInmovilizado As String
        Public Fecha As Date
        Public IDEstado As String
    End Class

    <Task()> Public Shared Function CambiarCondiciones(ByVal data As DatosCambioCondiciones, ByVal services As ServiceProvider) As Boolean
        'Cambia los campos Estado y fecha inicio de los elementos amortizables 
        'de los inmovilizados obtenidos en el rs
        'Si algún elemento tiene valor amortizado, se cancela el proceso.
        Dim dtElementos As DataTable = New ElementoAmortizable().Filter(New FilterItem("IDInmovilizado", FilterOperator.Equal, data.IDInmovilizado))
        If Not dtElementos Is Nothing AndAlso dtElementos.Rows.Count > 0 Then
            'Comprueba el valor amortizado
            Dim dr() As DataRow = dtElementos.Select("ValorAmortizadoElementoA > 0")
            If dr.Length = 0 Then
                For Each drElemento As DataRow In dtElementos.Rows
                    drElemento("IdEstado") = data.IDEstado
                    drElemento("FechaInicioContabilizacion") = data.Fecha
                Next
                BusinessHelper.UpdateTable(dtElementos)
                Return True
            Else : Return False
            End If
        End If
        Return True
    End Function

    <Task()> Public Shared Function CrearDtAmort(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim DtAmort As New DataTable
        DtAmort.Columns.Add("Año", GetType(Integer))
        DtAmort.Columns.Add("AmortContable", GetType(Double))
        DtAmort.Columns.Add("ValorNeto", GetType(Double))
        DtAmort.Columns.Add("AmortContableMensual", GetType(String))
        Return DtAmort
    End Function

    <Serializable()> _
    Public Class DatosAmortContAño
        Public IDInmovilizado As String
        Public Año As Integer
        Public BlnAño As Boolean
    End Class

    <Task()> Public Shared Function ObtenerAmortContAño(ByVal data As DatosAmortContAño, ByVal services As ServiceProvider) As Double
        Dim DtElementos As DataTable = New ElementoAmortizable().Filter(New FilterItem("IDInmovilizado", FilterOperator.Equal, data.IDInmovilizado))
        Dim DtAmort As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CrearDtAmort, Nothing, services)
        If Not DtElementos Is Nothing AndAlso DtElementos.Rows.Count > 0 Then
            Dim IntUltAñoAmort As Integer = 0
            For Each Dr As DataRow In DtElementos.Select
                Dim StAmortCont As New ElementoAmortizable.DataObtenerAmortCont(Dr("IDElemento"), DtAmort)
                DtAmort = ProcessServer.ExecuteTask(Of ElementoAmortizable.DataObtenerAmortCont, DataTable)(AddressOf ElementoAmortizable.ObtenerAmortCont, StAmortCont, services)
                If Length(Dr("FechaUltimaContabilizacion")) > 0 Then
                    If CDate(Dr("FechaUltimaContabilizacion")).Month = 12 AndAlso CDate(Dr("FechaUltimaContabilizacion")).Day = 31 Then
                        IntUltAñoAmort = CDate(Dr("FechaUltimaContabilizacion")).Year
                    Else
                        IntUltAñoAmort = CDate(Dr("FechaUltimaContabilizacioN")).Year - 1
                    End If
                End If
                Dim StAmort As New ElementoAmortizable.DataCalcAmort(Dr("IDElemento"), data.Año)
                DtAmort = ProcessServer.ExecuteTask(Of ElementoAmortizable.DataCalcAmort, DataTable)(AddressOf ElementoAmortizable.CalcularAmortizacion, StAmort, services)
                Dim DrSel() As DataRow = DtAmort.Select("Año = " & data.Año)
                If DrSel.Length > 0 Then
                    ObtenerAmortContAño += DrSel(0)("AmortAño")
                End If
                DtAmort = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CrearDtAmort, Nothing, services)
            Next
        End If
    End Function

#End Region

End Class