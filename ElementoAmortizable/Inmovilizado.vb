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
        If Length(data("IDInmovilizado")) = 0 Then ApplicationService.GenerateError("El identificador de Inmovilizado esta vac�o.")
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

#Region "Funciones P�blicas"

    <Serializable()> _
    Public Class DatosCambioCondiciones
        Public IDInmovilizado As String
        Public Fecha As Date
        Public IDEstado As String
    End Class

    <Task()> Public Shared Function CambiarCondiciones(ByVal data As DatosCambioCondiciones, ByVal services As ServiceProvider) As Boolean
        'Cambia los campos Estado y fecha inicio de los elementos amortizables 
        'de los inmovilizados obtenidos en el rs
        'Si alg�n elemento tiene valor amortizado, se cancela el proceso.
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
        DtAmort.Columns.Add("A�o", GetType(Integer))
        DtAmort.Columns.Add("AmortContable", GetType(Double))
        DtAmort.Columns.Add("ValorNeto", GetType(Double))
        DtAmort.Columns.Add("AmortContableMensual", GetType(String))
        Return DtAmort
    End Function

    <Serializable()> _
    Public Class DatosAmortContA�o
        Public IDInmovilizado As String
        Public A�o As Integer
        Public BlnA�o As Boolean
    End Class

    <Task()> Public Shared Function ObtenerAmortContA�o(ByVal data As DatosAmortContA�o, ByVal services As ServiceProvider) As Double
        Dim DtElementos As DataTable = New ElementoAmortizable().Filter(New FilterItem("IDInmovilizado", FilterOperator.Equal, data.IDInmovilizado))
        Dim DtAmort As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CrearDtAmort, Nothing, services)
        If Not DtElementos Is Nothing AndAlso DtElementos.Rows.Count > 0 Then
            Dim IntUltA�oAmort As Integer = 0
            For Each Dr As DataRow In DtElementos.Select
                Dim StAmortCont As New ElementoAmortizable.DataObtenerAmortCont(Dr("IDElemento"), DtAmort)
                DtAmort = ProcessServer.ExecuteTask(Of ElementoAmortizable.DataObtenerAmortCont, DataTable)(AddressOf ElementoAmortizable.ObtenerAmortCont, StAmortCont, services)
                If Length(Dr("FechaUltimaContabilizacion")) > 0 Then
                    If CDate(Dr("FechaUltimaContabilizacion")).Month = 12 AndAlso CDate(Dr("FechaUltimaContabilizacion")).Day = 31 Then
                        IntUltA�oAmort = CDate(Dr("FechaUltimaContabilizacion")).Year
                    Else
                        IntUltA�oAmort = CDate(Dr("FechaUltimaContabilizacioN")).Year - 1
                    End If
                End If
                Dim StAmort As New ElementoAmortizable.DataCalcAmort(Dr("IDElemento"), data.A�o)
                DtAmort = ProcessServer.ExecuteTask(Of ElementoAmortizable.DataCalcAmort, DataTable)(AddressOf ElementoAmortizable.CalcularAmortizacion, StAmort, services)
                Dim DrSel() As DataRow = DtAmort.Select("A�o = " & data.A�o)
                If DrSel.Length > 0 Then
                    ObtenerAmortContA�o += DrSel(0)("AmortA�o")
                End If
                DtAmort = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf CrearDtAmort, Nothing, services)
            Next
        End If
    End Function

#End Region

End Class