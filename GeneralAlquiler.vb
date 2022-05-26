Public Class GeneralAlquiler

    <Serializable()> _
    Public Class DatosActualizarFechaAlquilerAlbaran
        Public IDLineaAlbaran() As Object
        Public FechaActualizacion As Date

        Public Sub New(ByVal IDLineaAlbaran() As Object, ByVal FechaActualizacion As Date)
            Me.IDLineaAlbaran = IDLineaAlbaran
            Me.FechaActualizacion = FechaActualizacion
        End Sub
    End Class
    <Task()> Public Shared Sub ActualizarFechasAlquiler(ByVal data As DatosActualizarFechaAlquilerAlbaran, ByVal services As ServiceProvider)
        Dim dt As DataTable = New AlbaranVentaLinea().Filter(New InListFilterItem("IDLineaAlbaran", data.IDLineaAlbaran, FilterType.Numeric))
        If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                dr("FechaAlquiler") = data.FechaActualizacion
            Next
            BusinessHelper.UpdateTable(dt)
        End If
    End Sub

    <Task()> Public Shared Sub RecalcularFechaAlquiler(ByVal IDLineaAlbaran() As Object, ByVal services As ServiceProvider)
        If IDLineaAlbaran.Length > 0 Then
            Dim p As New Parametro
            Dim dtmLimHora As Date = CDate(Format(CDate(Nz(p.LimiteHoraAlquiler, 0)), "hh:mm:ss"))

            Dim f As New Filter
            f.Add(New InListFilterItem("IDLineaAlbaran", IDLineaAlbaran, FilterType.Numeric))
            Dim dtAVL As DataTable = New AlbaranVentaLinea().Filter(f)
            If Not IsNothing(dtAVL) AndAlso dtAVL.Rows.Count > 0 Then
                Dim blnActualizar As Boolean = False
                For Each dr As DataRow In dtAVL.Rows
                    Dim FechaAlquiler As Date = Nz(dr("FechaAlquiler"), Date.MinValue)
                    Dim TipoFactAlquiler As Integer = Nz(dr("TipoFactAlquiler"), enumTipoFacturacionAlquiler.enumTFASinAlquiler)
                    Dim IDArticulo As String = dr("IDArticulo") & String.Empty
                    Dim IDObra As Integer = Nz(dr("IDObra"), 0)
                    Dim HoraAlquiler As Date = Nz(dr("FechaAlquiler"), Date.MinValue)

                    Dim datadiasMinimos As New GeneralAlquiler.dataFechaRetornoDiasMinimos(FechaAlquiler, TipoFactAlquiler, IDArticulo, IDObra, HoraAlquiler, dtmLimHora)
                    Dim dtFechaRetorno As Date = ProcessServer.ExecuteTask(Of GeneralAlquiler.dataFechaRetornoDiasMinimos, Date)(AddressOf GeneralAlquiler.ObtenerFechaRetornoDiasMinimos, datadiasMinimos, services)
                    If Length(dtFechaRetorno) > 0 Then
                        blnActualizar = True
                        dr("FechaRetornoDiasMinimos") = dtFechaRetorno
                    End If
                Next
                If blnActualizar Then BusinessHelper.UpdateTable(dtAVL)
            End If
        End If
    End Sub

    <Serializable()> _
    Public Class dataFechaRetornoDiasMinimos
        Public FechaAlquiler, LimHora, HoraAlquiler As Date
        Public TipoFactAlquiler, IDObra As Integer
        Public IDArticulo As String

        Public Sub New(ByVal FechaAlquiler As Date, ByVal TipoFactAlquiler As Integer, ByVal IDArticulo As String, _
                       ByVal IDObra As Integer, ByVal HoraAlquiler As Date, ByVal LimiteHora As Date)

            Me.FechaAlquiler = FechaAlquiler
            Me.TipoFactAlquiler = TipoFactAlquiler
            Me.IDArticulo = IDArticulo
            Me.IDObra = IDObra
            Me.HoraAlquiler = HoraAlquiler
            Me.LimHora = LimiteHora
        End Sub
    End Class
    <Task()> Public Shared Function ObtenerFechaRetornoDiasMinimos(ByVal data As dataFechaRetornoDiasMinimos, ByVal services As ServiceProvider) As Date
        Dim blnObra As Boolean = False
        Dim DiasAux As Integer = 0
        If data.LimHora >= Format(data.HoraAlquiler, "hh:mm:ss") Then
            DiasAux = -1
        End If

        Dim Obra As BusinessHelper = BusinessHelper.CreateBusinessObject("ObraCabecera")
        Dim dtObra As DataTable = Obra.SelOnPrimaryKey(data.IDObra)
        If Not dtObra Is Nothing AndAlso dtObra.Rows.Count > 0 Then
            blnObra = True
        End If

        Dim dtArt As DataTable = New Articulo().SelOnPrimaryKey(data.IDArticulo)
        If Not dtArt Is Nothing AndAlso dtArt.Rows.Count > 0 Then
            'A partir de aquí se empieza a calcular la fecha necesaria para el cálculo de los días mínimos.
            '       1. Si el TipoFactAlquiler es de Dias Naturales --> FechaAlquiler + NºDias
            '       2. Si el TipoFactAlquiler es de Dias Laborables --> FechaAlquiler + NºDias + Festivos
            '       3. Si el TipoFactAlquiler es cualquier otro --> la fecha de Dias minimos es la misma que la FechaAlquiler

            If data.TipoFactAlquiler = enumTipoFacturacionAlquiler.enumTFADiasNaturales Or data.TipoFactAlquiler = enumTipoFacturacionAlquiler.enumTFAMeses Then
                Return data.FechaAlquiler.AddDays(Nz(dtArt.Rows(0)("DiasMinimosFactAlquiler"), 0)).AddDays(DiasAux)
            ElseIf data.TipoFactAlquiler = enumTipoFacturacionAlquiler.enumTFADiasLaborables Or data.TipoFactAlquiler = enumTipoFacturacionAlquiler.enumTFAMesesDiasLaborables Then
                Dim Calendario As BusinessHelper
                Dim f As New Filter
                If blnObra Then
                    Calendario = BusinessHelper.CreateBusinessObject("CalendarioObra")
                    f.Add(New NumberFilterItem("IDObra", data.IDObra))
                Else
                    Calendario = BusinessHelper.CreateBusinessObject("CalendarioEmpresa")
                End If
                If Nz(dtArt.Rows(0)("DiasMinimosFactAlquiler"), 0) > 0 Then
                    Dim i As Integer = 0
                    Dim DiasAciertos As Integer = 0
                    Do While DiasAciertos <> dtArt.Rows(0)("DiasMinimosFactAlquiler")
                        i = i + 1
                        Dim dtmFechaAUX As Date = data.FechaAlquiler.AddDays(i)
                        Dim oF1 As New Filter
                        oF1.Add(New DateFilterItem("Fecha", dtmFechaAUX))
                        If f.Count > 0 Then oF1.Add(f)

                        Dim dtCal As DataTable = Calendario.Filter(oF1)
                        If Not dtCal Is Nothing AndAlso dtCal.Rows.Count > 0 Then
                            If dtCal.Rows(0)("TipoDia") = 0 Then
                                'Esto significa que este día si se tiene en cuenta, ya que es un día normal
                                DiasAciertos = DiasAciertos + 1
                            End If
                        Else
                            'Esto tb significa que es un día normal pq si no está en la tabla es que ese día no tiene ninguna cosa excepcional.
                            DiasAciertos = DiasAciertos + 1
                        End If
                    Loop
                    Return data.FechaAlquiler.AddDays(i + DiasAux)
                Else
                    Return data.FechaAlquiler
                End If
            ElseIf data.TipoFactAlquiler = enumTipoFacturacionAlquiler.enumTFADiasLaborables Then
                Return data.FechaAlquiler
            End If
        Else
            Return data.FechaAlquiler
        End If
    End Function

End Class
