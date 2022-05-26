Public Class RepartoCGCabecera

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroRepartoCentroGestionCabecera"

#End Region

#Region "Funciones Públicas"

    Public Function RepartirCGAnalitica(ByVal strIDRepartoCG As String, ByRef dtAnalitica As DataTable, ByVal dblImporte As Double, ByVal strIDMoneda As String, ByVal dtFecha As Date) As DataTable
        'Dim fwnRepartoCG As RepartoCGLinea
        'Dim rcsRepartoCG As Recordset
        'Dim FwnParametro As Parametro
        'Dim blnanaliticaCC As Boolean
        'Dim blnAnaliticaCG As Boolean
        'Dim strCentroCoste As String
        'Dim fwnMoneda As moneda
        'Dim moneda As MonedaInfo
        'Dim dblCambioA As Double
        'Dim dblCambioB As Double
        'Dim lngDecimales As Double
        'Dim lngDecimalesA As Double
        'Dim lngDecimalesB As Double
        'Dim dblTotalImpAnalitica As Double

        ''Comprueba si analítica por centro gestión o centro coste
        'FwnParametro = New Parametro
        'blnanaliticaCC = FwnParametro.CAnaliticPredet()
        'blnAnaliticaCG = (Not blnanaliticaCC) And FwnParametro.CAnaliticGestion
        ''Solo se hará reparto con analiticaCG si el usuario NO ha metido líneas
        ''Solo se hará reparto con analiticacc si el usuario SI ha metido líneas
        'If rcsAnalitica Is Nothing Then
        '    blnanaliticaCC = False
        'Else
        '    rcsAnalitica.Sort = "IDCentroCoste"
        '    blnanaliticaCC = blnanaliticaCC And rcsAnalitica.RecordCount > 0
        '    blnAnaliticaCG = blnAnaliticaCG And rcsAnalitica.RecordCount = 0
        'End If
        'If blnAnaliticaCG Then
        '    strCentroCoste = FwnParametro.CCostePredet
        'End If
        ''UPGRADE_NOTE: El objeto FwnParametro no se puede destruir hasta que no se realice la recolección de los elementos no utilizados. Haga clic aquí para obtener más información: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1029"'
        'FwnParametro = Nothing
        'If blnanaliticaCC Or blnAnaliticaCG Then
        '    'Obtiene los datos de la moneda
        '    fwnMoneda = New moneda
        '    moneda = fwnMoneda.ObtenerMoneda(strIDMoneda, dtFecha)
        '    lngDecimales = moneda.NDecimalesImporte
        '    dblCambioA = moneda.CambioA
        '    dblCambioB = moneda.CambioB

        '    moneda = fwnMoneda.MonedaA
        '    lngDecimalesA = moneda.NDecimalesImporte

        '    moneda = fwnMoneda.MonedaB
        '    lngDecimalesB = moneda.NDecimalesImporte

        '    'Forma la estructura del rs a devolver
        '    RepartirCGAnalitica = New Recordset
        '    RepartirCGAnalitica.Columns.Add("IDCentroCoste", GetType(String))
        '    RepartirCGAnalitica.Columns.Add("IDCentroGestion", GetType(String))
        '    RepartirCGAnalitica.Columns.Add("Porcentaje", GetType(Double))
        '    RepartirCGAnalitica.Columns.Add("Importe", GetType(Double))
        '    RepartirCGAnalitica.Columns.Add("ImporteA", GetType(Double))
        '    RepartirCGAnalitica.Columns.Add("ImporteB", GetType(Double))
        '    RepartirCGAnalitica.Open()

        '    'Por cada centro gestión a repartir insertamos una línea por cada centro coste
        '    fwnRepartoCG = New RepartoCGLinea
        '    rcsRepartoCG = fwnRepartoCG.Filter("IDCentroGestion,Porcentaje", "IDRepartoCentroGestion='" & strIDRepartoCG & "'")
        '    While Not rcsRepartoCG.EOF
        '        If blnanaliticaCC Then
        '            rcsAnalitica.MoveFirst()
        '            While Not rcsAnalitica.EOF
        '                RepartirCGAnalitica.AddNew()
        '                RepartirCGAnalitica.Fields("IDCentroCoste").Value = rcsAnalitica.Fields("IDCentroCoste").Value
        '                RepartirCGAnalitica.Fields("IDCentroGestion").Value = rcsRepartoCG.Fields("IDCentroGestion").Value
        '                RepartirCGAnalitica.Fields("Porcentaje").Value = xRound(rcsRepartoCG.Fields("Porcentaje").Value * rcsAnalitica.Fields("Porcentaje").Value / 100, 2)
        '                RepartirCGAnalitica.Fields("Importe").Value = xRound(rcsAnalitica.Fields("Importe").Value * rcsRepartoCG.Fields("Porcentaje").Value / 100, lngDecimales)
        '                RepartirCGAnalitica.Fields("ImporteA").Value = xRound(rcsAnalitica.Fields("ImporteA").Value * rcsRepartoCG.Fields("Porcentaje").Value / 100, lngDecimalesA)
        '                RepartirCGAnalitica.Fields("ImporteB").Value = xRound(rcsAnalitica.Fields("ImporteB").Value * rcsRepartoCG.Fields("Porcentaje").Value / 100, lngDecimalesB)
        '                RepartirCGAnalitica.Update()
        '                dblTotalImpAnalitica = dblTotalImpAnalitica + RepartirCGAnalitica.Fields("Importe").Value
        '                rcsAnalitica.MoveNext()
        '            End While
        '        ElseIf blnAnaliticaCG Then
        '            RepartirCGAnalitica.AddNew()
        '            RepartirCGAnalitica.Fields("IDCentroCoste").Value = strCentroCoste
        '            RepartirCGAnalitica.Fields("IDCentroGestion").Value = rcsRepartoCG.Fields("IDCentroGestion").Value
        '            RepartirCGAnalitica.Fields("Porcentaje").Value = rcsRepartoCG.Fields("Porcentaje").Value
        '            RepartirCGAnalitica.Fields("Importe").Value = xRound(dblImporte * rcsRepartoCG.Fields("Porcentaje").Value / 100, lngDecimales)
        '            RepartirCGAnalitica.Fields("ImporteA").Value = xRound(dblImporte * dblCambioA * rcsRepartoCG.Fields("Porcentaje").Value / 100, lngDecimalesA)
        '            RepartirCGAnalitica.Fields("ImporteB").Value = xRound(dblImporte * dblCambioB * rcsRepartoCG.Fields("Porcentaje").Value / 100, lngDecimalesB)
        '            RepartirCGAnalitica.Update()
        '            dblTotalImpAnalitica = dblTotalImpAnalitica + RepartirCGAnalitica.Fields("Importe").Value
        '        End If
        '        rcsRepartoCG.MoveNext()
        '    End While
        '    'Se modifica la última línea para evitar problemas de decimales
        '    If dblImporte <> dblTotalImpAnalitica Then
        '        RepartirCGAnalitica.Fields("Importe").Value = xRound(RepartirCGAnalitica.Fields("Importe").Value.Item.Value - dblTotalImpAnalitica + dblImporte, lngDecimales)
        '        RepartirCGAnalitica.Fields("ImporteA").Value = xRound(RepartirCGAnalitica.Fields("Importe").Value.Item.Value * dblCambioA, lngDecimalesA)
        '        RepartirCGAnalitica.Fields("ImporteB").Value = xRound(RepartirCGAnalitica.Fields("Importe").Value.Item.Value * dblCambioB, lngDecimalesB)
        '    End If
        'End If

        'Exit Function

    End Function

#End Region

End Class