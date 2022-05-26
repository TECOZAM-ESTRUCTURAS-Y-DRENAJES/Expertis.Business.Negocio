Public Class ArticuloCosteEstandar

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbArticuloCosteEstandar"

#End Region

#Region "Definición Clases"

    <Serializable()> _
    Public Class udtCosteStdUnitario
        Public dblCosteStdA As Double
        Public dblCosteStdB As Double
    End Class

    <Serializable()> _
    Public Class udtCosteStdMat
        Public dblCosteMatA As Double
    End Class

    <Serializable()> _
    Public Class udtCosteStdOpe
        Public dblCosteIntA As Double
        Public dblCosteExtA As Double
        Public dblCosteFijo As Double
        Public dblCosteVariable As Double
        Public dblCosteDirecto As Double
        Public dblCosteIndirecto As Double
    End Class

    <Serializable()> _
    Public Class udtCosteStdVar
        Public dblCosteVarA As Double
        Public dblCosteFijo As Double
        Public dblCosteVariable As Double
        Public dblCosteDirecto As Double
        Public dblCosteIndirecto As Double
    End Class

    <Serializable()> _
    Public Class udtCosteStd
        Public udtMaterial As New udtCosteStdMat
        Public udtOperacion As New udtCosteStdOpe
        Public udtVarios As New udtCosteStdVar
    End Class

    <Serializable()> _
    Public Class udtCostesDT
        Public dtCosteStd As DataTable
        Public dtMaterial As DataTable
        Public dtOperacion As DataTable
        Public dtVarios As DataTable
    End Class

#End Region

    <Task()> Public Shared Sub AddArticuloCosteEstandar(ByVal data As DataRow, ByVal services As ServiceProvider)

        If Not IsNothing(data) AndAlso Length(data("IDArticulo")) > 0 Then
            Dim dt As DataTable = New ArticuloCosteEstandar().SelOnPrimaryKey(data("IDArticulo"))
            Dim acc As New ArticuloCosteEstandar()
            If IsNothing(dt) OrElse dt.Rows.Count = 0 Then
                dt = acc.AddNewForm
                dt.Rows(0)("IDArticulo") = data("IDArticulo")
                Dim DtArt As DataTable = AdminData.Filter("frmMntoArticulos", "IDRuta,IDTipoRuta,IDEstructura,IDTipoEstructura", "IDArticulo = '" & data("IDArticulo") & "'")
                If Not DtArt Is Nothing AndAlso DtArt.Rows.Count > 0 Then
                    dt.Rows(0)("IDRuta") = DtArt.Rows(0)("IDRuta")
                    dt.Rows(0)("IDTipoRuta") = DtArt.Rows(0)("IDTipoRuta")
                    dt.Rows(0)("IDEstructura") = DtArt.Rows(0)("IDEstructura")
                    dt.Rows(0)("IDTipoEstructura") = DtArt.Rows(0)("IDTipoEstructura")
                End If
            End If
            dt.Rows(0)("CosteStdA") = data("PrecioEstandarA")

            BusinessHelper.UpdateTable(dt)
        End If
    End Sub

#Region " Calculo CosteEstandar "

#Region " CosteEstandar "

    <Serializable()> _
    Public Class DataCosteEstandarDt
        Public DtEstandar As DataTable
        Public FechaCalculo As Date
        Public strError As String

        Public Sub New()
        End Sub
        Public Sub New(ByVal DtEstandar As DataTable, ByVal FechaCalculo As Date)
            Me.DtEstandar = DtEstandar
            Me.FechaCalculo = FechaCalculo
        End Sub
    End Class

    <Task()> Public Shared Function CosteEstandarDt(ByVal Data As DataCosteEstandarDt, ByVal services As ServiceProvider) As String
        If Not Data Is Nothing AndAlso Data.DtEstandar.Rows.Count > 0 Then
            Dim strError As String = String.Empty
            For Each dr As DataRow In Data.DtEstandar.Rows
                Try
                    Dim StData As New DataCosteEstandarIDArticulo(dr("IDArticulo"), Data.FechaCalculo)
                    ProcessServer.ExecuteTask(Of DataCosteEstandarIDArticulo, DataTable)(AddressOf CosteEstandarIDArticulo, StData, services)
                Catch ex As Exception
                    strError &= "Artículo " & dr("IDArticulo") & ": " & ex.Message & vbCrLf
                End Try
            Next
            Data.strError = strError
        End If
    End Function

    <Serializable()> _
    Public Class DataCosteEstandarIDArticulo
        Public IDArticulo As String
        Public FechaCalculo As Date

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDArticulo As String, ByVal FechaCalculo As Date)
            Me.IDArticulo = IDArticulo
            Me.FechaCalculo = FechaCalculo
        End Sub
    End Class

    <Task()> Public Shared Function CosteEstandarIDArticulo(ByVal data As DataCosteEstandarIDArticulo, ByVal services As ServiceProvider) As DataTable
        Dim udtDtCoste As New udtCostesDT
        If Length(data.IDArticulo) > 0 Then
            Dim dtACStd As DataTable = New ArticuloCosteEstandar().SelOnPrimaryKey(data.IDArticulo)
            If Not dtACStd Is Nothing AndAlso dtACStd.Rows.Count > 0 Then
                Dim dteUltimo As Date
                If Length(dtACStd.Rows(0)("FechaUltimo")) > 0 Then
                    dteUltimo = dtACStd.Rows(0)("FechaUltimo")
                End If
                udtDtCoste.dtCosteStd = dtACStd

                Dim dtComponentes As DataTable = AdminData.Execute("sp_CosteStdEstructuraExp", False, data.IDArticulo)
                If Not dtComponentes Is Nothing AndAlso dtComponentes.Rows.Count > 0 Then
                    Dim dtCopy As DataTable = dtComponentes.Copy

                    dtCopy.DefaultView.RowFilter = "IDComponente= '" & data.IDArticulo & "'"
                    If dtCopy.DefaultView.Count > 0 Then
                        If dtCopy.DefaultView(0).Row("Nivel") <> 0 Then
                            ApplicationService.GenerateError("Este componente no puede formar parte de su propia estructura")
                        Else
                            If dteUltimo <> Nz(dtACStd.Rows(0)("FechaEstandar")) Then
                                '///Lanzar un proceso de borrado sobre las tablas de historico de costes
                                AdminData.Execute("sp_CosteStdEliminarHistorico", False, data.IDArticulo, dteUltimo)
                            End If
                            Dim p As New Parametro
                            Dim intCriterioValoracion As Integer = p.CriterioValoracionCosteStd()
                            Dim udtCosteAcumulado As New udtCosteStd
                            Dim StDataComp As New DataCosteEstandarComponente(dtComponentes, data.IDArticulo, data.IDArticulo, "0", _
                                                                              String.Empty, 1, 1, 0, intCriterioValoracion, _
                                                                              udtCosteAcumulado, udtDtCoste, data.FechaCalculo)
                            ProcessServer.ExecuteTask(Of DataCosteEstandarComponente)(AddressOf CosteEstandarComponente, StDataComp, services)
                            ProcessServer.ExecuteTask(Of udtCostesDT)(AddressOf AplicarDecimales, udtDtCoste, services)

                            BusinessHelper.UpdateTable(StDataComp.udtDtNivel.dtCosteStd)
                            BusinessHelper.UpdateTable(StDataComp.udtDtNivel.dtMaterial)
                            BusinessHelper.UpdateTable(StDataComp.udtDtNivel.dtOperacion)
                            BusinessHelper.UpdateTable(StDataComp.udtDtNivel.dtVarios)
                        End If
                    End If
                Else : ApplicationService.GenerateError("Error en el procedimiento almacenado.")
                End If
            Else : ApplicationService.GenerateError("El artículo | no existe.", data.IDArticulo)
            End If
        Else : ApplicationService.GenerateError("El Artículo es un dato obligatorio.")
        End If
        Return udtDtCoste.dtCosteStd
    End Function

    <Serializable()> _
    Public Class DataCosteEstandarComponente
        Public dtComponentes As DataTable
        Public strArticuloPadre As String
        Public strArticulo As String
        Public strRamaArticulo As String
        Public strIDRutaPadre As String
        Public dblQComponente As Double
        Public dblQAcumulada As Double
        Public dblMerma As Double
        Public intCriterioValoracion As enumstdCriterioValoracion
        Public udtCosteAcumulado As udtCosteStd
        Public udtDtNivel As udtCostesDT
        Public FechaCalculo As Date

        Public Sub New()
        End Sub
        Public Sub New(ByVal dtComponentes As DataTable, ByVal strArticuloPadre As String, _
                       ByVal strArticulo As String, ByVal strRamaArticulo As String, _
                       ByVal strIDRutaPadre As String, ByVal dblQComponente As Double, _
                       ByVal dblQAcumulada As Double, ByVal dblMerma As Double, _
                       ByVal intCriterioValoracion As enumstdCriterioValoracion, _
                       ByVal udtCosteAcumulado As udtCosteStd, _
                       ByVal udtDtNivel As udtCostesDT, _
                       ByVal FechaCalculo As Date)
            Me.dtComponentes = dtComponentes
            Me.strArticuloPadre = strArticuloPadre
            Me.strArticulo = strArticulo
            Me.strRamaArticulo = strRamaArticulo
            Me.strIDRutaPadre = strIDRutaPadre
            Me.dblQComponente = dblQComponente
            Me.dblQAcumulada = dblQAcumulada
            Me.dblMerma = dblMerma
            Me.intCriterioValoracion = intCriterioValoracion
            Me.udtCosteAcumulado = udtCosteAcumulado
            Me.udtDtNivel = udtDtNivel
            Me.FechaCalculo = FechaCalculo
        End Sub
    End Class

    <Task()> Public Shared Function CosteEstandarComponente(ByVal data As DataCosteEstandarComponente, ByVal services As ServiceProvider) As udtCosteStdUnitario
        Dim StCoste As New udtCosteStdUnitario

        Dim udtCosteUnitario As New udtCosteStdUnitario
        Dim udtCosteOpe As New udtCosteStdOpe
        Dim udtCosteVar As New udtCosteStdVar
        Dim udtCosteNivel As New udtCosteStd
        Dim udtCosteAcumuladoNivel As New udtCosteStd

        Dim dblPrecioValoracionA As Double
        Dim dblStock As Double

        Dim strArticuloNivel As String = data.strArticulo
        Dim strIDEstructura As String
        Dim strIDTipoEstructura As String
        Dim dtComponentesCopy As DataTable = data.dtComponentes.Copy

        'Campo IDPadre (uno de los que devuelve el procedimiento almacenado) permite obtener la
        'estructura a primer nivel de cada uno de los componentes que forman la estructura
        Dim dvComponentes As DataView = dtComponentesCopy.DefaultView

        Dim Pos As Integer = Len(data.strRamaArticulo) - 1
        Dim f As New Filter
        f.Add(New StringFilterItem("IDPadre", FilterOperator.Equal, data.strArticulo))
        If Pos > 0 Then f.Add(New LikeFilterItem("Rama", data.strRamaArticulo & "%", True))
        Dim strWhere As String = f.Compose(New AdoFilterComposer)

        dvComponentes.RowFilter = strWhere
        '''aunque tenga estructura, si el artículo no es de tipo fábrica, no debe calcular nada más.
        Dim ArticuloI As New ArticuloInfo
        ArticuloI = ProcessServer.ExecuteTask(Of String, ArticuloInfo)(AddressOf Articulo.CaracteristicasArticuloInfo, data.strArticulo, services)
        If ArticuloI.Fabrica Or ArticuloI.Fantasma Then
            If dvComponentes.Count > 0 Then
                strIDEstructura = dvComponentes(0).Row("IDEstructura") & String.Empty
                strIDTipoEstructura = dvComponentes(0).Row("IDTipoEstructura") & String.Empty

                Dim dblQAcuNivel, dblQNeta As Double
                For Each drv As DataRowView In dvComponentes
                    data.strArticulo = drv("IDComponente")
                    data.strRamaArticulo = drv("Rama")
                    dblQNeta = drv("Cantidad") * (1 + (drv("Merma") / 100))
                    dblQAcuNivel = dblQNeta * data.dblQAcumulada

                    Dim StDataComp As New DataCosteEstandarComponente(data.dtComponentes, data.strArticuloPadre, data.strArticulo, data.strRamaArticulo, _
                                                     String.Empty, drv("Cantidad"), dblQAcuNivel, drv("Merma"), _
                                                     data.intCriterioValoracion, data.udtCosteAcumulado, data.udtDtNivel, data.FechaCalculo)
                    udtCosteUnitario = ProcessServer.ExecuteTask(Of DataCosteEstandarComponente, udtCosteStdUnitario)(AddressOf CosteEstandarComponente, StDataComp, services)
                    data.udtDtNivel = StDataComp.udtDtNivel

                    With udtCosteNivel.udtMaterial
                        .dblCosteMatA = .dblCosteMatA + udtCosteUnitario.dblCosteStdA
                    End With
                Next
            End If
        End If

        '///Control: El componente puede repetirse en disitntos puntos de la estructura
        Dim intOrden As Integer = 0
        Dim StrOrden(-1) As String
        If Not data.udtDtNivel.dtMaterial Is Nothing AndAlso data.udtDtNivel.dtMaterial.Rows.Count > 0 Then
            Dim dtMaterialCopy As DataTable = data.udtDtNivel.dtMaterial.Copy
            dtMaterialCopy.DefaultView.RowFilter = "IDComponente= '" & strArticuloNivel & "'"
            If dtMaterialCopy.DefaultView.Count > 0 Then
                ReDim StrOrden(dtMaterialCopy.DefaultView.Count - 1)
                Dim i As Integer = 0
                For Each Dr As DataRowView In dtMaterialCopy.DefaultView
                    StrOrden(i) = Dr("Orden")
                    i += 1
                Next
            End If
        End If
        f.Clear()

        f.Add(New StringFilterItem("IDComponente", FilterOperator.Equal, strArticuloNivel))
        If StrOrden.Length > 0 Then f.Add(New InListFilterItem("Orden", StrOrden, FilterType.String, False))
        strWhere = f.Compose(New AdoFilterComposer)

        dvComponentes.RowFilter = strWhere
        If dvComponentes.Count > 0 Then

            '///
            Dim strIDRuta, strIDTipoRuta As String

            Dim intNivel As Integer = dvComponentes(0).Row("Nivel")
            If intNivel <> 0 Then
                strIDEstructura = dvComponentes(0).Row("IDEstructura") & String.Empty
                strIDTipoEstructura = dvComponentes(0).Row("IDTipoEstructura") & String.Empty
            End If
            intOrden = dvComponentes(0).Row("Orden")

            If dvComponentes(0).Row("Fabrica") Then
                '///Coste del material
                With udtCosteAcumuladoNivel.udtMaterial
                    .dblCosteMatA = udtCosteNivel.udtMaterial.dblCosteMatA * data.dblQAcumulada
                End With
                '///Coste Operaciones
                Dim dblLote As Double = 1
                Dim aa As New ArticuloAlmacen
                Dim dtLote As DataTable = aa.Filter("LoteMinimo", "IDArticulo='" & strArticuloNivel & "' AND Predeterminado=1")
                If Not dtLote Is Nothing AndAlso dtLote.Rows.Count > 0 Then
                    dblLote = dtLote.Rows(0)("LoteMinimo")
                End If

                Dim StCosteOpe As New DataCosteStdOperaciones(data.strArticuloPadre, strArticuloNivel, dblLote, data.dblQAcumulada, _
                                                       intNivel, intOrden, data.udtDtNivel.dtOperacion, data.FechaCalculo, data.strIDRutaPadre)
                udtCosteOpe = ProcessServer.ExecuteTask(Of DataCosteStdOperaciones, udtCosteStdOpe)(AddressOf CosteEstandarOperaciones, StCosteOpe, services)

                data.udtDtNivel.dtOperacion = StCosteOpe.dtOperacion
                If Not data.udtDtNivel.dtOperacion Is Nothing AndAlso data.udtDtNivel.dtOperacion.Rows.Count > 0 Then
                    strIDRuta = data.udtDtNivel.dtOperacion.Rows(0)("IDRuta") & String.Empty
                    strIDTipoRuta = data.udtDtNivel.dtOperacion.Rows(0)("IDTipoRuta") & String.Empty
                End If

                With udtCosteNivel.udtOperacion
                    .dblCosteIntA = udtCosteOpe.dblCosteIntA
                    .dblCosteExtA = udtCosteOpe.dblCosteExtA
                    .dblCosteFijo = udtCosteOpe.dblCosteFijo
                    .dblCosteVariable = udtCosteOpe.dblCosteVariable
                    .dblCosteDirecto = udtCosteOpe.dblCosteDirecto
                    .dblCosteIndirecto = udtCosteOpe.dblCosteIndirecto
                End With

                '///Costes Varios
                Dim StCosteVar As New DataCosteStdVarios(data.strArticuloPadre, strArticuloNivel, data.dblQAcumulada, intNivel, _
                                                  intOrden, udtCosteNivel, data.udtDtNivel.dtVarios, data.FechaCalculo)
                udtCosteVar = ProcessServer.ExecuteTask(Of DataCosteStdVarios, udtCosteStdVar)(AddressOf CosteEstandarVarios, StCosteVar, services)

                data.udtDtNivel.dtVarios = StCosteVar.dtCoste
                With udtCosteNivel.udtVarios
                    .dblCosteVarA = udtCosteVar.dblCosteVarA
                    .dblCosteFijo = udtCosteVar.dblCosteFijo
                    .dblCosteVariable = udtCosteVar.dblCosteVariable
                    .dblCosteDirecto = udtCosteVar.dblCosteDirecto
                    .dblCosteIndirecto = udtCosteVar.dblCosteIndirecto
                End With

            ElseIf dvComponentes(0).Row("Compra") Then
                '///Coste del material
                If data.intCriterioValoracion = enumstdCriterioValoracion.stdPrecioEstandar Or (data.intCriterioValoracion = enumstdCriterioValoracion.stdCritValArticulo And dvComponentes(0).Row("CriterioValoracion") = enumtaValoracion.taPrecioEstandar) Then
                    dblPrecioValoracionA = dvComponentes(0).Row("PrecioStdA")
                ElseIf data.intCriterioValoracion = enumstdCriterioValoracion.stdCritValArticulo Then
                    If dvComponentes(0).Row("CriterioValoracion") = enumtaValoracion.taPrecioFIFOFecha Or dvComponentes(0).Row("CriterioValoracion") = enumtaValoracion.taPrecioFIFOMvto Or dvComponentes(0).Row("CriterioValoracion") = enumtaValoracion.taPrecioMedio Then
                        Dim strAlmacen As String = String.Empty
                        Dim FwnArtAlm As New ArticuloAlmacen
                        Dim dtAlmacen As DataTable = FwnArtAlm.Filter("IDAlmacen,StockFisico", "IDArticulo='" & strArticuloNivel & "' AND Predeterminado=1")
                        If Not dtAlmacen Is Nothing AndAlso dtAlmacen.Rows.Count > 0 Then
                            strAlmacen = dtAlmacen.Rows(0)("IDAlmacen") & String.Empty
                            dblStock = dtAlmacen.Rows(0)("StockFisico")
                        End If
                        If Len(strAlmacen) > 0 Then
                            Dim Precio As ValoracionPreciosInfo
                            Select Case CType(dvComponentes(0).Row("CriterioValoracion"), enumtaValoracion)
                                Case enumtaValoracion.taPrecioFIFOFecha
                                    Dim datosPrecio As New ProcesoStocks.DataValoracionFIFO(strArticuloNivel, strAlmacen, dblStock, dblStock, Today, enumstkValoracionFIFO.stkVFOrdenarPorFecha)
                                    Precio = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ProcesoStocks.ValoracionFIFO, datosPrecio, services)
                                Case enumtaValoracion.taPrecioFIFOMvto
                                    Dim datosPrecio As New ProcesoStocks.DataValoracionFIFO(strArticuloNivel, strAlmacen, dblStock, dblStock, Today, enumstkValoracionFIFO.stkVFOrdenarPorMvto)
                                    Precio = ProcessServer.ExecuteTask(Of ProcesoStocks.DataValoracionFIFO, ValoracionPreciosInfo)(AddressOf ProcesoStocks.ValoracionFIFO, datosPrecio, services)
                                Case enumtaValoracion.taPrecioMedio
                                    Dim datosValPMF As New DataArticuloAlmacenFecha(strArticuloNivel, strAlmacen, Today)
                                    Precio = ProcessServer.ExecuteTask(Of DataArticuloAlmacenFecha, ValoracionPreciosInfo)(AddressOf ProcesoStocks.ValoracionPrecioMedioAFecha, datosValPMF, services)
                            End Select
                            If Not Precio Is Nothing Then
                                dblPrecioValoracionA = Precio.PrecioA
                            End If
                        End If
                    ElseIf dvComponentes(0).Row("CriterioValoracion") = enumtaValoracion.taPrecioUltCompra Then
                        dblPrecioValoracionA = dvComponentes(0).Row("PrecioUltimaCompraA")
                    End If
                End If

                udtCosteUnitario.dblCosteStdA = dblPrecioValoracionA
                udtCosteNivel.udtMaterial.dblCosteMatA = dblPrecioValoracionA
                udtCosteAcumuladoNivel.udtMaterial.dblCosteMatA = dblPrecioValoracionA * data.dblQAcumulada
            End If

            '///Coste Ud. (nivel superior)
            With udtCosteNivel
                udtCosteUnitario.dblCosteStdA = data.dblQComponente * (1 + (data.dblMerma / 100)) * (.udtMaterial.dblCosteMatA + .udtOperacion.dblCosteIntA + .udtOperacion.dblCosteExtA + .udtVarios.dblCosteVarA)
            End With

            '///Valores acumulados (entre niveles)
            If dvComponentes(0).Row("Compra") And Not dvComponentes(0).Row("Fabrica") Then
                With data.udtCosteAcumulado.udtMaterial
                    .dblCosteMatA = .dblCosteMatA + udtCosteAcumuladoNivel.udtMaterial.dblCosteMatA
                End With
            End If
            With data.udtCosteAcumulado.udtOperacion
                .dblCosteIntA = .dblCosteIntA + (udtCosteOpe.dblCosteIntA * data.dblQAcumulada)
                .dblCosteExtA = .dblCosteExtA + (udtCosteOpe.dblCosteExtA * data.dblQAcumulada)
            End With
            With data.udtCosteAcumulado.udtVarios
                .dblCosteVarA = .dblCosteVarA + (udtCosteVar.dblCosteVarA * data.dblQAcumulada)
            End With

            Dim drNuevaLineaMaterial As DataRow
            'Dim drNuevaLineaCosteStd As DataRow
            Dim drComponente As DataRow = dvComponentes(0).Row

            If intNivel > 0 Then
                Dim StCosteMat As New DataLlenarCosteMat(data.strArticuloPadre, drComponente, data.dblQAcumulada, _
                                                           intOrden, strIDRuta, strIDTipoRuta, udtCosteUnitario, _
                                                           udtCosteNivel, udtCosteAcumuladoNivel, data.FechaCalculo)
                drNuevaLineaMaterial = ProcessServer.ExecuteTask(Of DataLlenarCosteMat, DataRow)(AddressOf LlenarCosteMaterial, StCosteMat, services)
            ElseIf intNivel = 0 Then
                Dim StCosteMat As New DataLlenarCosteMat(data.strArticuloPadre, drComponente, data.dblQAcumulada, _
                                                           intOrden, strIDRuta, strIDTipoRuta, udtCosteUnitario, _
                                                           udtCosteNivel, data.udtCosteAcumulado, data.FechaCalculo)
                drNuevaLineaMaterial = ProcessServer.ExecuteTask(Of DataLlenarCosteMat, DataRow)(AddressOf LlenarCosteMaterial, StCosteMat, services)

                Dim StCosteStd As New DataLlenarArtCosteStd(data.strArticuloPadre, strIDEstructura, strIDTipoEstructura, strIDRuta, strIDTipoRuta, _
                                       udtCosteUnitario, udtCosteNivel, data.udtCosteAcumulado, data.udtDtNivel, data.FechaCalculo)
                ProcessServer.ExecuteTask(Of DataLlenarArtCosteStd)(AddressOf LlenarArticuloCosteStd, StCosteStd, services)

            End If
            If Not drNuevaLineaMaterial Is Nothing Then
                If IsNothing(data.udtDtNivel.dtMaterial) Then data.udtDtNivel.dtMaterial = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf HistoricoMaterialAddNew, Nothing, services)
                data.udtDtNivel.dtMaterial.Rows.Add(drNuevaLineaMaterial.ItemArray)
            End If

            StCoste.dblCosteStdA = udtCosteUnitario.dblCosteStdA

            Dim dtMoneda As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf Moneda.ObtenerMonedaA, Nothing, services)
            If Not dtMoneda Is Nothing AndAlso dtMoneda.Rows.Count > 0 Then
                StCoste.dblCosteStdB = xRound(StCoste.dblCosteStdA * dtMoneda.Rows(0)("CambioB"), dtMoneda.Rows(0)("NDecimalesPrec"))
            End If
        End If
        Return StCoste
    End Function

    <Serializable()> _
    Public Class DataCosteEstandarSecuencia
        Public IDArticulo As String
        Public Secuencia As Integer
        Public Cantidad As Double

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDArticulo As String, ByVal Secuencia As Integer, ByVal Cantidad As Double)
            Me.IDArticulo = IDArticulo
            Me.Secuencia = Secuencia
            Me.Cantidad = Cantidad
        End Sub
    End Class

    <Task()> Public Shared Function CosteEstandarSecuencia(ByVal data As DataCosteEstandarSecuencia, ByVal services As ServiceProvider) As Double
        If Len(data.IDArticulo) > 0 Then
            Dim Coste As DataTable = New ArticuloCosteEstandar().SelOnPrimaryKey(data.IDArticulo)
            If Coste.Rows.Count > 0 Then
                If IsDate(Coste.Rows(0)("FechaEstandar")) Then
                    Dim costeMat, costeOpe, costeExt, costeVar As Double
                    Dim FechaCalculo As Date = Coste.Rows(0)("FechaEstandar")

                    '//En algun caso para saber los que hay que sumar, hay que saber si estan en la secuencia de la ruta
                    Dim componentes As DataTable = AdminData.Execute("sp_CosteStdEstructuraExp_v2", False, data.IDArticulo)
                    componentes.DefaultView.Sort = "Secuencia"
                    componentes.DefaultView.RowFilter = New NumberFilterItem("Secuencia", data.Secuencia).Compose(New AdoFilterComposer)

                    Dim f As New Filter
                    '//Coste de materiales
                    f.Add(New StringFilterItem("IDArticuloPadre", data.IDArticulo))
                    f.Add(New DateFilterItem("FechaCalculo", FechaCalculo))
                    Dim materiales As DataTable
                    materiales = New BE.DataEngine().Filter("vNegHistoricoCosteMaterial", f)
                    If materiales.Rows.Count > 0 Then
                        If componentes.DefaultView.Count >= 0 Then
                            Dim index As Integer
                            materiales.DefaultView.Sort = "IDComponente"
                            For Each dr As DataRow In componentes.Rows
                                index = materiales.DefaultView.Find(dr("IDComponente"))
                                If index >= 0 Then
                                    costeMat += materiales.DefaultView(index)("CosteTotal")
                                End If
                            Next
                        End If
                    End If

                    '//Coste de operaciones internas
                    f.Clear()
                    f.Add(New StringFilterItem("IDArticuloPadre", data.IDArticulo))
                    f.Add(New DateFilterItem("FechaCalculo", FechaCalculo))
                    f.Add(New NumberFilterItem("Secuencia", data.Secuencia))
                    f.Add(New NumberFilterItem("TipoOperacion", enumtrTipoOperacion.trInterna))
                    Dim operacionesInt As DataTable
                    operacionesInt = New BE.DataEngine().Filter("vNegHistoricoCosteOperacion", f)
                    If operacionesInt.Rows.Count > 0 Then
                        costeOpe = operacionesInt.Rows(0)("CosteOperacionA")
                    End If

                    '//Coste de operaciones externas
                    f.Clear()
                    f.Add(New StringFilterItem("IDArticuloPadre", data.IDArticulo))
                    f.Add(New DateFilterItem("FechaCalculo", FechaCalculo))
                    f.Add(New NumberFilterItem("Secuencia", data.Secuencia))
                    f.Add(New NumberFilterItem("TipoOperacion", enumtrTipoOperacion.trExterna))
                    Dim operacionesExt As DataTable
                    operacionesExt = New BE.DataEngine().Filter("vNegHistoricoCosteOperacion", f)
                    If operacionesExt.Rows.Count > 0 Then
                        costeExt = operacionesExt.Rows(0)("CosteOperacionA")
                    End If

                    '//Coste Varios
                    f.Clear()
                    f.Add(New StringFilterItem("IDArticuloPadre", data.IDArticulo))
                    f.Add(New DateFilterItem("FechaCalculo", FechaCalculo))
                    Dim varios As DataTable
                    varios = New BE.DataEngine().Filter("vNegHistoricoCosteVarios", f)
                    If varios.Rows.Count > 0 Then
                        If componentes.DefaultView.Count >= 0 Then
                            Dim index As Integer
                            varios.DefaultView.Sort = "IDArticulo"
                            For Each dr As DataRow In componentes.Rows
                                index = varios.DefaultView.Find(dr("IDComponente"))
                                If index >= 0 Then
                                    costeVar += varios.DefaultView(index)("CosteVariosA")
                                End If
                            Next
                        End If
                    End If

                    Return (costeMat + costeOpe + costeExt + costeVar)
                End If
            End If
        End If
    End Function

#End Region

    <Serializable()> _
    Public Class DataCosteEstandarPresupuesto
        Public IDArticulo As String
        Public IDRuta As String
        Public IDEstructura As String

        Public Sub New()
        End Sub
        Public Sub New(ByVal IDArticulo As String, ByVal IDRuta As String, ByVal IDEstructura As String)
            Me.IDArticulo = IDArticulo
            Me.IDRuta = IDRuta
            Me.IDEstructura = IDEstructura
        End Sub
    End Class

    <Task()> Public Shared Function CosteEstandarPresupuesto(ByVal data As DataCosteEstandarPresupuesto, ByVal services As ServiceProvider) As DataSet
        If Len(data.IDArticulo) > 0 Then
            Dim FechaCalculo As Date = CDate(Date.Today.ToShortDateString & " " & Now.ToShortTimeString)
            Dim dtACStd As DataTable = New ArticuloCosteEstandar().SelOnPrimaryKey(data.IDArticulo)
            If Not dtACStd Is Nothing AndAlso dtACStd.Rows.Count > 0 Then
                Dim udtDtCoste As New udtCostesDT
                udtDtCoste.dtCosteStd = dtACStd

                Dim blnCancel As Boolean
                Dim strCommand As String
                If Len(data.IDRuta) = 0 Then
                    blnCancel = True
                    ApplicationService.GenerateError("La ruta es obligatoria")
                Else
                    If Len(data.IDEstructura) > 0 Then
                    Else
                        blnCancel = True
                        ApplicationService.GenerateError("La estructura es obligatoria")
                    End If
                End If

                If Not blnCancel Then
                    Dim dtComponentes As DataTable = AdminData.Execute("sp_CosteStdEstructuraExp", False, data.IDArticulo, data.IDEstructura)
                    If Not dtComponentes Is Nothing AndAlso dtComponentes.Rows.Count > 0 Then
                        Dim dtCopy As DataTable = dtComponentes.Copy

                        dtCopy.DefaultView.RowFilter = "IDComponente= '" & data.IDArticulo & "'"
                        If dtCopy.DefaultView.Count > 0 Then
                            If dtCopy.DefaultView(0).Row("Nivel") <> 0 Then
                                ApplicationService.GenerateError("Este componente no puede formar parte de su propia estructura")
                            Else
                                Dim p As New Parametro
                                Dim intCriterioValoracion As Integer = p.CriterioValoracionCosteStd()
                                Dim udtCosteAcumulado As New udtCosteStd

                                Dim StDataComp As New DataCosteEstandarComponente(dtComponentes, data.IDArticulo, data.IDArticulo, "0", _
                                              data.IDRuta, 1, 1, 0, intCriterioValoracion, _
                                              udtCosteAcumulado, udtDtCoste, FechaCalculo)
                                ProcessServer.ExecuteTask(Of DataCosteEstandarComponente, udtCosteStdUnitario)(AddressOf CosteEstandarComponente, StDataComp, services)
                                ProcessServer.ExecuteTask(Of udtCostesDT)(AddressOf AplicarDecimales, udtDtCoste, services)

                                Dim ds As New DataSet
                                If Not udtDtCoste.dtCosteStd Is Nothing Then ds.Tables.Add(udtDtCoste.dtCosteStd)
                                If Not udtDtCoste.dtMaterial Is Nothing Then ds.Tables.Add(udtDtCoste.dtMaterial)
                                If Not udtDtCoste.dtOperacion Is Nothing Then ds.Tables.Add(udtDtCoste.dtOperacion)
                                If Not udtDtCoste.dtVarios Is Nothing Then ds.Tables.Add(udtDtCoste.dtVarios)
                                Return ds
                            End If
                        End If
                    Else : ApplicationService.GenerateError("Error en el procedimiento almacenado.")
                    End If
                End If
            Else : ApplicationService.GenerateError("El artículo | no existe.", data.IDArticulo)
            End If
        Else : ApplicationService.GenerateError("El Artículo es un dato obligatorio.")
        End If
    End Function

    <Serializable()> _
    Public Class DataCosteStdOperaciones
        Public strArticuloPadre As String
        Public strArticulo As String
        Public dblLote As Double
        Public dblQAcumulada As Double
        Public intNivel As Integer
        Public intOrden As Integer
        Public dtOperacion As DataTable
        Public FechaCalculo As Date
        Public strIDRuta As String = Nothing

        Public Sub New()
        End Sub
        Public Sub New(ByVal strArticuloPadre As String, ByVal strArticulo As String, _
                                              ByVal dblLote As Double, ByVal dblQAcumulada As Double, _
                                              ByVal intNivel As Integer, ByVal intOrden As Integer, _
                                              ByRef dtOperacion As DataTable, _
                                              ByVal FechaCalculo As Date, _
                                              Optional ByVal strIDRuta As String = Nothing)
            Me.strArticuloPadre = strArticuloPadre
            Me.strArticulo = strArticulo
            Me.dblLote = dblLote
            Me.dblQAcumulada = dblQAcumulada
            Me.intNivel = intNivel
            Me.intOrden = intOrden
            Me.dtOperacion = dtOperacion
            Me.FechaCalculo = FechaCalculo
            Me.strIDRuta = strIDRuta
        End Sub
    End Class

    <Task()> Public Shared Function CosteEstandarOperaciones(ByVal data As DataCosteStdOperaciones, ByVal services As ServiceProvider) As udtCosteStdOpe
        If Length(data.strArticulo) > 0 Then
            Const VIEW_NAME As String = "vNegCosteStdOpeExt"

            Dim TotalCostesOperaciones As New udtCosteStdOpe

            If data.dblLote = 0 Then data.dblLote = 1
            Dim dtRuta As DataTable

            If Length(data.strIDRuta) > 0 Then
                dtRuta = AdminData.Execute("sp_CosteStdRutaArticulo", False, data.strArticulo, data.strIDRuta)
            Else
                dtRuta = AdminData.Execute("sp_CosteStdRutaArticulo", False, data.strArticulo)
            End If

            If Not dtRuta Is Nothing AndAlso dtRuta.Rows.Count > 0 Then
                Dim dblFactorProduccion, dblIntTotalPrepA, dblIntTotalEjecA, dblIntTotalMODA, dblExtTotalA As Double
                Dim dblCosteFijo, dblCosteVariable, dblCosteDirecto, dblCosteIndirecto As Double
                Dim strIDProveedor As String
                Dim drNuevaLinea As DataRow
                Dim dblCosteA As Double

                For Each drRuta As DataRow In dtRuta.Rows
                    For Each c As DataColumn In dtRuta.Columns
                        If AreEquals("Tasa", Left(c.ColumnName, 4)) Then
                            If IsDBNull(drRuta(c)) Then drRuta(c) = 0
                        End If
                    Next
                    dblFactorProduccion = IIf(drRuta("FactorProduccion") > 0, drRuta("FactorProduccion"), 1)
                    dblCosteFijo = 0
                    dblCosteVariable = 0
                    dblCosteDirecto = 0
                    dblCosteIndirecto = 0
                    Select Case CType(drRuta("TipoOperacion"), enumtrTipoOperacion)
                        Case enumtrTipoOperacion.trInterna
                            '///Preparacion
                            Dim StDataTiempo As New DataTiempoOperacion(drRuta("TiempoPrep"), drRuta("UdTiempoPrep"))
                            Dim dblTiempo As Double = ProcessServer.ExecuteTask(Of DataTiempoOperacion, Double)(AddressOf TiempoOperacion, StDataTiempo, services)
                            Dim dblTasaPreparacionA As Double = ((dblTiempo * drRuta("TasaPreparacionA")) / data.dblLote) / dblFactorProduccion
                            dblIntTotalPrepA = dblIntTotalPrepA + dblTasaPreparacionA

                            '///Preparacion Fija
                            Dim dblPrepFija As Double = ((dblTiempo * drRuta("TasaPrepFija")) / data.dblLote) / dblFactorProduccion
                            dblCosteFijo = dblCosteFijo + dblPrepFija
                            '///Preparacion Variable
                            Dim dblPrepVar As Double = ((dblTiempo * drRuta("TasaPrepVar")) / data.dblLote) / dblFactorProduccion
                            dblCosteVariable = dblCosteVariable + dblPrepVar
                            '///Preparacion Directa
                            Dim dblPrepDir As Double = ((dblTiempo * drRuta("TasaPrepDir")) / data.dblLote) / dblFactorProduccion
                            dblCosteDirecto = dblCosteDirecto + dblPrepDir
                            '///Preparacion Indirecta
                            Dim dblPrepInd As Double = ((dblTiempo * drRuta("TasaPrepInd")) / data.dblLote) / dblFactorProduccion
                            dblCosteIndirecto = dblCosteIndirecto + dblPrepInd

                            '///Ejecucion
                            'Hay que diferenciar si son operaciones de ciclo o no
                            If drRuta("Ciclo") Then
                                Dim StData As New DataTiempoOperacion(drRuta("TiempoCiclo"), drRuta("UDTiempoCiclo"))
                                dblTiempo = ProcessServer.ExecuteTask(Of DataTiempoOperacion, Double)(AddressOf TiempoOperacion, StData, services)
                                dblTiempo = dblTiempo / IIf(Nz(drRuta("Loteciclo"), 1) > 0, Nz(drRuta("LoteCiclo"), 1), 1)
                                'Como TiempoEjecUnit no se rellena manualmente en las operaciones de ciclos en el Mnto.Articulo,
                                'se guarda en él, la relación TiempoCiclo/LoteCiclo (con su unidad) para mostrarla en Mnto.Coste Estándar.
                                drRuta("TiempoEjecUnit") = drRuta("TiempoCiclo") / IIf(Nz(drRuta("Loteciclo"), 1) > 0, Nz(drRuta("LoteCiclo"), 1), 1)
                                drRuta("UdTiempoEjec") = drRuta("UDTiempoCiclo")
                            Else
                                Dim StData As New DataTiempoOperacion(drRuta("TiempoEjecUnit"), drRuta("UdTiempoEjec"))
                                dblTiempo = ProcessServer.ExecuteTask(Of DataTiempoOperacion, Double)(AddressOf TiempoOperacion, StData, services)
                            End If
                            Dim dblTasaEjecucionA As Double = (dblTiempo * drRuta("TasaEjecucionA")) / dblFactorProduccion
                            dblIntTotalEjecA = dblIntTotalEjecA + dblTasaEjecucionA

                            Dim dblEjecFija As Double = (dblTiempo * drRuta("TasaEjecFija")) / dblFactorProduccion
                            dblCosteFijo = dblCosteFijo + dblEjecFija
                            Dim dblEjecVar As Double = (dblTiempo * drRuta("TasaEjecVar")) / dblFactorProduccion
                            dblCosteVariable = dblCosteVariable + dblEjecVar
                            Dim dblEjecDir As Double = (dblTiempo * drRuta("TasaEjecDir")) / dblFactorProduccion
                            dblCosteDirecto = dblCosteDirecto + dblEjecDir
                            Dim dblEjecInd As Double = (dblTiempo * drRuta("TasaEjecInd")) / dblFactorProduccion
                            dblCosteIndirecto = dblCosteIndirecto + dblEjecInd

                            '///MOD
                            Dim dblTasaMODA As Double = (drRuta("FactorHombre") * drRuta("TasaManoObraA") * dblTiempo) / dblFactorProduccion
                            dblIntTotalMODA = dblIntTotalMODA + dblTasaMODA

                            Dim dblMODFija As Double = (drRuta("FactorHombre") * drRuta("TasaMODFija") * dblTiempo) / dblFactorProduccion
                            dblCosteFijo = dblCosteFijo + dblMODFija
                            Dim dblMODVar As Double = (drRuta("FactorHombre") * drRuta("TasaMODVar") * dblTiempo) / dblFactorProduccion
                            dblCosteVariable = dblCosteVariable + dblMODVar
                            Dim dblMODDir As Double = (drRuta("FactorHombre") * drRuta("TasaMODDir") * dblTiempo) / dblFactorProduccion
                            dblCosteDirecto = dblCosteDirecto + dblMODDir
                            Dim dblMODInd As Double = (drRuta("FactorHombre") * drRuta("TasaMODInd") * dblTiempo) / dblFactorProduccion
                            dblCosteIndirecto = dblCosteIndirecto + dblMODInd

                            dblCosteA = dblTasaPreparacionA + dblTasaEjecucionA + dblTasaMODA
                            dblCosteA = dblCosteA * data.dblQAcumulada
                            dblCosteFijo = dblCosteFijo * data.dblQAcumulada
                            dblCosteVariable = dblCosteVariable * data.dblQAcumulada
                            dblCosteDirecto = dblCosteDirecto * data.dblQAcumulada
                            dblCosteIndirecto = dblCosteIndirecto * data.dblQAcumulada
                        Case enumtrTipoOperacion.trExterna
                            dblCosteA = 0
                            Dim dtOpeExterna As DataTable = AdminData.Filter(VIEW_NAME, , "IDRutaOp=" & drRuta("IDRutaOp"))
                            If Not dtOpeExterna Is Nothing AndAlso dtOpeExterna.Rows.Count > 0 Then
                                Dim dtOpeExternaCopy As DataTable = dtOpeExterna.Copy
                                Dim dvOpeExterna As DataView = dtOpeExternaCopy.DefaultView
                                dvOpeExterna.RowFilter = "Principal=1"

                                If dvOpeExterna.Count > 0 Then
                                    strIDProveedor = dvOpeExterna(0).Row("IDProveedor")
                                    dvOpeExterna.RowFilter = "IDProveedor='" & strIDProveedor & "'"
                                End If
                                If dvOpeExterna.Count > 0 Then
                                    Dim dblFactorProveedor As Double
                                    Dim StDatos As New ArticuloUnidadAB.DatosFactorConversion
                                    StDatos.IDArticulo = dvOpeExterna(0).Row("IDArticulo")
                                    StDatos.IDUdMedidaA = dvOpeExterna(0).Row("IDudProduccion")
                                    StDatos.IDUdMedidaB = dvOpeExterna(0).Row("IDudInterna")
                                    dblFactorProveedor = ProcessServer.ExecuteTask(Of ArticuloUnidadAB.DatosFactorConversion, Double)(AddressOf ArticuloUnidadAB.FactorDeConversion, StDatos, services)
                                    dblFactorProduccion = IIf(dblFactorProveedor > 0, dblFactorProveedor, dblFactorProduccion)
                                    dvOpeExterna.Sort = "QDesde"
                                    Dim udValoracion As Double
                                    If Nz(dvOpeExterna(0).Row("UDValoracion"), 1) = 0 Then
                                        udValoracion = 1
                                    Else
                                        udValoracion = Nz(dvOpeExterna(0).Row("UDValoracion"), 1)
                                    End If
                                    dblCosteA = (dvOpeExterna(0).Row("PrecioA") / udValoracion) / dblFactorProduccion
                                    For Each drv As DataRowView In dvOpeExterna
                                        If data.dblQAcumulada >= drv("QDesde") Then
                                            dblCosteA = (drv("PrecioA") / udValoracion) / dblFactorProduccion
                                        Else
                                            Exit For
                                        End If
                                    Next
                                End If
                            End If

                            dblExtTotalA = dblExtTotalA + dblCosteA
                            dblCosteA = dblCosteA * data.dblQAcumulada
                            dblCosteDirecto = dblCosteDirecto + dblCosteA
                            dblCosteVariable = dblCosteVariable + dblCosteA
                    End Select

                    Dim StCosteOper As New DataLlenarCosteOpe(data.strArticuloPadre, data.strArticulo, data.intNivel, data.intOrden, _
                                                        data.dblLote, dblCosteA, dblCosteDirecto, dblCosteIndirecto, _
                                                        dblCosteFijo, dblCosteVariable, strIDProveedor, drRuta, data.FechaCalculo, data.dblQAcumulada)
                    drNuevaLinea = ProcessServer.ExecuteTask(Of DataLlenarCosteOpe, DataRow)(AddressOf LlenarCosteOperacion, StCosteOper, services)

                    If Not drNuevaLinea Is Nothing Then
                        If IsNothing(data.dtOperacion) Then data.dtOperacion = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf HistoricoOperacionAddNew, Nothing, services)
                        data.dtOperacion.Rows.Add(drNuevaLinea.ItemArray)
                    End If
                    strIDProveedor = String.Empty
                Next

                With TotalCostesOperaciones
                    .dblCosteIntA = dblIntTotalPrepA + dblIntTotalEjecA + dblIntTotalMODA
                    .dblCosteExtA = dblExtTotalA
                    .dblCosteFijo = dblCosteFijo
                    .dblCosteDirecto = dblCosteDirecto
                    .dblCosteIndirecto = dblCosteIndirecto
                    .dblCosteVariable = dblCosteVariable
                End With
            End If
            Return TotalCostesOperaciones
        End If
    End Function

    <Serializable()> _
    Public Class DataCosteStdVarios
        Public strArticuloPadre As String
        Public strArticulo As String
        Public dblQAcumulada As Double
        Public intNivel As Integer
        Public intOrden As Integer
        Public udtCostes As udtCosteStd
        Public dtCoste As DataTable
        Public FechaCalculo As Date

        Public Sub New()
        End Sub
        Public Sub New(ByVal strArticuloPadre As String, ByVal strArticulo As String, _
                       ByVal dblQAcumulada As Double, ByVal intNivel As Integer, _
                       ByVal intOrden As Integer, ByRef udtCostes As udtCosteStd, _
                       ByRef dtCoste As DataTable, _
                       ByVal FechaCalculo As Date)
            Me.strArticuloPadre = strArticuloPadre
            Me.strArticulo = strArticulo
            Me.dblQAcumulada = dblQAcumulada
            Me.intNivel = intNivel
            Me.intOrden = intOrden
            Me.udtCostes = udtCostes
            Me.dtCoste = dtCoste
            Me.FechaCalculo = FechaCalculo
        End Sub
    End Class

    <Task()> Public Shared Function CosteEstandarVarios(ByVal data As DataCosteStdVarios, ByVal services As ServiceProvider) As udtCosteStdVar
        If Length(data.strArticulo) > 0 Then
            Dim TotalCostesVarios As New udtCosteStdVar
            Dim dblCosteVarA, dblCosteFijo, dblCosteVariable, dblCosteDirecto, dblCosteIndirecto As Double
            Dim av As New ArticuloVarios

            Dim dtVarios As DataTable = av.Filter(, "IDArticulo='" & data.strArticulo & "'")
            If Not dtVarios Is Nothing AndAlso dtVarios.Rows.Count > 0 Then
                Dim dblVarA, dblPVPA As Double
                Dim drNuevaLinea As DataRow

                For Each drVarios As DataRow In dtVarios.Rows
                    With data.udtCostes
                        Select Case CType(drVarios("Tipo"), enumCosteVarios)
                            Case enumCosteVarios.cvValor
                                dblVarA = drVarios("Valor")
                            Case enumCosteVarios.cvPorMaterial
                                dblVarA = .udtMaterial.dblCosteMatA * (drVarios("Valor") / 100)
                            Case enumCosteVarios.cvPorInterno
                                dblVarA = .udtOperacion.dblCosteIntA * (drVarios("Valor") / 100)
                            Case enumCosteVarios.cvPorExterno
                                dblVarA = .udtOperacion.dblCosteExtA * (drVarios("Valor") / 100)
                            Case enumCosteVarios.cvPorTotal
                                dblVarA = (.udtMaterial.dblCosteMatA + .udtOperacion.dblCosteIntA + .udtOperacion.dblCosteExtA) * (drVarios("Valor") / 100)
                            Case enumCosteVarios.cvPorPVP
                                dblPVPA = ProcessServer.ExecuteTask(Of String, Double)(AddressOf CosteEstandarPVP, data.strArticulo, services)
                                dblVarA = dblPVPA * (drVarios("Valor") / 100)
                        End Select
                        dblCosteVarA = dblCosteVarA + dblVarA
                        If IsDBNull(drVarios("TipoCosteFV")) OrElse drVarios("TipoCosteFV") = enumtcfvTipoCoste.tcfvFijo Then
                            dblCosteFijo = dblCosteFijo + dblVarA
                        Else
                            dblCosteVariable = dblCosteVariable + dblVarA
                        End If
                        If IsDBNull(drVarios("TipoCosteDI")) OrElse drVarios("TipoCosteDI") = enumtcdiTipoCoste.tcdiDirecto Then
                            dblCosteDirecto = dblCosteDirecto + dblVarA
                        Else
                            dblCosteIndirecto = dblCosteIndirecto + dblVarA
                        End If
                    End With

                    Dim StData As New DataLlenarCosteVar(data.strArticuloPadre, drVarios, data.intNivel, data.intOrden, dblVarA, _
                                                     dblCosteDirecto, dblCosteIndirecto, dblCosteFijo, dblCosteVariable, data.FechaCalculo)
                    drNuevaLinea = ProcessServer.ExecuteTask(Of DataLlenarCosteVar, DataRow)(AddressOf LlenarCosteVarios, StData, services)
                    If Not drNuevaLinea Is Nothing Then
                        If IsNothing(data.dtCoste) Then data.dtCoste = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf HistoricoVariosAddNew, Nothing, services)
                        data.dtCoste.Rows.Add(drNuevaLinea.ItemArray)
                    End If
                Next
                With TotalCostesVarios
                    .dblCosteVarA = dblCosteVarA
                    .dblCosteFijo = dblCosteFijo
                    .dblCosteVariable = dblCosteVariable
                    .dblCosteDirecto = dblCosteDirecto
                    .dblCosteIndirecto = dblCosteIndirecto
                End With
            End If
            Return (TotalCostesVarios)
        End If
    End Function

    <Task()> Public Shared Function CosteEstandarPVP(ByVal strIDArticulo As String, ByVal services As ServiceProvider) As Double
        Dim dblPVPA As Double
        If Len(strIDArticulo) > 0 Then
            Dim TA As New TarifaArticulo
            Dim dtPVP As DataTable = ProcessServer.ExecuteTask(Of String, DataTable)(AddressOf TA.PVP, strIDArticulo, services)
            If Not dtPVP Is Nothing AndAlso dtPVP.Rows.Count > 0 Then
                dblPVPA = dtPVP.Rows(0)("PVPA")
            End If
        End If
        Return dblPVPA
    End Function

    <Task()> Public Shared Function HistoricoMaterialAddNew(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim hcm As New HistoricoCosteMaterial
        Dim dt As DataTable = hcm.AddNew
        Return dt
    End Function

    <Task()> Public Shared Function HistoricoOperacionAddNew(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim hco As New HistoricoCosteOperacion
        Dim dt As DataTable = hco.AddNew
        Return dt
    End Function

    <Task()> Public Shared Function HistoricoVariosAddNew(ByVal data As Object, ByVal services As ServiceProvider) As DataTable
        Dim hcv As New HistoricoCosteVarios
        Dim dt As DataTable = hcv.AddNew
        Return dt
    End Function

    <Serializable()> _
    Public Class DataTiempoOperacion
        Public DblTiempo As Double
        Public UdTiempo As enumstdUdTiempo

        Public Sub New()
        End Sub
        Public Sub New(ByVal dblTiempo As Double, ByVal UdTiempo As enumstdUdTiempo)
            Me.DblTiempo = dblTiempo
            Me.UdTiempo = UdTiempo
        End Sub
    End Class

    <Task()> Public Shared Function TiempoOperacion(ByVal data As DataTiempoOperacion, ByVal services As ServiceProvider) As Double
        Dim dblCalculo As Double
        Select Case data.UdTiempo
            Case enumstdUdTiempo.Dias
                dblCalculo = data.DblTiempo * 24
            Case enumstdUdTiempo.Horas
                dblCalculo = data.DblTiempo
            Case enumstdUdTiempo.Minutos
                dblCalculo = data.DblTiempo / 60
            Case enumstdUdTiempo.Segundos
                dblCalculo = data.DblTiempo / 3600
        End Select
        Return dblCalculo
    End Function

    <Serializable()> _
    Public Class DataLlenarArtCosteStd
        Public strArticulo As String
        Public strIDEstructura As String
        Public strIDTipoEstructura As String
        Public strIDRuta As String
        Public strIDTipoRuta As String
        Public udtCosteUnitario As udtCosteStdUnitario
        Public udtCosteNivel As udtCosteStd
        Public udtCosteAcumulado As udtCosteStd
        Public udtDtNivel As udtCostesDT
        Public FechaCalculo As Date

        Public Sub New()
        End Sub
        Public Sub New(ByVal strArticulo As String, ByVal strIDEstructura As String, _
                       ByVal strIDTipoEstructura As String, ByVal strIDRuta As String, _
                       ByVal strIDTipoRuta As String, _
                       ByRef udtCosteUnitario As udtCosteStdUnitario, _
                       ByRef udtCosteNivel As udtCosteStd, _
                       ByRef udtCosteAcumulado As udtCosteStd, _
                       ByRef udtDtNivel As udtCostesDT, _
                       ByVal FechaCalculo As Date)
            Me.strArticulo = strArticulo
            Me.strIDEstructura = strIDEstructura
            Me.strIDTipoEstructura = strIDTipoEstructura
            Me.strIDRuta = strIDRuta
            Me.strIDTipoRuta = strIDTipoRuta
            Me.udtCosteUnitario = udtCosteUnitario
            Me.udtCosteNivel = udtCosteNivel
            Me.udtCosteAcumulado = udtCosteAcumulado
            Me.udtDtNivel = udtDtNivel
            Me.FechaCalculo = FechaCalculo
        End Sub
    End Class

    <Task()> Public Shared Sub LlenarArticuloCosteStd(ByVal data As DataLlenarArtCosteStd, ByVal services As ServiceProvider)
        Dim drArticulo As DataRow
        With data.udtDtNivel
            If IsNothing(.dtCosteStd) Then
                drArticulo = New ArticuloCosteEstandar().AddNewForm.NewRow
                drArticulo("IDArticulo") = data.strArticulo

                .dtCosteStd = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf HistoricoMaterialAddNew, Nothing, services)
                .dtCosteStd.Rows.Add(drArticulo.ItemArray)
            End If

            .dtCosteStd.Rows(0)("FechaUltimo") = data.FechaCalculo

            Dim DtArt As DataTable = AdminData.Filter("frmMntoArticulos", "IDRuta,IDTipoRuta,IDEstructura,IDTipoEstructura", "IDArticulo = '" & data.strArticulo & "'")
            If Not DtArt Is Nothing AndAlso DtArt.Rows.Count > 0 Then
                .dtCosteStd.Rows(0)("IDRuta") = DtArt.Rows(0)("IDRuta")
                .dtCosteStd.Rows(0)("IDTipoRuta") = DtArt.Rows(0)("IDTipoRuta")
                .dtCosteStd.Rows(0)("IDEstructura") = DtArt.Rows(0)("IDEstructura")
                .dtCosteStd.Rows(0)("IDTipoEstructura") = DtArt.Rows(0)("IDTipoEstructura")
            End If

            If Len(data.strIDEstructura) > 0 Then .dtCosteStd.Rows(0)("IDEstructura") = data.strIDEstructura
            If Len(data.strIDTipoEstructura) > 0 Then .dtCosteStd.Rows(0)("IDTipoEstructura") = data.strIDTipoEstructura
            If Len(data.strIDRuta) > 0 Then .dtCosteStd.Rows(0)("IDRuta") = data.strIDRuta
            If Len(data.strIDTipoRuta) > 0 Then .dtCosteStd.Rows(0)("IDTipoRuta") = data.strIDTipoRuta

            .dtCosteStd.Rows(0)("CosteUltimoA") = data.udtCosteUnitario.dblCosteStdA
            Dim dtMoneda As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf Moneda.ObtenerMonedaA, Nothing, services)
            If Not dtMoneda Is Nothing AndAlso dtMoneda.Rows.Count > 0 Then
                .dtCosteStd.Rows(0)("CosteUltimoB") = xRound(data.udtCosteUnitario.dblCosteStdA * dtMoneda.Rows(0)("CambioB"), dtMoneda.Rows(0)("NDecimalesPrec"))
            End If
            .dtCosteStd.Rows(0)("CosteMatUltA") = data.udtCosteNivel.udtMaterial.dblCosteMatA
            .dtCosteStd.Rows(0)("CosteOpeUltA") = data.udtCosteNivel.udtOperacion.dblCosteIntA
            .dtCosteStd.Rows(0)("CosteExtUltA") = data.udtCosteNivel.udtOperacion.dblCosteExtA
            .dtCosteStd.Rows(0)("CosteVarUltA") = data.udtCosteNivel.udtVarios.dblCosteVarA
            .dtCosteStd.Rows(0)("CosteAcuMatUltA") = data.udtCosteAcumulado.udtMaterial.dblCosteMatA
            .dtCosteStd.Rows(0)("CosteAcuOpeUltA") = data.udtCosteAcumulado.udtOperacion.dblCosteIntA
            .dtCosteStd.Rows(0)("CosteAcuExtUltA") = data.udtCosteAcumulado.udtOperacion.dblCosteExtA
            .dtCosteStd.Rows(0)("CosteAcuVarUltA") = data.udtCosteAcumulado.udtVarios.dblCosteVarA
            .dtCosteStd.Rows(0)("CosteVariableUltA") = .dtCosteStd.Rows(0)("CosteMatUltA") + data.udtCosteNivel.udtVarios.dblCosteVariable + data.udtCosteNivel.udtOperacion.dblCosteVariable
            .dtCosteStd.Rows(0)("CosteDirectoUltA") = .dtCosteStd.Rows(0)("CosteMatUltA") + data.udtCosteNivel.udtVarios.dblCosteDirecto + data.udtCosteNivel.udtOperacion.dblCosteDirecto
            .dtCosteStd.Rows(0)("CosteFijoUltA") = data.udtCosteNivel.udtVarios.dblCosteFijo + data.udtCosteNivel.udtOperacion.dblCosteFijo
            .dtCosteStd.Rows(0)("CosteIndirectoUltA") = data.udtCosteNivel.udtVarios.dblCosteIndirecto + data.udtCosteNivel.udtOperacion.dblCosteIndirecto
        End With

    End Sub

    <Serializable()> _
    Public Class DataLlenarCosteMat
        Public strArticuloPadre As String
        Public drComponente As DataRow
        Public dblQAcumulada As Double
        Public intOrden As Integer
        Public strIDRuta As String
        Public strIDTipoRuta As String
        Public udtCosteUnitario As udtCosteStdUnitario
        Public udtCosteNivel As udtCosteStd
        Public udtCosteAcumulado As udtCosteStd
        Public FechaCalculo As Date

        Public Sub New()
        End Sub
        Public Sub New(ByVal strArticuloPadre As String, ByVal drComponente As DataRow, _
                       ByVal dblQAcumulada As Double, ByVal intOrden As Integer, _
                       ByVal strIDRuta As String, ByVal strIDTipoRuta As String, _
                       ByRef udtCosteUnitario As udtCosteStdUnitario, _
                       ByRef udtCosteNivel As udtCosteStd, _
                       ByRef udtCosteAcumulado As udtCosteStd, ByVal FechaCalculo As Date)
            Me.strArticuloPadre = strArticuloPadre
            Me.drComponente = drComponente
            Me.dblQAcumulada = dblQAcumulada
            Me.intOrden = intOrden
            Me.strIDRuta = strIDRuta
            Me.strIDTipoRuta = strIDTipoRuta
            Me.udtCosteUnitario = udtCosteUnitario
            Me.udtCosteNivel = udtCosteNivel
            Me.udtCosteAcumulado = udtCosteAcumulado
            Me.FechaCalculo = FechaCalculo
        End Sub
    End Class

    <Task()> Public Shared Function LlenarCosteMaterial(ByVal data As DataLlenarCosteMat, ByVal services As ServiceProvider) As DataRow
        Dim drMaterial As DataRow = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf HistoricoMaterialAddNew, Nothing, services).NewRow

        drMaterial("IDCosteMaterial") = AdminData.GetAutoNumeric
        drMaterial("IDArticuloPadre") = data.strArticuloPadre
        If data.drComponente("Nivel") = 0 Then
            drMaterial("IDArticulo") = data.strArticuloPadre
        Else
            drMaterial("IDArticulo") = data.drComponente("IDPadre")
        End If
        drMaterial("FechaCalculo") = data.FechaCalculo
        drMaterial("IDEstrComp") = data.drComponente("IDEstrComp")
        drMaterial("IDEstructura") = data.drComponente("IDEstructura")
        drMaterial("IDComponente") = data.drComponente("IDComponente")
        drMaterial("Cantidad") = data.drComponente("Cantidad")
        drMaterial("Merma") = data.drComponente("Merma")
        drMaterial("Nivel") = data.drComponente("Nivel")
        drMaterial("CantidadAcumulada") = data.dblQAcumulada
        drMaterial("Orden") = data.intOrden
        If data.drComponente("Fabrica") Then
            drMaterial("Tipo") = CInt(enumacsTipoArticulo.acsFabrica)
        Else
            drMaterial("Tipo") = CInt(enumacsTipoArticulo.acsCompra)
        End If
        If Len(data.strIDRuta) > 0 Then drMaterial("IDRuta") = data.strIDRuta
        If Len(data.strIDTipoRuta) > 0 Then drMaterial("IDTipoRuta") = data.strIDTipoRuta

        drMaterial("CosteStdA") = data.udtCosteUnitario.dblCosteStdA
        drMaterial("CosteMatStdA") = data.udtCosteNivel.udtMaterial.dblCosteMatA
        drMaterial("CosteOpeStdA") = data.udtCosteNivel.udtOperacion.dblCosteIntA
        drMaterial("CosteExtStdA") = data.udtCosteNivel.udtOperacion.dblCosteExtA
        drMaterial("CosteVarStdA") = data.udtCosteNivel.udtVarios.dblCosteVarA

        drMaterial("CosteAcuMatStdA") = data.udtCosteAcumulado.udtMaterial.dblCosteMatA
        drMaterial("CosteAcuOpeStdA") = data.udtCosteAcumulado.udtOperacion.dblCosteIntA
        drMaterial("CosteAcuExtStdA") = data.udtCosteAcumulado.udtOperacion.dblCosteExtA
        drMaterial("CosteAcuVarStdA") = data.udtCosteAcumulado.udtVarios.dblCosteVarA

        drMaterial("CosteVariableStdA") = drMaterial("CosteMatStdA") * Nz(drMaterial("Cantidad"), 0) + data.udtCosteNivel.udtVarios.dblCosteVariable + data.udtCosteNivel.udtOperacion.dblCosteVariable
        drMaterial("CosteDirectoStdA") = drMaterial("CosteMatStdA") * Nz(drMaterial("Cantidad"), 0) + data.udtCosteNivel.udtVarios.dblCosteDirecto + data.udtCosteNivel.udtOperacion.dblCosteDirecto
        drMaterial("CosteFijoStdA") = data.udtCosteNivel.udtVarios.dblCosteFijo + data.udtCosteNivel.udtOperacion.dblCosteFijo
        drMaterial("CosteIndirectoStdA") = data.udtCosteNivel.udtVarios.dblCosteIndirecto + data.udtCosteNivel.udtOperacion.dblCosteIndirecto

        Return drMaterial
    End Function

    <Serializable()> _
    Public Class DataLlenarCosteOpe
        Public strArticuloPadre As String
        Public strArticulo As String
        Public intNivel As Integer
        Public intOrden As Integer
        Public dblLoteMinimo As Double
        Public dblCosteA As Double
        Public dblCosteDirecto As Double
        Public dblCosteIndirecto As Double
        Public dblCosteFijo As Double
        Public dblCosteVariable As Double
        Public strIDProveedor As String
        Public drRuta As DataRow
        Public FechaCalculo As Date
        Public QAcumulada As Double

        Public Sub New()
        End Sub
        Public Sub New(ByVal strArticuloPadre As String, ByVal strArticulo As String, _
                       ByVal intNivel As Integer, ByVal intOrden As Integer, _
                       ByVal dblLoteMinimo As Double, ByVal dblCosteA As Double, _
                       ByVal dblCosteDirecto As Double, ByVal dblCosteIndirecto As Double, _
                       ByVal dblCosteFijo As Double, ByVal dblCosteVariable As Double, _
                       ByVal strIDProveedor As String, ByVal drRuta As DataRow, _
                       ByVal FechaCalculo As Date, ByVal QAcumulada As Double)
            Me.strArticuloPadre = strArticuloPadre
            Me.strArticulo = strArticulo
            Me.intNivel = intNivel
            Me.intOrden = intOrden
            Me.dblLoteMinimo = dblLoteMinimo
            Me.dblCosteA = dblCosteA
            Me.dblCosteDirecto = dblCosteDirecto
            Me.dblCosteIndirecto = dblCosteIndirecto
            Me.dblCosteFijo = dblCosteFijo
            Me.dblCosteVariable = dblCosteVariable
            Me.strIDProveedor = strIDProveedor
            Me.drRuta = drRuta
            Me.FechaCalculo = FechaCalculo
            Me.QAcumulada = QAcumulada
        End Sub
    End Class

    <Task()> Public Shared Function LlenarCosteOperacion(ByVal data As DataLlenarCosteOpe, ByVal services As ServiceProvider) As DataRow
        Dim drOperacion As DataRow = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf HistoricoOperacionAddNew, Nothing, services).NewRow

        drOperacion("IDCosteOperacion") = AdminData.GetAutoNumeric
        drOperacion("IDArticuloPadre") = data.strArticuloPadre
        drOperacion("IDArticulo") = data.strArticulo
        drOperacion("FechaCalculo") = data.FechaCalculo
        drOperacion("IDRutaOp") = data.drRuta("IDRutaOp")
        drOperacion("IDRuta") = data.drRuta("IDRuta")
        drOperacion("IDTipoRuta") = data.drRuta("IDTipoRuta")
        drOperacion("Secuencia") = data.drRuta("Secuencia")
        drOperacion("TipoOperacion") = data.drRuta("TipoOperacion")
        drOperacion("IDOperacion") = data.drRuta("IDOperacion")
        drOperacion("DescOperacion") = data.drRuta("DescOperacion")
        drOperacion("IDCentro") = data.drRuta("IDCentro")
        drOperacion("FactorHombre") = data.drRuta("FactorHombre")
        drOperacion("TiempoPrep") = data.drRuta("TiempoPrep")
        drOperacion("UdTiempoPrep") = data.drRuta("UdTiempoPrep")
        drOperacion("TiempoEjecUnit") = data.drRuta("TiempoEjecUnit")
        drOperacion("UdTiempoEjec") = data.drRuta("UdTiempoEjec")
        drOperacion("FactorProduccion") = data.drRuta("FactorProduccion")

        drOperacion("TasaEjecucionA") = data.drRuta("TasaEjecucionA")
        drOperacion("TasaPreparacionA") = data.drRuta("TasaPreparacionA")
        drOperacion("TasaMODA") = data.drRuta("TasaManoObraA")
        drOperacion("CantidadAcumulada") = data.QAcumulada
        drOperacion("CosteOperacionA") = data.dblCosteA
        If drOperacion("TipoOperacion") = enumtrTipoOperacion.trExterna Then
            drOperacion("CosteVariableA") = data.dblCosteA
            drOperacion("CosteDirectoA") = data.dblCosteA
        Else
            drOperacion("CosteDirectoA") = data.dblCosteDirecto
            drOperacion("CosteIndirectoA") = data.dblCosteIndirecto
            drOperacion("CosteFijoA") = data.dblCosteFijo
            drOperacion("CosteVariableA") = data.dblCosteVariable
        End If
        drOperacion("Nivel") = data.intNivel
        drOperacion("Orden") = data.intOrden
        drOperacion("LoteMinimo") = data.dblLoteMinimo
        If Len(data.strIDProveedor) > 0 Then drOperacion("IDProveedor") = data.strIDProveedor
        'Campos para operaciones internas de ciclo
        drOperacion("TiempoCiclo") = data.drRuta("TiempoCiclo")
        drOperacion("UDTiempoCiclo") = data.drRuta("UDTiempoCiclo")
        drOperacion("Loteciclo") = data.drRuta("Loteciclo")
        drOperacion("Ciclo") = data.drRuta("Ciclo")

        Return drOperacion
    End Function

    <Serializable()> _
    Public Class DataLlenarCosteVar
        Public strArticuloPadre As String
        Public drVarios As DataRow
        Public intNivel As Integer
        Public intOrden As Integer
        Public dblCosteVarA As Double
        Public dblCosteDirecto As Double
        Public dblCosteIndirecto As Double
        Public dblCosteFijo As Double
        Public dblCosteVariable As Double
        Public FechaCalculo As Date

        Public Sub New()
        End Sub
        Public Sub New(ByVal strArticuloPadre As String, ByVal drVarios As DataRow, _
                       ByVal intNivel As Integer, ByVal intOrden As Integer, _
                       ByVal dblCosteVarA As Double, ByVal dblCosteDirecto As Double, _
                       ByVal dblCosteIndirecto As Double, ByVal dblCosteFijo As Double, _
                       ByVal dblCosteVariable As Double, _
                       ByVal FechaCalculo As Date)
            Me.strArticuloPadre = strArticuloPadre
            Me.drVarios = drVarios
            Me.intNivel = intNivel
            Me.intOrden = intOrden
            Me.dblCosteVarA = dblCosteVarA
            Me.dblCosteDirecto = dblCosteDirecto
            Me.dblCosteIndirecto = dblCosteFijo
            Me.dblCosteVariable = dblCosteVariable
            Me.FechaCalculo = FechaCalculo
        End Sub
    End Class

    <Task()> Public Shared Function LlenarCosteVarios(ByVal data As DataLlenarCosteVar, ByVal services As ServiceProvider) As DataRow
        Dim drCosteVarios As DataRow = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf HistoricoVariosAddNew, Nothing, services).NewRow

        drCosteVarios("IDCosteVarios") = AdminData.GetAutoNumeric
        drCosteVarios("IDArticuloPadre") = data.strArticuloPadre
        drCosteVarios("IDArticulo") = data.drVarios("IDArticulo")
        drCosteVarios("FechaCalculo") = data.FechaCalculo
        drCosteVarios("IDVarios") = data.drVarios("IDVarios")
        drCosteVarios("DescVarios") = data.drVarios("DescVarios")
        drCosteVarios("Nivel") = data.intNivel
        drCosteVarios("Orden") = data.intOrden
        drCosteVarios("Valor") = data.drVarios("Valor")
        drCosteVarios("Tipo") = data.drVarios("Tipo")
        drCosteVarios("CosteVariosA") = data.dblCosteVarA
        drCosteVarios("CosteDirectoA") = data.dblCosteDirecto
        drCosteVarios("CosteIndirectoA") = data.dblCosteIndirecto
        drCosteVarios("CosteFijoA") = data.dblCosteFijo
        drCosteVarios("CosteVariableA") = data.dblCosteVariable
        Return drCosteVarios
    End Function

    <Serializable()> _
    Public Class DataHistFilter
        Public strSelect As String = Nothing
        Public Where As Filter = Nothing
        Public strOrderBy As String = Nothing

        Public Sub New(Optional ByVal strSelect As String = Nothing, Optional ByVal Where As Filter = Nothing, Optional ByVal strOrderBy As String = Nothing)
            Me.strSelect = strSelect
            Me.Where = Where
            Me.strOrderBy = strOrderBy
        End Sub
    End Class

    <Task()> Public Shared Function HistoricoMaterialFilter(ByVal data As DataHistFilter, ByVal services As ServiceProvider) As DataTable
        HistoricoMaterialFilter = New BE.DataEngine().Filter("vNegHistoricoCosteMaterial", data.Where, data.strSelect, data.strOrderBy)
    End Function

    <Task()> Public Shared Function HistoricoOperacionFilter(ByVal data As DataHistFilter, ByVal services As ServiceProvider) As DataTable
        HistoricoOperacionFilter = New BE.DataEngine().Filter("vNegHistoricoCosteOperacion", data.Where, data.strSelect, data.strOrderBy)
    End Function

    <Task()> Public Shared Function HistoricoVariosFilter(ByVal data As DataHistFilter, ByVal services As ServiceProvider) As DataTable
        HistoricoVariosFilter = New BE.DataEngine().Filter("vNegHistoricoCosteVarios", data.Where, data.strSelect, data.strOrderBy)
    End Function

    <Task()> Public Shared Sub AplicarDecimales(ByVal udtDT As udtCostesDT, ByVal services As ServiceProvider)
        Dim intDecImpA As Integer
        Dim dtMoneda As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf Moneda.ObtenerMonedaA, Nothing, services)
        If Not dtMoneda Is Nothing AndAlso dtMoneda.Rows.Count > 0 Then
            intDecImpA = dtMoneda.Rows(0)("NDecimalesPrec")
        End If

        Dim dt As DataTable = udtDT.dtCosteStd
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            dt.Rows(0)("CosteUltimoA") = xRound(dt.Rows(0)("CosteUltimoA"), intDecImpA)
            dt.Rows(0)("CosteMatUltA") = xRound(dt.Rows(0)("CosteMatUltA"), intDecImpA)
            dt.Rows(0)("CosteOpeUltA") = xRound(dt.Rows(0)("CosteOpeUltA"), intDecImpA)
            dt.Rows(0)("CosteExtUltA") = xRound(dt.Rows(0)("CosteExtUltA"), intDecImpA)
            dt.Rows(0)("CosteVarUltA") = xRound(dt.Rows(0)("CosteVarUltA"), intDecImpA)
            dt.Rows(0)("CosteAcuMatUltA") = xRound(dt.Rows(0)("CosteAcuMatUltA"), intDecImpA)
            dt.Rows(0)("CosteAcuOpeUltA") = xRound(dt.Rows(0)("CosteAcuOpeUltA"), intDecImpA)
            dt.Rows(0)("CosteAcuExtUltA") = xRound(dt.Rows(0)("CosteAcuExtUltA"), intDecImpA)
            dt.Rows(0)("CosteAcuVarUltA") = xRound(dt.Rows(0)("CosteAcuVarUltA"), intDecImpA)
        End If

        dt = udtDT.dtMaterial
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            dt.Rows(0)("CosteStdA") = xRound(dt.Rows(0)("CosteStdA"), intDecImpA)
            dt.Rows(0)("CosteMatStdA") = xRound(dt.Rows(0)("CosteMatStdA"), intDecImpA)
            dt.Rows(0)("CosteOpeStdA") = xRound(dt.Rows(0)("CosteOpeStdA"), intDecImpA)
            dt.Rows(0)("CosteExtStdA") = xRound(dt.Rows(0)("CosteExtStdA"), intDecImpA)
            dt.Rows(0)("CosteVarStdA") = xRound(dt.Rows(0)("CosteVarStdA"), intDecImpA)
            dt.Rows(0)("CosteAcuMatStdA") = xRound(dt.Rows(0)("CosteAcuMatStdA"), intDecImpA)
            dt.Rows(0)("CosteAcuOpeStdA") = xRound(dt.Rows(0)("CosteAcuOpeStdA"), intDecImpA)
            dt.Rows(0)("CosteAcuExtStdA") = xRound(dt.Rows(0)("CosteAcuExtStdA"), intDecImpA)
            dt.Rows(0)("CosteAcuVarStdA") = xRound(dt.Rows(0)("CosteAcuVarStdA"), intDecImpA)
        End If

        dt = udtDT.dtOperacion
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            dt.Rows(0)("CosteOperacionA") = xRound(dt.Rows(0)("CosteOperacionA"), intDecImpA)
        End If

        dt = udtDT.dtVarios
        If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
            dt.Rows(0)("CosteVariosA") = xRound(dt.Rows(0)("CosteVariosA"), intDecImpA)
        End If

    End Sub

#End Region

#Region " CosteUltimoACosteStd "

    'Public Sub CosteUltimoACosteStd(ByVal IDArticulo As String)
    '    Dim services As New ServiceProvider
    '    If Len(IDArticulo) > 0 Then
    '        Dim f As New Filter
    '        f.Add(New StringFilterItem("IDArticulo", IDArticulo))
    '        Dim DtEstruc As DataTable = New BE.DataEngine().Filter("FrmMntoArticulos", f, "IDRuta, IDEstructura")
    '        Dim dtArticulo As DataTable = New Articulo().Filter(f, , "IDArticulo,PrecioEstandarA,PrecioEstandarB,FechaEstandar,IDTipoEstructura,IDTipoRuta")

    '        If Not dtArticulo Is Nothing AndAlso dtArticulo.Rows.Count > 0 Then
    '            Dim dt As DataTable = SelOnPrimaryKey(IDArticulo)
    '            If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then
    '                With dt

    '                    .Rows(0)("FechaEstandar") = .Rows(0)("FechaUltimo")
    '                    .Rows(0)("CosteStdA") = .Rows(0)("CosteUltimoA")

    '                    Dim dtMoneda As DataTable = ProcessServer.ExecuteTask(Of Object, DataTable)(AddressOf Moneda.ObtenerMonedaA, Nothing, services)
    '                    If Not dtMoneda Is Nothing AndAlso dtMoneda.Rows.Count > 0 Then
    '                        .Rows(0)("CosteStdB") = xRound(.Rows(0)("CosteStdA") * dtMoneda.Rows(0)("CambioB"), dtMoneda.Rows(0)("NDecimalesPrec"))
    '                    End If
    '                    .Rows(0)("CosteMatStdA") = .Rows(0)("CosteMatUltA")
    '                    .Rows(0)("CosteAcuMatStdA") = .Rows(0)("CosteAcuMatUltA")
    '                    .Rows(0)("CosteOpeStdA") = .Rows(0)("CosteOpeUltA")
    '                    .Rows(0)("CosteAcuOpeStdA") = .Rows(0)("CosteAcuOpeUltA")
    '                    .Rows(0)("CosteExtStdA") = .Rows(0)("CosteExtUltA")
    '                    .Rows(0)("CosteAcuExtStdA") = .Rows(0)("CosteAcuExtUltA")
    '                    .Rows(0)("CosteVarStdA") = .Rows(0)("CosteVarUltA")
    '                    .Rows(0)("CosteAcuVarStdA") = .Rows(0)("CosteAcuVarUltA")
    '                    .Rows(0)("CosteDirectoStdA") = .Rows(0)("CosteDirectoUltA")
    '                    .Rows(0)("CosteIndirectoStdA") = .Rows(0)("CosteIndirectoUltA")
    '                    .Rows(0)("CosteFijoStdA") = .Rows(0)("CosteFijoUltA")
    '                    .Rows(0)("CosteVariableStdA") = .Rows(0)("CosteVariableUltA")
    '                    .Rows(0)("IDEstructura") = DtEstruc.Rows(0)("IDEstructura") & String.Empty
    '                    .Rows(0)("IDTipoEstructura") = dtArticulo.Rows(0)("IDTipoEstructura") & String.Empty
    '                    .Rows(0)("IDRuta") = DtEstruc.Rows(0)("IDRuta") & String.Empty
    '                    .Rows(0)("IDTipoRuta") = dtArticulo.Rows(0)("IDTipoRuta") & String.Empty
    '                End With

    '                With dtArticulo
    '                    .Rows(0)("PrecioEstandarA") = dt.Rows(0)("CosteStdA")
    '                    .Rows(0)("PrecioEstandarB") = dt.Rows(0)("CosteStdB")
    '                    .Rows(0)("FechaEstandar") = CDate(Date.Today.ToShortDateString & " " & Now.ToShortTimeString)
    '                End With

    '                BusinessHelper.UpdateTable(dtArticulo)
    '                BusinessHelper.UpdateTable(dt)
    '            End If
    '        Else
    '            ApplicationService.GenerateError("El artículo no existe.")
    '        End If
    '    Else
    '        ApplicationService.GenerateError("El Artículo es un dato obligatorio.")
    '    End If
    'End Sub

    'Protected Overridable Function ValidarFechaPeriodo(ByVal dtMarcados As DataTable) As Boolean
    '    For Each row As DataRow In dtMarcados.Rows
    '        Dim dtDatosPeriodo As DataTable = New CierreInventario().SelOnPrimaryKey(Year(row("FechaUltimo")), Month(row("FechaUltimo")))
    '        If Not dtDatosPeriodo Is Nothing AndAlso dtDatosPeriodo.Rows.Count > 0 Then
    '            If dtDatosPeriodo.Rows(0)("Cerrado") Then
    '                ExpertisApp.GenerateMessage("La Fecha del último cálculo para el Artículo '|' pertenece a un periodo cerrado.", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Information, row("IDArticulo"))
    '                Return False
    '            End If
    '        End If
    '    Next
    '    Return True
    'End Function

    <Serializable()> _
    Public Class DataCosteUltimoACosteStdLog
        Public ArticulosActualizados As List(Of String)
        Public Errores As List(Of ClassErrors)
    End Class

    <Serializable()> _
    Public Class ProcInfoActualizarPrecioEstandar
        Public PermitirMovtoCantidad0 As Boolean
        Public RecalcularPrecioStdPosteriores As Boolean
    End Class

    <Serializable()> _
    Public Class DataCosteUltimoACosteStdMasivo
        Public dtArticulos As DataTable
        Public Fecha As Date
        Public Esquema As String

        Public Sub New(ByVal dtArticulos As DataTable, Optional ByVal Fecha As Date = cnMinDate, Optional ByVal Esquema As String = "")
            Me.dtArticulos = dtArticulos
            Me.Fecha = Fecha
            If Length(Esquema) > 0 Then Me.Esquema = Esquema
        End Sub
    End Class
    <Task()> Public Shared Function CosteUltimoACosteStdMasivo(ByVal data As DataCosteUltimoACosteStdMasivo, ByVal services As ServiceProvider) As DataCosteUltimoACosteStdLog
        If Not data Is Nothing AndAlso Not data.dtArticulos Is Nothing AndAlso data.dtArticulos.Rows.Count > 0 Then
            If Length(data.Esquema) = 0 Then
                data.Esquema = ProcessServer.ExecuteTask(Of Object, String)(AddressOf Business.General.Comunes.GetEsquemaBD, Nothing, services)
            End If

            For Each dr As DataRow In data.dtArticulos.Rows
                Dim stData As New DataCosteUltimoACosteStd(dr("IDArticulo"), data.Esquema)
                ProcessServer.ExecuteTask(Of DataCosteUltimoACosteStd)(AddressOf CosteUltimoACosteStd, stData, services)
            Next

            Return services.GetService(Of DataCosteUltimoACosteStdLog)()
        End If
    End Function

    <Serializable()> _
    Public Class DataCosteUltimoACosteStd
        Public IDArticulo As String
        Public Esquema As String

        Public Sub New(ByVal IDArticulo As String, Optional ByVal Esquema As String = "")
            Me.IDArticulo = IDArticulo
            If Length(Esquema) > 0 Then Me.Esquema = Esquema
        End Sub
    End Class

    <Task()> Public Shared Function CosteUltimoACosteStd(ByVal data As DataCosteUltimoACosteStd, ByVal services As ServiceProvider) As DataCosteUltimoACosteStdLog
        If Length(data.IDArticulo) > 0 Then
            Dim rslt As DataCosteUltimoACosteStdLog = services.GetService(Of DataCosteUltimoACosteStdLog)()
            Dim AppParams As ParametroStocks = services.GetService(Of ParametroStocks)()
            If Length(data.Esquema) = 0 Then
                data.Esquema = ProcessServer.ExecuteTask(Of Object, String)(AddressOf Business.General.Comunes.GetEsquemaBD, Nothing, services)
            End If

            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
            Dim dtEstruc As DataTable = New BE.DataEngine().Filter("FrmMntoArticulos", f, "IDRuta, IDEstructura")
            Dim dtArticulo As DataTable = New Articulo().Filter(f, , "IDArticulo,PrecioEstandarA,PrecioEstandarB,FechaEstandar,IDTipoEstructura,IDTipoRuta")
            If Not dtArticulo Is Nothing AndAlso dtArticulo.Rows.Count > 0 Then
                Dim dtArticuloCosteStd As DataTable = New ArticuloCosteEstandar().SelOnPrimaryKey(data.IDArticulo)
                Dim Monedas As MonedaCache = services.GetService(Of MonedaCache)()
                Dim MonInfoA As MonedaInfo = Monedas.MonedaA
                Dim MonInfoB As MonedaInfo = Monedas.MonedaB

                If Not dtArticuloCosteStd Is Nothing AndAlso dtArticuloCosteStd.Rows.Count > 0 Then
                    Dim FechaEstandar As Date = Nz(dtArticuloCosteStd.Rows(0)("FechaUltimo"), Today)

                    dtArticuloCosteStd.Rows(0)("FechaEstandar") = FechaEstandar
                    dtArticuloCosteStd.Rows(0)("CosteStdA") = dtArticuloCosteStd.Rows(0)("CosteUltimoA")
                    dtArticuloCosteStd.Rows(0)("CosteStdB") = xRound(dtArticuloCosteStd.Rows(0)("CosteStdA") * MonInfoA.CambioB, MonInfoB.NDecimalesPrecio)
                    dtArticuloCosteStd.Rows(0)("CosteMatStdA") = dtArticuloCosteStd.Rows(0)("CosteMatUltA")
                    dtArticuloCosteStd.Rows(0)("CosteAcuMatStdA") = dtArticuloCosteStd.Rows(0)("CosteAcuMatUltA")
                    dtArticuloCosteStd.Rows(0)("CosteOpeStdA") = dtArticuloCosteStd.Rows(0)("CosteOpeUltA")
                    dtArticuloCosteStd.Rows(0)("CosteAcuOpeStdA") = dtArticuloCosteStd.Rows(0)("CosteAcuOpeUltA")
                    dtArticuloCosteStd.Rows(0)("CosteExtStdA") = dtArticuloCosteStd.Rows(0)("CosteExtUltA")
                    dtArticuloCosteStd.Rows(0)("CosteAcuExtStdA") = dtArticuloCosteStd.Rows(0)("CosteAcuExtUltA")
                    dtArticuloCosteStd.Rows(0)("CosteVarStdA") = dtArticuloCosteStd.Rows(0)("CosteVarUltA")
                    dtArticuloCosteStd.Rows(0)("CosteAcuVarStdA") = dtArticuloCosteStd.Rows(0)("CosteAcuVarUltA")
                    dtArticuloCosteStd.Rows(0)("CosteDirectoStdA") = dtArticuloCosteStd.Rows(0)("CosteDirectoUltA")
                    dtArticuloCosteStd.Rows(0)("CosteIndirectoStdA") = dtArticuloCosteStd.Rows(0)("CosteIndirectoUltA")
                    dtArticuloCosteStd.Rows(0)("CosteFijoStdA") = dtArticuloCosteStd.Rows(0)("CosteFijoUltA")
                    dtArticuloCosteStd.Rows(0)("CosteVariableStdA") = dtArticuloCosteStd.Rows(0)("CosteVariableUltA")

                    dtArticuloCosteStd.Rows(0)("IDEstructura") = dtEstruc.Rows(0)("IDEstructura") & String.Empty
                    dtArticuloCosteStd.Rows(0)("IDTipoEstructura") = dtArticulo.Rows(0)("IDTipoEstructura") & String.Empty
                    dtArticuloCosteStd.Rows(0)("IDRuta") = dtEstruc.Rows(0)("IDRuta") & String.Empty
                    dtArticuloCosteStd.Rows(0)("IDTipoRuta") = dtArticulo.Rows(0)("IDTipoRuta") & String.Empty


                    dtArticulo.Rows(0)("PrecioEstandarA") = dtArticuloCosteStd.Rows(0)("CosteStdA")
                    dtArticulo.Rows(0)("PrecioEstandarB") = dtArticuloCosteStd.Rows(0)("CosteStdB")
                    dtArticulo.Rows(0)("FechaEstandar") = dtArticuloCosteStd.Rows(0)("FechaEstandar")

                    Try

                        AdminData.BeginTx()
                        BusinessHelper.UpdateTable(dtArticulo)
                        BusinessHelper.UpdateTable(dtArticuloCosteStd)

                        If AppParams.TipoMovimientoCantidad0 <> 0 AndAlso AppParams.TipoMovimientoCantidad0 > enumTipoMovimiento.tmSalContraActivos Then
                            Dim ProcInfo As ProcInfoActualizarPrecioEstandar = services.GetService(Of ProcInfoActualizarPrecioEstandar)()
                            ProcInfo.PermitirMovtoCantidad0 = True
                            ProcInfo.RecalcularPrecioStdPosteriores = True
                            Dim stData As New DataGenerarMovimiento(data.Esquema, dtArticulo.Rows(0)("IDArticulo"), FechaEstandar, dtArticulo.Rows(0)("PrecioEstandarA"), dtArticulo.Rows(0)("PrecioEstandarB"))
                            ProcessServer.ExecuteTask(Of DataGenerarMovimiento)(AddressOf GenerarMovimiento, stData, services)
                        End If

                        AdminData.CommitTx(True)

                        If rslt.ArticulosActualizados Is Nothing Then rslt.ArticulosActualizados = New List(Of String)
                        rslt.ArticulosActualizados.Add(data.IDArticulo)
                    Catch ex As Exception
                        AdminData.RollBackTx()

                        If rslt.Errores Is Nothing Then rslt.Errores = New List(Of ClassErrors)
                        Dim err As New ClassErrors(data.IDArticulo, ex.Message)
                        rslt.Errores.Add(err)
                    End Try
                End If
                Return rslt
            Else
                ApplicationService.GenerateError("El artículo no existe.")
            End If
        Else
            ApplicationService.GenerateError("El Artículo es un dato obligatorio.")
        End If
    End Function


    <Serializable()> _
    Public Class DataGenerarMovimiento
        Public Esquema As String
        Public IDArticulo As String
        Public Fecha As Date
        Public PrecioA As Double
        Public PrecioB As Double

        Public Sub New(ByVal Esquema As String, ByVal IDArticulo As String, ByVal Fecha As Date, ByVal PrecioA As Double, ByVal PrecioB As Double)
            Me.Esquema = Esquema
            Me.IDArticulo = IDArticulo
            Me.Fecha = Fecha
            Me.PrecioA = PrecioA
            Me.PrecioB = PrecioB
        End Sub
    End Class
    <Task()> Public Shared Sub GenerarMovimiento(ByVal data As DataGenerarMovimiento, ByVal services As ServiceProvider)
        Dim AppParams As ParametroStocks = services.GetService(Of ParametroStocks)()

        'Dim vStr As String = "SELECT * FROM tbHistoricoMovimiento " + _
        '                     " WHERE tbHistoricoMovimiento.IDLineaMovimiento IN " + _
        '                     " (SELECT  " & data.Esquema & ".fMovimientoValArticulo('" & Format(data.Fecha, "yyyyMMdd") & "', IDArticulo, IDAlmacen) AS MIDLineaMovimiento " + _
        '                     " FROM tbMaestroArticuloAlmacen WHERE " & data.Esquema & ".fMovimientoValArticulo('" & Format(data.Fecha, "yyyyMMdd") & "', IDArticulo, IDAlmacen)<>0) " + _
        '                     " AND tbHistoricoMovimiento.IDArticulo = " + Quoted(data.IDArticulo)

        Dim vStr As String = "SELECT hm.* FROM tbHistoricoMovimiento hm " + _
             " INNER JOIN (SELECT coalesce(" & data.Esquema & ".fMovimientoValArticulo('" & Format(data.Fecha, "yyyyMMdd") & "', IDArticulo, IDAlmacen), 0) AS MIDLineaMovimiento, IDArticulo, IDAlmacen " + _
             " FROM tbMaestroArticuloAlmacen) v ON hm.IDArticulo = v.IDArticulo AND hm.IDAlmacen = v.IDAlmacen AND hm.IDLineaMovimiento = v.MIDLineaMovimiento " + _
             " WHERE hm.IDArticulo = " + Quoted(data.IDArticulo)

        Dim IDLineaMov(-1) As Object
        Dim dtMov As DataTable = AdminData.Execute(vStr, ExecuteCommand.ExecuteReader)
        If Not dtMov Is Nothing AndAlso dtMov.Rows.Count > 0 Then

            Dim fArtAlmOR As New Filter(FilterUnionOperator.Or)
            Dim MovimientosGenerar(-1) As StockData
            Dim dtHistMov As DataTable = dtMov.Clone
            Dim Enlace As Integer
            Dim Articulos As EntityInfoCache(Of ArticuloInfo) = services.GetService(Of EntityInfoCache(Of ArticuloInfo))()
            For Each movimiento As DataRow In dtMov.Rows
                Enlace += 1
                Dim stk As New StockData
                stk.Enlace = Enlace
                stk.Articulo = movimiento("IDArticulo")
                stk.Almacen = movimiento("IDAlmacen")

                Dim fArtAlm As New Filter
                fArtAlm.Add(New StringFilterItem("IDArticulo", stk.Articulo))
                fArtAlm.Add(New StringFilterItem("IDAlmacen", stk.Almacen))
                fArtAlmOR.Add(fArtAlm)
                Dim ArtInfo As ArticuloInfo = Articulos.GetEntity(stk.Articulo)
                If ArtInfo.NSerieObligatorio Then
                    stk.NSerie = movimiento("Lote") & String.Empty
                    stk.EstadoNSerie = movimiento("IDEstadoActivo") & String.Empty
                    stk.Ubicacion = movimiento("Ubicacion") & String.Empty
                ElseIf ArtInfo.GestionStockPorLotes Then
                    stk.Lote = movimiento("Lote") & String.Empty
                    stk.Ubicacion = movimiento("Ubicacion") & String.Empty
                End If
                stk.Operario = movimiento("IDOperario") & String.Empty
                'stk.Obra = Nz(row.Cells("IDObra").Value)
                'stk.FechaCaducidad = Nz(row.Cells("FechaCaducidad").Value)
                stk.Cantidad = 0
                If ProcessServer.ExecuteTask(Of String, Boolean)(AddressOf ProcesoComunes.AplicarSegundaUnidad, stk.Articulo, services) Then
                    stk.Cantidad2 = 0
                End If
                stk.PrecioA = data.PrecioA
                stk.PrecioB = data.PrecioB
                stk.Documento = Nothing
                stk.Texto = AdminData.GetMessageText("Movimiento Actualización Coste Std")
                stk.FechaDocumento = data.Fecha
                stk.TipoMovimiento = AppParams.TipoMovimientoCantidad0

                ReDim Preserve MovimientosGenerar(MovimientosGenerar.Length)
                MovimientosGenerar(MovimientosGenerar.Length - 1) = stk
            Next

            Dim fMovtoCeroEnFecha As New Filter
            fMovtoCeroEnFecha.Add(New NumberFilterItem("IDTipoMovimiento", AppParams.TipoMovimientoCantidad0))
            fMovtoCeroEnFecha.Add(New DateFilterItem("FechaDocumento", data.Fecha))
            fMovtoCeroEnFecha.Add(fArtAlmOR)

            If MovimientosGenerar.Length > 0 Then
                '//REVISAR: ver otra forma de hacer esto?
                If Not fMovtoCeroEnFecha Is Nothing AndAlso fMovtoCeroEnFecha.Count >= 3 Then
                    Dim FilterDel As String = AdminData.ComposeFilter(fMovtoCeroEnFecha)
                    If Length(FilterDel) > 0 Then
                        FilterDel = " WHERE " & FilterDel
                        Dim sql As String = "DELETE FROM tbHistoricoMovimiento " & FilterDel
                        AdminData.Execute(sql)
                    End If
                End If

                Dim NumeroMovimiento As Integer = ProcessServer.ExecuteTask(Of Object, Integer)(AddressOf ProcesoStocks.NuevoNumeroMovimiento, Nothing, services)
                Dim datMovtos As New ProcesoStocks.DataMovimientosGenericosES(NumeroMovimiento, MovimientosGenerar)
                Dim updateData As StockUpdateData() = ProcessServer.ExecuteTask(Of ProcesoStocks.DataMovimientosGenericosES, StockUpdateData())(AddressOf ProcesoStocks.MovimientosGenericosES, datMovtos, services)
                If Not updateData Is Nothing AndAlso updateData.Count > 0 Then
                    Dim StockNoActualizado As List(Of StockUpdateData) = (From c In updateData.ToList Where c.Estado = EstadoStock.NoActualizado Select c).ToList
                    If Not StockNoActualizado Is Nothing AndAlso StockNoActualizado.Count > 0 Then
                        ApplicationService.GenerateError(StockNoActualizado(0).Detalle)
                    End If
                End If
            End If
        End If
    End Sub


#End Region



End Class