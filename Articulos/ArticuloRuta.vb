Public Class ArticuloRuta

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbArticuloRuta"

#End Region

#Region "Eventos RegisterValidateTasks"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarClaveDuplicada)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("IDRuta")) = 0 Then ApplicationService.GenerateError("El Identificador de Ruta es un dato obligatorio.")
        If Length(data("DescRuta")) = 0 Then ApplicationService.GenerateError("La descripción de la ruta es un dato obligatorio.")
        If Length(data("IDArticulo")) = 0 Then ApplicationService.GenerateError("El Artículo es obligatorio.")
    End Sub

    <Task()> Public Shared Sub ValidarClaveDuplicada(ByVal data As DataRow, ByVal services As ServiceProvider)
        If data.RowState = DataRowState.Added Then
            Dim FilAP As New Filter
            FilAP.Add("IDRuta", FilterOperator.Equal, data("IDRuta"), FilterType.String)
            FilAP.Add("IDArticulo", FilterOperator.Equal, data("IDArticulo"), FilterType.String)
            Dim DtAP As DataTable = New ArticuloRuta().Filter(FilAP)
            If Not DtAP Is Nothing AndAlso DtAP.Rows.Count > 0 Then
                ApplicationService.GenerateError("La ruta ya existe en la lista actual.", data("IDArticulo"), data("IDRuta"))
            End If
        End If
    End Sub

#End Region

#Region "Eventos RegisterUpdateTasks"

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataRow)(AddressOf TratarPrincipal)
    End Sub

    <Task()> Public Shared Sub TratarPrincipal(ByVal data As DataRow, ByVal services As ServiceProvider)
        Dim FilPrin As New Filter
        FilPrin.Add("IDArticulo", FilterOperator.Equal, data("IDArticulo"), FilterType.String)
        FilPrin.Add("Principal", FilterOperator.Equal, 1, FilterType.Boolean)
        Dim DtPrincipal As DataTable = New ArticuloRuta().Filter(FilPrin)
        If IsNothing(DtPrincipal) OrElse DtPrincipal.Rows.Count = 0 Then
            data("Principal") = True
        Else
            If Nz(data("Principal"), False) Then
                If data("IDRuta") <> DtPrincipal.Rows(0)("IDRuta") Then
                    DtPrincipal.Rows(0)("Principal") = False
                    BusinessHelper.UpdateTable(DtPrincipal)
                End If
            ElseIf data.RowState = DataRowState.Modified AndAlso data("Principal") <> data("Principal", DataRowVersion.Original) AndAlso DtPrincipal.Rows.Count = 1 Then
                data("Principal") = True
            End If
        End If
    End Sub

#End Region

#Region "Funciones Públicas"

    <Task()> Public Shared Function RutaPpal(ByVal data As String, ByVal services As ServiceProvider) As String
        Dim FilRuta As New Filter
        FilRuta.Add("IDArticulo", FilterOperator.Equal, data, FilterType.String)
        FilRuta.Add(New BooleanFilterItem("Principal", FilterOperator.Equal, True))
        Dim DtRuta As DataTable = New ArticuloRuta().Filter(FilRuta)
        If Not DtRuta Is Nothing AndAlso DtRuta.Rows.Count > 0 Then
            Return DtRuta.Rows(0)("IDRuta")
        End If
    End Function

    <Serializable()> _
    Public Class DatosEstRutaPrin
        Public IDArticulo As String
        Public IDRuta As String
    End Class

    <Task()> Public Shared Sub EstablecerRutaPrincipal(ByVal data As DatosEstRutaPrin, ByVal services As ServiceProvider)
        If Length(data.IDArticulo) > 0 AndAlso Length(data.IDRuta) > 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
            f.Add(New StringFilterItem("IDRuta", data.IDRuta))
            Dim ClsArtRuta As New ArticuloRuta
            Dim dtArticuloRuta As DataTable = ClsArtRuta.Filter(f)
            If Not dtArticuloRuta Is Nothing AndAlso dtArticuloRuta.Rows.Count > 0 Then
                'Quitar la estructura principal actual
                f.Clear()
                f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
                f.Add(New BooleanFilterItem("Principal", True))
                Dim dtRutaPpal As DataTable = ClsArtRuta.Filter(f)
                If Not dtRutaPpal Is Nothing AndAlso dtRutaPpal.Rows.Count > 0 Then
                    dtRutaPpal.Rows(0)("Principal") = False
                    ClsArtRuta.Update(dtRutaPpal)
                End If
                'Establecer la ruta principal
                dtArticuloRuta.Rows(0)("Principal") = True
                ClsArtRuta.Update(dtArticuloRuta)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub EliminarRuta(ByVal data As DatosEstRutaPrin, ByVal services As ServiceProvider)
        If Length(data.IDArticulo) > 0 AndAlso Length(data.IDRuta) > 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", data.IDArticulo))
            f.Add(New StringFilterItem("IDRuta", data.IDRuta))
            Dim ClsArtRuta As New ArticuloRuta
            Dim dtRuta As DataTable = ClsArtRuta.Filter(f)
            If Not dtRuta Is Nothing AndAlso dtRuta.Rows.Count > 0 Then
                ClsArtRuta.Delete(dtRuta)
            End If
        End If
    End Sub

    <Serializable()> _
    Public Class DatosCopiaRuta
        Public IDArticuloOrigen As String
        Public IDRutaOrigen As String
        Public IDArticuloDestino As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal IDArticuloOrigen As String, ByVal IDRutaOrigen As String, ByVal IDArticuloDestino As String)
            Me.IDArticuloOrigen = IDArticuloOrigen
            Me.IDRutaOrigen = IDRutaOrigen
            Me.IDArticuloDestino = IDArticuloDestino
        End Sub
    End Class
    <Task()> Public Shared Sub CopiarRuta(ByVal data As DatosCopiaRuta, ByVal services As ServiceProvider)
        If Length(data.IDRutaOrigen) > 0 And Length(data.IDArticuloDestino) > 0 And Length(data.IDArticuloOrigen) > 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", data.IDArticuloDestino))
            f.Add(New StringFilterItem("IDRuta", data.IDRutaOrigen))
            Dim ClsArtRuta As New ArticuloRuta
            Dim dtArticuloRutaNew As DataTable = ClsArtRuta.Filter(f)
            If dtArticuloRutaNew Is Nothing OrElse dtArticuloRutaNew.Rows.Count = 0 Then
                f.Clear()
                f.Add(New StringFilterItem("IDArticulo", data.IDArticuloOrigen))
                f.Add(New StringFilterItem("IDRuta", data.IDRutaOrigen))
                Dim dtArticuloRutaOrigen As DataTable = ClsArtRuta.Filter(f)
                If Not dtArticuloRutaOrigen Is Nothing AndAlso dtArticuloRutaOrigen.Rows.Count > 0 Then
                    dtArticuloRutaNew = ClsArtRuta.AddNewForm
                    dtArticuloRutaNew.Rows(0)("IDArticulo") = data.IDArticuloDestino
                    dtArticuloRutaNew.Rows(0)("IDRuta") = data.IDRutaOrigen
                    dtArticuloRutaNew.Rows(0)("DescRuta") = dtArticuloRutaOrigen.Rows(0)("DescRuta")
                    dtArticuloRutaNew.Rows(0)("Principal") = False
                    dtArticuloRutaNew.Rows(0)("FechaVigencia") = dtArticuloRutaOrigen.Rows(0)("FechaVigencia")
                    ClsArtRuta.Update(dtArticuloRutaNew)
                    Dim r As New Ruta
                    Dim dtRutaOrigen As DataTable = r.Filter(f)
                    If Not dtRutaOrigen Is Nothing AndAlso dtRutaOrigen.Rows.Count > 0 Then

                        For Each drRutaOrigen As DataRow In dtRutaOrigen.Rows
                            Dim dtRutaNew As DataTable = r.AddNew
                            Dim drRutaNew As DataRow = dtRutaNew.NewRow
                            drRutaNew.ItemArray = drRutaOrigen.ItemArray
                            drRutaNew("IDArticulo") = data.IDArticuloDestino
                            dtRutaNew.Rows.Add(drRutaNew)
                            r.Update(dtRutaNew)
                            Dim StDatos As New DatosCopiaRutaDr
                            Dim dtorigen As DataTable = drRutaOrigen.Table.Clone
                            Dim drorigen As DataRow = dtorigen.NewRow
                            drorigen.ItemArray = drRutaOrigen.ItemArray
                            dtorigen.Rows.Add(drorigen)


                            StDatos.RutaOrigen = dtorigen
                            StDatos.RutaDestino = dtRutaNew
                            ProcessServer.ExecuteTask(Of DatosCopiaRutaDr)(AddressOf CopiarUtillajes, StDatos, services)
                            ProcessServer.ExecuteTask(Of DatosCopiaRutaDr)(AddressOf CopiarOficios, StDatos, services)
                            ProcessServer.ExecuteTask(Of DatosCopiaRutaDr)(AddressOf CopiarParametros, StDatos, services)
                            ProcessServer.ExecuteTask(Of DatosCopiaRutaDr)(AddressOf CopiarProveedor, StDatos, services)
                            ProcessServer.ExecuteTask(Of DatosCopiaRutaDr)(AddressOf CopiarCentrosAlternativos, StDatos, services)
                        Next
                    End If
                End If
            End If
        End If
    End Sub

    <Serializable()> _
    Public Class DatosCopiaRutaOperacion
        Public IDArticuloOrigen As String
        Public IDRutaOrigen As String
        Public IDArticuloDestino As String
        Public IDOperacionOrigen As String
        Public SecuenciaOrigen As Integer
        Public IDRutaDestino As String

        Public Sub New(ByVal IDArticuloOrigen As String, ByVal IDRutaOrigen As String, ByVal IDOperacionOrigen As String, ByVal SecuenciaOrigen As Integer, ByVal IDArticuloDestino As String, ByVal IDRutaDestino As String)
            Me.IDArticuloOrigen = IDArticuloOrigen
            Me.IDRutaOrigen = IDRutaOrigen
            Me.IDOperacionOrigen = IDOperacionOrigen
            Me.IDArticuloDestino = IDArticuloDestino
            Me.SecuenciaOrigen = SecuenciaOrigen
            Me.IDRutaDestino = IDRutaDestino
        End Sub
    End Class
    <Task()> Public Shared Sub CopiarRutaOperacion(ByVal data As DatosCopiaRutaOperacion, ByVal services As ServiceProvider)
        If Length(data.IDArticuloOrigen) > 0 And Length(data.IDRutaOrigen) > 0 And Length(data.IDOperacionOrigen) > 0 And Length(data.IDArticuloDestino) > 0 Then
            Dim f As New Filter
            f.Add(New StringFilterItem("IDArticulo", data.IDArticuloOrigen))
            f.Add(New StringFilterItem("IDRuta", data.IDRutaOrigen))
            f.Add(New StringFilterItem("IDOperacion", data.IDOperacionOrigen))
            f.Add(New NumberFilterItem("Secuencia", data.SecuenciaOrigen))
            Dim r As New Ruta
            Dim dtRutaOrigen As DataTable = r.Filter(f)
            If Not dtRutaOrigen Is Nothing AndAlso dtRutaOrigen.Rows.Count > 0 Then
                Dim dtRutaNew As DataTable = r.AddNew
                For Each drRutaOrigen As DataRow In dtRutaOrigen.Rows
                    Dim drRutaNew As DataRow = dtRutaNew.NewRow
                    drRutaNew.ItemArray = drRutaOrigen.ItemArray
                    drRutaNew("IDArticulo") = data.IDArticuloDestino
                    drRutaNew("IDRuta") = data.IDRutaDestino

                    dtRutaNew.Rows.Add(drRutaNew)
                    r.Update(dtRutaNew)

                    Dim StDatos As New DatosCopiaRutaDr
                    StDatos.RutaOrigen = drRutaOrigen.Table
                    StDatos.RutaDestino = drRutaNew.Table
                    ProcessServer.ExecuteTask(Of DatosCopiaRutaDr)(AddressOf CopiarUtillajes, StDatos, services)
                    ProcessServer.ExecuteTask(Of DatosCopiaRutaDr)(AddressOf CopiarOficios, StDatos, services)
                    ProcessServer.ExecuteTask(Of DatosCopiaRutaDr)(AddressOf CopiarParametros, StDatos, services)
                    ProcessServer.ExecuteTask(Of DatosCopiaRutaDr)(AddressOf CopiarProveedor, StDatos, services)
                    ProcessServer.ExecuteTask(Of DatosCopiaRutaDr)(AddressOf CopiarCentrosAlternativos, StDatos, services)
                Next
            End If
        End If
    End Sub

    <Serializable()> _
    Public Class DatosCopiaRutaDr
        Public RutaOrigen As DataTable
        Public RutaDestino As DataTable
    End Class

    <Task()> Public Shared Sub CopiarUtillajes(ByVal data As DatosCopiaRutaDr, ByVal services As ServiceProvider)
        If Not data.RutaOrigen Is Nothing Then
            Dim u As New RutaUtillaje
            Dim dtNew As DataTable = u.AddNew
            Dim f As New Filter
            f.Add(New NumberFilterItem("IDRutaOP", data.RutaOrigen.Rows(0)("IDRutaOp")))
            Dim dtU As DataTable = u.Filter(f)
            If Not dtU Is Nothing AndAlso dtU.Rows.Count > 0 Then
                For Each drU As DataRow In dtU.Rows
                    Dim drNew As DataRow = dtNew.NewRow
                    drNew("IDRutaOP") = data.RutaDestino.Rows(0)("IDRutaOP")
                    drNew("IDUtillaje") = drU("IDUtillaje")
                    drNew("Texto") = drU("Texto")
                    drNew("Critico") = drU("Critico")

                    dtNew.Rows.Add(drNew)
                Next
                u.Update(dtNew)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CopiarOficios(ByVal data As DatosCopiaRutaDr, ByVal services As ServiceProvider)
        If Not data.RutaOrigen Is Nothing Then
            Dim o As BusinessHelper = BusinessHelper.CreateBusinessObject("RutaOficio")
            Dim dtNew As DataTable = o.AddNew
            Dim f As New Filter
            f.Add(New NumberFilterItem("IDRutaOP", data.RutaOrigen.Rows(0)("IDRutaOp")))
            Dim dtO As DataTable = o.Filter(f)
            If Not dtO Is Nothing AndAlso dtO.Rows.Count > 0 Then
                For Each drO As DataRow In dtO.Rows
                    Dim drNew As DataRow = dtNew.NewRow
                    drNew("IDRutaOP") = data.RutaDestino.Rows(0)("IDRutaOP")
                    drNew("IdOficio") = drO("IdOficio")
                    drNew("Porcentaje") = drO("Porcentaje")

                    dtNew.Rows.Add(drNew)
                Next
                o.Update(dtNew)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CopiarParametros(ByVal data As DatosCopiaRutaDr, ByVal services As ServiceProvider)
        If Not data.RutaOrigen Is Nothing Then
            Dim p As New RutaParametro
            Dim dtNew As DataTable = p.AddNew
            Dim f As New Filter
            f.Add(New NumberFilterItem("IDRutaOP", data.RutaOrigen.Rows(0)("IDRutaOp")))
            Dim dtP As DataTable = p.Filter(f)
            If Not dtP Is Nothing AndAlso dtP.Rows.Count > 0 Then
                For Each drP As DataRow In dtP.Rows
                    Dim fExisteParam As New Filter
                    fExisteParam.Add(New NumberFilterItem("IDRutaOP", data.RutaDestino.Rows(0)("IDRutaOP")))
                    fExisteParam.Add(New StringFilterItem("IDParametro", drP("IDParametro")))
                    Dim dtParam As DataTable = p.Filter(fExisteParam)
                    If dtParam.Rows.Count > 0 Then
                        dtParam.Rows(0)("DescParametro") = drP("DescParametro")
                        dtParam.Rows(0)("Secuencia") = drP("Secuencia")
                        dtParam.Rows(0)("Valor") = drP("Valor")
                        dtNew.ImportRow(dtParam.Rows(0))
                    Else
                        Dim drNew As DataRow = dtNew.NewRow
                        drNew("ID") = AdminData.GetAutoNumeric
                        drNew("IDRutaOP") = data.RutaDestino.Rows(0)("IDRutaOP")
                        drNew("IDParametro") = drP("IDParametro")
                        drNew("DescParametro") = drP("DescParametro")
                        drNew("Secuencia") = drP("Secuencia")
                        drNew("Valor") = drP("Valor")
                        dtNew.Rows.Add(drNew)
                    End If
                Next
                p.Update(dtNew)
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CopiarProveedor(ByVal data As DatosCopiaRutaDr, ByVal services As ServiceProvider)
        If Not data.RutaOrigen Is Nothing Then
            Dim p As BusinessHelper
            p = BusinessHelper.CreateBusinessObject("RutaProveedor")
            Dim dtNew As DataTable = p.AddNew
            Dim f As New Filter
            f.Add(New NumberFilterItem("IDRutaOP", data.RutaOrigen.Rows(0)("IDRutaOp")))
            Dim dtP As DataTable = p.Filter(f)
            If Not dtP Is Nothing AndAlso dtP.Rows.Count > 0 Then
                For Each drP As DataRow In dtP.Rows
                    Dim drNew As DataRow = dtNew.NewRow
                    drNew("IDRutaOP") = data.RutaDestino.Rows(0)("IDRutaOP")
                    drNew("IDProveedor") = drP("IDProveedor")
                    drNew("IDCentro") = drP("IDCentro")
                    drNew("UDValoracion") = drP("UDValoracion")
                    drNew("IDUDProduccion") = drP("IDUDProduccion")
                    drNew("Principal") = drP("Principal")
                    drNew("PlazoSub") = drP("PlazoSub")
                    drNew("UDTiempoPlazo") = drP("UDTiempoPlazo")
                    dtNew.Rows.Add(drNew)
                Next
                p.Update(dtNew)

                Dim pl As BusinessHelper
                pl = BusinessHelper.CreateBusinessObject("RutaProveedorLinea")

                Dim dtPL As DataTable = pl.Filter(f)
                If Not dtPL Is Nothing AndAlso dtPL.Rows.Count > 0 Then
                    dtNew = pl.AddNew
                    For Each drPL As DataRow In dtPL.Rows
                        Dim drNew As DataRow = dtNew.NewRow
                        drNew("IDRutaOP") = data.RutaDestino.Rows(0)("IDRutaOP")
                        drNew("IDProveedor") = drPL("IDProveedor")
                        drNew("IDCentro") = drPL("IDCentro")
                        drNew("QDesde") = drPL("QDesde")
                        drNew("Precio") = drPL("Precio")
                        drNew("PrecioA") = drPL("PrecioA")
                        drNew("PrecioB") = drPL("PrecioB")
                        dtNew.Rows.Add(drNew)
                    Next
                    pl.Update(dtNew)
                End If
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CopiarCentrosAlternativos(ByVal data As DatosCopiaRutaDr, ByVal services As ServiceProvider)
        If Not data.RutaOrigen Is Nothing Then
            Dim p As New RutaAlternativo
            Dim dtNew As DataTable = p.AddNew
            Dim f As New Filter
            f.Add(New NumberFilterItem("IDRutaOP", data.RutaOrigen.Rows(0)("IDRutaOP")))
            Dim dtP As DataTable = p.Filter(f)
            If Not dtP Is Nothing AndAlso dtP.Rows.Count > 0 Then
                For Each drP As DataRow In dtP.Rows
                    Dim drNew As DataRow = dtNew.NewRow
                    drNew.ItemArray = drP.ItemArray
                    drNew("ID") = AdminData.GetAutoNumeric
                    drNew("IDRutaOP") = data.RutaDestino.Rows(0)("IDRutaOP")
                    dtNew.Rows.Add(drNew)
                    p.Update(dtNew)
                    Dim StDatos As New DatosCopiaRutaDr
                    StDatos.RutaOrigen = data.RutaOrigen
                    StDatos.RutaDestino = drNew.Table
                    ProcessServer.ExecuteTask(Of DatosCopiaRutaDr)(AddressOf CopiarUtillajesAlternativos, StDatos, services)
                Next
            End If
        End If
    End Sub

    <Task()> Public Shared Sub CopiarUtillajesAlternativos(ByVal data As DatosCopiaRutaDr, ByVal services As ServiceProvider)
        If Not data.RutaOrigen Is Nothing Then
            Dim p As New RutaUtillajeAlternativo
            Dim dtNew As DataTable = p.AddNew
            Dim f As New Filter
            f.Add(New NumberFilterItem("ID", data.RutaDestino.Rows(0)("ID")))
            Dim dtP As DataTable = p.Filter(f)
            If Not dtP Is Nothing AndAlso dtP.Rows.Count > 0 Then
                For Each drP As DataRow In dtP.Rows
                    Dim drNew As DataRow = dtNew.NewRow
                    drNew.ItemArray = drP.ItemArray
                    drNew("IDRutaOP") = data.RutaOrigen.Rows(0)("IDRutaOP")
                    dtNew.Rows.Add(drNew)
                Next
                p.Update(dtNew)
            End If
        End If
    End Sub

#End Region

End Class