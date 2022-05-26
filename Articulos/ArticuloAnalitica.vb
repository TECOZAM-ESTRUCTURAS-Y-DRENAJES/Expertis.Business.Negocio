Public Class ArticuloAnalitica
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbArticuloAnalitica"

#Region " RegisterValidateTask "

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarAnalitica)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarRegistroExistente)
    End Sub

    <Task()> Public Shared Sub ValidarAnalitica(ByVal dr As DataRow, ByVal services As ServiceProvider)
        Dim AppParams As ParametroAnalitica = services.GetService(Of ParametroAnalitica)()
        If Not AppParams.AplicarAnalitica Then Exit Sub

        If dr.RowState = DataRowState.Added Then
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.ValidarArticuloObligatorio, dr, services)
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf NegocioGeneral.ValidarCentroCosteObligatorio, dr, services)
            ProcessServer.ExecuteTask(Of DataRow)(AddressOf ValidarRegistroExistente, dr, services)
        End If
    End Sub

    <Task()> Public Shared Sub ValidarRegistroExistente(ByVal dr As DataRow, ByVal services As ServiceProvider)
        If dr.RowState = DataRowState.Added Then
            Dim AA As New ArticuloAnalitica
            Dim dt As DataTable = AA.SelOnPrimaryKey(dr("IDArticulo"), dr("IDCentroCoste"))
            If Not IsNothing(dt) AndAlso dt.Rows.Count > 0 Then
                ApplicationService.GenerateError("El Centro Coste {0} ya está asignado en este artículo.", Quoted(dr("IDCentroCoste")))
            End If
        End If
    End Sub

#End Region

#Region " RegisterUpdateTask "

    Protected Overrides Sub RegisterUpdateTasks(ByVal updateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterUpdateTasks(updateProcess)
        updateProcess.AddTask(Of DataTable)(AddressOf GrabarAnaliticaMultinivel)
    End Sub

    <Task()> Public Shared Sub GrabarAnaliticaMultinivel(ByVal dt As DataTable, ByVal services As ServiceProvider)
        ProcessServer.ExecuteTask(Of Object)(AddressOf Business.General.Comunes.BeginTransaction, Nothing, services)
        Dim AppParams As ParametroAnalitica = services.GetService(Of ParametroAnalitica)()
        If Not AppParams.AplicarAnalitica Then Exit Sub
        For Each dr As DataRow In dt.Rows
            If dr.RowState = DataRowState.Added OrElse dr.RowState = DataRowState.Modified Then
                Dim ClsCoste As BusinessHelper = BusinessHelper.CreateBusinessObject("CentroCosteAnalitica")
                Dim DtCoste As DataTable = ClsCoste.SelOnPrimaryKey(dr("IDCentroCoste"))
                If DtCoste Is Nothing OrElse DtCoste.Rows.Count = 0 Then
                    Dim DtNew As DataTable = DtCoste.Clone
                    Dim DrNew As DataRow = DtNew.NewRow
                    DrNew("IDCentroCoste") = dr("IDCentroCoste")
                    Dim strDescCentroCoste As String
                    If AppParams.NivelesDeAnalitica >= 1 Then
                        DrNew("IDAnalitica1") = DrNew("IDCentroCoste").Substring(0, NegocioGeneral.cnLENGTH_NIVELES_ANALITICA)
                        Dim ClsAnalitica As BusinessHelper = BusinessHelper.CreateBusinessObject("AnaliticaNivel1")
                        Dim DtAnalitica As DataTable = ClsAnalitica.SelOnPrimaryKey(DrNew("IDAnalitica1"))
                        If DtAnalitica.Rows.Count > 0 Then
                            strDescCentroCoste = DtAnalitica.Rows(0)("DescAnalitica1")
                        End If
                    End If
                    If AppParams.NivelesDeAnalitica >= 2 Then
                        DrNew("IDAnalitica2") = DrNew("IDCentroCoste").Substring(NegocioGeneral.cnLENGTH_NIVELES_ANALITICA, NegocioGeneral.cnLENGTH_NIVELES_ANALITICA)
                        Dim ClsAnalitica As BusinessHelper = BusinessHelper.CreateBusinessObject("AnaliticaNivel2")
                        Dim DtAnalitica As DataTable = ClsAnalitica.SelOnPrimaryKey(DrNew("IDAnalitica2"))
                        If DtAnalitica.Rows.Count > 0 Then
                            strDescCentroCoste = strDescCentroCoste & " - " & DtAnalitica.Rows(0)("DescAnalitica2")
                        End If
                    End If
                    If AppParams.NivelesDeAnalitica >= 3 Then
                        DrNew("IDAnalitica3") = DrNew("IDCentroCoste").Substring((2 * NegocioGeneral.cnLENGTH_NIVELES_ANALITICA), NegocioGeneral.cnLENGTH_NIVELES_ANALITICA)
                        Dim ClsAnalitica As BusinessHelper = BusinessHelper.CreateBusinessObject("AnaliticaNivel3")
                        Dim DtAnalitica As DataTable = ClsAnalitica.SelOnPrimaryKey(DrNew("IDAnalitica3"))
                        If DtAnalitica.Rows.Count > 0 Then
                            strDescCentroCoste = strDescCentroCoste & " - " & DtAnalitica.Rows(0)("DescAnalitica3")
                        End If
                    End If
                    If AppParams.NivelesDeAnalitica >= 4 Then
                        DrNew("IDAnalitica4") = DrNew("IDCentroCoste").Substring((3 * NegocioGeneral.cnLENGTH_NIVELES_ANALITICA), NegocioGeneral.cnLENGTH_NIVELES_ANALITICA)
                        Dim ClsAnalitica As BusinessHelper = BusinessHelper.CreateBusinessObject("AnaliticaNivel4")
                        Dim DtAnalitica As DataTable = ClsAnalitica.SelOnPrimaryKey(DrNew("IDAnalitica4"))
                        If DtAnalitica.Rows.Count > 0 Then
                            strDescCentroCoste = strDescCentroCoste & " - " & DtAnalitica.Rows(0)("DescAnalitica4")
                        End If
                    End If
                    If AppParams.NivelesDeAnalitica >= 5 Then
                        DrNew("IDAnalitica5") = DrNew("IDCentroCoste").Substring((4 * NegocioGeneral.cnLENGTH_NIVELES_ANALITICA), NegocioGeneral.cnLENGTH_NIVELES_ANALITICA)
                        Dim ClsAnalitica As BusinessHelper = BusinessHelper.CreateBusinessObject("AnaliticaNivel5")
                        Dim DtAnalitica As DataTable = ClsAnalitica.SelOnPrimaryKey(DrNew("IDAnalitica5"))
                        If DtAnalitica.Rows.Count > 0 Then
                            strDescCentroCoste = strDescCentroCoste & " - " & DtAnalitica.Rows(0)("DescAnalitica5")
                        End If
                    End If
                    DrNew("DescCentroCoste") = strDescCentroCoste
                    DtNew.Rows.Add(DrNew)
                    ClsCoste.Update(DtNew)
                End If
            End If
        Next
        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.CommitTransaction, True, services)

    End Sub

#End Region

#Region " Business Rules "

    Public Overrides Function GetBusinessRules() As Engine.BE.BusinessRules
        Dim oBRL As New BusinessRules
        oBRL.Add("Porcentaje", AddressOf ProcesoComunes.ValidarValorNumerico)
        Return oBRL
    End Function

#End Region

End Class