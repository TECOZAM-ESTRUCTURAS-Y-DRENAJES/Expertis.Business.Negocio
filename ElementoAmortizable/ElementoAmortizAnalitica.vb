Public Class ElementoAmortizAnalitica

#Region "Constructor"

    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper

    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbMaestroElementoAmortizAnalitica"

#End Region

#Region "Eventos ElementoAmortizAnalitica"

    Protected Overrides Sub RegisterValidateTasks(ByVal validateProcess As Engine.BE.BusinessProcesses.Process)
        MyBase.RegisterValidateTasks(validateProcess)
        validateProcess.AddTask(Of DataRow)(AddressOf ValidarDatosObligatorios)
    End Sub

    <Task()> Public Shared Sub ValidarDatosObligatorios(ByVal data As DataRow, ByVal services As ServiceProvider)
        If Length(data("Porcentaje")) = 0 Then
            ApplicationService.GenerateError("El porcentaje es obligatorio")
        Else
            If data("Porcentaje") < 0 OrElse data("Porcentaje") > 100 Then
                ApplicationService.GenerateError("El porcentaje debe estar comprendido entre 0 y 100")
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
                    DrNew("DescCentroCoste") = "Desc. Centro: " & dr("IDCentroCoste")
                    If AppParams.NivelesDeAnalitica >= 1 Then
                        DrNew("IDAnalitica1") = DrNew("IDCentroCoste").Substring(0, NegocioGeneral.cnLENGTH_NIVELES_ANALITICA)
                    End If
                    If AppParams.NivelesDeAnalitica >= 2 Then
                        DrNew("IDAnalitica2") = DrNew("IDCentroCoste").Substring(NegocioGeneral.cnLENGTH_NIVELES_ANALITICA, NegocioGeneral.cnLENGTH_NIVELES_ANALITICA)
                    End If
                    If AppParams.NivelesDeAnalitica >= 3 Then
                        DrNew("IDAnalitica3") = DrNew("IDCentroCoste").Substring((2 * NegocioGeneral.cnLENGTH_NIVELES_ANALITICA), NegocioGeneral.cnLENGTH_NIVELES_ANALITICA)
                    End If
                    If AppParams.NivelesDeAnalitica >= 4 Then
                        DrNew("IDAnalitica4") = DrNew("IDCentroCoste").Substring((3 * NegocioGeneral.cnLENGTH_NIVELES_ANALITICA), NegocioGeneral.cnLENGTH_NIVELES_ANALITICA)
                    End If
                    If AppParams.NivelesDeAnalitica >= 5 Then
                        DrNew("IDAnalitica5") = DrNew("IDCentroCoste").Substring((4 * NegocioGeneral.cnLENGTH_NIVELES_ANALITICA), NegocioGeneral.cnLENGTH_NIVELES_ANALITICA)
                    End If
                    DtNew.Rows.Add(DrNew)
                    ClsCoste.Update(DtNew)
                End If
            End If
        Next
        ProcessServer.ExecuteTask(Of Boolean)(AddressOf Business.General.Comunes.CommitTransaction, True, services)

    End Sub

#End Region

End Class