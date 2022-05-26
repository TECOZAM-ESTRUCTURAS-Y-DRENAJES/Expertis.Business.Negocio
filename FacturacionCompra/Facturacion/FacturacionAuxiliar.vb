Public Class FacturacionAuxiliar
    Inherits Solmicro.Expertis.Engine.BE.BusinessHelper
#Region "Constructor"
    Public Sub New()
        MyBase.New(cnEntidad)
    End Sub

    Private Const cnEntidad As String = "tbFacturacionAux"
#End Region

End Class
