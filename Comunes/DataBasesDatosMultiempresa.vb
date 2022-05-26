<Serializable()> _
Public Class DataBasesDatosMultiempresa
    Public IDBaseDatosPrincipal As Guid?
    Public DescBaseDatosPrincipal As String
    Public BaseDatosPrincipal As String
    Public IDBaseDatosSecundaria As Guid?
    Public DescBaseDatosSecundaria As String
    Public BaseDatosSecundaria As String

    Public Sub New(ByVal IDBaseDatosPrincipal As Guid, ByVal IDBaseDatosSecundaria As Guid)
        Me.IDBaseDatosPrincipal = IDBaseDatosPrincipal
        Me.IDBaseDatosSecundaria = IDBaseDatosSecundaria
    End Sub

    Public Sub New()

    End Sub
End Class
