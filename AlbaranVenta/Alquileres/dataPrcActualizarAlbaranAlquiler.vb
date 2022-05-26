<Serializable()> _
Public Class dataPrcActualizarAlbaranAlquiler
    Public RstAlbaranAlquiler As ResultAlbaranAlquiler

    Public Conductores As DataTable
    Public Contadores As DataTable
    Public Avisos As DataTable
    Public Incidencias As DataTable
    Public Activos As DataTable
    Public ADDObraMaterial As DataTable
    Public Retornos As Boolean
    Public SalidaRetornos As Boolean
    Public CambioMaquina As Boolean

    Public Sub New(ByVal RstAlbaranAlquiler As ResultAlbaranAlquiler)
        Me.RstAlbaranAlquiler = RstAlbaranAlquiler
    End Sub

End Class
