<Serializable()> _
Public Class DataParamRiesgoCliente
    Public GestionAlquiler As Boolean
    Public TipoAlbaranRetornoAlquiler As String
    Public TipoAlbaranDeDeposito As String


    Public IDCliente As String
    Public RiesgoGrupo As Boolean
    Public IDClientesGrupo(-1) As Object
    Public IDClienteMatriz As String
    Public RiesgoConcedido As Double
    Public LimiteCapitalAsegurado As Double
    Public RiesgoInterno As Double

    Public IDProveedorAsociado As String
    Public DescProveedorAsociado As String
    Public PdteFacturar As Double
    Public PagosNoPagados As Double


End Class
