Public Class FacturaCompra
    Public idFC As String
    Public idLiquidacion As String
    Public idDTE As Integer
    Public fecha_Rem As Date
    Public fecha_vencimiento As Date
    Public rutVendedor As String
    Public nomVendedor As String
    Public pagada As Boolean
    Public idEgreso As String
    Public anulada As Boolean
    Public guia As Integer
    Public condicion As Integer
    Public neto As Integer
    Public iva As Integer
    Public porcentaje As Decimal
    Public prontoPago As Decimal
    Public comision As Integer
    Public administracion As Integer
    Public veterinario As Integer
    Public IVAcomision As Integer
    Public total As Integer
    Public totalDecomiso As Integer
    Public porPagar As Integer
    Public idRemate As Integer
    Public rutFletero As String
    Public montoFlete As Integer
    Public IVAretenido As Integer
    Public montoRetenido As Integer
    Public csv As Boolean

End Class
