Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class liquidacionFacturaController
    Public LiquidacionFacturaElectronica As New LiquidacionFacturaElectronica
    Public sql As String
    Public objadapter As OleDbDataAdapter
    Public objdataset As DataSet
    Public Function getAllParametro(ByRef parametro As String) As Boolean

        Try
            If ConectaBase() Then

                sql = "select  [idLiquidacion]
      ,[idDTE] as Folio
      ,[fecha_rem] as Fecha
      ,[fecha_vencimiento]
      ,[rutVendedor] as RUT
      ,[nomVendedor] as Nombre
      ,[pagada]
      ,[idEgreso]
      ,[anulada]
      ,[guia]
      ,[condicion]
      ,[neto] as Neto
      ,[iva] as IVA
      ,[porcentaje] 
      ,[prontoPago]
      ,[comision] as Comisión
      ,[administracion] as Administración
      ,[veterinario] as Veterinario
      ,[IVAcomision] as [IVA Comisión] 
      ,[total] as Total
      ,[totalDecomiso] as Decomiso
      ,[porPagar] as [Por Pagar]
      ,[idRemate]
      ,[rutFletero] as [RUT Fletero]
      ,[montoFlete] as Flete from LiquidacionFacturaElectronica " & parametro & " and  idRemate = " & id_remate & " order by idDTE desc"
                objadapter = New OleDbDataAdapter(sql, Conecta)
                objdataset = New DataSet
                objadapter.Fill(objdataset)

                If objdataset.Tables(0).Rows.Count > 0 Then

                    LiquidacionFacturaElectronica.idLiquidacion = objdataset.Tables(0).Rows(0).Item(0)
                    LiquidacionFacturaElectronica.idDTE = objdataset.Tables(0).Rows(0).Item(1)
                    LiquidacionFacturaElectronica.fecha_Rem = objdataset.Tables(0).Rows(0).Item(2)
                    LiquidacionFacturaElectronica.fecha_vencimiento = objdataset.Tables(0).Rows(0).Item(3)
                    LiquidacionFacturaElectronica.rutVendedor = objdataset.Tables(0).Rows(0).Item(4)
                    LiquidacionFacturaElectronica.nomVendedor = objdataset.Tables(0).Rows(0).Item(5)
                    LiquidacionFacturaElectronica.pagada = objdataset.Tables(0).Rows(0).Item(6)
                    LiquidacionFacturaElectronica.idEgreso = objdataset.Tables(0).Rows(0).Item(7)
                    LiquidacionFacturaElectronica.anulada = objdataset.Tables(0).Rows(0).Item(8)
                    LiquidacionFacturaElectronica.guia = objdataset.Tables(0).Rows(0).Item(9)
                    LiquidacionFacturaElectronica.condicion = objdataset.Tables(0).Rows(0).Item(10)
                    LiquidacionFacturaElectronica.neto = objdataset.Tables(0).Rows(0).Item(11)
                    LiquidacionFacturaElectronica.iva = objdataset.Tables(0).Rows(0).Item(12)
                    LiquidacionFacturaElectronica.porcentaje = objdataset.Tables(0).Rows(0).Item(13)
                    LiquidacionFacturaElectronica.prontoPago = objdataset.Tables(0).Rows(0).Item(14)
                    LiquidacionFacturaElectronica.comision = objdataset.Tables(0).Rows(0).Item(15)
                    LiquidacionFacturaElectronica.administracion = objdataset.Tables(0).Rows(0).Item(16)
                    LiquidacionFacturaElectronica.veterinario = objdataset.Tables(0).Rows(0).Item(17)
                    LiquidacionFacturaElectronica.IVAcomision = objdataset.Tables(0).Rows(0).Item(18)
                    LiquidacionFacturaElectronica.total = objdataset.Tables(0).Rows(0).Item(19)
                    LiquidacionFacturaElectronica.totalDecomiso = objdataset.Tables(0).Rows(0).Item(20)
                    LiquidacionFacturaElectronica.porPagar = objdataset.Tables(0).Rows(0).Item(21)
                    LiquidacionFacturaElectronica.idRemate = objdataset.Tables(0).Rows(0).Item(22)
                    LiquidacionFacturaElectronica.rutFletero = objdataset.Tables(0).Rows(0).Item(23)
                    LiquidacionFacturaElectronica.montoFlete = objdataset.Tables(0).Rows(0).Item(24)

                    Return True
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return False
    End Function

    Public Function updateCSVStatus(ByVal idDTE As String) As Boolean
        Try
            Using conexiones As New SqlConnection(conexion_string)
                conexiones.Open()

                Dim consulta As String = "UPDATE LiquidacionFacturaElectronica SET " &
                                    "csv = 1 WHERE idDte = @idDte"

                Using enunciado As New SqlCommand(consulta, conexiones)
                    ' Asignar valores a los parámetros
                    enunciado.Parameters.AddWithValue("@idDTE", LiquidacionFacturaElectronica.idDTE)

                    ' Ejecutar la consulta
                    enunciado.ExecuteNonQuery()
                End Using
            End Using
            Return True
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False
        End Try
    End Function

    Sub generar_csv_softland(ByRef bruto As Integer, ByRef factura As String, ByRef numRem As String, ByRef rutCliente As String, ByRef fechaEmision As String, ByRef fechaVencimiento As String, ByRef comision As Integer, ByRef otros As Integer, ByRef facturaS As String, ByRef nombre As String, ByRef neto As Integer, ByRef iva As Integer, ByRef ivaComision As Integer, ByRef administrativos As Integer, ByRef descuentos As Integer, ByRef vet As Integer, ByRef decomiso As Integer)


        Try
            Dim fFile As IO.File, fStream As IO.FileStream
            Dim fechaRemate As Date = fecha_remate
            Dim sFilename As String = "S:\fbiobio" & fechaRemate.ToString("yyyyMMdd") & ".CSV"

            'Eliminar el archivo si ya existe 
            If fFile.Exists(sFilename) Then
            Else

                Dim ANSIString1() As Byte, MyEncoder As New System.Text.ASCIIEncoding()
                Dim ANSIString2() As Byte
                Dim ANSIString3() As Byte
                Dim ANSIString4() As Byte
                Dim ANSIString5() As Byte
                Dim ANSIString6() As Byte
                Dim ANSIString7() As Byte
                Dim ANSIString8() As Byte
                Dim ANSIString9() As Byte
                Dim ANSIString10() As Byte
                Dim ANSIString11() As Byte

                Dim num As String = ""
                Dim correlativo As Integer = 0

                If idRecinto = 1 Then
                    num = "008"
                End If
                If idRecinto = 2 Then
                    num = "108"
                End If



                Dim remateController As New remateController
                If remateController.getRemateActual() Then
                    numRem = remateController.remate.remate
                End If

                If remateController.correlativoSoftland1 Then
                    correlativo = remateController.remate.correlativoSoftland
                Else
                    correlativo = 0
                End If





                Dim sSourceString1 As String = "001,4-9-1-01-01," & neto & ",0,LIQ. FACTURA N° " & factura & " REM. " & numRem & ",0,0,,,,,,,,,," & num & ",,0,,,," & fechaEmision & "," & fechaVencimiento & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,LIQUIDACION FACTURA REMATE " & numRem
                Dim sSourceString2 As String = "001,1-1-8-01-02," & iva & ",0," & nombre & ",0,0,,,,,,,,,,,,0,,,," & fechaEmision & "," & fechaVencimiento & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,LIQUIDACION FACTURA REMATE " & numRem
                Dim sSourceString3 As String = "001,2-1-3-01-01,0," & neto + iva & ",PROVEEDORES " & nombre.Substring(0, Math.Min(nombre.Length, 17)) & ",0,0,,,,,,,,,,,,0," & rutCliente.Remove(rutCliente.Length - 2) & ",CF," & factura & "," & fechaEmision & "," & fechaVencimiento & ",CF," & factura & "," & correlativo & "," & neto & "," & iva & ",0,0,0,0,0,0,0," & neto + iva & ",,," & factura & ",S,T,T,LIQUIDACION FACTURA REMATE " & numRem
                Dim sSourceString4 As String = "001,1-1-3-01-03," & comision + administrativos + vet + ivaComision + decomiso & ",0," & nombre.Substring(0, Math.Min(nombre.Length, 19)) & ",0,0,,,,,,,,,,,,0," & rutCliente.Remove(rutCliente.Length - 2) & ",VA," & factura & "," & fechaEmision & "," & fechaVencimiento & ",VA," & factura & ",,0,0," & comision & "," & administrativos & ",0," & decomiso & "," & vet & "," & ivaComision & ",0," & comision + administrativos + vet + ivaComision + decomiso & ",,," & factura & ",S,T,T,LIQUIDACION FACTURA REMATE " & numRem
                Dim sSourceString5 As String = "001,4-1-1-01-02,0," & comision & ",COMISIONES PERCIBIDAS VENDEDOR,0,0,,,,,,,,,," & num & ",,0,,,," & fechaEmision & "," & fechaVencimiento & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,LIQUIDACION FACTURA REMATE " & numRem
                Dim sSourceString6 As String = "001,4-1-2-01-02,0," & administrativos & ",INGRESOS ADMINISTRATIVOS VENDE,0,0,,,,,,,,,," & num & ",,0,,,," & fechaEmision & "," & fechaVencimiento & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,LIQUIDACION FACTURA REMATE " & numRem

                Dim sSourceString11 As String = "001,4-1-3-01-02,0," & vet & ",INGRESOS EXAMENES " & nombre.Substring(0, Math.Min(nombre.Length, 17)) & ",0,0,,,,,,,,,," & num & ",,0,,,," & fechaEmision & "," & fechaVencimiento & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,LIQUIDACION FACTURA REMATE " & numRem
                Dim sSourceString10 As String = "001,4-1-3-01-03,0," & decomiso & ",INGRESOS DECOMISOS " & nombre.Substring(0, Math.Min(nombre.Length, 17)) & ",0,0,,,,,,,,,," & num & ",,0,,,," & fechaEmision & "," & fechaVencimiento & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,LIQUIDACION FACTURA REMATE " & numRem


                Dim sSourceString7 As String = "001,2-1-4-01-01,0," & ivaComision & ",IVA DEBITO FISCAL " & nombre.Substring(0, Math.Min(nombre.Length, 17)) & ",0,0,,,,,,,,,,,,0,,,," & fechaEmision & "," & fechaVencimiento & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,LIQUIDACION FACTURA REMATE " & numRem
                Dim sSourceString8 As String = "001,2-1-3-01-01," & comision + administrativos + vet + ivaComision + decomiso & ",0,PROVEEDORES " & nombre.Substring(0, Math.Min(nombre.Length, 17)) & ",0,0,,,,,,,,,,,,0," & rutCliente.Remove(rutCliente.Length - 2) & ",VA," & factura & "," & fechaEmision & "," & fechaVencimiento & ",CF," & factura & ",,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,LIQUIDACION FACTURA REMATE " & numRem
                Dim sSourceString9 As String = "001,1-1-3-01-03,0," & comision + administrativos + vet + ivaComision + decomiso & ",CLIENTES " & nombre.Substring(0, Math.Min(nombre.Length, 17)) & ",0,0,,,,,,,,,,,,0," & rutCliente.Remove(rutCliente.Length - 2) & ",CF," & factura & "," & fechaEmision & "," & fechaVencimiento & ",VA," & factura & ",,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,LIQUIDACION FACTURA REMATE " & numRem



                Dim usu As Microsoft.VisualBasic.ApplicationServices.User
                usu = New Microsoft.VisualBasic.ApplicationServices.User
                Dim fecha As String = Date.Now
                Dim usuario() As String = Split(usu.Name, "\")


                '  Dim sFilename As String = "X:\Softland\LF_" & facturaS & ".CSV"


                ' Store the ANSI encoded string in ANSIString
                ANSIString1 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString1 + System.Environment.NewLine))
                ANSIString2 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString2 + System.Environment.NewLine))
                ANSIString3 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString3 + System.Environment.NewLine))
                ANSIString4 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString4 + System.Environment.NewLine))
                ANSIString5 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString5 + System.Environment.NewLine))
                ANSIString6 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString6 + System.Environment.NewLine))
                ANSIString11 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString11 + System.Environment.NewLine))
                ANSIString7 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString7 + System.Environment.NewLine))
                ANSIString8 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString8 + System.Environment.NewLine))
                ANSIString9 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString9 + System.Environment.NewLine))
                ANSIString10 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString10 + System.Environment.NewLine))



                ' Output the bytes
                fStream = fFile.OpenWrite(sFilename)
                fStream.Write(ANSIString1, 0, ANSIString1.Length)
                fStream.Write(ANSIString2, 0, ANSIString2.Length)
                fStream.Write(ANSIString3, 0, ANSIString3.Length)
                fStream.Write(ANSIString4, 0, ANSIString4.Length)
                fStream.Write(ANSIString5, 0, ANSIString5.Length)
                fStream.Write(ANSIString6, 0, ANSIString6.Length)


                If decomiso <> 0 Then
                    fStream.Write(ANSIString10, 0, ANSIString10.Length) 'Decomiso
                End If

                If vet <> 0 Then
                    fStream.Write(ANSIString11, 0, ANSIString11.Length) 'Examen
                End If

                fStream.Write(ANSIString7, 0, ANSIString7.Length)
                fStream.Write(ANSIString8, 0, ANSIString8.Length)
                fStream.Write(ANSIString9, 0, ANSIString9.Length)


                fStream.Close()

                updateCSVStatus(factura)
                Exit Sub
            End If



        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

End Class
