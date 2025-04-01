Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class facturaElectronicaController
    Public FacturaElectronica As New FacturaElectronica
    Public sql As String
    Public objadapter As OleDbDataAdapter
    Public objdataset As DataSet

    Public Function getAllParam(ByRef parametro As String) As Boolean

        Try
            If ConectaBase() Then

                sql = "select [idFE]
      ,[idDTE] as Folio
     ,FORMAT([fecha_rem], 'dd/MM/yyyy') as Fecha
      ,[fecha_vencimiento]
      ,[rutComprador] as RUT
      ,[nomComprador] as Nombre
      ,[pagada]
      ,[n_ingreso]
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
      ,[montoFlete] from FacturaElectronica " & parametro & " order by idDTE desc"
                objadapter = New OleDbDataAdapter(sql, Conecta)
                objdataset = New DataSet
                objadapter.Fill(objdataset)

                If objdataset.Tables(0).Rows.Count > 0 Then

                    FacturaElectronica.idFE = objdataset.Tables(0).Rows(0).Item(0)
                    FacturaElectronica.idDTE = objdataset.Tables(0).Rows(0).Item(1)
                    FacturaElectronica.fecha_Rem = objdataset.Tables(0).Rows(0).Item(2)
                    FacturaElectronica.fecha_vencimiento = objdataset.Tables(0).Rows(0).Item(3)
                    FacturaElectronica.rutComprador = objdataset.Tables(0).Rows(0).Item(4)
                    FacturaElectronica.nomComprador = objdataset.Tables(0).Rows(0).Item(5)
                    FacturaElectronica.pagada = objdataset.Tables(0).Rows(0).Item(6)
                    FacturaElectronica.n_ingreso = objdataset.Tables(0).Rows(0).Item(7)
                    FacturaElectronica.anulada = objdataset.Tables(0).Rows(0).Item(8)
                    FacturaElectronica.guia = objdataset.Tables(0).Rows(0).Item(9)
                    FacturaElectronica.condicion = objdataset.Tables(0).Rows(0).Item(10)
                    FacturaElectronica.neto = objdataset.Tables(0).Rows(0).Item(11)
                    FacturaElectronica.iva = objdataset.Tables(0).Rows(0).Item(12)
                    FacturaElectronica.porcentaje = objdataset.Tables(0).Rows(0).Item(13)
                    FacturaElectronica.prontoPago = objdataset.Tables(0).Rows(0).Item(14)
                    FacturaElectronica.comision = objdataset.Tables(0).Rows(0).Item(15)
                    FacturaElectronica.administracion = objdataset.Tables(0).Rows(0).Item(16)
                    FacturaElectronica.veterinario = objdataset.Tables(0).Rows(0).Item(17)
                    FacturaElectronica.IVAcomision = objdataset.Tables(0).Rows(0).Item(18)
                    FacturaElectronica.total = objdataset.Tables(0).Rows(0).Item(19)
                    FacturaElectronica.totalDecomiso = objdataset.Tables(0).Rows(0).Item(20)
                    FacturaElectronica.porPagar = objdataset.Tables(0).Rows(0).Item(21)
                    FacturaElectronica.idRemate = objdataset.Tables(0).Rows(0).Item(22)
                    FacturaElectronica.rutFletero = objdataset.Tables(0).Rows(0).Item(23)
                    FacturaElectronica.montoFlete = objdataset.Tables(0).Rows(0).Item(24)

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

                Dim consulta As String = "UPDATE FacturaElectronica SET " &
                                    "csv = 1 WHERE idDte = @idDte"

                Using enunciado As New SqlCommand(consulta, conexiones)
                    ' Asignar valores a los parámetros
                    enunciado.Parameters.AddWithValue("@idDTE", FacturaElectronica.idDTE)

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


    Sub generar_csv_softland(ByRef bruto As Integer, ByRef factura As String, ByRef numRem As String, ByRef rutCliente As String, ByRef fechaEmision As String, ByRef fechaVencimiento As String, ByRef comision As Integer, ByRef otros As Integer, ByRef facturaS As String, ByRef nombre As String, ByRef neto As Integer, ByRef iva As Integer, ByRef administracion As Integer, ByRef ivaComision As Integer, ByRef montoFlete As Integer)

        Dim ANSIString1() As Byte, MyEncoder As New System.Text.ASCIIEncoding()
        Dim ANSIString2() As Byte
        Dim ANSIString3() As Byte
        Dim ANSIString4() As Byte
        Dim ANSIString5() As Byte
        Dim ANSIString6() As Byte

        ' Obtener todos los valores necesarios una sola vez antes de los reintentos
        Dim num As String = If(idRecinto = 1, "008", If(idRecinto = 2, "108", ""))
        
        ' Obtener el usuario y fecha una sola vez
        Dim usu As Microsoft.VisualBasic.ApplicationServices.User = New Microsoft.VisualBasic.ApplicationServices.User
        Dim fecha As String = Date.Now
        Dim usuario() As String = Split(usu.Name, "\")
        Dim fechaArchivo As Date = Convert.ToDateTime(fecha_remate)

        ' Obtener solo el número de remate sin incrementar el correlativo
        Dim remateController As New remateController
        If remateController.getRemateActual() Then
            numRem = remateController.remate.remate
        End If

        ' Preparar el nombre del archivo
        Dim sFilename As String = "S:\fbiobio" & fechaArchivo.ToString("yyyyMMdd") & ".CSV"

        ' Asegurarnos de que las fechas estén en el formato correcto
        Dim fechaEmisionFormatted As String = Convert.ToDateTime(fechaEmision).ToString("yyyy-MM-dd")
        Dim fechaVencimientoFormatted As String = Convert.ToDateTime(fechaVencimiento).ToString("yyyy-MM-dd")

        ' Preparar todas las cadenas antes de intentar escribir
        Dim sSourceString1 As String = "001,1-1-3-01-03," & bruto & ",0,FACTURA VENTA N" & factura & " REM. " & numRem & ",0,0,,,,,,,,,,,,0," & rutCliente.Remove(rutCliente.Length - 2) & ",VB," & factura & ", " & fechaEmisionFormatted & ", " & fechaEmisionFormatted & ",VB," & factura & " ,," & neto & ", " & comision & ",0, " & administracion & " ," & montoFlete & ",0,0," & iva + ivaComision & ",0," & bruto & ",,," & factura & ",S,T,T,FACTURA VENTA REMATE " & numRem
        Dim sSourceString2 As String = "001,4-9-1-01-01,0," & neto & ",TRANSACCIONES " & nombre.Substring(0, Math.Min(nombre.Length, 15)) & ",0,0,,,,,,,,,," & num & ",,0,,,," & fechaEmisionFormatted & "," & fechaEmisionFormatted & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,FACTURA VENTA REMATE " & numRem
        Dim sSourceString3 As String = "001,4-1-1-01-01,0," & comision & ",COMISIONES PERCIBIDAS COMPRADO,0,0,,,,,,,,,," & num & ",,0,,,," & fechaEmisionFormatted & "," & fechaEmisionFormatted & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,FACTURA VENTA REMATE " & numRem
        Dim sSourceString4 As String = "001,4-1-2-01-01,0," & administracion & ",INGRESOS ADMINISTRATIVOS COMPR,0,0,,,,,,,,,," & num & ",,0,,,," & fechaEmisionFormatted & "," & fechaEmisionFormatted & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,FACTURA VENTA REMATE " & numRem
        Dim sSourceString6 As String = "001,4-1-3-01-01,0," & montoFlete & ",INGRESOS FLETES GANADO COMPR,0,0,,,,,,,,,," & num & ",,0,,,," & fechaEmisionFormatted & "," & fechaEmisionFormatted & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,FACTURA VENTA REMATE " & numRem
        Dim sSourceString5 As String = "001,2-1-4-01-01,0," & iva + ivaComision & ",IVA DEBITO FISCAL" & nombre.Substring(0, Math.Min(nombre.Length, 12)) & ",0,0,,,,,,,,,,,,0,,,," & fechaEmisionFormatted & "," & fechaEmisionFormatted & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,FACTURA VENTA REMATE " & numRem

        ' Convertir todas las cadenas a ANSI antes del bucle de reintento
        ANSIString1 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString1 + System.Environment.NewLine))
        ANSIString2 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString2 + System.Environment.NewLine))
        ANSIString3 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString3 + System.Environment.NewLine))
        ANSIString4 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString4 + System.Environment.NewLine))
        ANSIString6 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString6 + System.Environment.NewLine))
        ANSIString5 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString5 + System.Environment.NewLine))

        Dim retries As Integer = 0
        Dim maxRetries As Integer = 5
        Dim success As Boolean = False

        ' Intentar escribir el archivo con manejo de errores
        Do While Not success And retries < maxRetries
            Try
                If Not System.IO.File.Exists(sFilename) Then
                    Using fStream As New System.IO.FileStream(sFilename, System.IO.FileMode.Create, System.IO.FileAccess.Write)
                        If montoFlete > 0 Then
                            fStream.Write(ANSIString1, 0, ANSIString1.Length)
                            fStream.Write(ANSIString2, 0, ANSIString2.Length)
                            fStream.Write(ANSIString3, 0, ANSIString3.Length)
                            fStream.Write(ANSIString4, 0, ANSIString4.Length)
                            fStream.Write(ANSIString6, 0, ANSIString6.Length)
                            fStream.Write(ANSIString5, 0, ANSIString5.Length)
                        Else
                            fStream.Write(ANSIString1, 0, ANSIString1.Length)
                            fStream.Write(ANSIString2, 0, ANSIString2.Length)
                            fStream.Write(ANSIString3, 0, ANSIString3.Length)
                            fStream.Write(ANSIString4, 0, ANSIString4.Length)
                            fStream.Write(ANSIString5, 0, ANSIString5.Length)
                        End If
                    End Using
                    updateCSVStatus(factura)
                    success = True
                End If
            Catch ex As Exception
                retries += 1
                If retries < maxRetries Then
                    ' Esperar antes de reintentar
                    Threading.Thread.Sleep(5000)
                Else
                    MsgBox("Error al escribir el archivo después de varios intentos: " & ex.Message)
                End If
            End Try
        Loop
    End Sub


End Class
