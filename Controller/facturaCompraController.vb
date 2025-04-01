Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class facturaCompraController
    Public FacturaCompra As New FacturaCompra
    Public sql As String
    Public objadapter As OleDbDataAdapter
    Public objdataset As DataSet

    Public Function getAllParametro(ByRef parametro) As Boolean

        Try
            If ConectaBase() Then

                sql = "select [idFC]
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
      ,[montoFlete]
      ,IVAretenido
      ,montoRetenido  as [Monto Retenido] from FacturaCompra  " & parametro & " and  idRemate = " & id_remate & "   order by idDTE desc"
                objadapter = New OleDbDataAdapter(sql, Conecta)
                objdataset = New DataSet
                objadapter.Fill(objdataset)

                If objdataset.Tables(0).Rows.Count > 0 Then

                    FacturaCompra.idFC = objdataset.Tables(0).Rows(0).Item(0)
                    FacturaCompra.idDTE = objdataset.Tables(0).Rows(0).Item(1)
                    FacturaCompra.fecha_Rem = objdataset.Tables(0).Rows(0).Item(2)
                    FacturaCompra.fecha_vencimiento = objdataset.Tables(0).Rows(0).Item(3)
                    FacturaCompra.rutVendedor = objdataset.Tables(0).Rows(0).Item(4)
                    FacturaCompra.nomVendedor = objdataset.Tables(0).Rows(0).Item(5)
                    FacturaCompra.pagada = objdataset.Tables(0).Rows(0).Item(6)
                    FacturaCompra.idEgreso = objdataset.Tables(0).Rows(0).Item(7)
                    FacturaCompra.anulada = objdataset.Tables(0).Rows(0).Item(8)
                    FacturaCompra.guia = objdataset.Tables(0).Rows(0).Item(9)
                    FacturaCompra.condicion = objdataset.Tables(0).Rows(0).Item(10)
                    FacturaCompra.neto = objdataset.Tables(0).Rows(0).Item(11)
                    FacturaCompra.iva = objdataset.Tables(0).Rows(0).Item(12)
                    FacturaCompra.porcentaje = objdataset.Tables(0).Rows(0).Item(13)
                    FacturaCompra.prontoPago = objdataset.Tables(0).Rows(0).Item(14)
                    FacturaCompra.comision = objdataset.Tables(0).Rows(0).Item(15)
                    FacturaCompra.administracion = objdataset.Tables(0).Rows(0).Item(16)
                    FacturaCompra.veterinario = objdataset.Tables(0).Rows(0).Item(17)
                    FacturaCompra.IVAcomision = objdataset.Tables(0).Rows(0).Item(18)
                    FacturaCompra.total = objdataset.Tables(0).Rows(0).Item(19)
                    FacturaCompra.totalDecomiso = objdataset.Tables(0).Rows(0).Item(20)
                    FacturaCompra.porPagar = objdataset.Tables(0).Rows(0).Item(21)
                    FacturaCompra.idRemate = objdataset.Tables(0).Rows(0).Item(22)
                    FacturaCompra.rutFletero = objdataset.Tables(0).Rows(0).Item(23)
                    FacturaCompra.montoFlete = objdataset.Tables(0).Rows(0).Item(24)
                    FacturaCompra.IVAretenido = objdataset.Tables(0).Rows(0).Item(25)
                    FacturaCompra.montoRetenido = objdataset.Tables(0).Rows(0).Item(26)

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

                Dim consulta As String = "UPDATE FacturaCompra SET " &
                                    "csv = 1 WHERE idDte = @idDte"

                Using enunciado As New SqlCommand(consulta, conexiones)
                    ' Asignar valores a los parámetros
                    enunciado.Parameters.AddWithValue("@idDTE", FacturaCompra.idDTE)

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

    Sub generar_csv_softland(ByRef neto As Integer, ByRef factura As Integer, ByRef remate As Integer, ByRef iva As Integer, ByRef nombre As String, ByRef rut As String, ByRef totalDescuentos As Integer, ByRef comision As Integer, ByRef otros As Integer, ByRef examen As Integer, ByRef decomiso As Integer, ByRef ivaComision As Integer, ByRef fechaRem As Date, ByRef fechaVen As Date, ByRef administracion As Integer, ByRef montoFlete As Integer, ByRef montoRetenido As Integer, ByRef IVAretenido As Integer, ByRef porPagar As Integer)

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
        Dim ANSIString12() As Byte
        Dim ANSIString13() As Byte


        Dim fFile As IO.File, fStream As IO.FileStream
        Dim remateController As New remateController
        Dim correlativo As Integer = 0
        Dim num As String = ""

        If idRecinto = 1 Then
            num = "008"
        End If
        If idRecinto = 2 Then
            num = "108"
        End If

        If remateController.correlativoSoftland1 Then
            correlativo = remateController.remate.correlativoSoftland
        End If

        Dim decomisoString As String = ""



        If decomiso <> 0 Then
            decomisoString = decomiso
        End If


        If nombre.Length <= 21 Then
            'agregar espacios para completar 21 caracteres
            For i As Integer = nombre.Length To 21
                nombre = nombre & " "
            Next
        End If

        Dim codidoRetencion As String = ""
        Dim retencionString As String = ""

        Dim ivaDescuentos As Integer = (comision + administracion + montoFlete + examen + decomiso) * 0.19



        If remateController.getRemateActual() Then
            remate = remateController.remate.remate
        End If

        Dim sSourceString4 As String = ""

        If IVAretenido = 19 Then
            codidoRetencion = "2-1-4-01-02"
            retencionString = "IVA RETENIDO TOTAL "
            sSourceString4 = "001,2-1-3-01-01,0," & neto & ",PROVEEDORES " & nombre.Substring(0, Math.Min(nombre.Length, 19)) & ",0,0,,,,,,,,,,,,0," & rut.Remove(rut.Length - 2) & ",CD," & factura & "," & fechaRem & "," & fechaVen & ",CD," & factura & "," & correlativo & "," & neto & "," & iva & "," & iva & ",0,0,0,0,0,0," & neto & ",,," & factura & ",S,T,T,FACTURA COMPRA REMATE " & remate

        ElseIf IVAretenido = 8 Then
            codidoRetencion = "2-1-4-01-03"
            retencionString = "IVA RETENIDO PARCIAL "
            sSourceString4 = "001,2-1-3-01-01,0," & CInt(neto + (neto * 0.19) - iva) & ",PROVEEDORES " & nombre.Substring(0, Math.Min(nombre.Length, 19)) & ",0,0,,,,,,,,,,,,0," & rut.Remove(rut.Length - 2) & ",CD," & factura & "," & fechaRem & "," & fechaVen & ",CD," & factura & "," & correlativo & "," & neto & "," & CInt(neto * 0.19) & ",0," & montoRetenido & ",0,0,0,0,0," & CInt(neto + (neto * 0.19) - iva) & ",,," & factura & ",S,T,T,FACTURA COMPRA REMATE " & remate
        End If



        Dim sSourceString1 As String = "001,4-9-1-01-01," & neto & ",0,FACTURA COMPRA N° " & factura & " REM. " & remate & ",0,0,,,,,,,,,," & num & ",,0,,,," & fechaRem & "," & fechaVen & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,FACTURA COMPRA REMATE " & remate

        Dim sSourceString2 As String = ""

        If IVAretenido = 19 Then
            sSourceString2 = "001,1-1-8-01-02," & iva & ",0," & nombre & ",0,0,,,,,,,,,,,,0,,,," & fechaRem & "," & fechaVen & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,FACTURA COMPRA REMATE " & remate
        ElseIf IVAretenido = 8 Then
            sSourceString2 = "001,1-1-8-01-02," & CInt(neto * 0.19) & ",0," & nombre & ",0,0,,,,,,,,,,,,0,,,," & fechaRem & "," & fechaVen & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,FACTURA COMPRA REMATE " & remate

        End If


        Dim sSourceString3 As String = "001," & codidoRetencion & ",0," & montoRetenido & "," & retencionString & " " & nombre.Substring(0, 8) & ",0,0,,,,,,,,,,,,0,,,," & fechaRem & ", " & fechaRem & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,FACTURA COMPRA REMATE " & remate

        Dim sSourceString5 As String = "001,1-1-3-01-03," & ivaDescuentos + comision + administracion + montoFlete + examen + decomiso & ",0,CLIENTES " & nombre.Substring(0, Math.Min(nombre.Length, 21)) & ",0,0,,,,,,,,,,,,0," & rut.Remove(rut.Length - 2) & ",VA," & factura & "," & fechaRem & "," & fechaVen & ",VA," & factura & ",,0,0," & comision & "," & administracion & "," & montoFlete & ", " & decomiso & "," & examen & "," & ivaDescuentos & ",0," & ivaDescuentos + comision + administracion + montoFlete + examen + decomiso & ",,," & factura & ",S,T,T,FACTURA COMPRA REMATE " & remate
        Dim sSourceString6 As String = "001,4-1-1-01-02,0," & comision & ",COMISIONES PERCIBIDAS VENDEDOR,0,0,,,,,,,,,," & num & ",,0,,,," & fechaRem & "," & fechaVen & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,FACTURA COMPRA REMATE " & remate
        Dim sSourceString7 As String = "001,4-1-2-01-02,0," & administracion & ",INGRESOS ADMINISTRATIVOS VENDE,0,0,,,,,,,,,," & num & ",,0,,,," & fechaRem & "," & fechaVen & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,FACTURA COMPRA REMATE " & remate
        Dim sSourceString13 As String = "001,4-1-3-01-01,0," & montoFlete & ",INGRESOS FLETES GANADO  " & nombre.Substring(0, 4) & ",0,0,,,,,,,,,," & num & ",,0,,,," & fechaRem & "," & fechaVen & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,FACTURA COMPRA REMATE " & remate
        Dim sSourceString8 As String = "001,4-1-3-01-02,0," & examen & ",INGRESOS EXAMENES GANADO " & nombre.Substring(0, 4) & ",0,0,,,,,,,,,," & num & ",,0,,,," & fechaRem & "," & fechaVen & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,FACTURA COMPRA REMATE " & remate
        Dim sSourceString9 As String = "001,4-1-3-01-03,0," & decomisoString & ",INGRESOS POR DECOMISO " & nombre.Substring(0, 4) & ",0,0,,,,,,,,,," & num & ",,0,,,," & fechaRem & "," & fechaVen & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,FACTURA COMPRA REMATE " & remate
        Dim sSourceString10 As String = "001,2-1-4-01-01,0," & ivaDescuentos & ",IVA DEBITO FISCAL " & nombre.Substring(0, 12) & ",0,0,,,,,,,,,,,,0,,,," & fechaRem & "," & fechaVen & ",,,,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,FACTURA COMPRA REMATE " & remate
        Dim sSourceString11 As String = "001,2-1-3-01-01," & ivaDescuentos + comision + administracion + montoFlete + examen + decomiso & ",0,PROVEEDORES " & nombre.Substring(0, Math.Min(nombre.Length, 17)) & ",0,0,,,,,,,,,,,,0," & rut.Remove(rut.Length - 2) & ",VA," & factura & "," & fechaVen & "," & fechaVen & ",CD," & factura & ",,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,FACTURA COMPRA REMATE " & remate
        Dim sSourceString12 As String = "001,1-1-3-01-03,0," & ivaDescuentos + comision + administracion + montoFlete + examen + decomiso & ",CLIENTES " & nombre.Substring(0, Math.Min(nombre.Length, 20)) & ",0,0,,,,,,,,,,,,0," & rut.Remove(rut.Length - 2) & ",CD," & factura & "," & fechaVen & "," & fechaVen & ",VA," & factura & ",,0,0,0,0,0,0,0,0,0,0,,," & factura & ",N,T,T,FACTURA COMPRA REMATE " & remate




        sSourceString1 = formateoTexto(sSourceString1)
        sSourceString2 = formateoTexto(sSourceString2)
        sSourceString3 = formateoTexto(sSourceString3)
        sSourceString4 = formateoTexto(sSourceString4)
        sSourceString5 = formateoTexto(sSourceString5)
        sSourceString6 = formateoTexto(sSourceString6)
        sSourceString7 = formateoTexto(sSourceString7)
        sSourceString8 = formateoTexto(sSourceString8)
        sSourceString9 = formateoTexto(sSourceString9)
        sSourceString10 = formateoTexto(sSourceString10)
        sSourceString11 = formateoTexto(sSourceString11)
        sSourceString12 = formateoTexto(sSourceString12)
        sSourceString13 = formateoTexto(sSourceString13)


        Dim usu As Microsoft.VisualBasic.ApplicationServices.User
        usu = New Microsoft.VisualBasic.ApplicationServices.User
        Dim fecha As String = Date.Now
        Dim usuario() As String = Split(usu.Name, "\")

        Dim fechaRemate As Date = fecha_remate
        Dim sFilename As String = "S:\fbiobio" & fechaRemate.ToString("yyyyMMdd") & ".CSV"

        'Dim sFilename As String = "X:\Softland\FC_" & factura & ".CSV"


        ' Store the ANSI encoded string in ANSIString
        ANSIString1 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString1 + System.Environment.NewLine))
        ANSIString2 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString2 + System.Environment.NewLine))
        ANSIString3 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString3 + System.Environment.NewLine))
        ANSIString4 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString4 + System.Environment.NewLine))
        ANSIString5 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString5 + System.Environment.NewLine))
        ANSIString6 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString6 + System.Environment.NewLine))
        ANSIString7 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString7 + System.Environment.NewLine))
        ANSIString8 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString8 + System.Environment.NewLine))
        ANSIString9 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString9 + System.Environment.NewLine))
        ANSIString10 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString10 + System.Environment.NewLine))
        ANSIString11 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString11 + System.Environment.NewLine))
        ANSIString12 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString12 + System.Environment.NewLine))
        ANSIString13 = MyEncoder.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.ASCII, MyEncoder.GetBytes(sSourceString13 + System.Environment.NewLine))



        Dim retries As Integer = 0
        Dim maxRetries As Integer = 5
        Dim success As Boolean = False

        ' Intentar escribir el archivo con manejo de errores
        Do While Not success And retries < maxRetries

            Try


                'Eliminar el archivo si ya existe 
                If fFile.Exists(sFilename) Then

                Else

                    If montoFlete <> 0 Then
                        fStream = fFile.OpenWrite(sFilename)
                        fStream.Write(ANSIString1, 0, ANSIString1.Length)
                        fStream.Write(ANSIString2, 0, ANSIString2.Length)
                        fStream.Write(ANSIString3, 0, ANSIString3.Length)
                        fStream.Write(ANSIString4, 0, ANSIString4.Length)
                        fStream.Write(ANSIString5, 0, ANSIString5.Length)
                        fStream.Write(ANSIString6, 0, ANSIString6.Length)
                        fStream.Write(ANSIString7, 0, ANSIString7.Length)
                        fStream.Write(ANSIString13, 0, ANSIString13.Length) 'flete
                    Else
                        fStream = fFile.OpenWrite(sFilename)
                        fStream.Write(ANSIString1, 0, ANSIString1.Length)
                        fStream.Write(ANSIString2, 0, ANSIString2.Length)
                        fStream.Write(ANSIString3, 0, ANSIString3.Length)
                        fStream.Write(ANSIString4, 0, ANSIString4.Length)
                        fStream.Write(ANSIString5, 0, ANSIString5.Length)
                        fStream.Write(ANSIString6, 0, ANSIString6.Length)
                        fStream.Write(ANSIString7, 0, ANSIString7.Length)
                    End If



                    If decomiso <> 0 Then
                        If examen <> 0 Then
                            fStream.Write(ANSIString8, 0, ANSIString8.Length) ' examen
                        End If
                        fStream.Write(ANSIString9, 0, ANSIString9.Length) ' decomiso
                        fStream.Write(ANSIString10, 0, ANSIString10.Length)
                        fStream.Write(ANSIString11, 0, ANSIString11.Length)
                        fStream.Write(ANSIString12, 0, ANSIString12.Length)
                        fStream.Close()
                    Else
                        If examen <> 0 Then
                            fStream.Write(ANSIString8, 0, ANSIString8.Length) ' examen
                        End If
                        fStream.Write(ANSIString10, 0, ANSIString10.Length)
                        fStream.Write(ANSIString11, 0, ANSIString11.Length)
                        fStream.Write(ANSIString12, 0, ANSIString12.Length)
                        fStream.Close()
                    End If

                    updateCSVStatus(factura)
                    Exit Sub

                End If

            Catch ex As Exception
                retries += 1
                If retries < maxRetries Then
                    ' Esperar 5 segundos antes de reintentar
                    Threading.Thread.Sleep(5000)
                Else
                    ' Si no fue posible escribir el archivo después de varios intentos, mostrar el error
                    MsgBox("Error al escribir el archivo después de varios intentos: " & ex.Message)
                End If
            End Try
        Loop

    End Sub


End Class
