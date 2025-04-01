Imports System.Data.OleDb
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar

Public Class principalView
    Dim principalController As New principalController
    Private contador As Integer = 5
    Private procesandoDatos As Boolean = False

    Private Sub principalView_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If principalController.cargar_remate Then
            lb_remate.Text = num_remate
        End If

        Timer1.Interval = 1000
        Timer1.Enabled = True
        cargargrilla()
        actualizarContador()
    End Sub

    Private Async Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        contador -= 1
        actualizarContador()

        If contador <= 0 AndAlso Not procesandoDatos Then
            Timer1.Enabled = False ' Deshabilitar temporalmente el timer
            procesandoDatos = True

            Try
                Await cargargrilla()
            Finally
                contador = 5
                procesandoDatos = False
                Timer1.Enabled = True
            End Try
        End If
    End Sub

    Private Sub actualizarContador()
        If procesandoDatos Then
            lb_contador.Text = "Procesando datos..."
        Else
            lb_contador.Text = $"Próxima actualización en {contador} segundos"
        End If
    End Sub

    Private Async Function cargargrilla() As Task
        Try
            If Me.InvokeRequired Then
                Me.Invoke(Sub() grilla1.DataSource = Nothing)
            Else
                grilla1.DataSource = Nothing
            End If

            If principalController.estadoDocumentos(id_remate) Then
                Await ActualizarUI(Sub()
                                       grilla1.DataSource = principalController.objdataset.Tables(0)
                                       grilla1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

                                       Dim totalColumns As Integer = grilla1.Columns.Count
                                       If totalColumns >= 3 Then
                                           For i As Integer = totalColumns - 3 To totalColumns - 1
                                               grilla1.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                                           Next
                                       End If

                                       grilla1.Refresh()
                                   End Sub)

                ' Procesar documentos pendientes de forma asíncrona
                Await ProcesarDocumentosPendientes()
            End If
        Catch ex As Exception
            Me.Invoke(Sub() MessageBox.Show("Error al cargar la grilla: " & ex.Message))
        End Try
    End Function

    Private Async Function ActualizarUI(action As Action) As Task
        If Me.InvokeRequired Then
            Await Me.Invoke(action)
        Else
            action()
        End If
    End Function

    Private Async Function ProcesarDocumentosPendientes() As Task
        Try
            ' Verificar y procesar facturas electrónicas pendientes
            Await ProcesarFacturasElectronicas()
            Await ProcesarFacturasCompra()
            Await ProcesarLiquidaciones()
            Await ProcesarIngresos()
            Await ProcesarEgresos()
        Catch ex As Exception
            Me.Invoke(Sub() MessageBox.Show("Error procesando documentos: " & ex.Message))
        End Try
    End Function

    Private Async Function ProcesarFacturasElectronicas() As Task
        Try
            Dim facturaElectronicaController As New facturaElectronicaController
            Dim parametroVerificacion As String = "where isnull(csv,0) = 0 and idRemate = " & id_remate

            If facturaElectronicaController.getAllParam(parametroVerificacion) Then
                If facturaElectronicaController.objdataset IsNot Nothing AndAlso
                   facturaElectronicaController.objdataset.Tables(0).Rows.Count > 0 Then
                    Dim idFE As String = facturaElectronicaController.objdataset.Tables(0).Rows(0).Item(0).ToString()
                    Await Task.Run(Sub() consultarFacturaElectronica("where idFE = '" & idFE & "'"))
                End If
            End If
        Catch ex As Exception
            Me.Invoke(Sub() MessageBox.Show("Error procesando facturas electrónicas: " & ex.Message))
        End Try
    End Function

    Private Sub consultarFacturaElectronica(ByVal parametro As String)
        Dim facturaElectronicaController As New facturaElectronicaController

        If facturaElectronicaController.getAllParam(parametro) Then
            Try
                Dim dataTable As DataTable = facturaElectronicaController.objdataset.Tables(0)

                For Each fila As DataRow In dataTable.Rows
                    Dim idFE As String = fila.Item(0).ToString()
                    Dim bruto As Integer = Convert.ToInt32(fila.Item(19))
                    Dim factura As Integer = Convert.ToInt32(fila.Item(1))
                    Dim rutCliente As String = fila.Item(4).ToString()
                    Dim fechaEmision As String = fila.Item(2).ToString()
                    Dim fechaVencimiento As String = fila.Item(3).ToString()
                    Dim comision As Integer = Convert.ToInt32(fila.Item(15))
                    Dim otros As Integer = Convert.ToInt32(fila.Item(16)) + Convert.ToInt32(fila.Item(17)) + Convert.ToInt32(fila.Item(21))
                    Dim nombre As String = fila.Item(5).ToString()
                    Dim neto As Integer = Convert.ToInt32(fila.Item(11))
                    Dim iva As Integer = Convert.ToInt32(fila.Item(12))
                    Dim administracion As Integer = Convert.ToInt32(fila.Item(16))
                    Dim ivaComision As Integer = Convert.ToInt32(fila.Item(18))
                    Dim montoFlete As Integer = Convert.ToInt32(fila.Item(24))

                    facturaElectronicaController.generar_csv_softland(
                    bruto, factura, "", rutCliente, fechaEmision, fechaVencimiento,
                        comision, otros, factura.ToString(), nombre, neto, iva,
                        administracion, ivaComision, montoFlete)
                Next
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If
    End Sub

    Private Async Function ProcesarFacturasCompra() As Task
        Try
            Dim facturaCompraController As New facturaCompraController
            Dim parametro As String = "where isnull(csv,0) = 0 and idRemate = " & id_remate

            If facturaCompraController.getAllParametro(parametro) Then
                If facturaCompraController.objdataset IsNot Nothing AndAlso
                   facturaCompraController.objdataset.Tables(0).Rows.Count > 0 Then
                    Dim idFC As String = facturaCompraController.objdataset.Tables(0).Rows(0).Item(0).ToString()
                    Await Task.Run(Sub() consultarFacturaCompra("where idFC = '" & idFC & "'"))
                End If
            End If
        Catch ex As Exception
            Me.Invoke(Sub() MessageBox.Show("Error procesando facturas de compra: " & ex.Message))
        End Try
    End Function

    Private Async Function ProcesarLiquidaciones() As Task
        Try
            Dim liquidacionFacturaController As New liquidacionFacturaController
            Dim parametroLiquidacion As String = "where isnull(csv,0) = 0 and idRemate = " & id_remate

            If liquidacionFacturaController.getAllParametro(parametroLiquidacion) Then
                If liquidacionFacturaController.objdataset IsNot Nothing AndAlso
                   liquidacionFacturaController.objdataset.Tables(0).Rows.Count > 0 Then
                    Dim idLiquidacion As String = liquidacionFacturaController.objdataset.Tables(0).Rows(0).Item(0).ToString()
                    Await Task.Run(Sub() consultarLiquidacionFacturaElectronica("where idLiquidacion = '" & idLiquidacion & "'"))
                End If
            End If
        Catch ex As Exception
            Me.Invoke(Sub() MessageBox.Show("Error procesando liquidaciones: " & ex.Message))
        End Try
    End Function

    Private Async Function ProcesarIngresos() As Task
        Try
            Dim ingresoController As New ingresoController
            Dim parametroIngreso As String = "where isnull(csv,0) = 0 and i.anulado = 0 and idRemate = " & id_remate



            If ingresoController.getAllParam(parametroIngreso, "") Then
                If ingresoController.objdataset IsNot Nothing AndAlso
                   ingresoController.objdataset.Tables(0).Rows.Count > 0 Then
                    Dim idIngreso As String = ingresoController.objdataset.Tables(0).Rows(0).Item(0).ToString()

                    If ingresoController.getIniciarCSV(idIngreso) Then
                        Await Task.Run(Sub() consultarIngresos(parametroIngreso, idIngreso))
                    Else
                        Await Task.Run(Sub() consultarIngresos(parametroIngreso, idIngreso))
                    End If
                End If
            End If
        Catch ex As Exception
            Me.Invoke(Sub() MessageBox.Show("Error procesando ingresos: " & ex.Message))
        End Try
    End Function

    Private Async Function ProcesarEgresos() As Task
        Try
            Dim egresoController As New egresoController
            Dim parametroEgreso As String = "where isnull(csv,0) = 0 and isnull(e.compras,0) = 0 and e.anulado = 0 and idRemate = " & id_remate

            If egresoController.getAllParam(parametroEgreso) Then
                If egresoController.objdataset IsNot Nothing AndAlso
                   egresoController.objdataset.Tables(0).Rows.Count > 0 Then
                    Dim idEgreso As String = egresoController.objdataset.Tables(0).Rows(0).Item(0).ToString()
                    Await Task.Run(Sub() consultarEgresos("where e.idEgreso = '" & idEgreso & "'", idEgreso))
                End If
            End If
        Catch ex As Exception
            Me.Invoke(Sub() MessageBox.Show("Error procesando egresos: " & ex.Message))
        End Try

        Try
            Dim egresoController As New egresoController
            Dim parametroEgreso As String = "where isnull(csv,0) = 0 and isnull(e.compras,0) = 1 and e.anulado = 0 and idRemate = " & id_remate

            If egresoController.getAllParam(parametroEgreso) Then
                If egresoController.objdataset IsNot Nothing AndAlso
                   egresoController.objdataset.Tables(0).Rows.Count > 0 Then
                    Dim idEgreso As String = egresoController.objdataset.Tables(0).Rows(0).Item(0).ToString()
                    Await Task.Run(Sub() consultarEgresosCompras("where e.idEgreso = '" & idEgreso & "'", idEgreso))
                End If
            End If
        Catch ex As Exception
            Me.Invoke(Sub() MessageBox.Show("Error procesando egresos: " & ex.Message))
        End Try
    End Function

    Private Sub consultarFacturaCompra(ByVal parametro As String)
        Dim facturaCompraController As New facturaCompraController

        If facturaCompraController.getAllParametro(parametro) Then
            Try
                Dim dataTable As DataTable = facturaCompraController.objdataset.Tables(0)

                For Each fila As DataRow In dataTable.Rows
                    Dim neto As Integer = Convert.ToInt32(fila.Item(11))
                    Dim factura As Integer = Convert.ToInt32(fila.Item(1))
                    Dim remate As Integer = Convert.ToInt32(fila.Item(22))
                    Dim iva As Integer = Convert.ToInt32(fila.Item(12))
                    Dim nombre As String = fila.Item(5).ToString()
                    Dim rut As String = fila.Item(4).ToString()
                    Dim totalDescuentos As Integer = Convert.ToInt32(fila.Item(15)) + Convert.ToInt32(fila.Item(16)) + Convert.ToInt32(fila.Item(17)) + Convert.ToInt32(fila.Item(18))
                    Dim comision As Integer = Convert.ToInt32(fila.Item(15))
                    Dim otros As Integer = Convert.ToInt32(fila.Item(16)) + Convert.ToInt32(fila.Item(24))
                    Dim examen As Integer = Convert.ToInt32(fila.Item(17))
                    Dim decomiso As Integer = Convert.ToInt32(fila.Item(20))
                    Dim ivaComision As Integer = Convert.ToInt32(fila.Item(18))
                    Dim fechaRem As String = fila.Item(2).ToString()
                    Dim fechaVencimiento As String = fila.Item(3).ToString()
                    Dim administracion As Integer = Convert.ToInt32(fila.Item(16))
                    Dim montoFlete As Integer = Convert.ToInt32(fila.Item(24))
                    Dim IVAretenido As Integer = Convert.ToInt32(fila.Item(25))
                    Dim montoRetenido As Integer = Convert.ToInt32(fila.Item(26))
                    Dim porPagar As Integer = Convert.ToInt32(fila.Item(21))

                    facturaCompraController.generar_csv_softland(neto, factura, remate, iva, nombre, rut, totalDescuentos, comision, otros, examen, decomiso, ivaComision, fechaRem, fechaVencimiento, administracion, montoFlete, montoRetenido, IVAretenido, porPagar)
                Next
            Catch ex As Exception
                Me.Invoke(Sub() MessageBox.Show("Error procesando factura de compra: " & ex.Message))
            End Try
        End If
    End Sub

    Private Sub consultarLiquidacionFacturaElectronica(ByVal parametro As String)
        Dim liquidacionFacturaController As New liquidacionFacturaController

        If liquidacionFacturaController.getAllParametro(parametro) Then
            Try
                Dim dataTable As DataTable = liquidacionFacturaController.objdataset.Tables(0)

                For Each fila As DataRow In dataTable.Rows
                    Dim bruto As Integer = Convert.ToInt32(fila.Item(19))
                    Dim factura As Integer = Convert.ToInt32(fila.Item(1))
                    Dim numeroRem As Integer = Convert.ToInt32(fila.Item(22))
                    Dim rutCliente As String = fila.Item(4).ToString()
                    Dim fechaEmision As String = fila.Item(2).ToString()
                    Dim fechaVencimiento As String = fila.Item(3).ToString()
                    Dim comision As Integer = Convert.ToInt32(fila.Item(15))
                    Dim otros As Integer = Convert.ToInt32(fila.Item(17)) + Convert.ToInt32(fila.Item(20))
                    Dim nombre As String = fila.Item(5).ToString()
                    Dim neto As Integer = Convert.ToInt32(fila.Item(11))
                    Dim iva As Integer = Convert.ToInt32(fila.Item(12))
                    Dim ivaComision As Integer = Convert.ToInt32(fila.Item(18))
                    Dim administrativos As Integer = Convert.ToInt32(fila.Item(16))
                    Dim descuentos As Integer = Convert.ToInt32(fila.Item(15)) + Convert.ToInt32(fila.Item(16)) + Convert.ToInt32(fila.Item(17)) + Convert.ToInt32(fila.Item(18))
                    Dim vet As Integer = Convert.ToInt32(fila.Item(17))
                    Dim decomiso As Integer = Convert.ToInt32(fila.Item(20))

                    liquidacionFacturaController.generar_csv_softland(bruto, factura, num_remate, rutCliente, fechaEmision, fechaVencimiento, comision, otros, factura, nombre, neto, iva, ivaComision, administrativos, descuentos, vet, decomiso)
                Next
            Catch ex As Exception
                Me.Invoke(Sub() MessageBox.Show("Error procesando liquidación: " & ex.ToString))
            End Try
        End If
    End Sub

    Private Sub consultarIngresos(ByVal parametro As String, ByRef id As String)
        Dim ingresoController As New ingresoController

        If ingresoController.getAllParam(parametro, id) Then
            Try
                Dim dataTable As DataTable = ingresoController.objdataset.Tables(0)

                For Each fila As DataRow In dataTable.Rows
                    Dim idIngreso As String = fila(0).ToString
                    Dim count As Integer = 0

                    If fila(13).ToString() = "Efectivo" Or fila(13).ToString() = "Tranferencia" Or fila(13).ToString() = "Pago en Venta" Then
                        ingresoController.GuardarCSVefectivo(fila(0).ToString(), fila(1).ToString())

                    End If

                    If fila(13).ToString() = "Cheque a la Fecha" Then
                        Dim facturaElectronicaController As New facturaElectronicaController
                        Dim n_ingreso As String = "where n_ingreso = '" & idIngreso & "'"

                        If facturaElectronicaController.getAllParam(n_ingreso) Then
                            For Each row As DataRow In facturaElectronicaController.objdataset.Tables(0).Rows
                                count = count + 1
                            Next
                        End If

                        If count >= 2 Then
                            ingresoController.GuardarCSVchequeFechaMasDeUnaFactura(fila(0).ToString, fila(1).ToString())
                        Else
                            ingresoController.GuardarCSVchequeFecha(fila(0).ToString, fila(1).ToString())
                        End If

                    End If

                    If fila(13).ToString = "Cheque al dia" Then
                        ingresoController.GuardarCSVchequeDia(fila(0).ToString, fila(1).ToString)

                    End If

                    If fila(13).ToString = "Transferencia" Then
                        ingresoController.GuardarCSVtransferencia(fila(0).ToString, fila(1).ToString)

                    End If

                    If fila(13).ToString = "Pago con venta" Then
                        If ingresoController.getIniciarCSV(idIngreso) Then
                            ingresoController.GuardarCSVPagoVenta(fila(0).ToString, fila(1).ToString)

                        Else
                            ingresoController.updateCSVStatus(idIngreso)

                        End If
                    End If

                    ingresoController.updateCSVStatus(idIngreso)

                Next
            Catch ex As Exception
                Me.Invoke(Sub() MessageBox.Show("Error procesando ingreso: " & ex.Message))
            End Try
        End If
    End Sub

    Private Sub consultarEgresos(ByVal parametro As String, ByRef id As String)
        Dim egresoController As New egresoController

        If egresoController.getAllParam(parametro) Then
            Try
                Dim dataTable As DataTable = egresoController.objdataset.Tables(0)

                For Each fila As DataRow In dataTable.Rows
                    Dim idEgreso As String = fila(0).ToString
                    Dim count As Integer = 0

                    If fila(8).ToString = "Cheque al Día" Then
                        Try
                            Dim numeroEgreso As Integer
                            If Integer.TryParse(fila(1).ToString(), numeroEgreso) Then
                                ' Verificar si existe un anticipo previo para el rut
                                Dim rutCliente As String = fila(4).ToString()
                                If egresoController.verificarAnticipoRut(rutCliente) Then
                                    ' Obtener el idEgreso del anticipo
                                    Dim idEgresoAnticipo As String = egresoController.Egreso.idEgreso
                                    If Not String.IsNullOrEmpty(idEgresoAnticipo) Then
                                        '     MsgBox("Procesando cheque al día con anticipo. ID Egreso: " & fila(0).ToString() & ", ID Anticipo: " & idEgresoAnticipo)
                                        ' Si tiene anticipo, llamar al método con los 3 parámetros
                                        egresoController.GuardarCSVchequeAlDiaConAnticipo(fila(0).ToString(), idEgresoAnticipo, numeroEgreso)
                                    Else
                                        '   MsgBox("Se encontró anticipo pero no se pudo obtener su ID para el RUT: " & rutCliente)
                                        ' Si no se pudo obtener el ID del anticipo, procesar como cheque normal
                                        egresoController.GuardarCSVchequeAlDia(fila(0).ToString(), numeroEgreso)
                                    End If
                                Else
                                    ' Si no tiene anticipo, llamar al método original
                                    egresoController.GuardarCSVchequeAlDia(fila(0).ToString(), numeroEgreso)
                                End If
                            Else
                                Throw New Exception("El número de egreso no es válido: " & fila(1).ToString())
                            End If
                        Catch ex As Exception
                            Me.Invoke(Sub() MessageBox.Show("Error procesando egreso: " & ex.Message))
                        End Try

                    ElseIf fila(8).ToString = "Cheque a Fecha" Then
                        egresoController.GuardarCSVchequeAlaFecha(fila(0).ToString, fila(1).ToString)
                    Else
                        egresoController.GuardarCSVanticipo(fila(0).ToString, fila(1).ToString)
                    End If
                Next
            Catch ex As Exception
                Me.Invoke(Sub() MessageBox.Show("Error procesando egreso: " & ex.Message))
            End Try
        End If
    End Sub

    Private Sub consultarEgresosCompras(ByVal parametro As String, ByRef id As String)
        Dim egresoController As New egresoController

        If egresoController.getAllParam(parametro) Then
            Try
                Dim dataTable As DataTable = egresoController.objdataset.Tables(0)

                For Each fila As DataRow In dataTable.Rows
                    Dim idEgreso As String = fila(0).ToString
                    Dim count As Integer = 0

                    egresoController.GuardarCSVanticipoPagoVenta(fila(0).ToString, fila(1).ToString)
                Next
            Catch ex As Exception
                Me.Invoke(Sub() MessageBox.Show("Error procesando egreso: " & ex.Message))
            End Try
        End If
    End Sub

    Protected Overrides Sub OnFormClosing(ByVal e As FormClosingEventArgs)
        Timer1.Enabled = False
        MyBase.OnFormClosing(e)
    End Sub
End Class