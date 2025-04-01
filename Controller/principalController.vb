Imports System.Data.OleDb

Public Class principalController
    Public sql As String
    Public objadapter As OleDbDataAdapter
    Public objdataset As DataSet
    Public Function cargar_remate() As Boolean
        Try
            If ConectaBase() Then

                Dim sql As String
                Dim objadapter As OleDbDataAdapter
                Dim objdataset As DataSet

                sql = "SELECT id,remate,FORMAT(fecha, 'yyyy-MM-dd'), idRecinto FROM remate where estado= 1 "
                objadapter = New OleDbDataAdapter(sql, Conecta)
                objdataset = New DataSet
                objadapter.Fill(objdataset)

                id_remate = objdataset.Tables(0).Rows(0).Item(0)
                num_remate = objdataset.Tables(0).Rows(0).Item(1)
                fecha_remate = objdataset.Tables(0).Rows(0).Item(2)
                idRecinto = objdataset.Tables(0).Rows(0).Item(3)

                'id_remate = 42
                'num_remate = 13
                'fecha_remate = '2025-02-15'
                'idRecinto = 1

            Else
                MsgBox("¡No se han encontrado Registros!")
            End If

        Catch ex As Exception
            MsgBox("Debe iniciar Remate")
            Return False

        End Try
        Return True
    End Function

    Public Function estadoDocumentos(ByRef idRemate As Integer) As Boolean
        Try
            If ConectaBase() Then



                ' Construcción de la consulta SQL utilizando el valor de idRemate
                sql = $"SELECT 'Factura Electrónica' AS Tipo,
                           COUNT(*) AS Total,
                           SUM(CASE WHEN ISNULL(csv, 0) = 1 THEN 1 ELSE 0 END) AS Procesados,
                           COUNT(*) - SUM(CASE WHEN ISNULL(csv, 0) = 1 THEN 1 ELSE 0 END) AS Pendientes
                    FROM FacturaElectronica
                    WHERE idRemate = {idRemate}

                    UNION ALL

                    SELECT 'Liquidación Factura' AS Tipo,
                           COUNT(*) AS Total,
                           SUM(CASE WHEN ISNULL(csv, 0) = 1 THEN 1 ELSE 0 END) AS Procesados,
                           COUNT(*) - SUM(CASE WHEN ISNULL(csv, 0) = 1 THEN 1 ELSE 0 END) AS Pendientes
                    FROM LiquidacionFacturaElectronica
                    WHERE idRemate = {idRemate}

                    UNION ALL

                    SELECT 'Factura Compra' AS Tipo,
                           COUNT(*) AS Total,
                           SUM(CASE WHEN ISNULL(csv, 0) = 1 THEN 1 ELSE 0 END) AS Procesados,
                           COUNT(*) - SUM(CASE WHEN ISNULL(csv, 0) = 1 THEN 1 ELSE 0 END) AS Pendientes
                    FROM FacturaCompra
                    WHERE idRemate = {idRemate}

                    UNION ALL

                    SELECT 'Nota de Crédito' AS Tipo,
                           COUNT(*) AS Total,
                           SUM(CASE WHEN ISNULL(csv, 0) = 1 THEN 1 ELSE 0 END) AS Procesados,
                           COUNT(*) - SUM(CASE WHEN ISNULL(csv, 0) = 1 THEN 1 ELSE 0 END) AS Pendientes
                    FROM NotaCredito
                    WHERE idRemate = {idRemate}

                    UNION ALL

                    SELECT 'Egresos' AS Tipo,
                           COUNT(*) AS Total,
                           SUM(CASE WHEN ISNULL(csv, 0) = 1 THEN 1 ELSE 0 END) AS Procesados,
                           COUNT(*) - SUM(CASE WHEN ISNULL(csv, 0) = 1 THEN 1 ELSE 0 END) AS Pendientes
                    FROM Egresos
                    WHERE idRemate = {idRemate} and anulado = 0

                    UNION ALL

                    SELECT 'Ingresos' AS Tipo,
                           COUNT(*) AS Total,
                           SUM(CASE WHEN ISNULL(csv, 0) = 1 THEN 1 ELSE 0 END) AS Procesados,
                           COUNT(*) - SUM(CASE WHEN ISNULL(csv, 0) = 1 THEN 1 ELSE 0 END) AS Pendientes
                    FROM Ingresos
                    WHERE idRemate = {idRemate} and anulado = 0  ;"


                ' Ejecución de la consulta
                objadapter = New OleDbDataAdapter(Sql, Conecta)
                objdataset = New DataSet
                objadapter.Fill(objdataset)

            Else
                MsgBox("¡No se han encontrado Registros!")
            End If

        Catch ex As Exception
            MsgBox("Debe iniciar Remate")
            Return False
        End Try
        Return True
    End Function

End Class
