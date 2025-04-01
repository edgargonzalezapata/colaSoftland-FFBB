Imports System.Data.OleDb
Imports System.Data.SqlClient

Public Class remateController
    Public remate As New remate
    Public recientoFerial As New recientoFerial
    Public sql As String
    Public objadapter As OleDbDataAdapter
    Public objdataset As DataSet
    Public conexiones As SqlConnection

    Sub conexion()
        conexiones = New SqlConnection(conexion_string)
        conexiones.Open()
    End Sub
    Public Function getRemateActual() As Boolean

        Try
            If ConectaBase() Then


                sql = "select r.fecha,f.nombre as Recinto, f.rup, r.remate , r.correlativoSoftland
                        from remate as r 
                        inner JOIN recintoFerial as f on f.idRecinto = r.idRecinto
                        where r.estado = 1"
                objadapter = New OleDbDataAdapter(sql, Conecta)
                objdataset = New DataSet
                objadapter.Fill(objdataset)

                If objdataset.Tables(0).Rows.Count > 0 Then

                    remate.fecha = objdataset.Tables(0).Rows(0).Item(0)
                    recientoFerial.nombre = objdataset.Tables(0).Rows(0).Item(1)
                    recientoFerial.rup = objdataset.Tables(0).Rows(0).Item(2)
                    remate.remate = objdataset.Tables(0).Rows(0).Item(3)
                    remate.correlativoSoftland = objdataset.Tables(0).Rows(0).Item(4)
                    Return True
                End If

            End If
        Catch ex As Exception

        End Try
        Return False
    End Function

    Public Function correlativoSoftland1() As Boolean
        conexion()

        Try
            enunciado = New SqlCommand("update remate set correlativoSoftland = correlativoSoftland +1 where id = " & id_remate & ";select max(correlativoSoftland) as correlativoSoftland from remate where id = " & id_remate & "", conexiones)
            respuesta = enunciado.ExecuteReader()

            Try
                While respuesta.Read

                    remate.correlativoSoftland = respuesta(0)

                End While
                Return True
            Finally

                respuesta.Close()
                conexiones.Close()
            End Try
        Catch
            Return False
        Finally
            respuesta.Close()
            conexiones.Close()
        End Try
    End Function

End Class
