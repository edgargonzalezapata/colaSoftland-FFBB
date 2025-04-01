Imports System.Data.OleDb
Imports System.Data.Sql
Imports System.Data.SqlClient
Public Class cliente




    Public rut As String
    Public nombres As String
    Public apellidos As String
    Public direccion As String
    Public direccion2 As String
    Public fono As String
    Public email As String
    Public ciudad As String
    Public comuna As String
    Public giro As String
    Public com_venta As Decimal
    Public condi_venta As String
    Public com_compra As Decimal
    Public condi_compra As String
    Public tipo_retencion As String
    Public corredor_compras As String
    Public nom_corredor_compras As String
    Public check_compras As Boolean
    Public porcentaje_compras As Double
    Public corredor_ventas As String
    Public nom_corredor_ventas As String
    Public check_ventas As Boolean
    Public porcentaje_ventas As Double
    Public activo As String
    Public diferencial As String
    Public asegurado As String
    Public tipo_seg As String
    Public monto_uf As Double
    Public vigencia As String
    Public est_seguro As String
    Public fecha As String
    Public tipo As Integer
    Public id As Integer
    Public comision_v As String
    Public condicion_v As String
    Public monto_rem As Integer
    Public estado_seg As String




    Public comision As Decimal
    Public condicion As String
    Public retencion As Integer


    Public conexiones As SqlConnection
    Public enunciado As SqlCommand
    Public respuesta As SqlDataReader
    Public adaptador As SqlDataAdapter

    Public sql As String
    Public objadapter As OleDbDataAdapter
    Public objdataset As DataSet
    Public condicionS As String



    Public Function existe2() As Boolean

        Try

            conexiones = New SqlConnection(conexion_string)
            conexiones.Open()


            enunciado = New SqlCommand("select   id,rut,nombres,direccion,ciudad,comuna,giro  from clientes where rut  = '" & rut & "'", conexiones)
            respuesta = enunciado.ExecuteReader()

            Try
                While respuesta.Read

                    id = respuesta.GetInt32(0)
                    nombres = respuesta.GetString(2)
                    ' apellidos = respuesta.GetString(3)
                    direccion = respuesta.GetString(3)
                    ciudad = respuesta.GetString(4)
                    comuna = respuesta.GetString(5)
                    giro = respuesta.GetString(6)
                    Return True
                End While

            Finally

                respuesta.Close()
                conexiones.Close()
            End Try


        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return False
    End Function

    Public Function getRutCliente(ByRef cliente As String) As Boolean

        Try
            If ConectaBase() Then

                sql = "SELECT top 1 c.rut, c.nombres, c.comision, c.condicion, c.retencion, co.condicion  
FROM clientes as c inner join condicion as co on c.condicion = co.tipo  
where rut = '" & cliente & "' OR  c.nombres like '%" & cliente & "%' "
                objadapter = New OleDbDataAdapter(sql, Conecta)
                objdataset = New DataSet
                objadapter.Fill(objdataset)

                If objdataset.Tables(0).Rows.Count > 0 Then

                    rut = objdataset.Tables(0).Rows(0).Item(0)
                    nombres = objdataset.Tables(0).Rows(0).Item(1)
                    comision = objdataset.Tables(0).Rows(0).Item(2)
                    condicion = objdataset.Tables(0).Rows(0).Item(3)
                    retencion = objdataset.Tables(0).Rows(0).Item(4)
                    condicionS = objdataset.Tables(0).Rows(0).Item(5)

                    Return True
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.ToString)

        End Try
        Return False
    End Function


    Sub conexion()
        conexiones = New SqlConnection(conexion_string)
        conexiones.Open()
    End Sub
    Public Function existe() As Boolean

        Try
            conexion()


            enunciado = New SqlCommand("select   nombres,direccion,Fono,Email,ciudad,comuna,giro,isnull(comision_v,0), condicionv, isnull(comision,0), condicion, retencion, isnull(corredor_compras,0),isnull(porcentaje_compras,0), isnull(corredor_ventas,0), isnull(porcentaje_ventas,0), activo,isnull(diferencial,''), asegurado, isnull(Tipo_seg,0),isnull(montouf,0),isnull(vigencia,'1999-01-01') ,estado_seg,id,isnull(nom_corredor_compras,''), isnull(nom_corredor_ventas,'') ,isnull(direccion2,'') from clientes where rut='" & rut & "'", conexiones)
            respuesta = enunciado.ExecuteReader()

            Try
                While respuesta.Read

                    nombres = respuesta.GetString(0)
                    direccion = respuesta.GetString(1)
                    fono = respuesta.GetString(2)
                    email = respuesta.GetString(3)
                    ciudad = respuesta.GetString(4)
                    comuna = respuesta.GetString(5)
                    giro = respuesta(6)
                    com_venta = respuesta(7)
                    condi_venta = respuesta(8)
                    com_compra = respuesta(9)
                    condi_compra = respuesta(10)
                    tipo_retencion = respuesta(11)
                    corredor_compras = respuesta(12)
                    porcentaje_compras = respuesta(13)
                    corredor_ventas = respuesta(14)
                    porcentaje_ventas = respuesta(15)
                    activo = respuesta(16)
                    diferencial = respuesta(17)
                    asegurado = respuesta(18)
                    tipo_seg = respuesta(19)
                    monto_uf = respuesta(20)
                    vigencia = respuesta(21)
                    est_seguro = respuesta(22)
                    id = respuesta(23)
                    nom_corredor_compras = respuesta(24)
                    nom_corredor_ventas = respuesta(25)
                    direccion2 = respuesta.GetString(26)
                    Return True
                End While

            Finally

                respuesta.Close()
                conexiones.Close()
            End Try


        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return False
    End Function

    Public Function guardar_cliente() As Boolean
        conexion()




        Try
            enunciado = New SqlCommand("insert into clientes( rut,
nombres,
direccion,
direccion2,
Fono,
Email,
ciudad,
comuna,
giro,
comision_v, 
condicionv, 
comision, 
condicion, 
retencion,
corredor_compras,
porcentaje_compras,
corredor_ventas,
porcentaje_ventas, 
activo,
diferencial, 
asegurado,
Tipo_seg,
montouf,
vigencia ,
estado_seg) values ('" & rut & "',
'" & nombres & "',
'" & direccion & "', 
'" & direccion2 & "', 
'" & fono & "', 
'" & email & "', 
'" & ciudad & "', 
'" & comuna & "', 
'" & giro & "', 
" & Replace(com_venta, ",", ".") & ", 
" & condi_venta & ", 
" & Replace(com_compra, ",", ".") & ", 
" & condi_compra & ", 
" & tipo_retencion & ", 
'" & corredor_compras & "',
" & porcentaje_compras & ", 
'" & corredor_ventas & "',
" & porcentaje_ventas & ",
'" & activo & "', 
" & diferencial & ", 
'" & asegurado & "', 
'" & tipo_seg & "',
" & monto_uf & ", 
'" & vigencia & "' ,
'" & est_seguro & "' )", conexiones)
            respuesta = enunciado.ExecuteReader()
            Return True
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False
        Finally
            respuesta.Close()
            conexiones.Close()
        End Try
        Return False
    End Function


    Public Function update_cliente() As Boolean
        conexion()
        Dim com_ventas, com_compras As String
        com_ventas = Replace(com_venta, ",", ".")
        com_compras = Replace(com_compra, ",", ".")

        Try
            enunciado = New SqlCommand("update  clientes set  rut = '" & rut & "',
nombres = '" & nombres & "',
direccion = '" & direccion & "',
direccion2 = '" & direccion2 & "',
Fono = '" & fono & "',
Email = '" & email & "',
ciudad = '" & ciudad & "',
comuna ='" & comuna & "',
giro = '" & giro & "' ,
comision_v = " & com_ventas & ", 
condicionv =" & condi_venta & ", 
comision =" & com_compras & ", 
condicion =" & condi_compra & ", 
retencion = " & tipo_retencion & ",
corredor_compras = '" & corredor_compras & "',
nom_corredor_compras = '" & nom_corredor_compras & "',
porcentaje_compras =" & porcentaje_compras & ",
corredor_ventas =  '" & corredor_ventas & "',
nom_corredor_ventas =  '" & nom_corredor_ventas & "',
porcentaje_ventas = " & porcentaje_ventas & ", 
activo = '" & activo & "',
diferencial = " & diferencial & ", 
asegurado = '" & asegurado & "',
Tipo_seg = '" & tipo_seg & "',
montouf = " & monto_uf & ",
vigencia  = '" & vigencia & "' ,
estado_seg = '" & est_seguro & "' where rut = '" & rut & "'", conexiones)
            respuesta = enunciado.ExecuteReader()
            Return True
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False
        Finally
            respuesta.Close()
            conexiones.Close()
        End Try
        Return False
    End Function


    Public Function InsertarComision() As Boolean
        conexion()




        Try
            enunciado = New SqlCommand("insert into comisionCliente(comision,condicion,rut_cliente,tipo ) values ('" & Replace(com_venta, ",", ".") & "'," & condi_venta & ", '" & rut & "', " & tipo & ")", conexiones)
            respuesta = enunciado.ExecuteReader()
            Return True
        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False
        Finally
            respuesta.Close()
            conexiones.Close()
        End Try
        Return False
    End Function

End Class
