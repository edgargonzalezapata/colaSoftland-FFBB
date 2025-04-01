Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net.WebRequestMethods

Public Class ingresoController

    Public sql As String
    Public objadapter As OleDbDataAdapter
    Public objdataset As DataSet
    Public conexiones As SqlConnection

    Dim fFile As IO.File, fStream As IO.FileStream

    Public Function getAllParam(ByRef parametro As String, ByRef id As String) As Boolean
        Try
            If ConectaBase() Then

                sql = "select 
    i.idIngreso,
    i.numero as Numero,  
    FORMAT(i.fecha, 'dd-MM-yyyy') as Fecha, 
    FORMAT(i.fecha, 'HH:mm') as Hora, 
    i.rutCliente as RUT,
    c.nombres as Cliente,
    i.total as Total,
    i.motivo as Motivo,
    i.vb as VB,
    r.remate as Remate,
    i.anulado as Anulado,
    i.manual as Manual,
    i.contable as Contable,
    (select top 1 tp.descripcion 
     from tipoPago as tp 
     inner join IngresosDet as d on tp.cuentaIngreso = d.codigo 
     where d.idIngreso = i.idIngreso 
     and tp.descripcion <> 'Pago Clientes' 
     order by case when tp.descripcion = 'Pago con venta' then 1 else 2 end) as [Tipo Pago]
from 
    Ingresos as i 
inner join 
    clientes as c on i.rutCliente = c.rut
inner join 
    remate as r on r.id = i.idRemate
" & parametro & " group by 
    i.idIngreso, i.numero, i.fecha, i.rutCliente, c.nombres, i.total, i.motivo, 
    i.vb, i.idRemate, r.remate, i.anulado, i.manual, i.contable;"

                objadapter = New OleDbDataAdapter(sql, Conecta)
                objdataset = New DataSet
                objadapter.Fill(objdataset)

                If objdataset.Tables(0).Rows.Count > 0 Then
                    Return True
                Else
                    updateCSVStatus(id)
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return False

    End Function

    Public Function getIniciarCSV(ByRef idIngreso As String) As Boolean
        Try
            If ConectaBase() Then
                ' Definir la consulta SQL
                Dim sql As String = "select count(codigo) as totalMedios from IngresosDet " &
                                "where idIngreso = '" & idIngreso & "' and debito <> 0"

                ' Configurar el adaptador y el dataset
                objadapter = New OleDbDataAdapter(sql, Conecta)
                objdataset = New DataSet
                objadapter.Fill(objdataset)

                ' Verificar si el conteo es mayor o igual a 2
                If objdataset.Tables(0).Rows.Count > 0 AndAlso
               Convert.ToInt32(objdataset.Tables(0).Rows(0)("totalMedios")) >= 2 Then
                    Return True
                End If
            End If
        Catch ex As Exception
            ' Manejar excepciones
            MsgBox(ex.ToString)
        End Try
        Return False
    End Function


    Public Function updateCSVStatus(ByVal idIngreso As String) As Boolean
        Try
            Using conexiones As New SqlConnection(conexion_string)
                conexiones.Open()

                Dim consulta As String = "UPDATE Ingresos SET " &
                                    "csv = 1 WHERE idIngreso = @idIngreso"

                Using enunciado As New SqlCommand(consulta, conexiones)
                    ' Asignar valores a los parámetros
                    enunciado.Parameters.AddWithValue("@idIngreso", idIngreso)

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

    Public Function GuardarCSVefectivo(ByRef idIngreso As String, ByRef numero As Integer) As Boolean
        Try
            If ConectaBase() Then
                Dim sql As String = "WITH cte AS (
  SELECT d.cuentaSoft,
         d.debito AS Debe,
         d.credito AS Haber,
         i.idIngreso
  FROM Ingresos AS i
  FULL JOIN IngresosDet AS d ON i.idIngreso = d.idIngreso
  WHERE d.cuentaSoft IN ('2-1-3-01-01', '1-1-3-01-03')
)
SELECT '001' AS [Area de Negocio],
       d.cuentaSoft AS [Codigo Plan de cuenta],
       d.debito AS [Debe],
       CASE
           WHEN d.cuentaSoft = '1-1-3-01-03' 
           THEN d.credito - (SELECT ISNULL(SUM(Debe), 0) FROM cte WHERE cuentaSoft = '2-1-3-01-01' AND idIngreso = i.idIngreso)
           ELSE d.credito
       END AS [Haber],
       CASE
           WHEN tp.cuentaIngreso = '1-1-3-01-03' THEN CONCAT('F.V.- PAGO REM.:', r.remate, ' PAGO CLIENT')
           WHEN tp.cuentaIngreso = '2-1-3-01-01' THEN ('F.C. - PAGO FACTURA CON VENTA')
           ELSE CONCAT('CANCELACION FACTURAS REM.:', r.remate)
       END AS [Descripción],
       '0' AS EquivalenciaMoneda,
       '0' AS [Debe al debe Moneda],
       '' AS [Haber al haber Moneda],
       '' AS [Código Condición de Venta],
       '' AS [Código Vendedor],
       '' AS [Código Ubicación],
       '' AS [Código Concepto de Caja],
       '' AS [Código Instrumento Financiero],
       '' AS [Cantidad Instrumento Financiero],
       '' AS [Código Detalle de Gasto],
       '' AS [Cantidad Concepto de Gasto],
       '' AS [Código Centro de Costo],
       '' AS [Código Centro de Costo 2],
       CASE
           WHEN c.idCuenta = '1-1-3-01-03' THEN '0'
           ELSE '0'
       END AS [Nro. Docto. Conciliación],
       CASE
           WHEN tp.cuentaSoft = '1-1-3-01-03' THEN LEFT(i.rutCliente, LEN(i.rutCliente) - 2)
           ELSE ''
       END AS [Codigo Auxiliar],
       CASE
           WHEN tp.cuentaSoft = '1-1-1-01-01' THEN ''
           WHEN tp.cuentaSoft = '1-1-3-01-01' THEN 'PB'
           WHEN tp.cuentaSoft = '1-1-3-01-03' THEN 'PA'
           ELSE ''
       END AS [Tipo Documento],
       CASE
           WHEN tp.cuentaSoft = '1-1-1-01-01' THEN ''
           WHEN tp.cuentaSoft = '1-1-3-01-03' THEN CAST(r.remate AS NVARCHAR)
           ELSE ''
       END AS [Nro. Documento],
       CONVERT(NVARCHAR(10), i.fecha, 105) AS FechaEmision,
       CONVERT(NVARCHAR(10), i.fecha, 105) AS FechaVencimiento,
       CASE
           WHEN tp.cuentaSoft = '1-1-1-01-01' THEN ''
           WHEN tp.cuentaSoft = '1-1-3-01-01' THEN 'PB'
           WHEN tp.cuentaSoft = '1-1-3-01-03' THEN 'VB'
           ELSE ''
       END AS [Tipo Docto. Referencia],
       CASE
           WHEN tp.cuentaSoft = '1-1-3-01-03' THEN CAST(f.idDTE AS NVARCHAR)
           ELSE ''
       END AS [Nro. Docto. Referencia],
       '' AS [Correlativo Documento de Compra Interno],
       '0' AS [Monto 1],
       '0' AS [Monto 2],
       '0' AS [Monto 3],
       '0' AS [Monto 4],
       '0' AS [Monto 5],
       '0' AS [Monto 6],
       '0' AS [Monto 7],
       '0' AS [Monto 8],
       '0' AS [Monto 9],
       '0' AS [Monto Suma Detalle Libro],
       '' AS [Número Documento Desde],
       '' AS [Número Documento Hasta],
       CAST(i.numero AS NVARCHAR) AS [Nro. agrupación en igual comprobante],
       'N' AS [NºOpe.Doc.Aux.(Considera este campo sólo si se definió en los parámetros de CW)],
       'I' AS [Graba el detalle de libro (S/N)],
       'I' AS [Documento Nulo (S/N)],
       CONCAT('INGRESO REMATE ', r.remate, ' NRO. ', CAST(i.numero AS NVARCHAR)) AS [Tipo de comprobante (I, E, T)],
       '' AS [Descripción del comprobante]
FROM Ingresos AS i
FULL JOIN IngresosDet AS d ON i.idIngreso = d.idIngreso
FULL JOIN tipoPago AS tp ON tp.cuentaIngreso = d.codigo
FULL JOIN Cuentas AS c ON c.idCuenta = d.codigo
FULL JOIN remate AS r ON r.id = i.idRemate
FULL JOIN FacturaElectronica AS f ON f.n_ingreso = i.idIngreso
WHERE d.cuentaSoft != '2-1-3-01-01' 
AND i.idIngreso =  '" & idIngreso & "'
ORDER BY d.debito DESC;"


                Dim objadapter As New OleDbDataAdapter(sql, Conecta)
                Dim objdataset As New DataSet
                objadapter.Fill(objdataset)

                If objdataset.Tables(0).Rows.Count > 0 Then
                    ' Ruta donde se guardará el archivo CSV
                    Dim destino As String
                    Dim usu As Microsoft.VisualBasic.ApplicationServices.User
                    usu = New Microsoft.VisualBasic.ApplicationServices.User

                    Dim usuario() As String = Split(usu.Name, "\")

                    Dim fechaRemate As Date = fecha_remate
                    Dim sFilename As String = "S:\fbiobio" & fechaRemate.ToString("yyyyMMdd") & ".CSV"

                    destino = sFilename

                    ' Crear un diccionario para rastrear filas únicas por código de cuenta
                    Dim uniqueRows As New Dictionary(Of String, DataRow)

                    ' Filtrar filas duplicadas basadas en [Codigo Plan de cuenta]
                    For Each row As DataRow In objdataset.Tables(0).Rows
                        Dim codigoCuenta As String = row("Codigo Plan de cuenta").ToString()
                        If Not uniqueRows.ContainsKey(codigoCuenta) Then
                            uniqueRows.Add(codigoCuenta, row)
                        End If
                    Next

                    ' Crear un StreamWriter para escribir en el archivo CSV


                    Dim retries As Integer = 0
                    Dim maxRetries As Integer = 5
                    Dim success As Boolean = False

                    ' Intentar escribir el archivo con manejo de errores
                    Do While Not success And retries < maxRetries
                        Try


                            If fFile.Exists(sFilename) Then

                            Else
                                Using writer As New StreamWriter(destino)
                                    ' Escribir solo las filas únicas
                                    For Each row As DataRow In uniqueRows.Values
                                        Dim rowData As String = String.Join(",", row.ItemArray.Select(Function(value) value.ToString()))
                                        writer.WriteLine(rowData)
                                    Next
                                    writer.Close()
                                End Using

                                updateCSVStatus(idIngreso)
                                Return True
                            End If

                        Catch ex As Exception
                            ' MsgBox(ex.ToString)
                        End Try
                    Loop


                    Return True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return False
    End Function

    Public Function GuardarCSVchequeFecha(ByRef idIngreso As String, ByRef numero As Integer) As Boolean
        Try
            If ConectaBase() Then
                Dim sql As String = "WITH CTE AS (
    SELECT '001' AS [Area de Negocio],
           tp.cuentaSoft AS [Codigo Plan de cuenta],
           d.debito AS [Debe],
           d.credito AS [Haber],
           CASE
               WHEN  tp.cuentaSoft = '1-1-3-01-03' THEN concat('F.V.- PAGO REM.:', r.remate ,' PAGO CLIENT')
               WHEN  tp.cuentaSoft = '2-1-3-01-01' THEN ('F.C. -PAGO FACTURA CON VENTA')
               ELSE concat('CANCELACION FACTURAS REM.:', r.remate)
           END as [Descripción],
           '0' as EquivalenciaMoneda,
           '0' as [Debe al debe Moneda],
           '' as [Haber al haber Moneda],
           '' as [Código Condición de Venta],
           '' as [Código Vendedor],
           '' as [Código Ubicación],
           '' as [Código Concepto de Caja],
           '' as [Código Instrumento Financiero],
           '' as [Cantidad Instrumento Financiero],
           '' as [Código Detalle de Gasto],
           '' as [Cantidad Concepto de Gasto],
           '' as [Código Centro de Costo],
           '' as [Tipo Docto. Conciliación],
           CASE
               WHEN  tp.cuentaSoft = '1-1-3-01-03' THEN '0'
               ELSE '0'
           END AS [Nro. Docto. Conciliación],
           CASE
               WHEN  tp.cuentaSoft = '1-1-1-01-01' THEN LEFT(i.rutCliente, LEN(i.rutCliente) - 2)
               WHEN  tp.cuentaSoft = '1-1-3-01-01' THEN LEFT(i.rutCliente, LEN(i.rutCliente) - 2)
               WHEN  tp.cuentaSoft = '1-1-3-01-03' THEN LEFT(i.rutCliente, LEN(i.rutCliente) - 2)
               ELSE ''
           END AS [Codigo Auxiliar],
           'PB' AS [Tipo Documento],
           (select top 1 idc.numeroCheq from ingresosDetCheque as idc where (idc.idIngresoDet = d.idIngresoDet )) AS [Nro. Documento],
           CONVERT(NVARCHAR(10), i.fecha, 105) as FechaEmision,
           (select top 1 idc.fechaDeposito from ingresosDetCheque as idc where (idc.idIngresoDet = d.idIngresoDet )) as FechaVencimiento,
           
           CASE
               WHEN  tp.cuentaSoft = '1-1-3-01-03' THEN 'VB'
               ELSE 'PB'		  
           END AS [Tipo Docto. Referencia],
           
           CASE
               WHEN  tp.cuentaSoft = '1-1-3-01-03' THEN f.idDTE
               ELSE (select top 1 idc.numeroCheq from ingresosDetCheque as idc where (idc.idIngresoDet = d.idIngresoDet ))		  
           END AS [Nro. Docto. Referencia],
           '' as [Correlativo Documento de Compra Interno],
           '0' as [Monto 1],
           '0' as [Monto 2],
           '0' as [Monto 3],
           '0' as [Monto 4],
           '0' as [Monto 5],
           '0' as [Monto 6],
           '0' as [Monto 7],
           '0' as [Monto 8],
           '0' as [Monto 9],
           '0' as [Monto Suma Detalle Libro],
           '' as [Número Documento Desde],
           '' as [Número Documento Hasta],
           i.numero as [Nro. agrupación en igual comprobante],
           'N' as [NºOpe.Doc.Aux.(Considera este campo sólo si se definió en los parametros de CW)],
           'I' as [Graba el detalle de libro (S/N) ],
           'I' as [Documento Nulo (S/N)],
           concat('INGRESO REMATE ',r.remate ,' NRO. ', I.numero )  as [Tipo de comprobante (I, E, T)],
           '' AS [Descripción del comprobante],
           ROW_NUMBER() OVER (ORDER BY d.debito DESC) as RowNum
    FROM Ingresos AS i
    INNER JOIN IngresosDet AS d ON i.idIngreso = d.idIngreso
    FULL JOIN tipoPago as tp on tp.cuentaIngreso = d.codigo
    FULL JOIN Cuentas as c on c.idCuenta = d.codigo
    FULL JOIN remate as r on r.id = i.idRemate
    FULL JOIN FacturaElectronica as f on f.n_ingreso = i.idIngreso
    WHERE i.idIngreso = '" & idIngreso & "'
),
DebitoNull AS (
    -- Subconsulta para obtener el valor del Debe cuando tp.cuentaSoft es NULL
    SELECT SUM([Debe]) AS TotalDebeNull
    FROM CTE
    WHERE [Codigo Plan de cuenta] IS NULL
)
SELECT 
    [Area de Negocio],
    [Codigo Plan de cuenta],
    CASE 
        WHEN [Codigo Plan de cuenta] IS NOT NULL THEN [Debe]
        ELSE NULL 
    END AS [Debe],
    CASE 
        WHEN [Codigo Plan de cuenta] = '1-1-3-01-03' THEN 
            COALESCE([Haber], 0) - COALESCE((SELECT TotalDebeNull FROM DebitoNull), 0)
        ELSE COALESCE([Haber], 0)
    END AS [Haber],
    [Descripción],
    EquivalenciaMoneda,
    [Debe al debe Moneda],
    [Haber al haber Moneda],
    [Código Condición de Venta],
    [Código Vendedor],
    [Código Ubicación],
    [Código Concepto de Caja],
    [Código Instrumento Financiero],
    [Cantidad Instrumento Financiero],
    [Código Detalle de Gasto],
    [Cantidad Concepto de Gasto],
    [Código Centro de Costo],
    [Tipo Docto. Conciliación],
    [Nro. Docto. Conciliación],
    [Codigo Auxiliar],
    [Tipo Documento],
    COALESCE([Nro. Documento], 
        LAG([Nro. Documento]) OVER (ORDER BY RowNum)
    ) AS [Nro. Documento],
    FechaEmision,
    CONVERT(NVARCHAR(10), 
        COALESCE(FechaVencimiento, 
            LAG(FechaVencimiento) OVER (ORDER BY RowNum)
        ), 105) AS FechaVencimiento,
    [Tipo Docto. Referencia],
    [Nro. Docto. Referencia],
    [Correlativo Documento de Compra Interno],
    [Monto 1],
    [Monto 2],
    [Monto 3],
    [Monto 4],
    [Monto 5],
    [Monto 6],
    [Monto 7],
    [Monto 8],
    [Monto 9],
    [Monto Suma Detalle Libro],
    [Número Documento Desde],
    [Número Documento Hasta],
    [Nro. agrupación en igual comprobante],
    [NºOpe.Doc.Aux.(Considera este campo sólo si se definió en los parametros de CW)],
    [Graba el detalle de libro (S/N)],
    [Documento Nulo (S/N)],
    [Tipo de comprobante (I, E, T)],
    [Descripción del comprobante]
FROM CTE
WHERE [Codigo Plan de cuenta] IS NOT NULL
ORDER BY RowNum;"

                Dim objadapter As New OleDbDataAdapter(sql, Conecta)
                Dim objdataset As New DataSet
                objadapter.Fill(objdataset)

                If objdataset.Tables(0).Rows.Count > 0 Then
                    ' Ruta donde se guardará el archivo CSV
                    Dim destino As String
                    Dim usu As Microsoft.VisualBasic.ApplicationServices.User
                    usu = New Microsoft.VisualBasic.ApplicationServices.User

                    Dim usuario() As String = Split(usu.Name, "\")

                    Dim fechaRemate As Date = fecha_remate
                    Dim sFilename As String = "S:\fbiobio" & fechaRemate.ToString("yyyyMMdd") & ".CSV"

                    destino = sFilename

                    Dim retries As Integer = 0
                    Dim maxRetries As Integer = 5
                    Dim success As Boolean = False

                    ' Intentar escribir el archivo con manejo de errores
                    Do While Not success And retries < maxRetries
                        Try

                            If fFile.Exists(sFilename) Then
                            Else

                                ' Crear un StreamWriter para escribir en el archivo CSV
                                Using writer As New StreamWriter(destino)
                                    ' Recorrer las filas y escribir los datos en el archivo
                                    For Each row As DataRow In objdataset.Tables(0).Rows
                                        Dim rowData As String = String.Join(",", row.ItemArray.Select(Function(value) value.ToString()))
                                        writer.WriteLine(rowData)
                                    Next
                                End Using
                                updateCSVStatus(idIngreso)
                                Return True
                            End If
                        Catch ex As Exception

                        End Try

                    Loop
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Return False
    End Function

    Public Function GuardarCSVchequeFechaMasDeUnaFactura(ByRef idIngreso As String, ByRef numero As Integer) As Boolean
        Try
            If ConectaBase() Then
                Dim sql As String = "

select 
'001' as [01], 
'1-1-3-01-03' as [02],
'0' as [03], 
f.porPagar as [04] ,
concat('F.V.- PAGO REM.: ' , r.remate) as [05],
'0' as [06], '0' as [07] ,
'' as [08],
'' as [09],
'' as [10],
'' as [11],
'' as [12],
'' as [13],
'' as [14],
'' as [15],
'' as [16],
'' as [17],
'' as [18],
'0' as [19],
 LEFT(f.rutComprador, LEN(f.rutComprador) - 2) as [20],
 'PB' as [21],
 idc.numeroCheq  as [22],
  CONVERT(NVARCHAR(10), f.fecha_rem, 105) AS [23],
  CONVERT(NVARCHAR(10), f.fecha_rem, 105) AS [24],
  'VB' as [25],
  f.idDTE as [26],
  '' as [27],
   '0' as [28],
 '0' as [29],
 '0' as [30],
 '0' as [31],
 '0' as [32],
 '0' as [33],
 '0' as [34],
 '0' as [35],
 '0' as [36],
 '0' as [37],
 '' as [38],
 '' as [39],
  i.numero as [40],
 'N' as [41],
 'I' as [42],
 'I' as [43],
 concat('INGRESO REMATE ',  r.remate , 'NRO.', i.numero)

from IngresosDet as d 
inner join FacturaElectronica as f on d.idIngreso = f.n_ingreso 
inner join remate as r on r.id = f.idRemate
inner join ingresosDetCheque as idc on idc.idIngresoDet = d.idIngresoDet
inner join Ingresos as i on i.idIngreso = d.idIngreso
where d.idIngreso = '" & idIngreso & "'
group by f.porPagar, r.remate, f.rutComprador,  idc.numeroCheq ,  f.fecha_rem, f.idDTE , i.numero

union all

select '001' as [01], 
'1-1-3-01-01' as [02], 
i.total as [03], '0' as [04] ,
concat('F.V.- PAGO REM.: ' , r.remate) as [05],
'0' as [06], '0' as [07] ,
'' as [08],
'' as [09],
'' as [10],
'' as [11],
'' as [12],
'' as [13],
'' as [14],
'' as [15],
'' as [16],
'' as [17],
'' as [18],
'0' as [19],
LEFT(i.rutCliente, LEN(i.rutCliente) - 2) as [20],
'PB' as [21],
idc.numeroCheq as [22],
 CONVERT(NVARCHAR(10), i.fecha, 105) AS [23],
 CONVERT(NVARCHAR(10), i.fecha, 105) AS [24],
 'PB' as [25],
 idc.numeroCheq as [26],
 '' as [27],
 '0' as [28],
 '0' as [29],
 '0' as [30],
 '0' as [31],
 '0' as [32],
 '0' as [33],
 '0' as [34],
 '0' as [35],
 '0' as [36],
 '0' as [37],
 '' as [38],
 '' as [39],
 i.numero as [40],
 'N' as [41],
 'I' as [42],
 'I' as [43],
 concat('INGRESO REMATE ',  r.remate , 'NRO.', i.numero)
from Ingresos as i inner join remate as r on i.idRemate = r.id
inner join IngresosDet as id on id.idIngreso = i.idIngreso
inner join ingresosDetCheque as idc on idc.idIngresoDet = id.idIngresoDet
where i.idIngreso = '" & idIngreso & "'"


                Dim objadapter As New OleDbDataAdapter(sql, Conecta)
                Dim objdataset As New DataSet
                objadapter.Fill(objdataset)

                If objdataset.Tables(0).Rows.Count > 0 Then
                    ' Ruta donde se guardará el archivo CSV
                    Dim destino As String
                    Dim usu As Microsoft.VisualBasic.ApplicationServices.User
                    usu = New Microsoft.VisualBasic.ApplicationServices.User

                    Dim usuario() As String = Split(usu.Name, "\")

                    Dim fechaRemate As Date = fecha_remate
                    Dim sFilename As String = "S:\fbiobio" & fechaRemate.ToString("yyyyMMdd") & ".CSV"

                    destino = sFilename

                    Dim retries As Integer = 0
                    Dim maxRetries As Integer = 5
                    Dim success As Boolean = False

                    Do While Not success And retries < maxRetries
                        Try

                            If fFile.Exists(sFilename) Then
                            Else

                                ' Crear un StreamWriter para escribir en el archivo CSV
                                Using writer As New StreamWriter(destino)
                                    ' Recorrer las filas y escribir los datos en el archivo
                                    For Each row As DataRow In objdataset.Tables(0).Rows
                                        Dim rowData As String = String.Join(",", row.ItemArray.Select(Function(value) value.ToString()))
                                        writer.WriteLine(rowData)
                                    Next
                                End Using
                                updateCSVStatus(idIngreso)
                                Return True
                            End If
                        Catch ex As Exception

                        End Try

                    Loop

                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Return False
    End Function
    Public Function GuardarCSVchequeDia(ByRef idIngreso As String, ByRef numero As Integer) As Boolean
        Try
            If ConectaBase() Then
                Dim sql As String = "WITH CTE AS (
    SELECT '001' AS [Area de Negocio],
           tp.cuentaSoft AS [Codigo Plan de cuenta],
           d.debito AS [Debe],
           d.credito AS [Haber],
           CASE
               WHEN  tp.cuentaSoft = '1-1-3-01-03' THEN concat('F.V.- PAGO REM.:', r.remate ,' PAGO CLIENT')
               WHEN  tp.cuentaSoft = '2-1-3-01-01' THEN ('F.C. -PAGO FACTURA CON VENTA')
               ELSE concat('CANCELACION FACTURAS REM.:', r.remate)
           END as [Descripción],
           '0' as EquivalenciaMoneda,
           '0' as [Debe al debe Moneda],
           '' as [Haber al haber Moneda],
           '' as [Código Condición de Venta],
           '' as [Código Vendedor],
           '' as [Código Ubicación],
           '' as [Código Concepto de Caja],
           '' as [Código Instrumento Financiero],
           '' as [Cantidad Instrumento Financiero],
           '' as [Código Detalle de Gasto],
           '' as [Cantidad Concepto de Gasto],
           '' as [Código Centro de Costo],
           '' as [Tipo Docto. Conciliación],
           CASE
               WHEN  tp.cuentaSoft = '1-1-3-01-03' THEN '0'
               ELSE '0'
           END AS [Nro. Docto. Conciliación],
           CASE
               WHEN  tp.cuentaSoft = '1-1-1-01-01' THEN LEFT(i.rutCliente, LEN(i.rutCliente) - 2)
               WHEN  tp.cuentaSoft = '1-1-3-01-01' THEN LEFT(i.rutCliente, LEN(i.rutCliente) - 2)
               WHEN  tp.cuentaSoft = '1-1-3-01-03' THEN LEFT(i.rutCliente, LEN(i.rutCliente) - 2)
               ELSE ''
           END AS [Codigo Auxiliar],
           'PB' AS [Tipo Documento],
           (select top 1 idc.numeroCheq from ingresosDetCheque as idc where (idc.idIngresoDet = d.idIngresoDet )) AS [Nro. Documento],
           CONVERT(NVARCHAR(10), i.fecha, 105) as FechaEmision,
            CONVERT(NVARCHAR(10), i.fecha, 105) as FechaVencimiento,
           '' as descripcion,
           CASE
               WHEN  tp.cuentaSoft = '1-1-3-01-03' THEN 'VB'
               ELSE 'PB'		  
           END AS [Tipo Docto. Referencia],
           
		    CASE
               WHEN  tp.cuentaSoft = '1-1-3-01-03' THEN f.idDTE
               ELSE (select top 1 idc.numeroCheq from ingresosDetCheque as idc where (idc.idIngresoDet = d.idIngresoDet ))		  
           END    AS [Nro. Docto. Referencia],
           '' as [Correlativo Documento de Compra Interno],
           '0' as [Monto 1],
           '0' as [Monto 2],
           '0' as [Monto 3],
           '0' as [Monto 4],
           '0' as [Monto 5],
           '0' as [Monto 6],
           '0' as [Monto 7],
           '0' as [Monto 8],
           '0' as [Monto 9],
           '0' as [Monto Suma Detalle Libro],
           '' as [Número Documento Desde],
           '' as [Número Documento Hasta],
           i.numero as [Nro. agrupación en igual comprobante],
           'N' as [NºOpe.Doc.Aux.(Considera este campo sólo si se definió en los parametros de CW)],
           'I' as [Graba el detalle de libro (S/N) ],
           'I' as [Documento Nulo (S/N)],
           concat('INGRESO REMATE ',r.remate ,' NRO. ', I.numero )  as [Tipo de comprobante (I, E, T)],
           '' AS [Descripción del comprobante],
           ROW_NUMBER() OVER (ORDER BY d.debito DESC) as RowNum
    FROM Ingresos AS i
    INNER JOIN IngresosDet AS d ON i.idIngreso = d.idIngreso
    FULL JOIN tipoPago as tp on tp.cuentaIngreso = d.codigo
    FULL JOIN Cuentas as c on c.idCuenta = d.codigo
    FULL JOIN remate as r on r.id = i.idRemate
    FULL JOIN FacturaElectronica as f  on f.n_ingreso = i.idIngreso
    WHERE i.idIngreso = '" & idIngreso & "'
)
SELECT 
    [Area de Negocio],
    [Codigo Plan de cuenta],
    [Debe],
    [Haber],
    [Descripción],
    EquivalenciaMoneda,
    [Debe al debe Moneda],
    [Haber al haber Moneda],
    [Código Condición de Venta],
    [Código Vendedor],
    [Código Ubicación],
    [Código Concepto de Caja],
    [Código Instrumento Financiero],
    [Cantidad Instrumento Financiero],
    [Código Detalle de Gasto],
    [Cantidad Concepto de Gasto],
    [Código Centro de Costo],
    [Tipo Docto. Conciliación],
    [Nro. Docto. Conciliación],
    [Codigo Auxiliar],
    [Tipo Documento],
    COALESCE([Nro. Documento], 
        LAG([Nro. Documento]) OVER (ORDER BY RowNum)
    ) AS [Nro. Documento],
    FechaEmision,
    CONVERT(NVARCHAR(10), 
        COALESCE(FechaVencimiento, 
            LAG(FechaVencimiento) OVER (ORDER BY RowNum)
        ), 105) AS FechaVencimiento,
    descripcion,
    [Tipo Docto. Referencia],
    [Nro. Docto. Referencia],
    [Correlativo Documento de Compra Interno],
    [Monto 1],
    [Monto 2],
    [Monto 3],
    [Monto 4],
    [Monto 5],
    [Monto 6],
    [Monto 7],
    [Monto 8],
    [Monto 9],
    [Monto Suma Detalle Libro],
    [Número Documento Desde],
    [Número Documento Hasta],
    [Nro. agrupación en igual comprobante],
    [NºOpe.Doc.Aux.(Considera este campo sólo si se definió en los parametros de CW)],
    [Graba el detalle de libro (S/N) ],
    [Documento Nulo (S/N)],
    [Tipo de comprobante (I, E, T)],
    [Descripción del comprobante]
FROM CTE
ORDER BY RowNum"


                Dim objadapter As New OleDbDataAdapter(sql, Conecta)
                Dim objdataset As New DataSet
                objadapter.Fill(objdataset)

                If objdataset.Tables(0).Rows.Count > 0 Then
                    ' Ruta donde se guardará el archivo CSV
                    Dim destino As String
                    Dim usu As Microsoft.VisualBasic.ApplicationServices.User
                    usu = New Microsoft.VisualBasic.ApplicationServices.User

                    Dim usuario() As String = Split(usu.Name, "\")

                    Dim fechaRemate As Date = fecha_remate
                    Dim sFilename As String = "S:\fbiobio" & fechaRemate.ToString("yyyyMMdd") & ".CSV"

                    destino = sFilename

                    Dim retries As Integer = 0
                    Dim maxRetries As Integer = 5
                    Dim success As Boolean = False

                    ' Intentar escribir el archivo con manejo de errores
                    Do While Not success And retries < maxRetries

                        Try

                            If fFile.Exists(sFilename) Then
                            Else

                                ' Crear un StreamWriter para escribir en el archivo CSV
                                Using writer As New StreamWriter(destino)
                                    ' Recorrer las filas y escribir los datos en el archivo
                                    For Each row As DataRow In objdataset.Tables(0).Rows
                                        Dim rowData As String = String.Join(",", row.ItemArray.Select(Function(value) value.ToString()))
                                        writer.WriteLine(rowData)
                                    Next
                                End Using
                                updateCSVStatus(idIngreso)

                                Return True
                            End If
                        Catch ex As Exception

                        End Try

                    Loop
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Return False
    End Function
    Public Function GuardarCSVPagoVenta(ByRef idIngreso As String, ByRef numero As Integer) As Boolean
        Try
            If ConectaBase() Then
                Dim sql As String = "WITH GroupedData AS (
    SELECT   
        d.cuentaSoft,
        SUM(CASE 
            WHEN d.codigo != '1112101' THEN d.debito 
            ELSE 0 
        END) AS TotalDebe,
        SUM(CASE 
            WHEN d.codigo != '1112101' THEN d.credito 
            ELSE 0 
        END) AS TotalCredito,
        MIN(tp.cuentaIngreso) AS cuentaIngreso,
        MIN(r.remate) AS remate,
        MIN(i.numero) AS numero,
        MIN(i.rutCliente) AS rutCliente,
        MIN(i.fecha) AS fecha,
        MIN(f.idDTE) AS idDTE,
        MIN(i.idRemate) AS idRemate,
        MIN(tp.cuentaSoft) AS tpCuentaSoft,
        MIN(i.idIngreso) AS idIngreso
    FROM Ingresos AS i
    FULL JOIN IngresosDet AS d ON i.idIngreso = d.idIngreso
    FULL JOIN tipoPago AS tp ON tp.cuentaIngreso = d.codigo
    FULL JOIN Cuentas AS c ON c.idCuenta = d.codigo
    FULL JOIN remate AS r ON r.id = i.idRemate
    FULL JOIN FacturaElectronica AS f ON f.n_ingreso = i.idIngreso
    WHERE i.idIngreso ='" & idIngreso & "'
    AND d.codigo NOT IN ('1112101')
    GROUP BY d.cuentaSoft
)
SELECT
    '001' AS [Area de Negocio],
    cuentaSoft AS [Codigo Plan de cuenta],
    TotalDebe AS [Debe],
    CASE 
        WHEN TotalCredito <> 0 THEN
            TotalCredito - ISNULL((
                SELECT TOP 1 debito 
                FROM IngresosDet 
                WHERE codigo = '1112101' AND idIngreso = gd.idIngreso
            ), 0)
        ELSE TotalCredito
    END AS [Haber],
    CASE
        WHEN cuentaIngreso = '1-1-3-01-03' THEN CONCAT('F.V.- PAGO REM.:', remate, ' PAGO CLIENT')
        WHEN cuentaIngreso = '2-1-3-01-01' THEN 'F.C. -PAGO FACTURA CON VENTA'
        ELSE CONCAT('CANCELACION FACTURAS REM.:', remate)
    END AS [Descripción],
    '0' AS EquivalenciaMoneda,
    '0' AS [Debe al debe Moneda],
    '' AS [Haber al haber Moneda],
    '' AS [Código Condición de Venta],
    '' AS [Código Vendedor],
    '' AS [Código Ubicación],
    '' AS [Código Concepto de Caja],
    '' AS [Código Instrumento Financiero],
    '' AS [Cantidad Instrumento Financiero],
    '' AS [Código Detalle de Gasto],
    '' AS [Cantidad Concepto de Gasto],
    '' AS [Código Centro de Costo],
    CASE
        WHEN cuentaSoft = '1-1-1-02-02' THEN 'ID'
        ELSE ''
    END AS [Tipo Docto. Conciliación],
    CASE
        WHEN cuentaSoft = '1-1-1-02-02' THEN remate
        ELSE '0'
    END AS [Nro. Docto. Conciliación],
   
    CASE
        WHEN cuentaSoft = '1-1-1-01-01'  THEN '' else 
   LEFT(rutCliente, LEN(rutCliente) - 2)  end AS [Codigo Auxiliar],
    CASE
        WHEN tpCuentaSoft = '1-1-1-01-01' THEN ''
        WHEN tpCuentaSoft = '1-1-3-01-01' THEN 'PB'
        WHEN tpCuentaSoft = '1-1-3-01-03' THEN 'PA' ELSE '' end AS [Tipo Documento],
    '' AS [Nro. Documento],
    CONVERT(NVARCHAR(10), fecha, 105) AS FechaEmision,
    CONVERT(NVARCHAR(10), fecha, 105) AS FechaVencimiento,
    CASE
        WHEN tpCuentaSoft = '1-1-1-01-01' THEN ''
        WHEN tpCuentaSoft = '1-1-3-01-01' THEN 'PB'
        WHEN tpCuentaSoft = '1-1-3-01-03' THEN 'VB'
        ELSE 
        (SELECT TOP 1 'CF' AS TIPO FROM LiquidacionFacturaElectronica 
         WHERE rutVendedor = gd.rutCliente AND idRemate = gd.idRemate 
         UNION ALL 
         SELECT TOP 1 'CD' AS TIPO FROM FacturaCompra 
         WHERE rutVendedor = gd.rutCliente AND idRemate = gd.idRemate)
    END AS [Tipo Docto. Referencia],
  '' AS [Nro. Docto. Referencia],
    '' AS [Correlativo Documento de Compra Interno],
    '0' AS [Monto 1],
    '0' AS [Monto 2],
    '0' AS [Monto 3],
    '0' AS [Monto 4],
    '0' AS [Monto 5],
    '0' AS [Monto 6],
    '0' AS [Monto 7],
    '0' AS [Monto 8],
    '0' AS [Monto 9],
    '0' AS [Monto Suma Detalle Libro],
    '' AS [Numero Documento Desde],
    '' AS [Número Documento Hasta],
    CAST(numero AS NVARCHAR) AS [Nro. agrupación en igual comprobante],
    'N' AS [NºOpe.Doc.Aux.(Considera este campo sólo si se definió en los parametros de CW)'],
    'I' AS [Graba el detalle de libro (S/N)'],
    'I' AS [Documento Nulo (S/N)'],
    CONCAT('INGRESO REMATE ', remate, ' NRO. ', CAST(numero AS NVARCHAR)) AS [Tipo de comprobante (I, E, T)]
FROM GroupedData gd
ORDER BY [Debe] DESC;"

                Dim objadapter As New OleDbDataAdapter(sql, Conecta)
                Dim objdataset As New DataSet
                objadapter.Fill(objdataset)

                If objdataset.Tables(0).Rows.Count > 0 Then
                    ' Ruta donde se guardará el archivo CSV
                    Dim destino As String
                    Dim usu As Microsoft.VisualBasic.ApplicationServices.User
                    usu = New Microsoft.VisualBasic.ApplicationServices.User

                    Dim fechaRemate As Date = fecha_remate
                    Dim sFilename As String = "S:\fbiobio" & fechaRemate.ToString("yyyyMMdd") & ".CSV"

                    destino = sFilename

                    Dim retries As Integer = 0
                    Dim maxRetries As Integer = 5
                    Dim success As Boolean = False

                    ' Intentar escribir el archivo con manejo de errores
                    Do While Not success And retries < maxRetries

                        Try
                            If fFile.Exists(sFilename) Then
                            Else


                                ' Crear un StreamWriter para escribir en el archivo CSV
                                Using writer As New StreamWriter(destino)
                                    ' Recorrer las filas y escribir los datos en el archivo
                                    For Each row As DataRow In objdataset.Tables(0).Rows
                                        Dim rowData As String = String.Join(",", row.ItemArray.Select(Function(value) value.ToString()))
                                        writer.WriteLine(rowData)
                                    Next
                                End Using
                            End If
                            updateCSVStatus(idIngreso)
                            Return True
                        Catch ex As Exception

                        End Try
                    Loop
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Return False
    End Function

    Public Function GuardarCSVtransferencia(ByRef idIngreso As String, ByRef numero As Integer) As Boolean
        Try
            If ConectaBase() Then
                Dim sql As String = "SELECT '001' AS [Area de Negocio],
     d.cuentaSoft AS [Codigo Plan de cuenta],
       d.debito AS [Debe],
       d.credito AS [Haber],
       CASE
           WHEN tp.cuentaIngreso = '1-1-3-01-03' THEN concat('F.V.- PAGO REM.:', r.remate ,' PAGO CLIENT')
           WHEN tp.cuentaIngreso = '2-1-3-01-01' THEN ('F.C. -PAGO FACTURA CON VENTA')
           ELSE concat('CANCELACION FACTURAS REM.:', r.remate)
       END as [Descripción],
       '0' as EquivalenciaMoneda,
       '0' as [Debe al debe Moneda],
       '' as [Haber al haber Moneda],
       '' as [Código Condición de Venta],
       '' as [Código Vendedor],
       '' as [Código Ubicación],
       '' as [Código Concepto de Caja],
       '' as [Código Instrumento Financiero],
       '' as [Cantidad Instrumento Financiero],
       '' as [Código Detalle de Gasto],
       '' as [Cantidad Concepto de Gasto],
       '' as [Código Centro de Costo],
       CASE
           WHEN d.cuentaSoft = '1-1-1-02-02' THEN 'ID'
           ELSE ''
       END  as [Tipo Docto. Conciliación],
       CASE
           WHEN d.cuentaSoft = '1-1-1-02-02' THEN r.remate
           ELSE '0'
       END AS [Nro. Docto. Conciliación],
       CASE
           WHEN tp.cuentaSoft = '1-1-3-01-03' THEN LEFT(i.rutCliente, LEN(i.rutCliente) - 2)
           ELSE ''
       END AS [Codigo Auxiliar],
       CASE
           WHEN tp.cuentaSoft = '1-1-1-01-01' THEN ''
           WHEN tp.cuentaSoft = '1-1-3-01-01' THEN 'PB'
           WHEN tp.cuentaSoft = '1-1-3-01-03' THEN 'ID'
           ELSE ''
       END AS [Tipo Documento],
       CASE
           WHEN tp.cuentaSoft= '1-1-1-01-01' THEN ''
           WHEN tp.cuentaSoft = '1-1-3-01-03' THEN CAST(r.remate AS NVARCHAR)
           ELSE ''
       END AS [Nro. Documento],
       CONVERT(NVARCHAR(10), i.fecha, 105) as FechaEmision,
       CONVERT(NVARCHAR(10), i.fecha, 105) as FechaVencimiento,
       CASE
           WHEN tp.cuentaSoft = '1-1-1-01-01' THEN ''
           WHEN tp.cuentaSoft = '1-1-3-01-01' THEN 'PB'
           WHEN tp.cuentaSoft = '1-1-3-01-03' THEN 'VB'
           ELSE ''
       END AS [Tipo Docto. Referencia],
       
	    CASE
           WHEN tp.cuentaSoft = '1-1-3-01-03' THEN  CAST( f.idDTE AS NVARCHAR)
           ELSE ''
       END AS [Nro. Docto. Referencia],
       '' as [Correlativo Documento de Compra Interno],
       '0' as [Monto 1],
       '0' as [Monto 2],
       '0' as [Monto 3],
       '0' as [Monto 4],
       '0' as [Monto 5],
       '0' as [Monto 6],
       '0' as [Monto 7],
       '0' as [Monto 8],
       '0' as [Monto 9],
       '0' as [Monto Suma Detalle Libro],
       '' as [Número Documento Desde],
       '' as [Número Documento Hasta],
       CAST(i.numero AS NVARCHAR) as [Nro. agrupación en igual comprobante],
       'N' as [NºOpe.Doc.Aux.(Considera este campo sólo si se definió en los parametros de CW)],
       'I' as [Graba el detalle de libro (S/N)],
       'I' as [Documento Nulo (S/N)],
       concat('INGRESO REMATE ',r.remate ,' NRO. ', CAST(i.numero AS NVARCHAR))  as [Tipo de comprobante (I, E, T)],
       '' AS [Descripción del comprobante]
FROM Ingresos AS i
full JOIN IngresosDet AS d ON i.idIngreso = d.idIngreso
full join tipoPago as tp on tp.cuentaIngreso = d.codigo
full join Cuentas as c on c.idCuenta = d.codigo
full join remate as r on r.id = i.idRemate
full join FacturaElectronica as f  on f.n_ingreso = i.idIngreso
WHERE i.idIngreso = '" & idIngreso & "'
ORDER BY d.debito DESC"

                Dim objadapter As New OleDbDataAdapter(sql, Conecta)
                Dim objdataset As New DataSet
                objadapter.Fill(objdataset)

                If objdataset.Tables(0).Rows.Count > 0 Then
                    ' Ruta donde se guardará el archivo CSV
                    Dim destino As String
                    Dim usu As Microsoft.VisualBasic.ApplicationServices.User
                    usu = New Microsoft.VisualBasic.ApplicationServices.User

                    Dim fechaRemate As Date = fecha_remate
                    Dim sFilename As String = "S:\fbiobio" & fechaRemate.ToString("yyyyMMdd") & ".CSV"

                    destino = sFilename


                    Dim retries As Integer = 0
                    Dim maxRetries As Integer = 5
                    Dim success As Boolean = False

                    ' Intentar escribir el archivo con manejo de errores
                    Do While Not success And retries < maxRetries

                        Try
                            If fFile.Exists(sFilename) Then
                            Else

                                ' Crear un StreamWriter para escribir en el archivo CSV
                                Using writer As New StreamWriter(destino)
                                    ' Recorrer las filas y escribir los datos en el archivo
                                    For Each row As DataRow In objdataset.Tables(0).Rows
                                        Dim rowData As String = String.Join(",", row.ItemArray.Select(Function(value) value.ToString()))
                                        writer.WriteLine(rowData)
                                    Next
                                End Using
                                updateCSVStatus(idIngreso)
                                Return True
                            End If

                        Catch ex As Exception
                            MsgBox(ex.ToString)
                        End Try
                    Loop

                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        Return False
    End Function
End Class
