Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO

Public Class egresoController
    Public sql As String
    Public objadapter As OleDbDataAdapter
    Public objdataset As DataSet
    Public conexiones As SqlConnection
    Public Egreso As New Egresos
    Dim fFile As IO.File, fStream As IO.FileStream

    Public Function getAllParam(ByRef parametro As String) As Boolean

        Try
            If ConectaBase() Then

                sql = "SELECT e.idEgreso,
       e.numero AS Numero,  
       FORMAT(e.fechaHora, 'dd-MM-yyyy') AS Fecha, 
	   FORMAT(e.fechaHora, 'HH:mm') AS Hora, 
       e.rut AS RUT,
       c.nombres AS Cliente,
      e.total AS Total,
       e.motivo AS Motivo,
       CASE 
           WHEN e.motivo = 'ANTICIPO CLIENTES' THEN 'Anticipo'
           ELSE edc.tipo 
       END AS Tipo,
       r.remate AS Remate,
       e.anulado AS Anulado,
       edc.numeroCheq AS [Num.Cheque],
	   e.tipoDoc as Documento,
	   case when e.tipoDoc = 'LF'  
	   then (select top 1 idDTE from LiquidacionFacturaElectronica where idEgreso = e.idEgreso)
	   else (select top 1 idDTE from FacturaCompra where idEgreso = e.idEgreso) end as Folio
	   
FROM Egresos AS e
INNER JOIN clientes AS c ON e.rut = c.rut
INNER JOIN remate AS r ON r.id = e.idRemate
INNER JOIN egresosDet AS d ON d.idEgreso = e.idEgreso
FULL JOIN egresosDetCheque AS edc ON edc.idEgresoDetalle = d.idEgresoDetalle  " & parametro & "
  AND (CASE WHEN e.motivo = 'ANTICIPO CLIENTES' THEN 'Anticipo' ELSE edc.tipo END) IS NOT NULL
GROUP BY e.idEgreso, e.numero, e.fechaHora, e.rut, c.nombres, e.motivo, r.remate, e.anulado, edc.tipo, edc.numeroCheq, e.tipoDoc, e.total
ORDER BY e.numero ASC;"

                objadapter = New OleDbDataAdapter(sql, Conecta)
                objdataset = New DataSet
                objadapter.Fill(objdataset)

                If objdataset.Tables(0).Rows.Count > 0 Then
                    Return True
                End If

            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return False
    End Function

    Public Function updateCSVStatus(ByVal idEgreso As String) As Boolean
        Try
            Using conexiones As New SqlConnection(conexion_string)
                conexiones.Open()

                Dim consulta As String = "UPDATE Egresos SET " &
                                    "csv = 1 WHERE idEgreso = @idEgreso"

                Using enunciado As New SqlCommand(consulta, conexiones)
                    ' Asignar valores a los parámetros
                    enunciado.Parameters.AddWithValue("@idEgreso", idEgreso)

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

    Public Function GuardarCSVanticipo(ByRef idEgreso As String, ByRef numero As Integer) As Boolean
        Try
            If ConectaBase() Then
                Dim sql As String = "select '001' as [1] ,
c.idCuenta  as [2],
d.debito as [3],
d.credito as [4],
case 
when
c.idCuenta = '2-1-3-01-01' 
then concat(e.tipoDoc, ' REM.:', r.remate )
else c.descripcion end as [5],
'0' as [6],
       '0' as [7],
       '' as [8],
       '' as [9],
       '' as [10],
       '' as [11],
       '' as [12],
       '' as [13],
       ''as [14],
       '' as [15],
       '' as [16],
       '' as [17],
	   '' as [18],
	   '0' as [19],
	    CASE
	   WHEN c.idCuenta = '1-1-6-01-02' THEN LEFT(e.rut, LEN(e.rut) - 2)
           ELSE ''
       END as [20],
	   CASE 
	   WHEN c.idCuenta = '1-1-6-01-02' THEN 'AA' 
	   ELSE '' END as [21],
	   CASE
		   WHEN c.idCuenta = '1-1-6-01-02'  THEN CONVERT(VARCHAR(255), r.remate) else ''
       END AS [22],
	    CASE 
		WHEN
		c.idCuenta =  '1-1-6-01-02' THEN FORMAT( e.fechaHora, 'dd/MM/yyyy')
		else FORMAT( e.fechaHora, 'dd/MM/yyyy')
		end  as [23],
		 CASE 
		WHEN
		c.idCuenta =  '1-1-6-01-02' THEN FORMAT( e.fechaHora, 'dd/MM/yyyy')
		else FORMAT( e.fechaHora, 'dd/MM/yyyy')
		end  as [24],
			   CASE 
	   WHEN c.idCuenta = '1-1-6-01-02' THEN 'AA' 
	   ELSE '' END as [25],
	    CASE
		   WHEN c.idCuenta = '1-1-6-01-02'  THEN CONVERT(VARCHAR(255), r.remate) else ''
       END AS [26],
	   '' AS [27],
	    '0' AS [28],
		'0' AS [29],
		'0' AS [30],
		'0' AS [31],
		'0' AS [32],
		'0' AS [33],
		'0' AS [34],
		'0' AS [35],
		'0' AS [36],
		'0' AS [37],
		'' as [38],
		'' as [39],
		e.numero as [40],
		 'N' as [41],
		'E' as [42],
		'E' as [43],
	   concat('EGRESO REMATE ',r.remate ,' NRO. ', e.numero )  as [44]
from Egresos as e 
inner join egresosDet as d on d.idEgreso = e.idEgreso
inner join Cuentas as c on c.idCuenta = d.idCuenta
inner join remate as r on r.id = e.idRemate
full join egresosDetCheque as edc on edc.idEgresoDetalle = d.idEgresoDetalle
where e.idEgreso = '" & idEgreso & "'
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
                                    Dim lastNroDocumento As String = ""
                                    ' Recorrer las filas y escribir los datos en el archivo
                                    For i As Integer = 0 To objdataset.Tables(0).Rows.Count - 1
                                        Dim row As DataRow = objdataset.Tables(0).Rows(i)
                                        Dim values As New List(Of String)

                                        ' Procesar cada columna
                                        For j As Integer = 0 To row.ItemArray.Length - 1
                                            If j = 21 AndAlso row(1).ToString() = "1-1-6-01-02" Then
                                                ' Si es la columna 22 (índice 21) y la cuenta es 1-1-6-01-02
                                                values.Add(lastNroDocumento)
                                            Else
                                                values.Add(row(j).ToString())
                                                ' Guardar el número de documento si no es cuenta 1-1-6-01-02
                                                If j = 21 AndAlso row(1).ToString() <> "1-1-6-01-02" Then
                                                    lastNroDocumento = row(j).ToString()
                                                End If
                                            End If
                                        Next

                                        Dim rowData As String = String.Join(",", values)
                                        writer.WriteLine(rowData)
                                    Next
                                    writer.Close()
                                End Using
                                updateCSVStatus(idEgreso)
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

    Public Function GuardarCSVchequeAlDia(ByRef idEgreso As String, ByRef numero As Integer) As Boolean
        Try
            If ConectaBase() Then
                Dim sql As String = "WITH liquidaciones AS (
        SELECT 
            lf.idEgreso,
            lf.idDTE,
            lf.fecha_vencimiento,
            ROW_NUMBER() OVER (PARTITION BY lf.idEgreso ORDER BY lf.idDTE) AS RowNum,
            COUNT(*) OVER (PARTITION BY lf.idEgreso) AS TotalFacturas,
    		case when isnull(lf.porPagar,'') <> '' then lf.porPagar else (select lf.porPagar from LiquidacionFacturaElectronica as lf where lf.idEgreso = lf.idEgreso) end as porPagar
        FROM 
            LiquidacionFacturaElectronica lf
        WHERE 
            lf.idEgreso = '" & idEgreso & "'
			 UNION ALL

    SELECT 
        fc.idEgreso,
        fc.idDTE,
        fc.fecha_vencimiento,
        ROW_NUMBER() OVER (PARTITION BY fc.idEgreso ORDER BY fc.idDTE) AS RowNum,
        COUNT(*) OVER (PARTITION BY fc.idEgreso) AS TotalFacturas,
        CASE 
            WHEN ISNULL(fc.porPagar, '') <> '' THEN fc.porPagar 
            ELSE (SELECT fc.porPagar FROM FacturaCompra AS fc2 WHERE fc2.idEgreso = fc.idEgreso) 
        END AS porPagar
    FROM 
        FacturaCompra fc
    WHERE 
        fc.idEgreso = '" & idEgreso & "'

    ),
    Distribucion AS (
        SELECT 
            d.idEgresoDetalle,
            d.debito / f.TotalFacturas AS MontoDistribuido,
            f.idDTE,
            f.fecha_vencimiento,
            f.RowNum,
            f.TotalFacturas,
    		f.porPagar
        FROM 
            egresosDet d
        INNER JOIN liquidaciones f ON d.idEgreso = f.idEgreso
        WHERE 
            d.idCuenta = '2-1-3-01-01'
    )
    SELECT 
        '001' AS [1. AreaNegocio],
        c.idCuenta AS [2. idCuenta],
        CASE 
            WHEN d.idCuenta = '2-1-3-01-01' THEN dis.porPagar
            ELSE d.debito
        END AS [3. debito],
        d.credito AS [4. credito],
        CASE 
            WHEN c.idCuenta = '2-1-3-01-01' THEN CONCAT(e.tipoDoc, ' REM.:', r.remate)
            ELSE c.descripcion
        END AS [5. descripcion],
        '0' AS [6. EquivalenciaMoneda],
        '0' AS [7. Debe al debe Moneda],
        '' AS [8. Haber al haber Moneda],
        '' AS [9. Código Condición de Venta],
        '' AS [10. Código Vendedor],
        '' AS [11. Código Ubicación],
        '' AS [12. Código Concepto de Caja],
        '' AS [13. Código Instrumento Financiero],
        '' AS [14. Cantidad Instrumento Financiero],
        '' AS [15. Código Detalle de Gasto],
        '' AS [16. Cantidad Concepto de Gasto],
        '' AS [17. Código Centro de Costo],
        case when d.idCuenta <> '2-1-3-01-01' then 'PE' else '' end AS [18. Tipo Docto. Conciliación],
       case when d.idCuenta <> '2-1-3-01-01'then edc.numeroCheq else '0' end AS [19. Nro. Docto. Conciliación],
        case when d.idCuenta = '2-1-3-01-01'then  LEFT(e.rut, LEN(e.rut) - 2) else '' end AS [20. Codigo Auxiliar],
        case when d.idCuenta = '2-1-3-01-01'then 'PE' else '' end AS [21. Tipo Documento],

       case when isnull(edc.numeroCheq,'') = '' then (select idc.numeroCheq from Egresos as e1 
		 inner join egresosDet as d on e1.idEgreso = d.idEgreso 
		 inner join egresosDetCheque as idc on idc.idEgresoDetalle = d.idEgresoDetalle
		 where e1.idEgreso = e.idEgreso) else '' end AS [22. Nro. Documento]
		 
		 ,
        FORMAT(e.fechaHora, 'dd/MM/yyyy') AS [23. Fecha Emision],
        FORMAT(ISNULL(dis.fecha_vencimiento, e.fechaHora), 'dd/MM/yyyy') AS [24. Fecha Vencimiento],
        CASE 
            WHEN  c.idCuenta = '2-1-3-01-01' and e.tipoDoc = 'LF' THEN 'CF'
            WHEN  c.idCuenta = '2-1-3-01-01' and e.tipoDoc = 'FC' THEN 'CD'
        END AS [25. Tipo Doc Soft],
        ISNULL(CAST(dis.idDTE AS VARCHAR), ' ') AS [26. idDTE],
        '' AS [27. Monto 1],
        '0' AS [28. Monto 2],
        '0' AS [29. Monto 4],
        '0' AS [30. Monto 5],
        '0' AS [31. Monto 6],
        '0' AS [32. Monto 7],
        '0' AS [33. Monto 8],
        '0' AS [34. Monto 9],
        '0' AS [35. Monto Suma Detalle Libro],
        '0' AS [36. Número Documento Desde],
        '0' AS [37. Número Documento Hasta],
        '' AS [38. Monto 10],
        '' AS [39. Monto 11],
        e.numero AS [40. Nro. agrupación en igual comprobante],
        'N' AS [41. NºOpe.Doc.Aux.(Considera este campo sólo si se definió en los parametros de CW)],
        'E' AS [42. Graba el detalle de libro (S/N)],
        'E' AS [43. Documento Nulo (S/N)],
        CONCAT('EGRESO REMATE ', r.remate, ' NRO. ', e.numero) AS [44. Tipo de comprobante (I, E, T)]
    FROM 
        Egresos e
    INNER JOIN 
        egresosDet d ON d.idEgreso = e.idEgreso
    INNER JOIN 
        Cuentas c ON c.idCuenta = d.idCuenta
    INNER JOIN 
        remate r ON r.id = e.idRemate
    LEFT JOIN 
        egresosDetCheque edc ON edc.idEgresoDetalle = d.idEgresoDetalle
    LEFT JOIN 
        Distribucion dis ON dis.idEgresoDetalle = d.idEgresoDetalle

    WHERE 
        e.idEgreso = '" & idEgreso & "'
    ORDER BY 
         c.idCuenta desc"

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

                                Using writer As New StreamWriter(destino)
                                    ' Recorrer las filas y escribir los datos en el archivo
                                    For Each row As DataRow In objdataset.Tables(0).Rows
                                        Dim rowData As String = String.Join(",", row.ItemArray.Select(Function(value) value.ToString()))
                                        writer.WriteLine(rowData)
                                    Next
                                    writer.Close()
                                End Using

                                updateCSVStatus(idEgreso)
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
    Public Function GuardarCSVchequeAlDiaConAnticipo(ByRef idEgreso As String, ByRef idEgreso2 As String, ByRef numero As Integer) As Boolean
        Try
            If ConectaBase() Then
                Dim sql As String = "WITH liquidaciones AS (
    SELECT 
        lf.idEgreso,
        lf.idDTE,
        lf.fecha_vencimiento,
        ROW_NUMBER() OVER (PARTITION BY lf.idEgreso ORDER BY lf.idDTE) AS RowNum,
        COUNT(*) OVER (PARTITION BY lf.idEgreso) AS TotalFacturas,
        ISNULL(lf.porPagar, (SELECT TOP 1 porPagar FROM LiquidacionFacturaElectronica WHERE idEgreso = lf.idEgreso)) AS porPagar
    FROM LiquidacionFacturaElectronica lf
    WHERE lf.idEgreso IN ('" & idEgreso & "')
    
    UNION ALL

    SELECT 
        fc.idEgreso,
        fc.idDTE,
        fc.fecha_vencimiento,
        ROW_NUMBER() OVER (PARTITION BY fc.idEgreso ORDER BY fc.idDTE) AS RowNum,
        COUNT(*) OVER (PARTITION BY fc.idEgreso) AS TotalFacturas,
        ISNULL(fc.porPagar, (SELECT TOP 1 porPagar FROM FacturaCompra WHERE idEgreso = fc.idEgreso)) AS porPagar
    FROM FacturaCompra fc
    WHERE fc.idEgreso IN ('" & idEgreso & "')
),
Distribucion AS (
    SELECT 
        d.idEgresoDetalle,
        d.debito / f.TotalFacturas AS MontoDistribuido,
        f.idDTE,
        f.fecha_vencimiento,
        f.RowNum,
        f.TotalFacturas,
        f.porPagar
    FROM egresosDet d
    INNER JOIN liquidaciones f ON d.idEgreso = f.idEgreso
    WHERE d.idCuenta IN ('2-1-3-01-01', '1-1-6-01-02')
)
SELECT 
    '001' AS [1],
    CASE WHEN c.idCuenta = '1-1-1-01-01' THEN '1-1-6-01-02' ELSE c.idCuenta END AS [2],
    CASE 
        WHEN d.idCuenta = '2-1-3-01-01' THEN dis.porPagar
        ELSE d.debito
    END AS [3],
    d.credito AS [4],
    CASE 
        WHEN d.idCuenta IN ('2-1-3-01-01', '1-1-6-01-02') THEN CONCAT(e.tipoDoc, ' REM.:', r.remate)
        ELSE c.descripcion
    END AS [5],
    '0' AS [6],
    '0' AS [7],
    '' AS [8],
    '' AS [9],
    '' AS [10],
    '' AS [11],
    '' AS [12],
    '' AS [13],
    '' AS [14],
    '' AS [15],
    '' AS [16],
    '' AS [17],
    CASE WHEN d.idCuenta NOT IN ('2-1-3-01-01', '1-1-6-01-02','1-1-1-01-01') THEN 'PE' ELSE '' END AS [18],
    CASE WHEN d.idCuenta NOT IN ('2-1-3-01-01', '1-1-6-01-02','1-1-1-01-01') THEN edc.numeroCheq ELSE '0' END AS [19],
      CASE WHEN d.idCuenta IN ('2-1-3-01-01', '1-1-6-01-02', '1-1-1-01-01') THEN LEFT(e.rut, LEN(e.rut) - 2) ELSE '' END AS [20. Codigo Auxiliar],
     CASE WHEN d.idCuenta IN ('2-1-3-01-01', '1-1-6-01-02', '1-1-1-01-01') THEN 'PE' ELSE '' END AS [21. Tipo Documento],
    CASE 
        WHEN d.idCuenta = '2-1-3-01-01' THEN 
            (SELECT TOP 1 edc2.numeroCheq 
             FROM egresosDet d2 
             INNER JOIN egresosDetCheque edc2 ON d2.idEgresoDetalle = edc2.idEgresoDetalle 
             WHERE d2.idEgreso = e.idEgreso)
        ELSE ''
    END AS [22],
    FORMAT(e.fechaHora, 'dd/MM/yyyy') AS [23],
    FORMAT(ISNULL(dis.fecha_vencimiento, e.fechaHora), 'dd/MM/yyyy') AS [24],
    CASE 
        WHEN c.idCuenta = '2-1-3-01-01' AND e.tipoDoc = 'LF' THEN 'CF'
        WHEN c.idCuenta = '2-1-3-01-01' AND e.tipoDoc = 'FC' THEN 'CD'
        WHEN c.idCuenta = '1-1-1-01-01' THEN 'AA' 
        ELSE ''
    END AS [25],
   case when d.idCuenta = '1-1-1-02-08' then  '' else ISNULL(CAST(dis.idDTE AS VARCHAR),r.remate) end AS [26. idDTE],
    '' AS [27],
    '0' AS [28],
    '0' AS [29],
    '0' AS [30],
    '0' AS [31],
    '0' AS [32],
    '0' AS [33],
    '0' AS [34],
    '0' AS [35],
    '0' AS [36],
    '0' AS [37],
    '' AS [38],
    '' AS [39],
      MAX(e.numero) OVER () AS [40. Nro. agrupación en igual comprobante],
    'N' AS [41],
    'E' AS [42],
    'E' AS [43],
       CONCAT('EGRESO REMATE ', r.remate, ' NRO. ', MAX(e.numero) OVER ()) AS [44. Tipo de comprobante (I, E, T)]
FROM Egresos e
INNER JOIN egresosDet d ON d.idEgreso = e.idEgreso
INNER JOIN Cuentas c ON c.idCuenta = d.idCuenta
INNER JOIN remate r ON r.id = e.idRemate
LEFT JOIN egresosDetCheque edc ON edc.idEgresoDetalle = d.idEgresoDetalle
LEFT JOIN Distribucion dis ON dis.idEgresoDetalle = d.idEgresoDetalle
WHERE e.idEgreso IN ('" & idEgreso & "', '" & idEgreso2 & "')   AND c.idCuenta <> '1-1-6-01-02'
ORDER BY 
    CASE 
        WHEN c.descripcion = 'Efectivo' THEN 2
        WHEN c.descripcion LIKE 'Banco%' THEN 3
        ELSE 1
    END, 
    c.idCuenta DESC, 
    d.debito DESC;"

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
                                Using writer As New StreamWriter(destino)
                                    Dim lastNroDocumento As String = ""
                                    For i As Integer = 0 To objdataset.Tables(0).Rows.Count - 1
                                        Dim row As DataRow = objdataset.Tables(0).Rows(i)
                                        Dim values As New List(Of String)

                                        For j As Integer = 0 To row.ItemArray.Length - 1
                                            If j = 21 AndAlso row(1).ToString() = "1-1-6-01-02" Then
                                                values.Add(lastNroDocumento)
                                            Else
                                                values.Add(If(row(j) Is DBNull.Value, "", row(j).ToString()))
                                                If j = 21 AndAlso row(1).ToString() <> "1-1-6-01-02" Then
                                                    lastNroDocumento = If(row(j) Is DBNull.Value, "", row(j).ToString())
                                                End If
                                            End If
                                        Next

                                        Dim rowData As String = String.Join(",", values)
                                        writer.WriteLine(rowData)
                                    Next
                                    writer.Close()
                                End Using

                                updateCSVStatus(idEgreso)
                                updateCSVStatus(idEgreso2)
                                Return True
                            End If
                        Catch ex As Exception
                            MsgBox(ex.ToString)
                            retries += 1
                        End Try
                    Loop
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return False
    End Function


    Public Function GuardarCSVchequeAlaFecha(ByRef idEgreso As String, ByRef numero As Integer) As Boolean
        Try
            If ConectaBase() Then
                Dim sql As String = "WITH liquidaciones AS (
        SELECT 
            lf.idEgreso,
            lf.idDTE,
            lf.fecha_vencimiento,
            ROW_NUMBER() OVER (PARTITION BY lf.idEgreso ORDER BY lf.idDTE) AS RowNum,
            COUNT(*) OVER (PARTITION BY lf.idEgreso) AS TotalFacturas,
    		case when isnull(lf.porPagar,'') <> '' then lf.porPagar else (select lf.porPagar from LiquidacionFacturaElectronica as lf where lf.idEgreso = lf.idEgreso) end as porPagar
        FROM 
            LiquidacionFacturaElectronica lf
        WHERE 
            lf.idEgreso ='" & idEgreso & "'
			 UNION ALL

    SELECT 
        fc.idEgreso,
        fc.idDTE,
        fc.fecha_vencimiento,
        ROW_NUMBER() OVER (PARTITION BY fc.idEgreso ORDER BY fc.idDTE) AS RowNum,
        COUNT(*) OVER (PARTITION BY fc.idEgreso) AS TotalFacturas,
        CASE 
            WHEN ISNULL(fc.porPagar, '') <> '' THEN fc.porPagar 
            ELSE (SELECT fc.porPagar FROM FacturaCompra AS fc2 WHERE fc2.idEgreso = fc.idEgreso) 
        END AS porPagar
    FROM 
        FacturaCompra fc
    WHERE 
        fc.idEgreso = '" & idEgreso & "'

    ),
    Distribucion AS (
        SELECT 
            d.idEgresoDetalle,
            d.debito / f.TotalFacturas AS MontoDistribuido,
            f.idDTE,
            f.fecha_vencimiento,
            f.RowNum,
            f.TotalFacturas,
    		f.porPagar
        FROM 
            egresosDet d
        INNER JOIN liquidaciones f ON d.idEgreso = f.idEgreso
        WHERE 
            d.idCuenta = '2-1-3-01-01'
    )
    SELECT 
        '001' AS [1. AreaNegocio],
        c.idCuenta AS [2. idCuenta],
        CASE 
            WHEN d.idCuenta = '2-1-3-01-01' THEN dis.porPagar
            ELSE d.debito
        END AS [3. debito],
        d.credito AS [4. credito],
        CASE 
            WHEN c.idCuenta = '2-1-3-01-01' THEN CONCAT(e.tipoDoc, ' REM.:', r.remate)
            ELSE c.descripcion
        END AS [5. descripcion],
        '0' AS [6. EquivalenciaMoneda],
        '0' AS [7. Debe al debe Moneda],
        '' AS [8. Haber al haber Moneda],
        '' AS [9. Código Condición de Venta],
        '' AS [10. Código Vendedor],
        '' AS [11. Código Ubicación],
        '' AS [12. Código Concepto de Caja],
        '' AS [13. Código Instrumento Financiero],
        '' AS [14. Cantidad Instrumento Financiero],
        '' AS [15. Código Detalle de Gasto],
        '' AS [16. Cantidad Concepto de Gasto],
        '' AS [17. Código Centro de Costo],
        '' AS [18. Tipo Docto. Conciliación],
       case when d.idCuenta <> '2-1-3-01-01'then '' else '0' end AS [19. Nro. Docto. Conciliación],
        case when d.idCuenta = '2-1-2-01-09' THEN c.aux ELSE LEFT(e.rut, LEN(e.rut) - 2) END AS [20. Codigo Auxiliar],
        'PE'  AS [21. Tipo Documento],

       case when isnull(edc.numeroCheq,'') = '' then (select idc.numeroCheq from Egresos as e1 
		 inner join egresosDet as d on e1.idEgreso = d.idEgreso 
		 inner join egresosDetCheque as idc on idc.idEgresoDetalle = d.idEgresoDetalle
		 where e1.idEgreso = e.idEgreso) else edc.numeroCheq end AS [22. Nro. Documento]
		 
		 ,
        FORMAT(e.fechaHora, 'dd/MM/yyyy') AS [23. Fecha Emision],
       case WHEN  c.idCuenta = '2-1-2-01-09' then edc.fechaDeposito else   FORMAT(ISNULL(dis.fecha_vencimiento, e.fechaHora), 'dd/MM/yyyy') end AS [24. Fecha Vencimiento],
        CASE 
            WHEN  c.idCuenta = '2-1-3-01-01' and e.tipoDoc = 'LF' THEN 'CF'
            WHEN  c.idCuenta = '2-1-3-01-01' and e.tipoDoc = 'FC' THEN 'CD' else 'PE'
        END AS [25. Tipo Doc Soft],
        ISNULL(CAST(dis.idDTE AS VARCHAR), edc.numeroCheq) AS [26. idDTE],
        '' AS [27. Monto 1],
        '0' AS [28. Monto 2],
        '0' AS [29. Monto 4],
        '0' AS [30. Monto 5],
        '0' AS [31. Monto 6],
        '0' AS [32. Monto 7],
        '0' AS [33. Monto 8],
        '0' AS [34. Monto 9],
        '0' AS [35. Monto Suma Detalle Libro],
        '0' AS [36. Número Documento Desde],
        '0' AS [37. Número Documento Hasta],
        '' AS [38. Monto 10],
        '' AS [39. Monto 11],
        e.numero AS [40. Nro. agrupación en igual comprobante],
        'N' AS [41. NºOpe.Doc.Aux.(Considera este campo sólo si se definió en los parametros de CW)],
        'E' AS [42. Graba el detalle de libro (S/N)],
        'E' AS [43. Documento Nulo (S/N)],
        CONCAT('EGRESO REMATE ', r.remate, ' NRO. ', e.numero) AS [44. Tipo de comprobante (I, E, T)]
    FROM 
        Egresos e
    INNER JOIN 
        egresosDet d ON d.idEgreso = e.idEgreso
    INNER JOIN 
        Cuentas c ON c.idCuenta = d.idCuenta
    INNER JOIN 
        remate r ON r.id = e.idRemate
    LEFT JOIN 
        egresosDetCheque edc ON edc.idEgresoDetalle = d.idEgresoDetalle
    LEFT JOIN 
        Distribucion dis ON dis.idEgresoDetalle = d.idEgresoDetalle

    WHERE 
        e.idEgreso = '" & idEgreso & "'
    ORDER BY 
         c.idCuenta desc"

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
                    ' Intentar escribir el archivo con manejo de errores
                    Do While Not success And retries < maxRetries
                        Try

                            If fFile.Exists(sFilename) Then
                            Else

                                Using writer As New StreamWriter(destino)
                                    ' Recorrer las filas y escribir los datos en el archivo
                                    For Each row As DataRow In objdataset.Tables(0).Rows
                                        Dim rowData As String = String.Join(",", row.ItemArray.Select(Function(value) value.ToString()))
                                        writer.WriteLine(rowData)
                                    Next
                                    writer.Close()
                                End Using

                                updateCSVStatus(idEgreso)
                                Return True
                            End If

                        Catch ex As Exception
                            ' MsgBox(ex.ToString)
                        End Try
                    Loop
                End If
            End If
        Catch ex As Exception
            ' MsgBox(ex.ToString)
        End Try

        Return False
    End Function

    Public Function GuardarCSVanticipoPagoVenta(ByRef idEgreso As String, ByRef numero As Integer) As Boolean
        Try
            If ConectaBase() Then
                Dim sql As String = "WITH DatosEgreso AS (
    SELECT DISTINCT 
        f.rutVendedor,
        CONCAT('EGRESO REMATE ', (SELECT TOP 1 remate FROM remate WHERE id = f.idremate), ' NRO. ', 
               (SELECT top 1 numero FROM Egresos WHERE idEgreso = f.idEgreso)) AS tipoComprobante, e.numero
    FROM LiquidacionFacturaElectronica AS f
	inner join Egresos as e on e.idEgreso = f.idEgreso
    WHERE f.idEgreso = '" & idEgreso & "'
)

SELECT 
    '001' AS [1. AreaNegocio],
    '2-1-3-01-01' AS [2. idCuenta],
    f.porPagar AS [3. debito],
    '0' AS [4. credito],
    CONCAT('L.: ', f.idDTE) AS [5. descripcion],
    '0' AS [6. EquivalenciaMoneda],
    '0' AS [7. Debe al debe Moneda],
    '' AS [8. Haber al haber Moneda],
    '' AS [9. Código Condición de Venta],
    '' AS [10. Código Vendedor],
    '' AS [11. Código Ubicación],
    '' AS [12. Código Concepto de Caja],
    '' AS [13. Código Instrumento Financiero],
    '' AS [14. Cantidad Instrumento Financiero],
    '' AS [15. Código Detalle de Gasto],
    '' AS [16. Cantidad Concepto de Gasto],
    '' AS [17. Código Centro de Costo],
    '' AS [18. Tipo Docto. Conciliación],
    '0' AS [19. Nro. Docto. Conciliación],
    LEFT(f.rutVendedor, LEN(f.rutVendedor) - 2) AS [20. Codigo Auxiliar],
    'PE' AS [21. Tipo Documento],
    edc.numeroCheq AS [22. Nro. Documento],
    FORMAT(TRY_CONVERT(datetime, f.fecha_rem), 'dd/MM/yyyy') AS [23. Fecha Emision],
    FORMAT(TRY_CONVERT(datetime, f.fecha_vencimiento), 'dd/MM/yyyy') AS [24. Fecha Vencimiento],
    'CF' AS [25. Tipo Doc Soft],
    f.idDTE AS [26. idDTE],
    '' AS [27. Monto 1],
    '0' AS [28. Monto 2],
    '0' AS [29. Monto 4],
    '0' AS [30. Monto 5],
    '0' AS [31. Monto 6],
    '0' AS [32. Monto 7],
    '0' AS [33. Monto 8],
    '0' AS [34. Monto 9],
    '0' AS [35. Monto Suma Detalle Libro],
    '0' AS [36. Número Documento Desde],
    '0' AS [37. Número Documento Hasta],
    '' AS [38. Monto 10],
    '' AS [39. Monto 11],
    (SELECT TOP 1 numero FROM Egresos WHERE idEgreso = f.idEgreso) AS [40. Nro. agrupación en igual comprobante],
    'N' AS [41. NºOpe.Doc.Aux.(Considera este campo sólo si se definió en los parametros de CW)],
    'E' AS [42. Graba el detalle de libro (S/N)],
    'E' AS [43. Documento Nulo (S/N)],
    CONCAT('EGRESO REMATE ', (SELECT TOP 1 remate FROM remate WHERE id = f.idremate), ' NRO. ', 
           (SELECT TOP 1 numero FROM Egresos WHERE idEgreso = f.idEgreso)) AS [44. Tipo de comprobante (I, E, T)]
FROM 
    LiquidacionFacturaElectronica AS f 
	inner join Egresos as e on e.idEgreso = f.idEgreso
	inner join egresosDet as ed on ed.idEgreso = e.idEgreso
	inner join egresosDetCheque as edc on edc.idEgresoDetalle = ed.idEgresoDetalle
WHERE 
    f.idEgreso = '" & idEgreso & "'

UNION ALL

SELECT 
    '001' AS [1. AreaNegocio],
    '1-1-3-01-03' AS [2. idCuenta],
    '0' AS [3. debito],
    fe.porPagar AS [4. credito],
    CONCAT('FE.: ', fe.idDTE) AS [5. descripcion],
    '0' AS [6. EquivalenciaMoneda],
    '0' AS [7. Debe al debe Moneda],
    '' AS [8. Haber al haber Moneda],
    '' AS [9. Código Condición de Venta],
    '' AS [10. Código Vendedor],
    '' AS [11. Código Ubicación],
    '' AS [12. Código Concepto de Caja],
    '' AS [13. Código Instrumento Financiero],
    '' AS [14. Cantidad Instrumento Financiero],
    '' AS [15. Código Detalle de Gasto],
    '' AS [16. Cantidad Concepto de Gasto],
    '' AS [17. Código Centro de Costo],
    '' AS [18. Tipo Docto. Conciliación],
    '0' AS [19. Nro. Docto. Conciliación],
    LEFT(fe.rutComprador, LEN(fe.rutComprador) - 2) AS [20. Codigo Auxiliar],
    'TA' AS [21. Tipo Documento],
  (SELECT TOP 1 numero FROM DatosEgreso) AS [22. Nro. Documento],
    FORMAT(TRY_CONVERT(datetime, fe.fecha_rem), 'dd/MM/yyyy') AS [23. Fecha Emision],
    FORMAT(TRY_CONVERT(datetime, fe.fecha_vencimiento), 'dd/MM/yyyy') AS [24. Fecha Vencimiento],
    'VB' AS [25. Tipo Doc Soft],
    fe.idDTE AS [26. idDTE],
    '' AS [27. Monto 1],
    '0' AS [28. Monto 2],
    '0' AS [29. Monto 4],
    '0' AS [30. Monto 5],
    '0' AS [31. Monto 6],
    '0' AS [32. Monto 7],
    '0' AS [33. Monto 8],
    '0' AS [34. Monto 9],
    '0' AS [35. Monto Suma Detalle Libro],
    '0' AS [36. Número Documento Desde],
    '0' AS [37. Número Documento Hasta],
    '' AS [38. Monto 10],
    '' AS [39. Monto 11],
   (SELECT TOP 1 numero FROM DatosEgreso) AS [40. Nro. agrupación en igual comprobante],
    'N' AS [41. NºOpe.Doc.Aux.(Considera este campo sólo si se definió en los parametros de CW)],
    'E' AS [42. Graba el detalle de libro (S/N)],
    'E' AS [43. Documento Nulo (S/N)],
    (SELECT TOP 1 tipoComprobante FROM DatosEgreso) AS [44. Tipo de comprobante (I, E, T)]
FROM FacturaElectronica as fe 
INNER JOIN Ingresos as i ON fe.n_ingreso = i.idIngreso
INNER JOIN IngresosDet as id ON id.idIngreso = i.idIngreso
FULL JOIN tipoPago as tp ON tp.cuentaIngreso = id.codigo
WHERE fe.idRemate = " & id_remate & " 
    AND i.rutCliente IN (SELECT TOP 1 rutVendedor FROM DatosEgreso)
    AND tp.cuentaIngreso = 1110411

UNION ALL

SELECT 
    '001' AS [1. AreaNegocio],
    c.idCuenta AS [2. idCuenta],
    d.debito AS [3. debito],
    d.credito AS [4. credito],
    CASE 
        WHEN c.idCuenta = '2-1-3-01-01' THEN CONCAT(e.tipoDoc, ' REM.:', r.remate)
        ELSE c.descripcion 
    END AS [5. descripcion],
    '0' AS [6. EquivalenciaMoneda],
    '0' AS [7. Debe al debe Moneda],
    '' AS [8. Haber al haber Moneda],
    '' AS [9. Código Condición de Venta],
    '' AS [10. Código Vendedor],
    '' AS [11. Código Ubicación],
    '' AS [12. Código Concepto de Caja],
    '' AS [13. Código Instrumento Financiero],
    '' AS [14. Cantidad Instrumento Financiero],
    '' AS [15. Código Detalle de Gasto],
    '' AS [16. Cantidad Concepto de Gasto],
    '' AS [17. Código Centro de Costo],
    CASE 
        WHEN c.idCuenta <> '2-1-3-01-01' THEN 'PE' 
        ELSE '' 
    END AS [18. Tipo Docto. Conciliación],
    CASE
        WHEN c.idCuenta = '2-1-3-01-01' THEN '0'
        ELSE edc.numeroCheq
    END AS [19. Nro. Docto. Conciliación],
    CASE
        WHEN c.idCuenta = '2-1-3-01-01' THEN LEFT(e.rut, LEN(e.rut) - 2)
        ELSE ''
    END AS [20. Codigo Auxiliar],
    CASE
        WHEN c.idCuenta = '2-1-3-01-01' THEN 'PE'
        ELSE ''
    END AS [21. Tipo Documento],
    CASE
        WHEN c.idCuenta = '2-1-3-01-01' THEN 
            (SELECT TOP 1 edc.numeroCheq FROM egresosDet AS d1 
             FULL JOIN egresosDetCheque AS edc ON d1.idEgresoDetalle = edc.idEgresoDetalle
             WHERE d1.idEgreso = e.idEgreso AND ISNULL(edc.numeroCheq, '') <> '')
        ELSE ''
    END AS [22. Nro. Documento],
    FORMAT(TRY_CONVERT(datetime, e.fechaHora), 'dd/MM/yyyy') AS [23. Fecha Emision],
    CASE 
        WHEN e.tipoDoc = 'LF' THEN 
            FORMAT(TRY_CONVERT(datetime, (SELECT TOP 1 lf.fecha_vencimiento 
                                        FROM LiquidacionFacturaElectronica AS lf 
                                        WHERE lf.idEgreso = e.idEgreso)), 'dd/MM/yyyy')
        WHEN e.tipoDoc = 'FC' THEN 
            FORMAT(TRY_CONVERT(datetime, (SELECT TOP 1 fc.fecha_vencimiento 
                                        FROM FacturaCompra AS fc 
                                        WHERE fc.idEgreso = e.idEgreso)), 'dd/MM/yyyy')
        ELSE ''
    END AS [24. Fecha Vencimiento],
    CASE 
        WHEN e.tipoDoc = 'LF' AND d.idCuenta = '2-1-3-01-01' THEN 'CF'
        WHEN e.tipoDoc = 'FC' AND d.idCuenta = '2-1-3-01-01' THEN 'CD'
        ELSE '' 
    END AS [25. Tipo Doc Soft],
    CASE 
        WHEN e.tipoDoc = 'LF' AND d.idCuenta = '2-1-3-01-01' THEN 
            (SELECT TOP 1 lf.idDTE FROM LiquidacionFacturaElectronica AS lf WHERE lf.idEgreso = e.idEgreso)
        WHEN e.tipoDoc = 'FC' AND d.idCuenta = '2-1-3-01-01' THEN 
            (SELECT TOP 1 fc.idDTE FROM FacturaCompra AS fc WHERE fc.idEgreso = e.idEgreso)
        ELSE '' 
    END AS [26. idDTE],
    '' AS [27. Monto 1],
    '0' AS [28. Monto 2],
    '0' AS [29. Monto 4],
    '0' AS [30. Monto 5],
    '0' AS [31. Monto 6],
    '0' AS [32. Monto 7],
    '0' AS [33. Monto 8],
    '0' AS [34. Monto 9],
    '0' AS [35. Monto Suma Detalle Libro],
    '0' AS [36. Número Documento Desde],
    '0' AS [37. Número Documento Hasta],
    '' AS [38. Monto 10],
    '' AS [39. Monto 11],
    e.numero AS [40. Nro. agrupación en igual comprobante],
    'N' AS [41. NºOpe.Doc.Aux.(Considera este campo sólo si se definió en los parametros de CW)],
    'E' AS [42. Graba el detalle de libro (S/N)],
    'E' AS [43. Documento Nulo (S/N)],
    CONCAT('EGRESO REMATE ', r.remate, ' NRO. ', e.numero) AS [44. Tipo de comprobante (I, E, T)]
FROM 
    Egresos AS e 
INNER JOIN 
    egresosDet AS d ON d.idEgreso = e.idEgreso
INNER JOIN 
    Cuentas AS c ON c.idCuenta = d.idCuenta
INNER JOIN 
    remate AS r ON r.id = e.idRemate
FULL JOIN 
    egresosDetCheque AS edc ON edc.idEgresoDetalle = d.idEgresoDetalle
WHERE 
    e.idEgreso = '" & idEgreso & "' 
    AND c.idCuenta <> '2-1-3-01-01'"

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
                                    Dim lastNroDocumento As String = ""
                                    ' Recorrer las filas y escribir los datos en el archivo
                                    For i As Integer = 0 To objdataset.Tables(0).Rows.Count - 1
                                        Dim row As DataRow = objdataset.Tables(0).Rows(i)
                                        Dim values As New List(Of String)

                                        ' Procesar cada columna
                                        For j As Integer = 0 To row.ItemArray.Length - 1
                                            If j = 21 AndAlso row(1).ToString() = "1-1-6-01-02" Then
                                                ' Si es la columna 22 (índice 21) y la cuenta es 1-1-6-01-02
                                                values.Add(lastNroDocumento)
                                            Else
                                                values.Add(row(j).ToString())
                                                ' Guardar el número de documento si no es cuenta 1-1-6-01-02
                                                If j = 21 AndAlso row(1).ToString() <> "1-1-6-01-02" Then
                                                    lastNroDocumento = row(j).ToString()
                                                End If
                                            End If
                                        Next

                                        Dim rowData As String = String.Join(",", values)
                                        writer.WriteLine(rowData)
                                    Next
                                    writer.Close()
                                End Using
                                updateCSVStatus(idEgreso)
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
            ' MsgBox(ex.ToString)
        End Try

        Return False
    End Function

    Public Function verificarAnticipoRut(ByVal rut As String) As Boolean
        Try
            If ConectaBase() Then
                sql = "SELECT TOP 1 e.idEgreso 
                       FROM Egresos e 
                       WHERE e.rut = '" & rut & "' 
                       AND e.motivo = 'ANTICIPO CLIENTES' 
                       AND e.anulado = 0 
                       AND e.idRemate = " & id_remate & " 
                       ORDER BY e.fechaHora DESC"

                objadapter = New OleDbDataAdapter(sql, Conecta)
                objdataset = New DataSet
                objadapter.Fill(objdataset)



                If objdataset.Tables(0).Rows.Count > 0 Then

                    ' Egreso.idEgreso deeb ser igual al dato obteneido
                    Egreso.idEgreso = objdataset.Tables(0).Rows(0).Item(0).ToString()
                    Return True
                Else
                    Return False
                End If


            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return False
    End Function

    Public Function obtenerIdEgresoAnticipo(ByVal rut As String) As String
        Try
            If ConectaBase() Then
                sql = "SELECT TOP 1 e.idEgreso 
                       FROM Egresos e 
                       WHERE e.rut = '" & rut & "' 
                       AND e.motivo = 'ANTICIPO CLIENTES' 
                       AND e.anulado = 0 
                       AND e.idRemate = " & id_remate & " 
                       AND isnull(e.csv,0) = 0
                       ORDER BY e.fechaHora DESC"

                objadapter = New OleDbDataAdapter(sql, Conecta)
                objdataset = New DataSet
                objadapter.Fill(objdataset)

                If objdataset IsNot Nothing AndAlso
                   objdataset.Tables.Count > 0 AndAlso
                   objdataset.Tables(0).Rows.Count > 0 AndAlso
                   objdataset.Tables(0).Rows(0).Item(0) IsNot DBNull.Value Then

                    Dim idEgreso As String = objdataset.Tables(0).Rows(0).Item(0).ToString()
                    MsgBox("ID Egreso Anticipo encontrado: " & idEgreso) ' Debugging
                    Return idEgreso
                Else
                    MsgBox("No se encontró anticipo para el RUT: " & rut) ' Debugging
                End If
            End If
        Catch ex As Exception
            MsgBox("Error en obtenerIdEgresoAnticipo: " & ex.ToString)
        End Try
        Return Nothing
    End Function
End Class
