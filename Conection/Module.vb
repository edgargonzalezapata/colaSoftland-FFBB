Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Text

Module Module1

    Public user, id_fletero, examen_global As Integer
    Public rut_cliente, lado, nom_form As String
    Public decomiso_global As Decimal
    Public ds As New DataSet
    Public da As OleDbDataAdapter
    'Permitir conectarnos con nuestro archivo de excel'
    Public conn As OleDbConnection
    Public id_decomiso, id_remate, num_remate, id_tiket, fk_boleto, fk_diio As Integer
    Public id_hoja_loteo As Integer = 0
    Public marca, can_ani, can_diio, num_crias As String
    Public fecha_remate As String
    Public idRecinto As Integer
    Public bloqueoRomana As Boolean
    'Permitir conectarnos a nuestra base de datos sqlserver'
    Public cnn As SqlConnection
    Public sqlBC As SqlBulkCopy
    Public Conecta, conecta2, conecta3 As OleDb.OleDbConnection

    'Public conexion_string As String = "Server=notebook;Initial Catalog=Personal_FBB;Persist Security Info=True;User ID=sa;Password=Fbb7346La"

    Public conexion_string As String = "Server=ffbb.database.windows.net;Initial Catalog=Personal_FBB;Persist Security Info=False;User ID=egonzalez;Password=Fbb7346La;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;language=spanish"
    Public recinto As String = "Feria Los Angeles"
    Public rup_origen As String = "08.3.01.9000"
    Public admin As Integer = 3500
    Public enunciado As SqlCommand
    Public respuesta As SqlDataReader
    'Conectar a la base de datos sqlserver'
    Sub abrirConexion()
        Try
            cnn = New SqlConnection(conexion_string)
            cnn.Open()

        Catch ex As Exception
            MessageBox.Show("NO SE CONECTO: " + ex.ToString)
        End Try
    End Sub




    Private Function GetExcelConnection(ByVal FilePath As String) As OleDbConnection
        Dim connStr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & FilePath & ";Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;'"
        Return New OleDbConnection(connStr)
    End Function

    Sub ExcelToSqlServer()
        Dim myFileDialog As New OpenFileDialog()
        Dim xSheet As String = ""

        With myFileDialog
            .Filter = "Excel Files |*.xlsx"
            .Title = "Open File"
            .ShowDialog()
        End With

        If myFileDialog.FileName.ToString <> "" Then
            Dim ExcelFile As String = myFileDialog.FileName.ToString
            xSheet = InputBox("Digite el nombre de la Hoja que desea importar", "Complete")
            conn = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;" & "data source=" & ExcelFile & "; " & "Extended Properties='Excel 12.0 Xml;HDR=Yes'")


            Try
                conn.Open()
                da = New OleDbDataAdapter("SELECT * FROM  [" & xSheet & "$]", conn)
                ds = New DataSet
                da.Fill(ds)

                sqlBC = New SqlBulkCopy(cnn)
                sqlBC.DestinationTableName = "Informe$"
                sqlBC.WriteToServer(ds.Tables(0))
            Catch ex As Exception
                MsgBox("Error: " + ex.ToString, MsgBoxStyle.Information, "Informacion")
            Finally
                conn.Close()
            End Try
        End If
        MsgBox("Datos Importados Correctamente")
    End Sub


    ' Se genera un objeto de conexión



    ' Se genera una función de conexión
    Public Function ConectaBase() As Boolean

        Try

            Conecta = New OleDb.OleDbConnection("Provider=SQLOLEDB;Data Source=ffbb.database.windows.net;Initial Catalog=Personal_FBB;Persist Security Info=True;User ID=egonzalez;Password=Fbb7346La;language=spanish")
            Return ConnectionState.Open


        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False

        End Try

    End Function

    Public Function ConectaBaseCierra() As Boolean

        Try

            Conecta = New OleDb.OleDbConnection("Provider=SQLOLEDB;Data Source=ffbb.database.windows.net;Initial Catalog=Personal_FBB;Persist Security Info=True;User ID=egonzalez;Password=Fbb7346La;language=spanish")
            Return ConnectionState.Closed


        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False

        End Try

    End Function
    Public Function ConectaBase2() As Boolean

        Try

            conecta3 = New OleDb.OleDbConnection("Provider=SQLOLEDB;Data Source=200.72.147.46;Initial Catalog=FBIOBIO;Persist Security Info=True;User ID=sa;Password=Fbb7346La")
            Return ConnectionState.Open

        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False

        End Try

    End Function

    Public Function ConectaBase3() As Boolean

        Try

            conecta2 = New OleDb.OleDbConnection("Provider=SQLOLEDB;Data Source=200.72.147.46;Initial Catalog=FBIOBIO;Persist Security Info=True;User ID=sa;Password=Fbb7346La")
            Return ConnectionState.Open

        Catch ex As Exception
            MsgBox(ex.ToString)
            Return False

        End Try

    End Function

    Public Function formatoRUT(ByRef rut As String) As String
        Dim rutOriginal As String
        rutOriginal = rut
        Try

            rut = rut.Replace("-", "")
            rut = String.Format("{0:00\.000\.000\-0}", Long.Parse(rut))

            Return rut
        Catch ex As Exception
            Return rutOriginal
        End Try

    End Function

    Public Function separadorMiles(ByRef numero As Integer) As String
        Return numero.ToString("N0")
    End Function

    Function formateoTexto(inputString As String) As String
        ' Validar que la cadena de entrada no sea nula
        If inputString Is Nothing Then
            ' Si la cadena es nula, puedes manejar el caso como desees.
            ' En este ejemplo, simplemente se devuelve una cadena vacía.
            Return ""
        End If

        ' Normalize the string to decomposed Unicode format
        Dim normalizedString As String = inputString.Normalize(NormalizationForm.FormD)

        ' Create a StringBuilder to store the result
        Dim resultBuilder As New System.Text.StringBuilder()

        ' Iterate through each character in the normalized string
        For Each c As Char In normalizedString
            ' If the character is a non-spacing mark or the degree symbol (°), skip it
            If Char.GetUnicodeCategory(c) <> UnicodeCategory.NonSpacingMark AndAlso c <> "°" Then
                ' Convert the character to uppercase and append it to the result
                resultBuilder.Append(Char.ToUpper(c))
            End If
        Next

        ' Return the final result
        Return resultBuilder.ToString()
    End Function






End Module