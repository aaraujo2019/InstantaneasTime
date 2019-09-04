Imports System.Data.SqlClient
Imports Microsoft.Office.Interop

Public Class FrmCinstantaneas

    'Agregado al repositorio
    '04/08/2019
    'Alvaro Araujo

    Dim nombreHoja As String
    Dim conn As New ADODB.Connection()
    Dim rstlab As New ADODB.Recordset()

    Dim Cn As New SqlConnection("Server=SEGSVRSQL01;uid=admonDb_planta;pwd=GcgPlanta2019*.;database=PlantaBeneficio")
    Dim rst As New ADODB.Recordset()
    Dim cnStr As String
    Private TiempoRestante As Integer

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Me.WindowState = FormWindowState.Minimized
            Call TiempoEjecutar(5)
            LblFecha.Text = DateTime.Now.ToString("dd/MM/yyyy")
            LblHora.Text = TimeString
        Catch ex As Exception
            Me.Close()
        End Try
    End Sub

    Public Sub TimerOn(ByRef Interval As Short)
        If Interval > 0 Then
            Timer1.Enabled = True
        Else
            Timer1.Enabled = False
        End If
    End Sub

    Public Function TiempoEjecutar(ByVal Tiempo As Integer)
        TiempoEjecutar = ""
        TiempoRestante = Tiempo  ' 1 minutos=60 segundos 
        Timer1.Interval = 1000
        Call TimerOn(1000) ' Hechanos a andar el timer
    End Function





    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        If TiempoRestante >= 0 Then
            LblEjecutar.Text = TiempoRestante
            TiempoRestante = TiempoRestante - 1
            Label1.Text = TimeString
        Else
            LblFecha.Text = DateTime.Now.ToString("dd/MM/yyyy")
            Dim exists As Boolean
            exists = System.IO.Directory.Exists("C:\instantaneas")
            If exists = True Then
                RecorreDirectorio()
                'TiempoRestante = 5
                'Call TiempoEjecutar(3600)
                'Timer1.Enabled = False
                'Ejecuta tu función cuando termina el tiempo
                'borrardirectorio()
            End If
            Me.Close()

        End If

    End Sub
    Private Sub borrardirectorio()
        Try
            ' lista todos los archivos dll del directorio windows _  
            ' SearchAllSubDirectories : incluye los Subdirectorios  
            ' SearchTopLevelOnly : para buscar solo en el nivel actual  
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  
            For Each Archivo As String In My.Computer.FileSystem.GetFiles(
                                    "C:\instantaneas",
                                    FileIO.SearchOption.SearchAllSubDirectories,
                                    "*.xlsx")
                'MsgBox(Archivo)
                My.Computer.FileSystem.DeleteFile(Archivo)
                'ListBox1.Items.Add(Archivo)
            Next
            ' errores  
        Catch oe As Exception
            MsgBox(oe.Message, MsgBoxStyle.Critical)
        End Try
    End Sub



    Private Function ValidaSiExiste(ByVal fecha_v As Date, ByVal hora2_v As String, ByVal ubicacion_v As String) As Boolean
        Try
            Using cnn As New SqlConnection("Server=SEGSVRSQL01;uid=admonDb_planta;pwd=GcgPlanta2019*.;database=PlantaBeneficio")
                Dim sqlbuscar As String = String.Format("SELECT COUNT(*) FROM PB_Instantaneas WHERE fecha = @fecha and ubicacion = @ubicacion and hora = @hora")
                Dim cmd As New SqlCommand(sqlbuscar, cnn)
                cmd.Parameters.AddWithValue("@fecha", fecha_v)
                cmd.Parameters.AddWithValue("@ubicacion", ubicacion_v)
                cmd.Parameters.AddWithValue("@hora", hora2_v)
                cnn.Open()
                Dim Count As Integer = CInt(cmd.ExecuteScalar())
                cnn.Close()
                If Count > 0 Then
                    Return True
                Else
                    Return False
                End If
            End Using
        Catch ex As Exception
            Throw
        End Try
    End Function

    Private Function ObtenerNombrePrimeraHoja(ByVal rutaLibro As String) As String
        Dim app As Excel.Application = Nothing
        Try
            app = New Excel.Application()
            Dim wb As Excel.Workbook = app.Workbooks.Open(rutaLibro)
            Dim ws As Excel.Worksheet = CType(wb.Worksheets.Item(1), Excel.Worksheet)
            Dim name As String = ws.Name
            ws = Nothing
            wb.Close()
            wb = Nothing
            Return name
        Catch ex As Exception
            Throw
        Finally
            If (Not app Is Nothing) Then _
                app.Quit()
            Runtime.InteropServices.Marshal.ReleaseComObject(app)
            app = Nothing
        End Try
    End Function

    Private Sub CargarArchivo(ByVal Ruta As String)
        Try
            Using cnn As New SqlConnection("Server=SEGSVRSQL01;uid=admonDb_planta;pwd=GcgPlanta2019*.;database=PlantaBeneficio")
                Dim AppExcel As Excel.Application
                Dim LibroExcel As Excel.Workbook
                Dim HojaExcel As Excel.Worksheet
                Dim celda As String

                Try
                    Dim FicheroExcel As String
                    Dim NombreHoja As String
                    'variables de insercion
                    Dim sqlConnectiondb As New System.Data.SqlClient.SqlConnection("Server=SEGSVRSQL01;uid=admonDb_planta;pwd=GcgPlanta2019*.;database=PlantaBeneficio")
                    Dim cmd As New System.Data.SqlClient.SqlCommand
                    cmd.CommandType = System.Data.CommandType.Text
                    FicheroExcel = Ruta
                    NombreHoja = ObtenerNombrePrimeraHoja(Ruta)
                    AppExcel = New Excel.Application
                    LibroExcel = AppExcel.Workbooks.Open(FicheroExcel)
                    HojaExcel = CType(LibroExcel.Sheets(NombreHoja), Excel.Worksheet)
                    Dim limite As Integer
                    limite = 200
                    Dim fecha, hora As Date
                    Dim Ordentrabajo, ubicacion, TipoMuestra As String
                    TipoMuestra = "Solución"
                    Ordentrabajo = Replace(Convert.ToString(HojaExcel.Range("C3").Value), " ", "")
                    fecha = CDate(Convert.ToString((HojaExcel.Range("A10").Value)))
                    '  Dim dou_hora As Double = 0.22916666666667
                    hora = Date.FromOADate(CDbl(HojaExcel.Range("B10").Value))
                    Dim HoraAux, hora2 As String

                    For i As Integer = 10 To 65
                        celda = "A" & i
                        Dim tenor As Double
                        If CStr(HojaExcel.Range("B" & i).Value) = "" Then
                            HoraAux = FormatDateTime(hora, DateFormat.ShortTime)
                            hora2 = (String.Format(HoraAux, "HH:mm:ss"))
                        Else
                            hora = Date.FromOADate(CDbl(HojaExcel.Range("B" & i).Value))
                            HoraAux = FormatDateTime(hora, DateFormat.ShortTime)
                            hora2 = (String.Format(HoraAux, "HH:mm:ss"))
                        End If

                        If CStr(HojaExcel.Range("d" & i).Value) = "'<0.01" Or CStr(HojaExcel.Range("d" & i).Value) = "<0.01" Then
                            tenor = 0.01
                        ElseIf IsNumeric(Replace(Convert.ToString(HojaExcel.Range("d" & i).Value), "<0.01", "0.01")) Then
                            tenor = CDbl(Replace(Convert.ToString(HojaExcel.Range("d" & i).Value), "<0.01", "0.01"))
                        Else
                            tenor = 0
                        End If
                        If Replace(Convert.ToString(HojaExcel.Range("C" & i).Value), " ", "") = "CM" Then
                            ubicacion = "Cabeza Merril Crowe"
                        ElseIf Replace(Convert.ToString(HojaExcel.Range("C" & i).Value), " ", "") = "CMC" Then
                            ubicacion = "Cola Merril Crowe"

                        ElseIf Replace(Convert.ToString(HojaExcel.Range("C" & i).Value), " ", "") = "CAG1" Then
                            ubicacion = "Cabeza Agitador 1"

                        ElseIf Replace(Convert.ToString(HojaExcel.Range("C" & i).Value), " ", "") = "DAG1" Then
                            ubicacion = "Descarga Agitador 1"

                        ElseIf Replace(Convert.ToString(HojaExcel.Range("C" & i).Value), " ", "") = "DAG2" Then
                            ubicacion = "Descarga Agitador 2"
                        ElseIf Replace(Convert.ToString(HojaExcel.Range("C" & i).Value), " ", "") = "DAG3" Then
                            ubicacion = "Descarga Agitador 3"

                        ElseIf Replace(Convert.ToString(HojaExcel.Range("C" & i).Value), " ", "") = "E-5(descarga)" Then
                            ubicacion = "Descarga Espesador 5"
                        Else
                            ubicacion = CStr(HojaExcel.Range("C" & i).Value)
                        End If

                        If CStr(HojaExcel.Range("d" & i).Value) = "" Or IsDBNull(HojaExcel.Range("d" & i).Value) Or IsNothing((HojaExcel.Range("d" & i).Value)) Or tenor <= 0 Then


                        Else

                            If (ValidaSiExiste(fecha, hora2, ubicacion)) Then
                                cmd.CommandText = "UPDATE  PB_Instantaneas  SET tenor = '" & tenor & "' WHERE fecha= '" & fecha & "'  and hora= '" & hora2 & "' and hora= '" & ubicacion & "'  "
                            Else
                                cmd.CommandText = "INSERT INTO PB_Instantaneas (Ordentrabajo,hora,ubicacion,tenor,fecha,TipoMuestra)VALUES('" & Ordentrabajo & "','" & hora2 & "','" & ubicacion & "' , '" & tenor & "' , '" & fecha & "' , '" & TipoMuestra & "' )"
                            End If
                            cmd.Connection = sqlConnectiondb
                            sqlConnectiondb.Open()
                            cmd.ExecuteNonQuery()
                            sqlConnectiondb.Close()

                            If CStr(HojaExcel.Range("D" & i).Value) = "" Then
                                Exit For
                            End If
                        End If
                    Next

                    LibroExcel.Close()
                    AppExcel.Quit()
                    AppExcel = Nothing
                    LibroExcel = Nothing

                    For Each clsProcess As Process In Process.GetProcesses()
                        If clsProcess.ProcessName.Equals("EXCEL") Then
                            clsProcess.Kill()
                        End If
                    Next

                Catch ex As Exception
                    ' Handle the exception.
                    MessageBox.Show(ex.Message, "Instantaneas", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End Using
        Catch ex As Exception
            'Throw
            If conn.State <> ConnectionState.Closed Then
                conn.Close()
            End If
        End Try
    End Sub

    Private Sub RecorreDirectorio()
        Try
            ' lista todos los archivos dll del directorio windows _  
            ' SearchAllSubDirectories : incluye los Subdirectorios  
            ' SearchTopLevelOnly : para buscar solo en el nivel actual  
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  
            For Each Archivo As String In My.Computer.FileSystem.GetFiles(
                                    "C:\instantaneas",
                                    FileIO.SearchOption.SearchAllSubDirectories,
                                                                       "*.xlsx")
                If DateDiff(DateInterval.Day, FileDateTime(Archivo), Date.Now) > 4 Then
                Else
                    CargarArchivo(Archivo)
                End If

            Next
            ' errores  
        Catch oe As Exception
            '  MsgBox(oe.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

End Class
