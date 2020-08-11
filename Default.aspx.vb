Imports System.Data.SqlClient
Imports System.IO
Imports System
Imports System.Drawing
Imports System.Data.OleDb
Imports System.Diagnostics
Imports VB = Microsoft.VisualBasic
Imports System.Data
Imports ExcelDataReader
Imports Z.Dapper.Plus
'Imports Microsoft.Office.Interop

Partial Class _Default
    Inherits Page
    Public file12 As String
    Dim f_id As Integer = 0
    Dim Rd As OleDbDataReader
    Dim Olp As New SqlConnection
    'Dim xl As New Excel.Application
    'Dim xlsheet As Excel.Worksheet
    'Dim xlwbook As Excel.Workbook
    Dim bb As String
    Dim Sheetname As String
    Dim tables As DataTableCollection
    Dim labels(100) As Label
    Dim item, item2, item3, item4 As Integer
    Dim cc As Integer


    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        'On Error GoTo Errorhandler1
        Try

            Panel2.Controls.Clear()
            Panel3.Controls.Clear()

            Dim aa As Integer = 0

            Do
                aa = aa + 1

                Try

                    bb = "dbo.Survey_data"
                    TruncateTable3()

                    GridView1.DataSource = Nothing
                    Panel2.Controls.Clear()

                    Using stream = File.Open(Server.MapPath("/Files Uploaded/" & Label2.Text), FileMode.Open, FileAccess.Read)
                        Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
                            Dim result As DataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
                                                             .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                                                             .UseHeaderRow = True}})
                            tables = result.Tables

                            Dim dt As DataTable = tables(ListBox1.Items(aa - 1).ToString)
                            GridView1.DataSource = dt
                            GridView1.DataBind()

                        End Using
                    End Using

                    cc = GridView1.Rows(0).Cells.Count

                    Dim connection As New Data.SqlClient.SqlConnection
                    Dim command As New Data.SqlClient.SqlCommand
                    Dim sql_result3 As SqlDataReader

                    command.CommandText = "SELECT COUNT(*) as 'ColumnCount'
  FROM INFORMATION_SCHEMA.COLUMNS
 WHERE table_catalog = 'OpDb' AND table_name = 'Survey_data';"
                    connection.ConnectionString = "Server=tcp:opserver.database.windows.net,1433;Database=OpDb;Uid=openport@opserver;Pwd=smallKi+e83;Encrypt=yes;TrustServerCertificate=no;Connection Timeout=0;"

                    connection.Open()
                    command.Connection = connection

                    sql_result3 = command.ExecuteReader

                    If sql_result3.HasRows Then

                        Do While sql_result3.Read()
                            Dim query_result3 As String
                            query_result3 = sql_result3("ColumnCount")

                            item = query_result3
                        Loop

                    End If

                    If Not item = cc Then

                        item = 0

                    End If

                    sql_result3.Close()

                    Dim str As String = New String("")
                    Dim sb As New System.Text.StringBuilder()


                    For iii = 1 To item

                        If iii = item Then

                            str = sb.Append("@Column" & iii).ToString

                        Else

                            str = sb.Append("@Column" & iii).Append(",").ToString

                        End If


                    Next

                    command = New SqlCommand("insert into " & bb & " values (" & str & ") ", connection)

                    For i = 1 To item

                        command.Parameters.Add("@Column" & i, SqlDbType.NVarChar)

                    Next


                    For i As Integer = 0 To GridView1.Rows.Count - 1

                        For ii = 0 To item - 1

                            If IsDBNull(GridView1.Rows(i).Cells(ii).Text) Then

                                command.Parameters(ii).Value = DBNull.Value

                            Else

                                command.Parameters(ii).Value = GridView1.Rows(i).Cells(ii).Text

                            End If

                        Next

                        command.ExecuteNonQuery()

                    Next

                    connection.Close()

                Catch ex As Exception

IncorrectFormat:
                    'Label2.Text = ex.Message
                    Label2.Text = "The file doesn't have the agreed upon format." + Environment.NewLine
                    Label2.ForeColor = Color.Red
                    GridView1.DataSource = Nothing
                    GridView1.DataBind()

                    Page.MaintainScrollPositionOnPostBack = False
                    Page.SetFocus(Label2)

                    GoTo Ending

                End Try

                Label2.Text = "Successfully imported to UAT." + Environment.NewLine
                Label2.ForeColor = Color.Blue
                    GridView1.DataSource = Nothing
                    GridView1.DataBind()
                    Page.MaintainScrollPositionOnPostBack = False
                Page.SetFocus(Label2)
                Button3.Enabled = True
                Button4.Enabled = True

Ending:

            Loop Until aa = ListBox1.Items.Count

            ListBox1.Items.Clear()
            Button1.Enabled = False

        Catch ex As Exception

            ListBox1.Items.Clear()
            Label2.Text = ex.Message
            Label2.ForeColor = Color.Red
            GridView1.DataSource = Nothing
            GridView1.DataBind()

            Button1.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False

        End Try

        Dim proc As System.Diagnostics.Process

        For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
            proc.Kill()
            proc.Close()
        Next

        'Response.Redirect("~/Default.aspx")
    End Sub

    Protected Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If FileUpload1.HasFile Then

            Try

                Button3.Enabled = False
                Button4.Enabled = False
                Panel2.Controls.Clear()

                FileUpload1.SaveAs(Server.MapPath("/Files Uploaded/" & FileUpload1.FileName))
                file12 = Server.MapPath("/Files Uploaded/" & FileUpload1.FileName)

                Label2.Text = FileUpload1.FileName
                Label2.ForeColor = Color.Blue

                Dim proc As System.Diagnostics.Process

                For Each proc In System.Diagnostics.Process.GetProcessesByName("EXCEL")
                    proc.Kill()
                    proc.Close()
                Next

                Using stream = File.Open(file12, FileMode.Open, FileAccess.Read)
                    Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
                        Dim result As DataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
                                                                 .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                                                                 .UseHeaderRow = True}})
                        tables = result.Tables
                        ListBox1.Items.Clear()

                        For Each table As DataTable In tables
                            ListBox1.Items.Add(table.TableName)
                        Next
                    End Using
                End Using

                Button1.Enabled = True

            Catch ex As Exception


                Label2.Text = "Select a file to upload first."
                Label2.ForeColor = Color.Red
                Page.MaintainScrollPositionOnPostBack = False
                Page.SetFocus(Label2)
                ListBox1.Items.Clear()
                GridView1.DataSource = Nothing
                GridView1.DataBind()
                Button3.Enabled = False
                Button4.Enabled = False

            End Try

        Else

            Label2.Text = "Select a file to upload first."
            Label2.ForeColor = Color.Red
            Page.MaintainScrollPositionOnPostBack = False
            Page.SetFocus(Label2)
            ListBox1.Items.Clear()
            GridView1.DataSource = Nothing
            GridView1.DataBind()
            Button3.Enabled = False
            Button4.Enabled = False

        End If

    End Sub

    Function FileOpenTest(ByVal WorkBookName As String) As Boolean
        Dim fs As FileStream
        FileOpenTest = False
        Try
            fs = IO.File.OpenWrite(WorkBookName)
            fs.Close()
        Catch ex As Exception
            FileOpenTest = True
        End Try

        Return FileOpenTest
    End Function

    Public Function TableExists1(ByVal tableName As String) As Boolean
        Dim conStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=NWIND.mdb"

        Using connection = New OleDbConnection(conStr)
            connection.Open()
            Dim tables = connection.GetSchema("Tables")
            Dim tableExists = False

            For i = 0 To tables.Rows.Count - 1
                tableExists = String.Equals(tables.Rows(i)(2).ToString(), tableName, StringComparison.CurrentCultureIgnoreCase)
                If tableExists Then Exit For
            Next

            Return tableExists
        End Using
    End Function


    Public Sub TruncateTable3()

        Try

            Dim sql_connection As New SqlConnection
            Dim sql_query As New SqlCommand
            sql_connection.ConnectionString = "Server=tcp:opserver.database.windows.net,1433;Database=OpDb;Uid=openport@opserver;Pwd=smallKi+e83;Encrypt=yes;TrustServerCertificate=no;Connection Timeout=0;"
            sql_query.Connection = sql_connection
            sql_connection.Open()
            sql_query.CommandText = "Truncate Table dbo.Survey_data;"
            sql_query.CommandTimeout = 0

            sql_query.ExecuteNonQuery()

        Catch ex As Exception

            Label2.Text = ""
            Dim lbl As Label = New Label
            lbl.Text = ex.Message
            Panel3.Controls.Add(lbl)

        End Try


    End Sub



    Public Sub TruncateTable4()

        Try

            Dim sql_connection As New SqlConnection
            Dim sql_query As New SqlCommand
            sql_connection.ConnectionString = "Server=tcp:opserver.database.windows.net,1433;Database=OpDb;Uid=openport@opserver;Pwd=smallKi+e83;Encrypt=yes;TrustServerCertificate=no;Connection Timeout=0;"
            sql_query.Connection = sql_connection
            sql_connection.Open()
            sql_query.CommandText = "Truncate Table dbo.File2;"
            sql_query.CommandTimeout = 0

            sql_query.ExecuteNonQuery()

        Catch ex As Exception

            Label2.Text = ""
            Dim lbl As Label = New Label
            lbl.Text = ex.Message
            Panel3.Controls.Add(lbl)

        End Try


    End Sub

    Public Sub TruncateTable5()

        Try

            Dim sql_connection As New SqlConnection
            Dim sql_query As New SqlCommand
            sql_connection.ConnectionString = "Server=tcp:opserver.database.windows.net,1433;Database=OpDb;Uid=openport@opserver;Pwd=smallKi+e83;Encrypt=yes;TrustServerCertificate=no;Connection Timeout=0;"
            sql_query.Connection = sql_connection
            sql_connection.Open()
            sql_query.CommandText = "Truncate Table dbo.File3;"
            sql_query.CommandTimeout = 0

            sql_query.ExecuteNonQuery()

        Catch ex As Exception

            Label2.Text = ""
            Dim lbl As Label = New Label
            lbl.Text = ex.Message
            Panel3.Controls.Add(lbl)

        End Try


    End Sub

    Protected Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged

        Try

            If ListBox1.SelectedValue.ToString = "" Then
                Dim lbl5 As Label = New Label
                lbl5.Text = "Please select Item in the Listbox."
                Panel2.Controls.Add(lbl5)

                GoTo Ending

            End If

            GridView1.DataSource = Nothing
            Panel2.Controls.Clear()

            Using stream = File.Open(Server.MapPath("/Files Uploaded/" & Label2.Text), FileMode.Open, FileAccess.Read)
                Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
                    Dim result As DataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
                                                                 .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                                                                 .UseHeaderRow = True}})
                    tables = result.Tables

                    Dim dt As DataTable = tables(ListBox1.SelectedValue.ToString)
                    dt.MinimumCapacity = 242433
                    GridView1.DataSource = dt
                    GridView1.DataBind()

                End Using
            End Using


Ending:
        Catch ex As Exception

            Label2.Text = ex.Message

        End Try

    End Sub

    Private Sub GridView1_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles GridView1.RowDataBound



    End Sub
    Protected Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim url As String = "https://uat-dispatch.insightscs.com/dashboard/?customer=PHTXVC&server=PH&token=664e7f875af557df67d022e99fa6e5f25d940fa0&report=55"

        Process.Start(url)

    End Sub
    Protected Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim sql_connection As New SqlConnection
        Dim sql_query As New SqlCommand
        sql_connection.ConnectionString = "Server=tcp:opserver.database.windows.net,1433;Database=OpDb;Uid=openport@opserver;Pwd=smallKi+e83;Encrypt=yes;TrustServerCertificate=no;Connection Timeout=0;"
        sql_query.Connection = sql_connection
        sql_connection.Open()
        sql_query.CommandText = "INSERT INTO dbo.Survey_Data_Prod SELECT * FROM dbo.Survey_Data;"
        sql_query.CommandTimeout = 0

        sql_query.ExecuteNonQuery()

        Label2.Text = "Successfully imported to production." + Environment.NewLine
        Label2.ForeColor = Color.Blue
        GridView1.DataSource = Nothing
        GridView1.DataBind()
        Page.MaintainScrollPositionOnPostBack = False
        Page.SetFocus(Label2)
        Button3.Enabled = False
        Button4.Enabled = False

    End Sub
End Class