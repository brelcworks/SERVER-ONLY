Imports System.Data.SqlClient
Imports System.IO
Public Class errgrid
    Inherits System.Web.UI.Page
    Private CON_AM As New SqlConnection("Server=7d398a2f-1a2b-4338-bcc7-a66000a64b47.sqlserver.sequelizer.com;Database=db7d398a2f1a2b4338bcc7a66000a64b47;User ID=kjvstqndeaoallkm;Password=xsPrEXzwwVnd4TZxZ2Yag3qZbGjiipdL843dyHbK6AvazBnzikiGKxxCbWq7Nqoh;")
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    	Try
    		Dim line As String =""
    Dim table As DataTable = CreateTable()
 dim sr As New StreamReader(Server.MapPath("/App_Data/ERR/log.log"))
    	Do While Not SR.EndOfStream
                    Dim aryDataIn() = Split(SR.ReadLine, "~!~")
                    table.Rows.Add(aryDataIn)
                Loop
	sr.Close()
        GridView1.DataSource =Nothing
        GridView1.DataSource=table
        GridView1.DataBind()
    	Catch ex As Exception
    		MsgBox(ex.ToString)
    End Try
    End Sub
Private Function CreateTable() As DataTable
    Try
        Dim table As New DataTable()
        ' Declare DataColumn and DataRow variables.
        Dim column As DataColumn
        ' Create new DataColumn, set DataType, ColumnName
        ' and add to DataTable. 
        column = New DataColumn()
        column.DataType = System.Type.[GetType]("System.String")
        column.ColumnName = "EID"
        table.Columns.Add(column)
        ' Create second column.
        column = New DataColumn()
        column.DataType = Type.[GetType]("System.String")
        column.ColumnName = "ETIME"
        table.Columns.Add(column)
        column = New DataColumn()
        column.DataType = System.Type.[GetType]("System.String")
        column.ColumnName = "ERR"
        table.Columns.Add(column)
        Return table
    Catch ex As Exception
        Throw New Exception(ex.Message)
    End Try
End Function
End Class