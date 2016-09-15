Imports System.Data.SqlClient
Public Class errgrid
    Inherits System.Web.UI.Page
    Private CON_AM As New SqlConnection("Server=7d398a2f-1a2b-4338-bcc7-a66000a64b47.sqlserver.sequelizer.com;Database=db7d398a2f1a2b4338bcc7a66000a64b47;User ID=kjvstqndeaoallkm;Password=xsPrEXzwwVnd4TZxZ2Yag3qZbGjiipdL843dyHbK6AvazBnzikiGKxxCbWq7Nqoh;")
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If CON_AM.State <> ConnectionState.Open Then CON_AM.Open()
        Dim PMR_AM_DA As New SqlDataAdapter("SELECT * FROM ERR", CON_AM)
        Dim PMR_AM_DT As New DataTable
        PMR_AM_DA.Fill(PMR_AM_DT)
        GridView1.DataSource = PMR_AM_DT
        GridView1.DataBind()
    End Sub

End Class