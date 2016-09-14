Public Class errgrid
    Inherits System.Web.UI.Page
    Public CON5 As New System.Data.OleDb.OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.MapPath("\App_Data\ERR\ERR.accdb") & ";Persist Security Info=False;")
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If CON5.State <> ConnectionState.Open Then CON5.Open()
        Dim PMR_AM_DA As New OleDb.OleDbDataAdapter("SELECT * FROM ERR", CON5)
        Dim PMR_AM_DT As New DataTable
        PMR_AM_DA.Fill(PMR_AM_DT)
        GridView1.DataSource = PMR_AM_DT
        GridView1.DataBind()
    End Sub

End Class