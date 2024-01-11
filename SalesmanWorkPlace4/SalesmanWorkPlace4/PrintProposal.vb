Imports CrystalDecisions.CrystalReports.Engine
Imports System.Data.SqlClient

Public Class PrintProposal
    Public PropNum As String

    Private Sub PrintProposal_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim cryRpt As New ReportDocument
        'cryRpt.Load("Proposal010300rus.rpt")
        cryRpt.Load(My.Settings.CRPath)
        'Dim ConnectionString = "Data Source=spbdvl2;Initial Catalog=ScaDataDB;User ID=sa;Password=sqladmin;"
        Dim ConnectionString = Replace(Declarations.MyConnStr, "Provider=SQLOLEDB;", "")
        Dim con As SqlConnection = New SqlConnection(ConnectionString)
        Dim cmd As SqlCommand = New SqlCommand("exec spp_SalesWorkplace4_PrepareProposal N'" & PropNum & "'", con)
        Dim adapter As New SqlDataAdapter(cmd)
        Dim dtset As New DataSet
        adapter.Fill(dtset, "Dataset1")
        cryRpt.SetDataSource(dtset.Tables.Item(0))

        CrystalReportViewer1.ReportSource = cryRpt
        CrystalReportViewer1.Zoom(1)
        CrystalReportViewer1.Refresh()
    End Sub

End Class