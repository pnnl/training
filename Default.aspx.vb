Public Class Main
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents lblSearch As System.Web.UI.WebControls.Label
    Protected WithEvents txtSearch As System.Web.UI.WebControls.TextBox
    Protected WithEvents searchSubmit As System.Web.UI.WebControls.Button
    Protected WithEvents Td1 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents ContentIframe As System.Web.UI.HtmlControls.HtmlGenericControl

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here

        'Set the first page if not already set
        If Not IsPostBack Then
            If (Session("CurrentPage") Is Nothing) Then
                Session("CurrentPage") = "./Login.aspx"
            End If
        End If
    End Sub

    'Uses the PNNL search site to search for words typed in.
    Private Sub searchSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles searchSubmit.Click
        Dim queryString As String
        queryString = "http://amulet.pnl.gov/query.html?qp=url%3Ahttp%3A%2F%2Fwwwi.pnl.gov%2Fsafety%2F&qt="
        queryString += Me.txtSearch.Text
        Response.Redirect(queryString)
    End Sub
End Class
