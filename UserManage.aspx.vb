Public Class UserManage
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents imgTitle As System.Web.UI.WebControls.Image
    Protected WithEvents hlpServerSwitch As System.Web.UI.WebControls.Image
    Protected WithEvents chkServerSwitch As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblPlaceHolder As System.Web.UI.WebControls.Label
    Protected WithEvents lblNavBar As System.Web.UI.WebControls.Label
    Protected WithEvents btnUp As System.Web.UI.WebControls.Button
    Protected WithEvents pnlNavBar As System.Web.UI.WebControls.Panel
    Protected WithEvents hlpUser As System.Web.UI.WebControls.Image
    Protected WithEvents lblUser As System.Web.UI.WebControls.Label
    Protected WithEvents ddlUser As System.Web.UI.WebControls.DropDownList
    Protected WithEvents hlpHIDPayrollName As System.Web.UI.WebControls.Image
    Protected WithEvents lblHIDPayrollName As System.Web.UI.WebControls.Label
    Protected WithEvents txtHIDPayrollName As System.Web.UI.WebControls.TextBox
    Protected WithEvents hlpEmail As System.Web.UI.WebControls.Image
    Protected WithEvents lblEmail As System.Web.UI.WebControls.Label
    Protected WithEvents txtEmail As System.Web.UI.WebControls.TextBox
    Protected WithEvents hlpFirstName As System.Web.UI.WebControls.Image
    Protected WithEvents lblFirstName As System.Web.UI.WebControls.Label
    Protected WithEvents txtFirstName As System.Web.UI.WebControls.TextBox
    Protected WithEvents hlpLastName As System.Web.UI.WebControls.Image
    Protected WithEvents lblLastName As System.Web.UI.WebControls.Label
    Protected WithEvents txtLastName As System.Web.UI.WebControls.TextBox
    Protected WithEvents hlpMiddleName As System.Web.UI.WebControls.Image
    Protected WithEvents lblMiddleName As System.Web.UI.WebControls.Label
    Protected WithEvents txtMiddleName As System.Web.UI.WebControls.TextBox
    Protected WithEvents hlpSuffixName As System.Web.UI.WebControls.Image
    Protected WithEvents lblSuffixName As System.Web.UI.WebControls.Label
    Protected WithEvents txtSuffixName As System.Web.UI.WebControls.TextBox
    Protected WithEvents hlpPhone As System.Web.UI.WebControls.Image
    Protected WithEvents lblPhoneNumber As System.Web.UI.WebControls.Label
    Protected WithEvents txtPhoneNumber As System.Web.UI.WebControls.TextBox
    Protected WithEvents hlpCompany As System.Web.UI.WebControls.Image
    Protected WithEvents lblCompany As System.Web.UI.WebControls.Label
    Protected WithEvents txtCompany As System.Web.UI.WebControls.TextBox
    Protected WithEvents hlpUserLevel As System.Web.UI.WebControls.Image
    Protected WithEvents lblUserAccessLevel As System.Web.UI.WebControls.Label
    Protected WithEvents ddlUserLevel As System.Web.UI.WebControls.DropDownList
    Protected WithEvents hlpPNNLTraining As System.Web.UI.WebControls.Image
    Protected WithEvents chkTQSurvey As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblTraining As System.Web.UI.WebControls.Label
    Protected WithEvents hlpActive As System.Web.UI.WebControls.Image
    Protected WithEvents chkActive As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblActive As System.Web.UI.WebControls.Label
    Protected WithEvents hlpOwner As System.Web.UI.WebControls.Image
    Protected WithEvents chkOwner As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lb As System.Web.UI.WebControls.Label
    Protected WithEvents btnHIDPayrollNameFind As System.Web.UI.WebControls.Button
    Protected WithEvents lblError As System.Web.UI.WebControls.Label
    Protected WithEvents ddlPageOptions As System.Web.UI.WebControls.DropDownList
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label

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
        Response.Redirect(Session("CurrentPage"))
    End Sub

End Class
