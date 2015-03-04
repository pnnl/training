Public Class PersonInput
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()

    End Sub
    Protected myUtility As New Utility
    Protected WithEvents imgTitle As System.Web.UI.WebControls.Image
    Protected WithEvents lblIntroduction As System.Web.UI.WebControls.Label
    Protected WithEvents lbPersonList As System.Web.UI.WebControls.ListBox
    Protected WithEvents btnCancel As System.Web.UI.WebControls.Button
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents Td1 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents lblPrevPassword As System.Web.UI.WebControls.Label
    Protected WithEvents lblNewPassword As System.Web.UI.WebControls.Label
    Protected WithEvents lblConfirmPassword As System.Web.UI.WebControls.Label
    Protected WithEvents hlpFirstName As System.Web.UI.WebControls.Image
    Protected WithEvents txtLastName As System.Web.UI.WebControls.TextBox
    Protected WithEvents hlpLastName As System.Web.UI.WebControls.Image
    Protected WithEvents txtFirstName As System.Web.UI.WebControls.TextBox
    Protected WithEvents hlpMiddleInitial As System.Web.UI.WebControls.Image
    Protected WithEvents txtMiddleInitial As System.Web.UI.WebControls.TextBox
    Protected WithEvents RequiredFieldValidator1 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents RequiredFieldValidator2 As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents btnSubmit As System.Web.UI.WebControls.Button
    Protected WithEvents Email As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents FromPage As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents seqSurvey As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents intGroupID As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents currentServer As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected strHref As String = ""
    Protected WithEvents UserType As System.Web.UI.HtmlControls.HtmlInputHidden

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

#Region "Basic"
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load
        'Put user code to initialize the page here
        If Not (IsPostBack) Then
            'clear session
            myUtility.clearSession()

            'get everything off the request line
            myUtility.getRequest(Page.Request.Url.Query)

            'set the introduction text
            Me.lblIntroduction.Text = "Unable to locate any user information for " & _
                Session("PersonInputEmail") & ".  <p/>Please enter the following information " & _
                "and click submit to finish adding this user or click cancel to skip adding this user."

            'store the e-mail for return into a hidden field
            Me.Email.Value = Session("PersonInputEmail")
            Me.seqSurvey.Value = Session("seqSurvey")
            Me.intGroupID.Value = Session("intGroupID")
            Me.UserType.Value = Session("PersonInputUserType")
        End If
    End Sub
#End Region

#Region "Page Action"
    'cancels out of this pop-up window
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Response.Write("<script type='text/javascript'>window.close()</script>")
    End Sub

    'submits the information from the pop-up window back to the parent through session
    Private Sub btnSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Dim blnContinue As Boolean = True

        'check for possible malicious code entries
        If Not (myUtility.goodInput(Me.txtFirstName, True) And myUtility.goodInput(Me.txtLastName, True) _
                And myUtility.goodInput(Me.txtMiddleInitial, True)) Then
            blnContinue = False
            Session("Alert") = "Possible malicious code entry(s).  User input aborted."
            Response.Write("<script type='text/javascript'>window.opener.location.href = window.opener.location.pathname</script>")
            Response.Write("<script type='text/javascript'>window.close()</script>")
        End If

        If (blnContinue) Then
            Session("PersonInputFirstName") = Me.txtFirstName.Text
            Session("PersonInputLastName") = Me.txtLastName.Text
            Session("PersonInputMiddleInitial") = Me.txtMiddleInitial.Text
            Session("PersonInputEmailReturn") = Me.Email.Value
            Session("PersonInputUserType") = Me.UserType.Value
            Session("intGroupID") = Me.intGroupID.Value
            Session("seqSurvey") = Me.seqSurvey.Value
            Session("ProcessEmail") = "True"
            Response.Write("<script type='text/javascript'>window.opener.location.href = window.opener.location.pathname</script>")
            Response.Write("<script type='text/javascript'>window.close()</script>")
        End If
    End Sub
#End Region
End Class
