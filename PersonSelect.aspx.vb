Imports System.Collections.Specialized
Public Class PersonSelect
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()

    End Sub
    Protected WithEvents lblIntroduction As System.Web.UI.WebControls.Label
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents btnCancel As System.Web.UI.WebControls.Button
    Protected WithEvents btnContinue As System.Web.UI.WebControls.Button
    Protected requestItems As Array
    Protected WithEvents imgTitle As System.Web.UI.WebControls.Image
    Protected WithEvents lbPersonList As System.Web.UI.WebControls.ListBox
    Protected myUtility As Utility = New Utility
    Protected WithEvents JSAction As String

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
    'loads the page
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load
        'catch incoming alert messages
        alert(Session("Alert"))

        'clear session
        myUtility.clearSession()

        If (Session.IsNewSession) Then
            Session("Alert") = "Your session has timed out for security purposes.  Please log in again."
        End If

        'get the from page
        If Not (IsPostBack) Then
            Session("FromPage") = getFromPage()
        End If

        'Put user code to initialize the page here
        If (Session("isAuthenticated") <> "True") Then
            Session("CurrentPage") = "./Login.aspx"
            Session("pageSel") = "Login"
            redirect(Session("CurrentPage"))
        ElseIf (Session("PersonTable") Is Nothing) Then
            redirect(Session("CurrentPage"))
        Else
            Session("CurrentPage") = "./PersonSelect.aspx"
            Session("pageSel") = CStr(Session("FromPage")).Substring(CStr(Session("FromPage")).LastIndexOf("/") + 1)
            Session("pageSel") = CStr(Session("pageSel")).Substring(0, CStr(Session("pageSel")).LastIndexOf("."))
            If (CStr(Session("pageSel")).ToLower = "pubrequest") Then
                Session("pageSel") = "Template"
            End If
        End If

        'if the user's password has expired send them to the password reset page
        If (Session("isExpired") = "True") Then
            Session("CurrentPage") = "./Password.aspx"
            Session("pageSel") = "Password"
            redirect(Session("CurrentPage"))
        End If

        'Check for user selection from the comments list if we're sure it's not a simple post-back
        If Not (IsPostBack) Then
            'get anything we can off the request line
            myUtility.getRequest(Page.Request.Url.Query())

            'the users that are stored in the session
            loadData()
        End If

        Session("JSAction") = JSAction
    End Sub

    'Set location and re-direct
    Private Sub redirect(Optional ByVal currentPage As String = "")
        If (currentPage = "") Then
            Session("CurrentPage") = Session("FromPage")
            Response.Redirect(Session("CurrentPage"))
        Else
            Session("CurrentPage") = currentPage
            Response.Redirect(Session("CurrentPage"))
        End If
    End Sub

    'gets the referring or from page and returns it
    Public Function getFromPage() As String
        Dim coll As NameValueCollection
        ' Load ServerVariable collection into NameValueCollection object.
        coll = Request.ServerVariables
        ' Get referring page.
        Return (coll.Item("HTTP_REFERER"))
    End Function

    'handles error messages
    Public Sub alert(ByVal message As String)
        If (message <> "") Then
            JSAction = "<script type='text/javascript'>window.alert('" & message & "');</script>"
            Session("Alert") = ""
        End If
    End Sub
#End Region

#Region "Load Data"
    'loads the user data
    Private Sub loadData()
        Try
            'get the people returned and populate the drop down list
            Dim dt As DataTable = Session("PersonTable")
            Me.lbPersonList.DataSource = dt
            Me.lbPersonList.DataTextField = "NameEmail"
            Me.lbPersonList.DataValueField = "hanford_id"
            Me.lbPersonList.DataBind()
        Catch ex As Exception
            Session("Alert") = "User name listing not loaded properly."
            redirect(Session("FromPage"))
        End Try
    End Sub
#End Region

#Region "Page Action"
    'canceling out of the person select and going back to wherever and let the destination know that the process was stopped
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Session("CurrentPage") = Session("FromPage")
        Session("pageSel") = Session("FromPage")
        Session("ProcessStopped") = "True"
        redirect(Session("CurrentPage"))
    End Sub

    'the user may or may not have selected someone.  If they chose not to select anyone halt this process and alert the user.
    'otherwise go back to wherever and let the page know the process is going forward.
    Private Sub btnContinue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnContinue.Click
        If (Me.lbPersonList.SelectedIndex >= 0) Then
            Session("CurrentPage") = Session("FromPage")
            Session("pageSel") = Session("FromPage")
            Session("ProcessStopped") = "False"
            Session("SelectedPersonRow") = Me.lbPersonList.SelectedIndex
            redirect(Session("CurrentPage"))
        Else
            alert("You did not select anyone.  You must select a user to continue or cancel to go back to the page you came from.")
        End If
    End Sub
#End Region
End Class
