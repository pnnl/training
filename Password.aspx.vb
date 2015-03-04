Imports System.Collections.Specialized
Public Class Password
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Password))
        Me.sqlUsers = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlDeleteCommand = New System.Data.SqlClient.SqlCommand
        Me.commonConn = New System.Data.SqlClient.SqlConnection
        Me.SqlInsertCommand = New System.Data.SqlClient.SqlCommand
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlUpdateCommand1 = New System.Data.SqlClient.SqlCommand
        Me.dsCommon = New System.Data.DataSet
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'sqlUsers
        '
        Me.sqlUsers.DeleteCommand = Me.SqlDeleteCommand
        Me.sqlUsers.InsertCommand = Me.SqlInsertCommand
        Me.sqlUsers.SelectCommand = Me.SqlSelectCommand1
        Me.sqlUsers.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_USER", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("USER_KEY", "USER_KEY"), New System.Data.Common.DataColumnMapping("USER_CODE", "USER_CODE"), New System.Data.Common.DataColumnMapping("LAST_NAME", "LAST_NAME"), New System.Data.Common.DataColumnMapping("FIRST_NAME", "FIRST_NAME"), New System.Data.Common.DataColumnMapping("MIDDLE_NAME", "MIDDLE_NAME"), New System.Data.Common.DataColumnMapping("SUFFIX_NAME", "SUFFIX_NAME"), New System.Data.Common.DataColumnMapping("AUTHENTICATION_CODE", "AUTHENTICATION_CODE"), New System.Data.Common.DataColumnMapping("LAST_CHANGE_AUTH_CD_DATE", "LAST_CHANGE_AUTH_CD_DATE"), New System.Data.Common.DataColumnMapping("PRIMARY_WORK_PHONE", "PRIMARY_WORK_PHONE"), New System.Data.Common.DataColumnMapping("EMAIL_ADDRESS", "EMAIL_ADDRESS"), New System.Data.Common.DataColumnMapping("CREATE_DATE", "CREATE_DATE"), New System.Data.Common.DataColumnMapping("USED_DATE", "USED_DATE"), New System.Data.Common.DataColumnMapping("CREATOR_USER_KEY", "CREATOR_USER_KEY"), New System.Data.Common.DataColumnMapping("ACTIVE_IND", "ACTIVE_IND"), New System.Data.Common.DataColumnMapping("TRAINING_IND", "TRAINING_IND"), New System.Data.Common.DataColumnMapping("COMPANY_ABBREV", "COMPANY_ABBREV"), New System.Data.Common.DataColumnMapping("NO_EMAIL_IND", "NO_EMAIL_IND"), New System.Data.Common.DataColumnMapping("HANFORD_ID", "HANFORD_ID")})})
        Me.sqlUsers.UpdateCommand = Me.SqlUpdateCommand1
        '
        'SqlDeleteCommand
        '
        Me.SqlDeleteCommand.CommandText = "DELETE FROM VARCONN.[SAT_USER] WHERE (([USER_KEY] = @Original_USER_KEY))"
        Me.SqlDeleteCommand.Connection = Me.commonConn
        Me.SqlDeleteCommand.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@Original_USER_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "USER_KEY", System.Data.DataRowVersion.Original, Nothing)})
        '
        'commonConn
        '
        Me.commonConn.FireInfoMessageEventOnUserErrors = False
        '
        'SqlInsertCommand
        '
        Me.SqlInsertCommand.CommandText = resources.GetString("SqlInsertCommand.CommandText")
        Me.SqlInsertCommand.Connection = Me.commonConn
        Me.SqlInsertCommand.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@USER_CODE", System.Data.SqlDbType.Int, 0, "USER_CODE"), New System.Data.SqlClient.SqlParameter("@LAST_NAME", System.Data.SqlDbType.VarChar, 0, "LAST_NAME"), New System.Data.SqlClient.SqlParameter("@FIRST_NAME", System.Data.SqlDbType.VarChar, 0, "FIRST_NAME"), New System.Data.SqlClient.SqlParameter("@MIDDLE_NAME", System.Data.SqlDbType.VarChar, 0, "MIDDLE_NAME"), New System.Data.SqlClient.SqlParameter("@SUFFIX_NAME", System.Data.SqlDbType.VarChar, 0, "SUFFIX_NAME"), New System.Data.SqlClient.SqlParameter("@AUTHENTICATION_CODE", System.Data.SqlDbType.VarChar, 0, "AUTHENTICATION_CODE"), New System.Data.SqlClient.SqlParameter("@LAST_CHANGE_AUTH_CD_DATE", System.Data.SqlDbType.DateTime, 0, "LAST_CHANGE_AUTH_CD_DATE"), New System.Data.SqlClient.SqlParameter("@PRIMARY_WORK_PHONE", System.Data.SqlDbType.VarChar, 0, "PRIMARY_WORK_PHONE"), New System.Data.SqlClient.SqlParameter("@EMAIL_ADDRESS", System.Data.SqlDbType.VarChar, 0, "EMAIL_ADDRESS"), New System.Data.SqlClient.SqlParameter("@CREATE_DATE", System.Data.SqlDbType.DateTime, 0, "CREATE_DATE"), New System.Data.SqlClient.SqlParameter("@USED_DATE", System.Data.SqlDbType.DateTime, 0, "USED_DATE"), New System.Data.SqlClient.SqlParameter("@CREATOR_USER_KEY", System.Data.SqlDbType.Int, 0, "CREATOR_USER_KEY"), New System.Data.SqlClient.SqlParameter("@ACTIVE_IND", System.Data.SqlDbType.Bit, 0, "ACTIVE_IND"), New System.Data.SqlClient.SqlParameter("@TRAINING_IND", System.Data.SqlDbType.Bit, 0, "TRAINING_IND"), New System.Data.SqlClient.SqlParameter("@COMPANY_ABBREV", System.Data.SqlDbType.VarChar, 0, "COMPANY_ABBREV"), New System.Data.SqlClient.SqlParameter("@NO_EMAIL_IND", System.Data.SqlDbType.Bit, 0, "NO_EMAIL_IND"), New System.Data.SqlClient.SqlParameter("@HANFORD_ID", System.Data.SqlDbType.VarChar, 0, "HANFORD_ID")})
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = resources.GetString("SqlSelectCommand1.CommandText")
        Me.SqlSelectCommand1.Connection = Me.commonConn
        '
        'SqlUpdateCommand1
        '
        Me.SqlUpdateCommand1.CommandText = resources.GetString("SqlUpdateCommand1.CommandText")
        Me.SqlUpdateCommand1.Connection = Me.commonConn
        Me.SqlUpdateCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@USER_CODE", System.Data.SqlDbType.Int, 0, "USER_CODE"), New System.Data.SqlClient.SqlParameter("@LAST_NAME", System.Data.SqlDbType.VarChar, 0, "LAST_NAME"), New System.Data.SqlClient.SqlParameter("@FIRST_NAME", System.Data.SqlDbType.VarChar, 0, "FIRST_NAME"), New System.Data.SqlClient.SqlParameter("@MIDDLE_NAME", System.Data.SqlDbType.VarChar, 0, "MIDDLE_NAME"), New System.Data.SqlClient.SqlParameter("@SUFFIX_NAME", System.Data.SqlDbType.VarChar, 0, "SUFFIX_NAME"), New System.Data.SqlClient.SqlParameter("@AUTHENTICATION_CODE", System.Data.SqlDbType.VarChar, 0, "AUTHENTICATION_CODE"), New System.Data.SqlClient.SqlParameter("@LAST_CHANGE_AUTH_CD_DATE", System.Data.SqlDbType.DateTime, 0, "LAST_CHANGE_AUTH_CD_DATE"), New System.Data.SqlClient.SqlParameter("@PRIMARY_WORK_PHONE", System.Data.SqlDbType.VarChar, 0, "PRIMARY_WORK_PHONE"), New System.Data.SqlClient.SqlParameter("@EMAIL_ADDRESS", System.Data.SqlDbType.VarChar, 0, "EMAIL_ADDRESS"), New System.Data.SqlClient.SqlParameter("@CREATE_DATE", System.Data.SqlDbType.DateTime, 0, "CREATE_DATE"), New System.Data.SqlClient.SqlParameter("@USED_DATE", System.Data.SqlDbType.DateTime, 0, "USED_DATE"), New System.Data.SqlClient.SqlParameter("@CREATOR_USER_KEY", System.Data.SqlDbType.Int, 0, "CREATOR_USER_KEY"), New System.Data.SqlClient.SqlParameter("@ACTIVE_IND", System.Data.SqlDbType.Bit, 0, "ACTIVE_IND"), New System.Data.SqlClient.SqlParameter("@TRAINING_IND", System.Data.SqlDbType.Bit, 0, "TRAINING_IND"), New System.Data.SqlClient.SqlParameter("@COMPANY_ABBREV", System.Data.SqlDbType.VarChar, 0, "COMPANY_ABBREV"), New System.Data.SqlClient.SqlParameter("@NO_EMAIL_IND", System.Data.SqlDbType.Bit, 0, "NO_EMAIL_IND"), New System.Data.SqlClient.SqlParameter("@HANFORD_ID", System.Data.SqlDbType.VarChar, 0, "HANFORD_ID"), New System.Data.SqlClient.SqlParameter("@Original_USER_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "USER_KEY", System.Data.DataRowVersion.Original, Nothing)})
        '
        'dsCommon
        '
        Me.dsCommon.DataSetName = "NewDataSet"
        Me.dsCommon.Locale = New System.Globalization.CultureInfo("en-US")
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected WithEvents Image1 As System.Web.UI.WebControls.Image
    Protected WithEvents lblPrevPassword As System.Web.UI.WebControls.Label
    Protected WithEvents txtPrevPassword As System.Web.UI.WebControls.TextBox
    Protected WithEvents hlpPrevPassword As System.Web.UI.WebControls.Image
    Protected WithEvents hlpNewPassword As System.Web.UI.WebControls.Image
    Protected WithEvents hlpConfirmPassword As System.Web.UI.WebControls.Image
    Protected WithEvents lblNewPassword As System.Web.UI.WebControls.Label
    Protected WithEvents txtNewPassword As System.Web.UI.WebControls.TextBox
    Protected WithEvents lblConfirmPassword As System.Web.UI.WebControls.Label
    Protected WithEvents txtConfirmPassword As System.Web.UI.WebControls.TextBox
    Protected WithEvents lblErrPrevPassword As System.Web.UI.WebControls.Label
    Protected WithEvents lblErrNewPassword As System.Web.UI.WebControls.Label
    Protected WithEvents lblErrConfirmPassword As System.Web.UI.WebControls.Label
    Protected WithEvents lblErrSummary As System.Web.UI.WebControls.Label
    Protected WithEvents sqlUsers As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents dsCommon As System.Data.DataSet
    Protected WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlUpdateCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents Td1 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected myUtility As New Utility
    Protected WithEvents btnSubmit As System.Web.UI.WebControls.Button
    Protected WithEvents bntReset As System.Web.UI.WebControls.Button
    Protected WithEvents btnCancel As System.Web.UI.WebControls.Button
    Protected WithEvents commonConn As System.Data.SqlClient.SqlConnection
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Friend WithEvents SqlDeleteCommand As System.Data.SqlClient.SqlCommand
    Friend WithEvents SqlInsertCommand As System.Data.SqlClient.SqlCommand
    Protected WithEvents JSAction As String

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

#Region "Basic"
    'loads the page
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load
        Trace.Warn("Loading Page")
        InitializeComponent()
        Session("JSAction") = ""

        'catch incoming alert messages
        alert(Session("Alert"))

        'clear session
        myUtility.clearSession()

        If (Session.IsNewSession) Then
            Session("Alert") = "Your session has timed out for security purposes.  Please log in again."
        End If

        'get anything on the request line
        myUtility.getRequest(Page.Request.Url.Query())

        'Put user code to initialize the page here
        If (Session("isAuthenticated") <> "True") Then
            Session("pageSel") = "Login"
            redirect("./login.aspx")
        Else
            Session("CurrentPage") = "./Password.aspx"
            Session("pageSel") = "Password"
        End If

        'Set the connection based on the machine
        myUtility.getConnection(Me.commonConn)

        'get the from page
        If Not (IsPostBack) Then
            'get the page we came from
            Session("FromPage") = getFromPage()

            'set up ability to check if user really wants to post the information on the page
            Me.btnSubmit.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
        End If

        Session("JSAction") = JSAction
    End Sub

    'handles error messages
    Public Sub alert(ByVal message As String)
        If (message <> "") Then
            JSAction = "<script type='text/javascript'>window.alert('" & message & "');</script>"
            Session("Alert") = ""
        End If
    End Sub

    'clears the error strings
    Private Sub clearErrors()
        Me.lblErrSummary.Text = ""
        Me.lblErrConfirmPassword.Text = ""
        Me.lblErrNewPassword.Text = ""
        Me.lblErrPrevPassword.Text = ""
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
#End Region

#Region "Page Action"
    'Reacts to the user performing some kind of selection and submission
    Private Sub btnSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Session("JSAction") = ""

        'Clear the error strings
        clearErrors()

        'Change the user's password
        Me.SqlSelectCommand1.Connection = Me.commonConn
        If (myUtility.findUserName(Me.SqlSelectCommand1, Me.sqlUsers, Me.dsCommon)) Then
            If (Not inputError()) Then
                If (savePassword()) Then
                    Session("Alert") = "Password Changed."
                    Session("isExpired") = "False"
                    redirect()
                Else
                    alert("There was an error in saving your new password.")
                End If
            End If
        Else
            If (Me.dsCommon.Tables("Users").Rows.Count > 0) Then
                Me.lblErrSummary.Text = "* There appears to be more than one record for you.  This is an application error.<br/>"
            Else
                Me.lblErrSummary.Text = "* You appear to have been removed from the database.<br/>"
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Session("JSAction") = ""

        'Clear the error strings
        clearErrors()

        'return to the page the user came from
        redirect()
    End Sub

    'Reacts to the user performing a reset
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntReset.Click
        Session("JSAction") = ""

        'Clear the error strings
        clearErrors()

        'Reset all of the input fields
        Me.txtPrevPassword.Text = ""
        Me.txtConfirmPassword.Text = ""
        Me.txtNewPassword.Text = ""
    End Sub
#End Region
    
#Region "Password"
    'determines if there are input errors and sets up the error text based on the input errors
    Private Function inputError() As Boolean
        Dim isError As Boolean = False
        Me.lblErrSummary.Text = ""

        'check for malicious code in the confirmation password
        If Not (myUtility.goodInput(Me.txtConfirmPassword)) Then
            Me.lblErrSummary.Text &= "* You may have entered malicious code in the confirmation password field.<br/>"
            Me.lblErrConfirmPassword.Text = "*"
            isError = True
        End If

        'check for malicious code in the new password
        If Not (myUtility.goodInput(Me.txtNewPassword)) Then
            Me.lblErrSummary.Text &= "* You may have entered malicious code in the new password field.<br/>"
            Me.lblErrNewPassword.Text = "*"
            isError = True
        End If

        'check for malicious code in the previous password
        If Not (myUtility.goodInput(Me.txtPrevPassword)) Then
            Me.lblErrSummary.Text &= "* You may have entered malicious code in the previous password field.<br/>"
            Me.lblErrPrevPassword.Text = "*"
            isError = True
        End If

        'Check if the new password and confirmed password match
        If (Me.txtNewPassword.Text <> Me.txtConfirmPassword.Text) Then
            Me.lblErrSummary.Text &= "* Your new password does not match your confirmation password.<br/>"
            Me.lblErrPrevPassword.Text = "*"
            Me.lblErrNewPassword.Text = "*"
            isError = True
        End If

        'Check to make sure the user input text into all fields
        If (Me.txtPrevPassword.Text = "") Then
            Me.lblErrSummary.Text &= "* You must enter your old password.<br/>"
            Me.lblErrPrevPassword.Text &= "*"
            isError = True
        End If

        If (Me.txtNewPassword.Text = "") Then
            Me.lblErrSummary.Text &= "* You must enter your new password.<br/>"
            Me.lblErrNewPassword.Text &= "*"
            isError = True
        End If

        If (Me.txtConfirmPassword.Text = "") Then
            Me.lblErrSummary.Text &= "* You must confirm your new password.<br/>"
            Me.lblErrConfirmPassword.Text = "*"
            isError = True
        End If

        'check to see if the current and new passwords match
        If (Me.txtNewPassword.Text = Me.txtPrevPassword.Text) Then
            Me.lblErrSummary.Text &= "* You cannot change your password to the same password.<br/>"
            Me.lblErrNewPassword.Text &= "*"
            isError = True
        End If

        'Check the password length
        If (Me.txtNewPassword.Text.Length < 8) Then
            Me.lblErrSummary.Text &= "* Your new password must be at least 8 characters in length.<br/>"
            Me.lblErrNewPassword.Text &= "*"
            isError = True
        End If

        'Check for combinations of 1Upper, 1Digit, 1Special Character.  
        'Only 2 of the three types are required to be a valid password
        Dim containsUpper As Boolean = False
        Dim containsDigit As Boolean = False
        Dim containsSpecial As Boolean = False
        Dim tempChar As Char
        For Each tempChar In Me.txtNewPassword.Text
            If (tempChar >= "A" And tempChar <= "Z") Then
                containsUpper = True
            ElseIf (tempChar >= "0" And tempChar <= "9") Then
                containsDigit = True
            ElseIf (tempChar < "a" Or tempChar > "z") Then
                containsSpecial = True
            End If
        Next
        If Not ((containsUpper And containsDigit) Or (containsUpper And containsSpecial) Or (containsDigit And containsSpecial)) Then
            Me.lblErrSummary.Text &= "* You must at least have a combination of any two of these three types: Upper Case Character, Digit, Special Character.<br/>"
            Me.lblErrNewPassword.Text &= "*"
            Me.lblErrConfirmPassword.Text &= "*"
            isError = True
        End If

        'hash current password against entered password
        If Not (myUtility.checkPassword(Me.txtPrevPassword.Text, Me.dsCommon)) Then
            Me.lblErrSummary.Text &= "* You entered an old password that does not match your stored password.<br/>"
            Me.lblErrPrevPassword.Text &= "*"
            isError = True
        End If

        Return isError
    End Function

    'Saves the user's password returning false if it fails
    Private Function savePassword() As Boolean
        Dim strPassword As String
        strPassword = myUtility.hashPassword(Me.txtNewPassword.Text)
        If (strPassword <> "") Then
            'save the user's password
            Try
                Me.dsCommon.Tables("Users").Rows(0).Item("AUTHENTICATION_CODE") = strPassword
                Me.dsCommon.Tables("Users").Rows(0).Item("LAST_CHANGE_AUTH_CD_DATE") = Now
                Me.SqlUpdateCommand1.CommandText = Me.SqlUpdateCommand1.CommandText.Replace("VARCONN.", myUtility.getExtension())
                Me.SqlUpdateCommand1.Connection = Me.commonConn
                Me.sqlUsers.Update(Me.dsCommon, "Users")
                Me.dsCommon.AcceptChanges()
            Catch ex As Exception
                Return False
            End Try
        Else
            Return False
        End If
        Return True
    End Function
#End Region
End Class
