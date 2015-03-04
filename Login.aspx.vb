Imports System.Collections.Specialized
Public Class Login
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Login))
        Me.dsCommon = New System.Data.DataSet
        Me.sqlUsers = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlDeleteCommand1 = New System.Data.SqlClient.SqlCommand
        Me.commonConn = New System.Data.SqlClient.SqlConnection
        Me.SqlInsertCommand = New System.Data.SqlClient.SqlCommand
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlUpdateCommand1 = New System.Data.SqlClient.SqlCommand
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'dsCommon
        '
        Me.dsCommon.DataSetName = "NewDataSet"
        Me.dsCommon.Locale = New System.Globalization.CultureInfo("en-US")
        '
        'sqlUsers
        '
        Me.sqlUsers.DeleteCommand = Me.SqlDeleteCommand1
        Me.sqlUsers.InsertCommand = Me.SqlInsertCommand
        Me.sqlUsers.SelectCommand = Me.SqlSelectCommand1
        Me.sqlUsers.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_USER", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("USER_KEY", "USER_KEY"), New System.Data.Common.DataColumnMapping("USER_CODE", "USER_CODE"), New System.Data.Common.DataColumnMapping("LAST_NAME", "LAST_NAME"), New System.Data.Common.DataColumnMapping("FIRST_NAME", "FIRST_NAME"), New System.Data.Common.DataColumnMapping("MIDDLE_NAME", "MIDDLE_NAME"), New System.Data.Common.DataColumnMapping("SUFFIX_NAME", "SUFFIX_NAME"), New System.Data.Common.DataColumnMapping("AUTHENTICATION_CODE", "AUTHENTICATION_CODE"), New System.Data.Common.DataColumnMapping("PRIMARY_WORK_PHONE", "PRIMARY_WORK_PHONE"), New System.Data.Common.DataColumnMapping("EMAIL_ADDRESS", "EMAIL_ADDRESS"), New System.Data.Common.DataColumnMapping("CREATE_DATE", "CREATE_DATE"), New System.Data.Common.DataColumnMapping("USED_DATE", "USED_DATE"), New System.Data.Common.DataColumnMapping("CREATOR_USER_KEY", "CREATOR_USER_KEY"), New System.Data.Common.DataColumnMapping("ACTIVE_IND", "ACTIVE_IND"), New System.Data.Common.DataColumnMapping("TRAINING_IND", "TRAINING_IND"), New System.Data.Common.DataColumnMapping("COMPANY_ABBREV", "COMPANY_ABBREV"), New System.Data.Common.DataColumnMapping("NO_EMAIL_IND", "NO_EMAIL_IND"), New System.Data.Common.DataColumnMapping("HANFORD_ID", "HANFORD_ID"), New System.Data.Common.DataColumnMapping("INTERNET_PROVIDER_ADDRESS", "INTERNET_PROVIDER_ADDRESS")})})
        Me.sqlUsers.UpdateCommand = Me.SqlUpdateCommand1
        '
        'SqlDeleteCommand1
        '
        Me.SqlDeleteCommand1.CommandText = "DELETE FROM VARCONN.[SAT_USER] WHERE (([USER_KEY] = @Original_USER_KEY))" & _
            ""
        Me.SqlDeleteCommand1.Connection = Me.commonConn
        Me.SqlDeleteCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@Original_USER_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "USER_KEY", System.Data.DataRowVersion.Original, Nothing)})
        '
        'commonConn
        '
        Me.commonConn.FireInfoMessageEventOnUserErrors = False
        '
        'SqlInsertCommand
        '
        Me.SqlInsertCommand.CommandText = "INSERT INTO VARCONN.[SAT_USER] ([USER_CODE], [LAST_NAME], [FIRST_NAME], [MIDDLE_NAME], [SUFFIX_NAME], [AUTHENTICATION_CODE], [PRIMARY_WORK_PHONE], [EMAIL_ADDRESS], [CREATE_DATE], [USED_DATE], [CREATOR_USER_KEY], [ACTIVE_IND], [TRAINING_IND], [COMPANY_ABBREV], [NO_EMAIL_IND], [HANFORD_ID], [INTERNET_PROVIDER_ADDRESS]) VALUES (@USER_CODE, @LAST_NAME, @FIRST_NAME, @MIDDLE_NAME, @SUFFIX_NAME, @AUTHENTICATION_CODE, @PRIMARY_WORK_PHONE, @EMAIL_ADDRESS, @CREATE_DATE, @USED_DATE, @CREATOR_USER_KEY, @ACTIVE_IND, @TRAINING_IND, @COMPANY_ABBREV, @NO_EMAIL_IND, @HANFORD_ID, @INTERNET_PROVIDER_ADDRESS)"
        Me.SqlInsertCommand.Connection = Me.commonConn
        Me.SqlInsertCommand.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@USER_CODE", System.Data.SqlDbType.Int, 0, "USER_CODE"), New System.Data.SqlClient.SqlParameter("@LAST_NAME", System.Data.SqlDbType.VarChar, 0, "LAST_NAME"), New System.Data.SqlClient.SqlParameter("@FIRST_NAME", System.Data.SqlDbType.VarChar, 0, "FIRST_NAME"), New System.Data.SqlClient.SqlParameter("@MIDDLE_NAME", System.Data.SqlDbType.VarChar, 0, "MIDDLE_NAME"), New System.Data.SqlClient.SqlParameter("@SUFFIX_NAME", System.Data.SqlDbType.VarChar, 0, "SUFFIX_NAME"), New System.Data.SqlClient.SqlParameter("@AUTHENTICATION_CODE", System.Data.SqlDbType.VarChar, 0, "AUTHENTICATION_CODE"), New System.Data.SqlClient.SqlParameter("@PRIMARY_WORK_PHONE", System.Data.SqlDbType.VarChar, 0, "PRIMARY_WORK_PHONE"), New System.Data.SqlClient.SqlParameter("@EMAIL_ADDRESS", System.Data.SqlDbType.VarChar, 0, "EMAIL_ADDRESS"), New System.Data.SqlClient.SqlParameter("@CREATE_DATE", System.Data.SqlDbType.DateTime, 0, "CREATE_DATE"), New System.Data.SqlClient.SqlParameter("@USED_DATE", System.Data.SqlDbType.DateTime, 0, "USED_DATE"), New System.Data.SqlClient.SqlParameter("@CREATOR_USER_KEY", System.Data.SqlDbType.Int, 0, "CREATOR_USER_KEY"), New System.Data.SqlClient.SqlParameter("@ACTIVE_IND", System.Data.SqlDbType.Bit, 0, "ACTIVE_IND"), New System.Data.SqlClient.SqlParameter("@TRAINING_IND", System.Data.SqlDbType.Bit, 0, "TRAINING_IND"), New System.Data.SqlClient.SqlParameter("@COMPANY_ABBREV", System.Data.SqlDbType.VarChar, 0, "COMPANY_ABBREV"), New System.Data.SqlClient.SqlParameter("@NO_EMAIL_IND", System.Data.SqlDbType.Bit, 0, "NO_EMAIL_IND"), New System.Data.SqlClient.SqlParameter("@HANFORD_ID", System.Data.SqlDbType.VarChar, 0, "HANFORD_ID"), New System.Data.SqlClient.SqlParameter("@INTERNET_PROVIDER_ADDRESS", System.Data.SqlDbType.VarChar, 0, "INTERNET_PROVIDER_ADDRESS")})
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = "SELECT USER_KEY, USER_CODE, LAST_NAME, FIRST_NAME, MIDDLE_NAME, SUFFIX_NAME, AUTHENTICATION_CODE, PRIMARY_WORK_PHONE, EMAIL_ADDRESS, CREATE_DATE, USED_DATE, CREATOR_USER_KEY, ACTIVE_IND, TRAINING_IND, COMPANY_ABBREV, NO_EMAIL_IND, HANFORD_ID, INTERNET_PROVIDER_ADDRESS FROM VARCONN.SAT_USER"
        Me.SqlSelectCommand1.Connection = Me.commonConn
        '
        'SqlUpdateCommand1
        '
        Me.SqlUpdateCommand1.CommandText = "UPDATE VARCONN.[SAT_USER] SET [USER_CODE] = @USER_CODE, [LAST_NAME] = @LAST_NAME, [FIRST_NAME] = @FIRST_NAME, [MIDDLE_NAME] = @MIDDLE_NAME, [SUFFIX_NAME] = @SUFFIX_NAME, [AUTHENTICATION_CODE] = @AUTHENTICATION_CODE, [PRIMARY_WORK_PHONE] = @PRIMARY_WORK_PHONE, [EMAIL_ADDRESS] = @EMAIL_ADDRESS, [CREATE_DATE] = @CREATE_DATE, [USED_DATE] = @USED_DATE, [CREATOR_USER_KEY] = @CREATOR_USER_KEY, [ACTIVE_IND] = @ACTIVE_IND, [TRAINING_IND] = @TRAINING_IND, [COMPANY_ABBREV] = @COMPANY_ABBREV, [NO_EMAIL_IND] = @NO_EMAIL_IND, [HANFORD_ID] = @HANFORD_ID, [INTERNET_PROVIDER_ADDRESS] = @INTERNET_PROVIDER_ADDRESS WHERE (([USER_KEY] = @Original_USER_KEY))"
        Me.SqlUpdateCommand1.Connection = Me.commonConn
        Me.SqlUpdateCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@USER_CODE", System.Data.SqlDbType.Int, 0, "USER_CODE"), New System.Data.SqlClient.SqlParameter("@LAST_NAME", System.Data.SqlDbType.VarChar, 0, "LAST_NAME"), New System.Data.SqlClient.SqlParameter("@FIRST_NAME", System.Data.SqlDbType.VarChar, 0, "FIRST_NAME"), New System.Data.SqlClient.SqlParameter("@MIDDLE_NAME", System.Data.SqlDbType.VarChar, 0, "MIDDLE_NAME"), New System.Data.SqlClient.SqlParameter("@SUFFIX_NAME", System.Data.SqlDbType.VarChar, 0, "SUFFIX_NAME"), New System.Data.SqlClient.SqlParameter("@AUTHENTICATION_CODE", System.Data.SqlDbType.VarChar, 0, "AUTHENTICATION_CODE"), New System.Data.SqlClient.SqlParameter("@PRIMARY_WORK_PHONE", System.Data.SqlDbType.VarChar, 0, "PRIMARY_WORK_PHONE"), New System.Data.SqlClient.SqlParameter("@EMAIL_ADDRESS", System.Data.SqlDbType.VarChar, 0, "EMAIL_ADDRESS"), New System.Data.SqlClient.SqlParameter("@CREATE_DATE", System.Data.SqlDbType.DateTime, 0, "CREATE_DATE"), New System.Data.SqlClient.SqlParameter("@USED_DATE", System.Data.SqlDbType.DateTime, 0, "USED_DATE"), New System.Data.SqlClient.SqlParameter("@CREATOR_USER_KEY", System.Data.SqlDbType.Int, 0, "CREATOR_USER_KEY"), New System.Data.SqlClient.SqlParameter("@ACTIVE_IND", System.Data.SqlDbType.Bit, 0, "ACTIVE_IND"), New System.Data.SqlClient.SqlParameter("@TRAINING_IND", System.Data.SqlDbType.Bit, 0, "TRAINING_IND"), New System.Data.SqlClient.SqlParameter("@COMPANY_ABBREV", System.Data.SqlDbType.VarChar, 0, "COMPANY_ABBREV"), New System.Data.SqlClient.SqlParameter("@NO_EMAIL_IND", System.Data.SqlDbType.Bit, 0, "NO_EMAIL_IND"), New System.Data.SqlClient.SqlParameter("@HANFORD_ID", System.Data.SqlDbType.VarChar, 0, "HANFORD_ID"), New System.Data.SqlClient.SqlParameter("@INTERNET_PROVIDER_ADDRESS", System.Data.SqlDbType.VarChar, 0, "INTERNET_PROVIDER_ADDRESS"), New System.Data.SqlClient.SqlParameter("@Original_USER_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "USER_KEY", System.Data.DataRowVersion.Original, Nothing)})
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected WithEvents lblWelcomeText As System.Web.UI.WebControls.Label
    Protected WithEvents lblNoteText As System.Web.UI.WebControls.Label
    Protected WithEvents lblLoginText As System.Web.UI.WebControls.Label
    Protected WithEvents lblUserName As System.Web.UI.WebControls.Label
    Protected WithEvents txtUserName As System.Web.UI.WebControls.TextBox
    Protected WithEvents lblPassword As System.Web.UI.WebControls.Label
    Protected WithEvents txtPassword As System.Web.UI.WebControls.TextBox
    Protected WithEvents Image1 As System.Web.UI.WebControls.Image
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents lblPNNLText As System.Web.UI.WebControls.Label
    Protected WithEvents lblMailText As System.Web.UI.WebControls.Label
    Protected WithEvents btnSubmit As System.Web.UI.WebControls.Button
    Protected WithEvents btnReset As System.Web.UI.WebControls.Button
    Protected WithEvents dsCommon As System.Data.DataSet
    Protected WithEvents sqlUsers As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlUpdateCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlDeleteCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents txtSearch As System.Web.UI.WebControls.TextBox
    Protected WithEvents searchSubmit As System.Web.UI.WebControls.Button
    Protected WithEvents Td1 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected myUtility As New Utility
    Protected WithEvents lblDisclaimerText As System.Web.UI.WebControls.Label
    Protected WithEvents lblDisclaimer As System.Web.UI.WebControls.Label
    Protected WithEvents commonConn As System.Data.SqlClient.SqlConnection
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

        'Catch any alert coming in
        alert(Session("Alert"))

        'clear session
        myUtility.clearSession()

        If Not (IsPostBack) Then
            Session.Clear()
            Session("FromPage") = "./Login.aspx"
            Session("CurrentPage") = "./Login.aspx"
            Session("Machine") = "Development"
            Session("isAuthenticated") = "False"
            Session("UserType") = 0
        End If

        Session("JSAction") = JSAction

        'set the connections
        myUtility.getConnection(Me.commonConn)
        Me.SqlDeleteCommand1.Connection = Me.commonConn
        Me.SqlInsertCommand.Connection = Me.commonConn
        Me.SqlSelectCommand1.Connection = Me.commonConn
        Me.SqlUpdateCommand1.Connection = Me.commonConn
    End Sub

    'handles error messages
    Public Sub alert(ByVal message As String)
        If (message <> "") Then
            JSAction = "<script type='text/javascript'>window.alert('" & message & "');</script>"
            Session("Alert") = ""
        End If
    End Sub
#End Region

#Region "User Data"
    'Updates the changes to the record including the usage date
    Private Function updateUsageDate() As Boolean
        Trace.Warn("Updating Usage Date")
        Try
            Me.dsCommon.Tables("Users").Rows(0).Item("USED_DATE") = DateTime.Now
            Me.SqlUpdateCommand1.CommandText = Me.SqlUpdateCommand1.CommandText.Replace("VARCONN.", myUtility.getExtension())
            Me.sqlUsers.Update(Me.dsCommon, "Users")
            Me.dsCommon.AcceptChanges()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'Sees if there any users at all
    Private Function findUsers() As Boolean
        Trace.Warn("Seeing if there are any users at all")
        Try
            Me.SqlSelectCommand1.CommandText = Me.SqlSelectCommand1.CommandText.Remove(Me.SqlSelectCommand1.CommandText.IndexOf(" where"), Me.SqlSelectCommand1.CommandText.Length - Me.SqlSelectCommand1.CommandText.IndexOf(" where"))
            Me.SqlSelectCommand1.CommandText = Me.SqlSelectCommand1.CommandText.Replace("VARCONN.", myUtility.getExtension())
            Me.sqlUsers.Fill(Me.dsCommon, "Users")
            If (Me.dsCommon.Tables("Users").Rows.Count() > 0) Then
                Return True
            End If
            Return False
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'Generates the administrator
    Private Function generateAdmin() As Boolean
        Trace.Warn("Generating Administrator")
        Try
            Dim strUserName As String = UCase("survey.administrator@pnl.gov")
            Dim strTempPassword As String = "ILikeCheap2"
            Dim strPassword As String = ""

            If ((strPassword = myUtility.hashPassword(strTempPassword, strUserName)) <> "") Then
                'Generate the administrator and add to the database
                Dim dr As DataRow = Me.dsCommon.Tables("Users").NewRow()
                dr.Item("USER_CODE") = 4
                dr.Item("LAST_NAME") = "Administrator"
                dr.Item("FIRST_NAME") = "Survey"
                dr.Item("MIDDLE_NAME") = ""
                dr.Item("SUFFIX_NAME") = ""
                dr.Item("AUTHENTICATION_CODE") = strPassword
                dr.Item("PRIMARY_WORK_PHONE") = ""
                dr.Item("HANFORD_ID") = ""
                dr.Item("EMAIL_ADDRESS") = "survey.administrator@pnl.gov"
                dr.Item("CREATE_DATE") = Now
                dr.Item("USED_DATE") = Now
                dr.Item("CREATOR_USER_KEY") = 1
                dr.Item("ACTIVE_IND") = 1
                dr.Item("TRAINING_IND") = 1
                dr.Item("COMPANY_ABBREV") = "Battelle PNNL"
                dr.Item("NO_EMAIL_IND") = 0
                Me.dsCommon.Tables("Users").Rows.Add(dr)
                Me.SqlUpdateCommand1.CommandText = Me.SqlUpdateCommand1.CommandText.Replace("VARCONN.", myUtility.getExtension())
                Me.sqlUsers.Update(Me.dsCommon, "Users")
                Me.dsCommon.AcceptChanges()
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function
#End Region

#Region "Page Action"
    'Processes the user's attempt to login to the site
    Private Sub btnSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Session("JSAction") = ""

        'encode and check the user's input for hidden scripts
        If (myUtility.goodInput(Me.txtPassword) And myUtility.goodInput(Me.txtUserName)) Then
            'Store and locate the user name for authentication
            Session("UserName") = Me.txtUserName.Text
            If (myUtility.findUserName(Me.SqlSelectCommand1, Me.sqlUsers, Me.dsCommon) = True) Then
                'set the user's active status if they are a PNNL System Administrator
                If (Me.dsCommon.Tables("Users").Rows(0).Item("USER_CODE") = 4) Then
                    Me.dsCommon.Tables("Users").Rows(0).Item("ACTIVE_IND") = 1
                End If

                'If the user is not a common user authenticate their entered password
                If (Me.dsCommon.Tables("Users").Rows(0).Item("USER_CODE") > 0) Then
                    If (myUtility.checkPassword(Me.txtPassword.Text, Me.dsCommon, Me.txtUserName.Text) = True) Then
                        If (Me.dsCommon.Tables("Users").Rows(0).Item("ACTIVE_IND") = True) Then
                            Session("UserID") = Me.dsCommon.Tables("Users").Rows(0).Item("USER_KEY")
                            Session("UserType") = Me.dsCommon.Tables("Users").Rows(0).Item("USER_CODE")

                            'update the user's use date to the current date
                            If Not (updateUsageDate()) Then
                                alert("Failed to update your usage data.")
                            Else
                                'check to see if the user's password has expired
                                If (myUtility.isExpired(Session("UserID"), Me.commonConn)) Then
                                    'Transition to new page.
                                    Session("CurrentPage") = "./Password.aspx"
                                    Session("isAuthenticated") = "True"
                                    Session("isExpired") = "True"
                                    Response.Redirect(Session("CurrentPage"))
                                ElseIf (Session("expirationFailure") <> "True") Then
                                    'Transition to new page.
                                    Session("CurrentPage") = "./Main.aspx"
                                    Session("isAuthenticated") = "True"
                                    Session("isExpired") = "False"
                                    Response.Redirect(Session("CurrentPage"))
                                Else
                                    'stay here.
                                    Session("CurrentPage") = "./Login.aspx"
                                    Session("isAuthenticated") = "False"
                                    Session("isExpired") = "True"
                                End If
                            End If
                        Else
                            Me.lblLoginText.Text = "Your login information is correct but your access to this" & _
                            " application has been de-activated.<br/>  If you feel that this is an error then please" & _
                            " contact the " & _
                            "<a title=""Send mail to the Survey Administration Tool Manager."" href=""mailto:gene.gower@pnl.gov?subject=Survey Administration Access"">Survey Administration Tool Manager" & _
                            "</a>."
                        End If
                    Else
                        Me.lblLoginText.Text = "The password does not match the one on file, for that user name. " & _
                        "Please check your user name or re-enter the password.<br/> If you believe that you are receiving this " & _
                        "message in error, please contact the " & _
                        "<a title=""Send mail to the Survey Administration Tool Manager."" href=""mailto:gene.gower@pnl.gov?subject=Survey Administration Access"">Survey Administration Tool Manager" & _
                        "</a>."
                    End If
                Else
                    Me.lblLoginText.Text = "You do not have proper access to this site.<br/>  If you feel that you need access please contact" & _
                    " the <a title=""Send mail to the Survey Administration Tool Manager."" href=""mailto:gene.gower@pnl.gov?subject=Survey Administration Access"">Survey Administration Tool Manager</a>."
                End If
            Else
                'find out if there are any users & generate an administrator if there is none
                If (findUsers() = False) Then
                    If (generateAdmin() = True) Then
                        Me.lblLoginText.Text = "No Administrator existed, Survey User Database Initialized."
                    Else
                        Me.lblLoginText.Text = "No Users exist in this system.  Error generating Administrator.  Survey Database not " & _
                        " Initialized.<br/>  Please contact the " & _
                        "<a title=""Send mail to the Survey Administration Tool Manager."" href=""mailto:gene.gower@pnl.gov?subject=Survey Administration Access"">Survey Administration Tool Manager</a>."
                    End If
                Else
                    Me.lblLoginText.Text = "No user exists in this system with that user name.<br/> Please contact the " & _
                    "<a title=""Send mail to the Survey Administration Tool Manager."" href=""mailto:gene.gower@pnl.gov?subject=Survey Administration Access"">Survey Administration Tool Manager</a> if " & _
                    "you think this is an error."
                End If
            End If
        Else
            Me.lblLoginText.Text = "Possible malicious code detected in your input.  Request for login denied."
        End If
    End Sub

    'Resets the inputs
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Session("JSAction") = ""

        Me.txtPassword.Text = ""
        Me.txtUserName.Text = ""
        Me.lblLoginText.Text = "- Please Login -"
    End Sub
#End Region
End Class
