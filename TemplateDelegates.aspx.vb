Imports System.Web.Mail
Imports System.text
Imports System.Collections.Specialized
Public Class TemplateDelegates
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(TemplateDelegates))
        Me.dsCommon = New System.Data.DataSet
        Me.sqlUsers = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlDeleteCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlInsertCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlUpdateCommand1 = New System.Data.SqlClient.SqlCommand
        Me.commonConn = New System.Data.SqlClient.SqlConnection
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
        Me.sqlUsers.InsertCommand = Me.SqlInsertCommand1
        Me.sqlUsers.SelectCommand = Me.SqlSelectCommand1
        Me.sqlUsers.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_USER", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("USER_KEY", "USER_KEY"), New System.Data.Common.DataColumnMapping("USER_CODE", "USER_CODE"), New System.Data.Common.DataColumnMapping("LAST_NAME", "LAST_NAME"), New System.Data.Common.DataColumnMapping("FIRST_NAME", "FIRST_NAME"), New System.Data.Common.DataColumnMapping("MIDDLE_NAME", "MIDDLE_NAME"), New System.Data.Common.DataColumnMapping("SUFFIX_NAME", "SUFFIX_NAME"), New System.Data.Common.DataColumnMapping("AUTHENTICATION_CODE", "AUTHENTICATION_CODE"), New System.Data.Common.DataColumnMapping("PRIMARY_WORK_PHONE", "PRIMARY_WORK_PHONE"), New System.Data.Common.DataColumnMapping("HANFORD_ID", "HANFORD_ID"), New System.Data.Common.DataColumnMapping("EMAIL_ADDRESS", "EMAIL_ADDRESS"), New System.Data.Common.DataColumnMapping("CREATE_DATE", "CREATE_DATE"), New System.Data.Common.DataColumnMapping("USED_DATE", "USED_DATE"), New System.Data.Common.DataColumnMapping("CREATOR_USER_KEY", "CREATOR_USER_KEY"), New System.Data.Common.DataColumnMapping("ACTIVE_IND", "ACTIVE_IND"), New System.Data.Common.DataColumnMapping("TRAINING_IND", "TRAINING_IND"), New System.Data.Common.DataColumnMapping("COMPANY_ABBREV", "COMPANY_ABBREV"), New System.Data.Common.DataColumnMapping("NO_EMAIL_IND", "NO_EMAIL_IND")})})
        Me.sqlUsers.UpdateCommand = Me.SqlUpdateCommand1
        '
        'SqlDeleteCommand1
        '
        Me.SqlDeleteCommand1.CommandText = "DELETE FROM VARCONN.[SAT_USER] WHERE (([USER_KEY] = @Original_USER_KEY))" & _
            ""
        Me.SqlDeleteCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@Original_USER_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "USER_KEY", System.Data.DataRowVersion.Original, Nothing)})
        '
        'SqlInsertCommand1
        '
        Me.SqlInsertCommand1.CommandText = resources.GetString("SqlInsertCommand1.CommandText")
        Me.SqlInsertCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@USER_CODE", System.Data.SqlDbType.Int, 0, "USER_CODE"), New System.Data.SqlClient.SqlParameter("@LAST_NAME", System.Data.SqlDbType.VarChar, 0, "LAST_NAME"), New System.Data.SqlClient.SqlParameter("@FIRST_NAME", System.Data.SqlDbType.VarChar, 0, "FIRST_NAME"), New System.Data.SqlClient.SqlParameter("@MIDDLE_NAME", System.Data.SqlDbType.VarChar, 0, "MIDDLE_NAME"), New System.Data.SqlClient.SqlParameter("@SUFFIX_NAME", System.Data.SqlDbType.VarChar, 0, "SUFFIX_NAME"), New System.Data.SqlClient.SqlParameter("@AUTHENTICATION_CODE", System.Data.SqlDbType.VarChar, 0, "AUTHENTICATION_CODE"), New System.Data.SqlClient.SqlParameter("@PRIMARY_WORK_PHONE", System.Data.SqlDbType.VarChar, 0, "PRIMARY_WORK_PHONE"), New System.Data.SqlClient.SqlParameter("@HANFORD_ID", System.Data.SqlDbType.VarChar, 0, "HANFORD_ID"), New System.Data.SqlClient.SqlParameter("@EMAIL_ADDRESS", System.Data.SqlDbType.VarChar, 0, "EMAIL_ADDRESS"), New System.Data.SqlClient.SqlParameter("@CREATE_DATE", System.Data.SqlDbType.DateTime, 0, "CREATE_DATE"), New System.Data.SqlClient.SqlParameter("@USED_DATE", System.Data.SqlDbType.DateTime, 0, "USED_DATE"), New System.Data.SqlClient.SqlParameter("@CREATOR_USER_KEY", System.Data.SqlDbType.Int, 0, "CREATOR_USER_KEY"), New System.Data.SqlClient.SqlParameter("@ACTIVE_IND", System.Data.SqlDbType.Bit, 0, "ACTIVE_IND"), New System.Data.SqlClient.SqlParameter("@TRAINING_IND", System.Data.SqlDbType.Bit, 0, "TRAINING_IND"), New System.Data.SqlClient.SqlParameter("@COMPANY_ABBREV", System.Data.SqlDbType.VarChar, 0, "COMPANY_ABBREV"), New System.Data.SqlClient.SqlParameter("@NO_EMAIL_IND", System.Data.SqlDbType.Bit, 0, "NO_EMAIL_IND")})
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = resources.GetString("SqlSelectCommand1.CommandText")
        '
        'SqlUpdateCommand1
        '
        Me.SqlUpdateCommand1.CommandText = resources.GetString("SqlUpdateCommand1.CommandText")
        Me.SqlUpdateCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@USER_CODE", System.Data.SqlDbType.Int, 0, "USER_CODE"), New System.Data.SqlClient.SqlParameter("@LAST_NAME", System.Data.SqlDbType.VarChar, 0, "LAST_NAME"), New System.Data.SqlClient.SqlParameter("@FIRST_NAME", System.Data.SqlDbType.VarChar, 0, "FIRST_NAME"), New System.Data.SqlClient.SqlParameter("@MIDDLE_NAME", System.Data.SqlDbType.VarChar, 0, "MIDDLE_NAME"), New System.Data.SqlClient.SqlParameter("@SUFFIX_NAME", System.Data.SqlDbType.VarChar, 0, "SUFFIX_NAME"), New System.Data.SqlClient.SqlParameter("@AUTHENTICATION_CODE", System.Data.SqlDbType.VarChar, 0, "AUTHENTICATION_CODE"), New System.Data.SqlClient.SqlParameter("@PRIMARY_WORK_PHONE", System.Data.SqlDbType.VarChar, 0, "PRIMARY_WORK_PHONE"), New System.Data.SqlClient.SqlParameter("@HANFORD_ID", System.Data.SqlDbType.VarChar, 0, "HANFORD_ID"), New System.Data.SqlClient.SqlParameter("@EMAIL_ADDRESS", System.Data.SqlDbType.VarChar, 0, "EMAIL_ADDRESS"), New System.Data.SqlClient.SqlParameter("@CREATE_DATE", System.Data.SqlDbType.DateTime, 0, "CREATE_DATE"), New System.Data.SqlClient.SqlParameter("@USED_DATE", System.Data.SqlDbType.DateTime, 0, "USED_DATE"), New System.Data.SqlClient.SqlParameter("@CREATOR_USER_KEY", System.Data.SqlDbType.Int, 0, "CREATOR_USER_KEY"), New System.Data.SqlClient.SqlParameter("@ACTIVE_IND", System.Data.SqlDbType.Bit, 0, "ACTIVE_IND"), New System.Data.SqlClient.SqlParameter("@TRAINING_IND", System.Data.SqlDbType.Bit, 0, "TRAINING_IND"), New System.Data.SqlClient.SqlParameter("@COMPANY_ABBREV", System.Data.SqlDbType.VarChar, 0, "COMPANY_ABBREV"), New System.Data.SqlClient.SqlParameter("@NO_EMAIL_IND", System.Data.SqlDbType.Bit, 0, "NO_EMAIL_IND"), New System.Data.SqlClient.SqlParameter("@Original_USER_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "USER_KEY", System.Data.DataRowVersion.Original, Nothing)})
        '
        'commonConn
        '
        Me.commonConn.FireInfoMessageEventOnUserErrors = False
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected WithEvents imgTitle As System.Web.UI.WebControls.Image
    Protected WithEvents hlpServerSwitch As System.Web.UI.WebControls.Image
    Protected WithEvents chkServerSwitch As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblPlaceHolder As System.Web.UI.WebControls.Label
    Protected WithEvents lblTemplate As System.Web.UI.WebControls.Label
    Protected WithEvents lblNavBar As System.Web.UI.WebControls.Label
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents hlpUser As System.Web.UI.WebControls.Image
    Protected WithEvents lblUser As System.Web.UI.WebControls.Label
    Protected WithEvents hlpHIDPayrollName As System.Web.UI.WebControls.Image
    Protected WithEvents lblHIDPayrollName As System.Web.UI.WebControls.Label
    Protected WithEvents ddlUser As System.Web.UI.WebControls.DropDownList
    Protected WithEvents txtHIDPayrollName As System.Web.UI.WebControls.TextBox
    Protected WithEvents btnHIDPayrollNameFind As System.Web.UI.WebControls.Button
    Protected WithEvents hlpEmail As System.Web.UI.WebControls.Image
    Protected WithEvents hlpFirstName As System.Web.UI.WebControls.Image
    Protected WithEvents hlpLastName As System.Web.UI.WebControls.Image
    Protected WithEvents hlpMiddleName As System.Web.UI.WebControls.Image
    Protected WithEvents hlpSuffixName As System.Web.UI.WebControls.Image
    Protected WithEvents hlpPhone As System.Web.UI.WebControls.Image
    Protected WithEvents hlpCompany As System.Web.UI.WebControls.Image
    Protected WithEvents hlpPNNLTraining As System.Web.UI.WebControls.Image
    Protected WithEvents hlpActive As System.Web.UI.WebControls.Image
    Protected WithEvents hlpOwner As System.Web.UI.WebControls.Image
    Protected WithEvents lblEmail As System.Web.UI.WebControls.Label
    Protected WithEvents lblFirstName As System.Web.UI.WebControls.Label
    Protected WithEvents lblLastName As System.Web.UI.WebControls.Label
    Protected WithEvents lblMiddleName As System.Web.UI.WebControls.Label
    Protected WithEvents lblSuffixName As System.Web.UI.WebControls.Label
    Protected WithEvents lblPhoneNumber As System.Web.UI.WebControls.Label
    Protected WithEvents lblCompany As System.Web.UI.WebControls.Label
    Protected WithEvents lb As System.Web.UI.WebControls.Label
    Protected WithEvents lblTraining As System.Web.UI.WebControls.Label
    Protected WithEvents lblActive As System.Web.UI.WebControls.Label
    Protected WithEvents chkTQSurvey As System.Web.UI.WebControls.CheckBox
    Protected WithEvents chkActive As System.Web.UI.WebControls.CheckBox
    Protected WithEvents chkOwner As System.Web.UI.WebControls.CheckBox
    Protected WithEvents dsCommon As System.Data.DataSet
    Protected WithEvents sqlUsers As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlInsertCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlUpdateCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlDeleteCommand1 As System.Data.SqlClient.SqlCommand
    Protected requestItems As Array
    Protected myNavBar As Navigation = New Navigation
    Protected myUtility As Utility = New Utility
    Protected WithEvents btnTop As System.Web.UI.WebControls.Button
    Protected strNewPassword As String
    Protected WithEvents btnStart As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnPrev As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnReload As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnNext As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnLast As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnNew As System.Web.UI.WebControls.ImageButton
    Protected navButtons As Collection
    Protected sqlCommonAdapter As New System.Data.SqlClient.SqlDataAdapter
    Protected sqlCommonAction As New System.Data.SqlClient.SqlCommand
    Protected WithEvents txtCompany As System.Web.UI.WebControls.Label
    Protected WithEvents txtPhoneNumber As System.Web.UI.WebControls.Label
    Protected WithEvents txtSuffixName As System.Web.UI.WebControls.Label
    Protected WithEvents txtMiddleName As System.Web.UI.WebControls.Label
    Protected WithEvents txtLastName As System.Web.UI.WebControls.Label
    Protected WithEvents txtFirstName As System.Web.UI.WebControls.Label
    Protected WithEvents txtEmail As System.Web.UI.WebControls.Label
    Protected WithEvents commonConn As System.Data.SqlClient.SqlConnection
    Protected WithEvents btnInsert As System.Web.UI.WebControls.Button
    Protected WithEvents btnUpdate As System.Web.UI.WebControls.Button
    Protected WithEvents btnDelete As System.Web.UI.WebControls.Button
    Protected WithEvents btnReset As System.Web.UI.WebControls.Button
    Protected pageOptions As Integer
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
        Trace.Warn("Loading Page")
        Session("JSAction") = ""

        'catch incoming alert messages
        alert(Session("Alert"))

        'clear session
        myUtility.clearSession()

        If (Session.IsNewSession) Then
            Session("Alert") = "Your session has timed out for security purposes.  Please log in again."
            Session("CurrentPage") = "./Login.aspx"
            Session("pageSel") = "Login"
            redirect(Session("CurrentPage"))
        End If

        'get the from page
        If Not (IsPostBack) Then
            Session("FromPage") = getFromPage()
        End If

        'Set the connection based on the machine
        myUtility.getConnection(Me.commonConn)

        'determine template ownership if not defined
        If (Session("isTemplateOwner") Is Nothing) Then
            sqlCommonAction.Connection = Me.commonConn
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            myUtility.determineUserOwnership(sqlCommonAction, sqlCommonAdapter, True)
        End If

        'Put user code to initialize the page here
        If (Session("isAuthenticated") <> "True") Then
            Session("CurrentPage") = "./Login.aspx"
            Session("pageSel") = "Login"
            redirect(Session("CurrentPage"))
        ElseIf (myUtility.checkUserType(3, CType(Session("isTemplateOwner"), Boolean), CType(Session("isTemplateDelegate"), Boolean))) Then
            redirect(Session("CurrentPage"))
        Else
            Session("CurrentPage") = "./TemplateDelegates.aspx"
            Session("pageSel") = "Template"
        End If

        'if the user's password has expired send them to the password reset page
        If (Session("isExpired") = "True") Then
            Session("CurrentPage") = "./Password.aspx"
            Session("pageSel") = "Password"
            redirect(Session("CurrentPage"))
        End If

        'generate collection of navigational buttons
        myUtility.makeNavCollection(Me.navButtons, Me.btnStart, Me.btnPrev, Me.btnReload, Me.btnNext, Me.btnLast, Me.btnNew)

        'Check for user selection from the comments list if we're sure it's not a simple post-back
        If Not (IsPostBack) Then
            'get anything on the request line
            myUtility.getRequest(Page.Request.Url.Query())

            'if we're coming from the PersonSelect page and the process was not stopped load the selected user's data
            'otherwise load the page like normal.
            If (Session("ProcessStopped") = "False") Then
                'refresh the tables
                Dim oldID As Integer = Session("seqUserID")
                Session("seqUserID") = -1

                'load the page data
                If Not (loadData()) Then
                    alert("Failed to load the delegate/owners for this template.")
                Else
                    'to determine what of the nav bar buttons should be available
                    setNavBar()

                    'restore the person data and find the exact user selected
                    Dim dt As DataTable = Session("PersonTable")
                    Dim dr As DataRow = dt.Rows(Session("selectedPersonRow"))

                    'determine if the user already exists as a user and/or delegate/owner
                    Dim isDelegateOwner As Boolean = False
                    Dim isExistingUser As Boolean = False
                    isExistingUser = checkUserExistence(dr)

                    If (isExistingUser) Then
                        'if we found something determine if they are already an owner/delegate of the template
                        isDelegateOwner = checkOwnerDelegate(dr)
                    End If

                    'if the user already exists load their data
                    If (isExistingUser) Then
                        'if the user is already a delegate or owner go to their record otherwise make a new record
                        If (Me.dsCommon.Tables("ExistingUserODStatus").Rows.Count() > 0) Then
                            Session("seqUserID") = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("seqUserID")
                            Session("DelegateRow") = findRow("Delegates", Session("seqUserID"))
                        Else
                            'populate the controls with the selected individual's information
                            Me.dsCommon.Tables("Delegates").Rows.Add(Me.dsCommon.Tables("Delegates").NewRow())
                            Me.txtEmail.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("EMAIL_ADDRESS")
                            Me.txtFirstName.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("FIRST_NAME")
                            Me.txtLastName.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("LAST_NAME")
                            Me.txtMiddleName.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("MIDDLE_NAME")
                            Me.txtSuffixName.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("SUFFIX_NAME")
                            Me.txtPhoneNumber.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("PRIMARY_WORK_PHONE")
                            Me.txtCompany.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("COMPANY_ABBREV")
                            Me.chkTQSurvey.Checked = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("TRAINING_IND")
                            Me.chkActive.Checked = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("ACTIVE_IND")
                            Me.chkOwner.Checked = 0
                        End If
                    Else
                        'populate the controls with the selected individual's information
                        Me.dsCommon.Tables("Delegates").Rows.Add(Me.dsCommon.Tables("Delegates").NewRow())
                        Me.txtEmail.Text = dr.Item("EMAIL_ADDRESS")
                        Me.txtFirstName.Text = dr.Item("FIRST_NAME")
                        Me.txtLastName.Text = dr.Item("LAST_NAME")
                        Me.txtMiddleName.Text = dr.Item("MIDDLE_NAME")
                        Me.txtSuffixName.Text = dr.Item("SUFFIX_NAME")
                        Me.txtPhoneNumber.Text = dr.Item("PRIMARY_WORK_PHONE")
                        Me.txtCompany.Text = dr.Item("COMPANY_ABBREV")
                        Me.chkTQSurvey.Checked = 0
                        Me.chkActive.Checked = 1
                        Me.chkOwner.Checked = 0
                    End If

                    'get the template name - mr sneaky user has access and used a link
                    If (Session("TemplateName") Is Nothing Or Session("TemplateName") = "") Then
                        If Not (setTemplateName()) Then
                            alert("Unable to locate template name.")
                        End If
                    End If

                    'Populate the controls on the page
                    setControls("Override")

                    'set the page options
                    myUtility.optionPreSet(Session("seqUserID"), Session("DelegateMax"), Me.pageOptions)

                    Session("PageOptions") = Me.pageOptions

                    Session("ProcessStopped") = "True"

                    Me.btnInsert.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                    Me.btnUpdate.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                    Me.btnDelete.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                    Me.btnReset.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                End If
            Else
                'load the page data
                If Not (loadData()) Then
                    alert("Failed to load the delegate/owners for this template.")
                Else
                    'to determine what of the nav bar buttons should be available
                    setNavBar()

                    'get the template name - mr sneaky user has access and used a link
                    If (Session("TemplateName") Is Nothing Or Session("TemplateName") = "") Then
                        If Not (setTemplateName()) Then
                            alert("Unable to locate template name.")
                        End If
                    End If

                    'Populate the controls on the page
                    setControls()

                    'set the page options
                    myUtility.optionPreSet(Session("seqUserID"), Session("DelegateMax"), Me.pageOptions)

                    Session("PageOptions") = Me.pageOptions

                    Session("ProcessStopped") = "True"
                    Me.btnInsert.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                    Me.btnUpdate.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                    Me.btnDelete.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                    Me.btnReset.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                End If
            End If
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

    'set the template name
    Private Function setTemplateName() As Boolean
        sqlCommonAdapter.SelectCommand = sqlCommonAction
        sqlCommonAction.Connection = Me.commonConn

        Try
            sqlCommonAction.CommandText = "Select TEMPLATE_NAME from " & myUtility.getExtension() & "SAT_TEMPLATE where " & _
            "TEMPLATE_KEY = " & Session("seqTemplate")
            sqlCommonAdapter.Fill(Me.dsCommon, "TemplateName")
            Session("TemplateName") = Me.dsCommon.Tables("TemplateName").Rows(0).Item("TEMPLATE_NAME")
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function
#End Region

#Region "Data Load"
    'Loads data into the form
    Private Function loadData() As Boolean
        Trace.Warn("Loading Data")
        'Set up the common data adapter and select command
        sqlCommonAction.Connection = Me.commonConn
        sqlCommonAdapter.SelectCommand = sqlCommonAction

        'fill the delegates
        If (loadDelegates()) Then
            'Fill User List
            If (loadUsers()) Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    'loads the delegates
    Private Function loadDelegates() As Boolean
        Trace.Warn("Loading Delegates")
        Try
            'try to erase anything if it exists in the table so we don't end up with duplicate data
            If (Me.dsCommon.Tables.Contains("Delegates")) Then
                Me.dsCommon.Tables("Delegates").Clear()
            End If

            'select only the delegates for the current template
            Dim oldCommand As String
            oldCommand = Me.SqlSelectCommand1.CommandText
            Me.SqlSelectCommand1.CommandText = "SELECT fu.USER_KEY, fu.USER_CODE, fu.LAST_NAME, " & _
            "fu.FIRST_NAME, fu.MIDDLE_NAME, fu.SUFFIX_NAME, fu.AUTHENTICATION_CODE, fu.PRIMARY_WORK_PHONE, " & _
            "fu.HANFORD_ID, fu.EMAIL_ADDRESS, fu.CREATE_DATE, fu.USED_DATE, fu.CREATOR_USER_KEY, fu.ACTIVE_IND, " & _
            "fu.TRAINING_IND, fu.COMPANY_ABBREV, fu.NO_EMAIL_IND, fp.OWNER_IND, fp.DELEGATE_IND, fp.USER_IND FROM " & _
            myUtility.getExtension() & "SAT_USER fu, " & myUtility.getExtension() & "SAT_USER_PERMISSION fp Where fu.USER_KEY = fp.USER_KEY and fp.TEMPLATE_KEY = " & Session("seqTemplate") & _
            " and (fp.DELEGATE_IND = 1 or fp.OWNER_IND = 1) ORDER BY fu.LAST_NAME, fu.FIRST_NAME, fu.MIDDLE_NAME"
            Me.sqlUsers.SelectCommand.Connection = Me.commonConn
            Me.sqlUsers.Fill(Me.dsCommon, "Delegates")
            Me.SqlSelectCommand1.CommandText = oldCommand
            Session("DelegateMax") = Me.dsCommon.Tables("Delegates").Rows.Count()

            'determine if there is any data and if the user id has been set
            If (Session("DelegateMax") = 0) Then
                Session("seqUserID") = -1
            ElseIf (Session("seqUserID") Is Nothing) Then
                Session("seqUserID") = 1
            End If

            'make sure the delegate row never exceeds the maximum delegates
            If (Session("DelegateRow") > Session("DelegateMax")) Then
                Session("DelegateRow") = Session("DelegateMax") - 1
            End If

            'if the user id indicates a new record then move the delegate row to the maximum
            'else set it to just below the user id's number
            If (Session("seqUserID") = -1) Then
                Session("DelegateRow") = Session("DelegateMax")
            ElseIf Not (Session("DelegateMax") = Session("DelegateRow")) Then
                Session("DelegateRow") = findRow("Delegates", Session("seqUserID"))
            ElseIf (Session("DelegateMax") = Session("DelegateRow") And Session("seqUserID") <> -1) Then
                Session("seqUserID") = -1
            End If
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'loads the users in the system
    Private Function loadUsers() As Boolean
        Trace.Warn("Loading Users")
        Try
            'clear existing data
            If (Me.dsCommon.Tables.Contains("Users")) Then
                Me.dsCommon.Tables("Users").Rows.Clear()
            End If
            If (Me.dsCommon.Tables.Contains("ExistingDelegates")) Then
                Me.dsCommon.Tables("ExistingDelegates").Clear()
            End If

            'get a list of users that already exist
            sqlCommonAction.CommandText = "select USER_KEY from " & myUtility.getExtension() & "SAT_USER_PERMISSION where TEMPLATE_KEY = " & Session("seqTemplate") & _
            " and ((DELEGATE_IND = 1) or (OWNER_IND = 1))"
            sqlCommonAdapter.Fill(Me.dsCommon, "ExistingDelegates")

            'build the list of users for the next sql call
            Dim dr As DataRow
            Dim strList As String = ""
            For Each dr In Me.dsCommon.Tables("ExistingDelegates").Rows
                If (strList = "") Then
                    strList &= dr.Item("USER_KEY")
                Else
                    strList &= ", " & dr.Item("USER_KEY")
                End If
            Next

            If (strList <> "") Then
                'use the built list to get a list of users in the system that are not already delegates or owners
                sqlCommonAction.CommandText = "SELECT LAST_NAME + ', ' + FIRST_NAME + ' ' + MIDDLE_NAME + '-' + HANFORD_ID + " & _
                "'-' + EMAIL_ADDRESS as UserName, USER_KEY, USER_CODE, LAST_NAME, FIRST_NAME, MIDDLE_NAME, SUFFIX_NAME, " & _
                "AUTHENTICATION_CODE, PRIMARY_WORK_PHONE, HANFORD_ID, EMAIL_ADDRESS, CREATE_DATE, USED_DATE, CREATOR_USER_KEY, ACTIVE_IND, TRAINING_IND, " & _
                "COMPANY_ABBREV, NO_EMAIL_IND FROM " & myUtility.getExtension() & "SAT_USER where USER_KEY not in (" & strList & ") and not EMAIL_ADDRESS = ' ' " & _
                "and not LAST_NAME = ' ' and not FIRST_NAME = ' ' and HANFORD_ID <> '' and ACTIVE_IND = 1 ORDER BY UserName"
            Else
                'get a list of users in the system
                sqlCommonAction.CommandText = "SELECT LAST_NAME + ', ' + FIRST_NAME + ' ' + MIDDLE_NAME + '-' + HANFORD_ID + " & _
                "'-' + EMAIL_ADDRESS as UserName, USER_KEY, USER_CODE, LAST_NAME, FIRST_NAME, MIDDLE_NAME, SUFFIX_NAME, " & _
                "AUTHENTICATION_CODE, PRIMARY_WORK_PHONE, HANFORD_ID, EMAIL_ADDRESS, CREATE_DATE, USED_DATE, CREATOR_USER_KEY, ACTIVE_IND, TRAINING_IND, " & _
                "COMPANY_ABBREV, NO_EMAIL_IND FROM " & myUtility.getExtension() & "SAT_USER where not EMAIL_ADDRESS = ' ' " & _
                "and not LAST_NAME = ' ' and not FIRST_NAME = ' ' and HANFORD_ID <> '' and ACTIVE_IND = 1 ORDER BY UserName"
            End If

            sqlCommonAdapter.Fill(Me.dsCommon, "Users")
            Me.ddlUser.DataSource = Me.dsCommon.Tables("Users")
            Me.ddlUser.DataTextField = "UserName"
            Me.ddlUser.DataValueField = "USER_KEY"
            Me.ddlUser.DataBind()
            Dim li As System.Web.UI.WebControls.ListItem = New System.Web.UI.WebControls.ListItem
            li.Text = "--Select a User--"
            li.Value = -1
            Me.ddlUser.Items.Insert(0, li)
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Unable to load the users currently in the database.")
            Return False
        End Try
    End Function
#End Region

#Region "Control Settings"
    'populates the controls with the selected user's data
    Private Sub setControls(Optional ByVal override As String = "")
        Trace.Warn("Setting Controls")
        'set the basic page controls based on whether or not we are working with a new survey
        If (Session("seqUserID") > 0) Then
            Try
                Me.txtEmail.Text = Me.dsCommon.Tables("Delegates").Rows(Session("DelegateRow")).Item("EMAIL_ADDRESS")
                Me.txtFirstName.Text = Me.dsCommon.Tables("Delegates").Rows(Session("DelegateRow")).Item("FIRST_NAME")
                Me.txtLastName.Text = Me.dsCommon.Tables("Delegates").Rows(Session("DelegateRow")).Item("LAST_NAME")
                Me.txtMiddleName.Text = Me.dsCommon.Tables("Delegates").Rows(Session("DelegateRow")).Item("MIDDLE_NAME")
                Me.txtSuffixName.Text = Me.dsCommon.Tables("Delegates").Rows(Session("DelegateRow")).Item("SUFFIX_NAME")
                Me.txtPhoneNumber.Text = Me.dsCommon.Tables("Delegates").Rows(Session("DelegateRow")).Item("PRIMARY_WORK_PHONE")
                Me.txtCompany.Text = Me.dsCommon.Tables("Delegates").Rows(Session("DelegateRow")).Item("COMPANY_ABBREV")
                Me.chkTQSurvey.Checked = Me.dsCommon.Tables("Delegates").Rows(Session("DelegateRow")).Item("TRAINING_IND")
                Me.chkActive.Checked = Me.dsCommon.Tables("Delegates").Rows(Session("DelegateRow")).Item("ACTIVE_IND")
                Me.chkOwner.Checked = Me.dsCommon.Tables("Delegates").Rows(Session("DelegateRow")).Item("OWNER_IND")
                Me.lblNavBar.Text = "Delegate " & Session("DelegateRow") + 1 & " of " & Session("DelegateMax")
            Catch ex As Exception
                Trace.Warn(ex.ToString)
            End Try
        Else
            If (override = "") Then
                Me.txtEmail.Text = ""
                Me.txtFirstName.Text = ""
                Me.txtLastName.Text = ""
                Me.txtMiddleName.Text = ""
                Me.txtSuffixName.Text = ""
                Me.txtPhoneNumber.Text = ""
                Me.txtCompany.Text = ""
                Me.chkTQSurvey.Checked = 0
                Me.chkActive.Checked = 1
                Me.chkOwner.Checked = 0
            End If
            Me.lblNavBar.Text = "Delegate " & Session("DelegateRow") + 1 & " of " & Session("DelegateMax") + 1
        End If
    End Sub
#End Region

#Region "Checks"
    'checks to see if the user already exists in the survey system
    Private Function checkUserExistence(ByVal dr As DataRow) As Boolean
        Try
            'initialize the adhoc data collection items
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            sqlCommonAction.Connection = Me.commonConn

            'first check the current list of users and determine if that user already exists before polling the opwhse
            sqlCommonAction.CommandText = "select fu.HANFORD_ID, fu.USER_KEY, hp.hanf_pay_no, fu.FIRST_NAME, " & _
            "fu.LAST_NAME, fu.MIDDLE_NAME, fu.SUFFIX_NAME, " & _
            "fu.PRIMARY_WORK_PHONE, fu.EMAIL_ADDRESS, fu.COMPANY_ABBREV, " & _
            "fu.USER_CODE, fu.TRAINING_IND, fu.ACTIVE_IND, hp.active_sw from " & myUtility.getExtension() & "SAT_USER fu, " & _
            "opwhse.dbo.vw_pub_hanford_people hp where fu.LAST_NAME = hp.Last_Name and fu.FIRST_NAME = hp.first_name " & _
            "and fu.MIDDLE_NAME = hp.middle_initial and hp.active_sw = 'Y' and fu.HANFORD_ID = hp.hanford_id " & _
            "and hp.hanford_id = '" & dr.Item("HANFORD_ID") & "'"
            sqlCommonAdapter.Fill(Me.dsCommon, "ExistingUsers")

            'if we found something set alreadyExists to true
            If (Me.dsCommon.Tables("ExistingUsers").Rows.Count >= 1) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'checks to see if the user is an owner or delegate of the template
    Private Function checkOwnerDelegate(ByVal dr As DataRow) As Boolean
        Try
            'if we found something determine if they are already an owner/delegate of the template
            sqlCommonAction.CommandText = "select OWNER_IND from " & myUtility.getExtension() & "SAT_USER fu, " & _
            "opwhse.dbo.vw_pub_hanford_people hp, " & myUtility.getExtension() & "SAT_USER_PERMISSION fp where fu.LAST_NAME = hp.Last_Name and " & _
            "fu.FIRST_NAME = hp.first_name and fu.MIDDLE_NAME = hp.middle_initial and hp.active_sw = 'Y' and " & _
            "fu.USER_KEY = fp.USER_KEY And fp.TEMPLATE_KEY = " & Session("seqTemplate") & _
            " And (fp.OWNER_IND = 1 Or fp.DELEGATE_IND = 1) and hp.hanford_id = '" & dr.Item("hanford_id") & "'"
            sqlCommonAdapter.Fill(Me.dsCommon, "ExistingUserODStatus")

            'if we found something set alreadyExists to true
            If (Me.dsCommon.Tables("ExistingUserODStatus").Rows.Count >= 1) Then
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

#Region "Page Switch"
    'Takes the user back to the template they were working on
    Private Sub btnTop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTop.Click
        Session("JSAction") = ""
        Session("CurrentPage") = "./Template.aspx"
        Session("pageSel") = "Template"
        redirect(Session("CurrentPage"))
    End Sub
#End Region

#Region "Settings"
    'to determine what of the nav bar buttons should be available
    Private Sub setNavBar()
        'to determine what of the nav bar buttons should be available
        myNavBar.wnb_MoveTo(Me.navButtons, Session("DelegateMax"), Session("DelegateRow"), Switch)
    End Sub
#End Region

#Region "Find"
    'finds the row of the survey we're looking for
    Private Function findRow(ByVal strTable As String, ByVal userID As Integer) As Integer
        Dim intRow As Integer = 0
        Dim dr As DataRow
        For Each dr In Me.dsCommon.Tables(strTable).Rows()
            If (dr.Item("USER_KEY") = userID) Then
                Return intRow
            Else
                intRow += 1
            End If
        Next
        Return -1
    End Function

    'gets the hanford people information for the user we're looking for, returns true if it finds the person(s)
    Private Function getHanfordPeopleInfo(ByVal LookUpString As String) As Boolean
        'initialize the adhoc data collection items
        sqlCommonAdapter.SelectCommand = sqlCommonAction
        sqlCommonAction.Connection = Me.commonConn

        'initialize the adhoc data collection items
        If ((LookUpString.ToUpper.StartsWith("H") And IsNumeric(LookUpString.Substring(1))) Or (LookUpString.Length = 7 And IsNumeric(LookUpString))) Then
            'Process the apparent HID
            Return (processHID(LookUpString))
        ElseIf ((LookUpString.ToUpper.StartsWith("D") And LookUpString.Length = 6) Or LookUpString.Length = 5) Then
            If (IsNumeric(LookUpString.Substring(LookUpString.Length - 3))) Then
                'Process the apparent payroll
                Return (processPayroll(LookUpString))
            Else
                'Process the apparent last name
                Return (processName(LookUpString))
            End If
        Else
            'Process the apparent last name
            Return (processName(LookUpString))
        End If
    End Function

    'Process the HID to find the person
    Private Function processHID(ByVal LookUpString As String) As Boolean
        Try
            'strip the H from the front end if it exists
            If (LookUpString.ToUpper.StartsWith("H")) Then
                LookUpString = LookUpString.Substring(1)
            End If

            Dim alreadyExists As Boolean = False
            'first check the current list of users and determine if that user already exists before polling the opwhse
            sqlCommonAction.CommandText = "select fu.HANFORD_ID, fu.USER_KEY, hp.hanf_pay_no, fu.FIRST_NAME, " & _
            "fu.LAST_NAME, fu.MIDDLE_NAME, fu.SUFFIX_NAME, " & _
            "fu.PRIMARY_WORK_PHONE, fu.EMAIL_ADDRESS, fu.COMPANY_ABBREV, " & _
            "fu.USER_CODE, fu.TRAINING_IND, fu.ACTIVE_IND, hp.active_sw from " & myUtility.getExtension() & "SAT_USER fu, " & _
            "opwhse.dbo.vw_pub_hanford_people hp where fu.HANFORD_ID = hp.hanford_id and hp.active_sw = 'Y' " & _
            "and fu.HANFORD_ID = '" & LookUpString & "'" & " and fu.HANFORD_ID <> ''"
            sqlCommonAdapter.Fill(Me.dsCommon, "ExistingUsers")

            'if we found something then they already exist
            If (Me.dsCommon.Tables("ExistingUsers").Rows.Count >= 1) Then
                Session("PersonTable") = Me.dsCommon.Tables("ExistingUsers")
                alreadyExists = True
            End If

            If (alreadyExists) Then
                'if we found something determine if they are already an owner/delegate of the template
                sqlCommonAction.CommandText = "select OWNER_IND from " & myUtility.getExtension() & "SAT_USER fu, " & _
                "opwhse.dbo.vw_pub_hanford_people hp, " & myUtility.getExtension() & "SAT_USER_PERMISSION fp where fu.HANFORD_ID = hp.hanford_id " & _
                " and hp.active_sw = 'Y' and " & _
                "fu.USER_KEY = fp.USER_KEY And fp.TEMPLATE_KEY = " & Session("seqTemplate") & _
                " And (fp.OWNER_IND = 1 Or fp.DELEGATE_IND = 1) and fu.HANFORD_ID = '" & LookUpString & "'" & _
                " and fu.HANFORD_ID <> ''"
                sqlCommonAdapter.Fill(Me.dsCommon, "ExistingUserODStatus")
                Return True
            Else
                'set up and process the sql statement
                sqlCommonAction.CommandText = "select isnull(max(hanford_id), '') as hanford_id, isnull(max(hanf_pay_no), '') as " & _
                "hanf_pay_no, max(first_name) as first_name, max(last_name) as last_name, isnull(max(middle_initial), '') as " & _
                "MIDDLE_NAME, isnull(max(name_prefix), '') as name_prefix, isnull(max(name_suffix), '') as SUFFIX_NAME, " & _
                "isnull(max(phone_no), '') as PRIMARY_WORK_PHONE, isnull(max(internet_email_address), '') as EMAIL_ADDRESS, " & _
                "isnull(max(employed_by_hcid), '') as COMPANY_ABBREV, isNull(max(active_sw), 'N') as active_sw from  (" & _
                "select hanford_id, hanf_pay_no, first_name, last_name, middle_initial, NULL as name_prefix, name_suffix, phone_no, " & _
                "internet_email_address, employed_by_hcid, active_sw from opwhse.dbo.vw_pub_hanford_people union select hanford_id, " & _
                "pay_no as hanf_pay_no, first_name, last_name, middle_name as middle_initial, NULL as name_prefix, name_suffix, " & _
                "work_phone as phone_no, internet_address as internet_email_address, employed_by_hcid, NULL as active_sw " & _
                "from opwhse.dbo.vw_pub_pnnl_associate union select hanford_id, pnnl_pay_no as hanf_pay_no, first_name, last_name, " & _
                "middle_initial, name_prefix, name_suffix, contact_work_phone as phone_no, internet_email_address, company as " & _
                "employed_by_hcid, NULL as active_sw from opwhse.dbo.vw_pub_bmi_employee) a where active_sw = 'Y' and " & _
                "hanford_id = '" & LookUpString & "'" & _
                " group by hanford_id, last_name, first_name, middle_initial, internet_email_address"
                sqlCommonAdapter.Fill(Me.dsCommon, "PersonData")

                'if we found something return true otherwise return false
                If (Me.dsCommon.Tables("PersonData").Rows.Count >= 1) Then
                    Session("PersonTable") = Me.dsCommon.Tables("PersonData")
                    Return True
                Else
                    Session("PersonTable") = Nothing
                    Return False
                End If
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'Process the Payroll to find the person
    Private Function processPayroll(ByVal LookUpString As String) As Boolean
        Try
            'strip the D from the front end if it exists
            If (LookUpString.ToUpper.StartsWith("D")) Then
                LookUpString = LookUpString.Substring(1)
            End If

            Dim alreadyExists As Boolean = False
            'first check the current list of users and determine if that user already exists before polling the opwhse
            sqlCommonAction.CommandText = "select fu.HANFORD_ID, fu.USER_KEY, hp.hanf_pay_no, fu.FIRST_NAME, " & _
            "fu.LAST_NAME, fu.MIDDLE_NAME, fu.SUFFIX_NAME, " & _
            "fu.PRIMARY_WORK_PHONE, fu.EMAIL_ADDRESS, fu.COMPANY_ABBREV, " & _
            "fu.USER_CODE, fu.TRAINING_IND, fu.ACTIVE_IND, hp.active_sw from " & myUtility.getExtension() & "SAT_USER fu, " & _
            "opwhse.dbo.vw_pub_hanford_people hp where fu.HANFORD_ID = hp.hanford_id " & _
            "and hp.active_sw = 'Y' and hp.hanf_pay_no = '" & LookUpString & "'" & " and fu.HANFORD_ID <> ''"
            sqlCommonAdapter.Fill(Me.dsCommon, "ExistingUsers")

            'if we found something return true otherwise return false
            If (Me.dsCommon.Tables("ExistingUsers").Rows.Count >= 1) Then
                alreadyExists = True
            End If

            If (alreadyExists) Then
                'if we found something determine if they are already an owner/delegate of the template
                sqlCommonAction.CommandText = "select OWNER_IND from " & myUtility.getExtension() & "SAT_USER fu, " & _
                "opwhse.dbo.vw_pub_hanford_people hp, " & myUtility.getExtension() & "SAT_USER_PERMISSION fp where fu.HANFORD_ID = hp.hanford_id " & _
                "and hp.active_sw = 'Y' and " & _
                "fu.USER_KEY = fp.USER_KEY And fp.TEMPLATE_KEY = " & Session("seqTemplate") & _
                " And (fp.OWNER_IND = 1 Or fp.DELEGATE_IND = 1) and hp.hanf_pay_no = '" & LookUpString & "'" & _
                " and fu.HANFORD_ID <> ''"
                sqlCommonAdapter.Fill(Me.dsCommon, "ExistingUserODStatus")
                Return True
            Else
                'set up and process the sql statement
                sqlCommonAction.CommandText = "select isnull(max(hanford_id), '') as hanford_id, isnull(max(hanf_pay_no), '') as " & _
                "hanf_pay_no, max(first_name) as first_name, max(last_name) as last_name, isnull(max(middle_initial), '') as " & _
                "MIDDLE_NAME, isnull(max(name_prefix), '') as name_prefix, isnull(max(name_suffix), '') as SUFFIX_NAME, " & _
                "isnull(max(phone_no), '') as PRIMARY_WORK_PHONE, isnull(max(internet_email_address), '') as EMAIL_ADDRESS, " & _
                "isnull(max(employed_by_hcid), '') as COMPANY_ABBREV, isNull(max(active_sw), 'N') as active_sw from  (" & _
                "select hanford_id, hanf_pay_no, first_name, last_name, middle_initial, NULL as name_prefix, name_suffix, phone_no, " & _
                "internet_email_address, employed_by_hcid, active_sw from opwhse.dbo.vw_pub_hanford_people union select hanford_id, " & _
                "pay_no as hanf_pay_no, first_name, last_name, middle_name as middle_initial, NULL as name_prefix, name_suffix, " & _
                "work_phone as phone_no, internet_address as internet_email_address, employed_by_hcid, NULL as active_sw " & _
                "from opwhse.dbo.vw_pub_pnnl_associate union select hanford_id, pnnl_pay_no as hanf_pay_no, first_name, last_name, " & _
                "middle_initial, name_prefix, name_suffix, contact_work_phone as phone_no, internet_email_address, company as " & _
                "employed_by_hcid, NULL as active_sw from opwhse.dbo.vw_pub_bmi_employee) a where active_sw = 'Y' and " & _
                "hanf_pay_no = '" & LookUpString & "'" & _
                " group by hanford_id, last_name, first_name, middle_initial, internet_email_address"
                sqlCommonAdapter.Fill(Me.dsCommon, "PersonData")

                'if we found something return true otherwise return false
                If (Me.dsCommon.Tables("PersonData").Rows.Count >= 1) Then
                    Session("PersonTable") = Me.dsCommon.Tables("PersonData")
                    Return True
                Else
                    Session("PersonTable") = Nothing
                    Return False
                End If
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'Process the Name to find the person
    Private Function processName(ByVal LookUpString As String) As Boolean
        Try
            'fix any apostrophes that might be in the search string
            LookUpString = LookUpString.Replace("'", "''")

            'set up and process the sql statement
            sqlCommonAction.CommandText = "select isnull(max(hanford_id), '') as hanford_id, isnull(max(hanf_pay_no), '') as " & _
            "hanf_pay_no, max(first_name) as first_name, max(last_name) as last_name, isnull(max(middle_initial), '') as " & _
            "MIDDLE_NAME, isnull(max(name_prefix), '') as name_prefix, isnull(max(name_suffix), '') as SUFFIX_NAME, " & _
            "isnull(max(phone_no), '') as PRIMARY_WORK_PHONE, isnull(max(internet_email_address), '') as EMAIL_ADDRESS, " & _
            "isnull(max(employed_by_hcid), '') as COMPANY_ABBREV, isNull(max(active_sw), 'N') as active_sw, " & _
            "max(last_name) + ', ' + max(first_name) + ' ' + isNull(max(middle_initial), '') + ' - ' + " & _
            "isNull(max(internet_email_address), '') as NameEmail from  (" & _
            "select hanford_id, hanf_pay_no, first_name, last_name, middle_initial, NULL as name_prefix, name_suffix, phone_no, " & _
            "internet_email_address, employed_by_hcid, active_sw from opwhse.dbo.vw_pub_hanford_people union select hanford_id, " & _
            "pay_no as hanf_pay_no, first_name, last_name, middle_name as middle_initial, NULL as name_prefix, name_suffix, " & _
            "work_phone as phone_no, internet_address as internet_email_address, employed_by_hcid, NULL as active_sw " & _
            "from opwhse.dbo.vw_pub_pnnl_associate union select hanford_id, pnnl_pay_no as hanf_pay_no, first_name, last_name, " & _
            "middle_initial, name_prefix, name_suffix, contact_work_phone as phone_no, internet_email_address, company as " & _
            "employed_by_hcid, NULL as active_sw from opwhse.dbo.vw_pub_bmi_employee) a where active_sw = 'Y' and " & _
            "last_name like '" & LookUpString & "%'" & " and internet_email_address <> ''" & _
            " group by hanford_id, last_name, first_name, middle_initial, internet_email_address " & _
            "order by last_name, first_name"
            sqlCommonAdapter.Fill(Me.dsCommon, "PersonData")

            'if we found something return true otherwise return false
            If (Me.dsCommon.Tables("PersonData").Rows.Count = 1) Then
                Session("PersonTable") = Me.dsCommon.Tables("PersonData")
                Return True
            ElseIf (Me.dsCommon.Tables("PersonData").Rows.Count > 1) Then
                Session("PersonTable") = Me.dsCommon.Tables("PersonData")
                Session("CurrentPage") = "./PersonSelect.aspx"
                Session("pageSel") = "Template"
                redirect(Session("CurrentPage"))
            Else
                Session("PersonTable") = Nothing
                Return False
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function
#End Region

#Region "Nav Bar Events"
    'Moves the user to the first record
    Private Sub wnbDelegates_MoveToStart(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnStart.Click
        Session("JSAction") = ""

        If Not (loadData()) Then
            alert("Failed to load the delegates/owners for this survey.")
        Else
            Try
                Session("DelegateRow") = 0
                Session("seqUserID") = Me.dsCommon.Tables("Delegates").Rows(Session("DelegateRow")).Item("USER_KEY")
                setNavBar()
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("seqUserID"), Session("DelegateMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the first template delegate/owner.")
            End Try
        End If
    End Sub

    'Moves the user to the previous record
    Private Sub wnbDelegates_MovePrev(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnPrev.Click
        Session("JSAction") = ""

        If Not (loadData()) Then
            alert("Failed to load the delegates/owners for this survey.")
        Else
            Try
                Session("DelegateRow") -= 1
                If (Session("DelegateRow") = -1) Then
                    Session("DelegateRow") = 0
                End If
                Session("seqUserID") = Me.dsCommon.Tables("Delegates").Rows(Session("DelegateRow")).Item("USER_KEY")
                setNavBar()
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("seqUserID"), Session("DelegateMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the previous template delegate/owner.")
            End Try
        End If
    End Sub

    'reloads the page and the data from the database while attempting to locate the user's location
    Private Sub wnbDelegates_MoveReload(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnReload.Click
        Session("JSAction") = ""

        If Not (loadData()) Then
            alert("Failed to load the delegates/owners for this survey.")
        Else
            Try
                setNavBar()
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("seqUserID"), Session("DelegateMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error resetting the current template delegate/owner.")
            End Try
        End If
    End Sub

    'Moves the user to the next record
    Private Sub wnbDelegates_MoveNext(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnNext.Click
        Session("JSAction") = ""

        If Not (loadData()) Then
            alert("Failed to load the delegates/owners for this survey.")
        Else
            Try
                Session("DelegateRow") += 1
                Session("seqUserID") = Me.dsCommon.Tables("Delegates").Rows(Session("DelegateRow")).Item("USER_KEY")
                setNavBar()
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("seqUserID"), Session("DelegateMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the next template delegate/owner.")
            End Try
        End If
    End Sub

    'Moves the user to the last record
    Private Sub wnbDelegates_MoveToEnd(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnLast.Click
        Session("JSAction") = ""

        If Not (loadData()) Then
            alert("Failed to load the delegates/owners for this survey.")
        Else
            Try
                Session("DelegateRow") = Session("DelegateMax") - 1
                Session("seqUserID") = Me.dsCommon.Tables("Delegates").Rows(Session("DelegateRow")).Item("USER_KEY")
                setNavBar()
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("seqUserID"), Session("DelegateMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the last template delegate/owner.")
            End Try
        End If
    End Sub

    'Inserts a record into the table
    Private Sub wnbDelegates_InsertRow(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnNew.Click
        Session("JSAction") = ""

        If Not (loadData()) Then
            alert("Failed to load the delegates/owners for this survey.")
        Else
            Try
                Session("seqUserID") = -1
                Session("DelegateRow") = Session("DelegateMax")
                Me.dsCommon.Tables("Delegates").Rows.Add(Me.dsCommon.Tables("Delegates").NewRow())
                Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("EMAIL_ADDRESS") = ""
                Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("FIRST_NAME") = ""
                Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("LAST_NAME") = ""
                Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("MIDDLE_NAME") = ""
                Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("SUFFIX_NAME") = ""
                Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("PRIMARY_WORK_PHONE") = ""
                Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("COMPANY_ABBREV") = ""
                Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("USER_CODE") = -1
                Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("TRAINING_IND") = 0
                Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("ACTIVE_IND") = 1
                Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("OWNER_IND") = 0
                myNavBar.wnb_MoveTo(Me.navButtons, Session("DelegateMax"), Session("DelegateRow"))
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("seqUserID"), Session("DelegateMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error creating new template delegate/owner.")
            End Try
        End If
    End Sub
#End Region

#Region "Page Action"
    'loads the data for the selected user
    Private Sub ddlUser_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddlUser.SelectedIndexChanged
        Session("JSAction") = ""

        Try
            'refresh the tables
            Dim oldID As Integer = Session("seqUserID")
            Session("seqUserID") = -1
            Dim userID As Integer = Me.ddlUser.SelectedValue
            If (loadData()) Then
                'find the user information we'll need to load
                Dim userRow As Integer
                userRow = findRow("Users", userID)

                'If the user's information still exists then use it otherwise inform the user that the user's information is not retrievable
                If (userRow <> -1) Then
                    'if we're not working with a new record create one and populate it.  Otherwise re-populate the page.
                    If (Session("seqUserID") = -1) Then
                        Me.dsCommon.Tables("Delegates").Rows.Add(Me.dsCommon.Tables("Delegates").NewRow())
                        'populate the controls with the selected individual's information
                        Me.txtEmail.Text = Me.dsCommon.Tables("Users").Rows(userRow).Item("EMAIL_ADDRESS")
                        Me.txtFirstName.Text = Me.dsCommon.Tables("Users").Rows(userRow).Item("FIRST_NAME")
                        Me.txtLastName.Text = Me.dsCommon.Tables("Users").Rows(userRow).Item("LAST_NAME")
                        Me.txtMiddleName.Text = Me.dsCommon.Tables("Users").Rows(userRow).Item("MIDDLE_NAME")
                        Me.txtSuffixName.Text = Me.dsCommon.Tables("Users").Rows(userRow).Item("SUFFIX_NAME")
                        Me.txtPhoneNumber.Text = Me.dsCommon.Tables("Users").Rows(userRow).Item("PRIMARY_WORK_PHONE")
                        Me.txtCompany.Text = Me.dsCommon.Tables("Users").Rows(userRow).Item("COMPANY_ABBREV")
                        Me.chkTQSurvey.Checked = Me.dsCommon.Tables("Users").Rows(userRow).Item("TRAINING_IND")
                        Me.chkActive.Checked = Me.dsCommon.Tables("Users").Rows(userRow).Item("ACTIVE_IND")
                        Me.chkOwner.Checked = 0
                    End If

                    'set the page options
                    myUtility.optionPreSet(Session("seqUserID"), Session("DelegateMax"), Me.pageOptions)

                    Session("PageOptions") = Me.pageOptions

                    Session("seqUserID") = userID

                    'Populate the controls on the page
                    setControls()
                    Me.lblNavBar.Text = "Survey " & Session("DelegateRow") + 1 & " of " & Session("DelegateMax") + 1

                    'to determine what of the nav bar buttons should be available
                    setNavBar()
                Else
                    alert("Error, this user may have been deleted within the last few seconds.  User information is irretrievable for the selected individual.")
                    Session("seqUserID") = oldID
                End If
            Else
                alert("Unable to load data for selected user.")
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Unable to load data for selected user.")
        End Try
    End Sub

    'attempts to locate the user's information in the opwhse under the Hanford People view
    Private Sub btnHIDPayrollNameFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHIDPayrollNameFind.Click
        Session("JSAction") = ""

        Try
            If (myUtility.goodInput(Me.txtHIDPayrollName)) Then
                'refresh the tables
                Dim oldID As Integer = Session("seqUserID")
                Session("seqUserID") = -1
                If (loadData()) Then
                    'process the input only if the user actually put something in the find text field
                    If (Me.txtHIDPayrollName.Text <> "") Then
                        'find the user information we'll need to load
                        If (getHanfordPeopleInfo(Me.txtHIDPayrollName.Text)) Then
                            'if we're working with a new record create one and populate it.  Otherwise re-populate the page.
                            If (Session("seqUserID") = -1) Then
                                'if the user already exists load their data
                                If (Me.dsCommon.Tables.Contains("ExistingUsers")) Then
                                    If (Me.dsCommon.Tables("ExistingUsers").Rows.Count() > 0) Then
                                        'if the user is already a delegate or owner go to their record otherwise make a new record
                                        If (Me.dsCommon.Tables("ExistingUserODStatus").Rows.Count() > 0) Then
                                            Session("seqUserID") = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("USER_CODE")
                                            Session("DelegateRow") = findRow("Delegates", Session("seqUserID"))
                                        Else
                                            'populate the controls with the selected individual's information
                                            Me.dsCommon.Tables("Delegates").Rows.Add(Me.dsCommon.Tables("Delegates").NewRow())
                                            Me.txtEmail.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("EMAIL_ADDRESS")
                                            Me.txtFirstName.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("FIRST_NAME")
                                            Me.txtLastName.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("LAST_NAME")
                                            Me.txtMiddleName.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("MIDDLE_NAME")
                                            Me.txtSuffixName.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("SUFFIX_NAME")
                                            Me.txtPhoneNumber.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("PRIMARY_WORK_PHONE")
                                            Me.txtCompany.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("COMPANY_ABBREV")
                                            Me.chkTQSurvey.Checked = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("TRAINING_IND")
                                            Me.chkActive.Checked = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("ACTIVE_IND")
                                            Me.chkOwner.Checked = 0
                                        End If
                                    Else
                                        'populate the controls with the selected individual's information
                                        Me.dsCommon.Tables("Delegates").Rows.Add(Me.dsCommon.Tables("Delegates").NewRow())
                                        Me.txtEmail.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("EMAIL_ADDRESS")
                                        Me.txtFirstName.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("FIRST_NAME")
                                        Me.txtLastName.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("LAST_NAME")
                                        Me.txtMiddleName.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("MIDDLE_NAME")
                                        Me.txtSuffixName.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("SUFFIX_NAME")
                                        Me.txtPhoneNumber.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("PRIMARY_WORK_PHONE")
                                        Me.txtCompany.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("COMPANY_ABBREV")
                                        Me.chkTQSurvey.Checked = 0
                                        Me.chkActive.Checked = 1
                                        Me.chkOwner.Checked = 0
                                    End If
                                Else
                                    'populate the controls with the selected individual's information
                                    Me.dsCommon.Tables("Delegates").Rows.Add(Me.dsCommon.Tables("Delegates").NewRow())
                                    Me.txtEmail.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("EMAIL_ADDRESS")
                                    Me.txtFirstName.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("FIRST_NAME")
                                    Me.txtLastName.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("LAST_NAME")
                                    Me.txtMiddleName.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("MIDDLE_NAME")
                                    Me.txtSuffixName.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("SUFFIX_NAME")
                                    Me.txtPhoneNumber.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("PRIMARY_WORK_PHONE")
                                    Me.txtCompany.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("COMPANY_ABBREV")
                                    Me.chkTQSurvey.Checked = 0
                                    Me.chkActive.Checked = 1
                                    Me.chkOwner.Checked = 0
                                End If
                            End If

                            'Populate the controls on the page
                            setControls("override")
                            If (Me.dsCommon.Tables.Contains("ExistingUserODStatus")) Then
                                If (Me.dsCommon.Tables("ExistingUserODStatus").Rows.Count() > 0) Then
                                    Me.lblNavBar.Text = "Survey " & Session("DelegateRow") + 1 & " of " & Session("DelegateMax")
                                End If
                            Else
                                Me.lblNavBar.Text = "Survey " & Session("DelegateRow") + 1 & " of " & Session("DelegateMax") + 1
                            End If

                            'to determine what of the nav bar buttons should be available
                            setNavBar()

                            'set the page options
                            myUtility.optionPreSet(Session("seqUserID"), Session("DelegateMax"), Me.pageOptions)

                            Session("PageOptions") = Me.pageOptions
                            Me.txtHIDPayrollName.Text = ""
                        Else
                            alert("This person was not found.")
                            'to determine what of the nav bar buttons should be available
                            setNavBar()

                            'Populate the controls on the page
                            setControls()

                            'set the page options
                            myUtility.optionPreSet(Session("seqUserID"), Session("DelegateMax"), Me.pageOptions)

                            Session("PageOptions") = Me.pageOptions
                            Me.txtHIDPayrollName.Text = ""
                        End If
                    Else
                        alert("You must input some text into the HID, Employee ID, or Last Name field to use the find functionality.")
                        'to determine what of the nav bar buttons should be available
                        setNavBar()

                        'Populate the controls on the page
                        setControls()

                        'set the page options
                        myUtility.optionPreSet(Session("seqUserID"), Session("DelegateMax"), Me.pageOptions)

                        Session("PageOptions") = Me.pageOptions
                        Me.txtHIDPayrollName.Text = ""
                    End If
                Else
                    Me.txtHIDPayrollName.Text = ""
                    alert("Error loading delegate/owner information.")
                End If
            Else
                Me.txtHIDPayrollName.Text = ""
                alert("You may have entered malicious code into the HID/Emplid/Last Name text field.  Find aborted.")
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Me.txtHIDPayrollName.Text = ""
            alert("Error loading delegate/owner information.")
        End Try
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnInsert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsert.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'refresh the delegate information saving the type so that it can be restored
        If Not (loadData()) Then
            blnContinue = False
            alert("Error loading delegate/owner information. Insert aborted.")
        End If

        If (blnContinue) Then
            'do the insert
            failure = doInsert()

            'check the user type
            If (myUtility.checkUserType(3, CType(Session("isTemplateOwner"), Boolean), CType(Session("isTemplateDelegate"), Boolean))) Then
                Response.Redirect(Session("CurrentPage"))
            End If

            'reload the data
            If Not (loadData()) Then
                blnContinue = False
                alert("Error loading delegate/owner information.")
            End If

            If (blnContinue) Then
                'to determine what of the nav bar buttons should be available
                setNavBar()

                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
                Session("TransExists") = False
                'set the page options
                myUtility.optionPreSet(Session("seqUserID"), Session("DelegateMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'refresh the delegate information saving the type so that it can be restored
        If Not (loadData()) Then
            blnContinue = False
            alert("Error loading delegate/owner information. Update aborted.")
        End If

        If (blnContinue) Then
            'do the update
            failure = doUpdate()

            'check the user type
            If (myUtility.checkUserType(3, CType(Session("isTemplateOwner"), Boolean), CType(Session("isTemplateDelegate"), Boolean))) Then
                Response.Redirect(Session("CurrentPage"))
            End If

            'reload the data
            If Not (loadData()) Then
                blnContinue = False
                alert("Error loading delegate/owner information.")
            End If

            If (blnContinue) Then
                'to determine what of the nav bar buttons should be available
                setNavBar()

                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
                Session("TransExists") = False
                'set the page options
                myUtility.optionPreSet(Session("seqUserID"), Session("DelegateMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'refresh the delegate information saving the type so that it can be restored
        If Not (loadData()) Then
            blnContinue = False
            alert("Error loading delegate/owner information. Delete aborted.")
        End If

        If (blnContinue) Then
            'do the delete
            failure = doDelete()

            'check the user type
            If (myUtility.checkUserType(3, CType(Session("isTemplateOwner"), Boolean), CType(Session("isTemplateDelegate"), Boolean))) Then
                Response.Redirect(Session("CurrentPage"))
            End If

            'reload the data
            If Not (loadData()) Then
                blnContinue = False
                alert("Error loading delegate/owner information.")
            End If

            If (blnContinue) Then
                'to determine what of the nav bar buttons should be available
                setNavBar()

                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
                Session("TransExists") = False
                'set the page options
                myUtility.optionPreSet(Session("seqUserID"), Session("DelegateMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        If (blnContinue) Then
            'do the reset
            failure = doReset()

            'check the user type
            If (myUtility.checkUserType(3, CType(Session("isTemplateOwner"), Boolean), CType(Session("isTemplateDelegate"), Boolean))) Then
                Response.Redirect(Session("CurrentPage"))
            End If

            If (blnContinue) Then
                'to determine what of the nav bar buttons should be available
                setNavBar()

                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
                Session("TransExists") = False
                'set the page options
                myUtility.optionPreSet(Session("seqUserID"), Session("DelegateMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub
#End Region

#Region "Insert Code"
    'drives the commit and roll-back operations of the insert
    Private Function doInsert() As Boolean
        Try
            'start transaction
            If (myUtility.setupTransaction(sqlCommonAction, Me.commonConn)) Then
                Dim failure As Boolean = False
                'Check for actual information to insert
                If (Me.txtFirstName.Text <> "" And Me.txtLastName.Text <> "" And Me.txtEmail.Text <> "") Then
                    'If we're inserting a new record then add it onto the end of the recordset and add it to the existing Delegates
                    If (insertUser(sqlCommonAction, sqlCommonAdapter)) Then
                        If (insertPermission(sqlCommonAction, sqlCommonAdapter)) Then
                            sqlCommonAction.Transaction.Commit()
                            If (Me.dsCommon.Tables("UserList").Rows.Count <> 0) Then
                                If Not (sendMail(Me.txtEmail.Text, Session("newUserName"), True)) Then
                                    alert("Failed to notify the new Delegate/Owner of this template via e-mail.")
                                End If
                            Else
                                If Not (sendMail(Me.txtEmail.Text, Session("newUserName"))) Then
                                    alert("Failed to notify the new Delegate/Owner of this template via e-mail.")
                                End If
                            End If
                        Else
                            sqlCommonAction.Transaction.Rollback()
                            failure = True
                        End If
                    Else
                        sqlCommonAction.Transaction.Rollback()
                        failure = True
                    End If
                Else
                    alert("You must provide a first and last name as well as an e-mail address and select a user type at minimum.")
                    failure = True
                    sqlCommonAction.Transaction.Rollback()
                End If
                Return failure
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Failed to insert this new Owner/Delegate.")
                Return True
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Failed to insert this new Owner/Delegate.")
            Return True
        End Try
    End Function

    'attempts to insert a user, returns false if it cannot
    Private Function insertUser(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'find out if the user already exists in the system
            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER where EMAIL_ADDRESS = '" & Me.txtEmail.Text & "'"
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            sqlCommonAdapter.Fill(Me.dsCommon, "UserList")

            'if a user exists then return true instead of adding a new person to the user's table
            If (Me.dsCommon.Tables("UserList").Rows.Count = 1) Then
                'set the user's level of access accordingly
                If (Me.dsCommon.Tables("UserList").Rows(0).Item("USER_CODE") < 3) Then
                    sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_USER set USER_CODE = 3 where USER_KEY = " & _
                        Me.dsCommon.Tables("UserList").Rows(0).Item("USER_KEY")
                    sqlCommonAction.ExecuteNonQuery()
                End If

                'update the user's information
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_USER SET TRAINING_IND = " & IIf(Me.chkTQSurvey.Checked, 1, 0) & _
                ", ACTIVE_IND = " & IIf(Me.chkActive.Checked, 1, 0) & " where USER_KEY = " & _
                Me.dsCommon.Tables("UserList").Rows(0).Item("USER_KEY")
                sqlCommonAction.ExecuteNonQuery()

                'set the user id and return
                Session("seqUserID") = Me.dsCommon.Tables("UserLIst").Rows(0).Item("USER_KEY")
                Return True
            ElseIf (Me.dsCommon.Tables("UserList").Rows.Count > 1) Then
                alert("Not enough information to differentiate between two or more existing users.  The user you are trying to add may already exist in our users list.  Please select the user you are looking for from that list and update their information to indicate their new status for this template.")
                Return False
            End If

            'parameterized the text input to allow for things like quotes to get through
            Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@EMAIL_ADDRESS", System.Data.SqlDbType.VarChar, 128, "EMAIL_ADDRESS")
            param0.Value = Me.txtEmail.Text
            sqlCommonAction.Parameters.Add(param0)
            Dim param1 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@FIRST_NAME", System.Data.SqlDbType.VarChar, 50, "FIRST_NAME")
            param1.Value = Me.txtFirstName.Text
            sqlCommonAction.Parameters.Add(param1)
            Dim param2 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@LAST_NAME", System.Data.SqlDbType.VarChar, 50, "LAST_NAME")
            param2.Value = Me.txtLastName.Text
            sqlCommonAction.Parameters.Add(param2)
            Dim param3 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@MIDDLE_NAME", System.Data.SqlDbType.VarChar, 50, "MIDDLE_NAME")
            param3.Value = Me.txtMiddleName.Text
            sqlCommonAction.Parameters.Add(param3)
            Dim param4 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@SUFFIX_NAME", System.Data.SqlDbType.VarChar, 50, "SUFFIX_NAME")
            param4.Value = Me.txtSuffixName.Text
            sqlCommonAction.Parameters.Add(param4)
            Dim param5 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@PRIMARY_WORK_PHONE", System.Data.SqlDbType.VarChar, 32, "PRIMARY_WORK_PHONE")
            param5.Value = Me.txtPhoneNumber.Text
            sqlCommonAction.Parameters.Add(param5)
            Dim param6 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@COMPANY_ABBREV", System.Data.SqlDbType.VarChar, 32, "COMPANY_ABBREV")
            param6.Value = Me.txtCompany.Text
            sqlCommonAction.Parameters.Add(param6)
            Dim param7 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@HANFORD_ID", System.Data.SqlDbType.VarChar, 12, "HANFORD_ID")

            'get the user's HID if it exists
            If (Session("PersonTable") Is Nothing) Then
                'get the user's informaton - the user adding must have manually input their information
                'set up and process the sql statement
                sqlCommonAction.CommandText = "select isnull(max(hanford_id), '') as hanford_id, isnull(max(hanf_pay_no), '') as " & _
                "hanf_pay_no, max(first_name) as first_name, max(last_name) as last_name, isnull(max(middle_initial), '') as " & _
                "MIDDLE_NAME, isnull(max(name_prefix), '') as name_prefix, isnull(max(name_suffix), '') as SUFFIX_NAME, " & _
                "isnull(max(phone_no), '') as PRIMARY_WORK_PHONE, isnull(max(internet_email_address), '') as EMAIL_ADDRESS, " & _
                "isnull(max(employed_by_hcid), '') as COMPANY_ABBREV, isNull(max(active_sw), 'N') as active_sw from  (" & _
                "select hanford_id, hanf_pay_no, first_name, last_name, middle_initial, NULL as name_prefix, name_suffix, phone_no, " & _
                "internet_email_address, employed_by_hcid, active_sw from opwhse.dbo.vw_pub_hanford_people union select hanford_id, " & _
                "pay_no as hanf_pay_no, first_name, last_name, middle_name as middle_initial, NULL as name_prefix, name_suffix, " & _
                "work_phone as phone_no, internet_address as internet_email_address, employed_by_hcid, NULL as active_sw " & _
                "from opwhse.dbo.vw_pub_pnnl_associate union select hanford_id, pnnl_pay_no as hanf_pay_no, first_name, last_name, " & _
                "middle_initial, name_prefix, name_suffix, contact_work_phone as phone_no, internet_email_address, company as " & _
                "employed_by_hcid, NULL as active_sw from opwhse.dbo.vw_pub_bmi_employee) a where active_sw = 'Y' and " & _
                "internet_email_address = '" & Me.txtEmail.Text & "' group by hanford_id, " & _
                "last_name, first_name, middle_initial, internet_email_address"
                If (Me.dsCommon.Tables.Contains("PersonData")) Then
                    Me.dsCommon.Tables("PersonData").Rows.Clear()
                End If
                sqlCommonAdapter.Fill(Me.dsCommon, "PersonData")

                'if we found something return true otherwise return false
                If (Me.dsCommon.Tables("PersonData").Rows.Count < 1) Then
                    alert("This person does not match any active personnel.  If you think this may be an error then try the search capabilities of this form to locate the person faster.")
                    Return False
                Else
                    param7.Value = Me.dsCommon.Tables("PersonData").Rows(0).Item("hanford_id")
                End If
            Else
                'get the user's hid from the stored table
                Dim dt As DataTable = CType(Session("PersonTable"), DataTable)
                param7.Value = dt.Rows(Session("SelectedPersonRow")).Item("hanford_id")
            End If
            sqlCommonAction.Parameters.Add(param7)
            Dim param8 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@AUTHENTICATION_CODE", System.Data.SqlDbType.VarChar, 64, "AUTHENTICATION_CODE")

            'generate a password for this user
            If (Trim("" & param7.Value) = "") Then
                param8.Value = generatePassword(Me.txtEmail.Text)
                If (param8.Value = "") Then
                    Return False
                End If
                Session("newUserName") = Me.txtEmail.Text
            Else
                param8.Value = generatePassword(param7.Value)
                If (param8.Value = "") Then
                    Return False
                End If
                Session("newUserName") = param7.Value
            End If
            sqlCommonAction.Parameters.Add(param8)

            'insert the new user
            sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_USER (HANFORD_ID, EMAIL_ADDRESS, USER_CODE, LAST_NAME, FIRST_NAME, " & _
            "MIDDLE_NAME, SUFFIX_NAME, AUTHENTICATION_CODE, PRIMARY_WORK_PHONE, CREATE_DATE, USED_DATE, CREATOR_USER_KEY, ACTIVE_IND, " & _
            "TRAINING_IND, COMPANY_ABBREV, NO_EMAIL_IND) VALUES (@HANFORD_ID, @EMAIL_ADDRESS, 3, @LAST_NAME," & _
            "@FIRST_NAME, @MIDDLE_NAME, @SUFFIX_NAME, @AUTHENTICATION_CODE, @PRIMARY_WORK_PHONE,'" & Now & "','1/1/1970'," & _
            Session("UserID") & "," & IIf(Me.chkActive.Checked, 1, 0) & "," & IIf(Me.chkTQSurvey.Checked, 1, 0) & _
            ",@COMPANY_ABBREV,0)"
            sqlCommonAction.ExecuteNonQuery()

            'get the user's sequence number
            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER where MIDDLE_NAME = '" & Me.txtMiddleName.Text & "' and LAST_NAME = '" & _
            Me.txtLastName.Text & "' and FIRST_NAME = '" & Me.txtFirstName.Text & "' and SUFFIX_NAME = '" & Me.txtSuffixName.Text & "'"
            sqlCommonAdapter.Fill(Me.dsCommon, "UserIDFetch")
            Session("seqUserID") = Me.dsCommon.Tables("UserIDFetch").Rows(0).Item("USER_KEY")
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'attempts to insert permissions for the creator into the permissions table, returns false if it cannot
    Private Function insertPermission(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'parameterized the text input to allow for things like quotes to get through
            sqlCommonAction.Parameters.Clear()

            'determine if the user already exists
            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER_PERMISSION where USER_KEY = " & Session("seqUserID") & _
            " and TEMPLATE_KEY = " & Session("seqTemplate")
            sqlCommonAdapter.Fill(Me.dsCommon, "PermissionList")

            'if the user already exists under the template then update their record for that template
            If (Me.dsCommon.Tables("PermissionList").Rows.Count() > 0) Then
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_USER_PERMISSION set OWNER_IND = " & IIf(Me.chkOwner.Checked, 1, 0)
                sqlCommonAction.CommandText &= ", DELEGATE_IND = 1, USER_IND = 1"
                sqlCommonAction.CommandText &= " where USER_KEY = " & Session("seqUserID") & " and TEMPLATE_KEY = " & Session("seqTemplate")
            Else
                If (Me.dsCommon.Tables("ExistingDelegates").Rows.Count() > 0) Then
                    'Add permissions for the survey
                    sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_USER_PERMISSION (USER_KEY, TEMPLATE_KEY, SURVEY_KEY, CREATOR_USER_KEY, " & _
                        "CREATE_DATE, OWNER_IND, DELEGATE_IND, USER_IND) VALUES (" & Session("seqUserID") & ", " & Session("seqTemplate") & _
                        ", -1, " & Session("UserID") & ", '" & Now & "', " & IIf(Me.chkOwner.Checked, 1, 0) & ", 1, 1)"
                Else
                    'Add permissions for the survey
                    sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_USER_PERMISSION (USER_KEY, TEMPLATE_KEY, SURVEY_KEY, CREATOR_USER_KEY, " & _
                        "CREATE_DATE, OWNER_IND, DELEGATE_IND, USER_IND) VALUES (" & Session("seqUserID") & ", " & Session("seqTemplate") & _
                        ", -1, " & Session("UserID") & ", '" & Now & "', 1, 1, 1)"
                End If
            End If

            'process the command
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function
#End Region

#Region "Update Code"
    'drives the commit and roll-back operations of the update
    Private Function doUpdate() As Boolean
        Try
            'start transaction
            If (myUtility.setupTransaction(sqlCommonAction, Me.commonConn)) Then
                Dim failure As Boolean = False
                'Check for actual information to update
                If (Me.txtFirstName.Text <> "" And Me.txtLastName.Text <> "" And Me.txtEmail.Text <> "") Then
                    If (updateUser(sqlCommonAction, sqlCommonAdapter)) Then
                        If (updatePermission(sqlCommonAction, sqlCommonAdapter)) Then
                            sqlCommonAction.Transaction.Commit()
                            If (Session("seqNewOwnerID") >= 1) Then
                                If Not (sendMail(Session("newOwnerEmail"), "Update", True)) Then
                                    alert("Failed to notify the new Owner/Delegate of this template via e-mail")
                                End If
                                If Not (sendMail(Me.txtEmail.Text, "Update", True)) Then
                                    alert("Failed to notify the relieved Owner of this template via e-mail")
                                End If
                            Else
                                If Not (sendMail(Me.txtEmail.Text, "Update", True)) Then
                                    alert("Failed to notify the Owner/Delegate of this template via e-mail")
                                End If
                            End If
                            'determine if the modifed user is the current user and go back to the
                            'template page if the user is no longer an owner
                            If (Session("UserID") = Session("seqUserID") And Me.chkOwner.Checked = False And Session("UserType") <> 4) Then
                                Session("CurrentPage") = "./Template.aspx"
                                Session("pageSel") = "Template"
                                redirect(Session("CurrentPage"))
                            End If
                        Else
                            sqlCommonAction.Transaction.Rollback()
                            alert("Failed to update this Owner/Delegate.")
                            failure = True
                        End If
                    Else
                        sqlCommonAction.Transaction.Rollback()
                        If Not (Session("seqUserID") = Session("seqNewOwnerID") Or Session("seqNewOwnerID") = -1) Then
                            alert("Failed to update this Owner/Delegate.")
                            failure = True
                        End If
                    End If
                Else
                    alert("You must provide a first and last name as well as an e-mail address and select a user type at minimum.")
                    failure = True
                    sqlCommonAction.Transaction.Rollback()
                End If
                Return failure
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Failed to update this Owner/Delegate.")
                Return True
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Failed to update this Owner/Delegate.")
            Return True
        End Try
    End Function

    'attempts to update a survey, returns false if it cannot
    Private Function updateUser(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        'attempt to update the user's record in the database
        Try
            Session("seqNewOwnerID") = "-2"
            'if the owner checkbox is unchecked then check to see if they are an owner
            If (Me.chkOwner.Checked = False) Then
                'determine if the user is an owner and then determine if the user is the only owner/delegate left
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER_PERMISSION fp, " & myUtility.getExtension() & "SAT_USER fu where fp.TEMPLATE_KEY = " & _
                Session("seqTemplate") & " and (fp.DELEGATE_IND = 1 or fp.OWNER_IND = 1) and fp.USER_KEY = fu.USER_KEY"
                sqlCommonAdapter.SelectCommand = sqlCommonAction
                sqlCommonAdapter.Fill(Me.dsCommon, "TempTable")
                Dim isOwner As Boolean = False
                If (Me.dsCommon.Tables("TempTable").Rows.Count() > 1) Then
                    'determine if the current user is an owner
                    Dim dr As DataRow
                    For Each dr In Me.dsCommon.Tables("TempTable").Rows
                        If (dr.Item("seqUserID") = CType(Session("seqUserID"), Integer) And dr.Item("OWNER_IND")) Then
                            isOwner = True
                        End If
                    Next

                    'determine if any of the users are owners and not the current user
                    If (isOwner) Then
                        For Each dr In Me.dsCommon.Tables("TempTable").Rows
                            If (dr.Item("OWNER_IND") And dr.Item("USER_KEY") <> Session("seqUserID")) Then
                                Session("seqNewOwnerID") = dr.Item("USER_KEY")
                            End If
                        Next

                        'if there were no owners found then find the senior delegate and make that person the new owner
                        If (Session("seqNewOwnerID") = "-2") Then
                            Dim lowDate As DateTime = Now.AddYears(5)
                            For Each dr In Me.dsCommon.Tables("TempTable").Rows
                                If (dr.Item("CREATE_DATE") < lowDate) Then
                                    lowDate = dr.Item("CREATE_DATE")
                                    Session("seqNewOwnerID") = dr.Item("USER_KEY")
                                    Session("newOwnerEmail") = dr.Item("EMAIL_ADDRESS")
                                End If
                            Next
                        Else
                            Session("seqNewOwnerID") = "-2"
                        End If

                        'if the new owner is the current owner notify the user and abort
                        If (Session("seqUserID") = Session("seqNewOwnerID")) Then
                            alert("This person is the only senior Owner/Delegate left.  Try creating a new co-owner and then remove this person\'s status as owner.")
                            Return False
                        End If
                    Else
                        Session("seqNewOwnerID") = -2
                    End If
                Else
                    Session("seqNewOwnerID") = -1
                End If
            End If

            'if the user was not an owner or there is only 1 record indicating they are the only person with modification
            'access then alert the user otherwise allow the update to continue.
            If (Session("seqNewOwnerID") <> -1) Then
                'clear out table data before use
                If (Me.dsCommon.Tables.Contains("ExisitingUsers")) Then
                    Me.dsCommon.Tables("ExistingUsers").Rows.Clear()
                End If
                If (Me.dsCommon.Tables.Contains("PersonData")) Then
                    Me.dsCommon.Tables("PersonData").Rows.Clear()
                End If

                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_USER SET EMAIL_ADDRESS = @EMAIL_ADDRESS, FIRST_NAME = @FIRST_NAME" & _
                ", LAST_NAME = @LAST_NAME, MIDDLE_NAME = @MIDDLE_NAME, SUFFIX_NAME = @SUFFIX_NAME, " & _
                "PRIMARY_WORK_PHONE = @PRIMARY_WORK_PHONE, COMPANY_ABBREV = @COMPANY_ABBREV, " & _
                " TRAINING_IND = " & IIf(Me.chkTQSurvey.Checked, 1, 0) & ", ACTIVE_IND = " & IIf(Me.chkActive.Checked, 1, 0) & _
                " where USER_KEY = " & Session("seqUserID")

                'parameterized the text input to allow for things like quotes to get through
                Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@EMAIL_ADDRESS", System.Data.SqlDbType.VarChar, 128, "EMAIL_ADDRESS")
                param0.Value = Me.txtEmail.Text
                sqlCommonAction.Parameters.Add(param0)
                Dim param1 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@FIRST_NAME", System.Data.SqlDbType.VarChar, 50, "FIRST_NAME")
                param1.Value = Me.txtFirstName.Text
                sqlCommonAction.Parameters.Add(param1)
                Dim param2 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@LAST_NAME", System.Data.SqlDbType.VarChar, 50, "LAST_NAME")
                param2.Value = Me.txtLastName.Text
                sqlCommonAction.Parameters.Add(param2)
                Dim param3 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@MIDDLE_NAME", System.Data.SqlDbType.VarChar, 50, "MIDDLE_NAME")
                param3.Value = Me.txtMiddleName.Text
                sqlCommonAction.Parameters.Add(param3)
                Dim param4 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@SUFFIX_NAME", System.Data.SqlDbType.VarChar, 50, "SUFFIX_NAME")
                param4.Value = Me.txtSuffixName.Text
                sqlCommonAction.Parameters.Add(param4)
                Dim param5 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@PRIMARY_WORK_PHONE", System.Data.SqlDbType.VarChar, 32, "PRIMARY_WORK_PHONE")
                param5.Value = Me.txtPhoneNumber.Text
                sqlCommonAction.Parameters.Add(param5)
                Dim param6 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@COMPANY_ABBREV", System.Data.SqlDbType.VarChar, 32, "COMPANY_ABBREV")
                param6.Value = Me.txtCompany.Text
                sqlCommonAction.Parameters.Add(param6)

                sqlCommonAction.ExecuteNonQuery()
            Else
                alert("This is the only Owner/Delegate left.  You may not remove this person as owner at this time.  Try creating a new co-owner and then remove this owner.")
                Return False
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
        Return True
    End Function

    'attempts to update the permissions for the user, returns false if it cannot
    Private Function updatePermission(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        'attempt to update the permissions for the user in the database
        'remove the parameters from the previous function
        sqlCommonAction.Parameters.Clear()

        If (Session("seqNewOwnerID") >= 1) Then
            Try
                'try to update the user's record in the permissions table.
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_USER_PERMISSION SET DELEGATE_IND = 1, OWNER_IND = 1" & _
                " where USER_KEY = " & Session("seqNewOwnerID") & " and TEMPLATE_KEY = " & Session("seqTemplate")

                'do the update
                sqlCommonAction.ExecuteNonQuery()

                'if the user being modified is also the user using this page then update their permission
                If (Session("seqUserID") = Session("UserID") And Session("seqNewOwnerID") <> Session("UserID") And Session("UserType") <> 4) Then
                    Session("isTemplateOwner") = IIf(Me.chkOwner.Checked, True, False)
                End If

                'if the user being modified is not the senior delegate then go ahead and negate their owner permission
                If (Session("seqUserID") <> Session("seqNewOwnerID")) Then
                    sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_USER_PERMISSION Set DELEGATE_IND = 1, OWNER_IND = 0" & _
                    " where USER_KEY = " & Session("seqUserID") & " and TEMPLATE_KEY = " & Session("seqTemplate")

                    'do the update
                    sqlCommonAction.ExecuteNonQuery()
                End If
                alert("There are no owners left.  The senior delegate has been given owner permissions.")
                Return True
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                Return False
            End Try
        Else
            Dim blnFailure As Boolean = False
            Try
                'try to insert the user's record in the permissions table.
                sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_USER_PERMISSION (USER_KEY, TEMPLATE_KEY, SURVEY_KEY, CREATOR_USER_KEY, " & _
                "CREATE_DATE, OWNER_IND, DELEGATE_IND, USER_IND) VALUES (" & Session("seqUserID") & ", " & Session("seqTemplate") & _
                ", -1, " & Session("UserID") & ", '" & Now & "', " & IIf(Me.chkOwner.Checked, 1, 0) & ", 1, 1)"

                'do the insert
                sqlCommonAction.ExecuteNonQuery()
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                blnFailure = True
            End Try

            If (blnFailure) Then
                Try
                    'try to insert the user's record in the permissions table.
                    sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_USER_PERMISSION SET DELEGATE_IND = 1, OWNER_IND = " & _
                    IIf(Me.chkOwner.Checked, 1, 0) & ", USER_IND = 1 where USER_KEY = " & Session("seqUserID") & " and TEMPLATE_KEY = " & Session("seqTemplate")

                    'do the update
                    sqlCommonAction.ExecuteNonQuery()
                Catch ex As Exception
                    Trace.Warn(ex.ToString)
                    Return False
                End Try
            End If

            'if the user being modified is also the user using this page then update their permission
            If (Session("seqUserID") = Session("UserID") And Session("UserType") <> 4) Then
                Session("isTemplateOwner") = IIf(Me.chkOwner.Checked, True, False)
            End If
            Return True
        End If
    End Function
#End Region

#Region "Delete Code"
    'deletes the current delegate from the fempPermissions table
    Private Function doDelete() As Boolean
        Try
            'start transaction
            If (myUtility.setupTransaction(sqlCommonAction, Me.commonConn)) Then
                Dim failure As Boolean = False

                'do the deletion
                If (deletePermission(sqlCommonAction, sqlCommonAdapter)) Then
                    sqlCommonAction.Transaction.Commit()

                    'if the user being modified is also the user using this page then redirect to the template page and select no template
                    If (Session("seqUserID") = Session("UserID") And Session("UserType") <> 4) Then
                        Session("CurrentPage") = "./Template.aspx"
                        Session("pageSel") = "Template"
                        Session("seqTemplate") = -1
                        redirect(Session("CurrentPage"))
                    End If

                    'determine if we need to take a step back 
                    If (Session("DelegateMax") = 1) Then
                        Session("DelegateRow") = 0
                        Session("seqUserID") = Me.dsCommon.Tables("Delegates").Rows(Session("DelegateRow")).Item("USER_KEY")
                    ElseIf ((Session("DelegateRow") + 1) = Session("DelegateMax")) Then
                        Session("DelegateMax") -= 1
                        Session("DelegateRow") -= 1
                        Session("seqUserID") = Me.dsCommon.Tables("Delegates").Rows(Session("DelegateRow")).Item("USER_KEY")
                    Else
                        Session("DelegateMax") -= 1
                        Session("seqUserID") = Me.dsCommon.Tables("Delegates").Rows(Session("DelegateRow") + 1).Item("USER_KEY")
                    End If

                    'notify the user of their status change
                    If Not (sendMail(Me.txtEmail.Text, "Delete", True)) Then
                        alert("Failed to notify the Owner/Delegate of their status change")
                    End If
                Else
                    sqlCommonAction.Transaction.Rollback()
                    failure = True
                End If
                Return failure
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Failed to delete this Owner/Delegate.")
                Return True
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Failed to delete this Owner/Delegate.")
            Return True
        End Try
    End Function

    'attempts to delete a user from the fempPermission table.  Assigns ownership to the suvey administrator if
    'the person being deleted is the last person.  Returns false if it cannot for any reason.
    Private Function deletePermission(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'determine the type of user being deleted
            Dim row As DataRow
            For Each row In Me.dsCommon.Tables("Delegates").Rows()
                If (CType(row.Item("USER_KEY"), String) = Session("seqUserID")) Then
                    If (row.Item("OWNER_IND")) Then
                        'determine if the user being deleted is the only owner
                        sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER_PERMISSION where TEMPLATE_KEY = " & Session("seqTemplate") & _
                        " and OWNER_IND = 1"
                        sqlCommonAdapter.SelectCommand = sqlCommonAction
                        sqlCommonAdapter.Fill(Me.dsCommon, "OwnersCount")
                        If (Me.dsCommon.Tables("OwnersCount").Rows.Count() < 2) Then
                            alert("You may not delete the only remaining Owner/Delegate.  You must make someone else an owner of this template before removing this owner.")
                            Return False
                        End If
                    End If
                End If
            Next

            'Reduce their access to user
            sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_USER_PERMISSION set OWNER_IND = 0, " & _
            "DELEGATE_IND = 0, USER_IND = 1 where TEMPLATE_KEY = " & Session("seqTemplate") & _
            "and USER_KEY = " & Session("seqUserID")
            sqlCommonAction.ExecuteNonQuery()

            'if the user being modified is also the user using this page then update their permission
            If (Session("seqUserID") = Session("UserID") And Session("UserType") <> 4) Then
                Session("isTemplateOwner") = False
                Session("isTemplateDelegate") = False
            End If
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function
#End Region

#Region "Reset Code"
    'resets the page, will remove any new item being worked on at the end of the list
    Private Function doReset() As Boolean
        If Not (loadData()) Then
            alert("Failed to load the delegates/owners for this survey.")
        Else
            Try
                If (Session("seqUserID") = -1) Then
                    Me.dsCommon.Tables("Delegates").Rows.Add(Me.dsCommon.Tables("Delegates").NewRow())
                    Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("EMAIL_ADDRESS") = ""
                    Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("FIRST_NAME") = ""
                    Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("LAST_NAME") = ""
                    Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("MIDDLE_NAME") = ""
                    Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("SUFFIX_NAME") = ""
                    Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("PRIMARY_WORK_PHONE") = ""
                    Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("COMPANY_ABBREV") = ""
                    Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("USER_CODE") = -1
                    Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("TRAINING_IND") = 0
                    Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("ACTIVE_IND") = 1
                    Me.dsCommon.Tables("Delegates").Rows(Session("DelegateMax")).Item("OWNER_IND") = 0
                End If
                Return False
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error resetting the delegate/owner of this template.")
                Return True
            End Try
        End If
    End Function
#End Region

#Region "Password Generation"
    'Generates the user's password
    Private Function generatePassword(ByVal strValue As String) As String
        Trace.Warn("Generating Password")

        'Generate password for the user
        Dim proxyAuthentication As New Authentication.Service1
        strNewPassword = proxyAuthentication.createPasswordString()

        'Generate password hash for the user
        Dim strPassword As String = ""
        strPassword = myUtility.hashPassword(strNewPassword, strValue)

        Return strPassword
    End Function
#End Region

#Region "Send Mail"
    'Sends an e-mail to users with changed passwords, except level 0 users.
    Private Function sendMail(ByVal Email As String, ByVal strUserName As String, Optional ByVal inSystem As Boolean = False) As Boolean
        Trace.Warn("Sending user an email")
        'Set up the mail messsage
        Try
            Dim strFrom As String = "survey.administrator@pnl.gov"
            Dim strTo As String = Email
            Dim strbldrSubject As New StringBuilder("Survey Admin Tool Template Delegation.")

            'get the current user's name and e-mail
            Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
            Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
            sqlCommonAction.Connection = Me.commonConn
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER where USER_KEY = " & Session("UserID")
            sqlCommonAdapter.Fill(Me.dsCommon, "ModifyID")

            If (Me.dsCommon.Tables("ModifyID").Rows.Count < 1) Then
                alert("Unable to locate your information.  You may have been removed from the database in the last few minutes.  Contact the Survey Administrator immediately.")
                Return False
            End If

            Dim strbldrMessage As New StringBuilder
            'notify the user of their status change
            If (strUserName = "Delete") Then
                If (Me.chkOwner.Checked = True) Then
                    strbldrMessage.Append("You have been removed as a Owner of " & Session("TemplateName") & " template by")
                Else
                    strbldrMessage.Append("You have been removed as a Delegate of " & Session("TemplateName") & " template by")
                End If
                strbldrMessage.Append(" " & Me.dsCommon.Tables("ModifyID").Rows(0).Item("FIRST_NAME") & " ")
                strbldrMessage.Append(Me.dsCommon.Tables("ModifyID").Rows(0).Item("LAST_NAME") & ".")
                strbldrMessage.Append("  You may contact this person at this e-mail address: <b>")
                strbldrMessage.Append(Me.dsCommon.Tables("ModifyID").Rows(0).Item("EMAIL_ADDRESS") & "</b>")
            ElseIf (strUserName = "Update") Then
                strbldrMessage.Append("You have been given ")
                If (Me.chkOwner.Checked = True) Then
                    strbldrMessage.Append("Owner access to the ")
                    strbldrMessage.Append(Session("TemplateName") & " template or your user information as a Owner has been altered by")
                Else
                    strbldrMessage.Append("Delegate access to the ")
                    strbldrMessage.Append(Session("TemplateName") & " template or your user information as a Delegate has been altered by")
                End If
                strbldrMessage.Append(" " & Me.dsCommon.Tables("ModifyID").Rows(0).Item("FIRST_NAME") & " ")
                strbldrMessage.Append(Me.dsCommon.Tables("ModifyID").Rows(0).Item("LAST_NAME") & ".")
                strbldrMessage.Append("  You may contact this person at this e-mail address: <b>")
                strbldrMessage.Append(Me.dsCommon.Tables("ModifyID").Rows(0).Item("EMAIL_ADDRESS") & "</b><br/>")
            Else
                strbldrMessage.Append("You have been given ")
                If (Me.chkOwner.Checked = True) Then
                    strbldrMessage.Append("Owner access to the ")
                Else
                    strbldrMessage.Append("Delegate access to the ")
                End If
                strbldrMessage.Append(Session("TemplateName") & " template by")
                strbldrMessage.Append(" " & Me.dsCommon.Tables("ModifyID").Rows(0).Item("FIRST_NAME") & " ")
                strbldrMessage.Append(Me.dsCommon.Tables("ModifyID").Rows(0).Item("LAST_NAME") & ".")
                strbldrMessage.Append("  You may contact this person at this e-mail address: <b>")
                strbldrMessage.Append(Me.dsCommon.Tables("ModifyID").Rows(0).Item("EMAIL_ADDRESS") & "</b><br/>")
                If Not (inSystem) Then
                    strbldrMessage.Append("Please Make a note of your password.<br/> If you wish to change your password ")
                    strbldrMessage.Append("you may do so at the ")
                    strbldrMessage.Append("<a href='http://" & Request.ServerVariables.Item("SERVER_NAME") & "/Survey Admin/Default.aspx'>Survey Administration</a> website. <p></p>")
                    strbldrMessage.Append("Your User Name is: " & strUserName & "<br/>")
                    strbldrMessage.Append("Your Password is: " & strNewPassword)
                End If
            End If

            'send the e-mail
            If (myUtility.genericSendMail(strFrom, strTo, strbldrSubject.ToString, strbldrMessage.ToString)) Then
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
End Class