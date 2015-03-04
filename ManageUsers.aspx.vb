Imports System.Collections.Specialized
Imports System.Text
Public Class ManageUsers
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()
        Me.dsCommon = New System.Data.DataSet
        Me.commonConn = New System.Data.SqlClient.SqlConnection
        Me.sqlUsers = New System.Data.SqlClient.SqlDataAdapter
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'dsCommon
        '
        Me.dsCommon.DataSetName = "NewDataSet"
        Me.dsCommon.Locale = New System.Globalization.CultureInfo("en-US")
        '
        'sqlUsers
        '
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected WithEvents imgTitle As System.Web.UI.WebControls.Image
    Protected WithEvents hlpUser As System.Web.UI.WebControls.Image
    Protected WithEvents lblUser As System.Web.UI.WebControls.Label
    Protected WithEvents ddlUser As System.Web.UI.WebControls.DropDownList
    Protected WithEvents hlpHIDPayrollName As System.Web.UI.WebControls.Image
    Protected WithEvents lblHIDPayrollName As System.Web.UI.WebControls.Label
    Protected WithEvents txtHIDPayrollName As System.Web.UI.WebControls.TextBox
    Protected WithEvents btnHIDPayrollNameFind As System.Web.UI.WebControls.Button
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
    Protected WithEvents hlpUserType As System.Web.UI.WebControls.Image
    Protected WithEvents lblUserType As System.Web.UI.WebControls.Label
    Protected WithEvents ddlUserType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents hlpPNNLTraining As System.Web.UI.WebControls.Image
    Protected WithEvents chkTQSurvey As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblTraining As System.Web.UI.WebControls.Label
    Protected WithEvents hlpActive As System.Web.UI.WebControls.Image
    Protected WithEvents chkActive As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblActive As System.Web.UI.WebControls.Label
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents chkFindCurrentUsers As System.Web.UI.WebControls.CheckBox
    Protected WithEvents hlpImportList As System.Web.UI.WebControls.Image
    Protected WithEvents chkImportList As System.Web.UI.WebControls.CheckBox
    Protected WithEvents hlpUserList As System.Web.UI.WebControls.Image
    Protected WithEvents lblUserList As System.Web.UI.WebControls.Label
    Protected WithEvents txtUserList As System.Web.UI.WebControls.TextBox
    Protected myUtility As New Utility
    Protected WithEvents hlpPasswordReset As System.Web.UI.WebControls.Image
    Protected WithEvents chkPasswordReset As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblPasswordReset As System.Web.UI.WebControls.Label
    Protected WithEvents Td1 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents dsCommon As System.Data.DataSet
    Protected WithEvents sqlUsers As System.Data.SqlClient.SqlDataAdapter
    Protected sqlCommonAdapter As New System.Data.SqlClient.SqlDataAdapter
    Protected sqlCommonAction As New System.Data.SqlClient.SqlCommand
    Protected strNewPassword As String
    Protected WithEvents btnInsert As System.Web.UI.WebControls.Button
    Protected WithEvents btnUpdate As System.Web.UI.WebControls.Button
    Protected WithEvents btnDelete As System.Web.UI.WebControls.Button
    Protected WithEvents btnReset As System.Web.UI.WebControls.Button
    Protected WithEvents btnClear As System.Web.UI.WebControls.Button
    Protected WithEvents commonConn As System.Data.SqlClient.SqlConnection
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
        ElseIf (myUtility.checkUserType(4)) Then
            redirect(Session("CurrentPage"))
        Else
            Session("CurrentPage") = "./ManageUsers.aspx"
            Session("pageSel") = "ManageUsers"
        End If

        'if the user's password has expired send them to the password reset page
        If (Session("isExpired") = "True") Then
            Session("CurrentPage") = "./Password.aspx"
            Session("pageSel") = "Password"
            redirect(Session("CurrentPage"))
        End If

        'Set the connection based on the machine
        myUtility.getConnection(Me.commonConn)

        'Check for user selection from the comments list if we're sure it's not a simple post-back
        If Not (IsPostBack) Then
            'get anything we can off the request line if we're not coming back from processing an e-mail address
            If (Session("ProcessEmail") <> "True") Then
                myUtility.getRequest(Page.Request.Url.Query())
            End If

            'set the page to it's default layout
            If (Me.chkImportList.Checked) Then
                Session("ImportList") = True
            ElseIf (Session("ProcessEmail") = "True") Then
                Session("ImportList") = True
                Me.chkImportList.Checked = True
            Else
                Session("ImportList") = False
            End If

            'if we're coming from the PersonSelect page and the process was not stopped load the selected user's data
            'otherwise load the page like normal.
            If (Session("ProcessStopped") = "False") Then
                'refresh the tables
                Dim oldID As Integer = Session("manageUserID")
                Session("manageUserID") = -1

                'load the page data
                If Not (loadData()) Then
                    alert("Failed to load the Survey Tool users.")
                Else
                    'restore the person data and find the exact user selected
                    Dim dt As DataTable = Session("PersonTable")
                    Dim dr As DataRow = dt.Rows(Session("selectedPersonRow"))

                    'try to get the user's id.
                    'determine if the user already exists
                    Dim isExistingUser As Boolean = False
                    isExistingUser = checkUserExistence(dr)

                    If (isExistingUser) Then
                        Session("manageUserID") = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("USER_KEY")
                    End If

                    'if the user already exists load their data
                    If (isExistingUser) Then
                        'populate the controls with the selected individual's information
                        Me.txtEmail.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("EMAIL_ADDRESS")
                        Me.txtFirstName.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("FIRST_NAME")
                        Me.txtLastName.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("LAST_NAME")
                        Me.txtMiddleName.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("MIDDLE_NAME")
                        Me.txtSuffixName.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("SUFFIX_NAME")
                        Me.txtPhoneNumber.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("PRIMARY_WORK_PHONE")
                        Me.txtCompany.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("COMPANY_ABBREV")
                        Me.ddlUserType.SelectedIndex = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("USER_CODE") + 1
                        Me.chkTQSurvey.Checked = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("TRAINING_IND")
                        Me.chkActive.Checked = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("ACTIVE_IND")
                        Me.chkPasswordReset.Checked = 0
                        Me.chkFindCurrentUsers.Checked = CType(Session("CurrentUsers"), Boolean)
                        Session("UserRow") = findRow("Users", Session("manageUserID"))
                        Me.ddlUser.SelectedIndex = Session("UserRow") + 1
                    Else
                        'populate the controls with the selected individual's information
                        Me.txtEmail.Text = dr.Item("EMAIL_ADDRESS")
                        Me.txtFirstName.Text = dr.Item("FIRST_NAME")
                        Me.txtLastName.Text = dr.Item("LAST_NAME")
                        Me.txtMiddleName.Text = dr.Item("MIDDLE_NAME")
                        Me.txtSuffixName.Text = dr.Item("SUFFIX_NAME")
                        Me.txtPhoneNumber.Text = dr.Item("PRIMARY_WORK_PHONE")
                        Me.txtCompany.Text = dr.Item("COMPANY_ABBREV")
                        Me.ddlUserType.SelectedIndex = 0
                        Me.chkTQSurvey.Checked = 0
                        Me.chkActive.Checked = 1
                        Me.chkFindCurrentUsers.Checked = CType(Session("CurrentUsers"), Boolean)
                        Me.chkPasswordReset.Checked = 0
                        Session("UserRow") = findRow("Users", Session("manageUserID"))
                        Me.ddlUser.SelectedIndex = Session("UserRow") + 1
                    End If

                    'set the page options
                    setPageOptions()

                    Me.btnInsert.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                    Me.btnUpdate.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                    Me.btnDelete.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                    Me.btnReset.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                    Me.btnClear.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")

                    Session("ProcessStopped") = "True"
                End If
            Else
                'set the user id if the user id to -1 for initial page loads
                Session("manageUserID") = -1

                'load the page data
                If Not (loadData()) Then
                    alert("Failed to load the delegate/owners for this template.")
                Else
                    'Populate the controls on the page
                    setControls()

                    'set the page options
                    setPageOptions()

                    Session("ProcessStopped") = "True"
                    Me.btnInsert.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                    Me.btnUpdate.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                    Me.btnDelete.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                    Me.btnReset.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                    Me.btnClear.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                End If

                'process an incoming user input request from the person input page
                If (Session("ProcessEmail") = "True") Then
                    Me.txtUserList.Text = Session("PersonInputEmailReturn")
                    Me.ddlUserType.SelectedIndex = Session("PersonInputUserType") + 1
                    Dim failure As Boolean = False
                    If Not (doInsert()) Then
                        failure = True
                    End If
                    Session("ProcessEmail") = "False"

                    Dim blnContinue As Boolean = True
                    'determine if we need to reload the data
                    If Not (loadData()) Then
                        blnContinue = False
                        alert("Failed to load the groups for the current survey.")
                    End If

                    If (blnContinue) Then
                        'reset the page controls
                        If Not (failure) Then
                            setControls()
                        End If
                        Session("TransExists") = False

                        'set the page options
                        setPageOptions()
                    End If
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
#End Region

#Region "Data Load"
    'Loads data into the form
    Private Function loadData(Optional ByVal override As Boolean = False) As Boolean
        Trace.Warn("Loading Data")
        'Set up the common data adapter and select command
        Me.sqlCommonAction.Connection = Me.commonConn
        Me.sqlCommonAdapter.SelectCommand = sqlCommonAction

        If (Me.chkImportList.Checked) Then
            'Fill the User Level List
            If Not (override) Then
                If (loadUserTypes()) Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return True
            End If
        Else
            'Fill User List
            If (loadUsers(override)) Then
                'Fill the User Level List
                If Not (override) Then
                    If (loadUserTypes()) Then
                        Return True
                    Else
                        Return False
                    End If
                Else
                    Return True
                End If
            Else
                Return False
            End If
        End If
    End Function

    'loads the users in the system
    Private Function loadUsers(Optional ByVal override As Boolean = False) As Boolean
        Trace.Warn("Loading Users")
        Try
            'clear existing data
            If (Me.dsCommon.Tables.Contains("Users")) Then
                Me.dsCommon.Tables("Users").Rows.Clear()
            End If

            'get a list of users in the system
            Me.sqlCommonAction.CommandText = "SELECT LAST_NAME + ', ' + FIRST_NAME + ' ' + MIDDLE_NAME + '-' + HANFORD_ID + " & _
            "'-' + EMAIL_ADDRESS as UserName, USER_KEY, USER_CODE, LAST_NAME, FIRST_NAME, MIDDLE_NAME, SUFFIX_NAME, " & _
            "AUTHENTICATION_CODE, PRIMARY_WORK_PHONE, HANFORD_ID, EMAIL_ADDRESS, CREATE_DATE, USED_DATE, CREATOR_USER_KEY, ACTIVE_IND, TRAINING_IND, " & _
            "COMPANY_ABBREV, NO_EMAIL_IND FROM " & myUtility.getExtension() & "SAT_USER where not EMAIL_ADDRESS = ' ' " & _
            "and not LAST_NAME = ' ' and not FIRST_NAME = ' ' ORDER BY UserName"

            Me.sqlCommonAdapter.Fill(Me.dsCommon, "Users")
            Me.ddlUser.DataSource = Me.dsCommon.Tables("Users")
            Me.ddlUser.DataTextField = "UserName"
            Me.ddlUser.DataValueField = "USER_KEY"
            Me.ddlUser.DataBind()
            Dim li As System.Web.UI.WebControls.ListItem = New System.Web.UI.WebControls.ListItem
            li.Text = "--Select a User--"
            li.Value = -1
            Me.ddlUser.Items.Insert(0, li)

            'restore the selection of the current user
            If (override) Then
                Dim dr As DataRow
                Dim index As Integer = 0
                For Each dr In Me.dsCommon.Tables("Users").Rows
                    If (Session("manageUserID") = dr.Item("USER_KEY")) Then
                        Me.ddlUser.SelectedIndex = index + 1
                    End If
                    index += 1
                Next
            End If

            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Unable to load the users currently in the database.")
            Return False
        End Try
    End Function

    'loads the user levels in the system
    Private Function loadUserTypes() As Boolean
        Trace.Warn("Loading Levels")
        Try
            'clear out existing data
            If (Me.dsCommon.Tables.Contains("UserTypes")) Then
                Me.dsCommon.Tables("UserTypes").Rows.Clear()
            End If

            'get the user levels available
            Me.sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER_TYPE order by USER_CODE"
            Me.sqlCommonAdapter.Fill(Me.dsCommon, "UserTypes")
            Me.ddlUserType.DataSource = Me.dsCommon.Tables("UserTypes")
            Me.ddlUserType.DataTextField = "USER_TITLE"
            Me.ddlUserType.DataValueField = "USER_CODE"
            Me.ddlUserType.DataBind()
            Dim li As System.Web.UI.WebControls.ListItem = New System.Web.UI.WebControls.ListItem
            li.Text = "--Select a User Type--"
            li.Value = -1
            Me.ddlUserType.Items.Insert(0, li)
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Unable to load the user types")
            Return False
        End Try
    End Function

    'gets the user information for a specific user
    Private Function userInfo(ByVal strEmail As String) As Boolean
        'dump existing data
        If (Me.dsCommon.Tables.Contains("SpecificUserInfo")) Then
            Me.dsCommon.Tables("SpecificUserInfo").Rows.Clear()
        End If

        'get the current user's name and e-mail
        Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
        Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
        sqlCommonAdapter.SelectCommand = sqlCommonAction
        sqlCommonAction.Connection = Me.commonConn
        sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER where EMAIL_ADDRESS = '" & strEmail & "'"
        sqlCommonAdapter.Fill(Me.dsCommon, "SpecificUserInfo")

        If (Me.dsCommon.Tables("SpecificUserInfo").Rows.Count < 1) Then
            alert("Unable to locate " & strEmail & "'s information.")
            Return False
        Else
            Return True
        End If
    End Function
#End Region

#Region "Control Settings"
    'populates the controls with the selected user's data
    Private Sub setControls(Optional ByVal override As String = "")
        Trace.Warn("Setting Controls")
        'set the basic page controls based on whether or not we are working with a new survey
        If (Session("manageUserID") > 0) Then
            Try
                Me.txtEmail.Text = Me.dsCommon.Tables("Users").Rows(Session("UserRow")).Item("EMAIL_ADDRESS")
                Me.txtFirstName.Text = Me.dsCommon.Tables("Users").Rows(Session("UserRow")).Item("FIRST_NAME")
                Me.txtLastName.Text = Me.dsCommon.Tables("Users").Rows(Session("UserRow")).Item("LAST_NAME")
                Me.txtMiddleName.Text = Me.dsCommon.Tables("Users").Rows(Session("UserRow")).Item("MIDDLE_NAME")
                Me.txtSuffixName.Text = Me.dsCommon.Tables("Users").Rows(Session("UserRow")).Item("SUFFIX_NAME")
                Me.txtPhoneNumber.Text = Me.dsCommon.Tables("Users").Rows(Session("UserRow")).Item("PRIMARY_WORK_PHONE")
                Me.txtCompany.Text = Me.dsCommon.Tables("Users").Rows(Session("UserRow")).Item("COMPANY_ABBREV")
                Me.ddlUserType.SelectedIndex = Me.dsCommon.Tables("Users").Rows(Session("UserRow")).Item("USER_CODE") + 1
                Me.chkTQSurvey.Checked = Me.dsCommon.Tables("Users").Rows(Session("UserRow")).Item("TRAINING_IND")
                Me.chkActive.Checked = Me.dsCommon.Tables("Users").Rows(Session("UserRow")).Item("ACTIVE_IND")
                Me.ddlUser.SelectedIndex = findRow("Users", Session("ManageUserID")) + 1
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
                Me.ddlUserType.SelectedIndex = 0
                Me.chkTQSurvey.Checked = 0
                Me.chkActive.Checked = 1
                Me.chkPasswordReset.Checked = 0
            End If
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
            If (Trim("" & dr.Item("EMAIL_ADDRESS")) = "") Then
                sqlCommonAction.CommandText = "select fu.HANFORD_ID, fu.USER_KEY, fu.FIRST_NAME, " & _
                "fu.LAST_NAME, fu.MIDDLE_NAME, fu.SUFFIX_NAME, " & _
                "fu.PRIMARY_WORK_PHONE, fu.EMAIL_ADDRESS, fu.COMPANY_ABBREV, " & _
                "fu.USER_CODE, fu.TRAINING_IND, fu.ACTIVE_IND from " & myUtility.getExtension() & "SAT_USER fu " & _
                " where fu.EMAIL_ADDRESS = '" & dr.Item("HANFORD_ID") & "'"
            Else
                sqlCommonAction.CommandText = "select fu.HANFORD_ID, fu.USER_KEY, fu.FIRST_NAME, " & _
                "fu.LAST_NAME, fu.MIDDLE_NAME, fu.SUFFIX_NAME, " & _
                "fu.PRIMARY_WORK_PHONE, fu.EMAIL_ADDRESS, fu.COMPANY_ABBREV, " & _
                "fu.USER_CODE, fu.TRAINING_IND, fu.ACTIVE_IND from " & myUtility.getExtension() & "SAT_USER fu " & _
                " where fu.EMAIL_ADDRESS = '" & dr.Item("EMAIL_ADDRESS") & "'"
            End If
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

    'sets up the page for importing a list or not based on the status of the import list checkbox
    Private Sub chkImportList_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkImportList.CheckedChanged
        Session("JSAction") = ""

        If (Me.chkImportList.Checked) Then
            Session("ImportList") = True
            Me.ddlUserType.SelectedIndex = 0
            Session("manageUserID") = -1
            'determine if we need to reload the data
            Dim blnContinue As Boolean = True
            If Not (loadData()) Then
                blnContinue = False
                alert("Failed to load the groups for the current survey.")
            End If

            If (blnContinue) Then
                setControls()

                'set the page options
                setPageOptions()
            End If
        Else
            Session("ImportList") = False
            Me.ddlUserType.SelectedIndex = 0
            Session("manageUserID") = -1
            'determine if we need to reload the data
            Dim blnContinue As Boolean = True
            If Not (loadData()) Then
                blnContinue = False
                alert("Failed to load the groups for the current survey.")
            End If

            If (blnContinue) Then
                setControls()

                'set the page options
                setPageOptions()
            End If
        End If
    End Sub
#End Region

#Region "Settings"
    'set the page options
    Public Sub setPageOptions()
        If (Session("manageUserID") > 0) Then
            'existing users get update/delete
            Session("PageOptions") = 1
        Else
            'new users get insert
            Session("PageOptions") = 2
        End If
    End Sub
#End Region

#Region "Find"
    'finds the row of the user we're looking for
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
            "opwhse.dbo.vw_pub_hanford_people hp where fu.HANFORD_ID = hp.hanford_id " & _
            "and hp.active_sw = 'Y' and fu.HANFORD_ID = '" & LookUpString & "'"
            sqlCommonAdapter.Fill(Me.dsCommon, "ExistingUsers")

            'if we found something then they already exist
            If (Me.dsCommon.Tables("ExistingUsers").Rows.Count >= 1) Then
                Session("PersonTable") = Me.dsCommon.Tables("ExistingUsers")
                Session("manageUserID") = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("USER_KEY")
                alreadyExists = True
            End If

            If (alreadyExists) Then
                Return True
            Else
                If Not (Me.chkFindCurrentUsers.Checked) Then
                    'set up and process the sql statement
                    sqlCommonAction.CommandText = "select isnull(max(hanford_id), '') as hanford_id, isnull(max(hanf_pay_no), '') as " & _
                    "hanf_pay_no, max(first_name) as first_name, max(last_name) as last_name, isnull(max(middle_initial), '') as " & _
                    "MIDDLE_NAME, isnull(max(name_prefix), '') as name_prefix, isnull(max(name_suffix), '') as SUFFIX_NAME, " & _
                    "isnull(max(phone_no), '') as PRIMARY_WORK_PHONE, isnull(max(internet_email_address), '') as EMAIL_ADDRESS, " & _
                    "isnull(max(employed_by_hcid), '') as COMPANY_ABBREV, isNull(max(active_sw), 'N') as ACTIVE_IND from  (" & _
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
            " and hp.hanf_pay_no = '" & LookUpString & "'" & " and fu.HANFORD_ID <> ''"
            sqlCommonAdapter.Fill(Me.dsCommon, "ExistingUsers")

            'if we found something return true otherwise return false
            If (Me.dsCommon.Tables("ExistingUsers").Rows.Count >= 1) Then
                Session("PersonTable") = Me.dsCommon.Tables("ExistingUsers")
                Session("manageUserID") = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("USER_KEY")
                alreadyExists = True
            End If

            If (alreadyExists) Then
                Return True
            Else
                If Not (Me.chkFindCurrentUsers.Checked) Then
                    'set up and process the sql statement
                    sqlCommonAction.CommandText = "select isnull(max(hanford_id), '') as hanford_id, isnull(max(hanf_pay_no), '') as " & _
                    "hanf_pay_no, max(first_name) as first_name, max(last_name) as last_name, isnull(max(middle_initial), '') as " & _
                    "MIDDLE_NAME, isnull(max(name_prefix), '') as name_prefix, isnull(max(name_suffix), '') as SUFFIX_NAME, " & _
                    "isnull(max(phone_no), '') as PRIMARY_WORK_PHONE, isnull(max(internet_email_address), '') as email_address, " & _
                    "isnull(max(employed_by_hcid), '') as COMPANY_ABBREV, isNull(max(active_sw), 'N') as ACTIVE_IND from  (" & _
                    "select hanford_id, hanf_pay_no, first_name, last_name, middle_initial, NULL as name_prefix, name_suffix, phone_no, " & _
                    "internet_email_address, employed_by_hcid, active_sw from opwhse.dbo.vw_pub_hanford_people union select hanford_id, " & _
                    "pay_no as hanf_pay_no, first_name, last_name, middle_name as middle_initial, NULL as name_prefix, name_suffix, " & _
                    "work_phone as phone_no, internet_address as internet_email_address, employed_by_hcid, NULL as active_sw " & _
                    "from opwhse.dbo.vw_pub_pnnl_associate union select hanford_id, pnnl_pay_no as hanf_pay_no, first_name, last_name, " & _
                    "middle_initial, name_prefix, name_suffix, contact_work_phone as phone_no, internet_email_address, company as " & _
                    "employed_by_hcid, NULL as active_sw from opwhse.dbo.vw_pub_bmi_employee) a where active_sw = 'Y' and " & _
                    "hanf_pay_no = '" & LookUpString & "'" & " and hanford_id <> ''" & _
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

            If (Me.chkFindCurrentUsers.Checked) Then
                Dim alreadyExists As Boolean = False
                'first check the current list of users and determine if that user already exists before polling the opwhse
                sqlCommonAction.CommandText = "select fu.HANFORD_ID, fu.USER_KEY, fu.FIRST_NAME, " & _
                "fu.LAST_NAME, fu.MIDDLE_NAME, fu.SUFFIX_NAME, " & _
                "fu.PRIMARY_WORK_PHONE, fu.EMAIL_ADDRESS, " & _
                "fu.LAST_NAME + ', ' + fu.FIRST_NAME + ' ' + isNull(fu.MIDDLE_NAME, '') + ' - ' + " & _
                "fu.EMAIL_ADDRESS as NameEmail, fu.COMPANY_ABBREV, " & _
                "fu.USER_CODE, fu.TRAINING_IND, fu.ACTIVE_IND from " & myUtility.getExtension() & "SAT_USER fu " & _
                "where fu.LAST_NAME like '" & LookUpString & "%'"
                sqlCommonAdapter.Fill(Me.dsCommon, "ExistingUsers")

                'if we found something return true otherwise return false
                If (Me.dsCommon.Tables("ExistingUsers").Rows.Count = 1) Then
                    Session("PersonTable") = Me.dsCommon.Tables("ExistingUsers")
                    Session("manageUserID") = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("USER_KEY")
                    alreadyExists = True
                    Return True
                ElseIf (Me.dsCommon.Tables("ExistingUsers").Rows.Count > 1) Then
                    Session("PersonTable") = Me.dsCommon.Tables("ExistingUsers")
                    Session("CurrentUsers") = True
                    Session("CurrentPage") = "./PersonSelect.aspx"
                    Session("pageSel") = "PersonSelect"
                    Response.Redirect(Session("CurrentPage"))
                Else
                    Session("PersonTable") = Nothing
                    Return False
                End If
            Else
                'set up and process the sql statement
                sqlCommonAction.CommandText = "select * from (select isnull(max(hanford_id), '') " & _
                "as hanford_id, '' as USER_KEY, isnull(max(hanf_pay_no), '') as hanf_pay_no, " & _
                "max(first_name) as first_name, max(last_name) as last_name, " & _
                "isnull(max(middle_initial), '') as MIDDLE_NAME, isnull(max(name_prefix), '') " & _
                "as name_prefix, isnull(max(name_suffix), '') as SUFFIX_NAME, " & _
                "isnull(max(phone_no), '') as PRIMARY_WORK_PHONE, isnull(max(internet_email_address), '') " & _
                "as EMAIL_ADDRESS, isnull(max(employed_by_hcid), '') as COMPANY_ABBREV, " & _
                "max(CASE active_sw WHEN Null THEN 0 WHEN 'N' THEN 0 WHEN 'Y' THEN 1 END) as " & _
                "ACTIVE_IND, 0 as USER_CODE, 0 as TRAINING_IND, max(last_name) + ', ' + " & _
                "max(first_name) + ' ' + isNull(max(middle_initial), '') + ' - ' + " & _
                "isNull(max(internet_email_address), '') as NameEmail from(select hanford_id, " & _
                "hanf_pay_no, first_name, last_name, middle_initial, NULL as name_prefix, " & _
                "name_suffix, phone_no, internet_email_address, employed_by_hcid, active_sw " & _
                "from opwhse.dbo.vw_pub_hanford_people union select hanford_id, pay_no as " & _
                "hanf_pay_no, first_name, last_name, middle_name as middle_initial, NULL as " & _
                "name_prefix, name_suffix, work_phone as phone_no, internet_address as " & _
                "internet_email_address, employed_by_hcid, NULL as active_sw from " & _
                "opwhse.dbo.vw_pub_pnnl_associate union select hanford_id, pnnl_pay_no as " & _
                "hanf_pay_no, first_name, last_name, middle_initial, name_prefix, name_suffix, " & _
                "contact_work_phone as phone_no, internet_email_address, company as " & _
                "employed_by_hcid, NULL as active_sw from opwhse.dbo.vw_pub_bmi_employee) a " & _
                "where active_sw = 'Y' and last_name like '" & LookUpString & "%' " & _
                "group by hanford_id, last_name, first_name, middle_initial, " & _
                "internet_email_address union select fu.HANFORD_ID, fu.USER_KEY, '' " & _
                "as hanf_pay_no, fu.FIRST_NAME, fu.LAST_NAME, " & _
                "fu.MIDDLE_NAME, '' as name_prefix, fu.SUFFIX_NAME, " & _
                "fu.PRIMARY_WORK_PHONE, fu.EMAIL_ADDRESS, " & _
                "fu.COMPANY_ABBREV, fu.ACTIVE_IND, fu.USER_CODE, " & _
                "fu.TRAINING_IND, fu.LAST_NAME + ', ' + fu.FIRST_NAME + ' ' + " & _
                "isNull(fu.MIDDLE_NAME, '') + ' - ' + fu.EMAIL_ADDRESS as NameEmail from " & _
                myUtility.getExtension() & "SAT_USER fu where fu.LAST_NAME like '" & LookUpString & "%' and " & _
                "HANFORD_ID = '') b order by last_name, first_name"
                sqlCommonAdapter.Fill(Me.dsCommon, "PersonData")

                'if we found something return true otherwise return false
                If (Me.dsCommon.Tables("PersonData").Rows.Count = 1) Then
                    Session("PersonTable") = Me.dsCommon.Tables("PersonData")
                    Session("SelectedPersonRow") = 0
                    Return True
                ElseIf (Me.dsCommon.Tables("PersonData").Rows.Count > 1) Then
                    Session("PersonTable") = Me.dsCommon.Tables("PersonData")
                    Session("CurrentUsers") = False
                    Session("CurrentPage") = "./PersonSelect.aspx"
                    Session("pageSel") = "PersonSelect"
                    Response.Redirect(Session("CurrentPage"))
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

    'gets the hanford people information for the user we're looking for, returns true if it finds the person(s)
    Private Function listGetHanfordPeopleInfo(ByVal LookUpString As String) As Boolean
        Try
            'initialize the adhoc data collection items
            If ((LookUpString.ToUpper.StartsWith("H") And IsNumeric(LookUpString.Substring(1))) Or (LookUpString.Length = 7 And IsNumeric(LookUpString))) Then
                'Process the apparent HID
                Return (listProcessHID(LookUpString))
            ElseIf ((LookUpString.ToUpper.StartsWith("D") And LookUpString.Length = 6) Or LookUpString.Length = 5) Then
                If (IsNumeric(LookUpString.Substring(LookUpString.Length - 3))) Then
                    'Process the apparent payroll
                    Return (listProcessPayroll(LookUpString))
                Else
                    'Process the apparent last name
                    Return (listProcessEmail(LookUpString))
                End If
            Else
                'Process the apparent email
                Return (listProcessEmail(LookUpString))
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'Process the HID
    Private Function listProcessHID(ByVal LookUpString As String) As Boolean
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
            "opwhse.dbo.vw_pub_hanford_people hp where fu.HANFORD_ID = hp.hanford_id " & _
            "and fu.HANFORD_ID = '" & LookUpString & "'"
            sqlCommonAdapter.Fill(Me.dsCommon, "ExistingUsers")

            'if we found something then they already exist
            If (Me.dsCommon.Tables("ExistingUsers").Rows.Count >= 1) Then
                Session("PersonTable") = Me.dsCommon.Tables("ExistingUsers")
                alreadyExists = True
            End If

            If (alreadyExists) Then
                'if we found something then add their user id to the list of users.
                Return True
            Else
                'set up and process the sql statement
                sqlCommonAction.CommandText = "select isnull(max(hanford_id), '') as hanford_id, isnull(max(hanf_pay_no), '') as " & _
                "hanf_pay_no, max(first_name) as first_name, max(last_name) as last_name, isnull(max(middle_initial), '') as " & _
                "MIDDLE_NAME, isnull(max(name_prefix), '') as name_prefix, isnull(max(name_suffix), '') as SUFFIX_NAME, " & _
                "isnull(max(phone_no), '') as PRIMARY_WORK_PHONE, isnull(max(internet_email_address), '') as EMAIL_ADDRESS, " & _
                "isnull(max(employed_by_hcid), '') as COMPANY_ABBREV, isNull(max(active_sw), 'N') as ACTIVE_IND from  (" & _
                "select hanford_id, hanf_pay_no, first_name, last_name, middle_initial, NULL as name_prefix, name_suffix, phone_no, " & _
                "internet_email_address, employed_by_hcid, active_sw from opwhse.dbo.vw_pub_hanford_people union select hanford_id, " & _
                "pay_no as hanf_pay_no, first_name, last_name, middle_name as middle_initial, NULL as name_prefix, name_suffix, " & _
                "work_phone as phone_no, internet_address as internet_email_address, employed_by_hcid, NULL as active_sw " & _
                "from opwhse.dbo.vw_pub_pnnl_associate union select hanford_id, pnnl_pay_no as hanf_pay_no, first_name, last_name, " & _
                "middle_initial, name_prefix, name_suffix, contact_work_phone as phone_no, internet_email_address, company as " & _
                "employed_by_hcid, NULL as active_sw from opwhse.dbo.vw_pub_bmi_employee) a where " & _
                "hanford_id = '" & LookUpString & "' and active_sw = 'Y'" & _
                " group by hanford_id, last_name, first_name, middle_initial, internet_email_address"
                sqlCommonAdapter.Fill(Me.dsCommon, "PersonData")

                'if we found something return true otherwise return false
                If (Me.dsCommon.Tables("PersonData").Rows.Count >= 1) Then
                    Session("PersonTable") = Me.dsCommon.Tables("PersonData")
                    'create the new level 0 user
                    If (createUser(sqlCommonAction, sqlCommonAdapter)) Then
                        Return True
                    Else
                        Return False
                    End If
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

    'Process the Payroll
    Private Function listProcessPayroll(ByVal LookUpString As String) As Boolean
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
            "opwhse.dbo.vw_pub_hanford_people hp where fu.LAST_NAME = hp.Last_Name and fu.FIRST_NAME = hp.first_name " & _
            "and hp.hanf_pay_no = '" & LookUpString & "'"
            sqlCommonAdapter.Fill(Me.dsCommon, "ExistingUsers")

            'if we found something return true otherwise return false
            If (Me.dsCommon.Tables("ExistingUsers").Rows.Count >= 1) Then
                Session("PersonTable") = Me.dsCommon.Tables("ExistingUsers")
                alreadyExists = True
            End If

            If (alreadyExists) Then
                'if we found something then add their user id to the list of users.
                Return True
            Else
                'set up and process the sql statement
                sqlCommonAction.CommandText = "select isnull(max(hanford_id), '') as hanford_id, isnull(max(hanf_pay_no), '') as " & _
                "hanf_pay_no, max(first_name) as first_name, max(last_name) as last_name, isnull(max(middle_initial), '') as " & _
                "MIDDLE_NAME, isnull(max(name_prefix), '') as name_prefix, isnull(max(name_suffix), '') as SUFFIX_NAME, " & _
                "isnull(max(phone_no), '') as PRIMARY_WORK_PHONE, isnull(max(internet_email_address), '') as EMAIL_ADDRESS, " & _
                "isnull(max(employed_by_hcid), '') as COMPANY_ABBREV, isNull(max(active_sw), 'N') as ACTIVE_IND from  (" & _
                "select hanford_id, hanf_pay_no, first_name, last_name, middle_initial, NULL as name_prefix, name_suffix, phone_no, " & _
                "internet_email_address, employed_by_hcid, active_sw from opwhse.dbo.vw_pub_hanford_people union select hanford_id, " & _
                "pay_no as hanf_pay_no, first_name, last_name, middle_name as middle_initial, NULL as name_prefix, name_suffix, " & _
                "work_phone as phone_no, internet_address as internet_email_address, employed_by_hcid, NULL as active_sw " & _
                "from opwhse.dbo.vw_pub_pnnl_associate union select hanford_id, pnnl_pay_no as hanf_pay_no, first_name, last_name, " & _
                "middle_initial, name_prefix, name_suffix, contact_work_phone as phone_no, internet_email_address, company as " & _
                "employed_by_hcid, NULL as active_sw from opwhse.dbo.vw_pub_bmi_employee) a where " & _
                "hanf_pay_no = '" & LookUpString & "' and active_sw = 'Y'" & _
                " group by hanford_id, last_name, first_name, middle_initial, internet_email_address"
                sqlCommonAdapter.Fill(Me.dsCommon, "PersonData")

                'if we found something return true otherwise return false
                If (Me.dsCommon.Tables("PersonData").Rows.Count >= 1) Then
                    Session("PersonTable") = Me.dsCommon.Tables("PersonData")
                    'create the new level 0 user
                    If (createUser(sqlCommonAction, sqlCommonAdapter)) Then
                        Return True
                    Else
                        Return False
                    End If
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

    'Process the e-mail address
    Private Function listProcessEmail(ByVal LookUpString As String) As Boolean
        Try
            'destroy any data that may exist before using the table
            If (Me.dsCommon.Tables.Contains("ExisitingUsers")) Then
                Me.dsCommon.Tables("ExistingUsers").Rows.Clear()
            End If
            If (Me.dsCommon.Tables.Contains("PersonData")) Then
                Me.dsCommon.Tables("PersonData").Rows.Clear()
            End If

            'determine if the e-mail address looks valid
            If Not (isValidEmail(LookUpString)) Then
                Return False
            End If

            Dim alreadyExists As Boolean = False
            'first check the current list of users and determine if that user already exists before polling the opwhse
            sqlCommonAction.CommandText = "select fu.HANFORD_ID, fu.USER_KEY, fu.FIRST_NAME, " & _
            "fu.LAST_NAME, fu.MIDDLE_NAME, fu.SUFFIX_NAME, " & _
            "fu.PRIMARY_WORK_PHONE, fu.EMAIL_ADDRESS, fu.COMPANY_ABBREV, " & _
            "fu.USER_CODE, fu.TRAINING_IND, fu.ACTIVE_IND from " & myUtility.getExtension() & "SAT_USER fu " & _
            " where fu.EMAIL_ADDRESS = '" & LookUpString & "'"
            sqlCommonAdapter.Fill(Me.dsCommon, "ExistingUsers")

            'if we found something then they already exist
            If (Me.dsCommon.Tables("ExistingUsers").Rows.Count >= 1) Then
                Session("PersonTable") = Me.dsCommon.Tables("ExistingUsers")
                alreadyExists = True
            End If

            If (alreadyExists) Then
                'if we found something then add their user id to the list of users.
                Return True
            Else
                'set up and process the sql statement
                sqlCommonAction.CommandText = "select isnull(max(hanford_id), '') as hanford_id, isnull(max(hanf_pay_no), '') as " & _
                "hanf_pay_no, max(first_name) as first_name, max(last_name) as last_name, isnull(max(middle_initial), '') as " & _
                "MIDDLE_NAME, isnull(max(name_prefix), '') as name_prefix, isnull(max(name_suffix), '') as SUFFIX_NAME, " & _
                "isnull(max(phone_no), '') as PRIMARY_WORK_PHONE, isnull(max(internet_email_address), '') as EMAIL_ADDRESS, " & _
                "isnull(max(employed_by_hcid), '') as COMPANY_ABBREV, isNull(max(active_sw), 'N') as ACTIVE_IND, " & _
                "max(last_name) + ', ' + max(first_name) + ' ' + isNull(max(middle_initial), '') + ' - ' + " & _
                "isNull(max(internet_email_address), '') as NameEmail from  (" & _
                "select hanford_id, hanf_pay_no, first_name, last_name, middle_initial, NULL as name_prefix, name_suffix, phone_no, " & _
                "internet_email_address, employed_by_hcid, active_sw from opwhse.dbo.vw_pub_hanford_people union select hanford_id, " & _
                "pay_no as hanf_pay_no, first_name, last_name, middle_name as middle_initial, NULL as name_prefix, name_suffix, " & _
                "work_phone as phone_no, internet_address as internet_email_address, employed_by_hcid, NULL as active_sw " & _
                "from opwhse.dbo.vw_pub_pnnl_associate union select hanford_id, pnnl_pay_no as hanf_pay_no, first_name, last_name, " & _
                "middle_initial, name_prefix, name_suffix, contact_work_phone as phone_no, internet_email_address, company as " & _
                "employed_by_hcid, NULL as active_sw from opwhse.dbo.vw_pub_bmi_employee) a where " & _
                "internet_email_address = '" & LookUpString & "'" & _
                " group by hanford_id, last_name, first_name, middle_initial, internet_email_address " & _
                "order by last_name, first_name"
                sqlCommonAdapter.Fill(Me.dsCommon, "PersonData")

                'if we found something return true otherwise ask the user to provide a name for the user's e-mail
                If (Me.dsCommon.Tables("PersonData").Rows.Count >= 1) Then
                    Session("PersonTable") = Me.dsCommon.Tables("PersonData")
                    'create the new level 0 user
                    If (createUser(sqlCommonAction, sqlCommonAdapter)) Then
                        Return True
                    Else
                        Return False
                    End If
                ElseIf (Session("ProcessEmail") = "True") Then
                    Session("PersonTable") = Nothing
                    If (createUser(sqlCommonAction, sqlCommonAdapter)) Then
                        Return True
                    Else
                        Return False
                    End If
                Else
                    Session("PersonTable") = Nothing
                    Session("JSAction") &= "<script>newWin = window.open('./PersonInput.aspx?PersonInputEmail=" & LookUpString & "&PersonInputUserType=" & Me.ddlUserType.SelectedValue & "', '_blank', 'width=600,height=400,scrollbars=yes,toolbar=no,titlebar=no,status=no,resizable=yes,menubar=no');newWin.focus();</script>"
                End If
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'checks to be sure the e-mail address looks valid
    Private Function isValidEmail(ByVal strEmail As String) As Boolean
        Try
            'make sure that the last period is 3 away from the end
            If ((strEmail.Length - strEmail.LastIndexOf(".")) = 4) Then
                'make sure that the distance between the @ and the last period is greater than 2
                If ((strEmail.LastIndexOf(".") - strEmail.LastIndexOf("@")) > 2) Then
                    'make sure there is only one @ in the string
                    Dim ch As Char
                    Dim atCounter As Integer = 0
                    For Each ch In strEmail
                        If (ch = "@") Then
                            atCounter += 1
                        End If
                    Next
                    If (atCounter = 1) Then
                        'make sure there is only one period after the @ in the string
                        Dim strTemp As String = strEmail.Substring(strEmail.LastIndexOf("@") + 1)
                        Dim periodCounter As Integer = 0
                        For Each ch In strTemp
                            If (ch = ".") Then
                                periodCounter += 1
                            End If
                        Next
                        If (periodCounter = 1) Then
                            Return True
                        Else
                            Return False
                        End If
                    Else
                        Return False
                    End If
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'attempts to insert a user, returns false if it cannot
    Private Function createUser(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, Optional ByVal strEmail As String = "") As Boolean
        Try
            sqlCommonAction.Parameters.Clear()
            Dim accessLevel As Integer = Me.ddlUserType.SelectedIndex
            If Not (Session("PersonTable") Is Nothing) Then
                'parameterized the text input to allow for things like quotes to get through
                Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@EMAIL_ADDRESS", System.Data.SqlDbType.VarChar, 128, "EMAIL_ADDRESS")
                param0.Value = Me.dsCommon.Tables("PersonData").Rows(0).Item("internet_email_address")
                sqlCommonAction.Parameters.Add(param0)
                Dim param1 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@FIRST_NAME", System.Data.SqlDbType.VarChar, 50, "FIRST_NAME")
                param1.Value = Me.dsCommon.Tables("PersonData").Rows(0).Item("first_name")
                sqlCommonAction.Parameters.Add(param1)
                Dim param2 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@LAST_NAME", System.Data.SqlDbType.VarChar, 50, "LAST_NAME")
                param2.Value = Me.dsCommon.Tables("PersonData").Rows(0).Item("last_name")
                sqlCommonAction.Parameters.Add(param2)
                Dim param3 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@MIDDLE_NAME", System.Data.SqlDbType.VarChar, 50, "MIDDLE_NAME")
                param3.Value = Me.dsCommon.Tables("PersonData").Rows(0).Item("MIDDLE_NAME")
                sqlCommonAction.Parameters.Add(param3)
                Dim param4 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@SUFFIX_NAME", System.Data.SqlDbType.VarChar, 50, "SUFFIX_NAME")
                param4.Value = Me.dsCommon.Tables("PersonData").Rows(0).Item("SUFFIX_NAME")
                sqlCommonAction.Parameters.Add(param4)
                Dim param5 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@PRIMARY_WORK_PHONE", System.Data.SqlDbType.VarChar, 32, "PRIMARY_WORK_PHONE")
                param5.Value = Me.dsCommon.Tables("PersonData").Rows(0).Item("PRIMARY_WORK_PHONE")
                sqlCommonAction.Parameters.Add(param5)
                Dim param6 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@COMPANY_ABBREV", System.Data.SqlDbType.VarChar, 32, "COMPANY_ABBREV")
                param6.Value = Me.dsCommon.Tables("PersonData").Rows(0).Item("COMPANY_ABBREV")
                sqlCommonAction.Parameters.Add(param6)
                Dim param7 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@HANFORD_ID", System.Data.SqlDbType.VarChar, 12, "HANFORD_ID")

                'get the user's hid from the stored table
                Dim dt As DataTable = CType(Session("PersonTable"), DataTable)
                param7.Value = dt.Rows(0).Item("hanford_id")
                sqlCommonAction.Parameters.Add(param7)

                'generate a password for this user
                Dim param8 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@AUTHENTICATION_CODE", System.Data.SqlDbType.VarChar, 64, "AUTHENTICATION_CODE")
                If (param7.Value = "") Then
                    param8.Value = generatePassword(param0.value)
                Else
                    param8.Value = generatePassword(param7.Value)
                End If
                sqlCommonAction.Parameters.Add(param8)

                'determine if the user's level was appropriately selected.  if not bring it up to level 0
                If (Me.ddlUserType.SelectedItem.Value < 0) Then
                    Me.ddlUserType.SelectedIndex += 1
                    accessLevel = Me.ddlUserType.SelectedIndex
                End If
                accessLevel = Me.ddlUserType.SelectedIndex
                If (Session("ProcessEmail") = "True") Then
                    Me.ddlUserType.SelectedIndex = 1
                End If

                'insert the new user
                sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_USER (HANFORD_ID, EMAIL_ADDRESS, USER_CODE, LAST_NAME, FIRST_NAME, " & _
                "MIDDLE_NAME, SUFFIX_NAME, AUTHENTICATION_CODE, LAST_CHANGE_AUTH_CD_DATE, PRIMARY_WORK_PHONE, CREATE_DATE, USED_DATE, CREATOR_USER_KEY, ACTIVE_IND, " & _
                "TRAINING_IND, COMPANY_ABBREV, NO_EMAIL_IND) VALUES (@HANFORD_ID, @EMAIL_ADDRESS, " & Me.ddlUserType.SelectedValue & _
                ", @LAST_NAME, @FIRST_NAME, @MIDDLE_NAME, @SUFFIX_NAME, @AUTHENTICATION_CODE, '" & Now.AddMonths(-7) & "', @PRIMARY_WORK_PHONE,'" & Now & "','1/1/1970'," & _
                Session("UserID") & "," & _
                IIf(Me.dsCommon.Tables("PersonData").Rows(0).Item("active_sw") = "Y", 1, 0) & ", 0" & _
                ", @COMPANY_ABBREV, 0)"
                sqlCommonAction.ExecuteNonQuery()

                'get the user's sequence number
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER where MIDDLE_NAME = @MIDDLE_NAME and LAST_NAME = " & _
                "@LAST_NAME and FIRST_NAME = @FIRST_NAME and SUFFIX_NAME = @SUFFIX_NAME"
                sqlCommonAdapter.Fill(Me.dsCommon, "UserIDFetch")
                Session("mailListID") &= Me.dsCommon.Tables("UserIDFetch").Rows(0).Item("EMAIL_ADDRESS") & "," & strNewPassword & ","
                Return True
            Else
                'parameterized the text input to allow for things like quotes to get through
                Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@EMAIL_ADDRESS", System.Data.SqlDbType.VarChar, 128, "EMAIL_ADDRESS")
                param0.Value = Session("PersonInputEmailReturn")
                sqlCommonAction.Parameters.Add(param0)
                Dim param1 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@FIRST_NAME", System.Data.SqlDbType.VarChar, 50, "FIRST_NAME")
                param1.Value = Session("PersonInputFirstName")
                sqlCommonAction.Parameters.Add(param1)
                Dim param2 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@LAST_NAME", System.Data.SqlDbType.VarChar, 50, "LAST_NAME")
                param2.Value = Session("PersonInputLastName")
                sqlCommonAction.Parameters.Add(param2)
                Dim param3 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@MIDDLE_NAME", System.Data.SqlDbType.VarChar, 50, "MIDDLE_NAME")
                param3.Value = Session("PersonInputMiddleInitial")
                sqlCommonAction.Parameters.Add(param3)

                'determine if the user's level was appropriately selected.  if not bring it up to level 0
                If (Me.ddlUserType.SelectedItem.Value < 0) Then
                    Me.ddlUserType.SelectedIndex += 1
                    accessLevel = Me.ddlUserType.SelectedIndex
                End If
                accessLevel = Me.ddlUserType.SelectedIndex
                If (Session("ProcessEmail") = "True") Then
                    Me.ddlUserType.SelectedIndex = 1
                End If

                'insert the new user
                sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_USER (HANFORD_ID, EMAIL_ADDRESS, USER_CODE, LAST_NAME, FIRST_NAME, " & _
                "MIDDLE_NAME, SUFFIX_NAME, AUTHENTICATION_CODE, LAST_CHANGE_AUTH_CD_DATE, PRIMARY_WORK_PHONE, CREATE_DATE, USED_DATE, CREATOR_USER_KEY, ACTIVE_IND, " & _
                "TRAINING_IND, COMPANY_ABBREV, NO_EMAIL_IND) VALUES ('', @EMAIL_ADDRESS, " & Me.ddlUserType.SelectedValue & ", " & _
                "@LAST_NAME, @FIRST_NAME, @MIDDLE_NAME, '', '" & _
                generatePassword(param0.Value) & _
                "','" & Now.AddMonths(-7) & "', '','" & Now & "','1/1/1970'," & Session("UserID") & "," & _
                " 1, 0, '', 0)"
                sqlCommonAction.ExecuteNonQuery()

                'get the user's sequence number
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER where EMAIL_ADDRESS = @EMAIL_ADDRESS"
                sqlCommonAdapter.Fill(Me.dsCommon, "UserIDFetch")
                Session("mailListID") &= Me.dsCommon.Tables("UserIDFetch").Rows(0).Item("EMAIL_ADDRESS") & "," & strNewPassword & ","
                Return True
            End If

            'reset the access level
            Me.ddlUserType.SelectedIndex = accessLevel
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            sqlCommonAction.Parameters.Clear()
            Return False
        End Try
    End Function
#End Region

#Region "Page Action"
    'loads the data for the selected user
    Private Sub ddlUser_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddlUser.SelectedIndexChanged
        Session("JSAction") = ""

        Try
            'refresh the tables
            Dim oldID = Session("manageUserID")
            Session("manageUserID") = -1
            Dim userID As Integer = Me.ddlUser.SelectedValue
            If (loadData()) Then
                'find the user information we'll need to load
                Session("UserRow") = findRow("Users", userID)

                'If the user's information still exists then use it otherwise inform the user that the user's information is not retrievable
                If (Session("UserRow") <> -1) Then
                    'if we're not working with a new record create one and populate it.  Otherwise re-populate the page.
                    If (Session("manageUserID") = -1) Then
                        'populate the controls with the selected individual's information
                        Me.txtEmail.Text = Me.dsCommon.Tables("Users").Rows(Session("UserRow")).Item("EMAIL_ADDRESS")
                        Me.txtFirstName.Text = Me.dsCommon.Tables("Users").Rows(Session("UserRow")).Item("FIRST_NAME")
                        Me.txtLastName.Text = Me.dsCommon.Tables("Users").Rows(Session("UserRow")).Item("LAST_NAME")
                        Me.txtMiddleName.Text = Me.dsCommon.Tables("Users").Rows(Session("UserRow")).Item("MIDDLE_NAME")
                        Me.txtSuffixName.Text = Me.dsCommon.Tables("Users").Rows(Session("UserRow")).Item("SUFFIX_NAME")
                        Me.txtPhoneNumber.Text = Me.dsCommon.Tables("Users").Rows(Session("UserRow")).Item("PRIMARY_WORK_PHONE")
                        Me.txtCompany.Text = Me.dsCommon.Tables("Users").Rows(Session("UserRow")).Item("COMPANY_ABBREV")
                        Me.ddlUserType.SelectedIndex = Me.dsCommon.Tables("Users").Rows(Session("UserRow")).Item("USER_CODE") + 1
                        Me.chkTQSurvey.Checked = Me.dsCommon.Tables("Users").Rows(Session("UserRow")).Item("TRAINING_IND")
                        Me.chkActive.Checked = Me.dsCommon.Tables("Users").Rows(Session("UserRow")).Item("ACTIVE_IND")
                        Me.chkPasswordReset.Checked = 0
                    End If

                    Session("manageUserID") = userID

                    'Check to see if this is our first time on the page or not
                    If (Session("manageUserID") <> -1) Then
                        Me.ddlUser.Items.FindByValue(Session("manageUserID")).Selected = True
                    Else
                        Me.ddlUser.SelectedIndex = 0
                    End If

                    'set the page options
                    setPageOptions()

                    'Populate the controls on the page
                    setControls()
                ElseIf (userID = -1) Then
                    Session("manageUserID") = userID

                    'set the page options
                    setPageOptions()

                    'Populate the controls on the page
                    setControls()
                Else
                    alert("Error, this user may have been deleted within the last few seconds.  User information is irretrievable for the selected individual.")
                    Session("manageUserID") = oldID
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
                Dim oldID = Session("manageUserID")
                Session("manageUserID") = -1
                If (loadData()) Then
                    'process the input only if the user actually put something in the find text field
                    If (Me.txtHIDPayrollName.Text <> "") Then
                        'find the user information we'll need to load
                        If (getHanfordPeopleInfo(Me.txtHIDPayrollName.Text)) Then
                            'if we're working with a new record create one and populate it.  Otherwise re-populate the page.
                            If (Session("manageUserID") = -1) Then
                                'if the user already exists load their data
                                If (Me.dsCommon.Tables.Contains("ExistingUsers")) Then
                                    If (Me.dsCommon.Tables("ExistingUsers").Rows.Count() > 0) Then
                                        'populate the controls with the selected individual's information
                                        Me.txtEmail.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("EMAIL_ADDRESS")
                                        Me.txtFirstName.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("first_name")
                                        Me.txtLastName.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("last_name")
                                        Me.txtMiddleName.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("MIDDLE_NAME")
                                        Me.txtSuffixName.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("SUFFIX_NAME")
                                        Me.txtPhoneNumber.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("PRIMARY_WORK_PHONE")
                                        Me.txtCompany.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("COMPANY_ABBREV")
                                        Me.ddlUserType.SelectedIndex = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("USER_CODE") + 1
                                        Me.chkTQSurvey.Checked = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("TRAINING_IND")
                                        Me.chkActive.Checked = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("ACTIVE_IND")
                                        Me.chkPasswordReset.Checked = 0
                                        Session("UserRow") = findRow("Users", Session("manageUserID"))
                                        Me.ddlUser.SelectedIndex = Session("UserRow") + 1
                                    Else
                                        'populate the controls with the selected individual's information
                                        Me.txtEmail.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("EMAIL_ADDRESS")
                                        Me.txtFirstName.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("first_name")
                                        Me.txtLastName.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("last_name")
                                        Me.txtMiddleName.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("MIDDLE_NAME")
                                        Me.txtSuffixName.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("SUFFIX_NAME")
                                        Me.txtPhoneNumber.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("PRIMARY_WORK_PHONE")
                                        Me.txtCompany.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("COMPANY_ABBREV")
                                        Me.ddlUserType.SelectedIndex = 0
                                        Me.chkTQSurvey.Checked = 0
                                        Me.chkActive.Checked = 1
                                        Me.chkPasswordReset.Checked = 0
                                    End If
                                Else
                                    'restore the person data and find the exact user selected
                                    Dim dt As DataTable = Session("PersonTable")
                                    Dim dr As DataRow = dt.Rows(Session("selectedPersonRow"))

                                    'try to get the user's id.
                                    'determine if the user already exists
                                    Dim isExistingUser As Boolean = False
                                    isExistingUser = checkUserExistence(dr)

                                    If (isExistingUser) Then
                                        Session("manageUserID") = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("USER_KEY")
                                        Me.txtEmail.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("EMAIL_ADDRESS")
                                        Me.txtFirstName.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("first_name")
                                        Me.txtLastName.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("last_name")
                                        Me.txtMiddleName.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("MIDDLE_NAME")
                                        Me.txtSuffixName.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("SUFFIX_NAME")
                                        Me.txtPhoneNumber.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("PRIMARY_WORK_PHONE")
                                        Me.txtCompany.Text = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("COMPANY_ABBREV")
                                        Me.ddlUserType.SelectedIndex = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("USER_CODE") + 1
                                        Me.chkTQSurvey.Checked = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("TRAINING_IND")
                                        Me.chkActive.Checked = Me.dsCommon.Tables("ExistingUsers").Rows(0).Item("ACTIVE_IND")
                                        Me.chkPasswordReset.Checked = 0
                                        Session("UserRow") = findRow("Users", Session("manageUserID"))
                                        Me.ddlUser.SelectedIndex = Session("UserRow") + 1
                                    Else
                                        'populate the controls with the selected individual's information
                                        Me.txtEmail.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("EMAIL_ADDRESS")
                                        Me.txtFirstName.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("first_name")
                                        Me.txtLastName.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("last_name")
                                        Me.txtMiddleName.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("MIDDLE_NAME")
                                        Me.txtSuffixName.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("SUFFIX_NAME")
                                        Me.txtPhoneNumber.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("PRIMARY_WORK_PHONE")
                                        Me.txtCompany.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("COMPANY_ABBREV")
                                        Me.ddlUserType.SelectedIndex = 0
                                        Me.chkTQSurvey.Checked = 0
                                        Me.chkActive.Checked = 1
                                        Me.chkPasswordReset.Checked = 0
                                    End If
                                End If
                            Else
                                Session("UserRow") = findRow("Users", Session("manageUserID"))
                                Me.ddlUser.SelectedIndex = Session("UserRow") + 1
                            End If

                            'Populate the controls on the page
                            setControls("override")

                            'set the page options
                            setPageOptions()
                            Me.txtHIDPayrollName.Text = ""
                        Else
                            alert("This person was not found.")
                            Session("manageUserID") = oldID
                            'Populate the controls on the page
                            setControls("override")

                            'set the page options
                            setPageOptions()
                            Me.txtHIDPayrollName.Text = ""
                        End If
                    Else
                        alert("You must input some text into the HID, Payroll, or Last Name field to use the find functionality.")
                        Session("manageUserID") = oldID
                        'Populate the controls on the page
                        setControls("override")

                        'set the page options
                        setPageOptions()
                        Me.txtHIDPayrollName.Text = ""
                    End If
                Else
                    Me.txtHIDPayrollName.Text = ""
                    alert("Error loading user information.")
                End If
            Else
                Me.txtHIDPayrollName.Text = ""
                alert("You may have entered malicious code into the HID/Emplid/Last Name text field.  Find aborted.")
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Me.txtHIDPayrollName.Text = ""
            alert("Error loading user information.")
        End Try
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnInsert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsert.Click
        Session("JSAction") = ""
        Dim failure As Boolean
        Dim blnContinue As Boolean = True
        
        If Not (loadData(True)) Then
            blnContinue = False
            alert("Error loading user information.")
        End If

        'check for possible malicious code input
        If (Session("ImportList") <> True) Then
            If Not (myUtility.goodInput(Me.txtCompany) And myUtility.goodInput(Me.txtEmail) _
                And myUtility.goodInput(Me.txtFirstName) And myUtility.goodInput(Me.txtLastName) _
                And myUtility.goodInput(Me.txtMiddleName) And myUtility.goodInput(Me.txtPhoneNumber) _
                And myUtility.goodInput(Me.txtSuffixName)) Then
                blnContinue = False
                alert("Possible malicious code entry(s).  Insert aborted.")
            End If
        Else
            If Not (myUtility.goodInput(Me.txtUserList)) Then
                blnContinue = False
                alert("Possible malicious code entry(s).  Insert aborted.")
            End If
        End If

        If (blnContinue) Then
            'do an insert/update/delete/reset based on the page option selected
            failure = doInsert()
            Me.chkPasswordReset.Checked = False

            'determine if we need to reload the data
            If Not (loadData(True)) Then
                blnContinue = False
                alert("Error loading delegate/owner information.")
            End If

            If (blnContinue) Then
                'reset the page controls
                If Not (failure And Me.chkImportList.Checked) And Not (Me.chkImportList.Checked) Then
                    Session("UserRow") = findRow("Users", Session("manageUserID"))
                    setControls()
                Else
                    setControls("override")
                End If

                Session("TransExists") = False
                'set the page options
                setPageOptions()
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Session("JSAction") = ""
        Dim failure As Boolean
        Dim blnContinue As Boolean = True
        
        If Not (loadData(True)) Then
            blnContinue = False
            alert("Error loading user information.")
        End If

        'check for possible malicious code input
        If (Session("ImportList") <> True) Then
            If Not (myUtility.goodInput(Me.txtCompany) And myUtility.goodInput(Me.txtEmail) _
                And myUtility.goodInput(Me.txtFirstName) And myUtility.goodInput(Me.txtLastName) _
                And myUtility.goodInput(Me.txtMiddleName) And myUtility.goodInput(Me.txtPhoneNumber) _
                And myUtility.goodInput(Me.txtSuffixName)) Then
                blnContinue = False
                alert("Possible malicious code entry(s).  Update aborted.")
            End If
        Else
            If Not (myUtility.goodInput(Me.txtUserList)) Then
                blnContinue = False
                alert("Possible malicious code entry(s).  Update aborted.")
            End If
        End If

        If (blnContinue) Then
            'do the update
            failure = doUpdate()
            Me.chkPasswordReset.Checked = False
        
            'determine if we need to reload the data
            If Not (loadData(True)) Then
                blnContinue = False
                alert("Error loading delegate/owner information.")
            End If

            If (blnContinue) Then
                'reset the page controls
                If Not (failure And Me.chkImportList.Checked) And Not (Me.chkImportList.Checked) Then
                    Session("UserRow") = findRow("Users", Session("manageUserID"))
                    setControls()
                Else
                    setControls("override")
                End If

                Session("TransExists") = False
                'set the page options
                setPageOptions()
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Session("JSAction") = ""
        Dim failure As Boolean
        Dim blnContinue As Boolean = True
        
        If Not (loadData(True)) Then
            blnContinue = False
            alert("Error loading user information.")
        End If

        'check for possible malicious code input
        If (Session("ImportList") <> True) Then
            If Not (myUtility.goodInput(Me.txtCompany) And myUtility.goodInput(Me.txtEmail) _
                And myUtility.goodInput(Me.txtFirstName) And myUtility.goodInput(Me.txtLastName) _
                And myUtility.goodInput(Me.txtMiddleName) And myUtility.goodInput(Me.txtPhoneNumber) _
                And myUtility.goodInput(Me.txtSuffixName)) Then
                blnContinue = False
                alert("Possible malicious code entry(s).  Delete aborted.")
            End If
        Else
            If Not (myUtility.goodInput(Me.txtUserList)) Then
                blnContinue = False
                alert("Possible malicious code entry(s).  Delete aborted.")
            End If
        End If

        If (blnContinue) Then
            'do the delete
            failure = doDelete()
            Me.chkPasswordReset.Checked = False

            'determine if we need to reload the data
            If Not (loadData(True)) Then
                blnContinue = False
                alert("Error loading delegate/owner information.")
            End If

            If (blnContinue) Then
                'reset the page controls
                If Not (failure And Me.chkImportList.Checked) And Not (Me.chkImportList.Checked) Then
                    Session("UserRow") = findRow("Users", Session("manageUserID"))
                    setControls()
                Else
                    setControls("override")
                End If

                Session("TransExists") = False
                'set the page options
                setPageOptions()
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Session("JSAction") = ""
        Dim failure As Boolean
        Dim blnContinue As Boolean = True
        
        If Not (loadData(True)) Then
            blnContinue = False
            alert("Error loading user information.")
        End If

        'check for possible malicious code input
        If (Session("ImportList") <> True) Then
            If Not (myUtility.goodInput(Me.txtCompany) And myUtility.goodInput(Me.txtEmail) _
                And myUtility.goodInput(Me.txtFirstName) And myUtility.goodInput(Me.txtLastName) _
                And myUtility.goodInput(Me.txtMiddleName) And myUtility.goodInput(Me.txtPhoneNumber) _
                And myUtility.goodInput(Me.txtSuffixName)) Then
                blnContinue = False
                alert("Possible malicious code entry(s). Reset aborted.")
            End If
        Else
            If Not (myUtility.goodInput(Me.txtUserList)) Then
                blnContinue = False
                alert("Possible malicious code entry(s).  Reset aborted.")
            End If
        End If

        If (blnContinue) Then
            'do the reset
            failure = doReset()

            If (blnContinue) Then
                'reset the page controls
                If Not (failure And Me.chkImportList.Checked) And Not (Me.chkImportList.Checked) Then
                    Session("UserRow") = findRow("Users", Session("manageUserID"))
                    setControls()
                Else
                    setControls("override")
                End If

                Session("TransExists") = False
                'set the page options
                setPageOptions()
            End If
        End If
    End Sub

    'clears the page creating an opportunity to create a new user
    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Session("JSAction") = ""
        Session("manageUserID") = -1
        Me.ddlUser.SelectedIndex = 0
        Me.txtHIDPayrollName.Text = ""
        Me.txtUserList.Text = ""

        'set the page options
        setPageOptions()

        'Populate the controls on the page
        setControls()
    End Sub
#End Region

#Region "Insert Code"
    'drives the commit and roll-back operations of the insert
    Private Function doInsert() As Boolean
        Try
            'start transaction
            If (myUtility.setupTransaction(sqlCommonAction, Me.commonConn)) Then
                If (Session("ImportList") = True) Then
                    If (Me.ddlUserType.SelectedIndex <> 0 And Trim("" & Me.txtUserList.Text) <> "") Then
                        sqlCommonAdapter.SelectCommand = sqlCommonAction

                        Dim failure As Boolean = False
                        'process the text box for HID's, Emails, and Payrolls
                        If (processManualEntry(sqlCommonAction, sqlCommonAdapter)) Then
                            sqlCommonAction.Transaction.Commit()
                            Session("transExists") = False
                            If (Me.ddlUserType.SelectedIndex > 1) Then
                                'send a mass e-mail to all affected individuals
                                'create an array of e-mail addresses while trimming the white space
                                Dim requestItems As Array
                                Dim index As Integer = 0
                                If (Trim("" & Session("mailListID")) <> "") Then
                                    Dim strLine As String = ""
                                    'replace white space text
                                    requestItems = Session("mailListID").Split(",")
                                    While index < requestItems.Length
                                        requestItems(index) = Trim(CType(requestItems(index), String))
                                        index += 1
                                    End While
                                End If

                                'determine if we actually got anything.  If not then skip this function
                                If (requestItems Is Nothing) Then
                                    Return True
                                Else
                                    Session("mailListID") = ""
                                End If

                                'determine if the person exists as a current user.
                                index = 0
                                While index < (requestItems.Length - 1)
                                    If (userInfo(requestItems(index))) Then
                                        strNewPassword = requestItems(index + 1)
                                        If (Me.dsCommon.Tables("SpecificUserInfo").Rows(0).Item("HANFORD_ID") <> "") Then
                                            If Not (sendMail(Me.dsCommon.Tables("SpecificUserInfo").Rows(0).Item("EMAIL_ADDRESS"), Me.dsCommon.Tables("SpecificUserInfo").Rows(0).Item("HANFORD_ID"))) Then
                                                alert("Failed to notify the new user via e-mail.")
                                            End If
                                        Else
                                            If Not (sendMail(Me.dsCommon.Tables("SpecificUserInfo").Rows(0).Item("EMAIL_ADDRESS"), Me.dsCommon.Tables("SpecificUserInfo").Rows(0).Item("EMAIL_ADDRESS"))) Then
                                                alert("Failed to notify the new user via e-mail.")
                                            End If
                                        End If
                                    End If
                                    index += 2
                                End While
                            End If
                            Return False
                        Else
                            sqlCommonAction.Transaction.Rollback()
                            alert("Failed to process the user list.")
                            Return True
                        End If
                    Else
                        alert("You must select a user type and input at least one user to create.")
                        Return True
                    End If
                Else
                    Dim failure As Boolean = False
                    'Check for actual information to insert
                    Session("PersonTable") = Nothing
                    If (Me.txtFirstName.Text <> "" And Me.txtLastName.Text <> "" And Me.txtEmail.Text <> "" And Me.ddlUserType.SelectedValue <> -1) Then
                        'If we're inserting a new record then add it onto the end of the recordset and add it to the existing users
                        If (insertUser(sqlCommonAction, sqlCommonAdapter)) Then
                            sqlCommonAction.Transaction.Commit()
                            Session("transExists") = False
                            If (Me.dsCommon.Tables("UserList").Rows.Count = 0) Then
                                If (Me.dsCommon.Tables("UserIDFetch").Rows(0).Item("USER_CODE") > 0) Then
                                    If Not (sendMail(Me.txtEmail.Text, Session("newUserName"))) Then
                                        alert("Failed to notify the new user via e-mail.")
                                    End If
                                End If
                            End If
                        Else
                            sqlCommonAction.Transaction.Rollback()
                            alert("Failed to insert this new user.")
                            failure = True
                        End If
                    Else
                        alert("You must provide a first and last name as well as an e-mail address and select a user type at minimum.")
                        failure = True
                        sqlCommonAction.Transaction.Rollback()
                    End If
                    Return failure
                End If
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Failed to insert this new user.")
                Return True
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Failed to insert this new user.")
            Return True
        End Try
    End Function

    'attempts to insert a user, returns false if it cannot
    Private Function insertUser(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'clear out table data before use
            If (Me.dsCommon.Tables.Contains("UserList")) Then
                Me.dsCommon.Tables("UserList").Rows.Clear()
            End If
            If (Me.dsCommon.Tables.Contains("PersonData")) Then
                Me.dsCommon.Tables("PersonData").Rows.Clear()
            End If
            If (Me.dsCommon.Tables.Contains("UserIDFetch")) Then
                Me.dsCommon.Tables("UserIDFetch").Rows.Clear()
            End If

            'find out if the user already exists in the system
            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER where EMAIL_ADDRESS = '" & Me.txtEmail.Text & "'"
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            sqlCommonAdapter.Fill(Me.dsCommon, "UserList")

            'if a user exists then return true instead of adding a new person to the user's table
            If (Me.dsCommon.Tables("UserList").Rows.Count = 1) Then
                'determine if the user's level was appropriately selected.  if not bring it up to level 0
                If (Me.ddlUserType.SelectedItem.Value < Me.dsCommon.Tables("UserList").Rows(0).Item("USER_CODE") And Session("UserType") <> 4) Then
                    Me.ddlUserType.SelectedIndex = Me.dsCommon.Tables("UserList").Rows(0).Item("USER_CODE") + 1
                ElseIf (Me.ddlUserType.SelectedItem.Value < 0) Then
                    Me.ddlUserType.SelectedIndex = 1
                End If
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_USER set USER_CODE = " & Me.ddlUserType.SelectedValue
                sqlCommonAction.ExecuteNonQuery()
                Session("manageUserID") = Me.dsCommon.Tables("UserLIst").Rows(0).Item("USER_KEY")
                Return True
            ElseIf (Me.dsCommon.Tables("UserList").Rows.Count > 1) Then
                alert("Not enough information to differentiate between two or more existing users.  The user you are trying to add may already exist in our users list.  Please select the user you are looking for from that list and update their information to indicate their new status for this template.")
                Return False
            End If

            'attempt to add the new record to the database
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
                "isnull(max(employed_by_hcid), '') as COMPANY_ABBREV, isNull(max(active_sw), 'N') as ACTIVE_IND from  (" & _
                "select hanford_id, hanf_pay_no, first_name, last_name, middle_initial, NULL as name_prefix, name_suffix, phone_no, " & _
                "internet_email_address, employed_by_hcid, active_sw from opwhse.dbo.vw_pub_hanford_people union select hanford_id, " & _
                "pay_no as hanf_pay_no, first_name, last_name, middle_name as middle_initial, NULL as name_prefix, name_suffix, " & _
                "work_phone as phone_no, internet_address as internet_email_address, employed_by_hcid, NULL as active_sw " & _
                "from opwhse.dbo.vw_pub_pnnl_associate union select hanford_id, pnnl_pay_no as hanf_pay_no, first_name, last_name, " & _
                "middle_initial, name_prefix, name_suffix, contact_work_phone as phone_no, internet_email_address, company as " & _
                "employed_by_hcid, NULL as active_sw from opwhse.dbo.vw_pub_bmi_employee) a where active_sw = 'Y' and " & _
                "internet_email_address = '" & Me.txtEmail.Text & "' group by hanford_id, " & _
                "last_name, first_name, middle_initial, internet_email_address"
                sqlCommonAdapter.Fill(Me.dsCommon, "PersonData")

                'if we found something store the HID otherwise store the empty string
                If (Me.dsCommon.Tables("PersonData").Rows.Count < 1) Then
                    param7.Value = ""
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

            'determine if the user's level was appropriately selected.  if not bring it up to level 0
            If (Me.ddlUserType.SelectedItem.Value < 0) Then
                Me.ddlUserType.SelectedIndex = 1
            End If
            If (param7.Value = "") Then
                Me.ddlUserType.SelectedIndex = 1
            End If

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

            'check the user's level so that the user entering does not overstep their bounds
            If (Me.ddlUserType.SelectedValue > CType(Session("userType"), Integer)) Then
                Me.ddlUserType.SelectedIndex = CType(Session("userType"), Integer) + 1
            End If

            'insert the new user
            sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_USER (HANFORD_ID, EMAIL_ADDRESS, USER_CODE, LAST_NAME, FIRST_NAME, " & _
            "MIDDLE_NAME, SUFFIX_NAME, AUTHENTICATION_CODE, LAST_CHANGE_AUTH_CD_DATE, PRIMARY_WORK_PHONE, CREATE_DATE, USED_DATE, CREATOR_USER_KEY, ACTIVE_IND, " & _
            "TRAINING_IND, COMPANY_ABBREV, NO_EMAIL_IND) VALUES (@HANFORD_ID, @EMAIL_ADDRESS," & Me.ddlUserType.SelectedValue & ", @LAST_NAME," & _
            "@FIRST_NAME, @MIDDLE_NAME, @SUFFIX_NAME, @AUTHENTICATION_CODE,'" & Now.AddMonths(-7) & "', @PRIMARY_WORK_PHONE,'" & Now & "','1/1/1970'," & _
            Session("manageUserID") & "," & IIf(Me.chkActive.Checked, 1, 0) & "," & IIf(Me.chkTQSurvey.Checked, 1, 0) & _
            ",@COMPANY_ABBREV,0)"
            sqlCommonAction.ExecuteNonQuery()

            'get the user's sequence number
            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER where EMAIL_ADDRESS = '" & Me.txtEmail.Text & "'"
            sqlCommonAdapter.Fill(Me.dsCommon, "UserIDFetch")
            Session("manageUserID") = Me.dsCommon.Tables("UserIDFetch").Rows(0).Item("USER_KEY")
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'attempts to process the manual entry starting the e-mail string build
    Private Function processManualEntry(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'create an array of users while trimming the white space
            Dim requestItems As Array
            Dim index As Integer = 0
            If (Me.txtUserList.Text.Length > 0) Then
                Dim strLine As String = ""
                'replace white space text
                Me.txtUserList.Text = Me.txtUserList.Text.Replace(vbCrLf, ",")
                requestItems = Me.txtUserList.Text.Split(",")
                While index < requestItems.Length
                    requestItems(index) = Trim(CType(requestItems(index), String))
                    index += 1
                End While
            End If

            'determine if we actually got anything.  If not then skip this function
            If (requestItems Is Nothing) Then
                Return True
            End If

            'determine if the person exists as a current user.
            index = 0
            Me.txtUserList.Text = ""
            While index < requestItems.Length
                If (CType(requestItems(index), String) <> "") Then
                    If Not (listGetHanfordPeopleInfo(CType(requestItems(index), String))) Then
                        If (Session("Alert") <> "") Then
                            Session("Alert") &= ", " & CType(requestItems(index), String)
                        Else
                            Session("Alert") = "Could not find the following people: " & CType(requestItems(index), String)
                        End If
                    End If
                    If (Me.dsCommon.Tables.Contains("ExistingUsers")) Then
                        Me.dsCommon.Tables("ExistingUsers").Rows.Clear()
                    End If
                    If (Me.dsCommon.Tables.Contains("PersonData")) Then
                        Me.dsCommon.Tables("PersonData").Rows.Clear()
                    End If
                    If (Me.dsCommon.Tables.Contains("UserIDFetch")) Then
                        Me.dsCommon.Tables("UserIDFetch").Rows.Clear()
                    End If
                End If
                index += 1
            End While
            If (Session("Alert") <> "") Then
                alert(Session("Alert"))
                Session("Alert") = ""
            End If
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
                If (Me.txtFirstName.Text <> "" And Me.txtLastName.Text <> "" And Me.txtEmail.Text <> "" And Me.ddlUserType.SelectedValue <> -1) Then
                    Dim oldUserTypeIndex As Integer = Me.dsCommon.Tables("Users").Rows(Me.ddlUser.SelectedIndex - 1).Item("USER_CODE") + 1
                    If (updateUser(sqlCommonAction, sqlCommonAdapter)) Then
                        sqlCommonAction.Transaction.Commit()
                        If (Me.ddlUserType.SelectedIndex > 1 Or oldUserTypeIndex > 1) Then
                            If Not (sendMail(Me.txtEmail.Text, Session("newUserName"), False, True)) Then
                                alert("Failed to notify the user of the changes via e-mail.")
                            End If
                        End If
                    Else
                        sqlCommonAction.Transaction.Rollback()
                        alert("Failed to update this user.")
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
                alert("Failed to update this user.")
                Return True
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Failed to update this user.")
            Return True
        End Try
    End Function

    'attempts to update a user, returns false if it cannot
    Private Function updateUser(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        'attempt to update the user's record in the database
        Try
            'clear out table data before use
            If (Me.dsCommon.Tables.Contains("CurrentInfo")) Then
                Me.dsCommon.Tables("CurrentInfo").Rows.Clear()
            End If
            If (Me.dsCommon.Tables.Contains("UserList")) Then
                Me.dsCommon.Tables("UserList").Rows.Clear()
            End If

            'find out if the user already exists in the system
            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER where EMAIL_ADDRESS = '" & Me.txtEmail.Text & "'"
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            sqlCommonAdapter.Fill(Me.dsCommon, "UserList")

            'determine if the user owns or delegates any templates
            sqlCommonAction.CommandText = "select * from " & myUtility.getExtension() & "SAT_USER_PERMISSION where USER_KEY = " & _
                Session("ManageUserID") & " and (OWNER_IND = 1 or DELEGATE_IND = 1) and TEMPLATE_KEY <> -1"
            sqlCommonAdapter.Fill(Me.dsCommon, "TemplateOD")

            'determine if the user owns or delegates any surveys
            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER_PERMISSION where USER_KEY = " & _
                Session("ManageUserID") & " and (OWNER_IND = 1 or DELEGATE_IND = 1) and SURVEY_KEY <> -1"
            sqlCommonAdapter.Fill(Me.dsCommon, "SurveyOD")

            'update the user's record
            'users must be at least level 0
            If (Session("UserID") = Session("ManageUserID")) Then
                Me.ddlUserType.SelectedIndex = 5
            Else
                If (Me.dsCommon.Tables("UserList").Rows.Count() > 0) Then
                    If (Me.dsCommon.Tables("TemplateOD").Rows.Count > 0) Then
                        If (Me.ddlUserType.SelectedIndex < 4) Then
                            Me.ddlUserType.SelectedIndex = 4
                            alert("You may not lower this user\'s access below \""Template and Survey Administrator\"" because they own/delegate a template.")
                        End If
                    ElseIf (Me.dsCommon.Tables("SurveyOD").Rows.Count > 0) Then
                        If (Me.ddlUserType.SelectedIndex < 3) Then
                            Me.ddlUserType.SelectedIndex = 3
                            alert("You may not lower this user\'s access below \""Survey Administrator\"" because they own/delegate a survey.")
                        End If
                    ElseIf (Me.ddlUserType.SelectedItem.Value < 0) Then
                        Me.ddlUserType.SelectedIndex = 1
                    ElseIf (Me.ddlUserType.SelectedItem.Value > Session("UserType")) Then
                        Me.ddlUserType.SelectedIndex = Session("UserType") + 1
                        alert("You may not set this user\'s access above your own level of access.")
                    End If
                Else
                    If (Me.dsCommon.Tables("TemplateOD").Rows.Count > 0) Then
                        If (Me.ddlUserType.SelectedIndex < 4) Then
                            Me.ddlUserType.SelectedIndex = 4
                            alert("You may not lower this user\'s access below \""Template and Survey Administrator\"" because they own/delegate a template.")
                        End If
                    ElseIf (Me.dsCommon.Tables("SurveyOD").Rows.Count > 0) Then
                        If (Me.ddlUserType.SelectedIndex < 3) Then
                            Me.ddlUserType.SelectedIndex = 3
                            alert("You may not lower this user\'s access below \""Survey Administrator\"" because they own/delegate a survey.")
                        End If
                    ElseIf (Me.ddlUserType.SelectedItem.Value < 0) Then
                        Me.ddlUserType.SelectedIndex = 1
                    ElseIf (Me.ddlUserType.SelectedItem.Value > Session("UserType")) Then
                        Me.ddlUserType.SelectedIndex = Session("UserType") + 1
                        alert("You may not set this user\'s access above your own level of access.")
                    End If
                End If
            End If

            'get the user's HID, e-mail, and user type
            sqlCommonAction.CommandText = "select * from " & _
            myUtility.getExtension() & "SAT_USER where USER_KEY = " & Session("manageUserID")
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            sqlCommonAdapter.Fill(Me.dsCommon, "CurrentInfo")

            'if a user was returned then go ahead and update the user
            If (Me.dsCommon.Tables("CurrentInfo").Rows.Count() > 0) Then
                If (Me.dsCommon.Tables("CurrentInfo").Rows(0).Item("HANFORD_ID") <> "") Then
                    'check to see if this new person has the same HID as the old person
                    sqlCommonAction.CommandText = "select hanford_id, last_name, first_name, middle_initial as MIDDLE_NAME, internet_email_address as EMAIL_ADDRESS from (" & _
                    "select hanford_id, hanf_pay_no, first_name, last_name, middle_initial, NULL as name_prefix, name_suffix, phone_no, " & _
                    "internet_email_address, employed_by_hcid, active_sw from opwhse.dbo.vw_pub_hanford_people union select hanford_id, " & _
                    "pay_no as hanf_pay_no, first_name, last_name, middle_name as middle_initial, NULL as name_prefix, name_suffix, " & _
                    "work_phone as phone_no, internet_address as internet_email_address, employed_by_hcid, NULL as active_sw " & _
                    "from opwhse.dbo.vw_pub_pnnl_associate union select hanford_id, pnnl_pay_no as hanf_pay_no, first_name, last_name, " & _
                    "middle_initial, name_prefix, name_suffix, contact_work_phone as phone_no, internet_email_address, company as " & _
                    "employed_by_hcid, NULL as active_sw from opwhse.dbo.vw_pub_bmi_employee) a where active_sw = 'Y' and " & _
                    "internet_email_address = '" & Me.txtEmail.Text & "'" & _
                    " group by hanford_id, last_name, first_name, middle_initial, internet_email_address "
                    sqlCommonAdapter.Fill(Me.dsCommon, "ChangeUserCheck")
                    If (Me.dsCommon.Tables("ChangeUserCheck").Rows.Count > 0) Then
                        If (Me.dsCommon.Tables("ChangeUserCheck").Rows(0).Item("hanford_id") <> Me.dsCommon.Tables("CurrentInfo").Rows(0).Item("HANFORD_ID")) Then
                            alert("It appears that you are trying to change this user into someone completely different.  You cannot do this.  You must create a new user.")
                            Return False
                        End If
                    Else
                        alert("It appears that you are trying to change this user into someone completely different.  You cannot do this.  You must create a new user.")
                        Return False
                    End If
                Else
                    alert("You may not set non-staff user access levels any higher than survey respondent.")
                    Me.ddlUserType.SelectedIndex = 1
                End If

                'check for a change in e-mail address that conflicts with an existing address
                If (Me.dsCommon.Tables("CurrentInfo").Rows(0).Item("EMAIL_ADDRESS") <> Me.txtEmail.Text) Then
                    sqlCommonAction.CommandText = "select * from " & myUtility.getExtension() & "SAT_USER " & _
                    "where EMAIL_ADDRESS = '" & Me.txtEmail.Text & "'"
                    sqlCommonAdapter.Fill(Me.dsCommon, "CheckEmailCollision")
                    If (Me.dsCommon.Tables("CheckEmailCollision").Rows.Count() > 0) Then
                        alert("A user with this e-mail address already exists.  You cannot create multiple users with the same e-mail address.")
                        Return False
                    End If
                End If

                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_USER SET EMAIL_ADDRESS = @EMAIL_ADDRESS, FIRST_NAME = @FIRST_NAME" & _
                ", LAST_NAME = @LAST_NAME, MIDDLE_NAME = @MIDDLE_NAME, SUFFIX_NAME = @SUFFIX_NAME, " & _
                "PRIMARY_WORK_PHONE = @PRIMARY_WORK_PHONE, COMPANY_ABBREV = @COMPANY_ABBREV, USER_CODE = " & Me.ddlUserType.SelectedValue & _
                ", TRAINING_IND = " & IIf(Me.chkTQSurvey.Checked, 1, 0) & ", ACTIVE_IND = " & IIf(Me.chkActive.Checked, 1, 0)

                'if the password reset was checked then generate a new password and store it
                If (Me.chkPasswordReset.Checked) Then
                    sqlCommonAction.CommandText &= ", AUTHENTICATION_CODE = '"
                    If (Me.dsCommon.Tables("CurrentInfo").Rows(0).Item("HANFORD_ID") = "") Then
                        sqlCommonAction.CommandText &= generatePassword(Me.dsCommon.Tables("CurrentInfo").Rows(0).Item("EMAIL_ADDRESS")) & "'"
                        Session("newUserName") = Me.txtEmail.Text
                    Else
                        sqlCommonAction.CommandText &= generatePassword(Me.dsCommon.Tables("CurrentInfo").Rows(0).Item("HANFORD_ID")) & "'"
                        Session("newUserName") = Me.dsCommon.Tables("CurrentInfo").Rows(0).Item("HANFORD_ID")
                    End If
                    sqlCommonAction.CommandText &= ", LAST_CHANGE_AUTH_CD_DATE = '" & Now.AddMonths(-7) & "'"
                End If

                'determine if the active status has changed
                If (CType(Me.dsCommon.Tables("CurrentInfo").Rows(0).Item("ACTIVE_IND"), Boolean) <> Me.chkActive.Checked) Then
                    Session("ActiveStatusChanged") = True
                Else
                    Session("ActiveStatusChanged") = False
                End If

                'determine if the user type has changed
                If (Me.dsCommon.Tables("CurrentInfo").Rows(0).Item("USER_CODE") <> Me.ddlUserType.SelectedValue) Then
                    Session("UserTypeChanged") = True
                Else
                    Session("UserTypeChanged") = False
                End If

                'finish the sql statement
                sqlCommonAction.CommandText &= " where USER_KEY = " & Session("manageUserID")

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
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
        Return True
    End Function
#End Region

#Region "Delete Code"
    'deletes the current user
    Private Function doDelete() As Boolean
        Try
            If (Session("UserID") <> Session("ManageUserID")) Then
                'start transaction
                If (myUtility.setupTransaction(sqlCommonAction, Me.commonConn)) Then
                    Dim failure As Boolean = False

                    'do the deletion
                    If (deletePermission(sqlCommonAction, sqlCommonAdapter)) Then
                        If (deleteUser(sqlCommonAction, sqlCommonAdapter)) Then
                            sqlCommonAction.Transaction.Commit()

                            'if the user being modified is also the user using this page then 
                            'dump them onto the login page
                            If (Session("manageUserID") = Session("UserID")) Then
                                Session("CurrentPage") = "./Login.aspx"
                                Session("pageSel") = "Login"
                                redirect(Session("CurrentPage"))
                            End If

                            'notify the user of their status change
                            If (Me.ddlUserType.SelectedIndex > 1) Then
                                If Not (sendMail(Me.txtEmail.Text, "Delete", True)) Then
                                    alert("Failed to notify the user of their removal from the system.")
                                End If
                            End If

                            Session("manageUserID") = -1
                        Else
                            sqlCommonAction.Transaction.Rollback()
                            failure = True
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
                    alert("Failed to delete this user.")
                    Return True
                End If
            Else
                alert("You may not remove yourself from the tool.  If this must be done find another Survey Tool Administrator and they can do this for you.")
                Return True
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Failed to delete this user.")
            Return True
        End Try
    End Function

    'attempts to delete a user from the fempPermission table.  Returns false if it cannot for any reason.
    Private Function deletePermission(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'attempt to delete the record from the database
            sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_USER_PERMISSION where USER_KEY = " & _
            Session("manageUserID")
            sqlCommonAction.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'attempts to delete a user from the fempPermission table.  Returns false if it cannot for any reason.
    Private Function deleteUser(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'attempt to delete the record from the database
            sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_USER where USER_KEY = " & _
            Session("manageUserID")
            sqlCommonAction.ExecuteNonQuery()

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
            alert("Failed to load the users of the Survey Administration tool.")
        Else
            Return False
        End If
    End Function
#End Region

#Region "Password Generation"
    'Generates the user's password
    Private Function generatePassword(ByVal strValue As String) As String
        Trace.Warn("Generating Password")

        'Generate password for the user
        strNewPassword = myUtility.createPasswordString()

        'Generate password hash for the user
        Dim strPassword As String = ""
        strPassword = myUtility.hashPassword(strNewPassword, strValue)

        Return strPassword
    End Function
#End Region

#Region "Send Mail"
    'Sends an e-mail to users with changed passwords, except level 0 users.
    Private Function sendMail(ByVal Email As String, ByVal strUserName As String, Optional ByVal inSystem As Boolean = False, Optional ByVal isUpdate As Boolean = False) As Boolean
        Trace.Warn("Sending user an email")
        'Set up the mail messsage
        Try
            Dim strFrom As String = "survey.administrator@pnl.gov"
            Dim strTo As String = Email
            Dim strbldrSubject As New StringBuilder("Survey Admin Tool User Update Notice.")

            'get the current user's name and e-mail
            Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
            Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            sqlCommonAction.Connection = Me.commonConn
            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER where USER_KEY = " & Session("UserID")
            sqlCommonAdapter.Fill(Me.dsCommon, "ModifyID")

            If (Me.dsCommon.Tables("ModifyID").Rows.Count < 1) Then
                alert("Unable to locate your information.  You may have been removed from the database in the last few minutes.  Contact the Survey Administrator immediately.")
                Return False
            End If

            'to be removed!'
            'strTo = Me.dsCommon.Tables("ModifyID").Rows(0).Item("strEmail")

            Dim strbldrMessage As New StringBuilder
            'notify the user of their status change
            If (strUserName = "Delete") Then
                strbldrMessage.Append("You have been removed as a user of Survey Administration Tool by")
                strbldrMessage.Append(" " & Me.dsCommon.Tables("ModifyID").Rows(0).Item("FIRST_NAME") & " ")
                strbldrMessage.Append(Me.dsCommon.Tables("ModifyID").Rows(0).Item("LAST_NAME") & ".")
                strbldrMessage.Append("  You may contact this person at this e-mail address: <b>")
                strbldrMessage.Append(Me.dsCommon.Tables("ModifyID").Rows(0).Item("EMAIL_ADDRESS") & "</b>")
            ElseIf (isUpdate) Then
                If (Session("UserTypeChanged") Or Session("ActiveStatusChanged") Or Me.chkPasswordReset.Checked) Then
                    'notify the user of a change in user type
                    If (Session("UserTypeChanged")) Then
                        strbldrMessage.Append("Your user access level has been changed to <b>" & _
                        Me.ddlUserType.SelectedItem.Text & "</b>.<br/>")
                        If (Me.ddlUserType.SelectedValue = 0) Then
                            strbldrMessage.Append("As a result of this change you will not be able to " & _
                            "log into the Survey Administration tool.<br/>")
                        End If
                    End If

                    'notify the user of a change in status
                    If (Session("ActiveStatusChanged")) Then
                        If (Me.chkActive.Checked) Then
                            strbldrMessage.Append("You have been activated as a user in the Survey Administration tool.<br/>")
                        Else
                            strbldrMessage.Append("You have been de-activated as a user in the Survey Administration tool.<br/>" & _
                            "You wil be re-actived the next time you use the Survey Administration tool.<br/>")
                        End If
                    End If

                    'notify the external user of their e-mail address change
                    If (Session("NoHIDEmailChanged")) Then
                        strbldrMessage.Append("Your E-mail address has changed in the Survey Administration tool.<br/>")
                    End If

                    'notify the user of a change in password
                    If (Me.chkPasswordReset.Checked) Then
                        strbldrMessage.Append("Your password has been reset.<br/>")
                    End If

                    'inform the user of the modifying entity
                    strbldrMessage.Append("This change was made by " & Me.dsCommon.Tables("ModifyID").Rows(0).Item("FIRST_NAME") & " ")
                    strbldrMessage.Append(Me.dsCommon.Tables("ModifyID").Rows(0).Item("LAST_NAME") & ".")
                    strbldrMessage.Append("  You may contact this person at this e-mail address: <b>")
                    strbldrMessage.Append(Me.dsCommon.Tables("ModifyID").Rows(0).Item("EMAIL_ADDRESS") & "</b><br/>")

                    'notify the user of a change in password
                    If (Me.chkPasswordReset.Checked) Then
                        strbldrMessage.Append("Please Make a note of your password.<br/> If you wish to change your password ")
                        strbldrMessage.Append("you may do so at the ")
                        strbldrMessage.Append("<a href='http://" & Request.ServerVariables.Item("SERVER_NAME") & "/Survey Admin/Default.aspx'>Survey Administration</a> website. <p></p>")
                        strbldrMessage.Append("Your User Name is: " & strUserName & "<br/>")
                        strbldrMessage.Append("Your Password is: " & strNewPassword)
                    End If
                End If
            Else
                strbldrMessage.Append("You have been given " & Me.ddlUserType.SelectedItem.Text & " access to the Survey Administration Tool by")
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
            If (isUpdate And (Session("UserTypeChanged") Or Session("ActiveStatusChanged") Or Me.chkPasswordReset.Checked)) Then
                If (myUtility.genericSendMail(strFrom, strTo, strbldrSubject.ToString, strbldrMessage.ToString)) Then
                    Return True
                Else
                    Return False
                End If
            ElseIf Not (isUpdate) Then
                If (myUtility.genericSendMail(strFrom, strTo, strbldrSubject.ToString, strbldrMessage.ToString)) Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return True
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function
#End Region
End Class
