Imports System.Collections.Specialized
Public Class Permission
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Permission))
        Me.sqlUsers = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand
        Me.commonConn = New System.Data.SqlClient.SqlConnection
        Me.dsCommon = New System.Data.DataSet
        Me.sqlPermission = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlDeleteCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlInsertCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlSelectCommand2 = New System.Data.SqlClient.SqlCommand
        Me.SqlUpdateCommand1 = New System.Data.SqlClient.SqlCommand
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'sqlUsers
        '
        Me.sqlUsers.SelectCommand = Me.SqlSelectCommand1
        Me.sqlUsers.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_USER_PERMISSION", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("USER_KEY", "USER_KEY"), New System.Data.Common.DataColumnMapping("USER_CODE", "USER_CODE"), New System.Data.Common.DataColumnMapping("UserName", "UserName"), New System.Data.Common.DataColumnMapping("HANFORD_ID", "HANFORD_ID"), New System.Data.Common.DataColumnMapping("EMAIL_ADDRESS", "EMAIL_ADDRESS"), New System.Data.Common.DataColumnMapping("USER_IND", "USER_IND")})})
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = resources.GetString("SqlSelectCommand1.CommandText")
        Me.SqlSelectCommand1.Connection = Me.commonConn
        '
        'dsCommon
        '
        Me.dsCommon.DataSetName = "NewDataSet"
        Me.dsCommon.Locale = New System.Globalization.CultureInfo("en-US")
        '
        'sqlPermission
        '
        Me.sqlPermission.DeleteCommand = Me.SqlDeleteCommand1
        Me.sqlPermission.InsertCommand = Me.SqlInsertCommand1
        Me.sqlPermission.SelectCommand = Me.SqlSelectCommand2
        Me.sqlPermission.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_USER_PERMISSION", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("USER_KEY", "USER_KEY"), New System.Data.Common.DataColumnMapping("CREATOR_USER_KEY", "CREATOR_USER_KEY"), New System.Data.Common.DataColumnMapping("TEMPLATE_KEY", "TEMPLATE_KEY"), New System.Data.Common.DataColumnMapping("SURVEY_KEY", "SURVEY_KEY"), New System.Data.Common.DataColumnMapping("CREATE_DATE", "CREATE_DATE"), New System.Data.Common.DataColumnMapping("OWNER_IND", "OWNER_IND"), New System.Data.Common.DataColumnMapping("DELEGATE_IND", "DELEGATE_IND"), New System.Data.Common.DataColumnMapping("USER_IND", "USER_IND")})})
        Me.sqlPermission.UpdateCommand = Me.SqlUpdateCommand1
        '
        'SqlDeleteCommand1
        '
        Me.SqlDeleteCommand1.CommandText = "DELETE FROM VARCONN.[SAT_USER_PERMISSION] WHERE (([USER_KEY] = @Original" & _
            "_USER_KEY) AND ([TEMPLATE_KEY] = @Original_TEMPLATE_KEY) AND ([SURVEY_KEY] = @Or" & _
            "iginal_SURVEY_KEY))"
        Me.SqlDeleteCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@Original_USER_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "USER_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_SURVEY_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "SURVEY_KEY", System.Data.DataRowVersion.Original, Nothing)})
        '
        'SqlInsertCommand1
        '
        Me.SqlInsertCommand1.CommandText = resources.GetString("SqlInsertCommand1.CommandText")
        Me.SqlInsertCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@USER_KEY", System.Data.SqlDbType.Int, 0, "USER_KEY"), New System.Data.SqlClient.SqlParameter("@CREATOR_USER_KEY", System.Data.SqlDbType.Int, 0, "CREATOR_USER_KEY"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, "TEMPLATE_KEY"), New System.Data.SqlClient.SqlParameter("@SURVEY_KEY", System.Data.SqlDbType.Int, 0, "SURVEY_KEY"), New System.Data.SqlClient.SqlParameter("@CREATE_DATE", System.Data.SqlDbType.DateTime, 0, "CREATE_DATE"), New System.Data.SqlClient.SqlParameter("@OWNER_IND", System.Data.SqlDbType.Bit, 0, "OWNER_IND"), New System.Data.SqlClient.SqlParameter("@DELEGATE_IND", System.Data.SqlDbType.Bit, 0, "DELEGATE_IND"), New System.Data.SqlClient.SqlParameter("@USER_IND", System.Data.SqlDbType.Bit, 0, "USER_IND")})
        '
        'SqlSelectCommand2
        '
        Me.SqlSelectCommand2.CommandText = "SELECT USER_KEY, CREATOR_USER_KEY, TEMPLATE_KEY, SURVEY_KEY, CREATE_DATE, OWNER_I" & _
            "ND, DELEGATE_IND, USER_IND FROM VARCONN.SAT_USER_PERMISSION"
        '
        'SqlUpdateCommand1
        '
        Me.SqlUpdateCommand1.CommandText = resources.GetString("SqlUpdateCommand1.CommandText")
        Me.SqlUpdateCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@USER_KEY", System.Data.SqlDbType.Int, 0, "USER_KEY"), New System.Data.SqlClient.SqlParameter("@CREATOR_USER_KEY", System.Data.SqlDbType.Int, 0, "CREATOR_USER_KEY"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, "TEMPLATE_KEY"), New System.Data.SqlClient.SqlParameter("@SURVEY_KEY", System.Data.SqlDbType.Int, 0, "SURVEY_KEY"), New System.Data.SqlClient.SqlParameter("@CREATE_DATE", System.Data.SqlDbType.DateTime, 0, "CREATE_DATE"), New System.Data.SqlClient.SqlParameter("@OWNER_IND", System.Data.SqlDbType.Bit, 0, "OWNER_IND"), New System.Data.SqlClient.SqlParameter("@DELEGATE_IND", System.Data.SqlDbType.Bit, 0, "DELEGATE_IND"), New System.Data.SqlClient.SqlParameter("@USER_IND", System.Data.SqlDbType.Bit, 0, "USER_IND"), New System.Data.SqlClient.SqlParameter("@Original_USER_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "USER_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_SURVEY_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "SURVEY_KEY", System.Data.DataRowVersion.Original, Nothing)})
        '
        'commonConn
        '
        Me.commonConn.FireInfoMessageEventOnUserErrors = False
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected WithEvents Image1 As System.Web.UI.WebControls.Image
    Protected WithEvents lblTemplate As System.Web.UI.WebControls.Label
    Protected WithEvents hlpLevel0 As System.Web.UI.WebControls.Image
    Protected WithEvents ddlPageOptions As System.Web.UI.WebControls.DropDownList
    Protected WithEvents sqlUsers As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents dsCommon As System.Data.DataSet
    Protected WithEvents sqlPermission As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents SqlSelectCommand2 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlInsertCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlUpdateCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlDeleteCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents DataSet11 As SurveyAdmin.DataSet1
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents btnTop As System.Web.UI.WebControls.Button
    Protected WithEvents dgUsers As System.Web.UI.WebControls.DataGrid
    Protected myUtility As Utility = New Utility
    Protected WithEvents lblInGroup As System.Web.UI.WebControls.Label
    Protected WithEvents lblLevel0None As System.Web.UI.WebControls.Label
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents imgA1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgB1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgC1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgD1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgE1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgF1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgG1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgH1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgI1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgJ1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgK1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgL1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgM1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgN1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgO1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgP1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgQ1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgR1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgS1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgT1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgU1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgV1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgW1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgX1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgY1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgZ1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents Image2 As System.Web.UI.WebControls.Image
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents imgA2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgB2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgC2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgD2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgE2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgF2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgG2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgH2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgI2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgJ2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgK2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgL2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgM2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgN2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgO2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgP2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgQ2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgR2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgS2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgT2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgU2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgV2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgW2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgX2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgY2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents imgZ2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnSubmit As System.Web.UI.WebControls.Button
    Protected WithEvents dgTemplateUsers As System.Web.UI.WebControls.DataGrid
    Protected strUserList As String
    Protected WithEvents hlpAlphabet As System.Web.UI.WebControls.Image
    Protected WithEvents hlpAlphabet2 As System.Web.UI.WebControls.Image
    Protected WithEvents commonConn As System.Data.SqlClient.SqlConnection
    Protected WithEvents btnUpdate As System.Web.UI.WebControls.Button
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

        'if the session has timed out store an alert message we'll be needing it in a sec
        If (Session.IsNewSession) Then
            Session("Alert") = "Your session has timed out for security purposes.  Please log in again."
        End If

        'Put user code to initialize the page here
        If (Session("isAuthenticated") <> "True") Then
            Session("CurrentPage") = "./Login.aspx"
            Session("pageSel") = "Login"
            Response.Redirect("./Login.aspx")
        ElseIf (myUtility.checkUserType(3, CType(Session("isTemplateOwner"), Boolean), CType(Session("isTemplateDelegate"), Boolean), True)) Then
            Response.Redirect(Session("CurrentPage"))
        Else
            Session("CurrentPage") = "./Permission.aspx"
            Session("pageSel") = "Template"
        End If

        'if the user's password has expired send them to the password reset page
        If (Session("isExpired") = "True") Then
            Session("CurrentPage") = "./Password.aspx"
            Session("pageSel") = "Password"
            redirect(Session("CurrentPage"))
        End If

        'Finalize the SQL statements
        completeSQL()

        'Set the connection based on the machine
        myUtility.getConnection(Me.commonConn)

        'Fill the user grid
        If Not (IsPostBack) Then
            'get anything on the request line
            myUtility.getRequest(Page.Request.Url.Query())

            'set the initial alphabetical limiter
            If (Session("Alpha") Is Nothing Or Session("Alpha") = "") Then
                Session("Alpha") = "A"
            End If

            'load the user data
            loadData()

            'Checks the grid for content
            checkGrid()

            'add an onclick attribute to the submit button to make sure the user wants to submit the page
            Me.btnUpdate.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
            Me.btnReset.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
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

    'Finalize the SQL statements
    Private Sub completeSQL()
        Me.SqlSelectCommand1.CommandText &= " and fp.TEMPLATE_KEY = " & Session("seqTemplate") & " order by UserName, " & _
            "fu.HANFORD_ID, fu.EMAIL_ADDRESS"
    End Sub
#End Region

#Region "Data Load"
    'Loads data into the form
    Private Sub loadData(Optional ByVal override As String = "", Optional ByVal failure As Boolean = False)
        Trace.Warn("Loading Data")
        If (override <> "Reload") Then
            'Fill the Groups List
            loadTemplateUsers()

            'Fill the Group Users List
            loadUsers()
        Else
            'Fill the Groups List
            loadTemplateUsers(override, failure)

            'Fill the Group Users List
            loadUsers(override, failure)
        End If
    End Sub

    'loads the users
    Private Sub loadTemplateUsers(Optional ByVal override As String = "", Optional ByVal failure As Boolean = False)
        Trace.Warn("Loading Group Users")
        Try
            'setup the ad-hoc sql tools
            Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
            Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            sqlCommonAction.Connection = Me.commonConn

            'get the user information necessary for the group user listing
            sqlCommonAction.CommandText = "SELECT fu.USER_KEY, fu.USER_CODE, fu.LAST_NAME + ', ' + fu.FIRST_NAME + ' ' + fu.MIDDLE_NAME" & _
            " + ' ' + fu.SUFFIX_NAME AS Name, fu.HANFORD_ID, fu.EMAIL_ADDRESS FROM " & myUtility.getExtension() & "SAT_USER fu, " & _
            myUtility.getExtension() & "SAT_USER_PERMISSION fp " & _
            "WHERE fu.USER_KEY = fp.USER_KEY and USER_IND = 1 and fp.TEMPLATE_KEY = " & Session("seqTemplate") & " ORDER BY Name"

            sqlCommonAdapter.Fill(Me.dsCommon, "TemplateUsers")
            Me.dgTemplateUsers.DataSource = Me.dsCommon.Tables("TemplateUsers")
            Me.dgTemplateUsers.DataBind()
            myUtility.checkBoxes(Me.dgTemplateUsers)
            If (Me.dsCommon.Tables("TemplateUsers").Rows.Count() = 0) Then
                Session("TemplateUserListRows") = "None"
                strUserList = ""
            Else
                Session("TemplateUserListRows") = "Some"
                strUserList = ""
                Dim index As Integer = 0
                While (index < Me.dsCommon.Tables("TemplateUsers").Rows.Count())
                    If (index <> Me.dsCommon.Tables("TemplateUsers").Rows.Count() - 1) Then
                        strUserList &= Me.dsCommon.Tables("TemplateUsers").Rows(index).Item("USER_KEY") & ", "
                    Else
                        strUserList &= Me.dsCommon.Tables("TemplateUsers").Rows(index).Item("USER_KEY")
                    End If
                    index += 1
                End While
            End If
        Catch ex As Exception
            alert("Error getting group user information.")
            Session("TemplatepUserListRows") = "None"
            Trace.Warn(ex.ToString)
        End Try
    End Sub

    'loads the users
    Private Sub loadUsers(Optional ByVal override As String = "", Optional ByVal failure As Boolean = False)
        Trace.Warn("Loading Group Information")
        Try
            If (strUserList <> "") Then
                Me.SqlSelectCommand1.CommandText = "SELECT USER_KEY, USER_CODE, LAST_NAME + ', ' + FIRST_NAME + ' ' + MIDDLE_NAME" & _
                " + ' ' + SUFFIX_NAME AS Name, HANFORD_ID, EMAIL_ADDRESS FROM " & myUtility.getExtension() & "SAT_USER " & _
                "WHERE USER_KEY not in (" & strUserList & ") and upper(LAST_NAME) like '" & _
                Session("Alpha") & "%' and (HANFORD_ID <> '' or LAST_NAME = 'Administrator' " & _
                "or (LAST_NAME = 'User' and FIRST_NAME = 'Test')) ORDER BY Name"
            Else
                Me.SqlSelectCommand1.CommandText = "SELECT USER_KEY, USER_CODE, LAST_NAME + ', ' + FIRST_NAME + ' ' + MIDDLE_NAME" & _
                " + ' ' + SUFFIX_NAME AS Name, HANFORD_ID, EMAIL_ADDRESS FROM " & myUtility.getExtension() & "SAT_USER " & _
                "WHERE upper(LAST_NAME) like '" & _
                Session("Alpha") & "%' and (HANFORD_ID <> '' or LAST_NAME = 'Administrator' " & _
                "or (LAST_NAME = 'User' and FIRST_NAME = 'Test')) ORDER BY Name"
            End If

            Me.SqlSelectCommand1.Connection = Me.commonConn
            Me.sqlUsers.Fill(Me.dsCommon, "Users")
            Me.dgUsers.DataSource = Me.dsCommon.Tables("Users")
            Me.dgUsers.DataBind()
            If (Me.dsCommon.Tables("Users").Rows.Count() = 0) Then
                Session("GeneralUserListRows") = "None"
            Else
                Session("GeneralUserListRows") = "Some"
            End If
        Catch ex As Exception
            alert("Error getting user information.")
            Session("GeneralUserListRows") = "None"
            Trace.Warn(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Page Switch"
    'Returns the user back to the template
    Private Sub btnTop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTop.Click
        Session("JSAction") = ""
        Session("CurrentPage") = "./Template.aspx"
        Session("pageSel") = "Template"
        Response.Redirect("./Template.aspx")
    End Sub
#End Region

#Region "Grid"
    'Weigh station for processing the checking and unchecking of checkboxes in the grids 
    Public Sub commandBtnClick(ByVal sender As Object, ByVal e As CommandEventArgs)
        Trace.Warn("Header Button Click")
        Select Case e.CommandName
            Case "SelectUsers"
                myUtility.checkBoxes(Me.dgUsers)
            Case "SelectGroup"
                myUtility.checkBoxes(Me.dgTemplateUsers)
        End Select
    End Sub

    'makes changes to the user list based on selections in both grids
    Public Sub commandGridUserList(ByVal sender As Object, ByVal e As CommandEventArgs)
        Trace.Warn("Grid Alphabet Button Click")
        Session("Alpha") = e.CommandName

        Dim failure As Boolean = False

        'attempt to apply any changes made
        failure = applyUserPermissions()

        'loads the user data
        loadData("Reload", failure)

        'Checks the grid for content
        checkGrid()
    End Sub
#End Region

#Region "Checks"
    'Checks the user grid for content
    Private Sub checkGrid()
        Trace.Warn("Checking the grid")
        'Check level 0 Users
        If (Me.dgUsers.Items.Count() = 0) Then
            Session("UserAvailability") = "None"
        Else
            Session("UserAvailability") = "Some"
        End If
    End Sub
#End Region

#Region "Page Action"
    'Reacts to the user performing some kind of selection
    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        
        'do an update/reset based on the page option selected
        failure = applyUserPermissions()

        'determine if we need to reload the data
        loadData("Reload", failure)
        
        'Checks the grid for content
        checkGrid()
    End Sub

    'Reacts to the user performing some kind of selection
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Session("JSAction") = ""
        Dim failure As Boolean
        
        'do an update/reset based on the page option selected
        failure = doReset()
        
        'Checks the grid for content
        checkGrid()
    End Sub
#End Region

#Region "Insert/Update Code"
    'applies the permissions based on the checked users for the template
    Private Function applyUserPermissions() As Boolean
        Trace.Warn("Applying User Permissions")
        'Set up the common data adapter and delete command
        Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
        Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
        Try
            'set up the transaction
            If (myUtility.setupTransaction(sqlCommonAction, Me.commonConn)) Then
                'process the template user list for changes
                If (processTemplateUserList(sqlCommonAction, sqlCommonAdapter)) Then
                    'process the general user list for changes
                    If (processUserList(sqlCommonAction, sqlCommonAdapter)) Then
                        sqlCommonAction.Transaction.Commit()
                        alert("Updated the user status of users in the Survey Admin system.")
                        Return True
                    Else
                        sqlCommonAction.Transaction.Rollback()
                        alert("Failed to update the user status of users in the general user list")
                        Return False
                    End If
                Else
                    sqlCommonAction.Transaction.Rollback()
                    alert("Failed to update the user status of users in the template user list")
                    Return False
                End If
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Permission setting failure.")
                Return True
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Permission setting failure.")
            Return True
        End Try
    End Function

    'attempts to remove users who have been de-selected from the template user list, returns false if it cannot
    Private Function processTemplateUserList(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Trace.Warn("Checking Template Users")
        Dim Index As Integer = 0
        While (Index < Me.dgTemplateUsers.Items.Count())
            If (CType(Me.dgTemplateUsers.Items(Index).Controls(0).Controls(1), CheckBox).Checked = False) Then
                If Not (removeUserStatus(sqlCommonAction, sqlCommonAdapter, CType(Me.dgTemplateUsers.Items(Index).Controls(1), TableCell).Text)) Then
                    Return False
                End If
            End If
            Index += 1
        End While
        Return True
    End Function

    'attempts to insert/update permissions for users who have been selected in the general user list, returns false if it cannot
    Private Function processUserList(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Trace.Warn("Checking General Users")
        Dim Index As Integer = 0
        While (Index < Me.dgUsers.Items.Count())
            If (CType(Me.dgUsers.Items(Index).Controls(0).Controls(1), CheckBox).Checked = True) Then
                If Not (addUser(sqlCommonAction, sqlCommonAdapter, CType(Me.dgUsers.Items(Index).Controls(1), TableCell).Text)) Then
                    Return False
                End If
            End If
            Index += 1
        End While
        Return True
    End Function

    'attempts to remove a user's user permission for a template, returns false if it cannot
    Private Function removeUserStatus(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal userID As Integer) As Boolean
        'attempt to set the user flag to 0 to remove their user permissions
        Try
            'determine if the user is a reviewer
            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & _
            "SAT_TEMPLATE_SIGNATURE fss, " & myUtility.getExtension() & "SAT_USER fu " & _
            "where fss.HANFORD_ID = fu.HANFORD_ID and fu.USER_KEY = " & userID & " and fss.TEMPLATE_KEY = " & _
            Session("seqTemplate")
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            If (Me.dsCommon.Tables.Contains("Reviewer")) Then
                Me.dsCommon.Tables("Reviewer").Rows.Clear()
            End If
            sqlCommonAdapter.Fill(Me.dsCommon, "Reviewer")

            'determine if the user being deleted also owns or is a delegate of any surveys
            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY fs, " & _
            myUtility.getExtension() & "SAT_USER_PERMISSION fp, " & myUtility.getExtension() & "SAT_USER fu " & _
            "where fs.SURVEY_KEY = fp.SURVEY_KEY and fs.TEMPLATE_KEY = " & Session("seqTemplate") & _
            " and fp.USER_KEY = " & userID & " and fp.USER_KEY = fu.USER_KEY and (OWNER_IND = 1 or DELEGATE_IND = 1)"
            sqlCommonAdapter.Fill(Me.dsCommon, "UserSurveys")

            'update the user based on if there are results or if the template is approved
            If ((Me.dsCommon.Tables("Reviewer").Rows.Count = 0 And Me.dsCommon.Tables("UserSurveys").Rows.Count = 0) _
                    Or Session("isApproved")) Then
                sqlCommonAction.CommandText = "UPDATE " & myUtility.getExtension() & "SAT_USER_PERMISSION set USER_IND = 0 where " & _
                "TEMPLATE_KEY = " & Session("seqTemplate") & " and USER_KEY = " & userID
                sqlCommonAction.ExecuteNonQuery()
            Else
                If (Me.dsCommon.Tables("Reviewer").Rows.Count > 0) Then
                    alert("You may not remove " & Me.dsCommon.Tables("Reviewer").Rows(0).Item("NAME") & _
                      " at this time.  This template has not been fully approved and this person is a reviewer.")
                    Return False
                Else
                    Dim strAlert As String = ""
                    strAlert = "You may not remove " & Me.dsCommon.Tables("UserSurveys").Rows(0).Item("LAST_NAME") & _
                      ", " & Me.dsCommon.Tables("UserSurveys").Rows(0).Item("FIRST_NAME") & " "
                    If (Trim(Me.dsCommon.Tables("UserSurveys").Rows(0).Item("MIDDLE_NAME")) <> "") Then
                        strAlert &= Me.dsCommon.Tables("UserSurveys").Rows(0).Item("MIDDLE_NAME") & " "
                    End If
                    strAlert &= "at this time.  This template has not been fully approved and this person is a survey owner/delegate."
                    alert(strAlert)
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    Private Function addUser(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal userID As Integer) As Boolean
        Dim blnUpdated As Boolean = True
        Dim blnInserted As Boolean = True

        'get the user's level of access
        Try
            sqlCommonAction.CommandText = "Select USER_CODE from " & myUtility.getExtension() & "SAT_USER where USER_KEY = " & _
            userID
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            If (Me.dsCommon.Tables.Contains("UserType")) Then
                Me.dsCommon.Tables("UserType").Rows.Clear()
            End If
            sqlCommonAdapter.Fill(Me.dsCommon, "UserType")
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try

        'update the user's level of access if less than 2
        If (Me.dsCommon.Tables("UserType").Rows(0).Item("USER_CODE") < 2) Then
            Try
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_USER set USER_CODE = 2 where " & _
                "USER_KEY = " & userID
                sqlCommonAction.ExecuteNonQuery()
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                Return False
            End Try
        End If

        'Try to insert a new record
        Try
            sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_USER_PERMISSION (USER_KEY, TEMPLATE_KEY, SURVEY_KEY, CREATOR_USER_KEY, " & _
            "CREATE_DATE, OWNER_IND, DELEGATE_IND, USER_IND) VALUES (" & userID & ", " & _
            CType(Session("seqTemplate"), Integer) & ", -1, " & CType(Session("UserID"), Integer) & _
            ", '" & Now & "', 0, 0, 1)"
            sqlCommonAction.ExecuteNonQuery()
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            blnUpdated = False
        End Try

        'if the insert failed try to do an update
        If (blnUpdated = False) Then
            Try
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_USER_PERMISSION SET USER_IND = 1" & _
                " WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " AND USER_KEY = " & userID
                sqlCommonAction.ExecuteNonQuery()
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                blnInserted = False
            End Try
        End If

        'if both fail then alert the user and stop after rolling back the transaction
        If Not (blnUpdated) And Not (blnInserted) Then
            sqlCommonAction.Transaction.Rollback()
            Return False
        End If
        Return True
    End Function
#End Region

#Region "Reset Code"
    'resets the page, will reset the page to its defaults
    Private Function doReset() As Boolean
        loadData("Reload")
        Return True
    End Function
#End Region
End Class
