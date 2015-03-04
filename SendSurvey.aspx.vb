Imports System.Collections.Specialized
Public Class SendSurvey
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SendSurvey))
        Me.dsCommon = New System.Data.DataSet
        Me.SqlGroups = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand
        Me.commonConn = New System.Data.SqlClient.SqlConnection
        Me.SqlUsers = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlDeleteCommand2 = New System.Data.SqlClient.SqlCommand
        Me.SqlInsertCommand2 = New System.Data.SqlClient.SqlCommand
        Me.SqlSelectCommand2 = New System.Data.SqlClient.SqlCommand
        Me.SqlUpdateCommand2 = New System.Data.SqlClient.SqlCommand
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'dsCommon
        '
        Me.dsCommon.DataSetName = "NewDataSet"
        Me.dsCommon.Locale = New System.Globalization.CultureInfo("en-US")
        '
        'SqlGroups
        '
        Me.SqlGroups.SelectCommand = Me.SqlSelectCommand1
        Me.SqlGroups.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_SURVEY_DIST_GROUP", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("GROUP_TITLE", "GROUP_TITLE"), New System.Data.Common.DataColumnMapping("USER_KEY", "USER_KEY"), New System.Data.Common.DataColumnMapping("SURVEY_KEY", "SURVEY_KEY"), New System.Data.Common.DataColumnMapping("SURVEY_GROUP_ID", "SURVEY_GROUP_ID")})})
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = resources.GetString("SqlSelectCommand1.CommandText")
        Me.SqlSelectCommand1.Connection = Me.commonConn
        '
        'commonConn
        '
        Me.commonConn.FireInfoMessageEventOnUserErrors = False
        '
        'SqlUsers
        '
        Me.SqlUsers.DeleteCommand = Me.SqlDeleteCommand2
        Me.SqlUsers.InsertCommand = Me.SqlInsertCommand2
        Me.SqlUsers.SelectCommand = Me.SqlSelectCommand2
        Me.SqlUsers.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_USER", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("user_key", "user_key"), New System.Data.Common.DataColumnMapping("user_code", "user_code"), New System.Data.Common.DataColumnMapping("last_name", "last_name"), New System.Data.Common.DataColumnMapping("first_name", "first_name"), New System.Data.Common.DataColumnMapping("middle_name", "middle_name"), New System.Data.Common.DataColumnMapping("suffix_name", "suffix_name"), New System.Data.Common.DataColumnMapping("authentication_code", "authentication_code"), New System.Data.Common.DataColumnMapping("primary_work_phone", "primary_work_phone"), New System.Data.Common.DataColumnMapping("hanford_id", "hanford_id"), New System.Data.Common.DataColumnMapping("email_address", "email_address"), New System.Data.Common.DataColumnMapping("create_date", "create_date"), New System.Data.Common.DataColumnMapping("used_date", "used_date"), New System.Data.Common.DataColumnMapping("creator_user_key", "creator_user_key"), New System.Data.Common.DataColumnMapping("active_ind", "active_ind"), New System.Data.Common.DataColumnMapping("training_ind", "training_ind"), New System.Data.Common.DataColumnMapping("company_abbrev", "company_abbrev"), New System.Data.Common.DataColumnMapping("no_email_ind", "no_email_ind")})})
        Me.SqlUsers.UpdateCommand = Me.SqlUpdateCommand2
        '
        'SqlDeleteCommand2
        '
        Me.SqlDeleteCommand2.CommandText = "DELETE FROM VARCONN.[sat_user] WHERE (([user_key] = @Original_user_key))" & _
            ""
        Me.SqlDeleteCommand2.Connection = Me.commonConn
        Me.SqlDeleteCommand2.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@Original_user_key", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "user_key", System.Data.DataRowVersion.Original, Nothing)})
        '
        'SqlInsertCommand2
        '
        Me.SqlInsertCommand2.CommandText = resources.GetString("SqlInsertCommand2.CommandText")
        Me.SqlInsertCommand2.Connection = Me.commonConn
        Me.SqlInsertCommand2.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@user_code", System.Data.SqlDbType.Int, 0, "user_code"), New System.Data.SqlClient.SqlParameter("@last_name", System.Data.SqlDbType.VarChar, 0, "last_name"), New System.Data.SqlClient.SqlParameter("@first_name", System.Data.SqlDbType.VarChar, 0, "first_name"), New System.Data.SqlClient.SqlParameter("@middle_name", System.Data.SqlDbType.VarChar, 0, "middle_name"), New System.Data.SqlClient.SqlParameter("@suffix_name", System.Data.SqlDbType.VarChar, 0, "suffix_name"), New System.Data.SqlClient.SqlParameter("@authentication_code", System.Data.SqlDbType.VarChar, 0, "authentication_code"), New System.Data.SqlClient.SqlParameter("@primary_work_phone", System.Data.SqlDbType.VarChar, 0, "primary_work_phone"), New System.Data.SqlClient.SqlParameter("@hanford_id", System.Data.SqlDbType.VarChar, 0, "hanford_id"), New System.Data.SqlClient.SqlParameter("@email_address", System.Data.SqlDbType.VarChar, 0, "email_address"), New System.Data.SqlClient.SqlParameter("@create_date", System.Data.SqlDbType.DateTime, 0, "create_date"), New System.Data.SqlClient.SqlParameter("@used_date", System.Data.SqlDbType.DateTime, 0, "used_date"), New System.Data.SqlClient.SqlParameter("@creator_user_key", System.Data.SqlDbType.Int, 0, "creator_user_key"), New System.Data.SqlClient.SqlParameter("@active_ind", System.Data.SqlDbType.Bit, 0, "active_ind"), New System.Data.SqlClient.SqlParameter("@training_ind", System.Data.SqlDbType.Bit, 0, "training_ind"), New System.Data.SqlClient.SqlParameter("@company_abbrev", System.Data.SqlDbType.VarChar, 0, "company_abbrev"), New System.Data.SqlClient.SqlParameter("@no_email_ind", System.Data.SqlDbType.Bit, 0, "no_email_ind")})
        '
        'SqlSelectCommand2
        '
        Me.SqlSelectCommand2.CommandText = resources.GetString("SqlSelectCommand2.CommandText")
        Me.SqlSelectCommand2.Connection = Me.commonConn
        '
        'SqlUpdateCommand2
        '
        Me.SqlUpdateCommand2.CommandText = resources.GetString("SqlUpdateCommand2.CommandText")
        Me.SqlUpdateCommand2.Connection = Me.commonConn
        Me.SqlUpdateCommand2.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@user_code", System.Data.SqlDbType.Int, 0, "user_code"), New System.Data.SqlClient.SqlParameter("@last_name", System.Data.SqlDbType.VarChar, 0, "last_name"), New System.Data.SqlClient.SqlParameter("@first_name", System.Data.SqlDbType.VarChar, 0, "first_name"), New System.Data.SqlClient.SqlParameter("@middle_name", System.Data.SqlDbType.VarChar, 0, "middle_name"), New System.Data.SqlClient.SqlParameter("@suffix_name", System.Data.SqlDbType.VarChar, 0, "suffix_name"), New System.Data.SqlClient.SqlParameter("@authentication_code", System.Data.SqlDbType.VarChar, 0, "authentication_code"), New System.Data.SqlClient.SqlParameter("@primary_work_phone", System.Data.SqlDbType.VarChar, 0, "primary_work_phone"), New System.Data.SqlClient.SqlParameter("@hanford_id", System.Data.SqlDbType.VarChar, 0, "hanford_id"), New System.Data.SqlClient.SqlParameter("@email_address", System.Data.SqlDbType.VarChar, 0, "email_address"), New System.Data.SqlClient.SqlParameter("@create_date", System.Data.SqlDbType.DateTime, 0, "create_date"), New System.Data.SqlClient.SqlParameter("@used_date", System.Data.SqlDbType.DateTime, 0, "used_date"), New System.Data.SqlClient.SqlParameter("@creator_user_key", System.Data.SqlDbType.Int, 0, "creator_user_key"), New System.Data.SqlClient.SqlParameter("@active_ind", System.Data.SqlDbType.Bit, 0, "active_ind"), New System.Data.SqlClient.SqlParameter("@training_ind", System.Data.SqlDbType.Bit, 0, "training_ind"), New System.Data.SqlClient.SqlParameter("@company_abbrev", System.Data.SqlDbType.VarChar, 0, "company_abbrev"), New System.Data.SqlClient.SqlParameter("@no_email_ind", System.Data.SqlDbType.Bit, 0, "no_email_ind"), New System.Data.SqlClient.SqlParameter("@Original_user_key", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "user_key", System.Data.DataRowVersion.Original, Nothing)})
        '
        'commonConn
        '
        Me.commonConn.FireInfoMessageEventOnUserErrors = False
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected navButtons As Collection
    Protected requestItems As Array
    Protected myNavBar As Navigation = New Navigation
    Protected myUtility As Utility = New Utility
    Protected mailList As String
    Protected sqlCommonAdapter As New System.Data.SqlClient.SqlDataAdapter
    Protected sqlCommonAction As New System.Data.SqlClient.SqlCommand
    Protected WithEvents imgTitle As System.Web.UI.WebControls.Image
    Protected WithEvents hlpServerSwitch As System.Web.UI.WebControls.Image
    Protected WithEvents chkServerSwitch As System.Web.UI.WebControls.CheckBox
    Protected WithEvents hlpSelectSurvey As System.Web.UI.WebControls.Image
    Protected WithEvents ddlSurveyList As System.Web.UI.WebControls.DropDownList
    Protected WithEvents chkbxFollowUp As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblSelectSurvey As System.Web.UI.WebControls.Label
    Protected WithEvents lblFollowUp As System.Web.UI.WebControls.Label
    Protected WithEvents wdcFollowUp As Infragistics.WebUI.WebSchedule.WebDateChooser
    Protected WithEvents hlpSubject As System.Web.UI.WebControls.Image
    Protected WithEvents lblSubject As System.Web.UI.WebControls.Label
    Protected WithEvents txtSubject As System.Web.UI.WebControls.TextBox
    Protected WithEvents hlpMessage As System.Web.UI.WebControls.Image
    Protected WithEvents lblMessage As System.Web.UI.WebControls.Label
    Protected WithEvents txtMessage As System.Web.UI.WebControls.TextBox
    Protected WithEvents lblSurveyGroups As System.Web.UI.WebControls.Label
    Protected WithEvents hlpGroups As System.Web.UI.WebControls.Image
    Protected WithEvents dgSurveyGroups As System.Web.UI.WebControls.DataGrid
    Protected WithEvents lblSurveyGroupsNone As System.Web.UI.WebControls.Label
    Protected WithEvents lblSurveyRespondents As System.Web.UI.WebControls.Label
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
    Protected WithEvents hlpSurveyRespondents As System.Web.UI.WebControls.Image
    Protected WithEvents dgSurveyRespondents As System.Web.UI.WebControls.DataGrid
    Protected WithEvents lblSurveyRespondentsNone As System.Web.UI.WebControls.Label
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
    Protected WithEvents btnBuild1 As System.Web.UI.WebControls.Button
    Protected WithEvents btnBuild2 As System.Web.UI.WebControls.Button
    Protected WithEvents btnSend1 As System.Web.UI.WebControls.Button
    Protected WithEvents btnSend2 As System.Web.UI.WebControls.Button
    Protected WithEvents btnClear1 As System.Web.UI.WebControls.Button
    Protected WithEvents btnClear2 As System.Web.UI.WebControls.Button
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents dsCommon As System.Data.DataSet
    Protected WithEvents SqlGroups As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlUsers As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents SqlDeleteCommand2 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlInsertCommand2 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlSelectCommand2 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlUpdateCommand2 As System.Data.SqlClient.SqlCommand
    Protected WithEvents hlpchkbxFollowUp As System.Web.UI.WebControls.Image
    Protected WithEvents hlpFollowUp As System.Web.UI.WebControls.Image
    Protected WithEvents lblSelectedUsers As System.Web.UI.WebControls.Label
    Protected WithEvents Image1 As System.Web.UI.WebControls.Image
    Protected WithEvents txtRecipients As System.Web.UI.WebControls.TextBox
    Protected strRecipientFilter As String
    Protected strRecipientsList As String
    Protected WithEvents hlpchkbxReminder As System.Web.UI.WebControls.Image
    Protected WithEvents chkbxReminder As System.Web.UI.WebControls.CheckBox
    Protected WithEvents commonConn As System.Data.SqlClient.SqlConnection
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
        InitializeComponent()
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
        Else
            Session("CurrentPage") = "./SendSurvey.aspx"
            Session("pageSel") = "SendSurvey"
        End If

        'if the user's password has expired send them to the password reset page
        If (Session("isExpired") = "True") Then
            Session("CurrentPage") = "./Password.aspx"
            Session("pageSel") = "Password"
            redirect(Session("CurrentPage"))
        End If

        If (Session("seqTemplate") Is Nothing) Then
            Session("seqTemplate") = -1
        End If

        If (Session("seqSurvey") Is Nothing) Then
            Session("seqSurvey") = -1
        End If

        If (Session("isFollowUp") Is Nothing) Then
            Session("isFollowUp") = False
        End If

        myUtility.getConnection(Me.commonConn)

        'do initial set-up if we're coming here for the first time
        If Not (IsPostBack) Then
            Session("seqTemplate") = -1
            Session("seqSurvey") = -1

            'get anything we can off the request line if we're not coming back from processing an e-mail address
            If (Session("ProcessEmail") <> "True") Then
                myUtility.getRequest(Page.Request.Url.Query())
            End If

            Session("Alpha") = "A"
            Session("RecipientsSelected") = "NONE"

            'load the page data
            If Not (loadData()) Then
                alert("Failed to load the recipient and group data for this survey.")
            End If

            'Populate the controls on the page
            setControls()
            Me.btnBuild1.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
            Me.btnBuild2.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
            Me.btnClear1.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
            Me.btnClear2.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
            Me.btnSend1.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
            Me.btnSend2.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
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
#End Region

#Region "Data Load"
    'Loads data into the form
    Private Function loadData(Optional ByVal override As Boolean = False) As Boolean
        Trace.Warn("Loading Data")
        Try
            'Fill the surveys list
            If (loadSurveys()) Then
                If (Session("seqSurvey") <> -1) Then
                    'Fill the Groups List
                    If (loadGroups(override)) Then
                        'Fill the Users List
                        If (loadRespondents()) Then
                            Return True
                        Else
                            Return False
                        End If
                    Else
                        Return False
                    End If
                Else
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    'loads the survey list
    Private Function loadSurveys() As Boolean
        Trace.Warn("Loading Surveys")
        Try
            'try to erase anything if it exists in the table so we don't end up with duplicate data
            If (Me.dsCommon.Tables.Contains("Surveys")) Then
                Me.dsCommon.Tables("Surveys").Rows.Clear()
            End If

            'set up the query
            Me.sqlCommonAction = New SqlClient.SqlCommand
            Me.sqlCommonAdapter = New SqlClient.SqlDataAdapter
            If (Session("UserType") = 4) Then
                Me.sqlCommonAction.CommandText = "select * from (select fs.SURVEY_KEY, fs.SURVEY_COMMENT, fs.TEMPLATE_KEY from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY fs, " & _
                myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE fss, " & myUtility.getExtension() & "SAT_TEMPLATE ft where fs.TEMPLATE_KEY = fss.TEMPLATE_KEY and " & _
                "fss.TEMPLATE_KEY = ft.TEMPLATE_KEY and fss.CURRENT_IND = 1 and fss.SIGNED_IND = 1 and ft.ACTIVE_IND = 1 " & _
                "and fs.ACTIVE_IND = 1 and ft.CHANGE_IND = 0 and not exists (select * from " & myUtility.getExtension() & "SAT_SURVEY_TAG fts where fts.SURVEY_KEY = fs.SURVEY_KEY " & _
                "and rtrim(ltrim(SURVEY_TAG_ANSWER)) = '') " & _
                " group by fs.TEMPLATE_KEY, fs.SURVEY_KEY, fs.SURVEY_COMMENT " & _
                "having count(*) = 3) a order by SURVEY_COMMENT"
            Else
                Me.sqlCommonAction.CommandText = "select * from (select fs.SURVEY_KEY, fs.SURVEY_COMMENT, fs.TEMPLATE_KEY from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY fs, " & _
                myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE fss, " & myUtility.getExtension() & "SAT_TEMPLATE ft, " & myUtility.getProdExtension() & "SAT_USER_PERMISSION fp where " & _
                "fs.TEMPLATE_KEY = fss.TEMPLATE_KEY and fs.SURVEY_KEY = fp.SURVEY_KEY and fp.USER_KEY = " & Session("UserID") & _
                " and (fp.OWNER_IND = 1 or fp.DELEGATE_IND = 1) and fss.TEMPLATE_KEY = ft.TEMPLATE_KEY and fss.CURRENT_IND = 1 " & _
                "and fss.SIGNED_IND = 1 and ft.ACTIVE_IND = 1 " & _
                " and fs.ACTIVE_IND = 1 and ft.CHANGE_IND = 0 and not exists (select * from " & myUtility.getExtension() & "SAT_SURVEY_TAG fts where fts.SURVEY_KEY = fs.SURVEY_KEY " & _
                "and rtrim(ltrim(SURVEY_TAG_ANSWER)) = '') group by fs.TEMPLATE_KEY, fs.SURVEY_KEY, fs.SURVEY_COMMENT " & _
                "having count(*) = 3) a order by SURVEY_COMMENT"
            End If
            Trace.Warn(Me.sqlCommonAction.CommandText)
            Me.sqlCommonAdapter.SelectCommand = Me.sqlCommonAction
            Me.sqlCommonAction.Connection = Me.commonConn

            'get the data
            Me.sqlCommonAdapter.Fill(Me.dsCommon, "Surveys")

            'bind the data
            Me.ddlSurveyList.DataTextField = "SURVEY_COMMENT"
            Me.ddlSurveyList.DataValueField = "SURVEY_KEY"
            Me.ddlSurveyList.DataSource = Me.dsCommon.Tables("Surveys")
            Me.ddlSurveyList.DataBind()
            Dim li As System.Web.UI.WebControls.ListItem = New System.Web.UI.WebControls.ListItem
            li.Text = "-- Select a Survey --"
            li.Value = "-1"
            Me.ddlSurveyList.Items.Insert(0, li)

            'determine if any surveys were returned
            If (Me.dsCommon.Tables("Surveys").Rows.Count() > 0) Then
                Session("SurveysExist") = "Yes"
            Else
                Session("SurveysExist") = "No"
            End If

            'Check to see if this is our first time on the page or not
            If (Session("seqSurvey") <> -1 And Not Me.ddlSurveyList.Items.FindByValue(Session("seqSurvey")) Is Nothing) Then
                Me.ddlSurveyList.Items.FindByValue(Session("seqSurvey")).Selected = True
            Else
                Me.ddlSurveyList.SelectedIndex = 0
            End If

            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Session("SurveysExist") = "No"
            Return False
        End Try
    End Function

    'loads the groups list
    Private Function loadGroups(Optional ByVal override As Boolean = False) As Boolean
        Trace.Warn("Loading Groups")
        Try
            If Not (override) Then
                'try to erase anything if it exists in the table so we don't end up with duplicate data
                If (Me.dsCommon.Tables.Contains("Groups")) Then
                    Me.dsCommon.Tables("Groups").Rows.Clear()
                End If

                'set up the query
                Me.sqlCommonAction = New SqlClient.SqlCommand
                Me.sqlCommonAdapter = New SqlClient.SqlDataAdapter
                Me.sqlCommonAction.CommandText = "Select SSDG.SURVEY_KEY, SSDG.SURVEY_GROUP_ID, SSDG.GROUP_TITLE " & _
                        "From " & myUtility.getExtension() & "SAT_SURVEY_DIST_GROUP SSDG where SSDG.SURVEY_KEY = " & Session("seqSurvey") & _
                        " Order by SSDG.GROUP_TITLE"
                Me.sqlCommonAdapter.SelectCommand = Me.sqlCommonAction
                Me.sqlCommonAction.Connection = Me.commonConn

                'get the data
                Me.sqlCommonAdapter.Fill(Me.dsCommon, "Groups")

                'bind the data
                Me.dgSurveyGroups.DataSource = Me.dsCommon.Tables("Groups")
                Me.dgSurveyGroups.DataBind()

                'determine if any groups were returned
                If (Me.dsCommon.Tables("Groups").Rows.Count() > 0) Then
                    Session("GroupsRows") = "Some"
                Else
                    Session("GroupsRows") = "None"
                End If

                Return True
            Else
                Return True
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Session("GroupsRows") = "None"
            Return False
        End Try
    End Function

    'loads the users
    Private Function loadRespondents() As Boolean
        Trace.Warn("Loading Respondents")
        Try
            'try to erase anything if it exists in the table so we don't end up with duplicate data
            If (Me.dsCommon.Tables.Contains("Users")) Then
                Me.dsCommon.Tables("Users").Rows.Clear()
            End If

            'set up the query
            Me.sqlCommonAction.CommandText = "SELECT fu.USER_KEY, fu.USER_CODE, fu.LAST_NAME + ', '" & _
            "+ fu.FIRST_NAME + ' ' + fu.MIDDLE_NAME + ' ' + fu.SUFFIX_NAME AS Name, fu.HANFORD_ID, lower(fu.EMAIL_ADDRESS) as EMAIL_ADDRESS, " & _
            "convert(varchar, fr.CREATE_DATE, 101) as CREATE_DATE, case when convert(varchar, fr.LAST_COMPLETION_DATE, 101) " & _
            "= '01/01/1970' then NULL else convert(varchar, fr.LAST_COMPLETION_DATE, 101) end as LAST_COMPLETION_DATE, " & _
            "case when convert(varchar, fr.REMIND_DATE, 101) = '01/01/1970' then NULL else " & _
            "convert(varchar, fr.REMIND_DATE, 101) end as REMIND_DATE FROM " & myUtility.getExtension() & "SAT_USER fu left outer join " & _
            myUtility.getExtension() & "SAT_RESPONSE fr on fu.USER_KEY = fr.RESPONDENT_USER_KEY and fr.SURVEY_KEY = " & Session("seqSurvey") & _
            "WHERE upper(fu.LAST_NAME) like '" & Session("Alpha") & "%' and fu.EMAIL_ADDRESS <> '' ORDER BY Name"
            Me.sqlCommonAdapter.SelectCommand = Me.sqlCommonAction
            Me.sqlCommonAction.Connection = Me.commonConn

            'get the data
            Me.sqlCommonAdapter.Fill(Me.dsCommon, "Users")

            'bind the data
            Me.dgSurveyRespondents.DataSource = Me.dsCommon.Tables("Users")
            Me.dgSurveyRespondents.DataBind()

            'determine if any users were returned
            If (Me.dsCommon.Tables("Users").Rows.Count() = 0) Then
                Session("RespondentRows") = "None"
            Else
                Session("RespondentRows") = "Some"
            End If
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Session("RespondentRows") = "None"
            Return False
        End Try
    End Function
#End Region

#Region "Filters"
    'creates a recipient filter string based on the master list of recipients
    Private Function recipientFilter(Optional ByVal isSendSurvey As Boolean = False) As Boolean
        Trace.Warn("Setting Question Filter")
        Try
            Me.strRecipientFilter = ""
            Dim id As String

            'iterate through each id in the master list
            For Each id In CType(Session("MasterIDList"), ArrayList)
                Me.strRecipientFilter &= id & ","
            Next

            'if sending the survey get the current user's information too
            If (isSendSurvey) Then
                Me.strRecipientFilter &= Session("UserID") & ","
            End If

            'final string processing for the in clause
            Me.strRecipientFilter = Me.strRecipientFilter.Insert(0, " USER_KEY in (")
            Me.strRecipientFilter = Me.strRecipientFilter.Remove(Me.strRecipientFilter.Length - 1, 1)
            Me.strRecipientFilter = Me.strRecipientFilter.Insert(Me.strRecipientFilter.Length, ")")

            Return True
        Catch ex As Exception
            Session("Alert") = "Data filter creation failure.  " & ex.ToString
            Return False
        End Try
    End Function
#End Region

#Region "Control Settings"
    'populates the controls with the selected group information
    Private Sub setControls(Optional ByVal override As Boolean = False)
        Trace.Warn("Setting Controls")

        'reset the checkboxes on the grids
        If (override) Then
            myUtility.checkBoxes(Me.dgSurveyGroups, True)
            myUtility.checkBoxes(Me.dgSurveyRespondents, True)
            Session("RecipientIDList") = ""

            'reset the text on the form
            Me.txtMessage.Text = ""
            Me.txtSubject.Text = ""

            'determine what to do with the closing date
            If (Session("isFollowUp") = True) Then
                Session("isFollowUp") = False
                Me.chkbxFollowUp.Checked = False
            End If

            Me.chkbxReminder.Checked = False

            'reset the session lists
            Session("CheckedGroupsList") = New ArrayList
            Session("MasterIDList") = New ArrayList
            Session("RecipientsSelected") = "NONE"
        End If

        'sets the textbox to the appropriate text if there are no recipients
        If (Session("RecipientsSelected") = "NONE") Then
            Me.txtRecipients.Text = "There are currently no recipients.  Please select some recipients " & _
                "below or send the survey to receive an e-mail with text that will allow you to send the " & _
                "survey for anonymous results. (Do not try to enter text into this text box, the changes " & _
                "will not be saved)"
        End If

        'set the checked users based on the master list
        If Not (checkUsers()) Then
            alert("Failed to check the users you have previously selected that are in this section.")
        End If
    End Sub
#End Region

#Region "Checks"
    'handles displaying and setting the closing date
    Private Sub chkbxFollowUp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkbxFollowUp.CheckedChanged
        Trace.Warn("Follow-Up Selected")

        'determine what to do with the closing date
        If (Session("isFollowUp") = False) Then
            Session("isFollowUp") = True
            Me.wdcFollowUp.Value = Now.AddDays(14)
        Else
            Session("isFollowUp") = False
        End If

        'change control text
        setControls()
    End Sub
#End Region

#Region "Grid"
    'Weigh station for processing the checking and unchecking of checkboxes in the grids 
    Public Sub commandBtnClick(ByVal sender As Object, ByVal e As CommandEventArgs)
        Trace.Warn("Header Button Click")
        Select Case e.CommandName
            Case "SelectGroup"
                myUtility.checkBoxes(Me.dgSurveyGroups)
            Case "SelectUsers"
                myUtility.checkBoxes(Me.dgSurveyRespondents)
        End Select
    End Sub

    'resets the user grid to the value passed in from the button
    Public Sub commandGridUserList(ByVal sender As Object, ByVal e As CommandEventArgs)
        Trace.Warn("Grid Alphabet Button Click")
        Session("Alpha") = e.CommandName

        Dim failure As Boolean = False

        'do a build
        If (doBuild()) Then
            alert("Failed to save your recent changes during transition.")
        End If

        'load the page data
        If Not (loadData(True)) Then
            alert("Failed to load the group information for this survey.")
        End If

        'reset the page controls
        setControls()
    End Sub

    'checks the users in the recipient grid based on selected users and groups
    Private Function checkUsers() As Boolean
        Trace.Warn("Checking the Recipients")
        Try
            Dim Index As Integer = 0
            While (Index < Me.dgSurveyRespondents.Items.Count())
                Dim tempChkBx As CheckBox = CType(Me.dgSurveyRespondents.Items(Index).Controls(0).Controls(1), CheckBox)
                Dim tempString As String = CType(Me.dgSurveyRespondents.Items(Index).Controls(1), TableCell).Text
                Dim tempList As ArrayList = CType(Session("MasterIDList"), ArrayList)

                'check the recipients based on whether or not the user id is in the master list
                If (tempList.Contains(tempString)) Then
                    CType(Me.dgSurveyRespondents.Items(Index).Controls(0).Controls(1), CheckBox).Checked = True
                End If
                Index += 1
            End While

            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function
#End Region

#Region "Page Action"
    'handles the selection of surveys
    Private Sub ddlSurveyList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddlSurveyList.SelectedIndexChanged
        Session("JSAction") = ""

        'Store the recently selected survey sequence number
        Session("seqSurvey") = Me.ddlSurveyList.SelectedItem.Value()

        'loads the appropriate data
        loadData()

        'change control text
        setControls(True)
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnBuild_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBuild1.Click, btnBuild2.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'check for possible malicious code entries
        If Not (myUtility.goodInput(Me.txtMessage, True) And myUtility.goodInput(Me.txtSubject, True)) Then
            blnContinue = False
            alert("Possible malicious code entry(s).  Build aborted.")
        End If

        If (blnContinue) Then
            failure = doBuild()

            'reload the data
            If Not (loadData(True)) Then
                blnContinue = False
                alert("Failed to load the recipient and group data for this survey.")
            End If

            If (blnContinue) Then
                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnSend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSend1.Click, btnSend2.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'check for possible malicious code entries
        If Not (myUtility.goodInput(Me.txtMessage, True) And myUtility.goodInput(Me.txtSubject, True)) Then
            blnContinue = False
            alert("Possible malicious code entry(s).  Send aborted.")
        End If

        If (blnContinue) Then
            failure = doSendSurvey()

            'reload the data
            If Not (loadData(True)) Then
                blnContinue = False
                alert("Failed to load the recipient and group data for this survey.")
            End If

            If (blnContinue) Then
                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear1.Click, btnClear2.Click
        Session("JSAction") = ""

        Dim failure As Boolean

        'do an insert/update/delete/reset based on the page option selected
        failure = doReset()

        'reset the page controls
        If Not (failure) Then
            setControls(True)
        End If
    End Sub
#End Region

#Region "Build Code"
    'builds the list of recipients - stores a list of user ids in memory
    Private Function doBuild() As Boolean
        Dim tempMasterList As ArrayList = CType(Session("MasterIDList"), ArrayList)
        Dim checkedGroupsList As ArrayList = CType(Session("CheckedGroupsList"), ArrayList)

        Try
            'process the groups section - session of selected users in groups
            If (processRecipients()) Then
                'process the recipient section - session of selected users in recipient
                If (processGroups()) Then
                    'generate the recipient list
                    If Not (buildRecipientList()) Then
                        alert("failed to build the recipient list.")
                        Return True
                    End If
                Else
                    alert("failed to build the recipient list.  No changes were made.")
                    Session("MasterIDList") = tempMasterList
                    Session("CheckedGroupsList") = checkedGroupsList
                    Return True
                End If
            Else
                alert("failed to build the recipient list.  No changes were made.")
                Session("MasterIDList") = tempMasterList
                Session("CheckedGroupsList") = checkedGroupsList
                Return True
            End If
            Return False
        Catch ex As Exception
            alert("failed to build the recipient list.  No changes were made.")
            Session("MasterIDList") = tempMasterList
            Session("CheckedGroupsList") = checkedGroupsList
            Return True
        End Try
    End Function
#End Region

#Region "Build/Send Processing"
    'process the recipient section - session of selected users in recipient
    Private Function processRecipients() As Boolean
        Trace.Warn("Processing the Recipients")
        Try
            Dim Index As Integer = 0
            While (Index < Me.dgSurveyRespondents.Items.Count())
                Dim tempChkBx As CheckBox = CType(Me.dgSurveyRespondents.Items(Index).Controls(0).Controls(1), CheckBox)
                Dim tempString As String = CType(Me.dgSurveyRespondents.Items(Index).Controls(1), TableCell).Text

                'add or remove the id based on whether or not the user is checked
                If (tempChkBx.Checked = True) Then
                    If Not (addToLists(tempString)) Then
                        Return False
                    End If
                Else
                    If Not (subFromLists(tempString)) Then
                        Return False
                    End If
                End If
                Index += 1
            End While

            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'process the groups section - session of selected users in groups
    Private Function processGroups() As Boolean
        Trace.Warn("Procesing the Groups")
        Try
            Dim tempArr As ArrayList = CType(Session("CheckedGroupsList"), ArrayList)
            Dim Index As Integer = 0
            While (Index < Me.dgSurveyGroups.Items.Count())
                Dim tempChkBx As CheckBox = CType(Me.dgSurveyGroups.Items(Index).Controls(0).Controls(1), CheckBox)
                Dim tempID As Integer = CType(Me.dgSurveyGroups.Items(Index).Controls(2), TableCell).Text

                If (tempChkBx.Checked = True) Then
                    'if the group is not already in the list then add the ids
                    If Not (tempArr.Contains(tempID)) Then
                        If (addToLists(tempID, True)) Then
                            tempArr.Add(tempID)
                        Else
                            Return False
                        End If
                    End If
                Else
                    'if it was checked then remove it from the list and remove the ids
                    If (tempArr.Contains(tempID)) Then
                        If (subFromLists(tempID, True)) Then
                            tempArr.Remove(tempID)
                        Else
                            Return False
                        End If
                    End If
                End If
                Index += 1
            End While

            'save the changes to the checked group list
            Session("CheckedGroupsList") = tempArr

            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'appends a list of user id's
    Private Function addToLists(ByVal userIDs As String, Optional ByVal isGroupProcess As Boolean = False) As Boolean
        Try
            If (isGroupProcess) Then
                'set up the query
                Me.sqlCommonAction = New SqlClient.SqlCommand
                Me.sqlCommonAdapter = New SqlClient.SqlDataAdapter
                Me.sqlCommonAction.CommandText = "Select USER_KEY From " & myUtility.getExtension() & _
                        "SAT_SURVEY_DIST_GRP_LIST where SURVEY_KEY = " & Session("seqSurvey") & _
                        " Order by USER_KEY"
                Me.sqlCommonAdapter.SelectCommand = Me.sqlCommonAction
                Me.sqlCommonAction.Connection = Me.commonConn

                'get the data
                If (Me.dsCommon.Tables.Contains("GroupsList")) Then
                    Me.dsCommon.Tables.Remove("GroupsList")
                End If
                Me.sqlCommonAdapter.Fill(Me.dsCommon, "GroupList")

                'store the ids into an array list
                Dim tempList As ArrayList = CType(Session("MasterIDList"), ArrayList)

                'add each id to the array list object in session
                Dim row As DataRow
                For Each row In Me.dsCommon.Tables("GroupList").Rows
                    If (Not tempList.Contains(Trim(row.Item("USER_KEY")))) Then
                        tempList.Add(Trim(row.Item("USER_KEY")))
                    End If
                Next

                Session("MasterIDList") = tempList
                Return True
            Else
                'store the ids into an array list
                Dim tempList As ArrayList = CType(Session("MasterIDList"), ArrayList)

                'add each id to the array list object in session
                If (Not tempList.Contains(Trim(userIDs))) Then
                    tempList.Add(Trim(userIDs))
                End If

                Session("MasterIDList") = tempList
                Return True
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'removes ids from a list of user id's
    Private Function subFromLists(ByVal userIDs As String, Optional ByVal isGroupProcess As Boolean = False) As Boolean
        Try
            If (isGroupProcess) Then
                'set up the query
                Me.sqlCommonAction = New SqlClient.SqlCommand
                Me.sqlCommonAdapter = New SqlClient.SqlDataAdapter
                Me.sqlCommonAction.CommandText = "Select USER_KEY From " & myUtility.getExtension() & _
                        "SAT_SURVEY_DIST_GRP_LIST where SURVEY_KEY = " & Session("seqSurvey") & _
                        " Order by USER_KEY"
                Me.sqlCommonAdapter.SelectCommand = Me.sqlCommonAction
                Me.sqlCommonAction.Connection = Me.commonConn

                'get the data
                If (Me.dsCommon.Tables.Contains("GroupsList")) Then
                    Me.dsCommon.Tables.Remove("GroupsList")
                End If
                Me.sqlCommonAdapter.Fill(Me.dsCommon, "GroupList")

                'split the id's out of the string
                Dim tempList As ArrayList = CType(Session("MasterIDList"), ArrayList)

                'remove each id ini the array list object in session
                Dim row As DataRow
                For Each row In Me.dsCommon.Tables("GroupList").Rows
                    If (tempList.Contains(Trim(row.Item("USER_KEY")))) Then
                        tempList.Remove(Trim(row.Item("USER_KEY")))
                    End If
                Next

                Session("MasterIDList") = tempList
                Return True
            Else
                'split the id's out of the string
                Dim tempList As ArrayList = CType(Session("MasterIDList"), ArrayList)

                'remove each id ini the array list object in session
                If (tempList.Contains(Trim(userIDs))) Then
                    tempList.Remove(Trim(userIDs))
                End If

                Session("MasterIDList") = tempList
                Return True
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'generate the recipient list
    Private Function buildRecipientList(Optional ByVal isSendSurvey As Boolean = False) As Boolean
        Try
            'set the variable for the list based on the availability of selected items
            If (CType(Session("MasterIDList"), ArrayList).Count > 0) Then
                Session("RecipientsSelected") = "SOME"

                'try to erase anything if it exists in the table so we don't end up with duplicate data
                If (Me.dsCommon.Tables.Contains("MasterUserList")) Then
                    Me.dsCommon.Tables("MasterUserList").Rows.Clear()
                End If

                'build the recipient filter
                If Not (recipientFilter(isSendSurvey)) Then
                    Return False
                End If

                'set up the query
                Me.sqlCommonAction = New SqlClient.SqlCommand
                Me.sqlCommonAdapter = New SqlClient.SqlDataAdapter
                Me.sqlCommonAction.CommandText = "SELECT LAST_NAME + ', ' + FIRST_NAME + ' ' + MIDDLE_NAME" & _
                    " + ' ' + SUFFIX_NAME AS Name, EMAIL_ADDRESS, AUTHENTICATION_CODE, USER_KEY " & _
                    "FROM " & myUtility.getExtension() & "SAT_USER WHERE " & Me.strRecipientFilter & " ORDER BY Name"
                Me.sqlCommonAdapter.SelectCommand = Me.sqlCommonAction
                Me.sqlCommonAction.Connection = Me.commonConn

                'get the data
                Me.sqlCommonAdapter.Fill(Me.dsCommon, "MasterUserList")

                'process data
                If (isSendSurvey And Not CType(Session("MasterIDList"), ArrayList).Contains(Session("UserID"))) Then
                    'process the data if it is available
                    Me.txtRecipients.Text = ""
                    If (Me.dsCommon.Tables("MasterUserList").Rows.Count > 0) Then
                        Dim row As DataRow
                        For Each row In Me.dsCommon.Tables("MasterUserList").Rows
                            If (row.Item("USER_KEY") <> Session("UserID")) Then
                                Me.txtRecipients.Text &= row.Item("Name") & " - " & row.Item("EMAIL_ADDRESS") & vbCrLf
                                Me.strRecipientsList &= row.Item("Name") & " - " & row.Item("EMAIL_ADDRESS") & "<br/>"
                            End If
                        Next
                    Else
                        Return False
                    End If
                Else
                    'process the data if it is available
                    Me.txtRecipients.Text = ""
                    If (Me.dsCommon.Tables("MasterUserList").Rows.Count > 0) Then
                        Dim row As DataRow
                        For Each row In Me.dsCommon.Tables("MasterUserList").Rows
                            Me.txtRecipients.Text &= row.Item("Name") & " - " & row.Item("EMAIL_ADDRESS") & vbCrLf
                        Next
                    Else
                        Return False
                    End If
                End If

                Return True
            Else
                Session("RecipientsSelected") = "NONE"

                'if sending a survey at least get the current user's information
                If (isSendSurvey) Then
                    'try to erase anything if it exists in the table so we don't end up with duplicate data
                    If (Me.dsCommon.Tables.Contains("MasterUserList")) Then
                        Me.dsCommon.Tables("MasterUserList").Rows.Clear()
                    End If

                    'build the recipient filter
                    If Not (recipientFilter(True)) Then
                        Return False
                    End If

                    'set up the query
                    Me.sqlCommonAction = New SqlClient.SqlCommand
                    Me.sqlCommonAdapter = New SqlClient.SqlDataAdapter
                    Me.sqlCommonAction.CommandText = "SELECT LAST_NAME + ', ' + FIRST_NAME + ' ' + MIDDLE_NAME" & _
                        " + ' ' + SUFFIX_NAME AS Name, EMAIL_ADDRESS, AUTHENTICATION_CODE, USER_KEY " & _
                        "FROM " & myUtility.getExtension() & "SAT_USER WHERE " & Me.strRecipientFilter & " ORDER BY Name"
                    Me.sqlCommonAdapter.SelectCommand = Me.sqlCommonAction
                    Me.sqlCommonAction.Connection = Me.commonConn

                    'get the data
                    Me.sqlCommonAdapter.Fill(Me.dsCommon, "MasterUserList")
                End If
                Return True
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'sends the e-mails out based on location e-mails may or may not be sent to recipients
    Private Function sendSurveys() As Boolean
        'keep track of everyone who was sent the e-mail
        Dim SentList As String = ""

        Try
            Dim strMessage As String = ""
            Dim strSubject As String = ""
            Dim strTo As String = ""
            Dim strFrom As String = ""

            'send only 1 e-mail if on dev to the user sending the survey
            'otherwise deliver all of the e-mails.
            If (Session("Machine") = "Development") Then
                'set the subject for the e-mail
                strSubject = Me.txtSubject.Text

                'find the user's authenticator
                Dim row As DataRow
                Dim strAuthenticator As String = ""
                For Each row In Me.dsCommon.Tables("MasterUserList").Rows
                    If (row.Item("USER_KEY") = Session("UserID")) Then
                        strAuthenticator = row.Item("AUTHENTICATION_CODE")
                        strTo = row.Item("EMAIL_ADDRESS")
                        strFrom = "survey.administrator@pnl.gov"
                        Exit For
                    End If
                Next

                'Get the connection
                Me.sqlCommonAction.Connection = Me.commonConn

                'if doing a follow-up also modify the fempSurvey closing date
                If (CType(Session("isFollowUp"), Boolean)) Then
                    Me.sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY set SURVEY_CLOSING_DATE = '" & _
                    Me.wdcFollowUp.Text & "' where SURVEY_KEY = " & Session("seqSurvey")
                    Me.sqlCommonAction.ExecuteNonQuery()
                End If

                'only update fempResponse if the recipient is not also the sender
                Dim needsReminder As Boolean = False
                Dim skip As Boolean
                For Each row In Me.dsCommon.Tables("MasterUserList").Rows
                    'get the authenticator and destination e-mail address
                    strAuthenticator = row.Item("AUTHENTICATION_CODE")
                    strTo = row.Item("EMAIL_ADDRESS")
                    strMessage = ""
                    skip = False

                    If (row.Item("USER_KEY") <> Session("UserID")) Then
                        'determine if the user has a response record already created.
                        Me.sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_RESPONSE where RESPONDENT_USER_KEY = " & _
                        row.Item("USER_KEY") & " and SURVEY_KEY = " & Session("seqSurvey") & _
                        " order by SUBMISSION_COUNT Desc"
                        If (Me.dsCommon.Tables.Contains("ExistingResponse")) Then
                            Me.dsCommon.Tables("ExistingResponse").Rows.Clear()
                        End If
                        Me.sqlCommonAdapter.Fill(Me.dsCommon, "ExistingResponse")

                        'update if the response already exists or create a new record if it does not
                        If (Me.dsCommon.Tables("ExistingResponse").Rows.Count > 0) Then
                            If (Me.dsCommon.Tables("ExistingResponse").Rows(0).Item("LAST_COMPLETION_DATE") = "01/01/1970") Then
                                needsReminder = True
                            End If

                            'if doing a follow-up adjust created/remind/date-used
                            If (CType(Session("isFollowUp"), Boolean)) Then
                                If (needsReminder) Then
                                    Me.sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_RESPONSE set REMIND_DATE = '" & _
                                    Now & "' where RESPONDENT_USER_KEY = " & _
                                    row.Item("USER_KEY") & " and SURVEY_KEY = " & Session("seqSurvey")
                                Else
                                    Me.sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_RESPONSE set CREATE_DATE = '" & _
                                    Now & "', REMIND_DATE = '1/1/1970', LAST_USED_DATE = '1/1/1970'" & _
                                    ", SUBMISSION_COUNT = " & _
                                    Me.dsCommon.Tables("ExistingResponse").Rows(0).Item("SUBMISSION_COUNT") + 1 & _
                                    "where RESPONDENT_USER_KEY = " & row.Item("USER_KEY") & _
                                    " and SURVEY_KEY = " & Session("seqSurvey")
                                End If
                            ElseIf (Me.chkbxReminder.Checked) Then
                                If (needsReminder) Then
                                    Me.sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_RESPONSE set REMIND_DATE = '" & _
                                    Now & "' where RESPONDENT_USER_KEY = " & _
                                    row.Item("USER_KEY") & " and SURVEY_KEY = " & Session("seqSurvey")
                                Else
                                    skip = True
                                End If
                            Else
                                If (needsReminder) Then
                                    Me.sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_RESPONSE set REMIND_DATE = '" & _
                                    Now & "' where RESPONDENT_USER_KEY = " & _
                                    row.Item("USER_KEY") & " and SURVEY_KEY = " & Session("seqSurvey")
                                Else
                                    skip = True
                                End If
                            End If
                        Else
                            Me.sqlCommonAction.CommandText = "Insert into " & myUtility.getExtension() & "SAT_RESPONSE (RESPONDENT_USER_KEY, SURVEY_KEY, " & _
                            "CREATE_DATE, LAST_USED_DATE, REMIND_DATE, SUBMISSION_COUNT, LAST_COMPLETION_DATE) Values (" & _
                            row.Item("USER_KEY") & ", " & Session("seqSurvey") & ", '" & Now & "', " & _
                            "'1/1/1970', '1/1/1970', 0, '1/1/1970')"
                        End If

                        'modify the fempResponse table
                        Me.sqlCommonAction.ExecuteNonQuery()
                    End If
                Next

                'set the message body
                If (Session("RecipientsSelected") <> "NONE") Then
                    strMessage = "Surveys would have been sent to the following recipients: <br/><blockquote>" & _
                    Me.strRecipientsList & "</blockquote><br/>They would have received the following text " & _
                    "in the e-mail sent to them: <br/>" & _
                    "<DIV></DIV>Greetings,<BR><BR>We are asking you to complete an on-line " & _
                    "survey to help us understand your perceptions. " & Me.txtMessage.Text & _
                    " Important information about this survey follows. " & _
                    "<P>To complete the survey visit the following link using your internet browser: " & _
                    "<BLOCKQUOTE><A href='http://training.pnl.gov/Default.asp?seqUserID=NNN&amp;" & _
                    "seqSurvey=MM&amp;auth=XXXX'>" & _
                    "http://training.pnl.gov/Default.asp?seqUserID=NNN&amp;seqSurvey=MM&amp;auth=XXXX</A></BLOCKQUOTE> " & _
                    "<P>Your willingness to provide us with your candid input is appreciated.<BR>" & _
                    "Note: Each user would get their own survey link according to their survey respondent " & _
                    "account information in the Survey Administration Tool.<br/><hr/><br/>"
                End If

                sqlCommonAction.CommandText = "select distinct(php.internet_email_address) " & _
                                                  " from opwhse.dbo.vw_pub_rbac_role_all_members prra, " & _
                                                  " opwhse.dbo.vw_pub_hanford_people php where prra.child_role_name = " & _
                                                  " 'HumSub.Humn Subj Cont' and prra.hanford_id = php.hanford_id"
                'sqlCommonAction.CommandText = "Select * from fempUsers where strEmail = 'jesse.sharp@pnl.gov'"
                sqlCommonAdapter.SelectCommand = sqlCommonAction
                sqlCommonAdapter.Fill(Me.dsCommon, "hsrInfo")
                Dim HSREmailAddress As String = Me.dsCommon.Tables("hsrInfo").Rows(0).Item("internet_email_address")

                strMessage &= "You would have received the following text in the e-mail sent to you: <br/><br/>" & _
                "<DIV></DIV>Greetings,<BR><BR>[Insert your name and/or department here] is " & _
                "interested in gathering your opinions. We are asking you to complete an on-line " & _
                "survey to help us understand your perceptions. " & Me.txtMessage.Text & _
                " Important information about this survey follows. " & _
                "<P>To complete the survey visit the following link using your internet browser: " & _
                "<BLOCKQUOTE><A href='http://training.pnl.gov/Default.asp?seqUserID" & _
                Session("UserID") & "&amp;seqSurvey=" & Session("seqSurvey") & _
                "&amp;auth=" & strAuthenticator & "'>http://training.pnl.gov/Default.asp?seqSurvey=" & _
                Session("seqSurvey") & "&amp;auth=" & strAuthenticator & "</A></BLOCKQUOTE> " & _
                "<P>Your willingness to provide us with your candid input is appreciated.<BR> " & _
                "<BLOCKQUOTE><TABLE cellSpacing=0 cellPadding=0 width=550 border=0>" & _
                "<TBODY><TR><TD><P><STRONG>About this Survey</STRONG></P>" & _
                "<UL><LI><STRONG>Responses are anonymous</STRONG>Your responses to this " & _
                "on-line survey will be sent directly to the survey administration " & _
                "company server which is not associated with and cannot be accessed by " & _
                "your employer. This ensures that your specific responses will never be " & _
                "available to the organization or individuals that you work for. Your " & _
                "responses will only be available as aggregated group information.<BR>" & _
                "<LI><STRONG>Participation is Voluntary</STRONG>This survey is " & _
                "entirely voluntary, and you are free to choose at any time whether or " & _
                "not to provide responses to the survey or individual questions.<BR>" & _
                "<LI><STRONG>Your Rights </STRONG>If you have questions about your " & _
                "rights as a participant of this research survey or this website, " & _
                "please email the <A href='mailto:" & HSREmailAddress & "'>Institutional " & _
                "Review Board</A> at Pacific Northwest National Laboratory. A research " & _
                "specialist will respond to your question promptly.<BR></LI></UL></TD>" & _
                "</TR></TBODY></TABLE></BLOCKQUOTE><P>Sincerely, <P>[signature block]<BR/><br/><br/>" & _
                "Note: You can use this e-mail to send anonymous surveys.  Users will use your authentication " & _
                "to verify that they are allowed to take the survey.  This in no way gives them access" & _
                " to use the tool, and will not compromise your account."

                If Not (myUtility.genericSendMail(strFrom, strTo, strSubject, strMessage)) Then
                    Return False
                Else
                    Return True
                End If
            Else
                'set the subject for the e-mail
                strSubject = Me.txtSubject.Text

                'get the sender's e-mail address
                Dim row As DataRow
                For Each row In Me.dsCommon.Tables("MasterUserList").Rows
                    If (row.Item("USER_KEY") = Session("UserID")) Then
                        strFrom = row.Item("EMAIL_ADDRESS")
                        Exit For
                    End If
                Next

                'Get the connection
                Me.sqlCommonAction.Connection = Me.commonConn

                'if doing a follow-up also modify the fempSurvey closing date
                If (CType(Session("isFollowUp"), Boolean)) Then
                    Me.sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY set SURVEY_CLOSING_DATE = '" & _
                    Me.wdcFollowUp.Text & "' where SURVEY_KEY = " & Session("seqSurvey")
                    Me.sqlCommonAction.ExecuteNonQuery()
                End If

                'find the user's authenticator
                Dim strAuthenticator As String = ""
                Dim skip As Boolean
                For Each row In Me.dsCommon.Tables("MasterUserList").Rows
                    'get the authenticator and destination e-mail address
                    strAuthenticator = row.Item("AUTHENTICATION_CODE")
                    strTo = row.Item("EMAIL_ADDRESS")
                    strMessage = ""
                    skip = False

                    'only update fempResponse if the recipient is not also the sender
                    Dim needsReminder As Boolean = False
                    If (row.Item("USER_KEY") <> Session("UserID")) Then
                        'determine if the user has a response record already created.
                        Me.sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_RESPONSE where RESPONDENT_USER_KEY = " & _
                        row.Item("USER_KEY") & " and SURVEY_KEY = " & Session("seqSurvey") & _
                        " order by SUBMISSION_COUNT Desc"
                        If (Me.dsCommon.Tables.Contains("ExistingResponse")) Then
                            Me.dsCommon.Tables("ExistingResponse").Rows.Clear()
                        End If
                        Me.sqlCommonAdapter.Fill(Me.dsCommon, "ExistingResponse")

                        'update if the response already exists or create a new record if it does not
                        If (Me.dsCommon.Tables("ExistingResponse").Rows.Count > 0) Then
                            If (Me.dsCommon.Tables("ExistingResponse").Rows(0).Item("LAST_COMPLETION_DATE") = "01/01/1970") Then
                                needsReminder = True
                            End If

                            'if doing a follow-up adjust created/remind/date-used
                            If (CType(Session("isFollowUp"), Boolean)) Then
                                If (needsReminder) Then
                                    Me.sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_RESPONSE set REMIND_DATE = '" & _
                                    Now & "' where RESPONDENT_USER_KEY = " & _
                                    row.Item("USER_KEY") & " and SURVEY_KEY = " & Session("seqSurvey")
                                Else
                                    Me.sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_RESPONSE set CREATE_DATE = '" & _
                                    Now & "', REMIND_DATE = '1/1/1970', LAST_USED_DATE = '1/1/1970'" & _
                                    ", SUBMISSION_COUNT = " & _
                                    Me.dsCommon.Tables("ExistingResponse").Rows(0).Item("SUBMISSION_COUNT") + 1 & _
                                    "where RESPONDENT_USER_KEY = " & row.Item("USER_KEY") & _
                                    " and SURVEY_KEY = " & Session("seqSurvey")
                                End If
                            ElseIf (Me.chkbxReminder.Checked) Then
                                If (needsReminder) Then
                                    Me.sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_RESPONSE set REMIND_DATE = '" & _
                                    Now & "' where RESPONDENT_USER_KEY = " & _
                                    row.Item("USER_KEY") & " and SURVEY_KEY = " & Session("seqSurvey")
                                Else
                                    skip = True
                                End If
                            Else
                                If (needsReminder) Then
                                    Me.sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_RESPONSE set REMIND_DATE = '" & _
                                    Now & "' where RESPONDENT_USER_KEY = " & _
                                    row.Item("USER_KEY") & " and SURVEY_KEY = " & Session("seqSurvey")
                                Else
                                    skip = True
                                End If
                            End If
                        Else
                            Me.sqlCommonAction.CommandText = "Insert into " & myUtility.getExtension() & "SAT_RESPONSE (RESPONDENT_USER_KEY, SURVEY_KEY, " & _
                            "CREATE_DATE, LAST_USED_DATE, REMIND_DATE, SUBMISSION_COUNT, LAST_COMPLETION_DATE) Values (" & _
                            row.Item("USER_KEY") & ", " & Session("seqSurvey") & ", '" & Now & "', " & _
                            "'1/1/1970', '1/1/1970', 0, '1/1/1970')"
                        End If

                        'modify the fempResponse table
                        Me.sqlCommonAction.ExecuteNonQuery()
                    End If

                    'send the e-mail only if we did not skip this person
                    If Not (skip) Then
                        If (row.Item("USER_KEY") = Session("UserID")) Then
                            sqlCommonAction.CommandText = "select distinct(php.internet_email_address) " & _
                                                  " from opwhse.dbo.vw_pub_rbac_role_all_members prra, " & _
                                                  " opwhse.dbo.vw_pub_hanford_people php where prra.child_role_name = " & _
                                                  " 'HumSub.Humn Subj Cont' and prra.hanford_id = php.hanford_id"
                            'sqlCommonAction.CommandText = "Select * from fempUsers where strEmail = 'jesse.sharp@pnl.gov'"
                            sqlCommonAdapter.SelectCommand = sqlCommonAction
                            sqlCommonAdapter.Fill(Me.dsCommon, "hsrInfo")
                            Dim HSREmailAddress As String = Me.dsCommon.Tables("hsrInfo").Rows(0).Item("internet_email_address")

                            strMessage &= "<DIV></DIV>Greetings,<BR><BR>[Insert your name and/or department here] is " & _
                            "interested in gathering your opinions. We are asking you to complete an on-line " & _
                            "survey to help us understand your perceptions. " & Me.txtMessage.Text & _
                            " Important information about this survey follows. " & _
                            "<P>To complete the survey visit the following link using your internet browser: " & _
                            "<BLOCKQUOTE><A href='http://training.pnl.gov/Default.asp?seqUserID=" & _
                            row.Item("seqUserID") & "&amp;seqSurvey=" & Session("seqSurvey") & _
                            "&amp;auth=" & strAuthenticator & "'>http://training.pnl.gov/Default.asp?seqSurvey=" & _
                            Session("seqSurvey") & "&amp;auth=" & strAuthenticator & "</A></BLOCKQUOTE> " & _
                            "<P>Your willingness to provide us with your candid input is appreciated.<BR> " & _
                            "<BLOCKQUOTE><TABLE cellSpacing=0 cellPadding=0 width=550 border=0>" & _
                            "<TBODY><TR><TD><P><STRONG>About this Survey</STRONG></P>" & _
                            "<UL><LI><STRONG>Responses are anonymous</STRONG>Your responses to this " & _
                            "on-line survey will be sent directly to the survey administration " & _
                            "company server which is not associated with and cannot be accessed by " & _
                            "your employer. This ensures that your specific responses will never be " & _
                            "available to the organization or individuals that you work for. Your " & _
                            "responses will only be available as aggregated group information.<BR>" & _
                            "<LI><STRONG>Participation is Voluntary</STRONG>This survey is " & _
                            "entirely voluntary, and you are free to choose at any time whether or " & _
                            "not to provide responses to the survey or individual questions.<BR>" & _
                            "<LI><STRONG>Your Rights </STRONG>If you have questions about your " & _
                            "rights as a participant of this research survey or this website, " & _
                            "please email the <A href='mailto:" & HSREmailAddress & "'>Institutional " & _
                            "Review Board</A> at Pacific Northwest National Laboratory. A research " & _
                            "specialist will respond to your question promptly.<BR></LI></UL></TD>" & _
                            "</TR></TBODY></TABLE></BLOCKQUOTE><P>Sincerely, <P>[signature block]<BR/><br/>"
                        Else
                            If (Me.chkbxReminder.Checked = True And needsReminder) Then
                                strMessage = "<DIV></DIV>Greetings,<BR><BR>We are sending this survey as a reminder " & _
                                "that we are asking you to complete an on-line " & _
                                "survey to help us understand your perceptions. " & Me.txtMessage.Text & _
                                " Important information about this survey follows. " & _
                                "<P>To complete the survey visit the following link using your internet browser: " & _
                                "<BLOCKQUOTE><A href='http://training.pnl.gov/Default.asp?seqUserID=" & _
                                row.Item("USER_KEY") & "&amp;seqSurvey=" & Session("seqSurvey") & _
                                "&amp;auth=" & strAuthenticator & "'>" & _
                                "http://training.pnl.gov/Default.asp?seqUserID=" & row.Item("USER_KEY") & _
                                "&amp;seqSurvey=" & Session("seqSurvey") & "&amp;auth=" & strAuthenticator & _
                                "</A></BLOCKQUOTE><P>Your willingness to provide us with your candid input is appreciated."
                            Else
                                strMessage = "<DIV></DIV>Greetings,<BR><BR>We are asking you to complete an on-line " & _
                                "survey to help us understand your perceptions. " & Me.txtMessage.Text & _
                                " Important information about this survey follows. " & _
                                "<P>To complete the survey visit the following link using your internet browser: " & _
                                "<BLOCKQUOTE><A href='http://training.pnl.gov/Default.asp?seqUserID=" & _
                                row.Item("USER_KEY") & "&amp;seqSurvey=" & Session("seqSurvey") & _
                                "&amp;auth=" & strAuthenticator & "'>" & _
                                "http://training.pnl.gov/Default.asp?seqUserID=" & row.Item("USER_KEY") & _
                                "&amp;seqSurvey=" & Session("seqSurvey") & "&amp;auth=" & strAuthenticator & _
                                "</A></BLOCKQUOTE><P>Your willingness to provide us with your candid input is appreciated."
                            End If
                        End If

                        'TO DO: change to the actual user instead of the sender'
                        strTo = strFrom

                        'send the e-mail - if failure at any time alert the user and show who got the e-mail
                        If Not (myUtility.genericSendMail(strFrom, strTo, strSubject, strMessage)) Then
                            If (SentList <> "") Then
                                alert("Error sending surveys.  Not all surveys may have been sent. " & _
                                "The following surveys were sent: " & SentList)
                            Else
                                alert("Error sending surveys.  No surveys were sent.")
                            End If
                            Return False
                        Else
                            SentList &= row.Item("Name") & " - " & row.Item("EMAIL_ADDRESS") & "; "
                        End If
                    End If
                Next

                Return True
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            If (SentList <> "") Then
                alert("Error sending surveys.  Not all surveys may have been sent. " & _
                "The following surveys were sent: " & SentList)
            Else
                alert("Error sending surveys.  No surveys were sent.")
            End If
            Return False
        End Try
    End Function
#End Region

#Region "Send Survey"
    'sends the survey
    Private Function doSendSurvey() As Boolean
        Dim tempMasterList As ArrayList = CType(Session("MasterIDList"), ArrayList)
        Dim checkedGroupsList As ArrayList = CType(Session("CheckedGroupsList"), ArrayList)

        Try
            'process the groups section - session of selected users in groups
            If (processRecipients()) Then
                'process the recipient section - session of selected users in recipient
                If (processGroups()) Then
                    'generate the recipient list
                    If (buildRecipientList(True)) Then
                        'make sure we have a subject and message before sending
                        If (Trim(Me.txtSubject.Text) <> "" And Trim(Me.txtMessage.Text) <> "") Then
                            'send the surveys out
                            If Not (sendSurveys()) Then
                                Return True
                            Else
                                alert("Surveys sent.")
                                Return False
                            End If
                        Else
                            alert("You must have a subject and message before sending this survey.")
                            Return True
                        End If
                    Else
                        alert("failed to build recipient list.")
                        Return True
                    End If
                Else
                    alert("failed to process recipient groups.")
                    Session("MasterIDList") = tempMasterList
                    Session("CheckedGroupsList") = checkedGroupsList
                    Return True
                End If
            Else
                alert("failed to process recipients.")
                Session("MasterIDList") = tempMasterList
                Session("CheckedGroupsList") = checkedGroupsList
                Return True
            End If
            Return False
        Catch ex As Exception
            alert("failed to send the surveys.")
            Session("MasterIDList") = tempMasterList
            Session("CheckedGroupsList") = checkedGroupsList
            Return True
        End Try
    End Function
#End Region

#Region "Reset Code"
    'resets the page
    Private Function doReset() As Boolean
        If Not (loadData()) Then
            alert("Failed to load the groups for the current survey.")
            Return True
        Else
            Return False
        End If
    End Function
#End Region
End Class
