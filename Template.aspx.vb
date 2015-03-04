Imports System.Text
Imports System.Collections.Specialized
Public Class Template
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Template))
        Me.sqlTemplates = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand
        Me.dsCommon = New System.Data.DataSet
        Me.CommonConn = New System.Data.SqlClient.SqlConnection
        Me.SqlInsertCommand = New System.Data.SqlClient.SqlCommand
        Me.SqlUpdateCommand = New System.Data.SqlClient.SqlCommand
        Me.SqlDeleteCommand = New System.Data.SqlClient.SqlCommand
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'sqlTemplates
        '
        Me.sqlTemplates.DeleteCommand = Me.SqlDeleteCommand
        Me.sqlTemplates.InsertCommand = Me.SqlInsertCommand
        Me.sqlTemplates.SelectCommand = Me.SqlSelectCommand1
        Me.sqlTemplates.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_TEMPLATE", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("TEMPLATE_KEY", "TEMPLATE_KEY"), New System.Data.Common.DataColumnMapping("TEMPLATE_NAME", "TEMPLATE_NAME"), New System.Data.Common.DataColumnMapping("TEMPLATE_DESCRIPTION", "TEMPLATE_DESCRIPTION"), New System.Data.Common.DataColumnMapping("TEMPLATE_CREATE_DATE", "TEMPLATE_CREATE_DATE"), New System.Data.Common.DataColumnMapping("ACTIVE_IND", "ACTIVE_IND"), New System.Data.Common.DataColumnMapping("CREATOR_USER_KEY", "CREATOR_USER_KEY"), New System.Data.Common.DataColumnMapping("TRAINING_IND", "TRAINING_IND"), New System.Data.Common.DataColumnMapping("OPTIONAL_IND", "OPTIONAL_IND")})})
        Me.sqlTemplates.UpdateCommand = Me.SqlUpdateCommand
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = "SELECT TEMPLATE_KEY, TEMPLATE_NAME, TEMPLATE_DESCRIPTION, TEMPLATE_CREATE_DATE, A" & _
            "CTIVE_IND, CREATOR_USER_KEY, TRAINING_IND, OPTIONAL_IND FROM VARCONN.SAT_TE" & _
            "MPLATE"
        Me.SqlSelectCommand1.Connection = Me.CommonConn
        '
        'dsCommon
        '
        Me.dsCommon.DataSetName = "NewDataSet"
        Me.dsCommon.Locale = New System.Globalization.CultureInfo("en-US")
        '
        'CommonConn
        '
        Me.CommonConn.FireInfoMessageEventOnUserErrors = False
        '
        'SqlInsertCommand
        '
        Me.SqlInsertCommand.CommandText = resources.GetString("SqlInsertCommand.CommandText")
        Me.SqlInsertCommand.Connection = Me.CommonConn
        Me.SqlInsertCommand.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@TEMPLATE_NAME", System.Data.SqlDbType.VarChar, 0, "TEMPLATE_NAME"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_DESCRIPTION", System.Data.SqlDbType.VarChar, 0, "TEMPLATE_DESCRIPTION"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_CREATE_DATE", System.Data.SqlDbType.DateTime, 0, "TEMPLATE_CREATE_DATE"), New System.Data.SqlClient.SqlParameter("@ACTIVE_IND", System.Data.SqlDbType.Bit, 0, "ACTIVE_IND"), New System.Data.SqlClient.SqlParameter("@CREATOR_USER_KEY", System.Data.SqlDbType.Int, 0, "CREATOR_USER_KEY"), New System.Data.SqlClient.SqlParameter("@TRAINING_IND", System.Data.SqlDbType.Bit, 0, "TRAINING_IND"), New System.Data.SqlClient.SqlParameter("@OPTIONAL_IND", System.Data.SqlDbType.Bit, 0, "OPTIONAL_IND")})
        '
        'SqlUpdateCommand
        '
        Me.SqlUpdateCommand.CommandText = resources.GetString("SqlUpdateCommand.CommandText")
        Me.SqlUpdateCommand.Connection = Me.CommonConn
        Me.SqlUpdateCommand.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@TEMPLATE_NAME", System.Data.SqlDbType.VarChar, 0, "TEMPLATE_NAME"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_DESCRIPTION", System.Data.SqlDbType.VarChar, 0, "TEMPLATE_DESCRIPTION"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_CREATE_DATE", System.Data.SqlDbType.DateTime, 0, "TEMPLATE_CREATE_DATE"), New System.Data.SqlClient.SqlParameter("@ACTIVE_IND", System.Data.SqlDbType.Bit, 0, "ACTIVE_IND"), New System.Data.SqlClient.SqlParameter("@CREATOR_USER_KEY", System.Data.SqlDbType.Int, 0, "CREATOR_USER_KEY"), New System.Data.SqlClient.SqlParameter("@TRAINING_IND", System.Data.SqlDbType.Bit, 0, "TRAINING_IND"), New System.Data.SqlClient.SqlParameter("@OPTIONAL_IND", System.Data.SqlDbType.Bit, 0, "OPTIONAL_IND"), New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_CREATOR_USER_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CREATOR_USER_KEY", System.Data.DataRowVersion.Original, Nothing)})
        '
        'SqlDeleteCommand
        '
        Me.SqlDeleteCommand.CommandText = "DELETE FROM VARCONN.[SAT_TEMPLATE] WHERE (([TEMPLATE_KEY] = @Original_TE" & _
            "MPLATE_KEY) AND ([CREATOR_USER_KEY] = @Original_CREATOR_USER_KEY))"
        Me.SqlDeleteCommand.Connection = Me.CommonConn
        Me.SqlDeleteCommand.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_CREATOR_USER_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CREATOR_USER_KEY", System.Data.DataRowVersion.Original, Nothing)})
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected WithEvents hlpServerSwitch As System.Web.UI.WebControls.Image
    Protected WithEvents chkServerSwitch As System.Web.UI.WebControls.CheckBox
    Protected WithEvents Image1 As System.Web.UI.WebControls.Image
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents hlpCopy As System.Web.UI.WebControls.Image
    Protected WithEvents btnCopy As System.Web.UI.WebControls.Button
    Protected WithEvents hlpSelectTemplate As System.Web.UI.WebControls.Image
    Protected WithEvents sqlTemplates As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents dsCommon As System.Data.DataSet
    Protected WithEvents ddlTemplateList As System.Web.UI.WebControls.DropDownList
    Protected WithEvents lblSelectTemplate As System.Web.UI.WebControls.Label
    Protected WithEvents hlpTemplateName As System.Web.UI.WebControls.Image
    Protected WithEvents hlpTemplateDescription As System.Web.UI.WebControls.Image
    Protected WithEvents hlpTrainingSurvey As System.Web.UI.WebControls.Image
    Protected WithEvents hlpSurveyOptional As System.Web.UI.WebControls.Image
    Protected WithEvents hlpTemplateActive As System.Web.UI.WebControls.Image
    Protected WithEvents hlpSetPermissions As System.Web.UI.WebControls.Image
    Protected WithEvents hlpTemplatePreview As System.Web.UI.WebControls.Image
    Protected WithEvents hlpTags As System.Web.UI.WebControls.Image
    Protected WithEvents hlpQuestions As System.Web.UI.WebControls.Image
    Protected WithEvents hlpSurveys As System.Web.UI.WebControls.Image
    Protected WithEvents hlpDelegates As System.Web.UI.WebControls.Image
    Protected WithEvents lblTemplateSelect As System.Web.UI.WebControls.Label
    Protected WithEvents txtTemplateName As System.Web.UI.WebControls.TextBox
    Protected WithEvents lblTemplateDescription As System.Web.UI.WebControls.Label
    Protected WithEvents txtDescription As System.Web.UI.WebControls.TextBox
    Protected WithEvents chkTQSurvey As System.Web.UI.WebControls.CheckBox
    Protected WithEvents chkOptional As System.Web.UI.WebControls.CheckBox
    Protected WithEvents chkActive As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblSurveyOptional As System.Web.UI.WebControls.Label
    Protected WithEvents lblTemplateActive As System.Web.UI.WebControls.Label
    Protected WithEvents lblTrainingSurvey As System.Web.UI.WebControls.Label
    Protected WithEvents btnSetPermissions As System.Web.UI.WebControls.Button
    Protected WithEvents btnTemplatePreview As System.Web.UI.WebControls.Button
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label3 As System.Web.UI.WebControls.Label
    Protected WithEvents Label4 As System.Web.UI.WebControls.Label
    Protected WithEvents rptrQuestions As System.Web.UI.WebControls.Repeater
    Protected WithEvents rptrSurveys As System.Web.UI.WebControls.Repeater
    Protected WithEvents rptrDelegates As System.Web.UI.WebControls.Repeater
    Protected WithEvents rptrTags As System.Web.UI.WebControls.Repeater
    Protected strSurveyList As String
    Protected myUtility As New Utility
    Protected WithEvents Td1 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents txtDuplicate As String
    Protected WithEvents hlpTemplatePreview2 As System.Web.UI.WebControls.Image
    Protected WithEvents btnTemplatePreview2 As System.Web.UI.WebControls.Button
    Protected WithEvents btnSubmit As System.Web.UI.WebControls.Button
    Protected sqlCommonAdapter As New System.Data.SqlClient.SqlDataAdapter
    Protected sqlCommonAction As New System.Data.SqlClient.SqlCommand
    Protected WithEvents Label5 As System.Web.UI.WebControls.Label
    Protected WithEvents hlpCopy2 As System.Web.UI.WebControls.Image
    Protected WithEvents btnCopy2 As System.Web.UI.WebControls.Button
    Protected WithEvents hlpQuestionGroups As System.Web.UI.WebControls.Image
    Protected WithEvents lblQuestionGroups As System.Web.UI.WebControls.Label
    Protected WithEvents rptrQuestionGroups As System.Web.UI.WebControls.Repeater
    Protected WithEvents Label6 As System.Web.UI.WebControls.Label
    Protected WithEvents CommonConn As System.Data.SqlClient.SqlConnection
    Friend WithEvents SqlDeleteCommand As System.Data.SqlClient.SqlCommand
    Friend WithEvents SqlInsertCommand As System.Data.SqlClient.SqlCommand
    Friend WithEvents SqlUpdateCommand As System.Data.SqlClient.SqlCommand
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
        End If

        'get the from page
        If Not (IsPostBack) Then
            Session("FromPage") = getFromPage()
            Session("BackgroundColor") = "-1"
        End If

        'Put user code to initialize the page here
        If (Session("isAuthenticated") <> "True") Then
            Session("pageSel") = "Login"
            redirect("./Login.aspx")
        Else
            Session("CurrentPage") = "./Template.aspx"
            Session("pageSel") = "Template"
        End If

        'if the user's password has expired send them to the password reset page
        If (Session("isExpired") = "True") Then
            Session("CurrentPage") = "./Password.aspx"
            Session("pageSel") = "Password"
            redirect(Session("CurrentPage"))
        End If

        'Set the connection based on the machine
        myUtility.getConnection(Me.CommonConn)

        'Set the server switch text on initial page load
        If Not (IsPostBack) Then
            'get anything on the request line
            myUtility.getRequest(Page.Request.Url.Query())

            'set the template id if it is not set at all
            If (Session("seqTemplate") Is Nothing) Then
                Session("seqTemplate") = -1
            End If

            'Fill the template and survey drop-downs
            loadData()

            Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
            Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
            sqlCommonAction.Connection = Me.CommonConn
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            myUtility.determineUserOwnership(sqlCommonAction, sqlCommonAdapter, True)

            'set the page options
            myUtility.optionPreSetTemplate(Session("seqTemplate"), Me.dsCommon.Tables("Templates").Rows.Count(), Me.pageOptions, True)

            Session("PageOptions") = Me.pageOptions

            'alter the control text
            changeControlText()

            Me.btnInsert.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
            Me.btnUpdate.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
            Me.btnDelete.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
            Me.btnReset.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
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

#Region "Data Loading"
    'Loads data into the form
    Private Sub loadData(Optional ByVal failure As Boolean = False)
        Trace.Warn("Loading Data")

        Try
            'Set up the common data adapter and select command
            Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
            Dim sqlCommonSelect As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
            sqlCommonSelect.Connection = Me.CommonConn
            sqlCommonAdapter.SelectCommand = sqlCommonSelect

            'Fill the template list
            templateFill()

            'loads the rest of the data if a template has been selected
            If (Session("seqTemplate") <> -1 And Session("TemplatesExist") = "Yes") Then
                'Fill the Tags list
                If (tagsFill(sqlCommonSelect, sqlCommonAdapter)) Then
                    'Fill the Questions list
                    If (questionsFill(sqlCommonSelect, sqlCommonAdapter)) Then
                        'Fill the Question Group list
                        If (questionGroupsFill(sqlCommonSelect, sqlCommonAdapter)) Then
                            'Fill the Surveys list
                            If (surveysFill(sqlCommonSelect, sqlCommonAdapter)) Then
                                'List the Delegates/Owners
                                If (delegatesFill(sqlCommonSelect, sqlCommonAdapter)) Then
                                    'Populate the page with the template data
                                    If Not (failure) Then
                                        Me.txtTemplateName.Text = Me.dsCommon.Tables("Templates").Rows(Me.ddlTemplateList.SelectedIndex() - 1).Item("TEMPLATE_NAME")
                                        Me.txtDescription.Text = Me.dsCommon.Tables("Templates").Rows(Me.ddlTemplateList.SelectedIndex() - 1).Item("TEMPLATE_DESCRIPTION")
                                        Me.chkTQSurvey.Checked = Me.dsCommon.Tables("Templates").Rows(Me.ddlTemplateList.SelectedIndex() - 1).Item("TRAINING_IND")
                                        Me.chkOptional.Checked = Me.dsCommon.Tables("Templates").Rows(Me.ddlTemplateList.SelectedIndex() - 1).Item("OPTIONAL_IND")
                                        Me.chkActive.Checked = Me.dsCommon.Tables("Templates").Rows(Me.ddlTemplateList.SelectedIndex() - 1).Item("ACTIVE_IND")
                                    End If

                                    'determine the approval status of the selected template
                                    Session("isApproved") = myUtility.isApproved(Me.dsCommon, Me.CommonConn)
                                    If (Session("Alert") <> "") Then
                                        alert(Session("Alert"))
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else
                'Clear all fields
                Me.txtDescription.Text = ""
                Me.txtTemplateName.Text = ""
                Me.chkTQSurvey.Checked = False
                Me.chkOptional.Checked = False
                Me.chkActive.Checked = True
                Session("seqTemplate") = -1
            End If
        Catch ex As Exception
            alert("Failed to load all template information.")
        End Try
    End Sub

    'Fill the list of templates
    Private Sub templateFill()
        Trace.Warn("Template Fill")
        Try
            'Execute the SQL Call based on the user type
            If (Session("UserType") <> 4 And Session("UserType") <> 1) Then
                Me.SqlSelectCommand1.CommandText = "SELECT * FROM " & myUtility.getExtension() & "SAT_TEMPLATE ft, " & myUtility.getExtension() & "SAT_USER_PERMISSION fp where " & _
                "ft.TEMPLATE_KEY = fp.TEMPLATE_KEY AND fp.USER_KEY = " & Session("UserID") & _
                " Order by TEMPLATE_NAME"

                'clear out the old data if it's there
                If (Me.dsCommon.Tables.Contains("Templates")) Then
                    Me.dsCommon.Tables("Templates").Rows.Clear()
                End If
                Me.SqlSelectCommand1.Connection = Me.CommonConn
                Me.sqlTemplates.Fill(dsCommon, "Templates")
            Else
                Me.SqlSelectCommand1.CommandText = "SELECT * FROM " & myUtility.getExtension() & "SAT_TEMPLATE Order by " & _
                "TEMPLATE_NAME"
                'clear out the old data if it's there
                If (Me.dsCommon.Tables.Contains("Templates")) Then
                    Me.dsCommon.Tables("Templates").Rows.Clear()
                End If
                Me.SqlSelectCommand1.Connection = Me.CommonConn
                Me.sqlTemplates.Fill(dsCommon, "Templates")
            End If

            'if there are templates then build the list of templates otherwise set the existence session to "No"
            If (Me.dsCommon.Tables("Templates").Rows.Count() > 0) Then
                'Bind the data and add an additional item for no template selected
                Me.ddlTemplateList.DataSource = Me.dsCommon.Tables("Templates")
                Me.ddlTemplateList.DataValueField = "TEMPLATE_KEY"
                Me.ddlTemplateList.DataTextField = "TEMPLATE_NAME"
                Me.ddlTemplateList.DataBind()
                Dim li As System.Web.UI.WebControls.ListItem = New System.Web.UI.WebControls.ListItem
                li.Text = "-- Select a Template --"
                li.Value = "-1"
                Me.ddlTemplateList.Items.Insert(0, li)

                'Check to see if this is our first time on the page or not
                If (Session("seqTemplate") <> -1 And Not Me.ddlTemplateList.Items.FindByValue(Session("seqTemplate")) Is Nothing) Then
                    Me.ddlTemplateList.Items.FindByValue(Session("seqTemplate")).Selected = True
                    Session("isDirty") = Me.myUtility.isDirty(Me.dsCommon, Me.CommonConn)
                Else
                    Me.ddlTemplateList.SelectedIndex = 0
                End If

                Session("TemplatesExist") = "Yes"
            Else
                Session("TemplatesExist") = "No"
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Session("TemplatesExist") = "No"
        End Try
    End Sub

    'Fills the tags repeater
    Private Function tagsFill(ByVal sqlCommonAction As Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As Data.SqlClient.SqlDataAdapter) As Boolean
        Trace.Warn("Filling Tags")

        Try
            'clear out the questions if the are any residual
            If (Me.dsCommon.Tables.Contains("Tags")) Then
                Me.dsCommon.Tables("Tags").Rows.Clear()
            End If

            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_TEMPLATE_TAG where TEMPLATE_KEY = " & Session("seqTemplate")
            sqlCommonAdapter.Fill(Me.dsCommon, "Tags")
            Me.rptrTags.DataSource = Me.dsCommon.Tables("Tags")
            Me.rptrTags.DataBind()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to load the Template Tag Items.")
            Return False
        End Try
    End Function

    'Fills the questions repeater
    Private Function questionsFill(ByVal sqlCommonAction As Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As Data.SqlClient.SqlDataAdapter) As Boolean
        Trace.Warn("Filling Questions")

        Try
            'clear out the questions if the are any residual
            If (Me.dsCommon.Tables.Contains("Questions")) Then
                Me.dsCommon.Tables("Questions").Rows.Clear()
            End If

            sqlCommonAction.CommandText = "Select TEMPLATE_KEY,QUESTION_ID,REQUIRED_IND,PAGE_BREAK_IND" & _
            ",FILTER_IND,NOT_APPLICABLE_IND,QUESTION_TYPE_CODE," & _
            "CASE QUESTION_TEXT WHEN '<HR>' Then '__________' Else QUESTION_TEXT End as QUESTION_TEXT," & _
            "CHART_TYPE_CODE,RATING_COUNT" & _
            ",RATING_STEP_VALUE,RATING_INITIAL_VALUE,BRANCH_QUESTION_ID,QUESTION_GROUP_ID,QUERY_NAME" & _
            " from " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION where TEMPLATE_KEY = " & Session("seqTemplate")
            sqlCommonAdapter.Fill(Me.dsCommon, "Questions")
            Me.rptrQuestions.DataSource = Me.dsCommon.Tables("Questions")
            Me.rptrQuestions.DataBind()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to load the Template Questions.")
            Return False
        End Try
    End Function

    'Fills the question groups repeater
    Private Function questionGroupsFill(ByVal sqlCommonAction As Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As Data.SqlClient.SqlDataAdapter) As Boolean
        Trace.Warn("Filling Question Groups")

        Try
            'clear out the question groups if the are any residual
            If (Me.dsCommon.Tables.Contains("QuestionGroups")) Then
                Me.dsCommon.Tables("QuestionGroups").Rows.Clear()
            End If

            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION_GROUP where TEMPLATE_KEY = " & Session("seqTemplate")
            sqlCommonAdapter.Fill(Me.dsCommon, "QuestionGroups")
            Me.rptrQuestionGroups.DataSource = Me.dsCommon.Tables("QuestionGroups")
            Me.rptrQuestionGroups.DataBind()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to load the Template Question Groups.")
            Return False
        End Try
    End Function

    'Fills the surveys repeater and gets a table listing all surveys as well as a count of all surveys
    Private Function surveysFill(ByVal sqlCommonAction As Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As Data.SqlClient.SqlDataAdapter) As Boolean
        Trace.Warn("Filling Surveys")

        Try
            'clear out the surveys if there are any residual
            If (Me.dsCommon.Tables.Contains("Surveys")) Then
                Me.dsCommon.Tables("Surveys").Rows.Clear()
            End If

            'clear out the survey results if there are any residual
            If (Me.dsCommon.Tables.Contains("SurveyResults")) Then
                Me.dsCommon.Tables("SurveyResults").Rows.Clear()
            End If

            'Get count of survey results from production
            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY fs, " & _
            myUtility.getExtension() & "SAT_RESPONSE_RESULT fr where TEMPLATE_KEY = " & _
                Session("seqTemplate") & " and fs.SURVEY_KEY = fr.SURVEY_KEY"
            Trace.Warn(sqlCommonAction.CommandText)
            sqlCommonAdapter.Fill(Me.dsCommon, "SurveyResults")
            Session("Results") = Me.dsCommon.Tables("SurveyResults").Rows.Count()

            'Get only surveys available to the user
            If (Session("UserType") <> 4 And Session("UserType") <> 1) Then
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY fs, " & _
                    myUtility.getExtension() & "SAT_USER_PERMISSION fp where fs.TEMPLATE_KEY = " & _
                    Session("seqTemplate") & " and fs.SURVEY_KEY = fp.SURVEY_KEY and fp.USER_KEY = " & _
                    Session("UserID") & " and (fp.OWNER_IND = 1 or fp.DELEGATE_IND = 1)"
                sqlCommonAdapter.Fill(Me.dsCommon, "Surveys")
                Me.rptrSurveys.DataSource = Me.dsCommon.Tables("Surveys")
                Me.rptrSurveys.DataBind()
            Else
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY fs where fs.TEMPLATE_KEY = " & _
                    Session("seqTemplate")
                sqlCommonAdapter.Fill(Me.dsCommon, "Surveys")
                Me.rptrSurveys.DataSource = Me.dsCommon.Tables("Surveys")
                Me.rptrSurveys.DataBind()
            End If
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to load the Template Surveys.")
            Return False
        End Try
    End Function

    'Fills the delegates repeater
    Private Function delegatesFill(ByVal sqlCommonAction As Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As Data.SqlClient.SqlDataAdapter) As Boolean
        Trace.Warn("Filling Delegates")

        Try
            'clear out the questions if the are any residual
            If (Me.dsCommon.Tables.Contains("Delegates")) Then
                Me.dsCommon.Tables("Delegates").Rows.Clear()
            End If

            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER_PERMISSION fp, " & _
            myUtility.getExtension() & "SAT_USER fu where TEMPLATE_KEY = " & Session("seqTemplate") & " " & _
            "and fp.USER_KEY = fu.USER_KEY and (fp.OWNER_IND = 1 or fp.DELEGATE_IND = 1) order by fu.LAST_NAME, fu.FIRST_NAME"

            sqlCommonAdapter.Fill(Me.dsCommon, "Delegates")
            Me.rptrDelegates.DataSource = Me.dsCommon.Tables("Delegates")
            Me.rptrDelegates.DataBind()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to load the Template Delegates.")
            Return False
        End Try
    End Function

    'gets the list of surveys in the template for deletion
    Private Function getSurveyList(Optional ByVal tableName As String = "Surveys") As String
        'Get list of surveys if they exist or return true since there will be no responses
        Try
            If (Me.dsCommon.Tables(tableName).Rows.Count() > 0) Then
                Dim strSurveyList As System.Text.StringBuilder = New System.Text.StringBuilder
                Dim index As Integer = 0
                strSurveyList.Append("SURVEY_KEY in (")
                While (index < Me.dsCommon.Tables(tableName).Rows.Count())
                    If (index = 0) Then
                        strSurveyList.Append(Me.dsCommon.Tables(tableName).Rows(index).Item("SURVEY_KEY"))
                    Else
                        strSurveyList.Append("," & Me.dsCommon.Tables(tableName).Rows(index).Item("SURVEY_KEY"))
                    End If
                    index += 1
                End While
                strSurveyList.Append(") ")
                Return strSurveyList.ToString
            Else
                Return ""
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to determine the surveys that go with this Template.")
            Return ""
        End Try
    End Function
#End Region

#Region "Control Settings"
    'Changes text to reflect the current database
    Private Sub changeControlText()
        Trace.Warn("Change Control Text")
        Try
            'Set the visibility and interactibility of the controls based on template selection
            If (Session("seqTemplate") <> -1) Then
                If (Session("UserType") <> 4 And Session("UserType") <> 1) Then
                    'determine if the user is a delegate/owner
                    'Set up the common data adapter and select command
                    Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
                    Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
                    sqlCommonAction.Connection = Me.CommonConn
                    sqlCommonAdapter.SelectCommand = sqlCommonAction
                    myUtility.determineUserOwnership(sqlCommonAction, sqlCommonAdapter, True)

                    'set visibility and text of the controls based on the results
                    If ((CType(Session("isTemplateOwner"), Boolean) Or CType(Session("isTemplateDelegate"), Boolean)) And Session("AllSurveyCount") <= 0) Then
                        If (Not Session("isDirty") Or (Session("isDirty") And Session("isApproved"))) Then
                            Me.hlpCopy.ToolTip = "Clicking this button will take you to the publication request page."
                            Me.btnCopy.Text = "Re-Submit"
                            Me.btnCopy.ToolTip = "Re-Submit"
                            If (Session("UserType") >= 3) Then
                                setInteractivity(True)
                            Else
                                setInteractivity(False)
                            End If
                        ElseIf (Session("isDirty")) Then
                            Me.hlpCopy.ToolTip = "Clicking this button will take you to the publication request page."
                            Me.btnCopy.Text = "Publication Request"
                            Me.btnCopy.ToolTip = "Publication Request"
                            If (Session("UserType") >= 3) Then
                                setInteractivity(True)
                            Else
                                setInteractivity(False)
                            End If
                        End If
                    Else
                        setInteractivity(False)
                    End If
                Else
                    'determine if the user is a delegate/owner
                    'Set up the common data adapter and select command
                    Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
                    Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
                    sqlCommonAction.Connection = Me.CommonConn
                    sqlCommonAdapter.SelectCommand = sqlCommonAction
                    myUtility.determineUserOwnership(sqlCommonAction, sqlCommonAdapter, True)
                    If (Not Session("isDirty") Or (Session("isDirty") And Session("isApproved"))) Then
                        Me.hlpCopy.ToolTip = "Clicking this button will take you to the publication request page."
                        Me.btnCopy.Text = "Re-Submit"
                        Me.btnCopy.ToolTip = "Re-Submit"
                    ElseIf (Session("isDirty")) Then
                        Me.hlpCopy.ToolTip = "Clicking this button will take you to the publication request page."
                        Me.btnCopy.Text = "Publication Request"
                        Me.btnCopy.ToolTip = "Publication Request"
                    End If
                    If (Session("UserType") <> 1) Then
                        setInteractivity(True)
                    Else
                        setInteractivity(False)
                    End If
                End If
            Else
                'set the interactivity based on no template selected
                If (Session("UserType") < 3) Then
                    setInteractivity(False)
                Else
                    setInteractivity(True)
                End If
            End If
        Catch ex As Exception
            alert("Unable to determine your ownership status at this time.  This template will not be editable.")
        End Try
    End Sub

    'Sets the interactivity of the controls on the page
    Private Sub setInteractivity(ByVal Active As Boolean)
        Trace.Warn("Set Interactivity")

        If (Active) Then
            Me.chkActive.Enabled = True
            Me.hlpCopy.Visible = True
            Me.btnCopy.Visible = True
        Else
            Me.chkActive.Enabled = False
            Me.hlpCopy.Visible = False
            Me.btnCopy.Visible = False
        End If
    End Sub
#End Region

#Region "Checks"
    'Determines if the currently selected template is dirty
    Private Function isDirty() As Boolean
        'Execute the SQL Call based on the user type
        Try
            Me.SqlSelectCommand1.CommandText = "SELECT CHANGE_IND FROM " & myUtility.getExtension() & "SAT_TEMPLATE where " & _
                "TEMPLATE_KEY = " & Session("seqTemplate")
            'clear out the old data if it's there
            If (Me.dsCommon.Tables.Contains("Dirty")) Then
                Me.dsCommon.Tables("Dirty").Rows.Clear()
            End If
            Me.SqlSelectCommand1.Connection = Me.CommonConn
            Me.sqlTemplates.Fill(dsCommon, "Dirty")
            Return Me.dsCommon.Tables("Dirty").Rows(0).Item("CHANGE_IND")
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Unable to determine if this template has been altered or not.  Database error.  This template is assumed to have been altered.")
            Return True
        End Try
    End Function

    'determines whether the template name exists already
    Private Function doesTemplateNameExist(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, Optional ByVal isUpdate As Boolean = False, Optional ByVal isDuplicate As Boolean = False) As Boolean
        Trace.Warn("Verifying no other template has this same name.")
        Try
            'set up the sql command based on whether or not we're updating a current template or not
            If (isUpdate) Then
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_TEMPLATE where TEMPLATE_NAME = @TEMPLATE_NAME" & _
                       " AND TEMPLATE_KEY <> " & Session("seqTemplate")
            Else
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_TEMPLATE where TEMPLATE_NAME = @TEMPLATE_NAME"
            End If

            Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@TEMPLATE_NAME", System.Data.SqlDbType.VarChar, 50, "TEMPLATE_NAME")
            If (isDuplicate) Then
                param0.Value = Me.txtDuplicate
            Else
                param0.Value = Me.txtTemplateName.Text
            End If

            sqlCommonAction.Parameters.Add(param0)

            'select the template with the same name if it exists and return true of there others with the same name
            Dim ds As DataSet = New DataSet
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            sqlCommonAdapter.Fill(ds)
            If (ds.Tables(0).Rows.Count() > 0) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            alert("Unable to determine if template name exists or not.  Aborting.")
            Return True
        End Try
    End Function

    'determines if significant changes were made to the data
    Private Function significantChange() As Boolean
        Try
            'get the data
            Dim oldCommand As String = Me.sqlCommonAction.CommandText
            Me.sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_TEMPLATE where TEMPLATE_KEY = " & _
                Session("seqTemplate")
            If (Me.dsCommon.Tables.Contains("OriginalData")) Then
                Me.dsCommon.Tables("OriginalData").Rows.Clear()
            End If
            Me.sqlCommonAdapter.Fill(Me.dsCommon, "OriginalData")
            Me.sqlCommonAction.CommandText = oldCommand

            'compare data against alterable fields - only 1 row so loop runs once
            Dim dr As DataRow
            For Each dr In Me.dsCommon.Tables("OriginalData").Rows
                If (dr.Item("TEMPLATE_NAME") <> Me.txtTemplateName.Text) Then
                    Return True
                ElseIf (Me.txtDescription.Text <> dr.Item("TEMPLATE_DESCRIPTION")) Then
                    Return True
                ElseIf (Me.chkOptional.Checked <> CType(dr.Item("OPTIONAL_IND"), Boolean)) Then
                    Return True
                ElseIf (Me.chkTQSurvey.Checked <> CType(dr.Item("TRAINING_IND"), Boolean)) Then
                    Return True
                Else
                    Return False
                End If
            Next

            'default return of True if there was no data retrieved for the loop above to run through
            If (Me.dsCommon.Tables("OriginalData").Rows.Count = 0) Then
                alert("Unable to determine if significant changes were made.  This template has been marked dirty by default in this case.")
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            alert("Unable to determine if significant changes were made.  This template has been marked dirty by default in this case.")
            Return True
        End Try
    End Function
#End Region

#Region "Page Switch"
    'handles the user wanting to set permissions for users
    Private Sub btnSetPermissions_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSetPermissions.Click
        Session("JSAction") = ""
        Session("CurrentPage") = "./Permission.aspx"
        Session("pageSel") = "Permission"
        redirect(Session("CurrentPage"))
    End Sub

    'brings up a pop-up window with the template preview in it
    Private Sub btnTemplatePreview_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTemplatePreview.Click, btnTemplatePreview2.Click
        Session("JSAction") = "<script>newWin = window.open('./Preview.aspx', 'PreviewWindow', 'width=600,height=400,scrollbars=yes,toolbar=no,titlebar=no,status=no,resizable=yes,menubar=no');newWin.focus();</script>"
    End Sub
#End Region

#Region "Template Selection"
    'Fill the lists and populate the page with data
    Private Sub ddlTemplateList_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddlTemplateList.SelectedIndexChanged
        Session("JSAction") = ""

        'Store the recently selected template sequence number
        Session("seqTemplate") = Me.ddlTemplateList.SelectedItem.Value()
        Session("TemplateName") = Me.ddlTemplateList.SelectedItem.Text()
        Session("Results") = 0

        'loads the appropriate data
        loadData()

        'change control text
        changeControlText()

        Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
        Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
        sqlCommonAction.Connection = Me.CommonConn
        sqlCommonAdapter.SelectCommand = sqlCommonAction
        myUtility.determineUserOwnership(sqlCommonAction, sqlCommonAdapter, True)

        'set the page options
        myUtility.optionPreSetTemplate(Session("seqTemplate"), Me.dsCommon.Tables("Templates").Rows.Count(), Me.pageOptions, True)

        Session("PageOptions") = Me.pageOptions

        'reset lingering variables for approvers
        Session("isADCHSR") = Nothing
        Session("hsrHID") = Nothing
        Session("adcHID") = Nothing
    End Sub
#End Region

#Region "Page Action"
    'Handles the user attempting to do something with the template
    Private Sub btnInsert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsert.Click
        Session("JSAction") = ""

        'check for possible malicious code input
        Dim blnContinue As Boolean = True
        If Not (myUtility.goodInput(Me.txtDescription, True) And myUtility.goodInput(Me.txtTemplateName, True)) Then
            blnContinue = False
            alert("Possible malicious code entry(s).  Insert aborted.")
        End If

        If (blnContinue) Then
            'insert the template generating a new template even if it currently exists
            executeInsert()

            Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
            Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
            sqlCommonAction.Connection = Me.CommonConn
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            myUtility.determineUserOwnership(sqlCommonAction, sqlCommonAdapter, True)

            'set the page options
            myUtility.optionPreSetTemplate(Session("seqTemplate"), Me.dsCommon.Tables("Templates").Rows.Count(), Me.pageOptions, True)

            Session("PageOptions") = Me.pageOptions

            Session("TransExists") = False
        End If
    End Sub

    'Handles the user attempting to do something with the template
    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Session("JSAction") = ""

        'check for possible malicious code input
        Dim blnContinue As Boolean = True
        If Not (myUtility.goodInput(Me.txtDescription, True) And myUtility.goodInput(Me.txtTemplateName, True)) Then
            blnContinue = False
            alert("Possible malicious code entry(s).  Update aborted.")
        End If

        If (blnContinue) Then
            'determine if the template is dirty or not
            Dim tempConn As Data.SqlClient.SqlConnection = New Data.SqlClient.SqlConnection
            Session("isDirty") = myUtility.isDirty(Me.dsCommon, myUtility.getConnection(tempConn))
            If (myUtility.hasResults(Me.dsCommon, myUtility.getConnection(tempConn))) Then
                Session("hasResults") = "True"
            Else
                Session("hasResults") = "False"
            End If

            'Reset the form data to that in the db
            If (Session("hasResults") = "False") Then
                'Update the template
                executeUpdate()
            Else
                'duplicate the template 
                Me.txtDuplicate = Me.txtTemplateName.Text & "-New"
                executeDuplicate()
            End If

            Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
            Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
            sqlCommonAction.Connection = Me.CommonConn
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            myUtility.determineUserOwnership(sqlCommonAction, sqlCommonAdapter, True)

            'set the page options
            myUtility.optionPreSetTemplate(Session("seqTemplate"), Me.dsCommon.Tables("Templates").Rows.Count(), Me.pageOptions, True)

            Session("PageOptions") = Me.pageOptions

            Session("TransExists") = False
        End If
    End Sub

    'Handles the user attempting to do something with the template
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Session("JSAction") = ""

        'determine if the template is dirty or not
        Dim tempConn As Data.SqlClient.SqlConnection = New Data.SqlClient.SqlConnection
        Session("isDirty") = myUtility.isDirty(Me.dsCommon, myUtility.getConnection(tempConn))
        If (myUtility.hasResults(Me.dsCommon, myUtility.getConnection(tempConn))) Then
            Session("hasResults") = "True"
        Else
            Session("hasResults") = "False"
        End If

        'Reset the form data to that in the db
        If (Session("hasResults") = "False") Then
            'Delete the Template and all Surveys attached to it on development
            executeDelete()
        Else
            alert("You cannot delete a template that has surveys with results.  De-activate the template if you want to disable further use.")

            'Load fresh from the DB only if we're on our first pass through
            loadData()
        End If

        Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
        Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
        sqlCommonAction.Connection = Me.CommonConn
        sqlCommonAdapter.SelectCommand = sqlCommonAction
        myUtility.determineUserOwnership(sqlCommonAction, sqlCommonAdapter, True)

        'set the page options
        myUtility.optionPreSetTemplate(Session("seqTemplate"), Me.dsCommon.Tables("Templates").Rows.Count(), Me.pageOptions, True)

        Session("PageOptions") = Me.pageOptions

        Session("TransExists") = False
    End Sub

    'Handles the user attempting to do something with the template
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Session("JSAction") = ""

        'Reset the form data to that in the db
        loadData()

        Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
        Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
        sqlCommonAction.Connection = Me.CommonConn
        sqlCommonAdapter.SelectCommand = sqlCommonAction
        myUtility.determineUserOwnership(sqlCommonAction, sqlCommonAdapter, True)

        'set the page options
        myUtility.optionPreSetTemplate(Session("seqTemplate"), Me.dsCommon.Tables("Templates").Rows.Count(), Me.pageOptions, True)

        Session("PageOptions") = Me.pageOptions

        Session("TransExists") = False
    End Sub
#End Region

#Region "Insert Template"
    'executes the insert command
    Private Sub executeInsert()
        Try
            sqlCommonAction.Connection = Me.CommonConn

            Dim [continue] As Boolean = True
            'determine if the user gave this template a name.  If not alert the user and abort all transactions with the db
            If (Trim("" & Me.txtTemplateName.Text) = "") Then
                alert("Failed to add the new template. You must provide a name for your template.")
                [continue] = False
            End If

            'check to see if a template with the same name exists on development
            If (doesTemplateNameExist(sqlCommonAdapter, sqlCommonAction)) Then
                alert("Failed to add the new template. A template already exists with that name.")
                [continue] = False
            End If

            If ([continue]) Then
                'set up the transaction
                myUtility.setupTransaction(sqlCommonAction, Me.CommonConn)

                'Insert new Template
                If ([continue]) Then
                    If Not (InsertTemplate(sqlCommonAdapter, sqlCommonAction)) Then
                        [continue] = False
                    End If
                End If
                If ([continue]) Then
                    If Not (InsertPermissions(sqlCommonAdapter, sqlCommonAction)) Then
                        [continue] = False
                    End If
                End If
                If ([continue]) Then
                    If Not (InsertComments(sqlCommonAdapter, sqlCommonAction)) Then
                        [continue] = False
                    End If
                End If
                If ([continue]) Then
                    If Not (InsertQuestions(sqlCommonAdapter, sqlCommonAction)) Then
                        [continue] = False
                    End If
                End If
                If ([continue]) Then
                    If Not (InsertList(sqlCommonAdapter, sqlCommonAction)) Then
                        [continue] = False
                    End If
                End If
                If ([continue]) Then
                    If Not (InsertReportGroup(sqlCommonAdapter, sqlCommonAction)) Then
                        [continue] = False
                    End If
                End If
                If ([continue]) Then
                    Trace.Warn("Committing transaction")
                    'no errors?  Then we made it far enough to commit the data and we're nearly done.
                    sqlCommonAction.Transaction.Commit()

                    'set the template to the new template and reload the data from the database
                    Session("seqTemplate") = Session("newSeqTemplate")
                    Session("isDirty") = Me.myUtility.isDirty(Me.dsCommon, Me.CommonConn)
                    Session("TemplateName") = Me.txtTemplateName.Text
                    loadData()

                    'change control text
                    changeControlText()
                Else
                    'roll back the transaction
                    sqlCommonAction.Transaction.Rollback()
                    loadData(True)

                    'change control text
                    changeControlText()
                End If
            Else
                loadData(True)

                'change control text
                changeControlText()
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Template insertion failure.")
            loadData(True)

            'change control text
            changeControlText()
        End Try
    End Sub

    'Insert the template
    Private Function InsertTemplate(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand) As Boolean
        Trace.Warn("Inserting the Template")
        Try
            'Set up the insert query
            sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_TEMPLATE (TEMPLATE_NAME, TEMPLATE_DESCRIPTION, TEMPLATE_CREATE_DATE, ACTIVE_IND, " & _
                "CREATOR_USER_KEY, TRAINING_IND, OPTIONAL_IND, CHANGE_IND, HUMAN_SUBJECTS_RESCH_IND) VALUES (@TEMPLATE_NAME, @TEMPLATE_DESCRIPTION" & _
                ", '" & Now & "', " & IIf(Me.chkActive.Checked, 1, 0) & ", " & CType(Session("UserID"), Integer) & ", " & _
                IIf(Me.chkTQSurvey.Checked, 1, 0) & ", " & IIf(Me.chkOptional.Checked, 1, 0) & ", 1, 0)"

            'parameterized the text input to allow for things like quotes to get through
            sqlCommonAction.Parameters(0).Value = Me.txtTemplateName.Text
            Dim param1 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@TEMPLATE_DESCRIPTION", System.Data.SqlDbType.VarChar, 1800, "TEMPLATE_DESCRIPTION")
            param1.Value = Server.HtmlDecode(Me.txtDescription.Text)
            sqlCommonAction.Parameters.Add(param1)

            sqlCommonAction.ExecuteNonQuery()

            'if we make it this far also get the new template seqID
            Dim ds As DataSet = New DataSet
            sqlCommonAction.CommandText = "Select TEMPLATE_KEY from " & myUtility.getExtension() & "SAT_TEMPLATE where TEMPLATE_NAME = @TEMPLATE_NAME" & _
                " AND TEMPLATE_DESCRIPTION = @TEMPLATE_DESCRIPTION AND CREATOR_USER_KEY = " & Session("UserID") & _
                " AND ACTIVE_IND = " & IIf(Me.chkActive.Checked, 1, 0) & " AND TRAINING_IND = " & IIf(Me.chkTQSurvey.Checked, 1, 0) & _
                " AND OPTIONAL_IND = " & IIf(Me.chkOptional.Checked, 1, 0)
            sqlCommonAdapter.Fill(ds)
            Session("newSeqTemplate") = ds.Tables(0).Rows(0).Item("TEMPLATE_KEY")
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to add the new template.")
            Return False
        End Try
    End Function

    'Insert the template permissions
    Private Function InsertPermissions(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand) As Boolean
        Trace.Warn("Setting Template Permissions")

        'Add permissions for the template
        sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_USER_PERMISSION (USER_KEY, TEMPLATE_KEY, SURVEY_KEY, CREATOR_USER_KEY, " & _
            "CREATE_DATE, OWNER_IND, DELEGATE_IND, USER_IND) VALUES (" & Session("UserID") & ", " & Session("newSeqTemplate") & _
            ", -1, " & Session("UserID") & ", '" & Now & "', 1, 1, 1)"
        Trace.Warn(sqlCommonAction.CommandText)
        Try
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to add permissions for the new template.")
            Return False
        End Try
    End Function

    'Insert the template comments
    Private Function InsertComments(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand) As Boolean
        Trace.Warn("Inserting Template Comments")
        'Do only if duplicating an existing template
        If (Session("seqTemplate") <> -1) Then
            'copy the tag items over to the new template
            sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_TEMPLATE_TAG (TEMPLATE_KEY, TEMPLATE_TAG_ID, TEMPLATE_TAG_TITLE, TEMPLATE_TAG_QUESTION) " & _
                "SELECT TEMPLATE_KEY, TEMPLATE_TAG_ID, TEMPLATE_TAG_TITLE, TEMPLATE_TAG_QUESTION FROM " & _
                "(Select ft.TEMPLATE_KEY, fc.TEMPLATE_TAG_ID, fc.TEMPLATE_TAG_TITLE, fc.TEMPLATE_TAG_QUESTION from " & myUtility.getExtension() & "SAT_TEMPLATE_TAG fc, " & myUtility.getExtension() & "SAT_TEMPLATE " & _
                "ft WHERE fc.TEMPLATE_KEY = " & Session("seqTemplate") & " AND ft.TEMPLATE_KEY = " & Session("newSeqTemplate") & ") a"
            Trace.Warn(sqlCommonAction.CommandText)
            Try
                sqlCommonAction.ExecuteNonQuery()
                Return True
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Failed to insert the comments for the new template.")
                Return False
            End Try
        Else
            Return True
        End If
    End Function

    'Insert the template Questions
    Private Function InsertQuestions(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand) As Boolean
        Trace.Warn("Inserting Template Questions")
        'Do only if duplicating an existing template
        If (Session("seqTemplate") <> -1) Then
            'copy the questions over to the new template
            sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION (TEMPLATE_KEY, QUESTION_ID, REQUIRED_IND, PAGE_BREAK_IND, " & _
                "FILTER_IND, NOT_APPLICABLE_IND, QUESTION_TYPE_CODE, QUESTION_TEXT, CHART_TYPE_CODE, RATING_COUNT, RATING_STEP_VALUE, RATING_INITIAL_VALUE, " & _
                "BRANCH_QUESTION_ID) " & _
                "SELECT TEMPLATE_KEY, QUESTION_ID, REQUIRED_IND, PAGE_BREAK_IND, FILTER_IND, NOT_APPLICABLE_IND, QUESTION_TYPE_CODE, " & _
                "QUESTION_TEXT, CHART_TYPE_CODE, RATING_COUNT, RATING_STEP_VALUE, RATING_INITIAL_VALUE, BRANCH_QUESTION_ID FROM " & _
                "(Select ft.TEMPLATE_KEY, fq.QUESTION_ID, fq.REQUIRED_IND, fq.PAGE_BREAK_IND, fq.FILTER_IND, fq.NOT_APPLICABLE_IND, " & _
                "fq.QUESTION_TYPE_CODE, fq.QUESTION_TEXT, fq.CHART_TYPE_CODE, fq.RATING_COUNT, fq.RATING_STEP_VALUE, fq.RATING_INITIAL_VALUE, fq.BRANCH_QUESTION_ID " & _
                "from " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION fq, " & myUtility.getExtension() & "SAT_TEMPLATE ft WHERE fq.TEMPLATE_KEY = " & _
                Session("seqTemplate") & " AND ft.TEMPLATE_KEY = " & Session("newSeqTemplate") & ") a"
            Trace.Warn(sqlCommonAction.CommandText)
            Try
                sqlCommonAction.ExecuteNonQuery()
                Return True
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Failed to insert the questions for the new template.")
                Return False
            End Try
        Else
            Return True
        End If
    End Function

    'Insert the question lists
    Private Function InsertList(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand) As Boolean
        Trace.Warn("Inserting Template Lists for Questions")
        'Do only if duplicating an existing template
        If (Session("seqTemplate") <> -1) Then
            'copy the lists for the questions over to the new template
            sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM (TEMPLATE_KEY, QUESTION_ID, LIST_ITEM_ID, LIST_ITEM_VALUE, LIST_ITEM_TITLE, LIST_ITEM_IMAGE, " & _
                "LIST_ITEM_IMAGE_TYPE, BRANCH_QUESTION_ID) " & _
                "SELECT TEMPLATE_KEY, QUESTION_ID, LIST_ITEM_ID, LIST_ITEM_VALUE, LIST_ITEM_TITLE, LIST_ITEM_IMAGE, LIST_ITEM_IMAGE_TYPE, BRANCH_QUESTION_ID FROM " & _
                "(Select ft.TEMPLATE_KEY, fl.QUESTION_ID, fl.LIST_ITEM_ID, fl.LIST_ITEM_VALUE, fl.LIST_ITEM_TITLE, fl.LIST_ITEM_IMAGE, fl.LIST_ITEM_IMAGE_TYPE, " & _
                "fl.BRANCH_QUESTION_ID from " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM fl, " & myUtility.getExtension() & "SAT_TEMPLATE ft WHERE fl.TEMPLATE_KEY = " & Session("seqTemplate") & _
                " AND ft.TEMPLATE_KEY = " & Session("newSeqTemplate") & ") a"
            Trace.Warn(sqlCommonAction.CommandText)
            Try
                sqlCommonAction.ExecuteNonQuery()
                Return True
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Failed to insert the lists for the questions in the new template.")
                Return False
            End Try
        Else
            Return True
        End If
    End Function

    'Insert the report groups
    Private Function InsertReportGroup(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand) As Boolean
        Trace.Warn("Inserting Template Report Groups for Questions")
        'Do only if duplicating an existing template
        If (Session("seqTemplate") <> -1) Then
            'copy the report groups for the questions over to the new template
            sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION_GROUP (QUESTION_GROUP_ID, QUESTION_GROUP_TITLE, TEMPLATE_KEY) " & _
                "SELECT QUESTION_GROUP_ID, QUESTION_GROUP_TITLE, TEMPLATE_KEY FROM (Select ft.TEMPLATE_KEY, frg.QUESTION_GROUP_TITLE, " & _
                "frg.QUESTION_GROUP_ID from " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION_GROUP frg, " & myUtility.getExtension() & "SAT_TEMPLATE ft WHERE frg.TEMPLATE_KEY = " & _
                Session("seqTemplate") & " AND ft.TEMPLATE_KEY = " & Session("newSeqTemplate") & ") a"
            Trace.Warn(sqlCommonAction.CommandText)
            Try
                sqlCommonAction.ExecuteNonQuery()
                Return True
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Failed to insert the report groups for the questions in the new template.")
                Return False
            End Try
        Else
            Return True
        End If
    End Function
#End Region

#Region "Update Template"
    'executes the update command
    Private Sub executeUpdate()
        Try
            'set up the connection and db interaction tools
            sqlCommonAction.Parameters.Clear()

            Dim [continue] As Boolean = True
            [continue] = myUtility.setupTransaction(sqlCommonAction, Me.CommonConn)

            If ([continue]) Then
                If (Session("Machine") = "Development") Then
                    'determine if the user gave this template a name.  If not alert the user and abort all transactions with the db
                    If (Trim("" & Me.txtTemplateName.Text) = "") Then
                        alert("Failed to update the template. You must provide a name for your template.")
                        [continue] = False
                    End If

                    'check to see if a template with the same name exists on development
                    If (doesTemplateNameExist(sqlCommonAdapter, sqlCommonAction, True)) Then
                        alert("Failed to update the template. A template with the same name in the development database already exists.")
                        [continue] = False
                    End If
                End If

                If ([continue]) Then
                    sqlCommonAction.Connection = Me.CommonConn
                    sqlCommonAction.Parameters.Clear()
                    Me.sqlCommonAdapter.SelectCommand = sqlCommonAction

                    'try to do the update
                    sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE SET TEMPLATE_NAME = @TEMPLATE_NAME, " & _
                    "TEMPLATE_DESCRIPTION = @TEMPLATE_DESCRIPTION, ACTIVE_IND = " & IIf(Me.chkActive.Checked, 1, 0) & ", " & _
                    "TRAINING_IND = " & IIf(Me.chkTQSurvey.Checked, 1, 0) & ", OPTIONAL_IND = " & IIf(Me.chkOptional.Checked, 1, 0) & ", "

                    'some smart updating to avoid dirty bit settings when users are not lvl 4 users
                    If (Session("UserType") < 4 And Session("Results") > 0) Then
                        sqlCommonAction.CommandText &= "CHANGE_IND = 0 WHERE TEMPLATE_KEY = " & Session("seqTemplate")
                    Else
                        If (significantChange()) Then
                            sqlCommonAction.CommandText &= "CHANGE_IND = 1 WHERE TEMPLATE_KEY = " & Session("seqTemplate")
                        Else
                            sqlCommonAction.CommandText &= "CHANGE_IND = 0 WHERE TEMPLATE_KEY = " & Session("seqTemplate")
                        End If
                    End If

                    'parameterized the text input to allow for things like quotes to get through
                    Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@TEMPLATE_NAME", System.Data.SqlDbType.VarChar, 1800, "TEMPLATE_NAME")
                    param0.Value = Me.txtTemplateName.Text
                    sqlCommonAction.Parameters.Add(param0)
                    Dim param1 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@TEMPLATE_DESCRIPTION", System.Data.SqlDbType.VarChar, 1800, "TEMPLATE_DESCRIPTION")
                    param1.Value = Server.HtmlDecode(Me.txtDescription.Text)
                    sqlCommonAction.Parameters.Add(param1)

                    Trace.Warn(sqlCommonAction.CommandText)
                    sqlCommonAction.ExecuteNonQuery()

                    're-load the data from the database
                    sqlCommonAction.Transaction.Commit()
                    Trace.Warn("loading data from update")
                    loadData()

                    'change control text
                    changeControlText()
                Else
                    If (Session("transExists") = True) Then
                        sqlCommonAction.Transaction.Rollback()
                    End If
                    alert("Template update failure.")
                    loadData(True)

                    'change control text
                    changeControlText()
                End If
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Template update failure.")
            loadData(True)

            'change control text
            changeControlText()
        End Try
    End Sub
#End Region

#Region "Delete Template"
    'Executes the deletion of the Template
    Private Function executeDelete(Optional ByVal isCopy As Boolean = False) As Boolean
        Trace.Warn("Executing Delete")
        Try
            'Load fresh from the DB only if we're on our first pass through
            loadData()

            'get the list of surveys for the template
            strSurveyList = getSurveyList()

            'set up the transaction
            myUtility.setupTransaction(sqlCommonAction, Me.CommonConn)

            'do db data manipulation and commit if all manipulations are complete
            Dim blnContinue As Boolean = True
            Try
                'Remove all data pertaining to the template
                'if there are surveys attached to this template remove them and their stuff
                If (strSurveyList <> "") Then
                    If Not (removeTags(sqlCommonAdapter, sqlCommonAction, isCopy)) Then
                        blnContinue = False
                    End If
                    If (blnContinue) Then
                        If Not (removeSurveyPermissions(sqlCommonAdapter, sqlCommonAction, strSurveyList, isCopy)) Then
                            blnContinue = False
                        End If
                    End If
                    If (blnContinue) Then
                        If Not (removeResponses(sqlCommonAdapter, sqlCommonAction, strSurveyList, isCopy)) Then
                            blnContinue = False
                        End If
                    End If
                    If (blnContinue) Then
                        If Not (removeResults(sqlCommonAdapter, sqlCommonAction, strSurveyList, isCopy)) Then
                            blnContinue = False
                        End If
                    End If
                    If (blnContinue) Then
                        If Not (removeSurveyGroupsList(sqlCommonAdapter, sqlCommonAction, strSurveyList, isCopy)) Then
                            blnContinue = False
                        End If
                    End If
                    If (blnContinue) Then
                        If Not (removeSurveyGroups(sqlCommonAdapter, sqlCommonAction, strSurveyList, isCopy)) Then
                            blnContinue = False
                        End If
                    End If
                    If (blnContinue) Then
                        If Not (removeSurveys(sqlCommonAdapter, sqlCommonAction, isCopy)) Then
                            blnContinue = False
                        End If
                    End If
                End If
                'remove everything else related to the template
                If (blnContinue) Then
                    If Not (removeReportGroups(sqlCommonAdapter, sqlCommonAction, isCopy)) Then
                        blnContinue = False
                    End If
                End If
                If (blnContinue) Then
                    If Not (removeListItems(sqlCommonAdapter, sqlCommonAction, isCopy)) Then
                        blnContinue = False
                    End If
                End If
                If (blnContinue) Then
                    If Not (removeQuestions(sqlCommonAdapter, sqlCommonAction, isCopy)) Then
                        blnContinue = False
                    End If
                End If
                If (blnContinue) Then
                    If Not (removeComments(sqlCommonAdapter, sqlCommonAction, isCopy)) Then
                        blnContinue = False
                    End If
                End If
                If (blnContinue) Then
                    If Not (removeTemplatePermissions(sqlCommonAdapter, sqlCommonAction, isCopy)) Then
                        blnContinue = False
                    End If
                End If
                If (blnContinue) Then
                    If Not (removeTemplateSignatures(sqlCommonAdapter, sqlCommonAction, isCopy)) Then
                        blnContinue = False
                    End If
                End If
                If (blnContinue) Then
                    If (removeTemplate(sqlCommonAdapter, sqlCommonAction, isCopy)) Then
                        Trace.Warn("Committing transaction")
                        'no errors?  Then we made it far enough to commit the data and we're nearly done.
                        sqlCommonAction.Transaction.Commit()

                        'notify the survey owners of the loss of their surveys
                        If (Me.dsCommon.Tables.Contains("SurveyNotifyList")) Then
                            If (Me.dsCommon.Tables("SurveyNotifyList").Rows.Count > 0) Then
                                Dim row As DataRow
                                For Each row In Me.dsCommon.Tables("SurveyNotifyList").Rows()
                                    If Not (sendDeleteMail(row.Item("EMAIL_ADDRESS"), row.Item("SURVEY_COMMENT"), True)) Then
                                        alert("Failed to notify " & _
                                        row.Item("FIRST_NAME") & " " & row.Item("LAST_NAME") & _
                                        " about the deletion of the survey called " & _
                                        row.Item("SURVEY_COMMENT") & ".")
                                    End If
                                Next
                            End If
                        End If

                        'notify the template owners of the loss of their template
                        If (Me.dsCommon.Tables.Contains("TemplateNotifyList")) Then
                            If (Me.dsCommon.Tables("TemplateNotifyList").Rows.Count > 0) Then
                                Dim row As DataRow
                                For Each row In Me.dsCommon.Tables("TemplateNotifyList").Rows()
                                    If Not (sendDeleteMail(row.Item("EMAIL_ADDRESS"), row.Item("TEMPLATE_NAME"))) Then
                                        alert("Failed to notify " & _
                                        row.Item("FIRST_NAME") & " " & row.Item("LAST_NAME") & _
                                        " about the deletion of the template called " & _
                                        row.Item("TEMPLATE_NAME") & ".")
                                    End If
                                Next
                            End If
                        End If

                        'set the template to nothing and reload the data from the database
                        If (Not isCopy) Then
                            Session("seqTemplate") = -1
                            loadData()

                            'change control text
                            changeControlText()
                        End If
                        Return True
                    Else
                        'roll back the transaction
                        sqlCommonAction.Transaction.Rollback()
                        loadData(True)

                        'change control text
                        changeControlText()
                        Return False
                    End If
                Else
                    'roll back the transaction
                    sqlCommonAction.Transaction.Rollback()
                    loadData(True)

                    'change control text
                    changeControlText()
                    Return False
                End If
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                sqlCommonAction.Transaction.Rollback()
                loadData(True)

                'change control text
                changeControlText()
                Return False
            End Try
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Template delete failure.")
            loadData(True)

            'change control text
            changeControlText()
        End Try
    End Function

    'removes the tag items for the surveys
    Private Function removeTags(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal isCopy As Boolean) As Boolean
        Trace.Warn("Removing Tags")

        'Remove all Tags for the surveys if there are surveys otherwise return true
        sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_SURVEY_TAG where TEMPLATE_KEY = " & _
        Session("seqTemplate")

        Try
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to remove the tags for the template.")
            Return False
        End Try
    End Function

    'removes the permissions for the surveys
    Private Function removeSurveyPermissions(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal strSurveyList As String, ByVal isCopy As Boolean) As Boolean
        Trace.Warn("Removing Survey Permissions")
        'make a list of all individuals that will need to be notified of the loss of their surveys
        Try
            sqlCommonAction.CommandText = "SELECT fp.USER_KEY, fs.SURVEY_COMMENT, fu.EMAIL_ADDRESS, " & _
            "fu.LAST_NAME, fu.FIRST_NAME, fu.MIDDLE_NAME from " & _
            myUtility.getExtension() & "SAT_USER_PERMISSION fp, " & myUtility.getExtension() & _
            "SAT_TEMPLATE_SURVEY fs, " & myUtility.getExtension() & "SAT_USER fu where fp.USER_KEY = " & _
            "fu.USER_KEY And fs.SURVEY_KEY = fp.SURVEY_KEY And fs.TEMPLATE_KEY = " & _
            Session("seqTemplate") & " and (fp.OWNER_IND = 1 or fp.DELEGATE_IND = 1) order by fp.SURVEY_KEY"
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            sqlCommonAdapter.Fill(Me.dsCommon, "SurveyNotifyList")
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to create a notification list for the survey owners.")
            Return False
        End Try

        If (strSurveyList <> "") Then
            'Remove all permissions for the surveys if there are surveys otherwise return true
            sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_USER_PERMISSION where " & strSurveyList
            Trace.Warn(sqlCommonAction.CommandText)
            Try
                sqlCommonAction.ExecuteNonQuery()
                Return True
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Failed to remove the permissions for the surveys in the template.")
                Return False
            End Try
        Else
            Return True
        End If
    End Function

    'removes the responses for the surveys in the template
    Private Function removeResponses(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal strSurveyList As String, ByVal isCopy As Boolean) As Boolean
        Trace.Warn("Removing survey responses")

        'Remove all responses for the surveys if there are surveys otherwise return true
        sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_RESPONSE where " & strSurveyList

        Try
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to remove the responses for the surveys in this template.")
            Return False
        End Try
    End Function

    'removes the responses for the surveys in the template
    Private Function removeResults(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal strSurveyList As String, ByVal isCopy As Boolean) As Boolean
        Trace.Warn("Removing survey responses")

        'Remove all responses for the surveys if there are surveys otherwise return true
        sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_RESPONSE_RESULT where " & strSurveyList

        Try
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to remove the results for the surveys in this template.")
            Return False
        End Try
    End Function

    'removes the survey groups list for the surveys in the template
    Private Function removeSurveyGroupsList(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal strSurveyList As String, ByVal isCopy As Boolean) As Boolean
        Trace.Warn("Removing survey groups list")

        'Remove all survey groups list for the surveys if there are surveys otherwise return true
        sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_SURVEY_DIST_GRP_LIST where " & strSurveyList

        Try
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to remove the results for the surveys in this template.")
            Return False
        End Try
    End Function

    'removes the survey groups for the surveys in the template
    Private Function removeSurveyGroups(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal strSurveyList As String, ByVal isCopy As Boolean) As Boolean
        Trace.Warn("Removing survey groups")

        'Remove all survey groups list for the surveys if there are surveys otherwise return true
        sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_SURVEY_DIST_GROUP where " & strSurveyList

        Try
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to remove the results for the surveys in this template.")
            Return False
        End Try
    End Function

    'removes the surveys for the template
    Private Function removeSurveys(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal isCopy As Boolean) As Boolean
        Trace.Warn("Removing Surveys")

        'Remove all Surveys for the template
        sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY where TEMPLATE_KEY = " & Session("seqTemplate")

        Try
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to remove the surveys for the template.")
            Return False
        End Try
    End Function

    'removes the list items for the questions in the template
    Private Function removeListItems(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal isCopy As Boolean) As Boolean
        Trace.Warn("Removing List Items")

        'Remove all Surveys for the template
        sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM where TEMPLATE_KEY = " & Session("seqTemplate")

        Try
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to remove the list items for the questions in the template.")
            Return False
        End Try
    End Function

    'removes the report groups for the questions in the template
    Private Function removeReportGroups(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal isCopy As Boolean) As Boolean
        Trace.Warn("Removing Report Groups")

        'Remove all report groups for the template
        sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION_GROUP where TEMPLATE_KEY = " & Session("seqTemplate")

        Try
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to remove the report groups for the questions in the template.")
            Return False
        End Try
    End Function

    'removes the questions for the template
    Private Function removeQuestions(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal isCopy As Boolean) As Boolean
        Trace.Warn("Removing Questions")

        'Remove all Surveys for the template
        sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION where TEMPLATE_KEY = " & Session("seqTemplate")

        Try
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to remove the questions for the template.")
            Return False
        End Try
    End Function

    'removes the comments for the template
    Private Function removeComments(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal isCopy As Boolean) As Boolean
        Trace.Warn("Removing Comments")

        'Remove all Surveys for the template
        sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_TEMPLATE_TAG where TEMPLATE_KEY = " & Session("seqTemplate")

        Try
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to remove the comments for the template.")
            Return False
        End Try
    End Function

    'removes the permissions for the template
    Private Function removeTemplatePermissions(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal isCopy As Boolean) As Boolean
        Trace.Warn("Removing Template Permissions")

        'make a list of all individuals that will need to be notified of the loss of this Template
        Try
            sqlCommonAction.CommandText = "SELECT fp.USER_KEY, ft.TEMPLATE_NAME, fu.EMAIL_ADDRESS, " & _
            "fu.LAST_NAME, fu.FIRST_NAME, fu.MIDDLE_NAME from " & myUtility.getExtension() & "SAT_USER_PERMISSION fp, " & _
            myUtility.getExtension() & "SAT_TEMPLATE ft, " & myUtility.getExtension() & "SAT_USER fu where fp.USER_KEY = fu.USER_KEY " & _
            "And ft.TEMPLATE_KEY = fp.TEMPLATE_KEY And ft.TEMPLATE_KEY = " & _
            Session("seqTemplate") & " order by fp.TEMPLATE_KEY"
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            sqlCommonAdapter.Fill(Me.dsCommon, "TemplateNotifyList")
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to create a notification list for the template owners.")
            Return False
        End Try

        'Remove all permissions for the template
        sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_USER_PERMISSION where TEMPLATE_KEY = " & Session("seqTemplate")

        Try
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to remove the permissions for the template.")
            Return False
        End Try
    End Function

    'removes the the template signatures
    Private Function removeTemplateSignatures(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal isCopy As Boolean) As Boolean
        Trace.Warn("Removing Template Signatures")

        'Remove all signatures for this template.
        sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE where " & _
        " TEMPLATE_KEY = " & Session("seqTemplate")

        Try
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to remove the signatures for this template.")
            Return False
        End Try
    End Function

    'removes the the template
    Private Function removeTemplate(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal isCopy As Boolean) As Boolean
        Trace.Warn("Removing Templates")

        'Remove all Surveys for the template
        sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_TEMPLATE where TEMPLATE_KEY = " & Session("seqTemplate")

        Try
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to remove the template.")
            Return False
        End Try
    End Function
#End Region

#Region "Template Copy"
    'Handles a request to copy or the actual copy Development --> Production.
    Private Sub btnCopy_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCopy.Click
        Session("JSAction") = ""

        'if it's a publication request then send an e-mail to those that need it.
        If (Me.btnCopy.Text = "Publication Request" Or Me.btnCopy.Text = "Re-Submit") Then
            'go to the ADC picker page to find an ADC and then send e-mail to the ADC and HF
            Session("CurrentPage") = "./PubRequest.aspx"
            Session("pageSel") = "PubRequest"
            Response.Redirect("./PubRequest.aspx")
        End If
    End Sub
#End Region

#Region "Template Duplicate"
    'executes the insert command
    Private Sub executeDuplicate()
        Try
            sqlCommonAction.Connection = Me.CommonConn
            Dim [continue] As Boolean = True

            'check for possible malicious code entries
            If Not (myUtility.goodInput(Me.txtTemplateName, True)) Then
                [continue] = False
                alert("Possible malicious code entry for the new template name.  Template duplication aborted.")
            End If

            If ([continue]) Then
                'determine if the user gave this template a name.  If not alert the user and abort all transactions with the db
                If (Trim("" & Me.txtDuplicate) = "") Then
                    alert("Failed to add the new template. You must provide a name for your template.")
                    [continue] = False
                End If

                'check to see if a template with the same name exists on development
                If (doesTemplateNameExist(sqlCommonAdapter, sqlCommonAction, False, True)) Then
                    alert("Failed to add the new template. A template already exists with that name.")
                    [continue] = False
                End If

                If ([continue]) Then
                    'set up the transaction
                    myUtility.setupTransaction(sqlCommonAction, Me.CommonConn)

                    Try
                        'Insert new Template
                        If ([continue]) Then
                            If Not (DuplicateTemplate(sqlCommonAdapter, sqlCommonAction)) Then
                                [continue] = False
                            End If
                        End If
                        If ([continue]) Then
                            If Not (DuplicateTemplatePermissions(sqlCommonAdapter, sqlCommonAction)) Then
                                [continue] = False
                            End If
                        End If
                        If ([continue]) Then
                            If Not (DuplicateComments(sqlCommonAdapter, sqlCommonAction)) Then
                                [continue] = False
                            End If
                        End If
                        If ([continue]) Then
                            If Not (DuplicateQuestions(sqlCommonAdapter, sqlCommonAction)) Then
                                [continue] = False
                            End If
                        End If
                        If ([continue]) Then
                            If Not (DuplicateList(sqlCommonAdapter, sqlCommonAction)) Then
                                [continue] = False
                            End If
                        End If
                        If ([continue]) Then
                            If Not (DuplicateReportGroups(sqlCommonAdapter, sqlCommonAction)) Then
                                [continue] = False
                            End If
                        End If
                        If ([continue]) Then
                            If Not (DuplicateSurveys(sqlCommonAdapter, sqlCommonAction)) Then
                                [continue] = False
                            End If
                        End If
                        If ([continue]) Then
                            strSurveyList = newSurveyList(sqlCommonAdapter, sqlCommonAction)
                        End If
                        If ([continue] And Trim("" & strSurveyList) <> "") Then
                            If ([continue]) Then
                                If Not (DuplicateSurveyPermissions(sqlCommonAdapter, sqlCommonAction)) Then
                                    [continue] = False
                                End If
                            End If
                            If ([continue]) Then
                                If Not (DuplicateTags(sqlCommonAdapter, sqlCommonAction)) Then
                                    [continue] = False
                                End If
                            End If
                            If ([continue]) Then
                                If Not (DuplicateSurveyGroups(sqlCommonAdapter, sqlCommonAction)) Then
                                    [continue] = False
                                End If
                            End If
                            If ([continue]) Then
                                If Not (DuplicateSurveyGroupsList(sqlCommonAdapter, sqlCommonAction)) Then
                                    [continue] = False
                                End If
                            End If
                        End If
                        If ([continue]) Then
                            Trace.Warn("Committing transaction")
                            'no errors?  Then we made it far enough to commit the data and we're nearly done.
                            sqlCommonAction.Transaction.Commit()

                            'set the template to the new template and reload the data from the database
                            Session("seqTemplate") = Session("newSeqTemplate")
                            Session("isDirty") = Me.myUtility.isDirty(Me.dsCommon, Me.CommonConn)
                            Session("TemplateName") = Me.txtDuplicate
                            loadData()

                            'change control text
                            changeControlText()
                            Me.txtDuplicate = ""
                        Else
                            'roll back the transaction
                            sqlCommonAction.Transaction.Rollback()
                            loadData(True)

                            'change control text
                            changeControlText()
                        End If
                    Catch ex As Exception
                        Trace.Warn(ex.ToString)
                        loadData(True)

                        'change control text
                        changeControlText()
                    End Try
                Else
                    loadData(True)

                    'change control text
                    changeControlText()
                End If
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Template duplication failure.")

            'change control text
            changeControlText()
        End Try
    End Sub

    'Insert the template
    Private Function DuplicateTemplate(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand) As Boolean
        Trace.Warn("Inserting the Template")
        Try
            sqlCommonAction.Parameters(0).Value = Me.txtDuplicate

            Dim param1 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@TEMPLATE_DESCRIPTION", System.Data.SqlDbType.VarChar, 1800, "TEMPLATE_DESCRIPTION")
            param1.Value = Server.HtmlDecode(Me.txtDescription.Text)
            sqlCommonAction.Parameters.Add(param1)

            'Set up the insert query
            sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_TEMPLATE (TEMPLATE_NAME, TEMPLATE_DESCRIPTION, TEMPLATE_CREATE_DATE, ACTIVE_IND, " & _
                "CREATOR_USER_KEY, TRAINING_IND, OPTIONAL_IND, CHANGE_IND, HUMAN_SUBJECTS_RESCH_IND) VALUES (@TEMPLATE_NAME, @TEMPLATE_DESCRIPTION," & _
                " '" & Now & "', " & IIf(Me.chkActive.Checked, 1, 0) & ", " & CType(Session("UserID"), Integer) & ", " & _
                IIf(Me.chkTQSurvey.Checked, 1, 0) & ", " & IIf(Me.chkOptional.Checked, 1, 0) & ", 1, 0)"
            Trace.Warn(sqlCommonAction.CommandText)
            sqlCommonAction.ExecuteNonQuery()

            'if we make it this far also get the new template seqID
            Dim ds As DataSet = New DataSet
            sqlCommonAction.CommandText = "Select TEMPLATE_KEY from " & myUtility.getExtension() & "SAT_TEMPLATE where TEMPLATE_NAME = @TEMPLATE_NAME" & _
                " AND TEMPLATE_DESCRIPTION = @TEMPLATE_DESCRIPTION AND CREATOR_USER_KEY = " & Session("UserID") & _
                " AND ACTIVE_IND = " & IIf(Me.chkActive.Checked, 1, 0) & " AND TRAINING_IND = " & IIf(Me.chkTQSurvey.Checked, 1, 0) & _
                " AND OPTIONAL_IND = " & IIf(Me.chkOptional.Checked, 1, 0)
            sqlCommonAdapter.Fill(ds)
            Session("newSeqTemplate") = ds.Tables(0).Rows(0).Item("TEMPLATE_KEY")
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to add the new template.")
            Return False
        End Try
    End Function

    'Insert the template permissions
    Private Function DuplicateTemplatePermissions(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand) As Boolean
        Trace.Warn("Setting Template Permissions")

        'Add permissions for the template
        sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_USER_PERMISSION (USER_KEY, TEMPLATE_KEY, SURVEY_KEY, CREATOR_USER_KEY, " & _
            "CREATE_DATE, OWNER_IND, DELEGATE_IND, USER_IND) VALUES (" & Session("UserID") & ", " & Session("newSeqTemplate") & _
            ", -1, " & Session("UserID") & ", '" & Now & "', 1, 1, 1)"
        Trace.Warn(sqlCommonAction.CommandText)
        Try
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to add permissions for the new template.")
            Return False
        End Try
    End Function

    'Insert the template comments
    Private Function DuplicateComments(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand) As Boolean
        Trace.Warn("Inserting Template Comments")
        'Do only if duplicating an existing template
        If (Session("seqTemplate") <> -1) Then
            'copy the comments over to the new template
            sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_TEMPLATE_TAG (TEMPLATE_KEY, TEMPLATE_TAG_ID, TEMPLATE_TAG_TITLE, TEMPLATE_TAG_QUESTION) " & _
                "SELECT TEMPLATE_KEY, TEMPLATE_TAG_ID, TEMPLATE_TAG_TITLE, TEMPLATE_TAG_QUESTION FROM " & _
                "(Select " & Session("newSeqTemplate") & " as TEMPLATE_KEY, fc.TEMPLATE_TAG_ID, fc.TEMPLATE_TAG_TITLE, fc.TEMPLATE_TAG_QUESTION from " & myUtility.getExtension() & "SAT_TEMPLATE_TAG fc " & _
                "WHERE fc.TEMPLATE_KEY = " & Session("seqTemplate") & ") a"
            Trace.Warn(sqlCommonAction.CommandText)
            Try
                sqlCommonAction.ExecuteNonQuery()
                Return True
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Failed to insert the comments for the new template.")
                Return False
            End Try
        Else
            Return True
        End If
    End Function

    'Insert the template Questions
    Private Function DuplicateQuestions(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand) As Boolean
        Trace.Warn("Inserting Template Questions")
        'Do only if duplicating an existing template
        If (Session("seqTemplate") <> -1) Then
            'copy the questions over to the new template
            sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION (TEMPLATE_KEY, QUESTION_ID, REQUIRED_IND, PAGE_BREAK_IND, " & _
                "FILTER_IND, NOT_APPLICABLE_IND, QUESTION_TYPE_CODE, QUESTION_TEXT, CHART_TYPE_CODE, RATING_COUNT, RATING_STEP_VALUE, RATING_INITIAL_VALUE, BRANCH_QUESTION_ID, QUERY_NAME) " & _
                "SELECT TEMPLATE_KEY, QUESTION_ID, REQUIRED_IND, PAGE_BREAK_IND, FILTER_IND, NOT_APPLICABLE_IND, QUESTION_TYPE_CODE, " & _
                "QUESTION_TEXT, CHART_TYPE_CODE, RATING_COUNT, RATING_STEP_VALUE, RATING_INITIAL_VALUE, BRANCH_QUESTION_ID, QUERY_NAME FROM " & _
                "(Select ft.TEMPLATE_KEY, fq.QUESTION_ID, fq.REQUIRED_IND, fq.PAGE_BREAK_IND, fq.FILTER_IND, fq.NOT_APPLICABLE_IND, " & _
                "fq.QUESTION_TYPE_CODE, fq.QUESTION_TEXT, fq.CHART_TYPE_CODE, fq.RATING_COUNT, fq.RATING_STEP_VALUE, fq.RATING_INITIAL_VALUE, fq.BRANCH_QUESTION_ID, fq.QUERY_NAME " & _
                "from " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION fq, " & myUtility.getExtension() & "SAT_TEMPLATE ft WHERE fq.TEMPLATE_KEY = " & _
                Session("seqTemplate") & " AND ft.TEMPLATE_KEY = " & Session("newSeqTemplate") & ") a"
            Trace.Warn(sqlCommonAction.CommandText)
            Try
                sqlCommonAction.ExecuteNonQuery()
                Return True
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Failed to insert the questions for the new template.")
                Return False
            End Try
        Else
            Return True
        End If
    End Function

    'Insert the template report groups
    Private Function DuplicateReportGroups(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand) As Boolean
        Trace.Warn("Inserting Template Report Groups")
        'Do only if duplicating an existing template
        If (Session("seqTemplate") <> -1) Then
            'copy the report groups over to the new template
            sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION_GROUP (QUESTION_GROUP_ID, QUESTION_GROUP_TITLE, TEMPLATE_KEY) " & _
                "SELECT QUESTION_GROUP_ID, QUESTION_GROUP_TITLE, TEMPLATE_KEY FROM (Select ft.TEMPLATE_KEY, frg.QUESTION_GROUP_ID, " & _
                "frg.QUESTION_GROUP_TITLE from " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION_GROUP frg, " & myUtility.getExtension() & "SAT_TEMPLATE ft WHERE frg.TEMPLATE_KEY = " & _
                Session("seqTemplate") & " AND ft.TEMPLATE_KEY = " & Session("newSeqTemplate") & ") a"
            Trace.Warn(sqlCommonAction.CommandText)
            Try
                sqlCommonAction.ExecuteNonQuery()
                Return True
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Failed to insert the report groups for the new template.")
                Return False
            End Try
        Else
            Return True
        End If
    End Function

    'Insert the question lists
    Private Function DuplicateList(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand) As Boolean
        Trace.Warn("Inserting Template Lists for Questions")
        'Do only if duplicating an existing template
        If (Session("seqTemplate") <> -1) Then
            'copy the lists for the questions over to the new template
            sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM (TEMPLATE_KEY, QUESTION_ID, LIST_ITEM_ID, LIST_ITEM_VALUE, LIST_ITEM_TITLE, LIST_ITEM_IMAGE, " & _
                "LIST_ITEM_IMAGE_TYPE, BRANCH_QUESTION_ID) " & _
                "SELECT TEMPLATE_KEY, QUESTION_ID, LIST_ITEM_ID, LIST_ITEM_VALUE, LIST_ITEM_TITLE, LIST_ITEM_IMAGE, LIST_ITEM_IMAGE_TYPE, BRANCH_QUESTION_ID FROM " & _
                "(Select ft.TEMPLATE_KEY, fl.QUESTION_ID, fl.LIST_ITEM_ID, fl.LIST_ITEM_VALUE, fl.LIST_ITEM_TITLE, fl.LIST_ITEM_IMAGE, fl.LIST_ITEM_IMAGE_TYPE, " & _
                "fl.BRANCH_QUESTION_ID from " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM fl, " & myUtility.getExtension() & "SAT_TEMPLATE ft WHERE fl.TEMPLATE_KEY = " & Session("seqTemplate") & _
                " AND ft.TEMPLATE_KEY = " & Session("newSeqTemplate") & ") a"
            Trace.Warn(sqlCommonAction.CommandText)
            Try
                sqlCommonAction.ExecuteNonQuery()
                Return True
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Failed to insert the lists for the questions in the new template.")
                Return False
            End Try
        Else
            Return True
        End If
    End Function

    'Insert the Surveys
    Private Function DuplicateSurveys(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand) As Boolean
        Trace.Warn("Inserting Template Surveys")
        'Do only if duplicating an existing template
        If (Session("seqTemplate") <> -1) Then
            'copy the lists for the questions over to the new template
            sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY (CREATOR_USER_KEY, TEMPLATE_KEY, SURVEY_CLOSING_DATE, SURVEY_COMMENT, SURVEY_ID, " & _
                " COURSE, OPTIONAL_IND, ACTIVE_IND, SURVEY_COLOR_CODE, SUBMISSION_COUNT, SURVEY_CREATE_DATE) " & _
                "SELECT CREATOR_USER_KEY, TEMPLATE_KEY, SURVEY_CLOSING_DATE, SURVEY_COMMENT, SURVEY_ID, COURSE, OPTIONAL_IND, ACTIVE_IND, SURVEY_COLOR_CODE, " & _
                "SUBMISSION_COUNT, SURVEY_CREATE_DATE FROM " & _
                "(Select fs.CREATOR_USER_KEY, ft.TEMPLATE_KEY, fs.SURVEY_CLOSING_DATE, fs.SURVEY_COMMENT, fs.SURVEY_ID, FS.COURSE, fs.OPTIONAL_IND, " & _
                "fs.ACTIVE_IND, fs.SURVEY_COLOR_CODE, fs.SUBMISSION_COUNT, fs.SURVEY_CREATE_DATE from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY fs, " & myUtility.getExtension() & "SAT_TEMPLATE ft " & _
                "WHERE fs.TEMPLATE_KEY = " & Session("seqTemplate") & " AND ft.TEMPLATE_KEY = " & Session("newSeqTemplate") & ") a"
            Trace.Warn(sqlCommonAction.CommandText)
            Try
                sqlCommonAction.ExecuteNonQuery()
                Return True
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Failed to insert the surveys for the new template.")
                Return False
            End Try
        Else
            Return True
        End If
    End Function

    'Insert the survey permissions
    Private Function DuplicateSurveyPermissions(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand) As Boolean
        Trace.Warn("Inserting Survey Permissions")
        'attempt to insert permissions for each survey
        Try
            Dim dr As DataRow
            For Each dr In Me.dsCommon.Tables("Surveys").Rows
                'Add permissions for the surveys
                sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_USER_PERMISSION (USER_KEY, TEMPLATE_KEY, SURVEY_KEY, CREATOR_USER_KEY, " & _
                    "CREATE_DATE, OWNER_IND, DELEGATE_IND, USER_IND) VALUES (" & Session("UserID") & ", -1, " & _
                    dr.Item("SURVEY_KEY") & ", " & Session("UserID") & ", '" & Now & "', 1, 1, 1)"
                Trace.Warn(sqlCommonAction.CommandText)
                sqlCommonAction.ExecuteNonQuery()
            Next
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to add permissions for the new surveys in the template.")
            Return False
        End Try
    End Function

    'Insert the survey comments
    Private Function DuplicateTags(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand) As Boolean
        Trace.Warn("Inserting Survey Tags")
        'Do only if duplicating an existing template
        Try
            'get a list of surveys old and new
            sqlCommonAction.CommandText = "Select SURVEY_KEY from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY where TEMPLATE_KEY = " & _
            Session("newSeqTemplate") & " order by SURVEY_KEY"
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            sqlCommonAdapter.Fill(Me.dsCommon, "NewSurveys")
            sqlCommonAction.CommandText = "Select SURVEY_KEY from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY where TEMPLATE_KEY = " & _
            Session("SeqTemplate") & " order by SURVEY_KEY"
            sqlCommonAdapter.Fill(Me.dsCommon, "OldSurveys")

            Dim dr As DataRow
            Dim index As Integer = 0
            For Each dr In Me.dsCommon.Tables("NewSurveys").Rows
                If (Session("seqTemplate") <> -1) Then
                    'copy the tag items over to the new template
                    sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_SURVEY_TAG (SURVEY_KEY, TEMPLATE_KEY, SURVEY_TAG_ID, SURVEY_TAG_TITLE, SURVEY_TAG_ANSWER) " & _
                    "SELECT SURVEY_KEY, TEMPLATE_KEY, SURVEY_TAG_ID, SURVEY_TAG_TITLE, SURVEY_TAG_ANSWER FROM " & _
                    "(Select " & dr.Item("SURVEY_KEY") & " as SURVEY_KEY, " & Session("newSeqTemplate") & " as TEMPLATE_KEY, SURVEY_TAG_ID, " & _
                    "SURVEY_TAG_TITLE, SURVEY_TAG_ANSWER from " & myUtility.getExtension() & "SAT_SURVEY_TAG WHERE TEMPLATE_KEY = " & Session("seqTemplate") & _
                    " AND SURVEY_KEY = " & Me.dsCommon.Tables("OldSurveys").Rows(index).Item("SURVEY_KEY") & ") a"
                    Trace.Warn(sqlCommonAction.CommandText)
                    sqlCommonAction.ExecuteNonQuery()
                    index += 1
                End If
            Next
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to insert the tags for the new template.")
            Return False
        End Try
    End Function

    'Insert the survey report groups
    Private Function DuplicateSurveyGroups(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand) As Boolean
        Trace.Warn("Inserting Survey Groups")
        'Do only if duplicating an existing template
        Try
            Dim dr As DataRow
            Dim index As Integer = 0
            For Each dr In Me.dsCommon.Tables("NewSurveys").Rows
                If (Session("seqTemplate") <> -1) Then
                    'copy the tag items over to the new template
                    sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_SURVEY_DIST_GROUP " & _
                    "(SURVEY_KEY, SURVEY_GROUP_ID, GROUP_TITLE) " & _
                    "SELECT SURVEY_KEY, SURVEY_GROUP_ID, GROUP_TITLE FROM " & _
                    "(Select " & dr.Item("SURVEY_KEY") & " as SURVEY_KEY, SURVEY_GROUP_ID, " & _
                    "GROUP_TITLE from " & myUtility.getExtension() & "SAT_SURVEY_DIST_GROUP WHERE SURVEY_KEY = " & _
                    Me.dsCommon.Tables("OldSurveys").Rows(index).Item("SURVEY_KEY") & ") a"
                    Trace.Warn(sqlCommonAction.CommandText)
                    sqlCommonAction.ExecuteNonQuery()
                    index += 1
                End If
            Next
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to insert the survey groups for the new template.")
            Return False
        End Try
    End Function

    'Insert the survey report groups
    Private Function DuplicateSurveyGroupsList(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand) As Boolean
        Trace.Warn("Inserting Survey Groups List")
        'Do only if duplicating an existing template
        Try
            Dim dr As DataRow
            Dim index As Integer = 0
            For Each dr In Me.dsCommon.Tables("NewSurveys").Rows
                If (Session("seqTemplate") <> -1) Then
                    'copy the tag items over to the new template
                    sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_SURVEY_DIST_GRP_LIST " & _
                    "(SURVEY_KEY, SURVEY_GROUP_ID, USER_KEY) " & _
                    "SELECT SURVEY_KEY, SURVEY_GROUP_ID, USER_KEY FROM " & _
                    "(Select " & dr.Item("SURVEY_KEY") & " as SURVEY_KEY, SURVEY_GROUP_ID, " & _
                    "USER_KEY from " & myUtility.getExtension() & "SAT_SURVEY_DIST_GRP_LIST WHERE SURVEY_KEY = " & _
                    Me.dsCommon.Tables("OldSurveys").Rows(index).Item("SURVEY_KEY") & ") a"
                    Trace.Warn(sqlCommonAction.CommandText)
                    sqlCommonAction.ExecuteNonQuery()
                    index += 1
                End If
            Next
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to insert the survey groups list for the new template.")
            Return False
        End Try
    End Function

    'gets the newly created list of surveys
    Private Function newSurveyList(ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand) As String
        sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY where TEMPLATE_KEY = " & Session("newSeqTemplate")
        Trace.Warn(sqlCommonAction.CommandText)
        sqlCommonAdapter.SelectCommand = sqlCommonAction
        sqlCommonAdapter.Fill(Me.dsCommon, "Surveys")
        Return getSurveyList()
    End Function
#End Region

#Region "Send Mail"
    'Sends an e-mail to users with changed passwords, except level 0 users.
    Private Function sendDeleteMail(ByVal Email As String, ByVal strItemName As String, Optional ByVal isSurvey As Boolean = False) As Boolean
        Trace.Warn("Sending user an email")
        'Set up the mail messsage
        Dim strFrom As String = "survey.administrator@pnl.gov"
        Dim strTo As String = Email

        Dim strbldrSubject As New StringBuilder("Survey Admin Tool Template Deletion.")

        'get the current user's name and e-mail
        Try
            Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
            Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            sqlCommonAction.Connection = Me.CommonConn
            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER where USER_KEY = " & Session("UserID")
            sqlCommonAdapter.Fill(Me.dsCommon, "ModifyID")
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try

        If (Me.dsCommon.Tables("ModifyID").Rows.Count < 1) Then
            alert("Critical error getting sender's information.  Unable to locate your information.  You may have been removed from the database in the last few minutes.  Contact the Survey Administrator immediately.")
            Return False
        End If

        Dim strbldrMessage As New StringBuilder
        'notify the user of their status change
        If (isSurvey) Then
            strbldrMessage.Append("Survey: " & strItemName & " in Template: " & Me.txtTemplateName.Text & " has been deleted, by")
        Else
            strbldrMessage.Append("Template " & strItemName & " has been deleted, by")
        End If
        strbldrMessage.Append(" " & Me.dsCommon.Tables("ModifyID").Rows(0).Item("FIRST_NAME") & ", ")
        strbldrMessage.Append(Me.dsCommon.Tables("ModifyID").Rows(0).Item("LAST_NAME") & ".")
        strbldrMessage.Append("  You may contact this person at this e-mail address: <b>")
        strbldrMessage.Append(Me.dsCommon.Tables("ModifyID").Rows(0).Item("EMAIL_ADDRESS") & "</b>")

        If Not (myUtility.genericSendMail(strFrom, strTo, strbldrSubject.ToString, strbldrMessage.ToString)) Then
            Return False
        Else
            Return True
        End If
    End Function

    'Sends an e-mail to users with changed passwords, except level 0 users.
    Private Function sendCopyMail() As Boolean
        Trace.Warn("Sending user an email")
        'Set up the mail messsage
        Dim strFrom As String = "survey.administrator@pnl.gov"

        'get the owners and delegates of this template
        sqlCommonAction.CommandText = "Select fu.EMAIL_ADDRESS from " & myUtility.getExtension() & "SAT_USER_PERMISSION fp, " & myUtility.getExtension() & "SAT_USER fu where " & _
        "fp.TEMPLATE_KEY = " & Session("seqTemplate") & " and fp.USER_KEY = fu.USER_KEY"
        sqlCommonAction.Connection = Me.CommonConn
        sqlCommonAdapter.SelectCommand = sqlCommonAction
        sqlCommonAdapter.Fill(Me.dsCommon, "FromEmail")
        Dim dr As DataRow
        Dim strTo As String
        For Each dr In Me.dsCommon.Tables("FromEmail").Rows
            strTo &= dr.Item("EMAIL_ADDRESS") & ","
        Next

        Dim strbldrSubject As New StringBuilder("Survey Admin Tool Template Publication Notice.")

        'get the current user's name and e-mail
        Try
            Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
            Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            sqlCommonAction.Connection = Me.CommonConn
            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER where USER_KEY = " & Session("UserID")
            sqlCommonAdapter.Fill(Me.dsCommon, "ModifyID")
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to locate your information for sending an e-mail to the owners/delegates of this template.  Please send a courtesy e-mail to them.")
            Return False
        End Try

        If (Me.dsCommon.Tables("ModifyID").Rows.Count < 1) Then
            alert("Failed to locate your information for sending an e-mail to the owners/delegates of this template.  Please send a courtesy e-mail to them.")
            Return False
        End If

        Dim strbldrMessage As New StringBuilder
        strbldrMessage.Append("The template " & Me.txtTemplateName.Text & " has been published, by")
        strbldrMessage.Append(" " & Me.dsCommon.Tables("ModifyID").Rows(0).Item("FIRST_NAME") & " ")
        strbldrMessage.Append(Me.dsCommon.Tables("ModifyID").Rows(0).Item("LAST_NAME") & ".")
        strbldrMessage.Append("  You may begin sending surveys with this template.")

        If Not (myUtility.genericSendMail(strFrom, strTo, strbldrSubject.ToString, strbldrMessage.ToString)) Then
            Return False
        Else
            Return True
        End If
    End Function
#End Region
End Class
