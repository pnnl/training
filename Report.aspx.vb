Imports System.Collections.Specialized
Public Class Report
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Report))
        Me.sqlCommonAdapter = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand
        Me.dsCommon = New System.Data.DataSet
        Me.commonConn = New System.Data.SqlClient.SqlConnection
        Me.SqlInsertCommand = New System.Data.SqlClient.SqlCommand
        Me.SqlUpdateCommand = New System.Data.SqlClient.SqlCommand
        Me.SqlDeleteCommand = New System.Data.SqlClient.SqlCommand
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'sqlCommonAdapter
        '
        Me.sqlCommonAdapter.DeleteCommand = Me.SqlDeleteCommand
        Me.sqlCommonAdapter.InsertCommand = Me.SqlInsertCommand
        Me.sqlCommonAdapter.SelectCommand = Me.SqlSelectCommand1
        Me.sqlCommonAdapter.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_TEMPLATE", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("TEMPLATE_KEY", "TEMPLATE_KEY"), New System.Data.Common.DataColumnMapping("TEMPLATE_NAME", "TEMPLATE_NAME"), New System.Data.Common.DataColumnMapping("TEMPLATE_DESCRIPTION", "TEMPLATE_DESCRIPTION"), New System.Data.Common.DataColumnMapping("TEMPLATE_CREATE_DATE", "TEMPLATE_CREATE_DATE"), New System.Data.Common.DataColumnMapping("ACTIVE_IND", "ACTIVE_IND"), New System.Data.Common.DataColumnMapping("CREATOR_USER_KEY", "CREATOR_USER_KEY"), New System.Data.Common.DataColumnMapping("TRAINING_IND", "TRAINING_IND"), New System.Data.Common.DataColumnMapping("OPTIONAL_IND", "OPTIONAL_IND")})})
        Me.sqlCommonAdapter.UpdateCommand = Me.SqlUpdateCommand
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = "SELECT TEMPLATE_KEY, TEMPLATE_NAME, TEMPLATE_DESCRIPTION, TEMPLATE_CREATE_DATE, A" & _
            "CTIVE_IND, CREATOR_USER_KEY, TRAINING_IND, OPTIONAL_IND FROM VARCONN.SAT_TE" & _
            "MPLATE"
        Me.SqlSelectCommand1.Connection = Me.commonConn
        '
        'dsCommon
        '
        Me.dsCommon.DataSetName = "NewDataSet"
        Me.dsCommon.Locale = New System.Globalization.CultureInfo("en-US")
        '
        'commonConn
        '
        Me.commonConn.FireInfoMessageEventOnUserErrors = False
        '
        'SqlInsertCommand
        '
        Me.SqlInsertCommand.CommandText = resources.GetString("SqlInsertCommand.CommandText")
        Me.SqlInsertCommand.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@TEMPLATE_NAME", System.Data.SqlDbType.VarChar, 0, "TEMPLATE_NAME"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_DESCRIPTION", System.Data.SqlDbType.VarChar, 0, "TEMPLATE_DESCRIPTION"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_CREATE_DATE", System.Data.SqlDbType.DateTime, 0, "TEMPLATE_CREATE_DATE"), New System.Data.SqlClient.SqlParameter("@ACTIVE_IND", System.Data.SqlDbType.Bit, 0, "ACTIVE_IND"), New System.Data.SqlClient.SqlParameter("@CREATOR_USER_KEY", System.Data.SqlDbType.Int, 0, "CREATOR_USER_KEY"), New System.Data.SqlClient.SqlParameter("@TRAINING_IND", System.Data.SqlDbType.Bit, 0, "TRAINING_IND"), New System.Data.SqlClient.SqlParameter("@OPTIONAL_IND", System.Data.SqlDbType.Bit, 0, "OPTIONAL_IND")})
        Me.SqlInsertCommand.Connection = Me.commonConn
        '
        'SqlUpdateCommand
        '
        Me.SqlUpdateCommand.CommandText = resources.GetString("SqlUpdateCommand.CommandText")
        Me.SqlUpdateCommand.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@TEMPLATE_NAME", System.Data.SqlDbType.VarChar, 0, "TEMPLATE_NAME"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_DESCRIPTION", System.Data.SqlDbType.VarChar, 0, "TEMPLATE_DESCRIPTION"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_CREATE_DATE", System.Data.SqlDbType.DateTime, 0, "TEMPLATE_CREATE_DATE"), New System.Data.SqlClient.SqlParameter("@ACTIVE_IND", System.Data.SqlDbType.Bit, 0, "ACTIVE_IND"), New System.Data.SqlClient.SqlParameter("@CREATOR_USER_KEY", System.Data.SqlDbType.Int, 0, "CREATOR_USER_KEY"), New System.Data.SqlClient.SqlParameter("@TRAINING_IND", System.Data.SqlDbType.Bit, 0, "TRAINING_IND"), New System.Data.SqlClient.SqlParameter("@OPTIONAL_IND", System.Data.SqlDbType.Bit, 0, "OPTIONAL_IND"), New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_CREATOR_USER_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CREATOR_USER_KEY", System.Data.DataRowVersion.Original, Nothing)})
        Me.SqlUpdateCommand.Connection = Me.commonConn
        '
        'SqlDeleteCommand
        '
        Me.SqlDeleteCommand.CommandText = "DELETE FROM VARCONN.[SAT_TEMPLATE] WHERE (([TEMPLATE_KEY] = @Original_TE" & _
            "MPLATE_KEY) AND ([CREATOR_USER_KEY] = @Original_CREATOR_USER_KEY))"
        Me.SqlDeleteCommand.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_CREATOR_USER_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CREATOR_USER_KEY", System.Data.DataRowVersion.Original, Nothing)})
        Me.SqlDeleteCommand.Connection = Me.commonConn
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected WithEvents Image1 As System.Web.UI.WebControls.Image
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents chkServerSwitch As System.Web.UI.WebControls.CheckBox
    Protected WithEvents hlpServerSwitch As System.Web.UI.WebControls.Image
    Protected WithEvents oracleProdConn As System.Data.OracleClient.OracleConnection
    Protected WithEvents sqlCommonAction As New System.Data.SqlClient.SqlCommand
    Protected WithEvents sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents dsCommon As System.Data.DataSet
    Protected WithEvents hlpTemplateSelect As System.Web.UI.WebControls.Image
    Protected WithEvents lblTemplateSelectlblExpires As System.Web.UI.WebControls.Label
    Protected WithEvents ddlTemplateList As System.Web.UI.WebControls.DropDownList
    Protected WithEvents hlpSurveySelect As System.Web.UI.WebControls.Image
    Protected WithEvents lblSurveySelect As System.Web.UI.WebControls.Label
    Protected WithEvents lblNoSurveys As System.Web.UI.WebControls.Label
    Protected WithEvents lblNoTemplates As System.Web.UI.WebControls.Label
    Protected WithEvents ddlSurveyList As System.Web.UI.WebControls.DropDownList
    Protected WithEvents rptrTags As System.Web.UI.WebControls.Repeater
    Protected WithEvents Td1 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected myUtility As Utility = New Utility
    Protected WithEvents hlpDateFilter As System.Web.UI.WebControls.Image
    Protected WithEvents wdcStartDate As Infragistics.WebUI.WebSchedule.WebDateChooser
    Protected WithEvents wdcEndDate As Infragistics.WebUI.WebSchedule.WebDateChooser
    Protected WithEvents btnSubmit As System.Web.UI.WebControls.Button
    Protected WithEvents ReportFilter As System.Web.UI.WebControls.Table
    Protected WithEvents ReportData As System.Web.UI.WebControls.Table
    Protected WithEvents hlpSelectAllData As System.Web.UI.WebControls.Image
    Protected WithEvents btnSelectAll As System.Web.UI.WebControls.Button
    Protected WithEvents OutputSettings As System.Web.UI.WebControls.Table
    Protected WithEvents hlpTimeSplit As System.Web.UI.WebControls.Image
    Protected WithEvents ddlTimeSplit As System.Web.UI.WebControls.DropDownList
    Protected WithEvents chkTextOutput As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblPopup As System.Web.UI.WebControls.Label
    Protected WithEvents Table1 As System.Web.UI.WebControls.Table
    Protected WithEvents btnSubmit2 As System.Web.UI.WebControls.Button
    Protected WithEvents btnReset2 As System.Web.UI.WebControls.Button
    Protected WithEvents lblPopup2 As System.Web.UI.WebControls.Label
    Protected WithEvents btnReset As System.Web.UI.WebControls.Button
    Protected WithEvents commonConn As System.Data.SqlClient.SqlConnection
    Protected WithEvents ddlQuestionsPerPage As System.Web.UI.WebControls.DropDownList
    Friend WithEvents SqlDeleteCommand As System.Data.SqlClient.SqlCommand
    Friend WithEvents SqlInsertCommand As System.Data.SqlClient.SqlCommand
    Friend WithEvents SqlUpdateCommand As System.Data.SqlClient.SqlCommand
    Protected WithEvents JSAction As String

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
        Trace.Warn("page init")
        If (Session.IsNewSession) Then
            Session("Alert") = "Your session has timed out for security purposes.  Please log in again."
            Session("CurrentPage") = "./Login.aspx"
            Session("pageSel") = "Login"
            Response.Redirect("./Login.aspx")
        Else
            If (IsPostBack) Then
                'Set the connection based on the machine
                myUtility.getConnection(Me.commonConn)
                reportFilterFill()
                dataFilterFill()
            End If
        End If
    End Sub

#End Region

#Region "Basic"
    'loads the page
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load
        Trace.Warn("Loading Page")
        InitializeComponent()

        'catch incoming alert messages
        alert(Session("Alert"))

        'clear session
        myUtility.clearSession()

        'get the from page
        If Not (IsPostBack) Then
            Session("FromPage") = getFromPage()
            myUtility.getRequest(Page.Request.Url.Query())
        End If

        'Put user code to initialize the page here
        If (Session("isAuthenticated") <> "True") Then
            Session("CurrentPage") = "./Login.aspx"
            Session("pageSel") = "Login"
            Response.Redirect("./Login.aspx")
        ElseIf (myUtility.checkUserType(2)) Then
            redirect(Session("CurrentPage"))
        Else
            Session("CurrentPage") = "./Report.aspx"
            Session("pageSel") = "Report"
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

        'Set the connection based on the machine
        myUtility.getConnection(Me.commonConn)

        'Set the server switch text on initial page load
        If Not (IsPostBack) Then
            Session("seqTemplate") = -1
            Session("seqSurvey") = -1

            'Fill the template and survey drop-downs
            templateSurveyFill()
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
    'Handles filling the survey and template drop-downs for multiple subs
    Private Sub templateSurveyFill()
        'Fill Template List
        If (templateFill()) Then
            'Fill Survey List
            Trace.Warn(Session("TemplatesExist"))
            If (Session("TemplatesExist") = "Yes") Then
                If Not (surveyFill()) Then
                    Session("SurveysExist") = "No"
                End If
            Else
                Session("SurveysExist") = "No"
            End If
        Else
            Session("TemplatesExist") = "No"
        End If
    End Sub

    'Fills the template drop-down list
    Private Function templateFill() As Boolean
        Try
            'Set up the SQL call
            Dim strTempSQL As String = ""
            strTempSQL = "SELECT TEMPLATE_KEY from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY WHERE 0=0 "
            Me.SqlSelectCommand1.CommandText = "SELECT * from " & myUtility.getExtension() & "SAT_TEMPLATE Where 0=0 "
            If (Session("UserType") <= 3) Then
                Me.SqlSelectCommand1.CommandText &= "AND TEMPLATE_KEY in (Select TEMPLATE_KEY from " & myUtility.getExtension() & "SAT_USER_PERMISSION " & _
                "WHERE USER_KEY = " & Session("UserID") & ") "
            End If
            Me.SqlSelectCommand1.CommandText &= " AND TEMPLATE_KEY in (" & strTempSQL & ") ORDER BY TEMPLATE_NAME"
            Trace.Warn(Me.SqlSelectCommand1.CommandText)
            'Execute the SQL Call
            Me.SqlSelectCommand1.Connection = Me.commonConn
            Me.sqlCommonAdapter.SelectCommand = Me.SqlSelectCommand1
            Me.sqlCommonAdapter.Fill(dsCommon, "Templates")

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
                Me.ddlTemplateList.Items.FindByValue(Session("seqTemplate")).Selected = True
                Session("TemplatesExist") = "Yes"
            Else
                Session("TemplatesExist") = "No"
            End If
            Return True
        Catch ex As Exception
            alert("Could not get the templates.")
            Return False
        End Try
    End Function

    'Fills the survey drop-down list
    Private Function surveyFill() As Boolean
        Try
            'Set up the SQL call
            Dim strTempSQL As String = ""
            strTempSQL = "SELECT distinct SURVEY_KEY from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY WHERE 0=0 AND ACTIVE_IND = 1 "
            Me.SqlSelectCommand1.CommandText = "SELECT distinct fs.SURVEY_COMMENT, fs.SURVEY_KEY, fs.TEMPLATE_KEY " & _
                "from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY fs, " & myUtility.getExtension() & "SAT_RESPONSE fr, " & myUtility.getExtension() & "SAT_RESPONSE_RESULT frs WHERE 0=0 AND fs.ACTIVE_IND = 1 " & _
                " and fs.SURVEY_KEY = fr.SURVEY_KEY and fr.SURVEY_KEY = frs.SURVEY_KEY "

            'Make sure only users of level less than 4 can only see their surveys
            If (Session("UserType") <= 3) Then
                Me.SqlSelectCommand1.CommandText &= "AND fs.SURVEY_KEY in (SELECT SURVEY_KEY from " & myUtility.getExtension() & "SAT_USER_PERMISSION " & _
                    "where USER_KEY = " & Session("UserID") & ") "
            End If

            'If a template has been selected then limit the list of surveys based on the template selected
            If (Session("seqTemplate") <> -1) Then
                Me.SqlSelectCommand1.CommandText &= "AND fs.TEMPLATE_KEY = " & Session("seqTemplate") & " "
            End If

            Me.SqlSelectCommand1.CommandText &= " AND fs.SURVEY_KEY in (" & strTempSQL & ") order by fs.SURVEY_COMMENT"
            Trace.Warn(Me.SqlSelectCommand1.CommandText)

            'destroy any existing survey data
            If (Me.dsCommon.Tables.Contains("Surveys")) Then
                Me.dsCommon.Tables("Surveys").Rows.Clear()
            End If

            'Execute the SQL Call
            Me.SqlSelectCommand1.Connection = Me.commonConn
            Me.sqlCommonAdapter.SelectCommand = Me.SqlSelectCommand1
            Me.sqlCommonAdapter.Fill(dsCommon, "Surveys")
            Session("surveyData") = Me.dsCommon.Tables("Surveys")
            If (Me.dsCommon.Tables("Surveys").Rows.Count() > 0) Then
                'Bind the data and add an additional item for no template selected
                Me.ddlSurveyList.DataSource = Me.dsCommon.Tables("Surveys")
                Me.ddlSurveyList.DataValueField = "SURVEY_KEY"
                Me.ddlSurveyList.DataTextField = "SURVEY_COMMENT"
                Me.ddlSurveyList.DataBind()
                Dim li As System.Web.UI.WebControls.ListItem = New System.Web.UI.WebControls.ListItem
                li.Text = "-- Select a Survey --"
                li.Value = "-1"
                Me.ddlSurveyList.Items.Insert(0, li)
                Me.ddlSurveyList.Items.FindByValue(Session("seqSurvey")).Selected = True
                Session("SurveysExist") = "Yes"
            Else
                Session("SurveysExist") = "No"
            End If
            Return True
        Catch ex As Exception
            alert("Could not get the surveys for this template.")
            Return False
        End Try
    End Function

    'Fills the page with the appropriate tags and comments
    Private Function tagFill() As Boolean
        Try
            'Set up the SQL call
            Dim strTempSQL As String = ""
            strTempSQL = "SELECT SURVEY_TAG_TITLE, SURVEY_TAG_ANSWER from " & myUtility.getExtension() & "SAT_SURVEY_TAG ft, " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY fs WHERE ft.SURVEY_KEY = fs.SURVEY_KEY" & _
                " AND ft.SURVEY_KEY = " & Me.ddlSurveyList.SelectedValue
            strTempSQL &= " order by fs.SURVEY_CLOSING_DATE desc"
            Me.SqlSelectCommand1.CommandText = strTempSQL
            Me.SqlSelectCommand1.Connection = Me.commonConn
            Me.sqlCommonAdapter.SelectCommand = Me.SqlSelectCommand1
            Me.sqlCommonAdapter.Fill(dsCommon, "Tags")
            Dim dr As DataRow
            For Each dr In Me.dsCommon.Tables("Tags").Rows()
                If (dr.Item("SURVEY_TAG_ANSWER") = "") Then
                    dr.Item("SURVEY_TAG_ANSWER") = "(No Answer)"
                End If
            Next
            Me.rptrTags.DataSource = Me.dsCommon.Tables("Tags")
            Me.rptrTags.DataBind()
            Return True
        Catch ex As Exception
            alert("Could not get the tag data.")
            Return False
        End Try
    End Function

    'fills in the initial survey data
    Private Function getSurveyData() As Boolean
        Try
            'fill survey stats
            'get the number of times submitted
            Me.sqlCommonAction.CommandText = "select sum(SUBMISSION_COUNT+1) as clickedIn " & _
                                                " from " & myUtility.getExtension() & "SAT_RESPONSE where 0=0 " & _
                                                " and SURVEY_KEY = " & Session("seqSurvey")
            Me.sqlCommonAction.Connection = Me.commonConn
            Me.sqlCommonAdapter.SelectCommand = Me.sqlCommonAction
            If (Me.dsCommon.Tables.Contains("SurveySubmitted")) Then
                Me.dsCommon.Tables("SurveySubmitted").Rows.Clear()
            End If
            Me.sqlCommonAdapter.Fill(Me.dsCommon, "SurveySubmitted")
            Session("ClickedEmailLink") = Me.dsCommon.Tables("SurveySubmitted").Rows(0).Item(0)
            If IsDBNull(Session("ClickedEmailLink")) Then
                Session("ClickedEmailLink") = -1
            End If

            'get the number of completed surveys
            Me.sqlCommonAction.CommandText = "Select " & _
                    "(select count(frs.LAST_COMPLETION_DATE) from " & myUtility.getExtension() & "SAT_RESPONSE frs where frs.SURVEY_KEY = " & _
                    Session("seqSurvey") & " and frs.LAST_COMPLETION_DATE <> '1/1/1970' and LAST_USED_DATE > '1/1/1970')" & _
                    " as Responded from " & myUtility.getExtension() & "SAT_RESPONSE frs where SURVEY_KEY = " & Session("seqSurvey") & _
                    " and LAST_USED_DATE > '1/1/1970'"

            If (Me.dsCommon.Tables.Contains("SurveyCompleted")) Then
                Me.dsCommon.Tables("SurveyCompleted").Rows.Clear()
            End If
            Me.sqlCommonAdapter.Fill(Me.dsCommon, "SurveyCompleted")
            Session("CompletedSurveys") = Me.dsCommon.Tables("SurveyCompleted").Rows(0).Item(0)
            If IsDBNull(Session("CompletedSurveys")) Then
                Session("CompletedSurveys") = -1
            End If

            'get the earliest/latest data dates
            Me.sqlCommonAction.CommandText = "select min(CREATE_DATE) as beginData " & _
                                                " , max(CREATE_DATE) as endData " & _
                                                " from " & myUtility.getExtension() & "SAT_RESPONSE where 0=0 " & _
                                                " and SURVEY_KEY = " & Session("seqSurvey")
            If (Me.dsCommon.Tables.Contains("SurveyDate")) Then
                Me.dsCommon.Tables("SurveyDate").Rows.Clear()
            End If
            Me.sqlCommonAdapter.Fill(Me.dsCommon, "SurveyDate")
            Session("EarliestDate") = Me.dsCommon.Tables("SurveyDate").Rows(0).Item(0)
            If IsDBNull(Session("EarliestDate")) Then
                Session("EarliestDate") = "1/1/1970"
            End If
            Session("LatestDate") = Me.dsCommon.Tables("SurveyDate").Rows(0).Item(1)
            If IsDBNull(Session("LatestDate")) Then
                Session("LatestDate") = "1/1/1970"
            End If
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Unable to get survey data for this template.")
            Return False
        End Try
    End Function

    'gets the data for the filter section - does not include text based data
    Private Function reportFilterFill() As Boolean
        Try
            'define the variables
            Dim genericDDL As System.Web.UI.WebControls.DropDownList
            Dim genericChkBx As System.Web.UI.WebControls.CheckBox
            Dim genericLabel As System.Web.UI.WebControls.Label
            Dim genericCell As System.Web.UI.WebControls.TableCell
            Dim genericRow As System.Web.UI.WebControls.TableRow
            Dim genericItem As System.Web.UI.WebControls.ListItem

            If (Session("seqTemplate") <> -1 And Session("seqSurvey") <> -1) Then
                'eliminate previous data
                Dim index As Integer = Me.ReportFilter.Rows.Count() - 1
                While (index >= 2)
                    Me.ReportFilter.Rows.RemoveAt(index)
                    index -= 1
                End While

                'get all of the filterable items to include 1-line text fields that are marked for viewing at the
                'top of the report
                Me.sqlCommonAction.CommandText = "select * from " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION " & _
                    "where TEMPLATE_KEY = " & Session("seqTemplate") & " and QUESTION_TYPE_CODE in " & _
                    "('D', 'L', 'N', 'P', 'T', 'S', 'Z') and FILTER_IND = 1 ORDER BY QUESTION_ID"
                Trace.Warn(Me.sqlCommonAction.CommandText)
                Me.sqlCommonAction.Connection = Me.commonConn
                Me.sqlCommonAdapter.SelectCommand = Me.sqlCommonAction
                If (Me.dsCommon.Tables.Contains("TemplateQuestions")) Then
                    Me.dsCommon.Tables("TemplateQuestions").Rows.Clear()
                End If
                Me.sqlCommonAdapter.Fill(Me.dsCommon, "TemplateQuestions")
                Dim drTemplateQuestion As DataRow

                'for each question set up the filter fields for it
                Dim blnDoOnce As Boolean = False
                index = 0
                For Each drTemplateQuestion In Me.dsCommon.Tables("TemplateQuestions").Rows()
                    'build the table header
                    If Not (blnDoOnce) Then
                        genericRow = New System.Web.UI.WebControls.TableRow
                        genericRow.Height = Unit.Pixel(40)
                        genericRow.Cells.Clear()
                        genericCell = New System.Web.UI.WebControls.TableCell
                        genericCell.Text = "Question"
                        genericCell.HorizontalAlign = HorizontalAlign.Center
                        genericCell.ColumnSpan = 2
                        genericCell.CssClass = "boldHeader"
                        genericCell.Wrap = False
                        genericCell.EnableViewState = True
                        genericRow.Cells.Add(genericCell)
                        genericCell = New System.Web.UI.WebControls.TableCell
                        genericCell.Text = "Filter"
                        genericCell.HorizontalAlign = HorizontalAlign.Center
                        genericCell.CssClass = "boldHeader"
                        genericCell.Wrap = False
                        genericCell.EnableViewState = True
                        genericRow.Cells.Add(genericCell)
                        genericRow.BackColor = Color.LightGray
                        Me.ReportFilter.Rows.Add(genericRow)
                        blnDoOnce = True
                    End If

                    'get the question list items for building the filter
                    Me.sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM where QUESTION_ID = " & _
                        drTemplateQuestion("QUESTION_ID") & " and TEMPLATE_KEY = " & _
                        Session("seqTemplate") & " Order by LIST_ITEM_ID"
                    If (Me.dsCommon.Tables.Contains("ListItems")) Then
                        Me.dsCommon.Tables("ListItems").Rows.Clear()
                    End If
                    Me.sqlCommonAdapter.Fill(Me.dsCommon, "ListItems")
                    Dim drListItem As DataRow

                    'generate a row for the question text
                    genericRow = New System.Web.UI.WebControls.TableRow
                    genericCell = New System.Web.UI.WebControls.TableCell
                    genericLabel = New System.Web.UI.WebControls.Label
                    genericLabel.CssClass = "boldContent"
                    genericLabel.Text = drTemplateQuestion("QUESTION_TEXT")
                    genericLabel.ID = "QLabel" & drTemplateQuestion("QUESTION_ID")
                    genericCell.Controls.Add(genericLabel)
                    genericCell.Wrap = True
                    genericCell.ColumnSpan = 2
                    genericCell.EnableViewState = True
                    genericRow.Cells.Add(genericCell)

                    If (Me.dsCommon.Tables("ListItems").Rows.Count > 1) Then
                        'generate a new row, cell, and drop-down list control
                        genericCell = New System.Web.UI.WebControls.TableCell
                        genericDDL = New System.Web.UI.WebControls.DropDownList
                        genericDDL.ID = "QDDL^" & drTemplateQuestion("QUESTION_ID")
                        genericDDL.Attributes.Add("QuestionType", drTemplateQuestion("QUESTION_TYPE_CODE"))
                        genericDDL.Width = Unit.Percentage(100)
                        genericCell.Width = Unit.Percentage(25)
                        genericItem = New System.Web.UI.WebControls.ListItem
                        genericItem.Text = "No Filter"
                        genericItem.Value = -1
                        genericDDL.Items.Add(genericItem)

                        'Add the list items to list item controls
                        For Each drListItem In Me.dsCommon.Tables("ListItems").Rows()
                            genericItem = New System.Web.UI.WebControls.ListItem
                            If (Trim("" & drListItem("LIST_ITEM_TITLE")) = "") Then
                                genericItem.Text = drListItem("LIST_ITEM_VALUE")
                                genericItem.Value = drListItem("LIST_ITEM_VALUE")
                            Else
                                genericItem.Text = drListItem("LIST_ITEM_TITLE")
                                genericItem.Value = drListItem("LIST_ITEM_VALUE")
                            End If
                            genericDDL.Items.Add(genericItem)
                        Next

                        genericCell.Controls.Add(genericDDL)
                        genericCell.Wrap = False
                        genericCell.EnableViewState = True
                        genericRow.Cells.Add(genericCell)
                    Else
                        'get the result values for the question since it is not a type of list question
                        If (drTemplateQuestion("QUESTION_TYPE_CODE") = "D") Then
                            Me.sqlCommonAction.CommandText = "SELECT distinct cast(CAST(ANSWER_TEXT as " & _
                            "datetime) as int) as ANSWER_VALUE, ANSWER_TEXT FROM " & myUtility.getExtension() & "SAT_RESPONSE_RESULT " & _
                            "WHERE QUESTION_ID = " & drTemplateQuestion("QUESTION_ID") & " AND SURVEY_KEY = " & _
                            Session("seqSurvey") & " AND QUESTION_TYPE = '" & drTemplateQuestion("QUESTION_TYPE_CODE") & "' ORDER BY ANSWER_TEXT"
                        ElseIf (drTemplateQuestion("QUESTION_TYPE_CODE") = "S") Then
                            Me.sqlCommonAction.CommandText = "Select distinct ANSWER_VALUE, ANSWER_TEXT from " & myUtility.getExtension() & _
                            "SAT_RESPONSE_RESULT where QUESTION_ID = " & drTemplateQuestion("QUESTION_ID") & _
                            " and SURVEY_KEY = " & Session("seqSurvey") & " and ANSWER_TEXT <> '' AND QUESTION_TYPE = '" & _
                            drTemplateQuestion("QUESTION_TYPE_CODE") & "' Order by ANSWER_TEXT"
                        ElseIf (drTemplateQuestion("QUESTION_TYPE_CODE") = "Z") Then
                            Me.sqlCommonAction.CommandText = "Select distinct ANSWER_VALUE, ANSWER_TEXT, " & _
                            "substring(ANSWER_TEXT,9,len(ANSWER_TEXT)-8) as strActualResult from " & myUtility.getExtension() & _
                            "SAT_RESPONSE_RESULT where QUESTION_ID = " & drTemplateQuestion("QUESTION_ID") & _
                            " and SURVEY_KEY = " & Session("seqSurvey") & " and ANSWER_TEXT <> '' AND QUESTION_TYPE = '" & _
                            drTemplateQuestion("QUESTION_TYPE_CODE") & "' Order by strActualResult"
                        ElseIf (drTemplateQuestion("QUESTION_TYPE_CODE") = "N") Then
                            Me.sqlCommonAction.CommandText = "SELECT distinct ANSWER_VALUE, ANSWER_VALUE as ANSWER_TEXT FROM " & myUtility.getExtension() & "SAT_RESPONSE_RESULT " & _
                            "WHERE QUESTION_ID = " & drTemplateQuestion("QUESTION_ID") & " AND SURVEY_KEY = " & _
                            Session("seqSurvey") & " AND QUESTION_TYPE = '" & drTemplateQuestion("QUESTION_TYPE_CODE") & "' ORDER BY ANSWER_VALUE"
                        Else
                            Me.sqlCommonAction.CommandText = "SELECT distinct ANSWER_VALUE, ANSWER_TEXT FROM " & myUtility.getExtension() & "SAT_RESPONSE_RESULT " & _
                            "WHERE QUESTION_ID = " & drTemplateQuestion("QUESTION_ID") & " AND SURVEY_KEY = " & _
                            Session("seqSurvey") & " AND QUESTION_TYPE = '" & drTemplateQuestion("QUESTION_TYPE_CODE") & "' ORDER BY ANSWER_VALUE"
                        End If
                        If (Me.dsCommon.Tables.Contains("ResultSet")) Then
                            Me.dsCommon.Tables("ResultSet").Rows.Clear()
                        End If
                        Me.sqlCommonAdapter.Fill(Me.dsCommon, "ResultSet")
                        Dim drResultSet As DataRow

                        'generate the row, cell, and drop-down list
                        genericCell = New System.Web.UI.WebControls.TableCell
                        genericDDL = New System.Web.UI.WebControls.DropDownList
                        genericDDL.ID = "QDDL^" & drTemplateQuestion("QUESTION_ID")
                        genericDDL.Attributes.Add("QuestionType", drTemplateQuestion("QUESTION_TYPE_CODE"))
                        genericDDL.Width = Unit.Percentage(100)
                        genericCell.Width = Unit.Percentage(25)
                        genericItem = New System.Web.UI.WebControls.ListItem
                        genericItem.Text = "No Filter"
                        genericItem.Value = -1
                        genericDDL.Items.Add(genericItem)

                        'generate the options
                        Dim sTypeValue As String = 0
                        For Each drResultSet In Me.dsCommon.Tables("ResultSet").Rows()
                            genericItem = New System.Web.UI.WebControls.ListItem
                            If (drTemplateQuestion("QUESTION_TYPE_CODE") = "Z") Then
                                Dim myValueArray As Array = CType(drResultSet("ANSWER_TEXT"), String).Split("^")
                                If (myValueArray.Length > 1) Then
                                    genericItem.Text = myValueArray(1)
                                    sTypeValue = myValueArray(0) & "^" & myValueArray(1)
                                Else
                                    genericItem.Text = myValueArray(0)
                                End If
                            Else
                                genericItem.Text = drResultSet("ANSWER_TEXT")
                            End If

                            If (drTemplateQuestion("QUESTION_TYPE_CODE") = "S" _
                                Or drTemplateQuestion("QUESTION_TYPE_CODE") = "Z") Then
                                If (sTypeValue Is Nothing) Then
                                    sTypeValue = 0
                                End If
                                genericItem.Value = sTypeValue
                                'sTypeValue += 1
                            Else
                                genericItem.Value = drResultSet("ANSWER_VALUE")
                            End If
                            genericDDL.Items.Add(genericItem)
                        Next

                        genericCell.Controls.Add(genericDDL)
                        genericCell.Wrap = False
                        genericCell.EnableViewState = True
                        genericRow.Cells.Add(genericCell)
                    End If

                    'set an alternating background color
                    If ((index Mod 2) <> 0) Then
                        genericRow.BackColor = Color.LightGray
                    End If

                    'store the row
                    genericRow.Width = Unit.Percentage(100)
                    genericRow.EnableViewState = True
                    Me.ReportFilter.Rows.Add(genericRow)
                    Me.ReportFilter.EnableViewState = True
                    index += 1
                Next
            End If
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Could not get the report filters.")
            Return False
        End Try
    End Function

    'gets the data for the data section - includes all answerable question types
    Private Function dataFilterFill() As Boolean
        Try
            'define the variables
            Dim genericChkBx As System.Web.UI.WebControls.CheckBox
            Dim genericDDL As System.Web.UI.WebControls.DropDownList
            Dim genericItem As ListItem
            Dim genericLabel As System.Web.UI.WebControls.Label
            Dim genericCell As System.Web.UI.WebControls.TableCell
            Dim genericRow As System.Web.UI.WebControls.TableRow

            If (Session("seqTemplate") <> -1 And Session("seqSurvey") <> -1) Then
                Dim index As Integer = Me.ReportData.Rows.Count() - 1
                While (index >= 1)
                    Me.ReportData.Rows.RemoveAt(index)
                    index -= 1
                End While

                'get the number of times submitted
                Me.sqlCommonAction.CommandText = "select * from " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION " & _
                    "where TEMPLATE_KEY = " & Session("seqTemplate") & " and QUESTION_TYPE_CODE in " & _
                    "('C', 'D', 'L', 'N', 'P', 'R', 'T') ORDER BY QUESTION_ID"
                Me.sqlCommonAction.Connection = Me.commonConn
                Me.sqlCommonAdapter.SelectCommand = Me.sqlCommonAction
                If (Me.dsCommon.Tables.Contains("TemplateQuestions")) Then
                    Me.dsCommon.Tables("TemplateQuestions").Rows.Clear()
                End If
                Me.sqlCommonAdapter.Fill(Me.dsCommon, "TemplateQuestions")
                Dim drTemplateQuestion As DataRow

                'build the chart table
                Me.sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_CHART_TYPE " & _
                    "order by CHART_TYPE_SEQ_NO"
                If (Me.dsCommon.Tables.Contains("ChartTypes")) Then
                    Me.dsCommon.Tables("ChartTypes").Rows.Clear()
                End If
                Me.sqlCommonAdapter.Fill(Me.dsCommon, "ChartTypes")

                'get the user's preferences if they exist
                Me.sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_CHART_PREFERENCE " & _
                    "where TEMPLATE_KEY = " & Session("seqTemplate") & " and PREFERENCE_USER_KEY = " & _
                    Session("UserID") & " order by QUESTION_ID"
                If (Me.dsCommon.Tables.Contains("Preferences")) Then
                    Me.dsCommon.Tables("Preferences").Rows.Clear()
                End If
                Me.sqlCommonAdapter.Fill(Me.dsCommon, "Preferences")
                
                'build the table header
                genericRow = New System.Web.UI.WebControls.TableRow
                genericRow.Height = Unit.Pixel(40)
                genericRow.Cells.Clear()
                genericCell = New System.Web.UI.WebControls.TableCell
                genericCell.Text = "Question"
                genericCell.HorizontalAlign = HorizontalAlign.Center
                genericCell.CssClass = "boldHeader"
                genericCell.ColumnSpan = 2
                genericCell.Wrap = False
                genericCell.EnableViewState = True
                genericRow.Cells.Add(genericCell)
                genericCell = New System.Web.UI.WebControls.TableCell
                genericCell.Text = "Display as Percent?"
                genericCell.HorizontalAlign = HorizontalAlign.Center
                genericCell.CssClass = "boldHeader"
                genericCell.Wrap = False
                genericCell.EnableViewState = True
                genericRow.Cells.Add(genericCell)
                genericCell = New System.Web.UI.WebControls.TableCell
                genericCell.Text = "Chart Type"
                genericCell.HorizontalAlign = HorizontalAlign.Center
                genericCell.CssClass = "boldHeader"
                genericCell.Wrap = False
                genericCell.EnableViewState = True
                genericRow.Cells.Add(genericCell)
                genericRow.BackColor = Color.LightGray
                Me.ReportData.Rows.Add(genericRow)

                'for each question set up the filter fields for it
                index = 0
                For Each drTemplateQuestion In Me.dsCommon.Tables("TemplateQuestions").Rows()
                    'generate a new row, cell, and drop-down list control
                    genericRow = New System.Web.UI.WebControls.TableRow
                    genericRow.Cells.Clear()
                    genericCell = New System.Web.UI.WebControls.TableCell
                    genericChkBx = New System.Web.UI.WebControls.CheckBox
                    genericChkBx.ID = "QChk^" & drTemplateQuestion("QUESTION_ID")

                    'populate the checkbox
                    genericChkBx.Attributes.Add("Value", drTemplateQuestion("QUESTION_TYPE_CODE"))
                    genericChkBx.Attributes.Add("Group", drTemplateQuestion("QUESTION_GROUP_ID"))
                    genericChkBx.CssClass = "boldContent"

                    'add the checkbox
                    genericCell.Controls.Add(genericChkBx)
                    genericCell.Wrap = True
                    genericCell.EnableViewState = True
                    genericRow.Cells.Add(genericCell)

                    'add the question text
                    genericCell = New System.Web.UI.WebControls.TableCell
                    genericLabel = New System.Web.UI.WebControls.Label
                    genericLabel.CssClass = "boldContent"
                    genericLabel.Text = drTemplateQuestion("QUESTION_TEXT")
                    genericLabel.ID = "QChkLabel^" & drTemplateQuestion("QUESTION_ID")

                    'add the label
                    genericCell.Controls.Add(genericLabel)
                    genericCell.Wrap = True
                    genericCell.EnableViewState = True
                    genericRow.Cells.Add(genericCell)

                    'generate the checkbox
                    genericCell = New System.Web.UI.WebControls.TableCell
                    genericCell.HorizontalAlign = HorizontalAlign.Center
                    genericChkBx = New System.Web.UI.WebControls.CheckBox
                    genericChkBx.ID = "PercChkBx^" & drTemplateQuestion("QUESTION_ID")
                    genericChkBx.Text = ""
                    genericChkBx.CssClass = "boldContent"
                    genericChkBx.TextAlign = TextAlign.Right
                    genericChkBx.AutoPostBack = False
                    If (Me.dsCommon.Tables("Preferences").Rows.Count > 0) Then
                        Dim row As DataRow
                        For Each row In Me.dsCommon.Tables("Preferences").Rows
                            If (CType(row.Item("PERCENT_IND"), Boolean) And _
                                row.Item("QUESTION_ID") = drTemplateQuestion("QUESTION_ID")) Then
                                genericChkBx.Checked = True
                            End If
                        Next
                    End If
                    genericCell.Controls.Add(genericChkBx)
                    genericRow.Cells.Add(genericCell)

                    'generete the drop-down list chart override
                    genericCell = New System.Web.UI.WebControls.TableCell
                    genericDDL = New System.Web.UI.WebControls.DropDownList
                    genericDDL.ID = "CDDL^" & drTemplateQuestion("QUESTION_ID")

                    'Add the list items to list item controls
                    Dim drChartItem As DataRow
                    For Each drChartItem In Me.dsCommon.Tables("ChartTypes").Rows()
                        genericItem = New System.Web.UI.WebControls.ListItem
                        genericItem.Text = drChartItem("CHART_TYPE_TITLE")
                        genericItem.Value = drChartItem("CHART_TYPE_CODE")
                        genericDDL.Items.Add(genericItem)
                        If (Me.dsCommon.Tables("Preferences").Rows.Count > 0) Then
                            Dim row As DataRow
                            For Each row In Me.dsCommon.Tables("Preferences").Rows
                                If (drChartItem("CHART_TYPE_CODE") = row.Item("CHART_TYPE_CODE") And _
                                    row.Item("QUESTION_ID") = drTemplateQuestion("QUESTION_ID")) Then
                                    genericDDL.SelectedIndex = genericDDL.Items.Count - 1
                                End If
                            Next
                        Else
                            If (drChartItem("CHART_TYPE_CODE") = drTemplateQuestion.Item("CHART_TYPE_CODE")) Then
                                genericDDL.SelectedIndex = genericDDL.Items.Count - 1
                            End If
                        End If
                    Next

                    genericCell.Controls.Add(genericDDL)
                    genericRow.Cells.Add(genericCell)
                    genericRow.EnableViewState = True

                    'set an alternating background color
                    If ((index Mod 2) <> 0) Then
                        genericRow.BackColor = Color.LightGray
                    End If

                    'add the row and increment the index
                    Me.ReportData.Rows.Add(genericRow)
                    Me.ReportData.EnableViewState = True
                    index += 1
                Next
            End If
            Return True
        Catch ex As Exception
            alert("Could not get the data filters.")
            Return False
        End Try
    End Function

    'refreshes the page data on post-back where user makes a change
    Private Sub pageDataRefresh()
        If (Me.ddlSurveyList.SelectedIndex <> 0) Then
            Session("seqSurvey") = Me.ddlSurveyList.SelectedItem.Value
            Session("seqTemplate") = CType(Session("surveyData"), DataTable).Rows(Me.ddlSurveyList.SelectedIndex - 1).Item("TEMPLATE_KEY")
            If (surveyFill()) Then
                Me.ddlTemplateList.SelectedItem.Selected = False
                Me.ddlTemplateList.Items.FindByValue(Session("seqTemplate")).Selected = True
                If (tagFill()) Then
                    If (getSurveyData()) Then
                        If (reportFilterFill()) Then
                            dataFilterFill()
                        End If
                    End If
                End If
            Else
                Session("SurveysExist") = "No"
            End If
        Else
            Session("seqSurvey") = Me.ddlSurveyList.SelectedItem.Value
        End If
    End Sub
#End Region

#Region "Page Action"
    'Forces a survey Re-Fill
    Private Sub ddlTemplateList_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlTemplateList.SelectedIndexChanged
        Session("JSAction") = ""

        Session("seqTemplate") = Me.ddlTemplateList.SelectedItem.Value()
        Session("seqSurvey") = -1
        If Not (surveyFill()) Then
            Session("SurveysExist") = "No"
        End If
    End Sub

    'Handles the user selecting a survey
    Private Sub ddlSurveyList_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlSurveyList.SelectedIndexChanged
        Session("JSAction") = ""

        If (Me.ddlSurveyList.SelectedIndex <> 0) Then
            Session("seqSurvey") = Me.ddlSurveyList.SelectedItem.Value
            Session("seqTemplate") = CType(Session("surveyData"), DataTable).Rows(Me.ddlSurveyList.SelectedIndex - 1).Item("TEMPLATE_KEY")
            If (surveyFill()) Then
                Me.ddlTemplateList.SelectedItem.Selected = False
                Me.ddlTemplateList.Items.FindByValue(Session("seqTemplate")).Selected = True
                If (tagFill()) Then
                    If (getSurveyData()) Then
                        If (reportFilterFill()) Then
                            dataFilterFill()
                        End If
                    End If
                End If
            Else
                Session("SurveysExist") = "No"
            End If
        Else
            Session("seqSurvey") = Me.ddlSurveyList.SelectedItem.Value
        End If
    End Sub

    'resets the page
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click, btnReset2.Click
        Session("JSAction") = ""

        Me.wdcEndDate.Value = ""
        Me.wdcStartDate.Value = ""
        Me.chkTextOutput.Checked = False
        Me.ddlTimeSplit.SelectedIndex = 0

        pageDataRefresh()
    End Sub

    'sets the autocheck for the data section and stores any changes to the user's choices
    Private Sub btnSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectAll.Click
        Session("JSAction") = ""

        If (Session("reportAutoCheck") = True) Then
            Session("reportAutoCheck") = False
        Else
            Session("reportAutoCheck") = True
        End If

        'clear out the user's chart preferences in preparation for overwriting them
        Me.sqlCommonAction.Connection = Me.commonConn
        Me.sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_CHART_PREFERENCE where TEMPLATE_KEY = " & _
            Session("seqTemplate") & " and PREFERENCE_USER_KEY = " & Session("UserID")
        Me.sqlCommonAction.ExecuteNonQuery()
        
        'get the control section to save
        Dim myControl As Control
        myControl = Page.Controls(0).Controls(1).Controls(0).Controls(15)

        'store the user's changes
        Dim index As Integer = 2
        While (index < myControl.Controls.Count)
            Dim tempChkBx As CheckBox
            Dim tempPercChkBx As CheckBox
            Dim tempChartDDL As DropDownList
            tempChkBx = CType(myControl.Controls(index).Controls(0).Controls(0), CheckBox)
            tempPercChkBx = CType(myControl.Controls(index).Controls(2).Controls(0), CheckBox)
            tempChartDDL = CType(myControl.Controls(index).Controls(3).Controls(0), DropDownList)

            'write the saved choices for the user to the database
            Me.sqlCommonAction.CommandText = "Insert into " & myUtility.getExtension() & "SAT_CHART_PREFERENCE (TEMPLATE_KEY, PREFERENCE_USER_KEY, " & _
                "QUESTION_ID, CHART_TYPE_CODE, PERCENT_IND) values (" & Session("seqTemplate") & ", " & _
                Session("UserID") & ", " & tempChkBx.ID.Substring(tempChkBx.ID.LastIndexOf("^") + 1) & _
                ", '" & tempChartDDL.SelectedValue & "', " & IIf(tempPercChkBx.Checked, 1, 0) & ")"
            Me.sqlCommonAction.ExecuteNonQuery()
            index += 1
        End While

        pageDataRefresh()

        'build the session string for the data items
        index = 2
        While (index < myControl.Controls.Count)
            Dim tempChkBx As CheckBox
            tempChkBx = CType(myControl.Controls(index).Controls(0).Controls(0), CheckBox)

            'update the checkboxes of the dynamic checkbox controls
            tempChkBx.Checked = Session("reportAutoCheck")
            index += 1
        End While
    End Sub

    'submits the page's results for processing to a report output page.
    Private Sub btnSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubmit.Click, btnSubmit2.Click
        'store the hard coded control values in session
        Session("ReportHeader") = Me.ddlTemplateList.SelectedItem.Text
        Session("SurveyHeader") = Me.ddlSurveyList.SelectedItem.Text
        Session("StartDate") = Me.wdcStartDate.Value & ""
        Session("EndDate") = Me.wdcEndDate.Value & ""
        Session("QuestionsPerPage") = Me.ddlQuestionsPerPage.SelectedValue

        'store dynamic filter item control values in session
        Session("FilterItems") = ""
        Dim myControl As Control
        myControl = Page.Controls(0).Controls(1).Controls(0).Controls(11)
        Dim index As Integer = 3
        While (index < myControl.Controls.Count)
            Dim tempDDL As DropDownList
            Dim tempLbl As Label
            Dim course As String
            tempDDL = CType(myControl.Controls(index).Controls(1).Controls(0), DropDownList)
            tempLbl = CType(myControl.Controls(index).Controls(0).Controls(0), Label)
            If (tempDDL.SelectedValue.IndexOf("^") > 0) Then
                course = tempDDL.SelectedValue.Substring(0, tempDDL.SelectedValue.IndexOf("^"))
            Else
                course = tempDDL.SelectedValue
            End If
            Session("FilterItems") &= tempDDL.ID.Substring(tempDDL.ID.LastIndexOf("^") + 1) & "^" & _
                course & "^" & tempLbl.Text.Replace(",", "~") & "^" & _
                tempDDL.Attributes("QuestionType") & "^" & tempDDL.SelectedItem.Text.Replace(",", "~") & ","
            index += 1
        End While
        Session("FilterItems") = CType(Session("FilterItems"), String).Substring(0, CType(Session("FilterItems"), String).Length - 1)

        'store dynamic data item control values in session
        Session("DataItems") = ""
        Dim blnChecked As Boolean = False
        myControl = Page.Controls(0).Controls(1).Controls(0).Controls(15)

        'determine if all of the checkboxes are checked
        index = 2
        While (index < myControl.Controls.Count)
            Dim tempChkBx As CheckBox
            tempChkBx = CType(myControl.Controls(index).Controls(0).Controls(0), CheckBox)

            'make sure the application knows that the user picked some questions so it doesn't pick for the user
            If (tempChkBx.Checked) Then
                blnChecked = True
            End If
            index += 1
        End While

        'if the checkboxes are not checked then check them
        If Not (blnChecked) Then
            index = 2
            While (index < myControl.Controls.Count)
                CType(myControl.Controls(index).Controls(0).Controls(0), CheckBox).Checked = True
                index += 1
            End While
        End If

        'clear out the user's chart preferences in preparation for overwriting them
        Me.sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_CHART_PREFERENCE where TEMPLATE_KEY = " & _
            Session("seqTemplate") & " and PREFERENCE_USER_KEY = " & Session("UserID")
        Me.sqlCommonAction.ExecuteNonQuery()
        
        'build the session string for the data items
        index = 2
        While (index < myControl.Controls.Count)
            Dim tempChkBx As CheckBox
            Dim tempChkBxLabel As Label
            Dim tempPercChkBx As CheckBox
            Dim tempChartDDL As DropDownList
            tempChkBx = CType(myControl.Controls(index).Controls(0).Controls(0), CheckBox)
            tempChkBxLabel = CType(myControl.Controls(index).Controls(1).Controls(0), Label)
            tempPercChkBx = CType(myControl.Controls(index).Controls(2).Controls(0), CheckBox)
            tempChartDDL = CType(myControl.Controls(index).Controls(3).Controls(0), DropDownList)
            Session("DataItems") &= tempChkBx.ID.Substring(tempChkBx.ID.LastIndexOf("^") + 1) & "^" & _
                                        tempChkBx.Checked & "^" & tempChkBx.Attributes.Item("Value") & _
                                        "^" & tempChkBxLabel.Text.Replace(",", "~") & "^P-" & _
                                        tempPercChkBx.Checked & "^" & tempChkBx.Attributes.Item("Group") & _
                                        "^" & tempChartDDL.SelectedValue & ","

            'write the chart choice for the user to the database
            Me.sqlCommonAction.CommandText = "Insert into " & myUtility.getExtension() & "SAT_CHART_PREFERENCE (TEMPLATE_KEY, PREFERENCE_USER_KEY, " & _
                "QUESTION_ID, CHART_TYPE_CODE, PERCENT_IND) values (" & Session("seqTemplate") & ", " & _
                Session("UserID") & ", " & tempChkBx.ID.Substring(tempChkBx.ID.LastIndexOf("^") + 1) & _
                ", '" & tempChartDDL.SelectedValue & "', " & IIf(tempPercChkBx.Checked, 1, 0) & ")"
            Me.sqlCommonAction.ExecuteNonQuery()
            index += 1
        End While
        Session("DataItems") = CType(Session("DataItems"), String).Substring(0, CType(Session("DataItems"), String).Length - 1)

        'store the output options into session
        Session("SplitOutput") = Me.ddlTimeSplit.SelectedValue
        Session("TextOutput") = Me.chkTextOutput.Checked

        'brings up a pop-up window with the report output in it 
        Session("JSAction") = "<script>newWin = window.open('./ReportOutput.aspx', 'ReportOutput', 'width=800,height=600,scrollbars=yes,toolbar=yes,titlebar=no,status=no,resizable=yes,menubar=yes');newWin.focus();</script>"
    End Sub
#End Region
End Class