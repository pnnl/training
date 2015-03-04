Imports System.text
Imports System.IO
Imports System.Collections.Specialized
Public Class SurveyResponse
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SurveyResponse))
        Me.sqlQuestions = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand
        Me.sqlCommonAdapter = New System.Data.SqlClient.SqlDataAdapter
        Me.dsCommon = New System.Data.DataSet
        Me.commonConn = New System.Data.SqlClient.SqlConnection
        Me.sqlCommonAction = New System.Data.SqlClient.SqlCommand
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'sqlQuestions
        '
        Me.sqlQuestions.SelectCommand = Me.SqlSelectCommand1
        Me.sqlQuestions.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_QUESTION_LIST_ITEM", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("TEMPLATE_KEY", "TEMPLATE_KEY"), New System.Data.Common.DataColumnMapping("QUESTION_ID", "QUESTION_ID"), New System.Data.Common.DataColumnMapping("REQUIRED_IND", "REQUIRED_IND"), New System.Data.Common.DataColumnMapping("PAGE_BREAK_IND", "PAGE_BREAK_IND"), New System.Data.Common.DataColumnMapping("QUESTION_TYPE_CODE", "QUESTION_TYPE_CODE"), New System.Data.Common.DataColumnMapping("QUESTION_TEXT", "QUESTION_TEXT"), New System.Data.Common.DataColumnMapping("RATING_COUNT", "RATING_COUNT"), New System.Data.Common.DataColumnMapping("RATING_STEP_VALUE", "RATING_STEP_VALUE"), New System.Data.Common.DataColumnMapping("RATING_INITIAL_VALUE", "RATING_INITIAL_VALUE"), New System.Data.Common.DataColumnMapping("QUERY_NAME", "QUERY_NAME"), New System.Data.Common.DataColumnMapping("Expr1", "Expr1"), New System.Data.Common.DataColumnMapping("Expr2", "Expr2"), New System.Data.Common.DataColumnMapping("LIST_ITEM_ID", "LIST_ITEM_ID"), New System.Data.Common.DataColumnMapping("LIST_ITEM_TITLE", "LIST_ITEM_TITLE"), New System.Data.Common.DataColumnMapping("LIST_ITEM_IMAGE", "LIST_ITEM_IMAGE"), New System.Data.Common.DataColumnMapping("LIST_ITEM_IMAGE_TYPE", "LIST_ITEM_IMAGE_TYPE"), New System.Data.Common.DataColumnMapping("BRANCH_QUESTION_ID", "BRANCH_QUESTION_ID"), New System.Data.Common.DataColumnMapping("CHART_TYPE_CODE", "CHART_TYPE_CODE"), New System.Data.Common.DataColumnMapping("FILTER_IND", "FILTER_IND"), New System.Data.Common.DataColumnMapping("NOT_APPLICABLE_IND", "NOT_APPLICABLE_IND"), New System.Data.Common.DataColumnMapping("Expr3", "Expr3"), New System.Data.Common.DataColumnMapping("LIST_ITEM_VALUE", "LIST_ITEM_VALUE")})})
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
        'commonConn
        '
        Me.commonConn.FireInfoMessageEventOnUserErrors = False
        '
        'commonConn
        '
        Me.commonConn.ConnectionString = "Data Source=OLTPDEV02,915;Initial Catalog=TQDB_DEV;Persist Security Info=True;Use" & _
            "r ID=tqadm;Password=cas2tq"
        Me.commonConn.FireInfoMessageEventOnUserErrors = False
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents lblNoQuestions As System.Web.UI.WebControls.Label
    Protected WithEvents lblQuestions As System.Web.UI.WebControls.Label
    Protected WithEvents QuestionTable As System.Web.UI.WebControls.Table
    Protected WithEvents sqlQuestions As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents dsCommon As System.Data.DataSet
    Protected WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Protected myUtility As Utility = New Utility
    Protected WithEvents imgSurveyResponseTitle As System.Web.UI.WebControls.Image
    Protected WithEvents btnSubmit As System.Web.UI.WebControls.Button
    Protected WithEvents Image1 As System.Web.UI.WebControls.Image
    Protected WithEvents hlpServerSwitch As System.Web.UI.WebControls.Image
    Protected WithEvents chkServerSwitch As System.Web.UI.WebControls.CheckBox
    Protected WithEvents hlpTemplateSelect As System.Web.UI.WebControls.Image
    Protected WithEvents lblTemplateSelectlblExpires As System.Web.UI.WebControls.Label
    Protected WithEvents ddlTemplateList As System.Web.UI.WebControls.DropDownList
    Protected WithEvents hlpSurveySelect As System.Web.UI.WebControls.Image
    Protected WithEvents lblSurveySelect As System.Web.UI.WebControls.Label
    Protected WithEvents ddlSurveyList As System.Web.UI.WebControls.DropDownList
    Protected WithEvents lblNoSurveys As System.Web.UI.WebControls.Label
    Protected WithEvents lblNoTemplates As System.Web.UI.WebControls.Label
    Protected WithEvents rptrTags As System.Web.UI.WebControls.Repeater
    Protected WithEvents Td1 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents btnReset As System.Web.UI.WebControls.Button
    Protected WithEvents sqlCommonAction As System.Data.SqlClient.SqlCommand
    Protected WithEvents sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents hlpCompletedDate As System.Web.UI.WebControls.Image
    Protected WithEvents lblCompletedDate As System.Web.UI.WebControls.Label
    Protected WithEvents hlpInputSwitch As System.Web.UI.WebControls.Image
    Protected WithEvents chkInputSwitch As System.Web.UI.WebControls.CheckBox
    Protected WithEvents hlpNonRespondents As System.Web.UI.WebControls.Image
    Protected WithEvents lblNonRespondents As System.Web.UI.WebControls.Label
    Protected WithEvents wneNonRespondent As Infragistics.WebUI.WebDataInput.WebNumericEdit
    Protected WithEvents btnSubmit2 As System.Web.UI.WebControls.Button
    Protected WithEvents btnReset2 As System.Web.UI.WebControls.Button
    Protected WithEvents wdcCompleteDate As Infragistics.WebUI.WebSchedule.WebDateChooser
    Protected WithEvents commonConn As System.Data.SqlClient.SqlConnection
    Protected WithEvents JSAction As String

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
        If (Session.IsNewSession) Then
            Session("Alert") = "Your session has timed out for security purposes.  Please log in again."
            Session("CurrentPage") = "./Login.aspx"
            Session("pageSel") = "Login"
            Response.Redirect("./Login.aspx")
        Else
            If (IsPostBack) Then
                myUtility.getConnection(Me.commonConn)
                questionFill()
            End If
        End If
    End Sub

#End Region

#Region "Basic"
    'loads the page
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load
        Trace.Warn("Loading Page")

        'catch incoming alert messages
        alert(Session("Alert"))

        'clear session
        myUtility.clearSession()

        'if it's a new session store an alert because we'll need it in a sec
        If (Session.IsNewSession And Session("isAuthenticated") Is Nothing) Then
            Session("Alert") = "Your session has timed out for security purposes.  Please log in again."
        End If

        'get anything on the request URL
        If Not (IsPostBack) Then
            myUtility.getRequest(Page.Request.Url.Query())
        End If

        'set the default background color if no color has been selected
        If (Session("BackgroundColor") = "-1" Or Session("BackgroundColor") = "") Then
            Session("BackgroundColor") = "#0000FF"
        End If

        'Put user code to initialize the page here
        If (Session("isAuthenticated") <> "True") Then
            Session("CurrentPage") = "./Login.aspx"
            Session("pageSel") = "Login"
            Response.Redirect("./Login.aspx")
        ElseIf (myUtility.checkUserType(2)) Then
            redirect(Session("CurrentPage"))
        Else
            Session("CurrentPage") = "./SurveyResponse.aspx"
            Session("pageSel") = "SurveyInput"
        End If

        'if the user's password has expired send them to the password reset page
        If (Session("isExpired") = "True") Then
            Session("CurrentPage") = "./Password.aspx"
            Session("pageSel") = "Password"
            redirect(Session("CurrentPage"))
        End If

        'Set the connection based on the machine
        myUtility.getConnection(Me.commonConn)

        If (Session("seqTemplate") Is Nothing) Then
            Session("seqTemplate") = -1
        End If

        If (Session("seqSurvey") Is Nothing) Then
            Session("seqSurvey") = -1
        End If

        If (Session("InputSwitch") Is Nothing) Then
            Session("InputSwitch") = "Respondent"
        End If

        'Set the server switch text on initial page load
        If Not (IsPostBack) Then
            Session("seqTemplate") = -1
            Session("seqSurvey") = -1

            If (Session("InputSwitch") = "Respondent") Then
                Me.chkInputSwitch.Text = "Check for non-respondent input."
            Else
                Me.chkInputSwitch.Text = "Check for respondent input."
            End If

            'Fill the template and survey drop-downs
            loadData()

            Me.btnSubmit.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
            Me.btnSubmit2.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
            Me.btnReset.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
            Me.btnReset2.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
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
    'Initializes the question data on the page
    Private Sub loadData()
        Trace.Warn("Entering InitData")

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
            If (Session("UserType") = 4) Then
                strTempSQL = "SELECT Distinct ft.TEMPLATE_KEY, ft.TEMPLATE_NAME " & _
                    "from " & myUtility.getExtension() & "SAT_TEMPLATE ft, " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE fss " & _
                    "Where 0 = 0 AND ft.TEMPLATE_KEY = fss.TEMPLATE_KEY AND fss.CURRENT_IND = 1 " & _
                    "AND fss.SIGNED_IND = 1 AND ft.ACTIVE_IND = 1 AND ft.CHANGE_IND = 0 " & _
                    "GROUP BY ft.TEMPLATE_KEY, ft.TEMPLATE_NAME HAVING count(*) = 3 " & _
                    "ORDER BY TEMPLATE_NAME"
            Else
                strTempSQL = "SELECT Distinct ft.TEMPLATE_KEY, ft.TEMPLATE_NAME " & _
                   "from " & myUtility.getExtension() & "SAT_TEMPLATE ft, " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE fss, " & _
                   myUtility.getExtension() & "SAT_USER_PERMISSION fp " & _
                   "Where 0 = 0 AND ft.TEMPLATE_KEY = fss.TEMPLATE_KEY AND fss.TEMPLATE_KEY = fp.TEMPLATE_KEY " & _
                   "AND fss.CURRENT_IND = 1 AND fss.SIGNED_IND = 1 AND ft.ACTIVE_IND = 1 AND " & _
                   "ft.CHANGE_IND = 0 AND fp.USER_KEY = " & Session("UserID") & _
                   " AND fp.USER_IND = 1 " & _
                   " GROUP BY ft.TEMPLATE_KEY, ft.TEMPLATE_NAME HAVING count(*) = 3 " & _
                   "ORDER BY TEMPLATE_NAME"
            End If

            Me.sqlCommonAction.CommandText = strTempSQL
            Trace.Warn(Me.sqlCommonAction.CommandText)
            'Execute the SQL Call
            Me.sqlCommonAction.Connection = Me.commonConn
            Me.sqlCommonAdapter.SelectCommand = Me.sqlCommonAction
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
            Trace.Warn(ex.ToString)
            alert("Could not get the templates.")
            Return False
        End Try
    End Function

    'Fills the survey drop-down list
    Private Function surveyFill() As Boolean
        Try
            'Set up the SQL call
            Dim strTempSQL As String = ""

            If (Session("seqTemplate") = -1) Then
                'there is no chosen template yet so go ahead and load all surveys applicable to the user
                If (Session("UserType") = 4) Then
                    strTempSQL = "select * from (select fs.SURVEY_KEY, fs.SURVEY_COMMENT, fs.TEMPLATE_KEY from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY fs, " & _
                    myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE fss, " & myUtility.getExtension() & "SAT_TEMPLATE ft where fs.TEMPLATE_KEY = fss.TEMPLATE_KEY and " & _
                    "fss.TEMPLATE_KEY = ft.TEMPLATE_KEY and fss.CURRENT_IND = 1 and fss.SIGNED_IND = 1 and ft.ACTIVE_IND = 1 " & _
                    "and fs.ACTIVE_IND = 1 and ft.CHANGE_IND = 0 and not exists (select * from " & myUtility.getExtension() & "SAT_SURVEY_TAG fts where fts.SURVEY_KEY = fs.SURVEY_KEY " & _
                    "and rtrim(ltrim(SURVEY_TAG_ANSWER)) = '') " & _
                    "group by fs.TEMPLATE_KEY, fs.SURVEY_KEY, fs.SURVEY_COMMENT " & _
                    "having count(*) = 3) a order by SURVEY_COMMENT"
                Else
                    strTempSQL = "select * from (select fs.SURVEY_KEY, fs.SURVEY_COMMENT, fs.TEMPLATE_KEY from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY fs, " & _
                    myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE fss, " & myUtility.getExtension() & "SAT_TEMPLATE ft, " & myUtility.getExtension() & "SAT_USER_PERMISSION fp where " & _
                    "fs.TEMPLATE_KEY = fss.TEMPLATE_KEY and fs.SURVEY_KEY = fp.SURVEY_KEY and fp.USER_KEY = " & Session("UserID") & _
                    " and (fp.OWNER_IND = 1 or fp.DELEGATE_IND = 1) and fss.TEMPLATE_KEY = ft.TEMPLATE_KEY and fss.CURRENT_IND = 1 " & _
                    "and fss.SIGNED_IND = 1 and ft.ACTIVE_IND = 1 " & _
                    "and fs.ACTIVE_IND = 1 and ft.CHANGE_IND = 0 and not exists (select * from " & myUtility.getExtension() & "SAT_SURVEY_TAG fts where fts.SURVEY_KEY = fs.SURVEY_KEY " & _
                    "and rtrim(ltrim(SURVEY_TAG_ANSWER)) = '') group by fs.TEMPLATE_KEY, fs.SURVEY_KEY, fs.SURVEY_COMMENT " & _
                    "having count(*) = 3) a order by SURVEY_COMMENT"
                End If
            Else
                'a template has been chosen so load only those surveys applicable to the chosen template
                If (Session("UserType") = 4) Then
                    strTempSQL = "select * from (select fs.SURVEY_KEY, fs.SURVEY_COMMENT, fs.TEMPLATE_KEY from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY fs, " & _
                    myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE fss, " & myUtility.getExtension() & "SAT_TEMPLATE ft where fs.TEMPLATE_KEY = fss.TEMPLATE_KEY and " & _
                    "fss.TEMPLATE_KEY = ft.TEMPLATE_KEY and fss.CURRENT_IND = 1 and fss.SIGNED_IND = 1 and ft.ACTIVE_IND = 1 " & _
                    "and fs.ACTIVE_IND = 1 and ft.CHANGE_IND = 0 and not exists (select * from " & myUtility.getExtension() & "SAT_SURVEY_TAG fts where fts.SURVEY_KEY = fs.SURVEY_KEY " & _
                    "and rtrim(ltrim(SURVEY_TAG_ANSWER)) = '') and fs.TEMPLATE_KEY = " & Session("seqTemplate") & _
                    " group by fs.TEMPLATE_KEY, fs.SURVEY_KEY, fs.SURVEY_COMMENT " & _
                    "having count(*) = 3) a order by SURVEY_COMMENT"
                Else
                    strTempSQL = "select * from (select fs.SURVEY_KEY, fs.SURVEY_COMMENT, fs.TEMPLATE_KEY from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY fs, " & _
                    myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE fss, " & myUtility.getExtension() & "SAT_TEMPLATE ft, " & myUtility.getExtension() & "SAT_USER_PERMISSION fp where " & _
                    "fs.TEMPLATE_KEY = fss.TEMPLATE_KEY and fs.SURVEY_KEY = fp.SURVEY_KEY and fp.USER_KEY = " & Session("UserID") & _
                    " and (fp.OWNER_IND = 1 or fp.DELEGATE_IND = 1) and fss.TEMPLATE_KEY = ft.TEMPLATE_KEY and fss.CURRENT_IND = 1 " & _
                    "and fss.SIGNED_IND = 1 and ft.ACTIVE_IND = 1 and fs.TEMPLATE_KEY = " & Session("seqTemplate") & _
                    " and fs.ACTIVE_IND = 1 and ft.CHANGE_IND = 0 and not exists (select * from " & myUtility.getExtension() & "SAT_SURVEY_TAG fts where fts.SURVEY_KEY = fs.SURVEY_KEY " & _
                    "and rtrim(ltrim(SURVEY_TAG_ANSWER)) = '') group by fs.TEMPLATE_KEY, fs.SURVEY_KEY, fs.SURVEY_COMMENT " & _
                    "having count(*) = 3) a order by SURVEY_COMMENT"
                End If
            End If

            Me.sqlCommonAction.CommandText = strTempSQL
            Trace.Warn(Me.sqlCommonAction.CommandText)

            'destroy any existing survey data
            If (Me.dsCommon.Tables.Contains("Surveys")) Then
                Me.dsCommon.Tables("Surveys").Rows.Clear()
            End If

            'Execute the SQL Call
            Me.sqlCommonAction.Connection = Me.commonConn
            Me.sqlCommonAdapter.SelectCommand = Me.sqlCommonAction
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
            Trace.Warn(ex.ToString)
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
            Me.sqlCommonAction.CommandText = strTempSQL
            Me.sqlCommonAction.Connection = Me.commonConn
            Me.sqlCommonAdapter.SelectCommand = Me.sqlCommonAction
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
            Trace.Warn(ex.ToString)
            alert("Could not get the tag data.")
            Return False
        End Try
    End Function

    'Gets the questions for filling out the survey
    Private Function questionFill() As Boolean
        Try
            'clear out the old data and refill from DB
            Page.Controls(0).Controls(1).Controls(0).Controls(14).Controls.Clear()
            If (Me.QuestionTable.Rows.Count() > 0) Then
                Me.QuestionTable.Rows.Clear()
            End If

            Trace.Warn(Me.SqlSelectCommand1.CommandText)
            Me.sqlCommonAction.CommandText = Me.SqlSelectCommand1.CommandText & _
                " WHERE fq.TEMPLATE_KEY = " & Session("seqTemplate") & _
                " ORDER BY fq.QUESTION_ID, fl.LIST_ITEM_ID"
            Me.sqlCommonAction.CommandText = Me.sqlCommonAction.CommandText.Replace("VARCONN.", myUtility.getExtension())
            Me.sqlCommonAction.Connection = Me.commonConn
            Me.sqlQuestions.SelectCommand = Me.sqlCommonAction
            Trace.Warn(Me.sqlCommonAction.CommandText)

            'clear out any data that might be in the table
            If (Me.dsCommon.Tables.Contains("Questions")) Then
                Me.dsCommon.Tables("Questions").Rows.Clear()
            End If

            'get the question data
            Me.sqlQuestions.Fill(Me.dsCommon, "Questions")
            Dim dt As DataTable = Me.dsCommon.Tables("Questions")

            Dim cell As System.Web.UI.WebControls.TableCell
            Dim strQuestionType As String
            Dim index = 0
            Dim visibleQuestionNumber As Integer = 1
            If (dt.Rows.Count() > 0) Then
                While (index < dt.Rows.Count())
                    Dim row As New System.Web.UI.WebControls.TableRow
                    strQuestionType = dt.Rows(index).Item("QUESTION_TYPE_CODE")

                    'process the question
                    If (strQuestionType = "C") Then
                        'Choice Question
                        row.Cells.Clear()

                        'generate question label
                        cell = New System.Web.UI.WebControls.TableCell
                        cell.HorizontalAlign = HorizontalAlign.Left
                        Dim lblQuestionNumber As Label = New Label
                        lblQuestionNumber.Text = visibleQuestionNumber & ".  " & dt.Rows(index).Item("QUESTION_TEXT") & "<br/>"
                        lblQuestionNumber.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0, 153)
                        lblQuestionNumber.CssClass = "boldContentTeal"
                        lblQuestionNumber.ID = "lblQuestion" & dt.Rows(index).Item("QUESTION_ID")
                        cell.Controls.Add(lblQuestionNumber)

                        If (index < dt.Rows.Count()) Then
                            'Generate the check box list control for the question
                            Dim currentQuestion As Integer = dt.Rows(index).Item("QUESTION_ID")
                            Dim chkbxList As CheckBoxList = New CheckBoxList
                            Try
                                While (index < dt.Rows.Count())
                                    If (dt.Rows(index).Item("QUESTION_ID") = currentQuestion) Then
                                        Dim li As ListItem = New ListItem
                                        li.Text = dt.Rows(index).Item("LIST_ITEM_TITLE")
                                        li.Value = dt.Rows(index).Item("LIST_ITEM_VALUE")
                                        chkbxList.Items.Add(li)
                                        index += 1
                                    Else
                                        Exit While
                                    End If
                                End While
                                index -= 1
                            Catch ex As Exception
                                Trace.Warn(ex.ToString)
                            End Try

                            chkbxList.CssClass = "borderless"
                            chkbxList.ID = "chxbxList" & dt.Rows(index).Item("QUESTION_ID")
                            chkbxList.Attributes.Add("QuestionNumber", dt.Rows(index).Item("QUESTION_ID"))
                            chkbxList.Attributes.Add("QuestionType", strQuestionType)
                            cell.Controls.Add(chkbxList)
                            cell.Wrap = True
                            cell.EnableViewState = True
                            cell.CssClass = "borderless"
                            row.Cells.Add(cell)
                        End If

                        'increment the visible question number
                        visibleQuestionNumber += 1
                    ElseIf (strQuestionType = "D") Then
                        'DateField Question
                        row.Cells.Clear()

                        'generate question label
                        cell = New System.Web.UI.WebControls.TableCell
                        cell.Width = Unit.Percentage(100)
                        cell.HorizontalAlign = HorizontalAlign.Left
                        Dim lblQuestionNumber As Label = New Label
                        lblQuestionNumber.Text = visibleQuestionNumber & ".  " & dt.Rows(index).Item("QUESTION_TEXT") & "<br/>"
                        lblQuestionNumber.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0, 153)
                        lblQuestionNumber.CssClass = "boldContentTeal"
                        lblQuestionNumber.ID = "lblQuestion" & dt.Rows(index).Item("QUESTION_ID")
                        cell.Controls.Add(lblQuestionNumber)

                        'Generate the datefield control for the question
                        Dim currentQuestion As Integer = dt.Rows(index).Item("QUESTION_ID")
                        Dim datField As Infragistics.WebUI.WebSchedule.WebDateChooser = New Infragistics.WebUI.WebSchedule.WebDateChooser
                        datField.Value = " "
                        datField.NullDateLabel = " "
                        datField.Font.Size = FontUnit.Small
                        datField.CalendarLayout.CellPadding = 2
                        datField.BorderStyle = BorderStyle.Solid
                        datField.BorderWidth = Unit.Pixel(1)
                        datField.BorderColor = Color.Black
                        datField.CalendarLayout.ChangeMonthToDateClicked = True
                        datField.CalendarLayout.CalendarStyle.BorderStyle = BorderStyle.Ridge
                        datField.CalendarLayout.CalendarStyle.BorderWidth = Unit.Pixel(2)
                        datField.CalendarLayout.DayHeaderStyle.BackColor = Color.Black
                        datField.CalendarLayout.DayHeaderStyle.BorderStyle = BorderStyle.Ridge
                        datField.CalendarLayout.DayHeaderStyle.BorderWidth = Unit.Pixel(2)
                        datField.CalendarLayout.DayHeaderStyle.ForeColor = Color.White
                        datField.CalendarLayout.DayHeaderStyle.Width = Unit.Pixel(50)
                        datField.CalendarLayout.DayNameFormat = DayNameFormat.Short
                        datField.CalendarLayout.DayStyle.BackColor = Color.Tan
                        datField.CalendarLayout.DayStyle.BorderStyle = BorderStyle.Ridge
                        datField.CalendarLayout.DayStyle.BorderWidth = Unit.Pixel(2)
                        datField.CalendarLayout.DropDownYearsNumber = 20
                        datField.CalendarLayout.FooterFormat = "Today: {0:d}"
                        datField.CalendarLayout.GridLineColor = Color.Transparent
                        datField.CalendarLayout.NextMonthText = ">>"
                        datField.CalendarLayout.NextPrevFormat = NextPrevFormat.CustomText
                        datField.CalendarLayout.NextPrevStyle.BackColor = Color.Silver
                        datField.CalendarLayout.NextPrevStyle.ForeColor = Color.Yellow
                        datField.CalendarLayout.OtherMonthDayStyle.BackColor = Color.Silver
                        datField.CalendarLayout.OtherMonthDayStyle.ForeColor = Color.FromArgb(14737632)
                        datField.CalendarLayout.PrevMonthText = "<<"
                        datField.CalendarLayout.SelectedDayStyle.BackColor = Color.FromArgb(49344)
                        datField.CalendarLayout.ShowFooter = False
                        datField.CalendarLayout.TitleFormat = TitleFormat.MonthYear
                        datField.CalendarLayout.TitleStyle.BackColor = Color.Gray
                        datField.CalendarLayout.TitleStyle.ForeColor = Color.White
                        datField.CalendarLayout.TodayDayStyle.BackColor = Color.Lavender
                        datField.CalendarLayout.VisibleDayNames = Infragistics.WebUI.WebSchedule.VisibleDays.AllDays
                        datField.CalendarLayout.WeekendDayStyle.BackColor = Color.MistyRose
                        datField.CalendarLayout.WeekendDayStyle.ForeColor = Color.FromArgb(12583104)
                        datField.DropDownAlignment = Infragistics.WebUI.WebDropDown.DropDownAlignment.Left
                        datField.CalendarJavaScriptFileName = "./MyInfragisticsScripts/ig_calendar.js"
                        datField.ImageDirectory = "./MyInfragisticsImages/"
                        datField.JavaScriptFileName = "./MyInfragisticsScripts/ig_webdropdown.js"
                        datField.JavaScriptFileNameCommon = "./MyInfragisticsScripts/ig_shared.js"
                        datField.Width = Unit.Pixel(200)
                        datField.ID = "datField" & dt.Rows(index).Item("QUESTION_ID")
                        datField.Attributes.Add("QuestionNumber", dt.Rows(index).Item("QUESTION_ID"))
                        datField.Attributes.Add("QuestionType", strQuestionType)
                        cell.Controls.Add(datField)
                        cell.Wrap = False
                        cell.EnableViewState = True
                        cell.CssClass = "borderless"
                        row.Cells.Add(cell)

                        'increment the visible question number
                        visibleQuestionNumber += 1
                    ElseIf (strQuestionType = "H") Then
                        'Horizontal Rule Question
                        row.Cells.Clear()

                        'generate horizontal rule
                        cell = New System.Web.UI.WebControls.TableCell
                        cell.HorizontalAlign = HorizontalAlign.Left
                        Dim lblQuestion As Label = New Label
                        lblQuestion.Text = "<hr style=""BORDER-LEFT-COLOR: #000099; BORDER-BOTTOM-COLOR: #000099; BORDER-TOP-STYLE: solid; " & _
                            "BORDER-TOP-COLOR: #000099; BORDER-RIGHT-STYLE: solid; BORDER-LEFT-STYLE: solid; BORDER-RIGHT-COLOR: #000099; " & _
                            "BORDER-BOTTOM-STYLE: solid"" size=3>"
                        lblQuestion.CssClass = "borderless"
                        lblQuestion.ID = "HR" & dt.Rows(index).Item("QUESTION_ID")
                        cell.Controls.Add(lblQuestion)
                        cell.Wrap = True
                        cell.EnableViewState = True
                        cell.CssClass = "borderless"
                        row.Cells.Add(cell)
                    ElseIf (strQuestionType = "L") Then
                        'List Question
                        row.Cells.Clear()

                        'generate question label
                        cell = New System.Web.UI.WebControls.TableCell
                        cell.HorizontalAlign = HorizontalAlign.Left
                        Dim lblQuestionNumber As Label = New Label
                        lblQuestionNumber.Text = visibleQuestionNumber & ".  " & dt.Rows(index).Item("QUESTION_TEXT") & "<br/>"
                        lblQuestionNumber.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0, 153)
                        lblQuestionNumber.CssClass = "boldContentTeal"
                        lblQuestionNumber.ID = "lblQuestion" & dt.Rows(index).Item("QUESTION_ID")
                        cell.Controls.Add(lblQuestionNumber)

                        'Generate the drop down list control for the question
                        Dim currentQuestion As Integer = dt.Rows(index).Item("QUESTION_ID")
                        Dim ddlList As DropDownList = New DropDownList
                        Dim li As ListItem = New ListItem
                        li.Text = "--Select--"
                        li.Value = ""
                        ddlList.Items.Add(li)
                        Try
                            If (index < dt.Rows.Count()) Then
                                While (index < dt.Rows.Count())
                                    If (dt.Rows(index).Item("QUESTION_ID") = currentQuestion) Then
                                        li = New ListItem
                                        li.Text = dt.Rows(index).Item("LIST_ITEM_TITLE")
                                        li.Value = dt.Rows(index).Item("LIST_ITEM_VALUE")
                                        ddlList.Items.Add(li)
                                        index += 1
                                    Else
                                        Exit While
                                    End If
                                End While
                                index -= 1
                            End If
                        Catch ex As Exception
                            Trace.Warn(ex.ToString)
                        End Try

                        ddlList.CssClass = "borderless"
                        ddlList.ID = "ddlList" & dt.Rows(index).Item("QUESTION_ID")
                        ddlList.Attributes.Add("QuestionNumber", dt.Rows(index).Item("QUESTION_ID"))
                        ddlList.Attributes.Add("QuestionType", strQuestionType)
                        cell.Controls.Add(ddlList)
                        cell.Wrap = True
                        cell.EnableViewState = True
                        cell.CssClass = "borderless"
                        row.Cells.Add(cell)

                        'increment the visible question number
                        visibleQuestionNumber += 1
                    ElseIf (strQuestionType = "M") Then
                        'Textfield Question
                        row.Cells.Clear()

                        'generate question label
                        cell = New System.Web.UI.WebControls.TableCell
                        cell.HorizontalAlign = HorizontalAlign.Left
                        Dim lblQuestionNumber As Label = New Label
                        lblQuestionNumber.Text = visibleQuestionNumber & ".  " & dt.Rows(index).Item("QUESTION_TEXT") & "<br/>"
                        lblQuestionNumber.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0, 153)
                        lblQuestionNumber.CssClass = "boldContentTeal"
                        lblQuestionNumber.ID = "lblQuestion" & dt.Rows(index).Item("QUESTION_ID")
                        cell.Controls.Add(lblQuestionNumber)

                        'Generate the textfield control for the question
                        Dim currentQuestion As Integer = dt.Rows(index).Item("QUESTION_ID")
                        Dim txtField As TextBox = New TextBox
                        txtField.TextMode = TextBoxMode.MultiLine
                        txtField.Columns = 75
                        txtField.Rows = 10
                        txtField.CssClass = "content"
                        txtField.ID = "txtField" & dt.Rows(index).Item("QUESTION_ID")
                        txtField.Attributes.Add("QuestionNumber", dt.Rows(index).Item("QUESTION_ID"))
                        txtField.Attributes.Add("QuestionType", strQuestionType)
                        cell.Controls.Add(txtField)
                        cell.Wrap = True
                        cell.EnableViewState = True
                        cell.CssClass = "content"
                        row.Cells.Add(cell)

                        'increment the visible question number
                        visibleQuestionNumber += 1
                    ElseIf (strQuestionType = "N") Then
                        'Number Field Question
                        row.Cells.Clear()

                        'generate question label
                        cell = New System.Web.UI.WebControls.TableCell
                        cell.HorizontalAlign = HorizontalAlign.Left
                        Dim lblQuestionNumber As Label = New Label
                        lblQuestionNumber.Text = visibleQuestionNumber & ".  " & dt.Rows(index).Item("QUESTION_TEXT") & "<br/>"
                        lblQuestionNumber.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0, 153)
                        lblQuestionNumber.CssClass = "boldContentTeal"
                        lblQuestionNumber.ID = "lblQuestion" & dt.Rows(index).Item("QUESTION_ID")
                        cell.Controls.Add(lblQuestionNumber)

                        'Generate the datefield control for the question
                        Dim currentQuestion As Integer = dt.Rows(index).Item("QUESTION_ID")
                        Dim numField As Infragistics.WebUI.WebDataInput.WebNumericEdit = New Infragistics.WebUI.WebDataInput.WebNumericEdit
                        numField.CssClass = "borderless"
                        numField.Font.Size = FontUnit.Small
                        numField.BorderStyle = BorderStyle.Solid
                        numField.BorderWidth = Unit.Pixel(1)
                        numField.BorderColor = Color.Black
                        numField.Section508Compliant = True
                        numField.DataMode = Infragistics.WebUI.WebDataInput.NumericDataMode.Decimal
                        numField.MaxValue = 2147483647
                        numField.ImageDirectory = "./MyInfragisticsImages/images/"
                        numField.JavaScriptFileName = "./MyInfragisticsScripts/ig_edit.js"
                        numField.JavaScriptFileNameCommon = "./MyInfragisticsScripts/ig_shared.js"
                        numField.Width = Unit.Pixel(200)
                        numField.ID = "numField" & dt.Rows(index).Item("QUESTION_ID")
                        numField.Attributes.Add("QuestionNumber", dt.Rows(index).Item("QUESTION_ID"))
                        numField.Attributes.Add("QuestionType", strQuestionType)
                        cell.Controls.Add(numField)
                        cell.Wrap = False
                        cell.EnableViewState = True
                        cell.CssClass = "borderless"
                        row.Cells.Add(cell)

                        'restore the previous number value if this is to be shown at the top of the report
                        If (CType(dt.Rows(index).Item("FILTER_IND"), Boolean)) Then
                            numField.Text = Session("numField" & dt.Rows(index).Item("QUESTION_ID"))
                        End If

                        'increment the visible question number
                        visibleQuestionNumber += 1
                    ElseIf (strQuestionType = "P") Then
                        'Picture Question
                        row.Cells.Clear()

                        'generate question label
                        cell = New System.Web.UI.WebControls.TableCell
                        cell.HorizontalAlign = HorizontalAlign.Left
                        Dim lblQuestionNumber As Label = New Label
                        lblQuestionNumber.Text = visibleQuestionNumber & ".  " & dt.Rows(index).Item("QUESTION_TEXT") & "<br/>"
                        lblQuestionNumber.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0, 153)
                        lblQuestionNumber.CssClass = "boldContentTeal"
                        lblQuestionNumber.ID = "lblQuestion" & dt.Rows(index).Item("QUESTION_ID")
                        cell.Controls.Add(lblQuestionNumber)

                        'Generate the radio-button list control for the question
                        Dim currentQuestion As Integer = dt.Rows(index).Item("QUESTION_ID")
                        Dim radbtnList As RadioButtonList = New RadioButtonList
                        If (index < dt.Rows.Count()) Then
                            While (index < dt.Rows.Count())
                                If (dt.Rows(index).Item("QUESTION_ID") = currentQuestion) Then
                                    Dim li As ListItem = New ListItem
                                    If (Trim("" & dt.Rows(index).Item("LIST_ITEM_IMAGE_TYPE")) <> "") Then
                                        'resize the window to the image size + 20
                                        Dim myImage As System.Drawing.Image = Nothing
                                        Dim myStream As MemoryStream = New MemoryStream
                                        Dim imageWidth As Integer = 300
                                        Dim imageHeight As Integer = 300
                                        Dim processedMemStream As MemoryStream = New MemoryStream

                                        myStream.Write(CType(dt.Rows(index).Item("LIST_ITEM_IMAGE"), Byte()), 0, CType(dt.Rows(index).Item("LIST_ITEM_IMAGE"), Byte()).Length)
                                        myImage = System.Drawing.Image.FromStream(myStream)
                                        imageWidth = myImage.Width + 32
                                        imageHeight = myImage.Height + 32
                                        li.Text = dt.Rows(index).Item("LIST_ITEM_TITLE") & " <a class='pointer' onclick=""WinPopNew('./Picture.aspx?seqTemplate=" & Session("seqTemplate") & _
                                        "&intQuestion=" & dt.Rows(index).Item("QUESTION_ID") & "&intList=" & _
                                        dt.Rows(index).Item("LIST_ITEM_ID") & "', " & imageWidth & ", " & imageHeight & ")"">" & _
                                        "<img border=0 src=""./Images/QuestionLookup.png""></a><br/>"
                                    Else
                                        li.Text = dt.Rows(index).Item("LIST_ITEM_TITLE")
                                    End If

                                    li.Value = dt.Rows(index).Item("LIST_ITEM_VALUE")
                                    radbtnList.Items.Add(li)
                                    index += 1
                                Else
                                    Exit While
                                End If
                            End While
                            index -= 1
                        End If

                        radbtnList.CssClass = "borderless"
                        radbtnList.ID = "radbtnList" & dt.Rows(index).Item("QUESTION_ID")
                        radbtnList.Attributes.Add("QuestionNumber", dt.Rows(index).Item("QUESTION_ID"))
                        radbtnList.Attributes.Add("QuestionType", strQuestionType)
                        cell.Controls.Add(radbtnList)
                        cell.Wrap = True
                        cell.EnableViewState = True
                        cell.CssClass = "borderless"
                        row.Cells.Add(cell)

                        'increment the visible question number
                        visibleQuestionNumber += 1
                    ElseIf (strQuestionType = "Q") Then
                        'Hidden Field Question
                        'row.Cells.Clear()

                        'generate hidden field
                        'cell = New System.Web.UI.WebControls.TableCell
                        'cell.HorizontalAlign = HorizontalAlign.Left
                        'Dim lblQuestionHidden As Label = New Label
                        'lblQuestionHidden.Text = "Note: This field is not normally visible but contains the following: " & _
                        '        dt.Rows(index).Item("strValue") & "<br/>"
                        'lblQuestionHidden.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0, 153)
                        'lblQuestionHidden.CssClass = "content"
                        'lblQuestionHidden.ID = "lblQuestion" & dt.Rows(index).Item("intQuestion")
                        'cell.Controls.Add(lblQuestionHidden)
                        'cell.Wrap = True
                        'cell.EnableViewState = True
                        'cell.BorderStyle = BorderStyle.None
                        'cell.ColumnSpan = 2
                        'row.Cells.Add(cell)
                    ElseIf (strQuestionType = "R") Then
                        'Rating Question
                        row.Cells.Clear()

                        'generate question label
                        cell = New System.Web.UI.WebControls.TableCell
                        cell.HorizontalAlign = HorizontalAlign.Left
                        Dim lblQuestionNumber As Label = New Label
                        lblQuestionNumber.Text = visibleQuestionNumber & ".  " & dt.Rows(index).Item("QUESTION_TEXT") & "<br/>"
                        lblQuestionNumber.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0, 153)
                        lblQuestionNumber.CssClass = "boldContentTeal"
                        lblQuestionNumber.ID = "lblQuestion" & dt.Rows(index).Item("QUESTION_ID")
                        cell.Controls.Add(lblQuestionNumber)

                        'Generate the check box list control for the question
                        Dim currentQuestion As Integer = dt.Rows(index).Item("QUESTION_ID")
                        Dim radbtnList As RadioButtonList = New RadioButtonList
                        Dim initial As Integer = dt.Rows(index).Item("RATING_INITIAL_VALUE")
                        Dim interval As Integer = dt.Rows(index).Item("RATING_STEP_VALUE")
                        Dim count As Integer = dt.Rows(index).Item("RATING_COUNT")
                        If (CType(dt.Rows(index).Item("NOT_APPLICABLE_IND"), Boolean)) Then
                            Dim li As ListItem
                            While (count > 1)
                                li = New ListItem
                                li.Text = initial
                                li.Value = initial
                                radbtnList.Items.Add(li)
                                initial += interval
                                count -= 1
                            End While
                            li = New ListItem
                            li.Text = "N/A"
                            li.Value = ""
                            radbtnList.Items.Add(li)
                        Else
                            While (count > 0)
                                Dim li As ListItem = New ListItem
                                li.Text = initial
                                li.Value = initial
                                radbtnList.Items.Add(li)
                                initial += interval
                                count -= 1
                            End While
                        End If
                        radbtnList.CssClass = "borderless"
                        radbtnList.ID = "radbtnList" & dt.Rows(index).Item("QUESTION_ID")
                        radbtnList.Attributes.Add("QuestionNumber", dt.Rows(index).Item("QUESTION_ID"))
                        radbtnList.Attributes.Add("QuestionType", strQuestionType)
                        radbtnList.RepeatDirection = RepeatDirection.Horizontal
                        radbtnList.TextAlign = TextAlign.Right
                        cell.Controls.Add(radbtnList)
                        cell.Wrap = True
                        cell.EnableViewState = True
                        cell.CssClass = "borderless"
                        row.Cells.Add(cell)

                        'increment the visible question number
                        visibleQuestionNumber += 1
                    ElseIf (strQuestionType = "S") Then
                        '1-Line Field Question
                        row.Cells.Clear()

                        'generate question label
                        cell = New System.Web.UI.WebControls.TableCell
                        cell.HorizontalAlign = HorizontalAlign.Left
                        Dim lblQuestionNumber As Label = New Label
                        lblQuestionNumber.Text = visibleQuestionNumber & ".  " & dt.Rows(index).Item("QUESTION_TEXT") & "<br/>"
                        lblQuestionNumber.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0, 153)
                        lblQuestionNumber.CssClass = "boldContentTeal"
                        lblQuestionNumber.ID = "lblQuestion" & dt.Rows(index).Item("QUESTION_ID")
                        cell.Controls.Add(lblQuestionNumber)

                        'Generate the textfield control for the question
                        Dim currentQuestion As Integer = dt.Rows(index).Item("QUESTION_ID")
                        Dim txtField As TextBox = New TextBox
                        txtField.TextMode = TextBoxMode.SingleLine
                        txtField.Width = Unit.Pixel(300)
                        txtField.CssClass = "borderless"
                        txtField.ID = "txtField" & dt.Rows(index).Item("QUESTION_ID")
                        txtField.Attributes.Add("QuestionNumber", dt.Rows(index).Item("QUESTION_ID"))
                        txtField.Attributes.Add("QuestionType", strQuestionType)
                        txtField.Attributes.Add("ListReportTop", IIf(CType(dt.Rows(index).Item("FILTER_IND"), Boolean), 1, 0))

                        'restore the previous text value if this is to be shown at the top of the report
                        If (CType(dt.Rows(index).Item("FILTER_IND"), Boolean)) Then
                            txtField.Text = Session("txtField" & dt.Rows(index).Item("QUESTION_ID"))
                        End If

                        cell.Controls.Add(txtField)
                        cell.Wrap = True
                        cell.EnableViewState = True
                        cell.CssClass = "borderless"
                        row.Cells.Add(cell)

                        'increment the visible question number
                        visibleQuestionNumber += 1
                    ElseIf (strQuestionType = "T") Then
                        'Type Question
                        row.Cells.Clear()

                        'generate question label
                        cell = New System.Web.UI.WebControls.TableCell
                        cell.HorizontalAlign = HorizontalAlign.Left
                        Dim lblQuestionNumber As Label = New Label
                        lblQuestionNumber.Text = visibleQuestionNumber & ".  " & dt.Rows(index).Item("QUESTION_TEXT") & "<br/>"
                        lblQuestionNumber.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0, 153)
                        lblQuestionNumber.CssClass = "boldContentTeal"
                        lblQuestionNumber.ID = "lblQuestion" & dt.Rows(index).Item("QUESTION_ID")
                        cell.Controls.Add(lblQuestionNumber)

                        'Generate the radio-button list control for the question
                        Dim currentQuestion As Integer = dt.Rows(index).Item("QUESTION_ID")
                        Dim radbtnList As RadioButtonList = New RadioButtonList
                        If (index < dt.Rows.Count()) Then
                            While (index < dt.Rows.Count())
                                If (dt.Rows(index).Item("QUESTION_ID") = currentQuestion) Then
                                    Dim li As ListItem = New ListItem
                                    li.Text = dt.Rows(index).Item("LIST_ITEM_TITLE")
                                    li.Value = dt.Rows(index).Item("LIST_ITEM_VALUE")
                                    radbtnList.Items.Add(li)
                                    index += 1
                                Else
                                    Exit While
                                End If
                            End While
                            index -= 1
                        End If

                        radbtnList.CssClass = "borderless"
                        radbtnList.ID = "radbtnList" & dt.Rows(index).Item("QUESTION_ID")
                        radbtnList.Attributes.Add("QuestionNumber", dt.Rows(index).Item("QUESTION_ID"))
                        radbtnList.Attributes.Add("QuestionType", strQuestionType)
                        cell.Controls.Add(radbtnList)
                        cell.Wrap = True
                        cell.EnableViewState = True
                        cell.CssClass = "borderless"
                        row.Cells.Add(cell)

                        'increment the visible question number
                        visibleQuestionNumber += 1
                    ElseIf (strQuestionType = "X") Then
                        'Comment Question
                        row.Cells.Clear()

                        'generate question label
                        cell = New System.Web.UI.WebControls.TableCell
                        cell.HorizontalAlign = HorizontalAlign.Left
                        Dim lblQuestionComment As Label = New Label
                        lblQuestionComment.Text = "     " & dt.Rows(index).Item("QUESTION_TEXT") & "<br/>"
                        lblQuestionComment.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0, 153)
                        lblQuestionComment.CssClass = "boldContentTeal"
                        lblQuestionComment.ID = "lblQuestion" & dt.Rows(index).Item("QUESTION_ID")
                        cell.Controls.Add(lblQuestionComment)
                        cell.Wrap = True
                        cell.EnableViewState = True
                        cell.CssClass = "borderless"
                        row.Cells.Add(cell)
                    ElseIf (strQuestionType = "Z") Then
                        'Prepopulated Question
                        row.Cells.Clear()

                        'generate question label
                        cell = New System.Web.UI.WebControls.TableCell
                        cell.HorizontalAlign = HorizontalAlign.Left
                        Dim lblQuestionNumber As Label = New Label
                        lblQuestionNumber.Text = visibleQuestionNumber & ".  " & dt.Rows(index).Item("QUESTION_TEXT") & "<br/>"
                        lblQuestionNumber.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0, 153)
                        lblQuestionNumber.CssClass = "boldContentTeal"
                        lblQuestionNumber.ID = "lblQuestion" & dt.Rows(index).Item("QUESTION_ID")
                        cell.Controls.Add(lblQuestionNumber)

                        'get the query for this question
                        Dim SqlSelectCommandQuery As New SqlClient.SqlCommand
                        Dim SqlQueryDataAdapter As New SqlClient.SqlDataAdapter
                        SqlSelectCommandQuery.Connection = Me.commonConn
                        SqlSelectCommandQuery.CommandText = "Select * from " & myUtility.getExtension() & _
                                                            "SAT_TEMPLATE_QUESTION_QUERIES where QUERY_NAME = '" & _
                                                            dt.Rows(index).Item("QUERY_NAME") & "'"
                        Trace.Warn(SqlSelectCommandQuery.CommandText)
                        SqlQueryDataAdapter.SelectCommand = SqlSelectCommandQuery
                        If (Me.dsCommon.Tables.Contains("PrePopQuery")) Then
                            Me.dsCommon.Tables("PrePopQuery").Clear()
                        End If
                        SqlQueryDataAdapter.Fill(Me.dsCommon, "PrePopQuery")

                        'get the data from the query for this question
                        Dim myConnection As New OracleClient.OracleConnection
                        Dim OracleSelectCommandQuery As New OracleClient.OracleCommand
                        Dim OracleQueryDataAdapter As New OracleClient.OracleDataAdapter
                        myConnection = myUtility.getConnection(myConnection, Me.dsCommon.Tables("PrePopQuery").Rows(0).Item("QUERY_DATABASE_NAME"))
                        OracleSelectCommandQuery.CommandText = Me.dsCommon.Tables("PrePopQuery").Rows(0).Item("QUERY_STRING")
                        OracleSelectCommandQuery.Connection = myConnection
                        OracleQueryDataAdapter.SelectCommand = OracleSelectCommandQuery
                        Trace.Warn(OracleSelectCommandQuery.CommandText)
                        If (Me.dsCommon.Tables.Contains("PrePopQueryData")) Then
                            Me.dsCommon.Tables("PrePopQueryData").Clear()
                        End If
                        OracleQueryDataAdapter.Fill(Me.dsCommon, "PrePopQueryData")
                        Dim dt2 As DataTable = Me.dsCommon.Tables("PrePopQueryData")

                        'Generate the drop down list control for the question
                        Dim ddlList As DropDownList = New DropDownList
                        Dim li As ListItem = New ListItem
                        Dim PrePopIndex = 0
                        li.Text = "--Select--"
                        li.Value = ""
                        ddlList.Items.Add(li)
                        Try
                            If (PrePopIndex < dt2.Rows.Count()) Then
                                While (PrePopIndex < dt2.Rows.Count())
                                    li = New ListItem
                                    li.Text = dt2.Rows(PrePopIndex).Item(Me.dsCommon.Tables("PrePopQuery").Rows(0).Item("QUERY_RESULT_LABEL"))
                                    li.Value = dt2.Rows(PrePopIndex).Item(Me.dsCommon.Tables("PrePopQuery").Rows(0).Item("QUERY_RESULT_VALUE"))
                                    ddlList.Items.Add(li)
                                    PrePopIndex += 1
                                End While
                            End If
                        Catch ex As Exception
                            Trace.Warn(ex.ToString)
                        End Try

                        ddlList.CssClass = "borderless"
                        ddlList.ID = "ddlList" & dt.Rows(index).Item("QUESTION_ID")
                        ddlList.Attributes.Add("QuestionNumber", dt.Rows(index).Item("QUESTION_ID"))
                        ddlList.Attributes.Add("QuestionType", strQuestionType)
                        ddlList.Attributes.Add("ListReportTop", IIf(CType(dt.Rows(index).Item("FILTER_IND"), Boolean), 1, 0))

                        'restore the previous selection
                        If (CType(dt.Rows(index).Item("FILTER_IND"), Boolean)) Then
                            Dim tempIndex As Integer = ddlList.Items.IndexOf(ddlList.Items.FindByValue(Session("ddlList" & dt.Rows(index).Item("QUESTION_ID"))))
                            If (tempIndex < 0) Then
                                ddlList.SelectedIndex = 0
                            Else
                                ddlList.SelectedIndex = tempIndex
                            End If
                        End If

                        cell.Controls.Add(ddlList)
                        cell.Wrap = True
                        cell.EnableViewState = True
                        cell.CssClass = "borderless"
                        row.Cells.Add(cell)

                        'increment the visible question number
                        visibleQuestionNumber += 1
                    End If
                    'add the row to the table and increment to move to the next row of data
                    row.EnableViewState = True
                    row.CssClass = "borderless"
                    QuestionTable.Rows.Add(row)
                    QuestionTable.EnableViewState = True
                    index += 1
                End While
                Session("QuestionAvailability") = "Some"
            Else
                Session("QuestionAvailability") = "None"
            End If
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Session("QuestionAvailability") = "None"
            alert("Failed to load the questions for this survey.")
            Return False
        End Try
    End Function

    'called by other functions to reload the page data
    Private Sub reload(Optional ByVal skipFullLoad As Boolean = False)
        'Refill the template and survey drop-downs
        If Not (skipFullLoad) Then
            loadData()
        End If
        Me.ddlTemplateList.SelectedItem.Selected = False
        Me.ddlTemplateList.Items.FindByValue(Session("seqTemplate")).Selected = True
        If (tagFill()) Then
            If (Session("InputSwitch") <> "Non-Respondent") Then
                questionFill()
            End If
        End If
    End Sub

    'returns the next user id
    Private Function getNextUserID() As Integer
        Try
            'get the next user ID
            Me.sqlCommonAction.CommandText = "Select isnull(max(RESPONDENT_USER_KEY),0) as UserID from " & myUtility.getExtension() & "SAT_RESPONSE where SURVEY_KEY = " & _
                Session("seqSurvey")
            Me.sqlCommonAdapter.SelectCommand = Me.sqlCommonAction
            Me.sqlCommonAdapter.SelectCommand.Connection = Me.commonConn
            Me.sqlCommonAdapter.Fill(Me.dsCommon, "NextUserID")
            Return Me.dsCommon.Tables("NextUserID").Rows(0).Item("UserID") + 1
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return -1
        End Try
    End Function
#End Region

#Region "Checks"
    'Switch between respondent and non-respondent input
    Private Sub chkInputSwitch_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkInputSwitch.CheckedChanged
        If (Session("InputSwitch") = "Non-Respondent") Then
            Session("InputSwitch") = "Respondent"
            Me.chkInputSwitch.Text = "Check for non-respondent input."
            Me.chkInputSwitch.Checked = False
        Else
            holdResponseData()
            Session("InputSwitch") = "Non-Respondent"
            Me.chkInputSwitch.Text = "Check for respondent input."
            Me.chkInputSwitch.Checked = False
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

        'destroy the tree to avoid control loading failure on the viewstate
        Page.Controls(0).Controls(1).Controls(0).Controls(14).Controls.Clear()
    End Sub

    'Handles the user selecting a survey
    Private Sub ddlSurveyList_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlSurveyList.SelectedIndexChanged
        Session("JSAction") = ""

        If (Me.ddlSurveyList.SelectedIndex <> 0) Then
            Session("seqSurvey") = Me.ddlSurveyList.SelectedItem.Value
            Session("seqTemplate") = CType(Session("surveyData"), DataTable).Rows(Me.ddlSurveyList.SelectedIndex - 1).Item("TEMPLATE_KEY")
            If (surveyFill()) Then
                reload(True)
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

        If (surveyFill()) Then
            'reset the controls
            Me.wneNonRespondent.Text = ""
            reload(True)
        Else
            Session("SurveysExist") = "No"
        End If
    End Sub

    'submits the data on the page
    Private Sub btnSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubmit.Click, btnSubmit2.Click
        Session("JSAction") = ""

        If (Session("InputSwitch") = "Respondent") Then
            'store the respondent data
            If Not (storeRespondentData()) Then
                alert("Unable to save the respondent information.")
            Else
                alert("Respondent information saved.")
                reload(True)
            End If
        Else
            'store the non-respondent data
            If Not (storeNonRespondentData()) Then
                alert("Unable to save the data for the non-respondents.")
            Else
                alert("Non-respondents saved.")

                'reset the controls
                Me.wneNonRespondent.Text = ""
            End If
        End If
    End Sub
#End Region

#Region "Store Data"
    'stores the respondent data
    Private Function storeRespondentData() As Boolean
        Try
            'start transaction
            If (myUtility.setupTransaction(sqlCommonAction, Me.commonConn)) Then

                'get the next user ID
                Dim nextUserID As Integer = getNextUserID()

                'if it failed clean up and return false
                If (nextUserID = -1) Then
                    'clear the transaction if it exists
                    If (Session("transExists") = True) Then
                        Me.sqlCommonAction.Transaction.Rollback()
                    End If
                    Return False
                End If

                'set complete date to today if it is set to nothing
                If (Trim("" & Me.wdcCompleteDate.Value) = "") Then
                    Me.wdcCompleteDate.Value = Now
                End If

                'add respondent data
                Me.sqlCommonAction.CommandText = "Insert into " & myUtility.getExtension() & "SAT_RESPONSE (RESPONDENT_USER_KEY, SURVEY_KEY, CREATE_DATE, LAST_USED_DATE, " & _
                "REMIND_DATE, SUBMISSION_COUNT, LAST_COMPLETION_DATE) VALUES (" & nextUserID & ", " & Session("seqSurvey") & _
                ", '" & Me.wdcCompleteDate.Value & "', '" & Me.wdcCompleteDate.Value & "', '1/1/1970', 0, '" & Me.wdcCompleteDate.Value & "')"
                Me.sqlCommonAction.ExecuteNonQuery()

                'get the questions section
                Dim myControl As Control
                myControl = Page.Controls(0).Controls(1).Controls(0).Controls(14)

                'store the results based on question type
                Dim index As Integer = 0
                Dim radbtnList As RadioButtonList
                Dim txtField As TextBox
                Dim chxbxList As CheckBoxList
                Dim datField As Infragistics.WebUI.WebSchedule.WebDateChooser
                Dim ddlList As DropDownList
                Dim numField As Infragistics.WebUI.WebDataInput.WebNumericEdit
                Dim blnDataToStore As Boolean = False
                While (index < myControl.Controls.Count)
                    'skip hidden variables
                    If (myControl.Controls(index).Controls.Count() > 0) Then
                        'skip all formatting controls on the page
                        If (myControl.Controls(index).Controls(0).Controls.Count() = 2) Then
                            'determine control type by ID of the control
                            If (myControl.Controls(index).Controls(0).Controls(1).ID.StartsWith("radbtnList")) Then
                                'Get the radio button control information
                                radbtnList = CType(myControl.Controls(index).Controls(0).Controls(1), RadioButtonList)

                                If (Trim("" & radbtnList.SelectedValue) <> "") Then
                                    'build the insert query
                                    Me.sqlCommonAction.CommandText = "Insert into " & myUtility.getExtension() & "SAT_RESPONSE_RESULT (SURVEY_KEY, RESPONDENT_USER_KEY, RESPONSE_DATE, " & _
                                    "QUESTION_ID, QUESTION_TYPE, ANSWER_VALUE, SUBMISSION_COUNT, ANSWER_TEXT) VALUES (" & _
                                    Session("seqSurvey") & ", " & nextUserID & ", '" & Me.wdcCompleteDate.Value & "', " & _
                                    radbtnList.Attributes.Item("QuestionNumber") & ", '" & _
                                    radbtnList.Attributes.Item("QuestionType") & "', " & radbtnList.SelectedValue & ", 0, '')"

                                    'write the saved choices for the user to the database
                                    Me.sqlCommonAction.ExecuteNonQuery()
                                    blnDataToStore = True
                                End If
                            ElseIf (myControl.Controls(index).Controls(0).Controls(1).ID.StartsWith("txtField")) Then
                                'Get the text field control information
                                txtField = CType(myControl.Controls(index).Controls(0).Controls(1), TextBox)

                                If (Trim("" & txtField.Text) <> "") Then
                                    'parameterized the text input to allow for things like quotes to get through
                                    Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@ANSWER_TEXT", System.Data.SqlDbType.VarChar, 1024, "ANSWER_TEXT")
                                    param0.Value = txtField.Text
                                    Me.sqlCommonAction.Parameters.Add(param0)

                                    'build the insert query
                                    Me.sqlCommonAction.CommandText = "Insert into " & myUtility.getExtension() & "SAT_RESPONSE_RESULT (SURVEY_KEY, RESPONDENT_USER_KEY, RESPONSE_DATE, " & _
                                    "QUESTION_ID, QUESTION_TYPE, ANSWER_VALUE, SUBMISSION_COUNT, ANSWER_TEXT) VALUES (" & _
                                    Session("seqSurvey") & ", " & nextUserID & ", '" & Me.wdcCompleteDate.Value & "', " & _
                                    txtField.Attributes.Item("QuestionNumber") & ", '" & _
                                    txtField.Attributes.Item("QuestionType") & "', 0, 0, @ANSWER_TEXT)"

                                    'write the saved choices for the user to the database
                                    Me.sqlCommonAction.ExecuteNonQuery()
                                    Me.sqlCommonAction.Parameters.Clear()
                                    blnDataToStore = True

                                    'if we have a 1-line text field that is appearing at the top of
                                    'the report page then we need to maintain its value
                                    If (txtField.Attributes.Item("QuestionType") = "S") Then
                                        If (CType(txtField.Attributes.Item("ListReportTop"), Boolean)) Then
                                            Session("txtField" & txtField.Attributes.Item("QuestionNumber")) = txtField.Text
                                        End If
                                    End If
                                Else
                                    'if we have a 1-line text field that is appearing at the top of
                                    'the report page then we need to maintain its value
                                    If (txtField.Attributes.Item("QuestionType") = "S") Then
                                        If (CType(txtField.Attributes.Item("ListReportTop"), Boolean)) Then
                                            Session("txtField" & txtField.Attributes.Item("QuestionNumber")) = ""
                                        End If
                                    End If
                                End If
                            ElseIf (myControl.Controls(index).Controls(0).Controls(1).ID.StartsWith("chxbxList")) Then
                                'Get the check box list control information
                                chxbxList = CType(myControl.Controls(index).Controls(0).Controls(1), CheckBoxList)

                                'build the insert query
                                Me.sqlCommonAction.CommandText = "Insert into " & myUtility.getExtension() & "SAT_RESPONSE_RESULT (SURVEY_KEY, RESPONDENT_USER_KEY, RESPONSE_DATE, " & _
                                "QUESTION_ID, QUESTION_TYPE, ANSWER_VALUE, SUBMISSION_COUNT, ANSWER_TEXT) VALUES (" & _
                                Session("seqSurvey") & ", " & nextUserID & ", '" & Me.wdcCompleteDate.Value & "', " & _
                                chxbxList.Attributes.Item("QuestionNumber") & ", '" & _
                                chxbxList.Attributes.Item("QuestionType") & "', "

                                'get the total value and create a string of values for each selected value
                                Dim tempIndex As Integer = 0
                                Dim total As Integer = 0
                                Dim selectedString As String = ""
                                While (tempIndex < chxbxList.Items.Count())
                                    If (chxbxList.Items(tempIndex).Selected = True) Then
                                        total += chxbxList.Items(tempIndex).Value
                                        selectedString &= chxbxList.Items(tempIndex).Value & ", "
                                    End If
                                    tempIndex += 1
                                End While

                                If (Trim("" & selectedString) <> "") Then
                                    'finish the insert query
                                    selectedString = selectedString.Substring(0, selectedString.Length - 2)
                                    Me.sqlCommonAction.CommandText &= total & ", 0, '" & selectedString & "')"

                                    'write the saved choices for the user to the database
                                    Me.sqlCommonAction.ExecuteNonQuery()
                                    blnDataToStore = True
                                End If
                            ElseIf (myControl.Controls(index).Controls(0).Controls(1).ID.StartsWith("datField")) Then
                                'Get the infragistics date field control information
                                datField = CType(myControl.Controls(index).Controls(0).Controls(1), Infragistics.WebUI.WebSchedule.WebDateChooser)

                                If (Trim("" & datField.Value) <> "") Then
                                    'build the insert query
                                    Me.sqlCommonAction.CommandText = "Insert into " & myUtility.getExtension() & "SAT_RESPONSE_RESULT (SURVEY_KEY, RESPONDENT_USER_KEY, RESPONSE_DATE, " & _
                                    "QUESTION_ID, QUESTION_TYPE, ANSWER_VALUE, SUBMISSION_COUNT, ANSWER_TEXT) VALUES (" & _
                                    Session("seqSurvey") & ", " & nextUserID & ", '" & Me.wdcCompleteDate.Value & "', " & _
                                    datField.Attributes.Item("QuestionNumber") & ", '" & _
                                    datField.Attributes.Item("QuestionType") & "', 0, 0, '" & datField.Value & "')"

                                    'write the saved choices for the user to the database
                                    Me.sqlCommonAction.ExecuteNonQuery()
                                    blnDataToStore = True
                                End If
                            ElseIf (myControl.Controls(index).Controls(0).Controls(1).ID.StartsWith("ddlList")) Then
                                'Get the drop down list control information
                                ddlList = CType(myControl.Controls(index).Controls(0).Controls(1), DropDownList)

                                If (Trim("" & ddlList.SelectedValue) <> "") Then
                                    'build the insert query
                                    Me.sqlCommonAction.CommandText = "Insert into " & myUtility.getExtension() & "SAT_RESPONSE_RESULT (SURVEY_KEY, RESPONDENT_USER_KEY, RESPONSE_DATE, " & _
                                    "QUESTION_ID, QUESTION_TYPE, ANSWER_VALUE, SUBMISSION_COUNT, ANSWER_TEXT) VALUES (" & _
                                    Session("seqSurvey") & ", " & nextUserID & ", '" & Me.wdcCompleteDate.Value & "', " & _
                                    ddlList.Attributes.Item("QuestionNumber") & ", '"

                                    If (ddlList.Attributes.Item("QuestionType") = "S") Then
                                        Me.sqlCommonAction.CommandText &= ddlList.Attributes.Item("QuestionType") & "', " & ddlList.SelectedValue & ", 0, '')"
                                    ElseIf (ddlList.Attributes.Item("QuestionType") = "Z") Then
                                        Me.sqlCommonAction.CommandText &= ddlList.Attributes.Item("QuestionType") & "', 0, 0, '" & ddlList.SelectedValue & "^" & ddlList.SelectedItem.Text & "')"
                                    Else
                                        Me.sqlCommonAction.CommandText &= ddlList.Attributes.Item("QuestionType") & "', " & ddlList.SelectedValue & ", 0, '')"
                                    End If
                                    Trace.Warn(Me.sqlCommonAction.CommandText)
                                    'write the saved choices for the user to the database
                                    Me.sqlCommonAction.ExecuteNonQuery()
                                    blnDataToStore = True
                                End If

                                'if we have a prepopulated list field that is appearing at the top of
                                'the report page then we need to maintain its value
                                If (ddlList.Attributes.Item("QuestionType") = "Z" And Trim("" & ddlList.SelectedValue) <> "") Then
                                    If (CType(ddlList.Attributes.Item("ListReportTop"), Boolean)) Then
                                        Session("ddlList" & ddlList.Attributes.Item("QuestionNumber")) = ddlList.SelectedValue
                                        Session("ddlList" & ddlList.Attributes.Item("QuestionNumber") & "_1") = ddlList.SelectedValue & "^" & ddlList.SelectedItem.Text
                                    End If
                                End If
                            ElseIf (myControl.Controls(index).Controls(0).Controls(1).ID.StartsWith("numField")) Then
                                'Get the infragistics number field control information
                                numField = CType(myControl.Controls(index).Controls(0).Controls(1), Infragistics.WebUI.WebDataInput.WebNumericEdit)

                                If (Trim("" & numField.Text) <> "") Then
                                    'build the insert query
                                    Me.sqlCommonAction.CommandText = "Insert into " & myUtility.getExtension() & "SAT_RESPONSE_RESULT (SURVEY_KEY, RESPONDENT_USER_KEY, RESPONSE_DATE, " & _
                                    "QUESTION_ID, QUESTION_TYPE, ANSWER_VALUE, SUBMISSION_COUNT, ANSWER_TEXT) VALUES (" & _
                                    Session("seqSurvey") & ", " & nextUserID & ", '" & Me.wdcCompleteDate.Value & "', " & _
                                    numField.Attributes.Item("QuestionNumber") & ", '" & _
                                    numField.Attributes.Item("QuestionType") & "', " & numField.Value & ", 0, '')"

                                    'write the saved choices for the user to the database
                                    Me.sqlCommonAction.ExecuteNonQuery()
                                    blnDataToStore = True

                                    Session("numField" & numField.Attributes.Item("QuestionNumber")) = numField.Value
                                End If
                            End If
                        End If
                    End If

                    index += 1
                End While

                If (blnDataToStore) Then
                    'commit transaction and return
                    Me.sqlCommonAction.Transaction.Commit()
                    Session("transExists") = False
                    Return True
                Else
                    'clear the transaction if it exists
                    If (Session("transExists") = True) Then
                        Me.sqlCommonAction.Transaction.Rollback()
                    End If
                    Return False
                End If
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            'clear the transaction if it exists
            If (Session("transExists") = True) Then
                Me.sqlCommonAction.Transaction.Rollback()
            End If
            Return False
        End Try
    End Function

    'holds the respondent data
    Private Function holdResponseData() As Boolean
        Try
            'set complete date to today if it is set to nothing
            If (Trim("" & Me.wdcCompleteDate.Value) = "") Then
                Me.wdcCompleteDate.Value = Now
            End If

            'get the questions section
            Dim myControl As Control
            myControl = Page.Controls(0).Controls(1).Controls(0).Controls(14)

            'store the results based on question type
            Dim index As Integer = 0
            Dim radbtnList As RadioButtonList
            Dim txtField As TextBox
            Dim chxbxList As CheckBoxList
            Dim datField As Infragistics.WebUI.WebSchedule.WebDateChooser
            Dim ddlList As DropDownList
            Dim numField As Infragistics.WebUI.WebDataInput.WebNumericEdit
            Dim blnDataToStore As Boolean = False
            While (index < myControl.Controls.Count)
                'skip hidden variables
                If (myControl.Controls(index).Controls.Count() > 0) Then
                    'skip all formatting controls on the page
                    If (myControl.Controls(index).Controls(0).Controls.Count() = 2) Then
                        'determine control type by ID of the control
                        If (myControl.Controls(index).Controls(0).Controls(1).ID.StartsWith("txtField")) Then
                            'Get the text field control information
                            txtField = CType(myControl.Controls(index).Controls(0).Controls(1), TextBox)

                            If (Trim("" & txtField.Text) <> "") Then
                                'if we have a 1-line text field that is appearing at the top of
                                'the report page then we need to maintain its value
                                If (txtField.Attributes.Item("QuestionType") = "S") Then
                                    If (CType(txtField.Attributes.Item("ListReportTop"), Boolean)) Then
                                        Session("txtField" & txtField.Attributes.Item("QuestionNumber")) = txtField.Text
                                    End If
                                End If
                            Else
                                'if we have a 1-line text field that is appearing at the top of
                                'the report page then we need to maintain its value
                                If (txtField.Attributes.Item("QuestionType") = "S") Then
                                    If (CType(txtField.Attributes.Item("ListReportTop"), Boolean)) Then
                                        Session("txtField" & txtField.Attributes.Item("QuestionNumber")) = ""
                                    End If
                                End If
                            End If
                        ElseIf (myControl.Controls(index).Controls(0).Controls(1).ID.StartsWith("ddlList")) Then
                            'Get the drop down list control information
                            ddlList = CType(myControl.Controls(index).Controls(0).Controls(1), DropDownList)

                            'if we have a prepopulated list field that is appearing at the top of
                            'the report page then we need to maintain its value
                            If (ddlList.Attributes.Item("QuestionType") = "Z" And Trim("" & ddlList.SelectedValue) <> "") Then
                                If (CType(ddlList.Attributes.Item("ListReportTop"), Boolean)) Then
                                    Session("ddlList" & ddlList.Attributes.Item("QuestionNumber")) = ddlList.SelectedValue
                                    Session("ddlList" & ddlList.Attributes.Item("QuestionNumber") & "_1") = ddlList.SelectedValue & "^" & ddlList.SelectedItem.Text
                                End If
                            End If
                        ElseIf (myControl.Controls(index).Controls(0).Controls(1).ID.StartsWith("numField")) Then
                            'Get the infragistics number field control information
                            numField = CType(myControl.Controls(index).Controls(0).Controls(1), Infragistics.WebUI.WebDataInput.WebNumericEdit)

                            If (Trim("" & numField.Text) <> "") Then
                                Session("numField" & numField.Attributes.Item("QuestionNumber")) = numField.Value
                            End If
                        End If
                    End If
                End If

                index += 1
            End While

            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'stores the non-respondent data
    Private Function storeNonRespondentData() As Boolean
        Try
            'get the next user ID
            Dim nextUserID As Integer = getNextUserID()

            'if it failed clean up and return false
            If (nextUserID = -1) Then
                'clear the transaction if it exists
                If (Session("transExists") = True) Then
                    Me.sqlCommonAction.Transaction.Rollback()
                End If
                Return False
            End If

            'open the connection for the common action if it is closed
            Dim index As Integer = 0
            If (Me.sqlCommonAction.Connection.State = ConnectionState.Closed) Then
                Me.sqlCommonAction.Connection.Open()
            End If

            'set complete date to today if it is set to nothing
            If (Trim("" & Me.wdcCompleteDate.Value) = "") Then
                Me.wdcCompleteDate.Value = Now
            End If

            'get the questions to store results if there are any
            Me.sqlCommonAction.CommandText = "Select QUESTION_ID, QUESTION_TYPE_CODE from " & _
                myUtility.getExtension() & "SAT_TEMPLATE_QUESTION where TEMPLATE_KEY = " & Session("seqTemplate") & _
                " and QUESTION_TYPE_CODE in ('S','Z','N') and FILTER_IND = 1"
            Me.sqlCommonAdapter.SelectCommand = Me.sqlCommonAction
            Me.sqlCommonAdapter.Fill(Me.dsCommon, "ListReportTop")

            If Not (Trim("" & Me.wneNonRespondent.Text) = "") Then
                'loop until we run out of non-respondents
                While (index < Me.wneNonRespondent.Value)
                    'store the respondent
                    Me.sqlCommonAction.CommandText = "Insert into " & myUtility.getExtension() & "SAT_RESPONSE (RESPONDENT_USER_KEY, SURVEY_KEY, CREATE_DATE, LAST_USED_DATE, " & _
                        "REMIND_DATE, SUBMISSION_COUNT, LAST_COMPLETION_DATE) VALUES (" & nextUserID & ", " & Session("seqSurvey") & _
                        ", '" & Me.wdcCompleteDate.Value & "', '" & Me.wdcCompleteDate.Value & "', '1/1/1970', 0, '1/1/1970')"
                    Me.sqlCommonAction.ExecuteNonQuery()

                    If (Me.dsCommon.Tables("ListReportTop").Rows.Count > 0) Then
                        Dim dr As DataRow
                        For Each dr In Me.dsCommon.Tables("ListReportTop").Rows
                            If ((dr.Item("QUESTION_TYPE_CODE") <> "Z" And dr.Item("QUESTION_TYPE_CODE") <> "N") Or _
                                (dr.Item("QUESTION_TYPE_CODE") = "Z" And Session("ddlList" & dr.Item("QUESTION_ID") & "_1") <> "") Or _
                                (dr.Item("QUESTION_TYPE_CODE") = "N" And Trim("" & Session("numField" & dr.Item("QUESTION_ID"))) <> "")) Then
                                'store the data for the items to be listed at the report top
                                Me.sqlCommonAction.CommandText = "Insert into " & myUtility.getExtension() & "SAT_RESPONSE_RESULT (SURVEY_KEY, RESPONDENT_USER_KEY, RESPONSE_DATE, " & _
                                    "QUESTION_ID, QUESTION_TYPE, ANSWER_VALUE, SUBMISSION_COUNT, ANSWER_TEXT) VALUES (" & _
                                    Session("seqSurvey") & ", " & nextUserID & ", '" & Me.wdcCompleteDate.Value & "', " & _
                                    dr.Item("QUESTION_ID") & ", '" & dr.Item("QUESTION_TYPE_CODE") & "',"

                                If (dr.Item("QUESTION_TYPE_CODE") = "N") Then
                                    Me.sqlCommonAction.CommandText &= Session("numField" & dr.Item("QUESTION_ID")) & ", 0, '')"
                                Else
                                    Me.sqlCommonAction.CommandText &= " 0, 0, '"

                                    If (dr.Item("QUESTION_TYPE_CODE") = "S") Then
                                        Me.sqlCommonAction.CommandText &= Session("txtField" & dr.Item("QUESTION_ID")) & "')"
                                    ElseIf (dr.Item("QUESTION_TYPE_CODE") = "Z") Then
                                        Me.sqlCommonAction.CommandText &= Session("ddlList" & dr.Item("QUESTION_ID") & "_1") & "')"
                                    End If
                                End If

                                'write the saved choices for the user to the database
                                Me.sqlCommonAction.ExecuteNonQuery()
                            End If
                        Next
                    End If
                    index += 1
                    nextUserID += 1
                End While
                Me.sqlCommonAction.Connection.Close()

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
