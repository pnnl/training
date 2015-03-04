Imports System.Text
Imports System.Collections.Specialized
Public Class Survey
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Survey))
        Me.sqlSurvey = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlDeleteCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlInsertCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlUpdateCommand1 = New System.Data.SqlClient.SqlCommand
        Me.commonConn = New System.Data.SqlClient.SqlConnection
        Me.dsCommon = New System.Data.DataSet
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'sqlSurvey
        '
        Me.sqlSurvey.DeleteCommand = Me.SqlDeleteCommand1
        Me.sqlSurvey.InsertCommand = Me.SqlInsertCommand1
        Me.sqlSurvey.SelectCommand = Me.SqlSelectCommand1
        Me.sqlSurvey.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_TEMPLATE_SURVEY", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("SURVEY_KEY", "SURVEY_KEY"), New System.Data.Common.DataColumnMapping("CREATOR_USER_KEY", "CREATOR_USER_KEY"), New System.Data.Common.DataColumnMapping("TEMPLATE_KEY", "TEMPLATE_KEY"), New System.Data.Common.DataColumnMapping("SURVEY_CLOSING_DATE", "SURVEY_CLOSING_DATE"), New System.Data.Common.DataColumnMapping("SURVEY_COMMENT", "SURVEY_COMMENT"), New System.Data.Common.DataColumnMapping("SURVEY_ID", "SURVEY_ID"), New System.Data.Common.DataColumnMapping("COURSE", "COURSE"), New System.Data.Common.DataColumnMapping("OPTIONAL_IND", "OPTIONAL_IND"), New System.Data.Common.DataColumnMapping("ACTIVE_IND", "ACTIVE_IND"), New System.Data.Common.DataColumnMapping("SURVEY_COLOR_CODE", "SURVEY_COLOR_CODE"), New System.Data.Common.DataColumnMapping("SUBMISSION_COUNT", "SUBMISSION_COUNT"), New System.Data.Common.DataColumnMapping("SURVEY_CREATE_DATE", "SURVEY_CREATE_DATE")})})
        Me.sqlSurvey.UpdateCommand = Me.SqlUpdateCommand1
        '
        'SqlDeleteCommand1
        '
        Me.SqlDeleteCommand1.CommandText = "DELETE FROM VARCONN.[SAT_TEMPLATE_SURVEY] WHERE (([SURVEY_KEY] = @Origin" & _
            "al_SURVEY_KEY) AND ([CREATOR_USER_KEY] = @Original_CREATOR_USER_KEY))"
        Me.SqlDeleteCommand1.Connection = Me.commonConn
        Me.SqlDeleteCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@Original_SURVEY_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "SURVEY_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_CREATOR_USER_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CREATOR_USER_KEY", System.Data.DataRowVersion.Original, Nothing)})
        '
        'SqlInsertCommand1
        '
        Me.SqlInsertCommand1.CommandText = resources.GetString("SqlInsertCommand1.CommandText")
        Me.SqlInsertCommand1.Connection = Me.commonConn
        Me.SqlInsertCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@CREATOR_USER_KEY", System.Data.SqlDbType.Int, 0, "CREATOR_USER_KEY"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, "TEMPLATE_KEY"), New System.Data.SqlClient.SqlParameter("@SURVEY_CLOSING_DATE", System.Data.SqlDbType.DateTime, 0, "SURVEY_CLOSING_DATE"), New System.Data.SqlClient.SqlParameter("@SURVEY_COMMENT", System.Data.SqlDbType.VarChar, 0, "SURVEY_COMMENT"), New System.Data.SqlClient.SqlParameter("@SURVEY_ID", System.Data.SqlDbType.VarChar, 0, "SURVEY_ID"), New System.Data.SqlClient.SqlParameter("@COURSE", System.Data.SqlDbType.[Char], 0, "COURSE"), New System.Data.SqlClient.SqlParameter("@OPTIONAL_IND", System.Data.SqlDbType.Bit, 0, "OPTIONAL_IND"), New System.Data.SqlClient.SqlParameter("@ACTIVE_IND", System.Data.SqlDbType.Bit, 0, "ACTIVE_IND"), New System.Data.SqlClient.SqlParameter("@SURVEY_COLOR_CODE", System.Data.SqlDbType.VarChar, 0, "SURVEY_COLOR_CODE"), New System.Data.SqlClient.SqlParameter("@SUBMISSION_COUNT", System.Data.SqlDbType.Int, 0, "SUBMISSION_COUNT"), New System.Data.SqlClient.SqlParameter("@SURVEY_CREATE_DATE", System.Data.SqlDbType.DateTime, 0, "SURVEY_CREATE_DATE")})
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
        Me.SqlUpdateCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@CREATOR_USER_KEY", System.Data.SqlDbType.Int, 0, "CREATOR_USER_KEY"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, "TEMPLATE_KEY"), New System.Data.SqlClient.SqlParameter("@SURVEY_CLOSING_DATE", System.Data.SqlDbType.DateTime, 0, "SURVEY_CLOSING_DATE"), New System.Data.SqlClient.SqlParameter("@SURVEY_COMMENT", System.Data.SqlDbType.VarChar, 0, "SURVEY_COMMENT"), New System.Data.SqlClient.SqlParameter("@SURVEY_ID", System.Data.SqlDbType.VarChar, 0, "SURVEY_ID"), New System.Data.SqlClient.SqlParameter("@COURSE", System.Data.SqlDbType.[Char], 0, "COURSE"), New System.Data.SqlClient.SqlParameter("@OPTIONAL_IND", System.Data.SqlDbType.Bit, 0, "OPTIONAL_IND"), New System.Data.SqlClient.SqlParameter("@ACTIVE_IND", System.Data.SqlDbType.Bit, 0, "ACTIVE_IND"), New System.Data.SqlClient.SqlParameter("@SURVEY_COLOR_CODE", System.Data.SqlDbType.VarChar, 0, "SURVEY_COLOR_CODE"), New System.Data.SqlClient.SqlParameter("@SUBMISSION_COUNT", System.Data.SqlDbType.Int, 0, "SUBMISSION_COUNT"), New System.Data.SqlClient.SqlParameter("@SURVEY_CREATE_DATE", System.Data.SqlDbType.DateTime, 0, "SURVEY_CREATE_DATE"), New System.Data.SqlClient.SqlParameter("@Original_SURVEY_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "SURVEY_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_CREATOR_USER_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CREATOR_USER_KEY", System.Data.DataRowVersion.Original, Nothing)})
        '
        'commonConn
        '
        Me.commonConn.FireInfoMessageEventOnUserErrors = False
        '
        'dsCommon
        '
        Me.dsCommon.DataSetName = "NewDataSet"
        Me.dsCommon.Locale = New System.Globalization.CultureInfo("en-US")
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected WithEvents sqlSurvey As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlInsertCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlUpdateCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlDeleteCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents dsCommon As System.Data.DataSet
    Protected myNavBar As Navigation = New Navigation
    Protected myUtility As Utility = New Utility
    Protected WithEvents hlpTagItems As System.Web.UI.WebControls.Image
    Protected WithEvents rptrTags As System.Web.UI.WebControls.Repeater
    Protected WithEvents rptrSurveyGroups As System.Web.UI.WebControls.Repeater
    Protected WithEvents imgSurveyTitle As System.Web.UI.WebControls.Image
    Protected WithEvents hlpServerSwitch As System.Web.UI.WebControls.Image
    Protected WithEvents chkServerSwitch As System.Web.UI.WebControls.CheckBox
    Protected WithEvents hlpCopy As System.Web.UI.WebControls.Image
    Protected WithEvents btnCopy As System.Web.UI.WebControls.Button
    Protected WithEvents lblTemplate As System.Web.UI.WebControls.Label
    Protected WithEvents hlpTemplatePreview As System.Web.UI.WebControls.Image
    Protected WithEvents btnTemplatePreview As System.Web.UI.WebControls.Button
    Protected WithEvents lblNavBar As System.Web.UI.WebControls.Label
    Protected WithEvents hlpSurveyName As System.Web.UI.WebControls.Image
    Protected WithEvents lblSurveyName As System.Web.UI.WebControls.Label
    Protected WithEvents txtSurveyName As System.Web.UI.WebControls.TextBox
    Protected WithEvents hlpSurveyID As System.Web.UI.WebControls.Image
    Protected WithEvents lblSurveyID As System.Web.UI.WebControls.Label
    Protected WithEvents txtSurveyID As System.Web.UI.WebControls.TextBox
    Protected WithEvents hlpClosingDate As System.Web.UI.WebControls.Image
    Protected WithEvents lblClosingDate As System.Web.UI.WebControls.Label
    Protected WithEvents wdcClosingDate As Infragistics.WebUI.WebSchedule.WebDateChooser
    Protected WithEvents hlpCourseNumber As System.Web.UI.WebControls.Image
    Protected WithEvents lblCourseNumber As System.Web.UI.WebControls.Label
    Protected WithEvents hlpOptional As System.Web.UI.WebControls.Image
    Protected WithEvents chkOptional As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblOptional As System.Web.UI.WebControls.Label
    Protected WithEvents hlpActive As System.Web.UI.WebControls.Image
    Protected WithEvents chkActive As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblActive As System.Web.UI.WebControls.Label
    Protected WithEvents hlpSurveys As System.Web.UI.WebControls.Image
    Protected WithEvents lblTagItems As System.Web.UI.WebControls.Label
    Protected WithEvents hlpDelegates As System.Web.UI.WebControls.Image
    Protected WithEvents lblDelegates As System.Web.UI.WebControls.Label
    Protected WithEvents rptrDelegates As System.Web.UI.WebControls.Repeater
    Protected WithEvents hlpSurveyGroups As System.Web.UI.WebControls.Image
    Protected WithEvents lblSurveyGroups As System.Web.UI.WebControls.Label
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected requestItems As Array
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents btnStart As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnPrev As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnReload As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnNext As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnLast As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnNew As System.Web.UI.WebControls.ImageButton
    Protected navButtons As Collection
    Protected WithEvents hlpColorPicker As System.Web.UI.WebControls.Image
    Protected WithEvents lblColorPicker As System.Web.UI.WebControls.Label
    Protected WithEvents ddlColorPicker As System.Web.UI.HtmlControls.HtmlSelect
    Protected sqlCommonAdapter As New System.Data.SqlClient.SqlDataAdapter
    Protected sqlCommonAction As New System.Data.SqlClient.SqlCommand
    Protected WithEvents wneCourseNumber As Infragistics.WebUI.WebDataInput.WebNumericEdit
    Protected WithEvents btnTop As System.Web.UI.WebControls.Button
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
            myUtility.getRequest(Page.Request.Url.Query())
        End If

        'Set the connection based on the machine
        Me.commonConn = myUtility.getConnection(Me.commonConn)

        'determine survey ownership if not defined
        isOwner()

        'Put user code to initialize the page here
        If (Session("isAuthenticated") <> "True") Then
            Session("CurrentPage") = "./Login.aspx"
            Session("pageSel") = "Login"
            Response.Redirect(Session("CurrentPage"))
        ElseIf (myUtility.checkUserType(2, CType(Session("isSurveyOwner"), Boolean), CType(Session("isSurveyDelegate"), Boolean), True)) Then
            redirect(Session("CurrentPage"))
        Else
            Session("CurrentPage") = "./Survey.aspx"
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
            'load all Comments
            If Not (loadData()) Then
                alert("Failed to load the surveys for this template.")
            Else
                'get the survey results
                surveyResults()

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
                myUtility.optionPreSet(Session("seqSurvey"), Session("SurveyMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions

                Me.btnInsert.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                Me.btnUpdate.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                Me.btnDelete.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                Me.btnReset.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
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
    Private Function loadData(Optional ByVal override As Boolean = False) As Boolean
        Trace.Warn("Loading Data")
        Try
            'Set up the common data adapter and select command
            sqlCommonAction.Connection = Me.commonConn
            sqlCommonAdapter.SelectCommand = sqlCommonAction

            'load the colors
            If (loadColors(override)) Then
                'load the surveys
                If (loadSurveys(override)) Then
                    If (Session("seqSurvey") <> -1) Then
                        'Fill the Tags list
                        If (loadTags()) Then
                            'Fill the Delegates list
                            If (loadDelegates()) Then
                                'Fill the Survey Groups list
                                If (loadSurveyGroups()) Then
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
                        Return True
                    End If
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    'Loads data into the form
    Private Function loadDataLite(Optional ByVal override As Boolean = False) As Boolean
        Trace.Warn("Loading Data")
        Try
            'Set up the common data adapter and select command
            sqlCommonAction.Connection = Me.commonConn
            sqlCommonAdapter.SelectCommand = sqlCommonAction

            'load the colors
            If (loadColors(override)) Then
                'load the surveys
                If (Session("seqSurvey") <> -1) Then
                    'Fill the Tags list
                    If (loadTags()) Then
                        'Fill the Delegates list
                        If (loadDelegates()) Then
                            'Fill the Survey Groups list
                            If (loadSurveyGroups()) Then
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
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    'loads the surveys
    Private Function loadSurveys(Optional ByVal override As Boolean = False) As Boolean
        Trace.Warn("Loading Surveys")
        Try
            'try to erase anything if it exists in hte table so we don't end up with duplicate data
            If (Me.dsCommon.Tables.Contains("Surveys")) Then
                Me.dsCommon.Tables("Surveys").Rows.Clear()
            End If

            'select the surveys for the current template
            Dim oldCommand As String
            If (Session("UserType") = 4) Then
                oldCommand = Me.SqlSelectCommand1.CommandText
                Me.SqlSelectCommand1.CommandText &= " WHERE TEMPLATE_KEY = " & Session("seqTemplate")
            Else
                oldCommand = Me.SqlSelectCommand1.CommandText
                Me.SqlSelectCommand1.CommandText = "SELECT fs.SURVEY_KEY, fs.CREATOR_USER_KEY, fs.TEMPLATE_KEY, fs.SURVEY_CLOSING_DATE, " & _
                "fs.SURVEY_COMMENT, isnull(fs.SURVEY_ID, '') as SURVEY_ID, isnull(fs.COURSE, 0) as COURSE, fs.OPTIONAL_IND, fs.ACTIVE_IND, fs.SURVEY_COLOR_CODE, fs.SUBMISSION_COUNT, fs.SURVEY_CREATE_DATE " & _
                "FROM " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY fs inner join " & myUtility.getExtension() & "SAT_USER_PERMISSION fp on fs.SURVEY_KEY = fp.SURVEY_KEY " & _
                "Where (fs.TEMPLATE_KEY = " & Session("seqTemplate") & " And fp.USER_KEY = " & Session("UserID") & " And " & _
                "(fp.OWNER_IND = 1 Or fp.DELEGATE_IND = 1))"
            End If
            Me.sqlSurvey.SelectCommand.Connection = Me.commonConn
            Me.sqlSurvey.SelectCommand.CommandText = Me.sqlSurvey.SelectCommand.CommandText.Replace("VARCONN.", myUtility.getExtension())
            Me.sqlSurvey.Fill(Me.dsCommon, "Surveys")
            Me.SqlSelectCommand1.CommandText = oldCommand
            Session("SurveyMax") = Me.dsCommon.Tables("Surveys").Rows.Count()

            'determine if there is any data and if the survey id has been set
            If (Session("SurveyMax") = 0) Then
                Session("seqSurvey") = -1
            ElseIf (Session("seqSurvey") Is Nothing) Then
                Session("seqSurvey") = 1
            End If

            'make sure the survey row never exceeds the maximum surveys
            If (Session("SurveyRow") > Session("SurveyMax")) Then
                Session("SurveyRow") = Session("SurveyMax") - 1
            End If

            'if the survey id indicates a new record then move the survey row to the maximum
            'else set it to just below the survey id's number
            If (Session("seqSurvey") = -1) Then
                Session("SurveyRow") = Session("SurveyMax")
            Else
                Session("SurveyRow") = findRow("Surveys", Session("seqSurvey"))
            End If

            'set the survey name
            If (Session("seqSurvey") <> -1) Then
                Session("SurveyName") = Me.dsCommon.Tables("Surveys").Rows(Session("SurveyRow")).Item("SURVEY_COMMENT")
            End If

            'set the question type drop-down list
            If Not (override) And Session("seqSurvey") <> "-1" Then
                setColor()
            End If

            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'loads the tag items for the current survey
    Private Function loadTags() As Boolean
        Trace.Warn("Loading Tag Items")
        Try
            If (Me.dsCommon.Tables.Contains("Tags")) Then
                Me.dsCommon.Tables("Tags").Rows.Clear()
            End If
            sqlCommonAction.CommandText = "SELECT * from " & myUtility.getExtension() & "SAT_SURVEY_TAG WHERE TEMPLATE_KEY=" & Session("seqTemplate") & _
            " and SURVEY_KEY = " & Session("seqSurvey") & " ORDER BY SURVEY_TAG_ID"
            sqlCommonAdapter.Fill(Me.dsCommon, "Tags")
            Me.rptrTags.DataSource = Me.dsCommon.Tables("Tags")
            Me.rptrTags.DataBind()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Unable to load the tag items for this survey.")
            Return False
        End Try
    End Function

    'loads the delegates for the current survey
    Private Function loadDelegates() As Boolean
        Trace.Warn("Loading Delegates")
        Try
            If (Me.dsCommon.Tables.Contains("Delegates")) Then
                Me.dsCommon.Tables("Delegates").Rows.Clear()
            End If
            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER_PERMISSION fp, " & myUtility.getExtension() & "SAT_USER fu where SURVEY_KEY = " & Session("seqSurvey") & _
            " and fp.USER_KEY = fu.USER_KEY and (fp.OWNER_IND = 1 or fp.DELEGATE_IND = 1) order by fu.LAST_NAME, fu.FIRST_NAME"
            sqlCommonAdapter.Fill(Me.dsCommon, "Delegates")
            Me.rptrDelegates.DataSource = Me.dsCommon.Tables("Delegates")
            Me.rptrDelegates.DataBind()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Unable to load the delegates for this survey.")
            Return False
        End Try
    End Function

    'loads the survey groups for the current survey
    Private Function loadSurveyGroups() As Boolean
        Trace.Warn("Loading Survey Groups")
        Try
            If (Me.dsCommon.Tables.Contains("SurveyGroups")) Then
                Me.dsCommon.Tables("SurveyGroups").Rows.Clear()
            End If
            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_SURVEY_DIST_GROUP where SURVEY_KEY = " & Session("seqSurvey") & _
            " order by GROUP_TITLE"
            sqlCommonAdapter.Fill(Me.dsCommon, "SurveyGroups")
            Me.rptrSurveyGroups.DataSource = Me.dsCommon.Tables("SurveyGroups")
            Me.rptrSurveyGroups.DataBind()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Unable to load the survey groups for this survey.")
            Return False
        End Try
    End Function

    'loads the colors for the color picker
    Private Function loadColors(ByVal override As Boolean) As Boolean
        Trace.Warn("Loading Colors")
        Try
            If Not (override) Then
                'clear out existing data
                If (Me.dsCommon.Tables.Contains("Colors")) Then
                    Me.dsCommon.Tables("Colors").Rows.Clear()
                End If
                Me.ddlColorPicker.Items.Clear()

                'get the user levels available
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "webSafeColors order by red, green, blue"
                sqlCommonAdapter.Fill(Me.dsCommon, "Colors")

                'initialize the color drop-down
                Dim li As System.Web.UI.WebControls.ListItem = New System.Web.UI.WebControls.ListItem
                li.Text = "--Select a Color--"
                li.Value = -1
                Me.ddlColorPicker.Items.Insert(0, li)

                'load the colors into the drop down
                Dim dr As DataRow
                For Each dr In Me.dsCommon.Tables("Colors").Rows
                    li = New System.Web.UI.WebControls.ListItem
                    'build opposition color
                    Dim strbldrColor As New StringBuilder
                    strbldrColor.Append(Hex(255 - CType(dr.Item("Red"), Integer)).PadLeft(2, "0"))
                    strbldrColor.Append(Hex(255 - CType(dr.Item("Green"), Integer)).PadLeft(2, "0"))
                    strbldrColor.Append(Hex(255 - CType(dr.Item("Blue"), Integer)).PadLeft(2, "0"))

                    'set list item styles and save list item to drop-down list
                    li.Attributes.Add("style", "background-color:" & dr.Item("Hexadecimal") & ";color:#" & strbldrColor.ToString & ";text-align:center")
                    li.Text = dr.Item("Hexadecimal")
                    li.Value = dr.Item("Hexadecimal")
                    Me.ddlColorPicker.Items.Add(li)
                Next
                Me.ddlColorPicker.DataBind()
                Return True
            Else
                Return True
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Unable to load the colors")
            Return False
        End Try
    End Function

    'get the owner's information
    Private Function getOwnerInfo(ByRef strOwnerEmail As String, ByRef strOwnerName As String, ByRef intOwnerType As Integer, Optional ByVal isOwner As Boolean = False) As Boolean
        Try
            Dim ds As DataSet = New DataSet

            'get the owner's information
            If (isOwner) Then
                'get the user's information
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER where USER_KEY = " & Session("UserID")
                sqlCommonAdapter.Fill(ds, "UserInfo")
            Else
                'find the real owner of this template and his/her information
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER_PERMISSION fp where fp.TEMPLATE_KEY = " & _
                    Session("seqTemplate") & " and fp.OWNER_IND = 1"
                sqlCommonAdapter.Fill(ds, "Owner")
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER fu where fu.USER_KEY = " & _
                    ds.Tables("Owner").Rows(0).Item("USER_KEY")
                sqlCommonAdapter.Fill(ds, "UserInfo")
            End If

            'store the e-mail, user name, and user type
            intOwnerType = ds.Tables("UserInfo").Rows(0).Item("USER_CODE")
            strOwnerEmail = ds.Tables("UserInfo").Rows(0).Item("EMAIL_ADDRESS")
            'If the user has a name in the DB use it.  Otherwise use the HID if it exists. Otherwise use the stored email
            If ((Trim("" & ds.Tables("UserInfo").Rows(0).Item("FIRST_NAME")) = "") _
                    And (Trim("" & ds.Tables("UserInfo").Rows(0).Item("LAST_NAME")) = "") _
                    And (Trim("" & ds.Tables("UserInfo").Rows(0).Item("MIDDLE_NAME")) = "") _
                    And (Trim("" & ds.Tables("UserInfo").Rows(0).Item("SUFFIX_NAME")) = "")) Then
                If (Trim("" & ds.Tables("UserInfo").Rows(0).Item("HANFORD_ID")) = "") Then
                    strOwnerName = strOwnerEmail
                Else
                    strOwnerName = ds.Tables("UserInfo").Rows(0).Item("HANFORD_ID")
                End If
            Else
                strOwnerName = ds.Tables("UserInfo").Rows(0).Item("LAST_NAME") & ", " & _
                        ds.Tables("UserInfo").Rows(0).Item("FIRST_NAME") & " " & _
                        ds.Tables("UserInfo").Rows(0).Item("MIDDLE_NAME") & " " & _
                        ds.Tables("UserInfo").Rows(0).Item("SUFFIX_NAME")
            End If
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function
#End Region

#Region "Control Settings"
    'populates the controls with the selected comment item
    Private Sub setControls()
        Trace.Warn("Setting Controls")
        'set the basic page controls based on whether or not we are working with a new survey
        If (Session("seqSurvey") > 0) Then
            Try
                Me.txtSurveyName.Text = Me.dsCommon.Tables("Surveys").Rows(Session("SurveyRow")).Item("SURVEY_COMMENT")
                Me.wdcClosingDate.Value = Me.dsCommon.Tables("Surveys").Rows(Session("SurveyRow")).Item("SURVEY_CLOSING_DATE")
                Me.txtSurveyID.Text = Me.dsCommon.Tables("Surveys").Rows(Session("SurveyRow")).Item("SURVEY_ID")
                Me.wneCourseNumber.Text = Me.dsCommon.Tables("Surveys").Rows(Session("SurveyRow")).Item("COURSE")
                Me.chkActive.Checked = Me.dsCommon.Tables("Surveys").Rows(Session("SurveyRow")).Item("ACTIVE_IND")
                Me.chkOptional.Checked = Me.dsCommon.Tables("Surveys").Rows(Session("SurveyRow")).Item("OPTIONAL_IND")
                Me.lblNavBar.Text = "Survey " & Session("SurveyRow") + 1 & " of " & Session("SurveyMax")
            Catch ex As Exception
                Trace.Warn(ex.ToString)
            End Try
        Else
            Me.txtSurveyName.Text = ""
            Me.wdcClosingDate.Value = Now
            Me.txtSurveyID.Text = ""
            Me.wneCourseNumber.Text = ""
            Me.chkActive.Checked = 1
            Me.chkOptional.Checked = 0
            Me.lblNavBar.Text = "Survey " & Session("SurveyRow") + 1 & " of " & Session("SurveyMax") + 1
        End If

        'disable inputs on production
        If ((Session("Results") <= 0 Or Session("UserType") = 4)) Then
            Me.txtSurveyName.Enabled = True
            Me.wdcClosingDate.Enabled = True
            Me.txtSurveyID.Enabled = True
            Me.wneCourseNumber.Enabled = True
            Me.chkActive.Enabled = True
            Me.chkOptional.Enabled = True
            Me.ddlColorPicker.Disabled = False
        Else
            Me.txtSurveyName.Enabled = False
            Me.wdcClosingDate.Enabled = False
            Me.txtSurveyID.Enabled = False
            Me.wneCourseNumber.Enabled = False
            Me.chkActive.Enabled = True
            Me.chkOptional.Enabled = False
            Me.ddlColorPicker.Disabled = True
        End If

        'set the page options
        myUtility.optionPreSet(Session("seqSurvey"), Session("SurveyMax"), Me.pageOptions)

        Session("PageOptions") = Me.pageOptions
    End Sub
#End Region

#Region "Checks"
    'determines if the user has owner/delegate access
    Private Sub isOwner()
        sqlCommonAdapter.SelectCommand = sqlCommonAction
        sqlCommonAction.Connection = Me.commonConn
        myUtility.determineUserOwnership(sqlCommonAction, sqlCommonAdapter)
    End Sub

    'get the results for the current survey
    Private Sub surveyResults()
        'Get count of survey results from production
        sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY fs, " & myUtility.getExtension() & "SAT_RESPONSE_RESULT fr where TEMPLATE_KEY = " & _
            Session("seqTemplate") & " and fs.SURVEY_KEY = fr.SURVEY_KEY and fs.SURVEY_KEY = " & Session("seqSurvey")
        Trace.Warn(sqlCommonAction.CommandText)
        sqlCommonAdapter.Fill(Me.dsCommon, "SurveyResults")
        Session("Results") = Me.dsCommon.Tables("SurveyResults").Rows.Count()
    End Sub
#End Region

#Region "Find"
    'finds the row of the survey we're looking for
    Private Function findRow(ByVal strTable As String, ByVal userID As Integer) As Integer
        Dim intRow As Integer = 0
        Dim dr As DataRow
        For Each dr In Me.dsCommon.Tables(strTable).Rows()
            If (dr.Item("SURVEY_KEY") = userID) Then
                Return intRow
            Else
                intRow += 1
            End If
        Next
        Return -1
    End Function
#End Region

#Region "Page Switch"
    'Takes the user back to the template they were working on
    Private Sub btnTop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTop.Click
        Session("CurrentPage") = "./Template.aspx"
        Session("pageSel") = "Template"
        Response.Redirect(Session("CurrentPage"))
    End Sub

    'brings up a pop-up window with the template preview in it
    Private Sub btnTemplatePreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTemplatePreview.Click
        Session("JSAction") = ""
        Session("BackgroundColor") = Me.ddlColorPicker.Value
        Session("JSAction") = "<script>newWin = window.open('./Preview.aspx', 'PreviewWindow', 'width=600,height=400,scrollbars=yes,toolbar=no,titlebar=no,status=no,resizable=yes,menubar=no');newWin.focus();</script>"
        If (loadData()) Then
            'get the survey results
            surveyResults()

            'to determine what of the nav bar buttons should be available
            setNavBar()

            'reset the page controls
            setControls()
            Session("TransExists") = False
            'set the page options
            myUtility.optionPreSet(Session("seqSurvey"), Session("SurveyMax"), Me.pageOptions)

            Session("PageOptions") = Me.pageOptions
        Else
            alert("Error loading the surveys for this template.")
        End If
    End Sub
#End Region

#Region "Settings"
    'sets the selected color on the drop-down list
    Private Function setColor() As Boolean
        Trace.Warn("Setting Color")

        'locate and select the current color
        Dim item As ListItem
        Dim index As Integer = 0
        For Each item In Me.ddlColorPicker.Items
            If (Me.dsCommon.Tables("Surveys").Rows(Session("SurveyRow")).Item("SURVEY_COLOR_CODE") = item.Value) Then
                Me.ddlColorPicker.SelectedIndex = index
                Return True
            Else
                index += 1
            End If
        Next
        Me.ddlColorPicker.SelectedIndex = 0
        Return True
    End Function

    'to determine what of the nav bar buttons should be available
    Private Sub setNavBar()
        'to determine what of the nav bar buttons should be available
        myNavBar.wnb_MoveTo(Me.navButtons, Session("SurveyMax"), Session("SurveyRow"), Switch)
    End Sub
#End Region

#Region "Nav Bar Events"
    'Moves the user to the first record
    Private Sub wnbSurveys_MoveToStart(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnStart.Click
        Session("JSAction") = ""

        If Not (loadSurveys()) Then
            alert("Failed to load the delegates/owners for this survey.")
        Else
            Try
                Session("SurveyRow") = 0
                Session("seqSurvey") = Me.dsCommon.Tables("Surveys").Rows(Session("SurveyRow")).Item("SURVEY_KEY")
                If Not (loadDataLite()) Then
                    alert("Failed to load the delegates/owners for this survey.")
                Else
                    surveyResults()
                    isOwner()
                    Session("SurveyName") = Me.dsCommon.Tables("Surveys").Rows(Session("SurveyRow")).Item("SURVEY_COMMENT")
                    setColor()
                    setNavBar()
                    setControls()
                    'set the page options
                    myUtility.optionPreSet(Session("seqSurvey"), Session("SurveyMax"), Me.pageOptions)

                    Session("PageOptions") = Me.pageOptions
                End If
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the first survey.")
            End Try
        End If
    End Sub

    'Moves the user to the previous record
    Private Sub wnbSurveys_MovePrev(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnPrev.Click
        Session("JSAction") = ""

        If Not (loadSurveys()) Then
            alert("Failed to load the delegates/owners for this survey.")
        Else
            Try
                Session("SurveyRow") -= 1
                Session("seqSurvey") = Me.dsCommon.Tables("Surveys").Rows(Session("SurveyRow")).Item("SURVEY_KEY")
                If Not (loadDataLite()) Then
                    alert("Failed to load the delegates/owners for this survey.")
                Else
                    surveyResults()
                    isOwner()
                    Session("SurveyName") = Me.dsCommon.Tables("Surveys").Rows(Session("SurveyRow")).Item("SURVEY_COMMENT")
                    setColor()
                    setNavBar()
                    setControls()
                    'set the page options
                    myUtility.optionPreSet(Session("seqSurvey"), Session("SurveyMax"), Me.pageOptions)

                    Session("PageOptions") = Me.pageOptions
                End If
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the previous survey.")
            End Try
        End If
    End Sub

    'reloads the page and the data from the database while attempting to locate the user's location
    Private Sub wnbSurveys_MoveReload(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnReload.Click
        Session("JSAction") = ""

        If Not (loadData()) Then
            alert("Failed to load the delegates/owners for this survey.")
        Else
            Try
                isOwner()
                If (Session("seqSurvey") <> -1) Then
                    setColor()
                End If
                surveyResults()
                setNavBar()
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("seqSurvey"), Session("SurveyMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error resetting the current survey.")
            End Try
        End If
    End Sub

    'Moves the user to the next record
    Private Sub wnbSurveys_MoveNext(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnNext.Click
        Session("JSAction") = ""

        If Not (loadSurveys()) Then
            alert("Failed to load the delegates/owners for this survey.")
        Else
            Try
                Session("SurveyRow") += 1
                Session("seqSurvey") = Me.dsCommon.Tables("Surveys").Rows(Session("SurveyRow")).Item("SURVEY_KEY")
                If Not (loadDataLite()) Then
                    alert("Failed to load the delegates/owners for this survey.")
                Else
                    surveyResults()
                    isOwner()
                    Session("SurveyName") = Me.dsCommon.Tables("Surveys").Rows(Session("SurveyRow")).Item("SURVEY_COMMENT")
                    setColor()
                    setNavBar()
                    setControls()
                    'set the page options
                    myUtility.optionPreSet(Session("seqSurvey"), Session("SurveyMax"), Me.pageOptions)

                    Session("PageOptions") = Me.pageOptions
                End If
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the next survey.")
            End Try
        End If
    End Sub

    'Moves the user to the last record
    Private Sub wnbSurveys_MoveToEnd(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnLast.Click
        Session("JSAction") = ""

        If Not (loadSurveys()) Then
            alert("Failed to load the delegates/owners for this survey.")
        Else
            Try
                Session("SurveyRow") = Session("SurveyMax") - 1
                Session("seqSurvey") = Me.dsCommon.Tables("Surveys").Rows(Session("SurveyRow")).Item("SURVEY_KEY")
                If Not (loadDataLite()) Then
                    alert("Failed to load the delegates/owners for this survey.")
                Else
                    surveyResults()
                    isOwner()
                    Session("SurveyName") = Me.dsCommon.Tables("Surveys").Rows(Session("SurveyRow")).Item("SURVEY_COMMENT")
                    setColor()
                    setNavBar()
                    setControls()
                    'set the page options
                    myUtility.optionPreSet(Session("seqSurvey"), Session("SurveyMax"), Me.pageOptions)

                    Session("PageOptions") = Me.pageOptions
                End If
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the last survey.")
            End Try
        End If
    End Sub

    'Inserts a record into the table
    Private Sub wnbSurveys_InsertRow(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnNew.Click
        Session("JSAction") = ""

        If Not (loadSurveys()) Then
            alert("Failed to load the delegates/owners for this survey.")
        Else
            Try
                Session("seqSurvey") = -1
                If Not (loadDataLite()) Then
                    alert("Failed to load the delegates/owners for this survey.")
                Else
                    surveyResults()
                    Session("SurveyRow") = Session("SurveyMax")
                    Me.dsCommon.Tables("Surveys").Rows.Add(Me.dsCommon.Tables("Surveys").NewRow())
                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("SURVEY_KEY") = Session("seqSurvey")
                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("CREATOR_USER_KEY") = Session("UserID")
                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("TEMPLATE_KEY") = Session("seqTemplate")
                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("SURVEY_COMMENT") = ""
                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("SURVEY_ID") = ""
                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("COURSE") = ""
                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("OPTIONAL_IND") = 0
                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("ACTIVE_IND") = 1
                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("SUBMISSION_COUNT") = 0
                    myNavBar.wnb_MoveTo(Me.navButtons, Session("SurveyMax"), Session("SurveyRow"))
                    Me.ddlColorPicker.SelectedIndex = 0
                    setNavBar()
                    setControls()
                    'set the page options
                    myUtility.optionPreSet(Session("seqSurvey"), Session("SurveyMax"), Me.pageOptions)

                    Session("PageOptions") = Me.pageOptions
                End If
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error creating new survey.")
            End Try
        End If
    End Sub
#End Region

#Region "Page Action"
    'Reacts to the user performing some kind of selection
    Private Sub btnInsert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsert.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'refresh all of the page data
        If Not (loadData(True)) Then
            blnContinue = False
            alert("Failed to load the surveys for the current template. Insert aborted.")
        End If

        'check for possible malicious code entries
        If Not (myUtility.goodInput(Me.txtSurveyName, True)) Then
            blnContinue = False
            alert("Possible malicious code entry in the survey name.  Insert aborted.")
        End If

        If (blnContinue) Then
            'do the insert
            failure = doInsert()

            'reload the data
            If Not (loadData()) Then
                blnContinue = False
                alert("Failed to load the surveys for the current template.")
            End If

            If (blnContinue) Then
                'get the survey results
                surveyResults()

                'to determine what of the nav bar buttons should be available
                setNavBar()

                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
                Session("TransExists") = False
                'set the page options
                myUtility.optionPreSet(Session("seqSurvey"), Session("SurveyMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection
    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'refresh all of the page data
        If Not (loadData(True)) Then
            blnContinue = False
            alert("Failed to load the surveys for the current template. Update aborted.")
        End If

        'check for possible malicious code entries
        If Not (myUtility.goodInput(Me.txtSurveyName, True)) Then
            blnContinue = False
            alert("Possible malicious code entry in the survey name.  Update aborted.")
        End If

        If (blnContinue) Then
            'do the update
            failure = doUpdate()

            'reload the data
            If Not (loadData()) Then
                blnContinue = False
                alert("Failed to load the surveys for the current template.")
            End If

            If (blnContinue) Then
                'get the survey results
                surveyResults()

                'to determine what of the nav bar buttons should be available
                setNavBar()

                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
                Session("TransExists") = False
                'set the page options
                myUtility.optionPreSet(Session("seqSurvey"), Session("SurveyMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'refresh all of the page data
        If Not (loadData(True)) Then
            blnContinue = False
            alert("Failed to load the surveys for the current template. Delete aborted.")
        End If

        'check for possible malicious code entries
        If Not (myUtility.goodInput(Me.txtSurveyName, True)) Then
            blnContinue = False
            alert("Possible malicious code entry in the survey name.  Delete aborted.")
        End If

        If (blnContinue) Then
            'do the delete
            failure = doDelete()

            'reload the data
            If Not (loadData()) Then
                blnContinue = False
                alert("Failed to load the surveys for the current template.")
            End If

            If (blnContinue) Then
                'get the survey results
                surveyResults()

                'to determine what of the nav bar buttons should be available
                setNavBar()

                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
                Session("TransExists") = False
                'set the page options
                myUtility.optionPreSet(Session("seqSurvey"), Session("SurveyMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'check for possible malicious code entries
        If Not (myUtility.goodInput(Me.txtSurveyName, True)) Then
            blnContinue = False
            alert("Possible malicious code entry in the survey name.  Reset aborted.")
        End If

        If (blnContinue) Then
            'do the reset
            failure = doReset()

            If (blnContinue) Then
                'get the survey results
                surveyResults()

                'to determine what of the nav bar buttons should be available
                setNavBar()

                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
                Session("TransExists") = False
                'set the page options
                myUtility.optionPreSet(Session("seqSurvey"), Session("SurveyMax"), Me.pageOptions)

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
                If (Me.txtSurveyName.Text <> "") Then
                    If (Session("seqSurvey") = -1) Then
                        'If we're inserting a new record then add it onto the end of the recordset and add it to the existing surveys
                        If (insertSurvey(sqlCommonAction, sqlCommonAdapter)) Then
                            If (insertTags(sqlCommonAction, sqlCommonAdapter)) Then
                                If (insertPermissions(sqlCommonAction, sqlCommonAdapter)) Then
                                    Session("SurveyMax") += 1
                                    Session("seqSurvey") = Session("newSeqSurvey")
                                    sqlCommonAction.Transaction.Commit()
                                Else
                                    sqlCommonAction.Transaction.Rollback()
                                    alert("Error adding survey to the template.")
                                    failure = True
                                End If
                            Else
                                sqlCommonAction.Transaction.Rollback()
                                alert("Error adding survey to the template.")
                                failure = True
                            End If
                        Else
                            sqlCommonAction.Transaction.Rollback()
                            alert("Error adding survey to the template.")
                            failure = True
                        End If
                    Else
                        'If we're inserting an existing record then insert a copy of it into the database and add it to the existing surveys
                        If (insertSurvey(sqlCommonAction, sqlCommonAdapter)) Then
                            If (insertTags(sqlCommonAction, sqlCommonAdapter)) Then
                                If (insertPermissions(sqlCommonAction, sqlCommonAdapter)) Then
                                    Session("SurveyMax") += 1
                                    Session("seqSurvey") = Session("newSeqSurvey")
                                    sqlCommonAction.Transaction.Commit()
                                Else
                                    sqlCommonAction.Transaction.Rollback()
                                    alert("Error adding survey to the template.")
                                    failure = True
                                End If
                            Else
                                sqlCommonAction.Transaction.Rollback()
                                alert("Error adding survey to the template.")
                                failure = True
                            End If
                        Else
                            sqlCommonAction.Transaction.Rollback()
                            alert("Error adding survey to the template.")
                            failure = True
                        End If
                    End If
                Else
                    alert("You must have some text in the Survey Name field to add this Survey to the Template.")
                    failure = True
                    sqlCommonAction.Transaction.Rollback()
                End If
                Return failure
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Error adding survey to the template.")
                Return True
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Error adding survey to the template.")
            Return True
        End Try
    End Function

    'attempts to insert a survey, returns false if it cannot
    Private Function insertSurvey(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'parameterized the text input to allow for things like quotes to get through
            Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@SURVEY_COMMENT", System.Data.SqlDbType.VarChar, 1800, "SURVEY_COMMENT")
            param0.Value = Me.txtSurveyName.Text
            sqlCommonAction.Parameters.Add(param0)
            Dim param1 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@SURVEY_ID", System.Data.SqlDbType.VarChar, 16, "SURVEY_ID")
            param1.Value = Me.txtSurveyID.Text
            sqlCommonAction.Parameters.Add(param1)
            Dim param2 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@COURSE", System.Data.SqlDbType.Char, 6, "COURSE")
            param2.Value = Me.wneCourseNumber.Text
            sqlCommonAction.Parameters.Add(param2)

            'attempt to add the new record to the database
            If (Me.ddlColorPicker.Value <> "-1") Then
                sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY (CREATOR_USER_KEY, TEMPLATE_KEY, SURVEY_CLOSING_DATE, SURVEY_COMMENT, " & _
                "SURVEY_ID, COURSE, OPTIONAL_IND, ACTIVE_IND, SUBMISSION_COUNT, SURVEY_CREATE_DATE, SURVEY_COLOR_CODE) VALUES (" & Session("UserID") & ", " & _
                Session("seqTemplate") & ", '" & Me.wdcClosingDate.Value & "', @SURVEY_COMMENT, @SURVEY_ID, @COURSE, " & _
                IIf(Me.chkOptional.Checked, 1, 0) & ", " & IIf(Me.chkActive.Checked, 1, 0) & ", 0, '" & Now & "', '" & _
                Me.ddlColorPicker.Value & "')"
            Else
                sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY (CREATOR_USER_KEY, TEMPLATE_KEY, SURVEY_CLOSING_DATE, SURVEY_COMMENT, " & _
                "SURVEY_ID, COURSE, OPTIONAL_IND, ACTIVE_IND, SUBMISSION_COUNT, SURVEY_CREATE_DATE) VALUES (" & Session("UserID") & ", " & _
                Session("seqTemplate") & ", '" & Me.wdcClosingDate.Value & "', @SURVEY_COMMENT, @SURVEY_ID, @COURSE, " & _
                IIf(Me.chkOptional.Checked, 1, 0) & ", " & IIf(Me.chkActive.Checked, 1, 0) & ", 0, '" & Now & "')"
            End If
            sqlCommonAction.ExecuteNonQuery()

            'get the survey's sequence number
            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY where TEMPLATE_KEY = " & Session("seqTemplate") & _
            " and SURVEY_COMMENT = @SURVEY_COMMENT order by SURVEY_CREATE_DATE desc"
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            sqlCommonAdapter.Fill(Me.dsCommon, "NewSurvey")
            Session("newSeqSurvey") = Me.dsCommon.Tables("NewSurvey").Rows(0).Item("SURVEY_KEY")
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'attempts to insert the comments into the survey tags tables, returns false if it cannot
    Private Function insertTags(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'parameterized the text input to allow for things like quotes to get through
            sqlCommonAction.Parameters.Clear()
            Dim param3 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@SURVEY_TAG_TITLE", System.Data.SqlDbType.VarChar, 150, "SURVEY_TAG_TITLE")
            sqlCommonAction.Parameters.Add(param3)

            'add the template comments to the new survey
            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_TEMPLATE_TAG where TEMPLATE_KEY = " & Session("seqTemplate") & _
            " Order by TEMPLATE_TAG_ID"
            param3.Value = ""
            sqlCommonAdapter.Fill(Me.dsCommon, "Comments")

            'go through each comment and add each for the new survey to the tags table
            Dim dr As DataRow
            For Each dr In Me.dsCommon.Tables("Comments").Rows()
                sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_SURVEY_TAG (SURVEY_KEY, TEMPLATE_KEY, SURVEY_TAG_ID, SURVEY_TAG_TITLE, " & _
                "SURVEY_TAG_ANSWER) VALUES (" & Session("newSeqSurvey") & ", " & Session("seqTemplate") & ", " & dr.Item("TEMPLATE_TAG_ID") & _
                ", @SURVEY_TAG_TITLE, '')"
                param3.Value = dr.Item("TEMPLATE_TAG_TITLE")
                sqlCommonAction.ExecuteNonQuery()
            Next
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'attempts to insert permissions for the creator into the permissions table, returns false if it cannot
    Private Function insertPermissions(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'parameterized the text input to allow for things like quotes to get through
            sqlCommonAction.Parameters.Clear()

            'Add permissions for the survey
            sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_USER_PERMISSION (USER_KEY, TEMPLATE_KEY, SURVEY_KEY, CREATOR_USER_KEY, " & _
                "CREATE_DATE, OWNER_IND, DELEGATE_IND, USER_IND) VALUES (" & Session("UserID") & ", -1" & _
                ", " & Session("newSeqSurvey") & ", " & Session("UserID") & ", '" & Now & "', 1, 1, 1)"
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
                If (Me.txtSurveyName.Text <> "") Then
                    If (updateSurvey(sqlCommonAction, sqlCommonAdapter)) Then
                        sqlCommonAction.Transaction.Commit()
                    Else
                        sqlCommonAction.Transaction.Rollback()
                        alert("Error updating survey for this template.")
                        failure = True
                    End If
                Else
                    alert("You must have some text in the Survey Name field to update this Survey.")
                    failure = True
                    sqlCommonAction.Transaction.Rollback()
                End If
                Return failure
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Error updating survey in the template.")
                Return True
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Error updating survey in the template.")
            Return True
        End Try
    End Function

    'attempts to update a survey, returns false if it cannot
    Private Function updateSurvey(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        'attempt to update the record in the database
        Try
            'parameterized the text input to allow for things like quotes to get through
            Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@SURVEY_COMMENT", System.Data.SqlDbType.VarChar, 50, "SURVEY_COMMENT")
            param0.Value = Me.txtSurveyName.Text
            sqlCommonAction.Parameters.Add(param0)

            If (Me.ddlColorPicker.Value <> "-1") Then
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY SET SURVEY_COMMENT = @SURVEY_COMMENT, SURVEY_ID = '" & _
                Me.txtSurveyID.Text & "', COURSE = " & Me.wneCourseNumber.Value & _
                ", SURVEY_CLOSING_DATE = '" & Me.wdcClosingDate.Value & "', OPTIONAL_IND = " & IIf(Me.chkOptional.Checked, 1, 0) & _
                ", ACTIVE_IND = " & IIf(Me.chkActive.Checked, 1, 0) & ", SURVEY_COLOR_CODE = '" & Me.ddlColorPicker.Value & _
                "' WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and SURVEY_KEY = " & Session("seqSurvey")
            Else
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY SET SURVEY_COMMENT = @SURVEY_COMMENT, SURVEY_ID = '" & _
                Me.txtSurveyID.Text & "', COURSE = " & Me.wneCourseNumber.Value & _
                ", SURVEY_CLOSING_DATE = '" & Me.wdcClosingDate.Value & "', OPTIONAL_IND = " & IIf(Me.chkOptional.Checked, 1, 0) & _
                ", ACTIVE_IND = " & IIf(Me.chkActive.Checked, 1, 0) & ", SURVEY_COLOR_CODE = '#0000FF'" & _
                " WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and SURVEY_KEY = " & Session("seqSurvey")
            End If
            Trace.Warn(sqlCommonAction.CommandText)
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function
#End Region

#Region "Delete Code"
    'deletes the current survey from the survey, tags, delegates, and survey groups tables
    Private Function doDelete() As Boolean
        Try
            'set up the transaction
            If (myUtility.setupTransaction(sqlCommonAction, Me.commonConn)) Then
                Dim failure As Boolean = False

                'do the deletion
                If (deleteTags(sqlCommonAction, sqlCommonAdapter)) Then
                    If (deleteDelegates(sqlCommonAction, sqlCommonAdapter)) Then
                        If (deleteSurveyGroups(sqlCommonAction, sqlCommonAdapter)) Then
                            If (deleteSurvey(sqlCommonAction, sqlCommonAdapter)) Then
                                sqlCommonAction.Transaction.Commit()
                                'determine if we need to take a step back 
                                If (Session("SurveyMax") = 1) Then
                                    Session("SurveyRow") = 0
                                    Session("seqSurvey") = -1
                                    Session("SurveyMax") = 0
                                    Me.dsCommon.Tables("Surveys").Rows.Add(Me.dsCommon.Tables("Surveys").NewRow())
                                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("SURVEY_KEY") = Session("seqSurvey")
                                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("CREATOR_USER_KEY") = Session("UserID")
                                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("TEMPLATE_KEY") = Session("seqTemplate")
                                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("SURVEY_COMMENT") = ""
                                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("SURVEY_ID") = ""
                                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("COURSE") = ""
                                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("OPTIONAL_IND") = 0
                                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("ACTIVE_IND") = 1
                                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("SUBMISSION_COUNT") = 0
                                    Me.ddlColorPicker.SelectedIndex = 0
                                ElseIf ((Session("SurveyRow") + 1) = Session("SurveyMax")) Then
                                    Session("SurveyMax") -= 1
                                    Session("SurveyRow") -= 1
                                    Session("seqSurvey") = Me.dsCommon.Tables("Surveys").Rows(Session("SurveyRow")).Item("SURVEY_KEY")
                                Else
                                    Session("SurveyMax") -= 1
                                    Session("seqSurvey") = Me.dsCommon.Tables("Surveys").Rows(Session("SurveyRow") + 1).Item("SURVEY_KEY")
                                    Session("SurveyName") = Me.dsCommon.Tables("Surveys").Rows(Session("SurveyRow") + 1).Item("SURVEY_COMMENT")
                                End If

                                'notify the survey owners of the loss of their survey
                                If (Me.dsCommon.Tables.Contains("SurveyNotifyList")) Then
                                    If (Me.dsCommon.Tables("SurveyNotifyList").Rows.Count > 0) Then
                                        Dim row As DataRow
                                        For Each row In Me.dsCommon.Tables("SurveyNotifyList").Rows()
                                            If Not (sendDeleteMail(row.Item("EMAIL_ADDRESS"), row.Item("SURVEY_COMMENT"))) Then
                                                alert("Failed to notify " & _
                                                row.Item("FIRST_NAME") & " " & row.Item("LAST_NAME") & _
                                                " about the deletion of the survey called " & _
                                                row.Item("SURVEY_COMMENT") & ".")
                                            End If
                                        Next
                                    End If
                                End If
                            Else
                                sqlCommonAction.Transaction.Rollback()
                                alert("Error deleting survey from this template.")
                                failure = True
                            End If
                        Else
                            sqlCommonAction.Transaction.Rollback()
                            alert("Error deleting survey from this template.")
                            failure = True
                        End If
                    Else
                        sqlCommonAction.Transaction.Rollback()
                        alert("Error deleting survey from this template.")
                        failure = True
                    End If
                Else
                    sqlCommonAction.Transaction.Rollback()
                    alert("Error deleting survey from this template.")
                    failure = True
                End If
                Return failure
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Error deleting survey from template.")
                Return True
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Error deleting survey from template.")
            Return True
        End Try
    End Function

    'attempts to delete a survey, returns false if it cannot
    Private Function deleteSurvey(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, Optional ByVal isCopy As Boolean = False) As Boolean
        'attempt to delete the record from the database
        sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY where TEMPLATE_KEY = " & Session("seqTemplate") & _
            " and SURVEY_KEY = " & Session("seqSurvey")

        Try
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to remove the survey for the template.")
            Return False
        End Try
    End Function

    'attempts to delete the comments in the survey tags tables, returns false if it cannot
    Private Function deleteTags(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, Optional ByVal isCopy As Boolean = False) As Boolean
        'attempt to delete the record from the database
        sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_SURVEY_TAG where TEMPLATE_KEY = " & Session("seqTemplate") & _
                " and SURVEY_KEY = " & Session("seqSurvey")

        Try
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Failed to remove the survey for the template.")
            Return False
        End Try
    End Function

    'attempts to delete the delegates, owners, and users of the survey from the permissions table
    Private Function deleteDelegates(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        'attempt to delete the record from the database
        Try
            'determine if the survey still exists on the other side

            sqlCommonAction.CommandText = "SELECT * from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY where SURVEY_KEY = " & _
                Session("seqSurvey")
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            sqlCommonAdapter.Fill(Me.dsCommon, "SurveyExistence")

            If (Me.dsCommon.Tables("SurveyExistence").Rows.Count() = 0) Then
                'make a list of all individuals that will need to be notified of the loss of their surveys
                sqlCommonAction.CommandText = "SELECT fp.USER_KEY, fs.SURVEY_COMMENT, fu.EMAIL_ADDRESS, " & _
                "fu.LAST_NAME, fu.FIRST_NAME, fu.MIDDLE_NAME from " & _
                "" & myUtility.getExtension() & "SAT_USER_PERMISSION fp, " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY fs, " & _
                myUtility.getExtension() & "SAT_USER fu where fp.USER_KEY = " & _
                "fu.USER_KEY And fs.SURVEY_KEY = fp.SURVEY_KEY and fs.SURVEY_KEY = " & Session("seqSurvey") & _
                " and (fp.OWNER_IND = 1 or fp.DELEGATE_IND = 1) order by fp.SURVEY_KEY"
                sqlCommonAdapter.SelectCommand = sqlCommonAction
                sqlCommonAdapter.Fill(Me.dsCommon, "SurveyNotifyList")

                sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_USER_PERMISSION where TEMPLATE_KEY = -1 and SURVEY_KEY = " & Session("seqSurvey")
                sqlCommonAction.ExecuteNonQuery()
            End If
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'attempts to delete the survey groups set up for sending out this survey
    Private Function deleteSurveyGroups(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        'attempt to delete the record from the database
        Try
            sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_SURVEY_DIST_GROUP where SURVEY_KEY = " & Session("seqSurvey")
            sqlCommonAction.ExecuteNonQuery()
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
        Return True
    End Function
#End Region

#Region "Reset Code"
    'resets the page, will remove any new item being worked on at the end of the list
    Private Function doReset() As Boolean
        If Not (loadData()) Then
            alert("Failed to load the delegates/owners for this survey.")
        Else
            Try
                If (Session("seqSurvey") = -1) Then
                    Me.dsCommon.Tables("Surveys").Rows.Add(Me.dsCommon.Tables("Surveys").NewRow())
                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("SURVEY_KEY") = Session("seqSurvey")
                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("CREATOR_USER_KEY") = Session("UserID")
                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("TEMPLATE_KEY") = Session("seqTemplate")
                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("SURVEY_COMMENT") = ""
                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("SURVEY_ID") = ""
                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("COURSE") = ""
                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("OPTIONAL_IND") = 0
                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("ACTIVE_IND") = 1
                    Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax")).Item("SUBMISSION_COUNT") = 0
                    Me.ddlColorPicker.SelectedIndex = 0
                End If
                Return False
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error resetting the survey.")
                Return True
            End Try
        End If
    End Function
#End Region

#Region "Send Mail"
    'Sends an e-mail to users with changed passwords, except level 0 users.
    Private Function sendDeleteMail(ByVal Email As String, ByVal strItemName As String) As Boolean
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
            sqlCommonAction.Connection = Me.commonConn
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
        strbldrMessage.Append("Survey " & strItemName & " has been deleted, by")
        strbldrMessage.Append(" " & Me.dsCommon.Tables("ModifyID").Rows(0).Item("FIRST_NAME") & " ")
        strbldrMessage.Append(Me.dsCommon.Tables("ModifyID").Rows(0).Item("LAST_NAME") & ".")
        strbldrMessage.Append("  You may contact this person at this e-mail address: <b>")
        strbldrMessage.Append(Me.dsCommon.Tables("ModifyID").Rows(0).Item("EMAIL_ADDRESS") & "</b>")

        If Not (myUtility.genericSendMail(strFrom, strTo, strbldrSubject.ToString, strbldrMessage.ToString)) Then
            Return False
        Else
            Return True
        End If
    End Function
#End Region
End Class