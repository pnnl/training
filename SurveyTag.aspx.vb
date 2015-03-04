Imports System.Collections.Specialized
Public Class SurveyTag
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SurveyTag))
        Me.dsCommon = New System.Data.DataSet
        Me.commonConn = New System.Data.SqlClient.SqlConnection
        Me.sqlTags = New System.Data.SqlClient.SqlDataAdapter
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
        'commonConn
        '
        Me.commonConn.FireInfoMessageEventOnUserErrors = False
        '
        'sqlTags
        '
        Me.sqlTags.DeleteCommand = Me.SqlDeleteCommand2
        Me.sqlTags.InsertCommand = Me.SqlInsertCommand2
        Me.sqlTags.SelectCommand = Me.SqlSelectCommand2
        Me.sqlTags.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_SURVEY_TAG", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("SURVEY_KEY", "SURVEY_KEY"), New System.Data.Common.DataColumnMapping("TEMPLATE_KEY", "TEMPLATE_KEY"), New System.Data.Common.DataColumnMapping("SURVEY_TAG_ID", "SURVEY_TAG_ID"), New System.Data.Common.DataColumnMapping("SURVEY_TAG_TITLE", "SURVEY_TAG_TITLE"), New System.Data.Common.DataColumnMapping("SURVEY_TAG_ANSWER", "SURVEY_TAG_ANSWER")})})
        Me.sqlTags.UpdateCommand = Me.SqlUpdateCommand2
        '
        'SqlDeleteCommand2
        '
        Me.SqlDeleteCommand2.CommandText = "DELETE FROM VARCONN.[SAT_SURVEY_TAG] WHERE (([SURVEY_KEY] = @Original_SU" & _
            "RVEY_KEY) AND ([TEMPLATE_KEY] = @Original_TEMPLATE_KEY) AND ([SURVEY_TAG_ID] = @" & _
            "Original_SURVEY_TAG_ID))"
        Me.SqlDeleteCommand2.Connection = Me.commonConn
        Me.SqlDeleteCommand2.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@Original_SURVEY_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "SURVEY_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_SURVEY_TAG_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "SURVEY_TAG_ID", System.Data.DataRowVersion.Original, Nothing)})
        '
        'SqlInsertCommand2
        '
        Me.SqlInsertCommand2.CommandText = resources.GetString("SqlInsertCommand2.CommandText")
        Me.SqlInsertCommand2.Connection = Me.commonConn
        Me.SqlInsertCommand2.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@SURVEY_KEY", System.Data.SqlDbType.Int, 0, "SURVEY_KEY"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, "TEMPLATE_KEY"), New System.Data.SqlClient.SqlParameter("@SURVEY_TAG_ID", System.Data.SqlDbType.Int, 0, "SURVEY_TAG_ID"), New System.Data.SqlClient.SqlParameter("@SURVEY_TAG_TITLE", System.Data.SqlDbType.VarChar, 0, "SURVEY_TAG_TITLE"), New System.Data.SqlClient.SqlParameter("@SURVEY_TAG_ANSWER", System.Data.SqlDbType.VarChar, 0, "SURVEY_TAG_ANSWER")})
        '
        'SqlSelectCommand2
        '
        Me.SqlSelectCommand2.CommandText = "SELECT SURVEY_KEY, TEMPLATE_KEY, SURVEY_TAG_ID, SURVEY_TAG_TITLE, SURVEY_TAG_ANSW" & _
            "ER FROM VARCONN.SAT_SURVEY_TAG"
        Me.SqlSelectCommand2.Connection = Me.commonConn
        '
        'SqlUpdateCommand2
        '
        Me.SqlUpdateCommand2.CommandText = resources.GetString("SqlUpdateCommand2.CommandText")
        Me.SqlUpdateCommand2.Connection = Me.commonConn
        Me.SqlUpdateCommand2.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@SURVEY_KEY", System.Data.SqlDbType.Int, 0, "SURVEY_KEY"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, "TEMPLATE_KEY"), New System.Data.SqlClient.SqlParameter("@SURVEY_TAG_ID", System.Data.SqlDbType.Int, 0, "SURVEY_TAG_ID"), New System.Data.SqlClient.SqlParameter("@SURVEY_TAG_TITLE", System.Data.SqlDbType.VarChar, 0, "SURVEY_TAG_TITLE"), New System.Data.SqlClient.SqlParameter("@SURVEY_TAG_ANSWER", System.Data.SqlDbType.VarChar, 0, "SURVEY_TAG_ANSWER"), New System.Data.SqlClient.SqlParameter("@Original_SURVEY_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "SURVEY_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_SURVEY_TAG_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "SURVEY_TAG_ID", System.Data.DataRowVersion.Original, Nothing)})
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected WithEvents imgTitle As System.Web.UI.WebControls.Image
    Protected WithEvents hlpServerSwitch As System.Web.UI.WebControls.Image
    Protected WithEvents chkServerSwitch As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblTemplate As System.Web.UI.WebControls.Label
    Protected WithEvents hlpTemplatePreview As System.Web.UI.WebControls.Image
    Protected WithEvents btnTemplatePreview As System.Web.UI.WebControls.Button
    Protected WithEvents btnUp As System.Web.UI.WebControls.Button
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents dsCommon As System.Data.DataSet
    Protected WithEvents lblQuestion As System.Web.UI.WebControls.Label
    Protected WithEvents sqlTags As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents SqlDeleteCommand2 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlInsertCommand2 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlSelectCommand2 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlUpdateCommand2 As System.Data.SqlClient.SqlCommand
    Protected myNavBar As Navigation = New Navigation
    Protected myUtility As Utility = New Utility
    Protected requestItems As Array
    Protected WithEvents hlpTagQuestion As System.Web.UI.WebControls.Image
    Protected WithEvents lblTagQuestionName As System.Web.UI.WebControls.Label
    Protected WithEvents hlpTagAnswer As System.Web.UI.WebControls.Image
    Protected WithEvents lblTagAnswer As System.Web.UI.WebControls.Label
    Protected WithEvents txtTagAnswer As System.Web.UI.WebControls.TextBox
    Protected WithEvents lblTagQuestion As System.Web.UI.WebControls.Label
    Protected WithEvents btnStart As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnPrev As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnReload As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnNext As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnLast As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnNew As System.Web.UI.WebControls.ImageButton
    Protected navButtonsSurvey As Collection
    Protected navButtonsTagItem As Collection
    Protected sqlCommonAdapter As New System.Data.SqlClient.SqlDataAdapter
    Protected sqlCommonAction As New System.Data.SqlClient.SqlCommand
    Protected WithEvents lblNavBarSurvey As System.Web.UI.WebControls.Label
    Protected WithEvents btnTop As System.Web.UI.WebControls.Button
    Protected WithEvents lblNavBarTagItem As System.Web.UI.WebControls.Label
    Protected WithEvents btnStart2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnPrev2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnReload2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnNext2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnLast2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnNew2 As System.Web.UI.WebControls.ImageButton
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

        If (Session.IsNewSession) Then
            Session("Alert") = "Your session has timed out for security purposes.  Please log in again."
        End If

        'get the from page
        If Not (IsPostBack) Then
            Session("FromPage") = getFromPage()
        End If

        'Set the connection based on the machine
        Me.commonConn = myUtility.getConnection(Me.commonConn)

        'determine survey ownership if not defined
        If (Session("isSurveyOwner") Is Nothing) Then
            sqlCommonAdapter.SelectCommand = sqlCommonAction
        End If

        'Put user code to initialize the page here
        If (Session("isAuthenticated") <> "True") Then
            Session("CurrentPage") = "./Login.aspx"
            Session("pageSel") = "Login"
            redirect(Session("CurrentPage"))
        ElseIf (myUtility.checkUserType(2, CType(Session("isSurveyOwner"), Boolean), CType(Session("isSurveyDelegate"), Boolean), True)) Then
            redirect(Session("CurrentPage"))
        Else
            Session("CurrentPage") = "./SurveyTag.aspx"
            Session("pageSel") = "Template"
        End If

        'if the user's password has expired send them to the password reset page
        If (Session("isExpired") = "True") Then
            Session("CurrentPage") = "./Password.aspx"
            Session("pageSel") = "Password"
            redirect(Session("CurrentPage"))
        End If

        'generate collection of navigational buttons
        myUtility.makeNavCollection(Me.navButtonsSurvey, Me.btnStart, Me.btnPrev, Me.btnReload, Me.btnNext, Me.btnLast, Me.btnNew)
        myUtility.makeNavCollection(Me.navButtonsTagItem, Me.btnStart2, Me.btnPrev2, Me.btnReload2, Me.btnNext2, Me.btnLast2, Me.btnNew2)

        'Check for user selection from the Tags list if we're sure it's not a simple post-back
        If Not (IsPostBack) Then
            'get anything we can off the request line
            myUtility.getRequest(Page.Request.Url.Query())

            'load all Tags
            If Not (loadTags()) Then
                alert("Failed to load the tag items for this survey.")
            Else
                setNavBar()
            End If

            'get the template name - mr sneaky user has access and used a link
            If (Session("TemplateName") Is Nothing Or Session("TemplateName") = "") Then
                If Not (setTemplateName()) Then
                    alert("Unable to locate template name.")
                End If
            End If

            'get the survey name - mr sneaky user has access and used a link
            If (Session("SurveyName") Is Nothing Or Session("SurveyName") = "") Then
                If Not (setSurveyName()) Then
                    alert("Unable to locate survey name.")
                End If
            End If

            'Populate the controls on the page
            setControls()

            Me.btnUpdate.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
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

    'set the template name
    Private Function setSurveyName() As Boolean
        sqlCommonAdapter.SelectCommand = sqlCommonAction
        sqlCommonAction.Connection = Me.commonConn

        Try
            sqlCommonAction.CommandText = "Select SURVEY_COMMENT from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY where " & _
            "SURVEY_KEY = " & Session("seqSurvey")
            sqlCommonAdapter.Fill(Me.dsCommon, "SurveyName")
            Session("SurveyName") = Me.dsCommon.Tables("SurveyName").Rows(0).Item("SURVEY_COMMENT")
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function
#End Region

#Region "Data Load"
    'loads the comment tags
    Private Function loadTags() As Boolean
        Trace.Warn("Loading Tags")
        Try
            'dump the data that may exists before refilling
            If (Me.dsCommon.Tables.Contains("Comments")) Then
                Me.dsCommon.Tables("Comments").Rows.Clear()
            End If

            'set the rest of the command string here for stability of the code behind page
            Dim oldCommand As String = Me.SqlSelectCommand2.CommandText
            Me.SqlSelectCommand2.CommandText &= " WHERE TEMPLATE_KEY=" & Session("seqTemplate") & _
            " and SURVEY_KEY = " & Session("seqSurvey") & " ORDER BY SURVEY_TAG_ID"
            Me.SqlSelectCommand2.CommandText = Me.SqlSelectCommand2.CommandText.Replace("VARCONN.", myUtility.getExtension())
            Me.sqlTags.SelectCommand = Me.SqlSelectCommand2
            Trace.Warn(Me.sqlTags.SelectCommand.CommandText)
            Me.sqlTags.SelectCommand.Connection = Me.commonConn
            Me.sqlTags.Fill(Me.dsCommon, "Comments")
            Me.SqlSelectCommand2.CommandText = oldCommand
            Session("CommentMax") = Me.dsCommon.Tables("Comments").Rows.Count()

            'determine if there is any data and if the comment id has been set
            If (Session("CommentMax") = 0) Then
                Session("intCommentID") = -1
            ElseIf (Session("intCommentID") Is Nothing) Then
                Session("intCommentID") = 1
            End If

            'make sure the comment row never exceeds the maximum comments
            If (Session("CommentRow") > Session("CommentMax")) Then
                Session("CommentRow") = Session("CommentMax") - 1
            End If

            'if the message id indicates a new record then move the comment row to the maximum
            'else set it to just below the comment id's number
            If (Session("intCommentID") = -1) Then
                Session("CommentRow") = Session("CommentMax")
            Else
                Session("CommentRow") = Session("intCommentID") - 1
            End If
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'gets the survey number for the current template
    Private Function getSurvey(ByVal whatToGet As String) As Integer
        Try
            'get the surveys for this template based on who is accessing them
            If (Session("UserType") = 4) Then
                sqlCommonAction.CommandText = "SELECT SURVEY_KEY FROM " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY " & _
                " WHERE TEMPLATE_KEY = " & Session("seqTemplate")
            Else
                sqlCommonAction.CommandText = "SELECT fs.SURVEY_KEY FROM " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY fs " & _
                "inner join " & myUtility.getExtension() & "SAT_USER_PERMISSION fp on fs.SURVEY_KEY = fp.SURVEY_KEY " & _
                "Where (fs.TEMPLATE_KEY = " & Session("seqTemplate") & " And fp.USER_KEY = " & _
                Session("UserID") & " And " & "(fp.OWNER_IND = 1 Or fp.DELEGATE_IND = 1))"
            End If
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            sqlCommonAction.Connection = Me.commonConn
            sqlCommonAdapter.Fill(Me.dsCommon, "Surveys")

            'get the appropriate survey sequence number for the survey page on error assume new
            If (Me.dsCommon.Tables("Surveys").Rows.Count > 0) Then
                If (whatToGet = "First") Then
                    Return Me.dsCommon.Tables("Surveys").Rows(0).Item("SURVEY_KEY")
                ElseIf (whatToGet = "Previous") Then
                    Return Me.dsCommon.Tables("Surveys").Rows(Session("SurveyRow") - 1).Item("SURVEY_KEY")
                ElseIf (whatToGet = "Next") Then
                    Return Me.dsCommon.Tables("Surveys").Rows(Session("SurveyRow") + 1).Item("SURVEY_KEY")
                ElseIf (whatToGet = "Last") Then
                    Return Me.dsCommon.Tables("Surveys").Rows(Session("SurveyMax") - 1).Item("SURVEY_KEY")
                End If
            Else
                alert("Failed to locate the selected survey.")
                Return -1
            End If
        Catch ex As Exception
            'on error assume new
            Trace.Warn(ex.ToString)
            alert("Failed to locate the selected survey.")
            Return -1
        End Try
    End Function
#End Region

#Region "Control Settings"
    'populates the controls with the selected tag item
    Private Sub setControls()
        Trace.Warn("Setting Controls")
        Try
            Me.txtTagAnswer.Text = Me.dsCommon.Tables("Comments").Rows(Session("CommentRow")).Item("SURVEY_TAG_ANSWER")
            Me.lblTagQuestion.Text = Me.dsCommon.Tables("Comments").Rows(Session("CommentRow")).Item("SURVEY_TAG_TITLE")
            Me.lblNavBarSurvey.Text = "Survey " & Session("SurveyRow") + 1 & " of " & Session("SurveyMax")
            Me.lblNavBarTagItem.Text = "Tag Item " & Session("commentRow") + 1 & " of " & Session("CommentMax")
        Catch ex As Exception
            Trace.Warn(ex.ToString)
        End Try
    End Sub
#End Region

#Region "Page Switch"
    'Takes the user back to the template they were working on
    Private Sub btnTop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTop.Click
        Session("JSAction") = ""
        Session("CurrentPage") = "./Template.aspx"
        Session("pageSel") = "Template"
        redirect(Session("CurrentPage"))
    End Sub

    'Takes the user back to the template they were working on
    Private Sub btnUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUp.Click
        Session("JSAction") = ""
        Session("CurrentPage") = "./Survey.aspx"
        Session("pageSel") = "Template"
        redirect(Session("CurrentPage"))
    End Sub

    'brings up a pop-up window with the template preview in it
    Private Sub btnTemplatePreview_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTemplatePreview.Click
        Session("JSAction") = "<script>newWin = window.open('./Preview.aspx', 'PreviewWindow', 'width=600,height=400,scrollbars=yes,toolbar=no,titlebar=no,status=no,resizable=yes,menubar=no');newWin.focus();</script>"
    End Sub
#End Region

#Region "Settings"
    'to determine what of the nav bar buttons should be available
    Private Sub setNavBar()
        'to determine what of the nav bar buttons should be available
        myNavBar.wnb_MoveTo(Me.navButtonsSurvey, Session("SurveyMax"), Session("SurveyRow"), True, False)
        myNavBar.wnb_MoveTo(Me.navButtonsTagItem, Session("CommentMax"), Session("CommentRow"), True, True)
    End Sub
#End Region

#Region "Nav Bar Events Survey"
    'Moves the user to the first record
    Private Sub wnbSurveys_MoveToStart(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnStart.Click
        Session("JSAction") = ""
        Session("CurrentPage") = "./Survey.aspx?seqTemplate=" & Session("seqTemplate") & "&seqSurvey=" & getSurvey("First")
        Session("pageSel") = "Template"
        redirect(Session("CurrentPage"))
    End Sub

    'Moves the user to the previous record
    Private Sub wnbSurveys_MovePrev(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnPrev.Click
        Session("JSAction") = ""
        Session("CurrentPage") = "./Survey.aspx?seqTemplate=" & Session("seqTemplate") & "&seqSurvey=" & getSurvey("Previous")
        Session("pageSel") = "Template"
        redirect(Session("CurrentPage"))
    End Sub

    'reloads the page and the data from the database while attempting to locate the user's location
    Private Sub wnbSurveys_MoveReload(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnReload.Click
        Session("JSAction") = ""
        Session("CurrentPage") = "./Survey.aspx?seqTemplate=" & Session("seqTemplate") & "&seqSurvey=" & Session("seqSurvey")
        Session("pageSel") = "Template"
        redirect(Session("CurrentPage"))
    End Sub

    'Moves the user to the next record
    Private Sub wnbSurveys_MoveNext(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnNext.Click
        Session("JSAction") = ""
        Session("CurrentPage") = "./Survey.aspx?seqTemplate=" & Session("seqTemplate") & "&seqSurvey=" & getSurvey("Next")
        Session("pageSel") = "Template"
        redirect(Session("CurrentPage"))
    End Sub

    'Moves the user to the last record
    Private Sub wnbSurveys_MoveToEnd(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnLast.Click
        Session("JSAction") = ""
        Session("CurrentPage") = "./Survey.aspx?seqTemplate=" & Session("seqTemplate") & "&seqSurvey=" & getSurvey("Last")
        Session("pageSel") = "Template"
        redirect(Session("CurrentPage"))
    End Sub

    'Inserts a record into the table
    Private Sub wnbSurveys_InsertRow(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnNew.Click
        Session("JSAction") = ""
        Session("CurrentPage") = "./Survey.aspx?seqTemplate=" & Session("seqTemplate") & "&seqSurvey=-1"
        Session("pageSel") = "Template"
        redirect(Session("CurrentPage"))
    End Sub
#End Region

#Region "Nav Bar Events"
    'Moves the user to the first record
    Private Sub wnbTags_MoveToStart(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnStart2.Click
        Session("JSAction") = ""
        loadTags()
        Session("CommentRow") = 0
        Session("intCommentID") = Me.dsCommon.Tables("Comments").Rows(Session("CommentRow")).Item("SURVEY_TAG_ID")
        setNavBar()
        setControls()
    End Sub

    'Moves the user to the previous record
    Private Sub wnbTags_MovePrev(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnPrev2.Click
        Session("JSAction") = ""
        loadTags()
        Session("CommentRow") -= 1
        Session("intCommentID") = Me.dsCommon.Tables("Comments").Rows(Session("CommentRow")).Item("SURVEY_TAG_ID")
        setNavBar()
        setControls()
    End Sub

    'reloads the page and the data from the database while attempting to locate the user's location
    Private Sub wnbTags_MoveReload(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnReload2.Click
        Session("JSAction") = ""
        loadTags()
        setNavBar()
        setControls()
    End Sub

    'Moves the user to the next record
    Private Sub wnbTags_MoveNext(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnNext2.Click
        Session("JSAction") = ""
        loadTags()
        Session("CommentRow") += 1
        Session("intCommentID") = Me.dsCommon.Tables("Comments").Rows(Session("CommentRow")).Item("SURVEY_TAG_ID")
        setNavBar()
        setControls()
    End Sub

    'Moves the user to the last record
    Private Sub wnbTags_MoveToEnd(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnLast2.Click
        Session("JSAction") = ""
        loadTags()
        Session("CommentRow") = Session("CommentMax") - 1
        Session("intCommentID") = Me.dsCommon.Tables("Comments").Rows(Session("CommentRow")).Item("SURVEY_TAG_ID")
        setNavBar()
        setControls()
    End Sub
#End Region

#Region "Page Action"
    'Reacts to the user performing some kind of selection
    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'refresh the Tags table
        loadTags()

        'check for possible malicious code entries
        If Not (myUtility.goodInput(Me.txtTagAnswer, True)) Then
            blnContinue = False
            alert("Possible malicious code entry in the tag answer.  Update aborted.")
        End If

        'do the update
        If (blnContinue) Then
            failure = doUpdate()

            'reload the data
            loadTags()

            'to determine what of the nav bar buttons should be available
            setNavBar()

            'reset the page controls
            If Not (failure) Then
                setControls()
            End If
            Session("TransExists") = False
        End If
    End Sub

    'Reacts to the user performing some kind of selection
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'check for possible malicious code entries
        If Not (myUtility.goodInput(Me.txtTagAnswer, True)) Then
            blnContinue = False
            alert("Possible malicious code entry in the tag answer.  Reset aborted.")
        End If

        'do the reset
        If (blnContinue) Then
            failure = doReset()

            'to determine what of the nav bar buttons should be available
            setNavBar()

            'reset the page controls
            If Not (failure) Then
                setControls()
            End If
            Session("TransExists") = False
        End If
    End Sub
#End Region

#Region "Update Code"
    'drives the commit and roll-back operations of the update
    Private Function doUpdate() As Boolean
        Try
            'start transaction
            If (myUtility.setupTransaction(sqlCommonAction, Me.commonConn)) Then
                Dim failure As Boolean = False
                'Check for actual information to update
                If (Me.txtTagAnswer.Text <> "") Then
                    If (updateTags(sqlCommonAction, sqlCommonAdapter)) Then
                        sqlCommonAction.Transaction.Commit()
                    Else
                        sqlCommonAction.Transaction.Rollback()
                        alert("Tag update failure.")
                        failure = True
                    End If
                Else
                    sqlCommonAction.Transaction.Rollback()
                    alert("You must answer the question to update this tag item.")
                    failure = True
                End If
                Return failure
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Tag update failure.")
                Return True
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Tag update failure.")
            Return True
        End Try
    End Function

    'attempts to update the Tags in the survey tags tables, returns false if it cannot
    Private Function updateTags(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'parameterized the text input to allow for things like quotes to get through
            Dim param2 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@SURVEY_TAG_ANSWER", System.Data.SqlDbType.VarChar, 1200, "SURVEY_TAG_ANSWER")
            param2.Value = Me.txtTagAnswer.Text
            sqlCommonAction.Parameters.Add(param2)

            'update the comment in each survey
            Dim currentSurvey As Integer = -1
            Dim surveyMax As Integer = Me.dsCommon.Tables("Comments").Rows.Count()
            Dim surveyIndex As Integer = 0
            While (surveyIndex < surveyMax)
                If (Session("intCommentID") = Me.dsCommon.Tables("Comments").Rows(surveyIndex).Item("SURVEY_TAG_ID")) Then
                    Try
                        sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_SURVEY_TAG SET SURVEY_TAG_ANSWER = @SURVEY_TAG_ANSWER" & _
                        " WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and SURVEY_KEY = " & _
                        Session("seqSurvey") & " and SURVEY_TAG_ID = " & Me.dsCommon.Tables("Comments").Rows(surveyIndex).Item("SURVEY_TAG_ID")
                        sqlCommonAction.ExecuteNonQuery()
                    Catch ex As Exception
                        Trace.Warn(ex.ToString)
                        Return False
                    End Try
                End If
                surveyIndex += 1
            End While
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
        loadTags()
        Return False
    End Function
#End Region
End Class