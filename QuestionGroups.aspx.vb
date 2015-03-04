Imports System.Collections.Specialized
Public Class QuestionGroup
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(QuestionGroup))
        Me.dsCommon = New System.Data.DataSet
        Me.commonConn = New System.Data.SqlClient.SqlConnection
        Me.sqlReportGroups = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlDeleteCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlInsertCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlUpdateCommand1 = New System.Data.SqlClient.SqlCommand
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
        'sqlReportGroups
        '
        Me.sqlReportGroups.DeleteCommand = Me.SqlDeleteCommand1
        Me.sqlReportGroups.InsertCommand = Me.SqlInsertCommand1
        Me.sqlReportGroups.SelectCommand = Me.SqlSelectCommand1
        Me.sqlReportGroups.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_TEMPLATE_QUESTION_GROUP", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("QUESTION_GROUP_ID", "QUESTION_GROUP_ID"), New System.Data.Common.DataColumnMapping("QUESTION_GROUP_TITLE", "QUESTION_GROUP_TITLE"), New System.Data.Common.DataColumnMapping("TEMPLATE_KEY", "TEMPLATE_KEY")})})
        Me.sqlReportGroups.UpdateCommand = Me.SqlUpdateCommand1
        '
        'SqlDeleteCommand1
        '
        Me.SqlDeleteCommand1.CommandText = "DELETE FROM VARCONN.[SAT_TEMPLATE_QUESTION_GROUP] WHERE (([QUESTION_GROU" & _
            "P_ID] = @Original_QUESTION_GROUP_ID) AND ([TEMPLATE_KEY] = @Original_TEMPLATE_KE" & _
            "Y))"
        Me.SqlDeleteCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@Original_QUESTION_GROUP_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "QUESTION_GROUP_ID", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_KEY", System.Data.DataRowVersion.Original, Nothing)})
        '
        'SqlInsertCommand1
        '
        Me.SqlInsertCommand1.CommandText = "INSERT INTO VARCONN.[SAT_TEMPLATE_QUESTION_GROUP] ([QUESTION_GROUP_ID], " & _
            "[QUESTION_GROUP_TITLE], [TEMPLATE_KEY]) VALUES (@QUESTION_GROUP_ID, @QUESTION_GR" & _
            "OUP_TITLE, @TEMPLATE_KEY)"
        Me.SqlInsertCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@QUESTION_GROUP_ID", System.Data.SqlDbType.Int, 0, "QUESTION_GROUP_ID"), New System.Data.SqlClient.SqlParameter("@QUESTION_GROUP_TITLE", System.Data.SqlDbType.VarChar, 0, "QUESTION_GROUP_TITLE"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, "TEMPLATE_KEY")})
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = "SELECT QUESTION_GROUP_ID, QUESTION_GROUP_TITLE, TEMPLATE_KEY FROM VARCONN.SA" & _
            "T_TEMPLATE_QUESTION_GROUP"
        '
        'SqlUpdateCommand1
        '
        Me.SqlUpdateCommand1.CommandText = resources.GetString("SqlUpdateCommand1.CommandText")
        Me.SqlUpdateCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@QUESTION_GROUP_ID", System.Data.SqlDbType.Int, 0, "QUESTION_GROUP_ID"), New System.Data.SqlClient.SqlParameter("@QUESTION_GROUP_TITLE", System.Data.SqlDbType.VarChar, 0, "QUESTION_GROUP_TITLE"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, "TEMPLATE_KEY"), New System.Data.SqlClient.SqlParameter("@Original_QUESTION_GROUP_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "QUESTION_GROUP_ID", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_KEY", System.Data.DataRowVersion.Original, Nothing)})
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected WithEvents hlpServerSwitch As System.Web.UI.WebControls.Image
    Protected WithEvents chkServerSwitch As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblTemplate As System.Web.UI.WebControls.Label
    Protected WithEvents hlpTemplatePreview As System.Web.UI.WebControls.Image
    Protected WithEvents btnTemplatePreview As System.Web.UI.WebControls.Button
    Protected WithEvents lblNavBar As System.Web.UI.WebControls.Label
    Protected WithEvents btnStart As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnPrev As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnReload As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnNext As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnLast As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnNew As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnTop As System.Web.UI.WebControls.Button
    Protected WithEvents btnSubmit As System.Web.UI.WebControls.Button
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents dsCommon As System.Data.DataSet
    Protected WithEvents Td1 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected requestItems As Array
    Protected myNavBar As Navigation = New Navigation
    Protected myUtility As Utility = New Utility
    Protected navButtons As Collection
    Protected sqlCommonAdapter As New System.Data.SqlClient.SqlDataAdapter
    Protected sqlCommonAction As New System.Data.SqlClient.SqlCommand
    Protected WithEvents sqlReportGroups As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlInsertCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlUpdateCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlDeleteCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents lblQuestionGroupTitle As System.Web.UI.WebControls.Image
    Protected WithEvents hlpGroupName As System.Web.UI.WebControls.Image
    Protected WithEvents lblGroupName As System.Web.UI.WebControls.Label
    Protected WithEvents txtGroupName As System.Web.UI.WebControls.TextBox
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
    'loads the page content
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
        myUtility.getConnection(Me.commonConn)

        'determine template ownership if not defined
        If (Session("isTemplateOwner") Is Nothing) Then
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            myUtility.determineUserOwnership(sqlCommonAction, sqlCommonAdapter, True)
        End If

        'Put user code to initialize the page here
        If (Session("isAuthenticated") <> "True") Then
            Session("pageSel") = "Login"
            redirect("./Login.aspx")
        ElseIf (myUtility.checkUserType(3, CType(Session("isTemplateOwner"), Boolean), CType(Session("isTemplateDelegate"), Boolean), True)) Then
            redirect(Session("CurrentPage"))
        Else
            Session("CurrentPage") = "./Tag.aspx"
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

        'finalize the sql statements
        completeSQL()

        'Check for user selection from list items if we're sure it's not a simple post-back
        If Not (IsPostBack) Then
            'get anything on the request line
            myUtility.getRequest(Page.Request.Url.Query())

            'load all Questions
            If Not (loadData()) Then
                alert("Failed to load the question groups.")
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
                myUtility.optionPreSet(Session("intReportGroup"), Session("ReportGroupMax"), Me.pageOptions)

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

    'finalize the SQL statements - useful for not losing non-standard select statements when editing the aspx in design mode
    Private Sub completeSQL()
        Me.SqlSelectCommand1.CommandText &= " WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " ORDER BY QUESTION_GROUP_ID"
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
        'load the report group data 
        Try
            'get report group data
            If (loadQuestionGroups()) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    'loads the report groups for the current template
    Private Function loadQuestionGroups(Optional ByVal override As Boolean = False) As Boolean
        Trace.Warn("Loading report groups")
        Try
            'dump the data that may exist before refilling
            If (Me.dsCommon.Tables.Contains("QuestionGroups")) Then
                Me.dsCommon.Tables("QuestionGroups").Rows.Clear()
            End If

            Me.sqlReportGroups.SelectCommand.Connection = Me.commonConn
            Me.sqlReportGroups.SelectCommand.CommandText = Me.sqlReportGroups.SelectCommand.CommandText.Replace("VARCONN.", myUtility.getExtension())
            Me.sqlReportGroups.Fill(Me.dsCommon, "QuestionGroups")
            Session("ReportGroupMax") = Me.dsCommon.Tables("QuestionGroups").Rows.Count()

            'determine if there is any data and if the report group id has been set
            If (Session("ReportGroupMax") = 0) Then
                Session("intReportGroup") = -1
            ElseIf (Session("intReportGroup") Is Nothing) Then
                Session("intReportGroup") = 1
            End If

            'make sure the report group row never exceeeds the maximum report groups
            If (Session("ReportGroupRow") > Session("ReportGroupMax")) Then
                Session("ReportGroupRow") = Session("ReportGroupMax") - 1
            End If

            'if the report group id indicates a new record then move the report group row to the maximum
            'else set it to just below the report group id's number
            If (Session("intReportGroup") = -1) Then
                Session("ReportGroupRow") = Session("ReportGroupMax")
            Else
                Session("ReportGroupRow") = Session("intReportGroup") - 1
            End If

            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function
#End Region

#Region "Control Settings"
    'populates the controls with the selected Question item
    Private Sub setControls()
        Trace.Warn("Setting Controls")
        If (Session("intReportGroup") > 0) Then
            Try
                Me.txtGroupName.Text = Me.dsCommon.Tables("QuestionGroups").Rows(Session("ReportGroupRow")).Item("QUESTION_GROUP_TITLE")
                Me.lblNavBar.Text = "Question Group " & Session("ReportGroupRow") + 1 & " of " & Session("ReportGroupMax")
            Catch ex As Exception
                Trace.Warn(ex.ToString)
            End Try
        Else
            Me.txtGroupName.Text = ""
            Me.lblNavBar.Text = "Question Group " & Session("ReportGroupRow") + 1 & " of " & Session("ReportGroupMax") + 1
        End If

        'set the page options
        myUtility.optionPreSet(Session("intReportGroup"), Session("ReportGroupMax"), Me.pageOptions)

        Session("PageOptions") = Me.pageOptions
    End Sub
#End Region

#Region "Page Switch"
    'Takes the user back to the template they were working on
    Private Sub btnTop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTop.Click
        Session("CurrentPage") = "./Template.aspx"
        Session("pageSel") = "Template"
        Response.Redirect(Session("CurrentPage"))
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
        myNavBar.wnb_MoveTo(Me.navButtons, Session("ReportGroupMax"), Session("ReportGroupRow"), Switch)
    End Sub
#End Region

#Region "Nav Bar Events"
    'Moves the user to the first record
    Private Sub wnbQuestions_MoveToStart(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnStart.Click
        Session("JSAction") = ""

        If Not (loadQuestionGroups()) Then
            alert("Failed to load the question groups for this template.")
        Else
            Try
                Session("ReportGroupRow") = 0
                Session("intReportGroup") = Me.dsCommon.Tables("QuestionGroups").Rows(Session("ReportGroupRow")).Item("QUESTION_GROUP_ID")
                setNavBar()
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("intReportGroup"), Session("ReportGroupMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the first question group.")
            End Try
        End If
    End Sub

    'Moves the user to the previous record
    Private Sub wnbQuestions_MovePrev(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnPrev.Click
        Session("JSAction") = ""

        If Not (loadQuestionGroups()) Then
            alert("Failed to load the question groups for this template.")
        Else
            Try
                Session("ReportGroupRow") -= 1
                Session("intReportGroup") = Me.dsCommon.Tables("QuestionGroups").Rows(Session("ReportGroupRow")).Item("QUESTION_GROUP_ID")
                setNavBar()
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("intReportGroup"), Session("ReportGroupMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the previous question group.")
            End Try
        End If
    End Sub

    'reloads the page and the data from the database while attempting to locate the user's location
    Private Sub wnbQuestions_MoveReload(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnReload.Click
        Session("JSAction") = ""

        If Not (loadData()) Then
            alert("Failed to load the question groups for this template.")
        Else
            Try
                setNavBar()
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("intReportGroup"), Session("ReportGroupMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error resetting the current question group.")
            End Try
        End If
    End Sub

    'Moves the user to the next record
    Private Sub wnbQuestions_MoveNext(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnNext.Click
        Session("JSAction") = ""

        If Not (loadQuestionGroups()) Then
            alert("Failed to load the question groups for this template.")
        Else
            Try
                Session("ReportGroupRow") += 1
                Session("intReportGroup") = Me.dsCommon.Tables("QuestionGroups").Rows(Session("ReportGroupRow")).Item("QUESTION_GROUP_ID")
                setNavBar()
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("intReportGroup"), Session("ReportGroupMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the next question group.")
            End Try
        End If
    End Sub

    'Moves the user to the last record
    Private Sub wnbQuestions_MoveToEnd(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnLast.Click
        Session("JSAction") = ""

        If Not (loadQuestionGroups()) Then
            alert("Failed to load the question groups for this template.")
        Else
            Try
                Session("ReportGroupRow") = Session("ReportGroupMax") - 1
                Session("intReportGroup") = Me.dsCommon.Tables("QuestionGroups").Rows(Session("ReportGroupRow")).Item("QUESTION_GROUP_ID")
                setNavBar()
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("intReportGroup"), Session("ReportGroupMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the last question group.")
            End Try
        End If
    End Sub

    'Inserts a record into the table
    Private Sub wnbQuestions_InsertRow(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnNew.Click
        Session("JSAction") = ""

        If Not (loadData()) Then
            alert("Failed to load the question groups for this template.")
        Else
            Try
                Session("intReportGroup") = -1
                Session("ReportGroupRow") = Session("ReportGroupMax")
                Me.dsCommon.Tables("QuestionGroups").Rows.Add(Me.dsCommon.Tables("QuestionGroups").NewRow())
                Me.dsCommon.Tables("QuestionGroups").Rows(Session("ReportGroupMax")).Item("TEMPLATE_KEY") = Session("seqTemplate")
                Me.dsCommon.Tables("QuestionGroups").Rows(Session("ReportGroupMax")).Item("QUESTION_GROUP_ID") = Session("ReportGroupMax")
                Me.dsCommon.Tables("QuestionGroups").Rows(Session("ReportGroupMax")).Item("QUESTION_GROUP_TITLE") = ""
                myNavBar.wnb_MoveTo(Me.navButtons, Session("ReportGroupMax"), Session("ReportGroupRow"))
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("intReportGroup"), Session("ReportGroupMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error creating new question group.")
            End Try
        End If
    End Sub
#End Region

#Region "Page Action"
    'Reacts to the user performing some kind of selection and submission
    Private Sub btnInsert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsert.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'refresh the report groups table
        If Not (loadQuestionGroups(True)) Then
            blnContinue = False
            alert("Failed to load the question groups for the current template. Insert aborted.")
        End If

        'check for possible malicious code input
        If Not (myUtility.goodInput(Me.txtGroupName, True)) Then
            alert("Possible malicious code entry in the question group name.  Insert aborted.")
            blnContinue = False
        End If

        If (blnContinue) Then
            'determine if the template is dirty or not
            Session("isDirty") = myUtility.isDirty(Me.dsCommon, Me.commonConn)

            'do the insert
            failure = doInsert()

            'reload the data
            If Not (loadData()) Then
                blnContinue = False
                alert("Failed to load the question groups for the current template.")
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
                myUtility.optionPreSet(Session("intReportGroup"), Session("ReportGroupMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True
        
        'refresh the report groups table
        If Not (loadQuestionGroups(True)) Then
            blnContinue = False
            alert("Failed to load the question groups for the current template. Update aborted.")
        End If

        'check for possible malicious code input
        If Not (myUtility.goodInput(Me.txtGroupName, True)) Then
            alert("Possible malicious code entry in the question group name.  Update aborted.")
            blnContinue = False
        End If

        If (blnContinue) Then
            'determine if the template is dirty or not
            Session("isDirty") = myUtility.isDirty(Me.dsCommon, Me.commonConn)

            'do the update
            failure = doUpdate()

            'reload the data
            If Not (loadData()) Then
                blnContinue = False
                alert("Failed to load the question groups for the current template.")
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
                myUtility.optionPreSet(Session("intReportGroup"), Session("ReportGroupMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True
        
        'refresh the report groups table
        If Not (loadQuestionGroups(True)) Then
            blnContinue = False
            alert("Failed to load the question groups for the current template. Delete aborted.")
        End If

        'check for possible malicious code input
        If Not (myUtility.goodInput(Me.txtGroupName, True)) Then
            alert("Possible malicious code entry in the question group name. Delete aborted.")
            blnContinue = False
        End If

        If (blnContinue) Then
            'determine if the template is dirty or not
            Session("isDirty") = myUtility.isDirty(Me.dsCommon, Me.commonConn)

            'do the delete
            failure = doDelete()

            'reload the data
            If Not (loadData()) Then
                blnContinue = False
                alert("Failed to load the question groups for the current template.")
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
                myUtility.optionPreSet(Session("intReportGroup"), Session("ReportGroupMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True
        
        'check for possible malicious code input
        If Not (myUtility.goodInput(Me.txtGroupName, True)) Then
            alert("Possible malicious code entry in the question group name.  Reset aborted.")
            blnContinue = False
        End If

        If (blnContinue) Then
            'determine if the template is dirty or not
            Session("isDirty") = myUtility.isDirty(Me.dsCommon, Me.commonConn)

            'do the reset
            failure = doReset()

            If (blnContinue) Then
                'to determine what of the nav bar buttons should be available
                setNavBar()

                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
                Session("TransExists") = False
                'set the page options
                myUtility.optionPreSet(Session("intReportGroup"), Session("ReportGroupMax"), Me.pageOptions)

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
                If (Me.txtGroupName.Text <> "") Then
                    If (Session("intReportGroup") = -1) Then
                        'If we're inserting a new record then add it onto the end of the recordset
                        If (insertQuestionGroup(sqlCommonAction, sqlCommonAdapter, True)) Then
                            If Not (Session("isDirty")) Then
                                Dim tempConn As Data.SqlClient.SqlConnection = New Data.SqlClient.SqlConnection
                                If (myUtility.makeDirty(sqlCommonAction, sqlCommonAdapter, Me.dsCommon, myUtility.getConnection(tempConn))) Then
                                    If (Session("Alert") <> "") Then
                                        alert(Session("Alert"))
                                    End If
                                    Session("isDirty") = True
                                    Session("ReportGroupMax") += 1
                                    Session("intReportGroup") = Session("ReportGroupMax")
                                    sqlCommonAction.Transaction.Commit()
                                Else
                                    If (Session("Alert") <> "") Then
                                        alert(Session("Alert"))
                                    End If
                                    sqlCommonAction.Transaction.Rollback()
                                    alert("Error adding question group to the template.")
                                    failure = True
                                End If
                            Else
                                Session("ReportGroupMax") += 1
                                Session("intReportGroup") = Session("ReportGroupMax")
                                sqlCommonAction.Transaction.Commit()
                            End If
                        Else
                            sqlCommonAction.Transaction.Rollback()
                            alert("Error adding question group to the template.")
                            failure = True
                        End If
                    Else
                        'If we're inserting an existing record then inline insert it ito the record set and add it to the existing report groups
                        If (insertQuestionGroup(sqlCommonAction, sqlCommonAdapter)) Then
                            If Not (Session("isDirty")) Then
                                Dim tempConn As Data.SqlClient.SqlConnection = New Data.SqlClient.SqlConnection
                                If (myUtility.makeDirty(sqlCommonAction, sqlCommonAdapter, Me.dsCommon, myUtility.getConnection(tempConn))) Then
                                    If (Session("Alert") <> "") Then
                                        alert(Session("Alert"))
                                    End If
                                    Session("ReportGroupMax") += 1
                                    sqlCommonAction.Transaction.Commit()
                                    Session("isDirty") = True
                                Else
                                    If (Session("Alert") <> "") Then
                                        alert(Session("Alert"))
                                    End If
                                    sqlCommonAction.Transaction.Rollback()
                                    alert("Error adding question group to the template.")
                                    failure = True
                                End If
                            Else
                                Session("ReportGroupMax") += 1
                                sqlCommonAction.Transaction.Commit()
                            End If
                        Else
                            sqlCommonAction.Transaction.Rollback()
                            alert("Error adding question group to the template.")
                            failure = True
                        End If
                    End If
                Else
                    alert("You must have some text in the group name box.")
                    failure = True
                    sqlCommonAction.Transaction.Rollback()
                End If
                Return failure
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Error adding question group to the template.")
                Return True
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Error adding question group to the template.")
            Return True
        End Try
    End Function

    'attempts to insert a report group, returns false if it cannot
    Private Function insertQuestionGroup(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, Optional ByVal isNewRecord As Boolean = False) As Boolean
        Try
            'parameterized the text input to allow for things like quotes to get through
            Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@QUESTION_GROUP_TITLE", System.Data.SqlDbType.VarChar, 200, "QUESTION_GROUP_TITLE")
            param0.Value = Me.txtGroupName.Text
            sqlCommonAction.Parameters.Add(param0)

            If (isNewRecord) Then
                'attempt to add the new record to the database

                sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION_GROUP (TEMPLATE_KEY, QUESTION_GROUP_ID, " & _
                "QUESTION_GROUP_TITLE) VALUES (" & Session("seqTemplate") & ", " & Session("ReportGroupMax") + 1 & _
                ", @QUESTION_GROUP_TITLE)"

                sqlCommonAction.ExecuteNonQuery()
                Return True
            Else
                'update the report group id's
                Dim lowIndex As Integer = Session("ReportGroupRow")
                Dim highIndex As Integer = Me.dsCommon.Tables("QuestionGroups").Rows.Count() - 1

                'update the report group id's of the Questions Groups
                While (highIndex >= lowIndex)
                    sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION_GROUP SET QUESTION_GROUP_ID = (QUESTION_GROUP_ID+1) " & _
                    "WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and QUESTION_GROUP_ID = " & _
                    Me.dsCommon.Tables("QuestionGroups").Rows(highIndex).Item("QUESTION_GROUP_ID")
                    sqlCommonAction.ExecuteNonQuery()
                    highIndex -= 1
                End While

                'attempt to add the new record to the database

                sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION_GROUP (TEMPLATE_KEY, QUESTION_GROUP_ID, " & _
                "QUESTION_GROUP_TITLE) VALUES (" & Session("seqTemplate") & ", " & Session("ReportGroupRow") + 1 & _
                ", @QUESTION_GROUP_TITLE)"

                sqlCommonAction.ExecuteNonQuery()
                Return True
            End If
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
                If (Me.txtGroupName.Text <> "") Then
                    If (updateQuestionGroup(sqlCommonAction, sqlCommonAdapter)) Then
                        If Not (Session("isDirty")) Then
                            Dim tempConn As Data.SqlClient.SqlConnection = New Data.SqlClient.SqlConnection
                            If (myUtility.makeDirty(sqlCommonAction, sqlCommonAdapter, Me.dsCommon, myUtility.getConnection(tempConn))) Then
                                If (Session("Alert") <> "") Then
                                    alert(Session("Alert"))
                                End If
                                Session("isDirty") = True
                                sqlCommonAction.Transaction.Commit()
                            Else
                                If (Session("Alert") <> "") Then
                                    alert(Session("Alert"))
                                End If
                                sqlCommonAction.Transaction.Rollback()
                                alert("Error updating question group in the template.")
                                failure = True
                            End If
                        Else
                            sqlCommonAction.Transaction.Commit()
                        End If
                    Else
                        sqlCommonAction.Transaction.Rollback()
                        alert("Error updating question group in the template.")
                        failure = True
                    End If
                Else
                    sqlCommonAction.Transaction.Rollback()
                    alert("You must have some text in the question box.")
                    failure = True
                End If
                Return failure
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Error updating question group in the template.")
                Return True
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Error updating question group in the template.")
            Return True
        End Try
    End Function

    'attempts to update a Report Group, returns false if it cannot
    Private Function updateQuestionGroup(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'parameterized the text input to allow for things like quotes to get through
            Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@QUESTION_GROUP_TITLE", System.Data.SqlDbType.VarChar, 200, "QUESTION_GROUP_TITLE")
            param0.Value = Me.txtGroupName.Text
            sqlCommonAction.Parameters.Add(param0)

            'attempt to update the record in the database
            sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION_GROUP SET QUESTION_GROUP_TITLE = @QUESTION_GROUP_TITLE " & _
            "WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and QUESTION_GROUP_ID = " & Session("intReportGroup")
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
    'deletes the current item from the Question and List item tables
    Private Function doDelete() As Boolean
        Try
            'start transaction
            If (myUtility.setupTransaction(sqlCommonAction, Me.commonConn)) Then
                Dim failure As Boolean = False

                'do the deletion
                If (deleteQuestionGroup(sqlCommonAction, sqlCommonAdapter)) Then
                    If Not (Session("isDirty")) Then
                        Dim tempConn As Data.SqlClient.SqlConnection = New Data.SqlClient.SqlConnection
                        If (myUtility.makeDirty(sqlCommonAction, sqlCommonAdapter, Me.dsCommon, myUtility.getConnection(tempConn))) Then
                            If (Session("Alert") <> "") Then
                                alert(Session("Alert"))
                            End If
                            Session("isDirty") = True
                            sqlCommonAction.Transaction.Commit()
                            If (Session("ReportGroupMax") = 1) Then
                                Session("ReportGroupRow") = 0
                                Session("intReportGroup") = -1
                                Session("ReportGroupMax") = 0
                                Me.dsCommon.Tables("QuestionGroups").Rows.Add(Me.dsCommon.Tables("QuestionGroups").NewRow())
                                Me.dsCommon.Tables("QuestionGroups").Rows(Session("ReportGroupMax")).Item("TEMPLATE_KEY") = Session("seqTemplate")
                                Me.dsCommon.Tables("QuestionGroups").Rows(Session("ReportGroupMax")).Item("QUESTION_GROUP_ID") = Session("ReportGroupMax")
                                Me.dsCommon.Tables("QuestionGroups").Rows(Session("ReportGroupRow")).Item("QUESTION_GROUP_TITLE") = ""
                            ElseIf (Session("intReportGroup") = Session("ReportGroupMax") And Session("ReportGroupMax") - Session("ReportGroupRow") = 1) Then
                                Session("intReportGroup") -= 1
                                Session("ReportGroupMax") -= 1
                                Session("ReportGroupRow") -= 1
                            Else
                                Session("ReportGroupMax") -= 1
                            End If
                        Else
                            If (Session("Alert") <> "") Then
                                alert(Session("Alert"))
                            End If
                            sqlCommonAction.Transaction.Rollback()
                            alert("Error deleting question group in the template.")
                            failure = True
                        End If
                    Else
                        sqlCommonAction.Transaction.Commit()
                        If (Session("ReportGroupMax") = 1) Then
                            Session("ReportGroupRow") = 0
                            Session("intReportGroup") = -1
                            Session("ReportGroupMax") = 0
                            Me.dsCommon.Tables("QuestionGroups").Rows.Add(Me.dsCommon.Tables("QuestionGroups").NewRow())
                            Me.dsCommon.Tables("QuestionGroups").Rows(Session("ReportGroupRow")).Item("TEMPLATE_KEY") = Session("seqTemplate")
                            Me.dsCommon.Tables("QuestionGroups").Rows(Session("ReportGroupRow")).Item("QUESTION_GROUP_ID") = Session("ReportGroupMax")
                            Me.dsCommon.Tables("QuestionGroups").Rows(Session("ReportGroupRow")).Item("QUESTION_GROUP_TITLE") = ""
                        ElseIf (Session("intReportGroup") = Session("ReportGroupMax") And Session("ReportGroupMax") - Session("ReportGroupRow") = 1) Then
                            Session("intReportGroup") -= 1
                            Session("ReportGroupMax") -= 1
                            Session("ReportGroupRow") -= 1
                        Else
                            Session("ReportGroupMax") -= 1
                        End If
                    End If
                Else
                    sqlCommonAction.Transaction.Rollback()
                    alert("Error deleting question group in the template.")
                    failure = True
                End If
                Return failure
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Error deleting question group from template.")
                Return True
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Error deleting question group from template.")
            Return True
        End Try
    End Function

    'attempts to delete a Report Group, returns false if it cannot
    Private Function deleteQuestionGroup(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'attempt to delete the record from the database
            sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION_GROUP where TEMPLATE_KEY = " & Session("seqTemplate") & _
            " and QUESTION_GROUP_ID = " & Session("intReportGroup")
            sqlCommonAction.ExecuteNonQuery()

            'update the question id's
            Dim lowIndex As Integer = Session("ReportGroupRow")
            Dim highIndex As Integer = Me.dsCommon.Tables("QuestionGroups").Rows.Count()

            'update the ID's of the Report Groups
            While (lowIndex < highIndex)
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION_GROUP SET QUESTION_GROUP_ID = (QUESTION_GROUP_ID-1) " & _
                "WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and QUESTION_GROUP_ID = " & _
                Me.dsCommon.Tables("QuestionGroups").Rows(lowIndex).Item("QUESTION_GROUP_ID")
                sqlCommonAction.ExecuteNonQuery()
                lowIndex += 1
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
        If Not (loadData()) Then
            alert("Failed to load the question groups for this template.")
        Else
            Try
                If (Session("intReportGroup") = -1) Then
                    Me.dsCommon.Tables("QuestionGroups").Rows.Add(Me.dsCommon.Tables("QuestionGroups").NewRow())
                    Me.dsCommon.Tables("QuestionGroups").Rows(Session("ReportGroupMax")).Item("TEMPLATE_KEY") = Session("seqTemplate")
                    Me.dsCommon.Tables("QuestionGroups").Rows(Session("ReportGroupMax")).Item("QUESTION_GROUP_ID") = Session("ReportGroupMax")
                    Me.dsCommon.Tables("QuestionGroups").Rows(Session("ReportGroupMax")).Item("QUESTION_GROUP_TITLE") = ""
                End If
                Return False
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error resetting the question group.")
                Return True
            End Try
        End If
    End Function
#End Region
End Class
