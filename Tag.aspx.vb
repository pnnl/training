Imports System.Math
Imports System.Collections.Specialized
Public Class Tag
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Tag))
        Me.sqlComments = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlDeleteCommand1 = New System.Data.SqlClient.SqlCommand
        Me.commonConn = New System.Data.SqlClient.SqlConnection
        Me.SqlInsertCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlUpdateCommand1 = New System.Data.SqlClient.SqlCommand
        Me.dsCommon = New System.Data.DataSet
        Me.sqlTags = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlDeleteCommand2 = New System.Data.SqlClient.SqlCommand
        Me.SqlInsertCommand2 = New System.Data.SqlClient.SqlCommand
        Me.SqlSelectCommand2 = New System.Data.SqlClient.SqlCommand
        Me.SqlUpdateCommand2 = New System.Data.SqlClient.SqlCommand
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'sqlComments
        '
        Me.sqlComments.DeleteCommand = Me.SqlDeleteCommand1
        Me.sqlComments.InsertCommand = Me.SqlInsertCommand1
        Me.sqlComments.SelectCommand = Me.SqlSelectCommand1
        Me.sqlComments.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_TEMPLATE_TAG", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("TEMPLATE_KEY", "TEMPLATE_KEY"), New System.Data.Common.DataColumnMapping("TEMPLATE_TAG_ID", "TEMPLATE_TAG_ID"), New System.Data.Common.DataColumnMapping("TEMPLATE_TAG_TITLE", "TEMPLATE_TAG_TITLE"), New System.Data.Common.DataColumnMapping("TEMPLATE_TAG_QUESTION", "TEMPLATE_TAG_QUESTION")})})
        Me.sqlComments.UpdateCommand = Me.SqlUpdateCommand1
        '
        'SqlDeleteCommand1
        '
        Me.SqlDeleteCommand1.CommandText = "DELETE FROM VARCONN.[SAT_TEMPLATE_TAG] WHERE (([TEMPLATE_KEY] = @Origina" & _
            "l_TEMPLATE_KEY) AND ([TEMPLATE_TAG_ID] = @Original_TEMPLATE_TAG_ID))"
        Me.SqlDeleteCommand1.Connection = Me.commonConn
        Me.SqlDeleteCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_TAG_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_TAG_ID", System.Data.DataRowVersion.Original, Nothing)})
        '
        'commonConn
        '
        Me.commonConn.FireInfoMessageEventOnUserErrors = False
        '
        'SqlInsertCommand1
        '
        Me.SqlInsertCommand1.CommandText = resources.GetString("SqlInsertCommand1.CommandText")
        Me.SqlInsertCommand1.Connection = Me.commonConn
        Me.SqlInsertCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, "TEMPLATE_KEY"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_TAG_ID", System.Data.SqlDbType.Int, 0, "TEMPLATE_TAG_ID"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_TAG_TITLE", System.Data.SqlDbType.VarChar, 0, "TEMPLATE_TAG_TITLE"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_TAG_QUESTION", System.Data.SqlDbType.VarChar, 0, "TEMPLATE_TAG_QUESTION")})
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = "SELECT TEMPLATE_KEY, TEMPLATE_TAG_ID, TEMPLATE_TAG_TITLE, TEMPLATE_TAG_QUESTION F" & _
            "ROM VARCONN.SAT_TEMPLATE_TAG"
        Me.SqlSelectCommand1.Connection = Me.commonConn
        '
        'SqlUpdateCommand1
        '
        Me.SqlUpdateCommand1.CommandText = resources.GetString("SqlUpdateCommand1.CommandText")
        Me.SqlUpdateCommand1.Connection = Me.commonConn
        Me.SqlUpdateCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, "TEMPLATE_KEY"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_TAG_ID", System.Data.SqlDbType.Int, 0, "TEMPLATE_TAG_ID"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_TAG_TITLE", System.Data.SqlDbType.VarChar, 0, "TEMPLATE_TAG_TITLE"), New System.Data.SqlClient.SqlParameter("@TEMPLATE_TAG_QUESTION", System.Data.SqlDbType.VarChar, 0, "TEMPLATE_TAG_QUESTION"), New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_TAG_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_TAG_ID", System.Data.DataRowVersion.Original, Nothing)})
        '
        'dsCommon
        '
        Me.dsCommon.DataSetName = "NewDataSet"
        Me.dsCommon.Locale = New System.Globalization.CultureInfo("en-US")
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
        Me.SqlDeleteCommand2.CommandText = "DELETE FROM VARCONN.[SAT_SURVEY_TAG] WHERE (([SURVEY_KEY] = @Original_SURVEY_KEY)" & _
            " AND ([TEMPLATE_KEY] = @Original_TEMPLATE_KEY) AND ([SURVEY_TAG_ID] = @Original_" & _
            "SURVEY_TAG_ID))"
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
    Protected WithEvents Image1 As System.Web.UI.WebControls.Image
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents hlpServerSwitch As System.Web.UI.WebControls.Image
    Protected WithEvents chkServerSwitch As System.Web.UI.WebControls.CheckBox
    Protected WithEvents hlpTagName As System.Web.UI.WebControls.Image
    Protected WithEvents lblTagName As System.Web.UI.WebControls.Label
    Protected WithEvents hlpTagQuestion As System.Web.UI.WebControls.Image
    Protected WithEvents lblTagQuestion As System.Web.UI.WebControls.Label
    Protected WithEvents txtTagQuestion As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtTagName As System.Web.UI.WebControls.TextBox
    Protected WithEvents lblTemplate As System.Web.UI.WebControls.Label
    Protected WithEvents hlpPreview As System.Web.UI.WebControls.Image
    Protected WithEvents btnTemplatePreview As System.Web.UI.WebControls.Button
    Protected WithEvents sqlComments As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlInsertCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlUpdateCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlDeleteCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents dsCommon As System.Data.DataSet
    Protected WithEvents hlpTemplatePreview As System.Web.UI.WebControls.Image
    Protected requestItems As Array
    Protected WithEvents sqlTags As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents SqlSelectCommand2 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlInsertCommand2 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlUpdateCommand2 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlDeleteCommand2 As System.Data.SqlClient.SqlCommand
    Dim myNavBar As Navigation = New Navigation
    Dim myUtility As Utility = New Utility
    Protected WithEvents lblNavBar As System.Web.UI.WebControls.Label
    Protected WithEvents btnStart As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnPrev As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnReload As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnNext As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnLast As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnNew As System.Web.UI.WebControls.ImageButton
    Protected navButtons As Collection
    Protected WithEvents btnTop As System.Web.UI.WebControls.Button
    Protected WithEvents btnInsert As System.Web.UI.WebControls.Button
    Protected WithEvents btnUpdate As System.Web.UI.WebControls.Button
    Protected WithEvents btnDelete As System.Web.UI.WebControls.Button
    Protected WithEvents btnReset As System.Web.UI.WebControls.Button
    Protected sqlCommonAdapter As New System.Data.SqlClient.SqlDataAdapter
    Protected sqlCommonAction As New System.Data.SqlClient.SqlCommand
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

        'Set the connection based on the machine
        Me.commonConn = myUtility.getConnection(Me.commonConn)

        'determine template ownership if not defined
        If (Session("isTemplateOwner") Is Nothing) Then
            sqlCommonAction.Connection = Me.commonConn
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

        'finalize the SQL statements
        completeSQL()

        'Set the server switch text on initial page load
        If Not (IsPostBack) Then
            'get anything on the request line
            myUtility.getRequest(Page.Request.Url.Query())

            'load the comments
            If Not (loadComments()) Then
                alert("Failed to load the tag items for this template.")
            Else
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
                myUtility.optionPreSet(Session("intTagID"), Session("TagMax"), Me.pageOptions)

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
        Me.SqlSelectCommand1.CommandText &= " WHERE TEMPLATE_KEY = " & Session("seqTemplate")
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
    'loads the comment tags
    Private Function loadComments() As Boolean
        Trace.Warn("Loading Comments")
        Try
            'dump the data that may exist before refilling
            If (Me.dsCommon.Tables.Contains("Comments")) Then
                Me.dsCommon.Tables("Comments").Rows.Clear()
            End If

            Me.sqlComments.SelectCommand.Connection = Me.commonConn
            Me.sqlComments.SelectCommand.CommandText = Me.sqlComments.SelectCommand.CommandText.Replace("VARCONN.", myUtility.getExtension())
            Me.sqlComments.Fill(Me.dsCommon, "Comments")
            Session("TagMax") = Me.dsCommon.Tables("Comments").Rows.Count()

            'determine if there is any data and if the comment id has been set
            If (Session("TagMax") = 0) Then
                Session("intTagID") = -1
            ElseIf (Session("intTagID") Is Nothing) Then
                Session("intTagID") = 1
            End If

            'make sure the comment row never exceeds the maximum comments
            If (Session("TagRow") > Session("TagMax")) Then
                Session("TagRow") = Session("TagMax") - 1
            End If

            'if the message id indicates a new record then move the comment row to the maximum
            'else set it to just below the comment id's number
            If (Session("intTagID") = -1) Then
                Session("TagRow") = Session("TagMax")
            Else
                Session("TagRow") = Session("intTagID") - 1
            End If
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'loads the tag items for the current template
    Private Function loadTags(Optional ByVal override As String = "") As Boolean
        Trace.Warn("Loading List Items")
        Try
            Dim oldCommand As String = Me.SqlSelectCommand2.CommandText
            If (override <> "SurveyOrder") Then
                Me.SqlSelectCommand2.CommandText &= " WHERE TEMPLATE_KEY=" & Session("seqTemplate") & " ORDER BY SURVEY_TAG_ID DESC"
            Else
                Me.SqlSelectCommand2.CommandText &= " WHERE TEMPLATE_KEY=" & Session("seqTemplate") & " ORDER BY SURVEY_KEY, SURVEY_TAG_ID"
            End If
            Trace.Warn(Me.SqlSelectCommand2.CommandText)
            If (Me.dsCommon.Tables.Contains("Tags")) Then
                Me.dsCommon.Tables("Tags").Rows.Clear()
            End If
            Me.sqlTags.SelectCommand.Connection = Me.commonConn
            Me.sqlTags.SelectCommand.CommandText = Me.sqlTags.SelectCommand.CommandText.Replace("VARCONN.", myUtility.getExtension())
            Me.sqlTags.Fill(Me.dsCommon, "Tags")
            Me.SqlSelectCommand2.CommandText = oldCommand
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Unable to load the tag items for this template.")
            Return False
        End Try
    End Function
#End Region

#Region "Control Settings"
    'populates the controls with the selected comment item
    Private Sub setControls()
        Trace.Warn("Setting Controls")
        If (Session("intTagID") > 0) Then
            Try
                Me.txtTagName.Text = Me.dsCommon.Tables("Comments").Rows(Session("TagRow")).Item("TEMPLATE_TAG_TITLE")
                Me.txtTagQuestion.Text = Trim("" & Me.dsCommon.Tables("Comments").Rows(Session("TagRow")).Item("TEMPLATE_TAG_QUESTION"))
                Me.lblNavBar.Text = "Tag Item " & Session("TagRow") + 1 & " of " & Session("TagMax")
            Catch ex As Exception
                Trace.Warn(ex.ToString)
            End Try
        Else
            Me.txtTagName.Text = ""
            Me.txtTagQuestion.Text = ""
            Me.lblNavBar.Text = "Tag Item " & Session("TagRow") + 1 & " of " & Session("TagMax") + 1
        End If

        'set the page options
        myUtility.optionPreSet(Session("intTagID"), Session("TagMax"), Me.pageOptions)

        Session("PageOptions") = Me.pageOptions
    End Sub
#End Region

#Region "Page Switch"
    'Takes the user back to the template they were working on
    Private Sub btnTop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTop.Click
        Session("JSAction") = ""
        Session("CurrentPage") = "./Template.aspx"
        Session("pageSel") = "Template"
        redirect("./Template.aspx")
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
        myNavBar.wnb_MoveTo(Me.navButtons, Session("TagMax"), Session("TagRow"), Switch)
    End Sub
#End Region

#Region "Nav Bar Events"
    'Moves the user to the first record
    Private Sub wnbComments_MoveToStart(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnStart.Click
        Session("JSAction") = ""

        If Not (loadComments()) Then
            alert("Failed to load the tag items for this template.")
        Else
            Try
                Session("TagRow") = 0
                Session("intTagID") = Me.dsCommon.Tables("Comments").Rows(Session("TagRow")).Item("TEMPLATE_TAG_ID")
                setNavBar()
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("intTagID"), Session("TagMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the first tag item.")
            End Try
        End If
    End Sub

    'Moves the user to the previous record
    Private Sub wnbComments_MovePrev(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnPrev.Click
        Session("JSAction") = ""

        If Not (loadComments()) Then
            alert("Failed to load the tag items for this template.")
        Else
            Try
                Session("TagRow") -= 1
                Session("intTagID") = Me.dsCommon.Tables("Comments").Rows(Session("TagRow")).Item("TEMPLATE_TAG_ID")
                setNavBar()
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("intTagID"), Session("TagMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the previous tag item.")
            End Try
        End If
    End Sub

    'reloads the page and the data from the database while attempting to locate the user's location
    Private Sub wnbComments_MoveReload(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnReload.Click
        Session("JSAction") = ""

        If Not (loadComments()) Then
            alert("Failed to load the tag items for this template.")
        Else
            Try
                setNavBar()
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("intTagID"), Session("TagMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error resetting the current tag item.")
            End Try
        End If
    End Sub

    'Moves the user to the next record
    Private Sub wnbComments_MoveNext(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnNext.Click
        Session("JSAction") = ""

        If Not (loadComments()) Then
            alert("Failed to load the tag items for this template.")
        Else
            Try
                Session("TagRow") += 1
                Session("intTagID") = Me.dsCommon.Tables("Comments").Rows(Session("TagRow")).Item("TEMPLATE_TAG_ID")
                setNavBar()
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("intTagID"), Session("TagMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the previous tag item.")
            End Try
        End If
    End Sub

    'Moves the user to the last record
    Private Sub wnbComments_MoveToEnd(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnLast.Click
        Session("JSAction") = ""

        If Not (loadComments()) Then
            alert("Failed to load the tag items for this template.")
        Else
            Try
                Session("TagRow") = Session("TagMax") - 1
                Session("intTagID") = Me.dsCommon.Tables("Comments").Rows(Session("TagRow")).Item("TEMPLATE_TAG_ID")
                setNavBar()
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("intTagID"), Session("TagMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the last tag item.")
            End Try
        End If
    End Sub

    'Inserts a record into the table
    Private Sub wnbComments_InsertRow(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnNew.Click
        Session("JSAction") = ""

        If Not (loadComments()) Then
            alert("Failed to load the tag items for this template.")
        Else
            Try
                Session("intTagID") = -1
                Session("TagRow") = Session("TagMax")
                Me.dsCommon.Tables("Comments").Rows.Add(Me.dsCommon.Tables("Comments").NewRow())
                Me.dsCommon.Tables("Comments").Rows(Session("TagMax")).Item("TEMPLATE_KEY") = Session("seqTemplate")
                Me.dsCommon.Tables("Comments").Rows(Session("TagMax")).Item("TEMPLATE_TAG_ID") = Session("TagMax")
                Me.dsCommon.Tables("Comments").Rows(Session("TagMax")).Item("TEMPLATE_TAG_TITLE") = ""
                Me.dsCommon.Tables("Comments").Rows(Session("TagMax")).Item("TEMPLATE_TAG_QUESTION") = ""
                setNavBar()
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("intTagID"), Session("TagMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error creating new tag item.")
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
        
        'refresh the comments table
        If (loadComments()) Then
            'refresh the tags table
            If Not (loadTags("SurveyOrder")) Then
                blnContinue = False
                alert("Failed to load the tag items for the current template. Insert aborted.")
            End If
        Else
            blnContinue = False
            alert("Failed to load the tag items for the current template. Insert aborted.")
        End If

        'check for possible malicious code input
        If Not (myUtility.goodInput(Me.txtTagName, True) And myUtility.goodInput(Me.txtTagQuestion, True)) Then
            alert("Possible malicious code entry(s).  Insert aborted.")
            blnContinue = False
        End If

        If (blnContinue) Then
            'determine if the template is dirty or not
            Session("isDirty") = myUtility.isDirty(Me.dsCommon, Me.commonConn)

            'do the insert
            failure = doInsert()

            'reload the data
            If Not (loadComments()) Then
                blnContinue = False
                alert("Failed to load the tag items for the current template.")
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
                myUtility.optionPreSet(Session("intTagID"), Session("TagMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection
    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True
        
        'refresh the comments table
        If (loadComments()) Then
            'refresh the tags table
            If Not (loadTags("SurveyOrder")) Then
                blnContinue = False
                alert("Failed to load the tag items for the current template. Update aborted.")
            End If
        Else
            blnContinue = False
            alert("Failed to load the tag items for the current template. Update aborted.")
        End If

        'check for possible malicious code input
        If Not (myUtility.goodInput(Me.txtTagName, True) And myUtility.goodInput(Me.txtTagQuestion, True)) Then
            alert("Possible malicious code entry(s).  Update aborted.")
            blnContinue = False
        End If

        If (blnContinue) Then
            'determine if the template is dirty or not
            Session("isDirty") = myUtility.isDirty(Me.dsCommon, Me.commonConn)

            'do the update
            failure = doUpdate()

            'reload the data
            If Not (loadComments()) Then
                blnContinue = False
                alert("Failed to load the tag items for the current template.")
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
                myUtility.optionPreSet(Session("intTagID"), Session("TagMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True
        
        'refresh the comments table
        If (loadComments()) Then
            'refresh the tags table
            If Not (loadTags("SurveyOrder")) Then
                blnContinue = False
                alert("Failed to load the tag items for the current template. Delete aborted.")
            End If
        Else
            blnContinue = False
            alert("Failed to load the tag items for the current template. Delete aborted.")
        End If

        'check for possible malicious code input
        If Not (myUtility.goodInput(Me.txtTagName, True) And myUtility.goodInput(Me.txtTagQuestion, True)) Then
            alert("Possible malicious code entry(s).  Delete aborted.")
            blnContinue = False
        End If

        If (blnContinue) Then
            'determine if the template is dirty or not
            Session("isDirty") = myUtility.isDirty(Me.dsCommon, Me.commonConn)

            'do the delete
            failure = doDelete()

            'reload the data
            If Not (loadComments()) Then
                blnContinue = False
                alert("Failed to load the tag items for the current template.")
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
                myUtility.optionPreSet(Session("intTagID"), Session("TagMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'check for possible malicious code input
        If Not (myUtility.goodInput(Me.txtTagName, True) And myUtility.goodInput(Me.txtTagQuestion, True)) Then
            alert("Possible malicious code entry(s).  Reset aborted.")
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
                myUtility.optionPreSet(Session("intTagID"), Session("TagMax"), Me.pageOptions)

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
                If (Me.txtTagName.Text <> "" And Me.txtTagQuestion.Text <> "") Then
                    If (Session("intTagID") = -1) Then
                        'If we're inserting a new record then add it onto the end of the recordset and add it to the existing surveys
                        If (insertComment(sqlCommonAction, sqlCommonAdapter, True)) Then
                            If (insertTags(sqlCommonAction, sqlCommonAdapter, True)) Then
                                If Not (Session("isDirty")) Then
                                    Dim tempConn As Data.SqlClient.SqlConnection = New Data.SqlClient.SqlConnection
                                    If (myUtility.makeDirty(sqlCommonAction, sqlCommonAdapter, Me.dsCommon, myUtility.getConnection(tempConn))) Then
                                        If (Session("Alert") <> "") Then
                                            alert(Session("Alert"))
                                        End If
                                        Session("TagMax") += 1
                                        Session("intTagID") = Session("TagMax")
                                        sqlCommonAction.Transaction.Commit()
                                        Session("isDirty") = True
                                    Else
                                        If (Session("Alert") <> "") Then
                                            alert(Session("Alert"))
                                        End If
                                        sqlCommonAction.Transaction.Rollback()
                                        alert("Error adding tag item to the template.")
                                        failure = True
                                    End If
                                Else
                                    Session("TagMax") += 1
                                    Session("intTagID") = Session("TagMax")
                                    sqlCommonAction.Transaction.Commit()
                                End If
                            Else
                                sqlCommonAction.Transaction.Rollback()
                                alert("Error adding tag item to the template.")
                                failure = True
                            End If
                        Else
                            sqlCommonAction.Transaction.Rollback()
                            alert("Error adding tag item to the template.")
                            failure = True
                        End If
                    Else
                        'If we're inserting an existing record then inline insert it ito the record set and add it to the existing surveys
                        If (insertComment(sqlCommonAction, sqlCommonAdapter)) Then
                            If (insertTags(sqlCommonAction, sqlCommonAdapter)) Then
                                If Not (Session("isDirty")) Then
                                    Dim tempConn As Data.SqlClient.SqlConnection = New Data.SqlClient.SqlConnection
                                    If (myUtility.makeDirty(sqlCommonAction, sqlCommonAdapter, Me.dsCommon, myUtility.getConnection(tempConn))) Then
                                        If (Session("Alert") <> "") Then
                                            alert(Session("Alert"))
                                        End If
                                        Session("TagMax") += 1
                                        sqlCommonAction.Transaction.Commit()
                                        Session("isDirty") = True
                                    Else
                                        If (Session("Alert") <> "") Then
                                            alert(Session("Alert"))
                                        End If
                                        sqlCommonAction.Transaction.Rollback()
                                        alert("Error adding tag item to the template.")
                                        failure = True
                                    End If
                                Else
                                    Session("TagMax") += 1
                                    sqlCommonAction.Transaction.Commit()
                                End If
                            Else
                                sqlCommonAction.Transaction.Rollback()
                                alert("Error adding tag item to the template.")
                                failure = True
                            End If
                        Else
                            sqlCommonAction.Transaction.Rollback()
                            alert("Error adding tag item to the template.")
                            failure = True
                        End If
                    End If
                Else
                    alert("You must have some text in the Tag Name and Tag Question fields.")
                    failure = True
                    sqlCommonAction.Transaction.Rollback()
                End If
                Return failure
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Tag insertion failure.")
                Return True
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Tag insertion failure.")
            Return True
        End Try
    End Function

    'attempts to insert a comment, returns false if it cannot
    Private Function insertComment(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, Optional ByVal isNewRecord As Boolean = False) As Boolean
        'parameterized the text input to allow for things like quotes to get through
        Try
            Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@TEMPLATE_TAG_TITLE", System.Data.SqlDbType.VarChar, 150, "TEMPLATE_TAG_TITLE")
            param0.Value = Me.txtTagName.Text
            sqlCommonAction.Parameters.Add(param0)
            Dim param1 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@TEMPLATE_TAG_QUESTION", System.Data.SqlDbType.VarChar, 1800, "TEMPLATE_TAG_QUESTION")
            param1.Value = Me.txtTagQuestion.Text
            sqlCommonAction.Parameters.Add(param1)

            'process the insert based on whether or not we are inserting a new record
            If (isNewRecord) Then
                'attempt to add the new record to the database
                sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_TEMPLATE_TAG (TEMPLATE_KEY, TEMPLATE_TAG_ID, TEMPLATE_TAG_TITLE, " & _
                "TEMPLATE_TAG_QUESTION) VALUES (" & Session("seqTemplate") & ", " & Session("TagMax") + 1 & _
                ", @TEMPLATE_TAG_TITLE, @TEMPLATE_TAG_QUESTION)"
                sqlCommonAction.ExecuteNonQuery()
                Return True
            Else
                'update the tag id's
                Dim lowIndex As Integer = Session("TagRow")
                Dim highIndex As Integer = Session("TagMax") - 1

                'update the tagID's of the comments
                While (highIndex >= lowIndex)
                    sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_TAG SET TEMPLATE_TAG_ID = (TEMPLATE_TAG_ID+1) " & _
                    "WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and TEMPLATE_TAG_ID = " & _
                    Me.dsCommon.Tables("Comments").Rows(highIndex).Item("TEMPLATE_TAG_ID")
                    sqlCommonAction.ExecuteNonQuery()
                    highIndex -= 1
                End While

                'attempt to add the new record to the database
                sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_TEMPLATE_TAG (TEMPLATE_KEY, TEMPLATE_TAG_ID, TEMPLATE_TAG_TITLE, " & _
                    "TEMPLATE_TAG_QUESTION) VALUES (" & Session("seqTemplate") & ", " & Session("TagRow") + 1 & _
                ", @TEMPLATE_TAG_TITLE, @TEMPLATE_TAG_QUESTION)"
                sqlCommonAction.ExecuteNonQuery()
                Return True
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'attempts to insert the comments into the survey tags tables, returns false if it cannot
    Private Function insertTags(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, Optional ByVal isNewRecord As Boolean = False) As Boolean
        'if dealing with a new record (non-inline insert) get the existing surveys for this template
        Try
            If (isNewRecord) Then
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY where TEMPLATE_KEY = " & Session("seqTemplate")
                sqlCommonAdapter.SelectCommand = sqlCommonAction
                sqlCommonAdapter.Fill(Me.dsCommon, "Surveys")
            End If

            'parameterized the text input to allow for things like quotes to get through
            Dim param2 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@SURVEY_TAG_TITLE", System.Data.SqlDbType.VarChar, 150, "SURVEY_TAG_TITLE")
            param2.Value = Me.txtTagName.Text
            sqlCommonAction.Parameters.Add(param2)
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try

        'process the records based on whether or not we are inserting a new record
        If (isNewRecord) Then
            Try
                If (Me.dsCommon.Tables("Surveys").Rows.Count > 0) Then
                    Dim row As DataRow
                    For Each row In Me.dsCommon.Tables("Surveys").Rows()
                        'add the comment to each survey
                        Dim currentSurvey As Integer = -1
                        Dim oldTagMax As Integer = Me.dsCommon.Tables("Tags").Rows.Count()
                        Dim tagIndex As Integer = 0

                        While (tagIndex <= oldTagMax)
                            If (currentSurvey = -1) Then
                                currentSurvey = row.Item("seqSurvey")
                                sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_SURVEY_TAG (TEMPLATE_KEY, SURVEY_KEY, SURVEY_TAG_ID, SURVEY_TAG_TITLE, " & _
                                    "SURVEY_TAG_ANSWER) VALUES (" & Session("seqTemplate") & ", " & row.Item("seqSurvey") & _
                                    ", " & Session("TagRow") + 1 & ", @SURVEY_TAG_TITLE, '')"
                                sqlCommonAction.ExecuteNonQuery()
                            Else
                                If (currentSurvey <> row.Item("seqSurvey")) Then
                                    currentSurvey = row.Item("seqSurvey")
                                    sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_SURVEY_TAG (TEMPLATE_KEY, SURVEY_KEY, SUVEY_TAG_ID, SURVEY_TAG_TITLE, " & _
                                    "SURVEY_TAG_ANSWER) VALUES (" & Session("seqTemplate") & ", " & row.Item("seqSurvey") & _
                                    ", " & Session("TagRow") + 1 & ", @SURVEY_TAG_TITLE, '')"
                                    sqlCommonAction.ExecuteNonQuery()
                                End If
                            End If
                            tagIndex += 1
                        End While
                    Next
                End If
                Return True
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                Return False
            End Try
        Else
            'add the comment to each survey after updating the tag IDs for each survey
            Dim currentSurvey As Integer = -1
            Dim tagIndex As Integer = Me.dsCommon.Tables("Tags").Rows.Count() - 1

            'if there are no surveys this will fail and we return true
            Try
                currentSurvey = Me.dsCommon.Tables("Tags").Rows(tagIndex).Item("SURVEY_KEY")
            Catch ex As Exception
                Return True
            End Try

            'process the surveys
            Try
                While (tagIndex >= 0)
                    If (Me.dsCommon.Tables("Tags").Rows(tagIndex).Item("SURVEY_TAG_ID") > (Session("TagRow") + 1)) Then
                        sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_SURVEY_TAG SET SURVEY_TAG_ID = (SURVEY_TAG_ID+1) " & _
                        "WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and SURVEY_KEY = " & _
                        Me.dsCommon.Tables("Tags").Rows(tagIndex).Item("SURVEY_KEY") & _
                        " and SURVEY_TAG_ID = " & Me.dsCommon.Tables("Tags").Rows(tagIndex).Item("SURVEY_TAG_ID")
                        sqlCommonAction.ExecuteNonQuery()
                    ElseIf (Me.dsCommon.Tables("Tags").Rows(tagIndex).Item("SURVEY_TAG_ID") < (Session("TagRow") + 1)) Then
                        'do nothing don't want to change this data
                    Else
                        'update the row currently in the way
                        sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_SURVEY_TAG SET SURVEY_TAG_ID = (SURVEY_TAG_ID+1) " & _
                        "WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and SURVEY_KEY = " & _
                        Me.dsCommon.Tables("Tags").Rows(tagIndex).Item("SURVEY_KEY") & _
                        " and SURVEY_TAG_ID = " & Me.dsCommon.Tables("Tags").Rows(tagIndex).Item("SURVEY_TAG_ID")
                        sqlCommonAction.ExecuteNonQuery()

                        'add in the new row in-line where the row we just bumped existed
                        sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_SURVEY_TAG (TEMPLATE_KEY, SURVEY_KEY, SURVEY_TAG_ID, SURVEY_TAG_TITLE, " & _
                        "SURVEY_TAG_ANSWER) VALUES (" & Session("seqTemplate") & ", " & _
                        Me.dsCommon.Tables("Tags").Rows(tagIndex).Item("SURVEY_KEY") & _
                        ", " & Session("TagRow") + 1 & ", @SURVEY_TAG_TITLE, '')"
                        sqlCommonAction.ExecuteNonQuery()
                    End If
                    tagIndex -= 1
                End While
                Return True
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                Return False
            End Try
        End If
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
                If (Me.txtTagName.Text <> "" And Me.txtTagQuestion.Text <> "") Then
                    If (updateComment(sqlCommonAction, sqlCommonAdapter)) Then
                        If (updateTags(sqlCommonAction, sqlCommonAdapter)) Then
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
                                    alert("Error updating tag item in the template.")
                                    failure = True
                                End If
                            Else
                                sqlCommonAction.Transaction.Commit()
                            End If
                        Else
                            sqlCommonAction.Transaction.Rollback()
                            alert("Error updating tag item in the template.")
                            failure = True
                        End If
                    Else
                        sqlCommonAction.Transaction.Rollback()
                        alert("Error updating tag item in the template.")
                        failure = True
                    End If
                Else
                    sqlCommonAction.Transaction.Rollback()
                    alert("You must have some text in the Tag Name and Tag Question fields.")
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

    'attempts to update a comment, returns false if it cannot
    Private Function updateComment(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        'attempt to update the record in the database
        Try
            sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_TAG SET TEMPLATE_TAG_TITLE = @TEMPLATE_TAG_TITLE" & _
            ", TEMPLATE_TAG_QUESTION = @TEMPLATE_TAG_QUESTION" & _
            " WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and TEMPLATE_TAG_ID = " & Session("intTagID")

            'parameterized the text input to allow for things like quotes to get through
            Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@TEMPLATE_TAG_TITLE", System.Data.SqlDbType.VarChar, 150, "TEMPLATE_TAG_TITLE")
            param0.Value = Me.txtTagName.Text
            sqlCommonAction.Parameters.Add(param0)
            Dim param1 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@TEMPLATE_TAG_QUESTION", System.Data.SqlDbType.VarChar, 1800, "TEMPLATE_TAG_QUESTION")
            param1.Value = Me.txtTagQuestion.Text
            sqlCommonAction.Parameters.Add(param1)

            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'attempts to update the comments in the survey tags tables, returns false if it cannot
    Private Function updateTags(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        'parameterized the text input to allow for things like quotes to get through
        Try
            Dim param2 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@SURVEY_TAG_TITLE", System.Data.SqlDbType.VarChar, 150, "SURVEY_TAG_TITLE")
            param2.Value = Me.txtTagName.Text
            sqlCommonAction.Parameters.Add(param2)

            'update the comment in each survey
            Dim currentSurvey As Integer = -1
            Dim surveyMax As Integer = Me.dsCommon.Tables("Tags").Rows.Count()
            Dim surveyIndex As Integer = 0
            While (surveyIndex < surveyMax)
                If (Session("intTagID") = Me.dsCommon.Tables("Tags").Rows(surveyIndex).Item("SURVEY_TAG_ID")) Then
                    sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_SURVEY_TAG SET SURVEY_TAG_TITLE = @SURVEY_TAG_TITLE" & _
                    " WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and SURVEY_KEY = " & _
                    Me.dsCommon.Tables("Tags").Rows(surveyIndex).Item("seqSurvey") & _
                    " and SURVEY_TAG_ID = " & Me.dsCommon.Tables("Tags").Rows(surveyIndex).Item("SURVEY_TAG_ID")
                    sqlCommonAction.ExecuteNonQuery()
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

#Region "Delete Code"
    'deletes the current item from the comment and tag tables
    Private Function doDelete() As Boolean
        Try
            'start transaction
            If (myUtility.setupTransaction(sqlCommonAction, Me.commonConn)) Then
                Dim failure As Boolean = False
                'do the deletion
                If (deleteComment(sqlCommonAction, sqlCommonAdapter)) Then
                    If (deleteTags(sqlCommonAction, sqlCommonAdapter)) Then
                        If Not (Session("isDirty")) Then
                            Dim tempConn As Data.SqlClient.SqlConnection = New Data.SqlClient.SqlConnection
                            If (myUtility.makeDirty(sqlCommonAction, sqlCommonAdapter, Me.dsCommon, myUtility.getConnection(tempConn))) Then
                                If (Session("Alert") <> "") Then
                                    alert(Session("Alert"))
                                End If
                                Session("isDirty") = True
                                sqlCommonAction.Transaction.Commit()
                                'determine if we need to take a step back 
                                If (Session("TagMax") = 1) Then
                                    Session("TagRow") = 0
                                    Session("intTagID") = -1
                                    Session("TagMax") = 0
                                    Me.dsCommon.Tables("Comments").Rows.Add(Me.dsCommon.Tables("Comments").NewRow())
                                    Me.dsCommon.Tables("Comments").Rows(Session("TagMax")).Item("TEMPLATE_KEY") = Session("seqTemplate")
                                    Me.dsCommon.Tables("Comments").Rows(Session("TagMax")).Item("TEMPLATE_TAG_ID") = Session("TagMax")
                                    Me.dsCommon.Tables("Comments").Rows(Session("TagMax")).Item("TEMPLATE_TAG_TITLE") = ""
                                    Me.dsCommon.Tables("Comments").Rows(Session("TagMax")).Item("TEMPLATE_TAG_QUESTION") = ""
                                ElseIf (Session("intTagID") = Session("TagMax") And Session("TagMax") - Session("TagRow") = 1) Then
                                    Session("intTagID") -= 1
                                    Session("TagMax") -= 1
                                    Session("TagRow") -= 1
                                Else
                                    Session("TagMax") -= 1
                                End If
                            Else
                                If (Session("Alert") <> "") Then
                                    alert(Session("Alert"))
                                End If
                                sqlCommonAction.Transaction.Rollback()
                                alert("Error deleting tag item in the template.")
                                failure = True
                            End If
                        Else
                            sqlCommonAction.Transaction.Commit()
                            'determine if we need to take a step back 
                            If (Session("TagMax") = 1) Then
                                Session("TagRow") = 0
                                Session("intTagID") = -1
                                Session("TagMax") = 0
                                Me.dsCommon.Tables("Comments").Rows.Add(Me.dsCommon.Tables("Comments").NewRow())
                                Me.dsCommon.Tables("Comments").Rows(Session("TagMax")).Item("TEMPLATE_KEY") = Session("seqTemplate")
                                Me.dsCommon.Tables("Comments").Rows(Session("TagMax")).Item("TEMPLATE_TAG_ID") = Session("TagMax")
                                Me.dsCommon.Tables("Comments").Rows(Session("TagMax")).Item("TEMPLATE_TAG_TITLE") = ""
                                Me.dsCommon.Tables("Comments").Rows(Session("TagMax")).Item("TEMPLATE_TAG_QUESTION") = ""
                            ElseIf (Session("intTagID") = Session("TagMax") And Session("TagMax") - Session("TagRow") = 1) Then
                                Session("intTagID") -= 1
                                Session("TagMax") -= 1
                                Session("TagRow") -= 1
                            Else
                                Session("TagMax") -= 1
                            End If
                        End If
                    Else
                        sqlCommonAction.Transaction.Rollback()
                        alert("Error deleting tag item in the template.")
                        failure = True
                    End If
                Else
                    sqlCommonAction.Transaction.Rollback()
                    alert("Error deleting tag item in the template.")
                    failure = True
                End If
                Return failure
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Tag deletion failure.")
                Return True
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Tag deletion failure.")
            Return True
        End Try
    End Function

    'attempts to delete a comment, returns false if it cannot
    Private Function deleteComment(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        'attempt to delete the record from the database
        Try
            sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_TEMPLATE_TAG where TEMPLATE_KEY = " & Session("seqTemplate") & _
            " and TEMPLATE_TAG_ID = " & Session("intTagID")
            sqlCommonAction.ExecuteNonQuery()

            'update the tag id's
            Dim lowIndex As Integer = Session("TagRow")
            Dim highIndex As Integer = Me.dsCommon.Tables("Comments").Rows.Count()

            'update the tagID's of the comments
            While (lowIndex < highIndex)
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_TAG SET TEMPLATE_TAG_ID = (TEMPLATE_TAG_ID-1) " & _
                "WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and TEMPLATE_TAG_ID = " & _
                Me.dsCommon.Tables("Comments").Rows(lowIndex).Item("TEMPLATE_TAG_ID")
                sqlCommonAction.ExecuteNonQuery()
                lowIndex += 1
            End While
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'attempts to delete the comments in the survey tags tables, returns false if it cannot
    Private Function deleteTags(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        'delete the comment from each survey before updating the tag IDs for each survey
        Dim currentSurvey As Integer = -1
        Dim surveyIndex As Integer = 0

        'if there are no tags then this will fail and return true otherwise it will process any existing tags that
        'might need to be updated to reflect the deletion of the comment associated with them
        Try
            currentSurvey = Me.dsCommon.Tables("Tags").Rows(surveyIndex).Item("SURVEY_KEY")
        Catch ex As Exception
            Return True
        End Try

        'process the tags
        Try
            While (surveyIndex < Me.dsCommon.Tables("Tags").Rows.Count())
                If (Me.dsCommon.Tables("Tags").Rows(surveyIndex).Item("SURVEY_TAG_ID") = Session("intTagID")) Then
                    'attempt to delete the record from the database
                    sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_SURVEY_TAG where TEMPLATE_KEY = " & Session("seqTemplate") & _
                    " and SURVEY_KEY = " & Me.dsCommon.Tables("Tags").Rows(surveyIndex).Item("SURVEY_KEY") & _
                    " and SURVEY_TAG_ID = " & Session("intTagID")
                    sqlCommonAction.ExecuteNonQuery()

                    'only update the tag item after if one actually exists and it happens to be larger than the tag ID we're 
                    'getting rid of
                    If (surveyIndex < Me.dsCommon.Tables("Tags").Rows.Count()) Then
                        If (Me.dsCommon.Tables("Tags").Rows(surveyIndex).Item("SURVEY_TAG_ID") > Session("intTagID")) Then
                            sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_SURVEY_TAG SET SURVEY_TAG_ID = (SURVEY_TAG_ID-1) " & _
                            "WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and SURVEY_KEY = " & _
                            Me.dsCommon.Tables("Tags").Rows(surveyIndex).Item("SURVEY_KEY") & _
                            " and SURVEY_TAG_ID = " & Me.dsCommon.Tables("Tags").Rows(surveyIndex).Item("SURVEY_TAG_ID")
                            sqlCommonAction.ExecuteNonQuery()
                        End If
                    End If
                ElseIf (Me.dsCommon.Tables("Tags").Rows(surveyIndex).Item("SURVEY_TAG_ID") > Session("intTagID")) Then
                    sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_SURVEY_TAG SET SURVEY_TAG_ID = (SURVEY_TAG_ID-1) " & _
                    "WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and SURVEY_KEY = " & _
                    Me.dsCommon.Tables("Tags").Rows(surveyIndex).Item("SURVEY_KEY") & _
                    " and SURVEY_TAG_ID = " & Me.dsCommon.Tables("Tags").Rows(surveyIndex).Item("SURVEY_TAG_ID")
                    sqlCommonAction.ExecuteNonQuery()
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
        loadComments()
        If (Session("intTagID") = -1) Then
            Me.dsCommon.Tables("Comments").Rows.Add(Me.dsCommon.Tables("Comments").NewRow())
            Me.dsCommon.Tables("Comments").Rows(Session("TagMax")).Item("TEMPLATE_KEY") = Session("seqTemplate")
            Me.dsCommon.Tables("Comments").Rows(Session("TagMax")).Item("TEMPLATE_TAG_ID") = Session("TagMax")
            Me.dsCommon.Tables("Comments").Rows(Session("TagMax")).Item("TEMPLATE_TAG_TITLE") = ""
            Me.dsCommon.Tables("Comments").Rows(Session("TagMax")).Item("TEMPLATE_TAG_QUESTION") = ""
        End If
        Return False
    End Function
#End Region
End Class
