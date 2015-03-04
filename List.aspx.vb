Imports System.Math
Imports System.Text
Imports System.IO
Imports System.Collections.Specialized
Public Class List
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(List))
        Me.sqlQuestion = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand
        Me.sqlListItems = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlDeleteCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlInsertCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlSelectCommand2 = New System.Data.SqlClient.SqlCommand
        Me.SqlUpdateCommand1 = New System.Data.SqlClient.SqlCommand
        Me.dsCommon = New System.Data.DataSet
        Me.commonConn = New System.Data.SqlClient.SqlConnection
        Me.SqlInsertCommand = New System.Data.SqlClient.SqlCommand
        Me.SqlUpdateCommand = New System.Data.SqlClient.SqlCommand
        Me.SqlDeleteCommand = New System.Data.SqlClient.SqlCommand
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'sqlQuestion
        '
        Me.sqlQuestion.DeleteCommand = Me.SqlDeleteCommand
        Me.sqlQuestion.InsertCommand = Me.SqlInsertCommand
        Me.sqlQuestion.SelectCommand = Me.SqlSelectCommand1
        Me.sqlQuestion.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_TEMPLATE_QUESTION", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("TEMPLATE_KEY", "TEMPLATE_KEY"), New System.Data.Common.DataColumnMapping("QUESTION_ID", "QUESTION_ID"), New System.Data.Common.DataColumnMapping("REQUIRED_IND", "REQUIRED_IND"), New System.Data.Common.DataColumnMapping("QUESTION_TYPE_CODE", "QUESTION_TYPE_CODE"), New System.Data.Common.DataColumnMapping("QUESTION_TEXT", "QUESTION_TEXT"), New System.Data.Common.DataColumnMapping("RATING_COUNT", "RATING_COUNT"), New System.Data.Common.DataColumnMapping("RATING_STEP_VALUE", "RATING_STEP_VALUE"), New System.Data.Common.DataColumnMapping("RATING_INITIAL_VALUE", "RATING_INITIAL_VALUE"), New System.Data.Common.DataColumnMapping("BRANCH_QUESTION_ID", "BRANCH_QUESTION_ID")})})
        Me.sqlQuestion.UpdateCommand = Me.SqlUpdateCommand
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = "SELECT TEMPLATE_KEY, QUESTION_ID, REQUIRED_IND, QUESTION_TYPE_CODE, QUESTION_TEXT" & _
            ", RATING_COUNT, RATING_STEP_VALUE, RATING_INITIAL_VALUE, BRANCH_QUESTION_ID FROM" & _
            " VARCONN.SAT_TEMPLATE_QUESTION"
        Me.SqlSelectCommand1.Connection = Me.commonConn
        '
        'sqlListItems
        '
        Me.sqlListItems.DeleteCommand = Me.SqlDeleteCommand1
        Me.sqlListItems.InsertCommand = Me.SqlInsertCommand1
        Me.sqlListItems.SelectCommand = Me.SqlSelectCommand2
        Me.sqlListItems.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_QUESTION_LIST_ITEM", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("TEMPLATE_KEY", "TEMPLATE_KEY"), New System.Data.Common.DataColumnMapping("QUESTION_ID", "QUESTION_ID"), New System.Data.Common.DataColumnMapping("LIST_ITEM_ID", "LIST_ITEM_ID"), New System.Data.Common.DataColumnMapping("LIST_ITEM_VALUE", "LIST_ITEM_VALUE"), New System.Data.Common.DataColumnMapping("LIST_ITEM_TITLE", "LIST_ITEM_TITLE"), New System.Data.Common.DataColumnMapping("LIST_ITEM_IMAGE", "LIST_ITEM_IMAGE"), New System.Data.Common.DataColumnMapping("LIST_ITEM_IMAGE_TYPE", "LIST_ITEM_IMAGE_TYPE"), New System.Data.Common.DataColumnMapping("BRANCH_QUESTION_ID", "BRANCH_QUESTION_ID")})})
        Me.sqlListItems.UpdateCommand = Me.SqlUpdateCommand1
        '
        'SqlDeleteCommand1
        '
        Me.SqlDeleteCommand1.CommandText = "DELETE FROM VARCONN.[SAT_QUESTION_LIST_ITEM] WHERE (([TEMPLATE_KEY] = @O" & _
            "riginal_TEMPLATE_KEY) AND ([QUESTION_ID] = @Original_QUESTION_ID) AND ([LIST_ITE" & _
            "M_ID] = @Original_LIST_ITEM_ID))"
        Me.SqlDeleteCommand1.Connection = Me.commonConn
        Me.SqlDeleteCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_QUESTION_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "QUESTION_ID", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_LIST_ITEM_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LIST_ITEM_ID", System.Data.DataRowVersion.Original, Nothing)})
        '
        'SqlInsertCommand1
        '
        Me.SqlInsertCommand1.CommandText = resources.GetString("SqlInsertCommand1.CommandText")
        Me.SqlInsertCommand1.Connection = Me.commonConn
        Me.SqlInsertCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, "TEMPLATE_KEY"), New System.Data.SqlClient.SqlParameter("@QUESTION_ID", System.Data.SqlDbType.Int, 0, "QUESTION_ID"), New System.Data.SqlClient.SqlParameter("@LIST_ITEM_ID", System.Data.SqlDbType.Int, 0, "LIST_ITEM_ID"), New System.Data.SqlClient.SqlParameter("@LIST_ITEM_VALUE", System.Data.SqlDbType.Int, 0, "LIST_ITEM_VALUE"), New System.Data.SqlClient.SqlParameter("@LIST_ITEM_TITLE", System.Data.SqlDbType.VarChar, 0, "LIST_ITEM_TITLE"), New System.Data.SqlClient.SqlParameter("@LIST_ITEM_IMAGE", System.Data.SqlDbType.Image, 0, "LIST_ITEM_IMAGE"), New System.Data.SqlClient.SqlParameter("@LIST_ITEM_IMAGE_TYPE", System.Data.SqlDbType.VarChar, 0, "LIST_ITEM_IMAGE_TYPE"), New System.Data.SqlClient.SqlParameter("@BRANCH_QUESTION_ID", System.Data.SqlDbType.Int, 0, "BRANCH_QUESTION_ID")})
        '
        'SqlSelectCommand2
        '
        Me.SqlSelectCommand2.CommandText = "SELECT TEMPLATE_KEY, QUESTION_ID, LIST_ITEM_ID, LIST_ITEM_VALUE, LIST_ITEM_TITLE," & _
            " LIST_ITEM_IMAGE, LIST_ITEM_IMAGE_TYPE, BRANCH_QUESTION_ID FROM VARCONN.SAT" & _
            "_QUESTION_LIST_ITEM"
        Me.SqlSelectCommand2.Connection = Me.commonConn
        '
        'SqlUpdateCommand1
        '
        Me.SqlUpdateCommand1.CommandText = resources.GetString("SqlUpdateCommand1.CommandText")
        Me.SqlUpdateCommand1.Connection = Me.commonConn
        Me.SqlUpdateCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, "TEMPLATE_KEY"), New System.Data.SqlClient.SqlParameter("@QUESTION_ID", System.Data.SqlDbType.Int, 0, "QUESTION_ID"), New System.Data.SqlClient.SqlParameter("@LIST_ITEM_ID", System.Data.SqlDbType.Int, 0, "LIST_ITEM_ID"), New System.Data.SqlClient.SqlParameter("@LIST_ITEM_VALUE", System.Data.SqlDbType.Int, 0, "LIST_ITEM_VALUE"), New System.Data.SqlClient.SqlParameter("@LIST_ITEM_TITLE", System.Data.SqlDbType.VarChar, 0, "LIST_ITEM_TITLE"), New System.Data.SqlClient.SqlParameter("@LIST_ITEM_IMAGE", System.Data.SqlDbType.Image, 0, "LIST_ITEM_IMAGE"), New System.Data.SqlClient.SqlParameter("@LIST_ITEM_IMAGE_TYPE", System.Data.SqlDbType.VarChar, 0, "LIST_ITEM_IMAGE_TYPE"), New System.Data.SqlClient.SqlParameter("@BRANCH_QUESTION_ID", System.Data.SqlDbType.Int, 0, "BRANCH_QUESTION_ID"), New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_QUESTION_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "QUESTION_ID", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_LIST_ITEM_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LIST_ITEM_ID", System.Data.DataRowVersion.Original, Nothing)})
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
        Me.SqlInsertCommand.Connection = Me.commonConn
        Me.SqlInsertCommand.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, "TEMPLATE_KEY"), New System.Data.SqlClient.SqlParameter("@QUESTION_ID", System.Data.SqlDbType.Int, 0, "QUESTION_ID"), New System.Data.SqlClient.SqlParameter("@REQUIRED_IND", System.Data.SqlDbType.Bit, 0, "REQUIRED_IND"), New System.Data.SqlClient.SqlParameter("@QUESTION_TYPE_CODE", System.Data.SqlDbType.VarChar, 0, "QUESTION_TYPE_CODE"), New System.Data.SqlClient.SqlParameter("@QUESTION_TEXT", System.Data.SqlDbType.VarChar, 0, "QUESTION_TEXT"), New System.Data.SqlClient.SqlParameter("@RATING_COUNT", System.Data.SqlDbType.Int, 0, "RATING_COUNT"), New System.Data.SqlClient.SqlParameter("@RATING_STEP_VALUE", System.Data.SqlDbType.Int, 0, "RATING_STEP_VALUE"), New System.Data.SqlClient.SqlParameter("@RATING_INITIAL_VALUE", System.Data.SqlDbType.Int, 0, "RATING_INITIAL_VALUE"), New System.Data.SqlClient.SqlParameter("@BRANCH_QUESTION_ID", System.Data.SqlDbType.Int, 0, "BRANCH_QUESTION_ID")})
        '
        'SqlUpdateCommand
        '
        Me.SqlUpdateCommand.CommandText = resources.GetString("SqlUpdateCommand.CommandText")
        Me.SqlUpdateCommand.Connection = Me.commonConn
        Me.SqlUpdateCommand.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, "TEMPLATE_KEY"), New System.Data.SqlClient.SqlParameter("@QUESTION_ID", System.Data.SqlDbType.Int, 0, "QUESTION_ID"), New System.Data.SqlClient.SqlParameter("@REQUIRED_IND", System.Data.SqlDbType.Bit, 0, "REQUIRED_IND"), New System.Data.SqlClient.SqlParameter("@QUESTION_TYPE_CODE", System.Data.SqlDbType.VarChar, 0, "QUESTION_TYPE_CODE"), New System.Data.SqlClient.SqlParameter("@QUESTION_TEXT", System.Data.SqlDbType.VarChar, 0, "QUESTION_TEXT"), New System.Data.SqlClient.SqlParameter("@RATING_COUNT", System.Data.SqlDbType.Int, 0, "RATING_COUNT"), New System.Data.SqlClient.SqlParameter("@RATING_STEP_VALUE", System.Data.SqlDbType.Int, 0, "RATING_STEP_VALUE"), New System.Data.SqlClient.SqlParameter("@RATING_INITIAL_VALUE", System.Data.SqlDbType.Int, 0, "RATING_INITIAL_VALUE"), New System.Data.SqlClient.SqlParameter("@BRANCH_QUESTION_ID", System.Data.SqlDbType.Int, 0, "BRANCH_QUESTION_ID"), New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_QUESTION_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "QUESTION_ID", System.Data.DataRowVersion.Original, Nothing)})
        '
        'SqlDeleteCommand
        '
        Me.SqlDeleteCommand.CommandText = "DELETE FROM VARCONN.[SAT_TEMPLATE_QUESTION] WHERE (([TEMPLATE_KEY] = @Or" & _
            "iginal_TEMPLATE_KEY) AND ([QUESTION_ID] = @Original_QUESTION_ID))"
        Me.SqlDeleteCommand.Connection = Me.commonConn
        Me.SqlDeleteCommand.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_QUESTION_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "QUESTION_ID", System.Data.DataRowVersion.Original, Nothing)})
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected WithEvents hlpServerSwitch As System.Web.UI.WebControls.Image
    Protected WithEvents chkServerSwitch As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblTemplate As System.Web.UI.WebControls.Label
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents hlpTemplatePreview As System.Web.UI.WebControls.Image
    Protected WithEvents btnTemplatePreview As System.Web.UI.WebControls.Button
    Protected WithEvents hlpQuestionType As System.Web.UI.WebControls.Image
    Protected WithEvents lblQuestionType As System.Web.UI.WebControls.Label
    Protected WithEvents ddlQuestionType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents sqlQuestion As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents sqlListItems As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents SqlSelectCommand2 As System.Data.SqlClient.SqlCommand
    Protected WithEvents dsCommon As System.Data.DataSet
    Protected WithEvents lblQuestion As System.Web.UI.WebControls.Label
    Protected WithEvents lblQuestionTitle As System.Web.UI.WebControls.Image
    Protected requestItems As Array
    Protected WithEvents lblNavBarQuestion As System.Web.UI.WebControls.Label
    Protected WithEvents lblItemLabel As System.Web.UI.WebControls.Label
    Protected WithEvents hlpItemLabel As System.Web.UI.WebControls.Image
    Protected WithEvents hlpUpload As System.Web.UI.WebControls.Image
    Protected WithEvents txtItemLabel As System.Web.UI.WebControls.TextBox
    Protected WithEvents lblUpload As System.Web.UI.WebControls.Label
    Protected myNavBar As Navigation = New Navigation
    Protected myUtility As Utility = New Utility
    Protected WithEvents fuImageLoader As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents btnUp As System.Web.UI.WebControls.Button
    Protected WithEvents btnTop As System.Web.UI.WebControls.Button
    Protected WithEvents lblNavBar As System.Web.UI.WebControls.Label
    Protected WithEvents btnStart As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnPrev As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnReload As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnNext As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnLast As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnNew As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnStart2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnPrev2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnReload2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnNext2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnLast2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnNew2 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents lblNavBarList As System.Web.UI.WebControls.Label
    Protected navButtonsQuestion As Collection
    Protected navButtonsList As Collection
    Protected WithEvents hlpBranch As System.Web.UI.WebControls.Image
    Protected WithEvents lblBranch As System.Web.UI.WebControls.Label
    Protected WithEvents wneBranch As Infragistics.WebUI.WebDataInput.WebNumericEdit
    Protected WithEvents SqlInsertCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlUpdateCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlDeleteCommand1 As System.Data.SqlClient.SqlCommand
    Protected sqlCommonAdapter As New System.Data.SqlClient.SqlDataAdapter
    Protected sqlCommonAction As New System.Data.SqlClient.SqlCommand
    Protected WithEvents ListTable As System.Web.UI.WebControls.Table
    Protected WithEvents commonConn As System.Data.SqlClient.SqlConnection
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
        End If

        'Set the connection based on the machine
        myUtility.getConnection(Me.commonConn)

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
            Response.Redirect(Session("CurrentPage"))
        Else
            Session("CurrentPage") = "./List.aspx"
            Session("pageSel") = "Template"
        End If

        'if the user's password has expired send them to the password reset page
        If (Session("isExpired") = "True") Then
            Session("CurrentPage") = "./Password.aspx"
            Session("pageSel") = "Password"
            redirect(Session("CurrentPage"))
        End If

        'generate collection of navigational buttons
        myUtility.makeNavCollection(Me.navButtonsQuestion, Me.btnStart, Me.btnPrev, Me.btnReload, Me.btnNext, Me.btnLast, Me.btnNew)
        myUtility.makeNavCollection(Me.navButtonsList, Me.btnStart2, Me.btnPrev2, Me.btnReload2, Me.btnNext2, Me.btnLast2, Me.btnNew2)

        'Finalize the SQL statements
        completeSQL()

        'Check for user selection from list items if we're sure it's not a simple post-back
        If Not (IsPostBack) Then
            'get anything we can off the request line
            myUtility.getRequest(Page.Request.Url.Query())

            'load all data
            If Not (loadData()) Then
                alert("Failed to load the list items.")
            Else
                If (Session("SelectedQuestionType") = "P") Then
                    doSetQuestionLink()
                End If
                setNavBar()

                'get the template name - mr sneaky user has access and used a link
                If (Session("TemplateName") Is Nothing Or Session("TemplateName") = "") Then
                    If Not (setTemplateName()) Then
                        alert("Unable to locate template name.")
                    End If
                End If

                'get the question name - mr sneaky user has access and used a link
                If (Session("QuestionName") Is Nothing Or Session("QuestionName") = "") Then
                    If Not (setQuestionName()) Then
                        alert("Unable to locate question name.")
                    End If
                End If

                'Populate the controls on the page
                setControls()

                'set the page options based on question availability and the whether or not the page is dirty
                myUtility.optionPreSet(Session("intList"), Session("ListMax"), Me.pageOptions)

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
        Me.SqlSelectCommand1.CommandText &= " WHERE TEMPLATE_KEY=" & Session("seqTemplate") & " ORDER BY QUESTION_ID"
        Me.SqlSelectCommand2.CommandText &= " WHERE TEMPLATE_KEY=" & Session("seqTemplate") & " and QUESTION_ID = " & Session("intQuestion") & _
        " ORDER BY LIST_ITEM_ID"
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

    'set the question name
    Private Function setQuestionName() As Boolean
        sqlCommonAdapter.SelectCommand = sqlCommonAction
        sqlCommonAction.Connection = Me.commonConn

        Try
            sqlCommonAction.CommandText = "Select QUESTION_TEXT from " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION where " & _
            "TEMPLATE_KEY = " & Session("seqTemplate") & " and QUESTION_ID = " & Session("intQuestion")
            sqlCommonAdapter.Fill(Me.dsCommon, "QuestionName")
            Session("QuestionName") = Me.dsCommon.Tables("QuestionName").Rows(0).Item("QUESTION_TEXT")
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function
#End Region

#Region "Load Data"
    'Loads data into the form
    Private Function loadData() As Boolean
        Trace.Warn("Loading Data")
        'load the question data and list items
        Try
            'get question data
            If (loadQuestions()) Then
                'get list data
                If (loadListItems()) Then
                    Return True
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

    'loads the Questions for the current template
    Private Function loadQuestions() As Boolean
        Trace.Warn("Loading Questions")
        Try
            'dump the data that may exist before refilling
            If (Me.dsCommon.Tables.Contains("Questions")) Then
                Me.dsCommon.Tables("Questions").Rows.Clear()
            End If

            Me.SqlSelectCommand1.Connection = Me.commonConn
            Me.SqlSelectCommand1.CommandText = Me.SqlSelectCommand1.CommandText.Replace("VARCONN.", myUtility.getExtension())
            Me.sqlQuestion.Fill(Me.dsCommon, "Questions")
            Session("QuestionMax") = Me.dsCommon.Tables("Questions").Rows.Count()

            'determine if there is any data and if the question id has been set
            If (Session("QuestionMax") = 0) Then
                Session("intQuestion") = -1
            ElseIf (Session("intQuestion") Is Nothing) Then
                Session("intQuestion") = 1
            End If

            'make sure the question row never exceeeds the maximum questions
            If (Session("QuestionRow") > Session("QuestionMax")) Then
                Session("QuestionRow") = Session("QuestionMax") - 1
            End If

            'if the question id indicates a new record then move the question row to the maximum
            'else set it to just below the comment id's number
            If (Session("intQuestion") = -1) Then
                Session("QuestionRow") = Session("QuestionMax")
                Session("QuestionName") = ""
            Else
                Session("QuestionRow") = Session("intQuestion") - 1
                Session("QuestionName") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_TEXT")
            End If
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'loads the list items for the current question in the template
    Private Function loadListItems() As Boolean
        Trace.Warn("Loading List Items")
        Try
            'dump the data that may exist before refilling
            If (Me.dsCommon.Tables.Contains("ListItems")) Then
                Me.dsCommon.Tables("ListItems").Rows.Clear()
            End If

            Me.SqlSelectCommand2.Connection = Me.commonConn
            Me.SqlSelectCommand2.CommandText = Me.SqlSelectCommand2.CommandText.Replace("VARCONN.", myUtility.getExtension())
            Me.sqlListItems.Fill(Me.dsCommon, "ListItems")
            Session("ListMax") = Me.dsCommon.Tables("ListItems").Rows.Count()

            'determine if there is any data and if the list id has been set
            If (Session("ListMax") = 0) Then
                Session("intList") = -1
            ElseIf (Session("intList") Is Nothing) Then
                Session("intList") = 1
            End If

            'make sure the list row never exceeds the maximum list items
            If (Session("ListRow") > Session("ListMax")) Then
                Session("ListRow") = Session("ListMax") - 1
            End If

            'if the list item id indicates a new record then move the list item row to the maximum
            'else set it to just below the list item id's number
            If (Session("intList") = -1) Then
                Session("ListRow") = Session("ListMax")
            Else
                Session("ListRow") = Session("intList") - 1
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
        If (Session("intList") > 0) Then
            Try
                Me.txtItemLabel.Text = Me.dsCommon.Tables("ListItems").Rows(Session("ListRow")).Item("LIST_ITEM_TITLE")
                Me.wneBranch.Value = Me.dsCommon.Tables("ListItems").Rows(Session("ListRow")).Item("BRANCH_QUESTION_ID")
                Me.lblNavBarQuestion.Text = "Question " & Session("QuestionRow") + 1 & " of " & Session("QuestionMax")
                Me.lblNavBarList.Text = "List Item " & Session("ListRow") + 1 & " of " & Session("ListMax")
            Catch ex As Exception
                Trace.Warn(ex.ToString)
            End Try
        Else
            Me.txtItemLabel.Text = ""
            Me.wneBranch.Value = 0
            Me.lblNavBarQuestion.Text = "Question " & Session("QuestionRow") + 1 & " of " & Session("QuestionMax")
            Me.lblNavBarList.Text = "List Item " & Session("ListRow") + 1 & " of " & Session("ListMax") + 1
        End If

        'set the page options based on question availability and the whether or not the page is dirty
        myUtility.optionPreSet(Session("intList"), Session("ListMax"), Me.pageOptions)

        Session("PageOptions") = Me.pageOptions
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

    'Takes the user back to the question they were working on
    Private Sub btnUp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUp.Click
        Session("JSAction") = ""
        Session("CurrentPage") = "./Question.aspx"
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
        myNavBar.wnb_MoveTo(Me.navButtonsList, Session("ListMax"), Session("ListRow"), Switch)
        myNavBar.wnb_MoveTo(Me.navButtonsQuestion, Session("QuestionMax"), Session("QuestionRow"), Switch)
    End Sub

    'add the hyperlink to the question for the user
    Private Sub doSetQuestionLink()
        ListTable.Rows.Clear()
        Dim cell As System.Web.UI.WebControls.TableCell
        Dim row As New System.Web.UI.WebControls.TableRow
        row.Cells.Clear()
        cell = New System.Web.UI.WebControls.TableCell
        cell.HorizontalAlign = HorizontalAlign.Left
        Dim myLabel As New Label

        'resize the window to the image size + 20
        Dim myImage As System.Drawing.Image = Nothing
        Dim myStream As MemoryStream = New MemoryStream
        Dim imageWidth As Integer = 300
        Dim imageHeight As Integer = 300
        Dim processedMemStream As MemoryStream = New MemoryStream
        Try
            myStream.Write(CType(Me.dsCommon.Tables("ListItems").Rows(Session("ListRow")).Item("LIST_ITEM_IMAGE"), Byte()), 0, CType(Me.dsCommon.Tables("ListItems").Rows(Session("ListRow")).Item("LIST_ITEM_IMAGE"), Byte()).Length)
            myImage = System.Drawing.Image.FromStream(myStream)
            imageWidth = myImage.Width + 32
            imageHeight = myImage.Height + 32
        Catch ex As Exception
            Trace.Warn(ex.ToString)
        End Try

        myLabel.Text = "<a onclick=""WinPopNew('./Picture.aspx?seqTemplate=" & Session("seqTemplate") & _
        "&intQuestion=" & Session("intQuestion") & "&intList=" & _
        Session("intList") & "', " & imageWidth & ", " & imageHeight & ")"">" & _
        "<img border=0 src=""./Images/QuestionLookup.png""></a>"

        cell.Controls.Add(myLabel)
        cell.Wrap = False
        cell.EnableViewState = True
        cell.BorderStyle = BorderStyle.None
        row.Cells.Add(cell)
        ListTable.Rows.Add(row)
        ListTable.EnableViewState = True
    End Sub
#End Region

#Region "Nav Bar Events Question"
    'Moves the user to the first record
    Private Sub wnbQuestions_MoveToStart(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnStart.Click
        Session("CurrentPage") = "./Question.aspx?seqTemplate=" & Session("seqTemplate") & "&intQuestion=1"
        Session("pageSel") = "Template"
        redirect(Session("CurrentPage"))
    End Sub

    'Moves the user to the previous record
    Private Sub wnbQuestions_MovePrev(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnPrev.Click
        Session("CurrentPage") = "./Question.aspx?seqTemplate=" & Session("seqTemplate") & "&intQuestion=" & Session("intQuestion") - 1
        Session("pageSel") = "Template"
        redirect(Session("CurrentPage"))
    End Sub

    'reloads the page and the data from the database while attempting to locate the user's location
    Private Sub wnbQuestions_MoveReload(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnReload.Click
        Session("CurrentPage") = "./Question.aspx?seqTemplate=" & Session("seqTemplate") & "&intQuestion=" & Session("intQuestion")
        Session("pageSel") = "Template"
        redirect(Session("CurrentPage"))
    End Sub

    'Moves the user to the next record
    Private Sub wnbQuestions_MoveNext(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnNext.Click
        Session("CurrentPage") = "./Question.aspx?seqTemplate=" & Session("seqTemplate") & "&intQuestion=" & Session("intQuestion") + 1
        Session("pageSel") = "Template"
        redirect(Session("CurrentPage"))
    End Sub

    'Moves the user to the last record
    Private Sub wnbQuestions_MoveToEnd(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnLast.Click
        Session("CurrentPage") = "./Question.aspx?seqTemplate=" & Session("seqTemplate") & "&intQuestion=" & Session("QuestionMax")
        Session("pageSel") = "Template"
        redirect(Session("CurrentPage"))
    End Sub

    'Inserts a record into the table
    Private Sub wnbQuestions_InsertRow(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnNew.Click
        Session("CurrentPage") = "./Question.aspx?seqTemplate=" & Session("seqTemplate") & "&intQuestion=-1"
        Session("pageSel") = "Template"
        redirect(Session("CurrentPage"))
    End Sub
#End Region

#Region "Nav Bar Events List"
    'Moves the user to the first record
    Private Sub wnbList_MoveToStart(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnStart2.Click
        Session("JSAction") = ""

        If Not (loadData()) Then
            alert("Failed to load the list items for this question.")
        Else
            Try
                Session("ListRow") = 0
                Session("intList") = Me.dsCommon.Tables("ListItems").Rows(Session("ListRow")).Item("LIST_ITEM_ID")
                setNavBar()
                setControls()
                If (Session("SelectedQuestionType") = "P") Then
                    doSetQuestionLink()
                End If
                'set the page options based on question availability and the whether or not the page is dirty
                myUtility.optionPreSet(Session("intList"), Session("ListMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the first list item.")
            End Try
        End If
    End Sub

    'Moves the user to the previous record
    Private Sub wnbList_MovePrev(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnPrev2.Click
        Session("JSAction") = ""

        If Not (loadData()) Then
            alert("Failed to load the list items for this question.")
        Else
            Try
                Session("ListRow") -= 1
                Session("intList") = Me.dsCommon.Tables("ListItems").Rows(Session("ListRow")).Item("LIST_ITEM_ID")
                setNavBar()
                setControls()
                If (Session("SelectedQuestionType") = "P") Then
                    doSetQuestionLink()
                End If
                'set the page options based on question availability and the whether or not the page is dirty
                myUtility.optionPreSet(Session("intList"), Session("ListMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the previous list item.")
            End Try
        End If
    End Sub

    'reloads the page and the data from the database while attempting to locate the user's location
    Private Sub wnbList_MoveReload(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnReload2.Click
        Session("JSAction") = ""

        If Not (loadData()) Then
            alert("Failed to load the list items for this question.")
        Else
            Try
                setNavBar()
                setControls()
                If (Session("SelectedQuestionType") = "P") Then
                    doSetQuestionLink()
                End If
                'set the page options based on question availability and the whether or not the page is dirty
                myUtility.optionPreSet(Session("intList"), Session("ListMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error resetting the current list item.")
            End Try
        End If
    End Sub

    'Moves the user to the next record
    Private Sub wnbList_MoveNext(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnNext2.Click
        Session("JSAction") = ""

        If Not (loadData()) Then
            alert("Failed to load the list items for this question.")
        Else
            Try
                Session("ListRow") += 1
                Session("intList") = Me.dsCommon.Tables("ListItems").Rows(Session("ListRow")).Item("LIST_ITEM_ID")
                setNavBar()
                setControls()
                If (Session("SelectedQuestionType") = "P") Then
                    doSetQuestionLink()
                End If
                'set the page options based on question availability and the whether or not the page is dirty
                myUtility.optionPreSet(Session("intList"), Session("ListMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the next list item.")
            End Try
        End If
    End Sub

    'Moves the user to the last record
    Private Sub wnbList_MoveToEnd(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnLast2.Click
        Session("JSAction") = ""

        If Not (loadData()) Then
            alert("Failed to load the list items for this question.")
        Else
            Try
                Session("ListRow") = Session("ListMax") - 1
                Session("intList") = Me.dsCommon.Tables("ListItems").Rows(Session("ListRow")).Item("LIST_ITEM_ID")
                setNavBar()
                setControls()
                If (Session("SelectedQuestionType") = "P") Then
                    doSetQuestionLink()
                End If
                'set the page options based on question availability and the whether or not the page is dirty
                myUtility.optionPreSet(Session("intList"), Session("ListMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the last list item.")
            End Try
        End If
    End Sub

    'Inserts a record into the table
    Private Sub wnbList_InsertRow(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnNew2.Click
        Session("JSAction") = ""

        If Not (loadData()) Then
            alert("Failed to load the list items for this question.")
        Else
            Try
                Session("intList") = -1
                Session("ListRow") = Session("ListMax")
                Me.dsCommon.Tables("ListItems").Rows.Add(Me.dsCommon.Tables("ListItems").NewRow())
                Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("TEMPLATE_KEY") = Session("seqTemplate")
                Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("QUESTION_ID") = Session("intQuestion")
                Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("LIST_ITEM_ID") = Session("ListMax")
                Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("LIST_ITEM_VALUE") = -1
                Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("LIST_ITEM_TITLE") = ""
                Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("LIST_ITEM_IMAGE_TYPE") = ""
                Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("BRANCH_QUESTION_ID") = 0
                setNavBar()
                setControls()
                'set the page options based on question availability and the whether or not the page is dirty
                myUtility.optionPreSet(Session("intList"), Session("ListMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error creating new list item.")
            End Try
        End If
    End Sub
#End Region

#Region "Paqe Action"
    'Reacts to the user performing some kind of selection and submission
    Private Sub btnInsert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsert.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'refresh the data
        If (loadData()) Then
            If (Session("SelectedQuestionType") = "P") Then
                doSetQuestionLink()
            End If
        Else
            blnContinue = False
            alert("Failed to load the list items for the current question. Insert aborted.")
        End If

        'check for possible malicious code input
        If Not (myUtility.goodInput(Me.txtItemLabel, True)) Then
            alert("Possible malicious code entry in the list item text.  Insert aborted.")
            blnContinue = False
        End If

        If (blnContinue) Then
            'do the insert
            failure = doInsert()

            'reload the data
            If (loadData()) Then
                If (Session("SelectedQuestionType") = "P") Then
                    doSetQuestionLink()
                End If
            Else
                blnContinue = False
                alert("Failed to load the list items for the current question.")
            End If

            If (blnContinue) Then
                'to determine what of the nav bar buttons should be available
                setNavBar()

                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
                Session("TransExists") = False
                'set the page options based on question availability and the whether or not the page is dirty
                myUtility.optionPreSet(Session("intList"), Session("ListMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'refresh the data
        If (loadData()) Then
            If (Session("SelectedQuestionType") = "P") Then
                doSetQuestionLink()
            End If
        Else
            blnContinue = False
            alert("Failed to load the list items for the current question. Update aborted.")
        End If

        'check for possible malicious code input
        If Not (myUtility.goodInput(Me.txtItemLabel, True)) Then
            alert("Possible malicious code entry in the list item text.  Update aborted.")
            blnContinue = False
        End If

        If (blnContinue) Then
            'do the update
            failure = doUpdate()

            'reload the data
            If (loadData()) Then
                If (Session("SelectedQuestionType") = "P") Then
                    doSetQuestionLink()
                End If
            Else
                blnContinue = False
                alert("Failed to load the list items for the current question.")
            End If

            If (blnContinue) Then
                'to determine what of the nav bar buttons should be available
                setNavBar()

                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
                Session("TransExists") = False
                'set the page options based on question availability and the whether or not the page is dirty
                myUtility.optionPreSet(Session("intList"), Session("ListMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'refresh the data
        If (loadData()) Then
            If (Session("SelectedQuestionType") = "P") Then
                doSetQuestionLink()
            End If
        Else
            blnContinue = False
            alert("Failed to load the list items for the current question. Delete aborted.")
        End If

        'check for possible malicious code input
        If Not (myUtility.goodInput(Me.txtItemLabel, True)) Then
            alert("Possible malicious code entry in the list item text.  Delete aborted.")
            blnContinue = False
        End If

        If (blnContinue) Then
            'do the delete
            failure = doDelete()

            'reload the data
            If (loadData()) Then
                If (Session("SelectedQuestionType") = "P") Then
                    doSetQuestionLink()
                End If
            Else
                blnContinue = False
                alert("Failed to load the list items for the current question.")
            End If

            If (blnContinue) Then
                'to determine what of the nav bar buttons should be available
                setNavBar()

                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
                Session("TransExists") = False
                'set the page options based on question availability and the whether or not the page is dirty
                myUtility.optionPreSet(Session("intList"), Session("ListMax"), Me.pageOptions)

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
        If Not (myUtility.goodInput(Me.txtItemLabel, True)) Then
            alert("Possible malicious code entry in the list item text.  Reset aborted.")
            blnContinue = False
        End If

        If (blnContinue) Then
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
                'set the page options based on question availability and the whether or not the page is dirty
                myUtility.optionPreSet(Session("intList"), Session("ListMax"), Me.pageOptions)

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
                If (Me.txtItemLabel.Text <> "") Then
                    If (Session("intList") = -1) Then
                        'If we're inserting a new record then add it onto the end of the list items
                        If (insertList(sqlCommonAction, sqlCommonAdapter, True)) Then
                            If Not (Session("isDirty")) Then
                                Dim tempConn As Data.SqlClient.SqlConnection = New Data.SqlClient.SqlConnection
                                If (myUtility.makeDirty(sqlCommonAction, sqlCommonAdapter, Me.dsCommon, myUtility.getConnection(tempConn))) Then
                                    If (Session("Alert") <> "") Then
                                        alert(Session("Alert"))
                                    End If
                                    Session("ListMax") += 1
                                    Session("intList") = Session("ListMax")
                                    sqlCommonAction.Transaction.Commit()
                                    Session("isDirty") = True
                                Else
                                    If (Session("Alert") <> "") Then
                                        alert(Session("Alert"))
                                    End If
                                    sqlCommonAction.Transaction.Rollback()
                                    alert("Error adding list item to the question.")
                                    failure = True
                                End If
                            Else
                                Session("ListMax") += 1
                                Session("intList") = Session("ListMax")
                                sqlCommonAction.Transaction.Commit()
                            End If
                        Else
                            sqlCommonAction.Transaction.Rollback()
                            alert("Error adding list item to the question.")
                            failure = True
                        End If
                    Else
                        'If we're inserting an existing record then inline insert it ito the record set and add it to the existing surveys
                        If (insertList(sqlCommonAction, sqlCommonAdapter)) Then
                            If Not (Session("isDirty")) Then
                                Dim tempConn As Data.SqlClient.SqlConnection = New Data.SqlClient.SqlConnection
                                If (myUtility.makeDirty(sqlCommonAction, sqlCommonAdapter, Me.dsCommon, myUtility.getConnection(tempConn))) Then
                                    If (Session("Alert") <> "") Then
                                        alert(Session("Alert"))
                                    End If
                                    Session("ListMax") += 1
                                    sqlCommonAction.Transaction.Commit()
                                    Session("isDirty") = True
                                Else
                                    If (Session("Alert") <> "") Then
                                        alert(Session("Alert"))
                                    End If
                                    sqlCommonAction.Transaction.Rollback()
                                    alert("Error adding list item to the question.")
                                    failure = True
                                End If
                            Else
                                Session("ListMax") += 1
                                sqlCommonAction.Transaction.Commit()
                            End If
                        Else
                            sqlCommonAction.Transaction.Rollback()
                            alert("Error adding list item to the question.")
                            failure = True
                        End If
                    End If
                Else
                    alert("You must have some text for the list item you are adding to this question.")
                    failure = True
                    sqlCommonAction.Transaction.Rollback()
                End If
                Return failure
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("List item insertion failure.")
                Return True
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("List item insertion failure.")
            Return True
        End Try
    End Function

    'attempts to insert a list item, returns false if it cannot
    Private Function insertList(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, Optional ByVal isNewRecord As Boolean = False) As Boolean
        Try
            'parameterized the text input to allow for things like quotes to get through
            Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@LIST_ITEM_TITLE", System.Data.SqlDbType.VarChar, 64, "LIST_ITEM_TITLE")
            sqlCommonAction.Parameters.Add(param0)
            Dim param1 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@LIST_ITEM_IMAGE", System.Data.SqlDbType.VarBinary, 2147483647, "LIST_ITEM_IMAGE")
            sqlCommonAction.Parameters.Add(param1)
            Dim myByte(1) As Byte

            'process the list items based on whether or not we are working with a new item
            If (isNewRecord) Then
                'attempt to add the new record to the database
                sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM (TEMPLATE_KEY, QUESTION_ID, LIST_ITEM_ID, LIST_ITEM_VALUE, " & _
                "LIST_ITEM_TITLE, LIST_ITEM_IMAGE, LIST_ITEM_IMAGE_TYPE, BRANCH_QUESTION_ID) VALUES (" & Session("seqTemplate") & ", " & Session("intQuestion") & _
                ", " & Session("ListRow") + 1

                If (Session("SelectedQuestionType") = "C") Then
                    If (Session("ListRow") = 0) Then
                        sqlCommonAction.CommandText &= ", 1"
                    Else
                        sqlCommonAction.CommandText &= ", " & Pow(2, CType(Session("ListRow"), Integer))
                    End If
                Else
                    sqlCommonAction.CommandText &= ", " & Session("ListRow")
                End If

                sqlCommonAction.CommandText &= ", @LIST_ITEM_TITLE"
                param0.Value = Me.txtItemLabel.Text

                If (Session("SelectedQuestionType") = "P") Then
                    If (Not (Me.fuImageLoader.PostedFile Is Nothing) And (Me.fuImageLoader.PostedFile.ContentType.StartsWith("image/"))) Then
                        Dim theStream As Stream = Me.fuImageLoader.PostedFile.InputStream
                        Dim imgLength As Integer = Me.fuImageLoader.PostedFile.ContentLength
                        Dim buffer(imgLength) As Byte
                        Dim nBytesRead As Integer = theStream.Read(buffer, 0, imgLength)

                        sqlCommonAction.CommandText &= ", @LIST_ITEM_IMAGE, '" & Me.fuImageLoader.PostedFile.ContentType() & "',"
                        param1.Value = buffer
                    Else
                        param1.Value = myByte
                        sqlCommonAction.CommandText &= ", NULL, '',"
                    End If
                Else
                    param1.Value = myByte
                    sqlCommonAction.CommandText &= ", NULL, '',"
                End If

                sqlCommonAction.CommandText &= " " & Me.wneBranch.Value & ")"

                sqlCommonAction.ExecuteNonQuery()
                Return True
            Else
                'update the list ID's
                Dim lowIndex As Integer = Session("ListRow")
                Dim highIndex As Integer = Session("ListMax") - 1

                'update the list ID's and values of the list items
                While (highIndex >= lowIndex)
                    sqlCommonAction.CommandText = "Update SAT_QUESTION_LIST_ITEM SET LIST_ITEM_ID = (LIST_ITEM_ID+1) "


                    If (Session("SelectedQuestionType") = "C") Then
                        sqlCommonAction.CommandText &= ", LIST_ITEM_VALUE = " & Pow(2, highIndex + 1)
                    Else
                        sqlCommonAction.CommandText &= ", LIST_ITEM_VALUE = (LIST_ITEM_VALUE+1) "
                    End If

                    sqlCommonAction.CommandText &= "WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and QUESTION_ID = " & _
                    Session("intQuestion") & " and LIST_ITEM_ID = " & _
                    Me.dsCommon.Tables("ListItems").Rows(highIndex).Item("LIST_ITEM_ID")

                    'set dummy values
                    param0.Value = ""
                    param1.Value = myByte

                    sqlCommonAction.ExecuteNonQuery()
                    highIndex -= 1
                End While

                'attempt to add the new record to the database
                sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM (TEMPLATE_KEY, QUESTION_ID, LIST_ITEM_ID, LIST_ITEM_VALUE, " & _
                "LIST_ITEM_TITLE, LIST_ITEM_IMAGE, LIST_ITEM_IMAGE_TYPE, BRANCH_QUESTION_ID) VALUES (" & Session("seqTemplate") & ", " & Session("intQuestion") & _
                ", " & Session("ListRow") + 1 & ", " & Me.dsCommon.Tables("ListItems").Rows(lowIndex).Item("LIST_ITEM_VALUE") & _
                ", @LIST_ITEM_TITLE"

                param0.Value = Me.txtItemLabel.Text

                If (Session("SelectedQuestionType") = "P") Then
                    If (Not (Me.fuImageLoader.PostedFile Is Nothing) And (Me.fuImageLoader.PostedFile.ContentType.StartsWith("image/"))) Then
                        Dim theStream As Stream = Me.fuImageLoader.PostedFile.InputStream
                        Dim imgLength As Integer = Me.fuImageLoader.PostedFile.ContentLength
                        Dim buffer(imgLength) As Byte
                        Dim nBytesRead As Integer = theStream.Read(buffer, 0, imgLength)

                        sqlCommonAction.CommandText &= ", @LIST_ITEM_IMAGE, '" & Me.fuImageLoader.PostedFile.ContentType() & "',"
                        param1.Value = buffer
                    ElseIf (Trim("" & CType(Me.dsCommon.Tables("ListItems").Rows(lowIndex).Item("LIST_ITEM_IMAGE"), Byte()).ToString) <> "") Then
                        param1.Value = Me.dsCommon.Tables("ListItems").Rows(lowIndex).Item("LIST_ITEM_IMAGE")
                        sqlCommonAction.CommandText &= ", @LIST_ITEM_IMAGE, '" & _
                        Me.dsCommon.Tables("ListItems").Rows(lowIndex).Item("LIST_ITEM_IMAGE_TYPE") & "',"
                    Else
                        param1.Value = myByte
                        sqlCommonAction.CommandText &= ", NULL, '',"
                    End If
                Else
                    param1.Value = myByte
                    sqlCommonAction.CommandText &= ", NULL, '',"
                End If

                sqlCommonAction.CommandText &= " " & Me.wneBranch.Value & ")"

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
                If (Me.txtItemLabel.Text <> "") Then
                    If (updateList(sqlCommonAction, sqlCommonAdapter)) Then
                        If Not (Session("isDirty")) Then
                            Dim tempConn As Data.SqlClient.SqlConnection = New Data.SqlClient.SqlConnection
                            If (myUtility.makeDirty(sqlCommonAction, sqlCommonAdapter, Me.dsCommon, myUtility.getConnection(tempConn))) Then
                                If (Session("Alert") <> "") Then
                                    alert(Session("Alert"))
                                End If
                                sqlCommonAction.Transaction.Commit()
                                Session("isDirty") = True
                            Else
                                If (Session("Alert") <> "") Then
                                    alert(Session("Alert"))
                                End If
                                sqlCommonAction.Transaction.Rollback()
                                alert("Error updating the list item for this question.")
                                failure = True
                            End If
                        Else
                            sqlCommonAction.Transaction.Commit()
                        End If
                    Else
                        sqlCommonAction.Transaction.Rollback()
                        alert("Error updating the list item for this question.")
                        failure = True
                    End If
                Else
                    sqlCommonAction.Transaction.Rollback()
                    alert("You must have some text for the list item you are adding to this question.")
                    failure = True
                End If
                Return failure
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Error updating the list item for this question.")
                Return False
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Error updating the list item for this question.")
            Return False
        End Try
    End Function

    'attempts to update a list item, returns false if it cannot
    Private Function updateList(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'parameterized the text input to allow for things like quotes to get through
            Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@LIST_ITEM_TITLE", System.Data.SqlDbType.VarChar, 64, "LIST_ITEM_TITLE")
            sqlCommonAction.Parameters.Add(param0)
            Dim param1 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@LIST_ITEM_IMAGE", System.Data.SqlDbType.VarBinary, 2147483647, "LIST_ITEM_IMAGE")
            sqlCommonAction.Parameters.Add(param1)
            Dim myByte(1) As Byte

            'attempt to update the record in the database
            sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM SET LIST_ITEM_TITLE = @LIST_ITEM_TITLE "
            param0.Value = Me.txtItemLabel.Text

            If (Session("SelectedQuestionType") = "P") Then
                If (Not (Me.fuImageLoader.PostedFile Is Nothing) And (Me.fuImageLoader.PostedFile.ContentType.StartsWith("image/"))) Then
                    Dim theStream As Stream = Me.fuImageLoader.PostedFile.InputStream
                    Dim imgLength As Integer = Me.fuImageLoader.PostedFile.ContentLength
                    Dim buffer(imgLength) As Byte
                    Dim nBytesRead As Integer = theStream.Read(buffer, 0, imgLength)

                    sqlCommonAction.CommandText &= ", LIST_ITEM_IMAGE = @LIST_ITEM_IMAGE, LIST_ITEM_IMAGE_TYPE = '" & _
                    Me.fuImageLoader.PostedFile.ContentType & "'"
                    param1.Value = buffer
                Else
                    param1.Value = myByte
                End If
            Else
                param1.Value = myByte
            End If

            sqlCommonAction.CommandText &= ", BRANCH_QUESTION_ID = " & Me.wneBranch.Value

            sqlCommonAction.CommandText &= " WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and QUESTION_ID = " & _
            Session("intQuestion") & " and LIST_ITEM_ID = " & _
            Me.dsCommon.Tables("ListItems").Rows(Session("ListRow")).Item("LIST_ITEM_ID")

            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
        Return True
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
                If (deleteList(sqlCommonAction, sqlCommonAdapter)) Then
                    If Not (Session("isDirty")) Then
                        Dim tempConn As Data.SqlClient.SqlConnection = New Data.SqlClient.SqlConnection
                        If (myUtility.makeDirty(sqlCommonAction, sqlCommonAdapter, Me.dsCommon, myUtility.getConnection(tempConn))) Then
                            If (Session("Alert") <> "") Then
                                alert(Session("Alert"))
                            End If
                            Session("isDirty") = True
                            sqlCommonAction.Transaction.Commit()
                            'determine if we need to take a step back 
                            If (Session("ListMax") = 1) Then
                                Session("ListRow") = 0
                                Session("intList") = -1
                                Session("ListMax") = 0
                                Me.dsCommon.Tables("ListItems").Rows.Add(Me.dsCommon.Tables("ListItems").NewRow())
                                Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("TEMPLATE_KEY") = Session("seqTemplate")
                                Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("QUESTION_ID") = Session("intQuestion")
                                Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("LIST_ITEM_ID") = Session("ListMax")
                                Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("LIST_ITEM_VALUE") = -1
                                Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("LIST_ITEM_TITLE") = ""
                                Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("LIST_ITEM_IMAGE_TYPE") = ""
                                Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("BRANCH_QUESTION_ID") = 0
                            ElseIf (Session("intList") = Session("ListMax") And Session("ListMax") - Session("ListRow") = 1) Then
                                Session("intList") -= 1
                                Session("ListMax") -= 1
                                Session("ListRow") -= 1
                            Else
                                Session("ListMax") -= 1
                            End If
                        Else
                            If (Session("Alert") <> "") Then
                                alert(Session("Alert"))
                            End If
                            sqlCommonAction.Transaction.Rollback()
                            alert("Error deleting list item for this question.")
                            failure = True
                        End If
                    Else
                        sqlCommonAction.Transaction.Commit()
                        'determine if we need to take a step back 
                        If (Session("ListMax") = 1) Then
                            Session("ListRow") = 0
                            Session("intList") = -1
                            Session("ListMax") = 0
                            Me.dsCommon.Tables("ListItems").Rows.Add(Me.dsCommon.Tables("ListItems").NewRow())
                            Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("TEMPLATE_KEY") = Session("seqTemplate")
                            Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("QUESTION_ID") = Session("intQuestion")
                            Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("LIST_ITEM_ID") = Session("ListMax")
                            Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("LIST_ITEM_VALUE") = -1
                            Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("LIST_ITEM_TITLE") = ""
                            Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("LIST_ITEM_IMAGE_TYPE") = ""
                            Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("BRANCH_QUESTION_ID") = 0
                        ElseIf (Session("intList") = Session("ListMax") And Session("ListMax") - Session("ListRow") = 1) Then
                            Session("intList") -= 1
                            Session("ListMax") -= 1
                            Session("ListRow") -= 1
                        Else
                            Session("ListMax") -= 1
                        End If
                    End If
                Else
                    sqlCommonAction.Transaction.Rollback()
                    alert("Error deleting the list item for this question.")
                    failure = True
                End If
                Return failure
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Error deleting the list item for this question.")
                Return False
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Error deleting the list item for this question.")
            Return False
        End Try
    End Function

    'attempts to delete a comment, returns false if it cannot
    Private Function deleteList(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'attempt to delete the record from the database
            sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM where TEMPLATE_KEY = " & Session("seqTemplate") & _
                " and QUESTION_ID = " & Session("intQuestion") & " and LIST_ITEM_ID = " & Session("intList")
            sqlCommonAction.ExecuteNonQuery()

            'update the list ID's
            Dim lowIndex As Integer = Session("ListRow")
            Dim highIndex As Integer = Me.dsCommon.Tables("ListItems").Rows.Count()

            'update the list ID's and values of the list items
            While (lowIndex < highIndex)
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM SET LIST_ITEM_ID = (LIST_ITEM_ID-1) "

                If (Session("SelectedQuestionType") = "C") Then
                    sqlCommonAction.CommandText &= ", LIST_ITEM_VALUE = " & Pow(2, lowIndex - 1)
                Else
                    sqlCommonAction.CommandText &= ", LIST_ITEM_VALUE = (LIST_ITEM_VALUE-1) "
                End If

                sqlCommonAction.CommandText &= "WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and QUESTION_ID = " & _
                Session("intQuestion") & " and LIST_ITEM_ID = " & _
                Me.dsCommon.Tables("ListItems").Rows(lowIndex).Item("LIST_ITEM_ID")

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
        loadData()
        If (Session("intList") = -1) Then
            Me.dsCommon.Tables("ListItems").Rows.Add(Me.dsCommon.Tables("ListItems").NewRow())
            Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("TEMPLATE_KEY") = Session("seqTemplate")
            Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("QUESTION_ID") = Session("intQuestion")
            Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("LIST_ITEM_ID") = Session("ListMax")
            Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("LIST_ITEM_VALUE") = -1
            Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("LIST_ITEM_TITLE") = ""
            Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("LIST_ITEM_IMAGE_TYPE") = ""
            Me.dsCommon.Tables("ListItems").Rows(Session("ListMax")).Item("BRANCH_QUESTION_ID") = 0
        Else
            If (Session("SelectedQuestionType") = "P") Then
                doSetQuestionLink()
            End If
        End If
        Return False
    End Function
#End Region
End Class
