Imports System.Collections.Specialized
Public Class Question
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Question))
        Me.sqlQuestion = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlDeleteCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlInsertCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlUpdateCommand1 = New System.Data.SqlClient.SqlCommand
        Me.commonConn = New System.Data.SqlClient.SqlConnection
        Me.sqlListItems = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlSelectCommand2 = New System.Data.SqlClient.SqlCommand
        Me.dsCommon = New System.Data.DataSet
        Me.sqlQuestionType = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlSelectCommand3 = New System.Data.SqlClient.SqlCommand
        Me.sqlQuestionGroups = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlSelectCommand4 = New System.Data.SqlClient.SqlCommand
        Me.sqlChartTypes = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlSelectCommand5 = New System.Data.SqlClient.SqlCommand
        Me.sqlQueries = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlSelectCommand6 = New System.Data.SqlClient.SqlCommand
        Me.SqlInsertCommand = New System.Data.SqlClient.SqlCommand
        Me.SqlUpdateCommand = New System.Data.SqlClient.SqlCommand
        Me.SqlDeleteCommand = New System.Data.SqlClient.SqlCommand
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'sqlQuestion
        '
        Me.sqlQuestion.DeleteCommand = Me.SqlDeleteCommand1
        Me.sqlQuestion.InsertCommand = Me.SqlInsertCommand1
        Me.sqlQuestion.SelectCommand = Me.SqlSelectCommand1
        Me.sqlQuestion.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_TEMPLATE_QUESTION", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("TEMPLATE_KEY", "TEMPLATE_KEY"), New System.Data.Common.DataColumnMapping("QUESTION_ID", "QUESTION_ID"), New System.Data.Common.DataColumnMapping("REQUIRED_IND", "REQUIRED_IND"), New System.Data.Common.DataColumnMapping("PAGE_BREAK_IND", "PAGE_BREAK_IND"), New System.Data.Common.DataColumnMapping("FILTER_IND", "FILTER_IND"), New System.Data.Common.DataColumnMapping("NOT_APPLICABLE_IND", "NOT_APPLICABLE_IND"), New System.Data.Common.DataColumnMapping("QUESTION_TYPE_CODE", "QUESTION_TYPE_CODE"), New System.Data.Common.DataColumnMapping("QUESTION_TEXT", "QUESTION_TEXT"), New System.Data.Common.DataColumnMapping("CHART_TYPE_CODE", "CHART_TYPE_CODE"), New System.Data.Common.DataColumnMapping("RATING_COUNT", "RATING_COUNT"), New System.Data.Common.DataColumnMapping("RATING_STEP_VALUE", "RATING_STEP_VALUE"), New System.Data.Common.DataColumnMapping("RATING_INITIAL_VALUE", "RATING_INITIAL_VALUE"), New System.Data.Common.DataColumnMapping("BRANCH_QUESTION_ID", "BRANCH_QUESTION_ID"), New System.Data.Common.DataColumnMapping("QUESTION_GROUP_ID", "QUESTION_GROUP_ID"), New System.Data.Common.DataColumnMapping("QUERY_NAME", "QUERY_NAME")})})
        Me.sqlQuestion.UpdateCommand = Me.SqlUpdateCommand1
        '
        'SqlDeleteCommand1
        '
        Me.SqlDeleteCommand1.CommandText = "DELETE FROM VARCONN.[SAT_TEMPLATE_QUESTION] WHERE (([TEMPLATE_KEY] = @Or" & _
            "iginal_TEMPLATE_KEY) AND ([QUESTION_ID] = @Original_QUESTION_ID))"
        Me.SqlDeleteCommand1.Connection = Me.commonConn
        Me.SqlDeleteCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_QUESTION_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "QUESTION_ID", System.Data.DataRowVersion.Original, Nothing)})
        '
        'SqlInsertCommand1
        '
        Me.SqlInsertCommand1.CommandText = resources.GetString("SqlInsertCommand1.CommandText")
        Me.SqlInsertCommand1.Connection = Me.commonConn
        Me.SqlInsertCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, "TEMPLATE_KEY"), New System.Data.SqlClient.SqlParameter("@QUESTION_ID", System.Data.SqlDbType.Int, 0, "QUESTION_ID"), New System.Data.SqlClient.SqlParameter("@REQUIRED_IND", System.Data.SqlDbType.Bit, 0, "REQUIRED_IND"), New System.Data.SqlClient.SqlParameter("@PAGE_BREAK_IND", System.Data.SqlDbType.Bit, 0, "PAGE_BREAK_IND"), New System.Data.SqlClient.SqlParameter("@FILTER_IND", System.Data.SqlDbType.Bit, 0, "FILTER_IND"), New System.Data.SqlClient.SqlParameter("@NOT_APPLICABLE_IND", System.Data.SqlDbType.Bit, 0, "NOT_APPLICABLE_IND"), New System.Data.SqlClient.SqlParameter("@QUESTION_TYPE_CODE", System.Data.SqlDbType.VarChar, 0, "QUESTION_TYPE_CODE"), New System.Data.SqlClient.SqlParameter("@QUESTION_TEXT", System.Data.SqlDbType.VarChar, 0, "QUESTION_TEXT"), New System.Data.SqlClient.SqlParameter("@CHART_TYPE_CODE", System.Data.SqlDbType.VarChar, 0, "CHART_TYPE_CODE"), New System.Data.SqlClient.SqlParameter("@RATING_COUNT", System.Data.SqlDbType.Int, 0, "RATING_COUNT"), New System.Data.SqlClient.SqlParameter("@RATING_STEP_VALUE", System.Data.SqlDbType.Int, 0, "RATING_STEP_VALUE"), New System.Data.SqlClient.SqlParameter("@RATING_INITIAL_VALUE", System.Data.SqlDbType.Int, 0, "RATING_INITIAL_VALUE"), New System.Data.SqlClient.SqlParameter("@BRANCH_QUESTION_ID", System.Data.SqlDbType.Int, 0, "BRANCH_QUESTION_ID"), New System.Data.SqlClient.SqlParameter("@QUESTION_GROUP_ID", System.Data.SqlDbType.Int, 0, "QUESTION_GROUP_ID"), New System.Data.SqlClient.SqlParameter("@QUERY_NAME", System.Data.SqlDbType.VarChar, 0, "QUERY_NAME")})
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
        Me.SqlUpdateCommand1.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, "TEMPLATE_KEY"), New System.Data.SqlClient.SqlParameter("@QUESTION_ID", System.Data.SqlDbType.Int, 0, "QUESTION_ID"), New System.Data.SqlClient.SqlParameter("@REQUIRED_IND", System.Data.SqlDbType.Bit, 0, "REQUIRED_IND"), New System.Data.SqlClient.SqlParameter("@PAGE_BREAK_IND", System.Data.SqlDbType.Bit, 0, "PAGE_BREAK_IND"), New System.Data.SqlClient.SqlParameter("@FILTER_IND", System.Data.SqlDbType.Bit, 0, "FILTER_IND"), New System.Data.SqlClient.SqlParameter("@NOT_APPLICABLE_IND", System.Data.SqlDbType.Bit, 0, "NOT_APPLICABLE_IND"), New System.Data.SqlClient.SqlParameter("@QUESTION_TYPE_CODE", System.Data.SqlDbType.VarChar, 0, "QUESTION_TYPE_CODE"), New System.Data.SqlClient.SqlParameter("@QUESTION_TEXT", System.Data.SqlDbType.VarChar, 0, "QUESTION_TEXT"), New System.Data.SqlClient.SqlParameter("@CHART_TYPE_CODE", System.Data.SqlDbType.VarChar, 0, "CHART_TYPE_CODE"), New System.Data.SqlClient.SqlParameter("@RATING_COUNT", System.Data.SqlDbType.Int, 0, "RATING_COUNT"), New System.Data.SqlClient.SqlParameter("@RATING_STEP_VALUE", System.Data.SqlDbType.Int, 0, "RATING_STEP_VALUE"), New System.Data.SqlClient.SqlParameter("@RATING_INITIAL_VALUE", System.Data.SqlDbType.Int, 0, "RATING_INITIAL_VALUE"), New System.Data.SqlClient.SqlParameter("@BRANCH_QUESTION_ID", System.Data.SqlDbType.Int, 0, "BRANCH_QUESTION_ID"), New System.Data.SqlClient.SqlParameter("@QUESTION_GROUP_ID", System.Data.SqlDbType.Int, 0, "QUESTION_GROUP_ID"), New System.Data.SqlClient.SqlParameter("@QUERY_NAME", System.Data.SqlDbType.VarChar, 0, "QUERY_NAME"), New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_QUESTION_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "QUESTION_ID", System.Data.DataRowVersion.Original, Nothing)})
        '
        'commonConn
        '
        Me.commonConn.FireInfoMessageEventOnUserErrors = False
        '
        'sqlListItems
        '
        Me.sqlListItems.DeleteCommand = Me.SqlDeleteCommand
        Me.sqlListItems.InsertCommand = Me.SqlInsertCommand
        Me.sqlListItems.SelectCommand = Me.SqlSelectCommand2
        Me.sqlListItems.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_QUESTION_LIST_ITEM", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("TEMPLATE_KEY", "TEMPLATE_KEY"), New System.Data.Common.DataColumnMapping("QUESTION_ID", "QUESTION_ID"), New System.Data.Common.DataColumnMapping("LIST_ITEM_ID", "LIST_ITEM_ID"), New System.Data.Common.DataColumnMapping("LIST_ITEM_VALUE", "LIST_ITEM_VALUE"), New System.Data.Common.DataColumnMapping("LIST_ITEM_TITLE", "LIST_ITEM_TITLE"), New System.Data.Common.DataColumnMapping("LIST_ITEM_IMAGE", "LIST_ITEM_IMAGE"), New System.Data.Common.DataColumnMapping("LIST_ITEM_IMAGE_TYPE", "LIST_ITEM_IMAGE_TYPE"), New System.Data.Common.DataColumnMapping("BRANCH_QUESTION_ID", "BRANCH_QUESTION_ID")})})
        Me.sqlListItems.UpdateCommand = Me.SqlUpdateCommand
        '
        'SqlSelectCommand2
        '
        Me.SqlSelectCommand2.CommandText = "SELECT TEMPLATE_KEY, QUESTION_ID, LIST_ITEM_ID, LIST_ITEM_VALUE, LIST_ITEM_TITLE," & _
            " LIST_ITEM_IMAGE, LIST_ITEM_IMAGE_TYPE, BRANCH_QUESTION_ID FROM VARCONN.SAT" & _
            "_QUESTION_LIST_ITEM"
        Me.SqlSelectCommand2.Connection = Me.commonConn
        '
        'dsCommon
        '
        Me.dsCommon.DataSetName = "NewDataSet"
        Me.dsCommon.Locale = New System.Globalization.CultureInfo("en-US")
        '
        'sqlQuestionType
        '
        Me.sqlQuestionType.SelectCommand = Me.SqlSelectCommand3
        Me.sqlQuestionType.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_QUESTION_TYPE", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("QUESTION_TYPE_TITLE", "QUESTION_TYPE_TITLE"), New System.Data.Common.DataColumnMapping("QUESTION_TYPE_CODE", "QUESTION_TYPE_CODE")})})
        '
        'SqlSelectCommand3
        '
        Me.SqlSelectCommand3.CommandText = "SELECT QUESTION_TYPE_TITLE, QUESTION_TYPE_CODE FROM VARCONN.SAT_QUESTION_TYP" & _
            "E ORDER BY QUESTION_TYPE_CODE"
        Me.SqlSelectCommand3.Connection = Me.commonConn
        '
        'sqlQuestionGroups
        '
        Me.sqlQuestionGroups.SelectCommand = Me.SqlSelectCommand4
        Me.sqlQuestionGroups.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_TEMPLATE_QUESTION_GROUP", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("QUESTION_GROUP_ID", "QUESTION_GROUP_ID"), New System.Data.Common.DataColumnMapping("QUESTION_GROUP_TITLE", "QUESTION_GROUP_TITLE"), New System.Data.Common.DataColumnMapping("TEMPLATE_KEY", "TEMPLATE_KEY")})})
        '
        'SqlSelectCommand4
        '
        Me.SqlSelectCommand4.CommandText = "SELECT QUESTION_GROUP_ID, QUESTION_GROUP_TITLE, TEMPLATE_KEY FROM VARCONN.SA" & _
            "T_TEMPLATE_QUESTION_GROUP"
        Me.SqlSelectCommand4.Connection = Me.commonConn
        '
        'sqlChartTypes
        '
        Me.sqlChartTypes.SelectCommand = Me.SqlSelectCommand5
        Me.sqlChartTypes.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_CHART_TYPE", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("CHART_TYPE_CODE", "CHART_TYPE_CODE"), New System.Data.Common.DataColumnMapping("CHART_TYPE_TITLE", "CHART_TYPE_TITLE")})})
        '
        'SqlSelectCommand5
        '
        Me.SqlSelectCommand5.CommandText = "SELECT CHART_TYPE_CODE, CHART_TYPE_TITLE FROM VARCONN.SAT_CHART_TYPE ORDER B" & _
            "Y CHART_TYPE_SEQ_NO"
        Me.SqlSelectCommand5.Connection = Me.commonConn
        '
        'sqlQueries
        '
        Me.sqlQueries.SelectCommand = Me.SqlSelectCommand6
        Me.sqlQueries.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_TEMPLATE_QUESTION_QUERIES", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("QUERY_NAME", "QUERY_NAME"), New System.Data.Common.DataColumnMapping("QUERY_RESULT_LABEL", "QUERY_RESULT_LABEL"), New System.Data.Common.DataColumnMapping("QUERY_RESULT_VALUE", "QUERY_RESULT_VALUE"), New System.Data.Common.DataColumnMapping("QUERY_LIST_INITIAL_VALUE", "QUERY_LIST_INITIAL_VALUE"), New System.Data.Common.DataColumnMapping("QUERY_STRING", "QUERY_STRING"), New System.Data.Common.DataColumnMapping("QUERY_DATABASE_NAME", "QUERY_DATABASE_NAME")})})
        '
        'SqlSelectCommand6
        '
        Me.SqlSelectCommand6.CommandText = "SELECT QUERY_NAME, QUERY_RESULT_LABEL, QUERY_RESULT_VALUE, QUERY_LIST_INITIAL_VAL" & _
            "UE, QUERY_STRING, QUERY_DATABASE_NAME FROM VARCONN.SAT_TEMPLATE_QUESTION_QU" & _
            "ERIES ORDER BY QUERY_NAME"
        Me.SqlSelectCommand6.Connection = Me.commonConn
        '
        'SqlInsertCommand
        '
        Me.SqlInsertCommand.CommandText = resources.GetString("SqlInsertCommand.CommandText")
        Me.SqlInsertCommand.Connection = Me.commonConn
        Me.SqlInsertCommand.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, "TEMPLATE_KEY"), New System.Data.SqlClient.SqlParameter("@QUESTION_ID", System.Data.SqlDbType.Int, 0, "QUESTION_ID"), New System.Data.SqlClient.SqlParameter("@LIST_ITEM_ID", System.Data.SqlDbType.Int, 0, "LIST_ITEM_ID"), New System.Data.SqlClient.SqlParameter("@LIST_ITEM_VALUE", System.Data.SqlDbType.Int, 0, "LIST_ITEM_VALUE"), New System.Data.SqlClient.SqlParameter("@LIST_ITEM_TITLE", System.Data.SqlDbType.VarChar, 0, "LIST_ITEM_TITLE"), New System.Data.SqlClient.SqlParameter("@LIST_ITEM_IMAGE", System.Data.SqlDbType.Image, 0, "LIST_ITEM_IMAGE"), New System.Data.SqlClient.SqlParameter("@LIST_ITEM_IMAGE_TYPE", System.Data.SqlDbType.VarChar, 0, "LIST_ITEM_IMAGE_TYPE"), New System.Data.SqlClient.SqlParameter("@BRANCH_QUESTION_ID", System.Data.SqlDbType.Int, 0, "BRANCH_QUESTION_ID")})
        '
        'SqlUpdateCommand
        '
        Me.SqlUpdateCommand.CommandText = resources.GetString("SqlUpdateCommand.CommandText")
        Me.SqlUpdateCommand.Connection = Me.commonConn
        Me.SqlUpdateCommand.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, "TEMPLATE_KEY"), New System.Data.SqlClient.SqlParameter("@QUESTION_ID", System.Data.SqlDbType.Int, 0, "QUESTION_ID"), New System.Data.SqlClient.SqlParameter("@LIST_ITEM_ID", System.Data.SqlDbType.Int, 0, "LIST_ITEM_ID"), New System.Data.SqlClient.SqlParameter("@LIST_ITEM_VALUE", System.Data.SqlDbType.Int, 0, "LIST_ITEM_VALUE"), New System.Data.SqlClient.SqlParameter("@LIST_ITEM_TITLE", System.Data.SqlDbType.VarChar, 0, "LIST_ITEM_TITLE"), New System.Data.SqlClient.SqlParameter("@LIST_ITEM_IMAGE", System.Data.SqlDbType.Image, 0, "LIST_ITEM_IMAGE"), New System.Data.SqlClient.SqlParameter("@LIST_ITEM_IMAGE_TYPE", System.Data.SqlDbType.VarChar, 0, "LIST_ITEM_IMAGE_TYPE"), New System.Data.SqlClient.SqlParameter("@BRANCH_QUESTION_ID", System.Data.SqlDbType.Int, 0, "BRANCH_QUESTION_ID"), New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_QUESTION_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "QUESTION_ID", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_LIST_ITEM_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LIST_ITEM_ID", System.Data.DataRowVersion.Original, Nothing)})
        '
        'SqlDeleteCommand
        '
        Me.SqlDeleteCommand.CommandText = "DELETE FROM VARCONN.[SAT_QUESTION_LIST_ITEM] WHERE (([TEMPLATE_KEY] = @O" & _
            "riginal_TEMPLATE_KEY) AND ([QUESTION_ID] = @Original_QUESTION_ID) AND ([LIST_ITE" & _
            "M_ID] = @Original_LIST_ITEM_ID))"
        Me.SqlDeleteCommand.Connection = Me.commonConn
        Me.SqlDeleteCommand.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@Original_TEMPLATE_KEY", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "TEMPLATE_KEY", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_QUESTION_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "QUESTION_ID", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_LIST_ITEM_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "LIST_ITEM_ID", System.Data.DataRowVersion.Original, Nothing)})
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected WithEvents sqlQuestion As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents sqlListItems As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents dsCommon As System.Data.DataSet
    Protected requestItems As Array
    Protected WithEvents sqlQuestionType As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents SqlSelectCommand3 As System.Data.SqlClient.SqlCommand
    Protected myNavBar As Navigation = New Navigation
    Protected myUtility As Utility = New Utility
    Protected WithEvents Td1 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected navButtons As Collection
    Protected sqlCommonAdapter As New System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents lblPlaceHolder As System.Web.UI.WebControls.Label
    Protected WithEvents rptrListItems As System.Web.UI.WebControls.Repeater
    Protected WithEvents SqlSelectCommand2 As System.Data.SqlClient.SqlCommand
    Protected WithEvents sqlQuestionGroups As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents SqlSelectCommand4 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlInsertCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlUpdateCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlDeleteCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlSelectCommand5 As System.Data.SqlClient.SqlCommand
    Protected WithEvents sqlChartTypes As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents lblQuestionTitle As System.Web.UI.WebControls.Image
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
    Protected WithEvents hlpQuestionType As System.Web.UI.WebControls.Image
    Protected WithEvents lblQuestionType As System.Web.UI.WebControls.Label
    Protected WithEvents ddlQuestionType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents hlpQuestion As System.Web.UI.WebControls.Image
    Protected WithEvents lblQuestion As System.Web.UI.WebControls.Label
    Protected WithEvents txtQuestion As System.Web.UI.WebControls.TextBox
    Protected WithEvents hlpTopIsNotApp As System.Web.UI.WebControls.Image
    Protected WithEvents chkTopIsNotApp As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblTopIsNotApp As System.Web.UI.WebControls.Label
    Protected WithEvents hlpInitial As System.Web.UI.WebControls.Image
    Protected WithEvents lblInitial As System.Web.UI.WebControls.Label
    Protected WithEvents wneInitial As Infragistics.WebUI.WebDataInput.WebNumericEdit
    Protected WithEvents hlpNumber As System.Web.UI.WebControls.Image
    Protected WithEvents lblValues As System.Web.UI.WebControls.Label
    Protected WithEvents wneCount As Infragistics.WebUI.WebDataInput.WebNumericEdit
    Protected WithEvents hlpInterval As System.Web.UI.WebControls.Image
    Protected WithEvents lblInterval As System.Web.UI.WebControls.Label
    Protected WithEvents wneInterval As Infragistics.WebUI.WebDataInput.WebNumericEdit
    Protected WithEvents hlpRequired As System.Web.UI.WebControls.Image
    Protected WithEvents chkbxRequired As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblRequired As System.Web.UI.WebControls.Label
    Protected WithEvents hlpFilter As System.Web.UI.WebControls.Image
    Protected WithEvents chkbxFilter As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblFilter As System.Web.UI.WebControls.Label
    Protected WithEvents hlpPageBreak As System.Web.UI.WebControls.Image
    Protected WithEvents chkbxPageBreak As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblPageBreak As System.Web.UI.WebControls.Label
    Protected WithEvents hlpGraphType As System.Web.UI.WebControls.Image
    Protected WithEvents lblGraphType As System.Web.UI.WebControls.Label
    Protected WithEvents ddlGraphType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents hlpQuestionGroup As System.Web.UI.WebControls.Image
    Protected WithEvents lblQuestionGroup As System.Web.UI.WebControls.Label
    Protected WithEvents ddlQuestionGroup As System.Web.UI.WebControls.DropDownList
    Protected WithEvents hlpBranch As System.Web.UI.WebControls.Image
    Protected WithEvents lblBranch As System.Web.UI.WebControls.Label
    Protected WithEvents wneBranch As Infragistics.WebUI.WebDataInput.WebNumericEdit
    Protected WithEvents hlpPageBreak2 As System.Web.UI.WebControls.Image
    Protected WithEvents chkbxPageBreak2 As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblPageBreak2 As System.Web.UI.WebControls.Label
    Protected WithEvents hlpListItems As System.Web.UI.WebControls.Image
    Protected WithEvents lblListItems As System.Web.UI.WebControls.Label
    Protected WithEvents btnInsert As System.Web.UI.WebControls.Button
    Protected WithEvents btnUpdate As System.Web.UI.WebControls.Button
    Protected WithEvents btnDelete As System.Web.UI.WebControls.Button
    Protected WithEvents btnReset As System.Web.UI.WebControls.Button
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents commonConn As System.Data.SqlClient.SqlConnection
    Friend WithEvents SqlDeleteCommand As System.Data.SqlClient.SqlCommand
    Friend WithEvents SqlInsertCommand As System.Data.SqlClient.SqlCommand
    Friend WithEvents SqlUpdateCommand As System.Data.SqlClient.SqlCommand
    Protected WithEvents hlpQuerySelect As System.Web.UI.WebControls.Image
    Protected WithEvents lblQuerySelect As System.Web.UI.WebControls.Label
    Protected WithEvents ddlQuerySelect As System.Web.UI.WebControls.DropDownList
    Protected WithEvents sqlQueries As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents SqlSelectCommand6 As System.Data.SqlClient.SqlCommand
    Protected sqlCommonAction As New System.Data.SqlClient.SqlCommand
    Protected pageOptions As Integer
    Protected WithEvents JSAction As String


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
            Session("CurrentPage") = "./Question.aspx"
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
                alert("Failed to load the questions.")
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
                myUtility.optionPreSet(Session("intQuestion"), Session("QuestionMax"), Me.pageOptions)

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
        Me.SqlSelectCommand2.CommandText &= " WHERE TEMPLATE_KEY=" & Session("seqTemplate")
        Me.SqlSelectCommand4.CommandText &= " WHERE TEMPLATE_KEY=" & Session("seqTemplate") & " ORDER BY QUESTION_GROUP_ID"
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
        'load the question data and any list items necessary
        Try
            If Not (IsPostBack) Then
                If Not (loadTypes()) Then
                    Return False
                End If

                If Not (loadChartTypes()) Then
                    Return False
                End If

                If Not (loadQuestionGroups()) Then
                    Return False
                End If

                If Not (loadQueries()) Then
                    Return False
                End If
            End If

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

    'loads the types of questions
    Private Function loadTypes() As Boolean
        Trace.Warn("Loading Question Types")
        Try
            'dump the current data before refilling
            If (Me.dsCommon.Tables.Contains("Types")) Then
                Me.dsCommon.Tables("Types").Rows.Clear()
            End If

            'get the data and bind to a drop down list
            Me.sqlQuestionType.SelectCommand.Connection = Me.commonConn
            Me.sqlQuestionType.SelectCommand.CommandText = Me.sqlQuestionType.SelectCommand.CommandText.Replace("VARCONN.", myUtility.getExtension())
            Me.sqlQuestionType.Fill(Me.dsCommon, "Types")
            Me.ddlQuestionType.DataSource = Me.dsCommon.Tables("Types")
            Me.ddlQuestionType.DataTextField = Me.dsCommon.Tables("Types").Columns("QUESTION_TYPE_TITLE").ToString
            Me.ddlQuestionType.DataValueField = Me.dsCommon.Tables("Types").Columns("QUESTION_TYPE_CODE").ToString
            Me.ddlQuestionType.DataBind()
            Dim li As System.Web.UI.WebControls.ListItem = New System.Web.UI.WebControls.ListItem
            li.Text = "-- Select a Question Type --"
            li.Value = "*"
            Me.ddlQuestionType.Items.Insert(0, li)
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Unable to load the types of questions.")
            Return False
        End Try
    End Function

    'loads the types of charts available
    Private Function loadChartTypes() As Boolean
        Trace.Warn("Loading Chart Types")
        Try
            'dump the current data before refilling
            If (Me.dsCommon.Tables.Contains("ChartTypes")) Then
                Me.dsCommon.Tables("ChartTypes").Rows.Clear()
            End If

            'get the data and bind to a drop down list
            Me.sqlChartTypes.SelectCommand.Connection = Me.commonConn
            Me.sqlChartTypes.SelectCommand.CommandText = Me.sqlChartTypes.SelectCommand.CommandText.Replace("VARCONN.", myUtility.getExtension())
            Me.sqlChartTypes.Fill(Me.dsCommon, "ChartTypes")
            Me.ddlGraphType.DataSource = Me.dsCommon.Tables("ChartTypes")
            Me.ddlGraphType.DataTextField = Me.dsCommon.Tables("ChartTypes").Columns("CHART_TYPE_TITLE").ToString
            Me.ddlGraphType.DataValueField = Me.dsCommon.Tables("ChartTypes").Columns("CHART_TYPE_CODE").ToString
            Me.ddlGraphType.DataBind()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Unable to load the types of charts.")
            Return False
        End Try
    End Function

    'loads the Questions for the current template
    Private Function loadQuestions(Optional ByVal override As Boolean = False) As Boolean
        Trace.Warn("Loading Questions")
        Try
            'dump the data that may exist before refilling
            If (Me.dsCommon.Tables.Contains("Questions")) Then
                Me.dsCommon.Tables("Questions").Rows.Clear()
            End If

            Me.sqlQuestion.SelectCommand.Connection = Me.commonConn
            Me.sqlQuestion.SelectCommand.CommandText = Me.sqlQuestion.SelectCommand.CommandText.Replace("VARCONN.", myUtility.getExtension())
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
                If Not (override) Then
                    Session("SelectedQuestionType") = "*"
                    Session("SelectedChartType") = "PIE"
                End If
                Session("ReportGroupID") = 0
            Else
                Session("QuestionRow") = Session("intQuestion") - 1
                Session("QuestionName") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_TEXT")
                If Not (override) Then
                    Session("SelectedQuestionType") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_TYPE_CODE")
                    Session("SelectedChartType") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("CHART_TYPE_CODE")
                End If
                Session("ReportGroupID") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_GROUP_ID")
            End If

            'set the question type drop-down list
            If Not (override) Then
                setQuestionType()
                setChartType()
            End If
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'loads the list items for the current question in the template
    Private Function loadListItems(Optional ByVal override As String = "") As Boolean
        Trace.Warn("Loading List Items")
        If (override = "All") Then
            Try
                'load the list items ordered which is useful during inserts/updates
                Dim oldCommand As String = Me.SqlSelectCommand2.CommandText
                oldCommand = Me.SqlSelectCommand2.CommandText
                Me.SqlSelectCommand2.CommandText &= " ORDER BY TEMPLATE_KEY, QUESTION_ID, LIST_ITEM_ID"
                Me.SqlSelectCommand2.CommandText = Me.SqlSelectCommand2.CommandText.Replace("VARCONN.", myUtility.getExtension())
                Trace.Warn(Me.SqlSelectCommand2.CommandText)

                'dump the old data before filling
                If (Me.dsCommon.Tables.Contains("ListItems")) Then
                    Me.dsCommon.Tables("ListItems").Rows.Clear()
                End If

                'fill the data
                Me.sqlListItems.SelectCommand.Connection = Me.commonConn
                Me.sqlListItems.Fill(Me.dsCommon, "ListItems")
                Me.SqlSelectCommand2.CommandText = oldCommand
                Return True
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                Return False
            End Try
        Else
            Try
                'load the list items normally since we're not doing an insert/update
                Dim oldCommand As String = Me.SqlSelectCommand2.CommandText
                oldCommand = Me.SqlSelectCommand2.CommandText
                Me.SqlSelectCommand2.CommandText &= " AND QUESTION_ID=" & Session("intQuestion") & " ORDER BY TEMPLATE_KEY, QUESTION_ID, LIST_ITEM_ID"
                Me.SqlSelectCommand2.CommandText = Me.SqlSelectCommand2.CommandText.Replace("VARCONN.", myUtility.getExtension())
                Trace.Warn(Me.SqlSelectCommand2.CommandText)

                'make sure we get rid of any old data hanging around if we're reloading from a page selection
                If (override = "Reload") Then
                    'doesn't really matter if this fails or not we just want to make sure there's no data in it
                    Try
                        Me.dsCommon.Tables("ListItems").Rows.Clear()
                    Catch ex As Exception
                        Trace.Warn(ex.ToString)
                    End Try
                End If

                'dump the old data before filling
                If (Me.dsCommon.Tables.Contains("ListItems")) Then
                    Me.dsCommon.Tables("ListItems").Rows.Clear()
                End If

                'fill and bind the data to the repeater
                Me.sqlListItems.SelectCommand.Connection = Me.commonConn
                Me.sqlListItems.Fill(Me.dsCommon, "ListItems")
                Me.SqlSelectCommand2.CommandText = oldCommand
                Me.rptrListItems.DataSource = Me.dsCommon.Tables("ListItems")
                Me.rptrListItems.DataBind()
                Me.SqlSelectCommand2.CommandText = oldCommand
                Return True
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                Return False
            End Try
        End If
    End Function

    'loads the report groups for the current template
    Private Function loadQuestionGroups(Optional ByVal override As String = "") As Boolean
        Trace.Warn("Loading Question Groups")
        Try
            'dump the current data before refilling
            If (Me.dsCommon.Tables.Contains("QuestionGroups")) Then
                Me.dsCommon.Tables("QuestionGroups").Rows.Clear()
            End If

            'get the data and bind to a drop down list
            Me.sqlQuestionGroups.SelectCommand.Connection = Me.commonConn
            Me.sqlQuestionGroups.SelectCommand.CommandText = Me.sqlQuestionGroups.SelectCommand.CommandText.Replace("VARCONN.", myUtility.getExtension())
            Me.sqlQuestionGroups.Fill(Me.dsCommon, "QuestionGroups")
            Me.ddlQuestionGroup.DataSource = Me.dsCommon.Tables("QuestionGroups")
            Me.ddlQuestionGroup.DataTextField = Me.dsCommon.Tables("QuestionGroups").Columns("QUESTION_GROUP_TITLE").ToString
            Me.ddlQuestionGroup.DataValueField = Me.dsCommon.Tables("QuestionGroups").Columns("QUESTION_GROUP_ID").ToString
            Me.ddlQuestionGroup.DataBind()
            Dim li As System.Web.UI.WebControls.ListItem = New System.Web.UI.WebControls.ListItem
            li.Text = "-- Select a Question Group --"
            li.Value = "0"
            Me.ddlQuestionGroup.Items.Insert(0, li)

            If (Me.dsCommon.Tables("QuestionGroups").Rows.Count > 0) Then
                Session("QuestionGroupsExist") = True
            Else
                Session("QuestionGroupsExist") = False
            End If
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Session("QuestionGroupsExist") = False
            alert("Unable to load the question groups.")
            Return False
        End Try
    End Function

    'loads the query types for prepopulated fields
    Private Function loadQueries() As Boolean
        Trace.Warn("Loading Queryies")
        Try
            'load the queries
            Me.SqlSelectCommand6.CommandText = Me.SqlSelectCommand6.CommandText.Replace("VARCONN.", myUtility.getExtension())
            Trace.Warn(Me.SqlSelectCommand2.CommandText)

            'dump the old data before filling
            If (Me.dsCommon.Tables.Contains("Queries")) Then
                Me.dsCommon.Tables("Queries").Rows.Clear()
            End If

            'fill the data
            Me.sqlQueries.SelectCommand.Connection = Me.commonConn
            Me.sqlQueries.Fill(Me.dsCommon, "Queries")
            Me.ddlQuerySelect.DataSource = Me.dsCommon.Tables("Queries")
            Me.ddlQuerySelect.DataTextField = Me.dsCommon.Tables("Queries").Columns("QUERY_NAME").ToString
            Me.ddlQuerySelect.DataValueField = Me.dsCommon.Tables("Queries").Columns("QUERY_NAME").ToString
            Me.ddlQuerySelect.DataBind()
            Dim li As System.Web.UI.WebControls.ListItem = New System.Web.UI.WebControls.ListItem
            li.Text = "-- Select a Query --"
            li.Value = "*"
            Me.ddlQuerySelect.Items.Insert(0, li)

            If (Me.dsCommon.Tables("Queries").Rows.Count > 0) Then
                Session("QueriesExist") = True
            Else
                Session("QueriesExist") = False
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
        If (Session("intQuestion") > 0) Then
            Try
                Me.txtQuestion.Text = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_TEXT")
                Me.wneInitial.Value = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("RATING_INITIAL_VALUE")
                Me.wneInterval.Value = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("RATING_STEP_VALUE")
                Me.wneCount.Value = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("RATING_COUNT")
                Me.chkbxRequired.Checked = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("REQUIRED_IND")
                Me.chkbxPageBreak.Checked = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("PAGE_BREAK_IND")
                Me.chkbxPageBreak2.Checked = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("PAGE_BREAK_IND")
                Me.chkbxFilter.Checked = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("FILTER_IND")
                Me.chkTopIsNotApp.Checked = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("NOT_APPLICABLE_IND")
                If (Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("CHART_TYPE_CODE") <> "NONE") Then
                    Me.ddlGraphType.SelectedIndex = Me.ddlGraphType.Items.IndexOf(Me.ddlGraphType.Items.FindByValue(Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("CHART_TYPE_CODE")))
                Else
                    Me.ddlGraphType.SelectedIndex = 0
                End If
                If (Session("QuestionGroupsExist")) Then
                    Me.ddlQuestionGroup.SelectedIndex = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_GROUP_ID")
                End If
                Me.wneBranch.Value = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("BRANCH_QUESTION_ID")
                Me.lblNavBar.Text = "Question " & Session("QuestionRow") + 1 & " of " & Session("QuestionMax")
                If (Session("QueriesExist")) Then
                    Dim queryIndex As Integer = Me.ddlQuerySelect.Items.IndexOf(Me.ddlQuerySelect.Items.FindByValue(Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUERY_NAME")))
                    If (queryIndex >= 0) Then
                        Me.ddlQuerySelect.SelectedIndex = queryIndex
                    End If
                End If
            Catch ex As Exception
                Trace.Warn(ex.ToString)
            End Try
        Else
            Me.txtQuestion.Text = ""
            Me.wneInitial.Value = -1
            Me.wneInterval.Value = -1
            Me.wneCount.Value = -1
            Me.chkbxRequired.Checked = False
            Me.chkbxPageBreak.Checked = False
            Me.chkbxPageBreak2.Enabled = False
            Me.chkbxFilter.Checked = True
            Me.chkTopIsNotApp.Checked = False
            Me.ddlGraphType.SelectedIndex = 0
            If (Session("QuestionGroupsExist")) Then
                Me.ddlQuestionGroup.SelectedIndex = 0
            End If
            Me.wneBranch.Value = 0
            Me.lblNavBar.Text = "Question " & Session("QuestionRow") + 1 & " of " & Session("QuestionMax") + 1
            Me.ddlQuerySelect.SelectedIndex = 0
        End If

        'set the page options
        myUtility.optionPreSet(Session("intQuestion"), Session("QuestionMax"), Me.pageOptions)

        Session("PageOptions") = Me.pageOptions
    End Sub
#End Region

#Region "Page Switch"
    'Takes the user back to the template they were working on
    Private Sub btnTop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTop.Click
        Session("JSAction") = ""
        Session("CurrentPage") = "./Template.aspx"
        Session("pageSel") = "Template"
        Response.Redirect(Session("CurrentPage"))
    End Sub

    'brings up a pop-up window with the template preview in it
    Private Sub btnTemplatePreview_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTemplatePreview.Click
        Session("JSAction") = "<script>newWin = window.open('./Preview.aspx', 'PreviewWindow', 'width=600,height=400,scrollbars=yes,toolbar=no,titlebar=no,status=no,resizable=yes,menubar=no');newWin.focus();</script>"
    End Sub
#End Region

#Region "Settings"
    'sets the selected question type on the drop-down list
    Private Sub setQuestionType()
        Trace.Warn("Setting Question Type")

        'locate and select the current question type
        Dim item As ListItem
        Dim index As Integer = 0
        For Each item In Me.ddlQuestionType.Items
            If (Session("SelectedQuestionType") = item.Value) Then
                Me.ddlQuestionType.SelectedIndex = index
            Else
                index += 1
            End If
        Next
    End Sub

    'sets the selected chart type on the drop-down list
    Private Sub setChartType()
        Trace.Warn("Setting the chart type")

        'locate and select the current chart type
        Dim item As ListItem
        Dim index As Integer = 0
        For Each item In Me.ddlGraphType.Items
            If (Session("SelectedChartType") = item.Value) Then
                Me.ddlGraphType.SelectedIndex = index
            Else
                index += 1
            End If
        Next
    End Sub

    'sets the selected query on the drop-down list
    Private Sub setQuery()
        Trace.Warn("Setting the query")

        'locate and select the current query
        Dim item As ListItem
        Dim index As Integer = 0
        For Each item In Me.ddlQuerySelect.Items
            If (Session("SelectedQuery") = item.Value) Then
                Me.ddlQuerySelect.SelectedIndex = index
            Else
                index += 1
            End If
        Next
    End Sub

    'to determine what of the nav bar buttons should be available
    Private Sub setNavBar()
        'to determine what of the nav bar buttons should be available
        myNavBar.wnb_MoveTo(Me.navButtons, Session("QuestionMax"), Session("QuestionRow"), Switch)
    End Sub
#End Region

#Region "Nav Bar Events"
    'Moves the user to the first record
    Private Sub wnbQuestions_MoveToStart(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnStart.Click
        Session("JSAction") = ""

        If Not (loadQuestions()) Then
            alert("Failed to load the questions for this template.")
        Else
            Try
                Session("QuestionRow") = 0
                Session("intQuestion") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_ID")
                Session("QuestionName") = Me.dsCommon.Tables("Questions").Rows(0).Item("QUESTION_TEXT")
                If Not (loadListItems()) Then
                    alert("Failed to load the list items for the current question.")
                Else
                    Session("SelectedQuestionType") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_TYPE_CODE")
                    Session("SelectedChartType") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("CHART_TYPE_CODE")
                    Session("SelectedQuery") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUERY_NAME")
                    setQuestionType()
                    setChartType()
                    setQuery()
                    setNavBar()
                    setControls()
                    'set the page options
                    myUtility.optionPreSet(Session("intQuestion"), Session("QuestionMax"), Me.pageOptions)

                    Session("PageOptions") = Me.pageOptions
                End If
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the first question.")
            End Try
        End If
    End Sub

    'Moves the user to the previous record
    Private Sub wnbQuestions_MovePrev(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnPrev.Click
        Session("JSAction") = ""

        If Not (loadQuestions()) Then
            alert("Failed to load the questions for this template.")
        Else
            Try
                Session("QuestionRow") -= 1
                Session("intQuestion") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_ID")
                Session("QuestionName") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_TEXT")
                If Not (loadListItems()) Then
                    alert("Failed to load the list items for the current question.")
                Else
                    Session("SelectedQuestionType") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_TYPE_CODE")
                    Session("SelectedChartType") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("CHART_TYPE_CODE")
                    Session("SelectedQuery") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUERY_NAME")
                    setQuestionType()
                    setChartType()
                    setQuery()
                    setNavBar()
                    setControls()
                    'set the page options
                    myUtility.optionPreSet(Session("intQuestion"), Session("QuestionMax"), Me.pageOptions)

                    Session("PageOptions") = Me.pageOptions
                End If
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the previous question.")
            End Try
        End If
    End Sub

    'reloads the page and the data from the database while attempting to locate the user's location
    Private Sub wnbQuestions_MoveReload(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnReload.Click
        Session("JSAction") = ""

        If Not (loadData()) Then
            alert("Failed to load the questions for this template.")
        Else
            Try
                setNavBar()

                'reset the question type to nothing if making a new question
                If (Session("intQuestion") = -1) Then
                    Session("SelectedQuestionType") = "*"
                    Session("SelectedChartType") = "PIE"
                    Session("SelectedQuery") = "-- Select a Query --"
                    setQuestionType()
                    setChartType()
                    setQuery()
                End If

                setControls()
                'set the page options
                myUtility.optionPreSet(Session("intQuestion"), Session("QuestionMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error resetting the current question.")
            End Try
        End If
    End Sub

    'Moves the user to the next record
    Private Sub wnbQuestions_MoveNext(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnNext.Click
        Session("JSAction") = ""

        If Not (loadQuestions()) Then
            alert("Failed to load the questions for this template.")
        Else
            Try
                Session("QuestionRow") += 1
                Session("intQuestion") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_ID")
                Session("QuestionName") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_TEXT")
                If Not (loadListItems()) Then
                    alert("Failed to load the list items for the current question.")
                Else
                    Session("SelectedQuestionType") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_TYPE_CODE")
                    Session("SelectedChartType") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("CHART_TYPE_CODE")
                    Session("SelectedQuery") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUERY_NAME")
                    setQuestionType()
                    setChartType()
                    setQuery()
                    setNavBar()
                    setControls()
                    'set the page options
                    myUtility.optionPreSet(Session("intQuestion"), Session("QuestionMax"), Me.pageOptions)

                    Session("PageOptions") = Me.pageOptions
                End If
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the next question.")
            End Try
        End If
    End Sub

    'Moves the user to the last record
    Private Sub wnbQuestions_MoveToEnd(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnLast.Click
        Session("JSAction") = ""

        If Not (loadQuestions()) Then
            alert("Failed to load the questions for this template.")
        Else
            Try
                Session("QuestionRow") = Session("QuestionMax") - 1
                Session("intQuestion") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_ID")
                Session("QuestionName") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_TEXT")
                If Not (loadListItems()) Then
                    alert("Failed to load the list items for the current question.")
                Else
                    Session("SelectedQuestionType") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_TYPE_CODE")
                    Session("SelectedChartType") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("CHART_TYPE_CODE")
                    Session("SelectedQuery") = Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUERY_NAME")
                    setQuestionType()
                    setChartType()
                    setQuery()
                    setNavBar()
                    setControls()
                    'set the page options
                    myUtility.optionPreSet(Session("intQuestion"), Session("QuestionMax"), Me.pageOptions)

                    Session("PageOptions") = Me.pageOptions
                End If
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the last question.")
            End Try
        End If
    End Sub

    'Inserts a record into the table
    Private Sub wnbQuestions_InsertRow(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnNew.Click
        Session("JSAction") = ""

        If Not (loadData()) Then
            alert("Failed to load the questions for this template.")
        Else
            Try
                Session("intQuestion") = -1
                Session("QuestionRow") = Session("QuestionMax")
                Me.dsCommon.Tables("Questions").Rows.Add(Me.dsCommon.Tables("Questions").NewRow())
                Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("TEMPLATE_KEY") = Session("seqTemplate")
                Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("QUESTION_ID") = Session("QuestionMax")
                Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("REQUIRED_IND") = 0
                Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("FILTER_IND") = 1
                Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("NOT_APPLICABLE_IND") = 0
                Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("CHART_TYPE_CODE") = "PIE"
                Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("QUESTION_TYPE_CODE") = "*"
                Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("QUESTION_TEXT") = ""
                Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("RATING_INITIAL_VALUE") = -1
                Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("RATING_STEP_VALUE") = -1
                Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("RATING_COUNT") = -1
                Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("BRANCH_QUESTION_ID") = 0
                Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("QUESTION_GROUP_ID") = 0
                Session("QuestionName") = ""
                Session("SelectedQuestionType") = "*"
                Session("SelectedChartType") = "PIE"
                Session("SelectedQuery") = "-- Select a Query --"
                Session("ReportGroupID") = 0
                setQuestionType()
                setChartType()
                setQuery()
                myNavBar.wnb_MoveTo(Me.navButtons, Session("QuestionMax"), Session("QuestionRow"))
                setControls()
                'set the page options
                myUtility.optionPreSet(Session("intQuestion"), Session("QuestionMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error creating new question.")
            End Try
        End If
    End Sub
#End Region

#Region "Page Action"
    'handles changing the type of question
    Private Sub ddlQuestionType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddlQuestionType.SelectedIndexChanged
        Session("JSAction") = ""
        Session("OldQuestionType") = Session("SelectedQuestionType")
        Session("SelectedQuestionType") = Me.ddlQuestionType.SelectedItem.Value
        setControls()
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnInsert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsert.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'refresh the Questions table
        If (loadQuestions(True)) Then
            'refresh the tags table
            If Not (loadListItems("All")) Then
                blnContinue = False
                alert("Failed to load the questions for the current template. Insert aborted.")
            End If
        Else
            blnContinue = False
            alert("Failed to load the questions for the current template. Insert aborted.")
        End If

        'check for possible malicious code input
        If Not (myUtility.goodInput(Me.txtQuestion, True)) Then
            alert("Possible malicious code entry in the question text.  Insert aborted.")
            blnContinue = False
        End If

        If (blnContinue) Then
            'determine if the template is dirty or not
            Session("isDirty") = myUtility.isDirty(Me.dsCommon, Me.commonConn)

            'do the inert
            failure = doInsert()

            'determine if we need to reload the data
            If Not (loadData()) Then
                blnContinue = False
                alert("Failed to load the questions for the current template.")
            End If

            If (blnContinue) Then
                'if we succeeded then set the old question type to nothing
                If Not (failure) Then
                    Session("OldQuestionType") = ""
                End If

                'to determine what of the nav bar buttons should be available
                setNavBar()

                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
                Session("TransExists") = False
                'set the page options
                myUtility.optionPreSet(Session("intQuestion"), Session("QuestionMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'refresh the Questions table
        If (loadQuestions(True)) Then
            'refresh the tags table
            If Not (loadListItems("All")) Then
                blnContinue = False
                alert("Failed to load the questions for the current template. Update aborted.")
            End If
        Else
            blnContinue = False
            alert("Failed to load the questions for the current template. Update aborted.")
        End If

        'check for possible malicious code input
        If Not (myUtility.goodInput(Me.txtQuestion, True)) Then
            alert("Possible malicious code entry in the question text.  Update aborted.")
            blnContinue = False
        End If

        If (blnContinue) Then
            'determine if the template is dirty or not
            Session("isDirty") = myUtility.isDirty(Me.dsCommon, Me.commonConn)

            'do the update
            failure = doUpdate()

            'determine if we need to reload the data
            If Not (loadData()) Then
                blnContinue = False
                alert("Failed to load the questions for the current template.")
            End If

            If (blnContinue) Then
                'if we succeeded then set the old question type to nothing
                If Not (failure) Then
                    Session("OldQuestionType") = ""
                End If

                'to determine what of the nav bar buttons should be available
                setNavBar()

                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
                Session("TransExists") = False
                'set the page options
                myUtility.optionPreSet(Session("intQuestion"), Session("QuestionMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    'Reacts to the user performing some kind of selection and submission
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'refresh the Questions table
        If (loadQuestions(True)) Then
            'refresh the tags table
            If Not (loadListItems("All")) Then
                blnContinue = False
                alert("Failed to load the questions for the current template. Delete aborted.")
            End If
        Else
            blnContinue = False
            alert("Failed to load the questions for the current template. Delete aborted.")
        End If

        'check for possible malicious code input
        If Not (myUtility.goodInput(Me.txtQuestion, True)) Then
            alert("Possible malicious code entry in the question text.  Delete aborted.")
            blnContinue = False
        End If

        If (blnContinue) Then
            'determine if the template is dirty or not
            Session("isDirty") = myUtility.isDirty(Me.dsCommon, Me.commonConn)

            'do the delete
            failure = doDelete()

            'determine if we need to reload the data
            If Not (loadData()) Then
                blnContinue = False
                alert("Failed to load the questions for the current template.")
            End If

            If (blnContinue) Then
                'if we succeeded then set the old question type to nothing
                If Not (failure) Then
                    Session("OldQuestionType") = ""
                End If

                'to determine what of the nav bar buttons should be available
                setNavBar()

                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
                Session("TransExists") = False
                'set the page options
                myUtility.optionPreSet(Session("intQuestion"), Session("QuestionMax"), Me.pageOptions)

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
        If Not (myUtility.goodInput(Me.txtQuestion, True)) Then
            alert("Possible malicious code entry in the question text.  Reset aborted.")
            blnContinue = False
        End If

        If (blnContinue) Then
            'determine if the template is dirty or not
            Session("isDirty") = myUtility.isDirty(Me.dsCommon, Me.commonConn)

            'do the reset
            failure = doReset()

            If (blnContinue) Then
                'if we succeeded then set the old question type to nothing
                If Not (failure) Then
                    Session("OldQuestionType") = ""
                End If

                'to determine what of the nav bar buttons should be available
                setNavBar()

                'reset the page controls
                If Not (failure) Then
                    setControls()
                End If
                Session("TransExists") = False
                'set the page options
                myUtility.optionPreSet(Session("intQuestion"), Session("QuestionMax"), Me.pageOptions)

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
                If ((Session("SelectedQuestionType") = "H") Or (Me.txtQuestion.Text <> "")) Then
                    If (Session("intQuestion") = -1) Then
                        'If we're inserting a new record then add it onto the end of the recordset
                        If (insertQuestion(sqlCommonAction, sqlCommonAdapter, True)) Then
                            If (insertListItems(sqlCommonAction, sqlCommonAdapter, True)) Then
                                If Not (Session("isDirty")) Then
                                    Dim tempConn As Data.SqlClient.SqlConnection = New Data.SqlClient.SqlConnection
                                    If (myUtility.makeDirty(sqlCommonAction, sqlCommonAdapter, Me.dsCommon, myUtility.getConnection(tempConn))) Then
                                        If (Session("Alert") <> "") Then
                                            alert(Session("Alert"))
                                        End If
                                        Session("isDirty") = True
                                        Session("QuestionMax") += 1
                                        Session("intQuestion") = Session("QuestionMax")
                                        If (insertChartPref(sqlCommonAction, sqlCommonAdapter)) Then
                                            sqlCommonAction.Transaction.Commit()
                                        Else
                                            sqlCommonAction.Transaction.Rollback()
                                            alert("Error updating question in the template.")
                                            failure = True
                                        End If
                                    Else
                                        If (Session("Alert") <> "") Then
                                            alert(Session("Alert"))
                                        End If
                                        sqlCommonAction.Transaction.Rollback()
                                        alert("Error adding question to the template.")
                                        failure = True
                                    End If
                                Else
                                    Session("QuestionMax") += 1
                                    Session("intQuestion") = Session("QuestionMax")
                                    If (insertChartPref(sqlCommonAction, sqlCommonAdapter)) Then
                                        sqlCommonAction.Transaction.Commit()
                                    Else
                                        sqlCommonAction.Transaction.Rollback()
                                        alert("Error updating question in the template.")
                                        failure = True
                                    End If
                                End If
                            Else
                                sqlCommonAction.Transaction.Rollback()
                                alert("Error adding question to the template.")
                                failure = True
                            End If
                        Else
                            sqlCommonAction.Transaction.Rollback()
                            alert("Error adding question to the template.")
                            failure = True
                        End If
                    Else
                        'If we're inserting an existing record then inline insert it ito the record set and add it to the existing surveys
                        If (insertQuestion(sqlCommonAction, sqlCommonAdapter)) Then
                            If (insertListItems(sqlCommonAction, sqlCommonAdapter)) Then
                                If Not (Session("isDirty")) Then
                                    Dim tempConn As Data.SqlClient.SqlConnection = New Data.SqlClient.SqlConnection
                                    If (myUtility.makeDirty(sqlCommonAction, sqlCommonAdapter, Me.dsCommon, myUtility.getConnection(tempConn))) Then
                                        If (Session("Alert") <> "") Then
                                            alert(Session("Alert"))
                                        End If
                                        Session("QuestionMax") += 1
                                        If (insertChartPref(sqlCommonAction, sqlCommonAdapter)) Then
                                            sqlCommonAction.Transaction.Commit()
                                        Else
                                            sqlCommonAction.Transaction.Rollback()
                                            alert("Error updating question in the template.")
                                            failure = True
                                        End If
                                        Session("isDirty") = True
                                    Else
                                        If (Session("Alert") <> "") Then
                                            alert(Session("Alert"))
                                        End If
                                        sqlCommonAction.Transaction.Rollback()
                                        alert("Error adding question to the template.")
                                        failure = True
                                    End If
                                Else
                                    Session("QuestionMax") += 1
                                    If (insertChartPref(sqlCommonAction, sqlCommonAdapter)) Then
                                        sqlCommonAction.Transaction.Commit()
                                    Else
                                        sqlCommonAction.Transaction.Rollback()
                                        alert("Error updating question in the template.")
                                        failure = True
                                    End If
                                End If
                            Else
                                sqlCommonAction.Transaction.Rollback()
                                alert("Error adding question to the template.")
                                failure = True
                            End If
                        Else
                            sqlCommonAction.Transaction.Rollback()
                            alert("Error adding question to the template.")
                            failure = True
                        End If
                    End If
                Else
                    alert("You must select horizontal rule or at least have some text in the question box.")
                    failure = True
                    sqlCommonAction.Transaction.Rollback()
                End If
                Return failure
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Error adding question to the template.")
                Return True
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Error adding question to the template.")
            Return True
        End Try
    End Function

    'attempts to insert a Question, returns false if it cannot
    Private Function insertQuestion(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, Optional ByVal isNewRecord As Boolean = False) As Boolean
        Try
            'parameterized the text input to allow for things like quotes to get through
            Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@QUESTION_TEXT", System.Data.SqlDbType.VarChar, 1800, "QUESTION_TEXT")
            param0.Value = Me.txtQuestion.Text
            sqlCommonAction.Parameters.Add(param0)

            If (isNewRecord) Then
                'attempt to add the new record to the database

                sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION (TEMPLATE_KEY, QUESTION_ID, REQUIRED_IND, " & _
                "PAGE_BREAK_IND, FILTER_IND, NOT_APPLICABLE_IND, QUESTION_TYPE_CODE, QUESTION_TEXT, CHART_TYPE_CODE, RATING_COUNT, " & _
                "RATING_STEP_VALUE, RATING_INITIAL_VALUE, BRANCH_QUESTION_ID, QUESTION_GROUP_ID, QUERY_NAME) " & _
                "VALUES (" & Session("seqTemplate") & ", " & Session("QuestionMax") + 1

                'if dealing with a horizontal rule, comment, or hidden field then put a value of 0 in for blnRequired
                If (Session("SelectedQuestionType") <> "H" Or Session("SelectedQuestionType") <> "X" Or _
                    Session("SelectedQuestionType") <> "Q") Then
                    sqlCommonAction.CommandText &= ", " & IIf(Me.chkbxRequired.Checked, 1, 0) & _
                    ", "
                Else
                    sqlCommonAction.CommandText &= ", 0, "
                End If

                'use appropriate page break & filter checkboxes for data entry into the DB based on the question type
                If (Session("SelectedQuestionType") = "S" Or Session("SelectedQuestionType") = "C" Or _
                    Session("SelectedQuestionType") = "L" Or Session("SelectedQuestionType") = "P" Or _
                    Session("SelectedQuestionType") = "R" Or Session("SelectedQuestionType") = "M" Or _
                    Session("SelectedQuestionType") = "T" Or Session("SelectedQuestionType") = "D" Or _
                    Session("SelectedQuestionType") = "N") Then
                    sqlCommonAction.CommandText &= IIf(Me.chkbxPageBreak.Checked, 1, 0) & ", " & _
                        IIf(Me.chkbxFilter.Checked, 1, 0) & ", "
                Else
                    sqlCommonAction.CommandText &= IIf(Me.chkbxPageBreak2.Checked, 1, 0) & _
                        ", 0, "
                End If

                If (Session("SelectedQuestionType") = "R") Then
                    sqlCommonAction.CommandText &= IIf(Me.chkTopIsNotApp.Checked, 1, 0) & ", " & _
                        "'" & Session("SelectedQuestionType") & "'"
                Else
                    sqlCommonAction.CommandText &= "0, '" & Session("SelectedQuestionType") & "'"
                End If

                'if dealing with a horizontal rule use an empty string for the question
                If (Session("SelectedQuestionType") <> "H") Then
                    sqlCommonAction.CommandText &= ", @QUESTION_TEXT"
                Else
                    sqlCommonAction.CommandText &= ", '<HR>'"
                End If

                'use appropriate page break & filter checkboxes for data entry into the DB based on the question type
                If (Session("SelectedQuestionType") = "S" Or Session("SelectedQuestionType") = "C" Or _
                    Session("SelectedQuestionType") = "L" Or Session("SelectedQuestionType") = "P" Or _
                    Session("SelectedQuestionType") = "R" Or Session("SelectedQuestionType") = "M" Or _
                    Session("SelectedQuestionType") = "T" Or Session("SelectedQuestionType") = "D" Or _
                    Session("SelectedQuestionType") = "N") Then
                    sqlCommonAction.CommandText &= ", '" & Me.ddlGraphType.SelectedValue & "'"
                Else
                    sqlCommonAction.CommandText &= ", 'NONE'"
                End If

                'if dealing with a rating then use the input values otherwise put the dummy -1 values in for them
                If (Session("SelectedQuestionType") = "R") Then
                    sqlCommonAction.CommandText &= ", " & Me.wneCount.Value & ", " & Me.wneInterval.Value & _
                    ", " & Me.wneInitial.Value & ","
                Else
                    sqlCommonAction.CommandText &= ", -1, -1, -1,"
                End If

                sqlCommonAction.CommandText &= " " & Me.wneBranch.Value & ", " & _
                    Me.ddlQuestionGroup.SelectedValue & ", '" & Me.ddlQuerySelect.SelectedValue & "')"
                Trace.Warn(sqlCommonAction.CommandText)
                sqlCommonAction.ExecuteNonQuery()
                Return True
            Else
                'update the question id's
                Dim lowIndex As Integer = Session("QuestionRow")
                Dim highIndex As Integer = Me.dsCommon.Tables("Questions").Rows.Count() - 1

                'update the question id's of the Questions
                While (highIndex >= lowIndex)
                    sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION SET QUESTION_ID = (QUESTION_ID+1) " & _
                    "WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and QUESTION_ID = " & Me.dsCommon.Tables("Questions").Rows(highIndex).Item("QUESTION_ID")
                    sqlCommonAction.ExecuteNonQuery()
                    highIndex -= 1
                End While

                'attempt to add the new record to the database

                sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION (TEMPLATE_KEY, QUESTION_ID, REQUIRED_IND, " & _
                "PAGE_BREAK_IND, FILTER_IND, NOT_APPLICABLE_IND, QUESTION_TYPE_CODE, QUESTION_TEXT, CHART_TYPE_CODE, RATING_COUNT, " & _
                "RATING_STEP_VALUE, RATING_INITIAL_VALUE, BRANCH_QUESTION_ID, QUESTION_GROUP_ID, QUERY_NAME) VALUES (" & Session("seqTemplate") & ", " & Session("QuestionRow") + 1

                'if dealing with a horizontal rule, comment, or hidden field then put a value of 0 in for blnRequired
                If (Session("SelectedQuestionType") <> "H" Or Session("SelectedQuestionType") <> "X" Or _
                    Session("SelectedQuestionType") <> "Q") Then
                    sqlCommonAction.CommandText &= ", " & IIf(Me.chkbxRequired.Checked, 1, 0) & _
                    ", "
                Else
                    sqlCommonAction.CommandText &= ", 0, "
                End If

                'use appropriate page break checkbox for data entry into the DB based on the question type
                If (Session("SelectedQuestionType") = "S" Or Session("SelectedQuestionType") = "C" Or _
                    Session("SelectedQuestionType") = "L" Or Session("SelectedQuestionType") = "P" Or _
                    Session("SelectedQuestionType") = "R" Or Session("SelectedQuestionType") = "M" Or _
                    Session("SelectedQuestionType") = "T" Or Session("SelectedQuestionType") = "D" Or _
                    Session("SelectedQuestionType") = "N") Then
                    sqlCommonAction.CommandText &= IIf(Me.chkbxPageBreak.Checked, 1, 0) & ", " & _
                       IIf(Me.chkbxFilter.Checked, 1, 0) & ", "
                Else
                    sqlCommonAction.CommandText &= IIf(Me.chkbxPageBreak2.Checked, 1, 0) & _
                        ", 0, "
                End If

                If (Session("SelectedQuestionType") = "R") Then
                    sqlCommonAction.CommandText &= IIf(Me.chkTopIsNotApp.Checked, 1, 0) & ", " & _
                        "'" & Session("SelectedQuestionType") & "'"
                Else
                    sqlCommonAction.CommandText &= "0, '" & Session("SelectedQuestionType") & "'"
                End If

                'if dealing with a horizontal rule use an empty string for the question
                If (Session("SelectedQuestionType") <> "H") Then
                    sqlCommonAction.CommandText &= ", @QUESTION_TEXT"
                Else
                    sqlCommonAction.CommandText &= ", '<HR>'"
                End If

                'use appropriate page break & filter checkboxes for data entry into the DB based on the question type
                If (Session("SelectedQuestionType") = "S" Or Session("SelectedQuestionType") = "C" Or _
                    Session("SelectedQuestionType") = "L" Or Session("SelectedQuestionType") = "P" Or _
                    Session("SelectedQuestionType") = "R" Or Session("SelectedQuestionType") = "M" Or _
                    Session("SelectedQuestionType") = "T" Or Session("SelectedQuestionType") = "D" Or _
                    Session("SelectedQuestionType") = "N") Then
                    sqlCommonAction.CommandText &= ", '" & Me.ddlGraphType.SelectedValue & "'"
                Else
                    sqlCommonAction.CommandText &= ", 'NONE'"
                End If

                'if dealing with a rating then use the input values otherwise put the dummy -1 values in for them
                If (Session("SelectedQuestionType") = "R") Then
                    sqlCommonAction.CommandText &= ", " & Me.wneCount.Value & ", " & Me.wneInterval.Value & _
                    ", " & Me.wneInitial.Value & ","
                Else
                    sqlCommonAction.CommandText &= ", -1, -1, -1,"
                End If

                sqlCommonAction.CommandText &= " " & Me.wneBranch.Value & ", " & _
                    Me.ddlQuestionGroup.SelectedValue & ", '" & Me.ddlQuerySelect.SelectedValue & "')"

                Trace.Warn(sqlCommonAction.CommandText)
                sqlCommonAction.ExecuteNonQuery()
                Return True
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'attempts to insert the list items for the questions into the tables, returns false if it cannot
    Private Function insertListItems(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, Optional ByVal isNewRecord As Boolean = False) As Boolean
        If (isNewRecord) Then
            'no list items to update so skip this step on a new record and return true
            Return True
        Else
            'update the list items of questions after the one being inserted and if the question is also a list item
            'question duplicate the list items of the question it is being cloned from otherwise ignore any list
            'items belonging to the question it is being cloned from it won't need them.
            Dim currentQuestion As Integer = -1
            Dim questionIndex As Integer = Me.dsCommon.Tables("ListItems").Rows.Count() - 1

            'if there are no list items this will fail and we return true
            Try
                currentQuestion = Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("QUESTION_ID")
            Catch ex As Exception
                Return True
            End Try

            Try
                'declare the parameters outside so they are not reset
                Dim param1 As System.Data.SqlClient.SqlParameter
                Dim param2 As System.Data.SqlClient.SqlParameter

                While (questionIndex >= 0)
                    If (Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("QUESTION_ID") > Session("intQuestion")) Then
                        sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM SET QUESTION_ID = (QUESTION_ID+1) " & _
                        "WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and QUESTION_ID = " & _
                        Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("QUESTION_ID") & _
                        " and LIST_ITEM_ID = " & Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("LIST_ITEM_ID")
                        sqlCommonAction.ExecuteNonQuery()
                    ElseIf (Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("QUESTION_ID") < Session("intQuestion")) Then
                        'do nothing don't want to change this data
                        Exit While
                    Else
                        'update the current list item's question number to move it out of the way
                        sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM SET QUESTION_ID = (QUESTION_ID+1) " & _
                        "WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and QUESTION_ID = " & _
                        Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("QUESTION_ID") & _
                        " and LIST_ITEM_ID = " & Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("LIST_ITEM_ID")
                        sqlCommonAction.ExecuteNonQuery()

                        'if we have one of the four list types then clone the list item we just moved
                        If (Session("SelectedQuestionType") = "C" Or Session("SelectedQuestionType") = "L" Or _
                            Session("SelectedQuestionType") = "P" Or Session("SelectedQuestionType") = "T") Then

                            'parameterized the text input to allow for things like quotes to get through
                            sqlCommonAction.Parameters.Clear()
                            param1 = New System.Data.SqlClient.SqlParameter("@LIST_ITEM_TITLE", System.Data.SqlDbType.VarChar, 64, "LIST_ITEM_TITLE")
                            sqlCommonAction.Parameters.Add(param1)
                            param2 = New System.Data.SqlClient.SqlParameter("@LIST_ITEM_IMAGE", System.Data.SqlDbType.VarBinary, 2147483647, "LIST_ITEM_IMAGE")
                            sqlCommonAction.Parameters.Add(param2)

                            sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM (TEMPLATE_KEY, QUESTION_ID, LIST_ITEM_ID, LIST_ITEM_VALUE, LIST_ITEM_TITLE, " & _
                            "LIST_ITEM_IMAGE, LIST_ITEM_IMAGE_TYPE, BRANCH_QUESTION_ID) VALUES (" & Session("seqTemplate") & ", " & _
                            Session("intQuestion") & ", " & Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("LIST_ITEM_ID") & _
                            ", " & Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("LIST_ITEM_VALUE") & ", @LIST_ITEM_TITLE"

                            param1.Value = Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("LIST_ITEM_TITLE")
                            param2.Value = Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("LIST_ITEM_IMAGE")

                            'if we have a picture field and there are pictures to transfer do so otherwise just use Null
                            If ((Session("SelectedQuestionType") = Session("OldQuestionType") And Session("SelectedQuestionType") = "P") Or _
                               (Trim("" & Session("OldQuestionType")) = "" And Session("SelectedQuestionType") = "P")) Then
                                sqlCommonAction.CommandText &= ", @LIST_ITEM_IMAGE, '" & _
                                Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("LIST_ITEM_IMAGE_TYPE") & "',"
                            Else
                                sqlCommonAction.CommandText &= ", NULL, '',"
                            End If

                            sqlCommonAction.CommandText &= " " & Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("BRANCH_QUESTION_ID") & ")"

                            sqlCommonAction.ExecuteNonQuery()
                        End If
                    End If
                    questionIndex -= 1
                End While
                Return True
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                Return False
            End Try
        End If
    End Function

    'attempts to update the chart preferences for the users
    Private Function insertChartPref(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'adjust the question numbers up 1 for existing questions beyond the inserted question
            sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_CHART_PREFERENCE set QUESTION_ID = QUESTION_ID + 1 " & _
                "where TEMPLATE_KEY = " & Session("seqTemplate") & " and QUESTION_ID >= " & _
                Session("intQuestion")
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
                If ((Session("SelectedQuestionType") = "H") Or (Me.txtQuestion.Text <> "")) Then
                    If (updateQuestion(sqlCommonAction, sqlCommonAdapter)) Then
                        If ((Session("OldQuestionType") = "C" Or Session("OldQuestionType") = "L" Or _
                            Session("OldQuestionType") = "P" Or Session("OldQuestionType") = "T") And _
                            (Session("SelectedQuestionType") <> "C" And Session("SelectedQuestionType") <> "L" And _
                            Session("SelectedQuestionType") <> "P" And Session("SelectedQuestionType") <> "T")) Then
                            'ok we need to remove the list items since they are no longer needed because of a question
                            'type change from a list question to another type
                            If (updateListItems(sqlCommonAction, sqlCommonAdapter)) Then
                                If Not (Session("isDirty")) Then
                                    Dim tempConn As Data.SqlClient.SqlConnection = New Data.SqlClient.SqlConnection
                                    If (myUtility.makeDirty(sqlCommonAction, sqlCommonAdapter, Me.dsCommon, myUtility.getConnection(tempConn))) Then
                                        If (Session("Alert") <> "") Then
                                            alert(Session("Alert"))
                                        End If
                                        Session("isDirty") = True
                                        If (updateChartPref(sqlCommonAction, sqlCommonAdapter)) Then
                                            sqlCommonAction.Transaction.Commit()
                                        Else
                                            sqlCommonAction.Transaction.Rollback()
                                            alert("Error updating question in the template.")
                                            failure = True
                                        End If
                                    Else
                                        If (Session("Alert") <> "") Then
                                            alert(Session("Alert"))
                                        End If
                                        sqlCommonAction.Transaction.Rollback()
                                        alert("Error updating question in the template.")
                                        failure = True
                                    End If
                                Else
                                    If (updateChartPref(sqlCommonAction, sqlCommonAdapter)) Then
                                        sqlCommonAction.Transaction.Commit()
                                    Else
                                        sqlCommonAction.Transaction.Rollback()
                                        alert("Error updating question in the template.")
                                        failure = True
                                    End If
                                End If
                            Else
                                sqlCommonAction.Transaction.Rollback()
                                alert("Error updating question in the template.")
                                failure = True
                            End If
                        ElseIf (Session("OldQuestionType") = "P" And Session("SelectedQuestionType") <> "P") Then
                            'ok we just need to dump the binary data taking up space in the database
                            If (updateListItems(sqlCommonAction, sqlCommonAdapter, True)) Then
                                If Not (Session("isDirty")) Then
                                    Dim tempConn As Data.SqlClient.SqlConnection = New Data.SqlClient.SqlConnection
                                    If (myUtility.makeDirty(sqlCommonAction, sqlCommonAdapter, Me.dsCommon, myUtility.getConnection(tempConn))) Then
                                        If (Session("Alert") <> "") Then
                                            alert(Session("Alert"))
                                        End If
                                        Session("isDirty") = True
                                        If (updateChartPref(sqlCommonAction, sqlCommonAdapter)) Then
                                            sqlCommonAction.Transaction.Commit()
                                        Else
                                            sqlCommonAction.Transaction.Rollback()
                                            alert("Error updating question in the template.")
                                            failure = True
                                        End If
                                    Else
                                        If (Session("Alert") <> "") Then
                                            alert(Session("Alert"))
                                        End If
                                        sqlCommonAction.Transaction.Rollback()
                                        alert("Error updating question in the template.")
                                        failure = True
                                    End If
                                Else
                                    If (updateChartPref(sqlCommonAction, sqlCommonAdapter)) Then
                                        sqlCommonAction.Transaction.Commit()
                                    Else
                                        sqlCommonAction.Transaction.Rollback()
                                        alert("Error updating question in the template.")
                                        failure = True
                                    End If
                                End If
                            Else
                                sqlCommonAction.Transaction.Rollback()
                                alert("Error updating question in the template.")
                                failure = True
                            End If
                        Else
                            If Not (Session("isDirty")) Then
                                Dim tempConn As Data.SqlClient.SqlConnection = New Data.SqlClient.SqlConnection
                                If (myUtility.makeDirty(sqlCommonAction, sqlCommonAdapter, Me.dsCommon, myUtility.getConnection(tempConn))) Then
                                    If (Session("Alert") <> "") Then
                                        alert(Session("Alert"))
                                    End If
                                    Session("isDirty") = True
                                    If (updateChartPref(sqlCommonAction, sqlCommonAdapter)) Then
                                        sqlCommonAction.Transaction.Commit()
                                    Else
                                        sqlCommonAction.Transaction.Rollback()
                                        alert("Error updating question in the template.")
                                        failure = True
                                    End If
                                Else
                                    If (Session("Alert") <> "") Then
                                        alert(Session("Alert"))
                                    End If
                                    sqlCommonAction.Transaction.Rollback()
                                    alert("Error updating question in the template.")
                                    failure = True
                                End If
                            Else
                                If (updateChartPref(sqlCommonAction, sqlCommonAdapter)) Then
                                    sqlCommonAction.Transaction.Commit()
                                Else
                                    sqlCommonAction.Transaction.Rollback()
                                    alert("Error updating question in the template.")
                                    failure = True
                                End If
                            End If
                        End If
                    Else
                        sqlCommonAction.Transaction.Rollback()
                        alert("Error updating question in the template.")
                        failure = True
                    End If
                Else
                    sqlCommonAction.Transaction.Rollback()
                    alert("You must select horizontal rule or at least have some text in the question box.")
                    failure = True
                End If
                Return failure
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Error updating question in the template.")
                Return True
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Error updating question in the template.")
            Return True
        End Try
    End Function

    'attempts to update a Question, returns false if it cannot
    Private Function updateQuestion(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'parameterized the text input to allow for things like quotes to get through
            Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@QUESTION_TEXT", System.Data.SqlDbType.VarChar, 1800, "QUESTION_TEXT")
            param0.Value = Me.txtQuestion.Text
            sqlCommonAction.Parameters.Add(param0)

            'attempt to update the record in the database
            sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION SET "

            'if dealing with a horizontal rule, comment, or hidden field then put a value of 0 in for it
            If (Session("SelectedQuestionType") <> "H" And Session("SelectedQuestionType") <> "X" And _
                Session("SelectedQuestionType") <> "Q") Then
                sqlCommonAction.CommandText &= "REQUIRED_IND = " & IIf(Me.chkbxRequired.Checked, 1, 0) & _
                ", PAGE_BREAK_IND = " & IIf(Me.chkbxPageBreak.Checked, 1, 0) & _
                ", QUESTION_TYPE_CODE = '" & Session("SelectedQuestionType") & "'" & _
                ", FILTER_IND = " & IIf(Me.chkbxFilter.Checked, 1, 0) & _
                ", CHART_TYPE_CODE = '" & Me.ddlGraphType.SelectedValue & "'"
            Else
                sqlCommonAction.CommandText &= "REQUIRED_IND = 0, PAGE_BREAK_IND = " & _
                IIf(Me.chkbxPageBreak2.Checked, 1, 0) & _
                ", QUESTION_TYPE_CODE = '" & Session("SelectedQuestionType") & "'" & _
                ", FILTER_IND = 0, CHART_TYPE_CODE = 'NONE'"
            End If

            If (Session("SelectedQuestionType") = "R") Then
                sqlCommonAction.CommandText &= ", NOT_APPLICABLE_IND = " & IIf(Me.chkTopIsNotApp.Checked, 1, 0)
            Else
                sqlCommonAction.CommandText &= ", NOT_APPLICABLE_IND = 0"
            End If

            'if dealing with a horizontal rule use an empty string for the question
            If (Session("SelectedQuestionType") <> "H") Then
                sqlCommonAction.CommandText &= ", QUESTION_TEXT = @QUESTION_TEXT"
            Else
                sqlCommonAction.CommandText &= ", QUESTION_TEXT = '<HR>'"
            End If

            'if dealing with a rating then use the input values otherwise put the dummy -1 values in for them
            If (Session("SelectedQuestionType") = "R") Then
                sqlCommonAction.CommandText &= ", RATING_INITIAL_VALUE = " & Me.wneInitial.Value & ", RATING_STEP_VALUE = " & _
                Me.wneInterval.Value & ", RATING_COUNT = " & Me.wneCount.Value
            Else
                sqlCommonAction.CommandText &= ", RATING_INITIAL_VALUE = -1, RATING_STEP_VALUE = -1, RATING_COUNT = -1"
            End If

            sqlCommonAction.CommandText &= ", BRANCH_QUESTION_ID = " & Me.wneBranch.Value & _
                ", QUESTION_GROUP_ID = " & Me.ddlQuestionGroup.SelectedValue & _
                ", QUERY_NAME = '" & Me.ddlQuerySelect.SelectedValue & "'"

            sqlCommonAction.CommandText &= " WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and QUESTION_ID = " & _
Session("intQuestion")
            Trace.Warn(sqlCommonAction.CommandText)
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'attempts to update the list items in the list items table, returns false if it cannot
    Private Function updateListItems(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, Optional ByVal isRemovePictures As Boolean = False) As Boolean
        Try
            'if we are here to remove list items then do the delete otherwise we're here to remove the binary data
            If Not (isRemovePictures) Then
                'attempt to delete the records for the current question from the database
                sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM where TEMPLATE_KEY = " & Session("seqTemplate") & _
                    " and QUESTION_ID = " & Session("intQuestion")
                sqlCommonAction.ExecuteNonQuery()
            Else
                'attempt to remove the binary data from the records for the current question from the database
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM SET LIST_ITEM_IMAGE = NULL, LIST_ITEM_IMAGE_TYPE = ' ' " & _
                "WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and QUESTION_ID = " & Session("intQuestion")
                sqlCommonAction.ExecuteNonQuery()
            End If
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'attempts to update the chart preferences for the users
    Private Function updateChartPref(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'remove the chart preference if the item is no longer a chartable item
            If (Me.ddlQuestionType.SelectedItem.Text <> "C" And Me.ddlQuestionType.SelectedItem.Text <> "D" _
                Or Me.ddlQuestionType.SelectedItem.Text <> "L" And Me.ddlQuestionType.SelectedItem.Text <> "N" _
                Or Me.ddlQuestionType.SelectedItem.Text <> "P" And Me.ddlQuestionType.SelectedItem.Text <> "R" _
                Or Me.ddlQuestionType.SelectedItem.Text <> "T") Then
                sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_CHART_PREFERENCE where TEMPLATE_KEY = " & _
                    Session("seqTemplate") & " and QUESTION_ID = " & Session("intQuestion")
                sqlCommonAction.ExecuteNonQuery()
            End If
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
                If (deleteQuestion(sqlCommonAction, sqlCommonAdapter)) Then
                    If (deleteListItems(sqlCommonAction, sqlCommonAdapter)) Then
                        If Not (Session("isDirty")) Then
                            Dim tempConn As Data.SqlClient.SqlConnection = New Data.SqlClient.SqlConnection
                            If (myUtility.makeDirty(sqlCommonAction, sqlCommonAdapter, Me.dsCommon, myUtility.getConnection(tempConn))) Then
                                If (Session("Alert") <> "") Then
                                    alert(Session("Alert"))
                                End If
                                Session("isDirty") = True
                                If (deleteChartPref(sqlCommonAction, sqlCommonAdapter)) Then
                                    sqlCommonAction.Transaction.Commit()
                                Else
                                    sqlCommonAction.Transaction.Rollback()
                                    alert("Error updating question in the template.")
                                    failure = True
                                End If
                                If (Session("QuestionMax") = 1) Then
                                    Session("QuestionRow") = 0
                                    Session("intQuestion") = -1
                                    Session("QuestionMax") = 0
                                    Me.dsCommon.Tables("Questions").Rows.Add(Me.dsCommon.Tables("Questions").NewRow())
                                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("TEMPLATE_KEY") = Session("seqTemplate")
                                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("QUESTION_ID") = Session("QuestionMax")
                                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("REQUIRED_IND") = 0
                                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("FILTER_IND") = 1
                                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("NOT_APPLICABLE_IND") = 0
                                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("CHART_TYPE_CODE") = "PIE"
                                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_TYPE_CODE") = "*"
                                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_TEXT") = ""
                                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("RATING_INITIAL_VALUE") = -1
                                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("RATING_STEP_VALUE") = -1
                                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("RATING_COUNT") = -1
                                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("BRANCH_QUESTION_ID") = 0
                                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_GROUP_ID") = 0
                                    Session("QuestionName") = ""
                                    Session("SelectedQuestionType") = "*"
                                    Session("SelectedChartType") = "PIE"
                                    Session("ReportGroupID") = 0
                                    setQuestionType()
                                    setChartType()
                                ElseIf (Session("intQuestion") = Session("QuestionMax") And Session("QuestionMax") - Session("QuestionRow") = 1) Then
                                    Session("intQuestion") -= 1
                                    Session("QuestionMax") -= 1
                                    Session("QuestionRow") -= 1
                                Else
                                    Session("QuestionMax") -= 1
                                End If
                            Else
                                If (Session("Alert") <> "") Then
                                    alert(Session("Alert"))
                                End If
                                sqlCommonAction.Transaction.Rollback()
                                alert("Error deleting question in the template.")
                                failure = True
                            End If
                        Else
                            If (deleteChartPref(sqlCommonAction, sqlCommonAdapter)) Then
                                sqlCommonAction.Transaction.Commit()
                            Else
                                sqlCommonAction.Transaction.Rollback()
                                alert("Error updating question in the template.")
                                failure = True
                            End If
                            If (Session("QuestionMax") = 1) Then
                                Session("QuestionRow") = 0
                                Session("intQuestion") = -1
                                Session("QuestionMax") = 0
                                Me.dsCommon.Tables("Questions").Rows.Add(Me.dsCommon.Tables("Questions").NewRow())
                                Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("TEMPLATE_KEY") = Session("seqTemplate")
                                Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_ID") = Session("QuestionMax")
                                Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("REQUIRED_IND") = 0
                                Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("FILTER_IND") = 1
                                Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("NOT_APPLICABLE_IND") = 0
                                Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("CHART_TYPE_CODE") = "PIE"
                                Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_TYPE_CODE") = "*"
                                Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_TEXT") = ""
                                Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("RATING_INITIAL_VALUE") = -1
                                Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("RATING_STEP_VALUE") = -1
                                Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("RATING_COUNT") = -1
                                Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("BRANCH_QUESTION_ID") = 0
                                Me.dsCommon.Tables("Questions").Rows(Session("QuestionRow")).Item("QUESTION_GROUP_ID") = 0
                                Session("QuestionName") = ""
                                Session("SelectedQuestionType") = "*"
                                Session("SelectedChartType") = "PIE"
                                Session("ReportGroupID") = 0
                                setQuestionType()
                                setChartType()
                            ElseIf (Session("intQuestion") = Session("QuestionMax") And Session("QuestionMax") - Session("QuestionRow") = 1) Then
                                Session("intQuestion") -= 1
                                Session("QuestionMax") -= 1
                                Session("QuestionRow") -= 1
                            Else
                                Session("QuestionMax") -= 1
                            End If
                        End If
                    Else
                        sqlCommonAction.Transaction.Rollback()
                        alert("Error deleting question in the template.")
                        failure = True
                    End If
                Else
                    sqlCommonAction.Transaction.Rollback()
                    alert("Error deleting question in the template.")
                    failure = True
                End If
                Return failure
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Error deleting question from template.")
                Return True
            End If
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Error deleting question from template.")
            Return True
        End Try
    End Function

    'attempts to delete a Question, returns false if it cannot
    Private Function deleteQuestion(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'attempt to delete the record from the database
            sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION where TEMPLATE_KEY = " & Session("seqTemplate") & _
            " and QUESTION_ID = " & Session("intQuestion")
            sqlCommonAction.ExecuteNonQuery()

            'update the question id's
            Dim lowIndex As Integer = Session("QuestionRow")
            Dim highIndex As Integer = Me.dsCommon.Tables("Questions").Rows.Count()

            'update the ID's of the Questions
            While (lowIndex < highIndex)
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION SET QUESTION_ID = (QUESTION_ID-1) " & _
                "WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and QUESTION_ID = " & _
                Me.dsCommon.Tables("Questions").Rows(lowIndex).Item("QUESTION_ID")
                sqlCommonAction.ExecuteNonQuery()
                lowIndex += 1
            End While
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'attempts to delete the list items in the list items table, returns false if it cannot
    Private Function deleteListItems(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        'delete the list items from each survey before updating the tag IDs for each survey
        Dim currentQuestion As Integer = -1
        Dim questionIndex As Integer = 0

        'if there are no list items then this will fail and we'll return true
        Try
            currentQuestion = Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("QUESTION_ID")
        Catch ex As Exception
            Return True
        End Try

        Try
            'process the existing list items
            While (questionIndex < Me.dsCommon.Tables("ListItems").Rows.Count())
                If (Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("QUESTION_ID") = Session("intQuestion")) Then
                    'attempt to delete the record from the database
                    sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM where TEMPLATE_KEY = " & Session("seqTemplate") & _
                    " and QUESTION_ID = " & Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("QUESTION_ID") & _
                    " and LIST_ITEM_ID = " & Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("LIST_ITEM_ID")
                    sqlCommonAction.ExecuteNonQuery()

                    'only update the tag item after if one actually exists and it happens to be larger than the question ID we're 
                    'getting rid of
                    If (questionIndex < Me.dsCommon.Tables("ListItems").Rows.Count()) Then
                        If (Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("QUESTION_ID") > Session("intQuestion")) Then
                            sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM SET QUESTION_ID = (QUESTION_ID-1) " & _
                            "WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and QUESTION_ID = " & _
                            Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("QUESTION_ID") & " and LIST_ITEM_ID = " & _
                            Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("LIST_ITEM_ID")
                            sqlCommonAction.ExecuteNonQuery()
                        End If
                    End If
                ElseIf (Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("QUESTION_ID") > Session("intQuestion")) Then
                    'reduce the question numbers by 1 if the question number is greater than the question number in session
                    sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM SET QUESTION_ID = (QUESTION_ID-1) " & _
                    "WHERE TEMPLATE_KEY = " & Session("seqTemplate") & " and QUESTION_ID = " & _
                    Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("QUESTION_ID") & " and LIST_ITEM_ID = " & _
                    Me.dsCommon.Tables("ListItems").Rows(questionIndex).Item("LIST_ITEM_ID")
                    sqlCommonAction.ExecuteNonQuery()
                End If
                questionIndex += 1
            End While
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'attempts to update the chart preferences for the users
    Private Function deleteChartPref(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        Try
            'remove the chart preference if the item is no longer a chartable item
            Dim tempConn As Data.SqlClient.SqlConnection = New Data.SqlClient.SqlConnection
            Dim tempCommand As Data.SqlClient.SqlCommand = New Data.SqlClient.SqlCommand
            sqlCommonAction.CommandText = "Delete from " & myUtility.getProdExtension() & "fempReportPref where seqTemplate = " & _
                Session("seqTemplate") & " and intQuestion = " & Session("intQuestion")
            sqlCommonAction.ExecuteNonQuery()

            'adjust the question numbers down 1 for existing questions beyond the deleted question
            sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_CHART_PREFERENCE set QUESTION_ID = QUESTION_ID - 1 " & _
                "where TEMPLATE_KEY = " & Session("seqTemplate") & " and QUESTION_ID > " & _
                Session("intQuestion")
            sqlCommonAction.ExecuteNonQuery()
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
            alert("Failed to load the questions for this template.")
        Else
            Try
                If (Session("intQuestion") = -1) Then
                    Me.dsCommon.Tables("Questions").Rows.Add(Me.dsCommon.Tables("Questions").NewRow())
                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("TEMPLATE_KEY") = Session("seqTemplate")
                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("QUESTION_ID") = Session("QuestionMax")
                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("REQUIRED_IND") = 0
                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("FILTER_IND") = 1
                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("NOT_APPLICABLE_IND") = 0
                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("CHART_TYPE_CODE") = "PIE"
                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("QUESTION_TYPE_CODE") = "*"
                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("QUESTION_TEXT") = ""
                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("RATING_INITIAL_VALUE") = -1
                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("RATING_STEP_VALUE") = -1
                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("RATING_COUNT") = -1
                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("BRANCH_QUESTION_ID") = 0
                    Me.dsCommon.Tables("Questions").Rows(Session("QuestionMax")).Item("QUESTION_GROUP_ID") = 0
                    Session("QuestionName") = ""
                    Session("SelectedQuestionType") = "*"
                    Session("SelectedChartType") = "PIE"
                    Session("SelectedQuery") = "-- Select a Query --"
                    Session("ReportGroupID") = 0
                    setQuestionType()
                    setChartType()
                    setQuery()
                End If
                Return False
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error resetting the question.")
                Return True
            End Try
        End If
    End Function
#End Region
End Class
