Imports System.text
Imports System.IO
Imports System.Collections.Specialized
Public Class Preview
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Preview))
        Me.sqlQuestions = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand
        Me.dsCommon = New System.Data.DataSet
        Me.commonConn = New System.Data.SqlClient.SqlConnection
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
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents lblNoQuestions As System.Web.UI.WebControls.Label
    Protected WithEvents lblDisclaimer As System.Web.UI.WebControls.Label
    Protected WithEvents lblQuestions As System.Web.UI.WebControls.Label
    Protected WithEvents QuestionTable As System.Web.UI.WebControls.Table
    Protected WithEvents sqlQuestions As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents dsCommon As System.Data.DataSet
    Protected WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents btnRefresh As System.Web.UI.WebControls.Button
    Protected myUtility As Utility = New Utility
    Protected WithEvents btnAccept As System.Web.UI.WebControls.Button
    Protected WithEvents btnReject As System.Web.UI.WebControls.Button
    Protected WithEvents imgPreviewTitle As System.Web.UI.WebControls.Image
    Protected WithEvents lblQuestion As System.Web.UI.WebControls.Label
    Protected WithEvents lblRequired As System.Web.UI.WebControls.Label
    Protected WithEvents hlpComments As System.Web.UI.WebControls.Image
    Protected WithEvents hlpHSR As System.Web.UI.WebControls.Image
    Protected WithEvents chkbxHSR As System.Web.UI.WebControls.CheckBox
    Protected WithEvents txtComments As System.Web.UI.WebControls.TextBox
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

        'catch incoming alert messages
        alert(Session("Alert"))

        'clear session
        myUtility.clearSession()

        'if the user's password has expired send them to the password reset page
        If (Session("isExpired") = "True") Then
            Session("CurrentPage") = "./Password.aspx"
            Session("pageSel") = "Password"
            redirect(Session("CurrentPage"))
        End If

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

        If (CType(Session("isADCHSR"), Boolean) = False) Then
            'determine if a level 1 user is attempting to sign off
            If (isSigner()) Then
                Session("isADCHSR") = "True"
            Else
                Session("isADCHSR") = "False"
            End If

            'Put user code to initialize the page here
            If (Session("isAuthenticated") <> "True" And Session("UserType") <> 1) Then
                redirect("./Login.aspx?h=1")
            ElseIf ((myUtility.checkUserType(1, CType(Session("isTemplateOwner"), Boolean), CType(Session("isTemplateDelegate"), Boolean)) And Not CType(Session("isTemplateUser"), Boolean)) And (Session("UserType") <> 1 And Session("UserType") <> 4)) Then
                Response.Write("<script type='text/javascript'>window.opener.location.href = window.opener.location.href;</script>")
                Response.Write("<script type='text/javascript'>window.close()</script>")
            Else
                Session("Alert") = ""
            End If
        End If

        'Set the connection based on the machine
        myUtility.getConnection(Me.commonConn)

        'finalize the SQL statements
        completeSQL()

        'Put user code to initialize the page here
        If (Session("isADCHSR") <> "True") Then
            If (Session("isAuthenticated") <> "True") Then
                Session("CurrentPage") = "./Login.aspx"
                Session("pageSel") = "Login"
                Response.Write("<script type='text/javascript'>window.opener.location.href = window.opener.location.href;</script>")
                Response.Write("<script type='text/javascript'>window.close()</script>")
            Else
                loadData()
            End If
        ElseIf Not (Session("TransID") Is Nothing) Then
            'store the incoming data in the proper signature record
            If (storeSignature()) Then
                'determine if the survey template is completely signed off
                If (isSignedOff()) Then
                    'set the dirty flag to clean
                    If Not (resetDirtyFlag()) Then
                        alert("Unable to mark the survey template for production copy.  Please contact the Survey Administrator so that this survey template can be published.")
                    End If

                    'send an e-mail the user and the survey administrator about the acceptance
                    If Not (sendReplyEmail(Session("Comments"), False, True)) Then
                        alert("Failed to notify the user.  Please send the user a courtesy e-mail.")
                    Else
                        'close this window since we stored the signature
                        Session("TransID") = Nothing
                        Response.Write("<script type='text/javascript'>window.close()</script>")
                    End If
                Else
                    'send an e-mail the user and the survey administrator about the acceptance
                    If Not (sendReplyEmail(Session("Comments"))) Then
                        alert("Failed to notify the user.  Please send the user a courtesy e-mail.")
                    Else
                        'close this window since we stored the signature
                        Session("TransID") = Nothing
                        Response.Write("<script type='text/javascript'>window.close()</script>")
                    End If
                End If
            End If

            Session("TransID") = Nothing
            loadData()

            If (Not IsPostBack) Then
                'get the user's comments and check item if available
                populateFields()
            End If
        Else
            loadData()

            If (Not IsPostBack) Then
                'get the user's comments and check item if available
                populateFields()
            End If
        End If

        Me.btnAccept.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
        Me.btnReject.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")

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

    'handles error messages
    Public Sub alert(ByVal message As String)
        If (message <> "") Then
            JSAction = "<script type='text/javascript'>window.alert('" & message & "');</script>"
            Session("Alert") = ""
        End If
    End Sub

    'completes the sql statements
    Private Sub completeSQL()
        Me.SqlSelectCommand1.CommandText &= " WHERE fq.TEMPLATE_KEY = " & Session("seqTemplate") & _
            " ORDER BY fq.QUESTION_ID, fl.LIST_ITEM_ID"
    End Sub
#End Region

#Region "Data Load"
    'Initializes the question data on the page
    Private Sub loadData()
        Trace.Warn("Entering InitData")

        'clear out the old data and refill from the DB
        If (Me.QuestionTable.Rows.Count() > 0) Then
            Me.QuestionTable.Rows.Clear()
        End If

        Trace.Warn(Me.SqlSelectCommand1.CommandText)
        Me.SqlSelectCommand1.Connection = Me.commonConn
        Me.SqlSelectCommand1.CommandText = Me.SqlSelectCommand1.CommandText.Replace("VARCONN.", myUtility.getExtension())
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

                'if there is a page break simulate it for the user to see
                If (dt.Rows(index).Item("PAGE_BREAK_IND")) Then
                    row.Cells.Clear()

                    If (Session("Results") = 0 And ((CType(Session("isTemplateOwner"), Boolean) Or CType(Session("isTemplateDelegate"), Boolean)) Or (Session("UserType") = 4 And Session("isADCHSR") <> "True"))) Then
                        'generate cell 40 pixels wide
                        cell = New System.Web.UI.WebControls.TableCell
                        cell.Wrap = False
                        cell.Width = Unit.Pixel(40)
                        cell.EnableViewState = True
                        cell.BorderStyle = BorderStyle.None
                        cell.VerticalAlign = VerticalAlign.Top
                        row.Cells.Add(cell)
                    End If

                    'generate the pseudo page break
                    cell = New System.Web.UI.WebControls.TableCell
                    Dim lblPageBreak As Label = New Label
                    lblPageBreak.Text = "--------------Page Break--------------"
                    lblPageBreak.ForeColor = System.Drawing.Color.FromArgb(0, 153, 0, 0)
                    lblPageBreak.CssClass = "boldContentRed"
                    lblPageBreak.ID = "lblPageBreak" & dt.Rows(index).Item("QUESTION_ID")
                    cell.Controls.Add(lblPageBreak)
                    cell.HorizontalAlign = HorizontalAlign.Center
                    row.Cells.Add(cell)
                    QuestionTable.Rows.Add(row)
                    row = New System.Web.UI.WebControls.TableRow
                End If

                'process the question
                If (strQuestionType = "C") Then
                    'Choice Question
                    row.Cells.Clear()

                    'generate the link to the question
                    If (Session("Results") = 0 And ((CType(Session("isTemplateOwner"), Boolean) Or CType(Session("isTemplateDelegate"), Boolean)) Or (Session("UserType") = 4 And Session("isADCHSR") <> "True"))) Then
                        doAddQuestionLink(cell, row, index, dt)
                    End If

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
                        cell.Controls.Add(chkbxList)
                        cell.Wrap = True
                        cell.EnableViewState = True
                        cell.BorderStyle = BorderStyle.None
                        row.Cells.Add(cell)
                    End If

                    'increment the visible question number
                    visibleQuestionNumber += 1
                ElseIf (strQuestionType = "D") Then
                    'DateField Question
                    row.Cells.Clear()

                    'generate the link to the question
                    If (Session("Results") = 0 And ((CType(Session("isTemplateOwner"), Boolean) Or CType(Session("isTemplateDelegate"), Boolean)) Or (Session("UserType") = 4 And Session("isADCHSR") <> "True"))) Then
                        doAddQuestionLink(cell, row, index, dt)
                    End If

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
                    datField.Value = Now
                    datField.CssClass = "Dyanmic_Controsl_Font"
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
                    cell.Controls.Add(datField)
                    cell.Wrap = False
                    cell.EnableViewState = True
                    cell.BorderStyle = BorderStyle.None
                    row.Cells.Add(cell)

                    'increment the visible question number
                    visibleQuestionNumber += 1
                ElseIf (strQuestionType = "H") Then
                    'Horizontal Rule Question
                    row.Cells.Clear()

                    'generate the link to the question
                    If (Session("Results") = 0 And ((CType(Session("isTemplateOwner"), Boolean) Or CType(Session("isTemplateDelegate"), Boolean)) Or (Session("UserType") = 4 And Session("isADCHSR") <> "True"))) Then
                        doAddQuestionLink(cell, row, index, dt)
                    End If

                    'generate horizontal rule
                    cell = New System.Web.UI.WebControls.TableCell
                    cell.HorizontalAlign = HorizontalAlign.Left
                    Dim lblQuestion As Label = New Label
                    lblQuestion.Text = "<hr style=""BORDER-LEFT-COLOR: #000099; BORDER-BOTTOM-COLOR: #000099; BORDER-TOP-STYLE: solid; " & _
                        "BORDER-TOP-COLOR: #000099; BORDER-RIGHT-STYLE: solid; BORDER-LEFT-STYLE: solid; BORDER-RIGHT-COLOR: #000099; " & _
                        "BORDER-BOTTOM-STYLE: solid"" size=3>"
                    lblQuestion.CssClass = "boldContentTeal"
                    lblQuestion.ID = "HR" & dt.Rows(index).Item("QUESTION_ID")
                    cell.Controls.Add(lblQuestion)
                    cell.Wrap = True
                    cell.EnableViewState = True
                    cell.BorderStyle = BorderStyle.None
                    row.Cells.Add(cell)
                ElseIf (strQuestionType = "L") Then
                    'List Question
                    row.Cells.Clear()

                    'generate the link to the question
                    If (Session("Results") = 0 And ((CType(Session("isTemplateOwner"), Boolean) Or CType(Session("isTemplateDelegate"), Boolean)) Or (Session("UserType") = 4 And Session("isADCHSR") <> "True"))) Then
                        doAddQuestionLink(cell, row, index, dt)
                    End If

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
                    cell.Controls.Add(ddlList)
                    cell.Wrap = True
                    cell.EnableViewState = True
                    cell.BorderStyle = BorderStyle.None
                    row.Cells.Add(cell)

                    'increment the visible question number
                    visibleQuestionNumber += 1
                ElseIf (strQuestionType = "M") Then
                    'Textfield Question
                    row.Cells.Clear()

                    'generate the link to the question
                    If (Session("Results") = 0 And ((CType(Session("isTemplateOwner"), Boolean) Or CType(Session("isTemplateDelegate"), Boolean)) Or (Session("UserType") = 4 And Session("isADCHSR") <> "True"))) Then
                        doAddQuestionLink(cell, row, index, dt)
                    End If

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
                    cell.Controls.Add(txtField)
                    cell.Wrap = True
                    cell.EnableViewState = True
                    cell.BorderStyle = BorderStyle.None
                    row.Cells.Add(cell)

                    'increment the visible question number
                    visibleQuestionNumber += 1
                ElseIf (strQuestionType = "N") Then
                    'Number Field Question
                    row.Cells.Clear()

                    'generate the link to the question
                    If (Session("Results") = 0 And ((CType(Session("isTemplateOwner"), Boolean) Or CType(Session("isTemplateDelegate"), Boolean)) Or (Session("UserType") = 4 And Session("isADCHSR") <> "True"))) Then
                        doAddQuestionLink(cell, row, index, dt)
                    End If

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
                    cell.Controls.Add(numField)
                    cell.Wrap = False
                    cell.EnableViewState = True
                    cell.BorderStyle = BorderStyle.None
                    row.Cells.Add(cell)

                    'increment the visible question number
                    visibleQuestionNumber += 1
                ElseIf (strQuestionType = "P") Then
                    'Picture Question
                    row.Cells.Clear()

                    'generate the link to the question
                    If (Session("Results") = 0 And ((CType(Session("isTemplateOwner"), Boolean) Or CType(Session("isTemplateDelegate"), Boolean)) Or (Session("UserType") = 4 And Session("isADCHSR") <> "True"))) Then
                        doAddQuestionLink(cell, row, index, dt)
                    End If

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
                    Try
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
                                        Try
                                            myStream.Write(CType(dt.Rows(index).Item("LIST_ITEM_IMAGE"), Byte()), 0, CType(dt.Rows(index).Item("LIST_ITEM_IMAGE"), Byte()).Length)
                                            myImage = System.Drawing.Image.FromStream(myStream)
                                            imageWidth = myImage.Width + 32
                                            imageHeight = myImage.Height + 32
                                        Catch ex As Exception
                                            Trace.Warn(ex.ToString)
                                        End Try
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
                    Catch ex As Exception
                        Trace.Warn(ex.ToString)
                    End Try

                    radbtnList.CssClass = "borderless"
                    radbtnList.ID = "radbtnList" & dt.Rows(index).Item("QUESTION_ID")
                    cell.Controls.Add(radbtnList)
                    cell.Wrap = True
                    cell.EnableViewState = True
                    cell.BorderStyle = BorderStyle.None
                    row.Cells.Add(cell)

                    'increment the visible question number
                    visibleQuestionNumber += 1
                ElseIf (strQuestionType = "Q") Then
                    'Hidden Field Question
                    row.Cells.Clear()

                    'generate the link to the question
                    If (Session("Results") = 0 And ((CType(Session("isTemplateOwner"), Boolean) Or CType(Session("isTemplateDelegate"), Boolean)) Or (Session("UserType") = 4 And Session("isADCHSR") <> "True"))) Then
                        doAddQuestionLink(cell, row, index, dt)
                    End If

                    'generate hidden field
                    If (CType(Session("isADCHSR"), Boolean)) Then
                        'generate question label
                        row.Cells.Clear()
                        cell = New System.Web.UI.WebControls.TableCell
                        cell.HorizontalAlign = HorizontalAlign.Left
                        Dim lblQuestionHidden As Label = New Label
                        lblQuestionHidden.Text = "Note: This field is not normally visible but contains the following: " & _
                                dt.Rows(index).Item("QUESTION_TEXT") & "<br/>"
                        lblQuestionHidden.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0, 153)
                        lblQuestionHidden.CssClass = "boldContentTeal"
                        lblQuestionHidden.ID = "lblQuestion" & dt.Rows(index).Item("QUESTION_ID")
                        cell.Controls.Add(lblQuestionHidden)
                        cell.Wrap = True
                        cell.EnableViewState = True
                        cell.BorderStyle = BorderStyle.None
                        row.Cells.Add(cell)
                    Else
                        'generate question label
                        cell = New System.Web.UI.WebControls.TableCell
                        cell.HorizontalAlign = HorizontalAlign.Left
                        Dim lblQuestionHidden As Label = New Label
                        lblQuestionHidden.Text = "Note: This field is not normally visible but contains the following: " & _
                                dt.Rows(index).Item("QUESTION_TEXT") & "<br/>"
                        lblQuestionHidden.ForeColor = System.Drawing.Color.FromArgb(0, 0, 0, 153)
                        lblQuestionHidden.CssClass = "boldContentTeal"
                        lblQuestionHidden.ID = "lblQuestion" & dt.Rows(index).Item("QUESTION_ID")
                        cell.Controls.Add(lblQuestionHidden)
                        cell.Wrap = True
                        cell.EnableViewState = True
                        cell.BorderStyle = BorderStyle.None
                        cell.ColumnSpan = 2
                        row.Cells.Add(cell)
                    End If
                ElseIf (strQuestionType = "R") Then
                    'Rating Question
                    row.Cells.Clear()

                    'generate the link to the question
                    If (Session("Results") = 0 And ((CType(Session("isTemplateOwner"), Boolean) Or CType(Session("isTemplateDelegate"), Boolean)) Or (Session("UserType") = 4 And Session("isADCHSR") <> "True"))) Then
                        doAddQuestionLink(cell, row, index, dt)
                    End If

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
                    Dim initial = dt.Rows(index).Item("RATING_INITIAL_VALUE")
                    Dim interval = dt.Rows(index).Item("RATING_STEP_VALUE")
                    Dim count = dt.Rows(index).Item("RATING_COUNT")
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
                    radbtnList.RepeatDirection = RepeatDirection.Horizontal
                    radbtnList.TextAlign = TextAlign.Right
                    cell.Controls.Add(radbtnList)
                    cell.Wrap = True
                    cell.EnableViewState = True
                    cell.BorderStyle = BorderStyle.None
                    row.Cells.Add(cell)

                    'increment the visible question number
                    visibleQuestionNumber += 1
                ElseIf (strQuestionType = "S") Then
                    '1-Line Field Question
                    row.Cells.Clear()

                    'generate the link to the question
                    If (Session("Results") = 0 And ((CType(Session("isTemplateOwner"), Boolean) Or CType(Session("isTemplateDelegate"), Boolean)) Or (Session("UserType") = 4 And Session("isADCHSR") <> "True"))) Then
                        doAddQuestionLink(cell, row, index, dt)
                    End If

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
                    cell.Controls.Add(txtField)
                    cell.Wrap = True
                    cell.EnableViewState = True
                    cell.BorderStyle = BorderStyle.None
                    row.Cells.Add(cell)

                    'increment the visible question number
                    visibleQuestionNumber += 1
                ElseIf (strQuestionType = "T") Then
                    'Type Question
                    row.Cells.Clear()

                    'generate the link to the question
                    If (Session("Results") = 0 And ((CType(Session("isTemplateOwner"), Boolean) Or CType(Session("isTemplateDelegate"), Boolean)) Or (Session("UserType") = 4 And Session("isADCHSR") <> "True"))) Then
                        doAddQuestionLink(cell, row, index, dt)
                    End If

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
                    Try
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
                    Catch ex As Exception
                        Trace.Warn(ex.ToString)
                    End Try

                    radbtnList.CssClass = "borderless"
                    radbtnList.ID = "radbtnList" & dt.Rows(index).Item("QUESTION_ID")
                    cell.Controls.Add(radbtnList)
                    cell.Wrap = True
                    cell.EnableViewState = True
                    cell.BorderStyle = BorderStyle.None
                    row.Cells.Add(cell)

                    'increment the visible question number
                    visibleQuestionNumber += 1
                ElseIf (strQuestionType = "X") Then
                    'Comment Question
                    row.Cells.Clear()

                    'generate the link to the question
                    If (Session("Results") = 0 And ((CType(Session("isTemplateOwner"), Boolean) Or CType(Session("isTemplateDelegate"), Boolean)) Or (Session("UserType") = 4 And Session("isADCHSR") <> "True"))) Then
                        doAddQuestionLink(cell, row, index, dt)
                    End If

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
                    cell.BorderStyle = BorderStyle.None
                    row.Cells.Add(cell)
                ElseIf (strQuestionType = "Z") Then
                    'Prepopulated Question
                    row.Cells.Clear()

                    'generate the link to the question
                    If (Session("Results") = 0 And ((CType(Session("isTemplateOwner"), Boolean) Or CType(Session("isTemplateDelegate"), Boolean)) Or (Session("UserType") = 4 And Session("isADCHSR") <> "True"))) Then
                        doAddQuestionLink(cell, row, index, dt)
                    End If

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
                    cell.Controls.Add(ddlList)
                    cell.Wrap = True
                    cell.EnableViewState = True
                    cell.BorderStyle = BorderStyle.None
                    row.Cells.Add(cell)

                    'increment the visible question number
                    visibleQuestionNumber += 1
                End If
                'add the row to the table and increment to move to the next row of data
                row.EnableViewState = True
                QuestionTable.Rows.Add(row)
                QuestionTable.EnableViewState = True
                index += 1
            End While
        Else
            If (Session("CurrentPage") Is Nothing) Then
                alert("There are either no questions or this template does not exist.  If you feel that this is an error please contact the appropriate owner or delegate of the survey template or contact the Survey Administrator.")
            Else
                Session("Alert") = "There are either no questions or this template does not exist.  If you feel that this is an error please contact the appropriate owner or delegate of the survey template or contact the Survey Administrator."
                Response.Write("<script type='text/javascript'>window.opener.location.href = window.opener.location.href;</script>")
                Response.Write("<script type='text/javascript'>window.close()</script>")
            End If
        End If
    End Sub

    'returns the template's name
    Private Function getTemplateName() As String
        Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
        Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
        Try
            sqlCommonAdapter.SelectCommand = sqlCommonAction

            'attach the connection to the command
            sqlCommonAction.Connection = Me.commonConn

            'Open the connection
            If (sqlCommonAction.Connection.State = ConnectionState.Closed) Then
                sqlCommonAction.Connection.Open()
            End If

            'get the template name

            sqlCommonAction.CommandText = "Select TEMPLATE_NAME from " & myUtility.getExtension() & "SAT_TEMPLATE where " & _
            "TEMPLATE_KEY = " & Session("seqTemplate")
            sqlCommonAdapter.Fill(Me.dsCommon, "TemplateName")

            If (Me.dsCommon.Tables("TemplateName").Rows.Count > 0) Then
                Return Me.dsCommon.Tables("TemplateName").Rows(0).Item(0)
            Else
                Return ""
            End If
        Catch ex As Exception
            Return ""
        End Try
    End Function

    'returns whether or not it was able to get any survey template owner information
    Private Function getOwnersInfo() As Boolean
        Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
        Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
        Try
            sqlCommonAdapter.SelectCommand = sqlCommonAction

            'attach the connection to the command
            sqlCommonAction.Connection = Me.commonConn

            'Open the connection
            If (sqlCommonAction.Connection.State = ConnectionState.Closed) Then
                sqlCommonAction.Connection.Open()
            End If

            'get the template name
            sqlCommonAction.CommandText = "Select EMAIL_ADDRESS from " & myUtility.getExtension() & "SAT_USER fu, " & _
            myUtility.getExtension() & "SAT_USER_PERMISSION fp where fp.TEMPLATE_KEY = " & Session("seqTemplate") & _
            " and fp.OWNER_IND = 1 and fu.USER_KEY = fp.USER_KEY"
            sqlCommonAdapter.Fill(Me.dsCommon, "Owners")

            'determine if any records were returned
            If (Me.dsCommon.Tables("Owners").Rows.Count > 0) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    'returns the reviewer's e-mail address
    Private Function getReviewerInfo() As Boolean
        Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
        Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
        Try
            sqlCommonAdapter.SelectCommand = sqlCommonAction

            'attach the connection to the command
            sqlCommonAction.Connection = Me.commonConn

            'Open the connection
            If (sqlCommonAction.Connection.State = ConnectionState.Closed) Then
                sqlCommonAction.Connection.Open()
            End If

            sqlCommonAction.CommandText = "Select EMAIL_ADDRESS, AUTH_DERIV_CLASSIFIER_IND, HUMAN_SUBJECTS_RESCH_IND, PRIVACY_IND from " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE where " & _
            "TEMPLATE_KEY = " & Session("seqTemplate") & " and CURRENT_IND = 1"
            sqlCommonAdapter.Fill(Me.dsCommon, "ReviewerEmail")
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
#End Region

#Region "Control Settings"
    'populates the fields on this form from the database
    Private Sub populateFields()
        'start transaction
        Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
        Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
        sqlCommonAdapter.SelectCommand = sqlCommonAction
        sqlCommonAction.Connection = Me.commonConn

        'Open the connection
        If (sqlCommonAction.Connection.State = ConnectionState.Closed) Then
            sqlCommonAction.Connection.Open()
        End If

        If Not (Session("adcHID") Is Nothing) Then
            Try
                'populate the comments box
                sqlCommonAction.CommandText = "select SIGNER_COMMENT from " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE " & _
                " WHERE TEMPLATE_KEY = " & Session("seqTemplate") & _
                " and AUTH_DERIV_CLASSIFIER_IND = 1 and CURRENT_IND = 1"
                sqlCommonAdapter.Fill(Me.dsCommon, "Comments")

                Me.txtComments.Text = Me.dsCommon.Tables("Comments").Rows(0).Item(0)
            Catch ex As Exception
                alert("Unable to retrieve your reviewer information.")
            End Try
        ElseIf Not (Session("privHID") Is Nothing) Then
            Try
                'populate the comments box
                sqlCommonAction.CommandText = "select SIGNER_COMMENT from " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE " & _
                " WHERE TEMPLATE_KEY = " & Session("seqTemplate") & _
                " and PRIVACY_IND = 1 and CURRENT_IND = 1"
                sqlCommonAdapter.Fill(Me.dsCommon, "Comments")

                Me.txtComments.Text = Me.dsCommon.Tables("Comments").Rows(0).Item(0)
            Catch ex As Exception
                alert("Unable to retrieve your reviewer information.")
            End Try
        ElseIf Not (Session("hsrHID") Is Nothing) Then
            Try
                'populate the comments box
                sqlCommonAction.CommandText = "select SIGNER_COMMENT from " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE " & _
                " WHERE TEMPLATE_KEY = " & Session("seqTemplate") & _
                " and HUMAN_SUBJECTS_RESCH_IND = 1 and CURRENT_IND = 1"
                sqlCommonAdapter.Fill(Me.dsCommon, "Comments")

                Me.txtComments.Text = Me.dsCommon.Tables("Comments").Rows(0).Item(0)

                sqlCommonAction.CommandText = "select HUMAN_SUBJECTS_RESCH_IND from " & _
                myUtility.getExtension() & "SAT_TEMPLATE where TEMPLATE_KEY = " & Session("seqTemplate")
                sqlCommonAdapter.Fill(Me.dsCommon, "HSRItems")

                Me.chkbxHSR.Checked = Me.dsCommon.Tables("HSRItems").Rows(0).Item(0)
            Catch ex As Exception
                alert("Unable to retrieve your reviewer information.")
            End Try
        Else
            Session("Alert") = "1. You are not authorized to do this.  Please contact the Survey Administrator if this is an error"
            Session("FromPage") = "./Preview.aspx"
            Session("CurrentPage") = "./Login.aspx"
            If (Session("UserType") = 1) Then
                Response.Write("<script>window.close();</script>")
            Else
                Response.Redirect("./Login.aspx")
            End If
        End If
    End Sub

    'add the hyperlink to the question for the user
    Private Sub doAddQuestionLink(ByRef cell As System.Web.UI.WebControls.TableCell, ByVal row As System.Web.UI.WebControls.TableRow, ByVal index As Integer, ByVal dt As DataTable)
        cell = New System.Web.UI.WebControls.TableCell
        Dim lblQuestionLink As Label = New Label
        lblQuestionLink.Text = "<a class='pointer' onclick=""javascript:window.opener.location.href = './Question.aspx?seqTemplate=" & Session("seqTemplate") & _
        "&amp;intQuestion=" & dt.Rows(index).Item("QUESTION_ID") & "'""><img border=0 src=""./Images/QuestionLookup.png"" " & _
        "alt=""Question " & dt.Rows(index).Item("QUESTION_ID") & """ ></a>"
        lblQuestionLink.ID = "lblQuestionLink" & dt.Rows(index).Item("QUESTION_ID")
        cell.Controls.Add(lblQuestionLink)
        cell.Wrap = False
        cell.Width = Unit.Pixel(40)
        cell.EnableViewState = True
        cell.BorderStyle = BorderStyle.None
        cell.VerticalAlign = VerticalAlign.Top
        row.Cells.Add(cell)
    End Sub
#End Region

#Region "Checks"
    'determines if the user is here to sign the document
    Private Function isSigner() As Boolean
        'setup the basic data items
        Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
        Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
        Try
            sqlCommonAction.Connection = myUtility.getConnection(Me.commonConn)
            sqlCommonAdapter.SelectCommand = sqlCommonAction

            'get the user's email address
            If (Me.dsCommon.Tables.Contains("Level1Email")) Then
                Me.dsCommon.Tables("Level1Email").Rows.Clear()
            End If

            sqlCommonAction.CommandText = "Select EMAIL_ADDRESS from " & myUtility.getExtension() & "SAT_USER where USER_KEY = " & _
            Session("UserID")
            sqlCommonAdapter.Fill(Me.dsCommon, "Level1Email")

            'get the user's record from the survey signature table if it exists
            If (Me.dsCommon.Tables.Contains("Level1Record")) Then
                Me.dsCommon.Tables("Level1Record").Rows.Clear()
            End If
            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE where TEMPLATE_KEY = " & _
            Session("seqTemplate") & " and EMAIL_ADDRESS = '" & Me.dsCommon.Tables("Level1Email").Rows(0).Item(0) & _
            "' and CURRENT_IND = 1"
            sqlCommonAdapter.Fill(Me.dsCommon, "Level1Record")

            'if there is a record then return true
            If (Me.dsCommon.Tables("Level1Record").Rows.Count() = 1) Then
                If (Me.dsCommon.Tables("Level1Record").Rows(0).Item("AUTH_DERIV_CLASSIFIER_IND")) Then
                    Session("adcHID") = Me.dsCommon.Tables("Level1Record").Rows(0).Item("HANFORD_ID")
                ElseIf (Me.dsCommon.Tables("Level1Record").Rows(0).Item("PRIVACY_IND")) Then
                    Session("privHID") = Me.dsCommon.Tables("Level1Record").Rows(0).Item("HANFORD_ID")
                Else
                    Session("hsrHID") = Me.dsCommon.Tables("Level1Record").Rows(0).Item("HANFORD_ID")
                End If
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Unable to determine if you are here to sign this template or not.  Please contact the Survey Administrator if you are here to sign this template.")
            Return False
        End Try
    End Function

    'returns whether or not the survey template has been completely signed off
    Private Function isSignedOff() As Boolean
        'start transaction
        Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
        Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
        Try
            sqlCommonAdapter.SelectCommand = sqlCommonAction

            'attach the connection to the command
            sqlCommonAction.Connection = myUtility.getConnection(Me.commonConn)

            'Open the connection
            If (sqlCommonAction.Connection.State = ConnectionState.Closed) Then
                sqlCommonAction.Connection.Open()
            End If

            sqlCommonAction.CommandText = "select * from " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE WHERE " & _
            " TEMPLATE_KEY = " & Session("seqTemplate") & " and SIGNED_IND = 1 and " & _
            " CURRENT_IND = 1"
            sqlCommonAdapter.Fill(Me.dsCommon, "SignedOff")

            If (Me.dsCommon.Tables("SignedOff").Rows.Count() = 3) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
#End Region

#Region "Accept"
    'handles the user choosing to accept the survey template
    Private Sub btnAccept_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAccept.Click
        Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
        Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
        If (myUtility.goodInput(Me.txtComments, True)) Then
            Try
                If (myUtility.setupTransaction(sqlCommonAction, Me.commonConn)) Then
                    'store the comments
                    sqlCommonAdapter.SelectCommand = sqlCommonAction
                    Dim blnContinue As Boolean = True
                    If Not (storeComments(sqlCommonAction, sqlCommonAdapter)) Then
                        alert("Failed to store your response. Your signature request has been temporarily denied.")
                        blnContinue = False
                    End If
                    Session("Comments") = Me.txtComments.Text

                    'if this is the HSR representative then store the HSR Items value too
                    If (Not (Session("hsrHID") Is Nothing) And blnContinue) Then
                        If Not (storeHSRItem(sqlCommonAction, sqlCommonAdapter)) Then
                            alert("Failed to store your response. Your signature request has been temporarily denied.")
                            blnContinue = False
                        End If
                    End If

                    If (blnContinue) Then
                        sqlCommonAction.Transaction.Commit()
                        're-direct the user to the signature page.
                        Session("CurrentPage") = "./signature.asp"
                        If Not (Session("adcHID") Is Nothing) Then
                            Response.Redirect("./signature.asp?Return=Preview.aspx&HID=" & Session("adcHID") & "&str_Data=By digitally signing this form you indicate that this survey template is acceptable for distribution.")
                        ElseIf Not (Session("hsrHID") Is Nothing) Then
                            Response.Redirect("./signature.asp?Return=Preview.aspx&HID=" & Session("hsrHID") & "&str_Data=By digitally signing this form you indicate that this survey template is acceptable for distribution.")
                        ElseIf Not (Session("privHID") Is Nothing) Then
                            Response.Redirect("./signature.asp?Return=Preview.aspx&HID=" & Session("privHID") & "&str_Data=By digitally signing this form you indicate that this survey template is acceptable for distribution.")
                        Else
                            alert("You are not authorized to do this.  Please contact the Survey Administrator if this is an error.")
                        End If
                    Else
                        sqlCommonAction.Transaction.Rollback()
                    End If
                Else
                    alert("Unable to accept this template at this time.")
                End If
            Catch ex As Exception
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Unable to accept this template at this time.")
            End Try
        Else
            alert("Possible malicious code entry in the comments.  Publication approval aborted.")
        End If
    End Sub
#End Region

#Region "Reject"
    'handles the user choosing to reject the survey template
    Private Sub btnReject_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReject.Click
        'start transaction
        Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
        Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
        If (myUtility.goodInput(Me.txtComments, True)) Then
            Try
                sqlCommonAdapter.SelectCommand = sqlCommonAction

                'attach the connection to the command
                sqlCommonAction.Connection = Me.commonConn

                'Open the connection
                If (sqlCommonAction.Connection.State = ConnectionState.Closed) Then
                    sqlCommonAction.Connection.Open()
                End If

                'store the reviewer's comments
                If (storeComments(sqlCommonAction, sqlCommonAdapter)) Then
                    'send an e-mail the user and the survey administrator about the rejection
                    If Not (sendReplyEmail(Me.txtComments.Text, True)) Then
                        alert("Failed to notify the user.  Please send the user an e-mail with your comments.")
                    End If
                Else
                    alert("Failed to store your response. Your rejection request has been temporarily denied.")
                End If

                'update the signature record to indicate the rejection
                If Not (resetSignatures(sqlCommonAction, sqlCommonAdapter)) Then
                    alert("Failed to reset signature record.  Notify the Survey Tool Administrator.")
                End If

                'destroy the session and close the window in any event
                Response.Write("<script type='text/javascript'>window.close()</script>")
            Catch ex As Exception
                alert("Unable to reject this template at this time.")
            End Try
        Else
            alert("Possible malicious code entry in the comments.  Publication rejection aborted.")
        End If
    End Sub
#End Region

#Region "Reset Code"
    'Handles the click event.  Doesn't do anything the button's postback capability is all that is needed.
    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        Response.Write("<script type='text/javascript'>window.location.href = window.location.href;</script>")
    End Sub
#End Region

#Region "Store Data"
    'stores the user's comments
    Private Function storeComments(ByVal sqlCommonAction As SqlClient.SqlCommand, ByVal sqlCommonAdapter As SqlClient.SqlDataAdapter) As Boolean
        Try
            'parameterized the text input to allow for things like quotes to get through
            Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@SIGNER_COMMENT", System.Data.SqlDbType.VarChar, 3600, "SIGNER_COMMENT")
            param0.Value = Me.txtComments.Text
            sqlCommonAction.Parameters.Add(param0)

            If Not (Session("adcHID") Is Nothing) Then
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE SET SIGNER_COMMENT = " & _
                "@SIGNER_COMMENT WHERE TEMPLATE_KEY = " & Session("seqTemplate") & _
                " and AUTH_DERIV_CLASSIFIER_IND = 1 and CURRENT_IND = 1"
                sqlCommonAction.ExecuteNonQuery()
            ElseIf Not (Session("privHID") Is Nothing) Then
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE SET SIGNER_COMMENT = " & _
                "@SIGNER_COMMENT WHERE TEMPLATE_KEY = " & Session("seqTemplate") & _
                " and PRIVACY_IND = 1 and CURRENT_IND = 1"
                sqlCommonAction.ExecuteNonQuery()
            ElseIf Not (Session("hsrHID") Is Nothing) Then
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE SET SIGNER_COMMENT = " & _
                "@SIGNER_COMMENT WHERE TEMPLATE_KEY = " & Session("seqTemplate") & _
                " and HUMAN_SUBJECTS_RESCH_IND = 1 and CURRENT_IND = 1"
                sqlCommonAction.ExecuteNonQuery()
            Else
                alert("You are not authorized to do this.  Please contact the Survey Administrator if this is an error.")
                Return False
            End If
            'clear out the parameters
            sqlCommonAction.Parameters.Clear()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    'stores whether or not the survey contains HSR items
    Private Function storeHSRItem(ByVal sqlCommonAction As SqlClient.SqlCommand, ByVal sqlCommonAdapter As SqlClient.SqlDataAdapter) As Boolean
        Try
            sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE SET " & _
            " HUMAN_SUBJECTS_RESCH_IND = " & IIf(Me.chkbxHSR.Checked, 1, 0) & _
            " WHERE TEMPLATE_KEY = " & Session("seqTemplate")
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    'store the user's signature
    Private Function storeSignature() As Boolean
        'start transaction
        Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
        Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
        Try
            sqlCommonAdapter.SelectCommand = sqlCommonAction

            'attach the connection to the command
            sqlCommonAction.Connection = Me.commonConn

            'Open the connection
            If (sqlCommonAction.Connection.State = ConnectionState.Closed) Then
                sqlCommonAction.Connection.Open()
            End If

            If Not (Session("adcHID") Is Nothing) Then
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE SET TRANSACTION_ID = " & _
                Session("TransID") & ", SIGNED_IND = 1 WHERE TEMPLATE_KEY = " & _
                Session("seqTemplate") & " and AUTH_DERIV_CLASSIFIER_IND = 1 and CURRENT_IND = 1"
                sqlCommonAction.ExecuteNonQuery()
            ElseIf Not (Session("privHID") Is Nothing) Then
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE SET TRANSACTION_ID = " & _
                Session("TransID") & ", SIGNED_IND = 1 WHERE TEMPLATE_KEY = " & _
                Session("seqTemplate") & " and PRIVACY_IND = 1 and CURRENT_IND = 1"
                sqlCommonAction.ExecuteNonQuery()
            ElseIf Not (Session("hsrHID") Is Nothing) Then
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE SET TRANSACTION_ID = " & _
                Session("TransID") & ", SIGNED_IND = 1 WHERE TEMPLATE_KEY = " & _
                Session("seqTemplate") & " and HUMAN_SUBJECTS_RESCH_IND = 1 and CURRENT_IND = 1"
                sqlCommonAction.ExecuteNonQuery()
            Else
                alert("You are not authorized to do this.  Please contact the Survey Administrator if this is an error.")
                Return False
            End If
            Return True
        Catch ex As Exception
            alert("Failed to store your signature.  Your attempt to authorize this survey template has been temporarily denied.  Please contact the Survey Administrator.")
            Return False
        End Try
    End Function

    'resets the dirty bit to 0 for the survey template allowing it to be copied to production
    Private Function resetDirtyFlag() As Boolean
        Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
        Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
        Try
            sqlCommonAdapter.SelectCommand = sqlCommonAction

            'attach the connection to the command
            sqlCommonAction.Connection = Me.commonConn
            'Open the connection
            If (sqlCommonAction.Connection.State = ConnectionState.Closed) Then
                sqlCommonAction.Connection.Open()
            End If

            sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE SET " & _
            " CHANGE_IND = 0 WHERE TEMPLATE_KEY = " & Session("seqTemplate")
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    'resets the signatures on the template
    Private Function resetSignatures(ByVal sqlCommonAction As SqlClient.SqlCommand, ByVal sqlCommonAdapter As SqlClient.SqlDataAdapter) As Boolean
        Try
            'parameterized the text input to allow for things like quotes to get through
            sqlCommonAction.Parameters.Clear()

            If Not (Session("adcHID") Is Nothing) Or Not (Session("hsrHID") Is Nothing) Or Not (Session("privHID") Is Nothing) Then
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE SET SIGNED_IND = " & _
                "0, TRANSACTION_ID = 0 WHERE TEMPLATE_KEY = " & Session("seqTemplate") & _
                " and CURRENT_IND = 1"
                sqlCommonAction.ExecuteNonQuery()
            Else
                alert("You are not authorized to do this.  Please contact the Survey Administrator if this is an error.")
                Return False
            End If
            'clear out the parameters
            sqlCommonAction.Parameters.Clear()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
#End Region

#Region "SendMail"
    'sends a reply e-mail to the survey administrator and the user about approval/rejection
    Private Function sendReplyEmail(ByVal strComments As String, Optional ByVal isRejection As Boolean = False, Optional ByVal isSignedOff As Boolean = False) As Boolean
        Trace.Warn("Sending user an email")
        Try
            'Set up the mail messsage
            Dim strFrom As String = ""
            Dim strTo As String = ""
            Dim strbldrSubject As New StringBuilder
            Dim adcReviewer As String = ""
            Dim hsrReviewer As String = ""
            Dim privReviewer As String = ""

            'get the survey template name
            Dim strTemplateName As String = getTemplateName()
            If (strTemplateName <> "") Then
                If (isRejection) Then
                    strbldrSubject.Append("The " & CType(strTemplateName, String) & _
                    " survey template has been rejected")
                Else
                    If (isSignedOff) Then
                        strbldrSubject.Append("The " & CType(strTemplateName, String) & _
                        " survey template has been completely signed")
                    Else
                        strbldrSubject.Append("The " & CType(strTemplateName, String) & _
                        " survey template has been accepted")
                    End If
                End If
            Else
                alert("Failed to get the survey template data.  This survey template may have been recently deleted.")
                Return False
            End If

            'append who did the action to the subject.
            If Not (Session("adcHID") Is Nothing) Then
                strbldrSubject.Append(" by the Authorized Derivative Classifier representative")
            ElseIf Not (Session("hsrHID") Is Nothing) Then
                strbldrSubject.Append(" by the Human Subjects Research representative")
            ElseIf Not (Session("privHID") Is Nothing) Then
                strbldrSubject.Append(" by the Privacy representative")
            End If

            If (isSignedOff) Then
                strbldrSubject.Append(" and can be published to production.")
            Else
                strbldrSubject.Append(".")
            End If

            If (getReviewerInfo() And Me.dsCommon.Tables("ReviewerEmail").Rows.Count() > 0) Then
                'populate the two strings
                Dim dr As DataRow
                For Each dr In Me.dsCommon.Tables("ReviewerEmail").Rows
                    If (CType(dr.Item("HUMAN_SUBJECTS_RESCH_IND"), Boolean) = True) Then
                        hsrReviewer = dr.Item("EMAIL_ADDRESS")
                    ElseIf (CType(dr.Item("PRIVACY_IND"), Boolean) = True) Then
                        privReviewer = dr.Item("EMAIL_ADDRESS")
                    Else
                        adcReviewer = dr.Item("EMAIL_ADDRESS")
                    End If
                Next

                'set the from e-mail address
                If Not (Session("adcHID") Is Nothing) Then
                    strFrom = adcReviewer
                ElseIf Not (Session("hsrHID") Is Nothing) Then
                    strFrom = hsrReviewer
                ElseIf Not (Session("privHID") Is Nothing) Then
                    strFrom = privReviewer
                End If

                Dim strbldrMessage As New StringBuilder
                'set up the message text that will be sent to the owners and survey administrator
                If (isRejection) Then
                    strbldrMessage.Append(Me.txtComments.Text())
                Else
                    strbldrMessage.Append(Session("Comments"))
                End If

                'get the owner e-mail(s) so that they are informed of the rejection/acceptance
                If Not (getOwnersInfo()) Then
                    alert("Failed to get owner information.  Please contact the survey administrator.")
                    Return False
                End If
                Dim row As DataRow
                For Each row In Me.dsCommon.Tables("Owners").Rows
                    strTo &= row.Item("EMAIL_ADDRESS") & ","
                Next

                'make sure reviewers and the administrator get copies
                Dim strCC As String = ""
                For Each row In Me.dsCommon.Tables("ReviewerEmail").Rows
                    strCC &= row.Item("EMAIL_ADDRESS") & ","
                Next
                strCC &= "survey.administrator@pnl.gov"

                Return (myUtility.genericSendMail(strFrom, strTo, strbldrSubject.ToString, strbldrMessage.ToString, strCC))
                Return True
            Else
                alert("Failed to get the reviewer information.  Please contact the survey administrator.")
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
#End Region
End Class
