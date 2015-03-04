Imports System.Collections.Specialized
Public Class Main1
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Main1))
        Me.dsCommon = New System.Data.DataSet
        Me.sqlMessage = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlDeleteCommand = New System.Data.SqlClient.SqlCommand
        Me.commonConn = New System.Data.SqlClient.SqlConnection
        Me.SqlInsertCommand = New System.Data.SqlClient.SqlCommand
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlUpdateCommand = New System.Data.SqlClient.SqlCommand
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'dsCommon
        '
        Me.dsCommon.DataSetName = "NewDataSet"
        Me.dsCommon.Locale = New System.Globalization.CultureInfo("en-US")
        '
        'sqlMessage
        '
        Me.sqlMessage.DeleteCommand = Me.SqlDeleteCommand
        Me.sqlMessage.InsertCommand = Me.SqlInsertCommand
        Me.sqlMessage.SelectCommand = Me.SqlSelectCommand1
        Me.sqlMessage.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "SAT_MESSAGE", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("MESSAGE_ID", "MESSAGE_ID"), New System.Data.Common.DataColumnMapping("MESSAGE_TEXT", "MESSAGE_TEXT"), New System.Data.Common.DataColumnMapping("CREATE_DATE", "CREATE_DATE"), New System.Data.Common.DataColumnMapping("EXPIRATION_DATE", "EXPIRATION_DATE")})})
        Me.sqlMessage.UpdateCommand = Me.SqlUpdateCommand
        '
        'SqlDeleteCommand
        '
        Me.SqlDeleteCommand.CommandText = "DELETE FROM VARCONN.[SAT_MESSAGE] WHERE (([MESSAGE_ID] = @Original_MESSA" & _
            "GE_ID) AND ([CREATE_DATE] = @Original_CREATE_DATE) AND ([EXPIRATION_DATE] = @Ori" & _
            "ginal_EXPIRATION_DATE))"
        Me.SqlDeleteCommand.Connection = Me.commonConn
        Me.SqlDeleteCommand.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@Original_MESSAGE_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "MESSAGE_ID", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_CREATE_DATE", System.Data.SqlDbType.DateTime, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CREATE_DATE", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_EXPIRATION_DATE", System.Data.SqlDbType.DateTime, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "EXPIRATION_DATE", System.Data.DataRowVersion.Original, Nothing)})
        '
        'commonConn
        '
        Me.commonConn.FireInfoMessageEventOnUserErrors = False
        '
        'SqlInsertCommand
        '
        Me.SqlInsertCommand.CommandText = "INSERT INTO VARCONN.[SAT_MESSAGE] ([MESSAGE_ID], [MESSAGE_TEXT], [CREATE" & _
            "_DATE], [EXPIRATION_DATE]) VALUES (@MESSAGE_ID, @MESSAGE_TEXT, @CREATE_DATE, @EX" & _
            "PIRATION_DATE)"
        Me.SqlInsertCommand.Connection = Me.commonConn
        Me.SqlInsertCommand.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@MESSAGE_ID", System.Data.SqlDbType.Int, 0, "MESSAGE_ID"), New System.Data.SqlClient.SqlParameter("@MESSAGE_TEXT", System.Data.SqlDbType.[Char], 0, "MESSAGE_TEXT"), New System.Data.SqlClient.SqlParameter("@CREATE_DATE", System.Data.SqlDbType.DateTime, 0, "CREATE_DATE"), New System.Data.SqlClient.SqlParameter("@EXPIRATION_DATE", System.Data.SqlDbType.DateTime, 0, "EXPIRATION_DATE")})
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = "SELECT MESSAGE_ID, MESSAGE_TEXT, CREATE_DATE, EXPIRATION_DATE FROM VARCONN.S" & _
            "AT_MESSAGE"
        Me.SqlSelectCommand1.Connection = Me.commonConn
        '
        'SqlUpdateCommand
        '
        Me.SqlUpdateCommand.CommandText = resources.GetString("SqlUpdateCommand.CommandText")
        Me.SqlUpdateCommand.Connection = Me.commonConn
        Me.SqlUpdateCommand.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@MESSAGE_ID", System.Data.SqlDbType.Int, 0, "MESSAGE_ID"), New System.Data.SqlClient.SqlParameter("@MESSAGE_TEXT", System.Data.SqlDbType.[Char], 0, "MESSAGE_TEXT"), New System.Data.SqlClient.SqlParameter("@CREATE_DATE", System.Data.SqlDbType.DateTime, 0, "CREATE_DATE"), New System.Data.SqlClient.SqlParameter("@EXPIRATION_DATE", System.Data.SqlDbType.DateTime, 0, "EXPIRATION_DATE"), New System.Data.SqlClient.SqlParameter("@Original_MESSAGE_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "MESSAGE_ID", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_CREATE_DATE", System.Data.SqlDbType.DateTime, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CREATE_DATE", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_EXPIRATION_DATE", System.Data.SqlDbType.DateTime, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "EXPIRATION_DATE", System.Data.DataRowVersion.Original, Nothing)})
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected WithEvents Image1 As System.Web.UI.WebControls.Image
    Protected WithEvents lblPNNLText As System.Web.UI.WebControls.Label
    Protected WithEvents lblMailText As System.Web.UI.WebControls.Label
    Protected WithEvents lblNoteText As System.Web.UI.WebControls.Label
    Protected WithEvents lblUserName As System.Web.UI.WebControls.Label
    Protected WithEvents txtUserName As System.Web.UI.WebControls.TextBox
    Protected WithEvents lblPassword As System.Web.UI.WebControls.Label
    Protected WithEvents txtPassword As System.Web.UI.WebControls.TextBox
    Protected WithEvents btnSubmit As System.Web.UI.WebControls.Button
    Protected WithEvents btnReset As System.Web.UI.WebControls.Button
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents lblIntroduction As System.Web.UI.WebControls.Label
    Protected WithEvents lblMessage As System.Web.UI.HtmlControls.HtmlGenericControl
    Protected WithEvents lblIntroduction2 As System.Web.UI.WebControls.Label
    Protected WithEvents dsCommon As System.Data.DataSet
    Protected WithEvents sqlMessage As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents txtSearch As System.Web.UI.WebControls.TextBox
    Protected WithEvents searchSubmit As System.Web.UI.WebControls.Button
    Protected WithEvents Td1 As System.Web.UI.HtmlControls.HtmlTableCell
    Protected requestItems As Array
    Protected myUtility As Utility = New Utility
    Protected WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents commonConn As System.Data.SqlClient.SqlConnection
    Friend WithEvents SqlDeleteCommand As System.Data.SqlClient.SqlCommand
    Friend WithEvents SqlInsertCommand As System.Data.SqlClient.SqlCommand
    Friend WithEvents SqlUpdateCommand As System.Data.SqlClient.SqlCommand
    Protected WithEvents JSAction As String

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

#Region "Basic"
    'loads the page
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load
        Trace.Warn("Loading Page")
        InitializeComponent()
        Session("JSAction") = ""

        'catch incoming alert messages
        alert(Session("Alert"))

        'clear session
        myUtility.clearSession()

        If (Session.IsNewSession) Then
            Session("Alert") = "Your session has timed out for security purposes.  Please log in again."
        End If

        'get anything on the request line
        myUtility.getRequest(Page.Request.Url.Query())

        'Put user code to initialize the page here
        If (Session("isAuthenticated") <> "True") Then
            Session("CurrentPage") = "./Login.aspx"
            Session("pageSel") = "Login"
            Response.Redirect(Session("CurrentPage"))
        Else
            Session("CurrentPage") = "./Main.aspx"
            Session("pageSel") = "Main"
        End If

        'if the user's password has expired send them to the password reset page
        If (Session("isExpired") = "True") Then
            Session("CurrentPage") = "./Password.aspx"
            Session("pageSel") = "Password"
            redirect(Session("CurrentPage"))
        End If

        'Set the connection based on the machine
        myUtility.getConnection(Me.commonConn)

        If Not (IsPostBack) Then
            'load the messages of the day
            If Not (loadMessages()) Then
                alert("Error loading message(s) of the day.")
            End If

            'get the from page 
            Session("FromPage") = getFromPage()
        End If

        Session("JSAction") = JSAction
    End Sub

    'handles error messages
    Public Sub alert(ByVal message As String)
        If (message <> "") Then
            JSAction = "<script type='text/javascript'>window.alert('" & message & "');</script>"
            Session("Alert") = ""
        End If
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
#End Region

#Region "Message Handling"
    'loads the messages of the day
    Private Function loadMessages() As Boolean
        Try
            Dim blnUpdated As Boolean = False
            Me.SqlSelectCommand1.CommandText &= " order by CREATE_DATE DESC"
            Me.SqlSelectCommand1.CommandText = Me.SqlSelectCommand1.CommandText.Replace("VARCONN.", myUtility.getExtension())
            Me.SqlSelectCommand1.Connection = Me.CommonConn
            Me.sqlMessage.Fill(Me.dsCommon, "Messages")
            If (Me.dsCommon.Tables("Messages").Rows.Count() > 0) Then
                'Erase any expired messages
                Dim blnMessages As Boolean = False
                Dim intIndex As Integer = 0
                Me.lblMessage.InnerHtml = "<font class='medium'><br/>Messages of the Day<br/></font><hr width='33%' align='center'><br/>"
                While (intIndex < Me.dsCommon.Tables("Messages").Rows.Count())
                    'If the item is expired skip it and move on otherwise add it to the message of the day text
                    If (Me.dsCommon.Tables("Messages").Rows(intIndex).Item("EXPIRATION_DATE") >= Now) Then
                        Me.lblMessage.InnerHtml &= "<font class='small'>" & Me.dsCommon.Tables("Messages").Rows(intIndex).Item("MESSAGE_TEXT") & _
                            "</font><br/><br/><font class='small' size=3>As of: " & _
                            Me.dsCommon.Tables("Messages").Rows(intIndex).Item("CREATE_DATE") & "</font><br/><hr width='33%' align='center'><br/>"
                        blnMessages = True
                    End If
                    intIndex += 1
                End While

                'if there are no messages just hide the text completely.
                If Not (blnMessages) Then
                    Me.lblMessage.InnerHtml = ""
                End If
            End If
            Return True
        Catch ex As Exception
            Me.lblMessage.InnerHtml = ""
            Return False
        End Try
    End Function
#End Region
End Class
