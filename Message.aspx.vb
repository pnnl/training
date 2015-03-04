Imports System.Collections.Specialized
Public Class Message
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Message))
        Me.sqlMessage = New System.Data.SqlClient.SqlDataAdapter
        Me.SqlDeleteCommand = New System.Data.SqlClient.SqlCommand
        Me.commonConn = New System.Data.SqlClient.SqlConnection
        Me.SqlInsertCommand = New System.Data.SqlClient.SqlCommand
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand
        Me.SqlUpdateCommand = New System.Data.SqlClient.SqlCommand
        Me.dsCommon = New System.Data.DataSet
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).BeginInit()
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
        Me.SqlDeleteCommand.CommandText = "DELETE FROM VARCONN.[SAT_MESSAGE] WHERE (([MESSAGE_ID] = @Original_MESSAGE_ID) AN" & _
            "D ([CREATE_DATE] = @Original_CREATE_DATE) AND ([EXPIRATION_DATE] = @Original_EXP" & _
            "IRATION_DATE))"
        Me.SqlDeleteCommand.Connection = Me.commonConn
        Me.SqlDeleteCommand.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@Original_MESSAGE_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "MESSAGE_ID", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_CREATE_DATE", System.Data.SqlDbType.DateTime, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CREATE_DATE", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_EXPIRATION_DATE", System.Data.SqlDbType.DateTime, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "EXPIRATION_DATE", System.Data.DataRowVersion.Original, Nothing)})
        '
        'commonConn
        '
        Me.commonConn.FireInfoMessageEventOnUserErrors = False
        '
        'SqlInsertCommand
        '
        Me.SqlInsertCommand.CommandText = "INSERT INTO VARCONN.[SAT_MESSAGE] ([MESSAGE_ID], [MESSAGE_TEXT], [CREATE_DATE], [" & _
            "EXPIRATION_DATE]) VALUES (@MESSAGE_ID, @MESSAGE_TEXT, @CREATE_DATE, @EXPIRATION_" & _
            "DATE)"
        Me.SqlInsertCommand.Connection = Me.commonConn
        Me.SqlInsertCommand.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@MESSAGE_ID", System.Data.SqlDbType.Int, 0, "MESSAGE_ID"), New System.Data.SqlClient.SqlParameter("@MESSAGE_TEXT", System.Data.SqlDbType.[Char], 0, "MESSAGE_TEXT"), New System.Data.SqlClient.SqlParameter("@CREATE_DATE", System.Data.SqlDbType.DateTime, 0, "CREATE_DATE"), New System.Data.SqlClient.SqlParameter("@EXPIRATION_DATE", System.Data.SqlDbType.DateTime, 0, "EXPIRATION_DATE")})
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = "SELECT MESSAGE_ID, MESSAGE_TEXT, CREATE_DATE, EXPIRATION_DATE FROM VARCONN.SAT_ME" & _
            "SSAGE"
        Me.SqlSelectCommand1.Connection = Me.commonConn
        '
        'SqlUpdateCommand
        '
        Me.SqlUpdateCommand.CommandText = resources.GetString("SqlUpdateCommand.CommandText")
        Me.SqlUpdateCommand.Connection = Me.commonConn
        Me.SqlUpdateCommand.Parameters.AddRange(New System.Data.SqlClient.SqlParameter() {New System.Data.SqlClient.SqlParameter("@MESSAGE_ID", System.Data.SqlDbType.Int, 0, "MESSAGE_ID"), New System.Data.SqlClient.SqlParameter("@MESSAGE_TEXT", System.Data.SqlDbType.[Char], 0, "MESSAGE_TEXT"), New System.Data.SqlClient.SqlParameter("@CREATE_DATE", System.Data.SqlDbType.DateTime, 0, "CREATE_DATE"), New System.Data.SqlClient.SqlParameter("@EXPIRATION_DATE", System.Data.SqlDbType.DateTime, 0, "EXPIRATION_DATE"), New System.Data.SqlClient.SqlParameter("@Original_MESSAGE_ID", System.Data.SqlDbType.Int, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "MESSAGE_ID", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_CREATE_DATE", System.Data.SqlDbType.DateTime, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "CREATE_DATE", System.Data.DataRowVersion.Original, Nothing), New System.Data.SqlClient.SqlParameter("@Original_EXPIRATION_DATE", System.Data.SqlDbType.DateTime, 0, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "EXPIRATION_DATE", System.Data.DataRowVersion.Original, Nothing)})
        '
        'dsCommon
        '
        Me.dsCommon.DataSetName = "NewDataSet"
        Me.dsCommon.Locale = New System.Globalization.CultureInfo("en-US")
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected WithEvents Image1 As System.Web.UI.WebControls.Image
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents hlpPrevPassword As System.Web.UI.WebControls.Image
    Protected WithEvents Image2 As System.Web.UI.WebControls.Image
    Protected WithEvents lblExpires As System.Web.UI.WebControls.Label
    Protected WithEvents lblMessage As System.Web.UI.WebControls.Label
    Protected WithEvents txtMessage As System.Web.UI.WebControls.TextBox
    Protected WithEvents wdcExpirationDate As Infragistics.WebUI.WebSchedule.WebDateChooser
    Protected WithEvents hlpMessageSelect As System.Web.UI.WebControls.Image
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents sqlMessage As System.Data.SqlClient.SqlDataAdapter
    Protected WithEvents dsCommon As System.Data.DataSet
    Protected WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Protected requestItems As Array
    Protected navButtons As Collection
    Protected WithEvents lblNavBar As System.Web.UI.WebControls.Label
    Protected WithEvents rptrMessage As System.Web.UI.WebControls.Repeater
    Protected myNavBar As Navigation = New Navigation
    Protected myUtility As Utility = New Utility
    Protected WithEvents btnStart As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnPrev As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnReload As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnNext As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnLast As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnNew As System.Web.UI.WebControls.ImageButton
    Protected WithEvents btnInsert As System.Web.UI.WebControls.Button
    Protected WithEvents btnUpdate As System.Web.UI.WebControls.Button
    Protected WithEvents btnDelete As System.Web.UI.WebControls.Button
    Protected WithEvents btnReset As System.Web.UI.WebControls.Button
    Protected sqlCommonAction As New SqlClient.SqlCommand
    Protected sqlCommonAdapter As New SqlClient.SqlDataAdapter
    Protected WithEvents commonConn As System.Data.SqlClient.SqlConnection
    Friend WithEvents SqlDeleteCommand As System.Data.SqlClient.SqlCommand
    Friend WithEvents SqlInsertCommand As System.Data.SqlClient.SqlCommand
    Friend WithEvents SqlUpdateCommand As System.Data.SqlClient.SqlCommand
    Protected pageOptions As Integer
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

        'get the from page
        If Not (IsPostBack) Then
            Session("FromPage") = getFromPage()
        End If

        'Put user code to initialize the page here
        If (Session("isAuthenticated") <> "True") Then
            Session("pageSel") = "Login"
            redirect("./Login.aspx")
        ElseIf (myUtility.checkUserType(4)) Then
            redirect(Session("FromPage"))
        Else
            Session("CurrentPage") = "./Message.aspx"
            Session("pageSel") = "Message"
        End If

        'if the user's password has expired send them to the password reset page
        If (Session("isExpired") = "True") Then
            Session("CurrentPage") = "./Password.aspx"
            Session("pageSel") = "Password"
            redirect(Session("CurrentPage"))
        End If

        'generate collection of navigational buttons
        myUtility.makeNavCollection(Me.navButtons, Me.btnStart, Me.btnPrev, Me.btnReload, Me.btnNext, Me.btnLast, Me.btnNew)

        'Set the connection based on the machine
        myUtility.getConnection(Me.commonConn)

        'Check for user selection from the comments list if we're sure it's not a simple post-back
        If Not (IsPostBack) Then
            'get anything on the request line
            myUtility.getRequest(Page.Request.Url.Query())

            If Not (loadMessages()) Then
                alert("Failed to load the messages.")
            Else
                'to determine what of the nav bar buttons should be available
                myNavBar.wnb_MoveTo(Me.navButtons, Session("MessageMax"), Session("MessageRow"))

                'Populate the controls on the page
                setControls()

                'set the page options
                myUtility.optionPreSet(Session("seqMessageID"), Session("MessageMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions

                Me.btnInsert.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                Me.btnUpdate.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                Me.btnDelete.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
                Me.btnReset.Attributes.Add("onClick", "return confirm('Are you sure you want to perform this action.');")
            End If
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

#Region "Data Loading"
    'loads the comment tags
    Private Function loadMessages() As Boolean
        Trace.Warn("Loading Messages")
        Try
            'if we're doing a reload then get the data and mess with the current row only if somehow it ended up off the scale
            If (Me.dsCommon.Tables.Contains("Messages")) Then
                Me.dsCommon.Tables("Messages").Rows.Clear()
            End If

            'fill and bind to the repeater
            Me.SqlSelectCommand1.CommandText = Me.SqlSelectCommand1.CommandText.Replace("VARCONN.", myUtility.getExtension())
            Me.SqlSelectCommand1.Connection = Me.commonConn
            Me.sqlMessage.Fill(Me.dsCommon, "Messages")
            Session("MessageMax") = Me.dsCommon.Tables("Messages").Rows.Count()
            Me.rptrMessage.DataSource = Me.dsCommon.Tables("Messages")
            Me.rptrMessage.DataBind()

            'determine if there is any data and if the message id has been set
            If (Session("MessageMax") = 0) Then
                Session("seqMessageID") = -1
            ElseIf (Session("seqMessageID") Is Nothing) Then
                Session("seqMessageID") = 1
            End If

            'make sure the message row never exceeds the maximum messages
            If (Session("MessageRow") > Session("MessageMax")) Then
                Session("MessageRow") = Session("MessageMax") - 1
            End If

            'if the message id indicates a new record then move the message row to the maximum
            'else set it to just below the message id's number
            If (Session("seqMessageID") = -1) Then
                Session("MessageRow") = Session("MessageMax")
            Else
                Session("MessageRow") = Session("seqMessageID") - 1
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
    Private Sub setControls(Optional ByVal failure As Boolean = False)
        Trace.Warn("Setting Controls")
        If (Session("seqMessageID") > 0) Then
            Try
                If Not (failure) Then
                    Me.txtMessage.Text = Me.dsCommon.Tables("Messages").Rows(Session("MessageRow")).Item("MESSAGE_TEXT")
                    Me.wdcExpirationDate.Value = Me.dsCommon.Tables("Messages").Rows(Session("MessageRow")).Item("EXPIRATION_DATE")
                End If
                Me.lblNavBar.Text = "Message " & Session("MessageRow") + 1 & " of " & Session("MessageMax")
            Catch ex As Exception
                Trace.Warn(ex.ToString)
            End Try
        Else
            'set the text based on the message id and possible failure of a page command
            Me.txtMessage.Text = ""
            Me.wdcExpirationDate.Value = ""
            Me.lblNavBar.Text = "Message " & Session("MessageRow") + 1 & " of " & Session("MessageMax") + 1
        End If
    End Sub
#End Region

#Region "Nav Bar Events"
    'Moves the user to the first record
    Private Sub wnbMessages_MoveToStart(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnStart.Click
        Session("JSAction") = ""

        If Not (loadMessages()) Then
            alert("Failed to load the messages of the day.")
        Else
            Try
                Session("MessageRow") = 0
                Session("seqMessageID") = Me.dsCommon.Tables("Messages").Rows(Session("MessageRow")).Item("MESSAGE_ID")
                myNavBar.wnb_MoveTo(Me.navButtons, Session("MessageMax"), Session("MessageRow"))
                setControls()
                myUtility.optionPreSet(Session("seqMessageID"), Session("MessageMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the first message of the day.")
            End Try
        End If
    End Sub

    'Moves the user to the previous record
    Private Sub wnbMessages_MovePrev(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnPrev.Click
        Session("JSAction") = ""

        If Not (loadMessages()) Then
            alert("Failed to load the messages of the day.")
        Else
            Try
                Session("MessageRow") -= 1
                Session("seqMessageID") = Me.dsCommon.Tables("Messages").Rows(Session("MessageRow")).Item("MESSAGE_ID")
                myNavBar.wnb_MoveTo(Me.navButtons, Session("MessageMax"), Session("MessageRow"))
                setControls()
                myUtility.optionPreSet(Session("seqMessageID"), Session("MessageMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the previous message of the day.")
            End Try
        End If
    End Sub

    'reloads the page and the data from the database while attempting to locate the user's location
    Private Sub wnbMessages_MoveReload(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnReload.Click
        Session("JSAction") = ""

        If Not (loadMessages()) Then
            alert("Failed to load the messages of the day.")
        Else
            Try
                myNavBar.wnb_MoveTo(Me.navButtons, Session("MessageMax"), Session("MessageRow"))
                setControls()
                myUtility.optionPreSet(Session("seqMessageID"), Session("MessageMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error resetting the message of the day.")
            End Try
        End If
    End Sub

    'Moves the user to the next record
    Private Sub wnbMessages_MoveNext(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnNext.Click
        Session("JSAction") = ""

        If Not (loadMessages()) Then
            alert("Failed to load the messages of the day.")
        Else
            Try
                Session("MessageRow") += 1
                Session("seqMessageID") = Me.dsCommon.Tables("Messages").Rows(Session("MessageRow")).Item("MESSAGE_ID")
                myNavBar.wnb_MoveTo(Me.navButtons, Session("MessageMax"), Session("MessageRow"))
                setControls()
                myUtility.optionPreSet(Session("seqMessageID"), Session("MessageMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the next message of the day.")
            End Try
        End If
    End Sub

    'Moves the user to the last record
    Private Sub wnbMessages_MoveToEnd(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnLast.Click
        Session("JSAction") = ""

        If Not (loadMessages()) Then
            alert("Failed to load the messages of the day.")
        Else
            Try
                Session("MessageRow") = Session("MessageMax") - 1
                Session("seqMessageID") = Me.dsCommon.Tables("Messages").Rows(Session("MessageRow")).Item("MESSAGE_ID")
                myNavBar.wnb_MoveTo(Me.navButtons, Session("MessageMax"), Session("MessageRow"))
                setControls()
                myUtility.optionPreSet(Session("seqMessageID"), Session("MessageMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error moving to the last message of the day.")
            End Try
        End If
    End Sub

    'Inserts a record into the table
    Private Sub wnbMessages_InsertRow(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnNew.Click
        Session("JSAction") = ""

        If Not (loadMessages()) Then
            alert("Failed to load the messages of the day.")
        Else
            Try
                Session("seqMessageID") = -1
                Session("MessageRow") = Session("MessageMax")
                Me.dsCommon.Tables("Messages").Rows.Add(Me.dsCommon.Tables("Messages").NewRow())
                Me.dsCommon.Tables("Messages").Rows(Session("MessageMax")).Item("MESSAGE_ID") = Session("seqMessageID")
                Me.dsCommon.Tables("Messages").Rows(Session("MessageMax")).Item("CREATE_DATE") = Now
                Me.dsCommon.Tables("Messages").Rows(Session("MessageMax")).Item("MESSAGE_TEXT") = ""
                myNavBar.wnb_MoveTo(Me.navButtons, Session("MessageMax"), Session("MessageRow"))
                setControls()
                myUtility.optionPreSet(Session("seqMessageID"), Session("MessageMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                alert("Error creating new message of the day.")
            End Try
        End If
    End Sub
#End Region

#Region "Page Action"
    'Reacts to the user performing some kind of selection
    Private Sub btnUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUpdate.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'refresh the comments table
        If Not (loadMessages()) Then
            blnContinue = False
            alert("Failed to load the messages of the day. Update aborted.")
        End If

        'check for possible malicious code input
        myUtility.goodInput(Me.txtMessage, True)

        If (blnContinue) Then
            failure = doUpdate()

            'reload the data
            If Not (loadMessages()) Then
                blnContinue = False
                alert("Failed to load the messages of the day.")
            End If

            If (blnContinue) Then
                'to determine what of the nav bar buttons should be available
                myNavBar.wnb_MoveTo(Me.navButtons, Session("MessageMax"), Session("MessageRow"), Switch)

                'reset the page controls
                setControls()
                Session("TransExists") = False
                myUtility.optionPreSet(Session("seqMessageID"), Session("MessageMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'refresh the comments table
        If Not (loadMessages()) Then
            blnContinue = False
            alert("Failed to load the messages of the day. Delete aborted.")
        End If

        'check for possible malicious code input
        myUtility.goodInput(Me.txtMessage, True)

        If (blnContinue) Then
            'do a delete
            failure = doDelete()

            'reload the data
            If Not (loadMessages()) Then
                blnContinue = False
                alert("Failed to load the messages of the day.")
            End If

            If (blnContinue) Then
                'to determine what of the nav bar buttons should be available
                myNavBar.wnb_MoveTo(Me.navButtons, Session("MessageMax"), Session("MessageRow"), Switch)

                'reset the page controls
                setControls()
                Session("TransExists") = False
                myUtility.optionPreSet(Session("seqMessageID"), Session("MessageMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    Private Sub btnInsert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsert.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'refresh the comments table
        If Not (loadMessages()) Then
            blnContinue = False
            alert("Failed to load the messages of the day. Insert aborted.")
        End If

        'check for possible malicious code input
        myUtility.goodInput(Me.txtMessage, True)

        If (blnContinue) Then
            'do an insert
            failure = doInsert()

            'reload the data
               If Not (loadMessages()) Then
                blnContinue = False
                alert("Failed to load the messages of the day.")
            End If

            If (blnContinue) Then
                'to determine what of the nav bar buttons should be available
                myNavBar.wnb_MoveTo(Me.navButtons, Session("MessageMax"), Session("MessageRow"), Switch)

                'reset the page controls
                setControls()
                Session("TransExists") = False
                myUtility.optionPreSet(Session("seqMessageID"), Session("MessageMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub

    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Session("JSAction") = ""

        Dim failure As Boolean
        Dim blnContinue As Boolean = True

        'check for possible malicious code input
        myUtility.goodInput(Me.txtMessage, True)

        If (blnContinue) Then
            failure = doReset()

            If (blnContinue) Then
                'to determine what of the nav bar buttons should be available
                myNavBar.wnb_MoveTo(Me.navButtons, Session("MessageMax"), Session("MessageRow"), Switch)

                'reset the page controls
                setControls()
                Session("TransExists") = False
                myUtility.optionPreSet(Session("seqMessageID"), Session("MessageMax"), Me.pageOptions)

                Session("PageOptions") = Me.pageOptions
            End If
        End If
    End Sub
#End Region

#Region "Insert Code"
    'drives the commit and roll-back operations of the insert
    Private Function doInsert() As Boolean
        'start transaction
        Try
            Dim failure As Boolean = False
            If (myUtility.setupTransaction(sqlCommonAction, Me.commonConn)) Then
                'Check for actual information to insert
                If (Me.txtMessage.Text <> "") Then
                    If (insertMessage(sqlCommonAction, sqlCommonAdapter, True)) Then
                        Session("MessageMax") += 1
                        Session("seqMessageID") = Session("MessageMax")
                        Session("MessageRow") = Session("MessageMax") - 1
                        sqlCommonAction.Transaction.Commit()
                    Else
                        sqlCommonAction.Transaction.Rollback()
                        alert("Error adding message.")
                        failure = True
                    End If
                Else
                    alert("You must have some text for your message to insert this message.")
                    failure = True
                    sqlCommonAction.Transaction.Rollback()
                End If
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Error adding message.")
                failure = True
            End If
            Return failure
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Error adding message.")
        End Try
    End Function

    'attempts to insert a comment, returns false if it cannot
    Private Function insertMessage(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, Optional ByVal isNewRecord As Boolean = False) As Boolean
        Try
            sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_MESSAGE (MESSAGE_ID, CREATE_DATE, EXPIRATION_DATE, " & _
                "MESSAGE_TEXT) VALUES (" & Session("MessageMax") + 1 & ", '" & Now & "', '"

            'set the expiration for 15yrs from now if no date was supplied by the user
            If (Trim("" & Me.wdcExpirationDate.Text) = "") Then
                Trace.Warn("Expiration not set")
                Me.wdcExpirationDate.Value = Now.AddYears(15)
            End If

            sqlCommonAction.CommandText &= Me.wdcExpirationDate.Value & "', @MESSAGE_TEXT)"
            Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@MESSAGE_TEXT", System.Data.SqlDbType.VarChar, 4800, "MESSAGE_TEXT")
            param0.Value = Me.txtMessage.Text
            sqlCommonAction.Parameters.Add(param0)

            sqlCommonAction.ExecuteNonQuery()
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
        Return True
    End Function
#End Region

#Region "Update Code"
    'drives the commit and roll-back operations of the update
    Private Function doUpdate() As Boolean
        'start transaction
        Try
            Dim failure As Boolean = False
            If (myUtility.setupTransaction(sqlCommonAction, Me.commonConn)) Then
                'Check for actual information to update
                If (Me.txtMessage.Text <> "") Then
                    If (updateMessage(sqlCommonAction, sqlCommonAdapter)) Then
                        sqlCommonAction.Transaction.Commit()
                    Else
                        sqlCommonAction.Transaction.Rollback()
                        alert("Error updating message.")
                        failure = True
                    End If
                Else
                    sqlCommonAction.Transaction.Rollback()
                    alert("You must have some text in a message to update it.")
                    failure = True
                End If
            Else
                alert("Error updating message.")
                failure = True
            End If
            Return failure
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Error updating message.")
        End Try
    End Function

    'attempts to update a comment, returns false if it cannot
    Private Function updateMessage(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        'attempt to update the record in the database
        Try
            sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_MESSAGE SET MESSAGE_TEXT = @MESSAGE_TEXT, EXPIRATION_DATE = '"

            Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@MESSAGE_TEXT", System.Data.SqlDbType.VarChar, 4800, "MESSAGE_TEXT")
            param0.Value = Me.txtMessage.Text
            sqlCommonAction.Parameters.Add(param0)

            'set the expiration for 15yrs from now if no date was supplied by the user
            If (Trim("" & Me.wdcExpirationDate.Text) = "") Then
                Trace.Warn("Expiration not set")
                Me.wdcExpirationDate.Value = Now.AddYears(15)
            End If

            sqlCommonAction.CommandText &= Me.wdcExpirationDate.Value & _
            "' WHERE MESSAGE_ID = " & Session("seqMessageID")
            sqlCommonAction.ExecuteNonQuery()
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
        'start transaction
        Try
            Dim failure As Boolean = False
            If (myUtility.setupTransaction(sqlCommonAction, Me.commonConn)) Then
                'do the deletion
                If (deleteMessage(sqlCommonAction, sqlCommonAdapter)) Then
                    sqlCommonAction.Transaction.Commit()
                    'determine if we need to take a step back 
                    If (Session("MessageMax") = 1) Then
                        Session("MessageRow") = 0
                        Session("seqMessageID") = -1
                        Session("MessageMax") = 0
                        Me.dsCommon.Tables("Messages").Rows.Add(Me.dsCommon.Tables("Messages").NewRow())
                        Me.dsCommon.Tables("Messages").Rows(Session("MessageMax")).Item("MESSAGE_ID") = Session("seqMessageID")
                        Me.dsCommon.Tables("Messages").Rows(Session("MessageMax")).Item("CREATE_DATE") = Now
                        Me.wdcExpirationDate.Value = ""
                        Me.dsCommon.Tables("Messages").Rows(Session("MessageMax")).Item("MESSAGE_TEXT") = ""
                    ElseIf (Session("seqMessageID") = Session("MessageMax") And Session("MessageMax") - Session("MessageRow") = 1) Then
                        Session("seqMessageID") -= 1
                        Session("MessageMax") -= 1
                        Session("MessageRow") -= 1
                    Else
                        Session("MessageMax") -= 1
                    End If
                Else
                    sqlCommonAction.Transaction.Rollback()
                    alert("Error deleting message")
                    failure = True
                End If
            Else
                alert("Error deleting message.")
                failure = True
            End If
            Return failure
        Catch ex As Exception
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Error deleting message.")
        End Try
    End Function

    'attempts to delete a comment, returns false if it cannot
    Private Function deleteMessage(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter) As Boolean
        'attempt to delete the record from the database
        Try
            sqlCommonAction.CommandText = "Delete from " & myUtility.getExtension() & "SAT_MESSAGE where MESSAGE_ID = " & Session("seqMessageID")
            sqlCommonAction.ExecuteNonQuery()
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try

        'update the message ID's
        Dim lowIndex As Integer = Session("MessageRow")
        Dim highIndex As Integer = Session("MessageMax")

        'update the tagID's of the comments
        While (lowIndex < highIndex)
            Try
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_MESSAGE SET MESSAGE_ID = (MESSAGE_ID-1) " & _
                "WHERE MESSAGE_ID = " & _
                Me.dsCommon.Tables("Messages").Rows(lowIndex).Item("MESSAGE_ID")
                sqlCommonAction.ExecuteNonQuery()
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                Return False
            End Try
            lowIndex += 1
        End While
        Return True
    End Function
#End Region

#Region "Reset Code"
    'resets the page, will remove any new item being worked on at the end of the list
    Private Function doReset() As Boolean
        loadMessages()
        If (Session("seqMessageID") = -1) Then
            Me.dsCommon.Tables("Messages").Rows.Add(Me.dsCommon.Tables("Messages").NewRow())
            Me.dsCommon.Tables("Messages").Rows(Session("MessageMax")).Item("MESSAGE_ID") = Session("seqMessageID")
            Me.dsCommon.Tables("Messages").Rows(Session("MessageMax")).Item("CREATE_DATE") = Now
            Me.wdcExpirationDate.Value = ""
            Me.dsCommon.Tables("Messages").Rows(Session("MessageMax")).Item("MESSAGE_TEXT") = ""
        End If
        Return True
    End Function
#End Region
End Class