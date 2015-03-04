Imports System.Text
Imports System.Collections.Specialized
Imports System.Net

Public Class Utility
    Inherits System.Web.UI.Page

#Region "Approval"
    'Determines if the template has been approved before
    Public Function isApproved(ByVal dsCommon As DataSet, ByVal sqlConn As SqlClient.SqlConnection) As Boolean
        'Set up the common data adapter and delete command
        Try
            Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
            Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
            sqlCommonAction.Connection = sqlConn
            sqlCommonAdapter.SelectCommand = sqlCommonAction

            'Open the connection
            If (sqlCommonAction.Connection.State = ConnectionState.Closed) Then
                sqlCommonAction.Connection.Open()
            End If

            If (dsCommon.Tables.Contains("Signed")) Then
                dsCommon.Tables("Signed").Rows.Clear()
            End If
            'get the records for that date if they exist
            sqlCommonAction.CommandText = "Select * from " & getExtension() & "SAT_TEMPLATE_SIGNATURE where " & _
            "TEMPLATE_KEY = " & Session("seqTemplate") & " and CURRENT_IND = 1" & _
            " and SIGNED_IND = 1"
            sqlCommonAdapter.Fill(dsCommon, "Signed")

            If (dsCommon.Tables("Signed").Rows.Count() = 3) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Session("Alert") = "Failed to determine the signature status of this template."
            Return False
        End Try
    End Function

    'Determines if the template has been routed for signature
    Public Function isRouted(ByVal dsCommon As DataSet, ByVal sqlConn As SqlClient.SqlConnection) As Boolean
        'Set up the common data adapter and delete command
        Try
            Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
            Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
            sqlCommonAction.Connection = sqlConn

            'Open the connection
            If (sqlCommonAction.Connection.State = ConnectionState.Closed) Then
                sqlCommonAction.Connection.Open()
            End If

            sqlCommonAction.CommandText = "Select CURRENT_IND, SIGNED_IND, HANFORD_ID, TEMPLATE_KEY " & _
            "from " & getExtension() & "SAT_TEMPLATE_SIGNATURE where TEMPLATE_KEY = " & Session("seqTemplate") & _
            " and SIGNED_IND = 1 and CURRENT_IND = 1"
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            sqlCommonAdapter.Fill(dsCommon, "Signed")
            If (dsCommon.Tables("Signed").Rows.Count() > 0) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Session("Alert") = "Unable to determine if this template has been routed or not."
            Return False
        End Try
    End Function

    'makes the template dirty
    Public Function makeDirty(ByVal sqlCommonAction As System.Data.SqlClient.SqlCommand, ByVal sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter, ByVal dsCommon As DataSet, ByVal sqlConn As SqlClient.SqlConnection) As Boolean
        Try
            'update the template making it dirty
            sqlCommonAction.CommandText = "Update " & getExtension() & "SAT_TEMPLATE Set CHANGE_IND = 1 where " & _
                "TEMPLATE_KEY = " & Session("seqTemplate")
            sqlCommonAction.ExecuteNonQuery()

            'determine if the template has been routed
            If (isRouted(dsCommon, sqlConn)) Then
                'determine if the template has been approved, routed, or is only dirty
                If (Session("isApproved")) Then
                    Session("Alert") = "This template has been previously approved but you must send it for signature again to publish these changes to production."
                    Return True
                Else
                    Session("Alert") = "This template has been routed for signature but has been partially signed.  You must re-submit this template for signature."
                    'reset the signature flag and remove the transaction id
                    If Not (resetSignatures(sqlCommonAction)) Then
                        Session("Alert") = "Failed to reset the signatures.  Your request has been temporarily denied."
                        Return False
                    Else
                        Session("Alert") = "The signatures have been re-set.  You must re-submit this survey template for approval to notify the approvers of the changes."
                        Return True
                    End If
                End If
            Else
                Return True
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString())
            Session("Alert") = "Failed to mark the survey template as changed.  Your request has been temporarily denied."
            Return False
        End Try
    End Function

    'resets the signature flag and the transaction id's for the current signature process
    Public Function resetSignatures(ByVal sqlCommonAction As SqlClient.SqlCommand) As Boolean
        Try
            'Update both persons
            sqlCommonAction.CommandText = "Update " & getExtension() & "SAT_TEMPLATE_SIGNATURE " & _
            " set SIGNED_IND = 0, TRANSACTION_ID = 0 where TEMPLATE_KEY = " & _
            Session("seqTemplate") & " and CURRENT_IND = 1"
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString())
            Session("Alert") = "Failed to re-set the ADC, HSR, and Privacy representatives for the digital signature process.  Your request has been temporarily denied.  Please contact the Survey Tool Administrator and try again."
            Return False
        End Try
    End Function

    'determines if the template is dirty
    Public Function isDirty(ByVal dsCommon As DataSet, ByVal sqlConn As SqlClient.SqlConnection) As Boolean
        'Set up the common data adapter and delete command
        Try
            Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
            Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
            sqlCommonAction.Connection = sqlConn
            sqlCommonAdapter.SelectCommand = sqlCommonAction

            'Open the connection
            If (sqlCommonAction.Connection.State = ConnectionState.Closed) Then
                sqlCommonAction.Connection.Open()
            End If

            If (dsCommon.Tables.Contains("Dirty")) Then
                dsCommon.Tables("Dirty").Rows.Clear()
            End If
            'get the records for that date if they exist
            sqlCommonAction.CommandText = "Select CHANGE_IND from " & getExtension() & "SAT_TEMPLATE where " & _
            "TEMPLATE_KEY = " & Session("seqTemplate")
            sqlCommonAdapter.Fill(dsCommon, "Dirty")

            If (dsCommon.Tables("Dirty").Rows.Count() = 1) Then
                Return CType(dsCommon.Tables("Dirty").Rows(0).Item("CHANGE_IND"), Boolean)
            Else
                Return False
            End If
        Catch ex As Exception
            Session("Alert") = "Failed to determine if the template is dirty."
            Return False
        End Try
    End Function

    'determines if the template surveys have results
    Public Function hasResults(ByVal dsCommon As DataSet, ByVal sqlConn As SqlClient.SqlConnection) As Boolean
        'Set up the common data adapter and delete command
        Try
            Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
            Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
            sqlCommonAction.Connection = sqlConn
            sqlCommonAdapter.SelectCommand = sqlCommonAction

            'Open the connection
            If (sqlCommonAction.Connection.State = ConnectionState.Closed) Then
                sqlCommonAction.Connection.Open()
            End If

            If (dsCommon.Tables.Contains("resultsCount")) Then
                dsCommon.Tables("resultsCount").Rows.Clear()
            End If
            'get the records for that date if they exist
            sqlCommonAction.CommandText = "Select count(*) AS RESULT_COUNT from " & getExtension() & "SAT_RESPONSE SR," & _
            " SAT_TEMPLATE_SURVEY TS, SAT_TEMPLATE ST where TS.SURVEY_KEY = SR.SURVEY_KEY" & _
            " AND TS.TEMPLATE_KEY = ST.TEMPLATE_KEY AND ST.TEMPLATE_KEY = " & Session("seqTemplate")
            sqlCommonAdapter.Fill(dsCommon, "resultsCount")

            If (dsCommon.Tables("resultsCount").Rows.Count() = 1) Then
                Return CType(dsCommon.Tables("resultsCount").Rows(0).Item("RESULT_COUNT"), Boolean)
            Else
                Return False
            End If
        Catch ex As Exception
            Session("Alert") = "Failed to determine if the template has results for surveys."
            Return False
        End Try
    End Function
#End Region

#Region "Basic"
    'Creates Session variables for the query string if anything is in it and returns true
    'Usage: myUtility.getRequest(Page.Request.Url.Query())
    Public Sub getRequest(ByVal pageURL As String)
        'get count of records - to save time if there are no items available.
        Dim requestItems As Array
        Dim requestString As String = pageURL
        If (requestString.Length > 0) Then
            requestString = requestString.Substring(1)
            requestItems = requestString.Split("&")
            Dim str As String
            For Each str In requestItems
                Dim temparr As Array
                temparr = str.Split("=")
                Session("" & CType(temparr(0), String) & "") = CType(temparr(1), String)
            Next
        End If
    End Sub

    'Sets up the transaction
    Public Function setupTransaction(ByRef sqlCommonAction As System.Data.SqlClient.SqlCommand, ByRef sqlConnection As SqlClient.SqlConnection) As Boolean
        Trace.Warn("Setting up Transaction")
        Try
            'Declare a transaction instance
            Dim sqlTransaction As Data.SqlClient.SqlTransaction

            'attach the connection to the command
            sqlCommonAction.Connection = sqlConnection

            'Open the connection
            If (sqlCommonAction.Connection.State = ConnectionState.Closed) Then
                sqlCommonAction.Connection.Open()
            End If

            'Begin the transaction
            sqlTransaction = sqlCommonAction.Connection.BeginTransaction()

            'attach the transaction to the command
            sqlCommonAction.Transaction = sqlTransaction
            Session("transExists") = True
            Return True
        Catch ex As Exception
            Session("transExists") = False
            Return False
        End Try
    End Function

    'generate control array of buttons on the page
    Public Sub makeNavCollection(ByRef navButtons As Collection, ByVal btnStart As ImageButton, ByVal btnPrev As ImageButton, ByVal btnReload As ImageButton, ByVal btnNext As ImageButton, ByVal btnLast As ImageButton, ByVal btnNew As ImageButton)
        navButtons = New Collection
        navButtons.Add(btnStart)
        navButtons.Add(btnPrev)
        navButtons.Add(btnReload)
        navButtons.Add(btnNext)
        navButtons.Add(btnLast)
        navButtons.Add(btnNew)
    End Sub

    'Sets the page options available for the selected Comment
    Public Sub setPageOptions(ByVal OptionSetting As Integer, ByVal ddl As DropDownList, Optional ByVal strFirstItem As String = "", Optional ByVal isExistingRecord As Boolean = False)
        Trace.Warn("Set Page Options")

        ddl.Items.Clear()
        Dim li As System.Web.UI.WebControls.ListItem = New System.Web.UI.WebControls.ListItem
        li.Text = "--Select--"
        li.Value = "--Select--"
        ddl.Items.Insert(0, li)
        If (OptionSetting = 3) Then
            li = New System.Web.UI.WebControls.ListItem
            li.Text = "Insert"
            li.Value = "Insert"
            ddl.Items.Add(li)
        ElseIf (OptionSetting = 4) Then
            li = New System.Web.UI.WebControls.ListItem
            li.Text = "Update"
            li.Value = "Update"
            ddl.Items.Add(li)
            li = New System.Web.UI.WebControls.ListItem
            li.Text = "Delete"
            li.Value = "Delete"
            ddl.Items.Add(li)
        ElseIf (OptionSetting = 5) Then
            li = New System.Web.UI.WebControls.ListItem
            li.Text = "Insert"
            li.Value = "Insert"
            ddl.Items.Add(li)
            li = New System.Web.UI.WebControls.ListItem
            li.Text = "Update"
            li.Value = "Update"
            ddl.Items.Add(li)
        ElseIf (OptionSetting <> 2) Then
            If Not (isExistingRecord) Then
                li = New System.Web.UI.WebControls.ListItem
                If (strFirstItem <> "") Then
                    li.Text = strFirstItem
                    li.Value = strFirstItem
                Else
                    li.Text = "Insert"
                    li.Value = "Insert"
                End If
                ddl.Items.Add(li)
            End If
            If (OptionSetting <> 1) Then
                li = New System.Web.UI.WebControls.ListItem
                li.Text = "Update"
                li.Value = "Update"
                ddl.Items.Add(li)
                li = New System.Web.UI.WebControls.ListItem
                li.Text = "Delete"
                li.Value = "Delete"
                ddl.Items.Add(li)
            End If
        End If
        li = New System.Web.UI.WebControls.ListItem
        li.Text = "Reset"
        li.Value = "Reset"
        ddl.Items.Add(li)
    End Sub

    'determines how to set the page options
    Public Sub optionPreSet(ByVal itemID As Integer, ByVal itemMax As Integer, ByRef pgOpt As Integer)
        'set the page options based on question availability and the whether or not the page is dirty
        If (Session("Machine") = "Development") Then
            If (itemID > 0 And itemMax > 0) Then
                pgOpt = 0
            Else
                pgOpt = 1
            End If
        Else
            pgOpt = 2
        End If
    End Sub

    'Alter at your own peril!'
    'determines how to set the page options
    Public Sub optionPreSetTemplate(ByVal itemID As Integer, ByVal itemMax As Integer, ByRef pgOpt As Integer, ByVal isTemplate As Boolean)
        'set the page options based on question availability and the whether or not the page is dirty
        If (isTemplate) Then
            If (Session("CurrentPage") = "./Template.aspx") Then
                If (CType(Session("isTemplateOwner"), Boolean) Or CType(Session("isTemplateDelegate"), Boolean)) Then
                    If (Session("UserType") = 3) Then
                        If (itemID <> -1) Then
                            pgOpt = 1 'needs to be update instead of insert on this option
                        Else
                            pgOpt = 2
                        End If
                    ElseIf (Session("UserType") = 4) Then
                        If (itemID <> -1) Then
                            pgOpt = 1
                        Else
                            pgOpt = 2
                        End If
                    Else
                        pgOpt = 2
                    End If
                ElseIf (Session("UserType") = 4) Then
                    If (itemID <> -1) Then
                        pgOpt = 1
                    Else
                        pgOpt = 2
                    End If
                Else
                    pgOpt = 2
                End If
            Else
                pgOpt = 2
            End If
        Else
            If (Session("CurrentPage") = "./Survey.aspx") Then
                If (CType(Session("isSurveyOwner"), Boolean) Or CType(Session("isSurveyDelegate"), Boolean)) Then
                    If (Session("UserType") = 2 Or Session("UserType") = 3) Then
                        If (itemID <> -1) Then
                            pgOpt = 1 'needs to be insert on this option instead of insert
                        Else
                            pgOpt = 2
                        End If
                    ElseIf (Session("UserType") = 4) Then
                        If (itemID <> -1) Then
                            pgOpt = 4
                        Else
                            pgOpt = 2
                        End If
                    Else
                        pgOpt = 2
                    End If
                ElseIf (Session("UserType") = 4) Then
                    If (itemID <> -1) Then
                        pgOpt = 4
                    Else
                        pgOpt = 2
                    End If
                Else
                    pgOpt = 2
                End If
            Else
                pgOpt = 2
            End If
        End If
    End Sub

    'determines how to set the page options
    Public Sub optionPreSetDelegates(ByVal itemID As Integer, ByVal itemMax As Integer, ByVal ddl As DropDownList)
        'set the page options based on question availability and the whether or not the page is dirty
        If (Session("Machine") = "Development") Then
            If (itemID > 0 And itemMax > 0) Then
                setPageOptions(0, ddl, "", True)
            Else
                setPageOptions(1, ddl)
            End If
        Else
            setPageOptions(2, ddl)
        End If
    End Sub

    'clears specified session variables
    Public Sub clearSession()
        If (Session("CurrentPage") <> "./SurveyResponse.aspx") Then
            Session("InputSwitch") = Nothing
        End If
    End Sub
#End Region

#Region "Grid"
    'Processes checking and unchecking checkboxes in the grid passed
    Public Sub checkBoxes(ByRef dg As DataGrid, Optional ByVal Reset As Boolean = False)
        If (Reset = False) Then
            'Check to see if any item is unchecked
            Dim blnUnChecked As Boolean = False
            Dim Index As Integer = 0
            While (Index < dg.Items.Count())
                If (CType(dg.Items(Index).Controls(0).Controls(1), CheckBox).Checked = False) Then
                    blnUnChecked = True
                    Index = dg.Items.Count()
                End If
                Index += 1
            End While

            If (blnUnChecked = True) Then
                'Check All items
                Index = 0
                While (Index < dg.Items.Count())
                    CType(dg.Items(Index).Controls(0).Controls(1), CheckBox).Checked = True
                    Index += 1
                End While
            Else
                'Uncheck All items
                Index = 0
                While (Index < dg.Items.Count())
                    CType(dg.Items(Index).Controls(0).Controls(1), CheckBox).Checked = False
                    Index += 1
                End While
            End If
        Else
            'Uncheck All items
            Dim Index As Integer = 0
            While (Index < dg.Items.Count())
                CType(dg.Items(Index).Controls(0).Controls(1), CheckBox).Checked = False
                Index += 1
            End While
        End If
    End Sub
#End Region

#Region "Mail"
    'Sends an e-mail to users with a link to the preview page along with comments.
    Public Function sendMail(ByVal Email As String, ByVal strName As String, ByVal strHID As String, ByVal strInformation As String, ByVal sqlCommonAction As SqlClient.SqlCommand, ByVal sqlCommonAdapter As SqlClient.SqlDataAdapter, ByVal dsCommon As DataSet, ByVal strServerName As String, Optional ByVal blnOldADC As Boolean = False, Optional ByVal blnADC As Boolean = False, Optional ByVal blnPriv As Boolean = False) As Boolean
        Trace.Warn("Sending user an email")
        'Set up the mail messsage
        Dim strFrom As String
        'remove this
        Dim strTo As String = "survey.administrator@pnl.gov"
        'end remove this
        'activate this
        'Dim strTo As String = Email
        'end activate this
        Dim strbldrSubject As New StringBuilder("The " & CType(Session("TemplateName"), String) & _
            " survey template has been submitted for your review.")

        'get the current user's name and e-mail
        Try
            sqlCommonAction.CommandText = "Select * from " & getExtension() & "SAT_USER where USER_KEY = " & Session("UserID")
            sqlCommonAdapter.Fill(dsCommon, "ModifyID")
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try

        'check to be sure we got their information
        If (dsCommon.Tables("ModifyID").Rows.Count < 1) Then
            Session("Alert") = "Critical error getting your user information for sending review e-mails.  Unable to locate your information.  You may have been removed from the database in the last few minutes.  Contact the Survey Administrator immediately."
            Return False
        End If

        'generate the Modifier's Name
        Dim strByName As String = ""
        If (dsCommon.Tables("ModifyID").Rows(0).Item("MIDDLE_NAME") = "") Then
            strByName = dsCommon.Tables("ModifyID").Rows(0).Item("LAST_NAME") & _
                ", " & dsCommon.Tables("ModifyID").Rows(0).Item("FIRST_NAME") & "."
        Else
            strByName = dsCommon.Tables("ModifyID").Rows(0).Item("LAST_NAME") & _
                ", " & dsCommon.Tables("ModifyID").Rows(0).Item("FIRST_NAME") & " " & _
                dsCommon.Tables("ModifyID").Rows(0).Item("MIDDLE_NAME") & "."
        End If

        Dim strbldrMessage As New StringBuilder
        'remove this
        strbldrMessage.Append("To:" & Email & "<br/>")
        'end remove this
        'notify the user of their status change
        If (blnADC = True) Then
            'send e-mail to the ADC person
            strbldrMessage.Append("You have been designated as the ADC reviewer for the subject survey template by ")
            strbldrMessage.Append(strByName)
            strbldrMessage.Append("<p/>You may review the survey template at the following link and either accept or reject the survey template.")
            strbldrMessage.Append("<br/>If you reject the survey template you will be given the opportunity to explain the reason to the requestor.")
            strbldrMessage.Append("<br/>Survey Template Preview Link: " & _
                "<a href='http://" & strServerName & "/surveyadmin/preview.aspx?seqTemplate=" & Session("seqTemplate") & "&isADCHSR=True&adcHID=" & strHID & "&isAuthenticated=True'>")
            strbldrMessage.Append(Session("TemplateName") & "</a>")
            strbldrMessage.Append("<br/>Please turn off your pop-up blockers before attempting to review this template.  It may interfere with the preview window.")
            strbldrMessage.Append("<p></p>" & strInformation)
        ElseIf (blnPriv = True) Then
            'send e-mail to the Privacy person
            strbldrMessage.Append("You have been designated as the Privacy reviewer for the subject survey template by ")
            strbldrMessage.Append(strByName)
            strbldrMessage.Append("<p/>You may review the survey template at the following link and either accept or reject the survey template.")
            strbldrMessage.Append("<br/>If you reject the survey template you will be given the opportunity to explain the reason to the requestor.")
            strbldrMessage.Append("<br/>Survey Template Preview Link: " & _
                "<a href='http://" & strServerName & "/surveyadmin/preview.aspx?seqTemplate=" & Session("seqTemplate") & "&isADCHSR=True&privHID=" & strHID & "&isAuthenticated=True'>")
            strbldrMessage.Append(Session("TemplateName") & "</a>")
            strbldrMessage.Append("<br/>Please turn off your pop-up blockers before attempting to review this template.  It may interfere with the preview window.")
            strbldrMessage.Append("<p></p>" & strInformation)
        ElseIf (blnADC = False And blnOldADC = False And blnPriv = False) Then
            'send e-mail to the HSR person
            strbldrMessage.Append("You have been designated as the HSR reviewer for the subject survey template by ")
            strbldrMessage.Append(strByName)
            strbldrMessage.Append("<p/>You may review the survey template at the following link and either accept or reject the survey template.")
            strbldrMessage.Append("<br/>If you reject the survey template you will be given the opportunity to explain the reason to the requestor.")
            strbldrMessage.Append("<br/>Survey Template Preview Link: " & _
                "<a href='http://" & strServerName & "/surveyadmin/preview.aspx?seqTemplate=" & Session("seqTemplate") & "&isADCHSR=True&hsrHID=" & strHID & "&isAuthenticated=True'>")
            strbldrMessage.Append(Session("TemplateName") & "</a>")
            strbldrMessage.Append("<br/>Please turn off your pop-up blockers before attempting to review this template.  It may interfere with the preview window.")
            strbldrMessage.Append("<p></p>" & strInformation)
        ElseIf (blnOldADC = True) Then
            'send e-mail to the old ADC person
            strbldrSubject.Remove(0, strbldrSubject.Length())
            strbldrSubject.Append("Another ADC reviewer has been chosesn for the " & _
                CType(Session("TemplateName"), String) & _
                " survey template.")
            strbldrMessage.Append("You have been removed as the ADC reviewer for the subject survey template by ")
            strbldrMessage.Append(strByName)
            strbldrMessage.Append("  You may contact this person at this e-mail address: <a href='mailto:")
            strbldrMessage.Append(dsCommon.Tables("ModifyID").Rows(0).Item("EMAIL_ADDRESS") & _
                "'>" & dsCommon.Tables("ModifyID").Rows(0).Item("EMAIL_ADDRESS") & "</a><br/>")
            strbldrMessage.Append("<p></p>" & strInformation)
        End If

        Try
            ' Create a MailMessage object
            Dim objMail As Mail.SmtpClient = New Mail.SmtpClient
            Dim objMessage As Mail.MailMessage = New Mail.MailMessage
            Dim objAddress As Mail.MailAddress = New Mail.MailAddress(dsCommon.Tables("ModifyID").Rows(0).Item("EMAIL_ADDRESS"))
            objMail.Host = "mailhost.pnl.gov"
            objMail.Port = 25
            objMessage.From = objAddress
            objMessage.To.Add(strTo)
            objMessage.Subject = strbldrSubject.ToString
            objMessage.IsBodyHtml = True
            objMessage.Body = strbldrMessage.ToString

            ' Send the email
            objMail.Send(objMessage)
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'notifies the ADC, HSR, & Privacy representatives
    Public Function notify(ByVal sqlCommonAction As SqlClient.SqlCommand, ByVal sqlCommonAdapter As SqlClient.SqlDataAdapter, ByVal dsCommon As DataSet, ByVal strName As String, ByVal strEmail As String, ByVal strHID As String, ByVal AddtlInfo As String, ByVal strServerName As String) As Boolean
        Dim strADCName As String = strName

        'send the mail
        If (sendMail(strEmail, strADCName, strHID, AddtlInfo, sqlCommonAction, sqlCommonAdapter, dsCommon, strServerName, False, True)) Then
            'get the HSR's Information
            Try
                sqlCommonAction.CommandText = "select distinct(php.hanford_id, php.internet_email_address, php.first_name, " & _
                                  " php.last_name, php.middle_initial) " & _
                                  " from opwhse.dbo.vw_pub_rbac_role_all_members prra, " & _
                                  " opwhse.dbo.vw_pub_hanford_people php where prra.child_role_name = " & _
                                  " 'HumSub.Humn Subj Cont' and prra.hanford_id = php.hanford_id"
                'sqlCommonAction.CommandText = "Select * from fempUsers where strEmail = 'jesse.sharp@pnl.gov'"
                sqlCommonAdapter.SelectCommand = sqlCommonAction
                sqlCommonAdapter.Fill(dsCommon, "hsrInfo")
                Dim HSRHID As String = dsCommon.Tables("hsrInfo").Rows(0).Item("hanford_id")
                Dim HSREmailAddress As String = dsCommon.Tables("hsrInfo").Rows(0).Item("internet_email_address")
                Dim HSRFirstName As String = dsCommon.Tables("hsrInfo").Rows(0).Item("first_name")
                Dim HSRLastName As String = dsCommon.Tables("hsrInfo").Rows(0).Item("last_name")
                Dim HSRMiddleName As String = dsCommon.Tables("hsrInfo").Rows(0).Item("middle_initial")

                If (sendMail(HSREmailAddress, HSRLastName & ", " & HSRFirstName & " " & HSRMiddleName, HSRHID, AddtlInfo, sqlCommonAction, sqlCommonAdapter, dsCommon, strServerName)) Then
                    If (sendMail("alan.rither@pnl.gov", "Rither, Alan C", "0010720", AddtlInfo, sqlCommonAction, sqlCommonAdapter, dsCommon, strServerName, False, False, True)) Then
                        Return True
                    Else
                        Session("Alert") = "Failed to notify the Privacy representative via e-mail.  Your request has been temporarily denied.  Please contact the Survey Tool Administrator and try again."
                    End If
                Else
                    Session("Alert") = "Failed to notify the HSR & Privacy representatives via e-mail.  Your request has been temporarily denied.  Please contact the Survey Tool Administrator and try again."
                    Return False
                End If
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                Return False
            End Try
        Else
            Session("Alert") = "Failed to notify the ADC, HSR, & Privacy representatives via e-mail.  Your request has been temporarily denied.  Please contact the Survey Tool Administrator and try again."
            Return False
        End If
    End Function

    'notifies the old ADC and new ADC of the change in ADCs
    Public Sub notifyOldADC(ByVal sqlCommonAction As SqlClient.SqlCommand, ByVal sqlCommonAdapter As SqlClient.SqlDataAdapter, ByVal dsCommon As DataSet, ByVal AddtlInfo As String, ByVal strServerName As String)
        If Not (sendMail(dsCommon.Tables("Current").Rows(0).Item("EMAIL_ADDRESS"), _
                    dsCommon.Tables("Current").Rows(0).Item("NAME"), _
                    dsCommon.Tables("Current").Rows(0).Item("HANFORD_ID"), _
                    AddtlInfo, sqlCommonAction, sqlCommonAdapter, dsCommon, strServerName, True)) Then
            Session("Alert") = "Failed to notify the old ADC via e-mail.  Please send the old ADC (" & dsCommon.Tables("Signers").Rows(0).Item("EMAIL_ADDRESS") & ") a courtesy e-mail."
        End If
    End Sub

    'generic send mail function
    Public Function genericSendMail(ByVal strFrom As String, ByVal strTo As String, ByVal strSubject As String, ByVal strMessage As String, Optional ByVal strCC As String = "") As Boolean
        ' Create a MailMessage object
        Dim objMail As Mail.SmtpClient = New Mail.SmtpClient
        Dim objMessage As Mail.MailMessage = New Mail.MailMessage
        Dim objAddress As Mail.MailAddress = New Mail.MailAddress(strFrom)
        objMail.Host = "mailhost.pnl.gov"
        objMail.Port = 25
        objMessage.From = objAddress

        ' Create a MailMessage object
        Dim toList As Array = Split(strTo, ",")
        Dim index As Integer = 0
        While (index < toList.Length)
            'remove this
            strMessage &= "<br/>To list:<br/>" & toList(index)
            'end remove this
            'activate this
            'objMail.AddRecipient(toList(index), toList(index))
            'end activate this
            index += 1
        End While
        'remove this
        objMessage.To.Add("survey.administrator@pnl.gov")
        objMessage.To.Add("jesse.sharp@pnl.gov")
        'end remove this

        Dim ccList As Array = Split(strCC, ",")
        index = 0
        While (index < ccList.Length)
            'remove this
            strMessage &= "<br/>CC List:<br/>" & ccList(index)
            'end remove this
            'activate this
            'objMail.AddCC(ccList(index), ccList(index))
            'end activate this
            index += 1
        End While

        objMessage.Subject = strSubject.ToString
        objMessage.IsBodyHtml = True
        objMessage.Body = strMessage.ToString

        ' Send the email
        Try
            objMail.Send(objMessage)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
#End Region

#Region "Password"
    'hashes the input password text
    Public Function hashPassword(ByVal strTempPassword As String, Optional ByVal strUserName As String = "") As String
        Dim strPassword As String = ""
        Try
            'if the user name is blank then use the session for the current user
            If (strUserName = "") Then
                strUserName = Session("UserName")
            End If

            Dim proxyAuthentication As New Authentication.Service1
            'Dim myCred As New NetworkCredential("PNL\elearning", "Cay34bin")
            'Dim myCache As New CredentialCache
            'myCache.Add(New Uri("http://traindev.pnl.gov/services/authentication/"), "Basic", myCred)
            'proxyAuthentication.Credentials = myCred
            strPassword = proxyAuthentication.SHA256(strTempPassword, strUserName)
            Return strPassword
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return strPassword
        End Try
    End Function

    'returns a password string
    Public Function createPasswordString() As String
        Dim strPassword As String = ""
        Try
            Dim proxyAuthentication As New Authentication.Service1
            strPassword = proxyAuthentication.createPasswordString()
            Return strPassword
        Catch ex As Exception
            Return strPassword
        End Try
    End Function

    'Determines if the user's password is valid
    Public Function checkPassword(ByVal strEnteredPassword As String, ByVal dsCommon As DataSet, Optional ByVal strUserName As String = "") As Boolean
        Trace.Warn("Checking Password")
        Dim strPassword As String = ""

        'set the default user name to the current user if not supplied
        If (strUserName = "") Then
            strUserName = Session("UserName")
        End If

        strPassword = hashPassword(strEnteredPassword, strUserName)
        If (strPassword <> "") Then
            'compare the result to the DB and return true if they match
            If (strPassword = dsCommon.Tables("Users").Rows(0).Item("AUTHENTICATION_CODE")) Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    'Determines if the user's password is valid
    Public Function isExpired(ByVal userID As String, ByVal sqlConn As SqlClient.SqlConnection) As Boolean
        Trace.Warn("Checking Password Expiration")
        Try
            Dim sqlCommonAdapter As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter
            Dim sqlCommonAction As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand
            sqlCommonAction.Connection = sqlConn
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            Dim dsCommon As New DataSet

            'Open the connection
            If (sqlCommonAction.Connection.State = ConnectionState.Closed) Then
                sqlCommonAction.Connection.Open()
            End If

            'get the date records for the user
            sqlCommonAction.CommandText = "Select LAST_CHANGE_AUTH_CD_DATE from " & getExtension() & "SAT_USER where " & _
            "USER_KEY = " & userID
            sqlCommonAdapter.Fill(dsCommon, "AuthenticationDate")

            Dim dr As DataRow
            For Each dr In dsCommon.Tables("AuthenticationDate").Rows
                If (CType(dr.Item(0), DateTime).AddMonths(6) <= Now) Then
                    Return True
                Else
                    Return False
                End If
            Next
        Catch ex As Exception
            Session("expirationFailure") = "True"
            Session("Alert") = "Failed to determine if your password has expired or not."
            Return False
        End Try
    End Function
#End Region

#Region "User"
    'check the user type
    Public Function checkUserType(ByVal userTypeMin As Integer, Optional ByVal isOwner As Boolean = True, Optional ByVal isDelegate As Boolean = True, Optional ByVal isOr As Boolean = False) As Boolean
        If (isOr) Then
            If ((Session("UserType") < userTypeMin Or Not (isOwner Or isDelegate)) And Session("UserType") <> 4) Then
                Session("Alert") = "You are not authorized to be here."
                Return True
            Else
                Return False
            End If
        Else
            If ((Session("UserType") < userTypeMin Or Not (isOwner And isDelegate)) And Session("UserType") <> 4) Then
                Session("Alert") = "You are not authorized to be here."
                Return True
            Else
                Return False
            End If
        End If
    End Function

    'get user ownership status for a template or survey
    Public Sub determineUserOwnership(ByVal sqlCommonAction As SqlClient.SqlCommand, ByVal sqlCommonAdapter As SqlClient.SqlDataAdapter, Optional ByVal isTemplate As Boolean = False)
        'get the user information
        If ((isTemplate And Session("seqTemplate") <> -1) Or (Not isTemplate And Session("seqSurvey") <> -1)) Then
            Try
                If (isTemplate) Then
                    sqlCommonAction.CommandText = "Select * from " & getExtension() & "SAT_USER_PERMISSION fp, " & _
                    "" & getExtension() & "SAT_USER fu where TEMPLATE_KEY = " & Session("seqTemplate") & _
                    " and fp.USER_KEY = fu.USER_KEY " & _
                    "order by fu.LAST_NAME, fu.FIRST_NAME"
                Else
                    sqlCommonAction.CommandText = "Select * from " & getExtension() & "SAT_USER_PERMISSION fp, " & _
                    "" & getExtension() & "SAT_USER fu where SURVEY_KEY = " & Session("seqSurvey") & _
                    " and fp.USER_KEY = fu.USER_KEY " & _
                    "order by fu.LAST_NAME, fu.FIRST_NAME"
                End If

                'fill the dataset and proccess the results
                Dim ds As DataSet = New DataSet
                sqlCommonAdapter.Fill(ds, "OwnerDelegate")
                Dim index As Integer = 0
                While (index < ds.Tables("OwnerDelegate").Rows.Count())
                    If (ds.Tables("OwnerDelegate").Rows(index).Item("USER_KEY") = Session("UserID")) Then
                        If (isTemplate) Then
                            Session("isTemplateOwner") = IIf(ds.Tables("OwnerDelegate").Rows(index).Item("OWNER_IND"), 1, 0)
                            Session("isTemplateDelegate") = IIf(ds.Tables("OwnerDelegate").Rows(index).Item("DELEGATE_IND"), 1, 0)
                            Session("isTemplateUser") = IIf(ds.Tables("OwnerDelegate").Rows(index).Item("USER_IND"), 1, 0)
                        Else
                            Session("isSurveyOwner") = IIf(ds.Tables("OwnerDelegate").Rows(index).Item("OWNER_IND"), 1, 0)
                            Session("isSurveyDelegate") = IIf(ds.Tables("OwnerDelegate").Rows(index).Item("DELEGATE_IND"), 1, 0)
                            Session("isSurveyUser") = IIf(ds.Tables("OwnerDelegate").Rows(index).Item("USER_IND"), 1, 0)
                        End If
                        index = ds.Tables("OwnerDelegate").Rows.Count() + 10
                    Else
                        If (isTemplate) Then
                            Session("isTemplateOwner") = False
                            Session("isTemplateDelegate") = False
                            Session("isTemplateUser") = False
                        Else
                            Session("isSurveyOwner") = False
                            Session("isSurveyDelegate") = False
                            Session("isSurveyuser") = False
                        End If
                        index += 1
                    End If
                End While
            Catch ex As Exception
                Trace.Warn(ex.ToString)
                Session("Alert") = "Unable to determine your access status at this time."
            End Try
        Else
            If (isTemplate) Then
                Session("isTemplateOwner") = True
                Session("isTemplateDelegate") = True
                Session("isTemplateUser") = True
            Else
                Session("isSurveyOwner") = True
                Session("isSurveyDelegate") = True
                Session("isSurveyUser") = True
            End If
        End If
    End Sub

    'Determines if there are any users that match the supplied user name
    Public Function findUserName(ByVal sqlCommonAction As SqlClient.SqlCommand, ByVal sqlCommonAdapter As SqlClient.SqlDataAdapter, ByVal dsCommon As DataSet) As Boolean
        Trace.Warn("Finding User")
        Try
            If (CType(Session("UserName"), String).LastIndexOf("@") >= 0) Then
                sqlCommonAction.CommandText &= " where EMAIL_ADDRESS = '" & Session("UserName") & "'"
            Else
                sqlCommonAction.CommandText &= " where HANFORD_ID = '" & Session("UserName") & "' and ACTIVE_IND = 1"
            End If

            'get the proper extension
            sqlCommonAction.CommandText = sqlCommonAction.CommandText.Replace("VARCONN.", getExtension())
            'Trace.Write(sqlCommonAction.CommandText)
            'Check to see if the user was found
            sqlCommonAdapter.Fill(dsCommon, "Users")
            If (dsCommon.Tables("Users").Rows.Count() = 1) Then
                Return True
            End If
            Return False
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function
#End Region

#Region "Code"
    'checks for forbidden code
    Public Function goodInput(ByRef input As TextBox, Optional ByVal forDisplay As Boolean = False) As Boolean
        'create temporary lower case version of input
        Dim lowerInput As String = input.Text.ToLower

        'check for script input
        While (lowerInput.IndexOf("<script") <> -1)
            If (forDisplay) Then
                Dim lowIndex As Integer = lowerInput.IndexOf("<script")
                lowerInput = lowerInput.Remove(lowIndex, 7)
                lowerInput = lowerInput.Insert(lowIndex, "&lt;script")
                Dim indexGT As Integer = 0
                Dim indexLT As Integer = 0
                indexGT = lowerInput.IndexOf(">", lowIndex)
                If (indexGT > lowIndex) Then
                    lowerInput = lowerInput.Remove(indexGT, 1)
                    lowerInput = lowerInput.Insert(indexGT, "&gt;")
                End If
                lowerInput = lowerInput.Replace("</script>", "&lt;/script&gt;")
                input.Text = lowerInput
            Else
                Return False
            End If
        End While

        'check for on* methods of html controls
        Dim subString As String = input.Text.ToLower
        Dim startIndex As Integer = 0
        While (subString.IndexOf("on", startIndex) <> -1)
            startIndex = subString.IndexOf("on", startIndex)
            If (subString.IndexOf("=", startIndex) <> -1) Then
                Dim endIndex As Integer = subString.IndexOf("=", startIndex)
                Dim tempString As String
                tempString = subString.Substring(startIndex, endIndex - startIndex + 1)
                Dim character As Char
                Dim spaceCount As Integer = 0
                Dim SpaceCounted As Boolean = False
                For Each character In tempString
                    If (character = " ") Then
                        If Not (SpaceCounted) Then
                            spaceCount += 1
                        End If
                        SpaceCounted = True
                    Else
                        SpaceCounted = False
                    End If
                Next
                Dim offset As Integer = 0
                If (spaceCount <= 1) Then
                    tempString = LTrim(subString.Substring(endIndex + 1, subString.Length - endIndex - 1))
                    offset = subString.Substring(endIndex + 1, subString.Length - endIndex - 1).Length - tempString.Length
                    If (forDisplay) Then
                        If (tempString.Chars(0) = """") Then
                            endIndex = subString.IndexOf("""", endIndex + 2 + offset)
                            If (endIndex <> -1) Then
                                subString = subString.Remove(startIndex, endIndex - startIndex + 1)
                            Else
                                subString = subString.Remove(startIndex, subString.Length - startIndex)
                            End If
                            subString = subString.Insert(startIndex, "(HTML Event Handler Removed)")
                            input.Text = subString
                        ElseIf (tempString.Chars(0) = "'") Then
                            endIndex = subString.IndexOf("'", endIndex + 2 + offset)
                            If (endIndex <> -1) Then
                                subString = subString.Remove(startIndex, endIndex - startIndex + 1)
                            Else
                                subString = subString.Remove(startIndex, subString.Length - startIndex)
                            End If
                            subString = subString.Insert(startIndex, "(HTML Event Handler Removed)")
                            input.Text = subString
                        End If
                    Else
                        Return False
                    End If
                End If
            End If
            startIndex += 1
        End While
        Return True
    End Function
#End Region

#Region "Connection"
    'returns the proper connection for the current machine
    Public Function getConnection(ByVal conn As Data.SqlClient.SqlConnection, Optional ByVal serverName As String = "") As Data.SqlClient.SqlConnection
        'close the existing connection
        closeConnection(conn)

        'conn.ConnectionString = "packet size=4096;user id=tqadm;data source=OLTPDEV02,915;persist security info=True;initial catalog=TQDB_DEV;password=pwd;enlist=false"
        conn.ConnectionString = "Data Source=OLTPDEV02,915;Initial Catalog=TQDB_DEV;Persist Security Info=True;User ID=tqadm;Password=pwd"

        'open the connection
        openConnection(conn)

        Return conn
    End Function

    'returns the proper connection for the current machine
    Public Function getConnection(ByVal conn As Data.OracleClient.OracleConnection, Optional ByVal serverName As String = "") As Data.OracleClient.OracleConnection
        'close the existing connection
        closeConnection(conn)

        If (serverName.ToLower = "hrms.prod.irm") Then
            conn.ConnectionString = "user id=web_tqapps;data source=""hrms.prod.irm"";password=pwd"
        End If

        If (serverName.ToLower = "elm.prod.irm") Then
            conn.ConnectionString = "user id=el_user;data source=""elm.prod.irm"";password=pwd"
        End If

        'open the connection
        openConnection(conn)

        Return conn
    End Function

    'returns a production database prefix
    Public Function getProdExtension(Optional ByVal myRetiredVar As String = "") As String
        'get the connection
        'Return "sqlsrvsecure.tqdb_test_prod.dbo."
        Return "tqdb_test.dbo."
    End Function

    'returns a database prefix based on the machine
    Public Function getExtension(Optional ByVal switch As Boolean = False) As String
        If (switch = False) Then
            If (Session("Machine") = "Production") Then
                'Return "sqlsrvsecure.tqdb_test_prod.dbo."
                Return "tqdb_test.dbo."
            Else
                Return "tqdb_dev.dbo."
            End If
        Else
            If (Session("Machine") = "Production") Then
                Return "tqdb_dev.dbo."
            Else
                'Return "sqlsrvsecure.tqdb_test_prod.dbo."
                Return "tqdb_test.dbo."
            End If
        End If
    End Function

    'returns a development database prefix
    Public Function getDevExtension() As String
        Return "tqdb_dev.dbo."
    End Function

    'swaps the machine variable
    Public Sub swapMachine()
        If (Session("Machine") = "Production") Then
            Session("Machine") = "Development"
        Else
            Session("Machine") = "Production"
        End If
    End Sub

    'closes any open connection
    Private Sub closeConnection(ByVal conn As Data.SqlClient.SqlConnection)
        If (conn.State = ConnectionState.Open) Then
            conn.Close()
        End If
    End Sub

    'opens any closed connection
    Private Sub openConnection(ByVal conn As Data.SqlClient.SqlConnection)
        If (conn.State = ConnectionState.Closed) Then
            conn.Open()
        End If
    End Sub

    'closes any open connection
    Private Sub closeConnection(ByVal conn As Data.OracleClient.OracleConnection)
        If (conn.State = ConnectionState.Open) Then
            conn.Close()
        End If
    End Sub

    'opens any closed connection
    Private Sub openConnection(ByVal conn As Data.OracleClient.OracleConnection)
        If (conn.State = ConnectionState.Closed) Then
            conn.Open()
        End If
    End Sub
#End Region
End Class
