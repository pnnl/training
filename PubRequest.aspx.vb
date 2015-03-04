Imports System.text
Imports System.Collections.Specialized
Public Class PubRequest
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()
        Me.dsCommon = New System.Data.DataSet
        Me.commonConn = New System.Data.SqlClient.SqlConnection
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'dsCommon
        '
        Me.dsCommon.DataSetName = "NewDataSet"
        Me.dsCommon.Locale = New System.Globalization.CultureInfo("en-US")
        CType(Me.dsCommon, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub
    Protected WithEvents imgTitle As System.Web.UI.WebControls.Image
    Protected WithEvents hlpServerSwitch As System.Web.UI.WebControls.Image
    Protected WithEvents chkServerSwitch As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblPlaceHolder As System.Web.UI.WebControls.Label
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents lblHIDPayrollName As System.Web.UI.WebControls.Label
    Protected WithEvents txtHIDPayrollName As System.Web.UI.WebControls.TextBox
    Protected WithEvents btnHIDPayrollNameFind As System.Web.UI.WebControls.Button
    Protected WithEvents lblEmail As System.Web.UI.WebControls.Label
    Protected WithEvents lblFirstName As System.Web.UI.WebControls.Label
    Protected WithEvents lblLastName As System.Web.UI.WebControls.Label
    Protected WithEvents lblMiddleName As System.Web.UI.WebControls.Label
    Protected WithEvents lblPhoneNumber As System.Web.UI.WebControls.Label
    Protected WithEvents dsCommon As System.Data.DataSet
    Protected requestItems As Array
    Protected myUtility As Utility = New Utility
    Protected WithEvents btnSubmit As System.Web.UI.WebControls.Button
    Protected WithEvents btnCancel As System.Web.UI.WebControls.Button
    Protected WithEvents lblDepartment As System.Web.UI.WebControls.Label
    Protected WithEvents hlpHIDPayrollName As System.Web.UI.WebControls.Image
    Protected WithEvents lblEmailDisplay As System.Web.UI.WebControls.Label
    Protected WithEvents lblFirstNameDisplay As System.Web.UI.WebControls.Label
    Protected WithEvents lblLastNameDisplay As System.Web.UI.WebControls.Label
    Protected WithEvents lblMiddleNameDisplay As System.Web.UI.WebControls.Label
    Protected WithEvents lblDepartmentDisplay As System.Web.UI.WebControls.Label
    Protected WithEvents lblPhoneNumberDisplay As System.Web.UI.WebControls.Label
    Protected WithEvents lblIntroduction As System.Web.UI.WebControls.Label
    Protected WithEvents lblIntroduction2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents lblHID As System.Web.UI.WebControls.Label
    Protected WithEvents lblHIDDisplay As System.Web.UI.WebControls.Label
    Protected WithEvents hlpAddtlInfo As System.Web.UI.WebControls.Image
    Protected WithEvents lblAddtlInfo As System.Web.UI.WebControls.Label
    Protected WithEvents txtAddtlInfo As System.Web.UI.WebControls.TextBox
    Protected sqlCommonAdapter As New System.Data.SqlClient.SqlDataAdapter
    Protected sqlCommonAction As New System.Data.SqlClient.SqlCommand
    Protected strNewPassword As String
    Protected WithEvents commonConn As System.Data.SqlClient.SqlConnection
    Protected WithEvents JSAction As String
    Protected WithEvents HSREmailAddress As String
    Protected WithEvents HSRFirstName As String
    Protected WithEvents HSRLastName As String
    Protected WithEvents HSRMiddleName As String
    Protected WithEvents HSRPhone As String
    Protected WithEvents HSRHID As String

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
            Session("BackgroundColor") = "-1"
        End If

        'Put user code to initialize the page here
        If (Session("isAuthenticated") <> "True") Then
            Session("CurrentPage") = "./Login.aspx"
            Session("pageSel") = "Login"
            redirect(Session("CurrentPage"))
        ElseIf (myUtility.checkUserType(3, CType(Session("isTemplateOwner"), Boolean), CType(Session("isTemplateDelegate"), Boolean), True)) Then
            redirect(Session("CurrentPage"))
        Else
            Session("CurrentPage") = "./PubRequest.aspx"
            Session("pageSel") = "Template"
        End If

        'if the user's password has expired send them to the password reset page
        If (Session("isExpired") = "True") Then
            Session("CurrentPage") = "./Password.aspx"
            Session("pageSel") = "Password"
            redirect(Session("CurrentPage"))
        End If

        myUtility.getConnection(Me.commonConn)

        'Check for user selection from the comments list if we're sure it's not a simple post-back
        If Not (IsPostBack) Then
            'get anything we can off the request line
            myUtility.getRequest(Page.Request.Url.Query())

            'if we're coming from the PersonSelect page and the process was not stopped load the selected user's data
            'otherwise load the page like normal.
            If (Session("ProcessStopped") = "False") Then
                'initialize the adhoc data collection items
                sqlCommonAdapter.SelectCommand = sqlCommonAction
                sqlCommonAction.Connection = Me.commonConn

                'restore the person data and find the exact user selected
                Dim dt As DataTable = Session("PersonTable")
                Dim dr As DataRow = dt.Rows(Session("selectedPersonRow"))

                'Populate the controls on the page
                setControls("Override")
            Else
                If (Session("isDirty") And Not Session("isApproved")) Then
                    'get signer data
                    isInApproval()

                    'Populate the controls on the page
                    setControls("Override")
                Else
                    'Populate the controls on the page
                    setControls()
                End If
            End If

            Session("ProcessStopped") = "True"
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
#End Region

#Region "Control Settings"
    'populates the controls with the selected user's data
    Private Sub setControls(Optional ByVal override As String = "")
        Trace.Warn("Setting Controls")
        'set the basic page controls based on whether or not we are working with a new publication request
        If (override <> "") Then
            Try
                'restore the person data and find the exact user selected
                Dim dt As DataTable = Session("PersonTable")
                Dim dr As DataRow = dt.Rows(Session("selectedPersonRow"))

                'populate the controls with the selected individual's information
                Me.lblEmailDisplay.Text = dr.Item("internet_email_address")
                Me.lblFirstNameDisplay.Text = dr.Item("first_name")
                Me.lblLastNameDisplay.Text = dr.Item("last_name")
                Me.lblMiddleNameDisplay.Text = dr.Item("middle_initial")
                Me.lblDepartmentDisplay.Text = dr.Item("department")
                Me.lblPhoneNumberDisplay.Text = dr.Item("phone_no")
                Me.lblHIDDisplay.Text = dr.Item("hanford_id")
                If (Session("ProcessStopped") = "False") Then
                    Me.txtAddtlInfo.Text = Session("AddtlInfo")
                Else
                    Me.txtAddtlInfo.Text = dr.Item("strInformation")
                End If
            Catch ex As Exception
                Trace.Warn(ex.ToString)
            End Try
        Else
            If (override = "") Then
                Me.lblEmailDisplay.Text = ""
                Me.lblFirstNameDisplay.Text = ""
                Me.lblLastNameDisplay.Text = ""
                Me.lblMiddleNameDisplay.Text = ""
                Me.lblDepartmentDisplay.Text = ""
                Me.lblPhoneNumberDisplay.Text = ""
                Me.lblHIDDisplay.Text = ""
                Me.txtAddtlInfo.Text = ""
            End If
        End If
    End Sub
#End Region

#Region "Checks"
    'Determines if the template is in approval and load the data from the table
    Private Function isInApproval() As Boolean
        Try
            'Set up the common data adapter and delete command
            sqlCommonAction.Connection = Me.commonConn
            sqlCommonAdapter.SelectCommand = sqlCommonAction

            'Open the connection
            If (sqlCommonAction.Connection.State = ConnectionState.Closed) Then
                sqlCommonAction.Connection.Open()
            End If

            'get the current submittal
            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE " & _
                "Where TEMPLATE_KEY = " & Session("seqTemplate") & " and CURRENT_IND = 1"
            sqlCommonAdapter.Fill(Me.dsCommon, "Current")

            If (Me.dsCommon.Tables("Current").Rows.Count() > 0) Then
                'get the records for that date if they exist
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE where " & _
                    "TEMPLATE_KEY = " & Session("seqTemplate") & " and CURRENT_IND = 1" & _
                    " order by AUTH_DERIV_CLASSIFIER_IND Desc"
                sqlCommonAdapter.Fill(Me.dsCommon, "Signed")
            Else
                Return False
            End If

            'populate the user information
            If (Me.dsCommon.Tables("Signed").Rows.Count() = 3) Then
                If (getHanfordPeopleInfo(Me.dsCommon.Tables("Signed").Rows(0).Item("HANFORD_ID"))) Then
                    'populate the controls with the selected individual's information
                    Me.lblEmailDisplay.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("internet_email_address")
                    Me.lblFirstNameDisplay.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("first_name")
                    Me.lblLastNameDisplay.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("last_name")
                    Me.lblMiddleNameDisplay.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("middle_initial")
                    Me.lblDepartmentDisplay.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("department")
                    Me.lblPhoneNumberDisplay.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("phone_no")
                    Me.lblHIDDisplay.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("hanford_id")
                    Me.txtAddtlInfo.Text = Me.dsCommon.Tables("Signed").Rows(0).Item("AUTHOR_COMMENT")

                    'Populate the controls on the page
                    setControls("override")
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function
#End Region

#Region "Find"
    'gets the hanford people information for the user we're looking for, returns true if it finds the person(s)
    Private Function getHanfordPeopleInfo(ByVal LookUpString As String) As Boolean
        'initialize the adhoc data collection items
        sqlCommonAdapter.SelectCommand = sqlCommonAction
        sqlCommonAction.Connection = Me.commonConn

        If ((LookUpString.ToUpper.StartsWith("H") And IsNumeric(LookUpString.Substring(1))) Or _
            (LookUpString.Length = 7 And IsNumeric(LookUpString))) Then
            'Process the apparent HID
            Return (processHID(LookUpString))
        ElseIf ((LookUpString.ToUpper.StartsWith("D") And LookUpString.Length = 6) Or _
            LookUpString.Length = 5) Then
            If (IsNumeric(LookUpString.Substring(LookUpString.Length - 3))) Then
                'Process the apparent payroll
                Return (processPayroll(LookUpString))
            Else
                'Process the apparent last name
                Return (processName(LookUpString))
            End If
        Else
            'Process the apparent last name
            Return (processName(LookUpString))
        End If
    End Function

    'Process the HID
    Private Function processHID(ByVal LookUpString As String) As Boolean
        Try
            'strip the H from the front end if it exists
            If (LookUpString.ToUpper.StartsWith("H")) Then
                LookUpString = LookUpString.Substring(1)
            End If

            'set up and process the sql statement
            sqlCommonAction.CommandText = "select * from (select isnull(max(hanford_id), '') " & _
            "as hanford_id, isnull(max(hanf_pay_no), '') as hanf_pay_no, max(first_name) as " & _
            "first_name, max(last_name) as last_name, isnull(max(middle_initial), '') as " & _
            "middle_initial, isnull(max(name_prefix), '') as name_prefix, " & _
            "isnull(max(name_suffix), '') as name_suffix, isnull(max(phone_no), '') as " & _
            "phone_no, isnull(max(internet_email_address), '') as internet_email_address, " & _
            "isnull(max(employed_by_hcid), '') as employed_by_hcid, isNull(max(active_sw), 'N') " & _
            "as active_sw, max(last_name) + ', ' + max(first_name) + ' ' + " & _
            "isNull(max(middle_initial), '') + ' - ' + isNull(max(internet_email_address), '') " & _
            "as NameEmail, isNull(max(department), '') as department from (select hanford_id, " & _
            "hanf_pay_no, first_name, last_name, middle_initial, NULL as name_prefix, " & _
            "name_suffix, phone_no, internet_email_address, employed_by_hcid, active_sw, NULL " & _
            "as department from opwhse.dbo.vw_pub_hanford_people union select hanford_id, " & _
            "pay_no as hanf_pay_no, first_name, last_name, middle_name as middle_initial, NULL " & _
            "as name_prefix, name_suffix, work_phone as phone_no, internet_address as " & _
            "internet_email_address, employed_by_hcid, NULL as active_sw, NULL as department " & _
            "from opwhse.dbo.vw_pub_pnnl_associate union select hanford_id, pnnl_pay_no as " & _
            "hanf_pay_no, first_name, last_name, middle_initial, name_prefix, name_suffix, " & _
            "contact_work_phone as phone_no, internet_email_address, company as employed_by_hcid, " & _
            "NULL as active_sw, dept_cd as department from opwhse.dbo.vw_pub_bmi_employee) a " & _
            "where hanford_id = '" & LookUpString & "' group by hanford_id) b " & _
            "where active_sw = 'Y' and department <> '' order by last_name, first_name"
            sqlCommonAdapter.Fill(Me.dsCommon, "PersonData")

            'if we found something return true otherwise return false
            If (Me.dsCommon.Tables("PersonData").Rows.Count >= 1) Then
                Session("PersonTable") = Me.dsCommon.Tables("PersonData")
                Return True
            Else
                Session("PersonTable") = Nothing
                Return False
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'Process the Payroll
    Private Function processPayroll(ByVal LookUpString As String) As Boolean
        Try
            'strip the D from the front end if it exists
            If (LookUpString.ToUpper.StartsWith("D")) Then
                LookUpString = LookUpString.Substring(1)
            End If
            'set up and process the sql statement
            sqlCommonAction.CommandText = "select * from (select isnull(max(hanford_id), '') " & _
            "as hanford_id, isnull(max(hanf_pay_no), '') as hanf_pay_no, max(first_name) as " & _
            "first_name, max(last_name) as last_name, isnull(max(middle_initial), '') as " & _
            "middle_initial, isnull(max(name_prefix), '') as name_prefix, " & _
            "isnull(max(name_suffix), '') as name_suffix, isnull(max(phone_no), '') as " & _
            "phone_no, isnull(max(internet_email_address), '') as internet_email_address, " & _
            "isnull(max(employed_by_hcid), '') as employed_by_hcid, isNull(max(active_sw), 'N') " & _
            "as active_sw, max(last_name) + ', ' + max(first_name) + ' ' + " & _
            "isNull(max(middle_initial), '') + ' - ' + isNull(max(internet_email_address), '') " & _
            "as NameEmail, isNull(max(department), '') as department from (select hanford_id, " & _
            "hanf_pay_no, first_name, last_name, middle_initial, NULL as name_prefix, " & _
            "name_suffix, phone_no, internet_email_address, employed_by_hcid, active_sw, NULL " & _
            "as department from opwhse.dbo.vw_pub_hanford_people union select hanford_id, " & _
            "pay_no as hanf_pay_no, first_name, last_name, middle_name as middle_initial, NULL " & _
            "as name_prefix, name_suffix, work_phone as phone_no, internet_address as " & _
            "internet_email_address, employed_by_hcid, NULL as active_sw, NULL as department " & _
            "from opwhse.dbo.vw_pub_pnnl_associate union select hanford_id, pnnl_pay_no as " & _
            "hanf_pay_no, first_name, last_name, middle_initial, name_prefix, name_suffix, " & _
            "contact_work_phone as phone_no, internet_email_address, company as employed_by_hcid, " & _
            "NULL as active_sw, dept_cd as department from opwhse.dbo.vw_pub_bmi_employee) a " & _
            "where hanf_pay_no = '" & LookUpString & "' group by hanford_id) b " & _
            "where active_sw = 'Y' and department <> '' order by last_name, first_name"
            sqlCommonAdapter.Fill(Me.dsCommon, "PersonData")

            'if we found something return true otherwise return false
            If (Me.dsCommon.Tables("PersonData").Rows.Count >= 1) Then
                Session("PersonTable") = Me.dsCommon.Tables("PersonData")
                Return True
            Else
                Session("PersonTable") = Nothing
                Return False
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'Process the Name
    Private Function processName(ByVal LookUpString As String) As Boolean
        Try
            'fix any apostrophes that might be in the search string
            LookUpString = LookUpString.Replace("'", "''")

            'set up and process the sql statement
            sqlCommonAction.CommandText = "select * from (select isnull(max(hanford_id), '') " & _
            "as hanford_id, isnull(max(hanf_pay_no), '') as hanf_pay_no, max(first_name) as " & _
            "first_name, max(last_name) as last_name, isnull(max(middle_initial), '') as " & _
            "middle_initial, isnull(max(name_prefix), '') as name_prefix, " & _
            "isnull(max(name_suffix), '') as name_suffix, isnull(max(phone_no), '') as " & _
            "phone_no, isnull(max(internet_email_address), '') as internet_email_address, " & _
            "isnull(max(employed_by_hcid), '') as employed_by_hcid, isNull(max(active_sw), 'N') " & _
            "as active_sw, max(last_name) + ', ' + max(first_name) + ' ' + " & _
            "isNull(max(middle_initial), '') + ' - ' + isNull(max(internet_email_address), '') " & _
            "as NameEmail, isNull(max(department), '') as department from (select hanford_id, " & _
            "hanf_pay_no, first_name, last_name, middle_initial, NULL as name_prefix, " & _
            "name_suffix, phone_no, internet_email_address, employed_by_hcid, active_sw, NULL " & _
            "as department from opwhse.dbo.vw_pub_hanford_people union select hanford_id, " & _
            "pay_no as hanf_pay_no, first_name, last_name, middle_name as middle_initial, NULL " & _
            "as name_prefix, name_suffix, work_phone as phone_no, internet_address as " & _
            "internet_email_address, employed_by_hcid, NULL as active_sw, NULL as department " & _
            "from opwhse.dbo.vw_pub_pnnl_associate union select hanford_id, pnnl_pay_no as " & _
            "hanf_pay_no, first_name, last_name, middle_initial, name_prefix, name_suffix, " & _
            "contact_work_phone as phone_no, internet_email_address, company as employed_by_hcid, " & _
            "NULL as active_sw, dept_cd as department from opwhse.dbo.vw_pub_bmi_employee) a " & _
            "where last_name like '" & LookUpString & "%' group by hanford_id) b " & _
            "where active_sw = 'Y' and department <> '' order by last_name, first_name"
            sqlCommonAdapter.Fill(Me.dsCommon, "PersonData")

            'if we found something return true otherwise return false
            If (Me.dsCommon.Tables("PersonData").Rows.Count = 1) Then
                Session("PersonTable") = Me.dsCommon.Tables("PersonData")
                Session("selectedPersonRow") = 0
                Return True
            ElseIf (Me.dsCommon.Tables("PersonData").Rows.Count > 1) Then
                Session("PersonTable") = Me.dsCommon.Tables("PersonData")
                Session("AddtlInfo") = Me.txtAddtlInfo.Text
                Session("CurrentPage") = "./PersonSelect.aspx"
                Session("pageSel") = "template"
                redirect(Session("CurrentPage"))
            Else
                Session("PersonTable") = Nothing
                Return False
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function
#End Region

#Region "Page Action"
    'attempts to locate the user's information in the opwhse under the Hanford People view
    Private Sub btnHIDPayrollNameFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHIDPayrollNameFind.Click
        Session("JSAction") = ""

        If (myUtility.goodInput(Me.txtHIDPayrollName)) Then
            'process the input only if the user actually put something in the find text field
            If (Me.txtHIDPayrollName.Text <> "") Then
                'find the user information we'll need to load
                If (getHanfordPeopleInfo(Me.txtHIDPayrollName.Text)) Then
                    'populate the controls with the selected individual's information
                    Me.lblEmailDisplay.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("internet_email_address")
                    Me.lblFirstNameDisplay.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("first_name")
                    Me.lblLastNameDisplay.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("last_name")
                    Me.lblMiddleNameDisplay.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("middle_initial")
                    Me.lblDepartmentDisplay.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("department")
                    Me.lblPhoneNumberDisplay.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("phone_no")
                    Me.lblHIDDisplay.Text = Me.dsCommon.Tables("PersonData").Rows(0).Item("hanford_id")
                    If Not (Session("AddtlInfo") Is Nothing) Then
                        Me.txtAddtlInfo.Text = Session("AddtlInfo")
                    End If

                    'Populate the controls on the page
                    setControls("override")
                Else
                    alert("This person was not found.")
                End If
            Else
                alert("You must input some text into the HID, Payroll, or Last Name field to use the find functionality.")
            End If
        Else
            alert("You may have entered malicious code into the HID/Emplid/Last Name text field.  Find aborted.")
        End If
    End Sub

    'sets up the survey signed table to include the ADC and HSR representatives that need
    'to sign off on the survey template.
    Private Sub btnSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Session("JSAction") = ""

        Dim blnContinue As Boolean = True

        'check for possible malicious code entries
        If Not (myUtility.goodInput(Me.txtAddtlInfo, True)) Then
            blnContinue = False
            alert("Possible malicious code entry(s).  Publication request aborted.")
        End If

        If (blnContinue) Then
            If (doInsert()) Then
                Session("AddltInfo") = ""
                Session("CurrentPage") = "./Template.aspx"
                Session("pageSel") = "template"
                If (Session("Alert") Is Nothing Or Session("Alert") = "") Then
                    Session("Alert") = "You have successfully submitted this template for review."
                End If
                redirect(Session("CurrentPage"))
            End If
        End If
    End Sub

    'returns the user to the template they came from
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Session("JSAction") = ""
        Session("CurrentPage") = "./Template.aspx"
        Session("pageSel") = "template"
        redirect(Session("CurrentPage"))
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
                If (Me.lblFirstNameDisplay.Text <> "" And Me.lblLastNameDisplay.Text <> "" And Me.lblEmailDisplay.Text <> "" And Me.lblHIDDisplay.Text <> "") Then
                    'If we're inserting a new record then add it onto the end of the recordset and add it to the existing Delegates
                    If (insertUser()) Then
                        If (insertPermission()) Then
                            If (insertUser(True)) Then
                                If (insertPermission(True)) Then
                                    If (insertUser(False, True)) Then
                                        If (insertPermission(False, True)) Then
                                            sqlCommonAdapter.SelectCommand = sqlCommonAction

                                            'get the current signature record for the template
                                            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE " & _
                                                "Where TEMPLATE_KEY = " & Session("seqTemplate") & " and CURRENT_IND = 1" & _
                                                " order by AUTH_DERIV_CLASSIFIER_IND Desc"
                                            sqlCommonAdapter.Fill(Me.dsCommon, "Current")

                                            'parameterized the text input to allow for things like quotes to get through
                                            Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@AUTHOR_COMMENT", System.Data.SqlDbType.VarChar, 1800, "AUTHOR_COMMENT")
                                            param0.Value = Me.txtAddtlInfo.Text
                                            sqlCommonAction.Parameters.Add(param0)

                                            'determine if the records are fully signed off
                                            Dim row As DataRow
                                            Dim intSignCount As Integer = 0
                                            For Each row In Me.dsCommon.Tables("Current").Rows()
                                                If (row.Item("SIGNED_IND")) Then
                                                    intSignCount += 1
                                                End If
                                            Next

                                            'create two new records for signature
                                            'Insert the ADC person first
                                            If Not (enterADC()) Then
                                                sqlCommonAction.Transaction.Rollback()
                                                Return False
                                            End If

                                            'Insert the HSR person next
                                            If Not (enterHSR()) Then
                                                sqlCommonAction.Transaction.Rollback()
                                                Return False
                                            End If

                                            'Insert the HSR person next
                                            If Not (enterPriv()) Then
                                                sqlCommonAction.Transaction.Rollback()
                                                Return False
                                            End If

                                            'sets the last set of current signatures to be non-current
                                            'if it has been signed off already so that the new set
                                            'takes precedence.
                                            If (intSignCount = 3) Then
                                                If Not (removeOldCurrentSignatures()) Then
                                                    sqlCommonAction.Transaction.Rollback()
                                                    Return False
                                                End If
                                            End If

                                            'notify the ADC, HSR, & Privacy representatives of the submission
                                            'generate e-mails for inserted individuals
                                            Dim strName As String = generateName()
                                            If Not (myUtility.notify(sqlCommonAction, sqlCommonAdapter, Me.dsCommon, strName, Me.lblEmailDisplay.Text, Me.lblHIDDisplay.Text, Me.txtAddtlInfo.Text, Request.ServerVariables.Item("SERVER_NAME"))) Then
                                                sqlCommonAction.Transaction.Rollback()
                                                Return False
                                            End If

                                            'if we made it this far then commit the transaction
                                            sqlCommonAction.Transaction.Commit()
                                            Return True
                                        Else
                                            sqlCommonAction.Transaction.Rollback()
                                            alert("Failed to submit this template for review.")
                                            failure = False
                                        End If
                                    Else
                                        sqlCommonAction.Transaction.Rollback()
                                        alert("Failed to submit this template for review.")
                                        failure = False
                                    End If
                                Else
                                    sqlCommonAction.Transaction.Rollback()
                                    alert("Failed to submit this template for review.")
                                    failure = False
                                End If
                            Else
                                sqlCommonAction.Transaction.Rollback()
                                alert("Failed to submit this template for review.")
                                failure = False
                            End If
                        Else
                            sqlCommonAction.Transaction.Rollback()
                            alert("Failed to submit this template for review.")
                            failure = False
                        End If
                    Else
                        sqlCommonAction.Transaction.Rollback()
                        alert("Failed to submit this template for review.")
                        failure = False
                    End If
                Else
                    alert("You must find an ADC before submitting this template for review.")
                    failure = False
                    sqlCommonAction.Transaction.Rollback()
                End If
                Return failure
            Else
                If (Session("transExists") = True) Then
                    sqlCommonAction.Transaction.Rollback()
                End If
                alert("Failed to submit this template for review.")
                Return False
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            If (Session("transExists") = True) Then
                sqlCommonAction.Transaction.Rollback()
            End If
            alert("Failed to submit this template for review.")
            Return False
        End Try
    End Function

    'attempts to insert a user, returns false if it cannot
    Private Function insertUser(Optional ByVal isHSR As Boolean = False, Optional ByVal isPriv As Boolean = False) As Boolean
        Try
            'dump information that may exist
            If (Me.dsCommon.Tables.Contains("UserList")) Then
                Me.dsCommon.Tables("UserList").Rows.Clear()
            End If
            If (Me.dsCommon.Tables.Contains("PersonData")) Then
                Me.dsCommon.Tables("PersonData").Rows.Clear()
            End If
            If (Me.dsCommon.Tables.Contains("UserIDFetch")) Then
                Me.dsCommon.Tables("UserIDFetch").Rows.Clear()
            End If

            If (isHSR) Then
                'get the email address
                sqlCommonAction.CommandText = "select distinct(php.hanford_id, php.internet_email_address, php.first_name, " & _
                                   " php.last_name, php.middle_initial, phone_no) " & _
                                   " from opwhse.dbo.vw_pub_rbac_role_all_members prra, " & _
                                   " opwhse.dbo.vw_pub_hanford_people php where prra.child_role_name = " & _
                                   " 'HumSub.Humn Subj Cont' and prra.hanford_id = php.hanford_id"
                'sqlCommonAction.CommandText = "Select * from fempUsers where strEmail = 'jesse.sharp@pnl.gov'"
                sqlCommonAdapter.SelectCommand = sqlCommonAction
                sqlCommonAdapter.Fill(Me.dsCommon, "hsrInfo")
                HSRHID = Me.dsCommon.Tables("hsrInfo").Rows(0).Item("hanford_id")
                HSREmailAddress = Me.dsCommon.Tables("hsrInfo").Rows(0).Item("internet_email_address")
                HSRFirstName = Me.dsCommon.Tables("hsrInfo").Rows(0).Item("first_name")
                HSRLastName = Me.dsCommon.Tables("hsrInfo").Rows(0).Item("last_name")
                HSRMiddleName = Me.dsCommon.Tables("hsrInfo").Rows(0).Item("middle_initial")
                HSRPhone = Me.dsCommon.Tables("hsrInfo").Rows(0).Item("phone_no")

                'find out if the user already exists in the system
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER where EMAIL_ADDRESS = '" & HSREmailAddress & "'"
                'sqlCommonAction.CommandText = "Select * from fempUsers where strEmail = 'jesse.sharp@pnl.gov'"
                sqlCommonAdapter.SelectCommand = sqlCommonAction
                sqlCommonAdapter.Fill(Me.dsCommon, "UserList")
            ElseIf (isPriv) Then
                'find out if the user already exists in the system
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER where EMAIL_ADDRESS = 'alan.rither@pnl.gov'"
                'sqlCommonAction.CommandText = "Select * from fempUsers where strEmail = 'jesse.sharp@pnl.gov'"
                sqlCommonAdapter.SelectCommand = sqlCommonAction
                sqlCommonAdapter.Fill(Me.dsCommon, "UserList")
            Else
                'find out if the user already exists in the system
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER where EMAIL_ADDRESS = '" & Me.lblEmailDisplay.Text & "'"
                sqlCommonAdapter.SelectCommand = sqlCommonAction
                sqlCommonAdapter.Fill(Me.dsCommon, "UserList")
            End If

            'if a user exists then return true instead of adding a new person to the user's table
            If (Me.dsCommon.Tables("UserList").Rows.Count = 1) Then
                'determine if the user's level was appropriately selected.  if not bring it up to level 1
                If (Me.dsCommon.Tables("UserList").Rows(0).Item("USER_CODE") < 1) Then
                    If (isHSR) Then
                        sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_USER set USER_CODE = 1 WHERE " & _
                        "EMAIL_ADDRESS = '" & HSREmailAddress & "'"
                        '"strEmail = 'jesse.sharp@pnl.gov'"
                    ElseIf (isPriv) Then
                        sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_USER set USER_CODE = 1 WHERE " & _
                        "EMAIL_ADDRESS = 'alan.rither@pnl.gov'"
                        '"strEmail = 'jesse.sharp@pnl.gov'"
                    Else
                        sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_USER set USER_CODE = 1 WHERE " & _
                        "EMAIL_ADDRESS = '" & Me.lblEmailDisplay.Text & "'"
                    End If
                    sqlCommonAction.ExecuteNonQuery()
                End If
                If (isHSR) Then
                    Session("seqHSRUserID") = Me.dsCommon.Tables("UserLIst").Rows(0).Item("USER_KEY")
                ElseIf (isPriv) Then
                    Session("seqPrivUserID") = Me.dsCommon.Tables("UserLIst").Rows(0).Item("USER_KEY")
                Else
                    Session("seqADCUserID") = Me.dsCommon.Tables("UserLIst").Rows(0).Item("USER_KEY")
                End If
                Return True
            ElseIf (Me.dsCommon.Tables("UserList").Rows.Count > 1) Then
                alert("Not enough information to differentiate between two or more existing users.  The user you are trying to add may already exist in our users list.  Please select the user you are looking for from that list and update their information to indicate their new status for this template.")
                Return False
            End If

            'attempt to add the new record to the database
            'parameterized the text input to allow for things like quotes to get through
            Dim param0 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@EMAIL_ADDRESS", System.Data.SqlDbType.VarChar, 128, "EMAIL_ADDRESS")
            If (isHSR) Then
                param0.Value = HSREmailAddress
                'param0.Value = "jesse.sharp@pnl.gov"
            ElseIf (isPriv) Then
                param0.Value = "alan.rither@pnl.gov"
                'param0.Value = "jesse.sharp@pnl.gov"
            Else
                param0.Value = Me.lblEmailDisplay.Text
            End If
            sqlCommonAction.Parameters.Add(param0)
            Dim param1 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@FIRST_NAME", System.Data.SqlDbType.VarChar, 50, "FIRST_NAME")
            If (isHSR) Then
                param1.Value = HSRFirstName
            ElseIf (isPriv) Then
                param1.Value = "Alan"
            Else
                param1.Value = Me.lblFirstNameDisplay.Text
            End If
            sqlCommonAction.Parameters.Add(param1)
            Dim param2 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@LAST_NAME", System.Data.SqlDbType.VarChar, 50, "LAST_NAME")
            If (isHSR) Then
                param2.Value = HSRLastName
            ElseIf (isPriv) Then
                param2.Value = "Rither"
            Else
                param2.Value = Me.lblLastNameDisplay.Text
            End If
            sqlCommonAction.Parameters.Add(param2)
            Dim param3 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@MIDDLE_NAME", System.Data.SqlDbType.VarChar, 50, "MIDDLE_NAME")
            If (isHSR) Then
                param3.Value = HSRMiddleName
            ElseIf (isPriv) Then
                param3.Value = "C"
            Else
                param3.Value = Me.lblMiddleNameDisplay.Text
            End If
            sqlCommonAction.Parameters.Add(param3)
            Dim param5 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@PRIMARY_WORK_PHONE", System.Data.SqlDbType.VarChar, 32, "PRIMARY_WORK_PHONE")
            If (isHSR) Then
                param5.Value = HSRPhone
            ElseIf (isPriv) Then
                param5.Value = "375-2218"
            Else
                param5.Value = Me.lblPhoneNumberDisplay.Text
            End If
            sqlCommonAction.Parameters.Add(param5)
            Dim param7 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@HANFORD_ID", System.Data.SqlDbType.VarChar, 12, "HANFORD_ID")

            'get the user's HID if it exists
            If (Session("PersonTable") Is Nothing Or isHSR Or isPriv) Then
                If (isHSR) Then
                    'get the user's informaton - the user adding must have manually input their information
                    'set up and process the sql statement
                    sqlCommonAction.CommandText = "select isnull(max(hanford_id), '') as hanford_id, isnull(max(hanf_pay_no), '') as " & _
                    "hanf_pay_no, max(first_name) as first_name, max(last_name) as last_name, isnull(max(middle_initial), '') as " & _
                    "middle_initial, isnull(max(name_prefix), '') as name_prefix, isnull(max(name_suffix), '') as name_suffix, " & _
                    "isnull(max(phone_no), '') as phone_no, isnull(max(internet_email_address), '') as internet_email_address, " & _
                    "isnull(max(employed_by_hcid), '') as employed_by_hcid, isNull(max(active_sw), 'N') as active_sw from  (" & _
                    "select hanford_id, hanf_pay_no, first_name, last_name, middle_initial, NULL as name_prefix, name_suffix, phone_no, " & _
                    "internet_email_address, employed_by_hcid, active_sw from opwhse.dbo.vw_pub_hanford_people union select hanford_id, " & _
                    "pay_no as hanf_pay_no, first_name, last_name, middle_name as middle_initial, NULL as name_prefix, name_suffix, " & _
                    "work_phone as phone_no, internet_address as internet_email_address, employed_by_hcid, NULL as active_sw " & _
                    "from opwhse.dbo.vw_pub_pnnl_associate union select hanford_id, pnnl_pay_no as hanf_pay_no, first_name, last_name, " & _
                    "middle_initial, name_prefix, name_suffix, contact_work_phone as phone_no, internet_email_address, company as " & _
                    "employed_by_hcid, NULL as active_sw from opwhse.dbo.vw_pub_bmi_employee) a where active_sw = 'Y' and " & _
                    "first_name = '" & HSRFirstName & "' and last_name = '" & HSRLastName & "' and " & _
                    "internet_email_address = '" & HSREmailAddress & "' group by hanford_id, " & _
                    "last_name, first_name, middle_initial, internet_email_address"
                ElseIf (isPriv) Then
                    'get the user's informaton - the user adding must have manually input their information
                    'set up and process the sql statement
                    sqlCommonAction.CommandText = "select isnull(max(hanford_id), '') as hanford_id, isnull(max(hanf_pay_no), '') as " & _
                    "hanf_pay_no, max(first_name) as first_name, max(last_name) as last_name, isnull(max(middle_initial), '') as " & _
                    "middle_initial, isnull(max(name_prefix), '') as name_prefix, isnull(max(name_suffix), '') as name_suffix, " & _
                    "isnull(max(phone_no), '') as phone_no, isnull(max(internet_email_address), '') as internet_email_address, " & _
                    "isnull(max(employed_by_hcid), '') as employed_by_hcid, isNull(max(active_sw), 'N') as active_sw from  (" & _
                    "select hanford_id, hanf_pay_no, first_name, last_name, middle_initial, NULL as name_prefix, name_suffix, phone_no, " & _
                    "internet_email_address, employed_by_hcid, active_sw from opwhse.dbo.vw_pub_hanford_people union select hanford_id, " & _
                    "pay_no as hanf_pay_no, first_name, last_name, middle_name as middle_initial, NULL as name_prefix, name_suffix, " & _
                    "work_phone as phone_no, internet_address as internet_email_address, employed_by_hcid, NULL as active_sw " & _
                    "from opwhse.dbo.vw_pub_pnnl_associate union select hanford_id, pnnl_pay_no as hanf_pay_no, first_name, last_name, " & _
                    "middle_initial, name_prefix, name_suffix, contact_work_phone as phone_no, internet_email_address, company as " & _
                    "employed_by_hcid, NULL as active_sw from opwhse.dbo.vw_pub_bmi_employee) a where active_sw = 'Y' and " & _
                    "first_name = 'Alan' and last_name = 'Rither' and middle_initial = 'C' and " & _
                    "internet_email_address = 'alan.rither@pnl.gov' group by hanford_id, " & _
                    "last_name, first_name, middle_initial, internet_email_address"
                Else
                    'get the user's informaton - the user adding must have manually input their information
                    'set up and process the sql statement
                    sqlCommonAction.CommandText = "select isnull(max(hanford_id), '') as hanford_id, isnull(max(hanf_pay_no), '') as " & _
                    "hanf_pay_no, max(first_name) as first_name, max(last_name) as last_name, isnull(max(middle_initial), '') as " & _
                    "middle_initial, isnull(max(name_prefix), '') as name_prefix, isnull(max(name_suffix), '') as name_suffix, " & _
                    "isnull(max(phone_no), '') as phone_no, isnull(max(internet_email_address), '') as internet_email_address, " & _
                    "isnull(max(employed_by_hcid), '') as employed_by_hcid, isNull(max(active_sw), 'N') as active_sw from  (" & _
                    "select hanford_id, hanf_pay_no, first_name, last_name, middle_initial, NULL as name_prefix, name_suffix, phone_no, " & _
                    "internet_email_address, employed_by_hcid, active_sw from opwhse.dbo.vw_pub_hanford_people union select hanford_id, " & _
                    "pay_no as hanf_pay_no, first_name, last_name, middle_name as middle_initial, NULL as name_prefix, name_suffix, " & _
                    "work_phone as phone_no, internet_address as internet_email_address, employed_by_hcid, NULL as active_sw " & _
                    "from opwhse.dbo.vw_pub_pnnl_associate union select hanford_id, pnnl_pay_no as hanf_pay_no, first_name, last_name, " & _
                    "middle_initial, name_prefix, name_suffix, contact_work_phone as phone_no, internet_email_address, company as " & _
                    "employed_by_hcid, NULL as active_sw from opwhse.dbo.vw_pub_bmi_employee) a where active_sw = 'Y' and " & _
                    "first_name = '" & Me.lblFirstNameDisplay.Text & "' and last_name = '" & Me.lblLastNameDisplay.Text & "' and middle_initial = '" & _
                    Me.lblMiddleNameDisplay.Text & "' and internet_email_address = '" & Me.lblEmailDisplay.Text & "' group by hanford_id, " & _
                    "last_name, first_name, middle_initial, internet_email_address"
                End If
                sqlCommonAdapter.Fill(Me.dsCommon, "PersonData")

                'if we found something assign the HID otherwise return false
                If (Me.dsCommon.Tables("PersonData").Rows.Count < 1) Then
                    alert("This person does not match any active personnel.  If you think this may be an error then try the search capabilities of this form to locate the person faster.")
                    Return False
                Else
                    param7.Value = Me.dsCommon.Tables("PersonData").Rows(0).Item("hanford_id")
                End If
            Else
                'get the user's hid from the stored table
                Dim dt As DataTable = CType(Session("PersonTable"), DataTable)
                param7.Value = dt.Rows(Session("selectedPersonRow")).Item("hanford_id")
            End If
            sqlCommonAction.Parameters.Add(param7)
            Dim param8 As System.Data.SqlClient.SqlParameter = New System.Data.SqlClient.SqlParameter("@AUTHENTICATION_CODE", System.Data.SqlDbType.VarChar, 64, "AUTHENTICATION_CODE")

            'generate a password for this user
            Dim proxyAuthentication As New Authentication.Service1
            If (Trim("" & param7.Value) = "") Then
                If (isHSR) Then
                    param8.Value = myUtility.hashPassword(proxyAuthentication.createPasswordString(), HSREmailAddress)
                    'param8.Value = myUtility.hashPassword(myUtility.createPasswordString(), "jesse.sharp@pnl.gov")
                ElseIf (isPriv) Then
                    param8.Value = myUtility.hashPassword(proxyAuthentication.createPasswordString(), "alan.rither@pnl.gov")
                Else
                    param8.Value = myUtility.hashPassword(proxyAuthentication.createPasswordString(), Me.lblEmailDisplay.Text)
                End If
                If (param8.Value = "") Then
                    Return False
                End If
                If (isHSR) Then
                    Session("newUserName") = HSREmailAddress
                    'Session("newUserName") = "jesse.sharp@pnl.gov"
                ElseIf (isPriv) Then
                    Session("newUserName") = "alan.rither@pnl.gov"
                Else
                    Session("newUserName") = Me.lblEmailDisplay.Text
                End If
            Else
                param8.Value = myUtility.hashPassword(proxyAuthentication.createPasswordString(), param7.Value)
                If (param8.Value = "") Then
                    Return False
                End If
                Session("newUserName") = param7.Value
            End If
            sqlCommonAction.Parameters.Add(param8)

            'insert the new user
            sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_USER (HANFORD_ID, EMAIL_ADDRESS, USER_CODE, LAST_NAME, FIRST_NAME, " & _
            "MIDDLE_NAME, SUFFIX_NAME, AUTHENTICATION_CODE, PRIMARY_WORK_PHONE, CREATE_DATE, USED_DATE, CREATOR_USER_KEY, ACTIVE_IND, " & _
            "TRAINING_IND, COMPANY_ABBREV, NO_EMAIL_IND) VALUES (@HANFORD_ID, @EMAIL_ADDRESS, 1, @LAST_NAME," & _
            "@FIRST_NAME, @MIDDLE_NAME, '', @AUTHENTICATION_CODE, @PRIMARY_WORK_PHONE,'" & Now & "','1/1/1970'," & _
            Session("UserID") & ", 1, 0, '', 0)"
            sqlCommonAction.ExecuteNonQuery()

            'get the user's sequence number
            If (isHSR) Then
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER where EMAIL_ADDRESS = '" & HSREmailAddress & "'"
                'sqlCommonAction.CommandText = "Select * from fempUsers where strEmail = 'jesse.sharp@pnl.gov'"
                sqlCommonAdapter.Fill(Me.dsCommon, "UserIDFetch")
                Session("seqHSRUserID") = Me.dsCommon.Tables("UserIDFetch").Rows(0).Item("USER_KEY")
            ElseIf (isPriv) Then
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER where EMAIL_ADDRESS = 'alan.rither@pnl.gov'"
                'sqlCommonAction.CommandText = "Select * from fempUsers where strEmail = 'jesse.sharp@pnl.gov'"
                sqlCommonAdapter.Fill(Me.dsCommon, "UserIDFetch")
                Session("seqPrivUserID") = Me.dsCommon.Tables("UserIDFetch").Rows(0).Item("USER_KEY")
            Else
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER where EMAIL_ADDRESS = '" & _
                Me.lblEmailDisplay.Text & "'"
                sqlCommonAdapter.Fill(Me.dsCommon, "UserIDFetch")
                Session("seqADCUserID") = Me.dsCommon.Tables("UserIDFetch").Rows(0).Item("USER_KEY")
            End If

            'send the user an e-mail notifying them of their status
            If (isHSR) Then
                If Not (sendMail(HSREmailAddress, param7.Value, False)) Then
                    '   If Not (sendMail("jesse.sharp@pnl.gov", param7.Value, False)) Then
                    Return False
                End If
            ElseIf (isPriv) Then
                If Not (sendMail("alan.rither@pnl.gov", param7.Value, False)) Then
                    Return False
                End If
            Else
                If Not (sendMail(Me.lblEmailDisplay.Text, param7.Value, False)) Then
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'attempts to insert permissions for the creator into the permissions table, returns false if it cannot
    Private Function insertPermission(Optional ByVal isHSR As Boolean = False, Optional ByVal isPriv As Boolean = False) As Boolean
        Try
            'dump any old information that may exist
            If (Me.dsCommon.Tables.Contains("PermissionList")) Then
                Me.dsCommon.Tables("PermissionList").Rows.Clear()
            End If

            'parameterized the text input to allow for things like quotes to get through
            sqlCommonAction.Parameters.Clear()

            'determine if the user already exists
            If (isHSR) Then
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER_PERMISSION where USER_KEY = " & Session("seqHSRUserID") & _
                " and TEMPLATE_KEY = " & Session("seqTemplate")
            ElseIf (isPriv) Then
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER_PERMISSION where USER_KEY = " & Session("seqPrivUserID") & _
                " and TEMPLATE_KEY = " & Session("seqTemplate")
            Else
                sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER_PERMISSION where USER_KEY = " & Session("seqADCUserID") & _
                " and TEMPLATE_KEY = " & Session("seqTemplate")
            End If
            sqlCommonAdapter.Fill(Me.dsCommon, "PermissionList")

            'if the user already exists under the template then update their record for that template
            If (Me.dsCommon.Tables("PermissionList").Rows.Count() > 0) Then
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_USER_PERMISSION set USER_IND = 1"
                If (isHSR) Then
                    sqlCommonAction.CommandText &= " where USER_KEY = " & Session("seqHSRUserID") & " and TEMPLATE_KEY = " & Session("seqTemplate")
                ElseIf (isPriv) Then
                    sqlCommonAction.CommandText &= " where USER_KEY = " & Session("seqPrivUserID") & " and TEMPLATE_KEY = " & Session("seqTemplate")
                Else
                    sqlCommonAction.CommandText &= " where USER_KEY = " & Session("seqADCUserID") & " and TEMPLATE_KEY = " & Session("seqTemplate")
                End If
            Else
                'Add permissions for the survey
                If (isHSR) Then
                    sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_USER_PERMISSION (USER_KEY, TEMPLATE_KEY, SURVEY_KEY, CREATOR_USER_KEY, " & _
                    "CREATE_DATE, OWNER_IND, DELEGATE_IND, USER_IND) VALUES (" & Session("seqHSRUserID") & ", " & Session("seqTemplate") & _
                    ", -1, " & Session("UserID") & ", '" & Now & "', 0, 0, 1)"
                ElseIf (isPriv) Then
                    sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_USER_PERMISSION (USER_KEY, TEMPLATE_KEY, SURVEY_KEY, CREATOR_USER_KEY, " & _
                    "CREATE_DATE, OWNER_IND, DELEGATE_IND, USER_IND) VALUES (" & Session("seqPrivUserID") & ", " & Session("seqTemplate") & _
                    ", -1, " & Session("UserID") & ", '" & Now & "', 0, 0, 1)"
                Else
                    sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_USER_PERMISSION (USER_KEY, TEMPLATE_KEY, SURVEY_KEY, CREATOR_USER_KEY, " & _
                    "CREATE_DATE, OWNER_IND, DELEGATE_IND, USER_IND) VALUES (" & Session("seqADCUserID") & ", " & Session("seqTemplate") & _
                    ", -1, " & Session("UserID") & ", '" & Now & "', 0, 0, 1)"
                End If
            End If

            'process the command
            sqlCommonAction.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function

    'insert or update the ADC
    Function enterADC() As Boolean
        sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE SET HANFORD_ID = '" & _
        Me.lblHIDDisplay.Text & "', EMAIL_ADDRESS = '" & Me.lblEmailDisplay.Text & _
        "', NAME = '" & generateName() & _
        "', AUTHOR_COMMENT = @AUTHOR_COMMENT, SIGNED_IND = 0, " & _
        "TRANSACTION_ID = 0" & _
        " WHERE TEMPLATE_KEY = " & Session("seqTemplate") & _
        " and AUTH_DERIV_CLASSIFIER_IND = 1 and CURRENT_IND = 1"
        If (sqlCommonAction.ExecuteNonQuery() <= 0) Then
            Dim strADCName As String = ""
            'generate the ADC Name
            strADCName = generateName()

            'do the insert
            sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE " & _
            "(HANFORD_ID, TEMPLATE_KEY, EMAIL_ADDRESS, NAME, AUTH_DERIV_CLASSIFIER_IND, HUMAN_SUBJECTS_RESCH_IND, PRIVACY_IND, SIGNED_IND, CURRENT_IND, AUTHOR_COMMENT) " & _
            "VALUES ('" & Me.lblHIDDisplay.Text & "', " & _
            Session("seqTemplate") & ", '" & Me.lblEmailDisplay.Text & "', '" & _
            strADCName & "', 1, 0, 0, 0, 1, @AUTHOR_COMMENT)"
            If (sqlCommonAction.ExecuteNonQuery() <= 0) Then
                alert("Error adding ADC representative.")
                Return False
            Else
                Return True
            End If
        Else
            Return True
        End If
    End Function

    'insert or update the HSR
    Function enterHSR() As Boolean
        sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE SET AUTHOR_COMMENT = " & _
        " @AUTHOR_COMMENT, SIGNED_IND = 0, TRANSACTION_ID = 0 WHERE " & _
        " TEMPLATE_KEY = " & Session("seqTemplate") & " and HUMAN_SUBJECTS_RESCH_IND = 1 and " & _
        " CURRENT_IND = 1"
        If (sqlCommonAction.ExecuteNonQuery() <= 0) Then
            sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE " & _
            "(HANFORD_ID, TEMPLATE_KEY, EMAIL_ADDRESS, NAME, AUTH_DERIV_CLASSIFIER_IND, HUMAN_SUBJECTS_RESCH_IND, PRIVACY_IND, SIGNED_IND, CURRENT_IND, AUTHOR_COMMENT) " & _
            "VALUES ('" & HSRHID & "', " & Session("seqTemplate") & _
            ", '" & HSREmailAddress & "', '" & HSRLastName & ", " & HSRFirstName & " " & HSRMiddleName & "', 0, 1, 0, 0, 1, @AUTHOR_COMMENT)"
            If (sqlCommonAction.ExecuteNonQuery() <= 0) Then
                alert("Error adding HSR representative.")
                Return False
            Else
                Return True
            End If
        Else
            Return True
        End If
    End Function

    'insert or update the Privacy Person
    Function enterPriv() As Boolean
        sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE SET AUTHOR_COMMENT = " & _
        " @AUTHOR_COMMENT, SIGNED_IND = 0, TRANSACTION_ID = 0 WHERE " & _
        " TEMPLATE_KEY = " & Session("seqTemplate") & " and PRIVACY_IND = 1 and " & _
        " CURRENT_IND = 1"
        If (sqlCommonAction.ExecuteNonQuery() <= 0) Then
            sqlCommonAction.CommandText = "Insert INTO " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE " & _
            "(HANFORD_ID, TEMPLATE_KEY, EMAIL_ADDRESS, NAME, AUTH_DERIV_CLASSIFIER_IND, HUMAN_SUBJECTS_RESCH_IND, PRIVACY_IND, SIGNED_IND, CURRENT_IND, AUTHOR_COMMENT) " & _
            "VALUES ('0010720', " & Session("seqTemplate") & _
            ", 'alan.rither@pnl.gov', 'Rither, Alan C', 0, 0, 1, 0, 1, @AUTHOR_COMMENT)"
            If (sqlCommonAction.ExecuteNonQuery() <= 0) Then
                alert("Error adding Privacy representative.")
                Return False
            Else
                Return True
            End If
        Else
            Return True
        End If
    End Function

    'removes the old current signature flag to make room for the new records
    Function removeOldCurrentSignatures() As Boolean
        Try
            'remove the current flags for each of the records
            Dim row As DataRow
            For Each row In Me.dsCommon.Tables("Current").Rows
                sqlCommonAction.CommandText = "Update " & myUtility.getExtension() & "SAT_TEMPLATE_SIGNATURE set " & _
                "CURRENT_IND = 0 where TEMPLATE_KEY = " & Session("seqTemplate") & _
                " and TRANSACTION_ID > 0"
                sqlCommonAction.ExecuteNonQuery()
                Return True
            Next
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            alert("Error updating the signature records.")
            Return False
        End Try
    End Function
#End Region

#Region "Generate Name"
    'puts the name together for display
    Private Function generateName() As String
        'generate the ADC Name
        Dim strADCName As String = ""
        If (Me.lblMiddleNameDisplay.Text = "") Then
            strADCName = Me.lblLastNameDisplay.Text & ", " & Me.lblFirstNameDisplay.Text
        Else
            strADCName = Me.lblLastNameDisplay.Text & ", " & _
                Me.lblFirstNameDisplay.Text & " " & Me.lblMiddleNameDisplay.Text
        End If
        Return strADCName
    End Function
#End Region

#Region "Send Mail"
    'Sends an e-mail to users with changed passwords, except level 0 users.
    Private Function sendMail(ByVal Email As String, ByVal strUserName As String, Optional ByVal inSystem As Boolean = False) As Boolean
        Trace.Warn("Sending user an email")
        'Set up the mail messsage
        Try
            'dump any information that may exist
            If (Me.dsCommon.Tables.Contains("ModifyID")) Then
                Me.dsCommon.Tables("ModifyID").Rows.Clear()
            End If

            Dim strFrom As String = "survey.administrator@pnl.gov"
            Dim strTo As String = Email
            'Dim strTo As String = "jesse.sharp@pnl.gov"
            Dim strbldrSubject As New StringBuilder("Survey Admin Tool Template Delegation.")

            'get the current user's name and e-mail
            sqlCommonAdapter.SelectCommand = sqlCommonAction
            sqlCommonAction.Connection = Me.commonConn
            sqlCommonAction.CommandText = "Select * from " & myUtility.getExtension() & "SAT_USER where USER_KEY = " & Session("UserID")
            sqlCommonAdapter.Fill(Me.dsCommon, "ModifyID")

            If (Me.dsCommon.Tables("ModifyID").Rows.Count < 1) Then
                alert("Unable to locate your information.  You may have been removed from the database in the last few minutes.  Contact the Survey Administrator immediately.")
                Return False
            End If

            Dim strbldrMessage As New StringBuilder
            'notify the user of their status change
            strbldrMessage.Append("You have been given reviewer access to the survey administration tool by")
            strbldrMessage.Append(" " & Me.dsCommon.Tables("ModifyID").Rows(0).Item("FIRST_NAME") & " ")
            strbldrMessage.Append(Me.dsCommon.Tables("ModifyID").Rows(0).Item("LAST_NAME") & ".")
            strbldrMessage.Append("  You may contact this person at this e-mail address: <b>")
            strbldrMessage.Append(Me.dsCommon.Tables("ModifyID").Rows(0).Item("EMAIL_ADDRESS") & "</b><br/>")
            If Not (inSystem) Then
                strbldrMessage.Append("Please Make a note of your password.<br/> If you wish to change your password ")
                strbldrMessage.Append("you may do so at the ")
                strbldrMessage.Append("<a href='http://" & Request.ServerVariables.Item("SERVER_NAME") & "/SurveyAdmin/Login.aspx'>Survey Administration</a> website. <p></p>")
                strbldrMessage.Append("Your User Name is: " & strUserName & "<br/>")
                strbldrMessage.Append("Your Password is: " & strNewPassword)
            End If

            'send the e-mail
            If (myUtility.genericSendMail(strFrom, strTo, strbldrSubject.ToString, strbldrMessage.ToString)) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Return False
        End Try
    End Function
#End Region
End Class