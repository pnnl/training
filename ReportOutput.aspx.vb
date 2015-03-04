Public Class ReportOutput
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    Private Sub InitializeComponent()
        Me.commonConn = New System.Data.SqlClient.SqlConnection

    End Sub
    Protected WithEvents lblSecPriv As System.Web.UI.WebControls.Label
    Protected WithEvents lblWebmaster As System.Web.UI.WebControls.Label
    Protected WithEvents lblSurveyHeader As System.Web.UI.WebControls.Label
    Protected WithEvents lblReportHeader As System.Web.UI.WebControls.Label
    Protected WithEvents lblTotalRespondents As System.Web.UI.WebControls.Label
    Protected sqlCommonCommand As SqlClient.SqlCommand
    Protected sqlCommonAdapter As SqlClient.SqlDataAdapter
    Protected dsCommon As DataSet
    Protected WithEvents lblDateRange As System.Web.UI.WebControls.Label
    Protected WithEvents lblPeriodRecord As System.Web.UI.WebControls.Label
    Protected WithEvents lblNote As System.Web.UI.WebControls.Label
    Protected strDataFilter As String
    Protected strDateFilter As String
    Protected strQuestionFilter As String
    Protected WithEvents ReportFilter As System.Web.UI.WebControls.Table
    Protected WithEvents ReportFilterResults As System.Web.UI.WebControls.Table
    Protected WithEvents DataTable As System.Web.UI.WebControls.Table
    Protected WithEvents DataTable3 As System.Web.UI.WebControls.Table
    Protected WithEvents DataTable2 As System.Web.UI.WebControls.Table
    Protected WithEvents lblNoData As System.Web.UI.WebControls.Label
    Protected dateFilterArr As ArrayList = New ArrayList
    Protected WithEvents commonConn As System.Data.SqlClient.SqlConnection
    Protected myUtility As New Utility
    Protected WithEvents ReportFilterFooter As System.Web.UI.WebControls.Table

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
        Trace.Warn("Page Load")
        InitializeComponent()

        'clear session
        myUtility.clearSession()

        'Set the connection based on the machine
        myUtility.getConnection(Me.commonConn)

        'catch a session time-out
        If (Session.IsNewSession) Then
            Session("Alert") = "Your session has timed out for security purposes.  Please log in again."
            Session("CurrentPage") = "./Login.aspx"
            Response.Write("<script type='text/javascript'>window.opener.location.href = window.opener.location.href;</script>")
            Response.Write("<script type='text/javascript'>window.close()</script>")
        End If

        'process the report
        If Not (IsPostBack) Then
            If (setPreliminaryDateFilter()) Then
                If (setPreliminaryDataFilter(CType(Me.dateFilterArr(0), ArrayList)(2))) Then
                    If (processReportHeader()) Then
                        'process report results for the header
                        Me.lblReportHeader.Text = Session("ReportHeader")
                        Me.lblSurveyHeader.Text = Session("SurveyHeader")
                        Me.lblTotalRespondents.Text = getTotalRespondents()
                        Me.lblDateRange.Text = getDateRange()
                        Me.lblPeriodRecord.Text = getPeriodRecord()
                        If (setDataFilter(CType(Me.dateFilterArr(0), ArrayList)(2))) Then
                            If (setDateFilter()) Then
                                'get the colors for the charts and tables
                                If (getReportColors()) Then
                                    'generates the question filter
                                    Dim QFilter As ArrayList = New ArrayList
                                    If (setQuestionFilter(QFilter)) Then
                                        'generates the report filter
                                        If (processReportFilters()) Then
                                            'remove the first item only if there are other date filters
                                            If (Me.dateFilterArr.Count > 1) Then
                                                Me.dateFilterArr.RemoveAt(0)
                                            End If
                                            'builds the ratings rows for the first data table
                                            If (processRatings()) Then
                                                'builds the choice rows for the second data table
                                                If (processChoiceData()) Then
                                                    'builds the rest of the rows for the second table
                                                    If (processOtherData()) Then
                                                        'if there are no non 'R' type questions with data then don't show the data section
                                                        If (Session("RItemExists") = False And Session("NonRTypeSelected") = False) Then
                                                            Session("QuestionsSelected") = False
                                                        End If

                                                        'determine if we need to display the text section
                                                        If (CType(Session("TextOutput"), Boolean)) Then
                                                            If Not (processTextData()) Then
                                                                alert(Session("Alert"))
                                                                'Response.Write("<script type='text/javascript'>window.close()</script>")
                                                            End If
                                                        End If
                                                    Else
                                                        alert(Session("Alert"))
                                                        'Response.Write("<script type='text/javascript'>window.close()</script>")
                                                    End If
                                                Else
                                                    alert(Session("Alert"))
                                                    'Response.Write("<script type='text/javascript'>window.close()</script>")
                                                End If
                                            Else
                                                alert(Session("Alert"))
                                                'Response.Write("<script type='text/javascript'>window.close()</script>")
                                            End If
                                        Else
                                            alert(Session("Alert"))
                                            'Response.Write("<script type='text/javascript'>window.close()</script>")
                                        End If
                                    Else
                                        alert(Session("Alert"))
                                        'Response.Write("<script type='text/javascript'>window.close()</script>")
                                    End If
                                Else
                                    alert(Session("Alert"))
                                    'Response.Write("<script type='text/javascript'>window.close()</script>")
                                End If
                            Else
                                alert(Session("Alert"))
                                'Response.Write("<script type='text/javascript'>window.close()</script>")
                            End If
                        Else
                            alert(Session("Alert"))
                            'Response.Write("<script type='text/javascript'>window.close()</script>")
                        End If
                    Else
                        alert(Session("Alert"))
                        'Response.Write("<script type='text/javascript'>window.close()</script>")
                    End If
                Else
                    alert(Session("Alert"))
                    'Response.Write("<script type='text/javascript'>window.close()</script>")
                End If
            Else
                alert(Session("Alert"))
                'Response.Write("<script type='text/javascript'>window.close()</script>")
            End If
        End If
    End Sub

    'handles error messages
    Public Sub alert(ByVal message As String)
        If (message <> "") Then
            Response.Write("<script type='text/javascript'>window.alert('" & message & "');</script>")
            Session("Alert") = ""
        End If
    End Sub
#End Region
    
#Region "Data Load"
    'returns a formatted string for the respondents section of the report header
    Private Function getTotalRespondents() As String
        Trace.Warn("Getting Total Respondents")
        Dim returnString As String = ""
        Dim responded As Integer = Me.dsCommon.Tables("Total").Rows(0).Item("Responded")
        Dim total As Integer = Me.dsCommon.Tables("Total").Rows(0).Item("Total")
        returnString = responded & " of " & total & " (" & Math.Round(((responded / total) * 100), 2) & "%)"
        Return returnString
    End Function

    'returns a formatted string for the period of record
    Private Function getPeriodRecord() As String
        Trace.Warn("Getting Period of Record")
        Dim returnString As String = ""
        Dim minDate As String = CType(Me.dsCommon.Tables("Total").Rows(0).Item("Min1"), Date).Date & ""
        Dim maxDate As String = CType(Me.dsCommon.Tables("Total").Rows(0).Item("Max1"), Date).Date & ""
        If (minDate <> "") Then
            returnString = minDate & " - " & maxDate
        Else
            returnString = "No Data - No Data"
        End If
        Return returnString
    End Function

    'returns a formatted string for the date range filter
    Private Function getDateRange() As String
        Trace.Warn("Getting Date Range")
        Dim returnString As String = ""
        If (Session("StartDate") <> "") Then
            returnString &= CType(Session("StartDate"), Date).Date & " - "
        Else
            returnString &= "None - "
        End If

        If (Session("EndDate") <> "") Then
            returnString &= CType(Session("EndDate"), Date).Date
        Else
            returnString &= "None"
        End If
        Return returnString
    End Function

    'gets the colors for the report charts and tables
    Private Function getReportColors() As Boolean
        Trace.Warn("getting report colors")
        Try
            'colors ordered by intColorOrder
            Me.sqlCommonCommand.CommandText = "Select CHART_COLOR_CODE from " & myUtility.getExtension() & "SAT_CHART_COLOR " & _
                "Order by CHART_COLOR_CODE"
            Me.sqlCommonAdapter.Fill(Me.dsCommon, "Colors")

            Return True
        Catch ex As Exception
            Session("Alert") = "Failed to get report colors."
            Return False
        End Try
    End Function
#End Region

#Region "Filters"
    'sets the preliminary date filter
    Private Function setPreliminaryDateFilter() As Boolean
        Trace.Warn("Setting Date Filter")
        Try
            'no data split so do one big data filter
            Me.strDateFilter = ""
            If (Session("StartDate") <> "") Then
                If (Session("EndDate") <> "") Then
                    Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                        CType(Session("StartDate"), DateTime) & "' and frs.LAST_USED_DATE <= '" & _
                        CType(Session("EndDate"), DateTime).AddHours(23).AddMinutes(59).AddSeconds(59) & "'"
                Else
                    Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                       CType(Session("StartDate"), DateTime) & "'"
                End If
            ElseIf (Session("EndDate") <> "") Then
                Me.strDateFilter &= " and frs.LAST_USED_DATE <= '" & _
                    CType(Session("EndDate"), DateTime).AddHours(23).AddMinutes(59).AddSeconds(59) & "'"
            End If
            Dim tempArray As ArrayList = New ArrayList
            tempArray.Add(Session("StartDate"))
            tempArray.Add(Session("EndDate"))
            tempArray.Add(Me.strDateFilter)
            Me.dateFilterArr.Add(tempArray)
            Return True
        Catch ex As Exception
            Session("Alert") = "Preliminary date filter creation failure.  " & ex.ToString
            Return False
        End Try
    End Function

    'sets the date filter
    Private Function setDateFilter() As Boolean
        Trace.Warn("Setting Date Filter")
        Try
            Me.dateFilterArr.Clear()
            Dim tempArray As ArrayList = New ArrayList
            'no data split so do one big data filter
            Me.strDateFilter = ""
            If (Session("StartDate") <> "") Then
                If (Session("EndDate") <> "") Then
                    Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                        CType(Session("StartDate"), DateTime) & "' and frs.LAST_USED_DATE <= '" & _
                        CType(Session("EndDate"), DateTime).AddHours(23).AddMinutes(59).AddSeconds(59) & "'"
                Else
                    Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                       CType(Session("StartDate"), DateTime) & "'"
                End If
            ElseIf (Session("EndDate") <> "") Then
                Me.strDateFilter &= " and frs.LAST_USED_DATE <= '" & _
                    CType(Session("EndDate"), DateTime).AddHours(23).AddMinutes(59).AddSeconds(59) & "'"
            End If
            tempArray.Add(Session("StartDate"))
            tempArray.Add(Session("EndDate"))
            tempArray.Add(Me.strDateFilter)
            Me.dateFilterArr.Add(tempArray)

            If (Session("SplitOutput") = "YEAR") Then
                'split by year
                Dim minDate As DateTime = CType(Me.dsCommon.Tables("Total").Rows(0).Item("Min1"), Date).Date & ""
                Dim maxDate As DateTime = CType(Me.dsCommon.Tables("Total").Rows(0).Item("Max1"), Date).Date & ""

                'loop until we've generated date filters for each year
                While (minDate <= maxDate)
                    tempArray = New ArrayList

                    'if we're in a fiscal year that matches the current year number
                    If (minDate.Month < 10) Then
                        Me.strDateFilter = ""
                        Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                            minDate.Month & "/" & minDate.Day & "/" & minDate.Year & _
                            "' and frs.LAST_USED_DATE <= '"
                        tempArray.Add(minDate.Month & "/" & minDate.Day & "/" & minDate.Year)

                        If (maxDate.Month >= 10 Or maxDate.Year > minDate.Year) Then
                            'maximum date goes into the next fiscal year
                            Me.strDateFilter &= "10/01/" & minDate.Year & "'"
                            tempArray.Add("10/01/" & minDate.Year)
                            minDate = "10/01/" & minDate.Year
                        Else
                            'maximum date does not go into the next fiscal year
                            Me.strDateFilter &= maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year & "'"
                            tempArray.Add(maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year)
                            minDate = maxDate.AddDays(1)
                        End If
                    ElseIf (minDate.Month >= 10) Then
                        Me.strDateFilter = ""
                        Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                            minDate.Month & "/" & minDate.Day & "/" & minDate.Year & _
                            "' and frs.LAST_USED_DATE <= '"
                        tempArray.Add(minDate.Month & "/" & minDate.Day & "/" & minDate.Year)

                        If (maxDate.Month >= 10 And maxDate.Year >= minDate.Year + 1) Then
                            'maximum date goes into the next fiscal year
                            Me.strDateFilter &= "10/01/" & minDate.Year + 1 & "'"
                            tempArray.Add("10/01/" & minDate.Year + 1)
                            minDate = "10/01/" & minDate.Year + 1
                        ElseIf (maxDate.Year > minDate.Year + 1) Then
                            'maximum date goes into the next fiscal year
                            Me.strDateFilter &= "10/01/" & minDate.Year + 1 & "'"
                            tempArray.Add("10/01/" & minDate.Year + 1)
                            minDate = "10/01/" & minDate.Year + 1
                        Else
                            'maximum date does not go into the next fiscal year
                            Me.strDateFilter &= maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year & "'"
                            tempArray.Add(maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year)
                            minDate = maxDate.AddDays(1)
                        End If
                    End If

                    'add the date filter
                    tempArray.Add(Me.strDateFilter)
                    Me.dateFilterArr.Add(tempArray)
                End While
            ElseIf (Session("SplitOutput") = "SEMIANNUAL") Then
                Dim minDate As Date = CType(Me.dsCommon.Tables("Total").Rows(0).Item("Min1"), Date).Date & ""
                Dim maxDate As Date = CType(Me.dsCommon.Tables("Total").Rows(0).Item("Max1"), Date).Date & ""

                'loop until we've generated date filters for each half of the year
                While (minDate <= maxDate)
                    tempArray = New ArrayList

                    'if we're in a fiscal year that matches the current year number
                    If ((minDate.Month >= 10 And minDate.Month <= 12) Or _
                        (minDate.Month >= 1 And minDate.Month <= 3)) Then
                        Me.strDateFilter = ""
                        Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                            minDate.Month & "/" & minDate.Day & "/" & minDate.Year & _
                            "' and frs.LAST_USED_DATE <= '"
                        tempArray.Add(minDate.Month & "/" & minDate.Day & "/" & minDate.Year)

                        If (maxDate.Month > 3 Or maxDate.Year > minDate.year) Then
                            If (minDate.Month >= 10) Then
                                'maximum date goes into the next semi-annual portion of the year
                                Me.strDateFilter &= "3/31/" & minDate.Year + 1 & "'"
                                tempArray.Add("3/31/" & minDate.Year + 1)
                                minDate = "4/01/" & minDate.Year + 1
                            Else
                                'maximum date goes into the next semi-annual portion of the year
                                Me.strDateFilter &= "3/31/" & minDate.Year & "'"
                                tempArray.Add("3/31/" & minDate.Year)
                                minDate = "4/01/" & minDate.Year
                            End If
                        Else
                            'maximum date does not go into the next semi-annual portion of the year
                            Me.strDateFilter &= maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year & "'"
                            minDate = maxDate.AddDays(1)
                        End If
                    ElseIf (minDate.Month >= 4 And minDate.Month <= 9) Then
                        Me.strDateFilter = ""
                        Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                            minDate.Month & "/" & minDate.Day & "/" & minDate.Year & _
                            "' and frs.LAST_USED_DATE <= '"
                        tempArray.Add(minDate.Month & "/" & minDate.Day & "/" & minDate.Year)

                        If (maxDate.Month > 9 Or maxDate.Year > minDate.Year) Then
                            'maximum date goes into the next semi-annual portion of the year
                            Me.strDateFilter &= "9/30/" & minDate.Year & "'"
                            tempArray.Add("9/30/" & minDate.Year)
                            minDate = "10/01/" & minDate.Year
                        Else
                            'maximum date does not go into the next semi-annual portion of the year
                            Me.strDateFilter &= maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year & "'"
                            tempArray.Add(maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year)
                            minDate = maxDate.AddDays(1)
                        End If
                    End If

                    'add the date filter
                    tempArray.Add(Me.strDateFilter)
                    Me.dateFilterArr.Add(tempArray)
                End While
            ElseIf (Session("SplitOutput") = "QUARTER") Then
                Dim minDate As Date = CType(Me.dsCommon.Tables("Total").Rows(0).Item("Min1"), Date).Date & ""
                Dim maxDate As Date = CType(Me.dsCommon.Tables("Total").Rows(0).Item("Max1"), Date).Date & ""

                'loop until we've generated date filters for each quarter
                While (minDate <= maxDate)
                    tempArray = New ArrayList

                    'if we're in a fiscal year that matches the current year number
                    If (minDate.Month >= 10 And minDate.Month <= 12) Then
                        Me.strDateFilter = ""
                        Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                            minDate.Month & "/" & minDate.Day & "/" & minDate.Year & _
                            "' and frs.LAST_USED_DATE <= '"
                        tempArray.Add(minDate.Month & "/" & minDate.Day & "/" & minDate.Year)

                        If (maxDate.Month < 10 Or maxDate.Year > minDate.Year) Then
                            'maximum date goes into the next fiscal quarter
                            Me.strDateFilter &= "12/31/" & minDate.Year & "'"
                            tempArray.Add("12/31/" & minDate.Year)
                            minDate = "1/01/" & minDate.Year + 1
                        Else
                            'maximum date does not go into the next fiscal quarter
                            Me.strDateFilter &= maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year & "'"
                            tempArray.Add(maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year)
                            minDate = maxDate.AddDays(1)
                        End If
                    ElseIf (minDate.Month >= 1 And minDate.Month <= 3) Then
                        Me.strDateFilter = ""
                        Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                            minDate.Month & "/" & minDate.Day & "/" & minDate.Year & _
                            "' and frs.LAST_USED_DATE <= '"
                        tempArray.Add(minDate.Month & "/" & minDate.Day & "/" & minDate.Year)

                        If (maxDate.Month > 3 Or maxDate.Year > minDate.Year) Then
                            'maximum date goes into the next fiscal quarter
                            Me.strDateFilter &= "3/31/" & minDate.Year & "'"
                            tempArray.Add("3/31/" & minDate.Year)
                            minDate = "4/01/" & minDate.Year
                        Else
                            'maximum date does not go into the next fiscal quarter
                            Me.strDateFilter &= maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year & "'"
                            tempArray.Add(maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year)
                            minDate = maxDate.AddDays(1)
                        End If
                    ElseIf (minDate.Month >= 4 And minDate.Month <= 6) Then
                        Me.strDateFilter = ""
                        Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                            minDate.Month & "/" & minDate.Day & "/" & minDate.Year & _
                            "' and frs.LAST_USED_DATE <= '"
                        tempArray.Add(minDate.Month & "/" & minDate.Day & "/" & minDate.Year)

                        If (maxDate.Month > 6 Or maxDate.Year > minDate.Year) Then
                            'maximum date goes into the next fiscal quarter
                            Me.strDateFilter &= "6/30/" & minDate.Year & "'"
                            tempArray.Add("6/30/" & minDate.Year)
                            minDate = "7/01/" & minDate.Year
                        Else
                            'maximum date does not go into the next fiscal quarter
                            Me.strDateFilter &= maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year & "'"
                            tempArray.Add(maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year)
                            minDate = maxDate.AddDays(1)
                        End If
                    ElseIf (minDate.Month >= 7 And minDate.Month <= 9) Then
                        Me.strDateFilter = ""
                        Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                            minDate.Month & "/" & minDate.Day & "/" & minDate.Year & _
                            "' and frs.LAST_USED_DATE <= '"
                        tempArray.Add(minDate.Month & "/" & minDate.Day & "/" & minDate.Year)

                        If (maxDate.Month > 9 Or maxDate.Year > minDate.Year) Then
                            'maximum date goes into the next fiscal quarter
                            Me.strDateFilter &= "9/30/" & minDate.Year & "'"
                            tempArray.Add("9/30/" & minDate.Year)
                            minDate = "10/01/" & minDate.Year
                        Else
                            'maximum date does not go into the next fiscal quarter
                            Me.strDateFilter &= maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year & "'"
                            tempArray.Add(maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year)
                            minDate = maxDate.AddDays(1)
                        End If
                    End If

                    'add the date filter
                    tempArray.Add(Me.strDateFilter)
                    Me.dateFilterArr.Add(tempArray)
                End While
            ElseIf (Session("SplitOutput") = "MONTH") Then
                Dim minDate As Date = CType(Me.dsCommon.Tables("Total").Rows(0).Item("Min1"), Date).Date & ""
                Dim maxDate As Date = CType(Me.dsCommon.Tables("Total").Rows(0).Item("Max1"), Date).Date & ""

                'loop until we've generated date filters for each month
                While (minDate <= maxDate)
                    tempArray = New ArrayList

                    'if we're in a fiscal year that matches the current year number
                    If (minDate.Month = 1) Then
                        Me.strDateFilter = ""
                        Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                            minDate.Month & "/" & minDate.Day & "/" & minDate.Year & _
                            "' and frs.LAST_USED_DATE <= '"
                        tempArray.Add(minDate.Month & "/" & minDate.Day & "/" & minDate.Year)

                        If (maxDate.Month > 1 Or maxDate.Year > minDate.Year) Then
                            'maximum date goes into the next month
                            Me.strDateFilter &= "1/31/" & minDate.Year & "'"
                            tempArray.Add("1/31/" & minDate.Year)
                            minDate = "2/01/" & minDate.Year
                        Else
                            'maximum date does not go into the next month
                            Me.strDateFilter &= maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year & "'"
                            tempArray.Add(maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year)
                            minDate = maxDate.AddDays(1)
                        End If
                    ElseIf (minDate.Month = 2) Then
                        Me.strDateFilter = ""
                        Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                            minDate.Month & "/" & minDate.Day & "/" & minDate.Year & _
                            "' and frs.LAST_USED_DATE <= '"
                        tempArray.Add(minDate.Month & "/" & minDate.Day & "/" & minDate.Year)

                        If (maxDate.Month > 2 Or maxDate.Year > minDate.Year) Then
                            'maximum date goes into the next month
                            If (IsDate("2/29/" & minDate.Year)) Then
                                Me.strDateFilter &= "2/29/" & minDate.Year & "'"
                                tempArray.Add("2/29/" & minDate.Year)
                            Else
                                Me.strDateFilter &= "2/28/" & minDate.Year & "'"
                                tempArray.Add("2/28/" & minDate.Year)
                            End If
                            minDate = "3/01/" & minDate.Year
                        Else
                            'maximum date does not go into the next month
                            Me.strDateFilter &= maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year & "'"
                            tempArray.Add(maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year)
                            minDate = maxDate.AddDays(1)
                        End If
                    ElseIf (minDate.Month = 3) Then
                        Me.strDateFilter = ""
                        Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                            minDate.Month & "/" & minDate.Day & "/" & minDate.Year & _
                            "' and frs.LAST_USED_DATE <= '"
                        tempArray.Add(minDate.Month & "/" & minDate.Day & "/" & minDate.Year)

                        If (maxDate.Month > 3 Or maxDate.Year > minDate.Year) Then
                            'maximum date goes into the next month
                            Me.strDateFilter &= "3/31/" & minDate.Year & "'"
                            tempArray.Add("3/31/" & minDate.Year)
                            minDate = "4/01/" & minDate.Year
                        Else
                            'maximum date does not go into the next month
                            Me.strDateFilter &= maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year & "'"
                            tempArray.Add(maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year)
                            minDate = maxDate.AddDays(1)
                        End If
                    ElseIf (minDate.Month = 4) Then
                        Me.strDateFilter = ""
                        Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                            minDate.Month & "/" & minDate.Day & "/" & minDate.Year & _
                            "' and frs.LAST_USED_DATE <= '"
                        tempArray.Add(minDate.Month & "/" & minDate.Day & "/" & minDate.Year)

                        If (maxDate.Month > 4 Or maxDate.Year > minDate.Year) Then
                            'maximum date goes into the next month
                            Me.strDateFilter &= "4/30/" & minDate.Year & "'"
                            tempArray.Add("4/30/" & minDate.Year)
                            minDate = "5/01/" & minDate.Year
                        Else
                            'maximum date does not go into the next month
                            Me.strDateFilter &= maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year & "'"
                            tempArray.Add(maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year)
                            minDate = maxDate.AddDays(1)
                        End If
                    ElseIf (minDate.Month = 5) Then
                        Me.strDateFilter = ""
                        Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                            minDate.Month & "/" & minDate.Day & "/" & minDate.Year & _
                            "' and frs.LAST_USED_DATE <= '"
                        tempArray.Add(minDate.Month & "/" & minDate.Day & "/" & minDate.Year)

                        If (maxDate.Month > 5 Or maxDate.Year > minDate.Year) Then
                            'maximum date goes into the next month
                            Me.strDateFilter &= "5/31/" & minDate.Year & "'"
                            tempArray.Add("5/31/" & minDate.Year)
                            minDate = "6/01/" & minDate.Year
                        Else
                            'maximum date does not go into the next month
                            Me.strDateFilter &= maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year & "'"
                            tempArray.Add(maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year)
                            minDate = maxDate.AddDays(1)
                        End If
                    ElseIf (minDate.Month = 6) Then
                        Me.strDateFilter = ""
                        Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                            minDate.Month & "/" & minDate.Day & "/" & minDate.Year & _
                            "' and frs.LAST_USED_DATE <= '"
                        tempArray.Add(minDate.Month & "/" & minDate.Day & "/" & minDate.Year)

                        If (maxDate.Month > 6 Or maxDate.Year > minDate.Year) Then
                            'maximum date goes into the next month
                            Me.strDateFilter &= "6/30/" & minDate.Year & "'"
                            tempArray.Add("6/30/" & minDate.Year)
                            minDate = "7/01/" & minDate.Year
                        Else
                            'maximum date does not go into the next month
                            Me.strDateFilter &= maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year & "'"
                            tempArray.Add(maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year)
                            minDate = maxDate.AddDays(1)
                        End If
                    ElseIf (minDate.Month = 7) Then
                        Me.strDateFilter = ""
                        Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                            minDate.Month & "/" & minDate.Day & "/" & minDate.Year & _
                            "' and frs.LAST_USED_DATE <= '"
                        tempArray.Add(minDate.Month & "/" & minDate.Day & "/" & minDate.Year)

                        If (maxDate.Month > 7 Or maxDate.Year > minDate.Year) Then
                            'maximum date goes into the next month
                            Me.strDateFilter &= "7/31/" & minDate.Year & "'"
                            tempArray.Add("7/31/" & minDate.Year)
                            minDate = "8/01/" & minDate.Year
                        Else
                            'maximum date does not go into the next month
                            Me.strDateFilter &= maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year & "'"
                            tempArray.Add(maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year)
                            minDate = maxDate.AddDays(1)
                        End If
                    ElseIf (minDate.Month = 8) Then
                        Me.strDateFilter = ""
                        Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                            minDate.Month & "/" & minDate.Day & "/" & minDate.Year & _
                            "' and frs.LAST_USED_DATE <= '"
                        tempArray.Add(minDate.Month & "/" & minDate.Day & "/" & minDate.Year)

                        If (maxDate.Month > 8 Or maxDate.Year > minDate.Year) Then
                            'maximum date goes into the next month
                            Me.strDateFilter &= "8/31/" & minDate.Year & "'"
                            tempArray.Add("8/31/" & minDate.Year)
                            minDate = "9/01/" & minDate.Year
                        Else
                            'maximum date does not go into the next month
                            Me.strDateFilter &= maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year & "'"
                            tempArray.Add(maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year)
                            minDate = maxDate.AddDays(1)
                        End If
                    ElseIf (minDate.Month = 9) Then
                        Me.strDateFilter = ""
                        Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                            minDate.Month & "/" & minDate.Day & "/" & minDate.Year & _
                            "' and frs.LAST_USED_DATE <= '"
                        tempArray.Add(minDate.Month & "/" & minDate.Day & "/" & minDate.Year)

                        If (maxDate.Month > 9 Or maxDate.Year > minDate.Year) Then
                            'maximum date goes into the next month
                            Me.strDateFilter &= "9/30/" & minDate.Year & "'"
                            tempArray.Add("9/30/" & minDate.Year)
                            minDate = "10/01/" & minDate.Year
                        Else
                            'maximum date does not go into the next month
                            Me.strDateFilter &= maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year & "'"
                            tempArray.Add(maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year)
                            minDate = maxDate.AddDays(1)
                        End If
                    ElseIf (minDate.Month = 10) Then
                        Me.strDateFilter = ""
                        Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                            minDate.Month & "/" & minDate.Day & "/" & minDate.Year & _
                            "' and frs.LAST_USED_DATE <= '"
                        tempArray.Add(minDate.Month & "/" & minDate.Day & "/" & minDate.Year)

                        If (maxDate.Month > 10 Or maxDate.Year > minDate.Year) Then
                            'maximum date goes into the next month
                            Me.strDateFilter &= "10/31/" & minDate.Year & "'"
                            tempArray.Add("10/31/" & minDate.Year)
                            minDate = "11/01/" & minDate.Year
                        Else
                            'maximum date does not go into the next month
                            Me.strDateFilter &= maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year & "'"
                            tempArray.Add(maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year)
                            minDate = maxDate.AddDays(1)
                        End If
                    ElseIf (minDate.Month = 11) Then
                        Me.strDateFilter = ""
                        Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                            minDate.Month & "/" & minDate.Day & "/" & minDate.Year & _
                            "' and frs.LAST_USED_DATE <= '"
                        tempArray.Add(minDate.Month & "/" & minDate.Day & "/" & minDate.Year)

                        If (maxDate.Month > 11 Or maxDate.Year > minDate.Year) Then
                            'maximum date goes into the next month
                            Me.strDateFilter &= "11/30/" & minDate.Year & "'"
                            tempArray.Add("11/30/" & minDate.Year)
                            minDate = "12/01/" & minDate.Year
                        Else
                            'maximum date does not go into the next month
                            Me.strDateFilter &= maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year & "'"
                            tempArray.Add(maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year)
                            minDate = maxDate.AddDays(1)
                        End If
                    ElseIf (minDate.Month = 12) Then
                        Me.strDateFilter = ""
                        Me.strDateFilter &= " and frs.LAST_USED_DATE >= '" & _
                            minDate.Month & "/" & minDate.Day & "/" & minDate.Year & _
                            "' and frs.LAST_USED_DATE <= '"
                        tempArray.Add(minDate.Month & "/" & minDate.Day & "/" & minDate.Year)

                        If (maxDate.Month < 12 And maxDate.Year > minDate.Year) Then
                            'maximum date goes into the first month of the next year
                            Me.strDateFilter &= "12/31/" & minDate.Year & "'"
                            tempArray.Add("12/31/" & minDate.Year)
                            minDate = "1/01/" & minDate.Year + 1
                        Else
                            'maximum date does not go into the next month
                            Me.strDateFilter &= maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year & "'"
                            tempArray.Add(maxDate.Month & "/" & maxDate.Day & "/" & maxDate.Year)
                            minDate = maxDate.AddDays(1)
                        End If
                    End If

                    'add the date filter
                    tempArray.Add(Me.strDateFilter)
                    Me.dateFilterArr.Add(tempArray)
                End While
            End If
            Return True
        Catch ex As Exception
            Session("Alert") = "Date filter creation failure.  " & ex.ToString
            Return False
        End Try
    End Function

    'generates the data filter for the report results
    Private Function setPreliminaryDataFilter(ByVal dateFilter As String) As Boolean
        Trace.Warn("Setting Data Filter")
        Try
            'Set up the filter
            Dim filterArray As Array
            filterArray = CType(Session("FilterItems"), String).Split(",")
            Me.strDataFilter = ""

            'filter on filter items
            Dim filterItem As String
            For Each filterItem In filterArray
                Dim tempfilterItemArray As Array = filterItem.Split("^")

                If (tempfilterItemArray(1) <> -1) Then
                    strDataFilter &= " and frs.RESPONDENT_USER_KEY in (Select distinct fr.RESPONDENT_USER_KEY from " & myUtility.getExtension() & "SAT_RESPONSE_RESULT fr, " & _
                        myUtility.getExtension() & "SAT_RESPONSE frs"
                    If (tempfilterItemArray(3) = "C") Then
                        strDataFilter &= " where fr.QUESTION_ID = " & tempfilterItemArray(0) & " and fr.ANSWER_VALUE <> 0 " & _
                            "and charindex(fr.ANSWER_TEXT,'" & tempfilterItemArray(1) & "') <> 0 and fr.ANSWER_VALUE = '" & _
                            tempfilterItemArray(1) & "' and fr.SURVEY_KEY = " & _
                            Session("seqSurvey") & " fr.SURVEY_KEY = frs.SURVEY_KEY and fr.RESPONDENT_USER_KEY = frs.RESPONDENT_USER_KEY " & _
                            " " & dateFilter & ")"
                    ElseIf (tempfilterItemArray(3) = "D" Or tempfilterItemArray(3) = "S") Then
                        strDataFilter &= " where fr.QUESTION_ID = " & tempfilterItemArray(0) & " and fr.ANSWER_TEXT = '" & _
                            CType(tempfilterItemArray(4), String).Replace("'", "''").Replace("~", ",") & "' and fr.SURVEY_KEY = " & Session("seqSurvey") & " and " & _
                            "fr.SURVEY_KEY = frs.SURVEY_KEY and fr.RESPONDENT_USER_KEY = frs.RESPONDENT_USER_KEY " & _
                            dateFilter & ")"
                    ElseIf (tempfilterItemArray(3) = "Z") Then
                        strDataFilter &= " where fr.QUESTION_ID = " & tempfilterItemArray(0) & " and fr.ANSWER_TEXT = '" & _
                            CType(tempfilterItemArray(1), String).Replace("'", "''").Replace("~", ",") & "^" & _
                            CType(tempfilterItemArray(4), String).Replace("'", "''").Replace("~", ",") & "' and fr.SURVEY_KEY = " & Session("seqSurvey") & _
                            dateFilter & ")"
                    Else
                        strDataFilter &= " where fr.QUESTION_ID = " & tempfilterItemArray(0) & " and fr.ANSWER_VALUE = '" & _
                           tempfilterItemArray(1) & "' and fr.SURVEY_KEY = " & Session("seqSurvey") & " and " & _
                           "fr.SURVEY_KEY = frs.SURVEY_KEY and fr.RESPONDENT_USER_KEY = frs.RESPONDENT_USER_KEY " & _
                           dateFilter & ")"
                    End If
                End If
            Next
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Session("Alert") = "Data filter creation failure.  " & ex.ToString
            Return False
        End Try
    End Function

    'generates the data filter for the report results
    Private Function setDataFilter(ByVal dateFilter As String) As Boolean
        Trace.Warn("Setting Data Filter")
        Try
            'Set up the filter
            Dim filterArray As Array
            filterArray = CType(Session("FilterItems"), String).Split(",")
            Me.strDataFilter = ""

            'filter on filter items
            Dim filterItem As String
            For Each filterItem In filterArray
                Dim tempfilterItemArray As Array = filterItem.Split("^")

                If (tempfilterItemArray(1) <> -1) Then
                    strDataFilter &= " and frs.RESPONDENT_USER_KEY in (Select distinct fr.RESPONDENT_USER_KEY from " & myUtility.getExtension() & "SAT_RESPONSE_RESULT fr, " & _
                        myUtility.getExtension() & "SAT_RESPONSE frs"
                    If (tempfilterItemArray(3) = "C") Then
                        strDataFilter &= " where fr.QUESTION_ID = " & tempfilterItemArray(0) & " and fr.ANSWER_VALUE <> 0 " & _
                            "and charindex(fr.ANSWER_TEXT,'" & tempfilterItemArray(1) & "') <> 0 and fr.ANSWER_VALUE = '" & _
                            tempfilterItemArray(1) & "' and fr.SURVEY_KEY = " & _
                            Session("seqSurvey") & " fr.SURVEY_KEY = frs.SURVEY_KEY and fr.RESPONDENT_USER_KEY = frs.RESPONDENT_USER_KEY " & _
                            "and fr.RESPONSE_DATE <= frs.LAST_COMPLETION_DATE and frs.LAST_COMPLETION_DATE <> '1/1/1970' " & dateFilter & ")"
                    ElseIf (tempfilterItemArray(3) = "D" Or tempfilterItemArray(3) = "S") Then
                        strDataFilter &= " where fr.QUESTION_ID = " & tempfilterItemArray(0) & " and fr.ANSWER_TEXT = '" & _
                            CType(tempfilterItemArray(4), String).Replace("'", "''").Replace("~", ",") & "' and fr.SURVEY_KEY = " & Session("seqSurvey") & " and " & _
                            "fr.SURVEY_KEY = frs.SURVEY_KEY and fr.RESPONDENT_USER_KEY = frs.RESPONDENT_USER_KEY and fr.RESPONSE_DATE <= " & _
                            "frs.LAST_COMPLETION_DATE and frs.LAST_COMPLETION_DATE <> '1/1/1970' " & dateFilter & ")"
                    ElseIf (tempfilterItemArray(3) = "Z") Then
                        strDataFilter &= " where fr.QUESTION_ID = " & tempfilterItemArray(0) & " and fr.ANSWER_TEXT = '" & _
                            CType(tempfilterItemArray(1), String).Replace("'", "''").Replace("~", ",") & "^" & _
                            CType(tempfilterItemArray(4), String).Replace("'", "''").Replace("~", ",") & "' and fr.SURVEY_KEY = " & Session("seqSurvey") & " and " & _
                            "fr.SURVEY_KEY = frs.SURVEY_KEY and fr.RESPONDENT_USER_KEY = frs.RESPONDENT_USER_KEY and fr.RESPONSE_DATE <= " & _
                            "frs.LAST_COMPLETION_DATE and frs.LAST_COMPLETION_DATE <> '1/1/1970' " & dateFilter & ")"
                    Else
                        strDataFilter &= " where fr.QUESTION_ID = " & tempfilterItemArray(0) & " and fr.ANSWER_VALUE = '" & _
                           tempfilterItemArray(1) & "' and fr.SURVEY_KEY = " & Session("seqSurvey") & " and " & _
                           "fr.SURVEY_KEY = frs.SURVEY_KEY and fr.RESPONDENT_USER_KEY = frs.RESPONDENT_USER_KEY and fr.RESPONSE_DATE <= " & _
                           "frs.LAST_COMPLETION_DATE and frs.LAST_COMPLETION_DATE <> '1/1/1970' " & dateFilter & ")"
                    End If
                End If
            Next
            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Session("Alert") = "Data filter creation failure.  " & ex.ToString
            Return False
        End Try
    End Function

    'generates the question filter for the report results
    Private Function setQuestionFilter(ByVal QFilter As ArrayList, Optional ByVal textOutput As Boolean = False) As Boolean
        Trace.Warn("Setting Question Filter")
        Try
            Me.strQuestionFilter = ""
            If (QFilter.Count > 0) Then
                Dim item As Integer = 0
                For Each item In QFilter
                    Me.strQuestionFilter &= item & ","
                Next
            Else
                'Set up the filter
                Dim filterArray As Array
                filterArray = CType(Session("DataItems"), String).Split(",")
                Me.strQuestionFilter = ""

                'filter on filter items
                Dim filterItem As String
                For Each filterItem In filterArray
                    Dim tempfilterItemArray As Array = filterItem.Split("^")

                    If (tempfilterItemArray(1) = "True") Then
                        Me.strQuestionFilter &= tempfilterItemArray(0) & ","
                    End If
                Next
            End If

            'if we have questions we're showing then remove the last comma and add the in clause
            If (Me.strQuestionFilter.Length > 0) Then
                Me.strQuestionFilter = Me.strQuestionFilter.Substring(0, Me.strQuestionFilter.Length - 1)
                Me.strQuestionFilter = Me.strQuestionFilter.Insert(0, " and frs.QUESTION_ID in (")
                Me.strQuestionFilter = Me.strQuestionFilter.Insert(Me.strQuestionFilter.Length, ")")
                Dim tempString As String = ""

                'if doing text output then only get the M and S type questions.
                If (textOutput) Then
                    tempString = " and frs.QUESTION_TYPE in ('M','S') "
                Else
                    tempString = " and frs.QUESTION_TYPE in ('C','L','P','R','T','N','D') "
                End If
                Me.strQuestionFilter = Me.strQuestionFilter.Insert(0, tempString)
            End If
            Return True
        Catch ex As Exception
            Session("Alert") = "Data filter creation failure.  " & ex.ToString
            Return False
        End Try
    End Function
#End Region

#Region "Report Processing"
    'processes the report header information
    Private Function processReportHeader() As Boolean
        Trace.Warn("Processing report header")
        Try
            'set up the data adapter
            Me.sqlCommonAdapter = New SqlClient.SqlDataAdapter
            Me.sqlCommonCommand = New SqlClient.SqlCommand
            Me.dsCommon = New DataSet
            Me.sqlCommonAdapter.SelectCommand = Me.sqlCommonCommand
            Me.sqlCommonCommand.Connection = Me.commonConn

            'Responded/Total Count
            Me.sqlCommonCommand.CommandText = "Select count(*) as Total, " & _
                "(select count(frs.LAST_COMPLETION_DATE) from " & myUtility.getExtension() & "SAT_RESPONSE frs where frs.SURVEY_KEY = " & _
                Session("seqSurvey") & " and frs.LAST_COMPLETION_DATE <> '1/1/1970' " & _
                Me.strDateFilter & " " & Me.strDataFilter & ") as Responded, min(CREATE_DATE) as Min1, " & _
                "max(CREATE_DATE) as Max1 from " & myUtility.getExtension() & "SAT_RESPONSE frs where SURVEY_KEY = " & _
                Session("seqSurvey") & " " & Me.strDataFilter & " " & Me.strDateFilter
            Trace.Warn(Me.sqlCommonCommand.CommandText)
            Me.sqlCommonAdapter.Fill(Me.dsCommon, "Total")

            If (Me.dsCommon.Tables("Total").Rows(0).Item("Total") > 0) Then
                Return True
            Else
                Session("Alert") = "There is no data for time period selected."
                Return False
            End If
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Session("Alert") = "Failed to get survey data."
            Return False
        End Try
    End Function

    'processes the report filter section
    Private Function processReportFilters() As Boolean
        Trace.Warn("processing report filters")
        Try
            'define the variables
            Dim genericCell As System.Web.UI.WebControls.TableCell
            Dim genericRow As System.Web.UI.WebControls.TableRow

            'Apply the data filter if it is not the empty string
            Session("FiltersSelected") = False
            If (Me.strDataFilter <> "") Then
                Dim dateFilter As ArrayList
                For Each dateFilter In Me.dateFilterArr
                    'Responded Count
                    If (setDataFilter(dateFilter(2))) Then
                        Me.sqlCommonCommand.CommandText = "select count(*) as responded from (select fr.RESPONDENT_USER_KEY, " & _
                            "count(*) userCount from " & myUtility.getExtension() & "SAT_RESPONSE_RESULT fr, " & _
                            myUtility.getExtension() & "SAT_RESPONSE frs where fr.SURVEY_KEY = " & _
                            Session("seqSurvey") & " and fr.SURVEY_KEY = frs.SURVEY_KEY and fr.RESPONDENT_USER_KEY = frs.RESPONDENT_USER_KEY " & _
                            " and fr.RESPONSE_DATE <= frs.LAST_COMPLETION_DATE and frs.LAST_COMPLETION_DATE <> '1/1/1970' " & dateFilter(2) & _
                            " " & Me.strDataFilter & " group by fr.RESPONDENT_USER_KEY) a"
                        Trace.Warn(Me.sqlCommonCommand.CommandText)
                        Me.sqlCommonAdapter.Fill(Me.dsCommon, "Filters")

                        'Set up the filter
                        Dim filterArray As Array
                        filterArray = CType(Session("FilterItems"), String).Split(",")

                        'get filter item data
                        Dim filterItem As String
                        For Each filterItem In filterArray
                            Dim tempFilterItemArray As Array = filterItem.Split("^")
                            If (tempFilterItemArray(1) <> -1) Then
                                Session("FiltersSelected") = True

                                'generate a row for the question text
                                genericRow = New System.Web.UI.WebControls.TableRow
                                genericCell = New System.Web.UI.WebControls.TableCell
                                genericCell.CssClass = "Content"
                                'genericCell.Width = Unit.Percentage(50)
                                genericCell.Text = CType(tempFilterItemArray(2), String).Replace("~", ",")
                                genericCell.Wrap = True
                                genericCell.EnableViewState = True
                                genericRow.Cells.Add(genericCell)

                                'generate a row for the question answer
                                genericCell = New System.Web.UI.WebControls.TableCell
                                genericCell.CssClass = "ReportValueBlue"
                                'genericCell.Width = Unit.Percentage(50)
                                genericCell.Text = CType(tempFilterItemArray(4), String).Replace("~", ",")
                                genericCell.Wrap = False
                                genericCell.EnableViewState = True
                                genericRow.Cells.Add(genericCell)
                                If (Me.ReportFilter.Rows.Count Mod 2 = 0) Then
                                    genericRow.BackColor = Color.LightGray
                                Else
                                    genericRow.BackColor = Color.White
                                End If
                                genericRow.EnableViewState = True
                                Me.ReportFilter.Rows.Add(genericRow)
                                Me.ReportFilter.EnableViewState = True
                            End If
                        Next

                        If (Me.dsCommon.Tables("Filters").Rows.Count() > 0) Then
                            'generate a row for the filtered number of respondents
                            genericRow = New System.Web.UI.WebControls.TableRow
                            genericCell = New System.Web.UI.WebControls.TableCell
                            genericCell.CssClass = "Content"
                            genericCell.Text = "Filtered Respondents:"
                            genericCell.Wrap = False
                            genericCell.EnableViewState = True
                            genericCell.HorizontalAlign = HorizontalAlign.Right
                            genericRow.Cells.Add(genericCell)

                            'generate a cell for the number of filtered respondents
                            genericCell = New System.Web.UI.WebControls.TableCell
                            genericCell.CssClass = "ReportValueRed"
                            genericCell.Text = Me.dsCommon.Tables("Filters").Rows(0).Item("responded")
                            genericCell.Wrap = False
                            genericCell.EnableViewState = True
                            genericCell.HorizontalAlign = HorizontalAlign.Left
                            genericRow.Cells.Add(genericCell)

                            'Add the row to the table
                            genericRow.EnableViewState = True
                            Me.ReportFilterResults.Rows.Add(genericRow)
                            Me.ReportFilterResults.EnableViewState = True
                            Return True
                        Else
                            'generate a row for the filtered number of respondents
                            genericRow = New System.Web.UI.WebControls.TableRow
                            genericCell = New System.Web.UI.WebControls.TableCell
                            genericCell.CssClass = "Content"
                            genericCell.Text = "Filtered Respondents:"
                            genericCell.Wrap = False
                            genericCell.EnableViewState = True
                            genericCell.HorizontalAlign = HorizontalAlign.Right
                            genericRow.Cells.Add(genericCell)

                            'generate a row for the question answer
                            genericCell = New System.Web.UI.WebControls.TableCell
                            genericCell.CssClass = "ReportValueRed"
                            genericCell.Text = "0"
                            genericCell.Wrap = False
                            genericCell.EnableViewState = True
                            genericCell.HorizontalAlign = HorizontalAlign.Left
                            genericRow.Cells.Add(genericCell)

                            'Add the row to the table
                            genericRow.EnableViewState = True
                            Me.ReportFilterResults.Rows.Add(genericRow)
                            Me.ReportFilterResults.EnableViewState = True
                            Return True
                        End If
                    Else
                        Session("Alert") = "Failed to get survey data."
                        Return False
                    End If
                Next
            Else
                Return True
            End If
        Catch ex As Exception
            Session("Alert") = "Failed to get survey data.  " & ex.ToString
            Return False
        End Try
    End Function

    'processes the ratings part of the data section
    Private Function processRatings() As Boolean
        Trace.Warn("processing ratings section")
        Try
            'define the variables
            Dim genericCell As System.Web.UI.WebControls.TableCell
            Dim genericRow As System.Web.UI.WebControls.TableRow
            Dim genericTable As System.Web.UI.WebControls.Table
            Dim genericLabel As System.Web.UI.WebControls.Label
            Dim genericChart As Infragistics.WebUI.UltraWebChart.UltraChart

            Dim dateFilter As ArrayList
            Dim runCount As Integer = 0
            Dim blnDoOnce As Boolean = False
            Dim blnDoOnce2 As Boolean = False
            Dim blnDoOnce3 As Boolean = False
            Dim isFirstTable As Boolean = True
            Dim questionCounter As Integer = 1
            For Each dateFilter In Me.dateFilterArr
                If (setDataFilter(dateFilter(2))) Then
                    'Set up the filter
                    Dim filterArray As Array
                    Dim skipList As ArrayList = New ArrayList
                    Dim skipListText As ArrayList = New ArrayList
                    filterArray = CType(Session("DataItems"), String).Split(",")

                    If (Not blnDoOnce) Then
                        'generate a row for the question text
                        'genericRow = New System.Web.UI.WebControls.TableRow
                        'genericRow.CssClass = "BreakBefore"
                        'genericCell = New System.Web.UI.WebControls.TableCell
                        'genericLabel = New System.web.ui.webcontrols.Label
                        'genericLabel.Text = "<hr style='BORDER-TOP-STYLE: ridge' width='100%' SIZE='3'/>"
                        'genericCell.Controls.Add(genericLabel)
                        'genericCell.Wrap = False
                        'genericCell.EnableViewState = True
                        'genericCell.ColumnSpan = 2
                        'genericRow.Cells.Add(genericCell)
                        'genericRow.EnableViewState = True
                        'Me.DataTable.Rows.Add(genericRow)

                        'section header
                        genericRow = New System.Web.UI.WebControls.TableRow
                        genericCell = New System.Web.UI.WebControls.TableCell
                        genericCell.Wrap = False
                        genericCell.CssClass = "sectionHeadTeal"
                        genericCell.Text = "Ratings Results"
                        genericCell.EnableViewState = True
                        If (Session("SplitOutput") <> "NONE") Then
                            genericCell.ColumnSpan = 3
                        Else
                            genericCell.ColumnSpan = 2
                        End If
                        genericRow.Cells.Add(genericCell)
                        genericRow.EnableViewState = True
                        Me.DataTable.Rows.Add(genericRow)
                        blnDoOnce = True
                    End If

                    If (Session("SplitOutput") <> "NONE" And filterArray.Length > 0) Then
                        'buffer zone
                        'genericRow = New System.Web.UI.WebControls.TableRow
                        'genericCell = New System.Web.UI.WebControls.TableCell
                        'genericCell.Wrap = False
                        'genericCell.Height = Unit.Pixel(40)
                        'genericCell.EnableViewState = True
                        'genericCell.ColumnSpan = 3
                        'genericRow.Cells.Add(genericCell)
                        'genericRow.EnableViewState = True
                        'Me.DataTable.Rows.Add(genericRow)

                        'generate a row for the date range
                        genericRow = New System.Web.UI.WebControls.TableRow
                        genericCell = New System.Web.UI.WebControls.TableCell
                        genericLabel = New System.web.ui.webcontrols.Label
                        genericLabel.CssClass = "ReportHeaderData"
                        genericLabel.Text = "Period of Record Set: "
                        genericCell.Controls.Add(genericLabel)
                        genericLabel = New System.web.ui.webcontrols.Label
                        genericLabel.CssClass = "ReportHeaderBlue"
                        genericLabel.Text = dateFilter(0) & " - " & dateFilter(1)
                        genericCell.Controls.Add(genericLabel)
                        genericCell.Wrap = True
                        genericCell.EnableViewState = True
                        genericCell.ColumnSpan = 3
                        genericRow.Cells.Add(genericCell)
                        genericRow.EnableViewState = True
                        genericRow.BackColor = Color.LightGray
                        Me.DataTable.Rows.Add(genericRow)
                    End If

                    'get filter item data
                    Dim filterItem As String = ""
                    Session("QuestionsSelected") = False
                    Session("RItemExists") = False
                    isFirstTable = True
                    For Each filterItem In filterArray
                        blnDoOnce2 = False
                        Session("QuestionsSelected") = True
                        Dim blnPercentDisplay As Boolean = False
                        Dim tempFilterItemArray As Array = filterItem.Split("^")
                        If (CType(tempFilterItemArray(1), Boolean) And tempFilterItemArray(2) = "R") Then
                            Session("RItemExists") = True
                            'process type R questions with averages and the question label
                            Dim dr As DataRow

                            If Not (skipList.Contains(CType(tempFilterItemArray(0), String))) Then
                                'get the list of questions in the multi-group if the question chart type is multi
                                Dim tempSkipList As ArrayList = New ArrayList
                                Dim tempSkipListText As ArrayList = New ArrayList
                                If (tempFilterItemArray(6) = "MULTI") Then
                                    Dim tempItem As String = ""
                                    For Each tempItem In filterArray
                                        Dim tempSplit As Array = tempItem.Split("^")
                                        If (CType(tempSplit(0), Integer) > CType(tempFilterItemArray(0), Integer) And CType(tempSplit(1), Boolean) _
                                            And tempSplit(6) = "MULTI") Then
                                            If (tempSplit(5) = tempFilterItemArray(5)) Then
                                                tempSkipList.Add(tempSplit(0))
                                                tempSkipListText.Add(tempSplit(3))
                                                If (tempSplit(4) = "P-True") Then
                                                    blnPercentDisplay = True
                                                End If
                                            End If
                                        ElseIf (CType(tempSplit(0), Integer) = CType(tempFilterItemArray(0), Integer)) Then
                                            tempSkipListText.Add(tempSplit(3))
                                            If (tempSplit(4) = "P-True") Then
                                                blnPercentDisplay = True
                                            End If
                                        End If
                                    Next
                                Else
                                    If (tempFilterItemArray(4) = "P-True") Then
                                        blnPercentDisplay = True
                                    End If
                                End If

                                'build skiplist string for queries
                                Dim strSkipList As String = tempFilterItemArray(0)
                                Dim index As Integer = 0
                                While (index < tempSkipList.Count)
                                    strSkipList &= ", " & tempSkipList(index)
                                    index += 1
                                End While
                                tempSkipList.Add(tempFilterItemArray(0))

                                'add skip items to master list
                                index = 0
                                skipListText.Clear()
                                skipListText.Add(tempFilterItemArray(3))
                                While (index < tempSkipList.Count - 1)
                                    skipList.Add(tempSkipList(index))
                                    skipListText.Add(tempSkipListText(index))
                                    index += 1
                                End While
                                skipList.Add(tempFilterItemArray(0))

                                'set the question filter again
                                If (tempFilterItemArray(6) = "MULTI") Then
                                    setQuestionFilter(tempSkipList)
                                Else
                                    tempSkipList = New ArrayList
                                    setQuestionFilter(tempSkipList)
                                End If

                                Me.sqlCommonCommand.CommandText = "select frs.QUESTION_ID, frs.ANSWER_VALUE, count(frs.ANSWER_VALUE) as " & _
                                "ValueCount, frs.ANSWER_TEXT, fq.QUESTION_GROUP_ID, fq.QUESTION_TEXT as QuestionText from " & _
                                myUtility.getExtension() & "SAT_RESPONSE_RESULT frs, " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION fq " & _
                                "where frs.SURVEY_KEY = " & Session("seqSurvey") & " " & _
                                dateFilter(2).Replace("LAST_USED_DATE", "RESPONSE_DATE") & " " & Me.strDataFilter & " " & _
                                Me.strQuestionFilter & " and frs.QUESTION_ID = fq.QUESTION_ID and frs.QUESTION_TYPE = fq.QUESTION_TYPE_CODE " & _
                                "and fq.TEMPLATE_KEY = " & Session("seqTemplate") & " group by frs.QUESTION_ID, frs.ANSWER_VALUE, " & _
                                "frs.ANSWER_TEXT, fq.QUESTION_GROUP_ID, fq.QUESTION_TEXT order by frs.QUESTION_ID, frs.ANSWER_VALUE, frs.ANSWER_TEXT, " & _
                                "fq.QUESTION_GROUP_ID, fq.QUESTION_TEXT"
                                Trace.Warn(Me.sqlCommonCommand.CommandText)
                                If (Me.dsCommon.Tables.Contains("DataDump")) Then
                                    Me.dsCommon.Tables("DataDump").Rows.Clear()
                                End If
                                Me.sqlCommonAdapter.Fill(Me.dsCommon, "DataDump")

                                'get a count of questions in the group if dealing with a multi-chart set
                                Dim multiCount As Integer = 1
                                If (tempFilterItemArray(6) = "MULTI") Then
                                    multiCount = tempSkipList.Count()
                                End If

                                Dim sumAnswered(multiCount - 1) As Double
                                Dim sumAverage(multiCount - 1) As Double
                                Dim average(multiCount - 1) As Double

                                'generate the table
                                Dim dt As DataTable = New DataTable
                                dt.Columns.Add("Value")
                                dt.Columns("Value").DataType = Me.dsCommon.Tables("DataDump").Columns("ANSWER_VALUE").DataType
                                dt.TableName = "Question " & tempFilterItemArray(0)

                                'create the rest of the table dynamically
                                If (tempFilterItemArray(6) = "MULTI") Then
                                    index = 1
                                    While (index <= multiCount)
                                        'generated for the count of questions in the group
                                        dt.Columns.Add("People" & index)
                                        dt.Columns("People" & index).DataType = Me.dsCommon.Tables("DataDump").Columns("ValueCount").DataType
                                        dt.Columns.Add("Percentage" & index)
                                        dt.Columns("Percentage" & index).DataType = System.Type.GetType("System.Decimal", True, True)
                                        'dt.Columns.Add("Date"&index)
                                        'dt.Columns("Date"&index).DataType = Me.dsCommon.Tables("ChartQuestion").Columns("datDate").DataType
                                        index += 1
                                    End While
                                Else
                                    'generated for the count of questions in the group
                                    dt.Columns.Add("People")
                                    dt.Columns("People").DataType = Me.dsCommon.Tables("DataDump").Columns("ValueCount").DataType
                                    dt.Columns.Add("Percentage")
                                    dt.Columns("Percentage").DataType = System.Type.GetType("System.Decimal", True, True)
                                    'dt.Columns.Add("Date"&index)
                                    'dt.Columns("Date"&index).DataType = Me.dsCommon.Tables("ChartQuestion").Columns("datDate").DataType
                                End If

                                'get the rating information
                                Me.sqlCommonCommand.CommandText = "select RATING_COUNT, RATING_STEP_VALUE, RATING_INITIAL_VALUE, " & _
                                    "RATING_INITIAL_VALUE + (RATING_COUNT * RATING_STEP_VALUE) - 1 as maxValue, NOT_APPLICABLE_IND from " & _
                                    myUtility.getExtension() & "SAT_TEMPLATE_QUESTION where QUESTION_ID in (" & strSkipList & _
                                    ") and TEMPLATE_KEY = " & Session("seqTemplate")
                                Trace.Warn(Me.sqlCommonCommand.CommandText)
                                If (Me.dsCommon.Tables.Contains("RatingData")) Then
                                    Me.dsCommon.Tables("RatingData").Rows.Clear()
                                End If
                                Me.sqlCommonAdapter.Fill(Me.dsCommon, "RatingData")

                                'add default values and the item values to the table
                                Dim row As DataRow
                                If (tempFilterItemArray(6) = "MULTI") Then
                                    For Each row In Me.dsCommon.Tables("RatingData").Rows
                                        If (dt.Rows.Count = 0) Then
                                            Dim intCount As Integer = 0
                                            Dim intValue As Integer = row.Item("RATING_INITIAL_VALUE")
                                            While (intCount < row.Item("RATING_COUNT"))
                                                Dim tempRow As DataRow = dt.NewRow
                                                tempRow.Item(0) = intValue
                                                index = 1
                                                While (index <= multiCount)
                                                    tempRow.Item("People" & index) = 0
                                                    tempRow.Item("Percentage" & index) = 0
                                                    index += 1
                                                End While
                                                'tempRow.Item(2) = "1/1/1970"
                                                dt.Rows.Add(tempRow)
                                                intValue += row.Item("RATING_STEP_VALUE")
                                                intCount += 1
                                            End While
                                        Else
                                            Dim intCount As Integer = 0
                                            Dim intValue As Integer = row.Item("RATING_INITIAL_VALUE")
                                            While (intCount < row.Item("RATING_COUNT"))
                                                Dim rowIndex As Integer = dt.Rows.Count()
                                                While (rowIndex > 0)
                                                    If (intValue = CType(dt.Rows(rowIndex - 1).Item("Value"), Integer)) Then
                                                        Exit While
                                                    ElseIf (intValue > CType(dt.Rows(rowIndex - 1).Item("Value"), Integer)) Then
                                                        Dim tempRow As DataRow = dt.NewRow
                                                        tempRow.Item(0) = intValue
                                                        index = 1
                                                        While (index <= multiCount)
                                                            tempRow.Item("People" & index) = 0
                                                            tempRow.Item("Percentage" & index) = 0.0
                                                            index += 1
                                                        End While
                                                        'tempRow.Item(2) = "1/1/1970"
                                                        dt.Rows.InsertAt(tempRow, rowIndex)
                                                        Exit While
                                                    End If
                                                    rowIndex -= 1
                                                End While
                                                intValue += row.Item("RATING_STEP_VALUE")
                                                intCount += 1
                                            End While
                                        End If
                                    Next
                                Else
                                    'generate a table of the rating items from the returned record
                                    For Each row In Me.dsCommon.Tables("RatingData").Rows
                                        Dim intCount As Integer = 0
                                        Dim intValue As Integer = row.Item("RATING_INITIAL_VALUE")
                                        While (intCount < row.Item("RATING_COUNT"))
                                            Dim tempRow As DataRow = dt.NewRow
                                            tempRow.Item(0) = intValue
                                            tempRow.Item(1) = 0
                                            tempRow.Item(2) = 0
                                            'tempRow.Item(2) = "1/1/1970"
                                            dt.Rows.Add(tempRow)
                                            intValue += row.Item("RATING_STEP_VALUE")
                                            intCount += 1
                                        End While
                                    Next
                                End If

                                'fill in the rating counts
                                If (tempFilterItemArray(6) = "MULTI") Then
                                    'loop thru chart question and get the values and counts for each
                                    Dim currentQuestion As Integer = tempFilterItemArray(0)
                                    Dim currentIndex As Integer = 1
                                    For Each dr In Me.dsCommon.Tables("DataDump").Rows
                                        If (dr.Item("QUESTION_ID") <> currentQuestion _
                                            And dr.Item("QUESTION_GROUP_ID") = tempFilterItemArray(5)) Then
                                            currentQuestion = dr.Item("QUESTION_ID")
                                            currentIndex += 1
                                        End If

                                        For Each row In dt.Rows
                                            If (row.Item("Value") = dr.Item("ANSWER_VALUE")) Then
                                                row.Item("People" & currentIndex) += dr.Item("ValueCount")
                                                sumAnswered(currentIndex - 1) += dr.Item("ValueCount")
                                                'row.Item("Date") = dr.Item("datDate")
                                                Exit For
                                            End If
                                        Next
                                    Next
                                Else
                                    'loop thru chart question and get the values and counts for each
                                    For Each dr In Me.dsCommon.Tables("DataDump").Rows
                                        For Each row In dt.Rows
                                            If (row.Item("Value") = dr.Item("ANSWER_VALUE") _
                                                And dr.Item("QUESTION_ID") = tempFilterItemArray(0)) Then
                                                row.Item("People") += dr.Item("ValueCount")
                                                sumAnswered(0) += dr.Item("ValueCount")
                                                'row.Item("Date") = dr.Item("datDate")
                                                Exit For
                                            End If
                                        Next
                                    Next
                                End If

                                'calculate the percentages
                                If (tempFilterItemArray(6) = "MULTI") Then
                                    'determine if the user wants and can show the data as a percentage
                                    index = 1
                                    While (index <= multiCount)
                                        For Each dr In dt.Rows
                                            If (sumAnswered(index - 1) = 0) Then
                                                dr.Item("Percentage" & index) = 0
                                            Else
                                                dr.Item("Percentage" & index) = (dr.Item("People" & index) / sumAnswered(index - 1)) * 100
                                            End If
                                        Next
                                        index += 1
                                    End While
                                Else
                                    'determine if the user wants and can show the data as a percentage
                                    For Each dr In dt.Rows
                                        If (sumAnswered(0) = 0) Then
                                            dr.Item("Percentage") = 0
                                        Else
                                            dr.Item("Percentage") = (dr.Item("People") / sumAnswered(0)) * 100
                                        End If
                                    Next
                                End If

                                'calculate the averages
                                If (tempFilterItemArray(6) = "MULTI") Then
                                    index = 1
                                    Dim currentQuestion As Integer = tempFilterItemArray(0)
                                    While (index <= multiCount)
                                        For Each dr In Me.dsCommon.Tables("DataDump").Rows
                                            If (currentQuestion <> dr.Item("QUESTION_ID")) Then
                                                currentQuestion = dr.Item("QUESTION_ID")

                                                'calculate average
                                                If (sumAnswered(index - 1) = 0) Then
                                                    average(index - 1) = 0
                                                Else
                                                    average(index - 1) = sumAverage(index - 1) / sumAnswered(index - 1)
                                                End If

                                                index += 1
                                            End If

                                            If (dr.Item("ANSWER_VALUE") <> Me.dsCommon.Tables("RatingData").Rows(0).Item("MaxValue")) Then
                                                sumAverage(index - 1) += dr.Item("ANSWER_VALUE") * dr.Item("ValueCount")
                                            Else
                                                If (Not CType(Me.dsCommon.Tables("RatingData").Rows(0).Item("NOT_APPLICABLE_IND"), Boolean)) Then
                                                    sumAverage(index - 1) += dr.Item("ANSWER_VALUE") * dr.Item("ValueCount")
                                                End If
                                            End If
                                        Next
                                        'calculate average
                                        If (sumAnswered(index - 1) = 0) Then
                                            average(index - 1) = 0
                                        Else
                                            average(index - 1) = sumAverage(index - 1) / sumAnswered(index - 1)
                                        End If

                                        index = multiCount + 1
                                    End While
                                Else
                                    For Each dr In Me.dsCommon.Tables("DataDump").Rows
                                        If (dr.Item("QUESTION_ID") = tempFilterItemArray(0)) Then
                                            Trace.Warn(dr.Item("ANSWER_VALUE"))
                                            Trace.Warn(Me.dsCommon.Tables("RatingData").Rows(0).Item("MaxValue"))
                                            Trace.Warn(CType(Me.dsCommon.Tables("RatingData").Rows(0).Item("NOT_APPLICABLE_IND"), Boolean))
                                            If (dr.Item("ANSWER_VALUE") <> Me.dsCommon.Tables("RatingData").Rows(0).Item("MaxValue")) Then
                                                sumAverage(0) += dr.Item("ANSWER_VALUE") * dr.Item("ValueCount")
                                            Else
                                                If (Not CType(Me.dsCommon.Tables("RatingData").Rows(0).Item("NOT_APPLICABLE_IND"), Boolean)) Then
                                                    sumAverage(0) += dr.Item("ANSWER_VALUE") * dr.Item("ValueCount")
                                                End If
                                            End If
                                        ElseIf (dr.Item("QUESTION_ID") > tempFilterItemArray(0)) Then
                                            Exit For
                                        End If
                                    Next

                                    'calculate average
                                    If (sumAnswered(0) = 0) Then
                                        average(0) = 0
                                    Else
                                        average(0) = sumAverage(0) / sumAnswered(0)
                                    End If
                                End If

                                'add buffers only after the first table
                                If Not (isFirstTable) Then
                                    'buffer zone
                                    'genericRow = New System.Web.UI.WebControls.TableRow
                                    'genericCell = New System.Web.UI.WebControls.TableCell
                                    ' genericCell.Wrap = False
                                    'genericCell.Height = Unit.Pixel(40)
                                    'genericCell.EnableViewState = True
                                    'If (Session("SplitOutput") <> "NONE") Then
                                    'genericCell.ColumnSpan = 3
                                    'Else
                                    'genericCell.ColumnSpan = 2
                                    'End If
                                    'genericRow.Cells.Add(genericCell)
                                    'genericRow.EnableViewState = True
                                    'Me.DataTable.Rows.Add(genericRow)
                                Else
                                    isFirstTable = False
                                End If

                                If (Not blnDoOnce2 And Session("SplitOutput") <> "NONE") Then
                                    If (tempFilterItemArray(6) = "MULTI") Then
                                        'get the rest of the rows if missing some
                                        index = 0
                                        While (index < sumAnswered.Length)
                                            'generate a row for the question text
                                            genericRow = New System.Web.UI.WebControls.TableRow
                                            genericCell = New System.Web.UI.WebControls.TableCell
                                            genericLabel = New System.web.ui.webcontrols.Label
                                            genericLabel.CssClass = "boldContentRed"
                                            genericLabel.Text = index + 1 & ". "
                                            genericCell.Controls.Add(genericLabel)
                                            genericLabel = New System.web.ui.webcontrols.Label
                                            genericLabel.CssClass = "reportContent"
                                            genericLabel.Text = tempSkipListText(index).replace("~", ",")
                                            genericCell.Controls.Add(genericLabel)
                                            genericLabel = New System.web.ui.webcontrols.Label
                                            genericLabel.CssClass = "boldContentTeal"
                                            genericLabel.Text = "(" & sumAnswered(index) & ")"
                                            genericCell.Controls.Add(genericLabel)
                                            genericCell.Wrap = True
                                            genericCell.EnableViewState = True
                                            genericCell.ColumnSpan = 2
                                            genericRow.Cells.Add(genericCell)

                                            'generate a row for the question answer
                                            genericCell = New System.Web.UI.WebControls.TableCell
                                            genericCell.CssClass = "ReportValueBlue"
                                            genericCell.Text = "Average Answer: " & average(index).ToString("#0.#")
                                            genericCell.Wrap = False
                                            genericCell.HorizontalAlign = HorizontalAlign.Right
                                            genericCell.EnableViewState = True
                                            genericRow.Cells.Add(genericCell)

                                            genericRow.EnableViewState = True
                                            If (Session("SplitOutput") <> "NONE") Then
                                                genericRow.BackColor = Color.LemonChiffon
                                            End If
                                            Me.DataTable.Rows.Add(genericRow)
                                            blnDoOnce2 = True

                                            index += 1
                                        End While
                                    Else
                                        'generate a row for the question text
                                        genericRow = New System.Web.UI.WebControls.TableRow
                                        genericCell = New System.Web.UI.WebControls.TableCell
                                        genericLabel = New System.web.ui.webcontrols.Label
                                        genericLabel.CssClass = "reportContent"
                                        genericLabel.Text = CType(tempFilterItemArray(3), String).Replace("~", ",")
                                        genericCell.Controls.Add(genericLabel)
                                        genericLabel = New System.web.ui.webcontrols.Label
                                        genericLabel.CssClass = "boldContentTeal"
                                        genericLabel.Text = "(" & sumAnswered(0) & ")"
                                        genericCell.Controls.Add(genericLabel)
                                        genericCell.Wrap = True
                                        genericCell.EnableViewState = True
                                        genericCell.ColumnSpan = 2
                                        genericRow.Cells.Add(genericCell)

                                        'generate a row for the question answer
                                        genericCell = New System.Web.UI.WebControls.TableCell
                                        genericCell.CssClass = "ReportValueBlue"
                                        genericCell.Text = "Average Answer: " & average(0).ToString("#0.#")
                                        genericCell.Wrap = False
                                        genericCell.HorizontalAlign = HorizontalAlign.Right
                                        genericCell.EnableViewState = True
                                        genericRow.Cells.Add(genericCell)

                                        genericRow.EnableViewState = True
                                        If (Session("SplitOutput") <> "NONE") Then
                                            genericRow.BackColor = Color.LemonChiffon
                                        End If
                                        Me.DataTable.Rows.Add(genericRow)
                                        blnDoOnce2 = True
                                    End If
                                ElseIf (Session("SplitOutput") = "NONE") Then
                                    If (tempFilterItemArray(6) = "MULTI") Then
                                        Dim questionText As String = ""
                                        index = 0
                                        Dim currentQuestion As Integer = 0
                                        For Each row In Me.dsCommon.Tables("DataDump").Rows
                                            If (row.Item("QUESTION_GROUP_ID") = tempFilterItemArray(5) _
                                                And tempSkipList.Contains(CType(row.Item("QUESTION_ID"), String))) Then
                                                If (questionText = "") Then
                                                    questionText = row.Item("QuestionText")
                                                    currentQuestion = row.Item("QUESTION_ID")

                                                    'generate a row for the question text
                                                    genericRow = New System.Web.UI.WebControls.TableRow
                                                    genericCell = New System.Web.UI.WebControls.TableCell
                                                    genericLabel = New System.Web.UI.WebControls.Label
                                                    genericLabel.CssClass = "boldContentRed"
                                                    genericLabel.Text = index + 1 & ". "
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericLabel = New System.Web.UI.WebControls.Label
                                                    genericLabel.CssClass = "reportContent"
                                                    genericLabel.Text = questionText.Replace("~", ",")
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.Wrap = True
                                                    genericCell.EnableViewState = True
                                                    genericRow.Cells.Add(genericCell)
                                                    genericLabel = New System.Web.UI.WebControls.Label
                                                    genericLabel.CssClass = "boldContentTeal"
                                                    genericLabel.Text = "(" & sumAnswered(index) & ")"
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.Wrap = True
                                                    genericCell.EnableViewState = True
                                                    genericRow.Cells.Add(genericCell)

                                                    'generate a row for the question answer
                                                    genericCell = New System.Web.UI.WebControls.TableCell
                                                    genericCell.CssClass = "ReportValueBlue"
                                                    genericCell.Text = "Average Answer: " & average(index).ToString("#0.#")
                                                    genericCell.Wrap = False
                                                    genericCell.HorizontalAlign = HorizontalAlign.Right
                                                    genericCell.EnableViewState = True
                                                    genericRow.Cells.Add(genericCell)
                                                    genericRow.EnableViewState = True
                                                    Me.DataTable.Rows.Add(genericRow)
                                                ElseIf (currentQuestion <> row.Item("QUESTION_ID")) Then
                                                    questionText = row.Item("QuestionText")
                                                    currentQuestion = row.Item("QUESTION_ID")
                                                    index += 1

                                                    'generate a row for the question text
                                                    genericRow = New System.Web.UI.WebControls.TableRow
                                                    genericCell = New System.Web.UI.WebControls.TableCell
                                                    genericLabel = New System.Web.UI.WebControls.Label
                                                    genericLabel.CssClass = "boldContentRed"
                                                    genericLabel.Text = index + 1 & ". "
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericLabel = New System.Web.UI.WebControls.Label
                                                    genericLabel.CssClass = "reportContent"
                                                    genericLabel.Text = questionText.Replace("~", ",")
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.Wrap = True
                                                    genericCell.EnableViewState = True
                                                    genericRow.Cells.Add(genericCell)
                                                    genericLabel = New System.Web.UI.WebControls.Label
                                                    genericLabel.CssClass = "boldContentTeal"
                                                    genericLabel.Text = "(" & sumAnswered(index) & ")"
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.Wrap = True
                                                    genericCell.EnableViewState = True
                                                    genericRow.Cells.Add(genericCell)

                                                    'generate a row for the question answer
                                                    genericCell = New System.Web.UI.WebControls.TableCell
                                                    genericCell.CssClass = "ReportValueBlue"
                                                    genericCell.Text = "Average Answer: " & average(index).ToString("#0.#")
                                                    genericCell.Wrap = False
                                                    genericCell.HorizontalAlign = HorizontalAlign.Right
                                                    genericCell.EnableViewState = True
                                                    genericRow.Cells.Add(genericCell)
                                                    genericRow.EnableViewState = True
                                                    Me.DataTable.Rows.Add(genericRow)
                                                End If
                                            End If
                                        Next
                                    Else
                                        'generate a row for the question text
                                        genericRow = New System.Web.UI.WebControls.TableRow
                                        genericCell = New System.Web.UI.WebControls.TableCell
                                        genericLabel = New System.web.ui.webcontrols.Label
                                        genericLabel.CssClass = "reportContent"
                                        genericLabel.Text = CType(tempFilterItemArray(3), String).Replace("~", ",")
                                        genericCell.Controls.Add(genericLabel)
                                        genericCell.Wrap = True
                                        genericCell.EnableViewState = True
                                        genericRow.Cells.Add(genericCell)
                                        genericLabel = New System.web.ui.webcontrols.Label
                                        genericLabel.CssClass = "boldContentTeal"
                                        genericLabel.Text = "(" & sumAnswered(0) & ")"
                                        genericCell.Controls.Add(genericLabel)
                                        genericCell.Wrap = True
                                        genericCell.EnableViewState = True
                                        genericRow.Cells.Add(genericCell)

                                        'generate a row for the question answer
                                        genericCell = New System.Web.UI.WebControls.TableCell
                                        genericCell.CssClass = "ReportValueBlue"
                                        genericCell.Text = "Average Answer: " & average(0).ToString("#0.#")
                                        genericCell.Wrap = False
                                        genericCell.EnableViewState = True
                                        genericCell.HorizontalAlign = HorizontalAlign.Right
                                        genericRow.Cells.Add(genericCell)
                                        genericRow.EnableViewState = True
                                        Me.DataTable.Rows.Add(genericRow)
                                    End If
                                End If

                                'generate a row for the question answer
                                genericTable = New System.Web.UI.WebControls.Table
                                Dim colorIndex As Integer = 0
                                blnDoOnce3 = False
                                For Each row In dt.Rows
                                    'get the color for the current row
                                    Dim strCurrentColor As String = "FFFFFF"
                                    If (Me.dsCommon.Tables("Colors").Rows.Count() > 0) Then
                                        strCurrentColor = Me.dsCommon.Tables("Colors").Rows(colorIndex).Item("CHART_COLOR_CODE")
                                        colorIndex += 1
                                        If (colorIndex = Me.dsCommon.Tables("Colors").Rows.Count()) Then
                                            colorIndex = 0
                                        End If
                                    End If

                                    'set up the color for the row
                                    Dim Alpha As Integer = 0
                                    Dim Red As Integer = 0
                                    Dim Green As Integer = 0
                                    Dim Blue As Integer = 0
                                    If (strCurrentColor.Length = 8) Then
                                        Alpha = Integer.Parse(strCurrentColor.Substring(0, 2), Globalization.NumberStyles.HexNumber)
                                        Red = Integer.Parse(strCurrentColor.Substring(2, 2), Globalization.NumberStyles.HexNumber)
                                        Green = Integer.Parse(strCurrentColor.Substring(4, 2), Globalization.NumberStyles.HexNumber)
                                        Blue = Integer.Parse(strCurrentColor.Substring(6, 2), Globalization.NumberStyles.HexNumber)
                                    Else
                                        Red = Integer.Parse(strCurrentColor.Substring(0, 2), Globalization.NumberStyles.HexNumber)
                                        Green = Integer.Parse(strCurrentColor.Substring(2, 2), Globalization.NumberStyles.HexNumber)
                                        Blue = Integer.Parse(strCurrentColor.Substring(4, 2), Globalization.NumberStyles.HexNumber)
                                    End If

                                    'table header setup - do once
                                    If Not (blnDoOnce3) Then
                                        genericRow = New System.Web.UI.WebControls.TableRow
                                        genericCell = New System.Web.UI.WebControls.TableCell
                                        genericCell.Width = Unit.Pixel(20)
                                        genericCell.Wrap = False
                                        genericCell.EnableViewState = True
                                        genericRow.Cells.Add(genericCell)
                                        genericCell = New System.Web.UI.WebControls.TableCell
                                        genericLabel = New System.web.ui.webcontrols.Label
                                        genericLabel.CssClass = "reportContent"
                                        genericLabel.Text = "Value"
                                        genericCell.Controls.Add(genericLabel)
                                        genericCell.Wrap = True
                                        genericCell.EnableViewState = True
                                        genericCell.BackColor = System.Drawing.Color.AntiqueWhite
                                        genericRow.Cells.Add(genericCell)

                                        If (tempFilterItemArray(6) = "MULTI") Then
                                            index = 1
                                            While (index <= multiCount)
                                                genericCell = New System.Web.UI.WebControls.TableCell
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "reportContent"
                                                genericLabel.Text = index
                                                genericCell.Controls.Add(genericLabel)
                                                genericCell.HorizontalAlign = HorizontalAlign.Center
                                                genericCell.Wrap = False
                                                genericCell.EnableViewState = True
                                                genericCell.BackColor = System.Drawing.Color.AntiqueWhite
                                                genericRow.Cells.Add(genericCell)
                                                index += 1
                                            End While
                                        Else
                                            genericCell = New System.Web.UI.WebControls.TableCell
                                            genericLabel = New System.web.ui.webcontrols.Label
                                            genericLabel.CssClass = "reportContent"

                                            'change the header based on percent or count display
                                            If (blnPercentDisplay) Then
                                                genericLabel.Text = "Percent"
                                            Else
                                                genericLabel.Text = "Count"
                                            End If

                                            genericCell.Controls.Add(genericLabel)
                                            genericCell.HorizontalAlign = HorizontalAlign.Center
                                            genericCell.Wrap = False
                                            genericCell.EnableViewState = True
                                            genericCell.BackColor = System.Drawing.Color.AntiqueWhite
                                            genericRow.Cells.Add(genericCell)
                                        End If

                                        genericRow.EnableViewState = True
                                        genericTable.Rows.Add(genericRow)
                                        blnDoOnce3 = True
                                    End If

                                    'populate row data
                                    genericRow = New System.Web.UI.WebControls.TableRow
                                    genericCell = New System.Web.UI.WebControls.TableCell
                                    genericCell.Width = Unit.Pixel(20)
                                    genericCell.Wrap = False
                                    genericCell.EnableViewState = True
                                    genericRow.Cells.Add(genericCell)
                                    genericCell = New System.Web.UI.WebControls.TableCell
                                    genericLabel = New System.web.ui.webcontrols.Label
                                    genericLabel.CssClass = "reportContent"
                                    If (row.Item("Value") = Me.dsCommon.Tables("RatingData").Rows(0).Item("maxValue") And _
                                        CType(Me.dsCommon.Tables("RatingData").Rows(0).Item("NOT_APPLICABLE_IND"), Boolean)) Then
                                        genericLabel.Text = "N/A"
                                    Else
                                        genericLabel.Text = row.Item("Value")
                                    End If
                                    genericCell.Controls.Add(genericLabel)
                                    genericCell.Wrap = True
                                    genericCell.HorizontalAlign = HorizontalAlign.Center
                                    genericCell.EnableViewState = True
                                    genericCell.BackColor = System.Drawing.Color.FromArgb(Alpha, Red, Green, Blue)
                                    genericRow.Cells.Add(genericCell)

                                    If (tempFilterItemArray(6) = "MULTI") Then
                                        index = 1
                                        While (index <= multiCount)
                                            genericCell = New System.Web.UI.WebControls.TableCell
                                            genericLabel = New System.web.ui.webcontrols.Label
                                            genericLabel.CssClass = "reportContent"

                                            'format the value based on percent or count display
                                            If (blnPercentDisplay) Then
                                                genericLabel.Text = CType(row.Item("Percentage" & index), Double).ToString("#0.##")
                                            Else
                                                genericLabel.Text = row.Item("People" & index)
                                            End If

                                            genericCell.Controls.Add(genericLabel)
                                            genericCell.HorizontalAlign = HorizontalAlign.Right
                                            genericCell.Wrap = False
                                            genericCell.EnableViewState = True
                                            genericCell.BackColor = System.Drawing.Color.FromArgb(Alpha, Red, Green, Blue)
                                            genericRow.Cells.Add(genericCell)
                                            index += 1
                                        End While
                                    Else
                                        genericCell = New System.Web.UI.WebControls.TableCell
                                        genericLabel = New System.web.ui.webcontrols.Label
                                        genericLabel.CssClass = "reportContent"

                                        'format the value based on percent or count display
                                        If (blnPercentDisplay) Then
                                            genericLabel.Text = CType(row.Item("Percentage"), Double).ToString("#0.##")
                                        Else
                                            genericLabel.Text = row.Item("People")
                                        End If

                                        genericCell.Controls.Add(genericLabel)
                                        genericCell.HorizontalAlign = HorizontalAlign.Right
                                        genericCell.Wrap = False
                                        genericCell.EnableViewState = True
                                        genericCell.BackColor = System.Drawing.Color.FromArgb(Alpha, Red, Green, Blue)
                                        genericRow.Cells.Add(genericCell)
                                    End If

                                    genericRow.EnableViewState = True
                                    genericTable.Rows.Add(genericRow)
                                Next

                                genericRow = New System.Web.UI.WebControls.TableRow
                                genericCell = New System.Web.UI.WebControls.TableCell
                                genericCell.Controls.Add(genericTable)
                                genericCell.Wrap = False
                                genericCell.EnableViewState = True
                                If (Session("SplitOutput") <> "NONE") Then
                                    genericCell.ColumnSpan = 2
                                End If
                                genericCell.VerticalAlign = VerticalAlign.Top
                                genericRow.Cells.Add(genericCell)

                                'generate the chart for the Rating question
                                genericCell = New System.Web.UI.WebControls.TableCell
                                genericChart = New Infragistics.webui.UltraWebChart.UltraChart
                                If Not (generateChart(genericChart, tempFilterItemArray(0), tempFilterItemArray(6), dateFilter(2), dt, tempSkipList, runCount, False, True, False, blnPercentDisplay)) Then
                                    Session("Alert") = "Chart generation failure."
                                    Return False
                                End If
                                genericCell.Controls.Add(genericChart)
                                genericCell.Wrap = False
                                genericCell.EnableViewState = True
                                genericCell.VerticalAlign = VerticalAlign.Top
                                genericCell.HorizontalAlign = HorizontalAlign.Right
                                genericRow.Cells.Add(genericCell)
                                genericRow.EnableViewState = True

                                'Set to put a line break after 3 question results appear on the page.
                                If (questionCounter Mod Session("QuestionsPerPage") = 0) Then
                                    genericRow.CssClass = "Break"
                                End If
                                questionCounter += 1

                                Me.DataTable.Rows.Add(genericRow)
                                Me.DataTable.EnableViewState = True
                            End If
                        End If
                    Next

                    'remove ratings section data header if there is no data
                    If (CType(Session("RItemExists"), Boolean) <> True) Then
                        Me.DataTable.Rows.RemoveAt(Me.DataTable.Rows.Count - 1)
                    End If
                Else
                    Session("Alert") = "Failed to get survey data."
                    Return False
                End If
                runCount += 1
            Next

            'remove the section header if no rows exist
            If (Session("RItemExists") = False) Then
                While (Me.DataTable.Rows.Count() > 1)
                    Me.DataTable.Rows.RemoveAt(Me.DataTable.Rows.Count() - 1)
                End While
            End If
            Return True
        Catch ex As Exception
            Session("Alert") = "Failed to get survey data.  " & ex.ToString
            Return False
        End Try
    End Function

    'processes the choice data
    Private Function processChoiceData() As Boolean
        Trace.Warn("processing choice section")
        Try
            'define the variables
            Dim genericCell As System.Web.UI.WebControls.TableCell
            Dim genericRow As System.Web.UI.WebControls.TableRow
            Dim genericTable As System.Web.UI.WebControls.Table
            Dim genericLabel As System.Web.UI.WebControls.Label
            Dim genericChart As Infragistics.WebUI.UltraWebChart.UltraChart
            Dim noDataRows As ArrayList = New ArrayList

            'Set up the filter
            Dim filterArray As Array
            Dim skipList As ArrayList = New ArrayList
            Dim skipListText As ArrayList = New ArrayList
            filterArray = CType(Session("DataItems"), String).Split(",")

            'get filter item data
            Dim filterItem As String = ""
            Session("NonRTypeSelected") = False
            Dim blnDoOnce As Boolean = False
            Dim blnDoOnce3 As Boolean = False
            Dim filterItemCount As Integer = 0
            Dim questionCounter As Integer = 1
            For Each filterItem In filterArray
                Dim tempFilterItemArray As Array = filterItem.Split("^")
                'process type questions (except R type) with averages and the question label
                If (CType(tempFilterItemArray(1), Boolean) And tempFilterItemArray(2) <> "R") Then
                    Dim rowCount As Integer = 0
                    Dim dr As DataRow
                    If (tempFilterItemArray(2) = "C") Then
                        filterItemCount += 1
                        Dim blnPercentDisplay As Boolean = False
                        'skip any questions in the skip list
                        If Not (skipList.Contains(CType(tempFilterItemArray(0), String))) Then
                            'add a separator between the questions and the rest of the data
                            If (blnDoOnce = False) Then
                                blnDoOnce = True
                                If (Session("RItemExists") = True) Then
                                    'generate a row for the question text
                                    genericRow = New System.Web.UI.WebControls.TableRow
                                    genericRow.CssClass = "BreakBefore"
                                    genericCell = New System.Web.UI.WebControls.TableCell
                                    genericLabel = New System.web.ui.webcontrols.Label
                                    genericLabel.Text = "<hr style='BORDER-TOP-STYLE: ridge' width='100%' SIZE='3'/>"
                                    genericCell.Controls.Add(genericLabel)
                                    genericCell.Wrap = False
                                    genericCell.EnableViewState = True
                                    genericCell.ColumnSpan = 2
                                    genericRow.Cells.Add(genericCell)
                                    genericRow.EnableViewState = True
                                    Me.DataTable2.Rows.Add(genericRow)
                                End If
                            End If

                            'get the list of questions in the multi-group if the question chart type is multi
                            Dim tempSkipList As ArrayList = New ArrayList
                            Dim tempSkipListText As ArrayList = New ArrayList
                            If (tempFilterItemArray(6) = "MULTI") Then
                                Dim tempItem As String = ""
                                For Each tempItem In filterArray
                                    Dim tempSplit As Array = tempItem.Split("^")
                                    If (CType(tempSplit(0), Integer) > CType(tempFilterItemArray(0), Integer) And CType(tempSplit(1), Boolean) _
                                        And tempSplit(6) = "MULTI") Then
                                        If (tempSplit(5) = tempFilterItemArray(5)) Then
                                            tempSkipList.Add(tempSplit(0))
                                            tempSkipListText.Add(tempSplit(3))
                                            If (tempSplit(4) = "P-True") Then
                                                blnPercentDisplay = True
                                            End If
                                        End If
                                    ElseIf (CType(tempSplit(0), Integer) = CType(tempFilterItemArray(0), Integer)) Then
                                        tempSkipListText.Add(tempSplit(3))
                                        If (tempSplit(4) = "P-True") Then
                                            blnPercentDisplay = True
                                        End If
                                    End If
                                Next
                            Else
                                If (tempFilterItemArray(4) = "P-True") Then
                                    blnPercentDisplay = True
                                End If
                            End If

                            'build skiplist string for queries
                            Dim strSkipList As String = tempFilterItemArray(0)
                            Dim index As Integer = 0
                            While (index < tempSkipList.Count)
                                strSkipList &= ", " & tempSkipList(index)
                                index += 1
                            End While
                            tempSkipList.Add(tempFilterItemArray(0))

                            'add skip items to master list
                            index = 0
                            skipListText.Clear()
                            skipListText.Add(tempFilterItemArray(3))
                            While (index < tempSkipList.Count - 1)
                                skipList.Add(tempSkipList(index))
                                skipListText.Add(tempSkipListText(index))
                                index += 1
                            End While
                            skipList.Add(tempFilterItemArray(0))

                            'get the choices from the fempList table or fempResults table (for D Type questions)
                            If (tempFilterItemArray(6) = "MULTI") Then
                                Me.sqlCommonCommand.CommandText = "select distinct fl.QUESTION_ID, fl.LIST_ITEM_ID, fl.LIST_ITEM_VALUE, fl.LIST_ITEM_TITLE " & _
                                "from " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM fl, " & myUtility.getExtension() & "SAT_TEMPLATE_SURVEY fs, " & _
                                myUtility.getExtension() & "SAT_TEMPLATE_QUESTION fq where fs.SURVEY_KEY = " & Session("seqSurvey") & _
                                " And fs.TEMPLATE_KEY = fl.TEMPLATE_KEY and fl.QUESTION_ID in (" & strSkipList & _
                                ") and fl.TEMPLATE_KEY = fq.TEMPLATE_KEY and fl.QUESTION_ID = fq.QUESTION_ID " & _
                                "and fq.QUESTION_GROUP_ID = " & tempFilterItemArray(5) & _
                                " order by fl.LIST_ITEM_ID, fl.LIST_ITEM_VALUE, fl.QUESTION_ID"
                                Trace.Warn(Me.sqlCommonCommand.CommandText)
                            Else
                                Me.sqlCommonCommand.CommandText = "select fl.LIST_ITEM_ID, fl.LIST_ITEM_VALUE, fl.LIST_ITEM_TITLE " & _
                                "from " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM fl, " & myUtility.getExtension() & _
                                "SAT_TEMPLATE_SURVEY fs where fs.SURVEY_KEY = " & Session("seqSurvey") & _
                                " And fs.TEMPLATE_KEY = fl.TEMPLATE_KEY and fl.QUESTION_ID = " & tempFilterItemArray(0)
                                Trace.Warn(Me.sqlCommonCommand.CommandText)
                            End If
                            If (Me.dsCommon.Tables.Contains("ListItems")) Then
                                Me.dsCommon.Tables("ListItems").Rows.Clear()
                            End If
                            Me.sqlCommonAdapter.Fill(Me.dsCommon, "ListItems")

                            'set the question filter again
                            If (tempFilterItemArray(6) = "MULTI") Then
                                setQuestionFilter(tempSkipList)
                            Else
                                tempSkipList = New ArrayList
                                setQuestionFilter(tempSkipList)
                            End If

                            Dim dateFilter As ArrayList
                            Dim runCount As Integer = 0
                            Dim blnDoOnce2 As Boolean = False
                            Dim row As DataRow
                            For Each dateFilter In Me.dateFilterArr
                                If (setDataFilter(dateFilter(2))) Then
                                    Me.sqlCommonCommand.CommandText = "select frs.QUESTION_ID, frs.ANSWER_VALUE, count(frs.ANSWER_VALUE) as " & _
                                        "ValueCount, frs.ANSWER_TEXT, fq.CHART_TYPE_CODE, fq.QUESTION_GROUP_ID, fq.QUESTION_TEXT as QuestionText " & _
                                        "from " & myUtility.getExtension() & "SAT_RESPONSE_RESULT frs, " & myUtility.getExtension() & _
                                        "SAT_TEMPLATE_QUESTION fq where frs.SURVEY_KEY = " & Session("seqSurvey") & " " & _
                                        dateFilter(2).Replace("LAST_USED_DATE", "RESPONSE_DATE") & " " & Me.strDataFilter & " " & _
                                        Me.strQuestionFilter & " and frs.QUESTION_ID = fq.QUESTION_ID and frs.QUESTION_TYPE = fq.QUESTION_TYPE_CODE " & _
                                        "and fq.TEMPLATE_KEY = " & Session("seqTemplate") & " group by frs.QUESTION_ID, frs.ANSWER_VALUE, " & _
                                        "frs.ANSWER_TEXT, fq.CHART_TYPE_CODE, fq.QUESTION_GROUP_ID, fq.QUESTION_TEXT order by frs.QUESTION_ID, frs.ANSWER_VALUE, " & _
                                        "frs.ANSWER_TEXT, fq.CHART_TYPE_CODE, fq.QUESTION_GROUP_ID, fq.QUESTION_TEXT"
                                    Trace.Warn(Me.sqlCommonCommand.CommandText)
                                    If (Me.dsCommon.Tables.Contains("DataDump")) Then
                                        Me.dsCommon.Tables("DataDump").Rows.Clear()
                                    End If
                                    Me.sqlCommonAdapter.Fill(Me.dsCommon, "DataDump")

                                    'get a count of questions in the group if dealing with a multi-chart set
                                    Dim multiCount As Integer = 1
                                    If (tempFilterItemArray(6) = "MULTI") Then
                                        multiCount = tempSkipList.Count()
                                    End If

                                    Dim sumAnswered(multiCount - 1) As Double
                                    Dim sumAnswerCount(multiCount - 1) As Double

                                    'build a table of percentages for the question using the list items
                                    If (Me.dsCommon.Tables("ListItems").Rows.Count > 0) Then
                                        'build the skeletal structure
                                        Dim dt As DataTable = New DataTable
                                        dt.Columns.Add("ListLabel")
                                        dt.Columns("ListLabel").DataType = Me.dsCommon.Tables("ListItems").Columns("LIST_ITEM_TITLE").DataType
                                        dt.Columns.Add("ListValue")
                                        dt.Columns("ListValue").DataType = Me.dsCommon.Tables("ListItems").Columns("LIST_ITEM_VALUE").DataType

                                        If (tempFilterItemArray(6) = "MULTI") Then
                                            index = 1
                                            While (index <= multiCount)
                                                'generated for the count of questions in the group
                                                dt.Columns.Add("Percentage" & index)
                                                dt.Columns("Percentage" & index).DataType = System.Type.GetType("System.Decimal", True, True)
                                                dt.Columns.Add("AnswerCount" & index)
                                                dt.Columns("AnswerCount" & index).DataType = Me.dsCommon.Tables("ListItems").Columns("LIST_ITEM_ID").DataType
                                                index += 1
                                            End While
                                        Else
                                            'generated for the count of questions in the group
                                            dt.Columns.Add("Percentage")
                                            dt.Columns("Percentage").DataType = System.Type.GetType("System.Decimal", True, True)
                                            dt.Columns.Add("AnswerCount")
                                            dt.Columns("AnswerCount").DataType = Me.dsCommon.Tables("ListItems").Columns("LIST_ITEM_ID").DataType
                                        End If

                                        'add default values and the item value as well as the label to the table
                                        If (tempFilterItemArray(6) = "MULTI") Then
                                            dt.TableName = "Question" & tempFilterItemArray(0)
                                            For Each row In Me.dsCommon.Tables("ListItems").Rows
                                                Dim tempRow As DataRow = dt.NewRow
                                                tempRow.Item(0) = CType(row.Item("strLabel"), String).Replace("~", ",")
                                                tempRow.Item(1) = row.Item("intListValue")

                                                index = 1
                                                While (index <= multiCount)
                                                    tempRow.Item("Percentage" & index) = 0.0
                                                    tempRow.Item("AnswerCount" & index) = 0
                                                    index += 1
                                                End While
                                                dt.Rows.Add(tempRow)
                                            Next
                                        Else
                                            dt.TableName = "Question" & tempFilterItemArray(0)
                                            For Each row In Me.dsCommon.Tables("ListItems").Rows
                                                Dim tempRow As DataRow = dt.NewRow
                                                tempRow.Item(0) = CType(row.Item("LIST_ITEM_TITLE"), String).Replace("~", ",")
                                                tempRow.Item(1) = row.Item("LIST_ITEM_VALUE")
                                                tempRow.Item(2) = 0.0
                                                tempRow.Item(3) = 0
                                                dt.Rows.Add(tempRow)
                                            Next
                                        End If

                                        Dim blnDataExists As Boolean = False
                                        If (tempFilterItemArray(6) = "MULTI") Then
                                            'add the answer counts to the table
                                            Dim currentQuestion As Integer = tempFilterItemArray(0)
                                            Dim currentIndex As Integer = 1
                                            For Each dr In Me.dsCommon.Tables("DataDump").Rows
                                                If (dr.Item("QUESTION_ID") <> currentQuestion _
                                                    And dr.Item("QUESTION_GROUP_ID") = tempFilterItemArray(5)) Then
                                                    currentQuestion = dr.Item("QUESTION_ID")
                                                    currentIndex += 1
                                                End If

                                                'split out the result values
                                                Dim resultArray As Array = CType(dr.Item("ANSWER_TEXT"), String).Split(",")
                                                Dim resultItem As String = ""
                                                For Each resultItem In resultArray
                                                    resultItem = resultItem.Trim()

                                                    'get the list item label
                                                    Me.sqlCommonCommand.CommandText = "Select LIST_ITEM_TITLE from " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM " & _
                                                    "where QUESTION_ID = " & currentQuestion & _
                                                    " and TEMPLATE_KEY = " & Session("seqTemplate") & _
                                                    " and LIST_ITEM_VALUE = " & resultItem

                                                    If (Me.dsCommon.Tables.Contains("ListLabel")) Then
                                                        Me.dsCommon.Tables("ListLabel").Rows.Clear()
                                                    End If
                                                    Me.sqlCommonAdapter.Fill(Me.dsCommon, "ListLabel")

                                                    'populate the data table
                                                    If (dr.Item("QUESTION_ID") = currentQuestion) Then
                                                        blnDataExists = True
                                                        Session("NonRTypeSelected") = True
                                                        For Each row In dt.Rows()
                                                            If (Me.dsCommon.Tables("ListLabel").Rows.Count > 0) Then
                                                                If (row.Item("ListValue") = resultItem _
                                                                    And row.Item("ListLabel") = Me.dsCommon.Tables("ListLabel").Rows(0).Item("LIST_ITEM_TITLE")) Then
                                                                    row.Item("AnswerCount" & currentIndex) += dr.Item("ValueCount")
                                                                    sumAnswerCount(currentIndex - 1) += dr.Item("ValueCount")
                                                                    Exit For
                                                                End If
                                                            Else
                                                                If (row.Item("ListValue") = resultItem) Then
                                                                    row.Item("AnswerCount" & currentIndex) += dr.Item("ValueCount")
                                                                    sumAnswerCount(currentIndex - 1) += dr.Item("ValueCount")
                                                                    Exit For
                                                                End If
                                                            End If
                                                        Next
                                                    ElseIf (currentIndex > multiCount) Then
                                                        Exit For
                                                    End If
                                                Next
                                                sumAnswered(currentIndex - 1) += dr.Item("ValueCount")
                                            Next
                                        Else
                                            'add the answer counts to the table
                                            For Each dr In Me.dsCommon.Tables("DataDump").Rows
                                                If (dr.Item("QUESTION_ID") = tempFilterItemArray(0)) Then
                                                    blnDataExists = True
                                                    Session("NonRTypeSelected") = True

                                                    'split out the result values
                                                    Dim resultArray As Array = CType(dr.Item("ANSWER_TEXT"), String).Split(",")
                                                    Dim resultItem As String = ""
                                                    For Each resultItem In resultArray
                                                        resultItem = resultItem.Trim()
                                                        For Each row In dt.Rows()
                                                            If (row.Item("ListValue") = resultItem) Then
                                                                row.Item("AnswerCount") += dr.Item("ValueCount")
                                                                sumAnswerCount(0) += dr.Item("ValueCount")
                                                                Exit For
                                                            End If
                                                        Next
                                                    Next
                                                    sumAnswered(0) += dr.Item("ValueCount")
                                                ElseIf (dr.Item("QUESTION_ID") > tempFilterItemArray(0)) Then
                                                    Exit For
                                                End If
                                            Next
                                        End If

                                        If (blnDataExists) Then
                                            If (tempFilterItemArray(6) = "MULTI") Then
                                                'calculate the percentages for each row in the table
                                                index = 1
                                                While (index <= multiCount)
                                                    For Each row In dt.Rows
                                                        If (sumAnswerCount(index - 1) = 0) Then
                                                            row.Item("Percentage" & index) = 0
                                                        Else
                                                            row.Item("Percentage" & index) = (row.Item("AnswerCount" & index) / sumAnswerCount(index - 1)) * 100
                                                        End If
                                                    Next
                                                    index += 1
                                                End While
                                            Else
                                                'calculate the percentages for each row in the table
                                                For Each row In dt.Rows
                                                    If (sumAnswerCount(0) = 0) Then
                                                        row.Item("Percentage") = 0
                                                    Else
                                                        row.Item("Percentage") = (row.Item("AnswerCount") / sumAnswerCount(0)) * 100
                                                    End If
                                                Next
                                            End If

                                            If (Not blnDoOnce3) Then
                                                'section header
                                                genericRow = New System.Web.UI.WebControls.TableRow
                                                genericCell = New System.Web.UI.WebControls.TableCell
                                                genericCell.Wrap = False
                                                genericCell.CssClass = "sectionHeadTeal"
                                                genericCell.Text = "Choice Results"
                                                genericCell.EnableViewState = True
                                                genericCell.ColumnSpan = 2
                                                genericRow.Cells.Add(genericCell)
                                                genericRow.EnableViewState = True
                                                Me.DataTable2.Rows.Add(genericRow)
                                                blnDoOnce3 = True
                                            End If

                                            'buffer zone
                                            'genericRow = New System.Web.UI.WebControls.TableRow
                                            'genericCell = New System.Web.UI.WebControls.TableCell
                                            'genericCell.Wrap = False
                                            'genericCell.Height = Unit.Pixel(40)
                                            'genericCell.EnableViewState = True
                                            'genericCell.ColumnSpan = 2
                                            'genericRow.Cells.Add(genericCell)
                                            'genericRow.EnableViewState = True
                                            'Me.DataTable2.Rows.Add(genericRow)

                                            If (Session("SplitOutput") <> "NONE") Then
                                                If (tempFilterItemArray(6) = "MULTI") Then
                                                    'get the rest of the rows if missing some
                                                    index = 0
                                                    While (index < sumAnswered.Length)
                                                        'generate a row for the question text
                                                        genericRow = New System.Web.UI.WebControls.TableRow
                                                        genericCell = New System.Web.UI.WebControls.TableCell
                                                        genericLabel = New System.web.ui.webcontrols.Label
                                                        genericLabel.CssClass = "boldContentRed"
                                                        genericLabel.Text = index + 1 & ". "
                                                        genericCell.Controls.Add(genericLabel)
                                                        genericLabel = New System.web.ui.webcontrols.Label
                                                        genericLabel.CssClass = "reportContent"
                                                        genericLabel.Text = tempSkipListText(index).replace("~", ",")
                                                        genericCell.Controls.Add(genericLabel)
                                                        genericLabel = New System.web.ui.webcontrols.Label
                                                        genericLabel.CssClass = "boldContentTeal"
                                                        genericLabel.Text = "(" & sumAnswered(index) & ")"
                                                        genericCell.Controls.Add(genericLabel)
                                                        genericCell.Wrap = True
                                                        genericCell.EnableViewState = True
                                                        genericCell.ColumnSpan = 2
                                                        genericRow.Cells.Add(genericCell)
                                                        genericRow.EnableViewState = True
                                                        If (Session("SplitOutput") <> "NONE") Then
                                                            If (blnDoOnce2) Then
                                                                genericRow.BackColor = Color.LightSteelBlue
                                                            Else
                                                                genericRow.BackColor = Color.LemonChiffon
                                                            End If
                                                        End If
                                                        Me.DataTable2.Rows.Add(genericRow)

                                                        index += 1
                                                    End While
                                                    blnDoOnce2 = True
                                                Else
                                                    'generate a row for the question text
                                                    genericRow = New System.Web.UI.WebControls.TableRow
                                                    genericCell = New System.Web.UI.WebControls.TableCell
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "reportContent"
                                                    genericLabel.Text = CType(tempFilterItemArray(3), String).Replace("~", ",")
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "boldContentTeal"
                                                    genericLabel.Text = "(" & sumAnswered(0) & ")"
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.Wrap = True
                                                    genericCell.EnableViewState = True
                                                    genericCell.ColumnSpan = 2
                                                    genericRow.Cells.Add(genericCell)
                                                    genericRow.EnableViewState = True
                                                    If (Session("SplitOutput") <> "NONE") Then
                                                        If (blnDoOnce2) Then
                                                            genericRow.BackColor = Color.LightSteelBlue
                                                        Else
                                                            genericRow.BackColor = Color.LemonChiffon
                                                        End If
                                                    End If
                                                    Me.DataTable2.Rows.Add(genericRow)
                                                    blnDoOnce2 = True
                                                End If
                                            ElseIf (Session("SplitOutput") = "NONE") Then
                                                If (tempFilterItemArray(6) = "MULTI") Then
                                                    Dim questionText As String = ""
                                                    index = 0
                                                    Dim currentQuestion As Integer = 0
                                                    For Each row In Me.dsCommon.Tables("DataDump").Rows
                                                        If (row.Item("QUESTION_GROUP_ID") = tempFilterItemArray(5) _
                                                            And tempSkipList.Contains(CType(row.Item("QUESTION_ID"), String))) Then
                                                            If (questionText = "") Then
                                                                questionText = row.Item("QuestionText")
                                                                currentQuestion = row.Item("QUESTION_ID")

                                                                'generate a row for the question text
                                                                genericRow = New System.Web.UI.WebControls.TableRow
                                                                genericCell = New System.Web.UI.WebControls.TableCell
                                                                genericLabel = New System.Web.UI.WebControls.Label
                                                                genericLabel.CssClass = "boldContentRed"
                                                                genericLabel.Text = index + 1 & ". "
                                                                genericCell.Controls.Add(genericLabel)
                                                                genericLabel = New System.Web.UI.WebControls.Label
                                                                genericLabel.CssClass = "reportContent"
                                                                genericLabel.Text = questionText.Replace("~", ",")
                                                                genericCell.Controls.Add(genericLabel)
                                                                genericCell.Wrap = True
                                                                genericCell.EnableViewState = True
                                                                genericRow.Cells.Add(genericCell)
                                                                genericLabel = New System.Web.UI.WebControls.Label
                                                                genericLabel.CssClass = "boldContentTeal"
                                                                genericLabel.Text = "(" & sumAnswered(index) & ")"
                                                                genericCell.Controls.Add(genericLabel)
                                                                genericCell.Wrap = True
                                                                genericCell.EnableViewState = True
                                                                genericCell.ColumnSpan = 2
                                                                genericRow.Cells.Add(genericCell)
                                                                genericRow.EnableViewState = True
                                                                Me.DataTable2.Rows.Add(genericRow)
                                                            ElseIf (currentQuestion <> row.Item("QUESTION_ID")) Then
                                                                questionText = row.Item("QuestionText")
                                                                currentQuestion = row.Item("QUESTION_ID")
                                                                index += 1

                                                                'generate a row for the question text
                                                                genericRow = New System.Web.UI.WebControls.TableRow
                                                                genericCell = New System.Web.UI.WebControls.TableCell
                                                                genericLabel = New System.Web.UI.WebControls.Label
                                                                genericLabel.CssClass = "boldContentRed"
                                                                genericLabel.Text = index + 1 & ". "
                                                                genericCell.Controls.Add(genericLabel)
                                                                genericLabel = New System.Web.UI.WebControls.Label
                                                                genericLabel.CssClass = "reportContent"
                                                                genericLabel.Text = questionText.Replace("~", ",")
                                                                genericCell.Controls.Add(genericLabel)
                                                                genericCell.Wrap = True
                                                                genericCell.EnableViewState = True
                                                                genericRow.Cells.Add(genericCell)
                                                                genericLabel = New System.Web.UI.WebControls.Label
                                                                genericLabel.CssClass = "boldContentTeal"
                                                                genericLabel.Text = "(" & sumAnswered(index) & ")"
                                                                genericCell.Controls.Add(genericLabel)
                                                                genericCell.Wrap = True
                                                                genericCell.EnableViewState = True
                                                                genericCell.ColumnSpan = 2
                                                                genericRow.Cells.Add(genericCell)
                                                                genericRow.EnableViewState = True
                                                                Me.DataTable2.Rows.Add(genericRow)
                                                            End If
                                                        End If
                                                    Next
                                                Else
                                                    'generate a row for the question text
                                                    genericRow = New System.Web.UI.WebControls.TableRow
                                                    genericCell = New System.Web.UI.WebControls.TableCell
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "reportContent"
                                                    genericLabel.Text = CType(tempFilterItemArray(3), String).Replace("~", ",")
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.Wrap = True
                                                    genericCell.EnableViewState = True
                                                    genericRow.Cells.Add(genericCell)
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "boldContentTeal"
                                                    genericLabel.Text = "(" & sumAnswered(0) & ")"
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.Wrap = True
                                                    genericCell.EnableViewState = True
                                                    genericCell.ColumnSpan = 2
                                                    genericRow.Cells.Add(genericCell)
                                                    genericRow.EnableViewState = True
                                                    Me.DataTable2.Rows.Add(genericRow)
                                                End If
                                            End If

                                            If (Session("SplitOutput") <> "NONE") Then
                                                'generate a row for the date range
                                                genericRow = New System.Web.UI.WebControls.TableRow
                                                genericCell = New System.Web.UI.WebControls.TableCell
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "ReportHeaderData"
                                                genericLabel.Text = "Period of Record Set: "
                                                genericCell.Controls.Add(genericLabel)
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "ReportHeaderBlue"
                                                genericLabel.Text = dateFilter(0) & " - " & dateFilter(1)
                                                genericCell.Controls.Add(genericLabel)
                                                genericCell.Wrap = True
                                                genericCell.EnableViewState = True
                                                genericCell.ColumnSpan = 2
                                                genericRow.Cells.Add(genericCell)
                                                genericRow.EnableViewState = True
                                                genericRow.BackColor = Color.LightGray
                                                Me.DataTable2.Rows.Add(genericRow)
                                            End If

                                            'generate a row for the question answer
                                            genericTable = New System.Web.UI.WebControls.Table
                                            Dim colorIndex As Integer = 0
                                            blnDoOnce3 = False
                                            For Each row In dt.Rows
                                                'get the color for the current row
                                                Dim strCurrentColor As String = "FFFFFF"
                                                If (Me.dsCommon.Tables("Colors").Rows.Count() > 0) Then
                                                    strCurrentColor = Me.dsCommon.Tables("Colors").Rows(colorIndex).Item("CHART_COLOR_CODE")
                                                    colorIndex += 1
                                                    If (colorIndex = Me.dsCommon.Tables("Colors").Rows.Count()) Then
                                                        colorIndex = 0
                                                    End If
                                                End If

                                                'set up the color for the row
                                                Dim Alpha As Integer = 0
                                                Dim Red As Integer = 0
                                                Dim Green As Integer = 0
                                                Dim Blue As Integer = 0
                                                If (strCurrentColor.Length = 8) Then
                                                    Alpha = Integer.Parse(strCurrentColor.Substring(0, 2), Globalization.NumberStyles.HexNumber)
                                                    Red = Integer.Parse(strCurrentColor.Substring(2, 2), Globalization.NumberStyles.HexNumber)
                                                    Green = Integer.Parse(strCurrentColor.Substring(4, 2), Globalization.NumberStyles.HexNumber)
                                                    Blue = Integer.Parse(strCurrentColor.Substring(6, 2), Globalization.NumberStyles.HexNumber)
                                                Else
                                                    Red = Integer.Parse(strCurrentColor.Substring(0, 2), Globalization.NumberStyles.HexNumber)
                                                    Green = Integer.Parse(strCurrentColor.Substring(2, 2), Globalization.NumberStyles.HexNumber)
                                                    Blue = Integer.Parse(strCurrentColor.Substring(4, 2), Globalization.NumberStyles.HexNumber)
                                                End If

                                                'table header setup - do once
                                                If Not (blnDoOnce3) Then
                                                    genericRow = New System.Web.UI.WebControls.TableRow
                                                    genericCell = New System.Web.UI.WebControls.TableCell
                                                    genericCell.Width = Unit.Pixel(20)
                                                    genericCell.Wrap = False
                                                    genericCell.EnableViewState = True
                                                    genericRow.Cells.Add(genericCell)
                                                    genericCell = New System.Web.UI.WebControls.TableCell
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "reportContent"
                                                    genericLabel.Text = "Answer Text"
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.Wrap = True
                                                    genericCell.EnableViewState = True
                                                    genericCell.BackColor = System.Drawing.Color.AntiqueWhite
                                                    genericRow.Cells.Add(genericCell)

                                                    If (tempFilterItemArray(6) = "MULTI") Then
                                                        index = 1
                                                        While (index <= multiCount)
                                                            genericCell = New System.Web.UI.WebControls.TableCell
                                                            genericLabel = New System.web.ui.webcontrols.Label
                                                            genericLabel.CssClass = "reportContent"
                                                            genericLabel.Text = index
                                                            genericCell.Controls.Add(genericLabel)
                                                            genericCell.HorizontalAlign = HorizontalAlign.Center
                                                            genericCell.Wrap = False
                                                            genericCell.EnableViewState = True
                                                            genericCell.BackColor = System.Drawing.Color.AntiqueWhite
                                                            genericRow.Cells.Add(genericCell)
                                                            index += 1
                                                        End While
                                                    Else
                                                        genericCell = New System.Web.UI.WebControls.TableCell
                                                        genericLabel = New System.web.ui.webcontrols.Label
                                                        genericLabel.CssClass = "reportContent"

                                                        'change the header based on percent or count display
                                                        If (blnPercentDisplay) Then
                                                            genericLabel.Text = "Percent"
                                                        Else
                                                            genericLabel.Text = "Count"
                                                        End If

                                                        genericCell.Controls.Add(genericLabel)
                                                        genericCell.HorizontalAlign = HorizontalAlign.Center
                                                        genericCell.Wrap = False
                                                        genericCell.EnableViewState = True
                                                        genericCell.BackColor = System.Drawing.Color.AntiqueWhite
                                                        genericRow.Cells.Add(genericCell)
                                                    End If

                                                    genericRow.EnableViewState = True
                                                    genericTable.Rows.Add(genericRow)
                                                    blnDoOnce3 = True
                                                End If

                                                'populate row data
                                                genericRow = New System.Web.UI.WebControls.TableRow
                                                genericCell = New System.Web.UI.WebControls.TableCell
                                                genericCell.Width = Unit.Pixel(20)
                                                genericCell.Wrap = False
                                                genericCell.EnableViewState = True
                                                genericRow.Cells.Add(genericCell)
                                                genericCell = New System.Web.UI.WebControls.TableCell
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "reportContent"
                                                genericLabel.Text = row.Item("ListLabel")
                                                genericCell.Wrap = True
                                                genericCell.EnableViewState = True
                                                genericCell.BackColor = System.Drawing.Color.FromArgb(Alpha, Red, Green, Blue)
                                                genericCell.Controls.Add(genericLabel)
                                                genericRow.Cells.Add(genericCell)

                                                If (tempFilterItemArray(6) = "MULTI") Then
                                                    'format the value based on percent or count display
                                                    index = 1
                                                    While (index <= multiCount)
                                                        genericCell = New System.Web.UI.WebControls.TableCell
                                                        genericLabel = New System.web.ui.webcontrols.Label
                                                        genericLabel.CssClass = "reportContent"

                                                        If (blnPercentDisplay) Then
                                                            genericLabel.Text = CType(row.Item("Percentage" & index), Double).ToString("#0.##")
                                                        Else
                                                            genericLabel.Text = row.Item("AnswerCount" & index)
                                                        End If
                                                        genericCell.Controls.Add(genericLabel)
                                                        genericCell.HorizontalAlign = HorizontalAlign.Right
                                                        genericCell.Wrap = False
                                                        genericCell.EnableViewState = True
                                                        genericCell.BackColor = System.Drawing.Color.FromArgb(Alpha, Red, Green, Blue)
                                                        genericRow.Cells.Add(genericCell)
                                                        index += 1
                                                    End While
                                                Else
                                                    genericCell = New System.Web.UI.WebControls.TableCell
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "reportContent"

                                                    'format the value based on percent or count display
                                                    If (blnPercentDisplay) Then
                                                        genericLabel.Text = CType(row.Item("Percentage"), Double).ToString("#0.##")
                                                    Else
                                                        genericLabel.Text = row.Item("AnswerCount")
                                                    End If
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.HorizontalAlign = HorizontalAlign.Right
                                                    genericCell.Wrap = False
                                                    genericCell.EnableViewState = True
                                                    genericCell.BackColor = System.Drawing.Color.FromArgb(Alpha, Red, Green, Blue)
                                                    genericRow.Cells.Add(genericCell)
                                                End If

                                                genericRow.EnableViewState = True
                                                genericTable.Rows.Add(genericRow)
                                            Next

                                            genericRow = New System.Web.UI.WebControls.TableRow
                                            genericCell = New System.Web.UI.WebControls.TableCell
                                            genericCell.Controls.Add(genericTable)
                                            genericCell.Wrap = False
                                            genericCell.EnableViewState = True
                                            genericCell.VerticalAlign = VerticalAlign.Top
                                            genericRow.Cells.Add(genericCell)

                                            'generate the chart for the Rating question
                                            genericCell = New System.Web.UI.WebControls.TableCell
                                            genericChart = New Infragistics.webui.UltraWebChart.UltraChart
                                            If Not (generateChart(genericChart, tempFilterItemArray(0), tempFilterItemArray(6), dateFilter(2), dt, tempSkipList, runCount, True, False, False, blnPercentDisplay)) Then
                                                Session("Alert") = "Chart generation failure."
                                                Return False
                                            End If
                                            genericCell.Controls.Add(genericChart)
                                            genericCell.HorizontalAlign = HorizontalAlign.Right
                                            genericCell.VerticalAlign = VerticalAlign.Top
                                            genericCell.Wrap = False
                                            genericCell.EnableViewState = True

                                            'Set to put a line break after 3 question results appear on the page.
                                            If (questionCounter Mod Session("QuestionsPerPage") = 0) Then
                                                genericRow.CssClass = "Break"
                                            End If
                                            questionCounter += 1

                                            genericRow.Cells.Add(genericCell)

                                            'finalize the row
                                            genericRow.EnableViewState = True
                                            Me.DataTable2.Rows.Add(genericRow)
                                            genericRow = New System.Web.UI.WebControls.TableRow
                                            genericCell = New System.Web.UI.WebControls.TableCell
                                            genericCell.Height = Unit.Pixel(40)
                                            genericCell.Wrap = False
                                            genericCell.EnableViewState = True
                                            genericRow.Cells.Add(genericCell)
                                            genericRow.EnableViewState = True
                                            Me.DataTable2.Rows.Add(genericRow)
                                            Me.DataTable2.EnableViewState = True
                                            runCount += 1
                                        Else
                                            If (Not blnDoOnce3) Then
                                                'section header
                                                genericRow = New System.Web.UI.WebControls.TableRow
                                                genericCell = New System.Web.UI.WebControls.TableCell
                                                genericCell.Wrap = False
                                                genericCell.CssClass = "sectionHeadTeal"
                                                genericCell.Text = "Other Results"
                                                genericCell.EnableViewState = True
                                                genericCell.ColumnSpan = 2
                                                genericRow.Cells.Add(genericCell)
                                                genericRow.EnableViewState = True
                                                noDataRows.Add(genericRow)
                                                'Me.DataTable2.Rows.Add(genericRow)
                                                blnDoOnce3 = True
                                            End If

                                            'buffer zone
                                            'genericRow = New System.Web.UI.WebControls.TableRow
                                            'genericCell = New System.Web.UI.WebControls.TableCell
                                            'genericCell.Wrap = False
                                            'genericCell.Height = Unit.Pixel(40)
                                            'genericCell.EnableViewState = True
                                            'genericCell.ColumnSpan = 2
                                            'genericRow.Cells.Add(genericCell)
                                            'genericRow.EnableViewState = True
                                            'Me.DataTable2.Rows.Add(genericRow)

                                            If (Session("SplitOutput") <> "NONE") Then
                                                If (tempFilterItemArray(6) = "MULTI") Then
                                                    index = 0
                                                    While (index < skipListText.Count)
                                                        'generate a row for the question text
                                                        genericRow = New System.Web.UI.WebControls.TableRow
                                                        genericCell = New System.Web.UI.WebControls.TableCell
                                                        genericLabel = New System.web.ui.webcontrols.Label
                                                        genericLabel.CssClass = "reportContent"
                                                        genericLabel.Text = CType(tempSkipListText(index), String).Replace("~", ",")
                                                        genericCell.Controls.Add(genericLabel)
                                                        genericLabel = New System.web.ui.webcontrols.Label
                                                        genericLabel.CssClass = "boldContentTeal"
                                                        genericLabel.Text = "(" & sumAnswered(index) & ")"
                                                        genericCell.Controls.Add(genericLabel)
                                                        genericCell.Wrap = True
                                                        genericCell.EnableViewState = True
                                                        genericCell.ColumnSpan = 2
                                                        genericRow.Cells.Add(genericCell)
                                                        genericRow.EnableViewState = True
                                                        If (Session("SplitOutput") <> "NONE") Then
                                                            If (blnDoOnce2) Then
                                                                genericRow.BackColor = Color.LightSteelBlue
                                                            Else
                                                                genericRow.BackColor = Color.LemonChiffon
                                                            End If
                                                        End If
                                                        noDataRows.Add(genericRow)
                                                        'Me.DataTable2.Rows.Add(genericRow)
                                                        blnDoOnce2 = True
                                                        index += 1
                                                    End While
                                                Else
                                                    'generate a row for the question text
                                                    genericRow = New System.Web.UI.WebControls.TableRow
                                                    genericCell = New System.Web.UI.WebControls.TableCell
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "reportContent"
                                                    genericLabel.Text = CType(tempFilterItemArray(3), String).Replace("~", ",")
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "boldContentTeal"
                                                    genericLabel.Text = "(" & sumAnswered(index) & ")"
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.Wrap = True
                                                    genericCell.EnableViewState = True
                                                    genericCell.ColumnSpan = 2
                                                    genericRow.Cells.Add(genericCell)
                                                    genericRow.EnableViewState = True
                                                    If (Session("SplitOutput") <> "NONE") Then
                                                        If (blnDoOnce2) Then
                                                            genericRow.BackColor = Color.LightSteelBlue
                                                        Else
                                                            genericRow.BackColor = Color.LemonChiffon
                                                        End If
                                                    End If
                                                    noDataRows.Add(genericRow)
                                                    'Me.DataTable2.Rows.Add(genericRow)
                                                    blnDoOnce2 = True
                                                End If
                                            ElseIf (Session("SplitOutput") = "NONE") Then
                                                If (tempFilterItemArray(6) = "MULTI") Then
                                                    index = 0
                                                    While (index < skipListText.Count)
                                                        'generate a row for the question text
                                                        genericRow = New System.Web.UI.WebControls.TableRow
                                                        genericCell = New System.Web.UI.WebControls.TableCell
                                                        genericLabel = New System.web.ui.webcontrols.Label
                                                        genericLabel.CssClass = "reportContent"
                                                        genericLabel.Text = CType(tempSkipListText(index), String).Replace("~", ",")
                                                        genericCell.Controls.Add(genericLabel)
                                                        genericCell.Wrap = True
                                                        genericCell.EnableViewState = True
                                                        genericRow.Cells.Add(genericCell)
                                                        genericLabel = New System.web.ui.webcontrols.Label
                                                        genericLabel.CssClass = "boldContentTeal"
                                                        genericLabel.Text = "(" & sumAnswered(index) & ")"
                                                        genericCell.Controls.Add(genericLabel)
                                                        genericCell.Wrap = True
                                                        genericCell.EnableViewState = True
                                                        genericRow.Cells.Add(genericCell)
                                                        genericRow.EnableViewState = True
                                                        noDataRows.Add(genericRow)
                                                        'Me.DataTable2.Rows.Add(genericRow)
                                                        index += 1
                                                    End While
                                                Else
                                                    'generate a row for the question text
                                                    genericRow = New System.Web.UI.WebControls.TableRow
                                                    genericCell = New System.Web.UI.WebControls.TableCell
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "reportContent"
                                                    genericLabel.Text = CType(tempFilterItemArray(3), String).Replace("~", ",")
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.Wrap = True
                                                    genericCell.EnableViewState = True
                                                    genericRow.Cells.Add(genericCell)
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "boldContentTeal"
                                                    index = 0
                                                    Dim wholeAnswer As Integer = 0
                                                    While (index < sumAnswered.Length())
                                                        wholeAnswer += sumAnswered(index)
                                                        index += 1
                                                    End While
                                                    genericLabel.Text = "(" & wholeAnswer & ")"
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.Wrap = True
                                                    genericCell.EnableViewState = True
                                                    genericRow.Cells.Add(genericCell)
                                                    genericRow.EnableViewState = True
                                                    noDataRows.Add(genericRow)
                                                    'Me.DataTable2.Rows.Add(genericRow)
                                                End If
                                            End If

                                            If (Session("SplitOutput") <> "NONE") Then
                                                'generate a row for the date range
                                                genericRow = New System.Web.UI.WebControls.TableRow
                                                genericCell = New System.Web.UI.WebControls.TableCell
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "ReportHeaderData"
                                                genericLabel.Text = "Period of Record Set: "
                                                genericCell.Controls.Add(genericLabel)
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "ReportHeaderBlue"
                                                genericLabel.Text = dateFilter(0) & " - " & dateFilter(1)
                                                genericCell.Controls.Add(genericLabel)
                                                genericCell.Wrap = True
                                                genericCell.EnableViewState = True
                                                genericCell.ColumnSpan = 2
                                                genericRow.Cells.Add(genericCell)
                                                genericRow.EnableViewState = True
                                                genericRow.BackColor = Color.LightGray
                                                noDataRows.Add(genericRow)
                                                'Me.DataTable2.Rows.Add(genericRow)
                                            End If

                                            'generate to inform user that there is no data for this date range
                                            genericRow = New System.Web.UI.WebControls.TableRow
                                            genericTable = New System.Web.UI.WebControls.Table
                                            genericCell = New System.Web.UI.WebControls.TableCell
                                            genericCell.Width = Unit.Pixel(20)
                                            genericCell.Wrap = False
                                            genericCell.EnableViewState = True
                                            genericRow.Cells.Add(genericCell)
                                            genericCell = New System.Web.UI.WebControls.TableCell
                                            genericLabel = New System.web.ui.webcontrols.Label
                                            genericLabel.CssClass = "NoDataRed"
                                            genericLabel.Text = "This question has no data for this time period."
                                            genericCell.Controls.Add(genericLabel)
                                            'genericCell.ColumnSpan = 2
                                            'genericCell.HorizontalAlign = HorizontalAlign.Center
                                            genericRow.Cells.Add(genericCell)
                                            genericRow.EnableViewState = True
                                            genericTable.Rows.Add(genericRow)
                                            genericRow = New System.Web.UI.WebControls.TableRow
                                            genericCell = New System.Web.UI.WebControls.TableCell
                                            genericCell.ColumnSpan = 2
                                            genericCell.Controls.Add(genericTable)
                                            genericRow.Cells.Add(genericCell)
                                            noDataRows.Add(genericRow)
                                            'Me.DataTable2.Rows.Add(genericRow)
                                        End If
                                    Else
                                        If (tempFilterItemArray(2) = "D") Then
                                            If (Not blnDoOnce3) Then
                                                'section header
                                                genericRow = New System.Web.UI.WebControls.TableRow
                                                genericCell = New System.Web.UI.WebControls.TableCell
                                                genericCell.Wrap = False
                                                genericCell.CssClass = "sectionHeadTeal"
                                                genericCell.Text = "Other Results"
                                                genericCell.EnableViewState = True
                                                genericCell.ColumnSpan = 2
                                                genericRow.Cells.Add(genericCell)
                                                genericRow.EnableViewState = True
                                                noDataRows.Add(genericRow)
                                                'Me.DataTable2.Rows.Add(genericRow)
                                                blnDoOnce3 = True
                                            End If

                                            'buffer zone
                                            'genericRow = New System.Web.UI.WebControls.TableRow
                                            'genericCell = New System.Web.UI.WebControls.TableCell
                                            'genericCell.Wrap = False
                                            'genericCell.Height = Unit.Pixel(40)
                                            'genericCell.EnableViewState = True
                                            'genericCell.ColumnSpan = 2
                                            'genericRow.Cells.Add(genericCell)
                                            'genericRow.EnableViewState = True
                                            'Me.DataTable2.Rows.Add(genericRow)

                                            If (Session("SplitOutput") <> "NONE") Then
                                                If (tempFilterItemArray(6) = "MULTI") Then
                                                    index = 0
                                                    While (index < skipListText.Count)
                                                        'generate a row for the question text
                                                        genericRow = New System.Web.UI.WebControls.TableRow
                                                        genericCell = New System.Web.UI.WebControls.TableCell
                                                        genericLabel = New System.web.ui.webcontrols.Label
                                                        genericLabel.CssClass = "reportContent"
                                                        genericLabel.Text = CType(tempSkipListText(index), String).Replace("~", ",")
                                                        genericCell.Controls.Add(genericLabel)
                                                        genericLabel = New System.web.ui.webcontrols.Label
                                                        genericLabel.CssClass = "boldContentTeal"
                                                        genericLabel.Text = "(" & sumAnswered(index) & ")"
                                                        genericCell.Controls.Add(genericLabel)
                                                        genericCell.Wrap = True
                                                        genericCell.EnableViewState = True
                                                        genericCell.ColumnSpan = 2
                                                        genericRow.Cells.Add(genericCell)
                                                        genericRow.EnableViewState = True
                                                        If (Session("SplitOutput") <> "NONE") Then
                                                            If (blnDoOnce2) Then
                                                                genericRow.BackColor = Color.LightSteelBlue
                                                            Else
                                                                genericRow.BackColor = Color.LemonChiffon
                                                            End If
                                                        End If
                                                        noDataRows.Add(genericRow)
                                                        'Me.DataTable2.Rows.Add(genericRow)
                                                        blnDoOnce2 = True
                                                        index += 1
                                                    End While
                                                Else
                                                    'generate a row for the question text
                                                    genericRow = New System.Web.UI.WebControls.TableRow
                                                    genericCell = New System.Web.UI.WebControls.TableCell
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "reportContent"
                                                    genericLabel.Text = CType(tempFilterItemArray(3), String).Replace("~", ",")
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "boldContentTeal"
                                                    genericLabel.Text = "(" & sumAnswered(index) & ")"
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.Wrap = True
                                                    genericCell.EnableViewState = True
                                                    genericCell.ColumnSpan = 2
                                                    genericRow.Cells.Add(genericCell)
                                                    genericRow.EnableViewState = True
                                                    If (Session("SplitOutput") <> "NONE") Then
                                                        If (blnDoOnce2) Then
                                                            genericRow.BackColor = Color.LightSteelBlue
                                                        Else
                                                            genericRow.BackColor = Color.LemonChiffon
                                                        End If
                                                    End If
                                                    noDataRows.Add(genericRow)
                                                    'Me.DataTable2.Rows.Add(genericRow)
                                                    blnDoOnce2 = True
                                                End If
                                            ElseIf (Session("SplitOutput") = "NONE") Then
                                                If (tempFilterItemArray(6) = "MULTI") Then
                                                    index = 0
                                                    While (index < skipListText.Count)
                                                        'generate a row for the question text
                                                        genericRow = New System.Web.UI.WebControls.TableRow
                                                        genericCell = New System.Web.UI.WebControls.TableCell
                                                        genericLabel = New System.web.ui.webcontrols.Label
                                                        genericLabel.CssClass = "reportContent"
                                                        genericLabel.Text = CType(tempSkipListText(index), String).Replace("~", ",")
                                                        genericCell.Controls.Add(genericLabel)
                                                        genericCell.Wrap = True
                                                        genericCell.EnableViewState = True
                                                        genericRow.Cells.Add(genericCell)
                                                        genericLabel = New System.web.ui.webcontrols.Label
                                                        genericLabel.CssClass = "boldContentTeal"
                                                        genericLabel.Text = "(" & sumAnswered(index) & ")"
                                                        genericCell.Controls.Add(genericLabel)
                                                        genericCell.Wrap = True
                                                        genericCell.EnableViewState = True
                                                        genericRow.Cells.Add(genericCell)
                                                        genericRow.EnableViewState = True
                                                        noDataRows.Add(genericRow)
                                                        'Me.DataTable2.Rows.Add(genericRow)
                                                        index += 1
                                                    End While
                                                Else
                                                    'generate a row for the question text
                                                    genericRow = New System.Web.UI.WebControls.TableRow
                                                    genericCell = New System.Web.UI.WebControls.TableCell
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "reportContent"
                                                    genericLabel.Text = CType(tempFilterItemArray(3), String).Replace("~", ",")
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.Wrap = True
                                                    genericCell.EnableViewState = True
                                                    genericRow.Cells.Add(genericCell)
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "boldContentTeal"
                                                    index = 0
                                                    Dim wholeAnswer As Integer = 0
                                                    While (index < sumAnswered.Length())
                                                        wholeAnswer += sumAnswered(index)
                                                        index += 1
                                                    End While
                                                    genericLabel.Text = "(" & wholeAnswer & ")"
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.Wrap = True
                                                    genericCell.EnableViewState = True
                                                    genericRow.Cells.Add(genericCell)
                                                    genericRow.EnableViewState = True
                                                    noDataRows.Add(genericRow)
                                                    'Me.DataTable2.Rows.Add(genericRow)
                                                End If
                                            End If

                                            If (Session("SplitOutput") <> "NONE") Then
                                                'generate a row for the date range
                                                genericRow = New System.Web.UI.WebControls.TableRow
                                                genericCell = New System.Web.UI.WebControls.TableCell
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "ReportHeaderData"
                                                genericLabel.Text = "Period of Record Set: "
                                                genericCell.Controls.Add(genericLabel)
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "ReportHeaderBlue"
                                                genericLabel.Text = dateFilter(0) & " - " & dateFilter(1)
                                                genericCell.Controls.Add(genericLabel)
                                                genericCell.Wrap = True
                                                genericCell.EnableViewState = True
                                                genericCell.ColumnSpan = 2
                                                genericRow.Cells.Add(genericCell)
                                                genericRow.EnableViewState = True
                                                genericRow.BackColor = Color.LightGray
                                                noDataRows.Add(genericRow)
                                                'Me.DataTable2.Rows.Add(genericRow)
                                            End If

                                            'generate to inform user that there is no data for this date range
                                            genericRow = New System.Web.UI.WebControls.TableRow
                                            genericTable = New System.Web.UI.WebControls.Table
                                            genericCell = New System.Web.UI.WebControls.TableCell
                                            genericCell.Width = Unit.Pixel(20)
                                            genericCell.Wrap = False
                                            genericCell.EnableViewState = True
                                            genericRow.Cells.Add(genericCell)
                                            genericCell = New System.Web.UI.WebControls.TableCell
                                            genericLabel = New System.web.ui.webcontrols.Label
                                            genericLabel.CssClass = "NoDataRed"
                                            genericLabel.Text = "This question has no data for this time period."
                                            genericCell.Controls.Add(genericLabel)
                                            'genericCell.ColumnSpan = 2
                                            'genericCell.HorizontalAlign = HorizontalAlign.Center
                                            genericRow.Cells.Add(genericCell)
                                            genericRow.EnableViewState = True
                                            genericTable.Rows.Add(genericRow)
                                            genericRow = New System.Web.UI.WebControls.TableRow
                                            genericCell = New System.Web.UI.WebControls.TableCell
                                            genericCell.ColumnSpan = 2
                                            genericCell.Controls.Add(genericTable)
                                            genericRow.Cells.Add(genericCell)
                                            noDataRows.Add(genericRow)
                                            'Me.DataTable2.Rows.Add(genericRow)
                                        Else
                                            Session("Alert") = "Failed to get survey question list item data for choice questions."
                                            Return False
                                        End If
                                    End If
                                Else
                                    Session("Alert") = "Failed to get survey question list item data for choice questions."
                                    Return False
                                End If
                            Next
                        End If
                    End If
                End If
            Next

            'generate to inform user that there is no data for this date range
            'genericRow = New System.Web.UI.WebControls.TableRow
            'genericRow.CssClass = "Break"
            'genericRow.EnableViewState = True
            'noDataRows.Add(genericRow)

            'get the questions that have no data
            Dim tableRow As System.Web.UI.WebControls.TableRow
            For Each tableRow In noDataRows
                Me.DataTable2.Rows.Add(tableRow)
            Next

            'Dim lastRow As System.Web.UI.WebControls.TableRow = Me.DataTable2.Rows(Me.DataTable2.Rows.Count() - 1)
            'lastRow.CssClass = "Break"

            'If (filterItemCount = 0) Then
            'generate a line break
            'genericRow = New System.Web.UI.WebControls.TableRow
            'genericRow.CssClass = "Break"
            'genericRow.EnableViewState = True
            'Me.DataTable2.Rows.Add(genericRow)
            'End If

            Return True
        Catch ex As Exception
            Trace.Warn(ex.ToString)
            Session("Alert") = "Failed to get survey data.  " & ex.ToString
            Return False
        End Try
    End Function

    'processes the rest of the data
    Private Function processOtherData() As Boolean
        Trace.Warn("processing other section")
        Try
            'define the variables
            Dim genericCell As System.Web.UI.WebControls.TableCell
            Dim genericRow As System.Web.UI.WebControls.TableRow
            Dim genericTable As System.Web.UI.WebControls.Table
            Dim genericLabel As System.Web.UI.WebControls.Label
            Dim genericChart As Infragistics.WebUI.UltraWebChart.UltraChart
            Dim noDataRows As ArrayList = New ArrayList

            'Set up the filter
            Dim filterArray As Array
            Dim skipList As ArrayList = New ArrayList
            Dim skipListText As ArrayList = New ArrayList
            filterArray = CType(Session("DataItems"), String).Split(",")

            'get filter item data
            Dim filterItem As String = ""
            Dim blnDoOnce As Boolean = False
            Dim blnDoOnce3 As Boolean = False
            Dim questionCounter As Integer = 1
            For Each filterItem In filterArray
                Dim tempFilterItemArray As Array = filterItem.Split("^")
                'process type questions (except R type) with averages and the question label
                If (CType(tempFilterItemArray(1), Boolean) And tempFilterItemArray(2) <> "R" _
                    And tempFilterItemArray(2) <> "C") Then
                    Dim dr As DataRow
                    Dim blnPercentDisplay As Boolean = False

                    'skip any questions in the skip list
                    If Not (skipList.Contains(CType(tempFilterItemArray(0), String))) Then
                        'add a separator between the questions and the rest of the data
                        If (blnDoOnce = False) Then
                            blnDoOnce = True
                            If (Session("NonRTypeSelected") = True Or Session("RItemExists") = True) Then
                                'generate a row for the question text
                                genericRow = New System.Web.UI.WebControls.TableRow
                                Me.DataTable2.Rows(Me.DataTable2.Rows.Count() - 2).CssClass = ""
                                genericRow.CssClass = "BreakBefore"
                                genericCell = New System.Web.UI.WebControls.TableCell
                                genericLabel = New System.web.ui.webcontrols.Label
                                genericLabel.Text = "<hr style='BORDER-TOP-STYLE: ridge' width='100%' SIZE='3'/>"
                                genericCell.Controls.Add(genericLabel)
                                genericCell.Wrap = False
                                genericCell.EnableViewState = True
                                genericCell.ColumnSpan = 2
                                genericRow.Cells.Add(genericCell)
                                genericRow.EnableViewState = True
                                Me.DataTable2.Rows.Add(genericRow)
                            End If
                        End If

                        'get the list of questions in the multi-group if the question chart type is multi
                        Dim tempSkipList As ArrayList = New ArrayList
                        Dim tempSkipListText As ArrayList = New ArrayList
                        If (tempFilterItemArray(6) = "MULTI") Then
                            Dim tempItem As String = ""
                            For Each tempItem In filterArray
                                Dim tempSplit As Array = tempItem.Split("^")
                                If (CType(tempSplit(0), Integer) > CType(tempFilterItemArray(0), Integer) And CType(tempSplit(1), Boolean) _
                                    And tempSplit(6) = "MULTI") Then
                                    If (tempSplit(5) = tempFilterItemArray(5)) Then
                                        tempSkipList.Add(tempSplit(0))
                                        tempSkipListText.Add(tempSplit(3))
                                        If (tempSplit(4) = "P-True") Then
                                            blnPercentDisplay = True
                                        End If
                                    End If
                                ElseIf (CType(tempSplit(0), Integer) = CType(tempFilterItemArray(0), Integer)) Then
                                    tempSkipListText.Add(tempSplit(3))
                                    If (tempSplit(4) = "P-True") Then
                                        blnPercentDisplay = True
                                    End If
                                End If
                            Next
                        Else
                            If (tempFilterItemArray(4) = "P-True") Then
                                blnPercentDisplay = True
                            End If
                        End If

                        'build skiplist string for queries
                        Dim strSkipList As String = tempFilterItemArray(0)
                        Dim index As Integer = 0
                        While (index < tempSkipList.Count)
                            strSkipList &= ", " & tempSkipList(index)
                            index += 1
                        End While
                        tempSkipList.Add(tempFilterItemArray(0))

                        'add skip items to master list
                        index = 0
                        skipListText.Clear()
                        skipListText.Add(tempFilterItemArray(3))
                        While (index < tempSkipList.Count - 1)
                            skipList.Add(tempSkipList(index))
                            skipListText.Add(tempSkipListText(index))
                            index += 1
                        End While
                        skipList.Add(tempFilterItemArray(0))

                        'get the choices from the fempList table or fempResults table (for D Type questions)
                        If (tempFilterItemArray(6) = "MULTI") Then
                            If (tempFilterItemArray(2) = "D") Then
                                Me.sqlCommonCommand.CommandText = "select distinct cast(CAST(" & _
                                    "fr.ANSWER_TEXT as datetime) as int) as intListValue, fr.ANSWER_TEXT " & _
                                    "as strLabel from " & myUtility.getExtension() & "SAT_RESPONSE_RESULT fr, " & _
                                    myUtility.getExtension() & "SAT_TEMPLATE_SURVEY fs, " & myUtility.getExtension() & _
                                    "SAT_TEMPLATE_QUESTION fq where fr.SURVEY_KEY = " & _
                                    Session("seqSurvey") & " And fr.QUESTION_ID in (" & strSkipList & _
                                    ") and fr.QUESTION_ID = fq.QUESTION_ID " & _
                                    "and fr.SURVEY_KEY = fs.SURVEY_KEY and fs.TEMPLATE_KEY = fq.TEMPLATE_KEY " & _
                                    "and fq.QUESTION_GROUP_ID = " & tempFilterItemArray(5) & _
                                    " order by fr.ANSWER_TEXT"
                                Trace.Warn(Me.sqlCommonCommand.CommandText)
                            Else
                                Me.sqlCommonCommand.CommandText = "select distinct fl.LIST_ITEM_ID as intList, fl.LIST_ITEM_VALUE as intListValue, fl.LIST_ITEM_TITLE as strLabel " & _
                                    "from " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM fl, " & myUtility.getExtension() & _
                                    "SAT_TEMPLATE_SURVEY fs, " & myUtility.getExtension() & "SAT_TEMPLATE_QUESTION fq where fs.SURVEY_KEY = " & _
                                    Session("seqSurvey") & _
                                    " And fs.TEMPLATE_KEY = fl.TEMPLATE_KEY and fl.QUESTION_ID in (" & strSkipList & _
                                    ") and fl.TEMPLATE_KEY = fq.TEMPLATE_KEY and fl.QUESTION_ID = fq.QUESTION_ID " & _
                                    "and fq.QUESTION_GROUP_ID = " & tempFilterItemArray(5) & _
                                    " order by fl.LIST_ITEM_ID, fl.LIST_ITEM_VALUE"
                                Trace.Warn(Me.sqlCommonCommand.CommandText)
                            End If
                        Else
                            If (tempFilterItemArray(2) = "D") Then
                                Me.sqlCommonCommand.CommandText = "select distinct cast(CAST(" & _
                                    "ANSWER_TEXT as datetime) as int) as intListValue, ANSWER_TEXT " & _
                                    "as strLabel from " & myUtility.getExtension() & "SAT_RESPONSE_RESULT where SURVEY_KEY = " & _
                                    Session("seqSurvey") & _
                                    " And QUESTION_ID = " & tempFilterItemArray(0) & " and QUESTION_TYPE = '" & _
                                    tempFilterItemArray(2) & "' order by ANSWER_TEXT"
                                Trace.Warn(Me.sqlCommonCommand.CommandText)
                            Else
                                Me.sqlCommonCommand.CommandText = "select fl.LIST_ITEM_ID as intList, fl.LIST_ITEM_VALUE as intListValue, fl.LIST_ITEM_TITLE as strLabel " & _
                                    "from " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM fl, " & myUtility.getExtension() & _
                                    "SAT_TEMPLATE_SURVEY fs where fs.SURVEY_KEY = " & Session("seqSurvey") & _
                                    " And fs.TEMPLATE_KEY = fl.TEMPLATE_KEY and fl.QUESTION_ID = " & tempFilterItemArray(0)
                                Trace.Warn(Me.sqlCommonCommand.CommandText)
                            End If
                        End If
                        If (Me.dsCommon.Tables.Contains("ListItems")) Then
                            Me.dsCommon.Tables("ListItems").Rows.Clear()
                        End If
                        Me.sqlCommonAdapter.Fill(Me.dsCommon, "ListItems")

                        'set the question filter again
                        If (tempFilterItemArray(6) = "MULTI") Then
                            setQuestionFilter(tempSkipList)
                        Else
                            tempSkipList = New ArrayList
                            setQuestionFilter(tempSkipList)
                        End If

                        Dim dateFilter As ArrayList
                        Dim runCount As Integer = 0
                        Dim blnDoOnce2 As Boolean = False
                        Dim row As DataRow
                        For Each dateFilter In Me.dateFilterArr
                            If (setDataFilter(dateFilter(2))) Then
                                Me.sqlCommonCommand.CommandText = "select frs.QUESTION_ID, frs.ANSWER_VALUE, count(frs.ANSWER_VALUE) as " & _
                                    "ValueCount, frs.ANSWER_TEXT, fq.CHART_TYPE_CODE, fq.QUESTION_GROUP_ID, fq.QUESTION_TEXT as QuestionText " & _
                                    "from " & myUtility.getExtension() & "SAT_RESPONSE_RESULT frs, " & myUtility.getExtension() & _
                                    "SAT_TEMPLATE_QUESTION fq where frs.SURVEY_KEY = " & Session("seqSurvey") & " " & _
                                    dateFilter(2).Replace("LAST_USED_DATE", "RESPONSE_DATE") & " " & Me.strDataFilter & " " & _
                                    Me.strQuestionFilter & " and frs.QUESTION_ID = fq.QUESTION_ID and frs.QUESTION_TYPE = fq.QUESTION_TYPE_CODE " & _
                                    "and fq.TEMPLATE_KEY = " & Session("seqTemplate") & " group by frs.QUESTION_ID, frs.ANSWER_VALUE, " & _
                                    "frs.ANSWER_TEXT, fq.CHART_TYPE_CODE, fq.QUESTION_GROUP_ID, fq.QUESTION_TEXT order by frs.QUESTION_ID, frs.ANSWER_VALUE, " & _
                                    "frs.ANSWER_TEXT, fq.CHART_TYPE_CODE, fq.QUESTION_GROUP_ID, fq.QUESTION_TEXT"
                                Trace.Warn(Me.sqlCommonCommand.CommandText)
                                If (Me.dsCommon.Tables.Contains("DataDump")) Then
                                    Me.dsCommon.Tables("DataDump").Rows.Clear()
                                End If
                                Me.sqlCommonAdapter.Fill(Me.dsCommon, "DataDump")

                                'get a count of questions in the group if dealing with a multi-chart set
                                Dim multiCount As Integer = 1
                                If (tempFilterItemArray(6) = "MULTI") Then
                                    multiCount = tempSkipList.Count()
                                End If

                                Dim sumAnswered(multiCount - 1) As Double

                                'build a table of percentages for the question using the list items
                                If (Me.dsCommon.Tables("ListItems").Rows.Count > 0) Then
                                    'build the skeletal structure
                                    Dim dt As DataTable = New DataTable
                                    dt.Columns.Add("ListLabel")
                                    dt.Columns("ListLabel").DataType = Me.dsCommon.Tables("ListItems").Columns("strLabel").DataType
                                    dt.Columns.Add("ListValue")
                                    dt.Columns("ListValue").DataType = Me.dsCommon.Tables("ListItems").Columns("intListValue").DataType

                                    If (tempFilterItemArray(6) = "MULTI") Then
                                        index = 1
                                        While (index <= multiCount)
                                            'generated for the count of questions in the group
                                            dt.Columns.Add("Percentage" & index)
                                            dt.Columns("Percentage" & index).DataType = System.Type.GetType("System.Decimal", True, True)
                                            dt.Columns.Add("AnswerCount" & index)
                                            If (tempFilterItemArray(2) = "D") Then
                                                dt.Columns("AnswerCount" & index).DataType = Me.dsCommon.Tables("ListItems").Columns("intListValue").DataType
                                            Else
                                                dt.Columns("AnswerCount" & index).DataType = Me.dsCommon.Tables("ListItems").Columns("intList").DataType
                                            End If
                                            index += 1
                                        End While
                                    Else
                                        'generated for the count of questions in the group
                                        dt.Columns.Add("Percentage")
                                        dt.Columns("Percentage").DataType = System.Type.GetType("System.Decimal", True, True)
                                        dt.Columns.Add("AnswerCount")
                                        If (tempFilterItemArray(2) = "D") Then
                                            dt.Columns("AnswerCount").DataType = Me.dsCommon.Tables("ListItems").Columns("intListValue").DataType
                                        Else
                                            dt.Columns("AnswerCount").DataType = Me.dsCommon.Tables("ListItems").Columns("intList").DataType
                                        End If
                                    End If

                                    'add default values and the item value as well as the label to the table
                                    If (tempFilterItemArray(6) = "MULTI") Then
                                        dt.TableName = "Question" & tempFilterItemArray(0)
                                        For Each row In Me.dsCommon.Tables("ListItems").Rows
                                            Dim tempRow As DataRow = dt.NewRow
                                            tempRow.Item(0) = CType(row.Item("strLabel"), String).Replace("~", ",")
                                            tempRow.Item(1) = row.Item("intListValue")

                                            index = 1
                                            While (index <= multiCount)
                                                tempRow.Item("Percentage" & index) = 0.0
                                                tempRow.Item("AnswerCount" & index) = 0
                                                index += 1
                                            End While
                                            dt.Rows.Add(tempRow)
                                        Next
                                    Else
                                        dt.TableName = "Question" & tempFilterItemArray(0)
                                        For Each row In Me.dsCommon.Tables("ListItems").Rows
                                            Dim tempRow As DataRow = dt.NewRow
                                            tempRow.Item(0) = CType(row.Item("strLabel"), String).Replace("~", ",")
                                            tempRow.Item(1) = row.Item("intListValue")
                                            tempRow.Item(2) = 0.0
                                            tempRow.Item(3) = 0
                                            dt.Rows.Add(tempRow)
                                        Next
                                    End If

                                    Dim blnDataExists As Boolean = False
                                    If (tempFilterItemArray(6) = "MULTI") Then
                                        'add the answer counts to the table
                                        Dim currentQuestion As Integer = tempFilterItemArray(0)
                                        Dim currentIndex As Integer = 1
                                        For Each dr In Me.dsCommon.Tables("DataDump").Rows
                                            If (dr.Item("QUESTION_ID") <> currentQuestion _
                                                And dr.Item("QUESTION_GROUP_ID") = tempFilterItemArray(5)) Then
                                                currentQuestion = dr.Item("QUESTION_ID")
                                                currentIndex += 1
                                            End If

                                            'get the list item label if not a "D" type
                                            If (tempFilterItemArray(2) <> "D") Then
                                                Me.sqlCommonCommand.CommandText = "Select LIST_ITEM_TITLE from " & myUtility.getExtension() & "SAT_QUESTION_LIST_ITEM " & _
                                                "where QUESTION_ID = " & currentQuestion & _
                                                " and TEMPLATE_KEY = " & Session("seqTemplate") & _
                                                " and LIST_ITEM_VALUE = " & dr.Item("ANSWER_VALUE")

                                                If (Me.dsCommon.Tables.Contains("ListLabel")) Then
                                                    Me.dsCommon.Tables("ListLabel").Rows.Clear()
                                                End If
                                                Me.sqlCommonAdapter.Fill(Me.dsCommon, "ListLabel")
                                            End If

                                            'populate the data table
                                            If (dr.Item("QUESTION_ID") = currentQuestion) Then
                                                blnDataExists = True
                                                Session("NonRTypeSelected") = True
                                                For Each row In dt.Rows()
                                                    If (tempFilterItemArray(2) = "D") Then
                                                        If (row.Item("ListLabel") = dr.Item("ANSWER_TEXT")) Then
                                                            row.Item("AnswerCount" & currentIndex) += dr.Item("ValueCount")
                                                            Exit For
                                                        End If
                                                    Else
                                                        If (Me.dsCommon.Tables("ListLabel").Rows.Count > 0) Then
                                                            If (row.Item("ListValue") = dr.Item("ANSWER_VALUE") _
                                                                And row.Item("ListLabel") = Me.dsCommon.Tables("ListLabel").Rows(0).Item("strLabel")) Then
                                                                row.Item("AnswerCount" & currentIndex) += dr.Item("ValueCount")
                                                                Exit For
                                                            End If
                                                        Else
                                                            If (row.Item("ListValue") = dr.Item("ANSWER_VALUE")) Then
                                                                row.Item("AnswerCount" & currentIndex) += dr.Item("ValueCount")
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If
                                                Next
                                                sumAnswered(currentIndex - 1) += dr.Item("ValueCount")
                                            ElseIf (currentIndex > multiCount) Then
                                                Exit For
                                            End If
                                        Next
                                    Else
                                        'add the answer counts to the table
                                        For Each dr In Me.dsCommon.Tables("DataDump").Rows
                                            If (dr.Item("QUESTION_ID") = tempFilterItemArray(0)) Then
                                                blnDataExists = True
                                                Session("NonRTypeSelected") = True
                                                For Each row In dt.Rows()
                                                    If (tempFilterItemArray(2) = "D") Then
                                                        If (row.Item("ListLabel") = dr.Item("ANSWER_TEXT")) Then
                                                            row.Item("AnswerCount") += dr.Item("ValueCount")
                                                        End If
                                                    Else
                                                        If (row.Item("ListValue") = dr.Item("ANSWER_VALUE")) Then
                                                            row.Item("AnswerCount") += dr.Item("ValueCount")
                                                        End If
                                                    End If
                                                Next
                                                sumAnswered(0) += dr.Item("ValueCount")
                                            ElseIf (dr.Item("QUESTION_ID") > tempFilterItemArray(0)) Then
                                                Exit For
                                            End If
                                        Next
                                    End If

                                    If (blnDataExists) Then
                                        If (tempFilterItemArray(6) = "MULTI") Then
                                            'calculate the percentages for each row in the table
                                            index = 1
                                            While (index <= multiCount)
                                                For Each row In dt.Rows
                                                    If (sumAnswered(index - 1) = 0) Then
                                                        row.Item("Percentage" & index) = 0
                                                    Else
                                                        row.Item("Percentage" & index) = (row.Item("AnswerCount" & index) / sumAnswered(index - 1)) * 100
                                                    End If
                                                Next
                                                index += 1
                                            End While
                                        Else
                                            'calculate the percentages for each row in the table
                                            For Each row In dt.Rows
                                                If (sumAnswered(0) = 0) Then
                                                    row.Item("Percentage") = 0
                                                Else
                                                    row.Item("Percentage") = (row.Item("AnswerCount") / sumAnswered(0)) * 100
                                                End If
                                            Next
                                        End If

                                        If (Not blnDoOnce3) Then
                                            'section header
                                            genericRow = New System.Web.UI.WebControls.TableRow
                                            genericCell = New System.Web.UI.WebControls.TableCell
                                            genericCell.Wrap = False
                                            genericCell.CssClass = "sectionHeadTeal"
                                            genericCell.Text = "Other Results"
                                            genericCell.EnableViewState = True
                                            genericCell.ColumnSpan = 2
                                            genericRow.Cells.Add(genericCell)
                                            genericRow.EnableViewState = True
                                            Me.DataTable2.Rows.Add(genericRow)
                                            'genericRow = New System.Web.UI.WebControls.TableRow
                                            'genericCell = New System.Web.UI.WebControls.TableCell
                                            'genericCell.Wrap = False
                                            'genericCell.Text = "<HR style='BORDER-TOP-STYLE: ridge' width='100%' SIZE='3'>"
                                            'genericCell.ColumnSpan = 2
                                            'genericRow.Cells.Add(genericCell)
                                            'Me.DataTable2.Rows.Add(genericRow)
                                            blnDoOnce3 = True
                                        End If

                                        'buffer zone
                                        'genericRow = New System.Web.UI.WebControls.TableRow
                                        'genericCell = New System.Web.UI.WebControls.TableCell
                                        'genericCell.Wrap = False
                                        'genericCell.Height = Unit.Pixel(40)
                                        'genericCell.EnableViewState = True
                                        'genericCell.ColumnSpan = 2
                                        'genericRow.Cells.Add(genericCell)
                                        'genericRow.EnableViewState = True
                                        'Me.DataTable2.Rows.Add(genericRow)

                                        If (Session("SplitOutput") <> "NONE") Then
                                            If (tempFilterItemArray(6) = "MULTI") Then
                                                'get the rest of the rows if missing some
                                                index = 0
                                                While (index < sumAnswered.Length)
                                                    'generate a row for the question text
                                                    genericRow = New System.Web.UI.WebControls.TableRow
                                                    genericCell = New System.Web.UI.WebControls.TableCell
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "boldContentRed"
                                                    genericLabel.Text = index + 1 & ". "
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "reportContent"
                                                    genericLabel.Text = tempSkipListText(index).replace("~", ",")
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "boldContentTeal"
                                                    genericLabel.Text = "(" & sumAnswered(index) & ")"
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.Wrap = True
                                                    genericCell.EnableViewState = True
                                                    genericCell.ColumnSpan = 2
                                                    genericRow.Cells.Add(genericCell)
                                                    genericRow.EnableViewState = True
                                                    If (Session("SplitOutput") <> "NONE") Then
                                                        If (blnDoOnce2) Then
                                                            genericRow.BackColor = Color.LightSteelBlue
                                                        Else
                                                            genericRow.BackColor = Color.LemonChiffon
                                                        End If
                                                    End If
                                                    Me.DataTable2.Rows.Add(genericRow)

                                                    index += 1
                                                End While
                                                blnDoOnce2 = True
                                            Else
                                                'generate a row for the question text
                                                genericRow = New System.Web.UI.WebControls.TableRow
                                                genericCell = New System.Web.UI.WebControls.TableCell
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "reportContent"
                                                genericLabel.Text = CType(tempFilterItemArray(3), String).Replace("~", ",")
                                                genericCell.Controls.Add(genericLabel)
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "boldContentTeal"
                                                genericLabel.Text = "(" & sumAnswered(0) & ")"
                                                genericCell.Controls.Add(genericLabel)
                                                genericCell.Wrap = True
                                                genericCell.EnableViewState = True
                                                genericCell.ColumnSpan = 2
                                                genericRow.Cells.Add(genericCell)
                                                genericRow.EnableViewState = True
                                                If (Session("SplitOutput") <> "NONE") Then
                                                    If (blnDoOnce2) Then
                                                        genericRow.BackColor = Color.LightSteelBlue
                                                    Else
                                                        genericRow.BackColor = Color.LemonChiffon
                                                    End If
                                                End If
                                                Me.DataTable2.Rows.Add(genericRow)
                                                blnDoOnce2 = True
                                            End If
                                        ElseIf (Session("SplitOutput") = "NONE") Then
                                            If (tempFilterItemArray(6) = "MULTI") Then
                                                Dim questionText As String = ""
                                                index = 0
                                                Dim currentQuestion As Integer = 0
                                                For Each row In Me.dsCommon.Tables("DataDump").Rows
                                                    If (row.Item("QUESTION_GROUP_ID") = tempFilterItemArray(5) _
                                                        And tempSkipList.Contains(CType(row.Item("QUESTION_ID"), String))) Then
                                                        If (questionText = "") Then
                                                            questionText = row.Item("QuestionText")
                                                            currentQuestion = row.Item("QUESTION_ID")

                                                            'generate a row for the question text
                                                            genericRow = New System.Web.UI.WebControls.TableRow
                                                            genericCell = New System.Web.UI.WebControls.TableCell
                                                            genericLabel = New System.Web.UI.WebControls.Label
                                                            genericLabel.CssClass = "boldContentRed"
                                                            genericLabel.Text = index + 1 & ". "
                                                            genericCell.Controls.Add(genericLabel)
                                                            genericLabel = New System.Web.UI.WebControls.Label
                                                            genericLabel.CssClass = "reportContent"
                                                            genericLabel.Text = questionText.Replace("~", ",")
                                                            genericCell.Controls.Add(genericLabel)
                                                            genericCell.Wrap = True
                                                            genericCell.EnableViewState = True
                                                            genericRow.Cells.Add(genericCell)
                                                            genericLabel = New System.Web.UI.WebControls.Label
                                                            genericLabel.CssClass = "boldContentTeal"
                                                            genericLabel.Text = "(" & sumAnswered(index) & ")"
                                                            genericCell.Controls.Add(genericLabel)
                                                            genericCell.Wrap = True
                                                            genericCell.EnableViewState = True
                                                            genericCell.ColumnSpan = 2
                                                            genericRow.Cells.Add(genericCell)
                                                            genericRow.EnableViewState = True
                                                            Me.DataTable2.Rows.Add(genericRow)
                                                        ElseIf (currentQuestion <> row.Item("QUESTION_ID")) Then
                                                            questionText = row.Item("QuestionText")
                                                            currentQuestion = row.Item("QUESTION_ID")
                                                            index += 1

                                                            'generate a row for the question text
                                                            genericRow = New System.Web.UI.WebControls.TableRow
                                                            genericCell = New System.Web.UI.WebControls.TableCell
                                                            genericLabel = New System.Web.UI.WebControls.Label
                                                            genericLabel.CssClass = "boldContentRed"
                                                            genericLabel.Text = index + 1 & ". "
                                                            genericCell.Controls.Add(genericLabel)
                                                            genericLabel = New System.Web.UI.WebControls.Label
                                                            genericLabel.CssClass = "reportContent"
                                                            genericLabel.Text = questionText.Replace("~", ",")
                                                            genericCell.Controls.Add(genericLabel)
                                                            genericCell.Wrap = True
                                                            genericCell.EnableViewState = True
                                                            genericRow.Cells.Add(genericCell)
                                                            genericLabel = New System.Web.UI.WebControls.Label
                                                            genericLabel.CssClass = "boldContentTeal"
                                                            genericLabel.Text = "(" & sumAnswered(index) & ")"
                                                            genericCell.Controls.Add(genericLabel)
                                                            genericCell.Wrap = True
                                                            genericCell.EnableViewState = True
                                                            genericCell.ColumnSpan = 2
                                                            genericRow.Cells.Add(genericCell)
                                                            genericRow.EnableViewState = True
                                                            Me.DataTable2.Rows.Add(genericRow)
                                                        End If
                                                    End If
                                                Next
                                            Else
                                                'generate a row for the question text
                                                genericRow = New System.Web.UI.WebControls.TableRow
                                                genericCell = New System.Web.UI.WebControls.TableCell
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "reportContent"
                                                genericLabel.Text = CType(tempFilterItemArray(3), String).Replace("~", ",")
                                                genericCell.Controls.Add(genericLabel)
                                                genericCell.Wrap = True
                                                genericCell.EnableViewState = True
                                                genericRow.Cells.Add(genericCell)
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "boldContentTeal"
                                                genericLabel.Text = "(" & sumAnswered(0) & ")"
                                                genericCell.Controls.Add(genericLabel)
                                                genericCell.Wrap = True
                                                genericCell.EnableViewState = True
                                                genericCell.ColumnSpan = 2
                                                genericRow.Cells.Add(genericCell)
                                                genericRow.EnableViewState = True
                                                Me.DataTable2.Rows.Add(genericRow)
                                            End If
                                        End If

                                        If (Session("SplitOutput") <> "NONE") Then
                                            'generate a row for the date range
                                            genericRow = New System.Web.UI.WebControls.TableRow
                                            genericCell = New System.Web.UI.WebControls.TableCell
                                            genericLabel = New System.web.ui.webcontrols.Label
                                            genericLabel.CssClass = "ReportHeaderData"
                                            genericLabel.Text = "Period of Record Set: "
                                            genericCell.Controls.Add(genericLabel)
                                            genericLabel = New System.web.ui.webcontrols.Label
                                            genericLabel.CssClass = "ReportHeaderBlue"
                                            genericLabel.Text = dateFilter(0) & " - " & dateFilter(1)
                                            genericCell.Controls.Add(genericLabel)
                                            genericCell.Wrap = True
                                            genericCell.EnableViewState = True
                                            genericCell.ColumnSpan = 2
                                            genericRow.Cells.Add(genericCell)
                                            genericRow.EnableViewState = True
                                            genericRow.BackColor = Color.LightGray
                                            Me.DataTable2.Rows.Add(genericRow)
                                        End If

                                        'generate a row for the question answer
                                        genericTable = New System.Web.UI.WebControls.Table
                                        Dim colorIndex As Integer = 0
                                        blnDoOnce3 = False
                                        For Each row In dt.Rows
                                            'get the color for the current row
                                            Dim strCurrentColor As String = "FFFFFF"
                                            If (Me.dsCommon.Tables("Colors").Rows.Count() > 0) Then
                                                strCurrentColor = Me.dsCommon.Tables("Colors").Rows(colorIndex).Item("CHART_COLOR_CODE")
                                                colorIndex += 1
                                                If (colorIndex = Me.dsCommon.Tables("Colors").Rows.Count()) Then
                                                    colorIndex = 0
                                                End If
                                            End If

                                            'set up the color for the row
                                            Dim Alpha As Integer = 0
                                            Dim Red As Integer = 0
                                            Dim Green As Integer = 0
                                            Dim Blue As Integer = 0
                                            If (strCurrentColor.Length = 8) Then
                                                Alpha = Integer.Parse(strCurrentColor.Substring(0, 2), Globalization.NumberStyles.HexNumber)
                                                Red = Integer.Parse(strCurrentColor.Substring(2, 2), Globalization.NumberStyles.HexNumber)
                                                Green = Integer.Parse(strCurrentColor.Substring(4, 2), Globalization.NumberStyles.HexNumber)
                                                Blue = Integer.Parse(strCurrentColor.Substring(6, 2), Globalization.NumberStyles.HexNumber)
                                            Else
                                                Red = Integer.Parse(strCurrentColor.Substring(0, 2), Globalization.NumberStyles.HexNumber)
                                                Green = Integer.Parse(strCurrentColor.Substring(2, 2), Globalization.NumberStyles.HexNumber)
                                                Blue = Integer.Parse(strCurrentColor.Substring(4, 2), Globalization.NumberStyles.HexNumber)
                                            End If

                                            'table header setup - do once
                                            If Not (blnDoOnce3) Then
                                                genericRow = New System.Web.UI.WebControls.TableRow
                                                genericCell = New System.Web.UI.WebControls.TableCell
                                                genericCell.Width = Unit.Pixel(20)
                                                genericCell.Wrap = False
                                                genericCell.EnableViewState = True
                                                genericRow.Cells.Add(genericCell)
                                                genericCell = New System.Web.UI.WebControls.TableCell
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "reportContent"
                                                genericLabel.Text = "Answer Text"
                                                genericCell.Controls.Add(genericLabel)
                                                genericCell.Wrap = True
                                                genericCell.EnableViewState = True
                                                genericCell.BackColor = System.Drawing.Color.AntiqueWhite
                                                genericRow.Cells.Add(genericCell)

                                                If (tempFilterItemArray(6) = "MULTI") Then
                                                    index = 1
                                                    While (index <= multiCount)
                                                        genericCell = New System.Web.UI.WebControls.TableCell
                                                        genericLabel = New System.web.ui.webcontrols.Label
                                                        genericLabel.CssClass = "reportContent"
                                                        genericLabel.Text = index
                                                        genericCell.Controls.Add(genericLabel)
                                                        genericCell.HorizontalAlign = HorizontalAlign.Center
                                                        genericCell.Wrap = False
                                                        genericCell.EnableViewState = True
                                                        genericCell.BackColor = System.Drawing.Color.AntiqueWhite
                                                        genericRow.Cells.Add(genericCell)
                                                        index += 1
                                                    End While
                                                Else
                                                    genericCell = New System.Web.UI.WebControls.TableCell
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "reportContent"

                                                    'change the header based on percent or count display
                                                    If (blnPercentDisplay) Then
                                                        genericLabel.Text = "Percent"
                                                    Else
                                                        genericLabel.Text = "Count"
                                                    End If

                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.HorizontalAlign = HorizontalAlign.Center
                                                    genericCell.Wrap = False
                                                    genericCell.EnableViewState = True
                                                    genericCell.BackColor = System.Drawing.Color.AntiqueWhite
                                                    genericRow.Cells.Add(genericCell)
                                                End If

                                                genericRow.EnableViewState = True
                                                genericTable.Rows.Add(genericRow)
                                                blnDoOnce3 = True
                                            End If

                                            'populate row data
                                            genericRow = New System.Web.UI.WebControls.TableRow
                                            genericCell = New System.Web.UI.WebControls.TableCell
                                            genericCell.Width = Unit.Pixel(20)
                                            genericCell.Wrap = False
                                            genericCell.EnableViewState = True
                                            genericRow.Cells.Add(genericCell)
                                            genericCell = New System.Web.UI.WebControls.TableCell
                                            genericLabel = New System.web.ui.webcontrols.Label
                                            genericLabel.CssClass = "reportContent"
                                            genericLabel.Text = row.Item("ListLabel")
                                            genericCell.Wrap = True
                                            genericCell.EnableViewState = True
                                            genericCell.BackColor = System.Drawing.Color.FromArgb(Alpha, Red, Green, Blue)
                                            genericCell.Controls.Add(genericLabel)
                                            genericRow.Cells.Add(genericCell)

                                            If (tempFilterItemArray(6) = "MULTI") Then
                                                'format the value based on percent or count display
                                                index = 1
                                                While (index <= multiCount)
                                                    genericCell = New System.Web.UI.WebControls.TableCell
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "reportContent"

                                                    If (blnPercentDisplay) Then
                                                        genericLabel.Text = CType(row.Item("Percentage" & index), Double).ToString("#0.##")
                                                    Else
                                                        genericLabel.Text = row.Item("AnswerCount" & index)
                                                    End If
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.HorizontalAlign = HorizontalAlign.Right
                                                    genericCell.Wrap = False
                                                    genericCell.EnableViewState = True
                                                    genericCell.BackColor = System.Drawing.Color.FromArgb(Alpha, Red, Green, Blue)
                                                    genericRow.Cells.Add(genericCell)
                                                    index += 1
                                                End While
                                            Else
                                                genericCell = New System.Web.UI.WebControls.TableCell
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "reportContent"

                                                'format the value based on percent or count display
                                                If (blnPercentDisplay) Then
                                                    genericLabel.Text = CType(row.Item("Percentage"), Double).ToString("#0.##")
                                                Else
                                                    genericLabel.Text = row.Item("AnswerCount")
                                                End If
                                                genericCell.Controls.Add(genericLabel)
                                                genericCell.HorizontalAlign = HorizontalAlign.Right
                                                genericCell.Wrap = False
                                                genericCell.EnableViewState = True
                                                genericCell.BackColor = System.Drawing.Color.FromArgb(Alpha, Red, Green, Blue)
                                                genericRow.Cells.Add(genericCell)
                                            End If

                                            genericRow.EnableViewState = True
                                            genericTable.Rows.Add(genericRow)
                                        Next

                                        genericRow = New System.Web.UI.WebControls.TableRow
                                        genericCell = New System.Web.UI.WebControls.TableCell
                                        genericCell.Controls.Add(genericTable)
                                        genericCell.Wrap = False
                                        genericCell.EnableViewState = True
                                        genericCell.VerticalAlign = VerticalAlign.Top
                                        genericRow.Cells.Add(genericCell)

                                        'generate the chart for the Rating question
                                        genericCell = New System.Web.UI.WebControls.TableCell
                                        genericChart = New Infragistics.webui.UltraWebChart.UltraChart
                                        If Not (generateChart(genericChart, tempFilterItemArray(0), tempFilterItemArray(6), dateFilter(2), dt, tempSkipList, runCount, False, False, IIf(tempFilterItemArray(2) = "D", True, False), blnPercentDisplay)) Then
                                            Session("Alert") = "Chart generation failure."
                                            Return False
                                        End If
                                        genericCell.Controls.Add(genericChart)
                                        genericCell.Wrap = False
                                        genericCell.EnableViewState = True
                                        genericCell.HorizontalAlign = HorizontalAlign.Right
                                        genericCell.VerticalAlign = VerticalAlign.Top
                                        genericRow.Cells.Add(genericCell)

                                        'Set to put a line break after 3 question results appear on the page.
                                        If (questionCounter Mod Session("QuestionsPerPage") = 0) Then
                                            genericRow.CssClass = "Break"
                                        End If
                                        questionCounter += 1

                                        'finalize the row
                                        genericRow.EnableViewState = True
                                        Me.DataTable2.Rows.Add(genericRow)
                                        'genericRow = New System.Web.UI.WebControls.TableRow
                                        'genericCell = New System.Web.UI.WebControls.TableCell
                                        'genericCell.Height = Unit.Pixel(40)
                                        'genericCell.Wrap = False
                                        'genericCell.EnableViewState = True
                                        'genericRow.Cells.Add(genericCell)
                                        'genericRow.EnableViewState = True
                                        'Me.DataTable2.Rows.Add(genericRow)
                                        Me.DataTable2.EnableViewState = True
                                        runCount += 1
                                    Else
                                        If (Not blnDoOnce3) Then
                                            'section header
                                            genericRow = New System.Web.UI.WebControls.TableRow
                                            genericCell = New System.Web.UI.WebControls.TableCell
                                            genericCell.Wrap = False
                                            genericCell.CssClass = "sectionHeadTeal"
                                            genericCell.Text = "Other Results"
                                            genericCell.EnableViewState = True
                                            genericCell.ColumnSpan = 2
                                            genericRow.Cells.Add(genericCell)
                                            genericRow.EnableViewState = True
                                            noDataRows.Add(genericRow)
                                            'Me.DataTable2.Rows.Add(genericRow)
                                            blnDoOnce3 = True
                                        End If

                                        'buffer zone
                                        'genericRow = New System.Web.UI.WebControls.TableRow
                                        'genericCell = New System.Web.UI.WebControls.TableCell
                                        'genericCell.Wrap = False
                                        'genericCell.Height = Unit.Pixel(40)
                                        'genericCell.EnableViewState = True
                                        'genericCell.ColumnSpan = 2
                                        'genericRow.Cells.Add(genericCell)
                                        'genericRow.EnableViewState = True
                                        'Me.DataTable2.Rows.Add(genericRow)

                                        If (Session("SplitOutput") <> "NONE") Then
                                            If (tempFilterItemArray(6) = "MULTI") Then
                                                index = 0
                                                While (index < skipListText.Count)
                                                    'generate a row for the question text
                                                    genericRow = New System.Web.UI.WebControls.TableRow
                                                    genericCell = New System.Web.UI.WebControls.TableCell
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "reportContent"
                                                    genericLabel.Text = CType(tempSkipListText(index), String).Replace("~", ",")
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "boldContentTeal"
                                                    genericLabel.Text = "(" & sumAnswered(index) & ")"
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.Wrap = True
                                                    genericCell.EnableViewState = True
                                                    genericCell.ColumnSpan = 2
                                                    genericRow.Cells.Add(genericCell)
                                                    genericRow.EnableViewState = True
                                                    If (Session("SplitOutput") <> "NONE") Then
                                                        If (blnDoOnce2) Then
                                                            genericRow.BackColor = Color.LightSteelBlue
                                                        Else
                                                            genericRow.BackColor = Color.LemonChiffon
                                                        End If
                                                    End If
                                                    noDataRows.Add(genericRow)
                                                    'Me.DataTable2.Rows.Add(genericRow)
                                                    blnDoOnce2 = True
                                                    index += 1
                                                End While
                                            Else
                                                'generate a row for the question text
                                                genericRow = New System.Web.UI.WebControls.TableRow
                                                genericCell = New System.Web.UI.WebControls.TableCell
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "reportContent"
                                                genericLabel.Text = CType(tempFilterItemArray(3), String).Replace("~", ",")
                                                genericCell.Controls.Add(genericLabel)
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "boldContentTeal"
                                                genericLabel.Text = "(" & sumAnswered(index) & ")"
                                                genericCell.Controls.Add(genericLabel)
                                                genericCell.Wrap = True
                                                genericCell.EnableViewState = True
                                                genericCell.ColumnSpan = 2
                                                genericRow.Cells.Add(genericCell)
                                                genericRow.EnableViewState = True
                                                If (Session("SplitOutput") <> "NONE") Then
                                                    If (blnDoOnce2) Then
                                                        genericRow.BackColor = Color.LightSteelBlue
                                                    Else
                                                        genericRow.BackColor = Color.LemonChiffon
                                                    End If
                                                End If
                                                noDataRows.Add(genericRow)
                                                'Me.DataTable2.Rows.Add(genericRow)
                                                blnDoOnce2 = True
                                            End If
                                        ElseIf (Session("SplitOutput") = "NONE") Then
                                            If (tempFilterItemArray(6) = "MULTI") Then
                                                index = 0
                                                While (index < skipListText.Count)
                                                    'generate a row for the question text
                                                    genericRow = New System.Web.UI.WebControls.TableRow
                                                    genericCell = New System.Web.UI.WebControls.TableCell
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "reportContent"
                                                    genericLabel.Text = CType(tempSkipListText(index), String).Replace("~", ",")
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.Wrap = True
                                                    genericCell.EnableViewState = True
                                                    genericRow.Cells.Add(genericCell)
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "boldContentTeal"
                                                    genericLabel.Text = "(" & sumAnswered(index) & ")"
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.Wrap = True
                                                    genericCell.EnableViewState = True
                                                    genericRow.Cells.Add(genericCell)
                                                    genericRow.EnableViewState = True
                                                    noDataRows.Add(genericRow)
                                                    'Me.DataTable2.Rows.Add(genericRow)
                                                    index += 1
                                                End While
                                            Else
                                                'generate a row for the question text
                                                genericRow = New System.Web.UI.WebControls.TableRow
                                                genericCell = New System.Web.UI.WebControls.TableCell
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "reportContent"
                                                genericLabel.Text = CType(tempFilterItemArray(3), String).Replace("~", ",")
                                                genericCell.Controls.Add(genericLabel)
                                                genericCell.Wrap = True
                                                genericCell.EnableViewState = True
                                                genericRow.Cells.Add(genericCell)
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "boldContentTeal"
                                                index = 0
                                                Dim wholeAnswer As Integer = 0
                                                While (index < sumAnswered.Length())
                                                    wholeAnswer += sumAnswered(index)
                                                    index += 1
                                                End While
                                                genericLabel.Text = "(" & wholeAnswer & ")"
                                                genericCell.Controls.Add(genericLabel)
                                                genericCell.Wrap = True
                                                genericCell.EnableViewState = True
                                                genericRow.Cells.Add(genericCell)
                                                genericRow.EnableViewState = True
                                                noDataRows.Add(genericRow)
                                                'Me.DataTable2.Rows.Add(genericRow)
                                            End If
                                        End If

                                        If (Session("SplitOutput") <> "NONE") Then
                                            'generate a row for the date range
                                            genericRow = New System.Web.UI.WebControls.TableRow
                                            genericCell = New System.Web.UI.WebControls.TableCell
                                            genericLabel = New System.web.ui.webcontrols.Label
                                            genericLabel.CssClass = "ReportHeaderData"
                                            genericLabel.Text = "Period of Record Set: "
                                            genericCell.Controls.Add(genericLabel)
                                            genericLabel = New System.web.ui.webcontrols.Label
                                            genericLabel.CssClass = "ReportHeaderBlue"
                                            genericLabel.Text = dateFilter(0) & " - " & dateFilter(1)
                                            genericCell.Controls.Add(genericLabel)
                                            genericCell.Wrap = True
                                            genericCell.EnableViewState = True
                                            genericCell.ColumnSpan = 2
                                            genericRow.Cells.Add(genericCell)
                                            genericRow.EnableViewState = True
                                            genericRow.BackColor = Color.LightGray
                                            noDataRows.Add(genericRow)
                                            'Me.DataTable2.Rows.Add(genericRow)
                                        End If

                                        'generate to inform user that there is no data for this date range
                                        genericRow = New System.Web.UI.WebControls.TableRow
                                        genericTable = New System.Web.UI.WebControls.Table
                                        genericCell = New System.Web.UI.WebControls.TableCell
                                        genericCell.Width = Unit.Pixel(20)
                                        genericCell.Wrap = False
                                        genericCell.EnableViewState = True
                                        genericRow.Cells.Add(genericCell)
                                        genericCell = New System.Web.UI.WebControls.TableCell
                                        genericLabel = New System.web.ui.webcontrols.Label
                                        genericLabel.CssClass = "NoDataRed"
                                        genericLabel.Text = "This question has no data for this time period."
                                        genericCell.Controls.Add(genericLabel)
                                        'genericCell.ColumnSpan = 2
                                        'genericCell.HorizontalAlign = HorizontalAlign.Center
                                        genericRow.Cells.Add(genericCell)
                                        genericRow.EnableViewState = True
                                        genericTable.Rows.Add(genericRow)
                                        genericRow = New System.Web.UI.WebControls.TableRow
                                        genericCell = New System.Web.UI.WebControls.TableCell
                                        genericCell.ColumnSpan = 2
                                        genericCell.Controls.Add(genericTable)
                                        genericRow.Cells.Add(genericCell)
                                        noDataRows.Add(genericRow)
                                        'Me.DataTable2.Rows.Add(genericRow)
                                    End If
                                Else
                                    If (tempFilterItemArray(2) = "D") Then
                                        If (Not blnDoOnce3) Then
                                            'section header
                                            genericRow = New System.Web.UI.WebControls.TableRow
                                            genericCell = New System.Web.UI.WebControls.TableCell
                                            genericCell.Wrap = False
                                            genericCell.CssClass = "sectionHeadTeal"
                                            genericCell.Text = "Other Results"
                                            genericCell.EnableViewState = True
                                            genericCell.ColumnSpan = 2
                                            genericRow.Cells.Add(genericCell)
                                            genericRow.EnableViewState = True
                                            noDataRows.Add(genericRow)
                                            'Me.DataTable2.Rows.Add(genericRow)
                                            blnDoOnce3 = True
                                        End If

                                        'buffer zone
                                        'genericRow = New System.Web.UI.WebControls.TableRow
                                        'genericCell = New System.Web.UI.WebControls.TableCell
                                        'genericCell.Wrap = False
                                        'genericCell.Height = Unit.Pixel(40)
                                        'genericCell.EnableViewState = True
                                        'genericCell.ColumnSpan = 2
                                        'genericRow.Cells.Add(genericCell)
                                        'genericRow.EnableViewState = True
                                        'Me.DataTable2.Rows.Add(genericRow)

                                        If (Session("SplitOutput") <> "NONE") Then
                                            If (tempFilterItemArray(6) = "MULTI") Then
                                                index = 0
                                                While (index < skipListText.Count)
                                                    'generate a row for the question text
                                                    genericRow = New System.Web.UI.WebControls.TableRow
                                                    genericCell = New System.Web.UI.WebControls.TableCell
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "reportContent"
                                                    genericLabel.Text = CType(tempSkipListText(index), String).Replace("~", ",")
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "boldContentTeal"
                                                    genericLabel.Text = "(" & sumAnswered(index) & ")"
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.Wrap = True
                                                    genericCell.EnableViewState = True
                                                    genericCell.ColumnSpan = 2
                                                    genericRow.Cells.Add(genericCell)
                                                    genericRow.EnableViewState = True
                                                    If (Session("SplitOutput") <> "NONE") Then
                                                        If (blnDoOnce2) Then
                                                            genericRow.BackColor = Color.LightSteelBlue
                                                        Else
                                                            genericRow.BackColor = Color.LemonChiffon
                                                        End If
                                                    End If
                                                    noDataRows.Add(genericRow)
                                                    'Me.DataTable2.Rows.Add(genericRow)
                                                    blnDoOnce2 = True
                                                    index += 1
                                                End While
                                            Else
                                                'generate a row for the question text
                                                genericRow = New System.Web.UI.WebControls.TableRow
                                                genericCell = New System.Web.UI.WebControls.TableCell
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "reportContent"
                                                genericLabel.Text = CType(tempFilterItemArray(3), String).Replace("~", ",")
                                                genericCell.Controls.Add(genericLabel)
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "boldContentTeal"
                                                genericLabel.Text = "(" & sumAnswered(index) & ")"
                                                genericCell.Controls.Add(genericLabel)
                                                genericCell.Wrap = True
                                                genericCell.EnableViewState = True
                                                genericCell.ColumnSpan = 2
                                                genericRow.Cells.Add(genericCell)
                                                genericRow.EnableViewState = True
                                                If (Session("SplitOutput") <> "NONE") Then
                                                    If (blnDoOnce2) Then
                                                        genericRow.BackColor = Color.LightSteelBlue
                                                    Else
                                                        genericRow.BackColor = Color.LemonChiffon
                                                    End If
                                                End If
                                                noDataRows.Add(genericRow)
                                                'Me.DataTable2.Rows.Add(genericRow)
                                                blnDoOnce2 = True
                                            End If
                                        ElseIf (Session("SplitOutput") = "NONE") Then
                                            If (tempFilterItemArray(6) = "MULTI") Then
                                                index = 0
                                                While (index < skipListText.Count)
                                                    'generate a row for the question text
                                                    genericRow = New System.Web.UI.WebControls.TableRow
                                                    genericCell = New System.Web.UI.WebControls.TableCell
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "reportContent"
                                                    genericLabel.Text = CType(tempSkipListText(index), String).Replace("~", ",")
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.Wrap = True
                                                    genericCell.EnableViewState = True
                                                    genericRow.Cells.Add(genericCell)
                                                    genericLabel = New System.web.ui.webcontrols.Label
                                                    genericLabel.CssClass = "boldContentTeal"
                                                    genericLabel.Text = "(" & sumAnswered(index) & ")"
                                                    genericCell.Controls.Add(genericLabel)
                                                    genericCell.Wrap = True
                                                    genericCell.EnableViewState = True
                                                    genericRow.Cells.Add(genericCell)
                                                    genericRow.EnableViewState = True
                                                    noDataRows.Add(genericRow)
                                                    'Me.DataTable2.Rows.Add(genericRow)
                                                    index += 1
                                                End While
                                            Else
                                                'generate a row for the question text
                                                genericRow = New System.Web.UI.WebControls.TableRow
                                                genericCell = New System.Web.UI.WebControls.TableCell
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "reportContent"
                                                genericLabel.Text = CType(tempFilterItemArray(3), String).Replace("~", ",")
                                                genericCell.Controls.Add(genericLabel)
                                                genericCell.Wrap = True
                                                genericCell.EnableViewState = True
                                                genericRow.Cells.Add(genericCell)
                                                genericLabel = New System.web.ui.webcontrols.Label
                                                genericLabel.CssClass = "boldContentTeal"
                                                index = 0
                                                Dim wholeAnswer As Integer = 0
                                                While (index < sumAnswered.Length())
                                                    wholeAnswer += sumAnswered(index)
                                                    index += 1
                                                End While
                                                genericLabel.Text = "(" & wholeAnswer & ")"
                                                genericCell.Controls.Add(genericLabel)
                                                genericCell.Wrap = True
                                                genericCell.EnableViewState = True
                                                genericRow.Cells.Add(genericCell)
                                                genericRow.EnableViewState = True
                                                noDataRows.Add(genericRow)
                                                'Me.DataTable2.Rows.Add(genericRow)
                                            End If
                                        End If

                                        If (Session("SplitOutput") <> "NONE") Then
                                            'generate a row for the date range
                                            genericRow = New System.Web.UI.WebControls.TableRow
                                            genericCell = New System.Web.UI.WebControls.TableCell
                                            genericLabel = New System.web.ui.webcontrols.Label
                                            genericLabel.CssClass = "ReportHeaderData"
                                            genericLabel.Text = "Period of Record Set: "
                                            genericCell.Controls.Add(genericLabel)
                                            genericLabel = New System.web.ui.webcontrols.Label
                                            genericLabel.CssClass = "ReportHeaderBlue"
                                            genericLabel.Text = dateFilter(0) & " - " & dateFilter(1)
                                            genericCell.Controls.Add(genericLabel)
                                            genericCell.Wrap = True
                                            genericCell.EnableViewState = True
                                            genericCell.ColumnSpan = 2
                                            genericRow.Cells.Add(genericCell)
                                            genericRow.EnableViewState = True
                                            genericRow.BackColor = Color.LightGray
                                            noDataRows.Add(genericRow)
                                            'Me.DataTable2.Rows.Add(genericRow)
                                        End If

                                        'generate to inform user that there is no data for this date range
                                        genericRow = New System.Web.UI.WebControls.TableRow
                                        genericTable = New System.Web.UI.WebControls.Table
                                        genericCell = New System.Web.UI.WebControls.TableCell
                                        genericCell.Width = Unit.Pixel(20)
                                        genericCell.Wrap = False
                                        genericCell.EnableViewState = True
                                        genericRow.Cells.Add(genericCell)
                                        genericCell = New System.Web.UI.WebControls.TableCell
                                        genericLabel = New System.web.ui.webcontrols.Label
                                        genericLabel.CssClass = "NoDataRed"
                                        genericLabel.Text = "This question has no data for this time period."
                                        genericCell.Controls.Add(genericLabel)
                                        'genericCell.ColumnSpan = 2
                                        'genericCell.HorizontalAlign = HorizontalAlign.Center
                                        genericRow.Cells.Add(genericCell)
                                        genericRow.EnableViewState = True
                                        genericTable.Rows.Add(genericRow)
                                        genericRow = New System.Web.UI.WebControls.TableRow
                                        genericCell = New System.Web.UI.WebControls.TableCell
                                        genericCell.ColumnSpan = 2
                                        genericCell.Controls.Add(genericTable)
                                        genericRow.Cells.Add(genericCell)
                                        noDataRows.Add(genericRow)
                                        'Me.DataTable2.Rows.Add(genericRow)
                                        'Else
                                        'Session("Alert") = "Failed to get survey question list item data for others."
                                        'Return False
                                    End If
                                End If
                            Else
                                Session("Alert") = "Failed to get survey question list item data for others."
                                Return False
                            End If
                        Next
                    End If
                End If
            Next

            'get the questions that have no data
            Dim tableRow As System.Web.UI.WebControls.TableRow
            For Each tableRow In noDataRows
                Me.DataTable2.Rows.Add(tableRow)
            Next
            Return True
        Catch ex As Exception
            Session("Alert") = "Failed to get survey data.  " & ex.ToString
            Return False
        End Try
    End Function

    'processes the text results
    Private Function processTextData() As Boolean
        Trace.Warn("processing other section")
        Try
            'define the variables
            Dim genericCell As System.Web.UI.WebControls.TableCell
            Dim genericRow As System.Web.UI.WebControls.TableRow
            Dim genericTable As System.Web.UI.WebControls.Table
            Dim genericLabel As System.Web.UI.WebControls.Label

            'get a master list of text questions for the survey
            Me.sqlCommonCommand.CommandText = "select QUESTION_ID, QUESTION_TEXT from " & _
                myUtility.getExtension() & "SAT_TEMPLATE_QUESTION where " & _
                "TEMPLATE_KEY = " & Session("seqTemplate") & " and QUESTION_TYPE_CODE in ('M', 'S') " & _
                " and FILTER_IND = 0 order by QUESTION_ID"
            Trace.Warn(Me.sqlCommonCommand.CommandText)
            If (Me.dsCommon.Tables.Contains("TextItems")) Then
                Me.dsCommon.Tables("TextItems").Rows.Clear()
            End If
            Me.sqlCommonAdapter.Fill(Me.dsCommon, "TextItems")

            'get filter item data
            Dim row As DataRow
            Dim blnDoOnce As Boolean = False
            Dim blnDoOnce3 As Boolean = False
            For Each row In Me.dsCommon.Tables("TextItems").Rows
                Dim dr As DataRow

                'add a separator between the questions and the rest of the data
                If (blnDoOnce = False) Then
                    blnDoOnce = True
                    If (Session("NonRTypeSelected") = True) Then
                        'generate a row for the question text
                        genericRow = New System.Web.UI.WebControls.TableRow
                        'genericRow.CssClass = "BreakBefore"
                        genericCell = New System.Web.UI.WebControls.TableCell
                        genericLabel = New System.web.ui.webcontrols.Label
                        genericLabel.Text = "<hr style='BORDER-TOP-STYLE: ridge' width='100%' SIZE='3'/>"
                        genericCell.Controls.Add(genericLabel)
                        genericCell.Wrap = False
                        genericCell.EnableViewState = True
                        genericCell.ColumnSpan = 2
                        genericRow.Cells.Add(genericCell)
                        genericRow.EnableViewState = True
                        Me.DataTable2.Rows.Add(genericRow)
                    End If
                End If

                'set the question filter to bring back only M and S types in the data
                Dim tempSkipList As ArrayList = New ArrayList
                tempSkipList.Add(row.Item("QUESTION_ID"))
                setQuestionFilter(tempSkipList, True)

                Dim dateFilter As ArrayList
                Dim blnDoOnce2 As Boolean = False
                For Each dateFilter In Me.dateFilterArr
                    If (setDataFilter(dateFilter(2))) Then
                        Me.sqlCommonCommand.CommandText = "select frs.QUESTION_ID, frs.ANSWER_TEXT, fq.QUESTION_TEXT as QuestionText " & _
                        "from " & myUtility.getExtension() & "SAT_RESPONSE_RESULT frs, " & myUtility.getExtension() & _
                        "SAT_TEMPLATE_QUESTION fq where frs.SURVEY_KEY = " & Session("seqSurvey") & " " & _
                        dateFilter(2).Replace("LAST_USED_DATE", "RESPONSE_dATE") & " " & Me.strDataFilter & " " & _
                        Me.strQuestionFilter & " and frs.QUESTION_ID = fq.QUESTION_ID and frs.QUESTION_TYPE = fq.QUESTION_TYPE_CODE " & _
                        "and fq.TEMPLATE_KEY = " & Session("seqTemplate") & " order by frs.QUESTION_ID, frs.ANSWER_TEXT, fq.QUESTION_TEXT"
                        Trace.Warn(Me.sqlCommonCommand.CommandText)
                        If (Me.dsCommon.Tables.Contains("DataDump")) Then
                            Me.dsCommon.Tables("DataDump").Rows.Clear()
                        End If
                        Me.sqlCommonAdapter.Fill(Me.dsCommon, "DataDump")

                        Dim sumAnswered As Double = 0

                        'build a table of percentages for the question using the list items
                        If (Me.dsCommon.Tables("DataDump").Rows.Count > 0) Then
                            'build the skeletal structure
                            Session("QuestionsSelected") = True
                            Dim dt As DataTable = New DataTable
                            dt.Columns.Add("ResultText")
                            dt.Columns("ResultText").DataType = Me.dsCommon.Tables("DataDump").Columns("ANSWER_TEXT").DataType

                            'add the answer text and get the answer counts
                            dt.TableName = "QuestionText" & row.Item("QUESTION_ID")
                            For Each dr In Me.dsCommon.Tables("DataDump").Rows
                                Dim tempRow As DataRow = dt.NewRow
                                tempRow.Item(0) = CType(dr.Item("ANSWER_TEXT"), String)
                                dt.Rows.Add(tempRow)
                                sumAnswered += 1
                            Next

                            If (Not blnDoOnce3) Then
                                'section header
                                genericRow = New System.Web.UI.WebControls.TableRow
                                genericCell = New System.Web.UI.WebControls.TableCell
                                genericCell.Wrap = False
                                genericCell.CssClass = "sectionHeadTeal"
                                genericCell.Text = "Text Results"
                                genericCell.EnableViewState = True
                                genericCell.ColumnSpan = 2
                                genericRow.Cells.Add(genericCell)
                                genericRow.EnableViewState = True
                                Me.DataTable2.Rows.Add(genericRow)
                                blnDoOnce3 = True
                            End If

                            'buffer zone
                            genericRow = New System.Web.UI.WebControls.TableRow
                            genericCell = New System.Web.UI.WebControls.TableCell
                            genericCell.Wrap = False
                            genericCell.Height = Unit.Pixel(20)
                            genericCell.EnableViewState = True
                            genericCell.ColumnSpan = 2
                            genericRow.Cells.Add(genericCell)
                            genericRow.EnableViewState = True
                            Me.DataTable2.Rows.Add(genericRow)

                            If (Session("SplitOutput") <> "NONE") Then
                                'generate a row for the question text
                                genericRow = New System.Web.UI.WebControls.TableRow
                                genericCell = New System.Web.UI.WebControls.TableCell
                                genericLabel = New System.web.ui.webcontrols.Label
                                genericLabel.CssClass = "reportContent"
                                genericLabel.Text = Me.dsCommon.Tables("DataDump").Rows(0).Item("QuestionText")
                                genericCell.Controls.Add(genericLabel)
                                genericLabel = New System.web.ui.webcontrols.Label
                                genericLabel.CssClass = "boldContentTeal"
                                genericLabel.Text = "(" & sumAnswered & ")"
                                genericCell.Controls.Add(genericLabel)
                                genericCell.Wrap = True
                                genericCell.EnableViewState = True
                                genericCell.ColumnSpan = 2
                                genericRow.Cells.Add(genericCell)
                                genericRow.EnableViewState = True
                                If (Session("SplitOutput") <> "NONE") Then
                                    If (blnDoOnce2) Then
                                        genericRow.BackColor = Color.LightSteelBlue
                                    Else
                                        genericRow.BackColor = Color.LemonChiffon
                                    End If
                                End If
                                Me.DataTable2.Rows.Add(genericRow)
                                blnDoOnce2 = True
                            ElseIf (Session("SplitOutput") = "NONE") Then
                                'generate a row for the question text
                                genericRow = New System.Web.UI.WebControls.TableRow
                                genericCell = New System.Web.UI.WebControls.TableCell
                                genericLabel = New System.web.ui.webcontrols.Label
                                genericLabel.CssClass = "reportContent"
                                genericLabel.Text = Me.dsCommon.Tables("DataDump").Rows(0).Item("QuestionText")
                                genericCell.Controls.Add(genericLabel)
                                genericLabel = New System.web.ui.webcontrols.Label
                                genericLabel.CssClass = "boldContentTeal"
                                genericLabel.Text = "(" & sumAnswered & ")"
                                genericCell.Controls.Add(genericLabel)
                                genericCell.Wrap = True
                                genericCell.EnableViewState = True
                                genericCell.ColumnSpan = 2
                                genericRow.Cells.Add(genericCell)
                                genericRow.EnableViewState = True
                                Me.DataTable2.Rows.Add(genericRow)
                            End If

                            If (Session("SplitOutput") <> "NONE") Then
                                'generate a row for the date range
                                genericRow = New System.Web.UI.WebControls.TableRow
                                genericCell = New System.Web.UI.WebControls.TableCell
                                genericLabel = New System.web.ui.webcontrols.Label
                                genericLabel.CssClass = "ReportHeaderData"
                                genericLabel.Text = "Period of Record Set: "
                                genericCell.Controls.Add(genericLabel)
                                genericLabel = New System.web.ui.webcontrols.Label
                                genericLabel.CssClass = "ReportHeaderBlue"
                                genericLabel.Text = dateFilter(0) & " - " & dateFilter(1)
                                genericCell.Controls.Add(genericLabel)
                                genericCell.Wrap = True
                                genericCell.EnableViewState = True
                                genericCell.ColumnSpan = 2
                                genericRow.Cells.Add(genericCell)
                                genericRow.EnableViewState = True
                                genericRow.BackColor = Color.LightGray
                                Me.DataTable2.Rows.Add(genericRow)
                            End If

                            'generate a row for the question answer
                            genericTable = New System.Web.UI.WebControls.Table
                            Dim colorIndex As Integer = 0
                            For Each dr In dt.Rows
                                'populate row data
                                genericRow = New System.Web.UI.WebControls.TableRow
                                genericCell = New System.Web.UI.WebControls.TableCell
                                genericCell.Width = Unit.Pixel(20)
                                genericCell.Wrap = False
                                genericCell.EnableViewState = True
                                genericRow.Cells.Add(genericCell)
                                genericCell = New System.Web.UI.WebControls.TableCell
                                genericLabel = New System.web.ui.webcontrols.Label
                                genericLabel.CssClass = "reportContent"
                                genericLabel.Text = dr.Item("ResultText")
                                genericCell.Controls.Add(genericLabel)

                                'set the row color
                                If (Session("SplitOutput") <> "NONE") Then
                                    If (colorIndex Mod 2 = 0) Then
                                        genericCell.BackColor = Color.White
                                    Else
                                        genericCell.BackColor = Color.LightGray
                                    End If
                                Else
                                    If (colorIndex Mod 2 = 0) Then
                                        genericCell.BackColor = Color.LightGray
                                    Else
                                        genericCell.BackColor = Color.White
                                    End If
                                End If
                                colorIndex += 1

                                genericCell.Wrap = True
                                genericCell.EnableViewState = True
                                genericRow.Cells.Add(genericCell)
                                genericRow.EnableViewState = True
                                genericTable.Rows.Add(genericRow)
                            Next

                            genericRow = New System.Web.UI.WebControls.TableRow

                            'Set to put a line break after 3 question results appear on the page.
                            'If (questionCounter Mod Session("QuestionsPerPage") = 0) Then
                            'genericRow.CssClass = "Break"
                            'End If
                            'questionCounter += 1

                            genericCell = New System.Web.UI.WebControls.TableCell
                            genericCell.ColumnSpan = 2
                            genericCell.Controls.Add(genericTable)
                            genericCell.Wrap = False
                            genericCell.EnableViewState = True
                            genericCell.VerticalAlign = VerticalAlign.Top
                            genericRow.Cells.Add(genericCell)
                            Me.DataTable2.Rows.Add(genericRow)
                            Me.DataTable2.EnableViewState = True
                        Else
                            If (Not blnDoOnce3) Then
                                'section header
                                genericRow = New System.Web.UI.WebControls.TableRow
                                genericCell = New System.Web.UI.WebControls.TableCell
                                genericCell.Wrap = False
                                genericCell.CssClass = "sectionHeadTeal"
                                genericCell.Text = "Text Results"
                                genericCell.EnableViewState = True
                                genericCell.ColumnSpan = 2
                                genericRow.Cells.Add(genericCell)
                                genericRow.EnableViewState = True
                                Me.DataTable2.Rows.Add(genericRow)
                                blnDoOnce3 = True
                            End If

                            'buffer zone
                            genericRow = New System.Web.UI.WebControls.TableRow
                            genericCell = New System.Web.UI.WebControls.TableCell
                            genericCell.Wrap = False
                            genericCell.Height = Unit.Pixel(40)
                            genericCell.EnableViewState = True
                            genericCell.ColumnSpan = 2
                            genericRow.Cells.Add(genericCell)
                            genericRow.EnableViewState = True
                            Me.DataTable2.Rows.Add(genericRow)

                            If (Session("SplitOutput") <> "NONE") Then
                                'generate a row for the question text
                                genericRow = New System.Web.UI.WebControls.TableRow
                                genericCell = New System.Web.UI.WebControls.TableCell
                                genericLabel = New System.web.ui.webcontrols.Label
                                genericLabel.CssClass = "reportContent"
                                genericLabel.Text = row.Item("QUESTION_TEXT")
                                genericCell.Controls.Add(genericLabel)
                                genericLabel = New System.web.ui.webcontrols.Label
                                genericLabel.CssClass = "boldContentTeal"
                                genericLabel.Text = "(" & sumAnswered & ")"
                                genericCell.Controls.Add(genericLabel)
                                genericCell.Wrap = True
                                genericCell.EnableViewState = True
                                genericCell.ColumnSpan = 2
                                genericRow.Cells.Add(genericCell)
                                genericRow.EnableViewState = True
                                If (Session("SplitOutput") <> "NONE") Then
                                    If (blnDoOnce2) Then
                                        genericRow.BackColor = Color.LightSteelBlue
                                    Else
                                        genericRow.BackColor = Color.LemonChiffon
                                    End If
                                End If
                                Me.DataTable2.Rows.Add(genericRow)
                                blnDoOnce2 = True
                            ElseIf (Session("SplitOutput") = "NONE") Then
                                'generate a row for the question text
                                genericRow = New System.Web.UI.WebControls.TableRow
                                genericCell = New System.Web.UI.WebControls.TableCell
                                genericLabel = New System.web.ui.webcontrols.Label
                                genericLabel.CssClass = "reportContent"
                                genericLabel.Text = row.Item("QUESTION_TEXT")
                                genericCell.Controls.Add(genericLabel)
                                genericLabel = New System.web.ui.webcontrols.Label
                                genericLabel.CssClass = "boldContentTeal"
                                genericLabel.Text = "(" & sumAnswered & ")"
                                genericCell.Controls.Add(genericLabel)
                                genericCell.Wrap = True
                                genericCell.EnableViewState = True
                                genericCell.ColumnSpan = 2
                                genericRow.Cells.Add(genericCell)
                                genericRow.EnableViewState = True
                                Me.DataTable2.Rows.Add(genericRow)
                            End If

                            If (Session("SplitOutput") <> "NONE") Then
                                'generate a row for the date range
                                genericRow = New System.Web.UI.WebControls.TableRow
                                genericCell = New System.Web.UI.WebControls.TableCell
                                genericLabel = New System.web.ui.webcontrols.Label
                                genericLabel.CssClass = "ReportHeaderData"
                                genericLabel.Text = "Period of Record Set: "
                                genericCell.Controls.Add(genericLabel)
                                genericLabel = New System.web.ui.webcontrols.Label
                                genericLabel.CssClass = "ReportHeaderBlue"
                                genericLabel.Text = dateFilter(0) & " - " & dateFilter(1)
                                genericCell.Controls.Add(genericLabel)
                                genericCell.Wrap = True
                                genericCell.EnableViewState = True
                                genericCell.ColumnSpan = 2
                                genericRow.Cells.Add(genericCell)
                                genericRow.EnableViewState = True
                                genericRow.BackColor = Color.LightGray
                                Me.DataTable2.Rows.Add(genericRow)
                            End If

                            'generate to inform user that there is no data for this date range
                            genericRow = New System.Web.UI.WebControls.TableRow
                            genericTable = New System.Web.UI.WebControls.Table
                            genericCell = New System.Web.UI.WebControls.TableCell
                            genericCell.Width = Unit.Pixel(20)
                            genericCell.Wrap = False
                            genericCell.EnableViewState = True
                            genericRow.Cells.Add(genericCell)
                            genericCell = New System.Web.UI.WebControls.TableCell
                            genericLabel = New System.web.ui.webcontrols.Label
                            genericLabel.CssClass = "NoDataRed"
                            genericLabel.Text = "This question has no data for this time period."
                            genericCell.Controls.Add(genericLabel)
                            'genericCell.ColumnSpan = 2
                            'genericCell.HorizontalAlign = HorizontalAlign.Center
                            genericRow.Cells.Add(genericCell)
                            genericRow.EnableViewState = True
                            genericTable.Rows.Add(genericRow)
                            genericRow = New System.Web.UI.WebControls.TableRow
                            genericCell = New System.Web.UI.WebControls.TableCell
                            genericCell.ColumnSpan = 2
                            genericCell.Controls.Add(genericTable)
                            genericRow.Cells.Add(genericCell)
                            Me.DataTable2.Rows.Add(genericRow)
                        End If
                    Else
                        Session("Alert") = "Failed to get survey question list item data for text."
                        Return False
                    End If
                Next
            Next
            Return True
        Catch ex As Exception
            Session("Alert") = "Failed to get survey data.  " & ex.ToString
            Return False
        End Try
    End Function
#End Region

#Region "Chart Generation"
    'generates a chart for each question passed to it
    Private Function generateChart(ByRef genericChart As Infragistics.webui.UltraWebChart.UltraChart, ByVal intQuestion As Integer, ByVal chartType As String, ByVal dateFilter As String, ByVal dt As DataTable, ByVal questionList As ArrayList, ByVal runCount As Integer, Optional ByVal isChoice As Boolean = False, Optional ByVal isRating As Boolean = False, Optional ByVal isDate As Boolean = False, Optional ByVal percentDisplay As Boolean = False) As Boolean
        Trace.Warn("Generating Chart")
        Try
            If (isRating) Then
                If (chartType = "MULTI") Then
                    Dim index As Integer = 1
                    While (index <= questionList.Count())
                        If (percentDisplay) Then
                            dt.Columns.Remove("People" & index)
                            dt.Columns("Percentage" & index).ColumnName = index
                        Else
                            dt.Columns.Remove("Percentage" & index)
                            dt.Columns("People" & index).ColumnName = index
                        End If
                        index += 1
                    End While
                Else
                    If (percentDisplay) Then
                        dt.Columns.Remove("People")
                    Else
                        dt.Columns.Remove("Percentage")
                    End If
                End If
            Else
                dt.Columns.Remove("ListLabel")
                If (chartType = "MULTI") Then
                    Dim index As Integer = 1
                    While (index <= questionList.Count())
                        If (percentDisplay) Then
                            dt.Columns.Remove("AnswerCount" & index)
                            dt.Columns("Percentage" & index).ColumnName = index
                        Else
                            dt.Columns.Remove("Percentage" & index)
                            dt.Columns("AnswerCount" & index).ColumnName = index
                        End If
                        index += 1
                    End While
                Else
                    If (percentDisplay) Then
                        dt.Columns.Remove("AnswerCount")
                    Else
                        dt.Columns.Remove("Percentage")
                    End If
                End If
            End If

            If (dt.Rows.Count > 0) Then
                'get the maximum extent of the data
                Dim maxExtent As Integer = 0
                Dim tickInterval As Integer = 0
                Dim dr As DataRow
                For Each dr In dt.Rows
                    If (maxExtent < dr.Item(1)) Then
                        maxExtent = dr.Item(1)
                    End If
                Next
                If (maxExtent <= 10) Then
                    tickInterval = 1
                ElseIf (maxExtent <= 20) Then
                    tickInterval = 2
                ElseIf (maxExtent <= 100) Then
                    tickInterval = 10
                ElseIf (maxExtent <= 1000) Then
                    tickInterval = 100
                ElseIf (maxExtent <= 10000) Then
                    tickInterval = 1000
                ElseIf (maxExtent <= 100000) Then
                    tickInterval = 10000
                ElseIf (maxExtent <= 1000000) Then
                    tickInterval = 100000
                End If

                'get the chart type
                If (chartType = "3DBAR") Then
                    'chart data binding
                    'dt.Columns.Remove("Value")
                    genericChart.DataSource = dt
                    genericChart.DataBind()

                    'chart type specifics
                    genericChart.ChartType = Infragistics.UltraChart.[Shared].Styles.ChartType.ColumnChart3D
                    genericChart.ColumnChart3D.ColumnSpacing = 1

                    'axis specifics
                    genericChart.Axis.X.Labels.Visible = False
                    genericChart.Axis.X.TickmarkInterval = 0
                    genericChart.Axis.X.MajorGridLines.Visible = False
                    genericChart.Axis.X.Labels.Flip = False
                    genericChart.Axis.X.Labels.Orientation = Infragistics.UltraChart.[Shared].Styles.TextOrientation.Horizontal
                    genericChart.Axis.X.Labels.SeriesLabels.Visible = False
                    genericChart.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#0>"
                    genericChart.Axis.Y.Labels.ItemFormat = Infragistics.UltraChart.[Shared].Styles.AxisItemLabelFormat.Custom
                    genericChart.Axis.Y.TickmarkInterval = tickInterval
                    genericChart.Axis.Y.TickmarkPercentage = (tickInterval / maxExtent) * 100
                    genericChart.Axis.Y.TickmarkStyle = Infragistics.UltraChart.[Shared].Styles.AxisTickStyle.DataInterval
                    genericChart.Axis.Y.TickmarkIntervalType = Infragistics.UltraChart.[Shared].Styles.AxisIntervalType.Ticks
                    genericChart.Axis.Y.NumericAxisType = Infragistics.UltraChart.[Shared].Styles.NumericAxisType.Linear
                    If (percentDisplay) Then
                        genericChart.Axis.Y.RangeType = Infragistics.UltraChart.[Shared].Styles.AxisRangeType.Custom
                        genericChart.Axis.Y.RangeMin = 0
                        genericChart.Axis.Y.RangeMax = 100
                        genericChart.Axis.Y.TickmarkPercentage = 20
                        genericChart.Axis.Y.TickmarkInterval = 20
                    End If
                    genericChart.Axis.Z.Labels.Visible = False

                    'data specifics
                    genericChart.Data.ZeroAligned = True
                    genericChart.Data.SwapRowsAndColumns = False
                    genericChart.Data.UseRowLabelsColumn = True
                    genericChart.Data.RowLabelsColumn = 0

                    '3D Transforms
                    genericChart.Transform3D.Scale = 60
                    genericChart.Transform3D.XRotation = 30
                    genericChart.Transform3D.YRotation = 0
                    genericChart.Transform3D.ZRotation = 260
                ElseIf (chartType = "3DCOLUMN") Then
                    'chart data binding
                    'dt.Columns.Remove("Value")
                    genericChart.DataSource = dt
                    genericChart.DataBind()

                    'chart type specifics
                    genericChart.ChartType = Infragistics.UltraChart.[Shared].Styles.ChartType.ColumnChart3D
                    genericChart.ColumnChart3D.ColumnSpacing = 1

                    'axis specifics
                    genericChart.Axis.X.Labels.Visible = False
                    genericChart.Axis.X.TickmarkInterval = 0
                    genericChart.Axis.X.MajorGridLines.Visible = False
                    genericChart.Axis.X.Labels.Flip = False
                    genericChart.Axis.X.Labels.Orientation = Infragistics.UltraChart.[Shared].Styles.TextOrientation.Horizontal
                    genericChart.Axis.X.Labels.SeriesLabels.Visible = False
                    genericChart.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#0>"
                    genericChart.Axis.Y.Labels.ItemFormat = Infragistics.UltraChart.[Shared].Styles.AxisItemLabelFormat.Custom
                    genericChart.Axis.Y.TickmarkInterval = tickInterval
                    genericChart.Axis.Y.TickmarkPercentage = (tickInterval / maxExtent) * 100
                    genericChart.Axis.Y.TickmarkStyle = Infragistics.UltraChart.[Shared].Styles.AxisTickStyle.DataInterval
                    genericChart.Axis.Y.TickmarkIntervalType = Infragistics.UltraChart.[Shared].Styles.AxisIntervalType.Ticks
                    genericChart.Axis.Y.NumericAxisType = Infragistics.UltraChart.[Shared].Styles.NumericAxisType.Linear
                    If (percentDisplay) Then
                        genericChart.Axis.Y.RangeType = Infragistics.UltraChart.[Shared].Styles.AxisRangeType.Custom
                        genericChart.Axis.Y.RangeMin = 0
                        genericChart.Axis.Y.RangeMax = 100
                        genericChart.Axis.Y.TickmarkPercentage = 10
                        genericChart.Axis.Y.TickmarkInterval = 10
                    End If
                    genericChart.Axis.Z.Labels.Visible = False

                    'data specifics
                    genericChart.Data.ZeroAligned = True
                    genericChart.Data.SwapRowsAndColumns = False
                    genericChart.Data.UseRowLabelsColumn = True
                    genericChart.Data.RowLabelsColumn = 0

                    '3D Transforms
                    genericChart.Transform3D.Scale = 80
                    genericChart.Transform3D.XRotation = 144
                    genericChart.Transform3D.YRotation = 75
                    genericChart.Transform3D.ZRotation = 0
                ElseIf (chartType = "3DDOUGHNUT") Then
                    'chart data binding
                    genericChart.DataSource = dt
                    genericChart.DataBind()

                    'chart type specifics
                    genericChart.ChartType = Infragistics.UltraChart.[Shared].Styles.ChartType.DoughnutChart3D
                    genericChart.DoughnutChart3D.ColumnIndex = 1
                    genericChart.DoughnutChart3D.OthersCategoryPercent = 0
                    genericChart.DoughnutChart3D.StartAngle = 0
                    genericChart.DoughnutChart3D.PieThickness = 10

                    'data specifics
                    genericChart.Data.UseRowLabelsColumn = True
                    genericChart.Data.RowLabelsColumn = 0

                    '3D Transforms
                    genericChart.Transform3D.Scale = 100
                    genericChart.Transform3D.XRotation = 40
                    genericChart.Transform3D.YRotation = 180
                    genericChart.Transform3D.ZRotation = 0
                ElseIf (chartType = "3DBRKDOUGH") Then
                    'chart data binding
                    genericChart.DataSource = dt
                    genericChart.DataBind()

                    'chart type specifics
                    genericChart.ChartType = Infragistics.UltraChart.[Shared].Styles.ChartType.DoughnutChart3D
                    genericChart.DoughnutChart3D.ColumnIndex = 1
                    genericChart.DoughnutChart3D.OthersCategoryPercent = 0
                    genericChart.DoughnutChart3D.StartAngle = 0
                    genericChart.DoughnutChart3D.PieThickness = 10
                    genericChart.DoughnutChart3D.BreakAllSlices = True

                    'data specifics
                    genericChart.Data.UseRowLabelsColumn = True
                    genericChart.Data.RowLabelsColumn = 0

                    '3D Transforms
                    genericChart.Transform3D.Scale = 80
                    genericChart.Transform3D.XRotation = 40
                    genericChart.Transform3D.YRotation = 180
                    genericChart.Transform3D.ZRotation = 0
                ElseIf (chartType = "3DPIE") Then
                    'chart data binding
                    genericChart.DataSource = dt
                    genericChart.DataBind()

                    'chart type specifics
                    genericChart.ChartType = Infragistics.UltraChart.[Shared].Styles.ChartType.PieChart3D
                    genericChart.PieChart3D.ColumnIndex = 1
                    genericChart.PieChart3D.OthersCategoryPercent = 0.0
                    genericChart.PieChart3D.StartAngle = 0
                    genericChart.PieChart3D.PieThickness = 10
                    genericChart.PieChart3D.Labels.FormatString = "<PERCENT_VALUE:#0>%"

                    'data specifics
                    genericChart.Data.UseRowLabelsColumn = True
                    genericChart.Data.RowLabelsColumn = 0

                    '3D Transforms
                    'genericChart.Transform3D.Scale = 100
                    genericChart.Transform3D.XRotation = 60
                    genericChart.Transform3D.YRotation = 180
                    genericChart.Transform3D.ZRotation = 0
                ElseIf (chartType = "3DBRKPIE") Then
                    'chart data binding
                    genericChart.DataSource = dt
                    genericChart.DataBind()

                    'chart type specifics
                    genericChart.ChartType = Infragistics.UltraChart.[Shared].Styles.ChartType.PieChart3D
                    genericChart.PieChart3D.ColumnIndex = 1
                    genericChart.PieChart3D.OthersCategoryPercent = 0.0
                    genericChart.PieChart3D.StartAngle = 0
                    genericChart.PieChart3D.PieThickness = 10
                    genericChart.PieChart3D.BreakAllSlices = True
                    genericChart.PieChart3D.Labels.FormatString = "<PERCENT_VALUE:#0>%"

                    'data specifics
                    genericChart.Data.UseRowLabelsColumn = True
                    genericChart.Data.RowLabelsColumn = 0

                    '3D Transforms
                    genericChart.Transform3D.Scale = 80
                    genericChart.Transform3D.XRotation = 60
                    genericChart.Transform3D.YRotation = 180
                    genericChart.Transform3D.ZRotation = 0
                ElseIf (chartType = "BAR") Then
                    'chart data binding
                    genericChart.DataSource = dt
                    genericChart.DataBind()

                    'chart type specifics
                    genericChart.ChartType = Infragistics.UltraChart.[Shared].Styles.ChartType.BarChart
                    genericChart.BarChart.BarSpacing = 1

                    'axis specifics
                    genericChart.Axis.X.Labels.Visible = True
                    genericChart.Axis.X.Labels.ItemFormat = Infragistics.UltraChart.[Shared].Styles.AxisItemLabelFormat.Custom
                    genericChart.Axis.X.NumericAxisType = Infragistics.UltraChart.[Shared].Styles.NumericAxisType.Linear
                    genericChart.Axis.X.TickmarkStyle = Infragistics.UltraChart.[Shared].Styles.AxisTickStyle.DataInterval
                    genericChart.Axis.X.TickmarkIntervalType = Infragistics.UltraChart.[Shared].Styles.AxisIntervalType.Ticks
                    genericChart.Axis.X.Labels.ItemFormatString = "<DATA_VALUE:#0>"
                    genericChart.Axis.X.TickmarkInterval = tickInterval
                    genericChart.Axis.X.MajorGridLines.Visible = True
                    genericChart.Axis.X.Labels.Flip = False
                    genericChart.Axis.X.Labels.Orientation = Infragistics.UltraChart.[Shared].Styles.TextOrientation.Horizontal
                    genericChart.Axis.X.Labels.SeriesLabels.Visible = True
                    If (percentDisplay) Then
                        genericChart.Axis.X.RangeType = Infragistics.UltraChart.[Shared].Styles.AxisRangeType.Custom
                        genericChart.Axis.X.RangeMin = 0
                        genericChart.Axis.X.RangeMax = 100
                        genericChart.Axis.X.TickmarkInterval = 10
                    End If
                    genericChart.Axis.Y.MajorGridLines.Visible = False
                    genericChart.Axis.Y.Labels.Flip = False
                    genericChart.Axis.Y.Labels.SeriesLabels.Visible = False
                    genericChart.Axis.Y.Labels.Visible = False

                    'data specifics
                    genericChart.Data.ZeroAligned = True
                    genericChart.Data.RowLabelsColumn = 0
                    genericChart.Data.UseRowLabelsColumn = True
                    genericChart.Data.SwapRowsAndColumns = True
                ElseIf (chartType = "BRKDOUGH") Then
                    'chart data binding
                    genericChart.DataSource = dt
                    genericChart.DataBind()

                    'chart type specifics
                    genericChart.ChartType = Infragistics.UltraChart.[Shared].Styles.ChartType.DoughnutChart
                    genericChart.DoughnutChart.ColumnIndex = 1
                    genericChart.DoughnutChart.OthersCategoryPercent = 0
                    genericChart.DoughnutChart.StartAngle = 270
                    genericChart.DoughnutChart.BreakAllSlices = True
                ElseIf (chartType = "BRKPIE") Then
                    'chart data binding
                    genericChart.DataSource = dt
                    genericChart.DataBind()

                    'chart type specifics
                    genericChart.ChartType = Infragistics.UltraChart.[Shared].Styles.ChartType.PieChart
                    genericChart.PieChart.ColumnIndex = 1
                    genericChart.PieChart.OthersCategoryPercent = 0
                    genericChart.PieChart.StartAngle = 270
                    genericChart.PieChart.BreakAllSlices = True
                ElseIf (chartType = "BUBBLE") Then
                    'chart data binding
                    genericChart.DataSource = dt
                    genericChart.DataBind()

                    'chart type specifics
                    genericChart.ChartType = Infragistics.UltraChart.[Shared].Styles.ChartType.BubbleChart
                    genericChart.BubbleChart.ColumnX = 0
                    genericChart.BubbleChart.ColumnY = 1
                    genericChart.BubbleChart.ColumnZ = 1
                    genericChart.BubbleChart.BubbleShape = Infragistics.UltraChart.[Shared].Styles.BubbleShape.Circle
                    genericChart.BubbleChart.SortByRadius = Infragistics.UltraChart.[Shared].Styles.ChartSortType.None

                    'axis specifics
                    genericChart.Axis.X.Labels.Visible = False
                    If (isChoice Or isDate) Then
                        genericChart.Axis.X.NumericAxisType = Infragistics.UltraChart.[Shared].Styles.NumericAxisType.Logarithmic
                        genericChart.Axis.X.LogBase = 2
                    End If
                    genericChart.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#0>"
                    genericChart.Axis.Y.Labels.ItemFormat = Infragistics.UltraChart.[Shared].Styles.AxisItemLabelFormat.Custom
                    genericChart.Axis.Y.TickmarkInterval = tickInterval
                    genericChart.Axis.Y.TickmarkStyle = Infragistics.UltraChart.[Shared].Styles.AxisTickStyle.DataInterval
                    genericChart.Axis.Y.NumericAxisType = Infragistics.UltraChart.[Shared].Styles.NumericAxisType.Linear
                    If (percentDisplay) Then
                        genericChart.Axis.Y.RangeType = Infragistics.UltraChart.[Shared].Styles.AxisRangeType.Custom
                        genericChart.Axis.Y.RangeMin = 0
                        genericChart.Axis.Y.RangeMax = 100
                        genericChart.Axis.Y.TickmarkInterval = 10
                    End If


                    'data specifics
                    genericChart.Data.SwapRowsAndColumns = False
                    If (isChoice Or isDate) Then
                        genericChart.Data.ZeroAligned = False
                    Else
                        genericChart.Data.ZeroAligned = True
                    End If
                ElseIf (chartType = "COLUMN") Then
                    'chart data binding
                    genericChart.DataSource = dt
                    genericChart.DataBind()

                    'chart type specifics
                    genericChart.ChartType = Infragistics.UltraChart.[Shared].Styles.ChartType.ColumnChart
                    genericChart.ColumnChart.ColumnSpacing = 1

                    'axis specifics
                    genericChart.Axis.X.Labels.Visible = False
                    genericChart.Axis.X.TickmarkInterval = 0
                    genericChart.Axis.X.MajorGridLines.Visible = False
                    genericChart.Axis.X.Labels.Flip = False
                    genericChart.Axis.X.Labels.Orientation = Infragistics.UltraChart.[Shared].Styles.TextOrientation.Horizontal
                    genericChart.Axis.X.Labels.SeriesLabels.Visible = False
                    genericChart.Axis.Y.TickmarkInterval = tickInterval
                    genericChart.Axis.Y.TickmarkPercentage = 100
                    genericChart.Axis.Y.MajorGridLines.Visible = True
                    genericChart.Axis.Y.Labels.Flip = False
                    genericChart.Axis.Y.Labels.SeriesLabels.Visible = False
                    genericChart.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#0>"
                    genericChart.Axis.Y.Labels.ItemFormat = Infragistics.UltraChart.[Shared].Styles.AxisItemLabelFormat.Custom
                    genericChart.Axis.Y.NumericAxisType = Infragistics.UltraChart.[Shared].Styles.NumericAxisType.Linear
                    genericChart.Axis.Y.TickmarkStyle = Infragistics.UltraChart.[Shared].Styles.AxisTickStyle.DataInterval
                    genericChart.Axis.Y.TickmarkIntervalType = Infragistics.UltraChart.[Shared].Styles.AxisIntervalType.Ticks
                    If (percentDisplay) Then
                        genericChart.Axis.Y.RangeType = Infragistics.UltraChart.[Shared].Styles.AxisRangeType.Custom
                        genericChart.Axis.Y.RangeMin = 0
                        genericChart.Axis.Y.RangeMax = 100
                        genericChart.Axis.Y.TickmarkInterval = 10
                    End If

                    'data specifics
                    genericChart.Data.ZeroAligned = True
                    genericChart.Data.RowLabelsColumn = 0
                    genericChart.Data.UseRowLabelsColumn = True
                    genericChart.Data.SwapRowsAndColumns = True
                ElseIf (chartType = "DOUGHNUT") Then
                    'chart data binding
                    genericChart.DataSource = dt
                    genericChart.DataBind()

                    'chart type specifics
                    genericChart.ChartType = Infragistics.UltraChart.[Shared].Styles.ChartType.DoughnutChart
                    genericChart.DoughnutChart.ColumnIndex = 1
                    genericChart.DoughnutChart.OthersCategoryPercent = 0
                    genericChart.DoughnutChart.StartAngle = 270
                ElseIf (chartType = "PIE") Then
                    'chart data binding
                    genericChart.DataSource = dt
                    genericChart.DataBind()
                    genericChart.PieChart.Labels.FormatString = "<PERCENT_VALUE:#0>%"
                    'Dim myFont As New Font("Verdana", 6)
                    'genericChart.PieChart.Labels.Font = myFont

                    'chart type specifics
                    genericChart.ChartType = Infragistics.UltraChart.[Shared].Styles.ChartType.PieChart
                    genericChart.PieChart.ColumnIndex = 1
                    genericChart.PieChart.OthersCategoryPercent = 0
                    genericChart.PieChart.StartAngle = 270
                ElseIf (chartType = "MULTI") Then
                    'chart data binding
                    genericChart.DataSource = dt
                    genericChart.DataBind()

                    'chart type specifics
                    genericChart.ChartType = Infragistics.UltraChart.[Shared].Styles.ChartType.ColumnChart

                    'axis specifics
                    If (isRating) Then
                        genericChart.Axis.X.Labels.Visible = True
                        genericChart.Axis.X.TickmarkInterval = 0
                        genericChart.Axis.X.MajorGridLines.Visible = True
                        genericChart.Axis.X.Labels.Flip = False
                        genericChart.Axis.X.Labels.Orientation = Infragistics.UltraChart.[Shared].Styles.TextOrientation.VerticalRightFacing
                        genericChart.Axis.X.Labels.SeriesLabels.Visible = True
                    Else
                        genericChart.Axis.X.Labels.Visible = False
                        genericChart.Axis.X.TickmarkInterval = 0
                        genericChart.Axis.X.MajorGridLines.Visible = True
                        genericChart.Axis.X.Labels.Flip = False
                        genericChart.Axis.X.Labels.Orientation = Infragistics.UltraChart.[Shared].Styles.TextOrientation.VerticalRightFacing
                        genericChart.Axis.X.Labels.SeriesLabels.Visible = True
                    End If
                    genericChart.Axis.Y.TickmarkInterval = tickInterval
                    genericChart.Axis.Y.TickmarkPercentage = 100
                    genericChart.Axis.Y.MajorGridLines.Visible = True
                    genericChart.Axis.Y.Labels.Flip = False
                    genericChart.Axis.Y.Labels.SeriesLabels.Visible = False
                    genericChart.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#0>"
                    genericChart.Axis.Y.Labels.ItemFormat = Infragistics.UltraChart.[Shared].Styles.AxisItemLabelFormat.Custom
                    genericChart.Axis.Y.NumericAxisType = Infragistics.UltraChart.[Shared].Styles.NumericAxisType.Linear
                    genericChart.Axis.Y.TickmarkStyle = Infragistics.UltraChart.[Shared].Styles.AxisTickStyle.DataInterval
                    genericChart.Axis.Y.TickmarkIntervalType = Infragistics.UltraChart.[Shared].Styles.AxisIntervalType.Ticks
                    If (percentDisplay) Then
                        genericChart.Axis.Y.RangeType = Infragistics.UltraChart.[Shared].Styles.AxisRangeType.Custom
                        genericChart.Axis.Y.RangeMin = 0
                        genericChart.Axis.Y.RangeMax = 100
                        genericChart.Axis.Y.TickmarkInterval = 10
                    End If

                    'data specifics
                    genericChart.Data.ZeroAligned = True
                    genericChart.Data.RowLabelsColumn = 0
                    genericChart.Data.UseRowLabelsColumn = True
                    genericChart.Data.SwapRowsAndColumns = True
                ElseIf (chartType = "SCATTER") Then
                    'chart data binding
                    genericChart.DataSource = dt
                    genericChart.DataBind()

                    'chart type specifics
                    genericChart.ChartType = Infragistics.UltraChart.[Shared].Styles.ChartType.ScatterChart
                    genericChart.ScatterChart.ColumnX = 0
                    genericChart.ScatterChart.ColumnY = 1
                    genericChart.ScatterChart.Icon = Infragistics.UltraChart.[Shared].Styles.SymbolIcon.Circle
                    genericChart.ScatterChart.IconSize = Infragistics.UltraChart.[Shared].Styles.SymbolIconSize.Medium
                    genericChart.ScatterChart.UseGroupByColumn = True
                    genericChart.ScatterChart.GroupByColumn = 0

                    'axis specifics
                    genericChart.Axis.X.Labels.Visible = False
                    If (isChoice Or isDate) Then
                        genericChart.Axis.X.NumericAxisType = Infragistics.UltraChart.[Shared].Styles.NumericAxisType.Logarithmic
                        genericChart.Axis.X.LogBase = 2
                    End If
                    genericChart.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#0>"
                    genericChart.Axis.Y.Labels.ItemFormat = Infragistics.UltraChart.[Shared].Styles.AxisItemLabelFormat.Custom
                    genericChart.Axis.Y.TickmarkInterval = tickInterval
                    genericChart.Axis.Y.TickmarkStyle = Infragistics.UltraChart.[Shared].Styles.AxisTickStyle.DataInterval
                    genericChart.Axis.Y.NumericAxisType = Infragistics.UltraChart.[Shared].Styles.NumericAxisType.Linear
                    If (percentDisplay) Then
                        genericChart.Axis.Y.RangeType = Infragistics.UltraChart.[Shared].Styles.AxisRangeType.Custom
                        genericChart.Axis.Y.RangeMin = 0
                        genericChart.Axis.Y.RangeMax = 100
                        genericChart.Axis.Y.TickmarkInterval = 10
                    End If

                    'data specifics
                    genericChart.Data.SwapRowsAndColumns = False
                    If (isChoice Or isDate) Then
                        genericChart.Data.ZeroAligned = False
                    Else
                        genericChart.Data.ZeroAligned = True
                    End If
                End If
                genericChart.ID = "Chart" & intQuestion & "^" & runCount
                'genericChart.BackgroundImageStyle = Infragistics.UltraChart.[Shared].Styles.ImageFitStyle.StretchedFit
                'genericChart.CssClass = "content2"
                genericChart.Axis.X.Extent = 35
                genericChart.Axis.Y.Extent = 35
                genericChart.JavaScriptEnabled = True
                genericChart.JavaScriptFileName = "./MyInfragisticsScripts/ig_webchart.js"
                genericChart.Tooltips.Display = Infragistics.UltraChart.[Shared].Styles.TooltipDisplay.MouseClick
                genericChart.Tooltips.Format = Infragistics.UltraChart.[Shared].Styles.TooltipStyle.DataValue
                genericChart.Tooltips.FormatString = "<DATA_VALUE:#0>"
                genericChart.Tooltips.HotTrackingEnabled = True

                'get the colors for the chart
                Dim myColorArray(Me.dsCommon.Tables("Colors").Rows.Count()) As System.Drawing.Color
                Dim index As Integer = 0
                If (Me.dsCommon.Tables("Colors").Rows.Count() > 0) Then
                    For Each dr In Me.dsCommon.Tables("Colors").Rows
                        Dim Alpha As Integer = 0
                        Dim Red As Integer = 0
                        Dim Green As Integer = 0
                        Dim Blue As Integer = 0
                        If (dr.Item("CHART_COLOR_CODE").Length = 8) Then
                            Alpha = Integer.Parse(dr.Item("CHART_COLOR_CODE").Substring(0, 2), Globalization.NumberStyles.HexNumber)
                            Red = Integer.Parse(dr.Item("CHART_COLOR_CODE").Substring(2, 2), Globalization.NumberStyles.HexNumber)
                            Green = Integer.Parse(dr.Item("CHART_COLOR_CODE").Substring(4, 2), Globalization.NumberStyles.HexNumber)
                            Blue = Integer.Parse(dr.Item("CHART_COLOR_CODE").Substring(6, 2), Globalization.NumberStyles.HexNumber)
                        Else
                            Red = Integer.Parse(dr.Item("CHART_COLOR_CODE").Substring(0, 2), Globalization.NumberStyles.HexNumber)
                            Green = Integer.Parse(dr.Item("CHART_COLOR_CODE").Substring(2, 2), Globalization.NumberStyles.HexNumber)
                            Blue = Integer.Parse(dr.Item("CHART_COLOR_CODE").Substring(4, 2), Globalization.NumberStyles.HexNumber)
                        End If
                        Dim myColor As System.Drawing.Color = System.Drawing.Color.FromArgb(Alpha, Red, Green, Blue)
                        myColorArray(index) = myColor
                        index += 1
                    Next
                End If
                genericChart.ColorModel.CustomPalette = myColorArray
                genericChart.ColorModel.ModelStyle = Infragistics.UltraChart.[Shared].Styles.ColorModels.CustomLinear
                genericChart.ColorModel.Scaling = Infragistics.UltraChart.[Shared].Styles.ColorScaling.None
                genericChart.ColorModel.AlphaLevel = 255
                genericChart.Width = Unit.Pixel(200)
                genericChart.Height = Unit.Pixel(200)
                genericChart.Border.Thickness = 0
                Return True
            Else
                Session("Alert") = "Failed to generate chart."
                Return False
            End If
        Catch ex As Exception
            Session("Alert") = "Failed to generate chart."
            Return False
        End Try
    End Function
#End Region
End Class
