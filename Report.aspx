<%@ Page trace="false" Language="vb" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" Codebehind="Report.aspx.vb" Inherits="SurveyAdmin.Report" %>
<%@ Register TagPrefix="igsch" Namespace="Infragistics.WebUI.WebSchedule" Assembly="Infragistics2.WebUI.WebDateChooser.v8.2, Version=8.2.20082.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder" Runat="Server">
<%=Session("JSAction")%>
<form id="Form1" method="post" runat="server">
<h1>Generate Report</h1>
<p></p>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="0" cellPadding="2" class="borderless">
	<% if (Session("TemplatesExist") = "Yes") then %>
	<tr>
		<td><asp:image id="hlpTemplateSelect" runat="server" ImageUrl="./images/help.png" ToolTip="Select a template to limit the surveys in the list below to only those within a particular template."></asp:image></td>
		<td align="right"><asp:label id="lblTemplateSelectlblExpires" Runat="server" CssClass="boldContent">Template:</asp:label></td>
		<td align="left"><asp:dropdownlist id="ddlTemplateList" runat="server" CssClass="content" AutoPostBack="True"></asp:dropdownlist></td>
	</tr>
	<% if (Session("SurveysExist") = "Yes") then %>
	<tr>
		<td><asp:image id="hlpSurveySelect" runat="server" ImageUrl="./images/help.png" ToolTip="Select a survey to begin setting up the parameters for the report."></asp:image></td>
		<td align="right"><asp:label id="lblSurveySelect" Runat="server" CssClass="boldContent">Survey:</asp:label></td>
		<td align="left"><asp:dropdownlist id="ddlSurveyList" runat="server" CssClass="content" AutoPostBack="True"></asp:dropdownlist></td>
	</tr>
	<% else %>
	<tr>
		<td colSpan="3"><asp:label id="lblNoSurveys" Runat="server" CssClass="boldContent" ForeColor="Red">There are currently either no surveys that you have access to for this template<br />&nbsp;&nbsp;or there is no survey response data for the surveys created from this template.<br /><br />  Please either create a survey for this template<br />&nbsp;&nbsp;or send out a survey created from this template first.</asp:label></td>
	</tr>
	<% end if %>
	<% else %>
	<tr>
		<td colSpan="3"><asp:label id="lblNoTemplates" Runat="server" CssClass="boldContent" ForeColor="Red">There are currently no templates that you have access to or that have surveys made from them.<br />  Please create a template and/or survey first.</asp:label></td>
	</tr>
	<% end if %>
</table>
<% if (Session("seqSurvey") <> -1 AND Session("SurveysExist") = "Yes") then %>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table class="borderless" cellSpacing="0" cellPadding="2">
	<tr>
		<th class="boldContent" colSpan="2">Summary Information:</th>
	</tr>
	<tr>
		<td align="right" colSpan="3">Clicked E-mail Link:</td>
		<% if Session("ClickedEmailLink") = -1 then %>
		<td>0</td>
		<% else %>
		<td><%=Session("ClickedEmailLink")%></td>
		<% end if %>
	</tr>
	<tr>
		<td align="right" colSpan="3">Completed Surveys:</td>
		<% if Session("CompletedSurveys") = -1 then %>
		<td>0</td>
		<% else %>
		<td><%=Session("CompletedSurveys")%></td>
		<% end if %>
	</tr>
	<tr>
		<td align="right" colSpan="3">Earliest Data:</td>
		<% if Session("EarliestDate") = "1/1/1970" then %>
		<td>(No Data Available)</td>
		<% else %>
		<td><%=Session("EarliestDate")%></td>
		<% end if %>
	</tr>
	<tr>
		<td align="right" colSpan="3">Most Recent Data:</td>
		<% if Session("LatestDate") = "1/1/1970" then %>
		<td>(No Data Available)</td>
		<% else %>
		<td><%=Session("LatestDate")%></td>
		<% end if %>
	</tr>
</table>
<% end if %>
<% if (Session("SurveysExist") = "Yes" and Session("seqSurvey") <> -1) then %>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table class="borderless" cellSpacing="0" cellPadding="2">
	<TBODY>
		<tr align="left">
			<th colSpan="2">
				Survey Tag Item(s)</th></tr>
		<tr>
			<td><asp:repeater id="rptrTags" runat="server">
					<ItemTemplate>
						<tr>
							<td align="right"><%# Convert.ToString(DataBinder.Eval(Container.DataItem, "SURVEY_TAG_TITLE")) %>:</td>
							<td><%# Convert.ToString(DataBinder.Eval(Container.DataItem, "SURVEY_TAG_ANSWER")) %>
							</td>
						</tr>
					</ItemTemplate>
					<HeaderTemplate>
						<table align="left" class="borderless">
					</HeaderTemplate>
					<FooterTemplate>
                        </table>
                    </FooterTemplate>
                    <SeparatorTemplate></SeparatorTemplate>
                    </asp:repeater>
            </td>
       </tr>
   </TBODY>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<asp:table id="OutputSettings" runat="server" cssclass="borderless" cellPadding="2" cellSpacing="0">
	<asp:TableRow>
		<asp:TableCell Wrap="False" cssclass="boldContent" ColumnSpan="3">
			Report Output Settings - (You may choose more than one)
		</asp:TableCell>
	</asp:TableRow>
</asp:table>
<asp:table id="Table1" runat="server" cssclass="borderless" cellPadding="2" cellSpacing="0">
	<asp:TableRow>
		<asp:TableCell Wrap="False" cssclass="boldContent">
			<asp:image id="hlpTimeSplit" runat="server" ImageUrl="./images/help.png" ToolTip="Select a time split for your output.  Time splits may create multiple output for each question."></asp:image>
		</asp:TableCell>
		<asp:TableCell Wrap="False" cssclass="boldContent" HorizontalAlign="Right">
			Split Output On:
		</asp:TableCell>
		<asp:TableCell Wrap="False" cssclass="boldContent" HorizontalAlign="Left">
			<asp:DropDownList ID="ddlTimeSplit" Runat="server" CssClass="content">
				<asp:ListItem Value="NONE">-- Select --</asp:ListItem>
				<asp:ListItem Value="YEAR">Year</asp:ListItem>
				<asp:ListItem Value="SEMIANNUAL">Semi-Annual</asp:ListItem>
				<asp:ListItem Value="QUARTER">Quarter</asp:ListItem>
				<asp:ListItem Value="MONTH">Month</asp:ListItem>
			</asp:DropDownList>
		</asp:TableCell>
	</asp:TableRow>
	<asp:TableRow>
		<asp:TableCell Wrap="False" cssclass="boldContent">
			<asp:image id="hlpTextOutput" runat="server" ImageUrl="./images/help.png" ToolTip="Check this box to display textual responses from respondents."></asp:image>
		</asp:TableCell>
		<asp:TableCell Wrap="False" cssclass="boldContent" HorizontalAlign="Right">
			Display Text Results?
		</asp:TableCell>
		<asp:TableCell Wrap="False" cssclass="boldContent" HorizontalAlign="Left">
			<asp:CheckBox ID="chkTextOutput" Runat="server" AutoPostBack="False" CssClass="content"></asp:CheckBox>
		</asp:TableCell>
	</asp:TableRow>
	<asp:TableRow>
		<asp:TableCell Wrap="False" cssclass="boldContent">
			<asp:image id="hlpQuestionsPerPage" runat="server" ImageUrl="./images/help.png" ToolTip="Select the number of question results to show per page when printing the report."></asp:image>
		</asp:TableCell>
		<asp:TableCell Wrap="False" cssclass="boldContent" HorizontalAlign="Right">
			Questions Per Page:
		</asp:TableCell>
		<asp:TableCell Wrap="False" cssclass="boldContent" HorizontalAlign="Left">
			<asp:DropDownList ID="ddlQuestionsPerPage" Runat="server" CssClass="content">
				<asp:ListItem Value="1">1</asp:ListItem>
				<asp:ListItem Value="2">2</asp:ListItem>
				<asp:ListItem Value="3">3</asp:ListItem>
			</asp:DropDownList>
		</asp:TableCell>
	</asp:TableRow>
</asp:table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<asp:table id="ReportFilter" runat="server" CssClass="borderless" cellPadding="2" cellSpacing="0">
	<asp:TableRow HorizontalAlign="Left">
		<asp:TableCell ColumnSpan="3" cssclass="boldContent">
			Report Filters - (You may choose more than one)
		</asp:TableCell>
	</asp:TableRow>
	<asp:TableRow>
		<asp:TableCell>
			<asp:image id="hlpDateFilter" runat="server" ImageUrl="./images/help.png" ToolTip="Select a start and end date to limit the results of this report."></asp:image>
		</asp:TableCell>
		<asp:TableCell align="left" Wrap="False" cssclass="boldContent">
			<table cellSpacing="0" cellPadding="2" class="borderless">
				<tr>
					<td align="left" Wrap="False" class="boldContent">
						Start Date:
						<igsch:webdatechooser id="wdcStartDate" runat="server" Font-Names="Verdana" Font-Size="x-small" CalendarJavaScriptFileName="./MyInfragisticsScripts/ig_calendar.js"
							NullDateLabel=" " JavaScriptFileNameCommon="./MyInfragisticsScripts/ig_shared.js" JavaScriptFileName="./MyInfragisticsScripts/ig_webdropdown.js"
							ImageDirectory="./MyInfragisticsImages/">
							<CalendarLayout PrevMonthText="&lt;&lt;" FooterFormat="Today: {0:d}" NextMonthText="&gt;&gt;" ChangeMonthToDateClicked="True"
								MaxDate="" GridLineColor="Transparent" ShowFooter="False">
								<DayStyle BorderWidth="2px" Font-Size="x-small" Font-Names="Verdana" BorderStyle="Ridge" BackColor="Tan"></DayStyle>
								<SelectedDayStyle Font-Size="x-small" Font-Names="Verdana" BackColor="#00C0C0"></SelectedDayStyle>
								<OtherMonthDayStyle Font-Size="x-small" Font-Names="Verdana" ForeColor="#E0E0E0" BackColor="Silver"></OtherMonthDayStyle>
								<NextPrevStyle Font-Size="x-small" Font-Underline="True" Font-Names="Verdana" ForeColor="Yellow"
									BackColor="Silver"></NextPrevStyle>
								<CalendarStyle BorderWidth="2px" Font-Size="x-small" Font-Names="Verdana" BorderStyle="Ridge"></CalendarStyle>
								<WeekendDayStyle Font-Size="x-small" Font-Names="Verdana" ForeColor="#C000C0" BackColor="MistyRose"></WeekendDayStyle>
								<TodayDayStyle Font-Size="x-small" Font-Names="Verdana" BackColor="Lavender"></TodayDayStyle>
								<DayHeaderStyle Width="50px" BorderWidth="2px" Font-Size="x-small" Font-Names="Verdana" BorderStyle="Ridge"
									ForeColor="White" BackColor="Black"></DayHeaderStyle>
								<TitleStyle Font-Size="x-small" Font-Names="Verdana" ForeColor="White" BackColor="Gray"></TitleStyle>
							</CalendarLayout>
						</igsch:webdatechooser>
					</td>
					<td width="20px"></td>
					<td align="left" Wrap="False" class="boldContent">
						End Date:
						<igsch:webdatechooser id="wdcEndDate" runat="server" Font-Names="Verdana" Font-Size="x-small" CalendarJavaScriptFileName="./MyInfragisticsScripts/ig_calendar.js"
							NullDateLabel=" " JavaScriptFileNameCommon="./MyInfragisticsScripts/ig_shared.js" JavaScriptFileName="./MyInfragisticsScripts/ig_webdropdown.js"
							ImageDirectory="./MyInfragisticsImages/">
							<CalendarLayout PrevMonthText="&lt;&lt;" FooterFormat="Today: {0:d}" NextMonthText="&gt;&gt;" ChangeMonthToDateClicked="True"
								MaxDate="" GridLineColor="Transparent" ShowFooter="False">
								<DayStyle BorderWidth="2px" Font-Size="x-small" Font-Names="Verdana" BorderStyle="Ridge" BackColor="Tan"></DayStyle>
								<SelectedDayStyle Font-Size="x-small" Font-Names="Verdana" BackColor="#00C0C0"></SelectedDayStyle>
								<OtherMonthDayStyle Font-Size="x-small" Font-Names="Verdana" ForeColor="#E0E0E0" BackColor="Silver"></OtherMonthDayStyle>
								<NextPrevStyle Font-Size="x-small" Font-Underline="True" Font-Names="Verdana" ForeColor="Yellow"
									BackColor="Silver"></NextPrevStyle>
								<CalendarStyle BorderWidth="2px" Font-Size="x-small" Font-Names="Verdana" BorderStyle="Ridge"></CalendarStyle>
								<WeekendDayStyle Font-Size="x-small" Font-Names="Verdana" ForeColor="#C000C0" BackColor="MistyRose"></WeekendDayStyle>
								<TodayDayStyle Font-Size="x-small" Font-Names="Verdana" BackColor="Lavender"></TodayDayStyle>
								<DayHeaderStyle Width="50px" BorderWidth="2px" Font-Size="x-small" Font-Names="Verdana" BorderStyle="Ridge"
									ForeColor="White" BackColor="Black"></DayHeaderStyle>
								<TitleStyle Font-Size="x-small" Font-Names="Verdana" ForeColor="White" BackColor="Gray"></TitleStyle>
							</CalendarLayout>
						</igsch:webdatechooser>
					</td>
				</tr>
			</table>
		</asp:TableCell>
	</asp:TableRow>
</asp:table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="0" cellPadding="2" class="borderless">
	<TR>
		<TD colSpan="3"><asp:button id="btnSubmit2" Runat="server" CssClass="button" Text="Submit" ToolTip="Submit"></asp:button>&nbsp;
			<asp:button id="btnReset2" Runat="server" CssClass="button" Text="Reset" ToolTip="Reset"></asp:button>&nbsp;
			<asp:label id="lblPopup2" Runat="server" CssClass="ReportHeaderRed">(For enabled popup blockers, use Ctrl+Submit)</asp:label></TD>
	</TR>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<asp:table id="ReportData" runat="server" CssClass="borderless" cellPadding="2" cellSpacing="0">
	<asp:TableRow horizontalalign="left">
		<asp:TableCell columnSpan="3" cssclass="boldContent">
			Report Data - (You may choose more than one)<br />
			<asp:image id="hlpSelectAllData" runat="server" ImageUrl="./images/help.png" ToolTip="Click this button to check or uncheck all of the data checkboxes."></asp:image>
			&nbsp;
			<asp:button id="btnSelectAll" ToolTip="Check All" Runat="server" CssClass="button" Text="Check All"></asp:button>
	</asp:TableCell>
	</asp:TableRow>
</asp:table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="0" cellPadding="2" class="borderless">
	<TR>
		<TD colSpan="3"><asp:button id="btnSubmit" Runat="server" CssClass="button" Text="Submit" ToolTip="Submit"></asp:button>&nbsp;
			<asp:button id="btnReset" Runat="server" CssClass="button" Text="Reset" ToolTip="Reset"></asp:button>&nbsp;
			<asp:label id="lblPopup" Runat="server" CssClass="ReportHeaderRed">(For enabled popup blockers, use Ctrl+Submit)</asp:label></TD>
	</TR>
</table>
<% end if %>
</form>
</asp:Content>
