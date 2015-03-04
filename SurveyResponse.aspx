<%@ Register TagPrefix="igsch" Namespace="Infragistics.WebUI.WebSchedule" Assembly="Infragistics2.WebUI.WebDateChooser.v8.2, Version=8.2.20082.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" %>
<%@ Register TagPrefix="igtxt" Namespace="Infragistics.WebUI.WebDataInput" Assembly="Infragistics2.WebUI.WebDataInput.v8.2, Version=8.2.20082.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" %>
<%@ Page trace="false" Language="vb" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" Codebehind="SurveyResponse.aspx.vb" Inherits="SurveyAdmin.SurveyResponse" enableViewState="True"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder" Runat="Server">
<%=Session("JSAction")%>
<form id="Form1" method="post" runat="server">
<h1>Survey Input</h1>
<p></p>
<table cellSpacing="0" cellPadding="2" class="borderless">
	<tr>
		<td><asp:image id="hlpInputSwitch" runat="server" ImageUrl="./images/help.png" ToolTip="Check this box to switch between entering survey or non-respondent data."></asp:image></td>
		<td><asp:checkbox id="chkInputSwitch" runat="server" AutoPostBack="True" CssClass="content"></asp:checkbox></td>
	</tr>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="0" cellPadding="2" class="borderless">
	<% if (Session("TemplatesExist") = "Yes") then %>
	<tr>
		<td><asp:image id="hlpTemplateSelect" runat="server" ImageUrl="./images/help.png" ToolTip="Select a template to limit the surveys in the list below to only those within a particular template."></asp:image></td>
		<td align="right"><asp:label id="lblTemplateSelectlblExpires" Runat="server" CssClass="boldContent">Template:</asp:label></td>
		<td align="left"><asp:dropdownlist id="ddlTemplateList" runat="server" AutoPostBack="True" CssClass="content"></asp:dropdownlist></td>
	</tr>
	<% if (Session("SurveysExist") = "Yes") then %>
	<tr>
		<td><asp:image id="hlpSurveySelect" runat="server" ImageUrl="./images/help.png" ToolTip="Select a survey to begin setting up the parameters for the report."></asp:image></td>
		<td align="right"><asp:label id="lblSurveySelect" Runat="server" CssClass="boldContent">Survey:</asp:label></td>
		<td align="left"><asp:dropdownlist id="ddlSurveyList" runat="server" AutoPostBack="True" CssClass="content"></asp:dropdownlist></td>
	</tr>
	<% else %>
	<tr>
		<td colSpan="3"><asp:label id="lblNoSurveys" Runat="server" CssClass="boldContent" ForeColor="Red">There are currently either no surveys that you have access to or that have their tag items completely filled out, for this template<br />Please create a survey for this template and/or fill out the tag items for the survey you wish to use.<br />&nbsp;&nbsp;or get access to a survey under this template first.</asp:label></td>
	</tr>
	<% end if %>
	<% else %>
	<tr>
		<td colSpan="3"><asp:label id="lblNoTemplates" Runat="server" CssClass="boldContent" ForeColor="Red">There are currently no templates that you have access to or that have been fully approved for distribution.<br />  Please create a template, get it approved, and then generate a survey from the template.</asp:label></td>
	</tr>
	<% end if %>
</table>
<% if (Session("SurveysExist") = "Yes" and Session("seqSurvey") <> -1) then %>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="0" cellPadding="2" class="borderless">
	<TBODY>
		<tr align="left">
			<td vAlign="top" width="10"><asp:image id="hlpCompletedDate" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the date this survey was completed."></asp:image></td>
			<td vAlign="top" noWrap align="right"><asp:label id="lblCompletedDate" Runat="server" CssClass="boldContent">Date Completed:</asp:label></td>
			<td align="left"><igsch:webdatechooser id="wdcCompleteDate" runat="server" ImageDirectory="./MyInfragisticsImages/" JavaScriptFileName="./MyInfragisticsScripts/ig_webdropdown.js"
					JavaScriptFileNameCommon="./MyInfragisticsScripts/ig_shared.js" NullDateLabel=" " CalendarJavaScriptFileName="./MyInfragisticsScripts/ig_calendar.js"
					Font-Size="x-small" Font-Names="Verdana" Width="200px">
					<CalendarLayout PrevMonthText="&lt;&lt;" NextMonthText="&gt;&gt;" ChangeMonthToDateClicked="True"
						GridLineColor="Transparent" ShowFooter="False">
						<DayStyle BorderWidth="2px" Font-Size="X-Small" Font-Names="Verdana" BorderStyle="Ridge" BackColor="Tan"></DayStyle>
						<SelectedDayStyle Font-Size="X-Small" Font-Names="Verdana" BackColor="#00C0C0"></SelectedDayStyle>
						<OtherMonthDayStyle Font-Size="X-Small" Font-Names="Verdana" ForeColor="#E0E0E0" BackColor="Silver"></OtherMonthDayStyle>
						<NextPrevStyle Font-Size="X-Small" Font-Underline="True" Font-Names="Verdana" ForeColor="Yellow"
							BackColor="Silver"></NextPrevStyle>
						<CalendarStyle BorderWidth="2px" Font-Size="X-Small" Font-Names="Verdana" BorderStyle="Ridge"></CalendarStyle>
						<WeekendDayStyle Font-Size="X-Small" Font-Names="Verdana" ForeColor="#C000C0" BackColor="MistyRose"></WeekendDayStyle>
						<TodayDayStyle Font-Size="X-Small" Font-Names="Verdana" BackColor="Lavender"></TodayDayStyle>
						<DayHeaderStyle Width="50px" BorderWidth="2px" Font-Size="X-Small" Font-Names="Verdana" BorderStyle="Ridge"
							ForeColor="White" BackColor="Black"></DayHeaderStyle>
						<TitleStyle Font-Size="X-Small" Font-Names="Verdana" ForeColor="White" BackColor="Gray"></TitleStyle>
					</CalendarLayout>
				</igsch:webdatechooser></td>
		</tr>
		<tr align="left">
			<th colSpan="3">
				Survey Tag Item(s)</th></tr>
		<tr>
			<td colSpan="3"><asp:repeater id="rptrTags" runat="server" >
					<ItemTemplate>
						<tr>
							<td align="right"><%#Convert.ToString(DataBinder.Eval(Container.DataItem, "SURVEY_TAG_TITLE"))%>:</td>
							<td><%#Convert.ToString(DataBinder.Eval(Container.DataItem, "SURVEY_TAG_ANSWER"))%>
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
<table cellSpacing="0" class="borderless">
	<% if (Session("QuestionAvailability") <> "None" and Session("InputSwitch") = "Respondent") then %>
	<tr style="WIDTH: 100%; HEIGHT: 100%">
		<td><asp:table id="QuestionTable" Runat="server" CssClass="borderless" Width="100%"></asp:table></td>
	</tr>
	<tr>
		<td style="HEIGHT: 26px" align="center"><asp:button id="btnSubmit" Runat="server" ToolTip="Submit" CssClass="button" Text="Submit Page"></asp:button>&nbsp;
			<asp:button id="btnReset" Runat="server" ToolTip="Reset" CssClass="button" Text="Reset Page"></asp:button></td>
	</tr>
	<% else if (Session("InputSwitch") = "Respondent")%>
	<tr align="center">
		<td colSpan="2">
			<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
			<asp:label id="lblNoQuestions" runat="server" CssClass="boldContent" ForeColor="Red" Font-Bold="True"
				BackColor="LightGrey">There are currently no questions in this template.  If you feel that this is
							an error please contact the 
							<a title="Sned mail to the Survey Tool Administrator" href="mailto:gene.gower@pnl.gov?subject=Template Question Error">
					Survey Administrator</a>.</asp:label></td>
	</tr>
	<% else %>
	<tr>
		<td><asp:image id="hlpNonRespondents" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the number of non-resondents."></asp:image></td>
		<td noWrap align="right"><asp:label id="lblNonRespondents" Runat="server" CssClass="boldContent">Number of Non-Respondents:</asp:label></td>
		<td align="left"><igtxt:webnumericedit id="wneNonRespondent" runat="server" ToolTip="Enter the number of non-respondents."
				CssClass="content" ImageDirectory="./MyInfragisticsScripts/Images/" JavaScriptFileName="./MyInfragisticsScripts/ig_edit.js"
				JavaScriptFileNameCommon="./MyInfragisticsScripts/ig_shared.js" Width="152px" DataMode="Int" Section508Compliant="True"
				MaxValue="1000" MinValue="0"></igtxt:webnumericedit></td>
	</tr>
	<tr>
		<td style="HEIGHT: 26px" align="center" colSpan="3"><asp:button id="btnSubmit2" Runat="server" ToolTip="Submit" CssClass="button" Text="Submit Page"></asp:button>&nbsp;
			<asp:button id="btnReset2" Runat="server" ToolTip="Reset" CssClass="button" Text="Reset Page"></asp:button></td>
	</tr>
	<% end if %>
</table>
<% end if %>
</form>
</asp:Content>
