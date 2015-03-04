<%@ Register TagPrefix="igsch" Namespace="Infragistics.WebUI.WebSchedule" Assembly="Infragistics2.WebUI.WebDateChooser.v8.2, Version=8.2.20082.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" %>
<%@ Register TagPrefix="igtxt" Namespace="Infragistics.WebUI.WebDataInput" Assembly="Infragistics2.WebUI.WebDataInput.v8.2, Version=8.2.20082.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" %>
<%@ Page trace="false" Language="vb" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" Codebehind="Survey.aspx.vb" Inherits="SurveyAdmin.Survey" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder" Runat="Server">
<%=Session("JSAction")%>
<form id="Form1" method="post" runat="server">
<h1>Template Surveys</h1>
<p></p>
<script type="text/javascript">
	function isOwner(Owner)
	{
		if(Owner == "True"){
			document.write('<img src="./images/Owner.png" alt="This user is an Owner of this Survey.">');}
	}
</script>
<table class="borderless">
	<tr>
		<td align="left"><asp:label id="lblTemplate" Runat="server" CssClass="boldContent">Template:</asp:label></td>
		<td class="content" align="left"><%=Session("TemplateName")%></td>
		<td align="center" width="50"></td>
		<td align="right"><asp:image id="hlpTemplatePreview" runat="server" ImageUrl="./images/help.png" ToolTip="Click this button to preview the template."></asp:image></td>
		<td align="left"><asp:button id="btnTemplatePreview" runat="server" ToolTip="Preview" CssClass="button" Text="Preview"></asp:button></td>
	</tr>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="0" cellPadding="2" class="borderless">
	<tr class="navPanel">
		<td></td>
		<td noWrap align="right"><asp:label id="lblNavBar" runat="server" CssClass="lblNavPanel"></asp:label></td>
		<td><asp:imagebutton id="btnStart" accessKey="z" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/start_up.gif"
				AlternateText="First Survey" ToolTip="First Survey" CssClass="btnSpace" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnPrev" accessKey="x" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/prev_up.gif"
				AlternateText="Previous Survey" ToolTip="Previous Survey" CssClass="btnSpace" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnReload" accessKey="c" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/ig_menuCurrent_up.png"
				AlternateText="Reload Survey" ToolTip="Reload Survey" CssClass="btnSpace" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnNext" accessKey="v" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/next_up.gif"
				AlternateText="Next Survey" ToolTip="Next Survey" CssClass="btnSpace" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnLast" accessKey="b" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/end_up.gif"
				AlternateText="Last Survey" ToolTip="Last Survey" CssClass="btnSpace" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnNew" accessKey="n" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/add_up.gif"
				AlternateText="New Survey" ToolTip="New Survey" CssClass="btnSpace" CausesValidation="False"></asp:imagebutton></td>
		<td style="WIDTH: 100%" align="right"><asp:button id="btnTop" runat="server" ToolTip="Back to Top" CssClass="small" Text="^ Top ^"
				ForeColor="Blue" Font-Underline="True" BorderStyle="None" BackColor="Transparent"></asp:button></td>
	</tr>
	<tr>
		<td><asp:image id="hlpSurveyName" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the name of your survey."></asp:image></td>
		<td noWrap align="right"><asp:label id="lblSurveyName" Runat="server" CssClass="boldContent">Survey Name:</asp:label></td>
		<td align="left"><asp:textbox id="txtSurveyName" runat="server" CssClass="content" MaxLength="1800" Width="300px"></asp:textbox></td>
	</tr>
	<tr>
		<td vAlign="top"><asp:image id="hlpClosingDate" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the date this survey will expire."></asp:image></td>
		<td vAlign="top" noWrap align="right"><asp:label id="lblClosingDate" Runat="server" CssClass="boldContent">Closing Date:</asp:label></td>
		<td align="left"><igsch:webdatechooser id="wdcClosingDate" runat="server" Width="200px" ImageDirectory="./MyInfragisticsImages/"
			CalendarJavaScriptFileName="./MyInfragisticsScripts/ig_calendar.js" JavaScriptFileName="./MyInfragisticsScripts/ig_webdropdown.js" JavaScriptFileNameCommon="./MyInfragisticsScripts/ig_shared.js"
			NullDateLabel=" " Font-Size="Small" Font-Names="Verdana">
			<CalendarLayout PrevMonthText="&lt;&lt;" NextMonthText="&gt;&gt;" ChangeMonthToDateClicked="True" GridLineColor="Transparent" ShowFooter="False">
				<DayStyle BorderWidth="2px" Font-Size="Small" Font-Names="Verdana" BorderStyle="Ridge" BackColor="Tan"></DayStyle>
				<SelectedDayStyle Font-Size="Small" Font-Names="Verdana" BackColor="#00C0C0"></SelectedDayStyle>
				<OtherMonthDayStyle Font-Size="Small" Font-Names="Verdana" ForeColor="#E0E0E0" BackColor="Silver"></OtherMonthDayStyle>
				<NextPrevStyle Font-Size="Small" Font-Underline="True" Font-Names="Verdana" ForeColor="Yellow"
					BackColor="Silver"></NextPrevStyle>
				<CalendarStyle BorderWidth="2px" Font-Size="Small" Font-Names="Verdana" BorderStyle="Ridge" BackColor="Silver"></CalendarStyle>
				<WeekendDayStyle Font-Size="Medium" Font-Names="Verdana" ForeColor="#C000C0" BackColor="MistyRose"></WeekendDayStyle>
				<TodayDayStyle Font-Size="Small" Font-Names="Verdana" BackColor="Lavender"></TodayDayStyle>
				<DayHeaderStyle Width="50px" BorderWidth="2px" Font-Size="Small" Font-Names="Verdana" BorderStyle="Ridge"
					ForeColor="White" BackColor="Black"></DayHeaderStyle>
				<TitleStyle Font-Size="Small" Font-Names="Verdana" ForeColor="White" BackColor="Gray"></TitleStyle>
			</CalendarLayout>
			<DropDownStyle BorderWidth="1px" BorderColor="Black" BorderStyle="Solid"></DropDownStyle>
		</igsch:webdatechooser></td>
	</tr>
	<TR>
		<TD vAlign="top"><asp:image id="hlpSurveyID" runat="server" ImageUrl="./images/help.png" ToolTip="Enter an identifier for this survey."></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblSurveyID" Runat="server" CssClass="boldContent">Survey ID:</asp:label></TD>
		<TD align="left"><asp:TextBox ID="txtSurveyID" runat="server" CssClass="content" MaxLength="16" Width="200px" style="text-align:right;"></asp:TextBox></TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpCourseNumber" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the course number that this survey belongs to."></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblCourseNumber" Runat="server" CssClass="boldContent">Course Number:</asp:label></TD>
		<TD align="left"><igtxt:webnumericedit id="wneCourseNumber" runat="server" CssClass="content" Width="200px" ImageDirectory="./MyInfragisticsImages/images/"
				JavaScriptFileName="./MyInfragisticsScripts/ig_edit.js" JavaScriptFileNameCommon="./MyInfragisticsScripts/ig_shared.js"
				MinValue="0" DataMode="Int" Section508Compliant="True" MaxValue="999999"></igtxt:webnumericedit></TD>
	</TR>
	<TR>
		<TD style="HEIGHT: 27px" vAlign="top"><asp:image id="hlpColorPicker" runat="server" ImageUrl="./images/help.png" ToolTip="Select a background color from the drop-dwon list for your survey."></asp:image></TD>
		<TD style="HEIGHT: 27px" vAlign="top" noWrap align="right"><asp:label id="lblColorPicker" Runat="server" CssClass="boldContent">Survey Color:</asp:label></TD>
		<TD style="HEIGHT: 27px" align="left"><SELECT class="content" id="ddlColorPicker" runat="server"></SELECT>
		</TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpOptional" runat="server" ImageUrl="./images/help.png" ToolTip="Check this box to make the survey optional."></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:checkbox id="chkOptional" runat="server" CssClass="content"></asp:checkbox></TD>
		<TD align="left"><asp:label id="lblOptional" Runat="server" CssClass="boldContent">Optional?</asp:label></TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpActive" runat="server" ImageUrl="./images/help.png" ToolTip="Check this box to make the survey deliverable."></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:checkbox id="chkActive" runat="server" CssClass="content"></asp:checkbox></TD>
		<TD align="left"><asp:label id="lblActive" Runat="server" CssClass="boldContent">Active?</asp:label></TD>
	</TR>
</table>
<% if (Session("seqSurvey") <> -1) then %>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table class="borderless">
	<tr>
		<TD vAlign="top"><asp:image id="hlpSurveys" runat="server" ImageUrl="./images/help.png" ToolTip="Lists the tag items for the survey.  Note that these must be filled in before the survey can be sent."></asp:image></TD>
		<td vAlign="top" noWrap><asp:label id="lblTagItems" Runat="server" CssClass="boldContent">Tag Items:&nbsp;&nbsp;</asp:label></td>
		<td noWrap style="padding-left:20px;"><asp:repeater id="rptrTags" runat="server">
				<ItemTemplate>
					<% if (((Session("UserType") >= 2 and (CType(Session("isSurveyOwner"), Boolean) or CType(Session("isSurveyDelegate"), Boolean)) and Session("Results") <= 0) or Session("UserType") = 4)) then %>
					<li>
						<a title='<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "SURVEY_TAG_TITLE")) %>' href='./SurveyTag.aspx?seqTemplate=<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "TEMPLATE_KEY")) %>
						&amp;seqSurvey=<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "SURVEY_KEY")) %>
						&amp;intCommentID=<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "SURVEY_TAG_ID ")) %>'>
							<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "SURVEY_TAG_TITLE")) %>
						</a>
					</li>
					<% else %>
					<li>
						<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "SURVEY_TAG_TITLE")) %>
					</li>
					<% end if %>
				</ItemTemplate>
				<HeaderTemplate>
					<ol>
				</HeaderTemplate>
				<FooterTemplate>
					</ol>
				</FooterTemplate>
				<SeparatorTemplate>
				</SeparatorTemplate>
			</asp:repeater></td>
	</tr>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table class="borderless">
	<tr>
		<TD vAlign="top"><asp:image id="hlpDelegates" runat="server" ImageUrl="./images/help.png" ToolTip="Delegates are those that have the ability to alter your Survey.  You may also assign ownership of this Survey to others.  Owners will have full access to this Survey and may remove it. "></asp:image></TD>
		<td vAlign="top" noWrap><asp:label id="lblDelegates" Runat="server" CssClass="boldContent">Delegates:&nbsp;&nbsp;</asp:label></td>
		<td noWrap style="padding-left:20px;"><asp:repeater id="rptrDelegates" runat="server">
				<ItemTemplate>
					<% if ((((CType(Session("isSurveyOwner"), Boolean)) and Session("UserType") <> 1) or Session("UserType") = 4)) then %>
					<li>
						<a title='<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "FIRST_NAME")) %>	<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "MIDDLE_NAME")) %> <%# Convert.ToString(DataBinder.Eval(Container.DataItem, "LAST_NAME")) %> - <%# Convert.ToString(DataBinder.Eval(Container.DataItem, "HANFORD_ID")) %> - <%# Convert.ToString(DataBinder.Eval(Container.DataItem, "EMAIL_ADDRESS")) %>' href='./SurveyDelegates.aspx?seqSurvey=<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "SURVEY_KEY")) %>&amp;seqSurveyUserID=<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "USER_KEY")) %>'>
							<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "FIRST_NAME")) %>
							<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "MIDDLE_NAME")) %>
							<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "LAST_NAME")) %>
							-
							<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "HANFORD_ID")) %>
							-
							<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "EMAIL_ADDRESS")) %>
						</a>
						<script>isOwner("<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "OWNER_IND")) %>")</script>
					</li>
					<% else %>
					<li>
						<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "FIRST_NAME")) %>
						<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "MIDDLE_NAME")) %>
						<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "LAST_NAME")) %>
						-
						<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "HANFORD_ID")) %>
						-
						<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "EMAIL_ADDRESS")) %>
						<script>isOwner("<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "OWNER_IND")) %>")</script>
					</li>
					<% end if %>
				</ItemTemplate>
				<HeaderTemplate>
					<ol>
				</HeaderTemplate>
				<FooterTemplate>
					<% if ((((CType(Session("isSurveyOwner"), Boolean)) and Session("UserType") >= 2) or Session("UserType") = 4)) then %>
					<li>
						<a title="Create New Delegate" href='./SurveyDelegates.aspx?seqSurvey=<%=Session("seqSurvey")%>&amp;seqSurveyUserID=-1'>
							Create New Delegate</a>
					</li>
					<% end if %>
					</ol>
				</FooterTemplate>
				<SeparatorTemplate>
				</SeparatorTemplate>
			</asp:repeater></td>
	</tr>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table class="borderless">
	<tr>
		<TD vAlign="top"><asp:image id="hlpSurveyGroups" runat="server" ImageUrl="./images/help.png" ToolTip="Lists the survey distribution groups for this survey."></asp:image></TD>
		<td vAlign="top" noWrap><asp:label id="lblSurveyGroups" Runat="server" CssClass="boldContent">Distribution<br />&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Groups:</asp:label></td>
		<td noWrap style="padding-left:20px;"><asp:repeater id="rptrSurveyGroups" runat="server">
				<ItemTemplate>
					<% if (((Session("UserType") >= 2 and (CType(Session("isSurveyOwner"), Boolean) or CType(Session("isSurveyDelegate"), Boolean))) or Session("UserType") = 4)) then %>
					<li>
						<a title='<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "GROUP_TITLE")) %>' href='./Groups.aspx?seqSurvey=<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "SURVEY_KEY")) %>&amp;intGroupID=<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "SURVEY_GROUP_ID")) %>'>
							<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "GROUP_TITLE")) %>
						</a>
					</li>
					<% else %>
					<li>
						<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "GROUP_TITLE")) %>
					</li>
					<% end if %>
				</ItemTemplate>
				<HeaderTemplate>
					<ol>
				</HeaderTemplate>
				<FooterTemplate>
					<% if (((Session("UserType") >= 2 and (CType(Session("isSurveyOwner"), Boolean) or CType(Session("isSurveyDelegate"), Boolean))) or Session("UserType") = 4)) then %>
					<li>
						<a title="New Group" href='./Groups.aspx?seqSurvey=<%=Session("seqSurvey")%>&amp;intGroupID=-1'>
							Create New Group</a>
					</li>
					<% end if %>
					</ol>
				</FooterTemplate>
				<SeparatorTemplate>
				</SeparatorTemplate>
			</asp:repeater></td>
	</tr>
</table>
<% end if %>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table width="100%" align="center" class="borderless">
	<tr>
		<td>
		    <% If Session("PageOptions") <> 2 Then%>
		        <asp:button id="btnInsert" Runat="server" CssClass="button" ToolTip="Insert" Text="Insert"></asp:button>
		    <% end if %>
		    <% if session("PageOptions") <> 1 then %>
		        <asp:button id="btnUpdate" Runat="server" CssClass="button" ToolTip="Update" Text="Update"></asp:button>
		        <asp:button id="btnDelete" Runat="server" CssClass="button" ToolTip="Delete" Text="Delete"></asp:button>
		    <% end if %>
		    <asp:button id="btnReset" Runat="server" CssClass="button" ToolTip="Reset" Text="Reset"></asp:button>
		</td>
	</tr>
</table>
</form>
</asp:Content>