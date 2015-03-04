<%@ Register TagPrefix="igsch" Namespace="Infragistics.WebUI.WebSchedule" Assembly="Infragistics2.WebUI.WebDateChooser.v8.2, Version=8.2.20082.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" %>
<%@ Page trace="false" Language="vb" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" Codebehind="SendSurvey.aspx.vb" Inherits="SurveyAdmin.SendSurvey"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder" Runat="Server">
<%=Session("JSAction")%>
<form id="Form1" method="post" runat="server">
<h1>Send Survey</h1>
<p></p>
<table class="borderless">
	<tr>
		<td><asp:image id="hlpSelectSurvey" runat="server" ImageUrl="./images/help.png" ToolTip="Select a survey to send out."></asp:image></td>
		<% if (Session("SurveysExist") = "Yes") then %>
		<td colSpan="4"><asp:dropdownlist id="ddlSurveyList" runat="server" AutoPostBack="True"></asp:dropdownlist></td>
		<% else %>
		<td colSpan="4"><asp:label id="lblSelectSurvey" Runat="server" ForeColor="Red">There are currently no Surveys available for you to send.</asp:label></td>
		<% end if %>
	</tr>
	<% if (Session("SurveysExist") = "Yes") then %>
	<tr>
		<td><asp:image id="hlpchkbxFollowUp" runat="server" ImageUrl="./images/help.png" ToolTip="Check this item to mark this as a follow-up survey allowing an entry for the closing date of this survey to be sent."></asp:image></td>
		<td colSpan="4"><asp:checkbox id="chkbxFollowUp" Runat="server" AutoPostBack="True" CssClass="boldContent" Text="Follow-up?"></asp:checkbox></td>
	</tr>
	<tr>
		<td><asp:image id="hlpchkbxReminder" runat="server" ImageUrl="./images/help.png" ToolTip="Check this item to send reminders to individuals selected below who have yet to respond."></asp:image></td>
		<td colSpan="4"><asp:checkbox id="chkbxReminder" Runat="server" AutoPostBack="False" CssClass="boldContent" Text="Send Reminder?"></asp:checkbox></td>
	</tr>
	<% end if %>
</table>
<% if (Session("seqSurvey") <> -1) then %>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="0" cellPadding="2" class="borderless">
	<% If (CType(Session("isFollowUp"), boolean)) then %>
	<tr>
		<td><asp:image id="hlpFollowUp" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the closing date for this follow-up survey."></asp:image></td>
		<td noWrap align="right"><asp:label id="lblFollowUp" Runat="server" CssClass="boldContent">Closing Date:</asp:label></td>
		<td style="HEIGHT: 24px" align="left"><igsch:webdatechooser id="wdcFollowUp" runat="server" Width="200px" ImageDirectory="./MyInfragisticsImages/"
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
	<% End if %>
	<tr>
		<td><asp:image id="hlpSubject" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the subject for the e-mail that will be sent."></asp:image></td>
		<td noWrap align="right"><asp:label id="lblSubject" Runat="server" CssClass="boldContent"> Subject:</asp:label></td>
		<td align="left"><asp:textbox id="txtSubject" Runat="server" Width="400px" MaxLength="40"></asp:textbox></td>
	</tr>
	<tr>
		<td style="HEIGHT: 207px" vAlign="top"><asp:image id="hlpMessage" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the message for your e-mail."></asp:image></td>
		<td style="HEIGHT: 207px" vAlign="top" noWrap align="right"><asp:label id="lblMessage" Runat="server" CssClass="boldContent"> Message:</asp:label></td>
		<td style="HEIGHT: 207px" align="left"><asp:textbox id="txtMessage" runat="server" Width="400px" MaxLength="8000"
				Height="200px" TextMode="MultiLine" Rows="15" Columns="50"></asp:textbox></td>
	</tr>
</table>
<br>
<asp:button id="btnBuild1" Runat="server" ToolTip="Build" CssClass="button" Text="Build"></asp:button>
<asp:button id="btnSend1" Runat="server" ToolTip="Send" CssClass="button" Text="Send"></asp:button>
<asp:button id="btnClear1" Runat="server" ToolTip="Clear" CssClass="button" Text="Clear"></asp:button>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="5" cellPadding="5" width="100%" align="center" class="borderless">
	<TBODY>
		<TR align="center">
			<TD vAlign="top" colSpan="3">
				<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
				<asp:label id="lblSelectedUsers" runat="server" CssClass="boldContent" ForeColor="Red" Width="100%">Recipients List</asp:label><br>
			</TD>
		</TR>
		<tr>
			<td vAlign="top" width="20"><asp:image id="Image1" runat="server" ImageUrl="./images/help.png" ToolTip="This text box is for display purposes only.  This text box displays the list of recipients you have chosen."></asp:image></td>
			<td vAlign="top" align="center" colSpan="2"><asp:textbox id="txtRecipients" runat="server" Width="600px" Height="200px"
					TextMode="MultiLine" Rows="15" Columns="50" Enabled="True"></asp:textbox></td>
		</tr>
	</TBODY>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="5" cellPadding="5" width="100%" align="left" class="borderless">
	<TBODY>
		<% if (Session("GroupsRows") <> "None") then %>
		<TR align="center">
			<TD vAlign="top" colSpan="3">
				<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
				<asp:label id="lblSurveyGroups" runat="server" CssClass="boldContent" ForeColor="Red" Width="100%">Distribution Groups Set up for this Survey</asp:label><br>
			</TD>
		</TR>
		<tr>
			<td vAlign="top" width="20px"><asp:image id="hlpGroups" runat="server" ImageUrl="./images/help.png" ToolTip="Check groups to add the users contained in them to the mailing list above."></asp:image></td>
			<td vAlign="top" align="left" colSpan="2">
			    <asp:datagrid id="dgSurveyGroups" runat="server" CellPadding="0" CellSpacing="1"
					AutoGenerateColumns="False" Width="600px">
					<AlternatingItemStyle BackColor="#A7DCD2"></AlternatingItemStyle>
					<ItemStyle></ItemStyle>
					<HeaderStyle HorizontalAlign="Left" BorderWidth="5px" BorderStyle="Outset"
						BackColor="#60C0B4"></HeaderStyle>
					<Columns>
						<asp:TemplateColumn HeaderText="&lt;a href=&quot;./SendSurvey.aspx?selectAll=1&quot;&gt;Select All&lt;/a&gt;">
							<HeaderStyle HorizontalAlign="Center" Width="50px"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
							<HeaderTemplate>
								<asp:Button id="Button1" runat="server" ToolTip="Select All Groups" CssClass="button" Text="Select"
									OnCommand="commandBtnClick" CommandName="SelectGroup" CommandArgument="commandBtnClick"></asp:Button>
							</HeaderTemplate>
							<ItemTemplate>
								<asp:CheckBox id="Checkbox1" Runat="server"></asp:CheckBox>
							</ItemTemplate>
						</asp:TemplateColumn>
						<asp:BoundColumn Visible="False" DataField="SURVEY_KEY" HeaderText="SURVEY_KEY"></asp:BoundColumn>
						<asp:BoundColumn Visible="False" DataField="SURVEY_GROUP_ID" SortExpression="USER_KEY" HeaderText="SURVEY_GROUP_ID"></asp:BoundColumn>
						<asp:BoundColumn DataField="GROUP_TITLE" SortExpression="GROUP_TITLE" HeaderText="Group Name">
							<HeaderStyle></HeaderStyle>
						</asp:BoundColumn>
					</Columns>
				</asp:datagrid></td>
		</tr>
		<% else %>
		<tr align="center">
			<td colSpan="3">
				<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
				<asp:label id="lblSurveyGroupsNone" runat="server" CssClass="boldContent" ForeColor="Red" Font-Bold="True"
					BackColor="LightGrey">There are currently no distribution groups set up for this survey.</asp:label></td>
		</tr>
		<% end if %>
	</TBODY>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="5" cellPadding="5" width="100%" align="center" class="borderless">
	<TBODY>
		<TR align="center">
			<TD vAlign="top" colSpan="3">
				<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
				<asp:label id="lblSurveyRespondents" runat="server" CssClass="boldContent" ForeColor="Red"
					Width="100%">All Available Survey Respondents</asp:label></TD>
		</TR>
		<tr class="navPanel">
			<td class="navPanelLeft"></td>
			<td vAlign="middle" align="center" colSpan="3"><asp:imagebutton id="imgA1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/A.png" ToolTip="List all last names starting with A"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="A" AlternateText="A" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgB1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/B.png" ToolTip="List all last names starting with B"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="B" AlternateText="B" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgC1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/C.png" ToolTip="List all last names starting with C"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="C" AlternateText="C" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgD1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/D.png" ToolTip="List all last names starting with D"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="D" AlternateText="D" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgE1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/E.png" ToolTip="List all last names starting with E"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="E" AlternateText="E" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgF1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/F.png" ToolTip="List all last names starting with F"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="F" AlternateText="F" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgG1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/G.png" ToolTip="List all last names starting with G"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="G" AlternateText="G" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgH1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/H.png" ToolTip="List all last names starting with H"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="H" AlternateText="H" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgI1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/I.png" ToolTip="List all last names starting with I"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="I" AlternateText="I" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgJ1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/J.png" ToolTip="List all last names starting with J"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="J" AlternateText="J" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgK1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/K.png" ToolTip="List all last names starting with K"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="K" AlternateText="K" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgL1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/L.png" ToolTip="List all last names starting with L"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="L" AlternateText="L" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgM1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/M.png" ToolTip="List all last names starting with M"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="M" AlternateText="M" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgN1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/N.png" ToolTip="List all last names starting with N"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="N" AlternateText="N" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgO1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/O.png" ToolTip="List all last names starting with O"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="O" AlternateText="O" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgP1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/P.png" ToolTip="List all last names starting with P"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="P" AlternateText="P" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgQ1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/Q.png" ToolTip="List all last names starting with Q"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="Q" AlternateText="Q" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgR1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/R.png" ToolTip="List all last names starting with R"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="R" AlternateText="R" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgS1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/S.png" ToolTip="List all last names starting with S"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="S" AlternateText="S" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgT1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/T.png" ToolTip="List all last names starting with T"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="T" AlternateText="T" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgU1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/U.png" ToolTip="List all last names starting with U"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="U" AlternateText="U" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgV1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/V.png" ToolTip="List all last names starting with V"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="V" AlternateText="V" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgW1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/W.png" ToolTip="List all last names starting with W"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="W" AlternateText="W" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgX1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/X.png" ToolTip="List all last names starting with X"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="X" AlternateText="X" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgY1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/Y.png" ToolTip="List all last names starting with Y"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="Y" AlternateText="Y" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgZ1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/Z.png" ToolTip="List all last names starting with Z"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="Z" AlternateText="Z" CausesValidation="False"></asp:imagebutton></td>
		</tr>
		<% if (Session("RespondentRows") <> "None") then %>
		<tr>
			<td vAlign="top" width="20"><asp:image id="hlpSurveyRespondents" runat="server" ImageUrl="./images/help.png" ToolTip="Select respondents to add them to the send list."></asp:image></td>
			<td vAlign="top" align="left" colSpan="2">
			    <asp:datagrid id="dgSurveyRespondents" runat="server" CellPadding="0" CellSpacing="1"
					AutoGenerateColumns="False" Width="100%">
					<AlternatingItemStyle BackColor="#A7DCD2"></AlternatingItemStyle>
					<HeaderStyle HorizontalAlign="Left" BorderWidth="5px" BorderStyle="Outset"
						BackColor="#60C0B4"></HeaderStyle>
					<Columns>
						<asp:TemplateColumn HeaderText="&lt;a href=&quot;./SendSurvey.aspx?selectAll=1&quot;&gt;Select All&lt;/a&gt;">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
							<HeaderTemplate>
								<asp:Button id="Button2" runat="server" ToolTip="Select All Users" CssClass="button" Text="Select"
									OnCommand="commandBtnClick" CommandName="SelectUsers" CommandArgument="commandBtnClick"></asp:Button>
							</HeaderTemplate>
							<ItemTemplate>
								<asp:CheckBox id="Checkbox2" Runat="server"></asp:CheckBox>
							</ItemTemplate>
						</asp:TemplateColumn>
						<asp:BoundColumn Visible="False" DataField="USER_KEY" SortExpression="USER_KEY" HeaderText="USER_KEY"></asp:BoundColumn>
						<asp:BoundColumn Visible="False" DataField="USER_CODE" SortExpression="USER_CODE" HeaderText="USER_CODE"></asp:BoundColumn>
						<asp:BoundColumn DataField="Name" SortExpression="Name" HeaderText="User Name"></asp:BoundColumn>
						<asp:BoundColumn DataField="HANFORD_ID" SortExpression="HANFORD_ID" HeaderText="Hanford ID">
                            <HeaderStyle Font-Bold="False" Font-Italic="False" Font-Overline="False" Font-Strikeout="False"
                                Font-Underline="False" Wrap="False" />
                        </asp:BoundColumn>
						<asp:BoundColumn DataField="EMAIL_ADDRESS" SortExpression="EMAIL_ADDRESS" HeaderText="Email Address">
                            <ItemStyle Font-Bold="False" Font-Italic="False" Font-Overline="False"
                                Font-Strikeout="False" Font-Underline="False" />
                        </asp:BoundColumn>
						<asp:BoundColumn DataField="CREATE_DATE" SortExpression="CREATE_DATE" HeaderText="Sent"></asp:BoundColumn>
						<asp:BoundColumn DataField="LAST_COMPLETION_DATE" SortExpression="LAST_COMPLETION_DATE" HeaderText="Response"></asp:BoundColumn>
						<asp:BoundColumn DataField="REMIND_DATE" SortExpression="REMIND_DATE" HeaderText="Remind"></asp:BoundColumn>
					</Columns>
				</asp:datagrid>
			</td>
		</tr>
		<% else %>
		<tr align="center">
			<td width="20"></td>
			<td colSpan="2"><asp:label id="lblSurveyRespondentsNone" runat="server" CssClass="boldContent" ForeColor="Red"
					Font-Bold="True" BackColor="LightGrey">There are no users with a last name beginning with the letter specified.<br />If you just arrived on this page, then there are no persons in the system with a last name beginning with 'A'.</asp:label></td>
		</tr>
		<% end if %>
		<tr class="navPanel">
			<td class="navPanelLeft"></td>
			<td vAlign="middle" align="center" colSpan="3"><asp:imagebutton id="imgA2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/A.png" ToolTip="List all last names starting with A"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="A" AlternateText="A" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgB2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/B.png" ToolTip="List all last names starting with B"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="B" AlternateText="B" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgC2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/C.png" ToolTip="List all last names starting with C"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="C" AlternateText="C" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgD2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/D.png" ToolTip="List all last names starting with D"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="D" AlternateText="D" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgE2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/E.png" ToolTip="List all last names starting with E"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="E" AlternateText="E" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgF2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/F.png" ToolTip="List all last names starting with F"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="F" AlternateText="F" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgG2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/G.png" ToolTip="List all last names starting with G"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="G" AlternateText="G" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgH2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/H.png" ToolTip="List all last names starting with H"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="H" AlternateText="H" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgI2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/I.png" ToolTip="List all last names starting with I"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="I" AlternateText="I" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgJ2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/J.png" ToolTip="List all last names starting with J"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="J" AlternateText="J" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgK2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/K.png" ToolTip="List all last names starting with K"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="K" AlternateText="K" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgL2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/L.png" ToolTip="List all last names starting with L"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="L" AlternateText="L" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgM2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/M.png" ToolTip="List all last names starting with M"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="M" AlternateText="M" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgN2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/N.png" ToolTip="List all last names starting with N"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="N" AlternateText="N" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgO2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/O.png" ToolTip="List all last names starting with O"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="O" AlternateText="O" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgP2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/P.png" ToolTip="List all last names starting with P"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="P" AlternateText="P" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgQ2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/Q.png" ToolTip="List all last names starting with Q"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="Q" AlternateText="Q" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgR2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/R.png" ToolTip="List all last names starting with R"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="R" AlternateText="R" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgS2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/S.png" ToolTip="List all last names starting with S"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="S" AlternateText="S" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgT2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/T.png" ToolTip="List all last names starting with T"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="T" AlternateText="T" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgU2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/U.png" ToolTip="List all last names starting with U"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="U" AlternateText="U" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgV2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/V.png" ToolTip="List all last names starting with V"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="V" AlternateText="V" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgW2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/W.png" ToolTip="List all last names starting with W"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="W" AlternateText="W" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgX2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/X.png" ToolTip="List all last names starting with X"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="X" AlternateText="X" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgY2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/Y.png" ToolTip="List all last names starting with Y"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="Y" AlternateText="Y" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgZ2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/Z.png" ToolTip="List all last names starting with Z"
					CssClass="btnSpace" OnCommand="commandGridUserList" CommandName="Z" AlternateText="Z" CausesValidation="False"></asp:imagebutton></td>
		</tr>
	</TBODY>
</table>
<HR style="LEFT: 10px; BORDER-TOP-STYLE: ridge; TOP: 480px" width="100%" SIZE="3">
<TABLE style="LEFT: 10px; TOP: 536px" width="100%" align="center" class="borderless">
	<tr>
	    <td>
		    <asp:button id="btnBuild2" Runat="server" ToolTip="Build" CssClass="button" Text="Build"></asp:button>
		    <asp:button id="btnSend2" Runat="server" ToolTip="Send" CssClass="button" Text="Send"></asp:button>
		    <asp:button id="btnClear2" Runat="server" ToolTip="Clear" CssClass="button" Text="Clear"></asp:button>
	    </td>   
	</tr>
</TABLE>
<p></p>
<br>
<% End if %>
</form>
</asp:Content>