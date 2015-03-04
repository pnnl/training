<%@ Page trace="false" Language="vb" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" Codebehind="Message.aspx.vb" Inherits="SurveyAdmin.Message" %>
<%@ Register TagPrefix="igtxt" Namespace="Infragistics.WebUI.WebDataInput" Assembly="Infragistics2.WebUI.WebDataInput.v8.2, Version=8.2.20082.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" %>
<%@ Register TagPrefix="igsch" Namespace="Infragistics.WebUI.WebSchedule" Assembly="Infragistics2.WebUI.WebDateChooser.v8.2, Version=8.2.20082.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder" Runat="Server">
<%=Session("JSAction")%>
<form id="Form1" method="post" runat="server">
<h1>Message of the Day</h1>
<p></p>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="0" cellPadding="2" class="borderless">
	<tr class="navPanel">
		<td></td>
		<td noWrap align="right"><asp:label id="lblNavBar" runat="server" CssClass="lblNavPanel"></asp:label></td>
		<td><asp:imagebutton id="btnStart" accessKey="z" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/start_up.gif"
				CssClass="btnSpace" ToolTip="First Message" CausesValidation="False" AlternateText="First Message"></asp:imagebutton><asp:imagebutton id="btnPrev" accessKey="x" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/prev_up.gif"
				CssClass="btnSpace" ToolTip="Previous Message" CausesValidation="False" AlternateText="Previous Message"></asp:imagebutton><asp:imagebutton id="btnReload" accessKey="c" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/ig_menuCurrent_up.png"
				CssClass="btnSpace" ToolTip="Reload Message" CausesValidation="False" AlternateText="Reload Message"></asp:imagebutton><asp:imagebutton id="btnNext" accessKey="v" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/next_up.gif"
				CssClass="btnSpace" ToolTip="Next Message" CausesValidation="False" AlternateText="Next Message"></asp:imagebutton><asp:imagebutton id="btnLast" accessKey="b" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/end_up.gif"
				CssClass="btnSpace" ToolTip="Last Message" CausesValidation="False" AlternateText="Last Message"></asp:imagebutton><asp:imagebutton id="btnNew" accessKey="n" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/add_up.gif"
				CssClass="btnSpace" ToolTip="New Message" CausesValidation="False" AlternateText="New Message"></asp:imagebutton></td>
		<td style="WIDTH: 100%"></td>
	</tr>
	<tr>
		<td style="HEIGHT: 24px"><asp:image id="hlpPrevPassword" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the expiration date or leave this field blank to keep it around indefinitely."></asp:image></td>
		<td style="HEIGHT: 24px" align="right"><asp:label id="lblExpires" Runat="server">Expires:</asp:label></td>
		<td style="HEIGHT: 24px" align="left"><igsch:webdatechooser id="wdcExpirationDate" runat="server" Width="200px" ImageDirectory="./MyInfragisticsImages/"
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
	<tr>
		<td vAlign="top"><asp:image id="Image2" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the message to be displayed on the main page."></asp:image></td>
		<td vAlign="top" align="right"><asp:label id="lblMessage" Runat="server">Message:</asp:label></td>
		<td align="left"><asp:textbox id="txtMessage" runat="server" Width="400px" Height="200px" TextMode="MultiLine"
				MaxLength="4800" Columns="100" Rows="15" Wrap="False"></asp:textbox></td>
	</tr>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table class="borderless">
	<TBODY>
		<tr>
			<td vAlign="top"><asp:image id="hlpMessageSelect" runat="server" ImageUrl="./images/help.png" ToolTip="Select one of the messages to view it."></asp:image></td>
			<td style="WIDTH: 9px" vAlign="top" align="right"><asp:label id="Label2" Runat="server">Messages:</asp:label></td>
			<td vAlign="top" align="left" style="padding: 3px 0 0 5px;"><asp:repeater id="rptrMessage" runat="server">
					<ItemTemplate>
						<li>
							<a href='./Message.aspx?seqMessageID=<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "MESSAGE_ID")) %>'>
								<%#Left(Convert.ToString(DataBinder.Eval(Container.DataItem, "MESSAGE_TEXT")), 20)%>
								- Expires:
								<%#Convert.ToString(DataBinder.Eval(Container.DataItem, "EXPIRATION_DATE").ToShortDateString)%>
							</a>
						</li>
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
	</TBODY>
</table>
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
