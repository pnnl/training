<%@ Page trace="false" Language="vb" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" Codebehind="TemplateDelegates.aspx.vb" Inherits="SurveyAdmin.TemplateDelegates"  %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder" Runat="Server">
<%=Session("JSAction")%>
<form id="Form1" method="post" runat="server">
<h1>Template Delegates</h1>
<p></p>
<table class="borderless">
	<tr>
		<td align="left"><asp:label id="lblTemplate" Runat="server" CssClass="content">Template:</asp:label></td>
		<td class="content" align="left"><%=Session("TemplateName")%></td>
		<td align="center" width="50"></td>
		<td align="right"></td>
		<td align="left"></td>
	</tr>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="0" cellPadding="2" class="borderless">
	<tr class="navPanel">
		<td></td>
		<td noWrap align="right"><asp:label id="lblNavBar" runat="server" CssClass="lblNavPanel"></asp:label></td>
		<td><asp:imagebutton id="btnStart" accessKey="z" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/start_up.gif"
				CssClass="btnSpace" ToolTip="First Delegate" CausesValidation="False" AlternateText="First Delegate"></asp:imagebutton><asp:imagebutton id="btnPrev" accessKey="x" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/prev_up.gif"
				CssClass="btnSpace" ToolTip="Previous Delegate" CausesValidation="False" AlternateText="Previous Delegate"></asp:imagebutton><asp:imagebutton id="btnReload" accessKey="c" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/ig_menuCurrent_up.png"
				CssClass="btnSpace" ToolTip="Reload Delegate" CausesValidation="False" AlternateText="Reload Delegate"></asp:imagebutton><asp:imagebutton id="btnNext" accessKey="v" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/next_up.gif"
				CssClass="btnSpace" ToolTip="Next Delegate" CausesValidation="False" AlternateText="Next Delegate"></asp:imagebutton><asp:imagebutton id="btnLast" accessKey="b" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/end_up.gif"
				CssClass="btnSpace" ToolTip="Last Delegate" CausesValidation="False" AlternateText="Last Delegate"></asp:imagebutton><asp:imagebutton id="btnNew" accessKey="n" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/add_up.gif"
				CssClass="btnSpace" ToolTip="New Delegate" CausesValidation="False" AlternateText="New Delegate"></asp:imagebutton></td>
		<td style="WIDTH: 100%" align="right"><asp:button id="btnTop" runat="server" CssClass="small" BorderStyle="None" Font-Underline="True"
				ForeColor="Blue" Text="^ Top ^" BackColor="Transparent" ToolTip="Back to Top"></asp:button></td>
	</tr>
	<tr>
		<td><asp:image id="hlpUser" runat="server" ImageUrl="./images/help.png" ToolTip="Select an existing user from the pull-down to pre-populate this page with that individual's information."></asp:image></td>
		<td noWrap align="right"><asp:label id="lblUser" Runat="server" CssClass="boldContent">User:</asp:label></td>
		<td align="left"><asp:dropdownlist id="ddlUser" runat="server" CssClass="content" AutoPostBack="True"></asp:dropdownlist></td>
	</tr>
	<tr>
		<td vAlign="top"><asp:image id="hlpHIDPayrollName" runat="server" ImageUrl="./images/help.png" ToolTip="Enter an Hanford ID, Employee ID, or Last Name and click &quot;Find&quot; to pre-populate the page with that user's information.  If you chose to enter a last name you may be asked to select an individual from a list of names."></asp:image></td>
		<td vAlign="top" noWrap align="right"><asp:label id="lblHIDPayrollName" Runat="server" CssClass="boldContent">HID/Emplid/<br>Last Name:</asp:label></td>
		<td vAlign="middle" noWrap align="left"><asp:textbox id="txtHIDPayrollName" Runat="server" CssClass="content" MaxLength="100" Width="300px"></asp:textbox>&nbsp;<asp:button id="btnHIDPayrollNameFind" runat="server" CssClass="button" Text="Find"></asp:button></td>
	</tr>
	<TR>
		<TD vAlign="top"><asp:image id="hlpEmail" runat="server" ImageUrl="./images/help.png" ToolTip="The user's Email address."></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblEmail" Runat="server" CssClass="boldContent">E-Mail:</asp:label></TD>
		<TD align="left">
			<asp:Label id="txtEmail" runat="server" CssClass="content"></asp:Label></TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpFirstName" runat="server" ImageUrl="./images/help.png" ToolTip="The user's first name."></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblFirstName" Runat="server" CssClass="boldContent">First Name:</asp:label></TD>
		<TD align="left">
			<asp:Label id="txtFirstName" runat="server" CssClass="content"></asp:Label></TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpLastName" runat="server" ImageUrl="./images/help.png" ToolTip="The user's last name."></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblLastName" Runat="server" CssClass="boldContent">Last Name:</asp:label></TD>
		<TD align="left">
			<asp:Label id="txtLastName" runat="server" CssClass="content"></asp:Label></TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpMiddleName" runat="server" ImageUrl="./images/help.png" ToolTip="The user's middle name."></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblMiddleName" Runat="server" CssClass="boldContent">Middle Name:</asp:label></TD>
		<TD align="left">
			<asp:Label id="txtMiddleName" runat="server" CssClass="content"></asp:Label></TD>
	</TR>
	<TR>
		<TD vAlign="top" style="HEIGHT: 28px"><asp:image id="hlpSuffixName" runat="server" ImageUrl="./images/help.png" ToolTip="The user's suffix name (i.e., Sr., Jr., III, etc.)."></asp:image></TD>
		<TD vAlign="top" noWrap align="right" style="HEIGHT: 28px"><asp:label id="lblSuffixName" Runat="server" CssClass="boldContent">Suffix Name:</asp:label></TD>
		<TD align="left" style="HEIGHT: 28px">
			<asp:Label id="txtSuffixName" runat="server" CssClass="content"></asp:Label></TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpPhone" runat="server" ImageUrl="./images/help.png" ToolTip="The user's phone number."></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblPhoneNumber" Runat="server" CssClass="boldContent">Phone Number:</asp:label></TD>
		<TD align="left">
			<asp:Label id="txtPhoneNumber" runat="server" CssClass="content"></asp:Label></TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpCompany" runat="server" ImageUrl="./images/help.png" ToolTip="The name of the user's employer."></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblCompany" Runat="server" CssClass="boldContent">Company:</asp:label></TD>
		<TD align="left">
			<asp:Label id="txtCompany" runat="server" CssClass="content"></asp:Label></TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpPNNLTraining" runat="server" ImageUrl="./images/help.png" ToolTip="Will this user be delivering T&amp;Q or Training Related Surveys?"></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:checkbox id="chkTQSurvey" runat="server" CssClass="content"></asp:checkbox></TD>
		<TD align="left"><asp:label id="lblTraining" Runat="server" CssClass="boldContent">PNNL Training?</asp:label></TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpActive" runat="server" ImageUrl="./images/help.png" ToolTip="Is this user active?"></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:checkbox id="chkActive" runat="server" CssClass="content" Checked="True"></asp:checkbox></TD>
		<TD align="left"><asp:label id="lblActive" Runat="server" CssClass="boldContent">Active?</asp:label></TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpOwner" runat="server" ImageUrl="./images/help.png" ToolTip="Should this user be the new owner of this template?"></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:checkbox id="chkOwner" runat="server" CssClass="content"></asp:checkbox></TD>
		<TD align="left"><asp:label id="lb" Runat="server" CssClass="boldContent">Owner?</asp:label></TD>
	</TR>
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

