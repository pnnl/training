<%@ Page trace="false" Language="vb" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" Codebehind="ManageUsers.aspx.vb" Inherits="SurveyAdmin.ManageUsers"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder" Runat="Server">
<%=Session("JSAction")%>
<form id="Form1" method="post" runat="server">
<h1>User Management</h1>
<p></p>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table class="borderless">
	<tr>
		<td><asp:image id="hlpImportList" runat="server" ImageUrl="./images/help.png" ToolTip="Check this box to allow for a list of users to be input."></asp:image></td>
		<td colSpan="4"><asp:checkbox id="chkImportList" runat="server" CssClass="content" AutoPostBack="True" Text="Import a list of users?"></asp:checkbox></td>
	</tr>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="0" cellPadding="2" class="borderless">
	<% if Not Session("ImportList") %>
	<tr>
		<td><asp:image id="hlpUser" runat="server" ImageUrl="./images/help.png" ToolTip="Select an existing user or start creating a new user.  Note: Users who already exist in the system will not be inserted but their information will be retrieved from the system for you."></asp:image></td>
		<td noWrap align="right"><asp:label id="lblUser" Runat="server" CssClass="boldContent">User:</asp:label></td>
		<td align="left" colSpan="2"><asp:dropdownlist id="ddlUser" runat="server" CssClass="content" AutoPostBack="True"></asp:dropdownlist></td>
		<TD align="left"></TD>
	</tr>
	<tr>
		<td vAlign="top"><asp:image id="hlpHIDPayrollName" runat="server" ImageUrl="./images/help.png" ToolTip="Enter an Hanford ID, Employee ID, or Last Name and click &quot;Find&quot; to pre-populate the page with that user's information.  If you chose to enter a last name you may be asked to select an individual from a list of names."></asp:image></td>
		<td vAlign="top" noWrap align="right"><asp:label id="lblHIDPayrollName" Runat="server" CssClass="boldContent">HID/Emplid/<br>Last Name:</asp:label></td>
		<td vAlign="middle" noWrap align="left"><asp:textbox id="txtHIDPayrollName" Runat="server" CssClass="content" Width="300px" MaxLength="100"></asp:textbox>&nbsp;
			<asp:button id="btnHIDPayrollNameFind" runat="server" ToolTip="Find" CssClass="button" Text="Find"></asp:button></td>
		<TD vAlign="middle" noWrap align="left"><asp:checkbox id="chkFindCurrentUsers" runat="server" ToolTip="Check this box to find the users currently in the Survey Tool's list of users.  Leave this unchecked to search for employees and Hanford workers."
				CssClass="content" Text="Find in current users?"></asp:checkbox></TD>
	</tr>
	<TR>
		<TD vAlign="top"><asp:image id="hlpEmail" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the user's Email address."></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblEmail" Runat="server" CssClass="boldContent">E-Mail:</asp:label></TD>
		<TD align="left"><asp:textbox id="txtEmail" Runat="server" CssClass="content" Width="300px" MaxLength="128"></asp:textbox></TD>
		<TD align="left"></TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpFirstName" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the user's first name."></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblFirstName" Runat="server" CssClass="boldContent">First Name:</asp:label></TD>
		<TD align="left"><asp:textbox id="txtFirstName" Runat="server" CssClass="content" Width="300px" MaxLength="50"></asp:textbox></TD>
		<TD align="left"></TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpLastName" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the user's last name."></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblLastName" Runat="server" CssClass="boldContent">Last Name:</asp:label></TD>
		<TD align="left"><asp:textbox id="txtLastName" Runat="server" CssClass="content" Width="300px" MaxLength="50"></asp:textbox></TD>
		<TD align="left"></TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpMiddleName" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the user's middle name."></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblMiddleName" Runat="server" CssClass="boldContent">Middle Name:</asp:label></TD>
		<TD align="left"><asp:textbox id="txtMiddleName" Runat="server" CssClass="content" Width="300px" MaxLength="50"></asp:textbox></TD>
		<TD align="left"></TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpSuffixName" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the user's suffix name (i.e., Sr., Jr., III, etc.)."></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblSuffixName" Runat="server" CssClass="boldContent">Suffix Name:</asp:label></TD>
		<TD align="left"><asp:textbox id="txtSuffixName" Runat="server" CssClass="content" Width="300px" MaxLength="5"></asp:textbox></TD>
		<TD align="left"></TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpPhone" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the user's phone number."></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblPhoneNumber" Runat="server" CssClass="boldContent">Phone Number:</asp:label></TD>
		<TD align="left"><asp:textbox id="txtPhoneNumber" Runat="server" CssClass="content" Width="300px" MaxLength="32"></asp:textbox></TD>
		<TD align="left"></TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpCompany" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the name of the user's employer."></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblCompany" Runat="server" CssClass="boldContent">Company:</asp:label></TD>
		<TD align="left"><asp:textbox id="txtCompany" Runat="server" CssClass="content" Width="300px" MaxLength="64"></asp:textbox></TD>
		<TD align="left"></TD>
	</TR>
	<% end if %>
	<TR>
		<TD vAlign="top"><asp:image id="hlpUserType" runat="server" ImageUrl="./images/help.png" ToolTip="Select the user's type.  You may not make a user any type that has greater ability than your own."></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblUserType" Runat="server" CssClass="boldContent">User Type:</asp:label></TD>
		<TD align="left"><asp:dropdownlist id="ddlUserType" runat="server" CssClass="content"></asp:dropdownlist></TD>
		<TD align="left"></TD>
	</TR>
	<% if Not Session("ImportList") %>
	<TR>
		<TD vAlign="top"><asp:image id="hlpPNNLTraining" runat="server" ImageUrl="./images/help.png" ToolTip="Will this user be delivering T&amp;Q or Training Related Surveys?"></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:checkbox id="chkTQSurvey" runat="server" CssClass="content"></asp:checkbox></TD>
		<TD align="left"><asp:label id="lblTraining" Runat="server" CssClass="boldContent">PNNL Training?</asp:label></TD>
		<TD align="left"></TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpActive" runat="server" ImageUrl="./images/help.png" ToolTip="Is this user active?"></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:checkbox id="chkActive" runat="server" CssClass="content" Checked="True"></asp:checkbox></TD>
		<TD align="left"><asp:label id="lblActive" Runat="server" CssClass="boldContent">Active?</asp:label></TD>
		<TD align="left"></TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpPasswordReset" runat="server" ImageUrl="./images/help.png" ToolTip="Should this user's password be reset?"></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:checkbox id="chkPasswordReset" runat="server" CssClass="content"></asp:checkbox></TD>
		<TD align="left"><asp:label id="lblPasswordReset" Runat="server" CssClass="boldContent">Reset Password?</asp:label></TD>
		<TD align="left"></TD>
	</TR>
	<% else %>
	<TR>
		<TD vAlign="top"><asp:image id="hlpUserList" runat="server" ImageUrl="./images/help.png" ToolTip="Enter a series of comma and/or line separated items.  These items may include E-mail addresses (but will not recognize group e-mail addresses), Hanford ID's, and Employee IDs."></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblUserList" Runat="server" CssClass="boldContent">E-Mail List/<br>HID/Emplid:<br>(Separate your<br />items with commas<br /> and/or rows)</asp:label></TD>
		<TD align="left"><asp:textbox id="txtUserList" runat="server" CssClass="content" Width="400px" MaxLength="8000"
				Height="200px" TextMode="MultiLine" Rows="15" Columns="50"></asp:textbox></TD>
		<TD align="left"></TD>
	</TR>
	<% end if %>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table width="100%" align="center" class="borderless">
	<tr>
		<td>
		    <% if Session("PageOptions") = 2 then %>
		        <asp:button id="btnInsert" Runat="server" CssClass="button" ToolTip="Insert" Text="Insert"></asp:button>
		    <% end if %>
		    <% If Session("PageOptions") <> 2 Then%>
		        <asp:button id="btnUpdate" Runat="server" CssClass="button" ToolTip="Update" Text="Update"></asp:button>
		        <asp:button id="btnDelete" Runat="server" CssClass="button" ToolTip="Delete" Text="Delete"></asp:button>
		    <% end if %>
		    <asp:button id="btnReset" Runat="server" CssClass="button" ToolTip="Reset" Text="Reset"></asp:button>
		    <% If Session("PageOptions") <> 2 Then%>
		        <asp:button id="btnClear" Runat="server" ToolTip="New User Input" CssClass="button" Text="Clear"></asp:button>
		    <% end if %>
		</td>
	</tr>
</table>
</form>
</asp:Content>