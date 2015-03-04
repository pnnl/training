<%@ Page Language="vb" AutoEventWireup="false" Codebehind="UserManage.aspx.vb" Inherits="SurveyAdmin.UserManage" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>UserManage</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<SCRIPT src="./pop_up.js" type="text/javascript"></SCRIPT>
	</HEAD>
	<body ms_positioning="GridLayout">
		<form id="Form1" method="post" runat="server">
			<table align="center" border="0">
				<tr align="center">
					<td><asp:image id="imgTitle" Runat="server" ImageUrl="./images/TemplateDelegatesTitle.png"></asp:image></td>
				</tr>
			</table>
			<p></p>
			<table border="0">
				<% if (Session("UserType") = 4) then %>
				<tr>
					<td><asp:image id="hlpServerSwitch" runat="server" ImageUrl="./images/help.png" ToolTip="Check this box to change the server this report will be polling its information from. Note: Changing the server here will not affect the overall site server location."></asp:image></td>
					<td colSpan="4"><asp:checkbox id="chkServerSwitch" runat="server" AutoPostBack="True" Font-Names="Verdana" Font-Size="Small"></asp:checkbox></td>
				</tr>
				<% else %>
				<tr>
					<td><asp:label id="lblPlaceHolder" Runat="server">&nbsp;<br /></asp:label></td>
				</tr>
				<% end if %>
			</table>
			<HR style="LEFT: 10px; BORDER-TOP-STYLE: ridge; TOP: 176px" width="100%" SIZE="3">
			<asp:label id="lblNavBar" style="FONT-SIZE: medium; Z-INDEX: 102; LEFT: 8px; POSITION: absolute; TOP: 128px"
				runat="server" Font-Names="Verdana" Font-Size="Medium" Width="184px" BackColor="Transparent"
				Height="20px"></asp:label><br>
			<asp:panel id="pnlNavBar" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 128px" runat="server"
				Width="100%" BackColor="#FFE0C0" Height="25px" HorizontalAlign="Right">
				<asp:button id="btnUp" runat="server" Font-Size="Small" Font-Names="Verdana" BackColor="Transparent"
					BorderStyle="None" Font-Underline="True" ForeColor="Blue" Text="^ Top ^"></asp:button>
			</asp:panel><br>
			<table id="table" style="LEFT: 10px; TOP: 208px" cellSpacing="5" cellPadding="5" border="0">
				<tr>
					<td><asp:image id="hlpUser" runat="server" ImageUrl="./images/help.png" ToolTip="Select an existing user from the pull-down to pre-populate this page with that individual's information."></asp:image></td>
					<td align="right"><asp:label id="lblUser" Runat="server" Font-Names="Verdana" Font-Size="Medium">User:</asp:label></td>
					<td align="left">
						<asp:dropdownlist id="ddlUser" runat="server" Font-Size="Small" Font-Names="Verdana" AutoPostBack="True"></asp:dropdownlist></td>
				</tr>
				<tr>
					<td vAlign="top"><asp:image id="hlpHIDPayrollName" runat="server" ImageUrl="./images/help.png" ToolTip="Enter an Hanford ID, Payroll Number, or Last Name and click &quot;Find&quot; to pre-populate the page with that user's infomration.  If you chose to enter a last name you may be asked to select an individual from a list of names."></asp:image></td>
					<td vAlign="top" align="right"><asp:label id="lblHIDPayrollName" Runat="server" Font-Names="Verdana" Font-Size="Medium">HID/Payroll/<br>Last Name:</asp:label></td>
					<td align="left">
						<asp:textbox id="txtHIDPayrollName" Runat="server" Font-Size="Small" Font-Names="Verdana" Width="300px"
							MaxLength="100"></asp:textbox></td>
				</tr>
				<TR>
					<TD vAlign="top">
						<asp:image id="hlpEmail" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the user's Email address."></asp:image></TD>
					<TD vAlign="top" align="right">
						<asp:label id="lblEmail" Runat="server" Font-Size="Medium" Font-Names="Verdana">E-Mail:</asp:label></TD>
					<TD align="left">
						<asp:textbox id="txtEmail" Runat="server" Font-Size="Small" Font-Names="Verdana" Width="300px"
							MaxLength="128"></asp:textbox></TD>
				</TR>
				<TR>
					<TD vAlign="top">
						<asp:image id="hlpFirstName" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the user's first name."></asp:image></TD>
					<TD vAlign="top" align="right">
						<asp:label id="lblFirstName" Runat="server" Font-Size="Medium" Font-Names="Verdana">First Name:</asp:label></TD>
					<TD align="left">
						<asp:textbox id="txtFirstName" Runat="server" Font-Size="Small" Font-Names="Verdana" Width="300px"
							MaxLength="50"></asp:textbox></TD>
				</TR>
				<TR>
					<TD vAlign="top">
						<asp:image id="hlpLastName" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the user's last name."></asp:image></TD>
					<TD vAlign="top" align="right">
						<asp:label id="lblLastName" Runat="server" Font-Size="Medium" Font-Names="Verdana">Last Name:</asp:label></TD>
					<TD align="left">
						<asp:textbox id="txtLastName" Runat="server" Font-Size="Small" Font-Names="Verdana" Width="300px"
							MaxLength="50"></asp:textbox></TD>
				</TR>
				<TR>
					<TD vAlign="top">
						<asp:image id="hlpMiddleName" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the user's middle name."></asp:image></TD>
					<TD vAlign="top" align="right">
						<asp:label id="lblMiddleName" Runat="server" Font-Size="Medium" Font-Names="Verdana">Middle Name:</asp:label></TD>
					<TD align="left">
						<asp:textbox id="txtMiddleName" Runat="server" Font-Size="Small" Font-Names="Verdana" Width="300px"
							MaxLength="50"></asp:textbox></TD>
				</TR>
				<TR>
					<TD vAlign="top">
						<asp:image id="hlpSuffixName" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the user's suffix name (i.e., Sr., Jr., III, etc.)."></asp:image></TD>
					<TD vAlign="top" align="right">
						<asp:label id="lblSuffixName" Runat="server" Font-Size="Medium" Font-Names="Verdana">Suffix Name:</asp:label></TD>
					<TD align="left">
						<asp:textbox id="txtSuffixName" Runat="server" Font-Size="Small" Font-Names="Verdana" Width="300px"
							MaxLength="5"></asp:textbox></TD>
				</TR>
				<TR>
					<TD vAlign="top">
						<asp:image id="hlpPhone" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the user's phone number."></asp:image></TD>
					<TD vAlign="top" align="right">
						<asp:label id="lblPhoneNumber" Runat="server" Font-Size="Medium" Font-Names="Verdana">Phone Number:</asp:label></TD>
					<TD align="left">
						<asp:textbox id="txtPhoneNumber" Runat="server" Font-Size="Small" Font-Names="Verdana" Width="300px"
							MaxLength="32"></asp:textbox></TD>
				</TR>
				<TR>
					<TD vAlign="top">
						<asp:image id="hlpCompany" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the name of the user's employer."></asp:image></TD>
					<TD vAlign="top" align="right">
						<asp:label id="lblCompany" Runat="server" Font-Size="Medium" Font-Names="Verdana">Company:</asp:label></TD>
					<TD align="left">
						<asp:textbox id="txtCompany" Runat="server" Font-Size="Small" Font-Names="Verdana" Width="300px"
							MaxLength="64"></asp:textbox></TD>
				</TR>
				<TR>
					<TD vAlign="top">
						<asp:image id="hlpUserLevel" runat="server" ImageUrl="./images/help.png" ToolTip="Select the user's level of access.  You may not grant access any level of access higher than your own."></asp:image></TD>
					<TD vAlign="top" align="right">
						<asp:label id="lblUserAccessLevel" Runat="server" Font-Size="Medium" Font-Names="Verdana">User Level:</asp:label></TD>
					<TD align="left">
						<asp:dropdownlist id="ddlUserLevel" runat="server" Font-Size="Small" Font-Names="Verdana" AutoPostBack="True"></asp:dropdownlist></TD>
				</TR>
				<TR>
					<TD vAlign="top">
						<asp:image id="hlpPNNLTraining" runat="server" ImageUrl="./images/help.png" ToolTip="Will this user be delivering T&amp;Q or Training Related Surveys?"></asp:image></TD>
					<TD vAlign="top" align="right">
						<asp:checkbox id="chkTQSurvey" runat="server" Font-Size="Small" Font-Names="Verdana"></asp:checkbox></TD>
					<TD align="left">
						<asp:label id="lblTraining" Runat="server" Font-Size="Medium" Font-Names="Verdana">PNNL Training?</asp:label></TD>
				</TR>
				<TR>
					<TD vAlign="top">
						<asp:image id="hlpActive" runat="server" ImageUrl="./images/help.png" ToolTip="Is this user active?"></asp:image></TD>
					<TD vAlign="top" align="right">
						<asp:checkbox id="chkActive" runat="server" Font-Size="Small" Font-Names="Verdana" Checked="True"></asp:checkbox></TD>
					<TD align="left">
						<asp:label id="lblActive" Runat="server" Font-Size="Medium" Font-Names="Verdana">Active?</asp:label></TD>
				</TR>
				<TR>
					<TD vAlign="top">
						<asp:image id="hlpOwner" runat="server" ImageUrl="./images/help.png" ToolTip="Should this user be the new owner of this template?"></asp:image></TD>
					<TD vAlign="top" align="right">
						<asp:checkbox id="chkOwner" runat="server" Font-Size="Small" Font-Names="Verdana"></asp:checkbox></TD>
					<TD align="left">
						<asp:label id="lb" Runat="server" Font-Size="Medium" Font-Names="Verdana">Owner?</asp:label></TD>
				</TR>
			</table>
			<asp:Button id="btnHIDPayrollNameFind" style="Z-INDEX: 105; LEFT: 528px; POSITION: absolute; TOP: 216px"
				runat="server" Font-Size="Medium" Font-Names="Verdana" Text="Find"></asp:Button>
			<asp:label id="lblError" style="Z-INDEX: 104; LEFT: 616px; POSITION: absolute; TOP: 240px"
				runat="server" Font-Names="Verdana" Font-Size="Medium" ForeColor="Red" Visible="False"></asp:label>
			<HR style="LEFT: 10px; BORDER-TOP-STYLE: ridge; TOP: 480px" width="100%" SIZE="3">
			<TABLE style="LEFT: 10px; TOP: 536px" width="100%" align="center" border="0">
				<tr>
					<td><asp:dropdownlist id="ddlPageOptions" runat="server" AutoPostBack="True" Font-Names="Verdana" Font-Size="Small"></asp:dropdownlist></td>
				</tr>
			</TABLE>
			<p></p>
			<br>
			<p></p>
			<TABLE style="LEFT: 10px; TOP: 600px" width="100%" align="center" border="0">
				<tr align="center">
					<td align="left"><asp:label id="lblSecPriv" Runat="server" Font-Names="Verdana" Font-Size="Small">
							<a href="http://wss.pnl.gov/security-privacy/sec-priv.html" target="_blank">Security 
								& Privacy</a></asp:label></td>
					<td align="right"><asp:label id="lblWebmaster" Runat="server" Font-Names="Verdana" Font-Size="Small">
							<a href="mailto:gene.gower@pnl.gov?subject=Survey Administration Question">Webmaster</a></asp:label></td>
				</tr>
			</TABLE>
		</form>
		</TR></TBODY></TABLE></FORM>
	</body>
</HTML>
