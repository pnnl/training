<%@ Page trace="false"  Language="vb" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" Codebehind="PubRequest.aspx.vb" Inherits="SurveyAdmin.PubRequest" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder" Runat="Server">
<%=Session("JSAction")%>
<form id="Form1" method="post" runat="server">
<h1>Template Publication Request</h1>
<p></p>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table width="100%" class="borderless">
	<tr>
		<td><asp:label id="lblIntroduction" Runat="server" CssClass="content"><font style="COLOR: red">Note:</font>
		You must find the appropriate ADC (Authorized Derivative Classifier) representative for your department 
		from the following listing file in one of two formats   
		(<a title="ADC Document in Word" href="http://sass.pnl.gov/documents/adclisting.doc"
					target="_blank">Word</a> or 
		<a title="ADC Document in PDF" href="http://sass.pnl.gov/documents/adclisting.pdf"
					target="_blank">PDF</a>).
		</asp:label>
			<p><asp:label id="lblIntroduction2" Runat="server" CssClass="content">
		You must then enter their last name, payroll, or HID into the text box below and click the <i>Find</i> 
		button to locate their information.  Verify that the information returned is correct and then press the <i>Submit</i> button.
		</asp:label>
			<p><asp:label id="Label1" Runat="server" CssClass="content">
		Your request will then be sent to ADC and Human Factors for review.  You will be notified if it is rejected or 
		accepted.  Once the survey template has been completely approved by both individuals your survey template
		will then be published by the survey administrator to production where it can then be used in making and
		sending out surveys.
		</asp:label></p>
		</td>
	</tr>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="0" cellPadding="2" class="borderless">
	<tr>
		<td vAlign="top"><asp:image id="hlpHIDPayrollName" runat="server" ImageUrl="./images/help.png" ToolTip="Enter an Hanford ID, Employee ID, or Last Name and click &quot;Find&quot; to pre-populate the page with that user's information.  If you chose to enter a last name you may be asked to select an individual from a list of names."></asp:image></td>
		<td vAlign="top" noWrap align="right"><asp:label id="lblHIDPayrollName" Runat="server" CssClass="boldContent">HID/Emplid/<br>Last Name:</asp:label></td>
		<td vAlign="middle" noWrap align="left"><asp:textbox id="txtHIDPayrollName" Runat="server" CssClass="content" MaxLength="100" Width="300px"></asp:textbox>&nbsp;<asp:button id="btnHIDPayrollNameFind" runat="server" CssClass="button" Text="Find" ToolTip="Find"></asp:button></td>
	</tr>
	<TR>
		<TD vAlign="top"></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblEmail" Runat="server" CssClass="boldContent">E-Mail:</asp:label></TD>
		<TD align="left"><asp:label id="lblEmailDisplay" runat="server" CssClass="content"></asp:label></TD>
	</TR>
	<TR>
		<TD vAlign="top"></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblFirstName" Runat="server" CssClass="boldContent">First Name:</asp:label></TD>
		<TD align="left"><asp:label id="lblFirstNameDisplay" runat="server" CssClass="content"></asp:label></TD>
	</TR>
	<TR>
		<TD vAlign="top"></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblLastName" Runat="server" CssClass="boldContent">Last Name:</asp:label></TD>
		<TD align="left"><asp:label id="lblLastNameDisplay" runat="server" CssClass="content"></asp:label></TD>
	</TR>
	<TR>
		<TD style="HEIGHT: 22px" vAlign="top"></TD>
		<TD style="HEIGHT: 22px" vAlign="top" noWrap align="right"><asp:label id="lblMiddleName" Runat="server" CssClass="boldContent">Middle Name:</asp:label></TD>
		<TD style="HEIGHT: 22px" align="left"><asp:label id="lblMiddleNameDisplay" runat="server" CssClass="content"></asp:label></TD>
	</TR>
	<TR>
		<TD vAlign="top"></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblHID" Runat="server" CssClass="boldContent">Hanford ID:</asp:label></TD>
		<TD align="left"><asp:label id="lblHIDDisplay" runat="server" CssClass="content"></asp:label></TD>
	</TR>
	<TR>
		<TD vAlign="top"></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblDepartment" Runat="server" CssClass="boldContent"> Department:</asp:label></TD>
		<TD align="left"><asp:label id="lblDepartmentDisplay" runat="server" CssClass="content"></asp:label></TD>
	</TR>
	<TR>
		<TD vAlign="top"></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblPhoneNumber" Runat="server" CssClass="boldContent">Phone Number:</asp:label></TD>
		<TD align="left"><asp:label id="lblPhoneNumberDisplay" runat="server" CssClass="content"></asp:label></TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpAddtlInfo" runat="server" ImageUrl="./images/help.png" ToolTip="Enter additional information and links that you want to appear in the outgoing e-mail."></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblAddtlInfo" Runat="server" CssClass="boldContent">Additional Information:</asp:label></TD>
		<TD align="left"><asp:textbox id="txtAddtlInfo" runat="server" CssClass="content" MaxLength="1800" Width="400px"
				Columns="50" Rows="15" TextMode="MultiLine" Height="200px"></asp:textbox></TD>
	</TR>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<TABLE width="100%" align="center" class="borderless">
	<tr>
		<td align="center"><asp:button id="btnSubmit" runat="server" CssClass="button" Text="Submit" ToolTip="Submit"></asp:button><asp:button id="btnCancel" runat="server" CssClass="button" Text="Cancel" ToolTip="Cancel"></asp:button></td>
	</tr>
</TABLE>
</form>
</asp:Content>
