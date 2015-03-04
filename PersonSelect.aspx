<%@ Page trace="false" Language="vb" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" Codebehind="PersonSelect.aspx.vb" Inherits="SurveyAdmin.PersonSelect" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder" Runat="Server">
<%=Session("JSAction")%>
<form id="Form1" method="post" runat="server">
<h1>Person Select</h1>
<p></p>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table width="100%" class="borderless">
	<tr>
		<td><asp:label id="lblIntroduction" Runat="server" CssClass="boldContent">Your 
		search has returned multiple names.  Please select a person from the list below and click "Continue"
		to proceed.  You may also cancel out of this by clicking "Cancel".</asp:label></td>
	</tr>
</table>
<p></p>
<table cellSpacing="5" cellPadding="5" width="100%" align="center" class="borderless">
	<tr align="center">
		<td colSpan="2"><asp:listbox id="lbPersonList" runat="server" Rows="20" CssClass="content"></asp:listbox></td>
	</tr>
	<TR>
		<TD align="right" width="50%"><asp:button id="btnCancel" runat="server" Text="Cancel" CssClass="button" ToolTip="Cancel"></asp:button></TD>
		<td align="left" width="50%"><asp:button id="btnContinue" runat="server" Text="Continue" CssClass="button" ToolTip="Continue"></asp:button></td>
	</TR>
</table>
<p></p>
</form>
</asp:Content>
