<%@ Register TagPrefix="igtxt" Namespace="Infragistics.WebUI.WebDataInput" Assembly="Infragistics2.WebUI.WebDataInput.v8.2, Version=8.2.20082.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" %>
<%@ Page trace="false" Language="vb" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" Codebehind="QuestionGroups.aspx.vb" Inherits="SurveyAdmin.QuestionGroup"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder" Runat="Server">
<%=Session("JSAction")%>
<form id="Form1" method="post" runat="server">
<h1>Template Question Groups</h1>
<p></p>
<table class="borderless">
	<tr>
		<td align="left"><asp:label id="lblTemplate" Runat="server" CssClass="boldContent">Template:</asp:label></td>
		<td class="content" align="left"><%=Session("TemplateName")%></td>
		<td align="center" width="50"></td>
		<td align="right"><asp:image id="hlpTemplatePreview" runat="server" ImageUrl="./images/help.png" ToolTip="Click this button to preview the template."></asp:image></td>
		<td align="left"><asp:button id="btnTemplatePreview" runat="server" CssClass="button" Text="Preview" ToolTip="Preview"></asp:button></td>
	</tr>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="0" cellPadding="2" class="borderless">
	<tr class="navPanel">
		<td></td>
		<td noWrap align="right"><asp:label id="lblNavBar" runat="server" CssClass="lblNavPanel"></asp:label></td>
		<td noWrap><asp:imagebutton id="btnStart" accessKey="z" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/start_up.gif"
				ToolTip="First Question" CssClass="btnSpace" AlternateText="First Question" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnPrev" accessKey="x" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/prev_up.gif"
				ToolTip="Previous Question" CssClass="btnSpace" AlternateText="Previous Question" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnReload" accessKey="c" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/ig_menuCurrent_up.png"
				ToolTip="Reload Question" CssClass="btnSpace" AlternateText="Reload Question" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnNext" accessKey="v" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/next_up.gif"
				ToolTip="Next Question" CssClass="btnSpace" AlternateText="Next Question" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnLast" accessKey="b" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/end_up.gif"
				ToolTip="Last Question" CssClass="btnSpace" AlternateText="Last Question" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnNew" accessKey="n" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/add_up.gif"
				ToolTip="New Question" CssClass="btnSpace" AlternateText="New Question" CausesValidation="False"></asp:imagebutton></td>
		<td style="WIDTH: 100%" align="right"><asp:button id="btnTop" runat="server" CssClass="small" Text="^ Top ^" ForeColor="Blue" Font-Underline="True"
				BorderStyle="None" BackColor="Transparent" ToolTip="Back to Top"></asp:button></td>
	</tr>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="0" cellPadding="2" class="borderless">
	<tr>
		<td vAlign="top"><asp:image id="hlpGroupName" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the name of the group."></asp:image></td>
		<td vAlign="top" align="right"><asp:label id="lblGroupName" Runat="server" CssClass="boldContent">Group Name:</asp:label></td>
		<td align="left"><asp:textbox id="txtGroupName" runat="server" CssClass="content" Columns="50" MaxLength="200"
				Width="400px"></asp:textbox></td>
	</tr>
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
