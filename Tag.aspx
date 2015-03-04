<%@ Page trace="false" validateRequest="false" Language="vb" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" Codebehind="Tag.aspx.vb" Inherits="SurveyAdmin.Tag"  %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder" Runat="Server">
<%=Session("JSAction")%>
<form id="Form1" method="post" runat="server">
<h1>Template Tag Items</h1>
<p></p>
<table class="borderless">
	<tr>
		<td align="left"><asp:label id="lblTemplate" Runat="server" CssClass="content">Template:</asp:label></td>
		<td align="left" class='small'><%=Session("TemplateName")%></td>
		<td align="center" width="50"></td>
		<td align="right"><asp:image id="hlpTemplatePreview" runat="server" ImageUrl="./images/help.png" ToolTip="Click this button to preview the template."></asp:image></td>
		<td align="left"><asp:button id="btnTemplatePreview" runat="server" Text="Preview" CssClass="button" ToolTip="Preview"></asp:button></td>
	</tr>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table id="table" cellSpacing="0" cellPadding="2" class="borderless">
	<tr class="navPanel">
		<td></td>
		<td align="right" nowrap><asp:label id="lblNavBar" runat="server" CssClass="lblNavPanel"></asp:label></td>
		<td><asp:imagebutton id="btnStart" accessKey="z" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/start_up.gif"
				ToolTip="First Tag Item" CssClass="btnSpace" CausesValidation="False" AlternateText="First Tag Item"></asp:imagebutton><asp:imagebutton id="btnPrev" accessKey="x" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/prev_up.gif"
				ToolTip="Previous Tag Item" CssClass="btnSpace" CausesValidation="False" AlternateText="Previous Tag Item"></asp:imagebutton><asp:imagebutton id="btnReload" accessKey="c" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/ig_menuCurrent_up.png"
				ToolTip="Reload Tag Item" CssClass="btnSpace" CausesValidation="False" AlternateText="Reload Tag Item"></asp:imagebutton><asp:imagebutton id="btnNext" accessKey="v" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/next_up.gif"
				ToolTip="Next Tag Item" CssClass="btnSpace" CausesValidation="False" AlternateText="Next Tag Item"></asp:imagebutton><asp:imagebutton id="btnLast" accessKey="b" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/end_up.gif"
				ToolTip="Last Tag Item" CssClass="btnSpace" CausesValidation="False" AlternateText="Last Tag Item"></asp:imagebutton><asp:imagebutton id="btnNew" accessKey="n" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/add_up.gif"
				ToolTip="New Tag Item" CssClass="btnSpace" CausesValidation="False" AlternateText="New Tag Item"></asp:imagebutton></td>
		<td style="WIDTH: 100%" align="right"><asp:button id="btnTop" runat="server" Text="^ Top ^" BackColor="Transparent" BorderStyle="None"
				Font-Underline="True" ForeColor="Blue" CssClass="small" ToolTip="Back to Top"></asp:button></td>
	</tr>
	<tr>
		<td><asp:image id="hlpTagName" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the name for your tag item."></asp:image></td>
		<td align="right" noWrap><asp:label id="lblTagName" Runat="server" CssClass="boldContent">Tag Name:</asp:label></td>
		<td align="left"><asp:textbox id="txtTagName" runat="server" Width="300px" MaxLength="50" CssClass="content"></asp:textbox></td>
	</tr>
	<tr>
		<td vAlign="top"><asp:image id="hlpTagQuestion" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the question to be displayed to a user making a survey from this template."></asp:image></td>
		<td vAlign="top" align="right" noWrap><asp:label id="lblTagQuestion" Runat="server" CssClass="boldContent">Tag Question:</asp:label></td>
		<td align="left"><asp:textbox id="txtTagQuestion" runat="server" Width="400px" Height="200px" MaxLength="1800"
				TextMode="MultiLine" CssClass="content" Rows="15" Columns="50"></asp:textbox></td>
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