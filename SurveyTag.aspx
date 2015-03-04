<%@ Page trace="false" validateRequest="false" MasterPageFile="~/MasterPage.master" Language="vb" AutoEventWireup="false" Codebehind="SurveyTag.aspx.vb" Inherits="SurveyAdmin.SurveyTag" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder" Runat="Server">
<%=Session("JSAction")%>
<form id="Form1" method="post" runat="server">
<h1>Template Survey Tag Items</h1>
<p></p>
<table class="borderless">
	<tr>
		<td align="left"><asp:label id="lblTemplate" Runat="server" CssClass="boldContent">Template:</asp:label></td>
		<td class="content" align="left"><%=Session("TemplateName")%></td>
		<td align="center" width="50"></td>
		<td align="right"></td>
		<td align="left"></td>
	</tr>
	<tr>
		<td align="right"><asp:label id="lblQuestion" Runat="server" CssClass="boldContent">Survey:</asp:label></td>
		<td class="content" align="left"><%=Session("SurveyName")%></td>
		<td align="center" width="50"></td>
		<td align="right"><asp:image id="hlpTemplatePreview" runat="server" ImageUrl="./images/help.png" ToolTip="Click this button to preview the template."></asp:image></td>
		<td align="left"><asp:button id="btnTemplatePreview" runat="server" CssClass="button" Text="Preview" ToolTip="Preview"></asp:button></td>
	</tr>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="0" cellPadding="2" class="borderless">
	<tr class="navPanel">
		<td></td>
		<td noWrap align="right"><asp:label id="lblNavBarSurvey" runat="server" CssClass="lblNavPanel"></asp:label></td>
		<td style="WIDTH: 305px"><asp:imagebutton id="btnStart" accessKey="z" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/start_up.gif"
				ToolTip="First Survey" CssClass="btnSpace" AlternateText="First Question" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnPrev" accessKey="x" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/prev_up.gif"
				ToolTip="Previous Survey" CssClass="btnSpace" AlternateText="Previous Question" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnReload" accessKey="c" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/ig_menuCurrent_up.png"
				ToolTip="Reload Survey" CssClass="btnSpace" AlternateText="Reload Question" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnNext" accessKey="v" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/next_up.gif"
				ToolTip="Next Survey" CssClass="btnSpace" AlternateText="Next Question" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnLast" accessKey="b" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/end_up.gif"
				ToolTip="Last Survey" CssClass="btnSpace" AlternateText="Last Question" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnNew" accessKey="n" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/add_up.gif"
				ToolTip="New Survey" CssClass="btnSpace" AlternateText="New Question" CausesValidation="False"></asp:imagebutton></td>
		<td style="WIDTH: 100%" align="right"><asp:button id="btnTop" runat="server" CssClass="small" Text="^ Top ^" ForeColor="Blue" Font-Underline="True"
				BorderStyle="None" BackColor="Transparent" ToolTip="Back to Top"></asp:button></td>
	</tr>
	<tr class="navPanel2">
		<td></td>
		<td noWrap align="right"><asp:label id="lblNavBarTagItem" runat="server" CssClass="lblNavPanel"></asp:label></td>
		<td style="WIDTH: 305px"><asp:imagebutton id="btnStart2" accessKey="z" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/start_up.gif"
				ToolTip="First Tag Item" CssClass="btnSpace" AlternateText="First List Item" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnPrev2" accessKey="x" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/prev_up.gif"
				ToolTip="Previous Tag Item" CssClass="btnSpace" AlternateText="Previous List Item" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnReload2" accessKey="c" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/ig_menuCurrent_up.png"
				ToolTip="Reload Tag Item" CssClass="btnSpace" AlternateText="Reload List Item" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnNext2" accessKey="v" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/next_up.gif"
				ToolTip="Next Tag Item" CssClass="btnSpace" AlternateText="Next List Item" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnLast2" accessKey="b" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/end_up.gif"
				ToolTip="Last Tag Item" CssClass="btnSpace" AlternateText="Last List Item" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnNew2" accessKey="n" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/add_up.gif"
				ToolTip="New Tag Item" CssClass="btnSpace" AlternateText="New List Item" CausesValidation="False"></asp:imagebutton></td>
		<td style="WIDTH: 100%" align="right"><asp:button id="btnUp" runat="server" CssClass="small" Text="^ Up ^" ForeColor="Blue" Font-Underline="True"
				BorderStyle="None" BackColor="Transparent" ToolTip="Back Up One Level"></asp:button></td>
	</tr>
	<tr>
		<td><asp:image id="hlpTagQuestion" runat="server" ImageUrl="./images/help.png" ToolTip="Contains a question created by the template creator that you must answer before sending out this survey."></asp:image></td>
		<td noWrap align="right"><asp:label id="lblTagQuestionName" Runat="server" CssClass="boldContent">Tag Question:</asp:label></td>
		<td align="left"><asp:label id="lblTagQuestion" Runat="server" CssClass="content"></asp:label></td>
	</tr>
	<tr>
		<td vAlign="top"><asp:image id="hlpTagAnswer" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the answer to the question above in the text area."></asp:image></td>
		<td vAlign="top" noWrap align="right"><asp:label id="lblTagAnswer" Runat="server" CssClass="boldContent">Tag Answer:</asp:label></td>
		<td align="left"><asp:textbox id="txtTagAnswer" runat="server" CssClass="content" Columns="50" Rows="15" TextMode="MultiLine"
				MaxLength="1800" Height="200px" Width="400px"></asp:textbox></td>
	</tr>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table width="100%" align="center" class="borderless">
	<tr>
		<td>
		    <asp:button id="btnUpdate" Runat="server" CssClass="button" ToolTip="Update" Text="Update"></asp:button>
		    <asp:button id="btnReset" Runat="server" CssClass="button" ToolTip="Reset" Text="Reset"></asp:button>
		</td>
	</tr>
</table>
</form>
</asp:Content>

