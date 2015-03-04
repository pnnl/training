<%@ Register TagPrefix="igsch" Namespace="Infragistics.WebUI.WebSchedule" Assembly="Infragistics2.WebUI.WebDateChooser.v8.2, Version=8.2.20082.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" %>
<%@ Page trace="false" validateRequest="false" Language="vb" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" Codebehind="List.aspx.vb" Inherits="SurveyAdmin.List" %>
<%@ Register TagPrefix="igtxt" Namespace="Infragistics.WebUI.WebDataInput" Assembly="Infragistics2.WebUI.WebDataInput.v8.2, Version=8.2.20082.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder" Runat="Server">
<%=Session("JSAction")%>
<script type="text/javascript">
	var msg = null;
	var tipTimer;
	
	function clean_up() {
		if (msg != null) {
			if (msg.open) {
			msg.close();
			msg = null;
			}
		}
		window.clearTimeout(tipTimer);
	}

	function WinPopNew(strURL, intX, intY)
	{
		//used to destroy a popup so that the new popup will be the right size.
		//also checks to see if user is on a mac, to prevent clean_up which will break this. (JR_Sharp)
		var bAintMac = <%=lcase(cStr((InStr(lcase(Request.ServerVariables("ALL_HTTP")), "win") <> 0)))%>;
		var sServerName = "<%=lcase(Request.ServerVariables("SERVER_NAME"))%>";
		var bPopUp = true;
		var sTmpURL = strURL;
		if (bAintMac)
		{
			clean_up(); 
		}
		
		strOptions = 'menubar=0,directories=0,scrollbars=0,toolbar=0,status=0,resizable=1,';
		strOptions = strOptions + 'width=' + intX + ',height=' + intY + ',ontop=1';

		if (msg = open(sTmpURL, "frameWindow", strOptions))
		{
			msg.focus();
		}
		else
		{
			clean_up();
		}
		tipTimer=window.setTimeout("clean_up()", 10000);
	}
</script>
<form id="Form1" method="post" runat="server">
<h1>Template Question List Items</h1>
<p></p>
<table class="borderless">
	<tr>
		<td align="right"><asp:label id="lblTemplate" Runat="server" CssClass="bold_medium">Template:</asp:label></td>
		<td align="left" class='medium'><%=Session("TemplateName")%></td>
		<td align="center" width="50"></td>
		<td align="right"><asp:image id="hlpTemplatePreview" runat="server" ImageUrl="./images/help.png" ToolTip="Click this button to preview the template."></asp:image></td>
		<td align="left"><asp:button id="btnTemplatePreview" runat="server"  CssClass="button" Text="Preview"></asp:button></td>
	</tr>
	<tr>
		<td align="right"><asp:label id="lblQuestion" Runat="server" CssClass="bold_medium">Question:</asp:label></td>
		<td align="left" class='medium'><%=Session("QuestionName")%></td>
		<td align="center" width="50"></td>
		<td align="right"></td>
		<td align="left"></td>
	</tr>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="0" cellPadding="2" class="borderless">
	<tr class="navPanel">
		<td></td>
		<td noWrap align="right"><asp:label id="lblNavBarQuestion" runat="server" CssClass="lblNavPanel"></asp:label></td>
		<td style="WIDTH: 305px"><asp:imagebutton id="btnStart" accessKey="z" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/start_up.gif"
				ToolTip="First Question" CssClass="btnSpace" AlternateText="First Question" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnPrev" accessKey="x" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/prev_up.gif"
				ToolTip="Previous Question" CssClass="btnSpace" AlternateText="Previous Question" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnReload" accessKey="c" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/ig_menuCurrent_up.png"
				ToolTip="Reload Question" CssClass="btnSpace" AlternateText="Reload Question" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnNext" accessKey="v" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/next_up.gif"
				ToolTip="Next Question" CssClass="btnSpace" AlternateText="Next Question" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnLast" accessKey="b" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/end_up.gif"
				ToolTip="Last Question" CssClass="btnSpace" AlternateText="Last Question" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnNew" accessKey="n" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/add_up.gif"
				ToolTip="New Question" CssClass="btnSpace" AlternateText="New Question" CausesValidation="False"></asp:imagebutton></td>
		<td style="WIDTH: 100%" align="right"><asp:button id="btnTop" runat="server" CssClass="small" Text="^ Top ^" ForeColor="Blue" Font-Underline="True"
				BorderStyle="None" BackColor="Transparent"></asp:button></td>
	</tr>
	<tr class="navPanel2">
		<td></td>
		<td noWrap align="right"><asp:label id="lblNavBarList" runat="server" CssClass="lblNavPanel"></asp:label></td>
		<td style="WIDTH: 305px"><asp:imagebutton id="btnStart2" accessKey="z" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/start_up.gif"
				ToolTip="First List Item" CssClass="btnSpace" AlternateText="First List Item" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnPrev2" accessKey="x" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/prev_up.gif"
				ToolTip="Previous List Item" CssClass="btnSpace" AlternateText="Previous List Item" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnReload2" accessKey="c" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/ig_menuCurrent_up.png"
				ToolTip="Reload List Item" CssClass="btnSpace" AlternateText="Reload List Item" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnNext2" accessKey="v" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/next_up.gif"
				ToolTip="Next List Item" CssClass="btnSpace" AlternateText="Next List Item" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnLast2" accessKey="b" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/end_up.gif"
				ToolTip="Last List Item" CssClass="btnSpace" AlternateText="Last List Item" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnNew2" accessKey="n" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/add_up.gif"
				ToolTip="New List Item" CssClass="btnSpace" AlternateText="New List Item" CausesValidation="False"></asp:imagebutton></td>
		<td style="WIDTH: 100%" align="right"><asp:button id="btnUp" runat="server" CssClass="small" Text="^ Up ^" ForeColor="Blue" Font-Underline="True"
				BorderStyle="None" BackColor="Transparent"></asp:button></td>
	</tr>
	<tr>
		<td vAlign="top"><asp:image id="hlpItemLabel" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the text of the option the user will be selecting."></asp:image></td>
		<td vAlign="top" noWrap align="right"><asp:label id="lblItemLabel" Runat="server" CssClass="bold_medium">Item Label:</asp:label></td>
		<td style="WIDTH: 305px" align="left"><asp:textbox id="txtItemLabel" runat="server" CssClass="small" MaxLength="64" Width="300px"></asp:textbox></td>
	</tr>
	<% if (Session("selectedQuestionType") = "P") then %>
	<TR>
		<TD vAlign="top"><asp:image id="hlpUpload" runat="server" ImageUrl="./images/help.png" ToolTip="Please select an image file for upload by selecting the browse button and selecting an image file."></asp:image></TD>
		<TD vAlign="top" noWrap align="right"><asp:label id="lblUpload" Runat="server" CssClass="bold_medium">Upload Image:</asp:label></TD>
		<TD style="WIDTH: 305px" align="left"><INPUT id="fuImageLoader" style="FONT-SIZE: small; FILTER: image" type="file" name="fuImageLoader"
				runat="server"></TD>
		<td>
			<asp:Table id="ListTable" runat="server"></asp:Table></td>
	</TR>
	<% end if %>
	<tr>
		<td vAlign="top"><asp:image id="hlpBranch" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the question number the survey should branch to when the user selects this list item.  If you do not specify a branch then a default value of 0 will be entered and will be ignored by the survey page."></asp:image></td>
		<td vAlign="top" noWrap align="right"><asp:label id="lblBranch" Runat="server" CssClass="bold_medium">Branch to:</asp:label></td>
		<td style="WIDTH: 305px" align="left"><igtxt:webnumericedit id="wneBranch" runat="server" ToolTip="Enter the question number to branch to."
				Width="300px" MaxValue="1000" Section508Compliant="True" MinValue="0" DataMode="Int" CssClass="small"></igtxt:webnumericedit></td>
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
