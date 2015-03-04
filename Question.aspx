<%@ Page trace="false" validateRequest="false" MasterPageFile="~/MasterPage.master" Language="vb" AutoEventWireup="false" Codebehind="Question.aspx.vb" Inherits="SurveyAdmin.Question"  %>
<%@ Register TagPrefix="igtxt" Namespace="Infragistics.WebUI.WebDataInput" Assembly="Infragistics2.WebUI.WebDataInput.v8.2, Version=8.2.20082.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder" Runat="Server">
<%=Session("JSAction")%>
<form id="Form1" method="post" runat="server">
<h1>Template Questions</h1>
<p></p>
<table class="borderless">
	<tr>
		<td align="left"><asp:label id="lblTemplate" Runat="server" CssClass="boldContent">Template:</asp:label></td>
		<td class="content" align="left"><%=Session("TemplateName")%></td>
		<td align="center" width="50"></td>
		<td align="right"><asp:image id="hlpTemplatePreview" runat="server" ImageUrl="./images/help.png" ToolTip="Click this button to preview the template."></asp:image></td>
		<td align="left"><asp:button id="btnTemplatePreview" runat="server" CssClass="button" Text="Preview"></asp:button></td>
	</tr>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="0" cellPadding="2" class="borderless">
	<tr class="navPanel">
		<td></td>
		<td noWrap align="right"><asp:label id="lblNavBar" runat="server" CssClass="lblNavPanel"></asp:label></td>
		<td noWrap><asp:imagebutton id="btnStart" accessKey="z" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/start_up.gif"
				ToolTip="First Question" CssClass="btnSpace" CausesValidation="False" AlternateText="First Question"></asp:imagebutton><asp:imagebutton id="btnPrev" accessKey="x" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/prev_up.gif"
				ToolTip="Previous Question" CssClass="btnSpace" CausesValidation="False" AlternateText="Previous Question"></asp:imagebutton><asp:imagebutton id="btnReload" accessKey="c" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/ig_menuCurrent_up.png"
				ToolTip="Reload Question" CssClass="btnSpace" CausesValidation="False" AlternateText="Reload Question"></asp:imagebutton><asp:imagebutton id="btnNext" accessKey="v" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/next_up.gif"
				ToolTip="Next Question" CssClass="btnSpace" CausesValidation="False" AlternateText="Next Question"></asp:imagebutton><asp:imagebutton id="btnLast" accessKey="b" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/end_up.gif"
				ToolTip="Last Question" CssClass="btnSpace" CausesValidation="False" AlternateText="Last Question"></asp:imagebutton><asp:imagebutton id="btnNew" accessKey="n" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/add_up.gif"
				ToolTip="New Question" CssClass="btnSpace" CausesValidation="False" AlternateText="New Question"></asp:imagebutton></td>
		<td style="WIDTH: 100%" align="right"><asp:button id="btnTop" runat="server" ToolTip="Back to Top" CssClass="small" Text="^ Top ^"
				BackColor="Transparent" BorderStyle="None" Font-Underline="True" ForeColor="Blue"></asp:button></td>
	</tr>
	<tr>
		<td><asp:image id="hlpQuestionType" runat="server" ImageUrl="./images/help.png" ToolTip="Select a question type to change the current type of question you are editing."></asp:image></td>
		<td noWrap align="right"><asp:label id="lblQuestionType" Runat="server" CssClass="boldContent">Question Type:</asp:label></td>
		<td align="left"><asp:dropdownlist id="ddlQuestionType" runat="server" AutoPostBack="True" CssClass="content"></asp:dropdownlist></td>
	</tr>
</table>
<% if (Session("SelectedQuestionType") <> "*" AND Session("SelectedQuestionType") <> "") %>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="0" cellPadding="2" class="borderless">
	<% if (Session("SelectedQuestionType") <> "H") %>
	<tr>
		<td vAlign="top"><asp:image id="hlpQuestion" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the question text."></asp:image></td>
		<td vAlign="top" align="right"><asp:label id="lblQuestion" Runat="server" CssClass="boldContent">Question:</asp:label></td>
		<td align="left"><asp:textbox id="txtQuestion" runat="server" CssClass="content" Width="400px" Height="200px"
				MaxLength="1800" TextMode="MultiLine" Rows="15" Columns="50"></asp:textbox></td>
	</tr>
	<% end if %>
	<% if (Session("SelectedQuestionType") = "R") %>
	<TR>
		<TD vAlign="top"><asp:image id="hlpTopIsNotApp" runat="server" ImageUrl="./images/help.png" ToolTip="Check this box to make the top value Not Applicable (NA)."></asp:image></TD>
		<TD vAlign="top" align="right"><asp:checkbox id="chkTopIsNotApp" runat="server" CssClass="content"></asp:checkbox></TD>
		<TD align="left"><asp:label id="lblTopIsNotApp" Runat="server" CssClass="boldContent">Top value is not applicable?</asp:label></TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpInitial" runat="server" ImageUrl="./images/help.png" ToolTip="Fill in the initial value of the rating."></asp:image></TD>
		<TD vAlign="top" align="right"><asp:label id="lblInitial" Runat="server" CssClass="boldContent">Initial Value:</asp:label></TD>
		<TD align="left"><igtxt:webnumericedit id="wneInitial" runat="server" CssClass="content" JavaScriptFileNameCommon="./MyInfragisticsScripts/ig_shared.js"
				JavaScriptFileName="./MyInfragisticsScripts/ig_edit.js" ImageDirectory="./MyInfragisticsImages/images/" DataMode="Int"
				Section508Compliant="True" MaxValue="2147483647"></igtxt:webnumericedit></TD>
	</TR>
	<TR>
		<TD style="HEIGHT: 37px" vAlign="top"><asp:image id="hlpNumber" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the number of rating fields to provide."></asp:image></TD>
		<TD style="HEIGHT: 37px" vAlign="top" align="right"><asp:label id="lblValues" Runat="server" CssClass="boldContent">Number of Values:</asp:label></TD>
		<TD style="HEIGHT: 37px" align="left"><igtxt:webnumericedit id="wneCount" runat="server" CssClass="content" JavaScriptFileNameCommon="./MyInfragisticsScripts/ig_shared.js"
				JavaScriptFileName="./MyInfragisticsScripts/ig_edit.js" ImageDirectory="./MyInfragisticsScripts/images/" DataMode="Int" Section508Compliant="True"
				MaxValue="2147483647" MinValue="0"></igtxt:webnumericedit></TD>
	</TR>
	<TR>
		<TD vAlign="top"><asp:image id="hlpInterval" runat="server" ImageUrl="./images/help.png" ToolTip="This is the value that will be added to each preceding value starting with the initial value."></asp:image></TD>
		<TD vAlign="top" align="right"><asp:label id="lblInterval" Runat="server" CssClass="boldContent">Interval Value:</asp:label></TD>
		<TD align="left"><igtxt:webnumericedit id="wneInterval" runat="server" CssClass="content" JavaScriptFileNameCommon="./MyInfragisticsScripts/ig_shared.js"
				JavaScriptFileName="./MyInfragisticsScripts/ig_edit.js" ImageDirectory="./MyInfragisticsScripts/Images/" DataMode="Int"
				Section508Compliant="True" MaxValue="2147483647" MinValue="1"></igtxt:webnumericedit></TD>
	</TR>
	<% end if %>
	<% if (Session("SelectedQuestionType") = "S" or Session("SelectedQuestionType") = "C" or _
		Session("SelectedQuestionType") = "L" or Session("SelectedQuestionType") = "P" or _
		Session("SelectedQuestionType") = "R" or Session("SelectedQuestionType") = "M" or _
		Session("SelectedQuestionType") = "T" or Session("SelectedQuestionType") = "D" or _
		Session("SelectedQuestionType") = "N" or Session("SelectedQuestionType") = "Z") %>
	<% if (Session("SelectedQuestionType") = "Z") then %>
	<tr>
		<td><asp:image id="hlpQuerySelect" runat="server" ImageUrl="./images/help.png" ToolTip="Select the type of information that you want put into this prepopulated list."></asp:image></td>
		<td noWrap align="right"><asp:label id="lblQuerySelect" Runat="server" CssClass="boldContent">Query Select:</asp:label></td>
		<td align="left"><asp:dropdownlist id="ddlQuerySelect" runat="server" AutoPostBack="False" CssClass="content"></asp:dropdownlist></td>
	</tr>
	<% end if %>
	<TR>
		<TD vAlign="top"><asp:image id="hlpRequired" runat="server" ImageUrl="./images/help.png" ToolTip="Check this box to force the user to answer the question."></asp:image></TD>
		<TD vAlign="top" align="right"><asp:checkbox id="chkbxRequired" runat="server" CssClass="content"></asp:checkbox></TD>
		<TD align="left"><asp:label id="lblRequired" Runat="server" CssClass="boldContent">Required Question?</asp:label></TD>
	</TR>
	<% if (Session("SelectedQuestionType") <> "M") then %>
	<TR>
		<TD style="HEIGHT: 24px" vAlign="top"><asp:image id="hlpFilter" runat="server" ImageUrl="./images/help.png" ToolTip="Uncheck this box to exclude this question from the filter list on the reports page."></asp:image></TD>
		<TD style="HEIGHT: 24px" vAlign="top" align="right"><asp:checkbox id="chkbxFilter" runat="server" CssClass="content" Checked="True"></asp:checkbox></TD>
		<TD style="HEIGHT: 24px" align="left"><asp:label id="lblFilter" Runat="server" CssClass="boldContent">Allow Filter?</asp:label></TD>
	</TR>
	<% end if %>
	<TR>
		<TD vAlign="top"><asp:image id="hlpPageBreak" runat="server" ImageUrl="./images/help.png" ToolTip="Check this box to force a page break before this question."></asp:image></TD>
		<TD vAlign="top" align="right"><asp:checkbox id="chkbxPageBreak" runat="server" CssClass="content"></asp:checkbox></TD>
		<TD align="left"><asp:label id="lblPageBreak" Runat="server" CssClass="boldContent">Page Break?</asp:label></TD>
	</TR>
	<% if (Session("SelectedQuestionType") <> "M" and Session("SelectedQuestionType") <> "S") %>
	<tr>
		<td style="HEIGHT: 22px"><asp:image id="hlpGraphType" runat="server" ImageUrl="./images/help.png" ToolTip="Choose a graph type for the question.  The default is a pie graph."></asp:image></td>
		<td style="HEIGHT: 22px" noWrap align="right"><asp:label id="lblGraphType" Runat="server" CssClass="boldContent">Graph Type:</asp:label></td>
		<td style="HEIGHT: 23px" align="left"><asp:dropdownlist id="ddlGraphType" runat="server" AutoPostBack="False" CssClass="content"></asp:dropdownlist></td>
	</tr>
	<% end if %>
	<% if (CType(Session("QuestionGroupsExist"),boolean) and Session("SelectedQuestionType") <> "M" and Session("SelectedQuestionType") <> "S") %>
	<tr>
		<td><asp:image id="hlpQuestionGroup" runat="server" ImageUrl="./images/help.png" ToolTip="Choose a question group.  Question groups allow you to group questions together into a multi-graph on the reports page."></asp:image></td>
		<td noWrap align="right"><asp:label id="lblQuestionGroup" Runat="server" CssClass="boldContent">Question Group:</asp:label></td>
		<td align="left"><asp:dropdownlist id="ddlQuestionGroup" runat="server" AutoPostBack="False" CssClass="content"></asp:dropdownlist></td>
	</tr>
	<% end if %>
	<tr>
		<TD style="HEIGHT: 29px" vAlign="top"><asp:image id="hlpBranch" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the question number the survey should skip to when the user answers this question.  If you do not specify a question to skip to then a default value of 0 will be entered and will be ignored by the survey page."></asp:image></TD>
		<TD style="HEIGHT: 29px" vAlign="top" align="right"><asp:label id="lblBranch" Runat="server" CssClass="boldContent">Skip to:</asp:label></TD>
		<TD style="HEIGHT: 29px" align="left"><igtxt:webnumericedit id="wneBranch" runat="server" ToolTip="Enter the question number to branch to."
				CssClass="content" Width="152px" JavaScriptFileNameCommon="./MyInfragisticsScripts/ig_shared.js" JavaScriptFileName="./MyInfragisticsScripts/ig_edit.js"
				ImageDirectory="./MyInfragisticsScripts/Images/" DataMode="Int" Section508Compliant="True" MaxValue="1000" MinValue="0"></igtxt:webnumericedit></TD>
	</tr>
	<% else %>
	<TR>
		<TD vAlign="top"><asp:image id="hlpPageBreak2" runat="server" ImageUrl="./images/help.png" ToolTip="Check this box to force a page break before this question."></asp:image></TD>
		<TD vAlign="top" align="right"><asp:checkbox id="chkbxPageBreak2" runat="server" CssClass="content"></asp:checkbox></TD>
		<TD align="left"><asp:label id="lblPageBreak2" Runat="server" CssClass="boldContent">Page Break?</asp:label></TD>
	</TR>
	<% end if %>
</table>
<% if (((Session("SelectedQuestionType") = "C" or Session("SelectedQuestionType") = "L" or _
	Session("SelectedQuestionType") = "P" or Session("SelectedQuestionType") = "T") and _
	Session("intQuestion") <> -1)) %>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table class="borderless">
	<TBODY>
		<tr>
			<td style="WIDTH: 19px" vAlign="top"><asp:image id="hlpListItems" runat="server" ImageUrl="./images/help.png" ToolTip="Select one of the list items to view it."></asp:image></td>
			<td vAlign="top" align="right"><asp:label id="lblListItems" Runat="server" CssClass="boldContent">List Items:</asp:label></td>
			<td vAlign="top" align="left" style="padding-left:20px;"><asp:repeater id="rptrListItems" runat="server">
					<ItemTemplate>
						<% if (((Session("UserType") >= 2 and (CType(Session("isTemplateOwner"), Boolean) or CType(Session("isTemplateDelegate"), Boolean)) and Session("Results") <= 0) or Session("UserType") = 4)) then %>
						<li>
							<a title='<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "LIST_ITEM_TITLE"))%>' href='./List.aspx?seqTemplate=<%=Session("seqTemplate")%>&intQuestion=<%=Session("intQuestion")%>&intList=<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "LIST_ITEM_ID")) %>'>
								<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "LIST_ITEM_TITLE")) %>
							</a>
						</li>
						<% else %>
						<li>
							<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "LIST_ITEM_TITLE")) %>
						</li>
						<% end if %>
					</ItemTemplate>
					<HeaderTemplate>
						<ol>
					</HeaderTemplate>
					<FooterTemplate>
						<% if (((Session("UserType") >= 2 and (CType(Session("isTemplateOwner"), Boolean) or CType(Session("isTemplateDelegate"), Boolean)) and Session("Results") <= 0) or Session("UserType") = 4)) then %>
						<li>
							<a title="Create New List Item" href='./List.aspx?seqTemplate=<%=Session("seqTemplate")%>&intQuestion=<%=Session("intQuestion")%>&
							intList=-1'>Create New List Item</a>
						</li>
						<% end if %>
						</ol>
					</FooterTemplate>
					<SeparatorTemplate>
					</SeparatorTemplate>
				</asp:repeater></td>
		</tr>
	</TBODY>
</table>
<% end if %>
<% end if %>
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
