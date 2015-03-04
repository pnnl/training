<%@ Page Language="vb" trace="false" MasterPageFile="~/MasterPage.master"  AutoEventWireup="false" Codebehind="Template.aspx.vb" Inherits="SurveyAdmin.Template" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder" Runat="Server">
<%=Session("JSAction")%>
<form id="Form1" method="post" runat="server">
<h1>Generate Report</h1>
<p></p>
<script type="text/javascript">
	function isOwner(Owner)
	{
		if(Owner == "True"){
			document.write('<img src="./images/Owner.png" alt="This user is an Owner of this Template.">');}
	}
</script>
<table class="borderless">
    <% if (Session("UserType") >= 3 and Session("seqTemplate") <> -1) then %>
	<tr>
		<% if (((Session("UserType") > 2 and (Session("isTemplateOwner") or Session("isTemplateDelegate"))) or Session("UserType") = 4)) then %>
		    <% if ((Session("isDirty") and Session("Results") <= 0)) then %>
		        <TD><asp:image id="hlpCopy" runat="server" ImageUrl="./images/help.png"></asp:image></TD>
		        <td><asp:button id="btnCopy" Runat="server" CssClass="button"></asp:button></td>
		    <% else %>
		        <td></td>
		        <td><asp:label id="Label5" Runat="server" CssClass="boldContent">Your Template is Approved.</asp:label></td>
		    <% end if %>
		<% end if %>
	</tr>
	<% end if %>
	<tr>
		<td><asp:image id="hlpSelectTemplate" runat="server" ImageUrl="./images/help.png" ToolTip="Select a template to edit or create your own below."></asp:image></td>
		<% if (Session("TemplatesExist") = "Yes") then %>
		<td colSpan="4"><asp:dropdownlist id="ddlTemplateList" runat="server" AutoPostBack="True" CssClass="content"></asp:dropdownlist></td>
		<% else %>
		<td colSpan="4"><asp:label id="lblSelectTemplate" Runat="server" CssClass="content" ForeColor="Red">There are currently no Templates available for your use.</asp:label></td>
		<% end if %>
	</tr>
</table>
<% if (Session("TemplatesExist") = "Yes" or Session("UserType") >= 3 ) then %>
<% if (Session("seqTemplate") <> -1 or Session("UserType") >= 2) then %>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table class="borderless">
	<tr>
		<td><asp:image id="hlpTemplateName" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the name of the Template."></asp:image></td>
		<td align="right"><asp:label id="lblTemplateSelect" Runat="server" CssClass="boldContent">Name:</asp:label></td>
		<td colSpan="3"><asp:textbox id="txtTemplateName" Runat="server" CssClass="content" MaxLength="50" Width="400px"></asp:textbox></td>
	</tr>
	<tr>
		<TD vAlign="top"><asp:image id="hlpTemplateDescription" runat="server" ImageUrl="./images/help.png" ToolTip="Enter the description of the Template here."></asp:image></TD>
		<td vAlign="top" align="right"><asp:label id="lblTemplateDescription" Runat="server" CssClass="boldContent">Description:</asp:label></td>
		<td colSpan="3"><asp:textbox id="txtDescription" Runat="server" CssClass="content" MaxLength="1800" Height="200px"
				Width="400px" TextMode="MultiLine" Rows="15" Columns="50"></asp:textbox></td>
	</tr>
	<tr>
		<TD><asp:image id="hlpTrainingSurvey" runat="server" ImageUrl="./images/help.png" ToolTip="Is this a Training &amp; Qualifications Survey?"></asp:image></TD>
		<td noWrap align="right"><asp:checkbox id="chkTQSurvey" runat="server" CssClass="content"></asp:checkbox></td>
		<td noWrap align="left" colSpan="3"><asp:label id="lblTrainingSurvey" Runat="server" CssClass="boldContent">PNNL Training Survey.</asp:label></td>
	</tr>
	<tr>
		<TD><asp:image id="hlpSurveyOptional" runat="server" ImageUrl="./images/help.png" ToolTip="Will surveys for this Template be optional?"></asp:image></TD>
		<td noWrap align="right"><asp:checkbox id="chkOptional" runat="server" CssClass="content"></asp:checkbox></td>
		<td noWrap align="left" colSpan="3"><asp:label id="lblSurveyOptional" Runat="server" CssClass="boldContent">Surveys created from this Template may be optional.</asp:label></td>
	</tr>
	<tr>
		<TD><asp:image id="hlpTemplateActive" runat="server" ImageUrl="./images/help.png" ToolTip="Is the Template active?"></asp:image></TD>
		<td noWrap align="right"><asp:checkbox id="chkActive" runat="server" CssClass="content" Checked="True"></asp:checkbox></td>
		<td align="left" colSpan="3"><asp:label id="lblTemplateActive" Runat="server" CssClass="boldContent">Active?</asp:label></td>
	</tr>
	<% if (Session("UserType") > 2 and Session("seqTemplate") <> -1) then %>
	<tr>
		<td></td>
		<TD align="right"><% If (Session("isTemplateOwner") Or Session("isTemplateDelegate")) Or Session("UserType") = 4 Then%><asp:image id="hlpSetPermissions" runat="server" ImageUrl="./images/help.png" ToolTip="Click this button to set who may use this Template."></asp:image><% End If%></TD>
		<td align="left"><% If (Session("isTemplateOwner") Or Session("isTemplateDelegate")) Or Session("UserType") = 4 Then%><asp:button id="btnSetPermissions" Runat="server" CssClass="button" Text="Set Permissions" ToolTip="Set Permissions"></asp:button><% end if %></td>
		<td align="right"><asp:image id="hlpTemplatePreview" runat="server" ImageUrl="./images/help.png" ToolTip="Click this button to preview this Template as it will appear for users."></asp:image></td>
		<td align="left"><asp:button id="btnTemplatePreview" Runat="server" CssClass="button" Text="Preview" ToolTip="Preview"></asp:button></td>
	</tr>
	<% elseif (Session("seqTemplate") <> -1) %>
	<tr>
		<td align="right"><asp:image id="hlpTemplatePreview2" runat="server" ImageUrl="./images/help.png" ToolTip="Click this button to preview this Template as it will appear for users."></asp:image></td>
		<td></td>
		<td align="left"><asp:button id="btnTemplatePreview2" Runat="server" CssClass="button" Text="Preview" ToolTip="Preview"></asp:button></td>
	</tr>
	<% end if %>
</table>
<% if (Session("seqTemplate") <> -1) then %>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table class="borderless">
	<tr>
		<TD vAlign="top"><asp:image id="hlpTags" runat="server" ImageUrl="./images/help.png" ToolTip="Tag Items are your questions for the creators of surveys using this Template."></asp:image></TD>
		<td vAlign="top" noWrap><asp:label id="Label1" Runat="server" CssClass="boldContent">Tag Items:</asp:label></td>
		<td noWrap style="padding-left:20px;"><asp:repeater id="rptrTags" runat="server">
				<ItemTemplate>
					<% if (((Session("UserType") >= 2 and (CType(Session("isTemplateOwner"), Boolean) or CType(Session("isTemplateDelegate"), Boolean)) and Session("Results") <= 0) or Session("UserType") = 4)) then %>
					<li>
						<a title='<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "TEMPLATE_TAG_TITLE")) %>' href='./Tag.aspx?seqTemplate=<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "TEMPLATE_KEY")) %>&amp;intTagID=<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "TEMPLATE_TAG_ID")) %>'>
							<%#Convert.ToString(DataBinder.Eval(Container.DataItem, "TEMPLATE_TAG_TITLE"))%>
						</a>
					</li>
					<% else %>
					<li>
						<%#Convert.ToString(DataBinder.Eval(Container.DataItem, "TEMPLATE_TAG_TITLE"))%>
					</li>
					<% end if %>
				</ItemTemplate>
				<HeaderTemplate>
					<ol>
				</HeaderTemplate>
				<FooterTemplate>
					<% if (((Session("UserType") > 2 and (CType(Session("isTemplateOwner"), Boolean) or CType(Session("isTemplateDelegate"), Boolean)) and Session("Results") <= 0) or Session("UserType") = 4) and Session("Machine") <> "Production") then %>
					<li>
						<a title="New Tag Item" href='./Tag.aspx?seqTemplate=<%=Session("seqTemplate")%>&amp;intTagID=-1'>
							Create New Tag Item</a>
					</li>
					<% end if %>
					</ol>
				</FooterTemplate>
				<SeparatorTemplate>
				</SeparatorTemplate>
			</asp:repeater></td>
	</tr>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table class="borderless">
	<tr>
		<TD vAlign="top"><asp:image id="hlpQuestions" runat="server" ImageUrl="./images/help.png" ToolTip="Questions are the questions that will be asked of those taking surveys created from this Template."></asp:image></TD>
		<td vAlign="top" noWrap><asp:label id="Label2" Runat="server" CssClass="boldContent">Questions:</asp:label></td>
		<td noWrap style="padding-left:20px;"><asp:repeater id="rptrQuestions" runat="server">
				<ItemTemplate>
					<% if (((Session("UserType") >= 2 and (CType(Session("isTemplateOwner"), Boolean) or CType(Session("isTemplateDelegate"), Boolean)) and Session("Results") <= 0) or Session("UserType") = 4)) then %>
					<li>
						<a title='<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "QUESTION_TEXT")) %>' href='./Question.aspx?seqTemplate=<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "TEMPLATE_KEY")) %>&amp;intQuestion=<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "QUESTION_ID")) %>'>
							<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "QUESTION_TEXT")) %>
						</a>
					</li>
					<% else %>
					<li>
						<%#Convert.ToString(DataBinder.Eval(Container.DataItem, "QUESTION_TEXT"))%>
					</li>
					<% end if %>
				</ItemTemplate>
				<HeaderTemplate>
					<ol>
				</HeaderTemplate>
				<FooterTemplate>
					<% if (((Session("UserType") > 2 and (CType(Session("isTemplateOwner"), Boolean) or CType(Session("isTemplateDelegate"), Boolean)) and Session("Results") <= 0) or Session("UserType") = 4) and Session("Machine") <> "Production") then %>
					<li>
						<a title="Create New Question" href='./Question.aspx?seqTemplate=<%=Session("seqTemplate")%>&amp;intQuestion=-1'>
							Create New Question</a>
					</li>
					<% end if %>
					</ol>
				</FooterTemplate>
				<SeparatorTemplate>
				</SeparatorTemplate>
			</asp:repeater></td>
	</tr>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table class="borderless">
	<tr>
		<TD vAlign="top"><asp:image id="hlpQuestionGroups" runat="server" ImageUrl="./images/help.png" ToolTip="Questions groups are the groups that questions will appear in on the report page if you create question groups and assign them to questions on the Questions page."></asp:image></TD>
		<td vAlign="top" noWrap><asp:label id="lblQuestionGroups" Runat="server" CssClass="boldContent">Question<br />&nbsp;&nbsp;&nbsp;&nbsp;Groups:</asp:label></td>
		<td noWrap style="padding-left:20px;"><asp:repeater id="rptrQuestionGroups" runat="server">
				<ItemTemplate>
					<% if (((Session("UserType") >= 2 and (CType(Session("isTemplateOwner"), Boolean) or CType(Session("isTemplateDelegate"), Boolean)) and Session("Results") <= 0) or Session("UserType") = 4)) then %>
					<li>
						<a title='<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "QUESTION_GROUP_TITLE")) %>' href='./QuestionGroups.aspx?seqTemplate=<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "TEMPLATE_KEY")) %>&amp;intReportGroup=<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "QUESTION_GROUP_ID")) %>'>
							<%#Convert.ToString(DataBinder.Eval(Container.DataItem, "QUESTION_GROUP_TITLE"))%>
						</a>
					</li>
					<% else %>
					<li>
						<%#Convert.ToString(DataBinder.Eval(Container.DataItem, "QUESTION_GROUP_TITLE"))%>
					</li>
					<% end if %>
				</ItemTemplate>
				<HeaderTemplate>
					<ol>
				</HeaderTemplate>
				<FooterTemplate>
					<% if (((Session("UserType") > 2 and (CType(Session("isTemplateOwner"), Boolean) or CType(Session("isTemplateDelegate"), Boolean)) and Session("Results") <= 0) or Session("UserType") = 4) and Session("Machine") <> "Production") then %>
					<li>
						<a title="Create New Question Group" href='./QuestionGroups.aspx?seqTemplate=<%=Session("seqTemplate")%>&amp;intReportGroup=-1'>
							Create New Question Group</a>
					</li>
					<% end if %>
					</ol>
				</FooterTemplate>
				<SeparatorTemplate>
				</SeparatorTemplate>
			</asp:repeater></td>
	</tr>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table class="borderless">
	<tr>
		<TD vAlign="top"><asp:image id="hlpSurveys" runat="server" ImageUrl="./images/help.png" ToolTip="Surveys are the surveys crafted from this Template by users of this Template."></asp:image></TD>
		<td vAlign="top" noWrap><asp:label id="Label3" Runat="server" CssClass="boldContent">Surveys:&nbsp;&nbsp;&nbsp;</asp:label></td>
		<td noWrap style="padding-left:20px;"><asp:repeater id="rptrSurveys" runat="server">
				<ItemTemplate>
					<% if (Session("UserType") >= 2) then %>
					<li>
						<a title='<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "SURVEY_COMMENT")) %>' href='./Survey.aspx?seqTemplate=<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "TEMPLATE_KEY")) %>&amp;seqSurvey=<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "SURVEY_KEY")) %>'>
							<%#Convert.ToString(DataBinder.Eval(Container.DataItem, "SURVEY_COMMENT"))%>
						</a>
					</li>
					<% else %>
					<li>
						<%#Convert.ToString(DataBinder.Eval(Container.DataItem, "SURVEY_COMMENT"))%>
					</li>
					<% end if %>
				</ItemTemplate>
				<HeaderTemplate>
					<ol>
				</HeaderTemplate>
				<FooterTemplate>
					<% if (((Session("UserType") >= 2 and CType(Session("isTemplateUser"), Boolean)) or Session("UserType") = 4) and Session("Machine") <> "Production") %>
					<li>
						<a title="New Survey" href='./Survey.aspx?seqTemplate=<%=Session("seqTemplate")%>&amp;seqSurvey=-1'>
							Create New Survey</a>
					</li>
					<% end if %>
					</ol>
				</FooterTemplate>
				<SeparatorTemplate>
				</SeparatorTemplate>
			</asp:repeater></td>
	</tr>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table class="borderless">
	<tr>
		<TD vAlign="top"><asp:image id="hlpDelegates" runat="server" ImageUrl="./images/help.png" ToolTip="Delegates are those that have the ability to alter your Template.  You may also assign ownership of this Template to others.  Owners will have full access to this Template and may remove it. "></asp:image></TD>
		<td vAlign="top" noWrap><asp:label id="Label4" Runat="server" CssClass="boldContent">Delegates:</asp:label></td>
		<td noWrap style="padding-left:20px;"><asp:repeater id="rptrDelegates" runat="server">
				<ItemTemplate>
					<% if (((Session("UserType") > 2 And Session("isTemplateOwner") And Session("isTemplateDelegate")) Or Session("UserType") = 4) And Session("Machine") <> "Production") then %>
					<li>
						<a title='<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "FIRST_NAME")) %>	<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "MIDDLE_NAME")) %> <%# Convert.ToString(DataBinder.Eval(Container.DataItem, "LAST_NAME")) %> - <%# Convert.ToString(DataBinder.Eval(Container.DataItem, "HANFORD_ID")) %> - <%# Convert.ToString(DataBinder.Eval(Container.DataItem, "EMAIL_ADDRESS")) %>' href='./TemplateDelegates.aspx?seqTemplate=<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "TEMPLATE_KEY")) %>&amp;seqUserID=<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "USER_KEY")) %>'>
							<%#Convert.ToString(DataBinder.Eval(Container.DataItem, "FIRST_NAME"))%>
							<%#Convert.ToString(DataBinder.Eval(Container.DataItem, "MIDDLE_NAME"))%>
							<%#Convert.ToString(DataBinder.Eval(Container.DataItem, "LAST_NAME"))%>
							-
							<%#Convert.ToString(DataBinder.Eval(Container.DataItem, "HANFORD_ID"))%>
							-
							<%#Convert.ToString(DataBinder.Eval(Container.DataItem, "EMAIL_ADDRESS"))%>
						</a>
						<script type="text/javascript">isOwner("<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "OWNER_IND")) %>")</script>
					</li>
					<% else %>
					<li>
						<%#Convert.ToString(DataBinder.Eval(Container.DataItem, "FIRST_NAME"))%>
						<%#Convert.ToString(DataBinder.Eval(Container.DataItem, "MIDDLE_NAME"))%>
						<%#Convert.ToString(DataBinder.Eval(Container.DataItem, "LAST_NAME"))%>
						-
						<%#Convert.ToString(DataBinder.Eval(Container.DataItem, "HANFORD_ID"))%>
						-
						<%#Convert.ToString(DataBinder.Eval(Container.DataItem, "EMAIL_ADDRESS"))%>
						<script type="text/javascript">isOwner("<%# Convert.ToString(DataBinder.Eval(Container.DataItem, "OWNER_IND")) %>")</script>
					</li>
					<% end if %>
				</ItemTemplate>
				<HeaderTemplate>
					<ol>
				</HeaderTemplate>
				<FooterTemplate>
					<% if ((((CType(Session("isTemplateOwner"), Boolean)) and Session("UserType") > 2) or Session("UserType") = 4) and Session("Machine") <> "Production") then %>
					<li>
						<a title="Create New Delegate" href='./TemplateDelegates.aspx?seqTemplate=<%=Session("seqTemplate")%>&amp;seqUserID=-1'>
							Create New Delegate</a>
					</li>
					<% end if %>
					</ol>
				</FooterTemplate>
				<SeparatorTemplate>
				</SeparatorTemplate>
			</asp:repeater></td>
	</tr>
</table>
<% end if %>
<% if (Session("UserType") <> 1) then %>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table width="100%" align="center" class="borderless">
	<tr>
		<td>
		    <asp:button id="btnInsert" Runat="server" CssClass="button" ToolTip="Insert" Text="Insert"></asp:button>
		    <% If Session("PageOptions") <> 2 Then%>
		        <asp:button id="btnUpdate" Runat="server" CssClass="button" ToolTip="Update" Text="Update"></asp:button>
		        <asp:button id="btnDelete" Runat="server" CssClass="button" ToolTip="Delete" Text="Delete"></asp:button>
		    <% end if %>
		    <asp:button id="btnReset" Runat="server" CssClass="button" ToolTip="Reset" Text="Reset"></asp:button>
		</td>
	</tr>
</table>
<% end if %>
<% end if %>
<% end if %>
</form>
</asp:Content>
