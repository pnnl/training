<%@ Page trace="true" Language="vb" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" Codebehind="Groups.aspx.vb" Inherits="SurveyAdmin.Groups" smartNavigation="False" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder" Runat="Server">
<%=Session("JSAction")%>
<form id="Form1" method="post" runat="server">
<h1>Template Survey Distribution Groups</h1>
<p></p>
<table class="borderless">
	<tr>
		<td align="left" style="HEIGHT: 18px"><asp:label id="lblTemplate" Runat="server" CssClass="content">Template:</asp:label></td>
		<td class="content" align="left" style="HEIGHT: 18px"><%=Session("TemplateName")%></td>
		<td align="center" width="50" style="HEIGHT: 18px"></td>
		<td align="right" style="HEIGHT: 18px"></td>
		<td align="left" style="HEIGHT: 18px"></td>
	</tr>
	<tr>
		<td align="right"><asp:label id="lblQuestion" Runat="server" CssClass="content">Survey:</asp:label></td>
		<td class="content" align="left"><%=Session("SurveyName")%></td>
		<td align="center" width="50"></td>
		<td align="right"><asp:image id="hlpTemplatePreview" runat="server" ImageUrl="./images/help.png" ToolTip="Click this button to preview the template."></asp:image></td>
		<td align="left"><asp:button id="btnTemplatePreview" runat="server" CssClass="button" Text="Preview" ToolTip="Preview"></asp:button></td>
	</tr>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="0" cellPadding="2" class="borderless">
	<tr class="navPanel">
		<td style="WIDTH: 21px"></td>
		<td noWrap align="right"><asp:label id="lblNavBar" runat="server" CssClass="lblNavPanel"></asp:label></td>
		<td style="WIDTH: 305px"><asp:imagebutton id="btnStart" accessKey="z" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/start_up.gif"
				CssClass="btnSpace" ToolTip="First Group" AlternateText="First Group" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnPrev" accessKey="x" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/prev_up.gif"
				CssClass="btnSpace" ToolTip="Previous Group" AlternateText="Previous Group" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnReload" accessKey="c" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/ig_menuCurrent_up.png"
				CssClass="btnSpace" ToolTip="Reload Group" AlternateText="Reload Group" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnNext" accessKey="v" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/next_up.gif"
				CssClass="btnSpace" ToolTip="Next Group" AlternateText="Next Group" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnLast" accessKey="b" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/end_up.gif"
				CssClass="btnSpace" ToolTip="Last Group" AlternateText="Last Group" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="btnNew" accessKey="n" runat="server" ImageUrl="./MyInfragisticsImages/XpBlue/add_up.gif"
				CssClass="btnSpace" ToolTip="New Group" AlternateText="New Group" CausesValidation="False"></asp:imagebutton></td>
		<td style="WIDTH: 100%" align="right"><asp:button id="btnUp" runat="server" CssClass="small" Text="^ Up ^" BackColor="Transparent"
				BorderStyle="None" Font-Underline="True" ForeColor="Blue" ToolTip="Back Up One Level"></asp:button></td>
	</tr>
	<tr>
		<td style="WIDTH: 21px"><asp:image id="hlpGroupName" runat="server" ImageUrl="./images/help.png" ToolTip="Contains the name of the group you are creating."></asp:image></td>
		<td noWrap align="right"><asp:label id="lblGroupName" Runat="server" CssClass="boldContent">Group Name:</asp:label></td>
		<td align="left"><asp:textbox id="txtGroupName" Runat="server" CssClass="content" Width="400px" MaxLength="8000"></asp:textbox></td>
	</tr>
	<tr>
		<td style="WIDTH: 21px; HEIGHT: 207px" vAlign="top"><asp:image id="hlpGroupList" runat="server" ImageUrl="./images/help.png" ToolTip="Enter a series of comma and/or line separated items.  These items may include E-mail addresses (but will not recognize group e-mail addresses), Hanford IDs, and Employee IDs."></asp:image></td>
		<td style="HEIGHT: 207px" vAlign="top" noWrap align="right"><asp:label id="lblGroupList" Runat="server" CssClass="boldContent">E-Mail List/<br>HID/Emplid:<br>(Separate your<br />items with commas<br /> and/or rows)</asp:label></td>
		<td style="HEIGHT: 207px" align="left"><asp:textbox id="txtGroupList" runat="server" CssClass="content" Width="400px" MaxLength="8000"
				Height="200px" TextMode="MultiLine" Rows="15" Columns="50"></asp:textbox></td>
	</tr>
</table>
<br>
<TABLE style="LEFT: 10px; TOP: 536px" width="100%" align="center" class="borderless">
	<tr>
		<td>
		    <% If Session("PageOptions") <> 2 Then%>
		        <asp:button id="btnInsert1" Runat="server" CssClass="button" ToolTip="Insert" Text="Insert"></asp:button>
		    <% end if %>
		    <% if session("PageOptions") <> 1 then %>
		        <asp:button id="btnUpdate1" Runat="server" CssClass="button" ToolTip="Update" Text="Update"></asp:button>
		        <asp:button id="btnDelete1" Runat="server" CssClass="button" ToolTip="Delete" Text="Delete"></asp:button>
		    <% end if %>
		    <asp:button id="btnReset1" Runat="server" CssClass="button" ToolTip="Reset" Text="Reset"></asp:button>
		</td>
	</tr>
</TABLE>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table cellSpacing="5" cellPadding="5" width="100%" align="center" class="borderless">
	<% if (Session("GroupUserListRows") <> "None") then %>
	<TBODY>
		<TR align="center">
			<TD vAlign="top" colSpan="3">
				<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
				<asp:label id="lblInGroup" runat="server" CssClass="boldContent" ForeColor="Red" Width="100%">Survey Recipients in:</asp:label><br>
				<font class="boldContentRed">
					<%=Session("GroupName")%>
				</font>
			</TD>
		</TR>
		<tr>
			<td vAlign="top" width="20"><asp:image id="hlpLevel0" runat="server" ImageUrl="./images/help.png" ToolTip="Uncheck users to remove them from the group."></asp:image></td>
			<td vAlign="top" align="left" colSpan="2"><asp:datagrid id="dgGroup" runat="server" Width="100%" CellPadding="0" CellSpacing="1" AutoGenerateColumns="False">
					<AlternatingItemStyle BackColor="#A7DCD2"></AlternatingItemStyle>
					<ItemStyle CssClass="content"></ItemStyle>
					<HeaderStyle HorizontalAlign="Left" BorderWidth="5px" BorderStyle="Outset" CssClass="content"
						BackColor="#60C0B4"></HeaderStyle>
					<Columns>
						<asp:TemplateColumn HeaderText="&lt;a href=&quot;./Groups.aspx?selectAll=1&quot;&gt;Select All&lt;/a&gt;">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
							<HeaderTemplate>
								<asp:Button id="Button1" runat="server" CssClass="button" Text="Select All Users" OnCommand="commandBtnClick"
									CommandName="SelectGroup" CommandArgument="commandBtnClick"></asp:Button>
							</HeaderTemplate>
							<ItemTemplate>
								<asp:CheckBox id="Checkbox1" Runat="server"></asp:CheckBox>
							</ItemTemplate>
						</asp:TemplateColumn>
						<asp:BoundColumn Visible="False" DataField="seqUserID" SortExpression="seqUserID" HeaderText="seqUserID"></asp:BoundColumn>
						<asp:BoundColumn Visible="False" DataField="intUserType" SortExpression="intUserType" HeaderText="intUserType"></asp:BoundColumn>
						<asp:BoundColumn DataField="Name" SortExpression="Name" HeaderText="User Name"></asp:BoundColumn>
						<asp:BoundColumn DataField="strHID" SortExpression="strHID" HeaderText="Hanford ID"></asp:BoundColumn>
						<asp:BoundColumn DataField="strEmail" SortExpression="strEmail" HeaderText="Email Address"></asp:BoundColumn>
					</Columns>
				</asp:datagrid></td>
		</tr>
		<% else %>
		<tr align="center">
			<td colSpan="3">
				<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
				<asp:label id="lblLevel0None" runat="server" CssClass="boldContent" BackColor="LightGrey" ForeColor="Red"
					Font-Bold="True">There are currently no survey recipients in this group.</asp:label></td>
		</tr>
		<% end if %>
	</TBODY>
</table>
<% if (Session("Machine") = "Development") %>
<table cellSpacing="5" cellPadding="5" width="100%" align="center" class="borderless">
	<TBODY>
		<TR align="center">
			<TD vAlign="top" colSpan="3">
				<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
				<asp:label id="Label1" runat="server" CssClass="boldContent" ForeColor="Red" Width="100%">All available survey respondents</asp:label></TD>
		</TR>
		<tr class="navPanel">
			<td class="navPanelLeft"></td>
			<td vAlign="middle" align="center" colSpan="3"><asp:imagebutton id="imgA1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/A.png" CssClass="btnSpace"
					ToolTip="List all last names starting with A" AlternateText="A" CausesValidation="False" OnCommand="commandGridUserList" CommandName="A"></asp:imagebutton><asp:imagebutton id="imgB1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/B.png" CssClass="btnSpace"
					ToolTip="List all last names starting with B" AlternateText="B" CausesValidation="False" OnCommand="commandGridUserList" CommandName="B"></asp:imagebutton><asp:imagebutton id="imgC1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/C.png" CssClass="btnSpace"
					ToolTip="List all last names starting with C" AlternateText="C" CausesValidation="False" OnCommand="commandGridUserList" CommandName="C"></asp:imagebutton><asp:imagebutton id="imgD1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/D.png" CssClass="btnSpace"
					ToolTip="List all last names starting with D" AlternateText="D" CausesValidation="False" OnCommand="commandGridUserList" CommandName="D"></asp:imagebutton><asp:imagebutton id="imgE1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/E.png" CssClass="btnSpace"
					ToolTip="List all last names starting with E" AlternateText="E" CausesValidation="False" OnCommand="commandGridUserList" CommandName="E"></asp:imagebutton><asp:imagebutton id="imgF1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/F.png" CssClass="btnSpace"
					ToolTip="List all last names starting with F" AlternateText="F" CausesValidation="False" OnCommand="commandGridUserList" CommandName="F"></asp:imagebutton><asp:imagebutton id="imgG1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/G.png" CssClass="btnSpace"
					ToolTip="List all last names starting with G" AlternateText="G" CausesValidation="False" OnCommand="commandGridUserList" CommandName="G"></asp:imagebutton><asp:imagebutton id="imgH1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/H.png" CssClass="btnSpace"
					ToolTip="List all last names starting with H" AlternateText="H" CausesValidation="False" OnCommand="commandGridUserList" CommandName="H"></asp:imagebutton><asp:imagebutton id="imgI1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/I.png" CssClass="btnSpace"
					ToolTip="List all last names starting with I" AlternateText="I" CausesValidation="False" OnCommand="commandGridUserList" CommandName="I"></asp:imagebutton><asp:imagebutton id="imgJ1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/J.png" CssClass="btnSpace"
					ToolTip="List all last names starting with J" AlternateText="J" CausesValidation="False" OnCommand="commandGridUserList" CommandName="J"></asp:imagebutton><asp:imagebutton id="imgK1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/K.png" CssClass="btnSpace"
					ToolTip="List all last names starting with K" AlternateText="K" CausesValidation="False" OnCommand="commandGridUserList" CommandName="K"></asp:imagebutton><asp:imagebutton id="imgL1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/L.png" CssClass="btnSpace"
					ToolTip="List all last names starting with L" AlternateText="L" CausesValidation="False" OnCommand="commandGridUserList" CommandName="L"></asp:imagebutton><asp:imagebutton id="imgM1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/M.png" CssClass="btnSpace"
					ToolTip="List all last names starting with M" AlternateText="M" CausesValidation="False" OnCommand="commandGridUserList" CommandName="M"></asp:imagebutton><asp:imagebutton id="imgN1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/N.png" CssClass="btnSpace"
					ToolTip="List all last names starting with N" AlternateText="N" CausesValidation="False" OnCommand="commandGridUserList" CommandName="N"></asp:imagebutton><asp:imagebutton id="imgO1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/O.png" CssClass="btnSpace"
					ToolTip="List all last names starting with O" AlternateText="O" CausesValidation="False" OnCommand="commandGridUserList" CommandName="O"></asp:imagebutton><asp:imagebutton id="imgP1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/P.png" CssClass="btnSpace"
					ToolTip="List all last names starting with P" AlternateText="P" CausesValidation="False" OnCommand="commandGridUserList" CommandName="P"></asp:imagebutton><asp:imagebutton id="imgQ1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/Q.png" CssClass="btnSpace"
					ToolTip="List all last names starting with Q" AlternateText="Q" CausesValidation="False" OnCommand="commandGridUserList" CommandName="Q"></asp:imagebutton><asp:imagebutton id="imgR1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/R.png" CssClass="btnSpace"
					ToolTip="List all last names starting with R" AlternateText="R" CausesValidation="False" OnCommand="commandGridUserList" CommandName="R"></asp:imagebutton><asp:imagebutton id="imgS1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/S.png" CssClass="btnSpace"
					ToolTip="List all last names starting with S" AlternateText="S" CausesValidation="False" OnCommand="commandGridUserList" CommandName="S"></asp:imagebutton><asp:imagebutton id="imgT1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/T.png" CssClass="btnSpace"
					ToolTip="List all last names starting with T" AlternateText="T" CausesValidation="False" OnCommand="commandGridUserList" CommandName="T"></asp:imagebutton><asp:imagebutton id="imgU1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/U.png" CssClass="btnSpace"
					ToolTip="List all last names starting with U" AlternateText="U" CausesValidation="False" OnCommand="commandGridUserList" CommandName="U"></asp:imagebutton><asp:imagebutton id="imgV1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/V.png" CssClass="btnSpace"
					ToolTip="List all last names starting with V" AlternateText="V" CausesValidation="False" OnCommand="commandGridUserList" CommandName="V"></asp:imagebutton><asp:imagebutton id="imgW1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/W.png" CssClass="btnSpace"
					ToolTip="List all last names starting with W" AlternateText="W" CausesValidation="False" OnCommand="commandGridUserList" CommandName="W"></asp:imagebutton><asp:imagebutton id="imgX1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/X.png" CssClass="btnSpace"
					ToolTip="List all last names starting with X" AlternateText="X" CausesValidation="False" OnCommand="commandGridUserList" CommandName="X"></asp:imagebutton><asp:imagebutton id="imgY1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/Y.png" CssClass="btnSpace"
					ToolTip="List all last names starting with Y" AlternateText="Y" CausesValidation="False" OnCommand="commandGridUserList" CommandName="Y"></asp:imagebutton><asp:imagebutton id="imgZ1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/Z.png" CssClass="btnSpace"
					ToolTip="List all last names starting with Z" AlternateText="Z" CausesValidation="False" OnCommand="commandGridUserList" CommandName="Z"></asp:imagebutton></td>
		</tr>
		<% if (Session("GeneralUserListRows") <> "None") then %>
		<tr>
			<td vAlign="top" width="20"><asp:image id="Image1" runat="server" ImageUrl="./images/help.png" ToolTip="Select users to generate new passwords for them."></asp:image></td>
			<td vAlign="top" align="left" colSpan="2"><asp:datagrid id="dgUsers" runat="server" Width="100%" CellPadding="0" CellSpacing="1" AutoGenerateColumns="False">
					<AlternatingItemStyle BackColor="#A7DCD2"></AlternatingItemStyle>
					<ItemStyle CssClass="content"></ItemStyle>
					<HeaderStyle HorizontalAlign="Left" BorderWidth="5px" BorderStyle="Outset" CssClass="content"
						BackColor="#60C0B4"></HeaderStyle>
					<Columns>
						<asp:TemplateColumn HeaderText="&lt;a href=&quot;./Group.aspx?selectAll=1&quot;&gt;Select All&lt;/a&gt;">
							<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
							<ItemStyle HorizontalAlign="Center"></ItemStyle>
							<HeaderTemplate>
								<asp:Button id="Button2" runat="server" CssClass="button" Text="Select All Users" OnCommand="commandBtnClick"
									CommandName="SelectUsers" CommandArgument="commandBtnClick"></asp:Button>
							</HeaderTemplate>
							<ItemTemplate>
								<asp:CheckBox id="Checkbox2" Runat="server"></asp:CheckBox>
							</ItemTemplate>
						</asp:TemplateColumn>
						<asp:BoundColumn Visible="False" DataField="seqUserID" SortExpression="seqUserID" HeaderText="seqUserID"></asp:BoundColumn>
						<asp:BoundColumn Visible="False" DataField="intUserType" SortExpression="intUserType" HeaderText="intUserType"></asp:BoundColumn>
						<asp:BoundColumn DataField="Name" SortExpression="Name" HeaderText="User Name"></asp:BoundColumn>
						<asp:BoundColumn DataField="strHID" SortExpression="strHID" HeaderText="Hanford ID"></asp:BoundColumn>
						<asp:BoundColumn DataField="strEmail" SortExpression="strEmail" HeaderText="Email Address"></asp:BoundColumn>
					</Columns>
				</asp:datagrid></td>
		</tr>
		<% else %>
		<tr align="center">
			<td width="20"></td>
			<td colSpan="2"><asp:label id="Label2" runat="server" CssClass="boldContent" BackColor="LightGrey" ForeColor="Red"
					Font-Bold="True">There are no users with a last name beginning with the letter specified.<br />If you just arrived on this page, then there are no persons in the system with a last name beginning with 'A'.</asp:label></td>
		</tr>
		<% end if %>
		<tr class="navPanel">
			<td class="navPanelLeft"></td>
			<td vAlign="middle" align="center" colSpan="3"><asp:imagebutton id="imgA2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/A.png" CssClass="btnSpace"
					ToolTip="List all last names starting with A" AlternateText="A" CausesValidation="False" OnCommand="commandGridUserList" CommandName="A"></asp:imagebutton><asp:imagebutton id="imgB2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/B.png" CssClass="btnSpace"
					ToolTip="List all last names starting with B" AlternateText="B" CausesValidation="False" OnCommand="commandGridUserList" CommandName="B"></asp:imagebutton><asp:imagebutton id="imgC2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/C.png" CssClass="btnSpace"
					ToolTip="List all last names starting with C" AlternateText="C" CausesValidation="False" OnCommand="commandGridUserList" CommandName="C"></asp:imagebutton><asp:imagebutton id="imgD2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/D.png" CssClass="btnSpace"
					ToolTip="List all last names starting with D" AlternateText="D" CausesValidation="False" OnCommand="commandGridUserList" CommandName="D"></asp:imagebutton><asp:imagebutton id="imgE2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/E.png" CssClass="btnSpace"
					ToolTip="List all last names starting with E" AlternateText="E" CausesValidation="False" OnCommand="commandGridUserList" CommandName="E"></asp:imagebutton><asp:imagebutton id="imgF2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/F.png" CssClass="btnSpace"
					ToolTip="List all last names starting with F" AlternateText="F" CausesValidation="False" OnCommand="commandGridUserList" CommandName="F"></asp:imagebutton><asp:imagebutton id="imgG2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/G.png" CssClass="btnSpace"
					ToolTip="List all last names starting with G" AlternateText="G" CausesValidation="False" OnCommand="commandGridUserList" CommandName="G"></asp:imagebutton><asp:imagebutton id="imgH2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/H.png" CssClass="btnSpace"
					ToolTip="List all last names starting with H" AlternateText="H" CausesValidation="False" OnCommand="commandGridUserList" CommandName="H"></asp:imagebutton><asp:imagebutton id="imgI2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/I.png" CssClass="btnSpace"
					ToolTip="List all last names starting with I" AlternateText="I" CausesValidation="False" OnCommand="commandGridUserList" CommandName="I"></asp:imagebutton><asp:imagebutton id="imgJ2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/J.png" CssClass="btnSpace"
					ToolTip="List all last names starting with J" AlternateText="J" CausesValidation="False" OnCommand="commandGridUserList" CommandName="J"></asp:imagebutton><asp:imagebutton id="imgK2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/K.png" CssClass="btnSpace"
					ToolTip="List all last names starting with K" AlternateText="K" CausesValidation="False" OnCommand="commandGridUserList" CommandName="K"></asp:imagebutton><asp:imagebutton id="imgL2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/L.png" CssClass="btnSpace"
					ToolTip="List all last names starting with L" AlternateText="L" CausesValidation="False" OnCommand="commandGridUserList" CommandName="L"></asp:imagebutton><asp:imagebutton id="imgM2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/M.png" CssClass="btnSpace"
					ToolTip="List all last names starting with M" AlternateText="M" CausesValidation="False" OnCommand="commandGridUserList" CommandName="M"></asp:imagebutton><asp:imagebutton id="imgN2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/N.png" CssClass="btnSpace"
					ToolTip="List all last names starting with N" AlternateText="N" CausesValidation="False" OnCommand="commandGridUserList" CommandName="N"></asp:imagebutton><asp:imagebutton id="imgO2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/O.png" CssClass="btnSpace"
					ToolTip="List all last names starting with O" AlternateText="O" CausesValidation="False" OnCommand="commandGridUserList" CommandName="O"></asp:imagebutton><asp:imagebutton id="imgP2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/P.png" CssClass="btnSpace"
					ToolTip="List all last names starting with P" AlternateText="P" CausesValidation="False" OnCommand="commandGridUserList" CommandName="P"></asp:imagebutton><asp:imagebutton id="imgQ2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/Q.png" CssClass="btnSpace"
					ToolTip="List all last names starting with Q" AlternateText="Q" CausesValidation="False" OnCommand="commandGridUserList" CommandName="Q"></asp:imagebutton><asp:imagebutton id="imgR2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/R.png" CssClass="btnSpace"
					ToolTip="List all last names starting with R" AlternateText="R" CausesValidation="False" OnCommand="commandGridUserList" CommandName="R"></asp:imagebutton><asp:imagebutton id="imgS2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/S.png" CssClass="btnSpace"
					ToolTip="List all last names starting with S" AlternateText="S" CausesValidation="False" OnCommand="commandGridUserList" CommandName="S"></asp:imagebutton><asp:imagebutton id="imgT2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/T.png" CssClass="btnSpace"
					ToolTip="List all last names starting with T" AlternateText="T" CausesValidation="False" OnCommand="commandGridUserList" CommandName="T"></asp:imagebutton><asp:imagebutton id="imgU2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/U.png" CssClass="btnSpace"
					ToolTip="List all last names starting with U" AlternateText="U" CausesValidation="False" OnCommand="commandGridUserList" CommandName="U"></asp:imagebutton><asp:imagebutton id="imgV2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/V.png" CssClass="btnSpace"
					ToolTip="List all last names starting with V" AlternateText="V" CausesValidation="False" OnCommand="commandGridUserList" CommandName="V"></asp:imagebutton><asp:imagebutton id="imgW2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/W.png" CssClass="btnSpace"
					ToolTip="List all last names starting with W" AlternateText="W" CausesValidation="False" OnCommand="commandGridUserList" CommandName="W"></asp:imagebutton><asp:imagebutton id="imgX2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/X.png" CssClass="btnSpace"
					ToolTip="List all last names starting with X" AlternateText="X" CausesValidation="False" OnCommand="commandGridUserList" CommandName="X"></asp:imagebutton><asp:imagebutton id="imgY2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/Y.png" CssClass="btnSpace"
					ToolTip="List all last names starting with Y" AlternateText="Y" CausesValidation="False" OnCommand="commandGridUserList" CommandName="Y"></asp:imagebutton><asp:imagebutton id="imgZ2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/Z.png" CssClass="btnSpace"
					ToolTip="List all last names starting with Z" AlternateText="Z" CausesValidation="False" OnCommand="commandGridUserList" CommandName="Z"></asp:imagebutton></td>
		</tr>
	</TBODY>
</table>
<%end if%>
<HR style="LEFT: 10px; BORDER-TOP-STYLE: ridge; TOP: 480px" width="100%" SIZE="3">
<TABLE style="LEFT: 10px; TOP: 536px" width="100%" align="center" class="borderless">
	<tr>
		<td>
		    <% If Session("PageOptions") <> 2 Then%>
		        <asp:button id="btnInsert2" Runat="server" CssClass="button" ToolTip="Insert" Text="Insert"></asp:button>
		    <% end if %>
		    <% if session("PageOptions") <> 1 then %>
		        <asp:button id="btnUpdate2" Runat="server" CssClass="button" ToolTip="Update" Text="Update"></asp:button>
		        <asp:button id="btnDelete2" Runat="server" CssClass="button" ToolTip="Delete" Text="Delete"></asp:button>
		    <% end if %>
		    <asp:button id="btnReset2" Runat="server" CssClass="button" ToolTip="Reset" Text="Reset"></asp:button>
		</td>
	</tr>
</TABLE>
</form>
</asp:Content>


