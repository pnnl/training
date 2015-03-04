<%@ Page trace="false" Language="vb" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" Codebehind="Permission.aspx.vb" Inherits="SurveyAdmin.Permission" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder" Runat="Server">
<%=Session("JSAction")%>
<form id="Form1" method="post" runat="server">
<h1>Template Permissions</h1>
<p></p>
<table cellSpacing="0" cellPadding="2" class="borderless">
	<tr>
		<td align="left"><asp:label id="lblTemplate" Runat="server" CssClass="boldContent">Template:</asp:label></td>
		<td noWrap align="left" class="content"><%=Session("TemplateName")%></td>
	</tr>
	<tr class="navPanel">
		<td></td>
		<td noWrap align="right"></td>
		<td style="WIDTH: 305px"></td>
		<td style="WIDTH: 100%" align="right"><asp:button id="btnTop" runat="server" CssClass="small" ForeColor="Blue" Font-Underline="True"
				BorderStyle="None" BackColor="Transparent" Text="^ Top ^" ToolTip="Back to Top"></asp:button></td>
	</tr>
</table>
<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
<table width="100%" class="borderless">
	<tr>
		<td>
			<table cellSpacing="5" cellPadding="5" width="100%" align="center" class="borderless">
				<% if (Session("TemplateUserListRows") <> "None") then %>
				<TBODY>
					<TR align="center">
						<TD vAlign="top" colSpan="3"><asp:label id="lblInGroup" runat="server" CssClass="boldContent" ForeColor="Red" Width="100%">Current Template Users</asp:label><br>
							<font class="boldContentRed">
								<%=Session("TemplateName")%>
							</font>
						</TD>
					</TR>
					<tr>
						<td vAlign="top" width="20"><asp:image id="hlpLevel0" runat="server" ImageUrl="./images/help.png" ToolTip="Uncheck users to remove their status as a user of this template."></asp:image></td>
						<td vAlign="top" align="left" colSpan="2"><asp:datagrid id="dgTemplateUsers" runat="server" Width="100%" CellPadding="0" CellSpacing="1"
								AutoGenerateColumns="False">
								<AlternatingItemStyle BackColor="#A7DCD2"></AlternatingItemStyle>
								<ItemStyle CssClass="content"></ItemStyle>
								<HeaderStyle HorizontalAlign="Left" BorderWidth="5px" BorderStyle="Outset" CssClass="content"
									BackColor="#60C0B4"></HeaderStyle>
								<Columns>
									<asp:TemplateColumn HeaderText="&lt;a href=&quot;./Permission.aspx?selectAll=1&quot;&gt;Select All&lt;/a&gt;">
										<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
										<ItemStyle HorizontalAlign="Center"></ItemStyle>
										<HeaderTemplate>
											<asp:Button id="Button1" runat="server" CssClass="button" Text="Select All Users" ToolTip="Select All Users"
												OnCommand="commandBtnClick" CommandName="SelectGroup" CommandArgument="commandBtnClick"></asp:Button>
										</HeaderTemplate>
										<ItemTemplate>
											<asp:CheckBox id="Checkbox1" Runat="server"></asp:CheckBox>
										</ItemTemplate>
									</asp:TemplateColumn>
									<asp:BoundColumn Visible="False" DataField="USER_KEY" SortExpression="USER_KEY" HeaderText="USER_KEY"></asp:BoundColumn>
									<asp:BoundColumn Visible="False" DataField="USER_CODE" SortExpression="USER_CODE" HeaderText="USER_CODE"></asp:BoundColumn>
									<asp:BoundColumn DataField="Name" SortExpression="Name" HeaderText="User Name"></asp:BoundColumn>
									<asp:BoundColumn DataField="HANFORD_ID" SortExpression="HANFORD_ID" HeaderText="Hanford ID"></asp:BoundColumn>
									<asp:BoundColumn DataField="EMAIL_ADDRESS" SortExpression="EMAIL_ADDRESS" HeaderText="Email Address"></asp:BoundColumn>
								</Columns>
							</asp:datagrid></td>
					</tr>
					<% else %>
					<tr align="center">
						<td colSpan="3">
							<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
							<asp:label id="lblLevel0None" runat="server" CssClass="boldContent" ForeColor="Red" BackColor="LightGrey"
								Font-Bold="True">There are currently no users of this Template.</asp:label></td>
					</tr>
					<% end if %>
				</TBODY>
			</table>
			<table cellSpacing="5" cellPadding="5" width="100%" align="center" class="borderless">
				<TBODY>
					<TR align="center">
						<TD vAlign="top" colSpan="3">
							<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
							<asp:label id="Label1" runat="server" CssClass="boldContent" ForeColor="Red" Width="100%">All available Users</asp:label></TD>
					</TR>
					<tr class="navPanel">
						<td class="navPanelLeft">
							<asp:image id="hlpAlphabet" runat="server" ImageUrl="./images/help.png" ToolTip="Select an alphabet character to list users whose last names begin with that character."></asp:image></td>
						<td vAlign="middle" align="center" colSpan="3"><asp:imagebutton id="imgA1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/A.png" CssClass="btnSpace"
								ToolTip="List all last names starting with A" OnCommand="commandGridUserList" CommandName="A" AlternateText="A" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgB1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/B.png" CssClass="btnSpace"
								ToolTip="List all last names starting with B" OnCommand="commandGridUserList" CommandName="B" AlternateText="B" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgC1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/C.png" CssClass="btnSpace"
								ToolTip="List all last names starting with C" OnCommand="commandGridUserList" CommandName="C" AlternateText="C" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgD1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/D.png" CssClass="btnSpace"
								ToolTip="List all last names starting with D" OnCommand="commandGridUserList" CommandName="D" AlternateText="D" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgE1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/E.png" CssClass="btnSpace"
								ToolTip="List all last names starting with E" OnCommand="commandGridUserList" CommandName="E" AlternateText="E" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgF1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/F.png" CssClass="btnSpace"
								ToolTip="List all last names starting with F" OnCommand="commandGridUserList" CommandName="F" AlternateText="F" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgG1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/G.png" CssClass="btnSpace"
								ToolTip="List all last names starting with G" OnCommand="commandGridUserList" CommandName="G" AlternateText="G" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgH1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/H.png" CssClass="btnSpace"
								ToolTip="List all last names starting with H" OnCommand="commandGridUserList" CommandName="H" AlternateText="H" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgI1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/I.png" CssClass="btnSpace"
								ToolTip="List all last names starting with I" OnCommand="commandGridUserList" CommandName="I" AlternateText="I" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgJ1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/J.png" CssClass="btnSpace"
								ToolTip="List all last names starting with J" OnCommand="commandGridUserList" CommandName="J" AlternateText="J" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgK1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/K.png" CssClass="btnSpace"
								ToolTip="List all last names starting with K" OnCommand="commandGridUserList" CommandName="K" AlternateText="K" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgL1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/L.png" CssClass="btnSpace"
								ToolTip="List all last names starting with L" OnCommand="commandGridUserList" CommandName="L" AlternateText="L" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgM1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/M.png" CssClass="btnSpace"
								ToolTip="List all last names starting with M" OnCommand="commandGridUserList" CommandName="M" AlternateText="M" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgN1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/N.png" CssClass="btnSpace"
								ToolTip="List all last names starting with N" OnCommand="commandGridUserList" CommandName="N" AlternateText="N" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgO1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/O.png" CssClass="btnSpace"
								ToolTip="List all last names starting with O" OnCommand="commandGridUserList" CommandName="O" AlternateText="O" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgP1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/P.png" CssClass="btnSpace"
								ToolTip="List all last names starting with P" OnCommand="commandGridUserList" CommandName="P" AlternateText="P" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgQ1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/Q.png" CssClass="btnSpace"
								ToolTip="List all last names starting with Q" OnCommand="commandGridUserList" CommandName="Q" AlternateText="Q" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgR1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/R.png" CssClass="btnSpace"
								ToolTip="List all last names starting with R" OnCommand="commandGridUserList" CommandName="R" AlternateText="R" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgS1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/S.png" CssClass="btnSpace"
								ToolTip="List all last names starting with S" OnCommand="commandGridUserList" CommandName="S" AlternateText="S" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgT1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/T.png" CssClass="btnSpace"
								ToolTip="List all last names starting with T" OnCommand="commandGridUserList" CommandName="T" AlternateText="T" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgU1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/U.png" CssClass="btnSpace"
								ToolTip="List all last names starting with U" OnCommand="commandGridUserList" CommandName="U" AlternateText="U" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgV1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/V.png" CssClass="btnSpace"
								ToolTip="List all last names starting with V" OnCommand="commandGridUserList" CommandName="V" AlternateText="V" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgW1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/W.png" CssClass="btnSpace"
								ToolTip="List all last names starting with W" OnCommand="commandGridUserList" CommandName="W" AlternateText="W" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgX1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/X.png" CssClass="btnSpace"
								ToolTip="List all last names starting with X" OnCommand="commandGridUserList" CommandName="X" AlternateText="X" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgY1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/Y.png" CssClass="btnSpace"
								ToolTip="List all last names starting with Y" OnCommand="commandGridUserList" CommandName="Y" AlternateText="Y" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgZ1" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/Z.png" CssClass="btnSpace"
								ToolTip="List all last names starting with Z" OnCommand="commandGridUserList" CommandName="Z" AlternateText="Z" CausesValidation="False"></asp:imagebutton></td>
					</tr>
					<% if (Session("GeneralUserListRows") <> "None") then %>
					<tr>
						<td vAlign="top" width="20"><asp:image id="Image2" runat="server" ImageUrl="./images/help.png" ToolTip="Select users to make them users of this template."></asp:image></td>
						<td vAlign="top" align="left" colSpan="2"><asp:datagrid id="dgUsers" runat="server" Width="100%" CellPadding="0" CellSpacing="1" AutoGenerateColumns="False">
								<AlternatingItemStyle BackColor="#A7DCD2"></AlternatingItemStyle>
								<ItemStyle CssClass="content"></ItemStyle>
								<HeaderStyle HorizontalAlign="Left" BorderWidth="5px" BorderStyle="Outset" CssClass="content"
									BackColor="#60C0B4"></HeaderStyle>
								<Columns>
									<asp:TemplateColumn HeaderText="&lt;a href=&quot;./Permission.aspx?selectAll=1&quot;&gt;Select All&lt;/a&gt;">
										<HeaderStyle HorizontalAlign="Center"></HeaderStyle>
										<ItemStyle HorizontalAlign="Center"></ItemStyle>
										<HeaderTemplate>
											<asp:Button id="Button2" runat="server" CssClass="button" Text="Select All Users" ToolTip="Select All Users"
												OnCommand="commandBtnClick" CommandName="SelectUsers" CommandArgument="commandBtnClick"></asp:Button>
										</HeaderTemplate>
										<ItemTemplate>
											<asp:CheckBox id="Checkbox2" Runat="server"></asp:CheckBox>
										</ItemTemplate>
									</asp:TemplateColumn>
									<asp:BoundColumn Visible="False" DataField="USER_KEY" SortExpression="USER_KEY" HeaderText="USER_KEY"></asp:BoundColumn>
									<asp:BoundColumn Visible="False" DataField="USER_CODE" SortExpression="USER_CODE" HeaderText="USER_CODE"></asp:BoundColumn>
									<asp:BoundColumn DataField="Name" SortExpression="Name" HeaderText="User Name"></asp:BoundColumn>
									<asp:BoundColumn DataField="HANFORD_ID" SortExpression="HANFORD_ID" HeaderText="Hanford ID"></asp:BoundColumn>
									<asp:BoundColumn DataField="EMAIL_ADDRESS" SortExpression="EMAIL_ADDRESS" HeaderText="Email Address"></asp:BoundColumn>
								</Columns>
							</asp:datagrid></td>
					</tr>
					<% else %>
					<tr>
						<td vAlign="top" width="20"></td>
						<td vAlign="top" align="center" colSpan="2">
							<asp:label id="Label2" runat="server" CssClass="boldContent" ForeColor="Red" BackColor="LightGrey"
								Font-Bold="True">There are no users with a last name beginning with the letter specified.<br />If you just arrived on this page, then there are no persons in the system with a last name beginning with 'A'.</asp:label></td>
					</tr>
					<% end if %>
					<tr class="navPanel">
						<td class="navPanelLeft">
							<asp:image id="hlpAlphabet2" runat="server" ImageUrl="./images/help.png" ToolTip="Select an alphabet character to list users whose last names begin with that character."></asp:image></td>
						<td vAlign="middle" align="center" colSpan="3"><asp:imagebutton id="imgA2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/A.png" CssClass="btnSpace"
								ToolTip="List all last names starting with A" OnCommand="commandGridUserList" CommandName="A" AlternateText="A" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgB2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/B.png" CssClass="btnSpace"
								ToolTip="List all last names starting with B" OnCommand="commandGridUserList" CommandName="B" AlternateText="B" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgC2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/C.png" CssClass="btnSpace"
								ToolTip="List all last names starting with C" OnCommand="commandGridUserList" CommandName="C" AlternateText="C" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgD2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/D.png" CssClass="btnSpace"
								ToolTip="List all last names starting with D" OnCommand="commandGridUserList" CommandName="D" AlternateText="D" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgE2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/E.png" CssClass="btnSpace"
								ToolTip="List all last names starting with E" OnCommand="commandGridUserList" CommandName="E" AlternateText="E" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgF2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/F.png" CssClass="btnSpace"
								ToolTip="List all last names starting with F" OnCommand="commandGridUserList" CommandName="F" AlternateText="F" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgG2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/G.png" CssClass="btnSpace"
								ToolTip="List all last names starting with G" OnCommand="commandGridUserList" CommandName="G" AlternateText="G" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgH2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/H.png" CssClass="btnSpace"
								ToolTip="List all last names starting with H" OnCommand="commandGridUserList" CommandName="H" AlternateText="H" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgI2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/I.png" CssClass="btnSpace"
								ToolTip="List all last names starting with I" OnCommand="commandGridUserList" CommandName="I" AlternateText="I" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgJ2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/J.png" CssClass="btnSpace"
								ToolTip="List all last names starting with J" OnCommand="commandGridUserList" CommandName="J" AlternateText="J" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgK2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/K.png" CssClass="btnSpace"
								ToolTip="List all last names starting with K" OnCommand="commandGridUserList" CommandName="K" AlternateText="K" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgL2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/L.png" CssClass="btnSpace"
								ToolTip="List all last names starting with L" OnCommand="commandGridUserList" CommandName="L" AlternateText="L" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgM2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/M.png" CssClass="btnSpace"
								ToolTip="List all last names starting with M" OnCommand="commandGridUserList" CommandName="M" AlternateText="M" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgN2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/N.png" CssClass="btnSpace"
								ToolTip="List all last names starting with N" OnCommand="commandGridUserList" CommandName="N" AlternateText="N" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgO2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/O.png" CssClass="btnSpace"
								ToolTip="List all last names starting with O" OnCommand="commandGridUserList" CommandName="O" AlternateText="O" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgP2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/P.png" CssClass="btnSpace"
								ToolTip="List all last names starting with P" OnCommand="commandGridUserList" CommandName="P" AlternateText="P" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgQ2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/Q.png" CssClass="btnSpace"
								ToolTip="List all last names starting with Q" OnCommand="commandGridUserList" CommandName="Q" AlternateText="Q" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgR2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/R.png" CssClass="btnSpace"
								ToolTip="List all last names starting with R" OnCommand="commandGridUserList" CommandName="R" AlternateText="R" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgS2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/S.png" CssClass="btnSpace"
								ToolTip="List all last names starting with S" OnCommand="commandGridUserList" CommandName="S" AlternateText="S" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgT2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/T.png" CssClass="btnSpace"
								ToolTip="List all last names starting with T" OnCommand="commandGridUserList" CommandName="T" AlternateText="T" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgU2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/U.png" CssClass="btnSpace"
								ToolTip="List all last names starting with U" OnCommand="commandGridUserList" CommandName="U" AlternateText="U" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgV2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/V.png" CssClass="btnSpace"
								ToolTip="List all last names starting with V" OnCommand="commandGridUserList" CommandName="V" AlternateText="V" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgW2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/W.png" CssClass="btnSpace"
								ToolTip="List all last names starting with W" OnCommand="commandGridUserList" CommandName="W" AlternateText="W" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgX2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/X.png" CssClass="btnSpace"
								ToolTip="List all last names starting with X" OnCommand="commandGridUserList" CommandName="X" AlternateText="X" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgY2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/Y.png" CssClass="btnSpace"
								ToolTip="List all last names starting with Y" OnCommand="commandGridUserList" CommandName="Y" AlternateText="Y" CausesValidation="False"></asp:imagebutton><asp:imagebutton id="imgZ2" runat="server" ImageUrl="./MyInfragisticsImages/Alphabet/Z.png" CssClass="btnSpace"
								ToolTip="List all last names starting with Z" OnCommand="commandGridUserList" CommandName="Z" AlternateText="Z" CausesValidation="False"></asp:imagebutton></td>
					</tr>
				</TBODY>
			</table>
		</td>
	</tr>
	<tr>
		<td>
		    <asp:button id="btnUpdate" Runat="server" CssClass="button" ToolTip="Update" Text="Update"></asp:button>
		    <asp:button id="btnReset" Runat="server" CssClass="button" ToolTip="Reset" Text="Reset"></asp:button>
		</td>
	</tr>
</table>
</form>
</asp:Content>

