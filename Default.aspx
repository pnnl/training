<%@ Page Language="vb" Trace="False" AutoEventWireup="false" Codebehind="Default.aspx.vb" Inherits="SurveyAdmin.Main" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Survey Administration Tool</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<style type="text/css">@import url( ./styles.css ); 
		</style>
		<SCRIPT src="./pop_up.js" type="text/javascript"></SCRIPT>
	</HEAD>
	<body style="MARGIN: 0px">
		<form id="Form1" method="post" runat="server">
			<table style="BORDER-COLLAPSE: collapse; HEIGHT: 100%" cellSpacing="0" cellPadding="0"
				border="0">
				<tr style="WIDTH: 178px; HEIGHT: 131px">
					<td vAlign="top"><IMG style="WIDTH: 178px" alt="Training and Qualifications" src="./images/Shared/tqlogo.png"
							border="0"></td>
					<td vAlign="top">
						<table style="BACKGROUND-IMAGE: url(./images/Shared/BTile.png); WIDTH: 100%; BACKGROUND-REPEAT: repeat-x; BORDER-COLLAPSE: collapse"
							cellSpacing="0" cellPadding="0" border="0">
							<tr>
								<td vAlign="top"><IMG alt="Training and Qualifications" src="./images/Shared/banner.jpg" width="555" align="top"
										border="0">
									<% if Session("isAuthenticated") <> "True" %>
									<DIV id="div1" style="Z-INDEX: 1; LEFT: 185px; WIDTH: 100%; POSITION: absolute; TOP: 85px">
										<script type="text/javascript">write_top_links_disabled()</script>
									</DIV>
									<DIV id="div2" style="Z-INDEX: 1; LEFT: 185px; WIDTH: 100%; POSITION: absolute; TOP: 110px">
										<script type="text/javascript">write_bottom_links_disabled()</script>
									</DIV>
									<% else %>
									<DIV id="div3" style="Z-INDEX: 1; LEFT: 185px; WIDTH: 100%; POSITION: absolute; TOP: 85px">
										<script type="text/javascript">write_top_links(<%=Session("UserType")%>)</script>
									</DIV>
									<DIV id="div4" style="Z-INDEX: 1; LEFT: 185px; WIDTH: 100%; POSITION: absolute; TOP: 110px">
										<script type="text/javascript">write_bottom_links(<%=Session("UserType")%>,'<%=Session("Machine")%>')</script>
									</DIV>
									<% end if %>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr style="HEIGHT: 100%">
					<td style="BACKGROUND-IMAGE: url(./images/Shared/menu_top.png); WIDTH: 178px; BACKGROUND-REPEAT: repeat-y; HEIGHT: 100%"
						vAlign="top" colSpan="1">
						<table style="WIDTH: 178px; HEIGHT: 75px" cellSpacing="8" cellPadding="8" border="0">
							<tr style="WIDTH: 178px">
								<td></td>
							</tr>
						</table>
						<table style="BACKGROUND-IMAGE: url(./images/Shared/Tile.png); WIDTH: 178px; BORDER-COLLAPSE: collapse; HEIGHT: 100%"
							cellSpacing="0" cellPadding="0" border="0">
							<tr style="HEIGHT: 100%">
								<td style="BACKGROUND-IMAGE: url(./images/Shared/button_background2.png); WIDTH: 178px; BACKGROUND-REPEAT: no-repeat"
									vAlign="top" align="center"><br>
									<script type="text/javascript">write_side_menu_links()</script>
								</td>
							</tr>
						</table>
					</td>
					<td style="PADDING-RIGHT: 5px; PADDING-LEFT: 5px; WIDTH: 100%; PADDING-TOP: 5px"
						vAlign="top" align="left" colSpan="1" rowSpan="2" runat="server"></td>
				</tr>
			</table>
		</form>
	</body>
</HTML>
