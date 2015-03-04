<%@ Register TagPrefix="igchartprop" Namespace="Infragistics.UltraChart.Resources.Appearance" Assembly="Infragistics2.WebUI.UltraWebChart.v8.2, Version=8.2.20082.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" %>
<%@ Register TagPrefix="igchartdata" Namespace="Infragistics.UltraChart.Data" Assembly="Infragistics2.WebUI.UltraWebChart.v8.2, Version=8.2.20082.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" %>
<%@ Register TagPrefix="igchart" Namespace="Infragistics.WebUI.UltraWebChart" Assembly="Infragistics2.WebUI.UltraWebChart.v8.2, Version=8.2.20082.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" %>
<%@ Page trace="false" Language="vb" AutoEventWireup="false" Codebehind="ReportOutput.aspx.vb" Inherits="SurveyAdmin.ReportOutput"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>ReportOutput</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<style type="text/css">@import url( ./styles.css ); 
		</style>
		<SCRIPT src="./pop_up.js" type="text/javascript"></SCRIPT>
	</HEAD>
	<body>
		<form id="Form1" method="post" runat="server">
			<table align="center" border="0">
				<tr>
					<td align="center" colSpan="2"><asp:label id="lblReportHeader" runat="server" CssClass="boldContentTeal"></asp:label>
						<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
						<asp:label id="lblSurveyHeader" runat="server" CssClass="sectionHeadTeal"></asp:label></td>
				</tr>
				<tr>
					<td class="ReportHeaderData">Total Respondents:</td>
					<td><asp:label id="lblTotalRespondents" runat="server" CssClass="ReportHeaderRed"></asp:label></td>
				</tr>
				<tr>
					<td class="ReportHeaderData">Filtered Date Range:</td>
					<td><asp:label id="lblDateRange" runat="server" CssClass="ReportHeaderBlue"></asp:label></td>
				</tr>
				<tr>
					<td class="ReportHeaderData">Period of Record:</td>
					<td><asp:label id="lblPeriodRecord" runat="server" CssClass="ReportHeaderBlue"></asp:label></td>
				</tr>
				<tr>
					<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
				</tr>
			</table>
			<% IF CType(Session("FiltersSelected"), boolean) %>
			<asp:table id="ReportFilter" runat="server" cellPadding="2" cellSpacing="0" HorizontalAlign="Center">
				<asp:TableRow>
					<asp:TableCell HorizontalAlign="center" ColumnSpan="3" cssclass="sectionHeadTeal">Current Filters
						</asp:TableCell>
				</asp:TableRow>
			</asp:table><asp:table id="ReportFilterResults" runat="server" cellPadding="2" cellSpacing="0" HorizontalAlign="Center"></asp:table>
			<% End if %>
			<asp:table id="ReportFilterFooter" Runat="server" HorizontalAlign="Center">
				<asp:TableRow>
					<asp:TableCell ColumnSpan="2" HorizontalAlign="Center">
						<asp:Label runat="server" ID="lblNote" CssClass="boldContentTeal">
						(Number of actual Data Points in Teal)<br />Note: For any question with "Check all
						that apply" the Data Points may not match the tabular results.</asp:Label>
					</asp:TableCell>
				</asp:TableRow>
				<asp:TableRow>
					<asp:TableCell Text="
						&lt;HR style=&quot;BORDER-TOP-STYLE: ridge&quot; width=&quot;100%&quot; SIZE=&quot;3&quot;&gt;
					"></asp:TableCell>
				</asp:TableRow>
			</asp:table>
			<% IF CType(Session("QuestionsSelected"), boolean) %>
			<asp:table id="DataTable" runat="server" cellPadding="0" cellSpacing="0" Width="100%" GridLines=Both BorderStyle=Solid BorderWidth="0px">
				<asp:TableRow Width="100%" CssClass="BreakBefore">
					<asp:TableCell HorizontalAlign="center" ColumnSpan="3" cssclass="sectionHeadTeal">Data Section
						<HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
					</asp:TableCell>
				</asp:TableRow>
			</asp:table>
			<% Else %>
			<table width="100%" border="0">
				<tr align="center">
					<td align="center"><asp:label id="lblNoData" runat="server" CssClass="ReportHeaderRed">There is no data for the filter(s), date range, and/or question(s) you have selected.</asp:label></td>
				</tr>
			</table>
			<% End if %>
			<asp:table id="DataTable2" runat="server" cellPadding="0" cellSpacing="0" Width="100%" GridLines=Both BorderStyle=Solid BorderWidth="0px"></asp:table>
		</form>
	</body>
</HTML>
