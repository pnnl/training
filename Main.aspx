<%@ Page Trace="false" Language="vb" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" Codebehind="Main.aspx.vb" Inherits="SurveyAdmin.Main1" smartNavigation="False" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder" Runat="Server">
<%=Session("JSAction")%>
<h1>Welcome to the On-Line Survey Tool</h1>
<p></p>
<asp:label id="lblIntroduction" Runat="server"><span style="COLOR: red">INTRODUCTION:</span> This  
tool will allow you to assess the level of satisfaction of your customers.  The survey tool is designed to 
allow you to send the same survey questions on two or more occasions, and then compare the results in an 
output report format.  By using this tool you will be able to:</asp:label><br /><br />
	<ul>
		<li>
		Quantify and take credit for improvements that result in increased customer 
		satisfaction.
		<li>
		Identify and characterize problem areas of an existing or new service so that 
		solutions can be developed
		<li>
			Decrease customer complaints by acting upon information gained in the survey 
			responses.
		</li>
	</ul>
	<br>
	<asp:label id="lblIntroduction2" Runat="server">The survey process is as 
simple as sending email messages to your customers.  The email will contain a message from you and a URL link 
to a web site.  Customers then complete the simple on-line survey and results will be accessible to you on the 
reports page.  Detailed instructions are in the <i>Help</i> section of this site.</asp:label>
<p></p>
<asp:label id="lblNoteText" Runat="server"><font style="COLOR: red">TECHNICAL 
SUPPORT:</font>  
For live technical support, please send an email to the  
<a title="Send mail to the Survey Administration Tool Manager." href="mailto:gene.gower@pnl.gov?subject=Survey Administration Technical Support">Survey Administration Tool Manager</a>. 
Please include your phone and email information along with the best times to reach you and you will be contacted 
as quickly as possible.</asp:label>
<p></p>
<div id="lblMessage" runat="server"></div>
</asp:Content>
