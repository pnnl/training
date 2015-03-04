<%@ Page trace="false" Language="vb" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" Codebehind="Login.aspx.vb" Inherits="SurveyAdmin.Login" smartNavigation="False" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder" Runat="Server">
<%=Session("JSAction")%>
<form id="Form1" method="post" runat="server">
<h1>Survey Administration</h1>
<p></p>
<div style="text-align:left;">
    <asp:label id="lblWelcomeText" Runat="server">This site is for administering a survey hosted on our machine. 
    You will only need to access <b>this</b> site if you are a Survey Administrator.  If you are a Survey 
    Respondent then you will receive an e-mail that allows you to take the survey.  The Survey Administration 
    website requires you to login with the username and password that was sent to your e-mail address. If 
    you have not yet received an email from the survey project manager, and feel that you should 
    have, please contact the </asp:label><asp:label id="lblPNNLText" runat="server" ToolTip="Pacific Northwest National Laboratory"
    CssClass="info">PNNL</asp:label><asp:label id="lblMailText" runat="server">
    <a title="Send mail to the Survey Administration Tool Manager." href="mailto:gene.gower@pnl.gov?subject=Survey Administration Access">Survey 
    Administration Tool Manager</a>.</asp:label>
</div>
<br />
<div style="text-align:center;border-style:ridge;border-width:medium;">
    <asp:Label ID="lblDisclaimer" Runat="server" CssClass="ReportHeaderRed">Disclaimer:</asp:Label><br>
	<asp:label id="lblDisclaimerText" Runat="server">
	The Training and Qualification Department has made this tool available for use within the Laboratory environment, 
	T&amp;Q does not control or maintain the content of the templates and surveys generated by users of this site. 
	</asp:label>
</div>
<br />
<div style="text-align:left;">
    <asp:label id="lblNoteText" Runat="server" CssClass="noteText"><p class="noteText"><span class="noteHeader">Note: </span>If one or more of the email relay servers has mangled 
	the link within your email then you will end up here inadvertently.  The link in your email should 
	read something like: 
	<u>http://training.pnl.gov/Default.aspx?seqUserID=nnn&amp;strAuthenticator=XXXX&amp;seqSurvey=mm</u> 
	The link may not be correctly formed by word-wrapping, thus landing you here.  To correct the problem 
	type the link from your email message into the <b>Address</b> or <b>Location</b> line of your 
	browser without spaces or carriage returns.</p></asp:label>
</div>
<br />
<div style="text-align:center;">
    <asp:label id="lblLoginText" Runat="server" CssClass="boldContent">- Please Login -</asp:label>
</div>
<br />
<div style="text-align:center;">
<asp:label id="lblUserName" Runat="server" CssClass="boldContent">User Name:</asp:label>
<asp:textbox id="txtUserName" Runat="server" Width="150" MaxLength="128" tabIndex="1"></asp:textbox>
<br /><br /><asp:label id="lblPassword" Runat="server" CssClass="boldContent">&nbsp;&nbsp;Password:</asp:label>
<asp:textbox id="txtPassword" tabIndex="2" Runat="server" Width="150" MaxLength="32" TextMode="Password"></asp:textbox>
<br /><br /><asp:button id="btnSubmit" tabIndex="3" Runat="server" Text="Submit" CssClass="button" ToolTip="Submit"></asp:button>&nbsp;
<asp:button id="btnReset" tabIndex="4" Runat="server" Text="Reset" CssClass="button" ToolTip="Reset"></asp:button>
</div>
</form>
</asp:Content>