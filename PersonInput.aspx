<%@ Page trace="false" Language="vb" AutoEventWireup="false" Codebehind="PersonInput.aspx.vb" Inherits="SurveyAdmin.PersonInput"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
	<HEAD>
	    <!--#include virtual="/shared/global.inc"-->
		<title>PersonInput</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<style type="text/css">@import url( ./styles.css ); 
		</style>
		<SCRIPT src="./pop_up.js" type="text/javascript"></SCRIPT>
	</HEAD>
	<style type="text/css">
	    #main {
	    font-size:100%;
        line-height:150%;
        margin:0px 0px 0 0px;
        padding-top:2px;
        }
        #container {
            float:left;
            width:50%;
        }
    </style>
	<body class="col0" style="MARGIN: 0px">
	    <div id="page">
            <div id="container">
                <div id="main">
		            <form id="Form1" method="post" runat="server">
		                <h1>Person Input</h1>
			            <table style="BORDER-COLLAPSE: collapse; HEIGHT: 100%" cellSpacing="0" cellPadding="0"
				            class="borderless">
				            <tr style="HEIGHT: 100%">
					            <td id="Td1" style="PADDING-RIGHT: 5px; PADDING-LEFT: 5px; WIDTH: 100%; PADDING-TOP: 5px"
						            vAlign="top" align="left" colSpan="1" rowSpan="2" runat="server">
						            <table width="100%" class="borderless">
							            <tr>
								            <td><asp:label id="lblIntroduction" Runat="server" CssClass="boldContent"></asp:label></td>
							            </tr>
						            </table>
						            <p></p>
						            <table cellSpacing="5" cellPadding="5" width="100%" align="center" class="borderless">
							            <tr>
								            <td><asp:image id="hlpFirstName" runat="server" ImageUrl="./images/help.png" ToolTip="Please enter the last name of the individual."></asp:image></td>
								            <td align="right"><asp:label id="lblPrevPassword" Runat="server" cssclass="content">Last Name:</asp:label></td>
								            <td align="left"><asp:textbox id="txtLastName" Runat="server" MaxLength="60" Width="150"></asp:textbox><asp:requiredfieldvalidator id="RequiredFieldValidator1" runat="server" ErrorMessage="RequiredFieldValidator"
										            ControlToValidate="txtLastName">You must enter a last name.</asp:requiredfieldvalidator></td>
							            </tr>
							            <tr>
								            <td><asp:image id="hlpLastName" runat="server" ImageUrl="./images/help.png" ToolTip="Please enter the first name of the individual."></asp:image></td>
								            <td align="right"><asp:label id="lblNewPassword" Runat="server" cssclass="content">First Name:</asp:label></td>
								            <td align="left"><asp:textbox id="txtFirstName" Runat="server" MaxLength="60" Width="150"></asp:textbox><asp:requiredfieldvalidator id="RequiredFieldValidator2" runat="server" ErrorMessage="RequiredFieldValidator"
										            ControlToValidate="txtFirstName">You must enter a first name.</asp:requiredfieldvalidator></td>
							            </tr>
							            <tr>
								            <td><asp:image id="hlpMiddleInitial" runat="server" ImageUrl="./images/help.png" ToolTip="Please enter the middle initial of the individual.  (Note: This field is entirely optional)"></asp:image></td>
								            <td align="right"><asp:label id="lblConfirmPassword" Runat="server" cssclass="content">Middle Initial:</asp:label></td>
								            <td align="left"><asp:textbox id="txtMiddleInitial" Runat="server" MaxLength="60" Width="150"></asp:textbox></td>
							            </tr>
							            <TR>
								            <TD align="center" width="50%" colSpan="3"><asp:button id="btnCancel" runat="server" CssClass="button" Text="Cancel" CausesValidation="False"
										            ToolTip="Cancel"></asp:button>&nbsp;
									            <asp:button id="btnSubmit" runat="server" CssClass="button" Text="Submit" ToolTip="Submit"></asp:button></TD>
							            </TR>
						            </table>
						            <INPUT id="Email" type="hidden" runat="server"><INPUT id="UserType" type="hidden" runat="server">
						            <input id="seqSurvey" type="hidden" runat="server"> <input id="intGroupID" type="hidden" runat="server">
					            </td>
				            </tr>
			            </table>
		            </form>
                </div>
		    </div>
        </div>    
	</body>
</HTML>
