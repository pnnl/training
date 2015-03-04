<%@ Page Language="vb" trace="false" MasterPageFile="~/MasterPage.master" AutoEventWireup="false" Codebehind="Password.aspx.vb" Inherits="SurveyAdmin.Password" smartNavigation="False" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder" Runat="Server">
<%=Session("JSAction")%>
<form id="Form1" method="post" runat="server">
    <h1>Change Password</h1>
    <p></p>
    <table cellSpacing="5" cellPadding="5" align="center" width="100%" class="borderless">
	    <% if (Session("isExpired") <> "False") then %>
	    <tr>
		    <td colspan="3"><asp:label id="Label1" Runat="server">Your password has expired.  You must change it now.</asp:label></td>
	    </tr>
	    <% end if %>
	    <tr align="center">
		    <td align="right" width="150px"><asp:image id="hlpPrevPassword" runat="server" ImageUrl="./images/help.png" ToolTip="Please enter your current password."></asp:image></td>
		    <td align="right" width="150px"><asp:label id="lblPrevPassword" Runat="server">Previous Password:</asp:label></td>
		    <td align="left"><asp:textbox id="txtPrevPassword" Runat="server" TextMode="Password" Width="150" MaxLength="32"></asp:textbox><asp:label id="lblErrPrevPassword" runat="server" cssClass="boldContentRed"></asp:label></td>
	    </tr>
	    <tr align="center">
		    <td align="right" width="150px"><asp:image id="hlpNewPassword" runat="server" ImageUrl="./images/help.png" ToolTip="Please enter the new password that you want to use."></asp:image></td>
		    <td align="right" width="150px"><asp:label id="lblNewPassword" Runat="server">New Password:</asp:label></td>
		    <td align="left"><asp:textbox id="txtNewPassword" Runat="server" TextMode="Password" Width="150" MaxLength="32"></asp:textbox><asp:label id="lblErrNewPassword" runat="server" cssClass="boldContentRed"></asp:label></td>
	    </tr>
	    <tr align="center">
		    <td align="right" width="150px"><asp:image id="hlpConfirmPassword" runat="server" ImageUrl="./images/help.png" ToolTip="Please confirm your new password."></asp:image></td>
		    <td align="right" width="150px"><asp:label id="lblConfirmPassword" Runat="server">Confirm Password:</asp:label></td>
		    <td align="left"><asp:textbox id="txtConfirmPassword" Runat="server" TextMode="Password" Width="150" MaxLength="32"></asp:textbox><asp:label id="lblErrConfirmPassword" runat="server" cssClass="boldContentRed"></asp:label></td>
	    </tr>
	    <tr align="center">
		    <td colSpan="3">&nbsp;<asp:button id="btnSubmit" Runat="server" ToolTip="Submit" CssClass="button" Text="Submit"></asp:button>
			    <asp:button id="btnCancel" Runat="server" ToolTip="Cancel" CssClass="button" Text="Cancel"></asp:button>
			    <asp:button id="btnReset" Runat="server" ToolTip="Reset" CssClass="button" Text="Reset"></asp:button></td>
	    </tr>
    </table>
    <table cellSpacing="5" cellPadding="5" align="center" class="borderless">
	    <tr align="center">
		    <td align="left" colSpan="3"><asp:label id="lblErrSummary" runat="server" cssClass="boldContentRed"></asp:label></td>
	    </tr>
    </table>
</form>
</asp:Content>
