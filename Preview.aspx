<%@ Register TagPrefix="igsch" Namespace="Infragistics.WebUI.WebSchedule" Assembly="Infragistics2.WebUI.WebDateChooser.v8.2, Version=8.2.20082.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" %>
<%@ Register TagPrefix="igtxt" Namespace="Infragistics.WebUI.WebDataInput" Assembly="Infragistics2.WebUI.WebDataInput.v8.2, Version=8.2.20082.1000, Culture=neutral, PublicKeyToken=7dd5c3163f2cd0cb" %>
<%@ Page trace="false" Language="vb" AutoEventWireup="false" Codebehind="Preview.aspx.vb" Inherits="SurveyAdmin.Preview" smartNavigation="True" %>
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
	    <style type="text/css">
	        #main {
	        font-size:100%;
            line-height:150%;
            margin:0px 0px 0 0px;
            padding-top:2px;
            }
            #container {
                float:left;
                width:100%;
            }
        </style>
		<SCRIPT src="./pop_up.js" type="text/javascript"></SCRIPT>
		<script type="text/javascript">
			var msg = null;
			var tipTimer;
			
			function clean_up() {
				if (msg != null) {
					if (msg.open) {
					msg.close();
					msg = null;
					}
				}
				window.clearTimeout(tipTimer);
			}

			function WinPopNew(strURL, intX, intY)
			{
				//used to destroy a popup so that the new popup will be the right size.
				//also checks to see if user is on a mac, to prevent clean_up which will break this. (JR_Sharp)
				var bAintMac = <%=lcase(cStr((InStr(lcase(Request.ServerVariables("ALL_HTTP")), "win") <> 0)))%>;
				var sServerName = "<%=lcase(Request.ServerVariables("SERVER_NAME"))%>";
				var bPopUp = true;
				var sTmpURL = strURL;
				if (bAintMac)
				{
					clean_up(); 
				}
				
				strOptions = 'menubar=0,directories=0,scrollbars=0,toolbar=0,status=0,resizable=1,';
				strOptions = strOptions + 'width=' + intX + ',height=' + intY + ',ontop=1';

				if (msg = open(sTmpURL, "frameWindow", strOptions))
				{
					msg.focus();
				}
				else
				{
					clean_up();
				}
				tipTimer=window.setTimeout("clean_up()", 10000);
			}
		</script>
	</HEAD>
	<%=Session("JSAction")%>
	<body class="col0" style="MARGIN: 0px">
	    <div id="page">
            <div id="container">
                <div id="main">
		            <form id="Form1" method="post" runat="server">
		                <h1 align="center">Survey Preview</h1>
			            <table width="100%" align="center" class="borderless">
				            <tr align="center">
					            <td align="center"><br>
						            <asp:label id="lblDisclaimer" runat="server" CssClass="boldContent" Font-Bold="True" ForeColor="Red"
							            BackColor="LightGrey"> Note: The colors on this page are the  default colors. You may
							            use the default colors or you may pass a side banner image and/or color(s) to the survey page loading your
							            survey.</asp:label></td>
				            </tr>
			            </table>
			            <p></p>
			            <HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
			            <table cellSpacing="0" width="100%" class="borderless">
				            <% if (Session("QuestionAvailability") <> "None") then %>
				            <tr style="WIDTH: 100%; HEIGHT: 100%">
					            <td noWrap width="139" bgcolor='<%=Session("BackgroundColor")%>'></td>
					            <td><asp:table id="QuestionTable" Runat="server" CssClass="borderless" Width="100%"></asp:table></td>
				            </tr>
				            <tr>
					            <td noWrap width="139" style="HEIGHT: 26px"></td>
					            <% if ((Session("isADCHSR") <> "True") or (Session("isADCHSR") = "True" and (CType(Session("isTemplateOwner"),boolean) or CType(Session("isTemplateDelegate"),boolean)))) then %>
					            <td align="center" style="HEIGHT: 26px"><asp:button id="btnRefresh" Runat="server" CssClass="button" Text="Refresh Page" ToolTip="Refresh Page"></asp:button></td>
					            <% end if %>
				            </tr>
				            <% else %>
				            <tr align="center">
					            <td colSpan="3">
						            <HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
						            <asp:label id="lblNoQuestions" runat="server" CssClass="boldContent" Font-Bold="True" ForeColor="Red"
							            BackColor="LightGrey">There are currently no questions in this template.  If you feel that this is
								            an error please contact the 
								            <a title="Sned mail to the Survey Tool Administrator" href="mailto:gene.gower@pnl.gov?subject=Template Question Error">
								            Survey Administrator</a>.</asp:label></td>
				            </tr>
				            <% end if %>
				            <% if (Session("isADCHSR") = "True" and not (CType(Session("isTemplateOwner"),boolean) or CType(Session("isTemplateDelegate"),boolean))) then %>
				            <TR>
					            <td noWrap width="139"></td>
					            <TD align="left">
						            <HR style="BORDER-TOP-STYLE: ridge" width="100%" SIZE="3">
						            <font class="content">Reviewer Input Section</font>
						            <table cellSpacing="0" cellPadding="2" border="1">
							            <tr>
								            <td vAlign="top"><asp:image id="hlpComments" runat="server" ImageUrl="./images/help.png" ToolTip="Enter any comments you may have about the survey template here whether you approve or not.  These comments will be included in the e-mail sent to the user in either case."></asp:image></td>
								            <td vAlign="top" align="right"><asp:label id="lblQuestion" Runat="server" CssClass="boldContent">Comments:</asp:label></td>
								            <td align="left"><asp:textbox id="txtComments" runat="server" CssClass="content" Width="400px" Columns="50" Rows="15"
										            TextMode="MultiLine" MaxLength="1800" Height="200px"></asp:textbox></td>
							            </tr>
							            <% if not(Session("hsrHID") is nothing) then %>
							            <tr>
								            <TD vAlign="top"><asp:image id="hlpHSR" runat="server" ImageUrl="./images/help.png" ToolTip="Check this box to indicate that this survey does not include any Human Subjects Research items."></asp:image></TD>
								            <TD vAlign="top" align="right"><asp:checkbox id="chkbxHSR" runat="server" CssClass="content"></asp:checkbox></TD>
								            <TD align="left"><asp:label id="lblRequired" Runat="server" CssClass="boldContent">Does this survey template contain any Human Subject Research items?</asp:label></TD>
							            </tr>
							            <% end if %>
						            </table>
					            </TD>
				            </TR>
				            <tr>
					            <td></td>
					            <td noWrap><asp:button id="btnAccept" Runat="server" CssClass="button" Text="Accept" ToolTip="Accept"></asp:button><asp:button id="btnReject" Runat="server" CssClass="button" Text="Reject" ToolTip="Reject"></asp:button></td>
				            </tr>
				            <% end if %>
			            </table>
		            </form>
                </div>
		    </div>
        </div>  
	</body>
</HTML>
