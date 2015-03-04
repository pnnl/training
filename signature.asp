<%
Response.Expires = 0
Response.Buffer = True 
Response.AddHeader "Pragma", "no-cache"
%>    
<%
'
' Load the QueryString into Session Variables
'
for each item in request.querystring
  session(item) = trim(""&request(item))
next
for each item in request.form
  session(item) = trim(""&request(item))
next

str_Path = lcase(Request.ServerVariables("Path_Info"))
str_Path = left(str_Path, inStrRev(str_Path,"/"))

%>
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Transfer to Survey Administration Digital Signature</title>
<link rel="stylesheet" type="text/css" href="Styles.css"></link>
<SCRIPT src="./pop_up.js" type="text/javascript"></SCRIPT>
<script><!--
function SubmitForm() {
    if (fObj=findObj('myForm')) {
      fObj.submit();
    } else {
      document.location.href = '<%= str_Path & session("return") %>';
    }
}
//--></script>
</head>

<body>
<div align="center">
<h2><b>Transfer to Survey Administration Digital Signature</b></h2>
<p></p>
<table width="80%" height="100%"><tr><td>Our system uses the 
  <a target="_blank" href="http://infosource.pnl.gov/Network/services/digitalid/default.htm">PNNL Digital ID</a> 
  to electronically sign that you agree that this survey template is acceptable.&nbsp; 
  This process transfers you <b>away from our server</b> for the next step.&nbsp; 
  Please follow the link to accept this survey template<p align="center"><a href="JavaScript:SubmitForm()">Sign Form Electronically</a></p>
    <p align="left">If you are unable to complete the form electronically for 
    any reason, please contact the survey administrator.</p>
  <div align="center">
    <table border="0" width="80%" id="table1" bgcolor="#FFFFE1" cellpadding="4">
      <tr>
        <td valign="top"><b>Note</b>:</td>
        <td valign="top">It is normal for the Java program that accepts the 
        Digital ID PIN to clear the field and pause for up to 8 seconds while the validation routine processes.&nbsp; Please be 
        patient.</td>
      </tr>
    </table>
  </div>
<form method="POST" name="myForm" action="http://<%=Request.Servervariables("server_name")%>/signature/default.asp">
&nbsp;
<input type="hidden" name="HID" value="<%=session("HID")%>">
<input type="hidden" name="Return" value="<%= str_Path & session("return") %>">
<input type="hidden" name="str_Data" value="<%=session("str_Data") %>">
<input type="hidden" name="str_Title" value="<%=session("str_Title") %>">
<input type="hidden" name="str_AcknowledgmentDS" value="<%=replace(session("strAcknowledgmentDS"),"'","`")%>">
</form>

</td></tr></table>
</div>
</body>
</html>