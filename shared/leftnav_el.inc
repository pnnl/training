<!--#include virtual="/_include/pnl_inc.asp"-->
<!--#include virtual="/_include/hanford_info.asp"-->
<div id="leftNav">
<h2>Quick Links</h2>
<ul>
<li><a href="/default.asp" id="elhome">Enterprise Learning Home</a></li>
<li><a href="http://hr.pnl.gov/index.stm">HR Home</a></li>
<li><a href="http://ods.pnl.gov">ODS Home</a></li>
</ul>
 	
	<h2>Listen and Learn</h2>
	<ul>
		<li><a href="default.asp" id="default">Listen and Learn Home</a></li>
		<li><a href="sos1802.asp" id="vid1802">Manager SOS Briefing</a></li>
		<li><a href="em14001.asp" id="vid14001">ISO-14001 Environmental Management</a></li>
		<li><a href="badges.asp" id="badges">Security Badges</a></li>
	</ul>
	
	<h2>Human Resources</h2>
	<ul>
		<li><a href="http://hr.pnl.gov">Home</a></li>
	</ul>
	<hr />
<%
sHits = HitCounter( (session(strScriptName) <> "true") )
session(strScriptName) = "true"
%>
</div>