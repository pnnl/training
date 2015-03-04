<div id="title">Pacific Northwest National Laboratory: Enterprise Learning Website</div>
<div id="bannerBack">
<div id="banner">
	<div id="logoBanner"><img src="/shared/design-images/tqlogo.jpg" width="178" height="195" alt="Enterprise Learning" usemap="#skipit" ></div>
    <map name="skipit" id="skipit">
    <area shape="rect" coords="0,0,1,1" href="#skipNav" alt="Skip Repetitive Navigation" >
     <area shape="rect" coords="1,1,178,195" href="/default.asp" alt="Enterprise Learning Home" >
    </map>
<div id="mainBanner"><img src="/shared/design-images/banner.jpg" width="552" height="136" alt="Enterprise Learning" ></div> 
</div>  
</div>
<div id="buttonBackground">
	<div id="mainNav">
		<div class="navItem">
<form name="gs" method="get" action="http://search.pnl.gov/search" id="searchTraining">
<input type="hidden" name="as_sitesearch" value="training.pnl.gov" />
<input type="text" size="16" id="searchInput" name="q" maxlength="250" />
<input type="submit" id="searchSubmit" name="btnG" value="Search" />
<input type="hidden" name="access" value="a" /> 
<input type="hidden" name="output" value="xml_no_dtd" />
<input type="hidden" name="client" value="default_frontend" />
<input type="hidden" name="site" value="default_collection" />
<input type="hidden" name="proxystylesheet" value="default_frontend" />
<input type="hidden" name="as_dt" value="i" />
</form> 		
		</div>
		<div class="navItem"><a href="http://<%=request.servervariables("server_name")%>/default.asp?iPage=1&amp;fsb=0" id="resources">Training Resources/<br>Tools</a></div>
        <div class="navItem"><a href="http://<%=request.servervariables("server_name")%>/default.asp?iPage=2&amp;fsb=0">Organization<br>Points of Contact</a></div>	
		<div class="navItem"><a href="http://<%=request.servervariables("server_name")%>/default.asp?iPage=3&amp;fsb=0" id="mgmtSys">Management System</a></div>
        <div class="navItem"><a href="http://online-training.pnl.gov/default.asp">PNNL's External Training Website</a><span style="font-size: .75em; font-style: italic; ">(Off-site link.)</span></div>
        <div class="navItem"><a href="http://lessons.pnl.gov/">Lessons Learned</a><span style="font-size: .75em; font-style: italic; ">(Off-site link. Use the <span class="goldenrod">BACK</span> button to return)</span></div>
 <div class="navItem"><a href="/default.asp">T&amp;Q Home</a></div>
 
        <div class="navItem"><a href="http://wwwi.pnl.gov/safety">ESH&amp;Q Home</a></div>
		<div class="navItem"><a href="http://labweb.pnl.gov">LabWeb</a></div>
        <div class="navItem"><a href="mailto:training.administrator@pnl.gov">Webmaster</a></div>
		<div><p><a href="http://validator.w3.org/check?uri=referer"><img border='0' src="http://www.w3.org/Icons/valid-html401" alt="Valid HTML 4.01 Transitional" height="31" width="88"></a></p></div>
		<div><%=HitCounter(session("bAddHit")<>"F")%></div>
	</div>
</div>    

<div id="navBarTop">
    <div><a href="http://<%=request.servervariables("server_name")%>/onlinetraining.asp">Online Training</a></div>
    <div><a href="http://<%=request.servervariables("server_name")%>/reading/default.asp">Required Reading</a></div>
    <div><a href="http://<%=request.servervariables("server_name")%>/whatsnew.asp">Hot Off the Press!</a></div>
</div>

<div id="navBarBottom">
    <div><a href="http://<%=request.servervariables("server_name")%>/trainview/default.asp"><em>Your</em> Training Record</a></div>
    <div><a href="http://<%=request.servervariables("server_name")%>/catalog/default.asp">Course Catalog</a></div>
    <div><a href="http://<%=request.servervariables("server_name")%>/newsletters/default.asp">Newsletters</a></div>
    <div><a href="http://<%=request.servervariables("server_name")%>/pi-stats/default.asp">PI Stats</a></div>
</div>
<a name="skipNav" id="skipNav"></a>


