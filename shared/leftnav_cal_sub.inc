<div id="leftNav">
<h2 class="first">Enterprise Learning</h2>

<ul>
<li><a href="/default.asp" id="elhome">Home</a></li>
<% if (bConquistador) then %>
<li><A tabIndex="-1" id="calEL" href="../sectionManage.asp" alt="Manage and Create Class Sections" title="Manage and Create Sections">Section Management</A></li>
<% end if %>
</ul>
 	
<h2>Training</h2>
<ul>
<li><a tabIndex="-1" href="../default.asp" id="calHome">Training Calendar</a></li>
<li><a tabIndex="-1" href="../classes.asp" id="calClass">Available Classes</a></li>
<li><a tabIndex="-1" href="http://<%=application("MSRS")%>/ReportServer?%2fTraining+and+Qualification%2fAll+Classes&rs%3aCommand=Render">All Classes</a></li>
<li><a tabIndex="-1" href="http://<%=application("MSRS")%>/ReportServer?/Training+and+Qualification/My+Training&EMPLID=<%=user2emplid("")%>&rs:Command=Render&rc:Parameters=False">My Calendar Training</a></li>
</ul>
<% if (bAdmin) then %>
<h2>Admin Menu</h2>
<ul>
<li><A tabIndex="-1" id="calManage" href="../classManage.asp" alt="Manage and Create Classes" title="Manage and Create Classes">Class Management</A></li>
<li><A tabIndex="-1" id="calInstruct" href="../instructManage.asp" title="Create and Remove Instructors" alt="Create and Remove Instructors">Instructors</A></li>
</ul>
<% end if %>

<h2>Directorates</h2>
<ul>
<li><a href="http://business.pnl.gov/">Business Systems</a></li>
<li><a href="http://hr.pnl.gov/">Human Resources</a></li>
</ul>

</div>