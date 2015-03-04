<div id="leftNav">
	<h2 class="first">Enterprise Learning</h2>
	<ul>
		<li>
			<a href="/default.asp" id="elhome">Home</a>
		</li>
	</ul>
	<% if Session("isAuthenticated") = "True" then %>
		<h2>Training</h2>
			<ul>
				<script type="text/javascript">write_top_links(<%=Session("UserType")%>)</script>
				<script type="text/javascript">write_bottom_links(<%=Session("UserType")%>,'<%=Session("Machine")%>')</script>
			</ul>
	<% end if %>
</div>