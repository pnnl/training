/*****************************************************************************/
/*Pop-Up and follow box													     */
/*****************************************************************************/

/*sets the coordinates for the pop-up*/
var x, y;
// Example: obj = findObj("image1");
function findObj(theObj, theDoc)
{
	var p, i, foundObj;
		
	if(!theDoc) theDoc = document;
	if( (p = theObj.indexOf("?")) > 0 && parent.frames.length)
	{
		theDoc = parent.frames[theObj.substring(p+1)].document;
		theObj = theObj.substring(0,p);
	}
	if(!(foundObj = theDoc[theObj]) && theDoc.all) foundObj = theDoc.all[theObj];
	for (i=0; !foundObj && i < theDoc.forms.length; i++) 
		foundObj = theDoc.forms[i][theObj];
	for(i=0; !foundObj && theDoc.layers && i < theDoc.layers.length; i++) 
		foundObj = findObj(theObj,theDoc.layers[i].document);
	if(!foundObj && document.getElementById) foundObj = document.getElementById(theObj);
		
	return foundObj;
}
         
function write_top_links(userType)
{
	document.write("<li><a title='Go to main page.' href='\Main.aspx' id='mainpage'>Main</a></li>");
	document.write("<li><a title='Go to templates.' href='\Template.aspx' id='template'>Templates</a></li>");
	
	if(userType != 1)
	{
		document.write("<li><a title='Go to survey response manual input.' href='\SurveyResponse.aspx' id='surveyinput'>Survey Input</a></li>");
	}
		
	if(userType == 4)
	{
		document.write("<li><a title='Go manage users.' href='\ManageUsers.aspx' id='manageusers'>Manage Users</a></li>");
	}
}

function write_top_links_disabled()
{
    //do nothing;
}

function write_bottom_links(userType, Machine, approved)
{
	if(userType != 1)
	{
		document.write("<li><a title='Go generate reports.' href='\Report.aspx' id='report'>Reports</a></li>");
	}
	
	if(userType == 4)
	{
		document.write("<li><a title='Go set messages.' href='\Message.aspx' id='message'>Message</a></li>");
	}
	
	document.write("<li><a title='Go get help.' href='\Help.aspx' id='help'>Help</a></li>");
	document.write("<li><a title='Go reset your password.' href='\Password.aspx' id='password'>Password</a></li>");
	
	if(userType != 1)
	{
		document.write("<li><a title='Go send a survey.' href='\SendSurvey.aspx' id='sendsurvey'>Send Survey</a></li>");
	}

	document.write("<li><a style='COLOR: #CC0000' title='Go log out.' href='\login.aspx' id='logout'>Log Out</a><li>");
}

function write_bottom_links_disabled()
{
	//do nothing;
}