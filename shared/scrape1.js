function MM_preloadImages() //v3.0
{
	var d=document;
	if(d.images)
	{
		if(!d.MM_p) d.MM_p=new Array();
		var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
		if (a[i].indexOf("#")!=0)
		{
			d.MM_p[j]=new Image;
			d.MM_p[j++].src=a[i];
		}
	}
}

function MM_swapImgRestore() //v3.0
{
	var i,x,a=document.MM_sr;
	for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) //v3.0
{
	var p,i,x;
	if(!d) d=document;
	if((p=n.indexOf("?"))>0&&parent.frames.length)
	{
		d=parent.frames[n.substring(p+1)].document;
		n=n.substring(0,p);
	}
	if(!(x=d[n])&&d.all) x=d.all[n];
	for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
	for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
	return x;
}

function MM_swapImage() //v3.0
{
	var i,j=0,x,a=MM_swapImage.arguments;
	document.MM_sr=new Array;
	for(i=0;i<(a.length-2);i+=3)
	if ((x=MM_findObj(a[i]))!=null)
	{
		document.MM_sr[j++]=x;
		if(!x.oSrc) x.oSrc=x.src;
		x.src=a[i+2];
	}
}

function getPnlLink()
{
	if(window.location.protocol.toLowerCase() == "https:")
		document.location = "http://www.pnl.gov";
	else
		document.location = "http://labweb.pnl.gov/default.asp";
}

function saSiteMap( url, sa )
{
	// parameters for subject area site map pop-up window
	var map = window.open( '/standard/0003t010.htm?SAURL=' + encodeURIComponent(url), 'SubjectArea' + sa, 'top=10, left=10, width=732, height=517, resizable=no, scrollbars=yes, toolbars=no, location=no, directories=no, status=no, menubar=no, copyhistory=no' )
	map.focus();
}


//functions for the keyword index
var sbmsKwChar = "";
var sbmsKwReady = false;
var sbmsKwOrigKw = "";

function sbmsKeywordsReady()
{
	var kwForm = sbmsGetKwForm();
	if (kwForm == null)
	{
		setTimeout("sbmsKeywordsReady()", 200);
		return;
	}
	if (!kwForm.sbmsKwIndex.value)
	{
		sbmsHideKwFrame();
	}
	else
	{
		window.frames.sbmsKwFrame.showKw(kwForm.sbmsKwIndex.value);
	}
	if (kwForm.sbmsKwIndex.value != "")
		kwForm.sbmsKwIndex.focus();
	sbmsKwReady = true;
}

function sbmsGetKwForm()
{
	var i;
	for (i=0; i<document.forms.length; i++)
	{
		if (document.forms[i].sbmsKwIndex)
		{
			return (document.forms[i]);
		}
	}
	return null;
}

function sbmsHandleKwKeyUp(evt, val)
{
	var kwFrame = window.frames.sbmsKwFrame;
	var keyCode = ((evt.which)?evt.which:evt.keyCode);
	if (keyCode == 13) //enter key
	{
		sbmsGotoKw(val);
		return (false);
	}
	else if (keyCode == 38 && sbmsKwReady) //up arrow
	{
		if (sbmsKwOrigKw != "")
		{
			kwFrame.hlPrev();
		}
		return (false);
	}
	else if (keyCode == 40 && sbmsKwReady) //down arrow
	{
		if (sbmsKwOrigKw != "")
		{
			kwFrame.hlNext();
		}
		return (false);
	}
	else if (keyCode == 27) //escape key
	{
		sbmsHideKwFrame();
		return (false);
	}
	sbmsKwOrigKw = "";
	if (val.length)
	{
		sbmsSetKwFrameDisplay("inline");
		sbmsKwOrigKw = val;
		var kwForm = sbmsGetKwForm();
		var firstChr = val.slice(0,1).toLowerCase();
		if (firstChr != sbmsKwChar)
		{
			sbmsKwReady = false;
			sbmsKwChar = firstChr;
			newLoc = kwFrame.location.href.replace(/\?sbmsidx=[^&]*/i, "?sbmsidx=" + sbmsDoEncode(firstChr));
			kwFrame.location.replace(newLoc);
		}
		else if (sbmsKwReady)
		{
			kwFrame.showKw(val);
		}
	}
	else
	{
		sbmsHideKwFrame();
	}
	return (true);
}

function sbmsDoEncode(thestr)
{
	if (window.encodeURIComponent)
		return window.encodeURIComponent(thestr);
	else if (window.encodeURI)
		return window.encodeURI(thestr);
	else if (window.escape)
		return window.escape(thestr);
	else if (String.replace)
		return thestr.replace(/\W/g, "0");
	else
		return thestr;
}

function sbmsSetKwFrameHeight(newHeight)
{
	//var kwTemp = window.frames.sbmsKwFrame;
	var kwForm = sbmsGetKwForm();
	var kwFrame = document.getElementById("sbmsKwFrame");
	newHeight = parseInt(newHeight);

	kwFrame.style.width = kwForm.sbmsKwIndex.offsetWidth + "px";
	kwFrame.style.width = 300;
	//alert("kwTemp.document.body.offsetWidth");
	kwFrame.style.height = ((newHeight>150)?150:newHeight+35) + "px";
	kwFrame.style.borderWidth = "1px";
}

function sbmsHideKwFrame()
{
	sbmsKwOrigKw = "";
	document.getElementById("sbmsKwFrame").style.height="0px";
	document.getElementById("sbmsKwFrame").style.borderWidth = "0px";
	sbmsSetKwFrameDisplay("none");
}

function sbmsGotoKw(kw)
{
	document.location="/standard/0001t010.htm?kwIndex=" + encodeURIComponent(kw) ;
}

function sbmsSetKwFrameDisplay(newDisplay)
{
	if (document.all)
	{
		document.getElementById("sbmsKwFrame").style.display = newDisplay;
	}
}
function sbmsFillInKw(kw)
{
	var kwInput = sbmsGetKwForm().sbmsKwIndex;
	if (kw == "")
		kw = sbmsKwOrigKw;
		kwInput.value = kw;
	if (kwInput.createTextRange)
	{ //IE
		var tr = kwInput.createTextRange();
		tr.moveStart("character", sbmsKwOrigKw.length);
		tr.moveEnd("character", kw.length);
		tr.select();
	}
	else if (kwInput.setSelectionRange)
	{ //Mozilla?
		kwInput.setSelectionRange(sbmsKwOrigKw.length, kw.length)
	}
}

function expandHint()
{
	var hints = document.getElementById("hints");
	var hintsLink = document.getElementById("hintsLink");
	if (hints)
	{
		if (hints.style.display == "block")
		{
			hints.style.display = "none";
			hintsLink.style.display = "block";
		}
		else
		{
			hints.style.display = "block";
			hintsLink.style.display = "none";
		}
	}
}

function objDynamicLink(anch)
{
	this.dynaLink = anch;
	this.id = anch.getAttribute("dl:linkId");
	this.type = anch.getAttribute("dl:linkType");
	this.linkId = this.type + this.id;
	this.title = anch.getAttribute("title");

	this.wsUpdate = _wsUpdate;
}

function _wsUpdate(linkId, url, text, shortText, title, status, statusCode, contactInfo)
{
	if (this.linkId != linkId)
		return;
	if (url == "Error" || status != "A" || statusCode != "200")
		return;
	var oldhref = this.dynaLink.getAttribute("href");
	if (oldhref.indexOf("#") != -1)
	{
		url += oldhref.slice(oldhref.indexOf("#"));
	}
	this.dynaLink.setAttribute("href", url);
	if (this.dynaLink.getAttribute("dl:useThisText") == "N")
	{
		this.dynaLink.innerHTML = "";
		this.dynaLink.appendChild(document.createTextNode(text));
	}
	if (this.dynaLink.getAttribute("dl:useTitle") != "N")
	{
		this.dynaLink.setAttribute("title", title);
	}
	if (this.dynaLink.getAttribute("dl:useShortText") == "Y")
	{
		this.dynaLink.innerHTML = "";
		this.dynaLink.appendChild(document.createTextNode(shortText));
	}
	if (this.dynaLink.getAttribute("dl:linkClass") != null)
	{
		this.dynaLink.className = this.dynaLink.getAttribute("dl:linkClass");
	}
	if (this.dynaLink.getAttribute("dl:linkStyle") != null)
	{
		this.dynaLink.style.cssText = this.dynaLink.getAttribute("dl:linkStyle");
	}
	if (this.dynaLink.getAttribute("dl:linkAttributes") != null)
	{
		var attrs = this.dynaLink.getAttribute("dl:linkAttributes").split(";");
		var i;
		for (i=0; i<attrs.length; i++)
		{
			if (attrs[i].indexOf("=") != -1)
			{
				var attr = attrs[i].split("=");
				attr[0] = attr[0].replace(/(^\s+|\s+$)/g, "");
				attr[1] = attr[1].replace(/(^\s+|\s+$)/g, "");
				this.dynaLink.setAttribute(attr[0], attr[1]);
			}
		}
	}
}

var g_DynamicLinks = new Array();
var g_XmlHttp = null;

function getXmlHttp()
{
	if (g_XmlHttp)
		return g_XmlHttp;
	if (window.XMLHttpRequest)
		return new XMLHttpRequest();
	else if (window.ActiveXObject)
		return new ActiveXObject("MSXML2.XMLHTTP");
	else
		return null;
}

function getDynamicLinks()
{
	g_XmlHttp = getXmlHttp();
	if (g_XmlHttp == null)
		return;

	var i;
	var links = "<?xml version=\"1.0\"?>"
		+ "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">"
		+ "<soap:Body>"
		+ "<GetUrls xmlns=\"http://sbms.pnl.gov/ap/05/\">"
		+ "<links>";
	var anchs = document.getElementsByTagName("a");
	for (i=0; i<anchs.length; i++)
	{
		if (anchs[i].getAttribute("dl:linkId") != null)
		{
			var dynLink = new objDynamicLink(anchs[i]);
			links += "<string>" + dynLink.linkId + "</string>";
			g_DynamicLinks.push(dynLink);
		}
	}
	links += "</links></GetUrls></soap:Body></soap:Envelope>";

	g_XmlHttp.open("POST", "/ap/05/ap05d020.asmx", true);
	g_XmlHttp.setRequestHeader("Content-Type", "text/xml");
	g_XmlHttp.onreadystatechange = _updateLinks;
	g_XmlHttp.send(links);
}

function _updateLinks()
{
	if (g_XmlHttp.readyState != 4)
		return;
	if (g_XmlHttp.status != 200)
	{
		alert("Request for Dynamic Links resulted in result code: "
			+ g_XmlHttp.status + "\n"
			+ g_XmlHttp.responseText);
		return;
	}
	var dynLinks = g_XmlHttp.responseXML.getElementsByTagName("DynamicLink");
	var i, j;
	for (i=0; i<dynLinks.length; i++)
	{
		var dynLinkId =     getChildTagText(dynLinks[i], "LinkId");
		var dynHref =       getChildTagText(dynLinks[i], "Url");
		var dynText =       getChildTagText(dynLinks[i], "Text");
		var dynTitle =      getChildTagText(dynLinks[i], "Title");
		var dynShortText =  getChildTagText(dynLinks[i], "ShortText");
		var dynStatus =     getChildTagText(dynLinks[i], "Status");
		var dynStatusCode = getChildTagText(dynLinks[i], "StatusCode");
		var dynContact =    getChildTagText(dynLinks[i], "ContactName")
			+ "\n" + getChildTagText(dynLinks[i], "ContactPhone")
			+ "\n" + getChildTagText(dynLinks[i], "ContactEmail");
		//alert( dynLinkId
		//	+ "\n" + dynText
		//	+ "\n" + dynHref
		//	+ "\n" + dynStatus
		//	+ "\n" + dynStatusCode );
		for (j=0; j<g_DynamicLinks.length; j++)
		{
			if (g_DynamicLinks[j].linkId == dynLinkId)
			{
				g_DynamicLinks[j].wsUpdate(dynLinkId,
					dynHref,
					dynText,
					dynShortText,
					dynTitle,
					dynStatus,
					dynStatusCode,
					dynContact);
			}
		}
	}

	g_XmlHttp = null;

	function getChildTagText(node, tagName)
	{
		var tags = node.getElementsByTagName(tagName);
		if (tags.length == 0)
			return "";
		var tag = tags[0];
		var txt = "";
		var i;
		for (i=0; i<tag.childNodes.length; i++)
		{
			if (tag.childNodes.item(i).nodeType == 3)
				txt += tag.childNodes.item(i).data;
		}
		return txt;
	}
}

// JavaScript Document

function saOnload() //v3.0
{
	if( window.getDynamicLinks ){
		getDynamicLinks();
	}
}