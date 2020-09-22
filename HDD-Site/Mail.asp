<% Language=VBScript %>

<%
'FullString = Request.QueryString
'FullString = split(FullString,"?")
'HalfString = FullString(0)
'HalfString = split(HalfString,"=")

'YName = split(HalfString(1),"&")
'YEmail = split(HalfString(2),"&")
'FName = split(HalfString(3),"&")
'FEmail = split(HalfString(4),"&")
'YName(0) = replace(YName(0),"+"," ")
'FName(0) = replace(FName(0),"+"," ")
'Response.Write YName(0) & "<br>"
'Response.Write YEmail(0) & "<br>"
'Response.Write FName(0) & "<br>"
'Response.Write FEmail(0) & "<br>"

FromName = Request.Form("txtYName")
FromEmail = Request.Form("txtYEMail")
ToName = Request.Form("txtFName")
ToEMail = Request.Form("txtFEmail")

'FromEmail = YEmail(0)
'ToEMail = FEmail(0)
Subject = "Just Testing man.."
EMailMessage="This is testing.. so dont worry..."
Impt=1
Dim objMail
set objMail = CreateObject("CDONTS.NewMail")
objMail.Send FromEmail, ToEMail, Subject, EmailMessage, Impt
set objMail = nothing

Dim objMail1
'set objMail1 = CreateObject("CDONTS.NewMail")
'objMail1.Send FromEmail, ToEMail, Subject, EmailMessage, Impt
set objMail1 = nothing
%>

<html>
<head>
<META name=VI60_defaultClientScript content=VBScript>
<title>Hirdhav Digital Diary -- HOME</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="24%" bgcolor="#0A7AE6"><font face="Arial, Helvetica, sans-serif" size="+7"><b><img height=50 src="../Images/Hirdhav.jpg" width=220></b></font></td>
    <td width="76%" bgcolor="#0A7AE6">&nbsp;</td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr bgcolor="#0A7AE6"> 
    <td> 
      <style type="text/css"> <!--
	}A:hover {text-decoration: none; color: red
	}
	.submenuStyle {
    position: absolute;
    visibility: hidden;
}
.menuStyle {
	visibility: visable;
}
.m
{
    FONT-SIZE: 10px;
    FONT-WEIGHT: bolder;
    COLOR: white;
    BORDER-BOTTOM: white 0px solid;
    FONT-FAMILY: verdana;
    TEXT-DECORATION: none
}
//--></style>
      <script language="JavaScript"><!--

  
  // How many pixels to adjust down:
  var steps = 14;
  // How many total items untill the table is splits:
  var split = 10;
  // How many pixels to adjust left:
  var rsteps = 6;
  
  
  // Try these for vertical menus:
  // var steps = -6;
  // How many total items untill the table is splits:
  // var split = 200	;
  // How many pixels to adjust left:
  // var rsteps = -74;
 
// Create new main array.
 var menuitems =		[ ["http://www.hirdhav.com", "Hirdhav.com Home"],
						  ["href 2", "Past News", "herf 2", "Current News"],
						  ["herf 3", "HD Product Catalog", "herf 3", "Downloads"],
						  ["downloads/downloads.htm", "Download Center"],
						  ["herf 5", "Company Overview"],
						  ["herf 6", "Contact Us", "herf 4", "Profile Center"]
						]

 //Generate Span Layers -Charles Toepfer
function generate_layers()
{ 
  var x=0; 
  browser_type = navigator.appName;
 
  for (x=0; x< menuitems.length; x++) 
	  { 
		if (browser_type ==  "Microsoft Internet Explorer") { 
		document.writeln ('<div id=submenu'+ (x+1)  +' name=submenu'+ (x+1)  +' class=submenuStyle style="z-index:1;">');
		}
		else {
		document.writeln ('<div id=submenu'+ (x+1)  +' name=submenu'+ (x+1)  +' class=submenuStyle>');
		}		
		document.writeln ('<table border=0 bgcolor=#0A7AE6 bordercolor=#ffffff cellspacing=0 cellpadding=5><tr><td valign=top>')
		 var intmaxsplit = 50
			if ((menuitems[x].length/4) > (split/2)) { 
			var intmaxsplit = (menuitems[x].length/2) 
			if  (((intmaxsplit) % 2) == 0.5) {var intmaxsplit = intmaxsplit + 1;} 
			}
		  for (xx=0; xx< menuitems[x].length; xx++) 
		  { 
			if (browser_type ==  "Microsoft Internet Explorer") {
			document.writeln ('<TABLE width=120 border=0 cellspacing=0 cellpadding=0><TR>')
			document.writeln ('<TD style=width:100%; onMouseover=this.style.backgroundColor="#0A7AE6"; onMouseout=this.style.backgroundColor="";>')
			document.writeln ('<a href = ' + menuitems[x][xx] + ' class=m >');
			document.writeln ('' + menuitems[x][xx+1] + '</a><br>');
			document.writeln ('</TD></TR></TABLE>'); 
			}
			else {
			document.writeln ('<a href = ' + menuitems[x][xx] + ' class=m >');
			document.writeln ('' + menuitems[x][xx+1] + '</a><br>');
			}
		   xx = xx +1
		    if ((xx) > intmaxsplit - 2) { document.writeln('</td><td valign=top>') 
		    var intmaxsplit = 500
		    }   
		  }
		document.writeln ('<br>'); 
		document.writeln ('</td></tr></table></div>');  
		  } 
	 }
generate_layers();

if (browser_type == "Netscape" && (browser_version >= 4) && navigator.platform.indexOf("Mac") >= 0) { var steps = steps - 20 }

function getAnchorPosition(anchorname) {
	var useWindow = false;
	var coordinates = new Object();
	var x=0;
	var y=0;
	var w_gebi = false;
	var w_css = false;
	var w_layers = false;
	if (document.getElementById) { w_gebi = true; }
	else if (document.all) { w_css = true; }
	else if (document.layers) { w_layers = true; }

 	if (w_gebi && document.all) {
		x = AnchorPosition_getPageOffsetLeft(document.all[anchorname]);
		y = AnchorPosition_getPageOffsetTop(document.all[anchorname]);
		}
	else if (w_gebi) {
		var o = document.getElementById(anchorname);
		x = o.offsetLeft;
		y = o.offsetTop;
		}
 	else if (w_css) {
		x = AnchorPosition_getPageOffsetLeft(document.all[anchorname]);
		y = AnchorPosition_getPageOffsetTop(document.all[anchorname]);
		}
	else if (w_layers) {
		var found=0;
		for (var i=0; i<document.anchors.length; i++) {
			if (document.anchors[i].name == anchorname) {
				found=1;
				break;
				}
			}
		if (found == 0) {
			coordinates.x=0; coordinates.y=0; return coordinates;
			}
		x = document.anchors[i].x;
		y = document.anchors[i].y;
		}
	else {
		coordinates.x=0; coordinates.y=0; return coordinates;
		}
	coordinates.x = x;
	coordinates.y = y;
	return coordinates;
	}




function AnchorPosition_getPageOffsetLeft (el) {
	var ol = el.offsetLeft;
	while ((el = el.offsetParent) != null) { 
		ol += el.offsetLeft; 
		}
	return ol;
	}
function AnchorPosition_getWindowOffsetLeft (el) {
	var scrollamount = document.body.scrollLeft;
	return AnchorPosition_getPageOffsetLeft(el)-scrollamount;
	}	
function AnchorPosition_getPageOffsetTop (el) {
	var ot = el.offsetTop;
	while((el = el.offsetParent) != null) { 
		ot += el.offsetTop; 
		}
	return ot;
	}

function AnchorPosition_getWindowOffsetTop (el) {
	var scrollamount = document.body.scrollTop;
	return AnchorPosition_getPageOffsetTop(el)-scrollamount;
	}

function removeall() {
	for(i = 1; i <= menuitems.length; i++) { remove(i);  }
}

function show(object) {

	var c = getAnchorPosition(object.replace(/submenu/, "menu"));
	if (c.x < 1) { c.x = 21; }
	if (c.y < 1) { c.y = 155; }
	
    if (document.getElementById && document.getElementById(object) != null) {
        document.getElementById(object).style.top=c.y + steps+"px";
        document.getElementById(object).style.left= c.x - rsteps+"px";
        node = document.getElementById(object).style.visibility='visible';
        document.body.onclick= removeall;
    }
    else if (document.layers && document.layers[object] != null) {
		 document.layers[object].visibility = 'visible';
		 document.layers[object].top = c.y + steps;
		 document.layers[object].left = c.x - rsteps;
		 document.captureEvents(Event.MOUSEMOVE | Event.MOUSEDOWN);
		 document.onmousedown = removeall;
		 }
    else if (document.all) 
	  {
        document.all[object].style.visibility = 'visible';
        document.all[object].style.pixelTop = c.y + steps;
		document.all[object].style.pixelLeft = c.x - rsteps;
		document.body.onclick= removeall;
		}

	var id = object.replace(/submenu/, "");
	for(i = 1; i <= menuitems.length; i++)
	{
		if  (i != id) {
		remove(i);
		}
	}
	var c = null
}


function removeall() {
	for(i = 1; i <= menuitems.length; i++) { remove(i);  }
}



function remove(id) {
    
      if (document.getElementById && document.getElementById('submenu'+id) != null) {
        node = document.getElementById('submenu'+id).style.visibility='hidden';
    }
    else if (document.layers && document.layers['submenu'+id] != null) {
		 document.layers['submenu'+id].visibility = 'hidden';
		 }
    else if (document.all) 
	  {
        document.all['submenu'+id].style.visibility = 'hidden';
		}
}

//--></script>
      <table border=0 bgcolor=#0A7AE6 bordercolor=#0A7AE6 cellspacing=0 cellpadding=1>
        <tr> 
          <td width="84"><a href="../index.htm" onMouseOver="show('submenu1');this.style.color='red'" onMouseOut="this.style.color='white'" class="m" id="menu1" name="menu1"> 
            &nbsp Home </a><font face="verdana" color="white" size="1"> &nbsp;&nbsp|&nbsp;&nbsp 
            </font></td>
          <td width="75"><a href="index.htm" onMouseOver="show('submenu2');this.style.color='red'" onMouseOut="this.style.color='white'" class="m" id="menu2" name="menu2">News 
            </a><font face="verdana" color="white" size="1"> &nbsp;&nbsp|&nbsp;&nbsp 
            </font></td>
          <td width="93"><a href="index.htm" onMouseOver="show('submenu3');this.style.color='red'" onMouseOut="this.style.color='white'" class="m" id="menu3" name="menu3">Products 
            </a><font face="verdana" color="white" size="1"> &nbsp;&nbsp|&nbsp;&nbsp 
            </font></td>
          <td width="109"><a href="../downloads/downloads.htm" onMouseOver="show('submenu4');this.style.color='red'" onMouseOut="this.style.color='white'" class="m" id="menu4" name="menu4">Downloads 
            </a><font face="verdana" color="white" size="1"> &nbsp;&nbsp|&nbsp;&nbsp 
            </font></td>
          <td width="127"><a href="../About/abt.htm" onMouseOver="show('submenu5');this.style.color='red'" onMouseOut="this.style.color='white'" class="m" id="menu5" name="menu5">About 
            Hirdhav </a><font face="verdana" color="white" size="1"> &nbsp;&nbsp|&nbsp;&nbsp 
            </font></td>
          <td width="139"><a href="index.htm" onMouseOver="show('submenu6');this.style.color='red'" onMouseOut="this.style.color='white'" class="m" id="menu6" name="menu6">Hirdhav.com 
            Guide </a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#C1C1C1" width="23%" height="80"> 
      <table width="94%" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#F1F1F1">
        <tr>
          <td><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;HDD 
            Related</font></b>
            <table width="94%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr bgcolor="#F1F1F1"> 
                <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;History 
                  Of HDD</font></td>
              </tr>
              <tr bgcolor="#F1F1F1"> 
                <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;Next 
                  Version</font></td>
              </tr>
              <tr bgcolor="#F1F1F1">
                <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;Past 
                  Versions</font></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
    <td> 
      <table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td><font face="Verdana, Arial, Helvetica, sans-serif" size="5"><b>Hirdhav 
            Digital Diary (HDD)</b></font></td>
        </tr>
      </table>
      <img src="../Images/Line.jpg" width="400" height="2"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="23%" bgcolor="#C1C1C1" valign="top"> 
      <table width="94%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr bgcolor="#F1F1F1"> 
          <td>&nbsp;<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>Others</b></font>
            <table width="94%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr>
                <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;Screen 
                  Shots</font></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
    <td>
      <table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td>
            <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Hirdhav 
              Digital Diary is first official software of Hirdhav. And Hirdhav 
              has started its development since 19th September 2001 with its version 
              1.0 and now Hirdhav Digital Diary is at its version 1.3. Daily Hirdhav 
              Digital Diary is growing very fast. Hirdhav Digital Diary is also 
              known as HDD. HDD 1.3 is going to be re-launch on 15th May 2002. 
              We had released 1.3 on 31st March 2002, but due to some major bugs 
              we had to take it back. Explorer some of your questions from given 
              below answers.</font></p>
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>What 
              is Hirdhav Digital Diary (HDD)?</b></font> </p>
            <p> <font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;This 
              question has been asked to us more than 150 times and still people 
              are aking so we have desided to put it on the web, so people can 
              get to know what is Hirdhav Digital Diary.</font></p>
            <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Hirdhav 
              Digital Diary is also know as HDD. We have been working on this 
              software since September 2001. Hirdhav Digital Diary is a software 
              which keeps your personal information like your Contacts, Memos, 
              Reminders, Schedulers and lots, and the few thing about this software 
              which people likes most is that, this software is Freeware no ads 
              no evalution period, This software is having Auto Update future 
              so need to download all the time whole large setup files, Download 
              only once and rest it will do, MultiUser capablities, This software 
              is having MutiUser capablities, so it doesnt matter how many people 
              are using this software on one computer. You can create unlimited 
              users and also having lots of futures.</font></p>
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>Why 
              Hirdhav is Developing this software?</b></font></p>
            <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Originally 
              idea of this software came in mind of Dhaval Faria Chairman and 
              Chief Software Architect of this Company. So we have decided to 
              put Dhaval Faria's speech on this page.</font></p>
            <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&quot;Originally 
              idea of this software came in my mind, But Honestly speaking I was 
              working with Mukesh Parikh and he gave me idea of making calender 
              which stores personal reminder and scheduler of the user who uses 
              that diary. I accepted idea and I was thinking for it for 2 days 
              and suddenly one thought came in my mind, instead of making only 
              calender why not to make whole Personal Organizer, and after that 
              on the spot I started working on it from scratch. So origin is Mukesh 
              Parikh but original idea is of mine. To make this software more 
              powerfull lots of people have helped me. Here are there names: Mukesh 
              Parikh, Dhaval Faria (My Self), Ramnik Faria, Hiren Faria, Manjula 
              Faria, Vighnesh Prabhu, Kaustubh Gujar, John Couture. You can also 
              find name of this people in the Credits section of Hirdhav Digital 
              Diary. So its very simple that we are developing this software to 
              store the personal information of users.&quot;</font></p>
            <p align="center"><img src="../Images/Line.jpg" width="500" height="2"></p>
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2">To 
              Explore more about Hirdhav Digital Diary Visit to Below Links<br>
              </font></p>
            <p><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Hirdhav 
              Digital Diary 1.0<br>
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Hirdhav 
              Digital Diary 1.1<br>
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Hirdhav 
              Digital Diary 1.2<br>
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Hirdhav 
              Digital Diary 1.3</b></font></p>
            <p>&nbsp;</p>
            <table width="94%" border="0" cellspacing="0" cellpadding="0" align="center" height="34">
              <tr> 
                <td bgcolor="#0A7AE6" width="1%">&nbsp;</td>
                <td bgcolor="#0A7AE6" width="96%"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF">&nbsp;Spread 
                  The Word</font></b></td>
                <td bgcolor="#0A7AE6" width="1%">&nbsp;</td>
              </tr>
              <tr> 
                <td bgcolor="#0A7AE6" width="1%">&nbsp;</td>
                <td width="96%">
                  <div align="center"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Thanks 
                    you very much for Spreading Word.<br>Your E-Mail has been sent to <%Response.Write ToEMail%></font></b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b> 
                    </b></font> </div>
                </td>
                <td bgcolor="#0A7AE6" width="1%">&nbsp;</td>
              </tr>
              <tr>
                <td bgcolor="#0A7AE6" width="1%">&nbsp;</td>
                <td width="96%" bgcolor="#0A7AE6"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF">&nbsp;Spread 
                  The Word</font></b></td>
                <td bgcolor="#0A7AE6" width="1%">&nbsp;</td>
              </tr>
            </table>
            
          <p>&nbsp;</p></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#0A7AE6"> 
    <td> &nbsp;<font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b>Contact 
      Us<br>
      </b> &copy;2002 Hirdhav. All rights reserved. <b>Terms Of Use</b><br>
      <br>
      All the software names and company names used in this site are registered 
      trademart of there respective owner.<br>
      <br>
      This site can be best shown under screen resolution of 800 x 600 and with 
      the 16 bit or higher color.<br>
      To View this site in its original design please use Microsoft Internet Explorer 
      5.5 or higher.</font></td>
  </tr>
</table>
</body>
</html>