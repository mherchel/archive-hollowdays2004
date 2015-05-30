<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Guide - Web Wiz Journal
'**                                                              
'**  Copyright 2001-2002 Bruce Corkhill All Rights Reserved.                                
'**
'**  This program is free software; you can modify (at your own risk) any part of it 
'**  under the terms of the License that accompanies this software and use it both 
'**  privately and commercially.
'**
'**  All copyright notices must remain in tacked in the scripts and the 
'**  outputted HTML.
'**
'**  You may use parts of this program in your own private work, but you may NOT
'**  redistribute, repackage, or sell the whole or any part of this program even 
'**  if it is modified or reverse engineered in whole or in part without express 
'**  permission from the author.
'**
'**  You may not pass the whole or any part of this application off as your own work.
'**   
'**  All links to Web Wiz Guide and powered by logo's must remain unchanged and in place
'**  and must remain visible when the pages are viewed unless permission is first granted
'**  by the copyright holder.
'**
'**  This program is distributed in the hope that it will be useful,
'**  but WITHOUT ANY WARRANTY; without even the implied warranty of
'**  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE OR ANY OTHER 
'**  WARRANTIES WHETHER EXPRESSED OR IMPLIED.
'**
'**  You should have received a copy of the License along with this program; 
'**  if not, write to:- Web Wiz Guide, PO Box 4982, Bournemouth, BH8 8XP, United Kingdom.
'**    
'**
'**  No official support is available for this program but you may post support questions at: -
'**  http://www.webwizguide.info/forum
'**
'**  Support questions are NOT answered by e-mail ever!
'**
'**  For correspondence or non support questions contact: -
'**  info@webwizguide.com
'**
'**  or at: -
'**
'**  Web Wiz Guide, PO Box 4982, Bournemouth, BH8 8XP, United Kingdom
'**
'****************************************************************************************



'If the session variable is False or does not exsist then redirect the user to the unauthorised user page
If Session("blnIsUserGood") = False or IsNull(Session("blnIsUserGood")) = True then
	'Redirect to unathorised user page
	Response.Redirect"unauthorised_user_page.htm"
End If
%>
<html>
<head>
<title>Hollow Days News Update</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="copyright" content="Copyright (C) 2001-2002 Bruce Corkhill">
<!-- The Web Wiz Journal is written by Bruce Corkhill ©2001-2002
	If you want your own Web Wiz Journal then goto http://www.webwizguide.info -->
<style type="text/css">
<!--
a:link {
	color: #FFFFFF;
}
a:visited {
	color: #FFFFFF;
}
a:hover {
	color: #FFFFFF;
	text-decoration: none;
}
a:active {
	color: #FFFFFF;
}
-->
</style>
</head>
<body bgcolor="#000000" text="#CCCCCC">
<div align="left"> 
  <!--
///////////////////////////////////////////////////////////
//                                                       //
//              Website Design by Herchel                //
//                   www.herchel.com                     //
//                                                       //
//       Copyright 2002-2003 All Rights Reserved         //
//                                                       //
///////////////////////////////////////////////////////////
-->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="description" content="Hollow Days is a hard rock band out of Gainesville whose powerful, high energy songs are making a strong impression on the Florida music scene. Instead of trying to follow the trends of the moment, the guys in Hollow Days do what they know best: straight up solid hard rock. Inspired and influenced by bands like The Doors, Black Sabbath, Stone Temple Pilots and Guns n' Roses, Josh Mauldin (lead vocals), Joe Herchel (guitar), Jeremy Redding (bass, vocals, keyboard), and Luke Pidgeon (drums), craft songs that are at once original yet somehow familiar. Strong hooks, soaring choruses and killer groves conduct the power and emotion of these songs to the listener.">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

function MM_displayStatusMsg(msgStr) { //v1.0
  status=msgStr;
  document.MM_returnValue = true;
}
//-->
</script>
<link href="/global/main.css" rel="stylesheet" type="text/css">
<link href="/global/forms.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#000000" text="#CCCCCC">
<table width="740" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="46" height="1"></td>
    <td width="265"></td>
    <td width="107"></td>
    <td width="322" rowspan="2" valign="top"><div align="right"><span class="font16"><a href="/shows/">shows<img src="/global/guitar.gif" width="40" height="40" border="0"></a></span> 
        <span class="font16"><a href="/mp3s/">music<img src="/global/note.gif" width="40" height="40" border="0"></a></span> 
        <a href="/pictures/" class="font16">pictures<img src="/global/camera.gif" width="40" height="40" border="0"></a></div></td>
  </tr>
  <tr> 
    <td rowspan="3" valign="top"><a href="/"><img src="/global/h.gif" width="46" height="101" border="0"></a></td>
    <td rowspan="2" valign="top"><a href="/"><img src="/global/ollow.gif" width="265" height="80" border="0"></a></td>
    <td height="43">&nbsp;</td>
  </tr>
  <tr> 
    <td height="37">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr><td height="21" colspan="3" valign="top" bgcolor="#1F4B49"><div align="right" class="font16">&nbsp;&nbsp;<a href="/">home</a>&nbsp;&nbsp;::&nbsp;&nbsp;<a href="/news/">news</a>&nbsp;&nbsp;::&nbsp;&nbsp;<a href="#" onClick="MM_openBrWindow('/list/','','scrollbars=no,width=200,height=160')" onMouseOver="MM_displayStatusMsg('Join the Mailing List!');return document.MM_returnValue" onMouseOut="MM_displayStatusMsg('');return document.MM_returnValue">mailing&nbsp;list</a>&nbsp;&nbsp;::&nbsp;&nbsp;<a href="/shows/">shows</a>&nbsp;&nbsp;::&nbsp;&nbsp;<a href="/mp3s/">mp3s</a>&nbsp;&nbsp;::&nbsp;&nbsp;<a href="/press/">press</a>&nbsp;&nbsp;::&nbsp;&nbsp;<a href="/bios/">bios</a>&nbsp;&nbsp;::&nbsp;&nbsp;<a href="/pictures/">pictures</a>&nbsp;&nbsp;::&nbsp;&nbsp;<a href="/contact/">contact</a>&nbsp;&nbsp;::&nbsp;&nbsp;<a href="/lyrics/">lyrics</a>&nbsp;&nbsp;::&nbsp;&nbsp;<a href="/links/">links</a>&nbsp;</div></td></tr>
  <p><br>
  </p><table width="628" border="0" cellspacing="0" cellpadding="0">
    <!--DWLayoutTable-->
    <tr> 
      <td width="97" height="16"></td>
      <td width="531"></td>
    </tr>
    <tr>
      <td height="48"></td>
      <td valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="add_journal_form.asp?browser=IE" target="_self">Add 
        New News Item</a> (Windows IE 5+ WYSIWYG HTML Editor) &lt;---Use this 
        one!<br>
        Add New News Item to the web site</font></td>
      </tr>
    <tr> 
      <td height="37"></td>
      <td valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" size="2"><a href="select_journal_item.asp" target="_self">Amend 
        or Delete News Items </a><br>
        Amend or Delete News Items from the web site</font></td>
    </tr>
    <tr> 
      <td height="56"></td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <br>
  <!--#include file="../../global/footer.shtml" -->

</body>
</html>
