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
  <!--#include file="../../global/header.shtml" -->
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
