<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Hollow Days Admin Menu</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.tables {
	background-color: #1F4B49;
	height: 130px;
	width: 200px;
	border: 1px solid #CCCCCC;
}
.testfields {
	width: 120px;
	border: none;
	color: #000000;
	font-size: 12px;
}
.submit {
	color: #FFFFFF;
	background-color: #666666;
	width: 50px;
	border: 1px solid #000000;
}
a:hover {
	text-decoration: none;
}
-->
</style>
<style type="text/css">
<!--
.webstats {
	color: #FFFFFF;
	background-color: #1F4B49;
	border-top: 1px solid #EEEEEE;
	border-right: 1px solid #000000;
	border-bottom: 1px solid #000000;
	border-left: 1px solid #EEEEEE;
}
.webtable {
	background-color: #1F4B49;
	border: 1px solid #CCCCCC;
}
-->
</style>
<!--#include file="../global/header.shtml" -->

</head>

<body bgcolor="#000000" text="#CCCCCC" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<table width="712" border="0" cellpadding="0" cellspacing="0" class="13">
    <!--DWLayoutTable-->
    <tr> 
      <td width="34" height="18"></td>
      <td width="141"></td>
      <td width="59"></td>
      <td width="28"></td>
      <td width="200"></td>
      <td width="26"></td>
      <td width="39"></td>
      <td width="161"></td>
      <td width="24"></td>
    </tr>
    <tr> 
      <td height="130">&nbsp;</td>
      <td colspan="2" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="tables">
          <!--DWLayoutTable-->
          <tr bgcolor="#CCCCCC"> 
            <td height="21" colspan="4" valign="top"><div align="center"><font color="#000000">Email 
                Logon</font></div></td>
          </tr>
          <tr> 
            <td width="5" height="15"></td>
            <td width="69"></td>
            <td width="120"></td>
            <td width="4"></td>
          </tr>
          <tr> 
            <td height="80">&nbsp;</td>
            <td valign="top" class="13">Email&nbsp;&nbsp;<br>
              Password </td>
            <td valign="top"><form name="form1" method="post" action="http://mail.hollowdays.com/cgi-bin/sqwebmail">
                <div align="right"> 
                  <input name="username" type="text" class="testfields" id="username" value="band@hollowdays.com">
                  <input name="password" type="password" class="testfields" id="password">
                  <input name="Submit" type="submit" class="submit" value="Submit">
                </div>
              </form></td>
            <td></td>
          </tr>
          <tr> 
            <td height="12"></td>
            <td></td>
            <td></td>
            <td></td>
          </tr>
        </table></td>
      <td>&nbsp;</td>
      <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="tables">
          <!--DWLayoutTable-->
          <tr bgcolor="#CCCCCC"> 
            <td height="21" colspan="4" valign="top"><div align="center"><font color="#000000">Add 
                News Item</font></div></td>
          </tr>
          <tr> 
            <td width="4" height="15"></td>
            <td width="69"></td>
            <td width="120"></td>
            <td width="5"></td>
          </tr>
          <tr> 
            <td height="76">&nbsp;</td>
            <td valign="top">Username&nbsp;&nbsp;<br>
              Password</td>
            <td valign="top"><form name="form1" method="post" action="/news/journal/check_user.asp">
                <div align="right"> 
                  <input name="txtUserName" type="text" class="testfields" id="txtUserName">
                  <input name="txtUserPass" type="password" class="testfields" id="txtUserPass">
                  <input name="Submit2" type="submit" class="submit" value="Submit">
                </div>
              </form></td>
            <td></td>
          </tr>
          <tr> 
            <td height="16"></td>
            <td></td>
            <td></td>
            <td></td>
          </tr>
        </table></td>
      <td>&nbsp;</td>
      <td colspan="2" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="tables">
          <!--DWLayoutTable-->
          <tr bgcolor="#CCCCCC"> 
            <td height="21" colspan="4" valign="top"><div align="center"><font color="#000000">Send 
                Out Mailing list</font></div></td>
          </tr>
          <tr> 
            <td width="1" height="15"></td>
            <td width="69"></td>
            <td width="120"></td>
            <td width="8"></td>
          </tr>
          <tr> 
            <td height="78"></td>
            <td valign="top">Username&nbsp;&nbsp;<br>
              Password</td>
            <td valign="top"><form name="form1" method="post" action="/list/mailing_list/check_user.asp">
                <div align="right"> 
                  <input name="txtUserName" type="text" class="testfields" id="txtUserName">
                  <input name="txtUserPass" type="password" class="testfields" id="txtUserPass">
                  <input name="Submit3" type="submit" class="submit" value="Submit">
                </div>
              </form></td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td height="14"></td>
            <td></td>
            <td></td>
            <td></td>
          </tr>
        </table></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="21"></td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <tr> 
      <td height="83"></td>
      <td></td>
      <td colspan="5" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="webtable">
          <!--DWLayoutTable-->
          <tr> 
            <td width="352" height="21" valign="top" bgcolor="#CCCCCC"><div align="center"><font color="#000000">Webstats 
                are updated once a day automatically</font></div></td>
          </tr>
          <tr> 
            <td height="17"></td>
          </tr>
          <tr> 
            <td height="43" valign="top"><form method="POST" action="http://66.102.130.76/cp/awstats/awstats.pl?config=hollow">
                <p align="center"><font face="Arial" size="2"> 
                  <input name="B1" type="submit" class="webstats" value="View WebStats for hollowdays.com">
                  </font></p>
              </form></td>
          </tr>
        </table></td>
      <td>&nbsp;</td>
      <td></td>
    </tr>
    <tr> 
      <td height="20"></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
  </table><!--#include file="../global/footer.shtml" -->

</body>
</html>
