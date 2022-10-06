<%@ Language=VBScript %>
<%  Option Explicit   %>
<!-- #include file="smaConstants.inc" -->
<!-- #include file="smaProcs.inc" -->
<%
CheckLogon(strConst_Logon)
%>

<html>
<head>
<title>Service Management Application</title>

<script type="text/javascript">
<!--
window.moveTo(0,0);
window.resizeTo(screen.availWidth,screen.availHeight);

var StrCookiesEnabled = false;


// Purpose:		Check whether cookies enabled
// Created By:	Ian Harroitt

function frasetMain_onload()
{
document.cookie = "Enabled=true";
var StrcookieValid = document.cookie;
//if we are able to retrieve the value just set, then cookies are enabled
 if (StrcookieValid.indexOf("Enabled=true") != -1)
 {
  StrCookiesEnabled = true;
  }
  else
  {
   StrCookiesEnabled = false;
   alert("You need to enable cookies on your browser to use the SMA Application");
   }
}
//-->
</script>

<title>header</title>
</head>

<frameset frameborder="0" framespacing="0" border="0" cols="*" rows="85,*" onLoad="frasetMain_onload()">
  <frame marginwidth="0" marginheight="0" src="heading.htm" name="heading" noresize scrolling="no">
  <frame marginwidth="5" marginheight="5" src="SMA%20Version.htm" name="text" noresize>
<noframes>
<p>IE5 is either missing or misconfigured. Please contact your system administrator.</p>

</noframes>
</frameset>
</html>
