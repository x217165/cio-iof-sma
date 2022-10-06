<%@ Language=VbScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<TITLE></TITLE>
</HEAD>
<BODY>
<br>
<BR>
<BR>
<table border="0" cellPadding="0" cellSpacing="0" align="center" width="400" style="WIDTH: 400px">
 <tr bgcolor="#ffffff">
   	<td bgcolor="#ffffff" align="middle">
	  <img alt border="0" name="centrestage" src="Images/sma-logo-trans.gif" WIDTH="221" HEIGHT="191"> 
    </td>
 </tr>
</table>
<br>

<%


if instr(1,UCASE(Request.Cookies("UserInformation")("ConnectString")),"ESDD1")then
	Response.Write "<P align=center><FONT size=5 color=red>" & ("SMA2 Development Release Version 2.1.0") & "</FONT></P>"
end if	

if instr(1,UCASE(Request.Cookies("UserInformation")("ConnectString")),"ORAP1")then
	Response.Write "<P align=center><FONT size=5 color=red>" & ("SMA2 Production Release Version 2.1.0") & "</FONT></P>"
end if

if instr(1,UCASE(Request.Cookies("UserInformation")("ConnectString")),"ESDST1")then
	Response.Write "<P align=center><FONT size=5 color=red>" & ("SMA2 System Test Release Version 2.1.0") & "</FONT></P>"
end if

if instr(1,UCASE(Request.Cookies("UserInformation")("ConnectString")),"ESDTR1")then
	Response.Write "<P align=center><FONT size=5 color=red>" & ("SMA2 Training Release Version 2.1.0") & "</FONT></P>"
end if

if instr(1,UCASE(Request.Cookies("UserInformation")("ConnectString")),"ASF1")then
	Response.Write "<P align=center><FONT size=5 color=red>" & ("SMA2 Dev. Release Version 2.1.0") & "</FONT></P>"
end if

%>

<P align=center><FONT size=5 color=red></FONT></P>
<P align=center><FONT size=5 color=red></FONT></P>
<P align=center><FONT size=5 color=red></FONT></P>
<p>
<table border="0" cellPadding="0" cellSpacing="0" align="center" width="400">
 <tr>
   <td align=center><font color=Red size=2>&nbsp;</font></td>
 </tr>
 <tr bgcolor="#ffffff">
   	<td bgcolor="#32cd32" align="middle"> 
    &nbsp;
	</td>
 </tr>
</table></p>


</BODY>
</HTML>
