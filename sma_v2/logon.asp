<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<title>SMA - Logon</title>

<body>

<table border="0" cellPadding="0" cellSpacing="0" align="center" width="400" style="WIDTH: 400px">
 <tr bgcolor="#ffffff">
   	<td bgcolor="#ffffff" align="middle">
	  <img alt border="0" name="centrestage" src="Images/sma-logo-trans.gif" WIDTH="221" HEIGHT="191"> 
    </td>
 </tr>
 <tr> 
    <td colspan="3" align="middle"><strong><font size="6">Service Management Administration<br></font></strong></td>
 </tr>
</table>

<form method="post" action="smalogin.asp" name="frmUserLogon" id="frmUserLogon">
<table align="center" border="0" cellPadding="1" cellSpacing="1" style="WIDTH: 400px" width="400">
    <tr>
        <td></td>
        <td style="WIDTH: 30%" width="30%">
            <div align="right" ><strong>User ID:</strong></div></td>
        <td style="WIDTH: 50%" width="50%"> 
             <input name="txtusername" size="30" maxlength="40" value="" id="txtusername" style="HEIGHT: 22px; WIDTH: 202px"></td>
        <td></td></tr>
    <tr>
        <td></td>
        <td>
            <div align="right" ><strong>Password:</strong></div></td>
        <td>
            <input id="txtpassword" maxLength="40" name="txtpassword" style="HEIGHT: 22px; WIDTH: 201px" type="password" value=""></td>
        <td></td></tr>
    <tr>
        <td></td>
        <td colSpan="2">
            <div align="center"><br>
              <input name="btnSubmit" type="submit" value="Logon"> 
        <td></td></tr></table>
</form>
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
