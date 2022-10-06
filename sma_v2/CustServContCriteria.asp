<%@ Language=VBScript %>
<%Option Explicit		%>

<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--****************************************************************************** CSCriteria	* Purpose:		Create a Customer Service Search.*** Created By:	Sara Sangha	28/07/00*****************************************************************************-->

<html>
<head>
	<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<link rel="stylesheet" type="text/css" href="stylesheets/styles.css">
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<title>Customer Service Search </title>
	<script LANGUAGE="javascript" SRC="zoldcheckwhite.js">
	</script>
	<script type="text/javascript">
</script>
</head>
<body>
<form name="frmCSSearch" method="post" action="CustServContList.asp" target="fraResult" onSubmit="return validate(this)">
<table border="1">
    <tr>
        <td>Customer Service Name</td>
        <td><input id="text1" name="text1"> </td>
        <td>Customer Name</td>
        <td><input id="text2" name="text2"></td></tr>
    <tr>
        <td>Customer Service ID</td>
        <td><input id="text3" name="text3"></td>
        <td>Contact Name</td>
        <td><input id="text4" name="text4"></td></tr>
    <tr>
        <td></td>
        <td></td>
        <td>Active Only</td>
        <td><input type="checkbox" id="checkbox1" name="checkbox1"></td></tr>
    <tr>
        <td colSpan="4" align="right">
			<input id="btnSearch" name="btnSearch" type="submit" value="Search" style="HEIGHT: 24px; WIDTH: 62px"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input id="btnReset" name="btnReset" type="reset" value="Reset" style="HEIGHT: 24px; WIDTH: 62px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input id="btnAdd" name="btnAdd" type="button" value="Add New" style="HEIGHT: 24px; WIDTH: 65px" LANGUAGE="javascript" onclick="return btnAdd_onclick()">&nbsp;</td>

        </tr>
    </table>
</form>
</body>
</html>
