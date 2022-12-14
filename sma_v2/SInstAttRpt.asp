<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
*****************************************************************************************
* File Name:	STypeAttRpt.asp
* Author:		Linda chen
* Purpoase:		Display Service Attribute Report and Exportable to an Excel File
* Date:			August 2009
******************************************************************************************
-->

<%
Dim intAccessLevel, strRealUserID
Dim strAttID, strSql, objUsageRs, strExport
DIM strAttvID


strExport = Request("hdnExport")
intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
strAttID = Request("hdnselSInst")
strAttvID = Request("hdnselSInstv")
strSql ="select at.srvc_instnc_att_name att_name, av.srvc_instnc_att_val att_val " &_
		"from so.srvc_instnc_att at, " &_
     	"so.srvc_instnc_att_val av, " &_
     	"so.srvc_instnc_att_val_rule ar, "&_
     	"so.srvc_inst_att_val_rule_stat ars " &_
		"where at.srvc_instnc_att_id = ar.srvc_instnc_att_id "&_
		"and  av.srvc_instnc_att_val_id = ar.srvc_instnc_att_val_id " &_
		"and ar.srvc_instnc_att_val_rule_id = ars.srvc_instnc_att_val_rule_id " &_
		"and ars.srvc_inst_att_val_rule_stat_cd = 'A' " &_
		"and (ars.eff_stop_ts > sysdate or ars.eff_stop_ts = NULL)"

if (strAttID <> 0) then
	strSql = strSql  & "and ar.srvc_instnc_att_id = " & strAttID
end if
if (strAttvID <> 0) then
	strSql = strSql  & "and ar.srvc_instnc_att_val_id = " & strAttvID
end if

strSql = strSql & " order by att_name, att_val"

'response.write strSql
'response.end

set objUsageRs = objConn.Execute(strSql)
%>

<HTML>
<head>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK REL="stylesheet" TYPE="text/css" HREF="stylesheets/styles.css">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<TITLE>Service Attribute(s)</TITLE>
<SCRIPT LANGUAGE=javascript>
</SCRIPT>

</head>

<body>
<form id=frmSTAttRpt name=frmSTAttRpt method=post action="SInstAttRpt.asp" >
	<INPUT id="hdnExport" name="hdnExport" type=hidden  value="">
	<input id="hdnselSInst" name="hdnselSInst" type=hidden
	value=<%if (strAttID <> "") then  Response.Write(strAttID) else Response.Write """""" end if%>>
	<input id="hdnselSInstv" name="hdnselSInstv" type=hidden
	value=<%if (strAttvID <> "") then  Response.Write(strAttvID) else Response.Write """""" end if%>>


<% if strExport = "" then %>
  <IMG src="images/excel.gif" onClick="target='new';hdnExport.value='xls';submit(); hdnExport.value=''; target='_self';" width="32" height="32">
<% end if
if strExport <> "" then
'response.write(StrSql)
'response.end

	Response.Clear()
	Response.Buffer = True
	Response.AddHeader "Content-Disposition", "attachment;filename=AttReport.xls"
	Response.ContentType = "application/vnd.ms-excel"
end if %>

<table border=1 >
	<thead>
	<tr>
		<td> Service Instance Attribute(s)</td>
		<td> Attribute Value(s)</td>
	</tr>
	</thead>
	<tbody>
	<% Do while Not objUsageRs.EOF %>
	<tr>
		<td><% =objUsageRs(0) %></td>
		<td><% =objUsageRs(1) %></td>
	</tr>
	<% objUsageRs.MoveNext
	Loop %>
	</tbody>
</table>
</form>
<%
    objUsageRS.close
    set objUsageRS = Nothing
    objConn.close
    set ObjConn = Nothing
%>

</body>
</html>