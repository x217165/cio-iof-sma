<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<%
'************************************************************************************************
'* Page name:	STypeAttUsage1.asp																*
'* Purpose:		To display Service Attribute/Values Maintainance Screen							*
'*																								*
'* Created by:					Date															*
'* Linda Chen					07/01/2009														*
'*==============================================================================================*
'* Modifications By				Date				Modifcations								*
'*																								*
'* 																								*
'************************************************************************************************

Dim intAccessLevel, strRealUserID, strWinName
Dim strSQL, objRsSTAtt

strWinName = Request.Cookies("WinName")
intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
strRealUserID = Session("username")

if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access Service Type. Please contact your system administrator"
end if

' For service attribute dropdown list
strSQL = "SELECT SRVC_TYPE_ATT_NAME, " &_
				  "SRVC_TYPE_ATT_ID " &_
		  "FROM   CRP.SRVC_TYPE_ATT " &_
		  "WHERE  RECORD_STATUS_IND = 'A' " &_
		  "ORDER BY SRVC_TYPE_ATT_NAME"
 set objRsSTAtt = objConn.Execute(strSQL)
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK REL="stylesheet" TYPE="text/css" HREF="stylesheets/styles.css">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<TITLE>Service Attribute</TITLE>
<SCRIPT LANGUAGE=javascript>
<!--

//-->
</SCRIPT>
</HEAD>
<body>
<FORM id="frmSAttRM" name="frmSAttRM"  method="POST" action="STypeAttUsage1.asp" target="fraResult" >
	<INPUT type="hidden" id="hdnFrmAction" name="hdnFrmAction" value="">
<TABLE>
	<thead>
	</thead>
	<tbody>
	<tr>
		<td width="80%">Service Type Attribute</td>
		<td width="73%">
		<SELECT id=selmAtt name=selmAtt style="HEIGHT: 22; WIDTH: 600">
			<OPTION value=0 ></OPTION>
			<% objRsSTAtt.Movefirst
			Do while Not objRsSTAtt.EOF %>
	   		<option  value = <% =objRsSTAtt(1) %>
	  		 > <% =objRsSTAtt(0)%> </option>
				<%  objRsSTAtt.MoveNext
				Loop %>
		</SELECT>
		</td>
	</tr>
	</tbody>
	<tfoot>
	</tfoot>
</table>

</FORM>
<%
 'Clean up our ADO objects
 objRsSTAtt.close
 set objRsSTAtt = Nothing
 objConn.close
 set ObjConn = Nothing
%>
</BODY>
</HTML>