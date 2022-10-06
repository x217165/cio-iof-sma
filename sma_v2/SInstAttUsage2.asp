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
Dim strSQL, objRsSInstAttv

strWinName = Request.Cookies("WinName")
intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
strRealUserID = Session("username")

if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access Service Type. Please contact your system administrator"
end if

' For service attribute dropdown list
strSQL = "SELECT SRVC_INSTNC_ATT_VAL, " &_
				  "SRVC_INSTNC_ATT_VAL_ID " &_
		  "FROM   SO.SRVC_INSTNC_ATT_VAL " &_
		  "WHERE  RECORD_STATUS_IND = 'A' " &_
		  "ORDER BY upper(SRVC_INSTNC_ATT_VAL) "

 set objRsSInstAttv = objConn.Execute(strSQL)
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
<FORM id="frmInstAttRM" name="frmInstAttRM"  method="POST" action="SInstAttUsage1.asp" target="fraResult" >
	<INPUT type="hidden" id="hdnFrmAction" name="hdnFrmAction" value="">
<TABLE>
	<thead>
	</thead>
	<tbody>
	<tr>
		<td width="32%">Service Instance Attribute Value</td>
		<td width="67%">
		<SELECT id=selmAttv name=selmAttv style="HEIGHT: 22; WIDTH: 600">
			<OPTION value=0 ></OPTION>
			<% objRsSInstAttv.Movefirst
			Do while Not objRsSInstAttv.EOF %>
	   		<option  value = <% =objRsSInstAttv(1) %>
	  		 > <% =objRsSInstAttv(0)%> </option>
				<%  objRsSInstAttv.MoveNext
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
 objRsSInstAttv.close
 set objRsSInstAttv = Nothing
 objConn.close
 set ObjConn = Nothing
%>
</BODY>
</HTML>