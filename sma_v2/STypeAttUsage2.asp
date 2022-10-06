<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<%
'************************************************************************************************
'* Page name:	STypeAttUsage2.asp																*
'* Purpose:		To display Service Attribute Value Maintainance Screen							*
'*																								*
'* Created by:					Date															*
'* Linda Chen					07/01/2009														*
'*==============================================================================================*
'* Modifications By				Date				Modifcations								*
'*																								*
'* 																								*
'************************************************************************************************

Dim intAccessLevel, strRealUserID
Dim strSQL, strWinName
Dim objRsSTAvalue

strWinName = Request.Cookies("WinName")

intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
strRealUserID = Session("username")

if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access Service Type. Please contact your system administrator"
end if

' For service attribute values dropdown list
 strSQL = "SELECT SRVC_TYPE_ATT_VAL_NAME, " &_
				  "SRVC_TYPE_ATT_VAL_ID " &_
		  "FROM   CRP.SRVC_TYPE_ATT_VAL " &_
		  "WHERE  RECORD_STATUS_IND = 'A' "  &_
		  " ORDER BY SRVC_TYPE_ATT_VAL_NAME"

 set objRsSTAvalue = objConn.Execute(strSQL)
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

var bolSaveRequired = false;
var intAccessLevel=<%=intAccessLevel%>;
var intConst_MessageDisplay=<%=intConst_MessageDisplay%>;

//function btnUsage_Report()
//{
//document.location.href="STypeAttRpt.asp?hdnselSTAtt=" + document.frmSAttRM.selSTAtt.value;
//}

//-->
</SCRIPT>
</HEAD>

<body>
<FORM id="frmSAttRM" name="frmSAttRM"  method="POST" action="STypeAttRpt.asp" target="fraResult" >
<TABLE>
<tbody>
<thead>
<tr>
</thead>
<tbody>
<tr>
		<td width="80%">Service Type Attribute Value</td>
		<td width="73%">
		<SELECT id=selmAttv name=selmAttv style="HEIGHT: 22; WIDTH: 600">
			<OPTION value=0 ></OPTION>
				<% 'if (objRsSTAvalue0.RecordCount > 0) then
				   Do while Not objRsSTAvalue.EOF %>
		   		   <option  value = <% =objRsSTAvalue(1) %>> <% =objRsSTAvalue(0)%> </option>
				<%  objRsSTAvalue.MoveNext
					Loop
				 ' end if%>
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
 objRsSTAvalue.close
 set objRsSTAvalue = Nothing
 objConn.close
 set ObjConn = Nothing


%>


</BODY>
</HTML>