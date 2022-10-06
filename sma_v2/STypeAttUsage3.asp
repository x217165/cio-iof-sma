<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<%
'************************************************************************************************
'* Page name:	STypeAttUsage3.asp																*
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
Dim strSQL
Dim struAttID, objRsSTAtt, objRsuSTAvalue

strWinName = Request.Cookies("WinName")

intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
strRealUserID = Session("username")
struAttID=request("hdnseluSTAtt")

if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access Service Type. Please contact your system administrator"
end if

' For service attribute dropdown list
strSQL = "SELECT SRVC_TYPE_ATT_NAME, " &_
				  "SRVC_TYPE_ATT_ID " &_
		  "FROM   CRP.SRVC_TYPE_ATT " &_
		  "WHERE  RECORD_STATUS_IND = 'A' " &_
		  "ORDER BY SRVC_TYPE_ATT_NAME"

 'Create Recordset object
'response.write strSQL
'response.end
 set objRsSTAtt = objConn.Execute(strSQL)
 strSQL = "SELECT SRVC_TYPE_ATT_VAL_NAME, " &_
				  "SRVC_TYPE_ATT_VAL_ID " &_
		  "FROM   CRP.SRVC_TYPE_ATT_VAL " &_
		  "WHERE  RECORD_STATUS_IND = 'A' "
 if (struAttID <> 0) then
   strSQL = strSQL + " and SRVC_TYPE_ATT_VAL_ID in	" &_
		  "( SELECT SRVC_TYPE_ATT_VAL_ID	" &_
          " FROM crp.SRVC_TYPE_ATT_VAL_RULE r,  "&_
		  " crp.SRVC_TYPE_ATT_VAL_RULE_STAT rs " &_
   	  	  " WHERE r.srvc_type_att_id = " & struAttID  &_
		  " AND r.srvc_type_att_val_rule_id = rs.srvc_type_att_val_rule_id " &_
		  " AND rs.srvc_type_att_val_rule_stat_cd = 'A' "   &_
		  " and (rs.eff_stop_ts > sysdate or rs.eff_stop_ts=NULL))" &_
		  " ORDER BY SRVC_TYPE_ATT_VAL_NAME"
 end if
' response.write (strSQL)
' response.end
 set objRsuSTAvalue = objConn.Execute(strSQL)

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

function fct_onuChange(){
//**********************************************************************************************
// Function:	fct_onchange()
// Purpose:		set associated values for selected attribute.
// Creaded By:	Linda Chen  July 14th 2009
//**********************************************************************************************
// Set Ref to form
var sstattid=document.frmSAttRM.seluSTAtt
var hselSTAtt=document.frmSAttRM.hdnseluSTAtt;
// Reset field value
hselSTAtt.value=sstattid.value;
var	strURL = 'STypeAttUsage3.asp?hdnseluSTAtt=' + document.frmSAttRM.hdnseluSTAtt.value ;
self.document.location.href = strURL ;
}


//-->
</SCRIPT>
</HEAD>

<body>
<FORM id="frmSAttRM" name="frmSAttRM"  method="POST" action="STypeAttRpt.asp" target="fraResult" >
	<input id="hdnseluSTAtt" name="hdnseluSTAtt" type=hidden
			value=<%if (struAttID <> 0) then  Response.Write(struAttID) else Response.Write 0 end if%>>

<TABLE>
<thead>
</thead>
<tbody>
<tr>
	<td width="25%">Service Type Attribute</td>
	<td width="73%">
	<SELECT id=seluSTAtt name=seluSTAtt style="HEIGHT: 22; WIDTH: 600" onchange ="fct_onuChange();">
		<OPTION value=0 ></OPTION>
		<% objRsSTAtt.movefirst
			Do while Not objRsSTAtt.EOF %>
	   		<option  <% if struAttID <> 0 then
	   				if clng(struAttID) = clng(objRsSTAtt(1)) then
	              		response.write "selected "
	              	end if
	              end if %>
         		value = <% =objRsSTAtt(1) %>
	  		 > <% =objRsSTAtt(0)%> </option>
			<%  objRsSTAtt.MoveNext
			Loop %>
	</SELECT>
	</td>
</tr>
<tr>
	<td width="25%" >Service Type Attribute Value</td>
	<td width="73%" >
	<SELECT id=seluSTAttv name=seluSTAttv style="HEIGHT: 22; WIDTH: 600">
			<OPTION value=0></OPTION>
			<%'
			   Do while Not objRsuSTAvalue.EOF %>
			  <option   value= <% =objRsuSTAvalue(1)%>> <% =objRsuSTAvalue(0) %></option>
			<% objRsuSTAvalue.MoveNext
			  Loop
			'end if %>
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
	objRsuSTAvalue.close
    set objRsSTAtt =	Nothing
	set objRsuSTAvalue = Nothing
 objConn.close
 set ObjConn = Nothing
%>


</BODY>
</HTML>