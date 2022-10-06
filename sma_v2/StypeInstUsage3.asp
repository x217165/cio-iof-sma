<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<%
'************************************************************************************************
'* Page name:	STypeInstUsage.asp																*
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
Dim struAttID, objRsSIAtt, objRsuSIAvalue

strWinName = Request.Cookies("WinName")

intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
strRealUserID = Session("username")
struAttID=request("hdnseluSIAtt")

if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access Service Type. Please contact your system administrator"
end if

' For service instance attribute dropdown list
strSQL = "SELECT SRVC_INSTNC_ATT_NAME , " &_
				  "SRVC_INSTNC_ATT_ID " &_
		  "FROM   so.SRVC_INSTNC_ATT " &_
		  "WHERE  RECORD_STATUS_IND = 'A' " &_
		  "ORDER BY SRVC_INSTNC_ATT_NAME"

 'Create Recordset object
'response.write strSQL
'response.end
 set objRsSIAtt = objConn.Execute(strSQL)


 ' For service instance attribute values dropdown list
 strSQL = "SELECT SRVC_INSTNC_ATT_VAL, " &_
				  "SRVC_INSTNC_ATT_VAL_ID " &_
		  "FROM   so.SRVC_INSTNC_ATT_VAL " &_
		  "WHERE  RECORD_STATUS_IND = 'A' "

  if (struAttID <> 0) then
  strSQL = strSQL + " and SRVC_INSTNC_ATT_VAL_ID = 1 "
 '  strSQL = strSQL + " and SRVC_INSTNC_ATT_VAL_ID in	" &_
'		  "( SELECT SRVC_INSTNC_ATT_VAL_ID	" &_
'		  "FROM   SO.SRVC_INSTNC_ATT_VAL_RULE r,  " &_
'		  "SO.SRVC_INSTNC_ATT_VAL_RULE_STAT rs " &_
'		  "WHERE  r.SRVC_INSTNC_ATT_VAL_RULE_ID = rs.SRVC_INSTNC_ATT_VAL_RULE_ID " &_
'		  "AND r.RECORD_STATUS_IND = 'A'	" &_
'		  "AND r.EFF_STOP_TS > SYSDATE " &_
'  		 " AND SRVC_INSTNC_ATT_ID = " & strAttID & ")" &_
  	strSQL = strSQL + " ORDER BY SRVC_INSTNC_ATT_VAL"
 end if
'response.write (strSQL)
'response.end
 set objRsuSIAvalue = objConn.Execute(strSQL)

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
// Function:	fct_onuChange
// Purpose:		set associated values for selected attribute.
// Creaded By:	Linda Chen  July 14th 2009
//**********************************************************************************************
// Set Ref to form
var sSIAttid=document.frmSIAttRM.seluSIAtt
var hselSIAtt=document.frmSIAttRM.hdnseluSIAtt;
// Reset field value
hselSIAtt.value=sSIAttid.value;
var	strURL = 'STypeInstUsage3.asp?hdnseluSIAtt=' + document.frmSIAttRM.hdnseluSIAtt.value ;
self.document.location.href = strURL ;
}


//-->
</SCRIPT>
</HEAD>

<body>
<FORM id="frmSIAttRM" name="frmSIAttRM"  method="POST" action="STtypeAttRpt.asp" target="fraResult" >
	<input id="hdnseluSIAtt" name="hdnseluSIAtt" type=hidden
			value=<%if (struAttID <> 0) then  Response.Write(struAttID) else Response.Write 0 end if%>>

<TABLE>
<thead>
</thead>

<tbody>
<tr>
	<td width="28%">Service Instance Attribute</td>
	<td width="70%">
			<SELECT id=seluSIAtt name=seluSIAtt style="HEIGHT: 22; WIDTH: 272" onchange ="fct_onuChange();">
				<OPTION value=0 ></OPTION>
				<% objRsSIAtt.movefirst
				Do while Not objRsSIAtt.EOF %>
		   		<option  <% if struAttID <> 0 then
		   				if clng(struAttID) = clng(objRsSIAtt(1)) then
		              		response.write "selected "
		              	end if
		              end if %>
           		value = <% =objRsSIAtt(1) %>
		  		 > <% =objRsSIAtt(0)%> </option>
				<%  objRsSIAtt.MoveNext
				Loop %>
				</SELECT>
	</td>
</tr>
<tr>
	<td width="28%" >Service Instance Attribute Value</td>
	<td width="70%" >
		<SELECT id=seluSIAttv name=seluSIAttv style="HEIGHT: 22; WIDTH: 272">
			<OPTION></OPTION>
			<%'
			   Do while Not objRsuSIAvalue.EOF %>
			  <option   value= <% =objRsuSIAvalue(1)%>> <% =objRsuSIAvalue(0) %></option>
				<% objRsuSIAvalue.MoveNext
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
' if strAttID <> 0 then
    objRsSIAtt.close
	objRsuSIAvalue.close
    set objRsSIAtt =	Nothing
	set objRsuSIAvalue = Nothing
 'end if

 objConn.close
 set ObjConn = Nothing


%>


</BODY>
</HTML>