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

Dim intAccessLevel, strRealUserID
Dim strAttvID, strAttID, struAttID
Dim strSQL, strWinName, objRsSIAtt, objRsSIAvalue, objRsuSIAvalue
Dim strAction

strWinName = Request.Cookies("WinName")

intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
strRealUserID = Session("username")
strAttID=request("hdnselSIAtt")
struAttID=request("hdnseluSIAtt")


if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access Service Type. Please contact your system administrator"
end if

' For service instance attribute dropdown list
strSQL = "SELECT SRVC_INSTNC_ATT_NAME , " &_
				  "SRVC_INSTNC_ATT_ID " &_
		  "FROM   so.SRVC_INSTNC_ATT " &_
		  "WHERE  RECORD_STATUS_IND = 'A' " &_
		  "ORDER BY UPPER(SRVC_INSTNC_ATT_NAME)"

 'Create Recordset object
'response.write strSQL
'response.end
 set objRsSIAtt = objConn.Execute(strSQL)


 ' For service instance attribute values dropdown list
 strSQL = "SELECT SRVC_INSTNC_ATT_VAL, " &_
				  "SRVC_INSTNC_ATT_VAL_ID " &_
		  "FROM   so.SRVC_INSTNC_ATT_VAL " &_
		  "WHERE  RECORD_STATUS_IND = 'A' " &_
 		  " ORDER BY UPPER(SRVC_INSTNC_ATT_VAL)"
 set objRsSIAvalue = objConn.Execute(strSQL)

  if (struAttID <> 0) then
   strSQL="SELECT SRVC_INSTNC_ATT_VAL, " &_
		  "SRVC_INSTNC_ATT_VAL_ID " &_
		  "FROM   so.SRVC_INSTNC_ATT_VAL " &_
		  "WHERE  RECORD_STATUS_IND = 'A' " &_
          " and SRVC_INSTNC_ATT_VAL_ID in	" &_
		  "( SELECT SRVC_INSTNC_ATT_VAL_ID	" &_
		  " FROM SO.SRVC_INSTNC_ATT_VAL_RULE r,  "&_
		  " SO.SRVC_INST_ATT_VAL_RULE_STAT rs " &_
   	  	  " WHERE r.srvc_INSTNC_att_id = " & struAttID  &_
		  " AND r.srvc_INSTNC_att_val_rule_id = rs.srvc_INSTNC_att_val_rule_id " &_
		  " AND rs.srvc_INST_att_val_rule_stat_cd = 'A' "   &_
		  " and (rs.eff_stop_ts > sysdate or rs.eff_stop_ts=NULL))" &_
		  " ORDER BY UPPER(SRVC_INSTNC_ATT_VAL)"
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

//function btnUsage_Report()
//{
//document.location.href="STtypeAttRpt.asp?hdnselSIAtt=" + document.frmSIAttRM.selSIAtt.value;
//}

function fct_onChange(){
//**********************************************************************************************
// Function:	fct_onchange()
// Purpose:		set associated values for selected attribute.
// Creaded By:	Linda Chen  July 14th 2009
//**********************************************************************************************
// Set Ref to form
var sSIAttid=document.frmSIAttRM.selSIAtt;
var hselSIAtt=document.frmSIAttRM.hdnselSIAtt;
// Reset field value
hselSIAtt.value=sSIAttid.value;
var	strURL = 'STypeInstUsage.asp?hdnselSIAtt=' + document.frmSIAttRM.hdnselSIAtt.value ;
self.document.location.href = strURL ;
}

function fct_onuChange(){
//**********************************************************************************************
// Function:	fct_onchange()
// Purpose:		set associated values for selected attribute.
// Creaded By:	Linda Chen  July 14th 2009
//**********************************************************************************************
// Set Ref to form
var sSIAttid=document.frmSIAttRM.seluSIAtt
var hselSIAtt=document.frmSIAttRM.hdnseluSIAtt;
// Reset field value
hselSIAtt.value=sSIAttid.value;
var	strURL = 'SInstAttUsage3.asp?hdnseluSIAtt=' + document.frmSIAttRM.hdnseluSIAtt.value ;
self.document.location.href = strURL ;
}


function btnSrch_onclick() {
var v_AttID = document.frmSIAttRM.selmAtt;
if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
	alert('You do not have permission to CREATE a Service Instance Attribute.  Please contact your System Administrator.');
	return false;
}
if (v_AttID.value != 0){
		document.location.href ='SInstAttDetail.asp?hdnInstAttID=' + v_AttID.value;
}
else{
	alert('You Need Select the Attribute to be Searched!');
	return false;
}
}

function btnuSrch_onclick() {
var v_AttID = document.frmSIAttRM.seluSIAtt.value;
var v_AttvID = document.frmSIAttRM.seluSIAttv.value;
if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
	alert('You do not have permission to CREATE a Service Type Attribute.  Please contact your System Administrator.');
	return false;
}

if (v_AttID != 0 && v_AttvID != 0 ){
	document.location.href ='SAttUsageDetail.asp?hdnAttID=' + v_AttID +"&hdnAttvID=" + v_AttvID;
}
else
{
	alert('You Need Select the Attribute and Values to be Searched!');
	return false;
}
}


function btnvSrch_onclick() {
var v_AttId = document.frmSIAttRM.selmAttv.value;
if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
	alert('You do not have permission to CREATE a Service Type Attribute Value.  Please contact your System Administrator.');
	return false;
}
if (v_AttId != 0){
		document.location.href ='SInstAttvDetail.asp?hdnInstAttvID=' + v_AttId;
}
else{
	alert('You Need Select the Attribute Value to be Searched!');
	return false;
}
}



function btnClr_onclick(){
	document.frmSIAttRM.selmAtt.value="";
}

function btnuClr_onclick(){
	document.frmSIAttRM.seluSIAtt.value="";
	document.frmSIAttRM.seluSIAttv.value=""
}

function btnvClr_onclick(){
	document.frmSIAttRM.selmAttv.value="";
}


function btnNew_onclick(){
if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
	alert('You do not have permission to CREATE a Service Type Attribute.  Please contact your System Administrator.');
	return false;
}
document.location.href ="SInstAttDetail.asp?hdnInstAttID=0";
}

function btnvNew_onclick(){
if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
	alert('You do not have permission to CREATE a Service Type Attribute Value.  Please contact your System Administrator.');
	return false;
}
document.location.href ="SInstAttvDetail.asp?hdnInstAttvID=0";
}

function btnuNew_onclick(){
if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
	alert('You do not have permission to CREATE a Service Type Attribute Usage.  Please contact your System Administrator.');
	return false;
}
document.location.href ="SAttUsageDetail.asp?hdnAttID=0&hdnattvID=0";
}


//-->
</SCRIPT>
</HEAD>

<body>
<FORM id="frmSIAttRM" name="frmSIAttRM"  method="POST">
	<input id="hdnselSIAtt" name="hdnselSIAtt" type=hidden
			value=<%if (strAttID <> 0) then  Response.Write(strAttID) else Response.Write 0 end if%>>
	<input id="hdnseluSIAtt" name="hdnseluSIAtt" type=hidden
			value=<%if (struAttID <> 0) then  Response.Write(struAttID) else Response.Write 0 end if%>>

<TABLE>
<thead>
</thead>

<tbody>
<tr>
	<td width="28%">Service Instance Attribute</td>
	<td width="70%">
			<SELECT id=seluSIAtt name=seluSIAtt style="HEIGHT: 22; WIDTH: 600" onchange ="fct_onuChange();">
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
		<SELECT id=seluSIAttv name=seluSIAttv style="HEIGHT: 22; WIDTH: 600">
			<OPTION value=0></OPTION>
			<% if (struAttID <> 0) then
			   	   Do while Not objRsuSIAvalue.EOF %>
			  	   <option   value= <% =objRsuSIAvalue(1)%>> <% =objRsuSIAvalue(0) %></option>
					<% objRsuSIAvalue.MoveNext
			       Loop
			   else
			   	   Do while Not objRsSIAvalue.EOF %>
			  	     <option   value= <% =objRsSIAvalue(1)%>> <% =objRsSIAvalue(0) %></option>
					 <% objRsSIAvalue.MoveNext
			       Loop
			   end if %>
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
    objRsSIAvalue.close
	objRsuSIAvalue.close

    set objRsSIAtt =	Nothing
    set objRsSIAvalue = Nothing
	set objRsuSIAvalue = Nothing
 'end if

 objConn.close
 set ObjConn = Nothing


%>


</BODY>
</HTML>