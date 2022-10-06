<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<%
'************************************************************************************************
'* Page name:	STypeAttUsage.asp																*
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
Dim strSQL, strSQL0, strWinName, objRsSTAtt, objRsSTAvalue, objRsSTAvalue0, objRsuSTAvalue
'Dim strAction

strWinName = Request.Cookies("WinName")

intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
strRealUserID = Session("username")
strAttID=request("hdnselSTAtt")
struAttID=request("hdnseluSTAtt")

'strAction=request("hdnaction")

'response.write ("strTypeID is " + strAttvID)
'response.end

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


 ' For service attribute values dropdown list
 strSQL = "SELECT SRVC_TYPE_ATT_VAL_NAME, " &_
				  "SRVC_TYPE_ATT_VAL_ID " &_
		  "FROM   CRP.SRVC_TYPE_ATT_VAL " &_
		  "WHERE  RECORD_STATUS_IND = 'A' "
 strSQL0 = strSQL + " ORDER BY SRVC_TYPE_ATT_VAL_NAME"

 set objRsSTAvalue0 = objConn.Execute(strSQL0)

 if (strAttID <> 0) then
   strSQL = strSQL + " and SRVC_TYPE_ATT_VAL_ID in	" &_
		  "( SELECT SRVC_TYPE_ATT_VAL_ID	" &_
		  "FROM   CRP.SRVC_TYPE_ATT_VAL_USAGE	v  " &_
		  "WHERE  RECORD_STATUS_IND = 'A'	" &_
   		 " AND SRVC_TYPE_ATT_ID = " & strAttID & ")" &_
   		 " ORDER BY SRVC_TYPE_ATT_VAL_NAME"
 end if
 'response.write (strSQL)
 'response.end
 set objRsSTAvalue = objConn.Execute(strSQL)

  if (struAttID <> 0) then
   strSQL = strSQL + " and SRVC_TYPE_ATT_VAL_ID in	" &_
		  "( SELECT SRVC_TYPE_ATT_VAL_ID	" &_
		  "FROM   CRP.SRVC_TYPE_ATT_VAL_USAGE	v  " &_
		  "WHERE  RECORD_STATUS_IND = 'A'	" &_
   		 " AND SRVC_TYPE_ATT_ID = " & struAttID & ")" &_
   		 " ORDER BY SRVC_TYPE_ATT_VAL_NAME"
 end if
 'response.write (strSQL)
 'response.end
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

//function btnUsage_Report()
//{
//document.location.href="STypeAttRpt.asp?hdnselSTAtt=" + document.frmSAttRM.selSTAtt.value;
//}

function fct_onChange(){
//**********************************************************************************************
// Function:	fct_onchange()
// Purpose:		set associated values for selected attribute.
// Creaded By:	Linda Chen  July 14th 2009
//**********************************************************************************************
// Set Ref to form
var sstattid=document.frmSAttRM.selSTAtt;
var hselSTAtt=document.frmSAttRM.hdnselSTAtt;
// Reset field value
hselSTAtt.value=sstattid.value;
var	strURL = 'STypeAttUsage.asp?hdnselSTAtt=' + document.frmSAttRM.hdnselSTAtt.value ;
self.document.location.href = strURL ;
}

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
var	strURL = 'STypeAttUsage.asp?hdnseluSTAtt=' + document.frmSAttRM.hdnseluSTAtt.value ;
self.document.location.href = strURL ;
}

function btnclose_onclick() {
	document.location.href='STypeACriteria.asp';
}

function btnSrch_onclick() {
var v_AttID = document.frmSAttRM.selmAtt;
if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
	alert('You do not have permission to CREATE a Service Type Attribute.  Please contact your System Administrator.');
	return false;
}
if (v_AttID.value != 0){
		document.location.href ='SAttDetail.asp?hdnAttID=' + v_AttID.value;
}
else{
	alert('You Need Select the Attribute to be Searched!');
	return false;
}
}

function btnuSrch_onclick() {
var v_AttID = document.frmSAttRM.seluSTAtt.value;
var v_AttvID = document.frmSAttRM.seluSTAttv.value;
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
var v_AttId = document.frmSAttRM.selmAttv.value;
if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
	alert('You do not have permission to CREATE a Service Type Attribute Value.  Please contact your System Administrator.');
	return false;
}
if (v_AttId != 0){
		document.location.href ='SAttvDetail.asp?hdnAttvID=' + v_AttId;
}
else{
	alert('You Need Select the Attribute Value to be Searched!');
	return false;
}
}



function btnClr_onclick(){
	document.frmSAttRM.selmAtt.value="";
}

function btnuClr_onclick(){
	document.frmSAttRM.seluSTAtt.value="";
	document.frmSAttRM.seluSTAttv.value=""
}

function btnvClr_onclick(){
	document.frmSAttRM.selmAttv.value="";
}


function btnNew_onclick(){
if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
	alert('You do not have permission to CREATE a Service Type Attribute.  Please contact your System Administrator.');
	return false;
}
document.location.href ="SAttDetail.asp?hdnAttID=0";
}

function btnvNew_onclick(){
if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
	alert('You do not have permission to CREATE a Service Type Attribute Value.  Please contact your System Administrator.');
	return false;
}
document.location.href ="SAttvDetail.asp?hdnAttvID=0";
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
<FORM id="frmSAttRM" name="frmSAttRM"  method="POST" action="STypeAttRpt.asp" target="fraResult" >
	<input id="hdnselSTAtt" name="hdnselSTAtt" type=hidden
			value=<%if (strAttID <> 0) then  Response.Write(strAttID) else Response.Write 0 end if%>>
	<input id="hdnseluSTAtt" name="hdnseluSTAtt" type=hidden
			value=<%if (struAttID <> 0) then  Response.Write(struAttID) else Response.Write 0 end if%>>

<TABLE>
<tbody>
	<tr>
   	<td>
   	'T1 starts
	<table width="91%">
	<thead>
	<tr>
		<td colspan=3>Service Type Attribute Maintenance</td>
	</tr>
	</thead>
	<tbody>
	<tr>
		<td width="26%">Service Type Attribute</td>
		<td width="73%">
		<SELECT id=selmAtt name=selmAtt style="HEIGHT: 22; WIDTH: 272">
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
	<tr>
		<td colspan=2 align=center>
		<input id="btnnewSTAtt" name="btnnewSTAtt" type=button value="New" onclick="btnNew_onclick()">
		<input id="btnclrSTAtt" name="btnclrSTAtt" type=button value="Clear" onclick="btnClr_onclick()">
		<input id="btnsrchSTAtt" name="btnsrchSTAtt" type=button value="Update" onclick="btnSrch_onclick()">
		</td>
	</tr>
	</tfoot>
	</table>
	</td>
	</tr>
	<tr>
	<td>

	'T1 ends/T2 starts/
	<table width="76%">
	<thead>
	<tr>
		<td colspan=3 >Service Attribute Values Maintenance
		</td>
	</tr>
	</thead>
	<tbody>
	<tr>
		<td width="26%">Service Type Attribute Value</td>
		<td width="73%">
		<SELECT id=selmAttv name=selmAttv style="HEIGHT: 22; WIDTH: 272">
			<OPTION value=0 ></OPTION>
				<% 'if (objRsSTAvalue0.RecordCount > 0) then
				   Do while Not objRsSTAvalue0.EOF %>
		   		   <option  value = <% =objRsSTAvalue0(1) %>> <% =objRsSTAvalue0(0)%> </option>
				<%  objRsSTAvalue0.MoveNext
					Loop
				 ' end if%>
				</SELECT>
		</td>
	</tr>
	</tbody>
	<tfoot>
	<tr>
		<td colspan=2 align=center>
		<input id="btnnewSTAttv" name="btnnewSTAttv" type=button value="New" onclick="btnvNew_onclick()">
		<input id="btnclrSTAttv" name="btnclrSTAttv" type=button value="Clear" onclick="btnvClr_onclick()">
		<input id="btnsrchSTAttv" name="btnsrchSTAttv" type=button value="Update" onclick="btnvSrch_onclick()">
		</td>
	</tr>
	</tfoot>
	</table>
	</td>
	</tr>
	<tr>
	<td>
	't2 ends /t3 starts

	<TABLE>
	<tr>
    <td width="50%">
	<thead>
	<tr>
		<td colspan=3 >Service Type Attribute Usage Maintenance</td>
	</tr>
	</thead>
	<tbody>
	<tr>
		<td width="25%">Service Type Attribute</td>
		<td width="73%">
		<SELECT id=seluSTAtt name=seluSTAtt style="HEIGHT: 22; WIDTH: 272" onchange ="fct_onuChange();">
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
		<SELECT id=seluSTAttv name=seluSTAttv style="HEIGHT: 22; WIDTH: 272">
				<OPTION></OPTION>
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
	<tr>
		<td colspan=2 align=center>
			<input id="btnuNewSTAtt" name="btnuNewSTAtt" type=button value="New" onclick="btnuNew_onclick()">
			<input id="btnuClrSTAtt" name="btnuClrSTAtt" type=button value="Clear" onclick="btnuClr_onclick()">
			<input id="btnuSrchSTAtt" name="btnuSrchSTAtt" type=button value="Update" onclick="btnuSrch_onclick()">&nbsp;
		</td>
	</tr>
	</tfoot>
	</table>
	</td>
	</tr>
	<tr>
	<td>

	't3 ends/t4 starts
	<table width="107%">
	<thead >
	<tr>
		<td colspan=2>Service Type Attribute Usage Report </td>
	</tr>
	</thead>
	<tbody>
	<tr>
		<td>Service Attribute</td>
		<td>
			<SELECT id=selSTAtt name=selSTAtt style="HEIGHT: 22; WIDTH: 272" onchange ="fct_onChange();">
			<OPTION value=0 ></OPTION>
			<% objRsSTAtt.movefirst
				Do while Not objRsSTAtt.EOF %>
		   		<option  <% if strAttID <> 0 then
		   				if clng(strAttID) = clng(objRsSTAtt(1)) then
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
		<td width="31%">Attribute Values</td>
		<td width="65%">
				<SELECT id=selSTAttv name=selSTAttv style="HEIGHT: 22; WIDTH: 272">
				<OPTION></OPTION>
				<%'if objRsSTAvalue.RecordCount > 0 then
				  ' objRsSTAvalue.movefirst
				   Do while Not objRsSTAvalue.EOF %>
				  <option   value= <% =objRsSTAvalue(1)%>> <% =objRsSTAvalue(0) %></option>
				<% objRsSTAvalue.MoveNext
				  Loop
				'end if %>
				</SELECT>
		</td>
	</tr>
	</tbody>
	<tfoot>
	<TR>
		<td align=center colspan=2>
	  	   <input id=btnrpt name=btnrpt type=submit style="width:113;height:26" value="Search"
	  	   language=javascript>
	 	</td>
	</TR>
	</tfoot>
	</table>
</tbody>
<tfoot>
			<TR>
	  	 		<td align=right colspan=2>
				<INPUT id=btnClose name=btnClose type=button  style="width: 2cm" value="Close"
					language="javascript" onClick="btnClose_onClick()" >
			  	 </td>
			</TR>
</tfoot>
</table>

</FORM>
<%

 'Clean up our ADO objects
' if strAttID <> 0 then
    objRsSTAtt.close
    objRsSTAvalue.close
    objRsSTAvalue0.close
	objRsuSTAvalue.close

    set objRsSTAtt =	Nothing
    set objRsSTAvalue = Nothing
    set objRsSTAvalue0 = Nothing
	set objRsuSTAvalue = Nothing
 'end if

 objConn.close
 set ObjConn = Nothing


%>


</BODY>
</HTML>