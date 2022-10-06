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
Dim strSQL, strSQL0, strWinName, objRsSIAtt, objRsSIAvalue, objRsSIAvalue0, objRsuSIAvalue
Dim strAction

strWinName = Request.Cookies("WinName")

intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
strRealUserID =Session("username")
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
 strSQL0 = strSQL + " ORDER BY SRVC_INSTNC_ATT_VAL"

 set objRsSIAvalue0 = objConn.Execute(strSQL0)

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
 set objRsSIAvalue = objConn.Execute(strSQL)

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
var	strURL = 'STypeInstUsage.asp?hdnseluSIAtt=' + document.frmSIAttRM.hdnseluSIAtt.value ;
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
<FORM id="frmSIAttRM" name="frmSIAttRM"  method="POST" action="STtypeAttRpt.asp" target="fraResult" >
	<input id="hdnselSIAtt" name="hdnselSIAtt" type=hidden
			value=<%if (strAttID <> 0) then  Response.Write(strAttID) else Response.Write 0 end if%>>
	<input id="hdnseluSIAtt" name="hdnseluSIAtt" type=hidden
			value=<%if (struAttID <> 0) then  Response.Write(struAttID) else Response.Write 0 end if%>>

<TABLE>
<tr>
   <td>
		<table>
		<thead>
		<tr>
			<td colspan=3 >Service Instance Attribute Maintenance
			</td>
		</tr>
		</thead>

		<tbody>
		<tr>
			<td>Service Attribute</td>
			<td>
			<SELECT id=selmAtt name=selmAtt style="HEIGHT: 22; WIDTH: 281">
				<OPTION value=0 ></OPTION>
				<% objRsSIAtt.Movefirst
				Do while Not objRsSIAtt.EOF %>
		   		<option  value = <% =objRsSIAtt(1) %>
		  		 > <% =objRsSIAtt(0)%> </option>
				<%  objRsSIAtt.MoveNext
				Loop %>
				</SELECT>
			</td>
		</tr>

		</tbody>
		<tfoot>
		<tr>
			<td colspan=2 align=right>
				<input id="btnnewSIAtt" name="btnnewSIAtt" type=button value="New" onclick="btnNew_onclick()">
				<input id="btnclrSIAtt" name="btnclrSIAtt" type=button value="Clear" onclick="btnClr_onclick()">
				<input id="btnsrchSIAtt" name="btnsrchSIAtt" type=button value="Search" onclick="btnSrch_onclick()">
			</td>
			</tr>
		</tfoot>
		</table>
	</td>

	 <td>
		<table>
			<thead>
			<tr>
				<td colspan=3 >Service Instance Attribute Values Maintenance
				</td>
			</tr>
			</thead>

			<tr>
			<td>Attribute Value</td>
			<td>
			<SELECT id=selmAttv name=selmAttv style="HEIGHT: 22; WIDTH: 272">
				<OPTION value=0 ></OPTION>
				<% 'if (objRsSIAvalue0.RecordCount > 0) then
				   Do while Not objRsSIAvalue0.EOF %>
		   		   <option  value = <% =objRsSIAvalue0(1) %>> <% =objRsSIAvalue0(0)%> </option>
				<%  objRsSIAvalue0.MoveNext
					Loop
				 ' end if%>
				</SELECT>
				</td>
			</tr>
			</tbody>
			<tfoot>
			<tr>
				<td colspan=2 align=right>
				<input id="btnnewSIAttv" name="btnnewSIAttv" type=button value="New" onclick="btnvNew_onclick()">
				<input id="btnclrSIAttv" name="btnclrSIAttv" type=button value="Clear" onclick="btnvClr_onclick()">
				<input id="btnsrchSIAttv" name="btnsrchSIAttv" type=button value="Search" onclick="btnvSrch_onclick()">
				</td>
			</tr>
			</tfoot>
		</table>
	</td>
</tr>
</table>

<TABLE>
<tr>
   <td width="50%">
		<table width="107%">
			<thead >
			<tr>
				<td colspan=2 >Service Instance Attribute Usage Report </td>
			</tr>
			</thead>

			<tbody>
			<tr>
			<td>Service Attribute</td>
			<td>
				<SELECT id=selSIAtt name=selSIAtt style="HEIGHT: 22; WIDTH: 272" onchange ="fct_onChange();">
				<OPTION value=0 ></OPTION>
				<% objRsSIAtt.movefirst
				Do while Not objRsSIAtt.EOF %>
		   		<option  <% if strAttID <> 0 then
		   				if clng(strAttID) = clng(objRsSIAtt(1)) then
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
			<td width="31%">Attribute Values</td>
			<td width="65%">
				<SELECT id=selSIAttv name=selSIAttv style="HEIGHT: 22; WIDTH: 272">
				<OPTION></OPTION>
				<%'if objRsSIAvalue.RecordCount > 0 then
				  ' objRsSIAvalue.movefirst
				   Do while Not objRsSIAvalue.EOF %>
				  <option   value= <% =objRsSIAvalue(1)%>> <% =objRsSIAvalue(0) %></option>
				<% objRsSIAvalue.MoveNext
				  Loop
				'end if %>
				</SELECT>
			</td>
			</tr>
			</tbody>
			<tfoot>
			<TR>
	  	 		<td align=center colspan=2>
			  	   <input id=btnrpt name=btnrpt type=submit style="width:229;height:26" value="Attribute Usage Report"
			  	   language=javascript>

			  	 </td>
			</TR>
			</tfoot>
		</table>
	</td>

	 <td width="48%">
		<table>
			<thead>
			<tr>
				<td colspan=3 >Service Instance Attribute Usage Maintenance</td>
			</tr>
			</thead>

			<tbody>
			<tr>
			<td width="28%">Service Attribute</td>
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
			</tr>

			<tr>
			<td width="28%" >Attribute Value</td>
			<td width="70%" >
			<SELECT id=seluSIAttv name=seluSIAttv style="HEIGHT: 22; WIDTH: 272">
				<OPTION></OPTION>
				<%'
				   Do while Not objRsuSIAvalue.EOF %>
				  <option   value= <% =objRsuSIAvalue(1)%>> <% =objRsuSIAvalue(0) %></option>
				<% objRsuSIAvalue.MoveNext
				  Loop
				'end if %>
				</SELECT>			</tr>
			</tbody>
			<tfoot>
			<tr>
				<td colspan=2 align=center>
					<input id="btnuNewSIAtt" name="btnuNewSIAtt" type=button value="New" onclick="btnuNew_onclick()">
					<input id="btnuClrSIAtt" name="btnuClrSIAtt" type=button value="Clear" onclick="btnuClr_onclick()">
					<input id="btnuSrchSIAtt" name="btnuSrchSIAtt" type=button value="Search" onclick="btnuSrch_onclick()">&nbsp;
				</td>
			</tr>
			</tfoot>
		</table>
	</td>
</tr>
</table>

</FORM>
<%

 'Clean up our ADO objects
' if strAttID <> 0 then
    objRsSIAtt.close
    objRsSIAvalue.close
    objRsSIAvalue0.close
	objRsuSIAvalue.close

    set objRsSIAtt =	Nothing
    set objRsSIAvalue = Nothing
    set objRsSIAvalue0 = Nothing
	set objRsuSIAvalue = Nothing
 'end if

 objConn.close
 set ObjConn = Nothing


%>


</BODY>
</HTML>