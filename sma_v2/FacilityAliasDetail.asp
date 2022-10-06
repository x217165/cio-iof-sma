<%@ Language=VBScript %>
<% Option Explicit
 on error resume next
%>
<% Response.Buffer = true %>
<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->
<%


dim intAccessLevel,StrCircuitTyp
StrCircuitTyp = Request("FacType")

IF StrCircuitTyp = "ATMPVC" THEN
intAccessLevel = CInt(CheckLogon(strConst_PVC))
ELSE
 intAccessLevel = CInt(CheckLogon(strConst_Facilities))
END IF

if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to PVC/Facilities. Please contact your system administrator"
end if


dim strRealUserID
strRealUserID = Session("username")

'Response.Write "USER=" & strRealUserID
'Response.Write "TYPE=" & Request("FacType")&"<BR>"
'Response.Write "LEVEL=" & intAccessLevel &"<BR>"
'Response.Write "CHECK=" &intConst_Access_Update &"<BR>"



Dim StrAliasID,StrCircuitID,StrSql,strWhereClause,objRS,strNewFacility,strWinMessage,objRsCktProv,strUpdDate

 StrAliasID = Request("AliasID")
 StrCircuitID = Request("CircuitID")
 strNewFacility = Request("NewFacility")
 strUpdDate = Request("hdnUpdateDateTime")

 if strNewFacility = "NEW" then
  StrAliasID = 0
  StrCircuitID = 0
  strNewFacility =""
 end if

  select case Request("txtFrmAction")
	case "SAVE"
	 if (Request("AliasID") <>"") then
		 if ((intAccessLevel and intConst_Access_Update) <> intConst_Access_Update) then
		   DisplayError "BACK", "", 0, "UPDATE DENIED", "You don't have access to update Facility/PVC Alias. Please contact your system administrator"
		 end if

		   StrAliasID = Request("AliasID")

			'create command object for update stored proc
			dim cmdUpdateObj
			set cmdUpdateObj = server.CreateObject("ADODB.Command")
			set cmdUpdateObj.ActiveConnection = objConn
			cmdUpdateObj.CommandType = adCmdStoredProc
			cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_fac_inter.sp_fac_alias_update"
			'create parameters
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_user_id", adVarChar , adParamInput, 20, strRealUserID)
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_circuit_alias_id",adNumeric , adParamInput,, Clng(Request("AliasID")))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_circuit_id",adNumeric , adParamInput,, Clng(Request("CircuitID")))


			if Request("txtcktalias") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_circuit_alias_name", adVarChar,adParamInput, 50, Request("txtcktalias"))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_circuit_alias_name", adVarChar,adParamInput, 50, null)
			end if

			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , Cdate(Request("hdnUpdateDateTime")))

            if Request("selcktprov") <>"" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_circuit_provider", adVarChar,adParamInput, 6, Request("selcktprov"))
			else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_circuit_provider", adVarChar,adParamInput, 6, null)
			end if

			'Response.Write "updating..." & StrAliasID

			'Response.Write "<b> count = " & cmdInsertObj.Parameters.count & "<br>"

  			'dim nx
  			 'for nx=0 to cmdUpdateObj.Parameters.count-1
  			  ' Response.Write " parm value= " & cmdUpdateObj.Parameters.Item(nx) & "<br>"
  			 ' next

  			  'StrAliasID = Request("AliasID")
  			  'StrCircuitID = Request("CircuitID)

  	   'if objConn.Errors.Count <> 0 then
			'DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE FACILITY/PVC ALIAS - PARAMETER ERROR", objConn.Errors(0).Description
			'objConn.Errors.Clear
		'end if

			cmdUpdateObj.Execute


			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if

			strWinMessage = "Record saved successfully. You can now see the changes you made."
		  else
		   '(Request("AliasID")="" ) and
		   if ((intAccessLevel and intConst_Access_Create) <> intConst_Access_Create) then
		     DisplayError "BACK", "", 0, "INSERT DENIED", "You don't have access to create Facility/PVC. Please contact your system administrator"
		   end if
			'create a new record

			dim cmdInsertObj
			set cmdInsertObj = server.CreateObject("ADODB.Command")
			set cmdInsertObj.ActiveConnection = objConn
			cmdInsertObj.CommandType = adCmdStoredProc
			cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_fac_inter.sp_fac_alias_insert"

			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_user_id", adVarChar , adParamInput, 20, strRealUserID)
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_circuit_alias_id",adNumeric , adParamOutput,,null)
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_circuit_id",adNumeric , adParamInput,, Clng(Request("CircuitID")))


			if Request("txtcktalias") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_circuit_alias_name", adVarChar,adParamInput, 50, Request("txtcktalias"))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_circuit_alias_name", adVarChar,adParamInput, 50, null)
			end if


            if Request("selcktprov") <>"" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_circuit_provider", adVarChar,adParamInput, 6, Request("selcktprov"))
			else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_circuit_provider", adVarChar,adParamInput, 6, null)
			end if

		'if objConn.Errors.Count <> 0 then
			'DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT CREATE FACILITY/PVC ALIAS - PARAMETER ERROR", objConn.Errors(0).Description
			'objConn.Errors.Clear
		'end if

		'dim objparm
  		  ' for each objparm in cmdInsertObj.Parameters
  			  'Response.Write "<b>" & objparm.name & "</b>"
  			  'Response.Write " has size:  " & objparm.Size & " "
  			  'Response.Write " and value:  " & objparm.value & " "
  			  'Response.Write " and datatype:  " & objparm.Type & "<br> "
  		 ' next

		  'dim nx
  			 'for nx=0 to cmdUpdateObj.Parameters.count-1
  			  ' Response.Write " parm value= " & cmdUpdateObj.Parameters.Item(nx) & "<br>"
  			 ' next

			cmdInsertObj.Execute

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT CREATE FACILITY/PVC ALIAS", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
				StrAliasID =  cmdInsertObj.Parameters("p_circuit_alias_id").Value
				'Response.WRITE "ALIAS=" & StrAliasID
				'StrCircuitID = Request("hdncircuitid")
			end if
			strWinMessage = "Record created successfully. You can now see the new record."
		'else
		    'DisplayError "BACK", "", err.Number, "ACCESS DENIED. PLEASE CONTACT YOUR SYSTEM ADMINISTRATOR.", err.Description
	    end if

	    if err then
		  DisplayError "BACK", "", err.Number, "CANNOT CREATE FACILITY - TRY AGAIN", err.Description
	    end if

		case "DELETE"

		'delete record
		if ((intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete) then
		   DisplayError "BACK", "", 0, "DELETE DENIED", "You don't have access to delete Facility/PVC. Please contact your system administrator"
		end if

			dim cmdDeleteObj
			set cmdDeleteObj = server.CreateObject("ADODB.Command")
			set cmdDeleteObj.ActiveConnection = objConn
			'Response.Write Request("hdnUpdateDateTime")
			cmdDeleteObj.CommandType = adCmdStoredProc
			cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_fac_inter.sp_fac_alias_delete"
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_circuit_alias_id", adNumeric, adParamInput,,Clng(StrAliasID))
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(Request("hdnUpdateDateTime")))

			'call the insert stored proc
  			'cmdDeleteObj.Parameters.Refresh

  			'dim objparm
  		   'for each objparm in cmdDeleteObj.Parameters
  			 ' Response.Write "<b>" & objparm.name & "</b>"
  			  'Response.Write " has size:  " & objparm.Size & " "
  			  'Response.Write " and value:  " & objparm.value & " "
  			  'Response.Write " and datatype:  " & objparm.Type & "<br> "
  		  'next

  			'Response.Write "<b> count = " & cmdInsertObj.Parameters.count & "<br>"
  			'dim nx
  			 'for nx=0 to cmdDeleteObj.Parameters.count-1
  			   'Response.Write " parm value= " & cmdDeleteObj.Parameters.Item(nx) & "<br>"
  			' next

			cmdDeleteObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE FACILITY", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if
			StrAliasID = 0
			strWinMessage = "Record deleted successfully."
		  'else
		      'DisplayError "BACK", "", err.Number, "ACCESS DENIED. PLEASE CONTACT YOUR SYSTEM ADMINISTRATOR.", err.Description
	      'end if
       end select



 StrSql = "SELECT CIRCUIT_PROVIDER_CODE FROM CRP.CIRCUIT_PROVIDER WHERE RECORD_STATUS_IND = 'A' ORDER BY CIRCUIT_PROVIDER_CODE"

 'Create Recordset object
 set objRsCktProv = objConn.Execute(StrSql)

 if StrAliasID <> 0 then
   StrSql ="select "&_
         "CIRCUIT_NUMBER_ALIAS_ID," &_
         "CIRCUIT_ID," &_
         "CIRCUIT_NUMBER_ALIAS," &_
         "CIRCUIT_PROVIDER_CODE," &_
         "TO_CHAR(CREATE_DATE_TIME,'MON-DD-YYYY HH:MI:SS') CREATE_DATE_CONV," &_
         "sma_sp_userid.spk_sma_library.sf_get_full_username(CREATE_REAL_USERID) CREATE_REAL_USERID," &_
         "UPDATE_DATE_TIME," &_
         "TO_CHAR(UPDATE_DATE_TIME,'MON-DD-YYYY HH:MI:SS') UPDATE_DATE_CONV," &_
         "RECORD_STATUS_IND," &_
         "sma_sp_userid.spk_sma_library.sf_get_full_username(UPDATE_REAL_USERID) UPDATE_REAL_USERID" &_
         " from crp.CIRCUIT_NUMBER_ALIAS"



      strWhereClause =  "where CIRCUIT_NUMBER_ALIAS_ID =" & StrAliasID

      StrSql =  StrSql & " "& strWhereClause


      set objRS = objConn.Execute(StrSql)


    if err then
	   DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 32132", err.Description
    end if
  end if
   'Response.Write "SQL STATEMENT WIH WHERE=" & StrSql & "<p>"

   'Create the command object



%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK REL="stylesheet" TYPE="text/css" HREF="stylesheets/styles.css">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<TITLE>Alias Detail</TITLE>
<SCRIPT LANGUAGE=javascript>
<!--

var bolSaveRequired = false;
var intAccessLevel=<%=intAccessLevel%>;
var intConst_MessageDisplay=<%=intConst_MessageDisplay%>;




function fct_clearStatus() {
	window.status = "";
}

function fct_displayStatus(strMessage){
	window.status = strMessage;
	setTimeout('fct_clearStatus()',intConst_MessageDisplay);
}

function body_onLoad(strWinStatus){
	var strWinStatus='<%=strWinMessage%>';
	fct_displayStatus(strWinStatus);
}

function body_onUnload(){

	opener.document.fmfacDetail.btn_iFrameRefresh.click();
}



function btnClose_onclick(){
window.close();
parent.opener.iFrame_display();
}

//-->
</SCRIPT>


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function frmFacAlias_onsubmit() {
if	(((intAccessLevel & intConst_Access_Create == intConst_Access_Create) && (document.frmFacAlias.AliasID.value == "")) || ((intAccessLevel & intConst_Access_Update == intConst_Access_Update) && (document.frmFacAlias.AliasID.value != ""))) {
		if (document.frmFacAlias.txtcktalias.value != "")
			{document.frmFacAlias.txtFrmAction.value = "SAVE";
			bolSaveRequired = false;
			return(true);}
		else
			{alert(" Alias Name required");
			document.frmFacAlias.txtcktalias.focus();
			return(false);
			}
	} else {alert('Access denied. Please contact your system administrator.'); return(false);}


}


function btnSave_onclick()
{
var bolretval
bolretval= frmFacAlias_onsubmit();
if(bolretval)
document.frmFacAlias.submit();
}

function fct_onChange(){
if (intAccessLevel >= intConst_Access_Create){
 if (document.frmFacAlias.AliasID.value != "")
     {
      bolSaveRequired = true;
     }

    }
}


function btnNew_click(){
if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}
	self.document.location.href ="FacilityAliasDetail.asp?NewFacility=NEW";
}


function fct_onDelete(){
if (document.frmFacAlias.AliasID.value != '') {
 if (((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) || (document.frmFacAlias.txtRecordStatusInd.value == "D"))
  {alert('Access denied. Please contact your system administrator.');
   return;}
	if (confirm('Do you really want to delete this object?')){
		document.location = "FacilityAliasDetail.asp?txtFrmAction=DELETE&AliasID="+document.frmFacAlias.AliasID.value+"&hdnUpdateDateTime="+document.frmFacAlias.hdnUpdateDateTime.value;
	}
	} //null alias id
	else {
	fct_displayStatus('Unable to Delete record no Alias ID provided.');;
	return(false);
	}
}



function body_onBeforeUnload(){
    document.frmFacAlias.btnSave.focus();
	if (bolSaveRequired) {
		if	(((intAccessLevel & intConst_Access_Create == intConst_Access_Create) && (document.frmFacAlias.txtcktalias.value == "")) || ((intAccessLevel & intConst_Access_Update == intConst_Access_Update) && (document.frmFacAlias.txtcktalias.value != ""))) {
			event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
		}
	}
}

function fct_onReset(){
	bolSaveRequired = false;
	document.location = "FacilityAliasDetail.asp?AliasID=<%=StrAliasID%>";
}


//-->
</SCRIPT>
</HEAD>
<BODY onLoad="body_onLoad();" onBeforeUnload="body_onBeforeUnload();" onUnload="body_onUnload();" >

<FORM name=frmFacAlias LANGUAGE=javascript onsubmit="return frmFacAlias_onsubmit()">
<INPUT type="hidden" name=txtFrmAction value="">
<INPUT name=hdnUpdateDateTime type=hidden style="HEIGHT: 20px; WIDTH: 100px" value=<%if  StrAliasID <> 0 then  Response.Write """"&objRS("UPDATE_DATE_TIME")&"""" else Response.Write """""" end if%>  >&nbsp;&nbsp;
<INPUT  name=AliasID type=hidden style="HEIGHT: 21px; WIDTH: 200px" value= <%if StrAliasID <> 0 then  Response.Write """"&objRS("CIRCUIT_NUMBER_ALIAS_ID")&"""" else Response.Write """""" end if%> >
<INPUT  type=hidden name=CircuitID   style="HEIGHT: 21px; WIDTH: 200px" value= <%if StrAliasID <> 0 then  Response.Write """"&objRS("CIRCUIT_ID")&"""" else Response.Write """"&Request.Cookies("ParentCircuitID")&"""" end if%> >
<TABLE border=0 width=100%>
<thead>
	<TR ><TD colspan=2>Facility/PVC Alias Detail</td></tr>
</thead>
<tbody>
<TR>
	<TD ALIGN=RIGHT width=20% NOWRAP>Alias Name<font color=red>*</font></TD>
	<TD colspan=3 width=80%><INPUT  name=txtcktalias   style="HEIGHT: 21px; WIDTH: 500px" value= <%if StrAliasID <> 0 then  Response.Write """"&routineHtmlString(objRS("CIRCUIT_NUMBER_ALIAS"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();">
</TR>
<TR>
	<TD ALIGN=RIGHT width=20% NOWRAP>Provider Code</TD>
	<TD width=80%>
		<SELECT id=selcktprov name=selcktprov style="HEIGHT: 20px; WIDTH: 120px" onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%Do while Not objRsCktProv.EOF
		 Response.write "<OPTION "
		 if StrAliasID <> 0 then
		    if objRsCktProv("CIRCUIT_PROVIDER_CODE") = objRs("CIRCUIT_PROVIDER_CODE") then
				   Response.Write " selected "
			end if
		 end if
		  Response.Write " VALUE ="& objRsCktProv("CIRCUIT_PROVIDER_CODE") & ">" & objRsCktProv("CIRCUIT_PROVIDER_CODE") & "</OPTION>"
		   objRsCktProv.MoveNext
		 Loop
		%>
		</SELECT>
	</TD>
</TR>
</tbody>
</TABLE>

<TABLE>
	  <TR><TD align=right colspan=5>
			<INPUT id=btnClose name=btnClose  type=button value=Close LANGUAGE=javascript onclick="return btnClose_onclick()">&nbsp;&nbsp;
			<INPUT id=btnReset name=btnReset type=reset value=Reset onClick="fct_onReset();" style="HEIGHT: 24px; WIDTH: 51px">&nbsp;&nbsp;
			<INPUT id=btnAddNew  name=btnAddNew type=button value="New" LANGUAGE=javascript onclick="return btnNew_click()">&nbsp;&nbsp;
			<INPUT id=btnDelete  name=btnDelete type=button value=Delete LANGUAGE=javascript onclick="return fct_onDelete();">&nbsp;&nbsp;
			<INPUT  id=btnSave name=btnSave type=button value=Save onclick="btnSave_onclick();">&nbsp;&nbsp;
	  </TD></TR>
</table>

<FIELDSET >
	<LEGEND ALIGN=RIGHT><B>Audit Information</B></LEGEND>
	<Div SIZE=8pt ALIGN=RIGHT>
		Record Status Indicator:
		<INPUT align = left name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 18px"disabled value=<%if  StrAliasID <> 0 then  Response.Write """"&objRS("RECORD_STATUS_IND")&"""" else Response.Write """""" end if%>  >&nbsp;&nbsp;&nbsp;
		Create Date:&nbsp;&nbsp;
		<INPUT align = center name=txtcrdate type=text style="HEIGHT: 20px; WIDTH: 140px"disabled value=<%if  StrAliasID <> 0 then  Response.Write """"&objRS("CREATE_DATE_CONV")&"""" else Response.Write """""" end if%> >&nbsp;
		&nbsp;
		Created By:
		<INPUT align = right name=txtcrby type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if  StrAliasID <> 0 then  Response.Write """"&routineHtmlString(objRS("CREATE_REAL_USERID"))&"""" else Response.Write """""" end if%> ><BR>
		Update Date:
		<INPUT align= center name=txtupdate type=text style="HEIGHT: 20px; WIDTH: 140px"disabled value=<%if  StrAliasID <> 0 then  Response.Write """"&objRS("UPDATE_DATE_CONV")&"""" else Response.Write """""" end if%>  >&nbsp;&nbsp;
		Updated By:
		<INPUT align=right name=txtupby type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if  StrAliasID <> 0 then  Response.Write """"&routineHtmlString(objRS("UPDATE_REAL_USERID"))&"""" else Response.Write """""" end if%>  >
	</DIV>
</FIELDSET>

</FORM>
<%

 'Clean up our ADO objects
 if StrAliasID <> 0 then
    objRS.close
    set objRS = Nothing
 end if

    objRsCktProv.close
    set objRsCktProv = Nothing

    objConn.close
    set ObjConn = Nothing


%>


</BODY>
</HTML>
