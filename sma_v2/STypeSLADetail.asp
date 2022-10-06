<%@ Language=VBScript %>
<% Option Explicit
 on error resume next
%>
<% Response.Buffer = true %>
<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->
<%
'************************************************************************************************
'* Page name:	STypeDetail.asp																	*
'* Purpose:		To display the Service Type														*
'*				Chosen via STypeList.asp														*
'*																								*
'* Created by:					Date															*
'* Sara Sangha					02/15/2000														*
'*==============================================================================================*
'* Modifications By				Date				Modifcations								*
'*																								*
'* 																								*
'************************************************************************************************

Dim intAccessLevel, strRealUserID
Dim strXRefID, strServiceTypeID, strRegionLcode, strSLAID, strUpdateDateTime
Dim strSQL, strWhereClause, objRS, strWinMessage, objRsRegionLcode, objRsSLADesc

intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
strRealUserID = Session("username")
strXRefID = Request("XRefID")
strServiceTypeID = Request("ServiceTypeID")
strUpdateDateTime = Request("UpdateDateTime")


if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access Service Type. Please contact your system administrator"
end if


Select case Request("txtFrmAction")

	case "SAVE"

	 if (Request.Form("hdnXRefID") <> "") then

		'The XRefID is not null i.e. it is an existing record. So call the update procedure to update the record
		 if ((intAccessLevel and intConst_Access_Update) <> intConst_Access_Update) then
		   DisplayError "BACK", "", 0, "UPDATE DENIED", "You don't have access to update this record. Please contact your system administrator"
		 end if

		    strXRefID = Request.Form("hdnXRefID")

			dim cmdUpdateObj
			set cmdUpdateObj = server.CreateObject("ADODB.Command")
			set cmdUpdateObj.ActiveConnection = objConn
			cmdUpdateObj.CommandType = adCmdStoredProc
			cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_servtype_region_xref_update"

			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_user_id", adVarChar , adParamInput, 20, strRealUserID)
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_servicetype_region_xref_id",adNumeric , adParamInput,, clng(strXRefID))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_service_type_id",adNumeric , adParamInput,, Clng(Request("hdnServiceTypeID")))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_region_lcode",adNumeric , adParamInput,, Clng(Request("selRegionLcode")))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_sla_id",adNumeric , adParamInput,, Clng(Request("selSLA")))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_last_update_dt", adDBTimeStamp  , adParamInput,, Cdate(Request("hdnUpdateDateTime")))


			'****************************
			'check parameter values
  			'****************************

  			'dim objparm
  			'for each objparm in cmdUpdateObj.Parameters
  			'	  Response.Write "<b>" & objparm.name & "</b>"
  			'	  Response.Write " has size:  " & objparm.Size & " "
  			'	  Response.Write " and value:  " & objparm.value & " "
  			'	  Response.Write " and datatype:  " & objparm.Type & "<br> "
  			'next

  			'Response.Write "<b> count = " & cmdUpdateObj.Parameters.count & "<br>"
  			'dim nx
  			'for nx=0 to cmdUpdateObj.Parameters.count-1
  			'   Response.Write nx+1 & " parm value= " & cmdUpdateObj.Parameters.Item(nx).Value  & "<br>"
  			'next

			cmdUpdateObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE RECORD", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
				strWinMessage = "Record saved successfully. You can now see the changes you made."
			end if


	else 'create a new record

		   if ((intAccessLevel and intConst_Access_Create) <> intConst_Access_Create) then
		     DisplayError "BACK", "", 0, "INSERT DENIED", "You don't have access to add Default SLA. Please contact your system administrator"
		   end if

			dim cmdInsertObj
			set cmdInsertObj = server.CreateObject("ADODB.Command")
			set cmdInsertObj.ActiveConnection = objConn
			cmdInsertObj.CommandType = adCmdStoredProc
			cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_servtype_region_xref_insert"


			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_user_id", adVarChar , adParamInput, 20, strRealUserID)
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_servicetype_region_xref_id",adNumeric , adParamOutput,,null)
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_service_type_id",adNumeric , adParamInput,, Clng(Request("hdnServiceTypeID")))
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_region_lcode",adNumeric , adParamInput,, Clng(Request("selRegionLcode")))
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_sla_id",adNumeric , adParamInput,, Clng(Request("selSLA")))



			'****************************
			'check parameter values
  			'****************************

  			'dim objparm
  			'for each objparm in cmdInsertObj.Parameters
  			'	  Response.Write "<b>" & objparm.name & "</b>"
  			'	  Response.Write " has size:  " & objparm.Size & " "
  			'	  Response.Write " and value:  " & objparm.value & " "
  			'	  Response.Write " and datatype:  " & objparm.Type & "<br> "
  			'next

  			'Response.Write "<b> count = " & cmdInsertObj.Parameters.count & "<br>"
  			'dim nx
  			'for nx=0 to cmdInsertObj.Parameters.count-1
  			'   Response.Write nx+1 & " parm value= " & cmdInsertObj.Parameters.Item(nx).Value  & "<br>"
  			'next



			cmdInsertObj.Execute

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT ADD NEW RECORD", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
				 strXRefID =  cmdInsertObj.Parameters("p_servicetype_region_xref_id").Value
			end if
			strWinMessage = "Record created successfully. You can now see the new record."

	end if

		case "DELETE" 'delete record

		if ((intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete) then
		   DisplayError "BACK", "", 0, "DELETE DENIED", "You don't have access to delete Facility/PVC. Please contact your system administrator"
		end if

			dim cmdDeleteObj
			set cmdDeleteObj = server.CreateObject("ADODB.Command")
			set cmdDeleteObj.ActiveConnection = objConn
			cmdDeleteObj.CommandType = adCmdStoredProc

			cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_servtype_region_xref_delete"
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_servtype_region_xref_id", adNumeric, adParamInput,,Clng(strXRefID))
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(strUpdateDateTime))

			cmdDeleteObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE FACILITY", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
				strXRefID = 0
				strWinMessage = "Record deleted successfully."
			end if
		 end select

 strSQL = "SELECT REVENUE_REGION_LCODE, " &_
				  "REVENUE_REGION_DESC " &_
		  "FROM   SO.LCODE_REVENUE_REGION " &_
		  "WHERE  RECORD_STATUS_IND = 'A' " &_
		  "ORDER BY REVENUE_REGION_DESC"

 'Create Recordset object
 set objRsRegionLcode = objConn.Execute(strSQL)


 strSQL = "SELECT SERVICE_LEVEL_AGREEMENT_ID,	" &_
				  "SERVICE_LEVEL_AGREEMENT_DESC	" &_
		  "FROM   CRP.SERVICE_LEVEL_AGREEMENT	" &_
		  "WHERE  RECORD_STATUS_IND = 'A'	" &_
		  "ORDER BY SERVICE_LEVEL_AGREEMENT_DESC "

 set objRsSLADesc = objConn.Execute(strSQL)

 if (strXRefID <> 0 and strXRefID <> "") then

   strSQL =  "SELECT " &_
		     "  x.servicetype_region_xref_id," &_
			 "  s.service_type_id, " &_
			 "  x.region_lcode, " &_
			 "  x.service_level_agreement_id, " &_
			 "  r.revenue_region_desc, " &_
			 "  a.service_level_agreement_desc, " &_
			 "  TO_CHAR(X.CREATE_DATE_TIME, 'MON-DD-YYYY HH:MI:SS') CREATE_DATE_TIME, " &_
		     "  sma_sp_userid.spk_sma_library.sf_get_full_username(X.CREATE_REAL_USERID) CREATE_REAL_USERID, " &_
		     "  TO_CHAR(X.UPDATE_DATE_TIME,'MON-DD-YYYY HH:MI:SS') UPDATE_DATE_TIME, " &_
		     "  X.RECORD_STATUS_IND, " &_
		     "  sma_sp_userid.spk_sma_library.sf_get_full_username(X.UPDATE_REAL_USERID) UPDATE_REAL_USERID, " &_
		     "  x.update_date_time as update_date_time2 " &_
		" from   crp.servicetype_region_xref x, " &_
			 "  crp.service_type s, " &_
			 "  so.lcode_revenue_region r, " &_
			 "  crp.service_level_agreement a " &_
		" where  x.service_type_id = s.service_type_id " &_
		" and	x.REGION_LCODE = r.revenue_region_lcode " &_
		" and	x.service_level_agreement_id = a.service_level_agreement_id " &_
		" and	x.servicetype_region_xref_id = " & strXRefID &_
		" order by r.revenue_region_desc "


    'Response.Write (strSQL)
    'Response.End

    set objRS = objConn.Execute(StrSql)
    if err then
	   DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 32132", err.Description
    end if

  end if

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK REL="stylesheet" TYPE="text/css" HREF="stylesheets/styles.css">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<TITLE>Default SLA</TITLE>
<SCRIPT LANGUAGE=javascript>
<!--

var bolSaveRequired = false;
var intAccessLevel=<%=intAccessLevel%>;
var intConst_MessageDisplay=<%=intConst_MessageDisplay%>;


function fct_clearStatus() {
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		clear the message from window status bar.
//
// Creaded By:	Ian Harriott
//**********************************************************************************************
	window.status = "" ;
}

function fct_displayStatus(strMessage){
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		display a message in window status bar and then clear it after the set minutes.
//
// Creaded By:	Ian Harriott
//**********************************************************************************************
	window.status = strMessage;
	setTimeout('fct_clearStatus()',intConst_MessageDisplay);
}

function body_onLoad(strWinStatus){
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		Whenever the page is loaded it displays a message in window status bar.
//
// Creaded By:	Sara Sangha		Feb. 15th, 2001
//**********************************************************************************************
	var strWinStatus='<%=strWinMessage%>';
	fct_displayStatus(strWinStatus);
}

function body_onUnload(){
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		Refresh contents of iFrame in the base window.
//
// Creaded By:	Sara Sangha		Feb. 15th, 2001
//**********************************************************************************************

	opener.document.frmSTypeDetail.btn_iFrameRefresh.click();
}


function btnClose_onclick(){
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		close the pop up window and Refresh the contents of iFrame in the base window.
//
// Creaded By:	Sara Sangha		Feb. 15th, 2001
//**********************************************************************************************

	window.close();
	parent.opener.iFrame_display();

}

function frmSLADetail_onsubmit() {
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		set the frmAction to SAVE if the user has access to save the record.
//
// Creaded By:	Sara Sangha		Feb. 15th, 2001
//**********************************************************************************************

if	(((intAccessLevel & intConst_Access_Create == intConst_Access_Create) && (document.frmSLADetail.hdnXRefID.value == ""))
		|| ((intAccessLevel & intConst_Access_Update == intConst_Access_Update) && (document.frmSLADetail.hdnXRefID.value != ""))) {

			document.frmSLADetail.txtFrmAction.value = "SAVE";
			bolSaveRequired = false;
			return(true);
		}
   else {
		alert('Access denied. Please contact your system administrator.');
		return(false);
	}
}


function btnSave_onclick() {
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		close the pop up window and Refresh the iFrame in the base window.
//
// Creaded By:	Ian Harriott		Feb. 15th, 2001
//**********************************************************************************************
var bolretval

	bolretval= frmSLADetail_onsubmit();

	if(bolretval)
		document.frmSLADetail.submit();
}

function fct_onChange(){
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		set the bolSaveRequired flag if anything changes on the screen.
//
// Creaded By:	Sara Sangha		Feb. 15th, 2001
//**********************************************************************************************
if (intAccessLevel >= intConst_Access_Create){
if (document.frmSLADetail.hdnXRefID.value != "")
     {
     bolSaveRequired = true;
     }

    }
}


function btnNew_click(){
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		if the user has access to add new records then submit the page to itself with
//				XRefID = 0 so that it will display a blank page.
//
// Creaded By:	Sara Sangha		Feb. 15th, 2001
//**********************************************************************************************
var strURL ;

	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('Access denied. Please contact your system administrator.');
	    return;
	}

	strURL = 'STypeSLADetail.asp?XRefID=0&ServiceTypeID=' + document.frmSLADetail.hdnServiceTypeID.value ;
	self.document.location.href = strURL ;
}


function fct_onDelete(){
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		if the user has access to delete a record then set frmAction = 'DELETE' and pass in
//				in the required parameterst to delete a record
//
// Creaded By:	Sara Sangha		Feb. 15th, 2001
//**********************************************************************************************
var strURL ;

	if (document.frmSLADetail.hdnXRefID.value != "") {
		if (((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) || (document.frmSLADetail.txtRecordStatusInd.value == "D")){
			alert('Access denied. Please contact your system administrator.');
			return;
		}

		if (confirm('Do you really want to delete this object?')){

			strURL = 'STypeSLADetail.asp?txtFrmAction=DELETE&XRefID='
					+ document.frmSLADetail.hdnXRefID.value + '&UpdateDateTime='
					+ document.frmSLADetail.hdnUpdateDateTime.value + '&ServiceTypeID='
					+ document.frmSLADetail.hdnServiceTypeID.value;

			document.location = strURL ;
		}

	else {
		fct_displayStatus('Unable to Delete the record. No Record ID provided.');
		return(false);
	}
  }
}


function body_onBeforeUnload(){
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		Give a warrening message is there is unsaved data on the screen.
//
// Creaded By:	Sara Sangha		Feb. 15th, 2001
//**********************************************************************************************

    document.frmSLADetail.btnSave.focus();
	if (bolSaveRequired) {
		if	(((intAccessLevel & intConst_Access_Create == intConst_Access_Create) && (document.frmSLADetail.txtcktalias.value == "")) || ((intAccessLevel & intConst_Access_Update == intConst_Access_Update) && (document.frmSLADetail.txtcktalias.value != ""))) {
			event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
		}
	}
}


function fct_onReset(){
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		Refresh the contents on the screen from databaase.
//
// Creaded By:	Sara Sangha		Feb. 15th, 2001
//**********************************************************************************************
	bolSaveRequired = false;
	if (document.frmSLADetail.hdnXRefID.value != "")  {
		document.location = 'STypeSLADetail.asp?XRefID=' + document.frmSLADetail.hdnXRefID.value ;
	}
}


//-->
</SCRIPT>
</HEAD>

<BODY onLoad="body_onLoad();" onBeforeUnload="body_onBeforeUnload();" onUnload="body_onUnload();" >
<FORM  name=frmSLADetail action="STypeSLADetail.asp" method="POST" onsubmit="return frmSLADetail_onsubmit()">

	<INPUT type=hidden name=txtFrmAction      value="" >
	<INPUT type=hidden name=hdnUpdateDateTime value= <%if strXRefID <> 0 then  Response.Write """"&objRS("update_date_time2")&"""" else Response.Write """""" end if%> >
	<INPUT type=hidden name=hdnXRefID         value= <%if strXRefID <> 0 then  Response.Write strXRefID else Response.Write """""" end if%> >
	<INPUT type=hidden name=hdnServiceTypeID  value= <%if strXRefID <> 0 then  Response.Write """"&objRS("service_type_id")&"""" else Response.Write strServiceTypeID %>>
	<INPUT type=hidden name=hdnRegionLcode    value= <%if strXRefID <> 0 then  Response.Write """"&objRS("region_lcode")&""""  end if%> >
	<INPUT type=hidden name=hdnSLAID          value= <%if strXRefID <> 0 then  Response.Write """"&objRS("service_level_agreement_id")&""""  end if%> >

<TABLE>
<thead>
	<TR ><TD colspan=2>Default SLA</td></tr>
</thead>

<tbody>
<TR>
	<TD ALIGN=RIGHT width=20% NOWRAP>Revenue Region<font color=red>*</font></TD>
	<TD width=80%>
		<SELECT id=selRegionLcode name=selRegionLcode style="HEIGHT: 20px; WIDTH: 400px" onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%Do while Not objRsRegionLcode.EOF
		 Response.write "<OPTION "
		 if strXRefID <> 0 then
		    if clng(objRsRegionLcode("REVENUE_REGION_LCODE")) = clng(objRs("REGION_LCODE")) then
				   Response.Write " selected "
			end if
		 end if
		   Response.Write " VALUE = "& objRsRegionLcode("REVENUE_REGION_LCODE") & ">" & objRsRegionLcode("REVENUE_REGION_DESC") & "</OPTION>"
		   objRsRegionLcode.MoveNext
		 Loop
		%>
		</SELECT>
</TR>
<TR>
	<TD ALIGN=RIGHT width=20% NOWRAP>Default SLA<font color=red>*</font></TD>
	<TD width=80%>
		<SELECT id=selSLA name=selSLA style="HEIGHT: 20px; WIDTH: 400px" onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%Do while Not objRsSLADesc.EOF
		 Response.write "<OPTION "
		 if strXRefID <> 0 then
		    if clng(objRsSLADesc("service_level_agreement_id")) = clng(objRs("service_level_agreement_id")) then
				   Response.Write " selected "
			end if
		 end if
		   Response.Write " VALUE = "& objRsSLADesc("service_level_agreement_id") & ">" & objRsSLADesc("service_level_agreement_desc") & "</OPTION>"
		   objRsSLADesc.MoveNext
		 Loop
		%>
		</SELECT>
	</TD>
</TR>
</tbody>
</TABLE>

<TABLE>
	  <TR><TD align=right>
			<INPUT id=btnClose   name=btnClose  type=button style="width:2cm" value=Close  LANGUAGE=javascript onclick="return btnClose_onclick()"> &nbsp;&nbsp;
			<INPUT id=btnReset   name=btnReset  type=button style="width:2cm" value=Reset  LANGUAGE=javascript onClick="return fct_onReset();" >           &nbsp;&nbsp;
			<INPUT id=btnAddNew  name=btnAddNew type=button style="width:2cm" value=New    LANGUAGE=javascript onclick="return btnNew_click();">     &nbsp;&nbsp;
			<INPUT id=btnDelete  name=btnDelete type=button style="width:2cm" value=Delete LANGUAGE=javascript onclick="return fct_onDelete();">    &nbsp;&nbsp;
			<INPUT id=btnSave    name=btnSave   type=button style="width:2cm" value=Save   LANGUAGE=javascript onclick="return btnSave_onclick();">        &nbsp;&nbsp;
	  </TD></TR>
</table>

<FIELDSET >
	<LEGEND ALIGN=RIGHT><B>Audit Information</B></LEGEND>
	<Div SIZE=8pt ALIGN=RIGHT>
		Record Status Indicator:
		<INPUT align = left name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 18px"disabled value=<%if  strXRefID <> 0 then  Response.Write """"&objRS("RECORD_STATUS_IND")&"""" else Response.Write """""" end if%>  >&nbsp;&nbsp;&nbsp;
		Create Date:&nbsp;&nbsp;
		<INPUT align = center name=txtcrdate type=text style="HEIGHT: 20px; WIDTH: 140px"disabled value=<%if  strXRefID <> 0 then  Response.Write """"&objRS("CREATE_DATE_TIME")&"""" else Response.Write """""" end if%> >&nbsp;
		&nbsp;
		Created By:
		<INPUT align = right name=txtcrby type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if  strXRefID <> 0 then  Response.Write """"&routineHtmlString(objRS("CREATE_REAL_USERID"))&"""" else Response.Write """""" end if%> ><BR>
		Update Date:
		<INPUT align= center name=txtupdate type=text style="HEIGHT: 20px; WIDTH: 140px"disabled value=<%if  strXRefID <> 0 then  Response.Write """"&objRS("UPDATE_DATE_TIME")&"""" else Response.Write """""" end if%>  >&nbsp;&nbsp;
		Updated By:
		<INPUT align=right name=txtupby type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if  strXRefID <> 0 then  Response.Write """"&routineHtmlString(objRS("UPDATE_REAL_USERID"))&"""" else Response.Write """""" end if%>  >
	</DIV>
</FIELDSET>

</FORM>
<%

 'Clean up our ADO objects
 if strXRefID <> 0 then
    objRS.close
    set objRS = Nothing
 end if

 '   objRsCktProv.close
  '  set objRsCktProv = Nothing

    objConn.close
    set ObjConn = Nothing


%>


</BODY>
</HTML>
