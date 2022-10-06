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
'* Modifications By				Date				Modifications								*
'* Anthony Cheung				03/08/2013			Adding Attribute sequencing	(display_order)	*
'* 																								*
'************************************************************************************************

Dim intAccessLevel, strRealUserID
Dim strXRefID, strServiceTypeID, strAttID, strAttvID, strUsageID
Dim strSQL, objRS, strWinMessage, objRsSTAtt, objRsSTAvalue
Dim strhselAttId

intAccessLevel = CInt(CheckLogon(strConst_ServiceCatalogue))
strRealUserID = Session("username")
strXRefID = Request("hdnXRefID")
strServiceTypeID = Request("hdnServiceTypeID")
strAttID = Request("hdnstrAttID")
strAttvID = Request("hdnstrAttvID")
strUsageID = Request("hdnUsageID")
strhselAttID = Request("hdnselAttID")


if intAccessLevel and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access Service Type. Please contact your system administrator"
end if

if (strXRefID <> 0) then
    strSQL = "SELECT RECORD_STATUS_IND, " &_
		" CREATE_DATE_TIME, CREATE_REAL_USERID, " &_
		" UPDATE_DATE_TIME, UPDATE_REAL_USERID " &_
		" FROM CRP.SRVC_TYPE_ATT_VAL_XREF " &_
		" WHERE SRVC_TYPE_ATT_VAL_XREF_ID = " & strXRefID
	set objRs = objConn.Execute(strSQL)
	if err then
		DisplayError "BACK", "", err.Number, "ERROR IN SELECTING SERVICE TYPE USAGE INFORMATION", err.Description
	end if
end if


Select case Request("txtFrmAction")
response.write "action is " & equest("txtFrmAction")
response.end

	case "SAVE"

	 if (Request.Form("hdnXRefID") <> 0) then

		'The XRefID is not null i.e. it is an existing record. So call the update procedure to update the record
		 if ((intAccessLevel and intConst_Access_Update) <> intConst_Access_Update) then
		   DisplayError "BACK", "", 0, "UPDATE DENIED", "You don't have access to update this record. Please contact your system administrator"
		 end if

		    strXRefID = Request.Form("hdnXRefID")

			dim cmdUpdateObj
			set cmdUpdateObj = server.CreateObject("ADODB.Command")
			set cmdUpdateObj.ActiveConnection = objConn
			cmdUpdateObj.CommandType = adCmdStoredProc
			cmdUpdateObj.CommandText = "sma_sp_userid.Sp_Srvtype_Att_Val_Xref_Update"

			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_user_id", adVarChar , adParamInput, 20, strRealUserID)
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_servicetype_Att_Val_xref_id",adNumeric , adParamInput,, clng(strXRefID))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_srvc_type_att_id",adNumeric , adParamInput,, Clng(Request("selSTAtt")))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_srvc_type_att_val_id",adNumeric , adParamInput,, Clng(Request("selSTAttv")))


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
  			'response.write (cmdUpdateObj.CommandText)
			'response.end

			cmdUpdateObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE RECORD", objConn.Errors(0).Description
				objConn.Errors.Clear
				response.redirect("STypeAttDetail.asp")
			else
				 '20150511LC  strXRefID =  cmdUpdateObj.Parameters("p_servicetype_Att_Val_xref_id").Value
				 '20150511LC  strAttID =  cmdUpdateObj.Parameters("p_srvc_type_att_id").Value
				 '20150511LC  strAttvID =  cmdUpdateObj.Parameters("p_srvc_type_att_id").Value
                  strWinMessage = "Record Updated successfully. You can now see the changes you made."
                  response.write strWinMessage
                  response.end
        		 '20150511LC  response.write("<script language=""javascript"">window.close();parent.opener.iSTAFrame_display();</script>")

			end if

	else 'create a new record

		   if ((intAccessLevel and intConst_Access_Create) <> intConst_Access_Create) then
		     DisplayError "BACK", "", 0, "INSERT DENIED", "You don't have access to add Default Service Type Attribute. Please contact your system administrator"
		   end if

			dim cmdInsertObj
			set cmdInsertObj = server.CreateObject("ADODB.Command")
			set cmdInsertObj.ActiveConnection = objConn
			cmdInsertObj.CommandType = adCmdStoredProc
			cmdInsertObj.CommandText = "sma_sp_userid.Sp_Srvtype_Att_Val_Xref_Insert"

			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_user_id", adVarChar , adParamInput, 20, strRealUserID)
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_servicetype_Att_Val_xref_id",adNumeric, adParamOutput,,null)
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_service_type_id",adNumeric , adParamInput,, Clng(Request("hdnServiceTypeID")))
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_display_order", adNumeric , adParamInput,, null)
 			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_srvc_type_att_id",adNumeric , adParamInput,, Clng(Request("selSTAtt")))
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_srvc_type_att_val_id",adNumeric , adParamInput,, Clng(Request("selSTAttv")))

			'****************************
			'check parameter values
  			'****************************

  			'dim objparm
  			'for each objparm in cmdInsertObj.Parameters
  			'	  Response.Write "<b>" & objparm.name & "</b>"
  			'	  Response.Write " has size:  " & objparm.Size & " "
  			'	  Response.Write " and value:  " & objparm.value & " "
  			'	  Response.Write " and datatype:  " & objparm.type & "<br> "
  			'next

  			'Response.Write "<b> count = " & cmdInsertObj.Parameters.count & "<br>"
  			'dim nx
  			'for nx=0 to cmdInsertObj.Parameters.count-1
  			'   Response.Write nx+1 & " parm value= " & cmdInsertObj.Parameters.Item(nx).Value  & "<br>"
  			'next

  			'response.write (cmdInsertObj.CommandText)
			'response.end

			cmdInsertObj.Execute

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT ADD NEW RECORD", objConn.Errors(0).Description
				objConn.Errors.Clear
				response.redirect("STypeAttDetail.asp")
			else
				 '20150511LC   strXRefID =  cmdInsertObj.Parameters("p_servicetype_Att_Val_xref_id").Value
				 '20150511LC  strAttID =  cmdInsertObj.Parameters("p_srvc_type_att_id").Value
				 '20150511LC  strAttvID =  cmdInsertObj.Parameters("p_srvc_type_att_val_id").Value
 				 response.write("<script language=""javascript"">window.close();parent.opener.iSTAFrame_display();</script>")

			end if
			strWinMessage = "Record created successfully. You can now see the new record."
	end if
    'response.write("<script language=""javascript"">window.close();parent.opener.iSTAFrame_display();</script>")     'lc added on 20150511

 end select

 strSQL = "SELECT SRVC_TYPE_ATT_NAME, " &_
				  "SRVC_TYPE_ATT_ID " &_
		  "FROM   CRP.SRVC_TYPE_ATT " &_
		  "WHERE  RECORD_STATUS_IND = 'A' " &_
		  "ORDER BY upper(SRVC_TYPE_ATT_NAME)"

 'Create Recordset object
 set objRsSTAtt = objConn.Execute(strSQL)

 strSQL = "SELECT v.SRVC_TYPE_ATT_VAL_NAME ,	" &_
	  "v.SRVC_TYPE_ATT_VAL_ID	" &_
	  "FROM   CRP.SRVC_TYPE_ATT_VAL	v  "
 if (strhselAttID <> 0 OR strAttID <> 0) then
 	strSQL = strSQL + ", crp.SRVC_TYPE_ATT_VAL_RULE r, " &_
		  "crp.srvc_type_att_val_rule_stat rs " &_
		  "WHERE  v.record_status_ind = 'A' "&_
		  "AND v.SRVC_TYPE_ATT_VAL_ID=r.SRVC_TYPE_ATT_VAL_ID " &_
		  "AND r.srvc_type_att_val_rule_id = rs.srvc_type_att_val_rule_id " &_
		  "AND rs.srvc_type_att_val_rule_stat_cd ='A' " &_
 		  "AND (rs.eff_stop_ts>sysdate or rs.eff_stop_ts = NULL) "
 	if ( strhselAttID <> 0 ) then
 	   strSQL= StrSQL & " AND r.SRVC_TYPE_ATT_ID = " & strhselAttID
        else
    	   strSQL= StrSQL & " AND r.SRVC_TYPE_ATT_ID = " & strAttID
	end if
 end if
 strSQL = strSQL & " ORDER BY upper(SRVC_TYPE_ATT_VAL_NAME)"
 response.write(strSQL)
 'response.end
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
<TITLE>Default Service Type Attribute</TITLE>
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

	'20150511LC replace it with line below opener.document.frmSTypeDetail.btn_iFrameRefresh.click();
	 opener.document.frmSTypeDetail.btn_iSTAFrameRefresh.click();

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
	parent.opener.iSTAFrame_display();

}

function frmSAttDetail_onsubmit() {
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		set the frmAction to SAVE if the user has access to save the record.
//
// Creaded By:	Sara Sangha		Feb. 15th, 2001
//**********************************************************************************************

if	(((intAccessLevel & intConst_Access_Create == intConst_Access_Create) && (document.frmSAttDetail.hdnXRefID.value == ""))
		|| ((intAccessLevel & intConst_Access_Update == intConst_Access_Update) && (document.frmSAttDetail.hdnXRefID.value != ""))) {

			document.frmSAttDetail.txtFrmAction.value = "SAVE";
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
alert('Save button is clicked.');

	bolretval= frmSAttDetail_onsubmit();

	if(bolretval)
		document.frmSAttDetail.submit();

//	window.close();
//	parent.opener.iSTAFrame_display();

}

function fct_onChange(){
//**********************************************************************************************
// Function:	btnClose_onclick()
//
// Purpose:		set the bolSaveRequired flag if anything changes on the screen.
//
// Creaded By:	Sara Sangha		Feb. 15th, 2001
//**********************************************************************************************
var v_selAtt = document.frmSAttDetail.selSTAtt;
var v_hdnAtt = document.frmSAttDetail.hdnselAttID;
v_hdnAtt.value = v_selAtt.value;
var strURL
    strURL = 'STypeAttDetail.asp?hdnServiceTypeID=' + document.frmSAttDetail.hdnServiceTypeID.value;
	strURL = strURL + '&hdnselAttID=' + document.frmSAttDetail.hdnselAttID.value;
	strURL = strURL + '&hdnXRefID=' + document.frmSAttDetail.hdnXRefID.value;
	strURL = strURL + '&hdnstrAttID=' + document.frmSAttDetail.hdnstrAttID.value;
	strURL = strURL + '&hdnstrAttvID=' + document.frmSAttDetail.hdnstrAttvID.value;
	strURL = strURL + '&hdnUsageID=' + document.frmSAttDetail.hdnUsageID.value;
//if (intAccessLevel >= intConst_Access_Create){
//if (document.frmSAttDetail.hdnXRefID.value != "")
//     {
//     bolSaveRequired = true;
//     }

//    }
 self.document.location.href=strURL;
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

	strURL = 'STypeAttDetail.asp?XRefID=0&ServiceTypeID=' + document.frmSAttDetail.hdnServiceTypeID.value ;
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

	if (document.frmSAttDetail.hdnXRefID.value != "") {
		if (((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) || (document.frmSAttDetail.txtRecordStatusInd.value == "D")){
			alert('Access denied. Please contact your system administrator.');
			return;
		}

		if (confirm('Do you really want to delete this object?')){

			strURL = 'STypeAttDetail.asp?txtFrmAction=DELETE&XRefID='
					+ document.frmSAttDetail.hdnXRefID.value + '&UpdateDateTime='
					+ document.frmSAttDetail.hdnUpdateDateTime.value + '&ServiceTypeID='
					+ document.frmSAttDetail.hdnServiceTypeID.value;

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

    document.frmSAttDetail.btnSave.focus();
	if (bolSaveRequired) {
		if	(((intAccessLevel & intConst_Access_Create == intConst_Access_Create) && (document.frmSAttDetail.txtcktalias.value == "")) || ((intAccessLevel & intConst_Access_Update == intConst_Access_Update) && (document.frmSAttDetail.txtcktalias.value != ""))) {
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
	//When reset screen for Update
    if (document.frmSAttDetail.hdnXRefID.value != "")  {
	    document.frmSAttDetail.selSTAtt.value = document.frmSAttDetail.hdnstrAttID.value;
		document.frmSAttDetail.selSTAttv.value = document.frmSAttDetail.hdnstrAttvID.value;
	}
	//When reset screen for New
	else {

		document.frmSAttDetail.selSTAttv.value="";
		document.frmSAttDetail.selSTAtt.value="";
	}
}


//-->
</SCRIPT>
</HEAD>

<BODY onLoad="body_onLoad();" onBeforeUnload="body_onBeforeUnload();" onUnload="body_onUnload();" >
<FORM  name=frmSAttDetail action="STypeAttDetail.asp" method="POST" onsubmit="return frmSAttDetail_onsubmit()">
	<INPUT  name=txtFrmAction type=hidden value="" >
	<input name=txtcktalias type=hidden value="">
	<INPUT  name=hdnXRefID  type=hidden value= <%if strXRefID <> 0 then  Response.Write strXRefID else Response.Write 0 end if%> >
	<INPUT  name=hdnServiceTypeID type=hidden  value=<%if strServiceTypeId <> 0 then  Response.Write strServiceTypeId else Response.Write 0 end if%> >
 	<INPUT  name=hdnstrAttID type=hidden value= <%if strAttID <> 0 then  Response.Write strAttID else Response.Write 0  end if%> >
	<INPUT  name=hdnstrAttvID type=hidden  value= <%if strAttvID <> 0 then  Response.Write strAttvID else Response.Write 0 end if%> >
	<INPUT  name=hdnUsageID type=hidden value= <%if strUsageID <> 0 then  Response.Write strUsageID else Response.Write 0 end if%> >
    <INPUT  name=hdnselAttID type=hidden value= <%if strhselAttID <> 0 then  Response.Write strhselAttID  else Response.Write 0 end if%> >




<TABLE>
<thead>
	<TR ><TD colspan=2>Default Attribute</td></tr>
</thead>

<tbody>
<TR>
	<TD ALIGN=RIGHT width=20% NOWRAP>Attribute Type<font color=red>*</font></TD>
	<TD width=95%>
		<SELECT id=selSTAtt name=selSTAtt style="HEIGHT: 20px; WIDTH: 580px" onchange ="fct_onChange();">
		<OPTION ></OPTION>
		<%Do while Not objRsSTAtt.EOF %>
		   <option  <% if (strhselAttID <> 0) then
		   				 if clng(strhselAttID) = clng(objRsSTAtt(1)) then
		              		response.write "selected "
		              	 end if
		               else
		                 if (strAttID <> 0) then
		                    if clng(strAttID) = clng(objRsSTAtt(1)) then
		                		response.write "selected "
		              	    end if
		              	 end if
		              end if %>
           value = <% =objRsSTAtt(1) %>
		   > <% =objRsSTAtt(0)%> </option>
		<%  objRsSTAtt.MoveNext
		Loop
		%>

		</SELECT>
	</TD>
</TR>
<TR>
	<TD ALIGN=RIGHT width=20% NOWRAP>Attribute Value<font color=red>*</font></TD>
	<TD width=95%>
		<SELECT id=selSTAttv name=selSTAttv style="HEIGHT: 20px; WIDTH: 580px">
		<OPTION></OPTION>
		<%Do while Not objRsSTAvalue.EOF %>
		 <option <% if strXRefID <> 0 then
		               if clng(strAttvID) = clng(objRsSTAvalue(1)) then
		           			response.write "selected "
		           	   end if
		           end if %>
		  value= <% =objRsSTAvalue(1)%>
		 > <% =objRsSTAvalue(0) %></option>
		<%
		 objRsSTAvalue.MoveNext
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
 objRsSTAtt.close
 set objrsSTAtt = Nothing
 objRsSTAvalue.close
 set objRsSTAvalue = Nothing

 '   objRsCktProv.close
  '  set objRsCktProv = Nothing

    objConn.close
    set ObjConn = Nothing


%>


</BODY>
</HTML>
