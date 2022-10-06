<%@ Language=VBScript %>
<% Option Explicit %>
<% on error resume next %>
<% Response.Buffer = true %>
<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->

<!-- This is the child detail screen for the service location screen.
     It can have the following values passed into it:

     	Parameter			Details
     	------------------------------------------------------------------------
    	CustServContactID	the ID from the database of the service location contact
    						required for updates and deletes

    	CustServID			the ID of the service location from the parent screen
    						required for creates

    	CustName			the Customer's name form the parent screen. This is only used to select an appropriate customer
    						required for creates

    	NewContact			must have a value of 'NEW'
	    					required for new records

    	hdnUpdateDateTime	the updateDateTime from the database
							required for updates and deletes

********************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       22-Jan-01	 DTy		Increase contact priority from 10 to 30.
       19-Feb-02	 DTy		Provide extra space for email-address which had increased
                                  from 50 to 60 characters.
********************************************************************************************
-->


<%
Const ASP_NAME = "CustServContDetail.asp"
Const NO_ID    = "null"

dim intAccessLevel
Dim aRole

intAccessLevel = CInt(CheckLogon(strConst_CustomerServiceContact))
if intAccessLevel < intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to Service Location Contacts. Please contact your system administrator."
end if


Dim strRealUserID
strRealUserID = Session("username")

Dim strCustServContactID, strCustServID, strCustomerName, strNewContact, strWinMessage, strUpdDate
Dim strSql
Dim objRS, objRSContactRole, objCmd
Dim strContactInfo

strCustServContactID = Request("CustServContactID")
strCustServID = Request("CustServID")

strCustomerName = Request("txtWorkFor")
strNewContact = Request("NewContact")
strUpdDate = Request("hdnUpdateDateTime")


if strNewContact = "NEW" then
	strCustServContactID = NO_ID
end if

if strCustServContactID = "" then
	strCustServContactID = NO_ID
end if

dim nx
select case Request("hdnFrmAction")
	case "SAVE"
		if (strCustServContactID <> NO_ID) then
			if (intAccessLevel and intConst_Access_Update) <> intConst_Access_Update then
				DisplayError "BACK", "", 0, "UPDATE DENIED", "You don't have access to update Service Location Contacts. Please contact your system administrator."
			end if

			dim cmdUpdateObj
			set cmdUpdateObj = server.CreateObject("ADODB.Command")
			set cmdUpdateObj.ActiveConnection = objConn
			cmdUpdateObj.CommandType = adCmdStoredProc
			cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_cserv_inter.sp_cs_cont_update"

			aRole = split(Request("selRole"),"¿")

			'create params
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_userid", adVarChar, adParamInput, 20, strRealUserID)
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_customer_service_contact_id", adNumeric, adParamInput, , Clng(Request("CustServContactID")))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_cust_serv_contact_type_lcode", adVarChar, adParamInput, 8, aRole(0))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_customer_service_id", adNumeric, adParamInput, , Clng(Request("CustServID")))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_contact_id", adNumeric, adParamInput, , Clng(Request("hdnContactID")))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_contact_priority", adNumeric, adParamInput, , Clng(Request("selPriority")))
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_last_update", adDBTimeStamp, adParamInput, , CDate(Request("hdnUpdateDateTime")))

			'call the stored proc
			cmdUpdateObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if

			strCustServContactID = Request("CustServContactID")

			strWinMessage = "Record saved successfully. You can now see the changes you made."

		else
			'create a new record
			if (intAccessLevel and intConst_Access_Create) <> intConst_Access_Create then
				DisplayError "BACK", "", 0, "CREATE DENIED", "You don't have access to create Service Location Contacts. Please contact your system administrator."
			end if

			dim cmdInsertObj
			set cmdInsertObj = server.CreateObject("ADODB.Command")
			set cmdInsertObj.ActiveConnection = objConn
			cmdInsertObj.CommandType = adCmdStoredProc
			cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_cserv_inter.sp_cs_cont_insert"

			aRole = split(Request("selRole"),"¿")

			'create params
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_userid", adVarChar, adParamInput, 20, strRealUserID)
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_customer_service_contact_id", adNumeric, adParamOutput, , null)
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_cust_serv_contact_type_lcode", adVarChar, adParamInput, 8, aRole(0))
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_customer_service_id", adNumeric, adParamInput, , Clng(Request("CustServID")))
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_contact_id", adNumeric, adParamInput, , Clng(Request("hdnContactID")))
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_contact_priority", adNumeric, adParamInput, , Clng(Request("selPriority")))

			'call the stored proc
			cmdInsertObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
				strCustServContactID = cmdInsertObj.Parameters("p_customer_service_contact_id").Value
			end if

			strWinMessage = "Record created successfully. You can now see the new record."

		end if

	case "DELETE"
		'delete record
		if (intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete then
			DisplayError "BACK", "", 0, "DELETE DENIED", "You don't have access to delete Service Location Contacts. Please contact your system administrator."
		end if

		dim cmdDeleteObj
		set cmdDeleteObj = server.CreateObject("ADODB.Command")
		set cmdDeleteObj.ActiveConnection = objConn
		cmdDeleteObj.CommandType = adCmdStoredProc
		cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_cserv_inter.sp_cs_cont_delete"

		cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_customer_service_contact_id", adNumeric, adParamInput, , clng(Request("CustServContactID")))					'number(9)
		cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update", adDBTimeStamp, adParamInput, ,Cdate(Request("hdnUpdateDateTime")))		'Date
        cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("P_USER_ID", adVarChar , adParamInput, 30, strRealUserID)

		cmdDeleteObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if

		strCustServContactID = NO_ID

		StrWinMessage = "Record deleted successfully."

end select

if strCustServContactID <> NO_ID then
	'load the page data if required
	strSQL = " SELECT cc.customer_service_contact_id" &_
			 " ,      cc.customer_service_id" &_
			 " ,      cc.cust_serv_contact_type_lcode" &_
			 " ,      cc.contact_priority" &_
			 " ,      c.contact_id" &_
			 " ,      c.contact_name" &_
			 " ,      c.last_name" &_
			 " ,      c.first_name" &_
			 " ,      c.position_title" &_
			 " ,      c.work_number" &_
			 " ,      c.work_number_ext" &_
			 " ,      c.home_number" &_
			 " ,      c.cell_number" &_
			 " ,      c.pager_number" &_
			 " ,      c.fax_number" &_
			 " ,      c.email_address" &_
			 " ,      w.customer_name works_for" &_
			 " ,      a.building_name" &_
			 " ,      a.long_street_name" &_
			 " ,      a.municipality_name || ' ' || a.province_state_lcode || ' ' || a.country_lcode place" &_
			 " ,      a.postal_code_zip" &_
			 " ,      cc.record_status_ind" &_
			 " ,      to_char(cc.create_date_time, 'MON-DD-YYYY HH24:MI:SS') create_date" &_
			 " ,      sma_sp_userid.spk_sma_library.sf_get_full_username(cc.create_real_userid) as create_real_userid" &_
			 " ,      to_char(cc.update_date_time, 'MON-DD-YYYY HH24:MI:SS') update_date" &_
			 " ,      cc.update_date_time" &_
			 " ,      sma_sp_userid.spk_sma_library.sf_get_full_username(cc.update_real_userid) as update_real_userid" &_
			 " FROM   crp.Customer_service_contact cc" &_
			 " ,      crp.contact c" &_
			 " ,      crp.customer w" &_
			 " ,      crp.address  a" &_
			 " WHERE  c.contact_id = cc.contact_id" &_
			 " AND    cc.customer_service_contact_id = " & strCustServContactID &_
			 " AND    w.customer_id = c.work_for_customer_id" &_
			 " AND    a.address_id (+)= c.address_id"

   set objCmd = Server.CreateObject("ADODB.command")
       objCmd.ActiveConnection = objconn
	   objCmd.CommandText = strSql
	   objCmd.CommandType = adCmdText

	'Create Recordset object
	set objRs = objCmd.Execute

	'work number
	dim strWkArea, strWkMid, strWkEnd, strWP
	strWkArea = mid(objRs("work_number"),1,3)
	strWkMid = mid(objRs("work_number"),4,3)
	strWkEnd = mid(objRs("work_number"),7,10)
	if objRS("work_number") <> "" then
		strWP = "(" & strWkArea & ") " & strWkMid & "-" & strWkEnd
	end if

	'home number
	dim strHmArea, strHmMid, strHmEnd, strHP
	strHmArea = mid(objRs("home_number"),1,3)
	strHmMid = mid(objRs("home_number"),4,3)
	strHmEnd = mid(objRs("home_number"),7,10)
	if objRS("home_number") <> "" then
		strhP = "(" & strHmArea & ") " & strHmMid & "-" & strHmEnd
	end if

	'cell number
	dim strClArea, strClMid, strClEnd, strCP
	strClArea = mid(objRs("cell_number"),1,3)
	strClMid = mid(objRs("cell_number"),4,3)
	strClEnd = mid(objRs("cell_number"),7,10)
	if objRS("cell_number") <> "" then
		strCP = "(" & strClArea & ") " & strClMid & "-" & strClEnd
	end if

	'pager
	dim strPgArea, strPgMid, strPgEnd, strPP
	strPgArea = mid(objRs("pager_number"),1,3)
	strPgMid = mid(objRs("pager_number"),4,3)
	strPgEnd = mid(objRs("pager_number"),7,10)
	if objRS("pager_number") <> "" then
		strPP = "(" & strPgArea & ") " & strPgMid & "-" & strPgEnd
	end if

	'fax number
	dim strFxArea, strFxMid, strFxEnd, strFP
	strFxArea = mid(objRs("fax_number"),1,3)
	strFxMid = mid(objRs("fax_number"),4,3)
	strFxEnd = mid(objRs("fax_number"),7,10)
	if objRS("fax_number") <> "" then
		strFP = "(" & strFxArea & ") " & strFxMid & "-" & strFxEnd
	end if

	'get contact name, customer name

	'build text for contact info box
	strContactInfo = "Works for:" & vbTab & objRs("WORKS_FOR") & vbNewLine &_
					 "Position:" & vbTab & objRs("POSITION_TITLE") & vbNewLine &_
					 "Work # :" & vbTab & strWP & " Ext: " & objRs("WORK_NUMBER_EXT") & vbNewLine &_
					 "Cell # :" & vbTab & strCP & vbNewLine &_
					 "Pager # :" & vbTab & strPP & vbNewLine &_
					 "Fax # :" & vbTab & strFP & vbNewLine &_
					 "Email:" & vbTab & objRS("EMAIL_ADDRESS") & vbNewLine &_
					 "Building:" & vbTab & objRs("BUILDING_NAME") & vbNewLine &_
					 "Address:" & vbTab & objRs("LONG_STREET_NAME") & vbNewLine &_
					 vbTab & objRs("PLACE") & vbNewLine &_
					 vbTab & objRs("POSTAL_CODE_ZIP")



end if

'always load the Contact Roles

strsql = " SELECT cust_serv_contact_type_lcode" &_
		 " ,      cust_serv_contact_type_desc" &_
		 " FROM   crp.lcode_cust_serv_contact_type" &_
		 " ORDER  BY cust_serv_contact_type_lcode"

'Create the command object
set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = objconn
    objCmd.CommandText = strSql
    objCmd.CommandType = adCmdText

'Create Recordset object
set objRsContactRole = objCmd.Execute

%>

<HTML>
<HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK REL="stylesheet" TYPE="text/css" HREF="stylesheets/styles.css">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<TITLE>Service Location Contact Detail</TITLE>
<SCRIPT LANGUAGE=javascript>
<!--
var intAccessLevel = <%=intAccessLevel%>;
var boolNeedToSave = false;

function btnNew_click()
{
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create)
	{
		alert('Access Denied. Please contact your system administrator.');
		return;
	}

	self.document.location.href = "<%=ASP_NAME%>?NewFacility=NEW&txtWorkFor=" + document.frmCustServContact.txtWorkFor.value + "&CustServID=" + document.frmCustServContact.CustServID.value;
}

function fct_onDelete()
{
	if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete)
	{
		alert('Access Denied. Please contact your system administrator.');
		return;
	}

	if (confirm('Do you really want to delete this object?'))
	{
		boolNeedToSave = false;
		document.location = "<%=ASP_NAME%>?hdnFrmAction=DELETE&CustServContactID="+document.frmCustServContact.CustServContactID.value+"&hdnUpdateDateTime="+document.frmCustServContact.hdnUpdateDateTime.value;
	}
}

function btnClose_onclick()
{
	window.close();
	parent.opener.iFrame_display();
}

//-->
</SCRIPT>


<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
function body_onBeforeUnload()
{
	//must set focus to save button because if user has changed only one field and has not
	//left it the on_change event will not have fired and the flag that determines whether
	//you need to save or not will be false
	document.frmCustServContact.btnSave.focus();

	//before you do the code below check that the 'need to save' flag is true and check
	//the user's access for either insert or update depending on which is
	//appropriate (i.e. for most of us this means whether the main id = 0 for a new record
	//or a value < or > 0 for an existing record)
	if ((intAccessLevel & intConst_Access_Update) == intConst_Access_Update || (intAccessLevel & intConst_Access_Create) == intConst_Access_Create)
	{
		if (boolNeedToSave == true)
		{
			event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
		}
	}
}

function frmCustServContact_onsubmit()
{
	if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update || (intAccessLevel & intConst_Access_Create) != intConst_Access_Create)
	{
		alert('Access Denied. Please contact your system administrator.');
		return false;
	}

	document.frmCustServContact.hdnFrmAction.value = "SAVE";
	boolNeedToSave = false;
	document.forms[0].submit();
	return true;
}

function btnContactLookup_onClick()
{

	if (document.frmCustServContact.txtWorkFor.value != "")
	{
		SetCookie("WorkFor", document.frmCustServContact.txtWorkFor.value);
	}
	if (document.frmCustServContact.txtLName.value != "")
	{
		 SetCookie("LName", document.frmCustServContact.txtLName.value);
	}
	if (document.frmCustServContact.txtFName.value != "")
	{
		 SetCookie("FName", document.frmCustServContact.txtFName.value);
	}

	SetCookie("WinName", 'Popup');
	window.open('SearchFrame.asp?fraSrc=Contact', 'Popup', 'top=50, left=100, height=600, width=900' ) ;
	on_change();
}

function fct_onChangeRole() {

	var strWhole;
	var strRoleDesc, intStart, intIndex;

	intIndex = document.frmCustServContact.selRole.selectedIndex;
	strWhole = document.frmCustServContact.selRole.options[intIndex].value;
	intStart = strWhole.indexOf('<%=strDelimiter%>');
	document.frmCustServContact.txtRoleDesc.value = strWhole.substr(intStart+1);
	on_change();
}

function on_change()
{
	boolNeedToSave = true;
}
//-->
</SCRIPT>
</HEAD>
<BODY onload="parent.opener.iFrame_display();" onbeforeunload="body_onBeforeUnload();">

<FORM name=frmCustServContact ACTION="<%=ASP_NAME%>" METHOD="POST" LANGUAGE=javascript >

<INPUT name="hdnFrmAction"      type="hidden" value="">
<INPUT name="hdnUpdateDateTime" type="hidden" value="<%if strCustServContactID <> NO_ID then  Response.Write objRS("UPDATE_DATE_TIME")              else Response.Write null  end if%>" >
<INPUT name="CustServContactID" type="hidden" value="<%if strCustServContactID <> NO_ID then  Response.Write objRS("CUSTOMER_SERVICE_CONTACT_ID")   else Response.Write null  end if%>" >
<INPUT name="CustServID"        type="hidden" value="<%if strCustServContactID <> NO_ID then  Response.Write objRS("CUSTOMER_SERVICE_ID")           else Response.Write strCustServID end if%>">
<INPUT name="hdnContactID"      type="hidden" value="<%if strCustServContactID <> NO_ID then  Response.Write objRS("CONTACT_ID")                    else Response.Write null  end if%>" >
<INPUT name="txtLName"          type="hidden" value="<%if strCustServContactID <> NO_ID then  Response.Write routineHTMLString(objRS("LAST_NAME"))  else Response.Write null  end if%>" >
<INPUT name="txtFName"          type="hidden" value="<%if strCustServContactID <> NO_ID then  Response.Write routineHTMLString(objRS("FIRST_NAME")) else Response.Write null  end if%>" >
<INPUT name="txtWorkFor"        type="hidden" value="<%if strCustServContactID <> NO_ID then  Response.Write routineHTMLString(objRS("WORKS_FOR"))  else Response.Write routineHTMLString(strCustomerName) end if%>" >

<TABLE border=0 width=100%>
	<thead>
		<TR ><TD colspan=2>Customer Service Contact Detail</td></tr>
	</thead>
	<tbody>
	<TR>
		<TD ALIGN=RIGHT width=20% NOWRAP>Contact Role<font color=red>*</font></TD>
		<TD width=80%>
			<SELECT id=selRole name=SelRole style="HEIGHT: 20px; WIDTH: 120px" onChange="return fct_onChangeRole();">
			<%
			Dim roleDesc	'used to set the intial value of txtRoleDesc
			Do while Not objRSContactRole.EOF
				Response.write "<OPTION "
				if roleDesc = "" then
					roleDesc = objRSContactRole("CUST_SERV_CONTACT_TYPE_DESC")
				end if
				if strCustServContactID <> NO_ID then
					if objRSContactRole("CUST_SERV_CONTACT_TYPE_LCODE") = objRS("CUST_SERV_CONTACT_TYPE_LCODE") then
						Response.Write " selected "
						roleDesc = objRSContactRole("CUST_SERV_CONTACT_TYPE_DESC")
					end if
				end if
				Response.Write " VALUE =""" & routineHTMLString(objRSContactRole("CUST_SERV_CONTACT_TYPE_LCODE")& strDelimiter & objRSContactRole("CUST_SERV_CONTACT_TYPE_DESC")) & """>" & routineHTMLString(objRSContactRole("CUST_SERV_CONTACT_TYPE_LCODE")) & "</OPTION>"
				objRSContactRole.MoveNext
			Loop
			%>
			</SELECT>
			<INPUT id=txtRole name=txtRoleDesc value="<%=roleDesc%>" disabled style="WIDTH: 380px">
		</TD>
	</TR>
	<TR>
		<TD ALIGN=RIGHT width=20% NOWRAP>Contact Priority<font color=red>*</font></TD>
		<TD width=80%>
			<SELECT id=selPriority name=SelPriority onchange="return on_change();">
			<%
			dim i
			for i = 1 to 30
				 Response.write "<OPTION "
				if strCustServContactID <> NO_ID then
					if i = CLng(objRS("CONTACT_PRIORITY")) then
						Response.Write " selected "
					end if
				end if

				 Response.write "value=" & i & ">" & i & "</OPTION>" & vbNewLine
			next
			%>
			</SELECT>
		</TD>
	</TR>
	<TR>
		<TD ALIGN=RIGHT width=20% NOWRAP>Contact Name<font color=red>*</font></TD>
		<TD colspan=3 width=80%>
			<INPUT id=txtContactName    name=txtContactName    style="HEIGHT: 21px; WIDTH: 500px" disabled value="<%if strCustServContactID <> NO_ID then  Response.Write "" & objRS("CONTACT_NAME") else Response.Write null end if%>">
			<INPUT id=btnContactLookup  name=btnContactLookup  style="HEIGHT: 21px; WIDTH: 19px"  type=button value="..." onclick="return  btnContactLookup_onClick()">
		</TD>
	</TR>
	<TR>
		<td align=right valign=top width=20%>Contact Information</td>
		<td align=left width=50% colspan=2 ><textarea disabled name=txtContactInfo cols=85 style="HEIGHT: 200px"><%if strCustServContactID <> NO_ID then Response.Write "" & strContactInfo else Response.Write null end if%></textarea></td>
	</TR>
	</tbody>
</TABLE>

<TABLE>
	  <TR><TD align=right colspan=5>
			<INPUT id=btnClose  name=btnClose  type=button value=Close  style="WIDTH: 2cm" onclick="return btnClose_onclick()">&nbsp;&nbsp;
			<INPUT id=btnDelete name=btnDelete type=button value=Delete style="WIDTH: 2cm" onclick="return fct_onDelete();">&nbsp;&nbsp;
			<INPUT id=btnReset  name=btnReset  type=reset  value=Reset  style="WIDTH: 2cm" >&nbsp;&nbsp;
			<INPUT id=btnAddNew name=btnAddNew type=button value=New    style="WIDTH: 2cm" onclick="return btnNew_click()">&nbsp;&nbsp;
			<INPUT id=btnSave   name=btnSave   type=button value=Save   style="WIDTH: 2cm" onclick="return frmCustServContact_onsubmit();">&nbsp;&nbsp;
	  </TD></TR>
</table>

<FIELDSET >
	<LEGEND ALIGN=RIGHT><B>Audit Information</B></LEGEND>
	<Div SIZE=8pt ALIGN=RIGHT>
		Record Status Indicator:
		<INPUT align=left name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 18px"disabled value=<%if  strCustServContactID <> NO_ID then  Response.Write """"&objRS("RECORD_STATUS_IND")&"""" else Response.Write """""" end if%>  >&nbsp;&nbsp;&nbsp;
		Create Date:&nbsp;&nbsp;
		<INPUT align=center name=txtcrdate type=text style="HEIGHT: 20px; WIDTH: 150px"disabled value=<%if  strCustServContactID <> NO_ID then  Response.Write """"&objRS("CREATE_DATE")&"""" else Response.Write """""" end if%> >&nbsp;
		&nbsp;
		Created By:
		<INPUT align=right name=txtcrby type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if  strCustServContactID <> NO_ID then  Response.Write """"&objRS("CREATE_REAL_USERID")&"""" else Response.Write """""" end if%> ><BR>
		Update Date:
		<INPUT align= center name=txtupdate type=text style="HEIGHT: 20px; WIDTH: 150px"disabled value=<%if  strCustServContactID <> NO_ID then  Response.Write """"&objRS("UPDATE_DATE")&"""" else Response.Write """""" end if%>  >&nbsp;&nbsp;
		Updated By:
		<INPUT align=right name=txtupby type=text style="HEIGHT: 20px; WIDTH: 100px"disabled value=<%if  strCustServContactID <> NO_ID then  Response.Write """"&objRS("UPDATE_REAL_USERID")&"""" else Response.Write """""" end if%>  >
	</DIV>
</FIELDSET>

</FORM>
<%

if strCustServContactID <> NO_ID then

	'Clean up our ADO objects if they were opened
	objRS.close
	set objRS = Nothing

	objRSContactRole.close
	set objRSContactRole = Nothing

	objConn.close
	set ObjConn = Nothing

end if

%>


</BODY>
</HTML>
