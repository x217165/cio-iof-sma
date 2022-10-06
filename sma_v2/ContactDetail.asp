<%@ Language=VBScript %>
<% option explicit
 on error resume next
 Response.Buffer = true %>

<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--
**********************************************************************************
* Page name:		ContactDetail.asp
* Purpose:			To display the detailed information about a contact.
* Input parameter:
* Created by:		Nancy Mooney	08/18/2000
***************************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       19-Feb-02	  DTy		Increase email address size from 50 t0 60.
       08-Mar-02	  DTy		Add 'Contact ID' as a displayable field.
       03-Oct-07        ACheung		Add Reponsibility (50 chars) & PIN (1 char) fields to the contact table
**********************************************************************************
-->
<%
'*** SECURITY ********************************************************************

dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_Contact))
if ((intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly) then
	DisplayError "BACK","",0,"ACCESS DENIED","You do not have access to contact. Please contact your system administrator."
end if

'*********************************************************************************

dim strWinMessage

dim strRealUserID
strRealUserID = Session("username")
'Response.Write strRealUserID & "<BR>"

'get info. needed when submitting form to self
dim lngContactID, datUpdateDateTime, strWinLocation

	lngContactID = Request("ContactID")
	datUpdateDateTime = Request("UpdateDateTime")
	strWinLocation = "ContactDetail.asp?ContactID="&Request("hdnContactID")
	'Response.Write(Request("txtFrmAction"))

'Form action
select case Request("txtFrmAction")
	case "SAVE"
		'Response.Write intAccessLevel
		'Response.End

		if ((intAccessLevel and intConst_Access_Update) <> intConst_Access_Update) then
			DisplayError "BACK", "", 0, "UPDATE DENIED", "You don't have access to update contacts. Please contact your system administrator"
		end if

		'get value from checkbox
		dim strStaffFlag, strReceivePub, strPINAccess, strresponsibility

		if Lcase(Request.Form("chkStaffFlag")) = "on" then
			strStaffFlag = "Y"
		else
			strStaffFlag = "N"
		end if

		if Lcase(Request.Form("chkReceivePub")) = "on" then
			strReceivePub = "Y"
		else
			strReceivePub = "N"
		end if

		if Lcase(Request.Form("chkPINAccess")) = "on" then
			strPINAccess = "Y"
		else
			strPINAccess = "N"
		end if

		strresponsibility = Request.Form("txtResponsibility")

		'Response.Write " PIN: " & strPINAccess  & ""
		'Response.Write " Responsibility: " & strresponsibility  & ""
		'Response.Write " Rec Pub: " & strReceivePub  & ""
		'Response.Write " Staff Flag: " & strStaffFlag  & ""
		'Response.End

		'parse together phone numbers
		'work phone
		dim strWPhone,strWArea,strWMid,strWEnd
		strWArea = Request("txtWArea")
		strWMid = Request("txtWMid")
		strWEnd = Request("txtWEnd")
		if strWArea <> "" then
			strWPhone = strWPhone & strWArea
		end if
		if strWMid <> "" then
			strWPhone = strWPhone & strWMid
		end if
		if  strWEnd <> "" then
			strWPhone = strWPhone & strWend
		end if

		'cell phone
		dim strCPhone,strCArea,strCMid,strCEnd
		strCArea = Request("txtCArea")
		strCMid = Request("txtCMid")
		strCEnd = Request("txtCEnd")
		if strCArea <> "" then
			strCPhone = strCPhone & strCArea
		end if
		if strCMid <> "" then
			strCPhone = strCPhone & strCMid
		end if
		if  strCEnd <> "" then
			strCPhone = strCPhone & strCend
		end if

		'pager
		dim strPPhone,strPArea,strPMid,strPEnd
		strPArea = Request("txtPArea")
		strPMid = Request("txtPMid")
		strPEnd = Request("txtPEnd")
		if strPArea <> "" then
			strPPhone = strPPhone & strPArea
		end if
		if strPMid <> "" then
			strPPhone = strPPhone & strPMid
		end if
		if  strPEnd <> "" then
			strPPhone = strPPhone & strPend
		end if

		'fax
		dim strFPhone,strFArea,strFMid,strFEnd
		strFArea = Request("txtFArea")
		strFMid = Request("txtFMid")
		strFEnd = Request("txtFEnd")
		if strFArea <> "" then
			strFPhone = strFPhone & strFArea
		end if
		if strFMid <> "" then
			strFPhone = strFPhone & strFMid
		end if
		if  strFEnd <> "" then
			strFPhone = strFPhone & strFend
		end if

		'home phone
		dim strHPhone,strHArea,strHMid,strHEnd
		strHArea = Request("txtHArea")
		strHMid = Request("txtHMid")
		strHEnd = Request("txtHEnd")
		if strHArea <> "" then
			strHPhone = strHPhone & strHArea
		end if
		if strHMid <> "" then
			strHPhone = strHPhone & strHMid
		end if
		if  strHEnd <> "" then
			strHPhone = strHPhone & strHend
		end if

		if Request("hdnContactID") <> "" then	'*** EXISTING record --> UPDATE ***

			'Response.Write("entering update record")
			'Response.end

			'security check
			'Response.Write intAccessLevel
			'Response.end
			if (intAccessLevel and intConst_Access_Update <> intConst_Access_Update) then
				DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update contacts. Please contact your system administrator."
			end if

			'create command object for stored procedure
			dim cmdUpdateObj
			set cmdUpdateObj = server.CreateObject("ADODB.Command")
			set cmdUpdateObj.ActiveConnection = objConn
			cmdUpdateObj.CommandType = adCmdStoredProc
			cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_cust_inter.sp_contact_update"

			lngContactID = Request("hdnContactID")

			'create parameters
			'required fields
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_userid", adVarChar,adParamInput, 20, strRealUserID)						'varchar2(30) Real User ID
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_contact_id",adNumeric, adParamInput,9,CLng(Request("hdnContactId")))	'number(9) Contact id
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_customer_id",adNumeric, adParamInput,9,CLng(Request("hdnCustomerId")))	'number(9) Customer name
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_language_preference",adChar, adParamInput,2,Request("selLangPref"))		'char(2) Language Preference
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_staff_flag",adChar, adParamInput,1,strStaffFlag)	   'char(1) Staff flag
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_last_update",adDBTimeStamp, adParamInput,,CDate(Request("hdnUpdateDateTime")))'date Update date/time
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_last_name",adVarChar, adParamInput,50,Request("txtLName"))	'varchar(50) Last name
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_first_name",adVarChar, adParamInput,20,Request("txtFName"))	'varchar(20) First name

			'optional fields
			if Request("txtMName") <> "" then
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_middle_name",adVarChar, adParamInput,7,Request("txtMName"))	'varchar(7) Middle Name
			else
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_middle_name",adVarChar, adParamInput,7,null)
			end if
			if Request("selTitle") <> "" then
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_title",adVarChar, adParamInput,6,Request("selTitle")) 'varchar(6) Title
			else
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_title",adVarChar, adParamInput,6,null) 'varchar(6) Title
			end if
			if strWPhone <> "" then
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_work_number",adVarChar, adParamInput,24,strWPhone)    'varchar(24) Work Phone Number
			else
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_work_number",adVarChar, adParamInput,24,null)    'varchar(24) Work Phone Number
			end if
			if Request("txtExt") <> "" then
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_work_numbr_ext",adVarChar, adParamInput,10,Request("txtExt"))'varchar(10) Work Phone Extension
			else
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_work_numbr_ext",adVarChar, adParamInput,10,null)'varchar(10) Work Phone Extension
			end if
			if strHPhone <> "" then
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_home_number",adVarChar, adParamInput,24,strHPhone)    'varchar(24) Home Phone Number
			else
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_home_number",adVarChar, adParamInput,24,null)    'varchar(24) Home Phone Number
			end if
			if strCPhone <> "" then
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_cell_number",adVarChar, adParamInput,24,strCPhone)    'varchar(24) Cell Phone Number
			else
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_cell_number",adVarChar, adParamInput,24,null)    'varchar(24) Cell Phone Number
			end if
			if strPPhone <> "" then
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_pager_number",adVarChar, adParamInput,24,strPPhone)   'varchar(24) Pager Number
			else
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_pager_number",adVarChar, adParamInput,24,null)   'varchar(24) Pager Number
			end if
			if strFPhone <> "" then
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_fax_number",adVarChar, adParamInput,24,strFPhone)     'varchar(24) Fax Number
			else
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_fax_number",adVarChar, adParamInput,24,null)     'varchar(24) Fax Number
			end if
			if Request("txtEmail") <> "" then
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_email_address",adVarChar, adParamInput,80,Request("txtEmail"))	'varchar(60) Email address
			else
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_email_address",adVarChar, adParamInput,80,null)	'varchar(50) Email address
			end if
			if Request("txtWebSite") <> "" then
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_web_site",adVarChar, adParamInput,50,Request("txtWebSite"))     'varchar(60) Web Site
			else
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_web_site",adVarChar, adParamInput,50,null)     'varchar(50) Web Site
			end if
			if Request("txtPosition") <> "" then
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_position",adVarChar, adParamInput,50,Request("txtPosition"))    'varchar(50) Position
			else
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_position",adVarChar, adParamInput,50,null)    'varchar(50) Position
			end if
			if Request("hdnAddressID") <> "" then
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_address_id",adNumeric, adParamInput,9,CLng(Request("hdnAddressID"))) 'number(9) Address ID
			else
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_address_id",adNumeric, adParamInput,9,null) 'number(9) Address ID
			end if
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_receive_publications",adChar, adParamInput,1,strReceivePub)     'char(1) Receive Publications Flag
			if Request("selPrefContMeth") <> "" then
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_prefer_contact",adVarChar, adParamInput,6, Request("selPrefContMeth"))    'varchar(6) Preferred Contact Method
			else
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_prefer_contact",adVarChar, adParamInput,6, null)    'varchar(6) Preferred Contact Method
			end if
			if Request("selAvailSched") <> "" then
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_schedule_id",adNumeric, adParamInput,9,Request("selAvailSched"))    'number(9) Availability
			else
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_schedule_id",adNumeric, adParamInput,9,null)    'number(9) Availability
			end if

			if Request("txtComments") <> "" then
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_comments",adVarChar, adParamInput,2000,Request("txtComments"))    'varchar(2000) Comments
			else
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_comments",adVarChar, adParamInput,2000,null)    'varchar(2000) Comments
			end if
			if Request("txtResponsibility") <> "" then
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_responsibility",adVarChar, adParamInput,50,routineOraString(Request("txtResponsibility")))     'varchar(50) Responsibility
			else
				cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_responsibility",adVarChar, adParamInput,50,null)     'varchar(50) Responsibility
			end if
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter ("p_pinaccess",adChar, adParamInput,1,strPINAccess)     'char(1) PIN Access Flag

			'parameter check - development
			'cmdUpdateObj.Parameters.Refresh
			'dim objparm
			'for each objparm in cmdUpdateObj.Parameters
			'	Response.Write "<b>" & objparm.name & "</b>"
			'	Response.Write " and value: " & objparm.value & ""
			'	Response.Write " and datatype: " & objparm.Type & "<br>"
			'next

			'Response.Write "<b> count = " & cmdUpdateObj.Parameters.count & "<br>"
			'dim nx
			'for nx = 0 to cmdUpdateObj.Parameters.Count-1
			'	Response.Write "parm value = " & cmdUpdateObj.Parameters.Item(nx).Value & " <br>"
			'next
			'Response.end

			'call the update stored proc
			on error resume next
			cmdUpdateObj.Execute
			If objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if
			'Response.Write ("Record saved successfully")
			strWinMessage = "Record saved successfully."

		else '*** NEW record --> CREATE ***
			'Response.Write ("This is a new record")
			if ((intAccessLevel and intConst_Access_Create) <> intConst_Access_Create) then
				DisplayError "REFRESH", strWinLocation, 0, "INSERT DENIED", "You don't have access to create a contact. Please contact your system administrator."
			end if

			dim cmdInsertObj
			set cmdInsertObj = server.CreateObject("ADODB.Command")
			set cmdInsertObj.ActiveConnection = objConn
			cmdInsertObj.CommandType = adCmdStoredProc
			cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_cust_inter.sp_contact_insert"

			'create parameters
			'required fields
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_userid", adVarChar,adParamInput, 20, strRealUserID)						'varchar2(30) Real User ID
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_contact_id",adNumeric, adParamOutput ,,null)	'number(9) Contact id
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_customer_id",adNumeric, adParamInput,9,CLng(Request("hdnCustomerId")))	'number(9) Customer name
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_language_preference",adChar, adParamInput,2,Request("selLangPref"))		'char(2) Language Preference
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_staff_flag",adChar, adParamInput,1,strStaffFlag)	   'char(1) Staff flag
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_last_name",adVarChar, adParamInput,50,Request("txtLName"))	'varchar(50) Last name
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_first_name",adVarChar, adParamInput,20,Request("txtFName"))	'varchar(20) First name
			'optional fields
			if Request("txtMName") <> "" then
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_middle_name",adVarChar, adParamInput,7,Request("txtMName"))	'varchar(7) Middle Name
			else
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_middle_name",adVarChar, adParamInput,7,null)
			end if
			if Request("selTitle") <> "" then
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_title",adVarChar, adParamInput,6,Request("selTitle")) 'varchar(6) Title
			else
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_title",adVarChar, adParamInput,6,null) 'varchar(6) Title
			end if
			if strWPhone <> "" then
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_work_number",adVarChar, adParamInput,24,strWPhone)    'varchar(24) Work Phone Number
			else
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_work_number",adVarChar, adParamInput,24,null)    'varchar(24) Work Phone Number
			end if
			if Request("txtExt") <> "" then
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_work_numbr_ext",adVarChar, adParamInput,10,Request("txtExt"))'varchar(10) Work Phone Extension
			else
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_work_numbr_ext",adVarChar, adParamInput,10,null)'varchar(10) Work Phone Extension
			end if
			if strHPhone <> "" then
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_home_number",adVarChar, adParamInput,24,strHPhone)    'varchar(24) Home Phone Number
			else
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_home_number",adVarChar, adParamInput,24,null)    'varchar(24) Home Phone Number
			end if
			if strCPhone <> "" then
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_cell_number",adVarChar, adParamInput,24,strCPhone)    'varchar(24) Cell Phone Number
			else
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_cell_number",adVarChar, adParamInput,24,null)    'varchar(24) Cell Phone Number
			end if
			if strPPhone <> "" then
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_pager_number",adVarChar, adParamInput,24,strPPhone)   'varchar(24) Pager Number
			else
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_pager_number",adVarChar, adParamInput,24,null)   'varchar(24) Pager Number
			end if
			if strFPhone <> "" then
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_fax_number",adVarChar, adParamInput,24,strFPhone)     'varchar(24) Fax Number
			else
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_fax_number",adVarChar, adParamInput,24,null)     'varchar(24) Fax Number
			end if
			if Request("txtEmail") <> "" then
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_email_address",adVarChar, adParamInput,80,Request("txtEmail"))	'varchar(60) Email address
			else
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_email_address",adVarChar, adParamInput,80,null)	'varchar(60) Email address
			end if
			if Request("txtWebSite") <> "" then
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_web_site",adVarChar, adParamInput,50,Request("txtWebSite"))     'varchar(50) Web Site
			else
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_web_site",adVarChar, adParamInput,50,null)     'varchar(50) Web Site
			end if
			if Request("txtPosition") <> "" then
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_position",adVarChar, adParamInput,50,Request("txtPosition"))    'varchar(50) Position
			else
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_position",adVarChar, adParamInput,50,null)    'varchar(50) Position
			end if
			if Request("hdnAddressID") <> "" then
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_address_id",adNumeric, adParamInput,,CLng(Request("hdnAddressID"))) 'number(9) Address ID
			else
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_address_id",adNumeric, adParamInput,,null) 'number(9) Address ID
			end if

			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_receive_publications",adChar, adParamInput,1,strReceivePub)     'char(1) Receive Publications Flag

			if Request("selPrefContMeth") <> "" then
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_prefer_contact",adVarChar, adParamInput,6, Request("selPrefContMeth"))    'varchar(6) Preferred Contact Method
			else
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_prefer_contact",adVarChar, adParamInput,6, null)    'varchar(6) Preferred Contact Method
			end if
			if Request("selAvailSched") <> "" then
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_schedule_id",adNumeric, adParamInput,9,Request("selAvailSched"))    'number(9) Availability
			else
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_schedule_id",adNumeric, adParamInput,9,null)    'number(9) Availability
			end if
			if Request("txtComments") <> "" then
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_comments",adVarChar, adParamInput,2000,Request("txtComments"))    'varchar(2000) Comments
			else
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_comments",adVarChar, adParamInput,2000,null)    'varchar(2000) Comments
			end if

			if Request("txtResponsibility") <> "" then
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_responsibility",adVarChar, adParamInput,50,Request("txtResponsibility"))     'varchar(50) Responsibility
			else
				cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_responsibility",adVarChar, adParamInput,50,null)     'varchar(50) Responsibility
			end if
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter ("p_pinaccess",adChar, adParamInput,1,strPINAccess)     'char(1) PIN Access Flag


			'Response.Write ("Now write the parameters")
			'call the insert stored proc
  			'cmdInsertObj.Parameters.Refresh

  			'dim objparm
  			'for each objparm in cmdInsertObj.Parameters
  			'  Response.Write "<b>" & objparm.name & "</b>"
  			'  Response.Write " has size:  " & objparm.Size & " "
  			'  Response.Write " and value:  " & objparm.value & " "
  			'  Response.Write " and datatype:  " & objparm.Type & "<br> "
  		    'next

			'Response.Write "<b> count = " & cmdInsertObj.Parameters.count & "<br>"
			'dim nx
			'for nx = 0 to cmdInsertObj.Parameters.Count-1
			'	Response.Write (nx + 1) & " parm value = " & cmdInsertObj.Parameters.Item(nx).Value & " <br>"
			'next
			'Response.End

			'call the insert stored proc
			on error resume next
			cmdInsertObj.Execute

			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT CREATE NEW OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
				lngContactID = cmdInsertObj.Parameters("p_contact_id").Value
			end if
			strWinMessage = "Record created successfully."

		end if
	case "DELETE"
		if ((intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete) then
			DisplayError "REFRESH", strWinLocation, 0, "DELETE DENIED", "You don't have access to delete contacts. Please contact your system administrator."
		end if

		dim cmdDeleteObj
		set cmdDeleteObj = server.CreateObject("ADODB.Command")
		set cmdDeleteObj.ActiveConnection = objConn
		cmdDeleteObj.CommandType = adCmdStoredProc
		cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_cust_inter.sp_contact_delete"
		cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_contact_id", adNumeric, adParamInput, ,clng(lngContactID))	'Number(9)
		cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update", adDBTimeStamp, adParamInput,,cdate(datUpdateDateTime)) 'Date
         cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("P_USER_ID", adVarChar , adParamInput, 30, strRealUserID)

		'Response.Write "<b> count = & cmdDeleteObj.Parameters.count & <br>"
			'dim nx
			'for nx = 0 to cmdDeleteObj.Parameters.Count-1
			'	Response.Write "parm value = " & cmdDeleteObj.Parameters.Item(nx).Value & " <br>"
		'next
		on error resume next
		cmdDeleteObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if

		lngContactID = 0
		strWinMessage = "Record Deleted Successfully."

	end select

	if lngContactID <> 0 then

		'create SQL for populating fields
		Dim  strSQL, strSelectClause, strFromClause, strWhereClause

		strSelectClause = "select " &_
					"t1.contact_id, " & _
					"t1.work_for_customer_id, " & _
					"t1.contact_name, " & _
					"t1.last_name, " & _
					"t1.first_name, " & _
					"t1.middle_name, " & _
					"t1.name_prefix, " & _
					"t1.work_number, " & _
					"t1.work_number_ext, " & _
					"t1.home_number, " & _
					"t1.cell_number, " & _
					"t1.pager_number, " & _
					"t1.fax_number, " & _
					"t1.email_address, " & _
					"t1.web_site_url, " & _
					"t1.position_title, " & _
					"t1.address_id, " & _
					"t1.receive_publications_flag, " & _
					"t1.prefercontactmethodcode, " & _
					"t1.availablescheduleid, " & _
					"t1.comments, " & _
					"t1.language_preference_lcode, " & _
					"t1.staff_flag, " & _
					"t1.record_status_ind, " &_
					"to_char(t1.create_date_time,'MON-DD-YYYY HH24:MI:SS') as create_date, " & _
					"sma_sp_userid.spk_sma_library.sf_get_full_username(t1.create_real_userid) as create_real_userid, " & _
					"to_char(t1.update_date_time,'MON-DD-YYYY HH24:MI:SS') as update_date, " & _
					"sma_sp_userid.spk_sma_library.sf_get_full_username(t1.update_real_userid) as update_real_userid, " & _
					"t1.update_date_time as last_update_date_time, " & _
					"t2.customer_name, " & _
					"(t3.building_name || chr(10) || t3.street ||chr(10)|| " & _
					"t3.municipality_name||' '|| t3.province_state_lcode " & _
					"||' '|| t3.country_lcode ||chr(10) || t3.postal_code_zip) contact_address, " &_
					"t4.contact_method_code, " & _
					"t5.schedule_id, " & _
					"t6.language_preference_lcode, " & _
					"t1.responsibility, " & _
					"t1.PIN_access "

		strFromClause =	" from crp.contact t1, " &_
					"crp.customer t2, " & _
					"crp.v_address_consolidated_street t3, " & _
					"crp.contact_method t4, " & _
					"crp.schedule t5, " & _
					"crp.lcode_language_preference t6 "

		strWhereClause = " where " & _
					"t1.work_for_customer_id = t2.customer_id and " & _
					"t1.address_id = t3.address_id (+) and " & _
					"t1.prefercontactmethodcode = t4.contact_method_code (+) and " & _
					"t1.availablescheduleid = t5.schedule_id (+) and " & _
					"t1.language_preference_lcode = t6.language_preference_lcode (+) and " & _
					"t1.contact_id = " & lngContactID

		strSql =  strSelectClause & strFromClause & strWhereClause

		'Response.Write strSQL
		'response.end

		'get the contact recordset
		Dim rsContact
		set rsContact = Server.CreateObject("ADODB.Recordset")
		rsContact.CursorLocation = adUseClient
		rsContact.Open strSql,objConn
		if err then
			DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
		end if
		if rsContact.EOF then
			DisplayError "BACK", "", 999, "CANNOT FIND REQUESTED OBJECT", "EOF condition occurred in rsContact recordset."
		end if
		'set rsContact.ActiveConnection = nothing

		'parse out phone numbers

		'work number
		dim strWkArea, strWkMid, strWkEnd
		strWkArea = mid(rsContact("work_number"),1,3)
		strWkMid = mid(rsContact("work_number"),4,3)
		strWkEnd = mid(rsContact("work_number"),7,10)

		'home number
		dim strHmArea, strHmMid, strHmEnd
		strHmArea = mid(rsContact("home_number"),1,3)
		strHmMid = mid(rsContact("home_number"),4,3)
		strHmEnd = mid(rsContact("home_number"),7,10)

		'cell number
		dim strClArea, strClMid, strClEnd
		strClArea = mid(rsContact("cell_number"),1,3)
		strClMid = mid(rsContact("cell_number"),4,3)
		strClEnd = mid(rsContact("cell_number"),7,10)

		'pager
		dim strPgArea, strPgMid, strPgEnd
		strPgArea = mid(rsContact("pager_number"),1,3)
		strPgMid = mid(rsContact("pager_number"),4,3)
		strPgEnd = mid(rsContact("pager_number"),7,10)

		'fax number
		dim strFxArea, strFxMid, strFxEnd
		strFxArea = mid(rsContact("fax_number"),1,3)
		strFxMid = mid(rsContact("fax_number"),4,3)
		strFxEnd = mid(rsContact("fax_number"),7,10)

		'get contact name, customer name
		dim strContactName, strCustomerName
		strContactName = routineHtmlString(rsContact("CONTACT_NAME"))
		strCustomerName = routineHtmlString(rsContact("CUSTOMER_NAME"))

		'lists: get a list of roles assigned to this contact (= 3 recordsets)
		'get customer contact roles
		strSQL = "select s.customer_name, " &_
					"cs.customer_contact_type_lcode, " &_
					"l.customer_contact_type_desc, " &_
					"cs.contact_priority " &_
				"from crp.contact c, " &_
					"crp.customer_contact cs, " &_
					"crp.customer s, " &_
					"crp.lcode_customer_contact_type l  " &_
				"where c.contact_id = cs.contact_id  " &_
				"and	  cs.customer_id = s.customer_id " &_
				"and	  cs.customer_contact_type_lcode = l.customer_contact_type_lcode " &_
				"and c.contact_id = " & lngContactID &_
				"and cs.record_status_ind = 'A' " & _
				"order by s.customer_name, l.customer_contact_type_lcode, cs.contact_priority"

		Dim objRsCustContacts
		set objRsCustContacts = Server.CreateObject("ADODB.Recordset")
		objRsCustContacts.CursorLocation = adUseClient
		objRsCustContacts.Open strSQL,objConn
		if err then
			DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
		end if
		set objRsCustContacts.ActiveConnection = nothing

		'get customer service contacts
		strSQL = "select s.customer_service_desc,  " &_
					"cs.cust_serv_contact_type_lcode, " &_
 					"l.cust_serv_contact_type_desc, " &_
					"cs.contact_priority " &_
				"from crp.contact c, " &_
					"crp.customer_service_contact cs, " &_
					"crp.customer_service s, " &_
					"crp.lcode_cust_serv_contact_type l   " &_
				"where c.contact_id = cs.contact_id " &_
				"and  cs.customer_service_id = s.customer_service_id " &_
				"and  cs.cust_serv_contact_type_lcode = l.cust_serv_contact_type_lcode " &_
				"and c.contact_id =  " & lngContactID &_
				"and cs.record_status_ind = 'A'" &_
				"order by s.customer_service_desc, l.cust_serv_contact_type_lcode , cs.contact_priority"

		Dim objRsCustServContacts
		set objRsCustServContacts = Server.CreateObject("ADODB.Recordset")
		objRsCustServContacts.CursorLocation = adUseClient
		objRsCustServContacts.Open strSQL,objConn
		if err then
			DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
		end if
		set objRsCustServContacts.ActiveConnection = nothing

		'get service location contacts
		strSQL= "select s.service_location_name,  " &_
					"cs.serv_loc_contact_type_lcode, " &_
					"l.serv_loc_contact_type_desc, " &_
					"cs.contact_priority " &_
				"from crp.contact c, " &_
					"crp.service_location_contact cs, " &_
					"crp.service_location s, " &_
					"crp.lcode_serv_loc_contact_type l  " &_
				"where c.contact_id = cs.contact_id " &_
				"and	  cs.service_location_id = s.service_location_id " &_
				"and	  cs.serv_loc_contact_type_lcode = l.serv_loc_contact_type_lcode " &_
				"and c.contact_id = " & lngContactID &_
				"and cs.record_status_ind = 'A' " & _
				"order by s.service_location_name, l.serv_loc_contact_type_lcode, cs.contact_priority"

		Dim objRsServLocContacts
		set objRsServLocContacts = Server.CreateObject("ADODB.Recordset")
		objRsServLocContacts.CursorLocation = adUseClient
		objRsServLocContacts.Open strSQL,objConn
		if err then
			DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
		end if
		set objRsServLocContacts.ActiveConnection = nothing

	end if

	'get list items

	'Preferred Contact Method
	dim objRsPrefContMeth
	strSQL = "select contact_method_code, contact_method_desc from crp.contact_method where record_status_ind = 'A' order by contact_method_desc"
	set objRsPrefContMeth = Server.CreateObject("ADODB.Recordset")
	objRsPrefContMeth.CursorLocation = adUseClient
	objRsPrefContMeth.Open strSQL,objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
	end if
	set objRsPrefContMeth.ActiveConnection = nothing

	'Available Schedule
	dim objRsAvailSched
	strSQL = "select schedule_id, schedule_name from crp.schedule where record_status_ind = 'A' order by schedule_name"
	set objRsAvailSched = Server.CreateObject("ADODB.Recordset")
	objRsAvailSched.CursorLocation = adUseClient
	objRsAvailSched.Open strSQL,objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
	end if
	set objRsAvailSched.ActiveConnection = nothing

	'Language Preference
	dim rsLangPref
	strSQL = "select language_preference_lcode, language_preference_desc from crp.lcode_language_preference where record_status_ind = 'A' order by language_preference_desc"
	set rsLangPref = Server.CreateObject("ADODB.Recordset")
	rsLangPref.CursorLocation = adUseClient
	rsLangPref.Open strSQL,objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
	end if
	set rsLangPref.ActiveConnection = nothing
	'Response.write strSQL & "<BR>"
%>
<HTML>

<HEAD>
	<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
	<TITLE>SMA - Contact</TITLE>
	<script type = "text/Javascript" SRC="GeneralJavaFunctions.js"></script>
	<script type="text/javascript" SRC="AccessLevels.js"></script>
	<SCRIPT type = "text/javascript" LANGUAGE=javascript>

	<!--
	var bolNeedToSave = false ;
	var strWinMessage = "<%=strWinMessage%>";
	var intAccessLevel = "<%=intAccessLevel%>";
	//*********************************************************************************
	//set title
setPageTitle("SMA - Contacts");

	//*********************************************************************************

	function fct_selNavigate(){
	//***************************************************************************************************
	// Function:	fct_selNavigate															            *
	// Purpose:		To display the page selected by the user from Quick Navigation drop-down box. The	*
	//              To pass values to detail page use querystring; to list page use cookie.             *
	// Created By:	Nancy Mooney 08/31/2000															    *
	//																									*																				*
	//***************************************************************************************************

		var strPageName;
		var strLName;
		var strFName;
		var lngCustomerID;

		//alert ("In selNavigate");
		strPageName = document.frmContactDetail.selNavigate.item(document.frmContactDetail.selNavigate.selectedIndex).value ;

		switch(strPageName){
			case 'Address':
				//to detail
				document.frmContactDetail.selNavigate.selectedIndex=0;
				lngAddressID = document.frmContactDetail.hdnAddressID.value;
				if (lngAddressID != "" ) {
					self.location.href = "AddressDetail.asp?AddressID=" + lngAddressID; }
				else {
					alert('Cannot move to address as contact does not have an address.') ; }

				break;
			case 'Cust':
				//to detail
				document.frmContactDetail.selNavigate.selectedIndex=0;
				lngCustomerID = document.frmContactDetail.hdnCustomerID.value;
				self.location.href = "CustDetail.asp?CustomerID=" + lngCustomerID;
				break;
			case 'ContactRole':
				//to a list
				document.frmContactDetail.selNavigate.selectedIndex=0;
				strLName = document.frmContactDetail.txtLName.value;
				if(strLName != ""){SetCookie("LName", strLName)};
				strFName = document.frmContactDetail.txtFName.value;
				if(strFName != ""){SetCookie("FName", strFName)};
				self.location.href = "SearchFrame.asp?fraSrc=" + strPageName  ;
				break;
			case 'DEFAULT':
				//do nothing
		}
	}

//*********************************************************************************

	function fct_onChange(){
		bolNeedToSave = true;
	}

//*********************************************************************************

	function btnSave_onclick(){

		if (((intAccessLevel & intConst_Access_Update) == intConst_Access_Update) || ((intAccessLevel & intConst_Access_Create) == intConst_Access_Create)){

			//check required fields
			if (document.frmContactDetail.txtFName.value == "" ) {
				alert('Missing required field. Please enter a First Name');
				document.frmContactDetail.txtFName.focus();
				return(false);
			}
			if (document.frmContactDetail.txtLName.value == "" ) {
				alert('Missing required field. Please enter a Last Name');
				document.frmContactDetail.txtLName.focus();
				return(false);
			}
			if (document.frmContactDetail.txtCustomerName.value == "") {
				alert('Missing required field. Please enter a Works For Customer Name using the lookup button (...)');
				document.frmContactDetail.btnCustomerLookup.focus();
				return (false);
			}
			if (document.frmContactDetail.selLangPref.value == "" ) {
				alert('Missing required field. Please select a language preference using the drop down list.');
				document.frmContactDetail.selLangPref.focus();
				return(false);
			}

			//check that all phone numbers consist of numbers & 10 chars
			//work phone
			var WPhone;
			WPhone = document.frmContactDetail.txtWArea.value + document.frmContactDetail.txtWMid.value + document.frmContactDetail.txtWEnd.value;
			if (WPhone.length > 0) {
				if (isNaN(WPhone)) {
					alert('Work phone number must consist of digits only.');
					document.frmContactDetail.txtWArea.focus();
					return false;
				}
				if (WPhone.length < 10) {
					alert('Work phone number must consist of 10 digits (###) ###-####.');
					document.frmContactDetail.txtWArea.focus();
					return false;
				}
			}
			//work phone ext
			var WExt;
			WExt = document.frmContactDetail.txtExt.value;
			if (WExt.length > 0){
				if (isNaN(WExt)) {
					alert('Work phone extension must consist of digits only.');
					document.frmContactDetail.txtExt.focus();
					return false;
				}
			}
			//Cell phone
			var CPhone;
			CPhone = document.frmContactDetail.txtCArea.value + document.frmContactDetail.txtCMid.value + document.frmContactDetail.txtCEnd.value;
			if (CPhone.length > 0) {
				if (isNaN(CPhone)) {
					alert('Cell phone number must consist of digits only.');
					document.frmContactDetail.txtCArea.focus();
					return false;
				}
				if (CPhone.length < 10) {
					alert('Cell phone number must consist of 10 digits (###) ###-####.');
					document.frmContactDetail.txtCArea.focus();
					return false;
				}
			}
			//pager
			var PPhone;
			PPhone = document.frmContactDetail.txtPArea.value + document.frmContactDetail.txtPMid.value + document.frmContactDetail.txtPEnd.value;
			if (PPhone.length > 0) {
				if (isNaN(PPhone)) {
					alert('Pager number must consist of digits only.');
					document.frmContactDetail.txtPArea.focus();
					return false;
				}
				if (PPhone.length < 10) {
					alert('Pager number must consist of 10 digits (###) ###-####.');
					document.frmContactDetail.txtPArea.focus();
					return false;
				}
			}
			//Fax
			var FPhone;
			FPhone = document.frmContactDetail.txtFArea.value + document.frmContactDetail.txtFMid.value + document.frmContactDetail.txtFEnd.value;
			if (FPhone.length > 0) {
				if (isNaN(FPhone)) {
					alert('Fax number must consist of digits only.');
					document.frmContactDetail.txtFArea.focus();
					return false;
				}
				if (FPhone.length < 10) {
					alert('Fax number must consist of 10 digits (###) ###-####.');
					document.frmContactDetail.txtFArea.focus();
					return false;
				}
			}
			//home phone
			var HPhone;
			HPhone = document.frmContactDetail.txtHArea.value + document.frmContactDetail.txtHMid.value + document.frmContactDetail.txtHEnd.value;
			//alert(HPhone.length);
			if (HPhone.length > 0) {
				if (isNaN(HPhone)) {
					alert('Home phone number must consist of digits only.');
					document.frmContactDetail.txtHArea.focus();
					return false;
				}
				if (HPhone.length < 10) {
					alert('Home phone number must consist of 10 digits (###) ###-####.');
					document.frmContactDetail.txtHArea.focus();
					return false;
				}
			}

			//check that if preferred contact method has value, then that method has value
			var strPCM;
			strPCM = document.frmContactDetail.selPrefContMeth.value;
			if (strPCM != "") {
				switch (strPCM){
					case 'WORKPH':
						if (WPhone.length == 0){
							alert('Work phone has been selected as the preferred contact method. Please provide the work phone number or deselect the preferred contact method.');
							document.frmContactDetail.txtWArea.focus();
							return (false);
						}
						break;
					case 'EMAIL':
						if (document.frmContactDetail.txtEmail.value == "") {
							alert('Email has been selected as the preferred contact method. Please provide the email address or deselect the preferred contact method.');
							document.frmContactDetail.txtEmail.focus();
							return (false);
						}
						break;
					case 'CELL':
						if (CPhone.length == 0) {
							alert('Cellular has been selected as the preferred contact method. Please provide the cell phone number or deselect the preferred contact method.');
							document.frmContactDetail.txtCArea.focus();
							return (false);
						}
						break;
					case 'HOMEPH':
						if (HPhone.length == 0) {
							alert('Home phone has been selected as the preferred contact method. Please provide the home phone number or deselect the preferred contact method.');
							document.frmContactDetail.txtHArea.focus();
							return (false);
						}
						break;
					case 'FAX':
						if (FPhone.length == 0) {
							alert('Fax has been selected as the preferred contact method. Please provide the fax number or deselect the preferred contact method.');
							document.frmContactDetail.txtFArea.focus();
							return (false);
						}
						break;
					case 'PAGER':
						if (PPhone.length == 0) {
							alert('Pager has been selected as the preferred contact method. Please provide the pager number or deselect the preferred contact method.');
							document.frmContactDetail.txtPArea.focus();
							return (false);
						}
						break;
					case 'WEB':
						if (document.frmContactDetail.txtWebSite.value == "") {
							alert('Web Site Address has been selected as the preferred contact method. Please provide the web site address or deselect the preferred contact method.');
							document.frmContactDetail.txtPArea.focus();
							return (false);
						}
						break;
					default:
						//do nothing
						break;
				}//end switch
			}//end if

			//if Receive publications is checked - must enter address
			//alert(document.frmContactDetail.chkReceivePub.checked);
			if (document.frmContactDetail.chkReceivePub.checked == true) {
				if (document.frmContactDetail.hdnAddressID.value == ""){
					alert("Receive Publications has been checked. Please select an address for this customer using the lookup button (...).");
					return false;
				}
			}

			//comments
			var strComments = document.frmContactDetail.txtComments.value;
				if (strComments.length > 2000) {
					alert('Comments can be at most 2000 characters.\n\nYou entered ' + strComments.length + ' character(s).');
					document.frmContactDetail.txtComments.focus();
				return false;
			}

			bolNeedToSave = false; //bypass message asking if you want to save on window_onunload
			document.frmContactDetail.txtFrmAction.value = 'SAVE';
			document.frmContactDetail.submit();
			return(true);
		}
		else {
			alert('Access denied. Please contact your system administrator.');
			return(false);
		}
	}

//*********************************************************************************

	function body_onbeforeunload(){

		document.frmContactDetail.btnSave.focus();
		//alert("body_onbeforeunload: bolNeedToSave: " + bolNeedToSave);
		if ((bolNeedToSave == true) && ( ((intAccessLevel & intConst_Access_Create) == intConst_Access_Create)||((intAccessLevel & intConst_Access_Update) == intConst_Access_Update) )){
					event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
		}
	}

//*********************************************************************************

	function ClearStatus() {
		window.status = "";
	}

//*********************************************************************************

	function DisplayStatus(strWinStatus) {
		window.status=strWinStatus;
		setTimeout('ClearStatus()',"<%=intConst_MessageDisplay%>");
	}

//*********************************************************************************

	function btnAddressLookup_onclick()
	{
		if (document.frmContactDetail.txtCustomerName.value != "" ) {
			SetCookie("CustomerName",document.frmContactDetail.txtCustomerName.value ) ;
		}
		SetCookie("WinName", 'Popup');
		bolNeedToSave = true;
		window.open('SearchFrame.asp?fraSrc=Address', 'Popup', 'top=50, left=100, WIDTH=800, HEIGHT=600' ) ;
	}

//*********************************************************************************

	function btnAddressClear_onClick(){
		document.frmContactDetail.textAddress.value = "";
		document.frmContactDetail.hdnAddressID.value = "";
	}

//*********************************************************************************

	function btnCustomerLookup_onclick(CustService)
	{
		if (document.frmContactDetail.txtCustomerName.value != ""){
			 SetCookie("CustomerName", document.frmContactDetail.txtCustomerName.value);
		}
		SetCookie("WinName", 'Popup');
		SetCookie("ServiceEnd", CustService);
		bolNeedToSave = true;
		window.open('SearchFrame.asp?fraSrc=Cust', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
	}

//*********************************************************************************

	function btnReferences_onclick(){
		var strOwner = 'CRP';
		var strTableName = 'CONTACT';
		var strRecordID = document.frmContactDetail.hdnContactID.value;
		var URL;

		if (lngContactID = 0){
			alert("No references. This is a new record.");
		}
		else{
			URL='Dependency.asp?Owner='+strOwner+'&TableName='+strTableName+'&RecordID='+strRecordID;
			window.open(URL,'Popup','top=100,left=100,WIDTH=500,HEIGHT=300');
		}
	}

//*********************************************************************************

	function btnDelete_onclick()
	{
		var lngContactID = document.frmContactDetail.hdnContactID.value;
		var strUpdateDateTime = document.frmContactDetail.hdnUpdateDateTime.value;

		if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {alert('Access denied. Please contact your system administrator.'); return;}
		if(confirm('Do you really want to delete this contact?')){
			document.location = "ContactDetail.asp?txtFrmAction=DELETE&ContactID="+lngContactID+"&UpdateDateTime="+strUpdateDateTime;
		}
	}

//*********************************************************************************

	function btnReset_onclick()
	{
		if(confirm('All changes will be lost. Do you really want to reset the page?')){
			bolNeedToSave = false;
			document.location = "ContactDetail.asp?ContactID=<%=lngContactID%>";
		}
	}

//*********************************************************************************

	function btnNew_onclick()
	{
		if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {alert('Access denied. Please contact your system administrator.'); return;}
		self.document.location.href="ContactDetail.asp?ContactID=0";
	}

//*********************************************************************************
//-->end hide script

</SCRIPT>

</HEAD>

<BODY onLoad="DisplayStatus(strWinMessage);" onbeforeunload="body_onbeforeunload();" >
<FORM name=frmContactDetail action="ContactDetail.asp" method="POST" >

<!-- hidden variables -->

	<INPUT name=hdnContactID type=hidden value=<%if lngContactID <> 0 then Response.Write """"&rsContact("contact_id")&"""" else Response.Write """""" end if%>>
	<INPUT name=hdnCustomerID type=hidden value=<%if lngContactID <> 0 then Response.Write """"&rsContact("work_for_customer_id")&"""" else Response.Write """""" end if%>>
	<INPUT name=hdnAddressID type=hidden value=<%if lngContactID <> 0 then Response.Write """"&rsContact("address_id")&"""" else Response.Write null end if%>>
	<INPUT name=hdnUpdateDateTime type=hidden value=<%if lngContactID <> 0 then Response.Write """"&rsContact("last_update_date_time")&"""" else Response.Write """""" end if%>>
	<INPUT name=txtFrmAction type=hidden value="">

	<table border=0 width=100%>
	<thead>
		<tr><td align=left colspan=3 >Contact Detail</td>
			<td><SELECT align=right  valign=top name=selNavigate tabindex=38 LANGUAGE=javascript onchange="fct_selNavigate();" <%if lngContactID = 0 then Response.Write " disabled " end if%>>
					<OPTION value="DEFAULT">Quickly Goto ...</OPTION>
					<OPTION value="Address">Address</OPTION>
					<OPTION value="Cust">Customer</OPTION>
					<OPTION value="ContactRole" >Contact Role</OPTION>
			</SELECT></td>
	</thead>
	<tr><th align="left" colspan=4><%if lngContactID <> 0 then Response.Write ""& rsContact("contact_name")& "" else Response.Write null end if%></th></tr>
	<tr>
		<td  align=right >Title&nbsp;</td>
		<td colspan=2 align=left >
			<SELECT name=selTitle tabindex=1 onChange="fct_onChange();" >
				<OPTION value=></option>
				<OPTION <%if lngContactID <> 0 then if rsContact("name_prefix") = "Miss" then Response.Write " selected " end if end if%> value="Miss">Miss </option>
				<OPTION <%if lngContactID <> 0 then if rsContact("name_prefix") = "Mrs." then Response.Write " selected " end if end if%> value="Mrs.">Mrs.</option>
				<OPTION <%if lngContactID <> 0 then if rsContact("name_prefix") = "Ms." then Response.Write " selected " end if end if%> value="Ms.">Ms.</option>
				<OPTION <%if lngContactID <> 0 then if rsContact("name_prefix") = "Mr." then Response.Write " selected " end if end if%> value="Mr.">Mr.</option>
				<OPTION <%if lngContactID <> 0 then if rsContact("name_prefix") = "Dr." then Response.Write " selected " end if end if%> value="Dr.">Dr.</option>
				<OPTION <%if lngContactID <> 0 then if rsContact("name_prefix") = "Capt." then Response.Write " selected " end if end if%> value="Capt.">Capt.</option>
				<OPTION <%if lngContactID <> 0 then if rsContact("name_prefix") = "Prof." then Response.Write " selected " end if end if%> value="Prof.">Prof.</option>
			</SELECT>&nbsp;&nbsp;&nbsp;&nbsp;
			First Name<font color=red>*</font>
			<INPUT name=txtFName tabindex=2 maxlength=20 onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&rsContact("first_name")&"""" else Response.Write """""" end if%>></input>&nbsp;&nbsp;&nbsp;&nbsp;
			Last Name<font color=red>*</font>
			<INPUT name=txtLName tabindex=3 size=20 maxlength=20 onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&rsContact("last_name")&"""" else Response.Write """""" end if%>></input>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			Middle Name
		</td>
		<td align=left><INPUT name=txtMName tabindex=4 size=7 maxlength=7 onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&rsContact("middle_name")&"""" else Response.Write """""" end if%>></input></td>
	<TR>
		<td align="right" >Works For<font color=red>*</font></td>
		<td align="left" >
			<input type="text" name="txtCustomerName" size=50 maxlength="50" disabled onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&strCustomerName&"""" else Response.Write """""" end if%>></input>
			<INPUT align=right type="button" name=btnCustomerLookup  value="..." tabindex=5 onclick="return btnCustomerLookup_onclick('C')" >
		</td>
		<td  align="right">Language Preference<font color=red>*</font></td>
		<td >
			<SELECT name=selLangPref tabindex=11 onChange="fct_onChange();">
				<option value=></option>
				<%
				while not rsLangPref.EOF
					Response.write "<OPTION "
					if lngContactID <> 0 then
						if rsLangPref("language_preference_lcode") = rsContact("language_preference_lcode") then Response.Write " selected " end if
					else
						if rsLangPref("language_preference_lcode") = "EN" then Response.Write " selected " end if
					end if
					Response.Write "value=" & rsLangPref(0)& ">" & routineHtmlString(rsLangPref(1)) & "</OPTION>" & vbCrLf
					rsLangPref.MoveNext
				wend
				rsLangPref.Close
				set rsLangPref = nothing
				%>
			</SELECT>
		</td>
	</tr>
	<tr>
		<td align="right" >Position&nbsp;</td>
		<td ><INPUT name=txtPosition maxlength="50" size="50" tabindex=6 onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&rsContact("position_title")&"""" else Response.Write """""" end if%>></input></td>
		<td align="right" >Preferred Contact Method&nbsp;</td>
		<td >
			<SELECT name=selPrefContMeth tabindex=12 onChange="fct_onChange();" >
				<option value=></option>
				<%while not objRsPrefContMeth.EOF
					Response.Write "<OPTION "
					if lngContactID <> 0 then
						if objRsPrefContMeth("contact_method_code") = rsContact("prefercontactmethodcode") then Response.Write " selected " end if
					else
						if objRsPrefContMeth("contact_method_code") = "WORKPH" then Response.Write " selected " end if
					end if
					Response.Write "value=" & objRsPrefContMeth(0)& ">" & routineHtmlString(objRsPrefContMeth(1)) & "</OPTION>" & vbCrLf
					objRsPrefContMeth.MoveNext
				wend
				objRsPrefContMeth.Close
				set objRsPrefContMeth = nothing
				%>
			</SELECT>
		</td>
	</tr>
	<tr>
		<td align="right" >Address&nbsp;</td>
		<td rowspan=3 >
			<TEXTAREA cols=25 name=textAddress rows=4 disabled style="width: 360" onChange="fct_onChange();"><%if lngContactID <> 0 then Response.Write rsContact("contact_address") else Response.Write null end if%></TEXTAREA>
			<INPUT align=right name=btnAddressLookup type=button tabindex=7 value=... LANGUAGE=javascript onclick="return btnAddressLookup_onclick()">
			<INPUT align=right name=btnAddressClear type=button tabindex=8 value="X" LANGUAGE=javascript onclick = "btnAddressClear_onClick()">
		</td>
		<td align="right" >Availability </td>
		<td >
			<SELECT name=selAvailSched tabindex=14 onChange="fct_onChange();" >
				<option value=></option>
				<% while not objRsAvailSched.EOF
					Response.Write "<OPTION "
					if lngContactID <> 0 then
						if rsContact("availablescheduleid") <> "" then
							if CLng(objRsAvailSched("schedule_id"))= CLng(rsContact("availablescheduleid")) then
								Response.Write " selected "
							end if
						end if
					end if
					Response.Write "value=" & objRsAvailSched(0) & ">" & routineHtmlString(objRsAvailSched(1)) & "</OPTION>" & vbCrLf
					objRsAvailSched.MoveNext
				wend
				objRsAvailSched.Close
				set objRsAvailSched = nothing
				%>
			</SELECT>
		</td>
	</tr>
	<tr>
		<td ></td>
		<td align=right >TELUS Staff</td>
		<td><input name=chkStaffFlag type=checkbox tabindex=15 onChange="fct_onChange();"
			<%if lngContactID <> 0 then
				if rsContact("staff_flag") = "Y" then
					Response.Write " checked "
				end if
			end if %>>
		</td>
	</tr>
	<tr>
		<TD ></td>
		<td align=right>Receive Publications&nbsp;</td>
		<td align=left><INPUT name=chkReceivePub type=checkbox  tabindex=16 onChange="fct_onChange();"
			<%if lngContactID <> 0 then
				if rsContact("receive_publications_flag") = "Y" then
					Response.Write " checked "
				end if
			end if%>>
		</td>
	<tr>
		<td align="right" >Email&nbsp;</td>
		<td><INPUT name=txtEmail size=80 maxlength=80 style="WIDTH: 16cm"  onChange="fct_onChange();" tabindex=9 value=<%if lngContactID <> 0 then Response.Write """"&rsContact("email_address")&"""" else Response.Write """""" end if%>></input>
		<td align="right">Work&nbsp;</td>
		<td align="left" >(<INPUT name=txtWArea size=3 maxlength=3 tabindex=17 onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&strWkArea&"""" else Response.Write """""" end if%>></input>)
			<INPUT name=txtWMid size=3 maxlength=3 tabindex=18 onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&strWkMid&"""" else Response.Write """""" end if%>></input>
			-&nbsp;<INPUT name=txtWEnd size=4 maxlength=4 tabindex=19 onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&strWkEnd&"""" else Response.Write """""" end if%>></input>&nbsp;Ext
			<INPUT name=txtExt size=10 maxlength=10 tabindex=20 onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&rsContact("work_number_ext")&"""" else Response.Write """""" end if%>></input></td>
		</tr>
	<tr>
		<td align=right >Web Site&nbsp;</td>
		<td align=left><INPUT name=txtWebSite size=50 maxlength=50 tabindex=10 onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&rsContact("web_site_url")&"""" else Response.Write """""" end if%>></input></td>
		<td align=right>Cell&nbsp;</td>
		<td align=left>
			(<INPUT name=txtCArea size=3 maxlength=3 tabindex=21 onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&strClArea&"""" else Response.Write """""" end if%>></input>)
			<INPUT name=txtCMid size=3 maxlength=3 tabindex=22 onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&strClMid&"""" else Response.Write """""" end if%>></input>
			-&nbsp;<INPUT name=txtCEnd size=4 maxlength=4 tabindex=23 onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&strClEnd&"""" else Response.Write """""" end if%>></input></td>
	</tr>
	<tr>
		<td  align=right valign="top" >Comments&nbsp;</td>
		<td  rowspan=3 valign="top"><TEXTAREA cols=25 name=txtComments rows=6 style="width: 360" tabindex=10 onChange="fct_onChange();"><%if lngContactID <> 0 then Response.Write rsContact("comments") else Response.Write null end if%></TEXTAREA></td>
		<td  align="right">Pager&nbsp;</td>
		<td  align="left" >
			(<INPUT name=txtPArea size=3 maxlength=3 tabindex=24 onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&strPgArea&"""" else Response.Write """""" end if%>></input>)
			<INPUT name=txtPMid size=3 maxlength=3 tabindex=25 onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&strPgMid&"""" else Response.Write """""" end if%>></input>
			-&nbsp;<INPUT name=txtPEnd size=4 maxlength=4 tabindex=26 onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&strPgEnd&"""" else Response.Write """""" end if%>></input></td>
	</tr>
	<tr>
		<td ></td>
		<td  align="right">Fax&nbsp;</td>
		<td  align="left" >
			(<INPUT name=txtFArea size=3 maxlength=3 tabindex=27 onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&strFxArea&"""" else Response.Write """""" end if%>></input>)
			<INPUT name=txtFMid size=3 maxlength=3 tabindex=28 onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&strFxMid&"""" else Response.Write """""" end if%>></input>
			-&nbsp;<INPUT name=txtFEnd size=4 maxlength=4 tabindex=29 onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&strFxEnd&"""" else Response.Write """""" end if%>></input>
		</td>
	</tr>
	<tr>
		<td ></td>
	    <td  align="right">Home&nbsp;</td>
		<td  align="left" >(<INPUT name=txtHArea size=3 maxlength=3 tabindex=30 onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&strHmArea&"""" else Response.Write """""" end if%>></input>)
			<INPUT name=txtHMid size=3 maxlength=3 tabindex=31 onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&strHmMid&"""" else Response.Write """""" end if%>></input>
			-&nbsp;<INPUT name=txtHEnd size=4 maxlength=4 tabindex=32 onChange="fct_onChange();" value=<%if lngContactID <> 0 then Response.Write """"&strHmEnd&"""" else Response.Write """""" end if%>></input>
		</td>
	</tr>
	<tr>
		<td  align=right valign="top" >Responsibility&nbsp;</td>
		<td  valign="top"><INPUT name=txtResponsibility size=50 maxlength=50 tabindex=10 onChange="fct_onChange();"value=<%if lngContactID <> 0 then Response.Write """"&rsContact("responsibility")&"""" else Response.Write """""" end if%>></input></td>
		<td  align="right">PIN&nbsp;</td>
		<td  align="left" ><INPUT name=chkPINAccess type=checkbox  tabindex=16 onChange="fct_onChange();"
			<%if lngContactID <> 0 then
				if rsContact("PIN_access") = "Y" then
					Response.Write " checked "
				end if
			end if%>>
		</td>
	</tr>
<tfoot>
	<tr>
		<td  align=right colspan=4>
			<INPUT name=btnReferences type=button value=References style="WIDTH: 2cm" tabindex=33 onclick="return btnReferences_onclick();">&nbsp;&nbsp;
			<INPUT name=btnDelete	  type=button value=Delete	   style="WIDTH: 2cm" tabindex=34 onclick="return btnDelete_onclick();">&nbsp;&nbsp;
			<INPUT name=btnReset	  type=button value=Reset	   style="WIDTH: 2cm" tabindex=35 onclick="return btnReset_onclick();">&nbsp;&nbsp;
			<INPUT name=btnNew		  type=button value=New		   style="WIDTH: 2cm" tabindex=36 onclick="return btnNew_onclick();">&nbsp;&nbsp;
			<INPUT name=btnSave		  type=button value=Save	   style="WIDTH: 2cm" tabindex=37 onclick="return btnSave_onclick();">&nbsp;&nbsp;
		</td>
	</tr>
</tfoot>
</table>


<% if lngContactID <> 0 then
	IF not objRsCustContacts.eof then %>
	<TABLE border=1 cellPadding=2 cellSpacing=0 width=100%>
		<THEAD>
			<TR><td align=left colspan=4>Customer Contact Role(s)</td></tr>
			<TR>
				<TH align=left>Customer Name</TH>
				<TH align=left>Role</TH>
				<TH align=left>Role Description</TH>
				<TH align=left>Priority</TH></TR>
		</THEAD>
		<TBODY>
		<% do while not objRsCustContacts.eof %>
			<TR>
				<TD><%=objRsCustContacts(0)%></TD>
				<TD><%=objRsCustContacts(1)%></TD>
				<TD><%=objRsCustContacts(2)%></TD>
				<TD><%=objRsCustContacts(3)%></TD>
			</TR>
		<% objRsCustContacts.Movenext
		loop %>
		</BODY>
	</TABLE>
	<% end if
	objRsCustContacts.close
	set objRsCustContacts = nothing

	IF not objRsCustServContacts.eof  then %>
		<TABLE border=1 cellPadding=2 cellSpacing=0 width=100%>
		<THEAD>
			<TR><td align=left colspan=4>Customer Service Contact Role(s)</td></tr>
		    <TR>
				<TH align=left>Customer Service Name</TH>
				<TH align=left>Role</TH>
				<TH align=left>Role Description</TH>
				<TH align=left>Pirority</TH></TR>
		</THEAD>
		<TBODY>
		<% do while not objRsCustServContacts.eof %>
			<TR>
				<TD><%=objRsCustServContacts(0)%></TD>
				<TD><%=objRsCustServContacts(1)%></TD>
				<TD><%=objRsCustServContacts(2)%></TD>
				<TD><%=objRsCustServContacts(3)%></TD>
			</TR>
			<% objRsCustServContacts.Movenext
			loop %>
		</TBODY>
	</TABLE>
	<%end if
	objRsCustServContacts.close
	set objRsCustServContacts = nothing

	IF not objRsServLocContacts.eof  then %>
		<TABLE border=1 cellPadding=2 cellSpacing=0 width=100%>
		<THEAD>
			<TR><td align=left colspan=4>Service Location Contact Role(s)</td></tr>
			<TR>
				<TH align=left>Service Location Name</TH>
				<TH align=left>Role</TH>
				<TH align=left>Role Description</TH>
				<TH align=left>Priority</TH>
			</TR>
		</THEAD>
		<TBODY>
		<% do while not objRsServLocContacts.eof %>
		<TR>
			<TD><%=objRsServLocContacts(0)%></TD>
			<TD><%=objRsServLocContacts(1)%></TD>
			<TD><%=objRsServLocContacts(2)%></TD>
			<TD><%=objRsServLocContacts(3)%></TD>
		</TR>
		<% objRsServLocContacts.Movenext
		loop %>
		</TBODY>
	</TABLE>
	<% end if
	objRsServLocContacts.close
	set objRsServLocContacts = nothing 	%>
<% end if %>

<FIELDSET>
	<LEGEND align=right><B>Audit Information</B></LEGEND>
	<Div SIZE=8pt ALIGN=RIGHT>
		Record Status Indicator
		<INPUT align = left name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 18px"disabled value=<%if lngContactID <> 0 then Response.Write """"&rsContact("record_status_ind")&"""" else Response.Write """""" end if%>></input>&nbsp;&nbsp;&nbsp;
		Create Date
		<INPUT align = center name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 150px"disabled value=<%if lngContactID <> 0 then Response.Write """"&rsContact("create_date")&"""" else Response.Write """""" end if%>></input>&nbsp;
		Created By
		<INPUT align = right name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 200px"disabled value=<%if lngContactID <> 0 then Response.Write """"&rsContact("create_real_userid")&"""" else Response.Write """""" end if%>></input>&nbsp;&nbsp;<BR>
		Contact ID
		<INPUT align = left name=txtContactID type=text style="HEIGHT: 20px; WIDTH: 90px"disabled value=<%if lngContactID <> 0 then Response.Write """"&rsContact("contact_id")&"""" else Response.Write """""" end if%>></input>&nbsp;&nbsp;&nbsp;
		Update Date
		<INPUT align= center name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 150px"disabled value=<%if lngContactID <> 0 then Response.Write """"&rsContact("update_date")&"""" else Response.Write """""" end if%>></input>
		Updated By
		<INPUT align=right name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 200px"disabled value=<%if lngContactID <> 0 then Response.Write """"&rsContact("update_real_userid")&"""" else Response.Write """""" end if%>></input>&nbsp;&nbsp;
	</DIV>
</FIELDSET>

<% 'clean up ADO objects
	if lngContactID <> 0 then
		rsContact.close
		set rsContact = nothing
		objConn.close
		set objConn = nothing
	end if
%>
</FORM>
</BODY>
</HTML>
