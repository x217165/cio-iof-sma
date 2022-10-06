<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = True%>
<!--#include file = "smaConstants.inc"-->
<!--#include file = "smaProcs.inc"-->
<!--#include file = "databaseconnect.asp" -->
<!--
*************************************************************************************************
* Page name:	StaffDetail.asp																	*
* Purpose:		To display TAC Staff Information												*
*				Chosen via StaffList.asp														*
*																								*
* Created by:	Gilles Archer 10/27/2000														*
*																								*
*************************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       10-Apr-01	 DTy		Make employee number field mandatory and ensure
                                        only numeric entry of 8 positions are entered.
       19-Feb-02	 DTy		Increase email address size from 50 t0 60.
       16-Oct-07        ACheung         Add new columns "Responsibility" and "PIN"
*************************************************************************************************
-->
<%
Dim strWinMessage, strErrMessage, lIndex, arrNamePrefix
Dim objCommand, objRS, objLanguage, objDepartment, objRegion, objStaffStatus
Dim p_insert_userid, p_work_for_customer_id, p_contact_id, p_contact_name, p_last_update_dt, p_last_sec_update_dt, p_last_name, p_first_name, p_middle_name
Dim p_name_prefix, p_work_number, p_work_number_ext, p_home_number, p_cell_number, p_pager_number, p_fax_number
Dim p_email_address, p_position_title, p_address_id, p_client_rep_relationship
Dim p_default_noc_region, p_comments, p_language_preference_lcode, p_staff_flag
Dim p_employee_number, p_staff_status_lcode, p_department_id, p_old_userid, p_new_userid, p_password
Dim p_specific_location, p_manager_contact_id, p_phone_list_flag, p_def_remedy_support_group_id
Dim p_responsibility, p_pinaccess

Dim strContactID
Dim strWkArea, strWkMid, strWkEnd, strHmArea, strHmMid, strHmEnd
Dim strClArea, strClMid, strClEnd, strPgArea, strPgMid, strPgEnd
Dim strFxArea, strFxMid, strFxEnd

Dim strSQL, strFrom, strWhere
Dim intAccessLevel

	'Check user's rights
	intAccessLevel = CInt(CheckLogon(strConst_Security))
	If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly  Then
		DisplayError "BACK", "", 0, "ACCESS DENIED", "You do not have access to Staff Maintenance.  Please contact your system administrator."
	End If

	strWinMessage = ""
	arrNamePrefix = Split(strConst_NamePrefix, strDelimiter)
	strContactID = Request("ContactID")

	p_insert_userid = Session("username")

	If IsNumeric(Request.Form("hdnContactID")) Then
		p_contact_id = CLng(Request.Form("hdnContactID"))
	Else
		p_contact_id = Null
	End If

	If IsNumeric(Request.Form("hdnCustomerID")) Then
		p_work_for_customer_id = CLng(Request.Form("hdnCustomerID"))
	Else
		p_work_for_customer_id = Null
	End If

	If Len(Request.Form("txtContactName")) <> 0 Then
		p_contact_name = Trim(Request.Form("txtContactName"))
	Else
		p_contact_name = Null
	End If

	If IsDate(Request.Form("hdnUpdateDateTime")) Then
		p_last_update_dt = CDate(Request.Form("hdnUpdateDateTime"))
	Else
		p_last_update_dt = Null
	End If

	If IsDate(Request.Form("hdnSecUpdateDateTime")) Then
		p_last_sec_update_dt = CDate(Request.Form("hdnSecUpdateDateTime"))
	Else
		p_last_sec_update_dt = Null
	End If

	If Len(Request.Form("txtNameLast")) <> 0 Then
		p_last_name = Trim(Request.Form("txtNameLast"))
	Else
		p_last_name = Null
	End If

	If Len(Request.Form("txtNameFirst")) <> 0 Then
		p_first_name = Trim(Request.Form("txtNameFirst"))
	ELse
		p_first_name = Null
	End If

	If Len(Request.Form("txtNameMiddle")) <> 0 Then
		p_middle_name = Trim(Request.Form("txtNameMiddle"))
	Else
		p_middle_name = Null
	End If

	If Len(Request.Form("selNamePrefix")) <> 0 Then
		p_name_prefix = Trim(Request.Form("selNamePrefix"))
	Else
		p_name_prefix = Null
	End If

	p_work_number = Trim(Request.Form("txtWArea") & Request.Form("txtWMid") & Request.Form("txtWEnd"))
	If (Not IsNumeric(p_work_number)) And (Len(p_work_number) <> 10) Then
		p_work_number = Null
	End If

	If Len(Request.Form("txtExt")) <> 0 Then
		p_work_number_ext = Trim(Request.Form("txtExt"))
	Else
		p_work_number_ext = Null
	End If

	p_home_number = Trim(Request.Form("txtHArea") & Request.Form("txtHMid") & Request.Form("txtHEnd"))
	If (Not IsNumeric(p_home_number)) And (Len(p_home_number) <> 10) Then
		p_home_number = Null
	End If

	p_cell_number = Trim(Request.Form("txtCArea") & Request.Form("txtCMid") & Request.Form("txtCEnd"))
	If (Not IsNumeric(p_cell_number)) And (Len(p_cell_number) <> 10) Then
		p_cell_number = Null
	End If

	p_pager_number = Trim(Request.Form("txtPArea") & Request.Form("txtPMid") & Request.Form("txtPEnd"))
	If (Not IsNumeric(p_pager_number)) And (Len(p_pager_number) <> 10) Then
		p_pager_number = Null
	End If

	p_fax_number = Trim(Request.Form("txtFArea") & Request.Form("txtFMid") & Request.Form("txtFEnd"))
	If (Not IsNumeric(p_fax_number)) And (Len(p_fax_number) <> 10) Then
		p_fax_number = Null
	End If

	If Len(Request.Form("txtEmail")) <> 0 Then
		p_email_address = Trim(Request.Form("txtEmail"))
	Else
		p_email_address = Null
	End If

	If Len(Request.Form("txtPosition")) <> 0 Then
		p_position_title = Trim(Request.Form("txtPosition"))
	Else
		p_position_title = Null
	End If

	If IsNumeric(Request.Form("hdnAddressID")) Then
		p_address_id = CLng(Request.Form("hdnAddressID"))
	Else
		p_address_id = Null
	End If

	If Len(Request.Form("txtClientRepRelationship")) <> 0 Then
		p_client_rep_relationship = Trim(Request.Form("txtClientRepRelationship"))
	Else
		p_client_rep_relationship = Null
	End If

'new Lynx responsibility and pinaccess
	If Len(Request.Form("txtResponsibility")) <> 0 Then
		p_responsibility = Trim(Request.Form("txtResponsibility"))
	Else
		p_responsibility = Null
	End If

	if Lcase(Request.Form("chkPINAccess")) = "on" then
		p_pinaccess = "Y"
	else
		p_pinaccess = "N"
	end if

	'If Len(Request.Form("txtPINAccess")) <> 0 Then
	'	p_pinaccess = Trim(Request.Form("txtPINAccess"))
	'End If

	If Len(Request.Form("txtComments")) <> 0 Then
		p_comments = Trim(Request.Form("txtComments"))
	Else
		p_comments = Null
	End If

	If Len(Request.Form("selLanguage")) <> 0 Then
		p_language_preference_lcode = Trim(Request.Form("selLanguage"))
	Else
		p_language_preference_lcode = Null
	End If


	'If Len(Request.Form("chkStaffFlag")) <> 0 Then
		p_staff_flag = "Y"
	'Else
	'	p_staff_flag = "N"
	'End If

	If Len(Request.Form("txtEmpNo")) <> 0 Then
		p_employee_number = Trim(Request.Form("txtEmpNo"))
		Do While Len(p_employee_number) < 8
			p_employee_number = "0" & p_employee_number
		Loop
	Else
		p_employee_number = Null
	End If

	If Len(Request.Form("selStaffStatus")) <> 0 Then
		p_staff_status_lcode = Trim(Request.Form("selStaffStatus"))
	Else
		p_staff_status_lcode = Null
	End If

	If IsNumeric(Request.Form("selDepartment")) Then
		p_department_id = CLng(Request.Form("selDepartment"))
	Else
		p_department_id = Null
	End If

	If Len(Request.Form("hdnUserID")) <> 0 Then
		p_old_userid = LCase(Trim(Request.Form("hdnUserID")))
	Else
		p_old_userid = Null
	End If

	If Len(Request.Form("txtUserID")) <> 0 Then
		p_new_userid = LCase(Trim(Request.Form("txtUserID")))
	Else
		p_new_userid = Null
	End If

	If Len(Request.Form("txtLocation")) <> 0 Then
		p_specific_location = Trim(Request.Form("txtLocation"))
	Else
		p_specific_location = Null
	End If

	If IsNumeric(Request.Form("hdnManagerContactID")) Then
		p_manager_contact_id = CLng(Request.Form("hdnManagerContactID"))
	Else
		p_manager_contact_id = Null
	End If

	If Len(Request.Form("chkPhoneList")) <> 0 Then
		p_phone_list_flag = "Y"
	Else
		p_phone_list_flag = "N"
	End If

	If Len(Request.Form("selNOCRegion")) <> 0 Then
		p_default_noc_region = Trim(Request.Form("selNOCRegion"))
	Else
		p_default_noc_region = Null
	End If

	If Len(Request.Form("txtPassword")) <> 0 Then
		p_password = Trim(Request.Form("txtPassword"))
	Else
		p_password = Null
	End If

	If Len(Request.Form("selRemedySupport")) <> 0 Then
		p_def_remedy_support_group_id = Request.Form("selRemedySupport")
	Else
		p_def_remedy_support_group_id = Null
	End If

	Select Case UCase(Request("hdnFrmAction"))
		Case "SAVE"
			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc

			If IsNumeric(p_contact_id) Then	'Save existing Service Type
				If (intAccessLevel And intConst_Access_Update) <> intConst_Access_Update Then
					DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update Staff. Please contact your system administrator"
				End If

				objCommand.CommandText = "sma_sp_userid.spk_security_inter.sp_contact_update"
				objCommand.Parameters.Append objCommand.CreateParameter("p_insert_userid", adVarChar, adParamInput, 20, p_insert_userid)
				objCommand.Parameters.Append objCommand.CreateParameter("p_contact_id", adNumeric, adParamInput, , p_contact_id)
				objCommand.Parameters.Append objCommand.CreateParameter("p_contact_name", adVarChar, adParamInput, 50, p_contact_name)
				objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , p_last_update_dt)
				objCommand.Parameters.Append objCommand.CreateParameter("p_last_sec_update_dt", adDBTimeStamp, adParamInput, , p_last_sec_update_dt)
				objCommand.Parameters.Append objCommand.CreateParameter("p_old_userid", adVarChar, adParamInput, 8, p_old_userid)
				objCommand.Parameters.Append objCommand.CreateParameter("p_new_userid", adVarChar, adParamInput, 8, p_new_userid)
				objCommand.Parameters.Append objCommand.CreateParameter("p_password", adVarChar, adParamInput, 10, p_password)
				objCommand.Parameters.Append objCommand.CreateParameter("p_last_name", adVarChar, adParamInput, 20, p_last_name)
				objCommand.Parameters.Append objCommand.CreateParameter("p_first_name", adVarChar, adParamInput, 20, p_first_name)
				objCommand.Parameters.Append objCommand.CreateParameter("p_middle_name", adVarChar, adParamInput, 7, p_middle_name)
				objCommand.Parameters.Append objCommand.CreateParameter("p_name_prefix", adVarChar, adParamInput, 6, p_name_prefix)
				objCommand.Parameters.Append objCommand.CreateParameter("p_work_number", adVarChar, adParamInput, 24, p_work_number)
				objCommand.Parameters.Append objCommand.CreateParameter("p_work_number_ext", adVarChar, adParamInput, 10, p_work_number_ext)
				objCommand.Parameters.Append objCommand.CreateParameter("p_home_number", adVarChar, adParamInput, 24, p_home_number)
				objCommand.Parameters.Append objCommand.CreateParameter("p_cell_number", adVarChar, adParamInput, 24, p_cell_number)
				objCommand.Parameters.Append objCommand.CreateParameter("p_pager_number", adVarChar, adParamInput, 24, p_pager_number)
				objCommand.Parameters.Append objCommand.CreateParameter("p_fax_number", adVarChar, adParamInput, 24, p_fax_number)
				objCommand.Parameters.Append objCommand.CreateParameter("p_email_address", adVarChar, adParamInput, 60, p_email_address)
				objCommand.Parameters.Append objCommand.CreateParameter("p_web_site_url", adVarChar, adParamInput, 60, null)
				objCommand.Parameters.Append objCommand.CreateParameter("p_position_title", adVarChar, adParamInput, 50, p_position_title)
				objCommand.Parameters.Append objCommand.CreateParameter("p_address_id", adNumeric, adParamInput, , p_address_id)
				objCommand.Parameters.Append objCommand.CreateParameter("p_client_rep_relationship", adVarChar, adParamInput, 10, p_client_rep_relationship)
				objCommand.Parameters.Append objCommand.CreateParameter("p_receive_publications_flag", adChar, adParamInput, 1, "N")
				objCommand.Parameters.Append objCommand.CreateParameter("p_prefercontactmethodcode", adVarChar, adParamInput, 6, null)
				objCommand.Parameters.Append objCommand.CreateParameter("p_availablescheduleid", adNumeric, adParamInput, , null)
				objCommand.Parameters.Append objCommand.CreateParameter("p_comments", adVarChar, adParamInput, 2000, p_comments)
				objCommand.Parameters.Append objCommand.CreateParameter("p_responsibility", adVarChar, adParamInput, 50, p_responsibility)
				objCommand.Parameters.Append objCommand.CreateParameter("p_pinaccess", adChar, adParamInput, 1, p_pinaccess)
				objCommand.Parameters.Append objCommand.CreateParameter("p_language_preference_lcode", adChar, adParamInput, 2, p_language_preference_lcode)
				objCommand.Parameters.Append objCommand.CreateParameter("p_t2_group_flag", adChar, adParamInput, 1, "N")
				objCommand.Parameters.Append objCommand.CreateParameter("p_staff_flag", adChar, adParamInput, 1, p_staff_flag)
				objCommand.Parameters.Append objCommand.CreateParameter("p_employee_number", adVarChar, adParamInput, 8, p_employee_number)
				objCommand.Parameters.Append objCommand.CreateParameter("p_staff_status_lcode", adVarChar, adParamInput, 15, p_staff_status_lcode)
				objCommand.Parameters.Append objCommand.CreateParameter("p_department_id", adNumeric, adParamInput, , p_department_id)
				objCommand.Parameters.Append objCommand.CreateParameter("p_specific_location", adVarChar, adParamInput, 50, p_specific_location)
				objCommand.Parameters.Append objCommand.CreateParameter("p_manager_contact_id", adNumeric, adParamInput, , p_manager_contact_id)
				objCommand.Parameters.Append objCommand.CreateParameter("p_phone_list_flag", adChar, adParamInput, 1, p_phone_list_flag)
				objCommand.Parameters.Append objCommand.CreateParameter("p_default_noc_region", adVarChar, adParamInput, 8, p_default_noc_region)
				objCommand.Parameters.Append objCommand.CreateParameter("p_def_remedy_support_group_id", adVarChar, adParamInput, 15, p_def_remedy_support_group_id)

				strErrMessage = "CANNOT UPDATE OBJECT"

				On Error Resume Next
				objCommand.Execute
				If objConn.Errors.Count <> 0 Then
					DisplayError "BACK", "", objConn.Errors(0).NativeError, strErrMessage, objConn.Errors(0).Description
					objConn.Errors.Clear
				End If
				on error goto 0
				strContactID = CStr(objCommand.Parameters("p_contact_id").Value)

				strWinMessage = "Record saved successfully. You can now see the changes you made."
			Else										'Create a new Service Type
				If (intAccessLevel And intConst_Access_Create) <> intConst_Access_Create Then
					DisplayError "REFRESH", strWinLocation, 0, "INSERT DENIED", "You don't have access to create Staff. Please contact your system administrator"
				End If

				objCommand.CommandText = "sma_sp_userid.spk_security_inter.sp_contact_insert"
				objCommand.Parameters.Append objCommand.CreateParameter("p_insert_userid", adVarChar, adParamInput, 20, p_insert_userid)
				objCommand.Parameters.Append objCommand.CreateParameter("p_contact_id", adNumeric, adParamInputOutput, , p_contact_id)
				objCommand.Parameters.Append objCommand.CreateParameter("p_contact_name", adVarChar, adParamInput, 50, p_contact_name)
				objCommand.Parameters.Append objCommand.CreateParameter("p_new_userid", adVarChar, adParamInput, 8, p_new_userid)
				objCommand.Parameters.Append objCommand.CreateParameter("p_password", adVarChar, adParamInput, 10, p_password)
				objCommand.Parameters.Append objCommand.CreateParameter("p_last_name", adVarChar, adParamInput, 20, p_last_name)
				objCommand.Parameters.Append objCommand.CreateParameter("p_first_name", adVarChar, adParamInput, 20, p_first_name)
				objCommand.Parameters.Append objCommand.CreateParameter("p_middle_name", adVarChar, adParamInput, 7, p_middle_name)
				objCommand.Parameters.Append objCommand.CreateParameter("p_name_prefix", adVarChar, adParamInput, 6, p_name_prefix)
				objCommand.Parameters.Append objCommand.CreateParameter("p_work_number", adVarChar, adParamInput, 24, p_work_number)
				objCommand.Parameters.Append objCommand.CreateParameter("p_work_number_ext", adVarChar, adParamInput, 10, p_work_number_ext)
				objCommand.Parameters.Append objCommand.CreateParameter("p_home_number", adVarChar, adParamInput, 24, p_home_number)
				objCommand.Parameters.Append objCommand.CreateParameter("p_cell_number", adVarChar, adParamInput, 24, p_cell_number)
				objCommand.Parameters.Append objCommand.CreateParameter("p_pager_number", adVarChar, adParamInput, 24, p_pager_number)
				objCommand.Parameters.Append objCommand.CreateParameter("p_fax_number", adVarChar, adParamInput, 24, p_fax_number)
				objCommand.Parameters.Append objCommand.CreateParameter("p_email_address", adVarChar, adParamInput, 60, p_email_address)
				objCommand.Parameters.Append objCommand.CreateParameter("p_web_site_url", adVarChar, adParamInput, 60, null)
				objCommand.Parameters.Append objCommand.CreateParameter("p_position_title", adVarChar, adParamInput, 50, p_position_title)
				objCommand.Parameters.Append objCommand.CreateParameter("p_address_id", adNumeric, adParamInput, , p_address_id)
				objCommand.Parameters.Append objCommand.CreateParameter("p_client_rep_relationship", adVarChar, adParamInput, 10, p_client_rep_relationship)
				objCommand.Parameters.Append objCommand.CreateParameter("p_receive_publications_flag", adChar, adParamInput, 1, "N")
				objCommand.Parameters.Append objCommand.CreateParameter("p_prefercontactmethodcode", adVarChar, adParamInput, 6, null)
				objCommand.Parameters.Append objCommand.CreateParameter("p_availablescheduleid", adNumeric, adParamInput, , null)
				objCommand.Parameters.Append objCommand.CreateParameter("p_comments", adVarChar, adParamInput, 2000, p_comments)
				objCommand.Parameters.Append objCommand.CreateParameter("p_responsibility", adVarChar, adParamInput, 50, p_responsibility)
				objCommand.Parameters.Append objCommand.CreateParameter("p_pinaccess", adChar, adParamInput, 1, p_pinaccess)
				objCommand.Parameters.Append objCommand.CreateParameter("p_language_preference_lcode", adChar, adParamInput, 2, p_language_preference_lcode)
				objCommand.Parameters.Append objCommand.CreateParameter("p_t2_group_flag", adChar, adParamInput, 1, "N")
				objCommand.Parameters.Append objCommand.CreateParameter("p_staff_flag", adChar, adParamInput, 1, p_staff_flag)
				objCommand.Parameters.Append objCommand.CreateParameter("p_employee_number", adVarChar, adParamInput, 8, p_employee_number)
				objCommand.Parameters.Append objCommand.CreateParameter("p_staff_status_lcode", adVarChar, adParamInput, 15, p_staff_status_lcode)
				objCommand.Parameters.Append objCommand.CreateParameter("p_department_id", adNumeric, adParamInput, , p_department_id)
				objCommand.Parameters.Append objCommand.CreateParameter("p_specific_location", adVarChar, adParamInput, 50, p_specific_location)
				objCommand.Parameters.Append objCommand.CreateParameter("p_manager_contact_id", adNumeric, adParamInput, , p_manager_contact_id)
				objCommand.Parameters.Append objCommand.CreateParameter("p_phone_list_flag", adChar, adParamInput, 1, p_phone_list_flag)
				objCommand.Parameters.Append objCommand.CreateParameter("p_default_noc_region", adVarChar, adParamInput, 8, p_default_noc_region)
				objCommand.Parameters.Append objCommand.CreateParameter("p_def_remedy_support_group_id", adVarChar, adParamInput, 15, p_def_remedy_support_group_id)

				strErrMessage = "CANNOT CREATE OBJECT"

				On Error Resume Next
				objCommand.Execute
				If objConn.Errors.Count <> 0 Then
					DisplayError "BACK", "", objConn.Errors(0).NativeError, strErrMessage, objConn.Errors(0).Description
					objConn.Errors.Clear
				End If

				on error goto 0
				strContactID = CStr(objCommand.Parameters("p_contact_id").Value)

				strWinMessage = "Record saved successfully. You can now see the changes you made."
			End If


		Case "DELETE"
			If (intAccessLevel And intConst_Access_Delete) <> intConst_Access_Delete Then
				DisplayError "REFRESH", strWinLocation, 0, "DELETE DENIED", "You don't have access to delete Staff. Please contact your system administrator"
			End If

			Set objCommand = Server.CreateObject("ADODB.Command")
			Set objCommand.ActiveConnection = objConn
			objCommand.CommandType = adCmdStoredProc
			objCommand.CommandText = "sma_sp_userid.spk_security_inter.sp_contact_delete"
			objCommand.Parameters.Append objCommand.CreateParameter("p_contact_id", adNumeric, adParamInput, , p_contact_id)					'number(9)
			objCommand.Parameters.Append objCommand.CreateParameter("p_old_userid", adVarChar, adParamInput, 8, p_old_userid)
			objCommand.Parameters.Append objCommand.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , p_last_update_dt)		'Date
			objCommand.Parameters.Append objCommand.CreateParameter("p_last_sec_update_dt", adDBTimeStamp, adParamInput, , p_last_sec_update_dt)

  			On Error Resume Next
			objCommand.Execute
			If objConn.Errors.Count <> 0 Then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			End If
			on error goto 0
			strContactID = "DEL"
			strWinMessage = "Record deleted successfully."
	End Select

	If IsNumeric(strContactID) Then

		'create SQL for populating fields

		strSQL  = "SELECT " &_
					"CON.CONTACT_ID, " &_
					"CON.CONTACT_NAME, " &_
					"CON.LAST_NAME, " &_
					"CON.FIRST_NAME, " &_
					"CON.MIDDLE_NAME, " &_
					"CON.NAME_PREFIX, " &_
					"CON.WORK_NUMBER, " &_
					"CON.WORK_NUMBER_EXT, " &_
					"CON.HOME_NUMBER, " &_
					"CON.CELL_NUMBER, " &_
					"CON.PAGER_NUMBER, " &_
					"CON.FAX_NUMBER, " &_
					"CON.EMAIL_ADDRESS, " &_
					"CON.WEB_SITE_URL, " &_
					"CON.POSITION_TITLE, " &_
					"CON.ADDRESS_ID, " &_
					"CON.CLIENT_REP_RELATIONSHIP, " &_
					"CON.RECEIVE_PUBLICATIONS_FLAG, " &_
					"CON.PREFERCONTACTMETHODCODE, " &_
					"CON.AVAILABLESCHEDULEID, " &_
					"CON.COMMENTS, " &_
					"CON.LANGUAGE_PREFERENCE_LCODE, " &_
					"CON.T2_GROUP_FLAG, " &_
					"CON.STAFF_FLAG, " &_
					"CON.EMPLOYEE_NUMBER, " &_
					"CON.STAFF_STATUS_LCODE, " &_
					"CON.DEPARTMENT_ID, " &_
					"CON.SPECIFIC_LOCATION, " &_
					"CON.PHONE_LIST_FLAG, " &_
					"CUS.CUSTOMER_ID, " &_
					"CUS.CUSTOMER_NAME, " &_
					"MAN.CONTACT_ID AS MANAGER_CONTACT_ID, " &_
					"MAN.CONTACT_NAME AS MANAGER_CONTACT_NAME, " &_
					"SEC.USERID, " &_
					"SEC.PASSWORD, " &_
					"SEC.DEFAULT_NOC_REGION_LCODE, " &_
					"SEC.DEF_REMEDY_SUPPORT_GROUP_ID, " &_
					"SEC.UPDATE_DATE_TIME AS LAST_SEC_UPDATE_DATE_TIME, " &_
					"TO_CHAR(CON.CREATE_DATE_TIME,'MON-DD-YYYY HH24:MI:SS') AS CREATE_DATE_TIME, " & _
					"SMA_SP_USERID.SPK_SMA_LIBRARY.SF_GET_FULL_USERNAME(CON.CREATE_REAL_USERID) AS CREATE_REAL_USERID, " & _
					"TO_CHAR(CON.UPDATE_DATE_TIME,'MON-DD-YYYY HH24:MI:SS') AS UPDATE_DATE_TIME, " & _
					"SMA_SP_USERID.SPK_SMA_LIBRARY.SF_GET_FULL_USERNAME(CON.UPDATE_REAL_USERID) AS UPDATE_REAL_USERID, " & _
					"CON.RECORD_STATUS_IND, " &_
					"CON.UPDATE_DATE_TIME AS LAST_UPDATE_DATE_TIME, " & _
					"(ADR.BUILDING_NAME || CHR(10) || ADR.STREET ||CHR(10)|| " & _
					"ADR.MUNICIPALITY_NAME || ' ' || ADR.PROVINCE_STATE_LCODE || ' ' || " &_
					"ADR.COUNTRY_LCODE || CHR(10) || ADR.POSTAL_CODE_ZIP) CONTACT_ADDRESS, " &_
					"CON.RESPONSIBILITY, " &_
					"CON.PIN_ACCESS "

		strFrom =	"FROM CRP.CONTACT CON, " &_
					"CRP.CONTACT MAN, " &_
					"CRP.CUSTOMER CUS, " &_
					"CRP.V_ADDRESS_CONSOLIDATED_STREET ADR, " &_
					"MSACCESS.TBLSECURITY SEC "

		strWhere =	"WHERE CON.WORK_FOR_CUSTOMER_ID = CUS.CUSTOMER_ID (+) AND " & _
					"CON.MANAGER_CONTACT_ID = MAN.CONTACT_ID (+) AND " &_
					"CON.ADDRESS_ID = ADR.ADDRESS_ID (+) AND " & _
					"CON.CONTACT_ID = SEC.STAFFID (+) AND " &_
					"CON.CONTACT_ID = " & strContactID

		strSQL =  strSQL & strFrom & strWhere

		'Response.Write strSQL
		'Response.End

		'get the contact recordset
		On Error Resume Next
		Set objRS = Server.CreateObject("ADODB.Recordset")
		objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
		If err Then DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
		If objRS.EOF Then DisplayError "BACK", "", 999, "CANNOT FIND REQUESTED OBJECT", "EOF condition occurred in objRS recordset."

		on error goto 0
		'Keep old contact number
		'p_contact_id  = objRS.Fields("CONTACT_ID").Value

		'parse out phone numbers
		'work number
		strWkArea = Mid(objRS.Fields("WORK_NUMBER").Value, 1, 3)
		strWkMid = Mid(objRS.Fields("WORK_NUMBER").Value, 4, 3)
		strWkEnd = Mid(objRS.Fields("WORK_NUMBER").Value, 7, 10)

		'home number
		strHmArea = Mid(objRS.Fields("HOME_NUMBER").Value, 1, 3)
		strHmMid = Mid(objRS.Fields("HOME_NUMBER").Value, 4, 3)
		strHmEnd = Mid(objRS.Fields("HOME_NUMBER").Value, 7, 10)

		'cell number
		strClArea = Mid(objRS.Fields("CELL_NUMBER").Value, 1, 3)
		strClMid = Mid(objRS.Fields("CELL_NUMBER").Value, 4, 3)
		strClEnd = Mid(objRS.Fields("CELL_NUMBER").Value, 7, 10)

		'pager
		strPgArea = Mid(objRS.Fields("PAGER_NUMBER").Value, 1, 3)
		strPgMid = Mid(objRS.Fields("PAGER_NUMBER").Value, 4, 3)
		strPgEnd = Mid(objRS.Fields("PAGER_NUMBER").Value, 7, 10)

		'fax number
		strFxArea = Mid(objRS.Fields("FAX_NUMBER").Value, 1, 3)
		strFxMid = Mid(objRS.Fields("FAX_NUMBER").Value, 4, 3)
		strFxEnd = Mid(objRS.Fields("FAX_NUMBER").Value, 7, 10)
	End If

	'Language Preference
	strSQL = "SELECT LANGUAGE_PREFERENCE_LCODE, LANGUAGE_PREFERENCE_DESC " &_
			"FROM CRP.LCODE_LANGUAGE_PREFERENCE " &_
			"WHERE RECORD_STATUS_IND = 'A' " &_
			"ORDER BY LANGUAGE_PREFERENCE_DESC"

	Set objLanguage = Server.CreateObject("ADODB.Recordset")
	objLanguage.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If objConn.Errors.Count <> 0 Then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Language)", objConn.Errors(0).Description
		objConn.Errors.Clear
	End If

	'Department
	strSQL = "SELECT DEPARTMENT_ID, DEPARTMENT_DESC " &_
			"FROM CRP.DEPARTMENT_LOOKUP " &_
			"WHERE RECORD_STATUS_IND = 'A' " &_
			"ORDER BY DEPARTMENT_DESC"

	Set objDepartment = Server.CreateObject("ADODB.Recordset")
	objDepartment.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If objConn.Errors.Count <> 0 Then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Department)", objConn.Errors(0).Description
		objConn.Errors.Clear
	End If

	'Region
	strSQL = "SELECT NOC_REGION_LCODE, NOC_REGION_DESC " &_
			"FROM CRP.LCODE_NOC_REGION " &_
			"WHERE RECORD_STATUS_IND = 'A' " &_
			"ORDER BY NOC_REGION_DESC"

	Set objRegion = Server.CreateObject("ADODB.Recordset")
	objRegion.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If objConn.Errors.Count <> 0 Then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Region)", objConn.Errors(0).Description
		objConn.Errors.Clear
	End If

	'Staff Status
	strSQL = "SELECT STAFF_STATUS_LCODE, STAFF_STATUS_DESC " &_
			"FROM CRP.LCODE_STAFF_STATUS " &_
			"WHERE RECORD_STATUS_IND = 'A' " &_
			"ORDER BY STAFF_STATUS_DESC"

	Set objStaffStatus = Server.CreateObject("ADODB.Recordset")
	objStaffStatus.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	If objConn.Errors.Count <> 0 Then
		DisplayError "BACK", "", objConn.Errors(0).NativeError, "ERROR LOADING DATA (Staff Status)", objConn.Errors(0).Description
		objConn.Errors.Clear
	End If
%>
<HTML>
<HEAD>
<META name="Generator" content="Microsoft Visual Studio 6.0">
<META http-equiv="Pragma" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK REL="stylesheet" TYPE="text/css" HREF="stylesheets/styles.css">
<SCRIPT type="text/javascript" language="javascript" src="AccessLevels.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" src="GeneralJavaFunctions.js"></SCRIPT>
<SCRIPT type="text/javascript" language="javascript" id="clientEventHandlersJS">
<!--
var strWinMessage = "<%=strWinMessage%>";
var intAccessLevel = <%=intAccessLevel%>;
var bolSaveRequired = false;

setPageTitle("SMA - Staff Maintenance");

function fct_selNavigate() {
//***************************************************************************************************
// Function:	fct_selNavigate															            *
// Purpose:		To display the page selected by the user from Quick Navigation drop-down box. The	*
//              To pass values to detail page use querystring; to list page use cookie.             *
// Created By:	Gilles Archer Oct 27 2000															*
//																									*																				*
//***************************************************************************************************
var strPageName = document.frmStaffDetail.selNavigate.item(document.frmStaffDetail.selNavigate.selectedIndex).value;

	switch (strPageName) {
		case 'ROLES':
			document.frmStaffDetail.selNavigate.selectedIndex = 0;
			self.location.href = "StaffRoleDetail.asp?hdnContactID=" + document.frmStaffDetail.hdnContactID.value;
			break;		//do nothing
		case 'ADDRESS':
			document.frmStaffDetail.selNavigate.selectedIndex = 0;
			var strCustomerName = document.frmStaffDetail.txtCustomerName.value;
			SetCookie("CustomerName", strCustomerName);
			self.location.href = "SearchFrame.asp?fraSrc=Address";
			break;		//do nothing
		case 'EMPLOYER':
			document.frmStaffDetail.selNavigate.selectedIndex = 0;
			var strCustomerID = document.frmStaffDetail.hdnCustomerID.value;
			self.location.href = "CustDetail.asp?CustomerID=" + strCustomerID;
			break;		//do nothing
		case 'MANAGER':
			document.frmStaffDetail.selNavigate.selectedIndex = 0;
			var strContactID = document.frmStaffDetail.hdnManagerContactID.value;
			self.location.href = "StaffDetail.asp?ContactID=" + strContactID;
			break;		//do nothing
		case 'DEFAULT':
			document.frmStaffDetail.selNavigate.selectedIndex = 0;
			break;		//do nothing
		default:
			document.frmStaffDetail.selNavigate.selectedIndex = 0;
			break;		//do nothing
	}
}

function fct_onChange() {
	bolSaveRequired = true;
}

function btnSave_onClick() {

	if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
		alert('You do not have permission to UPDATE a Staff.  Please contact your System Administrator.');
		return false;
	}

	if (document.frmStaffDetail.txtNameFirst.value == "" ) {
		alert('Missing required field. Please enter a First Name');
		document.frmStaffDetail.txtNameFirst.focus();
		return false;
	}
	if (document.frmStaffDetail.txtNameLast.value == "" ) {
		alert('Missing required field. Please enter a Last Name');
		document.frmStaffDetail.txtNameLast.focus();
		return false;
	}
	if (document.frmStaffDetail.txtEmail.value == "" ) {
		alert('Missing required field. Please enter an e-mail address.');
		document.frmStaffDetail.txtEmail.focus();
		return false;
	}
	/*if (document.frmStaffDetail.selLanguage.selectedIndex == 0) {
		alert('Missing required field. Please select a Language using the drop down list.');
		document.frmStaffDetail.selLanguage.focus();
		return false;
	}
	*/
	if (document.frmStaffDetail.selDepartment.selectedIndex == 0) {
		alert('Missing required field. Please select a Department using the drop down list.');
		document.frmStaffDetail.selDepartment.focus();
		return false;
	}

	if (document.frmStaffDetail.txtEmpNo.value == "") {
		alert('Missing required field. Please enter a valid employee number');
		document.frmStaffDetail.txtEmpNo.focus();
		return false;
	}

	if (document.frmStaffDetail.txtUserID.value == "") {
		alert('Missing required field. Please enter the User ID');
		document.frmStaffDetail.txtUserID.focus();
		return false;
	}
	if (document.frmStaffDetail.txtPassword.value == "") {
		alert('Missing required field. Please enter the Password');
		document.frmStaffDetail.txtPassword.focus();
		return false;
	}
	if (document.frmStaffDetail.txtConfirm.value == "") {
		alert('Missing required field. Please re-enter the Password');
		document.frmStaffDetail.txtConfirm.focus();
		return false;
	}
	if (document.frmStaffDetail.txtPassword.value != document.frmStaffDetail.txtConfirm.value) {
		alert('The typed password and confirmation passwords do not match.\nThese passwords must match in order to save the record.');
		document.frmStaffDetail.txtConfirm.focus();
		document.frmStaffDetail.txtConfirm.select();
		return false;
	}

	//check that employee number consists of numbers & 8 chars
	var EMPNo = document.frmStaffDetail.txtEmpNo.value;
	if (EMPNo.length > 0) {
		if (isNaN(EMPNo)) {
			alert('Employee number must consist of digits only');
			document.frmStaffDetail.txtEmpNo.focus();
			document.frmStaffDetail.txtEmpNo.select();
			return false;
		}
		if (EMPNo.length != 8) {
			alert('Employee number must consist of 8 digits in "0009999" format');
			document.frmStaffDetail.txtEmpNo.focus();
			document.frmStaffDetail.txtEmpNo.select();
			return false;
		}
	}

	//check that all phone numbers consist of numbers & 10 chars
	//work phone
	var WPhone = document.frmStaffDetail.txtWArea.value + document.frmStaffDetail.txtWMid.value + document.frmStaffDetail.txtWEnd.value;
	if (WPhone.length > 0) {
		if (isNaN(WPhone)) {
			alert('Work phone number must consist of digits only.');
			document.frmStaffDetail.txtWArea.focus();
			document.frmStaffDetail.txtWArea.select();
			return false;
		}
		if (WPhone.length != 10) {
			alert('Work phone number must consist of 10 digits (###) ###-####.');
			document.frmStaffDetail.txtWArea.focus();
			document.frmStaffDetail.txtWArea.select();
			return false;
		}
	}
	//work phone ext
	var WExt = document.frmStaffDetail.txtExt.value;
	if (WExt.length > 0) {
		if (isNaN(WExt)) {
			alert('Work phone extension must consist of digits only.');
			document.frmStaffDetail.txtExt.focus();
			document.frmStaffDetail.txtExt.select();
			return false;
		}
	}
	//Cell phone
	var CPhone = document.frmStaffDetail.txtCArea.value + document.frmStaffDetail.txtCMid.value + document.frmStaffDetail.txtCEnd.value;
	if (CPhone.length > 0) {
		if (isNaN(CPhone)) {
			alert('Cell phone number must consist of digits only.');
			document.frmStaffDetail.txtCArea.focus();
			document.frmStaffDetail.txtCArea.select();
			return false;
		}
		if (CPhone.length != 10) {
			alert('Cell phone number must consist of 10 digits (###) ###-####.');
			document.frmStaffDetail.txtCArea.focus();
			document.frmStaffDetail.txtCArea.select();
			return false;
		}
	}
	//pager
	var PPhone = document.frmStaffDetail.txtPArea.value + document.frmStaffDetail.txtPMid.value + document.frmStaffDetail.txtPEnd.value;
	if (PPhone.length > 0) {
		if (isNaN(PPhone)) {
			alert('Pager number must consist of digits only.');
			document.frmStaffDetail.txtPArea.focus();
			document.frmStaffDetail.txtPArea.select();
			return false;
		}
		if (PPhone.length != 10) {
			alert('Pager number must consist of 10 digits (###) ###-####.');
			document.frmStaffDetail.txtPArea.focus();
			document.frmStaffDetail.txtPArea.select();
			return false;
		}
	}
	//Fax
	var FPhone = document.frmStaffDetail.txtFArea.value + document.frmStaffDetail.txtFMid.value + document.frmStaffDetail.txtFEnd.value;
	if (FPhone.length > 0) {
		if (isNaN(FPhone)) {
			alert('Fax number must consist of digits only.');
			document.frmStaffDetail.txtFArea.focus();
			document.frmStaffDetail.txtFArea.select();
			return false;
		}
		if (FPhone.length != 10) {
			alert('Fax number must consist of 10 digits (###) ###-####.');
			document.frmStaffDetail.txtFArea.focus();
			document.frmStaffDetail.txtFArea.select();
			return false;
		}
	}
	//home phone
	var HPhone = document.frmStaffDetail.txtHArea.value + document.frmStaffDetail.txtHMid.value + document.frmStaffDetail.txtHEnd.value;
	if (HPhone.length > 0) {
		if (isNaN(HPhone)) {
			alert('Home phone number must consist of digits only.');
			document.frmStaffDetail.txtHArea.focus();
			document.frmStaffDetail.txtHArea.select();
			return false;
		}
		if (HPhone.length != 10) {
			alert('Home phone number must consist of 10 digits (###) ###-####.');
			document.frmStaffDetail.txtHArea.focus();
			document.frmStaffDetail.txtHArea.select();
			return false;
		}
	}

	//comments
	var strComments = document.frmStaffDetail.txtComments.value;
	if (strComments.length > 2000) {
		alert('Comments can be at most 2000 characters.\n\nYou entered ' + strComments.length + ' character(s).');
		document.frmStaffDetail.txtComments.focus();
	return false;
	}

	bolSaveRequired = false; //bypass message asking if you want to save on window_onunload
	document.frmStaffDetail.hdnFrmAction.value = 'SAVE';
	document.frmStaffDetail.submit();
	return true;
}

function body_onBeforeUnload() {
	document.frmStaffDetail.btnSave.focus();
	if (bolSaveRequired) {
		event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
	}
}

function ClearStatus() {
	window.status = "";
}

function DisplayStatus(strWinStatus) {
	window.status = strWinStatus;
	setTimeout('ClearStatus()', 5000);
}

function btnAddressLookup_onClick() {
	if (document.frmStaffDetail.txtCustomerName.value != "" ) {
		SetCookie("CustomerName", document.frmStaffDetail.txtCustomerName.value);
	}
	SetCookie("WinName", "Popup");
	bolSaveRequired = true;
	window.open('SearchFrame.asp?fraSrc=Address', 'Popup', 'top=50, left=100, height=600, width=800');
}

function btnAddressClear_onClick() {
	document.frmStaffDetail.hdnAddressID.value = "";
	document.frmStaffDetail.textAddress.value = "";
}

function btnCustomerLookup_onClick(CustService) {
	if (document.frmStaffDetail.txtCustomerName.value != "") {
		 SetCookie("CustomerName", document.frmStaffDetail.txtCustomerName.value);
	}
	SetCookie("WinName", "Popup");
	SetCookie("ServiceEnd", CustService);
	bolSaveRequired = true;
	window.open('SearchFrame.asp?fraSrc=Cust', 'Popup', 'top=50, left=100, height=600, width=800');
}

function btnManagerClear_onClick() {
	document.frmStaffDetail.hdnManagerContactID.value = "";
	document.frmStaffDetail.txtManagerContactName.value = "";
}

function btnManagerLookup_onClick() {
	if (document.frmStaffDetail.txtCustomerName.value != "") {
		 SetCookie("WorkFor", document.frmStaffDetail.txtCustomerName.value);
	}
	SetCookie("Case", "M");
	SetCookie("TelusOnly", "yes");
	SetCookie("WinName", "Popup");
	bolSaveRequired = true;
	window.open('SearchFrame.asp?fraSrc=Contact', 'Popup', 'top=50, left=100, height=600, width=800');
}

function btnReferences_onClick() {
var strOwner = 'CRP';
var strTableName = 'CONTACT';
var strRecordID = document.frmStaffDetail.hdnContactID.value;
var URL = 'Dependency.asp?Owner=' + strOwner + '&TableName=' + strTableName + '&RecordID=' + strRecordID;

	if (strRecordID == "NEW") {
		alert("No references. This is a new record.");
		return false;
	}
	window.open(URL, 'Popup', 'top=100, left=100, width=500, height=300');
}

function btnDelete_onClick() {
	if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
		alert('Access denied. Please contact your system administrator.');
		return false;
	}
	if (confirm('Do you really want to delete this contact?')) {
		document.frmStaffDetail.hdnFrmAction.value = "DELETE";
		document.frmStaffDetail.submit();
	}
}

function btnReset_onClick() {
	if(confirm('All changes will be lost. Do you really want to reset the page?')){
		bolSaveRequired = false;
		document.location = "StaffDetail.asp?ContactID=<%=strContactID%>";
	}
}

function btnNew_onClick() {
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('Access denied. Please contact your system administrator.');
		return false;
	}
	self.document.location.href="StaffDetail.asp?ContactID=NEW";
}
-->
</SCRIPT>
</HEAD>
<BODY onLoad="DisplayStatus(strWinMessage);" onBeforeUnload="return body_onBeforeUnload();">
<FORM id="frmStaffDetail" name="frmStaffDetail" action="StaffDetail.asp" method="post">
	<INPUT type="hidden" id="hdnUserID" name="hdnUserID" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("USERID").Value%>">
	<INPUT type="hidden" id="hdnContactID" name="hdnContactID" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("CONTACT_ID").Value%>">
	<INPUT type="hidden" id="hdnManagerContactID" name="hdnManagerContactID" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("MANAGER_CONTACT_ID").Value%>">
	<INPUT type="hidden" id="hdnAddressID" name="hdnAddressID" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("ADDRESS_ID").Value%>">
	<INPUT type="hidden" id="hdnCustomerID" name="hdnCustomerID" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("CUSTOMER_ID").Value else Response.Write "6746" end if%>">
	<INPUT type="hidden" id="hdnUpdateDateTime" name="hdnUpdateDateTime" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("LAST_UPDATE_DATE_TIME").Value%>">
	<INPUT type="hidden" id="hdnSecUpdateDateTime" name="hdnSecUpdateDateTime" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("LAST_SEC_UPDATE_DATE_TIME").Value%>">
	<INPUT type="hidden" id="selRemedySupport" name="selRemedySupport" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("DEF_REMEDY_SUPPORT_GROUP_ID").Value%>">
	<INPUT type="hidden" id="hdnFrmAction" name="hdnFrmAction" value="">

<TABLE border="0" cellpadding="1" cellspacing="1" cols="6" width="100%">
<THEAD>
<TR valign="top">
	<TH align="left" colspan="3">Staff Detail</TH>
	<TH align="right">
		<SELECT id="selNavigate" name="selNavigate" onchange="fct_selNavigate();" <%if not isNumeric(strContactID) then Response.Write "disabled" end if%>>
			<OPTION selected value="DEFAULT">Quickly Goto ...</OPTION>
			<OPTION value="ROLES">Business Roles</OPTION>
			<OPTION value="ADDRESS">Address</OPTION>
			<OPTION value="EMPLOYER">Employer</OPTION>
			<OPTION value="MANAGER">Manager</OPTION>
		</SELECT>
	</TH>
</TR>
</THEAD>
<TBODY>
<TR valign="top">
	<TD align="left" nowrap colspan=2><STRONG><%If IsNumeric(strContactID) Then Response.Write objRS.Fields("CONTACT_NAME").Value else Response.Write "&nbsp;" end if%></STRONG></TD>

	<!--<TD align="left" nowrap><INPUT disabled id="txtContactName" name="txtContactName" maxlength="50" size="50" onChange="fct_onChange();" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("CONTACT_NAME").Value%>"></INPUT></TD>-->
	<td rowspan=4 colspan=2>
		<table>
			<thead>
				<th colspan=2 align=left>Logon Info</th>
			</thead>
			<tr>
				<TD align="right" nowrap>User ID<FONT color="red">*</FONT></TD>
				<TD align="left" nowrap><INPUT id="txtUserID" name="txtUserID" type="text" size="8" maxlength="8" onChange="fct_onChange();" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("USERID").Value%>"></TD>
			</tr>
				<TD align="right" nowrap>Password<FONT color="red">*</FONT></TD>
				<TD align="left" nowrap><INPUT id="txtPassword" name="txtPassword" type="password" size="10" maxlength="10" onChange="fct_onChange();" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("PASSWORD").Value%>"></TD>
			<tr>
			</tr>
			<tr>
				<TD align="right" nowrap>Confirm<FONT color="red">*</FONT></TD>
				<TD align="left" nowrap><INPUT id="txtConfirm" name="txtConfirm" type="password" size="10" maxlength="10" onChange="fct_onChange();" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("PASSWORD").Value%>"></TD>
			</tr>
		</table>
	</td>
</TR>
<TR valign="top">
	<TD align="right" nowrap>Title</TD>
	<TD align="left" nowrap>
		<SELECT id="selNamePrefix" name="selNamePrefix" onChange="fct_onChange();">
			<OPTION></OPTION>
			<%For lIndex = LBound(arrNamePrefix) To UBound(arrNamePrefix)
				If IsNumeric(strContactID) Then
					If StrComp(arrNamePrefix(lIndex), objRS.Fields("NAME_PREFIX").Value, 0) = 0 Then%>
						<OPTION selected value="<%=arrNamePrefix(lIndex)%>"><%=arrNamePrefix(lIndex)%></OPTION>
					<%Else%>
						<OPTION value="<%=arrNamePrefix(lIndex)%>"><%=arrNamePrefix(lIndex)%></OPTION>
					<%End If
				Else%>
					<OPTION value="<%=arrNamePrefix(lIndex)%>"><%=arrNamePrefix(lIndex)%></OPTION>
				<%End If
			Next%>
		</SELECT>
	</td>
</tr>
<tr>
	<td align="right">First Name<FONT color="red">*</FONT></td>
	<td><INPUT id="txtNameFirst" name="txtNameFirst" maxlength="20" size="20" onChange="fct_onChange();" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("FIRST_NAME").Value%>"></INPUT></td>
</tr>
<tr>
	<td align="right">Middle Name</td>
	<td><INPUT id="txtNameMiddle" name="txtNameMiddle" maxlength="7" size="7" onChange="fct_onChange();" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("MIDDLE_NAME").Value%>"></INPUT></TD>
</tr>
<TR valign="top">
	<TD align="right" nowrap>Last Name<FONT color="red">*</FONT></TD>
	<TD align="left" nowrap><INPUT id="txtNameLast" name="txtNameLast" maxlength="20" size="20" onChange="fct_onChange();" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("LAST_NAME").Value%>"></INPUT>
	<TD align="right" nowrap>Language<FONT color="red">*</FONT></TD>
	<TD align="left" nowrap><SELECT id="selLanguage" name="selLanguage" onChange="fct_onChange();">
			<%Do While Not objLanguage.EOF
				If IsNumeric(strContactID) Then
					If StrComp(objLanguage.Fields("LANGUAGE_PREFERENCE_LCODE").Value, objRS.Fields("LANGUAGE_PREFERENCE_LCODE").Value, 0) = 0 Then%>
						<OPTION selected value="<%=objLanguage.Fields("LANGUAGE_PREFERENCE_LCODE").Value%>"><%=routineHtmlString(objLanguage.Fields("LANGUAGE_PREFERENCE_DESC").Value)%></OPTION>
					<%Else%>
						<OPTION value="<%=objLanguage.Fields("LANGUAGE_PREFERENCE_LCODE").Value%>"><%=routineHtmlString(objLanguage.Fields("LANGUAGE_PREFERENCE_DESC").Value)%></OPTION>
					<%End If
				Else%>
				<OPTION value="<%=objLanguage.Fields("LANGUAGE_PREFERENCE_LCODE").Value%>"><%=routineHtmlString(objLanguage.Fields("LANGUAGE_PREFERENCE_DESC").Value)%></OPTION>
				<%End If
				objLanguage.MoveNext
			Loop
			objLanguage.Close
			Set objLanguage = Nothing%>
		</SELECT>
	</TD>
</TR>
<TR valign="top">
	<TD align="right" nowrap>Position</TD>
	<TD align="left" nowrap><INPUT id="txtPosition" name="txtPosition" maxlength="50" size="50" onChange="fct_onChange();" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("POSITION_TITLE").Value%>"></INPUT></TD>
	<TD align="right" nowrap>Employee No<FONT color="red">*</FONT></TD>
	<TD align="left" nowrap><INPUT id="txtEmpNo" name="txtEmpNo" type="text" size="8" maxlength="8" onChange="fct_onChange();" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("EMPLOYEE_NUMBER").Value%>"></TD>
</TR>
<TR valign="top">
	<TD align="right" nowrap>Manager</TD>
	<TD align="left" nowrap>
		<INPUT id="txtManagerContactName" name="txtManagerContactName" type="text" size="50" maxlength="50" disabled onChange="fct_onChange();" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("MANAGER_CONTACT_NAME").Value%>">
		<INPUT id="btnManagerLookup" name="btnManagerLookup" type="button" value="..." onClick="return btnManagerLookup_onClick();">
		<INPUT id="btnManagerClear" name="btnManagerClear" type="button" value="X" onClick="btnManagerClear_onClick();">
	</TD>
	<TD align="right" nowrap>Status</TD>
	<TD align="left" nowrap><SELECT id="selStaffStatus" name="selStaffStatus" onChange="fct_onChange();">
			<OPTION></OPTION>
			<%Do While Not objStaffStatus.EOF
				If IsNumeric(strContactID) Then
					If StrComp(objStaffStatus.Fields("STAFF_STATUS_LCODE").Value, objRS.Fields("STAFF_STATUS_LCODE").Value, 0) = 0 Then%>
						<OPTION selected value="<%=objStaffStatus.Fields("STAFF_STATUS_LCODE").Value%>"><%=routineHtmlString(objStaffStatus.Fields("STAFF_STATUS_DESC").Value)%></OPTION>
					<%Else%>
						<OPTION value="<%=objStaffStatus.Fields("STAFF_STATUS_LCODE").Value%>"><%=routineHtmlString(objStaffStatus.Fields("STAFF_STATUS_DESC").Value)%></OPTION>
					<%End If
				Else%>
				<OPTION value="<%=objStaffStatus.Fields("STAFF_STATUS_LCODE").Value%>"><%=routineHtmlString(objStaffStatus.Fields("STAFF_STATUS_DESC").Value)%></OPTION>
				<%End If
				objStaffStatus.MoveNext
			Loop
			objStaffStatus.Close
			Set objStaffStatus = Nothing%>
		</SELECT>
	</TD>
</TR>
<TR>
	<TD align="right" nowrap>Works For</TD>
	<TD align="left" nowrap>
		<INPUT type="text" id="txtCustomerName" name="txtCustomerName" size="50" maxlength="50" disabled onChange="fct_onChange();" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("CUSTOMER_NAME").Value else Response.Write "TELUS Advanced Communications DataInfo" end if%>"></INPUT>
		<!-- <INPUT type="button" id="btnCustomerLookup" name="btnCustomerLookup" value="..." onClick="return btnCustomerLookup_onClick('C');"> -->
	</TD>
	<TD align="right" nowrap>Department<font color=red>*</font></TD>
	<TD align="left" nowrap>
		<SELECT id="selDepartment" name="selDepartment" onChange="fct_onChange();">
			<OPTION></OPTION>
			<%Do While Not objDepartment.EOF
				Response.Write "<OPTION "
				If IsNumeric(strContactID) Then
					If objRS.Fields("DEPARTMENT_ID").Value <> "" Then
						If CLng(objDepartment.Fields("DEPARTMENT_ID").Value)= CLng(objRS.Fields("DEPARTMENT_ID").Value) Then
							Response.Write " selected "
						End If
					End If
				End If
				Response.Write "value=" & objDepartment.Fields("DEPARTMENT_ID").Value & ">" & routineHtmlString(objDepartment.Fields("DEPARTMENT_DESC").Value) & "</OPTION>" & vbCrLf
				objDepartment.MoveNext
			Loop
			objDepartment.Close
			Set objDepartment = Nothing%>
		</SELECT>
	</TD>
</tr>
<TR valign="top">
	<TD align="right" rowspan="3" nowrap>Address</TD>
	<TD align="left" rowspan="3" nowrap>
		<TEXTAREA id="textAddress" name="textAddress" cols="25" rows="4" disabled style="width: 363" onChange="fct_onChange();"><%If IsNumeric(strContactID) Then Response.Write objRS.Fields("CONTACT_ADDRESS").Value%></TEXTAREA>
		<INPUT id="btnAddressLookup" name="btnAddressLookup" type="button" value="..." onClick="return btnAddressLookup_onClick();">
		<INPUT id="btnAddressClear" name="btnAddressClear" type="button" value="X" onClick="btnAddressClear_onClick();">
	</TD>
	<td>&nbsp;</td>
</TR>
<TR valign="top">
	<TD align="right" nowrap>Work</TD>
	<TD align="left" nowrap >(<INPUT id="txtWArea" name="txtWArea" size="3" maxlength="3" onChange="fct_onChange();" value="<%=strWkArea%>"></INPUT>)
		<INPUT id="txtWMid" name="txtWMid" size="3" maxlength="3" onChange="fct_onChange();" value="<%=strWkMid%>"></INPUT>
		-&nbsp;<INPUT id="txtWEnd" name="txtWEnd" size="4" maxlength="4" onChange="fct_onChange();" value="<%=strWkEnd%>"></INPUT>
	Ext.<INPUT id="txtExt" name="txtExt" size="10" maxlength="10" onChange="fct_onChange();" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("WORK_NUMBER_EXT").Value%>"></INPUT></td>
</tr>
<tr>
	<TD align="right" nowrap>Fax</TD>
	<TD align="left" nowrap>
		(<INPUT id="txtFArea" name="txtFArea" size="3" maxlength="3" onChange="fct_onChange();" value="<%=strFxArea%>"></INPUT>)
		<INPUT id="txtFMid" name="txtFMid" size="3" maxlength="3" onChange="fct_onChange();" value="<%=strFxMid%>"></INPUT>
		-&nbsp;<INPUT id="txtFEnd" name="txtFEnd" size="4" maxlength="4" onChange="fct_onChange();" value="<%=strFxEnd%>"></INPUT>
	</TD>
</tr>
<TR valign="top">
	<TD align="right" nowrap>Location</TD>
	<TD align="left" nowrap><INPUT id="txtLocation" name="txtLocation" type="text" size="50" maxlength="50" onChange="fct_onChange();" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("SPECIFIC_LOCATION").Value%>"></TD>
	<TD align="right" nowrap>Cell</TD>
	<TD align="left" nowrap>
		(<INPUT id="txtCArea" name="txtCArea" size="3" maxlength="3" onChange="fct_onChange();" value="<%=strClArea%>"></INPUT>)
		<INPUT id="txtCMid" name="txtCMid" size="3" maxlength="3" onChange="fct_onChange();" value="<%=strClMid%>"></INPUT>
		-&nbsp;<INPUT id="txtCEnd" name="txtCEnd" size="4" maxlength="4" onChange="fct_onChange();" value="<%=strClEnd%>"></INPUT>
	</TD>
</TR>
<TR valign="top">
	<TD align="right" nowrap>Default NOC Region</TD>
	<TD align="left" nowrap>
		<SELECT id="selNOCRegion" name="selNOCRegion" onChange="fct_onChange();">
			<OPTION></OPTION>
			<%Do While Not objRegion.EOF
				Response.Write "<OPTION "
				If IsNumeric(strContactID) Then
					If objRS.Fields("DEFAULT_NOC_REGION_LCODE").Value <> "" Then
						If StrComp(objRegion.Fields("NOC_REGION_LCODE").Value, objRS.Fields("DEFAULT_NOC_REGION_LCODE").Value, 0) = 0 Then
							Response.Write " selected "
						End If
					End If
				End If
				Response.Write "value=" & objRegion.Fields("NOC_REGION_LCODE").Value & ">" & routineHtmlString(objRegion.Fields("NOC_REGION_DESC").Value) & "</OPTION>" & vbCrLf
				objRegion.MoveNext
			Loop
			objRegion.Close
			Set objRegion = Nothing%>
		</SELECT>
	</TD>
	<TD align="right" nowrap>Pager</TD>
	<TD align="left" nowrap>
		(<INPUT id="txtPArea" name="txtPArea" size="3" maxlength="3" onChange="fct_onChange();" value="<%=strPgArea%>"></INPUT>)
		<INPUT id="txtPMid" name="txtPMid" size="3" maxlength="3" onChange="fct_onChange();" value="<%=strPgMid%>"></INPUT>
		-&nbsp;<INPUT id="txtPEnd" name="txtPEnd" size="4" maxlength="4" onChange="fct_onChange();" value="<%=strPgEnd%>"></INPUT>
	</TD>
</TR>
<TR valign="top">
	<TD align="right" rowspan="3">Comments</TD>
	<TD align="left" rowspan="3" ><TEXTAREA id="txtComments" name="txtComments" cols="25" rows="4" style="width: 360" onChange="fct_onChange();"><%If IsNumeric(strContactID) Then Response.Write objRS.Fields("COMMENTS").Value%></TEXTAREA>
		<INPUT id="btnCommentsClear" name="btnCommentsClear" type="button" value="X" onClick="document.frmStaffDetail.txtComments.value = '';">
	</TD>
    <TD align="right" nowrap>Home</TD>
	<TD align="left" nowrap>(<INPUT id="txtHArea" name="txtHArea" size="3" maxlength="3" onChange="fct_onChange();" value="<%=strHmArea%>"></INPUT>)
		<INPUT id="txtHMid" name="txtHMid" size="3" maxlength="3" onChange="fct_onChange();" value="<%=strHmMid%>"></INPUT>
		-&nbsp;<INPUT id="txtHEnd" name="txtHEnd" size="4" maxlength="4" onChange="fct_onChange();" value="<%=strHmEnd%>"></INPUT>
	</TD>
</tr>
<TR valign="top">
	<TD align="right" nowrap>E-mail<font color=red>*</font></TD>
	<TD align="left" nowrap><INPUT id="txtEmail" name="txtEmail" size="60" maxlength="60" style="width=12cm" onChange="fct_onChange();" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("EMAIL_ADDRESS").Value%>"></INPUT>
</TR>
<tr>
	<TD align="right" nowrap>Phone List</TD>
	<TD align="left" nowrap><INPUT id="chkPhoneList" name="chkPhoneList" type="checkbox" onChange="fct_onChange();" <%If IsNumeric(strContactID) Then If objRS.Fields("PHONE_LIST_FLAG").Value = "Y" Then Response.Write "checked"%>></TD>
</tr>
<tr>
	<td  align="top" ><font color=purple>Fields below are used by external contacts only:</font>&nbsp;</td>
</tr>
<tr>
	<td  align=right valign="top" >Responsibility&nbsp;</td>
	<td  valign="top"><INPUT name=txtResponsibility size=50 maxlength=50 tabindex=10 onChange="fct_onChange();"value=<%if isNumeric(strContactID) then Response.Write """"&objRS.Fields("responsibility")&"""" else Response.Write """""" end if%>></input></td>
        	<!--<td  align="right">PIN&nbsp;</td>
		<td  align="left"><INPUT readonly style=color:silver name="txtPINAccess" type="text" style="width: 18px" value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("PIN_access").Value%>">&nbsp;</td>-->
	<td  align="right">PIN&nbsp;</td>
	<td  align="left" ><INPUT  name=chkPINAccess type=checkbox onChange="fct_onChange();"
		<%if isNumeric(strContactID) then
			if objRS.Fields("PIN_access") = "Y" then
				Response.Write " checked "
			end if
		end if%>>
	</td>
</tr>
<tr>
	<td>&nbsp;</td>
</tr>
</TBODY>
<TFOOT>
<TR valign="top">
	<TH align="right" colspan="6" nowrap>
	<INPUT id="btnReferences" name="btnReferences" type="button" value="References" style="width: 2.2cm" language="javascript" onClick="return btnReferences_onClick();">&nbsp;
	<INPUT id="btnDelete" name="btnDelete" type="button" value="Delete" style="width: 2cm" language="javascript" onClick="return btnDelete_onClick();">&nbsp;
	<INPUT id="btnReset" name="btnReset" type="button" value="Reset" style="width: 2cm" language="javascript" onClick="return btnReset_onClick();">&nbsp;
	<INPUT id="btnNew" name="btnNew" type="button" value="New" style="width: 2cm" language="javascript" onClick="return btnNew_onClick();">&nbsp;
	<INPUT id="btnSave" name="btnSave" type="button" value="Save" style="width: 2cm" onClick="return btnSave_onClick();">&nbsp;</TH>
</TR>
</TFOOT>
</TABLE>
<FIELDSET width="100%">
	<LEGEND align="right"><b>Audit Information</b></LEGEND>
	<DIV size="8pt" align="right">
	Record Status Indicator:<INPUT align="left" name="txtRecordStatusInd" type="text" style="width: 18px" disabled value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("RECORD_STATUS_IND").Value%>">&nbsp;&nbsp;&nbsp;
	Create Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("CREATE_DATE_TIME").Value%>">&nbsp;
	Created By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("CREATE_REAL_USERID").Value%>"><BR>
	Update Date:<INPUT align="center" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("UPDATE_DATE_TIME").Value%>">&nbsp;
	Updated By:<INPUT align="right" name="txtRecordStatusInd" type="text" style="width: 150px" disabled value="<%If IsNumeric(strContactID) Then Response.Write objRS.Fields("UPDATE_REAL_USERID").Value%>">
	</DIV>
</FIELDSET>
<% 'clean up ADO objects
	On Error Resume Next
	objRS.Close
	Set objRS = Nothing
	objConn.Close
	Set objConn = Nothing
%>
</FORM>
</BODY>
</HTML>
