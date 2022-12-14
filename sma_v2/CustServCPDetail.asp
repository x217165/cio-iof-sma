<%@ Language=VBScript %>
<% option explicit %>
<% Response.Buffer=true%>
<!--#include file="SmaConstants.inc"-->
<!--#include file="SMA_Env.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--
***************************************************************************************************
* Name:		CustServCPDetail.asp i.e. Customer Service List
*
* Purpose:	This page displays information about a customer service and allows the user to update it
*
* Created By:	Sara Sangha 08/01/00
***************************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       10-Apr-01	 DTy		Limit Service Name to 80 characters to prevent program
                                  crash.
	   20-Jul-01	 DTy  		Do not allow partial entry for Date Start Billing, Date
                                  Order Received, Completion Date, Date Configured and
								  Date Installed.
	   27-Nov-01     DTy        Add number of seats.
	   27-Feb-02	 DTy		Add Customer Service Alias iFrame.
	   28-Feb-02     DTy        Display 'No. of Seats' only entering a new Customer Service
	                              or updating ASP-related Customer Service.
	   10-Sept-04    MW         Add Repair Priority
	   10-Aug-12     ACheung	Add Service Type and Service Instance Attributes
   	   18-Jun-13     ACheung	Adding VPN info adapted from CustServDetail.asp
***************************************************************************************************
-->
<%

'check user's rights
dim intAccessLevel, intChildAccessLevel
dim intRowCount, intColCount, strInnerValues, strWinMessage, strWinLocation, strLANG, strSType, strSTypeEN
Dim  logCustomerServiceID, datUpdateDateTime, strRealUserID, strServiceTypeID
Dim bolCloned 'used to determine if this record was cloned or not

strLANG = Request.Cookies("UserInformation")("language_preference")
if (Len(strLANG) = 0) then strLANG = "EN"

intAccessLevel = CInt(CheckLogon(strConst_CustomerService))
if (intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to view customer service. Please contact your system administrator."
end if

intChildAccessLevel = CInt(CheckLogon(strConst_CustomerServiceContact))

logCustomerServiceID = Request("CustServID")
datUpdateDateTime = Request("UpdateDateTime")
'strRealUserID = Session("username")
strRealUserID = Session("username")
strWinLocation = "CustServDetail.asp?CustServID="&Request.Form("hdnCustomerServiceID")
bolCloned = (UCase(Request("NewCustServ")) = "CLONED")
strServiceTypeID = Request("ServiceTypeID")

select case Request("hdnFrmAction")
	case "SAVE"

	  if Request.Form("hdnCustomerServiceID")  <> "" then  ' it is an existing record so save the changes

		if (intAccessLevel and intConst_Access_Update) <> intConst_Access_Update then
			DisplayError "REFRESH", strWinLocation, 0, "UPDATE DENIED", "You don't have access to update customer service. Please contact your system administrator."
		end if

		dim cmdUpdateObj
		set cmdUpdateObj = server.CreateObject("ADODB.Command")
		set cmdUpdateObj.ActiveConnection = objConn
		cmdUpdateObj.CommandType = adCmdStoredProc
		cmdUpdateObj.CommandText = "sma_sp_userid.spk_sma_cserv_inter.sp_cs_update"

		logCustomerServiceID = Request("hdnCustomerServiceID")

		'create params
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_userid", adVarChar, adParamInput, 20 ,strRealUserID)				 						'real_userid
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_customer_service_id", adNumeric, adParamInput,,clng(Request("hdnCustomerServiceID")))	'customer_service_id
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_remedy_group", adVarChar, adParamInput, 20, Request("selSupportGroup"))					'
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_cs_desc", adVarChar, adParamInput , 80, left(Request("txtCustomerServiceName"), 80))		'
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_customer_id", adNumeric, adParamInput, , clng(Request("hdnCustomerID")))					'
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_service_type", adNumeric, adParamInput, , clng(Request("hdnServiceTypeID")))				'
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_sla_id", adNumeric,adParamInput, , clng(Request("hdnSLAID")))							'
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_service_status", adVarChar, adParamInput, 6, Request("hdnServiceStatusCode"))				'

		IF Request("hdnServLocID") <> "" THEN
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_sl_id",adNumeric, adParamInput, , clng(Request("hdnServLocID")))
		ELSE
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_sl_id", adNumeric, adParamInput, , NULL)
		END IF

		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_date_inservice", adVarChar, adParamInput,20 , cstr(Request("hdnDateInService")))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_date_terminated", adVarChar, adParamInput,20 , cStr(Request("hdnDateTerminated")))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_billing_date", adVarChar, adParamInput,20 , Request("hdnBillingStartDate"))

		IF	Request("txtComment") <> "" THEN
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_comments", adVarChar, adParamInput, 2000, Request("txtComment"))
		ELSE
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_comments", adVarChar, adParamInput, 2000, NULL)
		END IF

		IF	Request("txtNoOfSeats") <> "" THEN
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_No_Of_Seats", adNumeric, adParamInput, , clng(Request("txtNoOfSeats")))
		ELSE
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_No_Of_Seats", adNumeric, adParamInput, , NULL)
		END IF

		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_repair_priority", adVarChar, adParamInput, 20, Request("selRepairPriority"))				'

		IF Request("txtOrderNumber") <> "" THEN
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_order_no", adVarChar, adParamInput, 10, Request("txtOrderNumber"))
		ELSE
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_order_no", adVarChar, adParamInput, 10, NULL)
		END IF

		if Request("hdnContactID1") <> "" THEN
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_design_st", adNumeric, adParamInput, , clng(Request("hdnContactID1")))
		ELSE
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_design_st", adNumeric, adParamInput, , null)
		END IF

		if  Request("hdnContactID2") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_implement_st", adNumeric, adParamInput, , clng(Request("hdnContactID2")))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_implement_st", adNumeric, adParamInput, , null)
		end if

		if Request("hdnDateOrderRecieved") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_date_received", adVarChar, adParamInput,20 , Request("hdnDateOrderRecieved"))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_date_received", adVarChar, adParamInput,20 , null)
		end if

		if Request("hdnScheduledCompletionDate") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_date_pinstalled", adVarChar, adParamInput,20 , Request("hdnScheduledCompletionDate"))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_date_pinstalled", adVarChar, adParamInput,20 , null)
		end if

		if Request("hdnDateInstalled") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_date_installed", adVarChar, adParamInput,20 , Request("hdnDateInstalled"))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_date_installed", adVarChar, adParamInput,20 , null)
		end if


		if Request("hdnDateConfigured") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_date_config", adVarChar, adParamInput,20 , Request("hdnDateConfigured"))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_date_config", adVarChar, adParamInput,20 , null)
		end if

		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput,20 , cdate(Request("hdnUpdateDateTime")))
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_list", adVarChar, adParamOutput, 4000)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_subject", adVarChar, adParamOutput, 4000)
		cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_message", adVarChar, adParamOutput, 4000)

		if Request("hdnDatesocndate") <> "" then
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_date_socn", adVarChar, adParamInput,20 , Request("hdnDatesocndate"))
		else
			cmdUpdateObj.Parameters.Append cmdUpdateObj.CreateParameter("p_date_socn", adVarChar, adParamInput,20 , null)
		end if



		'call the update stored proc
  			'cmdUpdateObj.Parameters.Refresh

  		'	dim objparm
  		'	for each objparm in cmdUpdateObj.Parameters
  		'	  Response.Write "<b>" & objparm.name & "</b>"
  		'	  Response.Write " has size:  " & objparm.Size & " "
  		'	  Response.Write " and value:  " & objparm.value & " "
  		'	  Response.Write " and datatype:  " & objparm.Type & "<br> "
  		'    next


  		 '  Response.Write "<b> count = " & cmdUpdateObj.Parameters.count & "<br>"
  		'	dim nx
  		'	 for nx=0 to cmdUpdateObj.Parameters.count-1
  		'	   Response.Write nx+1 & " parm value= " & cmdUpdateObj.Parameters.Item(nx).Value  & "<br>"
  		'	  next
  			'response.end
		on error resume next
		cmdUpdateObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE Customer Service", objConn.Errors(0).Description
				objConn.Errors.Clear
			else
				' commented out because business users were not sure if they need this here or not.
				dim strEmailFrom, strEmailTo, strEmailSubject, strEmailBody
				'strEmailTo = cmdUpdateObj.Parameters("p_list").Value

				'if strEmailTo <> "" then
				'it's time to send an email
				'	strEmailSubject = cmdUpdateObj.Parameters("p_subject").Value
				'	strEmailBody = cmdUpdateObj.Parameters("p_message").Value
				'	Response.Cookies("txtEmailTo") = strEmailTo
				'	Response.Cookies("txtEmailSubject") = strEmailSubject
				'	Response.Cookies("txtEmailBody") = escape(strEmailBody)
				'end if
			strWinMessage = "Record saved successfully. You can now see the changes you made."
			end if


	else  'create a new record
		if (intAccessLevel and intConst_Access_Create) <> intConst_Access_Create then
			strWinLocation = "CustServDetail.asp?CustServID=0"
			DisplayError "REFRESH", strWinLocation, 0, "INSERT DENIED", "You don't have access to create service locations. Please contact your system administrator."
		end if

		dim cmdInsertObj
		set cmdInsertObj = server.CreateObject("ADODB.Command")
		set cmdInsertObj.ActiveConnection = objConn
		cmdInsertObj.CommandType = adCmdStoredProc
		cmdInsertObj.CommandText = "sma_sp_userid.spk_sma_cserv_inter.sp_cs_insert"

		'create params
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_userid", adVarChar, adParamInput, 20 ,strRealUserID)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_customer_service_id", adNumeric, adParamOutput,,null)
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_remedy_group", adVarChar, adParamInput, 20, Request("selSupportGroup"))
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_cs_desc", adVarChar, adParamInput , 80, left(Request("txtCustomerServiceName"), 80))
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_customer_id", adNumeric, adParamInput, , clng(Request("hdnCustomerID")))
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_service_type", adNumeric, adParamInput, , clng(Request("hdnServiceTypeID")))
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_sla_id", adNumeric ,adParamInput, , clng(Request("hdnSLAID")))
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_service_status", adVarChar, adParamInput, 6, Request("hdnServiceStatusCode"))

		IF Request("hdnServLocID") <> "" THEN
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_sl_id", adNumeric, adParamInput, , clng(Request("hdnServLocID")))
		ELSE
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_sl_id", adNumeric, adParamInput, , NULL)
		END IF

		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_date_inservice", adVarChar, adParamInput,20 , Request("hdnDateInService"))
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_date_terminated", adVarChar, adParamInput, 20 , Request("hdnDateTerminated"))
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_billing_date", adVarChar, adParamInput,20 , Request("hdnBillingStartDate"))

		IF	Request("txtComment") <> "" THEN
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_comments", adVarChar, adParamInput, 2000, Request("txtComment"))
		ELSE
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_comments", adVarChar, adParamInput, 2000, NULL)
		END IF

		IF	Request("txtNoOfSeats") <> "" THEN
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_No_Of_Seats", adNumeric, adParamInput, , clng(Request("txtNoOfSeats")))
		ELSE
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_No_Of_Seats", adNumeric, adParamInput, , NULL)
		END IF

		IF Request("txtOrderNumber") <> "" THEN
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_order_no", adVarChar, adParamInput, 10, Request("txtOrderNumber"))
		ELSE
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_order_no", adVarChar, adParamInput, 10, NULL)
		END IF

		if  Request("hdnContactID1") <> "" THEN
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_design_st", adNumeric, adParamInput, , clng(Request("hdnContactID1")))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_design_st", adNumeric, adParamInput, , null)
		end if

		if Request("hdnContactID2") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_implement_st", adNumeric, adParamInput, , clng(Request("hdnContactID2")))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_implement_st", adNumeric, adParamInput, , null)
		end if

		if Request("hdnDateOrderRecieved") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_date_received", adVarChar, adParamInput,20 , Request("hdnDateOrderRecieved") )
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_date_received", adVarChar, adParamInput,20 , null)
		end if

		if  Request("hdnScheduledCompletionDate") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_date_pinstalled", adVarChar, adParamInput,20 , Request("hdnScheduledCompletionDate"))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_date_pinstalled", adVarChar, adParamInput,20 , null)
		end if

		if  Request("hdnDateInstalled") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_date_installed", adVarChar, adParamInput,20 , Request("hdnDateInstalled"))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_date_installed", adVarChar, adParamInput,20 , null)
		end if

		if  Request("hdnDateConfigured") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_date_config", adVarChar, adParamInput,20 , Request("hdnDateConfigured"))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_date_config", adVarChar, adParamInput,20 , null)
		end if
		cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_repair_priority", adVarChar, adParamInput, 20, Request("selRepairPriority"))

		if  Request("hdnDatesocndate") <> "" then
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_date_socn", adVarChar, adParamInput,20 , Request("hdnDatesocndate"))
		else
			cmdInsertObj.Parameters.Append cmdInsertObj.CreateParameter("p_date_socn", adVarChar, adParamInput,20 , null)
		end if

		on error resume next
		cmdInsertObj.Execute

		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT CREATE Customer Service", objConn.Errors(0).Description
			objConn.Errors.Clear
		else
			logCustomerServiceID = cmdInsertObj.Parameters("p_customer_service_id").Value
		end if
		strWinMessage = "Record created successfully. You can now see the new record."

	on error goto 0

	end if

	case "DELETE"
			if (intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete then
				DisplayError "REFRESH", strWinLocation, 0, "DELETE DENIED", "You don't have access to delete managed objects. Please contact your system administrator"
			end if
			dim cmdDeleteObj
			set cmdDeleteObj = server.CreateObject("ADODB.Command")
			set cmdDeleteObj.ActiveConnection = objConn
			cmdDeleteObj.CommandType = adCmdStoredProc
			cmdDeleteObj.CommandText = "sma_sp_userid.spk_sma_cserv_inter.sp_cs_delete"
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_customer_service_id", adNumeric, adParamInput, , clng(logCustomerServiceID))					'number(9)
			cmdDeleteObj.Parameters.Append cmdDeleteObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, ,Cdate(datUpdateDateTime))		'Date

			'Response.Write "<b> count = " & cmdDeleteObj.Parameters.count & "<br>"
  			'dim nx
  			' for nx=0 to cmdDeleteObj.Parameters.count-1
  			'   Response.Write  " parm " & nx + 1 &  " value= " & cmdDeleteObj.Parameters.Item(nx).Value  & "<br>"
  			'  next
                        on error resume next
			cmdDeleteObj.Execute
			if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE CUSTOMER SERVICE", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if

			logCustomerServiceID = 0
			strWinMessage = "Record deleted successfully."

		on error goto 0

end select

Dim objRsRegion, objRsStatusCode, objRsSupportGroup
Dim strSQL
Dim lDisplayASP, objASP
Dim socnWrite, objsocn

'   strsql = "select b.SECURITY_ROLE_ID " &_
'			"from msaccess.tblsecurity a,  msaccess.staff_security_role  b, msaccess.security_role c " &_
'			"where b.STAFF_ID = a.STAFFID "&_
'			"and b.SECURITY_ROLE_ID = c.SECURITY_ROLE_ID " &_
'			"and c.SECURITY_ROLE_NAME='SMA2 - Customer Care/Service Management'	" &_
'			"and a.USERID = '" & strRealUserID & "'"
'   set objsocn = objConn.Execute(strSQL)
'   if objsocn.EOF then
'       socnWrite = "N"
'   else
'       socnWrite = "Y"
'   end if
   socnWrite = "Y"


   ' Check if Customer Service is ASP
   if logCustomerServiceID = 0 then
      lDisplayASP = "N"
   else
      strsql = "select cs.customer_service_id " &_
               "FROM crp.lob l, crp.service_category sc, crp.service_type st, crp.customer_service cs " &_
               "WHERE l.lob_code = 'ASP' AND l.lob_id = sc.lob_id AND sc.service_category_id = st.service_category_id AND " &_
	               "st.service_type_id = cs.service_type_id AND cs.customer_service_id = " & logCustomerServiceID
      set objASP = objConn.Execute(strSQL)
      if objASP.EOF then
         lDisplayASP = "N"
      else
         lDisplayASP = "Y"
      end if
   end if

   'get a list of region codes
   strsql = "select noc_region_lcode, noc_region_desc " &_
			"from crp.lcode_noc_region " &_
			"where record_status_ind = 'A' " &_
			"order by noc_region_desc"

   set objRsRegion = objConn.Execute(strSQL)

   'get a list of service status codes
   strSQL = "SELECT service_status_code, service_status_name " &_
			 "FROM crp.service_status " &_
			 "WHERE record_status_ind = 'A' " &_
			 "order by service_status_name "

   set objRsStatusCode = objConn.Execute(strSQL)

   'get a list of rememdy support groups
   strSQL = "SELECT remedy_support_group_id, group_name " &_
			  "FROM crp.v_remedy_support_group " &_
			  "order by group_name"

   set objRsSupportGroup = objConn.Execute(strSQL)

Dim strWhereClause, strServLocAddress
'Dim  logCustomerServiceID, strWhereClause, strServLocAddress
 if (logCustomerServiceID <> 0 and logCustomerServiceID <> "" ) then
	StrSql = "select " &_
			 "c.customer_id, " &_
			 "c.customer_name, " & _
			 "c.customer_short_name, " &_
			 "n.noc_region_desc, " &_
 			 "s.customer_service_id, " & _
			 "s.customer_service_desc, " & _
			 "s.service_type_id, " & _
			 "t.service_type_desc, " & _
			 "s.service_level_agreement_id, " &_
			 "sla.service_level_agreement_desc, " & _
			 "s.service_location_id, " & _
			 "l.service_location_name, " & _
			 "a.building_name, " & _
			 "a.street, " & _
			 "a.municipality_name, " & _
			 "z.clli_code, " &_
			 "a.province_state_lcode, " & _
			 "s.lynx_def_sev_lcode, " &_
			 "s.service_status_code, " & _
			 "s.project_code, " & _
			 "s.design_staff_id, " &_
			 "con1.contact_name AS design_contact_name, " &_
			 "con1.first_name as design_first_name, " &_
			 "con1.last_name as design_last_name, " &_
			 "s.implementation_staff_id, " &_
			 "con2.contact_name AS implementation_contact_name, " &_
			 "con2.first_name as implementation_first_name, " &_
			 "con2.last_name as implementation_last_name, " &_
			 "to_char(s.date_workorder_received, 'MON-DD-YYYY') AS date_workorder_received, " &_
			 "to_char(s.date_proposed_installed, 'MON-DD-YYYY') AS date_proposed_installed, " &_
			 "to_char(s.socn_date, 'MON-DD-YYYY') AS date_socn, " &_
			 "to_char(s.date_facility_ordered, 'MON-DD-YYYY') AS date_facility_ordered,  " &_
			 "to_char(s.date_facility_confirmed, 'MON-DD-YYYY') AS date_facility_confirmed,  " &_
			 "to_char(s.date_facility_due, 'MON-DD-YYYY') AS date_facility_due,  " &_
			 "to_char(s.date_facility_ready, 'MON-DD-YYYY') AS date_facility_ready,  " &_
			 "to_char(s.date_installed, 'MON-DD-YYYY') AS date_installed,  " &_
			 "to_char(s.date_in_service, 'MON-DD-YYYY') AS date_in_service, " &_
			 "to_char(s.date_in_service, 'mm/dd/yyyy') AS date_in_service_2, " &_
			 "to_char(s.date_configured, 'MON-DD-YYYY') AS date_configured, " &_
			 "to_char(s.date_terminated, 'MON-DD-YYYY') AS date_terminated, " &_
			 "to_char(s.date_terminated, 'mm/dd/yyyy') AS date_terminated_2, " &_
			 "to_char(s.date_to_start_billing, 'MON-DD-YYYY') AS date_to_start_billing,  " &_
			 "m.missed_installation_cause_desc, " &_
			 "s.remedy_support_group_id, " &_
			 "r.group_name, " &_
			 "s.comments, " &_
			 "s.no_of_seats, " &_
			 "s.record_status_ind, " &_
			 "to_char(s.create_date_time, 'MON-DD-YYYY HH24:MI:SS') AS create_date,  " &_
			 "sma_sp_userid.spk_sma_library.sf_get_full_username(s.create_real_userid) as create_real_userid, " &_
			 "to_char(s.update_date_time, 'MON-DD-YYYY HH24:MI:SS') AS update_date,  " &_
			 "sma_sp_userid.spk_sma_library.sf_get_full_username(s.update_real_userid) as update_real_userid, " &_
			 "s.update_date_time AS last_update_date_time " &_
			 "from crp.customer_service s,  " &_
				"crp.customer c, " &_
				"crp.lcode_noc_region n, " &_
				"crp.service_type t, " &_
				"crp.service_level_agreement sla, " &_
				"crp.service_location l, " &_
				"crp.contact con1, " &_
				"crp.contact con2, " &_
				"crp.v_remedy_support_group r, " &_
				"crp.missed_installation_cause m, " &_
				"crp.v_address_consolidated_street a, " &_
				"crp.municipality_lookup z "

	 strWhereClause = " where	s.customer_id = c.customer_id " &_
						"and	s.service_type_id = t.service_type_id " &_
						"and	s.service_level_agreement_id = sla.service_level_agreement_id " &_
						"and	s.service_location_id = l.service_location_id(+)    " &_
						"and	s.remedy_support_group_id = r.remedy_support_group_id(+) " &_
						"and	s.design_staff_id = con1.contact_id(+) " &_
						"and	s.implementation_staff_id = con2.contact_id(+) " &_
						"and	s.missed_installation_cause_id = m.missed_installation_cause_id(+) " &_
						"and	l.address_id = a.address_id(+) " &_
						"and	a.municipality_name = z.municipality_name(+) " &_
						"and	a.province_state_lcode = z.province_state_lcode(+) " &_
						"and    c.noc_region_lcode = n.noc_region_lcode " &_
						"and	s.customer_service_id = " & logCustomerServiceID

     strSQL =  StrSql & " "& strWhereClause
'   Response.Write strsql
'   Response.End

    if err then
		DisplayError "BACK", "", err.Number, "UNEXPECTED ERROR", err.Description
	end if

   'Create the command object
   dim objRsCustomerService
   set objRsCustomerService = server.CreateObject("ADODB.Recordset")
   objRsCustomerService.CursorLocation = adUseClient
   objRsCustomerService.Open strSQL, objConn
   if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
   end if

    if objRsCustomerService.EOF then
		DisplayError "BACK", "", 999, "CANNOT FIND REQUESTED CUSTOMER SERVICE", "EOF condition occured in rsNE recordset."
		Response.End
   END IF

   ' Check for and use the preferred language translation for the service type, if available
   if (len(strLANG) > 0 and len(objRsCustomerService("service_type_id")) > 0) then
        strSType = objRsCustomerService("service_type_desc")
		strSTypeEN = objRsCustomerService("service_type_desc")

        strSQL = "select service_type_lang_desc " &_
		         "from crp.service_type_lang " &_
				 "where language_preference_lcode like '" & strLANG & "' " &_
				 "and service_type_id = " & objRsCustomerService("service_type_id")

		dim objRsLang
		set objRsLang = server.CreateObject("ADODB.RecordSet")
'		objRsFrench.CursorLocation = adUseClient
		objRsLang.open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText

		if (not err and not objRsLang.EOF) then
 	        	strSType = objRsLang("service_type_lang_desc")
		end if

		objRsLang.Close
		set objRsLang = Nothing
    end if

	strServiceTypeID = objRsCustomerService("service_type_id")


'	strSQL = "select service_type_lang_desc " &_
'		 "from crp.service_type " &_
'		 "where service_type_id = " & objRsCustomerService("service_type_id")
'
'	dim objRsSTypeEN, strSTypeEN
'	set objRsSTypeEN = server.CreateObject("ADODB.RecordSet")
''		objRsFrench.CursorLocation = adUseClient
'	objRsSTypeEN.open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'	if (not err and not objRsSTypeEN.EOF) then
'			strSTypeEN = objRsLang("service_type_desc")
'	end if
'
'	objRsSTypeEN.Close
'	set objRsStypeEN = Nothing

	set objRsCustomerService.ActiveConnection = nothing
    strSQL = "select con.contact_id, " &_
				"c.cust_serv_contact_type_lcode,  " &_
				"l.cust_serv_contact_type_desc,  " &_
				"c.contact_priority, " &_
				"con.contact_name,  " &_
				"con.work_number,  " &_
				"con.work_number_ext,  " &_
				"con.cell_number,  " &_
				"con.pager_number,  " &_
				"con.fax_number,  " &_
				"con.email_address   " &_
			"from crp.customer_service s, " &_
				"crp.customer_service_contact c, " &_
				"crp.contact con, " &_
				"crp.lcode_cust_serv_contact_type l  " &_
			"where s.customer_service_id = c.customer_service_id " &_
			"and c.contact_id = con.contact_id " &_
			"and	c.cust_serv_contact_type_lcode = l.cust_serv_contact_type_lcode " &_
			"and s.customer_service_id = "  &  logCustomerServiceID &_
			"order by cust_serv_contact_type_lcode, contact_priority"

	'Response.Write (strSQL & "<p>")
	'Response.End
  dim objRsCustomerServiceContact
  set objRsCustomerServiceContact = objConn.Execute(strSQL)

  if not objRsCustomerService.EOF then

		if len(objRsCustomerService("building_name") ) > 0 then
			strServLocAddress = objRsCustomerService("building_name") & vbNewLine & objRsCustomerService("street") & vbNewLine &_
						   objRsCustomerService("municipality_name") & " " & objRsCustomerService("province_state_lcode")
		else
			strServLocAddress = objRsCustomerService("street") & vbNewLine & objRsCustomerService("municipality_name") & " " & objRsCustomerService("province_state_lcode")
		end if

	Dim strWPArea,strWPMid,strWPEnd,strWP
	Dim strCPArea,strCPMid,strCPEnd,strCP
	Dim strPPArea,strPPMid,strPPEnd,strPP
	Dim strFPArea,strFPMid,strFPEnd,strFP

		intRowCount = 0
		intColCount = 11
		strInnerValues = ""
		while not objRsCustomerServiceContact.EOF

		'Parse out the work phone number
	 	strWPArea = mid(objRsCustomerServiceContact("work_number"),1,3)
	 	strWPMid = mid(objRsCustomerServiceContact("work_number"),4,3)
	 	strWPEnd = mid(objRsCustomerServiceContact("work_number"),7,4)
	 	strWP = "(" & strWPArea & ") " & strWPMid & "-" & strWPEnd
	 	If strWP = "() -" then
	 		strWP = ""
	 	End If

		'Parse out the cell phone number
		strCPArea = mid(objRsCustomerServiceContact("cell_number"),1,3)
	 	strCPMid = mid(objRsCustomerServiceContact("cell_number"),4,3)
	 	strCPEnd = mid(objRsCustomerServiceContact("cell_number"),7,4)
	 	strCP = "(" & strCPArea & ") " & strCPMid & "-" & strCPEnd
	 	If strCP = "() -" then
	 		strCP = ""
	 	End If


		'Parse out the pager number
		strPPArea = mid(objRsCustomerServiceContact("pager_number"),1,3)
	 	strPPMid = mid(objRsCustomerServiceContact("pager_number"),4,3)
	 	strPPEnd = mid(objRsCustomerServiceContact("pager_number"),7,4)
	 	strPP = "(" & strPPArea & ") " & strPPMid & "-" & strPPEnd
	 	If strPP = "() -" then
	 		strPP = ""
	 	End If

		'Parse out the fax number
	 	strFPArea = mid(objRsCustomerServiceContact("fax_number"),1,3)
	 	strFPMid = mid(objRsCustomerServiceContact("fax_number"),4,3)
	 	strFPEnd = mid(objRsCustomerServiceContact("fax_number"),7,4)
	 	strFP = "(" & strFPArea & ") " & strFPMid & "-" & strFPEnd
	 	If strFP = "() -" then
	 		strFP = ""
	 	End If


			intRowCount = intRowCount + 1
			strInnerValues = strInnerValues & objRsCustomerServiceContact(0) &_
							 strDelimiter & objRsCustomerServiceContact(1) &_
							 strDelimiter & objRsCustomerServiceContact(2) &_
							 strDelimiter & objRsCustomerServiceContact(3) &_
							 strDelimiter & objRsCustomerServiceContact(4) &_
							 strDelimiter & strWP &_
							 strDelimiter & objRsCustomerServiceContact(6) &_
							 strDelimiter & strCP  &_
							 strDelimiter & strPP &_
							 strDelimiter & strFP &_
							 strDelimiter & objRsCustomerServiceContact(10) &_
							 strDelimiter
			objRsCustomerServiceContact.MoveNext
		wend
	objRsCustomerServiceContact.Close
 end if
end if

'get the LYNX default repair priority
dim rsLYNXrp

strSQL = "select LYNX_DEF_SEV_DESC, LYNX_DEF_SEV_LCODE from CRP.LCODE_LYNX_DEF_SEV where RECORD_STATUS_IND='A' ORDER BY LYNX_DEF_SEV_LCODE"
set rsLYNXrp=server.CreateObject("ADODB.Recordset")
rsLYNXrp.CursorLocation = adUseClient
rsLYNXrp.Open strSQL, objConn
if err then
	DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET", err.Description
end if

if rsLYNXrp.EOF then
	DisplayError "BACK", "", 999, "CANNOT CREATE CONTACT ROLE LIST", "EOF condition occured in rsLYNXrp recordset."
end if

set rsLYNXrp.ActiveConnection = nothing

'if (logCustomerServiceID <> 0 and logCustomerServiceID <> "" ) then
'	strServiceTypeID = objRsCustomerService("service_type_id")
'else
'	strServiceTypeID = 0
'end if

'Response.Write "logCustomerServiceID <b>" & logCustomerServiceID & "</b>"
'Response.Write "strServiceTypeID <b>" & strServiceTypeID & "</b>"
'Response.Write "hdnServiceTypeID <b>" & hdnServiceTypeID & "</b>"
'Response End
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript" SRC="AccessLevels.js"></script>
<script type="text/javascript">
<!--
//**************************************** JavaScipt Functions *********************************


var strWinMessage = "<%=strWinMessage%>";
var intAccessLevel = "<%=intAccessLevel%>" ;
var strLANG = "<%=strLANG%>";
var intChildAccessLevel = <%=intChildAccessLevel%>;
var bolNeedToSave = false ;

var intCustServID = <%=logCustomerServiceID%>;
var intServTypeID = '<%=strServiceTypeID%>';
<% if isnumeric(strServiceTypeID) then %>
		intServiceTypeID = <%=strServiceTypeID%> ;
<% end if %>

var strDelimiter='<%=strDelimiter%>';


//display Customer Service in the Heading frame
setPageTitle("SMA - Customer Service");

<%if strEmailSubject <> "" then%>
//pop-up the email window
var wndEmail = window.open('email.asp', 'PopupEmail', 'top=50, left=100, height=610, width=800' ) ;
<%end if%>
function iframe1_display(){
	window.frames["aifr1"].src = 'CustServAlias.asp?cs_id=<%response.write logCustomerServiceID%>';
}

function btn_ifrm1Add(){
	//open a blank form
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create){
		alert('Access denied. Please contact your system administrator.');
		return;
	}
	if (document.frmCustServCPDetail.hdnCustomerServiceID.value == "") {
		alert('At this time you cannot create a name alias. You must save the Customer Service first.');
		return;
	}
	var NewWin;
	var strMasterID = "<%=logCustomerServiceID%>";
	NewWin=window.open("CustServAliasDetail.asp?action=new&masterID="+strMasterID ,"NewWin","toolbar=no,status=yes,width=700px,height=175px,left=150px,top=200,menubar=no,resize=no");
	NewWin.focus();
}

function btn_ifrm1Update(){
	//open a detail form where the user can modify the alias
	if ((intAccessLevel & intConst_Access_Update) != intConst_Access_Update){
		alert('Access denied. Please contact your system administrator.');
		return;
	}
	var NewWin;
	var strAliasID = window.frames["aifr1"].contentDocument.frmIFR.hdnNameAliasID.value;
	if (strAliasID == "") {
		alert("Please select an alias or click ADD NEW to create a new alias.");
		return;
	}
	var strMasterID = "<%=logCustomerServiceID%>";
	NewWin=window.open("CustServAliasDetail.asp?action=update&aliasID="+strAliasID+"&masterID="+strMasterID ,"NewWin","toolbar=no,status=yes,width=700px,height=175px,left=150px,top=200,menubar=no,resize=yes");
	NewWin.focus();
}

function btn_ifrm1Delete(){
//delete selected row
	if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete){
		alert('Access denied. Please contact your system administrator.');
		return;
	}
	var strAliasID = window.frames["aifr1"].contentDocument.frmIFR.hdnNameAliasID.value;
	if (strAliasID == "") {
		alert("Please select an alias or click ADD NEW to create a new alias.");
		return;
	}
	var strLastUpdate = window.frames["aifr1"].contentDocument.frmIFR.hdnLastUpdate.value;
	if (confirm("Do you want to delete this alias?")){
		window.frames["aifr1"].contentDocument.location.href = "CustServAliasDetail.asp?action=delete&back=true&aliasID="+strAliasID+"&masterID=<%response.write logCustomerServiceID%>&hdnLastUpdate="+strLastUpdate;
	}
}

function iFrame_display()
{
//called whenever a refresh of the iframe is needed
	window.frames["aifr2"].src = 'CustServContList.asp?CustServID=' + document.frmCustServCPDetail.hdnCustomerServiceID.value;
}

function btn_ifrmAdd()
{

	if ((intChildAccessLevel & intConst_Access_Create) != intConst_Access_Create)
	{
		alert('Access denied.  Please contact your system administrator.');
		return false;
	}

	var NewWin;
	NewWin=window.open("CustServContDetail.asp?NewContact=NEW&CustServID=" + "<%=logCustomerServiceID%>" + "&txtWorkFor=" + document.frmCustServCPDetail.txtCustomerName.value, "NewWin","toolbar=no,status=no,width=700,height=430,menubar=no resize=no");
	NewWin.focus();
}


function btn_ifrmUpdate(){

	var NewWin;

	if ((intChildAccessLevel & intConst_Access_Update) != intConst_Access_Update)
	{
		alert('Access denied.  Please contact your system administrator.');
		return false;
	}

	if (window.frames["aifr2"].contentDocument.frmIFR.hdnContactID.value !="")
	{

		var strSource ="CustServContDetail.asp?CustServContactID="+window.frames["aifr2"].contentDocument.frmIFR.hdnContactID.value;
		NewWin=window.open(strSource ,"NewWin","toolbar=no,status=no,width=700,height=430,menubar=no resize=no");
		NewWin.focus();
	}
	else
	{
		alert('You must select a record to update!');
	}

}

function btn_ifrmDelete()
{

	if ((intChildAccessLevel & intConst_Access_Delete) != intConst_Access_Delete)
	{
		alert('Access denied.  Please contact your system administrator.');
		return false;
	}

	if (window.frames["aifr2"].contentDocument.frmIFR.hdnContactID.value !="")
	{
		if (confirm('Do you really want to delete this Contact?'))
		{
			window.frames["aifr2"].src = "CustServContList.asp?txtFrmAction=DELETE&CustServID=" + document.frmCustServCPDetail.hdnCustomerServiceID.value + "&ContactID="+window.frames["aifr2"].contentDocument.frmIFR.hdnContactID.value+"&hdnUpdateDateTime="+window.frames["aifr2"].contentDocument.frmIFR.hdnUpdateDateTime.value;
		}
	}
	else
	{
		alert('You must select a record to delete!');
	}
}

function window_onload() {

	iframe1_display();
	iFrame_display();
	iframe1a_display();
	iframe1b_display();
	iframe3_display();
	iframe4_display()
}

function fct_clearStatus() {
		window.status = "";
	}

function fct_displayStatus(strWinStatus){
		window.status=strWinStatus;
		setTimeout('fct_clearStatus()',5000);
	}


function btnServiceTypeLookup_onclick() {
	var logServiceTypeID = document.frmCustServCPDetail.hdnServiceTypeID.value ;
	var strServiceTypeDesc = document.frmCustServCPDetail.txtServiceType.value ;

	if ( logServiceTypeID != "" ) {
		SetCookie("ServiceType", logServiceTypeID) ;
		SetCookie("STypeDesc", strServiceTypeDesc );
	}
	SetCookie("WinName", "Popup") ;
	fct_onChange();
	window.open('SearchFrame.asp?fraSrc=ServiceType','Popup','top=50, left=100, height=600, width=850') ;
	}


function fct_lookupCustomer(CustService){

    SetCookie("ServiceEnd", CustService);
	if (document.frmCustServCPDetail.txtCustomerName.value != "")
		{SetCookie("CustomerName", document.frmCustServCPDetail.txtCustomerName.value);
		}

	fct_onChange();
	SetCookie("WinName", 'Popup');
	window.open('SearchFrame.asp?fraSrc=Cust', 'Popup', 'top=50, left=100, height=600, width=800' ) ;

}

function fct_onDelete() {
var logCustomerServiceID = document.frmCustServCPDetail.txtCustomerServiceID.value ;
var strUpdateDateTime = document.frmCustServCPDetail.hdnUpdateDateTime.value ;

 if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
	alert('Access denied. Please contact your system administrator.');
	return;
   }


	if (confirm('Do you really want to delete this object?')){
		document.location = "CustServDetail.asp?hdnFrmAction=DELETE&CustServID="+logCustomerServiceID+"&UpdateDateTime="+strUpdateDateTime ;
	}
}

function btnGuess_onclick() {
	var strCustomerShortName = document.frmCustServCPDetail.txtCustomerShortName.value ;
	var strCityShortName = document.frmCustServCPDetail.hdnClliCode.value ;
	var strProvinceCode = document.frmCustServCPDetail.hdnProvinceCode.value ;
	var strBuildingName = document.frmCustServCPDetail.hdnBuildingName.value ;
	var strStreetName =  document.frmCustServCPDetail.hdnStreetName.value ;
//	var strServiceType = document.frmCustServCPDetail.txtServiceType.value ;
	var strServiceType ;
	var strAddress ;
	var strSuggestedName ;
	var strLen ;

	if (document.frmCustServCPDetail.hdnSTypeEN.value == "") {
		strServiceType = document.frmCustServCPDetail.txtServiceType.value ; }
	else {
		strServiceType =  document.frmCustServCPDetail.hdnSTypeEN.value ;
	}

	strLen = strProvinceCode.length ;
	if (strLen > 2 ) {
		strProvinceCode= strProvinceCode.substr(0, 1) ;
	}

	if (strProvinceCode == "QC") {
		strProvinceCode = "PQ";
	}

	if (strProvinceCode == "NL") {
		strProvinceCode = "NF";
	}

	if ( document.frmCustServCPDetail.txtCustomerName.value == "" ) {
		alert('Please first select a Customer using the Lookup button.');
		document.frmCustServCPDetail.btnCustomerLookup.focus();
		return(false);
	}

	if ( document.frmCustServCPDetail.txtServLocName.value == "" ){
		alert('Please first select a Service Locaiton using the Lookup button.');
		document.frmCustServCPDetail.btnServiceLocationLookup.focus();
		return(false);
	}

	if (document.frmCustServCPDetail.txtServiceType.value == "") {

		alert('Please first select a Service Type using the Lookup button.');
		document.frmCustServCPDetail.btnServiceTypeLookup.focus();
		return(false);
	}

	if (strBuildingName == "") {
		strAddress = strStreetName ; }
	else {
		strAddress = strBuildingName ;
	}


	fct_onChange();
	strSuggestedName = strCustomerShortName + '_' + strCityShortName + strProvinceCode + '_' + strAddress + '_' + strServiceType ;

	document.frmCustServCPDetail.txtCustomerServiceName.value =  strSuggestedName;

}

function btnNetcrackerweblink_onclick() {
	var strCustomerServiceID = document.frmCustServCPDetail.hdnCustomerServiceID.value ;
	var strNetcrackerURL = '<%=strConstNetcrackerURL%>';
	var strNetcrackerURLCSID ;

	strNetcrackerURCSID = strNetcrackerURL + strCustomerServiceID

	window.open(strNetcrackerURCSID);
}

function btnSLALookup_onclick() {

var	strSLADesc = window.frmCustServCPDetail.txtServiceLevelAgreement.value ;

	if (strSLADesc != "" ) {
		SetCookie("SLADesc", strSLADesc);
	}
	SetCookie("WinName", "Popup") ;
	fct_onChange();
	window.open('SearchFrame.asp?fraSrc=SLA', 'Popup', 'top=50, left=100, WIDTH=800, HEIGHT=600'  ) ;
}

function btnServiceLocationLookup_onclick(strServiceEnd) {
  var strServiceLocationName = document.frmCustServCPDetail.txtServLocName.value ;
  var strCustomerName = document.frmCustServCPDetail.txtCustomerName.value ;


	if ( strCustomerName != "" ) {
		SetCookie("CustomerName", strCustomerName);  }

	SetCookie("IncludeTelus", "yes");
	SetCookie("ServiceEnd", strServiceEnd);
	SetCookie("WinName", "Popup") ;

	fct_onChange();
	window.open('SearchFrame.asp?fraSrc=ServLoc','Popup','top=50, left=100, height=600, width=800') ;

}

function DesignSpecialistContactlookup(){

	if (document.frmCustServCPDetail.txtLName1.value != ""){
		 SetCookie("LName", document.frmCustServCPDetail.txtLName1.value);
	}
	if (document.frmCustServCPDetail.txtFName1.value != ""){
		 SetCookie("FName", document.frmCustServCPDetail.txtFName1.value);
	}
	SetCookie("WinName", 'Popup');
	SetCookie("Case", "A");
	fct_onChange();
	window.open('SearchFrame.asp?fraSrc=Contact', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
}

function ImplementationConactlookup() {
	if (document.frmCustServCPDetail.txtLName2.value != ""){
		 SetCookie("LName", document.frmCustServCPDetail.txtLName2.value);
	}
	if (document.frmCustServCPDetail.txtFName2.value != ""){
		 SetCookie("FName", document.frmCustServCPDetail.txtFName2.value);
	}
	SetCookie("WinName", 'Popup');
	SetCookie("Case", "B");
	fct_onChange();
	window.open('SearchFrame.asp?fraSrc=Contact', 'Popup', 'top=50, left=100, height=600, width=800' ) ;
}

function selNavigate_onchange(){
//***********************************************************************************************
// Function:	selNavigate_onchange															*
//																								*
// Purpose:		To display the page selected by the user from Quick Navigation drop-down box.	*
//																								*
// Created By:	Sara Sangha	Aug. 25th, 2000														*
//																								*
// Updated By:																					*
//***********************************************************************************************

 var strPageName = document.frmCustServCPDetail.selNavigate.item(document.frmCustServCPDetail.selNavigate.selectedIndex).value ;
 var strCustomerID =  document.frmCustServCPDetail.hdnCustomerID.value ;
 var strServLocID = document.frmCustServCPDetail.hdnServLocID.value ;
 var strCustomerServiceName = document.frmCustServCPDetail.txtCustomerServiceName.value ;
 var strCustomerName = document.frmCustServCPDetail.txtCustomerName.value;
 var logCustomerServiceID = document.frmCustServCPDetail.txtCustomerServiceID.value ;
 var strServiceLocationName = document.frmCustServCPDetail.txtServLocName.value;
 var strAddress =  document.frmCustServCPDetail.txtServLocAddress.value

	switch ( strPageName ) {

	case 'Cust' :
	    document.frmCustServCPDetail.selNavigate.selectedIndex=0;
		self.location.href = 'CustDetail.asp?CustomerID=' + strCustomerID ;
		break ;

	case 'ServLoc' :
		if ( strServLocID  != "" ) {
		    document.frmCustServCPDetail.selNavigate.selectedIndex=0;
			self.location.href = 'ServLocDetail.asp?ServLocID=' + strServLocID ; }
		else
			{ alert("Unexpected Error: \nDo not have enough information to move forward"); }
		break ;

	case 'Facility' :
	    document.frmCustServCPDetail.selNavigate.selectedIndex=0;
	    SetCookie("CustomerServiceA", strCustomerServiceName);
		SetCookie("CustomerServA", strCustomerServiceName);
		SetCookie("CustomerServID", logCustomerServiceID);
		SetCookie("CustName", strCustomerName);
		SetCookie("CustID", strCustomerID);
		SetCookie("ServiceLocName",strServiceLocationName);
		SetCookie("ServiceLocID",strServLocID);
		SetCookie("Address",strAddress);

		self.location.href = 'SearchFrame.asp?fraSrc=' + strPageName ;
		break ;

	case 'FacilityPVC' :
	    document.frmCustServCPDetail.selNavigate.selectedIndex=0;
		SetCookie("CustomerServiceA", strCustomerServiceName);
		SetCookie("CustomerServA", strCustomerServiceName);
		SetCookie("CustomerServID", logCustomerServiceID);
		SetCookie("CustName", strCustomerName);
		SetCookie("CustID", strCustomerID);
		SetCookie("ServiceLocName",strServiceLocationName);
		SetCookie("ServiceLocID",strServLocID);
		SetCookie("Address",strAddress);
		self.location.href = 'SearchFrame.asp?fraSrc=' + strPageName ;
		break ;

	case 'Correlation' :
	    document.frmCustServCPDetail.selNavigate.selectedIndex=0;
		self.location.href = 'corrdetail.asp?CustomerServiceID=' + document.frmCustServCPDetail.txtCustomerServiceID.value ;
		break ;

	case 'CorrelationVpn' :
	    document.frmCustServCPDetail.selNavigate.selectedIndex=0;
		self.location.href = 'corrcpdetail.asp?CustomerServiceID=' + document.frmCustServCPDetail.txtCustomerServiceID.value ;
		break ;

	case 'OrderHistory' :
	    document.frmCustServCPDetail.selNavigate.selectedIndex=0;
	    SetCookie("CustomerServiceID", logCustomerServiceID);
		//SetCookie("CustomerServiceName", strCustomerServiceName);
		self.location.href = 'SearchFrame.asp?fraSrc=' + strPageName ;
		break ;

	case 'CorrelationRoot' :
	    document.frmCustServCPDetail.selNavigate.selectedIndex=0;
	    SetCookie("Type", "CustServ");
		SetCookie("ObjectName", document.frmCustServCPDetail.txtCustomerServiceName.value);
		self.location.href = 'SearchFrame.asp?fraSrc=Correlation'  ;
		break ;

	case 'ManagedObjects':  //to a list
	    document.frmCustServCPDetail.selNavigate.selectedIndex=0;
		SetCookie("CustomerName", strCustomerName);
		self.location.href = "SearchFrame.asp?fraSrc=" + strPageName  ;
		break;

	case 'DEFAULT' :
		// do nothing ;
	}

}

function btnImplemenationClear_onclick(){
	fct_onChange();
	document.frmCustServCPDetail.txtContactName2.value= "" ;
	document.frmCustServCPDetail.hdnContactID2.value = "";
	document.frmCustServCPDetail.txtLName2.value = "";
	document.frmCustServCPDetail.txtFName2.value = "";
}

function btnDesignSpecialistClear_onclick(){
	fct_onChange();
	document.frmCustServCPDetail.txtContactName1.value ="" ;
	document.frmCustServCPDetail.hdnContactID1.value ="";
	document.frmCustServCPDetail.txtLName1.value = "";
	document.frmCustServCPDetail.txtFName1.value = "";

}

function btnCalendar_onclick(intDateFieldNo) {
	var NewWin;
		fct_onChange();
	    SetCookie("Field",intDateFieldNo);
		NewWin=window.open("calendar.asp","NewWin","toolbar=no,status=no,width=260,height=225,menubar=no resize=no");
	NewWin.focus();
}

function fct_onClone() {

	if (document.frmCustServCPDetail.hdnCustomerServiceID.value == "" || document.frmCustServCPDetail.hdnCustomerServiceID.value == "0" || bolNeedToSave)
	{
		if (window.confirm ('There is unsaved data. To Save, press OK and then click on Clone to clone from saved record.') )
		{
			document.frmCustServCPDetail.btnSave.click();
		}

		return(false);
	}

	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('Access denied. Please contact your system administrator.');
		return(false);
	}

	document.location.href="CustServDetail.asp?NewCustServ=CLONED&CustServID=" + document.frmCustServCPDetail.hdnCustomerServiceID.value;

	alert("Record Cloned. Please make changes then save!");

}

function fct_onNew() {
	if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
		alert('Access denied. Please contact your system administrator.');
		return;
	}
	self.document.location.href ="CustServDetail.asp?CustServID=0" ;
}

function fct_onChange(){
	bolNeedToSave = true ;
}


function ClearStatus() {
	window.status = "";
}

function DisplayStatus(strWindowStatus){
	window.status=strWindowStatus;
	setTimeout('ClearStatus()', "<%=intConst_MessageDisplay%>");
}
function body_onbeforeunload() {

	//must set focus to save button because is user has changed only one field and has not left it the on_change event will not have fired and the flag that //determines whether you need to save or not will be false
	document.frmCustServCPDetail.btnSave.focus();
	if  ( bolNeedToSave == true ) {
		if (((intAccessLevel & "<%=intConst_Access_Create%>") == "<%=intConst_Access_Create%>") || ((intAccessLevel & "<%=intConst_Access_Update%>") == "<%=intConst_Access_Update%>") ){
				event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
		}
	}
}

function btnReset_onclick(){
	if(confirm('All changes will be lost. Do you really want to reset this page?')){
		bolNeedToSave = false;
		document.location = 'CustServCPDetail.asp?CustServID=' + "<%=logCustomerServiceID%>" ;
	}
}


function btnReferences_onclick() {
var strOwner = 'CRP' ;
var strTableName = 'CUSTOMER_SERVICE' ;
var strRecordID = document.frmCustServCPDetail.hdnCustomerServiceID.value ;
var URL ;

	if (strRecordID != "" ) {
		URL ='Dependency.asp?Owner=' + strOwner + '&TableName=' + strTableName + '&RecordID='+ strRecordID   ;
		window.open(URL, 'Popup', 'toolbar=no, status=no, top=100, left=100, WIDTH=500, HEIGHT=300, menubar=no '  ) ; }
	else {
		alert("No references. This is a new record."); }

}

function form_onsubmit(){
var strDay, strMonth, strYear, strDate

 if (((intAccessLevel & intConst_Access_Create) == intConst_Access_Create) || ((intAccessLevel & intConst_Access_Update) == intConst_Access_Update) )
  {
	if (document.frmCustServCPDetail.txtCustomerName.value == "" ) {
		alert('Please select a customer using lookup function');
		document.frmCustServCPDetail.btnCustomerLookup.focus();
		return(false);}

	if (document.frmCustServCPDetail.txtServLocName.value == "" ) {
		alert('Please select a Service Location using lookup function');
		document.frmCustServCPDetail.btnServiceLocationLookup.focus();
		return(false);}

	if (document.frmCustServCPDetail.txtServiceType.value == "" ) {
		alert('Please select a Service Type using lookup function');
		document.frmCustServCPDetail.btnServiceTypeLookup.focus() ;
		return(false);}

	if (document.frmCustServCPDetail.txtCustomerServiceName.value == "" ) {
		alert('Please enter a unique customer service name.');
		document.frmCustServCPDetail.txtCustomerServiceName.focus();
		return(false);}

	if (document.frmCustServCPDetail.txtServiceLevelAgreement.value == "" ) {
		alert('Please select an SLA using lookup function.');
		document.frmCustServCPDetail.btnSLALookup.focus();
		return(false);}

	if (document.frmCustServCPDetail.selSupportGroup.selectedIndex == 0  ) {
		alert('Please select a support group from the drop-down list.');
		document.frmCustServCPDetail.selSupportGroup.focus() ;
		return(false);}

	if (document.frmCustServCPDetail.selServiceStatus.selectedIndex == 0 ) {
		alert('Please enter a service status from the drop-down list.');
		document.frmCustServCPDetail.selServiceStatus.focus();
		return(false);}

	//Date Start Billing
	strDay = document.frmCustServCPDetail.selday.item(document.frmCustServCPDetail.selday.selectedIndex).value;
	strMonth = document.frmCustServCPDetail.selmonth.item(document.frmCustServCPDetail.selmonth.selectedIndex).value;
	strYear = document.frmCustServCPDetail.selyear.item(document.frmCustServCPDetail.selyear.selectedIndex).value;

	if ((strMonth != "") & (strDay !="") & (strYear !=""))
	  {
		strDate = strMonth + "/" + strDay + "/" + strYear;
		document.frmCustServCPDetail.hdnBillingStartDate.value = strDate;
	  }
	else
	  if ((strMonth != "")||(strDay != "" || strYear != ""  )) {
	      alert('Please enter a valid Date Start Billing');
	      document.frmCustServCPDetail.selmonth.focus();
	      return(false);
		  }
	  else
		 { document.frmCustServCPDetail.hdnBillingStartDate.value = ""; }

	//if (document.frmCustServCPDetail.selmonth2.item(document.frmCustServCPDetail.selmonth2.selectedIndex).value != "")
	//  {
	//	strDay = document.frmCustServCPDetail.selday2.item(document.frmCustServCPDetail.selday2.selectedIndex).value;
	//	strMonth = document.frmCustServCPDetail.selmonth2.item(document.frmCustServCPDetail.selmonth2.selectedIndex).value;
	//	strYear = document.frmCustServCPDetail.selyear2.item(document.frmCustServCPDetail.selyear2.selectedIndex).value;
	//
	//	strDate = strMonth + "/" + strDay + "/" + strYear;
	//	document.frmCustServCPDetail.hdnDateInService.value = strDate;
	// }
	//else
	//	{ document.frmCustServCPDetail.hdnDateInService.value = ""; }

	//if (document.frmCustServCPDetail.selmonth3.item(document.frmCustServCPDetail.selmonth3.selectedIndex).value != "")
	//  {
	//	strDay = document.frmCustServCPDetail.selday3.item(document.frmCustServCPDetail.selday3.selectedIndex).value;
	//	strMonth = document.frmCustServCPDetail.selmonth3.item(document.frmCustServCPDetail.selmonth3.selectedIndex).value;
	//	strYear = document.frmCustServCPDetail.selyear3.item(document.frmCustServCPDetail.selyear3.selectedIndex).value;
	//
	//	strDate = strMonth + "/" + strDay + "/" + strYear;
	//	document.frmCustServCPDetail.hdnDateTerminated.value = strDate;
	//  }
	//else
	//	{ document.frmCustServCPDetail.hdnDateTerminated.value = ""; }

	//Date Order Received
	strDay = document.frmCustServCPDetail.selday4.item(document.frmCustServCPDetail.selday4.selectedIndex).value;
	strMonth = document.frmCustServCPDetail.selmonth4.item(document.frmCustServCPDetail.selmonth4.selectedIndex).value;
	strYear = document.frmCustServCPDetail.selyear4.item(document.frmCustServCPDetail.selyear4.selectedIndex).value;

	if ((strMonth != "") & (strDay !="") & (strYear !=""))
	  {
		strDate = strMonth + "/" + strDay + "/" + strYear;
		document.frmCustServCPDetail.hdnDateOrderRecieved.value = strDate;
	  }
	else
	  if ((strMonth != "")||(strDay != "" || strYear != ""  )) {
	      alert('Please enter a valid Date Order Received');
	      document.frmCustServCPDetail.selmonth4.focus();
	      return(false);
		  }
	  else
		{ document.frmCustServCPDetail.hdnDateOrderRecieved.value = ""; }

	//Scheduled Completion Date
	strDay = document.frmCustServCPDetail.selday5.item(document.frmCustServCPDetail.selday5.selectedIndex).value;
	strMonth = document.frmCustServCPDetail.selmonth5.item(document.frmCustServCPDetail.selmonth5.selectedIndex).value;
	strYear = document.frmCustServCPDetail.selyear5.item(document.frmCustServCPDetail.selyear5.selectedIndex).value;

	if ((strMonth != "") & (strDay !="") & (strYear !=""))
	  {
		strDate = strMonth + "/" + strDay + "/" + strYear;
		document.frmCustServCPDetail.hdnScheduledCompletionDate.value = strDate;
	  }
	else
	  if ((strMonth != "")||(strDay != "" || strYear != ""  )) {
	      alert('Please enter a valid Scheduled Completion Date');
	      document.frmCustServCPDetail.selmonth5.focus();
	      return(false);
		  }
	  else
		{ document.frmCustServCPDetail.hdnScheduledCompletionDate.value = ""; }

	//Date Configured
	strDay = document.frmCustServCPDetail.selday6.item(document.frmCustServCPDetail.selday6.selectedIndex).value;
	strMonth = document.frmCustServCPDetail.selmonth6.item(document.frmCustServCPDetail.selmonth6.selectedIndex).value;
	strYear = document.frmCustServCPDetail.selyear6.item(document.frmCustServCPDetail.selyear6.selectedIndex).value;

	if ((strMonth != "") & (strDay !="") & (strYear !=""))
	  {
		strDate = strMonth + "/" + strDay + "/" + strYear;
		document.frmCustServCPDetail.hdnDateConfigured.value = strDate;
	  }
	else
	  if ((strMonth != "")||(strDay != "" || strYear != ""  )) {
	      alert('Please enter a valid Date Configured');
	      document.frmCustServCPDetail.selmonth6.focus();
	      return(false);
		  }
	  else
		{ document.frmCustServCPDetail.hdnDateConfigured.value = ""; }

	//Date Installed
	strDay = document.frmCustServCPDetail.selday7.item(document.frmCustServCPDetail.selday7.selectedIndex).value;
	strMonth = document.frmCustServCPDetail.selmonth7.item(document.frmCustServCPDetail.selmonth7.selectedIndex).value;
	strYear = document.frmCustServCPDetail.selyear7.item(document.frmCustServCPDetail.selyear7.selectedIndex).value;

	if ((strMonth != "") & (strDay !="") & (strYear !=""))
	  {
		strDate = strMonth + "/" + strDay + "/" + strYear;
		document.frmCustServCPDetail.hdnDateInstalled.value = strDate;
	  }
	else
	  if ((strMonth != "")||(strDay != "" || strYear != ""  )) {
	      alert('Please enter a valid Date Installed');
	      document.frmCustServCPDetail.selmonth7.focus();
	      return(false);
		  }
	  else
		{ document.frmCustServCPDetail.hdnDateInstalled.value = ""; }

	//date2b
	strDay = document.frmCustServCPDetail.selday8.item(document.frmCustServCPDetail.selday8.selectedIndex).value;
	strMonth = document.frmCustServCPDetail.selmonth8.item(document.frmCustServCPDetail.selmonth8.selectedIndex).value;
	strYear = document.frmCustServCPDetail.selyear8.item(document.frmCustServCPDetail.selyear8.selectedIndex).value;

	if ((strMonth != "") & (strDay !="") & (strYear !=""))
	  {
		strDate = strMonth + "/" + strDay + "/" + strYear;
		document.frmCustServCPDetail.hdnDatesocndate.value = strDate;
	  }
	else
	  if ((strMonth != "")||(strDay != "" || strYear != ""  )) {
	      alert('Please enter a valid Date Installed');
	      document.frmCustServCPDetail.selmonth8.focus();
	      return(false);
		  }
	  else
		{ document.frmCustServCPDetail.hdnDatesocndate.value = ""; }


	document.frmCustServCPDetail.hdnServiceStatusCode.value = document.frmCustServCPDetail.selServiceStatus.value ;
	bolNeedToSave = false  //otherwise you get a prompt to save on the onload function.
	document.frmCustServCPDetail.hdnFrmAction.value = "SAVE" ;
	document.frmCustServCPDetail.submit();
	return(true);
  	}
  	else {
  		alert('Access denied. Please contact your system administrator.');
  		return(false);
  	  }
}

function iframe1a_display(){
	window.frames["aifr1a"].src = 'CorrUsageList.asp?ServiceTypeID=' + intServTypeID;
}

function iframe1b_display(){
	window.frames["aifr1b"].src = 'CorrSOInstList.asp?ServiceTypeID=' + intServTypeID + '&CustomerServiceID=' + intCustServID;
}

function iframe3_display(){
	window.frames["waifr3"].src = 'CorrSOWInstList.asp?ServiceTypeID=' + intServTypeID + '&CustomerServiceID=' + intCustServID;
}

function iframe4_display()
{
//called whenever a refresh of the iframe is needed
	window.frames["aifrvpnws1"].src = 'CustServVpnList.asp?CustomerServiceID=' + intCustServID;
;
}


//***************************************** End of JavaScript Functions *************************
//-->
</SCRIPT>
</HEAD>
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<BODY LANGUAGE=javascript onload="window_onload();" onbeforeunload="body_onbeforeunload();">
<FORM name=frmCustServCPDetail action="CustServCPDetail.asp" method="POST" >

	<!--hidden variables -->
	<INPUT name=hdnCustomerServiceID type=hidden value=<%if logCustomerServiceID <> 0 and not bolCloned then  Response.Write """"&objRsCustomerService("customer_service_id")&"""" else Response.Write null end if%> >
	<INPUT name=hdnCustomerID type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("customer_id")&"""" else Response.Write null end if%> >
	<INPUT name=txtCustomerShortName type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("customer_short_name")&"""" else Response.Write null end if%> >
	<INPUT name=hdnServLocID type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("service_location_id")&"""" else Response.Write null end if%>>
	<INPUT name=hdnStatusCode type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("service_status_code")&"""" else Response.Write null end if%>>
	<INPUT name=hdnClliCode type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("clli_code")&"""" else Response.Write null end if%>>
	<INPUT name=hdnProvinceCode type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("province_state_lcode")&"""" else Response.Write null end if%>>
	<INPUT name=hdnBuildingName type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("building_name")&"""" else Response.Write null end if%> >
	<INPUT name=hdnStreetName type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("Street")&"""" else Response.Write null end if%> >
	<INPUT name=hdnServiceTypeDesc type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&routineHtmlString(objRsCustomerService("service_type_desc"))&"""" else Response.Write null end if%> >
	<INPUT name=hdnSTypeEN type=hidden value=<%if logCustomerServiceID <> 0 then Response.Write """"&routineHtmlString(strSTypeEN)&"""" else Response.write null end if%>>
	<INPUT name=hdnServiceTypeID type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("service_type_id")&"""" else Response.Write null end if%>>
	<INPUT name=hdnSLAID type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("service_level_agreement_id")&"""" else Response.Write null end if%>>
	<INPUT name=hdnSupportGroupID type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("remedy_support_group_id")&"""" else Response.Write null end if%>>
	<INPUT name=hdnContactID1 type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("design_staff_id")&"""" else Response.Write null end if%>>
	<INPUT name=txtFName1 type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("design_first_name")&"""" else Response.Write null end if%>>
	<INPUT name=txtLName1 type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("design_last_name")&"""" else Response.Write null end if%>>
	<INPUT name=hdnContactID2 type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("implementation_staff_id")&"""" else Response.Write null end if%>>
	<INPUT name=txtFName2 type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("implementation_first_name")&"""" else Response.Write null end if%>>
	<INPUT name=txtLName2 type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("implementation_last_name")&"""" else Response.Write null end if%>>
	<INPUT name=hdnFrmAction id=hdnFrmAction  type=hidden value= "">
	<INPUT name=hdnUpdateDateTime id=hdnUpdateDateTime type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("last_update_date_time")&"""" else Response.Write null end if%>>
	<INPUT name=hdnBillingStartDate id=hdnBillingStartDate type=hidden value="" >
	<INPUT name=hdnDateInService id=hdnDateInService type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("date_in_service_2")&"""" else Response.Write null end if%>>
	<INPUT name=hdnDateTerminated id=hdnDateTerminated type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("date_terminated_2")&"""" else Response.Write null end if%>>
	<INPUT name=hdnDateOrderRecieved  id=hdnDateOrderRecieved type=hidden value="">
	<INPUT name=hdnScheduledCompletionDate id=hdnScheduleDate type=hidden value="">
	<INPUT name=hdnDateConfigured id=hdnDateConfigured type=hidden value="">
	<INPUT name=hdnDateInstalled id=hdnDateInstalled type=hidden value="">
	<INPUT name=hdnDatesocndate id=hdnDatesocndate type=hidden value="">
	<INPUT name=hdnServiceStatusCode id=hdnServiceStatusCode type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("service_status_code")&"""" else Response.Write null end if%>>
	<INPUT name=hdnNoOfSeats type=hidden value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("no_of_seats")&"""" else Response.Write null end if%>>

<TABLE border=0>
	<thead>
		<tr><td align=left colspan=3>Customer Service Detail</td>
		<td><SELECT ALIGN=RIGHT id=selNavigate name=selNavigate tabindex=52  <%if logCustomerServiceID = 0 then  Response.Write " disabled " end if %> LANGUAGE=javascript onchange="return selNavigate_onchange()" tabindex=52 >
				<OPTION value='DEFAULT'>Quickly Goto...</OPTION>
				<OPTION value=Cust>Customer</OPTION>
				<OPTION value=ServLoc>Service Location</OPTION>
				<OPTION value=Facility>Facility</OPTION>
				<OPTION value=ManagedObjects>Managed Object</OPTION>
				<OPTION value=FacilityPVC>PVC</OPTION>
				<OPTION value=OrderHistory>Order History</OPTION>
				<OPTION value=Correlation>Correlation</OPTION>
				<OPTION value=CorrelationRoot>Correlation(Root)</OPTION>
				<OPTION value=CorrelationVpn>Correlation(VPN)</OPTION>
		   </SELECT></td>
		</tr>
	</thead>

	<TR>
		<TD align=right>Customer Service Name<font color=red>*</font></TD>
        <TD align=left>
			<INPUT id=txtCustomerServiceName name=txtCustomerServiceName tabindex=1 style="HEIGHT: 22px; WIDTH: 425px"
				value=<%if logCustomerServiceID <> 0 then  Response.Write """"&routineHtmlString(objRsCustomerService("customer_service_desc"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();">
			<INPUT id=btnGuess name=btnGuess tabindex=2 style="HEIGHT: 22px; WIDTH: 50px" type=button value=Guess LANGUAGE=javascript onclick="return btnGuess_onclick()"></TD>

		<td rowspan="5" valign="top" align="right">CS Name Alias</td>
		<td rowSpan="5" valign="top">
			<iframe id=aifr1 width=100% height=100% src="" scrolling=yes marginheight=1 marginwidth=1></iframe>
			<br>
			<input type="button" value="Refresh" name="btn_iframe1Refresh" onClick="iframe1_display();" class=button>
			<input type="button" value="New"     name="btn_iframe1Add"     onClick="btn_ifrm1Add();fct_onChange();"    class=button>
			<input type="button" value="Update"  name="btn_iframe1Update"  onClick="btn_ifrm1Update();" class=button>
			<input type="button" value="Delete"  name="btn_iframe1Delete"  onClick="btn_ifrm1Delete();" class=button>
		</td>
    </TR>

	<TR>
		<TD align=right>Customer Name<font color=red>*</font></TD>
		<TD align=left>
			<INPUT  name=txtCustomerName type=text style="WIDTH: 425px"  disabled maxlength=50
				value=<%if logCustomerServiceID <> 0 then  Response.Write """"&routineHtmlString(objRsCustomerService("customer_name"))&"""" else Response.Write """""" end if%> >
		    <INPUT  name=btnCustomerLookup type=button tabindex=3 value=... LANGUAGE=javascript onclick="fct_lookupCustomer('D')"></TD>
		<TD></TD>
	</TR>

    <TR>
		<TD align=right>Service Location<font color=red>*</font></TD>
		<TD align=left>
			<INPUT id=txtServLocName name=txtServLocName type=text disabled style="HEIGHT: 22px; WIDTH: 425px"
				value=<%if logCustomerServiceID <> 0 then  Response.Write """"&routineHtmlString(objRsCustomerService("service_location_name"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();">
			<INPUT id=btnServiceLocationLookup name=btnServiceLocationLookup tabindex=4 type=button value=... LANGUAGE=javascript onclick="return btnServiceLocationLookup_onclick('Z')"></TD>
	</TR>
	<TR>
		<TD align=right valign=top>Location Address</TD>
		<TD align=left rowspan=6 valign=top>
			<TEXTAREA  align=left rows=6 cols=25 style="WIDTH: 425px" id=txtServLocAddress name=txtServLocAddress disabled onchange ="fct_onChange();"><% if logCustomerServiceID <> 0 then  Response.Write strServLocAddress else Response.Write null end if%><% 'if logCustomerServiceID <> 0 then  Response.Write strServLocAddress else Response.Write null end if%></TEXTAREA></TD>
		<TR></TR>
		<TR></TR>
		<TR></TR>
		<TD></TD>
		<TD align=right>Service Status<font color=red>*</font></TD>
        <TD align=left>
            <SELECT id=selServiceStatus name=selServiceStatus tabindex=14 <%if logCustomerServiceID <> 0 and Not bolCloned then  Response.Write "disabled" %> style="HEIGHT: 22px; WIDTH: 157px" onchange ="fct_onChange();">
				<OPTION></OPTION>
				<%Do while Not objRsStatusCode.EOF
					Response.write "<OPTION "
					if clng(logCustomerServiceID) <> 0 then
						if objRsCustomerService("service_status_code") <> "" then
							if objRsCustomerService("service_status_code") = objRsStatusCode(0) and Not bolCloned then
								Response.Write " SELECTED "
							elseif bolCloned and objRsStatusCode(0) = "DESIGN" then
								Response.Write " SELECTED "
							END IF
						END IF
					end if
				Response.Write 	" VALUE=" &objRsStatusCode(0)& " >" &objRsStatusCode(1)& "</OPTION>" &vbCrLf
				objRsStatusCode.MoveNext
				Loop %>
            </SELECT></TD>
		<TR>
		<TD></TD>
        <TD align=right>Customer Region</TD>
		<TD align=left ><INPUT  name=txtRegion type=text style="WIDTH: 157x"  disabled value=<%if logCustomerServiceID <> 0 then  Response.Write """"&routineHtmlString(objRsCustomerService("noc_region_desc"))&"""" else Response.Write """""" end if%> ></TD>
		</TR>
	</TR>

	<TR>
	</TR>
	<TR>
	</TR>

	<TR>
	    <TD align=right>Service Type<font color=red>*</font>
		<TD align=left>
			<INPUT id=txtServiceType name=txtServiceType type=text style="HEIGHT: 22px; WIDTH: 425px" disabled
				value=<%if logCustomerServiceID <> 0 then  Response.Write """"&routineHtmlString(strSType)&"""" else Response.Write """""" end if%> onchange ="fct_onChange();">
			<INPUT id=btnServiceTypeLookup name=btnServiceTypeLookup tabindex=5 type=button value=... LANGUAGE=javascript onclick="return btnServiceTypeLookup_onclick()"></TD>
		<TD align=right>Customer Service ID</TD>
        <TD align=left><INPUT id=txtCustomerServiceID name=txtCustomerServiceID style="HEIGHT: 22px; WIDTH: 157px"  disabled onchange ="fct_onChange();"
			value=<%if logCustomerServiceID <> 0 and not bolCloned then  Response.Write """"&objRsCustomerService("customer_service_id")&"""" else Response.Write """""" end if%>></TD>
	</TR>
 <!--
        <TD align=left> <SELECT name=selmonth2  size=1 onchange ="fct_onChange();">
 <OPTION></OPTION>
 <%

 for k = 1 to 12
  Response.Write "<option "
  if logCustomerServiceID <> 0 then
	if k = month(objRsCustomerService("date_in_service")) then
		Response.Write " selected "
	end if
  end if
  if k < 10 then
  k="0"&k
  end if
  Response.write " VALUE ="& k & ">" &ucase(monthName(k,true)) & "</OPTION>"
  next
 %>
 </SELECT>

 <SELECT  name=selday2 size=1 onchange ="fct_onChange();">
 <OPTION></OPTION>
 <%

 for k = 1 to 31
  Response.Write "<option "
  if logCustomerServiceID <> 0 then
	if k = day(objRsCustomerService("date_in_service")) then
		Response.Write " selected "
	end if
  end if
  if k < 10 then
  k="0"&k
  end if
  Response.write " VALUE ="& k & ">" &k & "</OPTION>"
  next
 %>
 </SELECT>
 <SELECT  name=selyear2 size=1 onchange ="fct_onChange();">
 <OPTION></OPTION>
 <%
 baseYear = 1994
 for i = 0 to 30
   Response.Write "<option "
   if logCustomerServiceID <> 0 then
		if (baseYear+i) = year(objRsCustomerService("date_in_service")) then
			Response.Write " selected "
		end if
  end if
  Response.write " VALUE ="& baseYear+i & ">" &baseYear+i & "</OPTION>"
  next
 %>
 </SELECT>
 <INPUT type="button" value="..." id=btnCalendar name=btnCalendar LANGUAGE=javascript onclick="return btnCalendar_onclick(2)"> </TD>
 -->
	<TR>
	    <TD align=right>SLA<font color=red>*</font></TD>
        <TD align=left>
			<INPUT id=txtServiceLevelAgreement name=txtServiceLevelAgreement disabled style="HEIGHT: 22px; WIDTH: 425px"
				value=<%if logCustomerServiceID <> 0 then  Response.Write """"&routineHtmlString(objRsCustomerService("service_level_agreement_desc"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();">
			<INPUT id=btnSLALookup name=btnSLALookup tabindex=6 type=button value=... LANGUAGE=javascript onclick="return btnSLALookup_onclick()"></TD>

		<TD align=right>Order Number</TD>
        <TD align=left><INPUT id=txtOrderNumber name=txtOrderNumber tabindex=15 style="HEIGHT: 22px; WIDTH: 157px"
        value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("project_code")&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
    </TR>
 <!--
        <TD align=left> <SELECT name=selmonth3  onchange ="fct_onChange();">
 <OPTION></OPTION>
 <%

 for k = 1 to 12
   Response.Write "<option "
  IF logCustomerServiceID <> 0 then
	if k = month(objRsCustomerService("date_terminated")) then
		Response.Write " selected "
	end if
  end if
  if k < 10 then
  k="0"&k
  end if
  Response.write " VALUE ="& k & ">" &ucase(monthName(k,true)) & "</OPTION>"
  next
 %>
 </SELECT>

 <SELECT  name=selday3  onchange ="fct_onChange();">
 <OPTION></OPTION>
 <%

 for k = 1 to 31
   Response.Write "<option "
  if logCustomerServiceID <> 0 THEN
	if k = day(objRsCustomerService("date_terminated")) then
		Response.Write " selected "
	END IF
  end if
  if k < 10 then
  k="0"&k
  end if
  Response.write " VALUE ="& k & ">" &k & "</OPTION>"
  next
 %>
 </SELECT>
 <SELECT  name=selyear3  onchange ="fct_onChange();">
 <OPTION></OPTION>
 <%

 baseYear = 1994
 for i = 0 to 30
   Response.Write "<option "
   if logCustomerServiceID <> 0 then
		if (baseYear+i) = year(objRsCustomerService("date_terminated")) then
			Response.Write " selected "
		end if
	end if
  Response.write " VALUE ="& baseYear+i & ">" &baseYear+i & "</OPTION>"
  next
 %>
 </SELECT>
 <INPUT type="button" value="..." id=btnCalendar name=btnCalendar LANGUAGE=javascript onclick="return btnCalendar_onclick(3)">
 -->

    <TR>
		<TD align=right>Support Group<font color=red>*</font></TD>
		<TD align=left>
			<SELECT id=selSupportGroup name=selSupportGroup tabindex=7 style="HEIGHT: 22px; WIDTH: 425px" onchange ="fct_onChange();">
				<OPTION></OPTION>
				<%Do while Not objRsSupportGroup.EOF
				Response.write "<OPTION "
					if logCustomerServiceID <> 0 then
						if objRsCustomerService("remedy_support_group_id") <> "" then
							if Cint(objRsCustomerService("remedy_support_group_id")) = Cint(objRsSupportGroup(0)) then
								Response.Write " SELECTED "
							END IF
						END IF
					end if
				Response.Write 	" VALUE=" &objRsSupportGroup(0)& ">" &objRsSupportGroup(1)& "</OPTION>" &vbCrLf
				objRsSupportGroup.MoveNext
				Loop	%>
			</SELECT></TD>
        <TD align=right>Date In Service</TD>
		<TD align=left><INPUT id=txtDateInService name=txtDateInService disabled style="HEIGHT: 22px; WIDTH: 157px"
        value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("date_in_service")&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
    </TR>

	<TR>
		<TD align=right>Design Specialist</TD>
        <TD align=left><INPUT id=txtContactName1 name=txtContactName1 style=" WIDTH: 400px" disabled value =
				 <%if logCustomerServiceID <> 0 then  Response.Write """"&routineHtmlString(objRsCustomerService("design_contact_name"))&"""" else Response.Write """""" end if%> onchange ="fct_onChange();">
        <INPUT name=btnDesignSpecialistLookup tabindex=8 type=button value=... LANGUAGE=javascript onclick="DesignSpecialistContactlookup();">
        <INPUT name=btnDesignSpecialistClear tabindex=9 type=button value="X" LANGUAGE=javascript onclick="btnDesignSpecialistClear_onclick();" > </TD>

        <TD align=right>Date Start Billing</TD>
		<TD align=left><SELECT name=selmonth size=1 onchange ="fct_onChange();" tabindex=16>
		<OPTION></OPTION>
		<% dim k
			for k = 1 to 12
					Response.Write "<option "
					if logCustomerServiceID <> 0 then
						if k = month(objRsCustomerService("date_to_start_billing")) then
							Response.Write " selected "
						end if
					end if
				if k < 10 then
					k="0"&k
				end if
				Response.write " VALUE ="& k & ">" &ucase(monthName(k,true)) & "</OPTION>"
			next
		%>
			</SELECT>
			<SELECT  name=selday size=1 onchange ="fct_onChange();" tabindex=17 >
			<OPTION></OPTION>
		<%
			 for k = 1 to 31
				Response.Write "<option "
				if logCustomerServiceID <> 0 then
					if k = day(objRsCustomerService("date_to_start_billing")) then
						Response.Write " selected "
					end if
				end if
				if k < 10 then
					k="0"&k
				end if
				Response.write " VALUE ="& k & ">" &k & "</OPTION>"
			 next
		%>
		</SELECT>
		<SELECT  name=selyear size=1 onchange ="fct_onChange();" tabindex=18 >
		<OPTION></OPTION>
		<%
			dim i,baseYear
			baseYear = 1994
			for i = 0 to 30
				Response.Write "<option "
				if logCustomerServiceID <> 0 then
					if (baseYear+i) = year(objRsCustomerService("date_to_start_billing")) then
						Response.Write " selected "
					end if
				end if
				Response.write " VALUE ="& baseYear+i & ">" &baseYear+i & "</OPTION>"
			next
		 %>
		</SELECT>
		<INPUT type="button" value="..." id=btnCalendar name=btnCalendar tabindex=19 LANGUAGE=javascript onclick="return btnCalendar_onclick(1)"></TD>
 </TR>

 <TR>
	<TD align=right>Implementation Manager</TD>
    <TD align=left><INPUT id=txtContactName2 name=txtContactName2 style=" WIDTH: 400px" disabled
		value=<%if logCustomerServiceID <> 0 then  Response.Write """"&routineHtmlString(objRsCustomerService("implementation_contact_name"))& """" else Response.Write """""" end if%> onchange ="fct_onChange();">
		<INPUT name=btnImplemenationLookup tabindex=11 type=button value=... LANGUAGE=javascript onclick="ImplementationConactlookup();">
		<INPUT name=btnImplemenationClear tabindex=12 type=button value="X" LANGUAGE=javascript onclick="btnImplemenationClear_onclick();" > </TD>
    <TD align=right>Date Terminated</TD>
    <TD align=left><INPUT id=txtDateTerminated name=txtDateTerminated tabindex=20 disabled style="HEIGHT: 22px; WIDTH: 157px"
        value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("date_terminated")&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
 </TR>
 <TR>
	<TD align=right valign=top>Comment</TD>
	<TD align=left rowspan=3 colspan=1 valign=top> <TEXTAREA rows=6 tabindex=13 style="WIDTH: 425px" id=txtComment name=txtComment maxlength=2000 align=left onchange ="fct_onChange();"><%if logCustomerServiceID <> 0 then  Response.Write ""&routineHtmlString(objRsCustomerService("comments"))&"" else Response.Write null end if%></TEXTAREA></TD>

		<TD align=right>Date Order Received </TD>
		<TD align=left><SELECT name=selmonth4  size=1 onchange ="fct_onChange();" tabindex=21>
		<OPTION></OPTION>
		<%
		for k = 1 to 12
			 Response.Write "<option "
			 if logCustomerServiceID <> 0 then
				if k = month(objRsCustomerService("date_workorder_received")) then
					Response.Write " selected "
				end if
			 end if
			 if k < 10 then
			 k="0"&k
			 end if
			 Response.write " VALUE ="& k & ">" &ucase(monthName(k,true)) & "</OPTION>"
		next
		%>
		</SELECT>

		<SELECT  name=selday4 tabindex=22 size=1 onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%

		for k = 1 to 31
			 Response.Write "<option "
			 if logCustomerServiceID <> 0 then
				if k = day(objRsCustomerService("date_workorder_received")) then
					Response.Write " selected "
				end if
			 end if
			 if k < 10 then
			 k="0"&k
			 end if
			 Response.write " VALUE ="& k & ">" &k & "</OPTION>"
		next
		%>
		</SELECT>

		<SELECT  name=selyear4 tabindex=23 size=1 onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%
		baseYear = 1994
		for i = 0 to 30
			  Response.Write "<option "
			  if logCustomerServiceID <> 0 then
					if (baseYear+i) = year(objRsCustomerService("date_workorder_received")) then
						Response.Write " selected "
					end if
			 end if
			 Response.write " VALUE ="& baseYear+i & ">" &baseYear+i & "</OPTION>"
		next
		%>
		</SELECT>
		<INPUT type="button" value="..." id=btnCalendar name=btnCalendar tabindex=24 LANGUAGE=javascript onclick="return btnCalendar_onclick(4)"></TD>

	<TR>
		<TD></TD>
		<TD align=right>Scheduled Completion Date</TD>
		<TD align=left><SELECT name=selmonth5  tabindex=25 size=1 onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%
		for k = 1 to 12
			 Response.Write "<option "
			 if logCustomerServiceID <> 0 then
				if k = month(objRsCustomerService("date_proposed_installed")) then
					Response.Write " selected "
				end if
			 end if
			 if k < 10 then
			 k="0"&k
			 end if
			 Response.write " VALUE ="& k & ">" &ucase(monthName(k,true)) & "</OPTION>"
		next
		%>
		</SELECT>
		<SELECT  name=selday5 size=1 tabindex=26 onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%
		for k = 1 to 31
			 Response.Write "<option "
			 if logCustomerServiceID <> 0 then
				if k = day(objRsCustomerService("date_proposed_installed")) then
					Response.Write " selected "
				end if
			 end if
			 if k < 10 then
			 k="0"&k
			 end if
			 Response.write " VALUE ="& k & ">" &k & "</OPTION>"
		 next
		%>
		</SELECT>
		<SELECT  name=selyear5 size=1 tabindex=27 onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%
		baseYear = 1994
		for i = 0 to 30
			  Response.Write "<option "
			  if logCustomerServiceID <> 0 then
					if (baseYear+i) = year(objRsCustomerService("date_proposed_installed")) then
						Response.Write " selected "
					end if
			 end if
			 Response.write " VALUE ="& baseYear+i & ">" &baseYear+i & "</OPTION>"
		next
		%>
		</SELECT>
		<INPUT type="button" value="..." id=btnCalendar name=btnCalendar tabindex=28 LANGUAGE=javascript onclick="return btnCalendar_onclick(5)"></TD>
	</TR>

	<TR>
		<TD></TD>
        <TD align=right>Date Configured</TD>
        <TD align=left><SELECT name=selmonth6  tabindex=29 size=1 onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%
		for k = 1 to 12
			 Response.Write "<option "
			 if logCustomerServiceID <> 0 then
				if k = month(objRsCustomerService("date_configured")) then
					Response.Write " selected "
				end if
			 end if
			 if k < 10 then
			 k="0"&k
			 end if
			 Response.write " VALUE ="& k & ">" &ucase(monthName(k,true)) & "</OPTION>"
		next
		%>
		</SELECT>
 		<SELECT  name=selday6 size=1 tabindex=30 onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%
 		for k = 1 to 31
			 Response.Write "<option "
			 if logCustomerServiceID <> 0 then
				if k = day(objRsCustomerService("date_configured")) then
					Response.Write " selected "
				end if
			 end if
			 if k < 10 then
			 k="0"&k
			 end if
			 Response.write " VALUE ="& k & ">" &k & "</OPTION>"
		 next
		%>
		</SELECT>
		<SELECT  name=selyear6 size=1 tabindex=31 onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%
		baseYear = 1994
		for i = 0 to 30
			  Response.Write "<option "
			  if logCustomerServiceID <> 0 then
					if (baseYear+i) = year(objRsCustomerService("date_configured")) then
						Response.Write " selected "
					end if
			 end if
			 Response.write " VALUE ="& baseYear+i & ">" &baseYear+i & "</OPTION>"
		next
		%>
		</SELECT>
		<INPUT type="button" value="..." id=btnCalendar name=btnCalendar tabindex=32 LANGUAGE=javascript onclick="return btnCalendar_onclick(6)"></TD>
	</TR>

	<TR>
     	 <TD width=15% align=right nowrap>Repair Priority</TD>
         <TD width=20% ><SELECT id=selRepairPriority name=selRepairPriority  tabindex=5 style="HEIGHT: 22px; WIDTH: 160px" onchange ="fct_onChange();">
		<%
			while not rsLYNXrp.EOF
				Response.Write "<OPTION"
				if logCustomerServiceID <> 0 then if CLng(objRsCustomerService ("LYNX_DEF_SEV_LCODE")) = CLng(rsLYNXrp("LYNX_DEF_SEV_LCODE")) then Response.write " selected"
					Response.Write " VALUE="& rsLYNXrp("LYNX_DEF_SEV_LCODE") &">" & routineHtmlString(rsLYNXrp("LYNX_DEF_SEV_DESC")) & "</OPTION>" &vbCrLf
				rsLYNXrp.MoveNext
			wend
			rsLYNXrp.Close
		%>
		<SELECT></TD>

		<TD align=right valign=top >Date Installed</TD>
        <TD align=left valign=top ><SELECT name=selmonth7 tabindex=33 size=1 onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%
 		for k = 1 to 12
			 Response.Write "<option "
			 if logCustomerServiceID <> 0 then
				if k = month(objRsCustomerService("date_installed")) then
					Response.Write " selected "
				end if
			 end if
			 if k < 10 then
			 k="0"&k
			 end if
			 Response.write " VALUE ="& k & ">" &ucase(monthName(k,true)) & "</OPTION>"
		next
		%>
		</SELECT>
 		<SELECT  name=selday7 size=1 tabindex=34 onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%
		for k = 1 to 31
			 Response.Write "<option "
			 if logCustomerServiceID <> 0 then
				if k = day(objRsCustomerService("date_installed")) then
					Response.Write " selected "
				end if
			 end if
			 if k < 10 then
			 k="0"&k
			 end if
			 Response.write " VALUE ="& k & ">" &k & "</OPTION>"
		next
		%>
		</SELECT>
		<SELECT  name=selyear7 size=1 tabindex=35 onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%
		baseYear = 1994
		for i = 0 to 30
			  Response.Write "<option "
			  if logCustomerServiceID <> 0 then
					if (baseYear+i) = year(objRsCustomerService("date_installed")) then
						Response.Write " selected "
					end if
			 end if
			 Response.write " VALUE ="& baseYear+i & ">" &baseYear+i & "</OPTION>"
		 next
		%>
		</SELECT>
		<INPUT type="button" value="..." id=btnCalendar tabindex=36 name=btnCalendar LANGUAGE=javascript onclick="return btnCalendar_onclick(7)"></TD>
	</TR>
	<TR>
     	 <TD width=15% align=right nowrap>NetCracker Weblink</TD>
         <TD width=20% ><INPUT id=btnNetcrackerweblink name=btnNetcrackerweblink tabindex=2 style="HEIGHT: 22px; WIDTH: 180px" type=button value=NetCracker LANGUAGE=javascript onclick="return btnNetcrackerweblink_onclick()"></TD>
	</TR>

</TR>


<TR>
<%if lDisplayASP = "Y" then%>
	<TD align=right>No. of Seats</TD>
    <TD align=left><INPUT id=txtNoOfSeats name=txtNoOfSeats tabindex=37 style="HEIGHT: 22px; WIDTH: 155px"
        value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("no_of_seats")&"""" else Response.Write """""" end if%> onchange ="fct_onChange();"></TD>
<%else %>
	<td></td>
	<td></td>
<%end if%>
	<TD align=right valign=top >SOCN Date</TD>
		<TD align=left><SELECT  <%if socnWrite = "N" then  Response.write("disabled=""disabled""") end if%> name=selmonth8 tabindex=25 size=1 onchange ="fct_onChange();">
		<option></option>
		<%
		for k = 1 to 12
			 Response.Write "<option "
			 if logCustomerServiceID <> 0 then
				if k = month(objRsCustomerService("date_socn")) then
					Response.Write " selected "
				end if
			 end if
			 if k < 10 then
			 k="0"&k
			 end if
			 Response.write " VALUE ="& k & ">" &ucase(monthName(k,true)) & "</OPTION>"
		next
		%>
	    </SELECT>
		<SELECT  <%if socnWrite = "N" then  Response.write("disabled=""disabled""") end if%> name=selday8 size=1 tabindex=26 onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%
		for k = 1 to 31
			 Response.Write "<option "
			 if logCustomerServiceID <> 0 then
				if k = day(objRsCustomerService("date_socn")) then
					Response.Write " selected "
				end if
			 end if
			 if k < 10 then
			 k="0"&k
			 end if
			 Response.write " VALUE ="& k & ">" &k & "</OPTION>"
		 next
		%>
		</SELECT>
		<SELECT <%if socnWrite = "N" then  Response.write("disabled=""disabled""") end if%> name=selyear8 size=1 tabindex=27 onchange ="fct_onChange();">
		<OPTION></OPTION>
		<%
		baseYear = 1994
		for i = 0 to 30
			  Response.Write "<option "
			  if logCustomerServiceID <> 0 then
					if (baseYear+i) = year(objRsCustomerService("date_socn")) then
						Response.Write " selected "
					end if
			 end if
			 Response.write " VALUE ="& baseYear+i & ">" &baseYear+i & "</OPTION>"
		next
		%>
		</SELECT>
		<%if socnWrite = "Y" then %>
		<INPUT type="button" value="..." id=btnCalendar name=btnCalendar tabindex=28 LANGUAGE=javascript onclick="return btnCalendar_onclick(8)"></TD>
		<%end if%>
</TR>


	</table>

	<table>
	<thead><TR><TD align=left colspan=4>Customer Service Contacts</TD></TR></thead>
	<tbody>
		<TR>
			<TD colspan=4><iframe id=aifr2 width=100% height=100 src="" scrolling=yes marginheight=1 marginwidth=1></iframe>
			<br>
			<!-- The following buttons are disabled if this is a new record -->
			<input type="button" tabindex=37 style="width: 2cm" value="Delete"  <%if logCustomerServiceID = 0 or bolCloned then  Response.Write "DISABLED" end if%>   id="btn_iframeDelete"  name="btn_iframeDelete"  onClick="btn_ifrmDelete();">&nbsp;&nbsp;
			<input type="button" tabindex=38 style="width: 2cm" value="Refresh" <%if logCustomerServiceID = 0 or bolCloned then  Response.Write "DISABLED" end if%>   id="btn_iframeRefresh" name="btn_iframeRefresh" onClick="iFrame_display();">&nbsp;&nbsp;
			<input type="button" tabindex=39 style="width: 2cm" value="New"     <%if logCustomerServiceID = 0 or bolCloned then  Response.Write "DISABLED" end if%>   id="btn_iframeAdd"     name="btn_iframeAdd"     onClick="btn_ifrmAdd();">&nbsp;&nbsp;
			<input type="button" tabindex=40 style="width: 2cm" value="Update"  <%if logCustomerServiceID = 0 or bolCloned then  Response.Write "DISABLED" end if%>   id="btn_iframeupdate" name="btn_iframeupdate"   onClick="btn_ifrmUpdate();">
		</TR>
	</tbody>
	</TABLE>
	<TABLE>
	<thead><tr ><td colspan=6 align=left>Non Ecops Tracking</TD></TR></thead>
	<tbody>
    <TR>
        <TD align=right>Date Facility Ordered</TD>
        <TD align=left><INPUT id=txtDateFacilityOrdered name=txtDateFacilityOrdered  style="HEIGHT: 22px; WIDTH: 135px" disabled
			value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("date_facility_ordered")&"""" else Response.Write """""" end if%> ></TD>
        <TD align=right>Date Facility Due</TD>
        <TD align=left><INPUT id=txtDateFacilityDue name=txtDateFacilityDue  style="HEIGHT: 22px; WIDTH: 135px" disabled
			value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("date_facility_due")&"""" else Response.Write """""" end if%> >   </TD>
        <TD align=right>Missed Inst. Date Cause </TD>
        <TD rowspan=3 valign=top><TEXTAREA disabled id=textarea1 name=textarea1 >
			<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("missed_installation_cause_desc")&"""" else Response.Write NULL end if%></TEXTAREA> </TD>
    <TR>
		<TD align=right>Date Facility Confirmed</TD>
        <TD align=left><INPUT id=txtFacilityConfirmed name=txtFacilityConfirmed  style="HEIGHT: 22px; WIDTH: 135px" disabled
			value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("date_facility_confirmed")&"""" else Response.Write """""" end if%>></TD>
        <TD align=right>Date Facility Ready</TD>
        <TD align=left><INPUT id=txtFacilityReady name=txtFacilityReady  style="HEIGHT: 22px; WIDTH: 135px" disabled
			value=<%if logCustomerServiceID <> 0 then  Response.Write """"&objRsCustomerService("date_facility_ready")&"""" else Response.Write """""" end if%>></TD>
        <TD></TD></TR>
	<TR>
		<TD>&nbsp; </TD>
		<TD>&nbsp; </TD>
		<TD>&nbsp; </TD>
		<TD>&nbsp; </TD>
	</TR>
	</tbody>
	</TABLE>
<!-- New Frame begins -->
	<TABLE>
	<thead>
	<TR>
		<TD width="50%" align="left" colspan=2>Service Type Attributes and Values</TD>
		<TD width="70%" align="left" valign="top" colspan=2>Working Service Instance Attribute Values for this CSID</TD>
	</TR>
	</thead>
	<tbody>
		<TR>
			<TD colspan=2>
			<iframe id=aifr1a width=100% height=75 scrolling=yes marginheight=1 marginwidth=1></iframe></td>
			<TD width="70%" align="left" valign="top" colspan=2>
				<iframe id=waifr3 width=100% height=75 scrolling=yes marginheight=1 marginwidth=1></iframe>
			<br>
			</td>
		</TR>
	</tbody>


	<thead>
		<TR>
		<TD bgcolor=#FFFFCC align="left" colspan=2></TD>
		<TD width="70%" align="left" valign="top" colspan=2>Service Instance Attribute Values Requested for this CSID in this Order</TD>
	</TR>
	</thead>
	<tbody>
		<TR>
			<TD colspan=2></td>
			<TD width="50%" align="left" valign="top" colspan=2>
				<iframe id=aifr1b width=100% height=75 scrolling=yes marginheight=1 marginwidth=1></iframe>
			<br>
			</td>
		</TR>
	</tbody>

	<table>
	<thead><TR><TD align=left colspan=6>VPN Info</TD></TR></thead>
	<tbody>
		<TR>
			<TD colspan=6><iframe id=aifrvpnws1 width=100% height=100 src="" scrolling=yes marginheight=1 marginwidth=1></iframe>
			<br>
			<!-- The following buttons are disabled if this is a new record -->
		</TR>
	</tbody>
	</TABLE>

<!--New Frame ends -->
	<table>
	<tfoot>
    <TR>
		<TD align=right colspan=6>
			<input name=btnReferences tabindex=41 type=button value=References  tabindex=13 style= "width: 2.2cm" LANGUAGE=javascript onclick="return btnReferences_onclick()">&nbsp;&nbsp;
			<INPUT id=btnDelete name=btnDelete tabindex=42 type=button value=Delete style="WIDTH: 2cm" onclick="fct_onDelete();">&nbsp;&nbsp;
			<INPUT id=btnReset name=btnReset tabindex=43 type=button value=Reset style="WIDTH: 2cm" onclick="btnReset_onclick();">&nbsp;&nbsp;
			<INPUT id=btnAddNew name=btnAddNew tabindex=44 type=button value=New  style="WIDTH: 2cm"  onclick="fct_onNew();" >&nbsp;&nbsp;
			<INPUT id=btnClone name=btnClone tabindex=45 type=button value=Clone  style="WIDTH: 2cm" onclick="fct_onClone();">&nbsp;&nbsp;
			<INPUT id=btnSave name=btnSave tabindex=46 type=button value=Save style="WIDTH: 2cm" onclick="return form_onsubmit();" >&nbsp;&nbsp;
		</TD></TR>
	</tfoot>
	</table>
	<FIELDSET>
<% ' need to set the Customer service ID to nothing so the remaining fields will not fill
	if bolCloned then
		logCustomerServiceID = 0
	end if
%>
	<Div SIZE=8pt ALIGN=RIGHT>
		Record Status Indicator
		<INPUT align = left name=txtRecordStatusInd type=text style="HEIGHT: 20px; WIDTH: 18px" disabled
			value=<%if logCustomerServiceID <> 0 then  Response.Write """"&routineHtmlString(objRsCustomerService("record_status_ind"))&"""" else Response.Write """""" end if%> >&nbsp;&nbsp;&nbsp;
		Create Date&nbsp;
		<INPUT align = center name=txtCreateDate type=text style="HEIGHT: 20px; WIDTH: 150px" disabled
			value=<%if logCustomerServiceID <> 0 then  Response.Write """"&routineHtmlString(objRsCustomerService("create_date"))&"""" else Response.Write """""" end if%> >&nbsp;
		Created By
		<INPUT align = right name=txtCreateRealUserid type=text style="HEIGHT: 20px; WIDTH: 200px" disabled
			value=<%if logCustomerServiceID <> 0 then  Response.Write """"&routineHtmlString(objRsCustomerService("create_real_userid"))&"""" else Response.Write """""" end if%> ><BR>
		Update Date&nbsp;
		<INPUT align= center name=txtUpdateDate type=text style="HEIGHT: 20px; WIDTH: 150px" disabled
			value=<%if logCustomerServiceID <> 0 then  Response.Write """"&routineHtmlString(objRsCustomerService("update_date"))&"""" else Response.Write """""" end if%> >
		Updated By
		<INPUT align=right name=txtUpdateRealUserid type=text style="HEIGHT: 20px; WIDTH: 200px" disabled
			value=<%if logCustomerServiceID <> 0 then  Response.Write """"&routineHtmlString(objRsCustomerService("update_real_userid"))&"""" else Response.Write """""" end if%> >
	</DIV>
	</FIELDSET>


<%
 if logCustomerServiceID <> 0 then
	 'clean up ADO objects
	set objRsRegion = nothing
	set objRsCustomerService = nothing
	objConn.close
	set objConn = nothing

END IF %>

</FORM>
</BODY>
</HTML>
