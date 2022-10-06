<%@  language="VBScript" %>
<% option explicit %>
<% Response.Buffer = True %>

<!--#include file="SmaConstants.inc"-->
<!--#include file="SMA_Env.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--
*****************************************************************************************
* Page name:	CustDetail.asp
* Purpose:		To display the detailed information about a customer.
*				Customer chosen via CustList.asp
*
* Created by:	Nancy Mooney	08/03/2000
*
*****************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       19-Dec-01	DTy			Display Customer ID in the Audit section.
       19-Feb-02	DTy			Increase email address size from 50 t0 60.
       05-Mar-02	DTy			Change Customer Status drop down list sort sequence from
								Customer Status Description to Sort_Order field
***************************************************************************************************
-->
<%
'*************SECURITY********************************************************************
dim intAccessLevel, intNameAccessLevel, intAliasAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_Customer))
if (intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to customer. Please contact your system administrator"
end if
intNameAccessLevel = CInt(CheckLogon(strConst_CustomerName))
intAliasAccessLevel = CInt(CheckLogon(strConst_CustomerNameAlias))
'********************************************************************************************
dim strWinMessage, strRealUserID, strCustomerID, datUpdateDateTime, strWinLocation
strWinMessage = ""
strRealUserID =  Session("username")

strCustomerID = Request("CustomerID")
'Response.Write ("customer id:" & strCustomerID & "<BR><BR>")
datUpdateDateTime = Request("UpdateDateTime")
strWinLocation = "CustDetail.asp?CustID="&Request("hdnCustID")


dim hstring
hstring = Session("userRoles")
if len(hstring) = 0 then
	hstring ="hidden"
else
    hstring = ""
end if
 
'form action - txtFrmAction receives value before submitting to self
select case Request("txtFrmAction")
	case "SAVE"
		'create command object for stored procedures
		dim cmdObj, strErrMessage
		set cmdObj = server.CreateObject("ADODB.Command")
		'set cmdObj.ActiveConnection = objConn     LC: move it down
		cmdObj.CommandType = adCmdStoredProc

		'parse together phone numbers

		'customer phone
		dim strCPhone,strCPArea,strCPMid,strCPEnd
		strCPArea = Request("txtPArea")
		strCPMid = Request("txtPMid")
		strCPEnd = Request("txtPEnd")
		if strCPArea <> "" then
			strCPhone = strCPhone & strCPArea
		end if
		if strCPMid <> "" then
			strCPhone = strCPhone & strCPMid
		end if
		if  strCPEnd <> "" then
			strCPhone = strCPhone & strCPEnd
		end if
		if len(strCPhone)= 0 then strCPhone = null

		'fax
		dim strFxPhone,strFxArea,strFxMid,strFxEnd
		strFxArea = Request("txtFArea")
		strFxMid = Request("txtFMid")
		strFxEnd = Request("txtFEnd")
		if strFxArea <> "" then
			strFxPhone = strFxPhone & strFxArea
		end if
		if strFxMid <> "" then
			strFxPhone = strFxPhone & strFxMid
		end if
		if  strFxEnd <> "" then
			strFxPhone = strFxPhone & strFxEnd
		end if
		if len(strFxPhone)= 0 then strFxPhone = null

		'check non-required fields
		dim txtCustomerShortName, selIndustry, txtEmail, txtWebSite, txtComments
		if len(Request("txtCustomerShortName")) > 0 then
			txtCustomerShortName = Request("txtCustomerShortName")
		else
			txtCustomerShortName = Null
		end if

		if len(Request("selCustIndustry")) > 0 then
			selIndustry = Request("selCustIndustry")
		else
			selIndustry = null
		end if

		if len(Request("txtEmail")) > 0 then
			txtEmail = Request("txtEmail")
		else
			txtEmail = null
		end if

		if len(Request("txtWebSite")) > 0 then
			txtWebSite = Request("txtWebSite")
		else
			txtWebSite = null
		end if

		if len(Request("txtComments")) > 0 then
			txtComments = Request("txtComments")
		else
			txtComments = null
		end if




		if isNumeric(Request("hdnCustomerID"))then 'update existing record
			'security check
			if ((intAccessLevel and intConst_Access_Update) <> intConst_Access_Update) then
				DisplayError "BACK", "", 0, "UPDATE DENIED", "You don't have access to update customer. Please contact your system administrator"
			end if

			set cmdObj.ActiveConnection = objConn
			cmdObj.CommandText = "sma_sp_userid.spk_sma_cust_inter.sp_customer_update"
			'create parameters
			'required fields
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_userid", adVarChar,adParamInput, 20, strRealUserID)
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_customer_id",adNumeric, adParamInput, ,CLng(Request("hdnCustomerId")))
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_customer_type",adChar, adParamInput,3,Request("selCustType"))
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_customer_status",adVarChar, adParamInput,8,Request("selCustStatus"))
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_customer_name",adVarChar, adParamInput,50,Request("txtCustName"))
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_language_pref",adChar, adParamInput,2,Request("selCustLangPref"))
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_region",adVarChar, adParamInput,8,Request("selCustRegion"))
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_last_update_dt",adDBTimeStamp, adParamInput,,CDate(Request("hdnUpdateDateTime")))
			'optional fields
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_customer_short",adVarChar, adParamInput,15,txtCustomerShortName)
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_industry",adNumeric, adParamInput, ,selIndustry)
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_phone",adVarChar, adParamInput,24,strCPhone)
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_fax",adVarChar, adParamInput,24,strFxPhone)
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_email",adVarChar, adParamInput,50,txtEmail)
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_web_site",adVarChar, adParamInput,50,txtWebSite)
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_comments",adVarChar, adParamInput,2000,txtComments)

			strErrMessage = "CANNOT UPDATE CUSTOMER"

		else 'create new customer
			'security
			if ((intAccessLevel and intConst_Access_Create) <> intConst_Access_Create) then
				DisplayError "BACK", "", 0, "UPDATE DENIED", "You don't have access to create a customer. Please contact your system administrator"
			end if


			'need reset connection string if user has lcd roles
			' if the lcd role is required, get this role's oracle id
			dim strLCDString
			strLCDString = Request("selLCDName")
			strLCDString ="APP_SMA_" + ucase(strLCDString)


			Dim newCustConnectString
		 

			if strcomp(ucase(strLCDString),"APP_SMA_NON-LCD")=0 or strcomp(ucase(strLCDString),"APP_SMA_")=0 then
			      newCustConnectString = Decrypt(getConnString("strConstConnectString"))

			else
				 newCustConnectString =  Decrypt(getConnString(ucase(strLCDString)))



			end if

			set objConn = Server.CreateObject("ADODB.Connection")
			objConn.ConnectionString = newCustConnectString
			objConn.open


			'unexpected error, possible a database connection error
			if err then
				DisplayError "BACK", "", err.number, "UNEXPECTED ERROR - Possible database connection error", err.description
			end if



'			response.write newCustConnectString
'		response.end












			set cmdObj.ActiveConnection = objConn
			cmdObj.CommandText = "sma_sp_userid.spk_sma_cust_inter.sp_customer_insert"
			'create parameters
			'required fields
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_userid", adVarChar,adParamInput, 20, strRealUserID)
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_customer_id",adNumeric, adParamOutput)
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_customer_type",adChar, adParamInput,3,Request("selCustType"))
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_customer_status",adVarChar, adParamInput,8,Request("selCustStatus"))
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_customer_name",adVarChar, adParamInput,50,Request("txtCustName"))
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_language_pref",adChar, adParamInput,2,Request("selCustLangPref"))
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_region",adVarChar, adParamInput,8,Request("selCustRegion"))
			'optional fields
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_customer_short",adVarChar, adParamInput,15,txtCustomerShortName)
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_industry",adNumeric, adParamInput,9,selIndustry)
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_phone",adVarChar, adParamInput,24,strCPhone)
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_fax",adVarChar, adParamInput,24,strFxPhone)
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_email",adVarChar, adParamInput,50,txtEmail)
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_web_site",adVarChar, adParamInput,50,txtWebSite)
			cmdObj.Parameters.Append cmdObj.CreateParameter("p_comments",adVarChar, adParamInput,2000,txtComments)

			strErrMessage = "CANNOT CREATE CUSTOMER"

		end if

		'parameter check - development
			'cmdObj.Parameters.Refresh
			'dim objparm
			'for each objparm in cmdObj.Parameters
			'	Response.Write "<b>" & objparm.name & "</b>"
			'	Response.Write " and value: " & objparm.value & ""
			'	Response.Write " and datatype: " & objparm.Type & "<br>"
			'next

			'Response.Write "<b> count = " & cmdObj.Parameters.count & "<br>"
			'dim nx
			'for nx = 0 to cmdObj.Parameters.Count-1
			'Response.Write cmdObj.Parameters.Item(nx).Name & " = " & cmdObj.Parameters.Item(nx).Value & " <br>"
			'next
			'Response.end

		'call the stored proc
		on error resume next
		cmdObj.Execute
		If objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, strErrMessage, objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
		strCustomerID = CStr(cmdObj.Parameters("p_customer_id").value)
		strWinMessage = "Record saved successfully."

 		'no need to reset connection string back as objConn is cleared at end of this page
 		'set objConn = Server.CreateObject("ADODB.Connection")
		'objConn.ConnectionString = Session("ConnectString")
		'objConn.open










	case "DELETE"

		if ((intAccessLevel and intConst_Access_Delete) <> intConst_Access_Delete) then
			DisplayError "BACK", strWinLocation, 0, "DELETE DENIED", "You don't have access to delete contacts. Please contact your system administrator."
		end if


		set cmdObj = server.CreateObject("ADODB.Command")
		set cmdObj.ActiveConnection = objConn
	 	cmdObj.CommandType = adCmdStoredProc
		cmdObj.CommandText = "sma_sp_userid.spk_sma_cust_inter.sp_customer_delete"
		cmdObj.Parameters.Append cmdObj.CreateParameter("p_customer_id", adNumeric, adParamInput, ,clng(strCustomerID))
		cmdObj.Parameters.Append cmdObj.CreateParameter("p_last_update", adDBTimeStamp, adParamInput,,cdate(datUpdateDateTime))
        cmdObj.Parameters.Append cmdObj.CreateParameter("P_USER_ID", adVarChar , adParamInput, 30, strRealUserID)

		'Response.Write "<b> count = & cmdObj.Parameters.count & <br>"
		'	dim nx
		'	for nx = 0 to cmdObj.Parameters.Count-1
		'		Response.Write "parm value = " & cmdObj.Parameters.Item(nx).Value & " <br>"
		'next
	'response.end

		on error resume next
		cmdObj.Execute
		if objConn.Errors.Count <> 0 then
			DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
			objConn.Errors.Clear
		end if
		strCustomerID = "DEL"
		strWinMessage = "Record Deleted Successfully."

	end select

	'response.write "strCustomerID = " + strCustomerID
	'response.end

	'If NOT a new record - build sql and retrieve customer record
	If isNumeric(strCustomerID) then
		'create SQL for populating fields
		Dim strSQL, strSelectClause, strFromClause, strWhereClause
		Dim rsCustomer, rsCustRegion, rsCustStatus, rsCustIndustry, rsCustLangPref
 		'build query
		strSelectClause = "select " &_
					"t1.customer_id, " & _
					"t1.customer_type_ind, " & _
					"t1.customer_status_lcode, " & _
					"t2.customer_status_desc, " & _
					"t1.industry_id, " & _
					"t4.industry_desc, " & _
					"t1.customer_name, " & _
					"t1.phone_number, " & _
					"t1.fax_number, " & _
					"t1.email_address, " & _
					"t1.web_site_url, " & _
					"t1.customer_short_name, " & _
					"t1.language_preference_lcode, " & _
					"t5.language_preference_desc, " & _
					"t1.noc_region_lcode, " & _
					"t3.noc_region_desc, " & _
					"t1.comments, " & _
					"t1.record_status_ind, " &_
					"to_char(t1.create_date_time,'MON-DD-YYYY HH24:MI:SS') as create_date, " & _
					"sma_sp_userid.spk_sma_library.sf_get_full_username(t1.create_real_userid) as create_real_userid, " & _
					"to_char(t1.update_date_time,'MON-DD-YYYY HH24:MI:SS') as update_date, " & _
					"sma_sp_userid.spk_sma_library.sf_get_full_username(t1.update_real_userid) as update_real_userid, " & _
					"t1.update_date_time as last_update_date_time, " & _
					"nvl(t8.building_name,'(No building name)') || chr(10) || nvl(t8.street,'(No street address)') ||chr(10)|| " & _
					"nvl(t8.municipality_name,'(No municipality)')||' '|| nvl(t8.province_state_lcode,'(No prov/state)') " & _
					"||' '|| nvl(t8.country_lcode,'(No Country)') ||chr(10)|| nvl(t8.postal_code_zip,'(No postal Code)') primary_address, " &_
					"nvl(t9.building_name,'(No building name)') ||chr(10)|| nvl(t9.street,'(No street address)') ||chr(10)|| " & _
					"nvl(t9.municipality_name,'(No municipality)') ||' '|| nvl(t9.province_state_lcode,'(No prov/state)') " & _
					"||' '|| nvl(t9.country_lcode,'(No Country)')||chr(10)||nvl(t9.postal_code_zip,'(No postal Code)') billing_address, " &_
					"nvl(t10.building_name,'(No building name)') ||chr(10)|| nvl(t10.street,'(No street address)') ||chr(10)|| " & _
					"nvl(t10.municipality_name,'(No municipality)') ||' '|| nvl(t10.province_state_lcode,'(No prov/state)') " & _
					"||' '|| nvl(t10.country_lcode,'(No Country)') ||chr(10)|| nvl(t10.postal_code_zip,'(No postal Code)') mailing_address "

		strFromClause =	" from crp.customer t1, " &_
					"crp.lcode_customer_status t2, " & _
					"crp.lcode_noc_region t3, " & _
					"crp.industry_lookup t4, " & _
					"crp.lcode_language_preference t5, " & _
					"crp.v_address_consolidated_street t8, " &_
					"crp.v_address_consolidated_street t9, " &_
					"crp.v_address_consolidated_street t10 "

		 strWhereClause = " where rownum =1 and " & _
					"t1.customer_status_lcode = t2.customer_status_lcode and " & _
					"t1.noc_region_lcode = t3.noc_region_lcode and " & _
					"t1.industry_id = t4.industry_id(+) and " & _
					"t1.language_preference_lcode = t5.language_preference_lcode and " & _
					"t1.customer_id = t8.customer_id (+) and " & _
					"t1.customer_id = t9.customer_id (+) and " & _
					"t1.customer_id = t10.customer_id (+) and " & _
					"t8.primary_address_flag(+)= 'Y' and " & _
					"t9.billing_address_flag(+)= 'Y' and " & _
					"t10.mailing_address_flag(+) = 'Y' and " & _
					"t1.customer_id = " & strCustomerID

		strSQL =  strSelectClause & strFromClause & strWhereClause

		'Response.Write strSQL	'--debugging check
		'RESPONSE.END
	objConn.Errors.Clear

		'get the customer recordset
		set rsCustomer = Server.CreateObject("ADODB.Recordset")
		rsCustomer.CursorLocation = adUseClient
		rsCustomer.Open strSQL, objConn
		if err then
			DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 1", err.Description
		end if
		if rsCustomer.EOF then
			DisplayError "BACK", "", 999, "CANNOT FIND REQUESTED OBJECT", "EOF condition occurred in rsCustomer recordset."
		end if
		set rsCustomer.ActiveConnection = nothing

		'pars phone number
		dim strPArea, strPMid, strPEnd
		strPArea = mid(rsCustomer("phone_number"),1,3)
		strPMid = mid(rsCustomer("phone_number"),4,3)
		strPEnd = mid(rsCustomer("phone_number"),7,10)

		'pars fax number
		dim strFArea, strFMid, strFEnd
		strFArea = mid(rsCustomer("fax_number"),1,3)
		strFMid = mid(rsCustomer("fax_number"),4,3)
		strFEnd = mid(rsCustomer("fax_number"),7,10)

		'get the customer alias recordset
		dim rsAlias
		strSQL= "Select customer_name_alias_id, customer_name_alias_upper from crp.customer_name_alias where customer_id = " & strCustomerID
		set rsAlias=server.CreateObject("ADODB.Recordset")
		rsAlias.CursorLocation = adUseClient
		rsAlias.Open strSQL, objConn
		if err then
			DisplayError "BACK", "", err.Number, "Cannot create recordset rsAlias.", err.Description
		end if
		set rsAlias.ActiveConnection=nothing

		'create the innerValues for the iFrame
		dim intRowCount, intColCount, strInnerValues
			intRowCount = 0
			intColCount = 2
			strInnerValues = ""
		while not rsAlias.EOF
			intRowCount = intRowCount + 1
			strInnerValues =strInnerValues & rsAlias(0) & strDelimiter & rsAlias(1) & strDelimiter
			rsAlias.MoveNext
		wend
		rsAlias.Close
		set rsAlias = nothing

		'custcare (aka Service Management Rep) list
		dim rsCustCare, cList, intPgNum, inPgCnt, bolCustCareFlag
		strSQL = "select con.contact_name, " &_
					    "cucon.contact_priority, " & _
					    "con.work_number, " &_
					    "con.work_number_ext, " &_
					    "con.cell_number, " &_
					    "con.pager_number, " &_
					    "con.fax_number, " &_
					    "con.email_address " &_
			"from crp.customer_contact cucon, " &_
				 "crp.contact con " & _
			"where cucon.customer_id = " & strCustomerID & " and " &_
				  "cucon.customer_contact_type_lcode = 'custcare' and " & _
				  "cucon.record_status_ind = 'A' and " &_
			      "cucon.contact_id = con.contact_id and " & _
			      "con.record_status_ind = 'A'"
		set rsCustCare = Server.CreateObject("ADODB.Recordset")
		rsCustCare.CursorLocation = adUseClient
		rsCustCare.Open strSQL, objConn
		if err then
			DisplayError "BACK", "", err.Number, "Cannot create recordset rsCustCare.", err.Description
		end if

		if not rsCustCare.EOF then
			clist = rsCustCare.GetRows
			bolCustCareFlag = true
		else
			bolCustCareFlag = false
		end if

		'release and kill the recordset and the connection objects
		rsCustCare.Close
		set rsCustCare = nothing


		'address list (primary,billing,mailing)
		dim rsAddress, aList, intPageNumber, intPageCount, bolAddressFlag
		strSQL = "select distinct(a.address_id), " &_
			"a.billing_address_flag as billing, " &_
			"a.primary_address_flag as primary, " &_
			"a.mailing_address_flag as mailing, " &_
			"a.street, " &_
			"a.municipality_name, " &_
			"a.province_state_lcode, " &_
			"a.country_lcode, " &_
			"nvl(a.building_name, '(No building name)' ), " &_
			"p.province_state_name, " &_
			"c2.country_desc, " &_
			"a.postal_code_zip " & _
		"from crp.customer c, " &_
			 "crp.V_ADDRESS_CONSOLIDATED_STREET a, " &_
			 "crp.customer_name_alias c1, " &_
			 "crp.lcode_country c2, " &_
			 "crp.lcode_province_state p " & _
		"where c.customer_id = " & strCustomerID & _
			" and (a.primary_address_flag = 'Y' or a.mailing_address_flag = 'Y'or a.billing_address_flag = 'Y')" & _
			" and c.customer_id = a.customer_id" &_
			" and c.customer_id = c1.customer_id" &_
			" and a.province_state_lcode = p.province_state_lcode" &_
			" and a.country_lcode = c2.country_lcode" & _
			" and p.country_lcode = c2.country_lcode" & _
			" and a.record_status_ind = 'A'"
		set rsAddress = Server.CreateObject("ADODB.Recordset")
		rsAddress.CursorLocation = adUseClient
		rsAddress.Open strSQL, objConn
		if err then
			DisplayError "BACK", "", err.Number, "Cannot create recordset rsAddress.", err.Description
		end if

		if not rsAddress.EOF then
			alist = rsAddress.GetRows
			bolAddressFlag = true
		else
			bolAddressFlag = false
		end if

		'release and kill the recordset and the connection objects
		rsAddress.Close
		set rsAddress = nothing

	end if 'check for numeric strCustomerID

	'get list items NOTE: customer type is hard-coded because there is no corresponding lcode table
	'noc region
	strSQL = "select noc_region_lcode, noc_region_desc" & _
			 " from crp.lcode_noc_region" & _
			 " where record_status_ind = 'A'" & _
			 " order by noc_region_desc"
	set rsCustRegion = Server.CreateObject("ADODB.Recordset")
	rsCustRegion.CursorLocation = adUseClient
	rsCustRegion.Open strSQL, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 2", err.Description
	end if
	if rsCustRegion.EOF then
		DisplayError "BACK", "", 999, "CANNOT CREATE OBJECT TYPE LIST", "EOF condition occurred in rsCustRegion recorset."
	end if
	'release the active connection, keep the recordset open
	set rsCustRegion.ActiveConnection = nothing

	'customer status
	strSQL = "select customer_status_lcode, customer_status_desc from crp.lcode_customer_status where record_status_ind = 'A' order by sort_order"
	set rsCustStatus = Server.CreateObject("ADODB.Recordset")
	rsCustStatus.CursorLocation = adUseClient
	rsCustStatus.Open strSQL, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 3", err.Description
	end if
	if rsCustStatus.EOF then
		DisplayError "BACK", "", 999, "CANNOT CREATE OBJECT TYPE LIST", "EOF condition occurred in rsCustStatus recorset."
	end if
	'release the active connection, keep the recordset open
	set rsCustStatus.ActiveConnection = nothing

	'industry
	strSQL = "select industry_id, industry_desc from crp.industry_lookup where record_status_ind = 'A' order by industry_desc"
	set rsCustIndustry = Server.CreateObject("ADODB.Recordset")
	rsCustIndustry.CursorLocation = adUseClient
	rsCustIndustry.Open strSQL, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 4", err.Description
	end if
	if rsCustIndustry.EOF then
		DisplayError "BACK", "", 999, "CANNOT CREATE OBJECT TYPE LIST", "EOF condition occurred in rsCustIndustry recorset."
	end if
	'release the active connection, keep the recordset open
	set rsCustIndustry.ActiveConnection = nothing

	'language preference
	strSQL = "select language_preference_lcode, language_preference_desc from crp.lcode_language_preference where record_status_ind = 'A' order by language_preference_desc"
	set rsCustLangPref = Server.CreateObject("ADODB.Recordset")
	rsCustLangPref.CursorLocation = adUseClient
	rsCustLangPref.Open strSQL, objConn
	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 5", err.Description
	end if
	if rsCustLangPref.EOF then
		DisplayError "BACK", "", 999, "CANNOT CREATE OBJECT TYPE LIST", "EOF condition occurred in rsCustLangPref recorset."
	end if
	'release the active connection, keep the recordset open
	set rsCustLangPref.ActiveConnection = nothing
%>

<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link rel="stylesheet" type="text/css" href="stylesheets/styles.css">
    <script type="text/javascript" language="javascript" src="GeneralJavaFunctions.js"></script>
    <script type="text/javascript" language="javascript" src="AccessLevels.js"></script>
    <script type="text/javascript" language="javascript">

	<!--hide script

    var bolNeedToSave = false;
    var strWinMessage = "<%=strWinMessage%>";
    var intAccessLevel = "<%=intAccessLevel%>";
    var intNameAccessLevel = "<%=intNameAccessLevel%>";
    var intAliasAccessLevel = "<%=intAliasAccessLevel%>";
    var strCustomerID = "<%=strCustomerID%>";
    var struserRoles = "<%=Session("userRoles")%>";


    //set title
    setPageTitle("SMA - Customer");

    //javascript code related to iFrame functionality-----------------------------------------

    function iFrame_display() {
        //  if ((intAccessLevel & intConst_Access_ReadOnly) == intConst_Access_ReadOnly) {
        document.getElementById("aifr").src = 'CustAlias.asp?CustomerID=<%=strCustomerID%>';
        //  }
        // else { alert('Access Denied. You do not have access to name alias. Please contact your system administrator.') }
    }

    function CustOrgMClick() {
        var NewWin;
        //document.frames("aifr").document.location.href = 'CustAlias.asp?CustomerID=<%=strCustomerID%>';

        NewWin = window.open("COM.asp?action=new&CId=" + strCustomerID, "NewWin", "width=850px,scrollbars=1,resizable=1");
        NewWin.focus();
    }

    function btn_iFrmAdd() {
        //open a blank form
        if ((intAliasAccessLevel & intConst_Access_Create) != intConst_Access_Create) {
            alert('Access Denied - can not create alias. Please contact your system administrator.');
            return (false);
        }
        if (document.frmCustDetail.hdnCustomerID.value == "") {
            alert('At this time you cannot create a name alias. You must save the customer first.');
            return (false);
        }
        var NewWin;
        var strMasterID = "<%=strCustomerID%>";
        NewWin = window.open("CustAliasDetail.asp?action=new&masterID=" + strMasterID, "NewWin", "toolbar=no,status=yes,width=700px,height=175px,left=150px,top=200,menubar=no,resize=no");
        NewWin.focus();
    }

    function btn_iFrmUpdate() {
        //open a detail form where the user can modify the alias
        if ((intAliasAccessLevel & intConst_Access_Update) != intConst_Access_Update) {
            alert('Access Denied - can not update alias. Please contact your system administrator.');
            return;
        }
        var NewWin;

        var doc;
        var iframeObject = document.getElementById('aifr'); // MUST have an ID
        if (iframeObject.contentDocument) { // DOM
            doc = iframeObject.contentDocument;
        }
        else if (iframeObject.contentWindow) { // IE win
            doc = iframeObject.contentWindow.document;
        }

        var strAliasID = doc.getElementsByName("hdnNameAliasID")[0].value; // document.frames("aifr").document.frmIFR.hdnNameAliasID.value;
        if (strAliasID == "") {
            alert("Please select an alias or click NEW to create a new alias.");
            return;
        }
        var strMasterID = "<%=strCustomerID%>";
        NewWin = window.open("CustAliasDetail.asp?action=update&aliasID=" + strAliasID + "&masterID=" + strMasterID, "NewWin", "toolbar=no,status=yes,width=700px,height=175px,left=150px,top=200,menubar=no,resize=no");
        NewWin.focus();
    }

    function btn_iFrmDelete() {
        //delete selected row
        if ((intAliasAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
            alert('Access Denied - can not delete alias. Please contact your system administrator.');
            return;
        }

        var doc;
        var iframeObject = document.getElementById('aifr'); // MUST have an ID
        if (iframeObject.contentDocument) { // DOM
            doc = iframeObject.contentDocument;
        }
        else if (iframeObject.contentWindow) { // IE win
            doc = iframeObject.contentWindow.document;
        }



        var strAliasID = doc.getElementsByName("hdnNameAliasID")[0].value;// document.frames("aifr").document.frmIFR.hdnNameAliasID.value;
        if (strAliasID == "") {
            alert("Please select an alias or click ADD to create a new alias.");
            return;
        }
        var strLastUpdate = doc.getElementsByName("hdnLastUpdate")[0].value;// document.frames("aifr").document.frmIFR.hdnLastUpdate.value;
        if (confirm("Are you sure you want to delete this alias?")) {
            // document.frames("aifr").document.location.href = "CustAliasDetail.asp?action=delete&back=true&aliasID=" + strAliasID + "&masterID=<%=strCustomerID%>&hdnLastUpdate=" + strLastUpdate;
            document.getElementById("aifr").src = "CustAliasDetail.asp?action=delete&back=true&aliasID=" + strAliasID + "&masterID=<%=strCustomerID%>&hdnLastUpdate=" + strLastUpdate;
        }
    }

    //-----------------end of iFrame Javascript-----------------------------------------------
    function body_onLoad(strWinStatus) {
        DisplayStatus(strWinStatus);
        iFrame_display();
        return true;
    }

    function fct_selNavigate(strPageName) {
        //***************************************************************************************************
        // Function:	fct_selNavigate															            *
        // Purpose:		To display the page selected by the user from Quick Navigation drop-down box. The	*
        //              To pass values to detail page use querystring; to list page use cookie.             *
        // Created By:	Nancy Mooney 08/31/2000															    *
        //																									*																				*
        //***************************************************************************************************

        var strCustomerName, intCustomerID, strCustomerShortName;

        strCustomerName = document.frmCustDetail.txtCustomerName.value;
        strCustomerShortName = document.frmCustDetail.txtCustomerShortName.value;
        intCustomerID = document.frmCustDetail.hdnCustomerID.value;

        switch (strPageName) {
            case 'Address':
                //to a list
                document.frmCustDetail.selNavigate.selectedIndex = 0;
                if (strCustomerName != "") { SetCookie("CustomerName", strCustomerName) };
                self.location.href = "SearchFrame.asp?fraSrc=" + strPageName;
                break;
            case 'Asset':
                //to a list
                document.frmCustDetail.selNavigate.selectedIndex = 0;
                if (strCustomerName != "") { SetCookie("CustomerName", strCustomerName) };
                self.location.href = "SearchFrame.asp?fraSrc=" + strPageName;
                break;
            case 'Contact':
                //to a list
                document.frmCustDetail.selNavigate.selectedIndex = 0;
                if (strCustomerName != "") { SetCookie("WorkFor", strCustomerName) };
                self.location.href = "SearchFrame.asp?fraSrc=" + strPageName;
                break;
            case 'ContactRole':
                //to a list
                document.frmCustDetail.selNavigate.selectedIndex = 0;
                if (strCustomerName != "") { SetCookie("CustomerName", strCustomerName) };
                self.location.href = "SearchFrame.asp?fraSrc=" + strPageName;
                break;
            case 'Correlation':
                // to a list
                document.frmCustDetail.selNavigate.selectedIndex = 0;
                if (strCustomerName != "") { SetCookie("CustomerName", strCustomerName) };
                self.location.href = "SearchFrame.asp?fraSrc=" + strPageName;
                break;
            case 'CustServ':
                //to a list
                document.frmCustDetail.selNavigate.selectedIndex = 0;
                if (strCustomerName != "") { SetCookie("CustomerName", strCustomerName) };
                if (intCustomerID != "") { SetCookie("hdnCustomerID", intCustomerID) };
                if (strCustomerShortName != "") { SetCookie("CustomerShortName", strCustomerShortName) };
                self.location.href = "SearchFrame.asp?fraSrc=" + strPageName;
                break;
            case 'Facility':
                //to a list
                document.frmCustDetail.selNavigate.selectedIndex = 0;
                strCustomerName = document.frmCustDetail.txtCustomerName.value;
                if (strCustomerName != "") { SetCookie("CustomerA", strCustomerName) };
                self.location.href = "SearchFrame.asp?fraSrc=" + strPageName;
                break;
            case 'ManagedObjects':
                //to a list
                document.frmCustDetail.selNavigate.selectedIndex = 0;
                if (strCustomerName != "") { SetCookie("CustomerName", strCustomerName) };
                self.location.href = "SearchFrame.asp?fraSrc=" + strPageName;
                break;
            case 'FacilityPVC':
                //to a list
                document.frmCustDetail.selNavigate.selectedIndex = 0;
                if (strCustomerName != "") { SetCookie("CustomerA", strCustomerName) };
                self.location.href = "SearchFrame.asp?fraSrc=" + strPageName;
                break;
            case 'ServLoc':
                //to a list
                document.frmCustDetail.selNavigate.selectedIndex = 0;
                if (strCustomerName != "") { SetCookie("CustomerName", strCustomerName) };
                self.location.href = "SearchFrame.asp?fraSrc=" + strPageName;
                break;
            case 'DEFAULT':
                //do nothing
        }
    }

    function fct_onChange() {
        bolNeedToSave = true;
    }

    function fct_onChangeShortName() {
        document.frmCustDetail.txtCustomerShortName.value = document.frmCustDetail.txtShortName.value;
    }

    function fct_onChangeName() {
        document.frmCustDetail.txtCustName.value = document.frmCustDetail.txtCustomerName.value;
    }

    function fct_onChangeLCD() {
        document.frmCustDetail.selLCDName.value = document.frmCustDetail.selLCDName.value;
    }


    function fct_onSave() {

        if (struserRoles.length > 0 && strCustomerID == "NEW" && document.frmCustDetail.selLCDName.value == "") {
            alert('Missing LCD Name. You must set it as Non-LCD or select LCD Name for this new customer ');
            return false;
        }



        if (((intAccessLevel & intConst_Access_Update) == intConst_Access_Update) || ((intAccessLevel & intConst_Access_Create) == intConst_Access_Create)) {



            //check required fields
            if (document.frmCustDetail.txtCustomerName.value == "") {
                alert('Missing Required Field. Please enter a customer Name.');
                document.frmCustDetail.txtCustomerName.focus();
                return (false);
            }
            if (document.frmCustDetail.selCustStatus.value == "") {
                alert('Missing required field. Please select a Status');
                document.frmCustDetail.selCustStatus.focus();
                return (false);
            }
            if (document.frmCustDetail.selCustRegion.value == "") {
                alert('Missing required field. Please select a Region.');
                document.frmCustDetail.selCustRegion.focus();
                return (false);
            }
            if (document.frmCustDetail.selCustType.value == "") {
                alert('Missing required field. Please select a Type.');
                document.frmCustDetail.selCustType.focus();
                return (false);
            }
            if (document.frmCustDetail.selCustLangPref.value == "") {
                alert('Missing required field. Please select a language preference.');
                document.frmCustDetail.selCustLangPref.focus();
                return (false);
            }

            //check that all phone numbers consist of numbers & 10 chars
            //work phone
            var CustPhone;
            CustPhone = document.frmCustDetail.txtPArea.value + document.frmCustDetail.txtPMid.value + document.frmCustDetail.txtPEnd.value;
            if (CustPhone.length > 0) {
                if (isNaN(CustPhone)) {
                    alert('Phone number must consist of digits only.');
                    document.frmCustDetail.txtPArea.focus();
                    return false;
                }
                if (CustPhone.length < 10) {
                    alert('Phone number must consist of 10 digits (###) ###-####.');
                    document.frmCustDetail.txtPArea.focus();
                    return false;
                }
            }
            //Fax
            var FPhone;
            FPhone = document.frmCustDetail.txtFArea.value + document.frmCustDetail.txtFMid.value + document.frmCustDetail.txtFEnd.value;
            if (FPhone.length > 0) {
                if (isNaN(FPhone)) {
                    alert('Fax number must consist of digits only.');
                    document.frmCustDetail.txtFArea.focus();
                    return false;
                }
                if (FPhone.length < 10) {
                    alert('Fax number must consist of 10 digits (###) ###-####.');
                    document.frmCustDetail.txtFArea.focus();
                    return false;
                }
            }

            //comments field
            var strComments = document.frmCustDetail.txtComments.value;
            if (strComments.length > 2000) {
                alert('Comments can be at most 2000 characters.\n\nYou entered ' + strComments.length + ' character(s).');
                document.frmCustDetail.txtComments.focus();
                return false;
            }

            bolNeedToSave = false; //bypass message asking if you want to save on window_onunload
            //alert ("Need to save flag: " + bolNeedToSave);
            //alert ("fct_onSave success");
            document.frmCustDetail.txtFrmAction.value = 'SAVE';
            document.frmCustDetail.submit();
            return (true);
        }
        else {
            alert('Access denied. Please contact your system administrator.');
            return (false);
        }
    }

    function body_onbeforeunload() {

        document.frmCustDetail.btnSave.focus();
        //alert("body_onbeforeunload: bolNeedToSave: " + bolNeedToSave);
        if ((bolNeedToSave == true) && (((intAccessLevel & intConst_Access_Create) == intConst_Access_Create) || ((intAccessLevel & intConst_Access_Update) == intConst_Access_Update))) {
            event.returnValue = "There is unsaved data on the screen. To save changes, click CANCEL below then click SAVE on the main form.";
        }
    }

    function ClearStatus() {
        window.status = "";
    }

    function DisplayStatus(strWinStatus) {
        window.status = strWinStatus;
        setTimeout('ClearStatus()', "<%=intConst_MessageDisplay%>");
    }

    function btnReferences_onclick() {
        var strOwner = 'CRP';
        var strTableName = 'CUSTOMER';
        var strRecordID = document.frmCustDetail.hdnCustomerID.value;
        var URL;

        if (isNaN(strCustomerID)) {
            alert("No references. This is a new record.");
        }
        else {
            URL = 'Dependency.asp?Owner=' + strOwner + '&TableName=' + strTableName + '&RecordID=' + strRecordID;
            window.open(URL, 'Popup', 'top=100,left=100,WIDTH=500,HEIGHT=300');
        }
    }

    function fct_onDelete() {
        var strCustomerID = document.frmCustDetail.hdnCustomerID.value;
        var strUpdateDateTime = document.frmCustDetail.hdnUpdateDateTime.value;

        if ((intAccessLevel & intConst_Access_Delete) != intConst_Access_Delete) {
            alert('Access denied. Please contact your system administrator.');
            return false;
        }
        if (confirm('Do you really want to delete this customer?')) {
            self.document.location.href = "CustDetail.asp?txtFrmAction=DELETE&CustomerID=" + strCustomerID + "&UpdateDateTime=" + strUpdateDateTime;
        }
    }

    function fct_onReset() {
        if (confirm('All changes will be lost. Do you really want to reset this page?')) {
            bolNeedToSave = false;
            document.location = "CustDetail.asp?CustomerID=<%=strCustomerID%>";
        }
    }

    function fct_onNew() {
        if ((intAccessLevel & intConst_Access_Create) != intConst_Access_Create) { alert('Access denied. Please contact your system administrator.'); return; }
        self.document.location.href = "CustDetail.asp?CustomerID=NEW";
    }

    //-->
    </script>

</head>

<body onload="body_onLoad(strWinMessage);" onbeforeunload="return body_onbeforeunload();">
    <form name="frmCustDetail" action="CustDetail.asp" method="POST">
        <!-- hidden variables -->
        <input name="hdnCustomerID" type="hidden" value="<%if isNumeric(strCustomerID)then Response.Write rsCustomer("customer_id")end if%>">
        <input name="hdnCustLangPref" type="hidden" value="<%if isNumeric(strCustomerID)then Response.Write rsCustomer("language_preference_lcode")end if%>">
        <input name="hdnUpdateDateTime" type="hidden" value="<%if isNumeric(strCustomerID)then Response.Write rsCustomer("last_update_date_time")end if%>">
        <input name="txtFrmAction" type="hidden" value="">
        <input name="txtCustomerShortName" type="hidden" value="<%if isNumeric(strCustomerID) then Response.Write routineHtmlString(rsCustomer("customer_short_name"))end if%>">
        <input name="txtCustName" type="hidden" value="<%if isNumeric(strCustomerID) then Response.Write routineHtmlString(rsCustomer("customer_name"))end if%>">
        <table border="0">
            <thead>
                <tr>
                    <td align="left" colspan="3">Customer Detail</td>
                    <td>
                        <select align="RIGHT" valign="top" id="selNavigate" name="selNavigate" onchange="fct_selNavigate(this.value);" <%if not isNumeric(strCustomerID) then Response.Write " disabled " end if%> tabindex="27">
                            <option value="DEFAULT">Quickly Goto ...</option>
                            <option value="Address">Address</option>
                            <option value="Asset">Asset</option>
                            <option value="Contact">Contact</option>
                            <option value="ContactRole">Contact Role</option>
                            <option value="Correlation">Correlation</option>
                            <option value="CustServ">Customer Service</option>
                            <option value="Facility">Facility</option>
                            <option value="ManagedObjects">Managed Object</option>
                            <option value="FacilityPVC">PVC</option>
                            <option value="ServLoc">Service Location</option>
                        </select></td>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td align="right" width="15%">Name<font color="red">*</font></td>
                    <td align="left" width="35%">
                        <input name="txtCustomerName" tabindex="1" size="50" maxlength="50" onchange="fct_onChange();fct_onChangeName()"
                            <%if isNumeric(strCustomerID)then
					if (intNameAccessLevel and intConst_Access_Update) <> intConst_Access_Update then
						Response.Write " disabled "
					end if
				end if%>
                            value="<%if IsNumeric(strCustomerID) then Response.Write routineHtmlString(rsCustomer("customer_name"))%>">
                    </td>
                    <td align="right" width="15%">Status<font color="red">*</font></td>
                    <td align="left" width="35%">
                        <select name="selCustStatus" tabindex="16" onchange="fct_onChange();">
                            <%while not rsCustStatus.EOF
					Response.write "<OPTION"
					If isNumeric(strCustomerID) then
						if rsCustStatus("customer_status_lcode")= rsCustomer("customer_status_lcode") then
							Response.write " selected "
						end if
					else
						if rsCustStatus("customer_status_lcode") = "Current Customer" then
							Response.Write " selected "
						end if
					end if
					Response.write " value="& rsCustStatus(0) & ">" & routineHtmlString(rsCustStatus(1)) & "</option>" & vbCrLf
					rsCustStatus.MoveNext
				  wend
				  rsCustStatus.Close
				  set rsCustStatus = nothing
                            %>
                        </select>
                    </td>
                </tr>

                <%if  strcomp(hstring,"hidden") <> 0 then
        response.write" <TR> "
        response.write "<TD align=right width=15%" &">"
        'if strcomp(hstring,"hidden") <> 0 then response.write "LCD Name" else response.write "" end if
        response.write "LCD Name"
        response.write "<font color=red>"
        'if strcomp(hstring,"hidden") <> 0 then response.write "*" else response.write "" end if
        response.write "*"
        response.write "</font></TD></TD>"
	    response.write "<TD align=left width=35%" & ">"
	 	response.write "<select name=selLCDName>"

 	    if isNumeric(strCustomerID) then
			'StrSql = "SELECT access_group from crp_sec.access_cid_group"&_
			'		" WHERE customer_id = " & strCustomerID
			StrSql = "SELECT  CRP.sf_get_access_group(" &strCustomerID &") FROM DUAL"

			'Create Recordset object
			dim objRS
			set objRS = objConn.Execute(strSql)
			if strcomp(objRS(0),"NOGROUPCID")=0 or strcomp(objRS(0),"CHECKCID")=0  then
				response.write "CID " &strCustomerID &" has error, please contact SMA support!"
				response.end
			end if

			if strcomp(objRS(0),"NON-LCD")=0 then
				response.write "<option value=""NON-LCD"">Not LCD Customer</option>"
			else
				response.write "<option SELECTED value="&objRS(0) &">" &objRS(0) &"</option>"
			end if
			objConn.close
			set ObjConn = Nothing

        else
	 	     response.write ""
	 	     response.write "<option value=""""></option>"
	 	     response.write "<option value=""NON-LCD"">Not LCD Customer</option>"

	 	 	 dim a, x
	 	   	 a=Split(Session("userRoles"),";")
	       	 for each x in a
	       	  if strcomp(x,"NON-LCD")<>0 then
   		        response.write "<option value="""&x &""">"
   		        if strcomp (hstring,"hidden") <> 0 then
	 	     		response.write x
	 	     	else
	 	     	    response.write ""
	 	     	end if
	 	      end if
             next
        end if

        response.write "</TD> <td></td><td></td>	</tr>"
     end if  %>







                <tr>
                    <td align="right" width="15%">Short Name&nbsp;</td>
                    <td align="left" width="35%">
                        <input name="txtShortName" type="text" tabindex="2" size="15" maxlength="<%=intConst_CustShortNameLength%>" onchange="fct_onChange();fct_onChangeShortName();"
                            <%if isNumeric(strCustomerID)then
					if (intNameAccessLevel and intConst_Access_Update) <> intConst_Access_Update then
						Response.Write " disabled "
					end if
				end if%>
                            value="<%if isNumeric(strCustomerID) then Response.Write routineHtmlString(rsCustomer("customer_short_name"))end if%>"></td>
                    <td align="right" width="15%">Region<font color="red">*</font></td>
                    <td align="left" width="35%">
                        <select name="selCustRegion" tabindex="17" onchange="fct_onChange();">
                            <%while not rsCustRegion.EOF
					Response.write "<OPTION"
					if isNumeric(strCustomerID) then
						if rsCustRegion("noc_region_lcode")= rsCustomer("noc_region_lcode") then
							Response.write " selected "
						end if
					else
						if rsCustRegion("noc_region_lcode") = "AB" then
							Response.Write " selected "
						end if
					end if
					Response.write " value=" & rsCustRegion(0)& ">" & routineHtmlString(rsCustRegion(1)) & "</option>"&vbCrLf
					rsCustRegion.MoveNext
				  wend
				  rsCustRegion.Close
				  set rsCustRegion = nothing
                            %>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td align="right" width="15%">Phone&nbsp;</td>
                    <td width="35%">(<input name="txtPArea" tabindex="3" size="3" maxlength="3" style="height: 20px; width: 30px" onchange="fct_onChange();" value="<%if isNumeric(strCustomerID) then Response.Write strPArea end if%>">)
			<input name="txtPMid" tabindex="4" size="3" maxlength="3" style="height: 20px; width: 30px" onchange="fct_onChange();" value="<%if isNumeric(strCustomerID) then Response.Write strPMid end if%>">
                        -<input name="txtPEnd" tabindex="5" size="4" maxlength="4" style="height: 20px; width: 35px" onchange="fct_onChange();" value="<%if isNumeric(strCustomerID) then Response.Write strPEnd end if%>"></td>
                    <td align="right" width="15%">Type<font color="red">*</font></td>
                    <td align="left" width="35%" valign="top">
                        <select name="selCustType" tabindex="18" onchange="fct_onChange();">
                            <option <%if isNumeric(strCustomerID) then if rsCustomer("customer_type_ind") = "ORG" then Response.write " selected " end if else Response.Write " selected " end if%> value="ORG">
                            Organization
				<option <%if isNumeric(strCustomerID) then if rsCustomer("customer_type_ind") = "PER" then Response.write " selected " end if end if%> value="PER">
                            Person
                        </select>
                    </td>
                </tr>
                <tr>
                    <td align="right" width="15%">Fax&nbsp;</td>
                    <td width="35%">(<input name="txtFArea" tabindex="6" size="3" maxlength="3" onchange="fct_onChange();" style="height: 20px; width: 30px" value="<%if isNumeric(strCustomerID) then Response.Write strFArea end if%>">)
			<input name="txtFMid" tabindex="7" size="3" maxlength="3" onchange="fct_onChange();" style="height: 20px; width: 30px" value="<%if isNumeric(strCustomerID) then Response.Write strFMid end if%>">
                        -<input name="txtFEnd" tabindex="8" size="4" maxlength="4" onchange="fct_onChange();" style="height: 20px; width: 35px" value="<%if isNumeric(strCustomerID) then Response.Write strFEnd end if %>"></td>
                    <td align="right" width="15%">Language Pref<font color="red">*</font></td>
                    <td align="left" width="35%">
                        <select name="selCustLangPref" tabindex="19" onchange="fct_onChange();">
                            <% while not rsCustLangPref.EOF
					Response.write "<OPTION"
					if isNumeric(strCustomerID) then
						if rsCustLangPref("language_preference_lcode")= rsCustomer("language_preference_lcode") then
							Response.write " selected "
						end if
					else
						if rsCustLangPref("language_preference_lcode") = "EN" then
							Response.Write " selected "
						end if
					end if
					Response.write " value="& rsCustLangPref(0) & ">" & routineHtmlString(rsCustLangPref(1)) & "</option>"&vbCrLf
					rsCustLangPref.MoveNext
				  wend
				  rsCustLangPref.Close
				  set rsCustLangPref = nothing
                            %>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td align="right" width="15%">Email Address&nbsp;</td>
                    <td width="100%">
                        <input name="txtEmail" tabindex="9" size="50" maxlength="50" style="width: 10cm" onchange="fct_onChange();" value="<%if isNumeric(strCustomerID) then Response.write routineHTMLString(rsCustomer("email_address"))end if%>">
                    <td align="right" width="15%" valign="top">Industry&nbsp;</td>
                    <td width="35%" valign="top">
                        <select name="selCustIndustry" tabindex="20" onchange="fct_onChange();">
                            "<option></option>
                            "
				<%
				while not rsCustIndustry.EOF
					Response.write "<OPTION"
					if isNumeric(strCustomerID) then
						if (rsCustomer("industry_id") <> "")then
							if (cInt(rsCustIndustry("industry_id"))= cInt(rsCustomer("industry_id")))then
								Response.write " selected"
							end if
						end if
					end if
					Response.write " value="& rsCustIndustry("industry_id") & ">" & routineHtmlString(rsCustIndustry("industry_desc")) & "</option>"&vbCrLf
					rsCustIndustry.MoveNext
				wend
				rsCustIndustry.Close
				set rsCustIndustry = nothing
                %>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td align="right" colspan="2" width="35%">
                        <input name="btnCustOrgMaint" tabindex="11" onclick="CustOrgMClick();" type="button" value="Customer Organization"></td>
                    <td width="50%" colspan="2">&nbsp;</td>
                </tr>


                <tr>

                    <td align="right" width="15%">Web Site&nbsp;</td>
                    <td width="35%">
                        <input name="txtWebSite" tabindex="10" size="50" maxlength="50" onchange="fct_onChange();" value="<%if isNumeric(strCustomerID) then Response.Write routineHTMLString(rsCustomer("web_site_url"))end if%>"></td>
                    <td width="50%" colspan="2">&nbsp;</td>
                </tr>
                <tr>
                    <td width="15%" align="right" valign="top">Name alias&nbsp;</td>
                    <td width="35%" valign="top">
                        <iframe tabindex="11" id="aifr" width="100%" height="100" src="" scrolling="yes" marginheight="1" marginwidth="1"></iframe>
                        <br>
                        <input type="button" tabindex="12" style="width: 2cm" value="Delete" name="btn_iFrameDelete" onclick="btn_iFrmDelete();" class="button">&nbsp;
			<input type="button" tabindex="13" style="width: 2cm" value="Refresh" name="btn_iFrameRefresh" onclick="iFrame_display();" class="button">&nbsp;
			<input type="button" tabindex="14" style="width: 2cm" value="New" name="btn_iFrameAdd" onclick="btn_iFrmAdd();" class="button">&nbsp;
			<input type="button" tabindex="15" style="width: 2cm" value="Update" name="btn_iFrameUpdate" onclick="btn_iFrmUpdate();" class="button">
                    </td>
                    <td align="right" valign="top" width="15%">Comments&nbsp;</td>
                    <td width="35%" valign="top">
                        <textarea name="txtComments" tabindex="21" rows="6" onchange="fct_onChange();" style="width: 350px"><%if isNumeric(strCustomerID) then Response.Write routineHtmlString(rsCustomer("Comments"))end if%></textarea></td>
                </tr>
                <tr>
                    <th colspan="4" align="left">&nbsp;</th>
                </tr>
            </tbody>
            <tfoot>
                <tr>
                    <td align="right" colspan="4">
                        <input name="btnReferences" tabindex="22" type="button" value="References" style="width: 2cm" onclick="return btnReferences_onclick();">&nbsp;&nbsp;
			<input name="btnDelete" tabindex="23" type="button" value="Delete" style="width: 2cm" language="javascript" onclick="return fct_onDelete();">&nbsp;&nbsp;
			<input name="btnReset" tabindex="24" type="button" value="Reset" style="width: 2cm" language="javascript" onclick="fct_onReset();">&nbsp;&nbsp;
		    <input name="btnNew" tabindex="25" type="button" value="New" style="width: 2cm" language="javascript" onclick="fct_onNew();">&nbsp;&nbsp;
			<input name="btnSave" tabindex="26" type="button" value="Save" style="width: 2cm" language="javascript" onclick="fct_onSave();">&nbsp;&nbsp;
                    </td>
                </tr>
            </tfoot>
        </table>
        <%
if isNumeric(strCustomerID) then
	if bolAddressFlag = true then

		dim intTotal
		intTotal = UBound(alist,2) + 1
		'Response.Write "<TR align=left><TH>Total records: " & intTotal & "</TH></TR>"
        %>

        <table border="1" cellpadding="2" cellspacing="0" width="100%">
            <thead>
                <tr>
                    <td colspan="9">Key Addresses</td>
                </tr>
                <tr>
                    <th>Primary</th>
                    <th>Billing</th>
                    <th>Mailing</th>
                    <th align="left">Building</th>
                    <th align="left">Street</th>
                    <th align="left">City</th>
                    <th align="left">Prov/State</th>
                    <th align="left">Country</th>
                    <th align="left">Postal Code</th>
                </tr>
            </thead>
            <tbody>
                <%
				dim intCnt, strBilling, strPrimary, strMailing
				'display the table
				for intCnt=0 to (intTotal-1)

					'Alternate row background colour
					'if Int(intCnt/2) = intCnt/2 then
					'	Response.write "<TR bgcolor=White>"
					'else
					'	Response.write "<TR>"
					'end if
					Response.Write "<TR>"

					'set check boxes
					if alist(1,intCnt) = "Y" then
						strBilling = "=yes checked"
					else
						strBilling = ""
					end if

					if alist(2,intCnt) = "Y" then
						strPrimary = "=yes checked"
					else
						strPrimary = ""
					end if

					if alist(3, intCnt) = "Y" then
						strMailing = "=yes checked"
					else
						strMailing = ""
					end if

					'format postal code
					dim strPCBegin, strPCEnd, intPClen, strPC, strAddress
						strPC = alist(11,intCnt)
						if strPC <> "" then
							select case alist(7,intCnt)
								case "CA"
									strPCBegin = mid(strPC,1,3)
									strPCEnd = mid(strPC,4,3)
									strPC = strPCBegin & " " & strPCEnd
								case "US"
									intPClen = len(strPC)
									strPCBegin = mid(strPC,1,5)
									strPCEnd = mid(strPC,6,intPCLen-5)
									strPC = strPCBegin & " " & strPCEnd
							end select
						end if

					'display table row
					Response.Write "<TD NOWRAP disabled align=""center""><INPUT ID=""Primary""  name=""primary"" type=""checkbox""  VALUE" &strPrimary& "></TD>" &vbCrLf
					Response.Write "<TD NOWRAP disabled align=""center""><INPUT ID=""Billing""  name=""billing"" type=""checkbox"" VALUE" &strBilling& "></TD>" &vbCrLf
					Response.Write "<TD NOWRAP disabled align=""center""><INPUT ID=""Mailing""  name=""mailing"" type=""checkbox""  VALUE" &strMailing& "></TD>" &vbCrLf
					Response.Write "<TD NOWRAP >"&routineHtmlString(aList(8,intCnt))&"</TD>"&vbCrLf
					Response.Write "<TD NOWRAP >"&routineHtmlString(aList(4,intCnt))&"</TD>"&vbCrLf
					Response.Write "<TD NOWRAP >"&routineHtmlString(aList(5,intCnt))&"</TD>"&vbCrLf
					Response.Write "<TD NOWRAP >"&routineHtmlString(aList(9,intCnt))&"</TD>"&vbCrLf
					Response.Write "<TD NOWRAP >"&routineHtmlString(aList(10,intCnt))&"</TD>"&vbCrLf
					Response.Write "<TD NOWRAP >"& strPC &"</TD>"&vbCrLf
					Response.Write "</TR>"
				next
	end if 'bolAddressFlag

	if bolCustCareFlag = true then

		dim intTot
		intTot = UBound(clist,2) + 1
		'Response.Write "<TR align=left><TH>Total records: " & intTot & "</TH></TR>"
                %>

                <table border="1" cellpadding="2" cellspacing="0" width="100%">
                    <thead>
                        <tr>
                            <td colspan="8">Service Management Rep<%if intTot > 1 then Response.Write "s" end if%>&nbsp;(role = custcare)</td>
                        </tr>
                        <tr>
                            <th align="left">Name</th>
                            <th align="center">Priority</th>
                            <th align="left">Work Phone</th>
                            <th align="left">Ext</th>
                            <th align="left">Cell</th>
                            <th align="left">Pager</th>
                            <th align="left">Fax</th>
                            <th align="left">Email</th>
                        </tr>
                    </thead>
                    <tbody>
                        <% 		dim intCount
				'display the table
			for intCount=0 to (intTot-1)

				'Parse out the phone number
				if (cList(2,intCount) <> "") then
					Dim strWPArea,strWPMid,strWPEnd,strWP
			 		strWP = cList(2,intCount)
			 		strWPArea = mid(strWP,1,3)
			 		strWPMid = mid(strWP,4,3)
			 		strWPEnd = mid(strWP,7,4)
			 		strWP = "(" & strWPArea & ") " & strWPMid & "-" & strWPEnd
			 		If strWP = "() -" then
			 			strWP = ""
			 		End If
			 	end if

				'Parse out the cell phone number
				if (cList(4, intCount) <> "") then
					Dim strClPArea,strClPMid,strClPEnd,strClP
			 		strClP = cList(4,intCount)
			 		strClPArea = mid(strClP,1,3)
			 		strClPMid = mid(strClP,4,3)
			 		strClPEnd = mid(strClP,7,4)
			 		strClP = "(" & strClPArea & ") " & strClPMid & "-" & strClPEnd
			 		If strClP = "() -" then
			 			strClP = ""
			 		End If
			 	end if

				'Parse out the pager number
				if (cList(5,intCount) <> "") then
					Dim strPPArea,strPPMid,strPPEnd,strPP
			 		strPP = cList(5,intCount)
			 		strPPArea = mid(strPP,1,3)
			 		strPPMid = mid(strPP,4,3)
			 		strPPEnd = mid(strPP,7,4)
			 		strPP = "(" & strPPArea & ") " & strPPMid & "-" & strPPEnd
			 		If strPP = "() -" then
			 			strPP = ""
			 		End If
			 	end if

	 			'Parse out the fax number
	 			if (cList(6, intCount) <> "") then
	 				Dim strFPArea,strFPMid,strFPEnd,strFP
			 		strFP = cList(6,intCount)
			 		strFPArea = mid(strFP,1,3)
			 		strFPMid = mid(strFP,4,3)
			 		strFPEnd = mid(strFP,7,4)
			 		strFP = "(" & strFPArea & ") " & strFPMid & "-" & strFPEnd
			 		If strFP = "() -" then
			 			strFP = ""
			 		End If
			 	end if

				'display table row
				Response.Write "<TD NOWRAP >"&routineHtmlString(cList(0,intCount))&"&nbsp;</TD>"&vbCrLf
				Response.Write "<TD NOWRAP align=center >"&cList(1,intCount)&"&nbsp;</TD>"&vbCrLf
				Response.Write "<TD NOWRAP style='width: 100px'>"&routineHtmlString(strWP)&"&nbsp;</TD>"&vbCrLf
				Response.Write "<TD NOWRAP >"&routineHtmlString(cList(3,intCount))&"&nbsp;</TD>"&vbCrLf
				Response.Write "<TD NOWRAP style='width: 100px'>"&routineHtmlString(strClP)&"&nbsp;</TD>"&vbCrLf
				Response.Write "<TD NOWRAP style='width: 100px'>"&routineHtmlString(strPP)&"&nbsp;</TD>"&vbCrLf
				Response.Write "<TD NOWRAP style='width: 100px'>"&routineHtmlString(strFP)&"&nbsp;</TD>"&vbCrLf
				Response.Write "<TD NOWRAP >"&routineHtmlString(cList(7,intCount))&"&nbsp;</TD>"&vbCrLf
				Response.Write "</TR>"
			next
	else %>
                        <table border="1" cellpadding="2" cellspacing="0" width="100%">
                            <thead>
                                <tr>
                                    <td colspan="7">Currently there is no Service Management Rep (contact role = custcare) for this customer.</td>
                                </tr>
                            </thead>
                        </table>
                        <%end if 'bolCustCareFlag
end if 'isNumeric(strCustomerID)
                        %>
                    </tbody>
                </table>
                <fieldset>
                    <legend align="right"><b>Audit Information</b></legend>
                    <div size="8pt" align="RIGHT">
                        Record Status Indicator
		<input align="left" name="txtRecordStatusInd" type="text" style="height: 20px; width: 18px" disabled value="<%if isNumeric(strCustomerID) then Response.Write rsCustomer("record_status_ind")end if%>">&nbsp;&nbsp;&nbsp;
		Create Date
		<input align="center" name="txtRecordStatusInd" type="text" style="height: 20px; width: 150px" disabled value="<%if isNumeric(strCustomerID) then Response.Write rsCustomer("create_date")end if%>">&nbsp;
		Created By
		<input align="right" name="txtRecordStatusInd" type="text" style="height: 20px; width: 200px" disabled value="<%if isNumeric(strCustomerID) then Response.Write rsCustomer("create_real_userid")end if%>"><br>
                        Customer ID
		<input align="left" name="txtRecordStatusInd" type="text" style="height: 20px; width: 55px" disabled value="<%if isNumeric(strCustomerID) then Response.Write rsCustomer("customer_id")end if%>">&nbsp;
		Update Date
		<input align="center" name="txtRecordStatusInd" type="text" style="height: 20px; width: 150px" disabled value="<%if isNumeric(strCustomerID) then Response.Write rsCustomer("update_date")end if%>">&nbsp;
		Updated By
		<input align="right" name="txtRecordStatusInd" type="text" style="height: 20px; width: 200px" disabled value="<%if isNumeric(strCustomerID) then Response.Write rsCustomer("update_real_userid")end if%>">
                    </div>
                </fieldset>
    </form>
</body>
</html>
<%
	if isNumeric(strCustomerID)then
		rsCustomer.Close
		set rsCustomer = nothing
	end if
	'close the connection
	'objConn.close
	set objConn = nothing
%>