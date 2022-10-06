<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Buffer = true %>
<!-- #include file="../smaConstants.inc" -->
<!-- #include file="../smaProcs.inc" -->

<!--

********************************************************************************************
* Page name:	DoFacilitDetail.asp
* Purpose:	Act like a web service to create new Facility
*		Note:- Only supports ADSL
*
* In Param:	Action - Action to perform 'new' or 'update'
*           	UserName - User name
*		Password - Password
*		FacTyp - Facility Type
*		FacNumb - Facility Number
*		FacName - Facility Name
*		AdslTypCode - ADSL Type Code
*		AdslCpeOwn - CPE Ownership Flag
*		FacProvCode - Facility Provider Code
*		RegionCode - Region Code
*		OpStat - Operational Status
*		CustACsid - Customer Service Id A
*		CustBCsid - Customer Service Id B
*		StrAdslDue - ADSL Due Date
*		StrAdslSlot - ADSL Slot	
*		StrAdslShelf - ADSL Shelf		
*	
* Example:
*    http://abdev018:8080/sma2/webservices/DoFacilityDetail.asp?Action=new&UserName=t820429&Password=password&FacTyp=ADSL
*              &FacNumb=Peter6&FacName=PeterName&AdslTypCode=BUS4&AdslCpeOwn=N&FacProvCode=TELUS&RegionCode=BC
*              &OpStat=DEFINE&CustACsid=1037894&CustBCsid=1007611&AdslDue=12/07/2006&AdslSlot=123%20456&AdslShelf=456			
*
***************************************************************************************************
*   Date	Author		Changes/enhancements made
* 06 Dec 2006	Peter Smith	New page
* 11 Jan 2007	Peter Smith	Added new or update action.
**************************************************************************************************
-->
<%
Dim strUserName
Dim strUserPass
Dim strAction
Dim objConn
Dim objRs
Dim noLevel
Dim strSQL

' Validate the user
strUserName = Request("UserName")
strUserPass = Request("Password")
If (strUserPass = "") Or (strUserName = "") Then
	Response.Write ("User id or password is missing")
	Response.End
End If
' Validate the Action
strAction = Request("Action")
If (strAction <> "new") And (strAction <> "update") Then
	Response.Write ("No or invalid action - " & strAction)
	Response.End
End If
' Connect to database
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.ConnectionString = strConstConnectString
objConn.Open
If err Then
	Response.Write (err.Number & " Cannot connect to database - " & err.Description)
	Response.End
End If
' Check user
strSQL = "SELECT SEC.USERID " &_
		"FROM MSACCESS.TBLSECURITY SEC, CRP.CONTACT CON " &_
		"WHERE CON.CONTACT_ID = SEC.STAFFID AND " &_
		"SEC.USERID = '" & strUserName & "' AND " &_
		"SEC.PASSWORD = '" & strUserPass & "'"
Set objRs = Server.CreateObject("ADODB.Recordset")
objRs.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
If objRs.EOF Then
	Response.Write ("Incorrect User ID/Password, or User not defined in SMA.")
	Response.End
End If
objRs.Close

' Authenticate the user
strSQL = " SELECT nvl(bfa.access_level, 0) access_level" &_
	  " ,      b.business_function_id" &_
	  " FROM ( SELECT a.business_function_id" &_
	  "        ,      a.business_func_access_level_id" &_
	  "        ,      c.access_level" &_
	  "        FROM   msaccess.tblsecurity s" &_
	  "        ,      msaccess.staff_security_role r" &_
	  "        ,      msaccess.security_role t" &_
	  "        ,      msaccess.business_func_security_role a" &_
	  "        ,      msaccess.business_func_access_level c" &_
 	  "        WHERE  s.userid = '" & routineOraString(strUserName) & "'" &_
	  "        AND    s.staffid = r.staff_id" &_
	  "        AND    r.security_role_id = t.security_role_id" &_
	  "        AND    t.security_role_name LIKE 'SMA%'" &_
	  "        AND    t.security_role_id = a.security_role_id" &_
	  "        AND    a.business_func_access_level_id = c.business_func_access_level_id" &_
	  "      ) bfa" &_
	  " ,      msaccess.business_function b" &_
	  " ,      msaccess.application a" &_
	  " WHERE  bfa.access_level is not null and bfa.business_function_id (+)= b.business_function_id" &_
	  " AND    b.application_id = a.application_id" &_
	  " AND    a.application_name = 'SMA2' " &_
	  " ORDER BY business_function_id"
	
objRs.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
If err Then
	Response.Write (err.Number & "Cannot open record set..." & err.Description)
	Response.End
End If
If objRs.EOF Then
	Response.Write ("User has no assigned SMA role. - " & strUserName)
	Response.End 
End If
' Check the access level
noLevel = True
Do While Not objRs.EOF
	If ((objRs("ACCESS_LEVEL") And intConst_Access_Create) = intConst_Access_Create) Then
		noLevel = False
		Exit Do
	End If
	objRs.MoveNext			
Loop
If (noLevel) Then
	Response.Write ("Access level is too low for " & strUserName)
	Repsnse.End
End If
objRs.close
'trap unexpected error
If err Then
	Response.Write (err.Number & " Unexpected error. " & err.Description)
	Response.End
End If

'
' Now add the facility
'
Dim StrFacTyp
Dim StrFacNumb
Dim StrFacName
Dim StrAdslTypCode
Dim StrAdslCpeOwn
Dim StrFacProvCode
Dim StrRegionCode
Dim StrOpStat
Dim StrCustACsid
Dim StrCustAId
Dim StrCustALocId
Dim StrCustBCsid
Dim StrCustBLocId
Dim StrAdslDue
Dim StrAdslSlot
Dim StrAdslShelf

Dim StrCircuitID
Dim StrUpdateDateTime

Dim cmdUpdateObj
Dim cmdObj


' Get the request parameters
StrFacTyp = Request("FacTyp")
StrFacNumb = Request("FacNumb")
StrFacName = Request("FacName")
StrAdslTypCode = Request("AdslTypCode")
StrAdslCpeOwn = Request("AdslCpeOwn")
StrFacProvCode = Request("FacProvCode")
StrRegionCode = Request("RegionCode")
StrOpStat = Request("OpStat")
StrCustACsid = Request("CustACsid")
StrCustBCsid = Request("CustBCsid")
StrAdslDue = Request("AdslDue")
StrAdslSlot = Request("AdslSlot")
StrAdslShelf = Request("AdslShelf")

' Validate the Facility Type
If (StrFacTyp = "") Then
	Response.Write ("Missing Facility Type.")
	Response.End	
End If
If (StrFacTyp <> "ADSL") Then
	Response.Write ("Invalid Facility Type, must be ADSL.")
	Response.End	
End If
StrSql = "SELECT CIRCUIT_TYPE_CODE FROM CRP.CIRCUIT_TYPE WHERE CIRCUIT_TYPE_CODE = '" & StrFacTyp & "'"
objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
If (objRs.EOF) Then
	Response.Write ("Invalid Facility Type - " & StrFacTyp)
	Response.End
End If
objRs.close

' Validate Facility number
If (StrFacNumb = "") Then
	Response.Write ("Missing Facility Number.")
	Response.End	
End If

' Validate Facility Name
If (StrFacName = "") Then
	Response.Write ("Missing Facility Name.")
	Response.End	
End If

' Validate the ADSL Service Type
If (StrAdslTypCode = "") Then
	Response.Write ("Missing ADSL Service Type.")
	Response.End	
End If
StrSql = "SELECT ADSL_TYPE_CODE FROM CRP.ADSL_TYPE WHERE RECORD_STATUS_IND = 'A' AND ADSL_TYPE_CODE = '" & StrAdslTypCode & "'"
objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
If (objRs.EOF) Then
	Response.Write ("Invalid ADSL Service Type Code - " & StrAdslTypCode)
	Response.End
End If
objRs.close

' Validate the ADSL CPE Ownership
If (StrAdslCpeOwn = "") Then
	Response.Write ("Missing ADSL CPE Ownership flag.")
	Response.End	
End If
If (StrAdslCpeOwn <> "Y" And StrAdslCpeOwn <> "N") Then
	Response.Write ("Invalid ADSL CPE Ownership flag, must be 'Y' or 'N'. - " & StrAdslCpeOwn)
	Response.End	
End If    

' Validate the Facility Provider
If (StrFacProvCode = "") Then
	Response.Write ("Missing Facility Provider.")
	Response.End	
End If
StrSql = "SELECT CIRCUIT_PROVIDER_CODE FROM CRP.CIRCUIT_PROVIDER WHERE RECORD_STATUS_IND = 'A' AND CIRCUIT_PROVIDER_CODE = '" & StrFacProvCode & "'"
objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
If (objRs.EOF) Then
	Response.Write ("Invalid Facility Provider - " & StrFacProvCode)
	Response.End
End If
objRs.close

' Validate Region
If (StrRegionCode = "") Then
	Response.Write ("Missing Region.")
	Response.End	
End If
StrSql = "SELECT NOC_REGION_LCODE FROM CRP.LCODE_NOC_REGION WHERE RECORD_STATUS_IND = 'A' AND NOC_REGION_LCODE = '" & StrRegionCode & "'"
objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
If (objRs.EOF) Then
	Response.Write ("Invalid Region - " & StrRegionCode)
	Response.End
End If
objRs.close

' Validate Operational Status
If (StrOpStat = "") Then
	Response.Write ("Missing Operational Status.")
	Response.End	
End If
StrSql = "SELECT CIRCUIT_STATUS_CODE FROM CRP.CIRCUIT_STATUS WHERE RECORD_STATUS_IND = 'A' AND CIRCUIT_STATUS_CODE = '" & StrOpStat & "'"
objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
If (objRs.EOF) Then
	Response.Write ("Invalid Operational Status - " & StrOpStat)
	Response.End
End If
objRs.close

' Validate Customer A 
If (StrCustACsid = "") Then
	Response.Write ("Missing Customer Service Id A.")
	Response.End	
End If
StrSql = "SELECT CUSTOMER_ID, SERVICE_LOCATION_ID FROM CRP.CUSTOMER_SERVICE WHERE CUSTOMER_SERVICE_ID = '" & StrCustACsid & "'"
objRS.Open strSQL, objConn, adOpenForwardOnly, adLockReadOnly, adCmdText
If (objRs.EOF) Then
	Response.Write ("Invalid Customer Service Id A - " & StrCustACsid)
	Response.End
End If
StrCustAId = objRs("CUSTOMER_ID")
StrCustALocId = objRs("SERVICE_LOCATION_ID")
objRs.close

' Validate Customer B
If (StrCustBCsid = "") Then
	Response.Write ("Missing Customer Service Id B.")
	Response.End	
End If
StrSql = "SELECT SERVICE_LOCATION_ID FROM CRP.CUSTOMER_SERVICE WHERE CUSTOMER_SERVICE_ID = '" & StrCustBCsid & "'"
Set objRs = objConn.Execute(StrSql)
If (objRs.EOF) Then
	Response.Write ("Invalid Customer Service Id B - " & StrCustBCsid)
	Response.End
End If
StrCustBLocId = objRs("SERVICE_LOCATION_ID")
objRs.close

' Validate ADSL Due Date
If (StrAdslDue <> "") Then
	If (Not IsDate(StrAdslDue)) Then
		Response.Write ("Invalid ADSL Due Date - " & StrAdslDue)
		Response.End
	End If
End If

' Get the circuit id for updates
If (strAction <> "new") Then 
  StrSql = "SELECT CIRCUIT_ID, UPDATE_DATE_TIME FROM CRP.CIRCUIT WHERE CIRCUIT_TYPE_CODE = '" & StrFacTyp & "'" & _
           " AND CIRCUIT_NUMBER = '" & StrFacNumb & "'" & _ 
           " AND CIRCUIT_NAME = '" & StrFacName & "'"
  Set objRs = objConn.Execute(StrSql)
  If (objRs.EOF) Then
	Response.Write ("Cannot find circuit id  - " & StrFacName)
	Response.End
  End If
  StrCircuitID = objRs("CIRCUIT_ID")
  StrUpdateDateTime = objRs("UPDATE_DATE_TIME")
  objRs.close
End IF

'
' Create or update a record
'
Set cmdObj = server.CreateObject("ADODB.Command")
Set cmdObj.ActiveConnection = objConn
cmdObj.CommandType = adCmdStoredProc
' Set the procedure depending upon the action we want
If (strAction = "new") Then
	cmdObj.CommandText = "sma_sp_userid.spk_sma_fac_inter.sp_fac_insert" 
Else
	cmdObj.CommandText = "sma_sp_userid.spk_sma_fac_inter.sp_fac_update" 
End If
'create parameters
' Note:- The order the parameters are added matters, they must be in this order.
cmdObj.Parameters.Append cmdObj.CreateParameter("p_user_id", adVarChar , adParamInput, 30, strUserName)
If (strAction = "new") Then
  cmdObj.Parameters.Append cmdObj.CreateParameter("p_circuit_id",adNumeric , adParamOutput,,null) 
Else
  cmdObj.Parameters.Append cmdObj.CreateParameter("p_circuit_id",adNumeric , adParamInput,,Clng(StrCircuitID))
End If
cmdObj.Parameters.Append cmdObj.CreateParameter("p_circuit_number", adVarChar,adParamInput, 50, StrFacNumb)
cmdObj.Parameters.Append cmdObj.CreateParameter("p_circuit_name", adVarChar,adParamInput, 65, StrFacName)
cmdObj.Parameters.Append cmdObj.CreateParameter("p_circuit_type", adVarChar,adParamInput, 6, StrFacTyp)
cmdObj.Parameters.Append cmdObj.CreateParameter("p_circuit_provider", adVarChar,adParamInput, 6, StrFacProvCode)
cmdObj.Parameters.Append cmdObj.CreateParameter("p_circuit_ems", adVarChar,adParamInput, 10, null)
cmdObj.Parameters.Append cmdObj.CreateParameter("p_noc_region", adVarChar,adParamInput, 8, StrRegionCode)
If (strAction <> "new") Then
    cmdObj.Parameters.Append cmdObj.CreateParameter("p_last_update_dt", adDBTimeStamp, adParamInput, , CDate(StrUpdateDateTime))
End If
cmdObj.Parameters.Append cmdObj.CreateParameter("p_circuit_status", adVarChar, adParamInput, 6,StrOpStat)
cmdObj.Parameters.Append cmdObj.CreateParameter("p_adsl_cpe_flag", adChar, adParamInput, 1,  StrAdslCpeOwn)
cmdObj.Parameters.Append cmdObj.CreateParameter("p_customer_service_id_a",adNumeric , adParamInput,, Clng(StrCustACsid))
cmdObj.Parameters.Append cmdObj.CreateParameter("p_billing_customer_id_a",adNumeric , adParamInput,, Clng(StrCustAId)) 	
cmdObj.Parameters.Append cmdObj.CreateParameter("p_service_location_id_a",adNumeric , adParamInput,, Clng(StrCustALocId))
cmdObj.Parameters.Append cmdObj.CreateParameter("p_customer_service_id_b",adNumeric , adParamInput,, null) 
cmdObj.Parameters.Append cmdObj.CreateParameter("p_billing_customer_id_b",adNumeric , adParamInput,, null) 	
cmdObj.Parameters.Append cmdObj.CreateParameter("p_service_location_id_b",adNumeric , adParamInput,, Clng(StrCustBLocId))
cmdObj.Parameters.Append cmdObj.CreateParameter("p_circuit_start_dt",adVarChar,adParamInput,20 , null)
cmdObj.Parameters.Append cmdObj.CreateParameter("p_billing_type", adVarChar, adParamInput, 10, null)	
cmdObj.Parameters.Append cmdObj.CreateParameter("p_usage_calculation_type", adVarChar, adParamInput, 6, null) 
cmdObj.Parameters.Append cmdObj.CreateParameter("p_fr_dlci_from", adVarChar, adParamInput, 30 ,null)
cmdObj.Parameters.Append cmdObj.CreateParameter("p_fr_dlci_to", adVarChar, adParamInput, 30, null) 
cmdObj.Parameters.Append cmdObj.CreateParameter("p_fibre_order_no", adVarChar, adParamInput, 30,null)
cmdObj.Parameters.Append cmdObj.CreateParameter("p_fibre_check_no", adVarChar, adParamInput, 30, null)   
If (StrAdslShelf <> "") Then  
	cmdObj.Parameters.Append cmdObj.CreateParameter("p_adsl_shelf_no", adVarChar, adParamInput, 30, StrAdslShelf)
Else
	cmdObj.Parameters.Append cmdObj.CreateParameter("p_adsl_shelf_no", adVarChar, adParamInput, 30, null) 
End If 
If (StrAdslSlot <> "") Then  
	cmdObj.Parameters.Append cmdObj.CreateParameter("p_adsl_slot_no", adVarChar, adParamInput, 30, StrAdslSlot)
Else
	cmdObj.Parameters.Append cmdObj.CreateParameter("p_adsl_slot_no", adVarChar, adParamInput, 30, null)
End If 
cmdObj.Parameters.Append cmdObj.CreateParameter("p_adsl_loop_loss", adVarChar, adParamInput, 30 ,null)
cmdObj.Parameters.Append cmdObj.CreateParameter("p_adsl_trained_speed", adVarChar, adParamInput, 30 ,null) 
cmdObj.Parameters.Append cmdObj.CreateParameter("p_adsl_dist_block", adVarChar, adParamInput, 30, null) 
cmdObj.Parameters.Append cmdObj.CreateParameter("p_adsl_type", adVarChar, adParamInput, 6, StrAdslTypCode)
If (StrAdslDue <> "") Then  
	cmdObj.Parameters.Append cmdObj.CreateParameter("p_adsl_due_dt", adVarChar, adParamInput,20 , StrAdslDue)
Else
	cmdObj.Parameters.Append cmdObj.CreateParameter("p_adsl_due_dt", adVarChar, adParamInput,20 , null)
End If		
cmdObj.Parameters.Append cmdObj.CreateParameter("p_adsl_order_no", adVarChar, adParamInput,20 ,null)
cmdObj.Parameters.Append cmdObj.CreateParameter("p_comments", adVarChar, adParamInput, 2000 ,null) 


'call the insert stored proc 
'Response.Write cmdObj.CommandText & "<br>"
'dim objparm
'for each objparm in cmdObj.Parameters
' Response.Write "<b>" & objparm.name & "</b>"
' Response.Write " has size:  " & objparm.Size & " "
' Response.Write " and value:  <b>" & objparm.value & "</b> "
' Response.Write " and datatype:  " & objparm.Type & "<br> "
'next
'Response.End
 			
cmdObj.Execute
  			
if objConn.Errors.Count <> 0 then
	Response.Write (objConn.Errors(0).NativeError & " CANNOT CREATE NEW FACILITY " & objConn.Errors(0).Description)
	objConn.Errors.Clear
else
  If (strAction = "new") Then
	StrCircuitID = cmdObj.Parameters("p_circuit_id").Value
	Response.Write ("Facility id " & StrCircuitID & " created successfully.")
  Else
	Response.Write ("Facility id " & StrCircuitID & " updated successfully.")
  End If
end if
if err then
  Response.Write (err.Number & " CANNOT CREATE FACILITY" & err.Description)
end if
%>
