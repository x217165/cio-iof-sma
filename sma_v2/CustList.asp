<%@  language="VBScript" %>
<% option explicit %>
<!--% on error resume next%-->
<!--
********************************************************************************************
* Page name:	CustList.asp
* Purpose:		To display the results of a customer search.
*				Search criteria are chosen via CustCriteria.asp
*
* Created by:	Nancy Mooney	08/01/2000
*
********************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       04-30-01      AHaydey    Added Customer Short Name to the search criteria.

       07-20-01	     DTy		When 'Active Only' selected:
								  Exclude customers that are marked as soft deleted.
                                  Exclude addresses that are marked as soft deleted.
                                  Exclude constacts that are:
                                    Marked as soft deleted in CONTACT.,
                                       i.e., RECORD_STATUS_IND='D'.
                                    Staff who left their employer.
                                       i.e., CONTACT.STAFF_STATUS_LCODE='Departed'.
       18-Feb-02	 DTy		Active customers are those whose status is either
                                  'Prospect', 'OnHold' or 'Current'.
       29-Mar-02	 DTy		Add "Customer ID" column.
								Facilitate 'Customer Cleanup' Customer ID and Name lookup.
	   14-Apr-02     DTy        Fix '>', '>>', '<', '<<' buttons by pasing the bolActiveOnly value.
	   09-Sep-12	ACheung		Add strServiceEnd == 'E' to handle the customer service lookup
***************************************************************************************************
-->
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->
<%
'check user's rights
     Function SimpleBinaryToString(Binary)
    
    Dim DecryptedData
    if (Binary <> "" and Binary <> Empty) then
    
SimpleBinaryToString = DecryptWithKey("Constant",Binary)
    
    
    else
     SimpleBinaryToString=""
   
    end if
   
End Function

 
if CheckLogon(strConst_Customer) = 0 then
	Response.Write "You don't have access to Customer. Please contact your system administrator."
	Response.End
end if
 
dim strCustomerName, strSMRLName, strSMRFName, strRegion, strStatus, bolActiveOnly
dim rsCustList,detailrsCustList, detailaList,aList
dim strSQL, strSelectClause, strFromClause, strWhereClause, strRecordStatus, strOrderBy
    dim detailstrSQL, detailstrSelectClause, detailstrFromClause, detailstrWhereClause, detailstrRecordStatus, detailstrOrderBy
dim intPageNumber, intPageCount
dim strMyWinName, strBgColor,strServiceEnd, strCustShort

'get search criteria
	strMyWinName = Request("hdnWinName")
	strCustomerName = UCase(routineOraString((trim(Request("txtCustomerName")))))

	strCustShort = UCase(routineOraString((trim(Request("txtCustShort")))))
	strSMRLName = UCase(routineOraString((trim(Request("txtSMRLName")))))
	strSMRFName = UCase(routineOraString((trim(Request("txtSMRFName")))))
	strRegion = Request("selRegion")
	strStatus = Request("selStatus")
	bolActiveOnly = Request("chkActiveOnly")
	strServiceEnd = Request("hdnServiceEnd")

	if strServiceEnd = "" then
	 strServiceEnd = "OTHER"
	END IF

'build query
'no criteria selected - display all
    
    



	strSelectClause = "select distinct " & _
				"t1.customer_id, " & _
				"t1.customer_name, " & _
				"t1.customer_short_name, " & _
				"t1.noc_region_lcode, " &_
				"t5.noc_region_desc, " &_
				"t1.customer_status_lcode, " & _
				"t3.last_name, " & _
				"t3.first_name, " & _
				"t3.contact_name, " & _
				"t4.street, " & _
				"t4.municipality_name, " & _
				"t4.province_state_lcode "

	strFromClause = " from " & _
				"crp.customer t1,  " &_
				"crp.customer_contact t2," &_
				"crp.contact t3," & _
				"crp.v_address_consolidated_street t4, " &_
				"crp.lcode_noc_region t5"

	strWhereClause = " where " & _
				"t1.customer_id = t2.customer_id(+) and " & _
				"t2.customer_contact_type_lcode(+)='custcare' and " & _
				"t2.contact_id = t3.contact_id(+) and " & _
				"t1.customer_id = t4.customer_id(+) and " & _
				"t4.primary_address_flag(+)= 'Y' and " &_
				"t1.noc_region_lcode = t5.noc_region_lcode "

	'customer name entered
	If strCustomerName <> "" then
		'include alias table
		strFromClause = strFromClause &  _
				", crp.customer_name_alias t0 "
		'join alias table to customer table and specify customer search string
		if len(strCustomerName) = 50 then 'max search length, do not append '%'
			strWhereClause = strWhereClause &  " and " & _
				"t0.customer_id = t1.customer_id and " & _
				"t0.customer_name_alias_upper like '" & (strCustomerName)& "'"
		else
			strWhereClause = strWhereClause &  " and " & _
				"t0.customer_id = t1.customer_id and " & _
				"t0.customer_name_alias_upper like '" & (strCustomerName) & "%'"
		end if
	End If


	If len(strCustShort) > 0 then
		strWhereClause = strWhereClause & " and " & _
			"Upper(t1.customer_short_name) like '" & (strCustShort) & "%'"
	End If

	'service mgnt rep entered
	If len(strSMRLName) > 0 then
		strWhereClause = strWhereClause & " and " & _
			"Upper(t3.last_name) like '" & (strSMRLName) & "%'"
	End If

	If len(strSMRFName) > 0 then
		strWhereClause = strWhereClause & " and " & _
			"Upper(t3.first_name) like '" & (strSMRFName) & "%'"
	End If

	'region picked
	If strRegion <> "All" then
		strWhereClause = strWhereClause & " and " & _
			"t1.noc_region_lcode = '" & strRegion & "'"
	End If

	'status picked
	if strStatus <> "All" then
		strWhereClause = strWhereClause & " and t1.customer_status_lcode = '" & strStatus & "'"
	end if

	'see all records?
	If bolActiveOnly = "yes" then
		strRecordStatus = " and t1.customer_status_lcode IN ('Prospect', 'Current', 'OnHold')" &_
		                  " and t1.record_status_ind = 'A' and t2.record_status_ind (+) = 'A'" & _
		                  " and t3.record_status_ind (+) = 'A' and t4.record_status_ind (+) = 'A' and " & _
		                  "(t3.staff_status_lcode is null or " & _
		                  "(t3.staff_status_lcode is not null and t3.staff_status_lcode <> 'Departed')) "
	Else 'no
		strRecordStatus = " "
	End If
   
	strOrderBy = " order by Upper(t1.customer_name)"

	'join all pieces to make a complete query
	strSQL = strSelectClause & strFromClause & strWhereClause  & strRecordStatus & strOrderBy
'	Response.Write strSQL & vbCrLf	'show SQL for debugging
  'response.end
	'get the recordset
	set rsCustList=server.CreateObject("ADODB.Recordset")
	rsCustList.Open strSQL, objConn
	If err then
		DisplayError "BACK", "", err.Number, "CustList.asp - Cannot open database" , err.Description
	End if

	'put recordset into array
	if not rsCustList.EOF then
		aList = rsCustList.GetRows
	else
		Response.Write "0 Records Found"
		Response.End
	end if

	'release and kill the recordset and the connection objects
	rsCustList.Close
	set rsCustList = nothing
	
	intPageCount = Int(UBound(aList,2) / intConstDisplayPageSize) + 1
     
    if Request("hdnExport") <> "" then
						'get real userid
						dim strRealUserID
						strRealUserID = Session("username")

						'determine export path
						dim strExportPath, liLength
						strExportPath = Request.ServerVariables("PATH_TRANSLATED")


						While (Right(strExportPath, 1) <> "\" And Len(strExportPath) <> 0)
							liLength = Len(strExportPath) - 1
							strExportPath = Left(strExportPath, liLength)
						Wend
						strExportPath = strExportPath & "export\"


						'create scripting object
						dim objFSO, objTxtStream
						set objFSO = server.CreateObject("Scripting.FileSystemObject")
						'create export file (overwrite if exists)
						set objTxtStream = objFSO.CreateTextFile(strExportPath&strRealUserID&"-customer.xls", true, false)

						if err then
							DisplayError "CLOSE", "", err.Number, "CustList.asp - Cannot create Excel spreadsheet file due to the following reasons. Please contact your website administrator.", err.Description
						end if

						with objTxtStream
							.WriteLine "<table border=1>"

							'export the header
							.WriteLine "<THEAD>"
							.WriteLine "<TH>Customer ID</TH>"
							.WriteLine "<TH>Customer Name</TH>"
							.WriteLine "<TH>Short Name</TH>"
							.WriteLine "<TH>Region</TH>"
							.WriteLine "<TH>Status</TH>"
							.WriteLine "<TH>Service Mgnt Rep</TH>"
							.WriteLine "<TH>Primary Address</TH>"
							.WriteLine "<TH>City</TH>"
							.WriteLine "<TH>Prov/State</TH>"
							.WriteLine "</THEAD>"

							'export the body
							for k = 0 to UBound(aList, 2)
								.WriteLine "<TR>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(0,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(1,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(2,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(3,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(5,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(8,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(9,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(10,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(11,k))&"&nbsp;</TD>"
								.WriteLine "</TR>"
							next
							.WriteLine "</table>"
						end with
						objTxtStream.Close	
						set objTxtStream = Nothing
						set objFSO = Nothing

						strsql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-customer.xls"";</script>"
						Response.Write strsql
						Response.End
'						Response.redirect "export/"&strRealUserID&"-customer.xls"
					if Request("txtGoToPageNo") <> "" then
						intPageNumber = CInt(Request("txtGoToPageNo"))
					else
						intPageNumber = 1
					end if
    end if
    
	if request("detailhdnExport") <>"" then
   

    

    detailstrSelectClause = "SELECT DISTINCT NULL FEDERATED_ASSET_ID, NE.NETWORK_ELEMENT_NAME DEVICE_NAME, NPNA.NETWORK_PORT_NAME_ALIAS ALIAS_NAME,"&_
    " NEP.NETWORK_ELEMENT_PORT_NAME PORT_NAME, LCI.CTR_IN_VALUE,     LCO.CTR_OUT_VALUE,  LVI.VTR_IN_VALUE,"&_
    "LVO.VTR_OUT_VALUE,       NEP.QOS_NAME QOS_NAME,        LEI.ETR_IN_VALUE,   "&_
    "     LEO.ETR_OUT_VALUE, LMSP.MGMT_SPACE_NAME MGMT_SPACE, NE.MANAGED_IP_ADDRESS MGMT_IP_ADDRESS, NEP.NETWORK_ELEMENT_PORT_IP CUST_IP_ADDRESS, "&_ 
    "NES.SNMP_STRING SNMP_STRING, NES.SNMP_V3_USERNAME SNMP_V3_USERNAME, NES.SNMP_V3_ENGINEID SNMP_V3_ENGINEID, " &_
    " NES.SNMP_V3_CONTEXT_NAME SNMP_V3_CONTEXT_NAME, LSSL.SNMP_SECURITY_LVL_NAME SNMP_SECURITY_LVL_NAME,"&_
    " LSAP.SNMP_AUTH_PROT_NAME SNMP_AUTH_PROT_NAME, LSPP.SNMP_PRIV_PROT_NAME SNMP_PRIV_PROT_NAME, NES.SNMP_V3_AUTH_KEY SNMP_V3_AUTH_KEY,"&_
    " NES.SNMP_V3_PRIV_KEY SNMP_V3_PRIV_KEY, NES.SNMP_PORT SNMP_PORT, LCS.CI_STATUS_NAME CI_STATUS,( select listagg(S.MGMT_SYSTEM_NAME,',')   within group (order by S.MGMT_SYSTEM_ID)  MGMT_SYSTEMS " &_
" from CRP.NETWORK_ELEMENT_MGMT_SYS NEMS, CRP.LCODE_MGMT_SYSTEMS S " &_
  "where NEMS.MGMT_SYSTEM_ID = S.MGMT_SYSTEM_ID and  NEMS.NETWORK_ELEMENT_Port_ID = nep.NETWORK_ELEMENT_Port_ID), MK.MAKE_DESC DEVICE_VENDOR, MDL.MODEL_DESC DEVICE_MODEL, LNPF.NE_PORT_FUNCTION_NAME, ST.SERVICE_TYPE_DESC SERVICE_TYPE, NULL REPORTING_PACKAGE, LTC.TENANT_NAME TENANT, C.CUSTOMER_NAME COMPANY_NAME, C.CUSTOMER_SHORT_NAME COMPANY_CODE, CO.ORGANIZATION_NAME ORGANIZATION_NAME, CO.ORGANIZATION_CODE ORGANIZATION_CODE, lc.COUNTRY_DESC COUNTRY, A.COUNTRY_LCODE COUNTRY_CODE, LPS.PROVINCE_STATE_NAME PROVINCE, A.PROVINCE_STATE_LCODE PROVINCE_CODE, A.MUNICIPALITY_NAME CITY, ML.CLLI_CODE CITY_CODE, SNC.SITE_NAME SITE, SNC.SITE_CODE SITE_CODE, NULL CUSTOM_1, NULL CUSTOM_2, NULL CUSTOM_3, NULL CUSTOM_4, NULL CUSTOM_5 , NEP.VN_NAME VN_NAME"

    detailstrFromClause ="  FROM CRP.NETWORK_ELEMENT ne, CRP.NETWORK_PORT_NAME_ALIAS NPNA, CRP.CUSTOMER c, CRP.NETWORK_ELEMENT_NAME_ALIAS nena, CRP.NETWORK_ELEMENT_PORT nep, CRP.LCODE_MGMT_SPACE lmsp, CRP.NETWORK_ELEMENT_SNMP nes, CRP.LCODE_CI_STATUS lcs,  CRP.ASSET_CATALOGUE ac, CRP.MAKE mk, CRP.MODEL mdl, CRP.MANAGED_CORRELATION mc, CRP.Customer_Service cs, CRP.SERVICE_TYPE st, CRP.LCODE_TENANT_CODE ltc, CRP.CUSTOMER_ORGANIZATION co, CRP.LCODE_COUNTRY lc, CRP.LCODE_PROVINCE_STATE lps, CRP.MUNICIPALITY_LOOKUP ml, CRP.ADDRESS a, CRP.SERVICE_LOCATION sl, CRP.SITE_NAME_CODE snc, CRP.LCODE_SNMP_SECURITY_LVL lssl, CRP.LCODE_SNMP_AUTH_PROT lsap, CRP.LCODE_SNMP_PRIV_PROT lspp, crp.customer_contact ccon, crp.contact con, crp.v_address_consolidated_street acs, crp.lcode_noc_region lnr , crp.customer_name_alias cna , crp.LCODE_CTR_IN LCI, crp.LCODE_CTR_OUT LCO,   crp.LCODE_ETR_IN LEI, crp.LCODE_ETR_OUT LEO,       crp.LCODE_VTR_IN LVI,       crp.LCODE_VTR_OUT LVO,   CRP.LCODE_NE_PORT_FUNCTION LNPF"

	

	detailstrWhereClause = "  WHERE ne.customer_id = c.customer_Id AND ne.Network_Element_Id = nena.Network_Element_Id(+) AND ne.Network_Element_Id = nep.Network_Element_Id(+) AND ne.MGMT_SPACE_ID = lmsp.MGMT_SPACE_ID(+) AND MC.CUSTOMER_SERVICE_ID =CS.CUSTOMER_SERVICE_ID and ne.Network_Element_Id = nes.Network_Element_Id(+) AND nep.CI_STATUS_ID = lcs.CI_STATUS_ID(+)  AND ne.Asset_Catalogue_Id = ac.Asset_Catalogue_Id(+) AND ac.MAKE_ID = mk.MAKE_ID(+) AND ac.MODEL_ID = mdl.MODEL_ID(+) AND NE.NETWORK_ELEMENT_ID = MC.NETWORK_ELEMENT_ID and cs.SERVICE_TYPE_ID = st.SERVICE_TYPE_ID AND ne.TENANT_ID = ltc.TENANT_ID(+) AND nep.ORGANIZATION_ID = co.ORGANIZATION_ID(+) AND CS.SERVICE_LOCATION_ID = SL.SERVICE_LOCATION_ID AND A.ADDRESS_ID = SL.ADDRESS_ID AND A.COUNTRY_LCODE = LC.COUNTRY_LCODE(+) AND A.PROVINCE_STATE_LCODE = LPS.PROVINCE_STATE_LCODE AND A.MUNICIPALITY_NAME = ML.MUNICIPALITY_NAME(+) AND NEP.Site_id = SNC.Site_id(+) AND c.customer_id = ccon.customer_id(+) AND ccon.customer_contact_type_lcode(+) = 'custcare' AND ccon.contact_id = con.contact_id(+) AND c.customer_id = acs.customer_id(+) AND acs.primary_address_flag(+) = 'Y' AND c.noc_region_lcode = lnr.noc_region_lcode AND cna.customer_id = c.customer_id AND NES.SNMP_SECURITY_LVL_ID = LSSL.SNMP_SECURITY_LVL_ID (+) AND NES.SNMP_AUTH_PROT_ID = LSAP.SNMP_AUTH_PROT_ID (+) AND NES.SNMP_PRIV_PROT_ID = lspp.SNMP_PRIV_PROT_ID (+) and  NEP.CTR_IN_ID =  lci.CTR_IN_ID(+)       and NEP.CTR_OUT_ID= LCO.CTR_OUT_ID(+)       and NEP.ETR_IN_ID=LEI.ETR_IN_ID(+)       and NEP.ETR_OUT_ID= LEO.ETR_OUT_ID(+)       and NEP.VTR_IN_ID = LVI.VTR_IN_ID(+)       and NEP.VTR_OUT_ID = LVO.VTR_OUT_ID(+) and ML.COUNTRY_LCODE = A.COUNTRY_LCODE and ML.PROVINCE_STATE_LCODE = A.PROVINCE_STATE_LCODE and NEP.NETWORK_ELEMENT_PORT_ID = NPNA.NETWORK_ELEMENT_PORT_ID(+) AND NEP.NE_PORT_FUNCTION_LCODE = LNPF.NE_PORT_FUNCTION_LCODE(+)" 
   
    if Request("_txtStartDate") <> "" then
    detailstrWhereClause = detailstrWhereClause + " AND TRUNC( NEP.update_date_time) >= to_date('"+ Request("_txtStartDate")  + "','DD/MM/YYYY')   "
    'AND TRUNC( c.create_date_time) >    TO_DATE ('25/07/2011', 'DD/MM/YYYY')
    end if

     if Request("_txtEndDate") <> "" then
    detailstrWhereClause = detailstrWhereClause + " AND TRUNC( NEP.update_date_time) <= to_date('"+ Request("_txtEndDate")  + "','DD/MM/YYYY')   "
    'AND TRUNC( c.create_date_time) >    TO_DATE ('25/07/2011', 'DD/MM/YYYY')
    end if

     If strCustomerName <> "" then
		'include alias table
		'detailstrFromClause = detailstrFromClause &  _
				'", crp.customer_name_alias t0 "
		'join alias table to customer table and specify customer search string
		if len(strCustomerName) = 50 then 'max search length, do not append '%'
			detailstrWhereClause = detailstrWhereClause &  " and " & "cna.customer_name_alias_upper like '" & (strCustomerName)& "'"
		else
			detailstrWhereClause = detailstrWhereClause &  " and " & "cna.customer_name_alias_upper like '" & (strCustomerName) & "%'"
		end if
	End If


	If len(strCustShort) > 0 then
		detailstrWhereClause = detailstrWhereClause & " and " & _
			"Upper(C.customer_short_name) like '" & (strCustShort) & "%'"
	End If




	'service mgnt rep entered
	If len(strSMRLName) > 0 then
    
		detailstrWhereClause = detailstrWhereClause & " and " & _
			"Upper(con.last_name) like '" & (strSMRLName) & "%'"
	End If

	If len(strSMRFName) > 0 then

		detailstrWhereClause = detailstrWhereClause & " and " & _
			"Upper(con.first_name) like '" & (strSMRFName) & "%'"
	End If

	'region picked
	If strRegion <> "All" then
		detailstrWhereClause = detailstrWhereClause & " and " & _
			"C.noc_region_lcode = '" & strRegion & "'"
	End If

	'status picked
	if strStatus <> "All" then
		detailstrWhereClause = detailstrWhereClause & " and C.customer_status_lcode = '" & strStatus & "'"
	end if
    'see all records?
	If bolActiveOnly = "yes" then

   ' detailstrFromClause = detailstrFromClause & " ,crp.customer_contact t2,crp.v_address_consolidated_street t4 "

   ' detailstrWhereClause = detailstrWhereClause &  " and con.customer_id ="
		strRecordStatus = " and c.customer_status_lcode IN ('Prospect', 'Current', 'OnHold') " &_
		                " and c.record_status_ind = 'A'  and con.record_status_ind (+) = 'A'" & _
		                " and con.record_status_ind (+) = 'A' and acs.record_status_ind (+) = 'A'  " & _
		                 " and (con.staff_status_lcode is null or " & _
		                  "(con.staff_status_lcode is not null and con.staff_status_lcode <> 'Departed')) "
	Else 'no'
		strRecordStatus = " "
	End If
	
	
	strOrderBy = " order by Upper(C.customer_name)"
	'see all records?
	
	'join all pieces to make a complete query
	detailstrSQL = detailstrSelectClause & detailstrFromClause & detailstrWhereClause  & strRecordStatus '& strOrderBy
    
    
'	Response.Write strSQL & vbCrLf	'show SQL for debugging
  'response.end
	'get the recordset
	set detailrsCustList=server.CreateObject("ADODB.Recordset")
	detailrsCustList.Open detailstrSQL, objConn
	'release and kill the recordset and the connection objects
    
 	
If err then
		DisplayError "BACK", "", err.Number, "CustList.asp - Cannot open database" , err.Description
	End if

	'put recordset into array
     
	if not detailrsCustList.EOF then
		detailaList = detailrsCustList.GetRows
	else
		Response.Write "0 Records Found"
		Response.End
	end if

	'release and kill the recordset and the connection objects
	detailrsCustList.Close
	set detailrsCustList = nothing
    objConn.Close
	set objConn = nothing

	'calculate page number
	intPageCount = Int(UBound(detailaList,2) / intConstDisplayPageSize) + 1

                    dim detailstrRealUserID
						detailstrRealUserID = Session("username")

						'determine export path
						dim detailstrExportPath, detailliLength
						detailstrExportPath = Request.ServerVariables("PATH_TRANSLATED")


						While (Right(detailstrExportPath, 1) <> "\" And Len(detailstrExportPath) <> 0)
							detailliLength = Len(detailstrExportPath) - 1
							detailstrExportPath = Left(detailstrExportPath, detailliLength)
						Wend
						detailstrExportPath = detailstrExportPath & "export\"

						'create scripting object
						dim detailobjFSO, detailobjTxtStream
						set detailobjFSO = server.CreateObject("Scripting.FileSystemObject")
						'create export file (overwrite if exists)
					 set detailobjTxtStream = detailobjFSO.CreateTextFile(detailstrExportPath&detailstrRealUserID&"-detailcustomer.xls", true, false)

						if err then
							DisplayError "CLOSE", "", err.Number, "CustList.asp - Cannot create Excel spreadsheet file due to the following reasons. Please contact your website administrator.", err.Description
						end if
    
   						with detailobjTxtStream
							.WriteLine "<table border=1>"

							'export the header

							.WriteLine "<THEAD>"
    .WriteLine "<TH>FEDERATED_ASSET_ID</TH> " 
    .WriteLine "<TH>DEVICE_NAME</TH> " 
    .WriteLine "<TH>CUST_DEVICE_NAME</TH> "
     .WriteLine "<TH>SNC_DEVICE_NAME</TH> "
    .WriteLine "<TH>ALIAS_NAME</TH> " 
    .WriteLine "<TH>PORT_NAME</TH>"
    .WriteLine "<TH>CTR_IN</TH>"
    .WriteLine "<TH>CTR_Out</TH>"
    .WriteLine "<TH>VN_NAME</TH>"
.WriteLine "<TH>VTR_IN</TH>"
    .WriteLine "<TH>VTR_OUT</TH>"
    .WriteLine "<TH>QOS_NAME</TH>"
    .WriteLine "<TH>ETR_IN</TH>"
.WriteLine "<TH>ETR_OUT</TH>"
    .WriteLine "<TH>MGMT_SPACE</TH>"
    .WriteLine "<TH>MGMT_IP_ADDRESS</TH>"
    .WriteLine "<TH>CUST_IP_ADDRESS</TH>"
    .WriteLine "<TH>SNMP_STRING</TH>"
							.WriteLine "<TH>SNMP_V3_USERNAME</TH>"
							.WriteLine "<TH>SNMP_V3_ENGINEID</TH>"
							.WriteLine "<TH>SNMP_V3_CONTEXT_NAME</TH>"
    .WriteLine "<TH>SNMP_V3_SEC_LEVEL</TH>"
							.WriteLine "<TH>SNMP_V3_AUTH_PROTOCOL</TH>"
							.WriteLine "<TH>SNMP_V3_PRIV_PROTOCOL</TH>"
							.WriteLine "<TH>SNMP_V3_AUTH_KEY</TH>"
    .WriteLine "<TH>SNMP_V3_PRIV_KEY</TH>"
							.WriteLine "<TH>SNMP_PORT</TH>"
							.WriteLine "<TH>CI_STATUS</TH>"
                             .Writeline "<TH>MGMT_SYSTEM</TH>"
                            .WriteLine "<TH>DEVICE_VENDOR</TH>"
							.WriteLine "<TH>DEVICE_MODEL</TH>"
							
							.WriteLine "<TH>TECHNOLOGY</TH>"
                           .Writeline "<TH>SERVICE_TYPE</TH>"
                           .Writeline "<TH>REPORTING_PACKAGE</TH>"
							.WriteLine "<TH>TENANT</TH>"
                            .WriteLine "<TH>COMPANY_NAME</TH>"
                             .WriteLine "<TH>COMPANY_CODE</TH>"
                             .WriteLine "<TH>ORGANIZATION_NAME</TH>"
                             .WriteLine "<TH>ORGANIZATION_CODE</TH>"
                             .WriteLine "<TH>COUNTRY</TH>"
                             .WriteLine "<TH>COUNTRY_CODE</TH>"
                               .WriteLine "<TH>PROVINCE</TH>"
                               .WriteLine "<TH>PROVINCE_CODE</TH>"
                               .WriteLine "<TH>CITY</TH>"
                             .WriteLine "<TH>CITY_CODE</TH>"
                             .WriteLine "<TH>SITE</TH>"
                             .WriteLine "<TH>SITE_CODE</TH>"
                               .WriteLine "<TH>CUSTOM_1</TH>"
                                .WriteLine "<TH>CUSTOM_2</TH>"
                                 .WriteLine "<TH>CUSTOM_3</TH>"
                                .WriteLine "<TH>CUSTOM_4</TH>"
                                 .WriteLine "<TH>CUSTOM_5</TH>"
                             .WriteLine "</THEAD>"

							'export the body
							for k = 0 to UBound(detailaList, 2)
								.WriteLine "<TR>"
	.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(0,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(1,k))&"&nbsp;</TD>"
								 .WriteLine "<TD NOWRAP>&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(2,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(3,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(4,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(5,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(49,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(6,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(7,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(8,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(9,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(10,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(11,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(12,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(13,k))&"&nbsp;</TD>"	
    .WriteLine "<TD NOWRAP>"& SimpleBinaryToString(routineHtmlString(detailaList(14,k))) &"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&  SimpleBinaryToString(routineHtmlString(detailaList(15,k)))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"& SimpleBinaryToString(routineHtmlString(detailaList(16,k)))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"& SimpleBinaryToString(routineHtmlString(detailaList(17,k)))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(18,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(19,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(20,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"& SimpleBinaryToString(routineHtmlString(detailaList(21,k)))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"& SimpleBinaryToString(routineHtmlString(detailaList(22,k)))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(23,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(24,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(25,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(26,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(27,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(28,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(29,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(30,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(31,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(32,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(33,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(34,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(35,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(36,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(37,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(38,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(39,k))&"&nbsp;</TD>"
        .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(40,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(41,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(42,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(43,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(44,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(45,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(46,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(47,k))&"&nbsp;</TD>"
    .WriteLine "<TD NOWRAP>"&routineHtmlString(detailaList(48,k))&"&nbsp;</TD>"
								.WriteLine "</TR>"
							next
							.WriteLine "</table>"
						end with
						detailobjTxtStream.Close	
    
						set detailobjTxtStream = Nothing
						set detailobjFSO = Nothing

						detailstrsql = "<script type=""text/javascript"">document.location=""export/"&detailstrRealUserID&"-detailcustomer.xls"";</script>"
						Response.Write detailstrsql
						Response.End
'						Response.redirect "export/"&strRealUserID&"-customer.xls"
					
   
  end if
		
	select case Request("Action")
   
		case "<<"	intPageNumber = 1
		case "<"	intPageNumber = Request("txtPageNumber")-1
					if intPageNumber < 1 then intPageNumber = 1
		case ">"	intPageNumber = Request("txtPageNumber") + 1
					if intPageNumber > intPageCount then intPageNumber = intPageCount
		case ">>"	intPageNumber=intPageCount
		case  else				
	end select

	if intPageNumber < 1 then intPageNumber = 1
	if intPageNumber > intPageCount then intPageNumber = intPageCount

	dim k,m,n
	m = (intPageNumber - 1) * intConstDisplayPageSize
	n = (intPageNumber) * intConstDisplayPageSize - 1
	if n > UBound(aList,2) then
		n=UBound(aList,2)
	end if

	'check if the client is still connected just before sending any html to the browser
	if Response.IsClientConnected = false then
		Response.End
	end if

	'catch any unexpected error
	if err then
		DisplayError "BACK", "", err.Number, "Unexpected error.", err.Description
	end if
%>

<html>
<head>
    <style>
        /* The Modal (background) */
        .modal {
            display: none; /* Hidden by default */
            position: fixed; /* Stay in place */
            z-index: 1; /* Sit on top */
            padding-top: 100px; /* Location of the box */
            left: 0;
            top: 0;
            width: 100%; /* Full width */
            height: 100%; /* Full height */
            overflow: auto; /* Enable scroll if needed */
            background-color: rgb(0,0,0); /* Fallback color */
            background-color: rgba(0,0,0,0.4); /* Black w/ opacity */
            filter: alpha(opacity=50);
        }

        /* Modal Content */
        .modal-content {
            background-color: #fefefe;
            margin: auto;
            padding: 20px;
            border: 1px solid #888;
            width: 50%;
        }

        /* The Close Button */
        .close {
            color: #aaaaaa;
            float: right;
            font-size: 28px;
            font-weight: bold;
        }

            .close:hover,
            .close:focus {
                color: #000;
                text-decoration: none;
                cursor: pointer;
            }
    </style>
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link rel="stylesheet" type="text/css" href="stylesheets/styles.css" type="text/css">
    <title>Service Management Application</title>
    <script type="text/javascript" src="GeneralJavaFunctions.js"></script>

    <script id="clientEventHandlersJS" type="text/javascript">
	<!--
    setPageTitle("SMA - Customer");

    function go_back(strServiceEnd, lngCustomerID, strCustomerName, strCustomerShortName, strRegion) {
        //Response.Write ("inside go_back strServiceEnd is " & strServiceEnd)
        //alert (strServiceEnd);

        try {
            if (strServiceEnd == 'A') {
                parent.opener.document.forms[0].hdnCustomerIdA.value = lngCustomerID;
                parent.opener.document.forms[0].txtcustomera.value = strCustomerName;
            }
            else if (strServiceEnd == 'B') {
                parent.opener.document.forms[0].hdnCustomerIdB.value = lngCustomerID;
                parent.opener.document.forms[0].txtcustomerb.value = strCustomerName;
            }
            else if (strServiceEnd == 'C') { //this condition handles the customer service lookup
                parent.opener.document.forms[0].hdnCustomerID.value = lngCustomerID;
                parent.opener.document.forms[0].txtCustomerName.value = strCustomerName;
            }
            else if (strServiceEnd == 'D') { // Region is returned to CustServDetail.asp
                parent.opener.document.forms[0].hdnCustomerID.value = lngCustomerID;
                parent.opener.document.forms[0].txtCustomerName.value = strCustomerName;
                parent.opener.document.forms[0].txtRegion.value = strRegion;
                parent.opener.document.forms[0].txtCustomerShortName.value = strCustomerShortName;
            }
            else if (strServiceEnd == 'F') { // this condition handles FR Customer in CustCleanEntry.asp
                parent.opener.document.forms[0].txtFRCustomer.value = "(" + lngCustomerID + ") " + strCustomerName;
                parent.opener.document.forms[0].hdnFRCustomerID.value = lngCustomerID;
                parent.opener.document.forms[0].hhnFRCustomerName.value = strCustomerName;
            }
            else if (strServiceEnd == 'T') { // this condition handles TO Customer in CustCleanEntry.asp
                parent.opener.document.forms[0].txtTOCustomer.value = "(" + lngCustomerID + ") " + strCustomerName;
                parent.opener.document.forms[0].hdnTOCustomerID.value = lngCustomerID;
                parent.opener.document.forms[0].hdnTOCustomerName.value = strCustomerName;
            }
            else if (strServiceEnd == 'X') { // this condition handles Customer in XLSEntry.asp
                parent.opener.document.forms[0].txtCustomer.value = "(" + lngCustomerID + ") " + strCustomerName;
                parent.opener.document.forms[0].hdnCustomerID.value = lngCustomerID;
                parent.opener.document.forms[0].hdnCustomerName.value = strCustomerName;
            }
            else if (strServiceEnd == 'E') { //this condition handles the customer service lookup
                //alert (strCustomerName);
                parent.opener.document.forms[0].txtCustomerID.value = lngCustomerID;
                parent.opener.document.forms[0].txtCustomerName.value = strCustomerName;
                parent.opener.document.forms[0].txtCustomerShortName.value = strCustomerShortName;
            }
            else {
                parent.opener.document.forms[0].hdnCustomerID.value = lngCustomerID;
                parent.opener.document.forms[0].txtCustomerName.value = strCustomerName;
                parent.opener.document.forms[0].txtCustomerShortName.value = strCustomerShortName;
            }
        }
        catch (e) {
            //do nothing, most probably not all forms have CustomerShortName - needed in Managed Objects Details
        }
        parent.window.close();
    }

    function onExport() {

        // Get the modal
        var modal = document.getElementById('myModal');

        modal.style.display = "block";
        document.getElementsByName("_txtStartDate")[0].value = "";
        document.getElementsByName("_txtEndDate")[0].value = "";
        // Get the <span> element that closes the modal
        var span = document.getElementsByClassName("close")[0];
        // When the user clicks on <span> (x), close the modal
        span.onclick = function () {
            var modal = document.getElementById('myModal');
            modal.style.display = "none";
        }
       
        return false;
       
    }
    //-->

    function OnExportFilter() {
        document.frmCustList.target = 'new';
        var expElement = document.getElementsByName("detailhdnExport");
        expElement[0].value = 'xls';
        document.frmCustList.submit();
       // document.frmCustList.detailhdnexport.value = '';
       // document.frmCustList.target = '_self';

    }




    // When the user clicks anywhere outside of the modal, close it
    window.onclick = function (event) {
        var modal = document.getElementById('myModal');
        if (event.target == modal) {
            modal.style.display = "none";
        }
    }
    </script>

</head>
<body>
    <form name="frmCustList" action="CustList.asp" method="POST">
        <input type="hidden" name="hdnDate" value="">
        <input type="hidden" name="txtCustomerName" value="<%=strCustomerName%>">
        <input type="hidden" name="txtCustShort" value="<%=strCustShort%>">
        <input type="hidden" name="txtSMRLName" value="<%=strSMRLName%>">
        <input type="hidden" name="txtSMRFName" value="<%=strSMRFName%>">
        <input type="hidden" name="selRegion" value="<%=strRegion%>">
        <input type="hidden" name="selStatus" value="<%=strStatus%>">
        <input type="hidden" name="hdnServiceEnd" value="<%=strServiceEnd%>">
        <input type="hidden" name="hdnExport" value>
        <input type="hidden" name="detailhdnExport" value>
        <input type="hidden" name="chkActiveOnly" value="<%=bolActiveOnly%>">
        <table border="1" cellpadding="2" cellspacing="0" width="100%">
            <thead>
                <tr>
                    <th align="left" nowrap>Customer ID</th>
                    <th align="left" nowrap>Customer Name</th>
                    <th align="left" nowrap>Short Name</th>
                    <th align="left" nowrap>Region</th>
                    <th align="left" nowrap>Status</th>
                    <th align="left" nowrap>Service Mgnt Rep</th>
                    <th align="left" nowrap>Primary Address</th>
                    <th align="left" nowrap>City</th>
                    <th align="left" nowrap>Prov/State</th>

                </tr>
            </thead>
            <tbody>
                <%
'display the table
	for k = m to n
		'alternate row background color
		if Int(k/2) = k/2 then
			Response.Write "<tr bgcolor=White>"
		else
			Response.Write "<tr>"
		end if

		if strMyWinName = "Popup" then
			Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" &strServiceEnd& "','"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(2,k)) & "','" &routineJavascriptString(aList(4,k))& "')"">" &aList(0,k)& "</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" &strServiceEnd& "','"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(2,k)) & "','" &routineJavascriptString(aList(4,k))& "')"">" &aList(1,k)& "</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" &strServiceEnd& "','"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(2,k)) & "','" &routineJavascriptString(aList(4,k))& "')"">" &aList(2,k)& "</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" &strServiceEnd& "','"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(2,k)) & "','" &routineJavascriptString(aList(4,k))& "')"">" &aList(3,k)& "</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" &strServiceEnd& "','"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(2,k)) & "','" &routineJavascriptString(aList(4,k))& "')"">" &aList(5,k)& "</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" &strServiceEnd& "','"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(2,k)) & "','" &routineJavascriptString(aList(4,k))& "')"">" &aList(8,k)& "</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" &strServiceEnd& "','"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(2,k)) & "','" &routineJavascriptString(aList(4,k))& "')"">" &aList(9,k)& "</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" &strServiceEnd& "','"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(2,k)) & "','" &routineJavascriptString(aList(4,k))& "')"">" &aList(10,k)& "</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" &strServiceEnd& "','"&aList(0,k)& "', '" &routineJavascriptString(aList(1,k))& "','" &routineJavascriptString(aList(2,k)) & "','" &routineJavascriptString(aList(4,k))& "')"">" &aList(11,k)& "</a>&nbsp;</td>"&vbCrLf
			Response.Write "</tr>"
		else
			Response.Write "<td nowrap><a target=""_parent"" href=""CustDetail.asp?CustomerID="&aList(0,k)&""">"&aList(0,k)&"</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""CustDetail.asp?CustomerID="&aList(0,k)&""">"&aList(1,k)&"</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""CustDetail.asp?CustomerID="&aList(0,k)&""">"&aList(2,k)&"</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""CustDetail.asp?CustomerID="&aList(0,k)&""">"&aList(3,k)&"</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""CustDetail.asp?CustomerID="&aList(0,k)&""">"&aList(5,k)&"</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""CustDetail.asp?CustomerID="&aList(0,k)&""">"&aList(8,k)&"</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""CustDetail.asp?CustomerID="&aList(0,k)&""">"&aList(9,k)&"</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""CustDetail.asp?CustomerID="&aList(0,k)&""">"&aList(10,k)&"</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""CustDetail.asp?CustomerID="&aList(0,k)&""">"&aList(11,k)&"</a>&nbsp;</td>"&vbCrLf
			Response.Write "</tr>"
		end if
   next
                %>
            </tbody>
            <tfoot>
                <tr>
                    <td align="left" colspan="8">
                        <input type="hidden" name="hdnWinName" value="<%=strMyWinName%>">
                        <input type="hidden" name="txtPageNumber" value="<%=intPageNumber%>">

                        <input type="submit" name="action" value="&lt;&lt;">
                        <input type="submit" name="action" value="&lt;">
                        <input type="text" name="txtGoToPageNo" onclick="document.frmCustList.txtGoToPageNo.value = ''" title="You can jump to a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="height: 22px; width: 150px">
                        <input type="submit" name="action" value="&gt;">
                        <input type="submit" name="action" value="&gt;&gt;">&nbsp;&nbsp;
	<img src="images/excel.gif" onclick="document.frmCustList.target='new';document.frmCustList.hdnExport.value='xls';document.frmCustList.submit();document.frmCustList.hdnExport.value='';document.frmCustList.target='_self';" width="32" height="32">
                        <span>Customer List Report</span>
                        <img src="images/excel.gif" style="padding-left: 400px;" onclick="onExport()" width="32" height="32">
                        <span>OSS CPE configuration extract</span>

                    </td>
                </tr>
            </tfoot>
            <caption>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></caption>
        </table>
        <div id="myModal" class="modal">

            <!-- Modal content -->
            <div class="modal-content">
                <span class="close">×</span>
                <div>Please enter the below dates in DD/MM/YYYY formats</div>
                <div>
                    Start Date :
                    <input type="text" name="_txtStartDate" />
                    End Date : 
                    <input type="text" name="_txtEndDate" />
                </div>
                <button onclick="OnExportFilter()">Submit</button>
            </div>

        </div>
    </form>
</body>
</html>








