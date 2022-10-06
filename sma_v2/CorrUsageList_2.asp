<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Buffer = true %>
<!-- #include file="../smaConstants.inc" -->
<!-- #include file="../smaProcs.inc" -->
<%
Const ASP_NAME = "CorrUsageList.asp" 'only need to change this value when changing the filename

dim strCustServID, objRsServiceContact, strServTypeID, StrSql
Dim intAccessLevel, bolActiveOnly

strServTypeID = Request("ServiceTypeID")
strCustServID = Request("CustomerServiceID")

if strCustServID <> "" then
  strSQL =  "select sd.SO_DETAIL_ID," &_ 
			"     se.SO_ELEMENT_DESCRIPTION," &_
			"     sde.DETAIL_ELEMENT_TEXT," &_
			"     SE.ELEMENT_TYPE_LCODE," &_
			"     sde.SO_ELEMENT_ID," &_
			"     se.SO_ELEMENT_SMA_DISPLAY," &_
			"     st.SERVICE_TYPE_ID," &_
			"     sc.SERVICE_CONCENTRATION_NAME," &_
			"     sb.SERVICE_BASE_NAME" &_
			"from so.so_detail sd," &_
			"     so.so_detail_element sde," &_
	                "     so.so_element se," &_
                        "     crp.service_type st," &_
                        "     crp.service_con_base_xref scbx," &_
                        "     crp.lcode_service_concentration sc," &_
			"     crp.lcode_service_base sb" &_
			"where sd.SO_DETAIL_ID = sde.SO_DETAIL_ID," &_
			"and   sde.SO_ELEMENT_ID = se.SO_ELEMENT_ID," &_
			"and   st.service_type_id = scbx.service_type_id," &_
			"and   scbx.SERVICE_CONCENTRATION_LCODE = sc.SERVICE_CONCENTRATION_LCODE," &_
			"and   scbx.SERVICE_BASE_LCODE = sb.SERVICE_BASE_LCODE," &_
			"and   se.SO_ELEMENT_SMA_DISPLAY = 'Y' " &_
			"and   st.service_type_id = "& strServTypeID
			"and   sd.CUSTOMER_SERVICE_ID = "& strCustServID
	
	set objRsServiceOrderInfo = objConn.Execute(StrSql)
	if  objRsServiceOrderInfo.EOF then
		strCustServID = ""
	end if

	if err then
		DisplayError "BACK", "", err.Number, "CANNOT CREATE RECORDSET 32132", err.Description
	end if

end if

%>
