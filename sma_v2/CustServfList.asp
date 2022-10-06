<%@ Language=VBScript %>
<% option explicit
   'on error resume next
 %>
 <% Response.Buffer = true %>
<!--#include file="SmaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->

<!--
***************************************************************************************************
* Name:		CustServList.asp i.e. Customer Service List
*
* Purpose:	This page reads users's search critiera and bring back a list of matching Customer
*			Service records.
*
* Created By:	Sara Sangha 08/01/00
* Edited by:    Adam Haydey 01/25/01
*               Added Customer Service City and  Customer Service Address search fields.
***************************************************************************************************
		 Date		Author			Changes/enhancements made
		06-Mar-01	 DTy		Save 'ActiveOnly' cookie for use by CustServContList.asp.
		20-Jul-01	 DTy		When 'Active Only' is selected:
		                          Exclude Service Locations that are marked as soft deleted.
		                          Exclude Customers that are marked as soft deleted.
		                          Exclude Addresses that are marked as soft deleted.
       18-Feb-02	 DTy		Active customers are those whose status is either
                                  'Prospect', 'OnHold' or 'Current'.
       28-Feb-02	 DTy		Include Customer Service Desc Alias when searching for Customer
                                  Service names.
       26-Oct-03     DTy        Add Customer Service selection from ManObjPortDetail.asp
	   13-Sept-04	  MW    	Add Lynx default severity as search fields.
       10-Aug-12    ACheung		Add Customer ID and Customer Shortname
***************************************************************************************************
-->

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
<script type="text/javascript">

// End of script hiding -->
</script>
</HEAD>
 <%

    dim aList, intPageNumber, intPageCount
    dim strCustomerServiceDesc,  strServiceLocationName
    dim strStatusCode, intCustomerServiceID, strServiceType
    dim bolActiveOnly
    dim strSQL, strWhereClause, strRecordStatus,strOrderBy
    dim  intCustomerID, strCustomerShortName
    dim color
    Dim objRsResult

    strCustomerServiceDesc = UCase(trim(Request.Form("txtCustomerServiceDesc")))
	strServiceLocationName = UCase(trim(Request.Form("txtServiceLocationName")))
	strStatusCode = trim(Request.Form("SelStatus"))
	intCustomerServiceID = trim(Request.Form("txtCustomerServiceID"))
	strServiceType = UCase(trim(Request.Form("txtServiceType")))
	bolActiveOnly = trim(Request.Form("chkActiveOnly"))
	intCustomerID = trim(Request.Form("txtCustomerID"))
	strCustomerShortName = UCase(trim(Request.Form("txtCustomerShortName")))

	strSQL = "select corr.root_CUSTOMER_SERVICE_ID bcsid, CS.CUSTOMER_SERVICE_DESC bserv_desc, "&_
	         "ST.SERVICE_TYPE_DESC, L.SERVICE_LOCATION_NAME, CS.SERVICE_STATUS_CODE, " &_
             " corr.CUSTOMER_SERVICE_ID fcsid, FCS.CUSTOMER_SERVICE_DESC fcs_desc, FCS.SERVICE_STATUS_CODE " &_
        	 " from  CRP.MANAGED_CORRELATION corr," &_
         	 " CRP.CUSTOMER cust "&_
         	 ",crp.customer_service cs, CRP.SERVICE_TYPE st, CRP.SERVICE_LOCATION l "&_
        	 ", CRP.CUSTOMER_SERVICE fcs, CRP.SERVICE_TYPE fst " &_
			 " where CS.CUSTOMER_SERVICE_ID = CORR.ROOT_CUSTOMER_SERVICE_ID "&_
			 "and     ST.SERVICE_TYPE_ID =  CS.SERVICE_TYPE_ID "&_
			 "and     L.SERVICE_LOCATION_ID = CS.SERVICE_LOCATION_ID "&_
			 "and    FCS.CUSTOMER_SERVICE_ID=CORR.CUSTOMER_SERVICE_ID "&_
			 "and    FST.SERVICE_TYPE_ID=FCS.SERVICE_TYPE_ID "&_
			 "and    FST.SERVICE_TYPE_DESC IN ('WAN L3 VPN Features', 'Cooperators LAN Features') " &_
			 "and    CS.CUSTOMER_ID = CUST.CUSTOMER_ID "

	IF (len(intCustomerID)<> 0 ) then
	   strWhereClause = strWhereClause + "and    CUST.CUSTOMER_ID = " & intCustomerID
	end if
	if (len(strCustomerShortName) <> 0) then
	    strWhereClause = strWhereClause + " AND Upper(cust.customer_short_name)  LIKE '%" & routineOraString(strCustomerShortName) & "%' "
	end if

    if (len(intCustomerServiceID) <>0) then
        strWhereClause = strWhereClause + " and  CS.CUSTOMER_SERVICE_ID = " & intCustomerServiceID
    end if

	'add other search parameters to the where clause
	IF LEN(strCustomerServiceDesc) > 0 THEN
	  strWhereClause = strWhereClause & " AND cs.customer_service_id in (" &_
		            " select customer_service_id from crp.customer_service where " & rtRmvSpChr("customer_service_desc", "Y") & " like '%" & rtRmvSpChr(strCustomerServiceDesc, "N") & "%' union" &_
                    " select customer_service_id from crp.customer_service_desc_alias where " & rtRmvSpChr("customer_service_desc_alias", "Y") & " like '%" & rtRmvSpChr(strCustomerServiceDesc, "N") & "%')"

	END IF

	IF LEN(strServiceLocationName) > 0 THEN
      strWhereClause = strWhereClause & " AND UPPER(l.service_location_name) LIKE upper('%" & routineOraString(strServiceLocationName) &"%')"
	END IF

	IF LEN(strStatusCode) > 0 THEN
		if strStatusCode = "AllExceptTerm" then
			strWhereClause = strWhereClause & " AND cs.service_status_code <> 'TERM' AND fcs.service_status_code <> 'TERM'"
		else
			strWhereClause = strWhereClause & " AND fcs.service_status_code <> 'TERM' AND cs.service_status_code = '" & routineOraString(strStatusCode) & "'"
		end if
    END IF

	IF  LEN(strServiceType) > 0 THEN
      strWhereClause = strWhereClause & " AND Upper(st.service_type_desc)  LIKE upper('%" & routineOraString(strServiceType) & "%') "
	END IF

    'Response.Cookies ("ActiveOnly")=bolActiveOnly

	'if bolActiveOnly = "YES" then
		strRecordStatus = " and corr.record_status_ind (+) = 'A'" &_
						  " and st.record_status_ind (+) = 'A'" &_
						  " and fcs.record_status_ind (+) = 'A'" &_
						  " and fst.record_status_ind (+) = 'A'" &_
		                  " and cust.customer_status_lcode IN ('Prospect', 'Current', 'OnHold')" &_
                          " and cs.record_status_ind (+) = 'A' and l.record_status_ind (+) = 'A'" & _
		                  " and cust.record_status_ind = 'A'  "

	'else
		'display all record
	'	strRecordStatus = " "
	'end if

	strOrderBy = " order by upper(cs.customer_service_desc)"

	'join all pieces to make a complete query
	strsql = strSQL & strWhereClause & strRecordStatus & strOrderBy

 ' 	Response.Write( strsql )       'display SQL for debugging
 '	Response.end

	set objRsResult = objConn.Execute(strSql)
	if not objRsResult.EOF then
		aList = objRsResult.GetRows
	else
		Response.Write "0 records found"
		Response.end
	end if

   'release and kill the recordset and the connection objects
	objRsResult.Close
	set objRsResult = nothing

	objConn.close
	set objConn = nothing

   'calculate page number
	intPageCount = Int(UBound(aList, 2) / intConstDisplayPageSize) + 1

	select case Request("Action")

		case "<<"		intPageNumber = 1
		case "<"		color = Request("txtcolor")
		                intPageNumber = Request("txtPageNumber") - 1
					    if intPageNumber < 1 then intPageNumber = 1
		case ">"		color = Request("txtcolor")
						intPageNumber = Request("txtPageNumber") + 1
					    if intPageNumber > intPageCount then intPageNumber = intPageCount
		case ">>"		intPageNumber = intPageCount
		case else	    if Request("hdnExport") <> "" then
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
						set objTxtStream = objFSO.CreateTextFile(strExportPath&strRealUserID&"-CustomerService.xls", true, false)

						with objTxtStream
							.WriteLine "<table border=1>"

							'export the header

							.WriteLine "<TR>"
							.WriteLine "<TH>CSID</TD>"
							.WriteLine "<TH>Customer Service Name</TD>"
							.WriteLine "<TH>Service Type</TD>"
							.WriteLine "<TH>Service Location</TD>"
							.WriteLine "<TH>Service Status</TD>"
							.WriteLine "<TH>Feature CSID</TD>"
							.WriteLine "<TH>Feature Customer Service Name</TD>"
							.WriteLine "<TH>Feature Service Status</TD>"



							.WriteLine "<TD NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
                            .WriteLine "</TR>"

							'export the body
							for k = 0 to UBound(aList, 2)
								.WriteLine "<TR>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(0,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(1,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(2,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(3,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(4,k))&"</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(5,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(6,k))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(7,k))&"&nbsp;</TD>"



								.WriteLine "<TD NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
								.WriteLine "</TR>"
							next
							.WriteLine "</table>"
						end with
						objTxtStream.Close
						strsql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-CustomerService.xls"";</script>"
						Response.Write strsql
						Response.End
'						Response.redirect "export/"&strRealUserID&"-CustomerService.xls"
			elseif Request("txtGoToPageNo") <> "" then
						intPageNumber = CInt(Request("txtGoToPageNo"))
					else
						intPageNumber = 1
					end if
	end select

	if intPageNumber < 1 then
		intPageNumber = 1
	end if
	if intPageNumber > intPageCount then
		intPageNumber = intPageCount
	end if




	dim k, m, n
	m = (intPageNumber - 1 ) * intConstDisplayPageSize
	n = (intPageNumber) * intConstDisplayPageSize - 1
	if n > UBound(aList, 2) then
		n = UBound(aList, 2)
	end if

	'check if the client is still connected just before sending any html to the browser
	if response.isclientconnected = false then
		Response.End
	end if

	'catch any unexpected error
	if err then
		DisplayError "BACK", "", err.Number, "Unexpected error", err.Description
	end if

%>
<BODY>
<FORM method=post name=frmCustServfList action="CustServfList.asp">


    <input type=hidden name=txtCustomerServiceDesc value="<%=strCustomerServiceDesc%>">
    <input type=hidden name=txtServiceLocationName value="<%=strServiceLocationName%>">


    <input type=hidden name=SelStatus value="<%=strStatusCode%>">
    <input type=hidden name=txtCustomerServiceID value="<%=intCustomerServiceID%>">


    <input type=hidden name=chkActiveOnly value="<%=bolActiveOnly%>">


    <input type=hidden name=txtServiceType value="<%=strServiceType%>"  >



    <input type=hidden name="hdnExport" value>
    <input type=hidden name=txtCustomerID value="<%=intCustomerID%>">
    <input type=hidden name=txtCustomerShortName value="<%=strCustomerShortName%>">


<TABLE  border=1 cellPadding=2 cellSpacing=0 width="100%">
  <THEAD>
    <TR>

        <TH align=left nowrap>CSID</TH>
        <TH align=left nowrap>Customer Service Desc</TH>
        <TH align=left nowrap>Service Type</TH>
        <TH align=left nowrap>Service Location</TH>
        <TH align=left nowrap>Status</TH>
        <TH align=left nowrap>Feature CSID</TH>
        <TH align=left nowrap>Feature Service Desc</TH>
        <TH align=left nowrap>Feature Service Status</TH>


     </TR>
 </THEAD>
 <TBODY>
<%

    for k = m to n
	''Alternate row background colour
	if Int(k/2) = k/2 then
		Response.write "<TR>"
	else
		Response.write "<TR bgcolor=White>"
	end if






		Response.Write "<TD nowrap>" &routineJavascriptString(aList(0, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap>" &routineJavascriptString(aList(1, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap>" &routineJavascriptString(aList(2, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap>" &routineJavascriptString(aList(3, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap>" &routineJavascriptString(aList(4, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap>" &routineJavascriptString(aList(5, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap>" &routineJavascriptString(aList(6, k))& "&nbsp;</a></TD>" &vbCrLf
		Response.Write "<TD nowrap>" &routineJavascriptString(aList(7, k))& "&nbsp;</a></TD>" &vbCrLf

		Response.Write "</TR>"

  next
%>
</TBODY>
<TFOOT>
<TR>
<TD align=left colSpan=11>
	<input type=hidden name=txtPageNumber value=<%=intPageNumber%>>

	<input type="hidden" name="txtcolor" value=<%=color%>>

	<input type="submit" name="action" value="&lt;&lt;">
	<input type="submit" name="action" value="&lt;">
	<input type="text" name="txtGoToPageNo" title="You can jump to a specific page by typing the page number in this box" onclick="document.frmCustServList.txtGoToPageNo.value=''" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;">
	<img SRC="images/excel.gif" onclick="document.frmCustServfList.target='new'; document.frmCustServfList.hdnExport.value='xls';document.frmCustServfList.submit();document.frmCustServfList.hdnExport.value='';document.frmCustServfList.target='_self';" WIDTH="32" HEIGHT="32">
</TD>
</TR>
</TFOOT>
<CAPTION>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></CAPTION>
</TABLE>
</FORM>
</BODY>
</HTML>
