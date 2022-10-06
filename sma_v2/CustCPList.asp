<%@ Language=VBScript %>
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
'if CheckLogon(strConst_Customer) = 0 then
'	Response.Write "You don't have access to Customer. Please contact your system administrator."
'	Response.End
'end if

dim strCustomerName, strSMRLName, strSMRFName, strRegion, strStatus, bolActiveOnly, strCustomerProfileName, strCustomerProfileID
dim rsCustList, aList
dim strSQL, strSelectClause, strFromClause, strWhereClause, strRecordStatus, strOrderBy
dim intPageNumber, intPageCount
dim strMyWinName, strBgColor,strServiceEnd, strCustShort

'SOAP variables
dim strwsStatus,record_count,cidList(100)

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

	strCustomerProfileName = UCase(routineOraString((trim(Request("txtCustomerProfileName")))))
	strCustomerProfileID = UCase(routineOraString((trim(Request("txtCustomerProfileID")))))

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

	'CPID entered
	If strCustomerProfileID <> "" then
		'Response.Write strCustomerProfileID & "<br />" & vbCrLf

		If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

			strwsStatus = CP_GetCustomerID(strCustomerProfileID,100,record_count,cidList)
			'Response.write "<p>Status = " & strwsStatus & "</p>"
			'Response.write "<p>Size = " & record_count & "</p>"
			
			if 	strwsStatus <> 200 or record_count = 0 then
				strWhereClause = strWhereClause &  " and " & _
				"t1.customer_id in (-999)"   'assuming there is no customer id = -999
			else
				Dim  cidindex
				'for cidindex = 0 to record_count
				'	Response.write "<p>CID " & cidindex & " = " & cidList(cidindex) & "</p>"
				'next

				for cidindex = 0 to record_count
					'Response.write "<p>in cid loop " & cidindex & " = " & cidList(cidindex) &"</p>"
					if cidindex = 0 Then
						strWhereClause = strWhereClause &  " and " & _
						"t1.customer_id in (" & int(cidList(cidindex))
					elseif cidindex < record_count Then
						strWhereClause = strWhereClause & ", " & int(cidList(cidindex))
					elseif cidindex = record_count Then
						strWhereClause = strWhereClause & ") "
					end if
				next
			end if 'WSstatus
		end if 'WS
	End if	'CPID is not null


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
	'Response.Write strSQL & vbCrLf	'show SQL for debugging

	'get the recordset
	set rsCustList=server.CreateObject("ADODB.Recordset")
	rsCustList.Open strSQL, objConn
	If err then
		DisplayError "BACK", "", err.Number, "CustCPList.asp - Cannot open database" , err.Description
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
	objConn.Close
	set objConn = nothing

	'calculate page number
	intPageCount = Int(UBound(aList,2) / intConstDisplayPageSize) + 1
	select case Request("Action")
		case "<<"	intPageNumber = 1
		case "<"	intPageNumber = Request("txtPageNumber")-1
					if intPageNumber < 1 then intPageNumber = 1
		case ">"	intPageNumber = Request("txtPageNumber") + 1
					if intPageNumber > intPageCount then intPageNumber = intPageCount
		case ">>"	intPageNumber=intPageCount
		case else	if Request("hdnExport") <> "" then
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
							DisplayError "CLOSE", "", err.Number, "CustCPList.asp - Cannot create Excel spreadsheet file due to the following reasons. Please contact your website administrator.", err.Description
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
						strsql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-customer.xls"";</script>"
						Response.Write strsql
						Response.End
'						Response.redirect "export/"&strRealUserID&"-customer.xls"
					elseif Request("txtGoToPageNo") <> "" then
						intPageNumber = CInt(Request("txtGoToPageNo"))
					else
						intPageNumber = 1
					end if
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

            <HTML>

            <HEAD>
                <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
                <META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
                <LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css" type="text/css">
                <title>Service Management Application</title>
                <script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>

                <script ID=clientEventHandlersJS type="text/javascript">
                    <!--
                    setPageTitle("SMA - Customer");

                    function go_back(strServiceEnd, lngCustomerID, strCustomerName, strCustomerShortName, strRegion) {
                        //Response.Write ("inside go_back strServiceEnd is " & strServiceEnd)
                        //alert (strServiceEnd);

                        try {
                            if (strServiceEnd == 'A') {
                                parent.opener.document.forms[0].hdnCustomerIdA.value = lngCustomerID;
                                parent.opener.document.forms[0].txtcustomera.value = strCustomerName;
                            } else if (strServiceEnd == 'B') {
                                parent.opener.document.forms[0].hdnCustomerIdB.value = lngCustomerID;
                                parent.opener.document.forms[0].txtcustomerb.value = strCustomerName;
                            } else if (strServiceEnd == 'C') { //this condition handles the customer service lookup
                                parent.opener.document.forms[0].hdnCustomerID.value = lngCustomerID;
                                parent.opener.document.forms[0].txtCustomerName.value = strCustomerName;
                            } else if (strServiceEnd == 'D') { // Region is returned to CustServDetail.asp
                                parent.opener.document.forms[0].hdnCustomerID.value = lngCustomerID;
                                parent.opener.document.forms[0].txtCustomerName.value = strCustomerName;
                                parent.opener.document.forms[0].txtRegion.value = strRegion;
                                parent.opener.document.forms[0].txtCustomerShortName.value = strCustomerShortName;
                            } else if (strServiceEnd == 'F') { // this condition handles FR Customer in CustCleanEntry.asp
                                parent.opener.document.forms[0].txtFRCustomer.value = "(" + lngCustomerID + ") " + strCustomerName;
                                parent.opener.document.forms[0].hdnFRCustomerID.value = lngCustomerID;
                                parent.opener.document.forms[0].hhnFRCustomerName.value = strCustomerName;
                            } else if (strServiceEnd == 'T') { // this condition handles TO Customer in CustCleanEntry.asp
                                parent.opener.document.forms[0].txtTOCustomer.value = "(" + lngCustomerID + ") " + strCustomerName;
                                parent.opener.document.forms[0].hdnTOCustomerID.value = lngCustomerID;
                                parent.opener.document.forms[0].hdnTOCustomerName.value = strCustomerName;
                            } else if (strServiceEnd == 'X') { // this condition handles Customer in XLSEntry.asp
                                parent.opener.document.forms[0].txtCustomer.value = "(" + lngCustomerID + ") " + strCustomerName;
                                parent.opener.document.forms[0].hdnCustomerID.value = lngCustomerID;
                                parent.opener.document.forms[0].hdnCustomerName.value = strCustomerName;
                            } else if (strServiceEnd == 'E') { //this condition handles the customer service lookup
                                //alert (strCustomerName);
                                parent.opener.document.forms[0].txtCustomerID.value = lngCustomerID;
                                parent.opener.document.forms[0].txtCustomerName.value = strCustomerName;
                                parent.opener.document.forms[0].txtCustomerShortName.value = strCustomerShortName;
                            } else {
                                parent.opener.document.forms[0].hdnCustomerID.value = lngCustomerID;
                                parent.opener.document.forms[0].txtCustomerName.value = strCustomerName;
                                parent.opener.document.forms[0].txtCustomerShortName.value = strCustomerShortName;
                            }
                        } catch (e) {
                            //do nothing, most probably not all forms have CustomerShortName - needed in Managed Objects Details
                        }
                        parent.window.close();
                    }
                    //-->
                </SCRIPT>

            </head>

            <body>
                <form name=frmCustCPList action="CustCPList.asp" method=post>
                    <input type=hidden name=txtCustomerName value="<%=strCustomerName%>">
                    <input type=hidden name=txtCustShort value="<%=strCustShort%>">
                    <input type=hidden name=txtSMRLName value="<%=strSMRLName%>">
                    <input type=hidden name=txtSMRFName value="<%=strSMRFName%>">
                    <input type=hidden name=selRegion value="<%=strRegion%>">
                    <input type=hidden name=selStatus value="<%=strStatus%>">
                    <input type=hidden name=hdnServiceEnd value="<%=strServiceEnd%>">
                    <input type=hidden name="hdnExport" value>
                    <input type=hidden name="chkActiveOnly" value="<%=bolActiveOnly%>">
                    <input type=hidden name=txtCustomerProfileName value="<%=strCustomerProfileName%>">
                    <input type=hidden name=txtCustomerProfileID value="<%=strCustomerProfileID%>">

                    <TABLE border=1 cellPadding=2 cellSpacing=0 width="100%">
                        <THEAD>
                            <TR>
                                <TH align=left nowrap>Customer ID</TH>
                                <TH align=left nowrap>Customer Name</TH>
                                <TH align=left nowrap>Short Name</TH>
                                <TH align=left nowrap>Region</TH>
                                <TH align=left nowrap>Status</TH>
                                <TH align=left nowrap>Service Mgnt Rep</TH>
                                <TH align=left nowrap>Primary Address</TH>
                                <TH align=left nowrap>City</TH>
                                <TH align=left nowrap>Prov/State</TH>
                            </TR>
                        </THEAD>
                        <TBODY>
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
			Response.Write "<td nowrap><a target=""_parent"" href=""CustCPDetail.asp?CustomerID="&aList(0,k)&""">"&aList(0,k)&"</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""CustCPDetail.asp?CustomerID="&aList(0,k)&""">"&aList(1,k)&"</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""CustCPDetail.asp?CustomerID="&aList(0,k)&""">"&aList(2,k)&"</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""CustCPDetail.asp?CustomerID="&aList(0,k)&""">"&aList(3,k)&"</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""CustCPDetail.asp?CustomerID="&aList(0,k)&""">"&aList(5,k)&"</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""CustCPDetail.asp?CustomerID="&aList(0,k)&""">"&aList(8,k)&"</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""CustCPDetail.asp?CustomerID="&aList(0,k)&""">"&aList(9,k)&"</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""CustCPDetail.asp?CustomerID="&aList(0,k)&""">"&aList(10,k)&"</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""CustCPDetail.asp?CustomerID="&aList(0,k)&""">"&aList(11,k)&"</a>&nbsp;</td>"&vbCrLf
			Response.Write "</tr>"
		end if
   next
	%>
                        </TBODY>
                        <TFOOT>
                            <TR>
                                <TD align=left colSpan=9>
                                    <input type=hidden name=hdnWinName value="<%=strMyWinName%>">
                                    <input type=hidden name=txtPageNumber value=<%=intPageNumber%>>
                                    <input type="submit" name="action" value="&lt;&lt;">
                                    <input type="submit" name="action" value="&lt;">
                                    <input type="text" name="txtGoToPageNo" onClick="document.frmCustList.txtGoToPageNo.value=''" title="You can jump to a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
                                    <input type="submit" name="action" value="&gt;">
                                    <input type="submit" name="action" value="&gt;&gt;">&nbsp;&nbsp;
                                    <img SRC="images/excel.gif" onclick="document.frmCustList.target='new';document.frmCustList.hdnExport.value='xls';document.frmCustList.submit();document.frmCustList.hdnExport.value='';document.frmCustList.target='_self';" WIDTH="32" HEIGHT="32">
                                </TD>
                            </TR>
                        </TFOOT>
                        <CAPTION>Records
                            <%=m+1%> to
                                <%=n+1%> of
                                    <%=UBound(aList, 2)+1 & " records"%>
                        </CAPTION>
                    </table>
                </form>
            </body>

            </html>
