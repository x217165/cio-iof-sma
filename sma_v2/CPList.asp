<%@ Language=VBScript %>
<% option explicit %>
<!--% on error resume next%-->
<!--
********************************************************************************************
* Page name:	CPList.asp
* Purpose:		To display the results of a customer profile search.
*				Search criteria are chosen via CPCriteria.asp
*
* Created by:	Anthony Cheung	09/05/2013
*
********************************************************************************************
        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------


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

dim strCustomerProfileName, strCustomerProfileID
dim strSQL, strSelectClause, strFromClause, strWhereClause, strRecordStatus, strOrderBy
dim intPageNumber, intPageCount, intTotalCPcount
dim strMyWinName, strBgColor,strServiceEnd

'SOAP variables
dim strwsStatus, record_count, cpList(2,100)
dim cpname


'get search criteria
	strMyWinName = Request("hdnWinName")
	strServiceEnd = Request("ServiceEnd")

	strCustomerProfileName = UCase(routineOraString((trim(Request("txtCustomerProfileName")))))
	strCustomerProfileName = REPLACE(strCustomerProfileName, "*", "") ' replace * with space
	strCustomerProfileName = REPLACE(strCustomerProfileName, "%", "*") ' replace % with *
	strCustomerProfileID = UCase(routineOraString((trim(Request("txtCustomerProfileID")))))

	if strServiceEnd = "" then
	 strServiceEnd = "OTHER"
	END IF

'SOAP call
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

	If strCustomerProfileID <> "" Then 'CPID entered
		strwsStatus = CP_GetCPName (strCustomerProfileID,record_count,cpname)
		'Response.write "<p>CP_GetCPName = " & strwsStatus & "</p>"
		if  strwsStatus = 200 and record_count = 1 then
			cpList(0,0) = strCustomerProfileID
			cpList(1,0) = cpname
		else
			Response.Write "0 Records Found"
			Response.End
		end if
	elseif strCustomerProfileName <> "" Then 'CP name entered
		strwsStatus = CP_customers (strCustomerProfileName,100,record_count,cpList)
		if  strwsStatus <> 200 then
			Response.Write "0 Records Found"
			Response.End
		end if
		'Response.write "<p>CP_customers = " & strwsStatus & "</p>"
	else
		Response.Write "0 Records Found"
		Response.End
	end if

	'Response.write "<p>Status = " & strwsStatus & "</p>"
	'Response.write "<p>Size = " & record_count & "</p>"

End If

	'calculate page number
	intPageCount = int((record_count / intConstDisplayPageSize) + 1)
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
							DisplayError "CLOSE", "", err.Number, "CPList.asp - Cannot create Excel spreadsheet file due to the following reasons. Please contact your website administrator.", err.Description
						end if

						with objTxtStream
							.WriteLine "<table border=1>"

							'export the header
							.WriteLine "<THEAD>"
							.WriteLine "<TH>Customer Profile ID</TH>"
							.WriteLine "<TH>Customer Profile Name</TH>"
							.WriteLine "</THEAD>"

							'export the body
							for k = 0 to record_count-1
								.WriteLine "<TR>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(cpList(k,1))&"&nbsp;</TD>"
								.WriteLine "<TD NOWRAP>"&routineHtmlString(cpList(k,0))&"&nbsp;</TD>"
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
	if n > record_count  then
		n=record_count-1
	end if

	'Response.Write m & "<br />" & vbCrLf	'the smallest index
	'Response.Write n & "<br />" & vbCrLf	'the largest index

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
setPageTitle("SMA - Customer Profile");

	function go_back(strServiceEnd,lngCustomerProfileID, strCustomerProfileName){
	    //Response.Write ("inside go_back strServiceEnd is " & strServiceEnd)
		//alert (strServiceEnd);

		try
		{
			if (strServiceEnd == 'A'){ //this condition handles the customer - customer profile lookup
				parent.opener.document.forms[0].txtCustomerProfileID.value = lngCustomerProfileID;
				parent.opener.document.forms[0].txtCustomerProfileName.value = strCustomerProfileName;
			}
			else if (strServiceEnd == 'B'){ //this condition handles the customer service - customer profile lookup
				parent.opener.document.forms[0].txtCustomerProfileID.value = lngCustomerProfileID;
				parent.opener.document.forms[0].txtCustomerProfileName.value = strCustomerProfileName;
			}
			else {
				parent.opener.document.forms[0].txtCustomerProfileID.value = lngCustomerProfileID;
				parent.opener.document.forms[0].txtCustomerProfileName.value = strCustomerProfileName;
			}
		}
		catch(e)
		{
		 //do nothing, most probably not all forms have CustomerShortName - needed in Managed Objects Details
		}
		parent.window.close ();
	}
	//-->
	</SCRIPT>

</head>
<body>
<form name=frmCPList action="CPList.asp" method=post>
	<input type=hidden name=txtCustomerProfileName value="<%=strCustomerProfileName%>">
	<input type=hidden name=txtCustomerProfileID value="<%=strCustomerProfileID%>">
	<input type=hidden name=hdnServiceEnd value="<%=strServiceEnd%>">
<TABLE border=1 cellPadding=2 cellSpacing=0 width="100%">
	<THEAD>
	<TR>
		<TH align=left nowrap>Customer Profile ID</TH>
        <TH align=left nowrap>Customer Profile Name</TH>
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
			Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" &strServiceEnd& "','"&cpList(0,k)& "', '" &routineJavascriptString(cpList(1,k))& "')"">" &cpList(0,k)& "</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a href=""#"" onClick=""return go_back('" &strServiceEnd& "','"&cpList(0,k)& "', '" &routineJavascriptString(cpList(1,k))& "')"">" &cpList(1,k)& "</a>&nbsp;</td>"&vbCrLf
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
</TBODY>
<TFOOT>
<TR>
<TD align=left colSpan=8>
	<input type=hidden name=hdnWinName value="<%=strMyWinName%>">
	<input type=hidden name=txtPageNumber value=<%=intPageNumber%>>
	<input type="submit" name="action" value="&lt;&lt;">
	<input type="submit" name="action" value="&lt;">
	<input type="text" name="txtGoToPageNo" onClick="document.frmCPList.txtGoToPageNo.value=''" title="You can jump to a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;">&nbsp;&nbsp;
	<img SRC="images/excel.gif" onclick="document.frmCustCPList.target='new';document.frmCPList.hdnExport.value='xls';document.frmCPList.submit();document.frmCustCPList.hdnExport.value='';document.frmCPList.target='_self';" WIDTH="32" HEIGHT="32">
</TD>
</TR>
</TFOOT>
<CAPTION>Records <%=m+1%> to <%=n+1%> of <%=record_count & " records"%></CAPTION>
</table>
</form>
</body>
</html>








