<%@ Language=VBScript %>
<% option explicit %>
<%on error resume next%>

<!--
********************************************************************************************
* Page name:	RSAST3IPGenList.asp
* Purpose:		To display the results of a Gateway IP Search.
*				Search criteria are chosen via RSAST3IPGenCriteria.asp
*
* Created by: Dan Ty
********************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
********************************************************************************************

-->
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->
<%

dim strGatewayIP, strIPAddress, strCode,strLocation, strAvailable, bolActiveOnly
dim aList, strWinName
dim rsIPAddress

dim strSQL, strSelectClause, strFromClause, strWhereClause, strRecordStatus, strOrderBy
dim intPageNumber, intPageCount

strGatewayIP  = UCase(routineOraString(trim(Request("txtGatewayIP"))))
strIPAddress  = UCase(routineOraString(trim(Request("txtIPAddress"))))
strCode       = Request("selCode")
strLocation   = Request("selLocation")
strAvailable  = Request("selAvailable")
bolActiveOnly = Request("chkActiveOnly")

'get window name and WorkFor name
strWinName = Request("hdnWinName")

'build query
strSelectClause = "select ip_address_id, gateway_ip_address, ip_address, " &_
  "subnet_mask, code, location, available, comments, record_status_ind "

strFromClause = " from crp.rsas_ip_address "

strWhereClause = " where ip_address_id <> 0 "

'Gateway IP entered
If strGatewayIP <> "" then
	strWhereClause = strWhereClause & " and gateway_ip_address like '" & strGatewayIP & "%'"
End If

'IP Address entered
If strIPAddress <> "" then
	strWhereClause = strWhereClause & " and ip_address like '" & strIPAddress & "%'"
End If

'Code entered
If strCode <> "" then
	strWhereClause = strWhereClause & " and code like '" & strCode & "%'"
End If

'Location entered
If strLocation <> "" then
	strWhereClause = strWhereClause & " and location like '" & strLocation & "%'"
End If

'Available entered
If strAvailable <> "" then
	strWhereClause = strWhereClause & " and available like '" & strAvailable & "%'"
End If

'see all records?
If bolActiveOnly = "on" then
   strWhereClause = strWhereClause & " and record_status_ind = 'A'"
End If

	'order by what?
    strOrderBy =  " order by gateway_ip_address, ip_address"

'join all pieces to make a complete query
strSQL = strSelectClause & strFromClause & strWhereClause & strOrderBy

'get the recordset
set rsIPAddress=server.CreateObject("ADODB.Recordset")
rsIPAddress.Open strSQL, objConn

If err then
	DisplayError "BACK", "", err.Number, "RSAST3IPGenList.asp - Cannot open database" , err.Description
End if

'put recordset into array
if not rsIPAddress.EOF then
	aList = rsIPAddress.GetRows
else
	Response.Write "0 Record Found"
	Response.End
end if

'release and kill the recordset and the connection objects
rsIPAddress.Close

set rsIPAddress = nothing
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
                              strRealUserID =Session("username")

                              'determine export path

                              dim strExportPath, liLength
                              strExportPath =Request.ServerVariables("PATH_TRANSLATED")
                              While (Right(strExportPath, 1) <> "\" And Len(strExportPath) <> 0)
                                   liLength = Len(strExportPath) - 1
                                   strExportPath = Left(strExportPath, liLength)
                              Wend
                              strExportPath = strExportPath & "export\"

                              'create the scripting object

                              dim objFSO, objTxtStream
                              set objFSO = server.CreateObject("Scripting.FileSystemObject")

                              'create the export text file (overwrite if it already exists)

                              set objTxtStream = objFSO.CreateTextFile(strExportPath&strRealUserID&"-Gateway-IP.xls", true,false)

								if err then
										DisplayError "CLOSE", "", err.Number, "RSAST3IPGenList.asp - Cannot create Excel spreadsheet file due to the following reasons.  Please contact your website administrator.", err.Description
								end if

							  with objTxtStream

                                   .WriteLine "<table border=1>"

                                   'export the table header
                                   .WriteLine "<TR>"

                                   .WriteLine "<TH>Gateway IP</TH>"
                                   .WriteLine "<TH>IP Address</TH>"
                                   .WriteLine "<TH>IP Sub-mask</TH>"
                                   .WriteLine "<TH>Code</TH>"
                                   .WriteLine "<TH>Location</TH>"
                                   .WriteLine "<TH>Available</TH>"
                                   .WriteLine "<TH>Comments</TH>"

                                   	If bolActiveOnly <> "on" then
	                                   .WriteLine "<TH>Active Status</TH>"
	                                end if

								   'end the table header
                                   .WriteLine "</TR>"


                                   'export the body
                                   for k = 0 to UBound(aList, 2)
                                         'Alternate row background colour
                                         if Int(k/2) = k/2 then
'                                             .WriteLine "<TR bgcolor=#ffffcc>"
                                              .WriteLine "<TR>"
                                         else
'                                             .WriteLine "<TR bgcolor=#ffffff>"
                                              .WriteLine "<TR>"
                                         end if


                                         'fill the table with data

                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(1,k))&"</TD>"
                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(2,k))&"</TD>"
                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(3,k))&"</TD>"
                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(4,k))&"</TD>"
                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(5,k))&"</TD>"
                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(6,k))&"</TD>"
                                         .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(7,k))&"</TD>"

			                           	 If bolActiveOnly <> "on" then
                                            .WriteLine "<TD NOWRAP>"&routineHtmlString(aList(8,k))&"</TD>" 'status
					                     end if

                                         .WriteLine "</TR>"
                                   next
                                   .WriteLine "</table>"

                              end with

                              objTxtStream.Close
							strsql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-Gateway-IP.xls"";</script>"
							Response.Write strsql
							Response.End

					elseif Request("txtGoToPageNo") <> "" then
						intPageNumber=CInt(Request("txtGoToPageNo"))
					else
						intPageNumber=1
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
	<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css" type="text/css">
	<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
	<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>
	<script ID=clientEventHandlersJS type="text/javascript">
	<!--

setPageTitle("SMA - POS PLUS Admin");

	function btnEdit_onclick(lngCustomerContactID){
	var url ;

		url = 'RSAST3IPGenDetail.asp?IPAddressID=' + lngIPAddressID;
		self.open(url,'Popup','top=50, left=100, height=600, width=800' );
	}

	//-->
</SCRIPT>
</head>
<body>
<form name=frmRSAST3IPGenList action="RSAST3IPGenList.asp" method=post>

<!-- hidden variables -->
	<input type=hidden name=txtGatewayIP  value="<%=strGatewayIP%>">
	<input type=hidden name=txtIPAddress  value="<%=strIPAddress%>">
	<input type=hidden name=selCode       value="<%=strCode%>">
	<input type=hidden name=selLocation   value="<%=strLocation%>">
	<input type=hidden name=selAvailable  value="<%=strAvailable%>">
	<input type=hidden name=chkActiveOnly value="<%=bolActiveOnly%>">
	<input type=hidden name=hdnWinName    value="<%=strWinName%>">
    <input type="hidden" name="hdnExport" value>


<TABLE border="1" width=100% cellspacing=0 cellpadding=2 >
	<THEAD>
		<tr><th align=left colspan=11>RSAS GATEWAY IP Address Results</th></tr>
		<TR>
		    <TH align=left nowrap>Gateway IP  </TH>
		    <TH align=left nowrap>IP Address  </TH>
		    <TH align=left nowrap>Subnet Mask </TH>
		    <TH align=left nowrap>Code        </TH>
		    <TH align=left nowrap>Location    </TH>
		    <TH align=left nowrap>Is Available</TH>
		    <TH align=left nowrap>Comments    </TH>
		    <%if bolActiveOnly = "" then
				Response.Write ("<TH align=left nowrap>Active Status </TH>")
			end if%>
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

			Response.Write "<td nowrap><a target=""_parent"" href=""RSAST3IPGenDetail.asp?hdnIPAddressID="&aList(0,k)&""">"&aList(1,k)& "</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""RSAST3IPGenDetail.asp?hdnIPAddressID="&aList(0,k)&""">"&aList(2,k)& "</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""RSAST3IPGenDetail.asp?hdnIPAddressID="&aList(0,k)&""">"&aList(3,k)& "</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""RSAST3IPGenDetail.asp?hdnIPAddressID="&aList(0,k)&""">"&aList(4,k)& "</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""RSAST3IPGenDetail.asp?hdnIPAddressID="&aList(0,k)&""">"&aList(5,k)& "</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""RSAST3IPGenDetail.asp?hdnIPAddressID="&aList(0,k)&""">"&aList(6,k)& "</a>&nbsp;</td>"&vbCrLf
			Response.Write "<td nowrap><a target=""_parent"" href=""RSAST3IPGenDetail.asp?hdnIPAddressID="&aList(0,k)&""">"&aList(7,k)& "</a>&nbsp;</td>"&vbCrLf

			if bolActiveOnly = "" then
				Response.Write "<td nowrap align=center><a target=""_parent"" href=""RSAST3IPGenDetail.asp?hdnIPAddressID="&aList(0,k)&""">"&aList(8,k)& "</a>&nbsp;</td>"&vbCrLf
			end if
			Response.Write "</tr>"
	   next
		%>
	</TBODY>
	<TFOOT>
	<TR>
		<TD align=left colSpan=11 >
			<input type=hidden name=txtPageNumber value=<%=intPageNumber%>>
			<input type="submit" name="action" value="&lt;&lt;">
			<input type="submit" name="action" value="&lt;">
			<input type="text" name="txtGoToPageNo" onClick="document.frmRSAST3IPGenList.txtGoToPageNo.value=''" title="You can jump to a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
			<input type="submit" name="action" value="&gt;">
			<input type="submit" name="action" value="&gt;&gt;">
			<img SRC="images/excel.gif" onclick="document.frmRSAST3IPGenList.target='new';document.frmRSAST3IPGenList.hdnExport.value='xls';document.frmRSAST3IPGenList.submit();document.frmRSAST3IPGenList.hdnExport.value='';document.frmRSAST3IPGenList.target='_self';" WIDTH="32" HEIGHT="32">
		</TD>
	</TR>
	</TFOOT>
	<CAPTION>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></CAPTION>
</table>
</form>
</body>
</html>






















