<%@ Language=VBScript %>
<% option explicit %>
<!--% on error resume next%-->
<!--
********************************************************************************************
* Page name:	ContactCleanList.asp                                                       *
* Purpose:		To perform Contact amalgamation and display the results.                   *
*				Contact parameters entered through ContactCleanEntry.asp                   *
*                                                                                          *
* Created by:	Dan S. Ty	03/29/2002                                                     *
*                                                                                          *
********************************************************************************************
*       Date		Author			Changes/enhancements made                              *
*       -----		------		------------------------------------------------------     *
*                                                                                          *
********************************************************************************************
-->
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp"-->

<%
 

'Check Access rights - check other locations in this page.
dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_ESDCleanup))
If (intAccessLevel And intConst_Access_ReadOnly) <> intConst_Access_ReadOnly  Then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to ESD Cleanup functions. Please contact your system administrator"
End If

dim rsChangeList, aList, lngBatchNumber, strXLSFile
dim strSQL, strRealUserID, strAction, ExportPath
dim intPageNumber, intPageCount
dim strMyWinName, strBgColor

strRealUserID  = Session("username")
lngBatchNumber = Request.Cookies("BatchNumber")
strXLSFile     = Request("hdnExport")

strAction = request("selAction")

if Request("hdnFRContactID") <>"" then
   dim cmdExecObj
   set cmdExecObj = server.CreateObject("ADODB.Command")
   set cmdExecObj.ActiveConnection = objConn
   cmdExecObj.CommandType = adCmdStoredProc
   cmdExecObj.CommandText = "sma_sp_userid.spk_sma_admin_inter.sp_cont_amalgamate"

   'create params
   cmdExecObj.Parameters.Append cmdExecObj.CreateParameter("p_userid",        adVarChar, adParamInput, 20 , strRealUserID)
   cmdExecObj.Parameters.Append cmdExecObj.CreateParameter("p_batch_number",  adNumeric, adParamOutput,   , null)
   cmdExecObj.Parameters.Append cmdExecObj.CreateParameter("p_action",        adChar,    adParamInput, 1  , Request("selAction"))
   cmdExecObj.Parameters.Append cmdExecObj.CreateParameter("p_fr_contact_id", adNumeric, adParamInput,    , Request("hdnFRContactID"))
   cmdExecObj.Parameters.Append cmdExecObj.CreateParameter("p_to_contact_id", adNumeric, adParamInput,    , Request("hdnTOContactID"))
dim objparm
for each objparm in cmdExecObj.Parameters
  Response.Write "<b>" & objparm.name & "</b>"
  Response.Write " and value: " & objparm.value & ""
  Response.Write " and datatype: " & objparm.Type & "<br>"
next
'response.end

   on error resume next
   cmdExecObj.Execute

   if objConn.Errors.Count <> 0 then
	  DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT RUN Contact Cleanup", objConn.Errors(0).Description
	  objConn.Errors.Clear
   else
	 lngBatchNumber = cmdExecObj.Parameters("p_batch_number").Value
	 Response.Cookies("BatchNumber") = lngBatchNumber
	 strXLSFile = ""
   end if
end if

   'Build query to extract changes made.
   strSQL = "select batch_number, owner_name, table_name, column_name, rec_id, " &_
            "  old_value, new_value, error_message " &_
            "  from crp.change_audit"

   if lngBatchNumber <> 0 then
      strSQL = strSQL & " where batch_number = " & lngBatchNumber
   end if

   strSQL = strSQL & " order by batch_number, owner_name, table_name, column_name"

   'get the recordset
   on error resume next
   set rsChangeList=server.CreateObject("ADODB.Recordset")
   rsChangeList.Open strSQL, objConn
   If err then
	   DisplayError "BACK", "", err.Number, "contactCleanList.asp - Cannot run stored procedure" , err.Description
   End if

   'put recordset into array
   if not rsChangeList.EOF then
	  aList = rsChangeList.GetRows
   else
	  Response.Write "0 Record Found"
	  Response.End
   end if

   'release and kill the recordset and the connection objects
   rsChangeList.Close

   set rsChangeList = nothing
       objConn.Close

   set objConn = nothing

'Create the Contact Cleanup spreadsheet
if strXLSFile = "" and request("action") = "" then
	dim strExportPath, liLength
	strExportPath = Request.ServerVariables("PATH_TRANSLATED")
	While (Right(strExportPath, 1) <> "\" And Len(strExportPath) <> 0)
		liLength = Len(strExportPath) - 1
		strExportPath = Left(strExportPath, liLength)
	Wend
	strExportPath = strExportPath & "export\save\"

	'create scripting object
	dim objFSO, objTxtStream
	set objFSO = server.CreateObject("Scripting.FileSystemObject")

	'create export file and save for future use.
	strXLSFile =  "ContID" & request("hdnFRContactID") & "-to-" & Request ("hdnTOContactID") & "-Clean-" & year(now())  & "-" & month(now()) & "-" & day(now()) & "-" & hour(now()) & "-" & minute(now()) & "-" & second(now()) & "-" & strRealUserID & ".xls"
	set objTxtStream = objFSO.CreateTextFile(strExportPath & strXLSFile, false, false)
	if err then
		DisplayError "CLOSE", "", err.Number, "ContactCleanList.asp - Cannot create Excel spreadsheet file due to the following reasons. Please contact your website administrator.", err.Description
	end if

	with objTxtStream
		.WriteLine "<table border=1>"

		'export the header
		.WriteLine "<THEAD>"
		.WriteLine "<TH>Batch Number</TH>"
		.WriteLine "<TH>Schema</TH>"
		.WriteLine "<TH>Table Name</TH>"
		.WriteLine "<TH>Column Name</TH>"
		.WriteLine "<TH>Record ID</TH>"
		.WriteLine "<TH>Previous Value</TH>"
		.WriteLine "<TH>Current Value</TH>"
		.WriteLine "<TH>Error Message</TH>"
		.WriteLine "</THEAD>"

		'export the body
		for k = 0 to UBound(aList, 2)
			.WriteLine "<TR>"
			.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(0,k))&"&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(1,k))&"&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(2,k))&"&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(3,k))&"&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(4,k))&"&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(5,k))&"&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(6,k))&"&nbsp;</TD>" & vbCrLf
			.WriteLine "<TD NOWRAP>"&routineHtmlString(aList(7,k))&"&nbsp;</TD></TR>" & vbCrLf
		next
		.WriteLine "</table>"
	end with
	objTxtStream.Close
end if

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
						strSQL = "<script type=""text/javascript"">document.location=""export/save/" & strXLSFile & """;</script>"
						Response.Write strSQL
						Response.End
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
	<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css" type="text/css">
	<title>Service Management Application</title>
	<script type="text/javascript" SRC="GeneralJavaFunctions.js"></script>

	<script ID=clientEventHandlersJS type="text/javascript">
	<!--
setPageTitle("SMA - Contact Cleanup");
	//-->
	</SCRIPT>

</head>
<body>
<form name=frmContactCleanList action="contactCleanList.asp" method=post>
	<input type=hidden name="hdnExport" value=<%=strXLSFile%>>
	<input type=hidden name="hdnBatchNumber" value= <%=lngBatchNumber%>>
<TABLE border=1 cellPadding=2 cellSpacing=0 width="100%">
	<THEAD>
	<TR>
        <TH align=left nowrap>Batch Number</TH>
        <TH align=left nowrap>Schema</TH>
        <TH align=left nowrap>Table Name</TH>
        <TH align=left nowrap>Column Name</TH>
        <TH align=left nowrap>Record ID</TH>
        <TH align=left nowrap>Previous Value</TH>
        <TH align=left nowrap>New Value</TH>
        <TH align=left nowrap>Error Message</TH>
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

		Response.Write "<tr>"
		Response.Write "<td nowrap>" & aList(0, k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(1, k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(2, k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(3, k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(4, k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(5, k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(6, k) & "&nbsp;</td>" & vbCrLf
		Response.Write "<td nowrap>" & aList(7, k) & "&nbsp;</td></tr>" & vbCrLf

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
	<input type="text" name="txtGoToPageNo" onClick="document.frmContactCleanList.txtGoToPageNo.value=''" title="You can jump to a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;">&nbsp;&nbsp;
	<img SRC="images/excel.gif" onclick="document.frmContactCleanList.target='new';document.frmContactCleanList.hdnExport.value='<%=strXLSFile%>';document.frmContactCleanList.submit();document.frmContactCleanList.target='_self';" WIDTH="32" HEIGHT="32">
</TD>
</TR>
</TFOOT>
<CAPTION>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></CAPTION>
</table>
</form>
</body>
</html>
