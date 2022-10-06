<%@ Language=VBScript     %>
<% option explicit        %>
<% 'on error resume next   %>
<% Response.Buffer = true %>
<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->
<!--
********************************************************************************************
* Page name:	MakeList.asp
* Purpose:		To dynamically display the results of a search for an asset make.
*
* In Param:		This page reads following parameters
*				txtUserID - the userid to search for
*               txtLastName - the user's last name
*               txtFirstName - the user's first name
*               selRole - the ID for the security role to search for
*
* Out Param:    This screen is not coded to be used as a lookup
*
*
* Created by:	Chris Roe Oct. 31, 2000
********************************************************************************************

        Date		Author			Changes/enhancements made
        -----		------		------------------------------------------------------
       18-Feb-02	 DTy		Add Active Only variable.
                                Active customers are those whose status is either
                                  'Prospect', 'OnHold' or 'Current'.

********************************************************************************************
-->

<%
'check user's rights
Const ASP_NAME = "StaffRoleList.asp"
Const DETAIL_PAGE = "StaffRoleDetail.asp"

dim intAccessLevel
intAccessLevel = CInt(CheckLogon(strConst_Security))
if (intAccessLevel and intConst_Access_ReadOnly) <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to security. Please contact your system administrator."
end if

dim sql, strMyWinName
dim rsList

dim aList, intPageNumber, intPageCount
dim strCriteria

'get the caller
strMyWinName = Request("hdnWinName")

'get search criteria
Dim strFirstName
Dim strLastName
Dim strUserID
Dim strRoleID
Dim boolShowRoles

Dim chkActiveOnly
strFirstName = UCase(Trim(Request("txtFirstName")))
strLastName = UCase(Trim(Request("txtLastName")))
strUserID = UCase(Trim(Request("txtUserID")))
strRoleID = UCase(Trim(Request("selRole")))
boolShowRoles = Trim(Request("chkShowRoles")) <> ""

'build query
sql = " SELECT distinct s.contact_id" &_
	  " ,      s.last_name" &_
 	  " ,      s.first_name" &_
 	  " ,      s.userid" &_
 	  " ,      cust.customer_name"

if boolShowRoles then
	sql = sql & " , sr.security_role_name"
end if

sql = sql & " ,      s.work_for_customer_id"

sql = sql & " FROM   msaccess.staff_security_role ssr" &_
	  " ,      crp.contact                  s" &_
	  " ,      msaccess.security_role       sr" &_
	  " ,      msaccess.tblSecurity         t" &_
	  " ,      crp.customer                 cust" &_
	  " WHERE  ssr.staff_id (+)= s.contact_id" &_
	  " AND    ssr.security_role_id = sr.security_role_id (+)" &_
	  " AND    t.staffid = s.contact_id" &_
	  " AND    cust.customer_id = s.work_for_customer_id" &_
	  " AND    s.staff_flag= 'Y'"

	if strUserID <> "" then
		sql = sql & " AND    upper(t.userid) LIKE '" & routineOraString(strUserID) & "%'"
	end if

	if strFirstName <> "" then
		sql = sql & " AND    upper(s.first_name) LIKE '" & routineOraString(strFirstName) & "%'"
	end if

	if strLastName <> "" then
		sql = sql & " AND    upper(s.last_name) LIKE '" & routineOraString(strLastName) & "%'"
	end if

	if strRoleID <> "" then
		sql = sql & " AND    sr.security_role_id = " & strRoleID
	end if

	if chkActiveOnly = "On" then
	   sql = sql & "  and (t1.customer_status_lcode IN ('Prospect-Current-OnHold')"
	end if

sql = sql & " ORDER BY s.last_name, s.first_name"

if boolShowRoles then
	sql = sql & ", sr.security_role_name"
end if

'get the recordset

set rsList=server.CreateObject("ADODB.Recordset")
rsList.Open sql, objConn

if err then
	DisplayError "BACK", "", err.Number, ASP_NAME & " - Cannot open database", err.Description
end if

if not rsList.EOF then
	aList = rsList.GetRows
else
	Response.Write "0 records found"
	Response.end
end if

'release and kill the recordset and the connection objects
rsList.Close
set rsList = nothing

objConn.close
set objConn = nothing

'calculate page number
intPageCount = Int(UBound(aList, 2) / intConstDisplayPageSize) + 1
select case Request("Action")
	case "<<"		intPageNumber = 1
	case "<"		intPageNumber = Request("txtPageNumber") - 1
					if intPageNumber < 1 then intPageNumber = 1
	case ">"		intPageNumber = Request("txtPageNumber") + 1
					if intPageNumber > intPageCount then intPageNumber = intPageCount
	case ">>"		intPageNumber = intPageCount
	'Case "Export"
	case else
			if Request("hdnExport") <> "" then
				Dim strRealUserID
				Dim strExportPath
				Dim liLength
				Dim objFSO
				Dim objTxtStream
				strRealUserID = Session("username")
				'determine export path
				strExportPath = Request.ServerVariables("PATH_TRANSLATED")
				Do While (Right(strExportPath, 1) <> "\" And Len(strExportPath) <> 0)
					liLength = Len(strExportPath) - 1
					strExportPath = Left(strExportPath, liLength)
				Loop
				strExportPath = strExportPath & "export\"

				'create scripting object
				Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
				'create export file (overwrite if exists)
				Set objTxtStream = objFSO.CreateTextFile(strExportPath & strRealUserID & "-staffrole.xls", True, False)
				if err then
					DisplayError "CLOSE", "", err.Number, ASP_NAME & " - Cannot create Excel spreadsheet file due to the following reasons.  Please contact your website administrator.", err.Description
				end if

				with objTxtStream
					.WriteLine "<TABLE border=1>"
					.WriteLine "<THEAD>"
					.WriteLine "<TH>First Name</TH>"
					.WriteLine "<TH>Last Name</TH>"
					.WriteLine "<TH>User ID</TH>"
					.WriteLine "<TH>Works For</TH>"
					if boolShowRoles then
						.WriteLine "<TH>Security Role</TH>"
					end if
					.WriteLine "</THEAD>"

					'export the body
					For k = 0 To UBound(aList, 2)
						.WriteLine "<TR>"
						.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(1, k)) & "&nbsp;</TD>"
						.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(2, k)) & "&nbsp;</TD>"
						.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(3, k)) & "&nbsp;</TD>"
						.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(4, k)) & "&nbsp;</TD>"
						if boolShowRoles then
							.WriteLine "<TD NOWRAP>" & routineHtmlString(aList(5, k)) & "&nbsp;</TD>"
						end if
						.WriteLine "</TR>"
					Next
					.WriteLine "</TABLE>"
				End With
				objTxtStream.Close
				Set objTxtStream = Nothing
				Set objFSO = Nothing
				'Response.Write "<SCRIPT type='text/javascript' language='javascript'>"
				'Response.Write "window.open('" & "export/" & strRealUserID & ".xls" & "');"
				'Response.Write "</SCRIPT>"
					sql = "<script type=""text/javascript"">document.location=""export/"&strRealUserID&"-staffrole.xls"";</script>"
					Response.Write sql
					Response.End
'				Response.redirect "export/"&strRealUserID&".xls"
	'case else
			elseif Request("txtGoToPageNo") <> "" then
				intPageNumber = CInt(Request("txtGoToPageNo"))
			else
				intPageNumber = 1
			end if
end select

if intPageNumber < 1 then intPageNumber = 1
if intPageNumber > intPageCount then intPageNumber = intPageCount

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
<html>
<head>
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Asset Make Results</title>
<link href="stylesheets/styles.css" rel="stylesheet" type="text/css">
</head>
<script TYPE="TEXT/JAVASCRIPT">
function go_back(ID, Desc)
{
	parent.opener.document.forms[0].hdnMakeID.value = ID;
	parent.opener.document.forms[0].txtMakeDesc.value = Desc;
	parent.window.close ();
}

function btnExcel_onClick()
{
	document.forms[0].target='new';
	document.forms[0].hdnExport.value='xls';
	document.forms[0].submit();
	document.forms[0].hdnExport.value='';
	document.forms[0].target='_self';

}


</script>
<body>

<form name="frmACList" action="<%=ASP_NAME%>" method="POST">
    <input type="hidden" name="txtFirstName" value="<%=routineHTMLString(strFirstName)%>">
    <input type="hidden" name="txtLastName" value="<%=routineHTMLString(strLastName)%>">
    <input type="hidden" name="txtUserID" value="<%=routineHTMLString(strUserID)%>">
    <input type="hidden" name="selRole" value="<%=routineHTMLString(strRoleID)%>">
    <input type="hidden" name="hdnWinName" value="<%=strMyWinName%>">
    <input type="hidden" name="chkShowRoles" value="<%=Request("chkShowRoles")%>">

<table border="1" cellPadding="2" cellSpacing="0" width="100%">
<thead>
	<tr>
	   <!-- <TH align=left>Catalogue ID</TH> -->
		<th align="left">Last Name</th>
		<th align="left">First Name</th>
		<th align="left">User ID</th>
		<th align="left">Works For</th>
		<%if boolShowRoles then Response.Write ("<th align=""left"">Security Role</th>") end if %>
	</tr>
</thead>
<tbody>
<%
'display the table
for k = m to n
	'Alternate row background colour
	if Int(k/2) = k/2 then
		Response.write "<TR>"
	else
		Response.write "<TR bgcolor=White>"
	end if
	Response.Write "<TD NOWRAP><a target=""_parent"" href=""" & DETAIL_PAGE & "?hdnContactID=" & aList(0,k) & """>"&routineHtmlString(aList(1,k))&"&nbsp;</a></TD>"&vbCrLf
	Response.Write "<TD NOWRAP><a target=""_parent"" href=""" & DETAIL_PAGE & "?hdnContactID=" & aList(0,k) & """>"&routineHtmlString(aList(2,k))&"&nbsp;</a></TD>"&vbCrLf
	Response.Write "<TD NOWRAP><a target=""_parent"" href=""" & DETAIL_PAGE & "?hdnContactID=" & aList(0,k) & """>"&routineHtmlString(aList(3,k))&"&nbsp;</a></TD>"&vbCrLf
	Response.Write "<TD NOWRAP><a target=""_parent"" href=""" & DETAIL_PAGE & "?hdnContactID=" & aList(0,k) & """>"&routineHtmlString(aList(4,k))&"&nbsp;</a></TD>"&vbCrLf

	if boolShowRoles then
		Response.Write "<TD NOWRAP><a target=""_parent"" href=""" & DETAIL_PAGE & "?hdnContactID=" & aList(0,k) & """>"&routineHtmlString(aList(5,k))&"&nbsp;</a></TD>"&vbCrLf
	end if

	Response.Write "</TR>"
next
%>

</tbody>
<tfoot>
<tr>
<td align="left" colSpan="6">
	<input type="hidden" name="txtPageNumber" value="<%=intPageNumber%>">
	<input type="submit" name="action" value="&lt;&lt;">
	<input type="submit" name="action" value="&lt;">
	<input type="text" name="txtGoToPageNo" title="You can jump to a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;">
	<!--<INPUT type="submit" name="action" value="Export" title="Export this list to Excel"> -->
	<img src="images/excel.gif" onclick="btnExcel_onClick();" WIDTH="32" HEIGHT="32">
    <input type="hidden" name="hdnExport" value>
</td>
</tr>
</tfoot>
<caption>Records <%=m+1%> to <%=n+1%> of <%=UBound(aList, 2)+1 & " records"%></caption>
</table>
</form>
</body>
</html>
