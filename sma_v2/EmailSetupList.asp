<%@ Language=VBScript %>
<% option explicit%>
<%on error resume next%>

<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->

<%
'check the present user's rights

dim intAccessLevel

intAccessLevel = CInt(CheckLogon(strConst_AssetCatalogue))

if intAccessLevel < intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to asset catalogue. Please contact your system administrator."
end if

'****VARIABLES****

'declare the connection variables
dim sqlSelect,sqlFrom,sqlOrderBy,sqlString, rsEmailList

'declare variable to be used in array for rows
dim eList

'declare the caller variable to be used for gobacks, etc.
dim strMyWinName

'declare results variables for fields in bottom navigation 
dim intPageNumber,intPageCount

'****VARIABLES****

'get the caller variable value from the previous page
strMyWinName = Request("hdnWinName")



'connect to the database using the include file
'CONNECT using databaseconnect.asp

'extract the necessary data using sql query

sqlSelect = "SELECT " &_
			"SSC.Service_Status_Change_ID, "&_
			"SSC.From_Service_Status_Code, " &_
			"SSC.To_Service_Status_Code, " &_
			"SSC.Notify_Cust_Care_Staff_Flag, " &_
			"SSC.NOTIFY_PORTFOLIO_STAFF_FLAG, " &_
			"SSC.NOTIFY_DESIGN_STAFF_FLAG, " &_
			"SSC.NOTIFY_IMPLEMENT_MANAGER_FLAG, " &_
			"SSC.NOTIFY_IMPLEMENT_STAFF_FLAG, " &_
			"SSC.NOTIFY_INSTALLATION_STAFF_FLAG, " &_
			"SSC.NOTIFY_OPERATIONS_STAFF_FLAG "    
		
sqlFrom = "FROM " &_
		  "CRP.Service_Status_Change SSC "
		
sqlOrderBy = "ORDER BY " &_
			 "From_Service_Status_Code" 
			  
			
	    

sqlString = sqlSelect & sqlFrom & sqlOrderBy

	
'Response.Write (sqlString & "<p>") 
'Response.end



'set the recordset and parse through the data

set rsEmailList=server.CreateObject("ADODB.Recordset")
rsEmailList.Open sqlString, objConn

if err then
	DisplayError "BACK", "", err.Number, "EmailSetupList.asp - Cannot open database", err.Description
end if

'search through the recordset and get the data

if not rsEmailList.EOF then
	eList = rsEmailList.GetRows
else 
	Response.Write "0 records found"
	Response.end
end if



'calculate page number
intPageCount = Int(UBound(eList, 2) / intConstDisplayPageSize) + 1
select case Request("Action")
	case "<<"		intPageNumber = 1
	case "<"		intPageNumber = Request("txtPageNumber") - 1
					if intPageNumber < 1 then intPageNumber = 1
	case ">"		intPageNumber = Request("txtPageNumber") + 1
					if intPageNumber > intPageCount then intPageNumber = intPageCount
	case ">>"		intPageNumber = intPageCount
	case else		if Request("txtGoToPageNo") <> "" then 
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
if n > UBound(eList, 2) then 
	n = UBound(eList, 2)
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
<HTML>
<HEAD>
<TITLE>Email Setup List</TITLE>
<LINK href="stylesheets/styles.css" rel="stylesheet" type="text/css">
</HEAD>
<SCRIPT TYPE="TEXT/JAVASCRIPT">



//need to complete this function if this screen is used as a lookup

/*

function go_back(lngServStatChangeID, , ,)
{
	parent.opener.document.forms[0].hdnServiceStatusChangeID.value = lngServStatChangeID;
	--
	--
	parent.window.close ();
}

**/

</SCRIPT>

<BODY>


<FORM name="frmEmailSetupList" action="EmailSetupList.asp">

       <input type=hidden name=hdnWinName value="<%=strMyWinName%>">
    
<TABLE border=1 cellPadding=2 cellSpacing=0 width="100%">
	<THEAD> 
		<TR>
			<!-- <TH align=left>Catalogue ID</TH> -->
			<TH align=left>From Service Status Code</TH>
			<TH align=left>To Service Status Code</TH>
			<TH align=left>Customer Care</TH>
			<TH align=left>Portfolio Staff</TH>
			<TH align=left>Design Staff</TH>
			<TH align=left>Implement Manager</TH>
			<TH align=left>Implement Staff</TH>
			<TH align=left>Installation Staff</TH>
			<TH align=left>Operations Staff</TH>
			
		</TR>
	</THEAD>
<TBODY> 

<%

'display the table

dim strCustCare,strPortStaff,strDesignStaff,StrImplMgr,strImplStaff,strInstStaff,strOperStaff

for k = m to n
	'Alternate row background colour
	if Int(k/2) = k/2 then
		Response.write "<TR>"
	else
		Response.write "<TR bgcolor=White>"
	end if
	
	if elist(3,k) = "Y" then
		strCustCare = "=yes checked"
	else
		strCustCare = ""
	end if
	
	if elist(4,k) = "Y" then
		strPortStaff = "=yes checked"
	else
		strPortStaff = ""
	end if
	
	if elist(5,k) = "Y" then
		strDesignStaff = "=yes checked"
	else
		strDesignStaff = ""
	end if
	
	if elist(6,k) = "Y" then
		strImplMgr = "=yes checked"
	else
		strImplMgr = ""
	end if
	
	if elist(7,k) = "Y" then
		strImplStaff = "=yes checked"
	else
		strImplStaff = ""
	end if
	
	if elist(8,k) = "Y" then
		strInstStaff = "=yes checked"
	else
		strInstStaff = ""
	end if
	
	if elist(9,k) = "Y" then
		strOperStaff = "=yes checked"
	else
		strOperStaff = ""
	end if
	
	'this first condition is the list that appears in the popup window
	'if a lookup button is pressed.
	
	'if strMyWinName = "Popup" then
		
	   'sample only; currently this screen is never called as a lookup
	   'Response.Write "<td><a href=""#"" onClick=""return go_back('"&eList(0,k)& "', '" &routineJavascriptString(eList(1,k))& "','" &routineJavascriptString(aList(2,k))& "')"">" &eList(1,k)& "</a></td>"&vbCrLf
	   'Response.Write "</tr>"
		
	'this second condition is the list that appears upon initial load
	
	'else
		Response.Write "<TD NOWRAP><a target="""" href=""emailsetupdetail.asp?hdnServiceStatusChangeID="&eList(0,k)&""">"&routineHtmlString(eList(1,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP><a target="""" href=""emailsetupdetail.asp?hdnServiceStatusChangeID="&eList(0,k)&""">"&routineHtmlString(eList(2,k))&"&nbsp;</a></TD>"&vbCrLf
		Response.Write "<TD NOWRAP align=""center""><INPUT ID=""CustCare""  name=""CustCare"" type=""checkbox"" style=""HEIGHT: 22px; WIDTH: 22px"" DISABLED  VALUE" &strCustCare& "></TD>" &vbCrLf 
		Response.Write "<TD NOWRAP align=""center""><INPUT ID=""PortStaff""  name=""PortStaff"" type=""checkbox"" style=""HEIGHT: 22px; WIDTH: 22px"" DISABLED  VALUE" &strPortStaff& "></TD>" &vbCrLf 
		Response.Write "<TD NOWRAP align=""center""><INPUT ID=""DesignStaff""  name=""DesignStaff"" type=""checkbox"" style=""HEIGHT: 22px; WIDTH: 22px"" DISABLED  VALUE" &strDesignStaff& "></TD>" &vbCrLf 
		Response.Write "<TD NOWRAP align=""center""><INPUT ID=""ImplMgr""  name=""ImplMgr"" type=""checkbox"" style=""HEIGHT: 22px; WIDTH: 22px"" DISABLED  VALUE" &strImplMgr& "></TD>" &vbCrLf 
		Response.Write "<TD NOWRAP align=""center""><INPUT ID=""ImplStaff""  name=""ImplStaff"" type=""checkbox"" style=""HEIGHT: 22px; WIDTH: 22px"" DISABLED  VALUE" &strImplStaff& "></TD>" &vbCrLf 
		Response.Write "<TD NOWRAP align=""center""><INPUT ID=""InstStaff""  name=""InstStaff"" type=""checkbox"" style=""HEIGHT: 22px; WIDTH: 22px"" DISABLED  VALUE" &strInstStaff& "></TD>" &vbCrLf 
		Response.Write "<TD NOWRAP align=""center""><INPUT ID=""OperStaff""  name=""OperStaff"" type=""checkbox"" style=""HEIGHT: 22px; WIDTH: 22px"" DISABLED  VALUE" &strOperStaff& "></TD>" &vbCrLf 

	'end if
	
	Response.Write "</TR>"

next

%>

</TBODY>
<TFOOT>
<TR>
<TD align=left colSpan=9>
	<input type=hidden name=txtPageNumber value=<%=intPageNumber%>>
	<input type="submit" name="action" value="&lt;&lt;">
	<input type="submit" name="action" value="&lt;">
	<input type="text" name="txtGoToPageNo" onClick="document.frmEmailSetupList.txtGoToPageNo.value=''" title="You can jump to a specific page by typing the page number in this box" value="page <%=intPageNumber%> of <%=intPageCount%>" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;"> 
</TD>
</TR>
</TFOOT>
<CAPTION>Records <%=m+1%> to <%=n+1%> of <%=UBound(eList, 2)+1 & " records"%></CAPTION>
</TABLE>
</FORM>

<%

'close the recordset and the connection objects
rsEmailList.Close
set rsEmailList = nothing

objConn.close
set objConn = nothing


%>
</BODY>
</HTML>

