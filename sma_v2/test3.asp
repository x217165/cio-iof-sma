<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="smaConstants.inc"-->
<%
 Dim objConn,StrConnectString
 StrConnectString = strConstConnectString
 set objConn = Server.CreateObject("ADODB.Connection")
 objConn.ConnectionString = StrConnectString
 objConn.open
'<!--#include file="databaseconnect.asp"-->
%>
<HTML>
<HEAD>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
</HEAD>

<BODY>
<FORM  id=form1 name=form1 action=test3.asp>
<TABLE width="100%">
  <THEAD>
    <TR>
		<TH width="20%">Customer Name</TH>
		<TH width="20%">Street</TH>
        <TH width="20%">Building</TH>
        <TH width="20%">City</TH>
        <TH width="20%">Province</TH>
    </TR>
  </THEAD>
    
    <input type=hidden name=txtCustomerName value="<%=strCustomerName%>">
    <input type=hidden name=txtStreet value="<%=strStreet%>">
    <input type=hidden name=txtCity value="<%=strCity%>">
    <input type=hidden name=txtPostal value="<%=strPostal%>">
    <input type=hidden name=radAddressType value="<%=bolAddressType%>">
 <%
dim timeStart
timeStart=Timer()
 
dim strCustomerName, strStreet, strCity, strPostal, bolAddressType
dim strSQL, strSelectClause,strFromClause, strWhereClause, strRecordStatus, strOrderBy
dim intPageNumber, intPageCount
	
strCustomerName = UCase("CAD")
strStreet = UCase(trim(Request.Form("txtStreet")))
strCity = UCase(trim(Request.Form("txtCity")))
strPostal = UCase(trim(Request.Form("txtPostal")))
bolAddressType = trim(Request.Form("radAddressType"))
	
strSQL = "select distinct(a.address_id), " &_
 		"c.customer_name, " &_
 		"a.street_name, " &_
 		"a.building_name, " &_
 		"a.municipality_name, " &_
 		"a.province_state_lcode " &_
 	"from crp.customer c, " &_
 		 "crp.address a, " &_	
 		 "crp.customer_name_alias c1 "
			
strWhereClause =    "where c.customer_id = a.customer_id " &_
 					"and   c.customer_id = c1.customer_id " &_
 					"and   c.record_status_ind = 'A' "
			    		
						
if len(strCustomerName) > 0 then
 	strWhereClause = strWhereClause & "and c1.customer_name_alias_upper like '" & strCustomerName & "%'"
end if
	
if len(strStreet) > 0 then
 	strWhereClause = strWhereClause & "and UPPER(a.street_name) like '" & strStreet & "%'" 
end if
	
	
if len(strCity) > 0 then
 	strWhereClause = strWhereClause & "and UPPER(a.municipality_name) like '" & strCity & "%'" 
end if
	
if len(strPostal) > 0 then
	strWhereClause = strWhereClause & "and UPPER(a.postal_code_zip) like '" & strPostal & "%'" 
end if
	
select case bolAddressType
 	case  "billing"
 		strWhereClause = strWhereClause & " and a.billing_address_flag = 'Y' "
		
 	case "mailing"
 		strWhereClause = strWhereClause & " and a.mailing_address_flag = 'Y' "
		
 	case "primary"
 		strWhereClause = strWhereClause & " and a.primary_address_flag = 'Y' "		
end select
	
strsql = strsql & strWhereClause
	
Dim objRs, strbgcolor
     
'Create Recordset object  
set objRS = objConn.Execute(StrSql)
strbgcolor = "white"
dim aTest	
aTest = objRS.GetRows

'close and disconnect the recordset
objRS.close
set objRS = Nothing

'close and release the connection
objConn.close
set ObjConn = Nothing    

intPageCount = Int(UBound(aTest, 2) / intConstDisplayPageSize)
on error resume next
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
if intPageNumber > Int(UBound(aTest,2)/intConstDisplayPageSize) then intPageNumber = Int(UBound(aTest,2)/intConstDisplayPageSize)
dim k, m, n
m = (intPageNumber) * intConstDisplayPageSize
n = (intPageNumber + 1) * intConstDisplayPageSize
if n > UBound(aTest, 2) then 
	n = UBound(aTest, 2)
end if

'check if the client is still connected
if response.isclientconnected = false then
	Response.End
end if

for k = m to n
	'Alternate table background colour
	if strbgcolor ="white" then
		strbgcolor = "silver"
	else
		strbgcolor = "white"
	end if
	%>
	<TR bgcolor = "<%=strbgcolor%>">
	<TD align=left NOWRAP><a href ="AddressDetail.asp?AddressID=<%=aTest(0,k)%>" TARGET="_parent"><%=aTest(1,k)%></TD>
	<TD align=left NOWRAP><a href ="AddressDetail.asp?AddressID=<%=aTest(0,k)%>" TARGET="_parent"><%=aTest(2,k)%></TD>
	<TD align=left NOWRAP><a href ="AddressDetail.asp?AddressID=<%=aTest(0,k)%>" TARGET="_parent"><%=aTest(3,k)%></TD>
	<TD align=left NOWRAP><a href ="AddressDetail.asp?AddressID=<%=aTest(0,k)%>" TARGET="_parent"><%=aTest(4,k)%></TD>
	<TD align=left NOWRAP><a href ="AddressDetail.asp?AddressID=<%=aTest(0,k)%>" TARGET="_parent"><%=aTest(5,k)%></TD>
	<%

next
'display page navigation buttons if needed
%>
<tfoot>
<tr>
<td align=left colSpan=4>
	<input type=hidden name=txtPageNumber value=<%=intPageNumber%>>
	<input type="submit" name="action" value="&lt;&lt;">
	<input type="submit" name="action" value="&lt;">
	<input type="text" name="txtGoToPageNo" title="You can jump to a specific page by typing the page number in this box" value="page <%=intPageNumber%> of 80" style="HEIGHT: 22px; WIDTH: 150px">
	<input type="submit" name="action" value="&gt;">
	<input type="submit" name="action" value="&gt;&gt;"> 
</td>
<td align=right>Total:&nbsp;<%=UBound(aTest, 2) & " records"%>
</td>
</tr>
</tfoot>
<caption align=left>
	Use it for title or ...
</caption>
</table>

</FORM>
</BODY>
</HTML>
