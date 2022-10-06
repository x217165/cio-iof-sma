<%@ Language=VBScript %>
<% option explicit %>
<!--#include file="databaseconnect.asp"-->

<HTML>
<HEAD>
	<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
	<LINK rel="stylesheet" type="text/css" href="stylesheets/styles.css">
</HEAD>

<FORM  id=form1 name=form1>
<BODY>
<TABLE>
    <TR>
		<TH>Customer Name</TH>
		<TH>Street</TH>
        <TH>Building</TH>
        <TH>City</TH>
        <TH>Province</TH></TR>
 <%
 dim timeStart
 timeStart=Timer()

 dim strCustomerName, strStreet, strCity, strPostal, bolAddressType
 dim strSQL, strSelectClause,strFromClause, strWhereClause, strRecordStatus, strOrderBy
	
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
	END IF
	
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
	'Response.Write strsql
	
    Dim objRs,Recordcnt,strbgcolor
     
   'Create Recordset object  
   set objRS = objConn.Execute(StrSql)
   Recordcnt = 0
   strbgcolor = "white"
     
   
   Do while Not objRS.EOF
   
     'Alternate table background colour
	if strbgcolor ="white" then
      strbgcolor = "silver"
    else
      strbgcolor = "white"
	end if
      
%>
      
     <TR bgcolor = "<%=strbgcolor%>">
     <TD align=left NOWRAP><a href ="AddressDetail.asp?AddressID=<%=objRS(0)%>" TARGET="_parent"><%=objRS(1)%> </TD>
     <TD align=left NOWRAP><a href ="AddressDetail.asp?AddressID=<%=objRS(0)%>" TARGET="_parent"><%=objRS(2) %></TD>
     <TD align=left NOWRAP><a href ="AddressDetail.asp?AddressID=<%=objRS(0)%>" TARGET="_parent"><%=objRS(3) %> </TD>
     <TD align=left NOWRAP><a href ="AddressDetail.asp?AddressID=<%=objRS(0)%>" TARGET="_parent"><%= objRS(4)%></TD>
     <TD align=left NOWRAP><a href ="AddressDetail.asp?AddressID=<%=objRS(0)%>" TARGET="_parent"><%= objRS(5)%></TD>
<%
 
    objRS.MoveNext
    Recordcnt =Recordcnt+1
 Loop
 
  Response.write "<b>Total=" & Recordcnt & "</b><br>"
dim timeEnd
timeEnd = Timer()
Response.Write "RS: Elapsed (on the server) " + FormatNumber(TimeEnd-TimeStart,2) + "(seconds)"
 
 'Clean up our ADO objects
    objRS.close
    set objRS = Nothing

    objConn.close
    set ObjConn = Nothing    
       
      
%>

</FORM>
</TABLE>
</BODY>
</HTML>
