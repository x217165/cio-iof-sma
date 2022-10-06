<%@ Language=VBScript     %>
<% option explicit        %>
<% Response.Buffer = true %>
<!--#include file="databaseconnect.asp"-->
<!--#include file="SMAProcs.inc"-->
<!--#include file="SMAConstants.inc"-->
<!-- 
***********************************************************
 This is a temporary screen and will be dropped when the 
 whole security module is completed.
 
 The grid that is displayed is completely determined
 by the query that is executed.  To change a column name,
 alias it in the query.
***********************************************************
--> 

<%

Dim objRS, objRoleRS
Dim sql, selectClause, fromClause, whereClause
Dim roleCount

sql = " SELECT distinct sr.security_role_id" &_
	  " ,      sr.security_role_name" &_
	  " FROM   msaccess.security_role sr" &_
	  " ,      msaccess.business_func_security_role bfsr" &_
	  " ,      msaccess.business_function bf" &_
	  " ,      msaccess.application a" &_
	  " WHERE 1=1" &_
	  " AND   sr.security_role_id = bfsr.security_role_id" &_
	  " AND   bfsr.business_function_id = bf.business_function_id" &_
	  " AND   bf.application_id = a.application_id" &_
	  " AND   a.application_name = 'SMA2'" &_
	  " ORDER BY 1"

set objRoleRs = objConn.execute(sql)


selectClause = " select" &_
               " bf.business_function_name      ""Business Function""" 

fromClause = " from" &_
             " msaccess.business_function                      BF" 

whereClause = " where" &_
              " bf.application_id = 1" 

roleCount = 0
While not objRoleRS.EOF   

	selectClause = selectClause & " ,sub" & roleCount & ".access_level_desc           """ & Left(objRoleRS(1), 30) & """"

    fromClause = fromClause & " , (" &_
        " select  t1.business_function_id" &_
                " ,t1.security_role_id" &_
                " ,t1.business_func_access_level_id" &_
                " ,t2.access_level" &_
                " ,t2.access_level_desc" &_
        " from    msaccess.business_func_security_role    t1" &_
                " ,msaccess.business_func_access_level    t2" &_
        " where   t1.business_func_access_level_id=t2.business_func_access_level_id" &_
        " and     t1.security_role_id = " & objRoleRS(0) & ") " &_
                                                        " sub" & roleCount

	whereClause = whereClause & " and     bf.business_function_id=sub" & roleCount & ".business_function_id(+)"
	
	roleCount = roleCount + 1
	objRoleRS.MoveNext
	
wend

sql = selectClause & fromClause & whereClause & " order   by bf.business_function_id" 

set objRs = objConn.execute(sql)

%>
<html>
<head>
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<title>Security Matrix</title>
<link href="stylesheets/styles.css" rel="stylesheet" type="text/css">
</head>
<script TYPE="TEXT/JAVASCRIPT">

	top.heading.frmPageTitle.PageTitle.value = "SMA - Security Matrix";

</script>
<BODY>
<H2>SMA2 Security Matrix

<table border="1" cellPadding="2" cellSpacing="0" width="100%">
<%Dim i%>
<thead>
	<tr>
		<th bgColor=white></th>
		<th colspan=<%=objRs.Fields.Count-1%> align=center>Security Roles</th>
	<tr>
		<%for i = 0 to objRs.Fields.Count -1 %>	
			<th <%if i > 0 then Response.Write "width=100"%> align="center" valign="bottom"><%=objRs.Fields(i).name%></th>
		<%next%>
	</tr>

</thead>
<tbody>
<%
Dim counter
counter = 0
while not objRs.EOF
	counter = counter + 1
%>
	<tr <%if counter Mod 2 = 0 then Response.Write "bgcolor=white" end if %>>
		<%for i = 0 to objRs.Fields.Count -1 %>	
			<td NOWRAP <%if i > 0 then Response.Write "width=100"%>><%=objRs(i)%>&nbsp;</td>
		<%next%>
	</tr>
<%

	objRS.moveNext
wend
%>
</tbody>
</table>
<BR><BR>
<table border=1 cellpadding=2 cellspacing=0>
	<thead>
		<tr>
			<th align="center" colspan="2">Legend</th>
		</tr>
		<tr>
			<th>Symbol</th>
			<th>Access Level</th>
		</tr>
	</thead>
	<tbody>
		<tr>
			<td>S</td>
			<td>Read Access - you can see the data</td>
		</tr>
		<tr>
			<td>U</td>
			<td>Update Access - you can change existing data</td>
		</tr>
		<tr>
			<td>I</td>
			<td>Insert Access - you can create new data</td>
		</tr>
		<tr>
			<td>D</td>
			<td>Delete Access - you can delete data</td>
		</tr>
	</tbody>
</table>
<%
objRS.close
set objRS = Nothing

objConn.close
Set objConn = Nothing
%>

</BODY>
</HTML>
