<%@ Language=VBScript %>
<% Option Explicit %>
<!--#include Virtual="/adovbs.inc"-->


<TABLE border = 1>
 <TR>
  <TD><B>CUSTOMER</B></TD>
  <TD><B>CUSTOMER_TYPE_IND</B></TD>
  <TD><B>CUSTOMER_NAME</B></TD>
 </TR>

<%
 'Open up a connection to our Access database
  'We will use a DSN - less connection

 Dim objConn,user,password,strConnectstring,strSQL,strCust
  user="smatest" 'request.form("username")
  password= "smatest" 'request.form("password")
  strCust=request.form("customer")

'Setup connection object
 strConnectstring = "DSN=orad3;uid=" &user &";pwd=" &password
 set objConn = Server.CreateObject("ADODB.Connection")

' With objConn
 objConn.ConnectionString = strConnectstring
 objConn.CursorLocation = adUseServer
 objConn.open
 'End with
'30
strSQL = "Packcustomer.GetCustomer"
' Response.write "entry=" & strCust


'Command Object
 Dim objRS,cmdCust,objParam,counter
  set cmdCust = Server.CreateObject("ADODB.Command")
  set cmdCust.ActiveConnection = objConn
      cmdCust.CommandText = strSQL
      cmdCust.CommandType = adCmdStoredProc    
 set objParam = cmdCust.CreateParameter
      objParam.Name = "error_code"
     objParam.Type = adInteger
      objParam.Direction = adParamOutput
      objParam.size = 10
      cmdCust.Parameters.Append objParam
  set objParam = cmdCust.CreateParameter
      objParam.Name = "error_desc"
      objParam.Type = adVarChar
      objParam.Direction = adParamOutput
     objParam.size = 255
      cmdCust.Parameters.Append objParam
 set objParam = cmdCust.CreateParameter
      objParam.Name = "c_out"
      objParam.Type = adVarchar
      objParam.Direction = adParamOutput
     objParam.size = 255
      cmdCust.Parameters.Append objParam
'50   
  'set Param = cmdCust.createparameter("o_cust_nm", 129, 1,,"Cabletron")
  'set Param = cmdCust.CreateParameter("o_cust_id", 131, 2)
  'set Param = cmdCust.CreateParameter("o_cust_ind", 200, 2)
  'set Param = cmdCust.CreateParameter("o_cust_name", 200, 2)
  'cmdCust.Parameters.append Param
 ' cmdCust(0) = strCust


  'Set up Recordset Object
 ' set objRS = Server.CreateObject("ADODB.Recordset")
  'with objRS
   'objRS.CursorType = adOpenStatic
   'objRS.LockType = adLockReadOnly
  'End with

   Set objRS = cmdCust.execute

IF objConn.errors.count > 0 THEN
    response.write "Database Errors Occured" & "<P>"
   FOR counter = 0 to objConn.errors.count
      response.write "Error #" & objConn.errors(counter).number & "<p>"
      response.write "Error Desc. -> " & objConn.errors(counter).description & "<p>"
      response.write "Error Source -> " & objConn.errors(counter).Source & "<p>"
   next
  else
     response.write strSQL
     response.write "Everthing went fine."
 END IF

  

 'with cmdCust
  
 'end with

'
  'set objRS.Source = cmdCust
  'cmdCust(0) = 16
  'objRS.Open
 'call ErrorVBScriptReport("Calling Procedure")
 'call ErrorVBScriptReport(strSQL,objConn)


 'Display the contents of the Friends table

  Do while Not objRS.EOF
    Response.write "<TR><TD>" &  objRS.Fields("customer_id").Value & "</TD>"
    Response.write  "<TD>" &  objRS.Fields("customer_type_ind")>Value & "</TD>"
    Response.write   "<TD>" &  objRS.Fields("customer_name").Value & "</TD>" 

   'Move to the next row in the Friends table
    objRS.MoveNext
  Loop

   'Response.write "Error Code:" & cmdCust("error_code")
   'Response.write "Error Description:" & cmdCust("error_desc")

   'Clean up our ADO objects
    objRS.close
    set objRS = Nothing

    set cmdCust = Nothing

    objConn.close
    set ObjConn = Nothing
%>

</TABLE>