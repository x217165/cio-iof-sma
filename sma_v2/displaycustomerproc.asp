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

 Dim objConn,user,password,strConnectstring,strSQL
  user=request.form("username")
  password=request.form("password")

'Setup connection object
 strConnectstring = "DSN=orad3;uid=" &user &";pwd=" &password
 set objConn = Server.CreateObject("ADODB.Connection")

' With objConn
 objConn.ConnectionString = strConnectstring
 objConn.CursorLocation = adUseServer
 objConn.open
 'End with

 'Define call to store procedure
  strSQL = "{call PackCustomer.ALLCustomer({resultset 10000,o_cust_id,o_cust_ind,o_cust_name})}"

'Command Object
 Dim objRS,cmdCust
 set cmdCust = Server.CreateObject("ADODB.Command")

 'with cmdCust
  set cmdCust.ActiveConnection = objConn
      cmdCust.CommandText = strSQL
      cmdCust.CommandType = adCmdText
 'end with

'Set up Recordset Object
  set objRS = Server.CreateObject("ADODB.Recordset")
  'with objRS
   objRS.CursorType = adOpenStatic
   objRS.LockType = adLockReadOnly
  'End with

  set objRS.Source = cmdCust
  objRS.Open


 'Create a recordset object instance and retrieve the information
 'from Friends table.
  

 'Display the contents of the Friends table

  Do while Not objRS.EOF
    Response.write "<TR><TD>" & objRS(0) & "</TD>"
    Response.write  "<TD>" & objRS(1) & "</TD>"
    Response.write   "<TD>" & objRS(2)  & "</TD>" 

   'Move to the next row in the Friends table
    objRS.MoveNext
 Loop

   'Clean up our ADO objects
    objRS.close
    set objRS = Nothing

    set cmdCust = Nothing

    objConn.close
    set ObjConn = Nothing
%>