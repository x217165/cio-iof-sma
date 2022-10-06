<%@  language="VBScript" %>

<% Response.Buffer = true %>
<!--#include file="smaConstants.inc"-->
<!--#include file="smaProcs.inc"-->
<!--#include file="databaseconnect.asp" -->
<%
    Function SimpleBinaryToString(Binary)
     
    Dim DecryptedData
    if (Binary <> ""and Binary <> Empty) then
   
SimpleBinaryToString = DecryptWithKey("Constant",Binary)
    
    
    else
     SimpleBinaryToString=""
   
    end if
   
End Function

dim NEId, strWinMessage ,dt , selDt
    dim strRealUserID
strRealUserID =  Session("username") 
NEId = Request.QueryString("NEId")
    
  if (NEId = null or NEId ="" or NEId =Empty) then
    NEId = Request("NeId")
    end if

    if (NEId = null or NEId ="" or NEId =Empty) then
    NEId = Request("txtNEID")
    end if
    
dim intAccessLevelForSNMP_write
   
intAccessLevelForSNMP_write = CInt(CheckLogon(strConst_SNMP_write))
     if intAccessLevelForSNMP_write and intConst_Access_ReadOnly <> intConst_Access_ReadOnly then
	DisplayError "BACK", "", 0, "ACCESS DENIED", "You don't have access to SNMPview object. Please contact your system administrator"
end if

    dim cmdViewObj,aList,rsAlias,isEditable,strSQL,sqlPriv,sqlAuth,sqlSecurity
	
     'isEditable = (intAccessLevelForSNMP_write > 0)
    dim rsPRIV,rsAuth,rsSecurity , rsDt ,tmp
sqlPriv = "select SNMP_PRIV_PROT_ID,SNMP_PRIV_PROT_Name from CRP.LCODE_SNMP_PRIV_PROT"
    sqlAuth = "select SNMP_AUTH_PROT_ID,SNMP_AUTH_PROT_NAME from CRP.LCODE_SNMP_AUTH_PROT"
    sqlSecurity = "select SNMP_SECURITY_LVL_ID,SNMP_SECURITY_LVL_NAME from CRP.LCODE_SNMP_SECURITY_LVL"
    selDt = "select to_char(sysdate,'DD-MON-RRRR HH24:MI:SS') as dt from dual"

set rsPRIV=server.CreateObject("ADODB.Recordset")
'rsPRIV.CursorLocation = adUseClient
rsPRIV.Open sqlPriv, objConn
    
    set rsDt=Server.CreateObject("ADODB.command")
    rsDt.ActiveConnection = objConn
'rsPRIV.CursorLocation = adUseClient
'rsDt.Open selDt, objConn
    rsDt.CommandText = selDt

    set tmp = rsDt.Execute(,,adCmdText)
   dt =  tmp.Fields(0).Value
  '  if not rsDt.EOF then
	'	dt = rsDt.GetRows(1,0,"dt")
  '  else
  ' dt =""
  ' end if
   
    set rsAuth=server.CreateObject("ADODB.Recordset")
'rsAuth.CursorLocation = adUseClient
rsAuth.Open sqlAuth, objConn

    set rsSecurity=server.CreateObject("ADODB.Recordset")
'rsSecurity.CursorLocation = adUseClient
rsSecurity.Open sqlSecurity, objConn

    
     isEditable = (intAccessLevelForSNMP_write >0) 
    
    select case Request("txtFrmAction")
    case ""

	set cmdViewObj = server.CreateObject("ADODB.recordset")

		 strSQL =  "select UPDATE_REAL_USERID,CREATE_DATE_TIME,CREATE_REAL_USERID,UPDATE_DATE_TIME, SNMP_STRING	,SNMP_V3_USERNAME	,SNMP_V3_ENGINEID	,SNMP_V3_CONTEXT_NAME	,SNMP_SECURITY_LVL_ID	,SNMP_AUTH_PROT_ID	,SNMP_PRIV_PROT_ID	,SNMP_V3_AUTH_KEY	,SNMP_V3_PRIV_KEY	,SNMP_PORT,SNMP_CRED_LEVEL from CRP.NETWORK_ELEMENT_SNMP where NETWORK_ELEMENT_ID  =" & NEId &" and  rownum <2"

    
    cmdViewObj.Open strSQL, objConn
	If err then
		DisplayError "BACK", "", err.Number, "SNMPview.asp - Cannot open database" , err.Description
	End if
	'put recordset into array
	'if not cmdViewObj.EOF then
	'	aList = cmdViewObj.GetRows(1,0)
   ' rsAlias = aList
	'else
		'Response.Write "0 Records Found"
		'Response.End
	'end if
   ' for each x in cmdViewObj.fields
 ' response.write(x.name)
 ' response.write(" = ")
 ' response.write(x.value)
'next

   ' cmdViewObj.Close
	'set cmdViewObj = nothing
	'objConn.Close
	'set objConn = nothing

	case "save"
     
   
    Dim SNMP_CRED_LEVEL
     Dim CanEncryptSNMP_CRED_LEVEL
    SNMP_CRED_LEVEL=Request("selSNMP_CRED_LEVEL")


     if(SNMP_CRED_LEVEL=Empty or SNMP_CRED_LEVEL="") then
    CanEncryptSNMP_CRED_LEVEL=false
    else
    CanEncryptSNMP_CRED_LEVEL=true
    end if

    dim SNMP_SECURITY_LVL_ID,SNMP_PRIV_PROT_ID,SNMP_AUTH_PROT_ID

    if Request("selAuthProtocol") <> "" then
     SNMP_AUTH_PROT_ID = CInt(Request("selAuthProtocol")) 
    else
    SNMP_AUTH_PROT_ID = "null"
    end if

    if Request("selSecurityLevel") <> "" then
     SNMP_SECURITY_LVL_ID = CInt(Request("selSecurityLevel")) 
    else
    SNMP_SECURITY_LVL_ID = "null"
    end if

      if Request("selPrivProtocol") <> "" then
     SNMP_PRIV_PROT_ID = CInt(Request("selPrivProtocol")) 
    else
    SNMP_PRIV_PROT_ID = "null"
    end if
            dim cmdUpdateObj,SNMPId
			set cmdUpdateObj = server.CreateObject("ADODB.Command")
			set cmdUpdateObj.ActiveConnection = objConn
			cmdUpdateObj.CommandType = adCmdText
    Dim SNMPString
     Dim CanEncryptSNMPString
    SNMPString=Request("txtSNMPString")
    if(SNMPString=Empty or SNMPString="") then
    CanEncryptSNMPString=false
    else
    CanEncryptSNMPString=true
    end if


    Dim SNMPUserName
    Dim CanEncryptSNMPUserName
    SNMPUserName=Request("txtSNMP_V3_USERNAME")
     if(SNMPUserName=Empty or SNMPUserName="") then
    CanEncryptSNMPUserName=false
    else
    CanEncryptSNMPUserName=true
    end if

    Dim SNMP_V3_ENGINEID
     Dim CanEncryptSNMP_V3_ENGINEID
    SNMP_V3_ENGINEID=Request("txtSNMP_V3_ENGINEID")
     if(SNMP_V3_ENGINEID=Empty or SNMP_V3_ENGINEID="") then
    CanEncryptSNMP_V3_ENGINEID=false
    else
    CanEncryptSNMP_V3_ENGINEID=true
    end if

    Dim SNMP_V3_CONTEXT_NAME
     Dim CanEncryptSNMP_V3_CONTEXT_NAME
    SNMP_V3_CONTEXT_NAME=Request("txtSNMP_V3_CONTEXT_NAME")
     if(SNMP_V3_CONTEXT_NAME=Empty or SNMP_V3_CONTEXT_NAME="") then
    CanEncryptSNMP_V3_CONTEXT_NAME=false
    else
    CanEncryptSNMP_V3_CONTEXT_NAME=true
    end if

    Dim SNMP_V3_SEC_LEVEL
     Dim CanEncryptSNMP_V3_SEC_LEVEL
    SNMP_V3_SEC_LEVEL=Request("txtSNMP_V3_SEC_LEVEL")
     if(SNMP_V3_SEC_LEVEL=Empty or SNMP_V3_SEC_LEVEL="") then
    CanEncryptSNMP_V3_SEC_LEVEL=false
    else
    CanEncryptSNMP_V3_SEC_LEVEL=true
    end if

    Dim SNMP_V3_AUTH_PROTOCOL
     Dim CanEncryptSNMP_V3_AUTH_PROTOCOL
    SNMP_V3_AUTH_PROTOCOL=Request("txtSNMP_V3_AUTH_PROTOCOL")
     if(SNMP_V3_AUTH_PROTOCOL=Empty or SNMP_V3_AUTH_PROTOCOL="") then
    CanEncryptSNMP_V3_AUTH_PROTOCOL=false
    else
    CanEncryptSNMP_V3_AUTH_PROTOCOL=true
    end if

    Dim SNMP_V3_PRIV_PROTOCOL
     Dim CanEncryptSNMP_V3_PRIV_PROTOCOL
    SNMP_V3_PRIV_PROTOCOL=Request("txtSNMP_V3_PRIV_PROTOCOL")
     if(SNMP_V3_PRIV_PROTOCOL=Empty or SNMP_V3_PRIV_PROTOCOL="") then
    CanEncryptSNMP_V3_PRIV_PROTOCOL=false
    else
    CanEncryptSNMP_V3_PRIV_PROTOCOL=true
    end if

    Dim SNMP_V3_AUTH_KEY
     Dim CanEncryptSNMP_V3_AUTH_KEY
    SNMP_V3_AUTH_KEY=Request("txtSNMP_V3_AUTH_KEY")
     if(SNMP_V3_AUTH_KEY=Empty or SNMP_V3_AUTH_KEY="") then
    CanEncryptSNMP_V3_AUTH_KEY=false
    else
    CanEncryptSNMP_V3_AUTH_KEY=true
    end if

    Dim SNMP_V3_PRIV_KEY
     Dim CanEncryptSNMP_V3_PRIV_KEY
    SNMP_V3_PRIV_KEY=Request("txtSNMP_V3_PRIV_KEY")
     if(SNMP_V3_PRIV_KEY=Empty or SNMP_V3_PRIV_KEY="") then
    CanEncryptSNMP_V3_PRIV_KEY=false
    else
    CanEncryptSNMP_V3_PRIV_KEY=true
    end if

    Dim SNMP_PORT
     Dim CanEncryptSNMP_PORT
    SNMP_PORT=Request("txtSNMP_PORT")
     if(SNMP_PORT=Empty or SNMP_PORT="") then
    CanEncryptSNMP_PORT=false
    else
    CanEncryptSNMP_PORT=true
    end if
    




                      if(Request("hdnUpdate") <> "" and Request("hdnUpdate") ="True") then
    
   
	 cmdUpdateObj.CommandText = "Update CRP.NETWORK_ELEMENT_SNMP set SNMP_STRING = '"& EncryptWithKey("Constant", SNMPString)& "',SNMP_V3_USERNAME ='"& EncryptWithKey("Constant", SNMPUserName) &"',SNMP_V3_ENGINEID   = '"& EncryptWithKey("Constant",SNMP_V3_ENGINEID) & "' ,SNMP_V3_CONTEXT_NAME ='" & EncryptWithKey("Constant",  SNMP_V3_CONTEXT_NAME)& "'  ,SNMP_SECURITY_LVL_ID ="& SNMP_SECURITY_LVL_ID & "  ,SNMP_AUTH_PROT_ID =" & SNMP_AUTH_PROT_ID  &", SNMP_PRIV_PROT_ID = "& SNMP_PRIV_PROT_ID & "  ,SNMP_V3_AUTH_KEY = '" &  EncryptWithKey("Constant",  SNMP_V3_AUTH_KEY) 
     cmdUpdateObj.CommandText = cmdUpdateObj.CommandText + "'  ,SNMP_V3_PRIV_KEY ='" & EncryptWithKey("Constant", SNMP_V3_PRIV_KEY) & "'  ,SNMP_PORT = '" & SNMP_PORT &"'  , SNMP_CRED_LEVEL ='" &  SNMP_CRED_LEVEL 
     cmdUpdateObj.CommandText = cmdUpdateObj.CommandText + "', UPDATE_REAL_USERID = '" &  strRealUserID &"' , UPDATE_DATE_TIME = to_date('" & dt &"', 'dd-mon-yyyy hh24:mi:ss')  where NETWORK_ELEMENT_ID  =" & NEId 
                        else
    set cmdViewObj = server.CreateObject("ADODB.recordset")
	
		 strSQL =  "select max(SNMP_ID) as id from CRP.NETWORK_ELEMENT_SNMP "

    cmdViewObj.Open strSQL, objConn
    if(IsNull(cmdViewObj.Fields("ID")  )) then
    SNMPId =1
    else
    SNMPId = cmdViewObj.Fields("ID") 
    end if
     
                          '  cmdUpdateObj.CommandText = "insert into CRP.NETWORK_ELEMENT_SNMP (NETWORK_ELEMENT_ID,SNMP_STRING,SNMP_V3_USERNAME,SNMP_V3_ENGINEID	,SNMP_V3_CONTEXT_NAME ,SNMP_SECURITY_LVL_ID  ,SNMP_AUTH_PROT_ID ,SNMP_V3_PRIV_PROTOCOL ,SNMP_V3_AUTH_KEY ,SNMP_PRIV_PROT_ID ,SNMP_PORT, SNMP_CRED_LEVEL) values( "&  NEId & ",'"&EncryptWithKey("Constant",  Request("txtSNMPString")) & "','" &EncryptWithKey("Constant",  Request("txtSNMP_V3_USERNAME")) & "', '"&EncryptWithKey("Constant",  Request("txtSNMP_V3_ENGINEID")) & "','" &EncryptWithKey("Constant",  Request("txtSNMP_V3_CONTEXT_NAME")) &"', '"&EncryptWithKey("Constant",  Request("txtSNMP_V3_SEC_LEVEL")) & "','" &EncryptWithKey("Constant",  Request("txtSNMP_V3_AUTH_PROTOCOL")) &"', "&  CInt(  Request("selSecurityLevel") )& "," & CInt(Request("selAuthProtocol") )&", " & CInt(Request("selPrivProtocol") ) & ",'"&Request("txtSNMP_PORT")&"', '" &EncryptWithKey("Constant",  Request("txtSNMP_CRED_LEVEL")) &"')"
    cmdUpdateObj.CommandText ="insert into CRP.NETWORK_ELEMENT_SNMP ( NETWORK_ELEMENT_ID, SNMP_STRING, SNMP_V3_USERNAME, SNMP_V3_ENGINEID ,"
   cmdUpdateObj.CommandText = cmdUpdateObj.CommandText+ " SNMP_V3_CONTEXT_NAME , SNMP_V3_AUTH_KEY , SNMP_V3_PRIV_KEY , SNMP_PORT, SNMP_CRED_LEVEL, SNMP_SECURITY_LVL_ID , SNMP_AUTH_PROT_ID , SNMP_PRIV_PROT_ID ,CREATE_REAL_USERID ,CREATE_DATE_TIME ,SNMP_ID ) values( "
    cmdUpdateObj.CommandText = cmdUpdateObj.CommandText + NEId & ", '"&EncryptWithKey("Constant", Request("txtSNMPString")) & "', '" &EncryptWithKey("Constant", Request("txtSNMP_V3_USERNAME")) 
    cmdUpdateObj.CommandText = cmdUpdateObj.CommandText + "', '"&EncryptWithKey("Constant", Request("txtSNMP_V3_ENGINEID")) & "', '" &EncryptWithKey("Constant", Request("txtSNMP_V3_CONTEXT_NAME")) 
    cmdUpdateObj.CommandText = cmdUpdateObj.CommandText + "', '" &EncryptWithKey("Constant", Request("txtSNMP_V3_AUTH_KEY")) &"', '" &EncryptWithKey("Constant", Request("txtSNMP_V3_PRIV_KEY")) &"', '"
  cmdUpdateObj.CommandText = cmdUpdateObj.CommandText +  Request("txtSNMP_PORT")&"', '" &  SNMP_CRED_LEVEL &"'," & SNMP_SECURITY_LVL_ID& ", " & SNMP_AUTH_PROT_ID &", " & SNMP_PRIV_PROT_ID & ",'" &  strRealUserID &"', to_date('" & dt &"', 'dd-mon-yyyy hh24:mi:ss')"&"," &( CInt( SNMPId) +1)  &" )"

                        end if
   
                           cmdUpdateObj.Execute

    			        if err then
				        if instr(1, objConn.Errors(0).Description, "ORA-20040" ) then
				        	dim strWinLocation
				        	strWinLocation = "SNMPview.asp?NEId="&Request("txtNEID")
				        	DisplayError "REFRESH", strWinLmgocation, objConn.Errors(0).NativeError, "OBJECT UPDATED", objConn.Errors(0).Description
				        else
				        	DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
				        end if
    end if
				        objConn.Errors.Clear
			
	
		
	case "DELETE"
		'delete record
		
			dim cmdDeleteObj
			set cmdDeleteObj = server.CreateObject("ADODB.Command")
			set cmdDeleteObj.ActiveConnection = objConn
			cmdDeleteObj.CommandType = adCmdText
			cmdDeleteObj.CommandText = "delete from CRP.NETWORK_ELEMENT_SNMP  where NETWORK_ELEMENT_ID  =" & NEId 
			cmdDeleteObj.Execute
			
            if objConn.Errors.Count <> 0 then
				DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
				objConn.Errors.Clear
			end if
			strNE_ID = ""
			strWinMessage = "Record deleted successfully."
	end select
    		
    set cmdViewObj = server.CreateObject("ADODB.recordset")
	
		 strSQL =  "select * from CRP.NETWORK_ELEMENT_SNMP where NETWORK_ELEMENT_ID  =" & NEId &" and  rownum <2"

    cmdViewObj.Open strSQL, objConn	
%>


<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <link rel="stylesheet" type="text/css" href="stylesheets/styles.css">
    <script type="text/javascript" src="AccessLevels.js"></script>
    <script type="text/javascript" src="GeneralJavaFunctions.js"></script>

    <script language="javascript">
<!-- SNMP_STRING	,SNMP_V3_USERNAME	,SNMP_V3_ENGINEID	,SNMP_V3_CONTEXT_NAME	,SNMP_V3_SEC_LEVEL	,SNMP_V3_AUTH_PROTOCOL	,SNMP_V3_PRIV_PROTOCOL	,SNMP_V3_AUTH_KEY	,SNMP_V3_PRIV_KEY	,SNMP_PORT,SNMP_CRED_LEVEL
    //******************************************** End of Java Functions *****************************
    //-->
    //var intAccessLevelForSNMP_write = <% CStr(intAccessLevelForSNMP_write)%>;


    function btn_onDelete() {

        if (confirm('Do you really want to delete this object?')) {
            //submit the form
            document.frmSNMP.txtFrmAction.value = "DELETE";
            document.frmSNMP.submit();
        }
    }

    function btnClose_onclick() {
        window.close();
    }

    function frmAlias_onsubmit() {
        if (document.frmSNMP.txtNEID.value != "") {
            document.frmSNMP.txtFrmAction.value = "save";
            document.frmSNMP.submit();
            return (true);

        }
        else { alert('You cannot save an empty alias. Please click DELETE button if you want to delete the alias.'); return (false); }
    }

    function btnReset_onclick() {
        document.frmSNMP.txtFrmAction.value = "";
        document.frmSNMP.NeId = document.frmSNMP.txtNEID;
        document.frmSNMP.submit();

    }



    </script>
</head>
<body>
    <form name="frmSNMP" language="javascript" onsubmit="return frmAlias_onsubmit()">
        <input type="hidden" name="txtFrmAction" value="" />
        <input type="hidden" name="txtNEID" value="<%if (NEId <> "") then Response.write NEId  end if %>" />
        <input type="hidden" name="hdnUpdate" value="<%if (cmdViewObj.EOF <> true) then Response.write true else Response.write false  end if %>" />

        <table border="0" width="100%">
            <thead>
                <tr>
                    <th colspan="4">SNMP DEVICE INFORMATION</th>
                </tr>
            </thead>

            <tbody>
                <tr>
                    <td align="left" nowrap>SNMP STRING </td>
                    <td align="left">
                        <input size='40' maxlength='30' name='txtSNMPString' value='<%if (cmdViewObj.EOF <> true) then Response.write SimpleBinaryToString(cmdViewObj.fields("SNMP_STRING").Value) else Response.write ""  end if %>' <%if isEditable = false then  Response.write("disabled=""disabled""") end if%>>
                    </td>
                    <td align="left" nowrap>SNMP V3_USERNAME </td>
                    <td align="left">
                        <input size='40' maxlength='30' name='txtSNMP_V3_USERNAME' value='<%if (cmdViewObj.EOF <> true) then Response.Write  SimpleBinaryToString(cmdViewObj.fields("SNMP_V3_USERNAME").Value)  else Response.write ""  end if  %>' <%if isEditable = false then  Response.write("disabled=""disabled""") end if%>>
                    </td>


                </tr>

                <tr>

                    <td align="left" nowrap>SNMP V3 ENGINEID </td>
                    <td align="left">
                        <input size='40' maxlength='30' name='txtSNMP_V3_ENGINEID' value='<%if (cmdViewObj.EOF <> true) then Response.Write  SimpleBinaryToString(cmdViewObj.fields("SNMP_V3_ENGINEID").Value) else Response.write ""  end if %>' <%if isEditable = false then  Response.write("disabled=""disabled""") end if%>>
                    </td>

                    <td align="left" nowrap>SNMP V3 CONTEXT_NAME </td>
                    <td align="left">
                        <input size='40' maxlength='30' name='txtSNMP_V3_CONTEXT_NAME' value='<%if (cmdViewObj.EOF <> true) then Response.Write  SimpleBinaryToString(cmdViewObj.fields("SNMP_V3_CONTEXT_NAME").Value) else Response.write ""  end if  %>' <%if isEditable = false then  Response.write("disabled=""disabled""") end if%>>
                    </td>
                </tr>

                <tr>

                    <td align="left" nowrap>SNMP V3 SEC LEVEL </td>
                    <td align="left">

                        <select id="selSecurityLevel" name="selSecurityLevel" <%if isEditable = false then  Response.write("disabled=""disabled""") end if%>>
                            <option value=""></option>
                            <%    
                              
			        while not rsSecurity.EOF
					Response.Write "<OPTION"
					if cmdViewObj.EOF <> true then
                            If IsNull( cmdViewObj.fields("SNMP_SECURITY_LVL_ID").Value) <> true and IsEmpty( cmdViewObj.fields("SNMP_SECURITY_LVL_ID").Value) <> true then
                                 if CLng(rsSecurity("SNMP_SECURITY_LVL_ID").Value) = CLng(cmdViewObj.fields("SNMP_SECURITY_LVL_ID").Value) then Response.write " selected"
                                end if
                                end if
					   Response.write " value=" & rsSecurity("SNMP_SECURITY_LVL_ID").Value  & ">" & rsSecurity("SNMP_SECURITY_LVL_NAME").Value & "</option>" &vbCrLf
					rsSecurity.MoveNext
				wend
				
                                rsSecurity.Close
                            %>
                        </select>
                    </td>

                    <td align="left" nowrap>SNMP V3 AUTH PROTOCOL </td>
                    <td align="left">


                        <select id="selAuthProtocol" name="selAuthProtocol" <%if isEditable = false then  Response.write("disabled=""disabled""") end if%>>
                            <option value=""></option>
                            <%    

                                
			        while not rsAuth.EOF
					Response.Write "<OPTION"
					if cmdViewObj.EOF <> true then
                                     if  IsNull( cmdViewObj.fields("SNMP_AUTH_PROT_ID").Value)  <> true  and IsEmpty( cmdViewObj.fields("SNMP_AUTH_PROT_ID").Value) <> true then
                                      if CLng(rsAuth("SNMP_AUTH_PROT_ID").Value) = CLng(cmdViewObj.fields("SNMP_AUTH_PROT_ID").Value) then Response.write " selected"
                                       end if
                                end if
					   Response.write " value=" & rsAuth("SNMP_AUTH_PROT_ID").Value  & ">" & rsAuth("SNMP_AUTH_PROT_NAME").Value & "</option>" &vbCrLf
					rsAuth.MoveNext
				wend
				
                                rsAuth.Close
                            %>
                        </select>
                    </td>
                </tr>

                <tr>

                    <td align="left" nowrap>SNMP V3 PRIV PROTOCOL </td>
                    <td align="left">
                        <select name="selPrivProtocol" <%if isEditable = false then  Response.write("disabled=""disabled""") end if%>>
                            <option value=""></option>
                            <%
				if rsPRIV.EOF <> true then
					
				while not rsPRIV.EOF
					Response.Write "<OPTION"
					if cmdViewObj.EOF <> true then
                                if  IsNull( cmdViewObj.fields("SNMP_PRIV_PROT_ID").Value )  <> true and IsEmpty( cmdViewObj.fields("SNMP_PRIV_PROT_ID").Value ) <> true then 
                                if CLng(cmdViewObj.fields("SNMP_PRIV_PROT_ID").Value )= CLng(rsPRIV("SNMP_PRIV_PROT_ID").Value)  then Response.write " selected"
                                  end if
                                end if
					  Response.Write " value ='" & rsPRIV("SNMP_PRIV_PROT_ID").Value  & "' "
                                Response.Write ">" & rsPRIV("SNMP_PRIV_PROT_NAME").Value  & "</OPTION>"
					rsPRIV.MoveNext
				wend
				rsPRIV.Close
				end if
                            %>
                        </select>
                    </td>

                    <td align="left" nowrap>SNMP V3 AUTH KEY </td>
                    <td align="left">
                        <input size='40' maxlength='30' name='txtSNMP_V3_AUTH_KEY' value='<%if (cmdViewObj.EOF <> true) then Response.Write  SimpleBinaryToString(cmdViewObj.fields("SNMP_V3_AUTH_KEY").Value) else Response.write ""  end if %>' <%if isEditable = false then  Response.write("disabled=""disabled""") end if%>>
                    </td>
                </tr>

                <tr>

                    <td align="left" nowrap>SNMP V3 PRIV KEY </td>
                    <td align="left">
                        <input size='40' maxlength='30' name='txtSNMP_V3_PRIV_KEY' value='<%if (cmdViewObj.EOF <> true) then Response.Write  SimpleBinaryToString(cmdViewObj.fields("SNMP_V3_PRIV_KEY").Value) else Response.write ""  end if  %>' <%if isEditable = false then  Response.write("disabled=""disabled""") end if%>>
                    </td>

                    <td align="left" nowrap>SNMP PORT </td>
                    <td align="left">
                        <input size='40' maxlength='30' name='txtSNMP_PORT' value='<%if (cmdViewObj.EOF <> true) then Response.Write cmdViewObj.fields("SNMP_PORT").Value else Response.write ""  end if %>' <%if isEditable = false then  Response.write("disabled=""disabled""") end if%>>
                    </td>
                </tr>


                <tr>

                    <td align="left" nowrap>SNMP CRED LEVEL </td>
                    <td align="left">
                        <select name="selSNMP_CRED_LEVEL">
                            <option value=""></option>
                            <option value="Read" <% if (cmdViewObj.EOF <> true) then if cmdViewObj.fields("SNMP_CRED_LEVEL").Value = "Read" then Response.write " selected" %>>Read</option>
                            <option value="Write" <% if (cmdViewObj.EOF <> true) then if cmdViewObj.fields("SNMP_CRED_LEVEL").Value = "Write" then Response.write " selected" %>>Write</option>
                        </select>
                    </td>

                </tr>
            </tbody>

            <tfoot>
                <tr>
                    <td align="left" colspan="5">
                        <input type="button" name="btnClose" value="Close" style="width: 2cm" onclick="return btnClose_onclick();">&nbsp;&nbsp;
                        <% if isEditable then Response.Write "<input type='button' name='btnDelete' value='Delete' style='width: 2cm' onclick='return btn_onDelete();'>&nbsp;&nbsp;	  	<input type='button' name='btnReset' value='Reset' style='width: 2cm' onclick='return btnReset_onclick();'>&nbsp;&nbsp;	  	<input type='button' name='btnSave' value='Save' style='width: 2cm' onclick='return frmAlias_onsubmit();'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" end if %>
	  	
                    </td>
                </tr>
            </tfoot>
        </table>
        <fieldset>
            <!-- <%if bolClone then strNE_ID = ""%>-->
            <legend align="right"><b>Audit Information</b></legend>
            <div size="8pt" align="RIGHT">
                Create Date&nbsp;<input align="center" name="txtCreateDateTime" type="text" style="height: 20px; width: 150px" disabled value="<% if (cmdViewObj.EOF <> true) then Response.write  cmdViewObj.fields("CREATE_DATE_TIME").Value end if %>">&nbsp;
		Created By&nbsp;
                <input align="right" name="txtCreateRealUser" type="text" style="height: 20px; width: 200px" disabled value="<% if (cmdViewObj.EOF <> true) then Response.write  cmdViewObj.fields("CREATE_REAL_USERID").Value end if %>"><br>
                Update Date&nbsp;<input align="center" name="txtUpdateDateTime" type="text" style="height: 20px; width: 150px" disabled value="<% if (cmdViewObj.EOF <> true)then Response.write  cmdViewObj.fields("UPDATE_DATE_TIME").Value end if %>">
                Updated By&nbsp;
                <input align="right" name="txtUpdateRealUser" type="text" style="height: 20px; width: 200px" disabled value="<% if (cmdViewObj.EOF <> true) then Response.write  cmdViewObj.fields("UPDATE_REAL_USERID").Value end if %>">
            </div>
        </fieldset>

    </form>
</body>
</html>

