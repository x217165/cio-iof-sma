<%@  language="VBScript" %>
<% Option Explicit %>
<% on error resume next %>
<!--<% Response.Buffer = true %>-->
<!--#include file="sma_env.inc"-->
<!--#include file="smaConstants.inc"-->
<!--#include file="databaseconnect.asp"-->
<!--#include file="smaProcs.inc"-->

<style type="text/css">
    .auto-style1 {
        width: 58%;
    }

    .auto-style2 {
        height: 21px;
    }

    .Highlight {
        cursor: hand;
        background-color: #00974f;
        color: white;
    }
</style>

<%
     
     
dim CId, strWinMessage,OrgId,OrgType
CId = Request.QueryString("CId")
OrgId = Request.QueryString("OrgId")
    OrgType = Request.QueryString("OrgType")

    if CId = "Empty" or  CId = "" or CId = null   then
    CId = Request("CID")
    end if
    
dim intAccessLevelForCOM_write
    
     
intAccessLevelForCOM_write = CInt(CheckLogon(strConst_COM_write))
      
    dim cmdViewObj,aList,isEditable,isExists,rsGrid,sqlGrid,rsmaxCount,sqlMaxCount
	isEditable = (intAccessLevelForCOM_write >0)

   'select case Request("txtFrmAction")
    'case ""

	set cmdViewObj = server.CreateObject("ADODB.Recordset")
	set rsGrid = server.CreateObject("ADODB.Recordset")
    set rsmaxCount = server.CreateObject("ADODB.Recordset")
		dim strSQL
    
    strSQL =  "select  COALESCE(ORGANIZATION_NAME, '') as ORGANIZATION_NAME , COALESCE(ORGANIZATION_CODE, '') as ORGANIZATION_CODE  from CRP.CUSTOMER_ORGANIZATION where CUSTOMER_ID  =" & CId &" and  rownum <2"
    sqlGrid = " select ORGANIZATION_ID,COALESCE(ORGANIZATION_NAME, '') as ORGANIZATION_NAME , COALESCE(ORGANIZATION_CODE, '')  as ORGANIZATION_CODE from CRP.CUSTOMER_ORGANIZATION where CUSTOMER_ID  =" & CId 
    sqlMaxCount = " select max(ORGANIZATION_ID)+1 as count from CRP.CUSTOMER_ORGANIZATION"
   ' Response.Write("----"& strSQL)
    cmdViewObj.Open strSQL, objConn
    rsGrid.Open sqlGrid, objConn
    
	If err then
		DisplayError "BACK", "", err.Number, "CMO.asp - Cannot open database" , err.Description
	End if

   ' Response.Write("----"& cmdViewObj.Rows.Count)
	'put recordset into array
	if not cmdViewObj.EOF then
		aList = cmdViewObj.GetRows(1,0)
   ' rsAlias = aList
    isExists=true
	else
		isExists=false
		
	end if

    cmdViewObj.Close
	set cmdViewObj = nothing
	
    
    select case Request("txtFrmAction")
    
	case "SAVE"
		if (intAccessLevel and intConst_Access_Update <> intConst_Access_Update) then
				DisplayError "BACK", "", 0, "UPDATE DENIED", "You don't have access to update managed objects. Please contact your system administrator" & intAccessLevel & "dv"&  intConst_Access_Update
			end if
            dim cmdUpdateObj,ORGANIZATIONNAME, ORGANIZATIONCODE

   ' alert(Request("txtORGANIZATION_NAME"))
    
			set cmdUpdateObj = server.CreateObject("ADODB.Command")
			set cmdUpdateObj.ActiveConnection = objConn
			cmdUpdateObj.CommandType = adCmdText
     
      ORGANIZATIONNAME=Replace( Request("txtORGANIZATION_NAME"),"_"," ")
            
     ORGANIZATIONCODE= Replace( Request("txtORGANIZATION_CODE"),"_"," ")
    rsmaxCount.Open sqlMaxCount, objConn
     
                      
              if(OrgType ="UPDATE"  ) then
           
			                cmdUpdateObj.CommandText = "Update CRP.CUSTOMER_ORGANIZATION set ORGANIZATION_NAME ='"&  ORGANIZATIONNAME &"' , ORGANIZATION_CODE ='"&  ORGANIZATIONCODE &"' where ORGANIZATION_ID  ="& OrgId
                        else
                            cmdUpdateObj.CommandText = "insert into CRP.CUSTOMER_ORGANIZATION (CUSTOMER_ID,ORGANIZATION_NAME,ORGANIZATION_CODE,ORGANIZATION_ID) values("&  CId & ", '"&  ORGANIZATIONNAME & "','" &  ORGANIZATIONCODE &"',"&rsmaxCount("count") &")"
                       end if
    
                           cmdUpdateObj.Execute
    isExists = true
    
    			        if err then
				        if instr(1, objConn.Errors(0).Description, "ORA-20040" ) then
				        	dim strWinLocation
				        	strWinLocation = "COM.asp?CId="&Request("txtCID")
				        	DisplayError "REFRESH", strWinLocation, objConn.Errors(0).NativeError, "OBJECT UPDATED", objConn.Errors(0).Description
				        else
				        	DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT UPDATE OBJECT", objConn.Errors(0).Description
				        end if
    end if
				        objConn.Errors.Clear

   ' set strSQL =  "select ORGANIZATION_NAME,ORGANIZATION_CODE from CRP.CUSTOMER_ORGANIZATION where CUSTOMER_ID  =" & CId &" and  rownum <2"

			strWinMessage = "Record saved successfully. You can now see the changes you made."
    
	 
	case "DELETE"
		'delete record
     
		if (intAccessLevel and intConst_Access_Delete <> intConst_Access_Delete) then
			DisplayError "BACK", "", 0, "DELETE DENIED", "You don't have access to delete managed objects. Please contact your system administrator" & intAccessLevel& "zcfv" & intConst_Access_Delete
		end if
   
			dim cmdDeleteObj,strRealUserID
			set cmdDeleteObj = server.CreateObject("ADODB.Command")
			set cmdDeleteObj.ActiveConnection = objConn
			cmdDeleteObj.CommandType = adCmdText
         strRealUserID = Session("username") 
            cmdDeleteObj.CommandText = "update CRP.CUSTOMER_ORGANIZATION set update_real_userid ='"& strRealUserID &"' where ORGANIZATION_ID ="& OrgId &""
          cmdDeleteObj.Execute

          cmdDeleteObj.CommandText =   "delete from CRP.CUSTOMER_ORGANIZATION  where ORGANIZATION_ID   =" & OrgId 
			cmdDeleteObj.Execute
			
            if objConn.Errors.Count <> 0 then
				'DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description
                DisplayError "BACK", "", objConn.Errors(0).NativeError, "CANNOT DELETE OBJECT", objConn.Errors(0).Description   
				'objConn.Errors.Clear
			end if
			strNE_ID = ""
			strWinMessage = "Record deleted successfully."
    isExists = false
   end select
    			
    set rsGrid = server.CreateObject("ADODB.Recordset")
    rsGrid.Open sqlGrid, objConn

   ' set cmdViewObj = server.CreateObject("ADODB.Recordset")
   ' Response.Write("----"& strSQL)
   ' cmdViewObj.Open strSQL, objConn

   '' if not cmdViewObj.EOF then
		'aList = cmdViewObj.GetRows(1,0)
   ' rsAlias = aList
   ' else
  '  rsAlias = Nothing
		

   ' end if
%>


<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <link rel="stylesheet" type="text/css" href="stylesheets/styles.css">
    <script type="text/javascript" src="AccessLevels.js"></script>
    <script type="text/javascript" src="GeneralJavaFunctions.js"></script>

    <script language="javascript">


        function btn_onDelete() {


            if (confirm('Do you really want to delete this object?')) {
                //submit the form
                document.frmMODetails.txtFrmAction.value = "DELETE";
                document.frmMODetails.submit();
            }
        }

        function btn_onReset() {
            document.frmMODetails.txtFrmAction.value = "RESET";
            document.frmMODetails.submit();
        }

        function btnClose_onclick() {
            window.close();
        }

        function frmAlias_onsubmit() {

            var strMasterID = "<%=strCustomerID%>";
            // NewWin = window.open("CustAliasDetail.asp?action=update&aliasID=" + strAliasID + "&masterID=" + strMasterID, "NewWin", "toolbar=no,status=yes,width=700px,height=175px,left=150px,top=200,menubar=no,resize=no");
            // NewWin.focus();

            document.frmMODetails.txtFrmAction.value = "SAVE";
            document.frmMODetails.submit();
            return (true);

            // }
            //else { alert('Access denied. Please contact your system administrator.'); return (false); }
        }



        function OnButtonClick(type) {
            document.frmMODetails.txtFrmAction.value = type;
            document.frmMODetails.OrgType.value = type;
            if (type == "ADD") {
                document.frmMODetails.OrgId.value = "";
                document.frmMODetails.txtORGANIZATION_NAME.value = "";
                document.frmMODetails.txtORGANIZATION_CODE.value = "";
            }
            if (type == "UPDATE" || type == "ADD") {
                document.getElementById("tblOrganization").style.display = 'block';
                return false;
            }
            else {
                document.frmMODetails.submit();
            }
        }
        var oldHighlightedElement;
        var oldHighlightedElementClassName;

        function cell_onClick(organizationId, name, code) {

            document.frmMODetails.OrgId.value = organizationId;
            document.frmMODetails.txtORGANIZATION_NAME.value = name.replace("_", " ");
            document.frmMODetails.txtORGANIZATION_CODE.value = code.replace("_", " ");
            //highlight current record
            if (oldHighlightedElement != null)
            { oldHighlightedElement.className = oldHighlightedElementClassName }
            oldHighlightedElement = window.event.srcElement.parentElement;
            oldHighlightedElementClassName = oldHighlightedElement.className;
            oldHighlightedElement.className = "Highlight";
        }


    </script>
</head>
<body>
    <form name="frmMODetails" language="javascript">
        <input type="hidden" name="txtFrmAction" value="" />
        <input type="hidden" name="CID" value="<%Response.Write CId %>" />
        <input type="hidden" name="OrgId" value="" />
        <input type="hidden" name="OrgType" value="" />
        <!--<tr onclick='OnOrgclick(3335,abcd,1234)'>
             <td>abcd</td><td>1234</td></tr>
             <tr onclick='OnOrgclick(3334,abcd,1234)'><td>abcd</td><td>1234</td></tr>-->
        <table border="1" width="100%">
            <thead>
                <tr>
                    <th class="auto-style2">Organization Name</th>
                    <th class="auto-style2">Organization Code</th>


                </tr>
            </thead>


            <% while not rsGrid.EOF
                    dim onClick 
                if IsNull( rsGrid("ORGANIZATION_CODE").Value)  <> true then
                     onClick = "'"& routineHtmlString(rsGrid("ORGANIZATION_ID")) &"','"&routineHtmlString(Replace(rsGrid("ORGANIZATION_NAME")," ","_"))&"','"&routineHtmlString(Replace(rsGrid("ORGANIZATION_CODE")," ","_")) &"'"
                 else
                 onClick = "'"& routineHtmlString(rsGrid("ORGANIZATION_ID")) &"','"&routineHtmlString(Replace(rsGrid("ORGANIZATION_NAME")," ","_"))&"',''"
                end if
					

                 if IsNull( rsGrid("ORGANIZATION_CODE").Value)  <> true then 
                Response.Write "<tr><td onclick=cell_onClick("&onClick&")>" & routineHtmlString(Replace(rsGrid("ORGANIZATION_NAME")," "," ")) & "</td><td onclick=cell_onClick("&onClick&")>" & routineHtmlString(Replace(rsGrid("ORGANIZATION_CODE")," "," ")) & "</td></tr>"
					else
                 Response.Write "<tr><td onclick=cell_onClick("&onClick&")>" & routineHtmlString(Replace(rsGrid("ORGANIZATION_NAME")," "," ")) & "</td><td><label>&nbsp;</label></td></tr>"
                end if
					rsGrid.MoveNext

				wend
				rsGrid.Close  %>

            <tfoot>
                <tr>
                    <td colspan="3">
                        <button style="padding: 5px; padding-left: 30px;" type="button" onclick="OnButtonClick('DELETE')">Delete</button>

                        <button style="padding: 5px;" type="button" onclick="OnButtonClick('UPDATE')">Update</button>
                        <button style="padding: 5px;" type="button" onclick="OnButtonClick('ADD')">Add</button>
                        <button style="padding: 5px;" type="button" onclick="return btnClose_onclick();">Close</button>

                    </td>
                </tr>
            </tfoot>
        </table>
        <table border="0" width="100%" id="tblOrganization" style="display: none;">
            <thead>
                <tr>
                    <th colspan="4" class="auto-style2">CUSTOMER ORGANIZATION MAINTENANCE</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td align="RIGHT" width="20%" nowrap>ORGANIZATION NAME
                          
                        <%  if (isEditable) then if ( isExists) then Response.Write "<input size='40' maxlength='30' name='txtORGANIZATION_NAME' value=''  />" else  Response.Write "<input size='40' maxlength='30' name='txtORGANIZATION_NAME' value=''  />" end if%>
                       
                    </td>
                    <td align="LEFT" style="padding-left: 20px" width="20%" nowrap>ORGANIZATION CODE 
                        <% if isEditable then if(isExists) then Response.Write "<input size='40' maxlength='30' name='txtORGANIZATION_CODE' value='' />" else Response.Write "<input size='40' maxlength='30' name='txtORGANIZATION_CODE' value='' />"  end if%>
                       
                    </td>
                </tr>
            </tbody>
            <tfoot>
                <tr>
                    <td align="right" colspan="2">
                        <!-- <input type="button" tabindex=14 style="width: 2cm" value="New" name="btn_iFrameAdd" onClick="btn_iFrmAdd();" class=button>&nbsp;
                        <input type="button" tabindex=15 style="width: 2cm" value="Update" name="btn_iFrameUpdate" onCLick="btn_iFrmUpdate();" class=button> -->
                        <input type="button" name="btnClose" value="Close" style="width: 2cm" onclick="return btnClose_onclick();">&nbsp;&nbsp; 
                        <% if isEditable <> 1  then Response.Write "<input type='button' name='btnSave' value='Save' style='width: 2cm' onclick='return frmAlias_onsubmit();'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" end if %>
                    </td>
                </tr>
            </tfoot>
        </table>
    </form>
</body>
</html>

