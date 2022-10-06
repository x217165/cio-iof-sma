<%@ Language=VBScript %>
<% Response.Buffer = True%>
<!--#include file="smaProcs.inc"-->
<!--#include file="smaConstants.inc"-->
<!--#include file="SMA_Env.inc"-->

<html>
<head>
<title>Test Encrypt</title>
<body>

<%

      encrypted = strConstSConnectString
		  ' CAPICOM Constants                                                          
			Const CAPICOM_ENCRYPTION_ALGORITHM_RC2 = 0
			Const CAPICOM_ENCRYPTION_ALGORITHM_RC4 = 1
			Const CAPICOM_ENCRYPTION_ALGORITHM_DES = 2
			Const CAPICOM_ENCRYPTION_ALGORITHM_3DES = 3
			Const CAPICOM_ENCRYPTION_ALGORITHM_AES = 4
			Const CAPICOM_ENCRYPTION_KEY_LENGTH_MAXIMUM = 0
			Const CAPICOM_ENCRYPTION_KEY_LENGTH_40_BITS = 1
			Const CAPICOM_ENCRYPTION_KEY_LENGTH_56_BITS = 2
			Const CAPICOM_ENCRYPTION_KEY_LENGTH_128_BITS = 3
			Const CAPICOM_ENCRYPTION_KEY_LENGTH_192_BITS = 4
			Const CAPICOM_ENCRYPTION_KEY_LENGTH_256_BITS = 5 
			Const CAPICOM_SECRET_PASSWORD = 0
			Const CAPICOM_ENCODE_BASE64 = 0
			Const CAPICOM_ENCODE_BINARY = 1
			Const CAPICOM_ENCODE_ANY = -1
		
    	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    	Set colItems = objWMIService.ExecQuery ( "SELECT UUID FROM Win32_ComputerSystemProduct")

			Set oEncryptedData = Server.CreateObject("CAPICOM.EncryptedData")
			oEncryptedData.Algorithm.KeyLength = CAPICOM_ENCRYPTION_KEY_LENGTH_256_BITS
			oEncryptedData.Algorithm.Name = CAPICOM_ENCRYPTION_ALGORITHM_AES
			oEncryptedData.SetSecret colItems.ItemIndex(0).UUID, CAPICOM_SECRET_PASSWORD
			oEncryptedData.Decrypt(encrypted)
			decData = oEncryptedData.Content
			Response.Write("<p>Text=" + decData + "</p>")

       strConnect = Decrypt(strConstSConnectString)
	Response.Write("<p>Connect=" + strConnect + "</p>")
			  	  			
%>

</body>
</html>
