Option explicit 

'
' This utility will reead a file and encypt lines ending with "#Encrypt"
' For example ...
'     const a="Hello" #Encrypt
' Becomes ...
'     const a="kjfgvdav;lvdv"
'
' Run it like this ...
' C:\Windows\SysWOW64\cscript EncryptTool.vbs SMA_Env.inc

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

Const ForReading = 1
Const ForWriting = 2

Dim oFile
Dim oEncryptedData, objWMIService, colItems
Dim fso, tsFile, strNewLine, strTempFileName, strEnc
Dim strLines(), intSize, i

Dim decData

' Check the arguments are passed
If Wscript.Arguments.Count <> 1 Then
	Wscript.Echo "Run as:  EncryptTool.vbs {Filename}"
	Wscript.Quit
End If

' Get the file name 
Set fso = CreateObject("Scripting.FileSystemObject") 
Set oFile = fso.GetFile(Wscript.Arguments.Item(0))
' Open the file to read it into an array
set tsFile = oFile.OpenAsTextStream(ForReading)
intSize = 0
Do While not tsFile.AtEndOfStream
	      intSize = intSize + 1
				ReDim Preserve strLines(intSize)
     		strLines(intSize - 1) = tsFile.ReadLine
Loop 
tsFile.Close    		

Set oEncryptedData = CreateObject("CAPICOM.EncryptedData")			
oEncryptedData.Algorithm.KeyLength = CAPICOM_ENCRYPTION_KEY_LENGTH_256_BITS
oEncryptedData.Algorithm.Name = CAPICOM_ENCRYPTION_ALGORITHM_AES
' Set the secret to be used when deriving the key based on the unique server id
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery ( "SELECT UUID FROM Win32_ComputerSystemProduct")
Wscript.Echo("UUID=" + colItems.ItemIndex(0).UUID)
oEncryptedData.SetSecret colItems.ItemIndex(0).UUID, CAPICOM_SECRET_PASSWORD

' Now update lines
set tsFile = oFile.OpenAsTextStream(ForWriting)
For i=0 To intSize - 1
     		' Only encrypt non-comment lines with the key word 
     		If Left(Trim(strLines(i)),1) <> "'" And InStr(strLines(i),"#Encrypt") > 0 Then
     			Wscript.Echo("Encrypting:" & strLines(i))
     			oEncryptedData.Content = Mid(strLines(i), InStr(strLines(i),"""") + 1, InStrRev(strLines(i),"""") - InStr(strLines(i),"""") - 1)
     			
     			Wscript.Echo(oEncryptedData.Content)
     			
     			strEnc = oEncryptedData.Encrypt(CAPICOM_ENCODE_BASE64) 
     			strNewLine = Left(strLines(i), InStr(strLines(i),"""")) & strEnc & """"
     			strNewLine = Replace(strNewLine, Chr(13) & Chr(10), "")
     			Wscript.Echo("Outputting:" & strNewLine)
     			tsFile.WriteLine(strNewLine)
     		Else
     			tsFile.WriteLine(strLines(i))
     			Wscript.Echo("Skipping:" & strLines(i))
     		End If
Next
tsFile.Close
Wscript.Echo("Updated " & Wscript.Arguments.Item(0))
