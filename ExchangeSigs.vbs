 
Function GetUserDN(BYVAL UN, BYVAL DN)
       Set ObjTrans = CreateObject("NameTranslate")
       objTrans.init 1, DN
       objTrans.set 3, DN & "\" & UN
       strUserDN = objTrans.Get(1)
       GetUserDN = strUserDN
End Function

Sub SetDefaultSignature(strSigName, strProfile)
    Const HKEY_CURRENT_USER = &H80000001
    strComputer = "."
    
    If Not IsOutlookRunning Then
        Set objreg = GetObject("winmgmts:" & _
         "{impersonationLevel=impersonate}!\\" & _
         strComputer & "\root\default:StdRegProv")
       strKeyPath = "Software\Microsoft\Windows NT\" & _
                    "CurrentVersion\Windows " & _
                    "Messaging Subsystem\Profiles\"
        ' get default profile name if none specified
        If strProfile = "" Then
           objreg.GetStringValue HKEY_CURRENT_USER, _
             strKeyPath, "DefaultProfile", strProfile
        End If
        ' build array from signature name
        myArray = StringToByteArray(strSigName, True)
       strKeyPath = strKeyPath & strProfile & _
                    "\9375CFF0413111d3B88A00104B2A6676"
       objreg.EnumKey HKEY_CURRENT_USER, strKeyPath, _
                      arrProfileKeys
        For Each subkey In arrProfileKeys
            strsubkeypath = strKeyPath & "\" & subkey
           'On Error Resume Next
           objreg.SetBinaryValue HKEY_CURRENT_USER, _
             strsubkeypath, "New Signature", myArray
           objreg.SetBinaryValue HKEY_CURRENT_USER, _
             strsubkeypath, "Reply-Forward Signature", myArray
        Next
    Else
        strMsg = "Please shut down Outlook before " & _
                "running this script."
        MsgBox strMsg, vbExclamation, "SetDefaultSignature"
    End If
End Sub
 
Function IsOutlookRunning()
    strComputer = "."
    strQuery = "Select * from Win32_Process " & _
              "Where Name = 'Outlook.exe'"
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" _
        & strComputer & "\root\cimv2")
    Set colProcesses = objWMIService.ExecQuery(strQuery)
    For Each objProcess In colProcesses
        If UCase(objProcess.Name) = "OUTLOOK.EXE" Then
           IsOutlookRunning = True
        Else
           IsOutlookRunning = False
        End If
    Next
End Function
 
Public Function StringToByteArray _
                (Data, NeedNullTerminator)
    Dim strAll
    strAll = StringToHex4(Data)
    If NeedNullTerminator Then
        strAll = strAll & "0000"
    End If
    intLen = Len(strAll) \ 2
    ReDim arr(intLen - 1)
    For i = 1 To Len(strAll) \ 2
        arr(i - 1) = CByte _
                  ("&H" & Mid(strAll, (2 * i) - 1, 2))
    Next
    StringToByteArray = arr
End Function
 
Public Function StringToHex4(Data)
    ' Input: normal text
    ' Output: four-character string for each character,
   '         e.g. "3204" for lower-case Russian B,
   '        "6500" for ASCII e
    ' Output: correct characters
    ' needs to reverse order of bytes from 0432
    Dim strAll
    For i = 1 To Len(Data)
        ' get the four-character hex for each character
        strChar = Mid(Data, i, 1)
        strTemp = Right("00" & Hex(AscW(strChar)), 4)
        strAll = strAll & Right(strTemp, 2) & Left(strTemp, 2)
    Next
    StringToHex4 = strAll
End Function

Dim objFSO, objWsh, appDataPath, pathToCopyTo, plainTextFile, plainTextFilePath, richTextFile, richTextFilePath, htmlFile, htmlFilePath
 
Set objUser = CreateObject("WScript.Network")
userName = objUser.UserName
domainName = objUser.UserDomain 

Set objLDAPUser = GetObject("LDAP://" & GetUserDN(userName,domainName))
 
'Prepare to create some files
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objWsh = CreateObject("WScript.Shell")
 
appDataPath = objWsh.ExpandEnvironmentStrings("%APPDATA%")
pathToCopyTo = appDataPath & "\Microsoft\Signatures\"
 
If objFSO.FolderExists(pathToCopyTo) = false Then
  set ffolder = objFSO.CreateFolder(pathToCopyTo)
  set ffolder = nothing
End If
 
'Let's create the plain text signature
plainTextFilePath = pathToCopyTo & "Default.txt"
Set plainTextFile = objFSO.CreateTextFile(plainTextFilePath, TRUE)
 
  plainTextFile.WriteLine("-- ")
  plainTextFile.WriteLine(objLDAPUser.DisplayName)
  plainTextFile.WriteLine(objLDAPUser.title)
  plainTextFile.WriteLine(objLDAPUser.department)
  plainTextFile.WriteLine(objLDAPUser.company)
  plainTextFile.WriteLine(objLDAPUser.mail)
  plainTextFile.Write(objLDAPUser.website)
  plainTextFile.Close
 
Set plainTextFile = nothing
 
'Now we create the Rich Text signature
richTextFilePath = pathToCopyTo & "Default.rtf"
Set richTextFile = objFSO.CreateTextFile(richTextFilePath, TRUE)
 
  richTextFile.WriteLine("{\rtf1\ansi\ansicpg1252\deff0\deflang2057{\fonttbl{\f0\fswiss\fcharset0 Arial;}}")
  richTextFile.WriteLine("\viewkind4\uc1\pard\f0\fs20 -- \par")
  richTextFile.WriteLine(objLDAPUser.DisplayName & "\par")
  richTextFile.WriteLine(objLDAPUser.title & "\par")
  richTextFile.WriteLine(objLDAPUser.department & "\par")
  richTextFile.WriteLine(objLDAPUser.company & "\par")
  richTextFile.WriteLine(objLDAPUser.mail & "\par")
  richTextFile.WriteLine(objLDAPUser.website \par")
  richTextFile.Write("}")
  richTextFile.Close
 
Set richTextFile = nothing
 
'And finally, the HTML signature
htmlFilePath = pathToCopyTo & "Default.htm"
Set htmlFile = objFSO.CreateTextFile(htmlFilePath, TRUE)
 
  htmlfile.WriteLine("<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">")
  htmlfile.WriteLine("<html xmlns=""http://www.w3.org/1999/xhtml"" >" &_
                     "<style type=""text/stylesheet"" media=""screen"">" &_
                     " a:link, a:visited, a:hover, a:active { font-family:Arial,Helvetica,sans-serif;color:#333399; } " &_
                     "</style>")
  htmlfile.WriteLine("<body>")
  htmlfile.WriteLine("<div style=""font-size:10pt; font-family:Arial,Helvetica, sans-serif; color: #333399"">")
  ' name and title
 htmlfile.WriteLine("<p><strong>" & objLDAPUser.DisplayName & "<br/>")
  htmlfile.WriteLine(objLDAPUser.title & "<br/>")
  htmlfile.WriteLine(objLDAPUser.department & "</strong></p>")
 

  ' contact details
  htmlfile.WriteLine("<p>")
  htmlfile.WriteLine("Tel: <strong>" & objLDAPUser.telephoneNumber & "</strong><br />")
  htmlfile.WriteLine("Fax: <strong>" & objLDAPUser.faxNumber & "</strong></p>")
 
  htmlfile.WriteLine("<p><a href=""mailto:" & objLDAPUser.mail & """>" & objLDAPUser.mail & "</a><br/>")
  htmlfile.WriteLine("<a href=""" & objLDAPUser.website & """>" & objLDAPUser.website & "</a></p>")
  
  ' disclaimer
  htmlfile.WriteLine("<p style=""font-size:8pt;"">Your disclaimer goes here</p>")
  htmlfile.WriteLine("</div>")
  htmlfile.WriteLine("</body>")
  htmlfile.Write("</html>")
  htmlfile.Close
 
Set htmlFile = nothing
 
' Set the created files as default signatures in Outlook
 
' Use this version to set all accounts
' in the default mail profile
' to use a previously created signature
Call SetDefaultSignature("Default", "")

