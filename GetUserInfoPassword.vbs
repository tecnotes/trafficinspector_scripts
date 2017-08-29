'
' tecnotes.net, 2017
'
' VBScript for Traffic Inspector 2.xx, 3.xx
' ----------------------
' Скрипт для получения информации (группа, пароль) о пользователе по его имени
'
'
Set Args = WScript.Arguments.Unnamed
If Args.Count <> 3 Then
	WScript.Echo "cscript GetUserInfoPassword.vbs <Администратор> <Пароль> <Пользователь> "
	WScript.Quit        
End If
 
AdmID = Args(0)
AdmPass = Args(1)
User = Args(2)

Set Srv = CreateObject("TrafInsp.TrafInspAdmin")
Set Perm = Srv.QueryPermissions()
LogOn = Perm.DoSharedLogon(AdmID, AdmPass, "Script")

itUser = 3
conf_AttrLevelNormal = 0
IsUserPresent = 0

Set Dom1 = WScript.CreateObject("Msxml2.DOMDocument.6.0")
Dom1.LoadXML Srv.GetList(itUser, null, null, conf_AttrLevelDetail)

Set DocEl = Dom1.DocumentElement
	Set Nodes = DocEl.selectNodes("UserItem")
		For Each Node in Nodes
			If Node.getAttribute("DisplayName") = User Then
				WScript.Echo "Имя клиента        " & Chr(9) & Node.getAttribute("DisplayName")
				WScript.Echo "GUID             " & Chr(9) & Node.getAttribute("GUID")
				WScript.Echo "Группа             " & Chr(9) & Node.getAttribute("GroupDisplayName")
				WScript.Echo "Группa GUID        " & Chr(9) & Node.getAttribute("Group")
				WScript.Echo "Пароль 	         " & Chr(9) & Node.getAttribute("Password")
				IsUserPresent = 1
			End IF
		Next
		
If IsUserPresent = 0 then
	Wscript.Echo "Пользователеь с таким именем отсутствует"
End If
