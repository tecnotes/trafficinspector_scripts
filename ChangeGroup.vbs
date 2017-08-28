'
' tecnotes.net, 2017
'
' VBScript for Traffic Inspector 2.xx, 3.xx
' ----------------------
' Скрипт смены группы у пользователя.
'
' Укажите None в параметре <Группа> если необходимо переместить пользователя в группу: Пользователи вне группы
'

Set Args = WScript.Arguments.Unnamed
	If Args.Count <> 4 Then
	WScript.Echo "cscript ChangeGroup.vbs <Администратор> <Пароль> <Пользователь> <Группа>"
	WScript.Quit
	End If

AdmID = Args(0)
AdmPass = Args(1)
User = Args(2)
Group2 = Args (3)

Set Srv = CreateObject("TrafInsp.TrafInspAdmin")
Set Perm = Srv.QueryPermissions()
LogOn = Perm.DoSharedLogon(AdmID, AdmPass, "Script")
Set Dom = WScript.CreateObject("Msxml2.DOMDocument.6.0")

itUserGroup = 2
conf_AttrLevelDetail = 5
GroupIsTrue = 0
Dom.LoadXML Srv.GetList(itUserGroup, null, null, conf_AttrLevelDetail)
Set DocEl = Dom.DocumentElement
Set Nodes = DocEl.selectNodes("UserGroupItem")

If Group2 = "None" then
	Group2Name = ""
	Wscript.Echo "Группа 'Пользователи вне группы'"
	GroupIsTrue = 1
	Else
	
		For Each Node in Nodes

			If Node.getAttribute("DisplayName") <> Group2 then
			Else
				Group2Name = Group2
				Group2GUID = Node.getAttribute("GUID")
				GroupIsTrue = 1
			End if

		Next
	
	End if 

If GroupIsTrue = 0 then
	Wscript.Echo "Группа с именем '" & Group2 & "' не найдена."
	Wscript.Quit
Else
End if

itUser = 3
conf_AttrLevelNormal = 0
Set Dom1 = WScript.CreateObject("Msxml2.DOMDocument.6.0")
Dom1.LoadXML Srv.GetList(itUser, null, null, conf_AttrLevelDetail)
Set DocEl = Dom1.DocumentElement
Set Nodes = DocEl.selectNodes("UserItem")
For Each Node in Nodes

	If Node.getAttribute("DisplayName") <> User Then
		Wscript.Echo "Нет пользователя с такими именем: '" & User & "'."
	Else
		WScript.Echo "Имя клиента " & Chr(9) & Node.getAttribute("DisplayName")
		WScript.Echo "Группа " & Chr(9) & Node.getAttribute("GroupDisplayName")

		Node.setAttribute "GroupDisplayName", Group2Name
		Node.setAttribute "Group", Group2GUID
		Srv.UpdateList itUser, Dom1.xml

		WScript.Echo "Имя клиента " & Chr(9) & Node.getAttribute("DisplayName")
		
		If Group2 = "None" then
			WScript.Echo "Группа " & Chr(9) & "Пользователи вне группы"
		else
			WScript.Echo "Группа " & Chr(9) & Node.getAttribute("GroupDisplayName")
		End if
		
	End If

Next
