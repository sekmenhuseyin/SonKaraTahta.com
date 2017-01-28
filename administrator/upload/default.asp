<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>

<%
	' For disabling password security set wexPassword = "" 
	Const wexPassword = "mnfixed"
	' Root folder, it can be a physical or virtual folder: "/test", "/"
	Const wexRoot = "/"
	' Preferred character set
	Const wexCharSet = "ISO"
	' Set script timeout value to higher values (in seconds) if the script fails when uploading large files
	Server.ScriptTimeout = 300
' ------------------------------------------------------------

	Const appName = "Mini-NUKE Fixed Version  - Upload"
	Const appVersion = ""

	Dim scriptName
	scriptName = Request.ServerVariables("SCRIPT_NAME")
	
	Dim FSO
	Set FSO = server.CreateObject ("Scripting.FileSystemObject")

	Dim wexId
	wexId = appName & appVersion & "-" & FSO.GetParentFolderName(scriptName) & "-"
	
	Dim wexMessage, wexRootPath

	Const editableExtensions = "*htm*|*html*|*asp*|*asa*|*txt*|*inc*|*css*|*aspx*|*js*|*vbs*|*shtm*|*shtml*|*xml*|*xsl*|*log*"
	Const viewableExtensions = "*gif*|*jpg*|*jpeg*|*png*|*bmp*|*jpe*"
	
	Const iconFolderOpenBig = "<img align=absmiddle border=0 width=32 height=27 src=""./folder_open_big.gif"">"
	Const iconFolderUp = "<img align=absmiddle border=0 width=15 height=13  src=""./folder_up.gif"" alt=""Yukari"">"
	Const iconFolder = "<img align=absmiddle border=0 width=15 height=13 src=""./folder.gif"" alt=""Dosya Ayrintisi"">"
	Const iconFile = "<img align=absmiddle border=0 width=11 height=14 src=""./file.gif"" alt=""Dosya Ayrintisi"">"
	Const iconFileEditable = "<img align=absmiddle border=0 width=11 height=14 src=""./file_editable.gif"" alt=""Metin Dosyasini Düzenle"">"
	Const iconFileViewable = "<img align=absmiddle border=0 width=11 height=14 src=""./file_viewable.gif"" alt=""Resim Detayi"">"
	
	Const iconRefresh = "<img align=absmiddle border=0 width=21 height=20 src=""./refresh.gif"" alt=""Sayfayi Yenile"">"
	Const iconCreateFile = "<img align=absmiddle border=0 width=21 height=20 src=""./create_file.gif"" alt=""Yeni Dizin"">"
	Const iconCreateFolder = "<img align=absmiddle border=0 width=21 height=20 src=""./create_folder.gif"" alt=""Yeni Klasör"">"
	Const iconUpload = "<img align=absmiddle border=0 width=21 height=20 src=""./upload.gif"" alt=""Dosyalari Yükle"">"
	Const iconLogout = "<img align=absmiddle border=0 width=21 height=20 src=""./logout.gif"" alt=""ÇIKIS"">"
	Const iconDelete = "<img align=absmiddle border=0 width=21 height=20 src=""./delete.gif"" alt=""SiL"">"

' - UPLOAD functions ------------------------------------
	' Writes the html header
	Sub HtmlHeader (title, charset)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=<%=charSet%>">
<title>Mini-NUKE Fixed Version UPLOAD</title>
<%HtmlStyle%>
<%HtmlJavaScript%>
</head>
<body>
<%	
	End Sub
	
	' Writes the html footer
	Sub HtmlFooter ()
%>	
</body>
</html>
<%
	End Sub

	' Writes the copyright message
	Sub HtmlCopyright()
	%>
		<table cellspacing=0 cellpadding=0 border=0 align=center><tr><td>
			<a href="default.asp"><< Yönetim Sayfasý</a>
		</td></tr></table>
	<%
	End Sub
	
	' Writes the stylesheet
	Sub HtmlStyle
%>
<style>
BODY
{
    BACKGROUND-COLOR: #FFFFFF
}
TD
{
    FONT-WEIGHT: normal;
    FONT-SIZE: 10pt;
    COLOR: black;
    FONT-FAMILY: Verdana,Tahoma,Arial,Helvetica
}
.formClass
{
    BACKGROUND-COLOR: #C0C0C0;
    FONT-WEIGHT: normal;
    FONT-SIZE: 10pt;
    COLOR: black;
    FONT-FAMILY: Verdana,Tahoma,Arial,Helvetica
}
.lightRow {
	BACKGROUND-COLOR: #F4F4F4
}
.darkRow {
	BACKGROUND-COLOR: #E4E4E4
}
.titleRow {
	BACKGROUND-COLOR: #999999
}
.loginRow {
	border: black solid 1px;
	BACKGROUND-COLOR: #999999
}
.boldText
{
    FONT-WEIGHT: bold
}
A
{
    FONT-WEIGHT: bold;
    COLOR: #CC0000;
    TEXT-DECORATION: none
}
A:hover
{
    COLOR: #CC0000;
    TEXT-DECORATION: underline
}
A:visited
{
    TEXT-DECORATION: none
}
A:active
{
    COLOR: #000000;
    TEXT-DECORATION: none
}
</style>
<%
	End Sub
	
	' Writes the javascript code
	Sub HtmlJavaScript
%>
<script language=javascript>
	function Command(cmd, param) {
		var str;
		var someWin;
		switch (cmd) {
			case "NewFile":
				str = prompt("Lütfen Yeni Dizin ismi Girin", "Yeni Dizin");
				if(!str) return;
				document.forms.formBuffer.parameter.value = str;
				break;
			case "NewFolder":
				str = prompt("Lütfen Yeni Klasör ismi Girin", "Yeni Klasör");
				if(!str) return;
				document.forms.formBuffer.parameter.value = str;
				break;
			case "Edit":
				str = document.forms.formBuffer.folder.value + param;
				str = str.replace(/\W/gi,"_");
				someWin = openWin(cmd + str, "", 600, 440, false, false);
				createPage(someWin,cmd,param);
				someWin.focus(); 
				someWin = null;
				return;
				break;
			case "View":
				if(document.forms.formBuffer.virtual.value=="") {alert("Can not view image without web access !"); return;}
				str = document.forms.formBuffer.folder.value + param;
				str = str.replace(/\W/gi,"_");
				someWin = openWin(cmd + str, "", 600, 440, false, true);
				createPage(someWin,cmd,param);
				someWin.focus(); 
				someWin = null;
				return;
				break;
			case "FileDetails":
			case "FolderDetails":
				str = document.forms.formBuffer.folder.value + param;
				str = str.replace(/\W/gi,"_");
				someWin = openWin(cmd + str, "", 350, 220, false, false);
				createPage(someWin,cmd,param);
				someWin.focus(); 
				someWin = null;
				return;
				break;
			case "OpenFile":
				if(document.forms.formBuffer.virtual.value=="") {alert("No web access !"); return;}
				document.location = document.forms.formBuffer.virtual.value + param;
				return;
				break;
			case "Upload":
				someWin = openWin(cmd, "", 400, 150, true, false);
				createPage(someWin,cmd,param);
				someWin.focus(); 
				someWin = null;
				return;
				break;
			case "DeleteFolder":
				if (!confirm('Are you sure to delete the folder "' + param + '" and all its contens ?')) return;
				document.forms.formBuffer.parameter.value = param;
				break;
			case "DeleteFile":
				if (!confirm('Are you sure to delete "' + param + '" ?')) return;
				document.forms.formBuffer.parameter.value = param;				
				break;
			default:
				document.forms.formBuffer.parameter.value = param;
		}
		document.forms.formBuffer.command.value = cmd
		document.forms.formBuffer.submit();	
	}
	
	function Check() {
		if (document.forms.formBuffer.pwd.value == "") {
			alert("Lütfen Sifre Giriniz!"); 
			return false;
		} else return true;
	}

	function openWin(winName, urlLoc, w, h, showStatus, isViewer) {
		l = (screen.availWidth - w)/2;
		t = (screen.availHeight - h)/2;
		features  = "toolbar=no";      // yes|no 
		features += ",location=no";    // yes|no 
		features += ",directories=no"; // yes|no 
		features += ",status=" + (showStatus?"yes":"no");  // yes|no 
		features += ",menubar=no";     // yes|no 
		features += ",scrollbars=" + (isViewer?"yes":"no");   // auto|yes|no 
		features += ",resizable=" + (isViewer?"yes":"no");   // yes|no 
		features += ",dependent";      // close the parent, close the popup, omit if you want otherwise 
		features += ",height=" + h;
		features += ",width=" + w;
		features += ",left=" + l;
		features += ",top=" + t;
		return window.open(urlLoc,winName,features);
	} 
	
	function createPage (theWin, cmd, param, action){
		theWin.document.write ("<html><title>UPLOAD</title><body>");
		theWin.document.write ('<form name="formBuffer" action="' + document.forms.formBuffer.action + '" method="post">');
		theWin.document.write ('<input type="hidden" name="command" value="">');
		theWin.document.write ('<input type="hidden" name="parameter" value="">');
		theWin.document.write ('<input type="hidden" name="folder" value="">');
		theWin.document.write ('</form>');
		theWin.document.write ("</body></html>");
		theWin.document.forms.formBuffer.command.value = cmd;
		theWin.document.forms.formBuffer.parameter.value = param;
		theWin.document.forms.formBuffer.folder.value = document.forms.formBuffer.folder.value;
		theWin.document.forms.formBuffer.submit();
	}

	function EditorCommand (cmd) {
		switch (cmd) {
			case "Info":
				alert(document.forms.formBuffer.info.value.replace(/\|/gi,"\n"));
				break;
			case "Reload":
				document.forms.formBuffer.reset();
				break;
			case "Save":
				document.forms.formBuffer.subcommand.value = "Save";
				document.forms.formBuffer.submit();
				break;
			case "SaveAs":
				var str, oldname;
				oldname = document.forms.formBuffer.parameter.value;
				str = prompt("Bu isimle Farkli Kaydet :", oldname);
				if (!str || str==oldname) return;
				document.forms.formBuffer.parameter.value = str;
				document.forms.formBuffer.subcommand.value = "SaveAs";
				document.forms.formBuffer.submit();
				break;
		}
	}

	function ViewerCommand (cmd) {
		switch (cmd) {
			case "Info":
				alert(document.forms.formBuffer.info.value.replace(/\|/gi,"\n"));
				break;
			case "Reload":
				document.forms.formBuffer.submit();
				break;
		}
	}

	function Upload() {
		document.forms.formBuffer.submit();
	}
</script>
<%
	End Sub

	' Write file listing of the given folder
	Sub WriteListing (byref folder)
		Dim item, arr
		Dim rowType
		
		on error resume next
		
		If wexMessage="" Then wexMessage = "Toplam " & (folder.subfolders.Count + folder.files.Count) & " Parça..."
%>

<form name=formGlobal action="noaction">
<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<tr bgcolor="#CCCCCC">
		<td align=left>&nbsp;<font class=boldText>Mini-NUKE Fixed Version UPLOAD</font> - <a href="http://<%=Request.ServerVariables("SERVER_NAME")%>" target="_blank"><%=Request.ServerVariables("SERVER_NAME")%></a></td>
		<td align=right><font class=boldText>v<%=appVersion%></font> - <%=Date()%>&nbsp;</td>
	</tr>
</table>
<table cellspacing=0 cellpadding=1 border=0 width=100%>
	<tr class=lightRow height=60>
		<td>
			<div style="font-size:12pt;">&nbsp;<%=iconFolderOpenBig%>&nbsp;<font class=boldText><%=folder.Name%></font></div>
			<div style="font-size:8pt;">&nbsp;&nbsp;<%=folder.path%></div>
			<%If VirtualPath(destFolder)<>"" Then%>
				<div style="font-size:8pt;">&nbsp;&nbsp;(<a href="<%=VirtualPath(destFolder)%>" target="_blank"><%=VirtualPath(destFolder)%></a>)</div>
			<%Else%>
				<div style="font-size:8pt;">&nbsp;&nbsp;(no web access)</div>
			<%End If%>
		</td>
		<td>
			<font class=boldText><%=folder.subfolders.count%></font> KLASÖR(ler)<br>
			<font class=boldText><%=folder.files.count%></font> DiZiN(ler)
		</td>
		<td>
			Toplam Boyut: <font class=boldText><%If err.Number<>0 Then Response.Write "N/A" Else Response.Write FormatSize(folder.size)%></font>
		</td>
		<td>
			<a href="javascript:Command('Refresh', '');"><%=iconRefresh%></a>&nbsp;
			<a href="javascript:Command('NewFile', '');"><%=iconCreateFile%></a>&nbsp;
			<a href="javascript:Command('NewFolder', '');"><%=iconCreateFolder%></a>&nbsp;
			<a href="javascript:Command('Upload', '');"><%=iconUpload%></a>&nbsp;
<%If wexPassword <> "" Then%>
			<a href="javascript:Command('Logout', '');"><%=iconLogout%></a>
<%End If%>
			<br><input name=wexMessage type=text class=formClass size=20 value="<%=server.HTMLEncode(wexMessage)%>">
		</td>
		<td>&nbsp;</td>
	</tr>
	<tr class=titleRow>
		<td>&nbsp;<font class=boldText>Dosya veya Klasör Adý</font></td>
		<td>&nbsp;<font class=boldText>Boyut</font></td>
		<td>&nbsp;<font class=boldText>Tür</font></td>
		<td>&nbsp;<font class=boldText>Son Deðiþiklik</font></td>
		<td>&nbsp;</td>
	</tr>
<%
	rowType = "darkRow"

	If len(destFolder) > len(wexRootPath) Then
%>
	<tr class=<%=rowType%>><td>&nbsp;<a href="javascript:Command('LevelUp','')"><%=iconFolderUp%></a>&nbsp;<a href="javascript:Command('LevelUp','')">..</a></td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>
<%
		rowType = "lightRow"
	End If
	
	If (folder.subfolders.Count + folder.files.Count) = 0 Then
		' Do nothing when error occurs
	Else
		For each item in folder.subfolders
%>
	<tr class=<%=rowType%>><td>&nbsp;<%=GetIcon(item.Name, true)%>&nbsp;<a href="javascript:Command('OpenFolder','<%=item.Name%>')"><%=item.Name%></a></td><td>&nbsp;<%=FormatSize(item.Size)%></td><td>&nbsp;<%=item.Type%></td><td>&nbsp;<%=item.DateLastModified%></td><td>&nbsp;<a href="javascript:Command('DeleteFolder','<%=item.Name%>')"><%=iconDelete%></a></td></tr>
<%
			If rowType = "darkRow" Then rowType = "lightRow" Else rowType = "darkRow"
		Next

		For each item in folder.files
%>
	<tr class=<%=rowType%>><td>&nbsp;<%=GetIcon(item.Name, false)%>&nbsp;<a href="javascript:Command('OpenFile','<%=item.Name%>')"><%=item.Name%></a></td><td>&nbsp;<%=FormatSize(item.Size)%></td><td>&nbsp;<%=item.Type%></td><td>&nbsp;<%=item.DateLastModified%></td><td>&nbsp;<a href="javascript:Command('DeleteFile','<%=item.Name%>')"><%=iconDelete%></a></td></tr>
<%
			If rowType = "darkRow" Then rowType = "lightRow" Else rowType = "darkRow"
		Next
	End If
%>
	<tr></tr>

</table>
<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<tr class=titleRow>
		<td>&nbsp;</td>
	</tr>
</table>
</form>
<%
	End Sub

	' UPLOAD Login screen
	Sub	Login
		If Request.Form("command") = "Login" Then
			If Request.Form("pwd") = wexPassword Then
				Session(wexId & "Login") = true
				Exit Sub
			Else
				wexMessage = "HATALI SiFRE!"
			End If
		End If
		
		HtmlHeader appName, wexCharSet
		If(wexMessage<>"") Then Response.Write "<script language=""javascript"">alert('" & wexMessage & "');</script>"
%>
<%
If session("Member") = "" then
Response.Write "<center><font class=""tabloalt_font"">Lütfen Giriþ Yapýnýz.</font></center>"
Else
%>
<%
IF Session("Level") = "1" OR Session("Level") = "7" Then
%>
<form name=formBuffer method=post action="<%=scriptName%>" onSubmit="javascript:return(Check());">
<table border=0 cellspacing=0 cellpadding=0 width=400 align=center>
	<tr><td><br><br><br></td></tr>
	<tr><td>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=left>
					&nbsp;<b>Mini-NUKE Fixed Versi</b><font class=boldText>on - UPLOAD</font>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr align=center class=lightRow>
				<td>
					<br>
					<font class=boldText>Upload Yöneticisine Hoþ Geldiniz...</font>
					<br><br>
					<table cellspacing=0 cellpadding=5 border=0 class=loginRow>
						<tr>
							<td align=left>&nbsp;<font class=boldText>Sifre</font></td>
						</tr>
						<tr>
							<td align=center><input type="password" class=formClass name=pwd value="" size=21></td>
						</tr>
						<tr>
							<td align=right><input type=submit name=submitter value="Giris ->>" class=formClass></td>
						</tr>
					</table>
					<br><br><br>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=center>&nbsp;</td>
			</tr>
		</table>
	</td></tr>
</table>
<input type="hidden" name=command value="Login">
</form>
<% END IF%>
<% END IF%>
<%	
		HtmlFooter
		Response.End 
	End Sub
	
	' Checks if there is a valid login
	Function Secure()
		If wexPassword = "" Then
			Secure = true
		Else
			If Session(wexId & "Login") Then Secure = true Else Secure = false
		End If
	End Function
		
	' Logouts UPLOAD session
	Sub Logout
		Session(wexId & "Login") = false
		Login
	End Sub
		
	' Returns the icon of the file
	Function GetIcon(fileName, isFolder)
		Dim ext
		If isFolder Then
			GetIcon = "<a href=""javascript:Command('FolderDetails', '" & fileName & "');"">" & iconFolder & "</a>"
		Else
			ext = FSO.GetExtensionName(fileName)
			If InStr(1,editableExtensions, "*" & ext & "*", 1) <> 0 Then
				GetIcon = "<a href=""javascript:Command('Edit', '" & fileName & "');"">" & iconFileEditable & "</a>"
			ElseIf InStr(1,viewableExtensions, "*" & ext & "*", 1) <> 0 Then
				GetIcon = "<a href=""javascript:Command('View', '" & fileName & "');"">" & iconFileViewable & "</a>"
			Else
				GetIcon = "<a href=""javascript:Command('FileDetails', '" & fileName & "');"">" & iconFile & "</a>"
			End If
		End If
	End Function
	
	' Formats given size in bytes,KB,MB and GB
	Function FormatSize (givenSize)
		If (givenSize < 1024) Then
			FormatSize = givenSize & " B"
		ElseIf (givenSize < 1024*1024) Then
			FormatSize = FormatNumber(givenSize/1024,2) & " KB"
		ElseIf (givenSize < 1024*1024*1024) Then
			FormatSize = FormatNumber(givenSize/(1024*1024),2) & " MB"
		Else
			FormatSize = FormatNumber(givenSize/(1024*1024*1024),2) & " GB"
		End If
	End Function

	' Adds given type of the slash to the end of the path if required
	Function FixPath(path, slash)
		If Right(path, 1) <> slash Then
            FixPath = path & slash
        Else
			FixPath = path
        End If
	End Function

	' Converts the given path to physical path
	Function RealizePath(thePath)
		Dim path
		path = replace(thePath,"/","\")
		If left(path,1) = "\" Then
			on error resume next
			RealizePath = FixPath(server.MapPath(path),"\")
			If err.Number<>0 Then RealizePath = thePath
		Else
			If InStr(1,path, ":", 1) <> 0 Then
				RealizePath = FixPath(path,"\")
			Else
				RealizePath = thePath & "?"
			End If
		End If
	End Function	

	' Converts the given path to virtual path
	Function VirtualPath(thePath)
		Dim webRoot, path
		webRoot = FixPath(server.MapPath("/"),"\")
		path = FixPath(thePath,"\")
		VirtualPath = ""
		If left(wexRoot,1) = "/" Then
			VirtualPath = FixPath(wexRoot, "/")
			VirtualPath = VirtualPath & right(path, len(path) - len(wexRootPath))
			VirtualPath = replace(VirtualPath, "\", "/")
			VirtualPath = FixPath(VirtualPath,"/")
		ElseIf left(lcase(path), len(webRoot)) = lcase(webRoot) Then
			VirtualPath = "/" & right(path, len(path) - len(webRoot))
			VirtualPath = replace(VirtualPath, "\", "/")
			VirtualPath = FixPath(VirtualPath,"/")
		End If
	End Function
	
	' Makes sure that given file name does not contain path info
	Function SecureFileName(name)
		SecureFileName = replace(name,"/","?")
		SecureFileName = replace(SecureFileName,"\","?")
	End Function

	' Creates a folder or a file
	Function CreateItem()
		Dim itemType, itemName, itemPath
		itemType = Request.Form("command")
		itemName = SecureFileName(Request.Form("parameter"))
		itemPath = destFolder & itemName

		on error resume next
		
		Select Case itemType
			Case "NewFolder"
				If (FSO.FolderExists(itemPath) = false and FSO.FileExists(itemPath) = false ) Then 
					FSO.CreateFolder(itemPath)
					If (err.Number <> 0 ) Then 
						CreateItem = "HATA! < """ & itemName & """ > ADLI KLASÖR YARATILAMAZ..." 
					Else
						CreateItem = "YARATILAN KLASÖR """ & itemName & """..."
					End If
				Else
					CreateItem = "HATA! < """ & itemName & """ > ADLI KLASÖR YARATILAMAZ ÇÜNKÜ ZATEN BÖYLE BÝR KLASÖR VAR..."
				End If
			Case "NewFile"
				If (FSO.FolderExists(itemPath) = false and FSO.FileExists(itemPath) = false ) Then 
					FSO.CreateTextFile(itemPath)
					If (err.Number <> 0 ) Then 
						CreateItem = "HATA! < """ & itemName & """ > ADLI DiZiN YARATILAMAZ ÇÜNKÜ ZATEN BÖYLE BiR DiZiN VAR..."
					Else
						CreateItem = "YARATILAN DiZiN """ & itemName & """..."
					End If
				Else 
					CreateItem = "HATA! < """ & itemName & """ > ADLI DiZiN YARATILAMAZ ÇÜNKÜ ZATEN BÖYLE BiR DiZiN VAR..."
				End IF
		End Select
	End Function
	
	' Deletes a folder or a file
	Function DeleteItem
		Dim itemType, itemName, itemPath
		itemType = Request.Form("command")
		itemName = SecureFileName(Request.Form("parameter"))
		itemPath = destFolder & itemName

		on error resume next
		
		Select Case itemType
			Case "DeleteFolder"
				FSO.DeleteFolder itemPath, true
				If (err.Number <> 0 ) Then 
					DeleteItem = "HATA! < """ & itemName & """ > Klasörü Silinemez Çünkü Daha Önceden Silinmis..." 
				Else
					DeleteItem = "Silinen Klasör """ & itemName & """..."
				End If
			Case "DeleteFile"
				FSO.DeleteFile itemPath, true
				If (err.Number <> 0 ) Then 
					DeleteItem = "HATA! < """ & itemName & """ > Dizini Silinemez Çünkü Daha önceden Silinmis..." 
				Else
					DeleteItem = "Silinen Dosya """ & itemName & """..."
				End If
		End Select
	End Function
	
	' UPLOAD Editor
	Sub Editor
		Dim fileName, filePath, file
		
		on error resume next

		Select Case Request.Form("subcommand")
			Case "Save", "SaveAs"
				fileName = SecureFileName(Request.Form("parameter"))
				filePath = destFolder & fileName
				Set file = FSO.OpenTextFile (filePath,2,true,0)
				If (err.Number<>0) Then 
					wexMessage = "Can not write to the file """ & fileName & """, permission denied!"
					err.Clear
				Else
					file.write Request.Form("content")
				End If
				Set file = Nothing
				Set file = FSO.OpenTextFile (filePath,1,false,0)
			Case Else
				fileName = SecureFileName(Request.Form("parameter"))
				filePath = destFolder & fileName
				
				If not FSO.FileExists(filePath) Then
					wexMessage = "The file """ & fileName & """ does not exist"
					Set file = FSO.CreateTextFile (filePath, false, false)
					If err.Number<>0 Then 
						wexMessage = wexMessage & ", also unable to create new file."
						err.Clear 
					Else
						wexMessage = wexMessage & ", created new file."
					End If
				Else
					Set file = FSO.OpenTextFile (filePath,1,false,0)
					If err.Number<>0 Then 
						wexMessage = "Can not read from the file """ & fileName & """, permission denied!"
						err.Clear 
					End If
				End If
		End Select
		
		HtmlHeader appName, wexCharSet
		If(wexMessage<>"") Then Response.Write "<script language=""javascript"">alert('" & wexMessage & "');</script>"
%>
<form name=formBuffer method=post action="<%=scriptName%>">
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=left>
					&nbsp;<font class=boldText>Degisiklik Yap</font> - <%=fileName%>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100% height=90%>
			<tr align=center class=lightRow>
				<td valign=middle>
<textarea name=content class=formClass rows=22 cols=46 style="width:580; height:370;" wrap="off">
<%=Server.HTMLEncode(file.ReadAll)%></textarea>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=center>					
					<a href="javascript:EditorCommand('Save');">Kaydet</a> | <a href="javascript:EditorCommand('SaveAs');">Farkli Kaydet</a> | <a href="javascript:EditorCommand('Reload');">Tekrar Yükle</a> | <a href="javascript:EditorCommand('Info');">Özellikler</a> | <a href="javascript:this.close();">KAPAT</a>
				</td>
			</tr>
		</table>
<%
		Set file = Nothing
		Set file = FSO.GetFile (filePath)
%>
<input type="hidden" name=command value="Edit">
<input type="hidden" name=subcommand value="">
<input type="hidden" name=parameter value="<%=fileName%>">
<input type="hidden" name=folder value="<%=Request.Form("folder")%>">
<input type="hidden" name=info value="Boyut: <%=FormatSize(file.Size)%>|TüR: <%=file.Type%>|Son Degisiklik: <%=file.DateCreated%>|Son Giris: <%=file.DateLastAccessed%>|Olusturuldugu Tarih: <%=file.DateLastModified%>">
</form>
<%
		Set file = Nothing
		HtmlFooter
		Response.End 
	End Sub

	' UPLOAD Viewer
	Sub Viewer
		Dim fileName, filePath, file, imageSrc

		on error resume next
		fileName = Request.Form("parameter")
		filePath = destFolder & fileName
		imageSrc = replace(VirtualPath(destFolder) & fileName, " ", "%20")

		HtmlHeader appName, wexCharSet
%>
<form name=formBuffer method=post action="<%=scriptName%>">
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=left>
					&nbsp;<font class=boldText>Viewing</font> - <%=fileName%>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100% height=90%>
			<tr align=center class=lightRow>
				<td valign=middle>
					<img src="<%=imageSrc%>">
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=center>					
					<a href="javascript:ViewerCommand('Reload');">Reload</a> | <a href="javascript:ViewerCommand('Info');">Info</a> | <a href="javascript:this.close();">KAPAT</a>
				</td>
			</tr>
		</table>
<%
		Set file = FSO.GetFile (filePath)
%>
<input type="hidden" name=command value="View">
<input type="hidden" name=subcommand value="Refresh">
<input type="hidden" name=parameter value="<%=fileName%>">
<input type="hidden" name=folder value="<%=Request.Form("folder")%>">
<input type="hidden" name=info value="Size: <%=FormatSize(file.Size)%>|Type: <%=file.Type%>|Created: <%=file.DateCreated%>|Last Accessed: <%=file.DateLastAccessed%>|Last Modified: <%=file.DateLastModified%>">
</form>
<%
		Set file = Nothing
		HtmlFooter
		Response.End 
	End Sub

	' File/Folder Details
	Sub Details
		Dim fileName, filePath, file

		on error resume next
		fileName = Request.Form("parameter")
		filePath = destFolder & fileName

		HtmlHeader appName, wexCharSet
%>
<form name=formBuffer method=post action="<%=scriptName%>">
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=left>
					&nbsp;<font class=boldText>Ayrinti</font> - <%=fileName%>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100% height=80%>
			<tr align=center class=lightRow>
				<td valign=middle>
<%
		If Request.Form("command") = "FileDetails" Then
				Set file = FSO.GetFile (filePath)
		Else
				Set file = FSO.GetFolder (filePath)
		End If
%>
				<table border=0 cellspacing=5 cellpadding=0>
					<tr><td><font class=boldText>Boyut:</font></td><td><%=FormatSize(file.Size)%></td></tr>
					<tr><td><font class=boldText>TüR:</font></td><td><%=file.Type%></td></tr>
					<tr><td><font class=boldText>Son Degisiklik:</font></td><td><%=file.DateCreated%></td></tr>
					<tr><td><font class=boldText>Son Giris:</font></td><td><%=file.DateLastAccessed%></td></tr>
					<tr><td><font class=boldText>Son Çikis:</font></td><td><%=file.DateLastModified%></td></tr>
				</table>
<%
		Set file = Nothing
%>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=center>					
					<a href="javascript:this.close();">KAPAT</a>
				</td>
			</tr>
		</table>
<input type="hidden" name=command value="<%=Request.Form("command")%>">
<input type="hidden" name=parameter value="<%=fileName%>">
<input type="hidden" name=folder value="<%=Request.Form("folder")%>">
</form>
<%
		HtmlFooter
		Response.End 
	End Sub
	
	' Uploads a file
	Sub Upload
		
		If Request.QueryString("command")="DoUpload" Then 
			destFolder = wexRootPath & Request.QueryString("folder")
			destFolder = FixPath(destFolder, "\")
			If len(destFolder) < len(wexRootPath) Then Response.End 
		End If
		
		HtmlHeader appName, wexCharSet
%>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=left>
					&nbsp;<font class=boldText>YüKLE</font> - <%=FSO.GetBaseName(destFolder)%>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100% height=80%>
			<tr align=center class=lightRow>
				<td valign=middle>
					<font class=boldText>
<%
		If Request.QueryString("command")="DoUpload" Then
	
			Dim Uploader
			Set Uploader = New classUploader
	
			Uploader.Upload()
	
			If not Uploader.uploaded Then
				Response.Write "Gönderilecek Dosya Yok"
			Else
				If Uploader.uploadedFile.FileName="" Then
					Response.Write "Gönderilecek Dosya Yok"
				Else
					If Uploader.uploadedFile.Save(destFolder) Then
						Response.Write Uploader.uploadedFile.FileName & " Basariyla Yüklendi<br>Boyut: "
						Response.Write FormatSize(Uploader.uploadedFile.FileSize) & " (" & Uploader.uploadedFile.FileSize & " bytes)<br>"
					Else
						Response.Write Uploader.uploadedFile.FileName & " can not be written<br>"
					End If
				End If
			End If
%>
					</font>
					<form name=formBuffer method=post action="<%=scriptName%>">
						<input type=hidden name=command value="Upload">
						<input type=hidden name=folder value="<%=Request.QueryString("folder")%>">
					</form>
<%	
		Else
%>
					<form enctype="multipart/form-data" name=formBuffer method=post action="<%=scriptName%>?command=DoUpload&folder=<%=server.URLEncode(Request.Form("folder"))%>">
						<input type=file name=file class=formClass size="20">
					</form>
<%
		End If
%>					
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=center>					
					<a href="javascript:Upload();">YüKLE</a> | <a href="javascript:this.close();">KAPAT</a>
				</td>
			</tr>
		</table>
<%
		HtmlFooter
		Response.End 
	End Sub

	' Class containing upload data parsing functions
	Class classUploader
		Public uploaded
		Public uploadedFile
	
		Public Default Sub Upload()
			Dim biData, sInputName
			Dim nPosBegin, nPosEnd, nPos, vDataBounds, nDataBoundPos
			Dim nPosFile, nPosBound
	
			biData = Request.BinaryRead(Request.TotalBytes)
			nPosBegin = 1
			nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(13)))
			
			If (nPosEnd-nPosBegin) <= 0 Then Exit Sub
			 
			vDataBounds = MidB(biData, nPosBegin, nPosEnd-nPosBegin)
			nDataBoundPos = InstrB(1, biData, vDataBounds)
			
			Do Until nDataBoundPos = InstrB(biData, vDataBounds & CByteString("--"))
				
				nPos = InstrB(nDataBoundPos, biData, CByteString("Content-Disposition"))
				nPos = InstrB(nPos, biData, CByteString("name="))
				nPosBegin = nPos + 6
				nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(34)))
				sInputName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
				nPosFile = InstrB(nDataBoundPos, biData, CByteString("filename="))
				nPosBound = InstrB(nPosEnd, biData, vDataBounds)
				
				If nPosFile <> 0 And  nPosFile < nPosBound Then
					Dim sFileName
					Set uploadedFile = New classUploadedFile
					
					nPosBegin = nPosFile + 10
					nPosEnd =  InstrB(nPosBegin, biData, CByteString(Chr(34)))
					sFileName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
					uploadedFile.FileName = Right(sFileName, Len(sFileName)-InStrRev(sFileName, "\"))
	
					nPos = InstrB(nPosEnd, biData, CByteString("Content-Type:"))
					nPosBegin = nPos + 14
					nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(13)))
					
					nPosBegin = nPosEnd+4
					nPosEnd = InstrB(nPosBegin, biData, vDataBounds) - 2
					uploadedFile.FileData = MidB(biData, nPosBegin, nPosEnd-nPosBegin)
					
					If uploadedFile.FileSize > 0 Then uploaded = true Else uploaded = false
				End If

				nDataBoundPos = InstrB(nDataBoundPos + LenB(vDataBounds), biData, vDataBounds)
			Loop
		End Sub

		' String to byte string conversion
		Private Function CByteString(sString)
			Dim index
			For index = 1 to Len(sString)
				CByteString = CByteString & ChrB(AscB(Mid(sString,index,1)))
			Next
		End Function

		' Byte string to string conversion
		Private Function CWideString(bsString)
			Dim index
			CWideString =""
			For index = 1 to LenB(bsString)
				CWideString = CWideString & Chr(AscB(MidB(bsString,index,1))) 
			Next
		End Function
	End Class

	' Class containing file data writing functions
	Class classUploadedFile
		Public FileName, FileData
		
		Public Property Get FileSize()
			FileSize = LenB(FileData)
		End Property
	
		Public Function Save(path)
			Dim file
			Dim index
		
			If path = "" Or FileName = "" Then 
				Save = false
				Exit Function
			End If
			path = FixPath(path, "\")
		
			If Not FSO.FolderExists(path) Then
				Save = false
				Exit Function
			End If
		
			on error resume next
			Set file = FSO.CreateTextFile(path & FileName, True)
			
			For index = 1 to LenB(FileData)
			    file.Write Chr(AscB(MidB(FileData,index,1)))
			Next
	
			file.Close
			If err.Number<>0 Then
				Save = false
			Else
				Save = true
			End If
		End Function
	End Class
' ------------------------------------------------------------

' - UPLOAD Main -----------------------------------------
	If not Secure() Then Login
	
	Dim folder, destFolder

	wexRootPath = RealizePath(wexRoot)

	If Request.QueryString("command")="DoUpload" Then Upload()
	
	destFolder = wexRootPath & Request.Form("folder")
	destFolder = FixPath(destFolder, "\")
	If len(destFolder) < len(wexRootPath) Then Response.End 
	
	Select Case Request.Form("command")
		Case "NewFile", "NewFolder"
			wexMessage = CreateItem()
		Case "Edit"
			Editor()
		Case "View"
			Viewer()
		Case "FileDetails", "FolderDetails"
			Details()
		Case "Upload"
			Upload()
		Case "Logout"
			Logout()
		Case "DeleteFile", "DeleteFolder"
			wexMessage = DeleteItem()
		Case "OpenFolder"
			If Request.Form("folder") = "" Then
				destFolder = wexRootPath & Request.Form("parameter")
			Else	
				destFolder = wexRootPath & FixPath(Request.Form("folder"),"\") & Request.Form("parameter")
			End If
			destFolder = FixPath(destFolder, "\")
			If len(destFolder) < len(wexRootPath) Then Response.End 
		Case "LevelUp"
			destFolder = FSO.GetParentFolderName(destFolder)
			destFolder = FixPath(destFolder, "\")
			If len(destFolder) < len(wexRootPath) Then Response.End 
	End Select

	on error resume next
	Set folder = FSO.GetFolder(destFolder)
	if err.Number<>0 Then wexMessage = "Error opening folder """ & destFolder & """"

	HtmlHeader appName, wexCharSet
	WriteListing (folder)
%>
<form method=post action="<%=scriptName%>" name=formBuffer>
	<input type=hidden name=command value="">
	<input type=hidden name=parameter value="">
	<input type=hidden name=virtual value="<%=VirtualPath(destFolder)%>">
	<input type=hidden name=folder value="<%=replace(lcase(destfolder), lcase(wexRootPath), "")%>">
</form>
<%
	Set folder = Nothing
	HtmlCopyright
	HtmlFooter
%>