<%Function duz(Str)
Str=Replace(Str, "'", "´", 1, -1, 1)
Str=Replace(Str, "<", "&lt;", 1, -1, 1)
Str=Replace(Str, ">", "&gt;", 1, -1, 1)
Str=Replace(Str, vbCrLf, "<br />", 1, -1, 1)
Str=Replace(Str, "  ", " &nbsp;", 1, -1, 1)
Str=Replace(Str, "   ", " &nbsp; ", 1, -1, 1)
Str=Replace(Str, "[B]", "<b>", 1, -1, 1)
Str=Replace(Str, "[/B]", "</b>", 1, -1, 1)
Str=Replace(Str, "[I]", "<i>", 1, -1, 1)
Str=Replace(Str, "[/I]", "</i>", 1, -1, 1)
Str=Replace(Str, "[U]", "<u>", 1, -1, 1)
Str=Replace(Str, "[/U]", "</u>", 1, -1, 1)
Str=Replace(Str, "[CENTER]", "<center>", 1, -1, 1)
Str=Replace(Str, "[/CENTER]", "</center>", 1, -1, 1)
Str=ReplaceTag(Str, "[IMG]", "[/IMG]", "http://", "<img src=""%REPLACE%"" border=""0"">")
Str=ReplaceTag(Str, "[URL]", "[/URL]", "http://", "<a href=""%REPLACE%"" target=""_blank"">%REPLACE%</a>")
Do While InStr(1, Str, "[URL=", 1) > 0 AND InStr(1, Str, "[/URL]", 1) > 0
'Find the start position in the message of the [URL= code
lngStartPos=InStr(1, Str, "[URL=", 1)
'Find the position in the message for the [/URL] closing code
lngEndPos=InStr(lngStartPos, Str, "[/URL]", 1) + 6
'Make sure the end position is not in error
If lngEndPos < lngStartPos Then lngEndPos=lngStartPos + 7
'Read in the code to be converted into a hyperlink from the message
strMessageLink=Trim(Mid(Str, lngStartPos, (lngEndPos - lngStartPos)))
'Place the message link into the tempoary message variable
strTempMessage=strMessageLink
'Format the link into an HTML hyperlink
strTempMessage=Replace(strTempMessage, "[URL=", "<a href=""", 1, -1, 1)
'If there is no tag shut off place a > at the end
If InStr(1, strTempMessage, "[/URL]", 1) Then
strTempMessage=Replace(strTempMessage, "[/URL]", "</a>", 1, -1, 1)
strTempMessage=Replace(strTempMessage, "]", """>", 1, -1, 1)
Else
strTempMessage=strTempMessage & ">"
End If
'Place the new fromatted hyperlink into the message string body
Str=Replace(Str, strMessageLink, strTempMessage, 1, -1, 1)
Loop
Do While InStr(1, Str, "[EMAIL=", 1) > 0 AND InStr(1, Str, "[/EMAIL]", 1) > 0
'Find the start position in the message of the [EMAIL= code
lngStartPos=InStr(1, Str, "[EMAIL=", 1)
'Find the position in the message for the [/EMAIL] closing code
lngEndPos=InStr(lngStartPos, Str, "[/EMAIL]", 1) + 8
'Make sure the end position is not in error
If lngEndPos < lngStartPos Then lngEndPos=lngStartPos + 9
'Read in the code to be converted into a email link from the message
strMessageLink=Trim(Mid(Str, lngStartPos, (lngEndPos - lngStartPos)))
'Place the message link into the tempoary message variable
strTempMessage=strMessageLink
'Format the link into an HTML mailto link
strTempMessage=Replace(strTempMessage, "[EMAIL=", "<a href=""mailto:", 1, -1, 1)
'If there is no tag shut off place a > at the end
If InStr(1, strTempMessage, "[/EMAIL]", 1) Then
strTempMessage=Replace(strTempMessage, "[/EMAIL]", "</a>", 1, -1, 1)
strTempMessage=Replace(strTempMessage, "]", """>", 1, -1, 1)
Else
strTempMessage=strTempMessage & ">"
End If
'Place the new fromatted HTML mailto into the message string body
Str=Replace(Str, strMessageLink, strTempMessage, 1, -1, 1)
Loop
Do While InStr(1, Str, "[COLOR=", 1) > 0  AND InStr(1, Str, "[/COLOR]", 1) > 0
'Find the start position in the message of the [COLOR= code
lngStartPos=InStr(1, Str, "[COLOR=", 1)
'Find the position in the message for the [/COLOR] closing code
lngEndPos=InStr(lngStartPos, Str, "[/COLOR]", 1) + 8
'Make sure the end position is not in error
If lngEndPos < lngStartPos Then lngEndPos=lngStartPos + 9
'Read in the code to be converted into a font colour from the message
strMessageLink=Trim(Mid(Str, lngStartPos, (lngEndPos - lngStartPos)))
'Place the message colour into the tempoary message variable
strTempMessage=strMessageLink
'Format the link into an font colour HTML tag
strTempMessage=Replace(strTempMessage, "[COLOR=", "<font color=", 1, -1, 1)
'If there is no tag shut off place a > at the end
If InStr(1, strTempMessage, "[/COLOR]", 1) Then
strTempMessage=Replace(strTempMessage, "[/COLOR]", "</font>", 1, -1, 1)
strTempMessage=Replace(strTempMessage, "]", ">", 1, -1, 1)
Else
strTempMessage=strTempMessage & ">"
End If
'Place the new fromatted colour HTML tag into the message string body
Str=Replace(Str, strMessageLink, strTempMessage, 1, -1, 1)
Loop
'Hear for backward compatability with old colour codes abive
Str=Replace(Str, "[/COLOR]", "</font>", 1, -1, 1)
duz=Str
End Function

Function MakeURL(ByVal myString)
Set nnRegExp=New RegExp
nnRegExp.Pattern="\b( www | http:// | ftp | file:// )\S+\b" 
nnRegExp.IgnoreCase=True 
nnRegExp.Global=True 
Set nnMatches=nnRegExp.Execute(myString)
For Each nnMatch In nnMatches
If Left(nnMatch,3)=" www " Then
myString=Replace(myString,nnMatch,"<a href="" http:// "& nnMatch &""" target=""_blank"">"& nnMatch &"</a>",1,-1,1)
Else
myString=Replace(myString,nnMatch,"<a href="""& nnMatch &""" target=""_blank"">"& nnMatch &"</a>",1,-1,1)
End If
Next       
MakeURL=myString
End Function

Function Oku(Str)
Str=Replace(Str, "´", "'", 1, -1, 1)
Str=Replace(Str, "&lt;", "<", 1, -1, 1)
Str=Replace(Str, "&gt;", ">", 1, -1, 1)
Str=Replace(Str, "<br />", vbCrLf, 1, -1, 1)
Str=Replace(Str, "<b>", "[B]", 1, -1, 1)
Str=Replace(Str, "</b>", "[/B]", 1, -1, 1)
Str=Replace(Str, "<i>", "[I]", 1, -1, 1)
Str=Replace(Str, "</i>", "[/I]", 1, -1, 1)
Str=Replace(Str, "<u>", "[U]", 1, -1, 1)
Str=Replace(Str, "</u>", "[/U]", 1, -1, 1)
Str=Replace(Str, "<center>", "[CENTER]", 1, -1, 1)
Str=Replace(Str, "</center>", "[/CENTER]", 1, -1, 1)
Oku=Str
End Function

Function EmailCheck(Str)
Et_Isareti=InStr(2, Str , "@" )
If Et_Isareti=VBIsNull Then
EmailCheck=False
Else
Et_Isareti_Krakter_Sayisi=Et_Isareti
Et_Isareti=True
End If
If Et_Isareti=True Then
Nokta=InStr(Et_Isareti_Krakter_Sayisi + 2, Str , "." )
If Nokta=VBIsNull Then EmailCheck=False Else EmailCheck=True
Else
EmailCheck=False
End If
End Function%>