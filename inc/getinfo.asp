<%
  ua=Request.ServerVariables("HTTP_USER_AGENT")
  If Instr(ua, "MSIE") Then
   If Instr(ua, "MSIE 6.") Then
	 Browser="Microsoft Internet Explorer 6.x"
   ElseIf Instr(ua, "MSIE 5.") Then
	 Browser="Microsoft Internet Explorer 5.x"
   Elseif Instr(ua, "MSIE 4.") Then
	 Browser="Microsoft Internet Explorer 4.x"
   Elseif Instr(ua, "MSIE 3.") Then
	 Browser="Microsoft Internet Explorer 3.x"
		 If Instr(ua, "MSIE 3.02") Then
    		 Browser="Microsoft Internet Explorer 3.02"
		 End If
   Else
	 Browser="Microsoft Internet Explorer"
   End If
  ElseIf Instr(ua, "Mozilla") and Instr(ua, "compatible")=0 Then
	  If Instr(ua, "Mozilla/4") Then
		 Browser="Netscape Navigator 4.x"
	  Elseif Instr(ua, "Mozilla/3") Then
		 Browser="Netscape Navigator 3.x"
      Else
		 Browser="Netscape Navigator"
	  End If
  End If
   If Browser="" OR IsNull(os) OR Len(Browser) < 3 Then Browser="Bilinmeyen"

	  If Instr(ua, "Windows 95") or Instr(ua, "Win95") Then
		  os="Windows 95"
	  Elseif Instr(ua, "Windows 98") or Instr(ua, "Win98") Then
		  os="Windows 98"
	  Elseif Instr(ua, "Windows 2000") or Instr(u, "Win2000") Then
		  os="Windows 2000"
	  Elseif Instr(ua, "NT") or Instr(ua, "NT") Then
		  os="Windows NT"
	  Elseif Instr(ua, "XP") or Instr(ua, "Windows XP") Then
		  os="Windows XP"
	  Elseif Instr(ua, "Mac") Then
		  os="Mac"
			  If Instr(ua, "PowerPC") or Instr(ua, "PPC") Then
				 os="Mac PPC"
			  Elseif Instr(ua, "68000") or Instr(ua, "68K") Then
				 os="Mac 68K"
			  End If
	 Elseif Instr(ua, "X11") Then
			  os="UNIX"
	 End If
	If os="" OR IsNull(os) OR Len(os) < 3 Then os="Bilinmeyen"
%>