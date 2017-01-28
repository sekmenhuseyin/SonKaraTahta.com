<%'Kod Sahibi (Code Author) : Web Wiz Guide - Web Wiz Forums
'****************************************************************************************
'**  Copyright Notice    
'**  Web Wiz Guide - Web Wiz Forums
'**  Copyright 2001-2004 Bruce Corkhill All Rights Reserved.                                
'**  This program is free software; you can modify (at your own risk) any part of it 
'**  under the terms of the License that accompanies this software and use it both 
'**  privately and commercially.
'**  All copyright notices must remain in tacked in the scripts and the 
'**  outputted HTML.
'**  You may use parts of this program in your own private work, but you may NOT
'**  redistribute, repackage, or sell the whole or any part of this program even 
'**  if it is modified or reverse engineered in whole or in part without express 
'**  permission from the author.
'**  You may not pass the whole or any part of this application off as your own work.
'**  All links to Web Wiz Guide and powered by logo's must remain unchanged and in place
'**  and must remain visible when the pages are viewed unless permission is first granted
'**  by the copyright holder.
'**  This program is distributed in the hope that it will be useful,
'**  but WITHOUT ANY WARRANTY; without even the implied warranty of
'**  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE OR ANY OTHER 
'**  WARRANTIES WHETHER EXPRESSED OR IMPLIED.
'**  You should have received a copy of the License along with this program; 
'**  if not, write to:- Web Wiz Guide, PO Box 4982, Bournemouth, BH8 8XP, United Kingdom.
'**  No official support is available for this program but you may post support questions at: -
'**  http://www.webwizguide.info/forum
'**  Support questions are NOT answered by e-mail ever!
'**  For correspondence or non support questions contact: -
'**  info@webwizguide.info
'**  or at: -
'**  Web Wiz Guide, PO Box 4982, Bournemouth, BH8 8XP, United Kingdom
'****************************************************************************************
'Dimension variables
Dim saryHTMLtags(88)	 'If you add more disallowed HTML tags then increase the array size
'Initalise array values
saryHTMLtags(0)="html"
saryHTMLtags(1)="body"
saryHTMLtags(2)="head"
saryHTMLtags(3)="meta"
saryHTMLtags(4)="button"
saryHTMLtags(5)="input"
saryHTMLtags(6)="type"
saryHTMLtags(7)="select"
saryHTMLtags(8)="radio"
saryHTMLtags(9)="file"
saryHTMLtags(10)="hidden"
saryHTMLtags(11)="checkbox"
saryHTMLtags(12)="password"
saryHTMLtags(13)="checked"
saryHTMLtags(14)="fieldset"
saryHTMLtags(15)="language"
saryHTMLtags(16)="javascript"
saryHTMLtags(17)="vbscript"
saryHTMLtags(18)="script"
saryHTMLtags(19)="object"
saryHTMLtags(20)="applet"
saryHTMLtags(21)="embed"
saryHTMLtags(22)="event"
saryHTMLtags(23)="server"
saryHTMLtags(24)="function"	
saryHTMLtags(25)="document"
saryHTMLtags(26)="cookie"
saryHTMLtags(27)="onclick"
saryHTMLtags(28)="ondblclick"
saryHTMLtags(29)="onkeyup"
saryHTMLtags(30)="onkeydown"
saryHTMLtags(31)="onkeypress"
saryHTMLtags(32)="onkey"
saryHTMLtags(33)="onmouseenter"
saryHTMLtags(34)="onmouseleave"
saryHTMLtags(35)="onmousemove"
saryHTMLtags(36)="onmouseout"
saryHTMLtags(37)="onmouseover"
saryHTMLtags(38)="onrollover"
saryHTMLtags(39)="onmouse"
saryHTMLtags(40)="onchange"
saryHTMLtags(41)="onunloadhave"
saryHTMLtags(42)="onunload"
saryHTMLtags(43)="onsubmit"
saryHTMLtags(44)="onselect"
saryHTMLtags(45)="accesskey"
saryHTMLtags(46)="tabindex"
saryHTMLtags(47)="onfocus"
saryHTMLtags(48)="onblur"
saryHTMLtags(49)="onsubmit"
saryHTMLtags(50)="onreset"
saryHTMLtags(51)="form"
saryHTMLtags(52)="iframe"
saryHTMLtags(53)="ilayer"
saryHTMLtags(54)="textarea"
saryHTMLtags(55)="action"
saryHTMLtags(56)="enctype"
saryHTMLtags(57)="layer"
saryHTMLtags(58)="multicol"
saryHTMLtags(59)="frameset"
saryHTMLtags(60)="marquee"
saryHTMLtags(61)="blink"
saryHTMLtags(62)="filter"
saryHTMLtags(63)="overlay"
saryHTMLtags(64)="param"
saryHTMLtags(65)="bgsound"
saryHTMLtags(66)="behavior"
saryHTMLtags(67)="ismap"
saryHTMLtags(68)="sound"
saryHTMLtags(69)="disabled"
saryHTMLtags(70)="ENCTYPE"
saryHTMLtags(71)="!DOCTYPE"
saryHTMLtags(72)="BACKGROUND-COLOR"
saryHTMLtags(73)="base"
saryHTMLtags(74)="position" 'Thanks to Aytek Ustundag for his help (HolyOne, http://www.tahribat.com)
saryHTMLtags(75)="absolute"
saryHTMLtags(76)="z-index"
saryHTMLtags(77)="isindex"
saryHTMLtags(78)="xhtml"
saryHTMLtags(79)="xml"
saryHTMLtags(80)="class"
saryHTMLtags(81)="map"
saryHTMLtags(82)="option"
saryHTMLtags(83)="box"
saryHTMLtags(84)="SAMP"
saryHTMLtags(85)="data"
saryHTMLtags(86)="frame"
'If you add more disallowed HTML tags don't forget to increase the number in the Dim statement at the top!%>