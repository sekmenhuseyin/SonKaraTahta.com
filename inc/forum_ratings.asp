<%IF isnumeric(f_mem_msg)=false then f_mem_msg=0
IF f_mem_level="0"  Then

IF f_mem_msg < 50 Then
	fm_level=forum_level_0
	fm_img="0"
ElseIF f_mem_msg >= 50 AND f_mem_msg < 100 Then
	fm_level=forum_level_1
	fm_img="1"
ElseIF f_mem_msg >= 100 AND f_mem_msg < 200 Then
	fm_level=forum_level_2
	fm_img="2"
ElseIF f_mem_msg >= 200 AND f_mem_msg < 500 Then
	fm_level=forum_level_3
	fm_img="3"
ElseIF f_mem_msg >= 500 AND f_mem_msg < 1000 Then
	fm_level=forum_level_4
	fm_img="4"
ElseIF f_mem_msg >= 1000 Then
	fm_level=forum_level_5
	fm_img="5"
ElseIF f_mem_msg >= 1500 Then
	fm_level=forum_level_6
	fm_img="5"	
End IF

ELSE

IF f_mem_level="1" Then
	fm_level=level1
	fm_img="5"
ElseIF f_mem_level="2" Then
	fm_level=level2
	fm_img="4"
ElseIF f_mem_level="3" Then
	fm_level=level3
	fm_img="4"
ElseIF f_mem_level="4" Then
	fm_level=level4
	fm_img ="4"
ElseIF f_mem_level="5" Then
	fm_level=level5
	fm_img="5"
ElseIF f_mem_level="6" Then
	fm_level=level6
	fm_img="5"
ElseIF f_mem_level="7" Then
	fm_level=level7
	fm_img="5"
ElseIF f_mem_level="8" Then
	fm_level=level8
	fm_img="4"		
End IF

End IF

fm_include_img="<img src=""IMAGES/star_ratings/"&fm_img&"_star_rating.gif"">"
%>