<!--#include file="../../INC/b_include.asp" -->
<!-- puzzle baslangic -->
	<script language="javascript">
		var max = 3;
		var score = 0;
		var moves = 0;
		var ex = 3;
		var ey = 3;
		
		function getElement15(form, name) {
			var k;
			var elements = form.elements;
			for (k = 0; k < elements.length; k++) {
				if (elements[k].name == name) return elements[k];
			}
		}
		
		function press15(form, button) {
			name = button.name;
			x = name.substring(0,1);
			y = name.substring(2,3);
			play15(form, (x-1+1), (y-1+1));
		}
		
		function shuffle15(form, num) {
			for (i = 0; i < num; i++) {
				x = Math.floor(Math.random(4) * 4);
				if (x == 0) { toggle15(form, ex, ey, ex + 1, ey); }
				else if (x == 1) { toggle15(form, ex, ey, ex - 1, ey); }
				else if (x == 2) { toggle15(form, ex, ey, ex, ey + 1); }
				else if (x == 3) { toggle15(form, ex, ey, ex, ey - 1); }
			}
		}
		
		function play15(form, x, y) {			
			if (Math.abs(ex - x) + Math.abs(ey - y) == 1) {
				done = toggle15(form, x, y, x+1, y);
				if (!done) { done = toggle15(form, x, y, x-1, y); }
				if (!done) { done = toggle15(form, x, y, x, y+1); }
				if (!done) { done = toggle15(form, x, y, x, y-1);	}
				moves++;
				if (check15(form)) {
					alert('Kazandiniz ' + moves + ' harekette!');
					resetboard15(form);
				}
			}
			
		}

		function showrules15() {
			rules = 'Puzzle For Mini-NUKE \n\n' 
				+ 'Oyunun oynanis sekli \n' 
				+ 'blokta 1 den 15 kadar sayilar var \n'
				+ 'bos olan button u tiklayarak\n'
				+ 'numaralarin yerlerini degistirin.\n'
				+ '1 den 15 e kadar sirasiyla yapabildiginiz taktirde\n'
				+ 'oyun bitmis olucak \n'
				+ 'oyun ilk acilista siralidir karistir tusuna basiniz.'
				+ '\n\n (Orjinali IrcShow.Net tarafýndan PHP-Nuke için kodlanmýþtýr.)';
			alert(rules);
		}
		
		function resetboard15(form) {
			for (i = 0; i < 4; i++) {
				for (j = 0; j < 4; j++) {
					val = 1 + i + (4*j);
					if (val == 16) {
						getElement15(form,i + '_' + j).value = ' ';
					} else {
						getElement15(form,i + '_' + j).value = val;
					}
				}
			}
			score = 0;
			moves = 0;
			ex = 3;
			ey = 3;
		}
	
		function toggle15(form, x, y, x1, y1) {
			if (x < 0 || y < 0 || x > max || y > max) {
				return false;
			}
			if (x1 < 0 || y1 < 0 || x1 > max || y1 > max) {
				return false;
			}

			name = x + '_' + y;
			button = getElement15(form,name);
			name = x1 + '_' + y1;
			button1 = getElement15(form,name);
			if (button.value == ' ' || button1.value == ' ') {
				tmp = button.value;
				button.value = button1.value;
				button1.value = tmp;
				if (button.value == ' ') {
					ex = x;
					ey = y;
				} else {
					ex = x1;
					ey = y1;
				}
				return true;
			}
			return false;
		}
		
		function check15(form) {
			score = 0;
			for (i = 0; i < 4; i++) {
				for (j = 0; j < 4; j++) {
					val = 1 + i + (4*j);
					if (val < 16) {
						if (getElement15(form,i + '_' + j).value == val) {
							score++;
						}
					}
				}
			}
			return score == 15;
		}
	</script>
	
	<table border="0" cellspacing="0" cellpadding="1" width="100%">
		<form>
		<tr align="center">
			<td><input style="width:25px;" type="button" name="0_0" value="1" onclick="javascript:press15(this.form, this);" class="submit"></td>
			<td><input style="width:25px;" type="button" name="1_0" value="2" onclick="javascript:press15(this.form, this);" class="submit"></td>
			<td><input style="width:25px;" type="button" name="2_0" value="3" onclick="press15(this.form, this);" class="submit"></td>
			<td><input style="width:25px;" type="button" name="3_0" value="4" onclick="press15(this.form, this);" class="submit"></td>
		</tr>
		<tr align="center">
			<td><input style="width:25px;" type="button" name="0_1" value="5" onclick="press15(this.form, this);" class="submit"></td>
			<td><input style="width:25px;" type="button" name="1_1" value="6" onclick="press15(this.form, this);" class="submit"></td>
			<td><input style="width:25px;" type="button" name="2_1" value="7" onclick="press15(this.form, this);" class="submit"></td>
			<td><input style="width:25px;" type="button" name="3_1" value="8" onclick="press15(this.form, this);" class="submit"></td>
		</tr>
		<tr align="center">
			<td><input style="width:25px;" type="button" name="0_2" value="9" onclick="press15(this.form, this);" class="submit"></td>
			<td><input style="width:25px;" type="button" name="1_2" value="10" onclick="press15(this.form, this);" class="submit"></td>
			<td><input style="width:25px;" type="button" name="2_2" value="11" onclick="press15(this.form, this);" class="submit"></td>
			<td><input style="width:25px;" type="button" name="3_2" value="12" onclick="press15(this.form, this);" class="submit"></td>
		</tr>
		<tr align="center">
			<td><input style="width:25px;" type="button" name="0_3" value="13" onclick="press15(this.form, this);" class="submit"></td>
			<td><input style="width:25px;" type="button" name="1_3" value="14" onclick="press15(this.form, this);" class="submit"></td>
			<td><input style="width:25px;" type="button" name="2_3" value="15" onclick="press15(this.form, this);" class="submit"></td>
			<td><input style="width:25px;" type="button" name="3_3" value=" " onclick="press15(this.form, this);" class="submit"></td>
		</tr>
		<tr>
			<td colspan="2">
				<input style="width:100%" type="button" value="Reset" onclick="resetboard15(this.form);" class="submit">
			</td>
			<td colspan="2">
				<input style="width:100%" type="button" value="Kural" onclick="showrules15();" class="submit">
			</td>
		</tr>
		<tr>
			<td colspan="4">
				<input style="width:100%;" type="button" value="Karýþtýr" onclick="shuffle15(this.form,150);" class="submit">
			</td>
		</tr>
		<tr>
                        <td colspan="4" style="font-family:Verdana; font-size:10px;">
                               
                        </td>
                </tr>
                </form>
        </table>
<!-- puzzle son  -->
</body>
</html></td>

