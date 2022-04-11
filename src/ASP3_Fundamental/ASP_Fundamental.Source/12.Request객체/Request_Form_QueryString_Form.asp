<html>
<body>

<h4> 텍스트, 암호, 텍스트 영역 입력 폼 예제 </h4>

<form action="Request_Form_QueryString_Form.asp?formno=1" method="post">
    <input type="text" name="text" value="" size="30" maxlength="30"><br>
    <input type="password" name="pass" size="30"  value=""><br>
    <textarea name="area" cols="30" rows="3"></textarea><P>
    <input type="Submit" value="보내기">&nbsp;&nbsp;<input type="reset">
</form>

<hr> 
<h4> 라디오 버튼과 첵크 박스 입력 폼 예제  </h4>

<form action="Request_Form_QueryString_Process.asp?formno=2" method="post">

<table border = 0>
<tr>
    <td width= 100>
    <input type="Radio" name="Radio" value="Radio1"> Radio 1 <br>
    <input type="Radio" name="Radio" value="Radio2"> Radio 2 <br>
    <input type="Radio" name="Radio" value="Radio3" checked> Radio 3
    </td>
    <td>
    <input type="Checkbox" name="Check" value="Check1"> Check 1<br>
    <input type="Checkbox" name="Check" value="Check2"> Check 2<br>
    <input type="Checkbox" name="Check" value="Check3"> Check 3
    </td>
</tr>
</table>
<br>
<input type="Submit" value="보내기">&nbsp;&nbsp;<input type="reset">
</form>

<hr>
<h4> 리스트와 리스트 박스 입력 폼 예제 </h4>

<form action="vform_.asp?formno=3" method="post">

    <select name="List" size="1">
        <option value="List1"> List 1 
        <option value="List2"> List 2
        <option value="List3"> List 3
    </select>

    <select name="ListBox" size="3" multiple>
        <option value="ListBox1"> ListBox 1
        <option value="ListBox2"> ListBox 2
        <option value="ListBox3"> ListBox 3
    </select><P>

    <input type="submit" value="보내기">&nbsp;&nbsp;<input type="reset">
</form>
<hr>

</body>
</html>
