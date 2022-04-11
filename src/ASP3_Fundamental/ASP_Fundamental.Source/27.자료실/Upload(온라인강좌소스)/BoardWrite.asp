<pre>
<form action="./BoardWriteProcess.asp" method="post">
이름: <input type="text" name="Name">
이메일:<input type="text" name="Email">
홈페이지:<input type="text" name="Homepage">
제목:<input type="text" name="Title">
내용:<textarea name="Content" cols=40 rows=10></textarea>
인코딩:
	<select name="Encoding">
		<option value="Text">Text</option>
		<option value="HTML">HTML</option>
		<option value="Mixed">Mixed</option>
	</select>
<!--자료실에서 추가-->
파일첨부 : <input type="file" name="FileName">
비밀번호<input type="password" name="Password">
<input type="submit" value="전송">
</form>
</pre>