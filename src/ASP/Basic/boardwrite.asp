<%
'--------------------------------------------------
' Title : Basic ����
' Program Name : boardwrite.asp
' Program Description : �۾��� �� ������
' Include Files : None
' Copyright (C) 2004 Park Yong Jun
' E-mail: redplus@redplus.net
' Support: http://www.dotnetkorea.com/
'--------------------------------------------------
%>
<form action="./boardwriteprocess.asp" method="post">
<pre>
�̸� : <input type="text" name="Name">
�̸��� : <input type="text" name="Email">
Ȩ������ : <input type="text" name="Homepage">
���� : <input type="text" name="Title">
���� : <textarea name="Content" cols="40" rows="10"></textarea>
���ڵ� : <select name="Encoding"><option value="Text">Text</option><option value="HTML">HTML</option></select>
�н����� : <input type="password" name="Password">
<input type="submit" value="�ۼ� �Ϸ�">
</pre>
</form>