<%
'--------------------------------------------------
' Title : Basic ����
' Program Name : boarddelete.asp
' Program Description : ����� ��
' QueryString : �ݵ��  view.asp���� Num=1 ������ ����
' Include Files : None
' Copyright (C) 2004 Park Yong Jun
' E-mail: redplus@redplus.net
' Support: http://www.dotnetkorea.com/
'--------------------------------------------------
%>
<%
Dim Num: Num = Request("Num")
%>
<%=Num%>�� ���� �����Ͻðڽ��ϱ�?<br>
<form action="./boarddeleteprocess.asp"
method="post">
<!--��ȣ : --><input type="hidden" name="Num" 
		value="<%=Num%>"><br>
����� ��й�ȣ : <input type="password" name="Password">
<br>
<input type="submit" value="�����ϱ�">
<input type="button" value="���" onclick="history.go(-1);">
</form>