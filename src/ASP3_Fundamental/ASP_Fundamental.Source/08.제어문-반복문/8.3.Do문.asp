<%
     'Do~Loop������ 1~100������ �� �� ¦���� ���� ���ϴ� ���α׷�
     Dim Sum, Count

     Sum = 0
     Count = 1
     Do Until Count > 100   '���ǽ��� ������ ����...(���� ������)
          If Count Mod 2 = 0 Then
               Sum = Sum + Count
          End If
          Count = Count + 1   'ī���� ���� ����
     Loop

     Response.Write("<center>1���� 100������ ¦���� ���� " & Sum & " �Դϴ�</center>")

     Sum = 0
     Count = 1
     Do While Count <= 100   '���ǽ��� ���ϵ���...
          If Count Mod 2 = 0 Then
               Sum = Sum + Count
          End If
          Count = Count + 1   'ī���� ���� ����
     Loop

     Response.Write("<center>1���� 100������ ¦���� ���� " & Sum & " �Դϴ�</center>")

     Sum = 0
     Count = 1
     Do    '�ּ��� �ѹ��� �����Ѵ�.
          If Count Mod 2 = 0 Then
               Sum = Sum + Count
          End If
          Count = Count + 1   'ī���� ���� ����
     Loop Until Count > 100   '���ǽ��� ������ ����...(���� ������)

     Response.Write("<center>1���� 100������ ¦���� ���� " & Sum & " �Դϴ�</center>")

     Sum = 0
     Count = 1
     Do    '�ּ��� �ѹ��� �����Ѵ�.
          If Count Mod 2 = 0 Then
               Sum = Sum + Count
          End If
          Count = Count + 1   'ī���� ���� ����
     Loop While Count <= 100   '���ǽ��� ���� ����...(������ ������)

     Response.Write("<center>1���� 100������ ¦���� ���� " & Sum & " �Դϴ�</center>")
%>