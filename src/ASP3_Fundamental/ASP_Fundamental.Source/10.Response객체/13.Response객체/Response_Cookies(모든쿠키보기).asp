<h3> Cookie �ݷ��� </h3>
<hr>
<% 
   ' ��� ��Ű �����͸� �����´�.
   For Each Item in Request.Cookies

      ' Ű�� ������ �ִ� ���
      If Request.Cookies(Item).HasKeys Then

         ' ���� �׸��� ���� ���
         For Each ItemKey in Request.Cookies(Item)

             ' ���� �׸�� ���� �����ݴϴ�.
             Response.Write "<br>[���� �׸�] : " & Item & ":" & ItemKey
             Response.Write "=" & Request.Cookies(Item)(ItemKey)
         Next 
      Else
          ' �׸�� ���� �����ݴϴ�.
          Response.Write "<br>[�׸�] : " & Item
          Response.Write "=" & Request.Cookies(Item)
      End If

   Next 
%>

