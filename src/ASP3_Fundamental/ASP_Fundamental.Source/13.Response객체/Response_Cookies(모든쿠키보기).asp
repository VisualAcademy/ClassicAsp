<h3> Cookie 콜렉션 </h3>
<hr>
<% 
   ' 모든 쿠키 데이터를 가져온다.
   For Each Item in Request.Cookies

      ' 키를 가지고 있는 경우
      If Request.Cookies(Item).HasKeys Then

         ' 하위 항목을 가진 경우
         For Each ItemKey in Request.Cookies(Item)

             ' 하위 항목과 값을 보여줍니다.
             Response.Write "<br>[하위 항목] : " & Item & ":" & ItemKey
             Response.Write "=" & Request.Cookies(Item)(ItemKey)
         Next 
      Else
          ' 항목과 값을 보여줍니다.
          Response.Write "<br>[항목] : " & Item
          Response.Write "=" & Request.Cookies(Item)
      End If

   Next 
%>

