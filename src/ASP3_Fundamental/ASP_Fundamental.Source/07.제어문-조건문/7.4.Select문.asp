<%
'-------------------------------------------------------------------
' 인사말 기능을 ASP로 구현한 소스
' 작성자 : 박용준
' E-mail: redplus@redplus.net
' HomePage: http://www.redplus.net/
'-------------------------------------------------------------------
%>
<% 
'##############################[ 인삿말 지정 ]###################################
word00 = "오늘 또 하루가 지나가고 새 하루가 시작되었군요..."
word01 = "밤이 깊어가네요...출출하실텐데...라면 한그릇 어때요?"
word02 = "피곤하시죠? 어깨를 쭉펴고 기지개를 한번 펴 보세요... 아웅~"
word03 = "이 야심한밤에 이곳에 오시다니... 일찍 주무세요... "
word04 = "밤이 너무 깊었어요... 건강을 생각해서 이제 주무셔요..."
word05 = "아침이 밝아 옵니다. 알찬 계획을 세워 보세요..."
word06 = "좋은 아침~ Good morning?"
word07 = "일찍 오셨네요? 반가워요.... 헤헤"
word08 = "지금 어디시죠? 아마 집이시겠죠? 가족들께 안부 전해 주세요"
word09 = "아침 식사는 하셨나요? 식사 꼭 챙겨 드세요..."
word10 = "낮이 되어 가네요... 하루 계획을 세웁시다.. 알차게"
word11 = "배고프시죠... 조금만 참으세요... 즐거운 점심시간까지..."
word12 = "낮입니다. 영차 영차 열심히 일을 합시다...."
word13 = "점심은 드셨나요? 오늘도 즐거운 하루인가요?"
word14 = "해가 중천에 떳어요... 반나절이 지나가네요..."
word15 = "여러분들은 지금 무얼하고 있나요? 궁금하네요..."
word16 = "나른한 오후죠? 기분좋았던 일을 떠올려 보세요..."
word17 = "햐~ 벌써 하루가 지나가네요... 오늘 계획은 다 이루셨는지?"
word18 = "저녁 식사는 하셨나요? 건강을 위해선 식사를 꼭 하세요."
word19 = "지금 TV에 재미있는 프로는 하나요? 헤헤"
word20 = "밤이 되었습니다. 아직 못다한 일이 있으신지요?"
word21 = "뉴스에 뭐 좋은 소식이라도 나오나요? 있으면 알려 주세요..."
word22 = "오늘 하루 보람차게 보내셨나요? 하루를 되돌아 봅시다..."
word23 = "잠자리에 드실 시간이네요... 좋은꿈 꾸세요... "
'################################################################################

Select Case Hour(Now)
	Case 0 : insa = word00
	Case 8 : insa = word08
	Case 9 : insa = word09
	Case 10 : insa = word10
	Case 17 : insa = word17
	Case 18 : insa = word18
	Case 19 : insa = word19
	Case 20 : insa = word20
	Case 21 : insa = word21
	Case 22 : insa = word22
	Case 23 : insa = word23
	Case Else : insa = "다른시간대 메시지..."
End Select

'이 파일을 포함한 문서에서 
Response.Write(insa)'로 문자열(인사말)을 출력하면된다.
%>
 