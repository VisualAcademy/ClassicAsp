<%
Option Explicit
Dim strTest

strTest = " Abc Def Ghi ihG feD cbA "
%>

strTest에 저장된 문자열 : <%=strTest%> <br>
strTest에 저장된 문자열의 길이 : <%=Len(strTest)%> <br>
strTest에 저장된 문자열을 모두 대문자로 변환 : <%=UCase(strTest)%> <br>
strTest에 저장된 문자열을 모두 소문자로 변환 : <%=LCase(strTest)%> <br>
strTest에 저장된 문자열 중 Def의 시작위치 : <%=InStr(strTest,"Def")%> <br>
strTest에 저장된 문자열 중 f의 위치(앞에서) : <%=InStr(strTest,"f")%> <br>
strTest에 저장된 문자열 중 f의 위치(뒤에서) : <%=InStrRev(strTest,"f")%> <br>
strTest에 저장된 문자열 중 왼쪽에서 4자 검색(왼쪽에서) : <%=Left(strTest,4)%> <br>
strTest에 저장된 문자열 중 오른쪽에서 4자 검색(왼쪽에서) : <%=Right(strTest,4)%> <br>
strTest에 저장된 문자열 중 8번째에서 10번째까지(3문자)의 문자열 검색 : <%=Mid(strTest,8,3)%> <br>
strTest에 저장된 문자열 중 8번째에 이후의 문자열 검색 : <%=Mid(strTest,8)%> <br>
strTest에 저장된 문자열 중 왼쪽공백을 제거 : <%=LTrim(strTest)%> <br>
strTest에 저장된 문자열 중 왼쪽공백을 제거(문자열 길이) : <%=Len(LTrim(strTest))%> <br>
strTest에 저장된 문자열 중 오른쪽공백을 제거 : <%=RTrim(strTest)%> <br>
strTest에 저장된 문자열 중 오른쪽공백을 제거(문자열 길이) : <%=Len(RTrim(strTest))%> <br>
strTest에 저장된 문자열 중 양쪽공백을 제거 : <%=Trim(strTest)%> <br>
strTest에 저장된 문자열 중 양쪽공백을 제거(문자열 길이) : <%=Len(Trim(strTest))%> <br>