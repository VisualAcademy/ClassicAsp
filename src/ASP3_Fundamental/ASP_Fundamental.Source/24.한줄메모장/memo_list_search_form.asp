<%
'--------------------------------------
' MyMemo Education Version
' Copyright (C) 2001 Park Yong Jun
' E-mail: bittobio@empal.com
' Support: http://www.xpplus.net/
'--------------------------------------
%>
<HTML>
	<HEAD>
		<META name="GENERATOR" content="Microsoft Visual Studio 6.0">
	</HEAD>
	<BODY>
		<DIV align="center">
			<TABLE border="0">
				<FORM name="memo_list_search_form" method="post" action="./memo_list_search_output.asp">
					<TR>
						<TD>
							검색 : <SELECT name="strsearchfield" size="1">
								<OPTION value="name">
									이름</OPTION>
								<OPTION value="title">
									메모</OPTION>
							</SELECT><INPUT type="text" name="strsearchquery" size="20"> <INPUT type="submit" value="검 색">
						</TD>
					</TR>
				</FORM>
			</TABLE>
		</DIV>
	</BODY>
</HTML>
