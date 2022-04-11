<form name="SearchForm" action="./BoardSearchList.asp" method="post">

검색 : 
<select name="SearchField">
	<option value="Name">작성자</option>
	<option value="Title">제목</option>
	<option value="Content">내용</option>
	<option value="All">전체검색</option>
</select>

<input type="text" name="SearchQuery">

<input type="submit" value="검색">

</form>