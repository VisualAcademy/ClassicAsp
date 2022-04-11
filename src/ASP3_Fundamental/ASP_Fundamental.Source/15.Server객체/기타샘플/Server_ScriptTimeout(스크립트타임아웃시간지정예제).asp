<%

	'서버객체의 ScriptTimeout속성은 스크립트 처리시간을 제어하는 속성이다.
	'현재 이소스는 사이트갤럭시 업로드 컴포넌트를 사용해서 자료를 올리는 소스이다.
	'아래는 자료올리는 시간을 15분까지 지정한 구문이다.

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Server.ScriptTimeout = 900    
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''        

        Set UploadForm = Server.CreateObject("SiteGalaxyUpload.Form") 

        Set FSO = CreateObject("Scripting.FileSystemObject")

        file1 = UploadForm("file1") 
        file1 = Trim(file1)        

        IF Len(file1) > 0 Then '첨부한 파일이 있으면
            strDirectory = "D:\DATA\" '서버내 파일이 저장되는 위치지정-여기서는 D:\ 
            attach_file = UploadForm("file1").FilePath    '저장될 파일의 경로 얻음(클라이언트 파일 위치)            
            FileSize = UploadForm("file1").Size '파일 사이즈            
            FileName = Mid(attach_file, InstrRev(attach_file, "\")+1) '파일명.확장자 를 얻음!
            strName = Mid(FileName, 1, Instr(FileName, ".")-1) '파일명 추출
            strExt = Mid(FileName, Instr(FileName,".")+1) '확장자 추출

            bExist = True '같은 파일이 존재한다고 가정
            strFileName = strDirectory&strName&"."&strExt '서버에 저장할 경로와 파일 설정
            countFileName = 0 '같은 파일일때 파일명뒤에 붙일 숫자
            
            If bExist = True Then
                IF (FSO.FileExists(strFileName)) Then '같은 파일이 존재하면
                    countFileName = countFileName+1
                    FileName = strName&"_"&countFileName&"."&strExt
                    strFileName = strDirectory&FileName '실제 서버에 저장될 경로와 파일명
                Else
                    bExist = False
                End IF
            End IF
            Response.Write "실제서버에 저장될 파일명 : "&strFileName&"<br>"

            UploadForm("file1").SaveAs strFileName '서버에 파일을 저장
        End IF

        Set FSO = Nothing
        Set UploadForm = Nothing
        Response.Write "<br>"&strFileName&"이 저장되었습니다."
%>

