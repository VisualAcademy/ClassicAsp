<%

	'������ü�� ScriptTimeout�Ӽ��� ��ũ��Ʈ ó���ð��� �����ϴ� �Ӽ��̴�.
	'���� �̼ҽ��� ����Ʈ������ ���ε� ������Ʈ�� ����ؼ� �ڷḦ �ø��� �ҽ��̴�.
	'�Ʒ��� �ڷ�ø��� �ð��� 15�б��� ������ �����̴�.

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Server.ScriptTimeout = 900    
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''        

        Set UploadForm = Server.CreateObject("SiteGalaxyUpload.Form") 

        Set FSO = CreateObject("Scripting.FileSystemObject")

        file1 = UploadForm("file1") 
        file1 = Trim(file1)        

        IF Len(file1) > 0 Then '÷���� ������ ������
            strDirectory = "D:\DATA\" '������ ������ ����Ǵ� ��ġ����-���⼭�� D:\ 
            attach_file = UploadForm("file1").FilePath    '����� ������ ��� ����(Ŭ���̾�Ʈ ���� ��ġ)            
            FileSize = UploadForm("file1").Size '���� ������            
            FileName = Mid(attach_file, InstrRev(attach_file, "\")+1) '���ϸ�.Ȯ���� �� ����!
            strName = Mid(FileName, 1, Instr(FileName, ".")-1) '���ϸ� ����
            strExt = Mid(FileName, Instr(FileName,".")+1) 'Ȯ���� ����

            bExist = True '���� ������ �����Ѵٰ� ����
            strFileName = strDirectory&strName&"."&strExt '������ ������ ��ο� ���� ����
            countFileName = 0 '���� �����϶� ���ϸ�ڿ� ���� ����
            
            If bExist = True Then
                IF (FSO.FileExists(strFileName)) Then '���� ������ �����ϸ�
                    countFileName = countFileName+1
                    FileName = strName&"_"&countFileName&"."&strExt
                    strFileName = strDirectory&FileName '���� ������ ����� ��ο� ���ϸ�
                Else
                    bExist = False
                End IF
            End IF
            Response.Write "���������� ����� ���ϸ� : "&strFileName&"<br>"

            UploadForm("file1").SaveAs strFileName '������ ������ ����
        End IF

        Set FSO = Nothing
        Set UploadForm = Nothing
        Response.Write "<br>"&strFileName&"�� ����Ǿ����ϴ�."
%>

