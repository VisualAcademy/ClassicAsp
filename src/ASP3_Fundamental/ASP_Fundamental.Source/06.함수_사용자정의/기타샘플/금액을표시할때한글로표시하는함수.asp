�Ʒ� �ҽ��� ���� VBScript�� ���� ����� �ǵ���
�̷� �ڹٽ�ũ��Ʈ�� ��ȯ ��Ű�� �ɰ̴ϴ�.

function ConvertAmount()
Dim strNum(15)
Dim strMoney(4)
Dim strLength
Dim strTotalAmount
Dim strTemp
Dim Character
Dim PartCount
Dim i
Dim Count
Dim Amount

Amount = Pass.intAmount.value
strNum(1) = "��"
strNum(2) = "��"
strNum(3) = "��"
strNum(4) = "��"
strNum(5) = "��"
strNum(6) = "��"
strNum(7) = "ĥ"
strNum(8) = "��"
strNum(9) = "��"

strMoney(2) = "��"
strMoney(3) = "��"
strMoney(4) = "õ"


    strLength = Len(Amount)
        
    Count = 0
    PartCount = 0
    
    For i = strLength To 1 Step -1
        
        strTemp = Mid(Amount, i, 1)
        Count = Count + 1
        
        If Count = 6 Then
           if right(Amount,10)="00,000,000" then
           else
              strTotalAmount = "��" & strTotalAmount
            end if  
        ElseIf Count = 11 Then
            strTotalAmount = "��" & strTotalAmount
        ElseIf Count = 16 Then
            strTotalAmount = "��" & strTotalAmount
        End If
        
        If strTemp <> "," Then
            PartCount = PartCount + 1
            If PartCount > 1 Then
                If CInt(strTemp) <> 0 Then
                    strTotalAmount = strMoney(PartCount) & strTotalAmount
                End If
            End If
            If CInt(strTemp) <> 0 Then
                strTotalAmount = strNum(CInt(strTemp)) & strTotalAmount
            End If
        End If
        
        If PartCount = 4 Then
            PartCount = 0
        End If
        
        strTemp = ""
    Next
    
Pass.txtAmount1.value = strTotalAmount
Pass.txtAmount.value = strTotalAmount
end function
