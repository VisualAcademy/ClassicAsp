아래 소스는 제가 VBScript로 만들어서 사용한 건데요
이럴 자바스크립트로 변환 시키면 될겁니다.

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
strNum(1) = "일"
strNum(2) = "이"
strNum(3) = "삼"
strNum(4) = "사"
strNum(5) = "오"
strNum(6) = "육"
strNum(7) = "칠"
strNum(8) = "팔"
strNum(9) = "구"

strMoney(2) = "십"
strMoney(3) = "백"
strMoney(4) = "천"


    strLength = Len(Amount)
        
    Count = 0
    PartCount = 0
    
    For i = strLength To 1 Step -1
        
        strTemp = Mid(Amount, i, 1)
        Count = Count + 1
        
        If Count = 6 Then
           if right(Amount,10)="00,000,000" then
           else
              strTotalAmount = "만" & strTotalAmount
            end if  
        ElseIf Count = 11 Then
            strTotalAmount = "억" & strTotalAmount
        ElseIf Count = 16 Then
            strTotalAmount = "조" & strTotalAmount
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
