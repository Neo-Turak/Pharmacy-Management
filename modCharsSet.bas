Attribute VB_Name = "modCharsSet"

'Download by http://www.NewXing.com
Option Explicit

Public Const ERR_RESULT$ = "?"          ' �����Ĵ��󷵻�ֵ

'==================================================
' ����: ReturnSM
'
' ����: �����ַ����е�һ���ַ�����ĸ
'
' ע��: �ú���ֻ�ܴ���3755�����ú���(B0 - D7)
'       �����������ķ�Χ�����س��� ERR_RESULT$
'
' ���: Str     ��������ַ���
'
Public Function ReturnSM$(ByVal Str$)
'-------------------------------------------------
    Dim tmpStr$, tmpASCII&
    
    ' ȡ���ַ����еĵ�һ���ַ�
    tmpStr$ = Left(Str$, 1)
    
    ' ��tmpStr����Ϊ 0 ,�����޷���ֵ
    If Len(tmpStr$) <= 0 Then Exit Function
    
    ' �����ַ�ӳ����е��ַ���
    tmpASCII& = VBA.Asc(tmpStr$)
    
    ' ����tmpStr,����������ĸ,����������Χ,�򷵻ش���
    Select Case tmpASCII&
        Case &HB0A1 To &HB0C4
            
            ReturnSM$ = "A"
            
        Case &HB0C5 To &HB0FE, &HB1A1 To &HB1FE, _
             &HB2A1 To &HB2C0
            
            ReturnSM$ = "B"
            
        Case &HB2C1 To &HB2FE, &HB3A1 To &HB3FE, _
             &HB4A1 To &HB4ED
            
            ReturnSM$ = "C"
        
        Case &HB4EE To &HB4FE, &HB5A1 To &HB5FE, _
             &HB6A1 To &HB6E9
            
            ReturnSM$ = "D"
        
        Case &HB6EA To &HB6FE, &HB7A1
            
            ReturnSM$ = "E"
        
        Case &HB7A2 To &HB7FE, &HB8A1 To &HB8C0
            
            ReturnSM$ = "F"
        
        Case &HB8C1 To &HB8FE, &HB9A1 To &HB9FD
            
            ReturnSM$ = "G"
        
        Case &HB9FE, &HBAA1 To &HBAFE, &HBBA1 To &HBBF6
            
            ReturnSM$ = "H"
        
        Case &HBBF7 To &HBBFE, &HBCA1 To &HBCFE, _
             &HBDA1 To &HBDFE, &HBEA1 To &HBEFE, _
             &HBFA1 To &HBFA5
            
            ReturnSM$ = "J"
        
        Case &HBFA6 To &HBFFE, &HC0A1 To &HC0AB
            
            ReturnSM$ = "K"
        
        Case &HC0AC To &HC0FE, &HC1A1 To &HC1FE, _
             &HC2A1 To &HC2E7
            
            ReturnSM$ = "L"
        
        Case &HC2E8 To &HC2FE, &HC3A1 To &HC3FE, _
             &HC4A1 To &HC4C2
            
            ReturnSM$ = "M"
        
        Case &HC4C3 To &HC4FE, &HC5A1 To &HC5B5
            
            ReturnSM$ = "N"
        
        Case &HC5B6 To &HC5BD
            
            ReturnSM$ = "O"
        
        Case &HC5BE To &HC5FE, &HC6A1 To &HC6D9
            
            ReturnSM$ = "P"
        
        Case &HC6DA To &HC6FE, &HC7A1 To &HC7FE, _
             &HC8A1 To &HC8BA
            
            ReturnSM$ = "Q"
        
        Case &HC8BB To &HC8F5
            
            ReturnSM$ = "R"
        
        Case &HC8F6 To &HC8FE, &HC9A1 To &HC9FE, _
             &HCAA1 To &HCAFE, &HCBA1 To &HCBF9
            
            ReturnSM$ = "S"
        
        Case &HCBFA To &HCBFE, &HCCA1 To &HCCFE, _
             &HCDA1 To &HCDD9
            
            ReturnSM$ = "T"
        
        Case &HCDDA To &HCDFE, &HCEA1 To &HCEF3
            
            ReturnSM$ = "W"
        
        Case &HCEF4 To &HCEFE, &HCFA1 To &HCFFE, _
             &HD0A1 To &HD0FE, &HD1A1 To &HD1B8
            
            ReturnSM$ = "X"
        
        Case &HD1B9 To &HD1FE, &HD2A1 To &HD2FE, _
             &HD3A1 To &HD3FE, &HD4A1 To &HD4D0
            
            ReturnSM$ = "Y"
        
        Case &HD4D1 To &HD4FE, &HD5A1 To &HD5FE, _
             &HD6A1 To &HD6FE, &HD7A1 To &HD7F9
            
            ReturnSM$ = "Z"
        
        Case Else:  GoTo Err:
    End Select
    
    Exit Function
Err:
    ' ���������Ĵ���Χ
    ReturnSM$ = ERR_RESULT$
'-------------------------------------------------
End Function
'==================================================
