Attribute VB_Name = "modEncrypt"
'---------------------------------------------------------------------------------------
' Module    : ENCRYPT
' DateTime  : 04/08/2006 14:29
' Author    : Dean
' Purpose   :
'---------------------------------------------------------------------------------------
Option Explicit

'----------------------------------------
'Name: encryptstr
'Description:
'----------------------------------------
Function encryptstr(messagetext As String) As String
    Dim code1 As String
    Dim Encryptcode As String
    Dim a As Integer
    Dim b As Integer
    Dim C As Integer
    Dim d As Integer
    Dim cr As String

 
    code1 = "tiger"
    Encryptcode = "12345"
    a = 0: b = 0: C = 0: cr = vbNullString


    Do While a < Len(messagetext)
        a = a + 1
        b = b + 1: If b > Len(code1) Then b = 1
        C = C + 1: If C > Len(Encryptcode) Then C = 1
        d = Asc(Mid$(messagetext, a, 1)) + Asc(Mid$(code1, b, 1)) + Asc(Mid$(Encryptcode, C, 1))
        '         100                        116                             49   = 265 = 10
Loop2:
        
        If d > 255 Then
            d = d - 255
            GoTo Loop2
        End If
        cr = cr + Chr$(d)
    Loop
    
'"–
'Œ
'À –  "

    ' w5r5ldk11rt5n
    '                    v
    '  w   5  r   5   l  d   k   1  1  r   t   5  n
    '  29,208,13,206,20, 10,   7,203,202, 26, 26,208, 9
    ' 119,53,114, 53,108,101,107, 49, 49,114,116,53,110' original ascii
    ' 119,53,114, 53,108,100,107, 49, 49,114,116,53
    '  29
    ' w5r5lek11rt5n

    encryptstr = cr                              ' set the return value

    '======================================================
    'END routine error handler
    '======================================================
    On Error GoTo 0: Exit Function

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure encryptstr  " ' call a custom error-handler subroutine


End Function

'---------------------------------------------------------------------------------------
' Procedure : decryptstr
' Author    : beededea
' Date      : 27/04/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function decryptstr(messagetext As String) As String

    Dim code1 As String
    Dim Encryptcode As String
    Dim a As Integer
    Dim b As Integer
    Dim C As Integer
    Dim d As Integer
    Dim cr As String
 
    On Error GoTo decryptstr_Error

    code1 = "tiger"
    Encryptcode = "12345"

    a = 0
    b = 0
    C = 0

    '========================
    ' decrypt
    '========================
    Do While a < Len(messagetext)
        a = a + 1
        b = b + 1
        If b > Len(code1) Then b = 1
        C = C + 1
        If C > Len(Encryptcode) Then C = 1
        d = Asc(Mid$(messagetext, a, 1)) - (Asc(Mid$(code1, b, 1)) + Asc(Mid$(Encryptcode, C, 1)))
        '         10                        116                             49
        
Loop3:

        If d < 1 Then
            d = d + 255
            GoTo Loop3
        End If
        cr = cr + Chr$(d)
    Loop
    
'"–
'Œ
'À –  "

    ' w5r5ldk11rt5n
    '                    v
    '  w   5  r   5   l  d   k   1  1  r   t   5  n
    '  29,208,13,206,20, 10,  7,203,202,26,26,208, 9
    ' 119,53,114,53,108,101,107,49, 49,114,116,53,110
    ' 119,53,114,53,108,100,107,49, 49,114,116,53,110
    '
    ' w5r5lek11rt5n

    decryptstr = cr                              ' set the return value


    On Error GoTo 0
    Exit Function

decryptstr_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure decryptstr of Module ENCRYPT"

End Function

