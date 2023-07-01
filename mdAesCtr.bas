Attribute VB_Name = "modAES"
'--- mdAesCtr.bas
Option Explicit
DefObj A-Z

#Const ImplUseShared = False

'=========================================================================
' API
'=========================================================================

'--- for CryptStringToBinary
Private Const CRYPT_STRING_BASE64           As Long = 1
'--- for WideCharToMultiByte
Private Const CP_UTF8                       As Long = 65001
'--- for FormatMessage
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt" (phAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptGetProperty Lib "bcrypt" (ByVal hObject As Long, ByVal pszProperty As Long, pbOutput As Any, ByVal cbOutput As Long, cbResult As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptSetProperty Lib "bcrypt" (ByVal hObject As Long, ByVal pszProperty As Long, ByVal pbInput As Long, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptGenerateSymmetricKey Lib "bcrypt" (ByVal hAlgorithm As Long, phKey As Long, pbKeyObject As Any, ByVal cbKeyObject As Long, pbSecret As Any, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDestroyKey Lib "bcrypt" (ByVal hKey As Long) As Long
Private Declare Function BCryptEncrypt Lib "bcrypt" (ByVal hKey As Long, pbInput As Any, ByVal cbInput As Long, ByVal pPaddingInfo As Long, ByVal pbIV As Long, ByVal cbIV As Long, pbOutput As Any, ByVal cbOutput As Long, pcbResult As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDeriveKeyPBKDF2 Lib "bcrypt" (ByVal pPrf As Long, pbPassword As Any, ByVal cbPassword As Long, pbSalt As Any, ByVal cbSalt As Long, ByVal cIterations As Long, ByVal dwDummy As Long, pbDerivedKey As Any, ByVal cbDerivedKey As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptCreateHash Lib "bcrypt" (ByVal hAlgorithm As Long, phHash As Long, ByVal pbHashObject As Long, ByVal cbHashObject As Long, pbSecret As Any, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDestroyHash Lib "bcrypt" (ByVal hHash As Long) As Long
Private Declare Function BCryptHashData Lib "bcrypt" (ByVal hHash As Long, pbInput As Any, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptFinishHash Lib "bcrypt" (ByVal hHash As Long, pbOutput As Any, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
Private Declare Function htonl Lib "ws2_32" (ByVal hostlong As Long) As Long
Private Declare Function RtlGenRandom Lib "advapi32" Alias "SystemFunction036" (RandomBuffer As Any, ByVal RandomBufferLength As Long) As Long
#If Not ImplUseShared Then
    Private Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByRef pcbBinary As Long, ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long
    Private Declare Function CryptBinaryToString Lib "crypt32" Alias "CryptBinaryToStringW" (ByVal pbBinary As Long, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, pcchString As Long) As Long
    Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
    Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
    Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Args As Any) As Long
#End If

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const AES_BLOCK_SIZE        As Long = 16
Private Const AES_KEYLEN            As Long = 32                    '-- 32 -> AES-256, 24 -> AES-196, 16 -> AES-128
Private Const AES_IVLEN             As Long = AES_BLOCK_SIZE
Private Const KDF_SALTLEN           As Long = 8
Private Const KDF_ITER              As Long = 10000
Private Const KDF_HASH              As String = "SHA512"
Private Const HMAC_HASH             As String = "SHA256"
Private Const OPENSSL_MAGIC         As String = "Salted__"          '-- for openssl compatibility
Private Const OPENSSL_MAGICLEN      As Long = 8
Private Const ERR_UNSUPPORTED_ENCR  As String = "Unsupported encryption"

Private Type UcsCryptoContextType
    hPbkdf2Alg          As Long
    hHmacAlg            As Long
    hHmacHash           As Long
    HashLen             As Long
    hAesAlg             As Long
    hAesKey             As Long
    AesKeyObjData()     As Byte
    AesKeyObjLen        As Long
    Nonce(0 To 3)       As Long
    EncrData()          As Byte
    EncrPos             As Long
    LastError           As String
End Type

'=========================================================================
' Functions
'=========================================================================

'--- equivalent to `openssl aes-256-ctr -pbkdf2 -md sha512 -pass pass:{sPassword} -in {sText}.file -a`
Public Function AesEncryptString(sText As String, sPassword As String) As String
    Const PREFIXLEN     As Long = OPENSSL_MAGICLEN + KDF_SALTLEN
    Dim baData()        As Byte
    Dim baSalt(0 To KDF_SALTLEN - 1) As Byte
    Dim sError          As String
    
    baData = ToUtf8Array(sText)
    Call RtlGenRandom(baSalt(0), KDF_SALTLEN)
'    If Not AesCryptArray(baData, ToUtf8Array(sPassword), baSalt, Error:=sError) Then
'        err.Raise vbObjectError, , sError
'    End If
    ReDim Preserve baData(0 To UBound(baData) + PREFIXLEN) As Byte
    If UBound(baData) >= PREFIXLEN Then
        Call CopyMemory(baData(PREFIXLEN), baData(0), UBound(baData) + 1 - PREFIXLEN)
    End If
    Call CopyMemory(baData(OPENSSL_MAGICLEN), baSalt(0), KDF_SALTLEN)
    Call CopyMemory(baData(0), ByVal OPENSSL_MAGIC, 8)
    AesEncryptString = Replace(ToBase64Array(baData), vbCrLf, vbNullString)
End Function

'--- equivalent to `openssl aes-256-ctr -pbkdf2 -md sha512 -pass pass:{sPassword} -in {sEncr}.file -a -d`
Public Function AesDecryptString(sEncr As String, sPassword As String) As String
    Const PREFIXLEN     As Long = OPENSSL_MAGICLEN + KDF_SALTLEN
    Dim baData()        As Byte
    Dim baSalt()        As Byte
    Dim sMagic          As String
    Dim sError          As String
    
    baData = AESFromBase64Array(sEncr)
    baSalt = vbNullString
    If UBound(baData) >= PREFIXLEN - 1 Then
        sMagic = String$(OPENSSL_MAGICLEN, 0)
        Call CopyMemory(ByVal sMagic, baData(0), OPENSSL_MAGICLEN)
        If sMagic = OPENSSL_MAGIC Then
            ReDim baSalt(0 To KDF_SALTLEN - 1) As Byte
            Call CopyMemory(baSalt(0), baData(OPENSSL_MAGICLEN), KDF_SALTLEN)
            If UBound(baData) >= PREFIXLEN Then
                Call CopyMemory(baData(0), baData(PREFIXLEN), UBound(baData) + 1 - PREFIXLEN)
                ReDim Preserve baData(0 To UBound(baData) - PREFIXLEN) As Byte
            Else
                baData = vbNullString
            End If
        End If
    End If
'    If Not AesCryptArray(baData, ToUtf8Array(sPassword), baSalt, Error:=sError) Then
'        err.Raise vbObjectError, , sError
'    End If
    AesDecryptString = FromUtf8Array(baData)
End Function

Public Function AesCryptArray( _
            baData() As Byte, _
            baPass() As Byte, _
            Optional Salt As Variant, _
            Optional ByVal KeyLen As Long, _
            Optional Error As String, _
            Optional Hmac As Variant) As Boolean
    Const VT_BYREF      As Long = &H4000
    Dim uCtx            As UcsCryptoContextType
    Dim vErr            As Variant
    Dim bHashBefore     As Boolean
    Dim bHashAfter      As Boolean
    Dim baTemp()        As Byte
    Dim lPtr            As Long
    
    On Error GoTo EH
    If IsArray(Hmac) Then
        bHashBefore = (Hmac(0) <= 0)
        bHashAfter = (Hmac(0) > 0)
    End If
    If IsMissing(Salt) Then
        baTemp = vbNullString
    ElseIf IsArray(Salt) Then
        baTemp = Salt
    Else
        baTemp = ToUtf8Array(CStr(Salt))
    End If
    If KeyLen <= 0 Then
        KeyLen = AES_KEYLEN
    End If
    If Not pvCryptoAesCtrInit(uCtx, baPass, baTemp, KeyLen) Then
        Error = uCtx.LastError
        GoTo QH
    End If
    If Not pvCryptoAesCtrCrypt(uCtx, baData, HashBefore:=bHashBefore, HashAfter:=bHashAfter) Then
        Error = uCtx.LastError
        GoTo QH
    End If
    If IsArray(Hmac) Then
        baTemp = pvCryptoGetFinalHash(uCtx, UBound(Hmac) + 1)
        lPtr = Peek((VarPtr(Hmac) Xor &H80000000) + 8 Xor &H80000000)
        If (Peek(VarPtr(Hmac)) And VT_BYREF) <> 0 Then
            lPtr = Peek(lPtr)
        End If
        lPtr = Peek((lPtr Xor &H80000000) + 12 Xor &H80000000)
        Call CopyMemory(ByVal lPtr, baTemp(0), UBound(baTemp) + 1)
    End If
    '--- success
    AesCryptArray = True
QH:
    pvCryptoAesCtrTerminate uCtx
    Exit Function
EH:
    vErr = Array(err.Number, err.Source, err.Description)
    pvCryptoAesCtrTerminate uCtx
    err.Raise vErr(0), vErr(1), vErr(2)
End Function

'= private ===============================================================

Private Function pvCryptoAesCtrInit(uCtx As UcsCryptoContextType, baPass() As Byte, baSalt() As Byte, ByVal lKeyLen As Long) As Boolean
    Const MS_PRIMITIVE_PROVIDER         As String = "Microsoft Primitive Provider"
    Const BCRYPT_ALG_HANDLE_HMAC_FLAG   As Long = 8
    Const BCRYPT_HASH_REUSABLE_FLAG     As Long = &H20
    Dim baDerivedKey()  As Byte
    Dim lResult         As Long '--- discarded
    
    With uCtx
        '--- init member vars
        .EncrData = vbNullString
        .EncrPos = 0
        '--- generate RFC 2898 based derived key
        On Error GoTo EH_Unsupported '--- CNG API missing on XP
        If BCryptOpenAlgorithmProvider(.hPbkdf2Alg, StrPtr(KDF_HASH), StrPtr(MS_PRIMITIVE_PROVIDER), BCRYPT_ALG_HANDLE_HMAC_FLAG) <> 0 Then
            GoTo QH
        End If
        On Error GoTo 0
        ReDim baDerivedKey(0 To lKeyLen + AES_IVLEN - 1) As Byte
        On Error GoTo EH_Unsupported '--- PBKDF2 API missing on Vista
        If BCryptDeriveKeyPBKDF2(.hPbkdf2Alg, ByVal pvArrayPtr(baPass), pvArraySize(baPass), ByVal pvArrayPtr(baSalt), pvArraySize(baSalt), _
                KDF_ITER, 0, baDerivedKey(0), UBound(baDerivedKey) + 1, 0) <> 0 Then
            GoTo QH
        End If
        On Error GoTo 0
        '--- init AES key from first half of derived key
        If BCryptOpenAlgorithmProvider(.hAesAlg, StrPtr("AES"), StrPtr(MS_PRIMITIVE_PROVIDER), 0) <> 0 Then
            GoTo QH
        End If
        If BCryptGetProperty(.hAesAlg, StrPtr("ObjectLength"), .AesKeyObjLen, 4, lResult, 0) <> 0 Then
            GoTo QH
        End If
        If BCryptSetProperty(.hAesAlg, StrPtr("ChainingMode"), StrPtr("ChainingModeECB"), 30, 0) <> 0 Then ' 30 = LenB("ChainingModeECB")
            GoTo QH
        End If
        ReDim .AesKeyObjData(0 To .AesKeyObjLen - 1) As Byte
        If BCryptGenerateSymmetricKey(.hAesAlg, .hAesKey, .AesKeyObjData(0), .AesKeyObjLen, baDerivedKey(0), lKeyLen, 0) <> 0 Then
            GoTo QH
        End If
        '--- init AES IV from second half of derived key
        Call CopyMemory(.Nonce(0), baDerivedKey(lKeyLen), AES_IVLEN)
        '--- init HMAC key from last HashLen bytes of derived key
        If BCryptOpenAlgorithmProvider(.hHmacAlg, StrPtr(HMAC_HASH), StrPtr(MS_PRIMITIVE_PROVIDER), BCRYPT_ALG_HANDLE_HMAC_FLAG) <> 0 Then
            GoTo QH
        End If
        If BCryptGetProperty(.hHmacAlg, StrPtr("HashDigestLength"), .HashLen, 4, lResult, 0) <> 0 Then
            GoTo QH
        End If
        If BCryptCreateHash(.hHmacAlg, .hHmacHash, 0, 0, baDerivedKey(lKeyLen + AES_IVLEN - .HashLen), .HashLen, BCRYPT_HASH_REUSABLE_FLAG) <> 0 Then
            GoTo QH
        End If
    End With
    '--- success
    pvCryptoAesCtrInit = True
    Exit Function
QH:
    uCtx.LastError = GetSystemMessage(err.LastDllError)
    Exit Function
EH_Unsupported:
    uCtx.LastError = ERR_UNSUPPORTED_ENCR
End Function

Private Sub pvCryptoAesCtrTerminate(uCtx As UcsCryptoContextType)
    With uCtx
        If .hPbkdf2Alg <> 0 Then
            Call BCryptCloseAlgorithmProvider(.hPbkdf2Alg, 0)
            .hPbkdf2Alg = 0
        End If
        If .hHmacHash <> 0 Then
            Call BCryptDestroyHash(.hHmacHash)
            .hHmacHash = 0
        End If
        If .hHmacAlg <> 0 Then
            Call BCryptCloseAlgorithmProvider(.hHmacAlg, 0)
            .hHmacAlg = 0
        End If
        If .hAesKey <> 0 Then
            Call BCryptDestroyKey(.hAesKey)
            .hAesKey = 0
        End If
        If .hAesAlg <> 0 Then
            Call BCryptCloseAlgorithmProvider(.hAesAlg, 0)
            .hAesAlg = 0
        End If
    End With
End Sub

Private Function pvCryptoAesCtrCrypt( _
            uCtx As UcsCryptoContextType, _
            baData() As Byte, _
            Optional ByVal Offset As Long, _
            Optional ByVal Size As Long = -1, _
            Optional ByVal HashBefore As Boolean, _
            Optional ByVal HashAfter As Boolean) As Boolean
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lPadSize        As Long
    
    With uCtx
        If Size < 0 Then
            Size = pvArraySize(baData) - Offset
        End If
        If HashBefore Then
            If BCryptHashData(.hHmacHash, ByVal pvArrayPtr(baData, Offset), Size, 0) <> 0 Then
                GoTo QH
            End If
        End If
        '--- reuse .EncrData from prev call until next AES_BLOCK_SIZE boundary
        For lIdx = Offset To Offset + Size - 1
            If (.EncrPos And (AES_BLOCK_SIZE - 1)) = 0 Then
                Exit For
            End If
            baData(lIdx) = baData(lIdx) Xor .EncrData(.EncrPos)
            .EncrPos = .EncrPos + 1
        Next
        If lIdx < Offset + Size Then
            '--- pad remaining input size to AES_BLOCK_SIZE
            lPadSize = (Offset + Size - lIdx + AES_BLOCK_SIZE - 1) And -AES_BLOCK_SIZE
            If UBound(.EncrData) + 1 < lPadSize Then
                ReDim .EncrData(0 To lPadSize - 1) As Byte
            End If
            '--- encrypt incremental Nonce in .EncrData
            For lJdx = 0 To lPadSize - 1 Step AES_BLOCK_SIZE
                Call CopyMemory(.EncrData(lJdx), .Nonce(0), AES_BLOCK_SIZE)
                If pvInc(.Nonce(3)) Then
                    If pvInc(.Nonce(2)) Then
                        If pvInc(.Nonce(1)) Then
                            If pvInc(.Nonce(0)) Then
                                '--- do nothing
                            End If
                        End If
                    End If
                End If
            Next
            If BCryptEncrypt(.hAesKey, .EncrData(0), lPadSize, 0, 0, 0, .EncrData(0), lPadSize, lJdx, 0) <> 0 Then
                GoTo QH
            End If
            '--- XOR remaining input and leave anything extra in .EncrData for reuse
            For .EncrPos = 0 To Offset + Size - lIdx - 1
                baData(lIdx) = baData(lIdx) Xor .EncrData(.EncrPos)
                lIdx = lIdx + 1
            Next
        End If
        If HashAfter Then
            If BCryptHashData(.hHmacHash, ByVal pvArrayPtr(baData, Offset), Size, 0) <> 0 Then
                GoTo QH
            End If
        End If
    End With
    '--- success
    pvCryptoAesCtrCrypt = True
    Exit Function
QH:
    uCtx.LastError = GetSystemMessage(err.LastDllError)
End Function

Private Function pvCryptoGetFinalHash(uCtx As UcsCryptoContextType, ByVal lSize As Long) As Byte()
    Dim baResult()      As Byte
    
    ReDim baResult(0 To uCtx.HashLen - 1) As Byte
    Call BCryptFinishHash(uCtx.hHmacHash, baResult(0), uCtx.HashLen, 0)
    ReDim Preserve baResult(0 To lSize - 1) As Byte
    pvCryptoGetFinalHash = baResult
End Function

Private Function pvInc(lValue As Long) As Boolean
    lValue = htonl(lValue)
    If lValue = -1 Then
        lValue = 0
        '--- signal carry
        pvInc = True
    Else
        lValue = (lValue Xor &H80000000) + 1 Xor &H80000000
        lValue = htonl(lValue)
    End If
End Function

Private Property Get pvArrayPtr(baArray() As Byte, Optional ByVal Index As Long) As Long
    Dim lPtr            As Long
    
    '--- peek long at ArrPtr(baArray)
    Call CopyMemory(lPtr, ByVal ArrPtr(baArray), 4)
    If lPtr <> 0 Then
        If 0 <= Index And Index <= UBound(baArray) - LBound(baArray) Then
            pvArrayPtr = VarPtr(baArray(LBound(baArray) + Index))
        End If
    End If
End Property

Private Property Get pvArraySize(baArray() As Byte) As Long
    Dim lPtr            As Long
    
    '--- peek long at ArrPtr(baArray)
    Call CopyMemory(lPtr, ByVal ArrPtr(baArray), 4)
    If lPtr <> 0 Then
        pvArraySize = UBound(baArray) + 1 - LBound(baArray)
    End If
End Property

'= shared ================================================================

#If Not ImplUseShared Then
Public Function ToBase64Array(baData() As Byte) As String
    Dim lSize           As Long
    
    If UBound(baData) >= 0 Then
        ToBase64Array = String$(2 * UBound(baData) + 6, 0)
        lSize = Len(ToBase64Array) + 1
        Call CryptBinaryToString(VarPtr(baData(0)), UBound(baData) + 1, CRYPT_STRING_BASE64, StrPtr(ToBase64Array), lSize)
        ToBase64Array = Left$(ToBase64Array, lSize)
    End If
End Function

Public Function AESFromBase64Array(sText As String) As Byte()
    Dim lSize           As Long
    Dim baOutput()      As Byte
    
    lSize = Len(sText) + 1
    ReDim baOutput(0 To lSize - 1) As Byte
    Call CryptStringToBinary(StrPtr(sText), Len(sText), CRYPT_STRING_BASE64, VarPtr(baOutput(0)), lSize, 0, 0)
    If lSize > 0 Then
        ReDim Preserve baOutput(0 To lSize - 1) As Byte
        AESFromBase64Array = baOutput
    Else
        AESFromBase64Array = vbNullString
    End If
End Function

Public Function ToUtf8Array(sText As String) As Byte()
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), ByVal 0, 0, 0, 0)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), baRetVal(0), lSize, 0, 0)
    Else
        baRetVal = vbNullString
    End If
    ToUtf8Array = baRetVal
End Function

Public Function FromUtf8Array(baText() As Byte) As String
    Dim lSize           As Long
    
    If UBound(baText) >= 0 Then
        FromUtf8Array = String$(2 * UBound(baText), 0)
        lSize = MultiByteToWideChar(CP_UTF8, 0, baText(0), UBound(baText) + 1, StrPtr(FromUtf8Array), Len(FromUtf8Array))
        FromUtf8Array = Left$(FromUtf8Array, lSize)
    End If
End Function

Public Function GetSystemMessage(ByVal lLastDllError As Long) As String
    Dim lSize            As Long
   
    GetSystemMessage = Space$(2000)
    lSize = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lLastDllError, 0&, GetSystemMessage, Len(GetSystemMessage), 0&)
    If lSize > 2 Then
        If Mid$(GetSystemMessage, lSize - 1, 2) = vbCrLf Then
            lSize = lSize - 2
        End If
    End If
    GetSystemMessage = "[" & lLastDllError & "] " & Left$(GetSystemMessage, lSize)
End Function

Private Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
End Function
#End If
