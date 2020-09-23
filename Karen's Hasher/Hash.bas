Attribute VB_Name = "Hash"
Option Explicit
' Copyright © 2002, 2004-2005 Karen Kenworthy
' All Rights Reserved
' http://www.karenware.com/
' Version 2.2 7/13/2005

Private Const HASH_FREQ = 75

Public Enum HASH_TYPE
    HASH_NONE = -1
    HASH_MD5 = 1
    HASH_TYPE_LBOUND = HASH_MD5
    HASH_SHA1
    HASH_SHA224
    HASH_SHA256
    HASH_SHA384
    HASH_SHA512
    HASH_TYPE_UBOUND = HASH_SHA512
End Enum

Public Enum HASH_RESULT
    HASH_DEFAULT = -1
    HASH_OK = 0
    HASH_BADSIG = &H80000000
    HASH_BADFID = &H80000001
    HASH_BADSIGLEN = &H80000002
    HASH_IO_ERROR = &H80000003
    HASH_BADTYPE = &H80000004
    HASH_CANCELLED = &H80000009
End Enum

Public Enum HASH_DISP_FORMAT
    HASH_DISP_PLAIN = 0
    HASH_DISP_PRETTY = 1
    HASH_DISP_LCASE = 0
    HASH_DISP_UCASE = 2
End Enum

Private Declare Function PTHashDescA Lib "PTHash" ( _
    ByVal buf As String, _
    ByVal buflen As Long) As Long

Private Declare Function PTHashVersion Lib "PTHash" () As Long

Private Declare Function PTHashMajorVersion Lib "PTHash" () As Long

Private Declare Function PTHashMinorVersion Lib "PTHash" () As Long

Private Declare Function PTHashRevision Lib "PTHash" () As Long

Private Declare Function PTHashIOBlock Lib "PTHash" () As Long


''''''''''''''''''''''''''''
'''''''''  SHA1  '''''''''''
Private Declare Function PTMD5BlockLen Lib "PTHash" () As Long

Private Declare Function PTMD5SigLen Lib "PTHash" () As Long

Private Declare Function PTMD5CtxLen Lib "PTHash" () As Long

Private Declare Function PTMD5Init Lib "PTHash" ( _
    ByRef ctx As Byte) As HASH_RESULT

Private Declare Function PTMD5Input Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByRef buf As Byte, _
    ByVal buflen As Long) As HASH_RESULT

Private Declare Function PTMD5Fini Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByRef sig As Byte) As HASH_RESULT

'Private Declare Function PTMD5StringX Lib "PTHash" ( _
'    ByRef ctx As Byte, _
'    ByRef buf As Byte, _
'    ByVal buflen As Long, _
'    ByRef sig As Byte) As Long

Private Declare Function PTMD5String Lib "PTHash" ( _
    ByRef buf As Byte, _
    ByVal buflen As Long, _
    ByRef sig As Byte) As HASH_RESULT

Private Declare Function PTMD5FileA Lib "PTHash" ( _
    ByVal szfid As String, _
    ByRef sig As Byte) As HASH_RESULT

Private Declare Function PTMD5FileProgA Lib "PTHash" ( _
    ByVal szfid As String, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As HASH_RESULT

Private Declare Function PTMD5FileProgW Lib "PTHash" ( _
    ByRef szfid As Byte, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As HASH_RESULT

Private Declare Function PTMD5FilesA Lib "PTHash" ( _
    ByVal szfids As String, _
    ByRef sig As Byte) As HASH_RESULT

Private Declare Function PTMD5FilesProgInputA Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByVal szfids As String, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As HASH_RESULT

Private Declare Function PTMD5FilesProgA Lib "PTHash" ( _
    ByVal szfids As String, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As HASH_RESULT

Private Declare Function PTMD5FilesProgInputW Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByRef szfids As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As HASH_RESULT

Private Declare Function PTMD5FilesProgW Lib "PTHash" ( _
    ByRef szfids As Byte, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As HASH_RESULT

Private Declare Function PTMD5Hmac Lib "PTHash" ( _
    ByVal Text As String, _
    ByVal text_len As Long, _
    ByVal Key As String, _
    ByVal key_len As Long, _
    ByVal digest As String) As HASH_RESULT


''''''''''''''''''''''''''''
'''''''''  SHA1  '''''''''''
Private Declare Function PTSHA1BlockLen Lib "PTHash" () As Long

Private Declare Function PTSHA1SigLen Lib "PTHash" () As Long

Private Declare Function PTSHA1CtxLen Lib "PTHash" () As Long

Private Declare Function PTSHA1Reset Lib "PTHash" ( _
    ByRef ctx As Byte) As HASH_RESULT

Private Declare Function PTSHA1Input Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByRef buf As Byte, _
    ByVal buflen As Long) As HASH_RESULT

Private Declare Function PTSHA1Result Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByRef sig As Byte) As HASH_RESULT

Private Declare Function PTSHA1String Lib "PTHash" ( _
    ByRef buf As Byte, _
    ByVal buflen As Long, _
    ByRef sig As Byte) As Long

Private Declare Function PTSHA1FileA Lib "PTHash" ( _
    ByVal szfid As String, _
    ByRef sig As Byte) As Long

Private Declare Function PTSHA1FileProgA Lib "PTHash" ( _
    ByVal szfid As String, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As Long

Private Declare Function PTSHA1FileProgW Lib "PTHash" ( _
    ByRef szfid As Byte, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As Long

Private Declare Function PTSHA1FilesA Lib "PTHash" ( _
    ByVal szfids As String, _
    ByRef sig As Byte) As Long

Private Declare Function PTSHA1FilesProgInputA Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByVal szfids As String, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As HASH_RESULT

Private Declare Function PTSHA1FilesProgA Lib "PTHash" ( _
    ByVal szfids As String, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As Long

Private Declare Function PTSHA1FilesProgInputW Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByRef szfids As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As HASH_RESULT

Private Declare Function PTSHA1FilesProgW Lib "PTHash" ( _
    ByRef szfids As Byte, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As Long

Private Declare Function PTSHA1Hmac Lib "PTHash" ( _
    ByVal Text As String, _
    ByVal text_len As Long, _
    ByVal Key As String, _
    ByVal key_len As Long, _
    ByVal digest As String) As Long


''''''''''''''''''''''''''''''
'''''''''  SHA224  '''''''''''
Private Declare Function PTSHA224BlockLen Lib "PTHash" () As Long

Private Declare Function PTSHA224SigLen Lib "PTHash" () As Long

Private Declare Function PTSHA224CtxLen Lib "PTHash" () As Long

Private Declare Function PTSHA224Reset Lib "PTHash" ( _
    ByRef ctx As Byte) As HASH_RESULT

Private Declare Function PTSHA224Input Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByRef buf As Byte, _
    ByVal buflen As Long) As HASH_RESULT

Private Declare Function PTSHA224Result Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByRef sig As Byte) As HASH_RESULT

Private Declare Function PTSHA224String Lib "PTHash" ( _
    ByRef buf As Byte, _
    ByVal buflen As Long, _
    ByRef sig As Byte) As Long

Private Declare Function PTSHA224FileA Lib "PTHash" ( _
    ByVal szfid As String, _
    ByRef sig As Byte) As Long

Private Declare Function PTSHA224FileProgA Lib "PTHash" ( _
    ByVal szfid As String, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As Long

Private Declare Function PTSHA224FileProgW Lib "PTHash" ( _
    ByRef szfid As Byte, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As Long

Private Declare Function PTSHA224FilesA Lib "PTHash" ( _
    ByVal szfids As String, _
    ByRef sig As Byte) As Long

Private Declare Function PTSHA224FilesProgInputA Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByVal szfids As String, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As HASH_RESULT

Private Declare Function PTSHA224FilesProgA Lib "PTHash" ( _
    ByVal szfids As String, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As Long

Private Declare Function PTSHA224FilesProgInputW Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByRef szfids As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As HASH_RESULT

Private Declare Function PTSHA224FilesProgW Lib "PTHash" ( _
    ByRef szfids As Byte, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As Long

Private Declare Function PTSHA224Hmac Lib "PTHash" ( _
    ByVal Text As String, _
    ByVal text_len As Long, _
    ByVal Key As String, _
    ByVal key_len As Long, _
    ByVal digest As String) As Long


''''''''''''''''''''''''''''''
'''''''''  SHA256  '''''''''''
Private Declare Function PTSHA256BlockLen Lib "PTHash" () As Long

Private Declare Function PTSHA256SigLen Lib "PTHash" () As Long

Private Declare Function PTSHA256CtxLen Lib "PTHash" () As Long

Private Declare Function PTSHA256Reset Lib "PTHash" ( _
    ByRef ctx As Byte) As HASH_RESULT

Private Declare Function PTSHA256Input Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByRef buf As Byte, _
    ByVal buflen As Long) As HASH_RESULT

Private Declare Function PTSHA256Result Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByRef sig As Byte) As HASH_RESULT

Private Declare Function PTSHA256String Lib "PTHash" ( _
    ByRef buf As Byte, _
    ByVal buflen As Long, _
    ByRef sig As Byte) As Long

Private Declare Function PTSHA256FileA Lib "PTHash" ( _
    ByVal szfid As String, _
    ByRef sig As Byte) As Long

Private Declare Function PTSHA256FileProgA Lib "PTHash" ( _
    ByVal szfid As String, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As Long

Private Declare Function PTSHA256FileProgW Lib "PTHash" ( _
    ByRef szfid As Byte, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As Long

Private Declare Function PTSHA256FilesA Lib "PTHash" ( _
    ByVal szfids As String, _
    ByRef sig As Byte) As Long

Private Declare Function PTSHA256FilesProgInputA Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByVal szfids As String, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As HASH_RESULT

Private Declare Function PTSHA256FilesProgA Lib "PTHash" ( _
    ByVal szfids As String, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As Long

Private Declare Function PTSHA256FilesProgInputW Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByRef szfids As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As HASH_RESULT

Private Declare Function PTSHA256FilesProgW Lib "PTHash" ( _
    ByRef szfids As Byte, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As Long

Private Declare Function PTSHA256Hmac Lib "PTHash" ( _
    ByVal Text As String, _
    ByVal text_len As Long, _
    ByVal Key As String, _
    ByVal key_len As Long, _
    ByVal digest As String) As Long


''''''''''''''''''''''''''''''
'''''''''  SHA384  '''''''''''
Private Declare Function PTSHA384BlockLen Lib "PTHash" () As Long

Private Declare Function PTSHA384SigLen Lib "PTHash" () As Long

Private Declare Function PTSHA384CtxLen Lib "PTHash" () As Long

Private Declare Function PTSHA384Reset Lib "PTHash" ( _
    ByRef ctx As Byte) As HASH_RESULT

Private Declare Function PTSHA384Input Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByRef buf As Byte, _
    ByVal buflen As Long) As HASH_RESULT

Private Declare Function PTSHA384Result Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByRef sig As Byte) As HASH_RESULT

Private Declare Function PTSHA384String Lib "PTHash" ( _
    ByRef buf As Byte, _
    ByVal buflen As Long, _
    ByRef sig As Byte) As Long

Private Declare Function PTSHA384FileA Lib "PTHash" ( _
    ByVal szfid As String, _
    ByRef sig As Byte) As Long

Private Declare Function PTSHA384FileProgA Lib "PTHash" ( _
    ByVal szfid As String, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As Long

Private Declare Function PTSHA384FileProgW Lib "PTHash" ( _
    ByRef szfid As Byte, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As Long

Private Declare Function PTSHA384FilesA Lib "PTHash" ( _
    ByVal szfids As String, _
    ByRef sig As Byte) As Long

Private Declare Function PTSHA384FilesProgInputA Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByVal szfids As String, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As HASH_RESULT

Private Declare Function PTSHA384FilesProgA Lib "PTHash" ( _
    ByVal szfids As String, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As Long

Private Declare Function PTSHA384FilesProgInputW Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByRef szfids As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As HASH_RESULT

Private Declare Function PTSHA384FilesProgW Lib "PTHash" ( _
    ByRef szfids As Byte, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As Long

Private Declare Function PTSHA384Hmac Lib "PTHash" ( _
    ByVal Text As String, _
    ByVal text_len As Long, _
    ByVal Key As String, _
    ByVal key_len As Long, _
    ByVal digest As String) As Long


''''''''''''''''''''''''''''''
'''''''''  SHA512  '''''''''''
Private Declare Function PTSHA512BlockLen Lib "PTHash" () As Long

Private Declare Function PTSHA512SigLen Lib "PTHash" () As Long

Private Declare Function PTSHA512CtxLen Lib "PTHash" () As Long

Private Declare Function PTSHA512Reset Lib "PTHash" ( _
    ByRef ctx As Byte) As HASH_RESULT

Private Declare Function PTSHA512Input Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByRef buf As Byte, _
    ByVal buflen As Long) As HASH_RESULT

Private Declare Function PTSHA512Result Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByRef sig As Byte) As HASH_RESULT

Private Declare Function PTSHA512String Lib "PTHash" ( _
    ByRef buf As Byte, _
    ByVal buflen As Long, _
    ByRef sig As Byte) As Long

Private Declare Function PTSHA512FileA Lib "PTHash" ( _
    ByVal szfid As String, _
    ByRef sig As Byte) As Long

Private Declare Function PTSHA512FileProgA Lib "PTHash" ( _
    ByVal szfid As String, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As Long

Private Declare Function PTSHA512FileProgW Lib "PTHash" ( _
    ByRef szfid As Byte, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As Long

Private Declare Function PTSHA512FilesA Lib "PTHash" ( _
    ByVal szfids As String, _
    ByRef sig As Byte) As Long

Private Declare Function PTSHA512FilesProgInputA Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByVal szfids As String, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As HASH_RESULT

Private Declare Function PTSHA512FilesProgA Lib "PTHash" ( _
    ByVal szfids As String, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As Long

Private Declare Function PTSHA512FilesProgInputW Lib "PTHash" ( _
    ByRef ctx As Byte, _
    ByRef szfids As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As HASH_RESULT

Private Declare Function PTSHA512FilesProgW Lib "PTHash" ( _
    ByRef szfids As Byte, _
    ByRef sig As Byte, _
    ByVal fpCallBack As Long, _
    ByVal Freq As Long) As Long

Private Declare Function PTSHA512Hmac Lib "PTHash" ( _
    ByVal Text As String, _
    ByVal text_len As Long, _
    ByVal Key As String, _
    ByVal key_len As Long, _
    ByVal digest As String) As Long


Private Declare Function TestA Lib "PTHash" (ByVal szfid As String) As Long
Private Declare Function TestW Lib "PTHash" (ByRef szfid As Byte) As Long


Private InitDone As Boolean
Private HashOSInfo As PT_OS_INFO

Private mIOBlock As Long

Private mCtxBound(HASH_TYPE_UBOUND) As Long
Private mSigBound(HASH_TYPE_UBOUND) As Long
Private mSigTextLen(HASH_TYPE_UBOUND) As Long

Private MD5Ctx() As Byte
Private SHA1Ctx() As Byte
Private SHA224Ctx() As Byte
Private SHA256Ctx() As Byte
Private SHA384Ctx() As Byte
Private SHA512Ctx() As Byte

Private mPresent As Boolean
Private mBanner As String
Private mResult As HASH_RESULT
Private mCancel As Boolean

Public HashFileTot As Long
Public HashFileCnt As Long

Public HashPanel As Panel
Public Property Get HashResult() As HASH_RESULT
    HashResult = mResult
End Property
Public Property Get Banner() As String
    If Not InitDone Then Intialize
    Banner = mBanner
End Property
Public Property Get Present() As Boolean
    If Not InitDone Then Intialize
    Present = mPresent
End Property
Public Property Get Cancel() As Boolean
    Cancel = mCancel
End Property
Public Property Let Cancel(ByVal NewValue As Boolean)
    mCancel = NewValue
End Property
Public Sub HashDemo()
    Dim sig() As Byte
    Dim Hash As String
    Dim fn As Long
    Dim Fid As String
    Dim Vector As String
    Dim i As Long

    Vector = String(1000000, Chr(255))
    Vector = String(999999, Chr(255))

    ReDim sig(mSigBound(HASH_SHA512))

    Fid = App.Path & "\HashDemo1.txt"
    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Vector = String(1000000, Chr(255))
        For i = 1 To 1000
            Print #fn, Vector;
            DoEvents
        Next i
        Close fn
        sig = HashFile(HASH_MD5, Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    Debug.Print "8,000,000,000 1s: " & Hash

    Fid = App.Path & "\HashDemo2.txt"
    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Vector = String(1000000, Chr(255))
        For i = 1 To 999
            Print #fn, Vector;
            DoEvents
        Next i
        Print #fn, String(999999, Chr(255));
        Print #fn, Chr(254);
        Close fn
        sig = HashFile(HASH_MD5, Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    Debug.Print "7,999,999,999 1s and 1 0: " & Hash
End Sub
Public Function HashTest(lst As ListBox) As Boolean
    Dim sig() As Byte
    Dim Hash As String
    Dim fn As Long
    Dim Vector As String
    Dim Vector2 As String
    Dim result As String
    Dim Fid As String
    Dim Fid2 As String
    Dim Unicode() As Byte
    Dim TestCnt As Long
    Dim SuccessCnt As Long
    Dim FailCnt As Long
    Dim i As Long

    On Error Resume Next
    Fid = App.Path & "\HashTest.txt"
    Fid2 = App.Path & "\HashTest2.txt"

    ReDim sig(mSigBound(HASH_MD5))
    lst.AddItem "MD5 Tests Begin ..."
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    i = PTMD5CtxLen()
    lst.AddItem "MD5Context Len: " & CStr(i)

    Vector = ""
    result = "d41d8cd98f00b204e9800998ecf8427e"
    sig = HashString(HASH_MD5, Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "MD5 Test 1a Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "MD5 Test 1a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = HashFile(HASH_MD5, Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "MD5 Test 1b Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "MD5 Test 1b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = "a"
    result = "0cc175b9c0f1b6a831c399e269772661"
    sig = HashString(HASH_MD5, Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "MD5 Test 2a Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "MD5 Test 2a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = HashFile(HASH_MD5, Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "MD5 Test 2b Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "MD5 Test 2b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = "abc"
    result = "900150983cd24fb0d6963f7d28e17f72"
    sig = HashString(HASH_MD5, Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "MD5 Test 3a Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "MD5 Test 3a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = HashFile(HASH_MD5, Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "MD5 Test 3b Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "MD5 Test 3b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = "message digest"
    result = "f96b697d7cb7938d525a2f31aaf161d0"
    sig = HashString(HASH_MD5, Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "MD5 Test 4a Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "MD5 Test 4a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = HashFile(HASH_MD5, Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "MD5 Test 4b Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "MD5 Test 4b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = "abcdefghijklmnopqrstuvwxyz"
    result = "c3fcd3d76192e4007dfb496cca67e13b"
    sig = HashString(HASH_MD5, Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "MD5 Test 5a Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "MD5 Test 5a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = HashFile(HASH_MD5, Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "MD5 Test 5b Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "MD5 Test 5b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    result = "d174ab98d277d9f5a5611c2c9f419d9f"
    sig = HashString(HASH_MD5, Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "MD5 Test 6a Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "MD5 Test 6a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = HashFile(HASH_MD5, Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "MD5 Test 6b Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "MD5 Test 6b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = "12345678901234567890123456789012345678901234567890123456789012345678901234567890"
    result = "57edf4a22be3c955ac49da2e2107b67a"
    sig = HashString(HASH_MD5, Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "MD5 Test 7a Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "MD5 Test 7a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = HashFile(HASH_MD5, Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "MD5 Test 7b Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "MD5 Test 7b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = "1234567890123456789012345678901234567890"
    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = HashFiles(HASH_MD5, Fid & vbNullChar & Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "MD5 Test 7c Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "MD5 Test 7c Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If

    Vector = "1234567890123456789012345678901234567890"
    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        HashInit HASH_MD5
        HashFilesInput HASH_MD5, Fid & vbNullChar & Fid
        sig = HashFini(HASH_MD5)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "MD5 Test 7d Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "MD5 Test 7d Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    DoEvents

    lst.AddItem "MD5 Tests End ..."
    lst.TopIndex = lst.ListCount - 1
    DoEvents

' HMAC-MD5
'  key =         0x0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b
'  key_len =     16 bytes
'  Data = "Hi There"
'  data_len =    8  bytes
'  digest =      0x9294727a3638bb1c13f48ef8158bfc9d
'
'  Key = "Jefe"
'  Data = "what do ya want for nothing?"
'  data_len =    28 bytes
'  digest =      0x750c783e6ab0b503eaa86e310a5db738
'
'  key =         0xAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
'
'  key_len       16 bytes
'  data =        0xDDDDDDDDDDDDDDDDDDDD...
'                ..DDDDDDDDDDDDDDDDDDDD...
'                ..DDDDDDDDDDDDDDDDDDDD...
'                ..DDDDDDDDDDDDDDDDDDDD...
'                ..DDDDDDDDDDDDDDDDDDDD
'  data_len =    50 bytes
'  digest =      0x56be34521d144c88dbb8c733f0e8b3f6



    ReDim sig(mSigBound(HASH_SHA1))
    lst.AddItem ""
    lst.AddItem "SHA-1 Tests Begin ..."
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    i = PTSHA1CtxLen()
    lst.AddItem "SHA1Context Len: " & CStr(i)

    Vector = "abc"
    result = "A9993E364706816ABA3E25717850C26C9CD0D89D"
    sig = HashString(HASH_SHA1, Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-1 Test 1a Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-1 Test 1a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = HashFile(HASH_SHA1, Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-1 Test 1b Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-1 Test 1b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq"
    result = "84983E441C3BD26EBAAE4AA1F95129E5E54670F1"
    sig = HashString(HASH_SHA1, Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-1 Test 2a Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-1 Test 2a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = HashFile(HASH_SHA1, Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-1 Test 2b Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-1 Test 2b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = String(1000000, "a")
    result = "34AA973CD4C4DAA4F61EEB2BDBAD27316534016F"
    sig = HashString(HASH_SHA1, Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-1 Test 3a Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-1 Test 3a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = HashFile(HASH_SHA1, Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-1 Test 3b Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-1 Test 3b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = String(500000, "a")
    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = HashFiles(HASH_SHA1, Fid & vbNullChar & Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-1 Test 3c Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-1 Test 3c Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = "0123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567"
    result = "DEA356A2CDDD90C7A7ECEDC5EBB563934F460452"
    sig = HashString(HASH_SHA1, Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-1 Test 4a Failed: " & Hash
    Else
        lst.AddItem "SHA-1 Test 4a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = HashFile(HASH_SHA1, Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-1 Test 4b Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-1 Test 4b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    DoEvents

    Vector = "0123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567"
    Vector2 = "012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567012345670123456701234567"
    result = "DEA356A2CDDD90C7A7ECEDC5EBB563934F460452"
    Hash = ""
    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        fn = FreeFile()
        Open Fid2 For Output Access Write As fn
        If Err.Number = 0 Then
            Print #fn, Vector2;
            Close #fn
            HashInit HASH_SHA1
            HashFilesInput HASH_SHA1, Fid & vbNullChar & Fid2
            sig = HashFini(HASH_SHA1)
            Hash = HashSig2Text(sig)
        Else
            Hash = ""
        End If
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-1 Test 4c Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-1 Test 4c Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    DoEvents

    lst.AddItem "SHA-1 Tests End ..."
    lst.TopIndex = lst.ListCount - 1
    DoEvents



    ReDim sig(mSigBound(HASH_SHA224))
    lst.AddItem ""
    lst.AddItem "SHA-224 Tests Begin ..."
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    i = PTSHA224CtxLen()
    lst.AddItem "SHA224Context Len: " & CStr(i)

    Vector = "abc"
    result = "23097d223405d8228642a477bda255b32aadbce4bda0b3f7e36c9da7"
    sig = HashString(HASH_SHA224, Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-224 Test 1a Failed: " & Hash
    Else
        lst.AddItem "SHA-224 Test 1a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = HashFile(HASH_SHA224, Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-224 Test 1b Failed: " & Hash
    Else
        lst.AddItem "SHA-224 Test 1b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq"
    result = "75388b16512776cc5dba5da1fd890150b0c6455cb4f58b1952522525"
    sig = HashString(HASH_SHA224, Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-224 Test 2a Failed: " & Hash
    Else
        lst.AddItem "SHA-224 Test 2a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = HashFile(HASH_SHA224, Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-224 Test 2b Failed: " & Hash
    Else
        lst.AddItem "SHA-224 Test 2b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    HashInit HASH_SHA224
    HashFilesInput HASH_SHA224, Fid
    sig = HashFini(HASH_SHA224)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-224 Test 2c Failed: " & Hash
    Else
        lst.AddItem "SHA-224 Test 2c Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = String(1000000, "a")
    result = "20794655980c91d8bbb4c1ea97618a4bf03f42581948b2ee4ee7ad67"
    sig = HashString(HASH_SHA224, Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-224 Test 3a Failed: " & Hash
    Else
        lst.AddItem "SHA-224 Test 3a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = HashFile(HASH_SHA224, Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-224 Test 3b Failed: " & Hash
    Else
        lst.AddItem "SHA-224 Test 3b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    HashInit HASH_SHA224
    HashFilesInput HASH_SHA224, Fid
    sig = HashFini(HASH_SHA224)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-224 Test 3c Failed: " & Hash
    Else
        lst.AddItem "SHA-224 Test 3c Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = String(500000, "a")
    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = HashFiles(HASH_SHA224, Fid & vbNullChar & Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-224 Test 3d Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-224 Test 3d Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.AddItem "SHA-224 Tests End ..."
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    HashInit HASH_SHA224
    HashFilesInput HASH_SHA224, Fid & vbNullChar & Fid
    sig = HashFini(HASH_SHA224)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-224 Test 3e Failed: " & Hash
    Else
        lst.AddItem "SHA-224 Test 3e Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    ReDim sig(mSigBound(HASH_SHA256))
    lst.AddItem ""
    lst.AddItem "SHA-256 Tests Begin ..."
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    i = PTSHA256CtxLen()
    lst.AddItem "SHA256Context Len: " & CStr(i)

    Vector = "abc"
    result = "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
    sig = SHA256String(Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-256 Test 1a Failed: " & Hash
    Else
        lst.AddItem "SHA-256 Test 1a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = SHA256File(Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-256 Test 1b Failed: " & Hash
    Else
        lst.AddItem "SHA-256 Test 1b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = "abcdbcdecdefdefgefghfghighijhijkijkljklmklmnlmnomnopnopq"
    result = "248d6a61d20638b8e5c026930c3e6039a33ce45964ff2167f6ecedd419db06c1"
    sig = SHA256String(Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-256 Test 2a Failed: " & Hash
    Else
        lst.AddItem "SHA-256 Test 2a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = SHA256File(Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-256 Test 2b Failed: " & Hash
    Else
        lst.AddItem "SHA-256 Test 2b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = String(1000000, "a")
    result = "cdc76e5c9914fb9281a1c7e284d73e67f1809a48a497200e046d39ccc7112cd0"
    sig = SHA256String(Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-256 Test 3a Failed: " & Hash
    Else
        lst.AddItem "SHA-256 Test 3a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = SHA256File(Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-256 Test 3b Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-256 Test 3b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = String(500000, "a")
    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = SHA256Files(Fid & vbNullChar & Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-256 Test 3c Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-256 Test 3c Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    DoEvents

    HashInit HASH_SHA256
    HashFilesInput HASH_SHA256, Fid & vbNullChar & Fid
    sig = HashFini(HASH_SHA256)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-256 Test 3d Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-256 Test 3d Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.AddItem "SHA-256 Tests End ..."
    lst.TopIndex = lst.ListCount - 1
    DoEvents


    ReDim sig(mSigBound(HASH_SHA384))
    lst.AddItem ""
    lst.AddItem "SHA-384 Tests Begin ..."
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    i = PTSHA384CtxLen()
    lst.AddItem "SHA384Context Len: " & CStr(i)

    Vector = "abc"
    result = "cb00753f45a35e8bb5a03d699ac65007272c32ab0eded1631a8b605a43ff5bed8086072ba1e7cc2358baeca134c825a7"
    sig = HashString(HASH_SHA384, Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-384 Test 1a Failed: " & Hash
    Else
        lst.AddItem "SHA-384 Test 1a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = SHA384File(Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-384 Test 1b Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-384 Test 1b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = "abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstu"
    result = "09330c33f71147e83d192fc782cd1b4753111b173b3b05d22fa08086e3b0f712fcc7c71a557e2db966c3e9fa91746039"
    sig = SHA384String(Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-384 Test 2a Failed: " & Hash
    Else
        lst.AddItem "SHA-384 Test 2a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = SHA384File(Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-384 Test 2b Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-384 Test 2b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = String(1000000, "a")
    result = "9d0e1809716474cb086e834e310a4a1ced149e9c00f248527972cec5704c2a5b07b8b3dc38ecc4ebae97ddd87f3d8985"
    sig = SHA384String(Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-384 Test 3a Failed: " & Hash
    Else
        lst.AddItem "SHA-384 Test 3a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = SHA384File(Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-384 Test 3b Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-384 Test 3b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = String(500000, "a")
    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = SHA384Files(Fid & vbNullChar & Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-384 Test 3c Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-384 Test 3c Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    DoEvents

    HashInit HASH_SHA384
    HashFilesInput HASH_SHA384, Fid & vbNullChar & Fid
    sig = HashFini(HASH_SHA384)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-384 Test 3d Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-384 Test 3d Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If

    lst.AddItem "SHA-384 Tests End ..."
    lst.TopIndex = lst.ListCount - 1
    DoEvents


    ReDim sig(mSigBound(HASH_SHA512))
    lst.AddItem ""
    lst.AddItem "SHA-512 Tests Begin ..."
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    i = PTSHA512CtxLen()
    lst.AddItem "SHA512Context Len: " & CStr(i)

    Vector = "abc"
    result = "ddaf35a193617abacc417349ae20413112e6fa4e89a97ea20a9eeee64b55d39a2192992a274fc1a836ba3c23a3feebbd454d4423643ce80e2a9ac94fa54ca49f"
    sig = SHA512String(Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-512 Test 1a Failed: " & Hash
    Else
        lst.AddItem "SHA-512 Test 1a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = SHA512File(Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-512 Test 1b Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-512 Test 1b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = "abcdefghbcdefghicdefghijdefghijkefghijklfghijklmghijklmnhijklmnoijklmnopjklmnopqklmnopqrlmnopqrsmnopqrstnopqrstu"
    result = "8e959b75dae313da8cf4f72814fc143f8f7779c6eb9f7fa17299aeadb6889018501d289e4900f7e4331b99dec4b5433ac7d329eeb6dd26545e96e55b874be909"
    sig = SHA512String(Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-512 Test 2a Failed: " & Hash
    Else
        lst.AddItem "SHA-512 Test 2a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = SHA512File(Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-512 Test 2b Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-512 Test 2b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = String(1000000, "a")
    result = "e718483d0ce769644e2e42c7bc15b4638e1f98b13b2044285632a803afa973ebde0ff244877ea60a4cb0432ce577c31beb009c5c2c49aa2e4eadb217ad8cc09b"
    sig = SHA512String(Vector)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-512 Test 3a Failed: " & Hash
    Else
        lst.AddItem "SHA-512 Test 3a Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = SHA512File(Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-512 Test 3b Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-512 Test 3b Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Vector = String(500000, "a")
    fn = FreeFile()
    Open Fid For Output Access Write As fn
    If Err.Number = 0 Then
        Print #fn, Vector;
        Close fn
        sig = SHA512Files(Fid & vbNullChar & Fid)
        Hash = HashSig2Text(sig)
    Else
        Hash = ""
    End If
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-512 Test 3c Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-512 Test 3c Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    DoEvents

    HashInit HASH_SHA512
    HashFilesInput HASH_SHA512, Fid & vbNullChar & Fid
    sig = HashFini(HASH_SHA512)
    Hash = HashSig2Text(sig)
    TestCnt = TestCnt + 1
    If StrComp(Hash, result, vbTextCompare) <> 0 Then
        lst.AddItem "SHA-512 Test 3d Failed"
        FailCnt = FailCnt + 1
    Else
        lst.AddItem "SHA-512 Test 3d Succeeded"
        SuccessCnt = SuccessCnt + 1
    End If
    DoEvents

    lst.AddItem "SHA-512 Tests End ..."
    lst.TopIndex = lst.ListCount - 1
    DoEvents

    Kill Fid

    lst.AddItem ""
    If (TestCnt = SuccessCnt) And (FailCnt = 0) Then
        HashTest = True
        lst.AddItem "All " & FormatNumber(TestCnt, 0) & " Tests Completed Successfully"
    Else
        HashTest = False
        lst.AddItem FormatNumber(FailCnt, 0) & " of " & FormatNumber(TestCnt, 0) & " Tests Failed"
    End If
    lst.TopIndex = lst.ListCount - 1
End Function
Public Sub HashTestHmac()
'RFC 3174 - US Secure Hash Algorithm 1 (SHA1)
'2. Test Cases for HMAC-MD5
'
'test_case = 1
'key =           0x0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b
'key_len = 16
'Data = "Hi There"
'data_len = 8
'digest =        0x9294727a3638bb1c13f48ef8158bfc9d
'
'test_case = 2
'key = "Jefe"
'key_len = 4
'Data = "what do ya want for nothing?"
'data_len = 28
'digest =        0x750c783e6ab0b503eaa86e310a5db738
'
'test_case = 3
'key =           0xaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa
'key_len 16
'data =          0xdd repeated 50 times
'data_len = 50
'digest =        0x56be34521d144c88dbb8c733f0e8b3f6
'
'test_case = 4
'key =           0x0102030405060708090a0b0c0d0e0f10111213141516171819
'key_len 25
'data =          0xcd repeated 50 times
'data_len = 50
'digest =        0x697eaf0aca3a3aea3a75164746ffaa79
'
'
'test_case = 5
'key =           0x0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c
'key_len = 16
'Data = "Test With Truncation"
'data_len = 20
'digest =        0x56461ef2342edc00f9bab995690efd4c
'digest-96       0x56461ef2342edc00f9bab995
'
'test_case = 6
'key =           0xaa repeated 80 times
'key_len = 80
'Data = "Test Using Larger Than Block-Size Key - Hash Key First"
'data_len = 54
'digest =        0x6b1ab7fe4bd7bf8f0b62e6ce61b9d0cd
'
'test_case = 7
'key =           0xaa repeated 80 times
'key_len = 80
'Data = "Test Using Larger Than Block-Size Key and Larger Than One Block-Size Data"
'data_len = 73
'digest =        0x6f630fad67cda0ee1fb1f562db3aa53e
'
'3. Test Cases for HMAC-SHA-1
'
'test_case = 1
'key =           0x0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b0b
'key_len = 20
'Data = "Hi There"
'data_len = 8
'digest =        0xb617318655057264e28bc0b6fb378c8ef146be00
'
'test_case = 2
'key = "Jefe"
'key_len = 4
'Data = "what do ya want for nothing?"
'data_len = 28
'digest =        0xeffcdf6ae5eb2fa2d27416d5f184df9c259a7c79
'
'test_case = 3
'key =           0xaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa
'key_len = 20
'data =          0xdd repeated 50 times
'data_len = 50
'digest =        0x125d7342b9ac11cd91a39af48aa17b4f63f175d3
'
'
'test_case = 4
'key =           0x0102030405060708090a0b0c0d0e0f10111213141516171819
'key_len = 25
'data =          0xcd repeated 50 times
'data_len = 50
'digest =        0x4c9007f4026250c6bc8414f9bf50c86c2d7235da
'
'test_case = 5
'key =           0x0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c0c
'key_len = 20
'Data = "Test With Truncation"
'data_len = 20
'digest =        0x4c1a03424b55e07fe7f27be1d58bb9324a9a5a04
'digest-96 =     0x4c1a03424b55e07fe7f27be1
'
'test_case = 6
'key =           0xaa repeated 80 times
'key_len = 80
'Data = "Test Using Larger Than Block-Size Key - Hash Key First"
'data_len = 54
'digest =        0xaa4ae5e15272d00e95705637ce8a3b55ed402112
'
'test_case = 7
'key =           0xaa repeated 80 times
'key_len = 80
'Data = "Test Using Larger Than Block-Size Key and Larger Than One Block-Size Data"
'data_len = 73
'digest =        0xe8e99d0f45237d786d6bbaa7965c7808bbff1a91
'data_len = 20
'digest =        0x4c1a03424b55e07fe7f27be1d58bb9324a9a5a04
'digest-96 =     0x4c1a03424b55e07fe7f27be1
'
'test_case = 6
'key =           0xaa repeated 80 times
'key_len = 80
'Data = "Test Using Larger Than Block-Size Key - Hash Key First"
'data_len = 54
'digest =        0xaa4ae5e15272d00e95705637ce8a3b55ed402112
'
'test_case = 7
'key =           0xaa repeated 80 times
'key_len = 80
'Data = "Test Using Larger Than Block-Size Key and Larger Than One Block-Size Data"
'data_len = 73
'digest =        0xe8e99d0f45237d786d6bbaa7965c7808bbff1a91
End Sub
Private Sub Intialize()
    Dim buf As String
    Dim buflen As Long
    Dim result As Long

    HashOSInfo = ApiWinVersion()
    On Error Resume Next
    buflen = PTHashDescA(buf, 0)
    If Err.Number <> 0 Then
        mBanner = "File PTHash.dll is either missing, or not in correct directory." & vbCrLf & "Reinstall this program to correct problem."
        mPresent = False
        Exit Sub
    End If

    mPresent = True

    buf = String(buflen, 0)
    result = PTHashDescA(buf, Len(buf))
    If result > 0 Then
        mBanner = ApiTextStrip(buf)
    Else
        mBanner = "Error Retrieving PTHash.dll Version Info"
    End If

    mCtxBound(HASH_MD5) = PTMD5CtxLen() - 1
    mCtxBound(HASH_SHA1) = PTSHA1CtxLen() - 1
    mCtxBound(HASH_SHA224) = PTSHA224CtxLen() - 1
    mCtxBound(HASH_SHA256) = PTSHA256CtxLen() - 1
    mCtxBound(HASH_SHA384) = PTSHA384CtxLen() - 1
    mCtxBound(HASH_SHA512) = PTSHA512CtxLen() - 1

    mSigBound(HASH_MD5) = PTMD5SigLen() - 1
    mSigBound(HASH_SHA1) = PTSHA1SigLen() - 1
    mSigBound(HASH_SHA224) = PTSHA224SigLen() - 1
    mSigBound(HASH_SHA256) = PTSHA256SigLen() - 1
    mSigBound(HASH_SHA384) = PTSHA384SigLen() - 1
    mSigBound(HASH_SHA512) = PTSHA512SigLen() - 1

    mSigTextLen(HASH_MD5) = PTMD5SigLen() * 2
    mSigTextLen(HASH_SHA1) = PTSHA1SigLen() * 2
    mSigTextLen(HASH_SHA224) = PTSHA224SigLen() * 2
    mSigTextLen(HASH_SHA256) = PTSHA256SigLen() * 2
    mSigTextLen(HASH_SHA384) = PTSHA384SigLen() * 2
    mSigTextLen(HASH_SHA512) = PTSHA512SigLen() * 2

    mIOBlock = PTHashIOBlock()

    mCancel = False
    mResult = HASH_OK
    InitDone = True
End Sub
Public Sub HashCallBack(ByVal FileCnt As Long, ByVal FileTot As Long, ByVal BlockCnt As Long, ByVal BlockTot As Long, ByRef Cancel As Long)
    Dim s As String

    If HashPanel Is Nothing Then
        DoEvents
        Cancel = mCancel
        Exit Sub
    End If

    If HashFileTot > 0 Then
        If (BlockCnt = 0) Then HashFileCnt = HashFileCnt + 1
        If (HashFileCnt = HashFileTot) And (BlockCnt = BlockTot) Then
            HashPanel.Text = "" ' "Calculations Finished"
            HashFileTot = 0
            HashFileCnt = 0
        Else
            s = "Processing File " & FormatNumber(HashFileCnt, 0) & " of " & FormatNumber(HashFileTot, 0) & ": "
            If BlockTot > 0 Then
                HashPanel.Text = s & FormatPercent(BlockCnt / BlockTot, 2)
            Else
                HashPanel.Text = s & FormatPercent(1, 0)
            End If
        End If
    Else
        If (FileCnt = FileTot) And (BlockCnt = BlockTot) Then
            HashPanel.Text = "" ' "Calculations Finished"
            HashFileTot = 0
            HashFileCnt = 0
        Else
            s = "Processing File " & FormatNumber(FileCnt, 0) & " of " & FormatNumber(FileTot, 0) & ": "
            If BlockTot > 0 Then
                HashPanel.Text = s & FormatPercent(BlockCnt / BlockTot, 2)
            Else
                HashPanel.Text = s & FormatPercent(1, 0)
            End If
        End If
    End If
    DoEvents

    Cancel = mCancel
    Err.Clear
End Sub
Public Function HashInit(HashType As HASH_TYPE) As HASH_RESULT
    Dim hr As HASH_RESULT

    Select Case HashType
        Case HASH_MD5
            ReDim MD5Ctx(mCtxBound(HASH_MD5))
            hr = PTMD5Init(MD5Ctx(0))
        Case HASH_SHA1
            ReDim SHA1Ctx(mCtxBound(HASH_SHA1))
            hr = PTSHA1Reset(SHA1Ctx(0))
        Case HASH_SHA224
            ReDim SHA224Ctx(mCtxBound(HASH_SHA224))
            hr = PTSHA224Reset(SHA224Ctx(0))
        Case HASH_SHA256
            ReDim SHA256Ctx(mCtxBound(HASH_SHA256))
            hr = PTSHA256Reset(SHA256Ctx(0))
        Case HASH_SHA384
            ReDim SHA384Ctx(mCtxBound(HASH_SHA384))
            hr = PTSHA384Reset(SHA384Ctx(0))
        Case HASH_SHA512
            ReDim SHA512Ctx(mCtxBound(HASH_SHA512))
            hr = PTSHA512Reset(SHA512Ctx(0))
        Case Else
            hr = HASH_BADTYPE
    End Select

    HashInit = hr
End Function
Public Function HashFilesInput(HashType As HASH_TYPE, FileNames As Variant) As HASH_RESULT
    mResult = HASH_OK
    Select Case HashType
        Case HASH_MD5:      mResult = MD5FilesInput(FileNames)
        Case HASH_SHA1:     mResult = SHA1FilesInput(FileNames)
        Case HASH_SHA224:   mResult = SHA224FilesInput(FileNames)
        Case HASH_SHA256:   mResult = SHA256FilesInput(FileNames)
        Case HASH_SHA384:   mResult = SHA384FilesInput(FileNames)
        Case HASH_SHA512:   mResult = SHA512FilesInput(FileNames)
        Case Else:          mResult = HASH_BADTYPE
    End Select

    HashFilesInput = mResult
End Function
Public Function HashFini(HashType As HASH_TYPE) As Byte()
    Dim sig() As Byte
    Dim ResultHold

    ResultHold = mResult
    mResult = HASH_OK
    Select Case HashType
        Case HASH_MD5
            ReDim sig(mSigBound(HashType))
            mResult = PTMD5Fini(MD5Ctx(0), sig(0))
        Case HASH_SHA1
            ReDim sig(mSigBound(HashType))
            mResult = PTSHA1Result(SHA1Ctx(0), sig(0))
        Case HASH_SHA224
            ReDim sig(mSigBound(HashType))
            mResult = PTSHA224Result(SHA224Ctx(0), sig(0))
        Case HASH_SHA256
            ReDim sig(mSigBound(HashType))
            mResult = PTSHA256Result(SHA256Ctx(0), sig(0))
        Case HASH_SHA384
            ReDim sig(mSigBound(HashType))
            mResult = PTSHA384Result(SHA384Ctx(0), sig(0))
        Case HASH_SHA512
            ReDim sig(mSigBound(HashType))
            mResult = PTSHA512Result(SHA512Ctx(0), sig(0))
        Case Else
            Erase sig
            mResult = HASH_BADTYPE
    End Select

    If mResult = HASH_OK Then mResult = ResultHold
    HashFini = sig
End Function
Public Function HashString(HashType As HASH_TYPE, PlainText As Variant) As Byte()
    Dim sig() As Byte

    mResult = HASH_OK
    Select Case HashType
        Case HASH_MD5:      sig = MD5String(PlainText)
        Case HASH_SHA1:     sig = SHA1String(PlainText)
        Case HASH_SHA224:   sig = SHA224String(PlainText)
        Case HASH_SHA256:   sig = SHA256String(PlainText)
        Case HASH_SHA384:   sig = SHA384String(PlainText)
        Case HASH_SHA512:   sig = SHA512String(PlainText)
        Case Else:          mResult = HASH_BADTYPE
    End Select

    HashString = sig
End Function
Public Function HashFile(HashType As HASH_TYPE, Filename As String) As Byte()
    Dim sig() As Byte

    mResult = HASH_OK
    Select Case HashType
        Case HASH_MD5:      sig = MD5File(Filename)
        Case HASH_SHA1:     sig = SHA1File(Filename)
        Case HASH_SHA224:   sig = SHA224File(Filename)
        Case HASH_SHA256:   sig = SHA256File(Filename)
        Case HASH_SHA384:   sig = SHA384File(Filename)
        Case HASH_SHA512:   sig = SHA512File(Filename)
        Case Else:          mResult = HASH_BADTYPE
    End Select

    HashFile = sig
End Function
Public Function HashFiles(HashType As HASH_TYPE, FileNames As Variant) As Byte()
    Dim sig() As Byte

    mResult = HASH_OK
    Select Case HashType
        Case HASH_MD5:      sig = MD5Files(FileNames)
        Case HASH_SHA1:     sig = SHA1Files(FileNames)
        Case HASH_SHA224:   sig = SHA224Files(FileNames)
        Case HASH_SHA256:   sig = SHA256Files(FileNames)
        Case HASH_SHA384:   sig = SHA384Files(FileNames)
        Case HASH_SHA512:   sig = SHA512Files(FileNames)
        Case Else:          mResult = HASH_BADTYPE
    End Select

    HashFiles = sig
End Function
Private Function MD5String(PlainText As Variant) As Byte()
    Dim buf() As Byte
    Dim buflen As Long
    Dim sig() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    If IsArray(PlainText) Then
        buf = PlainText
        buflen = UBound(buf) + 1
    ElseIf Len(PlainText) = 0 Then
        ReDim buf(0)
        buflen = 0
    Else
        buf = StrConv(PlainText, vbFromUnicode)
        buflen = UBound(buf) + 1
    End If

    ReDim sig(mSigBound(HASH_MD5))
    mResult = PTMD5String(buf(0), buflen, sig(0))
    MD5String = sig
End Function
'Public Function HashMD5StringCompare(PlainText As Variant, OrigSig() As Byte) As Boolean
'    Dim buf() As Byte
'    Dim buflen As Long
'    Dim sig() As Byte
'    Dim i As Long
'
'    If Not InitDone Then Intialize
'    mCancel = False
'
'    HashMD5StringCompare = False
'    If mSigBound(HASH_MD5) <> UBound(OrigSig) Then Exit Function
'
'    If IsArray(PlainText) Then
'        buf = PlainText
'        buflen = UBound(buf) + 1
'    ElseIf Len(PlainText) = 0 Then
'        ReDim buf(0)
'        buflen = 0
'    Else
'        buf = StrConv(PlainText, vbFromUnicode)
'        buflen = UBound(buf) + 1
'    End If
'
'    ReDim sig(mSigBound(HASH_MD5))
'    mResult = PTMD5String(buf(0), buflen, sig(0))
'
'    If mResult <> HASH_OK Then Exit Function
'
'    For i = 0 To UBound(sig)
'        If sig(i) <> OrigSig(i) Then Exit Function
'    Next i
'
'    HashMD5StringCompare = True
'End Function
Private Function MD5File(Filename As String) As Byte()
    Dim Fid As String
    Dim sig() As Byte
    Dim Unicode() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    Fid = Filename & vbNullChar
    ReDim sig(mSigBound(HASH_MD5))
    If HashOSInfo.Unicode Then
        Unicode = Fid
        mResult = PTMD5FileProgW(Unicode(0), sig(0), AddressOf HashCallBack, HASH_FREQ)
    Else
        mResult = PTMD5FileProgA(Fid, sig(0), AddressOf HashCallBack, HASH_FREQ)
    End If

    MD5File = sig
End Function
Private Function MD5Files(FileNames As Variant) As Byte()
    Dim fids As String
    Dim sig() As Byte
    Dim i As Long
    Dim Unicode() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    If IsArray(FileNames) Then
        For i = 0 To UBound(FileNames)
            fids = fids & FileNames(i) & vbNullChar
        Next i
    Else
        fids = FileNames
    End If
    fids = fids & vbNullChar & vbNullChar

    ReDim sig(mSigBound(HASH_MD5))
    If HashOSInfo.Unicode Then
        Unicode = fids
        mResult = PTMD5FilesProgW(Unicode(0), sig(0), AddressOf HashCallBack, HASH_FREQ)
    Else
        mResult = PTMD5FilesProgA(fids, sig(0), AddressOf HashCallBack, HASH_FREQ)
    End If
    MD5Files = sig
End Function
Private Function MD5FilesInput(FileNames As Variant) As HASH_RESULT
    Dim fids As String
    Dim sig() As Byte
    Dim i As Long
    Dim Unicode() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    If IsArray(FileNames) Then
        For i = 0 To UBound(FileNames)
            fids = fids & FileNames(i) & vbNullChar
        Next i
    Else
        fids = FileNames
    End If
    fids = fids & vbNullChar & vbNullChar

    ReDim sig(mSigBound(HASH_MD5))
    If HashOSInfo.Unicode Then
        Unicode = fids
        mResult = PTMD5FilesProgInputW(MD5Ctx(0), Unicode(0), AddressOf HashCallBack, HASH_FREQ)
    Else
        mResult = PTMD5FilesProgInputA(MD5Ctx(0), fids, AddressOf HashCallBack, HASH_FREQ)
    End If

    MD5FilesInput = mResult
End Function
Private Function SHA1String(PlainText As Variant) As Byte()
    Dim buf() As Byte
    Dim buflen As Long
    Dim sig() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    If IsArray(PlainText) Then
        buf = PlainText
        buflen = UBound(buf) + 1
    ElseIf Len(PlainText) = 0 Then
        ReDim buf(0)
        buflen = 0
    Else
        buf = StrConv(PlainText, vbFromUnicode)
        buflen = UBound(buf) + 1
    End If

    ReDim sig(mSigBound(HASH_SHA1))
    mResult = PTSHA1String(buf(0), buflen, sig(0))
    SHA1String = sig
End Function
Private Function SHA1File(Filename As String) As Byte()
    Dim Fid As String
    Dim sig() As Byte
    Dim Unicode() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    Fid = Filename & vbNullChar
    ReDim sig(mSigBound(HASH_SHA1))
    If HashOSInfo.Unicode Then
        Unicode = Fid
        mResult = PTSHA1FileProgW(Unicode(0), sig(0), AddressOf HashCallBack, HASH_FREQ)
    Else
        mResult = PTSHA1FileProgA(Fid, sig(0), AddressOf HashCallBack, HASH_FREQ)
    End If
    SHA1File = sig
End Function
Private Function SHA1Files(FileNames As Variant) As Byte()
    Dim fids As String
    Dim sig() As Byte
    Dim i As Long
    Dim Unicode() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    If IsArray(FileNames) Then
        For i = 0 To UBound(FileNames)
            fids = fids & FileNames(i) & vbNullChar
        Next i
    Else
        fids = FileNames
    End If
    fids = fids & vbNullChar & vbNullChar

    ReDim sig(mSigBound(HASH_SHA1))
    If HashOSInfo.Unicode Then
        Unicode = fids
        mResult = PTSHA1FilesProgW(Unicode(0), sig(0), AddressOf HashCallBack, HASH_FREQ)
    Else
        mResult = PTSHA1FilesProgA(fids, sig(0), AddressOf HashCallBack, HASH_FREQ)
    End If
    SHA1Files = sig
End Function
Private Function SHA1FilesInput(FileNames As Variant) As HASH_RESULT
    Dim fids As String
    Dim sig() As Byte
    Dim i As Long
    Dim Unicode() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    If IsArray(FileNames) Then
        For i = 0 To UBound(FileNames)
            fids = fids & FileNames(i) & vbNullChar
        Next i
    Else
        fids = FileNames
    End If
    fids = fids & vbNullChar & vbNullChar

    ReDim sig(mSigBound(HASH_SHA1))
    If HashOSInfo.Unicode Then
        Unicode = fids
        mResult = PTSHA1FilesProgInputW(SHA1Ctx(0), Unicode(0), AddressOf HashCallBack, HASH_FREQ)
    Else
        mResult = PTSHA1FilesProgInputA(SHA1Ctx(0), fids, AddressOf HashCallBack, HASH_FREQ)
    End If
    SHA1FilesInput = mResult
End Function
Private Function SHA224String(PlainText As Variant) As Byte()
    Dim buf() As Byte
    Dim buflen As Long
    Dim sig() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    If IsArray(PlainText) Then
        buf = PlainText
        buflen = UBound(buf) + 1
    ElseIf Len(PlainText) = 0 Then
        ReDim buf(0)
        buflen = 0
    Else
        buf = StrConv(PlainText, vbFromUnicode)
        buflen = UBound(buf) + 1
    End If

    ReDim sig(mSigBound(HASH_SHA224))
    mResult = PTSHA224String(buf(0), buflen, sig(0))
    SHA224String = sig
End Function
Private Function SHA224File(Filename As String) As Byte()
    Dim Fid As String
    Dim sig() As Byte
    Dim Unicode() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    Fid = Filename & vbNullChar
    ReDim sig(mSigBound(HASH_SHA224))
    If HashOSInfo.Unicode Then
        Unicode = Fid
        mResult = PTSHA224FileProgW(Unicode(0), sig(0), AddressOf HashCallBack, HASH_FREQ)
    Else
        mResult = PTSHA224FileProgA(Fid, sig(0), AddressOf HashCallBack, HASH_FREQ)
    End If
    SHA224File = sig
End Function
Private Function SHA224Files(FileNames As Variant) As Byte()
    Dim fids As String
    Dim sig() As Byte
    Dim i As Long
    Dim Unicode() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    If IsArray(FileNames) Then
        For i = 0 To UBound(FileNames)
            fids = fids & FileNames(i) & vbNullChar
        Next i
    Else
        fids = FileNames
    End If
    fids = fids & vbNullChar & vbNullChar

    ReDim sig(mSigBound(HASH_SHA224))
    If HashOSInfo.Unicode Then
        Unicode = fids
        mResult = PTSHA224FilesProgW(Unicode(0), sig(0), AddressOf HashCallBack, HASH_FREQ)
    Else
        mResult = PTSHA224FilesProgA(fids, sig(0), AddressOf HashCallBack, HASH_FREQ)
    End If
    SHA224Files = sig
End Function
Private Function SHA224FilesInput(FileNames As Variant) As HASH_RESULT
    Dim fids As String
    Dim sig() As Byte
    Dim i As Long
    Dim Unicode() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    If IsArray(FileNames) Then
        For i = 0 To UBound(FileNames)
            fids = fids & FileNames(i) & vbNullChar
        Next i
    Else
        fids = FileNames
    End If
    fids = fids & vbNullChar & vbNullChar

    ReDim sig(mSigBound(HASH_SHA224))
    If HashOSInfo.Unicode Then
        Unicode = fids
        mResult = PTSHA224FilesProgInputW(SHA224Ctx(0), Unicode(0), AddressOf HashCallBack, HASH_FREQ)
    Else
        mResult = PTSHA224FilesProgInputA(SHA224Ctx(0), fids, AddressOf HashCallBack, HASH_FREQ)
    End If
    SHA224FilesInput = mResult
End Function
Private Function SHA256String(PlainText As Variant) As Byte()
    Dim buf() As Byte
    Dim buflen As Long
    Dim sig() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    If IsArray(PlainText) Then
        buf = PlainText
        buflen = UBound(buf) + 1
    ElseIf Len(PlainText) = 0 Then
        ReDim buf(0)
        buflen = 0
    Else
        buf = StrConv(PlainText, vbFromUnicode)
        buflen = UBound(buf) + 1
    End If

    ReDim sig(mSigBound(HASH_SHA256))
    mResult = PTSHA256String(buf(0), buflen, sig(0))
    SHA256String = sig
End Function
Private Function SHA256File(Filename As String) As Byte()
    Dim Fid As String
    Dim sig() As Byte
    Dim Unicode() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    Fid = Filename & vbNullChar
    ReDim sig(mSigBound(HASH_SHA256))
    If HashOSInfo.Unicode Then
        Unicode = Fid
        mResult = PTSHA256FileProgW(Unicode(0), sig(0), AddressOf HashCallBack, HASH_FREQ)
    Else
        mResult = PTSHA256FileProgA(Fid, sig(0), AddressOf HashCallBack, HASH_FREQ)
    End If
    SHA256File = sig
End Function
Private Function SHA256Files(FileNames As Variant) As Byte()
    Dim fids As String
    Dim sig() As Byte
    Dim i As Long
    Dim Unicode() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    If IsArray(FileNames) Then
        For i = 0 To UBound(FileNames)
            fids = fids & FileNames(i) & vbNullChar
        Next i
    Else
        fids = FileNames
    End If
    fids = fids & vbNullChar & vbNullChar

    ReDim sig(mSigBound(HASH_SHA256))
    If HashOSInfo.Unicode Then
        Unicode = fids
        mResult = PTSHA256FilesProgW(Unicode(0), sig(0), AddressOf HashCallBack, HASH_FREQ)
    Else
        mResult = PTSHA256FilesProgA(fids, sig(0), AddressOf HashCallBack, HASH_FREQ)
    End If
    SHA256Files = sig
End Function
Private Function SHA256FilesInput(FileNames As Variant) As HASH_RESULT
    Dim fids As String
    Dim sig() As Byte
    Dim i As Long
    Dim Unicode() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    If IsArray(FileNames) Then
        For i = 0 To UBound(FileNames)
            fids = fids & FileNames(i) & vbNullChar
        Next i
    Else
        fids = FileNames
    End If
    fids = fids & vbNullChar & vbNullChar

    ReDim sig(mSigBound(HASH_SHA256))
    If HashOSInfo.Unicode Then
        Unicode = fids
        mResult = PTSHA256FilesProgInputW(SHA256Ctx(0), Unicode(0), AddressOf HashCallBack, HASH_FREQ)
    Else
        mResult = PTSHA256FilesProgInputA(SHA256Ctx(0), fids, AddressOf HashCallBack, HASH_FREQ)
    End If
    SHA256FilesInput = mResult
End Function
Private Function SHA384String(PlainText As Variant) As Byte()
    Dim buf() As Byte
    Dim buflen As Long
    Dim sig() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    If IsArray(PlainText) Then
        buf = PlainText
        buflen = UBound(buf) + 1
    ElseIf Len(PlainText) = 0 Then
        ReDim buf(0)
        buflen = 0
    Else
        buf = StrConv(PlainText, vbFromUnicode)
        buflen = UBound(buf) + 1
    End If

    ReDim sig(mSigBound(HASH_SHA384))
    mResult = PTSHA384String(buf(0), buflen, sig(0))
    SHA384String = sig
End Function
Private Function SHA384File(Filename As String) As Byte()
    Dim Fid As String
    Dim sig() As Byte
    Dim Unicode() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    Fid = Filename & vbNullChar
    ReDim sig(mSigBound(HASH_SHA384))
    If HashOSInfo.Unicode Then
        Unicode = Fid
        mResult = PTSHA384FileProgW(Unicode(0), sig(0), AddressOf HashCallBack, HASH_FREQ)
    Else
        mResult = PTSHA384FileProgA(Fid, sig(0), AddressOf HashCallBack, HASH_FREQ)
    End If
    SHA384File = sig
End Function
Private Function SHA384Files(FileNames As Variant) As Byte()
    Dim fids As String
    Dim sig() As Byte
    Dim i As Long
    Dim Unicode() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    If IsArray(FileNames) Then
        For i = 0 To UBound(FileNames)
            fids = fids & FileNames(i) & vbNullChar
        Next i
    Else
        fids = FileNames
    End If
    fids = fids & vbNullChar & vbNullChar

    ReDim sig(mSigBound(HASH_SHA384))
    If HashOSInfo.Unicode Then
        Unicode = fids
        mResult = PTSHA384FilesProgW(Unicode(0), sig(0), AddressOf HashCallBack, HASH_FREQ)
    Else
        mResult = PTSHA384FilesProgA(fids, sig(0), AddressOf HashCallBack, HASH_FREQ)
    End If
    SHA384Files = sig
End Function
Private Function SHA384FilesInput(FileNames As Variant) As HASH_RESULT
    Dim fids As String
    Dim sig() As Byte
    Dim i As Long
    Dim Unicode() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    If IsArray(FileNames) Then
        For i = 0 To UBound(FileNames)
            fids = fids & FileNames(i) & vbNullChar
        Next i
    Else
        fids = FileNames
    End If
    fids = fids & vbNullChar & vbNullChar

    ReDim sig(mSigBound(HASH_SHA384))
    If HashOSInfo.Unicode Then
        Unicode = fids
        mResult = PTSHA384FilesProgInputW(SHA384Ctx(0), Unicode(0), AddressOf HashCallBack, HASH_FREQ)
    Else
        mResult = PTSHA384FilesProgInputA(SHA384Ctx(0), fids, AddressOf HashCallBack, HASH_FREQ)
    End If
    SHA384FilesInput = mResult
End Function
Private Function SHA512String(PlainText As Variant) As Byte()
    Dim buf() As Byte
    Dim buflen As Long
    Dim sig() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    If IsArray(PlainText) Then
        buf = PlainText
        buflen = UBound(buf) + 1
    ElseIf Len(PlainText) = 0 Then
        ReDim buf(0)
        buflen = 0
    Else
        buf = StrConv(PlainText, vbFromUnicode)
        buflen = UBound(buf) + 1
    End If

    ReDim sig(mSigBound(HASH_SHA512))
    mResult = PTSHA512String(buf(0), buflen, sig(0))
    SHA512String = sig
End Function
Private Function SHA512File(Filename As String) As Byte()
    Dim Fid As String
    Dim sig() As Byte
    Dim Unicode() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    Fid = Filename & vbNullChar
    ReDim sig(mSigBound(HASH_SHA512))
    If HashOSInfo.Unicode Then
        Unicode = Fid
        mResult = PTSHA512FileProgW(Unicode(0), sig(0), AddressOf HashCallBack, HASH_FREQ)
    Else
        mResult = PTSHA512FileProgA(Fid, sig(0), AddressOf HashCallBack, HASH_FREQ)
    End If
    SHA512File = sig
End Function
Private Function SHA512Files(FileNames As Variant) As Byte()
    Dim fids As String
    Dim sig() As Byte
    Dim i As Long
    Dim Unicode() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    If IsArray(FileNames) Then
        For i = 0 To UBound(FileNames)
            fids = fids & FileNames(i) & vbNullChar
        Next i
    Else
        fids = FileNames
    End If
    fids = fids & vbNullChar & vbNullChar

    ReDim sig(mSigBound(HASH_SHA512))
    If HashOSInfo.Unicode Then
        Unicode = fids
        mResult = PTSHA512FilesProgW(Unicode(0), sig(0), AddressOf HashCallBack, HASH_FREQ)
    Else
        mResult = PTSHA512FilesProgA(fids, sig(0), AddressOf HashCallBack, HASH_FREQ)
    End If
    SHA512Files = sig
End Function
Private Function SHA512FilesInput(FileNames As Variant) As HASH_RESULT
    Dim fids As String
    Dim sig() As Byte
    Dim i As Long
    Dim Unicode() As Byte

    If Not InitDone Then Intialize
    mCancel = False

    If IsArray(FileNames) Then
        For i = 0 To UBound(FileNames)
            fids = fids & FileNames(i) & vbNullChar
        Next i
    Else
        fids = FileNames
    End If
    fids = fids & vbNullChar & vbNullChar

    ReDim sig(mSigBound(HASH_SHA512))
    If HashOSInfo.Unicode Then
        Unicode = fids
        mResult = PTSHA512FilesProgInputW(SHA512Ctx(0), Unicode(0), AddressOf HashCallBack, HASH_FREQ)
    Else
        mResult = PTSHA512FilesProgInputA(SHA512Ctx(0), fids, AddressOf HashCallBack, HASH_FREQ)
    End If
    SHA512FilesInput = mResult
End Function
Public Function HashErrorMsg(Optional HashError As HASH_RESULT = HASH_DEFAULT) As String
    Dim s As String

    If HashError = HASH_DEFAULT Then HashError = mResult
    s = ""
    Select Case HashError
        Case HASH_OK:           s = ""
        Case HASH_BADSIG:       s = "Invalid Hash Signature"
        Case HASH_BADFID:       s = "Could Not Open File"
        Case HASH_BADSIGLEN:    s = "Wrong Hash Signature Length"
        Case HASH_IO_ERROR:     s = "Error Reading File"
        Case HASH_BADTYPE:      s = "Unrecognized Hash Algorithm (Formula)"
        Case HASH_CANCELLED:    s = "Computation Cancelled by User"
        Case Else:              s = "Unknown Error"
    End Select

    HashErrorMsg = s
End Function
Public Function HashDesc(HashType As HASH_TYPE) As String
    Dim s As String

    Select Case HashType
        Case HASH_NONE:     s = "None/Last Used"
        Case HASH_MD5:      s = "MD5"
        Case HASH_SHA1:     s = "SHA-1"
        Case HASH_SHA224:   s = "SHA-224"
        Case HASH_SHA256:   s = "SHA-256"
        Case HASH_SHA384:   s = "SHA-384"
        Case HASH_SHA512:   s = "SHA-512"
        Case Else:          s = "Unknown"
    End Select

    HashDesc = s
End Function
Public Function HashDisp(sig() As Byte, Optional DispFormat As HASH_DISP_FORMAT = HASH_DISP_LCASE Or HASH_DISP_PLAIN) As String
    Dim s As String
    Dim c As String
    Dim i As Long
    Dim BlockLen As Long

    If Not InitDone Then Intialize

    If DispFormat And HASH_DISP_PRETTY Then
        BlockLen = 4
    Else
        BlockLen = 0
    End If

    s = ""
    For i = 0 To UBound(sig)
        c = Hex(sig(i))
        c = String(2 - Len(c), "0") & c
        s = s & c
        If BlockLen > 0 Then
            If ((i + 1) Mod BlockLen) = 0 Then
                s = s & " "
            End If
        End If
    Next i

    If DispFormat And HASH_DISP_UCASE Then
        s = UCase(s)
    Else
        s = LCase(s)
    End If
    HashDisp = s
End Function
Public Function HashSig2Text(sig() As Byte) As String
    Dim i As Long
    Dim s As String
    Dim c As String

    If Not InitDone Then Intialize

    s = ""
    For i = 0 To UBound(sig)
        c = Hex(sig(i))
        If Len(c) = 1 Then s = s & "0"
        s = s & c
    Next i

    HashSig2Text = s
End Function
Public Function HashText2Sig(SigText As String) As Byte()
    Dim i As Long
    Dim sig() As Byte
    Dim c As String

    If Not InitDone Then Intialize

    If (Len(SigText) Mod 2) <> 0 Then
        mResult = HASH_BADSIGLEN
        Exit Function
    End If

    ReDim sig((Len(SigText) / 2) - 1)
    For i = 0 To UBound(sig)
        c = "&h" & Mid(SigText, (i * 2) + 1, 2)
        sig(i) = Val(c)
    Next i

    HashText2Sig = sig
End Function
Public Function HashLen2Desc(SigText As String) As String
    Dim s As String

    Select Case Len(SigText)
        Case mSigTextLen(HASH_MD5):      s = "MD5"
        Case mSigTextLen(HASH_SHA1):     s = "SHA-1"
        Case mSigTextLen(HASH_SHA224):   s = "SHA-224"
        Case mSigTextLen(HASH_SHA256):   s = "SHA-256"
        Case mSigTextLen(HASH_SHA384):   s = "SHA-384"
        Case mSigTextLen(HASH_SHA512):   s = "SHA-512"
        Case Else:                          s = "Unknown"
    End Select

    HashLen2Desc = s
End Function
Public Function HashLen2Type(SigText As String) As HASH_TYPE
    Dim HashType As HASH_TYPE

    Select Case Len(SigText)
        Case mSigTextLen(HASH_MD5):      HashType = HASH_MD5
        Case mSigTextLen(HASH_SHA1):     HashType = HASH_SHA1
        Case mSigTextLen(HASH_SHA224):   HashType = HASH_SHA224
        Case mSigTextLen(HASH_SHA256):   HashType = HASH_SHA256
        Case mSigTextLen(HASH_SHA384):   HashType = HASH_SHA384
        Case mSigTextLen(HASH_SHA512):   HashType = HASH_SHA512
        Case Else:                          HashType = HASH_NONE
    End Select

    HashLen2Type = HashType
End Function
Public Function HashExt2Type(Ext As String) As HASH_TYPE
    Dim HashType As HASH_TYPE
    Dim s As String

    s = LCase(Ext)
    If Left(Ext, 1) <> "." Then s = "." & s
    Select Case s
        Case ".md5":    HashType = HASH_MD5
        Case ".sha1":   HashType = HASH_SHA1
        Case ".sha224": HashType = HASH_SHA224
        Case ".sha256": HashType = HASH_SHA256
        Case ".sha384": HashType = HASH_SHA384
        Case ".sha512": HashType = HASH_SHA512
        Case Else:      HashType = HASH_NONE
    End Select

    HashExt2Type = HashType
End Function
Public Function HashType2Ext(HashType As HASH_TYPE) As String
    Dim s As String

    Select Case HashType
        Case HASH_MD5:      s = ".md5"
        Case HASH_SHA1:     s = ".sha1"
        Case HASH_SHA224:   s = ".sha224"
        Case HASH_SHA256:   s = ".sha256"
        Case HASH_SHA384:   s = ".sha384"
        Case HASH_SHA512:   s = ".sha512"
        Case Else:          s = ""
    End Select

    HashType2Ext = s
End Function
Public Function HashLegal(SigText As String, HashType As HASH_TYPE) As Boolean
    Dim i As Long
    Dim c As String
    Dim lim As Long

    HashLegal = False
    lim = Len(SigText)
    If lim <> mSigTextLen(HashType) Then Exit Function

    For i = 1 To lim
        c = Mid(SigText, i, 1)
        Select Case c
            Case "0" To "9":
            Case "A" To "F":
            Case "a" To "f":
            Case Else:  Exit Function
        End Select
    Next i

    HashLegal = True
End Function

