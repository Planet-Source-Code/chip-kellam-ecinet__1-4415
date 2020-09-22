VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Upload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private MyRequest As Request
Private MyResponse As Response

Private glngMaxFileBytes As Long
Private gstrUploadPath As String
Private gblnCustomForm As Boolean
Private gblnUserSetPath As Boolean

Private Sub Class_Initialize()
    glngMaxFileBytes = 0
    gblnCustomForm = False
    gblnUserSetPath = False
    gstrUploadPath = ""
End Sub

Public Sub OnStartPage(sc As ScriptingContext)
    Set MyRequest = sc.Request
    Set MyResponse = sc.Response
End Sub

Public Sub Upload()
    Select Case CInt(MyRequest.QueryString("ul"))
        Case 1
            Call UploadFile
        Case Else
            Call ShowUploadForm
    End Select
End Sub

Private Sub ShowUploadForm(Optional strFilename As String = "", Optional lngTotalBytes As Long)
Dim strHTML As String
    
    strHTML = "<font face=""Verdana,Helvetica"" size=2>"
    If strFilename <> "" Then strHTML = strHTML & "File &quot;" & strFilename & "&quot; successfully uploaded.<br>"
    
    If Not gblnCustomForm Then
        strHTML = strHTML & "<form name=""frmECUpload"" action=""ecupload.asp?ul=1"" method=""POST"" enctype=""multipart/form-data"">" _
            & "EC Upload 1.0<br><a href=""http://www.eclink.com"">EC Link, Inc.</a><br><br>" _
            & "<input type=file size=35 name=""ecuData"" style=""font-family:Verdana,Helvetica;font-size:8pt;""><br>" _
            & "[<a href=""javascript:document.frmECUpload.submit();""><font size=1>Upload</font></a>]</form></font>"
    Else
        strHTML = strHTML & "</font>"
    End If
    
    MyResponse.Write strHTML
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' UploadFile
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub UploadFile()
On Error GoTo ErrorCode
    
Dim lngFileDataStart As Long
Dim lngFileDataEnd As Long
Dim lngFilePos As Long
Dim strHeadData As String
Dim strDelimeter As String
Dim strUploadData As String
Dim strFileExtension As String
Dim intFileArgumentStart As Integer
Dim intFileNameEnd As Integer
Dim intFileNameStart As Integer
    
Dim intCount As Integer
Dim intCount2 As Integer
Dim lngCount As Long
Dim strTemp As String
    
Dim varByteCount
Dim binArray() As Byte
    
Dim strFilename As String
Dim lngTotalByteCount As Long
Dim lngTotalFileByteCount As Long
    
    '~~~~~ BYTE COUNT OF RAW FORM DATA ~~~~~~~~~~~
    varByteCount = MyRequest.TotalBytes
    lngTotalByteCount = CLng(varByteCount)
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    '~~~~~ PLACE RAW DATA INTO BYTE ARRAY ~~~~~~~~
    ReDim binArray(varByteCount)
    binArray = MyRequest.BinaryRead(varByteCount)
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    '~~~~~ PARSE HEADER FROM BYTE ARRAY ~~~~~~~~~~
    Do Until Right(strHeadData, 4) = vbCrLf & vbCrLf
        intCount = intCount + 1
        strHeadData = strHeadData & Chr(binArray(intCount))
    Loop
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    '~~~~~ PARSE FILE NAME FROM HEADER ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~ #1 Find the beginning of the name in the HTTP header name/value argument
    intFileArgumentStart = InStr(strHeadData, "ecuData")
    If intFileArgumentStart = 0 Then
        strTemp = "<CENTER><BR><CENTER><HR WIDTH = 50%><FONT FACE=""Verdana, Arial, Helvetica"" COLOR = White SIZE=3>"
        strTemp = strTemp & "<B>FORM INPUT NAME NOT FOUND</B></FONT><HR WIDTH = 50%></CENTER><BR>"
        MyResponse.Write (strTemp)
        Exit Sub
    End If
        
    '~~~~~ #2 Next find the end of the file name sent by users
    intFileNameEnd = InStr(intFileArgumentStart, strHeadData, Chr(34) & vbCrLf)
    intFileNameStart = InStrRev(strHeadData, "\", intFileNameEnd) + 1
    
    '~~~~~ Now parse the file name, 1 indcating no file name sent
    If Not intFileNameStart = 1 Then
        strFilename = Mid(strHeadData, intFileNameStart, intFileNameEnd - intFileNameStart)
    Else
        strTemp = "<CENTER><BR><CENTER><HR WIDTH = 50%><FONT FACE=""Verdana, Arial, Helvetica"" COLOR = White SIZE=3>"
        strTemp = strTemp & "<B>NO FILE SENT</B></FONT><HR WIDTH = 50%></CENTER><BR>"
        MyResponse.Write (strTemp)
        Exit Sub
    End If
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 
    '~~~~~PARSE FILE EXTENSION ~~~~~~~~~~~~~~~~~~~~
    If InStr(strFilename, ".") Then
        strFileExtension = Mid(strFilename, InStrRev(strFilename, ".") + 1)
    Else
        strFileExtension = ""
    End If
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    '~~~~~ CHECK FILE EXTENSIONS ACCEPTABILITY ~~~~
    'If Not blnAcceptableExtension(strFileExtension) Then Exit Sub
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    '~~~~~ PARSE DELIMETER FROM HEADER ~~~~~~~~~~~
    strDelimeter = Left(strHeadData, InStr(strHeadData, vbCrLf) - 1)
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    '~~~~~ START & END OF UPLOAD FILE DATA ~~~~~~~
    lngFileDataStart = Len(strHeadData) + 1
    lngFileDataEnd = CLng(varByteCount) - (Len(strDelimeter) + 8)
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    '~~~~~ SET PATH IF NOT SET BY USER ~~~~~~~~~~~
    If Not gblnUserSetPath Then gstrUploadPath = SetDefaultUpLoadPath
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    '~~~~~ FILE BYTE LENGTH ~~~~~~~~~~~~~~~~~~~~~~
    lngTotalFileByteCount = lngFileDataEnd - lngFileDataStart + 1
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    
    '~~~~~ EXIT IF FILE IS SIZE NOT ALLOWED ~~~~~~
    If Not glngMaxFileBytes = 0 Then
        If lngTotalFileByteCount > glngMaxFileBytes Then
            strTemp = "<font face=""Verdana, Arial, Helvetica"" size=2>"
            strTemp = strTemp & "<b>File length  [" & CStr(lngTotalFileByteCount) & " bytes] is larger<BR>than permitted settings [" & CStr(glngMaxFileBytes) & " bytes].</b></font><br>"
            MyResponse.Write (strTemp)
            Exit Sub
        End If
    End If
    
    '~~~~~ SAVE THE FILE DATA
    Open gstrUploadPath & strFilename For Binary Access Write As #1
    For lngCount = lngFileDataStart To lngFileDataEnd
        Put #1, , (binArray(lngCount))
    Next
    Close #1

    Call ShowUploadForm(strFilename, lngTotalFileByteCount)
Exit Sub

ErrorCode:
    strTemp = "<br><b>" & Err.Description & "</b><br>"
    MyResponse.Write (strTemp)
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' SetDefaultUpLoadPath
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Function SetDefaultUpLoadPath() As String
Dim intCount As Integer

    SetDefaultUpLoadPath = MyRequest.ServerVariables("PATH_TRANSLATED")
    For intCount = Len(SetDefaultUpLoadPath) To 1 Step -1
        If (Mid(SetDefaultUpLoadPath, intCount, 1)) = "\" Then
            SetDefaultUpLoadPath = Left(SetDefaultUpLoadPath, intCount)
            Exit For
        End If
    Next
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' PROPERTIES
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Property Let MaxFileSizeInBytes(ByVal vNewValue As Long)
    glngMaxFileBytes = vNewValue
End Property

Public Property Let UploadPath(ByVal vNewValue As String)
    gstrUploadPath = vNewValue
    gblnUserSetPath = True
End Property

Public Property Let CustomForm(ByVal vNewValue As Boolean)
    gblnCustomForm = vNewValue
End Property

