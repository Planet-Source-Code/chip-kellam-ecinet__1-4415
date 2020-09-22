VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Email"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' CONSTANTS
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Const APP_NAME = "ecMail"
Private Const APP_VERSION = "1.04.201"
Private Const APP_MAILER_ID = APP_NAME & " " & APP_VERSION & ""

Private Const SEND_TO = 1
Private Const SEND_CC = 2
Private Const SEND_BCC = 3

Private Const FORMAT_TEXT = 1
Private Const FORMAT_HTML = 2

Private Const ENCODE_BASE64 = 1
Private Const ENCODE_UUENCODE = 2
Private Const ENCODE_QUOTED_PRINTABLE = 3
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' GLOBAL VARIABLE DECLARATION
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private MyResponse As Response

Private gstrHost As String
Private gstrFrom As String, gstrFromName As String, gstrReplyTo As String
Private gstrTo() As String, gstrCc() As String, gstrBcc() As String
Private gstrSubject As String, gstrBody As String
Private gstrPriority As Integer       '1=high / 3=normal/ 5=low

Private gstrBodyFormat As String, gstrEncoding As String, gstrCharset As String

Private blnTo As Boolean, blnCc As Boolean, blnBcc As Boolean

Private gstrAttachments() As String
Private gstrAttachmentName() As String
Private gstrAttachmentType() As String
Private blnAttachments As Boolean

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Class Initializer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Class_Initialize()
    gstrHost = ""
    gstrFrom = "": gstrFromName = ""
    gstrSubject = "": gstrBody = ""
    gstrPriority = 3
    
    gstrBodyFormat = "text/plain"
    gstrCharset = "iso-8859-1"
    gstrEncoding = "base64"
    
    blnTo = False
    blnCc = False
    blnBcc = False
    blnAttachments = False
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' ASP Initializer
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Sub OnStartPage(sc As ScriptingContext)
    Set MyResponse = sc.Response()
End Sub

Public Sub AddRecipient(method As Integer, ByVal email As String, Optional ByVal name As String = "")
Dim sname As String
Dim i As Integer

    On Error Resume Next
    If name <> "" Then sname = Chr(34) & name & Chr(34) & " "
    i = 0
    
    Select Case method
        Case SEND_TO
            i = UBound(gstrTo)
            ReDim Preserve gstrTo(i + 1)
            gstrTo(i) = sname & "<" & Trim(email) & ">"
            blnTo = True
        Case SEND_CC
            i = UBound(gstrCc)
            ReDim Preserve gstrCc(i + 1)
            gstrCc(i) = sname & "<" & Trim(email) & ">"
            blnCc = True
        Case SEND_BCC
            i = UBound(gstrBcc)
            ReDim Preserve gstrBcc(i + 1)
            gstrBcc(i) = sname & "<" & Trim(email) & ">"
            blnBcc = True
    End Select
End Sub

Public Sub RemoveRecipients(Optional ByVal method As Integer = 0)
    Select Case method
        Case SEND_TO
            Erase gstrTo: blnTo = False
        Case SEND_CC
            Erase gstrCc: blnCc = False
        Case SEND_BCC
            Erase gstrBcc: blnBcc = False
        Case 0
            Erase gstrTo: blnTo = False
            Erase gstrCc: blnCc = False
            Erase gstrBcc: blnBcc = False
    End Select
End Sub

Public Sub AddAttachment(ByVal sFilePath As String, Optional iEncodeScheme As Integer = ENCODE_BASE64)
Dim i As Integer
Dim pos As Long

    i = 0
    If blnAttachments Then
        i = UBound(gstrAttachments)
        ReDim Preserve gstrAttachments(i + 1)
        ReDim Preserve gstrAttachmentName(i + 1)
        ReDim Preserve gstrAttachmentType(i + 1)
    Else
        blnAttachments = True
        ReDim gstrAttachments(i + 1)
        ReDim gstrAttachmentName(i + 1)
        ReDim gstrAttachmentType(i + 1)
    End If
    
    'get filename
    pos = InStrRev(sFilePath, "/")
    If pos <= 0 Then pos = InStrRev(sFilePath, "\")
    gstrAttachmentName(i) = Mid(sFilePath, pos + 1)
    
    'get MIME type
    pos = InStrRev(gstrAttachmentName(i), ".")
    If pos <= 0 Then
        gstrAttachmentType(i) = "application/octet-stream"
    Else
        gstrAttachmentType(i) = MIMETypeLookup(Mid(gstrAttachmentName(i), pos + 1))
    End If
    
    Select Case iEncodeScheme
        Case ENCODE_BASE64
            gstrAttachments(i) = Base64Encode(sFilePath)
    End Select
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Send
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Public Function Send() As Long
    If blnAttachments Then
        Send = SendWithAttachments
    Else
        Send = SendWithoutAttachments
    End If
End Function

Private Function SendWithoutAttachments() As Long
Dim Status, crlf$, i, s
Dim pos As Integer
Dim strTo As String, strCc As String
   
   On Error GoTo err_handler
   strTo = "": strCc = ""
   
   If Len(gstrFrom) < 1 Or Len(gstrBody) < 1 Or Len(gstrHost) < 1 Or ((Not blnTo) And (Not blnCc) And (Not blnBcc)) Then
      SendWithoutAttachments = -2: Exit Function
   End If

   s = ConnectSock(gstrHost, "smtp", 0, False)
   If s = INVALID_SOCKET Then SendWithoutAttachments = -1: Exit Function
   
   crlf$ = Chr(13) & Chr(10)
   Status = -3
   
   Do
      Status = WaitForStatus(s)
      If Status <> 220 Then Exit Do
      
      i = SendData(s, "helo " & GetLocalHostName() & crlf$)
      Status = WaitForStatus(s)
      If Status <> 250 Then Exit Do
      
      i = SendData(s, "mail from: <" & gstrFrom & ">" & crlf$)
      Status = WaitForStatus(s)
      If Status <> 250 And Status <> 251 Then Exit Do
      
      'if there are send to recipients...
      If blnTo Then
        For pos = 0 To UBound(gstrTo) - 1
          'the mid call here is to parse out the name of the recipient if it is present
          i = SendData(s, "rcpt to:" & Mid(gstrTo(pos), InStr(1, gstrTo(pos), "<")) & crlf$)
          Status = WaitForStatus(s)
          If Status <> 250 And Status <> 251 Then Exit Do
          
          strTo = strTo & gstrTo(pos)
          If pos < UBound(gstrTo) - 1 Then strTo = strTo & ","
        Next
      End If
      
      'if there are carbon copy recipients...
      If blnCc Then
        For pos = 0 To UBound(gstrCc) - 1
          i = SendData(s, "rcpt to:" & Mid(gstrCc(pos), InStr(1, gstrCc(pos), "<")) & crlf$)
          Status = WaitForStatus(s)
          If Status <> 250 And Status <> 251 Then Exit Do
      
          strCc = strCc & gstrCc(pos)
          If pos < UBound(gstrCc) - 1 Then strCc = strCc & ","
        Next
      End If
      
      'if there are blind carbon copy recipients...
      If blnBcc Then
        For pos = 0 To UBound(gstrBcc) - 1
          i = SendData(s, "rcpt to:" & Mid(gstrBcc(pos), InStr(1, gstrBcc(pos), "<")) & crlf$)
          Status = WaitForStatus(s)
          If Status <> 250 And Status <> 251 Then Exit Do
        Next
      End If
      
      i = SendData(s, "data" & crlf$)
      Status = WaitForStatus(s)
      If Status <> 354 Then Exit Do

      i = SendData(s, "MIME-Version: 1.0" & crlf$)
      i = SendData(s, "Content-Type: " & gstrBodyFormat & "; charset=" & Chr(34) & gstrCharset & Chr(34) & crlf$)
      i = SendData(s, "Date: " & Now() & crlf$)
      i = SendData(s, "To: " & strTo & crlf$)
      If strCc <> "" Then i = SendData(s, "Cc: " & strCc & crlf$)
      If gstrFromName <> "" Then
        i = SendData(s, "From: " & Chr(34) & gstrFromName & Chr(34) & " <" & gstrFrom & ">" & crlf$)
      Else
        i = SendData(s, "From: <" & gstrFrom & ">" & crlf$)
      End If
      If gstrReplyTo <> "" Then i = SendData(s, "Reply-To: <" & gstrReplyTo & ">" & crlf$)
      If Len(gstrSubject) > 0 Then i = SendData(s, "Subject: " & gstrSubject & crlf$)
      i = SendData(s, "X-Priority: " & gstrPriority & crlf$)
      i = SendData(s, "X-Mailer: " & APP_MAILER_ID & crlf$)
      i = SendData(s, crlf$)
      i = SendData(s, gstrBody & crlf$)
      i = SendData(s, "." & crlf$)
      
      Status = WaitForStatus(s)
      If Status <> 250 Then Exit Do

      i = SendData(s, "quit" & crlf$)
      Status = WaitForStatus(s)
      Status = 0
      Exit Do
   Loop
   SendWithoutAttachments = Status
   Exit Function
    
err_handler:
    MyResponse.Write "<font color=""#ff0000"" size=2 face=""Verdana""><b>An error occured.<br>" & Err.number & ": " & Err.Description & "</b></font>"
End Function

Private Function SendWithAttachments() As Long
Dim Status, crlf$, i, s
Dim pos As Integer
Dim strTo As String, strCc As String
Dim b As String
Dim newBody As String
   
   On Error GoTo err_handler
   strTo = "": strCc = ""
   
   If Len(gstrFrom) < 1 Or Len(gstrHost) < 1 Or ((Not blnTo) And (Not blnCc) And (Not blnBcc)) Then
      SendWithAttachments = -2: Exit Function
   End If

   s = ConnectSock(gstrHost, "smtp", 0, False)
   If s = INVALID_SOCKET Then SendWithAttachments = -1: Exit Function
   
   crlf$ = Chr(13) & Chr(10)
   Status = -3
   
   Do
      Status = WaitForStatus(s)
      If Status <> 220 Then Exit Do
      
      i = SendData(s, "helo " & GetLocalHostName() & crlf$)
      Status = WaitForStatus(s)
      If Status <> 250 Then Exit Do
      
      i = SendData(s, "mail from: <" & gstrFrom & ">" & crlf$)
      Status = WaitForStatus(s)
      If Status <> 250 And Status <> 251 Then Exit Do
      
      'if there are send to recipients...
      If blnTo Then
        For pos = 0 To UBound(gstrTo) - 1
          'the mid call here is to parse out the name of the recipient if it is present
          i = SendData(s, "rcpt to:" & Mid(gstrTo(pos), InStr(1, gstrTo(pos), "<")) & crlf$)
          Status = WaitForStatus(s)
          If Status <> 250 And Status <> 251 Then Exit Do
          
          strTo = strTo & gstrTo(pos)
          If pos < UBound(gstrTo) - 1 Then strTo = strTo & ","
        Next
      End If
      
      'if there are carbon copy recipients...
      If blnCc Then
        For pos = 0 To UBound(gstrCc) - 1
          i = SendData(s, "rcpt to:" & Mid(gstrCc(pos), InStr(1, gstrCc(pos), "<")) & crlf$)
          Status = WaitForStatus(s)
          If Status <> 250 And Status <> 251 Then Exit Do
      
          strCc = strCc & gstrCc(pos)
          If pos < UBound(gstrCc) - 1 Then strCc = strCc & ","
        Next
      End If
      
      'if there are blind carbon copy recipients...
      If blnBcc Then
        For pos = 0 To UBound(gstrBcc) - 1
          i = SendData(s, "rcpt to:" & Mid(gstrBcc(pos), InStr(1, gstrBcc(pos), "<")) & crlf$)
          Status = WaitForStatus(s)
          If Status <> 250 And Status <> 251 Then Exit Do
        Next
      End If
      
      i = SendData(s, "data" & crlf$)
      Status = WaitForStatus(s)
      If Status <> 354 Then Exit Do

      b = GenerateMIMEBoundary
      
      newBody = "This is a multi-part message in MIME format."
      If Len(gstrBody) > 0 Then
        newBody = newBody & crlf$ & crlf$ & "--" & b & crlf$
        newBody = newBody & "Content-Type: " & gstrBodyFormat & ";" & crlf$ & "        charset=" & Chr(34) & gstrCharset & Chr(34) & crlf$
        newBody = newBody & "Content-Transfer-Encoding: 7bit" & crlf$ & crlf$
        newBody = newBody & gstrBody
      End If
      For pos = 0 To UBound(gstrAttachments) - 1
        newBody = newBody & crlf$ & crlf$ & "--" & b & crlf$
        newBody = newBody & "Content-Type: " & gstrAttachmentType(pos) & ";" & crlf$ & "        name=" & Chr(34) & gstrAttachmentName(pos) & Chr(34) & crlf$
        newBody = newBody & "Content-Transfer-Encoding: base64" & crlf$
        newBody = newBody & "Content-Disposition: attachment; filename=" & Chr(34) & gstrAttachmentName(pos) & Chr(34) & crlf$ & crlf$
        newBody = newBody & gstrAttachments(pos)
      Next
      newBody = newBody & crlf$ & crlf$ & "--" & b & "--" & crlf$
      
      If gstrFromName <> "" Then
        i = SendData(s, "From: " & Chr(34) & gstrFromName & Chr(34) & " <" & gstrFrom & ">" & crlf$)
      Else
        i = SendData(s, "From: <" & gstrFrom & ">" & crlf$)
      End If
      If gstrReplyTo <> "" Then i = SendData(s, "Reply-To: <" & gstrReplyTo & ">" & crlf$)
      i = SendData(s, "To: " & strTo & crlf$)
      If strCc <> "" Then i = SendData(s, "Cc: " & strCc & crlf$)
      If Len(gstrSubject) > 0 Then i = SendData(s, "Subject: " & gstrSubject & crlf$)
      i = SendData(s, "Date: " & Now() & crlf$)
      i = SendData(s, "MIME-Version: 1.0" & crlf$)
      i = SendData(s, "Content-Type: multipart/mixed;" & crlf$ & "        boundary=" & Chr(34) & b & Chr(34) & crlf$) '";" & crlf$ & "        type=" & Chr(34) & "text/plain" & Chr(34) & crlf$)
      i = SendData(s, "X-Priority: " & gstrPriority & crlf$)
      i = SendData(s, "X-Mailer: " & APP_MAILER_ID & crlf$)
      'i = SendData(s, "X-MimeOLE: Produced by ecMimeOLE 1.03" & crlf$)
      i = SendData(s, crlf$)
      i = SendData(s, newBody & crlf$)
      i = SendData(s, "." & crlf$)
      
      Status = WaitForStatus(s)
      If Status <> 250 Then Exit Do

      i = SendData(s, "quit" & crlf$)
      Status = WaitForStatus(s)
      Status = 0
      Exit Do
   Loop
   SendWithAttachments = Status
   Exit Function
    
err_handler:
    MyResponse.Write "<font color=""#ff0000"" size=2 face=""Verdana""><b>An error occured.<br>" & Err.number & ": " & Err.Description & "</b></font>"
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' PROPERTIES
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Public Property Let host(ByVal sNewValue As String)
    gstrHost = sNewValue
End Property

Public Property Let from(ByVal sNewValue As String)
    gstrFrom = sNewValue
End Property

Public Property Let FromName(ByVal sNewValue As String)
    gstrFromName = sNewValue
End Property

Public Property Let ReplyTo(ByVal sNewValue As String)
    gstrReplyTo = sNewValue
End Property

Public Property Let sendto(ByVal sNewValue As String)
Dim pos As Long
Dim i As Integer

    i = 1
    pos = InStr(1, sNewValue, ",")
    If pos > 0 Then
        Do While pos > 0
            ReDim Preserve gstrTo(i)
            gstrTo(i - 1) = "<" & Left(LTrim(sNewValue), pos - 1) & ">"
            sNewValue = LTrim(Mid(sNewValue, pos + 1))
            pos = InStr(1, sNewValue, ",")
            i = i + 1
        Loop
    End If
    
    If Trim(sNewValue) <> "" Then
        ReDim Preserve gstrTo(i)
        gstrTo(i - 1) = "<" & Trim(sNewValue) & ">"
        blnTo = True
    End If
End Property

Public Property Get sendto() As String
    sendto = Join(gstrTo, ",")
End Property

Public Property Let Cc(ByVal sNewValue As String)
Dim pos As Long
Dim i As Integer

    i = 1
    pos = InStr(1, sNewValue, ",")
    If pos > 0 Then
        Do While pos > 0
            ReDim Preserve gstrCc(i)
            gstrCc(i - 1) = "<" & Left(LTrim(sNewValue), pos - 1) & ">"
            sNewValue = LTrim(Mid(sNewValue, pos + 1))
            pos = InStr(1, sNewValue, ",")
            i = i + 1
        Loop
    End If
    
    If Trim(sNewValue) <> "" Then
        ReDim Preserve gstrCc(i)
        gstrCc(i - 1) = "<" & Trim(sNewValue) & ">"
        blnCc = True
    End If
End Property

Public Property Get Cc() As String
    Cc = Join(gstrCc, ",")
End Property

Public Property Let Bcc(ByVal sNewValue As String)
Dim pos As Long
Dim i As Integer

    i = 1
    pos = InStr(1, sNewValue, ",")
    If pos > 0 Then
        Do While pos > 0
            ReDim Preserve gstrBcc(i)
            gstrBcc(i - 1) = "<" & Left(LTrim(sNewValue), pos - 1) & ">"
            sNewValue = LTrim(Mid(sNewValue, pos + 1))
            pos = InStr(1, sNewValue, ",")
            i = i + 1
        Loop
    End If
    
    If Trim(sNewValue) <> "" Then
        ReDim Preserve gstrBcc(i)
        gstrBcc(i - 1) = "<" & Trim(sNewValue) & ">"
        blnBcc = True
    End If
End Property

Public Property Get Bcc() As String
    Bcc = Join(gstrBcc, ",")
End Property

Public Property Let Subject(ByVal sNewValue As String)
    gstrSubject = sNewValue
End Property

Public Property Let Body(ByVal sNewValue As String)
    gstrBody = sNewValue
End Property

Public Property Let BodyFormat(ByVal sNewValue As Integer)
    Select Case sNewValue
        Case FORMAT_TEXT
            gstrBodyFormat = "text/plain"
        Case FORMAT_HTML
            gstrBodyFormat = "text/html"
    End Select
End Property

Public Property Let Charset(ByVal sNewValue As String)
    gstrCharset = sNewValue
End Property

Public Property Let Priority(ByVal sNewValue As Integer)
    gstrPriority = sNewValue
End Property
