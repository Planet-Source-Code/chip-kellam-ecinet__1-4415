Attribute VB_Name = "MIME"
Option Explicit
Private Const ecMIME_VERSION = "0001.503.2"
Private MIMETypes As String

Public Sub InitMIMETypes()
    MIMETypes = ""
    
    'TEXT TYPES
    MIMETypes = MIMETypes & "text/html  <html> <htm> <shtml> <cgi> <asp> <inc>" & vbCrLf
    MIMETypes = MIMETypes & "text/plain  <txt>" & vbCrLf
    MIMETypes = MIMETypes & "text/richtext  <rtx>" & vbCrLf
    MIMETypes = MIMETypes & "text/tab-separated-values  <tsv>" & vbCrLf
    MIMETypes = MIMETypes & "text/x-setext  <etx>" & vbCrLf
    MIMETypes = MIMETypes & "text/x-sgml  <sgml> <sgm>" & vbCrLf
    
    'IMAGE TYPES
    MIMETypes = MIMETypes & "image/gif  <gif>" & vbCrLf
    MIMETypes = MIMETypes & "image/ief  <ief>" & vbCrLf
    MIMETypes = MIMETypes & "image/jpeg  <jpeg> <jpg> <jpe>" & vbCrLf
    MIMETypes = MIMETypes & "image/png  <png>" & vbCrLf
    MIMETypes = MIMETypes & "image/tiff  <tiff> <tif>" & vbCrLf
    
    'APPLICATION TYPES
    MIMETypes = MIMETypes & "application/mac-binhex40  <hqx>" & vbCrLf
    MIMETypes = MIMETypes & "application/msword  <doc>" & vbCrLf
    MIMETypes = MIMETypes & "application/octet-stream  <bin> <dms> <lha> <lzh> <exe> <class>" & vbCrLf
    MIMETypes = MIMETypes & "application/oda  <oda>" & vbCrLf
    MIMETypes = MIMETypes & "application/pdf  <pdf>" & vbCrLf
    MIMETypes = MIMETypes & "application/postscript  <ai> <eps> <ps>" & vbCrLf
    MIMETypes = MIMETypes & "application/powerpoint  <ppt>" & vbCrLf
    MIMETypes = MIMETypes & "application/rtf  <rtf>" & vbCrLf
    MIMETypes = MIMETypes & "application/x-stuffit  <sit>" & vbCrLf
    MIMETypes = MIMETypes & "application/x-tar  <tar>" & vbCrLf
    MIMETypes = MIMETypes & "application/x-wais-source  <src>" & vbCrLf
    MIMETypes = MIMETypes & "application/zip  <zip>" & vbCrLf
    
    'AUDIO TYPES
    MIMETypes = MIMETypes & "audio/basic  <au> <snd>" & vbCrLf
    MIMETypes = MIMETypes & "audio/mpeg  <mpga> <mp2> <mp3>" & vbCrLf
    MIMETypes = MIMETypes & "audio/x-aiff  <aif> <aiff> <aifc>" & vbCrLf
    MIMETypes = MIMETypes & "audio/x-pn-realaudio  <ram>" & vbCrLf
    MIMETypes = MIMETypes & "audio/x-pn-realaudio-plugin  <rpm>" & vbCrLf
    MIMETypes = MIMETypes & "audio/x-realaudio  <ra>" & vbCrLf
    MIMETypes = MIMETypes & "audio/x-wav  <wav>" & vbCrLf
    
    'VIDEO TYPES
    MIMETypes = MIMETypes & "video/mpeg  <mpeg> <mpg> <mpe>" & vbCrLf
    MIMETypes = MIMETypes & "video/quicktime  <qt> <mov>" & vbCrLf
    MIMETypes = MIMETypes & "video/x-msvideo  <avi>" & vbCrLf
End Sub

Public Function MIMETypeLookup(ByVal sExt As String) As String
Dim pos As Long
Dim sTemp As String

    sExt = LCase(sExt)
    'make sure string is initialized
    If MIMETypes = "" Then InitMIMETypes
    'make sure we have the extension in the right format
    If Left(sExt, 1) = "." Then sExt = Mid(sExt, 2)
    If Left(sExt, 1) <> "<" Then sExt = "<" & sExt
    If Right(sExt, 1) <> ">" Then sExt = sExt & ">"
    
    'begin lookup
    pos = InStr(1, MIMETypes, sExt)
    If pos > 0 Then
        pos = InStrRev(MIMETypes, vbCrLf, pos)
        MIMETypeLookup = Mid(MIMETypes, pos + 2, InStr(pos + 2, MIMETypes, "  ") - pos - 2)
    Else
        MIMETypeLookup = "mime/x-type-unknown"
    End If
End Function

Public Function GenerateMIMEBoundary() As String
Dim seed As Long

    Randomize
    seed = Int((98712 * Rnd) + 500)
    
    GenerateMIMEBoundary = "NextPart=" & ecMIME_VERSION & "_" & Trim(Str(seed))
End Function
