Attribute VB_Name = "Base64"
Const Base64Chars$ = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

Public Function Base64Encode(ByVal filepath As String) As String
    Open filepath For Binary As #1
    
    If LOF(1) - Loc(1) > 20000 Then
        InBuffer$ = String$(20000, 0)
    Else
        InBuffer$ = String$(LOF(1) - Loc(1), 0)
    End If
    Get #1, , InBuffer$
    Base64Encode = encode(InBuffer$)
    
    Close #1
End Function

Public Function Base64Decode(ByVal data As String) As String
    Base64Decode = decode(data)
End Function

Private Function decode$(ss$)
   If Len(ss$) Mod 4 > 0 Then ss$ = ss$ + String$(4 - (Len(ss$) Mod 4), " ")

   p% = 0
   tt$ = ""
   For i = 1 To Len(ss$) Step 4
      t$ = "   "
      s$ = Mid$(ss$, i, 4)
      Byte1% = InStr(Base64Chars$, Mid$(s$, 1, 1)) - 1
      Byte2% = InStr(Base64Chars$, Mid$(s$, 2, 1)) - 1
      Byte3% = InStr(Base64Chars$, Mid$(s$, 3, 1)) - 1
      Byte4% = InStr(Base64Chars$, Mid$(s$, 4, 1)) - 1

      Mid$(t$, 1, 1) = Chr$(((Byte2% And 48) \ 16) Or (Byte1% * 4) And &HFF)
      Mid$(t$, 2, 1) = Chr$(((Byte3% And 60) \ 4) Or (Byte2% * 16) And &HFF)
      Mid$(t$, 3, 1) = Chr$((((Byte3% And 3) * 64) And &HFF) Or (Byte4% And 63))

      tt$ = tt$ + t$
      p% = p% + 1: If p% >= 19 Then p% = 0: ss$ = Mid$(ss$, 2)
    Next i
   decode$ = tt$
End Function

Private Function encode$(ss$)
   If Len(ss$) Mod 3 > 0 Then ss$ = ss$ + String$(3 - (Len(ss$) Mod 3), " ")
   
   p% = 0
   tt$ = ""
   For i = 1 To Len(ss$) Step 3
      t$ = "    "
      s$ = Mid$(ss$, i, 3)
      
      Char1% = Asc(Mid$(s$, 1, 1)): SaveBits1% = Char1% And 3
      Char2% = Asc(Mid$(s$, 2, 1)): SaveBits2% = Char2% And 15
      Char3% = Asc(Mid$(s$, 3, 1))
      
      Mid$(t$, 1) = Mid$(Base64Chars$, ((Char1% And 252) \ 4) + 1, 1)
      Mid$(t$, 2) = Mid$(Base64Chars$, (((Char2% And 240) \ 16) Or (SaveBits1% * 16) And &HFF) + 1, 1)
      Mid$(t$, 3) = Mid$(Base64Chars$, (((Char3% And 192) \ 64) Or (SaveBits2% * 4) And &HFF) + 1, 1)
      Mid$(t$, 4) = Mid$(Base64Chars$, (Char3% And 63) + 1, 1)
      tt$ = tt$ + t$
      p% = p% + 1: If p% >= 19 Then p% = 0: tt$ = tt$ & vbCrLf
   Next
   encode$ = tt$
End Function
