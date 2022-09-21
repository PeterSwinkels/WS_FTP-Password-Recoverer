Attribute VB_Name = "WSFTPRecoverModule"
'This module contains this program's procedures and interface.
Option Explicit

Private Const ENCODED_PASSWORD_PREFIX As String = "PWD=V"   'Defines the prefix for encoded passwords.


'This procedure decodes the specified encoded password and returns the decoded result.
Private Function Decode(Encoded As String) As String
On Error GoTo ErrorTrap
Dim Decoded As String
Dim DecodedValue As Long
Dim Position As Long
Dim Salt As String
Dim SaltValue As Long

   Decoded = vbNullString
   Encoded = Mid$(Encoded, Len(ENCODED_PASSWORD_PREFIX) + 1)
   Salt = Left$(Encoded, 32)
   Encoded = Mid$(Encoded, Len(Salt) + 1)
   For Position = 1 To Len(Encoded) Step 2
      SaltValue = CLng(Val("&H" & Mid$(Salt, (Position \ 2) + 1, 1) & "&"))
      SaltValue = SaltValue + 47
      SaltValue = SaltValue Mod 57
      
      DecodedValue = Val("&H" & Mid$(Encoded, Position, 2) & "&")
      DecodedValue = DecodedValue - ((Position \ 2) + 1)
      DecodedValue = DecodedValue - SaltValue
      
      Decoded = Decoded & Chr$(DecodedValue)
   Next Position
   
EndRoutine:
   Decode = Decoded
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure encodes the specified password and returns the encoded result.
Private Function Encode(Unencoded As String) As String
On Error GoTo ErrorTrap
Dim Encoded As String
Dim EncodedValue As Long
Dim Position As Long
Dim Salt As String
Dim SaltValue As Long
  
   Salt = vbNullString
   For Position = 1 To 32
      Salt = Salt & UCase$(Hex$(Int(Rnd * 15)))
   Next Position
   
   Encoded = ENCODED_PASSWORD_PREFIX & Salt
   For Position = 1 To Len(Unencoded)
      SaltValue = Val("&H" & Mid$(Salt, Position, 1) & "&")
      SaltValue = SaltValue + 47
      SaltValue = SaltValue Mod 57
      
      EncodedValue = Asc(Mid$(Unencoded, Position, 1))
      EncodedValue = EncodedValue + Position
      EncodedValue = EncodedValue + SaltValue
      
      Encoded = Encoded & UCase$(Hex$(EncodedValue))
   Next Position

EndRoutine:
   Encode = Encoded
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure handles any errors that occur.
Private Sub HandleError()
Dim ErrorCode As Long
Dim Message As String

   ErrorCode = Err.Number
   Message = Err.Description
   
   On Error Resume Next
   
   If MsgBox("Error: " & CStr(ErrorCode) & vbCr & Message, vbOKCancel Or vbExclamation) = vbCancel Then
      End
   End If
End Sub


'This procedure is executed when this program is started.
Public Sub Main()
On Error GoTo ErrorTrap
Dim Password As String

   Randomize
   
   Password = vbNullString
   Do
      Password = InputBox$("Specify a password (include """ & ENCODED_PASSWORD_PREFIX & """ for encoded passwords):", , Password)
      If Not Password = vbNullString Then
         If Left$(Password, Len(ENCODED_PASSWORD_PREFIX)) = ENCODED_PASSWORD_PREFIX Then
            Password = Decode(Password)
         Else
            Password = Encode(Password)
         End If
      End If
   Loop Until Password = vbNullString
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


