/'
   Basic Windows Registry Read/Write Functions
   Translated from VB6 source by
      -Vincent DeCampo, 2010
   
   Original author UNKOWN    
'/

#Include "windows.bi"
#Include "vbcompat.bi"

' Possible registry data types
Enum InTypes
   ValNull = 0
   ValString = 1
   ValXString = 2
   ValBinary = 3
   ValDWord = 4
   ValLink = 6
   ValMultiString = 7
   ValResList = 8
End Enum

' Registry section definitions
'Const HKEY_CLASSES_ROOT = &H80000000
'Const HKEY_CURRENT_USER = &H80000001
'Const HKEY_LOCAL_MACHINE = &H80000002
'Const HKEY_USERS = &H80000003
'Const HKEY_PERFORMANCE_DATA = &H80000004
'Const HKEY_CURRENT_CONFIG = &H80000005

Function ReadRegistry(ByVal Group as HKEY, ByVal Section As LPCSTR, ByVal Key As LPCSTR) As String
Dim as DWORD lDataTypeValue, lValueLength
Dim sValue As String * 2048
Dim As String Tstr1, Tstr2  
Dim lKeyValue As HKEY
Dim lResult as Integer
Dim td As Double

   sValue = ""
   
   lResult      = RegOpenKey(Group, Section, @lKeyValue)
   lValueLength = Len(sValue)
   lResult      = RegQueryValueEx(lKeyValue, Key, 0&, @lDataTypeValue, Cast(Byte Ptr,@sValue), @lValueLength)
   
   If (lResult = 0) Then

      Select Case lDataTypeValue
         case REG_DWORD 
            td = Asc(Mid(sValue, 1, 1)) + &H100& * Asc(Mid(sValue, 2, 1)) + &H10000 * Asc(Mid(sValue, 3, 1)) + &H1000000 * CDbl(Asc(Mid(sValue, 4, 1)))
            sValue = Format(td, "000")
         case REG_BINARY 
            ' Return a binary field as a hex string (2 chars per byte)
            Tstr2 = ""
            For I As Integer = 1 To lValueLength
               Tstr1 = Hex(Asc(Mid(sValue, I, 1)))
               If Len(Tstr1) = 1 Then Tstr1 = "0" & Tstr1
               Tstr2 += Tstr1
            Next
            sValue = Tstr2
         Case Else
            sValue = Left(sValue, lValueLength - 1)
      End Select
   
   End If

   lResult = RegCloseKey(lKeyValue)
   
   Return sValue

End Function

Sub WriteRegistry(ByVal Group as HKEY, ByVal Section As LPCSTR, ByVal Key As LPCSTR, ByVal ValType As InTypes, value As String)
Dim lResult as Integer
Dim lKeyValue As HKEY
Dim lNewVal as DWORD
Dim sNewVal As String * 2048

   lResult = RegCreateKey(Group, Section, @lKeyValue)

   If ValType = ValDWord Then
      lNewVal = CUInt(value)
      lResult = RegSetValueEx(lKeyValue, Key, 0&, ValType, Cast(Byte Ptr,@lNewVal), SizeOf(DWORD))
   Else
      If ValType = ValString Or ValType = ValXString Then
         sNewVal = value & Chr(0)
         lResult = RegSetValueEx(lKeyValue, Key, 0&, ValType, Cast(Byte Ptr,@sNewVal), Len(sNewVal))
      EndIf
   End If

   lResult = RegFlushKey(lKeyValue)
   lResult = RegCloseKey(lKeyValue)

End Sub