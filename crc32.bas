Attribute VB_Name = "Module1"
Option Explicit
Option Compare Text

Private Crc32Table(0 To 255) As Long

Private Function InitCrc32(Optional ByVal Seed As Long = &HEDB88320, Optional ByVal Precondition As Long = &HFFFFFFFF) As Long
  Dim iBytes As Integer, iBits As Integer, lCrc32 As Long, lTempCrc32 As Long
  On Error Resume Next

  For iBytes = 0 To 255
    lCrc32 = iBytes

    For iBits = 0 To 7
      lTempCrc32 = lCrc32 And &HFFFFFFFE
      lTempCrc32 = lTempCrc32 \ &H2
      lTempCrc32 = lTempCrc32 And &H7FFFFFFF

      If (lCrc32 And &H1) <> 0 Then
        lCrc32 = lTempCrc32 Xor Seed
      Else
        lCrc32 = lTempCrc32
      End If
    Next
    Crc32Table(iBytes) = lCrc32
  Next
  InitCrc32 = Precondition
End Function

Private Function AddCrc32(ByVal Item As String, ByVal Crc32 As Long) As Long
  Dim bCharValue As Byte, iCounter As Integer, lIndex As Long
  Dim lAccValue As Long, lTableValue As Long
  On Error Resume Next

  For iCounter = 1 To Len(Item)
    bCharValue = Asc(Mid$(Item, iCounter, 1))
    lAccValue = Crc32 And &HFFFFFF00
    lAccValue = lAccValue \ &H100
    lAccValue = lAccValue And &HFFFFFF
    lIndex = Crc32 And &HFF
    lIndex = lIndex Xor bCharValue
    lTableValue = Crc32Table(lIndex)
    Crc32 = lAccValue Xor lTableValue
  Next
  AddCrc32 = Crc32
End Function

Private Function GetCrc32(ByVal Crc32 As Long) As Long
  On Error Resume Next
  GetCrc32 = Crc32 Xor &HFFFFFFFF
End Function

Public Function Compute(ToGet As String) As String
  Dim lCrc32Value As Long
  On Error Resume Next
  lCrc32Value = InitCrc32()
  lCrc32Value = AddCrc32(ToGet, lCrc32Value)
  Compute = Hex$(GetCrc32(lCrc32Value))
End Function



