Attribute VB_Name = "modGDI"
Option Explicit

'  Ternary raster operations
Public Enum enmBrush
  SRCCOPY = &HCC0020 ' (DWORD) dest = source
  SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
  SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
  SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
  SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
  NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
  NOTSRCERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)
  MERGECOPY = &HC000CA       ' (DWORD) dest = (source AND pattern)
  MERGEPAINT = &HBB0226      ' (DWORD) dest = (NOT source) OR dest
  PATCOPY = &HF00021 ' (DWORD) dest = pattern
  PATPAINT = &HFB0A09        ' (DWORD) dest = DPSnoo
  PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
  DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)
  BLACKNESS = &H42 ' (DWORD) dest = BLACK
  WHITENESS = &HFF0062       ' (DWORD) dest = WHITE
End Enum

Public Declare Function BitBlt Lib "gdi32" _
        (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, _
        ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long


