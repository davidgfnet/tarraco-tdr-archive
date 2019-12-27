VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FDIBPointer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const BMP_MAGIC_COOKIE As Integer = 19778
Private Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

Private Type BITMAPINFOHEADER
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Private m_memBits() As Byte
Private m_memBitmapInfo() As Byte
Private m_bih As BITMAPINFOHEADER
Private m_bfh As BITMAPFILEHEADER

Public Sub CreateFromPackedDIBPointer(ByRef pDIB As Long)
Call CopyMemory(ByVal VarPtr(m_bih.biSize), ByVal pDIB, Len(m_bih))

If m_bih.biSizeImage < 1 Then Exit Sub
ReDim m_memBits(0 To m_bih.biSizeImage - 1)
Call CopyMemory(m_memBits(0), ByVal pDIB + 40, m_bih.biSizeImage)
End Sub

Public Sub CreateTexture(texture As Direct3DTexture8)
Set texture = Direct3DX.CreateTexture(Device, m_bih.biWidth, m_bih.biHeight, 1, 0, D3DFMT_X8R8G8B8, D3DPOOL_MANAGED)

Dim Rect1 As RECT, Rect2 As D3DLOCKED_RECT
Dim Salidas() As Byte

Rect1.left = 0: Rect1.top = 0: Rect1.bottom = m_bih.biHeight: Rect1.right = m_bih.biWidth
texture.LockRect 0, Rect2, Rect1, 0

ReDim Salidas(Rect2.Pitch * 4 * m_bih.biHeight)

formalisetex m_memBits(0), m_bih.biWidth, m_bih.biHeight, Rect2.Pitch, Salidas(0)

DXCopyMemory ByVal Rect2.pBits, Salidas(0), Rect2.Pitch * 4 * m_bih.biHeight
 
texture.UnlockRect 0
End Sub
