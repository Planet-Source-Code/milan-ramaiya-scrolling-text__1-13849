VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CScroller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' initially developed by Steve from vbtutor.com
' further enhancement by milan ramaiya of decadence of evolution.com
' usage:
'   set picDest = actual scrolling frame
'   set picSrc = any picturebox (will be hidden)
'   set Default... properties
'   add text using AddLine
'   Call Render()
'
'  updated 12.24.00

Option Explicit

Private m_Active As Boolean
Private m_Cancel As Boolean
Private m_PicSrc As PictureBox
Private m_PicDest As PictureBox
Private m_DefaultTextColor As Long
Private m_DefaultFontName As String
Private m_DefaultFontSize As Long
Private m_DefaultFontBold As Boolean
Private m_DefaultFontItalic As Boolean
Private m_DefaultFontUnderline As Boolean
Private m_DefaultHAlign As ESCR_HALIGN
Private m_DefaultShadowColor As Long
Private m_DefaultShadowDirection As ESCR_SHADOWDIRECTION
Private m_DefaultShadowLength As Long

Public Enum ESCR_BOOLEAN
    eTRUE = True
    eFALSE = False
    eDefault
End Enum

Public Enum ESCR_HALIGN
    eHAlign_Left
    eHAlign_Right
    eHAlign_Center
End Enum

Public Enum ESCR_SHADOWDIRECTION
'    eUp = 1
'    eDown = 2      'down currently causes problems
    eSD_Left = 4
    eSD_Right = 8
End Enum

Private Enum ESCR_LINETYPE
    eLT_Text
    eLT_Image
End Enum

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Const SRCCOPY = &HCC0020

Private asText() As SCR_TextLine

Private Type SCR_TextLine
    LineType As ESCR_LINETYPE
    Text As String
    Picture As StdPicture
    
    Color As Long
    FontName As String
    FontSize As Long
    FontBold As Boolean
    FontItalic As Boolean
    FontUnderline As Boolean
    
    HAlign As ESCR_HALIGN
    
    ShadowColor As Long
    ShadowDirection As ESCR_SHADOWDIRECTION
    ShadowLength As Long
End Type

Private Sub AddLine(TextLine As SCR_TextLine)
'add to array, check if non-itialized array
On Error Resume Next
Dim i As Long
i = 0
i = UBound(asText) + 1
ReDim Preserve asText(i)
asText(i) = TextLine
On Error GoTo 0
End Sub
Public Sub AddLineImage(Pic As StdPicture, Optional HAlign As ESCR_HALIGN = -1)
Dim TextLine As SCR_TextLine
With TextLine
    .LineType = eLT_Image
    Set .Picture = Pic
    
    If HAlign = -1 Then
        .HAlign = DefaultHAlign
    Else
        .HAlign = HAlign
    End If
End With

AddLine TextLine
End Sub
Public Sub ClearLines()
Dim EmptyArray() As SCR_TextLine
asText() = EmptyArray()
End Sub

Public Property Get Active() As Boolean
Active = m_Active
End Property


Public Sub AddLineText(Optional Text As String = "", _
                Optional Color As Long = -1, _
                Optional FontName As String = "", _
                Optional FontSize As Long = -1, _
                Optional FontBold As ESCR_BOOLEAN = eDefault, _
                Optional FontItalic As ESCR_BOOLEAN = eDefault, _
                Optional FontUnderline As ESCR_BOOLEAN = eDefault, _
                Optional HAlign As ESCR_HALIGN = -1, _
                Optional ShadowColor As Long = -1, _
                Optional ShadowDirection As ESCR_SHADOWDIRECTION = -1, _
                Optional ShadowLength As Long = -1)

Dim TextLine As SCR_TextLine
With TextLine
    .Text = Text
    If Color = -1 Then
        .Color = m_DefaultTextColor
    Else
        .Color = Color
    End If
    
    If HAlign = -1 Then
        .HAlign = m_DefaultHAlign
    Else
        .HAlign = HAlign
    End If
    
    If FontName = "" Then
        .FontName = m_DefaultFontName
    Else
        .FontName = FontName
    End If
    
    If FontSize = -1 Then
        .FontSize = m_DefaultFontSize
    Else
        .FontSize = FontSize
    End If
    
    If FontBold = eDefault Then
        .FontBold = m_DefaultFontBold
    Else
        .FontBold = FontBold
    End If
    
    If FontItalic = eDefault Then
        .FontItalic = m_DefaultFontItalic
    Else
        .FontItalic = FontItalic
    End If
    
    If FontUnderline = eDefault Then
        .FontUnderline = m_DefaultFontUnderline
    Else
        .FontUnderline = FontUnderline
    End If
    
    If ShadowColor = -1 Then
        .ShadowColor = m_DefaultShadowColor
    Else
        .ShadowColor = ShadowColor
    End If
    
    If ShadowDirection = -1 Then
        .ShadowDirection = m_DefaultShadowDirection
    Else
        .ShadowDirection = ShadowDirection
    End If
    
    If ShadowLength = -1 Then
        .ShadowLength = m_DefaultShadowLength
    Else
        .ShadowLength = ShadowLength
    End If
End With

AddLine TextLine

End Sub
Public Sub AttachPicSrc(PicSrc As PictureBox)
Set m_PicSrc = PicSrc
End Sub
Public Sub AttachPicDest(PicDest As PictureBox)
Set m_PicDest = PicDest
End Sub

Public Sub Cancel()
m_Active = False
End Sub
Private Sub DrawImage(TextLine As SCR_TextLine)

Dim Picture As StdPicture, wid As Single, hgt As Single
Set Picture = TextLine.Picture

With m_PicSrc

    wid = .ScaleX(Picture.Width, vbHimetric, vbPixels)
    hgt = .ScaleX(Picture.Height, vbHimetric, vbPixels)

    'set alignment
    Select Case TextLine.HAlign
        Case eHAlign_Center
            .CurrentX = (.ScaleWidth - wid) / 2
        Case eHAlign_Left
            .CurrentX = 0
        Case eHAlign_Right
            .CurrentX = .ScaleWidth - wid
    End Select

    .PaintPicture Picture, .CurrentX, .CurrentY

    .CurrentY = .CurrentY + hgt

End With
End Sub
Private Sub DrawLine(TextLine As SCR_TextLine)
Select Case TextLine.LineType
    Case eLT_Text
        DrawText TextLine
    Case eLT_Image
        DrawImage TextLine
End Select
End Sub

Private Sub DrawText(TextLine As SCR_TextLine)
Dim Text As String
Text = TextLine.Text

'shadow lengths
Dim dx As Integer, dy As Integer

dy = 1
'shadow is drawn first, so numbers are opposite
If TextLine.ShadowDirection And eSD_Left Then dx = 1
If TextLine.ShadowDirection And eSD_Right Then dx = -1

With m_PicSrc
    'set font
    .FontName = TextLine.FontName
    .FontSize = TextLine.FontSize
    .FontBold = TextLine.FontBold
    .FontItalic = TextLine.FontItalic
    
    'set alignment
    Select Case TextLine.HAlign
        Case eHAlign_Center
            .CurrentX = (.ScaleWidth / 2) - (.TextWidth(Text) / 2)
        Case eHAlign_Left
            .CurrentX = 0
        Case eHAlign_Right
            .CurrentX = .ScaleWidth - .TextWidth(Text)
    End Select
    
    'draw shadow
    Dim x As Single, y As Single, i As Long
    .ForeColor = TextLine.ShadowColor
    x = .CurrentX
    y = .CurrentY
    For i = 1 To TextLine.ShadowLength
        m_PicSrc.Print Text
        x = x + dx
        y = y + dy
        .CurrentX = x
        .CurrentY = y
    Next i
    
    'draw text
    .ForeColor = TextLine.Color
    m_PicSrc.Print Text
End With
End Sub
Public Property Get DefaultTextColor() As Long
DefaultTextColor = m_DefaultTextColor
End Property
Public Property Let DefaultTextColor(TextColor As Long)
m_DefaultTextColor = TextColor
End Property

Public Property Get DefaultFontName() As String
DefaultFontName = m_DefaultFontName
End Property
Public Property Let DefaultFontName(FontName As String)
m_DefaultFontName = FontName
End Property

Public Property Get DefaultFontSize() As Long
DefaultFontSize = m_DefaultFontSize
End Property
Public Property Let DefaultFontSize(FontSize As Long)
m_DefaultFontSize = FontSize
End Property

Public Property Get DefaultFontBold() As Boolean
DefaultFontBold = m_DefaultFontBold
End Property
Public Property Let DefaultFontBold(FontBold As Boolean)
m_DefaultFontBold = FontBold
End Property

Public Property Get DefaultFontItalic() As Boolean
DefaultFontItalic = m_DefaultFontItalic
End Property
Public Property Let DefaultFontItalic(FontItalic As Boolean)
m_DefaultFontItalic = FontItalic
End Property

Public Property Get DefaultFontUnderline() As Boolean
DefaultFontUnderline = m_DefaultFontUnderline
End Property
Public Property Let DefaultFontUnderline(FontUnderline As Boolean)
m_DefaultFontUnderline = FontUnderline
End Property

Public Property Get DefaultHAlign() As Long
DefaultHAlign = m_DefaultHAlign
End Property
Public Property Let DefaultHAlign(HAlign As ESCR_HALIGN)
m_DefaultHAlign = HAlign
End Property

Public Property Get DefaultShadowColor() As Long
DefaultShadowColor = m_DefaultShadowColor
End Property
Public Property Let DefaultShadowColor(ShadowColor As Long)
m_DefaultShadowColor = ShadowColor
End Property

Public Property Get DefaultShadowDirection() As ESCR_SHADOWDIRECTION
DefaultShadowDirection = m_DefaultShadowDirection
End Property
Public Property Let DefaultShadowDirection(ShadowDirection As ESCR_SHADOWDIRECTION)
m_DefaultShadowDirection = ShadowDirection
End Property

Public Property Get DefaultShadowLength() As Long
DefaultShadowLength = m_DefaultShadowLength
End Property
Public Property Let DefaultShadowLength(ShadowLength As Long)
m_DefaultShadowLength = ShadowLength
End Property

Private Function GetLineHeight(TextLine As SCR_TextLine)
With m_PicSrc
    Select Case TextLine.LineType
        Case eLT_Text
            .FontName = TextLine.FontName
            .FontSize = TextLine.FontSize
            GetLineHeight = .TextHeight("Test")
        Case eLT_Image
            GetLineHeight = .ScaleY(TextLine.Picture.Height, vbHimetric, vbPixels)
    End Select
End With
End Function
Public Property Set PicSrc(PicSrc As PictureBox)
Set m_PicSrc = PicSrc
End Property
Public Property Set PicDest(PicDest As PictureBox)
Set m_PicDest = PicDest
End Property

Public Property Get PicSrc() As PictureBox
Set PicSrc = m_PicSrc
End Property
Public Property Get PicDest() As PictureBox
Set PicDest = m_PicDest
End Property

Public Sub Render(Optional FrameLapse As Long = 10)
'FrameLapse: milliseconds between frame

'current draw coordinates, hgt will be used doubly as temp storage for picbox height
Dim x As Long, y As Long, wid As Single, hgt As Single

m_Active = True

With m_PicSrc
    
    m_PicDest.Cls
    
    .Visible = False
    .Cls
    
    .AutoRedraw = True
    'for a correct output, match the destination surface
    .Width = m_PicDest.Width
    .BackColor = m_PicDest.BackColor
    
    'save the current scale modes, to be restored later
    Dim oldFormMode As Long, oldSrcMode As Long, oldDestMode As Long
    oldFormMode = m_PicSrc.Parent.ScaleMode
    oldSrcMode = m_PicSrc.ScaleMode
    oldDestMode = m_PicDest.ScaleMode
    m_PicSrc.ScaleMode = vbPixels
    m_PicDest.ScaleMode = vbPixels
    m_PicSrc.Parent.ScaleMode = vbPixels

    'get necessary height of picturebox, multiply by 2 for automatic clearing
    Dim i As Long
    For i = 0 To UBound(asText)
        hgt = hgt + GetLineHeight(asText(i))
    Next i
    m_PicSrc.Height = hgt * 2

    For i = 0 To UBound(asText)
        DrawLine asText(i)
    Next i

    'set initial x,y coordinates
    x = 0
    y = m_PicDest.ScaleHeight
    
    'set width, height of picSrc to blit from
    wid = m_PicSrc.ScaleWidth
    hgt = m_PicSrc.ScaleHeight
    
    m_PicSrc.ScaleMode = oldSrcMode
    m_PicDest.ScaleMode = oldDestMode
    m_PicSrc.Parent.ScaleMode = oldFormMode
    
    'continue until full image in invisible
    Dim tmeStart As Long
    Do Until y < -hgt
        tmeStart = timeGetTime

        i = BitBlt(m_PicDest.hDC, x, y, wid, hgt, m_PicSrc.hDC, 0, 0, SRCCOPY)
        y = y - 1
        
        'break if cancelled
        If m_Active = False Then Exit Do
        
        'lock at specified framerate
        Do Until timeGetTime - tmeStart >= FrameLapse
            DoEvents
        Loop
        
    Loop
End With

m_Active = False
End Sub
Private Sub Class_Initialize()
m_DefaultTextColor = vbBlack
m_DefaultFontName = "Arial"
m_DefaultFontSize = 12
m_DefaultFontBold = False
m_DefaultFontItalic = False
m_DefaultFontUnderline = False
m_DefaultHAlign = eHAlign_Center
m_DefaultShadowColor = &H808080
m_DefaultShadowDirection = eSD_Right
m_DefaultShadowLength = 3
End Sub
Private Sub Class_Terminate()
Set m_PicSrc = Nothing
Set m_PicDest = Nothing
End Sub
