VERSION 5.00
Begin VB.Form FMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CScroller DEMO"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_Speed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "10"
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton cmd_Play 
      Caption         =   "Play"
      Height          =   255
      Left            =   6600
      TabIndex        =   2
      Top             =   3000
      Width           =   735
   End
   Begin VB.PictureBox pic_Src 
      Height          =   255
      Left            =   1080
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   2760
      Width           =   255
   End
   Begin VB.PictureBox pic_Dest 
      BackColor       =   &H00000000&
      Height          =   2415
      Left            =   1080
      ScaleHeight     =   2355
      ScaleWidth      =   6315
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
   Begin VB.HScrollBar hs_Speed 
      Height          =   255
      LargeChange     =   5
      Left            =   5160
      Max             =   30
      TabIndex        =   3
      Top             =   3120
      Value           =   10
      Width           =   1335
   End
   Begin VB.Image img_Venus 
      Height          =   870
      Left            =   120
      Picture         =   "FMain.frx":0000
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image img_Triskelion 
      Height          =   615
      Left            =   240
      Picture         =   "FMain.frx":045D
      Top             =   120
      Width           =   675
   End
   Begin VB.Image img_Fish 
      Height          =   705
      Left            =   0
      Picture         =   "FMain.frx":0859
      Top             =   840
      Width           =   1035
   End
   Begin VB.Label lbl_Website 
      Caption         =   "http://unformed.hypermart.net"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "unformed@hotmail.com"
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label lbl_Speed 
      Caption         =   "scrolling speed:"
      Height          =   255
      Left            =   5160
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim scroller As New CScroller
Private Sub cmd_Play_Click()
If scroller.Active Then scroller.Cancel

With scroller
    'attach working pictureboxes
    'may also be attached by:
    '  Set .PicDest = pic_Dest
    '  Set .PicSrc = pic_Src
  
    .AttachPicDest pic_Dest
    .AttachPicSrc pic_Src
    
    'set defaults
    .DefaultFontSize = 10
    .DefaultTextColor = vbRed
    .DefaultShadowColor = &H808080
    .DefaultShadowLength = 0
    
    .ClearLines
    .AddLineText "CScrollerDEMO", , "Times New Roman", 14, True, ShadowLength:=3
    .AddLineText "originally written by Steve from vbtutor.com", FontSize:=8
    .AddLineText "further enhanced by milan ramaiya / decadence of evolution", FontSize:=8
    .AddLineText "@ http://unformed.hypermart.net", FontSize:=8
    .AddLineText
    .AddLineText "powerful text-scrolling class"
    .AddLineText "supports all different types of text:"
    
    .DefaultTextColor = vbBlue
    .AddLineText "BOLD", FontBold:=eTRUE
    .AddLineText "ITALIC", FontItalic:=eTRUE
    .AddLineText "UNDERLINE", FontUnderline:=eTRUE
    .AddLineText "SMALL", FontSize:=6
    .AddLineText "MEDIUM", FontSize:=12
    .AddLineText "LARGE", FontSize:=20
    
    .DefaultTextColor = vbWhite
    .AddLineText "and anywhere in between"
    .AddLineText "and in any color!"
    
    .AddLineText
    .AddLineText "there's different alignments:", vbRed
    .DefaultTextColor = vbBlue
    .AddLineText "LEFT-ALIGNED", HAlign:=eHAlign_Left
    .AddLineText "CENTER-ALIGNED", HAlign:=eHAlign_Center
    .AddLineText "RIGHT-ALIGNED", HAlign:=eHAlign_Right
    
    .AddLineText
    .AddLineText "and different fonts:", vbRed
    .DefaultTextColor = vbBlue
    .AddLineText Screen.Fonts(1), FontName:=Screen.Fonts(1)
    .AddLineText Screen.Fonts(2), FontName:=Screen.Fonts(2)
    .AddLineText Screen.Fonts(3), FontName:=Screen.Fonts(3)
        
    .AddLineText
    .AddLineText "AND even more!!!", vbRed
    .AddLineText "Different shadow colors:", vbRed
    .AddLineText "RED", ShadowColor:=vbRed, ShadowLength:=3
    .AddLineText "YELLOW", ShadowColor:=vbYellow, ShadowLength:=3
    .AddLineText "&H2456AF&", ShadowColor:=&H2456AF, ShadowLength:=3
    
    .AddLineText
    .AddLineText "Different shadow lengths:", vbRed
    .AddLineText "0", ShadowLength:=0
    .AddLineText "3", ShadowLength:=3
    .AddLineText "10", ShadowLength:=10
    
    .AddLineText
    .AddLineText "and directions:", vbRed
    .AddLineText "LEFT", ShadowLength:=5
    .AddLineText "RIGHT", ShadowLength:=5

    .AddLineText
    .AddLineText "and even custom images, at any alignment!"
    .AddLineImage img_Triskelion.Picture, eHAlign_Left
    .AddLineImage img_Fish.Picture, eHAlign_Center
    .AddLineImage img_Venus.Picture, eHAlign_Right
    
    .AddLineText
    .AddLineText "You can even set the speed (see box below)"
    
    .AddLineText "(c) christmas eve 2000 decadence of evolution", , , 8
    .AddLineText "http://unformed.hypermart.net", , , 8
    .AddLineText "unformed@hotmail.com", , , 8
    
    .Render txt_Speed.Text
End With

End Sub
Private Sub Form_Unload(Cancel As Integer)
If scroller.Active Then scroller.Cancel
End
End Sub


Private Sub hs_Speed_Change()
txt_Speed.Text = hs_Speed.Value
End Sub

