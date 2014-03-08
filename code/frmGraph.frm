VERSION 5.00
Begin VB.Form frmGraph 
   Caption         =   "Graph"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   9240
   Begin VB.VScrollBar VScroll1 
      Height          =   5625
      LargeChange     =   50
      Left            =   8595
      Min             =   -32767
      SmallChange     =   10
      TabIndex        =   2
      Top             =   180
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   50
      Left            =   3480
      Min             =   -32767
      SmallChange     =   10
      TabIndex        =   1
      Top             =   5760
      Width           =   4770
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000008&
      Height          =   4875
      Left            =   3300
      ScaleHeight     =   4815
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   780
      Width           =   5175
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lPictureLeftAdjust As Long
Dim lPictureTopAdjust As Long

Dim lXMidPnt As Long
Dim lYMidPnt As Long


Sub prcVars()
    lPictureLeftAdjust = 150
    lPictureTopAdjust = 15
End Sub

Private Sub Form_Load()
    prcVars
End Sub

Private Sub Form_Resize()

On Error Resume Next
  Me.Show
  Picture1.Enabled = False
  Picture1.Left = lPictureLeftAdjust

  Picture1.Width = frmGraph.ScaleWidth - lPictureLeftAdjust
  Picture1.top = 0
  Picture1.Height = frmGraph.ScaleHeight

  HScroll1.top = frmGraph.ScaleHeight - lPictureTopAdjust
  HScroll1.Left = lPictureLeftAdjust
  HScroll1.Width = frmGraph.ScaleWidth - lPictureLeftAdjust

  VScroll1.top = 0
  VScroll1.Left = frmGraph.ScaleWidth - lPictureTopAdjust
  VScroll1.Height = frmGraph.ScaleHeight - lPictureTopAdjust

  lXMidPnt = (Picture1.ScaleWidth / 2) + HScroll1.Value
  lYMidPnt = (Picture1.ScaleHeight / 2) + VScroll1.Value

  'ste = val(Text1.Text)
  Picture1.Cls
  
  'horizonal line
  Picture1.Line (0, lYMidPnt)-(Picture1.ScaleWidth, lYMidPnt), RGB(255, 255, 0)
  'vertical line
  Picture1.Line (lXMidPnt, 0)-(lXMidPnt, Picture1.ScaleHeight), RGB(255, 255, 0)

  frmGraph.SetFocus
  
  '
  'Call drawit
  '
  
  Picture1.Enabled = True
  
  
End Sub
