VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spinning Image"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   2640
   ScaleWidth      =   3495
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CMN 
      Left            =   360
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Select Image"
      Filter          =   "Image Files |*.gif*;*.jpg*;*.bmp*"
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "None!"
      Top             =   1920
      Width           =   2175
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   100
      Left            =   840
      Max             =   499
      SmallChange     =   10
      TabIndex        =   2
      Top             =   1560
      Value           =   100
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2040
      Top             =   960
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   0
      Picture         =   "Form1.frx":D382
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   30
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox Board 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   1
      Top             =   0
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Image:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Dim W, WS

Private Sub Command1_Click()
CMN.ShowOpen
If CMN.FileName = "" Then Exit Sub
Text1.Text = CMN.FileName
On Error Resume Next
Picture1.Picture = LoadPicture(Text1.Text)
WS = -5
W = Picture1.ScaleWidth
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub HScroll1_Change()
Timer1.Interval = 500 - HScroll1.Value
End Sub

Private Sub Timer1_Timer()
Board.Cls
W = W + WS
If W < -Picture1.ScaleWidth Then WS = 5
If W > Picture1.ScaleWidth Then WS = -5

StretchBlt Board.hdc, Board.ScaleWidth \ 2 - W \ 2, Board.ScaleHeight \ 2 - Picture1.ScaleHeight \ 2, W, Picture1.ScaleHeight, Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, vbSrcCopy
End Sub
