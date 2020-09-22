VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Treeview background Example"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.TreeView TreeView1 
      Height          =   3015
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5318
      _Version        =   327682
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "This treeview controls should have a picture."
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.Image Img 
      Height          =   2880
      Left            =   4680
      Picture         =   "Form1.frx":000C
      Top             =   360
      Visible         =   0   'False
      Width           =   2880
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type PAINTSTRUCT
    hDC As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved As Byte
End Type

Private Declare Function BeginPaint Lib "user32" _
    (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" _
    (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" _
    (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" _
    (ByVal hDC As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" _
    (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" (ByVal hWnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Private Declare Function BitBlt Lib "gdi32" _
    (ByVal hDestDC As Long, ByVal x As Long, _
    ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" _
    (ByVal hObject As Long) As Long
Private Declare Function InvalidateRect Lib "user32" _
    (ByVal hWnd As Long, ByVal lpRect As Long, _
    ByVal bErase As Long) As Long

Private Const WM_PAINT = &HF
Private Const WM_ERASEBKGND = &H14
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115
Private Const WM_MOUSEWHEEL = &H20A


Private Sub Form_Load()

'Subclass the TreeView to trap messages
'that we'll need to respond to
Subclass Me, TreeView1

Dim Root As Node

'Add some items
With TreeView1.Nodes
Set Root = .Add(, , , "Top-level Node #1")
.Add Root.Index, tvwChild, , "Child Node #1"
.Add Root.Index, tvwChild, , "Child Node #2"
.Add Root.Index, tvwChild, , "Child Node #3"
Set Root = .Add(, , , "Top-level Node #2")
.Add Root.Index, tvwChild, , "Child Node #1"
.Add Root.Index, tvwChild, , "Child Node #2"
.Add Root.Index, tvwChild, , "Child Node #3"
Set Root = .Add(, , , "Top-level Node #3")
.Add Root.Index, tvwChild, , "Child Node #1"
.Add Root.Index, tvwChild, , "Child Node #2"
.Add Root.Index, tvwChild, , "Child Node #3"
Set Root = .Add(, , , "Top-level Node #4")
.Add Root.Index, tvwChild, , "Child Node #1"
.Add Root.Index, tvwChild, , "Child Node #2"
.Add Root.Index, tvwChild, , "Child Node #3"
End With

End Sub

Public Sub TreeViewMessage(ByVal hWnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long, RetVal As Long, _
    UseRetVal As Boolean)

'Prevent recursion with this variable
Static InProc As Boolean

Dim ps As PAINTSTRUCT
Dim TVDC As Long, drawDC1 As Long, drawDC2 As Long
Dim oldBMP1 As Long, drawBMP1 As Long
Dim oldBMP2 As Long, drawBMP2 As Long
Dim x As Long, y As Long, w As Long, h As Long
Dim TVWidth As Long, TVHeight As Long

If wMsg = WM_PAINT Then
    If InProc = True Then
        Exit Sub
    End If
    InProc = True
    'Prepare some variables we'll use
    TVWidth = TreeView1.width \ Screen.TwipsPerPixelX
    TVHeight = TreeView1.Height \ Screen.TwipsPerPixelY

    w = ScaleX(Img.Picture.width, vbHimetric, vbPixels)
    h = ScaleY(Img.Picture.Height, vbHimetric, vbPixels)

    'Begin painting. This API must be called in
    'response to the WM_PAINT message or you'll see
    'some odd visual effects :-)
    Call BeginPaint(hWnd, ps)
    TVDC = ps.hDC

    'Create a few canvases in memory to
    'draw on
    drawDC1 = CreateCompatibleDC(TVDC)
    drawBMP1 = CreateCompatibleBitmap(TVDC, TVWidth, TVHeight)
    oldBMP1 = SelectObject(drawDC1, drawBMP1)

    drawDC2 = CreateCompatibleDC(TVDC)
    drawBMP2 = CreateCompatibleBitmap(TVDC, TVWidth, TVHeight)
    oldBMP2 = SelectObject(drawDC2, drawBMP2)

    'This actually causes the TreeView to paint
    'itself onto our memory DC!
    SendMessage hWnd, WM_PAINT, drawDC1, ByVal 0&
    'Tile the bitmap and draw the TreeView
    'over it transparently
    For y = 0 To TVHeight Step h
        For x = 0 To TVWidth Step w
            PaintNormalStdPic drawDC2, x, y, w, h, _
                Img.Picture, 0, 0
        Next
    Next
    PaintTransparentDC drawDC2, 0, 0, TVWidth, TVHeight, _
        drawDC1, 0, 0, TranslateColor(vbWindowBackground)
    'Draw to the target DC
    BitBlt TVDC, 0, 0, TVWidth, TVHeight, _
        drawDC2, 0, 0, vbSrcCopy

    'Cleanup
    SelectObject drawDC1, oldBMP1
    SelectObject drawDC2, oldBMP2
    DeleteObject drawBMP1
    DeleteObject drawBMP2

    EndPaint hWnd, ps

    RetVal = 0
    UseRetVal = True
    InProc = False

ElseIf wMsg = WM_ERASEBKGND Then
    'Return TRUE
    RetVal = 1
    UseRetVal = True

ElseIf wMsg = WM_HSCROLL Or wMsg = WM_VSCROLL Or wMsg = WM_MOUSEWHEEL Then
    'Force a repaint to keep the bitmap
    'tiles lined up
    InvalidateRect hWnd, 0, 0

End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
MsgBox "If you liked it then please vote for me.", , "Treeview Example"
End Sub

Private Sub Form_Unload(Cancel As Integer)

'Kill subclassing routine for exit
UnSubclass TreeView1

End Sub


