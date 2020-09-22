VERSION 5.00
Begin VB.Form fMain 
   Caption         =   "Ordered dither test"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10980
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   569
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   732
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPalette 
      Caption         =   "Import palette"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   2970
      Begin VB.TextBox txtOptimalColors 
         Height          =   300
         Left            =   975
         MaxLength       =   3
         TabIndex        =   6
         Top             =   1575
         Width           =   900
      End
      Begin VB.ComboBox cbHalftoneLevels 
         Height          =   315
         ItemData        =   "fMain.frx":0000
         Left            =   975
         List            =   "fMain.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   690
         Width           =   1590
      End
      Begin VB.OptionButton optPalette 
         Caption         =   "Optimal"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Top             =   1230
         Width           =   1530
      End
      Begin VB.OptionButton optPalette 
         Caption         =   "Halftone"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   405
         Value           =   -1  'True
         Width           =   1530
      End
      Begin VB.CheckBox chkWeightChannels 
         Caption         =   "Weight channels"
         Height          =   240
         Left            =   420
         TabIndex        =   7
         Top             =   2025
         Width           =   2250
      End
      Begin VB.Label Label1 
         Caption         =   "Levels"
         Height          =   225
         Left            =   420
         TabIndex        =   2
         Top             =   750
         Width           =   510
      End
      Begin VB.Label lblColors 
         Caption         =   "Colors"
         Height          =   225
         Left            =   420
         TabIndex        =   5
         Top             =   1620
         Width           =   1125
      End
   End
   Begin VB.Frame fraSaving 
      Caption         =   "Saving"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   90
      TabIndex        =   14
      Top             =   5445
      Width           =   2970
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save 8-bpp DIB"
         Enabled         =   0   'False
         Height          =   525
         Left            =   750
         TabIndex        =   16
         Top             =   885
         Width           =   1395
      End
      Begin VB.CommandButton cmdSaveGIF 
         Caption         =   "Save GIF"
         Enabled         =   0   'False
         Height          =   525
         Left            =   750
         TabIndex        =   17
         Top             =   1515
         Width           =   1395
      End
      Begin VB.CheckBox chkOptimizeGIFPalette 
         Caption         =   "Optimize GIF palette"
         Height          =   390
         Left            =   150
         TabIndex        =   15
         Top             =   345
         Width           =   2025
      End
      Begin VB.Label lblDIBsize 
         Caption         =   "Last saved DIB size:"
         Height          =   360
         Left            =   180
         TabIndex        =   18
         Top             =   2280
         Width           =   2760
      End
      Begin VB.Label lblGIFsize 
         Caption         =   "Last saved GIF size:"
         Height          =   330
         Left            =   180
         TabIndex        =   19
         Top             =   2595
         Width           =   2760
      End
   End
   Begin VB.Frame fraDither 
      Caption         =   "Dithering"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      Left            =   90
      TabIndex        =   8
      Top             =   2655
      Width           =   2970
      Begin VB.CheckBox chkPreserveExactColors 
         Caption         =   "Preserve exact colors"
         Enabled         =   0   'False
         Height          =   270
         Left            =   150
         TabIndex        =   10
         Top             =   735
         Width           =   2460
      End
      Begin VB.CommandButton cmdDither 
         Caption         =   "&Dither"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   525
         Left            =   750
         TabIndex        =   11
         Top             =   1245
         Width           =   1395
      End
      Begin VB.CheckBox chkOrderedDither 
         Caption         =   "Ordered dither"
         Height          =   330
         Left            =   150
         TabIndex        =   9
         Top             =   345
         Width           =   1770
      End
      Begin VB.Label lblRemapTime 
         Caption         =   "Remap time:"
         Height          =   300
         Left            =   180
         TabIndex        =   13
         Top             =   2325
         Width           =   2610
      End
      Begin VB.Label lblPaletteExtractionTime 
         Caption         =   "Palette extraction time:"
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   2025
         Width           =   2610
      End
   End
   Begin OrderedDither2.ucCanvas ucCanvas 
      Height          =   3375
      Left            =   3180
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   225
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   5953
   End
   Begin VB.Menu mnuLoad 
      Caption         =   "&Load"
   End
   Begin VB.Menu mnuPaste 
      Caption         =   "&Paste"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
' Project:       Ordered Dither 2
' Author:        Carles P.V.
' Last revision: 2003.11.15
'
' Special thanks to Ron van Tilburg for the fantastic
' mGIFSave module.
'====================================================
'
' This is a test. Play as you want with code, specialy
' routines where dither weights are calculated.
'
' Notes: Contrary to Halftone palettes, optimal extracted
'        palettes don't cover homogeneously all color space,
'        so resulting dithered color patterns will not
'        always become valid.
'
'        Another issue is optimal palette extraction itself.
'        Here, it has been simplified by reducing source
'        image size, and pre-reducing color space to a 4096
'        one (16 levels) before apply dithering.
'
'        Suggestions will be welcome.



Option Explicit

Private Declare Function timeGetTime Lib "winmm" () As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private m_oDIB32 As cDIB     ' Buffer DIB (32-bpp)
Private m_oDIB08 As cDIB     ' Target remaped DIB (8-bpp)
Private m_oPal08 As cPal8bpp ' Palette

Private m_LastFile As String ' Last file path



Private Sub Form_Load()
    
    Set m_oDIB32 = New cDIB
    Set m_oDIB08 = New cDIB
    Set m_oPal08 = New cPal8bpp
    
    mGIFSave.InitMasks
    
    cbHalftoneLevels.ListIndex = 4
    txtOptimalColors.Text = "256"
    
    mWheel.HookWheel ' Comment this call if you want to play with code.
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    ucCanvas.Move 210, 15, ScaleWidth - 215, ScaleHeight - 20
    On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set fMain = Nothing
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

'//

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyAdd:      PostMessage Me.hWnd, WM_MOUSEWHEEL, 1, 0
        Case vbKeySubtract: PostMessage Me.hWnd, WM_MOUSEWHEEL, 0, 0
    End Select
End Sub

'//

Private Sub mnuLoad_Click()

  Dim sTmpFilename As String
  
    '-- Show open file dialog
    sTmpFilename = mDialogFile.GetFileName(Me.hWnd, m_LastFile, "Supported files|*.bmp;*.jpg;*.gif", , "Load", -1)
    
    On Error GoTo ErrH
    
    If (Len(sTmpFilename)) Then
        m_LastFile = sTmpFilename
        
        '-- Load from file (-> 32bpp)
        DoEvents
        If (ucCanvas.DIB.CreateFromStdPicture(LoadPicture(sTmpFilename), True)) Then
            Call pvInitialize
        End If
    End If
    Exit Sub
    
ErrH:
    MsgBox "Unexpected error.", vbExclamation
End Sub

Private Sub mnuPaste_Click()

    On Error GoTo ErrH
    
    '-- Paste from clipboard (-> 32bpp)
    If (Clipboard.GetFormat(vbCFBitmap)) Then
        If (ucCanvas.DIB.CreateFromStdPicture(Clipboard.GetData(vbCFBitmap), True)) Then
            Call pvInitialize
        End If
    End If
    Exit Sub
    
ErrH:
    MsgBox "Unexpected error.", vbExclamation
End Sub

Private Sub chkOrderedDither_Click()
    chkPreserveExactColors.Enabled = -chkOrderedDither
End Sub

Private Sub cmdDither_Click()

  Dim t    As Long
  Dim oDIB As cDIB
  
    Screen.MousePointer = vbHourglass
    
    t = timeGetTime
    
    '-- Build palette
    If (optPalette(1)) Then
    
        If (Not IsNumeric(txtOptimalColors.Text) Or _
           (Val(txtOptimalColors.Text) < 8 Or Val(txtOptimalColors.Text) > 256)) Then
            Screen.MousePointer = vbDefault
            MsgBox "Enter a valid number of colors [8-256]", vbExclamation
            With txtOptimalColors
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
            Exit Sub
        End If
        
        '-- Temp. 32-bpp DIB (reduced) for extracting palette from (speed up) [*]
        '-- [*]: Use m_oDIB32 for extracting palette of full-size source DIB (better/slower)
        Set oDIB = New cDIB
        m_oDIB32.CloneTo oDIB: Call pvFitDIB(oDIB, 150, 150)
        m_oPal08.CreateOptimal oDIB, txtOptimalColors.Text, 8, IIf(chkWeightChannels, 0.36, 1), IIf(chkWeightChannels, 0.436, 1), IIf(chkWeightChannels, 0.341, 1)
    
      Else
        m_oPal08.CreateHalftone (cbHalftoneLevels.ListIndex)
    End If
    
    '-- Palette extraction time
    lblPaletteExtractionTime.Caption = "Palette extraction time: " & Format$(timeGetTime - t, "0 ms")
    
    t = timeGetTime
    
    '-- Initialize LUTs, set dither mode and remap
    If (optPalette(1)) Then
    
        mRemap8bpp.Initialize_ODM_O_LUT m_oPal08.Entries
        mRemap8bpp.Build_RGB4096InvIdx_LUT m_oPal08
        mRemap8bpp.ImportPalette = [ipOptimal]
        mRemap8bpp.DitherMethod = IIf(chkOrderedDither, [dmOrdered], [dmNone])
        mRemap8bpp.PreserveExactColors = -chkPreserveExactColors
        mRemap8bpp.Remap m_oDIB32, m_oDIB08, m_oPal08

      Else
        mRemap8bpp.Initialize_ODM_H_LUT (cbHalftoneLevels.ListIndex + 2)
        mRemap8bpp.ImportPalette = [ipHalftone]
        mRemap8bpp.DitherMethod = IIf(chkOrderedDither, [dmOrdered], [dmNone])
        mRemap8bpp.PreserveExactColors = -chkPreserveExactColors
        mRemap8bpp.Remap m_oDIB32, m_oDIB08, m_oPal08
    End If
    
    '-- Remap time
    lblRemapTime.Caption = "Remap time: " & Format$(timeGetTime - t, "0 ms")
    
    '-- Paint our 8-bpp DIB to canvas 32-bpp DIB (speed up painting/zooming)
    m_oDIB08.Stretch ucCanvas.DIB.hDIBDC, 0, 0, ucCanvas.DIB.Width, ucCanvas.DIB.Height
    ucCanvas.Repaint
    
    cmdSave.Enabled = True
    cmdSaveGIF.Enabled = True
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSave_Click()
    
    Screen.MousePointer = vbHourglass

    m_oDIB08.Save pvAppPath & "Test.bmp"
    lblDIBsize.Caption = "Last saved DIB size: " & Format$(FileLen(pvAppPath & "Test.bmp"), "#,# bytes")
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSaveGIF_Click()

    Screen.MousePointer = vbHourglass
    
    If (chkOptimizeGIFPalette) Then
        mRemap8bpp.OptimizePalette m_oDIB08, m_oPal08
    End If
    mGIFSave.SaveGIF pvAppPath & "Test.gif", m_oDIB08, m_oPal08
    lblGIFsize.Caption = "Last saved GIF size: " & Format$(FileLen(pvAppPath & "Test.gif"), "#,# bytes")
    
    Screen.MousePointer = vbDefault
End Sub

'//

Private Sub pvInitialize()

    '-- Initialize/refresh canvas
    ucCanvas.Resize
    ucCanvas.Repaint
    '-- Temp. 32-bpp buffer DIB
    ucCanvas.DIB.CloneTo m_oDIB32
    '-- Target 8-bpp DIB
    m_oDIB08.Create ucCanvas.DIB.Width, ucCanvas.DIB.Height, [08_bpp]
    
    '-- Update command buttons
    cmdDither.Enabled = True
    cmdSave.Enabled = False
    cmdSaveGIF.Enabled = False
End Sub

Private Sub pvFitDIB(oDIB As cDIB, ByVal maxWidth As Long, maxHeight As Long)
  
  Dim bfW As Long, bfH As Long
  Dim bfx As Long, bfy As Long
    
    oDIB.GetBestFitInfo maxWidth, maxHeight, bfx, bfy, bfW, bfH
    oDIB.Resize bfW, bfH
End Sub

Private Function pvAppPath() As String
    pvAppPath = App.Path & IIf(Right$(App.Path, 1) <> "\", "\", vbNullString)
End Function
