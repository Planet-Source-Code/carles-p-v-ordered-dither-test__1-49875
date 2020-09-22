Attribute VB_Name = "mRemap8bpp"
'================================================
' Module:        mRemap8bpp.bas
' Author:        Carles P.V.
' Dependencies:  cDIB.cls
'                cPal8bpp.cls
' Last revision: 2003.11.15
'================================================

Option Explicit

'-- API:

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound   As Long
End Type

Private Type SAFEARRAY2D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    Bounds(1)  As SAFEARRAYBOUND
End Type

Private Type RGBQUAD
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type

Private Declare Function VarPtrArray Lib "msvbvm50" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, lpSrc As Any, ByVal Length As Long)

'//

'-- Public enums.:

Public Enum impDitherMethodCts
    [dmNone] = 0
    [dmOrdered]
End Enum
Public Enum impImportPaletteCts
    [ipOptimal] = 0
    [ipHalftone]
End Enum

'-- Property variables:

Private m_DitherMethod        As impDitherMethodCts
Private m_ImportPalette       As impImportPaletteCts
Private m_PreserveExactColors As Boolean

'-- Private variables:

Private m_tPal(&HFF) As RGBQUAD     '  8-bpp current palette entries
Private m_tSA32      As SAFEARRAY2D ' 32-bpp SA
Private m_Bits32()   As RGBQUAD     ' 32-bpp maped bits
Private m_tSA08      As SAFEARRAY2D '  8-bpp SA
Private m_Bits08()   As Byte        '  8-bpp maped bits
Private m_x          As Long
Private m_y          As Long
Private m_W          As Long
Private m_H          As Long

'//

Private m_ODM_O(7, 7)                As Long ' Ordered dither matrix (Bayer 8x8)
Private m_ODM_H(7, 7)                As Long ' Ordered dither matrix (Bayer 8x8)

Private m_RGB4096_Inv(&HF, &HF, &HF) As Byte ' RGB4096 palette inverse index LUT
Private m_RGB4096_Trn()              As Long ' RGB4096 translation LUT

Private m_HT_Inv()                   As Byte ' Halftone palette inverse index LUT
Private m_HT_Trn()                   As Long ' Halftone translation LUT




'========================================================================================
' LUTs initialization
'========================================================================================

Public Sub Initialize_ODM_O_LUT(ByVal Colors As Long)

  Dim aBPP As Byte   ' Color depth
  Dim sDW  As Single ' Dither weight
  Dim tArr As Variant
  Dim x    As Long
  Dim y    As Long
  
  Dim lIdx As Long
  Dim lOff As Long
  
    '-- Dither weight = f(bpp)
    Do: aBPP = aBPP + 1: Loop Until 2 ^ aBPP >= Colors
    sDW = (&H11 * (9 - aBPP)) / &H33

    '-- Ordered dither matrix (Bayer 8x8. [-25, 25])
    tArr = Array(0, 38, 9, 47, 2, 40, 11, 50, 25, 12, 35, 22, 27, 15, 37, 24, 6, 44, 3, 41, 8, 47, 5, 43, 31, 19, 28, 15, 34, 21, 31, 18, 1, 39, 11, 49, 0, 39, 10, 48, 27, 14, 36, 23, 26, 13, 35, 23, 7, 46, 4, 43, 7, 45, 3, 42, 33, 20, 30, 17, 32, 19, 29, 16)
    For x = 0 To 7
        For y = 0 To 7
            m_ODM_O(x, y) = sDW * (tArr(lIdx) - 25): lIdx = lIdx + 1
        Next y
    Next x
    
    '-- Prepare RGB4096 LUT
    lOff = 25 * sDW + 1
    ReDim m_RGB4096_Trn(-lOff To &HFF + lOff)
    
    '-- RGB4096 translation LUT
    For lIdx = -lOff To &HFF + lOff
        m_RGB4096_Trn(lIdx) = (lIdx + &H8) \ &H11
        If (m_RGB4096_Trn(lIdx) < &H0) Then m_RGB4096_Trn(lIdx) = &H0
        If (m_RGB4096_Trn(lIdx) > &HF) Then m_RGB4096_Trn(lIdx) = &HF
    Next lIdx
End Sub

Public Sub Initialize_ODM_H_LUT(ByVal Levels As Long)

  Dim sDW  As Single ' Dither weight
  Dim tArr As Variant
  Dim x    As Long
  Dim y    As Long
  
  Dim lLev As Long
  Dim lStp As Long, lStpDIV2 As Long
  Dim R    As Long
  Dim G    As Long
  Dim B    As Long
  
  Dim lIdx As Long
  Dim lOff As Long
    
    '-- Dither weight = f(levels)
    Select Case Levels
        Case 2: sDW = &H100 / &H33
        Case 3: sDW = &H80 / &H33
        Case 4: sDW = &H55 / &H33
        Case 5: sDW = &H40 / &H33
        Case 6: sDW = &H33 / &H33
    End Select
    
    '-- Ordered dither matrix (Bayer 8x8. [-25, 25])
    tArr = Array(0, 38, 9, 47, 2, 40, 11, 50, 25, 12, 35, 22, 27, 15, 37, 24, 6, 44, 3, 41, 8, 47, 5, 43, 31, 19, 28, 15, 34, 21, 31, 18, 1, 39, 11, 49, 0, 39, 10, 48, 27, 14, 36, 23, 26, 13, 35, 23, 7, 46, 4, 43, 7, 45, 3, 42, 33, 20, 30, 17, 32, 19, 29, 16)
    For x = 0 To 7
        For y = 0 To 7
            m_ODM_H(x, y) = sDW * (tArr(lIdx) - 25): lIdx = lIdx + 1
        Next y
    Next x
    
    '-- Prepare inverse index and translation LUTs
    lLev = Levels - 1
    lOff = 25 * sDW + 1
    lStp = 256 \ lLev
    lStpDIV2 = 128 \ lLev
    ReDim m_HT_Inv(lLev, lLev, lLev)
    ReDim m_HT_Trn(-lOff To &HFF + lOff)
    
    '-- Halftone inverse index LUT
    lIdx = 0
    For B = 0 To lLev
        For G = 0 To lLev
            For R = 0 To lLev
                '-- Set palette inverse index
                m_HT_Inv(R, G, B) = lIdx
                lIdx = lIdx + 1
            Next R
        Next G
    Next B
    
    '-- Halftone translation LUT
    For lIdx = -lOff To &HFF + lOff
        m_HT_Trn(lIdx) = (lIdx + lStpDIV2) \ lStp
        If (m_HT_Trn(lIdx) < 0) Then m_HT_Trn(lIdx) = 0
        If (m_HT_Trn(lIdx) > lLev) Then m_HT_Trn(lIdx) = lLev
    Next lIdx
End Sub

'========================================================================================
' Properties
'========================================================================================

Public Property Get DitherMethod() As impDitherMethodCts
    DitherMethod = m_DitherMethod
End Property
Public Property Let DitherMethod(ByVal New_DitherMethod As impDitherMethodCts)
    m_DitherMethod = New_DitherMethod
End Property

Public Property Get ImportPalette() As impImportPaletteCts
    ImportPalette = m_ImportPalette
End Property
Public Property Let ImportPalette(ByVal New_ImportPalette As impImportPaletteCts)
    m_ImportPalette = New_ImportPalette
End Property

Public Property Get PreserveExactColors() As Boolean
    PreserveExactColors = m_PreserveExactColors
End Property
Public Property Let PreserveExactColors(ByVal New_PreserveExactColors As Boolean)
    m_PreserveExactColors = New_PreserveExactColors
End Property

'========================================================================================
' Methods
'========================================================================================

Public Sub Remap(oDIB32 As cDIB, oDIB08 As cDIB, oPal08 As cPal8bpp)
  
  Dim aPal(1023) As Byte
  
    '-- Fill temp. palette copy (speed up)
    CopyMemory m_tPal(0), ByVal oPal08.lpPalette, 1024
    
    '-- Rebuild 8-bpp target DIB (create and set current palette)
    CopyMemory aPal(0), m_tPal(0), 1024
    oDIB08.SetPalette aPal()
    
    '-- Map source and target DIB bits (32-bpp/8-bpp)
    Call pvBuild_32bppSA(m_tSA32, oDIB32)
    Call pvBuild_08bppSA(m_tSA08, oDIB08)
    CopyMemory ByVal VarPtrArray(m_Bits32()), VarPtr(m_tSA32), 4
    CopyMemory ByVal VarPtrArray(m_Bits08()), VarPtr(m_tSA08), 4
   
    '-- Get dimensions
    m_W = oDIB32.Width - 1
    m_H = oDIB32.Height - 1
   
    '-- Dither...
    Select Case m_ImportPalette
        Case [ipOptimal]
            Select Case m_DitherMethod
                Case [dmNone]:    Call pvDitherToPalette
                Case [dmOrdered]: Call pvDitherToPalette_Ordered
            End Select
        Case [ipHalftone]
            Select Case m_DitherMethod
                Case [dmNone]:    Call pvDitherToHalftonePalette
                Case [dmOrdered]: Call pvDitherToHalftonePalette_Ordered
            End Select
    End Select
    
    '-- Unmap DIB bits
    CopyMemory ByVal VarPtrArray(m_Bits32()), 0&, 4
    CopyMemory ByVal VarPtrArray(m_Bits08()), 0&, 4
End Sub

Public Sub Build_RGB4096InvIdx_LUT(oPal08 As cPal8bpp)
  
  Dim R As Long
  Dim G As Long
  Dim B As Long

    '-- Build 4096-colors palette inverse indexes LUT
    For R = 0 To &HF
    For G = 0 To &HF
    For B = 0 To &HF
        oPal08.ClosestIndex R * &H11, G * &H11, B * &H11, m_RGB4096_Inv(R, G, B)
    Next B, G, R
End Sub

Public Function OptimizePalette(oDIB08 As cDIB, oPal8bpp As cPal8bpp) As Integer

  Dim nBefore As Integer
  Dim aPal()  As Byte
  Dim bUsed() As Boolean
  Dim bTrnE() As Byte
  Dim bInvE() As Byte
  Dim lIdx    As Long
  Dim lMax    As Long
  
    '-- Store current number of entries and initialize arrays
    nBefore = oPal8bpp.Entries
    ReDim bUsed(nBefore - 1)
    ReDim bTrnE(nBefore - 1)
    ReDim bInvE(nBefore - 1)

    '-- Map current 8-bpp DIB bits
    Call pvBuild_08bppSA(m_tSA08, oDIB08)
    CopyMemory ByVal VarPtrArray(m_Bits08()), VarPtr(m_tSA08), 4
   
    '-- Get dimensions
    m_W = oDIB08.Width - 1
    m_H = oDIB08.Height - 1
    
    '-- Check used entries...
    For m_y = 0 To m_H
        For m_x = 0 To m_W
            bUsed(m_Bits08(m_x, m_y)) = -1
        Next m_x
    Next m_y
    
    '-- 'Strecth' palette...
    For lIdx = 0 To oPal8bpp.Entries - 1
        If (bUsed(lIdx)) Then
            bTrnE(lMax) = lIdx ' New index
            bInvE(lIdx) = lMax ' Inverse index
            lMax = lMax + 1    ' Current count
        End If
    Next lIdx
    
    '-- Any entry removed [?]
    If (lMax < oPal8bpp.Entries) Then
    
        If (lMax > 1) Then lMax = lMax - 1
        '-- Build temp. palette with only used entries
        ReDim aPal(4 * (lMax + 1) - 1)
        With oPal8bpp
            For lIdx = 0 To lMax
                aPal(4 * lIdx + 0) = .rgbB(bTrnE(lIdx))
                aPal(4 * lIdx + 1) = .rgbG(bTrnE(lIdx))
                aPal(4 * lIdx + 2) = .rgbR(bTrnE(lIdx))
            Next lIdx
        End With
        
        '-- Rebuild current palette
        oPal8bpp.Initialize (lMax + 1)
        CopyMemory ByVal oPal8bpp.lpPalette, aPal(0), 4 * (lMax + 1)
        '-- and set as XOR DIB palette
        oDIB08.SetPalette aPal()
        
        '-- Set new indexes...
        For m_y = 0 To m_H
            For m_x = 0 To m_W
                m_Bits08(m_x, m_y) = bInvE(m_Bits08(m_x, m_y))
            Next m_x
        Next m_y
    End If

    '-- Unmap DIB bits
    CopyMemory ByVal VarPtrArray(m_Bits08()), 0&, 4
    
    '-- Return removed entries
    OptimizePalette = (nBefore - oPal8bpp.Entries)
End Function

'========================================================================================
' Private
'========================================================================================

Private Sub pvDitherToPalette()
    
    For m_y = 0 To m_H
        For m_x = 0 To m_W
            '-- Set 8-bpp palette index
            m_Bits08(m_x, m_y) = m_RGB4096_Inv(m_RGB4096_Trn(m_Bits32(m_x, m_y).R), m_RGB4096_Trn(m_Bits32(m_x, m_y).G), m_RGB4096_Trn(m_Bits32(m_x, m_y).B))
        Next m_x
    Next m_y
End Sub

Private Sub pvDitherToPalette_Ordered()

  Dim lODx   As Long
  Dim lODy   As Long
  Dim lODInc As Long
  Dim aInvID As Byte
  
    '-- Dither...
    If (m_PreserveExactColors) Then
    
        For m_y = 0 To m_H
            For m_x = 0 To m_W
    
                With m_Bits32(m_x, m_y)
    
                    '-- Inv. index
                    aInvID = m_RGB4096_Inv(m_RGB4096_Trn(.R), m_RGB4096_Trn(.G), m_RGB4096_Trn(.B))
                    
                    '-- Match to any palette color [?]
                    If (m_tPal(aInvID).R <> .R Or m_tPal(aInvID).G <> .G Or m_tPal(aInvID).B <> .B) Then
                        '-- Not: dither
                        lODInc = m_ODM_O(lODx, lODy)
                        m_Bits08(m_x, m_y) = m_RGB4096_Inv(m_RGB4096_Trn(.R + lODInc), m_RGB4096_Trn(.G + lODInc), m_RGB4096_Trn(.B + lODInc))
                      Else
                        '-- Yes: do not dither
                        m_Bits08(m_x, m_y) = aInvID
                    End If
                End With
    
                '-- Inc. ord. matrix column
                lODx = lODx + 1: If (lODx = 8) Then lODx = 0
            Next m_x
    
            '-- Inc. ord. matrix row
            lODx = 0
            lODy = lODy + 1: If (lODy = 8) Then lODy = 0
        Next m_y
        
      Else
        For m_y = 0 To m_H
            For m_x = 0 To m_W
    
                lODInc = m_ODM_O(lODx, lODy)
                With m_Bits32(m_x, m_y)
                     m_Bits08(m_x, m_y) = m_RGB4096_Inv(m_RGB4096_Trn(.R + lODInc), m_RGB4096_Trn(.G + lODInc), m_RGB4096_Trn(.B + lODInc))
                End With
    
                '-- Inc. ord. matrix column
                lODx = lODx + 1: If (lODx = 8) Then lODx = 0
            Next m_x
    
            '-- Inc. ord. matrix row
            lODx = 0
            lODy = lODy + 1: If (lODy = 8) Then lODy = 0
        Next m_y
    End If
End Sub

Private Sub pvDitherToHalftonePalette()

    '-- Dither...
    For m_y = 0 To m_H
        For m_x = 0 To m_W
            
            With m_Bits32(m_x, m_y)
                 m_Bits08(m_x, m_y) = m_HT_Inv(m_HT_Trn(.R), m_HT_Trn(.G), m_HT_Trn(.B))
            End With
        Next m_x
    Next m_y
End Sub

Private Sub pvDitherToHalftonePalette_Ordered()

  Dim lODx   As Long
  Dim lODy   As Long
  Dim lODInc As Long
  Dim aInvID As Byte
  
    '-- Dither...
    If (m_PreserveExactColors) Then
    
        For m_y = 0 To m_H
            For m_x = 0 To m_W
            
                With m_Bits32(m_x, m_y)
    
                    '-- Inv. index
                    aInvID = m_HT_Inv(m_HT_Trn(.R), m_HT_Trn(.G), m_HT_Trn(.B))
    
                    '-- Match to any palette color [?]
                    If (m_tPal(aInvID).R <> .R Or m_tPal(aInvID).G <> .G Or m_tPal(aInvID).B <> .B) Then
                        '-- Not: dither
                        lODInc = m_ODM_H(lODx, lODy)
                        m_Bits08(m_x, m_y) = m_HT_Inv(m_HT_Trn(.R + lODInc), m_HT_Trn(.G + lODInc), m_HT_Trn(.B + lODInc))
                      Else
                        '-- Yes: do not dither
                        m_Bits08(m_x, m_y) = aInvID
                    End If
                End With
    
                '-- Inc. ord. matrix column
                lODx = lODx + 1: If (lODx = 8) Then lODx = 0
            Next m_x

            '-- Inc. ord. matrix row
            lODx = 0
            lODy = lODy + 1: If (lODy = 8) Then lODy = 0
        Next m_y
        
      Else
        For m_y = 0 To m_H
            For m_x = 0 To m_W
            
                lODInc = m_ODM_H(lODx, lODy)
                With m_Bits32(m_x, m_y)
                     m_Bits08(m_x, m_y) = m_HT_Inv(m_HT_Trn(.R + lODInc), m_HT_Trn(.G + lODInc), m_HT_Trn(.B + lODInc))
                End With
    
                '-- Inc. ord. matrix column
                lODx = lODx + 1: If (lODx = 8) Then lODx = 0
            Next m_x

            '-- Inc. ord. matrix row
            lODx = 0
            lODy = lODy + 1: If (lODy = 8) Then lODy = 0
        Next m_y
    End If
End Sub

'//

Private Sub pvBuild_08bppSA(tSA As SAFEARRAY2D, oDIB As cDIB)

    '-- 8-bpp DIB mapping
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = oDIB.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = oDIB.BytesPerScanline
        .pvData = oDIB.lpBits
    End With
End Sub

Private Sub pvBuild_32bppSA(tSA As SAFEARRAY2D, oDIB As cDIB)

    '-- 32-bpp DIB mapping
    With tSA
        .cbElements = IIf(App.LogMode = 1, 1, 4)
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = oDIB.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = oDIB.Width
        .pvData = oDIB.lpBits
    End With
End Sub
