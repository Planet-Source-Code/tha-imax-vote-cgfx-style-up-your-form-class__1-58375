VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGfX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------
'*********imaX Class*********-
'*********Binary Crew********-
'************2oo3************-
'*********#82186397**********- .icQ
'      the_imax@yahoo.de     - .emaiL
'-----------------------------
'Help Comments:

'How to Initialize Class?
'     eg. dim GfX as new cGfX
'   -----------
'     if you want to use the fade effects you've to
'     set the CurrentRate at first to 255
'     Write the next lines to Form_Load()'     eg. gfX.SetCurrentRate 255
'   -----------
'What functions\subs are supportet?
'##   PolyTrans(PictureBox, Form)
'         the PolyTrans Sub fits your form like the Picture in the PictureBox.
'         the pixel in the upper left corner sets the transperent color.
'         CreatePoly is a Sub to individualize your Form very easy.
'##   DoTrans(HwnD, Rate)
'         Make a form transperent with the Rate from 1 to 255.
'         0 is invisible.
'         255 is not transperent.
'         And between 0 and 255...Try it ;)
'##   FadeIn(HwnD,[Steps],[Waitfor])
'         The FadeIn Sub sets the transperent of the selectet form in steps.
'         So it is a soft Rate Change.
'         [Steps] = How fast will the fade be done?
'         [WaitFor] = How long will the sub wait every run?
'##   FadeOut(HwnD,[Steps],[Waitfor])
'         The same like FadeIn but with fading out effect.
'##   FadeTo(HwnD, FadeRate, [Steps], [WaitFor])
'         The Same like the upperones but you can set the FadeRate.
'         FadeRate means the Rate you want to fade to.
'##   MoveForm(Form)
'         If you havn't a TitleBar in your Form, you can't move your form.
'         Insert this Sub in a MouseDown Sub, you'll able to move the form.
'   ----------
'   Any Questions? Ask me! #82186397
'   tHa_imaX [Digital Death Crew - Binary Crew]

Private Declare Function GetPixel Lib "gdi32" _
(ByVal hDC As Long, ByVal lngX As Long, ByVal lngY As Long) As Long

Private Declare Function SetWindowRgn Lib "user32" _
(ByVal HwnD As Long, ByVal hRgn As Long, _
ByVal bRedraw As Boolean) As Long

Private Declare Function CreateRectRgn Lib "gdi32" _
(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
ByVal Y2 As Long) As Long

Private Declare Function CombineRgn Lib "gdi32" _
(ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, _
ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal HwnD As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function DeleteObject Lib "gdi32" _
(ByVal hObject As Long) As Long

Private Const RGN_OR = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2


Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal HwnD As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal HwnD As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal HwnD As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal HwnD As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, crKey As Long, ByVal pblend As Long, ByVal dwFlags As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type SIZE
    cx As Long
    cy As Long
End Type

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const ULW_COLORKEY = &H1
Private Const ULW_ALPHA = &H2
Private Const ULW_OPAQUE = &H4
Private Const AC_SRC_OVER = &H0
Private Const AC_SRC_ALPHA = &H1
Private Const AC_SRC_NO_PREMULT_ALPHA = &H1
Private Const AC_SRC_NO_ALPHA = &H2
Private Const AC_DST_NO_PREMULT_ALPHA = &H10
Private Const AC_DST_NO_ALPHA = &H20
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2

Private CurrentRate  As Byte

Public Sub DoTrans(HwnD As Long, Rate As Integer)
    CurrentRate = Rate
    Dim WinInfo As Long
    WinInfo = GetWindowLong(HwnD, GWL_EXSTYLE)
    If Rate < 255 Then
        WinInfo = WinInfo Or WS_EX_LAYERED
        SetWindowLong HwnD, GWL_EXSTYLE, WinInfo
        SetLayeredWindowAttributes HwnD, 0, Rate, LWA_ALPHA
    Else
        WinInfo = WinInfo Xor WS_EX_LAYERED
        SetWindowLong HwnD, GWL_EXSTYLE, WinInfo
    End If
End Sub


Private Function CreatePoly(objBild As PictureBox) As Long
    
    Dim lngTransparenteFarbe As Long
    Dim bolErsterBereich As Boolean
    Dim bolEingangsLinie As Boolean
    Dim hDC As Long
    Dim lngBildWeite As Long
    Dim lngBildHoehe As Long
    Dim lngX As Long
    Dim lngY As Long
    Dim lngStartLinie_X As Long
    Dim lngGesamtBereich As Long
    Dim lngLinienBereich As Long

    
    hDC = objBild.hDC
    lngBildWeite = objBild.ScaleWidth
    lngBildHoehe = objBild.ScaleHeight
    
    bolErsterBereich = True: bolEingangsLinie = False
    lngX = lngY = lngStartLinie_X = 0
    'Die transparente Farbe ist die Farbe die sich in der
    'PictureBox oben Links befindet...
    lngTransparenteFarbe = GetPixel(hDC, 0, 0)
    For lngY = 0 To lngBildHoehe - 1
    For lngX = 0 To lngBildWeite - 1
    If GetPixel(hDC, lngX, lngY) = lngTransparenteFarbe Or _
    lngX = lngBildWeite Then
    If bolEingangsLinie Then
    bolEingangsLinie = False
    lngLinienBereich = CreateRectRgn(lngStartLinie_X, lngY, lngX, lngY + 1)
    If bolErsterBereich Then
    lngGesamtBereich = lngLinienBereich
    bolErsterBereich = False
    Else
    CombineRgn lngGesamtBereich, lngGesamtBereich, lngLinienBereich, RGN_OR
    DeleteObject lngLinienBereich
    End If
    End If
    Else
    If Not bolEingangsLinie Then
    bolEingangsLinie = True
    lngStartLinie_X = lngX
    End If
    End If
    Next
    Next
    CreatePoly = lngGesamtBereich
End Function

Public Sub PolyTrans(PicBox As PictureBox, fr As Form)
  
  With PicBox
    .ScaleMode = vbPixels
    .AutoRedraw = True
    .AutoSize = True
    .BorderStyle = None
  End With
  
  Dim FensterBereich As Long
    fr.Width = PicBox.Width
    fr.Height = PicBox.Height
    FensterBereich = CreatePoly(PicBox)
    SetWindowRgn fr.HwnD, FensterBereich, True
End Sub


Public Sub HelpMe()
  MsgBox "Check your Debug window!", vbInformation, "HelpMe!"
  Debug.Print "imaX ClassModule"
  Debug.Print "Binary Crew 2oo3"
  Debug.Print "ICQ 82186397"
  Debug.Print "This Module is designed for Designer in VB"
  Debug.Print "Have fun. imaX"
End Sub

Public Sub FadeOut(Hwn As Long, Optional Steps As Byte = 5, Optional WaitFor As Integer = 10000)
  Dim i As Integer
  For i = 255 To 0 Step -Steps
    DoTrans Hwn, i
    For t = 1 To WaitFor: DoEvents: Next
  Next
End Sub

Public Sub FadeIn(Hwn As Long, Optional Steps As Byte = 5, Optional WaitFor As Integer = 10000)
  Dim i As Integer
  For i = 0 To 255 Step Steps
    DoTrans Hwn, i
    For t = 1 To WaitFor: DoEvents: Next
  Next
End Sub

Public Sub FadeTo(HwnD As Long, FadeRate As Byte, Optional Steps As Byte = 3, Optional WaitFor As Integer = 10000)
  Do
    If FadeRate > CurrentRate Then
      DoTrans HwnD, CurrentRate + Steps
      If FadeRate < CurrentRate Then Exit Sub
    Else
      DoTrans HwnD, CurrentRate - Steps
      If FadeRate > CurrentRate Then Exit Sub
    End If
    For t = 1 To WaitFor: DoEvents: Next
  Loop
End Sub

Public Sub SetCurrentRate(CurRate As Byte)
  CurrentRate = CurRate
End Sub

Public Sub MoveForm(fRm As Form)
    ReleaseCapture
    SendMessage fRm.HwnD, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
