Attribute VB_Name = "mdl3DForm"
'=============================================================
'=============================================================
'            [ Auther  : 'Jim Jose '           ]
'            [ Email   : jimjosev33@yahoo.com  ]
'            [ Created : 07/03/2005            ]
'=============================================================
'            [ Project : '3D Form'             ]
'            [ Page    : 'Not Set              ]
'=============================================================
'             'Please do not modify this Title'
'=============================================================
'* Usage >:
'   Add the function call to the form load event.
'* Working >:
'   Will load a new Picturebox on to the form and draw regions on that.
'* Optional Parameters >:
'   -1 for the startcolor will capture the color of lower-right corner.
'=============================================================

Option Explicit

'[APIs]
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'[This function is the 3D Maker ]
'=============================================================
Public Sub ProjectForm(Frm As Form, Optional ByVal ThicknessX As Long = 15, Optional ByVal ThicknessY As Long = 10, _
                            Optional ByVal StartColor As Long = -1, Optional ByVal EndColor As Long = 0, Optional ByVal Curvature As Double = 25, Optional ByVal Frames As Double = 25)
Dim PicDraw As PictureBox
Dim X As Long, xIncr As Double, yIncr As Double
Dim ScrX As Long, ScrY As Long
Dim hBrush As Long, fScale As Long
Dim fWidth As Long, fHeight As Long
Dim R1 As Long, G1 As Long, B1 As Long
Dim R2 As Long, G2 As Long, B2 As Long
Dim hRgn1 As Long, TempRgn As Long, frmRgn As Long
Dim Rincr As Double, Gincr As Double, Bincr As Double

    'Some initialisation
    fScale = Frm.ScaleMode: Frm.ScaleMode = 1
    ScrX = Screen.TwipsPerPixelX: ScrY = Screen.TwipsPerPixelY
    If StartColor = -1 Then StartColor = Frm.Point(Frm.ScaleWidth - 10, Frm.ScaleHeight - 10)
    
    'Setting the color gradients and XY increments
    GetRGB EndColor, R1, G1, B1: GetRGB StartColor, R2, G2, B2
    Rincr = (R2 - R1) / Frames: Gincr = (G2 - G1) / Frames: Bincr = (B2 - B1) / Frames
    xIncr = ThicknessX / Frames: yIncr = ThicknessY / Frames
    
    'Loading a new picturebox into the frm
    Set PicDraw = Frm.Controls.Add("vb.picturebox", "Pic3d"): PicDraw.AutoRedraw = True
    PicDraw.Move 0, 0, Frm.Width * 2, Frm.Height * 2: PicDraw.BorderStyle = 0
    
    'Creating the initilal regions
    fWidth = Frm.Width / ScrX: fHeight = Frm.Height / ScrY
    TempRgn = CreateRoundRectRgn(2, 2, fWidth, fHeight, Curvature, Curvature)
    frmRgn = CreateRoundRectRgn(2, 2, fWidth, fHeight, Curvature, Curvature)
    
    'Resizing the form with the 'Thickness'
    fWidth = fWidth + ThicknessX: fHeight = fHeight + ThicknessY
    Frm.Width = fWidth * ScrX: Frm.Height = fHeight * ScrY

    'Creating the graidient regions and saving all the Rgns together in 'frmRgn'
    For X = 0 To Frames
        hRgn1 = CreateRoundRectRgn((Frames - X) * xIncr, (Frames - X) * yIncr, fWidth - X * xIncr, fHeight - X * yIncr, Curvature, Curvature)
        CombineRgn frmRgn, frmRgn, hRgn1, 2
        hBrush = CreateSolidBrush(RGB(R1 + X * Rincr, G1 + X * Gincr, B1 + X * Bincr))
        FillRgn PicDraw.hdc, hRgn1, hBrush
        DeleteObject hRgn1: DeleteObject hBrush
    Next X
    
    'Combining the initial frmRgn and tempRgn to prepare the new regions for 'PicDraw' and 'Frm'
    SetWindowRgn Frm.hwnd, frmRgn, True
    frmRgn = CreateRoundRectRgn(0, 0, 2 * fWidth, 2 * fHeight, Curvature, Curvature)
    CombineRgn frmRgn, frmRgn, TempRgn, 3: DeleteObject TempRgn
    SetWindowRgn PicDraw.hwnd, frmRgn, True

    'Finalising
    DeleteObject frmRgn: DeleteObject TempRgn
    PicDraw.Visible = True: Frm.ScaleMode = fScale
End Sub
'=============================================================

'[ Gets the RGB values ]
'=============================================================
Public Sub GetRGB(ByVal LngCol As Long, r As Long, g As Long, b As Long)
  r = LngCol Mod 256
  g = (LngCol And vbGreen) / 256 'Green
  b = (LngCol And vbBlue) / 65536 'Blue
End Sub
'=============================================================

