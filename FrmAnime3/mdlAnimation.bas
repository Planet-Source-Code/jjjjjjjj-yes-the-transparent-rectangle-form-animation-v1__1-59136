Attribute VB_Name = "mdlAnimation"
'=============================================================
'            [ Auther : Jim Jose              ]
'            [ Email  : jimjosev33@yahoo.com  ]
'            [ Date   : 25/2/2005             ]
'=============================================================
'Hi,
'This code is made for all my friends in PSC. I uploaded this
'code inorder to get useful for anyone. If you found it useful
'please inform me. Your +Ve comments are my motivation. Good Luck!
'=============================================================
'****** Note ******
'I got the animation style from the 'Microsoft Word'. I realy wonder
'how they can refresh the screen so smoothly.

'This Form animation function 'AnimateForm' uses a form 'frmRect'.
'In the first phase of this code I tried to draw the rectangles
'in the screen. But that was too dificult to REFRESH the screen.
'So I use a form for that. We only need its fullscreen LOAD
'capaility and no even one line of form-code.

'****** Instructions ******
'1)You can use any SPEED by selecting the frame time.
'2)The FRAMES property is to controll the number of rectangles drawn
'3)You can also select the 'BORDER WIDTH' and BORDER COLOR
'4)Use low value FRAMES to reduse CPU Load ( for low performance systems )

'****** Importance ******
'The main problem that occures in the case of FORM ANIMATIONs
'is that the speed is affected by the computer performance and
'graphics load. The problem is solved to an extend by using the
'sleep API to controll the speed.
'The code is optimised for Zero memory leakage.
'=============================================================

Option Explicit

'[Types]
Public Type POINTAPI
        X As Long
        y As Long
End Type

'[Event Enum]
Public Enum AnimeEventEnum
    aUnload = 0
    aload = 1
End Enum

'[APIs]
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'[This function is the animation maker ]
'======================================================================================================
Public Sub AnimateForm(Frm As Form, aEvent As AnimeEventEnum, Optional ByVal FrameTime As Long = 3, _
                             Optional ByVal BorderWidth As Long = 2, Optional ByVal Frames As Long = 25, Optional BorderColor As Long = 0)
Static MousePos As POINTAPI
Dim X1 As Long, iNow As Long
Dim hrgn1 As Long, hrgn2 As Long
Dim ScrX As Long, ScrY As Long
Dim XIncr As Double, YIncr As Double
Dim WIncr As Double, HIncr As Double
Dim XValue As Long, Yvalue As Long

    frmRect.Show: frmRect.BackColor = BorderColor
    ScrX = Screen.TwipsPerPixelX
    ScrY = Screen.TwipsPerPixelY
    If aEvent = aload Then GetCursorPos MousePos
    XIncr = (Frm.Left / ScrX - MousePos.X) / Frames
    YIncr = (Frm.Top / ScrY - MousePos.y) / Frames
    WIncr = Frm.Width / ScrX / Frames
    HIncr = Frm.Height / ScrY / Frames
    
    For X1 = 0 To Frames
        If aEvent = aload Then iNow = X1 Else iNow = Frames - X1
        XValue = MousePos.X + iNow * XIncr: Yvalue = MousePos.y + iNow * YIncr
        hrgn1 = CreateRectRgn(XValue, Yvalue, XValue + iNow * WIncr, Yvalue + iNow * HIncr)
        hrgn2 = CreateRectRgn(XValue - BorderWidth, Yvalue - BorderWidth, XValue + iNow * WIncr + BorderWidth, Yvalue + iNow * HIncr + BorderWidth)
        CombineRgn hrgn1, hrgn1, hrgn2, 3
        SetWindowRgn frmRect.hwnd, hrgn1, True: DoEvents
        DeleteObject hrgn1: DeleteObject hrgn2
        Sleep FrameTime
    Next X1
    Unload frmRect
    
End Sub
'======================================================================================================


