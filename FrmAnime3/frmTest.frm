VERSION 5.00
Begin VB.Form frmTest 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Sample Form"
   ClientHeight    =   6465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTest.frx":0000
   ScaleHeight     =   431
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   454
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AnimateForm Me, aload, Val(frmMain.txtFrameTime), Val(frmMain.txtBorder), Val(frmMain.txtFrames)
    End Sub

Private Sub Form_Unload(Cancel As Integer)
    AnimateForm Me, aUnload, Val(frmMain.txtFrameTime), Val(frmMain.txtBorder), Val(frmMain.txtFrames)
End Sub
