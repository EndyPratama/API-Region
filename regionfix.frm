VERSION 5.00
Begin VB.Form regionfix 
   Caption         =   "Form1"
   ClientHeight    =   11070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16290
   LinkTopic       =   "Form1"
   ScaleHeight     =   11070
   ScaleWidth      =   16290
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "regionfix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Sub Form_Load()

'HURUF PERTAMA

q1 = CreateEllipticRgn(100, 50, 300, 300)    'base
q2 = CreateEllipticRgn(150, 100, 250, 250)   'bidang cut
q2a = CreateEllipticRgn(-17, 30, 163, 180)
q3 = CreateRectRgn(50, 250, 65, 400)
q4 = CreateRectRgn(150, 170, 165, 310)
q5 = CreateRectRgn(85, 225, 235, 235)
q6 = CreateRectRgn(90, 300, 230, 310)
q7 = CreateRectRgn(100, 297, 115, 400)
q8 = CreateRectRgn(200, 297, 215, 400)
q9 = CreateRectRgn(90, 380, 230, 390)
q10 = CreateEllipticRgn(170, 50, 305, 300)
q11 = CreateEllipticRgn(170, 50, 290, 300)
q12 = CreateEllipticRgn(210, 30, 390, 180)
q13 = CreateEllipticRgn(180, 130, 355, 410)
q14 = CreateEllipticRgn(170, 130, 340, 405)
q15 = CreateRectRgn(170, 130, 440, 256)
q16 = CreateRectRgn(280, 240, 369, 256)
q17 = CreateEllipticRgn(275, 145, 450, 425)
q18 = CreateEllipticRgn(283, 125, 445, 405)
q19 = CreateEllipticRgn(193, 125, 355, 235)
q20 = CreateEllipticRgn(303, 135, 465, 415)
q21 = CreateEllipticRgn(303, 400, 435, 515)

'HURUF PERTAMA
CombineRgn q1, q1, q2, 4 'perpotongan atas miring kekiri
'CombineRgn q1, q1, q2a, 4 'perpotongan atas miring kekiri
CombineRgn q1, q1, q3, 2
CombineRgn q1, q1, q4, 2
CombineRgn q1, q1, q5, 2
CombineRgn q1, q1, q6, 2
CombineRgn q1, q1, q7, 2
CombineRgn q1, q1, q8, 2
CombineRgn q1, q1, q9, 2
'CombineRgn q2, q10, q11, 4
'CombineRgn q2, q2, q12, 4
'CombineRgn q3, q13, q14, 4
'CombineRgn q3, q3, q15, 4
'CombineRgn q3, q3, q16, 2
'CombineRgn q4, q17, q18, 4 'perpotongan bawah pojok miring ke kanan
'CombineRgn q4, q4, q19, 4 'perpotongan bawah pojok miring ke kanan
'CombineRgn q4, q4, q20, 4 'perpotongan bawah pojok miring ke kanan
'CombineRgn q4, q4, q21, 4 'perpotongan bawah pojok miring ke kanan
'CombineRgn q1, q1, q2, 2
'CombineRgn q1, q1, q3, 2
'CombineRgn q1, q1, q4, 2

SetWindowRgn Me.hwnd, q1, True

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

ReleaseCapture

SendMessage Me.hwnd, &HA1, 2, 0&

End Sub
