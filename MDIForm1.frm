VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
Form1.Show
Form2.Show
'Debug.Print Me.Left, Me.Width, Me.Top, Me.Height
Form1.Left = 0 'Me.Left
Form1.Width = Me.Width - 200
Form1.Top = 0 'Me.Top
Form1.Height = (Me.Height - 500) / 4 * 3
'Debug.Print Form1.Left, Form1.Width, Form1.Top, Form1.Height
Form2.Left = 0 'Me.Left
Form2.Width = Me.Width - 200
Form2.Top = Form1.Height
Form2.Height = Me.Height - (Form1.Height + 500) '(Me.Height - 500) / 3
End Sub

Private Sub MDIForm_Resize()
Form1.Left = 0 'Me.Left
Form1.Width = Me.Width - 200
Form1.Top = 0 'Me.Top
Form1.Height = (Me.Height - 500) / 4 * 3
'Debug.Print Form1.Left, Form1.Width, Form1.Top, Form1.Height
Form2.Left = 0 'Me.Left
Form2.Width = Me.Width - 200
Form2.Top = Form1.Height
Form2.Height = Me.Height - (Form1.Height + 500) '(Me.Height - 500) / 3

End Sub
