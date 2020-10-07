VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Tb_larg 
      Height          =   285
      Left            =   1275
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   750
   End
   Begin VB.TextBox Tb_long 
      Height          =   285
      Left            =   1275
      TabIndex        =   1
      Top             =   720
      Width           =   750
   End
   Begin ComCtl2.UpDown UpDown1 
      Height          =   255
      Left            =   2520
      TabIndex        =   0
      Top             =   360
      Width           =   255
      _ExtentX        =   423
      _ExtentY        =   450
      _Version        =   327681
      Max             =   10000
      Enabled         =   -1  'True
   End
   Begin VB.Label Lb_intLarg 
      Caption         =   "Largeur"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   375
      Width           =   615
   End
   Begin VB.Label Lb_intLong 
      Caption         =   "Longueur"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   735
      Width           =   735
   End
   Begin VB.Label Lb_uLarg 
      Caption         =   "m"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   375
      Width           =   855
   End
   Begin VB.Label Lb_uLong 
      Caption         =   "m"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   735
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
