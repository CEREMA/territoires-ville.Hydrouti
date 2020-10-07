VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmPrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Feuille d'édition"
   ClientHeight    =   15825
   ClientLeft      =   150
   ClientTop       =   615
   ClientWidth     =   11805
   Icon            =   "FrmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   15661.27
   ScaleMode       =   0  'User
   ScaleWidth      =   11805
   Begin MSComDlg.CommonDialog cdlg1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Cmd_annul 
      Cancel          =   -1  'True
      Caption         =   "Fermer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   258
      Left            =   9840
      TabIndex        =   98
      Top             =   840
      Width           =   1000
   End
   Begin VB.CommandButton Cmd_apres 
      Caption         =   ">"
      Height          =   258
      Left            =   10440
      TabIndex        =   97
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton Cmd_avant 
      Caption         =   "<"
      Height          =   258
      Left            =   9960
      TabIndex        =   96
      Top             =   1320
      Width           =   255
   End
   Begin VB.Frame Frm_par42 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame42"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   1200
      TabIndex        =   88
      Top             =   9240
      Width           =   9015
      Begin VB.Label Lb_ldon421 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Libellé"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   95
         Top             =   350
         Width           =   3150
      End
      Begin VB.Label Lb_don421 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   3400
         TabIndex        =   94
         Top             =   350
         Width           =   875
      End
      Begin VB.Label Lb_unit421 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   4275
         TabIndex        =   93
         Top             =   350
         Width           =   775
      End
      Begin VB.Label Lb_don423 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   7090
         TabIndex        =   92
         Top             =   350
         Width           =   925
      End
      Begin VB.Label Lb_unit423 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   8015
         TabIndex        =   91
         Top             =   350
         Width           =   775
      End
      Begin VB.Label Lb_don422 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   5270
         TabIndex        =   90
         Top             =   350
         Width           =   875
      End
      Begin VB.Label Lb_unit422 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   6145
         TabIndex        =   89
         Top             =   350
         Width           =   775
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimer"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   258
      Left            =   9840
      TabIndex        =   1
      Top             =   360
      Width           =   1000
   End
   Begin VB.Frame Frm_par3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   1200
      TabIndex        =   12
      Top             =   3960
      Width           =   9015
      Begin VB.Label Lb_ldon3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Libellé"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   800
         TabIndex        =   15
         Top             =   350
         Width           =   4665
      End
      Begin VB.Label Lb_don3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   5500
         TabIndex        =   14
         Top             =   350
         Width           =   1700
      End
      Begin VB.Label Lb_unit3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   7300
         TabIndex        =   13
         Top             =   350
         Width           =   720
      End
   End
   Begin VB.Frame Frm_par2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   1200
      TabIndex        =   8
      Top             =   3600
      Width           =   9015
      Begin VB.Label Lb_unit2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   7300
         TabIndex        =   11
         Top             =   350
         Width           =   720
      End
      Begin VB.Label Lb_don2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   5500
         TabIndex        =   10
         Top             =   350
         Width           =   1700
      End
      Begin VB.Label Lb_ldon2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Libellé"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   800
         TabIndex        =   9
         Top             =   350
         Width           =   4665
      End
   End
   Begin VB.Frame Frm_par1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   1200
      TabIndex        =   4
      Top             =   3240
      Visible         =   0   'False
      Width           =   9015
      Begin VB.Label Lb_ldon1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Libellé"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   800
         TabIndex        =   7
         Top             =   350
         Width           =   4665
      End
      Begin VB.Label Lb_don1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   5500
         TabIndex        =   6
         Top             =   345
         Width           =   1700
      End
      Begin VB.Label Lb_unit1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   7300
         TabIndex        =   5
         Top             =   350
         Width           =   720
      End
   End
   Begin VB.Frame Frm_par5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   1200
      TabIndex        =   20
      Top             =   4680
      Width           =   9015
      Begin VB.Label Lb_ldon5 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Libellé"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   800
         TabIndex        =   23
         Top             =   350
         Width           =   4665
      End
      Begin VB.Label Lb_don5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   5500
         TabIndex        =   22
         Top             =   350
         Width           =   1700
      End
      Begin VB.Label Lb_unit5 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   7300
         TabIndex        =   21
         Top             =   350
         Width           =   720
      End
   End
   Begin VB.Frame Frm_par4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   1200
      TabIndex        =   16
      Top             =   4320
      Width           =   9015
      Begin VB.Label Lb_unit4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   7300
         TabIndex        =   19
         Top             =   350
         Width           =   720
      End
      Begin VB.Label Lb_don4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   5500
         TabIndex        =   18
         Top             =   350
         Width           =   1700
      End
      Begin VB.Label Lb_ldon4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Libellé"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   800
         TabIndex        =   17
         Top             =   350
         Width           =   4665
      End
   End
   Begin VB.Frame Frm_par6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   1200
      TabIndex        =   24
      Top             =   5040
      Width           =   9015
      Begin VB.Label Lb_unit6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   7300
         TabIndex        =   27
         Top             =   350
         Width           =   720
      End
      Begin VB.Label Lb_don6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   5500
         TabIndex        =   26
         Top             =   350
         Width           =   1700
      End
      Begin VB.Label Lb_ldon6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Libellé"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   800
         TabIndex        =   25
         Top             =   350
         Width           =   4665
      End
   End
   Begin VB.Frame Frm_par7 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   1200
      TabIndex        =   28
      Top             =   5400
      Width           =   9015
      Begin VB.Label Lb_ldon7 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Libellé"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   800
         TabIndex        =   31
         Top             =   350
         Width           =   4665
      End
      Begin VB.Label Lb_don7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   5500
         TabIndex        =   30
         Top             =   350
         Width           =   1700
      End
      Begin VB.Label Lb_unit7 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   7300
         TabIndex        =   29
         Top             =   350
         Width           =   720
      End
   End
   Begin VB.Frame Frm_par21 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame21"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   1200
      TabIndex        =   34
      Top             =   6000
      Width           =   9015
      Begin VB.Label Lb_unit21 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   8060
         TabIndex        =   39
         Top             =   350
         Width           =   820
      End
      Begin VB.Label Lb_don21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   6360
         TabIndex        =   38
         Top             =   350
         Width           =   1695
      End
      Begin VB.Label Lb_unit11 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   5540
         TabIndex        =   37
         Top             =   350
         Width           =   820
      End
      Begin VB.Label Lb_don11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   3840
         TabIndex        =   36
         Top             =   350
         Width           =   1695
      End
      Begin VB.Label Lb_ldon21 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Libellé"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   120
         TabIndex        =   35
         Top             =   350
         Width           =   3705
      End
   End
   Begin VB.Frame Frm_par22 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame22"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   1200
      TabIndex        =   40
      Top             =   6360
      Width           =   9015
      Begin VB.Label Lb_ldon22 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Libellé"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   120
         TabIndex        =   45
         Top             =   350
         Width           =   3705
      End
      Begin VB.Label Lb_don12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   3840
         TabIndex        =   44
         Top             =   350
         Width           =   1695
      End
      Begin VB.Label Lb_unit12 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   5540
         TabIndex        =   43
         Top             =   350
         Width           =   820
      End
      Begin VB.Label Lb_don22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   6360
         TabIndex        =   42
         Top             =   350
         Width           =   1695
      End
      Begin VB.Label Lb_unit22 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   8060
         TabIndex        =   41
         Top             =   350
         Width           =   820
      End
   End
   Begin VB.Frame Frm_par23 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame23"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   1200
      TabIndex        =   46
      Top             =   6720
      Width           =   9015
      Begin VB.Label Lb_unit23 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   8060
         TabIndex        =   51
         Top             =   350
         Width           =   820
      End
      Begin VB.Label Lb_don23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   6360
         TabIndex        =   50
         Top             =   350
         Width           =   1695
      End
      Begin VB.Label Lb_unit13 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   5535
         TabIndex        =   49
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Lb_don13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   3840
         TabIndex        =   48
         Top             =   350
         Width           =   1695
      End
      Begin VB.Label Lb_ldon23 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Libellé"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   120
         TabIndex        =   47
         Top             =   350
         Width           =   3705
      End
   End
   Begin VB.Frame Frm_par31 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame31"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   1200
      TabIndex        =   59
      Top             =   7320
      Width           =   9015
      Begin VB.Label Lb_ldon312 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Libellé"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   65
         Top             =   360
         Width           =   2580
      End
      Begin VB.Label Lb_ldon311 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Libellé"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   64
         Top             =   350
         Width           =   2580
      End
      Begin VB.Label Lb_don311 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   63
         Top             =   350
         Width           =   825
      End
      Begin VB.Label Lb_unit311 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   3585
         TabIndex        =   62
         Top             =   350
         Width           =   825
      End
      Begin VB.Label Lb_don312 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   7200
         TabIndex        =   61
         Top             =   350
         Width           =   825
      End
      Begin VB.Label Lb_unit312 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   8060
         TabIndex        =   60
         Top             =   350
         Width           =   820
      End
   End
   Begin VB.Frame Frm_par33 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame33"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   1200
      TabIndex        =   73
      Top             =   8040
      Width           =   9015
      Begin VB.Label Lb_ldon332 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Libellé"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   79
         Top             =   360
         Width           =   2780
      End
      Begin VB.Label Lb_ldon331 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Libellé"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   78
         Top             =   350
         Width           =   2580
      End
      Begin VB.Label Lb_don331 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   77
         Top             =   350
         Width           =   825
      End
      Begin VB.Label Lb_unit331 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   3585
         TabIndex        =   76
         Top             =   350
         Width           =   825
      End
      Begin VB.Label Lb_don332 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   7200
         TabIndex        =   75
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Lb_unit332 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   8060
         TabIndex        =   74
         Top             =   350
         Width           =   820
      End
   End
   Begin VB.Frame Frm_par32 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame32"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   1200
      TabIndex        =   66
      Top             =   7680
      Width           =   9015
      Begin VB.Label Lb_unit322 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   250
         Index           =   0
         Left            =   8060
         TabIndex        =   72
         Top             =   350
         Width           =   820
      End
      Begin VB.Label Lb_don322 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   7200
         TabIndex        =   71
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Lb_unit321 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   3585
         TabIndex        =   70
         Top             =   350
         Width           =   825
      End
      Begin VB.Label Lb_don321 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   69
         Top             =   350
         Width           =   825
      End
      Begin VB.Label Lb_ldon321 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Libellé"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   68
         Top             =   350
         Width           =   2780
      End
      Begin VB.Label Lb_ldon322 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Libellé"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   4560
         TabIndex        =   67
         Top             =   360
         Width           =   2580
      End
   End
   Begin VB.Frame Frm_par41 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame41"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   1200
      TabIndex        =   80
      Top             =   8760
      Width           =   9015
      Begin VB.Label Lb_unit412 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   6145
         TabIndex        =   87
         Top             =   350
         Width           =   775
      End
      Begin VB.Label Lb_don412 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   5270
         TabIndex        =   86
         Top             =   350
         Width           =   875
      End
      Begin VB.Label Lb_unit413 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   8015
         TabIndex        =   85
         Top             =   350
         Width           =   775
      End
      Begin VB.Label Lb_don413 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   7090
         TabIndex        =   84
         Top             =   350
         Width           =   925
      End
      Begin VB.Label Lb_unit411 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Unite"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   4275
         TabIndex        =   83
         Top             =   350
         Width           =   775
      End
      Begin VB.Label Lb_don411 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Donnee"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   3400
         TabIndex        =   82
         Top             =   350
         Width           =   875
      End
      Begin VB.Label Lb_ldon411 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Libellé"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   81
         Top             =   360
         Width           =   3150
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   3375
      Left            =   1560
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   535
      TabIndex        =   3
      Top             =   11040
      Width           =   8055
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   1680
      ScaleHeight     =   3345
      ScaleWidth      =   8025
      TabIndex        =   53
      Top             =   11040
      Width           =   8055
   End
   Begin VB.Label Dossier 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Dossier"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   100
      Top             =   15240
      Width           =   10000
   End
   Begin VB.Label Etude 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Etude"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   99
      Top             =   1920
      Width           =   10000
   End
   Begin VB.Label Lb_des1b 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   202
      Left            =   1560
      TabIndex        =   58
      Top             =   14880
      Width           =   8055
   End
   Begin VB.Label Lb_des1h 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   202
      Left            =   1560
      TabIndex        =   57
      Top             =   14760
      Width           =   8055
   End
   Begin VB.Label Lb_des2b 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   202
      Left            =   1560
      TabIndex        =   56
      Top             =   14640
      Width           =   8055
   End
   Begin VB.Label Lb_des2h 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000000&
      Height          =   202
      Left            =   1560
      TabIndex        =   55
      Top             =   14520
      Width           =   8055
   End
   Begin VB.Label Lb_titre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lb_titre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   54
      Top             =   2760
      Width           =   10185
   End
   Begin VB.Label Lb_ent 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   2760
      TabIndex        =   52
      Top             =   480
      Width           =   6855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "HYDROUTI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   360
      TabIndex        =   32
      Top             =   720
      Width           =   2175
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   240
      X2              =   11040
      Y1              =   15438.6
      Y2              =   15438.6
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   11040
      X2              =   11040
      Y1              =   237.517
      Y2              =   15438.6
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   11040
      X2              =   240
      Y1              =   237.517
      Y2              =   237.517
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   240
      X2              =   240
      Y1              =   237.517
      Y2              =   15438.6
   End
   Begin VB.Label Titre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Titre"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   10185
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   240
      X2              =   11040
      Y1              =   1781.377
      Y2              =   1781.377
   End
   Begin VB.Label Entete 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Entete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   1080
      Width           =   6855
   End
   Begin VB.Label Label2 
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   360
      TabIndex        =   33
      Top             =   360
      Width           =   2175
   End
   Begin VB.Menu mnuetude 
      Caption         =   "&Etude"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "FrmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private owner As MDIFrm_menu
Public Type1 As String
Public nomobjet As String
Public titre1 As String
Public sstitre1 As String
Public ssTitre2 As String
Public ssTitre3 As String
Public ssTitre4 As String
Public ssTitre5 As String
Public ssTitre6 As String
Public des1_titrh As String
Public des1_titrb As String
Public des2_titrh As String
Public des2_titrb As String
Private form_top As Double
Private Sub ini_form()
Dim larg As Double, haut As Double
If Type1 = "deversoir" Or Type1 = "pompe" Or Type1 = "deversoiror" Then
    Cmd_avant.Visible = True
    Cmd_apres.Visible = True
    Cmd_avant.Enabled = False
    Cmd_apres.Enabled = True
Else
    Cmd_avant.Visible = False
    Cmd_apres.Visible = False
End If
    Frm_par1.Visible = False
    Frm_par2.Visible = False
    Frm_par3.Visible = False
    Frm_par4.Visible = False
    Frm_par5.Visible = False
    Frm_par6.Visible = False
    Frm_par7.Visible = False
    Frm_par21.Visible = False
    Frm_par22.Visible = False
    Frm_par23.Visible = False
    Frm_par31.Visible = False
    Frm_par32.Visible = False
    Frm_par33.Visible = False
    Frm_par41.Visible = False
    Frm_par42.Visible = False
    Picture2.Visible = False
    Lb_des2h.Visible = False
    Lb_des2b.Visible = False
    Lb_des1h.Visible = False
    Lb_des1b.Visible = False
End Sub
Private Sub ini_form1()
    Call ini_form
        larg = Frm_desprint.UC_graphique2.lire_width
        haut = Frm_desprint.UC_graphique2.lire_height
    Picture2.Left = Frm_par1.Left + ((Frm_par1.Width - larg) / 2)
    Picture2.Height = haut
    Picture2.Width = larg
    larg = Frm_desprint.UC_graphique1.lire_width
    haut = Frm_desprint.UC_graphique1.lire_height
    Picture1.Left = Frm_par1.Left + ((Frm_par1.Width - larg) / 2)
    Picture1.Height = haut
    Picture1.Width = larg
    Lb_ent.Caption = "Boite à Outils Hydrologie,Hydraulique et Assainissement"
    Entete(0).Caption = text_serv1 '"Centre d'Etudes Techniques de l'Equipement de L'Est"
    Load Entete(1)
    With Entete(1)
        .Caption = text_serv2 '"Laboratoire Régional de Nancy"
        .Top = Entete(0).Top + Entete(0).Height + 10
        .Visible = True
    End With

End Sub

Private Sub Cmd_annul_Click()
    Unload Me
End Sub

Private Sub Cmd_apres_Click()
Dim ytop As Double, ydec As Double
Dim nometude As String
nometude = "Etude : " + nom_etude
Me.MousePointer = 11
    Call ini_form
    Cmd_avant.Enabled = True
    Cmd_apres.Enabled = False
    ytop = 3500
    ydec = 300
    Etude.Caption = nometude
    Dossier.Caption = "Dossier : " + nom_fich_edit
Select Case Type1
        Case Is = "deversoir"
'unload des infos de la page 1
    unl_frm "Frm_par1", "list_don1"
    unl_frm "Frm_par2", "list_don2"
    unl_frmq "Frm_par41", "list_don3"
    unl_frmq "Frm_par42", "list_don4"
'**********
    Titre(0).Caption = titre1
    Lb_titre.Caption = nomobjet + "  --page 2/2--"
    ytop = ecr_frmd("Frm_par21", ssTitre5, "list_don5", ytop)
    ytop = ytop + ydec
    Picture1.Visible = True
    Picture1.Top = ytop
    ytop = Picture1.Top + Picture1.Height
    ytop = ytop + ydec + 50
    ytop = ecr_frmd("Frm_par22", ssTitre6, "list_don6", ytop)
    ytop = ytop + ydec
    Picture2.Visible = True
    Picture2.Top = ytop
        Case Is = "deversoiror"
    ytop = 3600
    ydec = 1000
'unload des infos de la page 1
    unl_frm "Frm_par1", "list_don1"
    unl_frmq "Frm_par41", "list_don3"
    unl_frm "Frm_par2", "list_don4"
    unl_frm "Frm_par3", "list_don5"
'**********
    Titre(0).Caption = titre1
    Titre(0).FontSize = 13.5
    Lb_titre.Caption = nomobjet + "  --page 2/2--"
    ytop = ecr_frm("Frm_par4", ssTitre6, "list_don6", ytop)
    ytop = ytop + ydec
    Picture1.Visible = True
    Picture1.Height = 6000
    Picture1.Top = ytop
        Case Is = "pompe"
'unload des infos de la page 1
    ytop = 3200
    ydec = 200
    unl_frmd "Frm_par21", "list_don1"
    unl_frm "Frm_par2", "list_don2"
    unl_frmq "Frm_par41", "list_don3"
    unl_frm "Frm_par4", "list_don4"
'**********
    Titre(0).Caption = titre1
    Lb_titre.Caption = nomobjet + "  --page 2/2--"
    ytop = ecr_frmd("Frm_par22", ssTitre5, "list_don5", ytop)
    ytop = ytop + ydec
    Picture1.Visible = True
    Picture1.Height = 6000
    Picture1.Top = ytop

End Select
Me.MousePointer = 1

End Sub

Private Sub Cmd_avant_Click()
Dim ytop As Double, ydec As Double
Dim nometude As String
     nometude = "Etude : " + nom_etude
  Call ini_form
    Cmd_avant.Enabled = False
    Cmd_apres.Enabled = True
    ytop = 3500
    ydec = 300
    If Type1 = "pompe" Then
        ytop = 3200
        ydec = 150
    End If
    If Lb_titre.Caption = nomobjet + "  --page 2/2--" Then
'unload des infos de la page 2
        Select Case Type1
            Case Is = "deversoir"
                unl_frmd "Frm_par21", "list_don5"
                unl_frmd "Frm_par22", "list_don6"
            Case Is = "deversoiror"
                unl_frmd "Frm_par4", "list_don6"
            Case Is = "pompe"
                unl_frmd "Frm_par22", "list_don5"
       End Select
''**********
    End If
    Etude.Caption = nometude
    Dossier.Caption = "Dossier : " + nom_fich_edit
    Titre(0).Caption = titre1
    Lb_titre.Caption = nomobjet + "  --page 1/2--"
    Select Case Type1
    
    Case Is = "deversoir"
        ytop = ecr_frm("Frm_par1", sstitre1, "list_don1", ytop)
        ytop = ytop + ydec
        ytop = ecr_frm("Frm_par2", ssTitre2, "list_don2", ytop)
        ytop = ytop + ydec
        ytop = ecr_frmq("Frm_par41", ssTitre3, "list_don3", ytop)
        ytop = ytop + ydec
        ytop = ecr_frmq("Frm_par42", ssTitre4, "list_don4", ytop)
    Case Is = "deversoiror"
        Titre(0).FontSize = 13.5
        ytop = ecr_frm("Frm_par1", sstitre1, "list_don1", ytop)
        ytop = ytop + ydec
'        ytop = ecr_frm("Frm_par2", ssTitre2, "list_don2", ytop)
'        ytop = ytop + ydec
        ytop = ecr_frmq("Frm_par41", ssTitre3, "list_don3", ytop)
        ytop = ytop + ydec
        ytop = ecr_frm("Frm_par2", ssTitre4, "list_don4", ytop)
        ytop = ytop + ydec
        ytop = ecr_frm("Frm_par3", ssTitre5, "list_don5", ytop)
    Case Is = "pompe"
        ytop = ecr_frmd("Frm_par21", sstitre1, "list_don1", ytop)
        ytop = ytop + ydec
        ytop = ecr_frm("Frm_par2", ssTitre2, "list_don2", ytop)
        ytop = ytop + ydec
'        ytop = ecr_frm("Frm_par3", ssTitre3, "list_don3", ytop)
        ytop = ecr_frmq("Frm_par41", ssTitre3, "list_don3", ytop)
        ytop = ytop + ydec
        ytop = ecr_frm("Frm_par4", ssTitre4, "list_don4", ytop)
   End Select
    Picture1.Visible = False
    
End Sub

Private Sub Command1_Click()
'Dim Im As Picture
'Set im = Me.Image
On Error GoTo erreur:
Printer.TrackDefault = True
cdlg1.PrinterDefault = True
'Cdlg1.Flags = cdlPDPrintSetup Or cdlPDReturnDefault
'Cdlg1.ShowPrinter
cdlg1.Flags = cdlPDPrintSetup ' Or cdlPDReturnDC  'Or cdlPDReturnDefault
cdlg1.CancelError = True
While Printer.Orientation = cdlLandscape
    
'If Cdlg1.Orientation = cdlLandscape Then
    MsgBox "l'impression doit se faire en mode portrait", vbExclamation, _
        "Configuration imprimante"
'    Cdlg1.Orientation = cdlPortrait
    cdlg1.CancelError = True
   cdlg1.ShowPrinter
'End If
Wend

Command1.Visible = False
Cmd_annul.Visible = False
If Type1 = "deversoir" Or Type1 = "deversoiror" Or Type1 = "pompe" Then
    If Cmd_avant.Enabled Then
        Call Cmd_avant_Click
    End If
    Cmd_avant.Visible = False
    Cmd_apres.Visible = False
    Me.PrintForm
    Call Cmd_apres_Click
    Cmd_avant.Visible = False
    Cmd_apres.Visible = False
    Me.PrintForm
    Cmd_avant.Visible = True
    Cmd_apres.Visible = True
Else
    Me.PrintForm
End If
Command1.Visible = True
Cmd_annul.Visible = True
Exit Sub
erreur:
'bannul = True
Resume Next
End Sub

Public Sub paint_picture(ByRef pict1 As Picture)
Me.Picture1.Cls
Set Me.Picture1.Picture = pict1
'Me.Picture1.PaintPicture pict1, 0, 0  ', Picture1.Width, Picture1.Height, 0, 0
End Sub
Public Sub paint_picture2(ByRef pict2 As Picture)
Me.Picture2.Cls
Set Me.Picture2.Picture = pict2
End Sub
Private Function ecr_frm(ByVal nom_frm As String, ByVal Titre As String, ByVal nom_list As String, ByVal ytop As Double) As Double
  Dim liste() As Variant
  Dim i As Integer, nb As Integer
  Dim frm_par As Object
  Dim ok As Boolean
  ok = True
'  Dim lb_ldon As Collection
'  Dim lb_don As Label
'  Dim lb_unit As Label
  Select Case nom_frm
  Case Is = "Frm_par1"
        Set frm_par = Frm_par1
        Set lb_ldon = Lb_ldon1
        Set lb_don = Lb_don1
        Set lb_unit = Lb_unit1
   Case Is = "Frm_par2"
        Set frm_par = Frm_par2
        Set lb_ldon = Lb_ldon2
        Set lb_don = Lb_don2
        Set lb_unit = Lb_unit2
   Case Is = "Frm_par3"
        Set frm_par = Frm_par3
        Set lb_ldon = Lb_ldon3
        Set lb_don = Lb_don3
        Set lb_unit = Lb_unit3
   Case Is = "Frm_par4"
        Set frm_par = Frm_par4
        Set lb_ldon = Lb_ldon4
        Set lb_don = Lb_don4
        Set lb_unit = Lb_unit4
   Case Is = "Frm_par5"
        Set frm_par = Frm_par5
        Set lb_ldon = Lb_ldon5
        Set lb_don = Lb_don5
        Set lb_unit = Lb_unit5
   Case Is = "Frm_par6"
        Set frm_par = Frm_par6
        Set lb_ldon = Lb_ldon6
        Set lb_don = Lb_don6
        Set lb_unit = Lb_unit6
   Case Is = "Frm_par7"
        Set frm_par = Frm_par7
        Set lb_ldon = Lb_ldon7
        Set lb_don = Lb_don7
        Set lb_unit = Lb_unit7
  Case Else
       ok = False
End Select
If ok Then
      
    frm_par.Top = ytop
    frm_par.Visible = True
    frm_par.Caption = Titre
    liste = owner.fobjet.lect_list(nom_list)
    nb = UBound(liste)
'        Frm_par1.Height = (nb + 3) * (Lb_ldon1(0).Height + 10)
    lb_ldon(0).Caption = Trim(liste(0, 1))
 '       Lb_ldon1(0).Caption = Trim(liste(0, 1))
    lb_don(0).Caption = Trim(liste(0, 2)) + " "
    lb_unit(0).Caption = Trim(liste(0, 3))
    frm_par.Height = lb_unit(0).Top + lb_unit(0).Height + 50
    If nb > 0 Then
        For i = 1 To nb
            Load lb_ldon(i)
            Load lb_don(i)
            Load lb_unit(i)
            With lb_ldon(i)
                .Top = lb_ldon(i - 1).Top + lb_ldon(i - 1).Height + 10
                .Caption = Trim(liste(i, 1))
                .Visible = True
            End With
            With lb_don(i)
                .Top = lb_don(i - 1).Top + lb_don(i - 1).Height + 10
                .Caption = Trim(liste(i, 2)) + " "
                .Visible = True
            End With
            With lb_unit(i)
                .Top = lb_unit(i - 1).Top + lb_unit(i - 1).Height + 10
                .Caption = Trim(liste(i, 3))
                .Visible = True
                frm_par.Height = .Top + .Height + 50
            End With
        Next
    End If
    If Type1 = "pompe" Then
        frm_par.Height = frm_par.Height - 50
    End If
    ytop = frm_par.Top + frm_par.Height
    Set frm_par = Nothing
    Set lb_ldon = Nothing
    Set lb_don = Nothing
    Set lb_unit = Nothing
End If
ecr_frm = ytop

 End Function
Private Function unl_frm(ByVal nom_frm As String, ByVal nom_list As String)
  Dim i As Integer, nb As Integer
  Dim frm_par As Object
  Dim ok As Boolean
  ok = True
'  Dim lb_ldon As Collection
'  Dim lb_don As Label
'  Dim lb_unit As Label
  Select Case nom_frm
  Case Is = "Frm_par1"
        Set frm_par = Frm_par1
        Set lb_ldon = Lb_ldon1
        Set lb_don = Lb_don1
        Set lb_unit = Lb_unit1
   Case Is = "Frm_par2"
        Set frm_par = Frm_par2
        Set lb_ldon = Lb_ldon2
        Set lb_don = Lb_don2
        Set lb_unit = Lb_unit2
   Case Is = "Frm_par3"
        Set frm_par = Frm_par3
        Set lb_ldon = Lb_ldon3
        Set lb_don = Lb_don3
        Set lb_unit = Lb_unit3
   Case Is = "Frm_par4"
        Set frm_par = Frm_par4
        Set lb_ldon = Lb_ldon4
        Set lb_don = Lb_don4
        Set lb_unit = Lb_unit4
   Case Is = "Frm_par5"
        Set frm_par = Frm_par5
        Set lb_ldon = Lb_ldon5
        Set lb_don = Lb_don5
        Set lb_unit = Lb_unit5
   Case Is = "Frm_par6"
        Set frm_par = Frm_par6
        Set lb_ldon = Lb_ldon6
        Set lb_don = Lb_don6
        Set lb_unit = Lb_unit6
   Case Is = "Frm_par7"
        Set frm_par = Frm_par7
        Set lb_ldon = Lb_ldon7
        Set lb_don = Lb_don7
        Set lb_unit = Lb_unit7
  Case Else
       ok = False
End Select
If ok Then
      
    liste = owner.fobjet.lect_list(nom_list)
    nb = UBound(liste)
'        Frm_par1.Height = (nb + 3) * (Lb_ldon1(0).Height + 10)
    If nb > 0 Then
        For i = 1 To nb
            Unload lb_ldon(i)
            Unload lb_don(i)
            Unload lb_unit(i)
        Next
    End If
    Set frm_par = Nothing
    Set lb_ldon = Nothing
    Set lb_don = Nothing
    Set lb_unit = Nothing
End If

 End Function
Private Function ecr_frmd(ByVal nom_frm As String, ByVal Titre As String, ByVal nom_list As String, ByVal ytop As Double) As Double
  Dim liste() As Variant
  Dim i As Integer, nb As Integer
  Dim frm_par As Object
  Dim ok As Boolean
  ok = True
'  Dim lb_ldon As Collection
'  Dim lb_don As Label
'  Dim lb_unit As Label
  Select Case nom_frm
  Case Is = "Frm_par21"
        Set frm_par = Frm_par21
        Set lb_ldon = Lb_ldon21
        Set lb1_don = Lb_don11
        Set lb1_unit = Lb_unit11
        Set lb2_don = Lb_don21
        Set lb2_unit = Lb_unit21
   Case Is = "Frm_par22"
        Set frm_par = Frm_par22
        Set lb_ldon = Lb_ldon22
        Set lb1_don = Lb_don12
        Set lb1_unit = Lb_unit12
        Set lb2_don = Lb_don22
        Set lb2_unit = Lb_unit22
   Case Is = "Frm_par23"
        Set frm_par = Frm_par23
        Set lb_ldon = Lb_ldon23
        Set lb1_don = Lb_don13
        Set lb1_unit = Lb_unit13
        Set lb2_don = Lb_don23
        Set lb2_unit = Lb_unit23
  Case Else
       ok = False
End Select
If ok Then
      
    frm_par.Top = ytop
    frm_par.Visible = True
    frm_par.Caption = Titre
    If Trim(Titre) = "Détail des singularités" Then
        lb1_unit(0).Alignment = 1
    End If
    liste = owner.fobjet.lect_list(nom_list)
    nb = UBound(liste)
'        Frm_par1.Height = (nb + 3) * (Lb_ldon1(0).Height + 10)
    lb_ldon(0).Caption = Trim(liste(0, 1))
 '       Lb_ldon1(0).Caption = Trim(liste(0, 1))
    lb1_don(0).Caption = Trim(liste(0, 2)) + " "
    lb1_unit(0).Caption = Trim(liste(0, 3))
    lb2_don(0).Caption = Trim(liste(0, 4)) + " "
    lb2_unit(0).Caption = Trim(liste(0, 5))
    frm_par.Height = lb1_unit(0).Top + lb1_unit(0).Height + 50
    If nb > 0 Then
        For i = 1 To nb
            Load lb_ldon(i)
            Load lb1_don(i)
            Load lb1_unit(i)
            Load lb2_don(i)
            Load lb2_unit(i)
            With lb_ldon(i)
                .Top = lb_ldon(i - 1).Top + lb_ldon(i - 1).Height + 10
                .Caption = Trim(liste(i, 1))
                .Visible = True
            End With
            With lb1_don(i)
                .Top = lb1_don(i - 1).Top + lb1_don(i - 1).Height + 10
                .Caption = Trim(liste(i, 2)) + " "
                .Visible = True
            End With
            With lb1_unit(i)
               .Top = lb1_unit(i - 1).Top + lb1_unit(i - 1).Height + 10
                .Caption = Trim(liste(i, 3)) + " "
                .Visible = True
            End With
             With lb2_don(i)
               .Top = lb2_don(i - 1).Top + lb2_don(i - 1).Height + 10
                .Caption = Trim(liste(i, 4)) + " "
                .Visible = True
            End With
            With lb2_unit(i)
               .Top = lb2_unit(i - 1).Top + lb2_unit(i - 1).Height + 10
                .Caption = Trim(liste(i, 5)) + " "
                .Visible = True
                frm_par.Height = .Top + .Height + 50
            End With
       Next
    End If
    ytop = frm_par.Top + frm_par.Height
    Set frm_par = Nothing
    Set lb_ldon = Nothing
    Set lb1_don = Nothing
    Set lb1_unit = Nothing
    Set lb2_don = Nothing
    Set lb2_unit = Nothing
End If
ecr_frmd = ytop

 End Function
Private Function unl_frmd(ByVal nom_frm As String, ByVal nom_list As String)
  Dim liste() As Variant
  Dim i As Integer, nb As Integer
  Dim frm_par As Object
  Dim ok As Boolean
  ok = True
  Select Case nom_frm
  Case Is = "Frm_par21"
        Set frm_par = Frm_par21
        Set lb_ldon = Lb_ldon21
        Set lb1_don = Lb_don11
        Set lb1_unit = Lb_unit11
        Set lb2_don = Lb_don21
        Set lb2_unit = Lb_unit21
   Case Is = "Frm_par22"
        Set frm_par = Frm_par22
        Set lb_ldon = Lb_ldon22
        Set lb1_don = Lb_don12
        Set lb1_unit = Lb_unit12
        Set lb2_don = Lb_don22
        Set lb2_unit = Lb_unit22
   Case Is = "Frm_par23"
        Set frm_par = Frm_par23
        Set lb_ldon = Lb_ldon23
        Set lb1_don = Lb_don13
        Set lb1_unit = Lb_unit13
        Set lb2_don = Lb_don23
        Set lb2_unit = Lb_unit23
  Case Else
       ok = False
End Select
If ok Then
    liste = owner.fobjet.lect_list(nom_list)
    nb = UBound(liste)
    If nb > 0 Then
        For i = 1 To nb
            Unload lb_ldon(i)
            Unload lb1_don(i)
            Unload lb1_unit(i)
            Unload lb2_don(i)
           Unload lb2_unit(i)
       Next
    End If
    Set frm_par = Nothing
    Set lb_ldon = Nothing
    Set lb1_don = Nothing
    Set lb1_unit = Nothing
    Set lb2_don = Nothing
    Set lb2_unit = Nothing
End If

 End Function
Private Function ecr_frmt(ByVal nom_frm As String, ByVal Titre As String, ByVal nom_list As String, ByVal ytop As Double) As Double
  Dim liste() As Variant
  Dim i As Integer, nb As Integer
  Dim frm_par As Object
  Dim ok As Boolean
  ok = True
'  Dim lb_ldon As Collection
'  Dim lb_don As Label
'  Dim lb_unit As Label
  Select Case nom_frm
  Case Is = "Frm_par31"
        Set frm_par = Frm_par31
        Set lb1_ldon = Lb_ldon311
        Set lb1_don = Lb_don311
        Set lb1_unit = Lb_unit311
        Set lb2_ldon = Lb_ldon312
        Set lb2_don = Lb_don312
        Set lb2_unit = Lb_unit312
   Case Is = "Frm_par32"
        Set frm_par = Frm_par32
        Set lb1_ldon = Lb_ldon321
        Set lb1_don = Lb_don321
        Set lb1_unit = Lb_unit321
        Set lb2_ldon = Lb_ldon322
        Set lb2_don = Lb_don322
        Set lb2_unit = Lb_unit322
   Case Is = "Frm_par33"
        Set frm_par = Frm_par33
        Set lb1_ldon = Lb_ldon331
        Set lb1_don = Lb_don331
        Set lb1_unit = Lb_unit331
        Set lb2_ldon = Lb_ldon332
        Set lb2_don = Lb_don332
        Set lb2_unit = Lb_unit332
  Case Else
       ok = False
End Select
If ok Then
      
    frm_par.Top = ytop
    frm_par.Visible = True
    frm_par.Caption = Titre
'    If Type1 = "versant" Then
    If Not owner.fbassin Is Nothing Then

        liste = owner.fbassin.lect_list(nom_list)
    Else
        liste = owner.fobjet.lect_list(nom_list)
    End If
    nb = UBound(liste)
    lb1_ldon(0).Caption = Trim(liste(0, 1))
    lb1_don(0).Caption = Trim(liste(0, 2)) + " "
    lb1_unit(0).Caption = Trim(liste(0, 3))
    lb2_ldon(0).Caption = Trim(liste(0, 4))
    lb2_don(0).Caption = Trim(liste(0, 5)) + " "
    lb2_unit(0).Caption = Trim(liste(0, 6))
    frm_par.Height = lb1_unit(0).Top + lb1_unit(0).Height + 50
    If nb > 0 Then
        For i = 1 To nb
            Load lb1_ldon(i)
            Load lb1_don(i)
            Load lb1_unit(i)
            Load lb2_ldon(i)
            Load lb2_don(i)
            Load lb2_unit(i)
            With lb1_ldon(i)
                .Top = lb1_ldon(i - 1).Top + lb1_ldon(i - 1).Height + 10
                .Caption = Trim(liste(i, 1))
                .Visible = True
            End With
            With lb1_don(i)
                .Top = lb1_don(i - 1).Top + lb1_don(i - 1).Height + 10
                .Caption = Trim(liste(i, 2)) + " "
                .Visible = True
            End With
            With lb1_unit(i)
                .Top = lb1_unit(i - 1).Top + lb1_unit(i - 1).Height + 10
                .Caption = Trim(liste(i, 3)) + " "
                .Visible = True
            End With
            With lb2_ldon(i)
                .Top = lb2_ldon(i - 1).Top + lb2_ldon(i - 1).Height + 10
                .Caption = Trim(liste(i, 4))
                .Visible = True
            End With
             With lb2_don(i)
                .Top = lb2_don(i - 1).Top + lb2_don(i - 1).Height + 10
                .Caption = Trim(liste(i, 5)) + " "
                .Visible = True
            End With
            With lb2_unit(i)
                .Top = lb2_unit(i - 1).Top + lb2_unit(i - 1).Height + 10
                .Caption = Trim(liste(i, 6)) + " "
                .Visible = True
                frm_par.Height = .Top + .Height + 50
            End With
       Next
    End If
    ytop = frm_par.Top + frm_par.Height
    Set frm_par = Nothing
    Set lb1_ldon = Nothing
    Set lb1_don = Nothing
    Set lb1_unit = Nothing
    Set lb2_ldon = Nothing
    Set lb2_don = Nothing
    Set lb2_unit = Nothing
End If
ecr_frmt = ytop

 End Function
Private Function unl_frmt(ByVal nom_frm As String, ByVal nom_list As String)
  Dim liste() As Variant
  Dim i As Integer, nb As Integer
  Dim frm_par As Object
  Dim ok As Boolean
  ok = True
  Select Case nom_frm
  Case Is = "Frm_par31"
        Set frm_par = Frm_par31
        Set lb1_ldon = Lb_ldon311
        Set lb1_don = Lb_don311
        Set lb1_unit = Lb_unit311
        Set lb2_ldon = Lb_ldon312
        Set lb2_don = Lb_don312
        Set lb2_unit = Lb_unit312
   Case Is = "Frm_par32"
        Set frm_par = Frm_par32
        Set lb1_ldon = Lb_ldon321
        Set lb1_don = Lb_don321
        Set lb1_unit = Lb_unit321
        Set lb2_ldon = Lb_ldon322
        Set lb2_don = Lb_don322
        Set lb2_unit = Lb_unit322
   Case Is = "Frm_par33"
        Set frm_par = Frm_par33
        Set lb1_ldon = Lb_ldon331
        Set lb1_don = Lb_don331
        Set lb1_unit = Lb_unit331
        Set lb2_ldon = Lb_ldon332
        Set lb2_don = Lb_don332
        Set lb2_unit = Lb_unit332
  Case Else
       ok = False
End Select
If ok Then
      
'    If Type1 = "versant" Then
    If Not owner.fbassin Is Nothing Then
        liste = owner.fbassin.lect_list(nom_list)
    Else
        liste = owner.fobjet.lect_list(nom_list)
    End If
    nb = UBound(liste)
    If nb > 0 Then
        For i = 1 To nb
            Unload lb1_ldon(i)
            Unload lb1_don(i)
            Unload lb1_unit(i)
            Unload lb2_ldon(i)
            Unload lb2_don(i)
            Unload lb2_unit(i)
       Next
    End If
    ytop = frm_par.Top + frm_par.Height
    Set frm_par = Nothing
    Set lb1_ldon = Nothing
    Set lb1_don = Nothing
    Set lb1_unit = Nothing
    Set lb2_ldon = Nothing
    Set lb2_don = Nothing
    Set lb2_unit = Nothing
End If

 End Function
Private Function ecr_frmq(ByVal nom_frm As String, ByVal Titre As String, ByVal nom_list As String, ByVal ytop As Double) As Double
  Dim liste() As Variant
  Dim i As Integer, nb As Integer
  Dim frm_par As Object
  Dim ok As Boolean
  ok = True
'  Dim lb_ldon As Collection
'  Dim lb_don As Label
'  Dim lb_unit As Label
  Select Case nom_frm
  Case Is = "Frm_par41"
        Set frm_par = Frm_par41
        Set lb1_ldon = Lb_ldon411
        Set lb1_don = Lb_don411
        Set lb1_unit = Lb_unit411
        Set lb2_don = Lb_don412
        Set lb2_unit = Lb_unit412
        Set lb3_don = Lb_don413
        Set lb3_unit = Lb_unit413
   Case Is = "Frm_par42"
        Set frm_par = Frm_par42
        Set lb1_ldon = Lb_ldon421
        Set lb1_don = Lb_don421
        Set lb1_unit = Lb_unit421
        Set lb2_don = Lb_don422
        Set lb2_unit = Lb_unit422
        Set lb3_don = Lb_don423
        Set lb3_unit = Lb_unit423
'   Case Is = "Frm_par43"
'        Set frm_par = Frm_par33
'        Set lb1_ldon = Lb_ldon331
'        Set lb1_don = Lb_don331
'        Set lb1_unit = Lb_unit331
'        Set lb2_ldon = Lb_ldon332
'        Set lb2_don = Lb_don332
'        Set lb2_unit = Lb_unit332
  Case Else
       ok = False
End Select
If ok Then
      
    frm_par.Top = ytop
    frm_par.Visible = True
    frm_par.Caption = Titre
'    If Type1 = "versant" Then
    If Not owner.fbassin Is Nothing Then
        liste = owner.fbassin.lect_list(nom_list)
    Else
        liste = owner.fobjet.lect_list(nom_list)
    End If
    nb = UBound(liste)
    lb1_ldon(0).Caption = Trim(liste(0, 1))
    lb1_don(0).Caption = Trim(liste(0, 2)) + " "
                If Type1 = "pompe" Then
                    lb1_don(0).Alignment = 2
                End If
    lb1_unit(0).Caption = Trim(liste(0, 3))
                If Type1 = "pompe" Then
                    lb1_unit(0).Alignment = 2
                End If
    lb2_don(0).Caption = Trim(liste(0, 4)) + " "
                If Type1 = "pompe" Then
                    lb2_don(0).Alignment = 2
                End If
    lb2_unit(0).Caption = Trim(liste(0, 5))
                If Type1 = "pompe" Then
                    lb2_unit(0).Alignment = 2
                End If
    lb3_don(0).Caption = Trim(liste(0, 6)) + " "
                If Type1 = "pompe" Then
                    lb3_don(0).Alignment = 2
                End If
    lb3_unit(0).Caption = Trim(liste(0, 7))
                If Type1 = "pompe" Then
                    lb3_unit(0).Alignment = 2
                End If
    frm_par.Height = lb1_unit(0).Top + lb1_unit(0).Height + 50
    If nb > 0 Then
        For i = 1 To nb
            Load lb1_ldon(i)
            Load lb1_don(i)
            Load lb1_unit(i)
            Load lb2_don(i)
            Load lb2_unit(i)
            Load lb3_don(i)
            Load lb3_unit(i)
           With lb1_ldon(i)
                .Top = lb1_ldon(i - 1).Top + lb1_ldon(i - 1).Height + 10
                .Caption = Trim(liste(i, 1))
                .Visible = True
            End With
            With lb1_don(i)
                If Type1 = "pompe" Then
                    .Alignment = 2
                End If
               .Top = lb1_don(i - 1).Top + lb1_don(i - 1).Height + 10
                .Caption = Trim(liste(i, 2)) + " "
                .Visible = True
            End With
            With lb1_unit(i)
                If Type1 = "pompe" Then
                    .Alignment = 2
                End If
                .Top = lb1_unit(i - 1).Top + lb1_unit(i - 1).Height + 10
                .Caption = Trim(liste(i, 3)) + " "
                .Visible = True
            End With
             With lb2_don(i)
                If Type1 = "pompe" Then
                    .Alignment = 2
                End If
                .Top = lb2_don(i - 1).Top + lb2_don(i - 1).Height + 10
                .Caption = Trim(liste(i, 4)) + " "
                .Visible = True
            End With
            With lb2_unit(i)
                If Type1 = "pompe" Then
                    .Alignment = 2
                End If
                .Top = lb2_unit(i - 1).Top + lb2_unit(i - 1).Height + 10
                .Caption = Trim(liste(i, 5)) + " "
                .Visible = True
                frm_par.Height = .Top + .Height + 50
            End With
              With lb3_don(i)
                If Type1 = "pompe" Then
                    .Alignment = 2
                End If
                .Top = lb3_don(i - 1).Top + lb3_don(i - 1).Height + 10
                .Caption = Trim(liste(i, 6)) + " "
                .Visible = True
            End With
            With lb3_unit(i)
                If Type1 = "pompe" Then
                    .Alignment = 2
                End If
                .Top = lb3_unit(i - 1).Top + lb3_unit(i - 1).Height + 10
                .Caption = Trim(liste(i, 7)) + " "
                .Visible = True
                frm_par.Height = .Top + .Height + 50
            End With
      Next
    End If
    ytop = frm_par.Top + frm_par.Height
    Set frm_par = Nothing
    Set lb1_ldon = Nothing
    Set lb1_don = Nothing
    Set lb1_unit = Nothing
    Set lb2_don = Nothing
    Set lb2_unit = Nothing
    Set lb3_don = Nothing
    Set lb3_unit = Nothing
End If
ecr_frmq = ytop

 End Function
Private Function unl_frmq(ByVal nom_frm As String, ByVal nom_list As String)
  Dim liste() As Variant
  Dim i As Integer, nb As Integer
  Dim frm_par As Object
  Dim ok As Boolean
  ok = True
  Select Case nom_frm
  Case Is = "Frm_par41"
        Set frm_par = Frm_par41
        Set lb1_ldon = Lb_ldon411
        Set lb1_don = Lb_don411
        Set lb1_unit = Lb_unit411
        Set lb2_don = Lb_don412
        Set lb2_unit = Lb_unit412
        Set lb3_don = Lb_don413
        Set lb3_unit = Lb_unit413
   Case Is = "Frm_par42"
        Set frm_par = Frm_par42
        Set lb1_ldon = Lb_ldon421
        Set lb1_don = Lb_don421
        Set lb1_unit = Lb_unit421
        Set lb2_don = Lb_don422
        Set lb2_unit = Lb_unit422
        Set lb3_don = Lb_don423
        Set lb3_unit = Lb_unit423
'   Case Is = "Frm_par43"
'        Set frm_par = Frm_par33
'        Set lb1_ldon = Lb_ldon331
'        Set lb1_don = Lb_don331
'        Set lb1_unit = Lb_unit331
'        Set lb2_ldon = Lb_ldon332
'        Set lb2_don = Lb_don332
'        Set lb2_unit = Lb_unit332
  Case Else
       ok = False
End Select
If ok Then
      
'    If Type1 = "versant" Then
    If Not owner.fbassin Is Nothing Then
        liste = owner.fbassin.lect_list(nom_list)
    Else
        liste = owner.fobjet.lect_list(nom_list)
    End If
    nb = UBound(liste)
    If nb > 0 Then
        For i = 1 To nb
            Unload lb1_ldon(i)
            Unload lb1_don(i)
            Unload lb1_unit(i)
            Unload lb2_don(i)
            Unload lb2_unit(i)
            Unload lb3_don(i)
            Unload lb3_unit(i)
      Next
    End If
    Set frm_par = Nothing
    Set lb1_ldon = Nothing
    Set lb1_don = Nothing
    Set lb1_unit = Nothing
    Set lb2_don = Nothing
    Set lb2_unit = Nothing
    Set lb3_don = Nothing
    Set lb3_unit = Nothing
End If

 End Function

Private Sub Form_Activate()
    owner.fdessin.Enabled = False
    If Not owner.fbassin Is Nothing Then
         owner.fbassin.Enabled = False
    Else
        owner.fobjet.Enabled = False
    End If
    owner.fcom.Enabled = False

End Sub

Private Sub Form_Load()
Dim ytop As Double, ydec As Double
Dim ok_sing As Boolean
Dim nometude As String
nometude = "Etude : " + nom_etude
Me.MousePointer = 11
If Printers.count > 0 Then
    Command1.Enabled = True
    
Else
    Command1.Enabled = False
End If

form_top = Me.Top
    Set owner = MDIFrm_menu.rec_owner
'    owner.fdessin.Enabled = False
''    If Type1 = "versant" Then
'    If Not owner.fbassin Is Nothing Then
'         owner.fbassin.Enabled = False
'    Else
'        owner.fobjet.Enabled = False
'    End If
'    owner.fcom.Enabled = False
    Call ini_form1
    Etude.Caption = nometude
    Dossier.Caption = "Dossier : " + nom_fich_edit
Select Case Type1
    Case Is = "decant"
        ytop = 4000
        ydec = 200
        Titre(0).Caption = titre1
        Lb_titre.Caption = nomobjet
        ytop = ecr_frm("Frm_par1", sstitre1, "list_don1", ytop)
        ytop = ytop + ydec
        ytop = ecr_frm("Frm_par2", ssTitre2, "list_int1", ytop)
        ytop = ytop + ydec
        ytop = ecr_frm("Frm_par3", ssTitre3, "list_resu1", ytop)
        ytop = ytop + 1500
        Picture1.Top = ytop
     Case Is = "stockage"
        ytop = 3500
        ydec = 200
        Titre(0).Caption = titre1
        Lb_titre.Caption = nomobjet
        ytop = ecr_frm("Frm_par1", sstitre1, "list_don1", ytop)
        ytop = ytop + ydec
        ytop = ecr_frm("Frm_par2", ssTitre2, "list_int1", ytop)
        ytop = ytop + ydec
        ytop = ecr_frm("Frm_par3", ssTitre3, "list_resu1", ytop)
        ytop = ytop + 500
        Picture1.Top = ytop
    Case Is = "chute"
        ytop = 3500
        ydec = 200
        Titre(0).Caption = titre1
        Lb_titre.Caption = nomobjet
        ytop = ecr_frmd("Frm_par21", sstitre1, "list_don1", ytop)
        ytop = ytop + ydec
        ytop = ecr_frm("Frm_par1", ssTitre2, "list_don2", ytop)
        ytop = ytop + ydec
        ytop = ecr_frmd("Frm_par22", ssTitre3, "list_int1", ytop)
        ytop = ytop + ydec
        ytop = ecr_frm("Frm_par2", ssTitre4, "list_resu1", ytop)
        ytop = ytop + ydec '1000
        Picture1.Top = ytop
        Picture1.Height = 4500
    Case Is = "conduite"
        ytop = 4000
        ydec = 200
        Titre(0).Caption = titre1
        Lb_titre.Caption = nomobjet
        ytop = ecr_frm("Frm_par1", sstitre1, "list_don1", ytop)
        If Trim(ssTitre3) <> "" Then
        ytop = ytop + ydec
        ytop = ecr_frm("Frm_par2", ssTitre2, "list_don2", ytop)
        ytop = ytop + ydec
        ytop = ecr_frm("Frm_par3", ssTitre3, "list_int1", ytop)
        End If
        ytop = ytop + 1500
        Picture1.Top = ytop
    Case Is = "siphon"
        ytop = 3200
        ydec = 200
        ok_sing = False
        Titre(0).Caption = titre1
        Lb_titre.Caption = nomobjet
        ytop = ecr_frmd("Frm_par21", sstitre1, "list_don1", ytop)
        ytop = ytop + ydec
        ytop = ecr_frm("Frm_par1", ssTitre2, "list_don2", ytop)
        ytop = ytop + ydec
        If Trim(ssTitre3) <> "" Then
            ok_sing = True
            ytop = ecr_frmd("Frm_par22", ssTitre3, "list_don3", ytop)
            ytop = ytop + ydec
        End If
        ytop = ecr_frm("Frm_par2", ssTitre4, "list_int1", ytop)
        If ok_sing Then
            ytop = ytop + 300
        Else
            ytop = ytop + 1000
        End If
        Picture1.Top = ytop
         Picture1.Height = 4200
    Case Is = "retention"
        ytop = 3000 '3200
        ydec = 100
        Titre(0).Caption = titre1
        Lb_titre.Caption = nomobjet
        ytop = ecr_frm("Frm_par1", sstitre1, "list_don1", ytop)
        ytop = ytop + ydec
        ytop = ecr_frm("Frm_par2", ssTitre2, "list_int1", ytop)
        ytop = ytop + 50
        Picture2.Visible = True
        Picture2.Top = ytop
        ytop = Picture2.Top + Picture2.Height
        Lb_des2b.Visible = True
        Lb_des2b.Caption = Me.des2_titrb
        Lb_des2b.Top = ytop
        ytop = ytop + Lb_des2b.Height
        ytop = ecr_frm("Frm_par3", ssTitre3, "list_resu1", ytop)
        ytop = ytop + 50
        Picture1.Top = ytop
    Case Is = "versant"
        ytop = 3500
        ydec = 200
        Titre(0).Caption = titre1
        Lb_titre.Caption = nomobjet
        ytop = ecr_frmt("Frm_par31", sstitre1, "list_don1", ytop)
        ytop = ytop + ydec
        ytop = ecr_frmt("Frm_par32", ssTitre2, "list_don2", ytop)
        ytop = ytop + ydec
        ytop = ecr_frmt("Frm_par33", ssTitre3, "list_resu1", ytop)
        ytop = ytop + 1000
        Lb_des1h.Visible = True
        Lb_des1h.Caption = Me.des1_titrh
        Lb_des1h.Top = ytop
        ytop = ytop + Lb_des1h.Height
        Picture1.Top = ytop
       ytop = Picture1.Top + Picture1.Height
        Lb_des1b.Visible = True
        Lb_des1b.Caption = Me.des1_titrb
        Lb_des1b.Top = ytop
        ytop = ytop + Lb_des1b.Height
    Case Is = "deversoir"
        Call Cmd_avant_Click
    Case Is = "deversoiror"
        Call Cmd_avant_Click
    Case Is = "pompe"
        Call Cmd_avant_Click
End Select
Me.MousePointer = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Me.Top = form_top
If Not owner.fcom Is Nothing Then
    owner.fcom.Enabled = True
    owner.fcom.retailler
End If

If Not owner.fdessin Is Nothing Then
    owner.fdessin.retailler
   owner.fdessin.Enabled = True
End If
'If Type1 = "versant" Then
    If Not owner.fbassin Is Nothing Then
        owner.fbassin.retailler
        owner.fbassin.Enabled = True
    End If
'Else
    If Not owner.fobjet Is Nothing Then
        owner.fobjet.retailler
        owner.fobjet.Enabled = True
    End If
'End If
End Sub
