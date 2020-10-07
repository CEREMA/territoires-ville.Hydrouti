VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Frm_pompe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hydraulique Pompe"
   ClientHeight    =   4305
   ClientLeft      =   150
   ClientTop       =   615
   ClientWidth     =   9825
   Icon            =   "Frm_pompe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   9825
   Begin VB.TextBox Tb_Geom 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   3720
      TabIndex        =   145
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   850
   End
   Begin VB.Frame Frame5 
      Height          =   4050
      Left            =   5450
      TabIndex        =   77
      Top             =   0
      Width           =   4245
      Begin VB.TextBox Tb_Singul 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   159
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   3480
         Width           =   850
      End
      Begin VB.TextBox Tb_Hmt 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   155
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   3720
         Width           =   850
      End
      Begin VB.TextBox Tb_Tsejh 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   133
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   3180
         Width           =   850
      End
      Begin VB.TextBox Tb_Vmy 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   130
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   2865
         Width           =   850
      End
      Begin VB.TextBox Tb_Nbcyc 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         MaxLength       =   6
         TabIndex        =   128
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   2520
         Width           =   645
      End
      Begin VB.TextBox Tb_Tvidange 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   125
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Tb_T1cyc 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         MaxLength       =   7
         TabIndex        =   124
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   2520
         Width           =   700
      End
      Begin VB.TextBox Tb_nrdph 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         MaxLength       =   5
         TabIndex        =   122
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   2160
         Width           =   850
      End
      Begin VB.TextBox Tb_denivr 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   156
         Text            =   "Text1"
         Top             =   1560
         Width           =   850
      End
      Begin VB.TextBox Tb_vurba 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   113
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   1875
         Width           =   850
      End
      Begin VB.TextBox Tb_Qpomp 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   152
         Top             =   240
         Width           =   850
      End
      Begin VB.TextBox Tb_Qpompc 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   240
         Width           =   850
      End
      Begin VB.TextBox Tb_Drflt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   154
         Top             =   550
         Width           =   850
      End
      Begin VB.TextBox Tb_VitRflt 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   860
         Width           =   850
      End
      Begin VB.TextBox Tb_Jmpkm 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   1200
         Width           =   850
      End
      Begin VB.Label Lb_unit_Singul 
         Caption         =   "m"
         Height          =   285
         Left            =   3720
         TabIndex        =   161
         Top             =   3510
         Width           =   465
      End
      Begin VB.Label Lb_int_Singul 
         Caption         =   "Perte de charge singulière"
         Height          =   285
         Left            =   120
         TabIndex        =   160
         Top             =   3510
         Width           =   2175
      End
      Begin VB.Label Lb_unit_Hmt 
         Caption         =   "m"
         Height          =   285
         Left            =   3720
         TabIndex        =   158
         Top             =   3750
         Width           =   465
      End
      Begin VB.Label Lb_int_Hmt 
         Caption         =   "Hauteur manométrique totale"
         Height          =   285
         Left            =   120
         TabIndex        =   157
         Top             =   3750
         Width           =   2175
      End
      Begin VB.Label Lb_int_T1cyc 
         Caption         =   "Cycle"
         Height          =   285
         Left            =   2040
         TabIndex        =   123
         Top             =   2550
         Width           =   495
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   4080
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Label Lb_unit_Tsejh 
         Caption         =   "h"
         Height          =   285
         Left            =   3675
         TabIndex        =   138
         Top             =   3210
         Width           =   225
      End
      Begin VB.Label Lb_unit_Vmy 
         Caption         =   "m/s"
         Height          =   285
         Left            =   3675
         TabIndex        =   137
         Top             =   2895
         Width           =   465
      End
      Begin VB.Label Lb_unit_T1cyc 
         Caption         =   "h"
         Height          =   285
         Left            =   3240
         TabIndex        =   136
         Top             =   2550
         Width           =   225
      End
      Begin VB.Label Lb_unit_tvidange 
         Caption         =   "h"
         Height          =   285
         Left            =   1920
         TabIndex        =   135
         Top             =   2550
         Width           =   195
      End
      Begin VB.Label Lb_unit_Jmpkm 
         Caption         =   "m/km"
         Height          =   285
         Left            =   3680
         TabIndex        =   134
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Lb_int_Tsejh 
         Caption         =   "Temps de séjour"
         Height          =   285
         Left            =   120
         TabIndex        =   132
         Top             =   3210
         Width           =   2055
      End
      Begin VB.Label Lb_int_Vmy 
         Caption         =   "Vitesse moyenne d'écoulement"
         Height          =   285
         Left            =   120
         TabIndex        =   129
         Top             =   2895
         Width           =   2295
      End
      Begin VB.Label Lb_int_Nbcyc 
         Caption         =   "NbC"
         Height          =   285
         Left            =   3720
         TabIndex        =   127
         Top             =   2280
         Width           =   330
      End
      Begin VB.Label Lb_int_tvidange 
         Caption         =   "Tps vidange"
         Height          =   285
         Left            =   120
         TabIndex        =   126
         Top             =   2550
         Width           =   975
      End
      Begin VB.Label Lb_int_nrdph 
         Caption         =   "Nb réel de démarrage(s) /h"
         Height          =   285
         Left            =   120
         TabIndex        =   121
         Top             =   2190
         Width           =   2295
      End
      Begin VB.Label Lb_unit_denivr 
         Caption         =   "m"
         Height          =   285
         Left            =   3675
         TabIndex        =   120
         Top             =   1590
         Width           =   465
      End
      Begin VB.Label Lb_int_denivr 
         Caption         =   "Tranche de pompage retenue"
         Height          =   285
         Left            =   120
         TabIndex        =   119
         Top             =   1590
         Width           =   2175
      End
      Begin VB.Label Lb_unit_vurba 
         Caption         =   "m3"
         Height          =   255
         Left            =   3675
         TabIndex        =   114
         Top             =   1905
         Width           =   420
      End
      Begin VB.Label Lb_int_vurba 
         Caption         =   "Volume utile de la bâche"
         Height          =   285
         Left            =   120
         TabIndex        =   112
         Top             =   1905
         Width           =   2295
      End
      Begin VB.Label Lb_int_Qpomp 
         Caption         =   "Débit de pompage"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   86
         Top             =   270
         Width           =   1395
      End
      Begin VB.Label Lb_unit_Qpompc 
         Caption         =   "m3/h"
         Height          =   255
         Index           =   1
         Left            =   3680
         TabIndex        =   85
         Top             =   270
         Width           =   435
      End
      Begin VB.Label Lb_unit_Qpomp 
         Caption         =   "l/s  "
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   84
         Top             =   270
         Width           =   195
      End
      Begin VB.Label Lb_Int_Drflt 
         Caption         =   "Canalisation "
         Height          =   285
         Left            =   120
         TabIndex        =   83
         Top             =   580
         Width           =   1095
      End
      Begin VB.Label Lb_Unit_Drflt 
         Caption         =   "mm"
         Height          =   285
         Left            =   3680
         TabIndex        =   82
         Top             =   580
         Width           =   315
      End
      Begin VB.Label Lb_materiau 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         TabIndex        =   81
         Top             =   580
         Width           =   1035
      End
      Begin VB.Label Lb_int_Vitrflt 
         Caption         =   "Vitesse en régime  permanent"
         Height          =   285
         Left            =   120
         TabIndex        =   80
         Top             =   890
         Width           =   2355
      End
      Begin VB.Label Lb_unit_Vitrflt 
         Caption         =   "m/s"
         Height          =   285
         Left            =   3680
         TabIndex        =   79
         Top             =   890
         Width           =   465
      End
      Begin VB.Label Lb_int_Jmpkm 
         Caption         =   "Perte de charges  linéaires"
         Height          =   285
         Left            =   120
         TabIndex        =   78
         Top             =   1200
         Width           =   2175
      End
   End
   Begin VB.TextBox Tb_titre 
      Height          =   300
      Left            =   6480
      MaxLength       =   30
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3915
      Left            =   120
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   120
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   6906
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   794
      TabCaption(0)   =   "Débits Caract."
      TabPicture(0)   =   "Frm_pompe.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lb_unit_Qpomp(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Lb_unit_Qpompc(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Lb_int_Qpomp(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Tb_Qpompc(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Tb_Qpomp(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Données géométr."
      TabPicture(1)   =   "Frm_pompe.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frm_geom2"
      Tab(1).Control(1)=   "Frm_geom1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Points singul."
      TabPicture(2)   =   "Frm_pompe.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Données tech."
      TabPicture(3)   =   "Frm_pompe.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Lb_resu"
      Tab(3).Control(1)=   "Lb_int_Vutba"
      Tab(3).Control(2)=   "Lb_int_denivt"
      Tab(3).Control(3)=   "Lb_unit_vutba"
      Tab(3).Control(4)=   "Lb_unit_denivt"
      Tab(3).Control(5)=   "Frame6"
      Tab(3).Control(6)=   "Tb_Vutba"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Frm_bache"
      Tab(3).Control(8)=   "Tb_denivt"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Frame8"
      Tab(3).ControlCount=   10
      Begin VB.Frame Frame8 
         Height          =   975
         Left            =   -74760
         TabIndex        =   146
         Top             =   2760
         Width           =   4700
         Begin VB.TextBox Tb_denivbas 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2880
            MaxLength       =   8
            TabIndex        =   150
            Top             =   560
            Width           =   855
         End
         Begin VB.TextBox Tb_denivhau 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2880
            MaxLength       =   8
            TabIndex        =   147
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Lb_int_denivbas 
            Caption         =   "Garde au fond"
            Height          =   285
            Left            =   120
            TabIndex        =   153
            Top             =   590
            Width           =   2220
         End
         Begin VB.Label Lb_unit_denivbas 
            Caption         =   "m"
            Height          =   255
            Left            =   3840
            TabIndex        =   151
            Top             =   590
            Width           =   300
         End
         Begin VB.Label Lb_int_denivhau 
            Caption         =   "Garde à l'égout"
            Height          =   285
            Left            =   120
            TabIndex        =   149
            Top             =   270
            Width           =   2220
         End
         Begin VB.Label Lb_unit_denivhau 
            Caption         =   "m"
            Height          =   255
            Left            =   3840
            TabIndex        =   148
            Top             =   270
            Width           =   300
         End
      End
      Begin VB.TextBox Tb_denivt 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71880
         MaxLength       =   8
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   2480
         Width           =   850
      End
      Begin VB.Frame Frm_bache 
         Caption         =   "Section de la bâche"
         Height          =   975
         Left            =   -74760
         TabIndex        =   95
         Top             =   1440
         Width           =   4700
         Begin VB.TextBox Tb_larg 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2880
            MaxLength       =   8
            TabIndex        =   106
            Top             =   560
            Width           =   850
         End
         Begin VB.TextBox Tb_long 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2880
            MaxLength       =   8
            TabIndex        =   102
            Top             =   240
            Width           =   850
         End
         Begin VB.OptionButton Opt_sect_ba 
            Caption         =   "Rectangulaire"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Opt_sect_ba 
            Caption         =   "Circulaire"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   96
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox Tb_diam 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2880
            MaxLength       =   8
            TabIndex        =   99
            Top             =   240
            Width           =   850
         End
         Begin VB.Label Lb_unit_long 
            Caption         =   "m"
            Height          =   285
            Left            =   3840
            TabIndex        =   110
            Top             =   270
            Width           =   300
         End
         Begin VB.Label Lb_unit_larg 
            Caption         =   "m"
            Height          =   285
            Left            =   3840
            TabIndex        =   108
            Top             =   590
            Width           =   300
         End
         Begin VB.Label Lb_int_larg 
            Caption         =   "Largeur"
            Height          =   285
            Left            =   1920
            TabIndex        =   103
            Top             =   590
            Width           =   900
         End
         Begin VB.Label Lb_int_long 
            Caption         =   "Longueur"
            Height          =   285
            Left            =   1920
            TabIndex        =   101
            Top             =   240
            Width           =   900
         End
         Begin VB.Label Lb_unit_diam 
            Caption         =   "m"
            Height          =   285
            Left            =   3840
            TabIndex        =   100
            Top             =   270
            Width           =   300
         End
         Begin VB.Label Lb_int_diam 
            Caption         =   "Diamétre"
            Height          =   285
            Left            =   1920
            TabIndex        =   98
            Top             =   270
            Width           =   900
         End
      End
      Begin VB.TextBox Tb_Vutba 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   -71880
         MaxLength       =   8
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   1150
         Width           =   850
      End
      Begin VB.Frame Frame6 
         Height          =   615
         Left            =   -74760
         TabIndex        =   87
         Top             =   480
         Width           =   4700
         Begin VB.TextBox Tb_Ntdph 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3900
            MaxLength       =   3
            TabIndex        =   90
            Top             =   240
            Width           =   400
         End
         Begin VB.TextBox Tb_Nbpom 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1500
            MaxLength       =   3
            TabIndex        =   88
            Top             =   240
            Width           =   400
         End
         Begin VB.Label Lb_Int_Ntdph 
            Caption         =   "Nb de démarrage(s) /h"
            Height          =   285
            Left            =   2000
            TabIndex        =   91
            Top             =   270
            Width           =   1800
         End
         Begin VB.Label Lb_Int_Nbpom 
            Caption         =   "Nb de pompe(s)"
            Height          =   285
            Left            =   150
            TabIndex        =   89
            Top             =   270
            Width           =   1200
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   120
         TabIndex        =   56
         Top             =   480
         Width           =   5055
         Begin VB.TextBox Tb_debit 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   2160
            MaxLength       =   8
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   2280
            Width           =   850
         End
         Begin VB.TextBox Tb_debit 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   2160
            MaxLength       =   8
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   1920
            Width           =   850
         End
         Begin VB.TextBox Tb_debit 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   2150
            MaxLength       =   8
            TabIndex        =   3
            Top             =   1440
            Width           =   850
         End
         Begin VB.TextBox Tb_debit 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   2150
            MaxLength       =   8
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   1080
            Width           =   850
         End
         Begin VB.TextBox Tb_FPointe 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2160
            MaxLength       =   8
            TabIndex        =   1
            Top             =   720
            Width           =   850
         End
         Begin VB.TextBox Tb_debit 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   2160
            MaxLength       =   8
            TabIndex        =   0
            Top             =   300
            Width           =   850
         End
         Begin VB.TextBox Tb_Debitc 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   3600
            MaxLength       =   8
            TabIndex        =   104
            TabStop         =   0   'False
            Text            =   "Text1"
            Top             =   360
            Width           =   850
         End
         Begin VB.TextBox Tb_Debitc 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   3600
            MaxLength       =   8
            TabIndex        =   105
            TabStop         =   0   'False
            Text            =   "Text1"
            Top             =   1080
            Width           =   850
         End
         Begin VB.TextBox Tb_Debitc 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   3600
            MaxLength       =   8
            TabIndex        =   107
            TabStop         =   0   'False
            Text            =   "Text1"
            Top             =   1440
            Width           =   850
         End
         Begin VB.TextBox Tb_Debitc 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   3600
            MaxLength       =   8
            TabIndex        =   109
            TabStop         =   0   'False
            Text            =   "Text1"
            Top             =   1860
            Width           =   850
         End
         Begin VB.TextBox Tb_Debitc 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   3600
            MaxLength       =   8
            TabIndex        =   111
            TabStop         =   0   'False
            Text            =   "Text1"
            Top             =   2250
            Width           =   850
         End
         Begin VB.Label Lb_int_debit 
            Caption         =   "Débit de pointe de temps sec"
            Height          =   255
            Index           =   4
            Left            =   50
            TabIndex        =   72
            Top             =   2280
            Width           =   2100
         End
         Begin VB.Label Lb_int_debit 
            Caption         =   "Débit moyen de temps sec"
            Height          =   255
            Index           =   3
            Left            =   50
            TabIndex        =   71
            Top             =   1890
            Width           =   2100
         End
         Begin VB.Label Lb_int_debit 
            Caption         =   "Débit des eaux parasites"
            Height          =   255
            Index           =   2
            Left            =   50
            TabIndex        =   70
            Top             =   1500
            Width           =   2100
         End
         Begin VB.Label Lb_int_debit 
            Caption         =   "Débit pointe des eaux usées"
            Height          =   255
            Index           =   1
            Left            =   50
            TabIndex        =   69
            Top             =   1110
            Width           =   2100
         End
         Begin VB.Label Lb_int_Fpointe 
            Caption         =   "Facteur de pointe"
            Height          =   255
            Left            =   50
            TabIndex        =   68
            Top             =   720
            Width           =   2100
         End
         Begin VB.Label Lb_int_debit 
            Caption         =   "Débit moyen des eaux usées"
            Height          =   255
            Index           =   0
            Left            =   45
            TabIndex        =   67
            Top             =   360
            Width           =   2100
         End
         Begin VB.Label Lb_unit_debit 
            Caption         =   "l/s soit "
            Height          =   255
            Index           =   0
            Left            =   3045
            LinkItem        =   " "
            TabIndex        =   66
            Top             =   330
            Width           =   495
         End
         Begin VB.Label Lb_unit_debit 
            Caption         =   "l/s soit "
            Height          =   255
            Index           =   1
            Left            =   3050
            LinkItem        =   " "
            TabIndex        =   65
            Top             =   1110
            Width           =   500
         End
         Begin VB.Label Lb_unit_debit 
            Caption         =   "l/s soit "
            Height          =   255
            Index           =   2
            Left            =   3050
            LinkItem        =   " "
            TabIndex        =   64
            Top             =   1500
            Width           =   500
         End
         Begin VB.Label Lb_unit_debit 
            Caption         =   "l/s soit "
            Height          =   255
            Index           =   3
            Left            =   3050
            LinkItem        =   " "
            TabIndex        =   63
            Top             =   1890
            Width           =   500
         End
         Begin VB.Label Lb_unit_debit 
            Caption         =   "l/s soit "
            Height          =   255
            Index           =   4
            Left            =   3050
            LinkItem        =   " "
            TabIndex        =   62
            Top             =   2280
            Width           =   500
         End
         Begin VB.Label Lb_unit_debitc 
            Caption         =   "m3/h"
            Height          =   255
            Index           =   4
            Left            =   4550
            TabIndex        =   61
            Top             =   2280
            Width           =   460
         End
         Begin VB.Label Lb_unit_debitc 
            Caption         =   "m3/h"
            Height          =   255
            Index           =   3
            Left            =   4550
            TabIndex        =   60
            Top             =   1890
            Width           =   460
         End
         Begin VB.Label Lb_unit_debitc 
            Caption         =   "m3/h"
            Height          =   255
            Index           =   2
            Left            =   4550
            TabIndex        =   59
            Top             =   1500
            Width           =   460
         End
         Begin VB.Label Lb_unit_debitc 
            Caption         =   "m3/h"
            Height          =   255
            Index           =   1
            Left            =   4560
            TabIndex        =   58
            Top             =   1110
            Width           =   465
         End
         Begin VB.Label Lb_unit_debitc 
            Caption         =   "m3/h"
            Height          =   255
            Index           =   0
            Left            =   4550
            TabIndex        =   57
            Top             =   330
            Width           =   460
         End
      End
      Begin VB.Frame Frm_geom2 
         Caption         =   "Niveaux "
         Height          =   1725
         Left            =   -74880
         TabIndex        =   47
         Top             =   1920
         Width           =   5050
         Begin VB.TextBox Tb_Geom 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   3420
            TabIndex        =   11
            Top             =   240
            Width           =   850
         End
         Begin VB.TextBox Tb_Geom 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   3420
            TabIndex        =   12
            Top             =   560
            Width           =   850
         End
         Begin VB.TextBox Tb_Geom 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   3420
            TabIndex        =   13
            Top             =   880
            Width           =   850
         End
         Begin VB.TextBox Tb_Geom 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   3420
            TabIndex        =   14
            Top             =   1200
            Width           =   850
         End
         Begin VB.Label Lb_Int_Geom 
            Caption         =   "Terrain naturel"
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   55
            Top             =   270
            Width           =   2520
         End
         Begin VB.Label Lb_Unit_Geom 
            Caption         =   "m"
            Height          =   285
            Index           =   3
            Left            =   4410
            TabIndex        =   54
            Top             =   270
            Width           =   315
         End
         Begin VB.Label Lb_Int_Geom 
            Caption         =   "Fil d'eau d'arrivée"
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   53
            Top             =   590
            Width           =   2520
         End
         Begin VB.Label Lb_Unit_Geom 
            Caption         =   "m"
            Height          =   285
            Index           =   4
            Left            =   4410
            TabIndex        =   52
            Top             =   590
            Width           =   315
         End
         Begin VB.Label Lb_Int_Geom 
            Caption         =   "Fil d'eau de sortie"
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   51
            Top             =   910
            Width           =   2520
         End
         Begin VB.Label Lb_Unit_Geom 
            Caption         =   "m"
            Height          =   285
            Index           =   5
            Left            =   4410
            TabIndex        =   50
            Top             =   915
            Width           =   315
         End
         Begin VB.Label Lb_Int_Geom 
            Caption         =   "Fil d'eau extrémité du refoulement"
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   49
            Top             =   1200
            Width           =   2520
         End
         Begin VB.Label Lb_Unit_Geom 
            Caption         =   "m"
            Height          =   285
            Index           =   6
            Left            =   4410
            TabIndex        =   48
            Top             =   1230
            Width           =   315
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   34
         Top             =   600
         Width           =   5050
         Begin VB.TextBox Tb_PtSing 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   0
            Left            =   1600
            TabIndex        =   15
            Text            =   "Text1"
            Top             =   600
            Width           =   380
         End
         Begin VB.TextBox Tb_PtSing 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   1
            Left            =   2300
            TabIndex        =   16
            Text            =   "Text1"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Tb_PtSing 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   2
            Left            =   3000
            TabIndex        =   17
            Text            =   "Text1"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Tb_PtSing 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   3
            Left            =   3700
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Tb_PtSing 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   4
            Left            =   4400
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   600
            Width           =   375
         End
         Begin VB.TextBox Tb_PtSing 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   5
            Left            =   2700
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   1080
            Width           =   375
         End
         Begin VB.TextBox Tb_PtSing 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   6
            Left            =   2700
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   1450
            Width           =   375
         End
         Begin VB.TextBox Tb_PtSing 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   7
            Left            =   2700
            TabIndex        =   22
            Text            =   "Text1"
            Top             =   1800
            Width           =   375
         End
         Begin VB.TextBox Tb_PtSing 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   8
            Left            =   2700
            TabIndex        =   23
            Text            =   "Text1"
            Top             =   2150
            Width           =   375
         End
         Begin VB.Frame Frame4 
            Caption         =   "Anti-Bélier"
            Height          =   1455
            Left            =   3360
            TabIndex        =   35
            Top             =   1080
            Width           =   1305
            Begin VB.OptionButton Opt_PtSing 
               Caption         =   "OUI"
               Height          =   255
               Index           =   0
               Left            =   180
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   360
               Width           =   885
            End
            Begin VB.OptionButton Opt_PtSing 
               Caption         =   "NON"
               Height          =   255
               Index           =   1
               Left            =   180
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   960
               Width           =   885
            End
         End
         Begin VB.Label Lb_Ptsing 
            Caption         =   "Nbre de coude(s)"
            Height          =   255
            Left            =   100
            TabIndex        =   76
            Top             =   630
            Width           =   1335
         End
         Begin VB.Label Lb_int_PtSing 
            Caption         =   " 11°15"
            Height          =   255
            Index           =   0
            Left            =   1500
            TabIndex        =   46
            Top             =   345
            Width           =   580
         End
         Begin VB.Label Lb_int_PtSing 
            Caption         =   " 22°30"
            Height          =   255
            Index           =   1
            Left            =   2250
            TabIndex        =   45
            Top             =   345
            Width           =   585
         End
         Begin VB.Label Lb_int_PtSing 
            Caption         =   "30°"
            Height          =   255
            Index           =   2
            Left            =   3100
            TabIndex        =   44
            Top             =   360
            Width           =   465
         End
         Begin VB.Label Lb_int_PtSing 
            Caption         =   " 45°"
            Height          =   255
            Index           =   3
            Left            =   3750
            TabIndex        =   43
            Top             =   340
            Width           =   460
         End
         Begin VB.Label Lb_int_PtSing 
            Caption         =   "90°"
            Height          =   255
            Index           =   4
            Left            =   4500
            TabIndex        =   42
            Top             =   340
            Width           =   460
         End
         Begin VB.Label Lb_int_PtSing 
            Caption         =   "Nbre de vanne(s)"
            Height          =   255
            Index           =   5
            Left            =   100
            TabIndex        =   41
            Top             =   1130
            Width           =   2260
         End
         Begin VB.Label Lb_int_PtSing 
            Caption         =   "Nbre de clapet(s) anti-retour"
            Height          =   255
            Index           =   6
            Left            =   100
            TabIndex        =   40
            Top             =   1480
            Width           =   2260
         End
         Begin VB.Label Lb_int_PtSing 
            Caption         =   "Nbre de système(s) de vidange"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   39
            Top             =   1830
            Width           =   2260
         End
         Begin VB.Label Lb_int_PtSing 
            Caption         =   "Nbre de ventouse(s)"
            Height          =   255
            Index           =   8
            Left            =   100
            TabIndex        =   38
            Top             =   2180
            Width           =   2260
         End
      End
      Begin VB.TextBox Tb_Qpomp 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2250
         MaxLength       =   8
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   3300
         Width           =   850
      End
      Begin VB.TextBox Tb_Qpompc 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3720
         MaxLength       =   8
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   3300
         Width           =   850
      End
      Begin VB.Frame Frm_geom1 
         Caption         =   "Conduite de refoulement"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   139
         Top             =   480
         Width           =   5055
         Begin VB.TextBox Tb_Geom 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   3420
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   950
            Width           =   850
         End
         Begin VB.TextBox Tb_Geom 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   3420
            TabIndex        =   8
            Top             =   240
            Width           =   850
         End
         Begin VB.ComboBox Cb_Materiau 
            Height          =   315
            Left            =   3420
            TabIndex        =   9
            Top             =   600
            Width           =   1305
         End
         Begin VB.Label Lb_Unit_Geom 
            Caption         =   "mm"
            Height          =   285
            Index           =   2
            Left            =   4410
            TabIndex        =   144
            Top             =   980
            Width           =   315
         End
         Begin VB.Label Lb_Int_Geom 
            Caption         =   "Diamètre théorique (V=1 m/s)"
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   143
            Top             =   960
            Width           =   3000
         End
         Begin VB.Label Lb_Int_Geom 
            Caption         =   "Nature du tuyau"
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   142
            Top             =   600
            Width           =   3000
         End
         Begin VB.Label Lb_Unit_Geom 
            Caption         =   "m"
            Height          =   285
            Index           =   0
            Left            =   4440
            TabIndex        =   141
            Top             =   270
            Width           =   315
         End
         Begin VB.Label Lb_Int_Geom 
            Caption         =   "Longueur"
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   140
            Top             =   270
            Width           =   3000
         End
      End
      Begin VB.Label Lb_unit_denivt 
         Caption         =   "m"
         Height          =   255
         Left            =   -70920
         TabIndex        =   118
         Top             =   2520
         Width           =   300
      End
      Begin VB.Label Lb_unit_vutba 
         Caption         =   "m3"
         Height          =   255
         Left            =   -70920
         TabIndex        =   117
         Top             =   1200
         Width           =   300
      End
      Begin VB.Label Lb_int_denivt 
         Caption         =   "Tranche de pompage théorique"
         Height          =   285
         Left            =   -74640
         TabIndex        =   115
         Top             =   2520
         Width           =   2580
      End
      Begin VB.Label Lb_int_Vutba 
         Caption         =   "Volume utile théorique de la bâche"
         Height          =   285
         Left            =   -74640
         TabIndex        =   93
         Top             =   1180
         Width           =   2580
      End
      Begin VB.Label Lb_resu 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -74400
         TabIndex        =   92
         Top             =   3240
         Width           =   3735
      End
      Begin VB.Label Lb_int_Qpomp 
         Caption         =   "Débit de pompage théorique"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   75
         Top             =   3360
         Width           =   2115
      End
      Begin VB.Label Lb_unit_Qpompc 
         Caption         =   "m3/h"
         Height          =   255
         Index           =   0
         Left            =   4680
         TabIndex        =   74
         Top             =   3360
         Width           =   465
      End
      Begin VB.Label Lb_unit_Qpomp 
         Caption         =   "l/s soit "
         Height          =   255
         Index           =   0
         Left            =   3200
         TabIndex        =   73
         Top             =   3360
         Width           =   495
      End
   End
   Begin VB.ComboBox Cb_pompe 
      Height          =   315
      Left            =   240
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   120
      Width           =   4000
   End
   Begin VB.CommandButton Cmd_calcul 
      Caption         =   "Calculer"
      Height          =   255
      Left            =   3120
      TabIndex        =   131
      TabStop         =   0   'False
      ToolTipText     =   "Calcul de la chute"
      Top             =   4000
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.Label Lb_Unit_Geom 
      Caption         =   "m"
      Height          =   285
      Index           =   1
      Left            =   4560
      TabIndex        =   162
      Top             =   1200
      Width           =   315
   End
   Begin VB.Label Lb_amo 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lb_amont"
      Height          =   615
      Left            =   360
      TabIndex        =   32
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label Lb_temp 
      Caption         =   "Lb_temp"
      Height          =   375
      Left            =   1200
      TabIndex        =   31
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Lb_pompe 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lb_pompe"
      Height          =   495
      Left            =   3480
      TabIndex        =   30
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Lb_ava 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Lb_ava"
      Height          =   750
      Left            =   6600
      TabIndex        =   29
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Menu mnufichier 
      Caption         =   "&Station de pompage"
      Begin VB.Menu mnunouv 
         Caption         =   "&Nouveau"
      End
      Begin VB.Menu mnuouv 
         Caption         =   "&Ouvrir..."
      End
      Begin VB.Menu f1 
         Caption         =   "-"
      End
      Begin VB.Menu mnusave 
         Caption         =   "&Enregistrer"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnusaves 
         Caption         =   "En&registrer sous..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnusuppr 
         Caption         =   "&Supprimer..."
         Enabled         =   0   'False
      End
      Begin VB.Menu f2 
         Caption         =   "-"
      End
      Begin VB.Menu Mnuprint 
         Caption         =   "Im&primer..."
         Enabled         =   0   'False
      End
      Begin VB.Menu f3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quitter module"
      End
   End
End
Attribute VB_Name = "Frm_pompe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private okg As Boolean
Private owner As MDIFrm_menu
Private esave As st_savpompe
Public nom_ouvrage As String
'Private nom_fich As String
Public nom_type As String
Private lhFicDbf As Long
Private FileLength As Integer
Private list_don1() As Variant
Private list_don2() As Variant
Private list_don3() As Variant
Private list_don4() As Variant
Private list_don5() As Variant
Private ch_texte As String
Private fen_titre As String
Public titre_sav As String
Private list_tb() As Variant
Private sval_champ As String
Private iSels As Integer
Private iSell As Integer
Private bKP As Boolean
Private label_prec As String
Private mes_prec As String
Private index_prec As Integer
Private change_coul As Boolean
Private Vutba As Double
Private Vurba As Double
Private SECBA As Double
Private DENIV As Double
Private Nrdph As Double
Private T1cyc As Double
Private Tvidange As Double
Private Nbcyc As Double
Private VitRflt As Double
Private Tsejh As Double
Private Vmy As Double
Private Denivhau As Double
Private Denivbas As Double
Private Typo As Integer
Private Singul As Double
Private Hmt As Double
Private jmpkm As Double
Private pi As Double
Private ok_saisie_denivr As Boolean
Private Sub meAffiche()
    DoEvents
    Me.Show
End Sub
Private Sub Change_Couleur(nom As String, Index As Integer)
'Dim coul As ColorConstants, coulp As ColorConstants
'Dim Index1 As Integer
'Dim nom1 As String
'coulp = vbBlack
'coul = Couleur_Change
'nom1 = nom
'Select Case nom
'    Case Is = "Tb_amo"
'         nom1 = "Lb_intam"
'    Case Is = "Tb_ava"
'         nom1 = "Lb_intava"
'    Case Is = "Tb_Qmax"
'         nom1 = "Lb_Qmax"
'End Select
'Select Case label_prec
'    Case Is = "Lb_intam"
'         Lb_intam(index_prec).ForeColor = coulp
'    Case Is = "Lb_intava"
'         Lb_intava(index_prec).ForeColor = coulp
'    Case Is = "Lb_Qmax"
'         Lb_Qmax.ForeColor = coulp
'    Case Is = "Frm_Amont"
'         Frm_Amont.ForeColor = coulp
'    Case Is = "Frm_Aval"
'         Frm_Aval.ForeColor = coulp
'End Select
'Select Case nom1
'    Case Is = "Me"
'         Me.SetFocus
'    Case Is = "Lb_intam"
'         Lb_intam(Index).ForeColor = coul
'    Case Is = "Lb_intava"
'         Lb_intava(Index).ForeColor = coul
'    Case Is = "Lb_Qmax"
'         Lb_Qmax.ForeColor = coul
'    Case Is = "Frm_Amont"
'         Frm_Amont.ForeColor = coul
'   Case Is = "Frm_Aval"
'         Frm_Aval.ForeColor = coul
'End Select
'label_prec = nom1
'index_prec = Index
'change_coul = True
End Sub
Private Sub Change_Focus(nom As String, Index As Integer)
Dim coul As ColorConstants, coulp As ColorConstants
Dim Index1 As Integer
Dim nom1 As String
coulp = vbBlack
coul = Couleur_Change
nom1 = nom
Select Case nom1
    Case Is = "Me"
         Me.SetFocus
'    Case Is = "Lb_intam"
'         Tb_amo(Index).SetFocus
'    Case Is = "Lb_intava"
'         Tb_ava(Index).SetFocus
'    Case Is = "Lb_Qmax"
'         Tb_Qmax.SetFocus
'    Case Is = "Frm_Amont"
'         Tb_amo(0).SetFocus
'   Case Is = "Frm_Aval"
'         Tb_ava(0).SetFocus
End Select
End Sub
Private Function Rec_Mes(nom As String, Index As Integer)
Dim mes As String
mes = ""
Select Case nom
    Case Is = "Lb_int_debit", "Lb_unit_debit", "Tb_debit", "Tb_Debitc"
    Select Case Index
        Case Is = 0
        mes = IDhlp_PompeDebitsCaracteristiques '"Débit moyen des eaux usées"
        Case Is = 1
        mes = IDhlp_PompeDebitsCaracteristiques '"Débit pointe des eaux usées"
        Case Is = 2
        mes = IDhlp_PompeDebitsCaracteristiques '"Débit des eaux parasites"
        Case Is = 3
        mes = IDhlp_PompeDebitsCaracteristiques '"Débit moyen de temps sec"
        Case Is = 4
        mes = IDhlp_PompeDebitsCaracteristiques '"Débit de pointe de temps sec"
    End Select
    Case Is = "Lb_int_Fpointe", "Tb_FPointe"
        mes = IDhlp_PompeDebitsCaracteristiques '"Facteur de pointe"
    Case Is = "Frm_geom1"
        mes = IDhlp_PompeDonneesGeometriques ' '"Conduite de refoulement"
    Case Is = "Frm_geom2"
        mes = IDhlp_PompeDonneesGeometriques '"niveaux"
    Case Is = "Lb_Int_Geom", "Lb_Unit_Geom", "Tb_Geom"
    Select Case Index
        Case Is = 0
        mes = IDhlp_PompeDonneesGeometriques '"Longueur"
        Case Is = 1
        mes = IDhlp_PompeDonneesGeometriques '"Diamètre théorique (V=1.5 m/s)"
        Case Is = 2
        mes = IDhlp_PompeDonneesGeometriques '"Nature du tuyau"
        Case Is = 3
        mes = IDhlp_PompeDonneesGeometriques '"Terrain naturel"
        Case Is = 4
        mes = IDhlp_PompeDonneesGeometriques '"Fil d'eau d'arrivée"
        Case Is = 5
        mes = IDhlp_PompeDonneesGeometriques '"Fil d'eau de sortie"
        Case Is = 6
        mes = IDhlp_PompeDonneesGeometriques '"Fil d'eau extrémité du refoulement"
    End Select
    Case Is = "Cb_Materiau"
        mes = IDhlp_PompeDonneesGeometriques '"Nature du tuyau"
    Case Is = "Frame3"
        mes = IDhlp_PompePointsSinguliers
    Case Is = "Lb_int_PtSing", "Tb_PtSing", "Lb_Ptsing"
    Select Case Index
        Case Is = 0
        mes = IDhlp_PompePointsSinguliers '"Nbre de coude(s) 11°15"
        Case Is = 1
        mes = IDhlp_PompePointsSinguliers '"Nbre de coude(s)  22°30"
        Case Is = 2
        mes = IDhlp_PompePointsSinguliers '"Nbre de coude(s)  30°"
        Case Is = 3
        mes = IDhlp_PompePointsSinguliers '"Nbre de coude(s)  45°"
        Case Is = 4
        mes = IDhlp_PompePointsSinguliers '"Nbre de coude(s)  90°"
        Case Is = 5
        mes = IDhlp_PompePointsSinguliers '"Nbre de vanne(s)"
        Case Is = 6
        mes = IDhlp_PompePointsSinguliers '"Nbre de clapet(s) anti-retour"
        Case Is = 7
        mes = IDhlp_PompePointsSinguliers '"Nbre de système(s) de vidange"
        Case Is = 8
        mes = IDhlp_PompePointsSinguliers '"Nbre de ventouse(s)"
    End Select
    Case Is = "Opt_PtSing"
    Select Case Index
        Case Is = 0
        mes = IDhlp_PompePointsSinguliersProtection '"Anti-Bélier OUI"
        Case Is = 1
        mes = IDhlp_PompePointsSinguliersProtection '"Anti-Bélier NON"
    End Select
    Case Is = "Frame4"
        mes = IDhlp_PompePointsSinguliersProtection '"Anti-Bélier"
    Case Is = "Frame6"
         mes = IDhlp_PompeDonneesTechniques '"Longueur"
    Case Is = "Lb_Int_Nbpom", "Tb_Nbpom"
         mes = IDhlp_PompeDonneesTechniques '"Nb de pompe(s)"
    Case Is = "Lb_Int_Ntdph", "Tb_Ntdph"
         mes = IDhlp_PompeDonneesTechniques '"Nb de démarrage(s) /h"
    Case Is = "Lb_int_Vutba", "Tb_Vutba"
         mes = IDhlp_PompeDonneesTechniques '"Volume utile théorique de la bâche"
    Case Is = "Frm_bache"
        mes = IDhlp_PompeDonneesTechniques '"Section de la bâche"
    Case Is = "Opt_sect_ba"
    Select Case Index
        Case Is = 0
         mes = IDhlp_PompeDonneesTechniques '"Circulaire"
        Case Is = 1
         mes = IDhlp_PompeDonneesTechniques '"Rectangulaire"
    End Select
    Case Is = "Frame8"
         mes = IDhlp_PompeDonneesTechniques '"Longueur"
    Case Is = "Lb_int_long", "Lb_unit_long", "Tb_long"
         mes = IDhlp_PompeDonneesTechniques '"Longueur"
    Case Is = "Lb_int_larg", "Lb_unit_larg", "Tb_larg"
         mes = IDhlp_PompeDonneesTechniques '"largeur"
    Case Is = "Lb_int_diam", "Lb_unit_diam", "Tb_diam"
         mes = IDhlp_PompeDonneesTechniques '"Diamétre"
    Case Is = "Lb_int_denivt", "Tb_denivt", "Lb_unit_denivt"
         mes = IDhlp_PompeDonneesTechniques '"Tranche de pompage théorique"
    Case Is = "Lb_int_denivhau", "Tb_denivhau", "Lb_unit_denivhau"
         mes = IDhlp_PompeDonneesTechniques '"Garde à l'égout"
    Case Is = "Lb_int_denivbas", "Tb_denivbas", "Lb_unit_denivbas"
         mes = IDhlp_PompeDonneesTechniques '"Garde au fond"
    Case Is = "Lb_int_Qpomp", "Lb_unit_Qpomp", "Tb_Qpomp", "Tb_Qpompc", "Lb_unit_Qpompc"
    Select Case Index
        Case Is = 0
         mes = IDhlp_PompeDonneesTechniques2 '"Débit de pompage théorique"
        Case Is = 1
         mes = IDhlp_PompeDonneesTechniques2 '"Débit de pompage"
    End Select
    Case Is = "Lb_Int_Drflt", "Lb_Unit_Drflt", "Lb_materiau", "Tb_Drflt"
         mes = IDhlp_PompeDonneesTechniques2 '"Canalisation Lbl_materiau"
    Case Is = "Lb_int_Vitrflt", "Lb_unit_Vitrflt", "Tb_VitRflt"
         mes = IDhlp_PompeDonneesTechniques2 '"Vitesse en régime  permanent"
    Case Is = "Lb_int_Jmpkm", "Lb_unit_Jmpkm", "Tb_Jmpkm"
         mes = IDhlp_PompeDonneesTechniques2 '"Perte de charges  linéaires"
    Case Is = "Lb_int_denivr", "Lb_unit_denivr", "Tb_denivr"
         mes = IDhlp_PompeDonneesTechniques2 '"Tranche de pompage retenue"
    Case Is = "Lb_int_vurba", "Lb_unit_vurba", "Tb_vurba"
         mes = IDhlp_PompeDonneesTechniques2 '"Volume utile de la bâche"
    Case Is = "Lb_int_nrdph", "Tb_nrdph"
         mes = IDhlp_PompeDonneesTechniques2 '"Nb réel de démarrage(s) /h"
    Case Is = "Lb_int_tvidange", "Lb_unit_tvidange", "Tb_Tvidange"
         mes = IDhlp_PompeDonneesTechniques2 '"Tps vidange"
    Case Is = "Lb_int_T1cyc", "Lb_unit_T1cyc", "Tb_T1cyc"
         mes = IDhlp_PompeDonneesTechniques2 '"Cycle"
    Case Is = "Lb_Int_Nbcyc", "Lb_unit_Nbcyc", "Tb_Nbcyc"
         mes = IDhlp_PompeDonneesTechniques2 '"NbC"
    Case Is = "Lb_int_Vmy", "Lb_unit_Vmy", "Tb_Vmy"
         mes = IDhlp_PompeDonneesTechniques2 '"Vitesse moyenne d'écoulement"
    Case Is = "Lb_int_Tsejh", "Lb_unit_Tsejh", "Tb_Tsejh"
         mes = IDhlp_PompeDonneesTechniques2 '"Temps de séjour"
    Case Is = "Lb_int_Singul", "Lb_unit_Singul", "Tb_Singul"
         mes = IDhlp_PompeDonneesTechniques2 '"Perte de charge singulière"
    Case Is = "Lb_int_Hmt", "Lb_unit_Hmt", "Tb_Hmt"
         mes = IDhlp_PompeDonneesTechniques2 '"Hauteur manométrique totale"
End Select
mes_prec = mes
Rec_Mes = mes
End Function
Public Function get_l_tb() As Variant
get_l_tb = list_tb
End Function
Private Sub init_l_tab()
Dim l0() As Variant, l1() As Variant, l2() As Variant, l3() As Variant, l4() As Variant
Dim l5() As Variant
l0 = Array(0)
l1 = Array(0, "TB_debit", "TB_FPointe")
l2 = Array(0, "TB_Geom")
l3 = Array(0, "TB_PtSing")
l4 = Array(0, "Tb_Nbpom", "Tb_Ntdph", "Tb_long", "Tb_larg", "Tb_diam", "Tb_denivhau", "Tb_denivbas")
l5 = Array(0, "Tb_Qpomp", "Tb_Drflt", "Tb_denivr")
ReDim list_tb(0 To UBound(l0), 0 To UBound(l1), 0 To UBound(l2), 0 To UBound(l3), 0 To UBound(l4), 0 To UBound(l5))
list_tb = Array(l0, l1, l2, l4, l5)

End Sub

Public Sub retailler()
retaille

End Sub
Public Function recup_mnuprint()
    recup_mnuprint = Me.mnuprint.Enabled
End Function
Private Sub retaille()
    Me.Left = owner.fcom.Width + owner.fcom.Left
    Me.Top = 0
    Me.Width = maximum(larg_mini, owner.Width - owner.fcom.Width - owner.fcom.Left - l_decal_asc)  ' 10040
    Me.Height = maximum(haut_mini, owner.fdessin.Top) '4600
End Sub


Private Sub Cb_Materiau_Click()
    Lb_materiau.Caption = Cb_Materiau.Text & " de D = "
    ebpompe.don_geometrie.NatRflt = Cb_Materiau.Text
    ebpompe.resultat.NatRflr = Lb_materiau.Caption
    Call calcul_resu
'    Tb_Drflt.Text = Format(val(Tb_Geom(2).Text), "###0")
'    ebpompe.resultat.Drflr = txtVersNum(Me.Tb_Drflt.Text)
'''    Call calc_vit_pcl
'''            Call change_hmt
Dim mes As String
Dim nom As String
Dim Index As Integer
nom = "Cb_Materiau"
Index = 0
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes


End Sub

Private Sub Cb_Materiau_GotFocus()
Dim mes As String
Dim nom As String
Dim Index As Integer
nom = "Cb_Materiau"
Index = 0
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Cb_pompe_Change()
    Cb_pompe.Text = ch_texte
End Sub

Private Sub Cb_pompe_KeyDown(KeyCode As Integer, Shift As Integer)
    ch_texte = Cb_pompe.Text
    Cb_pompe.Text = ch_texte

End Sub

Private Sub Cb_pompe_KeyPress(KeyAscii As Integer)
    ch_texte = Cb_pompe.Text
End Sub



Private Sub Cmd_calcul_Click()
   Call calcul_resu
End Sub
Function dessin_pompe()
Dim mes As String
'hgarde = 0.15
'hdeniv = 0.8
'hfond = 0.5
'diam = 3#
If ebpompe.don_geometrie.NivTN > 0 And ebpompe.don_geometrie.NivEX > 0 And ebpompe.don_geometrie.NivEN > 0 _
    And ebpompe.don_geometrie.NivSO > 0 And ebpompe.don_techniques.Denivhau > 0 _
    And ebpompe.resultat.Denivr > 0 And ebpompe.don_techniques.Denivbas > 0 _
    And ((ebpompe.don_techniques.Sectb = 0 And ebpompe.don_techniques.Diamb > 0) _
    Or (ebpompe.don_techniques.Sectb = 1 And ebpompe.don_techniques.Largb > 0)) _
    And ebpompe.resultat.Hmt > 0 Then
    Call init_graph(owner.fdessin.UC_graphique1)
    Call init_graph(Frm_desprint.UC_graphique1)
    Call dess_pompe(owner.fdessin.UC_graphique1)
    Call dess_pompe(Frm_desprint.UC_graphique1)

Else
'    mes = "toutes les données ne sont pas renseignées"
'    MsgBox mes
    owner.fdessin.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.graphique_clear

End If

End Function
Private Sub lect_fich()
Dim za As st_savpompe
Dim za1 As st_savpom1
Call funlockb
 
    lhFicDbf = FreeFile
    Cb_pompe.Clear
    On Error GoTo test_Error
    Open nom_fich For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
    If Not EOF(lhFicDbf) Then
        za = za1.stsavpo
        If Trim(za.type) = nom_type Then
            Cb_pompe.AddItem (Trim(za.nom))
        End If
   End If
Loop
Close #lhFicDbf
ch_texte = Cb_pompe.list(0)
Cb_pompe.Text = Cb_pompe.list(0)
Cb_pompe.Refresh

Call flockb(nom_fich)
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + nom_fich + " est déjà en cours d'utilisation.")
    End If
 
Call flockb(nom_fich)
End Sub


Private Sub Cmd_calcul_GotFocus()
Dim nom As String
Dim mes As String
    nom = "Cmd_calcul"
    mes = Rec_Mes(nom, 0)
    owner.affich_aide Me.Name, mes
End Sub

Private Sub Form_Activate()
    change_coul = False

'    owner.affich_aide Me.Name, mes_prec
End Sub

Private Sub Form_Click()
    owner.affich_aide Me.Name, ""  'Dimensionnement d'une pompe"
    Change_Couleur "Me", 0
End Sub

Private Sub m_quitter_Click()
    Unload Me
    Unload owner
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
owner.fcom.Form_KeyAide KeyCode, Shift
Me.SetFocus
End Sub

Private Sub Frame3_Click()
Dim mes As String
Dim nom As String
nom = "Frame3"
mes = Rec_Mes(nom, Index)
owner.affich_aide Me.Name, mes
Change_Couleur nom, 0
Change_Focus nom, 0

End Sub
Private Sub Frame4_Click()
Dim mes As String
Dim nom As String
nom = "Frame4"
mes = Rec_Mes(nom, Index)
owner.affich_aide Me.Name, mes
Change_Couleur nom, 0
Change_Focus nom, 0

End Sub
Private Sub Frame6_Click()
Dim mes As String
Dim nom As String
nom = "Frame6"
mes = Rec_Mes(nom, Index)
owner.affich_aide Me.Name, mes
Change_Couleur nom, 0
Change_Focus nom, 0
End Sub
Private Sub Frame8_Click()
Dim mes As String
Dim nom As String
nom = "Frame8"
mes = Rec_Mes(nom, Index)
owner.affich_aide Me.Name, mes
Change_Couleur nom, 0
Change_Focus nom, 0
End Sub

Private Sub Frm_bache_Click()
Dim mes As String
Dim nom As String
nom = "Frm_bache"
mes = Rec_Mes(nom, Index)
owner.affich_aide Me.Name, mes
Change_Couleur nom, 0
End Sub
Private Sub Frm_geom1_Click()
Dim mes As String
Dim nom As String
nom = "Frm_geom1"
mes = Rec_Mes(nom, Index)
owner.affich_aide Me.Name, mes
Change_Couleur nom, 0
Change_Focus nom, 0
End Sub
Private Sub Frm_geom2_Click()
Dim mes As String
Dim nom As String
nom = "Frm_geom2"
mes = Rec_Mes(nom, Index)
owner.affich_aide Me.Name, mes
Change_Couleur nom, 0
Change_Focus nom, 0
End Sub
Private Sub Lb_int_debit_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_int_debit"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_int_denivbas_Click()
Dim mes As String
Dim nom As String
nom = "Lb_int_denivbas"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes
End Sub

Private Sub Lb_int_denivhau_Click()
Dim mes As String
Dim nom As String
nom = "Lb_int_denivhau"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_int_denivr_Click()
Dim mes As String
Dim nom As String
nom = "Lb_int_denivr"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_int_denivt_Click()
Dim mes As String
Dim nom As String
nom = "Lb_int_denivt"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes
End Sub

Private Sub Lb_int_diam_Click()
Dim mes As String
Dim nom As String
nom = "Lb_int_diam"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_Int_Drflt_Click()
Dim mes As String
Dim nom As String
nom = "Lb_Int_Drflt"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_int_Fpointe_Click()
Dim mes As String
Dim nom As String
nom = "Lb_int_Fpointe"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_Int_Geom_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_Int_Geom"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes
End Sub

Private Sub Lb_int_Hmt_Click()
Dim mes As String
Dim nom As String
nom = "Lb_int_Hmt"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_int_Jmpkm_Click()
Dim mes As String
Dim nom As String
nom = "Lb_int_Jmpkm"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes
End Sub

Private Sub Lb_int_larg_Click()
Dim mes As String
Dim nom As String
nom = "Lb_int_larg"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_int_long_Click()
Dim mes As String
Dim nom As String
nom = "Lb_int_long"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_int_Nbcyc_Click()
Dim mes As String
Dim nom As String
nom = "Lb_Int_Nbcyc"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_Int_Nbpom_Click()
Dim mes As String
Dim nom As String
nom = "Lb_Int_Nbpom"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_int_nrdph_Click()
Dim mes As String
Dim nom As String
nom = "Lb_int_nrdph"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_Int_Ntdph_Click()
Dim mes As String
Dim nom As String
nom = "Lb_Int_Ntdph"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_int_PtSing_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_int_PtSing"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_int_Qpomp_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_int_Qpomp"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_int_Singul_Click()
Dim mes As String
Dim nom As String
nom = "Lb_int_Singul"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes
End Sub

Private Sub Lb_int_T1cyc_Click()
Dim mes As String
Dim nom As String
nom = "Lb_int_T1cyc"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_int_Tsejh_Click()
Dim mes As String
Dim nom As String
nom = "Lb_int_Tsejh"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes
End Sub

Private Sub Lb_int_tvidange_Click()
Dim mes As String
Dim nom As String
nom = "Lb_int_tvidange"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_int_Vitrflt_Click()
Dim mes As String
Dim nom As String
nom = "Lb_int_Vitrflt"
'nom = Me.Lb_int_Vitrflt.Name
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_int_Vmy_Click()
Dim mes As String
Dim nom As String
nom = "Lb_int_Vmy"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_int_vurba_Click()
Dim mes As String
Dim nom As String
nom = "Lb_int_vurba"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_int_Vutba_Click()
Dim mes As String
Dim nom As String
nom = "Lb_int_Vutba"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_materiau_Click()
Dim mes As String
Dim nom As String
nom = "Lb_materiau"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0

End Sub

Private Sub Lb_Ptsing_Click()
Dim mes As String
Dim nom As String
nom = "Lb_Ptsing"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_unit_debit_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_unit_debit"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes
End Sub

Private Sub Lb_unit_debitc_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_unit_debit"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes
End Sub
Private Sub Lb_unit_denivbas_Click()
Dim mes As String
Dim nom As String
nom = "Lb_unit_denivbas"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes
End Sub
Private Sub Lb_unit_denivhau_Click()
Dim mes As String
Dim nom As String
nom = "Lb_unit_denivhau"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_unit_denivr_Click()
Dim mes As String
Dim nom As String
nom = "Lb_unit_denivr"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes
End Sub

Private Sub Lb_unit_denivt_Click()
Dim mes As String
Dim nom As String
nom = "Lb_unit_denivt"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes
End Sub

Private Sub Lb_unit_diam_Click()
Dim mes As String
Dim nom As String
nom = "Lb_unit_diam"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
Call sel_text(Tb_Geom(Index))

End Sub
Private Sub Lb_Unit_Drflt_Click()
Dim mes As String
Dim nom As String
nom = "Lb_Unit_Drflt"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_Unit_Geom_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_Unit_Geom"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
Call sel_text(Tb_Geom(Index))

End Sub

Private Sub Lb_unit_Hmt_Click()
Dim mes As String
Dim nom As String
nom = "Lb_unit_Hmt"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_unit_Jmpkm_Click()
Dim mes As String
Dim nom As String
nom = "Lb_unit_Jmpkm"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_unit_larg_Click()
Dim mes As String
Dim nom As String
nom = "Lb_unit_larg"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_unit_long_Click()
Dim mes As String
Dim nom As String
nom = "Lb_unit_long"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_unit_Qpomp_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_unit_Qpomp"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_unit_Qpompc_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_unit_Qpompc"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes
End Sub

Private Sub Lb_unit_Singul_Click()
Dim mes As String
Dim nom As String
nom = "Lb_unit_Singul"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes
End Sub

Private Sub Lb_unit_T1cyc_Click()
Dim mes As String
Dim nom As String
nom = "Lb_unit_T1cyc"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes
End Sub

Private Sub Lb_unit_Tsejh_Click()
Dim mes As String
Dim nom As String
nom = "Lb_unit_Tsejh"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes
End Sub

Private Sub Lb_unit_Vitrflt_Click()
Dim mes As String
Dim nom As String
nom = "Lb_unit_Vitrflt"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_unit_Vmy_Click()
Dim mes As String
Dim nom As String
nom = "Lb_unit_Vmy"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_unit_vurba_Click()
Dim mes As String
Dim nom As String
nom = "Lb_unit_vurba"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Change_Focus nom, 0
owner.affich_aide Me.Name, mes
End Sub

Private Sub mnufichier_Click()
    If ouv_sauve Or save_fich Then  '(Not ouv_sauve And Not save_fich) Then
        Me.mnusave.Enabled = True
        Me.mnusaves.Enabled = True
        Me.mnusuppr.Enabled = True
'        Me.mnuprint.Enabled = True
    Else
        Me.mnusave.Enabled = False
        Me.mnusaves.Enabled = False
        Me.mnusuppr.Enabled = False
        Me.mnuprint.Enabled = False
   End If
End Sub

Private Sub mnuinfo_Click()
    Frm_saisie.Show 1
End Sub

Private Sub mnunouv_Click()
Dim reponse As Integer
'modif FO   ' If ProtectCheck(2) <> 0 Then End
If ouv_sauve Then 'And save_fich Then
    reponse = MsgBox("La pompe n'a pas été enregistrée" + Chr(10) _
        + "Voulez vous la sauvegarder?", 3, "Sauvegarde d'une pompe")
    Select Case reponse
        Case Is = 6  ' 6=oui,7=non,2=annuler
            Call mnusave_Click
            Call debut0
        Case Is = 7
            Call debut0
    End Select
Else
    Call debut0
End If
End Sub

Private Sub mnuouv_Click()
Dim reponse As Integer
Dim frmf As Frm_lectfich
Set frmf = New Frm_lectfich
Dim nom As String
'modif FO   ' If ProtectCheck(2) <> 0 Then End
fich_lect = nom_fich
If nom_fich_edit <> "" Then
    nom = "Etude " + nom_fich_edit
Else
    nom = " Nouvelle étude "
End If
If ouv_sauve Then 'And save_fich Then
    reponse = MsgBox("La pompe n'a pas été enregistrée" + Chr(10) _
        + "Voulez vous la sauvegarder?", 3, "Sauvegarde d'une pompe")
    Select Case reponse
        Case Is = 6  ' 6=oui,7=non,2=annuler
            Call mnusave_Click
'            Cb_pompe.Visible = True
            frmf.Label1.Caption = "Recherche d'une pompe "
            frmf.Caption = nom
            frmf.Show 1
            If frmf.nomfich <> "" Then
                Me.nom_ouvrage = frmf.nomfich
                Call Me.Cb_pompe_click
            End If
        Case Is = 7
'            Cb_pompe.Visible = True
            frmf.Label1.Caption = "Recherche d'une pompe "
            frmf.Caption = nom
            frmf.Show 1
            If frmf.nomfich <> "" Then
                Me.nom_ouvrage = frmf.nomfich
                Call Me.Cb_pompe_click
            End If
    End Select
Else
'    Cb_pompe.Visible = True
            frmf.Label1.Caption = "Recherche d'une pompe "
            frmf.Caption = nom
            frmf.Show 1
            If frmf.nomfich <> "" Then
                Me.nom_ouvrage = frmf.nomfich
                Call Me.Cb_pompe_click
            End If
End If
Set frmf = Nothing
End Sub

Public Sub cre_list_don1()
Dim i As Integer, j As Integer
ReDim list_don1(7, 5)
For i = 0 To Tb_debit.count - 1
    j = i
    If i > 0 Then
        j = i + 1
    End If
    list_don1(j, 1) = Lb_int_debit(i).Caption
    list_don1(j, 2) = str$(Int(val(Tb_debit(i).Text)))
    list_don1(j, 3) = Lb_unit_debit(i).Caption
    list_don1(j, 4) = Tb_Debitc(i).Text
    list_don1(j, 5) = Lb_unit_debitc(i).Caption
Next
    list_don1(1, 1) = Lb_int_Fpointe.Caption
    list_don1(1, 2) = Tb_FPointe.Text
    list_don1(1, 3) = ""
    list_don1(1, 4) = ""
    list_don1(1, 5) = ""
    list_don1(6, 1) = Lb_int_Qpomp(0).Caption
    list_don1(6, 2) = str$(Int(val(Tb_Qpomp(0).Text)))
    list_don1(6, 3) = Lb_unit_Qpomp(0).Caption
    list_don1(6, 4) = Tb_Qpompc(0).Text
    list_don1(6, 5) = Lb_unit_Qpompc(0).Caption
End Sub
Public Sub cre_list_don2()
Dim i As Integer, j As Integer
ReDim list_don2(9, 3)
    list_don2(0, 1) = "---" + Frm_geom1.Caption + "---"
    list_don2(0, 2) = ""
    list_don2(0, 3) = ""
    list_don2(1, 1) = Lb_Int_Geom(0).Caption
    list_don2(1, 2) = Tb_Geom(0).Text
    list_don2(1, 3) = Lb_Unit_Geom(0).Caption
    list_don2(2, 1) = Lb_Int_Geom(1).Caption
    list_don2(2, 2) = Cb_Materiau.Text
    list_don2(2, 3) = ""
    list_don2(3, 1) = Lb_Int_Geom(2).Caption
    list_don2(3, 2) = Tb_Geom(2).Text
    list_don2(3, 3) = Lb_Unit_Geom(2).Caption
    list_don2(4, 1) = "---" + Frm_geom2.Caption + "---"
    list_don2(4, 2) = ""
    list_don2(4, 3) = ""

For i = 3 To Tb_Geom.count - 1
    j = i + 2
    list_don2(j, 1) = Lb_Int_Geom(i).Caption
    list_don2(j, 2) = Tb_Geom(i).Text
    list_don2(j, 3) = Lb_Unit_Geom(i).Caption
Next
End Sub
Public Sub cre_list_don4()
Dim i As Integer
ReDim list_don4(9, 3)
    i = 0
    list_don4(i, 1) = Lb_Int_Nbpom.Caption
    list_don4(i, 2) = Tb_Nbpom.Text
    list_don4(i, 3) = ""
    i = i + 1
    list_don4(i, 1) = Lb_Int_Ntdph.Caption
    list_don4(i, 2) = Tb_Ntdph.Text
    list_don4(i, 3) = ""
    i = i + 1
    list_don4(i, 1) = Lb_int_Vutba.Caption
    list_don4(i, 2) = Tb_Vutba.Text
    list_don4(i, 3) = Lb_unit_vutba.Caption
    If Opt_sect_ba(0) Then
        i = i + 1
        list_don4(i, 1) = "--- Bâche circulaire ---"
        list_don4(i, 2) = ""
        list_don4(i, 3) = ""
        i = i + 1
        list_don4(i, 1) = "--- " + Lb_int_diam.Caption
        list_don4(i, 2) = Tb_diam.Text
        list_don4(i, 3) = Lb_unit_diam.Caption
        i = i + 1
        list_don4(i, 1) = ""
        list_don4(i, 2) = ""
        list_don4(i, 3) = ""
    Else
        i = i + 1
        list_don4(i, 1) = "--- Bâche rectangulaire ---"
        list_don4(i, 2) = ""
        list_don4(i, 3) = ""
        i = i + 1
        list_don4(i, 1) = "--- " + Lb_int_long.Caption
        list_don4(i, 2) = Tb_long.Text
        list_don4(i, 3) = Lb_unit_long.Caption
        i = i + 1
        list_don4(i, 1) = "--- " + Lb_int_larg.Caption
        list_don4(i, 2) = Tb_larg.Text
        list_don4(i, 3) = Lb_unit_larg.Caption
    End If
        i = i + 1
        list_don4(i, 1) = Lb_int_denivt.Caption
        list_don4(i, 2) = Tb_denivt.Text
        list_don4(i, 3) = Lb_unit_denivt.Caption
        i = i + 1
        list_don4(i, 1) = Lb_int_denivhau.Caption
        list_don4(i, 2) = Tb_denivhau.Text
        list_don4(i, 3) = Lb_unit_denivhau.Caption
        i = i + 1
        list_don4(i, 1) = Lb_int_denivbas.Caption
        list_don4(i, 2) = Tb_denivbas.Text
        list_don4(i, 3) = Lb_unit_denivbas.Caption
    
    
End Sub
Public Sub cre_list_don3()
Dim i As Integer, j As Integer
Dim libel As String
ReDim list_don3(7, 7)
    If Opt_PtSing(0) Then
        libel = "Anti-bélier (oui)"
    Else
        libel = "Anti-bélier (non)"
    End If
    list_don3(0, 1) = Lb_Ptsing.Caption + " de "
    list_don3(0, 2) = Lb_int_PtSing(0).Caption
    list_don3(0, 3) = Lb_int_PtSing(1).Caption
    list_don3(0, 4) = Lb_int_PtSing(2).Caption
    list_don3(0, 5) = Lb_int_PtSing(3).Caption
    list_don3(0, 6) = Lb_int_PtSing(4).Caption
    list_don3(0, 7) = ""
    list_don3(1, 1) = ""
    list_don3(1, 2) = Tb_PtSing(0).Text
    list_don3(1, 3) = Tb_PtSing(1).Text
    list_don3(1, 4) = Tb_PtSing(2).Text
    list_don3(1, 5) = Tb_PtSing(3).Text
    list_don3(1, 6) = Tb_PtSing(4).Text
    list_don3(1, 7) = ""
For i = 5 To Tb_PtSing.count - 1
    j = i - 3
    list_don3(j, 1) = Lb_int_PtSing(i).Caption
    list_don3(j, 2) = Tb_PtSing(i).Text
    list_don3(j, 3) = ""
    list_don3(j, 4) = ""
    list_don3(j, 5) = ""
    list_don3(j, 6) = ""
    list_don3(j, 7) = ""
Next
    list_don3(6, 1) = ""
    list_don3(6, 2) = ""
    list_don3(6, 3) = ""
    list_don3(6, 4) = ""
    list_don3(6, 5) = ""
    list_don3(6, 6) = ""
    list_don3(6, 7) = ""
    list_don3(7, 1) = libel
    list_don3(7, 2) = ""
    list_don3(7, 3) = ""
    list_don3(7, 4) = ""
    list_don3(7, 5) = ""
    list_don3(7, 6) = ""
    list_don3(7, 7) = ""

End Sub
Public Sub cre_list_don5()
Dim i As Integer
Dim plin As Double
Dim xplin As String
ReDim list_don5(15, 5)
    i = 0
    list_don5(i, 1) = Lb_int_Qpomp(1).Caption
    list_don5(i, 2) = str$(Int(val(Tb_Qpomp(1).Text)))
    list_don5(i, 3) = Lb_unit_Qpomp(1).Caption + " soit"
    list_don5(i, 4) = Tb_Qpompc(1).Text
    list_don5(i, 5) = Lb_unit_Qpompc(1).Caption
    i = i + 1
    list_don5(i, 1) = Lb_Int_Drflt.Caption + "  " + Lb_materiau.Caption
    list_don5(i, 2) = Tb_Drflt.Text
    list_don5(i, 3) = Lb_Unit_Drflt.Caption
    list_don5(i, 4) = ""
    list_don5(i, 5) = ""
    i = i + 1
    list_don5(i, 1) = Lb_int_Vitrflt.Caption
    list_don5(i, 2) = Tb_VitRflt.Text
    list_don5(i, 3) = Lb_unit_Vitrflt.Caption
    list_don5(i, 4) = ""
    list_don5(i, 5) = ""
    i = i + 1
    list_don5(i, 1) = Lb_int_Jmpkm.Caption
    plin = Format(ebpompe.resultat.jmpkm * ebpompe.don_geometrie.Lrflt / 1000#, "##0.000")
 '   plin = ebpompe.resultat.jmpkm * ebpompe.don_geometrie.Lrflt / 1000#
    xplin = Trim$(str(plin))
    If Left$(xplin, 1) = "." Then
        xplin = "0" + xplin
    End If
    list_don5(i, 2) = xplin 'Tb_Jmpkm.Text
    list_don5(i, 3) = "m" 'Lb_unit_Jmpkm.Caption
    list_don5(i, 4) = ""
    list_don5(i, 5) = ""
    i = i + 1
    list_don5(i, 1) = Lb_int_Singul.Caption
    list_don5(i, 2) = Tb_Singul.Text
    list_don5(i, 3) = Lb_unit_Singul.Caption
    list_don5(i, 4) = ""
    list_don5(i, 5) = ""
    i = i + 1
    list_don5(i, 1) = Lb_int_Hmt.Caption
    list_don5(i, 2) = Tb_Hmt.Text
    list_don5(i, 3) = Lb_unit_Hmt.Caption
    list_don5(i, 4) = ""
    list_don5(i, 5) = ""
    i = i + 1
    list_don5(i, 1) = ""
    list_don5(i, 2) = ""
    list_don5(i, 3) = ""
    list_don5(i, 4) = ""
    list_don5(i, 5) = ""
    i = i + 1
    list_don5(i, 1) = Lb_int_denivr.Caption
    list_don5(i, 2) = Tb_denivr.Text
    list_don5(i, 3) = Lb_unit_denivr.Caption
    list_don5(i, 4) = ""
    list_don5(i, 5) = ""
    i = i + 1
    list_don5(i, 1) = Lb_int_vurba.Caption
    list_don5(i, 2) = Tb_vurba.Text
    list_don5(i, 3) = Lb_unit_vurba.Caption
    list_don5(i, 4) = ""
    list_don5(i, 5) = ""
    i = i + 1
    list_don5(i, 1) = Lb_int_nrdph.Caption
    list_don5(i, 2) = Tb_nrdph.Text
    list_don5(i, 3) = ""
    list_don5(i, 4) = ""
    list_don5(i, 5) = ""
    i = i + 1
    list_don5(i, 1) = "Temps de vidange" 'Lb_int_tvidange.Caption
    list_don5(i, 2) = Tb_Tvidange.Text
    list_don5(i, 3) = Lb_unit_tvidange.Caption
    list_don5(i, 4) = ""
    list_don5(i, 5) = ""
    i = i + 1
    list_don5(i, 1) = "Durée totale d'un cycle " 'Lb_int_T1cyc.Caption
    list_don5(i, 2) = Tb_T1cyc.Text
    list_don5(i, 3) = Lb_unit_T1cyc.Caption
    list_don5(i, 4) = ""
    list_don5(i, 5) = ""
    i = i + 1
    list_don5(i, 1) = "Nombre de cycles par heure" 'Lb_int_Nbcyc.Caption
    list_don5(i, 2) = Tb_Nbcyc.Text
    list_don5(i, 3) = ""
    list_don5(i, 4) = ""
    list_don5(i, 5) = ""
    i = i + 1
    list_don5(i, 1) = Lb_int_Vmy.Caption
    list_don5(i, 2) = Tb_Vmy.Text
    list_don5(i, 3) = Lb_unit_Vmy.Caption
    list_don5(i, 4) = ""
    list_don5(i, 5) = ""
    i = i + 1
    list_don5(i, 1) = Lb_int_Tsejh.Caption
    list_don5(i, 2) = Tb_Tsejh.Text
    list_don5(i, 3) = Lb_unit_Tsejh.Caption
    list_don5(i, 4) = ""
    list_don5(i, 5) = ""
   
    
End Sub
Private Sub mnuprint_Click()
Dim pict1 As New StdPicture
Dim i As Integer, nb As Integer
'modif FO   ' If ProtectCheck(2) <> 0 Then End
FrmPrint.Type1 = "pompe"
FrmPrint.nomobjet = Trim(Tb_titre.Text)
FrmPrint.titre1 = "FICHE HYDRAULIQUE station de pompage"
FrmPrint.sstitre1 = "Débits caractéristiques"
FrmPrint.ssTitre2 = "Données géométriques"
FrmPrint.ssTitre3 = "Points singuliers"
FrmPrint.ssTitre4 = "Données techniques"
FrmPrint.ssTitre5 = "Résultats"
Frm_imp.Type1 = "pompe"
Frm_imp.nomobjet = Trim(Tb_titre.Text)
Frm_imp.titre1 = "FICHE HYDRAULIQUE station de pompage"
Frm_imp.sstitre1 = "Débits caractéristiques"
Frm_imp.ssTitre2 = "Données géométriques"
Frm_imp.ssTitre3 = "Points singuliers"
Frm_imp.ssTitre4 = "Données techniques"
Frm_imp.ssTitre5 = "Résultats"
cre_list_don1
cre_list_don2
cre_list_don3
cre_list_don4
cre_list_don5
'Frm_desprint.Show
Set pict1 = Frm_desprint.UC_graphique1.lire_pict1()
FrmPrint.paint_picture pict1
SavePicture pict1, chemin_app + "dess.bmp"
Frm_imp.Show 1
End Sub
Private Function complet_listd_don(ByVal liste1 As Variant, ByVal liste2 As Variant, ByVal liste3 As Variant) As Variant
Dim liste() As Variant
Dim i As Integer, j As Integer
i = -1
ReDim liste(UBound(liste1) + 2, 5)
For j = 0 To UBound(liste1)
    i = i + 1
    liste(i, 1) = liste1(j, 1)
    liste(i, 2) = liste1(j, 2)
    liste(i, 3) = liste1(j, 3)
    liste(i, 4) = liste1(j, 4)
    liste(i, 5) = liste1(j, 5)
Next
'i = i + 1
'liste(i, 1) = ""
'liste(i, 2) = ""
'liste(i, 3) = ""
i = i + 1
liste(i, 1) = liste2(0, 1)
liste(i, 2) = liste2(0, 2)
liste(i, 3) = liste2(0, 3)
liste(i, 4) = liste3(0, 2)
liste(i, 5) = liste3(0, 3)
i = i + 1
liste(i, 1) = liste2(1, 1)
liste(i, 2) = liste2(1, 2)
liste(i, 3) = liste2(1, 3)
liste(i, 4) = liste3(1, 2)
liste(i, 5) = liste3(1, 3)
complet_listd_don = liste
End Function
Private Function complet_listd_int1(ByVal liste1 As Variant, ByVal liste2 As Variant, ByVal liste3 As Variant) As Variant
Dim liste() As Variant
Dim i As Integer, j As Integer
        ReDim liste(2, 5)
        i = 0
        liste(i, 1) = ""
        liste(i, 2) = "Conduite amont"
        liste(i, 3) = ""
        liste(i, 4) = "Conduite aval"
        liste(i, 5) = ""
        i = i + 1
        liste(i, 1) = liste2(2, 1)
        liste(i, 2) = liste2(2, 2)
        liste(i, 3) = liste2(2, 3)
        liste(i, 4) = liste3(2, 2)
        liste(i, 5) = liste3(2, 3)
        i = i + 1
        liste(i, 1) = liste2(3, 1)
        liste(i, 2) = liste2(3, 2)
        liste(i, 3) = liste2(3, 3)
        liste(i, 4) = liste3(3, 2)
        liste(i, 5) = liste3(3, 3)
complet_listd_int1 = liste
End Function
Public Function lect_list(ByVal nom As String) As Variant
Select Case nom
Case Is = "list_don1"
    lect_list = list_don1
Case Is = "list_don2"
    lect_list = list_don2
Case Is = "list_don3"
    lect_list = list_don3
Case Is = "list_don4"
    lect_list = list_don4
Case Is = "list_don5"
    lect_list = list_don5
End Select
End Function

Private Sub mnusaves_Click()
'    Me.Enabled = False
'modif FO   ' If ProtectCheck(2) <> 0 Then End
    If fich_lect = nom_fich Or Trim(Tb_titre.Text) = "" Or fich_lect = "" Then
        Frm_titre.Label2.Caption = "Sauvegarde d'une pompe "
        Frm_titre.Label3.Caption = ""
    Else
         Frm_titre.Label2.Caption = "Sauvegarde de la pompe " & Me.Tb_titre.Text
         Frm_titre.Label3.Caption = " de l'étude " & fich_lect_edit
   
    End If
    Frm_titre.Caption = "Etude " + nom_fich_edit
    Frm_titre.Label1.Caption = "Nom de la pompe (30car. maxi)"
    Frm_titre.Text1.Text = Me.Tb_titre.Text
    Frm_titre.Show 1
End Sub
Public Sub save(ByVal bsous As Boolean)
Dim za As st_savpompe
Dim za1 As st_savpom1
Dim i As Integer, isave As Integer
Dim reponse As Integer
If Trim(Tb_titre.Text) <> "" Then
    Call funlockb

   lhFicDbf = FreeFile
    On Error GoTo test_Error
    Open nom_fich For Random Access Read Write Lock Read Write As #lhFicDbf Len = Len(za1)
    i = 0
    isave = 0
    Do While Not EOF(lhFicDbf)
        Get #lhFicDbf, , za1
        If Not EOF(lhFicDbf) Then
            i = i + 1
            za = za1.stsavpo
            If Trim(za.type) = nom_type And Trim(za.nom) = Trim(Tb_titre.Text) Then
                isave = i
            End If
        End If
    Loop
    If isave > 0 Then
        If bsous Then
           reponse = MsgBox("Le nom est déjà utilisé. Le remplacer?", 4, "Sauvegarde d'une pompe")
        Else
           reponse = 6
        End If
        If reponse = 6 Then
            za.type = "pompe"
            za.nom = Tb_titre.Text
            za.pompe = ebpompe
            za1.stsavpo = za
            Put #lhFicDbf, isave, za1
            ouv_sauve = False
            save_fich = True
            fich_lect = nom_fich
        Else
            Unload Frm_titre
            Call mnusaves_Click
        End If
    Else
        za.type = "pompe"
        za.nom = Tb_titre.Text
        za.pompe = ebpompe
        za1.stsavpo = za
        FileLength = (LOF(lhFicDbf) / Len(za1)) + 1
        Put #lhFicDbf, FileLength, za1
        ouv_sauve = False
        save_fich = True
        fich_lect = nom_fich
   End If
        Close #lhFicDbf
        Call flockb(nom_fich)

        Call lect_fich
        ch_texte = Trim(Tb_titre.Text)
        Cb_pompe.Text = Trim(Tb_titre.Text)
Else
    reponse = MsgBox("Le nom de la pompe n'est pas renseigné.", , "Sauvegarde d'une pompe")
End If
 
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + nom_fich + " est déjà en cours d'utilisation.")
    End If
 
Call flockb(nom_fich)
End Sub
Public Sub mnusave_Click()
'modif FO   ' If ProtectCheck(2) <> 0 Then End
    If Trim(Tb_titre.Text) <> "" And fich_lect = nom_fich Then
        Call save(False)
    Else
        Call mnusaves_Click
    End If

End Sub


Private Sub mnusuppr_Click()
Dim za As st_savpompe
Dim za1 As st_savpom1
Dim nom As String
Dim lhFicDbf1 As Integer, reponse As Integer
'modif FO   ' If ProtectCheck(2) <> 0 Then End
 
If Trim(Cb_pompe.Text) <> "" Then
    Call funlockb
    reponse = MsgBox(Trim(Cb_pompe.Text) + " va être supprimé .", 4, "Suppression d'une pompe")
    If reponse = 6 Then  '6=oui,7=non
    save_fich = True
    ouv_sauve = False
    nom = chemin_app + "tempbas.bin"
    lhFicDbf = FreeFile
    On Error GoTo test_Error
    Open nom_fich For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
    lhFicDbf1 = FreeFile
    Open nom For Random Access Write As #lhFicDbf1 Len = Len(za1)
    Do While Not EOF(lhFicDbf)
    '   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
        Get #lhFicDbf, , za1
        If Not EOF(lhFicDbf) Then
            za = za1.stsavpo
            If Trim(za.type) <> nom_type Or (Trim(za.type) = nom_type And Trim(za.nom) <> Trim(Cb_pompe.Text)) Then
                FileLength = LOF(lhFicDbf1) / Len(za1) + 1
                Put #lhFicDbf1, FileLength, za1
            End If
        End If
    Loop
    Close #lhFicDbf
    Close #lhFicDbf1
    Kill nom_fich
    lhFicDbf = FreeFile
    On Error GoTo test_Error
    Open nom_fich For Random Access Write Lock Read Write As #lhFicDbf Len = Len(za1)
    lhFicDbf1 = FreeFile
    Open nom For Random Access Read As #lhFicDbf1 Len = Len(za1)
    Do While Not EOF(lhFicDbf1)
    '   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
        Get #lhFicDbf1, , za1
       If Not EOF(lhFicDbf1) Then
            FileLength = LOF(lhFicDbf) / Len(za1) + 1
            Put #lhFicDbf, FileLength, za1
       End If
    Loop
    Close #lhFicDbf
    Call flockb(nom_fich)
    
    Close #lhFicDbf1
    Kill nom
    Call lect_fich
    Me.Tb_titre.Text = ""
    Me.Caption = fen_titre
    Call reini_valeurs
    Call ini_ebpompe
    Call ini_form
    ouv_sauve = False
    save_fich = False

    End If
End If
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + nom_fich + " est déjà en cours d'utilisation.")
    End If

Call flockb(nom_fich)
End Sub
Public Sub Cb_pompe_click()
Dim za As st_savpompe
Dim za1 As st_savpom1
Call funlockb

'    Cb_pompe.Visible = False
'    For i = 0 To Cb_pompe.ListCount - 1
'        If Trim(Cb_pompe.list(i)) = Trim(nom_ouvrage) Then
'            ch_texte = Cb_pompe.list(i)
'            Cb_pompe.Text = Cb_pompe.list(i)
'        End If
'    Next
    Me.Tb_titre.Text = Trim(nom_ouvrage)
    ch_texte = Trim(nom_ouvrage)
    Cb_pompe.Text = Trim(nom_ouvrage)
    lhFicDbf = FreeFile
    On Error GoTo test_Error
    Open fich_lect For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
    If Not EOF(lhFicDbf) Then
        za = za1.stsavpo
        If Trim(za.type) = nom_type And Trim(za.nom) = Trim(Cb_pompe.Text) Then
            Tb_titre = Trim(za.nom)
            Me.Caption = fen_titre + " : " + Tb_titre.Text
            ebpompe = za.pompe
            Call ini_form
            Call dessin_pompe
            owner.affich_aide Me.Name, mes

'            Call reini_valeurs
'           Me.Cmd_del.Visible = True
            If Cmd_calcul.Enabled Then
'                Call Cmd_calcul_Click
            End If
            Me.mnuprint.Enabled = True
            save_fich = True
            ouv_sauve = False
            If fich_lect <> nom_fich Then
                ouv_sauve = True
            End If
        End If
    End If

Loop
Close #lhFicDbf
If fich_lect <> nom_fich Then
    Kill fich_lect
End If

Call flockb(nom_fich)
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + fich_lect + " est déjà en cours d'utilisation.")
    End If

Call flockb(nom_fich)


End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim reponse As Integer
If ouv_sauve Then 'And save_fich Then
    reponse = MsgBox("La pompe n'a pas été enregistrée" + Chr(10) _
        + "Voulez vous la sauvegarder?", 3, "Sauvegarde d'une pompe")
    Select Case reponse
        Case Is = 6  ' 6=oui,7=non,2=annuler
            Call mnusave_Click
        Case Is = 7
            ouv_sauve = False
        Case Is = 2
            Cancel = True
    End Select
End If
 '   Cancel = True
End Sub
Private Sub Form_Load()
ichar = 0
pi = 3.14159
  okg = True
  Me.KeyPreview = True
    Call ini_tooltip_pompe(Me)
    ouv_sauve = False
    save_fich = False
'    save_fich = True
    nom_ouvrage = ""
    Set owner = MDIFrm_menu.rec_owner
    Call retaille
'    Me.mnusave.Enabled = False
'    Me.mnusaves.Enabled = False
 '   Me.mnuprint.Enabled = True
'    Me.mnusuppr.Enabled = False
'''''    owner.affich_aide Me.Name, "pompe"
'    nom_fich = chemin_app + "ouvrages.bin"
'    nom_fich = chemin_app + "etude.boa"
    nom_type = "pompe"
    fen_titre = Me.Caption
'   lecture fichier
    If Dir(nom_fich) <> "" Then
        Call lect_fich
    End If

    Cb_pompe.Visible = False
    Frm_desprint.Show
    Frm_desprint.Visible = False
    Call debut
End Sub
Private Sub debut0()
    Cb_pompe.Text = ""
    Tb_titre.Text = ""
    Me.Caption = fen_titre
'    ouv_sauve = False
    Call debut
End Sub
Private Sub debut()
Dim itab As Integer
Dim i As Integer
    bKP = False
    sval_champ = ""
Call init_l_tab
Call donne_focus(Me)
Call ini_ebpompe
Typo = 1
Vutba = 0#
Vurba = 0#
SECBA = 0#
DENIV = 0#
Nrdph = 0#
T1cyc = 0#
Tvidange = 0#
Nbcyc = 0#
Tsejh = 0#
Vmy = 0#
Denivhau = 0.2
Denivbas = 0.8
Singul = 0#
Hmt = 0#
jmpkm = 0#
VitRflt = 0#
ok_saisie_denivr = False
For i = 0 To 4
    Tb_debit(i).Text = "0.00"
    Select Case i
        Case 1
            Tb_debit(i).Enabled = False
        Case 3 To 4
            Tb_debit(i).Enabled = False
    End Select
    Tb_Debitc(i).Enabled = False
    Tb_Debitc(i).Text = "0.00"
Next
    Tb_FPointe.Text = "0.00"
    Tb_Qpomp(0).Text = "0.00"
    Tb_Qpomp(1).Text = "0.00"
    Tb_Qpomp(0).Enabled = False
    Tb_Qpompc(0).Enabled = False
    Tb_Qpompc(0).Text = "0.00"
    Tb_Qpompc(1).Enabled = False
    Tb_Qpompc(1).Text = "0.00"

With Cb_Materiau
    .Clear
    .AddItem "FONTE"
    .AddItem "PEHD"
    .AddItem "PVC"
    .AddItem "ACIER"
    .Text = ""
End With
For i = 0 To 6
    Tb_Geom(i).Text = "0.00"
Next
    Tb_Geom(2).Text = "0"
Cb_Materiau.Text = "FONTE"
SSTab1.Tab = 0
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
Tb_Qpomp(1).Enabled = False
Tb_Drflt.Enabled = False
For i = 0 To 8
    Tb_PtSing(i).Text = "0"
Next
    Opt_PtSing(0).Value = False
    Opt_PtSing(1).Value = True
    Tb_Drflt.Text = "0"
    Tb_VitRflt.Text = "0.00"
    Tb_Jmpkm.Text = "0.00"
    Tb_Nbpom.Text = "0"
    Tb_Ntdph.Text = "0"
    Tb_Vutba.Text = "0.00"
    Tb_vurba.Text = "0.00"
 '   Frm_bache.Visible = False
    Opt_sect_ba(0).Value = True
    Opt_sect_ba(1).Value = False
    Lb_int_diam.Visible = True
    Tb_diam.Visible = True
    Lb_unit_diam.Visible = True
    Lb_int_long.Visible = False
    Tb_long.Visible = False
    Lb_unit_long.Visible = False
    Lb_int_larg.Visible = False
    Tb_larg.Visible = False
    Lb_unit_larg.Visible = False
'    Lb_int_denivt.Visible = False
'    Tb_denivt.Visible = False
'    Lb_unit_denivt.Visible = False
    Tb_diam.Text = "0.00"
    Tb_larg.Text = "0.00"
    Tb_long.Text = "0.00"
    Tb_denivt.Text = "0.00"
    Tb_denivr.Text = "0.00"
    Tb_nrdph.Text = "0.00"
    Tb_Tvidange.Text = "0.00"
    Tb_T1cyc.Text = "0.00"
    Tb_Nbcyc.Text = "0.00"
    Tb_Vmy.Text = "0.00"
    Tb_Tsejh.Text = "0.00"
    Tb_denivhau.Text = "0.2"
    Tb_denivbas.Text = "0.8"
    Tb_Singul.Text = "0.000"
    Tb_Hmt.Text = "0.000"
    owner.fdessin.mnu_fichier.Caption = Me.mnufichier.Caption
    owner.fdessin.UC_graphique2.Visible = False
    owner.fdessin.UC_graphiqueB.Visible = False
    owner.fdessin.Image1.Visible = False
    owner.fdessin.Image2.Visible = False
    owner.fdessin.Image3.Visible = False
    owner.fdessin.UC_graphique1.Visible = True
    owner.fdessin.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.Height = 6000
'    owner.fdessin.UC_graphique1.Top = 0
'    owner.fdessin.UC_graphique1.Left = 1440
'    owner.fdessin.UC_graphique1.Height = 4210
'    owner.fdessin.UC_graphique1.Width = 7800
'    owner.fdessin.UC_graphique1.reinit 7, "Arial"
'    owner.fdessin.UC_graphique1.init_title
'    owner.fdessin.UC_graphique1.init_titleh ""
'    owner.fdessin.UC_graphique1.init_titleb ""
    Call reini_valeurs
 '   Call ini_ebpompe
 owner.affich_aide Me.Name, mes

    ouv_sauve = False
    save_fich = False
    fich_lect = ""
    change_coul = False
End Sub
Private Sub calc_val_reelles()
''' If val(Tb_denivr.Text) > 0 Then
    Vurba = ebpompe.resultat.Denivr * SECBA
'''    Tb_vurba.Text = Format(Vurba, "###0.00")
    Nrdph = ebpompe.don_techniques.Ntdph * Vutba / Vurba
'''    Tb_nrdph.Text = Format(Nrdph, "##0.00")
    
'T1CYC = VURBA * (1 / QMEULS + 1 / (ebpompe.resultat.Qpomr - QMEULS)) / 3.6
    T1cyc = (Vurba * (1 / ebpompe.debits_car.Qtsm) / 3.6) + (Vurba * (1 / (ebpompe.resultat.Qpomr - ebpompe.debits_car.Qtsm)) / 3.6)
'TVIDANGE = VURBA / (3.6 * (Qpomp - QMEULS))
    Tvidange = Vurba / (3.6 * (ebpompe.resultat.Qpomr - ebpompe.debits_car.Qtsm))
 '    DMM=diametre de la conduite de refoulement
'NBCYC = (Lrflt * PI * (0.001 * DMM) ^ 2 / 4) / (Qpomp * TVIDANGE * 3.6)
    Nbcyc = ((ebpompe.don_geometrie.Lrflt * pi * (0.001 * ebpompe.resultat.Drflr) ^ 2 / 4) / (ebpompe.resultat.Qpomr * Tvidange * 3.6))
    
    ' a revoir (nb cycle + vidange)
  '  Tsehj = Int(Nbcyc) * T1cyc + (Nbcyc - Int(Nbcyc)) * Tvidange
    
    Tsejh = T1cyc * Nbcyc
    Vmy = ebpompe.don_geometrie.Lrflt / Tsejh / 3600

'Debug.Print "Vitesse moyenne d'‚coulement sur 24 heures ........ :##.## m/s"; Vmy
''''VMY = (ebpompe.resultat.Qpomr * TVIDANGE * 3.6) / T1CYC / (PI * (0.001 * ebpompe.resultat.Drflr) ^ 2 / 4) / 3600
'''    Me.Tb_T1cyc.Text = Format(T1cyc, "###0.00")
'''    Me.Tb_Tvidange.Text = Format(Tvidange, "###0.00")
'''    Me.Tb_Tsejh.Text = Format(Tsejh, "###0.00")
'''    Me.Tb_Nbcyc.Text = Format(Nbcyc, "###0.00")
'''    Me.Tb_Vmy.Text = Format(Vmy, "###0.00")
' 20050513 à voir
    Nrdph = 1 / (T1cyc * val(Me.Tb_Nbpom.Text))
'''    Tb_nrdph.Text = Format(Nrdph, "##0.00")
    
    ebpompe.resultat.Vurba = Vurba
    ebpompe.resultat.Nrdph = Nrdph
    ebpompe.resultat.Tvidange = Tvidange
    ebpompe.resultat.T1cyc = T1cyc
    ebpompe.resultat.Nbcyc = Nbcyc
    ebpompe.resultat.Vmy = Vmy
    ebpompe.resultat.Tsejh = Tsejh
'''    Hmt = calc_hmt
'''Else
'''    Tb_vurba.Text = Format(0, "###0.00")
'''    Tb_nrdph.Text = Format(0, "##0.00")
'''     Me.Tb_T1cyc.Text = Format(0, "###0.00")
'''    Me.Tb_Tvidange.Text = Format(0, "###0.00")
'''    Me.Tb_Tsejh.Text = Format(0, "###0.00")
'''    Me.Tb_Nbcyc.Text = Format(0, "###0.00")
'''    Me.Tb_Vmy.Text = Format(0, "###0.00")
'''    Hmt = calc_hmt
'''
'''End If
End Sub
Function calc_hmt() As Double
Dim HGS As Double, LINEAI As Double
    Singul = calc_singul
    HGS = ebpompe.don_geometrie.NivEX - ebpompe.don_geometrie.NivEN + ebpompe.don_techniques.Denivhau + ebpompe.resultat.Denivr / 2 'Hauteur géomètrique statique (m)
    LINEAI = ebpompe.don_geometrie.Lrflt * ebpompe.resultat.jmpkm / 1000  'Pertes de charge linéaires  (m)
    Hmt = HGS + LINEAI + Singul
'    Tb_Hmt.Text = Format(Hmt, "###0.000")
'    ebpompe.resultat.Hmt = Hmt
'''    calc_hmt = Hmt
'''    Call dessin_pompe
 calc_hmt = Hmt
   If Hmt > 8 And ebpompe.pts_singuliers.Antb = 0 Then
        mes = "ATTENTION ! UN DISPOSITIF ANTI-BELIER EST SANS DOUTE NECESSAIRE."
        MsgBox mes
    End If
    
End Function
Function change_ch_sing() As Double
Dim Singul As Double, Hmt As Double
    Singul = calc_singul
  Hmt = calc_hmt()
End Function
Function change_hmt()
Dim Hmt As Double
  Hmt = calc_hmt()
    Tb_Hmt.Text = Format(Hmt, "###0.000")
    ebpompe.resultat.Hmt = Hmt
Call dessin_pompe
End Function



Function calc_singul() As Double
Dim RSURD As Double, KCLA As Integer, KSORTIE As Integer
Dim KC11 As Double, KC22 As Double, KC30 As Double, KC45 As Double
Dim KC90 As Double, KVAN As Double, KY As Double, KDAB As Double
Dim KVID As Double, KVEN As Double, vit As Double, KPIED As Double
Dim NDAB As Integer, nc11  As Integer, nc22 As Integer, nc30 As Integer, nc45 As Integer, nc90 As Integer
Dim nvan As Integer, nvid As Integer, nven As Integer, ncla As Integer
Dim g As Double, Singul As Double
g = 9.810001
    RSURD = 1#  ' Rem RSURD=RAYON DE COURBURE/DIAMETRE INTERIEUR DU TUYAU
    KCLA = 2
    With ebpompe.pts_singuliers
        NDAB = .Antb
        nc11 = .Nbc1
        nc22 = .Nbc2
        nc30 = .Nbc3
        nc45 = .Nbc4
        nc90 = .Nbc9
        nvan = .Nbva
        nvid = .Nbvi
        nven = .Nbve
        ncla = .Nbcl
    End With
    KSORTIE = 1
    KC90 = (0.131 + 1.847 * (2 * RSURD) ^ (-3.5))

    KC11 = KC90 * 11.25 / 90
    KC22 = KC90 * 22.5 / 90
    KC30 = KC90 * 30 / 90
    KC45 = KC90 * 45 / 90
    KVAN = 0.12
    KY = 0.5
    KDAB = 0.044
    KVID = 0.04
    KVEN = 0.04
    vit = ebpompe.resultat.VitRflt
    KPIED = 0.3 ' Rem Note žA : si Y->KY=0.11 si T->KY=0.50
    KPOSTE = KPIED + KC90 + KCLA + KVAN + KC90 + KY + NDAB * KDAB + KSORTIE
    Singul = (nc11 * KC11 + nc22 * KC22 + nc30 * KC30 + nc45 * KC45 + nc90 * KC90 + ncla * KCLA + nvan * KVAN + nvid * KVID + nven * KVEN + KPOSTE) * vit * vit / 2 / g
''    Tb_Singul.Text = Format(Singul, "###0.000")
    ebpompe.resultat.Singul = Singul
    calc_singul = Singul
End Function

Function calc_vit_pcl()
Dim krugo  As Double
''If ebpompe.resultat.Qpomr > 0 And val(Tb_Drflt.Text) > 0 Then
    VitRflt = 4000 * ebpompe.resultat.Qpomr / pi / (ebpompe.resultat.Drflr ^ 2)
'''    ebpompe.resultat.Drflr = txtVersNum(Me.Tb_Drflt.Text)
    krugo = rech_rugo(ebpompe.don_geometrie.NatRflt)
    jmpkm = perte_charge_lin(VitRflt, ebpompe.resultat.Drflr, krugo)
'''    ebpompe.resultat.Vitrflt = Vitrflt
'''    ebpompe.resultat.jmpkm = jmpkm
'''    Me.Tb_Jmpkm.Text = rempl_virgule(Format(jmpkm, "####0.00"))
'''    Me.Tb_VitRflt.Text = rempl_virgule(Format(Vitrflt, "####0.00"))
'''End If
End Function

Private Sub calc_secba()
    If ebpompe.don_techniques.Sectb = 0 Then
        SECBA = pi * ebpompe.don_techniques.Diamb ^ 2 / 4
    Else
        If ebpompe.don_techniques.Largb > 0 And ebpompe.don_techniques.Longb > 0 Then
            SECBA = ebpompe.don_techniques.Largb * ebpompe.don_techniques.Longb
        End If
    End If
End Sub
Private Sub calc_denivt()
'DENIV = 0
'If SECBA > 0 Then
    DENIV = 0.01 * Int(100 * (Vutba / SECBA))
'End If
'
'    Tb_denivt.Text = Format(DENIV, "###0.00")
'    Tb_denivr.Text = Format(DENIV, "###0.00")
'    ebpompe.don_techniques.Denivt = DENIV

End Sub
Private Sub calc_vutba()
Dim ds As Double, de As Double, dp As Double
Dim np As Integer
Dim nc As Double
    If ebpompe.don_techniques.Nbpom > 0 And ebpompe.don_techniques.Ntdph > 0 And ebpompe.resultat.Qpomr > 0 Then
dp = ebpompe.resultat.Qpomr
de = ebpompe.debits_car.qeum
ds = dp - de
np = ebpompe.don_techniques.Nbpom
nc = ebpompe.don_techniques.Ntdph
'''If ebpompe.don_techniques.Nbpom > 0 And ebpompe.don_techniques.Ntdph > 0 And ebpompe.resultat.Qpomr > 0 Then
'    VUTBA = Qpomp * 3.6 / 4 / NBPOM / NTDPH '(nbpompes/nb demarrages)
    Vutba = ebpompe.resultat.Qpomr * 3.6 / 4 / ebpompe.don_techniques.Nbpom / ebpompe.don_techniques.Ntdph
    Vutba = dp * 3.6 / 4 / np / nc
' 20050513 à voir
'    VUTBA = (ebpompe.resultat.Qpomr - val(Tb_Qpomp(0).Text)) * 3.6 / 4 / ebpompe.don_techniques.Nbpom / ebpompe.don_techniques.Ntdph
' 20050517 à voir
'    Vutba = (ebpompe.resultat.Qpomr - ebpompe.debits_car.qeum) / (ebpompe.resultat.Qpomr * ebpompe.don_techniques.Nbpom * ebpompe.don_techniques.Ntdph) * 3.6
'   Vutba = (ds * de) / (dp * np * nc) * 3.6
'    Frm_bache.Visible = True
    Lb_int_denivt.Visible = True
    Tb_denivt.Visible = True
    Lb_unit_denivt.Visible = True
'''Else
'''    Vutba = 0
'''    Tb_Vutba.Text = Format(Vutba, "###0.00")
'    Frm_bache.Visible = False
 '   Lb_int_denivt.Visible = False
 '   Tb_denivt.Visible = False
 '   Lb_unit_denivt.Visible = False
End If

End Sub
Private Sub ini_form()
    Me.Tb_FPointe.Text = rempl_virgule(Format(ebpompe.debits_car.Fp, "###0.00"))
    Me.Tb_debit(0).Text = rempl_virgule(Format(ebpompe.debits_car.qeum, "###0.00"))
    Me.Tb_debit(1).Text = rempl_virgule(Format(ebpompe.debits_car.Qeu, "###0.00"))
    Me.Tb_debit(2).Text = rempl_virgule(Format(ebpompe.debits_car.Qecp, "###0.00"))
    Me.Tb_debit(3).Text = rempl_virgule(Format(ebpompe.debits_car.Qtsm, "###0.00"))
    Me.Tb_debit(4).Text = rempl_virgule(Format(ebpompe.debits_car.Qts, "###0.00"))
    Me.Tb_Qpomp(0).Text = rempl_virgule(Format(ebpompe.debits_car.Qpomp, "###0.00"))
    Me.Tb_Geom(0).Text = rempl_virgule(Format(ebpompe.don_geometrie.Lrflt, "###0.00"))
    Me.Tb_Geom(2).Text = rempl_virgule(Format(ebpompe.don_geometrie.Drflt, "####0"))
    Me.Tb_Geom(3).Text = rempl_virgule(Format(ebpompe.don_geometrie.NivTN, "####0.00"))
    Me.Tb_Geom(4).Text = rempl_virgule(Format(ebpompe.don_geometrie.NivEN, "###0.00"))
    Me.Tb_Geom(5).Text = rempl_virgule(Format(ebpompe.don_geometrie.NivSO, "###0.00"))
    Me.Tb_Geom(6).Text = rempl_virgule(Format(ebpompe.don_geometrie.NivEX, "###0.00"))
    Cb_Materiau.Text = ebpompe.don_geometrie.NatRflt
    Me.Tb_PtSing(0).Text = rempl_virgule(Format(ebpompe.pts_singuliers.Nbc1, "#0"))
    Me.Tb_PtSing(1).Text = rempl_virgule(Format(ebpompe.pts_singuliers.Nbc2, "#0"))
    Me.Tb_PtSing(2).Text = rempl_virgule(Format(ebpompe.pts_singuliers.Nbc3, "#0"))
    Me.Tb_PtSing(3).Text = rempl_virgule(Format(ebpompe.pts_singuliers.Nbc4, "#0"))
    Me.Tb_PtSing(4).Text = rempl_virgule(Format(ebpompe.pts_singuliers.Nbc9, "#0"))
    Me.Tb_PtSing(5).Text = rempl_virgule(Format(ebpompe.pts_singuliers.Nbva, "#0"))
    Me.Tb_PtSing(6).Text = rempl_virgule(Format(ebpompe.pts_singuliers.Nbcl, "#0"))
    Me.Tb_PtSing(7).Text = rempl_virgule(Format(ebpompe.pts_singuliers.Nbvi, "#0"))
    Me.Tb_PtSing(8).Text = rempl_virgule(Format(ebpompe.pts_singuliers.Nbve, "#0"))
    If ebpompe.pts_singuliers.Antb = 1 Then
        Me.Opt_PtSing(0).Value = True
        Me.Opt_PtSing(1).Value = False
    Else
        Me.Opt_PtSing(0).Value = False
        Me.Opt_PtSing(1).Value = True
    End If
    Me.Tb_Nbpom.Text = rempl_virgule(Format(ebpompe.don_techniques.Nbpom, "##0"))
    Me.Tb_Ntdph.Text = rempl_virgule(Format(ebpompe.don_techniques.Ntdph, "##0"))
    Me.Tb_Vutba.Text = rempl_virgule(Format(ebpompe.don_techniques.Vutba, "###0.00"))
    Me.Tb_diam.Text = rempl_virgule(Format(ebpompe.don_techniques.Diamb, "###0.00"))
    Me.Tb_long.Text = rempl_virgule(Format(ebpompe.don_techniques.Longb, "###0.00"))
    Me.Tb_larg.Text = rempl_virgule(Format(ebpompe.don_techniques.Largb, "###0.00"))
    Me.Tb_denivt.Text = rempl_virgule(Format(ebpompe.don_techniques.Denivt, "###0.00"))
    Me.Tb_denivhau.Text = rempl_virgule(Format(ebpompe.don_techniques.Denivhau, "###0.00"))
    Me.Tb_denivbas.Text = rempl_virgule(Format(ebpompe.don_techniques.Denivbas, "###0.00"))
    If ebpompe.don_techniques.Sectb = 0 Then
        Me.Opt_sect_ba(0).Value = True
        Me.Opt_sect_ba(1).Value = False
    Else
        Me.Opt_sect_ba(0).Value = False
        Me.Opt_sect_ba(1).Value = True
    End If
    Me.Tb_Qpomp(1).Text = rempl_virgule(Format(ebpompe.resultat.Qpomr, "###0.00"))
    Me.Tb_Drflt.Text = rempl_virgule(Format(ebpompe.resultat.Drflr, "####0"))
    Lb_materiau.Caption = ebpompe.resultat.NatRflr
    Me.Tb_VitRflt.Text = rempl_virgule(Format(ebpompe.resultat.VitRflt, "###0.00"))
    Me.Tb_Jmpkm.Text = rempl_virgule(Format(ebpompe.resultat.jmpkm, "####0.00"))
    Me.Tb_denivr.Text = rempl_virgule(Format(ebpompe.resultat.Denivr, "####0.00"))
    Me.Tb_vurba.Text = rempl_virgule(Format(ebpompe.resultat.Vurba, "####0.00"))
    Me.Tb_nrdph.Text = rempl_virgule(Format(ebpompe.resultat.Nrdph, "####0.00"))
    Me.Tb_Tvidange.Text = rempl_virgule(Format(ebpompe.resultat.Tvidange, "####0.00"))
    Me.Tb_T1cyc.Text = rempl_virgule(Format(ebpompe.resultat.T1cyc, "####0.00"))
    Me.Tb_Nbcyc.Text = rempl_virgule(Format(ebpompe.resultat.Nbcyc, "####0.00"))
    Me.Tb_Vmy.Text = rempl_virgule(Format(ebpompe.resultat.Vmy, "####0.00"))
    Me.Tb_Tsejh.Text = rempl_virgule(Format(ebpompe.resultat.Tsejh, "####0.00"))
    Me.Tb_Singul.Text = rempl_virgule(Format(ebpompe.resultat.Singul, "###0.000"))
    Me.Tb_Hmt.Text = rempl_virgule(Format(ebpompe.resultat.Hmt, "###0.000"))
owner.affich_aide Me.Name, mes

End Sub
Private Sub init_graph(ByRef uc_g As UC_graphique)
Dim ok As Boolean
Dim ecx As Double
Dim hgarde As Double, hdeniv As Double, hfond As Double, diam As Double
Dim plin As Double, psing As Double, hmtot As Double, hpertes As Double
Dim cote_ex As Double, cote_maxi As Double
Dim i As Integer
    hgarde = ebpompe.don_techniques.Denivhau
    hdeniv = ebpompe.resultat.Denivr
    hfond = ebpompe.don_techniques.Denivbas
    If ebpompe.don_techniques.Sectb = 0 Then
        diam = ebpompe.don_techniques.Diamb
    Else
        diam = ebpompe.don_techniques.Largb
    End If
plin = ebpompe.resultat.jmpkm * ebpompe.don_geometrie.Lrflt / 1000#
psing = ebpompe.resultat.Singul
hpertes = plin + psing
If hpertes > 3# Then
    hpertes = 3#
End If
hmtot = ebpompe.resultat.Hmt
cote_ex = ebpompe.don_geometrie.NivEX
ok = False
If ebpompe.don_geometrie.NivEX - ebpompe.don_geometrie.NivTN > 3 Then
    cote_ex = ebpompe.don_geometrie.NivTN + 3#
End If
cote_maxi = cote_ex
If ebpompe.don_geometrie.NivTN > cote_ex Then
    cote_maxi = ebpompe.don_geometrie.NivTN
End If
uc_g.graphique_clear
uc_g.reinit 7, "Arial"
uc_g.init_titleh ""
uc_g.init_titleb ""
uc_g.init_arrondi_X 2
uc_g.init_arrondi_y 3
uc_g.init_MinX -2#
uc_g.init_MaxX 5 * diam
uc_g.init_EchXn 1
ecx = uc_g.lire_EchXn()
uc_g.init_MaxY ((cote_maxi - ebpompe.don_geometrie.NivEN) + hgarde + hdeniv + hfond + hpertes + 2)
uc_g.init_MinY 1#
uc_g.init_EchYn 1
  
End Sub
Private Sub Form_Unload(Cancel As Integer)
'    frm_menu.Enabled = True
    ouv_sauve = False
    Unload Frm_desprint
    Unload owner.fdessin
    owner.recharge_commentaire
End Sub

Private Sub MnuQuit_Click()
    Unload Me
End Sub
Public Sub Mquitter()
    MnuQuit_Click
End Sub
Public Sub Msupprimer()
    mnusuppr_Click
End Sub
Public Sub Menregistrer()
    mnusave_Click
End Sub
Public Sub Mimprimer()
    mnuprint_Click
End Sub
Public Sub Mnouveau()
    mnunouv_Click
End Sub
Public Sub Menregsous()
    mnusaves_Click
End Sub
Public Sub Mouvrir()
    mnuouv_Click
End Sub
Public Sub Minfo()
    mnuinfo_Click
End Sub
Public Sub Mquit()
    m_quitter_Click
End Sub
Public Sub reini_valeurs()
' impression false
  Me.mnuprint.Enabled = False
owner.fdessin.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.graphique_clear
    Call ini_lbresu
'     If ebpompe.dam > 0 And ebpompe.iRadam > 0 And ebpompe.Kam > 0 _
'        And ebpompe.Rdav > 0 Then
'        Me.Cmd_amo.Enabled = True
'    Else
'        Me.Cmd_amo.Enabled = False
'    End If
'     If ebpompe.dav > 0 And ebpompe.iradav > 0 And ebpompe.kav > 0 _
'        And ebpompe.Rdam > 0 Then
'        Me.Cmd_ava.Enabled = True
'    Else
'        Me.Cmd_ava.Enabled = False
'    End If
'   If ebpompe.dam > 0 And ebpompe.iRadam > 0 And ebpompe.Kam > 0 _
'        And ebpompe.dav > 0 And ebpompe.iradav > 0 And ebpompe.kav > 0 _
'        And ebpompe.Rdav > 0 And ebpompe.Rdam > 0 And ebpompe.Qmax > 0 Then
'        Me.Cmd_calcul.Enabled = True
'
'    Else
'        Me.Cmd_calcul.Enabled = False
'
'    End If
    ouv_sauve = True

End Sub

Public Sub maj_debit()
        ebpompe.debits_car.Qeu = ebpompe.debits_car.qeum * ebpompe.debits_car.Fp
        ebpompe.debits_car.Qtsm = ebpompe.debits_car.qeum + ebpompe.debits_car.Qecp
        ebpompe.debits_car.Qts = ebpompe.debits_car.Qeu + ebpompe.debits_car.Qecp
        ebpompe.debits_car.Qpomp = Maxi(3 * (ebpompe.debits_car.qeum + ebpompe.debits_car.Qecp) + 1, ebpompe.debits_car.Qts)
        bKP = False
        Tb_debit(1).Text = Format(ebpompe.debits_car.Qeu, "###0.00")
        Tb_debit(3).Text = Format(ebpompe.debits_car.Qtsm, "##0.00")
        Tb_debit(4).Text = Format(ebpompe.debits_car.Qts, "##0.00")
        Tb_Qpomp(0).Text = Format(ebpompe.debits_car.Qpomp, "##0.00")
        Tb_Qpompc(0).Text = Format(3.6 * ebpompe.debits_car.Qpomp, "##0.00")
        bKP = True

End Sub
Public Sub calcul_resu()
Dim oKp As Boolean
oKp = bKP
bKP = False
    Call reini_resu_ebpompe
    Call maj_resu
    Vutba = 0#
'    If ebpompe.resultat.Qpomr > 0 Then
    If ebpompe.don_techniques.Nbpom > 0 And ebpompe.don_techniques.Ntdph > 0 And ebpompe.resultat.Qpomr > 0 Then
        Call calc_vutba
    End If
'   Tb_Vutba.Text = Format(Vutba, "###0.00")
    ebpompe.don_techniques.Vutba = Vutba
    SECBA = 0#
    Call calc_secba
    If SECBA > 0 Then
        DENIV = 0#
        Call calc_denivt
'   Tb_denivt.Text = Format(DENIV, "###0.00")
'   Tb_denivr.Text = Format(DENIV, "###0.00")
        ebpompe.don_techniques.Denivt = DENIV
        If ok_saisie_denivr = False Then
            ebpompe.resultat.Denivr = 0
'            If DENIV > ebpompe.resultat.Denivr Then
'                ebpompe.resultat.Denivr = DENIV
'            End If
        End If
    End If

    If ebpompe.resultat.Denivr > 0 And ebpompe.don_geometrie.Lrflt > 0 And ebpompe.resultat.Drflr > 0 _
        And ebpompe.debits_car.Qtsm > 0 And ebpompe.don_techniques.Ntdph > 0 And SECBA > 0 _
        And Vutba > 0 And ebpompe.resultat.Qpomr > 0 Then
            Call calc_val_reelles
    End If
    VitRflt = 0#
    jmpkm = 0#
    If ebpompe.resultat.Qpomr > 0 And ebpompe.resultat.Drflr > 0 And ebpompe.don_geometrie.NatRflt <> "" Then
        Call calc_vit_pcl
    End If
    ebpompe.resultat.VitRflt = VitRflt
    ebpompe.resultat.jmpkm = jmpkm
'    Me.Tb_Jmpkm.Text = rempl_virgule(Format(jmpkm, "####0.00"))
'    Me.Tb_VitRflt.Text = rempl_virgule(Format(Vitrflt, "####0.00"))
    Hmt = 0#
    If ebpompe.don_geometrie.NivEX > 0 And ebpompe.don_geometrie.NivEN > 0 And ebpompe.don_techniques.Denivhau > 0 _
         And ebpompe.resultat.Denivr > 0 And ebpompe.don_geometrie.Lrflt > 0 And ebpompe.resultat.jmpkm > 0 Then
        Call calc_hmt
    End If
'    Tb_Hmt.Text = Format(Hmt, "###0.000")
    ebpompe.resultat.Hmt = Hmt

    Call maj_resu
'    End If
    Call dessin_pompe
 Me.mnuprint.Enabled = True
 ok_saisie_denivr = False
bKP = oKp
End Sub
Public Sub calcul_maj(ByVal nom As String)
Dim oKp As Boolean
oKp = bKP
bKP = False
Select Case nom
        Case Is = "Drflt"

If ebpompe.resultat.Drflr >= 100 Then
'    If bKP Then
    Call calcul_resu
'    Call calc_vit_pcl
'    Call calc_val_reelles
'    Call change_hmt
    Lb_Int_Drflt.Caption = "Canalisation"
    Lb_Int_Drflt.ForeColor = &H80000012
    Tb_Drflt.ForeColor = &H80000012
'    End If
Else
'    If bKP Then
    Lb_Int_Drflt.Caption = "  <100  "
    Lb_Int_Drflt.ForeColor = 255
    Tb_Drflt.ForeColor = 255
'    Me.Tb_Drflt.SetFocus
    Call calcul_resu

'   End If
'   Call reini_resu1
End If

         Case Is = "Qpomp"
                If ebpompe.resultat.Qpomr >= ebpompe.debits_car.Qts Then
' 20071025
'                    diam = 2000 * (ebpompe.resultat.Qpomr / (1000 * pi * 1.5)) ^ 0.5
                    diam = 2000 * (ebpompe.resultat.Qpomr / (1000 * pi * 1#)) ^ 0.5
                    If diam < 100 Then
                        diam = 100
                    End If
' 20080318
                    Lb_Int_Drflt.Caption = "Canalisation"
                    Lb_Int_Drflt.ForeColor = &H80000012
                    Tb_Drflt.ForeColor = &H80000012

                    Tb_Geom(2).Text = Format(diam, "####")
                    ebpompe.don_geometrie.Drflt = txtVersNum(Me.Tb_Geom(2).Text)
                    Tb_Drflt.Text = Format(ebpompe.don_geometrie.Drflt, "###0")
                    ebpompe.resultat.Drflr = ebpompe.don_geometrie.Drflt
                    Lb_materiau.Caption = Cb_Materiau.Text & " de D = "
                    ebpompe.resultat.NatRflr = Lb_materiau.Caption
                    Call calcul_resu
                    Lb_int_Qpomp(1).Caption = "Débit de pompage"
                    Lb_int_Qpomp(1).ForeColor = &H80000012
                    Tb_Qpomp(1).ForeColor = &H80000012
                Else
'                Else
                    Call reini_resu_ebpompe
                    ebpompe.resultat.Denivr = 0#
                    ebpompe.resultat.Drflr = 0
                    Call maj_resu
                    Lb_int_Qpomp(1).Caption = " < Qts  "
                    Lb_int_Qpomp(1).ForeColor = 255
                    Tb_Qpomp(1).ForeColor = 255
                    Me.Tb_Qpomp(1).SetFocus

'                    Me.Tb_VitRflt.Text = "0.00"
'                    Me.Tb_Jmpkm.Text = "0.00"
'                    Me.Tb_Drflt.Text = "0"
'                    ebpompe.resultat.VitRflt = 0#
'                    ebpompe.resultat.jmpkm = 0#
'                    ebpompe.resultat.Drflr = 0
'                    Me.Tb_Vutba.Text = "0.00"
'                    ebpompe.don_techniques.Vutba = 0#
'                    Me.Tb_denivt.Text = "0.00"
'                    ebpompe.don_techniques.Denivt = 0#
'                    Me.Tb_denivr.Text = "0.00"
'                    ebpompe.resultat.Denivr = 0#
               End If

            
    End Select
'    Call maj_resu
bKP = oKp

End Sub


Public Sub reini_resu1()
    ebpompe.resultat.VitRflt = 0#
    ebpompe.resultat.jmpkm = 0#
    Me.Tb_VitRflt.Text = rempl_virgule(Format(ebpompe.resultat.VitRflt, "###0.00"))
    Me.Tb_Jmpkm.Text = rempl_virgule(Format(ebpompe.resultat.jmpkm, "####0.00"))
End Sub
Public Sub reini_resu2()

'                    Me.Tb_VitRflt.Text = "0.00"
'                    Me.Tb_Jmpkm.Text = "0.00"
'                    Me.Tb_Drflt.Text = "0"
'                    ebpompe.resultat.VitRflt = 0#
'                    ebpompe.resultat.jmpkm = 0#
'                    ebpompe.resultat.Drflr = 0
'                    Me.Tb_Vutba.Text = "0.00"
'                    ebpompe.don_techniques.Vutba = 0#
'                    Me.Tb_denivt.Text = "0.00"
'                    ebpompe.don_techniques.Denivt = 0#
'                    Me.Tb_denivr.Text = "0.00"
'                    ebpompe.resultat.Denivr = 0#


    ebpompe.resultat.Vurba = 0#
    ebpompe.resultat.Nrdph = 0#
    ebpompe.resultat.Tvidange = 0#
    ebpompe.resultat.T1cyc = 0#
    ebpompe.resultat.Nbcyc = 0#
    ebpompe.resultat.Vmy = 0#
    ebpompe.resultat.Tsejh = 0#
    ebpompe.resultat.Singul = 0#
    ebpompe.resultat.Hmt = 0#
    Me.Tb_vurba.Text = rempl_virgule(Format(ebpompe.resultat.Vurba, "####0.00"))
    Me.Tb_nrdph.Text = rempl_virgule(Format(ebpompe.resultat.Nrdph, "####0.00"))
    Me.Tb_Tvidange.Text = rempl_virgule(Format(ebpompe.resultat.Tvidange, "####0.00"))
    Me.Tb_T1cyc.Text = rempl_virgule(Format(ebpompe.resultat.T1cyc, "####0.00"))
    Me.Tb_Nbcyc.Text = rempl_virgule(Format(ebpompe.resultat.Nbcyc, "####0.00"))
    Me.Tb_Vmy.Text = rempl_virgule(Format(ebpompe.resultat.Vmy, "####0.00"))
    Me.Tb_Tsejh.Text = rempl_virgule(Format(ebpompe.resultat.Tsejh, "####0.00"))
    Me.Tb_Singul.Text = rempl_virgule(Format(ebpompe.resultat.Singul, "###0.000"))
    Me.Tb_Hmt.Text = rempl_virgule(Format(ebpompe.resultat.Hmt, "###0.000"))
End Sub

Private Sub ecrire_tb(ByRef tb As TextBox, ByVal sval As String)
Dim oKp As Boolean
oKp = bKP
bKP = False
tb.Text = sval
bKP = oKp
End Sub
Public Sub ini_ebpompe()
    ebpompe.debits_car.qeum = 0#
    ebpompe.debits_car.Fp = 0#
    ebpompe.debits_car.Qeu = 0#
    ebpompe.debits_car.Qecp = 0#
    ebpompe.debits_car.Qtsm = 0#
    ebpompe.debits_car.Qts = 0#
    ebpompe.debits_car.Qpomp = 0#
    ebpompe.don_geometrie.Lrflt = 0#
    ebpompe.don_geometrie.NatRflt = ""
    ebpompe.don_geometrie.Drflt = 0
    ebpompe.don_geometrie.NivTN = 0#
    ebpompe.don_geometrie.NivEN = 0#
    ebpompe.don_geometrie.NivSO = 0#
    ebpompe.don_geometrie.NivEX = 0#
    ebpompe.pts_singuliers.Nbc1 = 0
    ebpompe.pts_singuliers.Nbc2 = 0
    ebpompe.pts_singuliers.Nbc3 = 0
    ebpompe.pts_singuliers.Nbc4 = 0
    ebpompe.pts_singuliers.Nbc9 = 0
    ebpompe.pts_singuliers.Nbva = 0
    ebpompe.pts_singuliers.Nbcl = 0
    ebpompe.pts_singuliers.Nbvi = 0
    ebpompe.pts_singuliers.Nbve = 0
    ebpompe.pts_singuliers.Antb = 0
    ebpompe.don_techniques.Nbpom = 0
    ebpompe.don_techniques.Ntdph = 0
    ebpompe.don_techniques.Vutba = 0#
    ebpompe.don_techniques.Diamb = 0#
    ebpompe.don_techniques.Largb = 0#
    ebpompe.don_techniques.Longb = 0#
    ebpompe.don_techniques.Denivt = 0#
    ebpompe.don_techniques.Denivhau = 0.2
    ebpompe.don_techniques.Denivbas = 0.8
    ebpompe.don_techniques.Sectb = 0
    ebpompe.resultat.Qpomr = 0#
    ebpompe.resultat.Drflr = 0
    ebpompe.resultat.NatRflr = ""
    ebpompe.resultat.VitRflt = 0#
    ebpompe.resultat.jmpkm = 0#
    ebpompe.resultat.Denivr = 0#
    ebpompe.resultat.Vurba = 0#
    ebpompe.resultat.Nrdph = 0#
    ebpompe.resultat.Tvidange = 0#
    ebpompe.resultat.T1cyc = 0#
    ebpompe.resultat.Nbcyc = 0#
    ebpompe.resultat.Vmy = 0#
    ebpompe.resultat.Tsejh = 0#
    ebpompe.resultat.Singul = 0#
    ebpompe.resultat.Hmt = 0#

End Sub
Public Sub reini_resu_ebpompe()
'    ebpompe.resultat.Qpomr = 0#
'    ebpompe.resultat.NatRflr = ""
 '   ebpompe.resultat.Drflr = 0
    ebpompe.resultat.VitRflt = 0#
    ebpompe.resultat.jmpkm = 0#
    ebpompe.don_techniques.Vutba = 0#
    ebpompe.don_techniques.Denivt = 0#
    ebpompe.resultat.Vurba = 0#
    ebpompe.resultat.Nrdph = 0#
    ebpompe.resultat.Tvidange = 0#
    ebpompe.resultat.T1cyc = 0#
    ebpompe.resultat.Nbcyc = 0#
    ebpompe.resultat.Vmy = 0#
    ebpompe.resultat.Tsejh = 0#
    ebpompe.resultat.Singul = 0#
    ebpompe.resultat.Hmt = 0#
'    If ok_saisie_denivr = False Then
'        ebpompe.resultat.Denivr = 0#
'    End If
End Sub
Public Sub maj_resu()
 '   Me.Tb_Qpomp(1).Text = rempl_virgule(Format(ebpompe.resultat.Qpomr, "###0.00"))
'    Tb_Qpompc(1).Text = 3.6 * ebpompe.resultat.Qpomr
    Me.Tb_Drflt.Text = rempl_virgule(Format(ebpompe.resultat.Drflr, "####0"))
    Lb_materiau.Caption = ebpompe.resultat.NatRflr
    Me.Tb_VitRflt.Text = rempl_virgule(Format(ebpompe.resultat.VitRflt, "###0.00"))
    Me.Tb_Jmpkm.Text = rempl_virgule(Format(ebpompe.resultat.jmpkm, "####0.00"))
    Me.Tb_vurba.Text = rempl_virgule(Format(ebpompe.resultat.Vurba, "####0.00"))
    Me.Tb_nrdph.Text = rempl_virgule(Format(ebpompe.resultat.Nrdph, "####0.00"))
    Me.Tb_Tvidange.Text = rempl_virgule(Format(ebpompe.resultat.Tvidange, "####0.00"))
    Me.Tb_T1cyc.Text = rempl_virgule(Format(ebpompe.resultat.T1cyc, "####0.00"))
    Me.Tb_Nbcyc.Text = rempl_virgule(Format(ebpompe.resultat.Nbcyc, "####0.00"))
    Me.Tb_Vmy.Text = rempl_virgule(Format(ebpompe.resultat.Vmy, "####0.00"))
    Me.Tb_Tsejh.Text = rempl_virgule(Format(ebpompe.resultat.Tsejh, "####0.00"))
    Me.Tb_Singul.Text = rempl_virgule(Format(ebpompe.resultat.Singul, "###0.000"))
    Me.Tb_Hmt.Text = rempl_virgule(Format(ebpompe.resultat.Hmt, "###0.000"))
    Me.Tb_Vutba.Text = rempl_virgule(Format(ebpompe.don_techniques.Vutba, "###0.00"))
    Me.Tb_denivt.Text = rempl_virgule(Format(ebpompe.don_techniques.Denivt, "###0.00"))
    If ok_saisie_denivr = False Then
        Me.Tb_denivr.Text = rempl_virgule(Format(ebpompe.resultat.Denivr, "####0.00"))
    End If
End Sub


Private Sub Opt_PtSing_Click(Index As Integer)
    If Me.Opt_PtSing(0) Then
        ebpompe.pts_singuliers.Antb = 1
    Else
        ebpompe.pts_singuliers.Antb = 0
    End If
Dim mes As String
Dim nom As String
nom = "Opt_PtSing"
mes = Rec_Mes(nom, Index)
'Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
'
    If bKP Then
        Call calcul_resu
    End If
''Call change_ch_sing

End Sub


Private Sub Opt_PtSing_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim mes As String
Dim nom As String
nom = "Opt_PtSing"
If Me.Opt_PtSing(Index) Then
mes = Rec_Mes(nom, Index)
'Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
Else
bKP = True
Opt_PtSing.Item(Index) = True

End If

End Sub

Private Sub Opt_sect_ba_Click(Index As Integer)
    If Me.Opt_sect_ba(0) Then
        Lb_int_diam.Visible = True
        Tb_diam.Visible = True
        Lb_unit_diam.Visible = True
        Lb_int_long.Visible = False
        Tb_long.Visible = False
        Lb_unit_long.Visible = False
        Lb_int_larg.Visible = False
        Tb_larg.Visible = False
        Lb_unit_larg.Visible = False
        ebpompe.don_techniques.Sectb = 0
'        Call Tb_diam_Change
    Else
        Lb_int_diam.Visible = False
        Tb_diam.Visible = False
        Lb_unit_diam.Visible = False
        Lb_int_long.Visible = True
        Tb_long.Visible = True
        Lb_unit_long.Visible = True
        Lb_int_larg.Visible = True
        Tb_larg.Visible = True
        Lb_unit_larg.Visible = True
        ebpompe.don_techniques.Sectb = 1
'         Call Tb_long_Change
   End If
Dim mes As String
Dim nom As String
nom = "Opt_sect_ba"
mes = Rec_Mes(nom, Index)
'Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
   If bKP Then
        Call calcul_resu
    End If
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
    Case Is = 0
        mes = IDhlp_PompeDebitsCaracteristiques
    Case Is = 1
        mes = IDhlp_PompeDonneesGeometriques
    Case Is = 2
        mes = IDhlp_PompePointsSinguliers
    Case Is = 3
        mes = IDhlp_PompeDonneesTechniques
End Select

If owner.fcom.Name = "Frm_ss_commentaire" Then
    Change_Couleur "SSTab1", 0
    DoEvents
    owner.affich_aide Me.Name, mes
    DoEvents
End If
End Sub

Private Sub Tb_debit_Change(Index As Integer)
Dim nom As String
nom = "a"
    Call reini_valeurs

If bKP Then
        Select Case Index
            Case Is = 0
                nom = verif_cart0(Tb_debit(Index).Text, "Saisie du débit moyen des eaux usées", "R")
'            Case Is = 1
'               nom = verif_cart0(Tb_Debit(Index).Text, "Saisie pente conduite Amont", "I")
            Case Is = 2
               nom = verif_cart0(Tb_debit(Index).Text, "Saisie du débit des eaux parasites", "R")
'            Case Is = 3
'               nom = verif_cart0(Tb_Debit(Index).Text, "Saisie cote radier Aval", "R")
        End Select
  If nom = "" Then
    Tb_debit(Index).Text = sval_champ
    Tb_debit(Index).SelStart = iSels
    Tb_debit(Index).SelLength = iSell
  End If
End If
'    Tb_Debitc(Index).Text = val(Tb_debit(Index)) * 3.6
    Tb_Debitc(Index).Text = rempl_virgule(Format(val(Tb_debit(Index)) * 3.6, "###0.00"))
If bKP Then
    Select Case Index
            Case Is = 0
                ebpompe.debits_car.qeum = txtVersNum(Me.Tb_debit(0).Text)
                If ebpompe.debits_car.qeum > 0 Then
                    ebpompe.debits_car.Fp = (1.5 + 2.5 / (ebpompe.debits_car.qeum) ^ 0.5)
                Else
                    ebpompe.debits_car.Fp = 0#
                End If
    '        Case Is = 1
    '            ebpompe.debits_car.Qeu = txtVersNum(Me.Tb_Debit(1).Text)
            Case Is = 2
                 ebpompe.debits_car.Qecp = txtVersNum(Me.Tb_debit(2).Text)
    End Select
    If Index = 0 Or Index = 2 Then
        Call maj_debit
        If ebpompe.resultat.Qpomr < ebpompe.debits_car.Qts Then
            Lb_int_Qpomp(1).Caption = " < Qts  "
            Lb_int_Qpomp(1).ForeColor = 255
            Tb_Qpomp(1).ForeColor = 255
'            Me.Tb_Qpomp(1).SetFocus
        Else
            Lb_int_Qpomp(1).Caption = "Débit de pompage"
            Lb_int_Qpomp(1).ForeColor = &H80000012
            Tb_Qpomp(1).ForeColor = &H80000012
        End If

'                ebpompe.debits_car.Qeu = ebpompe.debits_car.qeum * ebpompe.debits_car.Fp
'                ebpompe.debits_car.Qtsm = ebpompe.debits_car.qeum + ebpompe.debits_car.Qecp
'                ebpompe.debits_car.Qts = ebpompe.debits_car.Qeu + ebpompe.debits_car.Qecp
'                ebpompe.debits_car.Qpomp = Maxi(3 * ebpompe.debits_car.qeum + ebpompe.debits_car.Qecp, ebpompe.debits_car.Qts)
'                bKP = False
'                Tb_FPointe.Text = Format(ebpompe.debits_car.Fp, "###0.00")
'                Tb_debit(1).Text = Format(ebpompe.debits_car.Qeu, "###0.00")
'                Tb_debit(3).Text = Format(ebpompe.debits_car.Qtsm, "##0.00")
'                Tb_debit(4).Text = Format(ebpompe.debits_car.Qts, "##0.00")
'                Tb_Qpomp(0).Text = Format(ebpompe.debits_car.Qpomp, "##0.00")
'                Tb_Qpompc(0).Text = Format(3.6 * ebpompe.debits_car.Qpomp, "##0.00")
'                bKP = True
    End If
'        If Index = 0 And val(Tb_Debit(0).Text) > 0 Then
        If Index = 0 Then
            Call ecrire_tb(Tb_FPointe, Format(ebpompe.debits_car.Fp, "##0.00"))
'            bKP = False
'               Tb_FPointe.Text = Format(ebpompe.debits_car.Fp, "##0.00")
'            bKP = True
        End If
        
''''Debit de pointe EU (index=1)= Produit du débit moyen EU (index= 0)par Facteur de Pointe (index=1)
''''Qeu
''''Debug.Print Tb_debit(0).Text
'''    Tb_Debit(1).Text = Format(val(Tb_Debit(0).Text) * val(Tb_FPointe.Text), "###0.00")
''''Debit moyen de temps sec (index=3)= Somme du débit moyen EU (index= 0)et du débit des eaux parasites (index= 2)
''''Qmts
'''    Tb_Debit(3).Text = Format(val(Tb_Debit(0).Text) + val(Tb_Debit(2).Text), "##0.00")
''''Debit de pointe de temps sec (index=4)= Somme du débit Pointe EU (index= 2)et du débit des eaux parasites (index= 3)
''''Qts
'''    Tb_Debit(4).Text = Format(val(Tb_Debit(1).Text) + val(Tb_Debit(2).Text), "##0.00")
'''
''''Débit de pompage = 3 fois le débit moyen de temps sec(index=3)
''''Qpomp
''' '   Tb_Qpomp(0).Text = Format(3 * val(Tb_Debit(3).Text), "##0.00")
'''    Tb_Qpomp(0).Text = Format(Maxi(3 * val(Tb_Debit(0).Text) + val(Tb_Debit(2).Text), val(Tb_Debit(4).Text)), "##0.00")

''''
'''Select Case Index
'''        Case Is = 0
'''            ebpompe.debits_car.qeum = txtVersNum(Me.Tb_Debit(0).Text)
'''            ebpompe.debits_car.Fp = txtVersNum(Me.Tb_FPointe.Text)
'''        Case Is = 1
'''            ebpompe.debits_car.Qeu = txtVersNum(Me.Tb_Debit(1).Text)
'''        Case Is = 2
'''             ebpompe.debits_car.Qecp = txtVersNum(Me.Tb_Debit(2).Text)
'''       Case Is = 3
'''            ebpompe.debits_car.Qtsm = txtVersNum(Me.Tb_Debit(3).Text)
'''            ebpompe.debits_car.Qpomp = txtVersNum(Me.Tb_Qpomp(0).Text)
'''       Case Is = 4
'''            ebpompe.debits_car.Qts = txtVersNum(Me.Tb_Debit(4).Text)
'''           ebpompe.debits_car.Qpomp = txtVersNum(Me.Tb_Qpomp(0).Text)
'''End Select
End If

If ebpompe.debits_car.Qpomp >= ebpompe.debits_car.Qts Then
    If bKP Then
        Call calcul_resu
    End If
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(3) = True
    Tb_Qpomp(1).Enabled = True
    Tb_Drflt.Enabled = True
Else
    SSTab1.TabEnabled(1) = False
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(3) = False
    Tb_Qpomp(1).Enabled = False
    Tb_Drflt.Enabled = False
End If

'End If
    sval_champ = ""
    bKP = False
End Sub

Private Sub Tb_debit_GotFocus(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_debit"
If SSTab1.Tab <> 0 Then
'    SSTab1.Tab = 0
End If
Call sel_text(Tb_debit(Index))
'If change_coul Then
'    Change_Couleur nom, Index
'    mes = Rec_Mes(nom, Index)
'    owner.affich_aide Me.Name, mes
'End If
''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub


Private Sub Tb_debit_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_debit(Index).Text
    iSels = Tb_debit(Index).SelStart
    iSell = Tb_debit(Index).SelLength
End If
End Sub
Private Sub Tb_debit_Click(Index As Integer)

Dim mes As String
Dim nom As String
nom = "Tb_debit"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
Call meAffiche
Call sel_text(Tb_debit(Index))
'''owner.affich_aide Me.Name, "pompe Conduite Amont"
End Sub

Private Sub ini_lbresu()
Dim sresult As String
'
            sresult = "   Durée du cycle "
            sresult = sresult + Chr(13) + "   Nombre de cycle(s) "
            sresult = sresult + Chr(13) + "   Vitesse moyenne d'écoulement sur 24h "
            sresult = sresult + Chr(13) + "   Temps de séjour "
    Me.Lb_resu.BorderStyle = 1
    Me.Lb_resu.Caption = sresult
'  Me.Lb_amo.BackColor = &H8000000B
    Me.Lb_amo.BorderStyle = 1
    Me.Lb_amo.Caption = ""
'    Me.Lb_ava.BackColor = &H8000000B
    Me.Lb_ava.BorderStyle = 1
    Me.Lb_ava.Caption = ""
'    Me.lb_pompe.BackColor = &H8000000B
    Me.Lb_pompe.BorderStyle = 1
    Me.Lb_pompe.Caption = ""
End Sub
Private Sub modi_res_cana()
'    Me.Lb_amo.BackColor = &H80000009
    Me.Lb_amo.BorderStyle = 1
'    Me.Lb_ava.BackColor = &H80000009
    Me.Lb_ava.BorderStyle = 1
End Sub
Private Sub modi_res_pompe()
'    Me.lb_pompe.BackColor = &H80000009
    Me.Lb_pompe.BorderStyle = 1
End Sub
Private Sub calcul_amont_aval()
'Dim z1 As Double, z2 As Double, h1 As Double, h2 As Double, h0 As Double
'Dim v1 As Double, x0 As Double, X As Double, g As Double
'Dim sresult As String, sresult1 As String, sresult2 As String
'Dim troamo As troncon, troava As troncon
'Dim cana_amo As conduite
'Dim res_amo As debit_conduit
'Dim res_ava As debit_conduit
'Dim cana_ava As conduite
'Dim qv As deb_vit, qvps_amo As deb_vit, qvps_ava As deb_vit
'g = 9.81
'' conduite amont -> troncon amont
'    cana_amo.Diametre = ebpompe.dam / 1000#
'    cana_amo.Longueur = 5
'    cana_amo.pente = ebpompe.iRadam / 10000#
'    cana_amo.rugosite = ebpompe.Kam
'    cana_amo.typ = 2
'    With troamo
'      .Absamo = 0#
'      .Absava = .Absamo + cana_amo.Longueur
'      .conduit = cana_amo
'      .radava = ebpompe.Rdav
'      .radamo = ebpompe.Rdav + cana_amo.Longueur * cana_amo.pente '0.3 '
'    End With
'    ebpompe.tron_amo = troamo
'    qvps_amo = debvit_ps(ebpompe.tron_amo.conduit)
'    res_amo = calc_debit_tr(ebpompe.tron_amo, ebpompe.Qmax)
''    Debug.Print res_amo.charge, res_amo.debit, res_amo.hauteur, res_amo.vitesse
'    cana_ava.Diametre = ebpompe.dav / 1000#
'    cana_ava.Longueur = 5
'    cana_ava.pente = ebpompe.iradav / 10000#
'    cana_ava.rugosite = ebpompe.kav
'    cana_ava.typ = 2
'    With troava
'      .Absava = 0#
'      .Absava = .Absava + cana_ava.Longueur
'      .conduit = cana_ava
'      .radamo = ebpompe.Rdam
'      .radava = ebpompe.Rdam - cana_ava.Longueur * cana_ava.pente ' 0.3 '
'    End With
'    ebpompe.tron_ava = troava
'    qvps_ava = debvit_ps(ebpompe.tron_ava.conduit)
'    res_ava = calc_debit_tr(ebpompe.tron_ava, ebpompe.Qmax)
''    Debug.Print res_ava.charge, res_ava.debit, res_ava.hauteur, res_ava.vitesse
'    h1 = res_amo.hauteur
'    h2 = res_ava.hauteur
'    v1 = res_amo.vitesse
'    z1 = ebpompe.tron_amo.radava + h1
'    z2 = ebpompe.tron_ava.radamo + h2
''    Me.Lb_debitam.Caption = "Débit PS " + Trim(Str(Round(qvps_amo.debit, 3))) + " m3/s"
''    Me.Lb_debitav.Caption = "Débit PS " + Trim(Str(Round(qvps_ava.debit, 3))) + " m3/s"
''    Me.Lb_Vitam.Caption = "Vitesse PS " + Trim(Str(Round(qvps_amo.vitesse, 2))) + " m/s"
''    Me.Lb_Vitav.Caption = "Vitesse PS " + Trim(Str(Round(qvps_ava.vitesse, 2))) + " m/s"
'    Call modi_res_cana
'    sresult = "  Débit pleine section = " + ajout_zero(Trim(Str(Round(qvps_amo.debit, 3)))) + " m3/s"
'    sresult1 = "  Débit pleine section = " + ajout_zero(Trim(Str(Round(qvps_ava.debit, 3)))) + " m3/s"
'    sresult = sresult + Chr(13) + "   Vitesse pleine section = " + ajout_zero(Trim(Str(Round(qvps_amo.vitesse, 2)))) + " m/s"
'    sresult1 = sresult1 + Chr(13) + "   Vitesse pleine section = " + ajout_zero(Trim(Str(Round(qvps_ava.vitesse, 2)))) + " m/s"
'
'    Call init_graph(owner.fdessin.UC_graphique1)
'    Call init_graph(Frm_desprint.UC_graphique1)
'
'    If res_amo.charge Then
''        Me.Lb_Hautam.Caption = "Conduite en charge"
'       sresult = sresult + Chr(13) + Chr(13) + "   Conduite en charge"
'    Else
''        Me.Lb_Hautam.Caption = " Hauteur    " + Trim(Str(Round(res_amo.hauteur, 2))) + " m"
'        sresult = sresult + Chr(13) + Chr(13) + "   Hauteur  = " + ajout_zero(Trim(Str(Round(res_amo.hauteur, 2)))) + " m"
'        sresult = sresult + Chr(13) + "   Vitesse = " + ajout_zero(Trim(Str(Round(res_amo.vitesse, 2)))) + " m/s"
'        If res_ava.charge Then
''            Me.Lb_Hautav.Caption = "Conduite en charge"
'           sresult1 = sresult1 + Chr(13) + Chr(13) + "   Conduite en charge"
'        Else
'           Call modi_res_pompe
''            Me.Lb_Hautav.Caption = "Hauteur    " + Trim(Str(Round(res_ava.hauteur, 2))) + " m"
'           sresult1 = sresult1 + Chr(13) + Chr(13) + "   Hauteur = " + ajout_zero(Trim(Str(Round(res_ava.hauteur, 2)))) + " m"
'           sresult1 = sresult1 + Chr(13) + "   Vitesse = " + ajout_zero(Trim(Str(Round(res_ava.vitesse, 2)))) + " m/s"
'            h0 = z1 - z2
''            Me.Lb_Haut.Caption = "Dénivelée du liquide   " + Trim(Str(Round(h0, 2))) + " m"
''           sresult2 = "  Dénivelée = " + ajout_zero(Trim(Str(Round(h0, 2)))) + " m"
'            ebpompe.h0 = Round(h0, 2)
'            Dim hc As Double, he As Double, le As Double
'            he = 0: le = 0
'            hc = long_pompe(ebpompe, res_amo, res_ava, he, le)
'            h0 = ebpompe.h0
''            sresult2 = "  Réduction de la hauteur d'eau = 0"
''            sresult2 = "  Hauteur initiale  = " + ajout_zero(Trim(Str(Round(he, 2)))) + " m "
''            sresult2 = sresult2 + Chr(13) + "   Longueur de variation = " + ajout_zero(Trim(Str(Round(le, 2)))) + " m "
''            sresult2 = sresult2 + Chr(13) + Chr(13) + "   Dénivelée = " + ajout_zero(Trim(Str(Round(h0, 2)))) + " m"
'            ebpompe.h0 = Round(h0, 2)
'            x0 = ((h0 * v1 ^ 2) / g) ^ 0.5
'            X = x0 * 2#
'            X = hc
''            Me.Lb_Long.Caption = "Longueur du dispositif   " + Trim(Str(Round(X, 2))) + " m"
'           sresult2 = "  Longueur du dispositif = " + ajout_zero(Trim(Str(Round(X, 2)))) + " m"
'            sresult2 = sresult2 + Chr(13) + "   Dénivelée = " + ajout_zero(Trim(Str(Round(h0, 2)))) + " m"
'            sresult2 = sresult2 + Chr(13) + Chr(13) + "   Hauteur initiale = " + ajout_zero(Trim(Str(Round(he, 2)))) + " m "
'            sresult2 = sresult2 + Chr(13) + "   Longueur de variation = " + ajout_zero(Trim(Str(Round(le, 2)))) + " m "
'            ebpompe.Long = Round(X, 2)
'            Me.Cmd_calcul.Enabled = False
'            'impression true
''            Me.mnuprint.Enabled = True
'            Me.Lb_pompe.Caption = sresult2
'            Call dess_pompe(owner.fdessin.UC_graphique1, h1, h2, v1)
'            Call dess_pompe1(owner.fdessin.UC_graphique1, ebpompe, res_amo, res_ava)
'
'            Call dess_pompe(Frm_desprint.UC_graphique1, h1, h2, v1)
'            Call dess_pompe1(Frm_desprint.UC_graphique1, ebpompe, res_amo, res_ava)
'        End If
'   End If
'    Me.Lb_amo.Caption = sresult
'    Me.Lb_ava.Caption = sresult1
End Sub
Public Sub dess_pompe(ByRef uc_g As UC_graphique)
Dim xam As Double, yam As Double, xav As Double, yav As Double
Dim HT As Double, henb As Double, henh As Double, hsob As Double, hsoh As Double
Dim xdpomp As Double, xfpomp As Double, hexb As Double, hexh As Double
Dim hgarde As Double, hdeniv As Double, hfond As Double, diam As Double
Dim plin As Double, psing As Double, hmtot As Double, hpertes As Double
Dim diam_cond As Double, decal As Double
Dim cote_ex As Double, okpertes As Boolean
Dim okhaut As Boolean, nb_diam As Integer, text_pertes As String, text_hmt As String
Dim yhmt As Double, yhhmt As Double, ybhmt As Double
Dim yam1 As Double, yav1 As Double
Dim yam2 As Double, yav2 As Double
Dim hmax As Double
okhaut = True
okpertes = True
nb_diam = 2
hmax = uc_g.lire_MaxYn
    hgarde = ebpompe.don_techniques.Denivhau
    hdeniv = ebpompe.resultat.Denivr
    hfond = ebpompe.don_techniques.Denivbas
    If ebpompe.don_techniques.Sectb = 0 Then
        diam = ebpompe.don_techniques.Diamb
    Else
        diam = ebpompe.don_techniques.Largb
    End If
    plin = ebpompe.resultat.jmpkm * ebpompe.don_geometrie.Lrflt / 1000#
    psing = ebpompe.resultat.Singul
    hpertes = plin + psing
    If hpertes > 3# Then
        okpertes = False
        hpertes = 3#
    End If

    hmtot = ebpompe.resultat.Hmt
    cote_ex = ebpompe.don_geometrie.NivEX
    ok = False
    If ebpompe.don_geometrie.NivEX - ebpompe.don_geometrie.NivTN > 3 Then
        okhaut = False
        nb_diam = 5
        cote_ex = ebpompe.don_geometrie.NivTN + 3#
    End If

diam_cond = 0.4
decal = 1#
xdpomp = 3 * diam
xfpomp = 4 * diam
HT = (ebpompe.don_geometrie.NivTN - ebpompe.don_geometrie.NivEN) + hgarde + hdeniv + hfond
henb = hgarde + hdeniv + hfond
henh = henb + diam_cond
hsob = henb + (ebpompe.don_geometrie.NivSO - ebpompe.don_geometrie.NivEN)
hsoh = hsob + diam_cond
hexb = HT + (cote_ex - ebpompe.don_geometrie.NivTN) 'HT + 1.5
hexh = hexb + diam_cond
uc_g.redef_drwidth 8
'dessin des contours de la station----------------------------
xam = xdpomp
yam = HT + decal
xav = xdpomp
yav = decal
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 8
xam = xdpomp
yam = decal
xav = xfpomp - diam / 2.5
yav = decal
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 8
xam = xfpomp - diam / 10#
yam = decal
xav = xfpomp
yav = decal
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 8
xam = xfpomp
yam = decal
xav = xfpomp
yav = henb + decal
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 8
xam = xfpomp
yam = henh + decal
xav = xfpomp
yav = HT + decal
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 8
xam = xdpomp
yam = HT + decal
xav = xdpomp + diam / 3#
yav = HT + decal
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 8
xam = xfpomp - diam / 3#
yam = HT + decal
xav = xfpomp
yav = HT + decal
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 8
'-------compléments en traits fin haut------------
uc_g.redef_drwidth 2
xam = xdpomp + diam / 3#
yam = HT + decal
xav = xfpomp - diam / 3#
yav = HT + decal
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
'-------compléments en traits fin bas------------
xam = xfpomp - diam / 2.5
yam = decal
xav = xfpomp - diam / 10#
yav = decal
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xfpomp - diam / 2.5
yam = decal - 0.3
xav = xfpomp - diam / 10#
yav = decal - 0.3
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xfpomp - diam / 2.5
yam = decal - 0.3
xav = xfpomp - diam / 2.5
yav = decal
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xfpomp - diam / 10#
yam = decal - 0.3
xav = xfpomp - diam / 10#
yav = decal
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
'dessin conduite arrivee--------------------------------------
xam = xfpomp - diam / 10#
yam = henb + decal
xav = xfpomp + diam / 2.5
yav = henb + decal
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xfpomp - diam / 10#
yam = henh + decal
xav = xfpomp + diam / 3.5
yav = henh + decal
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xfpomp - diam / 10#
yam = henb + decal
xav = xfpomp - diam / 10#
yav = henh + decal
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xfpomp + diam / 2.5
yam = henb + decal
xav = xfpomp + diam / 3.5
yav = henh + decal
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
'dessin conduite sortie------------------------------------
uc_g.redef_drwidth 2
xam = xdpomp - diam * 0.53 '0.6---------------------
yam = hsob + decal
xav = xdpomp + diam / 10#
yav = hsob + decal
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xdpomp - diam * 0.5
yam = hsoh + decal
xav = xdpomp + diam / 10#
yav = hsoh + decal
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xdpomp + diam / 10#
yam = hsob + decal
xav = xdpomp + diam / 10#
yav = hsoh + decal
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
'xam = xdpomp - diam * 0.6
'yam = hsob + decal
'xav = xdpomp - diam * 0.5
'yav = hsoh + decal
'uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
'dessin conduite refoulement----------------------------------
uc_g.redef_drwidth 2
xam = xdpomp - diam * 2.2
yam = hexb + decal
xav = xdpomp - diam * 2.03
yav = hexb + decal
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xdpomp - diam * 2.2
yam = hexh + decal
xav = xdpomp - diam * 2#
yav = hexh + decal
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xdpomp - diam * 2.2
yam = hexb + decal
xav = xdpomp - diam * 2.2
yav = hexh + decal
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
'raccord conduite refoulement avec conduite sortie-------
'''raccord haut biais----------------------------------
    uc_g.redef_drwidth 1
If okhaut Then
'si hauteur juste------------------------------
    xam = xdpomp - diam * 2#
    yam = hexh + decal
    xav = xdpomp - diam * 1.6
    yav = hexh + decal - nb_diam * diam_cond
    uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
    xam = xdpomp - diam * 2.03
    yam = hexb + decal
    xav = xdpomp - diam * 1.62
    yav = hexh + decal - (nb_diam + 1) * diam_cond
    uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
Else
'si hauteur limitée----------------------------
'-----haut----------------------------------------------
    xam = xdpomp - diam * 2#
    yam = hexh + decal
    xav = xdpomp - diam * 1.81   '1.82
    yam1 = hexh + decal - nb_diam * diam_cond
    yav1 = (yam + yam1) / 2
    yav = yav1 '- diam_cond * 0.05 '0.2 '+ diam_cond * 0.1  '0.2
    uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
    xam = xdpomp - diam * 1.82
    yam = yav
    xav = xdpomp - diam * 1.77  '1.78
    yav = yav1 - diam_cond * 0.4 '0.5  '0.2
    uc_g.redef_drwidth 1
'    uc_g.dess_lign_point xam, yam, xav, yav, couleur.noir
    uc_g.redef_drwidth 2
    xam = xdpomp - diam * 1.77 '1.78
    yam = yav
    xav = xdpomp - diam * 1.6
    yav = hexh + decal - nb_diam * diam_cond
    uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
'----bas---------------------------------------------------
    xam = xdpomp - diam * 2.03
    yam = hexb + decal
    xav = xdpomp - diam * 1.85
    yam2 = hexh + decal - (nb_diam + 1) * diam_cond
    yav2 = (yam + yam2) / 2
    yav = yav2 + diam_cond * 0.2
    uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
    xam = xdpomp - diam * 1.85
    yam = yav
    xav = xdpomp - diam * 1.81
    yav = yav2 - diam_cond * 0.3  '0.2
    uc_g.redef_drwidth 1
'    uc_g.dess_lign_point xam, yam, xav, yav, couleur.noir
    uc_g.redef_drwidth 2
    xam = xdpomp - diam * 1.81
    yam = yav
    xav = xdpomp - diam * 1.62
    yav = hexh + decal - (nb_diam + 1) * diam_cond
    uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
'------lignes biais   de découpe------------------------
    uc_g.redef_drwidth 1
    xam = xdpomp - diam * 1.76   '1.82
    yam = yav1 + diam_cond * 0.6
    xav = xdpomp - diam * 1.91 '1.85
    yav = yav2 - diam_cond * 0.5
    uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 1
    xam = xdpomp - diam * 1.72  '1.78
    yam = yav1 + diam_cond * 0.4  '0.2
    xav = xdpomp - diam * 1.87  '1.81
    yav = yav2 - diam_cond * 0.7  '0.2
    uc_g.redef_drwidth 1
    uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 1
    uc_g.redef_drwidth 2



End If
'''raccord bas biais--------------------------------------
    xam = xdpomp - diam * 1.3
    yam = hexh + decal - nb_diam * diam_cond
    xav = xdpomp - diam * 0.5
    yav = hsoh + decal
    uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
    xam = xdpomp - diam * 1.32
    yam = hexh + decal - (nb_diam + 1) * diam_cond
    xav = xdpomp - diam * 0.53
    yav = hsob + decal
    uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
''''raccord haut horizontal----------------------------------
    xam = xdpomp - diam * 1.58  '1.6
    yam = hexh + decal - nb_diam * diam_cond
    xav = xdpomp - diam * 1.5
    yav = hexh + decal - nb_diam * diam_cond
    uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
    xam = xdpomp - diam * 1.5
    yam = hexh + decal - nb_diam * diam_cond
    xav = xdpomp - diam * 1.4
    yav = hexh + decal - nb_diam * diam_cond
    uc_g.redef_drwidth 1
   uc_g.dess_lign_point xam, yam, xav, yav, couleur.noir
    uc_g.redef_drwidth 2
    xam = xdpomp - diam * 1.4
    yam = hexh + decal - nb_diam * diam_cond
    xav = xdpomp - diam * 1.3
    yav = hexh + decal - nb_diam * diam_cond
    uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
'''raccord bas horizontal-----------------------------------
    xam = xdpomp - diam * 1.62
    yam = hexh + decal - (nb_diam + 1) * diam_cond
    xav = xdpomp - diam * 1.52
    yav = hexh + decal - (nb_diam + 1) * diam_cond
    uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
    xam = xdpomp - diam * 1.52
    yam = hexh + decal - (nb_diam + 1) * diam_cond
    xav = xdpomp - diam * 1.42
    yav = hexh + decal - (nb_diam + 1) * diam_cond
    uc_g.redef_drwidth 1
    uc_g.dess_lign_point xam, yam, xav, yav, couleur.noir
    uc_g.redef_drwidth 2
    xam = xdpomp - diam * 1.42
    yam = hexh + decal - (nb_diam + 1) * diam_cond
    xav = xdpomp - diam * 1.32
    yav = hexh + decal - (nb_diam + 1) * diam_cond
    uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
'''''''
    uc_g.redef_drwidth 1
    xam = xdpomp - diam * 1.48
    yam = hexh + decal - nb_diam * 0.8 * diam_cond
    xav = xdpomp - diam * 1.5
    yav = hexh + decal - (nb_diam + 1) * 1.2 * diam_cond
    uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 1
    xam = xdpomp - diam * 1.42
    yam = hexh + decal - nb_diam * 0.8 * diam_cond
    xav = xdpomp - diam * 1.45 '1.44
    yav = hexh + decal - (nb_diam + 1) * 1.2 * diam_cond
    uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 1

'End If
''cotation
uc_g.redef_drwidth 1
xam = xdpomp - diam * 2.2
yam = hexb + decal + diam_cond / 2#
xav = xfpomp + diam * 0.4
yav = hexb + decal + diam_cond / 2#
uc_g.dess_lign_point xam, yam, xav, yav, couleur.bleu 'ligne horiz h hgeom
xam = xdpomp - diam * 0.1
yam = hfond + hdeniv / 2# + decal
xav = xfpomp + diam * 0.9
yav = hfond + hdeniv / 2# + decal
ybhmt = yav
uc_g.dess_lign_point xam, yam, xav, yav, couleur.bleu 'ligne horiz b hgeom + Hmt
'xam = xfpomp + diam * 0.35
'yam = hfond + hdeniv / 2# + decal
'xav = xfpomp + diam * 0.35
'yav = hexb + decal + diam_cond / 2#
'uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 1 'ligne verticale  hgeom
xam = xfpomp + diam * 0.2
yam = hexb + decal + diam_cond / 2# + hpertes
xav = xfpomp + diam * 0.9
yav = hexb + decal + diam_cond / 2# + hpertes
uc_g.dess_lign_point xam, yam, xav, yav, couleur.bleu ' ligne horiz h Hmt

'**********************
xam = -3  'xdpomp - diam * 2.5 '2.8
xam = xdpomp - diam * 2.5 '2.8
yam = hmax - 0.4   'hexb + decal + diam_cond / 2# + hpertes + 0.2
text_pertes = "PCL-S = Pertes de charges linéaires (" + Format(plin, "##0.000") + " m)"
text_pertes = text_pertes + " ; singuliéres (" + Format(psing, "##0.000") + " m)"
uc_g.dess_text_aligne xam, "G", "H", yam, text_pertes, couleur.bleu
text_hmt = "HMT = Hauteur manométrique totale (" + Format(ebpompe.resultat.Hmt, "##0.000") + " m)"
yam = hmax - 0.6 'yam = hexb + decal + diam_cond / 2# + hpertes
uc_g.dess_text_aligne xam, "G", "B", yam, text_hmt, couleur.bleu

If okpertes Then
xam = xfpomp + diam * 0.35
yam = hexb + decal + diam_cond / 2#
xav = xfpomp + diam * 0.35
yav = hexb + decal + diam_cond / 2# + hpertes
yhhmt = yav
uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 1 'ligne verticale  pertes
'text_pertes = "Pertes de charges linéaires (" + Format(plin, "##0.000") + " m)"
'text_pertes = text_pertes + " ; singuliéres (" + Format(psing, "##0.000") + " m)"
text_pertes = "PCL-S"
uc_g.dess_text_aligne xam, "D", "C", (yam + yav) / 2, text_pertes, couleur.bleu
Else
xam = xfpomp + diam * 0.35
yam = hexb + decal + diam_cond / 2#
xav = xfpomp + diam * 0.35
yav1 = hexb + decal + diam_cond / 2# + hpertes
yhhmt = yav1
yav2 = (yam + yav1) / 2
yav = yav2 - diam_cond * 0.2
uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 1 'ligne verticale  pertes
xam = xfpomp + diam * 0.35
yam = yav2 + diam_cond * 0.2
xav = xfpomp + diam * 0.35
yav = yav1
uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 1 'ligne verticale  pertes

xam = xfpomp + diam * 0.25
yam = yav2 - diam_cond * 0.2 - 0.2
xav = xfpomp + diam * 0.45
yav = yav2 - diam_cond * 0.2 + 0.2
uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 1 'ligne biais haut  Hmt
xam = xfpomp + diam * 0.25
yam = yav2 + diam_cond * 0.2 - 0.2
xav = xfpomp + diam * 0.45
yav = yav2 + diam_cond * 0.2 + 0.2
uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 1 'ligne biais bas  Hmt


'text_pertes = "Pertes de charges linéaires (" + Format(plin, "##0.000") + " m)"
'text_pertes = text_pertes + " ; singuliéres (" + Format(psing, "##0.000") + " m)"
text_pertes = "PCL-S"
uc_g.dess_text_aligne xam, "D", "C", yav2, text_pertes, couleur.bleu
End If
If okhaut Then
xam = xfpomp + diam * 0.85
yam = hfond + hdeniv / 2# + decal
xav = xfpomp + diam * 0.85
yav = hexb + decal + diam_cond / 2# + hpertes
uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 1 'ligne verticale  Hmt
Else
xam = xfpomp + diam * 0.85
yam = hfond + hdeniv / 2# + decal
xav = xfpomp + diam * 0.85
yav = HT + decal + diam_cond * 2#
uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 1 'ligne verticale  Hmt
xam = xfpomp + diam * 0.85
yam = HT + decal + diam_cond * 2.5
xav = xfpomp + diam * 0.85
yav = hexb + decal + diam_cond / 2# + hpertes
uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 1 'ligne verticale  Hmt
xam = xfpomp + diam * 0.95
yam = HT + decal + diam_cond * 2.5 + 0.2
xav = xfpomp + diam * 0.75
yav = HT + decal + diam_cond * 2.5 - 0.2
uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 1 'ligne biais haut  Hmt
xam = xfpomp + diam * 0.95
yam = HT + decal + diam_cond * 2# + 0.2
xav = xfpomp + diam * 0.75
yav = HT + decal + diam_cond * 2# - 0.2
uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 1 'ligne biais bas  Hmt
End If
If (hexb + diam_cond / 2#) - HT > 2 * diam_cond Then
yam = (hexb + decal + diam_cond / 2#) - 2 * diam_cond
Else
yam = hexb + decal + 2 * diam_cond
End If
xam = xfpomp + diam * 0.85
'text_hmt = "Hauteur manométrique totale (" + Format(ebpompe.resultat.Hmt, "##0.000") + " m)"
text_hmt = "HMT"
uc_g.dess_text_aligne xam, "D", "C", yam, text_hmt, couleur.bleu

xam = xdpomp - diam * 2.5 '2.8
yam = hexb + decal
xav = xdpomp - diam * 2.2  '2.5
yav = hexb + decal
uc_g.dess_lign_point xam, yam, xav, yav, couleur.bleu
uc_g.dess_text_aligne xam, "D", "C", yam, "Niveau extrémité (" + Format(ebpompe.don_geometrie.NivEX, "####0.00") + " m)", couleur.bleu
xam = xdpomp - diam / 3
yam = HT + decal
xav = xfpomp - diam / 10#
yav = HT + decal
uc_g.dess_lign_point xam, yam, xav, yav, couleur.bleu
uc_g.dess_text_aligne xam, "D", "C", yam, Format(ebpompe.don_geometrie.NivTN, "####0.00") + " m", couleur.bleu
xam = xdpomp - diam * 0.6
yam = hsob + decal
xav = xdpomp - diam * 0.9
yav = hsob + decal
uc_g.dess_lign_point xam, yam, xav, yav, couleur.bleu
uc_g.dess_text_aligne xav, "D", "C", yam, "Niveau de sortie (" + Format(ebpompe.don_geometrie.NivSO, "####0.00") + " m)", couleur.bleu
xam = xdpomp - diam / 2#
yam = henb + decal
xav = xfpomp - diam / 10#
yav = henb + decal
uc_g.dess_lign_point xam, yam, xav, yav, couleur.bleu
uc_g.dess_text_aligne xam, "D", "H", yam, "Niveau d'arrivée (" + Format(ebpompe.don_geometrie.NivEN, "####0.00") + " m)", couleur.bleu
xam = xdpomp - diam / 2#
yam = henb - hgarde + decal
xav = xfpomp
yav = henb - hgarde + decal
uc_g.dess_lign_point xam, yam, xav, yav, couleur.bleu
uc_g.dess_text_aligne xam, "D", "B", yam, "Niveau de démarrage (" + Format((ebpompe.don_geometrie.NivEN - hgarde), "####0.00") + " m)", couleur.bleu
xam = xdpomp - diam / 2#
yam = henb - hgarde - hdeniv + decal
xav = xfpomp
yav = henb - hgarde - hdeniv + decal
uc_g.dess_lign_point xam, yam, xav, yav, couleur.bleu
uc_g.dess_text_aligne xam, "D", "C", yam, "Niveau d'arrêt (" + Format((ebpompe.don_geometrie.NivEN - hgarde - hdeniv), "####0.00") + " m)", couleur.bleu
xam = xdpomp - diam / 3
yam = henb - hgarde - hdeniv - hfond + decal
xav = xfpomp - diam / 10#
yav = henb - hgarde - hdeniv - hfond + decal
uc_g.dess_lign_point xam, yam, xav, yav, couleur.bleu
uc_g.dess_text_aligne xam, "D", "C", yam, Format((ebpompe.don_geometrie.NivEN - hgarde - hdeniv - hfond), "####0.00") + " m", couleur.bleu
End Sub
Public Sub Init_ss_commentaire()
    owner.affich_aide Me.Name, ""  'Dimensionnement d'une pompe"
End Sub

Private Sub Tb_debit_LostFocus(Index As Integer)
'If Index = 2 Then
'Me.Tb_Debit(0).SetFocus
'End If
'Tb_debit_LostFocus(Index As Integer)
End Sub
Private Sub Tb_denivbas_Change()
Dim nom As String
nom = "a"
    Call reini_valeurs

If bKP Then
        nom = verif_cart0(Tb_denivbas.Text, "Saisie de la garde au fond", "R")
  If nom = "" Then
    Tb_denivbas.Text = sval_champ
    Tb_denivbas.SelStart = iSels
    Tb_denivbas.SelLength = iSell
  End If
End If
ebpompe.don_techniques.Denivbas = txtVersNum(Me.Tb_denivbas.Text)
    sval_champ = ""
    bKP = False
'Call Me.change_hmt
Call dessin_pompe


End Sub

Private Sub Tb_denivbas_Click()
Dim mes As String
Dim nom As String
nom = "Tb_denivbas"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
owner.affich_aide Me.Name, mes
Call meAffiche
Call sel_text(Tb_denivbas)
'''owner.affich_aide Me.Name, "pompe Conduite Amont"


End Sub

Private Sub Tb_denivbas_GotFocus()
Dim mes As String
Dim nom As String
nom = "Tb_denivbas"
Call sel_text(Tb_denivbas)
'If change_coul Then
'    Change_Couleur nom, 0
'    mes = Rec_Mes(nom, 0)
'    owner.affich_aide Me.Name, mes
'End If
''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub

Private Sub Tb_denivbas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_denivbas.Text
    iSels = Tb_denivbas.SelStart
    iSell = Tb_denivbas.SelLength
End If

End Sub

Private Sub Tb_denivhau_Change()
Dim nom As String
nom = "a"
If bKP Then
        nom = verif_cart0(Tb_denivhau.Text, "Saisie de la garde à l'égout", "R")
  If nom = "" Then
    Tb_denivhau.Text = sval_champ
    Tb_denivhau.SelStart = iSels
    Tb_denivhau.SelLength = iSell
  End If
End If
ebpompe.don_techniques.Denivhau = txtVersNum(Me.Tb_denivhau.Text)
    sval_champ = ""
    bKP = False
Call Me.change_hmt
'Call dessin_pompe

End Sub

Private Sub Tb_denivhau_Click()
Dim mes As String
Dim nom As String
' 20061114
    Call reini_valeurs
nom = "Tb_denivhau"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
owner.affich_aide Me.Name, mes
Call meAffiche
Call sel_text(Tb_denivhau)
'''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub

Private Sub Tb_denivhau_GotFocus()
Dim mes As String
Dim nom As String
nom = "Tb_denivhau"
Call sel_text(Tb_denivhau)
'If change_coul Then
'    Change_Couleur nom, 0
'    mes = Rec_Mes(nom, 0)
'    owner.affich_aide Me.Name, mes
'End If
''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub

Private Sub Tb_denivhau_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_denivhau.Text
    iSels = Tb_denivhau.SelStart
    iSell = Tb_denivhau.SelLength
End If

End Sub

Private Sub Tb_denivr_Change()
Dim nom As String
    Call reini_valeurs
nom = "a"
If bKP Then
        nom = verif_cart0(Tb_denivr.Text, "Saisie du dénivelé retenu", "R")
  If nom = "" Then
    Tb_denivr.Text = sval_champ
    Tb_denivr.SelStart = iSels
    Tb_denivr.SelLength = iSell
  End If
    ebpompe.resultat.Denivr = txtVersNum(Me.Tb_denivr.Text)
    ok_saisie_denivr = True
    Call calcul_resu
End If
'''If ebpompe.resultat.Denivr > 0 Then
'''    Call calc_val_reelles
'''    Call dessin_pompe
'''Else
'''    Call reini_resu2
'''End If
' If val(Tb_denivr.Text) > 0 Then
'    VURBA = val(Tb_denivr.Text) * SECBA
'    Tb_vurba.Text = Format(VURBA, "###0.00")
'    NRDPH = val(Tb_Ntdph.Text) * VUTBA / VURBA
'    Tb_nrdph.Text = Format(NRDPH, "##0.00")
'End If
''''   Tb_Debit(1).Text = Format(val(Tb_Debit(0).Text) * val(Tb_FPointe.Text), "###0.00")
'End If
'    Call reini_valeurs
    sval_champ = ""
    bKP = False


End Sub

Private Sub Tb_denivr_Click()
Dim mes As String
Dim nom As String
nom = "Tb_denivr"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
owner.affich_aide Me.Name, mes
Call meAffiche
Call sel_text(Tb_denivr)
'''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub

Private Sub Tb_denivr_GotFocus()
Dim mes As String
Dim nom As String
nom = "Tb_denivr"
Call sel_text(Tb_denivr)
'If change_coul Then
'    Change_Couleur nom, 0
'    mes = Rec_Mes(nom, 0)
'    owner.affich_aide Me.Name, mes
'End If
''owner.affich_aide Me.Name, "pompe Conduite Amont"


End Sub

Private Sub Tb_denivr_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_denivr.Text
    iSels = Tb_denivr.SelStart
    iSell = Tb_denivr.SelLength
End If

End Sub
Private Sub Tb_diam_Change()
Dim nom As String
    Call reini_valeurs
nom = "a"
If bKP Then
        nom = verif_cart0(Tb_diam.Text, "Saisie du diamétre de la bâche", "R")
  If nom = "" Then
    Tb_diam.Text = sval_champ
    Tb_diam.SelStart = iSels
    Tb_diam.SelLength = iSell
  End If
    ebpompe.don_techniques.Diamb = txtVersNum(Me.Tb_diam.Text)
    Call calcul_resu
End If
'If val(Tb_diam.Text) > 0 Then
'    SECBA = pi * val(Tb_diam.Text) ^ 2 / 4
'    Call calc_denivt
'End If
'    Call reini_valeurs
    sval_champ = ""
    bKP = False

End Sub
Private Sub Tb_diam_Click()
Dim mes As String
Dim nom As String
nom = "Tb_diam"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
owner.affich_aide Me.Name, mes
Call meAffiche
Call sel_text(Tb_diam)
'''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub

Private Sub Tb_diam_GotFocus()
Dim mes As String
Dim nom As String
nom = "Tb_diam"
Call sel_text(Tb_diam)
'If change_coul Then
'    Change_Couleur nom, 0
'    mes = Rec_Mes(nom, 0)
'    owner.affich_aide Me.Name, mes
'End If
''owner.affich_aide Me.Name, "pompe Conduite Amont"


End Sub

Private Sub Tb_diam_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_diam.Text
    iSels = Tb_diam.SelStart
    iSell = Tb_diam.SelLength
End If

End Sub

Private Sub Tb_Drflt_Change()
Dim nom As String, mes As String
    Call reini_valeurs

nom = "a"
If bKP Then
        nom = verif_cart0(Tb_Drflt.Text, "Saisie du diamètre de la canalisation retenue ", "I")
  If nom = "" Then
    Tb_Drflt.Text = sval_champ
    Tb_Drflt.SelStart = iSels
    Tb_Drflt.SelLength = iSell
  End If
'End If
'Me.Tb_VitRflt.Text = "0.00"
'Me.Tb_Jmpkm.Text = "0.00"
ebpompe.resultat.Drflr = val(Tb_Drflt.Text)
         Call calcul_maj("Drflt")

'''If val(Tb_Drflt.Text) >= 100 Then
'''    If bKP Then
'''    Call calc_vit_pcl
'''    Call calc_val_reelles
'''    Call change_hmt
'''    Lb_Int_Drflt.Caption = "Canalisation"
'''    Lb_Int_Drflt.ForeColor = &H80000012
'''    Tb_Drflt.ForeColor = &H80000012
'''    End If
'''Else
'''    If bKP Then
'''    Lb_Int_Drflt.Caption = "  <100  "
'''    Lb_Int_Drflt.ForeColor = 255
'''    Tb_Drflt.ForeColor = 255
'''    Me.Tb_Drflt.SetFocus
'''   End If
'''   Call reini_resu1
'''End If
'End If
End If

  '  Call reini_valeurs
    sval_champ = ""
    bKP = False

End Sub

Private Sub Tb_Drflt_Click()
Dim mes As String
Dim nom As String
nom = "Tb_Drflt"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
owner.affich_aide Me.Name, mes
Call meAffiche
Call sel_text(Tb_Drflt)
'''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub

Private Sub Tb_Drflt_LostFocus()
Dim mes As String
Dim nom As String
nom = "Tb_Drflt"
Call sel_text(Tb_Drflt)
'If change_coul Then
'    Change_Couleur nom, 0
'    mes = Rec_Mes(nom, 0)
'    owner.affich_aide Me.Name, mes
'End If
''owner.affich_aide Me.Name, "pompe Conduite Amont"
If val(Tb_Drflt.Text) < 100 Then
'    Lb_Int_Drflt.Caption = "<100"
'    Lb_Int_Drflt.ForeColor = 255
'    Tb_Drflt.ForeColor = &H8000000D
''    mes = "Le diamètre de la canalisation est < 100"
''    MsgBox mes
''    Me.Tb_Drflt.SetFocus
End If

End Sub

Private Sub Tb_Drflt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_Drflt.Text
    iSels = Tb_Drflt.SelStart
    iSell = Tb_Drflt.SelLength
End If

End Sub

Private Sub Tb_FPointe_Change()
Dim nom As String
nom = "a"
    Call reini_valeurs

If bKP Then
        nom = verif_cart0(Tb_FPointe.Text, "Saisie du facteur de pointe", "R")
  If nom = "" Then
    Tb_FPointe.Text = sval_champ
    Tb_FPointe.SelStart = iSels
    Tb_FPointe.SelLength = iSell
  End If

    ebpompe.debits_car.Fp = val(Tb_FPointe.Text)
    Call maj_debit
'   tb_Debit(1).Text = Format(val(tb_Debit(0).Text) * val(Tb_FPointe.Text), "###0.00")
End If
    sval_champ = ""
    bKP = False

End Sub

Private Sub Tb_FPointe_Click()
Dim mes As String
Dim nom As String
nom = "Tb_FPointe"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
owner.affich_aide Me.Name, mes
Call meAffiche
Call sel_text(Tb_FPointe)
'''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub

Private Sub Tb_FPointe_GotFocus()
Dim mes As String
Dim nom As String
nom = "Tb_FPointe"
Call sel_text(Tb_FPointe)
'If change_coul Then
'    Change_Couleur nom, 0
'    mes = Rec_Mes(nom, 0)
'    owner.affich_aide Me.Name, mes
'End If
''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub

Private Sub Tb_FPointe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_FPointe.Text
    iSels = Tb_FPointe.SelStart
    iSell = Tb_FPointe.SelLength
End If


End Sub

Private Sub Tb_Geom_Change(Index As Integer)
Dim nom As String
    Call reini_valeurs

nom = "a"
If bKP Then
        Select Case Index
            Case Is = 0
                nom = verif_cart0(Tb_Geom(Index).Text, "Saisie de la longueur du refoulement", "R")
            Case Is = 2
               nom = verif_cart0(Tb_Geom(Index).Text, "Saisie du diamètre théorique de la canalisation", "I")
            Case Is = 3
               nom = verif_cart0(Tb_Geom(Index).Text, "Saisie du niveau du terrain naturel", "R")
            Case Is = 4
               nom = verif_cart0(Tb_Geom(Index).Text, "Saisie du niveau du fil d'eau d'arrivée", "R")
            Case Is = 5
               nom = verif_cart0(Tb_Geom(Index).Text, "Saisie du niveau du fil d'eau de sortie", "R")
            Case Is = 6
               nom = verif_cart0(Tb_Geom(Index).Text, "Saisie du niveau du fil d'eau à l'extrémité du refoulement", "R")
        End Select
  If nom = "" Then
    Tb_Geom(Index).Text = sval_champ
    Tb_Geom(Index).SelStart = iSels
    Tb_Geom(Index).SelLength = iSell
  End If
'End If
Select Case Index
        Case Is = 0
            ebpompe.don_geometrie.Lrflt = txtVersNum(Me.Tb_Geom(0).Text)
        Case Is = 2
            ebpompe.don_geometrie.Drflt = txtVersNum(Me.Tb_Geom(2).Text)
        Case Is = 3
            ebpompe.don_geometrie.NivTN = txtVersNum(Me.Tb_Geom(3).Text)
        Case Is = 4
            ebpompe.don_geometrie.NivEN = txtVersNum(Me.Tb_Geom(4).Text)
        Case Is = 5
            ebpompe.don_geometrie.NivSO = txtVersNum(Me.Tb_Geom(5).Text)
        Case Is = 6
            ebpompe.don_geometrie.NivEX = txtVersNum(Me.Tb_Geom(6).Text)
   End Select
    If Index = 0 Or Index > 2 Then
        Call calcul_resu
'        If ebpompe.don_geometrie.Lrflt > 0 And ebpompe.don_geometrie.Drflt > 0 Then
'            Call calc_vit_pcl
'            Call calc_val_reelles
'            Call change_hmt
'        Else
'            Call reini_resu1
'        End If
 '       Call dessin_pompe
    End If
End If

    sval_champ = ""
    bKP = False

End Sub

Private Sub Tb_Geom_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_Geom"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
Call meAffiche
Call sel_text(Tb_Geom(Index))

End Sub

Private Sub Tb_Geom_GotFocus(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_Geom"
If SSTab1.Tab <> 1 Then
    If Me.SSTab1.TabEnabled(1) Then
        SSTab1.Tab = 1
    Else
        Me.Tb_Geom(0).SetFocus
    End If
'    Me.tb_Debit(0).SetFocus
End If
Call sel_text(Tb_Geom(Index))
'If change_coul Then
'    Change_Couleur nom, Index
'    mes = Rec_Mes(nom, Index)
'    owner.affich_aide Me.Name, mes
'End If
''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub

Private Sub Tb_Geom_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_Geom(Index).Text
    iSels = Tb_Geom(Index).SelStart
    iSell = Tb_Geom(Index).SelLength
End If

End Sub
Private Sub Tb_Jmpkm_Click()
Dim mes As String
Dim nom As String
nom = "Tb_Geom"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
owner.affich_aide Me.Name, mes
'Call sel_text(Tb_Geom(Index))

End Sub
Private Sub Tb_larg_Change()
Dim nom As String
Call reini_valeurs

nom = "a"
If bKP Then
        nom = verif_cart0(Tb_larg.Text, "Saisie de la largeur de la bâche", "R")
  If nom = "" Then
    Tb_larg.Text = sval_champ
    Tb_larg.SelStart = iSels
    Tb_larg.SelLength = iSell
  End If
    ebpompe.don_techniques.Largb = txtVersNum(Me.Tb_larg.Text)
    Call calcul_resu
End If
''If val(Tb_larg.Text) > 0 And val(Tb_long.Text) > 0 Then
'    SECBA = val(Tb_larg.Text) * val(Tb_long.Text)
'    Call calc_denivt
''End If
    sval_champ = ""
    bKP = False

End Sub


Private Sub Tb_larg_Click()
Dim mes As String
Dim nom As String
nom = "Tb_larg"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
owner.affich_aide Me.Name, mes
Call meAffiche
Call sel_text(Tb_larg)
'''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub

Private Sub Tb_larg_GotFocus()
Dim mes As String
Dim nom As String
nom = "Tb_larg"
Call sel_text(Tb_larg)
'If change_coul Then
'    Change_Couleur nom, 0
'    mes = Rec_Mes(nom, 0)
'    owner.affich_aide Me.Name, mes
'End If
''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub

Private Sub Tb_larg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_larg.Text
    iSels = Tb_larg.SelStart
    iSell = Tb_larg.SelLength
End If

End Sub

Private Sub Tb_long_Change()
Dim nom As String
nom = "a"
If bKP Then
        nom = verif_cart0(Tb_long.Text, "Saisie de la largeur de la bâche", "R")
  If nom = "" Then
    Tb_long.Text = sval_champ
    Tb_long.SelStart = iSels
    Tb_long.SelLength = iSell
  End If
ebpompe.don_techniques.Longb = txtVersNum(Me.Tb_long.Text)
Call calcul_resu
End If
''If val(Tb_larg.Text) > 0 And val(Tb_long.Text) > 0 Then
'    SECBA = val(Tb_larg.Text) * val(Tb_long.Text)
'    Call calc_denivt
''End If
    sval_champ = ""
    bKP = False

End Sub

Private Sub Tb_long_Click()
Dim mes As String
Dim nom As String
nom = "Tb_long"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
owner.affich_aide Me.Name, mes
Call meAffiche
Call sel_text(Tb_long)
'''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub

Private Sub Tb_long_GotFocus()
Dim mes As String
Dim nom As String
nom = "Tb_long"
Call sel_text(Tb_long)
'If change_coul Then
'    Change_Couleur nom, 0
'    mes = Rec_Mes(nom, 0)
'    owner.affich_aide Me.Name, mes
'End If
''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub

Private Sub Tb_long_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_long.Text
    iSels = Tb_long.SelStart
    iSell = Tb_long.SelLength
End If

End Sub
Private Sub Tb_Ntdph_Change()
Dim nom As String
    Call reini_valeurs
nom = "a"
If bKP Then
        nom = verif_cart0(Tb_Ntdph.Text, "Saisie du nombre de démarrages", "I")
  If nom = "" Then
    Tb_Ntdph.Text = sval_champ
    Tb_Ntdph.SelStart = iSels
    Tb_Ntdph.SelLength = iSell
  End If
ebpompe.don_techniques.Ntdph = txtVersNum(Me.Tb_Ntdph.Text)
Call calcul_resu
End If
'Call dessin_pompe
'If val(Tb_Nbpom.Text) > 0 And val(Tb_Ntdph.Text) > 0 And val(Tb_Qpomp(1)) > 0 Then
''    VUTBA = Qpomp * 3.6 / 4 / NBPOM / NTDPH '(nbpompes/nb demarrages)
'    VUTBA = val(Tb_Qpomp(1).Text) * 3.6 / 4 / val(Tb_Nbpom.Text) / val(Tb_Ntdph.Text)
'    Tb_Vutba.Text = Format(VUTBA, "###0.00")
'    Frm_bache.Visible = True
'Else
'    Tb_Vutba.Text = "0.00"
'    Frm_bache.Visible = False
'End If
'    Call reini_valeurs
    sval_champ = ""
    bKP = False


End Sub

Private Sub Tb_Ntdph_Click()
Dim mes As String
Dim nom As String
nom = "Tb_Ntdph"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
owner.affich_aide Me.Name, mes
Call sel_text(Tb_Ntdph)
'''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub

Private Sub Tb_Ntdph_GotFocus()
Dim mes As String
Dim nom As String
nom = "Tb_Ntdph"
Call sel_text(Tb_Ntdph)
'If change_coul Then
'    Change_Couleur nom, 0
'    mes = Rec_Mes(nom, 0)
'    owner.affich_aide Me.Name, mes
'End If
''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub

Private Sub Tb_Ntdph_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_Ntdph.Text
    iSels = Tb_Ntdph.SelStart
    iSell = Tb_Ntdph.SelLength
End If

End Sub

Private Sub Tb_Nbpom_Change()
Dim nom As String
    Call reini_valeurs
nom = "a"
If bKP Then
        nom = verif_cart0(Tb_Nbpom.Text, "Saisie du nombre de pompes", "I")
  If nom = "" Then
    Tb_Nbpom.Text = sval_champ
    Tb_Nbpom.SelStart = iSels
    Tb_Nbpom.SelLength = iSell
  End If
ebpompe.don_techniques.Nbpom = txtVersNum(Me.Tb_Nbpom.Text)
Call calcul_resu
End If
'If val(Tb_Nbpom.Text) > 0 And val(Tb_Ntdph.Text) > 0 And val(Tb_Qpomp(1)) > 0 Then
''    VUTBA = Qpomp * 3.6 / 4 / NBPOM / NTDPH '(nbpompes/nb demarrages)
'    VUTBA = val(Tb_Qpomp(1).Text) * 3.6 / 4 / val(Tb_Nbpom.Text) / val(Tb_Ntdph.Text)
'' 20050513 à voir
''    VUTBA = (val(Tb_Qpomp(1).Text) - val(Tb_Qpomp(0).Text)) * 3.6 / 4 / val(Tb_Nbpom.Text) / val(Tb_Ntdph.Text)
'    Tb_Vutba.Text = Format(VUTBA, "###0.00")
'    Frm_bache.Visible = True
'Else
'    Tb_Vutba.Text = "0.00"
'    Frm_bache.Visible = False
'End If
'End If
    sval_champ = ""
    bKP = False

End Sub

Private Sub Tb_Nbpom_Click()
Dim mes As String
Dim nom As String
nom = "Tb_Nbpom"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
owner.affich_aide Me.Name, mes
Call sel_text(Tb_Nbpom)
'''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub

Private Sub Tb_Nbpom_GotFocus()
Dim mes As String
Dim nom As String
nom = "Tb_Nbpom"
If SSTab1.Tab <> 3 Then
    SSTab1.Tab = 3
End If

Call sel_text(Tb_Nbpom)
'If change_coul Then
'    Change_Couleur nom, 0
'    mes = Rec_Mes(nom, 0)
'    owner.affich_aide Me.Name, mes
'End If
''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub

Private Sub Tb_Nbpom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_Nbpom.Text
    iSels = Tb_Nbpom.SelStart
    iSell = Tb_Nbpom.SelLength
End If

End Sub



Private Sub Tb_PtSing_Change(Index As Integer)
Dim nom As String
    Call reini_valeurs

nom = "a"
If bKP Then
        Select Case Index
            Case Is = 0
                nom = verif_cart0(Tb_PtSing(Index).Text, "Saisie du nombre de coudes à 11°15", "I")
            Case Is = 1
               nom = verif_cart0(Tb_PtSing(Index).Text, "Saisie du nombre de coudes à 22°30", "I")
            Case Is = 2
               nom = verif_cart0(Tb_PtSing(Index).Text, "Saisie du nombre de coudes à 30°", "I")
            Case Is = 3
               nom = verif_cart0(Tb_PtSing(Index).Text, "Saisie du nombre de coudes à 45°", "I")
            Case Is = 4
               nom = verif_cart0(Tb_PtSing(Index).Text, "Saisie du nombre de coudes à 90°", "I")
            Case Is = 5
               nom = verif_cart0(Tb_PtSing(Index).Text, "Saisie du nombre de vannes", "I")
            Case Is = 6
               nom = verif_cart0(Tb_PtSing(Index).Text, "Saisie du nombre de clapets anti-retour", "I")
            Case Is = 7
               nom = verif_cart0(Tb_PtSing(Index).Text, "Saisie du nombre de systèmes de vidange", "I")
            Case Is = 8
               nom = verif_cart0(Tb_PtSing(Index).Text, "Saisie du nombre de ventouses", "I")
        End Select
  If nom = "" Then
    Tb_PtSing(Index).Text = sval_champ
    Tb_PtSing(Index).SelStart = iSels
    Tb_PtSing(Index).SelLength = iSell
  End If
'End If
Select Case Index
        Case Is = 0
            ebpompe.pts_singuliers.Nbc1 = txtVersNum(Me.Tb_PtSing(0).Text)
        Case Is = 1
            ebpompe.pts_singuliers.Nbc2 = txtVersNum(Me.Tb_PtSing(1).Text)
        Case Is = 2
            ebpompe.pts_singuliers.Nbc3 = txtVersNum(Me.Tb_PtSing(2).Text)
        Case Is = 3
            ebpompe.pts_singuliers.Nbc4 = txtVersNum(Me.Tb_PtSing(3).Text)
        Case Is = 4
            ebpompe.pts_singuliers.Nbc9 = txtVersNum(Me.Tb_PtSing(4).Text)
        Case Is = 5
            ebpompe.pts_singuliers.Nbva = txtVersNum(Me.Tb_PtSing(5).Text)
        Case Is = 6
            ebpompe.pts_singuliers.Nbcl = txtVersNum(Me.Tb_PtSing(6).Text)
        Case Is = 7
            ebpompe.pts_singuliers.Nbvi = txtVersNum(Me.Tb_PtSing(7).Text)
        Case Is = 8
            ebpompe.pts_singuliers.Nbve = txtVersNum(Me.Tb_PtSing(8).Text)
    End Select
    sval_champ = ""
    Call calcul_resu
''    If bKP Then
''        Call change_ch_sing
''    End If
 End If
   bKP = False


End Sub

Private Sub Tb_PtSing_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_PtSing"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
Call meAffiche
Call sel_text(Tb_PtSing(Index))
'''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub

Private Sub Tb_PtSing_GotFocus(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_PtSing"
If SSTab1.Tab <> 2 Then
    SSTab1.Tab = 2
End If

Call sel_text(Tb_PtSing(Index))
'If change_coul Then
'    Change_Couleur nom, Index
'    mes = Rec_Mes(nom, Index)
'    owner.affich_aide Me.Name, mes
'End If
''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub

Private Sub Tb_PtSing_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_PtSing(Index).Text
    iSels = Tb_PtSing(Index).SelStart
    iSell = Tb_PtSing(Index).SelLength
End If

End Sub

Private Sub Tb_Qpomp_Change(Index As Integer)
Dim diam As Double
    Select Case Index
        Case Is = 0
            Tb_Qpompc(0).Text = 3.6 * val(Tb_Qpomp(0).Text)
            Tb_Qpompc(0).Text = rempl_virgule(Format(val(Tb_Qpomp(0).Text) * 3.6, "###0.00"))
'            If bKP Then
'            Call reini_resu
'            End If
        Case Is = 1
            nom = "a"
            If bKP Then
                        nom = verif_cart0(Tb_Qpomp(Index).Text, "Saisie du débit de pompage retenu", "R")
                If nom = "" Then
                    Tb_Qpomp(Index).Text = sval_champ
                    Tb_Qpomp(Index).SelStart = iSels
                    Tb_Qpomp(Index).SelLength = iSell
                End If
 '           End If
               ebpompe.resultat.Qpomr = txtVersNum(Me.Tb_Qpomp(1).Text)
            'calcul du diamètre de la canalisation de refoulement avec une vitesse par défaut de 1.5 m/s
            ' Diamètre égal
'            If val(Tb_Qpomp(1).Text) = 0 Then
'                Me.Tb_VitRflt.Text = "0.00"
'                Me.Tb_Jmpkm.Text = "0.00"
'                Me.Tb_Drflt.Text = "0"
'                ebpompe.resultat.VitRflt = 0#
'                ebpompe.resultat.jmpkm = 0#
'                ebpompe.resultat.Drflr = 0
'            Else
'                If val(Tb_Qpomp(1).Text) >= val(Tb_Qpomp(0).Text) Then
'           If bKP Then
             Call calcul_maj("Qpomp")
'                 If ebpompe.resultat.Qpomr >= ebpompe.debits_car.Qts Then
'                    diam = 2000 * (ebpompe.resultat.Qpomr / (1000 * pi * 1.5)) ^ 0.5
'                    If diam < 100 Then
'                        diam = 100
'                    End If
'                    Tb_Geom(2).Text = Format(diam, "####")
'                    ebpompe.don_geometrie.Drflt = txtVersNum(Me.Tb_Geom(2).Text)
'                    Tb_Drflt.Text = Format(ebpompe.don_geometrie.Drflt, "###0")
'                    ebpompe.resultat.Drflr = ebpompe.don_geometrie.Drflt
'                    Lb_materiau.Caption = Cb_Materiau.Text & " de D = "
'                    ebpompe.resultat.NatRflr = Lb_materiau.Caption
'                    Call calc_vutba
'                    Lb_int_Qpomp(1).Caption = "Débit de pompage"
'                    Lb_int_Qpomp(1).ForeColor = &H80000012
'                    Tb_Qpomp(1).ForeColor = &H80000012
'                Else
''                Else
'                    Lb_int_Qpomp(1).Caption = " < Qts  "
'                    Lb_int_Qpomp(1).ForeColor = 255
'                    Tb_Qpomp(1).ForeColor = 255
'                    Me.Tb_Qpomp(1).SetFocus
'
'                    Me.Tb_VitRflt.Text = "0.00"
'                    Me.Tb_Jmpkm.Text = "0.00"
'                    Me.Tb_Drflt.Text = "0"
'                    ebpompe.resultat.VitRflt = 0#
'                    ebpompe.resultat.jmpkm = 0#
'                    ebpompe.resultat.Drflr = 0
'                    Me.Tb_Vutba.Text = "0.00"
'                    ebpompe.don_techniques.Vutba = 0#
'                    Me.Tb_denivt.Text = "0.00"
'                    ebpompe.don_techniques.Denivt = 0#
'                    Me.Tb_denivr.Text = "0.00"
'                    ebpompe.resultat.Denivr = 0#
'               End If
'            End If
             End If
                 Tb_Qpompc(1).Text = 3.6 * ebpompe.resultat.Qpomr
                 Tb_Qpompc(1).Text = rempl_virgule(Format(ebpompe.resultat.Qpomr * 3.6, "###0.000"))
 End Select

End Sub

Private Sub Tb_Qpomp_Click(Index As Integer)
Dim mes As String
Dim nom As String
If Index = 1 Then
    nom = "Tb_Qpomp"
    mes = Rec_Mes(nom, Index)
    Change_Couleur nom, Index
    owner.affich_aide Me.Name, mes
Call meAffiche
    Call sel_text(Tb_Qpomp(Index))
    '''owner.affich_aide Me.Name, "pompe Conduite Amont"
End If
End Sub

Private Sub Tb_Qpomp_GotFocus(Index As Integer)
Dim mes As String
Dim nom As String
If Index = 1 Then
    nom = "Tb_Qpomp"
    Call sel_text(Tb_Qpomp(Index))
    'If change_coul Then
    '    Change_Couleur nom, Index
    '    mes = Rec_Mes(nom, Index)
    '    owner.affich_aide Me.Name, mes
    'End If
    ''owner.affich_aide Me.Name, "pompe Conduite Amont"
End If
End Sub

Private Sub Tb_Qpomp_LostFocus(Index As Integer)
    'vérifier que le débit de pompage est supérieur au débit de pointe de temps sec
        If val(Tb_debit(4).Text) > val(Tb_Qpomp(1).Text) Then
'            mes$ = "Le débit de pompage est inférieur au débit de pointe de temps sec" & Chr(13)
'            mes$ = mes$ + "SOUHAITEZ-VOUS MODIFIER ?"
'            MsgBox mes
'            SSTab1.Tab = 0
'            Tb_Qpomp(1).SetFocus
'            Tb_Qpomp(1).SelStart = 0
'            Tb_Qpomp(1).SelLength = Len(Tb_Qpomp(0).Text)
        End If

End Sub


Private Sub Tb_Qpomp_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 1 Then
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_PtSing(Index).Text
    iSels = Tb_PtSing(Index).SelStart
    iSell = Tb_PtSing(Index).SelLength
End If
End If

End Sub
Private Sub Tb_Qpompc_Click(Index As Integer)
Dim mes As String
Dim nom As String
If Index = 1 Then
    nom = "Tb_Qpompc"
    mes = Rec_Mes(nom, Index)
    Change_Couleur nom, Index
    owner.affich_aide Me.Name, mes
Call meAffiche
    Call sel_text(Tb_Qpomp(Index))
    '''owner.affich_aide Me.Name, "pompe Conduite Amont"
End If
End Sub
Private Sub Tb_Singul_Click()
Dim mes As String
Dim nom As String
nom = "Tb_Singul"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
Call meAffiche
owner.affich_aide Me.Name, mes
'Call sel_text(Tb_VitRflt)

End Sub

Private Sub Tb_titre_Change()
    Me.Caption = fen_titre + " : " + Tb_titre.Text
End Sub
Private Sub sel_text(tb_objet As TextBox)
    tb_objet.SelStart = 0
    
    tb_objet.SelLength = Len(tb_objet.Text)
End Sub
Private Sub Tb_VitRflt_Click()
Dim mes As String
Dim nom As String
nom = "Tb_VitRflt"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
owner.affich_aide Me.Name, mes
Call meAffiche
Call sel_text(Tb_VitRflt)
'''owner.affich_aide Me.Name, "pompe Conduite Amont"
End Sub
Private Sub Tb_Vutba_Change()
'    If SECBA > 0 And (Me.Tb_Vutba.Text) > 0 Then
''    Call calcul_resu
'        Call calc_denivt
'    End If
End Sub
Private Sub Tb_Vutba_Click()
Dim mes As String
Dim nom As String
nom = "Tb_Vutba"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
owner.affich_aide Me.Name, mes
Call meAffiche
Call sel_text(Tb_Vutba)
'''owner.affich_aide Me.Name, "pompe Conduite Amont"


End Sub
