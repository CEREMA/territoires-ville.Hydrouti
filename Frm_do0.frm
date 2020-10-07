VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_do 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hydraulique Déversoir d'Orage"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10650
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   10650
   Begin VB.ComboBox Cb_deversoir 
      Height          =   315
      Left            =   1920
      TabIndex        =   91
      Top             =   240
      Width           =   4215
   End
   Begin VB.TextBox Tb_titre 
      Height          =   300
      Left            =   6360
      MaxLength       =   30
      TabIndex        =   89
      Top             =   240
      Width           =   4020
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   8070
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Bassin Versant"
      TabPicture(0)   =   "Frm_do.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Cmd_Sel_Bv"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frm_bv"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Contraintes"
      TabPicture(1)   =   "Frm_do.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Canal. Amont"
      TabPicture(2)   =   "Frm_do.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Cmd_annulm"
      Tab(2).Control(1)=   "Cmd_resum"
      Tab(2).Control(2)=   "Frm_condam"
      Tab(2).Control(3)=   "Lb_mesm"
      Tab(2).Control(4)=   "Lb_hqpluiem"
      Tab(2).Control(5)=   "Lb_vqpluiem"
      Tab(2).Control(6)=   "Label1"
      Tab(2).Control(7)=   "Lb_hqrinm"
      Tab(2).Control(8)=   "Lb_vqrinm"
      Tab(2).Control(9)=   "Lb_hqtsm"
      Tab(2).Control(10)=   "Lb_vqtsm"
      Tab(2).Control(11)=   "Lb_dpsm"
      Tab(2).Control(12)=   "Lb_vpsm"
      Tab(2).ControlCount=   13
      TabCaption(3)   =   "Canal. Aval"
      TabPicture(3)   =   "Frm_do.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Cmd_annulv"
      Tab(3).Control(1)=   "Cmd_Pvm"
      Tab(3).Control(2)=   "Cmd_Pvp"
      Tab(3).Control(3)=   "Cmd_resuv"
      Tab(3).Control(4)=   "Frm_condav"
      Tab(3).Control(5)=   "Lb_Vpsv"
      Tab(3).Control(6)=   "Lb_Dpsv"
      Tab(3).Control(7)=   "Lb_Vqtsv"
      Tab(3).Control(8)=   "Lb_Hqtsv"
      Tab(3).Control(9)=   "Lb_Vqrinv"
      Tab(3).Control(10)=   "Lb_Hqrinv"
      Tab(3).Control(11)=   "Lb_Vqpluiev"
      Tab(3).Control(12)=   "Lb_Hqpluiev"
      Tab(3).Control(13)=   "Lb_mesv"
      Tab(3).ControlCount=   14
      TabCaption(4)   =   "Déversoir"
      TabPicture(4)   =   "Frm_do.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Lb_preste"
      Tab(4).Control(1)=   "Lb_longce(0)"
      Tab(4).Control(2)=   "Lb_longce(1)"
      Tab(4).Control(3)=   "SSTab_result"
      Tab(4).Control(4)=   "Cmd_reinit"
      Tab(4).Control(5)=   "Cmd_VerifDo"
      Tab(4).Control(6)=   "Txtb_deversoir"
      Tab(4).Control(7)=   "Frame3"
      Tab(4).Control(8)=   "Cmd_resudo"
      Tab(4).Control(9)=   "Frame4"
      Tab(4).ControlCount=   10
      TabCaption(5)   =   "Décharge"
      TabPicture(5)   =   "Frm_do.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Lb_vpsdech"
      Tab(5).Control(1)=   "Lb_Dpsdech"
      Tab(5).Control(2)=   "Lb_vqdev"
      Tab(5).Control(3)=   "Lb_hqdev"
      Tab(5).Control(4)=   "lb_Qdev"
      Tab(5).Control(5)=   "Frm_dech"
      Tab(5).Control(6)=   "Cmd_resudech"
      Tab(5).Control(7)=   "Cmd_verifdech"
      Tab(5).Control(8)=   "Txtb_decharge"
      Tab(5).ControlCount=   9
      Begin VB.Frame Frame4 
         Caption         =   "Dessin lignes"
         Height          =   612
         Left            =   -68400
         TabIndex        =   121
         Top             =   3600
         Width           =   3252
         Begin VB.CommandButton OK_lignes 
            Caption         =   "OK"
            Height          =   255
            Left            =   2520
            TabIndex        =   154
            Top             =   240
            Width           =   495
         End
         Begin VB.CheckBox Chk_eau 
            Caption         =   "Eau"
            Height          =   255
            Left            =   1680
            TabIndex        =   124
            Top             =   240
            Width           =   732
         End
         Begin VB.CheckBox Chk_charge 
            Caption         =   "Charge"
            Height          =   255
            Left            =   840
            TabIndex        =   123
            Top             =   240
            Width           =   852
         End
         Begin VB.CheckBox Chk_piezo 
            Caption         =   "Piezo"
            Height          =   255
            Left            =   120
            TabIndex        =   122
            Top             =   240
            Width           =   852
         End
      End
      Begin VB.TextBox Txtb_decharge 
         Height          =   2175
         Left            =   -69960
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   156
         Top             =   720
         Width           =   5055
      End
      Begin VB.CommandButton Cmd_resudo 
         Caption         =   "Calcul"
         Height          =   255
         Left            =   -70920
         TabIndex        =   149
         Top             =   3000
         Width           =   855
      End
      Begin VB.Frame Frame3 
         Caption         =   "caractéristiques"
         Height          =   2295
         Left            =   -74760
         TabIndex        =   88
         Top             =   600
         Width           =   3732
         Begin VB.TextBox Tb_dev 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   3
            Left            =   2400
            MaxLength       =   6
            TabIndex        =   96
            Top             =   1680
            Width           =   750
         End
         Begin VB.TextBox Tb_dev 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   95
            Top             =   1080
            Width           =   750
         End
         Begin VB.TextBox Tb_dev 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   2400
            MaxLength       =   6
            TabIndex        =   94
            Top             =   720
            Width           =   750
         End
         Begin VB.TextBox Tb_dev 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   2400
            MaxLength       =   6
            TabIndex        =   93
            Top             =   360
            Width           =   750
         End
         Begin VB.Label Lb_udev 
            Caption         =   "m"
            Height          =   255
            Index           =   3
            Left            =   3240
            TabIndex        =   143
            Top             =   1725
            Width           =   360
         End
         Begin VB.Label Lb_udev 
            Caption         =   "m/m"
            Height          =   252
            Index           =   2
            Left            =   3240
            TabIndex        =   142
            Top             =   1128
            Width           =   360
         End
         Begin VB.Label Lb_udev 
            Caption         =   "m"
            Height          =   252
            Index           =   1
            Left            =   3240
            TabIndex        =   141
            Top             =   768
            Width           =   360
         End
         Begin VB.Label Lb_udev 
            Caption         =   "m"
            Height          =   252
            Index           =   0
            Left            =   3240
            TabIndex        =   140
            Top             =   360
            Width           =   396
         End
         Begin VB.Label Lb_intdev 
            Caption         =   "Tirant d'eau amont admissible"
            Height          =   300
            Index           =   3
            Left            =   120
            TabIndex        =   120
            Top             =   1725
            Width           =   2295
         End
         Begin VB.Label Lb_intdev 
            Caption         =   " Pente du DO"
            Height          =   300
            Index           =   2
            Left            =   120
            TabIndex        =   99
            Top             =   1130
            Width           =   2175
         End
         Begin VB.Label Lb_intdev 
            Caption         =   "Hauteur de la crête "
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   98
            Top             =   720
            Width           =   2172
         End
         Begin VB.Label Lb_intdev 
            Caption         =   "Longueur du DO"
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   97
            Top             =   410
            Width           =   2175
         End
      End
      Begin VB.TextBox Txtb_deversoir 
         Height          =   2175
         Left            =   -68400
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   153
         Top             =   720
         Width           =   3492
      End
      Begin VB.CommandButton Cmd_VerifDo 
         Caption         =   "Vérif"
         Height          =   255
         Left            =   -68400
         TabIndex        =   151
         Top             =   3000
         Width           =   852
      End
      Begin VB.CommandButton Cmd_reinit 
         Caption         =   "Reinit"
         Height          =   255
         Left            =   -74760
         TabIndex        =   148
         Top             =   3000
         Width           =   855
      End
      Begin VB.CommandButton Cmd_annulv 
         Caption         =   "Annuler"
         Height          =   255
         Left            =   -73200
         TabIndex        =   119
         Top             =   3240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Cmd_annulm 
         Caption         =   "Annuler"
         Height          =   255
         Left            =   -72600
         TabIndex        =   118
         Top             =   3240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Cmd_Pvm 
         Caption         =   "-"
         Height          =   195
         Left            =   -70080
         TabIndex        =   116
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton Cmd_Pvp 
         Caption         =   "+"
         Height          =   195
         Left            =   -70080
         TabIndex        =   115
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton Cmd_verifdech 
         Caption         =   "Vérif"
         Height          =   255
         Left            =   -68880
         TabIndex        =   114
         Top             =   3000
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.CommandButton Cmd_resudech 
         Caption         =   "Calcul"
         Height          =   255
         Left            =   -69960
         TabIndex        =   101
         Top             =   3000
         Width           =   855
      End
      Begin VB.Frame Frm_dech 
         Caption         =   "Conduite"
         Height          =   2055
         Left            =   -74760
         TabIndex        =   100
         Top             =   720
         Width           =   4575
         Begin VB.TextBox TB_dech 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   3
            Left            =   2640
            MaxLength       =   8
            TabIndex        =   105
            Top             =   1440
            Width           =   900
         End
         Begin VB.TextBox TB_dech 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   2640
            MaxLength       =   4
            TabIndex        =   104
            Top             =   1080
            Width           =   900
         End
         Begin VB.TextBox TB_dech 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   2640
            MaxLength       =   4
            TabIndex        =   103
            Top             =   720
            Width           =   900
         End
         Begin VB.TextBox TB_dech 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   2640
            MaxLength       =   6
            TabIndex        =   102
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Lb_udech 
            Height          =   255
            Index           =   2
            Left            =   3720
            TabIndex        =   159
            Top             =   1130
            Width           =   495
         End
         Begin VB.Label Lb_udech 
            Caption         =   "m"
            Height          =   255
            Index           =   3
            Left            =   3720
            TabIndex        =   146
            Top             =   1490
            Width           =   735
         End
         Begin VB.Label Lb_udech 
            Caption         =   "1/10000"
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   145
            Top             =   770
            Width           =   735
         End
         Begin VB.Label Lb_udech 
            Caption         =   "mm"
            Height          =   255
            Index           =   0
            Left            =   3720
            TabIndex        =   144
            Top             =   410
            Width           =   735
         End
         Begin VB.Label Lb_intdech 
            Caption         =   "Longueur de la canalisation "
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   113
            Top             =   1440
            Width           =   2415
         End
         Begin VB.Label Lb_intdech 
            Caption         =   "Coefficient de Manning-Strickler"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   108
            Top             =   1125
            Width           =   2415
         End
         Begin VB.Label Lb_intdech 
            Caption         =   "Pente "
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   107
            Top             =   765
            Width           =   2415
         End
         Begin VB.Label Lb_intdech 
            Caption         =   "Diamètre "
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   106
            Top             =   410
            Width           =   2415
         End
      End
      Begin VB.CommandButton Cmd_resuv 
         Caption         =   "Calcul"
         Height          =   255
         Left            =   -71880
         TabIndex        =   85
         Top             =   3240
         Width           =   855
      End
      Begin VB.Frame Frm_condav 
         Caption         =   "Conduite"
         Height          =   2055
         Left            =   -74760
         TabIndex        =   68
         Top             =   840
         Width           =   4575
         Begin VB.TextBox Tb_ava 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   2640
            MaxLength       =   6
            TabIndex        =   69
            Top             =   360
            Width           =   900
         End
         Begin VB.TextBox Tb_ava 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   2640
            MaxLength       =   4
            TabIndex        =   71
            Top             =   720
            Width           =   900
         End
         Begin VB.TextBox Tb_ava 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   2640
            MaxLength       =   4
            TabIndex        =   73
            Top             =   1080
            Width           =   900
         End
         Begin VB.TextBox Tb_ava 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   3
            Left            =   2640
            MaxLength       =   8
            TabIndex        =   75
            Top             =   1440
            Width           =   900
         End
         Begin VB.Label Lb_uava 
            Height          =   255
            Index           =   2
            Left            =   3720
            TabIndex        =   158
            Top             =   1130
            Width           =   615
         End
         Begin VB.Label Lb_uava 
            Caption         =   "m"
            Height          =   255
            Index           =   3
            Left            =   3720
            TabIndex        =   139
            Top             =   1490
            Width           =   735
         End
         Begin VB.Label Lb_uava 
            Caption         =   "1/10000"
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   138
            Top             =   770
            Width           =   735
         End
         Begin VB.Label Lb_uava 
            Caption         =   "mm"
            Height          =   255
            Index           =   0
            Left            =   3720
            TabIndex        =   137
            Top             =   410
            Width           =   735
         End
         Begin VB.Label Lb_intava 
            Caption         =   "Diamètre "
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   76
            Top             =   410
            Width           =   2415
         End
         Begin VB.Label Lb_intava 
            Caption         =   "Pente "
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   74
            Top             =   770
            Width           =   2415
         End
         Begin VB.Label Lb_intava 
            Caption         =   "Coefficient de Manning-Strickler"
            Height          =   300
            Index           =   2
            Left            =   120
            TabIndex        =   72
            Top             =   1130
            Width           =   2415
         End
         Begin VB.Label Lb_intava 
            Caption         =   "Longueur de la canalisation "
            Height          =   300
            Index           =   3
            Left            =   120
            TabIndex        =   70
            Top             =   1490
            Width           =   2415
         End
      End
      Begin VB.CommandButton Cmd_resum 
         Caption         =   "Calcul"
         Height          =   255
         Left            =   -71280
         TabIndex        =   67
         Top             =   3240
         Width           =   855
      End
      Begin VB.Frame Frm_condam 
         Caption         =   "Conduite"
         Height          =   2055
         Left            =   -74760
         TabIndex        =   49
         Top             =   840
         Width           =   4575
         Begin VB.TextBox Tb_amo 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   3
            Left            =   2640
            MaxLength       =   8
            TabIndex        =   57
            Top             =   1440
            Width           =   900
         End
         Begin VB.TextBox Tb_amo 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   2640
            MaxLength       =   4
            TabIndex        =   56
            Top             =   1080
            Width           =   900
         End
         Begin VB.TextBox Tb_amo 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   2640
            MaxLength       =   4
            TabIndex        =   55
            Top             =   720
            Width           =   900
         End
         Begin VB.TextBox Tb_amo 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   2640
            MaxLength       =   6
            TabIndex        =   54
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Lb_uamo 
            Height          =   255
            Index           =   2
            Left            =   3720
            TabIndex        =   157
            Top             =   1130
            Width           =   495
         End
         Begin VB.Label Lb_uamo 
            Caption         =   "m"
            Height          =   255
            Index           =   3
            Left            =   3720
            TabIndex        =   136
            Top             =   1490
            Width           =   735
         End
         Begin VB.Label Lb_uamo 
            Caption         =   "1/10000"
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   135
            Top             =   770
            Width           =   735
         End
         Begin VB.Label Lb_uamo 
            Caption         =   "mm"
            Height          =   255
            Index           =   0
            Left            =   3720
            TabIndex        =   134
            Top             =   410
            Width           =   735
         End
         Begin VB.Label Lb_intamo 
            Caption         =   "Longueur de la canalisation "
            Height          =   300
            Index           =   3
            Left            =   120
            TabIndex        =   53
            Top             =   1490
            Width           =   2415
         End
         Begin VB.Label Lb_intamo 
            Caption         =   "Coefficient de Manning-Strickler"
            Height          =   300
            Index           =   2
            Left            =   120
            TabIndex        =   52
            Top             =   1125
            Width           =   2415
         End
         Begin VB.Label Lb_intamo 
            Caption         =   "Pente "
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   51
            Top             =   770
            Width           =   2415
         End
         Begin VB.Label Lb_intamo 
            Caption         =   "Diamètre "
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   405
            Width           =   2415
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000080&
         Height          =   4095
         Left            =   -70680
         TabIndex        =   22
         Top             =   360
         Width           =   5865
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   135
            TabIndex        =   48
            Top             =   600
            Width           =   5055
         End
         Begin VB.Shape Forme1 
            Height          =   240
            Index           =   3
            Left            =   1260
            Top             =   2025
            Width           =   240
         End
         Begin VB.Shape Forme1 
            Height          =   330
            Index           =   4
            Left            =   2340
            Top             =   1935
            Width           =   600
         End
         Begin VB.Line Line1 
            X1              =   1485
            X2              =   2340
            Y1              =   2070
            Y2              =   2070
         End
         Begin VB.Line Line2 
            X1              =   1485
            X2              =   2340
            Y1              =   2205
            Y2              =   2205
         End
         Begin VB.Shape Forme1 
            BorderColor     =   &H000000C0&
            Height          =   195
            Index           =   5
            Left            =   3690
            Top             =   1845
            Width           =   240
         End
         Begin VB.Line Line3 
            BorderColor     =   &H000000C0&
            X1              =   2925
            X2              =   3690
            Y1              =   1935
            Y2              =   1935
         End
         Begin VB.Line Line4 
            BorderColor     =   &H000000C0&
            X1              =   2925
            X2              =   3690
            Y1              =   1980
            Y2              =   1980
         End
         Begin VB.Line Line5 
            BorderColor     =   &H0000FF00&
            X1              =   2430
            X2              =   3375
            Y1              =   1935
            Y2              =   765
         End
         Begin VB.Line Line6 
            BorderColor     =   &H0000FF00&
            X1              =   2565
            X2              =   3510
            Y1              =   1935
            Y2              =   765
         End
         Begin VB.Line Line7 
            X1              =   1260
            X2              =   1035
            Y1              =   2070
            Y2              =   2070
         End
         Begin VB.Line Line8 
            X1              =   1260
            X2              =   1035
            Y1              =   2205
            Y2              =   2205
         End
         Begin VB.Line Line9 
            BorderColor     =   &H000000C0&
            X1              =   3915
            X2              =   4320
            Y1              =   1890
            Y2              =   1890
         End
         Begin VB.Line Line10 
            BorderColor     =   &H000000C0&
            X1              =   3960
            X2              =   4365
            Y1              =   1980
            Y2              =   1980
         End
         Begin VB.Line Line11 
            X1              =   2340
            X2              =   2925
            Y1              =   2070
            Y2              =   1980
         End
         Begin VB.Line Line12 
            BorderStyle     =   3  'Dot
            X1              =   1350
            X2              =   1350
            Y1              =   2340
            Y2              =   3240
         End
         Begin VB.Line Line13 
            BorderStyle     =   3  'Dot
            X1              =   3780
            X2              =   3780
            Y1              =   2115
            Y2              =   3195
         End
         Begin VB.Line Line14 
            X1              =   1350
            X2              =   3735
            Y1              =   3150
            Y2              =   3150
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Longueur disponible"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1800
            TabIndex        =   47
            Top             =   2925
            Width           =   1590
         End
         Begin VB.Label Etiquette29 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   8
            Left            =   1845
            TabIndex        =   46
            Top             =   1800
            Width           =   195
         End
         Begin VB.Label Etiquette29 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   9
            Left            =   150
            TabIndex        =   45
            Top             =   3450
            Width           =   195
         End
         Begin VB.Label Etiquette30 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Canalisation amont unitaire"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   4
            Left            =   375
            TabIndex        =   44
            Top             =   3480
            Width           =   2040
         End
         Begin VB.Label Etiquette29 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   10
            Left            =   2520
            TabIndex        =   43
            Top             =   2250
            Width           =   195
         End
         Begin VB.Label Etiquette29 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   11
            Left            =   3285
            TabIndex        =   42
            Top             =   1980
            Width           =   195
         End
         Begin VB.Label Etiquette29 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   12
            Left            =   2655
            TabIndex        =   41
            Top             =   1170
            Width           =   195
         End
         Begin VB.Label Etiquette29 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   13
            Left            =   150
            TabIndex        =   40
            Top             =   3720
            Width           =   195
         End
         Begin VB.Label Etiquette29 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   14
            Left            =   2550
            TabIndex        =   39
            Top             =   3450
            Width           =   195
         End
         Begin VB.Label Etiquette29 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   15
            Left            =   2550
            TabIndex        =   38
            Top             =   3720
            Width           =   195
         End
         Begin VB.Label Etiquette30 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Déversoir d'orage"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   5
            Left            =   375
            TabIndex        =   37
            Top             =   3720
            Width           =   1725
         End
         Begin VB.Label Etiquette30 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Canalisation aval eaux usées"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   6
            Left            =   2775
            TabIndex        =   36
            Top             =   3480
            Width           =   2715
         End
         Begin VB.Label Etiquette30 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Canalisation de décharge eaux pluviales"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   7
            Left            =   2775
            TabIndex        =   35
            Top             =   3720
            Width           =   3000
         End
         Begin VB.Label Etiquette31 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cote de radier obligé aval"
            ForeColor       =   &H000000C0&
            Height          =   465
            Index           =   2
            Left            =   4095
            TabIndex        =   34
            Top             =   1755
            Width           =   1005
         End
         Begin VB.Label Etiquette31 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cote de radier obligé amont"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   90
            TabIndex        =   33
            Top             =   1665
            Width           =   1140
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Exutoire (fossé, ruisseau, rivière,...)"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   225
            TabIndex        =   32
            Top             =   585
            Width           =   2580
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Longueur décharge"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3240
            TabIndex        =   31
            Top             =   1125
            Width           =   1770
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cote des plus hautes eaux"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   270
            TabIndex        =   30
            Top             =   315
            Width           =   2040
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cote radier obligé exutoire"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   3240
            TabIndex        =   29
            Top             =   180
            Width           =   1905
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   180
            TabIndex        =   28
            Top             =   2160
            Width           =   600
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   4320
            TabIndex        =   27
            Top             =   2280
            Width           =   600
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3420
            TabIndex        =   26
            Top             =   1350
            Width           =   600
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3600
            TabIndex        =   25
            Top             =   360
            Width           =   690
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2205
            TabIndex        =   24
            Top             =   315
            Width           =   600
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2205
            TabIndex        =   23
            Top             =   3195
            Width           =   600
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   9
         Top             =   720
         Width           =   3975
         Begin VB.TextBox Tb_cont 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   5
            Left            =   2580
            MaxLength       =   8
            TabIndex        =   15
            Top             =   2160
            Width           =   900
         End
         Begin VB.TextBox Tb_cont 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   4
            Left            =   2580
            MaxLength       =   7
            TabIndex        =   14
            Top             =   1800
            Width           =   900
         End
         Begin VB.TextBox Tb_cont 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   3
            Left            =   2580
            MaxLength       =   7
            TabIndex        =   13
            Top             =   1440
            Width           =   900
         End
         Begin VB.TextBox Tb_cont 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   2580
            MaxLength       =   8
            TabIndex        =   12
            Top             =   1080
            Width           =   900
         End
         Begin VB.TextBox Tb_cont 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   2580
            MaxLength       =   7
            TabIndex        =   11
            Top             =   720
            Width           =   900
         End
         Begin VB.TextBox Tb_cont 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   2580
            MaxLength       =   7
            TabIndex        =   10
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Lb_ucont 
            Caption         =   "m"
            Height          =   255
            Index           =   5
            Left            =   3600
            TabIndex        =   133
            Top             =   2210
            Width           =   200
         End
         Begin VB.Label Lb_ucont 
            Caption         =   "m"
            Height          =   255
            Index           =   4
            Left            =   3600
            TabIndex        =   132
            Top             =   1850
            Width           =   200
         End
         Begin VB.Label Lb_ucont 
            Caption         =   "m"
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   131
            Top             =   1490
            Width           =   200
         End
         Begin VB.Label Lb_ucont 
            Caption         =   "m"
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   130
            Top             =   1130
            Width           =   200
         End
         Begin VB.Label Lb_ucont 
            Caption         =   "m"
            Height          =   255
            Index           =   1
            Left            =   3600
            TabIndex        =   129
            Top             =   770
            Width           =   200
         End
         Begin VB.Label Lb_ucont 
            Caption         =   "m"
            Height          =   255
            Index           =   0
            Left            =   3600
            TabIndex        =   128
            Top             =   410
            Width           =   200
         End
         Begin VB.Label Lb_intcont 
            Caption         =   "Cote de radier obligé à l'exutoire"
            Height          =   300
            Index           =   4
            Left            =   200
            TabIndex        =   21
            Top             =   1850
            Width           =   2295
         End
         Begin VB.Label Lb_intcont 
            Caption         =   "Longueur de la canalisation"
            Height          =   300
            Index           =   5
            Left            =   200
            TabIndex        =   20
            Top             =   2210
            Width           =   2295
         End
         Begin VB.Label Lb_intcont 
            Caption         =   "Cote de radier obligé amont"
            Height          =   300
            Index           =   0
            Left            =   200
            TabIndex        =   19
            Top             =   410
            Width           =   2295
         End
         Begin VB.Label Lb_intcont 
            Caption         =   "Cote de radier obligé aval"
            Height          =   300
            Index           =   1
            Left            =   200
            TabIndex        =   18
            Top             =   770
            Width           =   2295
         End
         Begin VB.Label Lb_intcont 
            Caption         =   "Longueur disponible"
            Height          =   300
            Index           =   2
            Left            =   200
            TabIndex        =   17
            Top             =   1130
            Width           =   2295
         End
         Begin VB.Label Lb_intcont 
            Caption         =   "Cote des P.H.E. à l'exutoire"
            Height          =   300
            Index           =   3
            Left            =   200
            TabIndex        =   16
            Top             =   1490
            Width           =   2295
         End
      End
      Begin VB.Frame Frm_bv 
         Caption         =   "Hydraulique du B.V. "
         Height          =   2055
         Left            =   5280
         TabIndex        =   2
         Top             =   1080
         Width           =   4575
         Begin VB.TextBox Tb_debit 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   2925
            MaxLength       =   8
            TabIndex        =   3
            Top             =   480
            Width           =   900
         End
         Begin VB.TextBox Tb_debit 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   2925
            MaxLength       =   8
            TabIndex        =   5
            Top             =   960
            Width           =   900
         End
         Begin VB.TextBox Tb_debit 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   2925
            MaxLength       =   8
            TabIndex        =   7
            Top             =   1320
            Width           =   900
         End
         Begin VB.Label Lb_udebit 
            Caption         =   "l/s"
            Height          =   255
            Index           =   2
            Left            =   3960
            TabIndex        =   127
            Top             =   1365
            Width           =   405
         End
         Begin VB.Label Lb_udebit 
            Caption         =   "l/s"
            Height          =   255
            Index           =   1
            Left            =   3960
            TabIndex        =   126
            Top             =   1005
            Width           =   405
         End
         Begin VB.Label Lb_udebit 
            Caption         =   "l/s"
            Height          =   255
            Index           =   0
            Left            =   3960
            TabIndex        =   125
            Top             =   540
            Width           =   405
         End
         Begin VB.Label Lb_intdebit 
            Caption         =   "Débit d'eau pluviale "
            Height          =   495
            Index           =   0
            Left            =   600
            TabIndex        =   8
            Top             =   450
            Width           =   2055
         End
         Begin VB.Label Lb_intdebit 
            Caption         =   "Débit de temps sec "
            Height          =   300
            Index           =   1
            Left            =   600
            TabIndex        =   6
            Top             =   1005
            Width           =   2055
         End
         Begin VB.Label Lb_intdebit 
            Caption         =   "Débit de rinçage "
            Height          =   300
            Index           =   2
            Left            =   600
            TabIndex        =   4
            Top             =   1365
            Width           =   2055
         End
      End
      Begin VB.CommandButton Cmd_Sel_Bv 
         Caption         =   "Sélection d'un bassin versant"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   1800
         Width           =   3855
      End
      Begin TabDlg.SSTab SSTab_result 
         Height          =   1815
         Left            =   -72240
         TabIndex        =   147
         Top             =   5160
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   3201
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabHeight       =   420
         TabCaption(0)   =   "Calcul"
         TabPicture(0)   =   "Frm_do.frx":00A8
         Tab(0).ControlEnabled=   0   'False
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "Vérification"
         TabPicture(1)   =   "Frm_do.frx":00C4
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).ControlCount=   0
      End
      Begin VB.Label lb_Qdev 
         Caption         =   "Qdev"
         Height          =   225
         Left            =   -69720
         TabIndex        =   155
         Top             =   1680
         Width           =   4260
      End
      Begin VB.Label Lb_longce 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   1335
         Index           =   1
         Left            =   -67920
         TabIndex        =   152
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Lb_longce 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   2175
         Index           =   0
         Left            =   -70920
         TabIndex        =   150
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Lb_Vpsv 
         Caption         =   "Vps"
         Height          =   225
         Left            =   -69435
         TabIndex        =   84
         Top             =   840
         Width           =   4500
      End
      Begin VB.Label Lb_Dpsv 
         Caption         =   "Dps"
         Height          =   225
         Left            =   -69435
         TabIndex        =   83
         Top             =   1080
         Width           =   4500
      End
      Begin VB.Label Lb_Vqtsv 
         Caption         =   "Vqts"
         Height          =   225
         Left            =   -69435
         TabIndex        =   82
         Top             =   1440
         Width           =   4500
      End
      Begin VB.Label Lb_Hqtsv 
         Caption         =   "Hqts"
         Height          =   225
         Left            =   -69435
         TabIndex        =   81
         Top             =   1680
         Width           =   4500
      End
      Begin VB.Label Lb_Vqrinv 
         Caption         =   "Vqrin"
         Height          =   225
         Left            =   -69435
         TabIndex        =   80
         Top             =   2040
         Width           =   4500
      End
      Begin VB.Label Lb_Hqrinv 
         Caption         =   "Hqrin"
         Height          =   225
         Left            =   -69435
         TabIndex        =   79
         Top             =   2280
         Width           =   4500
      End
      Begin VB.Label Lb_Vqpluiev 
         Caption         =   "Vqpluie"
         Height          =   225
         Left            =   -69435
         TabIndex        =   78
         Top             =   2640
         Width           =   4500
      End
      Begin VB.Label Lb_Hqpluiev 
         Caption         =   "Hqpluie"
         Height          =   225
         Left            =   -69435
         TabIndex        =   77
         Top             =   2880
         Width           =   4500
      End
      Begin VB.Label Lb_preste 
         Caption         =   " "
         Height          =   252
         Left            =   -70680
         TabIndex        =   117
         Top             =   480
         Width           =   4068
      End
      Begin VB.Label Lb_hqdev 
         Caption         =   "Hqdev"
         Height          =   225
         Left            =   -69720
         TabIndex        =   112
         Top             =   2160
         Width           =   4395
      End
      Begin VB.Label Lb_vqdev 
         Caption         =   "Vqdev"
         Height          =   225
         Left            =   -69720
         TabIndex        =   111
         Top             =   1920
         Width           =   4260
      End
      Begin VB.Label Lb_Dpsdech 
         Caption         =   "Dps"
         Height          =   225
         Left            =   -69720
         TabIndex        =   110
         Top             =   1200
         Width           =   4500
      End
      Begin VB.Label Lb_vpsdech 
         Caption         =   "Vps"
         Height          =   225
         Left            =   -69720
         TabIndex        =   109
         Top             =   960
         Width           =   4500
      End
      Begin VB.Label Lb_mesv 
         Height          =   300
         Left            =   -74715
         TabIndex        =   87
         Top             =   5760
         Width           =   9735
      End
      Begin VB.Label Lb_mesm 
         Height          =   300
         Left            =   -74715
         TabIndex        =   86
         Top             =   5760
         Width           =   9735
      End
      Begin VB.Label Lb_hqpluiem 
         Caption         =   "Hqpluie"
         Height          =   225
         Left            =   -69675
         TabIndex        =   66
         Top             =   3000
         Width           =   4395
      End
      Begin VB.Label Lb_vqpluiem 
         Caption         =   "Vqpluie"
         Height          =   225
         Left            =   -69675
         TabIndex        =   65
         Top             =   2760
         Width           =   4395
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   135
         Left            =   -74520
         TabIndex        =   64
         Top             =   5760
         Width           =   15
      End
      Begin VB.Label Lb_hqrinm 
         Caption         =   "Hqrin"
         Height          =   225
         Left            =   -69675
         TabIndex        =   63
         Top             =   2400
         Width           =   4395
      End
      Begin VB.Label Lb_vqrinm 
         Caption         =   "Vqrin"
         Height          =   225
         Left            =   -69675
         TabIndex        =   62
         Top             =   2160
         Width           =   4395
      End
      Begin VB.Label Lb_hqtsm 
         Caption         =   "Hqts"
         Height          =   225
         Left            =   -69675
         TabIndex        =   61
         Top             =   1800
         Width           =   4395
      End
      Begin VB.Label Lb_vqtsm 
         Caption         =   "Vqts"
         Height          =   225
         Left            =   -69675
         TabIndex        =   60
         Top             =   1560
         Width           =   4395
      End
      Begin VB.Label Lb_dpsm 
         Caption         =   "Dps"
         Height          =   300
         Left            =   -69675
         TabIndex        =   59
         Top             =   1200
         Width           =   4395
      End
      Begin VB.Label Lb_vpsm 
         Caption         =   "Vps"
         Height          =   300
         Left            =   -69675
         TabIndex        =   58
         Top             =   960
         Width           =   4395
      End
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "Liste des déversoirs sauvegardés"
      Height          =   255
      Left            =   1920
      TabIndex        =   92
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "Nom du déversoir (30 caract. maxi)"
      Height          =   255
      Left            =   6360
      TabIndex        =   90
      Top             =   0
      Width           =   3975
   End
   Begin VB.Menu mnufichier 
      Caption         =   "&Fichier"
      Begin VB.Menu mnusave 
         Caption         =   "&Enregistrer"
      End
      Begin VB.Menu mnusuppr 
         Caption         =   "&Supprimer"
      End
      Begin VB.Menu mnuquit 
         Caption         =   "&Quitter"
      End
   End
End
Attribute VB_Name = "Frm_do"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private owner As MDIFrm_menu

Private esave As st_savdo
Private nom_fich As String
Private lhFicDbf As Long
Private FileLength As Integer

Public Sub retailler()
retaille

End Sub
Private Sub retaille()
    Me.Left = owner.fcom.Width
    Me.Top = 0
    Me.Width = owner.Width - owner.fcom.Width - 200
    Me.Height = owner.fdessin.Top
End Sub



Private Sub Cb_deversoir_click()
Dim za As st_savdo
    lhFicDbf = FreeFile
    Open nom_fich For Random Access Read As #lhFicDbf Len = Len(za)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za
   If Not EOF(lhFicDbf) Then
       If Trim(za.nom) = Trim(Cb_deversoir.Text) Then
        Tb_titre = za.nom
        edo = za.edo
'        Debug.Print ebv.lghydr
        edessdo = za.edessdo
        Call ini_form_exist
        do_sauve = False
    End If
   End If

Loop
Close #lhFicDbf

End Sub

Private Sub Cmd_annulm_Click()
    Me.Tb_amo(0).Text = Format(edessdo.tron_amo.conduit.Diametre * 1000, "###0")
    Me.Tb_amo(1).Text = Format(edessdo.tron_amo.conduit.pente * 10000, "###0")
    Me.Tb_amo(3).Text = Format(edessdo.tron_amo.conduit.Longueur, "###0")
    Me.Tb_amo(2).Text = Format(edessdo.tron_amo.conduit.rugosite, "###0")
    Me.Cmd_annulm.Visible = False
    Me.Cmd_resum.Enabled = False
    Call calc_amont
End Sub


Private Sub Cmd_annulv_Click()
    Me.Tb_ava(0).Text = Format(edessdo.tron_ava.conduit.Diametre * 1000, "###0")
    Me.Tb_ava(1).Text = Format(edessdo.tron_ava.conduit.pente * 10000, "###0")
    Me.Tb_ava(3).Text = Format(edessdo.tron_ava.conduit.Longueur, "###0")
    Me.Tb_ava(2).Text = Format(edessdo.tron_ava.conduit.rugosite, "###0")
    Me.Cmd_annulv.Visible = False
    Me.Cmd_resuv.Enabled = False
    Call calc_aval
End Sub

Private Sub Cmd_Pvm_Click()
    Me.Tb_ava(1).Text = Format(edessdo.iradav - 1, "###0")
    Me.Cmd_annulv.Visible = True
    Me.Cmd_resuv.Enabled = True
'Call calcul_aval
End Sub

Private Sub Cmd_Pvp_Click()
    Me.Tb_ava(1).Text = Format(edessdo.iradav + 1, "###0")
    Me.Cmd_annulv.Visible = True
    Me.Cmd_resuv.Enabled = True
'Call calcul_aval

End Sub

Private Sub Cmd_reinit_Click()
    Me.Tb_dev(0).Text = "0.0"
    Me.Tb_dev(1).Text = "0.0"
    Me.Tb_dev(2).Text = "0.0"
    Me.Tb_dev(3).Text = "0.0"
    Me.Cmd_resudo.Enabled = True
    Call ini_longce
    Call init_graphdo(owner.fdessin.UC_graphique1)
    Call dess_troncon(owner.fdessin.UC_graphique1, edessdo.tron_amo, couleur.noir)
    Call dess_troncon(owner.fdessin.UC_graphique1, edessdo.tron_ava, couleur.noir)
'calcul hydraulique aval
    Call calc_aval
'dessin des lignes hydrauliques aval
    Call dess_piezo(owner.fdessin.UC_graphique1, edessdo.tron_amo, edessdo.Qts, couleur.bleu) ' vbBlue)
    Call dess_piezo(owner.fdessin.UC_graphique1, edessdo.tron_amo, edessdo.Qrin, couleur.rouge) ' vbRed)
    Call dess_piezo(owner.fdessin.UC_graphique1, edessdo.tron_amo, edessdo.Qpluie, couleur.magenta) ' vbMagenta)
    Call dess_piezo(owner.fdessin.UC_graphique1, edessdo.tron_ava, edessdo.Qts, couleur.bleu) ' vbBlue)
    Call dess_piezo(owner.fdessin.UC_graphique1, edessdo.tron_ava, edessdo.Qrin, couleur.rouge) ' vbRed)
    Call dess_piezo(owner.fdessin.UC_graphique1, edessdo.tron_ava, edessdo.Qpluie, couleur.magenta) ' vbMagenta)
'    Call init_graphdo(UC_graphique4)
'    Call dess_troncon(UC_graphique4, edessdo.tron_amo, couleur.rouge) ' vbRed)
'    Call dess_troncon(UC_graphique4, edessdo.tron_ava, couleur.rouge) ' vbRed)
End Sub

Private Sub Cmd_resudech_Click()
Call calcul_dech
End Sub
Private Sub calcul_dech()
Dim qv As deb_vit
Dim canal As conduite
Dim Qdev As Double
Dim qps As Double, vps As Double, hautdech As Double, vdech As Double
Dim ltc() As Variant, ct() As Variant, qvm(5) As Variant
Dim qcal As Double

Dim lambda As Double, dH As Double, hcr As Double, hmin As Double, Nen As Double, Ned As Double
Dim nok As Boolean
Dim i As Integer
Dim ok1 As Boolean
Dim mes_Res As String, mes_Res1 As String, mes_res_do As String
nok = True
i = 0
mes_Res = ""
mes_Res1 = ""
edessdo.tron_dech.conduit.typ = 2
edessdo.tron_dech.Absamo = edo.Absava
'edessdo.tron_dech.radamo = edo.radava
edessdo.tron_dech.Absava = edessdo.tron_dech.Absamo + edessdo.tron_dech.conduit.Longueur
'edessdo.tron_dech.radava = edessdo.tron_dech.radamo - edessdo.tron_dech.conduit.Longueur * edessdo.tron_dech.conduit.pente
edessdo.tron_dech.radava = edessdo.rdoex
edessdo.tron_dech.radamo = edessdo.tron_dech.radava + edessdo.tron_dech.conduit.Longueur * edessdo.tron_dech.conduit.pente
canal = edessdo.tron_dech.conduit
'dessin dans ucg5
'Call init_graphdech(UC_graphique5)
'
'Call dess_troncon(UC_graphique5, edessdo.tron_amo, couleur.noir) ' vbRed)
'Call dess_troncon(UC_graphique5, edessdo.tron_dech, couleur.noir)  'vbmagenta)
'
'Call dess_do(Me.UC_graphique5, edo, couleur.noir)
'Call dess_troncon(UC_graphique5, edessdo.tron_ava, couleur.noir) 'vbRed)

'dessin dans frmdessin
Call init_graphdech(owner.fdessin.UC_graphique2)
Call dess_troncon(owner.fdessin.UC_graphique2, edessdo.tron_ava, couleur.gris_clair) 'vbRed)
Call dess_troncon(owner.fdessin.UC_graphique2, edessdo.tron_amo, couleur.gris) ' vbRed)
Call dess_do(owner.fdessin.UC_graphique2, edo, couleur.gris)
Call dess_troncon(owner.fdessin.UC_graphique2, edessdo.tron_dech, couleur.noir) 'vbmagenta)

qv = debvit_ps(canal)
qps = qv.debit
vps = qv.vitesse
Me.Lb_vpsdech.Caption = "Vitesse pleine section = " + Str(Round(qv.vitesse, 3)) + " m/s"
Me.Lb_Dpsdech.Caption = "Débit pleine section = " + Str(Round(qv.debit, 3)) + " m3/s"
mes_Res = "Débit pleine section = " + Str(Round(qv.debit, 3)) + " m3/s"
mes_Res = mes_Res + Chr(13) + Chr(10) + "Vitesse pleine section = " + Str(Round(qv.vitesse, 3)) + " m/s"

' calcul du débit déversé
'a revoir
 'qdev= edessdo.qpluie - qavp
'Qdev = edessdo.Qpluie - 1.3 * edessdo.Qrin

While nok And i < 20
    mes_Res1 = ""
    nok = False
    Qdev = edo_res.Qdev
    Me.lb_Qdev.Caption = "Débit déversé : " + Str(Round(Qdev, 3)) + " m3/s"
        mes_Res1 = Chr(13) + Chr(10) + Chr(13) + Chr(10) + "Débit déversé : " + Str(Round(Qdev, 3)) + " m3/s"
    If Qdev > qv.debit Then
        Me.Lb_vqdev.Caption = "Vitesse d'écoulement à Qdev > Qps = " + "   " + " m/s"
        Me.Lb_hqdev.Caption = "Hauteur d'eau à Qdev > Qps = " + "   " + " m"
        mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Vitesse d'écoulement à Qdev > Qps = " + "   " + " m/s"
        mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Hauteur d'eau à Qdev > Qps = " + "   " + " m"
    Else
    
        Call cana(canal, ct)
        ltc = calc_par(canal)
        qvi = caltran1(Qdev * 1000, ct, ltc)
        hautdech = qvi(5)
        vdech = qvi(2)
        '    Debug.Print "qvi : 1= débit"; qvi(1); " 2=vitesse"; qvi(2); "qvi"; " 3=acceleration"; qvi(3)
        '    Debug.Print " 4=Largeur libre"; qvi(4); " 5 = hauteur; "; qvi(5); ""
        Me.Lb_vqdev.Caption = "Vitesse d'écoulement à Q déversé = " + Str(Round(qvi(2), 3)) + " m/s"
        Me.Lb_hqdev.Caption = "Hauteur d'eau Q déversé = " + Str(Round(qvi(5), 3)) + " m"
        mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Vitesse d'écoulement à Q déversé = " + Str(Round(qvi(2), 3)) + " m/s"
        mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Hauteur d'eau Q déversé = " + Str(Round(qvi(5), 3)) + " m"
    End If
    ' determination du niveau de la surface d'eau amont zradier + hauteur
    ' par rapport au niveau des plus hautes eaux aval
    If edessdo.phex > edessdo.tron_dech.radamo + hautdech Then
            MsgBox "remous", vbOKOnly
            mes_Res = mes_Res1 + Chr(13) + Chr(10) + "Remous"
       
    Else
    ' determination du regime torrentiel ou fluvial
        regime = verif_regime(Qdev, canal)
        Select Case regime
            Case "TORREN."
                ' calcul des niveaux d'energie
                'a revoir saisie du lambda
                mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Regime Torrentiel"
                lambda = 0.75
                'niveau d'energie nécessaire
                'calcul de la charge necessaire a l'entree coefficient d'entonnement 0.75
                dH = lambda * vdech ^ 2 / (2 * 9.81)
                Nen = edessdo.tron_dech.radamo + hautdech + dH
                'niveau d'energie disponible
                hcr = (((Qdev / edo.Longueur) ^ 2) / 9.81) ^ (1# / 3#)
                vcr = (9.81 * hcr) ^ 0.5
                hmin = hcr + (vcr ^ 2) / (2 * 9.81)
                Ned = edo.radamo + edo.hauteur + hmin
                If Ned >= Nen Then
'                    MsgBox "Nv energie disponible " & Str(Round(Ned, 3)) & " > Nv necessaire" & Str(Round(Nen, 3))
                    mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Nv energie disponible " & Str(Round(Ned, 3)) & " > Nv necessaire" & Str(Round(Nen, 3))
                Else
                    MsgBox "Ecoulement torrentiel " + Chr(13) + " Nv energie disponible " & Str(Round(Ned, 3)) & " < Ne necessaire " & Str(Round(Nen, 3)) & Chr(10) & " Modifier le dispositif "
                    mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Nv energie disponible " & Str(Round(Ned, 3)) & " < Ne necessaire " & Str(Round(Nen, 3)) & Chr(10) & " Modifier le dispositif "
                End If
                
                'a revoir suite torrentiel
                
            Case "FLUVIAL"
                mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Regime Fluvial"
                ' calcul des niveaux d'energie
                'a revoir saisie du lambda
                lambda = 0.75
                'niveau d'energie nécessaire
                'calcul de la charge necessaire a l'entree coefficient d'entonnement 0.75
                dH = lambda * vdech ^ 2 / (2 * 9.81)
                Nen = edessdo.tron_dech.radamo + hautdech + dH
                'niveau du seuil
                Dim zSeuil As Double, h As Double
                zSeuil = edo.radamo + edo.hauteur
                If zSeuil >= Nen Then
'                    MsgBox "nappe libre", vbOKOnly
                    mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "nappe libre"
                Else
'                    MsgBox "nappe noyée " + Str(i), vbOKOnly
                    mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "nappe noyée"
                    h = Nen - zSeuil
                    ' calcul du coef C de nappe noyee h/hm
                    Debug.Print h / edo_res.HM
                    edo_res.c = recup_do_C(h / edo_res.HM)
                    nok = True
                    i = i + 1
'                     mes_res_do = verif_do_charge(edo, edo_res, edessdo.tron_amo, edessdo.Qpluie / 1000, edessdo.Qrin / 1000, UC_graphique4)
                     mes_res_do = verif_do_charge(edo, edo_res, edessdo.tron_amo, edessdo.Qpluie / 1000, edessdo.Qrin / 1000, owner.fdessin.UC_graphique1)
                    Me.Txtb_deversoir.Text = mes_res_do
                    If Abs(edo_res.Qdev - Qdev) < 0.0001 Then
                       nok = False
                    End If
                End If
                'niveau d'energie disponible
        End Select
    End If
Wend
If Not nok Then
mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Hauteur de la lame d'eau :"
mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Amont  :" + Str(Round(edo_res.Ham, 3)) + " m"
mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Aval  :" + Str(Round(edo_res.Hav, 3)) + " m"
mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Hauteur de la charge :"
mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Amont  :" + Str(Round(edo_res.Haam, 3)) + " m"
mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Aval  :" + Str(Round(edo_res.Haav, 3)) + " m"

mes_Res = mes_Res + mes_Res1
Txtb_decharge.Text = mes_Res
'Dessin du fonctionnement dans l'onglet Decharge
'dessin de la ligne d'eau
' conduite amont

'dessin des lignes de charges

 Dim zplam_am As Double, zplam_av As Double, zplav_am As Double, zplav_av As Double
'Dim qvm(5) As Variant, haut As Double, pentmot As Double
Dim tr As troncon, uc_g As UC_graphique
Dim res_conduit As debit_conduit
'Dim qcal
'dessin de la ligne dans le deversoir
'Set uc_g = UC_graphique5
'dans la frmdessin
Set uc_g = owner.fdessin.UC_graphique2
'zplam_av = edo.radamo + edo_res.Tram
zplam_av = edo.radamo + edessdo.Tram

tr = edessdo.tron_amo
qcal = edessdo.Qpluie / 1000
res_conduit = calc_debit_tr(tr, qcal)

'dessin troncon amont
    'dessin des lignes d'eau
    zplam_am = res_conduit.hauteur + tr.radamo
' uc_g.dess_lign tr.Absamo, zplam_am, tr.Absava, zplam_av, couleur.bleu
 res_conduit.zphe_ava = zplam_av
 Call inter_piezo_eau(tr, res_conduit)
' uc_g.dess_lign tr.Absamo, zplam_am, res_conduit.piezointer.X, res_conduit.piezointer.Y, couleur.bleu, 2
' uc_g.dess_lign res_conduit.piezointer.X, res_conduit.piezointer.Y, tr.Absava, zplam_av, couleur.bleu, 2
 uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.zeau_amo.Y, res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, couleur.bleu, 2
 uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, res_conduit.zeau_ava.X, res_conduit.zeau_ava.Y, couleur.bleu, 2


'    zplam_av = res_conduit.hauteur + tr.radava
    'dessin charge
    uc_g.dess_lign tr.Absamo, zplam_am + res_conduit.dcharge, tr.Absava, edo.radamo + edo_res.Haam, couleur.rouge, 2

'dessin ligne d'eau sur la lame
    'dessin des lignes d'eau
'    uc_g.dess_lign edo.Absamo, edo.radamo + edo_res.Tram, edo.Absava, edo.radamo + edo.hauteur + edo_res.Hav, couleur.bleu, 2
    uc_g.dess_lign edo.Absamo, edo.radamo + edessdo.Tram, edo.Absava, edo.radamo + edo.hauteur + edo_res.Hav, couleur.bleu, 2
    'dessin charge
    uc_g.dess_lign edo.Absamo, edo.radamo + edo_res.Haam, edo.Absava, edo.radava + edo_res.Haav, couleur.rouge, 2
 
 'dessin troncon décharge
 
    tr = edessdo.tron_dech
     
'     If edessdo.phex > (edessdo.tron_dech.radava + hautdech) Then
'        zplam_av = edessdo.phex
'    Else
        zplam_av = edessdo.tron_dech.radava + hautdech
'    End If
    zplam_am = edessdo.tron_dech.radamo + hautdech
    'dessin des lignes d'eau
    uc_g.dess_lign tr.Absamo, zplam_am, tr.Absava, zplam_av, couleur.jaune, 2
    'dessin charge
    res_conduit = calc_debit_tr(edessdo.tron_dech, Qdev)
    res_conduit.zphe_ava = edessdo.phex
     Call inter_piezo_eau(tr, res_conduit)
' uc_g.dess_lign tr.Absamo, zplam_am, res_conduit.piezointer.X, res_conduit.piezointer.Y, couleur.bleu, 2
' uc_g.dess_lign res_conduit.piezointer.X, res_conduit.piezointer.Y, tr.Absava, zplam_av, couleur.bleu, 2
 uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.zeau_amo.Y, res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, couleur.bleu, 2
 uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, res_conduit.zeau_ava.X, res_conduit.zeau_ava.Y, couleur.bleu, 2

     If edessdo.phex > (edessdo.tron_dech.radava + hautdech) Then
        zplam_av = edessdo.phex
    Else
        zplam_av = res_conduit.chargeava
    End If
    
    zplam_am = res_conduit.chargeamo
    uc_g.dess_lign tr.Absamo, zplam_am, tr.Absava, zplam_av, couleur.rouge, 2
Else
  mes_Res = mes_Res + Chr(13) + Chr(10) + Chr(13) + Chr(10) + " Anomalie de fonctionnement " + Chr(13) + Chr(10) + "Redimensionner !" + mes_Res1
  Txtb_decharge.Text = mes_Res

End If

End Sub
Private Function recup_do_C(ByVal rap As Double) As Double
Dim list_hC(11, 2) As Double
Dim a As Double

Dim i As Integer

list_hC(1, 1) = 0
list_hC(2, 1) = 0.1
list_hC(3, 1) = 0.2
list_hC(4, 1) = 0.3
list_hC(5, 1) = 0.4
list_hC(6, 1) = 0.5
list_hC(7, 1) = 0.6
list_hC(8, 1) = 0.7
list_hC(9, 1) = 0.8
list_hC(10, 1) = 0.9
list_hC(11, 1) = 1#
list_hC(1, 2) = 1#
list_hC(2, 2) = 0.99
list_hC(3, 2) = 0.98
list_hC(4, 2) = 0.97
list_hC(5, 2) = 0.96
list_hC(6, 2) = 0.94
list_hC(7, 2) = 0.91
list_hC(8, 2) = 0.86
list_hC(9, 2) = 0.78
list_hC(10, 2) = 0.62
list_hC(11, 2) = 0



i = 1
While rap > list_hC(i, 1) And i < UBound(list_hC)
    i = i + 1
    
Wend
If i = 1 Then
    i = 2
End If
a = (rap - list_hC(i - 1, 1)) * (list_hC(i, 2) - list_hC(i - 1, 2)) / (list_hC(i, 1) - list_hC(i - 1, 1)) + list_hC(i - 1, 2)
If a < 0 Then
    a = 0.1
End If
recup_do_C = a
End Function

Private Sub Cmd_resudo_Click()
Dim longdo As Double, hautdo As Double, pentedo As Double
Dim lgreste As Double, preste As Double
'Dim edo As deversoir
Dim sresult As String
Dim troamo As troncon, troava As troncon
longdo = txtVersNum(Me.Tb_dev(0).Text)
hautdo = txtVersNum(Me.Tb_dev(1).Text)
pentedo = txtVersNum(Me.Tb_dev(2).Text)
edo.Longueur = longdo
edo.hauteur = hautdo
edo.pente = pentedo
Call pre_dimdo(edo)
Me.Tb_dev(0).Text = Format(edo.Longueur, "##0.00")
Me.Tb_dev(1).Text = Format(edo.hauteur, "###0.000")
Me.Tb_dev(2).Text = Format(edo.pente, "#0.0000")
Call modi_longce
sresult = pre_calculdo(edo)
    Me.Lb_longce(0).Caption = sresult
    Me.SSTab_result.Tab = 0
    Me.Txtb_deversoir.Text = ""

'calcul longueur restante
lgreste = edessdo.lgdisp - edo.tron_ava.Absava
If lgreste > 0 Then
'calcul pente restante
    preste = (edo.tron_ava.radava - edessdo.rdoav) / lgreste * 10000#
    Me.Lb_preste.Caption = "Pente disponible = " + Str(Round(preste)) + "  1/10000"
Else
    Me.Lb_preste.Caption = ""
End If
' dessin dans l UCG4


'dessin dans  frmdessin
Call dessin_do


Me.Cmd_resudo.Enabled = False
Me.SSTab1.TabEnabled(5) = True
End Sub
Sub dessin_do()
Call init_graphdo(owner.fdessin.UC_graphique1)
Call init_graphdo(owner.fdessin.UC_graphique2)

Call dessin_do_objet(owner.fdessin.UC_graphique1)

Call dessin_do_hydrau(owner.fdessin.UC_graphique1)
End Sub
Sub dessin_do_objet(ByRef uc_g As UC_graphique)

Call dess_troncon(uc_g, edessdo.tron_amo, couleur.gris) ' vbBlack)
Call dess_predo(uc_g, edo, couleur.noir)
Call dess_cot(uc_g, couleur.noir) ' vbBlack)
'Call dess_troncon(UC_graphique4, edessdo.tron_ava, vbRed)
End Sub
Sub dessin_do_hydrau(ByRef uc_g As UC_graphique)
    Call dess_piezo(uc_g, edessdo.tron_amo, edessdo.Qts, couleur.magenta) ' vbGreen) ' RGB(128, 64, 128))
    Call dess_piezo(uc_g, edessdo.tron_amo, edessdo.Qrin, couleur.magenta_clair) ' vbRed)
    Call dess_piezo(uc_g, edessdo.tron_amo, edessdo.Qpluie, couleur.vert) ' vbMagenta)
   Call dess_piezo(uc_g, edo.tron_ava, edessdo.Qts, couleur.magenta) ' vbGreen)
'    Call dess_piezo(uc_g, edo.tron_ava, edessdo.Qrin, couleur.magenta_clair) ' vbRed)
 '   Call dess_piezo(UC_graphique4, troava, edessdo.Qpluie, vbMagenta)
    Call dess_charge(uc_g, edo.tron_ava, edessdo.Qrin, couleur.orange) ' vbCyan)
    Call dess_charge(uc_g, edessdo.tron_amo, edessdo.Qrin, couleur.orange) ' vbCyan)


End Sub

Private Sub lect_fich()
Dim za As st_savdo
    lhFicDbf = FreeFile
    Cb_deversoir.Clear
    Open nom_fich For Random Access Read As #lhFicDbf Len = Len(za)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za
   If Not EOF(lhFicDbf) Then
       Cb_deversoir.AddItem (za.nom)
   End If
Loop
Close #lhFicDbf
Cb_deversoir.Text = ""
Cb_deversoir.Refresh
End Sub



Private Sub Cmd_verifdech_Click()
Dim ok As Boolean
ok = True
ok = ok And verif_dech()

End Sub

Private Sub Cmd_VerifDo_Click()
Dim ok1 As Boolean
Dim message As String
Dim reponse As Integer
Dim sresult As String

Call modi_longce(2)
message = ""
Me.Txtb_deversoir.Text = message
ok1 = True
message = message + verif_remous_do()
message = message + verif_ecoul_am_cr()
message = message + verif_ecoul_av_cr()
message = message + verif_Hauteur_am_cr()
message = message + verif_Hauteur_cphe()
edessdo.Tram = txtVersNum(Me.Tb_dev(3))
'edo_res.Tram = txtVersNum(Me.tb_dev(3))
edo_res.c = 1#
sresult = verif_do_charge(edo, edo_res, edessdo.tron_amo, edessdo.Qpluie / 1000#, edessdo.Qrin / 1000#, owner.fdessin.UC_graphique1)
Me.Lb_longce(1).Caption = sresult
Me.SSTab_result.Tab = 1
Me.Txtb_deversoir.Text = sresult + Chr(13) + Chr(10) + message
 If message <> "" Then
    
'Me.Txtb_deversoir.Text = sresult + Chr(13) + Chr(10) + message
    reponse = MsgBox(message, 0, "Vérification du déversoir")
End If
End Sub


Private Sub MnuSave_Click()
Dim za As st_savdo
Dim i As Integer, isave As Integer
Dim reponse As Integer
If Trim(Tb_titre.Text) <> "" Then
   lhFicDbf = FreeFile
'   Debug.Print Len(esave)
    Open nom_fich For Random Access Read Write As #lhFicDbf Len = Len(esave)
    i = 0
    isave = 0
    Do While Not EOF(lhFicDbf)
        Get #lhFicDbf, , za
        If Not EOF(lhFicDbf) Then
            i = i + 1
            If Trim(za.nom) = Trim(Tb_titre.Text) Then
                isave = i
            End If
       End If
    Loop
    If isave > 0 Then
    reponse = MsgBox("Le nom est déjà utilisé. Le remplacer?", 4, "Sauvegarde d'un déversoir")
        If reponse = 6 Then
            esave.type = "déversoir"
            esave.nom = Tb_titre.Text
            esave.edessdo = edessdo
            esave.edo = edo
           Put #lhFicDbf, isave, esave
        End If
    Else
        esave.type = "déversoir"
        esave.nom = Tb_titre.Text
        esave.edessdo = edessdo
        esave.edo = edo
        FileLength = LOF(lhFicDbf) / Len(esave) + 1
        Put #lhFicDbf, FileLength, esave
    End If
        Close #lhFicDbf
        Call lect_fich
        Cb_deversoir.Text = Trim(Tb_titre.Text)
        do_sauve = False
Else
    reponse = MsgBox("Le nom du déversoir n'est pas renseigné.", , "Sauvegarde d'un déversoir")
End If
End Sub

Private Sub Cmd_resum_Click()
    calcul_amont
End Sub
Private Sub ini_longce(Optional ByVal i As Integer)
If i = 0 Then
    Me.Lb_longce(0).BackColor = &H8000000B
    Me.Lb_longce(0).BorderStyle = 0
    Me.Lb_longce(0).Caption = ""
    Me.Lb_longce(1).BackColor = &H8000000B
    Me.Lb_longce(1).BorderStyle = 0
    Me.Lb_longce(1).Caption = ""
    Txtb_deversoir.Text = ""
    Else
    Me.Lb_longce(i - 1).BackColor = &H8000000B
    Me.Lb_longce(i - 1).BorderStyle = 0
    Me.Lb_longce(i - 1).Caption = ""
    Txtb_deversoir.Text = ""
End If
End Sub
Private Sub modi_longce(Optional ByVal i As Integer)
If i = 0 Then
    Me.Lb_longce(0).BackColor = &H80000009
    Me.Lb_longce(0).BorderStyle = 1
    Me.Lb_longce(1).BackColor = &H80000009
    Me.Lb_longce(1).BorderStyle = 1
    Else
    Me.Lb_longce(i - 1).BackColor = &H80000009
    Me.Lb_longce(i - 1).BorderStyle = 1
End If
End Sub
Private Sub ini_resum()
        If edessdo.dam > 0 And edessdo.iRadam > 0 And edessdo.Kam > 0 _
            And edessdo.Lam > 0 Then
            If Cmd_resum.Enabled = False Then
                Call ini_valeurs_m
                Me.Cmd_resum.Enabled = True
                Me.Cmd_annulm.Visible = True
            End If
        End If

End Sub
Private Sub ini_resuv()
        If edessdo.dav > 0 And edessdo.iradav > 0 And edessdo.kav > 0 _
            And edessdo.Lav > 0 Then
            If Cmd_resuv.Enabled = False Then
                Call ini_valeurs_v
                Me.Cmd_resuv.Enabled = True
                Me.Cmd_annulv.Visible = True
            End If
        End If

End Sub
Private Sub calc_amont()
Dim Qts As Double, Qrin As Double, Qpluie As Double
Dim qv As deb_vit
Dim ltc() As Variant, ct() As Variant, qvm(5) As Variant
Dim qcal As Double
Dim ecoulam As String, betam As Double
Dim message As String

Dim cana_amo As conduite

cana_amo = edessdo.tron_amo.conduit
'wagner Call calcul_condam(cana_amo)

Qts = edessdo.Qts
Qrin = edessdo.Qrin
Qpluie = edessdo.Qpluie

qv = debvit_ps(cana_amo)
'Debug.Print qv.debit * 1000, qv.vitesse
Me.Lb_vpsm.Caption = "Vitesse pleine section = " + Str(Round(qv.vitesse, 3)) + " m/s"
Me.Lb_dpsm.Caption = "Débit pleine section = " + Str(Round(qv.debit, 3)) + " m3/s"
If Qpluie > qv.debit * 1000 Or Qts > qv.debit * 1000 Or Qrin > qv.debit * 1000 Then
    If Qpluie > qv.debit * 1000 Then
        message = message + "Débit d'eau pluviale "
    End If
    If Qts > qv.debit * 1000 Then
        If Trim(message) <> "" Then
            message = message + ", "
        End If
        message = message + "Débit de temps sec "
    End If
    If Qrin > qv.debit * 1000 Then
        If Trim(message) <> "" Then
            message = message + ", "
        End If
        message = message + "Débit de rinçage "
    End If
    message = message + "> Débit de pleine section"
    MsgBox message, vbOKOnly, "Condition d'écoulement amont"
    Me.Lb_vqtsm.Caption = "Vitesse d'écoulement à QTS = " + "   " + " m/s"
    Me.Lb_hqtsm.Caption = "Hauteur d'eau QTS = " + "   " + " m"
    Me.Lb_vqrinm.Caption = "Vitesse d'écoulement à QRIN = " + "   " + " m/s"
    Me.Lb_hqrinm.Caption = "Hauteur d'eau QRIN = " + "   " + " m"
    Me.Lb_vqpluiem.Caption = "Vitesse d'écoulement à QPLUIE = " + "   " + " m/s"
    Me.Lb_hqpluiem.Caption = "Hauteur d'eau QPLUIE = " + "   " + " m"

Else
    qcal = Qts
    Call cana(cana_amo, ct)
    ltc = calc_par(cana_amo)
    qvi = caltran1(qcal, ct, ltc)
    Me.Lb_vqtsm.Caption = "Vitesse d'écoulement à QTS = " + Str(Round(qvi(2), 3)) + " m/s"
    Me.Lb_hqtsm.Caption = "Hauteur d'eau QTS = " + Str(Round(qvi(5), 3)) + " m"
    Debug.Print "qvi : 1= débit"; qvi(1); " 2=vitesse"; qvi(2); "qvi"; " 3=acceleration"; qvi(3)
    Debug.Print " 4=Largeur libre"; qvi(4); " 5 = hauteur; "; qvi(5); ""
'                vitmax = qvm(2)
'                qvm = caltran1(qps / 10#, ct, ltc)
'                vit10 = qvm(2)
'                qvm = caltran1(qps / 100#, ct, ltc)
'                vit100 = qvm(2)
    qcal = Qrin
    Call cana(cana_amo, ct)
    ltc = calc_par(cana_amo)
    qvi = caltran1(qcal, ct, ltc)
    Me.Lb_vqrinm.Caption = "Vitesse d'écoulement à QRIN = " + Str(Round(qvi(2), 3)) + " m/s"
    Me.Lb_hqrinm.Caption = "Hauteur d'eau QRIN = " + Str(Round(qvi(5), 3)) + " m"
    qcal = Qpluie
    Call cana(cana_amo, ct)
    ltc = calc_par(cana_amo)
    qvi = caltran1(qcal, ct, ltc)
    Me.Lb_vqpluiem.Caption = "Vitesse d'écoulement à QPLUIE = " + Str(Round(qvi(2), 3)) + " m/s"
    Me.Lb_hqpluiem.Caption = "Hauteur d'eau QPLUIE = " + Str(Round(qvi(5), 3)) + " m"
' calcul du regime amont
 ' a revoir
        betam = angle(Qpluie / (qv.debit * 1000))
        betam = beta
        ecoulam = calcul_ecoul(Qpluie / 1000, cana_amo.Diametre, betam)
        
  Me.SSTab1.TabEnabled(3) = True
End If

End Sub

Private Sub calcul_amont()
Dim troamo As troncon, troava As troncon
Dim cana_amo As conduite
Dim cana_ava As conduite
' conduite amont -> troncon amont
    cana_amo.Diametre = edessdo.dam / 1000#
    cana_amo.Longueur = edessdo.Lam
    cana_amo.pente = edessdo.iRadam / 10000#
    cana_amo.rugosite = edessdo.Kam
    cana_amo.typ = 2
    With troamo
      .Absamo = 0#
      .Absava = .Absamo + cana_amo.Longueur
      .conduit = cana_amo
      .radamo = edessdo.rdoam
      .radava = edessdo.rdoam - cana_amo.Longueur * cana_amo.pente
    End With
    edessdo.tron_amo = troamo


' calcul hydraulique
    Call calc_amont
    
'repositionnement de cana_ava
    cana_ava = edessdo.tron_ava.conduit
    If cana_ava.Diametre > 0 And cana_ava.Longueur > 0 And cana_ava.pente > 0 And cana_ava.rugosite > 0 Then
       With edessdo.tron_ava
            .Absamo = edessdo.tron_amo.Absava + edo.Longueur
            .Absava = .Absamo + cana_ava.Longueur
            .conduit = cana_ava
'       .radava = edessdo.rdoav
'       .radamo = edessdo.rdoav + cana_ava.Longueur * cana_ava.pente
            .radamo = edessdo.tron_amo.radava - edo.Longueur * edo.pente
            .radava = .radamo - cana_ava.Longueur * cana_ava.pente
        End With
        
        Me.Cmd_resuv.Enabled = True
End If
' dessin a partir de canal amont
 Call dessin_amont
' reinitialisation des
    Call ini_longce
    Me.Cmd_annulm.Visible = False
    Me.Cmd_resum.Enabled = False
    Me.Cmd_resudo.Enabled = True
If troamo.radava <= edessdo.rdoav Then
    MsgBox "cote aval canalisation amont inférieure à cote radier obligé aval", vbOKOnly
End If
End Sub
Private Sub dessin_amont()
Dim troamo As troncon, troava As troncon
Dim cana_ava As conduite
  troamo = edessdo.tron_amo
' dessin a partir de canal amont
' dessin de troncon amont
'    Call ini_gra
    Call init_graphdo(owner.fdessin.UC_graphique1)
'    Call dess_tronc(troamo, couleur.noir) ' vbBlack)
    Call dess_troncon(owner.fdessin.UC_graphique1, troamo, couleur.noir)
' dessin de troncon aval s'il existe
    cana_ava = edessdo.tron_ava.conduit
    If cana_ava.Diametre > 0 And cana_ava.Longueur > 0 And cana_ava.pente > 0 And cana_ava.rugosite > 0 Then
        troava = edessdo.tron_ava
'        Call dess_tronc(troava, couleur.noir) ' vbBlack)
        Call dess_troncon(owner.fdessin.UC_graphique1, troava, couleur.noir)
    End If
    
' dessin des lignes hydrauliques
    Call dess_piezo(owner.fdessin.UC_graphique1, troamo, edessdo.Qts, couleur.bleu) ' vbBlue)
    Call dess_piezo(owner.fdessin.UC_graphique1, troamo, edessdo.Qrin, couleur.rouge) ' vbRed)
    Call dess_piezo(owner.fdessin.UC_graphique1, troamo, edessdo.Qpluie, couleur.magenta) ' vbMagenta)
End Sub
Private Sub ini_valeurs_v()
    Me.Lb_Vpsv.Caption = ""
    Me.Lb_Dpsv.Caption = ""
    Me.Lb_Vqtsv.Caption = ""
    Me.Lb_Hqtsv.Caption = ""
    Me.Lb_Vqrinv.Caption = ""
    Me.Lb_Hqrinv.Caption = ""
    Me.Lb_Vqpluiev.Caption = ""
    Me.Lb_Hqpluiev.Caption = ""

End Sub
Private Sub ini_valeurs_m()
    Me.Lb_vpsm.Caption = ""
    Me.Lb_dpsm.Caption = ""
    Me.Lb_vqtsm.Caption = ""
    Me.Lb_hqtsm.Caption = ""
    Me.Lb_vqrinm.Caption = ""
    Me.Lb_hqrinm.Caption = ""
    Me.Lb_vqpluiem.Caption = ""
    Me.Lb_hqpluiem.Caption = ""

End Sub

Private Sub Cmd_resuv_Click()
Call calcul_aval
End Sub
Private Sub calc_aval()
Dim qv As deb_vit
Dim ltc() As Variant, ct() As Variant, qvm(5) As Variant
Dim qcal As Double
Dim Qts As Double, Qrin As Double, Qpluie As Double
Dim ecoulam As String, betam As Double
Dim message As String
Dim cana_ava As conduite
cana_ava = edessdo.tron_ava.conduit

Qts = edessdo.Qts
Qrin = edessdo.Qrin
Qpluie = edessdo.Qpluie

qv = debvit_ps(cana_ava)
Me.Lb_Vpsv.Caption = "Vitesse pleine section = " + Str(Round(qv.vitesse, 3)) + " m/s"
Me.Lb_Dpsv.Caption = "Débit pleine section = " + Str(Round(qv.debit, 3)) + " m3/s"
If Qts > qv.debit * 1000 Then
    Me.Lb_Vqtsv.Caption = "Vitesse d'écoulement à QTS = " + "   " + " m/s"
    Me.Lb_Hqtsv.Caption = "Hauteur d'eau QTS = " + "   " + " m"
Else

    qcal = Qts
    Call cana(cana_ava, ct)
    ltc = calc_par(cana_ava)
    qvi = caltran1(qcal, ct, ltc)
    Me.Lb_Vqtsv.Caption = "Vitesse d'écoulement à QTS = " + Str(Round(qvi(2), 3)) + " m/s"
    Me.Lb_Hqtsv.Caption = "Hauteur d'eau QTS = " + Str(Round(qvi(5), 3)) + " m"
    Debug.Print "qvi : 1= débit"; qvi(1); " 2=vitesse"; qvi(2); "qvi"; " 3=acceleration"; qvi(3)
    Debug.Print " 4=Largeur libre"; qvi(4); " 5 = hauteur; "; qvi(5); ""
End If
If Qrin > qv.debit * 1000 Then
    Me.Lb_Vqrinv.Caption = "Vitesse d'écoulement à QRIN = " + "   " + " m/s"
    Me.Lb_Hqrinv.Caption = "Hauteur d'eau QRIN = " + "   " + " m"
Else
    qcal = Qrin
    Call cana(cana_ava, ct)
    ltc = calc_par(cana_ava)
    qvi = caltran1(qcal, ct, ltc)
    Me.Lb_Vqrinv.Caption = "Vitesse d'écoulement à QRIN = " + Str(Round(qvi(2), 3)) + " m/s"
    Me.Lb_Hqrinv.Caption = "Hauteur d'eau QRIN = " + Str(Round(qvi(5), 3)) + " m"
End If
If Qpluie > qv.debit * 1000 Then
    Me.Lb_Vqpluiev.Caption = "Vitesse d'écoulement à QPLUIE = " + "   " + " m/s"
    Me.Lb_Hqpluiev.Caption = "Hauteur d'eau QPLUIE = " + "   " + " m"
Else
    qcal = Qpluie
    Call cana(cana_ava, ct)
    ltc = calc_par(cana_ava)
    qvi = caltran1(qcal, ct, ltc)
    Me.Lb_Vqpluiev.Caption = "Vitesse d'écoulement à QPLUIE = " + Str(Round(qvi(2), 3)) + " m/s"
    Me.Lb_Hqpluiev.Caption = "Hauteur d'eau QPLUIE = " + Str(Round(qvi(5), 3)) + " m"
End If

End Sub

Private Sub calcul_aval()
Dim troava As troncon, troamo As troncon
Dim cana_ava As conduite

' conduite aval -> troncon aval
    cana_ava.Diametre = edessdo.dav / 1000#
    cana_ava.Longueur = edessdo.Lav
    cana_ava.pente = edessdo.iradav / 10000#
    cana_ava.rugosite = edessdo.kav
    cana_ava.typ = 2
    With troava
'      .Absamo = edessdo.lgdisp - cana_ava.Longueur
      .Absamo = edessdo.tron_amo.Absava + edo.Longueur
      .Absava = .Absamo + cana_ava.Longueur
      .conduit = cana_ava
'      .radava = edessdo.rdoav
'      .radamo = edessdo.rdoav + cana_ava.Longueur * cana_ava.pente
      .radamo = edessdo.tron_amo.radava - edo.Longueur * edo.pente
      .radava = .radamo - cana_ava.Longueur * cana_ava.pente
    End With
    edessdo.tron_ava = troava
'dessin troncon amont troncon aval
    troamo = edessdo.tron_amo
'    Call ini_gra
'    Call dess_tronc(troamo, couleur.noir) ' vbBlack)
'    Call dess_tronc(troava, couleur.noir) ' vbBlack)
    Call init_graphdo(owner.fdessin.UC_graphique1)
    Call dess_troncon(owner.fdessin.UC_graphique1, troamo, couleur.noir)
    Call dess_troncon(owner.fdessin.UC_graphique1, troava, couleur.noir)
'calcul hydraulique aval
    Call calc_aval
'dessin des lignes hydrauliques aval
    Call dess_piezo(owner.fdessin.UC_graphique1, troamo, edessdo.Qts, couleur.bleu) ' vbBlue)
    Call dess_piezo(owner.fdessin.UC_graphique1, troamo, edessdo.Qrin, couleur.rouge) ' vbRed)
    Call dess_piezo(owner.fdessin.UC_graphique1, troamo, edessdo.Qpluie, couleur.magenta) ' vbMagenta)
    Call dess_piezo(owner.fdessin.UC_graphique1, troava, edessdo.Qts, couleur.bleu) ' vbBlue)
    Call dess_piezo(owner.fdessin.UC_graphique1, troava, edessdo.Qrin, couleur.rouge) ' vbRed)
    Call dess_piezo(owner.fdessin.UC_graphique1, troava, edessdo.Qpluie, couleur.magenta) ' vbMagenta)
'reinitailisation des autres onglets
    Call ini_longce
    Me.SSTab1.TabEnabled(4) = True
    Me.Cmd_annulv.Visible = False
    Me.Cmd_resuv.Enabled = False
    Me.Cmd_resudo.Enabled = True
    
'If troava.radamo > troamo.radava Then
'    MsgBox "cote amont canalisation aval  supérieure à cote aval canalisation amont", vbOKOnly
'End If
End Sub

Private Sub Cmd_Sel_Bv_Click()
    Me.Enabled = False
    do_bv = True
    
    frm_bv2.Show
'    Frm_bv2.Init_ss_commentaire

    owner.affich_aide frm_bv2.Name, "Calcul de débit de bassin versant"
End Sub

Private Sub dess_tronc(ByRef tr As troncon, ByRef ocolor As ColorConstants)
'Call dess_troncon(Me.UC_graphique2, tr, ocolor)
'Call dess_troncon(Me.UC_graphique3, tr, ocolor)
'Call dess_troncon(Me.UC_graphique4, tr, ocolor)

End Sub

Private Sub init_graphdech(ByRef uc_graph As UC_graphique)

Dim ecx As Double
Dim i As Integer
Dim maxX As Double

uc_graph.graphique_clear
uc_graph.reinit 7, "Arial"
uc_graph.init_arrondi_X 2
uc_graph.init_arrondi_y 3
uc_graph.init_MinX 0#
If edessdo.lgca > edo.tron_ava.conduit.Longueur Then
maxX = edessdo.tron_amo.conduit.Longueur + edo.Longueur + edessdo.lgca
Else
maxX = edessdo.tron_amo.conduit.Longueur + edo.Longueur + edo.tron_ava.conduit.Longueur
End If
uc_graph.init_MaxX maxX
uc_graph.init_EchXn 1
ecx = uc_graph.lire_EchXn()
uc_graph.init_EchY ecx * 10
While Not ok
    ok = True
    uc_graph.init_MinY Int(edessdo.rdoex)
    uc_graph.init_Ech_MaxYn
    If uc_graph.lire_MaxYn < edessdo.rdoam + 1.3 * edessdo.tron_amo.conduit.Diametre Then
        uc_graph.init_EchY uc_graph.lire_EchYn / 2
        ok = False
    End If
Wend

'uc_graph.init_MinY Int(edessdo.rdoex)
uc_graph.init_Ech_MaxYn

uc_graph.dess_lign 0, edessdo.rdoam, maxX, edessdo.rdoex, couleur.vert, 1 ' vbGreen
End Sub
Private Sub Form_Load()
Dim nom As String
    nom = chemin_app + "do_bassin.bmp"
    Set owner = MDIFrm_menu.rec_owner
    Call retaille
'    owner.affich_aide Me.Name, "Deversoir"



nom_fich = chemin_app + "deversoir.bin"
do_sauve = True
' lecture fichier
    If Dir(nom_fich) <> "" Then
        Call lect_fich
    End If
Call ini_longce
Me.SSTab1.Tab = 0
Me.SSTab1.TabEnabled(1) = False
Me.SSTab1.TabEnabled(2) = False
Me.SSTab1.TabEnabled(3) = False
Me.SSTab1.TabEnabled(4) = False
Me.SSTab1.TabEnabled(5) = False
owner.fdessin.Image2.Visible = False
owner.fdessin.Image1.Visible = False
owner.fdessin.UC_graphiqueB.Visible = True
owner.fdessin.UC_graphiqueB.reinit 7, "Arial"
owner.fdessin.UC_graphiqueB.graphique_clear
owner.fdessin.UC_graphiqueB.init_title
owner.fdessin.UC_graphiqueB.init_titleh ""
owner.fdessin.UC_graphiqueB.init_titleb ""
'    owner.fdessin.UC_graphiqueB.Top = 0
'    owner.fdessin.UC_graphiqueB.Left = 2500
'    owner.fdessin.UC_graphiqueB.Height = 4210
'    owner.fdessin.UC_graphiqueB.Width = 7800
owner.fdessin.UC_graphiqueB.init_fond nom
owner.fdessin.UC_graphique1.Visible = False
owner.fdessin.UC_graphique2.Visible = False
Call init_graphique1
owner.fdessin.UC_graphique2.reinit 7, "Arial"
owner.fdessin.UC_graphique2.graphique_clear
'owner.fdessin.UC_graphique2.Top = 0
'owner.fdessin.UC_graphique2.Left = 350 '240
'owner.fdessin.UC_graphique2.Height = 4400 '4210
'owner.fdessin.UC_graphique2.Width = 10000 '9855
owner.fdessin.UC_graphique2.init_title
owner.fdessin.UC_graphique2.init_titleh ""
owner.fdessin.UC_graphique2.init_titleb ""
'    UC_graphique1.ecr_texta 300, 1200, "SURFACE = 120 Ha", "G", "B"
Call ini_edessdo
Call ini_form
'Call ini_debit
Me.Tb_debit(0) = "0.0"
Me.Tb_debit(2) = "0.0"
Me.Tb_debit(1) = "0.0"
End Sub
Private Sub init_graphique1()
owner.fdessin.UC_graphique1.graphique_clear
'owner.fdessin.UC_graphique1.Top = 0
'owner.fdessin.UC_graphique1.Left = 350 '240
'owner.fdessin.UC_graphique1.Height = 4400 '4210
'owner.fdessin.UC_graphique1.Width = 10000 '9855
owner.fdessin.UC_graphique1.init_title
owner.fdessin.UC_graphique1.init_titleh ""
owner.fdessin.UC_graphique1.init_titleb ""
owner.fdessin.UC_graphique1.reinit 7, "Arial"

End Sub
Public Sub ini_edessdo()
Dim canal As conduite
Dim tronc As troncon

    edessdo.Qts = 0#
    edessdo.Qrin = 0#
    edessdo.Qpluie = 0#
    edessdo.rdoam = 0#
    edessdo.rdoav = 0#
    edessdo.lgdisp = 0#
    edessdo.phex = 0#
    edessdo.rdoex = 0#
    edessdo.lgca = 0#
    edessdo.dam = 0
    edessdo.iRadam = 0
    edessdo.Kam = 0
    edessdo.Lam = 0
    edessdo.dav = 0
    edessdo.iradav = 0
    edessdo.kav = 0
    edessdo.Lav = 0
    edo.Longueur = 0#
    edo.hauteur = 0#
    edo.pente = 0#
End Sub
Public Sub ini_debit(ByVal nom As String)
Call init_graphique1
    owner.fdessin.Image1.Visible = False
    owner.fdessin.Image2.Visible = False
    owner.fdessin.UC_graphiqueB.Visible = True
    owner.fdessin.UC_graphique1.Visible = False
    owner.fdessin.UC_graphique2.Visible = False
 
    If Trim(ebv.Qchoisi) <> "" Then
        Select Case ebv.Qchoisi
            Case Is = "CAQUOT"
                Me.Tb_debit(0).Text = Format(ebv.Qcor * 1000, "####0.0")
                Me.Lb_intdebit(0).Caption = "Débit d'eau pluviale (CAQUOT)"
            Case Is = "RATION"
                Me.Tb_debit(0).Text = Format(ebv.Qmr * 1000, "####0.0")
                Me.Lb_intdebit(0).Caption = "Débit d'eau pluviale (Rationnelle)"
            Case Is = "HYDROG"
                Me.Tb_debit(0).Text = Format(ebv.Qhydro * 1000, "####0.0")
                Me.Lb_intdebit(0).Caption = "Débit d'eau pluviale (Hydrogramme)"
        End Select
        Me.Tb_debit(1).Text = Format(ebv.Qts, "###0.0")
        Me.Tb_debit(2).Text = Format(ebv.Qrin, "###0.0")
        owner.fdessin.UC_graphiqueB.ecr_texta 1665, 1140, "Surface = " + Str(ebv.surface) + " Ha", "G", "B"
        owner.fdessin.UC_graphiqueB.ecr_texta 1620, 1530, "Coef. de ruissellement = " + Str(ebv.imper), "G", "B"
        owner.fdessin.UC_graphiqueB.ecr_texta 2055, 1995, "Nombre d'habitants = " + Str(ebv.nhab), "G", "B"
        owner.fdessin.UC_graphiqueB.ecr_texta 2160, 2505, "Taux de dilution = " + Str(ebv.tdilu), "G", "B"
        owner.fdessin.UC_graphiqueB.ecr_texta 2875, 540, "Longueur = " + Str(ebv.lghydr) + " m", "G", "B"
        owner.fdessin.UC_graphiqueB.ecr_texta 3145, 810, "Pente = " + Str(ebv.phydr) + " (1/10000)", "G", "B"
        Me.SSTab1.TabEnabled(1) = True
    Else
        Me.Tb_debit(0) = "0.0"
        Me.Tb_debit(2) = "0.0"
        Me.Tb_debit(1) = "0.0"
        owner.fdessin.UC_graphiqueB.ecr_texta 1665, 1140, "Surface", "G", "B"
        owner.fdessin.UC_graphiqueB.ecr_texta 1620, 1530, "Coef. de ruissellement", "G", "B"
        owner.fdessin.UC_graphiqueB.ecr_texta 2055, 1995, "Nombre d'habitants", "G", "B"
        owner.fdessin.UC_graphiqueB.ecr_texta 2160, 2505, "Taux de dilution", "G", "B"
        owner.fdessin.UC_graphiqueB.ecr_texta 2875, 540, "Longueur", "G", "B"
        owner.fdessin.UC_graphiqueB.ecr_texta 3145, 810, "Pente", "G", "B"
        Me.SSTab1.TabEnabled(1) = False
   End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
'    frm_menu.Enabled = True
Unload owner.fdessin
owner.recharge_commentaire
End Sub
Private Sub MnuQuit_Click()
    Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim reponse As Integer
If do_sauve Then
    reponse = MsgBox("Le déversoir n'a pas été enregistré" + Chr(10) _
        + "Voulez vous le sauvegarder?", 4, "Sauvegarde du déversoir")
    If reponse = 6 Then ' 6=oui,7=non,2=annuler
        Call MnuSave_Click
    End If
End If
 '   Cancel = True
End Sub
Private Sub reini_form(ntab As Integer)
Select Case ntab
    Case Is = 1
        Me.SSTab1.TabEnabled(3) = False
        Me.SSTab1.TabEnabled(4) = False
        Me.SSTab1.TabEnabled(5) = False
        If edessdo.Qpluie > 0 And edessdo.Qts > 0 And edessdo.Qrin > 0 Then
            Me.SSTab1.TabEnabled(1) = True
        Else
            Me.SSTab1.TabEnabled(1) = False
        End If
    Case Is = 2
        Me.SSTab1.TabEnabled(3) = False
        Me.SSTab1.TabEnabled(4) = False
        Me.SSTab1.TabEnabled(5) = False
       If edessdo.lgca > 0 And edessdo.lgdisp > 0 And edessdo.phex > 0 _
            And edessdo.rdoam > 0 And edessdo.rdoav > 0 And edessdo.rdoex > 0 Then
           Me.SSTab1.TabEnabled(2) = True
           Call ini_cana
           Call ini_gra
        Else
            Me.SSTab1.TabEnabled(2) = False
        End If
End Select
End Sub
Private Sub ini_form()
    Me.Tb_debit(1).Text = Format(edessdo.Qts, "###0.0")
    Me.Tb_debit(2).Text = Format(edessdo.Qrin, "###0.0")
    Me.Tb_debit(0).Text = Format(edessdo.Qpluie, "###0.0")
    Me.Tb_cont(0).Text = Format(edessdo.rdoam, "###0.0")
    Me.Tb_cont(1).Text = Format(edessdo.rdoav, "###0.0")
    Me.Tb_cont(2).Text = Format(edessdo.lgdisp, "###0.0")
    Me.Tb_cont(3).Text = Format(edessdo.phex, "###0.0")
    Me.Tb_cont(4).Text = Format(edessdo.rdoex, "###0.0")
    Me.Tb_cont(5).Text = Format(edessdo.lgca, "###0.0")
    Me.Tb_amo(0).Text = Format(edessdo.dam, "###0")
    Me.Tb_amo(1).Text = Format(edessdo.iRadam, "###0")
    Me.Tb_amo(3).Text = Format(edessdo.Lam, "###0")
    Me.Tb_amo(2).Text = Format(edessdo.Kam, "###0")
    Me.Tb_ava(0).Text = Format(edessdo.dav, "###0")
    Me.Tb_ava(1).Text = Format(edessdo.iradav, "###0")
    Me.Tb_ava(3).Text = Format(edessdo.Lav, "###0")
    Me.Tb_ava(2).Text = Format(edessdo.kav, "###0")
    Me.Tb_dev(0).Text = Format(edo.Longueur, "##0.00")
    Me.Tb_dev(1).Text = Format(edo.hauteur, "###0.000")
    Me.Tb_dev(2).Text = Format(edo.pente, "##0.0000")
    Me.Tb_dev(3).Text = Format(edessdo.Tram, "##0.00")
    Me.TB_dech(0).Text = Format(edessdo.tron_dech.conduit.Diametre * 1000, "###0")
    Me.TB_dech(1).Text = Format(edessdo.tron_dech.conduit.pente * 10000, "###0")
    Me.TB_dech(3).Text = Format(edessdo.tron_dech.conduit.Longueur, "###0")
    Me.TB_dech(2).Text = Format(edessdo.tron_dech.conduit.rugosite, "###0")

End Sub
Private Sub ini_form_exist()
Call ini_form
reini_form 1
reini_form 2
calcul_amont
calcul_aval
Me.SSTab1.Tab = 0
End Sub
Private Sub ini_cana()
edessdo.tron_amo.conduit.Diametre = 0
edessdo.tron_amo.conduit.Longueur = 0
edessdo.tron_amo.conduit.pente = 0
edessdo.tron_amo.conduit.rugosite = 0
edessdo.tron_ava.conduit.Diametre = 0
edessdo.tron_ava.conduit.Longueur = 0
edessdo.tron_ava.conduit.pente = 0
edessdo.tron_ava.conduit.rugosite = 0
End Sub
Private Sub ini_gra()
End Sub
Private Sub reini_valeurs()
Call ini_valeurs_m
Call ini_valeurs_v
Call ini_longce
Me.Cmd_annulm.Visible = False
Me.Cmd_resum.Enabled = True
Me.Cmd_annulv.Visible = False
Me.Cmd_resuv.Enabled = True
Me.Cmd_resudo.Enabled = True
End Sub


Private Sub mnusuppr_Click()
Dim za As st_savdo
Dim lhFicDbf1 As Integer
If Trim(Cb_deversoir.Text) <> "" Then
    nom = chemin_app + "tempbas.bin"
    lhFicDbf = FreeFile
    Open nom_fich For Random Access Read As #lhFicDbf Len = Len(za)
    lhFicDbf1 = FreeFile
    Open nom For Random Access Write As #lhFicDbf1 Len = Len(za)
    Do While Not EOF(lhFicDbf)
    '   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
        Get #lhFicDbf, , za
       If Not EOF(lhFicDbf) Then
           If Trim(za.nom) <> Trim(Cb_deversoir.Text) Then
            FileLength = LOF(lhFicDbf1) / Len(za) + 1
            Put #lhFicDbf1, FileLength, za
        End If
       End If
    Loop
    Close #lhFicDbf
    Close #lhFicDbf1
    Kill nom_fich
    lhFicDbf = FreeFile
    Open nom_fich For Random Access Write As #lhFicDbf Len = Len(za)
    lhFicDbf1 = FreeFile
    Open nom For Random Access Read As #lhFicDbf1 Len = Len(za)
    Do While Not EOF(lhFicDbf1)
    '   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
        Get #lhFicDbf1, , za
       If Not EOF(lhFicDbf1) Then
            FileLength = LOF(lhFicDbf) / Len(za) + 1
            Put #lhFicDbf, FileLength, za
       End If
    Loop
    Close #lhFicDbf
    Close #lhFicDbf1
    Kill nom
    Call lect_fich
    Me.SSTab1.Tab = 0
    Me.SSTab1.TabEnabled(1) = False
    Me.SSTab1.TabEnabled(2) = False
    Me.SSTab1.TabEnabled(3) = False
    Me.SSTab1.TabEnabled(4) = False
    Me.SSTab1.TabEnabled(5) = False
    Me.Tb_titre.Text = ""
'    Me.Cmd_del.Visible = False
    Call ini_edessdo
    Call ini_form
    Call ini_debit(" ")
End If
End Sub

Private Sub OK_lignes_Click()
'------
'dessin dans ucg4
'Call init_graphdo(UC_graphique4)
'Call dess_troncon(UC_graphique4, edessdo.tron_amo, couleur.noir) ' vbBlack)
'Call dess_predo(Me.UC_graphique4, edo, couleur.noir)
'Call dess_cot(Me.UC_graphique4, couleur.noir) ' vbBlack)
'If Chk_piezo.Value = 1 Then
'    Call dess_piezo(UC_graphique4, edessdo.tron_amo, edessdo.Qts, couleur.magenta) ' vbGreen) ' RGB(128, 64, 128))
'    Call dess_piezo(UC_graphique4, edessdo.tron_amo, edessdo.Qrin, couleur.magenta_clair) ' vbRed)
'    Call dess_piezo(UC_graphique4, edessdo.tron_amo, edessdo.Qpluie, couleur.vert) ' vbMagenta)
'    Call dess_piezo(UC_graphique4, edo.tron_ava, edessdo.Qts, couleur.magenta) ' vbGreen)
'    Call dess_piezo(UC_graphique4, edo.tron_ava, edessdo.Qrin, couleur.magenta_clair) ' vbRed)
' '   Call dess_piezo(UC_graphique4, troava, edessdo.Qpluie, vbMagenta)
'End If
'If Chk_charge.Value = 1 Then
'    Call dess_charge(UC_graphique4, edo.tron_ava, edessdo.Qrin, couleur.orange) ' vbCyan)
'    Call dess_charge(UC_graphique4, edessdo.tron_amo, edessdo.Qrin, couleur.orange) ' vbCyan)
'End If
'If Chk_charge.Value = 1 Then
'End If

'dessin dans  owner.fdessin.UC_graphique1
Call init_graphdo(owner.fdessin.UC_graphique1)
Call dess_troncon(owner.fdessin.UC_graphique1, edessdo.tron_amo, couleur.noir) ' vbBlack)
Call dess_predo(owner.fdessin.UC_graphique1, edo, couleur.noir)
Call dess_cot(owner.fdessin.UC_graphique1, couleur.noir) ' vbBlack)
If Chk_piezo.Value = 1 Then
    Call dess_piezo(owner.fdessin.UC_graphique1, edessdo.tron_amo, edessdo.Qts, couleur.magenta) ' vbGreen) ' RGB(128, 64, 128))
    Call dess_piezo(owner.fdessin.UC_graphique1, edessdo.tron_amo, edessdo.Qrin, couleur.magenta_clair) ' vbRed)
    Call dess_piezo(owner.fdessin.UC_graphique1, edessdo.tron_amo, edessdo.Qpluie, couleur.vert) ' vbMagenta)
    Call dess_piezo(owner.fdessin.UC_graphique1, edo.tron_ava, edessdo.Qts, couleur.magenta) ' vbGreen)
    Call dess_piezo(owner.fdessin.UC_graphique1, edo.tron_ava, edessdo.Qrin, couleur.magenta_clair) ' vbRed)
 '   Call dess_piezo(owner.fdessin.UC_graphique1, troava, edessdo.Qpluie, vbMagenta)
End If
If Chk_charge.Value = 1 Then
    Call dess_charge(owner.fdessin.UC_graphique1, edo.tron_ava, edessdo.Qrin, couleur.orange) ' vbCyan)
    Call dess_charge(owner.fdessin.UC_graphique1, edessdo.tron_amo, edessdo.Qrin, couleur.orange) ' vbCyan)
End If
If Chk_charge.Value = 1 Then
End If

End Sub



Private Sub SSTab1_Click(PreviousTab As Integer)
Dim mes As String
Select Case SSTab1.Tab
    Case Is = 0
        mes = "DO:Bassin versant"
        owner.fdessin.Image1.Visible = False
         owner.fdessin.UC_graphiqueB.Visible = True
       owner.fdessin.UC_graphique1.Visible = False
        owner.fdessin.UC_graphique2.Visible = False
   Case Is = 1
   mes = "DO:Contraintes"
        owner.fdessin.Image1.Visible = False
         owner.fdessin.UC_graphiqueB.Visible = True
        owner.fdessin.UC_graphique1.Visible = False
        owner.fdessin.UC_graphique2.Visible = False
   Case Is = 2
   mes = "DO:Conduite amont"
        owner.fdessin.Image1.Visible = False
        owner.fdessin.UC_graphique1.Visible = True
         owner.fdessin.UC_graphiqueB.Visible = False
        owner.fdessin.UC_graphique2.Visible = False
   Case Is = 3
   mes = "DO:Conduite aval étranglée"
        owner.fdessin.Image1.Visible = False
         owner.fdessin.UC_graphiqueB.Visible = False
        owner.fdessin.UC_graphique1.Visible = True
        owner.fdessin.UC_graphique2.Visible = False
   Case Is = 4
   mes = "DO:Déversoir"
        owner.fdessin.Image1.Visible = False
         owner.fdessin.UC_graphiqueB.Visible = False
        owner.fdessin.UC_graphique1.Visible = True
        owner.fdessin.UC_graphique2.Visible = False
   Case Is = 5
   mes = "DO:Conduite de décharge"
        owner.fdessin.Image1.Visible = False
         owner.fdessin.UC_graphiqueB.Visible = False
        owner.fdessin.UC_graphique1.Visible = False
        owner.fdessin.UC_graphique2.Visible = True
End Select
owner.affich_aide Me.Name, mes

'Debug.Print PreviousTab
'Debug.Print SSTab1.Tab
End Sub



Private Sub tb_dech_Change(Index As Integer)
Select Case Index
    Case Is = 0
        edessdo.tron_dech.conduit.Diametre = txtVersNum(TB_dech(0).Text) / 1000#
    Case Is = 1
        edessdo.tron_dech.conduit.pente = txtVersNum(TB_dech(1).Text) / 10000
    Case Is = 2
        edessdo.tron_dech.conduit.rugosite = txtVersNum(TB_dech(2).Text)
    Case Is = 3
        edessdo.tron_dech.conduit.Longueur = txtVersNum(TB_dech(3).Text)
End Select
End Sub
Private Sub Tb_dech_KeyPress(Index As Integer, KeyAscii As Integer)
Dim reponse As Integer
If Len(TB_dech(Index).Text) < TB_dech(Index).MaxLength Then
   Select Case Index
        Case Is = 0
            KeyAscii = verif_car(KeyAscii, "Saisie diamétre canalisation de décharge", "I")
        Case Is = 1
            KeyAscii = verif_car(KeyAscii, "Saisie pente canalisation de décharge", "I")
        Case Is = 2
            KeyAscii = verif_car(KeyAscii, "Saisie coefficient canalisation de décharge", "I")
        Case Is = 3
            KeyAscii = verif_car(KeyAscii, "Saisie longueur canalisation de décharge", "I")
    End Select
End If
End Sub

Private Sub Tb_dev_Change(Index As Integer)
Select Case Index
    Case Is = 3
        edessdo.Tram = txtVersNum(Me.Tb_dev(3).Text)
End Select
Me.Cmd_resudo.Enabled = True
Call ini_longce
End Sub

Private Sub Tb_dev_KeyPress(Index As Integer, KeyAscii As Integer)
Dim reponse As Integer
If Len(Tb_dev(Index).Text) < Tb_dev(Index).MaxLength Then
   Select Case Index
        Case Is = 0
            KeyAscii = verif_car(KeyAscii, "Saisie longueur DO", "R")
        Case Is = 1
            KeyAscii = verif_car(KeyAscii, "Saisie hauteur de la crête", "R")
        Case Is = 2
            KeyAscii = verif_car(KeyAscii, "Saisie pente DO", "R")
        Case Is = 3
            KeyAscii = verif_car(KeyAscii, "Saisie tirant d'eau amont admissible", "R")
    End Select
End If
End Sub
Private Sub Tb_debit_Change(Index As Integer)
    Select Case Index
        Case Is = 0
            edessdo.Qpluie = txtVersNum(Me.Tb_debit(0).Text)
        Case Is = 1
            edessdo.Qts = txtVersNum(Me.Tb_debit(1).Text)
        Case Is = 2
            edessdo.Qrin = txtVersNum(Me.Tb_debit(2).Text)
    End Select
    Call reini_form(1)
    Call reini_valeurs
End Sub

Private Sub Tb_debit_KeyPress(Index As Integer, KeyAscii As Integer)
Dim reponse As Integer
If Len(Tb_debit(Index).Text) < Tb_debit(Index).MaxLength Then
   Select Case Index
        Case Is = 0
            KeyAscii = verif_car(KeyAscii, "Saisie débit d'eau pluviale", "R")
        Case Is = 1
            KeyAscii = verif_car(KeyAscii, "Saisie débit de temps sec", "R")
        Case Is = 2
            KeyAscii = verif_car(KeyAscii, "Saisie débit de rinçage", "R")
    End Select
End If
End Sub
Private Sub Tb_cont_Change(Index As Integer)
If Trim(Tb_cont(Index).Text) = "" Then
    Tb_cont(Index).Text = 0#
End If
Select Case Index
    Case Is = 0
        Me.Label8.Caption = Trim(Tb_cont(0).Text)
        edessdo.rdoam = txtVersNum(Me.Tb_cont(0).Text)
    Case Is = 1
        Me.Label9.Caption = Trim(Tb_cont(1).Text)
        edessdo.rdoav = txtVersNum(Me.Tb_cont(1).Text)
    Case Is = 2
        Me.Label13.Caption = Trim(Tb_cont(2).Text)
        edessdo.lgdisp = txtVersNum(Me.Tb_cont(2).Text)
    Case Is = 3
        Me.Label12.Caption = Trim(Tb_cont(3).Text)
        edessdo.phex = txtVersNum(Me.Tb_cont(3).Text)
    Case Is = 4
        Me.Label11.Caption = Trim(Tb_cont(4).Text)
        edessdo.rdoex = txtVersNum(Me.Tb_cont(4).Text)
    Case Is = 5
        Me.Label10.Caption = Trim(Tb_cont(5).Text)
        edessdo.lgca = txtVersNum(Me.Tb_cont(5).Text)
End Select
Call reini_form(2)
Call reini_valeurs
End Sub

Private Sub Tb_cont_KeyPress(Index As Integer, KeyAscii As Integer)
Dim reponse As Integer
If Len(Tb_cont(Index).Text) < Tb_cont(Index).MaxLength Then
   Select Case Index
        Case Is = 0
            KeyAscii = verif_car(KeyAscii, "Saisie cote radier amont", "R")
        Case Is = 1
            KeyAscii = verif_car(KeyAscii, "Saisie cote radier aval", "R")
        Case Is = 2
            KeyAscii = verif_car(KeyAscii, "Saisie longueur disponible", "R")
        Case Is = 3
            KeyAscii = verif_car(KeyAscii, "Saisie cote des PHE à l'exutoire", "R")
        Case Is = 4
            KeyAscii = verif_car(KeyAscii, "Saisie cote radier à l'exutoire", "R")
       Case Is = 5
            KeyAscii = verif_car(KeyAscii, "Saisie longueur de la canalisation", "R")
    End Select
End If
End Sub
Private Sub Tb_amo_Change(Index As Integer)
Select Case Index
    Case Is = 0
        edessdo.dam = txtVersNum(Me.Tb_amo(0).Text)
    Case Is = 1
        edessdo.iRadam = txtVersNum(Me.Tb_amo(1).Text)
    Case Is = 2
        edessdo.Kam = txtVersNum(Me.Tb_amo(2).Text)
    Case Is = 3
        edessdo.Lam = txtVersNum(Me.Tb_amo(3).Text)
End Select
    Call ini_resum
End Sub

Private Sub Tb_amo_KeyPress(Index As Integer, KeyAscii As Integer)
Dim reponse As Integer
If Len(Tb_amo(Index).Text) < Tb_amo(Index).MaxLength Then
   Select Case Index
        Case Is = 0
            KeyAscii = verif_car(KeyAscii, "Saisie diamétre canalisation amont", "I")
        Case Is = 1
            KeyAscii = verif_car(KeyAscii, "Saisie pente canalisation amont", "I")
        Case Is = 2
            KeyAscii = verif_car(KeyAscii, "Saisie coefficient canalisation amont", "I")
        Case Is = 3
            KeyAscii = verif_car(KeyAscii, "Saisie longueur canalisation amont", "I")
    End Select
End If
End Sub
Private Sub Tb_ava_Change(Index As Integer)
Select Case Index
    Case Is = 0
        edessdo.dav = txtVersNum(Me.Tb_ava(0).Text)
    Case Is = 1
        edessdo.iradav = txtVersNum(Me.Tb_ava(1).Text)
    Case Is = 2
        edessdo.kav = txtVersNum(Me.Tb_ava(2).Text)
    Case Is = 3
        edessdo.Lav = txtVersNum(Me.Tb_ava(3).Text)
End Select
    Call ini_resuv
End Sub

Private Sub Tb_ava_KeyPress(Index As Integer, KeyAscii As Integer)
Dim reponse As Integer
If Len(Tb_ava(Index).Text) < Tb_ava(Index).MaxLength Then
   Select Case Index
        Case Is = 0
    KeyAscii = verif_car(KeyAscii, "Saisie diamétre canalisation aval", "I")
        Case Is = 1
    KeyAscii = verif_car(KeyAscii, "Saisie pente canalisation aval", "I")
        Case Is = 2
    KeyAscii = verif_car(KeyAscii, "Saisie coefficient canalisation aval", "I")
        Case Is = 3
    KeyAscii = verif_car(KeyAscii, "Saisie longueur canalisation aval", "I")
    End Select
End If
End Sub

Public Sub Mquitter()
    MnuQuit_Click
End Sub
Public Sub Msupprimer()
    mnusuppr_Click
End Sub
Public Sub Menregistrer()
    MnuSave_Click
End Sub
Public Sub Init_ss_commentaire()
    owner.affich_com Me.Name, "DO"
    owner.affich_aide Me.Name, "Dimensionnement d'un déversoir d'orage"

End Sub


