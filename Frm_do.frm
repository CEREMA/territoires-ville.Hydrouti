VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Frm_do 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hydraulique Déversoir d'Orage à crête haute"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9825
   Icon            =   "Frm_do.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   9825
   Begin VB.ComboBox Cb_deversoir 
      Height          =   315
      Left            =   360
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   360
      Width           =   4000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   459
      TabCaption(0)   =   "Bassin Versant"
      TabPicture(0)   =   "Frm_do.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Lab_bas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Cmd_Sel_Bv"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frm_bv"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Contraintes"
      TabPicture(1)   =   "Frm_do.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Canal. Amont"
      TabPicture(2)   =   "Frm_do.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lb_vpsm"
      Tab(2).Control(1)=   "Lb_dpsm"
      Tab(2).Control(2)=   "Lb_vqtsm"
      Tab(2).Control(3)=   "Lb_hqtsm"
      Tab(2).Control(4)=   "Lb_vqrinm"
      Tab(2).Control(5)=   "Lb_hqrinm"
      Tab(2).Control(6)=   "Label1"
      Tab(2).Control(7)=   "Lb_vqpluiem"
      Tab(2).Control(8)=   "Lb_hqpluiem"
      Tab(2).Control(9)=   "Lb_mesm"
      Tab(2).Control(10)=   "Frm_condam"
      Tab(2).Control(11)=   "Cmd_resum"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Cmd_annulm"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).ControlCount=   13
      TabCaption(3)   =   "Canal. Aval"
      TabPicture(3)   =   "Frm_do.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Lb_mesv"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Lb_Hqpluiev"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Lb_Vqpluiev"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Lb_Hqrinv"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Lb_Vqrinv"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Lb_Hqtsv"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Lb_Vqtsv"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Lb_Dpsv"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Lb_Vpsv"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Frm_condav"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Cmd_resuv"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Cmd_Pvp"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Cmd_Pvm"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Cmd_annulv"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).ControlCount=   14
      TabCaption(4)   =   "Déversoir"
      TabPicture(4)   =   "Frm_do.frx":093A
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Lb_preste"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Lb_longce(0)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Lb_longce(1)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "SSTab_result"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Cmd_reinit"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Cmd_VerifDo"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Txtb_deversoir"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Frm_dev"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Cmd_resudo"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Frame4"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).ControlCount=   10
      TabCaption(5)   =   "Décharge"
      TabPicture(5)   =   "Frm_do.frx":0956
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Lb_vpsdech"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Lb_Dpsdech"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Lb_vqdev"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Lb_hqdev"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "lb_Qdev"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Frm_dech"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "Cmd_resudech"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "Cmd_verifdech"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "Txtb_decharge"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).ControlCount=   9
      Begin VB.Frame Frame4 
         Caption         =   "Dessin lignes"
         Height          =   972
         Left            =   4800
         TabIndex        =   116
         Top             =   3000
         Width           =   3495
         Begin VB.CheckBox Chk_Qpluie 
            Caption         =   "Qpluie"
            Height          =   312
            Left            =   1920
            TabIndex        =   154
            Top             =   600
            Width           =   924
         End
         Begin VB.CheckBox Chk_Qrin 
            Caption         =   "Qrin"
            Height          =   192
            Left            =   960
            TabIndex        =   153
            Top             =   600
            Width           =   876
         End
         Begin VB.CheckBox Chk_Qts 
            Caption         =   "Qts"
            Height          =   192
            Left            =   120
            TabIndex        =   152
            Top             =   600
            Width           =   732
         End
         Begin VB.CommandButton OK_lignes 
            Caption         =   "OK"
            Height          =   255
            Left            =   3000
            TabIndex        =   156
            Top             =   240
            Visible         =   0   'False
            Width           =   372
         End
         Begin VB.CheckBox Chk_eau 
            Caption         =   "Eau"
            Height          =   255
            Left            =   120
            TabIndex        =   115
            Top             =   240
            Width           =   732
         End
         Begin VB.CheckBox Chk_charge 
            Caption         =   "Charge"
            Height          =   255
            Left            =   1920
            TabIndex        =   118
            Top             =   240
            Width           =   852
         End
         Begin VB.CheckBox Chk_piezo 
            Caption         =   "Piezo"
            Height          =   255
            Left            =   960
            TabIndex        =   117
            Top             =   240
            Width           =   852
         End
      End
      Begin VB.TextBox Txtb_decharge 
         BackColor       =   &H80000016&
         Height          =   2295
         Left            =   -70560
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   720
         Width           =   5055
      End
      Begin VB.CommandButton Cmd_resudo 
         Caption         =   "Calculer"
         Height          =   255
         Left            =   3600
         TabIndex        =   95
         TabStop         =   0   'False
         ToolTipText     =   "Calcul du dimensionnement de la conduite étranglée"
         Top             =   3000
         Width           =   1000
      End
      Begin VB.Frame Frm_dev 
         Caption         =   "Caractéristiques"
         Height          =   2250
         Left            =   120
         TabIndex        =   88
         Top             =   630
         Width           =   3255
         Begin VB.TextBox Tb_dev 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   3
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   94
            Top             =   1680
            Width           =   750
         End
         Begin VB.TextBox Tb_dev 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   1920
            MaxLength       =   8
            TabIndex        =   93
            Top             =   1080
            Width           =   750
         End
         Begin VB.TextBox Tb_dev 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   92
            Top             =   720
            Width           =   750
         End
         Begin VB.TextBox Tb_dev 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   91
            Top             =   360
            Width           =   750
         End
         Begin VB.Label Lb_udev 
            Caption         =   "m"
            Height          =   255
            Index           =   3
            Left            =   2760
            TabIndex        =   137
            Top             =   1725
            Width           =   360
         End
         Begin VB.Label Lb_udev 
            Caption         =   "m/m"
            Height          =   255
            Index           =   2
            Left            =   2760
            TabIndex        =   136
            Top             =   1125
            Width           =   360
         End
         Begin VB.Label Lb_udev 
            Caption         =   "m"
            Height          =   255
            Index           =   1
            Left            =   2760
            TabIndex        =   135
            Top             =   765
            Width           =   360
         End
         Begin VB.Label Lb_udev 
            Caption         =   "m"
            Height          =   255
            Index           =   0
            Left            =   2760
            TabIndex        =   134
            Top             =   360
            Width           =   390
         End
         Begin VB.Label Lb_intdev 
            Caption         =   "Tirant d'eau amont admissible"
            Height          =   465
            Index           =   3
            Left            =   120
            TabIndex        =   114
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Lb_intdev 
            Caption         =   " Pente du DO"
            Height          =   300
            Index           =   2
            Left            =   120
            TabIndex        =   98
            Top             =   1130
            Width           =   1575
         End
         Begin VB.Label Lb_intdev 
            Caption         =   "Hauteur de la crête "
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   97
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Lb_intdev 
            Caption         =   "Longueur du DO"
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   96
            Top             =   410
            Width           =   1575
         End
      End
      Begin VB.TextBox Txtb_deversoir 
         BackColor       =   &H80000016&
         Height          =   2175
         Left            =   6000
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   146
         TabStop         =   0   'False
         Top             =   720
         Width           =   3492
      End
      Begin VB.CommandButton Cmd_VerifDo 
         Caption         =   "Vérifier"
         Height          =   255
         Left            =   8520
         TabIndex        =   144
         TabStop         =   0   'False
         ToolTipText     =   "Vérification et calcul  à débit maxi "
         Top             =   3000
         Width           =   1000
      End
      Begin VB.CommandButton Cmd_reinit 
         Caption         =   "Ré-initialiser"
         Height          =   255
         Left            =   360
         TabIndex        =   142
         TabStop         =   0   'False
         ToolTipText     =   "Réinitialisation des valeurs"
         Top             =   3000
         Width           =   1000
      End
      Begin VB.CommandButton Cmd_annulv 
         Caption         =   "Annuler"
         Height          =   255
         Left            =   -72840
         TabIndex        =   113
         TabStop         =   0   'False
         ToolTipText     =   "Retour aux valeurs précédentes"
         Top             =   3240
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton Cmd_annulm 
         Caption         =   "Annuler"
         Height          =   255
         Left            =   -72840
         TabIndex        =   112
         TabStop         =   0   'False
         ToolTipText     =   "Retour aux valeurs précédentes"
         Top             =   3240
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton Cmd_Pvm 
         Caption         =   "-"
         Height          =   195
         Left            =   -70560
         TabIndex        =   110
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Cmd_Pvp 
         Caption         =   "+"
         Height          =   195
         Left            =   -70560
         TabIndex        =   109
         Top             =   1440
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Cmd_verifdech 
         Caption         =   "Vérifier"
         Height          =   255
         Left            =   -69360
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   3140
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.CommandButton Cmd_resudech 
         Caption         =   "Calculer"
         Height          =   255
         Left            =   -70560
         TabIndex        =   165
         TabStop         =   0   'False
         ToolTipText     =   "Calcul et vérification du DO à Qmax"
         Top             =   3140
         Width           =   1000
      End
      Begin VB.Frame Frm_dech 
         Caption         =   "Conduite"
         Height          =   2895
         Left            =   -74880
         TabIndex        =   99
         Top             =   630
         Width           =   4215
         Begin VB.CommandButton Cmd_dech 
            Caption         =   "Courbe..."
            Height          =   255
            Left            =   3120
            TabIndex        =   168
            TabStop         =   0   'False
            ToolTipText     =   "Courbe de débit de la conduite de décharge"
            Top             =   2520
            Width           =   990
         End
         Begin VB.ComboBox Cb_centon 
            CausesValidation=   0   'False
            Height          =   315
            ItemData        =   "Frm_do.frx":0972
            Left            =   2400
            List            =   "Frm_do.frx":0974
            Sorted          =   -1  'True
            TabIndex        =   164
            Top             =   2040
            Width           =   900
         End
         Begin VB.TextBox Tb_dech 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   4
            Left            =   2400
            MaxLength       =   6
            TabIndex        =   155
            Top             =   2040
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.TextBox Tb_dech 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   3
            Left            =   2400
            MaxLength       =   8
            TabIndex        =   163
            Top             =   1440
            Width           =   900
         End
         Begin VB.TextBox Tb_dech 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   162
            Top             =   1080
            Width           =   900
         End
         Begin VB.TextBox Tb_dech 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   161
            Top             =   720
            Width           =   900
         End
         Begin VB.TextBox Tb_dech 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   2400
            MaxLength       =   6
            TabIndex        =   160
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Label16 
            Caption         =   "Chambre d'entonnement"
            Height          =   255
            Left            =   120
            TabIndex        =   159
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label Lb_udech 
            Height          =   195
            Index           =   4
            Left            =   3360
            TabIndex        =   158
            Top             =   2085
            Width           =   255
         End
         Begin VB.Label Lb_intdech 
            Caption         =   "Coefficient de perte à l'entrée"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   157
            Top             =   2090
            Width           =   2175
         End
         Begin VB.Label Lb_udech 
            Height          =   255
            Index           =   2
            Left            =   3360
            TabIndex        =   151
            Top             =   1125
            Width           =   495
         End
         Begin VB.Label Lb_udech 
            Caption         =   "m"
            Height          =   255
            Index           =   3
            Left            =   3360
            TabIndex        =   140
            Top             =   1485
            Width           =   735
         End
         Begin VB.Label Lb_udech 
            Caption         =   "1/10000"
            Height          =   255
            Index           =   1
            Left            =   3360
            TabIndex        =   139
            Top             =   765
            Width           =   735
         End
         Begin VB.Label Lb_udech 
            Caption         =   "mm"
            Height          =   255
            Index           =   0
            Left            =   3360
            TabIndex        =   138
            Top             =   405
            Width           =   735
         End
         Begin VB.Label Lb_intdech 
            Caption         =   "Longueur  canalisation "
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   107
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Label Lb_intdech 
            Caption         =   "Coef. Manning-Strickler"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   102
            Top             =   1125
            Width           =   2175
         End
         Begin VB.Label Lb_intdech 
            Caption         =   "Pente "
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   101
            Top             =   765
            Width           =   2175
         End
         Begin VB.Label Lb_intdech 
            Caption         =   "Diamètre "
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   100
            Top             =   410
            Width           =   2175
         End
      End
      Begin VB.CommandButton Cmd_resuv 
         Caption         =   "Calculer"
         Height          =   255
         Left            =   -71640
         TabIndex        =   76
         TabStop         =   0   'False
         ToolTipText     =   "Contrôle de la conduite aval"
         Top             =   3240
         Width           =   1000
      End
      Begin VB.Frame Frm_condav 
         Caption         =   "Conduite"
         Height          =   2295
         Left            =   -74760
         TabIndex        =   68
         Top             =   840
         Width           =   4095
         Begin VB.CommandButton Cmd_ava 
            Caption         =   "Courbe..."
            Height          =   255
            Left            =   3000
            TabIndex        =   167
            TabStop         =   0   'False
            ToolTipText     =   "Courbe de débit de la conduite aval"
            Top             =   1920
            Width           =   990
         End
         Begin VB.TextBox Tb_ava 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   2160
            MaxLength       =   6
            TabIndex        =   69
            Top             =   360
            Width           =   900
         End
         Begin VB.TextBox Tb_ava 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   2160
            MaxLength       =   4
            TabIndex        =   71
            Top             =   720
            Width           =   900
         End
         Begin VB.TextBox Tb_ava 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   2160
            MaxLength       =   4
            TabIndex        =   73
            Top             =   1080
            Width           =   900
         End
         Begin VB.TextBox Tb_ava 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   3
            Left            =   2160
            MaxLength       =   8
            TabIndex        =   75
            Top             =   1440
            Width           =   900
         End
         Begin VB.Label Lb_uava 
            Height          =   255
            Index           =   2
            Left            =   3240
            TabIndex        =   150
            Top             =   1125
            Width           =   615
         End
         Begin VB.Label Lb_uava 
            Caption         =   "m"
            Height          =   255
            Index           =   3
            Left            =   3240
            TabIndex        =   133
            Top             =   1485
            Width           =   735
         End
         Begin VB.Label Lb_uava 
            Caption         =   "1/10000"
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   132
            Top             =   765
            Width           =   735
         End
         Begin VB.Label Lb_uava 
            Caption         =   "mm"
            Height          =   255
            Index           =   0
            Left            =   3240
            TabIndex        =   131
            Top             =   405
            Width           =   735
         End
         Begin VB.Label Lb_intava 
            Caption         =   "Diamètre "
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   77
            Top             =   410
            Width           =   1815
         End
         Begin VB.Label Lb_intava 
            Caption         =   "Pente "
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   74
            Top             =   770
            Width           =   1815
         End
         Begin VB.Label Lb_intava 
            Caption         =   "Coeff. de  Strickler"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   72
            Top             =   1125
            Width           =   1815
         End
         Begin VB.Label Lb_intava 
            Caption         =   "Longueur canalisation "
            Height          =   300
            Index           =   3
            Left            =   120
            TabIndex        =   70
            Top             =   1490
            Width           =   1815
         End
      End
      Begin VB.CommandButton Cmd_resum 
         Caption         =   "Calculer"
         Height          =   255
         Left            =   -71640
         TabIndex        =   58
         TabStop         =   0   'False
         ToolTipText     =   "Contrôle  de la conduite amont"
         Top             =   3240
         Width           =   1000
      End
      Begin VB.Frame Frm_condam 
         Caption         =   "Conduite"
         Height          =   2295
         Left            =   -74760
         TabIndex        =   49
         Top             =   840
         Width           =   4095
         Begin VB.CommandButton Cmd_amo 
            Caption         =   "Courbe..."
            Height          =   255
            Left            =   3000
            TabIndex        =   166
            TabStop         =   0   'False
            ToolTipText     =   "Courbe de débit de la conduite amont"
            Top             =   1920
            Width           =   990
         End
         Begin VB.TextBox Tb_amo 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   3
            Left            =   2160
            MaxLength       =   8
            TabIndex        =   57
            Top             =   1440
            Width           =   900
         End
         Begin VB.TextBox Tb_amo 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   2160
            MaxLength       =   4
            TabIndex        =   56
            Top             =   1080
            Width           =   900
         End
         Begin VB.TextBox Tb_amo 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   2160
            MaxLength       =   4
            TabIndex        =   55
            Top             =   720
            Width           =   900
         End
         Begin VB.TextBox Tb_amo 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   2160
            MaxLength       =   6
            TabIndex        =   54
            Top             =   360
            Width           =   900
         End
         Begin VB.Label Lb_uamo 
            Height          =   255
            Index           =   2
            Left            =   3240
            TabIndex        =   149
            Top             =   1125
            Width           =   495
         End
         Begin VB.Label Lb_uamo 
            Caption         =   "m"
            Height          =   255
            Index           =   3
            Left            =   3240
            TabIndex        =   130
            Top             =   1485
            Width           =   735
         End
         Begin VB.Label Lb_uamo 
            Caption         =   "1/10000"
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   129
            Top             =   765
            Width           =   735
         End
         Begin VB.Label Lb_uamo 
            Caption         =   "mm"
            Height          =   255
            Index           =   0
            Left            =   3240
            TabIndex        =   128
            Top             =   405
            Width           =   735
         End
         Begin VB.Label Lb_intamo 
            Caption         =   "Longueur  canalisation "
            Height          =   300
            Index           =   3
            Left            =   120
            TabIndex        =   53
            Top             =   1490
            Width           =   1815
         End
         Begin VB.Label Lb_intamo 
            Caption         =   "Coeff.  de  Strickler"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   52
            Top             =   1125
            Width           =   1815
         End
         Begin VB.Label Lb_intamo 
            Caption         =   "Pente "
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   51
            Top             =   770
            Width           =   1815
         End
         Begin VB.Label Lb_intamo 
            Caption         =   "Diamètre "
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   405
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00000080&
         Height          =   3495
         Left            =   -70680
         TabIndex        =   22
         Top             =   360
         Width           =   5265
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
            Left            =   1380
            Top             =   1785
            Width           =   240
         End
         Begin VB.Shape Forme1 
            Height          =   330
            Index           =   4
            Left            =   2460
            Top             =   1695
            Width           =   600
         End
         Begin VB.Line Line1 
            X1              =   1605
            X2              =   2460
            Y1              =   1830
            Y2              =   1830
         End
         Begin VB.Line Line2 
            X1              =   1605
            X2              =   2460
            Y1              =   1965
            Y2              =   1965
         End
         Begin VB.Shape Forme1 
            BorderColor     =   &H000000C0&
            Height          =   195
            Index           =   5
            Left            =   3810
            Top             =   1850
            Width           =   240
         End
         Begin VB.Line Line3 
            BorderColor     =   &H000000C0&
            X1              =   3045
            X2              =   3810
            Y1              =   2000
            Y2              =   2000
         End
         Begin VB.Line Line4 
            BorderColor     =   &H000000C0&
            X1              =   3045
            X2              =   3810
            Y1              =   1945
            Y2              =   1945
         End
         Begin VB.Line Line5 
            BorderColor     =   &H0000FF00&
            X1              =   2550
            X2              =   3360
            Y1              =   1695
            Y2              =   720
         End
         Begin VB.Line Line6 
            BorderColor     =   &H0000FF00&
            X1              =   2685
            X2              =   3480
            Y1              =   1695
            Y2              =   720
         End
         Begin VB.Line Line7 
            X1              =   1380
            X2              =   1155
            Y1              =   1830
            Y2              =   1830
         End
         Begin VB.Line Line8 
            X1              =   1380
            X2              =   1155
            Y1              =   1965
            Y2              =   1965
         End
         Begin VB.Line Line9 
            BorderColor     =   &H000000C0&
            X1              =   4035
            X2              =   4440
            Y1              =   1965
            Y2              =   1965
         End
         Begin VB.Line Line10 
            BorderColor     =   &H000000C0&
            X1              =   4035
            X2              =   4440
            Y1              =   2020
            Y2              =   2020
         End
         Begin VB.Line Line11 
            X1              =   2460
            X2              =   3045
            Y1              =   1800
            Y2              =   1910
         End
         Begin VB.Line Line12 
            BorderStyle     =   3  'Dot
            X1              =   1470
            X2              =   1470
            Y1              =   2100
            Y2              =   2520
         End
         Begin VB.Line Line13 
            BorderStyle     =   3  'Dot
            X1              =   3930
            X2              =   3930
            Y1              =   2100
            Y2              =   2520
         End
         Begin VB.Line Line14 
            X1              =   1470
            X2              =   3960
            Y1              =   2550
            Y2              =   2550
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Longueur disponible"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   1920
            TabIndex        =   47
            Top             =   2325
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
            Left            =   1965
            TabIndex        =   46
            Top             =   1560
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
            Left            =   135
            TabIndex        =   45
            Top             =   2850
            Width           =   195
         End
         Begin VB.Label Etiquette30 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Canalisation amont unitaire"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   4
            Left            =   360
            TabIndex        =   44
            Top             =   2880
            Width           =   1920
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
            Left            =   2640
            TabIndex        =   43
            Top             =   2010
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
            Left            =   3405
            TabIndex        =   42
            Top             =   1680
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
            Left            =   2775
            TabIndex        =   41
            Top             =   930
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
            Left            =   135
            TabIndex        =   40
            Top             =   3120
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
            Left            =   2310
            TabIndex        =   39
            Top             =   2880
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
            Left            =   1920
            TabIndex        =   38
            Top             =   3120
            Width           =   195
         End
         Begin VB.Label Etiquette30 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Déversoir d'orage"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   5
            Left            =   360
            TabIndex        =   37
            Top             =   3120
            Width           =   1365
         End
         Begin VB.Label Etiquette30 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Canalisation aval eaux usées"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   6
            Left            =   2535
            TabIndex        =   36
            Top             =   2880
            Width           =   2715
         End
         Begin VB.Label Etiquette30 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Canalisation de décharge eaux pluviales"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   7
            Left            =   2160
            TabIndex        =   35
            Top             =   3120
            Width           =   2880
         End
         Begin VB.Label Etiquette31 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cote de radier obligé aval"
            ForeColor       =   &H000000C0&
            Height          =   465
            Index           =   2
            Left            =   4215
            TabIndex        =   34
            Top             =   1515
            Width           =   1005
         End
         Begin VB.Label Etiquette31 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cote de radier obligé amont"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   210
            TabIndex        =   33
            Top             =   1425
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
            Left            =   3360
            TabIndex        =   31
            Top             =   885
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
            Left            =   300
            TabIndex        =   28
            Top             =   1920
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
            Left            =   4440
            TabIndex        =   27
            Top             =   2040
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
            Left            =   3540
            TabIndex        =   26
            Top             =   1110
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
            Left            =   2325
            TabIndex        =   23
            Top             =   2595
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
            TabIndex        =   127
            Top             =   2210
            Width           =   200
         End
         Begin VB.Label Lb_ucont 
            Caption         =   "m"
            Height          =   255
            Index           =   4
            Left            =   3600
            TabIndex        =   126
            Top             =   1850
            Width           =   200
         End
         Begin VB.Label Lb_ucont 
            Caption         =   "m"
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   125
            Top             =   1490
            Width           =   200
         End
         Begin VB.Label Lb_ucont 
            Caption         =   "m"
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   124
            Top             =   1130
            Width           =   200
         End
         Begin VB.Label Lb_ucont 
            Caption         =   "m"
            Height          =   255
            Index           =   1
            Left            =   3600
            TabIndex        =   123
            Top             =   770
            Width           =   200
         End
         Begin VB.Label Lb_ucont 
            Caption         =   "m"
            Height          =   255
            Index           =   0
            Left            =   3600
            TabIndex        =   122
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
            Caption         =   "Longueur de la canalisation de décharge"
            Height          =   420
            Index           =   5
            Left            =   195
            TabIndex        =   20
            Top             =   2205
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
         Left            =   -70200
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
            TabIndex        =   121
            Top             =   1365
            Width           =   405
         End
         Begin VB.Label Lb_udebit 
            Caption         =   "l/s"
            Height          =   255
            Index           =   1
            Left            =   3960
            TabIndex        =   120
            Top             =   1005
            Width           =   405
         End
         Begin VB.Label Lb_udebit 
            Caption         =   "l/s"
            Height          =   255
            Index           =   0
            Left            =   3960
            TabIndex        =   119
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
            Caption         =   "Débit de référence "
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
         Height          =   255
         Left            =   -74520
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1200
         Width           =   3855
      End
      Begin TabDlg.SSTab SSTab_result 
         Height          =   1815
         Left            =   2760
         TabIndex        =   141
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
         Tab(0).ControlEnabled=   0   'False
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "Vérification"
         TabPicture(1)   =   "Frm_do.frx":0976
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).ControlCount=   0
      End
      Begin VB.Label Lab_bas 
         Height          =   1335
         Left            =   -74520
         TabIndex        =   169
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label lb_Qdev 
         Caption         =   "Qdev"
         Height          =   225
         Left            =   -69720
         TabIndex        =   147
         Top             =   1680
         Width           =   4260
      End
      Begin VB.Label Lb_longce 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Height          =   1335
         Index           =   1
         Left            =   7080
         TabIndex        =   145
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Lb_longce 
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Height          =   2175
         Index           =   0
         Left            =   3480
         TabIndex        =   143
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Lb_Vpsv 
         Caption         =   "Vps"
         Height          =   225
         Left            =   -70155
         TabIndex        =   85
         Top             =   840
         Width           =   4500
      End
      Begin VB.Label Lb_Dpsv 
         Caption         =   "Dps"
         Height          =   225
         Left            =   -70155
         TabIndex        =   84
         Top             =   1080
         Width           =   4500
      End
      Begin VB.Label Lb_Vqtsv 
         Caption         =   "Vqts"
         Height          =   225
         Left            =   -70155
         TabIndex        =   83
         Top             =   1440
         Width           =   4500
      End
      Begin VB.Label Lb_Hqtsv 
         Caption         =   "Hqts"
         Height          =   225
         Left            =   -70155
         TabIndex        =   82
         Top             =   1680
         Width           =   4500
      End
      Begin VB.Label Lb_Vqrinv 
         Caption         =   "Vqref"
         Height          =   225
         Left            =   -70155
         TabIndex        =   81
         Top             =   2040
         Width           =   4500
      End
      Begin VB.Label Lb_Hqrinv 
         Caption         =   "Hqref"
         Height          =   225
         Left            =   -70155
         TabIndex        =   80
         Top             =   2280
         Width           =   4500
      End
      Begin VB.Label Lb_Vqpluiev 
         Caption         =   "Vqora"
         Height          =   225
         Left            =   -70155
         TabIndex        =   79
         Top             =   2640
         Width           =   4500
      End
      Begin VB.Label Lb_Hqpluiev 
         Caption         =   "Hqora"
         Height          =   225
         Left            =   -70155
         TabIndex        =   78
         Top             =   2880
         Width           =   4500
      End
      Begin VB.Label Lb_preste 
         Caption         =   " "
         Height          =   252
         Left            =   4320
         TabIndex        =   111
         Top             =   480
         Width           =   4068
      End
      Begin VB.Label Lb_hqdev 
         Caption         =   "Hqdev"
         Height          =   225
         Left            =   -70320
         TabIndex        =   106
         Top             =   2160
         Width           =   4395
      End
      Begin VB.Label Lb_vqdev 
         Caption         =   "Vqdev"
         Height          =   225
         Left            =   -69720
         TabIndex        =   105
         Top             =   1920
         Width           =   4260
      End
      Begin VB.Label Lb_Dpsdech 
         Caption         =   "Dps"
         Height          =   225
         Left            =   -70320
         TabIndex        =   104
         Top             =   1200
         Width           =   4500
      End
      Begin VB.Label Lb_vpsdech 
         Caption         =   "Vps"
         Height          =   225
         Left            =   -70440
         TabIndex        =   103
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
         Caption         =   "Hqora"
         Height          =   225
         Left            =   -70155
         TabIndex        =   67
         Top             =   2880
         Width           =   4395
      End
      Begin VB.Label Lb_vqpluiem 
         Caption         =   "Vqora"
         Height          =   225
         Left            =   -70155
         TabIndex        =   66
         Top             =   2640
         Width           =   4395
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   135
         Left            =   -74520
         TabIndex        =   65
         Top             =   5760
         Width           =   15
      End
      Begin VB.Label Lb_hqrinm 
         Caption         =   "Hqref"
         Height          =   225
         Left            =   -70155
         TabIndex        =   64
         Top             =   2280
         Width           =   4395
      End
      Begin VB.Label Lb_vqrinm 
         Caption         =   "Vqref"
         Height          =   225
         Left            =   -70155
         TabIndex        =   63
         Top             =   2040
         Width           =   4395
      End
      Begin VB.Label Lb_hqtsm 
         Caption         =   "Hqts"
         Height          =   225
         Left            =   -70155
         TabIndex        =   62
         Top             =   1680
         Width           =   4395
      End
      Begin VB.Label Lb_vqtsm 
         Caption         =   "Vqts"
         Height          =   225
         Left            =   -70155
         TabIndex        =   61
         Top             =   1440
         Width           =   4395
      End
      Begin VB.Label Lb_dpsm 
         Caption         =   "Dps"
         Height          =   300
         Left            =   -70155
         TabIndex        =   60
         Top             =   1080
         Width           =   4395
      End
      Begin VB.Label Lb_vpsm 
         Caption         =   "Vps"
         Height          =   300
         Left            =   -70155
         TabIndex        =   59
         Top             =   840
         Width           =   4395
      End
   End
   Begin VB.TextBox Tb_titre 
      Height          =   300
      Left            =   5520
      MaxLength       =   30
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.Menu mnufichier 
      Caption         =   "&Déversoir"
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
      End
      Begin VB.Menu mnusaves 
         Caption         =   "En&registrer sous..."
      End
      Begin VB.Menu mnusuppr 
         Caption         =   "&Supprimer..."
      End
      Begin VB.Menu f2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuprint 
         Caption         =   "Im&primer..."
      End
      Begin VB.Menu f3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quitter module"
      End
   End
End
Attribute VB_Name = "Frm_do"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private okg As Boolean
Private owner As MDIFrm_menu
Private ok_imp As Boolean
Private esave As st_savdo
Public nom_ouvrage As String
'Private nom_fich As String
Public nom_type As String
Private nom_dessin As String
Private lhFicDbf As Long
Private FileLength As Integer
Private nombassin As String
Private list_don1() As Variant
Private list_don2() As Variant
Private list_don3() As Variant
Private list_don4() As Variant
Private list_don5() As Variant
Private list_don6() As Variant
Private list_int1() As Variant
Private list_resu1() As Variant
Private Resuintdev As resu_intdev
Private Resuudev As resu_udev
Private Resup_do As Resudo
Private ok_tab1 As Boolean
Private ok_tab2 As Boolean
Private ok_tab3 As Boolean
Private ok_tab4 As Boolean
Private ok_tab5 As Boolean
Private centon_texte As String
Private dev_texte As String
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
Private Sub sel_text(tb_objet As TextBox)
    tb_objet.SelStart = 0
    
    tb_objet.SelLength = Len(tb_objet.Text)


End Sub
Private Sub Change_Couleur(nom As String, Index As Integer)
'Dim coul As ColorConstants, coulp As ColorConstants
'Dim Index1 As Integer
'Dim nom1 As String
'coulp = vbBlack
'coul = Couleur_Change
'nom1 = nom
'Select Case nom
'    Case Is = "Tb_debit"
'         nom1 = "Lb_intdebit"
'    Case Is = "Tb_cont"
'         nom1 = "Lb_intcont"
'    Case Is = "Tb_amo"
'         nom1 = "Lb_intamo"
'    Case Is = "Tb_ava"
'         nom1 = "Lb_intava"
'    Case Is = "Tb_dev"
'         nom1 = "Lb_intdev"
'    Case Is = "Tb_dech"
'         nom1 = "Lb_intdech"
'End Select
'Select Case label_prec
'    Case Is = "Lb_intdebit"
'         Lb_intdebit(index_prec).ForeColor = coulp
'    Case Is = "Lb_intcont"
'         Lb_intcont(index_prec).ForeColor = coulp
'    Case Is = "Lb_intamo"
'         Lb_intamo(index_prec).ForeColor = coulp
'    Case Is = "Lb_intava"
'         Lb_intava(index_prec).ForeColor = coulp
'    Case Is = "Lb_intdev"
'         Lb_intdev(index_prec).ForeColor = coulp
'    Case Is = "Lb_intdech"
'         Lb_intdech(index_prec).ForeColor = coulp
'    Case Is = "Frm_condam"
'         Frm_condam.ForeColor = coulp
'    Case Is = "Frm_condav"
'         Frm_condav.ForeColor = coulp
'    Case Is = "Frm_dev"
'         Frm_dev.ForeColor = coulp
'    Case Is = "Frm_dech"
'         Frm_dech.ForeColor = coulp
'    Case Is = "Frm_bv"
'         Frm_bv.ForeColor = coulp
'End Select
'Select Case nom1
'    Case Is = "Me"
'         Me.SetFocus
'    Case Is = "Lb_intdebit"
'         Lb_intdebit(Index).ForeColor = coul
''         Tb_debit(Index).SetFocus
'    Case Is = "Lb_intcont"
'         Lb_intcont(Index).ForeColor = coul
''         Tb_cont(Index).SetFocus
'    Case Is = "Lb_intamo"
'         Lb_intamo(Index).ForeColor = coul
''         Tb_amo(Index).SetFocus
'    Case Is = "Lb_intava"
'         Lb_intava(Index).ForeColor = coul
''         Tb_ava(Index).SetFocus
'    Case Is = "Lb_intdev"
'         Lb_intdev(Index).ForeColor = coul
''         Tb_dev(Index).SetFocus
'    Case Is = "Lb_intdech"
'         Lb_intdech(Index).ForeColor = coul
''        Select Case Index
''            Case Is = 4
''                Cb_centon.SetFocus
''            Case Is = 0, 1, 2, 3
''                Tb_dech(Index).SetFocus
''        End Select
'    Case Is = "Frm_bv"
''         Tb_debit(0).SetFocus
''         DoEvents
''         Lb_intdebit(0).ForeColor = coulp
'         Frm_bv.ForeColor = coul
'    Case Is = "Frm_condam"
''         Tb_amo(0).SetFocus
''         DoEvents
''         Lb_intamo(0).ForeColor = coulp
'         Frm_condam.ForeColor = coul
'    Case Is = "Frm_condav"
''         Tb_ava(0).SetFocus
''         DoEvents
''         Lb_intava(0).ForeColor = coulp
'         Frm_condav.ForeColor = coul
'    Case Is = "Frm_dev"
' '        Tb_dev(0).SetFocus
''         DoEvents
''         Lb_intdev(0).ForeColor = coulp
'         Frm_dev.ForeColor = coul
'    Case Is = "Frm_dech"
' '        Tb_dech(0).SetFocus
''         DoEvents
''         Lb_intdech(0).ForeColor = coulp
'         Frm_dech.ForeColor = coul
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
    Case Is = "Lb_intdebit"
         Tb_debit(Index).SetFocus
    Case Is = "Lb_intcont"
         Tb_cont(Index).SetFocus
    Case Is = "Lb_intamo"
         Tb_amo(Index).SetFocus
    Case Is = "Lb_intava"
         Tb_ava(Index).SetFocus
    Case Is = "Lb_intdev"
         Tb_dev(Index).SetFocus
    Case Is = "Lb_intdech"
        Select Case Index
            Case Is = 4
                Cb_centon.SetFocus
            Case Is = 0, 1, 2, 3
                Tb_dech(Index).SetFocus
        End Select
    Case Is = "Frm_bv"
         Tb_debit(0).SetFocus
    Case Is = "Frm_condam"
         Tb_amo(0).SetFocus
    Case Is = "Frm_condav"
         Tb_ava(0).SetFocus
    Case Is = "Frm_dev"
         Tb_dev(0).SetFocus
    Case Is = "Frm_dech"
         Tb_dech(0).SetFocus
    Case Is = "Cmd_Sel_Bv"
         Cmd_Sel_Bv.SetFocus
        
End Select
End Sub
Private Function Rec_Mes(nom As String, Index As Integer)
Dim mes As String
mes = ""
Select Case nom
    Case Is = "Lb_intdebit", "Tb_debit", "Frm_bv", "Cmd_Sel_Bv"
                mes = IDhlp_DODonneesBase  '"Données de base"
    Case Is = "Lb_intcont", "Tb_cont"
                mes = IDhlp_DOContraintes  '"Déversoir à seuil haut"
    Case Is = "Lb_intamo", "Tb_amo", "Frm_condam"
                mes = IDhlp_DOConduiteAmenee  '"Conduite d'amenée"
    Case Is = "Lb_intava", "Tb_ava", "Frm_condav"
                mes = IDhlp_DOConduiteDebitConserve  '"Conduite de débit conservé"
    Case Is = "Lb_intdev", "Tb_dev", "Frm_dev"
                mes = IDhlp_DOChambreDeversement  '"Chambre de déversement"
    Case Is = "Lb_intdech", "Tb_dech", "Frm_dech"
                mes = IDhlp_DOConduiteDecharge  '"Conduite de décharge"
End Select
mes_prec = mes
Rec_Mes = mes
End Function

Public Function get_l_tb() As Variant
get_l_tb = list_tb
End Function
Private Sub init_l_tab()
Dim l0() As Variant, l1() As Variant, l2() As Variant
Dim l3() As Variant, l4() As Variant, l5() As Variant
Dim l6() As Variant, l7() As Variant
l0 = Array(0)
l1 = Array(0, "TB_debit")
l2 = Array(0, "TB_cont")
l3 = Array(0, "TB_amo") ', "CMD_resum")
l4 = Array(0, "TB_ava") ', "CMD_resuv")
l5 = Array(0, "TB_dev") ', "CMD_resudo")
l6 = Array(0, "CHK_eau", "CHK_piezo", "CHK_charge", "CHK_qts", "CHK_qrin", "CHK_qpluie", "OK_lignes")
l7 = Array(0, "TB_dech", "CB_centon") ', "CMD_resudech")
ReDim list_tb(0 To UBound(l0), 0 To UBound(l1), 0 To UBound(l2), 0 To UBound(l3), 0 To UBound(l4), 0 To UBound(l5), 0 To UBound(l6), 0 To UBound(l7))
list_tb = Array(l0, l1, l2, l3, l4, l5, l6, l7)
End Sub
Public Sub retailler()
retaille

End Sub
Private Sub retaille()
    Me.Left = owner.fcom.Width + owner.fcom.Left
    Me.Top = 0
'    Me.Width = owner.Width - owner.fcom.Width - 200
'    Me.Height = owner.fdessin.Top
    Me.Width = maximum(larg_mini, owner.Width - owner.fcom.Width - owner.fcom.Left - l_decal_asc) ' 10040
    Me.Height = maximum(haut_mini, owner.fdessin.Top) '4600
End Sub
Private Sub Cb_centon_Change()
    centon_texte = Cb_centon.Text
    Tb_dech(4).Text = Cb_centon.Text
End Sub

Private Sub Cb_centon_Click()
Dim mes As String
Dim nom As String
nom = "Tb_dech"
mes = Rec_Mes(nom, 4)
Change_Couleur nom, 4
owner.affich_aide Me.Name, mes
    bKP = True
    Tb_dech(4).Text = Cb_centon.Text
End Sub

Private Sub Cb_centon_GotFocus()
Dim mes As String
Dim nom As String
nom = "Tb_dech"
If change_coul Then
    Change_Couleur nom, 4
    mes = Rec_Mes(nom, 4)
    owner.affich_aide Me.Name, mes
End If

End Sub

Private Sub Cb_centon_KeyDown(KeyCode As Integer, Shift As Integer)
    centon_texte = Cb_centon.Text
    Cb_centon.Text = centon_texte

End Sub

Private Sub Cb_centon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    centon_texte = Cb_centon.Text
End If
End Sub
Private Sub Cb_deversoir_Change()
    Cb_deversoir.Text = dev_texte
End Sub

Public Sub Cb_deversoir_click()
Dim za As st_savdo
Dim za1 As st_savdo1
Dim sresult As String
Call funlockb
 
    Me.Tb_titre.Text = Trim(nom_ouvrage)
    dev_texte = Trim(nom_ouvrage)
    Cb_deversoir.Text = Trim(nom_ouvrage)
    lhFicDbf = FreeFile
    On Error GoTo test_Error
    Open fich_lect For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
    If Not EOF(lhFicDbf) Then
        za = za1.stsavdo
        If Trim(za.type) = nom_type And Trim(za.nom) = Trim(Cb_deversoir.Text) Then
            Tb_titre = Trim(za.nom)
            Me.Caption = fen_titre + " : " + Tb_titre.Text
            edo = za.edo
            edessdo = za.edessdo
            nombassin = edessdo.nombv
            Call ini_bv
            If Trim(nombassin) <> "" Then
                Close #lhFicDbf
                If Not rec_bassin(nombassin, "versant") Then
                End If
                
                
            End If
            Call init_graphique
    
          If Trim(ebv.Qchoisi) <> "" Then
           Me.Frm_bv.Caption = "Hydraulique B.V. : " + Trim(nombassin) 'Trim(ebv.nom)
'           owner.fdessin.UC_graphiqueB.ecr_texta 1665, 1140, "Surface = " + ajout_zero(Trim(Str(ebv.surface))) + " Ha", "G", "B"
'           owner.fdessin.UC_graphiqueB.ecr_texta 1620, 1530, "Coef. de ruissellement = " + ajout_zero(Trim(Str(ebv.imper))), "G", "B"
'           owner.fdessin.UC_graphiqueB.ecr_texta 2055, 1995, "Nombre d'habitants = " + ajout_zero(Trim(Str(ebv.nhab))), "G", "B"
'           owner.fdessin.UC_graphiqueB.ecr_texta 2160, 2505, "Taux de dilution = " + ajout_zero(Trim(Str(ebv.tdilu))), "G", "B"
'           owner.fdessin.UC_graphiqueB.ecr_texta 2875, 540, "Longueur = " + ajout_zero(Trim(Str(ebv.lghydr))) + " m", "G", "B"
'           owner.fdessin.UC_graphiqueB.ecr_texta 3145, 810, "Pente = " + ajout_zero(Trim(Str(ebv.phydr))) + " (1/10000)", "G", "B"
            sresult = "Surface = " + ajout_zero(Trim(str(ebv.surface))) + " Ha"
            sresult = sresult + Chr(13) + Chr(10) + "Longueur = " + ajout_zero(Trim(str(ebv.lghydr))) + " m"
            sresult = sresult + Chr(13) + Chr(10) + "Pente = " + ajout_zero(Trim(str(ebv.phydr))) + " (1/10000)"
            sresult = sresult + Chr(13) + Chr(10) + "Coef. de ruissellement = " + ajout_zero(Trim(str(ebv.imper)))
            If ebv.tdilu > 0 Then
                sresult = sresult + Chr(13) + Chr(10) + "Taux de dilution = " + ajout_zero(Trim(str(ebv.tdilu)))
            End If
            If ebv.nhab > 0 Then
                sresult = sresult + Chr(13) + Chr(10) + "Nombre d'habitants = " + ajout_zero(Trim(str(ebv.nhab)))
             End If
            If ebv.ceau > 0 Then
                sresult = sresult + Chr(13) + Chr(10) + "Consommation eau = " + ajout_zero(Trim(str(ebv.ceau))) + " l/hab/j"
            End If
            Me.Lab_bas.Caption = sresult
            Me.SSTab1.TabEnabled(1) = True
          Else
''           owner.fdessin.UC_graphiqueB.ecr_texta 1665, 1140, "Surface", "G", "B"
''           owner.fdessin.UC_graphiqueB.ecr_texta 1620, 1530, "Coef. de ruissellement", "G", "B"
''           owner.fdessin.UC_graphiqueB.ecr_texta 2055, 1995, "Nombre d'habitants", "G", "B"
''           owner.fdessin.UC_graphiqueB.ecr_texta 2160, 2505, "Taux de dilution", "G", "B"
''           owner.fdessin.UC_graphiqueB.ecr_texta 2875, 540, "Longueur", "G", "B"
''           owner.fdessin.UC_graphiqueB.ecr_texta 3145, 810, "Pente", "G", "B"
               Me.Frm_bv.Caption = "Hydraulique B.V.  " + Trim(nombassin)
            Me.Lab_bas.Caption = ""
            Me.SSTab1.TabEnabled(1) = False
          End If
            Call ini_form_exist
            ouv_sauve = False
            save_fich = True
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

Private Sub Cb_deversoir_KeyDown(KeyCode As Integer, Shift As Integer)
    dev_texte = Cb_deversoir.Text
    Cb_deversoir.Text = dev_texte
End Sub

Private Sub Cb_deversoir_KeyPress(KeyAscii As Integer)
    dev_texte = Cb_deversoir.Text
End Sub

Private Sub Chk_charge_Click()
Call OK_lignes_Click

End Sub

Private Sub Chk_eau_Click()
Call OK_lignes_Click
End Sub

Private Sub Chk_piezo_Click()
Call OK_lignes_Click

End Sub

Private Sub Chk_Qpluie_Click()
Call OK_lignes_Click
End Sub

Private Sub Chk_Qrin_Click()
Call OK_lignes_Click
End Sub

Private Sub Chk_Qts_Click()
Call OK_lignes_Click
End Sub

Private Sub Cmd_amo_Click()
Call dessin_courbe_amo
End Sub

Private Sub Cmd_annulm_Click()
    Me.Tb_amo(0).Text = rempl_virgule(Format(edessdo.tron_amo.conduit.Diametre * 1000, "###0"))
    Me.Tb_amo(1).Text = rempl_virgule(Format(edessdo.tron_amo.conduit.pente * 10000, "###0"))
    Me.Tb_amo(3).Text = rempl_virgule(Format(edessdo.tron_amo.conduit.Longueur, "###0.00"))
    Me.Tb_amo(2).Text = rempl_virgule(Format(edessdo.tron_amo.conduit.rugosite, "###0"))
    Me.Cmd_annulm.Visible = False
    Me.Cmd_resum.Enabled = False
    Call calcul_amont
End Sub


Private Sub Cmd_annulv_Click()
    Me.Tb_ava(0).Text = rempl_virgule(Format(edessdo.tron_ava.conduit.Diametre * 1000, "###0"))
    Me.Tb_ava(1).Text = rempl_virgule(Format(edessdo.tron_ava.conduit.pente * 10000, "###0"))
    Me.Tb_ava(3).Text = rempl_virgule(Format(edessdo.tron_ava.conduit.Longueur, "###0.00"))
    Me.Tb_ava(2).Text = rempl_virgule(Format(edessdo.tron_ava.conduit.rugosite, "###0"))
    Me.Cmd_annulv.Visible = False
    Me.Cmd_resuv.Enabled = False
    Call calcul_aval
End Sub

Private Sub Cmd_ava_Click()
Call dessin_courbe_ava
End Sub

Private Sub Cmd_dech_Click()
Call dessin_courbe_dech
End Sub

Private Sub Cmd_Pvm_Click()
    Me.Tb_ava(1).Text = rempl_virgule(Format(edessdo.iradav - 1, "###0"))
    Me.Cmd_annulv.Visible = True
    Me.Cmd_resuv.Enabled = True
'Call calcul_aval
End Sub

Private Sub Cmd_Pvp_Click()
    Me.Tb_ava(1).Text = rempl_virgule(Format(edessdo.iradav + 1, "###0"))
    Me.Cmd_annulv.Visible = True
    Me.Cmd_resuv.Enabled = True
'Call calcul_aval

End Sub

Private Sub Cmd_reinit_Click()


    ok_imp = False
    Me.Tb_dev(0).Text = "0.0"
    Me.Tb_dev(1).Text = "0.0"
    Me.Tb_dev(2).Text = "0.0"
    Me.Tb_dev(3).Text = "0.0"
    Me.Cmd_resudo.Enabled = True
        edessdo.tron_ava.Absamo = edessdo.tron_amo.Absava
        edessdo.tron_ava.Absava = edessdo.tron_ava.Absamo + edessdo.tron_ava.conduit.Longueur
        edo.tron_ava.Absamo = edessdo.tron_amo.Absava
        edo.tron_ava.Absava = edo.tron_ava.Absamo + edessdo.tron_ava.conduit.Longueur
        edo.Longueur = 0
        edo.pente = 0
        edo.hauteur = 0
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
    ok_imp = True
    Call calcul_dech
    Cmd_resudech.Enabled = True
    Cmd_resudech.SetFocus
    Call key13(Me)
    Cmd_resudech.Enabled = False
    ouv_sauv = True
End Sub
Private Sub calcul_dech()
Dim qv As deb_vit
Dim canal As conduite
Dim Qdev As Double
Dim qps As Double, vps As Double, hautdech As Double, vdech As Double
Dim ltc() As Variant, ct() As Variant, qvm(5) As Variant
Dim qcal As Double
Dim lambda As Double, dh As Double, hcr As Double, hmin As Double, Nen As Double, Ned As Double
Dim nok As Boolean
Dim i As Integer
Dim ok1 As Boolean, ok_torr As Boolean, ok_fluv As Boolean, ok_suite As Boolean
Dim mes_Res As String, mes_Res1 As String, mes_res_do As String
Dim zSeuil As Double, h As Double
'julienne
    Dim tr As troncon
    Dim res_conduit As debit_conduit


ok_tab5 = True
nok = True
'' tests pour verif_fonctionnement autre debit
'edessdo.Qpluie = 450
'edo_res.Qdev = 0.3
'' fin  tests pour verif_fonctionnement autre debit


ok_torr = False
ok_fluv = False
ok_suite = False
i = 0
mes_Res = ""
mes_Res1 = ""
edessdo.tron_dech.conduit.typ = 2
edessdo.tron_dech.conduit.Diametre = txtVersNum(Tb_dech(0).Text) / 1000#
edessdo.tron_dech.conduit.pente = txtVersNum(Tb_dech(1).Text) / 10000
edessdo.tron_dech.conduit.rugosite = txtVersNum(Tb_dech(2).Text)
edessdo.tron_dech.conduit.Longueur = txtVersNum(Tb_dech(3).Text)
'resudev.ddech = TB_dech(0).Text
'resudev.iraddech = TB_dech(1).Text
'resudev.kdech = TB_dech(2).Text
'resudev.Ldech = TB_dech(3).Text
edessdo.tron_dech.Absamo = edo.Absava
'edessdo.tron_dech.radamo = edo.radava
edessdo.tron_dech.Absava = edessdo.tron_dech.Absamo + edessdo.tron_dech.conduit.Longueur
'edessdo.tron_dech.radava = edessdo.tron_dech.radamo - edessdo.tron_dech.conduit.Longueur * edessdo.tron_dech.conduit.pente
edessdo.tron_dech.radava = edessdo.rdoex
edessdo.tron_dech.radamo = edessdo.tron_dech.radava + edessdo.tron_dech.conduit.Longueur * edessdo.tron_dech.conduit.pente
canal = edessdo.tron_dech.conduit
'dessin dans frmdessin
    Call init_graphdech(owner.fdessin.UC_graphique2)
    Call dess_troncon(owner.fdessin.UC_graphique2, edessdo.tron_ava, couleur.gris_clair) 'vbRed)
    Call dess_troncon(owner.fdessin.UC_graphique2, edo.tron_ava, couleur.gris)   'vbRed)
    Call dess_troncon(owner.fdessin.UC_graphique2, edessdo.tron_amo, couleur.gris) ' vbRed)
    Call dess_do(owner.fdessin.UC_graphique2, edo, couleur.gris)
    Call dess_troncon(owner.fdessin.UC_graphique2, edessdo.tron_dech, couleur.noir) 'vbmagenta)
'calcul des conditions d'ecoulement
    qv = debvit_ps(canal)
    qps = qv.debit
    vps = qv.vitesse
'affichage
    Me.Lb_vpsdech.Caption = "Vitesse pleine section = " + ajout_zero(Trim(str(Round(qv.vitesse, 3)))) + " m/s"
    Me.Lb_Dpsdech.Caption = "Débit pleine section = " + ajout_zero(Trim(str(Round(qv.debit, 3)))) + " m3/s"
    mes_Res = "Débit pleine section = " + ajout_zero(Trim(str(Round(qv.debit, 3)))) + " m3/s"
    mes_Res = mes_Res + Chr(13) + Chr(10) + "Vitesse pleine section = " + ajout_zero(Trim(str(Round(qv.vitesse, 3)))) + " m/s"
    resudev.dpsdech = ajout_zero(Trim(str(Round(qv.debit, 3))))
    resudev.vpsdech = ajout_zero(Trim(str(Round(qv.vitesse, 3))))
' calcul du débit déversé

While nok And i < 20
    mes_Res1 = ""
    nok = False
    Qdev = edo_res.Qdev
    
    Me.lb_Qdev.Caption = "Débit déversé : " + ajout_zero(Trim(str(Round(Qdev, 3)))) + " m3/s"
    mes_Res1 = Chr(13) + Chr(10) + Chr(13) + Chr(10) + "Débit déversé : " + ajout_zero(Trim(str(Round(Qdev, 3)))) + " m3/s"
    
    If Qdev > qv.debit Then
        Me.Lb_vqdev.Caption = "Vitesse d'écoulement à Qdev > Qps = " + "   " + " m/s"
        Me.Lb_hqdev.Caption = "Hauteur d'eau à Qdev > Qps = " + "   " + " m"
        mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Vitesse d'écoulement à Qdev > Qps = " + "   " + " m/s"
        mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Hauteur d'eau à Qdev > Qps = " + "   " + " m"
        hautdech = canal.Diametre
        vdech = Qdev * qv.vitesse / qv.debit
    Else
        Call cana(canal, ct)
        ltc = calc_par(canal)
        qvi = caltran1(Qdev * 1000, ct, ltc)
        hautdech = qvi(5)
        vdech = qvi(2)
        
        Me.Lb_vqdev.Caption = "Vitesse d'écoulement à Q déversé = " + ajout_zero(Trim(str(Round(qvi(2), 3)))) + " m/s"
        Me.Lb_hqdev.Caption = "Hauteur d'eau Q déversé = " + ajout_zero(Trim(str(Round(qvi(5), 3)))) + " m"
        mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Vitesse d'écoulement à Q déversé = " + ajout_zero(Trim(str(Round(qvi(2), 3)))) + " m/s"
        mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Hauteur d'eau Q déversé = " + ajout_zero(Trim(str(Round(qvi(5), 3)))) + " m"
    End If
        resudev.vqdev = ajout_zero(Trim(str(Round(vdech, 3))))
        resudev.hqdev = ajout_zero(Trim(str(Round(hautdech, 3))))
 
    
    ' determination du niveau de la surface d'eau amont zradier + hauteur
    ' par rapport au niveau des plus hautes eaux aval
 '   If edessdo.phex > edessdo.tron_dech.radamo + hautdech Then
 ' julienne a verifier

 
 ' a redefinir comme dessin_decharge:  interpiezo,intercharge ...
    If edessdo.phex > edessdo.tron_dech.radava + hautdech Then
   '     MsgBox "remous", vbOKOnly
        mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Remous"
 '       ok_suite = False
        
        
   End If
'    Else
    ' determination du regime torrentiel ou fluvial
    tr = edessdo.tron_dech
    res_conduit = calc_debit_tr(tr, Qdev)
    res_conduit.zphe_ava = edessdo.phex
    Call inter_piezo_eau(tr, res_conduit)
    Call inter_charge_tr(tr, res_conduit)
    hautdech = (res_conduit.zeau_amo.Y - tr.radamo)
    vdech = calcul_vitesse(tr, res_conduit, hautdech)

        regime = verif_regime(Qdev, canal, hautdech)
 
  '      regime = calcul_ecoul(Qdev, canal.Diametre, beta)
        
        resudev.regime = regime
        
        Select Case regime
            Case "TORREN."
                ok_torr = True
                mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Regime Torrentiel"
               ' calcul des niveaux d'energie
                'a revoir saisie du lambda
                lambda = 0.75
                lambda = edessdo.Centon
                'niveau d'energie nécessaire
                'calcul de la charge necessaire a l'entree (coefficient d'entonnement 0.75)
                dh = lambda * (vdech ^ 2) / (2 * 9.81)
                Nen = edessdo.tron_dech.radamo + hautdech + dh
                'niveau d'energie disponible
                hcr = (((Qdev / edo.Longueur) ^ 2) / 9.81) ^ (1# / 3#)
                vcr = (9.81 * hcr) ^ 0.5
                hmin = hcr + (vcr ^ 2) / (2 * 9.81)
                Ned = edo.radamo + edo.hauteur + hmin
                 If Ned >= Nen Then
                    mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Nv energie disponible " & ajout_zero(Trim(str(Round(Ned, 3)))) & " > Nv necessaire" & ajout_zero(Trim(str(Round(Nen, 3))))
                    ok_suite = True
                    nok = False
                    'a revoir suite torrentiel
                Else
                    MsgBox "Ecoulement torrentiel " + Chr(13) + " Nv energie disponible " & ajout_zero(Trim(str(Round(Ned, 3)))) & " < Ne necessaire " & ajout_zero(Trim(str(Round(Nen, 3)))) & Chr(10) & " Modifier le dispositif "
                    mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Nv energie disponible " & ajout_zero(Trim(str(Round(Ned, 3)))) & " < Ne necessaire " & ajout_zero(Trim(str(Round(Nen, 3)))) & Chr(10) & " Modifier le dispositif "
                    ok_suite = False
'                    nok = True
'                    i = i + 1
                End If
                'a revoir suite torrentiel -> non ca doit etre ok
            Case "FLUVIAL"
                ok_fluv = True
                mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Regime Fluvial"
                ' calcul des niveaux d'energie
                'a revoir saisie du lambda
                lambda = 0.75
                lambda = edessdo.Centon
                'niveau d'energie nécessaire
                'calcul de la charge necessaire a l'entree coefficient d'entonnement 0.75
                dh = lambda * vdech ^ 2 / (2 * 9.81)
                Nen = edessdo.tron_dech.radamo + hautdech + dh
                'niveau du seuil
                zSeuil = edo.radamo + edo.hauteur
                If zSeuil >= Nen Then
'                    MsgBox "nappe libre", vbOKOnly
                    mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "nappe libre"
                    resudev.regime = regime + "  nappe libre"
                     edo_res.c = 1
                     nok = True
                     mes_res_do = verif_do_charge(edo, edo_res, edessdo.tron_amo, edessdo.Qpluie / 1000, edessdo.Qrin / 1000, owner.fdessin.UC_graphique1)
                     If Abs(edo_res.Qdev - Qdev) < 0.0005 Then
                       nok = False
                       ok_suite = True
                    End If
             Else
'                    MsgBox "nappe noyée " + Str(i), vbOKOnly
                    mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "nappe noyée"
                    resudev.regime = regime + "  nappe noyée"
                    h = Nen - zSeuil
                    ' calcul du coef C de nappe noyee h/hm
 '                   Debug.Print h / edo_res.HM
                    edo_res.c = recup_do_C(h / edo_res.HM)
                    nok = True
                    i = i + 1
'                     mes_res_do = verif_do_charge(edo, edo_res, edessdo.tron_amo, edessdo.Qpluie / 1000, edessdo.Qrin / 1000, UC_graphique4)
                     mes_res_do = verif_do_charge(edo, edo_res, edessdo.tron_amo, edessdo.Qpluie / 1000, edessdo.Qrin / 1000, owner.fdessin.UC_graphique1)
                    Me.Txtb_deversoir.Text = mes_res_do
                    If Abs(edo_res.Qdev - Qdev) < 0.0005 Then
                       nok = False
                       ok_suite = True
                    End If
                End If
                'niveau d'energie disponible
        End Select
'    End If
Wend
If Not nok Then
    'affichage du résultat
    mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Hauteur de la lame d'eau :"
    mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Amont  :" + ajout_zero(Trim(str(Round(edo_res.Ham, 3)))) + " m"
    mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Aval  :" + ajout_zero(Trim(str(Round(edo_res.Hav, 3)))) + " m"
    mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Hauteur de la charge :"
    mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Amont  :" + ajout_zero(Trim(str(Round(edo_res.Haam, 3)))) + " m"
    mes_Res1 = mes_Res1 + Chr(13) + Chr(10) + "Aval  :" + ajout_zero(Trim(str(Round(edo_res.Haav, 3)))) + " m"
    resudev.Ham = ajout_zero(Trim(str(Round(edo_res.Ham, 3))))
    resudev.Hav = ajout_zero(Trim(str(Round(edo_res.Hav, 3))))
    resudev.Haam = ajout_zero(Trim(str(Round(edo_res.Haam, 3))))
    resudev.Haav = ajout_zero(Trim(str(Round(edo_res.Haav, 3))))
    mes_Res = mes_Res + mes_Res1
    Txtb_decharge.Text = mes_Res
    
    'Dessin du fonctionnement dans l'onglet Decharge
    'dessin de la ligne d'eau
    ' conduite amont
    
    'dessin des lignes de charges
    Call dessin_decharge(owner.fdessin.UC_graphique2)
    
    Else
      mes_Res = mes_Res + Chr(13) + Chr(10) + Chr(13) + Chr(10) + " Anomalie de fonctionnement " + Chr(13) + Chr(10) + "Redimensionner !" + mes_Res1
      Txtb_decharge.Text = mes_Res
        ok_imp = False

End If
    Me.Cmd_resudech.Enabled = False
        ' impression true
                    Me.mnuprint.Enabled = True
End Sub



Private Sub dessin_decharge_0(ByRef uc_g As UC_graphique)
Dim zplam_am As Double, zplam_av As Double, zplav_am As Double, zplav_av As Double
'Dim qvm(5) As Variant, haut As Double, pentmot As Double
Dim tr As troncon ', uc_g As UC_graphique
Dim res_conduit As debit_conduit
Dim qcal
'dessin de la ligne dans le deversoir
'dans la frmdessin
'Set uc_g = owner.fdessin.UC_graphique2
'zplam_av = edo.radamo + edo_res.Tram
'dessin troncon amont
    zplam_av = edo.radamo + edessdo.Tram
    tr = edessdo.tron_amo
    qcal = edessdo.Qpluie / 1000
    res_conduit = calc_debit_tr(tr, qcal)
    'dessin des lignes d'eau
        zplam_am = res_conduit.hauteur + tr.radamo
        ' uc_g.dess_lign tr.Absamo, zplam_am, tr.Absava, zplam_av, couleur.bleu
        res_conduit.zphe_ava = zplam_av
        Call inter_piezo_eau(tr, res_conduit)
'dessin ligne piezo amont
'        uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.zeau_amo.Y, res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, couleur.orange, 2
'        uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, res_conduit.p_Eau_inter0.X, res_conduit.p_Eau_inter0.Y, couleur.orange, 2
'        uc_g.dess_lign res_conduit.p_Eau_inter0.X, res_conduit.p_Eau_inter0.Y, res_conduit.zeau_ava.X, res_conduit.piezoava, couleur.orange, 2
'dessin ligne d'eau amont
         uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.zeau_amo.Y, res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, couleur.bleu, 2
        uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, res_conduit.p_Eau_inter0.X, res_conduit.p_Eau_inter0.Y, couleur.bleu, 2
        uc_g.dess_lign res_conduit.p_Eau_inter0.X, res_conduit.p_Eau_inter0.Y, res_conduit.zeau_ava.X, res_conduit.zeau_ava.Y, couleur.bleu, 2
   'dessin charge
        'recalcul de charge amont en fonction de hauteur d'eau
        ' verification vitesse d'écoulement amont pour qcri
        Call inter_charge_tr(tr, res_conduit)
        uc_g.dess_lign tr.Absamo, res_conduit.chargeamo, res_conduit.piezointer.X, res_conduit.chargeinter, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer.X, res_conduit.chargeinter, res_conduit.piezointer0.X, res_conduit.chargeinter0, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer0.X, res_conduit.chargeinter0, tr.Absava, res_conduit.chargeava, couleur.rouge, 2
    
'dessin ligne d'eau sur la lame
    'dessin des lignes d'eau
        'uc_g.dess_lign edo.Absamo, edo.radamo + edo_res.Tram, edo.Absava, edo.radamo + edo.hauteur + edo_res.Hav, couleur.bleu, 2
        uc_g.dess_lign edo.Absamo, edo.radamo + edessdo.Tram, edo.Absava, edo.radamo + edo.hauteur + edo_res.Hav, couleur.bleu, 2
    'dessin charge
        uc_g.dess_lign edo.Absamo, edo.radamo + edo_res.Haam, edo.Absava, edo.radava + edo_res.Haav, couleur.rouge, 2
     
'dessin troncon décharge
    tr = edessdo.tron_dech
'    '     If edessdo.phex > (edessdo.tron_dech.radava + hautdech) Then
'    '        zplam_av = edessdo.phex
'    '    Else
'            zplam_av = edessdo.tron_dech.radava + hautdech
'    '    End If
'        zplam_am = edessdo.tron_dech.radamo + hautdech
'    'dessin des lignes d'eau
'        'uc_g.dess_lign tr.Absamo, zplam_am, tr.Absava, zplam_av, couleur.jaune, 2
res_conduit = calc_debit_tr(edessdo.tron_dech, edo_res.Qdev)

    'dessin des lignes d'eau
        res_conduit.zphe_ava = edessdo.phex
        Call inter_piezo_eau(tr, res_conduit)
        ' uc_g.dess_lign tr.Absamo, zplam_am, res_conduit.piezointer.X, res_conduit.piezointer.Y, couleur.bleu, 2
        ' uc_g.dess_lign res_conduit.piezointer.X, res_conduit.piezointer.Y, tr.Absava, zplam_av, couleur.bleu, 2
        uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.zeau_amo.Y, res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, couleur.bleu, 2
        uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, res_conduit.p_Eau_inter0.X, res_conduit.p_Eau_inter0.Y, couleur.bleu, 2
        uc_g.dess_lign res_conduit.p_Eau_inter0.X, res_conduit.p_Eau_inter0.Y, res_conduit.zeau_ava.X, res_conduit.zeau_ava.Y, couleur.bleu, 2
'dessin ligne piezo
'        uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.piezoamo, res_conduit.p_Eau_inter.X, res_conduit.piezointer.Y, couleur.orange, 2
'        uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.piezointer.Y, res_conduit.piezointer0.X, res_conduit.piezointer0.Y, couleur.orange, 2
'        uc_g.dess_lign res_conduit.piezointer0.X, res_conduit.piezointer0.Y, res_conduit.zeau_ava.X, res_conduit.piezoava, couleur.orange, 2
    
    'dessin charge
        ' dessin de la charge repris par inter_charge_pr
        Call inter_charge_tr(tr, res_conduit)
        uc_g.dess_lign tr.Absamo, res_conduit.chargeamo, res_conduit.piezointer.X, res_conduit.chargeinter, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer.X, res_conduit.chargeinter, res_conduit.piezointer0.X, res_conduit.chargeinter0, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer0.X, res_conduit.chargeinter0, tr.Absava, res_conduit.chargeava, couleur.rouge, 2
End Sub
'Private Function complet_list_do(ByVal liste1 As Variant) As Variant
'Dim liste() As Variant
'Dim i As Integer, j As Integer
'i = -1
''Select Case ebstock_dess.type
''     Case Is = "circ", "cond"
''        ReDim liste(UBound(liste1) + 4, 3)
''     Case Is = "rect"
'        ReDim liste(UBound(liste1) + 5, 3)
'' End Select
'For j = 0 To UBound(liste1)
'    i = i + 1
''    ReDim Preserve liste(i, 3)
'    liste(i, 1) = liste1(j, 1)
'    liste(i, 2) = liste1(j, 2)
'    liste(i, 3) = liste1(j, 3)
'Next
''i = i + 1
'''ReDim Preserve liste(i, 3)
''liste(i, 1) = ""
''liste(i, 2) = ""
''liste(i, 3) = ""
'i = i + 1
''ReDim Preserve liste(i, 3)
'liste(i, 1) = "Type de bassin"
'liste(i, 3) = ""
''Select Case ebstock_dess.type
''     Case Is = "circ"
''     Case Is = "rect"
'         liste(i, 2) = "rectangulaire"
'         i = i + 1
'         liste(i, 1) = Lb_intlong.Caption
'         liste(i, 2) = txtVersNum(Me.Tb_long.Text)
'         liste(i, 3) = Lb_ulong.Caption
'         i = i + 1
'         liste(i, 1) = Lb_intlarg.Caption
'         liste(i, 2) = txtVersNum(Me.Tb_larg.Text)
'         liste(i, 3) = Lb_ularg.Caption
'         i = i + 1
'         liste(i, 1) = Lb_intprof.Caption
'         liste(i, 2) = txtVersNum(Me.Tb_prof.Text)
'         liste(i, 3) = Lb_uprof.Caption
'         i = i + 1
'         liste(i, 1) = Lb_intrap.Caption
'         liste(i, 2) = txtVersNum(Me.Tb_rap.Text)
'         liste(i, 3) = Lb_urap.Caption
''End Select
'complet_list_do = liste
'End Function
Public Function rec_listdo(ByVal nom As String) As Variant
Dim liste() As Variant, list() As Variant
Dim chaine As String, nom1 As String
Dim i As Integer, j As Integer, ij As Integer
chaine = ""
j = -1
ReDim liste(0)
For i = 1 To Len(nom)
    If Mid(nom, i, 1) = Chr(13) Or Mid(nom, i, 1) = Chr(10) Then
        If Len(Trim(chaine)) > 0 Then
            j = j + 1
            ReDim Preserve liste(j)
            liste(j) = Trim(chaine)
        End If
            chaine = ""
        Else
        chaine = chaine + Mid(nom, i, 1)
    End If
Next
If Len(Trim(chaine)) > 0 Then
    j = j + 1
    ReDim Preserve liste(j)
    liste(j) = Trim(chaine)
End If
ReDim list(j, 3)
For i = 0 To UBound(liste)
    nom1 = liste(i) + "  "
    j = InStr(1, nom1, " = ")
    If j > 0 Then
        list(i, 1) = Trim(Mid(nom1, 1, j - 1))
        ij = j + 2
        nom1 = Right(nom1, Len(nom1) - ij)
        If Trim(nom1) <> "" Then
        j = InStr(1, nom1, " ")
        list(i, 2) = Trim(Mid(nom1, 1, j - 1))
    '    ij = j + 1
        nom1 = Right(nom1, Len(nom1) - j)
        list(i, 3) = Trim(nom1)
        Else
        list(i, 2) = ""
        list(i, 3) = ""
        End If
    Else
        list(i, 1) = Trim(nom1)
        list(i, 2) = ""
        list(i, 3) = ""
    End If
Next
rec_listdo = list
End Function
Public Function recup_mnuprint()
    recup_mnuprint = Me.mnuprint.Enabled
End Function
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
ok_imp = False
ok_tab4 = True
Call ini_resudo
longdo = txtVersNum(Me.Tb_dev(0).Text)
hautdo = txtVersNum(Me.Tb_dev(1).Text)
pentedo = txtVersNum(Me.Tb_dev(2).Text)
resudev.Ldev = Me.Tb_dev(0).Text
resudev.Hcret = Me.Tb_dev(1).Text
resudev.Pdev = Me.Tb_dev(2).Text
edo.Longueur = longdo
edo.hauteur = hautdo
edo.pente = pentedo
' test déversoir latéral ***
'If Not test_do_lat(edessdo) Then
'houpie 20040114 test aval en charge
Dim qcri As Double, qv As deb_vit
qv = debvit_ps(edessdo.tron_ava.conduit)
qcri = edessdo.Qrin
If qv.debit * 1000 < qcri Then
    Call pre_dimdo(edo)
    Me.Tb_dev(0).Text = rempl_virgule(Format(edo.Longueur, "##0.00"))
    Me.Tb_dev(1).Text = rempl_virgule(Format(edo.hauteur, "###0.000"))
    Me.Tb_dev(2).Text = rempl_virgule(Format(edo.pente, "#0.0000"))
    Call modi_longce
    sresult = pre_calculdo(edo, Resup_do)
        Me.Lb_longce(0).Caption = sresult
'        Me.SSTab_result.Tab = 0
        Me.Txtb_deversoir.Text = ""
    
    'calcul longueur restante
    lgreste = edessdo.lgdisp - edo.tron_ava.Absava
    If lgreste > 0 Then
    'calcul pente restante
        preste = (edo.tron_ava.radava - edessdo.rdoav) / lgreste * 10000#
        Me.Lb_preste.Caption = "Pente disponible = " + ajout_zero(Trim(str(Round(preste)))) + "  1/10000"
    Else
        Me.Lb_preste.Caption = ""
    End If
    
    ok = verif_resu(Resup_do)
    ' dessin dans l UCG4
    
    
    'dessin dans  frmdessin
    Call dessin_do(False)
 Else
  ok = False
  MsgBox "la conduite aval n'est pas en charge à QREF!", vbExclamation, "Calcul DO"
End If

    If ok Then
        Me.Frame4.Enabled = False
        Me.Chk_Qpluie.Enabled = False
        Me.Chk_Qpluie.Value = 0
        Me.Chk_Qts.Value = 1
        Me.Chk_Qrin = 1
        If Me.Chk_eau.Value + Me.Chk_piezo.Value + Me.Chk_charge.Value = 0 Then
            Me.Chk_eau.Value = 1
            Me.Chk_charge.Value = 1
            Me.Chk_piezo.Value = 1
        End If
        Me.Frame4.Enabled = True
        Me.Frame4.Visible = True
        Call OK_lignes_Click
     '   If edessdo.Tram > 0 Then
            Me.Cmd_VerifDo.Enabled = True
     '   End If
    End If
    
    Me.Cmd_resudo.Enabled = False
    Me.Cmd_resudo.Enabled = True
    Me.Cmd_resudo.SetFocus
    Call key13(Me)
    Me.Cmd_resudo.Enabled = False
    ouv_sauv = True

    'Me.SSTab1.TabEnabled(5) = True
    
'Else
'' calcul do_lat_ conduite aval libreav
'Call calcul_do_lat(edessdo, edo)
'    Me.Tb_dev(0).Text = rempl_virgule(Format(edo.Longueur, "##0.00"))
'    Me.Tb_dev(1).Text = rempl_virgule(Format(edo.hauteur, "###0.000"))
'    Me.Tb_dev(2).Text = rempl_virgule(Format(edo.pente, "#0.0000"))
'
'    Call dessin_do(False)
'
'End If
End Sub
Private Sub dessin_courbe_dech()
Dim troamo As troncon
Dim canal As conduite
   canal.Diametre = edessdo.tron_dech.conduit.Diametre ' / 1000#
    canal.Longueur = 5
    canal.pente = edessdo.tron_dech.conduit.pente '/ 10000#
    canal.rugosite = edessdo.tron_dech.conduit.rugosite
    canal.typ = 2
    With troamo
      .Absamo = 0#
      .Absava = .Absamo + canal.Longueur
      .conduit = canal
      .radava = 100#
      .radamo = 100# + 0.3  'cana_amo.Longueur * cana_amo.pente
    End With
    Call dess_courbe_debit_tr(troamo, 0, "Courbe débit conduite de décharge")
End Sub
Private Sub dessin_courbe_amo()
Dim troamo As troncon
Dim canal As conduite
   canal.Diametre = edessdo.dam / 1000#
    canal.Longueur = 5
    canal.pente = edessdo.iRadam / 10000#
    canal.rugosite = edessdo.Kam
    canal.typ = 2
    With troamo
      .Absamo = 0#
      .Absava = .Absamo + canal.Longueur
      .conduit = canal
      .radava = edessdo.rdoav
      .radamo = edessdo.rdoav + 0.3 'cana_amo.Longueur * cana_amo.pente
    End With
    Call dess_courbe_debit_tr(troamo, 0, "Courbe débit conduite amont")
End Sub
Private Sub dessin_courbe_ava()
Dim troamo As troncon
Dim canal As conduite
   canal.Diametre = edessdo.dav / 1000#
    canal.Longueur = 5
    canal.pente = edessdo.iradav / 10000#
    canal.rugosite = edessdo.kav
    canal.typ = 2
    With troamo
      .Absamo = 0#
      .Absava = .Absamo + canal.Longueur
      .conduit = canal
      .radava = edessdo.rdoam - 0.3
      .radamo = edessdo.rdoam 'cana_amo.Longueur * cana_amo.pente
    End With
    Call dess_courbe_debit_tr(troamo, 0, "Courbe débit conduite aval")
End Sub
Private Sub dess_dech_print(ByRef uc_g As UC_graphique)
    Call init_graphdech(uc_g)
    Call dess_troncon(uc_g, edessdo.tron_ava, couleur.gris_clair) 'vbRed)
    Call dess_troncon(uc_g, edo.tron_ava, couleur.gris)   'vbRed)
    Call dess_troncon(uc_g, edessdo.tron_amo, couleur.gris) ' vbRed)
    Call dess_do(uc_g, edo, couleur.gris)
    Call dess_troncon(uc_g, edessdo.tron_dech, couleur.noir) 'vbmagenta)
    Call dessin_decharge(uc_g)
End Sub
Private Sub dess_do_print(ByRef uc_g As UC_graphique, ByVal ok1 As Boolean, _
    ByVal ok2 As Boolean, ByVal ok3 As Boolean)
    Call init_graphdo(uc_g)
    Call dess_troncon(uc_g, edessdo.tron_amo, couleur.gris) ' vbBlack)
    Call dess_predo(uc_g, edo, couleur.noir)
    Call dess_cot(uc_g, couleur.noir) ' vbBlack)
    Call dessin_do_debpointe(uc_g, ok1, ok2, ok3) 'okcharge, okpiezo, okeau
End Sub

Sub dessin_do(Optional ByVal ok As Boolean)
Call init_graphdo(owner.fdessin.UC_graphique1)
Call init_graphdo(owner.fdessin.UC_graphique2)

Call dessin_do_objet(owner.fdessin.UC_graphique1)
If ok Then
Call dessin_do_hydrau(owner.fdessin.UC_graphique1, True, True, True, True, True, False)
End If
End Sub
Sub dessin_do_objet(ByRef uc_g As UC_graphique)

Call dess_troncon(uc_g, edessdo.tron_amo, couleur.gris) ' vbBlack)
Call dess_predo(uc_g, edo, couleur.noir)
Call dess_cot(uc_g, couleur.noir) ' vbBlack)
'Call dess_troncon(UC_graphique4, edessdo.tron_ava, vbRed)
End Sub
Sub dessin_do_hydrau(ByRef uc_g As UC_graphique, ByVal okcharge As Boolean, ByVal okpiezo As Boolean, ByVal okeau As Boolean, ByVal okqts As Boolean, okqrin As Boolean, okqpluie As Boolean)
    Dim res_conduit As debit_conduit
    Dim tr As troncon
    Dim qcal As Double
    Dim edptam(3) As points

    '    Call dess_piezo(uc_g, edessdo.tron_amo, edessdo.Qts, couleur.magenta) ' vbGreen) ' RGB(128, 64, 128))
    If okqrin Then
        qcal = (edessdo.Qrin + edessdo.Qts) / 1000#
        tr = edessdo.tron_amo
        res_conduit = calc_debit_tr(tr, qcal)
        res_conduit.zphe_ava = edo.radamo + edo.hauteur
        Call inter_piezo_eau(tr, res_conduit)
        If okpiezo Then
            uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.piezoamo, res_conduit.p_Eau_inter.X, res_conduit.piezointer.Y, couleur.orange, 2
'            uc_g.dess_lign res_conduit.p_Eau_inter.x, res_conduit.piezointer.y, res_conduit.p_Eau_inter0.x, res_conduit.piezointer0.y, couleur.orange, 2
            uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.piezointer.Y, res_conduit.p_Eau_inter2.X, res_conduit.piezointer2.Y, couleur.orange, 2
            uc_g.dess_lign res_conduit.p_Eau_inter2.X, res_conduit.piezointer2.Y, res_conduit.p_Eau_inter1.X, res_conduit.piezointer1.Y, couleur.orange, 2
            uc_g.dess_lign res_conduit.p_Eau_inter1.X, res_conduit.piezointer1.Y, res_conduit.p_Eau_inter0.X, res_conduit.piezointer0.Y, couleur.orange, 2
            uc_g.dess_lign res_conduit.p_Eau_inter0.X, res_conduit.piezointer0.Y, res_conduit.zeau_ava.X, res_conduit.piezoava, couleur.orange, 2
            edptam(1).X = tr.Absava
            edptam(1).Y = res_conduit.piezoava
        End If
        If okeau Then
            uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.zeau_amo.Y, res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, couleur.bleu, 2
'            uc_g.dess_lign res_conduit.p_Eau_inter.x, res_conduit.piezointer.y, res_conduit.p_Eau_inter0.x, res_conduit.piezointer0.y, couleur.bleu, 2
            uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.piezointer.Y, res_conduit.p_Eau_inter2.X, res_conduit.piezointer2.Y, couleur.bleu, 2
            uc_g.dess_lign res_conduit.p_Eau_inter2.X, res_conduit.piezointer2.Y, res_conduit.p_Eau_inter1.X, res_conduit.piezointer1.Y, couleur.bleu, 2
            uc_g.dess_lign res_conduit.p_Eau_inter1.X, res_conduit.piezointer1.Y, res_conduit.p_Eau_inter0.X, res_conduit.piezointer0.Y, couleur.bleu, 2
            uc_g.dess_lign res_conduit.p_Eau_inter0.X, res_conduit.piezointer0.Y, res_conduit.zeau_ava.X, res_conduit.zeau_ava.Y, couleur.bleu, 2
            edptam(2).X = tr.Absava
            edptam(2).Y = res_conduit.zeau_ava.Y
        End If
    '   Call dess_piezo(uc_g, edessdo.tron_amo, edessdo.Qrin, couleur.magenta_clair) ' vbRed)
        
    '    Call dess_piezo(uc_g, edessdo.tron_amo, edessdo.Qpluie, couleur.vert) ' vbMagenta)
    '
    '    Call dess_piezo(uc_g, edo.tron_ava, edessdo.Qts, couleur.magenta) ' vbGreen)
    ''    Call dess_piezo(uc_g, edo.tron_ava, edessdo.Qrin, couleur.magenta_clair) ' vbRed)
    ' '   Call dess_piezo(UC_graphique4, troava, edessdo.Qpluie, vbMagenta)
        If okcharge Then
           Call inter_charge_tr(tr, res_conduit)
            uc_g.dess_lign tr.Absamo, res_conduit.chargeamo, res_conduit.piezointer.X, res_conduit.chargeinter, couleur.rouge, 2
'        uc_g.dess_lign res_conduit.piezointer.x, res_conduit.chargeinter, res_conduit.piezointer0.x, res_conduit.chargeinter0, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer.X, res_conduit.chargeinter, res_conduit.piezointer2.X, res_conduit.chargeinter2, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer2.X, res_conduit.chargeinter2, res_conduit.piezointer1.X, res_conduit.chargeinter1, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer1.X, res_conduit.chargeinter1, res_conduit.piezointer0.X, res_conduit.chargeinter0, couleur.rouge, 2
            uc_g.dess_lign res_conduit.piezointer0.X, res_conduit.chargeinter0, tr.Absava, res_conduit.chargeava, couleur.rouge, 2
            edptam(3).X = tr.Absava
            edptam(3).Y = res_conduit.chargeava
        End If
        tr = edo.tron_ava
        res_conduit = calc_debit_tr(tr, qcal)
        res_conduit.zphe_ava = res_conduit.piezoava
        Call inter_piezo_eau(tr, res_conduit)
    
        If okpiezo Then
             uc_g.dess_lign edptam(1).X, edptam(1).Y, res_conduit.zeau_amo.X, res_conduit.piezoamo, couleur.orange, 2
           
            uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.piezoamo, res_conduit.p_Eau_inter.X, res_conduit.piezointer.Y, couleur.orange, 2
'            uc_g.dess_lign  res_conduit.p_Eau_inter.x, res_conduit.p_Eau_inter.y, res_conduit.zeau_ava.x, res_conduit.piezoava, couleur.orange, 2
            uc_g.dess_lign res_conduit.piezointer.X, res_conduit.piezointer.Y, res_conduit.piezointer0.X, res_conduit.piezointer0.Y, couleur.orange, 2
            uc_g.dess_lign res_conduit.piezointer0.X, res_conduit.piezointer0.Y, res_conduit.zeau_ava.X, res_conduit.piezoava, couleur.orange, 2
     '       Call dess_piezo(uc_g, edo.tron_ava, edessdo.Qrin, couleur.magenta_clair) ' vbRed)
        End If
        If okeau Then
            uc_g.dess_lign edptam(2).X, edptam(2).Y, edo.Absava, edo.radava + edo.tav, couleur.bleu, 2
            
            uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.zeau_amo.Y, res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, couleur.bleu, 2
            uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, res_conduit.zeau_ava.X, res_conduit.zeau_ava.Y, couleur.bleu, 2
        End If
        If okcharge Then
            Call inter_charge_tr(tr, res_conduit)
            uc_g.dess_lign edptam(3).X, edptam(3).Y, edo.Absava, edo.radava + edo.tav, couleur.rouge, 2
            uc_g.dess_lign tr.Absamo, res_conduit.chargeamo, res_conduit.piezointer.X, res_conduit.chargeinter, couleur.rouge, 2
'        uc_g.dess_lign res_conduit.piezointer.x, res_conduit.chargeinter, res_conduit.piezointer0.x, res_conduit.chargeinter0, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer.X, res_conduit.chargeinter, res_conduit.piezointer2.X, res_conduit.chargeinter2, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer2.X, res_conduit.chargeinter2, res_conduit.piezointer1.X, res_conduit.chargeinter1, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer1.X, res_conduit.chargeinter1, res_conduit.piezointer0.X, res_conduit.chargeinter0, couleur.rouge, 2
           uc_g.dess_lign res_conduit.piezointer0.X, res_conduit.chargeinter0, tr.Absava, res_conduit.chargeava, couleur.rouge, 2
    '       Call dess_charge(uc_g, edo.tron_ava, edessdo.Qrin, couleur.cyan) ' vbCyan)
        End If
    '    Call dess_charge(uc_g, edessdo.tron_amo, edessdo.Qrin, couleur.orange) ' vbCyan)
    
End If
If okqpluie Then

   Call dessin_do_debpointe(uc_g, okcharge, okpiezo, okeau)


End If
If okqts Then
        qcal = edessdo.Qts / 1000#
        tr = edessdo.tron_amo
        res_conduit = calc_debit_tr(tr, qcal)
        res_conduit.zphe_ava = res_conduit.piezoava
        Call inter_piezo_eau(tr, res_conduit)
        If okpiezo Then
            uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.piezoamo, res_conduit.p_Eau_inter.X, res_conduit.piezointer.Y, couleur.orange, 2
            uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, res_conduit.zeau_ava.X, res_conduit.piezoava, couleur.orange, 2
            edptam(1).X = tr.Absava
            edptam(1).Y = res_conduit.piezoava
        End If
        If okeau Then
            uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.zeau_amo.Y, res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, couleur.bleu, 2
            uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.piezointer.Y, res_conduit.zeau_ava.X, res_conduit.zeau_ava.Y, couleur.bleu, 2
            edptam(2).X = tr.Absava
            edptam(2).Y = res_conduit.zeau_ava.Y
        End If
    '   Call dess_piezo(uc_g, edessdo.tron_amo, edessdo.Qrin, couleur.magenta_clair) ' vbRed)
        
    '    Call dess_piezo(uc_g, edessdo.tron_amo, edessdo.Qpluie, couleur.vert) ' vbMagenta)
    '
    '    Call dess_piezo(uc_g, edo.tron_ava, edessdo.Qts, couleur.magenta) ' vbGreen)
    ''    Call dess_piezo(uc_g, edo.tron_ava, edessdo.Qrin, couleur.magenta_clair) ' vbRed)
    ' '   Call dess_piezo(UC_graphique4, troava, edessdo.Qpluie, vbMagenta)
        If okcharge Then
           Call inter_charge_tr(tr, res_conduit)
            uc_g.dess_lign tr.Absamo, res_conduit.chargeamo, res_conduit.piezointer.X, res_conduit.chargeinter, couleur.rouge, 2
           uc_g.dess_lign res_conduit.piezointer.X, res_conduit.chargeinter, res_conduit.piezointer0.X, res_conduit.chargeinter0, couleur.rouge, 2
           uc_g.dess_lign res_conduit.piezointer0.X, res_conduit.chargeinter0, tr.Absava, res_conduit.chargeava, couleur.rouge, 2
             edptam(3).X = tr.Absava
            edptam(3).Y = res_conduit.chargeava
       End If
        tr = edo.tron_ava
'        tr = edessdo.tron_ava
        res_conduit = calc_debit_tr(tr, qcal)
        res_conduit.zphe_ava = res_conduit.piezoava
        Call inter_piezo_eau(tr, res_conduit)
    
        If okpiezo Then
            uc_g.dess_lign edptam(1).X, edptam(1).Y, res_conduit.zeau_amo.X, res_conduit.piezoamo, couleur.orange, 2
            uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.piezoamo, res_conduit.p_Eau_inter.X, res_conduit.piezointer.Y, couleur.orange, 2
            uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, res_conduit.zeau_ava.X, res_conduit.piezoava, couleur.orange, 2
     '       Call dess_piezo(uc_g, edo.tron_ava, edessdo.Qrin, couleur.magenta_clair) ' vbRed)
        End If
        If okeau Then
             uc_g.dess_lign edptam(2).X, edptam(2).Y, res_conduit.zeau_amo.X, res_conduit.zeau_amo.Y, couleur.bleu, 2
            uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.zeau_amo.Y, res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, couleur.bleu, 2
            uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.piezointer.Y, res_conduit.zeau_ava.X, res_conduit.zeau_ava.Y, couleur.bleu, 2
        End If
        If okcharge Then
            Call inter_charge_tr(tr, res_conduit)
             uc_g.dess_lign edptam(3).X, edptam(3).Y, tr.Absamo, res_conduit.chargeamo, couleur.rouge, 2
            uc_g.dess_lign tr.Absamo, res_conduit.chargeamo, res_conduit.piezointer.X, res_conduit.chargeinter, couleur.rouge, 2
           uc_g.dess_lign res_conduit.piezointer.X, res_conduit.chargeinter, res_conduit.piezointer0.X, res_conduit.chargeinter0, couleur.rouge, 2
           uc_g.dess_lign res_conduit.piezointer0.X, res_conduit.chargeinter0, tr.Absava, res_conduit.chargeava, couleur.rouge, 2
    '       Call dess_charge(uc_g, edo.tron_ava, edessdo.Qrin, couleur.cyan) ' vbCyan)
        End If
    '    Call dess_charge(uc_g, edessdo.tron_amo, edessdo.Qrin, couleur.orange) ' vbCyan)
    
End If
End Sub

Private Sub lect_fich()
Dim za As st_savdo
Dim za1 As st_savdo1
Call funlockb
 
    lhFicDbf = FreeFile
    Cb_deversoir.Clear
    On Error GoTo test_Error
    Open nom_fich For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
   If Not EOF(lhFicDbf) Then
        za = za1.stsavdo
        If Trim(za.type) = nom_type Then
            Cb_deversoir.AddItem (Trim(za.nom))
        End If
   End If
Loop
Close #lhFicDbf
dev_texte = Cb_deversoir.list(0)
Cb_deversoir.Text = Cb_deversoir.list(0)
Cb_deversoir.Refresh
 
 Call flockb(nom_fich)
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + nom_fich + " est déjà en cours d'utilisation.")
    End If
 
Call flockb(nom_fich)
End Sub
Public Sub cre_list_don1()
Dim i As Integer
ReDim list_don1(Tb_debit.count - 1, 3)
For i = 0 To Tb_debit.count - 1
    list_don1(i, 1) = Lb_intdebit(i).Caption
    list_don1(i, 2) = Tb_debit(i).Text
    list_don1(i, 3) = Lb_udebit(i).Caption
Next
End Sub
Public Sub cre_list_don2()
Dim i As Integer
ReDim list_don2(Tb_cont.count - 1, 3)
For i = 0 To Tb_cont.count - 1
    list_don2(i, 1) = Lb_intcont(i).Caption
    list_don2(i, 2) = Tb_cont(i).Text
    list_don2(i, 3) = Lb_ucont(i).Caption
Next
End Sub
Public Sub cre_list_don3()
Dim i As Integer
ReDim list_don3(Tb_amo.count + 2, 7)
    list_don3(0, 2) = "Amont"
    list_don3(0, 4) = "Aval"
    list_don3(0, 6) = "Décharge"

For i = 0 To Tb_amo.count - 1
    list_don3(i + 1, 1) = Lb_intamo(i).Caption
    list_don3(i + 1, 2) = Tb_amo(i).Text
    list_don3(i + 1, 3) = Lb_uamo(i).Caption
    list_don3(i + 1, 4) = Tb_ava(i).Text
    list_don3(i + 1, 5) = Lb_uava(i).Caption
    list_don3(i + 1, 6) = Tb_dech(i).Text
    list_don3(i + 1, 7) = Lb_udech(i).Caption
Next
list_don3(4, 4) = resudev.longetranglee
i = i + 1
    list_don3(i, 1) = Resuintdev.vpsm  'list_int1(0, 1)
    list_don3(i, 2) = resudev.vpsm  'list_int1(0, 2)
    list_don3(i, 3) = Resuudev.vpsm  'list_int1(0, 3)
    list_don3(i, 4) = resudev.vpsv  'list_int1(0, 2)
    list_don3(i, 5) = Resuudev.vpsv  'list_int1(0, 3)
    list_don3(i, 6) = resudev.vpsdech 'list_int1(0, 2)
    list_don3(i, 7) = Resuudev.vpsdech  'list_int1(0, 3)
i = i + 1
    list_don3(i, 1) = Resuintdev.dpsm  'list_int1(0, 1)
    list_don3(i, 2) = resudev.dpsm  'list_int1(0, 2)
    list_don3(i, 3) = Resuudev.dpsm  'list_int1(0, 3)
    list_don3(i, 4) = resudev.dpsv  'list_int1(0, 2)
    list_don3(i, 5) = Resuudev.dpsv  'list_int1(0, 3)
    list_don3(i, 6) = resudev.dpsdech 'list_int1(0, 2)
    list_don3(i, 7) = Resuudev.dpsdech  'list_int1(0, 3)
End Sub
Public Sub cre_list_don4()
Dim i As Integer
Dim cana_ava As conduite
Dim dc As debit_conduit
ReDim list_don4(14, 7)
    list_don4(0, 2) = "Amont"
    list_don4(0, 4) = "Aval"
    list_don4(0, 6) = "Décharge"
i = 1
    list_don4(i, 1) = "---Temps sec---"
    list_don4(i, 2) = ""
    list_don4(i, 3) = ""
    list_don4(i, 4) = ""
    list_don4(i, 5) = ""
    list_don4(i, 6) = ""
    list_don4(i, 7) = ""
i = i + 1
    list_don4(i, 1) = "Débit"
    list_don4(i, 2) = Tb_debit(1).Text
    list_don4(i, 3) = Lb_udebit(1).Caption
    list_don4(i, 4) = Tb_debit(1).Text
    list_don4(i, 5) = Lb_udebit(1).Caption
    list_don4(i, 6) = ""
    list_don4(i, 7) = ""
i = i + 1
    list_don4(i, 1) = "Hauteur d'eau"
    list_don4(i, 2) = resudev.hqtsm
    list_don4(i, 3) = Resuudev.hqtsm
    If Trim(resudev.hqtsv) <> "0.0" Then
        list_don4(i, 4) = resudev.hqtsv
    Else
        list_don4(i, 4) = ajout_zero(Trim(str(txtVersNum(Tb_ava(0).Text) / 1000#)))
    End If
    list_don4(i, 5) = Resuudev.hqtsv
    list_don4(i, 6) = ""
    list_don4(i, 7) = ""
i = i + 1
    list_don4(i, 1) = "Vitesse d'écoulement"
    list_don4(i, 2) = resudev.vqtsm
    list_don4(i, 3) = Resuudev.vqtsm
    If Trim(resudev.vqtsv) <> "0.0" Then
        list_don4(i, 4) = resudev.vqtsv
    Else
        val1 = ((txtVersNum(Tb_ava(0).Text) / 2000) ^ 2) * pi
        list_don4(i, 4) = ajout_zero(Trim(str(Round(txtVersNum(Tb_debit(1).Text) / 1000# / val1, 2))))
    End If
    list_don4(i, 5) = Resuudev.vqtsv
    list_don4(i, 6) = ""
    list_don4(i, 7) = ""
i = i + 1
    list_don4(i, 1) = "---Rinçage---"
    list_don4(i, 2) = ""
    list_don4(i, 3) = ""
    list_don4(i, 4) = ""
    list_don4(i, 5) = ""
    list_don4(i, 6) = ""
    list_don4(i, 7) = ""
i = i + 1
    list_don4(i, 1) = "Débit"
    list_don4(i, 2) = Tb_debit(2).Text
    list_don4(i, 3) = Lb_udebit(2).Caption
    list_don4(i, 4) = Tb_debit(2).Text
    list_don4(i, 5) = Lb_udebit(2).Caption
    list_don4(i, 6) = ""
    list_don4(i, 7) = ""
i = i + 1
    list_don4(i, 1) = "Hauteur d'eau"
    list_don4(i, 2) = resudev.hqrinm
    list_don4(i, 3) = Resuudev.hqrinm
    If Trim(resudev.hqrinv) <> "0.0" Then
        list_don4(i, 4) = resudev.hqrinv
    Else
        list_don4(i, 4) = ajout_zero(Trim(str(txtVersNum(Tb_ava(0).Text) / 1000#)))
    End If
    list_don4(i, 5) = Resuudev.hqrinv
    list_don4(i, 6) = ""
    list_don4(i, 7) = ""
i = i + 1
    list_don4(i, 1) = "Vitesse d'écoulement"
    list_don4(i, 2) = resudev.vqrinm
    list_don4(i, 3) = Resuudev.vqrinm
    If Trim(resudev.vqrinv) <> "0.0" Then
        list_don4(i, 4) = resudev.vqrinv
    Else
        val1 = ((txtVersNum(Tb_ava(0).Text) / 2000#) ^ 2) * pi
        list_don4(i, 4) = ajout_zero(Trim(str(Round(txtVersNum(Tb_debit(2).Text) / 1000# / val1, 2))))
    End If
    list_don4(i, 5) = Resuudev.vqrinv
    list_don4(i, 6) = ""
    list_don4(i, 7) = ""
i = i + 1
    list_don4(i, 1) = "---Pluie---"
    list_don4(i, 2) = ""
    list_don4(i, 3) = ""
    list_don4(i, 4) = ""
    list_don4(i, 5) = ""
    list_don4(i, 6) = ""
    list_don4(i, 7) = ""
i = i + 1
    list_don4(i, 1) = "Débit"
    list_don4(i, 2) = Tb_debit(0).Text
    list_don4(i, 3) = Lb_udebit(0).Caption
    list_don4(i, 4) = resudev.debetranglee
    list_don4(i, 5) = Resuudev.debetranglee
    list_don4(i, 6) = resudev.debdeverse
    list_don4(i, 7) = Resuudev.debdeverse
i = i + 1
    list_don4(i, 1) = "Hauteur d'eau amont"
    list_don4(i, 2) = resudev.hqpluiem
    list_don4(i, 3) = Resuudev.hqpluiem
    cana_ava = edessdo.tron_ava.conduit
    cana_ava.Longueur = txtVersNum(resudev.longetranglee)
    
'    Call cana(cana_ava, ct)
'    ltc = calc_par(cana_ava)
'    qvi = caltran1(txtVersNum(resudev.debetranglee) * 1000#, ct, ltc)
'    list_don4(i, 4) = ajout_zero(Trim(Str(Round(qvi(5), 3))))
    dc = calc_debit_tr(edessdo.tron_ava, txtVersNum(resudev.debetranglee))
    list_don4(i, 4) = dc.hauteur
    list_don4(i, 5) = Resuudev.hqpluiev
    list_don4(i, 6) = resudev.hqdev
    list_don4(i, 7) = Resuudev.hqdev
i = i + 1
    list_don4(i, 1) = "Vitesse d'écoulement amont"
    list_don4(i, 2) = resudev.vqpluiem
    list_don4(i, 3) = Resuudev.vqpluiem
'    list_don4(i, 4) = ajout_zero(Trim(Str(Round(qvi(2), 3))))
    list_don4(i, 4) = ajout_zero(Trim(str(Round(dc.vitesse, 3))))
    list_don4(i, 5) = Resuudev.vqpluiev
    list_don4(i, 6) = resudev.vqdev
    list_don4(i, 7) = Resuudev.vqdev
i = i + 1
    list_don4(i, 1) = "Hauteur d'eau aval"
    list_don4(i, 2) = resudev.hqpluiemav
    list_don4(i, 3) = Resuudev.hqpluiemav
    list_don4(i, 4) = ""
    list_don4(i, 5) = ""
    list_don4(i, 6) = resudev.hqdevav
    list_don4(i, 7) = Resuudev.hqdevav
i = i + 1
    list_don4(i, 1) = "Vitesse d'écoulement aval"
    list_don4(i, 2) = resudev.vqpluiemav
    list_don4(i, 3) = Resuudev.vqpluiemav
    list_don4(i, 4) = ""
    list_don4(i, 5) = ""
    list_don4(i, 6) = resudev.vqdevav
    list_don4(i, 7) = Resuudev.vqdevav
End Sub
Public Sub cre_list_don5()
ReDim list_don5(7, 5)
    list_don5(0, 1) = Resuintdev.Ldev
    list_don5(0, 2) = ""
    list_don5(0, 3) = ""
    list_don5(0, 4) = resudev.Ldev
    list_don5(0, 5) = Resuudev.Ldev
    list_don5(1, 1) = Resuintdev.Hcret
    list_don5(1, 2) = ""
    list_don5(1, 3) = ""
    list_don5(1, 4) = resudev.Hcret
    list_don5(1, 5) = Resuudev.Hcret
    list_don5(2, 1) = Resuintdev.Pdev
    list_don5(2, 2) = ""
    list_don5(2, 3) = ""
    list_don5(2, 4) = resudev.Pdev
    list_don5(2, 5) = Resuudev.Pdev
    list_don5(3, 1) = ""
    list_don5(3, 2) = ""
    list_don5(3, 3) = ""
    list_don5(3, 4) = ""
    list_don5(3, 5) = ""
    list_don5(4, 1) = Resuintdev.Tram
    list_don5(4, 2) = ""
    list_don5(4, 3) = ""
    list_don5(4, 4) = resudev.Tram
    list_don5(4, 5) = Resuudev.Tram
    list_don5(5, 1) = ""
    list_don5(5, 2) = "Amont"
    list_don5(5, 3) = ""
    list_don5(5, 4) = "Aval"
    list_don5(5, 5) = ""
    list_don5(6, 1) = Resuintdev.Ham
    list_don5(6, 2) = resudev.Ham
    list_don5(6, 3) = Resuudev.Ham
    list_don5(6, 4) = resudev.Hav
    list_don5(6, 5) = Resuudev.Hav
    list_don5(7, 1) = Resuintdev.Haam
    list_don5(7, 2) = resudev.Haam
    list_don5(7, 3) = Resuudev.Haam
    list_don5(7, 4) = resudev.Haav
    list_don5(7, 5) = Resuudev.Haav
End Sub
Public Sub cre_list_don6()
ReDim list_don6(1, 5)
    list_don6(0, 1) = Resuintdev.regime + " " + resudev.regime
    list_don6(0, 2) = ""
    list_don6(0, 3) = ""
    list_don6(0, 4) = ""
    list_don6(0, 5) = ""
    list_don6(1, 1) = ""
    list_don6(1, 2) = ""
    list_don6(1, 3) = ""
    list_don6(1, 4) = ""
    list_don6(1, 5) = ""
'    list_don6(1, 1) = ""
'    list_don6(1, 2) = "Amont"
'    list_don6(1, 3) = ""
'    list_don6(1, 4) = "Aval"
'    list_don6(1, 5) = ""
'    list_don6(2, 1) = Resuintdev.Ham
'    list_don6(2, 2) = resudev.Ham
'    list_don6(2, 3) = Resuudev.Ham
'    list_don6(2, 4) = resudev.Hav
'    list_don6(2, 5) = Resuudev.Hav
'    list_don6(3, 1) = Resuintdev.Haam
'    list_don6(3, 2) = resudev.Haam
'    list_don6(3, 3) = Resuudev.Haam
'    list_don6(3, 4) = resudev.Haav
'    list_don6(3, 5) = Resuudev.Haav
End Sub
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
Case Is = "list_don6"
    lect_list = list_don6
Case Is = "list_int1"
    lect_list = list_int1
Case Is = "list_resu1"
    lect_list = list_resu1
End Select
End Function
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
ok_imp = False
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
resudev.Tram = Me.Tb_dev(3)
'edo_res.Tram = txtVersNum(Me.tb_dev(3))
edo_res.c = 1#
sresult = verif_do_charge(edo, edo_res, edessdo.tron_amo, edessdo.Qpluie / 1000#, edessdo.Qrin / 1000#, owner.fdessin.UC_graphique1)
'Me.Lb_longce(1).Caption = sresult
'Me.SSTab_result.Tab = 1
Me.Txtb_deversoir.Text = sresult + Chr(13) + Chr(10) + message

Me.Frame4.Visible = True
'Me.Frame4.Enabled = False

If message <> "" Then
    
'Me.Txtb_deversoir.Text = sresult + Chr(13) + Chr(10) + message
    reponse = MsgBox(message, 0, "Vérification du déversoir")
    Me.Chk_Qrin.Value = 0
    Me.Chk_Qts.Value = 0
    Me.Chk_Qpluie.Enabled = True
    Me.Chk_charge.Value = 1
    Me.Chk_piezo.Value = 1
    Me.Chk_eau.Value = 1
    Me.Chk_Qpluie = 1
'julienne
    Me.SSTab1.TabEnabled(5) = True 'False
    Me.Cmd_VerifDo.Enabled = False
    Me.Cmd_resudech.Enabled = True 'False
    Me.Txtb_decharge.Text = ""
Else
    Me.Cmd_VerifDo.Enabled = False
    Me.Chk_Qrin.Value = 0
    Me.Chk_Qts.Value = 0
    Me.Chk_Qpluie.Enabled = True
    Me.Chk_charge.Value = 1
    Me.Chk_piezo.Value = 1
    Me.Chk_eau.Value = 1
    Me.Chk_Qpluie = 1
    Me.SSTab1.TabEnabled(5) = True
    Me.Txtb_decharge.Text = ""
    If edessdo.tron_dech.conduit.Diametre > 0 And edessdo.tron_dech.conduit.Longueur > 0 _
        And edessdo.tron_dech.conduit.pente > 0 And edessdo.tron_dech.conduit.rugosite > 0 Then
        Me.Cmd_resudech.Enabled = True
        Me.Cmd_dech.Enabled = True
    End If
End If
'Me.Frame4.Visible = True
Me.Frame4.Enabled = True
Call OK_lignes_Click
End Sub
Private Sub m_quitter_Click()
    Unload owner
End Sub

Private Sub Form_Activate()
    change_coul = False
'    owner.affich_aide Me.Name, mes_prec
End Sub

Private Sub Form_Click()
    owner.affich_aide Me.Name, ""  'Déversoir d'orage"
    Change_Couleur "Me", 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
owner.fcom.Form_KeyAide KeyCode, Shift
Me.SetFocus
End Sub

Private Sub Frm_bv_Click()
Dim mes As String
Dim nom As String
nom = "Frm_bv"
mes = Rec_Mes(nom, 0)
owner.affich_aide Me.Name, mes
Change_Couleur nom, 0
Change_Focus nom, 0

End Sub


Private Sub Frm_condam_Click()
Dim mes As String
Dim nom As String
nom = "Frm_condam"
mes = Rec_Mes(nom, 0)
owner.affich_aide Me.Name, mes
Change_Couleur nom, 0
Change_Focus nom, 0

End Sub


Private Sub Frm_condav_Click()
Dim mes As String
Dim nom As String
nom = "Frm_condav"
mes = Rec_Mes(nom, 0)
owner.affich_aide Me.Name, mes
Change_Couleur nom, 0
Change_Focus nom, 0

End Sub

Private Sub Frm_dech_Click()
Dim mes As String
Dim nom As String
nom = "Frm_dech"
mes = Rec_Mes(nom, 0)
owner.affich_aide Me.Name, mes
Change_Couleur nom, 0
Change_Focus nom, 0

End Sub

Private Sub Frm_dev_Click()
Dim mes As String
Dim nom As String
nom = "Frm_dev"
mes = Rec_Mes(nom, 0)
owner.affich_aide Me.Name, mes
Change_Couleur nom, 0
Change_Focus nom, 0

End Sub


Private Sub Lb_intamo_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_intamo"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_intava_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_intava"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_intcont_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_intcont"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_intdebit_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_intdebit"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_intdech_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_intdech"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_intdev_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Lb_intdev"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes
End Sub

Private Sub mnufichier_Click()
    If ouv_sauve Or save_fich Then
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
    reponse = MsgBox("Le déversoir n'a pas été enregistré" + Chr(10) _
        + "Voulez vous le sauvegarder?", 3, "Sauvegarde du déversoir")
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
    reponse = MsgBox("Le déversoir n'a pas été enregistré" + Chr(10) _
        + "Voulez vous le sauvegarder?", 3, "Sauvegarde d'un déversoir")
    Select Case reponse
        Case Is = 6  ' 6=oui,7=non,2=annuler
            Call mnusave_Click
'            Cb_deversoir.Visible = True
            frmf.Label1.Caption = "Recherche d'un déversoir "
            frmf.Caption = nom
            frmf.Show 1
            If frmf.nomfich <> "" Then
                Me.nom_ouvrage = frmf.nomfich
                Call Me.Cb_deversoir_click
            End If
       Case Is = 7
'            Cb_deversoir.Visible = True
            frmf.Label1.Caption = "Recherche d'un déversoir "
            frmf.Caption = nom
            frmf.Show (1)
            If frmf.nomfich <> "" Then
                Me.nom_ouvrage = frmf.nomfich
                Call Me.Cb_deversoir_click
            End If
    End Select
Else
'    Cb_deversoir.Visible = True
    frmf.Label1.Caption = "Recherche d'un déversoir "
    frmf.Caption = nom
    frmf.Show 1
    If frmf.nomfich <> "" Then
        Me.nom_ouvrage = frmf.nomfich
        Call Me.Cb_deversoir_click
    End If
End If
Set frmf = Nothing
End Sub

Private Sub mnuprint_Click()
Dim pict1 As New StdPicture
Dim pict2 As New StdPicture
Dim i As Integer, nb As Integer, j As Integer
Dim val1 As Double
Dim ltc() As Variant, ct() As Variant, qvm(5) As Variant
'modif FO   ' If ProtectCheck(2) <> 0 Then End
If ok_imp Then
Me.MousePointer = 11
FrmPrint.Type1 = "deversoir"
FrmPrint.nomobjet = Tb_titre.Text
FrmPrint.titre1 = "FICHE HYDRAULIQUE DEVERSOIR d' ORAGE"
FrmPrint.sstitre1 = "Caractéristiques " + Frm_bv.Caption
Frm_imp.Type1 = "deversoir"
Frm_imp.nomobjet = Tb_titre.Text
Frm_imp.titre1 = "FICHE HYDRAULIQUE DEVERSOIR d' ORAGE"
Frm_imp.sstitre1 = "Caractéristiques " + Frm_bv.Caption
FrmPrint.ssTitre2 = "Contraintes"
Frm_imp.ssTitre2 = "Contraintes"
FrmPrint.ssTitre3 = "Conduites"
Frm_imp.ssTitre3 = "Conduites"
FrmPrint.ssTitre4 = "Résultats de fonctionnement"
Frm_imp.ssTitre4 = "Résultats de fonctionnement"
FrmPrint.ssTitre5 = "Déversoir"
Frm_imp.ssTitre5 = "Déversoir"
FrmPrint.ssTitre6 = "Décharge"
Frm_imp.ssTitre6 = "Décharge"
cre_list_don1
cre_list_don2
cre_list_don3
cre_list_don4
cre_list_don5
cre_list_don6
Call dess_do_print(Frm_desprint.UC_graphique1, False, False, True) 'okcharge, okpiezo, okeau
Call dess_dech_print(Frm_desprint.UC_graphique2)

Set pict1 = Frm_desprint.UC_graphique1.lire_pict1()
Set pict2 = Frm_desprint.UC_graphique2.lire_pict1()
FrmPrint.paint_picture pict1
FrmPrint.paint_picture2 pict2
SavePicture pict1, chemin_app + "dess.bmp"
SavePicture pict2, chemin_app + "dess1.bmp"
''FrmPrint.Show
Frm_imp.Show 1
Me.MousePointer = 1
End If
End Sub

Private Sub mnusave_Click()
'modif FO   ' If ProtectCheck(2) <> 0 Then End
    If Trim(Tb_titre.Text) <> "" And fich_lect = nom_fich Then
        Call save(False)
    Else
        Call mnusaves_Click
    End If
End Sub
Public Sub save(ByVal bsous As Boolean)
Dim za As st_savdo
Dim za1 As st_savdo1
Dim i As Integer, isave As Integer
Dim reponse As Integer
 
If Trim(Tb_titre.Text) <> "" Then
    If (Trim(ebv.Qchoisi) = "CAQUOT" And edessdo.Qpluie <> Round(ebv.Qcor * 1000, 1)) Or _
        (Trim(ebv.Qchoisi) = "RATION" And edessdo.Qpluie <> Round(ebv.Qmr * 1000, 1)) Or _
        (Trim(ebv.Qchoisi) = "HYDROG" And edessdo.Qpluie <> Round(ebv.Qhydro * 1000, 1)) Or _
        edessdo.Qts <> Round(ebv.Qts, 1) Or edessdo.Qrin <> Round(ebv.Qrin, 1) Then
           nombassin = ""
    End If

    Call funlockb
    edessdo.nombv = nombassin
    edessdo.nom = ebv.nom
    lhFicDbf = FreeFile
    On Error GoTo test_Error
    Open nom_fich For Random Access Read Write Lock Read Write As #lhFicDbf Len = Len(za1)
    i = 0
    isave = 0
    Do While Not EOF(lhFicDbf)
        Get #lhFicDbf, , za1
        If Not EOF(lhFicDbf) Then
            i = i + 1
            za = za1.stsavdo
            If Trim(za.type) = nom_type And Trim(za.nom) = Trim(Tb_titre.Text) Then
                isave = i
            End If
       End If
    Loop
    If isave > 0 Then
        If bsous Then
           reponse = MsgBox("Le nom est déjà utilisé. Le remplacer?", 4, "Sauvegarde d'un déversoir")
           Else
           reponse = 6
        End If
        If reponse = 6 Then
            esave.type = "deversoir"
            esave.nom = Tb_titre.Text
            esave.edessdo = edessdo
            esave.edo = edo
            za1.stsavdo = esave
            Put #lhFicDbf, isave, za1
            ouv_sauve = False
            save_fich = True
            fich_lect = nom_fich
       Else
            Unload Frm_titre
            Call mnusaves_Click
        End If
    Else
        esave.type = "deversoir"
        esave.nom = Tb_titre.Text
        esave.edessdo = edessdo
        esave.edo = edo
        za1.stsavdo = esave
        FileLength = LOF(lhFicDbf) / Len(za1) + 1
        Put #lhFicDbf, FileLength, za1
        ouv_sauve = False
        save_fich = True
        fich_lect = nom_fich
    End If
        Close #lhFicDbf
        Call flockb(nom_fich)
        Call lect_fich
        dev_texte = Trim(Tb_titre.Text)
        Cb_deversoir.Text = Trim(Tb_titre.Text)
Else
    reponse = MsgBox("Le nom du déversoir n'est pas renseigné.", , "Sauvegarde d'un déversoir")
End If
 
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + nom_fich + " est déjà en cours d'utilisation.")
    End If
 
Call flockb(nom_fich)
End Sub

Private Sub Cmd_resum_Click()
Dim ok As Boolean ', ok1 As Boolean
    ok_imp = False
    ok = calcul_amont
    Cmd_resum.Enabled = True
    Cmd_resum.SetFocus
    Call key13(Me)
    Cmd_resum.Enabled = False
     Cmd_annulm.Enabled = True
   ouv_sauv = True
End Sub
Private Sub ini_longce(Optional ByVal i As Integer)
If i = 0 Then
'    Me.Lb_longce(0).BackColor = &H80000005   '&H8000000B
    Me.Lb_longce(0).BorderStyle = 0
    Me.Lb_longce(0).Caption = ""
'    Me.Lb_longce(1).BackColor = &H80000005    ' &H8000000B
    Me.Lb_longce(1).BorderStyle = 0
    Me.Lb_longce(1).Caption = ""
    Txtb_deversoir.Text = ""
    Else
'    Me.Lb_longce(i - 1).BackColor = &H80000005  '&H8000000B
    Me.Lb_longce(i - 1).BorderStyle = 0
    Me.Lb_longce(i - 1).Caption = ""
    Txtb_deversoir.Text = ""
End If
End Sub
Private Sub modi_longce(Optional ByVal i As Integer)
If i = 0 Then
'    Me.Lb_longce(0).BackColor = &H80000005
    Me.Lb_longce(0).BorderStyle = 1
'    Me.Lb_longce(1).BackColor = &H80000005
    Me.Lb_longce(1).BorderStyle = 1
    Else
'    Me.Lb_longce(i - 1).BackColor = &H80000005
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
Private Function calc_amont() As Boolean
Dim Qts As Double, Qrin As Double, Qpluie As Double
Dim qv As deb_vit
Dim ltc() As Variant, ct() As Variant, qvm(5) As Variant
Dim qcal As Double
Dim ecoulam As String, betam As Double
Dim message As String

Dim cana_amo As conduite
calc_amont = False
cana_amo = edessdo.tron_amo.conduit
'wagner Call calcul_condam(cana_amo)

Qts = edessdo.Qts
Qrin = edessdo.Qrin + edessdo.Qts
Qpluie = edessdo.Qpluie

qv = debvit_ps(cana_amo)
Me.Lb_vpsm.Caption = "Vitesse pleine section = " + ajout_zero(Trim(str(Round(qv.vitesse, 3)))) + " m/s"
Me.Lb_dpsm.Caption = "Débit pleine section = " + ajout_zero(Trim(str(Round(qv.debit, 3)))) + " m3/s"
resudev.vpsm = ajout_zero(Trim(str(Round(qv.vitesse, 3))))
resudev.dpsm = ajout_zero(Trim(str(Round(qv.debit, 3))))
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
    Me.Lb_vqrinm.Caption = "Vitesse d'écoulement à QREF = " + "   " + " m/s"
    Me.Lb_hqrinm.Caption = "Hauteur d'eau QREF = " + "   " + " m"
    Me.Lb_vqpluiem.Caption = "Vitesse d'écoulement à QORA = " + "   " + " m/s"
    Me.Lb_hqpluiem.Caption = "Hauteur d'eau QORA = " + "   " + " m"

Else
    qcal = Qts
    Call cana(cana_amo, ct)
    ltc = calc_par(cana_amo)
    qvi = caltran1(qcal, ct, ltc)
    Me.Lb_vqtsm.Caption = "Vitesse d'écoulement à QTS = " + ajout_zero(Trim(str(Round(qvi(2), 3)))) + " m/s"
    Me.Lb_hqtsm.Caption = "Hauteur d'eau QTS = " + ajout_zero(Trim(str(Round(qvi(5), 3)))) + " m"
    resudev.vqtsm = ajout_zero(Trim(str(Round(qvi(2), 3))))
    resudev.hqtsm = ajout_zero(Trim(str(Round(qvi(5), 3))))
    qcal = Qrin
    Call cana(cana_amo, ct)
    ltc = calc_par(cana_amo)
    qvi = caltran1(qcal, ct, ltc)
    Me.Lb_vqrinm.Caption = "Vitesse d'écoulement à QREF = " + ajout_zero(Trim(str(Round(qvi(2), 3)))) + " m/s"
    Me.Lb_hqrinm.Caption = "Hauteur d'eau QREF = " + ajout_zero(Trim(str(Round(qvi(5), 3)))) + " m"
    resudev.vqrinm = ajout_zero(Trim(str(Round(qvi(2), 3))))
    resudev.hqrinm = ajout_zero(Trim(str(Round(qvi(5), 3))))
    qcal = Qpluie
    Call cana(cana_amo, ct)
    ltc = calc_par(cana_amo)
    qvi = caltran1(qcal, ct, ltc)
    Me.Lb_vqpluiem.Caption = "Vitesse d'écoulement à QORA = " + ajout_zero(Trim(str(Round(qvi(2), 3)))) + " m/s"
    Me.Lb_hqpluiem.Caption = "Hauteur d'eau QORA = " + ajout_zero(Trim(str(Round(qvi(5), 3)))) + " m"
    resudev.vqpluiem = ajout_zero(Trim(str(Round(qvi(2), 3))))
    resudev.hqpluiem = ajout_zero(Trim(str(Round(qvi(5), 3))))
' calcul du regime amont
 ' a revoir
        betam = angle(Qpluie / (qv.debit * 1000))
       betam = beta
        ecoulam = calcul_ecoul(Qpluie / 1000, cana_amo.Diametre, betam)
        If ecoulam = "TORREN." Then
            MsgBox "Ecoulement a débit de pointe Torrentiel !" + Chr(13) + "Diminnuez la pente ou prevoir un ressaut ", vbOKOnly, "Vérification d'écoulement"
        End If
 ''''' Me.SSTab1.TabEnabled(3) = True
  calc_amont = True
End If

End Function

Private Function calcul_amont() As Boolean
Dim troamo As troncon, troava As troncon
Dim cana_amo As conduite
Dim cana_ava As conduite
    ok_tab2 = True

    calcul_amont = False
' conduite amont -> troncon amont
    cana_amo.Diametre = edessdo.dam / 1000#
    cana_amo.Longueur = edessdo.Lam
    cana_amo.pente = edessdo.iRadam / 10000#
    cana_amo.rugosite = edessdo.Kam
'    resudev.dam = Str(edessdo.dam)
'    resudev.Lam = Str(edessdo.Lam)
'    resudev.iRadam = Str(edessdo.iRadam)
'    resudev.Kam = Str(edessdo.Kam)
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
    calcul_amont = calc_amont
    
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
        
'        Me.Cmd_resuv.Enabled = True
End If
' dessin a partir de canal amont
 Call dessin_amont
' reinitialisation des
    Call reini_form(21)
'    Call ini_longce
'    Me.Cmd_annulm.Visible = False
'    Me.Cmd_resum.Enabled = False
'    Me.Cmd_resudo.Enabled = True
'        Me.Frame4.Enabled = False
'        Me.Frame4.Visible = False
If troamo.radava <= edessdo.rdoav Then
    MsgBox "cote aval canalisation amont inférieure à cote radier obligé aval", vbOKOnly
End If
End Function
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
Dim ok As Boolean
    ok_imp = False
    ok = calcul_aval
    Cmd_resuv.Enabled = True
    Cmd_resuv.SetFocus
    Call key13(Me)
    Cmd_resuv.Enabled = False
    Cmd_annulv.Enabled = True
   ouv_sauv = True
End Sub
Private Function calc_aval() As Boolean
Dim qv As deb_vit
Dim ltc() As Variant, ct() As Variant, qvm(5) As Variant
Dim qcal As Double
Dim Qts As Double, Qrin As Double, Qpluie As Double
Dim ecoulam As String, betam As Double
Dim message As String
Dim cana_ava As conduite
calc_aval = True
cana_ava = edessdo.tron_ava.conduit

Qts = edessdo.Qts
Qrin = edessdo.Qrin + edessdo.Qts
Qpluie = edessdo.Qpluie

qv = debvit_ps(cana_ava)
Me.Lb_Vpsv.Caption = "Vitesse pleine section = " + ajout_zero(Trim(str(Round(qv.vitesse, 3)))) + " m/s"
Me.Lb_Dpsv.Caption = "Débit pleine section = " + ajout_zero(Trim(str(Round(qv.debit, 3)))) + " m3/s"
resudev.vpsv = ajout_zero(Trim(str(Round(qv.vitesse, 3))))
resudev.dpsv = ajout_zero(Trim(str(Round(qv.debit, 3))))
If Qts > qv.debit * 1000 Then
    Me.Lb_Vqtsv.Caption = "Vitesse d'écoulement à QTS = " + "   " + " m/s"
    Me.Lb_Hqtsv.Caption = "Hauteur d'eau QTS = " + "   " + " m"
    calc_aval = False
Else

    qcal = Qts
    Call cana(cana_ava, ct)
    ltc = calc_par(cana_ava)
    qvi = caltran1(qcal, ct, ltc)
    Me.Lb_Vqtsv.Caption = "Vitesse d'écoulement à QTS = " + ajout_zero(Trim(str(Round(qvi(2), 3)))) + " m/s"
    Me.Lb_Hqtsv.Caption = "Hauteur d'eau QTS = " + ajout_zero(Trim(str(Round(qvi(5), 3)))) + " m"
    resudev.vqtsv = ajout_zero(Trim(str(Round(qvi(2), 3))))
    resudev.hqtsv = ajout_zero(Trim(str(Round(qvi(5), 3))))
End If
If Qrin > qv.debit * 1000 Then
    Me.Lb_Vqrinv.Caption = "Vitesse d'écoulement à QREF = " + "   " + " m/s"
    Me.Lb_Hqrinv.Caption = "Hauteur d'eau QREF = " + "   " + " m"
Else
    qcal = Qrin
    Call cana(cana_ava, ct)
    ltc = calc_par(cana_ava)
    qvi = caltran1(qcal, ct, ltc)
    Me.Lb_Vqrinv.Caption = "Vitesse d'écoulement à QREF = " + ajout_zero(Trim(str(Round(qvi(2), 3)))) + " m/s"
    Me.Lb_Hqrinv.Caption = "Hauteur d'eau QREF = " + ajout_zero(Trim(str(Round(qvi(5), 3)))) + " m"
    resudev.vqrinv = ajout_zero(Trim(str(Round(qvi(2), 3))))
    resudev.hqrinv = ajout_zero(Trim(str(Round(qvi(5), 3))))
End If
If Qpluie > qv.debit * 1000 Then
    Me.Lb_Vqpluiev.Caption = "Vitesse d'écoulement à QORA = " + "   " + " m/s"
    Me.Lb_Hqpluiev.Caption = "Hauteur d'eau QORA = " + "   " + " m"
Else
    qcal = Qpluie
    Call cana(cana_ava, ct)
    ltc = calc_par(cana_ava)
    qvi = caltran1(qcal, ct, ltc)
    Me.Lb_Vqpluiev.Caption = "Vitesse d'écoulement à QORA = " + ajout_zero(Trim(str(Round(qvi(2), 3)))) + " m/s"
    Me.Lb_Hqpluiev.Caption = "Hauteur d'eau QORA = " + ajout_zero(Trim(str(Round(qvi(5), 3)))) + " m"
    resudev.vqpluiev = ajout_zero(Trim(str(Round(qvi(2), 3))))
    resudev.hqpluiev = ajout_zero(Trim(str(Round(qvi(5), 3))))
End If

End Function

Private Function calcul_aval() As Boolean
Dim troava As troncon, troamo As troncon
Dim cana_ava As conduite
Dim ok As Boolean
ok = False
    ok_tab3 = True

' conduite aval -> troncon aval
    cana_ava.Diametre = edessdo.dav / 1000#
    cana_ava.Longueur = edessdo.Lav
    cana_ava.pente = edessdo.iradav / 10000#
    cana_ava.rugosite = edessdo.kav
'    resudev.dav = edessdo.dav
'    resudev.Lav = edessdo.Lav
'    resudev.iradav = edessdo.iradav
'    resudev.kav = edessdo.kav
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
  ok = calc_aval
  If ok Then
'dessin des lignes hydrauliques aval
    Call dess_piezo(owner.fdessin.UC_graphique1, troamo, edessdo.Qts, couleur.bleu) ' vbBlue)
    Call dess_piezo(owner.fdessin.UC_graphique1, troamo, edessdo.Qrin, couleur.rouge) ' vbRed)
    Call dess_piezo(owner.fdessin.UC_graphique1, troamo, edessdo.Qpluie, couleur.magenta) ' vbMagenta)
    Call dess_piezo(owner.fdessin.UC_graphique1, troava, edessdo.Qts, couleur.bleu) ' vbBlue)
    Call dess_piezo(owner.fdessin.UC_graphique1, troava, edessdo.Qrin, couleur.rouge) ' vbRed)
    Call dess_piezo(owner.fdessin.UC_graphique1, troava, edessdo.Qpluie, couleur.magenta) ' vbMagenta)
'reinitailisation des autres onglets
Call reini_form(31)
'    Call ini_longce
'    Me.SSTab1.TabEnabled(4) = True
'    Me.Cmd_annulv.Visible = False
'    Me.Cmd_resuv.Enabled = False
'    If edo.Longueur > 0 And edo.hauteur > 0 And edo.pente > 0 Then
'        Me.Cmd_resudo.Enabled = True
'    End If
'     Me.Frame4.Enabled = False
'     Me.Frame4.Visible = False
 End If
'If troava.radamo > troamo.radava Then
'    MsgBox "cote amont canalisation aval  supérieure à cote aval canalisation amont", vbOKOnly
'End If
calcul_aval = ok
End Function

Private Sub Cmd_Sel_Bv_Click()
     Dim pict1 As New StdPicture
 Dim mes As String
Dim nom As String
nom = "Cmd_Sel_Bv"
mes = Rec_Mes(nom, 0)
Change_Focus nom, 0
owner.affich_aide Me.Name, mes
   dess_anc = chemin_app + "dessanc.bmp"
    If Dir(dess_anc) <> "" Then
        Kill dess_anc
    End If
    Set pict1 = owner.fdessin.UC_graphique1.lire_pict1()
    SavePicture pict1, chemin_app + "dessanc.bmp"
    owner.fdessin.UC_graphique1.graphique_clear
    owner.fdessin.UC_graphiqueB.Visible = False
    owner.fdessin.Image3.Visible = False
    owner.fdessin.UC_graphique1.Visible = True
    owner.fdessin.UC_graphique2.Visible = False
    Me.Enabled = False
    do_bv = True
    Set owner.fbassin = New Frm_bv2
    owner.fbassin.Show
    owner.fbassin.nom_ouvrage = nombassin
    owner.fbassin.Cmd_retour.Visible = True
    owner.fbassin.Cmd_retour.Caption = "Retour au déversoir"
    fich_lect = nom_fich
    Call owner.fbassin.rec_bassin_versant
    owner.affich_aide owner.fbassin.Name, "Module" ' "Calcul de débit de bassin versant"
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
maxX = edessdo.tron_amo.conduit.Longueur + edo.Longueur + maximum(edessdo.lgca, edo.tron_ava.conduit.Longueur)
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
Dim list_centon() As Variant
list_centon = Array("0.75", "1.00", "1.25")
For i = 0 To UBound(list_centon)
    Cb_centon.AddItem (list_centon(i))
Next
    okg = True
    Me.KeyPreview = True
    Call ini_tooltip_do(Me)
    nom_ouvrage = ""
    ouv_sauve = False
    save_fich = False
    nom_dessin = chemin_app + "do_bassin.bmp"
'    nom_fich = chemin_app + "deversoir.bin"
'    nom_fich = chemin_app + "etude.bin"
    nom_type = "deversoir"
    fen_titre = Me.Caption
    Set owner = MDIFrm_menu.rec_owner
    Call retaille
'    owner.affich_aide Me.Name, "Deversoir"
    Cb_deversoir.Visible = False
    Frm_desprint.Show
    Frm_desprint.Visible = False
    nombassin = ""
    If Dir(nom_fich) <> "" Then
        Call lect_fich
    End If
    Call debut
End Sub
Private Sub debut0()
    Cb_deversoir.Text = ""
    Tb_titre.Text = ""
    Me.Caption = fen_titre
    Me.Lab_bas = ""
'    ouv_sauve = False
    Call debut
End Sub
Private Sub debut()
    bKP = False
    sval_champ = ""
    Call init_l_tab
    ok_imp = False
    Call ini_resuintdv
    Me.SSTab1.Tab = 0
    ok_tab1 = False
    ok_tab2 = False
    ok_tab3 = False
    ok_tab4 = False
    ok_tab5 = False
    Call ini_bv
    nombassin = ""
    Call ini_edessdo
    Call reini_form(0)
    Call init_graphique
    Call ini_form
    Cb_centon.Text = "0.75"
    Tb_dech(4).Text = Cb_centon.Text
    Me.Tb_debit(0) = "0.0"
    Me.Tb_debit(2) = "0.0"
    Me.Tb_debit(1) = "0.0"
    ouv_sauve = False
    save_fich = False
    fich_lect = ""
    change_coul = False
End Sub
Private Sub init_graphique()
    owner.fdessin.mnu_fichier.Caption = Me.mnufichier.Caption
    owner.fdessin.Image1.Visible = False
    owner.fdessin.Image2.Visible = False
    owner.fdessin.Image3.Visible = True
    owner.fdessin.UC_graphiqueB.Visible = False
    owner.fdessin.UC_graphique1.Visible = False
    owner.fdessin.UC_graphique2.Visible = False
'    owner.fdessin.UC_graphiqueB.graphique_clear
 '   owner.fdessin.UC_graphiqueB.init_fond nom_dessin
    owner.fdessin.Image3.Picture = LoadPicture(nom_dessin)
    owner.fdessin.UC_graphique1.graphique_clear
    owner.fdessin.UC_graphique1.reinit 7, "Arial"
    owner.fdessin.UC_graphique1.init_title
    owner.fdessin.UC_graphique1.init_titleh ""
    owner.fdessin.UC_graphique1.init_titleb ""
'owner.fdessin.UC_graphique1.Top = 0
'owner.fdessin.UC_graphique1.Left = 350 '240
'owner.fdessin.UC_graphique1.Height = 4400 '4210
'owner.fdessin.UC_graphique1.Width = 10000 '9855
'owner.fdessin.UC_graphiqueB.reinit 7, "Arial"
'owner.fdessin.UC_graphiqueB.init_title
'owner.fdessin.UC_graphiqueB.init_titleh ""
'owner.fdessin.UC_graphiqueB.init_titleb ""
''    owner.fdessin.UC_graphiqueB.Top = 0
''    owner.fdessin.UC_graphiqueB.Left = 2500
''    owner.fdessin.UC_graphiqueB.Height = 4210
''    owner.fdessin.UC_graphiqueB.Width = 7800
owner.fdessin.UC_graphique2.reinit 7, "Arial"
owner.fdessin.UC_graphique2.graphique_clear
owner.fdessin.UC_graphique2.init_title
owner.fdessin.UC_graphique2.init_titleh ""
owner.fdessin.UC_graphique2.init_titleb ""
'owner.fdessin.UC_graphique2.Top = 0
'owner.fdessin.UC_graphique2.Left = 350 '240
'owner.fdessin.UC_graphique2.Height = 4400 '4210
'owner.fdessin.UC_graphique2.Width = 10000 '9855
End Sub
Public Sub ini_edessdo()
Dim canal As conduite
Dim tronc As troncon

    edessdo.nombv = ""
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
    edessdo.Lam = 0#
    edessdo.dav = 0
    edessdo.iradav = 0
    edessdo.kav = 0
    edessdo.Lav = 0#
    edessdo.Tram = 0#
    edessdo.Centon = 0#
Call ini_canamo
Call ini_canava
Call ini_canadech
    edo.Longueur = 0#
    edo.hauteur = 0#
    edo.pente = 0#
End Sub
Public Sub ini_debit(ByVal nom As String)
Dim sresult As String
   Call init_graphique
 
    If Trim(ebv.Qchoisi) <> "" Then
        nombassin = nom
        Me.Frm_bv.Caption = "Hydraulique du B.V : " + Trim(nombassin) 'Trim(ebv.nom)
       Select Case ebv.Qchoisi
            Case Is = "CAQUOT"
                Me.Tb_debit(0).Text = rempl_virgule(Format(ebv.Qcor * 1000, "####0.0"))
                Me.Lb_intdebit(0).Caption = "Débit d'eau pluviale (CAQUOT)"
            Case Is = "RATION"
                Me.Tb_debit(0).Text = rempl_virgule(Format(ebv.Qmr * 1000, "####0.0"))
                Me.Lb_intdebit(0).Caption = "Débit d'eau pluviale (Rationnelle)"
            Case Is = "HYDROG"
                Me.Tb_debit(0).Text = rempl_virgule(Format(ebv.Qhydro * 1000, "####0.0"))
                Me.Lb_intdebit(0).Caption = "Débit d'eau pluviale (Hydrogramme)"
        End Select
        Me.Tb_debit(1).Text = rempl_virgule(Format(ebv.Qts, "###0.0"))
        Me.Tb_debit(2).Text = rempl_virgule(Format(ebv.Qrin, "###0.0"))
'        owner.fdessin.UC_graphiqueB.ecr_texta 1665, 1140, "Surface = " + ajout_zero(Trim(Str(ebv.surface))) + " Ha", "G", "B"
'        owner.fdessin.UC_graphiqueB.ecr_texta 1620, 1530, "Coef. de ruissellement = " + ajout_zero(Trim(Str(ebv.imper))), "G", "B"
'        owner.fdessin.UC_graphiqueB.ecr_texta 2055, 1995, "Nombre d'habitants = " + ajout_zero(Trim(Str(ebv.nhab))), "G", "B"
'        owner.fdessin.UC_graphiqueB.ecr_texta 2160, 2505, "Taux de dilution = " + ajout_zero(Trim(Str(ebv.tdilu))), "G", "B"
'        owner.fdessin.UC_graphiqueB.ecr_texta 2875, 540, "Longueur = " + ajout_zero(Trim(Str(ebv.lghydr))) + " m", "G", "B"
'        owner.fdessin.UC_graphiqueB.ecr_texta 3145, 810, "Pente = " + ajout_zero(Trim(Str(ebv.phydr))) + " (1/10000)", "G", "B"
        
     sresult = "Surface = " + ajout_zero(Trim(str(ebv.surface))) + " Ha"
    sresult = sresult + Chr(13) + Chr(10) + "Longueur = " + ajout_zero(Trim(str(ebv.lghydr))) + " m"
    sresult = sresult + Chr(13) + Chr(10) + "Pente = " + ajout_zero(Trim(str(ebv.phydr))) + " (1/10000)"
    sresult = sresult + Chr(13) + Chr(10) + "Coef. de ruissellement = " + ajout_zero(Trim(str(ebv.imper)))
    If ebv.tdilu > 0 Then
    sresult = sresult + Chr(13) + Chr(10) + "Taux de dilution = " + ajout_zero(Trim(str(ebv.tdilu)))
    End If
    If ebv.nhab > 0 Then
    sresult = sresult + Chr(13) + Chr(10) + "Nombre d'habitants = " + ajout_zero(Trim(str(ebv.nhab)))
    End If
    If ebv.ceau > 0 Then
    sresult = sresult + Chr(13) + Chr(10) + "Consommation eau = " + ajout_zero(Trim(str(ebv.ceau))) + " l/hab/j"
    End If
    Me.Lab_bas.Caption = sresult
        
        
        
        Me.SSTab1.TabEnabled(1) = True
    Else
        nombassin = ""
        Me.Tb_debit(0) = "0.0"
        Me.Tb_debit(2) = "0.0"
        Me.Tb_debit(1) = "0.0"
'        owner.fdessin.UC_graphiqueB.ecr_texta 1665, 1140, "Surface", "G", "B"
'        owner.fdessin.UC_graphiqueB.ecr_texta 1620, 1530, "Coef. de ruissellement", "G", "B"
'        owner.fdessin.UC_graphiqueB.ecr_texta 2055, 1995, "Nombre d'habitants", "G", "B"
'        owner.fdessin.UC_graphiqueB.ecr_texta 2160, 2505, "Taux de dilution", "G", "B"
'        owner.fdessin.UC_graphiqueB.ecr_texta 2875, 540, "Longueur", "G", "B"
'        owner.fdessin.UC_graphiqueB.ecr_texta 3145, 810, "Pente", "G", "B"
        Me.Frm_bv.Caption = "Hydraulique du B.V : "
        Me.Lab_bas.Caption = ""
        Me.SSTab1.TabEnabled(1) = False
   End If

   SSTab1.Tab = 0
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
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim reponse As Integer
If ouv_sauve Then 'And sav_fich Then
    reponse = MsgBox("Le déversoir n'a pas été enregistré" + Chr(10) _
        + "Voulez vous le sauvegarder?", 3, "Sauvegarde du déversoir")
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
Private Sub reini_form(ntab As Integer)
'Me.Chk_charge.Value = 0
'Me.Chk_eau.Value = 0
'Me.Chk_piezo.Value = 0
'Me.Chk_Qpluie.Value = 0
'Me.Chk_Qrin.Value = 0
'Me.Chk_Qts.Value = 0

Select Case ntab
    Case Is = 0
        Call init_graphique

        Me.SSTab1.TabEnabled(2) = False
        Me.SSTab1.TabEnabled(3) = False
        Me.SSTab1.TabEnabled(4) = False
        Me.SSTab1.TabEnabled(5) = False
        Me.Frame4.Visible = False
        Me.Frame4.Enabled = False
        If edessdo.Qpluie > 0 And edessdo.Qts > 0 And edessdo.Qrin > 0 Then
            Me.SSTab1.TabEnabled(1) = True
            ouv_sauve = True
        Else
            Me.SSTab1.TabEnabled(1) = False
        End If
        Call ini_valeurs_m
        Call ini_valeurs_v
        Call ini_longce
        Txtb_deversoir.Text = ""
        Txtb_decharge.Text = ""
        Me.Cmd_annulm.Visible = False
        Me.Cmd_annulm.Enabled = False
        Me.Cmd_resum.Enabled = False
        Me.Cmd_annulv.Visible = False
        Me.Cmd_annulv.Enabled = False
        Me.Cmd_resuv.Enabled = False
        Me.Cmd_resudo.Enabled = False
        Me.Cmd_VerifDo.Enabled = False
        Me.Cmd_resudech.Enabled = False
    Case Is = 1
        Call init_graphique

        Call ini_valeurs_m
        Call ini_valeurs_v
        Call ini_longce
        Txtb_deversoir.Text = ""
        Txtb_decharge.Text = ""
        Me.Cmd_annulm.Visible = False
        Me.Cmd_resum.Enabled = False
        Me.Cmd_annulv.Visible = False
        Me.Cmd_resuv.Enabled = False
        Me.Cmd_resudo.Enabled = False
        Me.Cmd_VerifDo.Enabled = False
        Me.Cmd_resudech.Enabled = False
        Me.SSTab1.TabEnabled(2) = False
        Me.SSTab1.TabEnabled(3) = False
        Me.SSTab1.TabEnabled(4) = False
        Me.SSTab1.TabEnabled(5) = False
        Me.Frame4.Visible = False
        Me.Frame4.Enabled = False
        If Me.SSTab1.TabEnabled(1) Then
       If edessdo.lgca > 0 And edessdo.lgdisp > 0 And edessdo.phex > 0 _
            And edessdo.rdoam > 0 And edessdo.rdoav > 0 And edessdo.rdoex > 0 Then
           Me.SSTab1.TabEnabled(2) = True
           ok_tab1 = True
           If ok_tab2 Then
                Me.Cmd_resum.Enabled = True
           End If
        End If
        End If
    Case Is = 2
        Call ini_valeurs_m
        Call ini_valeurs_v
        Call ini_longce
        Txtb_deversoir.Text = ""
        Txtb_decharge.Text = ""
        Me.Cmd_annulv.Visible = False
        Me.Cmd_resuv.Enabled = False
        Me.Cmd_resudo.Enabled = False
        Me.Cmd_VerifDo.Enabled = False
        Me.Cmd_resudech.Enabled = False
        Me.SSTab1.TabEnabled(4) = False
        Me.SSTab1.TabEnabled(5) = False
        Me.SSTab1.TabEnabled(3) = False
        owner.fdessin.UC_graphique1.graphique_clear
        owner.fdessin.UC_graphique2.graphique_clear
        Me.Frame4.Visible = False
        Me.Frame4.Enabled = False
        If edessdo.dam > 0 And edessdo.iRadam > 0 And edessdo.Kam > 0 _
            And edessdo.Lam > 0 Then
 '          Me.SSTab1.TabEnabled(3) = True
            If Cmd_resum.Enabled = False Then
 '               Call ini_valeurs_m
                Me.Cmd_resum.Enabled = True
                If Me.Cmd_annulm.Enabled Then
                    Me.Cmd_annulm.Visible = False
                End If
            End If
'            If ok_tab3 Then
'                 Me.Cmd_resuv.Enabled = True
'            End If
            Me.Cmd_amo.Enabled = True
        Else
            Me.Cmd_amo.Enabled = False
            Me.Cmd_resum.Enabled = False
            If Me.Cmd_annulm.Enabled Then
                Me.Cmd_annulm.Visible = True
            End If
       End If
    Case Is = 21
        Call ini_valeurs_v
        Call ini_longce
        Txtb_deversoir.Text = ""
        Txtb_decharge.Text = ""
        Me.Cmd_annulm.Visible = False
        Me.Cmd_resum.Enabled = False
        Me.Cmd_annulv.Visible = False
        Me.Cmd_resuv.Enabled = False
        Me.Cmd_resudo.Enabled = False
        Me.Cmd_VerifDo.Enabled = False
        Me.Cmd_resudech.Enabled = False
        Me.Frame4.Enabled = False
        Me.Frame4.Visible = False
        Me.SSTab1.TabEnabled(3) = True
        Me.SSTab1.TabEnabled(4) = False
        Me.SSTab1.TabEnabled(5) = False
        If ok_tab3 Then
             Me.Cmd_resuv.Enabled = True
        End If
    Case Is = 3
        Call ini_valeurs_v
        Call ini_longce
        Txtb_deversoir.Text = ""
        Txtb_decharge.Text = ""
        Me.Cmd_resudo.Enabled = True
        Me.Cmd_VerifDo.Enabled = False
        Me.Cmd_resudech.Enabled = False
        Me.SSTab1.TabEnabled(4) = False
        Me.SSTab1.TabEnabled(5) = False
        Me.Frame4.Visible = False
        Me.Frame4.Enabled = False
        owner.fdessin.UC_graphique1.graphique_clear
        owner.fdessin.UC_graphique2.graphique_clear
        If edessdo.dav > 0 And edessdo.iradav > 0 And edessdo.kav > 0 _
            And edessdo.Lav > 0 Then
'           Me.SSTab1.TabEnabled(4) = True
            If Cmd_resuv.Enabled = False Then
'                Call ini_valeurs_v
                Me.Cmd_resuv.Enabled = True
                If Me.Cmd_annulv.Enabled Then
                    Me.Cmd_annulv.Visible = True
                End If
            End If
'            If ok_tab4 Then
'                 Me.Cmd_resudo.Enabled = True
'            End If
             Me.Cmd_ava.Enabled = True
        Else
            Me.Cmd_ava.Enabled = False
            Me.Cmd_resuv.Enabled = False
            If Me.Cmd_annulv.Enabled Then
                Me.Cmd_annulv.Visible = True
            End If
       End If
    Case Is = 31
        Call ini_longce
        Txtb_deversoir.Text = ""
        Txtb_decharge.Text = ""
        Me.Cmd_annulv.Visible = False
        Me.Cmd_resuv.Enabled = False
        Me.Cmd_resudo.Enabled = True
        Me.Cmd_VerifDo.Enabled = False
        Me.Cmd_resudech.Enabled = False
        Me.SSTab1.TabEnabled(4) = True
        Me.SSTab1.TabEnabled(5) = False
        Me.Frame4.Enabled = False
        Me.Frame4.Visible = False
        If ok_tab4 Then
            Me.Cmd_resudo.Enabled = True
        End If
   Case Is = 40
        Txtb_deversoir.Text = ""
        Txtb_decharge.Text = ""
        Call ini_longce
        Me.Frame4.Enabled = False
        Me.Frame4.Visible = False
        owner.fdessin.UC_graphique1.graphique_clear
        owner.fdessin.UC_graphique2.graphique_clear
        Me.Cmd_resudo.Enabled = False
        Me.Cmd_VerifDo.Enabled = False
        Me.SSTab1.TabEnabled(5) = False
        Me.Cmd_resudech.Enabled = False
    Case Is = 41
        Me.Txtb_decharge.Text = ""
        Me.Txtb_deversoir.Text = ""
        Me.Frame4.Visible = False
        Me.Frame4.Enabled = False
        owner.fdessin.UC_graphique1.graphique_clear
        owner.fdessin.UC_graphique2.graphique_clear
        Me.Chk_Qpluie.Enabled = False
        Me.Chk_Qpluie.Value = 0
        Me.SSTab1.TabEnabled(5) = False
        Me.Cmd_resudech.Enabled = False
        If Me.Cmd_resudo.Enabled = False Then
            Me.Cmd_VerifDo.Enabled = True
        End If
   Case Is = 42
        Txtb_deversoir.Text = ""
        Txtb_decharge.Text = ""
        Call ini_longce
        Me.Frame4.Enabled = False
        Me.Frame4.Visible = False
        owner.fdessin.UC_graphique1.graphique_clear
        owner.fdessin.UC_graphique2.graphique_clear
        Me.Cmd_resudo.Enabled = True
        Me.Cmd_VerifDo.Enabled = False
        Me.SSTab1.TabEnabled(5) = False
        Me.Cmd_resudech.Enabled = False
   Case Is = 50
        Me.Txtb_decharge.Text = ""
        owner.fdessin.UC_graphique2.graphique_clear
        If txtVersNum(Tb_dech(0).Text) > 0 And txtVersNum(Tb_dech(1).Text) > 0 And txtVersNum(Tb_dech(2).Text) > 0 _
            And txtVersNum(Tb_dech(3).Text) > 0 Then
             Me.Cmd_dech.Enabled = True
        Else
            Me.Cmd_dech.Enabled = False
       End If
        Me.Cmd_resudech.Enabled = False
   Case Is = 51
        Me.Txtb_decharge.Text = ""
        owner.fdessin.UC_graphique2.graphique_clear
        If txtVersNum(Tb_dech(0).Text) > 0 And txtVersNum(Tb_dech(1).Text) > 0 And txtVersNum(Tb_dech(2).Text) > 0 _
            And txtVersNum(Tb_dech(3).Text) > 0 Then
             Me.Cmd_dech.Enabled = True
        Else
            Me.Cmd_dech.Enabled = False
       End If
        Me.Cmd_resudech.Enabled = True

End Select
' impression false
    Me.mnuprint.Enabled = False
    ouv_sauve = True
End Sub
Private Sub ini_form()
    Me.Tb_debit(1).Text = rempl_virgule(Format(edessdo.Qts, "###0.0"))
    Me.Tb_debit(2).Text = rempl_virgule(Format(edessdo.Qrin, "###0.0"))
    Me.Tb_debit(0).Text = rempl_virgule(Format(edessdo.Qpluie, "###0.0"))
    Me.Tb_cont(0).Text = rempl_virgule(Format(edessdo.rdoam, "###0.0"))
    Me.Tb_cont(1).Text = rempl_virgule(Format(edessdo.rdoav, "###0.0"))
    Me.Tb_cont(2).Text = rempl_virgule(Format(edessdo.lgdisp, "###0.0"))
    Me.Tb_cont(3).Text = rempl_virgule(Format(edessdo.phex, "###0.0"))
    Me.Tb_cont(4).Text = rempl_virgule(Format(edessdo.rdoex, "###0.0"))
    Me.Tb_cont(5).Text = rempl_virgule(Format(edessdo.lgca, "###0.0"))
    Me.Tb_amo(0).Text = rempl_virgule(Format(edessdo.dam, "###0"))
    Me.Tb_amo(1).Text = rempl_virgule(Format(edessdo.iRadam, "###0"))
    Me.Tb_amo(3).Text = rempl_virgule(Format(edessdo.Lam, "###0.00"))
    Me.Tb_amo(2).Text = rempl_virgule(Format(edessdo.Kam, "###0"))
    Me.Tb_ava(0).Text = rempl_virgule(Format(edessdo.dav, "###0"))
    Me.Tb_ava(1).Text = rempl_virgule(Format(edessdo.iradav, "###0"))
    Me.Tb_ava(3).Text = rempl_virgule(Format(edessdo.Lav, "###0.00"))
    Me.Tb_ava(2).Text = rempl_virgule(Format(edessdo.kav, "###0"))
    Me.Tb_dev(0).Text = rempl_virgule(Format(edo.Longueur, "##0.00"))
    Me.Tb_dev(1).Text = rempl_virgule(Format(edo.hauteur, "###0.000"))
    Me.Tb_dev(2).Text = rempl_virgule(Format(edo.pente, "##0.0000"))
    Me.Tb_dev(3).Text = rempl_virgule(Format(edessdo.Tram, "##0.00"))
    Me.Tb_dech(0).Text = rempl_virgule(Format(edessdo.tron_dech.conduit.Diametre * 1000, "###0"))
    Me.Tb_dech(1).Text = rempl_virgule(Format(edessdo.tron_dech.conduit.pente * 10000, "###0"))
    Me.Tb_dech(3).Text = rempl_virgule(Format(edessdo.tron_dech.conduit.Longueur, "###0.00"))
    Me.Tb_dech(2).Text = rempl_virgule(Format(edessdo.tron_dech.conduit.rugosite, "###0"))
    Me.Tb_dech(4).Text = rempl_virgule(Format(edessdo.Centon, "##0.00"))
    Cb_centon.Text = Tb_dech(4).Text
    Cb_centon.Refresh
    Me.Frm_bv.Caption = "Hydraulique du B.V : " + Trim(nombassin) 'edessdo.nombv) 'Trim(ebv.nom)

End Sub

Private Sub ini_resuintdv()
With Resuintdev
'    .dam = "Diamètre"
'    .iRadam = "Pente"
'    .Kam = "Coefficient de Manning-Strickler"
'    .Lam = "Longueur"
'    .dav = "Diamètre"
'    .iradav = "Pente"
'    .kav = "Coefficient de Manning-Strickler"
'    .Lav = "Longueur"
'    .ddech = "Diamètre"
'    .iraddech = "Pente"
'    .kdech = "Coefficient de Manning-Strickler"
'    .Ldech = "Longueur"
    .Ldev = " Longueur du DO"
    .Hcret = " Hauteur de la crête"
    .Pdev = " Pente du DO"
    .Tram = " Tirant d'eau amont admissible"
    .dpsm = "Débit pleine section"
    .vpsm = "Vitesse pleine section"
    .vqtsm = "Vitesse d'écoulement à QTS"
    .hqtsm = "Hauteur d'eau QTS"
    .vqrinm = "Vitesse d'écoulement à QRIN"
    .hqrinm = "Hauteur d'eau QRIN"
    .vqpluiem = "Vitesse d'écoulement amont à QPLUIE"
    .hqpluiem = "Hauteur d'eau amont QPLUIE"
    .vqpluiemav = "Vitesse d'écoulement aval à QPLUIE"
    .hqpluiemav = "Hauteur d'eau aval QPLUIE"
    .dpsv = "Débit pleine section"
    .vpsv = "Vitesse pleine section"
    .vqtsv = "Vitesse d'écoulement à QTS"
    .hqtsv = "Hauteur d'eau QTS"
    .vqrinv = "Vitesse d'écoulement à QRIN"
    .hqrinv = "Hauteur d'eau QRIN"
    .vqpluiev = "Vitesse d'écoulement à QPLUIE"
    .hqpluiev = "Hauteur d'eau QPLUIE"
    .longetranglee = " Longueur conduite étranglée"
    .debetranglee = "Débit dans la conduite étranglée"
    .debdeverse = "Débit déversé"
    .dpsdech = "Débit pleine section décharge"
    .vpsdech = "Vitesse pleine section décharge"
    .vqdev = "Vitesse d'écoulement amont pour débit déversé"
    .hqdev = "Hauteur d'eau amont pour débit déversé"
    .vqdevav = "Vitesse d'écoulement aval pour débit déversé"
    .hqdevav = "Hauteur d'eau aval pour débit déversé"
    .regime = "Régime"
    .Ham = "Hauteur de la lame d'eau"
    .Hav = "Hauteur de la lame d'eau"
    .Haam = "Hauteur de la charge"
    .Haav = "Hauteur de la charge"
End With
With Resuudev
'    .dam = "mm"
'    .iRadam = "1/10000"
'    .Kam = ""
'    .Lam = "m"
'    .dav = "mm"
'    .iradav = "1/10000"
'    .kav = ""
'    .Lav = "m"
'    .ddech = "mm"
'    .iraddech = "1/10000"
'    .kdech = ""
'    .Ldech = "m"
    .Ldev = " m"
    .Hcret = " m"
    .Pdev = " m/m"
    .Tram = " m"
    .dpsm = "m3/s"
    .vpsm = "m/s"
    .vqtsm = "m/s"
    .hqtsm = " m"
    .vqrinm = "m/s"
    .hqrinm = " m"
    .vqpluiem = "m/s"
    .hqpluiem = " m"
    .vqpluiemav = "m/s"
    .hqpluiemav = " m"
    .dpsv = "m3/s"
    .vpsv = "m/s"
    .vqtsv = "m/s"
    .hqtsv = " m"
    .vqrinv = "m/s"
    .hqrinv = " m"
    .vqpluiev = "m/s"
    .hqpluiev = " m"
    .longetranglee = " m"
    .debetranglee = "m3/s"
    .debdeverse = "m3/s"
    .dpsdech = "m3/s"
    .vpsdech = "m/s"
    .vqdev = "m/s"
    .hqdev = "m"
    .vqdevav = "m/s"
    .hqdevav = "m"
    .regime = ""
    .Ham = "m"
    .Hav = "m"
    .Haam = "m"
    .Haav = "m"
End With
With resudev
'    .dam = "0"
'    .iRadam = "0"
'    .Kam = "0"
'    .Lam = "0.0"
'    .dav = "0"
'    .iradav = "0"
'    .kav = "0"
'    .Lav = "0.0"
'    .ddech = "0"
'    .iraddech = "0"
'    .kdech = "0"
'    .Ldech = "0.0"
    .Ldev = "0.0"
    .Hcret = "0.0"
    .Pdev = "0.0"
    .Tram = "0.0"
    .dpsm = "0.0"
    .vpsm = "0.0"
    .vqtsm = "0.0"
    .hqtsm = "0.0"
    .vqrinm = "0.0"
    .hqrinm = "0.0"
    .vqpluiem = "0.0"
    .hqpluiem = "0.0"
    .dpsv = "0.0"
    .vpsv = "0.0"
    .vqtsv = "0.0"
    .hqtsv = "0.0"
    .vqrinv = "0.0"
    .hqrinv = "0.0"
    .vqpluiev = "0.0"
    .hqpluiev = "0.0"
    .longetranglee = "0.0"
    .debetranglee = "0.0"
    .debdeverse = "0.0"
    .dpsdech = "0.0"
    .vpsdech = "0.0"
    .vqdev = "0.0"
    .hqdev = "0.0"
    .regime = ""
    .Ham = "0.0"
    .Hav = "0.0"
    .Haam = "0.0"
    .Haav = "0.0"
End With
End Sub


Private Function verif_resu(ByRef resu As Resudo) As Boolean
    verif_resu = True
    
    With resu
        If .ldav <= 0 Or .ddav <= 0 Or .pdav <= 0 Or .dlongdo <= 0 Or .dpentedo <= 0 Then
            verif_resu = False
        End If
    End With
End Function
Public Sub ini_resudo()
    With Resup_do
        .ldav = 0#
        .pdav = 0#
        .ddav = 0#
        .dlongdo = 0#
        .dpentedo = 0#
    End With
End Sub
Public Sub ini_bv()
    ebv.nom = ""
    ebv.Qchoisi = ""
    ebv.type = "U"
    ebv.surface = 0
    ebv.imper = 0
    ebv.lghydr = 0
    ebv.phydr = 0
    ebv.nhab = 0
    ebv.tdilu = 0
    ebv.ceau = 0
    ebv.perti = 0
    ebv.vinf = 0
    ebv.ahorton = 0#
    ebv.bhorton = 0
    ebv.trep = 0
    ebv.Qbrut = 0#
    ebv.Qcor = 0#
    ebv.Qmr = 0#
    ebv.Qhydro = 0#
    ebv.Qeu = 0#
    ebv.Qecp = 0#
    ebv.Qts = 0#
    ebv.Qprin = 0#
    ebv.Qrin = 0#
    ebv.tc = 0#
    ebv.qfuite = 0
    eph.amontana = 0#
    eph.bmontana = 0#
    eph.lcrin = 0
    eph.ceau = 0
    eph.aeu = 0
    eph.beu = 0
    eph.a1montana = 0#
    eph.b1montana = 0#
    eph.Seuil = 0#
    ehyd.DM = 0
    ehyd.dt = 0
    ehyd.HM = 0#
    ehyd.HT = 0#
    ehyd.pas = 5
    ehyd.Teta = 0.5
    ehyd.kdesbor = 0#
    ehyd.qfuite = 0
    ehyd.vst = 0#
    ehyd.vstock = 0#
End Sub
Private Sub ini_form_exist()
Dim ok As Boolean
Dim s As Double
Call ini_form
reini_form 0
reini_form 1
If Me.SSTab1.TabEnabled(2) Then
If edessdo.dam > 0 And edessdo.iRadam > 0 And edessdo.Kam > 0 _
    And edessdo.Lam > 0 Then

ok = calcul_amont
Me.SSTab1.Tab = 2
If ok Then
    If edessdo.dav > 0 And edessdo.iradav > 0 And edessdo.kav > 0 _
        And edessdo.Lav > 0 Then
    ok = calcul_aval
    Me.SSTab1.Tab = 3
    If ok Then
      Call Tb_dev_Change(0)
        If Cmd_resudo.Enabled Then
           Cmd_resudo_Click
            Me.SSTab1.Tab = 4
           If Me.Cmd_VerifDo.Enabled Then
                Cmd_VerifDo_Click
                If Me.Cmd_resudech.Enabled Then
                    owner.fdessin.UC_graphique1.Visible = False
                    owner.fdessin.UC_graphique2.Visible = True
                    Cmd_resudech_Click
                    Me.SSTab1.Tab = 5
                End If
            End If
        End If
    End If
    End If
End If
End If
End If
'Me.SSTab1.Tab = 0
'Me.SSTab1.TabEnabled(1) = True
'Me.SSTab1.TabEnabled(2) = True
'Me.SSTab1.TabEnabled(3) = True
'Me.SSTab1.TabEnabled(4) = True
'Me.SSTab1.TabEnabled(5) = True
'If calcul_amont Then
'    calcul_aval
'    Cmd_resudo_Click
'    Cmd_VerifDo_Click
'    calcul_dech
'    ok_imp = True
'End If
End Sub
Private Sub ini_canamo()
edessdo.tron_amo.conduit.Diametre = 0
edessdo.tron_amo.conduit.Longueur = 0#
edessdo.tron_amo.conduit.pente = 0
edessdo.tron_amo.conduit.rugosite = 0
End Sub
Private Sub ini_canava()
edessdo.tron_ava.conduit.Diametre = 0
edessdo.tron_ava.conduit.Longueur = 0#
edessdo.tron_ava.conduit.pente = 0
edessdo.tron_ava.conduit.rugosite = 0
End Sub
Private Sub ini_canadech()
edessdo.tron_dech.conduit.Diametre = 0
edessdo.tron_dech.conduit.Longueur = 0#
edessdo.tron_dech.conduit.pente = 0
edessdo.tron_dech.conduit.rugosite = 0
End Sub
Private Sub reini_valeurs(ntab As Integer)
Select Case ntab
    Case Is = 0
        Call ini_valeurs_m
        Call ini_valeurs_v
        Call ini_longce
        Txtb_deversoir.Text = ""
        Txtb_decharge.Text = ""
        Me.Cmd_annulm.Visible = False
        Me.Cmd_resum.Enabled = False
        Me.Cmd_annulv.Visible = False
        Me.Cmd_resuv.Enabled = False
        Me.Cmd_resudo.Enabled = False
        Me.Cmd_VerifDo.Enabled = False
        Me.Cmd_resudech.Enabled = False
    Case Is = 1
        Call ini_valeurs_v
        Call ini_longce
        Txtb_deversoir.Text = ""
        Txtb_decharge.Text = ""
        Me.Cmd_annulv.Visible = False
        Me.Cmd_resuv.Enabled = False
        Me.Cmd_resudo.Enabled = False
        Me.Cmd_VerifDo.Enabled = False
        Me.Cmd_resudech.Enabled = False
End Select
End Sub


Private Sub mnusaves_Click()
'modif FO   ' If ProtectCheck(2) <> 0 Then End
    If fich_lect = nom_fich Or Trim(Tb_titre.Text) = "" Or fich_lect = "" Then
        Frm_titre.Label2.Caption = "Sauvegarde d'un déversoir "
        Frm_titre.Label3.Caption = ""
    Else
         Frm_titre.Label2.Caption = "Sauvegarde du déversoir " & Me.Tb_titre.Text
         Frm_titre.Label3.Caption = " de l'étude " & fich_lect_edit
   
    End If
    Frm_titre.Caption = "Etude " + nom_fich_edit
    Frm_titre.Label1.Caption = "Nom du déversoir (30car. maxi)"
    Frm_titre.Text1.Text = Me.Tb_titre.Text
    Frm_titre.Show 1
End Sub

Private Sub mnusuppr_Click()
Dim za As st_savdo
Dim za1 As st_savdo1
Dim lhFicDbf1 As Integer, reponse As Integer
Dim nom As String

'modif FO   ' If ProtectCheck(2) <> 0 Then End

 If Trim(Cb_deversoir.Text) <> "" Then
    Call funlockb
    reponse = MsgBox(Trim(Cb_deversoir.Text) + " va être supprimé .", 4, "Suppression d'un déversoir")
    If reponse = 6 Then  '6=oui,7=non
    save_fich = True
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
            za = za1.stsavdo
            If Trim(za.type) <> nom_type Or (Trim(za.type) = nom_type And Trim(za.nom) <> Trim(Cb_deversoir.Text)) Then
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
    Me.SSTab1.Tab = 0
    Me.SSTab1.TabEnabled(1) = False
    Me.SSTab1.TabEnabled(2) = False
    Me.SSTab1.TabEnabled(3) = False
    Me.SSTab1.TabEnabled(4) = False
    Me.SSTab1.TabEnabled(5) = False
    Me.Tb_titre.Text = ""
    Me.Caption = fen_titre
'    Me.Cmd_del.Visible = False
    nombassin = ""
    Call ini_edessdo
    Call ini_resuintdv
    Call ini_form
'    Call ini_debit
        nombassin = ""
        Me.Tb_debit(0) = "0.0"
        Me.Tb_debit(2) = "0.0"
        Me.Tb_debit(1) = "0.0"
        Call init_graphique
        Me.Lab_bas.Caption = ""

'        owner.fdessin.UC_graphiqueB.ecr_texta 1665, 1140, "Surface", "G", "B"
'        owner.fdessin.UC_graphiqueB.ecr_texta 1620, 1530, "Coef. de ruissellement", "G", "B"
'        owner.fdessin.UC_graphiqueB.ecr_texta 2055, 1995, "Nombre d'habitants", "G", "B"
'        owner.fdessin.UC_graphiqueB.ecr_texta 2160, 2505, "Taux de dilution", "G", "B"
'        owner.fdessin.UC_graphiqueB.ecr_texta 2875, 540, "Longueur", "G", "B"
'        owner.fdessin.UC_graphiqueB.ecr_texta 3145, 810, "Pente", "G", "B"
        Me.SSTab1.TabEnabled(1) = False
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

Private Sub OK_lignes_Click()
If Me.Frame4.Enabled = True Then

'dessin dans  owner.fdessin.UC_graphique1
Call init_graphdo(owner.fdessin.UC_graphique1)
Call dess_troncon(owner.fdessin.UC_graphique1, edessdo.tron_amo, couleur.noir) ' vbBlack)
Call dess_predo(owner.fdessin.UC_graphique1, edo, couleur.noir)
Call dess_cot(owner.fdessin.UC_graphique1, couleur.noir) ' vbBlack)
Call dessin_do_hydrau(owner.fdessin.UC_graphique1, (Chk_charge = 1), (Chk_piezo = 1), (Chk_eau = 1), (Chk_Qts = 1), (Chk_Qrin = 1), (Chk_Qpluie = 1))
'If Chk_piezo.Value = 1 Then
'
'    Call dess_piezo(owner.fdessin.UC_graphique1, edessdo.tron_amo, edessdo.Qts, couleur.magenta) ' vbGreen) ' RGB(128, 64, 128))
'    Call dess_piezo(owner.fdessin.UC_graphique1, edessdo.tron_amo, edessdo.Qrin, couleur.magenta_clair) ' vbRed)
'    Call dess_piezo(owner.fdessin.UC_graphique1, edessdo.tron_amo, edessdo.Qpluie, couleur.vert) ' vbMagenta)
'    Call dess_piezo(owner.fdessin.UC_graphique1, edo.tron_ava, edessdo.Qts, couleur.magenta) ' vbGreen)
'    Call dess_piezo(owner.fdessin.UC_graphique1, edo.tron_ava, edessdo.Qrin, couleur.magenta_clair) ' vbRed)
' '   Call dess_piezo(owner.fdessin.UC_graphique1, troava, edessdo.Qpluie, vbMagenta)
'End If
'If Chk_charge.Value = 1 Then
'
'
'    Call dess_charge(owner.fdessin.UC_graphique1, edo.tron_ava, edessdo.Qrin, couleur.orange) ' vbCyan)
'    Call dess_charge(owner.fdessin.UC_graphique1, edessdo.tron_amo, edessdo.Qrin, couleur.orange) ' vbCyan)
'End If
'If Chk_charge.Value = 1 Then
'End If
'
End If
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
Dim mes As String
Select Case SSTab1.Tab
    Case Is = 0
        mes = IDhlp_DODonneesBase
        owner.fdessin.Image1.Visible = False
        owner.fdessin.UC_graphiqueB.Visible = False
        owner.fdessin.Image3.Visible = True
        owner.fdessin.UC_graphique1.Visible = False
        owner.fdessin.UC_graphique2.Visible = False
   Case Is = 1
        mes = IDhlp_DOContraintes
        owner.fdessin.Image1.Visible = False
        owner.fdessin.UC_graphiqueB.Visible = False
        owner.fdessin.Image3.Visible = True
        owner.fdessin.UC_graphique1.Visible = False
        owner.fdessin.UC_graphique2.Visible = False
   Case Is = 2
        mes = IDhlp_DOConduiteAmenee
        owner.fdessin.Image1.Visible = False
        owner.fdessin.UC_graphique1.Visible = True
        owner.fdessin.UC_graphiqueB.Visible = False
        owner.fdessin.Image3.Visible = False
        owner.fdessin.UC_graphique2.Visible = False
   Case Is = 3
        mes = IDhlp_DOConduiteDebitConserve
        owner.fdessin.Image1.Visible = False
        owner.fdessin.UC_graphiqueB.Visible = False
        owner.fdessin.UC_graphique1.Visible = True
        owner.fdessin.UC_graphique2.Visible = False
   Case Is = 4
        mes = IDhlp_DOChambreDeversement
        owner.fdessin.Image1.Visible = False
        owner.fdessin.UC_graphiqueB.Visible = False
        owner.fdessin.Image3.Visible = False
        owner.fdessin.UC_graphique1.Visible = True
        owner.fdessin.UC_graphique2.Visible = False
   Case Is = 5
        mes = IDhlp_DOConduiteDecharge
        owner.fdessin.Image1.Visible = False
        owner.fdessin.UC_graphiqueB.Visible = False
        owner.fdessin.Image3.Visible = False
        owner.fdessin.UC_graphique1.Visible = False
        owner.fdessin.UC_graphique2.Visible = True
End Select

If owner.fcom.Name = "Frm_ss_commentaire" Then
    Change_Couleur "SSTab1", 0
    DoEvents
    owner.affich_aide Me.Name, mes
    DoEvents
End If
End Sub

Private Sub Tb_amo_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_amo"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
Call sel_text(Tb_amo(Index))
End Sub

Private Sub Tb_amo_GotFocus(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_amo"
If SSTab1.Tab <> 2 Then
    If Me.SSTab1.TabEnabled(2) Then
        SSTab1.Tab = 2
    Else
        Me.Tb_debit(0).SetFocus
    End If
'    Me.Tb_Debit(0).SetFocus
End If
Call sel_text(Tb_amo(Index))
If change_coul Then
    Change_Couleur nom, Index
    mes = Rec_Mes(nom, Index)
    owner.affich_aide Me.Name, mes
End If

End Sub

Private Sub Tb_amo_LostFocus(Index As Integer)
Dim ok As Boolean
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_amo", Index, txtVersNum(Tb_amo(Index).Text))
    If Not ok Then
        Tb_amo(Index).SetFocus
        DoEvents
    End If
    okg = True
End If

End Sub

Private Sub Tb_ava_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_ava"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
Call sel_text(Tb_ava(Index))
End Sub

Private Sub Tb_ava_GotFocus(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_ava"
If SSTab1.Tab <> 3 Then
    If Me.SSTab1.TabEnabled(3) Then
        SSTab1.Tab = 3
    Else
        Me.Tb_debit(0).SetFocus
    End If
'    Me.Tb_Debit(0).SetFocus
End If
Call sel_text(Tb_ava(Index))
If change_coul Then
    Change_Couleur nom, Index
    mes = Rec_Mes(nom, Index)
    owner.affich_aide Me.Name, mes
End If

End Sub

Private Sub Tb_ava_LostFocus(Index As Integer)
Dim ok As Boolean
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_ava", Index, txtVersNum(Tb_ava(Index).Text))
    If Not ok Then
        Tb_ava(Index).SetFocus
        DoEvents
    End If
    okg = True
End If

End Sub

Private Sub Tb_cont_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_cont"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
Call sel_text(Tb_cont(Index))
End Sub

Private Sub Tb_cont_GotFocus(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_cont"
If SSTab1.Tab <> 1 Then
    If Me.SSTab1.TabEnabled(1) Then
        SSTab1.Tab = 1
    Else
        Me.Tb_debit(0).SetFocus
    End If
'    Me.Tb_Debit(0).SetFocus
End If
Call sel_text(Tb_cont(Index))
If change_coul Then
    Change_Couleur nom, Index
    mes = Rec_Mes(nom, Index)
    owner.affich_aide Me.Name, mes
End If

End Sub

Private Sub Tb_cont_LostFocus(Index As Integer)
Dim ok As Boolean
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_cont", Index, txtVersNum(Tb_cont(Index).Text))
    If Not ok Then
        Tb_cont(Index).SetFocus
        DoEvents
    End If
    okg = True
End If

End Sub

Private Sub Tb_debit_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_debit"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
Call sel_text(Tb_debit(Index))
End Sub

Private Sub Tb_debit_GotFocus(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_debit"
If SSTab1.Tab <> 0 Then
'    SSTab1.Tab = 0
End If
Call sel_text(Tb_debit(Index))
If change_coul Then
    Change_Couleur nom, Index
    mes = Rec_Mes(nom, Index)
    owner.affich_aide Me.Name, mes
End If

End Sub

Private Sub Tb_debit_LostFocus(Index As Integer)
Dim ok As Boolean
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_debit", Index, txtVersNum(Tb_debit(Index).Text))
    If Not ok Then
        Tb_debit(Index).SetFocus
        DoEvents
    End If
    okg = True
End If

End Sub

Private Sub tb_dech_Change(Index As Integer)
Dim s As Double
Dim nom As String

If bKP Then
        Select Case Index
            Case Is = 0
                nom = verif_cart0(Tb_dech(Index).Text, "Saisie diamétre canalisation de décharge", "I")
            Case Is = 1
                nom = verif_cart0(Tb_dech(Index).Text, "Saisie pente canalisation de décharge", "I")
            Case Is = 2
                nom = verif_cart0(Tb_dech(Index).Text, "Saisie coefficient canalisation de décharge", "I")
            Case Is = 3
                nom = verif_cart0(Tb_dech(Index).Text, "Saisie longueur canalisation de décharge", "R")
            Case Is = 4
                nom = "ok"
        End Select
  If nom = "" Then
    Tb_dech(Index).Text = sval_champ
    Tb_dech(Index).SelStart = iSels
    Tb_dech(Index).SelLength = iSell
  Else
'  End If
'End If
'****

s = 1
For i = 0 To 4
    s = s * txtVersNum(Me.Tb_dech(i).Text)
Next

    Select Case Index
        Case Is = 0
            edessdo.tron_dech.conduit.Diametre = txtVersNum(Tb_dech(0).Text) / 1000#
        Case Is = 1
            edessdo.tron_dech.conduit.pente = txtVersNum(Tb_dech(1).Text) / 10000
        Case Is = 2
            edessdo.tron_dech.conduit.rugosite = txtVersNum(Tb_dech(2).Text)
        Case Is = 3
            edessdo.tron_dech.conduit.Longueur = txtVersNum(Tb_dech(3).Text)
        Case Is = 4
            centon_texte = Tb_dech(4).Text
            edessdo.Centon = txtVersNum(Tb_dech(4).Text)
            centon_texte = Tb_dech(4).Text
'            Cb_centon.Text = Tb_dech(4).Text
'            Cb_centon.Refresh
    End Select
If s > 0 Then
    
    Call reini_form(51)
'    Me.Cmd_resudech.Enabled = True
'    Me.Txtb_decharge.Text = ""
Else
    Call reini_form(50)
'    Me.Cmd_resudech.Enabled = False
'    Me.Txtb_decharge.Text = ""
End If
  End If
End If

 sval_champ = ""
    bKP = False

End Sub

Private Sub Tb_dech_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_dech"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
DoEvents
owner.affich_aide Me.Name, mes
Call sel_text(Tb_dech(Index))
End Sub

Private Sub Tb_dech_GotFocus(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_dech"
If SSTab1.Tab <> 5 Then
    If Me.SSTab1.TabEnabled(5) Then
        SSTab1.Tab = 5
    Else
        Me.Tb_debit(0).SetFocus
    End If
'    Me.Tb_Debit(0).SetFocus
End If
Call sel_text(Tb_dech(Index))
If change_coul Then
    Change_Couleur nom, Index
    mes = Rec_Mes(nom, Index)
    owner.affich_aide Me.Name, mes
End If

End Sub

Private Sub Tb_dech_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    sval_champ = Tb_dech(Index).Text
    iSels = Tb_dech(Index).SelStart
    iSell = Tb_dech(Index).SelLength
    bKP = True
'   If Len(Tb_dech(Index).Text) <= Tb_dech(Index).MaxLength Then
'       Select Case Index
'            Case Is = 0
'                KeyAscii = verif_car(Tb_dech(Index).Text, KeyAscii, "Saisie diamétre canalisation de décharge", "I")
'            Case Is = 1
'                KeyAscii = verif_car(Tb_dech(Index).Text, KeyAscii, "Saisie pente canalisation de décharge", "I")
'            Case Is = 2
'                KeyAscii = verif_car(Tb_dech(Index).Text, KeyAscii, "Saisie coefficient canalisation de décharge", "I")
'            Case Is = 3
'                KeyAscii = verif_car(Tb_dech(Index).Text, KeyAscii, "Saisie longueur canalisation de décharge", "R")
'        End Select
'    End If
End If
End Sub

Private Sub TB_dech_LostFocus(Index As Integer)
Dim ok As Boolean
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_dech", Index, txtVersNum(Tb_dech(Index).Text))
    If Not ok Then
        Tb_dech(Index).SetFocus
        DoEvents
    End If
    okg = True
End If

End Sub

Private Sub Tb_dev_Change(Index As Integer)
Dim s As Double
Dim nom As String

If bKP Then
        Select Case Index
            Case Is = 0
                nom = verif_cart0(Tb_dev(Index).Text, "Saisie longueur DO", "R")
            Case Is = 1
                nom = verif_cart0(Tb_dev(Index).Text, "Saisie hauteur de la crête", "R")
            Case Is = 2
                nom = verif_cart0(Tb_dev(Index).Text, "Saisie pente DO", "R")
            Case Is = 3
                nom = verif_cart0(Tb_dev(Index).Text, "Saisie tirant d'eau amont admissible", "R")
        End Select
  If nom = "" Then
    Tb_dev(Index).Text = sval_champ
    Tb_dev(Index).SelStart = iSels
    Tb_dev(Index).SelLength = iSell
  Else
'  End If
'End If
'****

s = 1
For i = 0 To 2
    s = s * txtVersNum(Me.Tb_dev(i).Text)
Next
If s > 0 Then
Select Case Index
    Case Is = 3
' houpie 20040112 début
'      If Not Cmd_resudo.Enabled  Then
      If Not Cmd_resudo.Enabled And verif_resu(Resup_do) Then
' houpie 20040112 fin
        Call reini_form(41)
        edessdo.Tram = txtVersNum(Me.Tb_dev(3).Text)
        
        Me.Chk_Qts.Value = 1
        Me.Chk_Qrin = 1
' houpie 20040112 début
        If Me.Chk_eau.Value + Me.Chk_piezo.Value + Me.Chk_charge.Value = 0 Then
            Me.Chk_eau.Value = 1
            Me.Chk_charge.Value = 1
            Me.Chk_piezo.Value = 1
        End If
' houpie 20040112 fin
        
        Me.Frame4.Enabled = True
        Me.Frame4.Visible = True
                Call OK_lignes_Click

'        If Me.Cmd_resudo.Enabled = False Then
'            Me.Cmd_VerifDo.Enabled = True
'        End If
'        Me.Chk_Qpluie.Enabled = False
'        Me.Chk_Qpluie.Value = 0
'        Me.Txtb_deversoir.Text = ""
'        Me.Frame4.Enabled = False
'        Me.Frame4.Visible = False
'        Me.SSTab1.TabEnabled(5) = False
'        Me.Cmd_resudech.Enabled = False
'        Me.Txtb_decharge.Text = ""
    End If
  Case Else
  
        Call reini_form(42)
'        Me.Cmd_resudo.Enabled = True
'        Me.Cmd_VerifDo.Enabled = false
'        Me.Frame4.Enabled = False
'        Me.Frame4.Visible = False
'        Me.SSTab1.TabEnabled(5) = False
'        Me.Cmd_resudech.Enabled = False
'        Me.Txtb_decharge.Text = ""
'        Call ini_longce
End Select
Else
        Call reini_form(40)
'        Me.Cmd_resudo.Enabled = False
'        Me.Cmd_VerifDo.Enabled = False
'        Me.Frame4.Enabled = False
'        Me.Frame4.Visible = False
'        Me.SSTab1.TabEnabled(5) = False
'        Me.Cmd_resudech.Enabled = False
'        Call ini_longce
End If
  End If
End If

 sval_champ = ""
    bKP = False

End Sub

Private Sub Tb_dev_Click(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_dev"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
Call sel_text(Tb_dev(Index))
End Sub

Private Sub Tb_dev_GotFocus(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_dev"
If SSTab1.Tab <> 4 Then
    If Me.SSTab1.TabEnabled(4) Then
        SSTab1.Tab = 4
    Else
        Me.Tb_debit(0).SetFocus
    End If
'    Me.Tb_Debit(0).SetFocus
End If
Call sel_text(Tb_dev(Index))
If change_coul Then
    Change_Couleur nom, Index
    mes = Rec_Mes(nom, Index)
    owner.affich_aide Me.Name, mes
End If

End Sub

Private Sub Tb_dev_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
     sval_champ = Tb_dev(Index).Text
    iSels = Tb_dev(Index).SelStart
    iSell = Tb_dev(Index).SelLength
    bKP = True
'   If Len(Tb_dev(Index).Text) <= Tb_dev(Index).MaxLength Then
'       Select Case Index
'            Case Is = 0
'                KeyAscii = verif_car(Tb_dev(Index).Text, KeyAscii, "Saisie longueur DO", "R")
'            Case Is = 1
'                KeyAscii = verif_car(Tb_dev(Index).Text, KeyAscii, "Saisie hauteur de la crête", "R")
'            Case Is = 2
'                KeyAscii = verif_car(Tb_dev(Index).Text, KeyAscii, "Saisie pente DO", "R")
'            Case Is = 3
'                KeyAscii = verif_car(Tb_dev(Index).Text, KeyAscii, "Saisie tirant d'eau amont admissible", "R")
'        End Select
'    End If
End If
End Sub
Private Sub Tb_debit_Change(Index As Integer)
Dim nom As String

If bKP Then
        Select Case Index
             Case Is = 0
                nom = verif_cart0(Tb_debit(Index).Text, "Saisie débit d'eau pluviale", "R")
            Case Is = 1
                nom = verif_cart0(Tb_debit(Index).Text, "Saisie débit de temps sec", "R")
            Case Is = 2
                nom = verif_cart0(Tb_debit(Index).Text, "Saisie débit de rinçage", "R")
        End Select
  If nom = "" Then
    Tb_debit(Index).Text = sval_champ
    Tb_debit(Index).SelStart = iSels
    Tb_debit(Index).SelLength = iSell
  End If
End If
'****

    Select Case Index
        Case Is = 0
            edessdo.Qpluie = txtVersNum(Me.Tb_debit(0).Text)
        Case Is = 1
            edessdo.Qts = txtVersNum(Me.Tb_debit(1).Text)
        Case Is = 2
            edessdo.Qrin = txtVersNum(Me.Tb_debit(2).Text)
    End Select
'    nombassin = ""
'    Me.Lab_bas.Caption = ""
   Call reini_form(0)
   Call reini_form(1)
'    Call reini_valeurs(0)
   sval_champ = ""
   bKP = False

End Sub

Private Sub Tb_debit_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    sval_champ = Tb_debit(Index).Text
    iSels = Tb_debit(Index).SelStart
    iSell = Tb_debit(Index).SelLength
    bKP = True
'    If Len(Tb_debit(Index).Text) <= Tb_debit(Index).MaxLength Then
'       Select Case Index
'            Case Is = 0
'                KeyAscii = verif_car(Tb_debit(Index).Text, KeyAscii, "Saisie débit d'eau pluviale", "R")
'            Case Is = 1
'                KeyAscii = verif_car(Tb_debit(Index).Text, KeyAscii, "Saisie débit de temps sec", "R")
'            Case Is = 2
'                KeyAscii = verif_car(Tb_debit(Index).Text, KeyAscii, "Saisie débit de rinçage", "R")
'        End Select
'    End If
End If
End Sub
Private Sub Tb_cont_Change(Index As Integer)
Dim nform As Integer
Dim nom As String

If bKP Then
        Select Case Index
            Case Is = 0
                nom = verif_cart0(Tb_cont(Index).Text, "Saisie cote radier amont", "R")
            Case Is = 1
                nom = verif_cart0(Tb_cont(Index).Text, "Saisie cote radier aval", "R")
            Case Is = 2
                nom = verif_cart0(Tb_cont(Index).Text, "Saisie longueur disponible", "R")
            Case Is = 3
                nom = verif_cart0(Tb_cont(Index).Text, "Saisie cote des PHE à l'exutoire", "R")
            Case Is = 4
                nom = verif_cart0(Tb_cont(Index).Text, "Saisie cote radier à l'exutoire", "R")
           Case Is = 5
                nom = verif_cart0(Tb_cont(Index).Text, "Saisie longueur de la canalisation", "R")
        End Select
  If nom = "" Then
    Tb_cont(Index).Text = sval_champ
    Tb_cont(Index).SelStart = iSels
    Tb_cont(Index).SelLength = iSell
  End If
End If
'****

If Trim(Tb_cont(Index).Text) = "" Then
    Tb_cont(Index).Text = 0#
End If
nform = 1
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
        If SSTab1.TabEnabled(2) Then
            nform = 51
        End If
        Me.Label12.Caption = Trim(Tb_cont(3).Text)
        edessdo.phex = txtVersNum(Me.Tb_cont(3).Text)
    Case Is = 4
        If SSTab1.TabEnabled(2) Then
            nform = 51
        End If
        Me.Label11.Caption = Trim(Tb_cont(4).Text)
        edessdo.rdoex = txtVersNum(Me.Tb_cont(4).Text)
    Case Is = 5
        If SSTab1.TabEnabled(2) Then
            nform = 51
        End If
        Me.Label10.Caption = Trim(Tb_cont(5).Text)
        edessdo.lgca = txtVersNum(Me.Tb_cont(5).Text)
End Select
Call reini_form(nform)
'Call reini_valeurs(2)
 sval_champ = ""
    bKP = False

End Sub

Private Sub Tb_cont_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
     sval_champ = Tb_cont(Index).Text
    iSels = Tb_cont(Index).SelStart
    iSell = Tb_cont(Index).SelLength
    bKP = True
'   If Len(Tb_cont(Index).Text) <= Tb_cont(Index).MaxLength Then
'       Select Case Index
'            Case Is = 0
'                KeyAscii = verif_car(Tb_cont(Index).Text, KeyAscii, "Saisie cote radier amont", "R")
'            Case Is = 1
'                KeyAscii = verif_car(Tb_cont(Index).Text, KeyAscii, "Saisie cote radier aval", "R")
'            Case Is = 2
'                KeyAscii = verif_car(Tb_cont(Index).Text, KeyAscii, "Saisie longueur disponible", "R")
'            Case Is = 3
'                KeyAscii = verif_car(Tb_cont(Index).Text, KeyAscii, "Saisie cote des PHE à l'exutoire", "R")
'            Case Is = 4
'                KeyAscii = verif_car(Tb_cont(Index).Text, KeyAscii, "Saisie cote radier à l'exutoire", "R")
'           Case Is = 5
'                KeyAscii = verif_car(Tb_cont(Index).Text, KeyAscii, "Saisie longueur de la canalisation", "R")
'        End Select
'    End If
End If
End Sub
Private Sub Tb_amo_Change(Index As Integer)
Dim nom As String

If bKP Then
        Select Case Index
            Case Is = 0
                nom = verif_cart0(Tb_amo(Index).Text, "Saisie diamétre canalisation amont", "I")
            Case Is = 1
                nom = verif_cart0(Tb_amo(Index).Text, "Saisie pente canalisation amont", "I")
            Case Is = 2
                nom = verif_cart0(Tb_amo(Index).Text, "Saisie coefficient canalisation amont", "I")
            Case Is = 3
                nom = verif_cart0(Tb_amo(Index).Text, "Saisie longueur canalisation amont", "R")
        End Select
  If nom = "" Then
    Tb_amo(Index).Text = sval_champ
    Tb_amo(Index).SelStart = iSels
    Tb_amo(Index).SelLength = iSell
  End If
End If
'****

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
    Call reini_form(2)
'    Call ini_resum
 sval_champ = ""
    bKP = False

End Sub

Private Sub Tb_amo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    sval_champ = Tb_amo(Index).Text
    iSels = Tb_amo(Index).SelStart
    iSell = Tb_amo(Index).SelLength
    bKP = True
'    If Len(Tb_amo(Index).Text) <= Tb_amo(Index).MaxLength Then
'       Select Case Index
'            Case Is = 0
'                KeyAscii = verif_car(Tb_amo(Index).Text, KeyAscii, "Saisie diamétre canalisation amont", "I")
'            Case Is = 1
'                KeyAscii = verif_car(Tb_amo(Index).Text, KeyAscii, "Saisie pente canalisation amont", "I")
'            Case Is = 2
'                KeyAscii = verif_car(Tb_amo(Index).Text, KeyAscii, "Saisie coefficient canalisation amont", "I")
'            Case Is = 3
'                KeyAscii = verif_car(Tb_amo(Index).Text, KeyAscii, "Saisie longueur canalisation amont", "R")
'        End Select
'    End If
End If
End Sub
Private Sub Tb_ava_Change(Index As Integer)
Dim nom As String

If bKP Then
        Select Case Index
            Case Is = 0
        nom = verif_cart0(Tb_ava(Index).Text, "Saisie diamétre canalisation aval", "I")
            Case Is = 1
        nom = verif_cart0(Tb_ava(Index).Text, "Saisie pente canalisation aval", "I")
            Case Is = 2
        nom = verif_cart0(Tb_ava(Index).Text, "Saisie coefficient canalisation aval", "I")
            Case Is = 3
        nom = verif_cart0(Tb_ava(Index).Text, "Saisie longueur canalisation aval", "R")
        End Select
  If nom = "" Then
    Tb_ava(Index).Text = sval_champ
    Tb_ava(Index).SelStart = iSels
    Tb_ava(Index).SelLength = iSell
  End If
End If
'****

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
    Call reini_form(3)
'    Call ini_resuv
 sval_champ = ""
    bKP = False

End Sub

Private Sub Tb_ava_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
     sval_champ = Tb_ava(Index).Text
    iSels = Tb_ava(Index).SelStart
    iSell = Tb_ava(Index).SelLength
    bKP = True
'   If Len(Tb_ava(Index).Text) <= Tb_ava(Index).MaxLength Then
'       Select Case Index
'            Case Is = 0
'        KeyAscii = verif_car(Tb_ava(Index).Text, KeyAscii, "Saisie diamétre canalisation aval", "I")
'            Case Is = 1
'        KeyAscii = verif_car(Tb_ava(Index).Text, KeyAscii, "Saisie pente canalisation aval", "I")
'            Case Is = 2
'        KeyAscii = verif_car(Tb_ava(Index).Text, KeyAscii, "Saisie coefficient canalisation aval", "I")
'            Case Is = 3
'        KeyAscii = verif_car(Tb_ava(Index).Text, KeyAscii, "Saisie longueur canalisation aval", "R")
'        End Select
'    End If
End If
End Sub

Public Sub Mquitter()
    MnuQuit_Click
End Sub
Public Sub Mquit()
    m_quitter_Click
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
Public Sub Init_ss_commentaire()
    owner.affich_aide Me.Name, ""  'Déversoir d'orage"
End Sub
Private Sub Tb_dev_LostFocus(Index As Integer)
Dim ok As Boolean
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_dev", Index, txtVersNum(Tb_dev(Index).Text))
    If Not ok Then
        Tb_dev(Index).SetFocus
        DoEvents
    End If
    okg = True
End If

End Sub

Private Sub Tb_titre_Change()
    Me.Caption = fen_titre + " : " + Tb_titre.Text
End Sub
