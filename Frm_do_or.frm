VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Begin VB.Form Frm_do_or 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hydraulique Déversoir d'Orage à ouverture de radier"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9825
   Icon            =   "Frm_do_or.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   9825
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   459
      TabCaption(0)   =   "Bassin Versant"
      TabPicture(0)   =   "Frm_do_or.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frm_bv"
      Tab(0).Control(1)=   "Cmd_Sel_Bv"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Lab_bas"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Contraintes"
      TabPicture(1)   =   "Frm_do_or.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Canalisations"
      TabPicture(2)   =   "Frm_do_or.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(1)=   "Lb_mesm"
      Tab(2).Control(2)=   "Frm_condam"
      Tab(2).Control(3)=   "Frm_dech"
      Tab(2).Control(4)=   "Frm_condav"
      Tab(2).Control(5)=   "Frm_contraintes"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Déversoir"
      TabPicture(3)   =   "Frm_do_or.frx":091E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Lb_mesv"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Txtb_resu"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frm_dev"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Txtb_deversoir"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Frame4"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Cmd_recalc"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Cmd_mini"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Cmd_Leaping_Wear"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Frm_Ouverture"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).ControlCount=   9
      Begin VB.Frame Frm_Ouverture 
         Caption         =   "Caractéristiques de l'ouverture"
         Height          =   1320
         Left            =   240
         TabIndex        =   133
         Top             =   370
         Width           =   3350
         Begin VB.TextBox Tb_larg 
            Height          =   285
            Left            =   915
            Locked          =   -1  'True
            TabIndex        =   136
            Top             =   360
            Width           =   750
         End
         Begin VB.TextBox Tb_long 
            Height          =   285
            Left            =   915
            TabIndex        =   135
            Top             =   720
            Width           =   750
         End
         Begin ComCtl2.UpDown UpDown1 
            Height          =   255
            Left            =   2160
            TabIndex        =   134
            Top             =   375
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   327681
            Max             =   10000
            Enabled         =   -1  'True
         End
         Begin VB.Label Lb_intLarg 
            Caption         =   "Largeur"
            Height          =   255
            Left            =   120
            TabIndex        =   140
            Top             =   375
            Width           =   615
         End
         Begin VB.Label Lb_intLong 
            Caption         =   "Longueur"
            Height          =   255
            Left            =   120
            TabIndex        =   139
            Top             =   735
            Width           =   735
         End
         Begin VB.Label Lb_uLarg 
            Caption         =   "m"
            Height          =   255
            Left            =   1800
            TabIndex        =   138
            Top             =   375
            Width           =   855
         End
         Begin VB.Label Lb_uLong 
            Caption         =   "m"
            Height          =   255
            Left            =   1800
            TabIndex        =   137
            Top             =   735
            Width           =   615
         End
      End
      Begin VB.CommandButton Cmd_Leaping_Wear 
         Caption         =   "Calcul initialisation"
         Height          =   255
         Left            =   1870
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1695
      End
      Begin VB.CommandButton Cmd_mini 
         Caption         =   "Dessin mini"
         Height          =   255
         Left            =   1990
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   3720
         Visible         =   0   'False
         Width           =   1570
      End
      Begin VB.CommandButton Cmd_recalc 
         Caption         =   "Dessin recalculé"
         Height          =   255
         Left            =   240
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   3720
         Visible         =   0   'False
         Width           =   1570
      End
      Begin VB.Frame Frm_contraintes 
         Caption         =   "Contraintes"
         Height          =   1750
         Left            =   -70200
         TabIndex        =   73
         Top             =   2150
         Width           =   4575
         Begin VB.TextBox Tb_cont 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   3120
            MaxLength       =   7
            TabIndex        =   16
            Top             =   600
            Width           =   900
         End
         Begin VB.TextBox Tb_hmin 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   3105
            MaxLength       =   8
            TabIndex        =   17
            Top             =   1200
            Width           =   900
         End
         Begin VB.Label Lb_intcont 
            Caption         =   "Conduite arrivée : cote radier aval"
            Height          =   300
            Index           =   1
            Left            =   360
            TabIndex        =   77
            Top             =   645
            Width           =   2655
         End
         Begin VB.Label Lb_ucont 
            Caption         =   "m"
            Height          =   255
            Index           =   1
            Left            =   4125
            TabIndex        =   76
            Top             =   645
            Width           =   195
         End
         Begin VB.Label Lb_uHmin 
            Caption         =   "m"
            Height          =   255
            Left            =   4125
            TabIndex        =   75
            Top             =   1240
            Width           =   255
         End
         Begin VB.Label Lb_intHmin 
            Caption         =   "Hauteur entre canalisations départ et déversement"
            Height          =   495
            Left            =   360
            TabIndex        =   74
            Top             =   1125
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Dessin lignes"
         Height          =   735
         Left            =   240
         TabIndex        =   68
         Top             =   2880
         Width           =   3350
         Begin VB.CheckBox Chk_cri 
            Caption         =   "Débit de référence"
            Height          =   255
            Left            =   1560
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton OK_lignes 
            Caption         =   "OK"
            Height          =   255
            Left            =   2760
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   372
         End
         Begin VB.CheckBox Chk_max 
            Caption         =   "Débit d'orage"
            Height          =   192
            Left            =   120
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   360
            Width           =   1250
         End
      End
      Begin VB.TextBox Txtb_deversoir 
         BackColor       =   &H80000016&
         Height          =   2295
         Left            =   3720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   480
         Width           =   5655
      End
      Begin VB.Frame Frm_condav 
         Caption         =   "Conduite de départ"
         Height          =   1750
         Left            =   -74640
         TabIndex        =   60
         Top             =   2150
         Width           =   4245
         Begin VB.TextBox Tb_ava 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   3
            Left            =   1800
            MaxLength       =   8
            TabIndex        =   11
            Top             =   1320
            Width           =   900
         End
         Begin VB.TextBox Tb_ava 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   10
            Top             =   960
            Width           =   900
         End
         Begin VB.TextBox Tb_ava 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   9
            Top             =   600
            Width           =   900
         End
         Begin VB.TextBox Tb_ava 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   1800
            MaxLength       =   6
            TabIndex        =   8
            Top             =   240
            Width           =   900
         End
         Begin VB.CommandButton Cmd_ava 
            Caption         =   "Courbe..."
            Height          =   255
            Left            =   3120
            TabIndex        =   121
            TabStop         =   0   'False
            ToolTipText     =   "Courbe de débit de la conduite aval"
            Top             =   1200
            Width           =   990
         End
         Begin VB.Label Lb_intava 
            Caption         =   "Longueur "
            Height          =   300
            Index           =   3
            Left            =   120
            TabIndex        =   117
            Top             =   1365
            Width           =   1455
         End
         Begin VB.Label Lb_intava 
            Caption         =   "Coeff. de  Strickler"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   116
            Top             =   1005
            Width           =   1455
         End
         Begin VB.Label Lb_intava 
            Caption         =   "Pente "
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   115
            Top             =   645
            Width           =   1455
         End
         Begin VB.Label Lb_intava 
            Caption         =   "Diamètre "
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   114
            Top             =   285
            Width           =   1455
         End
         Begin VB.Label Lb_uava 
            Caption         =   "mm"
            Height          =   255
            Index           =   0
            Left            =   2800
            TabIndex        =   118
            Top             =   285
            Width           =   735
         End
         Begin VB.Label Lb_uava 
            Caption         =   "1/10000"
            Height          =   255
            Index           =   1
            Left            =   2800
            TabIndex        =   119
            Top             =   645
            Width           =   735
         End
         Begin VB.Label Lb_uava 
            Caption         =   "m"
            Height          =   255
            Index           =   3
            Left            =   2805
            TabIndex        =   120
            Top             =   1365
            Width           =   255
         End
         Begin VB.Label Lb_uava 
            Height          =   255
            Index           =   2
            Left            =   3240
            TabIndex        =   61
            Top             =   1125
            Width           =   615
         End
      End
      Begin VB.Frame Frm_dech 
         Caption         =   "Conduite de déversement"
         Height          =   1750
         Left            =   -70080
         TabIndex        =   57
         Top             =   360
         Width           =   4245
         Begin VB.TextBox Tb_dech 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   1800
            MaxLength       =   6
            TabIndex        =   12
            Top             =   240
            Width           =   900
         End
         Begin VB.TextBox Tb_dech 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   13
            Top             =   600
            Width           =   900
         End
         Begin VB.TextBox Tb_dech 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   14
            Top             =   960
            Width           =   900
         End
         Begin VB.TextBox Tb_dech 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   3
            Left            =   1800
            MaxLength       =   8
            TabIndex        =   15
            Top             =   1320
            Width           =   900
         End
         Begin VB.CommandButton Cmd_dech 
            Caption         =   "Courbe..."
            Height          =   255
            Left            =   3120
            TabIndex        =   129
            TabStop         =   0   'False
            ToolTipText     =   "Courbe de débit de la conduite de décharge"
            Top             =   1200
            Width           =   990
         End
         Begin VB.Label Lb_intdech 
            Caption         =   "Diamètre "
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   122
            Top             =   285
            Width           =   1455
         End
         Begin VB.Label Lb_intdech 
            Caption         =   "Pente "
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   123
            Top             =   645
            Width           =   1455
         End
         Begin VB.Label Lb_intdech 
            Caption         =   "Coeff. de Strickler"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   124
            Top             =   1005
            Width           =   1455
         End
         Begin VB.Label Lb_intdech 
            Caption         =   "Longueur  "
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   125
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Lb_udech 
            Caption         =   "mm"
            Height          =   255
            Index           =   0
            Left            =   2800
            TabIndex        =   126
            Top             =   285
            Width           =   735
         End
         Begin VB.Label Lb_udech 
            Caption         =   "1/10000"
            Height          =   255
            Index           =   1
            Left            =   2800
            TabIndex        =   127
            Top             =   645
            Width           =   735
         End
         Begin VB.Label Lb_udech 
            Caption         =   "m"
            Height          =   255
            Index           =   3
            Left            =   2805
            TabIndex        =   128
            Top             =   1365
            Width           =   255
         End
         Begin VB.Label Lb_udech 
            Height          =   255
            Index           =   2
            Left            =   3360
            TabIndex        =   59
            Top             =   1125
            Width           =   495
         End
         Begin VB.Label Lb_udech 
            Height          =   195
            Index           =   4
            Left            =   3360
            TabIndex        =   58
            Top             =   2085
            Width           =   255
         End
      End
      Begin VB.Frame Frm_condam 
         Caption         =   "Conduite d'arrivée"
         Height          =   1750
         Left            =   -74640
         TabIndex        =   49
         Top             =   360
         Width           =   4245
         Begin VB.CommandButton Cmd_amo 
            Caption         =   "Courbe..."
            Height          =   255
            Left            =   3120
            TabIndex        =   113
            TabStop         =   0   'False
            ToolTipText     =   "Courbe de débit de la conduite amont"
            Top             =   1200
            Width           =   990
         End
         Begin VB.TextBox Tb_amo 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   3
            Left            =   1800
            MaxLength       =   8
            TabIndex        =   7
            Top             =   1320
            Width           =   900
         End
         Begin VB.TextBox Tb_amo 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   6
            Top             =   960
            Width           =   900
         End
         Begin VB.TextBox Tb_amo 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   5
            Top             =   600
            Width           =   900
         End
         Begin VB.TextBox Tb_amo 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   1800
            MaxLength       =   6
            TabIndex        =   4
            Top             =   240
            Width           =   900
         End
         Begin VB.Label Lb_uamo 
            Height          =   255
            Index           =   2
            Left            =   3240
            TabIndex        =   55
            Top             =   1125
            Width           =   495
         End
         Begin VB.Label Lb_uamo 
            Caption         =   "m"
            Height          =   255
            Index           =   3
            Left            =   2800
            TabIndex        =   112
            Top             =   1365
            Width           =   255
         End
         Begin VB.Label Lb_uamo 
            Caption         =   "1/10000"
            Height          =   255
            Index           =   1
            Left            =   2800
            TabIndex        =   111
            Top             =   645
            Width           =   735
         End
         Begin VB.Label Lb_uamo 
            Caption         =   "mm"
            Height          =   255
            Index           =   0
            Left            =   2800
            TabIndex        =   110
            Top             =   285
            Width           =   735
         End
         Begin VB.Label Lb_intamo 
            Caption         =   "Longueur "
            Height          =   300
            Index           =   3
            Left            =   120
            TabIndex        =   109
            Top             =   1365
            Width           =   1455
         End
         Begin VB.Label Lb_intamo 
            Caption         =   "Coeff.  de  Strickler"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   108
            Top             =   1005
            Width           =   1455
         End
         Begin VB.Label Lb_intamo 
            Caption         =   "Pente "
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   107
            Top             =   645
            Width           =   1455
         End
         Begin VB.Label Lb_intamo 
            Caption         =   "Diamètre "
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   106
            Top             =   285
            Width           =   1455
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
         Height          =   3135
         Left            =   -74880
         TabIndex        =   21
         Top             =   600
         Width           =   3975
         Begin VB.TextBox Tb_cont 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   2760
            MaxLength       =   7
            TabIndex        =   101
            TabStop         =   0   'False
            Text            =   "100.00"
            Top             =   2400
            Width           =   900
         End
         Begin VB.TextBox Tb_cont 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Index           =   4
            Left            =   2580
            MaxLength       =   7
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   1800
            Width           =   900
         End
         Begin VB.TextBox Tb_cont 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Index           =   3
            Left            =   2580
            MaxLength       =   7
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   1440
            Width           =   900
         End
         Begin VB.TextBox Tb_cont 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Index           =   2
            Left            =   2580
            MaxLength       =   8
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label Lb_ucont 
            Caption         =   "m"
            Height          =   255
            Index           =   0
            Left            =   3720
            TabIndex        =   105
            Top             =   2400
            Width           =   195
         End
         Begin VB.Label Lb_intcont 
            Caption         =   "Conduite arrivée : cote radier amont"
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   97
            Top             =   2400
            Width           =   2655
         End
         Begin VB.Label Lb_ucont 
            Caption         =   "m"
            Height          =   255
            Index           =   4
            Left            =   3600
            TabIndex        =   104
            Top             =   1850
            Width           =   200
         End
         Begin VB.Label Lb_ucont 
            Caption         =   "m"
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   103
            Top             =   1490
            Width           =   200
         End
         Begin VB.Label Lb_ucont 
            Caption         =   "m"
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   102
            Top             =   1130
            Width           =   200
         End
         Begin VB.Label Lb_intcont 
            Caption         =   "Cote de radier obligé à l'exutoire"
            Height          =   300
            Index           =   4
            Left            =   200
            TabIndex        =   96
            Top             =   1850
            Width           =   2295
         End
         Begin VB.Label Lb_intcont 
            Caption         =   "Longueur disponible"
            Height          =   300
            Index           =   2
            Left            =   200
            TabIndex        =   94
            Top             =   1130
            Width           =   2295
         End
         Begin VB.Label Lb_intcont 
            Caption         =   "Cote des P.H.E. à l'exutoire"
            Height          =   300
            Index           =   3
            Left            =   200
            TabIndex        =   95
            Top             =   1490
            Width           =   2295
         End
      End
      Begin VB.Frame Frm_bv 
         Caption         =   "Hydraulique du B.V. "
         Height          =   2055
         Left            =   -70200
         TabIndex        =   20
         Top             =   1080
         Width           =   4575
         Begin VB.TextBox Tb_debit 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   2925
            MaxLength       =   8
            TabIndex        =   1
            Top             =   480
            Width           =   900
         End
         Begin VB.TextBox Tb_debit 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   2925
            MaxLength       =   8
            TabIndex        =   2
            Top             =   960
            Width           =   900
         End
         Begin VB.TextBox Tb_debit 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   2925
            MaxLength       =   8
            TabIndex        =   3
            Top             =   1320
            Width           =   900
         End
         Begin VB.Label Lb_udebit 
            Caption         =   "l/s"
            Height          =   255
            Index           =   2
            Left            =   3960
            TabIndex        =   93
            Top             =   1365
            Width           =   405
         End
         Begin VB.Label Lb_udebit 
            Caption         =   "l/s"
            Height          =   255
            Index           =   1
            Left            =   3960
            TabIndex        =   92
            Top             =   1005
            Width           =   405
         End
         Begin VB.Label Lb_udebit 
            Caption         =   "l/s"
            Height          =   255
            Index           =   0
            Left            =   3960
            TabIndex        =   91
            Top             =   540
            Width           =   405
         End
         Begin VB.Label Lb_intdebit 
            Caption         =   "Débit d'eau pluviale "
            Height          =   495
            Index           =   0
            Left            =   600
            TabIndex        =   88
            Top             =   450
            Width           =   2055
         End
         Begin VB.Label Lb_intdebit 
            Caption         =   "Débit de temps sec "
            Height          =   300
            Index           =   1
            Left            =   600
            TabIndex        =   89
            Top             =   1005
            Width           =   2055
         End
         Begin VB.Label Lb_intdebit 
            Caption         =   "Débit de référence "
            Height          =   300
            Index           =   2
            Left            =   600
            TabIndex        =   90
            Top             =   1365
            Width           =   2055
         End
      End
      Begin VB.CommandButton Cmd_Sel_Bv 
         Caption         =   "Sélection d'un bassin versant"
         Height          =   255
         Left            =   -74520
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Frame Frm_dev 
         Caption         =   "Caractéristiques de la chambre de déversoir"
         Height          =   930
         Left            =   240
         TabIndex        =   63
         Top             =   1840
         Width           =   3350
         Begin VB.TextBox Tb_dev 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   915
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   18
            Top             =   405
            Width           =   750
         End
         Begin VB.TextBox Tb_dev 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   1395
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   19
            Top             =   885
            Visible         =   0   'False
            Width           =   750
         End
         Begin ComCtl2.UpDown UpDown2 
            Height          =   255
            Left            =   2160
            TabIndex        =   130
            Top             =   435
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   327681
            Increment       =   10
            Max             =   10000
            Enabled         =   -1  'True
         End
         Begin ComCtl2.UpDown UpDown3 
            Height          =   255
            Left            =   2880
            TabIndex        =   131
            Top             =   840
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   327681
            Max             =   10000
            Enabled         =   -1  'True
         End
         Begin VB.Label Lb_uLchambre 
            Caption         =   "m"
            Height          =   255
            Left            =   1800
            TabIndex        =   132
            Top             =   465
            Width           =   375
         End
         Begin VB.Label Label17 
            Caption         =   "<="
            Height          =   195
            Left            =   1080
            TabIndex        =   86
            Top             =   930
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label16 
            Caption         =   "<"
            Height          =   195
            Left            =   2280
            TabIndex        =   85
            Top             =   930
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label15 
            Caption         =   ">="
            Height          =   195
            Left            =   2160
            TabIndex        =   84
            Top             =   165
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label14 
            Caption         =   ">="
            Height          =   195
            Left            =   1920
            TabIndex        =   83
            Top             =   165
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Lb_udev0 
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   82
            Top             =   930
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label Lb_udev0 
            Height          =   195
            Index           =   0
            Left            =   2640
            TabIndex        =   81
            Top             =   165
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label Lb_intdev 
            Caption         =   "Longueur "
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   67
            Top             =   465
            Width           =   855
         End
         Begin VB.Label Lb_intdev 
            Caption         =   "Profondeur "
            Height          =   200
            Index           =   1
            Left            =   120
            TabIndex        =   66
            Top             =   880
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Lb_udev 
            Height          =   195
            Index           =   0
            Left            =   1440
            TabIndex        =   65
            Top             =   360
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label Lb_udev 
            Height          =   195
            Index           =   1
            Left            =   2670
            TabIndex        =   64
            Top             =   930
            Visible         =   0   'False
            Width           =   600
         End
      End
      Begin VB.TextBox Txtb_resu 
         BackColor       =   &H80000016&
         Height          =   735
         Left            =   3720
         MultiLine       =   -1  'True
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   2880
         Width           =   5655
      End
      Begin VB.Label Lab_bas 
         Height          =   1335
         Left            =   -74520
         TabIndex        =   56
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label Lb_mesv 
         Height          =   300
         Left            =   285
         TabIndex        =   52
         Top             =   5760
         Width           =   9735
      End
      Begin VB.Label Lb_mesm 
         Height          =   300
         Left            =   -74715
         TabIndex        =   51
         Top             =   5760
         Width           =   9735
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   135
         Left            =   -74520
         TabIndex        =   50
         Top             =   5760
         Width           =   15
      End
   End
   Begin VB.TextBox Tb_titre 
      Height          =   300
      Left            =   5520
      MaxLength       =   30
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   4000
   End
   Begin VB.ComboBox Cb_deversoir 
      Height          =   315
      Left            =   360
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   360
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
Attribute VB_Name = "Frm_do_or"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private okg As Boolean
Private owner As MDIFrm_menu
Private ok_imp As Boolean
Private ok_longueur As Boolean
Private ok_hauteur As Boolean
Private ok_largeur As Boolean
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
Private ecoulam As String
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
Private Function Rec_Mes0(nom As String, Index As Integer)
Dim mes As String
mes = ""
Select Case nom
    Case Is = "Frm_bv"
                 mes = "Hydraulique du BV"
    Case Is = "Lb_intdebit", "Tb_debit"
                mes = "Données de base"
        Select Case Index
            Case Is = 0
                mes = "Débit d'orage des eaux pluviales"
            Case Is = 1
                mes = "Débit de temps sec"
            Case Is = 2
                mes = "Débit de référence"
        End Select
   Case Is = "Frm_condam"
                 mes = "Conduite arrivée"
    Case Is = "Lb_intamo", "Tb_amo"
        Select Case Index
            Case Is = 0
                mes = "Diamètre"
            Case Is = 1
                mes = "Pente"
            Case Is = 2
                mes = "Coefficient de Manning-Strickler"
            Case Is = 3
                mes = "Longueur"
        End Select
   Case Is = "Frm_condav"
                mes = "Conduite départ"
   Case Is = "Lb_intava", "Tb_ava"
        Select Case Index
            Case Is = 0
                mes = "Diamètre"
            Case Is = 1
                mes = "Pente"
            Case Is = 2
                mes = "Coefficient de Manning-Strickler"
            Case Is = 3
                mes = "Longueur"
        End Select
   Case Is = "Frm_dech"
                mes = "Conduite de déversement"
   Case Is = "Lb_intdech", "Tb_dech"
        Select Case Index
            Case Is = 0
                mes = "Diamètre"
            Case Is = 1
                mes = "Pente"
            Case Is = 2
                mes = "Coefficient de Manning-Strickler"
            Case Is = 3
                mes = "Longueur"
        End Select
   Case Is = "Lb_intcont", "Tb_cont"
        Select Case Index
            Case Is = 1
                mes = "cote radier aval,conduite arrivée"
        End Select
   Case Is = "Lb_intHmin", "Tb_hmin"
                mes = "Hauteur entre canalisations"
   Case Is = "Frm_dev"
                mes = "Caractéristiques"
   Case Is = "Lb_intdev", "Tb_dev"
        Select Case Index
            Case Is = 0
                mes = "Longueur"
            Case Is = 1
                mes = "Profondeur"
        End Select
   Case Is = "Lb_intLarg", "Tb_larg"
                mes = "Largeur ouverture"
   Case Is = "Lb_intLong", "Tb_long"
                mes = "Longueur ouverture"
   Case Is = "Frame4"
                mes = "Dessin lignes"
   Case Is = "Chk_cri"
                mes = "Débit critique"
   Case Is = "Chk_max"
                mes = "Débit d'orage"
   Case Is = "Cmd_recalc"
                mes = "Dessin recalculé"
   Case Is = "Cmd_mini"
                mes = "Dessin mini"
End Select
'''ligne à supprimer si aide
                mes = ""
'''''''''
mes_prec = mes
Rec_Mes0 = mes
End Function
Private Function Rec_Mes(nom As String, Index As Integer)
Dim mes As String
mes = ""
' voir aussi SStab1
Select Case nom
    Case Is = "Lb_intdebit", "Tb_debit", "Frm_bv", "Cmd_Sel_Bv"
                mes = IDhlp_DOORDonneesBase  '"Hydraulique du bassin versant"
    Case Is = "Lb_intcont", "Tb_cont", "Lb_intHmin", "Tb_hmin", "Frm_contraintes"
                ' ajouter hauteur entre canalisation Tb et Lb
                mes = IDhlp_DOORContraintes  '"Déversoir à seuil haut"
    Case Is = "Lb_intamo", "Tb_amo", "Frm_condam"
                mes = IDhlp_DOORConduiteArrivee  '"Conduite d'arrivée"
    Case Is = "Lb_intava", "Tb_ava", "Frm_condav"
                mes = IDhlp_DOORConduiteDepart  '"Conduite de départ"
    Case Is = "Lb_intdev", "Tb_dev", "Frm_dev", "Tb_larg", "Lb_intLarg", "Tb_long", "Lb_intLong", "Frm_Ouverture"
         ' ajouter "Tb_larg Tb_long et lb associés
                mes = IDhlp_DOOROuvrageDeversoir  '"L'ouvrage déversoir"
    Case Is = "Lb_intdech", "Tb_dech", "Frm_dech"
                mes = IDhlp_DOORConduiteDeversement  '"Conduite de décharge"
     Case Is = "Txtb_deversoir"
                mes = IDhlp_DOORMethodeDimensionnement  '"Méthode de dimensionnement"
               
                'Txtb_deversoir
End Select
mes_prec = mes
Rec_Mes = mes
End Function


Public Function get_l_tb() As Variant
get_l_tb = list_tb
End Function
Private Sub init_l_tab()
Dim l0() As Variant, l1() As Variant, l2() As Variant
Dim l3() As Variant
l0 = Array(0)
'l1 = Array(0, "TB_debit")
'l2 = Array(0, "TB_cont")
'l3 = Array(0, "TB_amo")
'l4 = Array(0, "TB_ava")
l1 = Array(0, "TB_debit")
l2 = Array(0, "TB_amo", "TB_ava", "TB_dech", "TB_cont", "TB_hmin")
l3 = Array(0, "TB_dev")
'l6 = Array(0, "CHK_eau", "CHK_piezo", "CHK_charge", "CHK_qts", "CHK_qrin", "CHK_qpluie", "OK_lignes")
'ReDim list_tb(0 To UBound(l0), 0 To UBound(l1), 0 To UBound(l2), 0 To UBound(l3), 0 To UBound(l4), 0 To UBound(l5), 0 To UBound(l6), 0 To UBound(l7))
'list_tb = Array(l0, l1, l2, l3, l4, l5, l6, l7)
'ReDim list_tb(0 To UBound(l0), 0 To UBound(l1), 0 To UBound(l2), 0 To UBound(l3), 0 To UBound(l4), 0 To UBound(l7))
'list_tb = Array(l0, l1, l2, l3, l4, l7)
ReDim list_tb(0 To UBound(l0), 0 To UBound(l1), 0 To UBound(l2), 0 To UBound(l3))
list_tb = Array(l0, l1, l2, l3)
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
'            Me.SSTab1.TabEnabled(1) = True
          Else
''           owner.fdessin.UC_graphiqueB.ecr_texta 1665, 1140, "Surface", "G", "B"
''           owner.fdessin.UC_graphiqueB.ecr_texta 1620, 1530, "Coef. de ruissellement", "G", "B"
''           owner.fdessin.UC_graphiqueB.ecr_texta 2055, 1995, "Nombre d'habitants", "G", "B"
''           owner.fdessin.UC_graphiqueB.ecr_texta 2160, 2505, "Taux de dilution", "G", "B"
''           owner.fdessin.UC_graphiqueB.ecr_texta 2875, 540, "Longueur", "G", "B"
''           owner.fdessin.UC_graphiqueB.ecr_texta 3145, 810, "Pente", "G", "B"
            Me.Frm_bv.Caption = "Hydraulique B.V.  " + Trim(nombassin)
            Me.Lab_bas.Caption = ""
'            Me.SSTab1.TabEnabled(1) = False
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
Private Sub Chk_cri_Click()
Call OK_lignes_Click
End Sub

Private Sub Chk_max_Click()
Call OK_lignes_Click
End Sub

Private Sub Cmd_amo_Click()
Call dessin_courbe_amo
End Sub

Private Sub Cmd_ava_Click()
Call dessin_courbe_ava
End Sub

Private Sub Cmd_dech_Click()
Call dessin_courbe_dech
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

Sub dessin_door(Optional ByVal okmax As Boolean, Optional ByVal okcri As Boolean)
    Call init_graphdoor(owner.fdessin.UC_graphique1)
    Call init_graphdoor(Frm_desprint.UC_graphique1)
    Call dessin_door_objet(owner.fdessin.UC_graphique1)
    Call dessin_door_objet(Frm_desprint.UC_graphique1)
If okmax Then
    Call dess_debit_max_or(owner.fdessin.UC_graphique1)
    Call dess_debit_max_or(Frm_desprint.UC_graphique1)
End If
If okcri Then
    Call dess_debit_cri_or(owner.fdessin.UC_graphique1)
    Call dess_debit_cri_or(Frm_desprint.UC_graphique1)
End If
End Sub
Sub dessin_door_objet(ByRef uc_g As UC_graphique)
Call dess_troncon_or(uc_g, edessdo.tron_amo, couleur.gris, "D")
'Call dess_door(uc_g, edo, couleur.noir)
Call dess_troncon_or(uc_g, edessdo.tron_amo, couleur.gris, "F")
Call dess_troncon_or(uc_g, edessdo.tron_amo, couleur.gris, "C")
Call dess_door(uc_g, edo, couleur.noir)
'Call dess_cot(uc_g, couleur.noir) ' vbBlack)
End Sub

Private Sub lect_fich()
Dim za As st_savdoor
Dim za1 As st_savdoor1
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
        za = za1.stsavdoor
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
Public Sub cre_list_don4()
Dim i As Integer
ReDim list_don4(13, 3)
i = 0
    list_don4(i, 1) = "------Résultats pour le débit critique " + ajout_zero(Trim(str(Round(edessdo.Qrin / 1000, 3)))) + " m3/s"
    list_don4(i, 2) = ""
    list_don4(i, 3) = ""
i = i + 1
    list_don4(i, 1) = "Hauteur d'eau à l'amont"
    list_don4(i, 2) = Round(edoor_res.Ham_cri, 3)
    list_don4(i, 3) = " m"
i = i + 1
    list_don4(i, 1) = "Vitesse à l'amont"
    list_don4(i, 2) = Round(edoor_res.Vam_cri, 3)
    list_don4(i, 3) = " m/s"
i = i + 1
    list_don4(i, 1) = "Hauteur d'eau à l'ouverture"
    list_don4(i, 2) = Round(edoor_res.hc_cri, 3)
    list_don4(i, 3) = " m"
i = i + 1
    list_don4(i, 1) = "Vitesse à l'ouverture"
    list_don4(i, 2) = Round(edoor_res.vc_cri, 3)
    list_don4(i, 3) = " m/s"
i = i + 1
    list_don4(i, 1) = "------Résultats pour le débit d'orage " + ajout_zero(Trim(str(Round(edessdo.Qpluie / 1000, 3)))) + " m3/s"
    list_don4(i, 2) = ""
    list_don4(i, 3) = ""
i = i + 1
    list_don4(i, 1) = "Hauteur d'eau à l'amont"
    list_don4(i, 2) = Round(edoor_res.Ham, 3)
    list_don4(i, 3) = " m"
i = i + 1
    list_don4(i, 1) = "Vitesse à l'amont"
    list_don4(i, 2) = Round(edoor_res.Vam, 3)
    list_don4(i, 3) = " m/s"
i = i + 1
    list_don4(i, 1) = "Hauteur d'eau à l'ouverture"
    list_don4(i, 2) = Round(edoor_res.hc, 3)
    list_don4(i, 3) = " m"
i = i + 1
    list_don4(i, 1) = "Vitesse à l'ouverture"
    list_don4(i, 2) = Round(edoor_res.vc, 3)
    list_don4(i, 3) = " m/s"
i = i + 1
    list_don4(i, 1) = " "
    list_don4(i, 2) = " "
    list_don4(i, 3) = " "
i = i + 1
    list_don4(i, 1) = "Longueur de l'ouverture"
    list_don4(i, 2) = Round(edoor_res.l_ouverture, 3)
    list_don4(i, 3) = " m"
'i = i + 1
'    list_don4(i, 1) = "Débit conservé à l'aval: théorique"
'    list_don4(i, 2) = Round(edoor_res.Qbavth, 3)
'    list_don4(i, 3) = " m3/s"
i = i + 1
    list_don4(i, 1) = "ratio Qconservé sur Q critique : "
    
    list_don4(i, 2) = Round(((1000 * edoor_res.Qbaveff - edessdo.Qrin) / edessdo.Qrin), 3)
    list_don4(i, 3) = "  "
i = i + 1
    list_don4(i, 1) = "Débit conservé à l'aval: effectif"
    list_don4(i, 2) = Round(edoor_res.Qbaveff, 3)
    list_don4(i, 3) = " m3/s"

End Sub
Public Sub cre_list_don3()
Dim i As Integer
Dim qv As deb_vit

ReDim list_don3(Tb_amo.count + 2, 7)
    list_don3(0, 2) = "Amont"
    list_don3(0, 4) = "Aval"
    list_don3(0, 6) = "Décharge"

For i = 0 To Tb_amo.count - 2
    list_don3(i + 1, 1) = Lb_intamo(i).Caption
    list_don3(i + 1, 2) = Tb_amo(i).Text
    list_don3(i + 1, 3) = Lb_uamo(i).Caption
    list_don3(i + 1, 4) = Tb_ava(i).Text
    list_don3(i + 1, 5) = Lb_uava(i).Caption
    list_don3(i + 1, 6) = Tb_dech(i).Text
    list_don3(i + 1, 7) = Lb_udech(i).Caption
Next
resudev.vpsm = ajout_zero(Trim(str(Round(qv.vitesse, 3))))
resudev.dpsm = ajout_zero(Trim(str(Round(qv.debit, 3))))

'list_don3(4, 4) = resudev.longetranglee
i = i + 1
    list_don3(i, 1) = "Vitesse pleine section"
    qv = debvit_ps(edessdo.tron_amo.conduit)
    list_don3(i, 2) = ajout_zero(Trim(str(Round(qv.vitesse, 3))))
    list_don3(i, 3) = "m/s"
    qv = debvit_ps(edessdo.tron_ava.conduit)
    list_don3(i, 4) = ajout_zero(Trim(str(Round(qv.vitesse, 3))))
    list_don3(i, 5) = "m/s"
    qv = debvit_ps(edessdo.tron_dech.conduit)
    list_don3(i, 6) = ajout_zero(Trim(str(Round(qv.vitesse, 3))))
    list_don3(i, 7) = "m/s"
i = i + 1
    list_don3(i, 1) = "Débit pleine section"
    qv = debvit_ps(edessdo.tron_amo.conduit)
    list_don3(i, 2) = ajout_zero(Trim(str(Round(qv.debit, 3))))
    list_don3(i, 3) = "m3/s"
    qv = debvit_ps(edessdo.tron_ava.conduit)
    list_don3(i, 4) = ajout_zero(Trim(str(Round(qv.debit, 3))))
    list_don3(i, 5) = "m3/s"
    qv = debvit_ps(edessdo.tron_dech.conduit)
    list_don3(i, 6) = ajout_zero(Trim(str(Round(qv.debit, 3))))
    list_don3(i, 7) = "m3/s"
End Sub
Public Sub cre_list_don2()
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
Dim i As Integer
ReDim list_don5(2, 3)
i = 0
    list_don5(i, 1) = "Longueur de la chambre "
'    list_don5(i, 2) = Round(edoor_res.l_chambre1, 3)
    list_don5(i, 2) = Round(txtVersNum(Me.Lb_udev0(0).Caption), 3)
    list_don5(i, 3) = " m"
i = i + 1
    list_don5(i, 1) = "Hauteur de la chambre"
'    list_don5(i, 2) = Round((edessdo.Centon + edessdo.tron_ava.conduit.Diametre), 3)
    list_don5(i, 2) = Round(txtVersNum(Me.Lb_udev0(1).Caption), 3)
    list_don5(i, 3) = " m"
End Sub
Public Sub cre_list_don6()
Dim i As Integer
ReDim list_don6(2, 3)
i = 0
    list_don6(i, 1) = "Longueur de la chambre "
    list_don6(i, 2) = Round(txtVersNum(Me.Tb_dev(0).Text), 3)
    list_don6(i, 3) = " m"
i = i + 1
    list_don6(i, 1) = "Hauteur de la chambre"
    list_don6(i, 2) = Round(txtVersNum(Me.Tb_dev(1).Text), 3)
    list_don6(i, 3) = " m"
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

Private Sub m_quitter_Click()
    Unload owner
End Sub


Private Sub Cmd_Leaping_Wear_Click()
Dim ok As Boolean
If Me.Cmd_Leaping_Wear.Caption = "Calcul initialisation" Then
    ok = Calcul_initialisation
    If ok Then
'    Me.Cmd_mini.Enabled = True
'    Me.Cmd_recalc.Enabled = True
    Me.Frame4.Visible = True
    Me.Cmd_Leaping_Wear.Caption = "Calcul"
    Me.Cmd_Leaping_Wear.Visible = False
    ok_longueur = False
    ok_hauteur = False
    ok_largeur = False
    End If
ElseIf Me.Cmd_Leaping_Wear.Caption = "Calcul ouvrir" Then
'   ok = Calcul_Ouverture_Definie
        Dim sValLongueur As String
        ok = True
         Me.Frame4.Visible = True
        sValLongueur = Me.Tb_dev(0)
       Call calcul_avec_largeur
       Me.Tb_dev(0) = sValLongueur
       Me.UpDown2.Value = val(sValLongueur) * 1000
        Call calcul_avec_longueur
        Me.Cmd_Leaping_Wear.Caption = "Calcul"
        Me.Cmd_Leaping_Wear.Visible = False
        ok_longueur = False
        ok_largeur = False
        bKP = False
Else
    ok = True
   If ok_longueur Then
        Call calcul_avec_longueur
        ok_longueur = False
        bKP = False

    End If
   If ok_hauteur Then
        Call calcul_avec_hauteur
        ok_hauteur = False
        bKP = False
   End If
   If ok_largeur Then
        Call calcul_avec_largeur
        ok_largeur = False
        bKP = False
   End If
End If
If ok Then
    edessdo.tron_dech.conduit.typ = 2
    edessdo.tron_dech.Absamo = edessdo.tron_ava.Absamo
    edessdo.tron_dech.radamo = edessdo.tron_ava.radamo + edessdo.tron_ava.conduit.Diametre + edo.tav
    edessdo.tron_dech.Absava = edessdo.tron_dech.Absamo + edessdo.tron_dech.conduit.Longueur
    edessdo.tron_dech.radava = edessdo.tron_dech.radamo - edessdo.tron_dech.conduit.Longueur * edessdo.tron_dech.conduit.pente
'    Chk_max.Value = 1
'    Chk_cri.Value = 1
    Call OK_lignes_Click
'    ok_imp = True
'    Me.mnuprint.Enabled = True
'Else
'    ok_imp = False
 '   Me.mnuprint.Enabled = False

End If
End Sub
Private Function Calcul_initialisation(Optional ByVal codeCalc As String) As Boolean
Dim ok As Boolean, ok_amont As Boolean
Dim mes As String
    Me.Txtb_resu.Text = ""
    Me.Txtb_deversoir.Text = ""

If edessdo.Centon = 0 Then
'  MsgBox "la conduite aval est  en charge à Qbav!", vbExclamation, "Verif Aval"
  MsgBox "La hauteur entre canalisations n'a pas été saisie", vbExclamation, "Calcul déversoir"
    Calcul_initialisation = False
Else
''''vérif si  canal amont écoulement torrentiel
ok_amont = verif_amont
    Calcul_initialisation = ok_amont
Me.Tb_dev(0).Text = "0"
If ok_amont Then
    If ecoulam = "TORREN." Then
'    mes = sres + Chr(13) + Chr$(10) + " Ecoulement a débit de pointe Torrentiel " + ajout_zero(Trim(Str(Round(Qav, 3)))) + "m3/s"
        If codeCalc = "LARG" Then
            edoor_res.l_largOuverture = val(Me.Tb_larg.Text)
        End If
        ok = calcul_LEAPING_WEAR(mes, codeCalc)
        If ok Then
        Dim X As Double
        X = val(Me.Tb_dev(0).Text)
            If val(Me.Tb_dev(0).Text) = 0 Then
'                Me.Cmd_mini.Enabled = False
'                Me.Cmd_recalc.Enabled = False
                bKP = False
                Me.Tb_larg.Text = rempl_virgule(Format(edoor_res.l_largOuverture, "###0.000"))
                Me.UpDown1.Value = edoor_res.l_largOuverture * 100
                Me.Tb_long.Text = rempl_virgule(Format(edoor_res.l_ouverture, "###0.000"))
                Me.Tb_dev(0).Text = rempl_virgule(Format(Round(edoor_res.l_chambre1, 2), "###0.000"))
                Me.UpDown2.Value = Round(edoor_res.l_chambre1, 2) * 1000
                Me.Tb_dev(1).Text = rempl_virgule(Format(edessdo.Centon + edessdo.tron_ava.conduit.Diametre, "###0.000"))
                Me.Lb_udev(0).Caption = Format(edoor_res.l_chambre1, "###0.000")
                Me.Lb_udev(1).Caption = Format(edessdo.Centon + edessdo.tron_ava.conduit.Diametre, "###0.000")
            End If
            Me.Lb_udev0(0).Caption = Format(edoor_res.l_chambre1, "###0.000")
            Me.Lb_udev0(1).Caption = Format(edessdo.Centon + edessdo.tron_ava.conduit.Diametre, "###0.000")
            edessdo.tron_dech.conduit.typ = 2
            edessdo.tron_dech.Absamo = edessdo.tron_ava.Absamo
            edessdo.tron_dech.radamo = edessdo.tron_ava.radamo + edessdo.tron_ava.conduit.Diametre + edessdo.Centon
            edessdo.tron_dech.Absava = edessdo.tron_dech.Absamo + edessdo.tron_dech.conduit.Longueur
            edessdo.tron_dech.radava = edessdo.tron_dech.radamo - edessdo.tron_dech.conduit.Longueur * edessdo.tron_dech.conduit.pente
            Chk_max.Value = 1
            Chk_cri.Value = 1
            
            Call OK_lignes_Click
            ok_imp = True
            Me.mnuprint.Enabled = True
        Else
            mes = mes + Chr(13) + Chr$(10) + " La conduite aval est  en charge à Qbav "
            ok_imp = False
            Me.mnuprint.Enabled = False
'modif 12102006
            Calcul_initialisation = False
        End If
    Else
        mes = Chr(13) + Chr$(10) + " Ecoulement amont à débit de pointe non Torrentiel "
        ok_imp = False
        Me.mnuprint.Enabled = False
'modif 12102006
        Calcul_initialisation = False
    End If
End If
    Me.Txtb_deversoir.Text = mes
    ' 20070903
         mes = "Longueur de la chambre = " + ajout_zero(Trim(str(Round(edo.Longueur, 3)))) + " m"
     mes = mes + Chr(13) + Chr(10) + "Hauteur de la chambre = " + ajout_zero(Trim(str(Round(edo.hauteur, 3)))) + " m"
    Me.Txtb_resu.Text = mes

End If

End Function
Private Sub calcul_avec_largeur()
Dim mes As String
Dim ok1 As Boolean
Dim hmin As Double
mes = ""
Calcul_initialisation ("LARG")
End Sub
Private Sub calcul_avec_hauteur()
Dim mes As String
Dim ok1 As Boolean
Dim hmin As Double
mes = ""
hmin = edo.hauteur - edessdo.tron_ava.conduit.Diametre
edo.tav = hmin
ok1 = calcul_longueur(hmin)
Me.Lb_udev(0).Caption = rempl_virgule(Format(Round(edoor_res.l_chambre1, 3), "###0.000"))
bKP = False
'Me.Tb_dev(0).Text = rempl_virgule(Format(edoor_res.l_chambre1, "###0.000"))
Dim X As Double
Dim s As String
s = Me.Lb_udev(0).Caption
s = Me.Lb_udev(1).Caption
X = val(Me.Lb_udev(1).Caption)
If edo.hauteur > val(Me.Lb_udev(1).Caption) Then
    ok1 = calcul_door(mes)
    Me.Txtb_resu.Text = mes
    Me.Tb_dev(0).Text = rempl_virgule(Format(Round(edoor_res.l_chambre1, 3), "###0.000"))
s = Me.Tb_dev(0).Text
'Else
End If
    edo.Longueur = val(Me.Tb_dev(0).Text)
    mes = ""
     mes = mes + Chr(13) + Chr(10) + "Longueur de la chambre = " + ajout_zero(Trim(str(Round(edo.Longueur, 3)))) + " m"
     mes = mes + Chr(13) + Chr(10) + "Hauteur de la chambre = " + ajout_zero(Trim(str(Round(edo.hauteur, 3)))) + " m"
    Me.Txtb_resu.Text = mes
'    edoor_res.l_chambre1 = l_chambre1
'    edoor_res.l_jetaval_h = l_jetaval_h
'    edoor_res.l_jetaval_b = l_jetaval_b
'    edo.hauteur = hmin + edessdo.tron_ava.conduit.Diametre
'    edo.Absamo = edessdo.tron_amo.Absava
'    edo.Longueur = l_chambre1
    edo.Absava = edo.Absamo + edo.Longueur

    edo.radava = (edessdo.tron_amo.radava - edo.Longueur * edessdo.tron_amo.conduit.pente) - edessdo.tron_ava.conduit.Diametre - hmin
    'edo.pente = edessdo.tron_ava.conduit.pente
    edo.radamo = edo.radava + edo.Longueur * edo.pente
    edessdo.tron_ava.Absamo = edo.Absava
    edessdo.tron_ava.radamo = edo.radava
    edessdo.tron_ava.Absava = edessdo.tron_ava.Absamo + edessdo.tron_ava.conduit.Longueur
    edessdo.tron_ava.radava = edessdo.tron_ava.radamo - edessdo.tron_ava.conduit.Longueur * edessdo.tron_ava.conduit.pente
'End If

ok1 = calcul_courbes
End Sub
Private Sub calcul_avec_longueur()
Dim mes As String
Dim ok1 As Boolean
mes = ""
'hmin = edo.hauteur - edessdo.tron_ava.conduit.Diametre
'edo.tav = hmin
ok1 = calcul_hauteur(edo.Longueur)
Me.Lb_udev(1).Caption = rempl_virgule(Format(Round(edo.tav + edessdo.tron_ava.conduit.Diametre, 3), "###0.000"))
bKP = False
'Me.Tb_dev(1).Text = rempl_virgule(Format(edo.tav + edessdo.tron_ava.conduit.Diametre, "###0.000"))

Dim s As String
s = Me.Lb_udev(0).Caption
If edo.Longueur < val(Me.Lb_udev(0).Caption) Then
    ok1 = calcul_door(mes)
    Me.Txtb_resu.Text = mes
    Me.Tb_dev(1).Text = rempl_virgule(Format(Round(edo.tav + edessdo.tron_ava.conduit.Diametre, 3), "###0.000"))
'Else
End If
mes = ""
Dim X As Double
X = val(Me.Lb_udev(0).Caption)
X = val(Me.Tb_dev(0).Text)
X = val(Me.Tb_dev(1).Text)
edo.tav = val(Me.Tb_dev(1).Text) - edessdo.tron_ava.conduit.Diametre
hmin = edo.tav
'    edoor_res.l_chambre1 = l_chambre1
'    edoor_res.l_jetaval_h = l_jetaval_h
'    edoor_res.l_jetaval_b = l_jetaval_b
     edo.hauteur = hmin + edessdo.tron_ava.conduit.Diametre
     mes = mes + "Longueur de la chambre = " + ajout_zero(Trim(str(Round(edo.Longueur, 3)))) + " m"
     mes = mes + Chr(13) + Chr(10) + "Hauteur de la chambre = " + ajout_zero(Trim(str(Round(edo.hauteur, 3)))) + " m"
    Me.Txtb_resu.Text = mes
'    edo.Absamo = edessdo.tron_amo.Absava
'    edo.Longueur = l_chambre1
    edo.Absava = edo.Absamo + edo.Longueur

 '   edo.radava = (edessdo.tron_amo.radava - edo.Longueur * edessdo.tron_amo.conduit.pente) - edessdo.tron_ava.conduit.Diametre - hmin
    edo.radava = (edessdo.tron_amo.radava - edo.Longueur * edo.pente) - edessdo.tron_ava.conduit.Diametre - hmin
'    edo.pente = edessdo.tron_ava.conduit.pente
    edo.radamo = edo.radava + edo.Longueur * edo.pente
    edessdo.tron_ava.Absamo = edo.Absava
    edessdo.tron_ava.radamo = edo.radava
    edessdo.tron_ava.Absava = edessdo.tron_ava.Absamo + edessdo.tron_ava.conduit.Longueur
    edessdo.tron_ava.radava = edessdo.tron_ava.radamo - edessdo.tron_ava.conduit.Longueur * edessdo.tron_ava.conduit.pente
'End If
ok1 = calcul_courbes

End Sub

Private Sub Cmd_mini_Click()
Dim ok1 As Boolean
Dim mes As String
mes = ""
ok1 = calcul_mini(mes, edessdo.Centon)
edo.tav = edessdo.Centon
            edessdo.tron_dech.conduit.typ = 2
            edessdo.tron_dech.Absamo = edessdo.tron_ava.Absamo
            edessdo.tron_dech.radamo = edessdo.tron_ava.radamo + edessdo.tron_ava.conduit.Diametre + edo.tav
            edessdo.tron_dech.Absava = edessdo.tron_dech.Absamo + edessdo.tron_dech.conduit.Longueur
            edessdo.tron_dech.radava = edessdo.tron_dech.radamo - edessdo.tron_dech.conduit.Longueur * edessdo.tron_dech.conduit.pente
'            edessdo.tron_dech.conduit.Diametre = edessdo.tron_amo.conduit.Diametre
'            edessdo.tron_dech.conduit.pente = edessdo.tron_amo.conduit.pente
'            edessdo.tron_dech.conduit.rugosite = edessdo.tron_amo.conduit.rugosite
            Chk_max.Value = 1
            Chk_cri.Value = 1
            Call OK_lignes_Click

End Sub

Private Sub Cmd_recalc_Click()
    edo.hauteur = txtVersNum(Me.Tb_dev(1).Text)
    Call calcul_avec_hauteur
            edessdo.tron_dech.conduit.typ = 2
            edessdo.tron_dech.Absamo = edessdo.tron_ava.Absamo
            edessdo.tron_dech.radamo = edessdo.tron_ava.radamo + edessdo.tron_ava.conduit.Diametre + edo.tav
            edessdo.tron_dech.Absava = edessdo.tron_dech.Absamo + edessdo.tron_dech.conduit.Longueur
            edessdo.tron_dech.radava = edessdo.tron_dech.radamo - edessdo.tron_dech.conduit.Longueur * edessdo.tron_dech.conduit.pente
'            edessdo.tron_dech.conduit.Diametre = edessdo.tron_amo.conduit.Diametre
'            edessdo.tron_dech.conduit.pente = edessdo.tron_amo.conduit.pente
'            edessdo.tron_dech.conduit.rugosite = edessdo.tron_amo.conduit.rugosite
            Chk_max.Value = 1
            Chk_cri.Value = 1
            Call OK_lignes_Click
End Sub

Private Sub Form_Activate()
    change_coul = False
'    owner.affich_aide Me.Name, mes_prec
End Sub

Private Sub Form_Click()
    owner.affich_aide Me.Name, ""  'Déversoir d'orage"
    Change_Couleur "Me", 0
DoEvents
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

Private Sub Frm_contraintes_Click()
Dim mes As String
Dim nom As String
nom = "Frm_contraintes"
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


Private Sub Frm_Ouverture_Click()
Dim mes As String
Dim nom As String
nom = "Frm_Ouverture"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

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

Private Sub Lb_intHmin_Click()
Dim mes As String
Dim nom As String
nom = "Lb_intHmin"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_intLarg_Click()
Dim mes As String
Dim nom As String
nom = "Lb_intLarg"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes
End Sub

Private Sub Lb_intLong_Click()
Dim mes As String
Dim nom As String
nom = "Lb_intLong"
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
FrmPrint.Type1 = "deversoiror"
FrmPrint.nomobjet = Tb_titre.Text
FrmPrint.titre1 = "FICHE HYDRAULIQUE DEVERSOIR à ouverture de radier"
FrmPrint.sstitre1 = "Caractéristiques " + Frm_bv.Caption
Frm_imp.Type1 = "deversoiror"
Frm_imp.nomobjet = Tb_titre.Text
Frm_imp.titre1 = "FICHE HYDRAULIQUE DEVERSOIR à ouverture de radier"
Frm_imp.sstitre1 = "Caractéristiques " + Frm_bv.Caption
'FrmPrint.ssTitre2 = "Contraintes"
'Frm_imp.ssTitre2 = "Contraintes"
FrmPrint.ssTitre3 = "Conduites"
Frm_imp.ssTitre3 = "Conduites"
FrmPrint.ssTitre4 = "Résultats de fonctionnement"
Frm_imp.ssTitre4 = "Résultats de fonctionnement"
FrmPrint.ssTitre5 = "Dimensions minimum"
Frm_imp.ssTitre5 = "Dimensions minimum"
FrmPrint.ssTitre6 = "Dimensions retenues"
Frm_imp.ssTitre6 = "Dimensions retenues"
cre_list_don1
'cre_list_don2
cre_list_don3
cre_list_don4
cre_list_don5
cre_list_don6
'Call dess_do_print(Frm_desprint.UC_graphique1, False, False, True) 'okcharge, okpiezo, okeau
'Call dess_dech_print(Frm_desprint.UC_graphique2)
Call Cmd_recalc_Click
Set pict1 = Frm_desprint.UC_graphique1.lire_pict1()
FrmPrint.paint_picture pict1
SavePicture pict1, chemin_app + "dess.bmp"
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
    edessdo.nom = nom_ouvrage 'ebv.nom
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
    edo.Longueur = txtVersNum(Me.Tb_dev(0).Text)
    edo.hauteur = txtVersNum(Me.Tb_dev(1).Text)
    If isave > 0 Then
        If bsous Then
           reponse = MsgBox("Le nom est déjà utilisé. Le remplacer?", 4, "Sauvegarde d'un déversoir")
           Else
           reponse = 6
        End If
        If reponse = 6 Then
            esave.type = "deversoiror"
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
        esave.type = "deversoiror"
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

Private Function calc_amont() As Boolean
Dim Qts As Double, Qrin As Double, Qpluie As Double
Dim qv As deb_vit
Dim ltc() As Variant, ct() As Variant, qvm(5) As Variant
Dim qcal As Double
Dim betam As Double
Dim message As String

Dim cana_amo As conduite
calc_amont = False
cana_amo = edessdo.tron_amo.conduit
'wagner Call calcul_condam(cana_amo)
' a voir Qref et Qrin
Qts = edessdo.Qts
Qrin = edessdo.Qrin '+ edessdo.Qts
Qpluie = edessdo.Qpluie

qv = debvit_ps(cana_amo)
'Me.Lb_vpsm.Caption = "Vitesse pleine section = " + ajout_zero(Trim(Str(Round(qv.vitesse, 3)))) + " m/s"
'Me.Lb_dpsm.Caption = "Débit pleine section = " + ajout_zero(Trim(Str(Round(qv.debit, 3)))) + " m3/s"
'resudev.vpsm = ajout_zero(Trim(Str(Round(qv.vitesse, 3))))
'resudev.dpsm = ajout_zero(Trim(Str(Round(qv.debit, 3))))
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
'    Me.Lb_vqtsm.Caption = "Vitesse d'écoulement à QTS = " + "   " + " m/s"
'    Me.Lb_hqtsm.Caption = "Hauteur d'eau QTS = " + "   " + " m"
'    Me.Lb_vqrinm.Caption = "Vitesse d'écoulement à QREF = " + "   " + " m/s"
'    Me.Lb_hqrinm.Caption = "Hauteur d'eau QREF = " + "   " + " m"
'    Me.Lb_vqpluiem.Caption = "Vitesse d'écoulement à QORA = " + "   " + " m/s"
'    Me.Lb_hqpluiem.Caption = "Hauteur d'eau QORA = " + "   " + " m"
  calc_amont = False

Else
    qcal = Qts
    Call cana(cana_amo, ct)
    ltc = calc_par(cana_amo)
    qvi = caltran1(qcal, ct, ltc)
'    Me.Lb_vqtsm.Caption = "Vitesse d'écoulement à QTS = " + ajout_zero(Trim(Str(Round(qvi(2), 3)))) + " m/s"
'    Me.Lb_hqtsm.Caption = "Hauteur d'eau QTS = " + ajout_zero(Trim(Str(Round(qvi(5), 3)))) + " m"
'    resudev.vqtsm = ajout_zero(Trim(Str(Round(qvi(2), 3))))
'    resudev.hqtsm = ajout_zero(Trim(Str(Round(qvi(5), 3))))
'    Debug.Print "qvi : 1= débit"; qvi(1); " 2=vitesse"; qvi(2); "qvi"; " 3=acceleration"; qvi(3)
'    Debug.Print " 4=Largeur libre"; qvi(4); " 5 = hauteur; "; qvi(5); ""
'                vitmax = qvm(2)
'                qvm = caltran1(qps / 10#, ct, ltc)
'                vit10 = qvm(2)
'                qvm = caltran1(qps / 100#, ct, ltc)
'                vit100 = qvm(2)
    qcal = Qrin
    Call cana(cana_amo, ct)
    ltc = calc_par(cana_amo)
    qvi = caltran1(qcal, ct, ltc)
         Dim nFroude   As Double, vit As Double, hr As Double
        vit = qvi(2)
        hr = (qvi(1) / vit) / qvi(4)
        nFroude = calcul_Froude1(vit, hr)
        hr = qvi(5)
        nFroude = calcul_Froude(qcal / 1000, hr, cana_amo.Diametre)
'        Debug.Print nFroude
'    Me.Lb_vqrinm.Caption = "Vitesse d'écoulement à QREF = " + ajout_zero(Trim(Str(Round(qvi(2), 3)))) + " m/s"
'    Me.Lb_hqrinm.Caption = "Hauteur d'eau QREF = " + ajout_zero(Trim(Str(Round(qvi(5), 3)))) + " m"
'    resudev.vqrinm = ajout_zero(Trim(Str(Round(qvi(2), 3))))
'    resudev.hqrinm = ajout_zero(Trim(Str(Round(qvi(5), 3))))
        message = ""
        If nFroude < 1 Then
    '        MsgBox "Ecoulement a débit de pointe Torrentiel !" + Chr(13) + "Diminnuez la pente ou prevoir un ressaut ", vbOKOnly, "Vérification d'écoulement"
            message = message + " Régime d'écoulement fluvial ! " + Chr(13) + " Ce type de déversoir n'est pas adapté "
            MsgBox message, vbOKOnly, "Condition d'écoulement amont"
            calc_amont = False
            ecoulam = "FLUVIAL"
        ElseIf nFroude < 1.5 Then
            message = message + " Régime d'écoulement torrentiel " + Chr(13) + " Le nombre de Foude est inférieur à 1.5 "
            message = message + Chr(13) + " valeur minimale recommandée. "
            MsgBox message, vbOKOnly, "Condition d'écoulement amont"
            calc_amont = True
            ecoulam = "TORREN."
        Else
            calc_amont = True
            ecoulam = "TORREN."
        End If
  
        



'    qcal = Qpluie
'    Call cana(cana_amo, ct)
'    ltc = calc_par(cana_amo)
'    qvi = caltran1(qcal, ct, ltc)


'    Me.Lb_vqpluiem.Caption = "Vitesse d'écoulement à QORA = " + ajout_zero(Trim(Str(Round(qvi(2), 3)))) + " m/s"
'    Me.Lb_hqpluiem.Caption = "Hauteur d'eau QORA = " + ajout_zero(Trim(Str(Round(qvi(5), 3)))) + " m"
'    resudev.vqpluiem = ajout_zero(Trim(Str(Round(qvi(2), 3))))
'    resudev.hqpluiem = ajout_zero(Trim(Str(Round(qvi(5), 3))))
' calcul du regime amont
 ' a revoir
'        betam = angle(Qpluie / (qv.debit * 1000))
'        betam = beta
'        ecoulam = calcul_ecoul(Qpluie / 1000, cana_amo.Diametre, betam)
'        vit = qvi(2)
'        hr = (qvi(1) / vit) / qvi(4)
'        nFroude = calcul_Froude(vit, hr)
'        Debug.Print nFroude
'        If ecoulam <> "TORREN." Then
'    '        MsgBox "Ecoulement a débit de pointe Torrentiel !" + Chr(13) + "Diminnuez la pente ou prevoir un ressaut ", vbOKOnly, "Vérification d'écoulement"
'            message = message + "> Froude dépasse"
'            MsgBox message, vbOKOnly, "Condition d'écoulement amont"
'        calc_amont = False
'        Else
'            calc_amont = True
'        End If
' ''''' Me.SSTab1.TabEnabled(3) = True
 '           calc_amont = True

End If

End Function

Private Function verif_amont() As Boolean
Dim troamo As troncon, troava As troncon
Dim cana_amo As conduite
Dim cana_ava As conduite

    verif_amont = False
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
''si prise en compte contrainte cote radier amont
'      .radamo = edessdo.rdoam
'      .radava = edessdo.rdoam - cana_amo.Longueur * cana_amo.pente
''si prise en compte contrainte cote radier aval
      .radava = edessdo.rdoav
      .radamo = edessdo.rdoav + cana_amo.Longueur * cana_amo.pente
    End With
    edessdo.tron_amo = troamo


' calcul hydraulique
    verif_amont = calc_amont
    

'If troamo.radava <= edessdo.rdoav Then
'    MsgBox "cote aval canalisation amont inférieure à cote radier obligé aval", vbOKOnly
'End If
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
    door_bv = True
    Set owner.fbassin = New Frm_bv2
    owner.fbassin.Show
    owner.fbassin.nom_ouvrage = nombassin
    owner.fbassin.Cmd_retour.Visible = True
    owner.fbassin.Cmd_retour.Caption = "Retour au déversoir"
    fich_lect = nom_fich
    Call owner.fbassin.rec_bassin_versant
    owner.affich_aide owner.fbassin.Name, "Module" ' "Calcul de débit de bassin versant"
End Sub

Private Sub Form_Load()
    okg = True
    Me.KeyPreview = True
    Call ini_tooltip_door(Me)
    nom_ouvrage = ""
    ouv_sauve = False
    save_fich = False
    nom_dessin = chemin_app + "do_bassin.bmp"
'    nom_fich = chemin_app + "deversoir.bin"
'    nom_fich = chemin_app + "etude.bin"
    nom_type = "deversoiror"
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
    ok_longueur = False
    ok_hauteur = False
    Me.Tb_debit(1).Visible = False
    Me.Lb_intdebit(1).Visible = False
    Me.Lb_udebit(1).Visible = False
'    Me.SSTab1.TabEnabled(1) = False
    Me.SSTab1.TabVisible(1) = False
    Call ini_bv
    Call ini_resuintdv
    Me.SSTab1.Tab = 0
    nombassin = ""
    Call ini_edessdo
    Cmd_Leaping_Wear.Enabled = False
    Call reini_form(0)
    Call init_graphique
    Call ini_form
    owner.fdessin.UC_graphique1.Visible = True
    owner.fdessin.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.Height = 6000
    Me.Tb_debit(0) = "0.0"
    Me.Tb_debit(2) = "0.0"
    Me.Tb_debit(1) = "10.0"
    edessdo.Qts = txtVersNum(Me.Tb_debit(1).Text)
    ouv_sauve = False
    save_fich = False
    fich_lect = ""
    change_coul = False
End Sub
Private Sub init_graphique()
    owner.fdessin.mnu_fichier.Caption = Me.mnufichier.Caption
    owner.fdessin.Image1.Visible = False
    owner.fdessin.Image2.Visible = False
    owner.fdessin.Image3.Visible = False
    owner.fdessin.UC_graphiqueB.Visible = False
    owner.fdessin.UC_graphique1.Visible = True
    owner.fdessin.UC_graphique2.Visible = False
'    owner.fdessin.UC_graphiqueB.graphique_clear
 '   owner.fdessin.UC_graphiqueB.init_fond nom_dessin
    owner.fdessin.Image3.Picture = LoadPicture(nom_dessin)
    owner.fdessin.UC_graphique1.graphique_clear
    owner.fdessin.UC_graphique1.reinit 7, "Arial"
    owner.fdessin.UC_graphique1.init_title
    owner.fdessin.UC_graphique1.init_titleh ""
    owner.fdessin.UC_graphique1.init_titleb ""
    owner.fdessin.UC_graphique2.reinit 7, "Arial"
    owner.fdessin.UC_graphique2.graphique_clear
    owner.fdessin.UC_graphique2.init_title
    owner.fdessin.UC_graphique2.init_titleh ""
    owner.fdessin.UC_graphique2.init_titleb ""
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
owner.fdessin.UC_graphique1.graphique_clear
'owner.fdessin.UC_graphique2.graphique_clear
Txtb_deversoir.Text = ""
Txtb_resu.Text = ""
'Me.Cmd_mini.Enabled = False
'Me.Cmd_recalc.Enabled = False
Me.Frame4.Visible = False
bKP = False
Me.Tb_dev(0).Text = rempl_virgule(Format(edo.Longueur, "##0.000"))
Me.Tb_dev(1).Text = rempl_virgule(Format(edo.hauteur, "###0.000"))
Me.Lb_udev(0).Caption = ""
Me.Lb_udev(1).Caption = ""
Me.Lb_udev0(0).Caption = ""
Me.Lb_udev0(1).Caption = ""
Me.Cmd_Leaping_Wear.Caption = "Calcul initialisation"
Me.Cmd_Leaping_Wear.Visible = True
Select Case ntab
    Case Is = 5
'        Call init_graphique

        Me.Tb_dev(0).Text = "0.000"
        Me.Tb_dev(1).Text = "0.000"
        Me.Tb_long.Text = "0.00"
        Me.Tb_larg.Text = "0.00"
        If edessdo.Qpluie > 0 And edessdo.Qts > 0 And edessdo.Qrin > 0 Then
            ouv_sauve = True
        End If
   Case Is = 1
        Me.Tb_dev(0).Text = "0.000"
        Me.Tb_dev(1).Text = "0.000"
        Me.Tb_long.Text = "0.00"
        Me.Tb_larg.Text = "0.00"

    Case Is = 2
        Me.Tb_dev(0).Text = "0.000"
        Me.Tb_dev(1).Text = "0.000"
        Me.Tb_long.Text = "0.00"
        Me.Tb_larg.Text = "0.00"

        If edessdo.dam > 0 And edessdo.iRadam > 0 And edessdo.Kam > 0 _
            And edessdo.Lam > 0 Then
            Me.Cmd_amo.Enabled = True
        Else
            Me.Cmd_amo.Enabled = False
       End If
    Case Is = 3
        Me.Tb_dev(0).Text = "0.000"
        Me.Tb_dev(1).Text = "0.000"
        Me.Tb_long.Text = "0.00"
        Me.Tb_larg.Text = "0.00"
        If edessdo.dav > 0 And edessdo.iradav > 0 And edessdo.kav > 0 _
            And edessdo.Lav > 0 Then
             Me.Cmd_ava.Enabled = True
        Else
            Me.Cmd_ava.Enabled = False
       End If
   Case Is = 4
        Me.Tb_dev(0).Text = "0.000"
        Me.Tb_dev(1).Text = "0.000"
        Me.Tb_long.Text = "0.00"
        Me.Tb_larg.Text = "0.00"
        If val(Tb_dech(0).Text) > 0 And val(Tb_dech(1).Text) > 0 And val(Tb_dech(2).Text) > 0 _
            And val(Tb_dech(3).Text) > 0 Then
             Me.Cmd_dech.Enabled = True
        Else
            Me.Cmd_dech.Enabled = False
       End If

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
'    Me.Tb_cont(2).Text = rempl_virgule(Format(edessdo.lgdisp, "###0.0"))
'    Me.Tb_cont(3).Text = rempl_virgule(Format(edessdo.phex, "###0.0"))
'    Me.Tb_cont(4).Text = rempl_virgule(Format(edessdo.rdoex, "###0.0"))
'    Me.Tb_cont(5).Text = rempl_virgule(Format(edessdo.lgca, "###0.0"))
    Me.Tb_hmin.Text = rempl_virgule(Format(edessdo.Centon, "###0.000"))
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
    Me.Tb_long.Text = rempl_virgule(Format(edessdo.lgdisp, "###0.000"))
    Me.Tb_larg.Text = rempl_virgule(Format(edessdo.lgca, "###0.000"))
    Me.UpDown1.Value = edessdo.lgca * 100
'    Me.Tb_dev(2).Text = rempl_virgule(Format(edo.pente, "##0.0000"))
'    Me.Tb_dev(3).Text = rempl_virgule(Format(edessdo.Tram, "##0.00"))
    Me.Tb_dech(0).Text = rempl_virgule(Format(edessdo.tron_dech.conduit.Diametre * 1000, "###0"))
    Me.Tb_dech(1).Text = rempl_virgule(Format(edessdo.tron_dech.conduit.pente * 10000, "###0"))
    Me.Tb_dech(3).Text = rempl_virgule(Format(edessdo.tron_dech.conduit.Longueur, "###0.00"))
    Me.Tb_dech(2).Text = rempl_virgule(Format(edessdo.tron_dech.conduit.rugosite, "###0"))
    Me.Frm_bv.Caption = "Hydraulique du B.V : " + Trim(nombassin) ' Trim(ebv.nom)

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
Dim xlong As Double, hmin As Double
Dim xprof As Double
Call ini_form
reini_form 0
'reini_form 1
'If Me.SSTab1.TabEnabled(2) Then
Call ini_Cmd_Leaping_Wear
If Cmd_Leaping_Wear.Enabled = True Then
    If val(Me.Tb_dev(0).Text) > 0 And val(Me.Tb_dev(1).Text) > 0 Then
        Me.Cmd_Leaping_Wear.Caption = "Calcul ouvrir"
        Me.Cmd_Leaping_Wear.Visible = False
        ok_longueur = True
        ok_largeur = False
    End If
    Call Cmd_Leaping_Wear_Click
'    Call OK_lignes_Click
    xlong = txtVersNum(Me.Tb_dev(0).Text)
    xprof = txtVersNum(Me.Tb_dev(1).Text)
    If xlong > 0 And xprof > 0 And xlong <> edo.Longueur And xprof <> edo.hauteur Then
        edo.hauteur = txtVersNum(Me.Tb_dev(1).Text)
        hmin = edo.hauteur - edessdo.tron_ava.conduit.Diametre
        edo.tav = hmin
        ok = calcul_longueur(hmin)
        Me.Lb_udev(0).Caption = Format(Round(edoor_res.l_chambre1, 2), "###0.000")
        edo.Longueur = txtVersNum(Me.Tb_dev(0).Text)
        ok = calcul_hauteur(edo.Longueur)
        Me.Lb_udev(1).Caption = Format(edo.tav + edessdo.tron_ava.conduit.Diametre, "###0.000")
        ok_hauteur = True
        ok_longueur = False

        Call Cmd_Leaping_Wear_Click

'    If edo.Longueur > 0 And edo.hauteur > 0 Then
    End If
   Me.SSTab1.Tab = 3
End If
'End If
End Sub
Private Sub ini_canamo()
edessdo.tron_amo.conduit.typ = 2
edessdo.tron_amo.conduit.Diametre = 0
edessdo.tron_amo.conduit.Longueur = 0#
edessdo.tron_amo.conduit.pente = 0
edessdo.tron_amo.conduit.rugosite = 0
End Sub
Private Sub ini_canava()
edessdo.tron_ava.conduit.typ = 2
edessdo.tron_ava.conduit.Diametre = 0
edessdo.tron_ava.conduit.Longueur = 0#
edessdo.tron_ava.conduit.pente = 0
edessdo.tron_ava.conduit.rugosite = 0
End Sub
Private Sub ini_canadech()
edessdo.tron_dech.conduit.typ = 2
edessdo.tron_dech.conduit.Diametre = 0
edessdo.tron_dech.conduit.Longueur = 0#
edessdo.tron_dech.conduit.pente = 0
edessdo.tron_dech.conduit.rugosite = 0
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
        Me.Frm_bv.Caption = "Hydraulique du B.V : "
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
        Me.Lab_bas.Caption = ""
        Me.SSTab1.TabEnabled(1) = False
   End If
End Sub
Private Sub reini_valeurs(ntab As Integer)
Select Case ntab
    Case Is = 0
        Txtb_deversoir.Text = ""
        Txtb_decharge.Text = ""
    Case Is = 1
        Txtb_deversoir.Text = ""
        Txtb_decharge.Text = ""
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
'    Me.SSTab1.Tab = 0
'    Me.SSTab1.TabEnabled(1) = False
'    Me.SSTab1.TabEnabled(2) = False
'    Me.SSTab1.TabEnabled(3) = False
'    Me.SSTab1.TabEnabled(4) = False
'    Me.SSTab1.TabEnabled(5) = False
    Me.Tb_titre.Text = ""
    Me.Caption = fen_titre
'    Me.Cmd_del.Visible = False
    nombassin = ""
    Call ini_edessdo
    Call ini_resuintdv
    Call ini_form
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
'        Me.SSTab1.TabEnabled(1) = False
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
Dim okmax As Boolean, okcri As Boolean
okmax = True
okcri = True
If Chk_max.Value = 0 Then
    okmax = False
End If
If Chk_cri.Value = 0 Then
    okcri = False
End If
' a enlever
'okmax = False
Call dessin_door(okmax, okcri)

End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
Dim mes As String
'Me.Cmd_Leaping_Wear.Enabled = True
Select Case SSTab1.Tab
    Case Is = 0
        mes = IDhlp_DOORDonneesBase '"Hydraulique du bassin versant"
'        owner.fdessin.Image1.Visible = False
'        owner.fdessin.UC_graphiqueB.Visible = False
'        owner.fdessin.Image3.Visible = True
'        owner.fdessin.UC_graphique1.Visible = True
'        owner.fdessin.UC_graphique2.Visible = False
   Case Is = 1
        mes = IDhlp_DOORContraintes
'        owner.fdessin.Image1.Visible = False
'        owner.fdessin.UC_graphiqueB.Visible = False
'        owner.fdessin.Image3.Visible = True
'        owner.fdessin.UC_graphique1.Visible = True
'        owner.fdessin.UC_graphique2.Visible = False
   Case Is = 2
        mes = IDhlp_DOORConduiteArrivee '"Conduite d'arrivée"
'        owner.fdessin.Image1.Visible = False
'        owner.fdessin.UC_graphique1.Visible = True
'        owner.fdessin.UC_graphiqueB.Visible = False
'        owner.fdessin.Image3.Visible = False
'        owner.fdessin.UC_graphique2.Visible = False
   Case Is = 3
   
        mes = IDhlp_DOOROuvrageDeversoir '"L'ouvrage déversoir"
'        owner.fdessin.Image1.Visible = False
'        owner.fdessin.UC_graphiqueB.Visible = False
'        owner.fdessin.UC_graphique1.Visible = True
'        owner.fdessin.UC_graphique2.Visible = False
'   Case Is = 4
'        mes = "Chambre de déversement"
'        owner.fdessin.Image1.Visible = False
'        owner.fdessin.UC_graphiqueB.Visible = False
'        owner.fdessin.Image3.Visible = False
'        owner.fdessin.UC_graphique1.Visible = True
'        owner.fdessin.UC_graphique2.Visible = False
End Select
If owner.fcom.Name = "Frm_ss_commentaire" Then
    Change_Couleur "SSTab1", 0
  '  owner.affich_aide Me.Name, mes
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
 '   SSTab1.Tab = 0
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

    Select Case Index
        Case Is = 0
            edessdo.tron_dech.conduit.Diametre = txtVersNum(Tb_dech(0).Text) / 1000#
        Case Is = 1
            edessdo.tron_dech.conduit.pente = txtVersNum(Tb_dech(1).Text) / 10000
        Case Is = 2
            edessdo.tron_dech.conduit.rugosite = txtVersNum(Tb_dech(2).Text)
        Case Is = 3
            edessdo.tron_dech.conduit.Longueur = txtVersNum(Tb_dech(3).Text)
    End Select
Call ini_Cmd_Leaping_Wear
    Call reini_form(4)
'    Me.Cmd_resudech.Enabled = True
'    Me.Txtb_decharge.Text = ""
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
owner.affich_aide Me.Name, mes
Call sel_text(Tb_dech(Index))
End Sub

Private Sub Tb_dech_GotFocus(Index As Integer)
Dim mes As String
Dim nom As String
nom = "Tb_dech"
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
                nom = verif_cart0(Tb_dev(Index).Text, "Saisie profondeur DO", "R")
        End Select
  If nom = "" Then
    Tb_dev(Index).Text = sval_champ
    Tb_dev(Index).SelStart = iSels
    Tb_dev(Index).SelLength = iSell
  Else
'  End If
'End If
'****

Select Case Index
    Case Is = 0
        edo.Longueur = txtVersNum(Me.Tb_dev(Index).Text)
        ok_longueur = True
'        Call calcul_avec_longueur

    Case Is = 1
        edo.hauteur = txtVersNum(Me.Tb_dev(Index).Text)
        ok_hauteur = True
'        Call calcul_avec_hauteur
    End Select

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
If SSTab1.Tab <> 3 Then
    If Me.SSTab1.TabEnabled(3) Then
        SSTab1.Tab = 3
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
Private Sub ini_Cmd_Leaping_Wear()
Cmd_Leaping_Wear.Enabled = True
If edessdo.Qpluie = 0# Or edessdo.Qts = 0# Or edessdo.Qrin = 0# Or edo.tav = 0# Or _
edessdo.dam = 0# Or edessdo.iRadam = 0# Or edessdo.Kam = 0# Or edessdo.Lam = 0# Or _
edessdo.dav = 0# Or edessdo.iradav = 0# Or edessdo.kav = 0# Or edessdo.Lav = 0# Or _
edessdo.tron_dech.conduit.Diametre = 0# Or edessdo.tron_dech.conduit.pente = 0# _
Or edessdo.tron_dech.conduit.rugosite = 0# Or edessdo.tron_dech.conduit.Longueur = 0# Then
    Cmd_Leaping_Wear.Enabled = False
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
            edessdo.Qts = txtVersNum("20.0")
        Case Is = 1
            edessdo.Qts = txtVersNum(Me.Tb_debit(1).Text)
        Case Is = 2
            edessdo.Qrin = txtVersNum(Me.Tb_debit(2).Text)
    End Select
'    nombassin = ""
'    Me.Lab_bas.Caption = ""
If bKP Then
    Me.Lab_bas.Caption = ""
End If
Call ini_Cmd_Leaping_Wear
Call reini_form(5)
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
End Select
Call ini_Cmd_Leaping_Wear
Call reini_form(1)
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
 Call ini_Cmd_Leaping_Wear
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
Call ini_Cmd_Leaping_Wear
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
    owner.affich_aide Me.Name, " "  'Déversoir d'orage"
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

Private Sub Tb_larg_Change()
Dim nform As Integer
Dim nom As String

If bKP Then
                nom = verif_cart0(Tb_larg.Text, "Saisie largeur ouverture", "R")
  If nom = "" Then
    Tb_larg.Text = sval_champ
    Tb_larg.SelStart = iSels
    Tb_larg.SelLength = iSell
  End If
End If
'****

If Trim(Tb_larg.Text) = "" Then
    Tb_larg.Text = 0#
End If
    edessdo.lgca = txtVersNum(Me.Tb_larg.Text)
    ok_largeur = True
'    Call ini_Cmd_Leaping_Wear
'    Call reini_form(1)
    sval_champ = ""
    bKP = False



End Sub

Private Sub Tb_larg_Click()
Dim mes As String
Dim nom As String
nom = "Tb_larg"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
'Call sel_text(Tb_larg(Index))
Call sel_text(Tb_larg)
End Sub

Private Sub Tb_larg_GotFocus()
Dim mes As String
Dim nom As String
nom = "Tb_larg"
If SSTab1.Tab <> 3 Then
    If Me.SSTab1.TabEnabled(3) Then
        SSTab1.Tab = 3
    Else
        Me.Tb_debit(0).SetFocus
    End If
'    Me.Tb_Debit(0).SetFocus
End If
Call sel_text(Tb_larg)
If change_coul Then
    Change_Couleur nom, Index
    mes = Rec_Mes(nom, Index)
    owner.affich_aide Me.Name, mes
End If

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
Dim nform As Integer
Dim nom As String

If bKP Then
                nom = verif_cart0(Tb_long.Text, "Saisie longueur ouverture", "R")
  If nom = "" Then
    Tb_long.Text = sval_champ
    Tb_long.SelStart = iSels
    Tb_long.SelLength = iSell
  End If
End If
'****

If Trim(Tb_long.Text) = "" Then
    Tb_long.Text = 0#
End If
    edessdo.lgdisp = txtVersNum(Me.Tb_long.Text)
'    Call ini_Cmd_Leaping_Wear
'    Call reini_form(1)
    sval_champ = ""
    bKP = False




End Sub

Private Sub Tb_hmin_Change()
Dim nform As Integer
Dim nom As String

If bKP Then
                nom = verif_cart0(Tb_hmin.Text, "Saisie hauteur entre canalisations", "R")
  If nom = "" Then
    Tb_hmin.Text = sval_champ
    Tb_hmin.SelStart = iSels
    Tb_hmin.SelLength = iSell
  End If
End If
'****

If Trim(Tb_hmin.Text) = "" Then
    Tb_hmin.Text = 0#
End If
'        edoor_res.hmin = txtVersNum(Me.Tb_hmin.Text)
    edessdo.Centon = txtVersNum(Me.Tb_hmin.Text)
    edo.tav = txtVersNum(Me.Tb_hmin.Text)
    Call ini_Cmd_Leaping_Wear
    Call reini_form(1)
    sval_champ = ""
    bKP = False


End Sub

Private Sub Tb_hmin_Click()
Dim mes As String
Dim nom As String
nom = "Tb_hmin"
mes = Rec_Mes(nom, 0)
Change_Couleur nom, 0
owner.affich_aide Me.Name, mes
Call sel_text(Tb_hmin)
'''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub

Private Sub Tb_hmin_GotFocus()
Dim mes As String
Dim nom As String
nom = "Tb_hmin"
Call sel_text(Tb_hmin)
If change_coul Then
    Change_Couleur nom, 0
    mes = Rec_Mes(nom, 0)
    owner.affich_aide Me.Name, mes
End If
''owner.affich_aide Me.Name, "pompe Conduite Amont"

End Sub

Private Sub Tb_hmin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    bKP = True
    sval_champ = Tb_hmin.Text
    iSels = Tb_hmin.SelStart
    iSell = Tb_hmin.SelLength
End If

End Sub

Private Sub Tb_long_Click()
Dim mes As String
Dim nom As String
nom = "Tb_long"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
owner.affich_aide Me.Name, mes
Call sel_text(Tb_long)

End Sub

Private Sub Tb_long_GotFocus()
Dim mes As String
Dim nom As String
nom = "Tb_long"
If SSTab1.Tab <> 3 Then
    If Me.SSTab1.TabEnabled(3) Then
        SSTab1.Tab = 3
    Else
        Me.Tb_debit(0).SetFocus
    End If
'    Me.Tb_Debit(0).SetFocus
End If
Call sel_text(Tb_long)
If change_coul Then
    Change_Couleur nom, Index
    mes = Rec_Mes(nom, Index)
    owner.affich_aide Me.Name, mes
End If

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

Private Sub Tb_titre_Change()
    Me.Caption = fen_titre + " : " + Tb_titre.Text
End Sub

Private Sub Txtb_deversoir_Click()
Dim mes As String
Dim nom As String
nom = "Txtb_deversoir"
mes = Rec_Mes(nom, Index)
Change_Couleur nom, Index
Change_Focus nom, Index
owner.affich_aide Me.Name, mes

End Sub

Private Sub UpDown1_Change()
Me.Tb_larg.Text = rempl_virgule(Format(Me.UpDown1.Value / 100, "###0.00"))
End Sub


Private Sub UpDown1_DownClick()
ouv_sauve = True
    Call Cmd_Leaping_Wear_Click
End Sub

Private Sub UpDown1_UpClick()
ouv_sauve = True
    Call Cmd_Leaping_Wear_Click
End Sub
Private Sub UpDown2_Change()
bKP = True
Me.Tb_dev(0) = rempl_virgule(Format(Me.UpDown2.Value / 1000, "###0.000"))
End Sub
Private Sub UpDown2_DownClick()
bKP = True
ouv_sauve = True
    Call Cmd_Leaping_Wear_Click
    bKP = False
End Sub

Private Sub UpDown2_UpClick()
bKP = True
ouv_sauve = True
    Call Cmd_Leaping_Wear_Click
    bKP = False
End Sub

