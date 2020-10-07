VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_pomp1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Station de pompage"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9825
   Icon            =   "Frm_pomp1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4530
   ScaleMode       =   0  'User
   ScaleWidth      =   7509.189
   Begin VB.Frame Frame5 
      Height          =   3825
      Left            =   6600
      TabIndex        =   71
      Top             =   480
      Width           =   3165
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   345
         Left            =   1750
         TabIndex        =   92
         Text            =   "Text1"
         Top             =   2800
         Width           =   825
      End
      Begin VB.TextBox Tb_VitRflt 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   1750
         TabIndex        =   89
         Text            =   "Text1"
         Top             =   2190
         Width           =   825
      End
      Begin VB.TextBox Tb_Drflt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1750
         TabIndex        =   84
         Top             =   1410
         Width           =   825
      End
      Begin VB.TextBox Tb_Qpomp 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   2040
         TabIndex        =   77
         Text            =   "Text1"
         Top             =   780
         Width           =   495
      End
      Begin VB.TextBox Tb_Qpomp 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   2040
         TabIndex        =   4
         Top             =   450
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   " linéaires"
         Height          =   315
         Left            =   720
         TabIndex        =   95
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   " permanent "
         Height          =   255
         Left            =   600
         TabIndex        =   94
         Top             =   2400
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   " soit "
         Height          =   255
         Index           =   14
         Left            =   1560
         TabIndex        =   93
         Top             =   810
         Width           =   315
      End
      Begin VB.Label Label4 
         Caption         =   "Perte de charges "
         Height          =   315
         Left            =   120
         TabIndex        =   91
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "m/s"
         Height          =   285
         Left            =   2640
         TabIndex        =   90
         Top             =   2220
         Width           =   345
      End
      Begin VB.Label Label2 
         Caption         =   "Vitesse en régime  "
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   2160
         Width           =   1395
      End
      Begin VB.Label Lbl_materiau 
         Caption         =   "Lbl_materiau"
         Height          =   315
         Left            =   600
         TabIndex        =   87
         Top             =   1560
         Width           =   1035
      End
      Begin VB.Label Lbl_UnitGeom 
         Caption         =   "mm"
         Height          =   285
         Index           =   6
         Left            =   2640
         TabIndex        =   86
         Top             =   1440
         Width           =   315
      End
      Begin VB.Label Lbl_IntGéom 
         Caption         =   "Canalisation retenue :"
         Height          =   285
         Index           =   7
         Left            =   60
         TabIndex        =   85
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "l/s  "
         Height          =   255
         Index           =   10
         Left            =   2640
         TabIndex        =   80
         Top             =   480
         Width           =   315
      End
      Begin VB.Label Label1 
         Caption         =   "m3/h"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   79
         Top             =   810
         Width           =   435
      End
      Begin VB.Label Lbl_intdebit 
         Caption         =   "Débit de pompage retenu"
         Height          =   255
         Index           =   7
         Left            =   60
         TabIndex        =   78
         Top             =   480
         Width           =   1875
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3915
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6906
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Débits Caractéristiques"
      TabPicture(0)   =   "Frm_pomp1.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Lbl_intdebit(4)"
      Tab(0).Control(1)=   "Label1(12)"
      Tab(0).Control(2)=   "Label1(13)"
      Tab(0).Control(3)=   "Frame1"
      Tab(0).Control(4)=   "Tb_Qpomp(0)"
      Tab(0).Control(5)=   "Tb_Qpompc(0)"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Données géomètriques"
      TabPicture(1)   =   "Frm_pomp1.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Points singuliers"
      TabPicture(2)   =   "Frm_pomp1.frx":0902
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame3"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.TextBox Tb_Qpompc 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -70320
         MaxLength       =   8
         TabIndex        =   73
         Text            =   "Text1"
         Top             =   3480
         Width           =   495
      End
      Begin VB.TextBox Tb_Qpomp 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -71640
         MaxLength       =   8
         TabIndex        =   72
         Text            =   "Text1"
         Top             =   3480
         Width           =   495
      End
      Begin VB.Frame Frame3 
         Height          =   3500
         Left            =   160
         TabIndex        =   49
         Top             =   320
         Width           =   6000
         Begin VB.Frame Frame4 
            Caption         =   "ANTI-BELIER"
            Height          =   1575
            Left            =   4560
            TabIndex        =   68
            Top             =   240
            Width           =   1305
            Begin VB.OptionButton Opt_PtSing 
               Caption         =   "NON"
               Height          =   255
               Index           =   1
               Left            =   180
               TabIndex        =   70
               Top             =   960
               Width           =   885
            End
            Begin VB.OptionButton Opt_PtSing 
               Caption         =   "OUI"
               Height          =   255
               Index           =   0
               Left            =   180
               TabIndex        =   69
               Top             =   330
               Width           =   885
            End
         End
         Begin VB.TextBox Tb_PtSing 
            Height          =   315
            Index           =   8
            Left            =   3900
            TabIndex        =   67
            Text            =   "Text1"
            Top             =   3060
            Width           =   375
         End
         Begin VB.TextBox Tb_PtSing 
            Height          =   315
            Index           =   7
            Left            =   3900
            TabIndex        =   65
            Text            =   "Text1"
            Top             =   2704
            Width           =   375
         End
         Begin VB.TextBox Tb_PtSing 
            Height          =   315
            Index           =   6
            Left            =   3900
            TabIndex        =   63
            Text            =   "Text1"
            Top             =   2352
            Width           =   375
         End
         Begin VB.TextBox Tb_PtSing 
            Height          =   315
            Index           =   5
            Left            =   3900
            TabIndex        =   61
            Text            =   "Text1"
            Top             =   2000
            Width           =   375
         End
         Begin VB.TextBox Tb_PtSing 
            Height          =   315
            Index           =   4
            Left            =   3900
            TabIndex        =   59
            Text            =   "Text1"
            Top             =   1648
            Width           =   375
         End
         Begin VB.TextBox Tb_PtSing 
            Height          =   315
            Index           =   3
            Left            =   3900
            TabIndex        =   57
            Text            =   "Text1"
            Top             =   1296
            Width           =   375
         End
         Begin VB.TextBox Tb_PtSing 
            Height          =   315
            Index           =   2
            Left            =   3900
            TabIndex        =   55
            Text            =   "Text1"
            Top             =   944
            Width           =   375
         End
         Begin VB.TextBox Tb_PtSing 
            Height          =   315
            Index           =   1
            Left            =   3900
            TabIndex        =   53
            Text            =   "Text1"
            Top             =   592
            Width           =   375
         End
         Begin VB.TextBox Tb_PtSing 
            Height          =   315
            Index           =   0
            Left            =   3900
            TabIndex        =   51
            Text            =   "Text1"
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Lbl_PtSing 
            Caption         =   "Nbre de ventouse(s)"
            Height          =   255
            Index           =   8
            Left            =   210
            TabIndex        =   66
            Top             =   3060
            Width           =   3255
         End
         Begin VB.Label Lbl_PtSing 
            Caption         =   "Nbre de système(s) de vidange"
            Height          =   255
            Index           =   7
            Left            =   210
            TabIndex        =   64
            Top             =   2715
            Width           =   3255
         End
         Begin VB.Label Lbl_PtSing 
            Caption         =   "Nbre de clapet(s) anti-retour"
            Height          =   255
            Index           =   6
            Left            =   210
            TabIndex        =   62
            Top             =   2370
            Width           =   3255
         End
         Begin VB.Label Lbl_PtSing 
            Caption         =   "Nbre de vanne(s)"
            Height          =   255
            Index           =   5
            Left            =   210
            TabIndex        =   60
            Top             =   2025
            Width           =   3255
         End
         Begin VB.Label Lbl_PtSing 
            Caption         =   "Nbre de coude(s) à 90°"
            Height          =   255
            Index           =   4
            Left            =   210
            TabIndex        =   58
            Top             =   1680
            Width           =   3255
         End
         Begin VB.Label Lbl_PtSing 
            Caption         =   "Nbre de coude(s) à 45°"
            Height          =   255
            Index           =   3
            Left            =   210
            TabIndex        =   56
            Top             =   1335
            Width           =   3255
         End
         Begin VB.Label Lbl_PtSing 
            Caption         =   "Nbre de coude(s) à 30°"
            Height          =   255
            Index           =   2
            Left            =   210
            TabIndex        =   54
            Top             =   990
            Width           =   3255
         End
         Begin VB.Label Lbl_PtSing 
            Caption         =   "Nbre de coude(s) à 22°30"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   52
            Top             =   645
            Width           =   3255
         End
         Begin VB.Label Lbl_PtSing 
            Caption         =   "Nbre de coude(s) à 11°15"
            Height          =   255
            Index           =   0
            Left            =   210
            TabIndex        =   50
            Top             =   300
            Width           =   3255
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2925
         Left            =   -74730
         TabIndex        =   30
         Top             =   600
         Width           =   5900
         Begin VB.TextBox Tb_Geom 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   3960
            TabIndex        =   81
            Top             =   900
            Width           =   825
         End
         Begin VB.TextBox Tb_Geom 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   3900
            TabIndex        =   47
            Top             =   2160
            Width           =   825
         End
         Begin VB.TextBox Tb_Geom 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   3900
            TabIndex        =   44
            Top             =   1860
            Width           =   825
         End
         Begin VB.TextBox Tb_Geom 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   3930
            TabIndex        =   41
            Top             =   1530
            Width           =   825
         End
         Begin VB.TextBox Tb_Geom 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   3930
            TabIndex        =   38
            Top             =   1200
            Width           =   825
         End
         Begin VB.ComboBox Cb_Materiau 
            Height          =   315
            Left            =   3960
            TabIndex        =   36
            Text            =   "Combo1"
            Top             =   540
            Width           =   1305
         End
         Begin VB.TextBox Tb_Geom 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   3960
            TabIndex        =   35
            Top             =   540
            Width           =   825
         End
         Begin VB.TextBox Tb_Geom 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   3900
            TabIndex        =   32
            Top             =   210
            Width           =   825
         End
         Begin VB.Label Lbl_UnitGeom 
            Caption         =   "mm"
            Height          =   285
            Index           =   1
            Left            =   4770
            TabIndex        =   83
            Top             =   930
            Width           =   555
         End
         Begin VB.Label Lbl_IntGéom 
            Caption         =   "Diamètre théorique de la canalisation (V=1.5 m/s)"
            Height          =   285
            Index           =   2
            Left            =   210
            TabIndex        =   82
            Top             =   930
            Width           =   3495
         End
         Begin VB.Label Lbl_UnitGeom 
            Caption         =   "m"
            Height          =   285
            Index           =   5
            Left            =   4740
            TabIndex        =   48
            Top             =   2160
            Width           =   555
         End
         Begin VB.Label Lbl_IntGéom 
            Caption         =   "Niveau du fil d'eau à l'extrémité du refoulement"
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   46
            Top             =   2280
            Width           =   3645
         End
         Begin VB.Label Lbl_UnitGeom 
            Caption         =   "m"
            Height          =   285
            Index           =   4
            Left            =   4740
            TabIndex        =   45
            Top             =   1860
            Width           =   555
         End
         Begin VB.Label Lbl_IntGéom 
            Caption         =   "Niveau du fil d'eau de sortie"
            Height          =   285
            Index           =   5
            Left            =   150
            TabIndex        =   43
            Top             =   1890
            Width           =   3165
         End
         Begin VB.Label Lbl_UnitGeom 
            Caption         =   "m"
            Height          =   285
            Index           =   3
            Left            =   4770
            TabIndex        =   42
            Top             =   1530
            Width           =   555
         End
         Begin VB.Label Lbl_IntGéom 
            Caption         =   "Niveau du fil d'eau d'arrivée"
            Height          =   285
            Index           =   4
            Left            =   180
            TabIndex        =   40
            Top             =   1560
            Width           =   3165
         End
         Begin VB.Label Lbl_UnitGeom 
            Caption         =   "m"
            Height          =   285
            Index           =   2
            Left            =   4770
            TabIndex        =   39
            Top             =   1200
            Width           =   555
         End
         Begin VB.Label Lbl_IntGéom 
            Caption         =   "Niveau du terrain naturel"
            Height          =   285
            Index           =   3
            Left            =   180
            TabIndex        =   37
            Top             =   1230
            Width           =   3165
         End
         Begin VB.Label Lbl_IntGéom 
            Caption         =   "Nature du tuyau"
            Height          =   285
            Index           =   1
            Left            =   210
            TabIndex        =   34
            Top             =   570
            Width           =   3165
         End
         Begin VB.Label Lbl_UnitGeom 
            Caption         =   "m"
            Height          =   285
            Index           =   0
            Left            =   4740
            TabIndex        =   33
            Top             =   210
            Width           =   555
         End
         Begin VB.Label Lbl_IntGéom 
            Caption         =   "Longueur du refoulement"
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   31
            Top             =   270
            Width           =   3165
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
         Height          =   3015
         Left            =   -74640
         TabIndex        =   5
         Top             =   360
         Width           =   5535
         Begin VB.TextBox Tb_Debitc 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   4320
            MaxLength       =   8
            TabIndex        =   13
            Text            =   "Text1"
            Top             =   2580
            Width           =   500
         End
         Begin VB.TextBox Tb_Debitc 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   4320
            MaxLength       =   8
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   2160
            Width           =   500
         End
         Begin VB.TextBox Tb_Debitc 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   4320
            MaxLength       =   8
            TabIndex        =   11
            Text            =   "Text1"
            Top             =   1620
            Width           =   500
         End
         Begin VB.TextBox Tb_Debitc 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   4320
            MaxLength       =   8
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   1200
            Width           =   500
         End
         Begin VB.TextBox Tb_Debitc 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   4320
            MaxLength       =   8
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   300
            Width           =   500
         End
         Begin VB.TextBox Tb_Debit 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   3030
            MaxLength       =   8
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   300
            Width           =   500
         End
         Begin VB.TextBox Tb_FPointe 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3030
            MaxLength       =   8
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   735
            Width           =   500
         End
         Begin VB.TextBox Tb_Debit 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   3030
            MaxLength       =   8
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   1200
            Width           =   500
         End
         Begin VB.TextBox Tb_Debit 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   3030
            MaxLength       =   8
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   1620
            Width           =   500
         End
         Begin VB.TextBox Tb_Debit 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   3030
            MaxLength       =   8
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   2160
            Width           =   500
         End
         Begin VB.TextBox Tb_Debit 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   3030
            MaxLength       =   8
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   2580
            Width           =   500
         End
         Begin VB.Label Label1 
            Caption         =   "m3/h"
            Height          =   255
            Index           =   11
            Left            =   4920
            TabIndex        =   29
            Top             =   330
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "m3/h"
            Height          =   255
            Index           =   9
            Left            =   4860
            TabIndex        =   28
            Top             =   1230
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "m3/h"
            Height          =   255
            Index           =   8
            Left            =   4890
            TabIndex        =   27
            Top             =   1650
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "m3/h"
            Height          =   255
            Index           =   7
            Left            =   4920
            TabIndex        =   26
            Top             =   2160
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "m3/h"
            Height          =   255
            Index           =   6
            Left            =   4920
            TabIndex        =   25
            Top             =   2610
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "l/s soit "
            Height          =   255
            Index           =   5
            Left            =   3630
            TabIndex        =   24
            Top             =   2610
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "l/s soit "
            Height          =   255
            Index           =   4
            Left            =   3630
            TabIndex        =   23
            Top             =   2160
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "l/s soit "
            Height          =   255
            Index           =   3
            Left            =   3630
            TabIndex        =   22
            Top             =   1650
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "l/s soit "
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   21
            Top             =   1230
            Width           =   585
         End
         Begin VB.Label Label1 
            Caption         =   "l/s soit "
            Height          =   255
            Index           =   0
            Left            =   3660
            TabIndex        =   20
            Top             =   330
            Width           =   585
         End
         Begin VB.Label Lbl_intdebit 
            Caption         =   "Débit moyen des eaux usées"
            Height          =   255
            Index           =   0
            Left            =   300
            TabIndex        =   19
            Top             =   330
            Width           =   2235
         End
         Begin VB.Label Lbl_intdebit 
            Caption         =   "Facteur de pointe"
            Height          =   255
            Index           =   1
            Left            =   300
            TabIndex        =   18
            Top             =   780
            Width           =   2235
         End
         Begin VB.Label Lbl_intdebit 
            Caption         =   "Débit pointe des eaux usées"
            Height          =   255
            Index           =   2
            Left            =   300
            TabIndex        =   17
            Top             =   1200
            Width           =   2235
         End
         Begin VB.Label Lbl_intdebit 
            Caption         =   "Débit des eaux parasites"
            Height          =   255
            Index           =   3
            Left            =   300
            TabIndex        =   16
            Top             =   1650
            Width           =   2235
         End
         Begin VB.Label Lbl_intdebit 
            Caption         =   "Débit moyen de temps sec"
            Height          =   255
            Index           =   5
            Left            =   300
            TabIndex        =   15
            Top             =   2130
            Width           =   2235
         End
         Begin VB.Label Lbl_intdebit 
            Caption         =   "Débit de pointe de temps sec"
            Height          =   255
            Index           =   6
            Left            =   300
            TabIndex        =   14
            Top             =   2580
            Width           =   2235
         End
      End
      Begin VB.Label Label1 
         Caption         =   "l/s soit "
         Height          =   255
         Index           =   13
         Left            =   -71040
         TabIndex        =   76
         Top             =   3510
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "m3/h"
         Height          =   255
         Index           =   12
         Left            =   -69780
         TabIndex        =   75
         Top             =   3510
         Width           =   585
      End
      Begin VB.Label Lbl_intdebit 
         Caption         =   "Débit de pompage théorique"
         Height          =   255
         Index           =   4
         Left            =   -74400
         TabIndex        =   74
         Top             =   3480
         Width           =   2235
      End
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
   End
End
Attribute VB_Name = "Frm_pomp1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cb_Materiau_Click()
    Lbl_materiau.Caption = Cb_Materiau.Text & " de D = "
    Tb_Drflt.Text = Format(val(Tb_Geom(2).Text), "####")
End Sub

Private Sub Form_Load()
For i% = 0 To 9
    Tb_Debit(i%).Text = ""
    Select Case i%
        Case 1
            Tb_Debit(i%).Enabled = False
        Case 3 To 9
            Tb_Debit(i%).Enabled = False
            
    End Select
Next
With Cb_Materiau
    .Clear
    .AddItem "FONTE"
    .AddItem "PEHD"
    .AddItem "PVC"
    .AddItem "ACIER"
    .Text = ""
End With
SSTab1.Tab = 0
For i% = 0 To 8
    Tb_PtSing(i%).Text = ""
Next
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case PreviousTab
    Case 0
        
End Select

End Sub

Private Sub tb_debit_Change(Index As Integer)
If Index < 5 Then
    Tb_Debit(Index + 5).Text = val(Tb_Debit(Index)) * 3.6
End If
'Facteur de pointe =somme de 1.5 et du quotient de 2.5 par la racine du débit moyen des EU (index=0)
'En général limité à 4 mais peut être imposé par le projeteur
'Fp
    If Index = 0 And val(Tb_Debit(0).Text) > 0 Then Tb_FPointe.Text = Format(1.5 + 2.5 / ((val(Tb_Debit(0).Text)) ^ 0.5), "###.#")
'Debit de pointe EU (index=1)= Produit du débit moyen EU (index= 0)par Facteur de Pointe (index=1)
'Qeu
    Tb_Debit(1).Text = Format(val(Tb_Debit(0).Text) * val(Tb_FPointe.Text), "###.#")
'Debit moyen de temps sec (index=3)= Somme du débit moyen EU (index= 0)et du débit des eaux parasites (index= 2)
'Qmts
    Tb_Debit(3).Text = Format(val(Tb_Debit(0).Text) + val(Tb_Debit(2).Text), "###.#")
'Debit de pointe de temps sec (index=4)= Somme du débit Pointe EU (index= 2)et du débit des eaux parasites (index= 3)
'Qts
    Tb_Debit(4).Text = Format(val(Tb_Debit(1).Text) + val(Tb_Debit(2).Text), "###.#")

'Débit de pompage = 3 fois le débit moyen de temps sec(index=3)
'Qpomp
    Tb_Qpomp(0).Text = 3 * val(Tb_Debit(3).Text)
End Sub

Private Sub Tb_Drflt_Change()
    Tb_VitRflt.Text = Format(4000 * val(Tb_Qpomp(2).Text) / 3.14 / ((val(Tb_Drflt.Text)) ^ 2), "##.##")
End Sub

Private Sub Tb_FPointe_Change()
    Tb_Debit(1).Text = Format(val(Tb_Debit(0).Text) * val(Tb_FPointe.Text), "###.#")

End Sub


Private Sub Tb_Qpomp_Change(Index As Integer)
    Tb_Qpomp(1).Text = 3.6 * val(Tb_Qpomp(0).Text)
    Tb_Qpomp(3).Text = 3.6 * val(Tb_Qpomp(2).Text)
    'calcul du diamètre de la canalisation de refoulement avec une vitesse par défaut de 1.5 m/s
    ' Diamètre égal
    Tb_Geom(2).Text = Format(2000 * (val(Tb_Qpomp(2).Text) / (1000 * 3.14 * 1.5)) ^ 0.5, "####")
    
End Sub

Private Sub Tb_Qpomp_LostFocus(Index As Integer)
    'vérifier que le débit de pompage est supérieur au débit de pointe de temps sec
        If val(Tb_Debit(4).Text) > val(Tb_Qpomp(2).Text) Then
            mes$ = "Le débit de pompage est inférieur au débit de pointe de temps sec" & Chr(13)
            mes$ = mes$ + "SOUHAITEZ-VOUS MODIFIER ?"
            MsgBox mes
            SSTab1.Tab = 0
            Tb_Qpomp(2).SetFocus
            Tb_Qpomp(2).SelStart = 0
            Tb_Qpomp(2).SelLength = Len(Tb_Qpomp(0).Text)
        End If

End Sub

