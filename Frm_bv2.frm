VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frm_bv2 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hydraulique Bassin Versant"
   ClientHeight    =   4320
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9825
   Icon            =   "Frm_bv2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   9825
   Begin VB.TextBox Tb_txtqf 
      Height          =   285
      Left            =   2520
      MaxLength       =   7
      TabIndex        =   113
      TabStop         =   0   'False
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Cmd_retour 
      Caption         =   "Retour"
      Height          =   255
      Left            =   6480
      TabIndex        =   111
      Top             =   0
      Width           =   2895
   End
   Begin VB.ComboBox Cb_bassin 
      Height          =   315
      Left            =   360
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   0
      Width           =   4000
   End
   Begin VB.TextBox Tb_debit1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Index           =   6
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   3840
      Width           =   800
   End
   Begin VB.TextBox Tb_debit1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Index           =   5
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   3525
      Width           =   800
   End
   Begin VB.CommandButton Cmd_resu 
      Caption         =   "Calculer"
      Height          =   255
      Left            =   7680
      TabIndex        =   58
      TabStop         =   0   'False
      ToolTipText     =   "Calcul du bassin versant"
      Top             =   3960
      Width           =   1000
   End
   Begin VB.TextBox Tb_Qbrut 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   4920
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Frame Frm_debit 
      Caption         =   "D�bit d'eau pluviale"
      Height          =   1335
      Left            =   360
      TabIndex        =   47
      Top             =   2880
      Width           =   3375
      Begin VB.TextBox Tb_debit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   240
         Width           =   800
      End
      Begin VB.TextBox Tb_debit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   560
         Width           =   800
      End
      Begin VB.TextBox Tb_debit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   880
         Width           =   800
      End
      Begin VB.OptionButton Ob_Caquot 
         Height          =   255
         Left            =   3000
         TabIndex        =   50
         TabStop         =   0   'False
         ToolTipText     =   "S�lection du d�bit � retourner � l'ouvrage appelant"
         Top             =   280
         Width           =   250
      End
      Begin VB.OptionButton Ob_Mr 
         Height          =   255
         Left            =   3000
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   600
         Width           =   250
      End
      Begin VB.OptionButton Ob_Mh 
         Height          =   255
         Left            =   3000
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   920
         Width           =   250
      End
      Begin VB.Label Lb_udebit 
         Caption         =   "l/s"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   103
         Top             =   920
         Width           =   255
      End
      Begin VB.Label Lb_udebit 
         Caption         =   "l/s"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   102
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Lb_udebit 
         Caption         =   "l/s"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   101
         Top             =   280
         Width           =   255
      End
      Begin VB.Label Lb_debit 
         Caption         =   "M�thode Rationnelle"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   53
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Lb_debit 
         Caption         =   "M�thode de Caquot"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   52
         Top             =   280
         Width           =   1455
      End
      Begin VB.Label Lb_debit 
         Caption         =   "M�thode Hydrogramme"
         Height          =   420
         Index           =   2
         Left            =   120
         TabIndex        =   51
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.TextBox Tb_debit1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Index           =   3
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   3210
      Width           =   800
   End
   Begin VB.TextBox Tb_debit1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Index           =   1
      Left            =   8400
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   2880
      Width           =   800
   End
   Begin VB.TextBox Tb_debit1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Index           =   4
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   3525
      Width           =   800
   End
   Begin VB.TextBox Tb_debit1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Index           =   0
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   2880
      Width           =   800
   End
   Begin VB.TextBox Tb_debit1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Index           =   2
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3210
      Width           =   800
   End
   Begin VB.OptionButton Opt_rural 
      Caption         =   "Rural"
      Height          =   255
      Left            =   1440
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   0
      Width           =   855
   End
   Begin VB.OptionButton Opt_urbain 
      Caption         =   "Urbain"
      Height          =   255
      Left            =   360
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   0
      Value           =   -1  'True
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2410
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4260
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   406
      BackColor       =   0
      TabCaption(0)   =   "Caract�ristiques"
      TabPicture(0)   =   "Frm_bv2.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image21"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frm_cep"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frm_ceu"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frm_cbr"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Param�tres"
      TabPicture(1)   =   "Frm_bv2.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Cmd_calc"
      Tab(1).Control(1)=   "Tb_par_pl(5)"
      Tab(1).Control(2)=   "Tb_par_pl(4)"
      Tab(1).Control(3)=   "Tb_par_pl(3)"
      Tab(1).Control(4)=   "Tb_par_pl(2)"
      Tab(1).Control(5)=   "Tb_par_pl(1)"
      Tab(1).Control(6)=   "Tb_par_pl(0)"
      Tab(1).Control(7)=   "Frm_peu"
      Tab(1).Control(8)=   "Frm_pep"
      Tab(1).Control(9)=   "Frm_ppr"
      Tab(1).Control(10)=   "Lb_upar_pl(4)"
      Tab(1).ControlCount=   11
      Begin VB.CommandButton Cmd_calc 
         Caption         =   ">"
         Height          =   255
         Left            =   -65640
         TabIndex        =   75
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox Tb_par_pl 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   -68040
         MaxLength       =   3
         TabIndex        =   74
         Text            =   "5"
         Top             =   1995
         Width           =   900
      End
      Begin VB.TextBox Tb_par_pl 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   -68040
         MaxLength       =   4
         TabIndex        =   73
         Text            =   "Teta"
         Top             =   1680
         Width           =   900
      End
      Begin VB.TextBox Tb_par_pl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Height          =   285
         Index           =   3
         Left            =   -68040
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   72
         Text            =   "Vruis"
         Top             =   1365
         Width           =   900
      End
      Begin VB.TextBox Tb_par_pl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Height          =   285
         Index           =   2
         Left            =   -68040
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   71
         Text            =   "Qmax"
         Top             =   1035
         Width           =   900
      End
      Begin VB.TextBox Tb_par_pl 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   -68040
         MaxLength       =   7
         TabIndex        =   70
         Text            =   "Vruis"
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox Tb_par_pl 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   -68040
         MaxLength       =   7
         TabIndex        =   69
         Text            =   " Qmax"
         Top             =   405
         Width           =   900
      End
      Begin VB.Frame Frm_cbr 
         Caption         =   "Caract�ristiques"
         Height          =   1950
         Left            =   4320
         TabIndex        =   60
         Top             =   300
         Visible         =   0   'False
         Width           =   4455
         Begin VB.TextBox Tb_carep_rur 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   4
            Left            =   2520
            MaxLength       =   4
            TabIndex        =   12
            Text            =   "Trep"
            Top             =   1560
            Width           =   900
         End
         Begin VB.TextBox Tb_carep_rur 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   3
            Left            =   3000
            MaxLength       =   6
            TabIndex        =   11
            Text            =   "b"
            Top             =   1200
            Width           =   700
         End
         Begin VB.TextBox Tb_carep_rur 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   2160
            MaxLength       =   6
            TabIndex        =   10
            Text            =   "a"
            Top             =   1200
            Width           =   700
         End
         Begin VB.TextBox Tb_carep_rur 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   2520
            MaxLength       =   4
            TabIndex        =   9
            Text            =   "Vinf"
            Top             =   600
            Width           =   900
         End
         Begin VB.TextBox Tb_carep_rur 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   2520
            MaxLength       =   4
            TabIndex        =   8
            Text            =   "Pert"
            Top             =   240
            Width           =   900
         End
         Begin VB.Label Lb_carep_rur 
            Caption         =   "Loi de HORTON"
            Height          =   135
            Index           =   2
            Left            =   240
            TabIndex        =   89
            Top             =   960
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Lb_ucarep_rur 
            Caption         =   "mn"
            Height          =   255
            Index           =   4
            Left            =   3600
            TabIndex        =   88
            Top             =   1605
            Width           =   615
         End
         Begin VB.Label Lb_ucarep_rur 
            Height          =   255
            Index           =   3
            Left            =   4080
            TabIndex        =   87
            Top             =   1245
            Width           =   135
         End
         Begin VB.Label Lb_ucarep_rur 
            Height          =   255
            Index           =   2
            Left            =   3840
            TabIndex        =   86
            Top             =   1245
            Width           =   135
         End
         Begin VB.Label Lb_ucarep_rur 
            Caption         =   "mm/h"
            Height          =   255
            Index           =   1
            Left            =   3600
            TabIndex        =   85
            Top             =   645
            Width           =   615
         End
         Begin VB.Label Lb_ucarep_rur 
            Caption         =   "mm"
            Height          =   255
            Index           =   0
            Left            =   3600
            TabIndex        =   84
            Top             =   285
            Width           =   615
         End
         Begin VB.Label Lb_carep_rur 
            Caption         =   "Temps de r�ponse"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   66
            Top             =   1605
            Width           =   1815
         End
         Begin VB.Label Lb_bHorton 
            Alignment       =   2  'Center
            Caption         =   "b"
            Height          =   255
            Left            =   3000
            TabIndex        =   65
            Top             =   960
            Width           =   705
         End
         Begin VB.Label Lb_aHorton 
            Alignment       =   2  'Center
            Caption         =   "a"
            Height          =   255
            Left            =   2160
            TabIndex        =   64
            Top             =   960
            Width           =   705
         End
         Begin VB.Label Lb_carep_rur 
            Caption         =   "Loi de HORTON "
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   63
            Top             =   1245
            Width           =   1695
         End
         Begin VB.Label Lb_carep_rur 
            Caption         =   "Vitesse limite d'infiltration "
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   62
            Top             =   645
            Width           =   2055
         End
         Begin VB.Label Lb_carep_rur 
            Caption         =   "Pertes initiales "
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   61
            Top             =   285
            Width           =   1455
         End
      End
      Begin VB.Frame Frm_peu 
         Caption         =   "Eau us�e"
         Height          =   1260
         Left            =   -74790
         TabIndex        =   30
         Top             =   1070
         Width           =   4335
         Begin VB.TextBox Tb_par_eu 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   2400
            MaxLength       =   6
            TabIndex        =   15
            Text            =   "Lcrin"
            Top             =   240
            Width           =   900
         End
         Begin VB.TextBox Tb_par_eu 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   16
            Text            =   "a"
            Top             =   840
            Width           =   525
         End
         Begin VB.TextBox Tb_par_eu 
            Alignment       =   1  'Right Justify
            DataField       =   "  "
            Height          =   300
            Index           =   2
            Left            =   2760
            MaxLength       =   4
            TabIndex        =   17
            Text            =   "b"
            Top             =   840
            Width           =   525
         End
         Begin VB.Label Lb_upar_eu 
            Height          =   255
            Index           =   2
            Left            =   3840
            TabIndex        =   99
            Top             =   890
            Width           =   135
         End
         Begin VB.Label Lb_upar_eu 
            Height          =   255
            Index           =   1
            Left            =   3480
            TabIndex        =   98
            Top             =   890
            Width           =   135
         End
         Begin VB.Label Lb_upar_eu 
            Caption         =   "l/hab/s"
            Height          =   255
            Index           =   0
            Left            =   3480
            TabIndex        =   97
            Top             =   290
            Width           =   735
         End
         Begin VB.Label Lb_par_eu 
            Caption         =   "Coefficient pointe EU"
            Height          =   135
            Index           =   1
            Left            =   240
            TabIndex        =   94
            Top             =   600
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Lb_par_eu 
            Caption         =   "Intensit� pluie de rin�age"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   34
            Top             =   290
            Width           =   1935
         End
         Begin VB.Label Lb_par_eu 
            Caption         =   "Coefficient pointe EU"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   33
            Top             =   890
            Width           =   1600
         End
         Begin VB.Label Lb_aeu 
            Alignment       =   2  'Center
            Caption         =   "a"
            Height          =   300
            Left            =   2040
            TabIndex        =   32
            Top             =   600
            Width           =   525
         End
         Begin VB.Label Lb_beu 
            Alignment       =   2  'Center
            Caption         =   "b"
            Height          =   300
            Left            =   2760
            TabIndex        =   31
            Top             =   600
            Width           =   525
         End
      End
      Begin VB.Frame Frm_pep 
         Caption         =   "Eau pluviale"
         Height          =   780
         Left            =   -74790
         TabIndex        =   26
         Top             =   300
         Width           =   4335
         Begin VB.TextBox tb_par_ep 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   2040
            MaxLength       =   4
            TabIndex        =   13
            Text            =   "a"
            Top             =   360
            Width           =   525
         End
         Begin VB.TextBox tb_par_ep 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   2760
            MaxLength       =   4
            TabIndex        =   14
            Text            =   "b"
            Top             =   360
            Width           =   525
         End
         Begin VB.Label Lb_upar_ep 
            Height          =   255
            Index           =   1
            Left            =   3840
            TabIndex        =   96
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Lb_upar_ep 
            Height          =   255
            Index           =   0
            Left            =   3480
            TabIndex        =   95
            Top             =   360
            Width           =   135
         End
         Begin VB.Label Lb_par_ep 
            Caption         =   "Coefficient Montana"
            Height          =   15
            Index           =   0
            Left            =   1200
            TabIndex        =   93
            Top             =   240
            Visible         =   0   'False
            Width           =   800
         End
         Begin VB.Label Lb_par_ep 
            BackColor       =   &H80000004&
            Caption         =   "Coefficient Montana"
            Height          =   300
            Index           =   1
            Left            =   240
            TabIndex        =   29
            Top             =   410
            Width           =   1600
         End
         Begin VB.Label Lb_amontana 
            Alignment       =   2  'Center
            Caption         =   "a"
            Height          =   300
            Index           =   1
            Left            =   2040
            TabIndex        =   28
            Top             =   120
            Width           =   525
         End
         Begin VB.Label Lb_bmontana 
            Alignment       =   2  'Center
            Caption         =   "b"
            Height          =   300
            Index           =   2
            Left            =   2760
            TabIndex        =   27
            Top             =   120
            Width           =   525
         End
      End
      Begin VB.Frame Frm_ceu 
         Caption         =   "Eau us�e"
         Height          =   1455
         Left            =   4680
         TabIndex        =   22
         Top             =   600
         Width           =   3975
         Begin VB.TextBox Tb_car_eu 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   6
            Text            =   "Ceau"
            Top             =   600
            Width           =   900
         End
         Begin VB.TextBox Tb_car_eu 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   7
            Text            =   "Tdilu"
            Top             =   960
            Width           =   900
         End
         Begin VB.TextBox Tb_car_eu 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   1920
            MaxLength       =   9
            TabIndex        =   5
            Text            =   "Nhab"
            Top             =   240
            Width           =   900
         End
         Begin VB.Label Lb_ucar_eu 
            Caption         =   "%"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   92
            Top             =   1010
            Width           =   735
         End
         Begin VB.Label Lb_ucar_eu 
            Caption         =   "l/hab/j"
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   91
            Top             =   650
            Width           =   735
         End
         Begin VB.Label Lb_ucar_eu 
            Height          =   255
            Index           =   0
            Left            =   3000
            TabIndex        =   90
            Top             =   290
            Width           =   735
         End
         Begin VB.Label Lb_car_eu 
            Caption         =   "Consommation eau"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   59
            Top             =   650
            Width           =   1575
         End
         Begin VB.Label Lb_car_eu 
            Caption         =   "Taux de dilution"
            Height          =   300
            Index           =   2
            Left            =   240
            TabIndex        =   25
            Top             =   1010
            Width           =   1575
         End
         Begin VB.Label Lb_car_eu 
            Caption         =   "Nbre d'habitants"
            Height          =   300
            Index           =   0
            Left            =   240
            TabIndex        =   24
            Top             =   290
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            Height          =   135
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Width           =   15
         End
      End
      Begin VB.Frame Frm_cep 
         Caption         =   "Eau pluviale"
         Height          =   1950
         Left            =   240
         TabIndex        =   4
         Top             =   300
         Width           =   3975
         Begin VB.TextBox Tb_car_ep 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   114
            Text            =   "S"
            Top             =   360
            Width           =   900
         End
         Begin VB.TextBox Tb_car_ep 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   3
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   3
            Text            =   "C"
            Top             =   1440
            Width           =   900
         End
         Begin VB.TextBox Tb_car_ep 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   2
            Text            =   "P"
            Top             =   1080
            Width           =   900
         End
         Begin VB.TextBox Tb_car_ep 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   1920
            MaxLength       =   6
            TabIndex        =   1
            Text            =   "L"
            Top             =   720
            Width           =   900
         End
         Begin VB.Label Lb_ucar_ep 
            Caption         =   "%"
            Height          =   255
            Index           =   3
            Left            =   3000
            TabIndex        =   83
            Top             =   1490
            Width           =   735
         End
         Begin VB.Label Lb_ucar_ep 
            Caption         =   "1/10000"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   82
            Top             =   1125
            Width           =   735
         End
         Begin VB.Label Lb_ucar_ep 
            Caption         =   "m"
            Height          =   255
            Index           =   1
            Left            =   3000
            TabIndex        =   81
            Top             =   770
            Width           =   735
         End
         Begin VB.Label Lb_ucar_ep 
            Caption         =   "Ha"
            Height          =   255
            Index           =   0
            Left            =   3000
            TabIndex        =   80
            Top             =   410
            Width           =   735
         End
         Begin VB.Label Lb_car_ep 
            Caption         =   "Coef. de ruissellement"
            Height          =   300
            Index           =   3
            Left            =   240
            TabIndex        =   21
            Top             =   1490
            Width           =   1695
         End
         Begin VB.Label Lb_car_ep 
            Caption         =   "Pente "
            Height          =   300
            Index           =   2
            Left            =   240
            TabIndex        =   20
            Top             =   1130
            Width           =   1695
         End
         Begin VB.Label Lb_car_ep 
            Caption         =   "Longueur "
            Height          =   300
            Index           =   1
            Left            =   240
            TabIndex        =   19
            Top             =   770
            Width           =   1695
         End
         Begin VB.Label Lb_car_ep 
            Caption         =   "Surface "
            Height          =   300
            Index           =   0
            Left            =   240
            TabIndex        =   18
            Top             =   410
            Width           =   1695
         End
      End
      Begin VB.Frame Frm_ppr 
         Caption         =   "Pluie de projet"
         Height          =   2100
         Left            =   -70320
         TabIndex        =   115
         Top             =   240
         Width           =   4095
         Begin VB.Label Lb_upar_pl 
            Caption         =   "mn"
            Height          =   255
            Index           =   0
            Left            =   3400
            TabIndex        =   126
            Top             =   240
            Width           =   550
         End
         Begin VB.Label Lb_upar_pl 
            Caption         =   "mn"
            Height          =   255
            Index           =   1
            Left            =   3400
            TabIndex        =   125
            Top             =   510
            Width           =   550
         End
         Begin VB.Label Lb_upar_pl 
            Caption         =   "mm"
            Height          =   255
            Index           =   2
            Left            =   3400
            TabIndex        =   124
            Top             =   840
            Width           =   550
         End
         Begin VB.Label Lb_upar_pl 
            Caption         =   "mm"
            Height          =   255
            Index           =   3
            Left            =   3400
            TabIndex        =   123
            Top             =   1170
            Width           =   550
         End
         Begin VB.Label Lb_upar_pl 
            Caption         =   "mn"
            Height          =   255
            Index           =   5
            Left            =   3400
            TabIndex        =   122
            Top             =   1800
            Width           =   550
         End
         Begin VB.Label Lb_par_pl 
            Caption         =   "Hauteur totale"
            Height          =   255
            Index           =   2
            Left            =   250
            TabIndex        =   121
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Lb_par_pl 
            Caption         =   "Hauteur intense"
            Height          =   255
            Index           =   3
            Left            =   250
            TabIndex        =   120
            Top             =   1170
            Width           =   1575
         End
         Begin VB.Label Lb_par_pl 
            Caption         =   "D�calage de la pointe"
            Height          =   255
            Index           =   4
            Left            =   250
            TabIndex        =   119
            Top             =   1500
            Width           =   1575
         End
         Begin VB.Label Lb_par_pl 
            Caption         =   "Pas de calcul"
            Height          =   255
            Index           =   5
            Left            =   250
            TabIndex        =   118
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Lb_par_pl 
            Caption         =   "Dur�e intense"
            Height          =   255
            Index           =   1
            Left            =   250
            TabIndex        =   117
            Top             =   510
            Width           =   1575
         End
         Begin VB.Label Lb_par_pl 
            Caption         =   "Dur�e totale"
            Height          =   255
            Index           =   0
            Left            =   255
            TabIndex        =   116
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Label Lb_upar_pl 
         Height          =   255
         Index           =   4
         Left            =   -67200
         TabIndex        =   100
         Top             =   1845
         Width           =   630
      End
      Begin VB.Image Image21 
         Height          =   375
         Left            =   480
         Picture         =   "Frm_bv2.frx":0902
         Top             =   6480
         Width           =   300
      End
   End
   Begin VB.TextBox Tb_titre 
      Height          =   300
      Left            =   5760
      MaxLength       =   30
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   3180
   End
   Begin VB.TextBox Tb_ruic 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Height          =   285
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   127
      Top             =   3525
      Width           =   800
   End
   Begin VB.Label Lb_txtvstock 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      Caption         =   "0"
      Height          =   240
      Left            =   3600
      TabIndex        =   112
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Lb_udebit1 
      Caption         =   "l/s"
      Height          =   255
      Index           =   6
      Left            =   6600
      TabIndex        =   110
      Top             =   3885
      Width           =   255
   End
   Begin VB.Label Lb_udebit1 
      Caption         =   "l/s"
      Height          =   255
      Index           =   5
      Left            =   9240
      TabIndex        =   109
      Top             =   3555
      Width           =   255
   End
   Begin VB.Label Lb_udebit1 
      Caption         =   "l/s"
      Height          =   255
      Index           =   4
      Left            =   6600
      TabIndex        =   108
      Top             =   3555
      Width           =   255
   End
   Begin VB.Label Lb_udebit1 
      Caption         =   "l/s"
      Height          =   255
      Index           =   3
      Left            =   9240
      TabIndex        =   107
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Lb_udebit1 
      Caption         =   "l/s"
      Height          =   255
      Index           =   2
      Left            =   6600
      TabIndex        =   106
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Lb_udebit1 
      Caption         =   "l/s"
      Height          =   255
      Index           =   1
      Left            =   9240
      TabIndex        =   105
      Top             =   2925
      Width           =   255
   End
   Begin VB.Label Lb_udebit1 
      Caption         =   "l/s"
      Height          =   255
      Index           =   0
      Left            =   6600
      TabIndex        =   104
      Top             =   2925
      Width           =   255
   End
   Begin VB.Label Lb_debit1 
      Caption         =   "Volume total ruissel�"
      Height          =   240
      Index           =   6
      Left            =   3840
      TabIndex        =   78
      Top             =   3885
      Width           =   1575
   End
   Begin VB.Label Lb_debit1 
      Caption         =   "D�bit de pointe"
      Height          =   240
      Index           =   5
      Left            =   6960
      TabIndex        =   76
      Top             =   3555
      Width           =   1215
   End
   Begin VB.Label Lb_debit1 
      Caption         =   "D�bit de rin�age"
      Height          =   300
      Index           =   3
      Left            =   6960
      TabIndex        =   46
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Lb_debit1 
      Caption         =   "D�bit de temps sec"
      Height          =   300
      Index           =   1
      Left            =   6960
      TabIndex        =   44
      Top             =   2925
      Width           =   1455
   End
   Begin VB.Label Lb_debit1 
      Caption         =   "D�bit de pluie de rin�age"
      Height          =   300
      Index           =   4
      Left            =   3840
      TabIndex        =   43
      Top             =   3555
      Width           =   1815
   End
   Begin VB.Label Lb_debit1 
      Caption         =   "D�bit des eaux us�es"
      Height          =   300
      Index           =   0
      Left            =   3840
      TabIndex        =   40
      Top             =   2925
      Width           =   1665
   End
   Begin VB.Label Lb_debit1 
      Caption         =   "D�bit des eaux claires"
      Height          =   300
      Index           =   2
      Left            =   3840
      TabIndex        =   39
      Top             =   3240
      Width           =   1665
   End
   Begin VB.Label Lb_ruic 
      Caption         =   "Coef. de ruisselement"
      Height          =   300
      Left            =   3840
      TabIndex        =   128
      Top             =   3555
      Width           =   1815
   End
   Begin VB.Label Lb_uruic 
      Caption         =   "l/s"
      Height          =   255
      Left            =   6600
      TabIndex        =   129
      Top             =   3555
      Width           =   255
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "&Bassin versant"
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
      Begin VB.Menu MnuPrint 
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
Attribute VB_Name = "frm_bv2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private okg As Boolean
Private owner As MDIFrm_menu
Private esave As st_save
Private bv_charge As Boolean
Public nom_ouvrage As String
Public nom_type As String
'Private nom_fich As String
Private lhFicDbf As Long
Private FileLength As Integer
Private list_don1() As Variant
Private list_don2() As Variant
Private list_int1() As Variant
Private list_resu1() As Variant
Private ba_texte As String
Private fen_titre As String
Public titre_sav As String
Private list_tb() As Variant
Private sval_champ As String
Private iSels As Integer
Private iSell As Integer
Private bKP As Boolean

Public Function get_l_tb() As Variant
get_l_tb = list_tb
End Function

Public Sub retailler()
retaille
End Sub
Private Sub retaille()
    Me.Left = owner.fcom.Width
    Me.Top = 0
'    Me.Width = owner.Width - owner.fcom.Width - 200
'    Me.Height = owner.fdessin.Top
    Me.Width = maximum(9940, owner.Width - owner.fcom.Width - 200) ' 10040
    Me.Height = maximum(3850, owner.fdessin.Top) '4600
End Sub
Private Sub affiche_debit_ep(ByRef ebv As st_Bv)
Me.Tb_Qbrut.Text = rempl_virgule(Format(ebv.Qbrut * 1000, "####0.0"))
Me.Tb_debit(0).Text = rempl_virgule(Format(ebv.Qcor * 1000, "####0.0"))
Me.Ob_Caquot.Enabled = True
End Sub
Private Sub affiche_debit_epmr(ByRef ebv As st_Bv)
Me.Tb_debit(1).Text = rempl_virgule(Format(ebv.Qmr * 1000, "####0.0"))
Me.Ob_Mr.Enabled = True
'Me.tb_debit(0).Text = Format(ebv.Qcor * 1000, "####0.0")
End Sub
Private Sub affiche_debit_eu(ByRef ebv As st_Bv)
Me.Tb_debit1(0).Text = rempl_virgule(Format(ebv.Qeu, "###0.0"))
Me.Tb_debit1(2).Text = rempl_virgule(Format(ebv.Qecp, "###0.0"))
Me.Tb_debit1(1).Text = rempl_virgule(Format(ebv.Qts, "###0.0"))
Me.Tb_debit1(4).Text = rempl_virgule(Format(ebv.Qprin, "###0.0"))
Me.Tb_debit1(3).Text = rempl_virgule(Format(ebv.Qrin, "###0.0"))
End Sub
Public Sub ini_bv0()
    ebv.nom = ""
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
Sub ini_schema()
    Me.Tb_par_pl(0).Text = "0"
    Me.Tb_par_pl(1).Text = "0"
    Me.Tb_par_pl(2).Text = "0"
    Me.Tb_par_pl(3).Text = "0"
    Me.Tb_debit1(5).Text = "0"
    Me.Tb_debit1(6).Text = "0"
    Me.Tb_ruic.Text = "0.0"
    Me.Tb_txtqf.Text = "0"
    Me.Tb_par_pl(5) = "5"
    Me.Tb_par_pl(4) = "0.5"
    Me.Lb_txtvstock.Caption = 0
'    Me.hydro.Visible = False
'    Me.hyeto.Visible = False
    owner.fdessin.UC_graphique1.Visible = False
End Sub
Sub C_bv()
 ebv.surface = val(Me.Tb_car_ep(0).Text)
 ebv.imper = val(Me.Tb_car_ep(3).Text)
 ebv.lghydr = val(Me.Tb_car_ep(1).Text)
 ebv.phydr = val(Me.Tb_car_ep(2).Text)
 ebv.nhab = val(Me.Tb_car_eu(0).Text)
 ebv.tdilu = val(Me.Tb_car_eu(2).Text)
 ebv.ceau = val(Me.Tb_car_eu(1).Text)
 ebv.perti = val(Me.Tb_carep_rur(0).Text)
 ebv.vinf = val(Me.Tb_carep_rur(1).Text)
 ebv.ahorton = val(Me.Tb_carep_rur(2).Text)
 ebv.bhorton = val(Me.Tb_carep_rur(3).Text)
 ebv.trep = val(Me.Tb_carep_rur(4).Text)
End Sub
Sub c_ph()
eph.amontana = val(Me.tb_par_ep(0).Text)
eph.bmontana = val(Me.tb_par_ep(1).Text)
eph.lcrin = val(Me.Tb_par_eu(0).Text)
eph.ceau = val(Me.Tb_car_eu(1).Text)
eph.aeu = val(Me.Tb_par_eu(1).Text)
eph.beu = val(Me.Tb_par_eu(2).Text)
ehyd.Teta = 0.5
End Sub
Private Sub ini_pluie(ByVal visib As Boolean)
Me.Frm_ppr.Visible = visib

    Me.Lb_par_pl(0).Visible = visib
    Me.Tb_par_pl(0).Visible = visib
    Me.Lb_upar_pl(0).Visible = visib
    Me.Lb_par_pl(1).Visible = visib
    Me.Tb_par_pl(1).Visible = visib
    Me.Lb_upar_pl(1).Visible = visib
    Me.Lb_par_pl(2).Visible = visib
    Me.Tb_par_pl(2).Visible = visib
    Me.Lb_upar_pl(2).Visible = visib
    Me.Lb_par_pl(3).Visible = visib
    Me.Tb_par_pl(3).Visible = visib
    Me.Lb_upar_pl(3).Visible = visib
    Me.Lb_par_pl(4).Visible = visib
    Me.Tb_par_pl(4).Visible = visib
    Me.Lb_upar_pl(4).Visible = visib
    Me.Lb_par_pl(5).Visible = visib
    Me.Tb_par_pl(5).Visible = visib
    Me.Lb_upar_pl(5).Visible = visib
    Me.Cmd_calc.Visible = visib
End Sub
Private Sub ini_form()
    Me.Tb_car_ep(0).Text = ajout_zero(Trim(Str(ebv.surface)))
    Me.Tb_car_ep(3).Text = ajout_zero(Trim(Str(ebv.imper)))
    Me.Tb_car_ep(2).Text = ajout_zero(Trim(Str(ebv.phydr)))
    Me.Tb_car_ep(1).Text = ajout_zero(Trim(Str(ebv.lghydr)))
    Me.Tb_car_eu(0).Text = ajout_zero(Trim(Str(ebv.nhab)))
    Me.Tb_car_eu(2).Text = ajout_zero(Trim(Str(ebv.tdilu)))
    Me.tb_par_ep(0).Text = ajout_zero(Trim(Str(eph.amontana)))
    Me.tb_par_ep(1).Text = ajout_zero(Trim(Str(eph.bmontana)))
    Me.Tb_par_eu(0).Text = ajout_zero(Trim(Str(eph.lcrin)))
    Me.Tb_car_eu(1).Text = ajout_zero(Trim(Str(ebv.ceau)))
    Me.Tb_carep_rur(0).Text = ajout_zero(Trim(Str(ebv.perti)))
    Me.Tb_carep_rur(1).Text = ajout_zero(Trim(Str(ebv.vinf)))
    Me.Tb_carep_rur(2).Text = ajout_zero(Trim(Str(ebv.ahorton)))
    Me.Tb_carep_rur(3).Text = ajout_zero(Trim(Str(ebv.bhorton)))
    Me.Tb_carep_rur(4).Text = ajout_zero(Trim(Str(ebv.trep)))
    Me.Tb_par_eu(1).Text = ajout_zero(Trim(Str(eph.aeu)))
    Me.Tb_par_eu(2).Text = ajout_zero(Trim(Str(eph.beu)))
    Me.Tb_Qbrut.Text = rempl_virgule(Format(ebv.Qbrut * 1000#, "####0.0"))
    Me.Tb_debit(0).Text = rempl_virgule(Format(ebv.Qcor * 1000, "####0.0"))
    Me.Tb_debit(1).Text = rempl_virgule(Format(ebv.Qmr * 1000, "####0.0"))
    Me.Tb_debit(2).Text = rempl_virgule(Format(ebv.Qhydro * 1000, "####0.0"))
    Me.Tb_debit1(0).Text = rempl_virgule(Format(ebv.Qeu, "###0.0"))
    Me.Tb_debit1(2).Text = rempl_virgule(Format(ebv.Qecp, "###0.0"))
    Me.Tb_debit1(1).Text = rempl_virgule(Format(ebv.Qts, "###0.0"))
    Me.Tb_debit1(4).Text = rempl_virgule(Format(ebv.Qprin, "###0.0"))
    Me.Tb_debit1(3).Text = rempl_virgule(Format(ebv.Qrin, "###0.0"))
    Me.Tb_par_pl(4).Text = rempl_virgule(Format(ehyd.Teta, "###0,00"))
    Me.Tb_ruic.Visible = False
    Me.Lb_ruic.Visible = False
    Me.Lb_uruic.Visible = False
    opt_cli = False
    If ebv.type = "U" Then
        Call ini_urbain
    Else
        Call ini_rural
   End If
'    ehyd.dt = 0
'    ehyd.hm = 0#
'    ehyd.ht = 0#
'    ehyd.pas = ebv.pas
'    ehyd.teta = ebv.teta
'    ehyd.kdesbor = 0#
'    ehyd.vst = 0#
'    ehyd.vstock = 0#
'   ehyd.qfuite = ebv.qfuite
If ebv.surface <> 0 And ebv.lghydr <> 0 And ebv.phydr <> 0 And eph.amontana <> 0 _
    And eph.bmontana <> 0 Then
'    Me.Cmd_hydro.Visible = True
    Me.Tb_par_pl(1).Text = rempl_virgule(Format(ebv.tc, "#####0"))
    Me.Tb_par_pl(0).Text = rempl_virgule(Format(4 * ebv.tc, "#####0"))
    Me.Tb_par_pl(1).Text = rempl_virgule(Format(ehyd.DM, "#####0"))
    Me.Tb_par_pl(0).Text = rempl_virgule(Format(ehyd.dt, "#####0"))
    Me.Tb_txtqf.Text = rempl_virgule(Format(ehyd.qfuite, "#####0"))
    Me.Tb_par_pl(5) = rempl_virgule(Format(ehyd.pas, "#####0"))
    Me.Tb_par_pl(4) = rempl_virgule(Format(ehyd.Teta, "#####0.0"))
'    form_ouv = True
    Call calc_hyd
    Me.MnuPrint.Enabled = True
    SSTab1.Tab = 1
End If
Select Case ebv.Qchoisi
    Case Is = "CAQUOT"
        Me.Ob_Caquot.Value = True
        Me.Tb_debit(0).Enabled = True
        Me.Ob_Mr.Value = False
        Me.Tb_debit(1).Enabled = False
        Me.Ob_Mh.Value = False
        Me.Tb_debit(2).Enabled = False
    Case Is = "RATION"
        Me.Ob_Caquot.Value = False
        Me.Tb_debit(0).Enabled = False
        Me.Ob_Mr.Value = True
        Me.Tb_debit(1).Enabled = True
        Me.Ob_Mh.Value = False
        Me.Tb_debit(2).Enabled = False
    Case Is = "HYDROG"
        Me.Ob_Caquot.Value = False
        Me.Tb_debit(0).Enabled = False
        Me.Ob_Mr.Value = False
        Me.Tb_debit(1).Enabled = False
        Me.Ob_Mh.Value = True
        Me.Tb_debit(2).Enabled = True
End Select
   opt_cli = True
'If ebv.surface <> 0 And ebv.lghydr <> 0 And ebv.phydr <> 0 And eph.amontana <> 0 _
'    And eph.bmontana <> 0 Then
'    Me.Cmd_hydro.Visible = True
'End If

End Sub
Private Sub ini_urbain()
    Me.Opt_urbain.Value = True
    Me.Opt_rural.Value = False
    Me.Frm_ceu.Visible = True
    Me.Tb_car_eu(0).Visible = True
    Me.Tb_car_eu(1).Visible = True
    Me.Tb_car_eu(2).Visible = True
    Me.Lb_car_eu(0).Visible = True
    Me.Lb_car_eu(1).Visible = True
    Me.Lb_car_eu(2).Visible = True
    Me.Lb_ucar_eu(0).Visible = True
    Me.Lb_ucar_eu(1).Visible = True
    Me.Lb_ucar_eu(2).Visible = True
    Me.Frm_peu.Visible = True
    Me.Tb_par_eu(0).Visible = True
    Me.Tb_par_eu(1).Visible = True
    Me.Tb_par_eu(2).Visible = True
    Me.Lb_par_eu(0).Visible = True
    Me.Lb_par_eu(1).Visible = False
    Me.Lb_par_eu(2).Visible = True
    Me.Lb_upar_eu(0).Visible = True
    Me.Lb_upar_eu(1).Visible = True
    Me.Lb_upar_eu(2).Visible = True
    Me.Lb_aeu.Visible = True
    Me.Lb_beu.Visible = True
    Me.Lb_debit1(0).Visible = True
    Me.Lb_debit1(2).Visible = True
    Me.Lb_debit1(4).Visible = True
    Me.Lb_debit1(1).Visible = True
    Me.Lb_debit1(3).Visible = True
    Me.Tb_debit1(0).Visible = True
    Me.Tb_debit1(2).Visible = True
    Me.Tb_debit1(4).Visible = True
    Me.Tb_debit1(1).Visible = True
    Me.Tb_debit1(3).Visible = True
    Me.Lb_udebit1(0).Visible = True
    Me.Lb_udebit1(2).Visible = True
    Me.Lb_udebit1(4).Visible = True
    Me.Lb_udebit1(1).Visible = True
    Me.Lb_udebit1(3).Visible = True
    Me.Tb_ruic.Visible = False
    Me.Lb_ruic.Visible = False
    Me.Lb_uruic.Visible = False
    Me.Tb_ruic.Text = "0.0"
    
'    Me.Frm_debit.Left = 120
    Me.Frm_cbr.Visible = False
    Me.Tb_carep_rur(0).Visible = False
    Me.Tb_carep_rur(1).Visible = False
    Me.Tb_carep_rur(2).Visible = False
    Me.Tb_carep_rur(3).Visible = False
    Me.Tb_carep_rur(4).Visible = False
    Me.Lb_carep_rur(0).Visible = False
    Me.Lb_carep_rur(1).Visible = False
    Me.Lb_carep_rur(2).Visible = False
    Me.Lb_carep_rur(3).Visible = False
    Me.Lb_carep_rur(4).Visible = False
    Me.Lb_ucarep_rur(0).Visible = False
    Me.Lb_ucarep_rur(1).Visible = False
    Me.Lb_ucarep_rur(2).Visible = False
    Me.Lb_ucarep_rur(3).Visible = False
    Me.Lb_ucarep_rur(4).Visible = False
    Me.Lb_aHorton.Visible = False
    Me.Lb_bHorton.Visible = False
'    Me.Frm_ceu.Top = 2760
    Me.Ob_Mr.Enabled = False
    Me.Ob_Caquot.Enabled = False
    Me.Ob_Mh.Enabled = False
    Me.Ob_Mr.Value = False
    Me.Ob_Caquot.Value = False
    Me.Ob_Mh.Value = False
    If ebv.Qcor > 0 Then
        Me.Ob_Caquot.Enabled = True
    End If
    If ebv.Qmr > 0 Then
        Me.Ob_Mr.Enabled = True
    End If
    If ebv.Qhydro > 0 Then
        Me.Ob_Mh.Enabled = True
    End If
End Sub
Private Sub ini_rural()
    Me.Opt_urbain.Value = False
    Me.Opt_rural.Value = True
    Me.Frm_ceu.Visible = False
    Me.Tb_car_eu(0).Visible = False
    Me.Tb_car_eu(1).Visible = False
    Me.Tb_car_eu(2).Visible = False
    Me.Lb_car_eu(0).Visible = False
    Me.Lb_car_eu(1).Visible = False
    Me.Lb_car_eu(2).Visible = False
    Me.Lb_ucar_eu(0).Visible = False
    Me.Lb_ucar_eu(1).Visible = False
    Me.Lb_ucar_eu(2).Visible = False
    Me.Frm_peu.Visible = False
    Me.Tb_par_eu(0).Visible = False
    Me.Tb_par_eu(1).Visible = False
    Me.Tb_par_eu(2).Visible = False
    Me.Lb_par_eu(0).Visible = False
    Me.Lb_par_eu(1).Visible = False
    Me.Lb_par_eu(2).Visible = False
    Me.Lb_upar_eu(0).Visible = False
    Me.Lb_upar_eu(1).Visible = False
    Me.Lb_upar_eu(2).Visible = False
    Me.Lb_aeu.Visible = False
    Me.Lb_beu.Visible = False
    Me.Lb_debit1(0).Visible = False
    Me.Lb_debit1(2).Visible = False
    Me.Lb_debit1(4).Visible = False
    Me.Lb_debit1(1).Visible = False
    Me.Lb_debit1(3).Visible = False
    Me.Tb_debit1(0).Visible = False
    Me.Tb_debit1(2).Visible = False
    Me.Tb_debit1(4).Visible = False
    Me.Tb_debit1(1).Visible = False
    Me.Tb_debit1(3).Visible = False
    Me.Lb_udebit1(0).Visible = False
    Me.Lb_udebit1(2).Visible = False
    Me.Lb_udebit1(4).Visible = False
    Me.Lb_udebit1(1).Visible = False
    Me.Lb_udebit1(3).Visible = False
    Me.Tb_ruic.Visible = True
    Me.Lb_ruic.Visible = True
    Me.Lb_uruic.Visible = True
    Me.Tb_ruic.Text = "0.0"
'    Me.Frm_debit.Left = 3700
    Me.Frm_cbr.Visible = True
    Me.Tb_carep_rur(0).Visible = True
    Me.Tb_carep_rur(1).Visible = True
    Me.Tb_carep_rur(2).Visible = True
    Me.Tb_carep_rur(3).Visible = True
    Me.Tb_carep_rur(4).Visible = True
    Me.Lb_carep_rur(0).Visible = True
    Me.Lb_carep_rur(1).Visible = True
    Me.Lb_carep_rur(2).Visible = False
    Me.Lb_carep_rur(3).Visible = True
    Me.Lb_carep_rur(4).Visible = True
    Me.Lb_ucarep_rur(0).Visible = True
    Me.Lb_ucarep_rur(1).Visible = True
    Me.Lb_ucarep_rur(2).Visible = True
    Me.Lb_ucarep_rur(3).Visible = True
    Me.Lb_ucarep_rur(4).Visible = True
    Me.Lb_aHorton.Visible = True
    Me.Lb_bHorton.Visible = True
'    Me.Frm_cbr.Top = 2520
    Me.Ob_Mr.Enabled = False
    Me.Ob_Caquot.Enabled = False
    Me.Ob_Mh.Enabled = False
    Me.Ob_Mr.Value = False
    Me.Ob_Caquot.Value = False
    Me.Ob_Mh.Value = False
    If ebv.Qcor > 0 Then
        Me.Ob_Caquot.Enabled = True
    End If
    If ebv.Qmr > 0 Then
        Me.Ob_Mr.Enabled = True
    End If
    If ebv.Qhydro > 0 Then
        Me.Ob_Mh.Enabled = True
    End If
End Sub
Private Sub Cb_bassin_Change()
    Cb_bassin.Text = ba_texte
End Sub

Private Sub Cb_bassin_KeyDown(KeyCode As Integer, Shift As Integer)
    ba_texte = Cb_bassin.Text
    Cb_bassin.Text = ba_texte

End Sub

Private Sub Cb_bassin_KeyPress(KeyAscii As Integer)
    ba_texte = Cb_bassin.Text
End Sub

Private Sub Cmd_calc_Click()
    Call calc_hyd

End Sub


Private Sub Cmd_retour_Click()
Unload Me
End Sub


Public Sub Init_ss_commentaire()
    owner.affich_com Me.Name, ""
    owner.affich_aide Me.Name, "Calcul de d�bit de bassin versant"

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim reponse As Integer
    If do_bv Or sto_bv Or ret_bv Then
        owner.fdessin.Image1.Visible = False
        owner.fdessin.UC_graphique1.Visible = True
        If do_bv Then
            owner.fobjet.Enabled = True
            owner.fdessin.UC_graphique1.graphique_clear
            owner.fdessin.UC_graphique1.init_title
            owner.fdessin.UC_graphique1.init_titleh ""
            owner.fdessin.UC_graphique1.init_titleb ""
            If Trim(Me.Tb_titre.Text) <> "" Then
                reponse = MsgBox("", 4, "Transfert des donn�es vers le d�versoir d'orage.")
                If reponse = 6 Then ' 6=oui,7=non
                    Call owner.fobjet.ini_debit(Me.Tb_titre.Text)
                Else
                    owner.fdessin.UC_graphique1.init_fond dess_anc
'                   owner.fdessin.UC_graphiqueB.Visible = True
                    owner.fdessin.Image3.Visible = True
                    owner.fdessin.UC_graphique1.Visible = False
                    owner.fdessin.UC_graphique2.Visible = False
                End If
            Else
                owner.fdessin.UC_graphique1.init_fond dess_anc
'                owner.fdessin.UC_graphiqueB.Visible = True
                owner.fdessin.Image3.Visible = True
                owner.fdessin.UC_graphique1.Visible = False
                owner.fdessin.UC_graphique2.Visible = False
            End If
            do_bv = False
            owner.affich_aide owner.fobjet.Name, "DO:Bassin versant"
        End If
        If sto_bv Then
            owner.fobjet.Enabled = True
            owner.fdessin.UC_graphique1.graphique_clear
            owner.fdessin.UC_graphique1.init_title
            owner.fdessin.UC_graphique1.init_titleh ""
            owner.fdessin.UC_graphique1.init_titleb ""
            If Trim(Me.Tb_titre.Text) <> "" Then
                reponse = MsgBox("", 4, "Transfert des donn�es vers le bassin de stockage.")
                If reponse = 6 Then ' 6=oui,7=non
                    Call owner.fobjet.ini_debit(Me.Tb_titre.Text)
                Else
                    owner.fdessin.UC_graphique1.init_fond dess_anc
                End If
            Else
                owner.fdessin.UC_graphique1.init_fond dess_anc
            End If
            sto_bv = False
            owner.affich_aide owner.fobjet.Name, "Dimensionnement d'un bassin de stockage"
        End If
        If ret_bv Then
            owner.fobjet.Enabled = True
            owner.fdessin.UC_graphique1.graphique_clear
            owner.fdessin.UC_graphique1.init_title
            owner.fdessin.UC_graphique1.init_titleh ""
            owner.fdessin.UC_graphique1.init_titleb ""
            If Trim(Me.Tb_titre.Text) <> "" Then
                reponse = MsgBox("", 4, "Transfert des donn�es vers le bassin de r�tention.")
                If reponse = 6 Then ' 6=oui,7=non
                    Call owner.fobjet.ini_debit(Me.Tb_titre.Text)
                Else
                    owner.fdessin.UC_graphique1.init_fond dess_anc
                End If
            Else
                owner.fdessin.UC_graphique1.init_fond dess_anc
            End If
            ret_bv = False
            owner.affich_aide owner.fobjet.Name, "Dimensionnement d'un bassin de r�tention"
        End If
    Else
        Unload Frm_desprint
        Unload owner.fdessin
        owner.recharge_commentaire
        
    End If
        If Not owner.fbassin Is Nothing Then
            Set owner.fbassin = Nothing
        End If
End Sub

Private Sub image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Debug.Print X, Y
End Sub



Private Sub Frm_cep_Click()
owner.affich_aide Me.Name, "Caract�ristiques d'un BV"

End Sub

Private Sub Frm_ceu_Click()
Dim mes As String
mes = "M�thodes d'�valuation des d�bits de temps sec"
owner.affich_aide Me.Name, mes

End Sub

Private Sub Frm_debit_Click()
Dim mes As String
mes = "D�bits caract�ristiques"
owner.affich_aide Me.Name, mes

End Sub



Private Sub Lb_car_ep_Click(Index As Integer)
Dim mes As String
mes = "Caract�ristiques d'un BV"
Select Case Index
 Case Is = 3
   mes = "Tableaux coefficients de ruissellement"
End Select
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_car_eu_Click(Index As Integer)
Dim mes As String
mes = "M�thodes d'�valuation des d�bits de temps sec"
owner.affich_aide Me.Name, mes

End Sub

Private Sub Lb_debit_Click(Index As Integer)
Dim mes As String
Select Case Index
 Case Is = 0
   mes = "M�thode superficielle de Caquot"
 Case Is = 1
   mes = "M�thode Rationnelle "
 Case Is = 2
   mes = "M�thode de l'hydrogramme"
End Select
owner.affich_aide Me.Name, mes

End Sub



Private Sub m_quitter_Click()
    Unload owner
End Sub

Private Sub mnufichier_Click()
    If ouv_sauve Or save_fich Then ' Or (Not ouv_sauve And Not save_fich) Then
        Me.mnusave.Enabled = True
        Me.mnusaves.Enabled = True
        Me.mnusuppr.Enabled = True
'        Me.Mnuprint.Enabled = True
    Else
        Me.mnusave.Enabled = False
        Me.mnusaves.Enabled = False
        Me.mnusuppr.Enabled = False
        Me.MnuPrint.Enabled = False
   End If
End Sub

Private Sub mnuinfo_Click()
    Frm_saisie.Show 1
End Sub

Private Sub mnunouv_Click()
Dim reponse As Integer
If ProtectCheck(2) <> 0 Then End
If ouv_sauve Then 'And save_fich Then
    reponse = MsgBox("Le bassin n'a pas �t� enregistr�" + Chr(10) _
        + "Voulez vous le sauvegarder?", 3, "Sauvegarde du bassin")
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
If ProtectCheck(2) <> 0 Then End
fich_lect = nom_fich
If nom_fich_edit <> "" Then
    nom = "Etude " + nom_fich_edit
Else
    nom = " Nouvelle �tude "
End If
If ouv_sauve Then 'And save_fich Then
    reponse = MsgBox("Le bassin n'a pas �t� enregistr�" + Chr(10) _
        + "Voulez vous le sauvegarder?", 3, "Sauvegarde d'un bassin versant")
    Select Case reponse
        Case Is = 6  ' 6=oui,7=non,2=annuler
            Call mnusave_Click
'                Cb_bassin.Visible = True
            frmf.Label1.Caption = "Recherche d'un bassin versant "
            frmf.Caption = nom
            frmf.Show 1
            If frmf.nomfich <> "" Then
                Me.nom_ouvrage = frmf.nomfich
                Call Me.Cb_bassin_Click
            End If
        Case Is = 7
'                Cb_bassin.Visible = True
            frmf.Label1.Caption = "Recherche d'un bassin versant "
            frmf.Caption = nom
            frmf.Show 1
            If frmf.nomfich <> "" Then
                Me.nom_ouvrage = frmf.nomfich
                Call Me.Cb_bassin_Click
            End If
    End Select
Else
'    Cb_bassin.Visible = True
    frmf.Label1.Caption = "Recherche d'un bassin versant "
    frmf.Caption = nom
    frmf.Show 1
    If frmf.nomfich <> "" Then
        Me.nom_ouvrage = frmf.nomfich
        Call Me.Cb_bassin_Click
    End If
End If
Set frmf = Nothing
End Sub

Private Sub mnusave_Click()
If ProtectCheck(2) <> 0 Then End
    If Trim(Tb_titre.Text) <> "" And fich_lect = nom_fich Then
        Call save(False)
    Else
        Call mnusaves_Click
    End If

End Sub
Private Sub mnusave0_Click()
Dim za As st_save
Dim i As Integer, isave As Integer
Dim reponse As Integer

If Trim(Tb_titre.Text) <> "" Then
    Call funlockb

    ebv.nom = Mid(Trim(Me.Tb_titre.Text), 1, 10)
    ebv.qfuite = ehyd.qfuite
    ebv.Teta = ehyd.Teta
    ebv.pas = ehyd.pas
   lhFicDbf = FreeFile
'   Debug.Print Len(esave)
    Open nom_fich For Random Access Read Write As #lhFicDbf Len = Len(za)
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
            If Not ouv_sauve Then
                Call Cmd_resu_Click
            End If
            esave.nom = Tb_titre.Text
            esave.bv = ebv
            esave.hydro = eph
             esave.hydro1 = ehyd
           Put #lhFicDbf, isave, esave
    End If
        Close #lhFicDbf
        
        Call flockb(nom_fich)
        
        ouv_sauve = False
        Call lect_fich
        Cb_bassin.Text = Trim(Tb_titre.Text)
'        Me.Cmd_del.Visible = True
Else
    reponse = MsgBox("Le nom du bassin n'est pas renseign�.", , "Sauvegarde d'un bassin")
End If


End Sub

Private Sub Opt_rural_GotFocus()
owner.affich_aide Me.Name, "Type de bassins"

End Sub

Private Sub Opt_urbain_GotFocus()
owner.affich_aide Me.Name, "Type de bassins"

End Sub

Private Sub Tb_car_Change(Index As Integer)
   Call reini_form(1)
End Sub

Private Sub Tb_car_ep_change(Index As Integer)
Dim nom As String

If bKP Then
        Select Case Index
            Case Is = 0
                 nom = verif_cart0(Tb_car_ep(Index).Text, "Saisie Surface", "R")
            Case Is = 1
                 nom = verif_cart0(Tb_car_ep(Index).Text, "Saisie Longueur", "R")
            Case Is = 2
                 nom = verif_cart0(Tb_car_ep(Index).Text, "Saisie Pente", "I")
            Case Is = 3
                 nom = verif_cart0(Tb_car_ep(Index).Text, "Saisie Coef. de ruissellement", "I")
        End Select
  If nom = "" Then
    Tb_car_ep(Index).Text = sval_champ
    Tb_car_ep(Index).SelStart = iSels
    Tb_car_ep(Index).SelLength = iSell
  End If
End If
'****

    Call reini_form(1)
     sval_champ = ""
    bKP = False

End Sub

Private Sub mnuprint_Click()
Dim pict1 As New StdPicture
Dim i As Integer
Dim chaine As String
If ProtectCheck(2) <> 0 Then End
chaine = " rural"
If Me.Opt_urbain Then
    chaine = " urbain"
End If

FrmPrint.Type1 = "versant"
FrmPrint.nomobjet = Trim(Tb_titre.Text)
FrmPrint.titre1 = "FICHE HYDRAULIQUE BASSIN VERSANT" + chaine
FrmPrint.sstitre1 = "Caract�ristiques"
FrmPrint.ssTitre2 = "Param�tres"
FrmPrint.ssTitre3 = ""
Frm_imp.Type1 = "versant"
Frm_imp.nomobjet = Trim(Tb_titre.Text)
Frm_imp.titre1 = "FICHE HYDRAULIQUE BASSIN VERSANT" + chaine
Frm_imp.sstitre1 = "Caract�ristiques"
Frm_imp.ssTitre2 = "Param�tres"
Frm_imp.ssTitre3 = ""
If Me.Opt_rural Then
    ReDim list_don1(Tb_carep_rur.Count - 1, 6)
    For i = 0 To Tb_car_ep.Count - 1
        list_don1(i, 1) = Lb_car_ep(i).Caption
        list_don1(i, 2) = Tb_car_ep(i).Text
        list_don1(i, 3) = Lb_ucar_ep(i).Caption
    Next
    For i = 0 To Tb_carep_rur.Count - 1
        list_don1(i, 4) = Lb_carep_rur(i).Caption
        If i = 2 Then
            list_don1(i, 4) = Lb_carep_rur(i).Caption + " a"
        End If
        If i = 3 Then
            list_don1(i, 4) = Lb_carep_rur(i).Caption + " b"
        End If
        list_don1(i, 5) = Tb_carep_rur(i).Text
        list_don1(i, 6) = Lb_ucarep_rur(i).Caption
    Next
    ReDim list_don2(Tb_par_pl.Count - 1, 6)
        list_don2(0, 1) = Frm_pep.Caption
        list_don2(0, 2) = ""
        list_don2(0, 3) = ""
    For i = 0 To tb_par_ep.Count - 1
        If i = 0 Then
            list_don2(i + 1, 1) = Lb_par_ep(i).Caption + " a"
        End If
        If i = 1 Then
            list_don2(i + 1, 1) = Lb_par_ep(i).Caption + " b"
        End If
        list_don2(i + 1, 2) = tb_par_ep(i).Text
        list_don2(i + 1, 3) = Lb_upar_ep(i).Caption
    Next
    For i = 0 To Tb_par_pl.Count - 1
        list_don2(i, 4) = Lb_par_pl(i).Caption
        list_don2(i, 5) = Tb_par_pl(i).Text
        list_don2(i, 6) = Lb_upar_pl(i).Caption
    Next
    ReDim list_resu1(Tb_debit.Count, 6)
        list_resu1(0, 1) = Frm_debit.Caption
        list_resu1(0, 2) = ""
        list_resu1(0, 3) = ""
    For i = 0 To Tb_debit.Count - 1
        list_resu1(i + 1, 1) = Lb_debit(i).Caption
        list_resu1(i + 1, 2) = Tb_debit(i).Text
        list_resu1(i + 1, 3) = Lb_udebit(i).Caption
    Next
        list_resu1(0, 4) = Lb_debit1(5).Caption
        list_resu1(0, 5) = Tb_debit1(5).Text
        list_resu1(0, 6) = Lb_udebit1(5).Caption
        list_resu1(1, 4) = Lb_debit1(6).Caption
        list_resu1(1, 5) = Tb_debit1(6).Text
        list_resu1(1, 6) = Lb_udebit1(6).Caption
Else
    ReDim list_don1(Tb_car_ep.Count - 1, 6)
    For i = 0 To Tb_car_ep.Count - 1
        list_don1(i, 1) = Lb_car_ep(i).Caption
        list_don1(i, 2) = Tb_car_ep(i).Text
        list_don1(i, 3) = Lb_ucar_ep(i).Caption
    Next
    For i = 0 To Tb_car_eu.Count - 1
        list_don1(i, 4) = Lb_car_eu(i).Caption
        list_don1(i, 5) = Tb_car_eu(i).Text
        list_don1(i, 6) = Lb_ucar_eu(i).Caption
    Next
    ReDim list_don2(Tb_par_pl.Count, 6)
        list_don2(0, 1) = Frm_pep.Caption
        list_don2(0, 2) = ""
        list_don2(0, 3) = ""
    For i = 0 To tb_par_ep.Count - 1
        If i = 0 Then
            list_don2(i + 1, 1) = Lb_par_ep(i).Caption + " a"
        End If
        If i = 1 Then
            list_don2(i + 1, 1) = Lb_par_ep(i).Caption + " b"
        End If
        list_don2(i + 1, 2) = tb_par_ep(i).Text
        list_don2(i + 1, 3) = Lb_upar_ep(i).Caption
    Next
    j = i + 1
    list_don2(j, 1) = Frm_peu.Caption
    list_don2(j, 2) = ""
    list_don2(j, 3) = ""
    j = j + 1
     For i = 0 To Tb_par_eu.Count - 1
        list_don2(i + j, 1) = Lb_par_eu(i).Caption
        If i = 1 Then
            list_don2(i + j, 1) = Lb_par_eu(i).Caption + " a"
        End If
        If i = 2 Then
            list_don2(i + j, 1) = Lb_par_eu(i).Caption + " b"
        End If
        list_don2(i + j, 2) = Tb_par_eu(i).Text
        list_don2(i + j, 3) = Lb_upar_eu(i).Caption
    Next
    For i = 0 To Tb_par_pl.Count - 1
        list_don2(i, 4) = Lb_par_pl(i).Caption
        list_don2(i, 5) = Tb_par_pl(i).Text
        list_don2(i, 6) = Lb_upar_pl(i).Caption
    Next
    ReDim list_resu1(Tb_debit1.Count - 1, 6)
        list_resu1(0, 1) = Frm_debit.Caption
        list_resu1(0, 2) = ""
        list_resu1(0, 3) = ""
    For i = 0 To Tb_debit.Count - 1
        list_resu1(i + 1, 1) = Lb_debit(i).Caption
        list_resu1(i + 1, 2) = Tb_debit(i).Text
        list_resu1(i + 1, 3) = Lb_udebit(i).Caption
    Next
    For i = 0 To Tb_debit1.Count - 1
        list_resu1(i, 4) = Lb_debit1(i).Caption
        list_resu1(i, 5) = Tb_debit1(i).Text
        list_resu1(i, 6) = Lb_udebit1(i).Caption
    Next
End If
    FrmPrint.des1_titrh = "HYETOGRAMME DE LA PLUIE"
    FrmPrint.des1_titrb = "HYDROGRAMME DE RUISSELEMENT"
    Frm_imp.des1_titrh = "HYETOGRAMME DE LA PLUIE"
    Frm_imp.des1_titrb = "HYDROGRAMME DE RUISSELEMENT"

Set pict1 = Frm_desprint.UC_graphique1.lire_pict1()
FrmPrint.paint_picture pict1
SavePicture pict1, chemin_app + "dess.bmp"
Frm_imp.Show 1
End Sub
Public Function lect_list(ByVal nom As String) As Variant
Select Case nom
Case Is = "list_don1"
    lect_list = list_don1
Case Is = "list_don2"
    lect_list = list_don2
Case Is = "list_int1"
    lect_list = list_int1
Case Is = "list_resu1"
    lect_list = list_resu1
End Select
End Function

Private Sub MnuQuit_Click()
    Unload Me
End Sub
Public Sub Cb_bassin_Click()
  Call rec_bassin_versant
End Sub
Public Function recup_mnuprint()
    recup_mnuprint = Me.MnuPrint.Enabled
End Function
Public Sub rec_bassin_versant()
Dim za As st_save
Dim za1 As st_save1
Dim ok As Boolean
    Call funlockb

'    Cb_bassin.Visible = False
'    For i = 0 To Cb_bassin.ListCount - 1
'        If Trim(Cb_bassin.list(i)) = Trim(nom_ouvrage) Then
'            ba_texte = Cb_bassin.list(i)
'            Cb_bassin.Text = Cb_bassin.list(i)
'        End If
'    Next
    Me.Tb_titre.Text = Trim(nom_ouvrage)
    ba_texte = Trim(nom_ouvrage)
    Cb_bassin.Text = Trim(nom_ouvrage)
    lhFicDbf = FreeFile
    On Error GoTo test_Error
If Trim(nom_ouvrage) = "" Then
    ok = True
Else
    ok = False
    Open fich_lect For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
    If Not EOF(lhFicDbf) Then
        If Trim(za1.type) = nom_type Then
            za = za1.stsave
            If Trim(za.nom) = Trim(Cb_bassin.Text) Then
                Tb_titre = Trim(za.nom)
                Me.Caption = fen_titre + " : " + Tb_titre.Text
                ok = True
                ebv = za.bv
'               Debug.Print ebv.lghydr
                eph = za.hydro
                ehyd = za.hydro1
                bv_charge = True
                Call ini_form
                Call ini_pluie(True)
'               Me.Cmd_del.Visible = True
                ouv_sauve = False
                save_fich = True
                If fich_lect <> nom_fich Then
                    ouv_sauve = True
                End If
            End If
        End If
    End If

Loop
Close #lhFicDbf
If fich_lect <> nom_fich Then
    Kill fich_lect
End If
End If
Call flockb(nom_fich)
If Not ok Then
    MsgBox "Le bassin n'existe pas dans l'�tude courante", vbExclamation, "S�lection d'un bassin"
End If
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + fich_lect + " est d�j� en cours d'utilisation.")
    End If

Call flockb(nom_fich)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim reponse As Integer
If ouv_sauve Then 'And save_fich Then
    reponse = MsgBox("Le bassin n'a pas �t� enregistr�" + Chr(10) _
        + "Voulez vous le sauvegarder?", 3, "Sauvegarde du bassin")
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


Private Sub mnusaves_Click()
If ProtectCheck(2) <> 0 Then End
    If fich_lect = nom_fich Or Trim(Tb_titre.Text) = "" Or fich_lect = "" Then
        Frm_titre.Label2.Caption = "Sauvegarde d'un bassin versant "
        Frm_titre.Label3.Caption = ""
    Else
         Frm_titre.Label2.Caption = "Sauvegarde du bassin versant " & Me.Tb_titre.Text
         Frm_titre.Label3.Caption = " de l'�tude " & fich_lect
   
    End If
    Frm_titre.Caption = "Etude " + nom_fich_edit
    Frm_titre.Label1.Caption = "Nom du bassin (30car. maxi)"
    Frm_titre.Text1.Text = Me.Tb_titre.Text
    Frm_titre.Show 1
End Sub
Public Sub save(ByVal bsous As Boolean)
Dim za As st_save
Dim za1 As st_save1
Dim i As Integer, isave As Integer
Dim reponse As Integer

If Trim(Tb_titre.Text) <> "" Then
    Call funlockb

   ebv.nom = Mid(Trim(Me.Tb_titre.Text), 1, 10)
    ebv.qfuite = ehyd.qfuite
    ebv.Teta = ehyd.Teta
    ebv.pas = ehyd.pas
   lhFicDbf = FreeFile
'   Debug.Print Len(esave)
    On Error GoTo test_Error
    Open nom_fich For Random Access Read Write Lock Read Write As #lhFicDbf Len = Len(za1)
    i = 0
    isave = 0
    Do While Not EOF(lhFicDbf)
        Get #lhFicDbf, , za1
        If Not EOF(lhFicDbf) Then
            i = i + 1
            If Trim(za1.type) = nom_type Then
                za = za1.stsave
                If Trim(za.nom) = Trim(Tb_titre.Text) Then
                    isave = i
                End If
            End If
       End If
    Loop
    If isave > 0 Then
        If bsous Then
           reponse = MsgBox("Le nom est d�j� utilis�. Le remplacer?", 4, "Sauvegarde d'un bassin versant")
           Else
           reponse = 6
        End If
        If reponse = 6 Then
            If ouv_sauve Then
                Call Cmd_resu_Click
            End If
            esave.nom = Tb_titre.Text
            esave.bv = ebv
            esave.hydro = eph
            esave.hydro1 = ehyd
            za1.type = "versant"
            za1.stsave = esave
            Put #lhFicDbf, isave, za1
            ouv_sauve = False
            save_fich = True
            fich_lect = nom_fich
        Else
            Unload Frm_titre
            Call mnusaves_Click
        End If
    Else
        If ouv_sauve Then
            Call Cmd_resu_Click
        End If
        esave.nom = Tb_titre.Text
        esave.bv = ebv
        esave.hydro = eph
        esave.hydro1 = ehyd
        za1.type = "versant"
        za1.stsave = esave
        FileLength = LOF(lhFicDbf) / Len(za1) + 1
        Put #lhFicDbf, FileLength, za1
        ouv_sauve = False
        save_fich = True
        fich_lect = nom_fich
    End If
        Close #lhFicDbf
        Call flockb(nom_fich)
        Call lect_fich
        ba_texte = Trim(Tb_titre.Text)
        Cb_bassin.Text = Trim(Tb_titre.Text)
'        Me.Cmd_del.Visible = True
Else
    reponse = MsgBox("Le nom du bassin n'est pas renseign�.", , "Sauvegarde d'un bassin versant")
End If

Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + nom_fich + " est d�j� en cours d'utilisation.")
    End If

Call flockb(nom_fich)
End Sub
Private Sub Form_Load()
    okg = True
    Cmd_retour.Visible = False
    nom_ouvrage = ""
'    nom_fich = chemin_app + "bassin.bin"
'    nom_fich = chemin_app + "etude.bin"
    nom_type = "versant"
    Set owner = MDIFrm_menu.rec_owner
    Call retaille
    fen_titre = Me.Caption
    ouv_sauve = False
    save_fich = False
    If Dir(nom_fich) <> "" Then
        Call lect_fich
    End If
    If do_bv Or sto_bv Or ret_bv Then
        Me.Ob_Caquot.Visible = True
        Me.Ob_Mh.Visible = True
        Me.Ob_Mr.Visible = True
    Else
        Me.Ob_Caquot.Visible = False
        Me.Ob_Mh.Visible = False
        Me.Ob_Mr.Visible = False
    End If
    Cb_bassin.Visible = False
    Frm_desprint.Show
    Frm_desprint.Visible = False
'    owner.affich_aide Me.Name, "Calcul de d�bit de bassin versant"
    Call debut
End Sub
Private Sub debut0()
    Cb_bassin.Text = ""
    Tb_titre.Text = ""
    Me.Caption = fen_titre
'    ouv_sauve = True 'False
    Call debut
End Sub

Private Sub init_l_tab()
Dim l0() As Variant, l1() As Variant, l2() As Variant
l0 = Array(0)
l1 = Array(0, "TB_car_ep", "TB_car_eu", "TB_carep_rur")
l2 = Array(0, "TB_par_ep", "TB_par_eu", "TB_par_pl")
ReDim list_tb(0 To UBound(l0), 0 To UBound(l1), 0 To UBound(l2))
list_tb = Array(l0, l1, l2)
End Sub
Private Sub debut()
Dim nom As String
Dim i As Integer
    bKP = False
    sval_champ = ""
Call init_l_tab

   nom = chemin_app + "bv.bmp"
    bv_charge = False
    Me.SSTab1.Tab = 0
'    Me.SSTab1.TabVisible(2) = False
    owner.fdessin.Image2.Visible = False
    owner.fdessin.UC_graphique2.Visible = False
    owner.fdessin.UC_graphiqueB.Visible = False
    owner.fdessin.Image1.Visible = True
    owner.fdessin.Image3.Visible = False
    owner.fdessin.Image1.Picture = LoadPicture(nom)
    owner.fdessin.UC_graphique1.Visible = False
    owner.fdessin.UC_graphique1.graphique_clear
'    owner.fdessin.UC_graphique1.Top = 0
'    owner.fdessin.UC_graphique1.Left = 1560
'    owner.fdessin.UC_graphique1.Height = 4210
'    owner.fdessin.UC_graphique1.Width = 7800
    Me.Tb_titre.Text = ""
'LblTitre.Caption = "BASSIN VERSANT TEST"
'    ehyd.qfuite = 0
'    Me.Tb_car_ep(0).Text = "20"
'    Me.Tb_car_ep(3).Text = "30"
'    Me.Tb_car_ep(2).Text = "50"
'    Me.Tb_car_ep(1).Text = "800"
    Me.Tb_car_ep(0).Text = "0"
    Me.Tb_car_ep(3).Text = "0"
    Me.Tb_car_ep(2).Text = "0"
    Me.Tb_car_ep(1).Text = "0"
    Me.Tb_car_eu(0).Text = "0"
'    Me.Tb_car_eu(2).Text = "100"
'    Me.tb_par_ep(0).Text = "5.9"
'    Me.tb_par_ep(1).Text = "0.59"
'    Me.Tb_par_eu(0).Text = "15"
'    Me.Tb_car_eu(1).Text = "150"
    Me.Tb_car_eu(2).Text = "0"
    Me.tb_par_ep(0).Text = "0.0"
    Me.tb_par_ep(1).Text = "0.0"
    Me.Tb_par_eu(0).Text = "0"
    Me.Tb_car_eu(1).Text = "0"
    Me.Tb_carep_rur(0).Text = "0"
    Me.Tb_carep_rur(1).Text = "0"
    Me.Tb_carep_rur(2).Text = "0.0"
    Me.Tb_carep_rur(3).Text = "0"
    Me.Tb_carep_rur(4).Text = "0"
'    Me.Tb_par_eu(1).Text = "1.5"
'    Me.Tb_par_eu(2).Text = "2.5"
    Me.Tb_par_eu(1).Text = "0.0"
    Me.Tb_par_eu(2).Text = "0.0"
    Me.Tb_Qbrut.Text = "0.0"
    Me.Tb_debit(0).Text = "0.0"
    Me.Tb_debit(1).Text = "0.0"
    Me.Tb_debit(2).Text = "0.0"
    Me.Tb_debit1(0).Text = "0.0"
    Me.Tb_debit1(2).Text = "0.0"
    Me.Tb_debit1(1).Text = "0.0"
    Me.Tb_debit1(4).Text = "0.0"
    Me.Tb_debit1(3).Text = "0.0"
    Me.Tb_ruic.Text = "0.0"
    tc = 0
    DM = 0
    dt = 0
    opt_cli = False
    Me.Ob_Mr.Enabled = False
    Me.Ob_Caquot.Enabled = False
    Me.Ob_Mh.Enabled = False
'    Me.Cmd_del.Visible = False
    Me.Opt_rural.Value = False
    Me.Opt_urbain.Value = True
    Me.Frm_ceu.Visible = True
    For i = 1 To Tb_car_eu.Count
    Me.Tb_car_eu(i - 1).Visible = True
    Next
    Me.Frm_peu.Visible = True
    For i = 1 To Tb_par_eu.Count
    Me.Tb_par_eu(i - 1).Visible = True
    Next
    Me.Lb_par_eu(1).Visible = False
    Me.Frm_cbr.Visible = False
    For i = 1 To Tb_carep_rur.Count
    Me.Tb_carep_rur(i - 1).Visible = False
    Next
    
   Call ini_pluie(False)
    Call ini_schema
    ini_bv
    opt_cli = True
    ouv_sauve = False
    save_fich = False
    fich_lect = ""
End Sub
Private Sub lect_fich()
Dim za As st_save
Dim za1 As st_save1
 Call funlockb
   lhFicDbf = FreeFile
    Cb_bassin.Clear
    On Error GoTo test_Error
    Open nom_fich For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
'   Open nom For Binary Access Read As #lhFicDbf
Do While Not EOF(lhFicDbf)
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
    Get #lhFicDbf, , za1
   If Not EOF(lhFicDbf) Then
        If Trim(za1.type) = nom_type Then
            za = za1.stsave
            Cb_bassin.AddItem (Trim(za.nom))
        End If
   End If
Loop
Close #lhFicDbf
ba_texte = Cb_bassin.list(0)
Cb_bassin.Text = Cb_bassin.list(0)
Cb_bassin.Refresh

Call flockb(nom_fich)
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + nom_fich + " est d�j� en cours d'utilisation.")
    End If

Call flockb(nom_fich)

End Sub
Private Sub reini_form(ByVal i As Integer)
    Select Case i
         Case Is = 1
             Me.Tb_Qbrut.Text = "0.0"
             Me.Tb_debit(0).Text = "0.0"
             Me.Tb_debit(1).Text = "0.0"
             Me.Tb_debit(2).Text = "0.0"
             Me.Tb_debit1(0).Text = "0.0"
             Me.Tb_debit1(2).Text = "0.0"
             Me.Tb_debit1(1).Text = "0.0"
             Me.Tb_debit1(4).Text = "0.0"
             Me.Tb_debit1(3).Text = "0.0"
             Me.Tb_debit1(5).Text = "0.0"
             Me.Tb_debit1(6).Text = "0.0"
             Me.Tb_ruic.Text = "0.0"
            Me.Ob_Mr.Enabled = False
            Me.Ob_Caquot.Enabled = False
            Me.Ob_Mh.Enabled = False
        Case Is = 2
             Me.Tb_debit1(0).Text = "0.0"
             Me.Tb_debit1(2).Text = "0.0"
             Me.Tb_debit1(1).Text = "0.0"
             Me.Tb_debit1(4).Text = "0.0"
             Me.Tb_debit1(3).Text = "0.0"
             Me.Tb_debit1(5).Text = "0.0"
             Me.Tb_debit1(6).Text = "0.0"
             Me.Tb_ruic.Text = "0.0"
         Case Is = 3
             Me.Tb_Qbrut.Text = "0.0"
             Me.Tb_debit(0).Text = "0.0"
             Me.Tb_debit(1).Text = "0.0"
             Me.Tb_debit(2).Text = "0.0"
            Me.Ob_Mr.Enabled = False
            Me.Ob_Caquot.Enabled = False
            Me.Ob_Mh.Enabled = False
    End Select
        ' impression false
                    Me.MnuPrint.Enabled = False
    Call ini_schema
'    ini_bv
    ouv_sauve = True
'    Cmd_resu_Click
'    ehyd.qfuite = 0
End Sub
Private Sub reini_form1()
    owner.fdessin.UC_graphique1.Visible = False
    Me.MnuPrint.Enabled = False
     Me.Tb_debit(2).Text = "0.0"
     Me.Tb_debit1(5).Text = "0.0"
     Me.Tb_debit1(6).Text = "0.0"
     Me.Tb_ruic.Text = "0.0"
End Sub

Private Sub mnusuppr_Click()
Dim za As st_save
Dim za1 As st_save1
Dim lhFicDbf1 As Integer, reponse As Integer
If ProtectCheck(2) <> 0 Then End

If Trim(Cb_bassin.Text) <> "" Then
    Call funlockb

    reponse = MsgBox(Trim(Cb_bassin.Text) + " va �tre supprim� .", 4, "Suppression d'un bassin versant")
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
            If Trim(za1.type) = nom_type Then
                za = za1.stsave
                If Trim(za.nom) <> Trim(Cb_bassin.Text) Then
                    FileLength = LOF(lhFicDbf1) / Len(za1) + 1
                    Put #lhFicDbf1, FileLength, za1
                End If
            Else
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
    owner.fdessin.Image2.Visible = False
    Me.Tb_titre.Text = ""
    Me.Caption = fen_titre
    Me.Tb_car_ep(0).Text = "0"
    Me.Tb_car_ep(3).Text = "0"
    Me.Tb_car_ep(2).Text = "0"
    Me.Tb_car_ep(1).Text = "0"
    Me.Tb_car_eu(0).Text = "0"
    Me.Tb_car_eu(2).Text = "0"
    Me.tb_par_ep(0).Text = "0.0"
    Me.tb_par_ep(1).Text = "0.0"
    Me.Tb_par_eu(0).Text = "0"
    Me.Tb_car_eu(1).Text = "0"
    Me.Tb_carep_rur(0).Text = "0"
    Me.Tb_carep_rur(1).Text = "0"
    Me.Tb_carep_rur(2).Text = "0.0"
    Me.Tb_carep_rur(3).Text = "0"
    Me.Tb_carep_rur(4).Text = "0"
    Me.Tb_par_eu(1).Text = "0.0"
    Me.Tb_par_eu(2).Text = "0.0"
    Me.Tb_Qbrut.Text = "0.0"
    Me.Tb_debit(0).Text = "0.0"
    Me.Tb_debit(1).Text = "0.0"
    Me.Tb_debit(2).Text = "0.0"
    Me.Tb_debit1(0).Text = "0.0"
    Me.Tb_debit1(2).Text = "0.0"
    Me.Tb_debit1(1).Text = "0.0"
    Me.Tb_debit1(4).Text = "0.0"
    Me.Tb_debit1(3).Text = "0.0"
    'Me.Opt_rural.Value = False
    'Me.Opt_urbain.Value = True
    'Me.Ob_Mr.Enabled = False
    'Me.Ob_Caquot.Enabled = False
    'Me.Ob_Mh.Enabled = False
    'Me.Cmd_del.Visible = False
    'Call ini_schema
    'ini_bv
    tc = 0
    DM = 0
    dt = 0
    opt_cli = False
    Me.Ob_Mr.Enabled = False
    Me.Ob_Caquot.Enabled = False
    Me.Ob_Mh.Enabled = False
'    Me.Cmd_del.Visible = False
    Me.Opt_rural.Value = False
    Me.Opt_urbain.Value = True
    Call ini_schema
    ini_bv
    opt_cli = True
    ouv_sauve = False
    save_fich = False
    End If


End If
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + nom_fich + " est d�j� en cours d'utilisation.")
    End If

Call flockb(nom_fich)

End Sub

Private Sub Ob_Caquot_Click()
If opt_cli Then
    ebv.Qchoisi = "CAQUOT"
    Call affiche_debit_sel
End If
End Sub
Private Sub Ob_mr_Click()
If opt_cli Then
     ebv.Qchoisi = "RATION"
   Call affiche_debit_sel
End If
End Sub
Private Sub Ob_mh_Click()
If opt_cli Then
    ebv.Qchoisi = "HYDROG"
    Call affiche_debit_sel
End If
End Sub
Private Sub affiche_debit_sel()

    Me.Tb_debit(1).Enabled = Me.Ob_Mr.Value
    Me.Tb_debit(0).Enabled = Me.Ob_Caquot.Value
    Me.Tb_debit(2).Enabled = Me.Ob_Mh.Value
    

End Sub
Private Sub Opt_rural_Click()
Dim i As Integer
owner.affich_aide Me.Name, "Type de bassins"
If opt_cli Then
   Me.Frm_ceu.Visible = False
    For i = 1 To Tb_car_eu.Count
        Me.Tb_car_eu(i - 1).Visible = False
        Me.Lb_car_eu(i - 1).Visible = False
        Me.Lb_ucar_eu(i - 1).Visible = False
    Next
    Me.Frm_peu.Visible = False
    For i = 1 To Tb_par_eu.Count
        Me.Tb_par_eu(i - 1).Visible = False
        Me.Lb_par_eu(i - 1).Visible = False
        Me.Lb_upar_eu(i - 1).Visible = False
    Next
    Me.Lb_aeu.Visible = False
    Me.Lb_beu.Visible = False
    Me.Frm_cbr.Visible = True
    For i = 1 To Tb_carep_rur.Count
        Me.Tb_carep_rur(i - 1).Visible = True
        Me.Lb_carep_rur(i - 1).Visible = True
        Me.Lb_ucarep_rur(i - 1).Visible = True
    Next
    Me.Lb_carep_rur(2).Visible = False
    Me.Lb_aHorton.Visible = True
    Me.Lb_bHorton.Visible = True
    Me.Lb_debit1(0).Visible = False
    Me.Lb_debit1(2).Visible = False
    Me.Lb_debit1(4).Visible = False
    Me.Lb_debit1(1).Visible = False
    Me.Lb_debit1(3).Visible = False
    Me.Tb_debit1(0).Visible = False
    Me.Tb_debit1(2).Visible = False
    Me.Tb_debit1(4).Visible = False
    Me.Tb_debit1(1).Visible = False
    Me.Tb_debit1(3).Visible = False
    Me.Lb_udebit1(0).Visible = False
    Me.Lb_udebit1(2).Visible = False
    Me.Lb_udebit1(4).Visible = False
    Me.Lb_udebit1(1).Visible = False
    Me.Lb_udebit1(3).Visible = False
    Me.Tb_ruic.Visible = True
    Me.Lb_ruic.Visible = True
    Me.Lb_uruic.Visible = True
    Me.Tb_ruic.Text = "0.0"
'    Me.Frm_debit.Left = 3700
    Me.Tb_Qbrut.Text = "0.0"
    Me.Tb_debit(0).Text = "0.0"
    Me.Tb_debit(1).Text = "0.0"
    Me.Tb_debit(2).Text = "0.0"
    Me.Tb_debit1(0).Text = "0.0"
    Me.Tb_debit1(2).Text = "0.0"
    Me.Tb_debit1(1).Text = "0.0"
    Me.Tb_debit1(4).Text = "0.0"
    Me.Tb_debit1(3).Text = "0.0"
    Me.Tb_car_eu(0).Text = "0"
'    Me.Frm_cbr.Top = 2520
    Me.Ob_Mr.Enabled = False
    Me.Ob_Caquot.Enabled = False
    Me.Ob_Mh.Enabled = False
    ini_bv
    ebv.type = "R"
    Call ini_schema
    Me.SSTab1.Tab = 0
End If
End Sub

Private Sub Opt_urbain_Click()
Dim i As Integer
owner.affich_aide Me.Name, "Type de bassins"
If opt_cli Then
    Me.Frm_ceu.Visible = True
    For i = 1 To Tb_car_eu.Count
        Me.Tb_car_eu(i - 1).Visible = True
        Me.Lb_car_eu(i - 1).Visible = True
        Me.Lb_ucar_eu(i - 1).Visible = True
    Next
    Me.Frm_peu.Visible = True
    For i = 1 To Tb_par_eu.Count
        Me.Tb_par_eu(i - 1).Visible = True
        Me.Lb_par_eu(i - 1).Visible = True
        Me.Lb_upar_eu(i - 1).Visible = True
    Next
    Me.Lb_par_eu(1).Visible = False
    Me.Lb_aeu.Visible = True
    Me.Lb_beu.Visible = True
    Me.Frm_cbr.Visible = False
    For i = 1 To Tb_carep_rur.Count
        Me.Tb_carep_rur(i - 1).Visible = False
        Me.Lb_carep_rur(i - 1).Visible = False
        Me.Lb_ucarep_rur(i - 1).Visible = False
    Next
    Me.Lb_aHorton.Visible = False
    Me.Lb_bHorton.Visible = False
    Me.Lb_debit1(0).Visible = True
    Me.Lb_debit1(2).Visible = True
    Me.Lb_debit1(4).Visible = True
    Me.Lb_debit1(1).Visible = True
    Me.Lb_debit1(3).Visible = True
    Me.Tb_debit1(0).Visible = True
    Me.Tb_debit1(2).Visible = True
    Me.Tb_debit1(4).Visible = True
    Me.Tb_debit1(1).Visible = True
    Me.Tb_debit1(3).Visible = True
    Me.Lb_udebit1(0).Visible = True
    Me.Lb_udebit1(2).Visible = True
    Me.Lb_udebit1(4).Visible = True
    Me.Lb_udebit1(1).Visible = True
    Me.Lb_udebit1(3).Visible = True
    Me.Tb_ruic.Visible = False
    Me.Lb_ruic.Visible = False
    Me.Lb_uruic.Visible = False
    Me.Tb_ruic.Text = "0.0"
'    Me.Frm_debit.Left = 120
    Me.Tb_Qbrut.Text = "0.0"
    Me.Tb_debit(0).Text = "0.0"
    Me.Tb_debit(1).Text = "0.0"
    Me.Tb_debit(2).Text = "0.0"
    Me.Tb_debit1(0).Text = "0.0"
    Me.Tb_debit1(2).Text = "0.0"
    Me.Tb_debit1(1).Text = "0.0"
    Me.Tb_debit1(4).Text = "0.0"
    Me.Tb_debit1(3).Text = "0.0"
'    Me.Frm_ceu.Top = 2760
    Me.Ob_Mr.Enabled = False
    Me.Ob_Caquot.Enabled = False
    Me.Ob_Mh.Enabled = False
    ini_bv
    Call ini_schema
    Me.SSTab1.Tab = 0
'    ebv.type = "U"
End If
End Sub
Private Sub Cmd_resu_Click()
Dim crui As Double
'construire Bv
'Construire ParHydrau
'If ebv.type = "U" Or (ebv.type = "R" And val(Me.Tb_carep_rur(0).Text) > 0 And val(Me.Tb_carep_rur(1).Text) > 0 _
'    And val(Me.Tb_carep_rur(2).Text) > 0 And val(Me.Tb_carep_rur(3).Text) > 0 And val(Me.Tb_carep_rur(4).Text) > 0) Then
If ebv.type = "U" Or (ebv.type = "R" And val(Me.Tb_carep_rur(1).Text) > 0 _
    And val(Me.Tb_carep_rur(2).Text) > 0 And val(Me.Tb_carep_rur(3).Text) > 0 And val(Me.Tb_carep_rur(4).Text) > 0) Then
If val(Me.Tb_par_pl(0)) <> 0 And val(Me.Tb_par_pl(1)) <> 0 And val(Me.Tb_par_pl(2)) <> 0 And val(Me.Tb_par_pl(3)) <> 0 Then
    Call calc_hyd
    SSTab1.Tab = 1
        ' impression true
                    Me.MnuPrint.Enabled = True
Else
    Call C_bv
    Call c_ph
'   a eliminer Call calc_resu
    Call calcul_debit_ep(ebv, eph)
    Call affiche_debit_ep(ebv)
    Call calcul_debit_epmr(ebv, eph)
    Call affiche_debit_epmr(ebv)
    If ebv.nhab > 0 Then
        Call calcul_debit_eu(ebv, eph)
        Call affiche_debit_eu(ebv)
    End If
    ouv_sauve = True
'*******************methode hyeto/hydro
    If ebv.surface <> 0 And ebv.lghydr <> 0 And ebv.phydr <> 0 And eph.amontana <> 0 _
        And eph.bmontana <> 0 Then
        Me.Tb_par_pl(1).Text = rempl_virgule(Format(ebv.tc, "#####0"))
        Me.Tb_par_pl(0).Text = rempl_virgule(Format(4 * ebv.tc, "#####0"))
        Me.Tb_txtqf.Text = rempl_virgule(Format(ehyd.qfuite, "#####0"))
        Me.Tb_par_pl(5) = rempl_virgule(Format(ehyd.pas, "#####0"))
        Me.Tb_par_pl(4).Text = rempl_virgule(Format(ehyd.Teta, "###0.00"))
    
        Call calc_hyd
    
        SSTab1.Tab = 1
        ' impression true
                    Me.MnuPrint.Enabled = True
    End If
End If
If Me.Ob_Caquot.Value = False And Me.Ob_Mh.Value = False And Me.Ob_Mr.Value = False Then
    Me.Ob_Caquot.Value = True
    Me.Tb_debit(0).Enabled = True
    ebv.Qchoisi = "CAQUOT"
End If
End If
crui = 0#
If val(Me.Tb_debit1(6).Text) > 0 And val(Me.Tb_par_pl(2).Text) > 0 And val(Me.Tb_car_ep(0).Text) > 0 Then
    crui = val(Me.Tb_debit1(6).Text) / (val(Me.Tb_par_pl(2).Text) * val(Me.Tb_car_ep(0).Text))
End If
Me.Tb_ruic.Text = rempl_virgule(Format(crui, "###0.00"))

End Sub
Private Sub calc_resu()
Call calcul_debit_ep(ebv, eph)
Call affiche_debit_ep(ebv)
Call calcul_debit_epmr(ebv, eph)
Call affiche_debit_epmr(ebv)
If ebv.nhab > 0 Then
    Call calcul_debit_eu(ebv, eph)
    Call affiche_debit_eu(ebv)
End If
ouv_sauve = True
If ebv.surface <> 0 And ebv.lghydr <> 0 And ebv.phydr <> 0 And eph.amontana <> 0 _
    And eph.bmontana <> 0 Then
'    Me.Cmd_hydro.Visible = True
    Me.Tb_par_pl(1).Text = rempl_virgule(Format(ebv.tc, "#####0"))
    Me.Tb_par_pl(0).Text = rempl_virgule(Format(4 * ebv.tc, "#####0"))
'    Me.Tb_Txtqf.Text = Format(ehyd.qfuite, "#####0")
'    Me.tb_par_pl(5) = "5"
'    form_ouv = True
    Call calc_hyd

 '   SSTab1.Tab = 2
End If

End Sub

Private Sub calc_hyd()

Dim DM As Double, dt As Double
Dim tt As Integer
'    Me.Cmd_resu.Visible = False
    owner.fdessin.UC_graphique1.Visible = True
    
    DM = val(Me.Tb_par_pl(1).Text) 'ebv.tc
    dt = val(Me.Tb_par_pl(0).Text) 'ebv.tc * 4
    
    ehyd.DM = DM
    ehyd.dt = dt
    ehyd.pas = val(Me.Tb_par_pl(5).Text)
    ehyd.Teta = txtVersNum(Me.Tb_par_pl(4).Text)  ' 0.5 '
    ehyd.qfuite = val(Me.Tb_txtqf.Text)
    
    If DM <> 0 And dt <> 0 Then
        Call ini_pluie(True)
        Call calcul_hydro1(ebv, eph, ehyd)
      
        Me.Tb_par_pl(2).Text = rempl_virgule(Format(ehyd.HT, "#####0"))
        Me.Tb_par_pl(3).Text = rempl_virgule(Format(ehyd.HM, "#####0"))

'        End With
        pas = val(Me.Tb_par_pl(5).Text)
        tt = calcul_hyeto(ehyd, pas)
        Call dessin_hyeto1
        ' le hyeto brut est stock� dabs la table globale hpluie()
        Call calcul_hydro(pas)
        '   l'hydro est stock� dans lae tableau global Q()
        ' calcul des valeurs qmax,vrui et vstock
        nbval = UBound(Q)
        Qmax = 0#
        vrui = 0#
        For i = 1 To nbval
            If Q(i) > Qmax Then
                Qmax = Q(i)
            End If
            vrui = vrui + Q(i) * pas * 60
        Next
        ' affiche valeur
        Call calcul_stock(ehyd, pas)
        Me.Tb_debit1(5).Text = rempl_virgule(Format(Qmax * 1000, "#########0.0"))
        Me.Tb_debit1(6).Text = rempl_virgule(Format(vrui, "#########0.0"))
        Me.Lb_txtvstock.Caption = rempl_virgule(Format(ehyd.vstock, "########00"))
    
        Me.Tb_debit(2).Text = rempl_virgule(Format(Qmax * 1000, "#########0.0"))
        Me.Ob_Mh.Enabled = True
        ebv.Qhydro = Qmax
'        Call dessin_hydro1(ehyd.qfuite, pas)
        
            Call dessine_hyeto_hydro
        ouv_sauve = True

    Else
    msg = "Les donn�es du BV ne sont pas d�finies"
    MsgBox (msg)
    End If
End Sub


Private Sub Tb_car_ep_Click(Index As Integer)
Dim mes As String
mes = "Caract�ristiques d'un BV"
Select Case Index
 Case Is = 3
   mes = "Tableaux coefficients de ruissellement"
End Select
owner.affich_aide Me.Name, mes
End Sub

Private Sub Tb_car_ep_GotFocus(Index As Integer)
     owner.fdessin.Image2.Visible = True
   Select Case Index
        Case Is = 0
            owner.fdessin.Image2.Left = owner.fdessin.Image1.Left + (owner.fdessin.Image1.Width / 17# * 7#)
            owner.fdessin.Image2.Top = owner.fdessin.Image1.Top + (owner.fdessin.Image1.Height / 9# * 1.1) - owner.fdessin.Image2.Height
        Case Is = 1
            owner.fdessin.Image2.Left = owner.fdessin.Image1.Left + (owner.fdessin.Image1.Width / 17# * 14.2)
            owner.fdessin.Image2.Top = owner.fdessin.Image1.Top + (owner.fdessin.Image1.Height / 9# * 3.5) - owner.fdessin.Image2.Height
        Case Is = 2
            owner.fdessin.Image2.Left = owner.fdessin.Image1.Left + (owner.fdessin.Image1.Width / 17# * 14.1)
            owner.fdessin.Image2.Top = owner.fdessin.Image1.Top + (owner.fdessin.Image1.Height / 9# * 4.4) - owner.fdessin.Image2.Height
        Case Is = 3
            owner.fdessin.Image2.Left = owner.fdessin.Image1.Left + (owner.fdessin.Image1.Width / 17# * 10.6)
            owner.fdessin.Image2.Top = owner.fdessin.Image1.Top + (owner.fdessin.Image1.Height / 9# * 1.5) - owner.fdessin.Image2.Height
'        Case Is = 0
'            owner.fdessin.Image2.Left = owner.fdessin.Image1.Left + 2820
'            owner.fdessin.Image2.Top = owner.fdessin.Image1.Top + 405 - owner.fdessin.Image2.Height
'        Case Is = 1
'            owner.fdessin.Image2.Left = owner.fdessin.Image1.Left + 5355
'            owner.fdessin.Image2.Top = owner.fdessin.Image1.Top + 1500 - owner.fdessin.Image2.Height
'        Case Is = 2
'            owner.fdessin.Image2.Left = owner.fdessin.Image1.Left + 5415
'            owner.fdessin.Image2.Top = owner.fdessin.Image1.Top + 1800 - owner.fdessin.Image2.Height
'        Case Is = 3
'            owner.fdessin.Image2.Left = owner.fdessin.Image1.Left + 3360
'            owner.fdessin.Image2.Top = owner.fdessin.Image1.Top + 840 - owner.fdessin.Image2.Height
    End Select
End Sub

Private Sub Tb_car_ep_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    sval_champ = Tb_car_ep(Index).Text
    iSels = Tb_car_ep(Index).SelStart
    iSell = Tb_car_ep(Index).SelLength
    bKP = True
'    If Len(Tb_car_ep(Index).Text) <= Tb_car_ep(Index).MaxLength Then
'        Select Case Index
'            Case Is = 0
'                 KeyAscii = verif_car(Tb_car_ep(Index).Text, KeyAscii, "Saisie Surface", "R")
'            Case Is = 1
'                 KeyAscii = verif_car(Tb_car_ep(Index).Text, KeyAscii, "Saisie Longueur", "R")
'            Case Is = 2
'                 KeyAscii = verif_car(Tb_car_ep(Index).Text, KeyAscii, "Saisie Pente", "I")
'            Case Is = 3
'                 KeyAscii = verif_car(Tb_car_ep(Index).Text, KeyAscii, "Saisie Coef. de ruissellement", "I")
'        End Select
'    End If
End If
End Sub

Private Sub Tb_car_ep_LostFocus(Index As Integer)
Dim ok As Boolean
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_car_ep", Index, txtVersNum(Tb_car_ep(Index).Text))
    If Not ok Then
        Tb_car_ep(Index).SetFocus
        DoEvents
    End If
    okg = True
End If

End Sub

Private Sub Tb_car_eu_Change(Index As Integer)
 Dim nom As String

If bKP Then
        Select Case Index
            Case Is = 0
                nom = verif_cart0(Tb_car_eu(Index).Text, "Saisie Nbre d'habitants", "I")
           Case Is = 1
                nom = verif_cart0(Tb_car_eu(Index).Text, "Saisie Consommation eau ", "I")
            Case Is = 2
                nom = verif_cart0(Tb_car_eu(Index).Text, "Saisie Taux de dilution", "R")
        End Select
  If nom = "" Then
    Tb_car_eu(Index).Text = sval_champ
    Tb_car_eu(Index).SelStart = iSels
    Tb_car_eu(Index).SelLength = iSell
  End If
End If
'****

   Call reini_form(2)
    sval_champ = ""
    bKP = False

End Sub

Private Sub Tb_car_eu_Click(Index As Integer)
Dim mes As String
mes = "M�thodes d'�valuation des d�bits de temps sec"
owner.affich_aide Me.Name, mes

End Sub

Private Sub Tb_car_eu_GotFocus(Index As Integer)
     owner.fdessin.Image2.Visible = True
    Select Case Index
        Case Is = 0
            owner.fdessin.Image2.Left = owner.fdessin.Image1.Left + (owner.fdessin.Image1.Width / 17# * 9.3)
            owner.fdessin.Image2.Top = owner.fdessin.Image1.Top + (owner.fdessin.Image1.Height / 9# * 5.4) - owner.fdessin.Image2.Height
        Case Is = 1
            owner.fdessin.Image2.Left = owner.fdessin.Image1.Left + (owner.fdessin.Image1.Width / 17# * 9.7)
            owner.fdessin.Image2.Top = owner.fdessin.Image1.Top + (owner.fdessin.Image1.Height / 9# * 3.8) - owner.fdessin.Image2.Height
        Case Is = 2
            owner.fdessin.Image2.Left = owner.fdessin.Image1.Left + (owner.fdessin.Image1.Width / 17# * 9.3)
            owner.fdessin.Image2.Top = owner.fdessin.Image1.Top + (owner.fdessin.Image1.Height / 9# * 2.3) - owner.fdessin.Image2.Height
'        Case Is = 0
'            owner.fdessin.Image2.Left = owner.fdessin.Image1.Left + 3540
'            owner.fdessin.Image2.Top = owner.fdessin.Image1.Top + 2085 - owner.fdessin.Image2.Height
'        Case Is = 1
'            owner.fdessin.Image2.Left = owner.fdessin.Image1.Left + 3675
'            owner.fdessin.Image2.Top = owner.fdessin.Image1.Top + 1680 - owner.fdessin.Image2.Height
'        Case Is = 2
'            owner.fdessin.Image2.Left = owner.fdessin.Image1.Left + 3450
'            owner.fdessin.Image2.Top = owner.fdessin.Image1.Top + 1080 - owner.fdessin.Image2.Height
    End Select
End Sub

Private Sub Tb_car_eu_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    sval_champ = Tb_car_eu(Index).Text
    iSels = Tb_car_eu(Index).SelStart
    iSell = Tb_car_eu(Index).SelLength
    bKP = True
'    If Len(Tb_car_eu(Index).Text) <= Tb_car_eu(Index).MaxLength Then
'       Select Case Index
'            Case Is = 0
'                KeyAscii = verif_car(Tb_car_eu(Index).Text, KeyAscii, "Saisie Nbre d'habitants", "I")
'           Case Is = 1
'                KeyAscii = verif_car(Tb_car_eu(Index).Text, KeyAscii, "Saisie Consommation eau ", "I")
'            Case Is = 2
'                KeyAscii = verif_car(Tb_car_eu(Index).Text, KeyAscii, "Saisie Taux de dilution", "R")
'        End Select
'    End If
End If
End Sub

Private Sub Tb_car_eu_LostFocus(Index As Integer)
Dim ok As Boolean
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_car_eu", Index, txtVersNum(Tb_car_eu(Index).Text))
    If Not ok Then
        Tb_car_eu(Index).SetFocus
        DoEvents
    End If
    okg = True
End If

End Sub

Private Sub Tb_carep_rur_Change(Index As Integer)
Dim nom As String

If bKP Then
        Select Case Index
            Case Is = 0
                nom = verif_cart0(Tb_carep_rur(Index).Text, "Saisie Pertes initiales", "I")
            Case Is = 1
                nom = verif_cart0(Tb_carep_rur(Index).Text, "Saisie Vitesse limite d'infiltration", "I")
            Case Is = 2
                nom = verif_cart0(Tb_carep_rur(Index).Text, "Saisie param�tre a de Horton", "R")
            Case Is = 3
                nom = verif_cart0(Tb_carep_rur(Index).Text, "Saisie param�tre b de Horton", "R")
            Case Is = 4
                nom = verif_cart0(Tb_carep_rur(Index).Text, "Saisie Temps de r�ponse", "I")
        End Select
  If nom = "" Then
    Tb_carep_rur(Index).Text = sval_champ
    Tb_carep_rur(Index).SelStart = iSels
    Tb_carep_rur(Index).SelLength = iSell
  End If
End If
'****

   Call reini_form(1)
    sval_champ = ""
    bKP = False

End Sub
Private Sub Tb_carep_rur_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    sval_champ = Tb_carep_rur(Index).Text
    iSels = Tb_carep_rur(Index).SelStart
    iSell = Tb_carep_rur(Index).SelLength
    bKP = True
'    If Len(Tb_carep_rur(Index).Text) <= Tb_carep_rur(Index).MaxLength Then
'        Select Case Index
'            Case Is = 0
'                KeyAscii = verif_car(Tb_carep_rur(Index).Text, KeyAscii, "Saisie Pertes initiales", "I")
'            Case Is = 1
'                KeyAscii = verif_car(Tb_carep_rur(Index).Text, KeyAscii, "Saisie Vitesse limite d'infiltration", "I")
'            Case Is = 2
'                KeyAscii = verif_car(Tb_carep_rur(Index).Text, KeyAscii, "Saisie param�tre a de Horton", "R")
'            Case Is = 3
'                KeyAscii = verif_car(Tb_carep_rur(Index).Text, KeyAscii, "Saisie param�tre b de Horton", "R")
'            Case Is = 4
'                KeyAscii = verif_car(Tb_carep_rur(Index).Text, KeyAscii, "Saisie Temps de r�ponse", "I")
'       End Select
'    End If
End If
End Sub

Private Sub Tb_carep_rur_LostFocus(Index As Integer)
Dim ok As Boolean
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_carep_rur", Index, txtVersNum(Tb_carep_rur(Index).Text))
    If Not ok Then
        Tb_carep_rur(Index).SetFocus
        DoEvents
    End If
    okg = True
End If

End Sub

Private Sub Tb_debit_Click(Index As Integer)
Dim mes As String
Select Case Index
 Case Is = 0
   mes = "M�thode superficielle de Caquot"
 Case Is = 1
   mes = "M�thode Rationnelle "
 Case Is = 2
   mes = "M�thode de l'hydrogramme"
End Select
owner.affich_aide Me.Name, mes

End Sub


Private Sub tb_par_ep_Change(Index As Integer)
Dim nom As String

If bKP Then
        Select Case Index
            Case Is = 0
                nom = verif_cart0(tb_par_ep(Index).Text, "Saisie Coefficient aMontana", "R")
            Case Is = 1
                nom = verif_cart0(tb_par_ep(Index).Text, "Saisie Coefficient bMontana", "R")
        End Select
  If nom = "" Then
    tb_par_ep(Index).Text = sval_champ
    tb_par_ep(Index).SelStart = iSels
    tb_par_ep(Index).SelLength = iSell
  End If
End If
'****

   Call reini_form(3)
    sval_champ = ""
    bKP = False

End Sub

Private Sub tb_par_ep_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
     sval_champ = tb_par_ep(Index).Text
    iSels = tb_par_ep(Index).SelStart
    iSell = tb_par_ep(Index).SelLength
    bKP = True
'   If Len(tb_par_ep(Index).Text) <= tb_par_ep(Index).MaxLength Then
'        Select Case Index
'            Case Is = 0
'                KeyAscii = verif_car(tb_par_ep(Index).Text, KeyAscii, "Saisie Coefficient aMontana", "R")
'            Case Is = 1
'                KeyAscii = verif_car(tb_par_ep(Index).Text, KeyAscii, "Saisie Coefficient bMontana", "R")
'        End Select
'    End If
End If
End Sub

Private Sub tb_par_ep_LostFocus(Index As Integer)
Dim ok As Boolean
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_par_ep", Index, txtVersNum(tb_par_ep(Index).Text))
    If Not ok Then
        tb_par_ep(Index).SetFocus
        DoEvents
    End If
    okg = True
End If
End Sub

Private Sub Tb_par_eu_Change(Index As Integer)
 Dim nom As String

If bKP Then
        Select Case Index
            Case Is = 0
                nom = verif_cart0(Tb_par_eu(Index).Text, "Saisie Intensit� pluie de rin�age", "I")
            Case Is = 1
                nom = verif_cart0(Tb_par_eu(Index).Text, "Saisie Coefficient pointe Aeu", "R")
            Case Is = 2
                nom = verif_cart0(Tb_par_eu(Index).Text, "Saisie Coefficient pointe Beu", "R")
        End Select
  If nom = "" Then
    Tb_par_eu(Index).Text = sval_champ
    Tb_par_eu(Index).SelStart = iSels
    Tb_par_eu(Index).SelLength = iSell
  End If
End If
'****

    Call reini_form(2)
     sval_champ = ""
    bKP = False

End Sub

Private Sub Tb_par_eu_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    sval_champ = Tb_par_eu(Index).Text
    iSels = Tb_par_eu(Index).SelStart
    iSell = Tb_par_eu(Index).SelLength
    bKP = True
'    If Len(Tb_par_eu(Index).Text) <= Tb_par_eu(Index).MaxLength Then
'        Select Case Index
'            Case Is = 0
'                KeyAscii = verif_car(Tb_par_eu(Index).Text, KeyAscii, "Saisie Intensit� pluie de rin�age", "I")
'            Case Is = 1
'                KeyAscii = verif_car(Tb_par_eu(Index).Text, KeyAscii, "Saisie Coefficient pointe Aeu", "R")
'            Case Is = 2
'                KeyAscii = verif_car(Tb_par_eu(Index).Text, KeyAscii, "Saisie Coefficient pointe Beu", "R")
'        End Select
'    End If
End If
End Sub

Private Sub Tb_par_eu_LostFocus(Index As Integer)
Dim ok As Boolean
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_par_eu", Index, txtVersNum(Tb_par_eu(Index).Text))
    If Not ok Then
        Tb_par_eu(Index).SetFocus
        DoEvents
    End If
    okg = True
End If

End Sub

Private Sub Tb_par_pl_Change(Index As Integer)
Dim nom As String

If bKP Then
        Select Case Index
            Case Is = 0
                nom = verif_cart0(Tb_par_pl(Index).Text, "Saisie Dur�e totale", "I")
            Case Is = 1
                nom = verif_cart0(Tb_par_pl(Index).Text, "Saisie Dur�e intense", "I")
           Case Is = 4
                nom = verif_cart0(Tb_par_pl(Index).Text, "Saisie D�calage de la pointe", "R")
           Case Is = 5
                nom = verif_cart0(Tb_par_pl(Index).Text, "Saisie Pas de calcul", "I")
        End Select
  If nom = "" Then
    Tb_par_pl(Index).Text = sval_champ
    Tb_par_pl(Index).SelStart = iSels
    Tb_par_pl(Index).SelLength = iSell
  Else
  End If
End If
'****

Select Case Index
    Case Is = 0
        If opt_cli Then
'            owner.fdessin.UC_graphique1.Visible = False
'            Me.MnuPrint.Enabled = False
            Call reini_form1
        End If
    Case Is = 1
        If opt_cli Then
'            owner.fdessin.UC_graphique1.Visible = False
'            Me.MnuPrint.Enabled = False
            Call reini_form1
        End If
    Case Is = 4
        If opt_cli Then
'            owner.fdessin.UC_graphique1.Visible = False
'            Me.MnuPrint.Enabled = False
            Call reini_form1
        End If
    Case Is = 5
        If opt_cli Then
            If Trim(Me.Tb_par_pl(5).Text) = "" Or Trim(Me.Tb_par_pl(5).Text) = "0" Then
                Me.Tb_par_pl(5).Text = "5"
           End If
'                owner.fdessin.UC_graphique1.Visible = False
'                Me.MnuPrint.Enabled = False
                Call reini_form1
        End If
End Select

    ouv_sauve = True
   
     sval_champ = ""
    bKP = False

End Sub

Private Sub Tb_par_pl_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Call key13(Me)
Else
    sval_champ = Tb_par_pl(Index).Text
    iSels = Tb_par_pl(Index).SelStart
    iSell = Tb_par_pl(Index).SelLength
    bKP = True
'    If Len(Tb_par_pl(Index).Text) <= Tb_par_pl(Index).MaxLength Then
'        Select Case Index
'            Case Is = 0
'                KeyAscii = verif_car(Tb_par_pl(Index).Text, KeyAscii, "Saisie Dur�e totale", "I")
'            Case Is = 1
'                KeyAscii = verif_car(Tb_par_pl(Index).Text, KeyAscii, "Saisie Dur�e intense", "I")
'           Case Is = 4
'                KeyAscii = verif_car(Tb_par_pl(Index).Text, KeyAscii, "Saisie D�calage de la pointe", "R")
'           Case Is = 5
'                KeyAscii = verif_car(Tb_par_pl(Index).Text, KeyAscii, "Saisie Pas de calcul", "I")
'        End Select
'    End If
End If
End Sub
Private Sub Tb_par_pl_LostFocus(Index As Integer)
Dim ok As Boolean
If Index = 0 Or Index = 1 Or Index = 4 Or Index = 5 Then
If okg Then
    okg = False
    ok = recup_defchamp(Me.Name, "Tb_par_pl", Index, txtVersNum(Tb_par_pl(Index).Text))
    If Not ok Then
        Tb_par_pl(Index).SetFocus
        DoEvents
    End If
    okg = True
End If
End If
End Sub

Private Sub Tb_titre_Change()
    Me.Caption = fen_titre + " : " + Tb_titre.Text
End Sub

Private Sub Tb_Txtqf_Change()
'Dim q1() As Variant, liste() As Variant
'Dim q10() As Variant, q11() As Variant, q12() As Variant
'Dim q0()
'ReDim q1(1)
'Dim i As Integer
'If Not bv_charge Then
'    ReDim q0(UBound(Q))
'    ReDim q10(UBound(Q), 2)
'    ReDim q11(UBound(Q), 2)
'    ReDim q12(UBound(Hpluie), 2)
'    If Trim(Tb_txtqf.Text) = "" Then
'        Tb_txtqf.Text = "0"
'    End If
    ehyd.qfuite = val(Me.Tb_txtqf.Text)
'
'    For i = 1 To UBound(Q)
'        liste = Array(i * ehyd.pas * 1#, Q(i) * 1000#)
'        q10(i, 1) = i * ehyd.pas * 1#
'        q10(i, 2) = Q(i) * 1000#
'        liste = Array(i * ehyd.pas, ehyd.qfuite)
'        q11(i, 1) = i * ehyd.pas
'        q11(i, 2) = ehyd.qfuite
'        q0(i) = Q(i) * 1000#
'    Next
'    For i = 1 To UBound(Hpluie)
'        liste = Array(i * ehyd.pas, Hpluie(i))
'        q12(i, 1) = i * ehyd.pas
'        q12(i, 2) = Hpluie(i)
'    Next
'    If opt_cli Then
'        Call calcul_stock(ehyd, val(Me.Tb_par_pl(5).Text))
'        Me.Lb_txtvstock.Caption = rempl_virgule(Format(ehyd.vstock, "########00"))
'
'     '   Call dessin_hydro1(ehyd.qfuite, Val(Me.tb_par_pl(5).Text)) 'calc_hyd
'        owner.fdessin.UC_graphique1.graphique_clear
'        owner.fdessin.UC_graphique1.dess_cadre 8, 2, 50, 6, 2, 10, 6, 1, 10
'    '    owner.fdessin.uc_graphique1.dess_courbe q0, "N", UBound(q0), &H80FF80
'    '    q1(1) = ehyd.qfuite
'    '    owner.fdessin.uc_graphique1.dess_courbe q1, "N", UBound(q0), &H80C0FF
'    '    owner.fdessin.uc_graphique1.dess_courbe hpluie, "I", UBound(hpluie), &HFFFF80
'        owner.fdessin.UC_graphique1.dess_courbe q10, "N", &H80FF80
'        owner.fdessin.UC_graphique1.dess_courbe q11, "N", &H80C0FF
'        owner.fdessin.UC_graphique1.dess_courbe q12, "I", &HFFFF80
'       ouv_sauve = True
'    End If
'Else
'    bv_charge = False
'End If
End Sub
Private Sub Tb_Txtqf_KeyPress(KeyAscii As Integer)
Dim reponse As Integer
If Len(Tb_txtqf.Text) <= Tb_txtqf.MaxLength Then
    KeyAscii = verif_car(Tb_txtqf.Text, KeyAscii, "Saisie D�bit de fuite", "R")
End If

End Sub
Private Sub dessine_hyeto_hydro()
Dim q1() As Variant, liste() As Variant
Dim q10() As Variant, q11() As Variant, q12() As Variant
Dim q0()
ReDim q1(1)
Dim i As Integer
ReDim q0(UBound(Q))
ReDim q10(UBound(Q), 2)
ReDim q11(UBound(Q), 2)
ReDim q12(UBound(Hpluie), 2)
For i = 1 To UBound(Q)
    liste = Array(i * ehyd.pas * 1#, Q(i) * 1000#)
    q10(i, 1) = i * ehyd.pas * 1#
    q10(i, 2) = Q(i) * 1000#
    liste = Array(i * ehyd.pas, ehyd.qfuite)
' dessin de la fuite
'    q11(i, 1) = i * ehyd.pas
'    q11(i, 2) = ehyd.qfuite
    q0(i) = Q(i) * 1000#
Next
For i = 1 To UBound(Hpluie)
    liste = Array(i * ehyd.pas, Hpluie(i))
    q12(i, 1) = i * ehyd.pas
    q12(i, 2) = Hpluie(i)
Next
'**dessin fen�tre**********
    owner.fdessin.UC_graphique1.reinit 7, "Arial"
    owner.fdessin.UC_graphique1.graphique_clear
    owner.fdessin.UC_graphique1.init_title
    owner.fdessin.UC_graphique1.init_titleh "HYETOGRAMME DE LA PLUIE"
    owner.fdessin.UC_graphique1.init_titleb "HYDROGRAMME DE RUISSELEMENT"
     owner.fdessin.UC_graphique1.init_arrondi_y 1
    owner.fdessin.UC_graphique1.init_MaxYn q10
    owner.fdessin.UC_graphique1.init_MaxYi q12
    owner.fdessin.UC_graphique1.init_EchYn 0.6
    owner.fdessin.UC_graphique1.init_EchYi 0.3
    owner.fdessin.UC_graphique1.init_MaxXn q10 'q0
'    owner.fdessin.uc_graphique1.init_MaxXn hpluie
    owner.fdessin.UC_graphique1.init_EchXn 1#
    owner.fdessin.UC_graphique1.dess_cadre 8, 2, 50, 6, 2, 10, 6, 1, 10
'    owner.fdessin.uc_graphique1.dess_courbe q0, "N", UBound(q0), &H80FF80
'    q1(1) = ehyd.qfuite
'    owner.fdessin.uc_graphique1.dess_courbe q1, "N", UBound(q), &H80C0FF
'    owner.fdessin.uc_graphique1.dess_courbe hpluie, "I", UBound(hpluie), &HFFFF80
    owner.fdessin.UC_graphique1.dess_courbe q10, "N", &H80FF80
' dessin de la fuite
'    owner.fdessin.UC_graphique1.dess_courbe q11, "N", &H80C0FF
    owner.fdessin.UC_graphique1.dess_courbe q12, "I", &HFFFF80
'**dessin temporaire*************
    Frm_desprint.UC_graphique1.reinit 7, "Arial"
    Frm_desprint.UC_graphique1.graphique_clear
    Frm_desprint.UC_graphique1.init_title
    Frm_desprint.UC_graphique1.init_titleh "HYETOGRAMME DE LA PLUIE"
    Frm_desprint.UC_graphique1.init_titleb "HYDROGRAMME DE RUISSELEMENT"
    Frm_desprint.UC_graphique1.init_arrondi_y 1
    Frm_desprint.UC_graphique1.init_MaxYn q10
    Frm_desprint.UC_graphique1.init_MaxYi q12
    Frm_desprint.UC_graphique1.init_EchYn 0.6
    Frm_desprint.UC_graphique1.init_EchYi 0.3
    Frm_desprint.UC_graphique1.init_MaxXn q10 'q0
'    owner.fdessin.uc_graphique1.init_MaxXn hpluie
    Frm_desprint.UC_graphique1.init_EchXn 1#
    Frm_desprint.UC_graphique1.dess_cadre 8, 2, 50, 6, 2, 10, 6, 1, 10
'    owner.fdessin.uc_graphique1.dess_courbe q0, "N", UBound(q0), &H80FF80
'    q1(1) = ehyd.qfuite
'    owner.fdessin.uc_graphique1.dess_courbe q1, "N", UBound(q), &H80C0FF
'    owner.fdessin.uc_graphique1.dess_courbe hpluie, "I", UBound(hpluie), &HFFFF80
    Frm_desprint.UC_graphique1.dess_courbe q10, "N", &H80FF80
    Frm_desprint.UC_graphique1.dess_courbe q12, "I", &HFFFF80
'**********************************
End Sub
Public Sub Mquitter()
    MnuQuit_Click
End Sub
Public Sub Mquit()
    m_quitter_Click
End Sub
Public Sub Menregistrer()
    mnusave_Click
End Sub
Public Sub Msupprimer()
    mnusuppr_Click
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



