VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "À propos de"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "À propos de"
   Begin VB.Timer Timer1 
      Interval        =   40
      Left            =   240
      Top             =   960
   End
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   120
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   461
      TabIndex        =   7
      Top             =   2520
      Width           =   6975
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":08CA
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   5880
      TabIndex        =   0
      Tag             =   "OK"
      Top             =   4065
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "Infos &système..."
      Height          =   345
      Left            =   5880
      TabIndex        =   1
      Tag             =   "Infos &système..."
      Top             =   4515
      Width           =   1245
   End
   Begin VB.Image Image1 
      Height          =   1980
      Left            =   4680
      Picture         =   "frmAbout.frx":1194
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2400
   End
   Begin VB.Label NumLicence 
      AutoSize        =   -1  'True
      Caption         =   "Licence N° "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   8
      Tag             =   "Version"
      Top             =   720
      Width           =   1185
   End
   Begin VB.Label lblDescription 
      Caption         =   "Boite à Outils Hydrologie,  Hydraulique et Assainissement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   930
      Left            =   960
      TabIndex        =   6
      Tag             =   "Description de l'application"
      Top             =   1200
      Width           =   2805
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Titre de l'application"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   960
      TabIndex        =   5
      Tag             =   "Titre de l'application"
      Top             =   240
      Width           =   2430
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      TabIndex        =   4
      Tag             =   "Version"
      Top             =   240
      Width           =   930
   End
   Begin VB.Label WarningLabel 
      Caption         =   "Avertissement:                            Logiciel protégé                      Toute reproduction interdite  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   705
      Left            =   135
      TabIndex        =   3
      Tag             =   "Avertissement: ..."
      Top             =   4065
      Width           =   2670
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Reg Key - Options de sécurité ...
Const KEY_ALL_ACCESS = &H2003F

' Reg Key - Types de ROOT...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' chaîne Unicode terminée par 0
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Const SRCCOPY = &HCC0020
Const ShowText$ = "Frank TRIFILETTI"
Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Long, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer
Dim ShowIt%, monIndMsg%
Dim monTabString(11) As String


Private Sub Form_Load()
    lblVersion.Caption = "version " & App.Major & "." & App.Minor & "." & App.Revision
    
    lblTitle.Caption = App.Title
    Caption = "A propos de " + App.Title
    
    lblVersion.Left = lblTitle.Left + lblTitle.Width + 60
    lblVersion.Top = lblTitle.Top + lblTitle.Height - lblVersion.Height
    
    'CentrerFenetreEcran Me
    
    'Affectation du contexte d'aide
    HelpContextID = IDhlp_WinAbout
    
    'Affichage du numéro de licence
'    NumLicence.Caption = "Licence N° : " + SerialNumber
    NumLicence.Caption = F_LBLICENCE + NumeroLicence
   
    'Traitement permettant de lister la boucle des intervenants
    unDecalage = "     "
    WarningLabel.Caption = "Avertissement :" + Chr(13) + unDecalage + "Logiciel protégé"
    WarningLabel.Caption = WarningLabel.Caption + Chr(13) + unDecalage + "Toute reproduction interdite"
    WarningLabel.Font.Bold = True
    'Initialisation de l'indice des messages listant les participants
    monIndMsg% = 0
    'Initialisation des noms des participants
    monTabString(0) = "Production du cahier des charges"
    monTabString(1) = "    CETE DE L'EST / Laboratoire des Ponts et Chaussées de Nancy"
    monTabString(2) = "    CETE DE L'EST / Département INFORMATIQUE"
    monTabString(3) = "Réalisation du développement du logiciel"
    monTabString(4) = "    CETE DE L'EST / Département INFORMATIQUE"
    monTabString(5) = "Diffusion et Assistance au logiciel"
    monTabString(6) = "    CERTU / Département SYSTEMES / Groupe Informatique Technique et Scientifique"
    monTabString(7) = "        - its.sys.certu@developpement-durable.gouv.fr"
    monTabString(8) = "    PND Hydro  / CETE DE L'EST & CETE NORD-PICARDIE"
    monTabString(9) = "        - Pnd-Hydro@developpement-durable.gouv.fr"
End Sub

Private Sub cmdSysInfo_Click()
        Call StartSysInfo
End Sub


Private Sub cmdOK_Click()
    Unload Me
End Sub


Public Sub StartSysInfo()
    On Error GoTo SysInfoErr


        Dim rc As Long
        Dim SysInfoPath As String
        

        ' Lecture dans la base de registres du chemin\nom du programme d'info système...
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Lecture dans la base de registres du chemin du programme d'info système...
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
                ' Valider l'existence d'une version du fichier 32 bits connue
                If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                        

                ' Erreur - Fichier introuvable ...
                Else
                        GoTo SysInfoErr
                End If
        ' Erreur - Entrée de la base de registres introuvable ...
        Else
                GoTo SysInfoErr
        End If
        

        Call Shell(SysInfoPath, vbNormalFocus)
        

        Exit Sub
SysInfoErr:
        MsgBox "Les informations sur le système ne sont pas disponibles pour l'instant", vbOKOnly
End Sub


Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long                                           ' Compteur de boucle
        Dim rc As Long                                          ' Code de retour
        Dim hKey As Long                                        ' Pointeur vers une clé de registre ouvert
        Dim hDepth As Long                                      '
        Dim KeyValType As Long                                  ' Type de données d'une clé de registre
        Dim tmpVal As String                                    ' Stockage temp. pour une valeur de clé de registre
        Dim KeyValSize As Long                                  ' Taille de la variable clé de registre
        '------------------------------------------------------------
        ' Ouvrir RegKey sous KeyRoot {HKEY_LOCAL_MACHINE...}
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Ouvrir clé de registre
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Gérer les erreurs...
        

        tmpVal = String$(1024, 0)                               ' Allouer l'espace pour la variable
        KeyValSize = 1024                                       ' Marquer la taille de la variable
        

        '------------------------------------------------------------
        ' Extraire la valeur de clé de registre...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Lire/créer validation de clé
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Gérer les erreurs
        

        If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 termine les chaînes par 0...
                tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null atteint, extraire de la chaîne
        Else                                                    ' WinNT ne termine pas les chaînes par 0...
                tmpVal = Left(tmpVal, KeyValSize)                   ' 0 non trouvé, extraire chaîne uniquement
        End If
        '---------------------------------------------------------------
        ' Determiner le type de la valeur de la clé pour la convertir...
        '---------------------------------------------------------------
        Select Case KeyValType                                  ' Rechercher types de données...
        Case REG_SZ                                             ' Type de données de clé de registre String
                KeyVal = tmpVal                                     ' Copier valeur de la chaîne
        Case REG_DWORD                                          ' Type de données de clé de registre Double Word
                For i = Len(tmpVal) To 1 Step -1                    ' Convertir chaque bit
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Construire valeur caractère par caractère
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' Convertir Double Word en String
        End Select
        

        GetKeyValue = True                                      ' Renvoyer Réussite
        rc = RegCloseKey(hKey)                                  ' Fermer la clé de registre
        Exit Function                                           ' Sortir

GetKeyError:    ' Nettoyage si erreur...
        KeyVal = ""                                             ' Affecter chaîne vide à la valeur de retour
        GetKeyValue = False                                     ' Renvoyer Échec
        rc = RegCloseKey(hKey)                                  ' Fermer la clé de registre
End Function



Private Sub Timer1_Timer()
    Dim i As Integer
    Dim uneString As String
    
    If (ShowIt% Mod 20 = 0) Then
        Picture1.CurrentX = 20
        Picture1.CurrentY = Picture1.ScaleHeight - 20
        'Affichage du participant d'indice monIndMsg%
        Picture1.Print monTabString(monIndMsg% Mod 10)
        ShowIt% = 1
        If monIndMsg% = 10 Then
            'Pour éviter un débordement de capacité des entiers
            monIndMsg% = 1
        Else
            'Permettra l'affichage du message suivant
            monIndMsg% = monIndMsg% + 1
        End If
    Else
        i = BitBlt(Picture1.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight - 1, Picture1.hDC, 0, 1, SRCCOPY)
        ShowIt% = ShowIt% + 1
    End If
End Sub
