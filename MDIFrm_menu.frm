VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm MDIFrm_menu 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "HYDROUTI"
   ClientHeight    =   5310
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8760
   Icon            =   "MDIFrm_menu.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Cdlg3 
      Left            =   720
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "cdlg3"
   End
   Begin MSComDlg.CommonDialog cdlg2 
      Left            =   2040
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog Cdlg1 
      Left            =   1080
      Top             =   7800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu m_etu 
      Caption         =   "&Etude"
      Begin VB.Menu mnuNouv 
         Caption         =   "&Nouveau"
      End
      Begin VB.Menu mnuOuv 
         Caption         =   "&Ouvrir..."
      End
      Begin VB.Menu f1 
         Caption         =   "-"
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
         Caption         =   "Configurer im&primante..."
      End
      Begin VB.Menu f3 
         Caption         =   "-"
      End
      Begin VB.Menu m_saisie_info 
         Caption         =   "&Informations..."
      End
      Begin VB.Menu f4 
         Caption         =   "-"
      End
      Begin VB.Menu m_Quitter 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu m_hyd 
      Caption         =   "&Hydrologie Hydraulique"
      Begin VB.Menu m_bassin 
         Caption         =   "Bassin &Versant..."
      End
      Begin VB.Menu m_chute 
         Caption         =   "&Chute..."
      End
      Begin VB.Menu m_siphon 
         Caption         =   "&Siphon..."
      End
      Begin VB.Menu m_pompe 
         Caption         =   "Station de &pompage..."
      End
      Begin VB.Menu m_conduite 
         Caption         =   "Con&duite..."
      End
   End
   Begin VB.Menu m_trait_qualit 
      Caption         =   "&Traitement Qualitatif"
      Begin VB.Menu m_DO 
         Caption         =   "Déversoir d'&Orage à crête haute..."
      End
      Begin VB.Menu m_DO_or 
         Caption         =   "Déversoir d'or&Age à ouverture de radier..."
      End
      Begin VB.Menu m_decantation 
         Caption         =   "Bassin de &Décantation..."
      End
      Begin VB.Menu m_Stockage 
         Caption         =   "Bassin de Stockage &Restitution..."
      End
   End
   Begin VB.Menu m_trait_quant 
      Caption         =   "Traitement &Quantitatif"
      Begin VB.Menu m_Retention 
         Caption         =   "Bassin de Retention &Pluviale..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&?"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "A &propos de Hydrouti..."
      End
   End
End
Attribute VB_Name = "MDIFrm_menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "User32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Public fcom As Object
Public fbassin As Object
Public fobjet As Object
Public fdessin As Frm_dessin
Public chemin_aide As String
Public nom_fichier_texte As String
Private nom_fichier_aide As String
Private nom_fichier_exemple As String
Private nom_fichier_com As String
Public lhFicDbf As Integer
Public lhFicDbf1 As Integer
Public FileLength As Integer
Private ouv_save As Boolean
'Private nouv_save As Boolean
Private chemin_init As String
Private lprinter As Printer
Public coef As Double
Private Sub init_etude()
Dim za1 As st_hydrouti
Dim i As Integer
i = 0
    lhFicDbf = FreeFile
    Open nom_fich For Random Access Read Write Lock Read Write As #lhFicDbf Len = Len(za1)
        za1.type = "hydrouti"
        FileLength = LOF(lhFicDbf) / Len(za1) + 1
        Put #lhFicDbf, FileLength, za1
    Close #lhFicDbf
End Sub
Private Function verif_etude() As String
Dim za1 As st_hydrouti
Dim num As Integer
 On Error GoTo test_Error
   lhFicDbf = FreeFile
    Open nom_fich_edit For Input Lock Read Write As #lhFicDbf ' Len = Len(za1)
    Input #lhFicDbf, num, za1.type
        message = ""
        If Trim(za1.type) <> "hydrouti" Then
            message = "Le fichier n'est pas de type HYDROUTI"
        End If
    Close #lhFicDbf
verif_etude = message
Exit Function
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + s + " est déjà en cours d'utilisation.")
    End If
    message = "Anomalie sur le traitement du fichier " + nom_fich_edit + " !"
    verif_etude = message
End Function
Private Sub lect_fich()
'Dim za As st_texte
'    lhFicDbf = FreeFile
'    Open nom_fichier_texte For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za)
'    Get #lhFicDbf, , za
'    text_serv1 = Trim(za.nom1)
'    text_serv2 = Trim(za.nom2)
    text_serv1 = GetSetting("Hydrouti", "Informations", "Info1")
    text_serv2 = GetSetting("Hydrouti", "Informations", "Info2")
'    Close #lhFicDbf
End Sub
Private Sub recopie_init(ByRef za1 As st_hydrouti)
'        FileLength = LOF(lhFicDbf1) / Len(za1) + 1
        Write #lhFicDbf1, FileLength, za1.type, ""
        'za1.reste
End Sub
Private Sub recopie_bv(ByRef za2 As st_save1)
Dim za As st_save
za = za2.stsave
        Write #lhFicDbf1, FileLength, za2.type, za.nom, za.bv.nom, za.bv.type _
        , za.bv.surface, za.bv.imper, za.bv.lghydr, za.bv.phydr _
        , za.bv.nhab, za.bv.tdilu, za.bv.ceau, za.bv.perti, za.bv.vinf _
        , za.bv.ahorton, za.bv.bhorton, za.bv.trep, za.bv.Qbrut _
        , za.bv.Qcor, za.bv.Qmr, za.bv.Qhydro, za.bv.Qeu _
        , za.bv.Qecp, za.bv.Qts, za.bv.Qprin, za.bv.Qrin _
        , za.bv.tc, za.bv.qfuite, za.bv.pas, za.bv.Teta _
        , za.bv.Qchoisi, za.hydro.amontana, za.hydro.bmontana _
        , za.hydro.lcrin, za.hydro.ceau, za.hydro.aeu, za.hydro.beu _
        , za.hydro.a1montana, za.hydro.b1montana, za.hydro.Seuil _
        , za.hydro1.DM, za.hydro1.dt, za.hydro1.HM, za.hydro1.HT _
        , za.hydro1.pas, za.hydro1.Teta, za.hydro1.kdesbor _
        , za.hydro1.qfuite, za.hydro1.vst, za.hydro1.vstock, ""
        'za2.reste
        

End Sub
Private Sub recopie_decant(ByRef za4 As st_savdec1)
Dim za As st_savdecant
za = za4.stsavdecant
        Write #lhFicDbf1, FileLength, za.type, za.nom, za.decant.Q _
        , za.decant.d, za.decant.X, za.decant.Psed, za.decant.Vhor _
        , za.decant.Long, za.decant.larg, za.decant.Hchamb _
        , za.decant.heau, za.decant.Vvert, za.decant.k, "", ""
        'za.reste, za4.reste

End Sub
Private Sub recopie_ret(ByRef za6 As st_savret1)
Dim za As st_savret
za = za6.stsavret
        Write #lhFicDbf1, FileLength, za.type, za.nom, za.retention.nom _
        , za.retention.nombv, za.retention.type_calcul, za.retention.surface _
        , za.retention.Ca, za.retention.qf, za.retention.amontana, za.retention.bmontana _
        , za.retention.deltaH, za.retention.volume, za.retention.a1montana _
        , za.retention.b1montana, za.retention.Seuil, za.retention.desssret.type _
        , za.retention.desssret.opt_long, za.retention.desssret.opt_larg _
        , za.retention.desssret.opt_prof, za.retention.desssret.opt_rap _
        , za.retention.desssret.Longueur, za.retention.desssret.Largeur _
        , za.retention.desssret.Profondeur, za.retention.desssret.Rapport _
        , za.retention.desssret.coef, za.retention.desssret.duree _
        , za.retention.desssret.Hpluie, za.retention.desssret.Hfuite _
        , "", ""
        ' za.reste, za6.reste

End Sub
Private Sub recopie_siphon(ByRef za7 As st_savsi1)
Dim za As st_savsi
za = za7.stsavsi
Write #lhFicDbf1, FileLength, za.type, za.nom, za.siphon.dam, za.siphon.iRadam, za.siphon.Kam, za.siphon.dav, za.siphon.iradav _
, za.siphon.kav, za.siphon.Rdav, za.siphon.Rdam, za.siphon.Jadm, za.siphon.ds, za.siphon.Ks, za.siphon.Qmax, za.siphon.ls _
, za.siphon.Kc, za.siphon.List_coude.coude(0).Nbre, za.siphon.List_coude.coude(0).type, za.siphon.List_coude.coude(0).angle _
, za.siphon.List_coude.coude(0).Rayon, za.siphon.List_coude.coude(1).Nbre, za.siphon.List_coude.coude(1).type, za.siphon.List_coude.coude(1).angle _
, za.siphon.List_coude.coude(1).Rayon, za.siphon.List_coude.coude(2).Nbre, za.siphon.List_coude.coude(2).type, za.siphon.List_coude.coude(2).angle _
, za.siphon.List_coude.coude(2).Rayon, za.siphon.List_coude.coude(3).Nbre, za.siphon.List_coude.coude(3).type, za.siphon.List_coude.coude(3).angle _
, za.siphon.List_coude.coude(3).Rayon, za.siphon.List_coude.coude(4).Nbre, za.siphon.List_coude.coude(4).type, za.siphon.List_coude.coude(4).angle _
, za.siphon.List_coude.coude(4).Rayon, za.siphon.List_coude.coude(5).Nbre, za.siphon.List_coude.coude(5).type, za.siphon.List_coude.coude(5).angle _
, za.siphon.List_coude.coude(5).Rayon, za.siphon.List_coude.coude(6).Nbre, za.siphon.List_coude.coude(6).type, za.siphon.List_coude.coude(6).angle _
, za.siphon.List_coude.coude(6).Rayon, za.siphon.List_coude.coude(7).Nbre, za.siphon.List_coude.coude(7).type, za.siphon.List_coude.coude(7).angle _
, za.siphon.List_coude.coude(7).Rayon, za.siphon.List_coude.coude(8).Nbre, za.siphon.List_coude.coude(8).type, za.siphon.List_coude.coude(8).angle _
, za.siphon.List_coude.coude(8).Rayon, za.siphon.List_coude.coude(9).Nbre, za.siphon.List_coude.coude(9).type, za.siphon.List_coude.coude(9).angle _
, za.siphon.List_coude.coude(9).Rayon, za.siphon.Ipl, za.siphon.deltaH1, za.siphon.deltaH2, za.siphon.IPs _
, za.siphon.tron_amo.Absamo, za.siphon.tron_amo.radamo, za.siphon.tron_amo.Absava _
, za.siphon.tron_amo.radava, za.siphon.tron_amo.conduit.Longueur, za.siphon.tron_amo.conduit.Diametre _
, za.siphon.tron_amo.conduit.pente, za.siphon.tron_amo.conduit.rugosite, za.siphon.tron_amo.conduit.typ _
, za.siphon.tron_ava.Absamo, za.siphon.tron_ava.radamo, za.siphon.tron_ava.Absava _
, za.siphon.tron_ava.radava, za.siphon.tron_ava.conduit.Longueur, za.siphon.tron_ava.conduit.Diametre _
, za.siphon.tron_ava.conduit.pente, za.siphon.tron_ava.conduit.rugosite, za.siphon.tron_ava.conduit.typ _
, "", ""
', za.reste, za7.reste
End Sub
Private Sub recopie_stock(ByRef za8 As st_savsto1)
Dim za As st_savstock
za = za8.stsavstock
        Write #lhFicDbf1, FileLength, za.type, za.nom, za.stockage.nom _
        , za.stockage.nombv, za.stockage.Qpluie, za.stockage.Qts _
        , za.stockage.Qrin, za.stockage.lcrin, za.stockage.surface, za.stockage.imper _
        , za.stockage.tc, za.stockage.Qav, za.stockage.Ipcav _
        , za.stockage.Vr, za.stockage.alphat, za.stockage.volume _
        , za.stockage.dessstock.type, za.stockage.dessstock.opt_long _
        , za.stockage.dessstock.opt_larg, za.stockage.dessstock.opt_prof _
        , za.stockage.dessstock.opt_rap, za.stockage.dessstock.Longueur _
        , za.stockage.dessstock.Largeur, za.stockage.dessstock.Profondeur _
        , za.stockage.dessstock.Rapport, za.stockage.dessstock.Diametre _
        , za.stockage.dessstock.hauteur, za.stockage.dessstock.Diametrec _
        , za.stockage.dessstock.Longueurc, za.stockage.dessstock.coef _
        , "", ""
'        , za.reste, za8.reste

End Sub
Private Sub recopie_deversoir(ByRef za5 As st_savdo1)
Dim za As st_savdo
za = za5.stsavdo
    Write #lhFicDbf1, FileLength, za.type, za.nom, za.edessdo.nom, za.edessdo.nombv _
    , za.edessdo.Qts, za.edessdo.Qrin, za.edessdo.Qpluie, za.edessdo.rdoam, za.edessdo.rdoav _
    , za.edessdo.lgdisp, za.edessdo.phex, za.edessdo.rdoex, za.edessdo.lgca, za.edessdo.dam _
    , za.edessdo.iRadam, za.edessdo.Kam, za.edessdo.Lam, za.edessdo.dav, za.edessdo.iradav _
    , za.edessdo.kav, za.edessdo.Lav, za.edessdo.Tram, za.edessdo.Centon _
    , za.edessdo.tron_amo.Absamo, za.edessdo.tron_amo.radamo, za.edessdo.tron_amo.Absava _
    , za.edessdo.tron_amo.radava, za.edessdo.tron_amo.conduit.Longueur, za.edessdo.tron_amo.conduit.Diametre _
    , za.edessdo.tron_amo.conduit.pente, za.edessdo.tron_amo.conduit.rugosite, za.edessdo.tron_amo.conduit.typ _
    , za.edessdo.tron_ava.Absamo, za.edessdo.tron_ava.radamo, za.edessdo.tron_ava.Absava _
    , za.edessdo.tron_ava.radava, za.edessdo.tron_ava.conduit.Longueur, za.edessdo.tron_ava.conduit.Diametre _
    , za.edessdo.tron_ava.conduit.pente, za.edessdo.tron_ava.conduit.rugosite, za.edessdo.tron_ava.conduit.typ _
    , za.edessdo.tron_dech.Absamo, za.edessdo.tron_dech.radamo, za.edessdo.tron_dech.Absava _
    , za.edessdo.tron_dech.radava, za.edessdo.tron_dech.conduit.Longueur, za.edessdo.tron_dech.conduit.Diametre _
    , za.edessdo.tron_dech.conduit.pente, za.edessdo.tron_dech.conduit.rugosite, za.edessdo.tron_dech.conduit.typ _
    , za.edo.Absamo, za.edo.radamo, za.edo.Absava, za.edo.radava, za.edo.Longueur, za.edo.pente, za.edo.hauteur _
    , za.edo.tron_ava.Absamo, za.edo.tron_ava.radamo, za.edo.tron_ava.Absava _
    , za.edo.tron_ava.radava, za.edo.tron_ava.conduit.Longueur, za.edo.tron_ava.conduit.Diametre _
    , za.edo.tron_ava.conduit.pente, za.edo.tron_ava.conduit.rugosite, za.edo.tron_ava.conduit.typ _
    , za.edo.tav, "", ""
    ' za.reste, za5.reste

End Sub
Private Sub recopie_deversoir_or(ByRef za10 As st_savdo1)
Dim za As st_savdo
za = za10.stsavdo
    Write #lhFicDbf1, FileLength, za.type, za.nom, za.edessdo.nom, za.edessdo.nombv _
    , za.edessdo.Qts, za.edessdo.Qrin, za.edessdo.Qpluie, za.edessdo.rdoam, za.edessdo.rdoav _
    , za.edessdo.lgdisp, za.edessdo.phex, za.edessdo.rdoex, za.edessdo.lgca, za.edessdo.dam _
    , za.edessdo.iRadam, za.edessdo.Kam, za.edessdo.Lam, za.edessdo.dav, za.edessdo.iradav _
    , za.edessdo.kav, za.edessdo.Lav, za.edessdo.Tram, za.edessdo.Centon _
    , za.edessdo.tron_amo.Absamo, za.edessdo.tron_amo.radamo, za.edessdo.tron_amo.Absava _
    , za.edessdo.tron_amo.radava, za.edessdo.tron_amo.conduit.Longueur, za.edessdo.tron_amo.conduit.Diametre _
    , za.edessdo.tron_amo.conduit.pente, za.edessdo.tron_amo.conduit.rugosite, za.edessdo.tron_amo.conduit.typ _
    , za.edessdo.tron_ava.Absamo, za.edessdo.tron_ava.radamo, za.edessdo.tron_ava.Absava _
    , za.edessdo.tron_ava.radava, za.edessdo.tron_ava.conduit.Longueur, za.edessdo.tron_ava.conduit.Diametre _
    , za.edessdo.tron_ava.conduit.pente, za.edessdo.tron_ava.conduit.rugosite, za.edessdo.tron_ava.conduit.typ _
    , za.edessdo.tron_dech.Absamo, za.edessdo.tron_dech.radamo, za.edessdo.tron_dech.Absava _
    , za.edessdo.tron_dech.radava, za.edessdo.tron_dech.conduit.Longueur, za.edessdo.tron_dech.conduit.Diametre _
    , za.edessdo.tron_dech.conduit.pente, za.edessdo.tron_dech.conduit.rugosite, za.edessdo.tron_dech.conduit.typ _
    , za.edo.Absamo, za.edo.radamo, za.edo.Absava, za.edo.radava, za.edo.Longueur, za.edo.pente, za.edo.hauteur _
    , za.edo.tron_ava.Absamo, za.edo.tron_ava.radamo, za.edo.tron_ava.Absava _
    , za.edo.tron_ava.radava, za.edo.tron_ava.conduit.Longueur, za.edo.tron_ava.conduit.Diametre _
    , za.edo.tron_ava.conduit.pente, za.edo.tron_ava.conduit.rugosite, za.edo.tron_ava.conduit.typ _
    , za.edo.tav, "", ""
    ' za.reste, za10.reste

End Sub
Private Sub recopie_chute(ByRef za3 As st_savch1)
Dim za As st_savchute
za = za3.stsavch
        Write #lhFicDbf1, FileLength, za.type, za.nom, za.chute.dam, za.chute.iRadam _
        , za.chute.Kam, za.chute.dav, za.chute.iradav, za.chute.kav _
        , za.chute.Rdav, za.chute.Rdam, za.chute.Qmax, za.chute.h0, za.chute.Long _
        , za.chute.tron_amo.Absamo, za.chute.tron_amo.radamo _
        , za.chute.tron_amo.Absava, za.chute.tron_amo.radava _
        , za.chute.tron_amo.conduit.Longueur, za.chute.tron_amo.conduit.Diametre _
        , za.chute.tron_amo.conduit.pente, za.chute.tron_amo.conduit.rugosite _
        , za.chute.tron_amo.conduit.typ _
        , za.chute.tron_ava.Absamo, za.chute.tron_ava.radamo _
        , za.chute.tron_ava.Absava, za.chute.tron_ava.radava _
        , za.chute.tron_ava.conduit.Longueur, za.chute.tron_ava.conduit.Diametre _
        , za.chute.tron_ava.conduit.pente, za.chute.tron_ava.conduit.rugosite _
        , za.chute.tron_ava.conduit.typ, "", ""
        'za.reste, za3.reste

End Sub
Private Sub recopie_pompe(ByRef za9 As st_savpom1)
Dim za As st_savpompe
za = za9.stsavpo
        Write #lhFicDbf1, FileLength, za.type, za.nom, za.pompe.debits_car.qeum, za.pompe.debits_car.Fp _
        , za.pompe.debits_car.Qeu, za.pompe.debits_car.Qecp, za.pompe.debits_car.Qtsm, za.pompe.debits_car.Qts _
        , za.pompe.debits_car.Qpomp, za.pompe.don_geometrie.Lrflt, za.pompe.don_geometrie.NatRflt, za.pompe.don_geometrie.Drflt, za.pompe.don_geometrie.NivTN _
        , za.pompe.don_geometrie.NivEN, za.pompe.don_geometrie.NivSO _
        , za.pompe.don_geometrie.NivEX, za.pompe.pts_singuliers.Nbc1 _
        , za.pompe.pts_singuliers.Nbc2, za.pompe.pts_singuliers.Nbc3 _
        , za.pompe.pts_singuliers.Nbc4, za.pompe.pts_singuliers.Nbc9 _
        , za.pompe.pts_singuliers.Nbva _
        , za.pompe.pts_singuliers.Nbcl, za.pompe.pts_singuliers.Nbvi _
        , za.pompe.pts_singuliers.Nbve, za.pompe.pts_singuliers.Antb _
        , za.pompe.don_techniques.Nbpom, za.pompe.don_techniques.Ntdph _
        , za.pompe.don_techniques.Vutba, za.pompe.don_techniques.Sectb _
        , za.pompe.don_techniques.Diamb, za.pompe.don_techniques.Longb _
        , za.pompe.don_techniques.Largb, za.pompe.don_techniques.Denivt _
        , za.pompe.don_techniques.Denivhau, za.pompe.don_techniques.Denivbas _
        , za.pompe.resultat.Qpomr, za.pompe.resultat.Drflr, za.pompe.resultat.NatRflr _
        , za.pompe.resultat.VitRflt, za.pompe.resultat.jmpkm, za.pompe.resultat.Denivr _
        , za.pompe.resultat.Vurba, za.pompe.resultat.Nrdph, za.pompe.resultat.Tvidange _
        , za.pompe.resultat.T1cyc, za.pompe.resultat.Nbcyc, za.pompe.resultat.Vmy _
        , za.pompe.resultat.Tsejh, za.pompe.resultat.Singul, za.pompe.resultat.Hmt, "", ""
        'za.reste, za9.reste

End Sub
Private Sub recopie_fich(ByVal nom As String)
'Dim nom As String
Dim list_enreg() As Variant
Dim ir As Integer
Dim za1 As st_hydrouti
Dim za2 As st_save1 'bassin versant
Dim za3 As st_savch1 'chute et conduite
Dim za4 As st_savdec1 'decantation
Dim za5 As st_savdo1 ' deversoir
Dim za6 As st_savret1 'retention
Dim za7 As st_savsi1 'siphon
Dim za8 As st_savsto1 'stockage
Dim za9 As st_savpom1 'stockage
Dim za10 As st_savdo1 ' deversoir
'nom = nom_fich_edit
ir = 0
 On Error GoTo test_Error
ReDim list_enreg(ir)
Call funlocka
Call funlockb
 '   nom = chemin_app + "etude.boh"
    If Dir(nom) <> "" Then
        Kill nom
    End If
   lhFicDbf = FreeFile
    Open nom_fich For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
    Do While Not EOF(lhFicDbf)
        Get #lhFicDbf, , za1
        If Not EOF(lhFicDbf) Then
            ir = ir + 1
            ReDim Preserve list_enreg(ir)
            list_enreg(ir) = za1.type
        End If
    Loop
    Close #lhFicDbf
   lhFicDbf = FreeFile
    Open nom_fich For Random Access Read Lock Read Write As #lhFicDbf Len = Len(za1)
   lhFicDbf1 = FreeFile
    Open nom For Output Lock Read Write As #lhFicDbf1
   ir = 1
Do While ir <= UBound(list_enreg)  'Not EOF(lhFicDbf) And
'    If ir <= UBound(list_enreg) Then
    Select Case Trim(list_enreg(ir))
        Case Is = "hydrouti"
            Get #lhFicDbf, , za1
        Case Is = "versant"
            Get #lhFicDbf, , za2
        Case Is = "chute", "conduite"
            Get #lhFicDbf, , za3
        Case Is = "decantation"
            Get #lhFicDbf, , za4
        Case Is = "deversoir"
            Get #lhFicDbf, , za5
        Case Is = "retention"
            Get #lhFicDbf, , za6
        Case Is = "siphon"
            Get #lhFicDbf, , za7
        Case Is = "stockage"
            Get #lhFicDbf, , za8
        Case Is = "pompe"
            Get #lhFicDbf, , za9
         Case Is = "deversoiror"
            Get #lhFicDbf, , za10
  End Select
'    If Not EOF(lhFicDbf) Then
    Select Case Trim(list_enreg(ir))
        Case Is = "hydrouti"
            Call recopie_init(za1)
        Case Is = "versant"
            Call recopie_bv(za2)
        Case Is = "chute", "conduite"
            Call recopie_chute(za3)
        Case Is = "decantation"
            Call recopie_decant(za4)
        Case Is = "deversoir"
            Call recopie_deversoir(za5)
        Case Is = "retention"
            Call recopie_ret(za6)
        Case Is = "siphon"
            Call recopie_siphon(za7)
        Case Is = "stockage"
            Call recopie_stock(za8)
        Case Is = "pompe"
            Call recopie_pompe(za9)
        Case Is = "deversoiror"
            Call recopie_deversoir_or(za10)
       End Select
'    End If
'    End If
    ir = ir + 1
Loop
    Close #lhFicDbf
    Close #lhFicDbf1
Exit Sub
test_Error:
   Call print_erreur("fonction recopie_fich  fichier " + nom)
End Sub
Public Function rec_owner() As Form
    Set rec_owner = Me
End Function
Public Sub change_taille()
If gVersionWindow = 5 Then
    h_decal = 880 '750
Else
    h_decal = 750
End If
fcom.Top = 0
If fcom.Name = "Frm_commentaire" Then
    If Me.Height >= h_decal Then
    If Me.Height < 3500 Then Me.Height = 3500
        fcom.Height = Me.Height - h_decal
    End If
End If

If Not fbassin Is Nothing Then
    fdessin.retailler
    fbassin.retailler
    If fcom.Name = "Frm_ss_commentaire" Then
        fcom.Height = fdessin.Height + fbassin.Height
    End If
End If
If Not fobjet Is Nothing Then
    fdessin.retailler
    fobjet.retailler
    If fcom.Name = "Frm_ss_commentaire" Then
        fcom.Height = fdessin.Height + fobjet.Height
    End If
End If
End Sub

Private Sub m_bassin_Click()
Set fdessin = New Frm_dessin
'Set fbassin = New frm_bv2
Set fobjet = New Frm_bv2
    fdessin.Show
    fobjet.Show
    Me.charge_ss_commentaire
End Sub
Private Sub m_decantation_Click()
Set fdessin = New Frm_dessin
Set fobjet = New Frm_decant
    fdessin.Show
    fobjet.Show
    Me.charge_ss_commentaire
End Sub

Private Sub m_do_Click()
Set fdessin = New Frm_dessin
Set fobjet = New Frm_do
    fdessin.Show
    fobjet.Show
    Me.charge_ss_commentaire
End Sub
Private Sub m_DO_or_Click()
Set fdessin = New Frm_dessin
Set fobjet = New Frm_do_or
    fdessin.Show
    fobjet.Show
    Me.charge_ss_commentaire

End Sub



Private Sub m_pompe_Click()
Set fdessin = New Frm_dessin
Set fobjet = New Frm_pompe
    fdessin.Show
    fobjet.Show
    Me.charge_ss_commentaire
End Sub


Private Sub m_quitter_Click()
    Unload Me
End Sub

    

Private Sub m_Retention_Click()
Set fdessin = New Frm_dessin
Set fobjet = New Frm_ret
    fdessin.Show
    fobjet.Show
    Me.charge_ss_commentaire
End Sub

Private Sub m_saisie_info_Click()
'    Me.Enabled = False
    Frm_saisie.Show 1
End Sub

Private Sub m_siphon_Click()
Set fdessin = New Frm_dessin
Set fobjet = New Frm_siphon
    fdessin.Show
    fobjet.Show
    Me.charge_ss_commentaire
End Sub
Private Sub m_conduite_Click()
Set fdessin = New Frm_dessin
Set fobjet = New Frm_conduite
    fdessin.Show
    fobjet.Show
    Me.charge_ss_commentaire
End Sub
Private Sub m_chute_Click()
Set fdessin = New Frm_dessin
Set fobjet = New Frm_chute
    fdessin.Show
    fobjet.Show
    Me.charge_ss_commentaire
End Sub
Private Sub m_stockage_Click()
Set fdessin = New Frm_dessin
Set fobjet = New Frm_stock
    fdessin.Show
    fobjet.Show
    Me.charge_ss_commentaire
End Sub


Private Sub MDIForm_Load()
Dim X As Printer, snp As String
Dim nomworexe As String
'''Dim X As Printer
''For Each X In Printers
''   If X.Orientation = vbPRORPortrait Then
''      ' Définit l'imprimante comme imprimante système
''      ' par défaut.
''      Set Printer = X
''      ' Cesse de rechercher des imprimantes.
''      Exit For
''   End If
''Next
''
Call InfoSysteme

    ok_wor = exist_word()

Couleur_Change = vbBlack 'vbMagenta
larg_mini = 9940 '9000 '9940
haut_mini = 4500 '3750 '3850
l_decal_asc = 200 '200
  maProtectVersion = False    ' version sans contrôle de protection
'  maProtectVersion = True    ' version avec contrôle de protection
'modif FO   ' If ProtectCheck(0) <> 0 Then End
If Printers.count > 0 Then
snp = Printer.DeviceName
Printer.TrackDefault = True
For Each X In Printers
   If X.DeviceName = snp Then
      ' Définit l'imprimante comme imprimante par
      ' défaut du système.
      Set lprinter = X

      Set Printer = X
      ' Cesse la recherche d'imprimante.
      Exit For
   End If
Next
Else
      Set lprinter = Nothing
End If
'coef = Screen.Width / 15360
coef = minimum(Int(Screen.Height / 115.2) / 100#, 1)



'Set lprinter = Printer
''' fichier EXCEL du détail des champs  = c:\hydraulique\bo_v4\defchamps.xls
'''                                 et  = c:\hydraulique\defchamps.xls
'''fichier utilisé = c:\hydraulique\bo_v4\defchamps.mdb
    Me.WindowState = 2 'plein ecran
    Call ini_color
    ouv_save = False
'    nouv_save = False
''''    If Left$(App.Path, 1) = "\" Then
    If Right$(App.Path, 1) = "\" Then
        chemin_app = App.Path
    Else
        chemin_app = App.Path + "\"
    End If
    chemin_init = chemin_app
    nom_fich = ""
    nom_fich_edit = ""
    nomword = ""
    Me.m_hyd.Enabled = False
    Me.m_trait_qualit.Enabled = False
    Me.m_trait_quant.Enabled = False
'    Me.mnusave.Enabled = False
    Me.mnusaves.Enabled = False
    Me.mnusuppr.Enabled = False
    Me.Caption = "HYDROUTI " & App.Major & "." & App.Minor & "." & App.Revision
'''If Dir(chemin_app + "Defchamps.dbf") <> "" Then
'''               xfile1 = FileDateTime(chemin_app + "Defchamps.hyo") 'recup date heure d'un fichier
'''               xfile2 = FileDateTime(chemin_app + "Defchamps.dbf") 'recup date heure d'un fichier
'''    If val(calc_date(xfile2)) < val(calc_date(xfile1)) Then
'''        FileCopy chemin_app + "Defchamps.hyo", chemin_app + "Defchamps.dbf"
'''
'''    End If
'''Else
'''    FileCopy chemin_app + "Defchamps.hyo", chemin_app + "Defchamps.dbf"
'''
'''End If
''''
'    If Dir(chemin_app + "bassin.bin") <> "" _
'        Or Dir(chemin_app + "ouvrages.bin") <> "" _
'        Or Dir(chemin_app + "conduites.bin") <> "" _
'        Or Dir(chemin_app + "ouvrages1.bin") <> "" _
'        Or Dir(chemin_app + "deversoir.bin") <> "" _
'        Or Dir(chemin_app + "retention.bin") <> "" _
'        Or Dir(chemin_app + "siphon.bin") <> "" _
'        Or Dir(chemin_app + "stockage.bin") <> "" Then
'        Me.m_etu.Enabled = False
'        Frm_recopie.Show 1
'        nom_fich_edit = chemin_app + "etude.boa"
'        nom_fich = Left(nom_fich_edit, Len(nom_fich_edit) - 1) + "h"
'''        nom_fich = chemin_app + "etude.boh"
'        Call recopie_fich0(nom_fich_edit)
'        Kill nom_fich
'        nom_fich = ""
'        nom_fich_edit = ""
'    End If
'    nom_fichier_texte = chemin_app + "service.bin"
    text_serv1 = ""
    text_serv2 = ""
    Call lect_fich
    Call ini_bv
    do_bv = False
    door_bv = False
    sto_bv = False
    ret_bv = False
'    chemin_aide = chemin_app + "aide\"
'    chemin_aide = chemin_app + "aide_html\"
    chemin_aide = chemin_app + "html\"
    Set fcom = New Frm_commentaire
    fcom.Top = 0
    fcom.Left = 0
    fcom.Show
  ' récupération de la ligne de commande (ouverture par double click fichier hyd
    Dim sfich As String, c1 As String, scom, i As Integer, lg As Integer
    scom = Command
    If scom <> "" Then
        lg = Len(scom)
        sfich = ""
        For i = 1 To lg
            c1 = Mid$(scom, i, 1)
'            If c1 <> " " And c1 <> vbTab And c1 <> Chr(34) Then
            If c1 <> vbTab And c1 <> Chr(34) Then
               sfich = sfich + c1
            End If
        Next
        Call ouv_etude(sfich)
    End If
    
End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim reponse As Integer
Dim okferme As Boolean
okferme = False
If Not fbassin Is Nothing Then
   okferme = True
    Unload fbassin
End If
If Not fobjet Is Nothing Then
    okferme = True
   Unload fobjet
End If
If fbassin Is Nothing And fobjet Is Nothing Then
'    If nouv_save Then 'And save_fich Then
'        reponse = MsgBox("L'étude n'a pas été enregistrée" + Chr(10) _
'            + "Voulez vous la sauvegarder?", 3, "Sauvegarde d'une étude")
'        Select Case reponse
'            Case Is = 6  ' 6=oui,7=non,2=annuler
'                Call mnusaves_Click
'            Case Is = 7
'                nouv_save = False
'            Case Is = 2
'                Cancel = True
'        End Select
'    End If
    If ouv_save Then
        Call recopie_fich(nom_fich_edit)
        ouv_save = False
    End If
Else
  If okferme Then
    Cancel = True
    End If
End If

Dim X As Printer, snp As String
'snp = Printer.DeviceName

Printer.TrackDefault = True

For Each X In Printers
   If X.DeviceName = lprinter.DeviceName Then ' snp
      ' Définit l'imprimante comme imprimante par
      ' défaut du système.
        Set Printer = X '      Set lprinter = X
      ' Cesse la recherche d'imprimante.
      Exit For
   End If
Next

'Cdlg3.PrinterDefault = true
'Cdlg3.Flags = cdlPDPrintSetup  'Or cdlPDReturnIC   'Or cdlPDReturnDefault
'Cdlg3.CancelError = False
'    Cdlg3.ShowPrinter
'   Set Printer = lprinter
  

End Sub

Private Sub MDIForm_Resize()
   
If Not fbassin Is Nothing Then
    fcom.Top = 0
    fbassin.Top = 0
    fdessin.Top = fbassin.Top + fbassin.Height
End If
If Not fobjet Is Nothing Then
    fcom.Top = 0
    fobjet.Top = 0
    fdessin.Top = fobjet.Top + fobjet.Height
End If
Call change_taille
End Sub
Public Sub calc_kc()
    fobjet.calc_kc
End Sub
Public Sub affich_aide(ByVal nom_form As String, ByVal nom_champ As String)
' If nom_champ <> "" Then
Select Case nom_form
    Case Is = "Frm_do"
                nom_fichier_aide = IDhlpDOFichier
                nom_fichier_exemple = IDhlpDOExempleFichier
    Case Is = "Frm_do_or"
                nom_fichier_aide = IDhlpDOORFichier '"do.rtf"
                nom_fichier_exemple = IDhlpDOORExempleFichier  '"do.rtf"
    Case Is = "Frm_chute"
                nom_fichier_aide = IDhlpChuteFichier
                nom_fichier_exemple = IDhlpChuteExempleFichier
   Case Is = "Frm_pompe"
                nom_fichier_aide = IDhlpPompeFichier
                nom_fichier_exemple = IDhlpPompeExempleFichier
    Case Is = "Frm_conduite"
                nom_fichier_aide = IDhlpConduiteFichier
                nom_fichier_exemple = IDhlpConduiteExempleFichier
    Case Is = "Frm_siphon"
                nom_fichier_aide = IDhlpSiphonFichier
                nom_fichier_exemple = IDhlpSiphonExempleFichier
    Case Is = "Frm_decant"
                nom_fichier_aide = IDhlpDecantationFichier
                nom_fichier_exemple = IDhlpDecantationExempleFichier
    Case Is = "Frm_ret"
                nom_fichier_aide = IDhlpRetentionFichier
                nom_fichier_exemple = IDhlpRetentionExempleFichier
    Case Is = "Frm_stock"
                nom_fichier_aide = IDhlpStockageFichier
                nom_fichier_exemple = IDhlpStockageExempleFichier
    Case Is = "Frm_bv2"
                nom_fichier_aide = IDhlpBVFichier
                nom_fichier_exemple = IDhlpBVExempleFichier
    Case Is = "MDIFrm_menu"
                nom_fichier_aide = IDhlpAideFichier
                nom_fichier_exemple = IDhlpAideExempleFichier
End Select
nom_fichier_aide = chemin_aide + nom_fichier_aide
nom_fichier_exemple = chemin_aide + nom_fichier_exemple
fcom.affich_aide nom_fichier_aide, nom_champ, nom_fichier_exemple
' End If
End Sub
Public Sub recharge_commentaire()
If Not fbassin Is Nothing Then
    Set fbassin = Nothing
End If
If Not fobjet Is Nothing Then
    Set fobjet = Nothing
 
End If
Unload fcom
Set fcom = Nothing
Set fcom = New Frm_commentaire
fcom.Top = 0
fcom.Left = 0
fcom.Show
affich_aide Me.Name, "Objet" '"Structure générale"

End Sub

Public Sub charge_ss_commentaire()
Unload fcom
Set fcom = Nothing
Set fcom = New Frm_ss_commentaire
fcom.Top = 0
fcom.Left = 0
fcom.Show
fcom.Refresh
Me.fdessin.Refresh
Me.MousePointer = vbHourglass
fcom.MousePointer = vbHourglass
If Not fbassin Is Nothing Then
     Me.fbassin.Refresh
    fbassin.Init_ss_commentaire
    fcom.mnu_fichier.Caption = fbassin.mnufichier.Caption
End If
If Not fobjet Is Nothing Then
Me.fobjet.Refresh
    fcom.mnu_fichier.Caption = fobjet.mnufichier.Caption
   fobjet.Init_ss_commentaire
'    fcom.mnu_fichier.Caption = fobjet.mnufichier.Caption
End If
Me.MousePointer = vbArrow
fcom.MousePointer = vbArrow
'fcom.Show

End Sub
Public Sub afficheScrollBars(ByVal ok As Boolean)
Me.fcom.Top = 0
Me.fdessin.Top = 6000
Me.fobjet.Top = 0
End Sub



Private Sub MDIForm_Unload(Cancel As Integer)

    If Trim(nom_fich) <> "" And Dir(nom_fich) <> "" Then
        Call funlockb
        Kill nom_fich
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpContents_Click()
    

    Dim nRet As Integer


    's'il n'y pas de fichier d'aide pour le projet, afficher un message à l'utilisateur
    'vous pouvez définir le fichier d'aide de votre application dans la boîte
    'de dialogue de propriétés du projet
    If Len(App.HelpFile) = 0 Then
        MsgBox "Impossible d'afficher le sommaire de l'aide. Il n'y a pas d'aide associée à ce projet.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        'trifiletti
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub

Private Sub mnuHelpSearch_Click()
    

    Dim nRet As Integer


    's'il n'y pas de fichier d'aide pour le projet, afficher un message à l'utilisateur
    'vous pouvez définir le fichier d'aide de votre application dans la boîte
    'de dialogue de propriétés du projet
    If Len(App.HelpFile) = 0 Then
        MsgBox "Impossible d'affichiez le sommaire de l'aide. Il n'y a pas d'aide associée à ce projet.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub

Private Sub mnunouv_Click()
'Dim reponse As Integer
'If nouv_save Then 'And save_fich Then
'    reponse = MsgBox("L'étude n'a pas été enregistrée" + Chr(10) _
'        + "Voulez vous la sauvegarder?", 3, "Sauvegarde d'une étude")
'    Select Case reponse
'        Case Is = 6  ' 6=oui,7=non,2=annuler
'            Call mnusave_Click
'            nouv_save = False
'            Call nouv_etude
'       Case Is = 7
'            nouv_save = False
'            Call nouv_etude
'    End Select
'Else
'modif FO   ' If ProtectCheck(2) <> 0 Then End
If ouv_save Then
   Call recopie_fich(nom_fich_edit)
   ouv_save = False
    Call nouv_etude
Else
    Call nouv_etude
End If
End Sub
Private Sub mnuouv_Click()
'Dim reponse As Integer
'If nouv_save Then 'And save_fich Then
'    reponse = MsgBox("L'étude n'a pas été enregistrée" + Chr(10) _
'        + "Voulez vous la sauvegarder?", 3, "Sauvegarde d'une étude")
'    Select Case reponse
'        Case Is = 6  ' 6=oui,7=non,2=annuler
'            Call mnusave_Click
'            nouv_save = False
'            Call ouv_etude
'       Case Is = 7
'            nouv_save = False
'            Call ouv_etude
'    End Select
'Else
'modif FO   ' If ProtectCheck(2) <> 0 Then End
If ouv_save Then
   '     Debug.Print "debut recopie"
        Call recopie_fich(nom_fich_edit)
'        Debug.Print "fin recopie"
        ouv_save = False
        Call ouv_etude("")
Else
    Call ouv_etude("")
End If
End Sub
    
Private Sub ouv_etude(ByVal sfichier As String)
Dim reponse As Integer
Dim message As String
Dim fs As Object
Dim s As String
Dim fsco As file_spec
Dim f As File
Dim d As Drive
Dim nom As String
''If nom_fich <> "" Then
''    Call recopie_fich
''End If
'If nom_fich <> "" Then
'    If Dir(nom_fich) <> "" Then
'        Call funlockb
'        Kill nom_fich
'    End If
'End If
'
On Error GoTo test_Error

If sfichier = "" Then
    cdlg1.DialogTitle = "Recherche d'une étude "
    cdlg1.FileName = ""
    cdlg1.Filter = "Fichiers HYDROUTI (*.hyd)|*.hyd"
    cdlg1.InitDir = chemin_init
    cdlg1.Flags = cdlOFNHideReadOnly Or cdlOFNFileMustExist
    cdlg1.ShowOpen
    s = cdlg1.FileName
Else
    s = sfichier
End If

If s <> "" Then
If nom_fich <> "" Then
    If Dir(nom_fich) <> "" Then
        Call funlockb
        Kill nom_fich
    End If
End If

fsco = create_fs(s)
    If fsco.dr_type = 1 Then
        message = "Fichier sur disquette;" + Chr(13) + Chr(10) + "Vérifier que la disquette n'est pas protégée en écriture."
        reponse = MsgBox(message, , "Saisie du nom de l'étude")
    End If
    If fsco.dr_type = 4 Then
        message = "Fichier sur CR-ROM;" + Chr(13) + Chr(10) + "Pas d'accés en écriture."
        reponse = MsgBox(message, , "Saisie du nom de l'étude")
        nom_fich_edit = ""
    ElseIf fsco.lecteur <> "" And fsco.Chemin <> "" Then
        If Trim(fsco.nom) <> "" Then
            If fsco.f_attr = 1 Or fsco.f_attr = 33 Then
                message = "Fichier en lecture seule."
                reponse = MsgBox(message, , "Saisie du nom de l'étude")
                nom_fich_edit = ""
            Else
                nom_fich_edit = Trim(fsco.nomcomplet)
            End If
        Else
            nom_fich_edit = ""
        End If
    Else
            nom_fich_edit = ""
    End If
    nom_etude = Left(fsco.nom, Len(fsco.nom) - 4)

    If nom_fich_edit <> "" Then
        message = ""
        If Right(nom_fich_edit, 4) <> ".hyd" Then
            nom_fich_edit = Left(nom_fich_edit, Len(nom_fich_edit) - 4) + ".hyd"
        End If
        If fsco.f_size > 0 Then
             message = verif_etude()
        End If
        If message = "" Then
            nom_fich = Left(nom_fich_edit, Len(nom_fich_edit) - 3) + "boh"  'chemin_app + "hydrouti.boh"
            If Dir(nom_fich) <> "" Then
                Kill nom_fich
            End If
            If fsco.f_size = 0 Then
                 Call init_etude
            Else
                Call recup_fich(nom_fich, nom_fich_edit)
            End If
             DoEvents
            Call flocka(nom_fich_edit)
            Call flockb(nom_fich)
            chemin_init = fsco.Chemin
            chemin_etude = fsco.lecteur + fsco.Chemin + "\"
            
             Me.Caption = "HYDROUTI " & App.Major & "." & App.Minor & "." & App.Revision & " -- Etude " & nom_fich_edit
     '      Me.Caption = "HYDROUTI -- Etude " + nom_fich_edit
            Me.m_hyd.Enabled = True
            Me.m_trait_qualit.Enabled = True
            Me.m_trait_quant.Enabled = True
'            Me.mnusave.Enabled = True
            Me.mnusaves.Enabled = True
            Me.mnusuppr.Enabled = True
            ouv_save = True

'           Call recopie_fich
        Else
            reponse = MsgBox(message, , "Saisie du nom de l'étude")
            Me.m_hyd.Enabled = False
            Me.m_trait_qualit.Enabled = False
            Me.m_trait_quant.Enabled = False
'            Me.mnusave.Enabled = False
            Me.mnusaves.Enabled = False
            Me.mnusuppr.Enabled = False
            ouv_save = False
        End If
    Else
            Me.m_hyd.Enabled = False
            Me.m_trait_qualit.Enabled = False
            Me.m_trait_quant.Enabled = False
'            Me.mnusave.Enabled = False
            Me.mnusaves.Enabled = False
            Me.mnusuppr.Enabled = False
            ouv_save = False
    End If
End If
Exit Sub
test_Error:
    If Err.Number = 70 Then
        Call print_erreur("Le fichier " + s + " est déjà en cours d'utilisation.")
    End If
    Unload Me
End Sub
Private Sub nouv_etude()
Dim reponse As Integer
Dim message As String
Dim fs As Object
Dim s As String
Dim fsco As file_spec
Dim fsce As file_spec
Dim f As File
Dim d As Drive
Dim nom As String
'If nom_fich <> "" Then
'    Call recopie_fich
'End If
''''    If nom_fich <> "" Then
''''        If Dir(nom_fich) <> "" Then
''''            Call funlockb
''''            Kill nom_fich
''''        End If
''''    End If
cdlg1.DialogTitle = "Nouvelle étude "
cdlg1.FileName = ""
cdlg1.Filter = "Fichiers HYDROUTI (*.hyd)|*.hyd"
'If nom_fich_edit <> "" Then
'    fsce = create_fs(nom_fich_edit)
'    cdlg1.FileName = fsce.nom
'    cdlg1.InitDir = chemin_init
'Else
    cdlg1.InitDir = chemin_init
'End If
'cdlg1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
'cdlg1.ShowSave
cdlg1.Flags = cdlOFNHideReadOnly 'Or cdlOFNFileMustExist
cdlg1.ShowOpen
s = cdlg1.FileName
If s <> "" Then
     If nom_fich <> "" Then
        If Dir(nom_fich) <> "" Then
            Call funlockb
            Kill nom_fich
        End If
    End If
   fsco = create_fs(s)
'Debug.Print "chemin++", fsco.Chemin, "++"
'Debug.Print "lecteur==", fsco.lecteur, "++"
'Debug.Print fsco.nom
'Debug.Print fsco.extension
'Debug.Print "+++", fsco.nomcomplet, "+++"
   nom_etude = Left(fsco.nom, Len(fsco.nom) - 4)
   chemin_init = fsco.Chemin
'    Call recopie_fich(fsco.nomcomplet)

    If Dir(fsco.nomcomplet) <> "" Then
         message = "Le fichier existe dèjà !!!"
        reponse = MsgBox(message, , "Saisie du nom d'une nouvelle étude")
    Else
        If fsco.dr_type = 4 Then
        message = "Fichier sur CR-ROM;" + Chr(13) + Chr(10) + "Pas d'accés en écriture."
        reponse = MsgBox(message, , "Saisie du nom d'une nouvelle étude")
        Else
            If fsco.dr_type = 1 Then
            message = "Fichier sur disquette;" + Chr(13) + Chr(10) + "Vérifier que la disquette n'est pas protégée en écriture."
            reponse = MsgBox(message, , "Saisie du nom d'une nouvelle étude")
            End If

            nom_fich_edit = fsco.nomcomplet
            nom_fich = Left(nom_fich_edit, Len(nom_fich_edit) - 3) + "boh" 'chemin_app + "hydrouti.boh"
            If Dir(nom_fich) <> "" Then
                Kill nom_fich
            End If
            Call init_etude
            Call flockb(nom_fich)
'           Me.Caption = " HYDROUTI -- Nouvelle étude "
              Me.Caption = "HYDROUTI " & App.Major & "." & App.Minor & "." & App.Revision & " -- Etude " & nom_fich_edit
'           Me.Caption = "HYDROUTI -- Etude " + nom_fich_edit
            Me.m_hyd.Enabled = True
            Me.m_trait_qualit.Enabled = True
            Me.m_trait_quant.Enabled = True
'           Me.mnusave.Enabled = True
            Me.mnusaves.Enabled = True
            Me.mnusuppr.Enabled = True
'           nouv_save = True
            ouv_save = True
        End If
    End If
End If
End Sub

Private Sub mnuprint_Click()
Dim a As Variant
On Error GoTo erreur:
'modif FO   ' If ProtectCheck(2) <> 0 Then End
Printer.TrackDefault = True
cdlg1.PrinterDefault = True
cdlg1.Flags = cdlPDPrintSetup 'Or cdlPDReturnIC   'Or cdlPDReturnDefault
cdlg1.CancelError = True
cdlg1.ShowPrinter
While Printer.Orientation = cdlLandscape
    MsgBox "l'impression doit se faire en mode portrait", vbExclamation, _
        "Configuration imprimante"
    cdlg1.CancelError = True
    cdlg1.ShowPrinter
Wend
'Debug.Print Printer.DeviceName
'Printer.TrackDefault = True
'For i = 0 To Printers.Count - 1
'        Set Printer = Printers(i)
' Debug.Print Printer.hDC, Cdlg1.hDC, Printers(i).TrackDefault
'Next
'For i = 0 To Printers.Count - 1
'' '  If Printers(i).hDC = Cdlg1.hDC Then
'  If Printers(i).DeviceName = "HP LaserJet 4/4M" Then
''      ' Définit l'imprimante comme imprimante par
''      ' défaut du système.
'      Set Printer = Printers(i)
'      Debug.Print Printer.hDC
''      ' Cesse la recherche d'imprimante.
'      Exit For
'   End If
'Next
'Debug.Print Printer.DeviceName
'Cdlg1.Flags = cdlPDPrintSetup Or cdlPDReturnDC   'Or cdlPDReturnDefault
'Cdlg1.hDC = Printer.hDC
'Cdlg1.ShowPrinter
Exit Sub
erreur:
'bannul = True
Resume Next
End Sub


Private Sub mnusave0_Click()
    If Trim(nom_fich_edit) <> "" Then
        Call recopie_fich(nom_fich_edit)
        ouv_save = False
    Else
        Call mnusaves_Click
    End If

End Sub

Private Sub mnusaves_Click()
Dim reponse As Integer
Dim message As String
Dim fs As Object
Dim s As String
Dim fsco As file_spec
Dim fsce As file_spec
Dim f As File
Dim d As Drive
'modif FO   ' If ProtectCheck(2) <> 0 Then End
cdlg1.DialogTitle = "Enregistrer l'étude sous "
cdlg1.FileName = ""
cdlg1.Filter = "Fichiers HYDROUTI (*.hyd)|*.hyd"
If nom_fich_edit <> "" Then
fsce = create_fs(nom_fich_edit)
cdlg1.FileName = fsce.nom
cdlg1.InitDir = chemin_init
Else
cdlg1.InitDir = chemin_init
End If
cdlg1.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
cdlg1.ShowSave
Print cdlg1.Tag
s = cdlg1.FileName
If s <> "" Then
fsco = create_fs(s)
'Debug.Print "chemin++", fsco.Chemin, "++"
'Debug.Print "lecteur==", fsco.lecteur, "++"
'Debug.Print fsco.nom
'Debug.Print fsco.extension
'Debug.Print "+++", fsco.nomcomplet, "+++"
    chemin_init = fsco.Chemin
    Call recopie_fich(fsco.nomcomplet)
    nom_fich_edit = fsco.nomcomplet
            If Dir(nom_fich) <> "" Then
                Kill nom_fich
            End If
            nom_fich = Left(nom_fich_edit, Len(nom_fich_edit) - 3) + "boh"  'chemin_app + "hydrouti.boh"
            If Dir(nom_fich) <> "" Then
                Kill nom_fich
            End If
                Call recup_fich(nom_fich, nom_fich_edit)
            Call flocka(nom_fich_edit)
            Call flockb(nom_fich)
             Me.Caption = "HYDROUTI " & App.Major & "." & App.Minor & "." & App.Revision & " -- Etude " & nom_fich_edit
'   Me.Caption = " HYDROUTI -- Etude " + nom_fich_edit
    Me.mnusuppr.Enabled = True
    ouv_save = False
'    nouv_save = False
End If
End Sub

Private Sub mnusuppr_Click()
Dim reponse As Integer
'modif FO   ' If ProtectCheck(2) <> 0 Then End
' 2004/03/16
    reponse = MsgBox("Confirmer la suppression de l'étude: " + nom_fich_edit, 3, "Suppression d'une étude")
    Select Case reponse
        Case Is = 6  ' 6=oui,7=non,2=annuler
' 2003/04/04
            Call funlocka
            Kill nom_fich_edit
            Call funlockb
            Kill nom_fich
            nom_fich_edit = ""
            nom_fich = ""
            ouv_save = False
'    nouv_save = False
            Me.mnusaves.Enabled = False
            Me.mnusuppr.Enabled = False
            Me.m_hyd.Enabled = False
            Me.m_trait_qualit.Enabled = False
            Me.m_trait_quant.Enabled = False
            Me.Caption = "HYDROUTI " & App.Major & "." & App.Minor & "." & App.Revision
'    Me.Caption = "HYDROUTI"
       Case Is = 7
    End Select

End Sub
Sub InfoSysteme()
Dim WinVer As MYVERSION
    WinVer = WindowsVersion()
    gVersionWindow = WinVer.lMajorVersion
End Sub


