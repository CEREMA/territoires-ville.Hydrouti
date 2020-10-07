Attribute VB_Name = "Global"
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long

Public Const ERROR_SUCCESS = 0&
Public Const ERROR_NO_MORE_ITEMS = 259&
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CLASSES_ROOT = &H80000000
'***
Public Type points
    X As Double
    Y As Double
End Type
'***
Public Type typ_Couleur
    blanc  As ColorConstants
    noir  As ColorConstants
    gris  As ColorConstants
    gris_clair  As ColorConstants
    rouge_clair  As ColorConstants
    rouge As ColorConstants
    orange_clair  As ColorConstants
    orange As ColorConstants
    vert_clair  As ColorConstants
    vert  As ColorConstants
    jaune_clair  As ColorConstants
    jaune  As ColorConstants
    cyan_clair  As ColorConstants
    cyan  As ColorConstants
    bleu_clair  As ColorConstants
    bleu  As ColorConstants
    magenta_clair  As ColorConstants
    magenta  As ColorConstants
End Type
Public Type defchamp
    Form As String
    Intitule As String
    Ancnom As String
    Nomchp As String
    Indexc As Integer
    taille As Integer
    Decimal As Integer
    OKmini As Boolean
    Mini As Double
    OKmaxi As Boolean
    Maxi As Double
    message As String
    Chplabel As String
    Label As String
    Chpunite As String
    Unite As String
End Type
Public Type stock_dess
    type As String * 4
    opt_long As Boolean
    opt_larg As Boolean
    opt_prof As Boolean
    opt_rap As Boolean
    Longueur As Double
    Largeur As Double
    Profondeur As Double
    Rapport As Double
    Diametre As Double
    hauteur As Double
    Diametrec As Double
    Longueurc As Double
    coef As Double
End Type
Public Type st_Stock
    nom As String * 10
    nombv As String * 30
    Qpluie As Double
    Qts As Double
    Qrin As Double
    lcrin As Double
    surface As Double
    imper As Integer
    tc As Double
    Qav As Double
    Ipcav As Double
    Vr As Double
    alphat As Double
    volume As Double
    dessstock As stock_dess
End Type
Public Type st_savstock
    type As String * 15
    nom As String * 30
    stockage As st_Stock
    reste As String * 50
End Type
Public Type st_savsto1
    stsavstock As st_savstock
    reste As String * 441
End Type
Public Type courbe_dess
    duree As Double
    quantite As Double
    volume As Double
    hauteur As Double
End Type
Public Type volume_dess
    coef As Double
    Diametre As Double
    hauteur As Double
    Diametrec As Double
    Longueur As Double
    Longueurc  As Double
    Largeur As Double
    Profondeur As Double
    Rapport As Double
End Type
Public Type ret_dess
    type As String * 4
    opt_long As Boolean
    opt_larg As Boolean
    opt_prof As Boolean
    opt_rap As Boolean
    Longueur As Double
    Largeur As Double
    Profondeur As Double
    Rapport As Double
    coef As Double
    duree As Double
    Hpluie As Double
    Hfuite As Double
End Type
Public Type st_Ret
    nom As String * 10
    nombv As String * 30
    type_calcul As String * 1
    surface As Double
    Ca As Integer
    qf As Double
    amontana As Double
    bmontana As Double
    deltaH As Double
    volume As Double
    a1montana As Double
    b1montana As Double
    Seuil As Double
    desssret As ret_dess
End Type
Public Type st_savret
    type As String * 15
    nom As String * 30
    retention As st_Ret
    reste As String * 50
End Type
Public Type st_savret1
    stsavret As st_savret
    reste As String * 464   '488
End Type
Public Type st_Bv
    nom As String * 10
    type As String * 1
    surface As Double
    imper As Integer
    lghydr As Double
    phydr As Integer
    nhab As Double
    tdilu As Double
    ceau As Double
    perti As Double
    vinf As Double
    ahorton As Double
    bhorton As Double
    trep As Double
    Qbrut As Double
    Qcor As Double
    Qmr As Double
    Qhydro As Double
    Qeu As Double
    Qecp As Double
    Qts As Double
    Qprin As Double
    Qrin As Double
    tc As Double
    qfuite As Double
    pas As Double
    Teta As Double
    Qchoisi As String * 6
    
    
End Type
Public Type st_ParHydro
'    nom As String * 30
    amontana As Double
    bmontana As Double
    lcrin As Double
    ceau As Double
    aeu As Double
    beu As Double
    a1montana As Double
    b1montana As Double
    Seuil As Double
End Type
Public Type st_ResHydro
'    nom As String * 30
    dt As Double
    DM As Double
    HT As Double
    HM As Double
    Teta As Double
    pas As Double
   
End Type
Public Type st_hydr
    DM As Double
    dt As Double
    HM As Double
    HT As Double
    pas As Double
    Teta As Double
    kdesbor As Double
    qfuite As Double
    vst As Double
    vstock As Double
End Type
Public Type st_Coude
    Nbre As Double
    type As String * 9
    angle As Double
    Rayon As Double
End Type
Public Type st_listcoude
    coude(9) As st_Coude
End Type

Public Type st_save 'bassin versant
'    type As String * 15
    nom As String * 30
    bv As st_Bv
    hydro As st_ParHydro
    hydro1 As st_hydr
'    reste As String * 50
End Type
Public Type st_save1 'bassin versant
    type As String * 15
    stsave As st_save
    reste As String * 348
'    reste As String * 372
End Type
Public Type st_ParHydroa
'    nom As String * 30
    amontana As Double
    bmontana As Double
    lcrin As Double
    ceau As Double
    aeu As Double
    beu As Double
End Type
Public Type st_savea 'bassin versant
'    type As String * 15
    nom As String * 30
    bv As st_Bv
    hydro As st_ParHydroa
    hydro1 As st_hydr
'    reste As String * 50
End Type
Public Type st_save1a 'bassin versant
    type As String * 15
    stsave As st_savea
    reste As String * 372
End Type

Public Type troncon
    Absamo As Double
    radamo As Double
    Absava As Double
    radava As Double
    conduit As conduite
End Type

Public Type st_dessdo
    nom As String * 10
    nombv As String * 30
    Qts As Double
    Qrin As Double
    Qpluie As Double
    rdoam As Double
    rdoav As Double
    lgdisp As Double
    phex As Double
    rdoex As Double
    lgca As Double
    dam As Double
    iRadam As Double
    Kam As Double
    Lam As Double
    dav As Double
    iradav As Double
    kav As Double
    Lav  As Double
    Tram As Double
    Centon As Double
    tron_amo As troncon
    tron_ava As troncon
    tron_dech As troncon
End Type
Public Type deversoir
    Absamo As Double
    radamo As Double
    Absava As Double
    radava As Double
    Longueur As Double
    pente As Double
    hauteur As Double
    tron_ava  As troncon
    tav As Double
End Type
Type resu_intdev
'    dam As String  'Diamètre
'    iRadam As String 'Pente
'    Kam As String 'Coefficient de Manning-Strickler
'    Lam As String 'Longueur
'    dav As String 'Diamètre
'    iradav As String 'Pente
'    kav As String 'Coefficient de Manning-Strickler
'    Lav  As String 'Longueur
'    ddech As String 'Diamètre
'    iraddech As String 'Pente
'    kdech As String 'Coefficient de Manning-Strickler
'    Ldech  As String 'Longueur
    Ldev As String ' Longueur du DO
    Hcret As String ' Hauteur de la crête
    Pdev As String ' Pente du DO
    Tram As String ' Tirant d'eau amont admissible
    dpsm As String 'Débit pleine section
    vpsm As String 'Vitesse pleine section
    vqtsm As String 'Vitesse d'écoulement à QTS
    hqtsm As String 'Hauteur d'eau QTS
    vqrinm As String 'Vitesse d'écoulement à QRIN
    hqrinm As String 'Hauteur d'eau QRIN
    vqpluiem As String 'Vitesse d'écoulement amontà QPLUIE
    hqpluiem As String 'Hauteur d'eau amont QPLUIE
    vqpluiemav As String 'Vitesse d'écoulement aval à QPLUIE
    hqpluiemav As String 'Hauteur d'eau aval QPLUIE
    dpsv As String 'Débit pleine section
    vpsv As String 'Vitesse pleine section
    vqtsv As String 'Vitesse d'écoulement à QTS
    hqtsv As String 'Hauteur d'eau QTS
    vqrinv As String 'Vitesse d'écoulement à QRIN
    hqrinv As String 'Hauteur d'eau QRIN
    vqpluiev As String 'Vitesse d'écoulement à QPLUIE
    hqpluiev As String 'Hauteur d'eau QPLUIE
    longetranglee As String ' Longueur conduite étranglée"
    debetranglee As String 'Débit dans la conduite étranglée
    debdeverse As String 'Débit déversé"
    dpsdech As String 'Débit pleine section décharge
    vpsdech As String 'Vitesse pleine section décharge
    vqdev As String 'Vitesse d'écoulement amont pour débit déversé
    hqdev As String 'Hauteur d'eau amont pour débit déversé
    vqdevav As String 'Vitesse d'écoulement aval pour débit déversé
    hqdevav As String 'Hauteur d'eau aval pour débit déversé
    regime As String
    Ham As String 'Hauteur de la lame d'eau amont
    Hav As String 'Hauteur de la lame d'eau aval
    Haam As String 'Hauteur de la charge amont
    Haav As String 'Hauteur de la charge aval
End Type
Type resu_dev
'    dam As String  'Diamètre
'    iRadam As String 'Pente
'    Kam As String 'Coefficient de Manning-Strickler
'    Lam As String 'Longueur
'    dav As String 'Diamètre
'    iradav As String 'Pente
'    kav As String 'Coefficient de Manning-Strickler
'    Lav  As String 'Longueur
'    ddech As String 'Diamètre
'    iraddech As String 'Pente
'    kdech As String 'Coefficient de Manning-Strickler
'    Ldech  As String 'Longueur
    Ldev As String ' Longueur du DO
    Hcret As String ' Hauteur de la crête
    Pdev As String ' Pente du DO
    Tram As String ' Tirant d'eau amont admissible
    dpsm As String 'Débit pleine section
    vpsm As String 'Vitesse pleine section
    vqtsm As String 'Vitesse d'écoulement à QTS
    hqtsm As String 'Hauteur d'eau QTS
    vqrinm As String 'Vitesse d'écoulement à QRIN
    hqrinm As String 'Hauteur d'eau QRIN
    vqpluiem As String 'Vitesse d'écoulement à QPLUIE
    hqpluiem As String 'Hauteur d'eau QPLUIE
    vqpluiemav As String 'Vitesse d'écoulement aval à QPLUIE
    hqpluiemav As String 'Hauteur d'eau aval QPLUIE
    dpsv As String 'Débit pleine section
    vpsv As String 'Vitesse pleine section
    vqtsv As String 'Vitesse d'écoulement à QTS
    hqtsv As String 'Hauteur d'eau QTS
    vqrinv As String 'Vitesse d'écoulement à QRIN
    hqrinv As String 'Hauteur d'eau QRIN
    vqpluiev As String 'Vitesse d'écoulement à QPLUIE
    hqpluiev As String 'Hauteur d'eau QPLUIE
    longetranglee As String 'Longueur conduite étranglée
    debetranglee As String 'Débit dans la conduite étranglée
    debdeverse As String 'Débit déversé"
    dpsdech As String 'Débit pleine section
    vpsdech As String 'Vitesse pleine section
    vqdev As String 'Vitesse d'écoulement pour débit déversé
    hqdev As String 'Hauteur d'eau pour débit déversé
    vqdevav As String 'Vitesse d'écoulement aval pour débit déversé
    hqdevav As String 'Hauteur d'eau aval pour débit déversé
    regime As String
    Ham As String 'Hauteur de la lame d'eau amont
    Hav As String 'Hauteur de la lame d'eau aval
    Haam As String 'Hauteur de la charge amont
    Haav As String 'Hauteur de la charge aval
End Type
Type resu_udev
'    dam As String  'mm
'    iRadam As String '1/10000
'    Kam As String '
'    Lam As String 'm
'    dav As String 'mm
'    iradav As String '1/10000
'    kav As String '
'    Lav  As String 'm
'    ddech As String 'mm
'    iraddech As String '1/10000
'    kdech As String '
'    Ldech  As String 'm
    Ldev As String ' m
    Hcret As String ' m
    Pdev As String ' m/m
    Tram As String ' m
    dpsm As String 'm3/s
    vpsm As String 'm/s
    vqtsm As String 'm/s
    hqtsm As String ' m
    vqrinm As String 'm/s
    hqrinm As String ' m
    vqpluiem As String 'm/s
    hqpluiem As String ' m
    vqpluiemav As String 'm/s
    hqpluiemav As String ' m
    dpsv As String 'm3/s
    vpsv As String 'm/s
    vqtsv As String 'm/s
    hqtsv As String ' m
    vqrinv As String 'm/s
    hqrinv As String ' m
    vqpluiev As String 'm/s
    hqpluiev As String ' m
    longetranglee As String ' m
    debetranglee As String 'm3/s
    debdeverse As String 'm3/s
    dpsdech As String 'm3/s
    vpsdech As String 'm/s
    vqdev As String 'm/s
    hqdev As String 'm
    vqdevav As String 'm/s
    hqdevav As String 'm
    regime As String
    Ham As String 'm
    Hav As String 'm
    Haam As String 'm
    Haav As String 'm
End Type
Public Type devor_courbe
    dx(51) As Double
    dy(51) As Double
End Type
Public Type deversoiror_resultat
    Ham_cri As Double
    Vam_cri As Double
    hc_cri As Double
    vc_cri As Double
    hav_cri As Double
    Ham As Double
    Vam As Double
    hc As Double
    vc As Double
    Hbav As Double
    Qbavth As Double
    Qbaveff As Double
    hmin As Double
    Hav As Double
    hdev As Double
    alpha As Double
    CosA As Double
    deltaa As Double
    l_ouverture As Double
    l_largOuverture As Double
    nbFroude As Double
    nbFroudeMax As Double
    l_chambre1 As Double
    l_jetaval_b As Double
    l_jetaval_h As Double
End Type
Public Type deversoir_resultat
    Tram As Double
    HM As Double
    Ham As Double
    Hav As Double
    Haav As Double
    Haavd As Double
    Haam As Double
    a As Double
    c As Double
    Qav As Double
    Qdev As Double
End Type

Public Type st_savdo
    type As String * 15
    nom As String * 30
    edessdo As st_dessdo
    edo As deversoir
    reste As String * 50
End Type
Public Type st_savdo1
    stsavdo As st_savdo
    reste As String * 135
End Type
Public Type st_savdoor
    type As String * 15
    nom As String * 30
    edessdo As st_dessdo
    edo As deversoir
    reste As String * 50
End Type
Public Type st_savdoor1
    stsavdoor As st_savdoor
    reste As String * 135
End Type
Public Type Resudo
    ldav As Double
    pdav As Double
    ddav As Double
    dlongdo As Double
    dpentedo As Double
End Type

Public Type st_Siphon
    dam As Double
    iRadam As Double
    Kam As Double
    dav As Double
    iradav As Double
    kav As Double
    Rdav As Double
    Rdam As Double
    Jadm As Double
    ds As Double
    Ks As Double
    Qmax As Double
    ls As Double
    Kc As Double
    List_coude As st_listcoude
    Ipl As Double
    deltaH1 As Double
    deltaH2 As Double
    IPs As Double
    tron_amo As troncon
    tron_ava As troncon
End Type
Public Type st_savsi
    type As String * 15
    nom As String * 30
    siphon As st_Siphon
    reste As String * 50
End Type
Public Type st_savsi1
    stsavsi As st_savsi
    reste As String * 49
End Type
Public Type pompe_car
    qeum As Double 'débit moyen des eaux usées
    Fp As Double 'facteur de pointe
    Qeu As Double 'débit de pointe des eaux usées
    Qecp As Double 'débit des eaux parasites
    Qtsm As Double 'débit moyen de temps sec
    Qts As Double 'débit de pointe de temps sec
    Qpomp As Double 'débit de pompage théorique (proposé)

End Type
Public Type pompe_geo
    Lrflt As Double 'longueur de refoulement
    NatRflt As String * 5 'nature du tuyau
    Drflt As Double 'diametre théorique
    NivTN As Double 'niveau TN
    NivEN As Double 'niveau fil d'eau arrivée
    NivSO As Double 'niveau fil d'eau sortie
    NivEX As Double 'niveau fil d'eau extrémité refoulement
End Type
Public Type pompe_sing
    Nbc1 As Integer 'Nb coudes 1
    Nbc2 As Integer 'Nb coudes 2
    Nbc3 As Integer 'Nb coudes 3
    Nbc4 As Integer 'Nb coudes 4
    Nbc9 As Integer 'Nb coudes 9
    Nbva As Integer 'Nb vannes
    Nbcl As Integer 'Nb clapets
    Nbvi As Integer 'Nb vidanges
    Nbve As Integer 'Nb ventouses
    Antb As Integer  'valeur anti belier
End Type
Public Type pompe_tech
    Nbpom As Integer 'nombre de pompes
    Ntdph As Integer 'nombre de démarrages
    Vutba As Double 'volume utile théorique de la bâche
    Sectb As Integer 'valeur section bâche
    Diamb As Double 'diamétre de la section de la bâche
    Longb As Double 'longueur de la section de la bâche
    Largb As Double 'largeur de la section de la bâche
    Denivt As Double 'dénivelé retenu
    Denivhau As Double 'garde à l'égout
    Denivbas As Double 'garde au fond
End Type
Public Type pompe_resu
    Qpomr As Double 'débit de pompage retenu
    Drflr As Double 'diametre retenu
    NatRflr As String * 11 'nature du tuyau retenue
    VitRflt As Double 'vitesse en régime permanentJmpkm
    jmpkm As Double 'pertes de charges linéaires
    Denivr As Double  'dénivelé retenu
    Vurba As Double 'volume utile retenu de la bâche
    Nrdph As Double 'nombre reel de démarrage
    Tvidange As Double 'temps de vidange
    T1cyc As Double 'durée du cycle
    Nbcyc As Double 'nombre de cycles
    Vmy As Double 'vitesse moyenne d'écoulement
    Tsejh As Double 'temps de séjour
    Singul As Double 'pertes de charges singulières
    Hmt As Double 'hauteur manométrique totale
End Type
Public Type st_Pompe
    debits_car As pompe_car
    don_geometrie As pompe_geo
    pts_singuliers As pompe_sing
    don_techniques As pompe_tech
    resultat As pompe_resu
End Type
Public Type st_savpompe
    type As String * 15
    nom As String * 30
    pompe As st_Pompe
    reste As String * 50
End Type
Public Type st_savpom1
    stsavpo As st_savpompe
    reste As String * 341  '(à voir l total=750 voir l pompe dans save)
End Type

Public Type st_Chute
    dam As Double
    iRadam As Double
    Kam As Double
    dav As Double
    iradav As Double
    kav As Double
    Rdav As Double
    Rdam As Double
    Qmax As Double
    h0 As Double
    Long As Double
    tron_amo As troncon
    tron_ava As troncon
    
End Type
Public Type st_savchute
    type As String * 15
    nom As String * 30
    chute As st_Chute
    reste As String * 50
End Type
Public Type st_savch1
    stsavch As st_savchute
    reste As String * 435
End Type
Public Type st_Decant
    Q As Double
    d As Double
    X As Double
    Psed As Double
    Vhor As Double
    Long As Double
    larg As Double
    Hchamb As Double
    heau As Double
    Vvert As Double
    k As Double
End Type
Public Type st_savdecant
    type As String * 15
    nom As String * 30
    decant As st_Decant
    reste As String * 50
End Type
Public Type st_savdec1
    stsavdecant As st_savdecant
    reste As String * 567
End Type
Public Type st_hydrouti
    type As String * 15
    reste As String * 735
End Type
Public Type debit_conduit
    charge As Boolean
    debit As Double
    vitesse As Double
    hauteur As Double
    largeurlibre   As Double
    surface As Double
    acceleration As Double
    pentemotrice As Double
    piezoamo As Double
    piezoava As Double
    piezointer As points
    piezointer0 As points
    piezointer1 As points
    piezointer2 As points
    dcharge As Double
    chargeamo As Double
    chargeava As Double
    chargeinter As Double
    chargeinter0 As Double
    chargeinter1 As Double
    chargeinter2 As Double
    zphe_ava As Double
    hautamo As Double
    hautava As Double
    vitamo As Double
    vitava As Double
    zeau_amo As points
    zeau_ava As points
    p_Eau_inter As points
    p_Eau_inter0 As points
    p_Eau_inter1 As points
    p_Eau_inter2 As points

End Type
Public Type st_texte
    nom1 As String * 60
    nom2 As String * 60
End Type
Public lhFicDba As Integer
Public lhFicDbb As Integer
Public lhFicooo As Integer
Public lhFicooo1 As Integer

Global annul As Boolean
Global couleur As typ_Couleur
Global edessdo As st_dessdo
Global edo As deversoir
Global edoor_res As deversoiror_resultat
Global edoor_courbe_max_dever As devor_courbe
Global edoor_courbe_max_bas As devor_courbe
Global edoor_courbe_max_haut As devor_courbe
Global edoor_courbe_cri_bas As devor_courbe
Global edoor_courbe_cri_haut As devor_courbe
Global edo_res As deversoir_resultat
Global resudev As resu_dev
Global ehyd As st_hydr
Global ebv As st_Bv
Global eph As st_ParHydro
Global ebstock As st_Stock
Global ebret As st_Ret
Global ebdecant As st_Decant
Global ebsiphon As st_Siphon
Global ebchute As st_Chute
Global ebpompe As st_Pompe
Global Listcoud As st_listcoude
Global chemin_app As String
Global chemin_OOO As String
Global chemin_etude As String
Global nom_etude As String
Global chemin_fiche As String
Global nom_fich As String
Global nom_fich_edit As String
Global fich_lect As String
Global fich_lect_edit As String
Global gnom_type As String
Global do_bv As Boolean
Global door_bv As Boolean
Global sto_bv As Boolean
Global ret_bv As Boolean
Global nomprojet As String
Global nomword As String
Global nomooo As String
Global ouv_sauve As Boolean
Global save_fich As Boolean
Global ok_tooltip As Boolean
Global opt_cli As Boolean
Global dess_anc As String
Global text_serv1 As String * 60
Global text_serv2 As String * 60
Global Hpluie() As Variant, Q() As Variant 'hydro
Global awd As Word.Application
Global Couleur_Change As ColorConstants
Global larg_mini As Integer
Global haut_mini As Integer
Global l_decal_asc As Integer
Global gVersionWindow As Integer
Global Index As Integer
Global ok_wor As Boolean
Public Sub ini_color()
couleur.blanc = "&H00FFFFFF"
couleur.noir = "&H00000000"
couleur.gris = "&H00808080"
couleur.gris_clair = "&H00E0E0E0"
couleur.rouge = "&H000000FF"
couleur.rouge_clair = "&H00C0C0FF"
couleur.vert = "&H0000C000"
couleur.vert_clair = "&H0080FF80"
couleur.orange = "&H000080FF"
couleur.orange_clair = "&H0080C0FF"
couleur.jaune = "&H0000FFFF"
couleur.jaune_clair = "&H00C0FFFF"
couleur.cyan = "&H00FFFF00"
couleur.cyan_clair = "&H00FFFFC0"
couleur.bleu = "&H00FF0000"
couleur.bleu_clair = "&H00FFC0C0"
couleur.magenta = "&H00FF00FF"
couleur.magenta_clair = "&H00FFC0FF"
End Sub



