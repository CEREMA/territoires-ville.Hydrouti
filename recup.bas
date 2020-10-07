Attribute VB_Name = "recup"
Public Function recup_init(ByVal lhFicDbf As Integer) As st_hydrouti
Dim za1 As st_hydrouti
Dim num As Integer
   On Error GoTo test_Error
    Input #lhFicDbf, num, za1.type, za1.reste
recup_init = za1
Exit Function
test_Error:
   Call print_erreur("fonction recup_init ")
recup_init = za1
End Function
Public Function recup_bv(ByVal lhFicDbf As Integer) As st_save1
Dim za As st_save, za2 As st_save1
Dim num As Integer
   On Error GoTo test_Error
       Input #lhFicDbf, num, za2.type, za.nom, za.bv.nom, za.bv.type _
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
        , za.hydro1.qfuite, za.hydro1.vst, za.hydro1.vstock, za2.reste
 za2.stsave = za
recup_bv = za2
Exit Function
test_Error:
   Call print_erreur("Anomalie dans la récupération d'un bassin versant. Les derniers éléments du fichiers seront perdus. ")
'   Call print_erreur("fonction recup_bv ")
recup_bv = za2
End Function
Public Function recup_chute(ByVal lhFicDbf As Integer) As st_savch1
Dim za3 As st_savch1, za As st_savchute
Dim num As Integer
   On Error GoTo test_Error
        Input #lhFicDbf, num, za.type, za.nom, za.chute.dam, za.chute.iRadam _
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
        , za.chute.tron_ava.conduit.typ, za.reste, za3.reste
za3.stsavch = za
recup_chute = za3
Exit Function
test_Error:
   Call print_erreur("Anomalie dans la récupération d'une chute. Les derniers éléments du fichiers seront perdus. ")
'   Call print_erreur("fonction recup_chute ")
recup_chute = za3
End Function
Public Function recup_pompe(ByVal lhFicDbf As Integer) As st_savpom1
Dim za9 As st_savpom1, za As st_savpompe
Dim num As Integer
   On Error GoTo test_Error
        Input #lhFicDbf, num, za.type, za.nom, za.pompe.debits_car.qeum, za.pompe.debits_car.Fp _
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
        , za.pompe.resultat.Tsejh, za.pompe.resultat.Singul, za.pompe.resultat.Hmt, za.reste, za9.reste
za9.stsavpo = za
recup_pompe = za9
Exit Function
test_Error:
   Call print_erreur("Anomalie dans la récupération d'une pompe. Les derniers éléments du fichiers seront perdus. ")
'   Call print_erreur("fonction recup_chute ")
recup_pompe = za9
End Function
Public Function recup_decant(ByVal lhFicDbf As Integer) As st_savdec1
Dim za As st_savdecant, za4 As st_savdec1
Dim num As Integer
   On Error GoTo test_Error
za = za4.stsavdecant
        Input #lhFicDbf, num, za.type, za.nom, za.decant.Q _
        , za.decant.d, za.decant.X, za.decant.Psed, za.decant.Vhor _
        , za.decant.Long, za.decant.larg, za.decant.Hchamb _
        , za.decant.heau, za.decant.Vvert, za.decant.k, za.reste, za4.reste
za4.stsavdecant = za
recup_decant = za4
Exit Function
test_Error:
   Call print_erreur("Anomalie dans la récupération d'un bassin de décantation. Les derniers éléments du fichiers seront perdus. ")
'   Call print_erreur("fonction recup_decant ")
recup_decant = za4
End Function
Public Function recup_deversoir(ByVal lhFicDbf As Integer) As st_savdo1
Dim za5 As st_savdo1, za As st_savdo
Dim num As Integer
   On Error GoTo test_Error
    Input #lhFicDbf, num, za.type, za.nom, za.edessdo.nom, za.edessdo.nombv _
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
    , za.edo.tav, za.reste, za5.reste
 za5.stsavdo = za
recup_deversoir = za5
Exit Function
test_Error:
   Call print_erreur("Anomalie dans la récupération d'un deversoir. Les derniers éléments du fichiers seront perdus. ")
'   Call print_erreur("fonction recup_deversoir ")
recup_deversoir = za5
End Function
Public Function recup_deversoir_or(ByVal lhFicDbf As Integer) As st_savdo1
Dim za10 As st_savdo1, za As st_savdo
Dim num As Integer
   On Error GoTo test_Error
    Input #lhFicDbf, num, za.type, za.nom, za.edessdo.nom, za.edessdo.nombv _
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
    , za.edo.tav, za.reste, za10.reste
 za10.stsavdo = za
recup_deversoir_or = za10
Exit Function
test_Error:
   Call print_erreur("Anomalie dans la récupération d'un deversoir. Les derniers éléments du fichiers seront perdus. ")
'   Call print_erreur("fonction recup_deversoir ")
recup_deversoir_or = za10
End Function
Public Function recup_ret(ByVal lhFicDbf As Integer) As st_savret1
Dim za6 As st_savret1, za As st_savret
Dim num As Integer
    On Error GoTo test_Error
       Input #lhFicDbf, num, za.type, za.nom, za.retention.nom _
        , za.retention.nombv, za.retention.type_calcul, za.retention.surface _
        , za.retention.Ca, za.retention.qf, za.retention.amontana, za.retention.bmontana _
        , za.retention.deltaH, za.retention.volume, za.retention.a1montana _
        , za.retention.b1montana, za.retention.Seuil, za.retention.desssret.type _
        , za.retention.desssret.opt_long, za.retention.desssret.opt_larg _
        , za.retention.desssret.opt_prof, za.retention.desssret.opt_rap _
        , za.retention.desssret.Longueur, za.retention.desssret.Largeur _
        , za.retention.desssret.Profondeur, za.retention.desssret.Rapport _
        , za.retention.desssret.coef, za.retention.desssret.duree _
        , za.retention.desssret.Hpluie, za.retention.desssret.Hfuite, za.reste, za6.reste
 za6.stsavret = za
 recup_ret = za6
Exit Function
test_Error:
   Call print_erreur("Anomalie dans la récupération d'un bassin de retention. Les derniers éléments du fichiers seront perdus. ")
 recup_ret = za6
End Function
Public Function recup_siphon(ByVal lhFicDbf As Integer) As st_savsi1
Dim za7 As st_savsi1, za As st_savsi
Dim num As Integer
   On Error GoTo test_Error
Input #lhFicDbf, num, za.type, za.nom, za.siphon.dam, za.siphon.iRadam, za.siphon.Kam, za.siphon.dav, za.siphon.iradav _
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
, za.reste, za7.reste
za7.stsavsi = za
recup_siphon = za7
Exit Function
test_Error:
   Call print_erreur("Anomalie dans la récupération d'un siphon. Les derniers éléments du fichiers seront perdus. ")
'   Call print_erreur("fonction recup_siphon ")
recup_siphon = za7
End Function
Public Function recup_stock(ByVal lhFicDbf As Integer) As st_savsto1
Dim za8 As st_savsto1, za As st_savstock
Dim num As Integer
   On Error GoTo test_Error
        Input #lhFicDbf, num, za.type, za.nom, za.stockage.nom _
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
        , za.reste, za8.reste
za8.stsavstock = za
recup_stock = za8
Exit Function
test_Error:
   Call print_erreur("Anomalie dans la récupération d'un bassin de stockage. Les derniers éléments du fichiers seront perdus. ")
'   Call print_erreur("fonction recup_stock ")
recup_stock = za8
End Function
Public Sub recup_fich(ByVal nom As String, ByVal nom1 As String)
'Dim nom As String, nom1 As String
Dim list_enreg() As Variant
Dim ir As Integer, num
Dim za1 As st_hydrouti
Dim za2 As st_save1 'bassin versant
Dim za3 As st_savch1 'chute et conduite
Dim za4 As st_savdec1 'decantation
Dim za5 As st_savdo1 ' deversoir
Dim za6 As st_savret1 'retention
Dim za7 As st_savsi1 'siphon
Dim za8 As st_savsto1 'stockage
Dim za9 As st_savpom1 'stockage
Dim texte As String
Dim za10 As st_savdo1 ' deversoir
'nom1 = chemin_app + "etude.boh"
'nom = chemin_app + "etude1.boa"
'nom = nom_fich
'nom1 = nom_fich_edit
 On Error GoTo test_Error
ir = 0
ReDim list_enreg(ir)
    If Dir(nom) <> "" Then
        Kill nom
    End If
   lhFicDbf = FreeFile
'    Open nom1 For Random Access Read As #lhFicDbf Len = Len(za1)
    Open nom1 For Input Lock Read Write As #lhFicDbf  ' Len = Len(za1)
    Do While Not EOF(lhFicDbf)
        Line Input #lhFicDbf, texte 'num, za1.type, za1.reste
        texte1 = ""
        For i = 4 To Len(texte)
        If Mid(texte, i, 1) <> " " Then
            texte1 = texte1 + Mid(texte, i, 1)
        Else
            i = Len(texte)
        End If
        Next
'        If Not EOF(lhFicDbf) Then
            ir = ir + 1
            ReDim Preserve list_enreg(ir)
            list_enreg(ir) = texte1
'        End If
    Loop
    Close #lhFicDbf
   lhFicDbf = FreeFile
'    Open nom_fich For Random Access Read As #lhFicDbf Len = Len(za1)
    Open nom1 For Input Lock Read Write As #lhFicDbf  ' Len = Len(za1)
   lhFicDbf1 = FreeFile
'    Open nom For Output Access Write As #lhFicDbf Len = Len(za1)
    Open nom For Random Access Write Lock Read Write As #lhFicDbf1 Len = Len(za1)
   ir = 1
Do While ir <= UBound(list_enreg) 'Not EOF(lhFicDbf) And
'   Input #lhFicDbf, ev.iden, ev.surface, ev.texte
'    If ir <= UBound(list_enreg) Then
    Select Case Trim(list_enreg(ir))
        Case Is = "hydrouti"
            za1 = recup_init(lhFicDbf)
        Case Is = "versant"
            za2 = recup_bv(lhFicDbf)
        Case Is = "chute", "conduite"
            za3 = recup_chute(lhFicDbf)
        Case Is = "decantation"
            za4 = recup_decant(lhFicDbf)
        Case Is = "deversoir"
            za5 = recup_deversoir(lhFicDbf)
        Case Is = "retention"
            za6 = recup_ret(lhFicDbf)
        Case Is = "siphon"
            za7 = recup_siphon(lhFicDbf)
        Case Is = "stockage"
            za8 = recup_stock(lhFicDbf)
        Case Is = "pompe"
            za9 = recup_pompe(lhFicDbf)
        Case Is = "deversoiror"
            za10 = recup_deversoir_or(lhFicDbf)
    End Select
'    If Not EOF(lhFicDbf) Then
    Select Case Trim(list_enreg(ir))
        Case Is = "hydrouti"
            FileLength = LOF(lhFicDbf1) / Len(za1) + 1
            Put #lhFicDbf1, FileLength, za1
        Case Is = "versant"
            FileLength = LOF(lhFicDbf1) / Len(za2) + 1
            Put #lhFicDbf1, FileLength, za2
        Case Is = "chute", "conduite"
            FileLength = LOF(lhFicDbf1) / Len(za3) + 1
            Put #lhFicDbf1, FileLength, za3
        Case Is = "decantation"
            FileLength = LOF(lhFicDbf1) / Len(za4) + 1
            Put #lhFicDbf1, FileLength, za4
        Case Is = "deversoir"
            FileLength = LOF(lhFicDbf1) / Len(za5) + 1
            Put #lhFicDbf1, FileLength, za5
        Case Is = "retention"
            FileLength = LOF(lhFicDbf1) / Len(za6) + 1
            Put #lhFicDbf1, FileLength, za6
        Case Is = "siphon"
            FileLength = LOF(lhFicDbf1) / Len(za7) + 1
            Put #lhFicDbf1, FileLength, za7
        Case Is = "stockage"
            FileLength = LOF(lhFicDbf1) / Len(za8) + 1
            Put #lhFicDbf1, FileLength, za8
        Case Is = "pompe"
            FileLength = LOF(lhFicDbf1) / Len(za9) + 1
            Put #lhFicDbf1, FileLength, za9
        Case Is = "deversoiror"
            FileLength = LOF(lhFicDbf1) / Len(za10) + 1
            Put #lhFicDbf1, FileLength, za10
        End Select
'    End If
'    End If
    ir = ir + 1
Loop
    Close #lhFicDbf
    Close #lhFicDbf1
Exit Sub
test_Error:
   Call print_erreur("fonction recup_fich  ")

End Sub
Public Sub flockb(nom)
Dim za1 As st_hydrouti
 On Error GoTo test_Error
lhFicDbb = FreeFile
Open nom For Random Access Read Write Lock Read Write As #lhFicDbb Len = Len(za1)
Exit Sub
test_Error:
   Call print_erreur("fonction flockb  fichier " + nom)
End Sub
Public Sub flocka(nom)
 On Error GoTo test_Error
    lhFicDba = FreeFile
    DoEvents
   Open nom For Input Lock Read Write As #lhFicDba
Exit Sub
test_Error:
   Call print_erreur("fonction flocka  fichier " + nom)
End Sub
Public Sub funlockb()
  On Error GoTo test_Error
   Close #lhFicDbb
Exit Sub
test_Error:
   Call print_erreur("fonction funlockb  fichier " + lhFicDbb)
End Sub
Public Sub funlocka()
 On Error GoTo test_Error
    Close #lhFicDba
Exit Sub
test_Error:
   Call print_erreur("fonction funlocka  fichier " + lhFicDba)
End Sub

