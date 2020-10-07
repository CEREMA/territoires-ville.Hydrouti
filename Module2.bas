Attribute VB_Name = "Module2"
Type file_spec
    lecteur As String
    Chemin As String
    nom As String
    extension As String
    nomcomplet As String
    dr_type As Integer
    f_attr As Integer
    f_size As Double
End Type
Public Function calcul_vitesse(ByRef tr As troncon, ByRef res_conduit As debit_conduit, ByVal zhe As Double) As Double
    Dim s As Double, v As Double, beta As Double
'    zhe = hautdech
    If zhe / tr.conduit.Diametre < 1 Then
        beta = 2 * arccosinus((1 - 2 * zhe / tr.conduit.Diametre))
    Else
        beta = 2 * pi
    End If
        s = (1 / 8) * (tr.conduit.Diametre ^ 2) * (beta - Sin(beta))
        v = res_conduit.debit / s
    calcul_vitesse = v

End Function

Public Function verif_regime(ByVal Q As Double, ByRef canal As conduite, ByVal tre As Double) As String
' q en m3/s
Dim regim As String
Dim betam As Double
Dim qv As deb_vit
'qv = debvit_ps(canal)
'        betam = angle(Q / qv.debit)
'        betam = beta
If tre / canal.Diametre < 1 Then
        betam = 2 * arccosinus(1 - 2 * tre / canal.Diametre)
Else
        betam = 2 * pi
End If
regim = calcul_ecoul(Q, canal.Diametre, betam)
verif_regime = regim



End Function
Public Function calcul_Froude1(vit As Double, hr As Double) As Double
calcul_Froude1 = 0
If hr > 0 Then
calcul_Froude1 = vit / (9.81 * hr) ^ 0.5
End If
End Function
Public Function calcul_Froude(Q As Double, hr As Double, d As Double) As Double
calcul_Froude = 0
If hr > 0 Then
calcul_Froude = Q / (hr ^ 2 * (9.81 * d) ^ 0.5)
End If
End Function
Public Function verif_regime0(ByVal Q As Double, ByRef canal As conduite) As String
' q en m3/s
Dim regim As String
Dim betam As Double
Dim qv As deb_vit
qv = debvit_ps(canal)
        betam = angle(Q / qv.debit)
        betam = beta
        regim = calcul_ecoul(Q, canal.Diametre, betam)
verif_regime0 = regim



End Function

Function verif_do_charge_0(ByRef edev As deversoir, ByRef edo_r As deversoir_resultat, ByRef tr As troncon, ByVal Qpl As Double, ByVal Qrin As Double, ByRef uc_g As UC_graphique) As String
Dim ok As Boolean, nok0 As Boolean, nok1 As Boolean
'Dim edo_res As deversoir_resultat
ok = True
Dim HM As Double, Ham As Double, Hav As Double, Haam As Double, Haav As Double
Dim dHa As Double
Dim Tram As Double, c As Double

Dim ed As deversoir
Dim beta As Double, s As Double, v As Double
Dim Qav As Double, Qdev As Double
Dim i0 As Integer
Dim sres As String
i = 0
sres = "  Fonctionnement à débit de pointe "
'Tram = edo_r.Tram
Tram = edessdo.Tram
'a revoir verification Tram/ haut d'eau< Tram<diametre ?
If Tram = 0 Then
    Tram = tr.conduit.Diametre * 0.9
End If
'fin a revoir
c = edo_r.c
nok0 = True
' debit en m3/s
Qav = Qrin * 1.3

While nok0 And i0 < 20
Qdev = Qpl - Qav
    'Qav = 0.172
    ed = edev
    
    'a revoir ? 0.01 pente amont-aval de la ligne d'energie
    dHa = 0.01 * ed.Longueur
    
    HM = (0.85 * Qdev / (c * ed.Longueur)) ^ (2# / 3#)
    Ham = Tram - ed.hauteur
    Hav = (4 * HM - Ham) / 3#
    Haav = Hav + ed.hauteur
    'HAam = rech_haut_do_vam(dam, qcri / 0.3)
    ' calcul de vitesse amont debit pluie
    
    ' verification vitesse d'écoulement amont pour qcri
    If Tram / tr.conduit.Diametre < 1 Then
        beta = 2 * arccosinus(1 - 2 * Tram / tr.conduit.Diametre)
    Else
        beta = 2 * pi
    End If
        s = (1 / 8) * (tr.conduit.Diametre ^ 2) * (beta - Sin(beta))
        v = Qpl / s
    Haam = Tram + (v ^ 2) / (2 * 9.81)
    Haavd = Haam - dHa
    If (Haam - Haav) < dHa Then
    sres = sres + Chr$(13) + Chr$(10) + " Perte de charge Amont Aval : augmenter la longueur du do "
    End If
'    MsgBox "Haam = " + Str(Haam) + Chr$(13) + " Haav disponible  = " + Str(Haavd) + Chr$(13) + " Haav   = " + Str(Haav), vbOKOnly, "verification charge"
    
    'calcul Qav" en fonctionnement
    Dim tav As Double, a As Double, Vavp As Double
    Dim Tavp As Double, Sav As Double, Qavp As Double
    Dim imot As Double, irad As Double
    Dim i1 As Integer, dQavp As Double, signe As Integer
    
    
    tav = Hav + ed.hauteur + ed.Longueur * ed.pente
    'calcul de a * v2 /2g
     '   tavdav = tav / edo.tron_ava.conduit.Diametre
        a = rech_do_A(tav, ed.tron_ava.conduit.Diametre)
    
    Sav = debvit_ps(ed.tron_ava.conduit).debit / debvit_ps(ed.tron_ava.conduit).vitesse
    
    Qavp = Qav
    dQavp = Qavp / 10
    nok1 = True
    i1 = 0
    signe = 0
    While nok1 And ii < 50
    Vavp = Qavp / Sav
    imot = pent_mot0(ed.tron_ava.conduit, Qavp)
    irad = ed.tron_ava.conduit.pente
    Tavp = a * Vavp ^ 2 / (2 * 9.81) + ed.tron_ava.conduit.Longueur * (imot - irad) + ed.tron_ava.conduit.Diametre
    If Abs(Tavp - tav) < 0.0001 Then
    nok1 = False
    Else
    If Tavp < tav Then
    'MsgBox "augmenter Qavp", vbOKOnly
    If (tav - Tavp) * signe < 0 Then
        dQavp = dQavp * 0.5
    End If
    signe = 1
    ElseIf Tavp > tav Then
    If (tav - Tavp) * signe < 0 Then
        dQavp = dQavp * 0.5
    End If
    signe = -1
    'MsgBox "diminuereQavp", vbOKOnly
    End If
    End If
    'MsgBox "Tav = " + Str(tav) + " Tav'= " + Str(Tavp), vbOKOnly, "verification charge"
    Qavp = Qavp + signe * dQavp
    i1 = ii + 1
    Wend
    If Abs(Qav - Qavp) < 0.0001 Then
        nok0 = False
    End If
    Qav = Qavp + (Qav - Qavp) / 4#
    i0 = i0 + 1
Wend

With edo_r
    .Tram = Tram
    .a = a
    .c = c
    .HM = HM
    .Hav = Hav
    .Ham = Ham
    .Haav = Haav
    .Haavd = Haavd
    .Haam = Haam
    .Qav = Qav
    .Qdev = Qdev
End With
Call init_graphdo(uc_g)
Call dess_troncon(uc_g, tr, couleur.gris) ' vbBlack)
Call dess_predo(uc_g, ed, couleur.noir)
Call dess_cot(uc_g, couleur.noir) ' vbBlack)

Call dessin_do_debpointe(uc_g, True, True, True)
''Dim zplam_am As Double, zplam_av As Double, zplav_am As Double, zplav_av As Double
''Dim ct() As Variant, qvm(5) As Variant, haut As Double, pentmot As Double
''zplam_av = ed.radamo + Tram
''
''Dim qcal
'''Dessin du fonctionnement dans l'onglet Deversoir
'''dessin de la ligne d'eau
''' conduite amont
''qcal = Qpl
''If qcal < debvit_ps(tr.conduit).debit Then
''    Call cana(tr.conduit, ct)
''    ltc = calc_par(tr.conduit)
''    qvi = caltran1(qcal * 1000, ct, ltc)
''    Debug.Print "qvi : 1= débit"; qvi(1); " 2=vitesse"; qvi(2); "qvi"; " 3=acceleration"; qvi(3)
''    Debug.Print " 4=Largeur libre"; qvi(4); " 5 = hauteur; "; qvi(5); ""
'''                vitmax = qvm(2)
'''                qvm = caltran1(qps / 10#, ct, ltc)
'''                vit10 = qvm(2)
'''                qvm = caltran1(qps / 100#, ct, ltc)
'''                vit100 = qvm(2)
''   haut = qvi(5)
''
''    zplam_am = tr.radamo + haut
'''    piezoava = tr.radava + haut
''Else
''    zplam_av = tr.radava + tr.conduit.Diametre
'''  piezoamo = tr.radamo + tr.conduit.Diametre
''   pentmot = pent_mot0(tr.conduit, qcal)
''    zplam_am = zplam_am + pentmot * tr.conduit.Longueur
''End If
'''dessin des lignes d'eau
'''Call init_graphdo(uc_g)
'''dessin ligne d'eau conduite amont
''uc_g.dess_lign tr.Absamo, zplam_am, tr.Absava, zplam_av, couleur.vert, 2
''
'''Call dess_piezo(uc_g, ed.tron_ava, Qavp * 1000, vbRed)
'''dessin ligne d'eau sur la lame
'' uc_g.dess_lign ed.Absamo, ed.radamo + Tram, ed.Absava, ed.radava + tav, couleur.vert, 2
''

sres = sres + Chr(13) + Chr$(10) + " Débit dans la conduite étranglée " + ajout_zero(Trim(str(Round(Qav, 3)))) + "m3/s"
sres = sres + Chr(13) + Chr$(10) + " Débit déversé : " + ajout_zero(Trim(str(Round(Qdev, 3)))) + " m3/s"
resudev.debetranglee = ajout_zero(Trim(str(Round(Qav, 3))))
resudev.debdeverse = ajout_zero(Trim(str(Round(Qdev, 3))))
verif_do_charge_0 = sres


End Function

Public Sub dessin_do_debpointe_0(ByRef uc_g As UC_graphique, ByVal okcharge As Boolean, ByVal okpiezo As Boolean, ByVal okeau As Boolean)
Dim zplam_am As Double, zplam_av As Double, zplav_am As Double, zplav_av As Double
'Dim qvm(5) As Variant, haut As Double, pentmot As Double
Dim tr As troncon
Dim res_conduit As debit_conduit
Dim qcal As Double
'Call init_graphdo(uc_g)

'dessin de la ligne dans le deversoir
'dans la frmdessin
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
        If okpiezo Then
'            uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.piezoamo, res_conduit.p_Eau_inter.X, res_conduit.piezointer.Y, couleur.orange, 2
'            uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, res_conduit.zeau_ava.X, res_conduit.piezoava, couleur.orange, 2
            uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.piezoamo, res_conduit.piezointer.X, res_conduit.piezointer.Y, couleur.orange, 2
            uc_g.dess_lign res_conduit.piezointer.X, res_conduit.piezointer.Y, res_conduit.piezointer0.X, res_conduit.piezointer0.Y, couleur.orange, 2
            uc_g.dess_lign res_conduit.piezointer0.X, res_conduit.piezointer0.Y, res_conduit.zeau_ava.X, res_conduit.piezoava, couleur.orange, 2
        End If
        If okeau Then
'            ' uc_g.dess_lign tr.Absamo, zplam_am, res_conduit.piezointer.X, res_conduit.piezointer.Y, couleur.bleu, 2
'            ' uc_g.dess_lign res_conduit.piezointer.X, res_conduit.piezointer.Y, tr.Absava, zplam_av, couleur.bleu, 2
'            uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.zeau_amo.Y, res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, couleur.bleu, 2
'            uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, res_conduit.zeau_ava.X, res_conduit.zeau_ava.Y, couleur.bleu, 2
            uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.zeau_amo.Y, res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, couleur.bleu, 2
            uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, res_conduit.p_Eau_inter0.X, res_conduit.p_Eau_inter0.Y, couleur.bleu, 2
            uc_g.dess_lign res_conduit.p_Eau_inter0.X, res_conduit.p_Eau_inter0.Y, res_conduit.zeau_ava.X, res_conduit.zeau_ava.Y, couleur.bleu, 2
        End If
    'dessin charge
        'recalcul de charge amont en fonction de hauteur d'eau
        ' verification vitesse d'écoulement amont pour qcri
        If okcharge Then
            Call inter_charge_tr(tr, res_conduit)
            uc_g.dess_lign tr.Absamo, res_conduit.chargeamo, res_conduit.piezointer.X, res_conduit.chargeinter, couleur.rouge, 2
            uc_g.dess_lign res_conduit.piezointer.X, res_conduit.chargeinter, res_conduit.piezointer0.X, res_conduit.chargeinter0, couleur.rouge, 2
            uc_g.dess_lign res_conduit.piezointer0.X, res_conduit.chargeinter0, tr.Absava, res_conduit.chargeava, couleur.rouge, 2
        End If
'dessin ligne d'eau sur la lame
    'dessin des lignes d'eau
        'uc_g.dess_lign edo.Absamo, edo.radamo + edo_res.Tram, edo.Absava, edo.radamo + edo.hauteur + edo_res.Hav, couleur.bleu, 2
        If okpiezo Then
            uc_g.dess_lign edo.Absamo, edo.radamo + edessdo.Tram, edo.Absava, edo.radamo + edo.hauteur + edo_res.Hav, couleur.orange, 2
        End If
        If okeau Then
            uc_g.dess_lign edo.Absamo, edo.radamo + edessdo.Tram, edo.Absava, edo.radamo + edo.hauteur + edo_res.Hav, couleur.bleu, 2
        End If
    'dessin charge
        If okcharge Then

            uc_g.dess_lign edo.Absamo, edo.radamo + edo_res.Haam, edo.Absava, edo.radava + edo_res.Haav, couleur.rouge, 2
        End If
''''dessin troncon aval
    tr = edo.tron_ava
res_conduit = calc_debit_tr(tr, edo_res.Qav)
'''
 'dessin des lignes d'eau
        res_conduit.zphe_ava = res_conduit.hauteur + tr.radava
        Call inter_piezo_eau(tr, res_conduit)
        If okpiezo Then
            uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.piezoamo, res_conduit.p_Eau_inter.X, res_conduit.piezointer.Y, couleur.orange, 2
            uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, res_conduit.zeau_ava.X, res_conduit.piezoava, couleur.orange, 2
        End If
        If okeau Then
'''        ' uc_g.dess_lign tr.Absamo, zplam_am, res_conduit.piezointer.X, res_conduit.piezointer.Y, couleur.bleu, 2
'''        ' uc_g.dess_lign res_conduit.piezointer.X, res_conduit.piezointer.Y, tr.Absava, zplam_av, couleur.bleu, 2
            uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.zeau_amo.Y, res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, couleur.bleu, 2
            uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, res_conduit.zeau_ava.X, res_conduit.zeau_ava.Y, couleur.bleu, 2
        End If
        If okcharge Then
'''
'''    'dessin charge
'''        ' dessin de la charge repris par inter_charge_pr
        Call inter_charge_tr(tr, res_conduit)
        uc_g.dess_lign tr.Absamo, res_conduit.chargeamo, res_conduit.piezointer.X, res_conduit.chargeinter, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer.X, res_conduit.chargeinter, res_conduit.piezointer0.X, res_conduit.chargeinter0, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer0.X, res_conduit.chargeinter0, tr.Absava, res_conduit.chargeava, couleur.rouge, 2
        End If
End Sub

Sub ini_bv()

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
    ebv.Qchoisi = ""
    eph.amontana = 0#
    eph.bmontana = 0#
    eph.lcrin = 0
    eph.ceau = 0
    eph.aeu = 0
    eph.beu = 0
    eph.a1montana = 0#
    eph.b1montana = 0#
    eph.Seuil = 0
    ehyd.DM = 0
    ehyd.dt = 0
    ehyd.HM = 0#
    ehyd.HT = 0#
    ehyd.pas = 1
    ehyd.Teta = 0.5
    ehyd.kdesbor = 0#
    ehyd.qfuite = 0
    ehyd.vst = 0#
    ehyd.vstock = 0#
End Sub
Sub Centre(frm As Form, Optional ref)
'Cette procédure centre la fenêtre frm par rapport
'à la position de la fenêtre ref
'si ref n'est pas mentionné, la fenêtre est centrée
'par rapport à l'écran.
Dim milX As Long, milY As Long
If IsMissing(ref) Then
    milX = Screen.Width / 2
    milY = Screen.Height / 2
Else
    milX = ref.Left + ref.Width / 2
    milY = ref.Top + ref.Height / 2
End If
frm.Left = milX - frm.Width / 2
frm.Top = milY - frm.Height / 2

End Sub
Function calcul_debit_ep(ByRef ebv As st_Bv, ByRef eph As st_ParHydro)
Dim a As Double, l As Double, p As Double, c As Double
Dim amontana As Double, bmontana As Double
Dim tc As Double, c1 As Double, k As Double, c_i As Double, c_c As Double
Dim c_a As Double, grdm As Double, malon As Double
Dim Qbrut As Double, Qcor As Double

a = ebv.surface
l = ebv.lghydr
p = ebv.phydr / 10000#
c = ebv.imper / 100
If a <> 0 And l <> 0 And p <> 0 And eph.amontana <> 0 And eph.bmontana <> 0 _
    And eph.a1montana <> 0 And eph.b1montana <> 0 And eph.Seuil <> 0 Then
' temps de concentration Formule etat de californie
    tc = 0.0195 * (l ^ 0.77) * p ^ -0.385
    'eaux usées
' houpie 2004/04/07
    If tc < eph.Seuil Then
        amontana = eph.amontana
        bmontana = -eph.bmontana
    Else
        amontana = eph.a1montana
        bmontana = -eph.b1montana
    End If
''''''''
' eaux pluviales
    ' caquot
    c1 = (1 / (1 + 0.29 * bmontana))
    k = ((amontana * 0.5 ^ bmontana) / (6.6)) ^ c1
    c_i = (bmontana * -0.41) * c1
    c_c = c1
    c_a = (bmontana * 0.51 + 0.95) * c1
    grdm = l / (100 * (a ^ 0.5))
    If grdm < 0.8 Then grdm = 0.8
    malon = (grdm / 2) ^ ((0.84 * bmontana) / (1 + 0.287 * bmontana))
    Qbrut = k * (p ^ c_i) * (c ^ c_c) * (a ^ c_a)
    Qcor = Qbrut * malon
   ' fin caquot
   'début hydrogramme
          'fin hydrogramme
          ebv.tc = tc
          ebv.Qbrut = Qbrut
          ebv.Qcor = Qcor
          ebv.Qeu = 0#
          ebv.Qecp = 0#
          ebv.Qts = 0#
          ebv.Qprin = 0#
          ebv.Qrin = 0#
End If
End Function
Function calcul_debit_epmr(ByRef ebv As st_Bv, ByRef eph As st_ParHydro)
Dim a As Double, l As Double, p As Double, c As Double
Dim amontana As Double, bmontana As Double
Dim tc As Double, c1 As Double, k As Double, c_i As Double, c_c As Double
Dim c_a As Double, grdm As Double, malon As Double
Dim Qbrut As Double, Qcor As Double, Qmr As Double
Dim ipluie As Double

a = ebv.surface
l = ebv.lghydr
p = ebv.phydr / 10000#
c = ebv.imper / 100
amontana = eph.amontana
bmontana = -eph.bmontana
If a <> 0 And l <> 0 And p <> 0 And eph.amontana <> 0 And eph.bmontana <> 0 _
    And eph.a1montana <> 0 And eph.b1montana <> 0 And eph.Seuil <> 0 Then
' temps de concentration Formule etat de californie
    tc = 0.0195 * (l ^ 0.77) * p ^ -0.385
    'eaux usées
' houpie 2004/04/07
    If tc < eph.Seuil Then
        amontana = eph.amontana
        bmontana = -eph.bmontana
    Else
        amontana = eph.a1montana
        bmontana = -eph.b1montana
    End If
'''''''''''''
  ' eaux pluviales Methode rationnelle californie
    ' intensité de la pluie
    ipluie = amontana * tc ^ bmontana
    
    Qmr = 0.167 * c * ipluie * a
   ' fin mr californie
   'début hydrogramme
          'fin hydrogramme
          ebv.tc = tc
          ebv.Qmr = Qmr
         ' ebv.Qcor = Qcor
          ebv.Qeu = 0#
          ebv.Qecp = 0#
          ebv.Qts = 0#
          ebv.Qprin = 0#
          ebv.Qrin = 0#
End If
End Function

Function calcul_debit_eu(ByRef ebv As st_Bv, ByRef eph As st_ParHydro)
Dim a As Double, l As Double, p As Double, c As Double, h As Double
Dim t_dilu As Double, icrin As Double
Dim tc As Double, c1 As Double, k As Double, c_i As Double, c_c As Double
Dim c_a As Double, grdm As Double, malon As Double, ceau As Double
Dim aeu As Double, beu As Double
Dim Qbrut As Double, Qcor As Double, qeum As Double, Qeu As Double
Dim Qecp As Double, Qts As Double, Qprin As Double, Qrin As Double
a = ebv.surface
l = ebv.lghydr
p = ebv.phydr / 10000#
c = ebv.imper / 100
h = ebv.nhab
t_dilu = ebv.tdilu
ceau = ebv.ceau
aeu = eph.aeu
beu = eph.beu
icrin = eph.lcrin
If a <> 0 And l <> 0 And p <> 0 Then
    tc = 0.0195 * (l ^ 0.77) * p ^ -0.385
    'eaux usées
    qeum = h * ceau / 86400
    If qeum <> 0 Then Qeu = (aeu + beu / (qeum ^ 0.5)) * qeum
    Qecp = qeum * t_dilu / 100
    Qts = Qeu + Qecp
    Qprin = icrin * a * c * (1 - tc / 200)
    Qrin = Qprin + Qts
    ebv.tc = tc
    ebv.Qeu = Qeu
    ebv.Qecp = Qecp
    ebv.Qts = Qts
    ebv.Qprin = Qprin
    ebv.Qrin = Qrin
    
End If
End Function

Function calcul_hydro1(ByRef ebv As st_Bv, ByRef eph As st_ParHydro, ByRef ehyd As st_hydr)
Dim a As Double, l As Double, p As Double, c As Double, h As Double
Dim amontana As Double, bmontana As Double, DM As Double, dt As Double
Dim HT As Double, HM As Double, k As Double, kdesbor As Double, qfuite As Double
Dim qf As Double, vst As Double
a = ebv.surface
l = ebv.lghydr
p = ebv.phydr / 10000#
c = ebv.imper / 100
h = ebv.nhab
DM = ehyd.DM
dt = ehyd.dt
qfuite = ehyd.qfuite
' 20040907
'If ebv.tc < eph.seuil Then
If DM < eph.Seuil Then
    amontana = eph.amontana
    bmontana = -eph.bmontana
Else
    amontana = eph.a1montana
    bmontana = -eph.b1montana
End If
HT = 0#
HM = 0#
' 20040323
HM = ehyd.HM
HT = ehyd.HT

kdesbor = 0#
vst = 0#
   'début hydrogramme
    If DM <> 0 And dt <> 0 And HM <> 0 And HT <> 0 Then
          kdesbor = 5.07 * (a ^ 0.18) * ((p * 100) ^ -0.36) * ((1 + c) ^ -1.9) * (DM ^ 0.21) * (l ^ 0.15) * (HM ^ -0.07)
        k = kdesbor
        If a < 250 Then kdesbor = 0.7 * (a ^ 0.09) * k
        If a < 6 Then kdesbor = 0.8 * k
 
    ElseIf DM <> 0 And dt <> 0 Then
'        HT = Int(dt * amontana * (dt) ^ bmontana)
        HT = Round((dt * amontana * (dt) ^ bmontana), 0)
'        HM = Int(DM * amontana * (DM) ^ bmontana)
        HM = Round((DM * amontana * (DM) ^ bmontana), 0)
        kdesbor = 5.07 * (a ^ 0.18) * ((p * 100) ^ -0.36) * ((1 + c) ^ -1.9) * (DM ^ 0.21) * (l ^ 0.15) * (HM ^ -0.07)
        k = kdesbor
        If a < 250 Then kdesbor = 0.7 * (a ^ 0.09) * k
        If a < 6 Then kdesbor = 0.8 * k
'        FrmBV.kbestxt.Text = kdesbor
    End If
    If qfuite = 0 Then
        qfuite = 10
    End If
    If (c * a) = 0 Then
        qf = 0
        vst = 0
    Else
        qf = 60 * qfuite / (c * a)
        vst = 10 * a * c * (-bmontana * qf / (bmontana + 1) * ((qf / (amontana * (bmontana + 1))) ^ (1 / bmontana)))
    End If
'        FrmBV.txtvstock(1).Text = Int(vst)
    ehyd.HM = HM
    ehyd.HT = HT
    ehyd.kdesbor = kdesbor
    ehyd.vst = vst
         'fin hydrogramme
End Function
Function calcul_hyeto(ByRef ehyd As st_hydr, ByVal pas As Integer)
'    Dim ehyd As st_hydr
    Dim imax As Integer, nb As Integer, rest As Integer, i As Integer
    Dim j As Integer
    Dim dy As Double, Y As Double
    
    Dim i1 As Double, i2 As Double
    Dim Teta As Double, DM As Double, dt As Double, HT As Double, HM As Double
    Dim nbpoint As Integer
Dim a(5)
Dim o(5)
Dim l_int() As Variant
ReDim l_int(0)
ReDim Hpluie(500)
On Error GoTo test_Error

    DM = ehyd.DM
    dt = ehyd.dt
    HT = ehyd.HT
    HM = ehyd.HM
  '  Teta = Int(ehyd.Teta * (dt - DM))
    Teta = ehyd.Teta * (dt)
'    i1 = HM / DM
'    i2 = (HT - HM) / (dt - DM)

Call Discret_hyeto(2, dt, HT, DM, HM, Teta, pas, l_int, Hpluie)
'''
'''imax = 0
'''i = 0
'''j = 2
'''a(0) = 0:               o(0) = 0
'''a(1) = 0:               o(1) = 0
'''a(2) = Teta:            o(2) = 2 * i2
'''a(3) = a(2) + DM / 2:   o(3) = 2 * (i1 - i2)
'''a(4) = a(2) + DM:       o(4) = 2 * i2
'''a(5) = dt:              o(5) = 0
'''nbpoint = dt / pas
''' Do While j <= 5
'''        Do While (i * pas) < a(j)
'''            i = i + 1
'''            dy = Y_de_X(a(j), a(j - 1), o(j), o(j - 1), i * pas)
'''            If i <= nbpoint Then
'''                hpluie(i) = 60# * (Y + dy) / 2#
'''                If hpluie(i) > imax Then imax = hpluie(i)
'''                Y = dy
'''            End If
'''       Loop
'''    j = j + 1
'''    Loop
'''    ReDim Preserve hpluie(nbpoint)
 '   Call dessin_hyeto1
    Exit Function
test_Error:
        Call print_erreur("anomalie dans calcul_hyeto")

End Function

Function dessin_hyeto1()
Dim i As Integer
Dim nbpoint As Integer
nbpoint = UBound(Hpluie)
'Frm_bv2.hyeto.RowCount = nbpoint
'i = 1
'    Do While i <= nbpoint
'        With Frm_bv2.hyeto
'            .Column = 1
'            .Row = i
'            .Data = -1 * hpluie(i)
'        End With
'        i = i + 1
'    Loop
End Function
Function qpas(a, c, hp, q0, pas)
Dim inpluie As Double, c1 As Double, kdesbor As Double, vstock As Double, qfuite As Double
kdesbor = ehyd.kdesbor
vstock = ehyd.vstock
qfuite = ehyd.qfuite
    inpluie = a * c * hp / 360#
    c1 = (Exp(-pas / kdesbor))
    qpas = c1 * q0 + (1 - c1) * inpluie
    If qpas > qfuite Then
        vstock = vstock + (qpas - qfuite / 1000) * pas * 60
'        vstock = vstock + (qpas - qfuite) * pas * 60
        ehyd.vstock = vstock
    End If
End Function
Function Y_de_X(ByVal x1 As Integer, ByVal x2 As Integer, ByVal y1 As Double, ByVal y2 As Double, ByVal X As Integer)
    Dim a As Double, b As Double
    a = (y2 - y1) / (x2 - x1)
    b = y1 - x1 * (y2 - y1) / (x2 - x1)
    Y_de_X = a * X + b
    'frmbv.Hyeto.Data = 60 * Y
End Function

Sub calcul_hydro(ByVal pas As Integer)
Dim qfuite As Double, a As Double, crui As Double, dt As Double
Dim vrui As Double, Qmax As Double, c As Variant, qq As Variant
Dim i As Integer, nbval As Integer
Dim j As Integer
Dim l_h() As Variant
ReDim Q(600)
On Error GoTo test_Error

i = 1

Q(0) = 0
qfuite = ehyd.qfuite
a = ebv.surface
'Debug.Print ebv.perti
dt = ehyd.dt
crui = ebv.imper
ehyd.vstock = 0

If ebv.type = "U" Then
Call Hye_Hyd_1r(Hpluie, pas, ehyd.kdesbor, a, crui / 100#, 1, Q)
Else
ReDim l_h(UBound(Hpluie))
For i = 1 To UBound(l_h)
    l_h(i) = Hpluie(i)
Next

       Call Perte_init(ebv.perti, l_h, pas)
       Call Calcul_Infiltr(l_h, pas, ebv.ahorton, ebv.bhorton, ebv.vinf)
ReDim Q(UBound(l_h))
For i = 1 To UBound(Q)
    Q(i) = l_h(i)
Next
        For i = 1 To UBound(Q)
            If Q(i) > Qmax Then
                Qmax = Q(i)
            End If
        Next
'' certu 20080903
'If ebv.trep = 0 Then ebv.trep = Int(ehyd.kdesbor)
      Call Hye_Hyd_2r(Q, pas, ebv.trep / 2#, a, 1)

End If
'dessin
    Exit Sub
test_Error:
        Call print_erreur("anomalie dans calcul_hyeto")

End Sub
Sub dessin_hydro1(ByVal qfuite As Double, ByVal pas As Double)
nbval = UBound(Q)
i = 0
'Close #lhFicDbf
'lhFicDbf = FreeFile
'Open fichtmp For Input As #lhFicDbf
'    With Frm_bv2.hydro
''        .ColumnCount = 2
'        .RowCount = nbval '+ 1
'        Do While i < nbval
'            i = i + 1
''            Line Input #lhFicDbf, c
'            .Column = 2
'            .Row = i
'            .Data = q(i) * 1000
'            .Column = 1
'            .Row = i
'            .Data = qfuite
'        Loop
'    End With
'  j = Frm_bv2.hyeto.RowCount
'  Frm_bv2.hyeto.RowCount = nbval
'    For i = j + 1 To nbval
'
'    Frm_bv2.hyeto.Row = i
'    Frm_bv2.hyeto.Data = 0
'    Next
   
End Sub

Public Sub calcul_stock(ByRef ehyd As st_hydr, ByVal pas As Double)
Dim i As Integer
Dim qpas As Double
Dim vstock As Double, qfuite As Double
vstock = 0
qfuite = ehyd.qfuite
For i = 1 To UBound(Q)
qpas = Q(i)
    If qpas > qfuite / 1000 Then
        vstock = vstock + (qpas - qfuite / 1000) * pas * 60
        ehyd.vstock = vstock
    End If
Next
ehyd.vstock = vstock

End Sub
 Function rech_do_A(ByVal tav As Double, ByVal diam As Double) As Double
 Dim tavdav As Double, a As Double
 
   tavdav = tav / diam
    If tavdav = 1 Then a = 0.4
    If tavdav > 1 And tavdav <= 1.2 Then
        a = 0.4 + ((0.75 - 0.4) / (1.2 - 1)) * (tavdav - 1)
    End If
    If tavdav > 1.2 And tavdav <= 1.4 Then
        a = 0.75 + ((0.95 - 0.75) / (1.4 - 1.2)) * (tavdav - 1.2)
    End If
    If tavdav > 1.4 And tavdav <= 1.6 Then
        a = 0.95 + ((1.1 - 0.95) / (1.6 - 1.4)) * (tavdav - 1.4)
    End If
    If tavdav > 1.6 And tavdav <= 1.8 Then
        a = 1.1 + ((1.15 - 1.1) / (1.8 - 1.6)) * (tavdav - 1.6)
    End If
    If tavdav > 1.8 And tavdav <= 2# Then
        a = 1.15 + ((1.2 - 1.15) / (2# - 1.8)) * (tavdav - 1.8)
    End If
    If tavdav > 2# And tavdav <= 3# Then
        a = 1.2 + ((1.35 - 1.2) / (3 - 2)) * (tavdav - 2)
    End If
    If tavdav > 3 And tavdav <= 5 Then
        a = 1.35 + (((1.45 - 1.35) / (5 - 3)) * (tavdav - 3))
    End If
rech_do_A = a
End Function

Function rech_dav_do(ByRef edo As deversoir, ByVal ld As Double) As Double
Dim v As Double
Dim s As Double
Dim vcri As Double
Dim tav As Double
Dim tcr As Double
Dim a As Double
Dim ddav As Double
Dim icri As Double
Dim dam As Double, kav As Double, dav As Double, iradav As Double
Dim ldav As Double, pdav As Double
Dim q10 As Double, qcri As Double
Dim laval As Double

Dim istop As Integer
istop = 0
' initialisation des donnees
'dam = edessdo.dam
'dav = edessdo.dav
'kav = edessdo.kav
'iradav = edessdo.iradav

dam = edessdo.tron_amo.conduit.Diametre
dav = edessdo.tron_ava.conduit.Diametre
kav = edessdo.tron_ava.conduit.rugosite
iradav = edessdo.tron_ava.conduit.pente


laval = edessdo.tron_ava.conduit.Longueur
q10 = edessdo.Qpluie
qcri = edessdo.Qrin ' julienne 20030725 + edessdo.Qts
longdo = edo.Longueur
hautdo = edo.hauteur
pentedo = edo.pente
ddav = 0.01
sign = 1
ldav = ld
If ldav - laval <> 0 Then
    sign = Abs((ldav - laval)) / ((ldav - laval))

'While Abs(ldav - laval) > 0.001 And ldav > 0 And istop < 2000 '0.001
While Abs(ldav - laval) > 0.001 And istop < 2000  '0.001
    istop = istop + 1
    If ldav <= 0 Then
    dav = dav + sign * ddav
    ddav = ddav / 2
    Else
    If sign * (ldav - laval) < 0 Then
        ddav = ddav / 2#
        sign = sign * -1
    End If
    End If
         dav = dav - sign * ddav
     vcri = (qcri / 1000) / ((3.14159 * (dav ^ 2)) / 4)
    icri = (vcri / (kav * (dav / 4) ^ (2 / 3))) ^ 2
    If tav = 0 Then
        tav = hautdo + longdo * pentedo
    End If
    a = rech_do_A(tav, dav)
    
'    Debug.Print (tav - (a * (vcri ^ 2) / 19.62) - dav - tcr), (icri - iradav)
    ldav = (tav - (a * (vcri ^ 2) / 19.62) - dav - tcr) / (icri - iradav)
  

Wend
End If

If istop >= 2000 Then
    MsgBox "non convergence", vbOKOnly, " rech hauteur"
End If

If ldav > 0 And istop < 2000 Then
rech_dav_do = dav

Else
rech_dav_do = 0
End If
End Function
Function rech_ham_do(ByRef edo As deversoir, ByVal ld As Double) As Double

Dim v As Double
Dim s As Double
Dim vcri As Double
Dim tav As Double
Dim tcr As Double
Dim a As Double
Dim ddav As Double
Dim icri As Double
Dim dam As Double, kav As Double, dav As Double, iradav As Double
Dim ldav As Double, pdav As Double
Dim q10 As Double, qcri As Double
Dim laval As Double
Dim istop As Integer
istop = 0
' initialisation des donnees
If ld > 0 Then
'dam = edessdo.dam
'dav = edessdo.dav
'kav = edessdo.kav
'iradav = edessdo.iradav

dam = edessdo.tron_amo.conduit.Diametre
dav = edessdo.tron_ava.conduit.Diametre
kav = edessdo.tron_ava.conduit.rugosite
iradav = edessdo.tron_ava.conduit.pente

laval = edessdo.tron_ava.conduit.Longueur
q10 = edessdo.Qpluie
qcri = edessdo.Qrin ' julienne 20030725 + edessdo.Qts

longdo = edo.Longueur
hautdo = edo.hauteur
pentedo = edo.pente
ddav = 0.01
sign = 1
ldav = ld
If ldav - laval <> 0 Then
    sign = Abs((ldav - laval)) / ((ldav - laval))

While Abs(ldav - laval) > 0.001 And ldav > 0 And istop < 2000  '0.001
    istop = istop + 1
  
    If sign * (ldav - laval) < 0 Then
        ddav = ddav / 2#
        sign = sign * -1
    End If
         hautdo = hautdo - sign * ddav
     vcri = (qcri / 1000) / ((3.14159 * (dav ^ 2)) / 4)
    icri = (vcri / (kav * (dav / 4) ^ (2 / 3))) ^ 2
'    If TAV = 0 Then
        tav = hautdo + longdo * pentedo
'    End If
    a = rech_do_A(tav, dav)
    
'    Debug.Print (tav - (a * (vcri ^ 2) / 19.62) - dav - tcr), (icri - iradav)
    ldav = (tav - (a * (vcri ^ 2) / 19.62) - dav - tcr) / (icri - iradav)
  

Wend
End If
If istop >= 2000 Then
    MsgBox "non convergence", vbOKOnly, " rech hauteur"
End If
If ldav > 0 And istop < 2000 Then
rech_ham_do = hautdo
End If
Else
rech_ham_do = 0
End If
End Function


Public Function calc_debit_tr(ByRef tr As troncon, ByVal Q As Double) As debit_conduit
Dim qv As deb_vit
Dim ltc() As Variant, ct() As Variant, qvm(5) As Variant
Dim qcal As Double
Dim haut As Double, larg As Double, acce As Double, piezoamo As Double, piezoaval As Double
Dim vit_amo As Double, vit_ava As Double, chargeamo As Double, chargeaval As Double
Dim canal As conduite
Dim pentmot As Double

canal = tr.conduit
qv = debvit_ps(canal)
'Debug.Print qv.debit


qcal = Q
'If qcal < qv.debit * 1000 Then
If qcal < qv.debit Then
    Call cana(canal, ct)
    ltc = calc_par(canal)
    qvi = caltran1(qcal * 1000, ct, ltc)
'    Debug.Print "qvi : 1= débit"; qvi(1); " 2=vitesse"; qvi(2); "qvi"; " 3=acceleration"; qvi(3)
'    Debug.Print " 4=Largeur libre"; qvi(4); " 5 = hauteur; "; qvi(5); ""
'                vitmax = qvm(2)
'                qvm = caltran1(qps / 10#, ct, ltc)
'                vit10 = qvm(2)
'                qvm = caltran1(qps / 100#, ct, ltc)
'                vit100 = qvm(2)
larg = qvi(4)
acce = qvi(3)
   haut = qvi(5)
   vit_amo = qvi(2)
   vit_ava = vit_amo
    piezoamo = tr.radamo + haut
    piezoava = tr.radava + haut
'a revoir
    pentmot = tr.conduit.pente
    pentmot = pent_mot0(canal, qcal)
    calc_debit_tr.charge = False
Else
    haut = tr.conduit.Diametre
    larg = 0
    acce = 0

    vit_amo = qcal / (qv.debit / qv.vitesse)
   vit_ava = vit_amo
   piezoava = tr.radava + tr.conduit.Diametre
'  piezoamo = tr.radamo + tr.conduit.Diametre
   pentmot = pent_mot0(canal, qcal)
  piezoamo = piezoava + pentmot * canal.Longueur
    calc_debit_tr.charge = True
End If
    dcharge = vit_ava * vit_ava / (2 * 9.81)

chargeamo = piezoamo + (vit_amo ^ 2 / (2 * 9.81))
chargeava = piezoava + (vit_ava ^ 2 / (2 * 9.81))
'  uc_g.dess_lign tr.Absamo, chargeamo, tr.Absava, chargeava, ocolor

calc_debit_tr.debit = qcal
calc_debit_tr.vitesse = vit_amo
calc_debit_tr.hauteur = haut
calc_debit_tr.largeurlibre = larg
calc_debit_tr.acceleration = acce
calc_debit_tr.chargeamo = chargeamo
calc_debit_tr.chargeava = chargeava
calc_debit_tr.dcharge = dcharge
calc_debit_tr.pentemotrice = pentmot
calc_debit_tr.piezoamo = piezoamo
calc_debit_tr.piezoava = piezoava
If vit_amo > 0 Then
calc_debit_tr.surface = qcal / vit_amo
Else
calc_debit_tr.surface = 0
End If
calc_debit_tr.hautava = haut
calc_debit_tr.hautamo = haut
calc_debit_tr.vitamo = vit_amo
calc_debit_tr.vitava = vit_ava
End Function
Public Function calc_hauteur_tr(ByRef tr As troncon, ByVal h As Double) As debit_conduit
Dim qv As deb_vit
Dim ltc() As Variant, ct() As Variant, qvm(5) As Variant
Dim qcal As Double
Dim haut As Double, larg As Double, acce As Double, piezoamo As Double, piezoaval As Double
Dim vit_amo As Double, vit_ava As Double, chargeamo As Double, chargeaval As Double
Dim canal As conduite
Dim pentmot As Double

canal = tr.conduit
qv = debvit_ps(canal)
'Debug.Print qv.debit


'qcal = Q
'If qcal < qv.debit * 1000 Then
If h < canal.Diametre Then
    Call cana(canal, ct)
    ltc = calc_par(canal)
'h = 0.124

     beta = 2 * arccosinus((1 - 2 * h / canal.Diametre))
'     Debug.Print (beta - Sin(beta))
    Sc = (1# / 8#) * (canal.Diametre ^ 2) * (beta - Sin(beta))
    lc = canal.Diametre * Sin(beta / 2)
   Dim peri As Double
   peri = beta * canal.Diametre / 2
'qcal = canal.rugosite * canal.pente ^ 0.5 * Sc * (Sc / peri) ^ (3# / 4#)
qcal = canal.rugosite * canal.pente ^ 0.5 * Sc * (Sc / peri) ^ (2# / 3#)
'qcal = canal.rugosite * canal.pente ^ 0.5 * canal.Diametre ^ (8 / 3) * 3 / 4 * (h / canal.Diametre) ^ 2 * (1 - 7 / 12 * (h / canal.Diametre) ^ 2)
'Debug.Print qcal
    qvi = caltran1(qcal * 1000, ct, ltc)
'    Debug.Print "qvi : 1= débit"; qvi(1); " 2=vitesse"; qvi(2); "qvi"; " 3=acceleration"; qvi(3)
'    Debug.Print " 4=Largeur libre"; qvi(4); " 5 = hauteur; "; qvi(5); ""
'                vitmax = qvm(2)
'                qvm = caltran1(qps / 10#, ct, ltc)
'                vit10 = qvm(2)
'                qvm = caltran1(qps / 100#, ct, ltc)
'                vit100 = qvm(2)
larg = qvi(4)
acce = qvi(3)
   haut = qvi(5)
   vit_amo = qvi(2)
   vit_ava = vit_amo
'    piezoamo = tr.radamo + haut
'    piezoava = tr.radava + haut
''a revoir
'    pentmot = tr.conduit.pente
'    pentmot = pent_mot0(canal, qcal)
'    calc_debit_tr.charge = False
'Else
'    haut = tr.conduit.Diametre
'    larg = 0
'    acce = 0
'
'    vit_amo = qcal / (qv.debit / qv.vitesse)
'   vit_ava = vit_amo
'   piezoava = tr.radava + tr.conduit.Diametre
''  piezoamo = tr.radamo + tr.conduit.Diametre
'   pentmot = pent_mot0(canal, qcal)
'  piezoamo = piezoava + pentmot * canal.Longueur
'    calc_debit_tr.charge = True
End If
'    dcharge = vit_ava * vit_ava / (2 * 9.81)
'
'chargeamo = piezoamo + (vit_amo ^ 2 / (2 * 9.81))
'chargeava = piezoava + (vit_ava ^ 2 / (2 * 9.81))
''  uc_g.dess_lign tr.Absamo, chargeamo, tr.Absava, chargeava, ocolor
'
'calc_debit_tr.debit = qcal
'calc_debit_tr.vitesse = vit_amo
'calc_debit_tr.hauteur = haut
'calc_debit_tr.largeurlibre = larg
'calc_debit_tr.acceleration = acce
calc_hauteur_tr.debit = qcal
calc_hauteur_tr.vitesse = vit_amo
calc_hauteur_tr.hauteur = haut
calc_hauteur_tr.largeurlibre = larg
calc_hauteur_tr.acceleration = acce
'calc_debit_tr.chargeamo = chargeamo
'calc_debit_tr.chargeava = chargeava
'calc_debit_tr.dcharge = dcharge
'calc_debit_tr.pentemotrice = pentmot
'calc_debit_tr.piezoamo = piezoamo
'calc_debit_tr.piezoava = piezoava
If vit_amo > 0 Then
    calc_hauteur_tr.surface = qcal / vit_amo
Else
    calc_hauteur_tr.surface = 0
End If
calc_hauteur_tr.hautava = haut
calc_hauteur_tr.hautamo = haut
calc_hauteur_tr.vitamo = vit_amo
'calc_hauteur.vitava = vit_ava
End Function
Public Function print_erreur(ByVal message As String)
    Dim reponse As Integer
    reponse = MsgBox(message, , "Ouverture d'un fichier")
End Function
Public Function dess_courbe_debit_tr(ByRef troamo As troncon, _
    ByVal Qmax As Double, ByVal Titre As String)
Dim qv As deb_vit, qvps_amo As deb_vit
    qvps_amo = debvit_ps(troamo.conduit)
'    owner.fdessin.UC_graphique1.dess_lign 0, ebchute.Qmax * 1000, ebchute.dam, ebchute.Qmax * 1000, couleur.rouge, 1
'    Call calc_courbe_debit_tr(owner.fdessin.UC_graphique1, ebchute.tron_amo)
'    owner.fdessin.UC_graphique1.dess_lign 0, ebchute.Qmax * 1000, ebchute.dam, ebchute.Qmax * 1000, couleur.rouge, 1
'    owner.fdessin.UC_graphique1.dess_lign 0, qvps_amo.debit * 1000, ebchute.dam, qvps_amo.debit * 1000, couleur.orange, 1
    Frm_graph.Caption = Titre
    If Qmax > 0 Then
    Frm_graph.UC_graphique1.dess_lign 0, Qmax * 1000, troamo.conduit.Diametre * 1000, Qmax * 1000, couleur.rouge, 1
    Else
    Frm_graph.UC_graphique1.dess_lign 0, Qmax * 1000, troamo.conduit.Diametre * 1000, Qmax * 1000, couleur.noir, 1
    End If
    Call calc_courbe_debit_tr(Frm_graph.UC_graphique1, troamo)
    If Qmax > 0 Then
    Frm_graph.UC_graphique1.dess_lign 0, Qmax * 1000, troamo.conduit.Diametre * 1000, Qmax * 1000, couleur.rouge, 1
    End If
    Frm_graph.UC_graphique1.dess_lign 0, qvps_amo.debit * 1000, troamo.conduit.Diametre * 1000, qvps_amo.debit * 1000, couleur.orange, 1
    Frm_graph.UC_graphique1.init_lbvb "l/s"
    Frm_graph.UC_graphique1.init_lbhg "mn"
    Frm_graph.Show 1

End Function

Public Function calc_courbe_debit_tr(uc_g As UC_graphique, ByRef tr As troncon) As debit_conduit
Dim qv As deb_vit
Dim ltc() As Variant, ct() As Variant, qvm() As Variant
Dim qcal As Double
Dim haut As Double, larg As Double, acce As Double, piezoamo As Double, piezoaval As Double
Dim vit_amo As Double, vit_ava As Double, chargeamo As Double, chargeaval As Double
Dim canal As conduite
Dim pentmot As Double
Dim coul As typ_Couleur
coul = couleur
canal = tr.conduit
qv = debvit_ps(canal)
'Debug.Print qv.debit
    Call cana(canal, ct)
    ltc = calc_par(canal)
    qvi = caltran_courbe(ct, ltc)
'    Debug.Print "qvi : 1= débit"; qvi(10, 1); " 2=vitesse"; qvi(10, 2); "qvi"; " 3=acceleration"; qvi(10, 3)
'    Debug.Print " 4=Largeur libre"; qvi(10, 4); " 5 = hauteur; "; qvi(10, 5); ""
ReDim qvm(UBound(qvi), 2)
For i = 1 To UBound(qvm)
qvm(i, 1) = qvi(i, 5) * 1000
qvm(i, 2) = qvi(i, 1) * 1000
Next
uc_g.graphique_clear
uc_g.init_titleh ""
uc_g.init_titleb ""
uc_g.redef_cadrs 600, 500, 200
uc_g.init_arrondi_y 1
uc_g.init_MaxX 0
uc_g.init_MaxY 0
uc_g.init_MinX 0
uc_g.init_MinY 0
uc_g.init_MaxXn qvm
uc_g.init_MaxYn qvm
 uc_g.init_EchYn 0.9
 uc_g.init_EchXn 1#
'   uc_g.dess_cadre 10, 2, 100, 2, 2, 1, 10, 2, 100
   uc_g.dess_cadre 10, 2, 100, 0, 0, 0, 10, 2, 100
 
uc_g.dess_poly qvm, "N", couleur.bleu, 1
 '   uc_g.init_EchYi 0.3




'                vitmax = qvm(2)
'                qvm = caltran1(qps / 10#, ct, ltc)
'                vit10 = qvm(2)
'                qvm = caltran1(qps / 100#, ct, ltc)
'                vit100 = qvm(2)
'calc_debit_tr.debit = qcal
'calc_debit_tr.vitesse = vit_amo
'calc_debit_tr.hauteur = haut
'calc_debit_tr.largeurlibre = larg
'calc_debit_tr.surface = qcal / vit_amo
  
End Function
Function caltran_courbe(b, ltc As Variant) As Variant
Dim lqv() As Double
Dim i As Integer
Dim Q, lm, haut, pen, ray, Ks, kg, lkg, louvr, epsi As Double
ReDim lqv(UBound(ltc), 5)
louvr = b(1)
ray = b(2) / 2#
pen = b(3)
Ks = b(4)
epsi = 0.005
kg = Ks * Sqr(pen)
lkg = kg * ray ^ (8# / 3#)
i = 1
For i = 1 To UBound(ltc)
a = ltc(i, 1) * (1000# * lkg)
ch = lkg / (ray * ray) * (ltc(i, 3))
lm = ray * (ltc(i, 4))
haut = ray * (ltc(i, 5))
'debug.print"Debit =",q," Vs= ",Vs," DT =",Louvr/Vs," s","lm = ",lm,"haut = ",haut
vs = kg * ray ^ (2# / 3#) * (ltc(i, 2))
' Debug.Print "Debit =", a, " Vs= ", Vs, " DT =", louvr / Vs, " s Celerite =", ch, "m/s"
lqv(i, 1) = a / 1000#
lqv(i, 2) = vs
lqv(i, 3) = ch
lqv(i, 4) = lm
lqv(i, 5) = haut
Next
'return(lqv)
caltran_courbe = lqv
End Function


Public Function inter_piezo_eau0(ByRef tr As troncon, ByRef tr_res As debit_conduit)
Dim xam As Double, yam As Double, xav As Double, yav As Double, pcana As Double
Dim xamp As Double, yamp As Double, xavp As Double, yavp As Double, ppiezo As Double
Dim xi As Double, yi As Double
Dim p As points
Dim ok As Boolean
' intersection avec la generaatrice superieure
    ppiezo = tr_res.pentemotrice
xam = tr.Absamo
yam = tr.radamo + tr.conduit.Diametre
xav = tr.Absava
yav = tr.radava + tr.conduit.Diametre
xavp = tr.Absava
yavp = maximum(tr_res.zphe_ava, tr.radava + tr_res.hauteur)
xamp = tr.Absamo
yamp = yavp + ((tr.conduit.Longueur) * ppiezo)
ok = inters(xam, yam, xav, yav, xamp, yamp, xavp, yavp, xi, yi)
' intersection avec la ligne d'eau theorique

If xi > xam And xi < xav Then
tr_res.piezointer0.X = xi
tr_res.piezointer0.Y = yi
tr_res.p_Eau_inter0.X = xi
tr_res.p_Eau_inter0.Y = yi

Else
tr_res.piezointer0.X = xavp
tr_res.piezointer0.Y = yavp
tr_res.p_Eau_inter0.X = xav
tr_res.p_Eau_inter0.Y = maximum(tr.radava + tr_res.hauteur, minimum(tr_res.zphe_ava, tr.radava + tr.conduit.Diametre))

End If


xam = tr.Absamo
yam = tr.radamo + tr_res.hauteur
xav = tr.Absava
yav = tr.radava + tr_res.hauteur
'If tr_res.charge Then
    ppiezo = tr_res.pentemotrice
'Else
 '   ppiezo = tr.conduit.pente
'End If
xavp = tr.Absava
yavp = maximum(tr_res.zphe_ava, yav)
xamp = tr.Absamo
yamp = yavp + ((tr.conduit.Longueur) * ppiezo)
ok = inters(xam, yam, xav, yav, xamp, yamp, xavp, yavp, xi, yi)
tr_res.zeau_amo.X = xam
tr_res.zeau_ava.X = xav
tr_res.piezoava = yavp
If yamp <= yam Then
    tr_res.piezoamo = yam
    tr_res.zeau_amo.Y = yam
Else
    tr_res.piezoamo = yamp
    If yamp > (tr.conduit.Diametre + tr.radamo) Then
        tr_res.zeau_amo.Y = tr.conduit.Diametre + tr.radamo
        Else
        tr_res.zeau_amo.Y = yamp
    End If
End If
If yavp > (tr.conduit.Diametre + tr.radava) Then
    tr_res.zeau_ava.Y = tr.conduit.Diametre + tr.radava
ElseIf yavp > yav Then
        tr_res.zeau_ava.Y = yavp
Else
    tr_res.zeau_ava.Y = yav
End If

'Debug.Print xi, yi

If xi < xam Then
    xi = xam
    yi = yamp
    tr_res.piezointer.X = xi
    tr_res.piezointer.Y = yi
    If yi > (tr.conduit.Diametre + tr.radamo) Then
        yi = tr.conduit.Diametre + tr.radamo
    Else
        yi = yamp
    End If
ElseIf xi > xav Then
    xi = xav
    yi = yavp
    tr_res.piezointer.X = xi
    tr_res.piezointer.Y = yi
    yi = yav
    Else
        tr_res.piezointer.X = xi
        tr_res.piezointer.Y = yi

End If

tr_res.p_Eau_inter.X = xi
tr_res.p_Eau_inter.Y = yi
tr_res.hautamo = tr_res.zeau_amo.Y - tr.radamo
tr_res.hautava = tr_res.zeau_ava.Y - tr.radava


End Function
Public Sub inter_charge_tr(ByRef tr As troncon, ByRef res_conduit As debit_conduit)
    Dim s As Double, v As Double, beta As Double, zhe As Double
    zhe = res_conduit.zeau_amo.Y - tr.radamo
    If zhe / tr.conduit.Diametre < 1 Then
        beta = 2 * arccosinus((1 - 2 * zhe / tr.conduit.Diametre))
    Else
        beta = 2 * pi
    End If
        s = (1 / 8) * (tr.conduit.Diametre ^ 2) * (beta - Sin(beta))
        v = res_conduit.debit / s
        res_conduit.vitamo = v
        res_conduit.dcharge = (v ^ 2) / (2 * 9.81)
        res_conduit.chargeamo = res_conduit.piezoamo + res_conduit.dcharge
    zhe = res_conduit.zeau_ava.Y - tr.radava
    If zhe / tr.conduit.Diametre < 1 Then
        beta = 2 * arccosinus((1 - 2 * (zhe / tr.conduit.Diametre)))
    Else
        beta = 2 * pi
    End If
        s = (1 / 8) * (tr.conduit.Diametre ^ 2) * (beta - Sin(beta))
        v = res_conduit.debit / s
        res_conduit.vitava = v
        res_conduit.dcharge = (v ^ 2) / (2 * 9.81)
        res_conduit.chargeava = maximum(res_conduit.piezoava, res_conduit.zeau_ava.Y) + res_conduit.dcharge
        
      zhe = res_conduit.piezointer.Y - tr.radamo + (res_conduit.piezointer.X - tr.Absamo) * tr.conduit.pente
    If zhe / tr.conduit.Diametre < 1 Then
        beta = 2 * arccosinus(1 - 2 * zhe / tr.conduit.Diametre)
    Else
        beta = 2 * pi
    End If
        s = (1 / 8) * (tr.conduit.Diametre ^ 2) * (beta - Sin(beta))
        v = res_conduit.debit / s
        res_conduit.dcharge = (v ^ 2) / (2 * 9.81)
        res_conduit.chargeinter = res_conduit.piezointer.Y + res_conduit.dcharge
  If res_conduit.piezointer2.X = 0 And res_conduit.piezointer2.Y = 0 And res_conduit.piezointer1.X = 0 And res_conduit.piezointer1.Y = 0 Then
          res_conduit.piezointer2 = res_conduit.piezointer
          res_conduit.piezointer1 = res_conduit.piezointer0
 End If
      zhe = res_conduit.piezointer2.Y - tr.radamo + (res_conduit.piezointer2.X - tr.Absamo) * tr.conduit.pente
    If zhe / tr.conduit.Diametre < 1 Then
        beta = 2 * arccosinus(1 - 2 * zhe / tr.conduit.Diametre)
    Else
        beta = 2 * pi
    End If
        s = (1 / 8) * (tr.conduit.Diametre ^ 2) * (beta - Sin(beta))
        v = res_conduit.debit / s
        res_conduit.dcharge = (v ^ 2) / (2 * 9.81)
        res_conduit.chargeinter2 = res_conduit.piezointer2.Y + res_conduit.dcharge
      
      zhe = res_conduit.piezointer1.Y - tr.radamo + (res_conduit.piezointer1.X - tr.Absamo) * tr.conduit.pente
    If zhe / tr.conduit.Diametre < 1 Then
        beta = 2 * arccosinus(1 - 2 * zhe / tr.conduit.Diametre)
    Else
        beta = 2 * pi
    End If
        s = (1 / 8) * (tr.conduit.Diametre ^ 2) * (beta - Sin(beta))
        v = res_conduit.debit / s
        res_conduit.dcharge = (v ^ 2) / (2 * 9.81)
        res_conduit.chargeinter1 = res_conduit.piezointer1.Y + res_conduit.dcharge

      zhe = res_conduit.piezointer0.Y - tr.radamo + (res_conduit.piezointer0.X - tr.Absamo) * tr.conduit.pente
    If zhe / tr.conduit.Diametre < 1 Then
        beta = 2 * arccosinus(1 - 2 * zhe / tr.conduit.Diametre)
    Else
        beta = 2 * pi
    End If
        s = (1 / 8) * (tr.conduit.Diametre ^ 2) * (beta - Sin(beta))
        v = res_conduit.debit / s
        res_conduit.dcharge = (v ^ 2) / (2 * 9.81)
        res_conduit.chargeinter0 = res_conduit.piezointer0.Y + res_conduit.dcharge
  
        

End Sub
Function create_fs(ByVal nfc As String) As file_spec
Dim fs As Object
Dim s As String
Dim fsc As file_spec
Dim f As File
Dim d As Drive
Set fs = CreateObject("Scripting.FileSystemObject")

'Debug.Print fs.Getabsolutepathname(s)
'Debug.Print fs.getdrivename(s)

'Debug.Print fs.Getabsolutepathname(s)

'Debug.Print fs.Getfilename(s)
'Debug.Print fs.Getextensionname(s)
'Debug.Print fs.Getparentfoldername(s)
With fsc
    .lecteur = fs.GetDriveName(nfc)
    .Chemin = Mid$(fs.GetParentFolderName(nfc), Len(.lecteur) + 1)
    .nom = fs.GetFileName(nfc)
    .extension = fs.GetExtensionName(nfc)
    .nomcomplet = fs.GetAbsolutePathName(nfc)
    If Trim(fs.GetDriveName(nfc)) <> "" Then
    Set d = fs.GetDrive(fs.GetDriveName(nfc))
'   type 0=inconnu,1=amovible,2=fixe,3=reseau,4=cd-rom,5=disque RAM
        .dr_type = d.DriveType
    Else
        .dr_type = 0
    End If
    If Trim(fs.GetFileName(nfc)) <> "" Then
        If Dir(fs.GetAbsolutePathName(nfc)) <> "" Then
            Set f = fs.GetFile(fs.GetAbsolutePathName(nfc))
' attributes 1=lecture seule
            .f_attr = f.Attributes
            .f_size = f.Size
        Else
        .f_attr = 0
        End If
    Else
        .f_attr = 0
    End If
End With
 
 create_fs = fsc
End Function

Public Function rec_list(ByVal nom As String) As Variant
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
rec_list = list
End Function

Public Function rec_bassin(ByVal nom1 As String, ByVal nomt As String) As Boolean
Dim lhFicDbf1 As Long
Dim nom As String
Dim za As st_save
Dim za1 As st_save1
'nom = chemin_app + "bassin.bin"
rec_bassin = False
lhFicDbf1 = FreeFile
Open nom_fich For Random Access Read Lock Read Write As #lhFicDbf1 Len = Len(za1)
'Open nom For Random Access Read As #lhFicDbf1 Len = Len(za)
Do While Not EOF(lhFicDbf1)
    Get #lhFicDbf1, , za1
    If Not EOF(lhFicDbf1) Then
        If Trim(za1.type) = nomt Then
            za = za1.stsave
            If Trim(za.nom) = Trim(nom1) Then
                ebv = za.bv
                eph = za.hydro
                ehyd = za.hydro1
                rec_bassin = True
            End If
        End If
    End If
Loop
Close #lhFicDbf1
End Function


