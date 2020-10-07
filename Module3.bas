Attribute VB_Name = "Module3"
Option Explicit
Private Function Rec_Mes(nom As String, Index As Integer)
Dim mes As String
mes = ""
Select Case nom
    Case Is = "Lb_car_ep"
        Select Case Index
            Case Is = 0
                mes = "Temps de concentration Tc"
            Case Is = 1
                mes = "Temps de concentration Tc"
            Case Is = 2
                mes = "Temps de concentration Tc"
            Case Is = 3
                mes = "Coefficients de ruissellement Cr"
        End Select
    Case Is = "Lb_car_eu"
        Select Case Index
            Case Is = 0
                mes = "Débit des eaux usées domestiques"
            Case Is = 1
                mes = "Débit des eaux usées domestiques"
            Case Is = 2
                mes = "Débit des eaux claires parasites"
        End Select
    Case Is = "Lb_debit"
        Select Case Index
            Case Is = 0
                mes = "Méthode superficielle de Caquot"
            Case Is = 1
               mes = "Méthode Rationnelle "
            Case Is = 2
               mes = "Méthode de l'hydrogramme"
        End Select
    Case Is = "Lb_debit1"
        Select Case Index
            Case Is = 0
                mes = "Débit des eaux usées domestiques"
            Case Is = 1
                mes = "Débit des eaux usées domestiques"
            Case Is = 2
                mes = "Débit des eaux claires parasites"
            Case Is = 3
                mes = "Le débit de référence QREF"
            Case Is = 4
                mes = "Le débit de référence QREF"
            Case Is = 5
                mes = "Le débit d'orage QORA"
            Case Is = 6
                mes = "Méthode de l'hydrogramme"
        End Select
    Case Is = "Lb_par_ep"
        mes = "Courbes Intensité-Durée-Fréquence(IDF)"
    Case Is = "Lb_par_eu"
        Select Case Index
            Case Is = 0
                mes = "Le débit de référence QREF"
            Case Is = 1
                mes = "Débit des eaux usées domestiques"
            Case Is = 2
                mes = "Débit des eaux usées domestiques"
        End Select
    Case Is = "Lb_carep_rur"
        Select Case Index
            Case Is = 0
                mes = "Estimation des pertes initiales"
            Case Is = 1
                mes = "Estimation des pertes continues"
            Case Is = 2
                mes = "Estimation des pertes continues"
            Case Is = 3
                mes = "Estimation des pertes continues"
            Case Is = 4
                mes = "Modèle de ruissellement ' réservoir linéaire '"
        End Select
    Case Is = "Frm_cep"
        mes = "Caractéristiques d'un BV"
    Case Is = "Frm_ceu"
        mes = "Méthodes d'évaluation des débits de temps sec"
    Case Is = "Frm_debit"
        mes = "Débits caractéristiques"

End Select
Rec_Mes = mes
End Function
Function verif_do_charge(ByRef edev As deversoir, ByRef edo_r As deversoir_resultat, ByRef tr As troncon, ByVal Qpl As Double, ByVal Qrin As Double, ByRef uc_g As UC_graphique) As String
Dim ok As Boolean, nok0 As Boolean, nok1 As Boolean
'Dim edo_res As deversoir_resultat
On Error GoTo ErrorHandler
ok = True
Dim HM As Double, Ham As Double, Hav As Double, Haam As Double, Haav As Double, Haavd As Double
Dim dHa As Double
Dim Tram As Double, c As Double
Dim tam As Double
Dim res_conduit As debit_conduit
Dim ed As deversoir
Dim beta As Double, s As Double, v As Double
Dim ecoulam As String
Dim Qav As Double, Qdev As Double
Dim i0 As Integer
Dim sres As String
Dim okTram As Boolean
Dim i As Integer
Dim dhmin As Double, tammin As Double

i = 0
sres = "  Fonctionnement à débit de pointe "
'Tram = edo_r.Tram
Tram = edessdo.Tram
'a revoir verification Tram/ haut d'eau< Tram<diametre ?
If Tram = 0 Then
    Tram = tr.conduit.Diametre * 0.9
End If
   
    res_conduit = calc_debit_tr(tr, Qpl)
tam = res_conduit.hauteur
tam = maximum(tam, edo.hauteur)
'fin a revoir
c = edo_r.c
nok0 = True
' debit en m3/s
Qav = Qrin * 1.3

While nok0 And i0 < 20
okTram = False
tam = res_conduit.hauteur
tam = maximum(tam, edo.hauteur)

Qdev = Qpl - Qav
    'Qav = 0.172
    
    ed = edev
    tammin = Tram
    dhmin = 10000
    While Not okTram And tam <= Tram
    
    'a revoir ? 0.01 pente amont-aval de la ligne d'energie
    dHa = 0.01 * ed.Longueur
    
    HM = (0.85 * Qdev / (c * ed.Longueur)) ^ (2# / 3#)
    Ham = tam - ed.hauteur
    Hav = (4 * HM - Ham) / 3#
    Haav = Hav + ed.hauteur
    'HAam = rech_haut_do_vam(dam, qcri / 0.3)
    ' calcul de vitesse amont debit pluie
    
    ' verification vitesse d'écoulement amont pour qcri
    If tam / tr.conduit.Diametre < 1 Then
        beta = 2 * arccosinus(1 - 2 * tam / tr.conduit.Diametre)
    Else
        beta = 2 * pi
    End If
        s = (1 / 8) * (tr.conduit.Diametre ^ 2) * (beta - Sin(beta))
        v = Qpl / s
        
  
    
    Haam = tam + (v ^ 2) / (2 * 9.81)
    Haavd = Haam - dHa
    If (Haavd - Haav) < dhmin And (Haavd - Haav) > 0 Then
            dhmin = (Haavd - Haav)
            tammin = tam
    End If
    
'    If (Haam - Haav) < dHa Then
'    sres = sres + Chr$(13) + Chr$(10) + " Perte de charge Amont Aval : augmenter la longueur du do "
'    End If
'    If ((Haam - Haav) - dHa) > 0.01 Then
    If Abs((Haavd - Haav)) < 0.0001 And (Haavd - Haav) > 0 Then
'    MsgBox "Haam = " + Str(Haam) + Chr$(13) + " Haav disponible  = " + Str(Haavd) + Chr$(13) + " Haav   = " + Str(Haav), vbOKOnly, "verification charge"
    okTram = True
    Else
  '  Tram = Tram - (Haavd - Haav) / 100
'    Tam = Tam + Abs(Haavd - Haav) / 10
    tam = tam + 0.001
    
    okTram = False
    End If
    
    Wend
    
 '   calcul a tammin
    tam = tammin
        
        
        dHa = 0.01 * ed.Longueur
    
    HM = (0.85 * Qdev / (c * ed.Longueur)) ^ (2# / 3#)
    Ham = tam - ed.hauteur
    Hav = (4 * HM - Ham) / 3#
    Haav = Hav + ed.hauteur
    'HAam = rech_haut_do_vam(dam, qcri / 0.3)
    ' calcul de vitesse amont debit pluie
    
    ' verification vitesse d'écoulement amont pour qcri
    If tam / tr.conduit.Diametre < 1 Then
        beta = 2 * arccosinus(1 - 2 * tam / tr.conduit.Diametre)
    Else
        beta = 2 * pi
    End If
        s = (1 / 8) * (tr.conduit.Diametre ^ 2) * (beta - Sin(beta))
        v = Qpl / s
        
      ' calcul ecoulement amont (torrentiel ou fluvial)
    ecoulam = calcul_ecoul(Qpl, tr.conduit.Diametre, beta)
    
    Haam = tam + (v ^ 2) / (2 * 9.81)
    Haavd = Haam - dHa
    If (Haam - Haav) < dHa Then
    sres = sres + Chr$(13) + Chr$(10) + " Perte de charge Amont Aval : augmenter la longueur du do "
    End If

    
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
    While nok1 And i1 < 50
    Vavp = Qavp / Sav
    imot = pent_mot0(ed.tron_ava.conduit, Qavp)
    irad = ed.tron_ava.conduit.pente
    Tavp = (a * Vavp ^ 2 / (2 * 9.81)) + (ed.tron_ava.conduit.Longueur * (imot - irad)) + ed.tron_ava.conduit.Diametre
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
    i1 = i1 + 1
    Wend
    If Abs(Qav - Qavp) < 0.0001 Then
        nok0 = False
    End If
    Qav = Qavp + (Qav - Qavp) / 4#
    i0 = i0 + 1
Wend

With edo_r
    .Tram = tam
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
sres = sres + Chr(13) + Chr$(10) + " Ecoulement canalisation amont " + ecoulam
sres = sres + Chr(13) + Chr$(10) + " Débit dans la conduite étranglée " + ajout_zero(Trim(str(Round(Qav, 3)))) + "m3/s"
sres = sres + Chr(13) + Chr$(10) + " Débit déversé : " + ajout_zero(Trim(str(Round(Qdev, 3)))) + " m3/s"

resudev.debetranglee = ajout_zero(Trim(str(Round(Qav, 3))))
resudev.debdeverse = ajout_zero(Trim(str(Round(Qdev, 3))))

verif_do_charge = sres
Exit Function
ErrorHandler:
MsgBox "erreur", vbOKOnly, "Module verification charge"
verif_do_charge = ""
End Function


Public Sub dessin_do_debpointe(ByRef uc_g As UC_graphique, ByVal okcharge As Boolean, ByVal okpiezo As Boolean, ByVal okeau As Boolean)
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
'    zplam_av = edo.radamo + edessdo.Tram
    zplam_av = edo.radamo + edo_res.Tram
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
'            uc_g.dess_lign res_conduit.piezointer.x, res_conduit.piezointer.y, res_conduit.piezointer0.x, res_conduit.piezointer0.y, couleur.orange, 2
            uc_g.dess_lign res_conduit.piezointer.X, res_conduit.piezointer.Y, res_conduit.piezointer2.X, res_conduit.piezointer2.Y, couleur.orange, 2
            uc_g.dess_lign res_conduit.piezointer2.X, res_conduit.piezointer2.Y, res_conduit.piezointer1.X, res_conduit.piezointer1.Y, couleur.orange, 2
            uc_g.dess_lign res_conduit.piezointer1.X, res_conduit.piezointer1.Y, res_conduit.piezointer0.X, res_conduit.piezointer0.Y, couleur.orange, 2
            uc_g.dess_lign res_conduit.piezointer0.X, res_conduit.piezointer0.Y, res_conduit.zeau_ava.X, res_conduit.piezoava, couleur.orange, 2
        End If
        If okeau Then
'            ' uc_g.dess_lign tr.Absamo, zplam_am, res_conduit.piezointer.X, res_conduit.piezointer.Y, couleur.bleu, 2
'            ' uc_g.dess_lign res_conduit.piezointer.X, res_conduit.piezointer.Y, tr.Absava, zplam_av, couleur.bleu, 2
'            uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.zeau_amo.Y, res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, couleur.bleu, 2
'            uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, res_conduit.zeau_ava.X, res_conduit.zeau_ava.Y, couleur.bleu, 2
            uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.zeau_amo.Y, res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, couleur.bleu, 2
            uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, res_conduit.p_Eau_inter2.X, res_conduit.p_Eau_inter2.Y, couleur.bleu, 2
            uc_g.dess_lign res_conduit.p_Eau_inter2.X, res_conduit.p_Eau_inter2.Y, res_conduit.p_Eau_inter1.X, res_conduit.p_Eau_inter1.Y, couleur.bleu, 2
            uc_g.dess_lign res_conduit.p_Eau_inter1.X, res_conduit.p_Eau_inter1.Y, res_conduit.p_Eau_inter0.X, res_conduit.p_Eau_inter0.Y, couleur.bleu, 2
            uc_g.dess_lign res_conduit.p_Eau_inter0.X, res_conduit.p_Eau_inter0.Y, res_conduit.zeau_ava.X, res_conduit.zeau_ava.Y, couleur.bleu, 2
        End If
    'dessin charge
        'recalcul de charge amont en fonction de hauteur d'eau
        ' verification vitesse d'écoulement amont pour qcri
        If okcharge Then
            Call inter_charge_tr(tr, res_conduit)
            uc_g.dess_lign tr.Absamo, res_conduit.chargeamo, res_conduit.piezointer.X, res_conduit.chargeinter, couleur.rouge, 2
'        uc_g.dess_lign res_conduit.piezointer.x, res_conduit.chargeinter, res_conduit.piezointer0.x, res_conduit.chargeinter0, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer.X, res_conduit.chargeinter, res_conduit.piezointer2.X, res_conduit.chargeinter2, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer2.X, res_conduit.chargeinter2, res_conduit.piezointer1.X, res_conduit.chargeinter1, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer1.X, res_conduit.chargeinter1, res_conduit.piezointer0.X, res_conduit.chargeinter0, couleur.rouge, 2
            uc_g.dess_lign res_conduit.piezointer0.X, res_conduit.chargeinter0, tr.Absava, res_conduit.chargeava, couleur.rouge, 2
        End If
'dessin ligne d'eau sur la lame
    'dessin des lignes d'eau
        'uc_g.dess_lign edo.Absamo, edo.radamo + edo_res.Tram, edo.Absava, edo.radamo + edo.hauteur + edo_res.Hav, couleur.bleu, 2
        If okpiezo Then
            uc_g.dess_lign edo.Absamo, edo.radamo + edo_res.Tram, edo.Absava, edo.radamo + edo.hauteur + edo_res.Hav, couleur.orange, 2
        End If
        If okeau Then
            uc_g.dess_lign edo.Absamo, edo.radamo + edo_res.Tram, edo.Absava, edo.radamo + edo.hauteur + edo_res.Hav, couleur.bleu, 2
        End If
    'dessin charge
        If okcharge Then

            uc_g.dess_lign edo.Absamo, edo.radamo + edo_res.Haam, edo.Absava, edo.radamo + edo_res.Haav, couleur.rouge, 2
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
            uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.piezointer.Y, res_conduit.zeau_ava.X, res_conduit.piezoava, couleur.orange, 2
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
Public Sub dessin_decharge(ByRef uc_g As UC_graphique)
Dim zplam_am As Double, zplam_av As Double, zplav_am As Double, zplav_av As Double
'Dim qvm(5) As Variant, haut As Double, pentmot As Double
Dim tr As troncon  ', uc_g As UC_graphique
Dim res_conduit As debit_conduit
Dim qcal
'dessin de la ligne dans le deversoir
'dans la frmdessin
'Set uc_g = owner.fdessin.UC_graphique2
'zplam_av = edo.radamo + edo_res.Tram
'dessin troncon amont
    zplam_av = edo.radamo + edo_res.Tram
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
'        uc_g.dess_lign res_conduit.p_Eau_inter.x, res_conduit.p_Eau_inter.y, res_conduit.p_Eau_inter0.x, res_conduit.p_Eau_inter0.y, couleur.bleu, 2
            uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.piezointer.Y, res_conduit.piezointer2.X, res_conduit.p_Eau_inter2.Y, couleur.bleu, 2
            uc_g.dess_lign res_conduit.p_Eau_inter2.X, res_conduit.p_Eau_inter2.Y, res_conduit.piezointer1.X, res_conduit.p_Eau_inter1.Y, couleur.bleu, 2
            uc_g.dess_lign res_conduit.p_Eau_inter1.X, res_conduit.p_Eau_inter1.Y, res_conduit.piezointer0.X, res_conduit.p_Eau_inter0.Y, couleur.bleu, 2
        uc_g.dess_lign res_conduit.p_Eau_inter0.X, res_conduit.p_Eau_inter0.Y, res_conduit.zeau_ava.X, res_conduit.zeau_ava.Y, couleur.bleu, 2
   'dessin charge
        'recalcul de charge amont en fonction de hauteur d'eau
        ' verification vitesse d'écoulement amont pour qcri
        Call inter_charge_tr(tr, res_conduit)
        
        resudev.hqpluiem = ajout_zero(Trim(str(Round(res_conduit.hautamo, 3))))
        resudev.hqpluiemav = ajout_zero(Trim(str(Round(res_conduit.hautava, 3))))
        resudev.vqpluiem = ajout_zero(Trim(str(Round(res_conduit.vitamo, 3))))
        resudev.vqpluiemav = ajout_zero(Trim(str(Round(res_conduit.vitava, 3))))
        uc_g.dess_lign tr.Absamo, res_conduit.chargeamo, res_conduit.piezointer.X, res_conduit.chargeinter, couleur.rouge, 2
'        uc_g.dess_lign res_conduit.piezointer.x, res_conduit.chargeinter, res_conduit.piezointer0.x, res_conduit.chargeinter0, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer.X, res_conduit.chargeinter, res_conduit.piezointer2.X, res_conduit.chargeinter2, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer2.X, res_conduit.chargeinter2, res_conduit.piezointer1.X, res_conduit.chargeinter1, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer1.X, res_conduit.chargeinter1, res_conduit.piezointer0.X, res_conduit.chargeinter0, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer0.X, res_conduit.chargeinter0, tr.Absava, res_conduit.chargeava, couleur.rouge, 2
    
'dessin ligne d'eau sur la lame
    'dessin des lignes d'eau
        'uc_g.dess_lign edo.Absamo, edo.radamo + edo_res.Tram, edo.Absava, edo.radamo + edo.hauteur + edo_res.Hav, couleur.bleu, 2
        uc_g.dess_lign edo.Absamo, edo.radamo + edo_res.Tram, edo.Absava, edo.radamo + edo.hauteur + edo_res.Hav, couleur.bleu, 2
    'dessin charge
        uc_g.dess_lign edo.Absamo, edo.radamo + edo_res.Haam, edo.Absava, edo.radamo + edo_res.Haav, couleur.rouge, 2
     
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
'        uc_g.dess_lign res_conduit.p_Eau_inter.x, res_conduit.p_Eau_inter.y, res_conduit.p_Eau_inter0.x, res_conduit.p_Eau_inter0.y, couleur.bleu, 2
        uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, res_conduit.p_Eau_inter2.X, res_conduit.p_Eau_inter2.Y, couleur.bleu, 2
         uc_g.dess_lign res_conduit.p_Eau_inter2.X, res_conduit.p_Eau_inter2.Y, res_conduit.p_Eau_inter1.X, res_conduit.p_Eau_inter1.Y, couleur.bleu, 2
        uc_g.dess_lign res_conduit.p_Eau_inter1.X, res_conduit.p_Eau_inter1.Y, res_conduit.p_Eau_inter0.X, res_conduit.p_Eau_inter0.Y, couleur.bleu, 2
       uc_g.dess_lign res_conduit.p_Eau_inter0.X, res_conduit.p_Eau_inter0.Y, res_conduit.zeau_ava.X, res_conduit.zeau_ava.Y, couleur.bleu, 2
'dessin ligne piezo
'        uc_g.dess_lign res_conduit.zeau_amo.X, res_conduit.zeau_amo.Y, res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, couleur.orange, 2
'        uc_g.dess_lign res_conduit.p_Eau_inter.X, res_conduit.p_Eau_inter.Y, res_conduit.p_Eau_inter0.X, res_conduit.p_Eau_inter0.Y, couleur.orange, 2
'        uc_g.dess_lign res_conduit.p_Eau_inter0.X, res_conduit.p_Eau_inter0.Y, res_conduit.zeau_ava.X, res_conduit.piezoava, couleur.orange, 2
    
    'dessin charge
        ' dessin de la charge repris par inter_charge_pr
        Call inter_charge_tr(tr, res_conduit)
        
        resudev.hqdev = ajout_zero(Trim(str(Round(res_conduit.hautamo, 3))))
        resudev.hqdevav = ajout_zero(Trim(str(Round(res_conduit.hautava, 3))))
        resudev.vqdev = ajout_zero(Trim(str(Round(res_conduit.vitamo, 3))))
        resudev.vqdevav = ajout_zero(Trim(str(Round(res_conduit.vitava, 3))))
        uc_g.dess_lign tr.Absamo, res_conduit.chargeamo, res_conduit.piezointer.X, res_conduit.chargeinter, couleur.rouge, 2
'        uc_g.dess_lign res_conduit.piezointer.x, res_conduit.chargeinter, res_conduit.piezointer0.x, res_conduit.chargeinter0, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer.X, res_conduit.chargeinter, res_conduit.piezointer2.X, res_conduit.chargeinter2, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer2.X, res_conduit.chargeinter2, res_conduit.piezointer1.X, res_conduit.chargeinter1, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer1.X, res_conduit.chargeinter1, res_conduit.piezointer0.X, res_conduit.chargeinter0, couleur.rouge, 2
        uc_g.dess_lign res_conduit.piezointer0.X, res_conduit.chargeinter0, tr.Absava, res_conduit.chargeava, couleur.rouge, 2
End Sub


