Attribute VB_Name = "fonction_graphique"
Public Sub dess_piezo(ByRef uc_g As UC_graphique, ByRef tr As troncon, ByVal q As Double, ByRef ocolor As ColorConstants)
Dim qv As deb_vit
Dim ltc() As Variant, ct() As Variant, qvm(5) As Variant
Dim qcal As Double
Dim haut As Double, piezoamo As Double, piezoaval As Double
Dim canal As conduite
Dim pentmot As Double
canal = tr.conduit
qv = debvit_ps(canal)
'Debug.Print qv.debit
qcal = q
If qcal < qv.debit * 1000 Then
    Call cana(canal, ct)
    ltc = calc_par(canal)
    qvi = caltran1(qcal, ct, ltc)
'    Debug.Print "qvi : 1= débit"; qvi(1); " 2=vitesse"; qvi(2); "qvi"; " 3=acceleration"; qvi(3)
'    Debug.Print " 4=Largeur libre"; qvi(4); " 5 = hauteur; "; qvi(5); ""
'                vitmax = qvm(2)
'                qvm = caltran1(qps / 10#, ct, ltc)
'                vit10 = qvm(2)
'                qvm = caltran1(qps / 100#, ct, ltc)
'                vit100 = qvm(2)
   haut = qvi(5)
   
    piezoamo = tr.radamo + haut
    piezoava = tr.radava + haut
Else
    piezoava = tr.radava + tr.conduit.Diametre
'  piezoamo = tr.radamo + tr.conduit.Diametre
   pentmot = pent_mot0(canal, qcal / 1000#)
  piezoamo = piezoava + pentmot * canal.Longueur
 

End If
  uc_g.dess_lign tr.Absamo, piezoamo, tr.Absava, piezoava, ocolor, 2
  
End Sub
Public Sub dess_charge(ByRef uc_g As UC_graphique, ByRef tr As troncon, ByVal q As Double, ByRef ocolor As ColorConstants)
Dim qv As deb_vit
Dim ltc() As Variant, ct() As Variant, qvm(5) As Variant
Dim qcal As Double
Dim haut As Double, piezoamo As Double, piezoaval As Double
Dim vit_amo As Double, vit_ava As Double, chargeamo As Double, chargeaval As Double
Dim canal As conduite
Dim pentmot As Double
canal = tr.conduit
qv = debvit_ps(canal)
'Debug.Print qv.debit
qcal = q
If qcal < qv.debit * 1000 Then
    Call cana(canal, ct)
    ltc = calc_par(canal)
    qvi = caltran1(qcal, ct, ltc)
'    Debug.Print "qvi : 1= débit"; qvi(1); " 2=vitesse"; qvi(2); "qvi"; " 3=acceleration"; qvi(3)
'    Debug.Print " 4=Largeur libre"; qvi(4); " 5 = hauteur; "; qvi(5); ""
'                vitmax = qvm(2)
'                qvm = caltran1(qps / 10#, ct, ltc)
'                vit10 = qvm(2)
'                qvm = caltran1(qps / 100#, ct, ltc)
'                vit100 = qvm(2)
   haut = qvi(5)
   vit_amo = qvi(2)
   vit_ava = vit_amo
   
    piezoamo = tr.radamo + haut
    piezoava = tr.radava + haut
Else
    vit_amo = (qcal / 1000) / (qv.debit / qv.vitesse)
   vit_ava = vit_amo
   piezoava = tr.radava + tr.conduit.Diametre
'  piezoamo = tr.radamo + tr.conduit.Diametre
   pentmot = pent_mot0(canal, qcal / 1000#)
  piezoamo = piezoava + pentmot * canal.Longueur
 

End If
chargeamo = piezoamo + (vit_amo ^ 2 / (2 * 9.81))
chargeava = piezoava + (vit_ava ^ 2 / (2 * 9.81))
  uc_g.dess_lign tr.Absamo, chargeamo, tr.Absava, chargeava, ocolor, 2
  
End Sub

Public Sub init_graphdo(ByRef uc_graph As UC_graphique)
Dim ok As Boolean
Dim ecx As Double
Dim i As Integer
ok = False
uc_graph.graphique_clear
uc_graph.reinit 7, "Arial"
uc_graph.init_arrondi_X 2
uc_graph.init_arrondi_y 3
uc_graph.init_MinX 0#
uc_graph.init_MaxX edessdo.lgdisp
uc_graph.init_EchXn 1
ecx = uc_graph.lire_EchXn()
uc_graph.init_EchY ecx * 10
While Not ok
    ok = True
    uc_graph.init_MinY Int(edessdo.rdoav)
    uc_graph.init_Ech_MaxYn
    If uc_graph.lire_MaxYn < edessdo.rdoam + 1.3 * edessdo.tron_amo.conduit.Diametre Then
        uc_graph.init_EchY uc_graph.lire_EchYn / 2
        ok = False
    End If
Wend
uc_graph.dess_lign 0, edessdo.rdoam, edessdo.lgdisp, edessdo.rdoav, couleur.vert, 1 'vbgreen
'uc_graph.dess_tiret 0, edessdo.rdoav, edessdo.lgdisp, edessdo.rdoav, couleur.noir 'vbblack
   
End Sub
Public Sub init_graphdoor(ByRef uc_graph As UC_graphique)
Dim ok As Boolean
Dim ecx As Double, lg As Double
Dim i As Integer
lg = 2 * edo.Longueur
ok = False

uc_graph.graphique_clear
uc_graph.reinit 7, "Arial"
uc_graph.init_arrondi_X 2
uc_graph.init_arrondi_y 3
uc_graph.init_MinX 0#
uc_graph.init_MaxX lg * 2 + edoor_res.l_jetaval_b

uc_graph.init_EchXn 1
ecx = uc_graph.lire_EchXn()
uc_graph.init_EchY ecx * 10
uc_graph.init_MaxY edessdo.tron_amo.radava + edessdo.tron_ava.conduit.Diametre + edoor_res.hmin + Maxi(edessdo.tron_amo.conduit.Diametre, edessdo.tron_dech.conduit.Diametre)  ' 2 * edessdo.tron_amo.conduit.Diametre)
uc_graph.init_MinY Int(edo.radava) - 0.2
uc_graph.init_EchYn 1

   
End Sub
Public Sub dess_troncon(ByRef uc_g As UC_graphique, ByRef tr As troncon, ByRef ocolor As ColorConstants)
uc_g.dess_lign tr.Absamo, tr.radamo, tr.Absava, tr.radava, ocolor, 2
uc_g.dess_lign tr.Absamo, tr.radamo + tr.conduit.Diametre, tr.Absava, tr.radava + tr.conduit.Diametre, ocolor, 2
'uc_g.dess_circle tr.Absamo, tr.radamo + tr.conduit.Diametre / 2#, tr.conduit.Diametre / 2#, ocolor
'uc_g.dess_circle tr.Absava, tr.radava + tr.conduit.Diametre / 2#, tr.conduit.Diametre / 2#, ocolor
uc_g.dess_lign tr.Absamo, tr.radamo, tr.Absamo, tr.radamo + tr.conduit.Diametre, ocolor, 2
uc_g.dess_lign tr.Absava, tr.radava, tr.Absava, tr.radava + tr.conduit.Diametre, ocolor, 2
End Sub
Public Sub dess_troncon_or(ByRef uc_g As UC_graphique, ByRef tr As troncon, ByRef ocolor As ColorConstants, ByVal pos As String)
Dim xd As Double, yd As Double, xf As Double, yf As Double, yf0 As Double, dc As Double
Dim lg As Double
lg = 2 * edo.Longueur
If pos = "D" Then
    xd = edessdo.tron_amo.Absamo
    yd = edessdo.tron_amo.radava + ((edessdo.tron_amo.radamo - edessdo.tron_amo.radava) / edessdo.tron_amo.conduit.Longueur * lg)
    xf = xd + lg
    yf = edessdo.tron_amo.radava
    dc = edessdo.tron_amo.conduit.Diametre
ElseIf pos = "F" Then
    xd = lg + edo.Longueur
    yd = edessdo.tron_ava.radamo
    xf = xd + lg + (edoor_res.l_jetaval_b - edoor_res.l_chambre1)
 '   yf = edessdo.tron_ava.radamo - edessdo.tron_ava.conduit.pente * lg
    yf = edessdo.tron_ava.radamo - edessdo.tron_ava.conduit.pente * (lg + (edoor_res.l_jetaval_b - edoor_res.l_chambre1))
    dc = edessdo.tron_ava.conduit.Diametre
Else  'pos="C"
    xd = lg + edo.Longueur
    yd = edessdo.tron_dech.radamo
    xf = xd + lg
    yf = edessdo.tron_dech.radamo - edessdo.tron_dech.conduit.pente * lg
    dc = edessdo.tron_dech.conduit.Diametre

End If
If pos = "F" Then
'   uc_g.dess_rect xd, yd + dc, xf, yf + dc + edo.tav, couleur.gris_clair, 2
End If
uc_g.dess_lign xd, yd, xf, yf, ocolor, 2
uc_g.dess_lign xd, yd + dc, xf, yf + dc, ocolor, 2
yf0 = yd + dc
If pos = "D" Then
    uc_g.dess_lign xf, yf, xf, yf + dc, couleur.noir, 1
    uc_g.dess_lign xf, yf + dc, xf, yf + dc + 0.3, couleur.noir, 3
    xf = xf - lg / 4
    yf = yd + edessdo.tron_amo.conduit.Diametre + 0.05 '* 1.1
'uc_g.dess_lign xd, yd + edoor_res.Ham, xf, yf + edoor_res.hc, couleur.bleu, 1
'    uc_g.dess_text_aligne xf, "D", "C", yf, "Conduite arrivée (" + Format(edessdo.tron_amo.conduit.Diametre * 1000, "####0000") + " mm)", couleur.bleu
    uc_g.dess_text_aligne xf, "D", "C", yf, "Conduite arrivée (" + str(edessdo.tron_amo.conduit.Diametre * 1000) + " mm)", couleur.bleu
End If
If pos = "F" Then
    xd = xd + lg / 4
    yd = yd + edessdo.tron_ava.conduit.Diametre + 0.05 '* 1.1
'    uc_g.dess_text_aligne xd, "A", "C", yd, "Conduite départ (" + Format(edessdo.tron_ava.conduit.Diametre * 1000, "####0000") + " mm)", couleur.bleu
    uc_g.dess_text_aligne xd, "A", "C", yd, "Conduite départ (" + str(edessdo.tron_ava.conduit.Diametre * 1000) + " mm)", couleur.bleu
End If
If pos = "C" Then
    uc_g.dess_lign xd, yd, xd, yd + dc, couleur.noir, 1
    uc_g.dess_lign xd, yd + dc, xd, yd + dc + 0.3, couleur.noir, 3
    xd = xd + lg / 4
    yd = yd + edessdo.tron_dech.conduit.Diametre + 0.05 '* 1.1
'    uc_g.dess_text_aligne xd, "A", "C", yd, "Conduite déversement (" + Format(edessdo.tron_dech.conduit.Diametre * 1000, "####0000") + " mm)", couleur.bleu
    uc_g.dess_text_aligne xd, "A", "C", yd, "Conduite déversement (" + str(edessdo.tron_dech.conduit.Diametre * 1000) + " mm)", couleur.bleu
'uc_g.dess_lign xd, yd + edoor_res.Ham, xf, yf + edoor_res.hc, couleur.bleu, 1
End If
'If pos = "F" Then
'xd = edessdo.tron_amo.Absamo + lg + edoor_res.l_jetaval_h
'yd = edessdo.tron_ava.radamo - ((edoor_res.l_jetaval_h - edoor_res.l_chambre1) * edessdo.tron_ava.conduit.pente) + edoor_res.Hav
'xf = edessdo.tron_amo.Absamo + lg + edoor_res.l_jetaval_b + lg
'yf = edessdo.tron_ava.radamo - (((edoor_res.l_jetaval_b - edoor_res.l_chambre1) + lg) * edessdo.tron_ava.conduit.pente) + edoor_res.Hav
'uc_g.dess_lign xd, yd, xf, yf, couleur.bleu, 1
'End If

End Sub
Public Sub dess_debit_max_or(ByRef uc_g As UC_graphique)
Dim xd As Double, yd As Double, xf As Double, yf As Double, dc As Double
Dim lg As Double, xfd As Double, yfd As Double
Dim i As Integer, j As Integer, np As Integer
lg = 2 * edo.Longueur
'dessin ligne eau dans conduite amont(arrivée)
xd = edessdo.tron_amo.Absamo
yd = edessdo.tron_amo.radava + ((edessdo.tron_amo.radamo - edessdo.tron_amo.radava) / edessdo.tron_amo.conduit.Longueur * lg)
xf = xd + lg
yf = edessdo.tron_amo.radava
dc = edessdo.tron_amo.conduit.Diametre
uc_g.dess_lign xd, yd + edoor_res.Ham, xf, yf + edoor_res.hc, couleur.bleu, 1
'dessin ligne eau dans conduite aval(départ)
xd = edessdo.tron_amo.Absamo + lg + edoor_res.l_jetaval_h
yd = edessdo.tron_ava.radamo - ((edoor_res.l_jetaval_h - edoor_res.l_chambre1) * edessdo.tron_ava.conduit.pente) + edoor_res.Hav
xf = edessdo.tron_amo.Absamo + lg + edoor_res.l_jetaval_b + lg
yf = edessdo.tron_ava.radamo - (((edoor_res.l_jetaval_b - edoor_res.l_chambre1) + lg) * edessdo.tron_ava.conduit.pente) + edoor_res.Hav
uc_g.dess_lign xd, yd, xf, yf, couleur.bleu, 1
np = UBound(edoor_courbe_max_haut.dx)
'dessin courbe vers deversement
For i = 1 To np - 1
    If edoor_courbe_max_dever.dx(i) < edoor_res.l_ouverture Then
'    xd = edessdo.tron_amo.Absamo + lg + edoor_courbe_max_dever.dx(i)
'    yd = edessdo.tron_amo.radava - ((edoor_courbe_max_dever.dx(i) * edessdo.tron_amo.conduit.pente) + edoor_courbe_max_dever.dy(i))
'    xf = edessdo.tron_amo.Absamo + lg + edoor_courbe_max_dever.dx(i + 1)
'    yf = edessdo.tron_amo.radava - ((edoor_courbe_max_dever.dx(i + 1) * edessdo.tron_amo.conduit.pente) + edoor_courbe_max_dever.dy(i + 1))
    xd = edessdo.tron_amo.Absamo + lg + edoor_courbe_max_dever.dx(i)
    yd = edessdo.tron_amo.radava - edoor_courbe_max_dever.dy(i)
    xf = edessdo.tron_amo.Absamo + lg + edoor_courbe_max_dever.dx(i + 1)
    yf = edessdo.tron_amo.radava - edoor_courbe_max_dever.dy(i + 1)
    uc_g.dess_lign_point xd, yd, xf, yf, couleur.bleu
'    uc_g.dess_lign xd, yd, xf, yf, couleur.bleu, 1
    xfd = xf
    yfd = yf
    End If
Next
'dessin courbe max haut
j = 0
For i = 1 To np - 1
    If edoor_courbe_max_haut.dx(i) > edoor_res.l_ouverture Then
    j = j + 1
    If j = 1 Then
    xd = edessdo.tron_amo.Absamo + lg + edoor_res.l_ouverture
'    yd = edessdo.tron_amo.radava - (edoor_res.l_ouverture * edessdo.tron_amo.conduit.pente)
    yd = edessdo.tron_amo.radava - (edoor_res.l_ouverture * edoor_res.deltaa)
    xf = edessdo.tron_amo.Absamo + lg + edoor_courbe_max_haut.dx(i)
 '   yf = edessdo.tron_amo.radava - ((edoor_courbe_max_haut.dx(i) * edessdo.tron_amo.conduit.pente) + edoor_courbe_max_haut.dy(i))
    yf = edessdo.tron_amo.radava - edoor_courbe_max_haut.dy(i)
    uc_g.dess_lign xd, yd, xf, yf, couleur.bleu, 1
    End If
'    xd = edessdo.tron_amo.Absamo + lg + edoor_courbe_max_haut.dx(i)
'    yd = edessdo.tron_amo.radava - ((edoor_courbe_max_haut.dx(i) * edessdo.tron_amo.conduit.pente) + edoor_courbe_max_haut.dy(i))
'    xf = edessdo.tron_amo.Absamo + lg + edoor_courbe_max_haut.dx(i + 1)
'    yf = edessdo.tron_amo.radava - ((edoor_courbe_max_haut.dx(i + 1) * edessdo.tron_amo.conduit.pente) + edoor_courbe_max_haut.dy(i + 1))
    xd = edessdo.tron_amo.Absamo + lg + edoor_courbe_max_haut.dx(i)
    yd = edessdo.tron_amo.radava - ((edoor_courbe_max_haut.dx(i) * 0) + edoor_courbe_max_haut.dy(i))
    xf = edessdo.tron_amo.Absamo + lg + edoor_courbe_max_haut.dx(i + 1)
    yf = edessdo.tron_amo.radava - ((edoor_courbe_max_haut.dx(i + 1) * 0) + edoor_courbe_max_haut.dy(i + 1))
    uc_g.dess_lign xd, yd, xf, yf, couleur.bleu, 1
    End If
Next
xd = xf
yd = yf
xf = edessdo.tron_amo.Absamo + lg + edoor_res.l_jetaval_b + lg
yf = edessdo.tron_ava.radamo - (((edoor_res.l_jetaval_b - edoor_res.l_chambre1) + lg) * edessdo.tron_ava.conduit.pente) + edoor_res.Hav
uc_g.dess_lign xd, yd, xf, yf, couleur.bleu, 1

'dessin courbe max bas
For i = 1 To np - 1
'    xd = edessdo.tron_amo.Absamo + lg + edoor_courbe_max_bas.dx(i)
'    yd = edessdo.tron_amo.radava - ((edoor_courbe_max_bas.dx(i) * edessdo.tron_amo.conduit.pente) + edoor_courbe_max_bas.dy(i))
'    xf = edessdo.tron_amo.Absamo + lg + edoor_courbe_max_bas.dx(i + 1)
'    yf = edessdo.tron_amo.radava - ((edoor_courbe_max_bas.dx(i + 1) * edessdo.tron_amo.conduit.pente) + edoor_courbe_max_bas.dy(i + 1))
    xd = edessdo.tron_amo.Absamo + lg + edoor_courbe_max_bas.dx(i)
    yd = edessdo.tron_amo.radava - ((edoor_courbe_max_bas.dx(i) * 0) + edoor_courbe_max_bas.dy(i))
    xf = edessdo.tron_amo.Absamo + lg + edoor_courbe_max_bas.dx(i + 1)
    yf = edessdo.tron_amo.radava - ((edoor_courbe_max_bas.dx(i + 1) * 0) + edoor_courbe_max_bas.dy(i + 1))
    uc_g.dess_lign xd, yd, xf, yf, couleur.bleu, 1
Next
'dessin ligne eau dans conduite deversement
    xd = lg + edo.Longueur
    yd = edessdo.tron_dech.radamo + edoor_res.hdev
    xf = xd + lg
    yf = edessdo.tron_dech.radamo - edessdo.tron_dech.conduit.pente * lg + edoor_res.hdev
    uc_g.dess_lign_point xfd, yfd, xd, yd, couleur.bleu
'    uc_g.dess_lign xfd, yfd, xd, yd, couleur.bleu, 1
    uc_g.dess_lign xd, yd, xf, yf, couleur.bleu, 1

End Sub
Public Sub dess_debit_cri_or(ByRef uc_g As UC_graphique)
Dim xd As Double, yd As Double, xf As Double, yf As Double, dc As Double
Dim lg As Double
Dim i As Integer, j As Integer, np As Integer
lg = 2 * edo.Longueur
'dessin ligne eau dans conduite amont(arrivée)
xd = edessdo.tron_amo.Absamo
yd = edessdo.tron_amo.radava + ((edessdo.tron_amo.radamo - edessdo.tron_amo.radava) / edessdo.tron_amo.conduit.Longueur * lg)
xf = xd + lg
yf = edessdo.tron_amo.radava
dc = edessdo.tron_amo.conduit.Diametre
uc_g.dess_lign xd, yd + edoor_res.Ham_cri, xf, yf + edoor_res.hc_cri, couleur.rouge, 1
np = UBound(edoor_courbe_max_haut.dx)
'dessin courbe cri haut
j = 0
For i = 1 To np - 1
'    xd = edessdo.tron_amo.Absamo + lg + edoor_courbe_cri_haut.dx(i)
'    yd = edessdo.tron_amo.radava - ((edoor_courbe_cri_haut.dx(i) * edessdo.tron_amo.conduit.pente) + edoor_courbe_cri_haut.dy(i))
'    xf = edessdo.tron_amo.Absamo + lg + edoor_courbe_cri_haut.dx(i + 1)
'    yf = edessdo.tron_amo.radava - ((edoor_courbe_cri_haut.dx(i + 1) * edessdo.tron_amo.conduit.pente) + edoor_courbe_cri_haut.dy(i + 1))
    xd = edessdo.tron_amo.Absamo + lg + edoor_courbe_cri_haut.dx(i)
    yd = edessdo.tron_amo.radava - ((edoor_courbe_cri_haut.dx(i) * 0) + edoor_courbe_cri_haut.dy(i))
    xf = edessdo.tron_amo.Absamo + lg + edoor_courbe_cri_haut.dx(i + 1)
    yf = edessdo.tron_amo.radava - ((edoor_courbe_cri_haut.dx(i + 1) * 0) + edoor_courbe_cri_haut.dy(i + 1))
    uc_g.dess_lign xd, yd, xf, yf, couleur.rouge, 1
Next
'dessin ligne eau dans conduite aval(départ)
xd = xf 'edessdo.tron_amo.Absamo + lg + edoor_res.l_jetaval_h
yd = yf 'edessdo.tron_ava.radamo - ((edoor_res.l_jetaval_h - edoor_res.l_chambre1) * edessdo.tron_ava.conduit.pente) + edoor_res.hav_cri
xf = edessdo.tron_amo.Absamo + lg + edoor_res.l_jetaval_b + lg
yf = edessdo.tron_ava.radamo - (((edoor_res.l_jetaval_b - edoor_res.l_chambre1) + lg) * edessdo.tron_ava.conduit.pente) + edoor_res.hav_cri
uc_g.dess_lign xd, yd, xf, yf, couleur.rouge, 1
'dessin courbe cri bas
For i = 1 To np - 1
'    xd = edessdo.tron_amo.Absamo + lg + edoor_courbe_cri_bas.dx(i)
'    yd = edessdo.tron_amo.radava - ((edoor_courbe_cri_bas.dx(i) * edessdo.tron_amo.conduit.pente) + edoor_courbe_cri_bas.dy(i))
'    xf = edessdo.tron_amo.Absamo + lg + edoor_courbe_cri_bas.dx(i + 1)
'    yf = edessdo.tron_amo.radava - ((edoor_courbe_cri_bas.dx(i + 1) * edessdo.tron_amo.conduit.pente) + edoor_courbe_cri_bas.dy(i + 1))
    xd = edessdo.tron_amo.Absamo + lg + edoor_courbe_cri_bas.dx(i)
    yd = edessdo.tron_amo.radava - ((edoor_courbe_cri_bas.dx(i) * 0) + edoor_courbe_cri_bas.dy(i))
    xf = edessdo.tron_amo.Absamo + lg + edoor_courbe_cri_bas.dx(i + 1)
    yf = edessdo.tron_amo.radava - ((edoor_courbe_cri_bas.dx(i + 1) * 0) + edoor_courbe_cri_bas.dy(i + 1))
    uc_g.dess_lign xd, yd, xf, yf, couleur.rouge, 1
Next

End Sub

Public Sub dess_troncon_c(ByRef uc_g As UC_graphique, ByRef tr As troncon, ByRef ocolor As ColorConstants)
uc_g.redef_drwidth 2
uc_g.dess_lign tr.Absamo, tr.radamo, tr.Absava, tr.radava, ocolor, 2
uc_g.dess_lign tr.Absamo, tr.radamo + tr.conduit.Diametre, tr.Absava, tr.radava + tr.conduit.Diametre, ocolor, 2
uc_g.dess_circle tr.Absava, tr.radava + tr.conduit.Diametre / 2#, tr.conduit.Diametre / 2#, 0, 0, ocolor
uc_g.dess_circle tr.Absamo, tr.radamo + tr.conduit.Diametre / 2#, tr.conduit.Diametre / 2#, 1.57, 4.71, ocolor
uc_g.redef_drwidth 1
uc_g.redef_drstyle 2
uc_g.dess_circle tr.Absamo, tr.radamo + tr.conduit.Diametre / 2#, tr.conduit.Diametre / 2#, 4.71, 1.57, ocolor
uc_g.redef_drstyle 0
'uc_g.dess_lign tr.Absamo, tr.radamo, tr.Absamo, tr.radamo + tr.conduit.Diametre, ocolor, 2
'uc_g.dess_lign tr.Absava, tr.radava, tr.Absava, tr.radava + tr.conduit.Diametre, ocolor, 2
End Sub
Public Sub dess_cot(ByRef uc_g As UC_graphique, ByRef ocolor As ColorConstants)
uc_g.dess_cote edessdo.tron_amo.Absamo, edessdo.tron_amo.Absava, edessdo.tron_amo.conduit.Longueur, ocolor
uc_g.dess_cote edo.Absamo, edo.Absava, edo.Longueur, ocolor
uc_g.dess_cote edo.tron_ava.Absamo, edo.tron_ava.Absava, edo.tron_ava.conduit.Longueur, ocolor
If edo.tron_ava.Absava < edessdo.lgdisp Then
    uc_g.dess_cote edo.tron_ava.Absava, edessdo.lgdisp, edessdo.lgdisp - edo.tron_ava.Absava, ocolor
End If
End Sub
Public Sub dess_predo(ByRef uc_g As UC_graphique, ByRef tr As deversoir, ByRef ocolor As ColorConstants)


uc_g.dess_lign tr.Absamo, tr.radamo, tr.Absava, tr.radava, ocolor, 2
uc_g.dess_lign tr.Absamo, tr.radamo + tr.hauteur, tr.Absava, tr.radamo + tr.hauteur, ocolor, 2
uc_g.dess_lign tr.Absamo, tr.radamo, tr.Absamo, tr.radamo + tr.hauteur, ocolor, 2
uc_g.dess_lign tr.Absava, tr.radava, tr.Absava, tr.radamo + tr.hauteur, ocolor, 2
'uc_g.dess_circle tr.Absamo, tr.radamo + tr.conduit.Diametre / 2#, tr.conduit.Diametre / 2#, vbRed
'uc_g.dess_circle tr.Absava, tr.radava + tr.conduit.Diametre / 2#, tr.conduit.Diametre / 2#, vbRed
Call dess_troncon(uc_g, tr.tron_ava, couleur.noir) ' vbBlack)

End Sub
Public Sub dess_do(ByRef uc_g As UC_graphique, ByRef tr As deversoir, ByRef ocolor As ColorConstants)


uc_g.dess_lign tr.Absamo, tr.radamo, tr.Absava, tr.radava, ocolor, 2
uc_g.dess_lign tr.Absamo, tr.radamo + tr.hauteur, tr.Absava, tr.radamo + tr.hauteur, ocolor, 2
uc_g.dess_lign tr.Absamo, tr.radamo, tr.Absamo, tr.radamo + tr.hauteur, ocolor, 2
uc_g.dess_lign tr.Absava, tr.radava, tr.Absava, tr.radamo + tr.hauteur, ocolor, 2
'uc_g.dess_circle tr.Absamo, tr.radamo + tr.conduit.Diametre / 2#, tr.conduit.Diametre / 2#, vbRed
'uc_g.dess_circle tr.Absava, tr.radava + tr.conduit.Diametre / 2#, tr.conduit.Diametre / 2#, vbRed
'Call dess_troncon(uc_g, tr.tron_ava)
End Sub
Public Sub dess_door(ByRef uc_g As UC_graphique, ByRef tr As deversoir, ByRef ocolor As ColorConstants)
Dim xd As Double, xf As Double, lg As Double, dec As Double, dec1 As Double
Dim i As Integer, np As Integer, j As Integer
Dim txt As String
Dim xd1 As Double, xf1 As Double
lg = 2 * edo.Longueur
dec = edo.Longueur / 2
dec1 = edo.Longueur / 10
xd = edessdo.tron_amo.Absamo + lg
xf = xd + edo.Longueur
uc_g.dess_lign xd, edo.radamo, xf, edo.radava, ocolor, 3  'ligne fond
uc_g.dess_lign xd, edo.radamo, xd, edessdo.tron_amo.radava, ocolor, 3  'ligne verticale devant
yf = edo.radava + edessdo.tron_ava.conduit.Diametre
uc_g.dess_lign xf, edo.radava, xf, yf, ocolor, 1  ' ligne verticale  conduite
yd = yf
yf = yf + edo.tav
uc_g.dess_lign xf, yd, xf, yf, ocolor, 3
xd = xd + edoor_res.l_ouverture
yd = yf + ((xf - xd) * edo.pente) 'ligne epres ouverture
uc_g.dess_lign xd, yd, xf, yf, ocolor, 3
' cote hauteur do
xf1 = edessdo.tron_amo.Absamo + lg
xd = (edessdo.tron_amo.Absamo + lg) - dec
yd = edo.radava + edo.hauteur
xf = xf1 + edo.Longueur
yf = edo.radava
uc_g.dess_lign_point xd, yd, xf, yd, couleur.bleu
uc_g.dess_lign_point xd, yf, xf1, yf, couleur.bleu
uc_g.dess_lign xd + dec1, yd, xd + dec1, yf, couleur.bleu, 1
Text = Format(edo.hauteur, "##0.000") + " m"
uc_g.dess_text_aligne xd, "D", "C", (yf + yd) / 2, Text, couleur.bleu
' cote l_ouverture
xd = edessdo.tron_amo.Absamo + lg
yd = edo.radava
yf1 = edo.radava + edo.hauteur
xf = xd + edoor_res.l_ouverture
yf = edo.radava - 0.15
uc_g.dess_lign_point xd, yd, xd, yf, couleur.bleu
uc_g.dess_lign_point xf, yf1, xf, yf, couleur.bleu
uc_g.dess_lign xd, yf + 0.05, xf, yf + 0.05, couleur.bleu, 1
Text = Format(edoor_res.l_ouverture, "##0.000") + " m"
uc_g.dess_text_aligne xd, "D", "C", yf, Text, couleur.bleu
' cote longueur
xd = edessdo.tron_amo.Absamo + lg
yd = edo.radava
'yf1 = edo.radava + edo.hauteur
xf = xd + edo.Longueur
yf = edo.radava - 0.3
uc_g.dess_lign_point xd, yd, xd, yf, couleur.bleu
uc_g.dess_lign_point xf, yd, xf, yf, couleur.bleu
uc_g.dess_lign xd, yf + 0.05, xf, yf + 0.05, couleur.bleu, 1
Text = Format(edo.Longueur, "##0.000") + " m"
'uc_g.dess_text_aligne (xd + xf) / 2, "D", "H", yf, Text, couleur.bleu
uc_g.dess_text xd, xf, " ", yf, Text, couleur.bleu
' cote radier aval deversoir
xd = edessdo.tron_amo.Absamo + lg + edo.Longueur
yd = edo.radava
xf = xd + dec1 * 2
yf = yd - dec1 '* 2
uc_g.dess_lign xd, yd, xf, yf, couleur.bleu, 1
uc_g.dess_text_aligne xf, "F", "C", yf, Format(edo.radava, "####0.000") + " m", couleur.bleu
' cote radier aval conduite amont
xd = edessdo.tron_amo.Absamo + lg
yd = edessdo.tron_amo.radava
xf = xd - dec1 * 6 ' 2
yf = yd - dec1 / 2 '* 2
uc_g.dess_lign xd, yd, xf, yf, couleur.bleu, 1
uc_g.dess_text_aligne xf, "D", "C", yf, Format(edessdo.tron_amo.radava, "####0.000") + " m", couleur.bleu

End Sub
Public Sub init_graph_courbe(ByRef uc_g As UC_graphique, ByRef eb As courbe_dess)
Dim qftot As Double, dtot As Double, htot As Double
dtot = (Int((eb.duree * 1.5) / 10) + 1) * 10
qftot = (eb.quantite * 60 * dtot) / 1000#
htot = (Int(qftot + eb.volume * 1.5) / 100) * 100
htot = Int(eb.hauteur * 1.5 / 100) * 100
If htot = 0 Then htot = 100
uc_g.graphique_clear
    uc_g.reinit 7, "Arial"
    uc_g.init_title
    uc_g.init_titleh ""
    uc_g.init_titleb ""
uc_g.redef_cadrs 600, 500, 200
uc_g.init_arrondi_y 1
uc_g.init_MinX 0
uc_g.init_MinY 0
uc_g.init_MaxX dtot
uc_g.init_MaxY htot
uc_g.init_EchYn 0.9
uc_g.init_EchXn 1#
'   uc_g.dess_cadre 10, 2, 100, 2, 2, 1, 10, 2, 100
uc_g.dess_cadre 10, 2, 50, 0, 0, 0, 10, 2, 10
uc_g.dess_lign 0, 0, dtot, qftot, couleur.rouge, 1



End Sub
Public Sub init_graph_circ(ByRef uc_graph As UC_graphique, ByRef eb As volume_dess)
Dim ecx As Double, diam As Double, haut As Double, nb1 As Double, nb2 As Double, nb As Double
Dim i As Integer
diam = eb.Diametre
haut = eb.hauteur + eb.coef
nb1 = diam
nb2 = haut + (haut * 0.25)
nb = maximum(nb1, nb2)
uc_graph.graphique_clear
uc_graph.reinit 7, "Arial"
    uc_graph.init_title
    uc_graph.init_titleh ""
    uc_graph.init_titleb ""
uc_graph.init_arrondi_X 2
uc_graph.init_arrondi_y 3
uc_graph.init_MinX -nb / 4
uc_graph.init_MaxX nb + nb / 8
uc_graph.init_EchXn 1
'ecx = owner.fdessin.UC_graphique1.lire_EchXn()
uc_graph.init_MaxY nb + nb / 8
uc_graph.init_MinY -nb / 2
uc_graph.init_EchYn 1
   
End Sub
Public Sub init_graph_cond(ByRef uc_graph As UC_graphique, ByRef eb As volume_dess)
Dim ecx As Double, diam As Double, xlong As Double, nb1 As Double, nb2 As Double, nb As Double
Dim i As Integer
diam = eb.Diametrec
xlong = eb.Longueurc * 1.4
nb2 = diam * 2
nb1 = -diam * 0.5
uc_graph.graphique_clear
uc_graph.reinit 7, "Arial"
    uc_graph.init_title
    uc_graph.init_titleh ""
    uc_graph.init_titleb ""
uc_graph.init_arrondi_X 2
uc_graph.init_arrondi_y 3
uc_graph.init_MinX 0 '-nb / 4
uc_graph.init_MaxX xlong 'nb + nb / 8
uc_graph.init_EchXn 1
'ecx = owner.fdessin.UC_graphique1.lire_EchXn()
uc_graph.init_MaxY nb2 ' nb + nb / 8
uc_graph.init_MinY nb1 ' -nb / 4
uc_graph.init_EchYn 1
   
End Sub
Public Sub init_graph_rect(ByRef uc_graph As UC_graphique, ByRef eb As volume_dess)
Dim xlon As Double, xlar As Double, haut As Double
Dim ecx As Double
Dim i As Integer
Dim decalx As Double
    xlon = eb.Longueur
    xlar = eb.Largeur
    haut = eb.Profondeur
uc_graph.graphique_clear
uc_graph.reinit 7, "Arial"
    uc_graph.init_title
    uc_graph.init_titleh ""
    uc_graph.init_titleb ""
uc_graph.init_arrondi_X 2
uc_graph.init_arrondi_y 3
decalx = xlon / 5#
If decalx < 3 Then
    decalx = 3
End If
uc_graph.init_MinX -decalx '4#
uc_graph.init_MaxX xlon + 1.5 * xlar
uc_graph.init_EchXn 1
ecx = uc_graph.lire_EchXn()
uc_graph.init_MaxY haut + 1
uc_graph.init_MinY -0.5
uc_graph.init_EchYn 1
   
End Sub
Public Sub dess_stock_cond(ByRef uc_g As UC_graphique, ByRef eb As volume_dess)
Dim xam As Double, yam As Double, xav As Double, yav As Double
Dim diam As Double, xlong As Double
Dim tr As troncon
diam = eb.Diametrec
xlong = eb.Longueurc
tr.Absamo = xlong * 0.2 '15
tr.radamo = 0
tr.Absava = xlong + tr.Absamo
tr.radava = 0
tr.conduit.Diametre = diam
Call dess_troncon_c(uc_g, tr, couleur.noir)
uc_g.dess_coth_text tr.Absamo, tr.radamo, tr.Absava, tr.radava, ajout_zero(Trim(str(xlong))) + " m", couleur_noir
uc_g.dess_cotv_texte tr.Absamo, 0, tr.Absamo, diam, ajout_zero(Trim(str(diam))) + " m ", couleur_noir


End Sub
Public Sub dess_stock_circ(ByRef uc_g As UC_graphique, ByRef eb As volume_dess)
Dim xam As Double, yam As Double, xav As Double, yav As Double
Dim diam As Double, haut As Double
diam = eb.Diametre
haut = eb.hauteur + eb.coef
uc_g.redef_drwidth 2
uc_g.dess_cercle_X diam / 2#, haut, diam / 2#, 0.25, 0, 0, couleur.noir
uc_g.dess_cercle_X diam / 2#, eb.hauteur, diam / 2#, 0.25, 3.14, 6.28, couleur.bleu
uc_g.dess_cercle_X diam / 2#, 0, diam / 2#, 0.25, 3.14, 6.28, couleur.noir
uc_g.dess_lign 0, 0, 0, haut, couleur.noir, 2
uc_g.dess_lign diam, 0, diam, haut, couleur.noir, 2
uc_g.redef_drwidth 1
uc_g.redef_drstyle 2
uc_g.dess_cercle_X diam / 2#, eb.hauteur, diam / 2#, 0.25, 0, 3.14, couleur.bleu
uc_g.dess_cercle_X diam / 2#, 0, diam / 2#, 0.25, 0, 3.14, couleur.noir
uc_g.redef_drstyle 0
'uc_g.dess_cot_cercle 0, haut, diam, haut, Trim(Str(diam)) + " m", couleur_noir
uc_g.dess_coth_text 0, 0, diam, 0, ajout_zero(Trim(str(diam))) + " m", couleur_noir
uc_g.dess_cotv_texte 0, 0, 0, eb.hauteur, ajout_zero(Trim(str(eb.hauteur))) + " m ", couleur_noir


End Sub

Public Sub dess_stock_rect(ByRef uc_g As UC_graphique, ByRef eb As volume_dess)
Dim xam As Double, yam As Double, xav As Double, yav As Double
Dim xlon As Double, xlar As Double, haut As Double, htot As Double
xlon = eb.Longueur
xlar = eb.Largeur
haut = eb.Profondeur
htot = eb.Profondeur + eb.coef
uc_g.redef_drwidth 2
xam = 0
yam = 0
xav = xlon
yav = 0
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = 0
yam = 0
xav = 0
yav = htot
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = 0
yam = htot
xav = xlon
yav = htot
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xlon
yam = 0
xav = xlon
yav = htot
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xlon
yam = 0
xav = xlon + xlar '2 * xlar
yav = 0.3
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xlon
yam = htot
xav = xlon + xlar ' 2 * xlar
yav = htot + 0.3
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xlon + xlar ' 2 * xlar
yam = 0.3
xav = xam
yav = htot + 0.3
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = 0
yam = htot
xav = xlar '2 * xlar
yav = htot + 0.3
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
xam = xlar '2 * xlar
yam = htot + 0.3
xav = xlon + xlar '2 * xlar
yav = htot + 0.3
uc_g.dess_lign xam, yam, xav, yav, couleur.noir, 2
uc_g.redef_drwidth 1
xam = 0
yam = 0
xav = xlar '2 * xlar
yav = 0.3
uc_g.dess_tiret xam, yam, xav, yav, couleur.noir
xam = xlar '2 * xlar
yam = 0.3
xav = xam
yav = htot + 0.3
uc_g.dess_tiret xam, yam, xav, yav, couleur.noir
xam = xlar '2 * xlar
yam = 0.3
xav = xlon + xlar '2 * xlar
yav = 0.3
uc_g.dess_tiret xam, yam, xav, yav, couleur.noir
uc_g.redef_drwidth 1
xam = 0
yam = haut
xav = xlon
yav = haut
uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 1
xam = xlon
yam = haut
xav = xlon + xlar ' 2 * xlar
yav = haut + 0.3
uc_g.dess_lign xam, yam, xav, yav, couleur.bleu, 1
    xam = 0
    yam = haut
    xav = xlar '2 * xlar
    yav = haut + 0.3
    uc_g.dess_tiret xam, yam, xav, yav, couleur.bleu
    xam = xlar '2 * xlar
    yam = haut + 0.3
    xav = xlon + xlar '2 * xlar
    yav = haut + 0.3
    uc_g.dess_tiret xam, yam, xav, yav, couleur.bleu

uc_g.redef_drwidth 1
'uc_g.dess_coth 0, 0, xlon, 0, xlon, couleur_noir
'uc_g.dess_cotv 0, 0, 0, haut, haut, couleur_noir
'uc_g.dess_cotb 0, htot, 2 * xlar, htot + 0.3, xlar, couleur_noir
'uc_g.dess_cotb 0, htot, xlar, htot + 0.3, xlar, couleur_noir
uc_g.dess_coth_text 0, 0, xlon, 0, ajout_zero(Trim(str(xlon))) + " m", couleur_noir
uc_g.dess_cotv_texte 0, 0, 0, haut, ajout_zero(Trim(str(haut))) + " m ", couleur_noir
uc_g.dess_cotb_text 0, htot, xlar, htot + 0.3, xlar, ajout_zero(Trim(str(xlar))) + " m  ", couleur_noir
End Sub

