Attribute VB_Name = "leaping_wear"
Option Explicit
Public Function calcul_long_ouverture(hcri As Double, h0 As Double, nbFr As Double, pent As Double, coefEcoul As Double) As Double
Dim hc22 As Double, dh As Double, dc As Double
Dim hcZ As Double, dcX As Double, ept As Double
Dim dvx As Double, hc_cri As Double
Dim delt As Double
Dim pente As Double
pente = 0
hc_cri = hcri
delt = hc_cri / h0
    hcZ = hc22 / h0
    dcX = 2 * (-1 / 3 + (1 / 9 + hcZ) ^ 0.5)
    dc = dcX * h0 * nbFr ^ 0.8
    dc = dc ' - (hc22 / edoor_res.deltaa)
    ept = 1 + 0.06 * dcX
     ept = ept * hc_cri
 Dim delta As Double
 delta = (1 / 3# - 0.06 * delt) ^ 2 + 1 * delt
 dcX = 2 * (-(1 / 3 - 0.06 * delt) + delta ^ 0.5)
     dc = dcX * h0 * nbFr ^ 0.8

 delta = (1 / 3# - 0.06 * delt - pente / nbFr ^ 0.8) ^ 2 + 1 * delt
dcX = 2 * (-(1 / 3 - 0.06 * delt - pente / nbFr ^ 0.8) + delta ^ 0.5)
    dc = dcX * h0 * nbFr ^ 0.8
 'calcul
 Dim a As Double, b As Double, c As Double
 Dim coef As Double
 coef = 1 / (h0 * nbFr ^ 0.8)
 a = 0.25 * coef ^ 2
 b = (1 / 3 - 0.06 * delt) * coef - pente / h0
 c = -delt
 
  delta = b ^ 2 - 4 * a * c
 ' rDelta = delta ^ 0.5
  dcX = (-b + delta ^ 0.5) / (2 * a)
   dc = dcX

' formule nlle
dcX = ((1) / (coefEcoul * coef ^ 1.5)) ^ (1 / 1.5)
dcX = (1 / (coefEcoul ^ (1 / 1.5))) / coef
  dc = dcX
 Dim d As Double
 Dim dx As Double
 Dim D0 As Double
 dx = 0.005
 d = 1
 D0 = 1
dc = dc - dx
While Abs(d) > 0.0001
dc = dc + dx
hcZ = -dc * pente
dh = (1 - coefEcoul * coef ^ 1.5 * dc ^ 1.5) * hcri
d = dh - hcZ
If d * D0 < 0 Then
dx = -dx / 2
End If
D0 = d
'dc = dc + dx

Wend
    calcul_long_ouverture = dc
End Function
Public Function calcul_X_pour_h(hcri As Double, h0 As Double, nbFr As Double, pent As Double, ByVal hy As Double, coefEcoul As Double) As Double
Dim hc22 As Double, dh As Double, dc As Double
Dim hcZ As Double, dcX As Double, ept As Double
Dim dvx As Double, hc_cri As Double
Dim delt As Double
Dim pente As Double
pente = pent
pente = 0

hc_cri = hcri
delt = hc_cri / h0
    hcZ = hc22 / h0
 'calcul
 Dim a As Double, b As Double, c As Double
 Dim coef As Double
 coef = 1 / (h0 * nbFr ^ 0.8)

' formule nlle
dcX = ((1) / (coefEcoul * coef ^ 1.5)) ^ (1 / 1.5)
dcX = ((1 - (hy / hcri)) / (coefEcoul * coef ^ 1.5)) ^ (1 / 1.5)

  dc = dcX
 Dim d As Double
 Dim dx As Double
 Dim D0 As Double
 dx = 0.01
 d = 1
 d = 0
 D0 = 1
While Abs(d) > 0.001
hcZ = -dc * pente
dh = (1 - coefEcoul * coef ^ 1.5 * dc ^ 1.5) * hcri
d = dh - hcZ
If d * D0 < 0 Then
dx = -dx / 2
End If
D0 = d
dc = dc + dx

Wend
    calcul_X_pour_h = dc
End Function
Public Function calcul_Y_pour_X(hcri As Double, h0 As Double, nbFr As Double, pent As Double, ByVal xl As Double, coefEcoul As Double) As Double
Dim hc22 As Double, dh As Double, dc As Double
Dim hcZ As Double, dcX As Double, ept As Double
Dim dvx As Double, hc_cri As Double
Dim delt As Double
Dim pente As Double
pente = 0

hc_cri = hcri
delt = hc_cri / h0
 'calcul
 Dim coef As Double
 coef = 1 / (h0 * nbFr ^ 0.8)

' formule nlle
dcX = xl
hcZ = (1 - coefEcoul * (dcX * coef) ^ 1.5) * hc_cri
    calcul_Y_pour_X = hcZ
End Function
Public Function calcul_longueur_bas_jet(hcri As Double, h0 As Double, nbFr As Double, pent As Double, deltaH As Double, coefEcoul As Double) As Double
Dim hc22 As Double, dh As Double, dc As Double
Dim hcZ As Double, dcX As Double, ept As Double
Dim dvx As Double, hc_cri As Double
Dim delt As Double
Dim pente As Double
pente = pent
pente = 0
hc_cri = hcri
 Dim coef As Double
 coef = 1 / (h0 * nbFr ^ 0.8)
' formule nlle
Dim hy As Double
hy = -deltaH + hc_cri
dcX = ((1) / (coefEcoul * coef ^ 1.5)) ^ (1 / 1.5)
dcX = ((1 - (hy / hcri)) / (coefEcoul * coef ^ 1.5)) ^ (1 / 1.5)
    ept = 1 + 0.06 * dcX * coef
    ept = ept * hc_cri

  dc = dcX
  hcZ = hy - dc * pente

 Dim d As Double
 Dim dx As Double
 Dim D0 As Double
 d = 1
 D0 = 1
dh = hy
 dx = 0.02
   If (dh - ept) < (hcZ - hc_cri) Then
     d = -1

    dx = -0.02
End If
D0 = d
 'd = 0
While Abs(d) > 0.0005
hcZ = hy - dc * pente
    ept = 1 + 0.06 * dc * coef
    ept = ept * hc_cri

dh = (1 - coefEcoul * coef ^ 1.5 * dc ^ 1.5) * hcri
dh = dh
d = dh - hcZ
d = (dh - ept) - (hcZ - hc_cri)
If d * D0 < 0 Then
dx = -dx / 2
End If
D0 = d
dc = dc + dx

Wend

    calcul_longueur_bas_jet = dc
End Function
Public Function calcul_hauteur_jet_chambre(hcri As Double, h0 As Double, nbFr As Double, pent As Double, xlong As Double, coefEcoul As Double) As Double
Dim hc22 As Double, dh As Double, dc As Double
Dim hcZ As Double, dcX As Double, ept As Double
Dim dvx As Double, hc_cri As Double
Dim delt As Double
Dim pente As Double
pente = pent
pente = 0

hc_cri = hcri
 Dim coef As Double
 coef = 1 / (h0 * nbFr ^ 0.8)
' formule nlle
Dim hy As Double
dc = xlong

hy = (1 - coefEcoul * (coef * dc) ^ 1.5) * hc_cri

    ept = 1 + 0.06 * dcX * coef
    ept = ept * hc_cri

  hcZ = hy - dc * pente - ept
    calcul_hauteur_jet_chambre = -hcZ
End Function


Public Function calcul_hauteur_bas_jet(hcri As Double, h0 As Double, nbFr As Double, pent As Double, lc As Double, deltaH As Double, penteA As Double, coefEcoul As Double) As Double
'Public Function calcul_hauteur_bas_jet(hcri As Double, h0 As Double, nbFr As Double, pente As Double, deltaH As Double) As Double
Dim hc22 As Double, dh As Double, dc As Double
Dim hcZ As Double, dcX As Double, ept As Double
Dim dvx As Double, hc_cri As Double
Dim delt As Double
Dim pente As Double
pente = pent
pente = 0
hc_cri = hcri
 Dim coef As Double
 coef = 1 / (h0 * nbFr ^ 0.8)
' formule nlle
Dim hy As Double
hy = -deltaH + hc_cri
dcX = ((1) / (coefEcoul * coef ^ 1.5)) ^ (1 / 1.5)
dcX = ((1 - (hy / hcri)) / (coefEcoul * coef ^ 1.5)) ^ (1 / 1.5)
    ept = 1 + 0.06 * dcX * coef
    ept = ept * hc_cri

  dc = dcX
hy = -deltaH '+ hc_cri
  
'  hcZ = hy - dc * pente
If dc < lc Then
    hcZ = hy - dc * pente
Else
    hcZ = hy - (lc * pente + (dc - lc) * penteA)
End If

 Dim d As Double
 Dim dx As Double
 Dim D0 As Double
 d = 1
 D0 = 1
dh = hy + hc_cri
 dx = 0.02
'   If (dh - ept) < (hcZ - hc_cri) Then
   If (dh - ept) < (hcZ) Then
     d = -1
        dx = -0.02
   End If
D0 = d
 'd = 0
dc = dc - dx
While Abs(d) > 0.001
    dc = dc + dx
'hcZ = hy - dc * pente
    If dc < lc Then
        hcZ = hy - dc * pente
    Else
        hcZ = hy - (lc * pente + (dc - lc) * penteA)
    End If
    ept = 1 + 0.06 * dc * coef
    ept = ept * hc_cri

dh = (1 - coefEcoul * coef ^ 1.5 * dc ^ 1.5) * hcri
dh = dh
d = dh - hcZ
d = (dh - ept) - (hcZ) '- hc_cri)
If d * D0 < 0 Then
    dx = -dx / 2
End If
D0 = d
'dc = dc + dx

Wend

    calcul_hauteur_bas_jet = (-hcZ + hc_cri - ept)
End Function

Public Function calcul_hauteur_bas_jet0(hcri As Double, h0 As Double, nbFr As Double, pente As Double, deltaH As Double) As Double
Dim hc22 As Double, dh As Double, dc As Double
Dim hcZ As Double, dcX As Double, ept As Double
Dim dvx As Double, hc_cri As Double
Dim delt As Double

hc_cri = hcri
delt = hc_cri / h0
'    hcZ = hc22 / edoor_res.Ham_cri
'    dcX = 2 * (-1 / 3 + (1 / 9 + hcZ) ^ 0.5)
    dc = dcX * h0 * nbFr ^ 0.8
    dc = dc ' - (hc22 / edoor_res.deltaa)
    ept = 1 + 0.06 * dcX
     ept = ept * hc_cri
 Dim delta As Double
 delta = (1 / 3# - 0.06 * delt) ^ 2 + 1 * delt
 dcX = 2 * (-(1 / 3 - 0.06 * delt) + delta ^ 0.5)
     dc = dcX * h0 * nbFr ^ 0.8

 delta = (1 / 3# - 0.06 * delt - pente / nbFr ^ 0.8) ^ 2 + 1 * delt
dcX = 2 * (-(1 / 3 - 0.06 * delt - pente / nbFr ^ 0.8) + delta ^ 0.5)
    dc = dcX * h0 * nbFr ^ 0.8
 'calcul
 Dim a As Double, b As Double, c As Double
 Dim coef As Double
 coef = 1 / (h0 * nbFr ^ 0.8)
 a = 0.25 * coef ^ 2
 b = (1 / 3) * coef - pente / h0
 c = -deltaH / h0
 
  delta = b ^ 2 - 4 * a * c
 ' rDelta = delta ^ 0.5
  dcX = (-b + delta ^ 0.5) / (2 * a)
   dc = dcX
 dcX = dcX * coef
hcZ = 1# / 3# * dcX + 1# / 4# * dcX ^ 2
hcZ = hcZ * h0
    
' formule nlle
Dim hy As Double
hy = -deltaH + hc_cri
dcX = ((1) / (0.54 * coef ^ 1.5)) ^ (1 / 1.5)
dcX = ((1 - (hy / hcri)) / (0.54 * coef ^ 1.5)) ^ (1 / 1.5)
    ept = 1 + 0.06 * dcX * coef
    ept = ept * hc_cri

  dc = dcX
  hcZ = hy - dc * pente

 Dim d As Double
 Dim dx As Double
 Dim D0 As Double
 d = 1
 D0 = 1

 dx = 0.02
   If (dh - ept) < (hcZ - hc_cri) Then
     d = -1

    dx = -0.02
End If
D0 = d
 'd = 0
While Abs(d) > 0.001
hcZ = hy - dc * pente
    ept = 1 + 0.06 * dc * coef
    ept = ept * hc_cri

dh = (1 - 0.54 * coef ^ 1.5 * dc ^ 1.5) * hcri
dh = dh
d = dh - hcZ
d = (dh - ept) - (hcZ - hc_cri)
If d * D0 < 0 Then
dx = -dx / 2
End If
D0 = d
dc = dc + dx

Wend

    calcul_hauteur_bas_jet0 = -hcZ + hc_cri
End Function
Public Function calcul_hauteur_haut_jet(hcri As Double, h0 As Double, nbFr As Double, pent As Double, lc As Double, deltaH As Double, penteA As Double, coefEcoul) As Double
'Public Function calcul_hauteur_haut_jet(hcri As Double, h0 As Double, nbFr As Double, pente As Double, deltaH As Double) As Double
Dim hc22 As Double, dh As Double, dc As Double
Dim hcZ As Double, dcX As Double, ept As Double
Dim dvx As Double, hc_cri As Double
Dim delt As Double
Dim pente As Double
pente = pent
pente = 0
hc_cri = hcri
delt = hc_cri / h0
'    hcZ = hc22 / edoor_res.Ham_cri
'    dcX = 2 * (-1 / 3 + (1 / 9 + hcZ) ^ 0.5)
    dc = dcX * h0 * nbFr ^ 0.8
    dc = dc ' - (hc22 / edoor_res.deltaa)
    ept = 1 + 0.06 * dcX
     ept = ept * hc_cri
 Dim delta As Double
 delta = (1 / 3# - 0.06 * delt) ^ 2 + 1 * delt
 dcX = 2 * (-(1 / 3 - 0.06 * delt) + delta ^ 0.5)
     dc = dcX * h0 * nbFr ^ 0.8

 delta = (1 / 3# - 0.06 * delt - pente / nbFr ^ 0.8) ^ 2 + 1 * delt
dcX = 2 * (-(1 / 3 - 0.06 * delt - pente / nbFr ^ 0.8) + delta ^ 0.5)
    dc = dcX * h0 * nbFr ^ 0.8
 'calcul
 Dim a As Double, b As Double, c As Double
 Dim coef As Double
 coef = 1 / (h0 * nbFr ^ 0.8)
 a = 0.25 * coef ^ 2
 b = (1 / 3) * coef - 0.06 * delt - pente / h0
 c = -delt - deltaH / h0
 
  delta = b ^ 2 - 4 * a * c
 ' rDelta = delta ^ 0.5
  dcX = (-b + delta ^ 0.5) / (2 * a)
   dc = dcX
 dcX = dcX * coef
hcZ = 1# / 3# * dcX + 1# / 4# * dcX ^ 2
hcZ = hcZ * h0
  
' formule nlle
Dim hy As Double
hy = -deltaH
dcX = ((1) / (coefEcoul * coef ^ 1.5)) ^ (1 / 1.5)
dcX = ((1 - (hy / hcri)) / (coefEcoul * coef ^ 1.5)) ^ (1 / 1.5)

  dc = dcX
 Dim d As Double
 Dim dx As Double
 Dim D0 As Double
 dx = 0.01
 d = 1
 'd = 0
 D0 = 1
 dh = hy
 dx = 0.02
'   If (dh - ept) < (hcZ - hc_cri) Then
If dc < lc Then
    hcZ = hy - (lc * pente + (dc - lc) * penteA)
Else
    hcZ = hy - (lc * pente + (dc - lc) * penteA)
End If

  If dh < hcZ Then
    d = -1
    dx = -0.02
End If
D0 = d
dc = dc - dx
While Abs(d) > 0.001
dc = dc + dx
If dc < lc Then
    hcZ = hy - dc * pente
    hcZ = hy - (lc * pente + (dc - lc) * penteA)
Else
    hcZ = hy - (lc * pente + (dc - lc) * penteA)
End If
dh = (1 - coefEcoul * coef ^ 1.5 * dc ^ 1.5) * hcri
d = dh - hcZ
If d * D0 < 0 Then
dx = -dx / 2
End If
D0 = d
dc = dc + dx

Wend
 
 
    calcul_hauteur_haut_jet = (-hcZ + hc_cri)
End Function
Public Function verif_aval() As Boolean
Dim qv As deb_vit, qcri As Double
Dim troava As troncon
Dim cana_ava As conduite
    cana_ava.Diametre = edessdo.dav / 1000#
    cana_ava.Longueur = edessdo.Lav
    cana_ava.pente = edessdo.iradav / 10000#
    cana_ava.rugosite = edessdo.kav
    cana_ava.typ = 2
    With troava
      .Absamo = edessdo.tron_amo.Absava '+ edo.Longueur
      .Absava = .Absamo + cana_ava.Longueur
      .conduit = cana_ava
      .radamo = edessdo.tron_amo.radava '- edo.Longueur * edo.pente
      .radava = .radamo - cana_ava.Longueur * cana_ava.pente
    End With
    edessdo.tron_ava = troava
verif_aval = True
qv = debvit_ps(edessdo.tron_ava.conduit)
qcri = edoor_res.Qbaveff

If qv.debit < qcri Then
'If qv.debit * 1000 < qcri Then
  verif_aval = False
  MsgBox "la conduite aval est  en charge à Qbav!", vbExclamation, "Verif Aval"
End If

End Function
Public Function calcul_mini(ByRef mes As String, ByVal hmin As Double) As Boolean
Dim l_chambre1 As Double, l_jetaval_b As Double, l_jetaval_h As Double
Dim hc22 As Double, dh As Double, dc As Double
Dim i As Integer, np As Integer
Dim g As Double, l_ouverture As Double
Dim hc As Double, hc1 As Double, hc2 As Double
Dim alpha As Double, CosA As Double, deltaa As Double, vc_cri As Double, vc As Double
Dim Hbav As Double, Hav As Double, hc_cri As Double, hav_cri As Double

Dim hcZ As Double, dcX As Double, ept As Double
Dim dvx As Double
Dim coefEcoul As Double
coefEcoul = 0.54

g = 9.81
'    edoor_res.Qbaveff = Qbaveff
'    edoor_res.Qbavth = Qbavth
'    edoor_res.Ham = Ham
'    edoor_res.hdev = hdev
'    edoor_res.Ham_cri = Ham
Hbav = edoor_res.Hbav
hc = edoor_res.hc ' hauteur à l'overture à Qmax
vc = edoor_res.vc ' vitesse à l'overture à Qmax
l_ouverture = edoor_res.l_ouverture
Hav = edoor_res.Hav
alpha = edoor_res.alpha
CosA = edoor_res.CosA
deltaa = edoor_res.deltaa
hc_cri = edoor_res.hc_cri ' hauteur à l'overture à Qref
vc_cri = edoor_res.vc_cri ' vitesse à l'overture à Qref
hav_cri = edoor_res.hav_cri
hc1 = edessdo.tron_ava.conduit.Diametre - Hbav + hmin
If hc1 > 0 Then
 Dim nb1 As Double, nb2 As Double
 'calcul longueur de chambre
l_chambre1 = calcul_X_pour_h(hc_cri, edoor_res.Ham_cri, edoor_res.nbFroude, edoor_res.deltaa, -hc1, coefEcoul)
l_chambre1 = calcul_longueur_bas_jet(hc_cri, edoor_res.Ham_cri, edoor_res.nbFroude, edoor_res.deltaa, (hc1 + Hbav), coefEcoul)
edo.Longueur = l_chambre1
l_chambre1 = Round(l_chambre1, 2)
edo.Longueur = l_chambre1
edo.tav = hmin
'edoor_res.l_chambre1 = l_chambre1
     mes = mes + Chr(13) + Chr(10) + "Longueur de la chambre = " + ajout_zero(Trim(str(Round(l_chambre1, 3)))) + " m"
     mes = mes + Chr(13) + Chr(10) + "Hauteur de la chambre = " + ajout_zero(Trim(str(Round((hmin + edessdo.tron_ava.conduit.Diametre), 3)))) + " m"

Dim ok As Boolean
ok = calcul_courbes_jet(mes, hmin)
If ok Then
'   edoor_res.l_chambre1 = l_chambre1
'    edoor_res.l_jetaval_h = l_jetaval_h
'    edoor_res.l_jetaval_b = l_jetaval_b
'    edo.hauteur = hmin + edessdo.tron_ava.conduit.Diametre
'    edo.Absamo = edessdo.tron_amo.Absava
'    edo.Longueur = l_chambre1
'    edo.Absava = edo.Absamo + edo.Longueur
'
'    edo.radava = (edessdo.tron_amo.radava - edo.Longueur * edessdo.tron_amo.conduit.pente) - edessdo.tron_ava.conduit.Diametre - hmin
'    edo.pente = edessdo.tron_ava.conduit.pente
'    edo.radamo = edo.radava + edo.Longueur * edo.pente
'    edessdo.tron_ava.Absamo = edo.Absava
'    edessdo.tron_ava.radamo = edo.radava
'    edessdo.tron_ava.Absava = edessdo.tron_ava.Absamo + edessdo.tron_ava.conduit.Longueur
'    edessdo.tron_ava.radava = edessdo.tron_ava.radamo - edessdo.tron_ava.conduit.Longueur * edessdo.tron_ava.conduit.pente
End If
calcul_mini = ok
End If
End Function
Public Function calcul_coefSep()
Dim dx As Double, X As Double, coef As Double, hc As Double, hc22 As Double, ept As Double
Dim coefEcoul As Double, zDev As Double
coefEcoul = 0.4
 coef = 1 / (edoor_res.Ham * edoor_res.nbFroudeMax ^ 0.8)
hc = edoor_res.hc
'zDev = hc22
hc22 = calcul_Y_pour_X(hc, edoor_res.Ham, edoor_res.nbFroudeMax, edoor_res.deltaa, edoor_res.l_ouverture, coefEcoul)
ept = (1 + 0.06 * edoor_res.l_ouverture * coef) * hc
zDev = ept - zDev
zDev = ept - hc22
'zDev = ept - hc22 - (edoor_res.deltaa * edoor_res.l_ouverture)
Dim coefSep As Double
coefSep = zDev / ept
coefSep = (zDev - (edoor_res.deltaa * edoor_res.l_ouverture)) / ept
calcul_coefSep = coefSep
End Function

Public Function calcul_courbes_jet(ByRef mes As String, ByVal hmin As Double) As Boolean
Dim l_chambre1 As Double, l_jetaval_b As Double, l_jetaval_h As Double
Dim hc22 As Double, dh As Double, dc As Double
Dim i As Integer, np As Integer
Dim g As Double, l_ouverture As Double
Dim hc As Double, hc1 As Double, hc2 As Double
Dim alpha As Double, CosA As Double, deltaa As Double, vc_cri As Double, vc As Double
Dim Hbav As Double, Hav As Double, hc_cri As Double, hav_cri As Double
'Dim hmin As Double
Dim hcZ As Double, dcX As Double, ept As Double
Dim dvx As Double
Dim Ham As Double, Ham_cri As Double
 Dim coef As Double
 Dim coefEcoul As Double
 coefEcoul = 0.54

g = 9.81
'    edoor_res.Qbaveff = Qbaveff
'    edoor_res.Qbavth = Qbavth
'    edoor_res.Ham = Ham
Ham = edoor_res.Ham
'    edoor_res.hdev = hdev
Ham_cri = edoor_res.Ham_cri
Hbav = edoor_res.Hbav
hc = edoor_res.hc ' hauteur à l'overture à Qmax
vc = edoor_res.vc ' vitesse à l'overture à Qmax
l_ouverture = edoor_res.l_ouverture
Hav = edoor_res.Hav
alpha = edoor_res.alpha
CosA = edoor_res.CosA
deltaa = edoor_res.deltaa
hc_cri = edoor_res.hc_cri ' hauteur à l'overture à Qref
vc_cri = edoor_res.vc_cri ' vitesse à l'overture à Qref
hav_cri = edoor_res.hav_cri
hmin = edo.tav
l_chambre1 = edo.Longueur
Dim nbFr As Double
nbFr = edoor_res.nbFroude
hc1 = edessdo.tron_ava.conduit.Diametre - Hbav + hmin

'Calcul courbes
' debit critique
' nouveau calcul

' calcul courbe haut critique
Dim seuil_ini As Double, Seuil As Double
 coef = 1 / (Ham_cri * nbFr ^ 0.8)

np = UBound(edoor_courbe_cri_haut.dx)

hc2 = hmin + edessdo.tron_ava.conduit.Diametre
hc2 = calcul_hauteur_haut_jet(hc_cri, edoor_res.Ham_cri, edoor_res.nbFroude, edoor_res.deltaa, l_chambre1, (hc2 - hav_cri), edessdo.tron_ava.conduit.pente, coefEcoul)
dh = hc2 / (np - 1)
edoor_courbe_cri_haut.dx(1) = 0
edoor_courbe_cri_haut.dy(1) = -hc_cri
For i = 1 To np - 1 '+ 10
    hc22 = dh * i
    hc22 = hc_cri - dh * i
    hcZ = hc22 / edoor_res.Ham_cri
'    dcX = 2 * (-1 / 3 + (1 / 9 + hcZ) ^ 0.5)
'    dc = dcX * edoor_res.Ham_cri * edoor_res.nbFroude ^ 0.8
    dc = calcul_X_pour_h(hc_cri, edoor_res.Ham_cri, edoor_res.nbFroude, edoor_res.deltaa, hc22, coefEcoul)
    
    dc = dc ' - (hc22 / edoor_res.deltaa)
    ept = 1 + 0.06 * dc * coef
    ept = ept * hc_cri
'    ept = ept * hc_cri
    edoor_courbe_cri_haut.dx(i + 1) = dc
 '   edoor_courbe_cri_haut.dy(i + 1) = hc22  - epT '+ dc * edoor_res.deltaa
 '   edoor_courbe_cri_haut.dy(i + 1) = -hc22  ' - epT '+ dc * edoor_res.deltaa
    edoor_courbe_cri_haut.dy(i + 1) = -hc22 + Mini(dc, l_chambre1) * edoor_res.deltaa ' - epT '+ dc * edoor_res.deltaa
Next

'calcul    edoor_courbe_cri_bas

hc2 = hmin + edessdo.tron_ava.conduit.Diametre + (l_chambre1 * edessdo.tron_amo.conduit.pente)
hc2 = hmin + edessdo.tron_ava.conduit.Diametre ' + (l_chambre1 * edessdo.tron_amo.conduit.pente)
'hc2 = calcul_hauteur_bas_jet(hc_cri, edoor_res.Ham_cri, edoor_res.nbFroude, edoor_res.deltaa, hc2)
hc2 = calcul_hauteur_bas_jet(hc_cri, edoor_res.Ham_cri, edoor_res.nbFroude, edoor_res.deltaa, l_chambre1, hc2, edessdo.tron_ava.conduit.pente, coefEcoul)

dh = hc2 / (np - 1)
    edoor_courbe_cri_bas.dx(1) = 0#
    edoor_courbe_cri_bas.dy(1) = 0#
 '       edoor_courbe_cri_haut.dx(1) = 0
  '  edoor_courbe_cri_haut.dy(1) = -hc_cri

dvx = edoor_res.l_ouverture * edoor_res.deltaa
dcX = edoor_res.l_ouverture / (edoor_res.Ham_cri * edoor_res.nbFroude ^ 0.8)
hcZ = 1# / 3# * dcX + 1# / 4# * dcX ^ 2
hcZ = hcZ * edoor_res.Ham_cri
ept = 1 + 0.06 * dcX
ept = ept * hc_cri

seuil_ini = hmin + edessdo.tron_ava.conduit.Diametre
For i = 1 To np - 1 '+ 10
hc22 = dh * i
hc22 = hc_cri - dh * i
    hcZ = hc22 / edoor_res.Ham_cri
''    dcX = 2 * (-1 / 3 + (1 / 9 + hcZ) ^ 0.5)
'    dc = dcX * edoor_res.Ham_cri * edoor_res.nbFroude ^ 0.8
'    dc = dc ' - (hc22 / edoor_res.deltaa)
    dc = calcul_X_pour_h(hc_cri, edoor_res.Ham_cri, edoor_res.nbFroude, edoor_res.deltaa, hc22, coefEcoul)
    ept = 1 + 0.06 * dc * coef
    ept = ept * hc_cri
    Seuil = seuil_ini + dc * edoor_res.deltaa
    edoor_courbe_cri_bas.dx(i + 1) = dc
'    edoor_courbe_cri_bas.dy(i + 1) = -hc22 + ept '+ dc * edoor_res.deltaa
    edoor_courbe_cri_bas.dy(i + 1) = -hc22 + ept + Mini(dc, l_chambre1) * edoor_res.deltaa '+ dc * edoor_res.deltaa
Next


'Debit max

'calcul    edoor_courbe_max_dever
coefEcoul = 0.4
hc2 = hc + hmin + (edessdo.tron_ava.conduit.Diametre - Hav) + (l_chambre1 * edessdo.tron_amo.conduit.pente)
    hc2 = hc + hmin + (edessdo.tron_ava.conduit.Diametre - Hav) '+ (l_chambre1 * edessdo.tron_amo.conduit.pente)
    np = UBound(edoor_courbe_max_haut.dx)
    dh = hc2 / (np - 1)
    edoor_courbe_max_dever.dx(1) = 0#
    edoor_courbe_max_dever.dy(1) = -hc
'    Dim coef As Double
Dim dx, X As Double
dx = edoor_res.l_ouverture / np
 coef = 1 / (Ham * edoor_res.nbFroudeMax ^ 0.8)
   For i = 1 To np - 1
    dc = dx * i
        edoor_courbe_max_dever.dx(i + 1) = dc
        hc22 = 1 - coefEcoul * (coef * dc) ^ 1.5
'        hc22 = hc22 * Ham
        hc22 = hc22 * hc
        
        edoor_courbe_max_dever.dy(i + 1) = -hc22 + dc * edoor_res.deltaa ' - hc
    Next
Dim zDev As Double
zDev = hc22
hc22 = calcul_Y_pour_X(hc, edoor_res.Ham, edoor_res.nbFroudeMax, edoor_res.deltaa, edoor_res.l_ouverture, coefEcoul)
ept = (1 + 0.06 * edoor_res.l_ouverture * coef) * hc
zDev = ept - zDev
zDev = ept - hc22
'zDev = ept - hc22 - (edoor_res.deltaa * edoor_res.l_ouverture)
Dim coefSep As Double
coefSep = zDev / ept
coefSep = (zDev - (edoor_res.deltaa * edoor_res.l_ouverture)) / ept
coefSep = (zDev - (0 * edoor_res.l_ouverture)) / ept
' calcul max bas
'    hc2 = Hbav + hmin + edessdo.tron_ava.conduit.Diametre
    hc2 = hmin + edessdo.tron_ava.conduit.Diametre
'    hc2 = calcul_hauteur_bas_jet(hc, edoor_res.Ham, edoor_res.nbFroudeMax, edoor_res.deltaa, hc2)
    hc2 = calcul_hauteur_bas_jet(hc, edoor_res.Ham, edoor_res.nbFroudeMax, edoor_res.deltaa, l_chambre1, hc2, edessdo.tron_ava.conduit.pente, coefEcoul)
''    hc2 = Hbav + hmin + edessdo.tron_ava.conduit.Diametre
'    hc2 = hmin + edessdo.tron_ava.conduit.Diametre
' 'calcul    edoor_courbe_max_bas
'dh = hc2 / (np - 1)

'*******
'calcul_courbe_qmax_bas:
dh = hc2 / (np - 1)
    edoor_courbe_max_bas.dx(1) = 0#
    edoor_courbe_max_haut.dx(1) = -hc

    edoor_courbe_max_bas.dy(1) = 0#
'        edoor_courbe_max_haut.dx(1) = 0
'    edoor_courbe_max_haut.dy(1) = -hc  '-Hbav ' ?????? -hc_cri


dvx = edoor_res.l_ouverture * edoor_res.deltaa
dcX = edoor_res.l_ouverture / (edoor_res.Hbav * edoor_res.nbFroudeMax ^ 0.8)
hcZ = 1# / 3# * dcX + 1# / 4# * dcX ^ 2
hcZ = hcZ * edoor_res.Hbav
ept = 1 + 0.06 * dcX
ept = ept * edoor_res.Hbav

Dim y0 As Double, ya As Double, yb As Double, y0c As Double, xi As Double, yi As Double
y0c = 0
For i = 1 To np - 1 '+ 10
hc22 = dh * i
hc22 = hc - dh * i
    hcZ = hc22 / edoor_res.Ham
''    dcX = 2 * (-1 / 3 + (1 / 9 + hcZ) ^ 0.5)
'    dc = dcX * edoor_res.Ham_cri * edoor_res.nbFroude ^ 0.8
'    dc = dc ' - (hc22 / edoor_res.deltaa)
    dc = calcul_X_pour_h(hc, edoor_res.Ham, edoor_res.nbFroudeMax, edoor_res.deltaa, hc22, coefEcoul)
    ept = 1 + 0.06 * dc * coef
    ept = ept * hc
    Seuil = seuil_ini + dc * edoor_res.deltaa
    edoor_courbe_max_bas.dx(i + 1) = dc
'    edoor_courbe_max_bas.dy(i + 1) = -hc22 + ept '+ dc * edoor_res.deltaa
    edoor_courbe_max_bas.dy(i + 1) = -hc22 + ept '+ dc * edoor_res.deltaa
    y0 = -hc22 + ept - coefSep * ept
    ya = (edessdo.tron_amo.radava - y0)
    ya = y0
 '   yb = (edessdo.tron_ava.radamo - (dc - edoor_res.l_chambre1) * edessdo.tron_ava.conduit.pente + edoor_res.Hav)
    

Dim pente0 As Double
'pente0 = edoor_res.deltaa
pente0 = 0
    If dc < l_chambre1 Then
    yb = -(0 - (dc * pente0) + edoor_res.Hav - hmin - edessdo.tron_ava.conduit.Diametre)
    yb = -(0 - (l_chambre1 * pente0 + (dc - l_chambre1) * edessdo.tron_ava.conduit.pente) + edoor_res.Hav - hmin - edessdo.tron_ava.conduit.Diametre)
    Else
    yb = -(0 - (l_chambre1 * pente0 + (dc - l_chambre1) * edessdo.tron_ava.conduit.pente) + edoor_res.Hav - hmin - edessdo.tron_ava.conduit.Diametre)
    yb = -(0 - (l_chambre1 * pente0 + (dc - l_chambre1) * edessdo.tron_ava.conduit.pente) + edoor_res.Hav - hmin - edessdo.tron_ava.conduit.Diametre)
    End If
'    If ya < yb Then
    If ya > yb Then
        If y0c = 0 Then
            Call inters(dc, ya, edoor_courbe_max_haut.dx(i), edoor_courbe_max_haut.dy(i), dc, yb, edoor_courbe_max_haut.dx(i), (yb - ((dc - edoor_courbe_max_haut.dx(i)) * edessdo.tron_ava.conduit.pente)), xi, yi)
            edoor_courbe_max_haut.dx(i + 1) = xi
            edoor_courbe_max_haut.dy(i + 1) = yi '(edessdo.tron_amo.radava - yi)
            y0c = yi
            l_jetaval_h = xi
      Else
        edoor_courbe_max_haut.dx(i + 1) = edoor_courbe_max_haut.dx(i)
        edoor_courbe_max_haut.dy(i + 1) = edoor_courbe_max_haut.dy(i)
       
        End If
        
    Else
    edoor_courbe_max_haut.dx(i + 1) = dc
    edoor_courbe_max_haut.dy(i + 1) = (-hc22 + ept - coefSep * ept) ' + dc * edoor_res.deltaa
    edoor_courbe_max_haut.dy(i + 1) = (-hc22 + ept - coefSep * ept) '+ dc * edoor_res.deltaa

    End If
    
    l_jetaval_b = dc
'    edoor_courbe_max_haut.dy(i + 1) = -hc22 + ept - coefSep * ept

     
 '   edoor_courbe_max_bas.dx(i + 1) = dc
  '  edoor_courbe_max_bas.dy(i + 1) = hc22
Next
For i = 1 To UBound(edoor_courbe_max_bas.dx)
    edoor_courbe_max_bas.dy(i) = edoor_courbe_max_bas.dy(i) + Mini(edoor_courbe_max_bas.dx(i), l_chambre1) * edoor_res.deltaa
Next
For i = 1 To UBound(edoor_courbe_max_haut.dx)
    edoor_courbe_max_haut.dy(i) = edoor_courbe_max_haut.dy(i) + Mini(edoor_courbe_max_haut.dx(i), l_chambre1) * edoor_res.deltaa
Next



'calcul    edoor_courbe_cri_haut



'     mes = mes + Chr(13) + Chr(10) + "Longueur jusqu'à intersection hauteur d'eau aval = " + ajout_zero(Trim(Str(Round(l_jetaval_h, 3)))) + " m"
'     mes = mes + Chr(13) + Chr(10) + "Longueur jusqu'à intersection radier aval = " + ajout_zero(Trim(Str(Round(l_jetaval_b, 3)))) + " m"

'    edoor_res.hmin = hmin
    edoor_res.l_chambre1 = l_chambre1
    edoor_res.l_jetaval_h = l_jetaval_h
    edoor_res.l_jetaval_b = l_jetaval_b
    edo.hauteur = hmin + edessdo.tron_ava.conduit.Diametre
    edo.Absamo = edessdo.tron_amo.Absava
    edo.Longueur = l_chambre1
    edo.Absava = edo.Absamo + edo.Longueur
   edo.pente = edessdo.tron_ava.conduit.pente
'    edo.pente = 0.005
    edo.pente = edoor_res.deltaa
    edo.radava = (edessdo.tron_amo.radava - edo.Longueur * edo.pente) - edessdo.tron_ava.conduit.Diametre - hmin
    edo.radamo = edo.radava + edo.Longueur * edo.pente
    edessdo.tron_ava.Absamo = edo.Absava
    edessdo.tron_ava.radamo = edo.radava
    edessdo.tron_ava.Absava = edessdo.tron_ava.Absamo + edessdo.tron_ava.conduit.Longueur
    edessdo.tron_ava.radava = edessdo.tron_ava.radamo - edessdo.tron_ava.conduit.Longueur * edessdo.tron_ava.conduit.pente

'End If
calcul_courbes_jet = True
End Function
Public Function calcul_mini1(ByRef mes As String, ByVal hmin As Double) As Boolean
Dim l_chambre1 As Double, l_jetaval_b As Double, l_jetaval_h As Double
Dim hc22 As Double, dh As Double, dc As Double
Dim i As Integer, np As Integer
Dim g As Double, l_ouverture As Double
Dim hc As Double, hc1 As Double, hc2 As Double
Dim alpha As Double, CosA As Double, deltaa As Double, vc_cri As Double, vc As Double
Dim Hbav As Double, Hav As Double, hc_cri As Double, hav_cri As Double

Dim hcZ As Double, dcX As Double, ept As Double
Dim dvx As Double


g = 9.81
'    edoor_res.Qbaveff = Qbaveff
'    edoor_res.Qbavth = Qbavth
'    edoor_res.Ham = Ham
Hbav = edoor_res.Hbav
hc = edoor_res.hc ' hauteur à l'overture à Qmax
vc = edoor_res.vc ' vitesse à l'overture à Qmax
l_ouverture = edoor_res.l_ouverture
Hav = edoor_res.Hav
'    edoor_res.hdev = hdev
alpha = edoor_res.alpha
CosA = edoor_res.CosA
deltaa = edoor_res.deltaa
'    edoor_res.Ham_cri = Ham
hc_cri = edoor_res.hc_cri ' hauteur à l'overture à Qref
vc_cri = edoor_res.vc_cri ' vitesse à l'overture à Qref
hav_cri = edoor_res.hav_cri
hc1 = Hbav + hmin
If hc1 > 0 Then
 Dim nb1 As Double, nb2 As Double
 'calcul longueur de chambre
    If deltaa < 0.1 Then
        l_chambre1 = alpha * vc * ((2 * hc1 / g) ^ 0.5)
    Else
        l_chambre1 = alpha * vc * ((2 * hc1 / (g * CosA)) ^ 0.5)
        l_chambre1 = l_chambre1 + hc1 * deltaa

    End If
    
'GoTo partiebassecrit
'calcul    edoor_courbe_max_dever
    hc2 = hc + hmin + (edessdo.tron_ava.conduit.Diametre - Hav) + (l_chambre1 * edessdo.tron_amo.conduit.pente)
    hc2 = hc + hmin + (edessdo.tron_ava.conduit.Diametre - Hav) '+ (l_chambre1 * edessdo.tron_amo.conduit.pente)
    np = UBound(edoor_courbe_max_haut.dx)
    dh = hc2 / (np - 1)
    edoor_courbe_max_dever.dx(1) = 0#
    edoor_courbe_max_dever.dy(1) = -hc
    For i = 1 To np - 1
    hc22 = dh * i
        If deltaa < 0.1 Then
            dc = alpha * vc * (2 * hc22 / g) ^ 0.5
        
        Else
            dc = alpha * vc * (2 * hc22 / (g * CosA)) ^ 0.5
        
            dc = dc + hc22 * deltaa
        End If
        edoor_courbe_max_dever.dx(i + 1) = dc
        edoor_courbe_max_dever.dy(i + 1) = hc22 - hc
    Next

    hc2 = Hbav + hmin + (edessdo.tron_ava.conduit.Diametre - Hav) + (l_chambre1 * edessdo.tron_amo.conduit.pente)
    If deltaa < 0.1 Then
        l_jetaval_h = alpha * vc * (2 * hc2 / g) ^ 0.5
    
    Else
        l_jetaval_h = alpha * vc * (2 * hc2 / (g * CosA)) ^ 0.5
    
        l_jetaval_h = l_jetaval_h + hc2 * deltaa
    End If
    hc2 = Hbav + hmin + (edessdo.tron_ava.conduit.Diametre - Hav) ' + (l_chambre1 * edessdo.tron_amo.conduit.pente)
'calcul    edoor_courbe_max_haut
'a voir avec bas
np = UBound(edoor_courbe_max_haut.dx)
dh = hc2 / (np - 1)
    edoor_courbe_max_haut.dx(1) = l_ouverture
    edoor_courbe_max_haut.dy(1) = -Hbav

For i = 1 To np - 1
hc22 = dh * i
    If deltaa < 0.1 Then
        dc = alpha * vc * (2 * hc22 / g) ^ 0.5
    
    Else
        dc = alpha * vc * (2 * hc22 / (g * CosA)) ^ 0.5
    
        dc = dc + hc22 * deltaa
    End If
    edoor_courbe_max_haut.dx(i + 1) = dc
    edoor_courbe_max_haut.dy(i + 1) = hc22 - Hbav
Next

' calcul max bas
'    hc2 = Hbav + hmin + edessdo.tron_ava.conduit.Diametre
    hc2 = hmin + edessdo.tron_ava.conduit.Diametre + (l_chambre1 * edessdo.tron_amo.conduit.pente)
    If deltaa < 0.1 Then
        l_jetaval_b = alpha * vc * (2 * hc2 / g) ^ 0.5
    
    Else
        l_jetaval_b = alpha * vc * (2 * hc2 / (g * CosA)) ^ 0.5
    
        l_jetaval_b = l_jetaval_b + hc2 * deltaa
    End If
''    hc2 = Hbav + hmin + edessdo.tron_ava.conduit.Diametre
'    hc2 = hmin + edessdo.tron_ava.conduit.Diametre
' 'calcul    edoor_courbe_max_bas
'dh = hc2 / (np - 1)

'*******
'calcul_courbe_qmax_bas:
hc2 = hmin + edessdo.tron_ava.conduit.Diametre + (l_chambre1 * edessdo.tron_amo.conduit.pente)
hc2 = hmin + edessdo.tron_ava.conduit.Diametre ' + (l_chambre1 * edessdo.tron_amo.conduit.pente)
dh = hc2 / (np - 1)
    edoor_courbe_max_bas.dx(1) = 0#
    edoor_courbe_max_bas.dy(1) = 0#
        edoor_courbe_max_haut.dx(1) = 0
    edoor_courbe_max_haut.dy(1) = -hc  '-Hbav ' ?????? -hc_cri


dvx = edoor_res.l_ouverture * edoor_res.deltaa
dcX = edoor_res.l_ouverture / (edoor_res.Hbav * edoor_res.nbFroudeMax ^ 0.8)
hcZ = 1# / 3# * dcX + 1# / 4# * dcX ^ 2
hcZ = hcZ * edoor_res.Hbav
ept = 1 + 0.06 * dcX
ept = ept * edoor_res.Hbav

For i = 1 To np - 1
hc22 = dh * i
GoTo suiteMax

'*******

    If deltaa < 0.1 Then
        dc = alpha * vc * (2 * hc22 / g) ^ 0.5
    
    Else
        dc = alpha * vc * (2 * hc22 / (g * CosA)) ^ 0.5
    
        dc = dc + hc22 * deltaa
    End If
suiteMax:

    hcZ = hc22 / edoor_res.Ham ' ???? edoor_res.Ham_cri
    dcX = 2 * (-1 / 3 + (1 / 9 + hcZ) ^ 0.5)
    dc = dcX * edoor_res.Ham * edoor_res.nbFroudeMax ^ 0.8
    dc = dc ' - (hc22 / edoor_res.deltaa)
    ept = 1 + 0.06 * dcX
     ept = ept * edoor_res.Ham ' ???? hc_cri
   ept = ept * edoor_res.Hbav / edoor_res.hc
    edoor_courbe_max_bas.dx(i + 1) = dc
    edoor_courbe_max_bas.dy(i + 1) = hc22 + (dc * edoor_res.deltaa)
    edoor_courbe_max_haut.dx(i + 1) = dc
    edoor_courbe_max_haut.dy(i + 1) = hc22 - ept + dc * edoor_res.deltaa
     
 '   edoor_courbe_max_bas.dx(i + 1) = dc
  '  edoor_courbe_max_bas.dy(i + 1) = hc22
Next





'calcul    edoor_courbe_cri_haut
' nouveau calcul
partiebassecrit:
GoTo calcul_courbe_qref_bas
hc2 = hc_cri + hmin + (edessdo.tron_ava.conduit.Diametre - hav_cri) + (l_chambre1 * edessdo.tron_amo.conduit.pente)
hc2 = hc_cri + hmin + (edessdo.tron_ava.conduit.Diametre - hav_cri) ' + (l_chambre1 * edessdo.tron_amo.conduit.pente)
np = UBound(edoor_courbe_max_haut.dx)
dh = hc2 / (np - 1)
    edoor_courbe_cri_haut.dx(1) = 0
    edoor_courbe_cri_haut.dy(1) = -hc_cri

For i = 1 To np - 1
hc22 = dh * i
    If deltaa < 0.1 Then
        dc = alpha * vc_cri * (2 * hc22 / g) ^ 0.5
    
    Else
        dc = alpha * vc_cri * (2 * hc22 / (g * CosA)) ^ 0.5
    
        dc = dc + hc22 * deltaa
    End If
    edoor_courbe_cri_haut.dx(i + 1) = dc
    edoor_courbe_cri_haut.dy(i + 1) = hc22 - hc_cri
Next

'calcul    edoor_courbe_cri_bas
calcul_courbe_qref_bas:
hc2 = hmin + edessdo.tron_ava.conduit.Diametre + (l_chambre1 * edessdo.tron_amo.conduit.pente)
hc2 = hmin + edessdo.tron_ava.conduit.Diametre ' + (l_chambre1 * edessdo.tron_amo.conduit.pente)
dh = hc2 / (np - 1)
    edoor_courbe_cri_bas.dx(1) = 0#
    edoor_courbe_cri_bas.dy(1) = 0#
        edoor_courbe_cri_haut.dx(1) = 0
    edoor_courbe_cri_haut.dy(1) = -hc_cri

dvx = edoor_res.l_ouverture * edoor_res.deltaa
dcX = edoor_res.l_ouverture / (edoor_res.Ham_cri * edoor_res.nbFroude ^ 0.8)
hcZ = 1# / 3# * dcX + 1# / 4# * dcX ^ 2
hcZ = hcZ * edoor_res.Ham_cri
ept = 1 + 0.06 * dcX
ept = ept * hc_cri

For i = 1 To np - 1
hc22 = dh * i
GoTo suite
    If deltaa < 0.1 Then
        dc = alpha * vc_cri * (2 * hc22 / g) ^ 0.5
    
    Else
        dc = alpha * vc_cri * (2 * hc22 / (g * CosA)) ^ 0.5
    
        dc = dc + hc22 * deltaa
    End If
    GoTo suite1
suite:
    hcZ = hc22 / edoor_res.Ham_cri
    dcX = 2 * (-1 / 3 + (1 / 9 + hcZ) ^ 0.5)
    dc = dcX * edoor_res.Ham_cri * edoor_res.nbFroude ^ 0.8
    dc = dc ' - (hc22 / edoor_res.deltaa)
    ept = 1 + 0.06 * dcX
     ept = ept * hc_cri
suite1:
    
    edoor_courbe_cri_bas.dx(i + 1) = dc
    edoor_courbe_cri_bas.dy(i + 1) = hc22 + dc * edoor_res.deltaa
    edoor_courbe_cri_haut.dx(i + 1) = dc
    edoor_courbe_cri_haut.dy(i + 1) = hc22 - ept + dc * edoor_res.deltaa
Next
     mes = mes + Chr(13) + Chr(10) + "Longueur de la chambre = " + ajout_zero(Trim(str(Round(l_chambre1, 3)))) + " m"
     mes = mes + Chr(13) + Chr(10) + "Hauteur de la chambre = " + ajout_zero(Trim(str(Round((hmin + edessdo.tron_ava.conduit.Diametre), 3)))) + " m"
'     mes = mes + Chr(13) + Chr(10) + "Longueur jusqu'à intersection hauteur d'eau aval = " + ajout_zero(Trim(Str(Round(l_jetaval_h, 3)))) + " m"
'     mes = mes + Chr(13) + Chr(10) + "Longueur jusqu'à intersection radier aval = " + ajout_zero(Trim(Str(Round(l_jetaval_b, 3)))) + " m"

'    edoor_res.hmin = hmin
    edoor_res.l_chambre1 = l_chambre1
    edoor_res.l_jetaval_h = l_jetaval_h
    edoor_res.l_jetaval_b = l_jetaval_b
    edo.hauteur = hmin + edessdo.tron_ava.conduit.Diametre
    edo.Absamo = edessdo.tron_amo.Absava
    edo.Longueur = l_chambre1
    edo.Absava = edo.Absamo + edo.Longueur

    edo.radava = (edessdo.tron_amo.radava - edo.Longueur * edessdo.tron_amo.conduit.pente) - edessdo.tron_ava.conduit.Diametre - hmin
    edo.pente = edessdo.tron_ava.conduit.pente
    edo.radamo = edo.radava + edo.Longueur * edo.pente
    edessdo.tron_ava.Absamo = edo.Absava
    edessdo.tron_ava.radamo = edo.radava
    edessdo.tron_ava.Absava = edessdo.tron_ava.Absamo + edessdo.tron_ava.conduit.Longueur
    edessdo.tron_ava.radava = edessdo.tron_ava.radamo - edessdo.tron_ava.conduit.Longueur * edessdo.tron_ava.conduit.pente

End If
End Function 'fin mini1
Public Function calcul_mini0(ByRef mes As String, ByVal hmin As Double) As Boolean
Dim l_chambre1 As Double, l_jetaval_b As Double, l_jetaval_h As Double
Dim hc22 As Double, dh As Double, dc As Double
Dim i As Integer, np As Integer
Dim g As Double, l_ouverture As Double
Dim hc As Double, hc1 As Double, hc2 As Double
Dim alpha As Double, CosA As Double, deltaa As Double, vc_cri As Double, vc As Double
Dim Hbav As Double, Hav As Double, hc_cri As Double, hav_cri As Double
g = 9.81
'    edoor_res.Qbaveff = Qbaveff
'    edoor_res.Qbavth = Qbavth
'    edoor_res.Ham = Ham
Hbav = edoor_res.Hbav
hc = edoor_res.hc ' hauteur à l'overture à Qmax
vc = edoor_res.vc ' vitesse à l'overture à Qmax
l_ouverture = edoor_res.l_ouverture
Hav = edoor_res.Hav
'    edoor_res.hdev = hdev
alpha = edoor_res.alpha
CosA = edoor_res.CosA
deltaa = edoor_res.deltaa
'    edoor_res.Ham_cri = Ham
hc_cri = edoor_res.hc_cri ' hauteur à l'overture à Qref
vc_cri = edoor_res.vc_cri ' vitesse à l'overture à Qref
hav_cri = edoor_res.hav_cri
hc1 = Hbav + hmin
If hc1 > 0 Then
 Dim nb1 As Double, nb2 As Double
    If deltaa < 0.1 Then
        l_chambre1 = alpha * vc * ((2 * hc1 / g) ^ 0.5)
    Else
        l_chambre1 = alpha * vc * ((2 * hc1 / (g * CosA)) ^ 0.5)
        l_chambre1 = l_chambre1 + hc1 * deltaa

    End If
'calcul    edoor_courbe_max_dever
    hc2 = hc + hmin + (edessdo.tron_ava.conduit.Diametre - Hav) + (l_chambre1 * edessdo.tron_amo.conduit.pente)
    hc2 = hc + hmin + (edessdo.tron_ava.conduit.Diametre - Hav) '+ (l_chambre1 * edessdo.tron_amo.conduit.pente)
    np = UBound(edoor_courbe_max_haut.dx)
    dh = hc2 / (np - 1)
    edoor_courbe_max_dever.dx(1) = 0#
    edoor_courbe_max_dever.dy(1) = -hc
    For i = 1 To np - 1
    hc22 = dh * i
        If deltaa < 0.1 Then
            dc = alpha * vc * (2 * hc22 / g) ^ 0.5
        
        Else
            dc = alpha * vc * (2 * hc22 / (g * CosA)) ^ 0.5
        
            dc = dc + hc22 * deltaa
        End If
        edoor_courbe_max_dever.dx(i + 1) = dc
        edoor_courbe_max_dever.dy(i + 1) = hc22 - hc
    Next

    hc2 = Hbav + hmin + (edessdo.tron_ava.conduit.Diametre - Hav) + (l_chambre1 * edessdo.tron_amo.conduit.pente)
    If deltaa < 0.1 Then
        l_jetaval_h = alpha * vc * (2 * hc2 / g) ^ 0.5
    
    Else
        l_jetaval_h = alpha * vc * (2 * hc2 / (g * CosA)) ^ 0.5
    
        l_jetaval_h = l_jetaval_h + hc2 * deltaa
    End If
    hc2 = Hbav + hmin + (edessdo.tron_ava.conduit.Diametre - Hav) ' + (l_chambre1 * edessdo.tron_amo.conduit.pente)
'calcul    edoor_courbe_max_haut
np = UBound(edoor_courbe_max_haut.dx)
dh = hc2 / (np - 1)
    edoor_courbe_max_haut.dx(1) = l_ouverture
    edoor_courbe_max_haut.dy(1) = -Hbav

For i = 1 To np - 1
hc22 = dh * i
    If deltaa < 0.1 Then
        dc = alpha * vc * (2 * hc22 / g) ^ 0.5
    
    Else
        dc = alpha * vc * (2 * hc22 / (g * CosA)) ^ 0.5
    
        dc = dc + hc22 * deltaa
    End If
    edoor_courbe_max_haut.dx(i + 1) = dc
    edoor_courbe_max_haut.dy(i + 1) = hc22 - Hbav
Next

' calcul max bas
'    hc2 = Hbav + hmin + edessdo.tron_ava.conduit.Diametre
    hc2 = hmin + edessdo.tron_ava.conduit.Diametre + (l_chambre1 * edessdo.tron_amo.conduit.pente)
    If deltaa < 0.1 Then
        l_jetaval_b = alpha * vc * (2 * hc2 / g) ^ 0.5
    
    Else
        l_jetaval_b = alpha * vc * (2 * hc2 / (g * CosA)) ^ 0.5
    
        l_jetaval_b = l_jetaval_b + hc2 * deltaa
    End If
'    hc2 = Hbav + hmin + edessdo.tron_ava.conduit.Diametre
    hc2 = hmin + edessdo.tron_ava.conduit.Diametre
 'calcul    edoor_courbe_max_bas
dh = hc2 / (np - 1)
    edoor_courbe_max_bas.dx(1) = 0#
    edoor_courbe_max_bas.dy(1) = 0#

For i = 1 To np - 1
hc22 = dh * i
    If deltaa < 0.1 Then
        dc = alpha * vc * (2 * hc22 / g) ^ 0.5
    
    Else
        dc = alpha * vc * (2 * hc22 / (g * CosA)) ^ 0.5
    
        dc = dc + hc22 * deltaa
    End If
    edoor_courbe_max_bas.dx(i + 1) = dc
    edoor_courbe_max_bas.dy(i + 1) = hc22
Next

'calcul    edoor_courbe_cri_haut
' nouveau calcul
GoTo calcul_courbe_qref_bas
hc2 = hc_cri + hmin + (edessdo.tron_ava.conduit.Diametre - hav_cri) + (l_chambre1 * edessdo.tron_amo.conduit.pente)
hc2 = hc_cri + hmin + (edessdo.tron_ava.conduit.Diametre - hav_cri) ' + (l_chambre1 * edessdo.tron_amo.conduit.pente)
np = UBound(edoor_courbe_max_haut.dx)
dh = hc2 / (np - 1)
    edoor_courbe_cri_haut.dx(1) = 0
    edoor_courbe_cri_haut.dy(1) = -hc_cri

For i = 1 To np - 1
hc22 = dh * i
    If deltaa < 0.1 Then
        dc = alpha * vc_cri * (2 * hc22 / g) ^ 0.5
    
    Else
        dc = alpha * vc_cri * (2 * hc22 / (g * CosA)) ^ 0.5
    
        dc = dc + hc22 * deltaa
    End If
    edoor_courbe_cri_haut.dx(i + 1) = dc
    edoor_courbe_cri_haut.dy(i + 1) = hc22 - hc_cri
Next

'calcul    edoor_courbe_cri_bas
calcul_courbe_qref_bas:
hc2 = hmin + edessdo.tron_ava.conduit.Diametre + (l_chambre1 * edessdo.tron_amo.conduit.pente)
 hc2 = hmin + edessdo.tron_ava.conduit.Diametre ' + (l_chambre1 * edessdo.tron_amo.conduit.pente)



dh = hc2 / (np - 1)
    edoor_courbe_cri_bas.dx(1) = 0#
    edoor_courbe_cri_bas.dy(1) = 0#
        edoor_courbe_cri_haut.dx(1) = 0
    edoor_courbe_cri_haut.dy(1) = -hc_cri

Dim hcZ As Double, dcX As Double, ept As Double
Dim dvx As Double
dvx = edoor_res.l_ouverture * edoor_res.deltaa
dcX = edoor_res.l_ouverture / (edoor_res.Ham_cri * edoor_res.nbFroude ^ 0.8)
hcZ = 1# / 3# * dcX + 1# / 4# * dcX ^ 2
hcZ = hcZ * edoor_res.Ham_cri
ept = 1 + 0.06 * dcX
ept = ept * hc_cri

For i = 1 To np - 1
hc22 = dh * i
GoTo suite
    If deltaa < 0.1 Then
        dc = alpha * vc_cri * (2 * hc22 / g) ^ 0.5
    
    Else
        dc = alpha * vc_cri * (2 * hc22 / (g * CosA)) ^ 0.5
    
        dc = dc + hc22 * deltaa
    End If
    GoTo suite1
suite:
    hcZ = hc22 / edoor_res.Ham_cri
    dcX = 2 * (-1 / 3 + (1 / 9 + hcZ) ^ 0.5)
    dc = dcX * edoor_res.Ham_cri * edoor_res.nbFroude ^ 0.8
    dc = dc ' - (hc22 / edoor_res.deltaa)
    ept = 1 + 0.06 * dcX
     ept = ept * hc_cri
suite1:
    
    edoor_courbe_cri_bas.dx(i + 1) = dc
    edoor_courbe_cri_bas.dy(i + 1) = hc22 + dc * edoor_res.deltaa
    edoor_courbe_cri_haut.dx(i + 1) = dc
    edoor_courbe_cri_haut.dy(i + 1) = hc22 - ept + dc * edoor_res.deltaa
Next
     mes = mes + Chr(13) + Chr(10) + "Longueur de la chambre = " + ajout_zero(Trim(str(Round(l_chambre1, 3)))) + " m"
     mes = mes + Chr(13) + Chr(10) + "Hauteur de la chambre = " + ajout_zero(Trim(str(Round((hmin + edessdo.tron_ava.conduit.Diametre), 3)))) + " m"
'     mes = mes + Chr(13) + Chr(10) + "Longueur jusqu'à intersection hauteur d'eau aval = " + ajout_zero(Trim(Str(Round(l_jetaval_h, 3)))) + " m"
'     mes = mes + Chr(13) + Chr(10) + "Longueur jusqu'à intersection radier aval = " + ajout_zero(Trim(Str(Round(l_jetaval_b, 3)))) + " m"

'    edoor_res.hmin = hmin
    edoor_res.l_chambre1 = l_chambre1
    edoor_res.l_jetaval_h = l_jetaval_h
    edoor_res.l_jetaval_b = l_jetaval_b
    edo.hauteur = hmin + edessdo.tron_ava.conduit.Diametre
    edo.Absamo = edessdo.tron_amo.Absava
    edo.Longueur = l_chambre1
    edo.Absava = edo.Absamo + edo.Longueur

    edo.radava = (edessdo.tron_amo.radava - edo.Longueur * edessdo.tron_amo.conduit.pente) - edessdo.tron_ava.conduit.Diametre - hmin
    edo.pente = edessdo.tron_ava.conduit.pente
    edo.radamo = edo.radava + edo.Longueur * edo.pente
    edessdo.tron_ava.Absamo = edo.Absava
    edessdo.tron_ava.radamo = edo.radava
    edessdo.tron_ava.Absava = edessdo.tron_ava.Absamo + edessdo.tron_ava.conduit.Longueur
    edessdo.tron_ava.radava = edessdo.tron_ava.radamo - edessdo.tron_ava.conduit.Longueur * edessdo.tron_ava.conduit.pente

End If
End Function

Public Function calcul_longueur(ByVal hmin As Double) As Boolean
Dim l_chambre1 As Double, l_jetaval_b As Double, l_jetaval_h As Double
Dim hc22 As Double, dh As Double, dc As Double
Dim i As Integer, np As Integer
Dim g As Double, l_ouverture As Double
Dim hc As Double, hc1 As Double, hc2 As Double
Dim alpha As Double, CosA As Double, deltaa As Double, vc_cri As Double, vc As Double
Dim Hbav As Double, Hav As Double, hc_cri As Double, hav_cri As Double
g = 9.81
Hbav = edoor_res.Hbav
vc = edoor_res.vc
l_ouverture = edoor_res.l_ouverture
Hav = edoor_res.Hav
alpha = edoor_res.alpha
CosA = edoor_res.CosA
deltaa = edoor_res.deltaa
hc_cri = edoor_res.hc_cri
vc_cri = edoor_res.vc_cri
hav_cri = edoor_res.hav_cri
hc1 = Hbav + hmin
hc1 = edessdo.tron_ava.conduit.Diametre - Hbav + hmin
Dim coefEcoul As Double
coefEcoul = 0.54
If hc1 > 0 Then
 Dim nb1 As Double, nb2 As Double
    
'l_chambre1 = calcul_X_pour_h(hc_cri, edoor_res.Ham_cri, edoor_res.nbFroude, edoor_res.deltaa, -hc1)

l_chambre1 = calcul_longueur_bas_jet(hc_cri, edoor_res.Ham_cri, edoor_res.nbFroude, edoor_res.deltaa, (hc1 + Hbav), coefEcoul)
    
edoor_res.l_chambre1 = l_chambre1
edo.Longueur = l_chambre1
End If
End Function
Public Function calcul_hauteur(ByVal xlong As Double) As Boolean
Dim l_chambre1 As Double, l_jetaval_b As Double, l_jetaval_h As Double
Dim hc22 As Double, dh As Double, dc As Double
Dim i As Integer, np As Integer
Dim g As Double, l_ouverture As Double
Dim hc As Double, hc1 As Double, hc2 As Double
Dim alpha As Double, CosA As Double, deltaa As Double, vc_cri As Double, vc As Double
Dim Hbav As Double, Hav As Double, hc_cri As Double, hav_cri As Double
Dim CosA2 As Double, sina As Double, sinA2 As Double
Dim hmin As Double
Dim Ham As Double
Dim Ham_cri As Double
Dim coefEcoul As Double
coefEcoul = 0.54
g = 9.81
'    edoor_res.Qbaveff = Qbaveff
'    edoor_res.Qbavth = Qbavth
Ham_cri = edoor_res.Ham_cri
Hbav = edoor_res.Hbav
'    edoor_res.hc = hc
vc = edoor_res.vc
l_ouverture = edoor_res.l_ouverture
Hav = edoor_res.Hav
'    edoor_res.hdev = hdev
alpha = edoor_res.alpha
CosA = edoor_res.CosA
deltaa = edoor_res.deltaa
'    edoor_res.Ham_cri = Ham
hc_cri = edoor_res.hc_cri
vc_cri = edoor_res.vc_cri
hav_cri = edoor_res.hav_cri
'hc1 = Hbav + hmin
edoor_res.l_chambre1 = xlong
l_chambre1 = xlong
''    If deltaa < 0.1 Then
'''        l_chambre1 = alpha * vc * ((2 * hc1 / g) ^ 0.5)
''        nb1 = (xlong / (alpha * vc)) ^ 2
''        hc1 = nb1 * g / 2
''    Else
''        l_chambre1 = alpha * vc * ((2 * hc1 / (g * CosA)) ^ 0.5)
''        l_chambre1 = l_chambre1 + hc1 * deltaa
''        nb1 = (l_chambre1 / (alpha * vc)) ^ 2
''        nb2 = nb1 * (g * CosA) / 2
''        Debug.Print nb2
''    End If
CosA2 = CosA * CosA

If deltaa < 0.1 Then
hc1 = ((l_chambre1 / vc) ^ 2) * g / 2#
hc1 = ((l_chambre1 / (vc * alpha)) ^ 2) * g / 2#
Else
sinA2 = 1# - CosA2
sina = sinA2 ^ 0.5

hc1 = (l_chambre1 ^ 2) * g * CosA / (2# * (((vc * alpha) ^ 2) + l_chambre1 * g * sina))
End If
'hcri As Double, h0 As Double, nbFr As Double, pente As Double, xlong As Double
hc1 = calcul_hauteur_jet_chambre(hc_cri, Ham_cri, edoor_res.nbFroude, edoor_res.deltaa, l_chambre1, coefEcoul)
hmin = hc1 - Hbav
hmin = hc1 - edessdo.tron_ava.conduit.Diametre

edo.tav = hmin
End Function

Public Function calcul_door(ByRef mes As String) As Boolean
Dim l_chambre1 As Double, l_jetaval_b As Double, l_jetaval_h As Double
Dim hc22 As Double, dh As Double, dc As Double
Dim i As Integer, np As Integer
Dim g As Double, l_ouverture As Double
Dim hc As Double, hc1 As Double, hc2 As Double
Dim alpha As Double, CosA As Double, deltaa As Double, vc_cri As Double, vc As Double
Dim Hbav As Double, Hav As Double, hc_cri As Double, hav_cri As Double
Dim CosA2 As Double, sina As Double, sinA2 As Double
Dim hmin As Double
g = 9.81
'    edoor_res.Qbaveff = Qbaveff
'    edoor_res.Qbavth = Qbavth
'    edoor_res.Ham = Ham
Hbav = edoor_res.Hbav
hc = edoor_res.hc
vc = edoor_res.vc
l_ouverture = edoor_res.l_ouverture
Hav = edoor_res.Hav
'    edoor_res.hdev = hdev
alpha = edoor_res.alpha
CosA = edoor_res.CosA
deltaa = edoor_res.deltaa
'    edoor_res.Ham_cri = Ham
hc_cri = edoor_res.hc_cri
vc_cri = edoor_res.vc_cri
hav_cri = edoor_res.hav_cri
'hc1 = Hbav + hmin
l_chambre1 = edoor_res.l_chambre1
hmin = edo.tav
Dim ok As Boolean
ok = calcul_courbes_jet(mes, hmin)
calcul_door = ok
End Function
Public Function calcul_courbes() As Boolean
Dim l_chambre1 As Double, l_jetaval_b As Double, l_jetaval_h As Double
Dim hc22 As Double, dh As Double, dc As Double
Dim i As Integer, np As Integer
Dim g As Double, l_ouverture As Double
Dim hc As Double, hc1 As Double, hc2 As Double
Dim alpha As Double, CosA As Double, deltaa As Double, vc_cri As Double, vc As Double
Dim Hbav As Double, Hav As Double, hc_cri As Double, hav_cri As Double
Dim CosA2 As Double, sina As Double, sinA2 As Double
Dim hmin As Double
g = 9.81
'    edoor_res.Qbaveff = Qbaveff
'    edoor_res.Qbavth = Qbavth
'    edoor_res.Ham = Ham
Hbav = edoor_res.Hbav
hc = edoor_res.hc
vc = edoor_res.vc
l_ouverture = edoor_res.l_ouverture
Hav = edoor_res.Hav
'    edoor_res.hdev = hdev
alpha = edoor_res.alpha
CosA = edoor_res.CosA
deltaa = edoor_res.deltaa
'    edoor_res.Ham_cri = Ham
hc_cri = edoor_res.hc_cri
vc_cri = edoor_res.vc_cri
hav_cri = edoor_res.hav_cri
'hc1 = Hbav + hmin
'l_chambre1 = edoor_res.l_chambre1
l_chambre1 = edo.Longueur
hmin = edo.tav

Dim ok As Boolean
Dim mes As String

mes = ""
ok = calcul_courbes_jet(mes, hmin)
GoTo fin
'calcul    edoor_courbe_max_dever
    hc2 = hc + hmin + (edessdo.tron_ava.conduit.Diametre - Hav) + (l_chambre1 * edessdo.tron_amo.conduit.pente)
'    hc2 = hc + hmin + (edessdo.tron_ava.conduit.Diametre - Hav) + (l_chambre1 * edessdo.tron_dech.conduit.pente)
    hc2 = hc + hmin + (edessdo.tron_ava.conduit.Diametre - Hav) '+ (l_chambre1 * edessdo.tron_amo.conduit.pente)

    np = UBound(edoor_courbe_max_haut.dx)
    dh = hc2 / (np - 1)
    edoor_courbe_max_dever.dx(1) = 0#
    edoor_courbe_max_dever.dy(1) = -hc
    For i = 1 To np - 1
    hc22 = dh * i
        If deltaa < 0.1 Then
            dc = alpha * vc * (2 * hc22 / g) ^ 0.5
        
        Else
            dc = alpha * vc * (2 * hc22 / (g * CosA)) ^ 0.5
        
            dc = dc + hc22 * deltaa
        End If
        edoor_courbe_max_dever.dx(i + 1) = dc
        edoor_courbe_max_dever.dy(i + 1) = hc22 - hc
    Next

    hc2 = Hbav + hmin + (edessdo.tron_ava.conduit.Diametre - Hav) + (l_chambre1 * edessdo.tron_amo.conduit.pente)
    If deltaa < 0.1 Then
        l_jetaval_h = alpha * vc * (2 * hc2 / g) ^ 0.5

    Else
        l_jetaval_h = alpha * vc * (2 * hc2 / (g * CosA)) ^ 0.5

        l_jetaval_h = l_jetaval_h + hc2 * deltaa
    End If
    hc2 = Hbav + hmin + (edessdo.tron_ava.conduit.Diametre - Hav) ' + (l_chambre1 * edessdo.tron_amo.conduit.pente)
'calcul    edoor_courbe_max_haut
np = UBound(edoor_courbe_max_haut.dx)
dh = hc2 / (np - 1)
    edoor_courbe_max_haut.dx(1) = l_ouverture
    edoor_courbe_max_haut.dy(1) = -Hbav

For i = 1 To np - 1
hc22 = dh * i
    If deltaa < 0.1 Then
        dc = alpha * vc * (2 * hc22 / g) ^ 0.5
    
    Else
        dc = alpha * vc * (2 * hc22 / (g * CosA)) ^ 0.5
    
        dc = dc + hc22 * deltaa
    End If
    edoor_courbe_max_haut.dx(i + 1) = dc
    edoor_courbe_max_haut.dy(i + 1) = hc22 - Hbav
Next
'calcul    edoor_courbe_cri_haut
hc2 = hc_cri + hmin + (edessdo.tron_ava.conduit.Diametre - hav_cri) + (l_chambre1 * edessdo.tron_amo.conduit.pente)
hc2 = hc_cri + hmin + (edessdo.tron_ava.conduit.Diametre - hav_cri) '+ (l_chambre1 * edessdo.tron_amo.conduit.pente)
np = UBound(edoor_courbe_max_haut.dx)
dh = hc2 / (np - 1)
    edoor_courbe_cri_haut.dx(1) = 0
    edoor_courbe_cri_haut.dy(1) = -hc_cri

For i = 1 To np - 1
hc22 = dh * i
    If deltaa < 0.1 Then
        dc = alpha * vc_cri * (2 * hc22 / g) ^ 0.5
    
    Else
        dc = alpha * vc_cri * (2 * hc22 / (g * CosA)) ^ 0.5
    
        dc = dc + hc22 * deltaa
    End If
    edoor_courbe_cri_haut.dx(i + 1) = dc
    edoor_courbe_cri_haut.dy(i + 1) = hc22 - hc_cri
Next


'    hc2 = Hbav + hmin + edessdo.tron_ava.conduit.Diametre
    hc2 = hmin + edessdo.tron_ava.conduit.Diametre + (l_chambre1 * edessdo.tron_amo.conduit.pente)
    If deltaa < 0.1 Then
        l_jetaval_b = alpha * vc * (2 * hc2 / g) ^ 0.5

    Else
        l_jetaval_b = alpha * vc * (2 * hc2 / (g * CosA)) ^ 0.5

        l_jetaval_b = l_jetaval_b + hc2 * deltaa
    End If
'calcul    edoor_courbe_max_bas
   hc2 = Hbav + hmin + edessdo.tron_ava.conduit.Diametre
   hc2 = hmin + edessdo.tron_ava.conduit.Diametre
dh = hc2 / (np - 1)
    edoor_courbe_max_bas.dx(1) = 0#
    edoor_courbe_max_bas.dy(1) = 0#

For i = 1 To np - 1
hc22 = dh * i
    If deltaa < 0.1 Then
        dc = alpha * vc * (2 * hc22 / g) ^ 0.5
    
    Else
        dc = alpha * vc * (2 * hc22 / (g * CosA)) ^ 0.5
    
        dc = dc + hc22 * deltaa
    End If
    edoor_courbe_max_bas.dx(i + 1) = dc
    edoor_courbe_max_bas.dy(i + 1) = hc22
Next
'calcul    edoor_courbe_cri_bas
hc2 = hmin + edessdo.tron_ava.conduit.Diametre + (l_chambre1 * edessdo.tron_amo.conduit.pente)
hc2 = hmin + edessdo.tron_ava.conduit.Diametre ' + (l_chambre1 * edessdo.tron_amo.conduit.pente)
dh = hc2 / (np - 1)
    edoor_courbe_cri_bas.dx(1) = 0#
    edoor_courbe_cri_bas.dy(1) = 0#

For i = 1 To np - 1
hc22 = dh * i
    If deltaa < 0.1 Then
        dc = alpha * vc_cri * (2 * hc22 / g) ^ 0.5
    
    Else
        dc = alpha * vc_cri * (2 * hc22 / (g * CosA)) ^ 0.5
    
        dc = dc + hc22 * deltaa
    End If
    edoor_courbe_cri_bas.dx(i + 1) = dc
    edoor_courbe_cri_bas.dy(i + 1) = hc22
Next
'     mes = mes + Chr(13) + Chr(10) + "Longueur de la chambre = " + ajout_zero(Trim(Str(Round(l_chambre1, 3)))) + " m"
'     mes = mes + Chr(13) + Chr(10) + "Hauteur de la chambre = " + ajout_zero(Trim(Str(Round((hmin + edessdo.tron_ava.conduit.Diametre), 3)))) + " m"
'     mes = mes + Chr(13) + Chr(10) + "Longueur jusqu'à intersection hauteur d'eau aval = " + ajout_zero(Trim(str(Round(l_jetaval_h, 3)))) + " m"
'    mes = mes + Chr(13) + Chr(10) + "Longueur jusqu'à intersection radier aval = " + ajout_zero(Trim(str(Round(l_jetaval_b, 3)))) + " m"
'
''    edoor_res.hmin = hmin
    edoor_res.l_chambre1 = l_chambre1
    edoor_res.l_jetaval_h = l_jetaval_h
    edoor_res.l_jetaval_b = l_jetaval_b
    edo.hauteur = hmin + edessdo.tron_ava.conduit.Diametre
    edo.Absamo = edessdo.tron_amo.Absava
    edo.Longueur = l_chambre1
    edo.Absava = edo.Absamo + edo.Longueur

    edo.radava = (edessdo.tron_amo.radava - edo.Longueur * edessdo.tron_amo.conduit.pente) - edessdo.tron_ava.conduit.Diametre - hmin
    edo.pente = edessdo.tron_ava.conduit.pente
    edo.radamo = edo.radava + edo.Longueur * edo.pente
    edessdo.tron_ava.Absamo = edo.Absava
    edessdo.tron_ava.radamo = edo.radava
    edessdo.tron_ava.Absava = edessdo.tron_ava.Absamo + edessdo.tron_ava.conduit.Longueur
    edessdo.tron_ava.radava = edessdo.tron_ava.radamo - edessdo.tron_ava.conduit.Longueur * edessdo.tron_ava.conduit.pente
fin:
End Function

Function rech_hauteur_amont(ByRef tr As troncon, ByVal hc0 As Double, ByRef Qcal1 As Double, ByVal dh As Double) As Double
Dim Ham As Double, Q As Double, dh0 As Double, v As Double, larg As Double, hc As Double, dq As Double
Dim res_cond As debit_conduit
Dim nbFroude As Double
hc = hc0 - dh
dq = Qcal1 * dh / hc0
dh0 = dh
Q = Qcal1
While Abs(dh) > 0.0001
Q = Q + dq
res_cond = calc_debit_tr(tr, Q)
Ham = res_cond.hauteur
v = res_cond.vitesse
larg = res_cond.largeurlibre
nbFroude = calcul_Froude1(v, (res_cond.surface / res_cond.largeurlibre))
nbFroude = calcul_Froude(Q, Ham, tr.conduit.Diametre)

Dim betaL As Double, Fbeta As Double

betaL = res_cond.largeurlibre / tr.conduit.Diametre

Fbeta = nbFroude / betaL ^ 0.5
hc = Ham * (2 * Fbeta ^ 2 / (1 + 2 * Fbeta ^ 2)) ^ (2# / 3#)
dh = hc0 - hc
If dh * dh0 < 0 Then
    dq = -dq / 2
    dh0 = dh
End If
Wend
Qcal1 = Q
rech_hauteur_amont = Ham
End Function
Public Function calcul_LEAPING_WEAR(ByRef mes As String, Optional codeCalc As String) As Boolean
' Calcul pour les débits critique Qcr (Rincage) : Longueur Maxi de l'ouverture
'  et d'Orage (Qmax) : Débit conservé Qbav
Dim Ham As Double, Vam As Double, Hav As Double, hdev As Double, hav_cri As Double
Dim Hbav As Double, Vav As Double, hc_cri As Double, vc_cri As Double
Dim Sav As Double
Dim Phil As Double, rbav As Double, sbav As Double, SSbav As Double
Dim Qbavth As Double, Qbaveff As Double, epsi As Double
Dim g As Double
Dim hc As Double, hc1 As Double, hc2 As Double
Dim Rham As Double
Dim Bouam As Double, iRadam As Double
Dim CosA2 As Double, CosA As Double, deltaa As Double
Dim sinA2 As Double, sina As Double
Dim qcal As Double, Qcr As Double, Qmax As Double
Dim Qreste As Double
Dim res_cond As debit_conduit
Dim tr As troncon
Dim Sc As Double, vc As Double, alpha As Double
Dim lc As Double
Dim l_ouverture As Double, hmin As Double, l_chambre1 As Double, l_jetaval_b As Double, l_jetaval_h As Double
Dim ok As Boolean, ok1 As Boolean
Dim hc22 As Double, dh As Double, dc As Double
Dim i As Integer, np As Integer
Dim coefEcoul As Double
coefEcoul = 0.54
g = 9.81
calcul_LEAPING_WEAR = True
'  calcul pour débit critique

Qcr = edessdo.Qrin / 1000
qcal = Qcr
mes = " Résultats pour le débit de référence " + ajout_zero(Trim(str(Round(qcal, 3)))) + " m3/s"

tr = edessdo.tron_amo
res_cond = calc_debit_tr(tr, qcal)
Ham = res_cond.hauteur
Vam = res_cond.vitesse
Dim largOuverture As Double
largOuverture = res_cond.largeurlibre
 deltaa = edoor_res.deltaa
If Len(codeCalc) = 0 Then
    deltaa = tr.conduit.pente
End If
If edoor_res.deltaa = 0 Then
    edoor_res.deltaa = tr.conduit.pente
    deltaa = tr.conduit.pente
End If
Dim nbFroude As Double
nbFroude = calcul_Froude1(res_cond.vitamo, (res_cond.surface / res_cond.largeurlibre))
nbFroude = calcul_Froude(qcal, Ham, tr.conduit.Diametre)
    
    mes = mes + Chr(13) + Chr(10) + "    Hauteur normale à l'amont = " + ajout_zero(Trim(str(Round(Ham, 3)))) + " m"
    mes = mes + Chr(13) + Chr(10) + "    Vitesse à l'amont = " + ajout_zero(Trim(str(Round(Vam, 3)))) + " m/s"
    mes = mes + Chr(13) + Chr(10) + "    Nombre de Froude : = " + ajout_zero(Trim(str(Round(nbFroude, 3))))
'deltaa = 0.06 'tr.conduit.pente
'deltaa = 0.005
'deltaa = 0.002
' nouveau calcul 20061219
Dim betaL As Double, Fbeta As Double

betaL = res_cond.largeurlibre / tr.conduit.Diametre

Fbeta = nbFroude / betaL ^ 0.5
hc = Ham * (2 * Fbeta ^ 2 / (1 + 2 * Fbeta ^ 2)) ^ (2# / 3#)


'    beta = 2 * arccosinus((1 - 2 * hc / tr.conduit.Diametre))
'    Sc = (1# / 8#) * (tr.conduit.Diametre ^ 2) * (beta - Sin(beta))
'    lc = tr.conduit.Diametre * Sin(beta / 2)
    beta = 2 * arccosinus((1 - 2 * Ham / tr.conduit.Diametre))
    Sc = (1# / 8#) * (tr.conduit.Diametre ^ 2) * (beta - Sin(beta))
    lc = tr.conduit.Diametre * Sin(beta / 2)
    largOuverture = lc
    '    v = res_conduit.debit / s
    vc = qcal / Sc
    mes = mes + Chr(13) + Chr(10) + "    Hauteur d'eau à l'ouverture = " + ajout_zero(Trim(str(Round(hc, 3)))) + " m"
    mes = mes + Chr(13) + Chr(10) + "    Vitesse à l'ouverture = " + ajout_zero(Trim(str(Round(vc, 3)))) + " m/s"

    
    
' arrondi lc
Dim lc1 As Double, hc0 As Double, ray As Double, Qcal1 As Double
Qcal1 = qcal
ray = tr.conduit.Diametre / 2#
lc1 = Round(lc, 2)
If codeCalc = "LARG" And edoor_res.l_largOuverture < tr.conduit.Diametre And edoor_res.l_largOuverture > 0 Then
lc1 = edoor_res.l_largOuverture
End If
mes = mes + Chr(13) + Chr(10) + "    Largeur de l'ouverture = " + ajout_zero(Trim(str(Round(lc, 3)))) + " m"
mes = mes + Chr(13) + Chr(10) + "    Largeur retenue = " + ajout_zero(Trim(str(Round(lc1, 2)))) + " m"
hc0 = ray - Sqr((ray ^ 2 - (lc1 / 2#) ^ 2))
dh = hc0 - hc
'Ham = rech_hauteur_amont(tr, hc0, Qcal1, dh)
'res_cond = calc_debit_tr(tr, Qcal1)
'Ham = res_cond.hauteur
'Vam = res_cond.vitesse
Ham = hc0

res_cond = calc_hauteur_tr(tr, hc0)
Qcal1 = res_cond.debit
res_cond = calc_debit_tr(tr, Qcal1)
nbFroude = calcul_Froude1(res_cond.vitamo, (res_cond.surface / res_cond.largeurlibre))
nbFroude = Qcal1 / (9.81 * tr.conduit.Diametre * hc0 ^ 4) ^ 0.5
nbFroude = calcul_Froude(Qcal1, Ham, tr.conduit.Diametre)

  mes = mes + Chr(13) + Chr(10) + "    débit limite = " + ajout_zero(Trim(str(Round(Qcal1, 3)))) + " m3/s"
  mes = mes + Chr(13) + Chr(10) + "    Hauteur normale à l'amont = " + ajout_zero(Trim(str(Round(Ham, 3)))) + " m"
    mes = mes + Chr(13) + Chr(10) + "    Vitesse à l'amont = " + ajout_zero(Trim(str(Round(Vam, 3)))) + " m/s"
    mes = mes + Chr(13) + Chr(10) + "    Nombre de Froude : = " + ajout_zero(Trim(str(Round(nbFroude, 3))))
betaL = res_cond.largeurlibre / tr.conduit.Diametre

Fbeta = nbFroude / betaL ^ 0.5
hc = Ham * (2 * Fbeta ^ 2 / (1 + 2 * Fbeta ^ 2)) ^ (2# / 3#)


    beta = 2 * arccosinus((1 - 2 * hc / tr.conduit.Diametre))
    Sc = (1# / 8#) * (tr.conduit.Diametre ^ 2) * (beta - Sin(beta))
    lc = tr.conduit.Diametre * Sin(beta / 2)
    largOuverture = lc1
'suite

'    v = res_conduit.debit / s
    vc = Qcal1 / Sc
    mes = mes + Chr(13) + Chr(10) + "    Hauteur d'eau à l'ouverture = " + ajout_zero(Trim(str(Round(hc, 3)))) + " m"
    mes = mes + Chr(13) + Chr(10) + "    Vitesse à l'ouverture = " + ajout_zero(Trim(str(Round(vc, 3)))) + " m/s"


hc_cri = hc
vc_cri = vc
alpha = recup_alpha_lw(deltaa)

suiteFroude:
  Dim okR As Integer

If res_cond.hauteur / tr.conduit.Diametre > 0.1 And res_cond.hauteur / tr.conduit.Diametre < 0.35 Then
    l_ouverture = res_cond.hauteur * nbFroude
    l_ouverture = calcul_long_ouverture(hc_cri, res_cond.hauteur, nbFroude, deltaa, coefEcoul) 'tr.conduit.pente)
Else
    okR = MsgBox("Le rapport hauteur/diamètre " + ajout_zero(Trim(str(Round(res_cond.hauteur / tr.conduit.Diametre, 3)))) + " ne respecte les conditions du calcul." + Chr(13) + "        0.1 < h/D < 0.35 " + Chr(13) + " Voulez-vous continuer le calcul?", vbYesNo, "DO ouverture de radier")
End If
If okR = 7 Then
calcul_LEAPING_WEAR = False

    Exit Function
End If
    l_ouverture = res_cond.hauteur * nbFroude
    l_ouverture = calcul_long_ouverture(hc_cri, res_cond.hauteur, nbFroude, deltaa, coefEcoul) 'tr.conduit.pente)

'GoTo suite1
hmin = edessdo.Centon
edo.tav = hmin
hc1 = hc + hmin
'revoir calcul chambre
Dim nb1 As Double, nb2 As Double
' hc, ham, nbFroude, deltaa,hc1 +
l_chambre1 = calcul_longueur_bas_jet(hc, Ham, nbFroude, deltaa, (hmin + edessdo.tron_ava.conduit.Diametre), coefEcoul)

edoor_res.l_largOuverture = largOuverture
edoor_res.nbFroude = nbFroude
edoor_res.l_chambre1 = l_chambre1
suite1:
'l_chambre1 = 1
    mes = mes + Chr(13) + Chr(10) + "    Longueur de l'ouverture = " + ajout_zero(Trim(str(Round(l_ouverture, 3)))) + " m"
'    mes = mes + Chr(13) + Chr(10) + "Longueur de la chambre = " + ajout_zero(Trim(Str(Round(l_chambre1, 3)))) + " m"
' l_chambre1= longueur de la chambre pour que le flux entre au mieux ds la conduite aval
' pour le débit critique
'GoTo suitePartage
    edoor_res.Ham_cri = Ham
    edoor_res.Vam_cri = Vam
    edoor_res.hc_cri = hc
    edoor_res.vc_cri = vc
'
''''''qbav=qcal
' Calcul pour débit max

Qmax = edessdo.Qpluie / 1000
qcal = Qmax
mes = mes + Chr(13) + Chr$(10) + " Résultats pour le débit d'orage " + ajout_zero(Trim(str(Round(qcal, 3)))) + " m3/s"
res_cond = calc_debit_tr(tr, qcal)

nbFroude = calcul_Froude1(res_cond.vitamo, (res_cond.surface / res_cond.largeurlibre))
nbFroude = calcul_Froude(qcal, res_cond.hauteur, tr.conduit.Diametre)
edoor_res.nbFroudeMax = nbFroude
Ham = res_cond.hauteur
Vam = res_cond.vitesse
    mes = mes + Chr(13) + Chr(10) + "    Hauteur d'eau à l'amont = " + ajout_zero(Trim(str(Round(Ham, 3)))) + " m"
    mes = mes + Chr(13) + Chr(10) + "    Vitesse à l'amont = " + ajout_zero(Trim(str(Round(Vam, 3)))) + " m/s"
    mes = mes + Chr(13) + Chr(10) + "    Nombre de Froude : = " + ajout_zero(Trim(str(Round(nbFroude, 3))))
Rham = (Vam / ((tr.conduit.pente ^ 0.5) * tr.conduit.rugosite)) ^ 1.5
Bouam = Vam / (g * Rham) ^ 0.5
hc = (2 * Bouam ^ 2 / (2 * Bouam ^ 2 + CosA2)) * Ham
    beta = 2 * arccosinus((1 - 2 * hc / tr.conduit.Diametre))
    Sc = (1 / 8) * (tr.conduit.Diametre ^ 2) * (beta - Sin(beta))
'    v = res_conduit.debit / s
    vc = qcal / Sc
' nouveau calcul 20061219
'Dim betaL As Double, Fbeta As Double

betaL = res_cond.largeurlibre / tr.conduit.Diametre

Fbeta = nbFroude / betaL ^ 0.5
hc = Ham * (2 * Fbeta ^ 2 / (1 + 2 * Fbeta ^ 2)) ^ (2# / 3#)
'hc = (2 * Bouam ^ 2 / (2 * Bouam ^ 2 + CosA2)) * Ham
    beta = 2 * arccosinus((1 - 2 * hc / tr.conduit.Diametre))
    Sc = (1 / 8) * (tr.conduit.Diametre ^ 2) * (beta - Sin(beta))
'    v = res_conduit.debit / s
    vc = qcal / Sc
    
    
'Lc = tr.conduit.Diametre * Sin(beta / 2)
    mes = mes + Chr(13) + Chr(10) + "    Hauteur d'eau à l'ouverture = " + ajout_zero(Trim(str(Round(hc, 3)))) + " m"
    mes = mes + Chr(13) + Chr(10) + "    Vitesse à l'ouverture = " + ajout_zero(Trim(str(Round(vc, 3)))) + " m/s"
'GoTo suitePartage

Vav = (Vam ^ 2 + (2 * g * Ham * CosA)) ^ 0.5
Sav = qcal / Vav
If deltaa < 0.1 Then
Hbav = ((l_ouverture / vc) ^ 2) * g / 2#
'Hbav = ((l_ouverture / (vc * alpha)) ^ 2) * g / 2#
Else
sinA2 = 1# - CosA2
sina = sinA2 ^ 0.5

Hbav = (l_ouverture ^ 2) * g * CosA / (2# * (((vc * alpha) ^ 2) + l_ouverture * g * sina))
Hbav = (l_ouverture ^ 2) * g * CosA / (2# * (((vc * alpha) ^ 2) + l_ouverture * g * sina))
End If
'GoTo suitePartage

rbav = (lc ^ 2) / (8 * Hbav) + (Hbav / 2)
Phil = 2 * (arccosinus((rbav - Hbav) / rbav))
sbav = Phil * rbav
SSbav = 0.5 * (sbav * rbav - lc * (rbav - Hbav))
Qbavth = SSbav * Vav
epsi = 2.02
epsi = recup_epsi_lw(Qmax / Qcr)
Qbaveff = epsi * Qbavth
''     mes = mes + Chr(13) + Chr(10) + "Hauteur d'eau conservée à l'aval = " + ajout_zero(Trim(Str(Round(Hbav, 3)))) + " m"
 '    mes = mes + Chr(13) + Chr(10) + "    Ex Débit conservé à l'aval: théorique  = " + ajout_zero(Trim(str(Round(Qbavth, 3)))) + " m3/s ; effectif = " + ajout_zero(Trim(str(Round(Qbaveff, 3)))) + " m3/s"
''    mes = mes + Chr(13) + Chr(10) + "Débit effectif conservé à l'aval = " + ajout_zero(Trim(Str(Round(Qbaveff, 3)))) + " m3/s"
    edoor_res.Qbaveff = Qbaveff
    edoor_res.Qbavth = Qbavth
    edoor_res.Ham = Ham
    edoor_res.Vam = Vam
    edoor_res.Hbav = Hbav
    edoor_res.hc = hc
    edoor_res.vc = vc
    edoor_res.l_ouverture = l_ouverture
    edoor_res.alpha = alpha
    edoor_res.CosA = CosA
    edoor_res.deltaa = deltaa
suitePartage:
'nouveau calcul
'Ham = 0.2967
'largOuverture = 0.61
'l_ouverture = 0.58

Qbaveff = (0.61 * largOuverture * l_ouverture * (2 * 9.81 * Ham) ^ 0.5) - (0.14 * (largOuverture ^ 3 / (tr.conduit.Diametre * Ham ^ 2)) ^ 0.5 * qcal)
Dim coefSep As Double
Dim qConserve As Double
coefSep = calcul_coefSep()
Dim scSep As Double
    beta = 2 * arccosinus((1 - 2 * hc * coefSep / tr.conduit.Diametre))
    scSep = (1 / 8) * (tr.conduit.Diametre ^ 2) * (beta - Sin(beta))
'    v = res_conduit.debit / s
qConserve = vc * scSep
 '   vc = qcal / Sc



Dim ratioQc As Double
ratioQc = (Qbaveff - Qcr) / Qcr
ratioQc = (Qbaveff - Qcal1) / Qcal1
    edoor_res.l_ouverture = l_ouverture
    edoor_res.Qbaveff = Qbaveff
 '   deltaa = tr.conduit.pente
 '   edoor_res.deltaa = 0 ' deltaa
    edoor_res.deltaa = deltaa
     mes = mes + Chr(13) + Chr(10) + "    Débit conservé à l'aval = " + ajout_zero(Trim(str(Round(Qbaveff, 3)))) + " m3/s"
If ratioQc > 0.3 Then
MsgBox "ratio Qconservé sur Q critique : " + ajout_zero(Trim(str(Round(ratioQc, 3)))) + " > 0.3 ", vbExclamation, "Débit conservé à Qmax"
End If
     mes = mes + Chr(13) + Chr(10) + "    ratio Qconservé sur Q critique = " + ajout_zero(Trim(str(Round(ratioQc, 3))))




ok = verif_aval()  ' avec Qbaveff


If ok Then
 ''calcul hauteur d'eau dans conduite aval(depart)
    res_cond = calc_debit_tr(edessdo.tron_ava, Qbaveff)
    Hav = res_cond.hauteur
    edoor_res.Hav = Hav
    res_cond = calc_debit_tr(edessdo.tron_ava, Qcr)
    hav_cri = res_cond.hauteur
    edoor_res.hav_cri = hav_cri
''calcul hauteur d'eau dans conduite deversement
    Qreste = qcal - Qbaveff
    res_cond = calc_debit_tr(edessdo.tron_dech, Qreste)
    hdev = res_cond.hauteur
    edoor_res.hdev = hdev
''''calcul avec hmin par défaut
    ok1 = calcul_mini(mes, hmin)

Else
calcul_LEAPING_WEAR = False
End If
End Function
Private Function recup_alpha_lw(ByVal Gran As Double) As Double
Dim list_alpha(41, 2) As Double
Dim a As Double
Dim X As Double
Dim i As Integer
list_alpha(1, 1) = 0#

'For i = 2 To 26
'list_alpha(i, 1) = (i - 1) * 0.002
'Next
list_alpha(1, 1) = 0#

list_alpha(2, 1) = 1#
For i = 2 To 41
    X = 1 + (i - 1) * 0.01
    list_alpha(i, 2) = X
    list_alpha(i, 1) = (10416.667 * X * X - 19875 * X + 9458.333) / 1000#
Next

i = 1
While Gran > list_alpha(i, 1) And i < UBound(list_alpha)
    i = i + 1
    
Wend
If i = 1 Then
    i = 2
End If
a = (Gran - list_alpha(i - 1, 1)) * (list_alpha(i, 2) - list_alpha(i - 1, 2)) / (list_alpha(i, 1) - list_alpha(i - 1, 1)) + list_alpha(i - 1, 2)
recup_alpha_lw = a
End Function
Private Function recup_epsi_lw(ByVal Gran As Double) As Double
Dim list_alpha(35, 2) As Double
Dim a As Double
Dim X As Double
Dim i As Integer
list_alpha(1, 1) = 1

list_alpha(2, 1) = 1#
For i = 2 To 35
    X = 1 + (i - 1) * 1
    list_alpha(i, 1) = X
Next
i = 1
i = i + 1 '2
list_alpha(i, 2) = 1.32
i = i + 1
list_alpha(i, 2) = 1.45
i = i + 1
list_alpha(i, 2) = 1.58
i = i + 1 '5
list_alpha(i, 2) = 1.66
i = i + 1
list_alpha(i, 2) = 1.74
i = i + 1
list_alpha(i, 2) = 1.82
i = i + 1
list_alpha(i, 2) = 1.88
i = i + 1
list_alpha(i, 2) = 1.94
i = i + 1 '10
list_alpha(i, 2) = 2
i = i + 1
list_alpha(i, 2) = 2.05
i = i + 1
list_alpha(i, 2) = 2.1
i = i + 1
list_alpha(i, 2) = 2.15
i = i + 1
list_alpha(i, 2) = 2.2
i = i + 1  '15
list_alpha(i, 2) = 2.24
i = i + 1
list_alpha(i, 2) = 2.29
i = i + 1
list_alpha(i, 2) = 2.33
i = i + 1
list_alpha(i, 2) = 2.37
i = i + 1
list_alpha(i, 2) = 2.41
i = i + 1  '20
list_alpha(i, 2) = 2.45
i = i + 1
list_alpha(i, 2) = 2.49
i = i + 1
list_alpha(i, 2) = 2.52
i = i + 1
list_alpha(i, 2) = 2.56
i = i + 1
list_alpha(i, 2) = 2.59
i = i + 1  '25
list_alpha(i, 2) = 2.63
i = i + 1
list_alpha(i, 2) = 2.66
i = i + 1
list_alpha(i, 2) = 2.7
i = i + 1
list_alpha(i, 2) = 2.73
i = i + 1
list_alpha(i, 2) = 2.76
i = i + 1  '30
list_alpha(i, 2) = 2.79
i = i + 1
list_alpha(i, 2) = 2.82
i = i + 1
list_alpha(i, 2) = 2.85
i = i + 1
list_alpha(i, 2) = 2.88
i = i + 1
list_alpha(i, 2) = 2.91
i = i + 1  '35
list_alpha(i, 2) = 2.94






i = 1
While Gran > list_alpha(i, 1) And i < UBound(list_alpha)
    i = i + 1
    
Wend
If i = 1 Then
    i = 2
End If
a = (Gran - list_alpha(i - 1, 1)) * (list_alpha(i, 2) - list_alpha(i - 1, 2)) / (list_alpha(i, 1) - list_alpha(i - 1, 1)) + list_alpha(i - 1, 2)
recup_epsi_lw = a
End Function

