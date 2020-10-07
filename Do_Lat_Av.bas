Attribute VB_Name = "Do_Lat_Av"
Public Function test_do_lat(ByRef eds As st_dessdo) As Boolean
Dim ok As Boolean
Dim Q As Double
Dim qv As deb_vit
Dim tr As troncon
Dim td As Double, beta As Double
Dim res_conduit As debit_conduit
Q = eds.Qpluie / 1000#
tr = eds.tron_amo
 res_conduit = calc_debit_tr(tr, Q)
 td = res_conduit.hauteur
    If td / tr.conduit.Diametre < 1 Then
        beta = 2 * arccosinus(1 - 2 * td / tr.conduit.Diametre)
    Else
        beta = 2 * pi
    End If
'        s = (1 / 8) * (tr.conduit.Diametre ^ 2) * (beta - Sin(beta))
'        v = Qpl / s
        
      ' calcul ecoulement amont (torrentiel ou fluvial)
    ecoulam = calcul_ecoul(Q, tr.conduit.Diametre, beta)
If ecoulam = "TORREN." Then
    ok = True
Else
    ok = False
End If
tr = eds.tron_ava
res_conduit = calc_debit_tr(tr, eds.Qrin / 1000#)
If res_conduit.charge Then
    ok = False And ok
Else
    ok = True And ok
End If
 test_do_lat = ok
End Function
Public Sub calcul_do_lat(ByRef eds As st_dessdo, ByRef edv As deversoir)
Dim edr As deversoir_resultat
Dim tr As troncon
Dim td As Double, beta As Double
Dim res_conduit As debit_conduit
Dim pam As Double, pav As Double, Ham As Double, Hav As Double, HM As Double
Dim Tram As Double, pentedo As Double
Dim longdo As Double
Dim smes As String
Dim oksuite As Boolean
Dim Q As Double

Dim Qdev As Double
Dim Longueur As Double
Dim c As Double
Dim tam As Double
Dim tamqp As Double
Dim hautam As Double
Dim okboucle As Boolean
'verif hauteur amont a qrin
tr = edessdo.tron_amo
res_conduit = calc_debit_tr(tr, eds.Qrin / 1000#)
tam = res_conduit.hauteur
res_conduit = calc_debit_tr(tr, eds.Qpluie / 1000#)
tamqp = res_conduit.hauteur
hautam = edo.hauteur

'verif hauteur aval a qrin
tr = edessdo.tron_ava
res_conduit = calc_debit_tr(tr, eds.Qrin / 1000#)
td = res_conduit.hauteur
longdo = edv.Longueur
pentedo = edv.pente

'julienne
pentedo = (pav - hautam) / longdo

pav = td
pam = pav - longdo * pentedo
'controle hauteur > 0.25
oksuite = True
'smes = verif_ecoul_am_cr
'smes = verif_ecoul_av_cr()
If pam < 0.25 Then
    smes = "Hauteur de lame " & str(Round(pam, 3)) + "  < 0.25 m "
    smes = smes + Chr(13) + "diminuer la pente de la conduite aval"
    MsgBox smes, vbOKOnly, "Vérification hauteur de seuil"
    oksuite = False
End If
If pam < tam Then
    smes = "Hauteur de lame " & str(Round(pam, 3)) + "  < Tirant d'eau amont " & str(Round(tam, 3)) + "  m "
    smes = smes + Chr(13) + "diminuer la pente du do"
    MsgBox smes, vbOKOnly, "Vérification hauteur de seuil débit critique"
    oksuite = False
End If
If oksuite Then
' julienne
'Tram = edessdo.Tram
okboucle = True
Dim i As Integer
i = 0
While okboucle And i < 20
i = i + 1
pav = td
pam = pav - longdo * pentedo
pam = maximum(0.25, pam)

pam = maximum(hautam, pam)
'pam = minimum(hautam, pam)
pam = minimum(maximum(hautam, 0.25), pam)

    Tram = tamqp
    HM = (Tram - pam) / 4
    
Qdev = (edessdo.Qpluie - edessdo.Qrin) / 1000#
c = 1#
Longueur = 0.85 * Qdev / (c * HM ^ 1.5)
pentedo = maximum(0.0001, Round(((pav - pam) / Longueur), 4))

If Abs(longdo - Longueur) < 0.001 Then
  okboucle = False
End If
longdo = Longueur

Wend
edv.Longueur = Round(Longueur, 2)
edv.pente = pentedo
edv.Absamo = eds.tron_amo.Absava
edv.Absava = edv.Absamo + edv.Longueur
edv.radamo = eds.tron_amo.radava
edv.radava = edv.radamo - edv.Longueur * edv.pente
edv.hauteur = pam
edv.tron_ava = eds.tron_ava
edv.tron_ava.Absamo = edv.Absava
edv.tron_ava.radamo = edv.radava
edv.tron_ava.Absava = edv.tron_ava.Absamo + edv.tron_ava.conduit.Longueur
edv.tron_ava.radava = edv.tron_ava.radamo - edv.tron_ava.conduit.Longueur * edv.tron_ava.conduit.pente
eds.tron_ava = edv.tron_ava
 End If

End Sub

