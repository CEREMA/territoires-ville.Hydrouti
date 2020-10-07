Attribute VB_Name = "Fonct_do"
Global beta As Double
Global Const pi = 3.14159

Global ecoulavts As String
Global ecoulavcri As String
Global ecoulav10 As String
'Global dav As Single
Global vpsav As Double
Global qpsav As Double
'Global iradav As Single
'Global kav As Single
'Global vavts As Double
'Global havts As Double
'Global betavts As Double
Global vavcri As Double
Global havcri As Double
'Global betavcri As Double
Global vav10 As Double
Global hav10 As Double
'Global betav10 As Double

'déversoir
Global longdo As Double
Global hautdo As Double
Global pentedo As Double
Public Sub pre_dimdo(ByRef edo As deversoir)
Dim msg As String
Dim sresult As String
Dim v As Double
Dim s As Double
Dim vcri As Double
Dim tav As Double
Dim tcr As Double
Dim a As Double
Dim icri As Double
Dim dam As Double, kav As Double, dav As Double, iradav As Double
Dim ldav As Double, pdav As Double
Dim ddav As Double, dhautdo As Double, dlongdo As Double, dpentedo As Double
Dim q10 As Double, qcri As Double

' initialisation des donnees
'dam = edessdo.dam
'dav = edessdo.dav
'kav = edessdo.kav
'iradav = edessdo.iradav

dam = edessdo.tron_amo.conduit.Diametre
dav = edessdo.tron_ava.conduit.Diametre
kav = edessdo.tron_ava.conduit.rugosite
iradav = edessdo.tron_ava.conduit.pente

q10 = edessdo.Qpluie
' julienne 20030725
qcri = edessdo.Qrin ' julienne 20030725 + edessdo.Qts

longdo = edo.Longueur
hautdo = edo.hauteur
pentedo = edo.pente

    If longdo = 0 Then
        longdo = Int(40 * q10 / (1000 * dam)) / 10
    End If
'radier entrée Do
    
' longueur du do
' hauteur de la crete 0.25 m mini ou 0.6 DAM
    If hautdo = 0 Then
        hautdo = 0.6 * dam '/ 1000
        If hautdo < 0.25 Then
            hautdo = 0.25
        End If
    End If
    If pentedo = 0 Then
        pentedo = 0.01
    End If
'    hautdo = 0.7
edo.hauteur = hautdo
edo.Longueur = longdo
edo.pente = pentedo
If tav = 0 Then
        tav = hautdo + longdo * pentedo
End If
edo.tav = tav
End Sub
Public Function pre_calculdo(ByRef edo As deversoir, ByRef rescalcdo As Resudo) As String
Dim msg As String
Dim sresult As String
Dim v As Double
Dim s As Double
Dim vcri As Double
Dim tav As Double
Dim tcr As Double
Dim a As Double
Dim icri As Double
Dim dam As Double, kav As Double, dav As Double, iradav As Double
Dim ldav As Double, pdav As Double
Dim ddav As Double, dhautdo As Double, dlongdo As Double, dpentedo As Double
Dim q10 As Double, qcri As Double

' initialisation des donnees
'dam = edessdo.dam
'dav = edessdo.dav
'kav = edessdo.kav
'iradav = edessdo.iradav

dam = edessdo.tron_amo.conduit.Diametre
dav = edessdo.tron_ava.conduit.Diametre
kav = edessdo.tron_ava.conduit.rugosite
iradav = edessdo.tron_ava.conduit.pente

q10 = edessdo.Qpluie

qcri = edessdo.Qrin ' julienne 20030725 + edessdo.Qts
longdo = edo.Longueur
hautdo = edo.hauteur
pentedo = edo.pente

    If longdo = 0 Then
        longdo = Int(40 * q10 / (1000 * dam)) / 10
    End If
'radier entrée Do
    
' longueur du do
' hauteur de la crete 0.25 m mini ou 0.6 DAM
    If hautdo = 0 Then
        hautdo = 0.6 * dam '/ 1000
        If hautdo < 0.25 Then
            hautdo = 0.25
        End If
    End If
    If pentedo = 0 Then
        pentedo = 0.01
    End If
'    hautdo = 0.7
edo.hauteur = hautdo
edo.Longueur = longdo
edo.pente = pentedo

If tav = 0 Then
        tav = hautdo + longdo * pentedo
End If
edo.tav = tav
ldav = rech_ldav_do(edo)
    ldav = Round(ldav, 2)
'    Frm_do.Lb_longce.Caption = "Longueur conduite aval étranglée = " + Str(Round(ldav, 2)) + " m"
pdav = rech_pdav_do(edo)
'     Frm_do.Lb_longce.Caption = sresult
    edo.hauteur = hautdo
    edo.Longueur = longdo
    edo.pente = pentedo
    edo.tav = tav
    tav = edessdo.tron_ava.conduit.Longueur * (icri - iradav / 10000) + (a * (vcri ^ 2) / 19.62) + (dav / 1000) + tcr
    tav = tav - longdo * pentedo
    dhautdo = rech_ham_do(edo, ldav)
    dlongdo = longdo + (dhautdo - edo.hauteur) / pentedo
    dpentedo = (dhautdo - edo.hauteur) / longdo + pentedo
    edo.tron_ava = edessdo.tron_ava
    edo.Absamo = edessdo.tron_amo.Absava
    edo.radamo = edessdo.tron_amo.radava
    edo.Absava = edo.Absamo + edo.Longueur
    edo.radava = edo.radamo - edo.Longueur * edo.pente
    edo.tron_ava.conduit = edessdo.tron_ava.conduit
    edo.tron_ava.conduit.Longueur = ldav
    edo.tron_ava.Absamo = edo.Absava
    edo.tron_ava.radamo = edo.radava
    edo.tron_ava.Absava = edo.tron_ava.Absamo + edo.tron_ava.conduit.Longueur
    edo.tron_ava.radava = edo.tron_ava.radamo - edo.tron_ava.conduit.Longueur * edessdo.tron_ava.conduit.pente
    ddav = rech_dav_do(edo, ldav)
    
    sresult = " Conduite aval étranglée :" + Chr(13) + " Longueur = " + ajout_zero(Trim(Str(ldav))) + " m"
    resudev.longetranglee = ajout_zero(Trim(Str(ldav)))
    If pdav > 0 Then
        sresult = sresult + Chr(13) + Chr(10) + " Modifier la pente :" + ajout_zero(Trim(Str(Round(pdav * 10000, 0)))) + "  1/10000"
    End If
    sresult = sresult + Chr(13) + Chr(10) + " Modifier le diamètre :" + ajout_zero(Trim(Str(Round(ddav * 1000, 0)))) + "mm"
    sresult = sresult + Chr(13) + Chr(10) + Chr(13) + Chr(10) + " Deversoir :" + Chr(13) + Chr(10) + "Modifier la hauteur :" + ajout_zero(Trim(Str(Round(dhautdo * 1000, 0)))) + "  mm"
    sresult = sresult + Chr(13) + Chr(10) + " Modifier la longueur  :" + ajout_zero(Trim(Str(Round(dlongdo, 3)))) + "  m"
    sresult = sresult + Chr(13) + Chr(10) + " Modifier la pente  :" + ajout_zero(Trim(Str(Round(dpentedo * 10000, 0)))) + "  1/10000"
 With rescalcdo
    .ldav = ldav
    .pdav = pdav
    .ddav = ddav
    .dlongdo = dlongdo
    .dpentedo = dpentedo
 End With
 
 pre_calculdo = sresult
    
'    Frm_do.Lb_longce(0).Caption = sresult
'    Frm_do.SSTab_result.Tab = 0

End Function
Public Function rech_ldav_do(ByRef edo As deversoir)
Dim msg As String
Dim sresult As String
Dim v As Double
Dim s As Double
Dim vcri As Double
Dim tav As Double
Dim tcr As Double
Dim a As Double
Dim icri As Double
Dim dam As Double, kav As Double, dav As Double, iradav As Double
Dim ldav As Double, pdav As Double
Dim ddav As Double, dhautdo As Double, dlongdo As Double, dpentedo As Double
Dim q10 As Double, qcri As Double
'dam = edessdo.dam
'dav = edessdo.dav
'kav = edessdo.kav
'iradav = edessdo.iradav

dam = edessdo.tron_amo.conduit.Diametre
dav = edessdo.tron_ava.conduit.Diametre
kav = edessdo.tron_ava.conduit.rugosite
iradav = edessdo.tron_ava.conduit.pente

q10 = edessdo.Qpluie
qcri = edessdo.Qrin ' julienne 20030725 + edessdo.Qts
longdo = edo.Longueur
hautdo = edo.hauteur
pentedo = edo.pente

If tav = 0 Then
        tav = hautdo + longdo * pentedo
End If

If dav > 0 Then
If ldav = 0 Then
    vcri = (qcri / 1000) / ((3.14159 * (dav ^ 2)) / 4)
    icri = (vcri / (kav * (dav / 4) ^ (2 / 3))) ^ 2
    a = rech_do_A(tav, dav)
    
    ldav = (tav - (a * (vcri ^ 2) / 19.62) - dav - tcr) / (icri - iradav)

End If
End If
rech_ldav_do = ldav
End Function
Public Function rech_pdav_do(ByRef edo As deversoir)
Dim msg As String
Dim sresult As String
Dim v As Double
Dim s As Double
Dim vcri As Double
Dim tav As Double
Dim tcr As Double
Dim a As Double
Dim icri As Double
Dim dam As Double, kav As Double, dav As Double, iradav As Double
Dim ldav As Double, pdav As Double
Dim ddav As Double, dhautdo As Double, dlongdo As Double, dpentedo As Double
Dim q10 As Double, qcri As Double
'dam = edessdo.dam
'dav = edessdo.dav
'kav = edessdo.kav
'iradav = edessdo.iradav

dam = edessdo.tron_amo.conduit.Diametre
dav = edessdo.tron_ava.conduit.Diametre
kav = edessdo.tron_ava.conduit.rugosite
iradav = edessdo.tron_ava.conduit.pente


q10 = edessdo.Qpluie
qcri = edessdo.Qrin ' julienne 20030725 + edessdo.Qts

longdo = edo.Longueur
hautdo = edo.hauteur
pentedo = edo.pente

If tav = 0 Then
        tav = hautdo + longdo * pentedo
End If

If dav > 0 Then
If ldav = 0 Then
    vcri = (qcri / 1000) / ((3.14159 * (dav ^ 2)) / 4)
    icri = (vcri / (kav * (dav / 4) ^ (2 / 3))) ^ 2
    a = rech_do_A(tav, dav)
    ldav = (tav - (a * (vcri ^ 2) / 19.62) - dav - tcr) / (icri - iradav)
    pdav = icri - (tav - (a * (vcri ^ 2) / 19.62) - dav - tcr) / edessdo.tron_ava.conduit.Longueur

End If
End If
If pdav > 0 And pdav < icri Then
    rech_pdav_do = pdav
Else
rech_pdav_do = 0
End If
End Function

Function rech_dav_do1() As Double

End Function
Function verif_remous_do() As String
Dim havts As Double, hamts As Double, vavts As Double, vamts As Double
Dim canal As conduite
Dim ltc() As Variant, ct() As Variant, qvm(5) As Variant
Dim qcal As Double
Dim message As String
message = ""
verif_remous_do = message
qcal = edessdo.Qts
canal = edessdo.tron_ava.conduit
Call cana(canal, ct)
ltc = calc_par(canal)
qvi = caltran1(qcal, ct, ltc)
havts = qvi(5)
vavts = qvi(2)
'    Debug.Print "qvi : 1= débit"; qvi(1); " 2=vitesse"; qvi(2); "qvi"; " 3=acceleration"; qvi(3)
'    Debug.Print " 4=Largeur libre"; qvi(4); " 5 = hauteur; "; qvi(5); ""
canal = edessdo.tron_amo.conduit
Call cana(canal, ct)
ltc = calc_par(canal)
qvi = caltran1(qcal, ct, ltc)
hamts = qvi(5)
vamts = qvi(2)

' vérification remou aval-amont
'    If havts / 1000 + (vavts ^ 2 - vamts ^ 2) / 19.62 > hamts / 1000 + longdo * pentedo Then
'        erreur(1) = 1
'    End If
    If (havts + ((vavts ^ 2 - vamts ^ 2) / 19.62)) > (hamts + (edo.Longueur * edo.pente)) Then
 '       Debug.Print "remou"
        message = Chr(13) + Chr(10) + "Il y a remous aval-amont"
        verif_remous_do = message
    End If

End Function
Function verif_ecoul_am_cr() As String
Dim beta As Double
Dim dam As Double
Dim hautdo As Double
Dim qcri As Double
Dim s, v As Double
Dim message As String
message = ""
verif_ecoul_am_cr = message
hautdo = edo.hauteur
dam = edessdo.tron_amo.conduit.Diametre
qcri = edessdo.Qrin ' julienne 20030725 + edessdo.Qts
' verification vitesse d'écoulement amont pour qcri
    beta = 2 * arccosinus(1 - 2 * hautdo / dam)
    s = (1 / 8) * (dam ^ 2) * (beta - Sin(beta))
    v = Int(100 * qcri / s / 1000) / 100
    If v < 0.3 Then
'        Debug.Print " vitesse < 0.3 m/s"
        hautdo = rech_haut_do_vam(dam, qcri / 0.3)
        message = Chr(13) + Chr(10) + "Vitesse amont " + ajout_zero(Trim(Str(Round(v, 2)))) + " pour le débit de référence < 0.3 m/s"
        message = message + Chr(13) + Chr(10) + "Hauteur limite de la lame < à :" + ajout_zero(Trim(Str(Int(hautdo * 1000)))) + " mm"
        verif_ecoul_am_cr = message
    End If

End Function
Function verif_ecoul_av_cr() As String
Dim qv As deb_vit
Dim canal As conduite
Dim qcri As Double
Dim message As String
message = ""
verif_ecoul_av_cr = message
qcri = edessdo.Qrin ' julienne 20030725 + edessdo.Qts / 1000#
canal = edessdo.tron_ava.conduit

qv = debvit_ps(canal)

'vérification écoulement aval libre ou charge
    If qv.debit > qcri Then
        message = Chr(13) + Chr(10) + "L'écoulement aval est libre"
        verif_ecoul_av_cr = message
    End If
End Function
Function verif_Hauteur_am_cr() As String
Dim hamcri As Double, hautdo As Double
Dim canal As conduite
Dim ltc() As Variant, ct() As Variant, qvm(5) As Variant
Dim qcal As Double
Dim message As String
message = ""
verif_Hauteur_am_cr = message
hautdo = edo.hauteur
qcal = edessdo.Qrin + edessdo.Qts
canal = edessdo.tron_amo.conduit
Call cana(canal, ct)
ltc = calc_par(canal)
qvi = caltran1(qcal, ct, ltc)
hamcri = qvi(5)

' vérification hauteur de crete>tirant d'eau amont qcri
    If hautdo < hamcri Then
'        Debug.Print "Hauteur de crête < Hauteur de débit de pluie de rincage"
        message = Chr(13) + Chr(10) + "Hauteur de crête < Hauteur de débit de pluie de rincage"
        verif_Hauteur_am_cr = message
    End If
End Function

Function verif_Hauteur_cphe() As String
Dim message As String
message = ""
verif_Hauteur_cphe = message
' vérification hauteur de crete>cote des phe
  
'    If raddo + hautdo < phe Then
'        erreur(5) = 5
'    End If

If edessdo.phex > edo.radamo + edo.hauteur Then
'    Debug.Print "Cote de la lame < Cote des plus hautes eaux"
    message = Chr(13) + Chr(10) + "La cote de la lame " + Str$(edo.radamo + edo.hauteur) + " est inférieure  à la cote des PHE" + Str$(edessdo.phex)
    verif_Hauteur_cphe = message
End If
End Function
Function verif_dech() As Boolean
verif_dech = True
' vérification hauteur de crete>cote des phe
  
'    If raddo + hautdo < phe Then
'        erreur(5) = 5
'    End If

If edessdo.phex > edo.radamo + edo.hauteur Then
'    Debug.Print "Cote de la lame < Cote des plus hautes eaux"
    verif_dech = False
End If
End Function

Sub verif_fonct()
Dim s As Double
Dim v As Double
Static erreur(10) As Integer
Static errmsg(10) As String

For i = 0 To 5
    erreur(i) = 0
    VERIF.Etiquette1(i) = ""
    VERIF.Commande1(i).Visible = False
Next i
' vérification remou aval-amont
    If havts / 1000 + (vavts ^ 2 - vamts ^ 2) / 19.62 > hamts / 1000 + longdo * pentedo Then
        erreur(1) = 1
    End If
' verification vitesse d'écoulement amont pour qcri
    beta = 2 * arccosinus(1 - 2 * hautdo / (dam / 1000))
    s = (1 / 8) * ((dam / 1000) ^ 2) * (beta - Sin(beta))
    v = Int(100 * qcri / s / 1000) / 100
    If v < 0.3 Then
        erreur(2) = 2
    End If
'vérification écoulement aval libre ou charge
    If qpsav > qcri Then
        erreur(3) = 3
    End If
' vérification hauteur de crete>tirant d'eau amont qcri
    If hautdo < hamcri / 1000 Then
        erreur(4) = 4
    End If
' vérification hauteur de crete>cote des phe
  
    If raddo + hautdo < phe Then
        erreur(5) = 5
    End If

    errmsg(1) = "Il y a remous aval-amont"
    errmsg(2) = "La vitesse amont pour le débit de référence : " + Str$(v) + " m/s"
    errmsg(3) = "L'ecoulement aval est libre"
    errmsg(4) = "Crête trop basse! Valeur mini = " + Str$(hamcri / 1000) + " m"
    errmsg(5) = "La cote de la crête " + Str$(raddo + hautdo) + " est inférieure  à la cote des PHE" + Str$(phe)
j = -1
For i = 1 To 5

    If erreur(i) <> 0 Then
            j = j + 1
            VERIF.Etiquette1(j) = errmsg(i)
            VERIF.Commande1(j).Visible = True
            VERIF.Commande1(j).Caption = "Aide"
    End If
    Next i
End Sub


Sub affi_condam()
Dim vts As Integer
Dim vcri As Double
Dim v10 As Double
Dim hts As Double
Dim hcri As Double
Dim h10 As Double
Dim ecoults As String
Dim ecoulcri As String
Dim ecoul10 As String
'affichage caractéristique conduite amont
For i = 6 To 11
    FeuilleDO.Textebte(i).Visible = False

Next i
If cadre = 2 Then
    FeuilleDO.Label(0).Caption = "Diamètre en mm"
    FeuilleDO.Label(1).Caption = "Pente d en 1/10000"
    FeuilleDO.Label(2).Caption = "Coefficient de Manning-Strickler"
    FeuilleDO.Label(3).Caption = "Longueur de la canalisation en m"
    FeuilleDO.Label(4).Caption = "Vitesse à pleine section en m/s"
    FeuilleDO.Label(5).Caption = "Débit à pleine section en l/s"
    FeuilleDO.Textebte(0).Text = dam
    FeuilleDO.Textebte(1).Text = iRadam
    FeuilleDO.Textebte(2).Text = kamon
    FeuilleDO.Textebte(3).Text = ldam

    FeuilleDO.Label(6).Caption = "Vitesse pleine section en m/s"
    FeuilleDO.Label2(6).Caption = Str$(vpsam)
    FeuilleDO.Label(7).Caption = "Débit pleine section en m3/s"
    FeuilleDO.Label2(7).Caption = Str$(qpsam)

    FeuilleDO.Label(8).Caption = "Vitesse d'écoulement à Qts en m/s"
    FeuilleDO.Label2(8).Caption = Str$(vamts) + " m/s"
    FeuilleDO.Label(9).Caption = "Hauteur d'eau Qts en m"
    FeuilleDO.Label2(9).Caption = Str$(hamts / 1000)
    FeuilleDO.Label(10).Caption = "Vitesse d'écoulement à Qcri en m/s"
    FeuilleDO.Label2(10).Caption = Str$(vamcri)
    FeuilleDO.Label(11).Caption = "Hauteur d'eau Qcri en m"
    FeuilleDO.Label2(11).Caption = Str$(hamcri / 1000)
    FeuilleDO.Label(12).Caption = "Vitesse d'écoulement à Qdix en m/s"
    FeuilleDO.Label2(12).Caption = Str$(vam10)
    FeuilleDO.Label(13).Caption = "Hauteur d'eau Qdix en m"
    FeuilleDO.Label2(13).Caption = Str$(ham10 / 1000)

    vts = vamts: vcri = vamcri: v10 = vam10
    hts = hamts: hcri = hamts: h10 = ham10
    ecoults = ecoulamts: ecoulcri = ecoulamcri: ecoul10 = ecoulam10
    Else
    FeuilleDO.Label(0).Caption = "Diamètre en mm"
    FeuilleDO.Label(1).Caption = "Pente d en 1/10000"
    FeuilleDO.Label(2).Caption = "Coefficient de Manning-Strickler"
    FeuilleDO.Label(3).Caption = "Longueur de la canalisation en m"
    FeuilleDO.Label(4).Caption = "Vitesse à pleine section en m/s"
    FeuilleDO.Label(5).Caption = "Débit à pleine section en l/s"
    FeuilleDO.Textebte(0).Text = dav
    FeuilleDO.Textebte(1).Text = iradav
    FeuilleDO.Textebte(2).Text = kav
    FeuilleDO.Textebte(3).Text = ldav

    FeuilleDO.Label(6).Caption = "Vitesse pleine section en m/s"
    FeuilleDO.Label2(6).Caption = Str$(vpsav)
    FeuilleDO.Label(7).Caption = "Débit pleine section en m3/s"
    FeuilleDO.Label2(7).Caption = Str$(qpsav)

    FeuilleDO.Label(8).Caption = "Vitesse d'écoulement à Qts en m/s"
    FeuilleDO.Label2(8).Caption = Str$(vavts) + " m/s"
    FeuilleDO.Label(9).Caption = "Hauteur d'eau Qts en m"
    FeuilleDO.Label2(9).Caption = Str$(havts / 1000)
    FeuilleDO.Label(10).Caption = "Vitesse d'écoulement à Qcri en m/s"
    FeuilleDO.Label2(10).Caption = Str$(vavcri)
    FeuilleDO.Label(11).Caption = "Hauteur d'eau Qcri en m"
    FeuilleDO.Label2(11).Caption = Str$(havcri / 1000)
    FeuilleDO.Label(12).Caption = "Vitesse d'écoulement à Qdix en m/s"
    FeuilleDO.Label2(12).Caption = Str$(vav10)
    FeuilleDO.Label(13).Caption = "Hauteur d'eau Qdix en m"
    FeuilleDO.Label2(13).Caption = Str$(hav10 / 1000)


    vts = vavts: vcri = vavcri: v10 = vav10
    hts = havts: hcri = havcri: h10 = hav10
    ecoults = ecoulavts: ecoulcri = ecoulavcri: ecoul10 = ecoulav10
End If
End Sub
Function fnvit(Q, r, b)
Dim s As Double
Dim s1 As Double
Dim SS As Double
s = (r ^ 2) * b / 2
s1 = (r ^ 2) * Sin(b)
SS = s - s1
'Debug.Print (Q / SS)
    fnvit = 2 * Q / r ^ 2 / (b - Sin(b))
'    fnvit = q / R ^ 2 / (b / 2 - Sin(b))
    hauteur = r * (1 - Cos(b / 2))
End Function
Sub calcul_condam(ByRef cana As conduite)
''Dim ecoulamts As String, ecoulamcri As String, ecoulam10 As String
''Dim vamcri As Double, hamcri As Double, betamcri As Double, vam10 As Double
''Dim ham10 As Double, betam10 As Double
''Dim dam As Single, iRadam As Single, kamon As Single
''Dim vpsam As Double, qpsam As Double, vamts As Double, hamts As Double, betamts As Double
''Dim Qts As Double, qcri As Double, q10 As Double
''Qts = txtVersNum(Frm_do.Tb_Qts.Text)
''qcri = txtVersNum(Frm_do.Tb_Qrin.Text)
''q10 = txtVersNum(Frm_do.Tb_Qpluie.Text)
''
''
'' dam = cana.Diametre * 1000
'' iRadam = cana.pente * 10000
'' kamon = cana.rugosite
'' vpsam = kamon * ((dam / 4000) ^ (2 / 3)) * ((iRadam / 10000) ^ 0.5)
''    qpsam = (1000 * vpsam * (PI * (dam / 1000) ^ 2) / 4#)
''
''    If Qts > 0 And qpsam > 0 Then
''        If Qts < qpsam Then
''        betamts = angle(Qts / qpsam)
''        betamts = beta
''        vamts = Int(100 * fnvit(Qts / 1000, dam / 2000, betamts)) / 100
''        hamts = Int(1000 * hauteur)
''        ecoulamts = calcul_ecoul(Qts / 1000, dam / 1000, betamts)
''        Else
''         ecoulamts = "Charge"
''        End If
''
''    End If
''    If qcri > 0 And qpsam > 0 Then
''        betamcri = angle(qcri / qpsam)
''        betamcri = beta
''        vamcri = Int(100 * fnvit(qcri / 1000, dam / 2000, betamcri)) / 100
''        hamcri = Int(1000 * hauteur)
''        ecoulamcri = calcul_ecoul(qcri / 1000, dam / 1000, betamcri)
''    End If
''    If q10 > 0 And qpsam > 0 Then
''        betam10 = angle(q10 / qpsam)
''        betam10 = beta
''        vam10 = Int(100 * fnvit(q10 / 1000, dam / 2000, betam10)) / 100
''        ham10 = Int(1000 * hauteur)
''        ecoulam10 = calcul_ecoul(q10 / 1000, dam / 1000, betam10)
''    End If
''
'''conditions aval
''
'''    vpsav = kav * ((dav / 4000) ^ (2 / 3)) * ((iradav / 10000) ^ 0.5)
'''    qpsav = (Int(1000 * vpsav * (PI * (dav / 1000) ^ 2) / 4))
'''    If qts > 0 And qpsav > 0 Then
'''        betavts = Angle(qts / qpsav)
'''        betavts = beta
'''        vavts = Int(100 * fnvit(qts / 1000, dav / 2000, betavts)) / 100
'''        havts = Int(1000 * hauteur)
'''        ecoulavts = calcul_ecoul(qts / 1000, dav / 1000, betavts)
'''    End If
'''    If qcri > 0 And qpsav > 0 Then
'''        betavcri = Angle(qcri / qpsav)
'''        betavcri = beta
'''        vavcri = Int(100 * fnvit(qcri / 1000, dav / 2000, betavcri)) / 100
'''        havcri = Int(1000 * hauteur)
'''        ecoulavcri = calcul_ecoul(qcri / 1000, dav / 1000, betavcri)
'''        ecoulavcri = ecoul
'''    End If
'''    If Q10 > 0 And qpsav > 0 Then
'''        betav10 = Angle(Q10 / qpsav)
'''        betav10 = beta
'''        vav10 = Int(100 * fnvit(Q10 / 1000, dav / 2000, betav10)) / 100
'''        hav10 = Int(1000 * hauteur)
'''        ecoulav10 = calcul_ecoul(Q10 / 1000, dav / 1000, betav10)
'''        ecoulav10 = ecoul
'''    End If
''
End Sub
Function calcul_ecoul(Q, d, b)
Dim c1 As Double
Dim c2 As Double

    c1 = ((16 * Q) / (9.81 ^ 0.5) / (d ^ (5 / 2))) ^ 2
    c2 = (((b - Sin(b)) ^ 3) * Sin(b / 2)) / (1 - Cos(b))
    If c1 < 0.75 * c2 Then
        calcul_ecoul = "FLUVIAL"
        Else
        calcul_ecoul = "TORREN."
    End If
End Function
Function angle(rq) As Double
Dim c1 As Double
Dim pas As Double
Dim xi As Integer
Dim sinbeta As Double
Dim lambda  As Double
'Dim beta As Double
 '      CALCUL DE BETA
beta = 0
If rq > 1 Then
 beta = 2 * pi
Else
If rq > 0 Then
          c1 = (2 * pi * rq) ^ 1.5
            beta = pi
            pas = pi / 2
            For xi = 1 To 15
            sinbeta = Sin(beta)
                If beta > pi Then
'                    lambda = (7 / 12) * sinbeta + (5 / 24) * Sin(2 * beta)
                    lambda = 0
                Else
                    lambda = 0
                End If
            If (((beta - sinbeta) ^ 2.5) / (beta - lambda)) > c1 Then
'julienne           If (((2 * (beta / 2 - sinbeta)) ^ 2.5) / (beta - lambda)) > c1 Then
                beta = beta - pas
            Else
                beta = beta + pas
            End If
            pas = pas / 2
            Next xi
End If
End If
angle = beta
End Function
Function rech_haut_do_vam(ByVal diam As Double, ByVal rq As Double) As Double
Dim c1 As Double, c2 As Double
Dim pas As Double
Dim xi As Integer
Dim sinbeta As Double
Dim lambda  As Double
Dim ray As Double
Dim haut As Double
ray = diam / 2#
 '      CALCUL DE BETA
 c1 = rq / 1000 / ray ^ 2
       
            beta = pi / 2
            pas = pi / 4
            c2 = (beta - sinbeta / 2)
            xi = 1
            While (Abs(c1 - c2) > 0.000001 Or c1 - c2 < 0) And xi < 50
                sinbeta = Sin(2 * beta)
                c2 = (beta - sinbeta / 2)
                If c2 > c1 Then
                    beta = beta - pas
                Else
                    beta = beta + pas
                End If
                pas = pas / 2
                xi = xi + 1
            Wend
'surface=  ray ^ 2 * (beta - sinbeta / 2)
haut = ray - ray * Cos(beta)
rech_haut_do_vam = haut
End Function

