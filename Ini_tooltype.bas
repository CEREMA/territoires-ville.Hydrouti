Attribute VB_Name = "Ini_tooltip"
Function ini_tooltip_bv(frm1 As Form)
'bassin versant
'Caractéristiques eau pluviale
frm1.Tb_car_ep(0).ToolTipText = "Surface totale du B.V."
frm1.Tb_car_ep(1).ToolTipText = "Longueur du plus long parcours hydraulique"
frm1.Tb_car_ep(2).ToolTipText = "Pente du plus long parcours hydraulique"
frm1.Tb_car_ep(3).ToolTipText = "Coefficient de ruissellement appliqué au B.V."
'Caractéristiques eau usées
frm1.Tb_car_eu(0).ToolTipText = "Nombre d'équivalent-habitants sur le B.V."
frm1.Tb_car_eu(1).ToolTipText = "Rejet moyen journalier d'eaux usées par équivalent-habitant"
frm1.Tb_car_eu(2).ToolTipText = "% d'eaux claires parasites par rapport au débit moyen des eaux usées"
'Caractéristiques caractéristiques
frm1.Tb_carep_rur(0).ToolTipText = "Hauteur de la lame d'eau absorbée par les pertes initiales"
frm1.Tb_carep_rur(1).ToolTipText = "Variable 'fc' de la loi de Horton "
frm1.Tb_carep_rur(2).ToolTipText = "Paramètre 'a' de la loi de Horton "
frm1.Tb_carep_rur(3).ToolTipText = "Paramètre 'b' de la loi de Horton "
frm1.Tb_carep_rur(4).ToolTipText = ""
'Paramètres eau pluviale
frm1.Tb_par_ep(0).ToolTipText = "Valeur du paramètre 'a' pour les averses de durée inférieure au seuil"
frm1.Tb_par_ep(1).ToolTipText = "Valeur du paramètre 'b' pour les averses de durée inférieure au seuil"
frm1.Tb_par_ep(2).ToolTipText = "Valeur du paramètre 'a' pour les averses de durée supérieure au seuil"
frm1.Tb_par_ep(3).ToolTipText = "Valeur du paramètre 'b' pour les averses de durée supérieure au seuil"
frm1.Tb_par_ep(4).ToolTipText = "Seuil d'exploitation statistiques de durée des averses"
'Paramètres eau usée
frm1.Tb_par_eu(0).ToolTipText = "Intensité de la pluie de référence"
frm1.Tb_par_eu(1).ToolTipText = "Valeur du paramètre 'a' pour évaluation du coefficient de pointe (1,5 par défaut)"
frm1.Tb_par_eu(2).ToolTipText = "Valeur du paramètre 'b' pour évaluation du coefficient de pointe (2,5 par défaut)"""
'Paramètres pluie de projet
frm1.Tb_par_pl(0).ToolTipText = "Durée totale de la pluie"
frm1.Tb_par_pl(1).ToolTipText = "Durée de la période intense"
frm1.Tb_par_pl(2).ToolTipText = "Hauteur totale précipitée - Double clic pour calcul avec loi de Montana"
frm1.Tb_par_pl(3).ToolTipText = "Hauteur précipitée pendant la période intense - Double clic pour calcul avec loi de Montana"
frm1.Tb_par_pl(4).ToolTipText = "Décalage de l'intant de la pointe- Par défaut 0.5(pluie centrée)"
frm1.Tb_par_pl(5).ToolTipText = "Pas de temps de discrétisation"
'Débit pluie d'orage
frm1.Tb_Debit(0).ToolTipText = "Débit calculé par la méthode de Caquot"
frm1.Tb_Debit(1).ToolTipText = "Débit calculé par la mèthode rationnelle"
frm1.Tb_Debit(2).ToolTipText = "Débit de pointe du calcul par la méthode de l'hydrogramme"
'Débit des eaux usées
frm1.Tb_debit1(0).ToolTipText = "Débit de pointe des eaux usées : Qeu"
'Débit de temps sec
frm1.Tb_debit1(1).ToolTipText = "Débit de pointe de temps sec : Qts = Qes + Qecp"
'Débit des eaux claires
frm1.Tb_debit1(2).ToolTipText = "Débit des eaux claires parasites : Qecp"
'Débit de référence
frm1.Tb_debit1(3).ToolTipText = "Débit de référence : Qref = Qpref + Qts"
'Débit de pluie de référence
frm1.Tb_debit1(4).ToolTipText = "Débit de la pluie de référence : Qpref"
'Débit d' orage
frm1.Tb_debit1(5).ToolTipText = "Débit de la pluie d'orage : Qora (choix Caquot, Rationnelle ou Hydrogramme)"
'Volume total ruisselé
frm1.Tb_debit1(6).ToolTipText = "Volume ruisselée sur le BV (méthode de l'hydrogramme)"
End Function
Function ini_tooltip_chute(frm1 As Form)
'chute
'Conduite amont
frm1.Tb_amo(0).ToolTipText = "Diamètre de la canalisation en amont de la chute"
frm1.Tb_amo(1).ToolTipText = "Pente de la canalisation en amont de la chute"
frm1.Tb_amo(2).ToolTipText = "Coefficient de Manning-Strickler de la canalisation en amont de la chute"
frm1.Tb_amo(3).ToolTipText = "Cote du fil d'eau d'arrivée de la canalisation amont"
'Conduite aval
frm1.Tb_ava(0).ToolTipText = "Diamètre de la canalisation en val de la chute"
frm1.Tb_ava(1).ToolTipText = "Pente de la canalisation en aval de la chute"
frm1.Tb_ava(2).ToolTipText = "Coefficient de Manning-Strickler de la canalisation en aval de la chute"
frm1.Tb_ava(3).ToolTipText = "Cote du fil d'eau de départ de la canalisation aval"
'Débit
frm1.Tb_Qmax.ToolTipText = "Débit maximal à transiter"
End Function
Function ini_tooltip_pompe(frm1 As Form)
'pompe
'Débits Caract.
frm1.Tb_Debit(0).ToolTipText = "Valeur du débit moyen journalier des eaux usées (Qeum)"
frm1.Tb_Debitc(0).ToolTipText = "Valeur du débit moyen journalier des eaux usées (Qeum)"
frm1.Tb_Debit(1).ToolTipText = "Valeur du débit de pointe des eaux usées (Qeu=Qeum x p)"
frm1.Tb_Debitc(1).ToolTipText = "Valeur du débit de pointe des eaux usées (Qeu=Qeum x p)"
frm1.Tb_Debit(2).ToolTipText = "Valeur du débit des eaux claires parasites (Qecp)"""
frm1.Tb_Debitc(2).ToolTipText = "Valeur du débit des eaux claires parasites (Qecp)"""
frm1.Tb_Debit(3).ToolTipText = "Valeur du débit  moyen de temps sec (Qmts=Qeum+Qecp)"
frm1.Tb_Debitc(3).ToolTipText = "Valeur du débit  moyen de temps sec (Qmts=Qeum+Qecp)"
frm1.Tb_Debit(4).ToolTipText = "Valeur du débit de pointe de temps sec (Qts=Qeu+Qecp)"
frm1.Tb_Debitc(4).ToolTipText = "Valeur du débit de pointe de temps sec (Qts=Qeu+Qecp)"
frm1.Tb_FPointe.ToolTipText = "Facteur de pointe des eaux usées (p)"
frm1.Tb_Qpomp(0).ToolTipText = "Débit de pompage théorique (QpompThé=3 x Qeum +1)"
frm1.Tb_Qpompc(0).ToolTipText = "Débit de pompage théorique (QpompThé=3 x Qeum +1)"
'Données géométr.
'***Conduite de refoulement
frm1.Tb_Geom(0).ToolTipText = "Longueur développée de la canalisation de refoulement"
frm1.Tb_Geom(1).ToolTipText = "Diamètre théorique de la canalisation pour une vitesse d'écoulement de1.5 m/s"
frm1.Cb_Materiau.ToolTipText = "Matériau choisi pour la canalisation de refoulement"
'***Niveaux
frm1.Tb_Geom(3).ToolTipText = "Cote du terrain naturel au droit du poste de pompage"
frm1.Tb_Geom(4).ToolTipText = "Cote du fil d'eau de la canalisation d'arrivée dans le poste"
frm1.Tb_Geom(5).ToolTipText = "Cote du fil d'eau de départ de la canalisation de refoulement"
frm1.Tb_Geom(6).ToolTipText = "Cote du fil d'eau à l'extrémité du refoulement"
'Points singul.
frm1.Tb_PtSing(0).ToolTipText = "Nbre de coude(s) à 11°15 prévu(s) sur le refoulement"
frm1.Tb_PtSing(1).ToolTipText = "Nbre de coude(s) à 22°30 prévu(s) sur le refoulement "
frm1.Tb_PtSing(2).ToolTipText = "Nbre de coude(s) à 30° prévu(s) sur le refoulement "
frm1.Tb_PtSing(3).ToolTipText = "Nbre de coude(s) à 45° prévu(s) sur le refoulement "
frm1.Tb_PtSing(4).ToolTipText = "Nbre de coude(s) à 90° prévu(s) sur le refoulement "
frm1.Tb_PtSing(5).ToolTipText = "Nbre de vanne(s) prévue(s) sur le refoulement "
frm1.Tb_PtSing(6).ToolTipText = "Nbre de clapet(s) anti-retour prévu(s) sur le refoulement "
frm1.Tb_PtSing(7).ToolTipText = "Nbre de système(s) de vidange prévu(s) sur le refoulement "
frm1.Tb_PtSing(8).ToolTipText = "Nbre de ventouse(s) prévue(s) sur le refoulement "
frm1.Opt_PtSing(0).ToolTipText = "Mise en place d'un système de protection Anti-Bélier (OUI/NON)"
frm1.Opt_PtSing(1).ToolTipText = "Mise en place d'un système de protection Anti-Bélier (OUI/NON)"
'Données tech.
frm1.Tb_Nbpom.ToolTipText = "Nb de pompe(s) installées (en général 2)"
frm1.Tb_Ntdph.ToolTipText = "Nb de démarrage(s) prévus par heure (2 à 6 suivant la puissance des pompes)"
frm1.Tb_Vutba.ToolTipText = "Volume utile théorique de la bâche de pompage"
'***Section de la bâche
frm1.Opt_sect_ba(0).ToolTipText = "Choix pour une bâche de section circulaire"
frm1.Opt_sect_ba(1).ToolTipText = "Choix pour une bâche de section rectangulaire"
frm1.Tb_long.ToolTipText = "Longueur de la bâche rectangulaire"
frm1.Tb_larg.ToolTipText = "Largeur de la bâche rectangulaire"
frm1.Tb_diam.ToolTipText = "Diamètre de la bâche circulaire"
'
frm1.Tb_denivt.ToolTipText = "Tranche de pompage théorique"
frm1.Tb_denivhau.ToolTipText = "Garde à l'égout : distance entre le niveau de l'arrivée des eaux et le niveau de démarrage"
frm1.Tb_denivbas.ToolTipText = "Garde au fond"
'Résultats
frm1.Tb_Qpomp(1).ToolTipText = "Débit de pompage retenu"
frm1.Tb_Qpompc(1).ToolTipText = "Débit de pompage retenu m3"
frm1.Tb_Drflt.ToolTipText = "Diamètre intérieur de la canalisation de refoulement"
frm1.Tb_VitRflt.ToolTipText = "Vitesse instantanée de l'écoulement en régime  permanent (conseillée entre 0.8 et 1.2 m/s)"
frm1.Tb_Jmpkm.ToolTipText = "Pertes de charge dues au linéaire de canalisation"
frm1.Tb_denivr.ToolTipText = "Dénivelée entre le capteur de démarrage et le capteur d'arrêt des pompes"
frm1.Tb_vurba.ToolTipText = "Volume utile de la bâche"
frm1.Tb_nrdph.ToolTipText = "Nombre réel de démarrage(s) par heure "
frm1.Tb_Tvidange.ToolTipText = "Temps de vidange du volume utile"
frm1.Tb_T1cyc.ToolTipText = "Durée totale d'un cycle (remplissage + vidange)"
frm1.Tb_Nbcyc.ToolTipText = "Nombre de cycles par heure"
frm1.Tb_Vmy.ToolTipText = "Vitesse moyenne d'écoulement"
frm1.Tb_Tsejh.ToolTipText = "Temps de séjour"
frm1.Tb_Singul.ToolTipText = "Pertes de charge dues aux singularités (coudes, vannes,etc...)"
frm1.Tb_Hmt.ToolTipText = "Hauteur manométrique totale"

End Function

Function ini_tooltip_conduite(frm1 As Form)
'Conduite
frm1.Tb_cond(0).ToolTipText = "Diamètre de la canalisation"
frm1.Tb_cond(1).ToolTipText = "Pente de la canalisation "
frm1.Tb_cond(2).ToolTipText = "Coefficient de Manning-Strickler de la canalisation"
'Débit
frm1.Tb_Qmax.ToolTipText = "Débit maximal à transiter"
End Function
Function ini_tooltip_decant(frm1 As Form)
'bassin de décantation
'Paramètres
frm1.Tb_dec(0).ToolTipText = "Débit entrant dans l'ouvrage"
frm1.Tb_dec(1).ToolTipText = "Taille des particules à décanter" + " (domaine de validité 0,125-0,315)"
frm1.Tb_dec(2).ToolTipText = "Rapport décrivant la section transversale de l'ouvrage"
frm1.Tb_dec(3).ToolTipText = "% souhaité de décantation des particules de tailles retenues" + " (domaine de validité 85-100)"
frm1.Tb_dec(4).ToolTipText = "Vitesse horizontale des particules"
End Function
Function ini_tooltip_do(frm1 As Form)
'déversoir d'orage
' Bassin versant / hydraulique du BV
frm1.Tb_Debit(0).ToolTipText = "Débit d'orage des eaux pluviales"
frm1.Tb_Debit(1).ToolTipText = "Débit de temps sec"
frm1.Tb_Debit(2).ToolTipText = "Débit de référence"
'Contraintes
frm1.Tb_cont(0).ToolTipText = "Côte imposée de départ en amont du système"
frm1.Tb_cont(1).ToolTipText = "Côte imposée d'arrivée à l'aval du système"
frm1.Tb_cont(2).ToolTipText = "Longueur disponible entre les contraintes amont et aval"
frm1.Tb_cont(3).ToolTipText = "Côte des plus hautes eaux à l'exutoire de la canalisation de décharge"
frm1.Tb_cont(4).ToolTipText = "Côte du fil d'eau à l'exutoire de la canalisation de décharge"
frm1.Tb_cont(5).ToolTipText = "Longueur de la canalisation de décharge"
'Canal. amont / Conduite
frm1.Tb_amo(0).ToolTipText = "Diamètre de la canalisation de tranquilisation en amont de la chambre de déversement"
frm1.Tb_amo(1).ToolTipText = "Pente de la canalisation de tranquilisation "
frm1.Tb_amo(2).ToolTipText = "Coefficient de Manning-Strickler de la canalisation"
frm1.Tb_amo(3).ToolTipText = "Longueur de la canalisation de tranquilisation (minimum L = 20 x D)"
'Canal. aval / Conduite
frm1.Tb_ava(0).ToolTipText = "Diamètre de la canalisation étranglée en aval de la chambre de déversement"
frm1.Tb_ava(1).ToolTipText = "Pente de la canalisation étranglée"
frm1.Tb_ava(2).ToolTipText = "Coefficient de Manning-Strickler de la canalisation"
frm1.Tb_ava(3).ToolTipText = "Longueur de la canalisation étranglée (50 à 60 m maximum)"
'Déversoir / Caractéristiques
frm1.Tb_dev(0).ToolTipText = "Longueur de la lame de déversement "
frm1.Tb_dev(1).ToolTipText = "Hauteur de la crête (Mini 0.20 m / Conseillée 0.6 x D amont)"
frm1.Tb_dev(2).ToolTipText = "Pente longitudinale du déversoir"
frm1.Tb_dev(3).ToolTipText = "Valeur maximale du niveau de la ligne d'eau tolérée "
'Décharge / Conduite
frm1.Tb_dech(0).ToolTipText = "Diamètre de la canalisation de décharge"
frm1.Tb_dech(1).ToolTipText = "Pente de la canalisation de décharge"
frm1.Tb_dech(2).ToolTipText = "Coefficient de Manning-Strickler de la canalisation de décharge"
frm1.Tb_dech(3).ToolTipText = ""
'Décharge / Conduite / Coefficient de perte à l'entrée
End Function
Function ini_tooltip_door(frm1 As Form)
'Déversoir d'orage à ouverture de radier
' Bassin versant
frm1.Tb_Debit(0).ToolTipText = "Débit d'orage des eaux pluviales"
frm1.Tb_Debit(1).ToolTipText = "Débit de temps sec"
frm1.Tb_Debit(2).ToolTipText = "Débit de référence"
'Conduite arrivée
frm1.Tb_amo(0).ToolTipText = "Diamètre "
frm1.Tb_amo(1).ToolTipText = "Pente "
frm1.Tb_amo(2).ToolTipText = "Coefficient de Manning-Strickler "
frm1.Tb_amo(3).ToolTipText = "Longueur "
'Conduite départ
frm1.Tb_ava(0).ToolTipText = "Diamètre "
frm1.Tb_ava(1).ToolTipText = "Pente "
frm1.Tb_ava(2).ToolTipText = "Coefficient de Manning-Strickler "
frm1.Tb_ava(3).ToolTipText = "Longueur "
'Conduite déversement
frm1.Tb_dech(0).ToolTipText = "Diamètre "
frm1.Tb_dech(1).ToolTipText = "Pente "
frm1.Tb_dech(2).ToolTipText = "Coefficient de Manning-Strickler"
frm1.Tb_dech(3).ToolTipText = "Longueur "
'Contraintes
frm1.Tb_cont(1).ToolTipText = "cote radier aval,conduite arrivée "
frm1.Tb_hmin.ToolTipText = "Hauteur entre canalisations "
'Déversoir
frm1.Tb_dev(0).ToolTipText = "Longueur"
frm1.Tb_dev(1).ToolTipText = "Profondeur "
frm1.Chk_cri.ToolTipText = "Débit critique"
frm1.Chk_max.ToolTipText = "Débit d'orage "
frm1.Cmd_recalc.ToolTipText = "Dessin recalculé"
frm1.Cmd_mini.ToolTipText = "Dessin mini "
End Function
Function ini_tooltip_ret(frm1 As Form)
'bassin de rétention
'Bassin versant
frm1.Tb_bv(0).ToolTipText = "Surface totale du bassin versant"
frm1.Tb_bv(1).ToolTipText = "Coefficient d'apport (Attention différent du coefficient de ruissellement)"
'Paramètres pluviométriques
frm1.Tb_par(0).ToolTipText = "Valeur du paramètre 'a' pour les averses de durée inférieure au seuil"
frm1.Tb_par(1).ToolTipText = "Valeur du paramètre 'b' pour les averses de durée inférieure au seuil"
frm1.Tb_par(2).ToolTipText = "Valeur du paramètre 'a' pour les averses de durée supérieure au seuil"
frm1.Tb_par(3).ToolTipText = "Valeur du paramètre 'b' pour les averses de durée supérieure au seuil"
frm1.Tb_par(4).ToolTipText = "Seuil de durée d'exploitation statistiques des averses"
'Débit de fuite de la retenue
frm1.Tb_Qf.ToolTipText = "Débit de vidange de l'ouvrage de rétention"
'Schéma
frm1.Tb_long.ToolTipText = ""
frm1.Tb_larg.ToolTipText = ""
frm1.Tb_prof.ToolTipText = ""
frm1.Tb_rap.ToolTipText = ""
End Function
Function ini_tooltip_siphon(frm1 As Form)
'siphon
'Conduite amont
frm1.Tb_amo(0).ToolTipText = "Diamètre de la canalisation en amont du siphon"
frm1.Tb_amo(1).ToolTipText = "Pente de la canalisation en amont du siphon"
frm1.Tb_amo(2).ToolTipText = "Coefficient de Manning-Strickler de la canalisation en amont du siphon"
frm1.Tb_amo(3).ToolTipText = "Cote du fil d'eau d'arrivée de la canalisation en amont du siphon"
'Conduite aval
frm1.Tb_ava(0).ToolTipText = "Diamètre de la canalisation en aval du siphon"
frm1.Tb_ava(1).ToolTipText = "Pente de la canalisation en aval du siphon"
frm1.Tb_ava(2).ToolTipText = "Coefficient de Manning-Strickler de la canalisation en aval du siphon"
frm1.Tb_ava(3).ToolTipText = "Cote du fil d'eau de départ de la canalisation à l'aval du siphon"
'Siphon
frm1.Tb_siph(0).ToolTipText = "Diamètre de la canalisation du siphon"
frm1.Tb_siph(1).ToolTipText = "Coefficient de Manning-Strickler de la canalisation du siphon"
frm1.Tb_siph(2).ToolTipText = "Débit à transiter"
frm1.Tb_siph(3).ToolTipText = "Longueur développée du siphon"
'Siphon / Coefficient singulatités
frm1.Tb_siph(4).ToolTipText = ""
End Function
Function ini_tooltip_stock(frm1 As Form)
'bassin de stockage
'Bassin /débits
frm1.Tb_bv(0).ToolTipText = "Débit de pointe d'orage "
frm1.Tb_bv(1).ToolTipText = "Débit de temps sec"
frm1.Tb_bv(2).ToolTipText = "Débit de référence"
'Bassin / Intensité pluie de rinçage
frm1.Tb_bv(3).ToolTipText = "Intensité de la pluie de référence"
'Bassin / Surface du B.V.
frm1.Tb_bv(4).ToolTipText = "Surface du bassin versant"
'Bassin / Coefficient de ruissellement du B.V.
frm1.Tb_bv(5).ToolTipText = "Coefficient de ruissellement appliquée au B.V."
'Bassin / Temps de concentration du B.V.
frm1.Tb_bv(6).ToolTipText = "Temps de concentration"
'Débit aval admissible
frm1.Tb_Qav.ToolTipText = "Débit maximal pouvant être évacué à l'aval de l'ouvrage"
'Schéma
frm1.Tb_long.ToolTipText = ""
frm1.Tb_larg.ToolTipText = ""
frm1.Tb_prof.ToolTipText = ""
frm1.Tb_rap.ToolTipText = ""
End Function


