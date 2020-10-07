Attribute VB_Name = "Ini_tooltip"
Function ini_tooltip_bv(frm1 As Form)
'bassin versant
'Caract�ristiques eau pluviale
frm1.Tb_car_ep(0).ToolTipText = "Surface totale du B.V."
frm1.Tb_car_ep(1).ToolTipText = "Longueur du plus long parcours hydraulique"
frm1.Tb_car_ep(2).ToolTipText = "Pente du plus long parcours hydraulique"
frm1.Tb_car_ep(3).ToolTipText = "Coefficient de ruissellement appliqu� au B.V."
'Caract�ristiques eau us�es
frm1.Tb_car_eu(0).ToolTipText = "Nombre d'�quivalent-habitants sur le B.V."
frm1.Tb_car_eu(1).ToolTipText = "Rejet moyen journalier d'eaux us�es par �quivalent-habitant"
frm1.Tb_car_eu(2).ToolTipText = "% d'eaux claires parasites par rapport au d�bit moyen des eaux us�es"
'Caract�ristiques caract�ristiques
frm1.Tb_carep_rur(0).ToolTipText = "Hauteur de la lame d'eau absorb�e par les pertes initiales"
frm1.Tb_carep_rur(1).ToolTipText = "Variable 'fc' de la loi de Horton "
frm1.Tb_carep_rur(2).ToolTipText = "Param�tre 'a' de la loi de Horton "
frm1.Tb_carep_rur(3).ToolTipText = "Param�tre 'b' de la loi de Horton "
frm1.Tb_carep_rur(4).ToolTipText = ""
'Param�tres eau pluviale
frm1.Tb_par_ep(0).ToolTipText = "Valeur du param�tre 'a' pour les averses de dur�e inf�rieure au seuil"
frm1.Tb_par_ep(1).ToolTipText = "Valeur du param�tre 'b' pour les averses de dur�e inf�rieure au seuil"
frm1.Tb_par_ep(2).ToolTipText = "Valeur du param�tre 'a' pour les averses de dur�e sup�rieure au seuil"
frm1.Tb_par_ep(3).ToolTipText = "Valeur du param�tre 'b' pour les averses de dur�e sup�rieure au seuil"
frm1.Tb_par_ep(4).ToolTipText = "Seuil d'exploitation statistiques de dur�e des averses"
'Param�tres eau us�e
frm1.Tb_par_eu(0).ToolTipText = "Intensit� de la pluie de r�f�rence"
frm1.Tb_par_eu(1).ToolTipText = "Valeur du param�tre 'a' pour �valuation du coefficient de pointe (1,5 par d�faut)"
frm1.Tb_par_eu(2).ToolTipText = "Valeur du param�tre 'b' pour �valuation du coefficient de pointe (2,5 par d�faut)"""
'Param�tres pluie de projet
frm1.Tb_par_pl(0).ToolTipText = "Dur�e totale de la pluie"
frm1.Tb_par_pl(1).ToolTipText = "Dur�e de la p�riode intense"
frm1.Tb_par_pl(2).ToolTipText = "Hauteur totale pr�cipit�e - Double clic pour calcul avec loi de Montana"
frm1.Tb_par_pl(3).ToolTipText = "Hauteur pr�cipit�e pendant la p�riode intense - Double clic pour calcul avec loi de Montana"
frm1.Tb_par_pl(4).ToolTipText = "D�calage de l'intant de la pointe- Par d�faut 0.5(pluie centr�e)"
frm1.Tb_par_pl(5).ToolTipText = "Pas de temps de discr�tisation"
'D�bit pluie d'orage
frm1.Tb_Debit(0).ToolTipText = "D�bit calcul� par la m�thode de Caquot"
frm1.Tb_Debit(1).ToolTipText = "D�bit calcul� par la m�thode rationnelle"
frm1.Tb_Debit(2).ToolTipText = "D�bit de pointe du calcul par la m�thode de l'hydrogramme"
'D�bit des eaux us�es
frm1.Tb_debit1(0).ToolTipText = "D�bit de pointe des eaux us�es : Qeu"
'D�bit de temps sec
frm1.Tb_debit1(1).ToolTipText = "D�bit de pointe de temps sec : Qts = Qes + Qecp"
'D�bit des eaux claires
frm1.Tb_debit1(2).ToolTipText = "D�bit des eaux claires parasites : Qecp"
'D�bit de r�f�rence
frm1.Tb_debit1(3).ToolTipText = "D�bit de r�f�rence : Qref = Qpref + Qts"
'D�bit de pluie de r�f�rence
frm1.Tb_debit1(4).ToolTipText = "D�bit de la pluie de r�f�rence : Qpref"
'D�bit d' orage
frm1.Tb_debit1(5).ToolTipText = "D�bit de la pluie d'orage : Qora (choix Caquot, Rationnelle ou Hydrogramme)"
'Volume total ruissel�
frm1.Tb_debit1(6).ToolTipText = "Volume ruissel�e sur le BV (m�thode de l'hydrogramme)"
End Function
Function ini_tooltip_chute(frm1 As Form)
'chute
'Conduite amont
frm1.Tb_amo(0).ToolTipText = "Diam�tre de la canalisation en amont de la chute"
frm1.Tb_amo(1).ToolTipText = "Pente de la canalisation en amont de la chute"
frm1.Tb_amo(2).ToolTipText = "Coefficient de Manning-Strickler de la canalisation en amont de la chute"
frm1.Tb_amo(3).ToolTipText = "Cote du fil d'eau d'arriv�e de la canalisation amont"
'Conduite aval
frm1.Tb_ava(0).ToolTipText = "Diam�tre de la canalisation en val de la chute"
frm1.Tb_ava(1).ToolTipText = "Pente de la canalisation en aval de la chute"
frm1.Tb_ava(2).ToolTipText = "Coefficient de Manning-Strickler de la canalisation en aval de la chute"
frm1.Tb_ava(3).ToolTipText = "Cote du fil d'eau de d�part de la canalisation aval"
'D�bit
frm1.Tb_Qmax.ToolTipText = "D�bit maximal � transiter"
End Function
Function ini_tooltip_pompe(frm1 As Form)
'pompe
'D�bits Caract.
frm1.Tb_Debit(0).ToolTipText = "Valeur du d�bit moyen journalier des eaux us�es (Qeum)"
frm1.Tb_Debitc(0).ToolTipText = "Valeur du d�bit moyen journalier des eaux us�es (Qeum)"
frm1.Tb_Debit(1).ToolTipText = "Valeur du d�bit de pointe des eaux us�es (Qeu=Qeum x p)"
frm1.Tb_Debitc(1).ToolTipText = "Valeur du d�bit de pointe des eaux us�es (Qeu=Qeum x p)"
frm1.Tb_Debit(2).ToolTipText = "Valeur du d�bit des eaux claires parasites (Qecp)"""
frm1.Tb_Debitc(2).ToolTipText = "Valeur du d�bit des eaux claires parasites (Qecp)"""
frm1.Tb_Debit(3).ToolTipText = "Valeur du d�bit  moyen de temps sec (Qmts=Qeum+Qecp)"
frm1.Tb_Debitc(3).ToolTipText = "Valeur du d�bit  moyen de temps sec (Qmts=Qeum+Qecp)"
frm1.Tb_Debit(4).ToolTipText = "Valeur du d�bit de pointe de temps sec (Qts=Qeu+Qecp)"
frm1.Tb_Debitc(4).ToolTipText = "Valeur du d�bit de pointe de temps sec (Qts=Qeu+Qecp)"
frm1.Tb_FPointe.ToolTipText = "Facteur de pointe des eaux us�es (p)"
frm1.Tb_Qpomp(0).ToolTipText = "D�bit de pompage th�orique (QpompTh�=3 x Qeum +1)"
frm1.Tb_Qpompc(0).ToolTipText = "D�bit de pompage th�orique (QpompTh�=3 x Qeum +1)"
'Donn�es g�om�tr.
'***Conduite de refoulement
frm1.Tb_Geom(0).ToolTipText = "Longueur d�velopp�e de la canalisation de refoulement"
frm1.Tb_Geom(1).ToolTipText = "Diam�tre th�orique de la canalisation pour une vitesse d'�coulement de1.5 m/s"
frm1.Cb_Materiau.ToolTipText = "Mat�riau choisi pour la canalisation de refoulement"
'***Niveaux
frm1.Tb_Geom(3).ToolTipText = "Cote du terrain naturel au droit du poste de pompage"
frm1.Tb_Geom(4).ToolTipText = "Cote du fil d'eau de la canalisation d'arriv�e dans le poste"
frm1.Tb_Geom(5).ToolTipText = "Cote du fil d'eau de d�part de la canalisation de refoulement"
frm1.Tb_Geom(6).ToolTipText = "Cote du fil d'eau � l'extr�mit� du refoulement"
'Points singul.
frm1.Tb_PtSing(0).ToolTipText = "Nbre de coude(s) � 11�15 pr�vu(s) sur le refoulement"
frm1.Tb_PtSing(1).ToolTipText = "Nbre de coude(s) � 22�30 pr�vu(s) sur le refoulement "
frm1.Tb_PtSing(2).ToolTipText = "Nbre de coude(s) � 30� pr�vu(s) sur le refoulement "
frm1.Tb_PtSing(3).ToolTipText = "Nbre de coude(s) � 45� pr�vu(s) sur le refoulement "
frm1.Tb_PtSing(4).ToolTipText = "Nbre de coude(s) � 90� pr�vu(s) sur le refoulement "
frm1.Tb_PtSing(5).ToolTipText = "Nbre de vanne(s) pr�vue(s) sur le refoulement "
frm1.Tb_PtSing(6).ToolTipText = "Nbre de clapet(s) anti-retour pr�vu(s) sur le refoulement "
frm1.Tb_PtSing(7).ToolTipText = "Nbre de syst�me(s) de vidange pr�vu(s) sur le refoulement "
frm1.Tb_PtSing(8).ToolTipText = "Nbre de ventouse(s) pr�vue(s) sur le refoulement "
frm1.Opt_PtSing(0).ToolTipText = "Mise en place d'un syst�me de protection Anti-B�lier (OUI/NON)"
frm1.Opt_PtSing(1).ToolTipText = "Mise en place d'un syst�me de protection Anti-B�lier (OUI/NON)"
'Donn�es tech.
frm1.Tb_Nbpom.ToolTipText = "Nb de pompe(s) install�es (en g�n�ral 2)"
frm1.Tb_Ntdph.ToolTipText = "Nb de d�marrage(s) pr�vus par heure (2 � 6 suivant la puissance des pompes)"
frm1.Tb_Vutba.ToolTipText = "Volume utile th�orique de la b�che de pompage"
'***Section de la b�che
frm1.Opt_sect_ba(0).ToolTipText = "Choix pour une b�che de section circulaire"
frm1.Opt_sect_ba(1).ToolTipText = "Choix pour une b�che de section rectangulaire"
frm1.Tb_long.ToolTipText = "Longueur de la b�che rectangulaire"
frm1.Tb_larg.ToolTipText = "Largeur de la b�che rectangulaire"
frm1.Tb_diam.ToolTipText = "Diam�tre de la b�che circulaire"
'
frm1.Tb_denivt.ToolTipText = "Tranche de pompage th�orique"
frm1.Tb_denivhau.ToolTipText = "Garde � l'�gout : distance entre le niveau de l'arriv�e des eaux et le niveau de d�marrage"
frm1.Tb_denivbas.ToolTipText = "Garde au fond"
'R�sultats
frm1.Tb_Qpomp(1).ToolTipText = "D�bit de pompage retenu"
frm1.Tb_Qpompc(1).ToolTipText = "D�bit de pompage retenu m3"
frm1.Tb_Drflt.ToolTipText = "Diam�tre int�rieur de la canalisation de refoulement"
frm1.Tb_VitRflt.ToolTipText = "Vitesse instantan�e de l'�coulement en r�gime  permanent (conseill�e entre 0.8 et 1.2 m/s)"
frm1.Tb_Jmpkm.ToolTipText = "Pertes de charge dues au lin�aire de canalisation"
frm1.Tb_denivr.ToolTipText = "D�nivel�e entre le capteur de d�marrage et le capteur d'arr�t des pompes"
frm1.Tb_vurba.ToolTipText = "Volume utile de la b�che"
frm1.Tb_nrdph.ToolTipText = "Nombre r�el de d�marrage(s) par heure "
frm1.Tb_Tvidange.ToolTipText = "Temps de vidange du volume utile"
frm1.Tb_T1cyc.ToolTipText = "Dur�e totale d'un cycle (remplissage + vidange)"
frm1.Tb_Nbcyc.ToolTipText = "Nombre de cycles par heure"
frm1.Tb_Vmy.ToolTipText = "Vitesse moyenne d'�coulement"
frm1.Tb_Tsejh.ToolTipText = "Temps de s�jour"
frm1.Tb_Singul.ToolTipText = "Pertes de charge dues aux singularit�s (coudes, vannes,etc...)"
frm1.Tb_Hmt.ToolTipText = "Hauteur manom�trique totale"

End Function

Function ini_tooltip_conduite(frm1 As Form)
'Conduite
frm1.Tb_cond(0).ToolTipText = "Diam�tre de la canalisation"
frm1.Tb_cond(1).ToolTipText = "Pente de la canalisation "
frm1.Tb_cond(2).ToolTipText = "Coefficient de Manning-Strickler de la canalisation"
'D�bit
frm1.Tb_Qmax.ToolTipText = "D�bit maximal � transiter"
End Function
Function ini_tooltip_decant(frm1 As Form)
'bassin de d�cantation
'Param�tres
frm1.Tb_dec(0).ToolTipText = "D�bit entrant dans l'ouvrage"
frm1.Tb_dec(1).ToolTipText = "Taille des particules � d�canter" + " (domaine de validit� 0,125-0,315)"
frm1.Tb_dec(2).ToolTipText = "Rapport d�crivant la section transversale de l'ouvrage"
frm1.Tb_dec(3).ToolTipText = "% souhait� de d�cantation des particules de tailles retenues" + " (domaine de validit� 85-100)"
frm1.Tb_dec(4).ToolTipText = "Vitesse horizontale des particules"
End Function
Function ini_tooltip_do(frm1 As Form)
'd�versoir d'orage
' Bassin versant / hydraulique du BV
frm1.Tb_Debit(0).ToolTipText = "D�bit d'orage des eaux pluviales"
frm1.Tb_Debit(1).ToolTipText = "D�bit de temps sec"
frm1.Tb_Debit(2).ToolTipText = "D�bit de r�f�rence"
'Contraintes
frm1.Tb_cont(0).ToolTipText = "C�te impos�e de d�part en amont du syst�me"
frm1.Tb_cont(1).ToolTipText = "C�te impos�e d'arriv�e � l'aval du syst�me"
frm1.Tb_cont(2).ToolTipText = "Longueur disponible entre les contraintes amont et aval"
frm1.Tb_cont(3).ToolTipText = "C�te des plus hautes eaux � l'exutoire de la canalisation de d�charge"
frm1.Tb_cont(4).ToolTipText = "C�te du fil d'eau � l'exutoire de la canalisation de d�charge"
frm1.Tb_cont(5).ToolTipText = "Longueur de la canalisation de d�charge"
'Canal. amont / Conduite
frm1.Tb_amo(0).ToolTipText = "Diam�tre de la canalisation de tranquilisation en amont de la chambre de d�versement"
frm1.Tb_amo(1).ToolTipText = "Pente de la canalisation de tranquilisation "
frm1.Tb_amo(2).ToolTipText = "Coefficient de Manning-Strickler de la canalisation"
frm1.Tb_amo(3).ToolTipText = "Longueur de la canalisation de tranquilisation (minimum L = 20 x D)"
'Canal. aval / Conduite
frm1.Tb_ava(0).ToolTipText = "Diam�tre de la canalisation �trangl�e en aval de la chambre de d�versement"
frm1.Tb_ava(1).ToolTipText = "Pente de la canalisation �trangl�e"
frm1.Tb_ava(2).ToolTipText = "Coefficient de Manning-Strickler de la canalisation"
frm1.Tb_ava(3).ToolTipText = "Longueur de la canalisation �trangl�e (50 � 60 m maximum)"
'D�versoir / Caract�ristiques
frm1.Tb_dev(0).ToolTipText = "Longueur de la lame de d�versement "
frm1.Tb_dev(1).ToolTipText = "Hauteur de la cr�te (Mini 0.20 m / Conseill�e 0.6 x D amont)"
frm1.Tb_dev(2).ToolTipText = "Pente longitudinale du d�versoir"
frm1.Tb_dev(3).ToolTipText = "Valeur maximale du niveau de la ligne d'eau tol�r�e "
'D�charge / Conduite
frm1.Tb_dech(0).ToolTipText = "Diam�tre de la canalisation de d�charge"
frm1.Tb_dech(1).ToolTipText = "Pente de la canalisation de d�charge"
frm1.Tb_dech(2).ToolTipText = "Coefficient de Manning-Strickler de la canalisation de d�charge"
frm1.Tb_dech(3).ToolTipText = ""
'D�charge / Conduite / Coefficient de perte � l'entr�e
End Function
Function ini_tooltip_door(frm1 As Form)
'D�versoir d'orage � ouverture de radier
' Bassin versant
frm1.Tb_Debit(0).ToolTipText = "D�bit d'orage des eaux pluviales"
frm1.Tb_Debit(1).ToolTipText = "D�bit de temps sec"
frm1.Tb_Debit(2).ToolTipText = "D�bit de r�f�rence"
'Conduite arriv�e
frm1.Tb_amo(0).ToolTipText = "Diam�tre "
frm1.Tb_amo(1).ToolTipText = "Pente "
frm1.Tb_amo(2).ToolTipText = "Coefficient de Manning-Strickler "
frm1.Tb_amo(3).ToolTipText = "Longueur "
'Conduite d�part
frm1.Tb_ava(0).ToolTipText = "Diam�tre "
frm1.Tb_ava(1).ToolTipText = "Pente "
frm1.Tb_ava(2).ToolTipText = "Coefficient de Manning-Strickler "
frm1.Tb_ava(3).ToolTipText = "Longueur "
'Conduite d�versement
frm1.Tb_dech(0).ToolTipText = "Diam�tre "
frm1.Tb_dech(1).ToolTipText = "Pente "
frm1.Tb_dech(2).ToolTipText = "Coefficient de Manning-Strickler"
frm1.Tb_dech(3).ToolTipText = "Longueur "
'Contraintes
frm1.Tb_cont(1).ToolTipText = "cote radier aval,conduite arriv�e "
frm1.Tb_hmin.ToolTipText = "Hauteur entre canalisations "
'D�versoir
frm1.Tb_dev(0).ToolTipText = "Longueur"
frm1.Tb_dev(1).ToolTipText = "Profondeur "
frm1.Chk_cri.ToolTipText = "D�bit critique"
frm1.Chk_max.ToolTipText = "D�bit d'orage "
frm1.Cmd_recalc.ToolTipText = "Dessin recalcul�"
frm1.Cmd_mini.ToolTipText = "Dessin mini "
End Function
Function ini_tooltip_ret(frm1 As Form)
'bassin de r�tention
'Bassin versant
frm1.Tb_bv(0).ToolTipText = "Surface totale du bassin versant"
frm1.Tb_bv(1).ToolTipText = "Coefficient d'apport (Attention diff�rent du coefficient de ruissellement)"
'Param�tres pluviom�triques
frm1.Tb_par(0).ToolTipText = "Valeur du param�tre 'a' pour les averses de dur�e inf�rieure au seuil"
frm1.Tb_par(1).ToolTipText = "Valeur du param�tre 'b' pour les averses de dur�e inf�rieure au seuil"
frm1.Tb_par(2).ToolTipText = "Valeur du param�tre 'a' pour les averses de dur�e sup�rieure au seuil"
frm1.Tb_par(3).ToolTipText = "Valeur du param�tre 'b' pour les averses de dur�e sup�rieure au seuil"
frm1.Tb_par(4).ToolTipText = "Seuil de dur�e d'exploitation statistiques des averses"
'D�bit de fuite de la retenue
frm1.Tb_Qf.ToolTipText = "D�bit de vidange de l'ouvrage de r�tention"
'Sch�ma
frm1.Tb_long.ToolTipText = ""
frm1.Tb_larg.ToolTipText = ""
frm1.Tb_prof.ToolTipText = ""
frm1.Tb_rap.ToolTipText = ""
End Function
Function ini_tooltip_siphon(frm1 As Form)
'siphon
'Conduite amont
frm1.Tb_amo(0).ToolTipText = "Diam�tre de la canalisation en amont du siphon"
frm1.Tb_amo(1).ToolTipText = "Pente de la canalisation en amont du siphon"
frm1.Tb_amo(2).ToolTipText = "Coefficient de Manning-Strickler de la canalisation en amont du siphon"
frm1.Tb_amo(3).ToolTipText = "Cote du fil d'eau d'arriv�e de la canalisation en amont du siphon"
'Conduite aval
frm1.Tb_ava(0).ToolTipText = "Diam�tre de la canalisation en aval du siphon"
frm1.Tb_ava(1).ToolTipText = "Pente de la canalisation en aval du siphon"
frm1.Tb_ava(2).ToolTipText = "Coefficient de Manning-Strickler de la canalisation en aval du siphon"
frm1.Tb_ava(3).ToolTipText = "Cote du fil d'eau de d�part de la canalisation � l'aval du siphon"
'Siphon
frm1.Tb_siph(0).ToolTipText = "Diam�tre de la canalisation du siphon"
frm1.Tb_siph(1).ToolTipText = "Coefficient de Manning-Strickler de la canalisation du siphon"
frm1.Tb_siph(2).ToolTipText = "D�bit � transiter"
frm1.Tb_siph(3).ToolTipText = "Longueur d�velopp�e du siphon"
'Siphon / Coefficient singulatit�s
frm1.Tb_siph(4).ToolTipText = ""
End Function
Function ini_tooltip_stock(frm1 As Form)
'bassin de stockage
'Bassin /d�bits
frm1.Tb_bv(0).ToolTipText = "D�bit de pointe d'orage "
frm1.Tb_bv(1).ToolTipText = "D�bit de temps sec"
frm1.Tb_bv(2).ToolTipText = "D�bit de r�f�rence"
'Bassin / Intensit� pluie de rin�age
frm1.Tb_bv(3).ToolTipText = "Intensit� de la pluie de r�f�rence"
'Bassin / Surface du B.V.
frm1.Tb_bv(4).ToolTipText = "Surface du bassin versant"
'Bassin / Coefficient de ruissellement du B.V.
frm1.Tb_bv(5).ToolTipText = "Coefficient de ruissellement appliqu�e au B.V."
'Bassin / Temps de concentration du B.V.
frm1.Tb_bv(6).ToolTipText = "Temps de concentration"
'D�bit aval admissible
frm1.Tb_Qav.ToolTipText = "D�bit maximal pouvant �tre �vacu� � l'aval de l'ouvrage"
'Sch�ma
frm1.Tb_long.ToolTipText = ""
frm1.Tb_larg.ToolTipText = ""
frm1.Tb_prof.ToolTipText = ""
frm1.Tb_rap.ToolTipText = ""
End Function


