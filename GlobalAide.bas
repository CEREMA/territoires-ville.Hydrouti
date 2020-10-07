Attribute VB_Name = "GlobalAide"
'G�n�ral
Public Const IDhlpAideFichier = "ch02s01.html" ' nom du fichier aide
Public Const IDhlpAideExempleFichier = "ch03s01.html" ' nom du fichier aide
''******version 1
''Chute
'Public Const IDhlpChuteFichier = "ch02s03.html" ' nom du fichier aide
'Public Const IDhlpChuteExempleFichier = "ch03s03.html" ' nom du fichier aide
'Public Const IDhlp_ChuteConduiteAmont = "N10B41" ' "Conduite Amont"
'Public Const IDhlp_ChuteConduiteAval = "N10B46" ' "Conduite Aval"
'Public Const IDhlp_ChuteRegard = "Chap233" '"Regard"
'Public Const IDhlp_ChuteEtudeProfil = "Chap233" ' ' "3. Etude du profil"
'Bassin Versant
'Public Const IDhlpBVFichier = "ch02s02.html" ' nom du fichier aide
'Public Const IDhlpBVExempleFichier = "ch03s02.html" ' nom du fichier aide
'Public Const IDhlp_BVTypeBassin = "N101D1" '"Type de bassin"
'Public Const IDhlp_BVTempsConcentration = "N10247"  '"Temps de concentration Tc"
'Public Const IDhlp_BVCoefficientRuissellement = "N10216" '"Coefficient de ruissellement Cr"
'Public Const IDhlp_BVDebitEauxUseesDomestiques = "N10930" '"D�bit des eaux us�es domestiques"
'Public Const IDhlp_BVDebitEauxClairesParasites = "N10A18" '"D�bit des eaux claires parasites"
'Public Const IDhlp_BVMethodeSuperficielleCaquot = "N104DC" '"M�thode superficielle de Caquot"
'Public Const IDhlp_BVMethodeRationnelle = "N1046F" '"M�thode Rationnelle "
'Public Const IDhlp_BVMethodeHydrogramme = "Chap22413" '"M�thode de l'hydrogramme"
'Public Const IDhlp_BVDebitReference = "Chap2243"  '"Le d�bit de r�f�rence QREF"
'Public Const IDhlp_BVDebitOrage = "N10B08" '"Le d�bit d'orage QORA"
'Public Const IDhlp_BVCourbesIntensiteDureeFrequence = "N103BB" '"Courbes Intensit�-Dur�e-Fr�quence (IDF)"
'Public Const IDhlp_BVEstimationPertesInitiales = "Chap22414" '"Estimation des pertes initiales"
'Public Const IDhlp_BVEstimationPertesContinues = "Chap22415" '"Estimation des pertes continues"
'Public Const IDhlp_BVModeleRuissellement = "N10797" '"Mod�le de ruissellement"  ' r�servoir lin�aire '"
'Public Const IDhlp_BVCaracteristiques = "N10204" '"Caract�ristiques d'un BV"
'Public Const IDhlp_BVdebitTempsSec = "N10927" '"Le d�bit de temps sec QTS"
'Public Const IDhlp_BVNonUrbain = "Chap22413" '"Pour les bassins versants de type non urbain"
'Public Const IDhlp_BVDebitsCaracteristiques = "Chap2212" '"D�bits caract�ristiques"
''Siphon
'Public Const IDhlpSiphonFichier = "ch02s04.html" ' nom du fichier aide
'Public Const IDhlpSiphonExempleFichier = "ch03s04.html" ' nom du fichier aide
'Public Const IDhlp_SiphonPrincipesHydrauliques = "N10E18" '"Principes hydrauliques"
'Public Const IDhlp_SiphonPertesChargesSingulieres = "Chap24723" '"Pertes de charge singuli�res dans le siphon (coudes)"
''Conduite
'Public Const IDhlpConduiteFichier = "ch02s06.html" '"ch02s05.html" ' nom du fichier aide
'Public Const IDhlpConduiteExempleFichier = "ch02s06.html" '"ch02s05.html" ' nom du fichier aide pas d'exemple
'Public Const IDhlp_ConduiteDimensionnement = "" '"Dimensionnement d'une conduite"
''Bassin de r�tention
'Public Const IDhlpRetentionFichier = "ch02s09.html" ' "ch02s08.html" ' nom du fichier aide
'Public Const IDhlpRetentionExempleFichier = "ch03s09.html" '"ch03s07.html" ' nom du fichier aide
'Public Const IDhlp_RetentionDimensionnementMethodePluies = "N1150B" '"Dimensionnement par la  m�thode des pluies"
'Public Const IDhlp_RetentionCoefficientsMontana = "N1164E" '"Choix des coefficients a et b de Montana"
'Public Const IDhlp_RetentionDimensionnementMethodeHydrogramme = "N1168D" '"Dimensionnement par la  m�thode de l'hydrogramme"
''Bassin de stockage
'Public Const IDhlpStockageFichier = "ch02s10.html" '"ch02s09.html" ' nom du fichier aide
'Public Const IDhlpStockageExempleFichier = "ch02s10.html" '"ch02s09.html" ' nom du fichier aide pas d'exemple
'Public Const IDhlp_StockageOrigineMethode = "N1169A" '"Origine de la m�thode"
'Public Const IDhlp_StockagePresentationMethodeCalcul = "N116B9" '"Pr�sentation de la m�thode de calcul"
''Bassin de d�cantation
'Public Const IDhlpDecantationFichier = "ch02s08.html" '"ch02s07.html" ' nom du fichier aide
'Public Const IDhlpDecantationExempleFichier = "ch03s08.html" '"ch03s06.html" ' nom du fichier aide
'Public Const IDhlp_DecantationModeCalcul = "N11444" '"Mode de calcul hydraulique d'un bassin de d�cantation"
''D�versoir d'orage a crete haute
'Public Const IDhlpDOFichier = "ch02s07.html" '"ch02s06.html" ' nom du fichier aide
'Public Const IDhlpDOExempleFichier = "ch03s06.html" '"ch03s05.html" ' nom du fichier aide
'Public Const IDhlp_DODonneesBase = "N1123B" '"Donn�es de base"
'Public Const IDhlp_DOContraintes = "N1121B" '"D�versoir � seuil haut"
'Public Const IDhlp_DOConduiteAmenee = "N1125D" '"Conduite d'amen�e"
'Public Const IDhlp_DOConduiteDebitConserve = "N11266" '"Conduite de d�bit conserv�"
'Public Const IDhlp_DOChambreDeversement = "N1132D" '"Chambre de d�versement"
'Public Const IDhlp_DOConduiteDecharge = "N113B7" '"Conduite de d�charge"
''D�versoir d'orage a ouverture de radier
'Public Const IDhlpDOORFichier = "ch02s07.html" '"ch02s06.html" ' nom du fichier aide
'Public Const IDhlpDOORExempleFichier = "ch03s07.html" '"ch03s05.html" ' nom du fichier aide
'Public Const IDhlp_DOORDonneesBase = "N1123B" '"Hydraulique du bassin versant"
'Public Const IDhlp_DOORContraintes = "N1121B" '"Contraintes"
'Public Const IDhlp_DOORConduiteArrivee = "N1125D" '"Conduite d'arriv�e"
'Public Const IDhlp_DOORConduiteDepart = "N11266" '"Conduite de d�part"
'Public Const IDhlp_DOOROuvrageDeversoir = "N1132D" '"L'ouvrage d�versoir"
'Public Const IDhlp_DOORConduiteDeversement = "N113B7" '"Conduite de d�versement"
'Public Const IDhlp_DOORMethodeDimensionnement = "N1123B" '"M�thode de dimensionnement"
'******version 2
'Chute
Public Const IDhlpChuteFichier = "ch02s03.html" ' nom du fichier aide
Public Const IDhlpChuteExempleFichier = "ch03s03.html" ' nom du fichier aide
Public Const IDhlp_ChuteConduiteAmont = "ChuteConduiteAmont" '"N10B81" ' "Conduite Amont"
Public Const IDhlp_ChuteConduiteAval = "ChuteConduiteAval" '"N10B86" ' "Conduite Aval"
Public Const IDhlp_ChuteRegard = "ChuteRegard" '"Chap233" '"Regard"
Public Const IDhlp_ChuteEtudeProfil = "ChuteEtudeProfil" '"Chap233" ' ' "3. Etude du profil"
'Bassin Versant
Public Const IDhlpBVFichier = "ch02s02.html" ' nom du fichier aide
Public Const IDhlpBVExempleFichier = "ch03s02.html" ' nom du fichier aide
Public Const IDhlp_BVTypeBassin = "BVTypeBassin" '"N10207" '"Type de bassin"
Public Const IDhlp_BVTempsConcentration = "BVTempsConcentration" '"N1027D"  '"Temps de concentration Tc"
Public Const IDhlp_BVCoefficientRuissellement = "BVCoefficientRuissellement" '"N1024C" '"Coefficient de ruissellement Cr"
Public Const IDhlp_BVDebitEauxUseesDomestiques = "BVDebitEauxUseesDomestiques" '"N1096E" '"D�bit des eaux us�es domestiques"
Public Const IDhlp_BVDebitEauxClairesParasites = "BVDebitEauxClairesParasites" '"N10A56" '"D�bit des eaux claires parasites"
Public Const IDhlp_BVPluieProjet = "BVPluieProjet"  '"Pluie de projet"
Public Const IDhlp_BVMethodeSuperficielleCaquot = "BVMethodeSuperficielleCaquot" '"N10512" '"M�thode superficielle de Caquot"
Public Const IDhlp_BVMethodeRationnelle = "BVMethodeRationnelle" '"N104A5" '"M�thode Rationnelle "
Public Const IDhlp_BVMethodeHydrogramme = "BVMethodeHydrogramme" '"Chap22413" '"M�thode de l'hydrogramme"
Public Const IDhlp_BVDebitReference = "BVDebitReference" '"Chap2243"  '"Le d�bit de r�f�rence QREF"
Public Const IDhlp_BVDebitOrage = "BVDebitOrage" '"N10B46" '"Le d�bit d'orage QORA"
Public Const IDhlp_BVCourbesIntensiteDureeFrequence = "BVCourbesIntensiteDureeFrequence" '"N103F1" '"Courbes Intensit�-Dur�e-Fr�quence (IDF)"
Public Const IDhlp_BVEstimationPertesInitiales = "BVEstimationPertesInitiales" '"Chap22414" '"Estimation des pertes initiales"
Public Const IDhlp_BVEstimationPertesContinues = "BVEstimationPertesContinues" '"Chap22415" '"Estimation des pertes continues"
Public Const IDhlp_BVModeleRuissellement = "BVModeleRuissellement" '"N107D5" '"Mod�le de ruissellement"  ' r�servoir lin�aire '"
Public Const IDhlp_BVCaracteristiques = "BVCaracteristiques" '"N1023A" '"Caract�ristiques d'un BV"
Public Const IDhlp_BVdebitTempsSec = "BVdebitTempsSec" '"N10965" '"Le d�bit de temps sec QTS"
Public Const IDhlp_BVNonUrbain = "BVNonUrbain" '"Chap22413" '"Pour les bassins versants de type non urbain"
Public Const IDhlp_BVDebitsCaracteristiques = "BVDebitsCaracteristiques" '"Chap2212" '"D�bits caract�ristiques"
'Siphon
Public Const IDhlpSiphonFichier = "ch02s04.html" ' nom du fichier aide
Public Const IDhlpSiphonExempleFichier = "ch03s04.html" ' nom du fichier aide
Public Const IDhlp_SiphonPrincipesHydrauliques = "SiphonPrincipesHydrauliques" '"N10E5A" '"Principes hydrauliques"
Public Const IDhlp_SiphonPertesChargesSingulieres = "SiphonPertesChargesSingulieres" '"Chap24723" '"Pertes de charge singuli�res dans le siphon (coudes)"
'Conduite
Public Const IDhlpConduiteFichier = "ch02s06.html" '"ch02s05.html" ' nom du fichier aide
Public Const IDhlpConduiteExempleFichier = "ch03s06.html" '"ch02s05.html" ' nom du fichier aide pas d'exemple
Public Const IDhlp_ConduiteDimensionnement = "" '"Dimensionnement d'une conduite"
'Bassin de r�tention
Public Const IDhlpRetentionFichier = "ch02s10.html" ' "ch02s08.html" ' nom du fichier aide
Public Const IDhlpRetentionExempleFichier = "ch03s10.html" '"ch03s07.html" ' nom du fichier aide
Public Const IDhlp_RetentionDimensionnementMethodePluies = "RetentionDimensionnementMethodePluies" '"N1160F" '"Dimensionnement par la  m�thode des pluies"
Public Const IDhlp_RetentionCoefficientsMontana = "RetentionCoefficientsMontana" '"N11752" '"Choix des coefficients a et b de Montana"
Public Const IDhlp_RetentionDimensionnementMethodeHydrogramme = "RetentionDimensionnementMethodeHydrogramme" '"N11791" '"Dimensionnement par la  m�thode de l'hydrogramme"
'Bassin de stockage
Public Const IDhlpStockageFichier = "ch02s11.html" '"ch02s09.html" ' nom du fichier aide
Public Const IDhlpStockageExempleFichier = "ch03s11.html" '"ch02s09.html" ' nom du fichier aide pas d'exemple
Public Const IDhlp_StockageOrigineMethode = "StockageOrigineMethode" '"N117A0" '"Origine de la m�thode"
Public Const IDhlp_StockagePresentationMethodeCalcul = "StockagePresentationMethodeCalcul" '"N117BF" '"Pr�sentation de la m�thode de calcul"
'Bassin de d�cantation
Public Const IDhlpDecantationFichier = "ch02s09.html" '"ch02s07.html" ' nom du fichier aide
Public Const IDhlpDecantationExempleFichier = "ch03s09.html" '"ch03s06.html" ' nom du fichier aide
Public Const IDhlp_DecantationModeCalcul = "DecantationModeCalcul" '"N11546" '"Mode de calcul hydraulique d'un bassin de d�cantation"
'D�versoir d'orage a crete haute
Public Const IDhlpDOFichier = "ch02s07.html" '"ch02s06.html" ' nom du fichier aide
Public Const IDhlpDOExempleFichier = "ch03s07.html" '"ch03s05.html" ' nom du fichier aide
Public Const IDhlp_DODonneesBase = "DODonneesBase" '"N11339" '"Donn�es de base"
Public Const IDhlp_DOContraintes = "DOContraintes" '"N11319" '"D�versoir � seuil haut"
Public Const IDhlp_DOConduiteAmenee = "DOConduiteAmenee" '"N1135B" '"Conduite d'amen�e"
Public Const IDhlp_DOConduiteDebitConserve = "DOConduiteDebitConserve" '"N11364" '"Conduite de d�bit conserv�"
Public Const IDhlp_DOChambreDeversement = "DOChambreDeversement" '"N1142D" '"Chambre de d�versement"
Public Const IDhlp_DOConduiteDecharge = "DOConduiteDecharge" '"N114B7" '"Conduite de d�charge"
'D�versoir d'orage a ouverture de radier
Public Const IDhlpDOORFichier = "ch02s08.html" '"ch02s06.html" ' nom du fichier aide
Public Const IDhlpDOORExempleFichier = "ch03s08.html" '"ch03s05.html" ' nom du fichier aide
Public Const IDhlp_DOORDonneesBase = "DOORDonneesBase" '"N12E5D" '"Hydraulique du bassin versant"
Public Const IDhlp_DOORContraintes = "DOORContraintes" '"N12E92" '"Contraintes"
Public Const IDhlp_DOORConduiteArrivee = "DOORConduiteArrivee" '"N12E73" '"Conduite d'arriv�e"
Public Const IDhlp_DOORConduiteDepart = "DOORConduiteDepart" '"N12E86" '"Conduite de d�part"
Public Const IDhlp_DOOROuvrageDeversoir = "DOOROuvrageDeversoir" ' "N12E97" '"L'ouvrage d�versoir"
Public Const IDhlp_DOORConduiteDeversement = "DOORConduiteDeversement" '"N12E8D" '"Conduite de d�versement"
Public Const IDhlp_DOORMethodeDimensionnement = "DOORMethodeDimensionnement" '"N12EB1" '"M�thode de dimensionnement"
'Station de pompage
Public Const IDhlpPompeFichier = "ch02s05.html" '"ch02s06.html" ' nom du fichier aide
Public Const IDhlpPompeExempleFichier = "ch03s05.html" '"ch03s05.html" ' nom du fichier aide
Public Const IDhlp_PompeDebitsCaracteristiques = "PompeDebitsCaracteristiques"
Public Const IDhlp_PompeDonneesGeometriques = "PompeDonneesGeometriques"
Public Const IDhlp_PompePointsSinguliers = "PompeCalculPertesCharges" '"PompePointsSinguliers"
Public Const IDhlp_PompePointsSinguliersProtection = "PompeDispositifProtection"
Public Const IDhlp_PompeDonneesTechniques = "PompeDonneesTechniques"
Public Const IDhlp_PompeDonneesTechniques2 = "PompeDonneesTechniques2"

