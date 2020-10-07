Attribute VB_Name = "GlobalAide"
'Général
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
'Public Const IDhlp_BVDebitEauxUseesDomestiques = "N10930" '"Débit des eaux usées domestiques"
'Public Const IDhlp_BVDebitEauxClairesParasites = "N10A18" '"Débit des eaux claires parasites"
'Public Const IDhlp_BVMethodeSuperficielleCaquot = "N104DC" '"Méthode superficielle de Caquot"
'Public Const IDhlp_BVMethodeRationnelle = "N1046F" '"Méthode Rationnelle "
'Public Const IDhlp_BVMethodeHydrogramme = "Chap22413" '"Méthode de l'hydrogramme"
'Public Const IDhlp_BVDebitReference = "Chap2243"  '"Le débit de référence QREF"
'Public Const IDhlp_BVDebitOrage = "N10B08" '"Le débit d'orage QORA"
'Public Const IDhlp_BVCourbesIntensiteDureeFrequence = "N103BB" '"Courbes Intensité-Durée-Fréquence (IDF)"
'Public Const IDhlp_BVEstimationPertesInitiales = "Chap22414" '"Estimation des pertes initiales"
'Public Const IDhlp_BVEstimationPertesContinues = "Chap22415" '"Estimation des pertes continues"
'Public Const IDhlp_BVModeleRuissellement = "N10797" '"Modèle de ruissellement"  ' réservoir linéaire '"
'Public Const IDhlp_BVCaracteristiques = "N10204" '"Caractéristiques d'un BV"
'Public Const IDhlp_BVdebitTempsSec = "N10927" '"Le débit de temps sec QTS"
'Public Const IDhlp_BVNonUrbain = "Chap22413" '"Pour les bassins versants de type non urbain"
'Public Const IDhlp_BVDebitsCaracteristiques = "Chap2212" '"Débits caractéristiques"
''Siphon
'Public Const IDhlpSiphonFichier = "ch02s04.html" ' nom du fichier aide
'Public Const IDhlpSiphonExempleFichier = "ch03s04.html" ' nom du fichier aide
'Public Const IDhlp_SiphonPrincipesHydrauliques = "N10E18" '"Principes hydrauliques"
'Public Const IDhlp_SiphonPertesChargesSingulieres = "Chap24723" '"Pertes de charge singulières dans le siphon (coudes)"
''Conduite
'Public Const IDhlpConduiteFichier = "ch02s06.html" '"ch02s05.html" ' nom du fichier aide
'Public Const IDhlpConduiteExempleFichier = "ch02s06.html" '"ch02s05.html" ' nom du fichier aide pas d'exemple
'Public Const IDhlp_ConduiteDimensionnement = "" '"Dimensionnement d'une conduite"
''Bassin de rétention
'Public Const IDhlpRetentionFichier = "ch02s09.html" ' "ch02s08.html" ' nom du fichier aide
'Public Const IDhlpRetentionExempleFichier = "ch03s09.html" '"ch03s07.html" ' nom du fichier aide
'Public Const IDhlp_RetentionDimensionnementMethodePluies = "N1150B" '"Dimensionnement par la  méthode des pluies"
'Public Const IDhlp_RetentionCoefficientsMontana = "N1164E" '"Choix des coefficients a et b de Montana"
'Public Const IDhlp_RetentionDimensionnementMethodeHydrogramme = "N1168D" '"Dimensionnement par la  méthode de l'hydrogramme"
''Bassin de stockage
'Public Const IDhlpStockageFichier = "ch02s10.html" '"ch02s09.html" ' nom du fichier aide
'Public Const IDhlpStockageExempleFichier = "ch02s10.html" '"ch02s09.html" ' nom du fichier aide pas d'exemple
'Public Const IDhlp_StockageOrigineMethode = "N1169A" '"Origine de la méthode"
'Public Const IDhlp_StockagePresentationMethodeCalcul = "N116B9" '"Présentation de la méthode de calcul"
''Bassin de décantation
'Public Const IDhlpDecantationFichier = "ch02s08.html" '"ch02s07.html" ' nom du fichier aide
'Public Const IDhlpDecantationExempleFichier = "ch03s08.html" '"ch03s06.html" ' nom du fichier aide
'Public Const IDhlp_DecantationModeCalcul = "N11444" '"Mode de calcul hydraulique d'un bassin de décantation"
''Déversoir d'orage a crete haute
'Public Const IDhlpDOFichier = "ch02s07.html" '"ch02s06.html" ' nom du fichier aide
'Public Const IDhlpDOExempleFichier = "ch03s06.html" '"ch03s05.html" ' nom du fichier aide
'Public Const IDhlp_DODonneesBase = "N1123B" '"Données de base"
'Public Const IDhlp_DOContraintes = "N1121B" '"Déversoir à seuil haut"
'Public Const IDhlp_DOConduiteAmenee = "N1125D" '"Conduite d'amenée"
'Public Const IDhlp_DOConduiteDebitConserve = "N11266" '"Conduite de débit conservé"
'Public Const IDhlp_DOChambreDeversement = "N1132D" '"Chambre de déversement"
'Public Const IDhlp_DOConduiteDecharge = "N113B7" '"Conduite de décharge"
''Déversoir d'orage a ouverture de radier
'Public Const IDhlpDOORFichier = "ch02s07.html" '"ch02s06.html" ' nom du fichier aide
'Public Const IDhlpDOORExempleFichier = "ch03s07.html" '"ch03s05.html" ' nom du fichier aide
'Public Const IDhlp_DOORDonneesBase = "N1123B" '"Hydraulique du bassin versant"
'Public Const IDhlp_DOORContraintes = "N1121B" '"Contraintes"
'Public Const IDhlp_DOORConduiteArrivee = "N1125D" '"Conduite d'arrivée"
'Public Const IDhlp_DOORConduiteDepart = "N11266" '"Conduite de départ"
'Public Const IDhlp_DOOROuvrageDeversoir = "N1132D" '"L'ouvrage déversoir"
'Public Const IDhlp_DOORConduiteDeversement = "N113B7" '"Conduite de déversement"
'Public Const IDhlp_DOORMethodeDimensionnement = "N1123B" '"Méthode de dimensionnement"
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
Public Const IDhlp_BVDebitEauxUseesDomestiques = "BVDebitEauxUseesDomestiques" '"N1096E" '"Débit des eaux usées domestiques"
Public Const IDhlp_BVDebitEauxClairesParasites = "BVDebitEauxClairesParasites" '"N10A56" '"Débit des eaux claires parasites"
Public Const IDhlp_BVPluieProjet = "BVPluieProjet"  '"Pluie de projet"
Public Const IDhlp_BVMethodeSuperficielleCaquot = "BVMethodeSuperficielleCaquot" '"N10512" '"Méthode superficielle de Caquot"
Public Const IDhlp_BVMethodeRationnelle = "BVMethodeRationnelle" '"N104A5" '"Méthode Rationnelle "
Public Const IDhlp_BVMethodeHydrogramme = "BVMethodeHydrogramme" '"Chap22413" '"Méthode de l'hydrogramme"
Public Const IDhlp_BVDebitReference = "BVDebitReference" '"Chap2243"  '"Le débit de référence QREF"
Public Const IDhlp_BVDebitOrage = "BVDebitOrage" '"N10B46" '"Le débit d'orage QORA"
Public Const IDhlp_BVCourbesIntensiteDureeFrequence = "BVCourbesIntensiteDureeFrequence" '"N103F1" '"Courbes Intensité-Durée-Fréquence (IDF)"
Public Const IDhlp_BVEstimationPertesInitiales = "BVEstimationPertesInitiales" '"Chap22414" '"Estimation des pertes initiales"
Public Const IDhlp_BVEstimationPertesContinues = "BVEstimationPertesContinues" '"Chap22415" '"Estimation des pertes continues"
Public Const IDhlp_BVModeleRuissellement = "BVModeleRuissellement" '"N107D5" '"Modèle de ruissellement"  ' réservoir linéaire '"
Public Const IDhlp_BVCaracteristiques = "BVCaracteristiques" '"N1023A" '"Caractéristiques d'un BV"
Public Const IDhlp_BVdebitTempsSec = "BVdebitTempsSec" '"N10965" '"Le débit de temps sec QTS"
Public Const IDhlp_BVNonUrbain = "BVNonUrbain" '"Chap22413" '"Pour les bassins versants de type non urbain"
Public Const IDhlp_BVDebitsCaracteristiques = "BVDebitsCaracteristiques" '"Chap2212" '"Débits caractéristiques"
'Siphon
Public Const IDhlpSiphonFichier = "ch02s04.html" ' nom du fichier aide
Public Const IDhlpSiphonExempleFichier = "ch03s04.html" ' nom du fichier aide
Public Const IDhlp_SiphonPrincipesHydrauliques = "SiphonPrincipesHydrauliques" '"N10E5A" '"Principes hydrauliques"
Public Const IDhlp_SiphonPertesChargesSingulieres = "SiphonPertesChargesSingulieres" '"Chap24723" '"Pertes de charge singulières dans le siphon (coudes)"
'Conduite
Public Const IDhlpConduiteFichier = "ch02s06.html" '"ch02s05.html" ' nom du fichier aide
Public Const IDhlpConduiteExempleFichier = "ch03s06.html" '"ch02s05.html" ' nom du fichier aide pas d'exemple
Public Const IDhlp_ConduiteDimensionnement = "" '"Dimensionnement d'une conduite"
'Bassin de rétention
Public Const IDhlpRetentionFichier = "ch02s10.html" ' "ch02s08.html" ' nom du fichier aide
Public Const IDhlpRetentionExempleFichier = "ch03s10.html" '"ch03s07.html" ' nom du fichier aide
Public Const IDhlp_RetentionDimensionnementMethodePluies = "RetentionDimensionnementMethodePluies" '"N1160F" '"Dimensionnement par la  méthode des pluies"
Public Const IDhlp_RetentionCoefficientsMontana = "RetentionCoefficientsMontana" '"N11752" '"Choix des coefficients a et b de Montana"
Public Const IDhlp_RetentionDimensionnementMethodeHydrogramme = "RetentionDimensionnementMethodeHydrogramme" '"N11791" '"Dimensionnement par la  méthode de l'hydrogramme"
'Bassin de stockage
Public Const IDhlpStockageFichier = "ch02s11.html" '"ch02s09.html" ' nom du fichier aide
Public Const IDhlpStockageExempleFichier = "ch03s11.html" '"ch02s09.html" ' nom du fichier aide pas d'exemple
Public Const IDhlp_StockageOrigineMethode = "StockageOrigineMethode" '"N117A0" '"Origine de la méthode"
Public Const IDhlp_StockagePresentationMethodeCalcul = "StockagePresentationMethodeCalcul" '"N117BF" '"Présentation de la méthode de calcul"
'Bassin de décantation
Public Const IDhlpDecantationFichier = "ch02s09.html" '"ch02s07.html" ' nom du fichier aide
Public Const IDhlpDecantationExempleFichier = "ch03s09.html" '"ch03s06.html" ' nom du fichier aide
Public Const IDhlp_DecantationModeCalcul = "DecantationModeCalcul" '"N11546" '"Mode de calcul hydraulique d'un bassin de décantation"
'Déversoir d'orage a crete haute
Public Const IDhlpDOFichier = "ch02s07.html" '"ch02s06.html" ' nom du fichier aide
Public Const IDhlpDOExempleFichier = "ch03s07.html" '"ch03s05.html" ' nom du fichier aide
Public Const IDhlp_DODonneesBase = "DODonneesBase" '"N11339" '"Données de base"
Public Const IDhlp_DOContraintes = "DOContraintes" '"N11319" '"Déversoir à seuil haut"
Public Const IDhlp_DOConduiteAmenee = "DOConduiteAmenee" '"N1135B" '"Conduite d'amenée"
Public Const IDhlp_DOConduiteDebitConserve = "DOConduiteDebitConserve" '"N11364" '"Conduite de débit conservé"
Public Const IDhlp_DOChambreDeversement = "DOChambreDeversement" '"N1142D" '"Chambre de déversement"
Public Const IDhlp_DOConduiteDecharge = "DOConduiteDecharge" '"N114B7" '"Conduite de décharge"
'Déversoir d'orage a ouverture de radier
Public Const IDhlpDOORFichier = "ch02s08.html" '"ch02s06.html" ' nom du fichier aide
Public Const IDhlpDOORExempleFichier = "ch03s08.html" '"ch03s05.html" ' nom du fichier aide
Public Const IDhlp_DOORDonneesBase = "DOORDonneesBase" '"N12E5D" '"Hydraulique du bassin versant"
Public Const IDhlp_DOORContraintes = "DOORContraintes" '"N12E92" '"Contraintes"
Public Const IDhlp_DOORConduiteArrivee = "DOORConduiteArrivee" '"N12E73" '"Conduite d'arrivée"
Public Const IDhlp_DOORConduiteDepart = "DOORConduiteDepart" '"N12E86" '"Conduite de départ"
Public Const IDhlp_DOOROuvrageDeversoir = "DOOROuvrageDeversoir" ' "N12E97" '"L'ouvrage déversoir"
Public Const IDhlp_DOORConduiteDeversement = "DOORConduiteDeversement" '"N12E8D" '"Conduite de déversement"
Public Const IDhlp_DOORMethodeDimensionnement = "DOORMethodeDimensionnement" '"N12EB1" '"Méthode de dimensionnement"
'Station de pompage
Public Const IDhlpPompeFichier = "ch02s05.html" '"ch02s06.html" ' nom du fichier aide
Public Const IDhlpPompeExempleFichier = "ch03s05.html" '"ch03s05.html" ' nom du fichier aide
Public Const IDhlp_PompeDebitsCaracteristiques = "PompeDebitsCaracteristiques"
Public Const IDhlp_PompeDonneesGeometriques = "PompeDonneesGeometriques"
Public Const IDhlp_PompePointsSinguliers = "PompeCalculPertesCharges" '"PompePointsSinguliers"
Public Const IDhlp_PompePointsSinguliersProtection = "PompeDispositifProtection"
Public Const IDhlp_PompeDonneesTechniques = "PompeDonneesTechniques"
Public Const IDhlp_PompeDonneesTechniques2 = "PompeDonneesTechniques2"

