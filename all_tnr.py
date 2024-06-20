import unittest
import os
import time
import datetime
import docx
from os import path
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException, NoSuchElementException
from selenium.common.exceptions import *
from selenium.webdriver.common.keys import Keys
import glob
import pandas as pd
import re
import openpyxl
import os
from os import path

import Utilities as util
from Utilities import *

class ensemble_des_tnr:
    #def setup(self):
    #    #util.lancer_SISTER()
    #    global wait, driver, wait_court
    #    driver = webdriver.Chrome("C:\webdrivers\chromedriver.exe")
#
    #    wait = WebDriverWait(driver, 40)
    #    wait_court = WebDriverWait(driver, 15)
    #    driver.get("https://xxx.com")
    #    driver.maximize_window()
    #    # driver.find_element_by_xpath("//*[@name='userid']").send_keys('RGA')
    def creation_article(self):
        fonctionnel.demarrer_tnrSISTER(self, 'TNR-CA', '/tnrCA_', 'CA')
        fn_tnr25.renseigner_classe_article(self, 'MOBILE')
        fn_tnr25.renseigner_Marketing(self)
        #fn_tnr25.renseigner_Dirco(self)
        fn_tnr25.renseigner_Achat(self)
        fn_tnr25.test(self)

    def Création_article_en_masse_69(self):
        fonctionnel.demarrer_tnrSISTER(self, 'TNR-69', '/tnr69_', '69')
        fonctionnel.gestion_des_articles(self, True, '69')

        fonctionnel.reaffectation_lot_article_au_bon_utilisateur(self, 'RGA', '69')

    def Unicité_des_champs_de_la_fiche_article_71(self):
        fonctionnel.demarrer_tnrSISTER(self, 'TNR-71', '/tnr71_', '71')

        util.rechercher_EAN13('227481', 'IPHONE 8 GRI 64 GO', 'FEH_IO_MASTER')
        util.cliquer("//a[contains(@id, 'p:table1:0:itmNoLnkForSearch') and text() = '227481']")
        util.cliquer("//a[text()='Spécifications']")
        util.cliquer("//a[text()='SISTER: Marketing']")

        fn_tnr71.assurer_EAN13_positif(self)
        fn_tnr71.obtenir_informations(self)

        fn_tnr71.copiercoller("//input[contains(@name, 'EAN_AG:0:XXSCFEANItemAT')]", '227481', 'a', 'c', '')  # Copier l'EAN 13 Article pour l'article 227481 pour la collée dans l'article 226352
        util.scroll("//div[@title = 'EAN']")
        util.driver.get_screenshot_as_file(util.Directory + "/EAN13_227481.png")
        fonctionnel.vers_Word("Vérification de l'EAN13 pour l'article 227481", util.Directory + "/EAN13_227481.png")
        util.driver.close()
        # ***** Deuxieme partie *****
        util.lancer_SISTER()
        fonctionnel.login(self, '71')

        util.rechercher_EAN13('226352', '', 'FEH_IO_MASTER')
        util.cliquer("//span/a[contains(@id, 'p:table1:0:itmNoLnkForSearch')]")  # 226352
        # U.Cliquer("//span[text()='FEH_IO_MASTER']/ancestor::table[contains(@summary, 'Résultats de la recherche d')]/descendant::tbody/tr[2]/td/span/a[text()='226352']")
        util.cliquer("//a[text()='Spécifications']")
        util.cliquer("//a[text()='SISTER: Marketing']")
        # Coller dans l'article 226352 avant de sauvegarder
        util.scroll("//div[@title = 'EAN']")
        util.driver.save_screenshot(util.Directory + "/EAN13_226352.png")
        fonctionnel.vers_Word("Vérification de l'EAN13 pour l'article 226352", util.Directory + "/EAN13_226352.png")
        #
        fn_tnr71.copiercoller("//input[contains(@name, '_EAN_AG:0:XXSCFEANItemAT')]", '226352', '', '', 'v')  # Coller
        util.driver.save_screenshot(util.Directory + "/EAN13_modifier.png")
        fonctionnel.vers_Word("L'EAN13 pour l'article 227481 est maintenant collée dans l'article 226352",
                              util.Directory + "/EAN13_modifier.png")
        util.cliquer("//*[text()='Sauvegarder']")
        #
        fn_tnr71.testsave(1)
        #
        # Comparaison des valeurs pour les 2 articles en question (* Optionel *)
        fn_tnr71.supprimer_remplacer_info(self)  # Effacer les valeurs des champs pour l'article 226352 et remplacer par celles de l'article (227481)
        util.cliquer("//*[text()='Sauvegarder']")
        fn_tnr71.testsave(2)
        #
        # Second test avec le reste des champs modifiés
        util.scroll("//input[contains(@name, 'ItemXxscfOrangeIcPrivateVOXXSCF_EAN_AG:0:XXSCFEANItemAT')]")
        util.remplir_champ("//input[contains(@name, 'ItemXxscfOrangeIcPrivateVOXXSCF_EAN_AG:0:XXSCFEANItemAT')]", util.EAN13Article_original)

        util.driver.save_screenshot(util.Directory + "/EAN13_226352.png")
        fonctionnel.vers_Word("Insérer l'EAN 13 originale pour l'article 226352 avant de vérifier si la sauvegarde aura lieu", util.Directory + "/EAN13_226352.png")
        util.cliquer("//*[text()='Sauvegarder']")
        fn_tnr71.dernier_test(self)

    def Format_du_champ_EAN13_de_la_fiche_article_72(self):
        fonctionnel.demarrer_tnrSISTER(self, 'TNR-72', '/tnr72_', '72')
        fn_tnr7_256.naviguer_vers_Accueil("//div[@title='Gestion des produits']", "//a[@title='Gestion des informations produit']")
        fn_tnr72.choisir_FEH_IO_MASTER(231203)

        cliquer("//a[text()='Spécifications']")
        cliquer("//a[text()='SISTER: Marketing']")
        fn_tnr72.modifier_utilisationEAN13_par_oui(self)

        fn_tnr72.nombre_chiffre_test("C:/Users/" + os.getlogin() + "/Desktop/TNR-72", 100_000_000_000, 1000_000_000_000, 12)
        fn_tnr72.nombre_chiffre_test("C:/Users/" + os.getlogin() + "/Desktop/TNR-72", 10_000_000_000_000, 100_000_000_000_000, 14)

        fn_tnr72.chiffre_a_13_errones("C:/Users/" + os.getlogin() + "/Desktop/TNR-72")

    def Historique_de_modication_de_fiche_article_Code_court_75(self):
        fonctionnel.demarrer_tnrSISTER(self, 'TNR-75', '/tnr75_', '75')
        fn_tnr7_256.naviguer_vers_Accueil("//div[@title = 'Outils']", "//a[contains(@title, 'Etats d')]")

        fn_tnr7_256.alimentation_des_champs('07/02/21', 'Avant', 'Product Hub', 'Article', '231203')
        util.driver.get_screenshot_as_file(util.Directory + "/Resultat_check.png")
        fonctionnel.vers_Word("Verification: Apparition de l'article modifié dans la partie 'Résultats de la recherche'", util.Directory + "/Resultat_check.png")

        fn_tnr7_256.cocher_tous_les_attributs(self)

        util.driver.get_screenshot_as_file(util.Directory + "/Tout_attributs_check.png")
        fonctionnel.vers_Word("Vérification visibilite des champs modifiés(Auteur, date/heure et Ancienne/nouvelle valeur)", util.Directory + "/Tout_attributs_check.png")

        #fn_tnr7_256.extraction_des_donnees_auto(self)
        fn_tnr7_256.telecharger_les_donnees(self)

    def Historique_de_modication_de_fiche_article_Date_76(self):
        fonctionnel.demarrer_tnrSISTER(self, 'TNR-76', '/tnr76_', '76')

        fn_tnr7_256.naviguer_vers_Accueil("//div[@title = 'Outils']", "//a[contains(@title, 'Etats d')]")
        fn_tnr7_256.alimentation_des_champs('07/11/20', 'Après', 'Product Hub', "Champ flexible extensible d'article", '')
        util.driver.get_screenshot_as_file(util.Directory + "/Resultat_check.png")
        fonctionnel.vers_Word("Verification: Apparition de l'article modifié dans la partie 'Résultats de la recherche'", util.Directory + "/Resultat_check.png")

        fn_tnr7_256.cocher_tous_les_attributs(self)

        # Vérifier que tous les articles créés et modifiés pour les tests de non regression de la release apparaissent
        util.driver.get_screenshot_as_file(util.Directory + "/Tout_attributs_check.png")
        fonctionnel.vers_Word("Vérification: Visibilité des champs(Auteur, Date/Heure, Ancienne/Nouvelle valeur)", util.Directory + "/Tout_attributs_check.png")

        #fn_tnr7_256.extraction_des_donnees_auto(self)
        fn_tnr7_256.telecharger_les_donnees(self)

    def Modification_article_en_masse_RGA_86(self):
        fonctionnel.demarrer_tnrSISTER(self, 'TNR-86', '/tnr86_', '86')
        fn_tnr69_8_67.gestion_des_articles(self, True, '86')
        fn_tnr69_8_67.reaffectation_lot_article_au_bon_utilisateur(self, 'RGA', '86')

        fn_tnr69_8_67.clique_Action(self, '86')
        fn_tnr69_8_67.purger_toutes_les_lignes(self, True, '86')

        fn_tnr69_8_67.parametrage(self, False)
        fn_tnr69_8_67.clique_Action(self,'86')

        fn_tnr69_8_67.chargement_fichier(self, '86', 'TEST_597_TNR_86_Modif_Masse_RGA.csv', 'test86')
        fn_tnr69_8_67.consulter_les_articles(self, '231507', False)
        fn_tnr69_8_67.verification_Tx_Depre(self, '86', '0', '231507')
        fn_tnr69_8_67.verification_Tx_Depre(self, '86', '1', '231507')
        fn_tnr69_8_67.consulter_les_articles(self, '228648', True)
        fn_tnr69_8_67.verification_Tx_Depre(self, '86', '2', '228648')
        fn_tnr69_8_67.verification_Tx_Depre(self, '86', '3', '228648')

    def Modification_article_en_masse_CDG_87(self):
        fonctionnel.demarrer_tnrSISTER(self, 'TNR-87', '/tnr87_', '87')
        fn_tnr69_8_67.gestion_des_articles(self, True, '87')
        fn_tnr69_8_67.reaffectation_lot_article_au_bon_utilisateur(self, 'CDG','87')
        fn_tnr69_8_67.clique_Action(self, '87')

        fn_tnr69_8_67.purger_toutes_les_lignes(self, True, '87')
        fn_tnr69_8_67.parametrage(self, False)
#
        fn_tnr69_8_67.clique_Action(self, '87')

        fn_tnr69_8_67.chargement_fichier(self, '87', 'TEST_599_Modif_Tx_Depre (1).csv', 'test87')
        fn_tnr69_8_67.consulter_les_articles(self, '230130', False)
        fn_tnr69_8_67.verification_Tx_Depre(self, '87', '0', '230130')
        fn_tnr69_8_67.verification_Tx_Depre(self, '87', '1', '230130')

    def Avancement_du_workflow_par_le_Dashboard_90(self):
        fonctionnel.demarrer_tnrSISTER(self, 'TNR-90', '/tnr90_', '90')

        fn_tnr90.acceder_aux_dossiers_partages(self)
        fn_tnr90.chercher_element_personalise(self)
        fn_tnr90.Dashboard_Nir(self)

        fn_tnr90.assurer_conformite_de_la_page(self)
        fn_tnr90.checker_les_3_parties_de_la_page(self)

        fn_tnr90.conversion_et_transfert_du_fichier_xlsx(self)
        fn_tnr90.choisir_mois_concerne(self)

        fn_tnr90.switch_frame(self)
        fn_tnr90.semaine_du_calendrier('d_1', '_6_1')

        fn_tnr90.checker_mois_et_fin_de_semaine_actuelle(self)
        fn_tnr90.checker_semaine_attendue_actuelle(self)

    def Contrôle_matrice_article_94(self):
        fonctionnel.demarrer_tnrSISTER(self, 'TNR-94', '/tnr94_', '94')
        #fonctionnel.gestion_des_articles(self)
        #fn_tnr94.ajouter_date_de_creation(self)
        #fn_tnr94.alimenter_les_champs(self)
        #fn_tnr94.test(self, 'MOBILE GP')
        fn_tnr94.test(self, 'MULTIMEDIA')

    def tearDown(self):
        # close the browser window
        util.driver.quit()

    if __name__ == '__main__':
       unittest.main()
