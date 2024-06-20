import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.common.action_chains import ActionChains
import sys
import random
import string
import re
from string import *
from random import randint
from selenium.common.exceptions import *
from os import path
import docx
from datetime import datetime

import Utilities as util
from Utilities import *
import HTMLTestRunner
import unittest
from all_tnr import *

class SISTERTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.driver = webdriver.Chrome("C:\webdrivers\chromedriver.exe")
    #    cls.driver.get("https://x-x.x.x.xx.com")
    #    cls.driver.maximize_window()
        cls.driver.quit()

    #def test_Login(self):
    #    driver = self.driver
    #    wait = WebDriverWait(driver, 20)
    #    driver.get("https://xxx")
    #    language = wait.until(EC.presence_of_element_located((By.NAME, 'Languages')))
    #    if language != "French - français":
    #        language.click()
    #        fr = wait.until(EC.element_to_be_clickable((By.XPATH, "//option[@value='fr-fr']")))
    #        fr.click()
    #    else:
    #        pass
    #    userid = wait.until(EC.presence_of_element_located((By.NAME, 'userid')))
    #    userid.send_keys('x')
    #    driver.find_element_by_xpath("//*[@name='password']").send_keys('xxx!')
    #    connex = wait.until(EC.presence_of_element_located((By.XPATH, "//button[@id = 'btnActive']")))
    #    connex.click()

    #def test_tnrCA(self):
    #    ensemble_des_tnr.creation_article(self)

    #def test_tnr69(self):
    #    ensemble_des_tnr.Création_article_en_masse_69(self)

    #def test_tnr71(self):
    #    ensemble_des_tnr.Unicité_des_champs_de_la_fiche_article_71(self)

    #def test_tnr72(self):
    #    ensemble_des_tnr.Format_du_champ_EAN13_de_la_fiche_article_72(self)
#
    #def test_tnr75(self):
    #    ensemble_des_tnr.Historique_de_modication_de_fiche_article_Code_court_75(self)
#
    #def test_tnr76(self):
    #    ensemble_des_tnr.Historique_de_modication_de_fiche_article_Date_76(self)
#
    # Le taux de depre doit rester inchange apres une nouvelle modification pour le tnr 87
    #def test_tnr86(self):
    #    ensemble_des_tnr.Modification_article_en_masse_RGA_86(self)
#
    ## Modifier le taux de depre dans le fichier TEST_599_Modif_Tx_Depre (1).csv pour confirmer si le changement a bien ete fait.
    #def test_tnr87(self):
    #    ensemble_des_tnr.Modification_article_en_masse_CDG_87(self)
##
    #def test_tnr90(self):
    #    ensemble_des_tnr.Avancement_du_workflow_par_le_Dashboard_90(self)

    def test_tnr94(self):
        ensemble_des_tnr.Contrôle_matrice_article_94(self)

    @classmethod
    def tearDownClass(cls):
        # close the browser window
        cls.driver.quit()

if __name__ == '__main__':
    unittest.main(testRunner=HTMLTestRunner.HTMLTestRunner())
