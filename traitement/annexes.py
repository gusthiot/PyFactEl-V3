#import os
#import shutil

from outils import Outils
from latex import Latex


class Annexes(object):
    """
    Classe pour la création des annexes
    """
# TODO : vérifier longueur des tableaux
# TODO : vérifier que données des csv "commentaires" pas utilisées
    @staticmethod
    def annexes(sommes, clients, edition, livraisons, acces, machines, reservations, prestations, comptes,
                dossier_annexe, plateforme, coefprests, coefmachines, generaux):
        """
        création des annexes de facture
        :param sommes: sommes calculées
        :param clients: clients importés
        :param edition: paramètres d'édition
        :param livraisons: livraisons importées
        :param acces: accès importés
        :param machines: machines importées
        :param reservations: réservations importées
        :param prestations: prestations importées
        :param comptes: comptes importés
        :param dossier_annexe: nom du dossier dans lequel enregistrer le dossier des annexes
        :param plateforme: OS utilisé
        :param coefprests: coefficients prestations importés
        :param coefmachines: coefficients machines importés
        :param generaux: paramètres généraux
        """
        prefixe = "annexe_"
        garde = r'''Annexes factures \newline Billing Appendices'''

        Annexes.creation_annexes(sommes, clients, edition, livraisons, acces, machines, reservations, prestations,
                                 comptes, dossier_annexe, plateforme, prefixe, coefprests, coefmachines, generaux, garde)

    @staticmethod
    def annexes_techniques(sommes, clients, edition, livraisons, acces, machines, reservations, prestations, comptes,
                           dossier_annexe, plateforme, coefprests, coefmachines, generaux):
        """
        création des annexes techniques
        :param sommes: sommes calculées
        :param clients: clients importés
        :param edition: paramètres d'édition
        :param livraisons: livraisons importées
        :param acces: accès importés
        :param machines: machines importées
        :param reservations: réservations importées
        :param prestations: prestations importées
        :param comptes: comptes importés
        :param dossier_annexe: nom du dossier dans lequel enregistrer le dossier des annexes
        :param plateforme: OS utilisé
        :param coefprests: coefficients prestations importés
        :param coefmachines: coefficients machines importés
        :param generaux: paramètres généraux
        """
        prefixe = "annexeT_"
        garde = r'''Annexes techniques \newline Technical Appendices'''

        Annexes.creation_annexes(sommes, clients, edition, livraisons, acces, machines, reservations, prestations,
                                 comptes, dossier_annexe, plateforme, prefixe, coefprests, coefmachines, generaux, garde)

    @staticmethod
    def creation_annexes(sommes, clients, edition, livraisons, acces, machines, reservations, prestations, comptes,
                         dossier_annexe, plateforme, prefixe, coefprests, coefmachines, generaux, garde):
        """
        création des annexes techniques
        :param sommes: sommes calculées
        :param clients: clients importés
        :param edition: paramètres d'édition
        :param livraisons: livraisons importées
        :param acces: accès importés
        :param machines: machines importées
        :param reservations: réservations importées
        :param prestations: prestations importées
        :param comptes: comptes importés
        :param dossier_annexe: nom du dossier dans lequel enregistrer les annexes
        :param plateforme: OS utilisé
        :param prefixe: prefixe de nom des annexes
        :param coefprests: coefficients prestations importés
        :param coefmachines: coefficients machines importés
        :param generaux: paramètres généraux
        :param garde: titre page de garde
        """

        if sommes.calculees == 0:
            info = "Vous devez d'abord faire toutes les sommes avant de pouvoir créer les annexes"
            print(info)
            Outils.affiche_message(info)
            return

        for code_client in sommes.sommes_clients.keys():
            contenu = Latex.entete(plateforme)
            contenu += r'''
                \usepackage[margin=10mm, includehead]{geometry}
                \usepackage{multirow}
                \usepackage{graphicx}
                \usepackage{longtable}
                \usepackage{dcolumn}
                \usepackage{changepage}
                \usepackage[scriptsize]{caption}
                \usepackage{fancyhdr}
                \pagestyle{fancy}

                \fancyhead{}
                \fancyfoot{}

                \renewcommand{\headrulewidth}{0pt}
                \renewcommand{\footrulewidth}{0pt}

                \fancyhead[L]{\leftmark \\ \rightmark}
                \fancyhead[R]{\thepage}

                \newcommand{\fakesection}[2]{
                    \markboth{#1}{#2}
                }

                \begin{document}
                \renewcommand{\arraystretch}{1.5}
                '''
            client = clients.donnees[code_client]
            nature = generaux.nature_client_par_code_n(client['type_labo'])
            reference = nature + str(edition.annee)[2:] + Outils.mois_string(edition.mois) + "." + code_client
            if edition.version != "0":
                reference += "-" + edition.version

            contenu += r'''
                \begin{titlepage}
                \vspace*{8cm}
                \begin{adjustwidth}{5cm}{}
                \Large\textsc{''' + garde + r'''}\newline
                \Large\textsc{''' + reference + r'''}\newline\newline\newline
                '''

            dic_entete = {'code': code_client, 'code_sap': client['code_sap'],
                          'nom': Latex.echappe_caracteres(client['abrev_labo']),
                          'date': edition.mois_txt + " " + str(edition.annee)}

            contenu += r'''Client %(code)s -  %(code_sap)s -  %(nom)s \newline
                 %(date)s
                \end{adjustwidth}
                \end{titlepage}''' % dic_entete
            contenu += Annexes.contenu_client(sommes, clients, code_client, edition, livraisons, acces, machines,
                                              reservations, prestations, comptes, coefprests, coefmachines, generaux)
            contenu += r'''\end{document}'''

            nom = prefixe + str(edition.annee) + "_" + Outils.mois_string(edition.mois) + "_" + \
                  str(edition.version) + "_" + code_client

            Latex.creer_latex_pdf(nom, contenu, dossier_annexe)

    @staticmethod
    def contenu_client(sommes, clients, code_client, edition, livraisons, acces, machines, reservations, prestations,
                       comptes, coefprests, coefmachines, generaux):
        """
        création du contenu de l'annexe pour un client
        :param sommes: sommes calculées
        :param clients: clients importés
        :param code_client: code du client pour l'annexe
        :param edition: paramètres d'édition
        :param livraisons: livraisons importées
        :param acces: accès importés
        :param machines: machines importées
        :param reservations: réservations importées
        :param prestations: prestations importées
        :param comptes: comptes importés
        :param coefprests: coefficients prestations importés
        :param coefmachines: coefficients machines importés
        :param generaux: paramètres généraux
        :return: contenu de l'annexe du client
        """

        contenu = ""

        scl = sommes.sommes_clients[code_client]
        client = clients.donnees[code_client]
        nature = generaux.nature_client_par_code_n(client['type_labo'])
        av_ds = generaux.avantage_ds_par_code_n(client['type_labo'])
        av_hc = generaux.avantage_hc_par_code_n(client['type_labo'])
        reference = nature + str(edition.annee)[2:] + Outils.mois_string(edition.mois) + "." + code_client
        if edition.version != "0":
            reference += "-" + edition.version

        contenu_bonus_compte = ""
        contenu_procedes_compte = ""
        contenu_recap_compte = ""
        contenu_fact_compte = ""
        inc_fact = 1

        contenu_prestations_client_tab = {}
        for article in generaux.articles_d3:
            contenu_prestations_client_tab[article.code_d] = ""

        client_comptes = sommes.sommes_comptes[code_client]
        titre_4 = "Annexe détaillée par compte"
        nombre_4 = "IV"
        titre_2 = "Récapitulatifs par compte"
        nombre_2 = "II"
        contenu_compte_annexe2 = ""
        contenu_compte_annexe4 = ""
        contenu_compte_annexe5 = ""

        for id_compte in sorted(client_comptes.keys()):

            # ## COMPTE

            sco = sommes.sommes_comptes[code_client][id_compte]
            compte = comptes.donnees[id_compte]
            intitule_compte = id_compte + " - " + Latex.echappe_caracteres(compte['intitule'])

            titre2 = titre_2 + " : " + id_compte
            contenu_compte_annexe2 += Annexes.section(code_client, client, edition, reference, titre2, nombre_2)

            titre4 = titre_4 + " : " + id_compte
            contenu_compte_annexe4 += Annexes.section(code_client, client, edition, reference, titre4, nombre_4)

            # ## ligne 1.1

            if sco['si_facture'] > 0:
                poste = inc_fact * 10
                intitule = intitule_compte + " - " + generaux.articles[2].intitule_long
                dico_fact_compte = {'intitule': intitule, 'poste': str(poste), 'mm': "%.2f" % sco['somme_j_mm'],
                                    'mr': "%.2f" % sco['somme_j_mr'], 'mj': "%.2f" % sco['mj']}
                contenu_fact_compte += r'''
                    %(poste)s & %(intitule)s & %(mm)s  & %(mr)s & %(mj)s \\
                    \hline
                    ''' % dico_fact_compte
                poste += 1

                for article in generaux.articles_d3:
                    categorie = article.code_d
                    if sco['sommes_cat_m'][categorie] > 0:
                        intitule = intitule_compte + " - " + Latex.echappe_caracteres(article.intitule_long)
                        dico_fact_compte = {'intitule': intitule, 'poste': str(poste),
                                            'mm': "%.2f" % sco['sommes_cat_m'][article.code_d],
                                            'mr': "%.2f" % sco['sommes_cat_r'][article.code_d],
                                            'mj': "%.2f" % sco['tot_cat'][article.code_d]}
                        contenu_fact_compte += r'''
                            %(poste)s & %(intitule)s & %(mm)s  & %(mr)s & %(mj)s \\
                            \hline
                            ''' % dico_fact_compte
                        poste += 1

                inc_fact += 1

            # ## ligne 1.5

            total = sco['mj']
            dico_recap_compte = {'compte': intitule_compte, 'procede': "%.2f" % sco['mj']}

            ligne = r'''%(compte)s & %(procede)s ''' % dico_recap_compte

            for categorie in generaux.codes_d3():
                total += sco['tot_cat'][categorie]
                ligne += r''' & ''' + "%.2f" % sco['tot_cat'][categorie]

            if total > 0:
                dico_recap_compte['total'] = "%.2f" % total
                ligne += r'''& %(total)s \\
                    \hline
                    ''' % dico_recap_compte
                contenu_recap_compte += ligne

            # ## 1.6

            rsj = client['rs'] * sco['somme_j_dsi']
            rhj = client['rh'] * sco['somme_j_dhi']
            dico_procedes_compte = {'intitule': intitule_compte, 'maij': "%.2f" % sco['somme_j_mai'],
                                    'mm': "%.2f" % sco['somme_j_mm'], 'mr': "%.2f" % sco['somme_j_mr'],
                                    'rsj': "%.2f" % rsj, 'rhj': "%.2f" % rhj,
                                    'moij': "%.2f" % sco['somme_j_moi'], 'mj': "%.2f" % sco['mj']}
            contenu_procedes_compte += r'''
                %(intitule)s & %(maij)s & %(moij)s & %(rsj)s & %(rhj)s & %(mm)s & %(mr)s & %(mj)s \\
                \hline
                ''' % dico_procedes_compte

            # ## ligne 1.7

            if code_client in livraisons.sommes and id_compte in livraisons.sommes[code_client]:
                for article in generaux.articles_d3:
                    if article.code_d in livraisons.sommes[code_client][id_compte]:
                        if contenu_prestations_client_tab[article.code_d] == "":
                            contenu_prestations_client_tab[article.code_d] = r'''
                                \cline{2-4}
                                \multicolumn{1}{c}{} & \multicolumn{3}{|c|}{''' + article.intitule_long + r'''} \\
                                \hline
                                Compte & Montant & Rabais & Montant net \\
                                \hline
                                '''
                        dico_prestations_client = {'intitule': intitule_compte,
                                                   'cmj': "%.2f" % sco['sommes_cat_m'][article.code_d],
                                                   'crj': "%.2f" % sco['sommes_cat_r'][article.code_d],
                                                   'cj': "%.2f" % sco['tot_cat'][article.code_d]}
                        contenu_prestations_client_tab[article.code_d] += r'''
                        %(intitule)s & %(cmj)s & %(crj)s & %(cj)s \\
                        \hline
                        ''' % dico_prestations_client

            # ## ligne 1.8

            if code_client in acces.sommes and id_compte in acces.sommes[code_client]['comptes']:
                bsj = client['bs'] * sco['somme_j_dsi']
                bhj = client['bh'] * sco['somme_j_dhi']
                dico_bonus_compte = {'compte': intitule_compte, 'bsj': "%.2f" % bsj, 'bhj': "%.2f" % bhj,
                                     'tot': "%.2f" % sco['somme_j_mb']}
                contenu_bonus_compte += r'''
                    %(compte)s & %(bsj)s & %(bhj)s & %(tot)s \\
                    \hline
                    ''' % dico_bonus_compte

            # ## 2.1

            structure_recap_poste = r'''{|l|c|c|c|}'''
            legende_recap_poste = r'''Table II.1 - Récapitulatif du compte'''

            dico_recap_poste = {'mm': "%.2f" % sco['somme_j_mm'], 'mr': "%.2f" % sco['somme_j_mr'],
                                'maij': "%.2f" % sco['somme_j_mai'], 'dsij': "%.2f" % sco['somme_j_dsi'],
                                'dhij': "%.2f" % sco['somme_j_dhi'], 'mj': "%.2f" % sco['mj'],
                                'nmij': "%.2f" % (sco['somme_j_mai']-sco['somme_j_mr']),
                                'moij': "%.2f" % sco['somme_j_moi'], 'int_proc': generaux.articles[2].intitule_long}

            contenu_recap_poste = r'''
                \cline{2-4}
                \multicolumn{1}{r|}{} & Montant & Rabais & Net \\
                \hline
                Machine & %(maij)s & %(mr)s  & %(nmij)s \\
                \hline
                Main d'oeuvre & %(moij)s & & %(moij)s  \\
                \hline
                %(int_proc)s & %(mm)s & %(mr)s & %(mj)s \\
                \hline
                ''' % dico_recap_poste

            total = sco['mj']
            for article in generaux.articles_d3:
                total += sco['tot_cat'][article.code_d]
                dico_recap_poste = {'intitule': Latex.echappe_caracteres(article.intitule_long),
                                    'cmj': "%.2f" % sco['sommes_cat_m'][article.code_d],
                                    'crj': "%.2f" % sco['sommes_cat_r'][article.code_d],
                                    'cj': "%.2f" % sco['tot_cat'][article.code_d]}
                contenu_recap_poste += r'''
                %(intitule)s & %(cmj)s & %(crj)s & %(cj)s \\
                \hline
                ''' % dico_recap_poste

            contenu_recap_poste += r'''\multicolumn{3}{|r|}{Total} & ''' + "%.2f" % total + r'''\\
                \hline
                '''

            contenu_compte_annexe2 += Latex.tableau(contenu_recap_poste, structure_recap_poste, legende_recap_poste)

            # ## 2.2

            if code_client in acces.sommes and id_compte in acces.sommes[code_client]['comptes']:
                structure_utilise_compte = r'''{|l|c|c|c|c|c|c|c|c|c|}'''
                legende_utilise_compte = r'''Table II.2 - Procédés (machine + main d'oeuvre)'''

                contenu_utilise_compte = r'''
                    \cline{3-10}
                    \multicolumn{2}{c}{} & \multicolumn{2}{|c|}{Machine} & \multicolumn{2}{l|}{PU [CHF/h]} & \multicolumn{2}{l|}{Montant [CHF]} & Déduc. Sp. & Déduc. HC \\
                    \hline
                    \multicolumn{2}{|l|}{\textbf{''' + intitule_compte + r'''}} & Mach. & Oper. & Mach. & Oper. & Mach. & Oper. & ''' + av_ds + r''' & ''' + av_hc + r''' \\
                    \hline
                    '''

                somme = acces.sommes[code_client]['comptes'][id_compte]

                machines_utilisees = {}
                for key in somme:
                    id_cout = machines.donnees[key]['id_cout']
                    nom = machines.donnees[key]['nom']
                    if id_cout not in machines_utilisees:
                        machines_utilisees[id_cout] = {}
                    machines_utilisees[id_cout][nom] = key

                for id_cout, mics in sorted(machines_utilisees.items()):
                    for nom, id_machine in sorted(mics.items()):
                        dico_machine = {'machine': Latex.echappe_caracteres(nom),
                                        'hp': Outils.format_heure(somme[id_machine]['duree_hp']),
                                        'hc': Outils.format_heure(somme[id_machine]['duree_hc']),
                                        'mo_hp': Outils.format_heure(somme[id_machine]['mo_hp']),
                                        'mo_hc': Outils.format_heure(somme[id_machine]['mo_hc']),
                                        'pu_m': "%.2f" % somme[id_machine]['pum'],
                                        'puo_hp': "%.2f" % somme[id_machine]['puo_hp'],
                                        'puo_hc' : "%.2f" % somme[id_machine]['puo_hc'],
                                        'mai_hp': "%.2f" % somme[id_machine]['mai_hp'],
                                        'mai_hc': "%.2f" % somme[id_machine]['mai_hc'],
                                        'moi_hp': "%.2f" % somme[id_machine]['moi_hp'],
                                        'moi_hc': "%.2f" % somme[id_machine]['moi_hc'],
                                        'dsi_hp': "%.2f" % somme[id_machine]['dsi_hp'],
                                        'dsi_hc': "%.2f" % somme[id_machine]['dsi_hc'],
                                        'dhi': "%.2f" % somme[id_machine]['dhi']}

                        if somme[id_machine]['duree_hp'] > 0 or somme[id_machine]['mo_hp'] > 0:
                            contenu_utilise_compte += r'''
                                %(machine)s & HP & %(hp)s & %(mo_hp)s & %(pu_m)s & %(puo_hp)s & %(mai_hp)s & %(moi_hp)s & %(dsi_hp)s & \\
                                \hline
                                ''' % dico_machine

                        if somme[id_machine]['duree_hc'] > 0 or somme[id_machine]['mo_hc'] > 0:
                            contenu_utilise_compte += r'''
                                %(machine)s & HC & %(hc)s & %(mo_hc)s & %(pu_m)s & %(puo_hc)s & %(mai_hc)s & %(moi_hc)s & %(dsi_hc)s & %(dhi)s \\
                                \hline
                                ''' % dico_machine

                dico_tot = {'maij': "%.2f" % sco['somme_j_mai'], 'moij': "%.2f" % sco['somme_j_moi'],
                            'dsij': "%.2f" % sco['somme_j_dsi'], 'dhij': "%.2f" % sco['somme_j_dhi'],
                            'rabais': "%.2f" % sco['somme_j_mr'], 'bonus': "%.2f" % sco['somme_j_mb']}
                contenu_utilise_compte += r'''
                    \multicolumn{6}{|r|}{Total} & %(maij)s & %(moij)s & %(dsij)s & %(dhij)s \\
                    \hline
                    \multicolumn{6}{r}{} & \multicolumn{2}{|c|}{Rabais total} & \multicolumn{2}{c|}{%(rabais)s} \\
                    \cline{7-10}
                    \multicolumn{6}{r}{} & \multicolumn{2}{|c|}{Bonus total} & \multicolumn{2}{c|}{%(bonus)s} \\
                    \cline{7-10}
                    ''' % dico_tot

                contenu_compte_annexe2 += Latex.tableau(contenu_utilise_compte, structure_utilise_compte, legende_utilise_compte)
            else:
                contenu_compte_annexe2 += r'''
                    \tiny{Table II.1 - Procédés (machine + main d'oeuvre) : table vide (pas d’utilisation machines)}
                    \newline
                    '''

            # ## 2.3

            if code_client in livraisons.sommes and id_compte in livraisons.sommes[code_client]:
                somme = livraisons.sommes[code_client][id_compte]
                structure_prests_compte = r'''{|l|c|c|c|c|c|}'''
                legende_prests_compte = r'''Table II.3 - Prestations livrées'''
                contenu_prests_compte = ""
                for article in generaux.articles_d3:
                    if article.code_d in somme:
                        if contenu_prests_compte != "":
                            contenu_prests_compte += r'''
                                \multicolumn{6}{c}{} \\
                                '''

                        contenu_prests_compte += r'''
                            \hline
                            \multicolumn{1}{|l|}{\scriptsize{\textbf{''' + intitule_compte + " - " + article.intitule_long + r'''
                            }}} & Quantité & Unité & P.U. & Montant & Rabais \\
                            \hline
                            '''
                        for no_prestation, sip in sorted(somme[article.code_d].items()):
                            dico_prestations = {'nom': Latex.echappe_caracteres(sip['nom']), 'num': no_prestation,
                                                'quantite': sip['quantite'], 'unite': sip['unite'],
                                                'pu': "%.2f" % sip['pu'], 'montant': "%.2f" % sip['montant'],
                                                'rabais': "%.2f" % sip['rabais']}
                            contenu_prests_compte += r'''
                                %(num)s - %(nom)s & \hspace{5mm} %(quantite)s & %(unite)s & %(pu)s & %(montant)s & %(rabais)s  \\
                                \hline
                                ''' % dico_prestations
                        dico_prestations = {'montant': "%.2f" % sco['sommes_cat_m'][article.code_d],
                                            'rabais': "%.2f" % sco['sommes_cat_r'][article.code_d]}
                        contenu_prests_compte += r'''
                            \multicolumn{4}{|r|}{Total} & %(montant)s & %(rabais)s  \\
                            \hline
                            ''' % dico_prestations
                contenu_compte_annexe2 += Latex.tableau(contenu_prests_compte , structure_prests_compte ,
                                 legende_prests_compte )
            else:
                contenu_compte_annexe2 += r'''
                    \tiny{Table II.3 - Prestations livrées : table vide (pas de prestations livrées)}
                    \newline
                    '''

            # ## 4.1

            if code_client in acces.sommes and id_compte in acces.sommes[code_client]['comptes']:

                structure_machuts_compte = r'''{|l|l|l|c|c|c|c|}'''
                legende_machuts_compte = r'''Table IV.1 - Détails des utilisations machines'''

                contenu_machuts_compte = r'''
                    \hline
                    \multicolumn{3}{|l|}{\multirow{2}{*}{\scriptsize{\textbf{''' + intitule_compte + r'''}}}} & \multicolumn{2}{c|}{Machine} & \multicolumn{2}{c|}{Main d'oeuvre} \\
                    \cline{4-7}
                    \multicolumn{3}{|l|}{} & HP & HC & HP & HC \\
                    \hline
                    '''

                somme = acces.sommes[code_client]['comptes'][id_compte]

                machines_utilisees= {}
                for key in somme:
                    id_cout = machines.donnees[key]['id_cout']
                    nom = machines.donnees[key]['nom']
                    if id_cout not in machines_utilisees:
                        machines_utilisees[id_cout] = {}
                    machines_utilisees[id_cout][nom] = key

                for id_cout, mics in sorted(machines_utilisees.items()):
                    for nom, id_machine in sorted(mics.items()):

                        dico_machine = {'machine': Latex.echappe_caracteres(nom),
                                        'hp': Outils.format_heure(somme[id_machine]['duree_hp']),
                                        'hc': Outils.format_heure(somme[id_machine]['duree_hc']),
                                        'mo_hp': Outils.format_heure(somme[id_machine]['mo_hp']),
                                        'mo_hc': Outils.format_heure(somme[id_machine]['mo_hc'])}
                        contenu_machuts_compte += r'''
                            \multicolumn{3}{|l|}{\textbf{%(machine)s}} & \hspace{5mm} %(hp)s & \hspace{5mm} %(hc)s &
                             \hspace{5mm} %(mo_hp)s & \hspace{5mm} %(mo_hc)s \\
                            \hline
                            ''' % dico_machine

                        users = {}
                        for key in somme[id_machine]['users']:
                            prenom = somme[id_machine]['users'][key]['prenom']
                            nom = somme[id_machine]['users'][key]['nom']
                            if nom not in users:
                                users[nom] = {}
                            if prenom not in users[nom]:
                                users[nom][prenom] = []
                            users[nom][prenom].append(key)

                        for nom, upi in sorted(users.items()):
                            for prenom, ids in sorted(upi.items()):
                                for id_user in sorted(ids):
                                    smu = somme[id_machine]['users'][id_user]
                                    dico_user = {'user': nom + " " + prenom,
                                                 'hp': Outils.format_heure(smu['duree_hp']),
                                                 'hc': Outils.format_heure(smu['duree_hc']),
                                                 'mo_hp': Outils.format_heure(smu['mo_hp']),
                                                 'mo_hc': Outils.format_heure(smu['mo_hc'])}
                                    contenu_machuts_compte += r'''
                                        \multicolumn{3}{|l|}{\hspace{5mm} %(user)s} & %(hp)s & %(hc)s & %(mo_hp)s & %(mo_hc)s \\
                                        \hline
                                    ''' % dico_user
                                    for pos in smu['data']:
                                        cae = acces.donnees[pos]
                                        login = Latex.echappe_caracteres(cae['date_login']).split()
                                        temps = login[0].split('-')
                                        date = temps[0]
                                        for pos in range(1, len(temps)):
                                            date = temps[pos] + '.' + date
                                        if len(login) > 1:
                                            heure = login[1]
                                        else:
                                            heure = ""

                                        rem = ""
                                        if id_user != cae['id_op']:
                                            rem += "op : " + cae['nom_op']
                                        if cae['remarque_op'] != "":
                                            if rem != "":
                                                rem += "; "
                                            rem += "rem op : " + Latex.echappe_caracteres(cae['remarque_op'])
                                        if cae['remarque_staff'] != "":
                                            if rem != "":
                                                rem += "; "
                                            rem += "rem CMi : " + Latex.echappe_caracteres(cae['remarque_staff'])

                                        dico_pos = {'date': date, 'heure': heure, 'rem': rem,
                                                    'hp': Outils.format_heure(cae['duree_machine_hp']),
                                                    'hc': Outils.format_heure(cae['duree_machine_hc']),
                                                    'mo_hp': Outils.format_heure(cae['duree_operateur_hp']),
                                                    'mo_hc': Outils.format_heure(cae['duree_operateur_hc'])}
                                        contenu_machuts_compte += r'''
                                            \hspace{10mm} %(date)s & %(heure)s & %(rem)s & %(hp)s \hspace{5mm} &
                                             %(hc)s \hspace{5mm} & %(mo_hp)s \hspace{5mm} & %(mo_hc)s \hspace{5mm} \\
                                            \hline
                                        ''' % dico_pos

                contenu_compte_annexe4 += Latex.long_tableau(contenu_machuts_compte, structure_machuts_compte, legende_machuts_compte)
            else:
                contenu_compte_annexe4 += r'''
                    \tiny{Table IV.1 - Détails des utilisations machines : table vide (pas d’utilisation machines)}
                    \newline
                    '''

            # ## 4.2

            if code_client in livraisons.sommes and id_compte in livraisons.sommes[code_client]:
                somme = livraisons.sommes[code_client][id_compte]

                structure_prestations_compte = r'''{|l|c|c|c|}'''
                legende_prestations_compte = r'''Table IV.2 - Détails des prestations livrées'''

                contenu_prestations_compte = r'''
                    '''

                i = 0
                for article in generaux.articles_d3:
                    if article.code_d in somme:
                        if i == 0:
                            i += 1
                        else:
                            contenu_prestations_compte += r'''\multicolumn{4}{c}{} \\
                                '''
                        contenu_prestations_compte += r'''
                            \hline
                            \multicolumn{1}{|l|}{\scriptsize{\textbf{''' + intitule_compte + " - " + article.intitule_long + r'''
                            }}} & Quantité & Unité & Rabais \\
                            \hline
                            '''
                        for no_prestation, sip in sorted(somme[article.code_d].items()):
                            dico_prestations = {'nom': Latex.echappe_caracteres(sip['nom']), 'num': no_prestation,
                                                'quantite': sip['quantite'], 'unite': sip['unite'],
                                                'rabais': "%.2f" % sip['rabais']}
                            contenu_prestations_compte += r'''
                                %(num)s - %(nom)s & \hspace{5mm} %(quantite)s & %(unite)s & \hspace{5mm} %(rabais)s  \\
                                \hline
                                ''' % dico_prestations

                            users = {}
                            for key in sip['users']:
                                prenom = sip['users'][key]['prenom']
                                nom = sip['users'][key]['nom']
                                if nom not in users:
                                    users[nom] = {}
                                if prenom not in users[nom]:
                                    users[nom][prenom] = []
                                users[nom][prenom].append(key)

                            for nom, upi in sorted(users.items()):
                                for prenom, ids in sorted(upi.items()):
                                    for id_user in sorted(ids):
                                        spu = sip['users'][id_user]
                                        dico_user = {'user': nom + " " + prenom, 'quantite': spu['quantite'],
                                                     'unite': sip['unite'], 'rabais': "%.2f" % spu['rabais']}
                                        contenu_prestations_compte += r'''
                                            \hspace{5mm} %(user)s & %(quantite)s & %(unite)s & %(rabais)s \\
                                            \hline
                                        ''' % dico_user

                                        for pos in spu['data']:
                                            liv = livraisons.donnees[pos]
                                            rem = ""
                                            dl = ""
                                            if liv['remarque'] != "":
                                                rem = "; Remarque : " + liv['remarque']
                                            if liv['date_livraison'] != "":
                                                dl = "Dt livraison: " + Latex.echappe_caracteres(liv['date_livraison']) + ";"
                                            dico_pos = {'date_liv': dl,
                                                        'quantite': liv['quantite'],'rabais': "%.2f" % liv['rabais_r'],
                                                        'id': liv['id_livraison'], 'unite': liv['unite'],
                                                        'responsable': Latex.echappe_caracteres(liv['responsable']),
                                                        'commande': Latex.echappe_caracteres(liv['date_commande']),
                                                        'remarque': Latex.echappe_caracteres(rem)}
                                            contenu_prestations_compte += r'''
                                                \hspace{10mm} %(date_liv)s N. livraison: %(id)s & %(quantite)s \hspace{5mm} & %(unite)s & %(rabais)s \hspace{5mm} \\

                                                \hspace{10mm} \scalebox{.8}{Commande: %(commande)s; Resp: %(responsable)s%(remarque)s} & & & \\
                                                \hline
                                            ''' % dico_pos

                contenu_compte_annexe4 += Latex.long_tableau(contenu_prestations_compte, structure_prestations_compte, legende_prestations_compte)

            # ## 5.x ? coûts éligibles

            # structure_eligibles_compte = r'''{|c|c|c|c|}'''
            # for i in range(1,4):
            #     legende_eligibles_compte = r'''Coûts d'utilisation (U''' + str(i) + r''') pour compte ''' + intitule_compte
            #
            #     contenu_eligibles_compte = r'''
            #         \hline
            #         Mach. U''' + str(i) + r''' & M.O. & Prest. Livr. & Total \\
            #         \hline
            #         '''
            #
            #     contenu_compte_annexe5 += Latex.tableau(contenu_eligibles_compte, structure_eligibles_compte, legende_eligibles_compte)

            # ## 5.x ? coûts éligibles machines

            # structure_machines_compte = r'''{|l|c|c|c|c|c|c|c|}'''
            # for i in range(1,4):
            #     legende_machines_compte = r'''Coûts procédés éligibles (U''' + str(i) + r''') pour compte ''' + intitule_compte
            #
            #     contenu_machines_compte = r'''
            #         \cline{3-8}
            #         \multicolumn{2}{r}{} & \multicolumn{2}{|c|}{Machine} & \multicolumn{2}{c|}{PU [CHF/h]} & \multicolumn{2}{c|}{Montant [CHF]} \\
            #         \hline
            #         ''' + intitule_compte + r''' & Mach. & Oper. & PU''' + str(i) + r''' & M.O. & Mach.& U''' + str(i) + r''' & M.O. \\
            #         \hline
            #         '''
            #
            #     contenu_machines_compte += r'''
            #         \multicolumn{6}{|r|}{Total} & & \\
            #         \hline
            #         '''
            #
            #     contenu_compte_annexe5 += Latex.tableau(contenu_machines_compte, structure_machines_compte, legende_machines_compte)

            # ## 5.x ? coûts éligibles prestations

            # structure_livraisons_compte = r'''{|l|c|}'''
            # legende_livraisons_compte = r'''Prestations livrées éligibles pour compte ''' + intitule_compte
            #
            # contenu_livraisons_compte = r'''
            #     \hline
            #     '''
            # total = 0
            # for article in generaux.articles_d3:
            #     total += sco['tot_cat'][article.code_d]
            #     contenu_livraisons_compte += r'''
            #         ''' + article.intitule_long + r'''&''' + "%.2f" % sco['tot_cat'][article.code_d] + r''' \\
            #         \hline
            #         '''
            # contenu_livraisons_compte += r'''
            #     Total Pres. Livr. &''' + "%.2f" % total + r''' \\
            #     \hline
            #     '''
            #
            # contenu_compte_annexe5 += Latex.tableau(contenu_livraisons_compte, structure_livraisons_compte, legende_livraisons_compte)

            contenu_compte_annexe2 += r'''\clearpage'''
            contenu_compte_annexe4 += r'''\clearpage'''
            contenu_compte_annexe5 += r'''\clearpage'''
            # ## compte

        # ## Début des tableaux

        # ## Annexe 1

        contenu += Annexes.titre_annexe(code_client, client, edition, reference, "Récapitulatif", "I")
        contenu += Annexes.section(code_client, client, edition, reference, "Récapitulatif", "I")

        # ## 1.1

        structure_recap_fact = r'''{|c|l|c|c|c|}'''
        legende_recap_fact = r'''Table I.1 - Récapitulatif des postes de la facture'''

        dico_recap_fact = {'emom': "%.2f" % scl['em'], 'emor': "%.2f" % scl['er'], 'emo': "%.2f" % scl['e'],
                           'resm': "%.2f" % scl['rm'], 'resr': "%.2f" % scl['rr'], 'res': "%.2f" % scl['r'],
                           'int_emo': generaux.articles[0].intitule_long,
                           'int_res': generaux.articles[1].intitule_long,
                           'p_emo': generaux.poste_emolument, 'p_res': generaux.poste_reservation}

        contenu_recap_fact = r'''
            \hline
            N. Poste & Poste & \multicolumn{1}{l|}{Montant} & \multicolumn{1}{l|}{Rabais} & \multicolumn{1}{l|}{Total} \\
            \hline
            %(p_emo)s & %(int_emo)s & %(emom)s & %(emor)s & %(emo)s \\
            \hline
            %(p_res)s & %(int_res)s & %(resm)s & %(resr)s & %(res)s \\
            \hline
            ''' % dico_recap_fact

        contenu_recap_fact += contenu_fact_compte

        contenu += Latex.tableau(contenu_recap_fact, structure_recap_fact, legende_recap_fact)

        # ## 1.2

        structure_recap_poste_cl = r'''{|l|c|c|c|}'''
        legende_recap_poste_cl = r'''Table I.2 - Récapitulatif des postes'''

        dico_recap_poste_cl = {'emom': "%.2f" % scl['em'], 'emor': "%.2f" % scl['er'], 'emo': "%.2f" % scl['e'],
                               'resm': "%.2f" % scl['rm'], 'resr': "%.2f" % scl['rr'], 'res': "%.2f" % scl['r'],
                               'int_emo': generaux.articles[0].intitule_long,
                               'int_res': generaux.articles[1].intitule_long,
                               'int_proc': generaux.articles[2].intitule_long, 'mm': "%.2f" % scl['somme_t_mm'],
                               'mr': "%.2f" % scl['somme_t_mr'], 'mt': "%.2f" % scl['mt']}

        contenu_recap_poste_cl = r'''
            \cline{2-4}
            \multicolumn{1}{l|}{} & \multicolumn{1}{l|}{Montant} & \multicolumn{1}{l|}{Rabais} & \multicolumn{1}{l|}{Total} \\
            \hline
            %(int_emo)s & %(emom)s & %(emor)s & %(emo)s \\
            \hline
            %(int_res)s & %(resm)s & %(resr)s & %(res)s \\
            \hline
            %(int_proc)s & %(mm)s & %(mr)s & %(mt)s \\
            \hline
            ''' % dico_recap_poste_cl

        for article in generaux.articles_d3:
            dico_recap_poste_cl = {'intitule': Latex.echappe_caracteres(article.intitule_long),
                                'mm': "%.2f" % scl['sommes_cat_m'][article.code_d],
                                'mr': "%.2f" % scl['sommes_cat_r'][article.code_d],
                                'mj': "%.2f" % scl['tot_cat'][article.code_d]}
            contenu_recap_poste_cl += r'''
                %(intitule)s & %(mm)s  & %(mr)s & %(mj)s \\
                \hline
                ''' % dico_recap_poste_cl

        contenu_recap_poste_cl += r'''\multicolumn{3}{|r|}{Total} & ''' + "%.2f" % (scl['somme_t'] + scl['e']) + r'''\\
            \hline
            '''

        contenu += Latex.tableau(contenu_recap_poste_cl, structure_recap_poste_cl, legende_recap_poste_cl)

        # ## 1.3

        structure_emolument = r'''{|r|r|l|r|r|r|r|r|}'''
        legende_emolument = r'''Table I.3 - Emolument mensuel'''

        dico_emolument = {'emb':  "%.2f" % client['emol_base_mens'], 'ef':  "%.2f" % client['emol_fixe'],
                          'pente': client['coef'], 'tot_eq_r': "%.2f" % scl['r'],
                          'tot_eq_m': "%.2f" % scl['mat'], 'tot_eq': "%.2f" % scl['somme_eq'],
                          'rabais': "%.2f" % scl['er'], 'emo': scl['e']}

        contenu_emolument = r'''
            \hline
            \multicolumn{1}{|l|}{Emolument de base} & \multicolumn{1}{l|}{Emolument fixe} & Pente
            & \multicolumn{1}{l|}{Total EQ R} & \multicolumn{1}{l|}{Total EQ M} & \multicolumn{1}{l|}{Total EQ} &
            \multicolumn{1}{l|}{Rabais émolument} & \multicolumn{1}{l|}{Emolument} \\
            \hline
            %(emb)s & %(ef)s & %(pente)s & %(tot_eq_r)s & %(tot_eq_m)s & %(tot_eq)s & %(rabais)s & %(emo)s \\
            \hline
            ''' % dico_emolument

        contenu += Latex.tableau(contenu_emolument, structure_emolument, legende_emolument)

        # ## 1.4

        structure_frais_client = r'''{|l|c|c|c|c|}'''
        legende_frais_client = r'''Table I.4 - Pénalités réservation'''

        contenu_frais_client = r'''
            \cline{3-5}
            \multicolumn{2}{c|}{} & Pénalités & PU & Montant \\
            \cline{3-5}
            \multicolumn{2}{c|}{} & Durée & CHF/h & CHF \\
            \hline
            '''

        machines_utilisees = {}

        for key in scl['res']:
            id_cout = machines.donnees[key]['id_cout']
            nom = machines.donnees[key]['nom']
            if id_cout not in machines_utilisees:
                machines_utilisees[id_cout] = {}
            machines_utilisees[id_cout][nom] = key

        for id_cout, mics in sorted(machines_utilisees.items()):
            for nom_machine, id_machine in sorted(mics.items()):
                re_somme = reservations.sommes[code_client]['machines'][id_machine]

                tot_hp = scl['res'][id_machine]['tot_hp']
                tot_hc = scl['res'][id_machine]['tot_hc']

                dico_machine = {'machine': Latex.echappe_caracteres(nom_machine),
                                'pu_hp': re_somme['pu_hp'], 'pu_hc': re_somme['pu_hc'],
                                'mont_hp': "%.2f" % scl['res'][id_machine]['mont_hp'],
                                'mont_hc': "%.2f" % scl['res'][id_machine]['mont_hc'],
                                'tot_hp': Outils.format_heure(tot_hp), 'tot_hc': Outils.format_heure(tot_hc)}

                if tot_hp > 0:
                    contenu_frais_client += r'''%(machine)s & HP & %(tot_hp)s & %(pu_hp)s & %(mont_hp)s \\
                         \hline
                         ''' % dico_machine

                if tot_hc > 0:
                    contenu_frais_client += r'''%(machine)s & HC & %(tot_hc)s & %(pu_hc)s & %(mont_hc)s \\
                         \hline
                         ''' % dico_machine

        dico_frais = {'rm': scl['rm'], 'r': scl['r'], 'rr': scl['rr']}
        contenu_frais_client += r'''
            \multicolumn{4}{|r|}{Total} & %(rm)s \\
            \hline
            \multicolumn{4}{|r|}{Rabais} & %(rr)s \\
            \hline
            \multicolumn{4}{|r|}{\textbf{Total à payer}} & \textbf{%(r)s} \\
            \hline
            ''' % dico_frais

        contenu += Latex.tableau(contenu_frais_client, structure_frais_client, legende_frais_client)

        # ## 1.5

        legende_recap = r'''Table I.5 - Récapitulatif des comptes'''

        structure_recap = r'''{|l|r|r|'''
        contenu_recap = r'''
            \hline
            Compte & \multicolumn{1}{l|}{Procédés}'''

        for article in generaux.articles_d3:
            structure_recap += r'''r|'''
            contenu_recap += r''' & \multicolumn{1}{l|}{
            ''' + Latex.echappe_caracteres(article.intitule_court) + r'''}'''
        structure_recap += r'''}'''
        contenu_recap += r'''& \multicolumn{1}{l|}{Total} \\
            \hline
            '''

        contenu_recap += contenu_recap_compte

        dico_recap = {'procedes': "%.2f" % scl['mt'], 'total': "%.2f" % (scl['somme_t']-scl['r'])}

        contenu_recap += r'''Total article & %(procedes)s''' % dico_recap

        for categorie in generaux.codes_d3():
            contenu_recap += r''' & ''' + "%.2f" % scl['tot_cat'][categorie]

        contenu_recap += r'''& %(total)s \\
            \hline
            ''' % dico_recap

        contenu += Latex.tableau(contenu_recap, structure_recap, legende_recap)

        # ## 1.6

        if code_client in acces.sommes:
            structure_procedes_client = r'''{|l|c|c|c|c|c|c|c|}'''
            legende_procedes_client = r'''Table I.6 - Récapitulatif des procédés'''

            contenu_procedes_client = r'''
                \cline{2-8}
                \multicolumn{1}{c}{} & \multicolumn{2}{|c|}{Procédés} & \multicolumn{2}{c|}{Rabais} & \multicolumn{2}{c|}{Facture} & Montant \\
                \cline{1-7}
                Compte & Machine & M.O. opér. & Déduc. Sp. & Déduc. HC & Montant & Rabais & net \\
                \hline
                '''

            contenu_procedes_client += contenu_procedes_compte

            rst = client['rs'] * scl['dst']
            rht = client['rh'] * scl['dht']
            dico_procedes_client = {'mat': "%.2f" % scl['mat'], 'mr': "%.2f" % scl['somme_t_mr'],
                                    'mm': "%.2f" % scl['somme_t_mm'], 'mot': "%.2f" % scl['mot'],
                                    'mt': "%.2f" % scl['mt'], 'rst': "%.2f" % rst, 'rht': "%.2f" % rht}
            contenu_procedes_client += r'''
                Total & %(mat)s & %(mot)s & %(rst)s & %(rht)s & %(mm)s & %(mr)s & %(mt)s \\
                \hline
                ''' % dico_procedes_client

            contenu += Latex.tableau(contenu_procedes_client, structure_procedes_client, legende_procedes_client)
        else:
            contenu += r'''
                \tiny{Table I.6 - Récapitulatif des procédés : table vide (pas d'utilisations machines)}
                \newline
                '''

        # ## 1.7

        if code_client in livraisons.sommes:
            structure_prestations_client_recap = r'''{|l|c|c|c|}'''
            legende_prestations_client_recap = r'''Table I.7 - Récapitulatif des prestations livrées'''

            contenu_prestations_client_recap = ""
            i = 0
            for article in generaux.articles_d3:
                if contenu_prestations_client_tab[article.code_d] != "":
                    dico_prestations_client = {'cmt': "%.2f" % scl['sommes_cat_m'][article.code_d],
                                               'crt': "%.2f" % scl['sommes_cat_r'][article.code_d],
                                               'ct': "%.2f" % scl['tot_cat'][article.code_d]}
                    contenu_prestations_client_tab[article.code_d] += r'''
                    Total & %(cmt)s & %(crt)s & %(ct)s \\
                    \hline
                    ''' % dico_prestations_client
                    if i == 0:
                        i += 1
                    else:
                        contenu_prestations_client_recap += r'''\multicolumn{4}{c}{} \\'''
                    contenu_prestations_client_recap += contenu_prestations_client_tab[article.code_d]

            contenu += Latex.tableau(contenu_prestations_client_recap, structure_prestations_client_recap, legende_prestations_client_recap)

        else:
            contenu += r'''
                \tiny{Table II.1 - Récapitulatif des prestations livrées : table vide (pas de prestations livrées)}
                \newline
                '''

        # ## 1.8

        if code_client in acces.sommes:
            structure_bonus = r'''{|l|c|c|c|}'''
            legende_bonus = r'''Table I.8 - Récapitulatif des bonus'''

            contenu_bonus = r'''
                \cline{2-4}
                \multicolumn{1}{c}{} & \multicolumn{3}{|c|}{Bonus} \\
                \hline
                Compte & Déduc. Sp. & Déduc. HC & Total \\
                \hline
                '''
            contenu_bonus += contenu_bonus_compte

            bst = client['bs'] * scl['dst']
            bht = client['bh'] * scl['dht']
            dico_bonus = {'bst': "%.2f" % bst, 'bht': "%.2f" % bht, 'mbt': "%.2f" % scl['somme_t_mb']}
            contenu_bonus += r'''Total & %(bst)s & %(bht)s & %(mbt)s \\
                \hline
                ''' % dico_bonus

            contenu += Latex.tableau(contenu_bonus, structure_bonus, legende_bonus)
        else:
            contenu += r'''
                \tiny{Table I.8 - Récapitulatif des bonus : table vide (pas d'utilisations machines)}
                \newline
                '''

        # ## Annexe 2

        contenu += Annexes.titre_annexe(code_client, client, edition, reference, titre_2, nombre_2)

        contenu += contenu_compte_annexe2

        # ## Annexe 3

        contenu += Annexes.titre_annexe(code_client, client, edition, reference,
                                        "Annexe détaillée des pénalités de réservation", "III")
        contenu += Annexes.section(code_client, client, edition, reference,
                                   "Annexe détaillée des pénalités de réservation", "III")

        # ## 3.1

        if scl['res'] and len(scl['res'].keys()) > 0:
            structure_stats_client = r'''{|l|c|c|c|c|c|c|}'''
            legende_stats_client = r'''Table III.1 - Statistiques des réservations et des utilisations machines'''
            contenu_stats_client = r'''
                \cline{3-7}
                \multicolumn{2}{c}{} & \multicolumn{3}{|c|}{Réservation} & Utilisation & Pénalités \\
                \hline
                 & & Durée & Taux & Util. Min. & Durée & Durée \\
                \hline'''

            machines_utilisees = {}
            ac_somme = None
            re_somme = reservations.sommes[code_client]['machines']

            if code_client in acces.sommes:
                ac_somme = acces.sommes[code_client]['machines']

            for key in scl['res']:
                id_cout = machines.donnees[key]['id_cout']
                nom = machines.donnees[key]['nom']
                if id_cout not in machines_utilisees:
                    machines_utilisees[id_cout] = {}
                machines_utilisees[id_cout][nom] = key

            for id_cout, mics in sorted(machines_utilisees.items()):
                for nom_machine, id_machine in sorted(mics.items()):
                    re_hp = re_somme[id_machine]['res_hp']
                    re_hc = re_somme[id_machine]['res_hc']
                    pu_hp = re_somme[id_machine]['pu_hp']
                    pu_hc = re_somme[id_machine]['pu_hc']
                    tx_hp = machines.donnees[id_machine]['tx_occ_eff_hp']
                    tx_hc = machines.donnees[id_machine]['tx_occ_eff_hc']
                    ok_hp = False
                    ok_hc = False
                    if re_hp > 0 and pu_hp > 0 and tx_hp > 0:
                        ok_hp = True
                    if re_hc > 0 and pu_hc > 0 and tx_hc > 0:
                        ok_hc = True

                    ac_hp = 0
                    ac_hc = 0
                    if ac_somme and id_machine in ac_somme:
                        ac_hp = ac_somme[id_machine]['duree_hp']
                        ac_hc = ac_somme[id_machine]['duree_hc']

                    tot_hp = scl['res'][id_machine]['tot_hp']
                    tot_hc = scl['res'][id_machine]['tot_hc']

                    dico_machine = {'machine': Latex.echappe_caracteres(nom_machine),
                                    'ac_hp': Outils.format_heure(ac_hp), 'ac_hc': Outils.format_heure(ac_hc),
                                    're_hp': Outils.format_heure(re_hp), 're_hc': Outils.format_heure(re_hc),
                                    'tot_hp': Outils.format_heure(tot_hp), 'tot_hc': Outils.format_heure(tot_hc)}

                    users = {}
                    sclu = scl['res'][id_machine]['users']

                    for key in sclu:
                        prenom = sclu[key]['prenom']
                        nom = sclu[key]['nom']
                        if nom not in users:
                            users[nom] = {}
                        if prenom not in users[nom]:
                            users[nom][prenom] = []
                        users[nom][prenom].append(key)

                    if ok_hp:
                        contenu_stats_client += r'''
                            %(machine)s & HP & \hspace{5mm} %(re_hp)s & & & \hspace{5mm} %(ac_hp)s & \hspace{5mm} %(tot_hp)s \\
                             \hline
                             ''' % dico_machine

                        for nom, upi in sorted(users.items()):
                            for prenom, ids in sorted(upi.items()):
                                for id_user in sorted(ids):
                                    ac = sclu[id_user]['ac_hp']
                                    re = sclu[id_user]['re_hp']
                                    mini = sclu[id_user]['mini_hp']
                                    tot = sclu[id_user]['tot_hp']
                                    if ac > 0 or re > 0:
                                        dico_user = {'user': nom + " " + prenom, 'ac': Outils.format_heure(ac),
                                                     're': Outils.format_heure(re), 'tx': tx_hp,
                                                     'mini': Outils.format_heure(mini), 'tot': Outils.format_heure(tot)}
                                        contenu_stats_client += r'''
                                            \hspace{5mm} %(user)s & HP & %(re)s & %(tx)s & %(mini)s & %(ac)s & %(tot)s \\
                                            \hline
                                            ''' % dico_user

                    if ok_hc:
                        contenu_stats_client += r'''
                            %(machine)s & HC & \hspace{5mm} %(re_hc)s & & & \hspace{5mm} %(ac_hc)s & \hspace{5mm} %(tot_hc)s  \\
                             \hline
                             ''' % dico_machine

                        for nom, upi in sorted(users.items()):
                            for prenom, ids in sorted(upi.items()):
                                for id_user in sorted(ids):
                                    ac = sclu[id_user]['ac_hc']
                                    re = sclu[id_user]['re_hc']
                                    mini = sclu[id_user]['mini_hc']
                                    tot = sclu[id_user]['tot_hc']
                                    if ac > 0 or re > 0:
                                        dico_user = {'user': nom + " " + prenom, 'ac': Outils.format_heure(ac),
                                                     're': Outils.format_heure(re), 'tx': tx_hc,
                                                     'mini': Outils.format_heure(mini), 'tot': Outils.format_heure(tot)}
                                        contenu_stats_client += r'''
                                            \hspace{5mm} %(user)s & HC & %(re)s & %(tx)s & %(mini)s & %(ac)s & %(tot)s \\
                                            \hline
                                            ''' % dico_user

            contenu += Latex.tableau(contenu_stats_client, structure_stats_client, legende_stats_client)
        else:
            contenu += r'''
                \tiny{Table III.1 - Statistiques des réservations et des utilisations machines : table vide (pas de pénalités de réservation)}
                \newline
                '''

        # ## 3.2

        if code_client in acces.sommes:
            structure_machuts_client = r'''{|l|c|c|}'''
            legende_machuts_client = r'''Table III.2 - Récapitulatif des utilisations machines par utilisateur'''
            contenu_machuts_client = r'''
                \cline{2-3}
                \multicolumn{1}{c|}{} & HP & HC \\
                \hline'''

            somme = acces.sommes[code_client]['machines']

            machines_utilisees = {}
            for key in somme:
                id_cout = machines.donnees[key]['id_cout']
                nom = machines.donnees[key]['nom']
                if id_cout not in machines_utilisees:
                    machines_utilisees[id_cout] = {}
                machines_utilisees[id_cout][nom] = key

            for id_cout, mics in sorted(machines_utilisees.items()):
                for nom, id_machine in sorted(mics.items()):
                    dico_machine = {'machine': Latex.echappe_caracteres(nom),
                                    'hp': Outils.format_heure(somme[id_machine]['duree_hp']),
                                    'hc': Outils.format_heure(somme[id_machine]['duree_hc'])}
                    contenu_machuts_client += r'''
                       \textbf{%(machine)s} & \hspace{5mm} %(hp)s & \hspace{5mm} %(hc)s \\
                        \hline
                        ''' % dico_machine

                    users = {}
                    for key in somme[id_machine]['users']:
                        prenom = somme[id_machine]['users'][key]['prenom']
                        nom = somme[id_machine]['users'][key]['nom']
                        if nom not in users:
                            users[nom] = {}
                        if prenom not in users[nom]:
                            users[nom][prenom] = []
                        users[nom][prenom].append(key)

                    for nom, upi in sorted(users.items()):
                        for prenom, ids in sorted(upi.items()):
                            for id_user in sorted(ids):
                                smu = somme[id_machine]['users'][id_user]
                                dico_user = {'user': nom + " " + prenom,
                                             'hp': Outils.format_heure(smu['duree_hp']),
                                             'hc': Outils.format_heure(smu['duree_hc'])}
                                contenu_machuts_client += r'''
                                    \hspace{5mm} %(user)s & %(hp)s & %(hc)s \\
                                    \hline
                                    ''' % dico_user

                                for id_compte in smu['comptes']:
                                    smuc = smu['comptes'][id_compte]
                                    dico_compte = {'compte': id_compte, 'hp': Outils.format_heure(smuc['duree_hp']),
                                                 'hc': Outils.format_heure(smuc['duree_hc'])}
                                    contenu_machuts_client += r'''
                                        \hspace{10mm} %(compte)s & %(hp)s \hspace{5mm} & %(hc)s \hspace{5mm} \\
                                        \hline
                                        ''' % dico_compte

            contenu += Latex.long_tableau(contenu_machuts_client, structure_machuts_client, legende_machuts_client)
        else:
            contenu += r'''
                \tiny{Table III.2 - Récapitulatif des utilisations machines par utilisateur : table vide (pas d’utilisation machines)}
                \newline
                '''

        # ## 3.3

        if code_client in reservations.sommes:
            structure_reserve_client = r'''{|c|c|c|c|c|}'''
            legende_reserve_client = r'''Table III.3 - Détail des réservations machines par utilisateur'''

            contenu_reserve_client = r'''
                \cline{4-5}
                \multicolumn{3}{c}{} & \multicolumn{2}{|c|}{Durée réservée} \\
                \cline{4-5}
                \multicolumn{3}{c|}{} & HP & HC \\
                \hline
                '''

            somme = reservations.sommes[code_client]['machines']

            machines_reservees = {}
            for key in somme:
                id_cout = machines.donnees[key]['id_cout']
                nom = machines.donnees[key]['nom']
                if id_cout not in machines_reservees:
                    machines_reservees[id_cout] = {}
                machines_reservees[id_cout][nom] = key

            for id_cout, mics in sorted(machines_reservees.items()):
                for nom, id_machine in sorted(mics.items()):

                    dico_machine = {'machine': Latex.echappe_caracteres(nom),
                                    'hp': Outils.format_heure(somme[id_machine]['res_hp']),
                                    'hc': Outils.format_heure(somme[id_machine]['res_hc'])}
                    contenu_reserve_client += r'''
                                \multicolumn{3}{|l|}{\textbf{%(machine)s}} & \hspace{5mm} %(hp)s &
                                \hspace{5mm} %(hc)s \\
                                \hline
                                ''' % dico_machine

                    users = {}
                    for key in somme[id_machine]['users']:
                        prenom = somme[id_machine]['users'][key]['prenom']
                        nom = somme[id_machine]['users'][key]['nom']
                        if nom not in users:
                            users[nom] = {}
                        if prenom not in users[nom]:
                            users[nom][prenom] = []
                        users[nom][prenom].append(key)

                    for nom, upi in sorted(users.items()):
                        for prenom, ids in sorted(upi.items()):
                            for id_user in sorted(ids):
                                smu = somme[id_machine]['users'][id_user]
                                dico_user = {'user': nom + " " + prenom,
                                             'hp': Outils.format_heure(smu['res_hp']),
                                             'hc': Outils.format_heure(smu['res_hc'])}
                                contenu_reserve_client += r'''
                                            \multicolumn{3}{|l|}{\hspace{5mm} %(user)s} & %(hp)s & %(hc)s \\
                                            \hline
                                        ''' % dico_user
                                for pos in smu['data']:
                                    res = reservations.donnees[pos]
                                    login = Latex.echappe_caracteres(res['date_debut']).split()
                                    temps = login[0].split('-')
                                    date = temps[0]
                                    for pos in range(1, len(temps)):
                                        date = temps[pos] + '.' + date
                                    if len(login) > 1:
                                        heure = login[1]
                                    else:
                                        heure = ""

                                    sup = ""
                                    if res['si_supprime'] == "OUI":
                                        sup = "Supprimé le : " + Latex.echappe_caracteres(res['date_suppression'])
                                    dico_pos = {'date': date, 'heure': heure, 'sup': sup,
                                                'hp': Outils.format_heure(res['duree_fact_hp']),
                                                'hc': Outils.format_heure(res['duree_fact_hc'])}
                                    contenu_reserve_client += r'''
                                                \hspace{10mm} %(date)s & %(heure)s & %(sup)s & %(hp)s \hspace{5mm} &
                                                 %(hc)s \hspace{5mm} \\
                                                \hline
                                            ''' % dico_pos

            contenu += Latex.long_tableau(contenu_reserve_client, structure_reserve_client, legende_reserve_client)
        else:
            contenu += r'''
                \tiny{Table III.3 - Détail des réservations machines par utilisateur : table vide (pas de réservation machines)}
                \newline
                '''

        # ## Annexe 4

        contenu += Annexes.titre_annexe(code_client, client, edition, reference, titre_4, nombre_4)

        contenu += contenu_compte_annexe4

        # ## Annexe 5

        contenu += Annexes.titre_annexe(code_client, client, edition, reference,
                                        "Justificatif des coûts d'utilisation par compte", "V")
        contenu += Annexes.section(code_client, client, edition, reference,
                                   "Justificatif des coûts d'utilisation par compte", "V")

        return contenu

    @staticmethod
    def titre_annexe(code_client, client, edition, reference, titre, nombre):
        dic_titre = {'code': code_client, 'code_sap': client['code_sap'],
                     'nom': Latex.echappe_caracteres(client['abrev_labo']),
                     'date': edition.mois_txt + " " + str(edition.annee),
                     'ref': reference, 'titre': titre, 'nombre': nombre}

        contenu = r'''
            \clearpage
            \begin{titlepage}
            %(ref)s \hspace*{4cm} Client %(code)s - %(code_sap)s - %(nom)s - %(date)s
            \vspace*{8cm}
            \begin{adjustwidth}{5cm}{}
            \Large\textsc{Annexe %(nombre)s} \newline\newline
            \Large\textsc{%(titre)s} \newpage
            \end{adjustwidth}
            \end{titlepage}
            ''' % dic_titre

        return contenu

    @staticmethod
    def section(code_client, client, edition, reference, titre, nombre):
        dic_section = {'code': code_client, 'code_sap': client['code_sap'],
                       'nom': Latex.echappe_caracteres(client['abrev_labo']),
                       'date': edition.mois_txt + " " + str(edition.annee),
                       'ref': reference, 'titre': titre, 'nombre': nombre}

        section = r'''
            \fakesection{%(ref)s \hspace*{4cm} Client %(code)s - %(code_sap)s - %(nom)s - %(date)s}{Annexe %(nombre)s - %(titre)s}
            ''' % dic_section

        return section