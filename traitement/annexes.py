#import os
#import shutil

from outils import Outils
from latex import Latex


class Annexes(object):
    """
    Classe pour la création des annexes
    """

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
            contenu += r'''\usepackage[margin=10mm, includefoot]{geometry}
                \usepackage{multirow}
                \usepackage{graphicx}
                \usepackage{longtable}
                \usepackage{dcolumn}
                \usepackage{changepage}
                \usepackage[scriptsize]{caption}
                \usepackage{fancyhdr}
                \pagestyle{fancy}

                \fancyfoot[R]{\thepage}
                \fancyfoot[L]{\leftmark}

                \newcommand{\fakesection}[1]{%
                \par\refstepcounter{section}
                \sectionmark{#1}
                }

                \begin{document}
                \renewcommand{\arraystretch}{1.5}
                '''
            contenu += r'''
                \begin{titlepage}
                \vspace*{8cm}
                \begin{adjustwidth}{5cm}{}
                \Large\textsc{''' + garde + r'''}\newline\newline'''

            client = clients.donnees[code_client]

            nature = generaux.nature_client_par_code_n(client['type_labo'])
            reference = nature + str(edition.annee)[2:] + Outils.mois_string(edition.mois) + "." + code_client
            if edition.version != "0":
                reference += "-" + edition.version
            dic_entete = {'code': code_client, 'code_sap': client['code_sap'],
                          'nom': Latex.echappe_caracteres(client['abrev_labo']),
                          'date': edition.mois_txt + " " + str(edition.annee),
                          'ref': reference}

            contenu += r''' %(code)s -  %(code_sap)s -  %(nom)s \newline
                 %(date)s \newline
                  %(ref)s
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
        reference = nature + str(edition.annee)[2:] + Outils.mois_string(edition.mois) + "." + code_client
        if edition.version != "0":
            reference += "-" + edition.version
        intitule_client = code_client + " - " + Latex.echappe_caracteres(client['abrev_labo'])

        # structure_recap_compte = r'''{|l|r|r|'''
        # contenu_recap_compte = r'''
        #     \hline
        #     Compte & \multicolumn{1}{l|}{Procédés}'''
        #
        # for article in generaux.articles_d3:
        #     structure_recap_compte += r'''r|'''
        #     contenu_recap_compte += r''' & \multicolumn{1}{l|}{
        #     ''' + Latex.echappe_caracteres(article.intitule_long) + r'''}'''
        # structure_recap_compte += r'''}'''
        # contenu_recap_compte += r'''& \multicolumn{1}{l|}{Total} \\
        #     \hline
        #     '''

        # contenu_procedes_client = r'''
        #     \cline{2-6}
        #     \multicolumn{1}{r}{} & \multicolumn{5}{|c|}{Procédés} \\
        #     \hline
        #     Compte & Machine & Déduc. Mach. & Net Machine & M.O. opérateur & Montant net \\
        #     \hline
        #     '''

        # contenu_prestations_client_tab = {}
        # for article in generaux.articles_d3:
        #     contenu_prestations_client_tab[article.code_d] = r'''
        #         \cline{2-4}
        #         \multicolumn{1}{c}{} & \multicolumn{3}{|c|}{''' + article. intitule_long + r'''} \\
        #         \hline
        #         Compte & Montant & Rabais & Montant net \\
        #         \hline
        #         '''

        client_comptes = sommes.sommes_comptes[code_client]
        contenu_compte_annexe2 = ""
        contenu_compte_annexe4 = ""
        contenu_compte_annexe5 = ""

        for id_compte in sorted(client_comptes.keys()):

            # ## COMPTE

            sco = sommes.sommes_comptes[code_client][id_compte]
            compte = comptes.donnees[id_compte]
            intitule_compte = id_compte + " - " + Latex.echappe_caracteres(compte['intitule'])

            # ## 1.1 ? ligne récapitulatif comptes pour client

            # dico_recap_compte = {'compte': intitule_compte, 'procede': "%.2f" % sco['mj'],
            #                      'total': "%.2f" % sco['mj']}
            #
            # contenu_recap_compte += r'''%(compte)s & %(procede)s ''' \
            #                         % dico_recap_compte
            #
            # for categorie in generaux.codes_d3():
            #     contenu_recap_compte += r''' & ''' + "%.2f" % sco['tot_cat'][categorie]
            #
            # contenu_recap_compte += r'''& %(total)s \\
            #     \hline
            #     ''' % dico_recap_compte

            # ## 1.6 ? ligne coûts procédés pour client

            # dico_procedes_client = {'intitule': intitule_compte, 'maij': "%.2f" % sco['somme_j_mai'],
            #                         'dtij': "%.2f" % sco['somme_j_mr'],
            #                         'nmij': "%.2f" % (sco['somme_j_mai'] - sco['somme_j_mr']),
            #                         'moij': "%.2f" % sco['somme_j_moi'], 'mj': "%.2f" % sco['mj']}
            # contenu_procedes_client += r'''
            #     %(intitule)s & %(maij)s & %(dtij)s & %(nmij)s & %(moij)s & %(mj)s \\
            #     \hline
            #     ''' % dico_procedes_client

            # ## 1.7 ? ligne prestations livrées pour client

            # for article in generaux.articles_d3:
            #     dico_prestations_client = {'intitule': intitule_compte,
            #                                'cmj': "%.2f" % sco['sommes_cat_m'][article.code_d],
            #                                'crj': "%.2f" % sco['sommes_cat_r'][article.code_d],
            #                                'cj': "%.2f" % sco['tot_cat'][article.code_d]}
            #     contenu_prestations_client_tab[article.code_d] += r'''
            #     %(intitule)s & %(cmj)s & %(crj)s & %(cj)s \\
            #     \hline
            #     ''' % dico_prestations_client

            # ## 2.1 ? récapitulatif postes pour compte

            # structure_recap_poste = r'''{|l|c|c|c|c|c|}'''
            # legende_recap_poste = r'''Récapitulatif postes pour compte ''' + intitule_compte
            #
            # dico_recap_poste = {'mm': "%.2f" % sco['somme_j_mm'], 'mr': "%.2f" % sco['somme_j_mr'],
            #                     'maij': "%.2f" % sco['somme_j_mai'], 'dsij': "%.2f" % sco['somme_j_dsi'],
            #                     'dhij': "%.2f" % sco['somme_j_dhi'], 'mj': "%.2f" % sco['mj'],
            #                     'nmij': "%.2f" % (sco['somme_j_mai']-sco['somme_j_mr']),
            #                     'moij': "%.2f" % sco['somme_j_moi']}
            #
            # contenu_recap_poste = r'''
            #     \cline{2-6}
            #     \multicolumn{1}{r|}{} & Montant & Déduc. Sp. & Déduc. HC & Rabais T. & Net \\
            #     \hline
            #     Procédés & %(mm)s &  &  & %(mr)s & %(mj)s \\
            #     \hline
            #     \hspace{5mm} \textit{Machine} & \textit{%(maij)s} & \textit{%(dsij)s} & \textit{%(dhij)s}  &
            #         \textit{%(mr)s}  & \textit{%(nmij)s} \\
            #     \hline
            #     \hspace{5mm} \textit{Main d'oeuvre} & \textit{%(moij)s} &  &  &  & \textit{%(moij)s}  \\
            #     \hline
            #     ''' % dico_recap_poste
            #
            # for article in generaux.articles_d3:
            #     contenu_recap_poste += Latex.echappe_caracteres(article.intitule_long)
            #     contenu_recap_poste += r''' & ''' + "%.2f" % sco['sommes_cat_m'][article.code_d]
            #     contenu_recap_poste += r''' & & &  ''' + "%.2f" % sco['sommes_cat_r'][article.code_d]
            #     contenu_recap_poste += r''' & ''' + "%.2f" % sco['tot_cat'][article.code_d]
            #     contenu_recap_poste += r''' \\
            #         \hline
            #         '''
            #
            # total = 0
            # contenu_recap_poste += r'''\multicolumn{5}{|r|}{Total} & ''' + "%.2f" % total + r'''\\
            #     \hline
            #     '''
            #
            # contenu_compte_annexe2 += Latex.tableau(contenu_recap_poste, structure_recap_poste, legende_recap_poste)

            # ## 2.2 ?  durée utilisée compte

            # structure_utilise_compte = r'''{|l|c|c|c|c|c|c|c|c|c|c|c|}'''
            # legende_utilise_compte = r'''Procédés (machine + main d'oeuvre) pour compte ''' + intitule_compte
            #
            # contenu_utilise_compte = r'''
            #     \cline{3-11}
            #     \multicolumn{2}{c}{} & \multicolumn{2}{|c|}{Machine} & \multicolumn{2}{l|}{PU [CHF/h]} & \multicolumn{5}{l|}{Montant [CHF]} \\
            #     \hline
            #     ''' + intitule_compte + r''' & & Mach. & Oper. & Mach. & Oper. & Machine & Déduc. Sp. & Déduc. HC & Net & M.O. \\
            #     \hline
            #     '''
            #
            # dico_utilise_compte = {'maij': "%.2f" % sco['somme_j_mai'], 'dsij': "%.2f" % sco['somme_j_dsi'],
            #                        'dhij': "%.2f" % sco['somme_j_dhi'], 'moij': "%.2f" % sco['somme_j_moi'],
            #                        'nmij': "%.2f" % (sco['somme_j_mai']-sco['somme_j_mr'])}
            # contenu_utilise_compte += r'''
            #     \multicolumn{6}{|r|}{Total} & %(maij)s & %(dsij)s & %(dhij)s & %(nmij)s & %(moij)s \\
            #     \hline
            #     ''' % dico_utilise_compte
            #
            # contenu_compte_annexe2 += Latex.tableau(contenu_utilise_compte, structure_utilise_compte, legende_utilise_compte)

            # ## 4.1

            if code_client in acces.sommes and id_compte in acces.sommes[code_client]['comptes']:

                structure_machuts_compte = r'''{|l|l|l|c|c|c|c|}'''
                legende_machuts_compte = r'''Table IV.1 - Détails des utilisations machines pour compte'''

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
                            users[nom][prenom] = key

                        for nom, upi in sorted(users.items()):
                            for prenom, id_user in sorted(upi.items()):
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
                                %(num)s - %(nom)s & \hspace{5mm} %(quantite)s & %(unite)s & %(rabais)s  \\
                                \hline
                                ''' % dico_prestations

                            users = {}
                            for key in sip['users']:
                                prenom = sip['users'][key]['prenom']
                                nom = sip['users'][key]['nom']
                                if nom not in users:
                                    users[nom] = {}
                                users[nom][prenom] = key

                            for nom, upi in sorted(users.items()):
                                for prenom, id_user in sorted(upi.items()):
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
                                            \hspace{10mm} %(date_liv)s N. livraison: %(id)s & %(quantite)s \hspace{5mm} & %(unite)s & %(rabais)s \\

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

        titre = Annexes.titre_annexe(code_client, client, edition, reference,
                                     "Annexe I - Récapitulatif")
        contenu += titre[0]
 #       contenu += titre[1]

        # ## 1.1 ? récapitulatif postes pour client

       #  structure_recap_poste_cl = r'''{|l|r|r|r|}'''
       #  legende_recap_poste_cl = r'''Récapitulatif postes pour client ''' + intitule_client
       #
       #  dico_recap_poste_cl = {'emom': "%.2f" % scl['em'], 'emor': "%.2f" % scl['er'], 'emo': "%.2f" % scl['e'],
       #                         'resm': "%.2f" % scl['rm'], 'resr': "%.2f" % scl['rr'], 'res': "%.2f" % scl['r'],
       #                         'mat': "%.2f" % scl['mat'], 'stmm': "%.2f" % scl['somme_t_mm'],
       #                         'stmr': "%.2f" % scl['somme_t_mr'], 'mt': "%.2f" % scl['mt'],
       #                         'mot': "%.2f" % scl['mot'], 'matr': "%.2f" % (scl['mat']-scl['somme_t_mr'])}
       #
       #  contenu_recap_poste_cl = r'''
       #      \hline
       #       & \multicolumn{1}{l|}{Montant} & \multicolumn{1}{l|}{Rabais} & \multicolumn{1}{l|}{Total} \\
       #      \hline
       #      Emolument & %(emom)s & %(emor)s & %(emo)s \\
       #      \hline
       #      Frais de réservation & %(resm)s & %(resr)s & %(res)s \\
       #      \hline
       #      Procédés & %(stmm)s & %(stmr)s & %(mt)s \\
       #      \hline
       #      \hspace{5mm} \textit{Machine} & \textit{%(mat)s} & \textit{%(stmr)s} & \textit{%(matr)s} \\
       #      \hline
       #      \hspace{5mm} \textit{Main d'oeuvre} & \textit{%(mot)s} & \textit{0.00} & \textit{%(mot)s}  \\
       #      \hline
       #      ''' % dico_recap_poste_cl
       #
       #  for article in generaux.articles_d3:
       #      contenu_recap_poste_cl += Latex.echappe_caracteres(article.intitule_long)
       #      contenu_recap_poste_cl += r''' & '''
       #      contenu_recap_poste_cl += "%.2f" % scl['sommes_cat_m'][article.code_d] + r''' & '''
       #      contenu_recap_poste_cl += "%.2f" % scl['sommes_cat_r'][article.code_d] + r''' & '''
       #      contenu_recap_poste_cl += "%.2f" % scl['tot_cat'][article.code_d] + r''' \\
       #          \hline
       #          '''
       #
       # contenu_recap_poste_cl += r'''\multicolumn{3}{|r|}{Total} & ''' + "%.2f" % (scl['somme_t'] + scl['e']) + r'''\\
       #      \hline
       #      '''
       #
       #  contenu += Latex.tableau(contenu_recap_poste_cl, structure_recap_poste_cl, legende_recap_poste_cl)

        # ## 1.3 ? émolument pour client

        # structure_emolument = r'''{|r|r|l|r|r|r|r|}'''
        # legende_emolument = r'''Emolument pour client ''' + intitule_client
        #
        # dico_emolument = {'emb':  "%.2f" % client['emol_base_mens'], 'ef':  "%.2f" % client['emol_fixe'],
        #                   'pente': client['coef'], 'tot_eq_r': "%.2f" % scl['r'],
        #                   'tot_eq_m': "%.2f" % (scl['mt']-scl['mot']), 'tot_eq': "%.2f" % scl['somme_eq'],
        #                   'rabais': "%.2f" % scl['er']}
        #
        # contenu_emolument = r'''
        #     \hline
        #     \multicolumn{1}{|l|}{Emolument de base} & \multicolumn{1}{l|}{Emolument fixe} & Pente
        #     & \multicolumn{1}{l|}{Total EQ R} & \multicolumn{1}{l|}{Total EQ M} & \multicolumn{1}{l|}{Total EQ} &
        #     \multicolumn{1}{l|}{Rabais émolument} \\
        #     \hline
        #     %(emb)s & %(ef)s & %(pente)s & %(tot_eq_r)s & %(tot_eq_m)s & %(tot_eq)s & %(rabais)s \\
        #     \hline
        #     ''' % dico_emolument
        #
        # contenu += Latex.tableau(contenu_emolument, structure_emolument, legende_emolument)

        # ## 1.4 ? pénalités réservation

        if scl['r'] > 0:
            structure_frais_client = r'''{|l|c|c|c|c|c|c|r|r|}'''
            legende_frais_client = r'''Pénalités réservation pour client : ''' + intitule_client

            contenu_frais_client = r'''
                \cline{3-9}
                \multicolumn{2}{c}{} & \multicolumn{3}{|c|}{Réservation} & Utilis. & Pénal. & PU & Montant \\
                \cline{3-9}
                \multicolumn{2}{c|}{} & Durée & Taux & Util Min & Durée & Durée & CHF/h & CHF \\
                \hline
                '''
            # machines_utilisees = {}
            # for key in scl['res']:
            #     machines_utilisees[key] = {'machine': machines.donnees[key]['nom']}
            #
            # for machine_t in sorted(machines_utilisees.items(), key=lambda k_v: k_v[1]['machine']):
            #     machine = machine_t[1]
            #     id_mach = machine_t[0]
            #     som_m = scl['res'][id_mach]
            #     if som_m['pen_hp'] > 0 or som_m['pen_hc'] > 0:
            #
            #         dico_frais_client = {
            #             'machine': Latex.echappe_caracteres(machine['machine']),
            #             'pen_hp': "%.1f" % som_m['pen_hp'], 'pen_hc': "%.1f" % som_m['pen_hc'],
            #             'mont_hp': Outils.format_si_nul(som_m['m_hp']), 'mont_hc': Outils.format_si_nul(som_m['m_hc']),
            #             'pu_hp': Outils.format_si_nul(reservations.sommes[code_client]['machines'][id_mach]['pu_hp']),
            #             'pu_hc': Outils.format_si_nul(reservations.sommes[code_client]['machines'][id_mach]['pu_hc'])}
            #
            #         if som_m['pen_hp'] > 0:
            #             contenu_frais_client += r'''%(machine)s & HP &  %(pen_hp)s & %(pu_hp)s & %(mont_hp)s \\
            #                 \hline
            #                 ''' % dico_frais_client
            #
            #         if som_m['pen_hc'] > 0:
            #             contenu_frais_client += r'''%(machine)s & HC &  %(pen_hc)s & %(pu_hc)s & %(mont_hc)s \\
            #                 \hline
            #                 ''' % dico_frais_client
            #
            # contenu_frais_client += r'''
            #     \multicolumn{4}{|r|}{Total} & ''' + Outils.format_si_nul(scl['r']) + r'''\\
            #     \hline
            #     '''

            contenu_frais_client += r'''
                \multicolumn{8}{|r|}{Total} & ''' + "%.2f" % scl['r'] + r''' \\
                \hline
                '''

#            contenu += Latex.tableau(contenu_frais_client, structure_frais_client, legende_frais_client)

        # ## 1.6 ? coûts procédés

        # structure_procedes_client = r'''{|l|c|c|c|c|c|}'''
        # legende_procedes_client = r'''Récapitulatif des coûts procédés pour client : ''' + intitule_client
        #
        # dico_procedes_client = {'mat': "%.2f" % scl['mat'], 'dti': "%.2f" % scl['somme_t_mr'],
        #                         'nmi': "%.2f" % (scl['mat']-scl['somme_t_mr']),
        #                         'mot': "%.2f" % scl['mot'], 'mt': "%.2f" % scl['mt']}
        # contenu_procedes_client += r'''
        #     Total article & %(mat)s & %(dti)s & %(nmi)s & %(mot)s & %(mt)s \\
        #     \hline
        #     ''' % dico_procedes_client
        #
        # contenu += Latex.tableau(contenu_procedes_client, structure_procedes_client, legende_procedes_client)

        # ## 1.7 prestations livrées

        # structure_prestations_client = r'''{|l|c|c|c|}'''
        # legende_prestations_client = r'''Récapitulatif des prestations livrées pour client : ''' + intitule_client
        #
        # contenu_prestations_client = ""
        # i = 0
        # for article in generaux.articles_d3:
        #     dico_prestations_client = {'cmt': "%.2f" % scl['sommes_cat_m'][article.code_d],
        #                                'crt': "%.2f" % scl['sommes_cat_r'][article.code_d],
        #                                'ct': "%.2f" % scl['tot_cat'][article.code_d]}
        #     contenu_prestations_client_tab[article.code_d] += r'''
        #     Total article & %(cmt)s & %(crt)s & %(ct)s \\
        #     \hline
        #     ''' % dico_prestations_client
        #     if i == 0:
        #         i += 1
        #     else:
        #         contenu_prestations_client += r'''\multicolumn{4}{c}{} \\'''
        #     contenu_prestations_client += contenu_prestations_client_tab[article.code_d]
        #
        # contenu += Latex.tableau(contenu_prestations_client, structure_prestations_client, legende_prestations_client)

        # ## 1.2 ? récapitulatif comptes pour client

        # legende_recap_compte = r'''Récapitulatif des comptes pour client ''' + intitule_client
        #
        # dico_recap_compte = {'procedes': "%.2f" % scl['mt'], 'total': "%.2f" % (scl['somme_t']-scl['r'])}
        #
        # contenu_recap_compte += r'''Total article & %(procedes)s''' % dico_recap_compte
        #
        # for categorie in generaux.codes_d3():
        #     contenu_recap_compte += r''' & ''' + "%.2f" % scl['tot_cat'][categorie]
        #
        # contenu_recap_compte += r'''& %(total)s \\
        #     \hline
        #     ''' % dico_recap_compte
        #
        # contenu += Latex.tableau(contenu_recap_compte, structure_recap_compte, legende_recap_compte)

        # ## Annexe 2

        titre = Annexes.titre_annexe(code_client, client, edition, reference,
                                     "Annexe II - Récapitulatifs par compte")
        contenu += titre[0]
#        contenu += titre[1]

        # ## Annexe 3

        titre = Annexes.titre_annexe(code_client, client, edition, reference,
                                     "Annexe III - Annexe détaillée des pénalités de réservation")
        contenu += titre[0]
 #       contenu += titre[1]

        # ## 3.1

        if code_client in acces.sommes or code_client in reservations.sommes:
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
            re_somme = None

            if code_client in acces.sommes:
                ac_somme = acces.sommes[code_client]['machines']
                for key in ac_somme:
                    id_cout = machines.donnees[key]['id_cout']
                    nom = machines.donnees[key]['nom']
                    if id_cout not in machines_utilisees:
                        machines_utilisees[id_cout] = {}
                        machines_utilisees[id_cout][nom] = key

            if code_client in reservations.sommes:
                re_somme = reservations.sommes[code_client]['machines']
                for key in re_somme:
                    id_cout = machines.donnees[key]['id_cout']
                    nom = machines.donnees[key]['nom']
                    if id_cout not in machines_utilisees:
                        machines_utilisees[id_cout] = {}
                        machines_utilisees[id_cout][nom] = key

            for id_cout, mics in sorted(machines_utilisees.items()):
                for nom, id_machine in sorted(mics.items()):
                    hc = machines.donnees[id_machine]['hc']
                    dico_machine = {'machine': Latex.echappe_caracteres(nom)}
                    contenu_stats_client += r'''
                        %(machine)s & HP/HC & & & & & \\
                         \hline
                         ''' % dico_machine

                    users = {}
                    if ac_somme and id_machine in ac_somme:
                        for key in ac_somme[id_machine]['users']:
                            prenom = ac_somme[id_machine]['users'][key]['prenom']
                            nom = ac_somme[id_machine]['users'][key]['nom']
                            if nom not in users:
                                users[nom] = {}
                            users[nom][prenom] = key
                    if re_somme and id_machine in re_somme:
                        for key in re_somme[id_machine]['users']:
                            prenom = re_somme[id_machine]['users'][key]['prenom']
                            nom = re_somme[id_machine]['users'][key]['nom']
                            if nom not in users:
                                users[nom] = {}
                            users[nom][prenom] = key

                    for nom, upi in sorted(users.items()):
                        for prenom, id_user in sorted(upi.items()):
                            dico_user = {'user': nom + " " + prenom}
                            contenu_stats_client += r'''
                                \hspace{5mm} %(user)s & HP/HC & & & & & \\
                                \hline
                                ''' % dico_user


            contenu += Latex.tableau(contenu_stats_client, structure_stats_client, legende_stats_client)

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
                        users[nom][prenom] = key

                    for nom, upi in sorted(users.items()):
                        for prenom, id_user in sorted(upi.items()):
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

            contenu += Latex.tableau(contenu_machuts_client, structure_machuts_client, legende_machuts_client)

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
                        users[nom][prenom] = key

                    for nom, upi in sorted(users.items()):
                        for prenom, id_user in sorted(upi.items()):
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

        # ## Annexe 4

        titre = Annexes.titre_annexe(code_client, client, edition, reference,
                                     "Annexe IV - Annexe détaillée par compte")
        contenu += titre[0]
#        contenu += titre[1]

        contenu += contenu_compte_annexe4

        # ## Annexe 5

        titre = Annexes.titre_annexe(code_client, client, edition, reference,
                                     "Annexe V - Justificatif des coûts d'utilisation par compte")
        contenu += titre[0]
       # contenu += titre[1]


        # ## 3.2 ? durée utilisée client

        # structure_utilise_client = r'''{|l|c|c|}'''
        # legende_utilise_client = r'''Récapitulatif heures machines par utilisateur pour client : ''' + intitule_client
        #
        # contenu_utilise_client = r'''
        #     \cline{2-3}
        #     \multicolumn{1}{c|}{} & HP & HC \\
        #     \hline
        #     '''
        #
        # contenu += Latex.tableau(contenu_utilise_client, structure_utilise_client, legende_utilise_client)

        return contenu

    @staticmethod
    def titre_annexe(code_client, client, edition, reference, titre):
        dic_entete = {'code': code_client, 'code_sap': client['code_sap'],
                      'nom': Latex.echappe_caracteres(client['abrev_labo']),
                      'date': edition.mois_txt + " " + str(edition.annee),
                      'ref': reference, 'titre': titre}
        entete = r'''
            %(titre)s pour %(code)s - %(code_sap)s - %(nom)s - %(date)s - %(ref)s
            ''' % dic_entete

        contenu = r'''
            \begin{titlepage}
            \vspace*{8cm}
            \begin{adjustwidth}{1cm}{}
            \Large\textsc{%(titre)s \newline %(code)s - %(code_sap)s - %(nom)s} \newline
            %(date)s \newline
            %(ref)s \newpage
            \end{adjustwidth}
            \end{titlepage}
            \fakesection{%(titre)s}
            ''' % dic_entete

        return contenu, entete
