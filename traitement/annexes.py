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
                \usepackage{longtable}
                \usepackage{dcolumn}
                \usepackage{changepage}
                \usepackage[scriptsize]{caption}

                \begin{document}
                \renewcommand{\arraystretch}{1.5}
                '''
            contenu += r'''
                \vspace*{8cm}
                \begin{adjustwidth}{5cm}{}
                \Large\textsc{''' + garde + r'''}\newline\newline'''
            nom = Latex.echappe_caracteres(clients.donnees[code_client]['abrev_labo'])
            code_sap = clients.donnees[code_client]['code_sap']

            contenu += code_client + " - " + code_sap + " - " + nom + r'''\newpage
                \end{adjustwidth}'''
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
        intitule_client = code_client + " - " + Latex.echappe_caracteres(client['abrev_labo'])

        structure_recap_compte = r'''{|l|r|r|'''
        contenu_recap_compte = r'''
            \hline
            Compte & \multicolumn{1}{l|}{Procédés}'''

        for article in generaux.articles_d3:
            structure_recap_compte += r'''r|'''
            contenu_recap_compte += r''' & \multicolumn{1}{l|}{
            ''' + Latex.echappe_caracteres(article.intitule_long) + r'''}'''
        structure_recap_compte += r'''}'''
        contenu_recap_compte += r'''& \multicolumn{1}{l|}{Total} \\
            \hline
            '''

        contenu_procedes_client = r'''
            \cline{2-6}
            \multicolumn{1}{r}{} & \multicolumn{5}{|c|}{Procédés} \\
            \hline
            Compte & Machine & Déduc. Mach. & Net Machine & M.O. opérateur & Montant net \\
            \hline
            '''

        contenu_prestations_client_tab = {}
        for article in generaux.articles_d3:
            contenu_prestations_client_tab[article.code_d] = r'''
                \cline{2-4}
                \multicolumn{1}{c}{} & \multicolumn{3}{|c|}{''' + article. intitule_long + r'''} \\
                \hline
                Compte & Montant & Rabais & Montant net \\
                \hline
                '''

        client_comptes = sommes.sommes_comptes[code_client]
        contenu_compte_annexe3 = ""
        contenu_compte_annexe4 = ""

        for id_compte in sorted(client_comptes.keys()):

            # ## COMPTE

            sco = sommes.sommes_comptes[code_client][id_compte]
            compte = comptes.donnees[id_compte]
            intitule_compte = id_compte + " - " + Latex.echappe_caracteres(compte['intitule'])

            # ## ligne récapitulatif comptes pour client

            dico_recap_compte = {'compte': intitule_compte, 'procede': "%.2f" % sco['mj'],
                                 'total': "%.2f" % sco['mj']}

            contenu_recap_compte += r'''%(compte)s & %(procede)s ''' \
                                    % dico_recap_compte

            for categorie in generaux.codes_d3():
                contenu_recap_compte += r''' & ''' + "%.2f" % sco['tot_cat'][categorie]

            contenu_recap_compte += r'''& %(total)s \\
                \hline
                ''' % dico_recap_compte

            # ## ligne coûts procédés pour client

            dico_procedes_client = {'intitule': intitule_compte, 'maij': "%.2f" % sco['somme_j_mai'],
                                    'dtij': "%.2f" % sco['somme_j_mr'],
                                    'nmij': "%.2f" % (sco['somme_j_mai'] - sco['somme_j_mr']),
                                    'moij': "%.2f" % sco['somme_j_moi'], 'mj': "%.2f" % sco['mj']}
            contenu_procedes_client += r'''
                %(intitule)s & %(maij)s & %(dtij)s & %(nmij)s & %(moij)s & %(mj)s \\
                \hline
                ''' % dico_procedes_client

            # ## ligne prestations livrées pour client

            for article in generaux.articles_d3:
                dico_prestations_client = {'intitule': intitule_compte,
                                           'cmj': "%.2f" % sco['sommes_cat_m'][article.code_d],
                                           'crj': "%.2f" % sco['sommes_cat_r'][article.code_d],
                                           'cj': "%.2f" % sco['tot_cat'][article.code_d]}
                contenu_prestations_client_tab[article.code_d] += r'''
                %(intitule)s & %(cmj)s & %(crj)s & %(cj)s \\
                \hline
                ''' % dico_prestations_client

            # ## Annexe 3

            dico_nom = {'labo': Latex.echappe_caracteres(client['abrev_labo']),
                        'compte': Latex.echappe_caracteres(compte['intitule']),
                        'date': edition.mois_txt + " " + str(edition.annee)}
            contenu_compte_annexe3 += r'''
                \clearpage
                Annexe 3 : %(labo)s - %(compte)s - %(date)s
                ''' % dico_nom

            # ## récapitulatif postes pour compte

            structure_recap_poste = r'''{|l|c|c|c|c|c|}'''
            legende_recap_poste = r'''Récapitulatif postes pour compte ''' + intitule_compte

            dico_recap_poste = {'mm': "%.2f" % sco['somme_j_mm'], 'mr': "%.2f" % sco['somme_j_mr'],
                                'maij': "%.2f" % sco['somme_j_mai'], 'dsij': "%.2f" % sco['somme_j_dsi'],
                                'dhij': "%.2f" % sco['somme_j_dhi'], 'mj': "%.2f" % sco['mj'],
                                'nmij': "%.2f" % (sco['somme_j_mai']-sco['somme_j_mr']),
                                'moij': "%.2f" % sco['somme_j_moi']}

            contenu_recap_poste = r'''
                \cline{2-6}
                \multicolumn{1}{r|}{} & Montant & Déduc. Sp. & Déduc. HC & Rabais T. & Net \\
                \hline
                Procédés & %(mm)s &  &  & %(mr)s & %(mj)s \\
                \hline
                \hspace{5mm} \textit{Machine} & \textit{%(maij)s} & \textit{%(dsij)s} & \textit{%(dhij)s}  &
                    \textit{%(mr)s}  & \textit{%(nmij)s} \\
                \hline
                \hspace{5mm} \textit{Main d'oeuvre} & \textit{%(moij)s} &  &  &  & \textit{%(moij)s}  \\
                \hline
                ''' % dico_recap_poste

            for article in generaux.articles_d3:
                contenu_recap_poste += Latex.echappe_caracteres(article.intitule_long)
                contenu_recap_poste += r''' & ''' + "%.2f" % sco['sommes_cat_m'][article.code_d]
                contenu_recap_poste += r''' & & &  ''' + "%.2f" % sco['sommes_cat_r'][article.code_d]
                contenu_recap_poste += r''' & ''' + "%.2f" % sco['tot_cat'][article.code_d]
                contenu_recap_poste += r''' \\
                    \hline
                    '''

            total = 0
            contenu_recap_poste += r'''\multicolumn{5}{|r|}{Total} & ''' + "%.2f" % total + r'''\\
                \hline
                '''

            contenu_compte_annexe3 += Latex.tableau(contenu_recap_poste, structure_recap_poste, legende_recap_poste)

            # ## T2 : durée utilisée compte

            structure_utilise_compte = r'''{|l|c|c|c|c|c|c|c|c|c|c|c|}'''
            legende_utilise_compte = r'''Procédés (machine + main d'oeuvre) pour compte ''' + intitule_compte

            contenu_utilise_compte = r'''
                \cline{3-11}
                \multicolumn{2}{c}{} & \multicolumn{2}{|c|}{Machine} & \multicolumn{2}{l|}{PU [CHF/h]} & \multicolumn{5}{l|}{Montant [CHF]} \\
                \hline
                ''' + intitule_compte + r''' & & Mach. & Oper. & Mach. & Oper. & Machine & Déduc. Sp. & Déduc. HC & Net & M.O. \\
                \hline
                '''

            dico_utilise_compte = {'maij': "%.2f" % sco['somme_j_mai'], 'dsij': "%.2f" % sco['somme_j_dsi'],
                                   'dhij': "%.2f" % sco['somme_j_dhi'], 'moij': "%.2f" % sco['somme_j_moi'],
                                   'nmij': "%.2f" % (sco['somme_j_mai']-sco['somme_j_mr'])
                                }
            contenu_utilise_compte += r'''
                \multicolumn{6}{|r|}{Total} & %(maij)s & %(dsij)s & %(dhij)s & %(nmij)s & %(moij)s \\
                \hline
                ''' % dico_utilise_compte

            contenu_compte_annexe3 += Latex.tableau(contenu_utilise_compte, structure_utilise_compte, legende_utilise_compte)

            # ## Annexe 4

            dico_nom = {'labo': Latex.echappe_caracteres(client['abrev_labo']),
                        'compte': Latex.echappe_caracteres(compte['intitule']),
                        'date': edition.mois_txt + " " + str(edition.annee)}
            contenu_compte_annexe4 += r'''
                \clearpage
                Annexe 4 : %(labo)s - %(compte)s - %(date)s
                ''' % dico_nom

            # ## T1 : prestations livrées compte

            structure_prestations_compte = r'''{|l|l|c|c|c|c|c|c|}'''
            legende_prestations_compte = r'''Détails prestations livrées pour compte ''' + intitule_compte

            contenu_prestations_compte = r'''
                \hline
                \multicolumn{2}{|l|}{''' + intitule_compte + r'''} & Q & Unité & PU & Montant & Rabais & Total \\
                \hline
                '''
            i = 0
            for article in generaux.articles_d3:
                dico_prestations_compte = {'smj': "%.2f" % sco['sommes_cat_m'][article.code_d],
                                           'srj': "%.2f" % sco['sommes_cat_r'][article.code_d],
                                           'sj': "%.2f" % sco['tot_cat'][article.code_d]}
                if i == 0:
                    i += 1
                else:
                    contenu_prestations_compte += r'''\multicolumn{8}{c}{} \\
                        \hline
                        '''
                contenu_prestations_compte += r'''
                    \multicolumn{5}{|l|}{''' + article.intitule_long + r'''} & %(smj)s & %(srj)s & %(sj)s \\
                    \hline
                    ''' % dico_prestations_compte

            contenu_compte_annexe4 += Latex.tableau(contenu_prestations_compte, structure_prestations_compte, legende_prestations_compte)

            # ## T2 : coûts éligibles

            structure_eligibles_compte = r'''{|c|c|c|c|}'''
            for i in range(1,4):
                legende_eligibles_compte = r'''Coûts d'utilisation (U''' + str(i) + r''') pour compte ''' + intitule_compte

                contenu_eligibles_compte = r'''
                    \hline
                    Mach. U''' + str(i) + r''' & M.O. & Prest. Livr. & Total \\
                    \hline
                    '''

                contenu_compte_annexe4 += Latex.tableau(contenu_eligibles_compte, structure_eligibles_compte, legende_eligibles_compte)

            # ## T3 : coûts éligibles machines

            structure_machines_compte = r'''{|l|c|c|c|c|c|c|c|}'''
            for i in range(1,4):
                legende_machines_compte = r'''Coûts procédés éligibles (U''' + str(i) + r''') pour compte ''' + intitule_compte

                contenu_machines_compte = r'''
                    \cline{3-8}
                    \multicolumn{2}{r}{} & \multicolumn{2}{|c|}{Machine} & \multicolumn{2}{c|}{PU [CHF/h]} & \multicolumn{2}{c|}{Montant [CHF]} \\
                    \hline
                    ''' + intitule_compte + r''' & Mach. & Oper. & PU''' + str(i) + r''' & M.O. & Mach.& U''' + str(i) + r''' & M.O. \\
                    \hline
                    '''

                contenu_machines_compte += r'''
                    \multicolumn{6}{|r|}{Total} & & \\
                    \hline
                    '''

                contenu_compte_annexe4 += Latex.tableau(contenu_machines_compte, structure_machines_compte, legende_machines_compte)

            # ## T4 : coûts éligibles prestations

            structure_livraisons_compte = r'''{|l|c|}'''
            legende_livraisons_compte = r'''Prestations livrées éligibles pour compte ''' + intitule_compte

            contenu_livraisons_compte = r'''
                \hline
                '''
            total = 0
            for article in generaux.articles_d3:
                total += sco['tot_cat'][article.code_d]
                contenu_livraisons_compte += r'''
                    ''' + article.intitule_long + r'''&''' + "%.2f" % sco['tot_cat'][article.code_d] + r''' \\
                    \hline
                    '''
            contenu_livraisons_compte += r'''
                Total Pres. Livr. &''' + "%.2f" % total + r''' \\
                \hline
                '''

            contenu_compte_annexe4 += Latex.tableau(contenu_livraisons_compte, structure_livraisons_compte, legende_livraisons_compte)

            # ## compte

        # ## Début des tableaux

        # ## Annexe 1
        dic_entete = {'code': code_client, 'code_sap': client['code_sap'],
                      'nom': Latex.echappe_caracteres(client['abrev_labo']),
                      'date': edition.mois_txt + " " + str(edition.annee)}
        entete = r'''
            Annexe 1 : %(code)s - %(code_sap)s - %(nom)s - %(date)s
            ''' % dic_entete

        contenu += entete

        # ## récapitulatif postes pour client

        structure_recap_poste_cl = r'''{|l|r|r|r|}'''
        legende_recap_poste_cl = r'''Récapitulatif postes pour client ''' + intitule_client

        dico_recap_poste_cl = {'emom': "%.2f" % scl['em'], 'emor': "%.2f" % scl['er'], 'emo': "%.2f" % scl['e'],
                               'resm': "%.2f" % scl['rm'], 'resr': "%.2f" % scl['rr'], 'res': "%.2f" % scl['r'],
                               'mat': "%.2f" % scl['mat'], 'stmm': "%.2f" % scl['somme_t_mm'],
                               'stmr': "%.2f" % scl['somme_t_mr'], 'mt': "%.2f" % scl['mt'],
                               'mot': "%.2f" % scl['mot'], 'matr': "%.2f" % (scl['mat']-scl['somme_t_mr'])}

        contenu_recap_poste_cl = r'''
            \hline
             & \multicolumn{1}{l|}{Montant} & \multicolumn{1}{l|}{Rabais} & \multicolumn{1}{l|}{Total} \\
            \hline
            Emolument & %(emom)s & %(emor)s & %(emo)s \\
            \hline
            Frais de réservation & %(resm)s & %(resr)s & %(res)s \\
            \hline
            Procédés & %(stmm)s & %(stmr)s & %(mt)s \\
            \hline
            \hspace{5mm} \textit{Machine} & \textit{%(mat)s} & \textit{%(stmr)s} & \textit{%(matr)s} \\
            \hline
            \hspace{5mm} \textit{Main d'oeuvre} & \textit{%(mot)s} & \textit{0.00} & \textit{%(mot)s}  \\
            \hline
            ''' % dico_recap_poste_cl

        for article in generaux.articles_d3:
            contenu_recap_poste_cl += Latex.echappe_caracteres(article.intitule_long)
            contenu_recap_poste_cl += r''' & '''
            contenu_recap_poste_cl += "%.2f" % scl['sommes_cat_m'][article.code_d] + r''' & '''
            contenu_recap_poste_cl += "%.2f" % scl['sommes_cat_r'][article.code_d] + r''' & '''
            contenu_recap_poste_cl += "%.2f" % scl['tot_cat'][article.code_d] + r''' \\
                \hline
                '''

        contenu_recap_poste_cl += r'''\multicolumn{3}{|r|}{Total} & ''' + "%.2f" % (scl['somme_t'] + scl['e']) + r'''\\
            \hline
            '''

        contenu += Latex.tableau(contenu_recap_poste_cl, structure_recap_poste_cl, legende_recap_poste_cl)

        # ## émolument pour client

        structure_emolument = r'''{|r|r|l|r|r|r|r|}'''
        legende_emolument = r'''Emolument pour client ''' + intitule_client

        dico_emolument = {'emb':  "%.2f" % client['emol_base_mens'], 'ef':  "%.2f" % client['emol_fixe'],
                          'pente': client['coef'], 'tot_eq_r': "%.2f" % scl['r'],
                          'tot_eq_m': "%.2f" % (scl['mt']-scl['mot']), 'tot_eq': "%.2f" % scl['somme_eq'],
                          'rabais': "%.2f" % scl['er']}

        contenu_emolument = r'''
            \hline
            \multicolumn{1}{|l|}{Emolument de base} & \multicolumn{1}{l|}{Emolument fixe} & Pente
            & \multicolumn{1}{l|}{Total EQ R} & \multicolumn{1}{l|}{Total EQ M} & \multicolumn{1}{l|}{Total EQ} &
            \multicolumn{1}{l|}{Rabais émolument} \\
            \hline
            %(emb)s & %(ef)s & %(pente)s & %(tot_eq_r)s & %(tot_eq_m)s & %(tot_eq)s & %(rabais)s \\
            \hline
            ''' % dico_emolument

        contenu += Latex.tableau(contenu_emolument, structure_emolument, legende_emolument)

        # ## pénalités réservation

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

            contenu += Latex.tableau(contenu_frais_client, structure_frais_client, legende_frais_client)

        # ## coûts procédés

        structure_procedes_client = r'''{|l|c|c|c|c|c|}'''
        legende_procedes_client = r'''Récapitulatif des coûts procédés pour client : ''' + intitule_client

        dico_procedes_client = {'mat': "%.2f" % scl['mat'], 'dti': "%.2f" % scl['somme_t_mr'],
                                'nmi': "%.2f" % (scl['mat']-scl['somme_t_mr']),
                                'mot': "%.2f" % scl['mot'], 'mt': "%.2f" % scl['mt']}
        contenu_procedes_client += r'''
            Total article & %(mat)s & %(dti)s & %(nmi)s & %(mot)s & %(mt)s \\
            \hline
            ''' % dico_procedes_client

        contenu += Latex.tableau(contenu_procedes_client, structure_procedes_client, legende_procedes_client)

        # ## prestations livrées

        structure_prestations_client = r'''{|l|c|c|c|}'''
        legende_prestations_client = r'''Récapitulatif des prestations livrées pour client : ''' + intitule_client

        contenu_prestations_client = ""
        i = 0
        for article in generaux.articles_d3:
            dico_prestations_client = {'cmt': "%.2f" % scl['sommes_cat_m'][article.code_d],
                                       'crt': "%.2f" % scl['sommes_cat_r'][article.code_d],
                                       'ct': "%.2f" % scl['tot_cat'][article.code_d]}
            contenu_prestations_client_tab[article.code_d] += r'''
            Total article & %(cmt)s & %(crt)s & %(ct)s \\
            \hline
            ''' % dico_prestations_client
            if i == 0:
                i += 1
            else:
                contenu_prestations_client += r'''\multicolumn{4}{c}{} \\'''
            contenu_prestations_client += contenu_prestations_client_tab[article.code_d]

        contenu += Latex.tableau(contenu_prestations_client, structure_prestations_client, legende_prestations_client)

        # ## récapitulatif comptes pour client

        legende_recap_compte = r'''Récapitulatif des comptes pour client ''' + intitule_client

        dico_recap_compte = {'procedes': "%.2f" % scl['mt'], 'total': "%.2f" % (scl['somme_t']-scl['r'])}

        contenu_recap_compte += r'''Total article & %(procedes)s''' % dico_recap_compte

        for categorie in generaux.codes_d3():
            contenu_recap_compte += r''' & ''' + "%.2f" % scl['tot_cat'][categorie]

        contenu_recap_compte += r'''& %(total)s \\
            \hline
            ''' % dico_recap_compte

        contenu += Latex.tableau(contenu_recap_compte, structure_recap_compte, legende_recap_compte)

        # ## Annexe 2
        dic_entete = {'code': code_client, 'code_sap': client['code_sap'],
                      'nom': Latex.echappe_caracteres(client['abrev_labo']),
                      'date': edition.mois_txt + " " + str(edition.annee)}
        entete = r'''
            \clearpage
            Annexe 2 : %(code)s - %(code_sap)s - %(nom)s - %(date)s
            ''' % dic_entete

        contenu += entete

        # ## T1 : durée réservée client

        if code_client in reservations.sommes:
            structure_reserve_client = r'''{|c|c|c|c|c|}'''
            legende_reserve_client = r'''Détail réservations machines pour client : ''' + intitule_client

            contenu_reserve_client = r'''
                \cline{4-5}
                \multicolumn{3}{c}{} & \multicolumn{2}{|c|}{Durée réservée} \\
                \cline{4-5}
                \multicolumn{3}{c|}{} & HP & HC \\
                \hline
                '''

            somme = reservations.sommes[code_client]['machines']

            machines_reservees= {}
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
                        \multicolumn{3}{|l|}{\textbf{%(machine)s}} & %(hp)s & %(hc)s \\
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
                            dico_user = {'user': smu['prenom'] + " " + smu['nom'],
                                         'hp': Outils.format_heure(smu['res_hp']),
                                         'hc': Outils.format_heure(smu['res_hc'])}
                            contenu_reserve_client += r'''
                                \multicolumn{3}{|l|}{%(user)s} & %(hp)s & %(hc)s \\
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
                                    %(date)s & %(heure)s & %(sup)s & %(hp)s & %(hc)s \\
                                    \hline
                                ''' % dico_pos

            contenu += Latex.tableau(contenu_reserve_client, structure_reserve_client, legende_reserve_client)

        # ## T2 : durée utilisée client

        structure_utilise_client = r'''{|l|c|c|}'''
        legende_utilise_client = r'''Récapitulatif heures machines par utilisateur pour client : ''' + intitule_client

        contenu_utilise_client = r'''
            \cline{2-3}
            \multicolumn{1}{c|}{} & HP & HC \\
            \hline
            '''

        contenu += Latex.tableau(contenu_utilise_client, structure_utilise_client, legende_utilise_client)

        contenu += contenu_compte_annexe3
        contenu += contenu_compte_annexe4

        return contenu

    @staticmethod
    def ligne_cae(cae, machine, coefmachine):
        """
        création d'une ligne de tableau pour un accès
        :param cae: accès particulier
        :param machine: machine concernée
        :param coefmachine: coefficients machine concernée
        :return: ligne de tableau latex
        """

        login = Latex.echappe_caracteres(cae['date_login']).split()
        temps = login[0].split('-')
        date = temps[0]
        for pos in range(1, len(temps)):
            date = temps[pos] + '.' + date
        if len(login) > 1:
            heure = login[1]
        else:
            heure = ""

        mai_hp = round(cae['duree_machine_hp']/60 * cae['pum'], 2)
        mai_hc = round(cae['duree_machine_hc']/60 * cae['pum'], 2)
        dsi_hp = round(cae['duree_machine_hp']/60 * coefmachine['coef_d'] * machine['d_h_machine_d'], 2)
        dsi_hc = round(cae['duree_machine_hc']/60 * coefmachine['coef_d'] * machine['d_h_machine_d'], 2)
        dhi = cae['dhi']
        moi_hp = round(cae['duree_operateur_hp']/60 * cae['puo_hp'], 2)
        moi_hc = round(cae['duree_operateur_hc']/60 * cae['puo_hc'], 2)
        m_hp = mai_hp + moi_hp - dsi_hp
        m_hc = mai_hc + moi_hc - dsi_hc - dhi

        dico = {'date': date, 'heure': heure,
                'machine': Latex.echappe_caracteres(cae['nom_machine']),
                'operateur': Latex.echappe_caracteres(cae['nom_op']),
                'rem_op': Latex.echappe_caracteres(cae['remarque_op']),
                'rem_staff': Latex.echappe_caracteres(cae['remarque_staff']),
                'tm_hp': Outils.format_heure(cae['duree_machine_hp']),
                'to_hp': Outils.format_heure(cae['duree_operateur_hp']),
                'tm_hc': Outils.format_heure(cae['duree_machine_hc']),
                'to_hc': Outils.format_heure(cae['duree_operateur_hc']),
                'pum': cae['pum'], 'puo_hp': cae['puo_hp'], 'puo_hc': cae['puo_hc'], 'mai_hp': "%.2f" % mai_hp,
                'mai_hc': "%.2f" % mai_hc, 'dsi_hp': Outils.format_si_nul(dsi_hp),
                'dsi_hc': Outils.format_si_nul(dsi_hc), 'dhi_hp': '-', 'dhi_hc': Outils.format_si_nul(dhi),
                'moi_hp': "%.2f" % moi_hp, 'moi_hc': "%.2f" % moi_hc, 'm_hp': "%.2f" % m_hp, 'm_hc': "%.2f" % m_hc}

        nb = 0
        if (cae['duree_machine_hp'] > 0) or (cae['duree_operateur_hp'] > 0):
            nb += 1

        if (cae['duree_machine_hc'] > 0) or (cae['duree_operateur_hc'] > 0):
            nb += 1

        if nb == 0:
            return "", [0, 0, 0, 0, 0]

        if (cae['remarque_staff'] != "") or (cae['remarque_op'] != ""):
            nb += 1

        if nb == 1:
            ligne = r'''%(date)s & %(heure)s''' % dico
        else:
            ligne = r'''\multirow{''' + str(nb) + r'''}{*}{%(date)s} & \multirow{''' % dico
            ligne += str(nb) + r'''}{*}{%(heure)s}''' % dico

        nb = 0
        if (cae['duree_machine_hp'] > 0) or (cae['duree_operateur_hp'] > 0):
            ligne += r''' & %(machine)s & HP & %(tm_hp)s & %(to_hp)s & %(pum)s & %(puo_hp)s & %(mai_hp)s &
                %(dsi_hp)s & %(dhi_hp)s & %(moi_hp)s & %(m_hp)s \\
                ''' % dico
            nb += 1

        if (cae['duree_machine_hc'] > 0) or (cae['duree_operateur_hc'] > 0):
            if nb > 0:
                ligne += r'''& &'''
            else:
                ligne += r'''& %(machine)s ''' % dico
            ligne += r''' & HC & %(tm_hc)s & %(to_hc)s & %(pum)s & %(puo_hc)s & %(mai_hc)s &
                %(dsi_hc)s & %(dhi_hc)s & %(moi_hc)s & %(m_hc)s \\
                ''' % dico

        if (cae['remarque_staff'] != "") or (cae['remarque_op'] != ""):
            ligne += r'''\cline{3-13}
                &  & \multicolumn{11}{l|}{%(operateur)s ; %(rem_op)s ; %(rem_staff)s}\\
                ''' % dico

        ligne += r'''\hline
            '''
        return ligne, [(mai_hp + mai_hc), (dsi_hp + dsi_hc), dhi, (moi_hp + moi_hc), (m_hp + m_hc)]

    @staticmethod
    def ligne_res(res):
        """
        création d'une ligne de tableau pour une réservation
        :param res: réservation particulière
        :return: ligne de tableau latex
        """
        login = Latex.echappe_caracteres(res['date_debut']).split()
        temps = login[0].split('-')
        date = temps[0]
        for pos in range(1, len(temps)):
            date = temps[pos] + '.' + date
        if len(login) > 1:
            heure = login[1]
        else:
            heure = ""

        dico = {'date': date, 'heure': heure,
                'machine': Latex.echappe_caracteres(res['nom_machine']),
                'reserve': Latex.echappe_caracteres(res['date_reservation']),
                'supprime': Latex.echappe_caracteres(res['date_suppression']),
                'shp': Outils.format_heure(res['duree_hp']), 'shc': Outils.format_heure(res['duree_hc']),
                'fhp': Outils.format_heure(res['duree_fact_hp']), 'fhc': Outils.format_heure(res['duree_fact_hc'])}

        nb = 0
        if res['duree_fact_hp'] > 0:
            nb += 1

        if res['duree_fact_hc'] > 0:
            nb += 1

        if nb == 0:
            return ""

        if res['date_suppression'] != "":
            nb += 1

        if nb == 1:
            ligne = r'''%(date)s & %(heure)s''' % dico
        else:
            ligne = r'''\multirow{''' + str(nb) + r'''}{*}{%(date)s} & \multirow{''' % dico
            ligne += str(nb) + r'''}{*}{%(heure)s}''' % dico

        nb = 0
        if res['duree_fact_hp'] > 0:
            ligne += r''' & %(machine)s & HP & %(shp)s & %(fhp)s\\
                ''' % dico
            nb += 1

        if res['duree_fact_hc'] > 0:
            if nb > 0:
                ligne += r'''& &'''
            else:
                ligne += r'''& %(machine)s ''' % dico
            ligne += r''' & HC & %(shc)s & %(fhc)s \\
                ''' % dico

        if res['date_suppression'] != "":
            ligne += r'''\cline{3-6}
                &  & \multicolumn{4}{l|}{Supprimé le : %(supprime)s} \\
                ''' % dico

        ligne += r'''\hline
            '''

        return ligne

    @staticmethod
    def ligne_liv(livraison):
        """
        création d'une ligne de tableau pour une livraison
        :param livraison: livraison particulière
        :return: ligne de tableau latex
        """
        total = livraison['montant'] - livraison['rabais_r']
        dico = {'date': Latex.echappe_caracteres(livraison['date_livraison']),
                'prestation': Latex.echappe_caracteres(livraison['designation']),
                'quantite': livraison['quantite'], 'unite': Latex.echappe_caracteres(livraison['unite']),
                'rapport': "%.2f" % livraison['prix_unit_client'], 'montant': "%.2f" % livraison['montant'],
                'rabais': "%.2f" % livraison['rabais_r'], 'total': "%.2f" % total, 'id': livraison['id_livraison'],
                'responsable': Latex.echappe_caracteres(livraison['responsable']),
                'commande': Latex.echappe_caracteres(livraison['date_commande']),
                'remarque': Latex.echappe_caracteres(livraison['remarque'])}

        return r'''\multirow{2}{*}{%(date)s} & %(prestation)s & %(quantite)s & %(unite)s & %(rapport)s & %(montant)s &
            %(rabais)s & %(total)s \\
            \cline{2-8}
             & \multicolumn{7}{l|}{Commande: %(commande)s; N. livraison: %(id)s; Resp: %(responsable)s; Remarque:
             %(remarque)s} \\
             \hline
             ''' % dico, total
