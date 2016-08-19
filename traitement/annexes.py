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

        client_comptes = sommes.sommes_comptes[code_client]
        contenu_compte_annexe3 = ""
        contenu_compte_annexe4 = ""

        for id_compte in sorted(client_comptes.keys()):

            # ## COMPTE

            sco = sommes.sommes_comptes[code_client][id_compte]
            compte = comptes.donnees[id_compte]
            intitule_compte = id_compte + " - " + Latex.echappe_caracteres(compte['intitule'])

            # ## Annexe 3

            dico_nom = {'labo': Latex.echappe_caracteres(client['abrev_labo']),
                        'compte': Latex.echappe_caracteres(compte['intitule']),
                        'date': edition.mois_txt + " " + str(edition.annee)}
            contenu_compte_annexe3 += r'''
                \clearpage
                Annexe 3 : %(labo)s - %(compte)s - %(date)s
                ''' % dico_nom

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

            # ## récapitulatif postes pour compte

            structure_recap_poste = r'''{|l|r|r|r|}'''
            legende_recap_poste = r'''Récapitulatif postes pour compte ''' + intitule_compte

            dico_recap_poste = {'maij': "%.2f" % sco['somme_j_mai'], 'moij': "%.2f" % sco['somme_j_moi'],
                                'mrj': "%.2f" % sco['somme_j_mr'], 'mj': "%.2f" % sco['mj']}

            contenu_recap_poste = r'''
                \hline
                Compte : ''' + intitule_compte + r''' & \multicolumn{1}{l|}{Montant} & \multicolumn{1}{l|}{Rabais}
                & \multicolumn{1}{l|}{Total} \\
                \hline
                Machine & %(maij)s & \multirow{2}{*}{%(mrj)s} & \multirow{2}{*}{%(mj)s} \\
                \cline{1-2}
                Main d'oeuvre & %(moij)s &  &  \\
                \hline
                ''' % dico_recap_poste

            for article in generaux.articles_d3:
                categorie = article.code_d
                contenu_recap_poste += Latex.echappe_caracteres(article.intitule_long)
                contenu_recap_poste += r''' & ''' + "%.2f" % sco['sommes_cat_m'][categorie]
                contenu_recap_poste += r''' & ''' + "%.2f" % sco['sommes_cat_r'][categorie]
                contenu_recap_poste += r''' & ''' + "%.2f" % sco['tot_cat'][categorie]
                contenu_recap_poste += r''' \\
                    \hline
                    '''

            contenu_compte_annexe3 += Latex.tableau(contenu_recap_poste, structure_recap_poste, legende_recap_poste)

            # ## T2 : durée utilisée compte

            # ## Annexe 4

            dico_nom = {'labo': Latex.echappe_caracteres(client['abrev_labo']),
                        'compte': Latex.echappe_caracteres(compte['intitule']),
                        'date': edition.mois_txt + " " + str(edition.annee)}
            contenu_compte_annexe4 += r'''
                \clearpage
                Annexe 4 : %(labo)s - %(compte)s - %(date)s
                ''' % dico_nom

            # ## T1 : prestations livrées compte

            # ## T2 : coûts éligibles

            # ## T3 : coûts éligibles machines

            # ## T4 : coûts éligibles prestations

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
            categorie = article.code_d
            contenu_recap_poste_cl += Latex.echappe_caracteres(article.intitule_long)
            contenu_recap_poste_cl += r''' & '''
            contenu_recap_poste_cl += "%.2f" % scl['sommes_cat_m'][categorie] + r''' & '''
            contenu_recap_poste_cl += "%.2f" % scl['sommes_cat_r'][categorie] + r''' & '''
            contenu_recap_poste_cl += "%.2f" % scl['tot_cat'][categorie] + r''' \\
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
            structure_frais_client = r'''{|l|c|r|r|r|}'''
            legende_frais_client = r'''Frais de réservation/utilisation par machine pour client : ''' + intitule_client

            contenu_frais_client = r'''
                    \hline
                    Equipement & & Pénalités [h] & PU & Montant \\
                    \hline
                    '''
            machines_utilisees = {}
            for key in scl['res']:
                machines_utilisees[key] = {'machine': machines.donnees[key]['nom']}

            for machine_t in sorted(machines_utilisees.items(), key=lambda k_v: k_v[1]['machine']):
                machine = machine_t[1]
                id_mach = machine_t[0]
                som_m = scl['res'][id_mach]
                if som_m['pen_hp'] > 0 or som_m['pen_hc'] > 0:

                    dico_frais_client = {
                        'machine': Latex.echappe_caracteres(machine['machine']),
                        'pen_hp': "%.1f" % som_m['pen_hp'], 'pen_hc': "%.1f" % som_m['pen_hc'],
                        'mont_hp': Outils.format_si_nul(som_m['m_hp']), 'mont_hc': Outils.format_si_nul(som_m['m_hc']),
                        'pu_hp': Outils.format_si_nul(reservations.sommes[code_client]['machines'][id_mach]['pu_hp']),
                        'pu_hc': Outils.format_si_nul(reservations.sommes[code_client]['machines'][id_mach]['pu_hc'])}

                    if som_m['pen_hp'] > 0:
                        contenu_frais_client += r'''%(machine)s & HP &  %(pen_hp)s & %(pu_hp)s & %(mont_hp)s \\
                            \hline
                            ''' % dico_frais_client

                    if som_m['pen_hc'] > 0:
                        contenu_frais_client += r'''%(machine)s & HC &  %(pen_hc)s & %(pu_hc)s & %(mont_hc)s \\
                            \hline
                            ''' % dico_frais_client

            contenu_frais_client += r'''
                \multicolumn{4}{|r|}{Total} & ''' + Outils.format_si_nul(scl['r']) + r'''\\
                \hline
                '''

        #    contenu += Latex.tableau(contenu_frais_client, structure_frais_client, legende_frais_client)

        # ## T4 : coûts procédés

        # ## T5 : prestations livrées

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

        # ## T2 : durée utilisée client

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
