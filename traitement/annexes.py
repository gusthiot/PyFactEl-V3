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
                dossier_annexe, plateforme, coefprests, coefmachines, generaux, nom_dossier='deprecated'):
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
        :param nom_dossier: nom du dossier dans lequel enregistrer le dossier des annexes
        """
        # dossier_annexe = Outils.chemin_dossier([nom_dossier, "annexes"], plateforme, generaux)
        prefixe = "annexe_"
        garde = r'''Annexes factures \newline Billing Appendices'''

        Annexes.creation_annexes(sommes, clients, edition, livraisons, acces, machines, reservations, prestations,
                                 comptes, dossier_annexe, plateforme, prefixe, coefprests, coefmachines, generaux, garde)

        """
        # tant que les annexes techniques et les annexes de factures sont identiques
        dossier_annexe_t = Outils.chemin_dossier([nom_dossier, "annexes_techniques"], plateforme, generaux)
        prefixe_t = "annexeT_"
        for file_t in os.listdir(dossier_annexe_t):
            if file_t.endswith(".pdf"):
                file = file_t.replace(prefixe_t, prefixe)
                shutil.copyfile(dossier_annexe_t + file_t, dossier_annexe + file)
        """

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
        sca = sommes.sommes_categories[code_client]
        intitule_client = code_client + " - " + Latex.echappe_caracteres(client['abrev_labo'])

        structure_recap_compte = r'''{|l|l|r|r|'''
        contenu_recap_compte = r'''
            \hline
            Compte & Catégorie & \multicolumn{1}{l|}{Procédés}'''

        for categorie in generaux.codes_d3():
            structure_recap_compte += r'''r|'''
            contenu_recap_compte += r''' & \multicolumn{1}{l|}{
            ''' + Latex.echappe_caracteres(coefprests.obtenir_noms_categories(categorie)) + r'''}'''

        structure_recap_compte += r'''}'''
        legende_recap_compte = r'''Récapitulatif des comptes pour client ''' + intitule_client
        contenu_recap_compte += r'''& \multicolumn{1}{l|}{Total cpte} \\
            \hline
            '''

        structure_recap_eligibles = r'''{|l|c|c|c|c|c|c|c|c|}'''
        legende_recap_eligibles = r'''Coûts procédés éligibles des comptes pour client ''' + intitule_client
        contenu_recap_eligibles = r'''
            \hline
            Compte & Coûts A & Coûts B & Coûts C & Coûts U1 & Coûts U2 & Coûts U3 & Coûts MO & Coûts CO \\
            \hline
            '''

        client_comptes = sommes.sommes_comptes[code_client]
        contenu_compte = ""

        for id_compte in sorted(client_comptes.keys()):
            # ## COMPTE

            compte = comptes.donnees[id_compte]
            intitule_compte = id_compte + " - " + Latex.echappe_caracteres(compte['intitule'])
            dico_nom = {'labo': Latex.echappe_caracteres(client['abrev_labo']),
                        'utilisateur': Latex.echappe_caracteres(compte['intitule']),
                        'date': edition.mois_txt + " " + str(edition.annee)}
            contenu_compte += r'''
                \clearpage
                %(labo)s - %(utilisateur)s - %(date)s
                ''' % dico_nom

            structure_recap_projet = r'''{|l|r|r|'''
            contenu_recap_projet = r'''
                \hline
                Projet & \multicolumn{1}{l|}{Procédés} '''
            for categorie in generaux.codes_d3():
                structure_recap_projet += r'''r|'''
                contenu_recap_projet += r''' & \multicolumn{1}{l|}{
                ''' + Latex.echappe_caracteres(coefprests.obtenir_noms_categories(categorie)) + r'''}'''
            structure_recap_projet += r'''}'''
            legende_recap_projet = r'''Récapitulatif compte ''' + intitule_compte
            contenu_recap_projet += r''' & \multicolumn{1}{l|}{Total projet} \\
                \hline
                '''
            client_compte_projet = sommes.sommes_projets[code_client][id_compte]
            contenu_projet = ""

            # ## RES
            structure_res = r'''{|l|l|l|l|l|l|}'''
            dico_res = {'compte': intitule_compte}
            contenu_res = r'''
                \hline
                \multicolumn{3}{|l|}{%(compte)s} & & \multicolumn{2}{l|}{hh:mm} \\
                \hline
                Date & Heure & Equipement & & slot & fact. \\
                \hline
                ''' % dico_res
            nombre_res = 0
            legende_res = r'''Récapitulatif Réservations : ''' + intitule_compte

            res_proj = reservations.reservations_pour_compte(id_compte, code_client)
            for res in res_proj:
                nombre_res += 1
                ligne = Annexes.ligne_res(res)
                contenu_res += ligne

            contenu_rsv = ""
            if nombre_res > 0:
                contenu_rsv = Latex.long_tableau(contenu_res, structure_res, legende_res)
            # ## res

            for num_projet in sorted(client_compte_projet.keys()):
                # ## PROJET
                sp = sommes.sommes_projets[code_client][id_compte][num_projet]
                intitule_projet = num_projet + " - " + Latex.echappe_caracteres(sp['intitule'])

                dico_recap_projet = {'num': intitule_projet, 'procede': "%.2f" % sp['mp']}

                total = sp['mp']
                contenu_recap_projet += r'''
                    \hline
                    %(num)s & %(procede)s''' % dico_recap_projet
                for categorie in generaux.codes_d3():
                    total += sp['tot_cat'][categorie]
                    contenu_recap_projet += r''' & ''' + "%.2f" % sp['tot_cat'][categorie]
                dico_recap_projet['total'] = "%.2f" % total

                contenu_recap_projet += r''' & %(total)s \\
                    \hline
                    ''' % dico_recap_projet

                ## CAE
                structure_cae = r'''{|l|l|l|c|c|c|c|c|c|c|c|c|c|}'''
                dico_cae = {'compte': intitule_compte, 'projet': intitule_projet}
                contenu_cae = r'''
                    \hline
                    \multicolumn{3}{|l|}{%(compte)s / %(projet)s} & & \multicolumn{2}{c|}{hh:mm} &
                    \multicolumn{2}{c|}{CHF/h} & Montant & \multicolumn{2}{c|}{Déductions} & Montant & Montant \\
                    \cline{1-8}
                    \cline{10-11}
                    Date & Heure & Equipement & & mach. & oper. & mach. & MO & machine &
                    spé. & HC & MO & net \\
                    \hline
                    ''' % dico_cae
                nombre_cae = 0
                legende_cae = r'''Récapitulatif Utilisation machines : ''' + intitule_compte + r''' / ''' +\
                              intitule_projet

                cae_proj = acces.acces_pour_projet(num_projet, id_compte, code_client)
                resultats = [0, 0, 0, 0, 0]
                for cae in cae_proj:
                    nombre_cae += 1
                    machine = machines.donnees[cae['id_machine']]
                    coefmachine = coefmachines.donnees[client['id_classe_tarif'] + machine['categorie']]
                    ligne, resultat = Annexes.ligne_cae(cae, machine, coefmachine)
                    resultats[0] += resultat[0]
                    resultats[1] += resultat[1]
                    resultats[2] += resultat[2]
                    resultats[3] += resultat[3]
                    resultats[4] += resultat[4]
                    contenu_cae += ligne

                contenu_cae += r'''
                    \multicolumn{8}{|r|}{Total} & ''' + Outils.format_si_nul(resultats[0]) + r'''
                    & ''' + Outils.format_si_nul(resultats[1]) + r'''
                    & ''' + Outils.format_si_nul(resultats[2]) + r'''
                    & ''' + Outils.format_si_nul(resultats[3]) + r'''
                    & ''' + Outils.format_si_nul(resultats[4]) + r''' \\
                    \hline
                    '''

                if nombre_cae > 0:
                    contenu_projet += Latex.long_tableau(contenu_cae, structure_cae, legende_cae)
                ## cae

                # ## LIV
                structure_liv = r'''{|l|l|l|l|r|r|r|r|}'''
                dico_liv = {'compte': intitule_compte, 'projet': intitule_projet}
                contenu_liv = r'''
                    \hline
                    \multicolumn{2}{|l|}{%(compte)s / %(projet)s} & & & & & &  \\
                    \hline
                    Date livr. & Désignation & Q & Unité & \multicolumn{1}{l|}{PU} & \multicolumn{1}{l|}{Montant}
                    & \multicolumn{1}{l|}{Rabais} & \multicolumn{1}{l|}{Total} \\
                    \hline
                    ''' % dico_liv
                nombre_liv = 0
                legende_liv = r'''Récapitulatif Livraisons : ''' + intitule_compte + r''' / ''' + intitule_projet

                liv_proj_cat = livraisons.livraisons_pour_projet_par_categorie(num_projet, id_compte, code_client,
                                                                               prestations)
                resultats = 0
                for categorie in generaux.codes_d3():
                    if categorie in liv_proj_cat:
                        livs = liv_proj_cat[categorie]
                        for liv in livs:
                            nombre_liv += 1
                            ligne, resultat = Annexes.ligne_liv(liv)
                            resultats += resultat
                            contenu_liv += ligne

                contenu_liv += r'''
                    \multicolumn{7}{|r|}{Total} & ''' + Outils.format_si_nul(resultats) + r'''\\
                    \hline
                    '''

                if nombre_liv > 0:
                    contenu_projet += Latex.long_tableau(contenu_liv, structure_liv, legende_liv)
                # ## liv

                # ## projet

            sco = sommes.sommes_comptes[code_client][id_compte]

            ligne = r'''\hline
                Total article & ''' + "%.2f" % sco['mj']

            sj = sco['mj']

            for categorie in generaux.codes_d3():
                ligne += r''' & ''' + "%.2f" % sco['tot_cat'][categorie]
                sj += sco['tot_cat'][categorie]

            ligne += r''' & ''' + "%.2f" % sj
            ligne += r'''\\
                \hline
                '''

            contenu_recap_projet += ligne

            contenu_compte += Latex.tableau(contenu_recap_projet, structure_recap_projet, legende_recap_projet)

            dico_recap_compte = {'compte': intitule_compte, 'type': compte['categorie'], 'procede': "%.2f" % sco['mj'],
                                 'total': "%.2f" % sj}

            contenu_recap_compte += r'''%(compte)s & %(type)s & %(procede)s ''' \
                                    % dico_recap_compte

            for categorie in generaux.codes_d3():
                contenu_recap_compte += r''' & ''' + "%.2f" % sco['tot_cat'][categorie]

            contenu_recap_compte += r'''& %(total)s \\
                    \hline
                    ''' % dico_recap_compte

            u1 = sco['somme_j_ai']
            u2 = u1 + sco['somme_j_bi']
            u3 = u2 + sco['somme_j_ci']
            couts_co = 0
            for cat, tt in sco['tot_cat'].items():
                couts_co += tt

            dico_recap_eligibles = {'compte': intitule_compte, 'couts_a': "%.2f" % sco['somme_j_ai'],
                                    'couts_mo': "%.2f" % sco['somme_j_oi'], 'couts_b': "%.2f" % sco['somme_j_bi'],
                                    'couts_c': "%.2f" % sco['somme_j_ci'], 'u1': "%.2f" % u1, 'u2': "%.2f" % u2,
                                    'u3': "%.2f" % u3, 'couts_co': "%.2f" % couts_co}

            contenu_recap_eligibles += r'''%(compte)s & %(couts_a)s & %(couts_b)s & %(couts_c)s &
                %(u1)s & %(u2)s & %(u3)s & %(couts_mo)s & %(couts_co)s \\
                \hline
                ''' % dico_recap_eligibles

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

            for categorie in generaux.codes_d3():
                contenu_recap_poste += Latex.echappe_caracteres(coefprests.obtenir_noms_categories(categorie))
                contenu_recap_poste += r''' & ''' + "%.2f" % sco['sommes_cat_m'][categorie]
                contenu_recap_poste += r''' & ''' + "%.2f" % sco['sommes_cat_r'][categorie]
                contenu_recap_poste += r''' & ''' + "%.2f" % sco['tot_cat'][categorie]
                contenu_recap_poste += r''' \\
                    \hline
                    '''

            contenu_compte += Latex.tableau(contenu_recap_poste, structure_recap_poste, legende_recap_poste)

            if nombre_res > 0:
                contenu_compte += contenu_rsv

            machines_utilisees = {}
            somme_res_compte = {}
            if code_client in reservations.sommes:
                if id_compte in reservations.sommes[code_client]:
                    somme_res_compte = reservations.sommes[code_client][id_compte]
                    for key in somme_res_compte.keys():
                        machines_utilisees[key] = {'machine': machines.donnees[key]['nom']}
            somme_cae_compte = {}
            if code_client in acces.sommes:
                if id_compte in acces.sommes[code_client]:
                    somme_cae_compte = acces.sommes[code_client][id_compte]
                    for key in somme_cae_compte.keys():
                        if key not in machines_utilisees:
                            machines_utilisees[key] = {'machine': machines.donnees[key]['nom']}

            if len(machines_utilisees) > 0:
                structure_stat_machines = r'''{|l|c|c|c|c|c|c|r|}'''
                legende_stat_machines = r'''Statistiques de réservation/utilisation par machine : ''' + intitule_compte
                contenu_stat_machines = r'''
                    \hline
                    Equipement & & Slot rés. & Slot ann. & Taux & Utilisation minimale & Machine & Pénalités [h] \\
                    \hline
                    '''

                for machine_t in sorted(machines_utilisees.items(), key=lambda k_v: k_v[1]['machine']):
                    machine = machine_t[1]
                    id_machine = machine_t[0]

                    taux_hp = machines.donnees[id_machine]['tx_occ_eff_hp']
                    taux_hc = machines.donnees[id_machine]['tx_occ_eff_hc']
                    duree_hp = 0
                    duree_hc = 0
                    res_hp = 0
                    res_hc = 0
                    ann_hp = 0
                    ann_hc = 0
                    if id_machine in somme_cae_compte:
                        duree_hp = somme_cae_compte[id_machine]['duree_hp']
                        duree_hc = somme_cae_compte[id_machine]['duree_hc']
                    if id_machine in somme_res_compte:
                        res_hp = somme_res_compte[id_machine]['res_hp']
                        res_hc = somme_res_compte[id_machine]['res_hc']
                        ann_hp = somme_res_compte[id_machine]['ann_hp']
                        ann_hc = somme_res_compte[id_machine]['ann_hc']

                    min_hp = (res_hp + ann_hp) * taux_hp / 100
                    min_hc = (res_hc + ann_hc) * taux_hc / 100

                    pen_hp = min_hp - duree_hp
                    pen_hc = min_hc - duree_hc

                    dico_stat_machines = {
                        'machine': Latex.echappe_caracteres(machine['machine']),
                        'res_hp': Outils.format_heure(res_hp), 'res_hc': Outils.format_heure(res_hc),
                        'ann_hp': Outils.format_heure(ann_hp), 'ann_hc': Outils.format_heure(ann_hc),
                        'duree_hp': Outils.format_heure(duree_hp), 'duree_hc': Outils.format_heure(duree_hc),
                        'taux_hp': str(int(taux_hp)) + '\%', 'taux_hc': str(int(taux_hc)) + '\%',
                        'min_hp': Outils.format_heure(min_hp), 'min_hc': Outils.format_heure(min_hc),
                        'pen_hp': "%.1f" % round(pen_hp/60, 1), 'pen_hc': "%.1f" % round(pen_hc/60, 1)}

                    if res_hp > 0 or ann_hp or duree_hp:
                        contenu_stat_machines += r'''%(machine)s & HP &  %(res_hp)s & %(ann_hp)s & %(taux_hp)s &
                            %(min_hp)s & %(duree_hp)s & %(pen_hp)s \\
                            \hline
                            ''' % dico_stat_machines
                    if res_hc > 0 or ann_hc or duree_hc:
                        contenu_stat_machines += r'''%(machine)s & HC &  %(res_hc)s & %(ann_hc)s & %(taux_hc)s &
                            %(min_hc)s & %(duree_hc)s & %(pen_hc)s \\
                            \hline
                            ''' % dico_stat_machines

                contenu_compte += Latex.tableau(contenu_stat_machines, structure_stat_machines,
                                                legende_stat_machines)

            contenu_compte += contenu_projet
            # ## compte

        dic_entete = {'code': code_client, 'code_sap': client['code_sap'],
                      'nom': Latex.echappe_caracteres(client['abrev_labo']),
                      'date': edition.mois_txt + " " + str(edition.annee)}
        entete = r'''
            %(code)s - %(code_sap)s - %(nom)s - %(date)s
            ''' % dic_entete

        contenu += entete

        structure_recap_poste_cl = r'''{|l|r|r|r|}'''
        legende_recap_poste_cl = r'''Récapitulatif postes pour client ''' + intitule_client

        dico_recap_poste_cl = {'emom': "%.2f" % scl['em'], 'emor': "%.2f" % scl['er'], 'emo': "%.2f" % scl['e'],
                               'resm': "%.2f" % scl['rm'], 'resr': "%.2f" % scl['rr'], 'res': "%.2f" % scl['r'],
                               'mat': "%.2f" % scl['mat'], 'stmr': "%.2f" % scl['somme_t_mr'], 'mt': "%.2f" % scl['mt'],
                               'mot': "%.2f" % scl['mot']}

        contenu_recap_poste_cl = r'''
            \hline
             & \multicolumn{1}{l|}{Montant} & \multicolumn{1}{l|}{Rabais} & \multicolumn{1}{l|}{Total} \\
            \hline
            Emolument & %(emom)s & %(emor)s & %(emo)s \\
            \hline
            Frais de réservation & %(resm)s & %(resr)s & %(res)s \\
            \hline
            Machine & %(mat)s & \multirow{2}{*}{%(stmr)s} & \multirow{2}{*}{%(mt)s} \\
            \cline{1-2}
            Main d'oeuvre & %(mot)s & & \\
            \hline
            ''' % dico_recap_poste_cl

        for categorie in generaux.codes_d3():
            contenu_recap_poste_cl += Latex.echappe_caracteres(coefprests.obtenir_noms_categories(categorie))
            contenu_recap_poste_cl += r''' & '''
            contenu_recap_poste_cl += "%.2f" % scl['sommes_cat_m'][categorie] + r''' & '''
            contenu_recap_poste_cl += "%.2f" % scl['sommes_cat_r'][categorie] + r''' & '''
            contenu_recap_poste_cl += "%.2f" % scl['tot_cat'][categorie] + r''' \\
                \hline
                '''

        contenu += Latex.tableau(contenu_recap_poste_cl, structure_recap_poste_cl, legende_recap_poste_cl)

        dico_emolument = {'emb':  "%.2f" % client['emol_base_mens'], 'ef':  "%.2f" % client['emol_fixe'],
                          'pente': client['coef'], 'tot_eq_r': "%.2f" % scl['r'],
                          'tot_eq_m': "%.2f" % (scl['mt']-scl['mot']), 'tot_eq': "%.2f" % scl['somme_eq'],
                          'rabais': "%.2f" % scl['er']}

        structure_emolument = r'''{|r|r|l|r|r|r|r|}'''
        legende_emolument = r'''Emolument pour client ''' + intitule_client
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
                id_machine = machine_t[0]
                som_m = scl['res'][id_machine]
                if som_m['pen_hp'] > 0 or som_m['pen_hc'] > 0:

                    dico_frais_client = {
                        'machine': Latex.echappe_caracteres(machine['machine']),
                        'pen_hp': "%.1f" % som_m['pen_hp'], 'pen_hc': "%.1f" % som_m['pen_hc'],
                        'mont_hp': Outils.format_si_nul(som_m['m_hp']), 'mont_hc': Outils.format_si_nul(som_m['m_hc']),
                        'pu_hp': Outils.format_si_nul(reservations.sommes[code_client][id_machine]['pu_hp']),
                        'pu_hc': Outils.format_si_nul(reservations.sommes[code_client][id_machine]['pu_hc'])}

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

            contenu += Latex.tableau(contenu_frais_client, structure_frais_client, legende_frais_client)

        dico_recap_compte = {'procedes': "%.2f" % scl['mt'], 'total': "%.2f" % scl['somme_t']}

        contenu_recap_compte += r'''Total article & & %(procedes)s''' % dico_recap_compte

        for categorie in generaux.codes_d3():
            contenu_recap_compte += r''' & ''' + "%.2f" % scl['tot_cat'][categorie]

        contenu_recap_compte += r'''& %(total)s \\
                \hline
                ''' % dico_recap_compte

        contenu += Latex.tableau(contenu_recap_compte, structure_recap_compte, legende_recap_compte)

        u1 = scl['somme_t_ai']
        u2 = u1 + scl['somme_t_bi']
        u3 = u2 + scl['somme_t_ci']
        tot_co = 0
        for cat, tt in scl['tot_cat'].items():
            tot_co += tt
        dico_recap_eligibles = {'ait': "%.2f" % scl['somme_t_ai'], 'moit': "%.2f" % scl['somme_t_oi'],
                                'bit': "%.2f" % scl['somme_t_bi'], 'cit': "%.2f" % scl['somme_t_ci'], 'u1': "%.2f" % u1,
                                'u2': "%.2f" % u2, 'u3': "%.2f" % u3, 'coit': "%.2f" % tot_co}

        contenu_recap_eligibles += r'''Total article & %(ait)s & %(bit)s & %(cit)s & %(u1)s & %(u2)s &
            %(u3)s & %(moit)s & %(coit)s \\
            \hline
            ''' % dico_recap_eligibles

        contenu += Latex.tableau(contenu_recap_eligibles, structure_recap_eligibles, legende_recap_eligibles)

        structure_recap_cat_cl = r'''{|l|r|r|r|}'''
        legende_recap_cat_cl = r'''Détail par catégorie pour client ''' + intitule_client

        dico_recap_cat_cl = {'mmk1': '0.00', 'mrk1': '0.00', 'mk1': '0.00', 'mmk2': '0.00', 'mrk2': '0.00',
                             'mk2': '0.00', 'mmk3': '0.00', 'mrk3': '0.00', 'mk3': '0.00', 'mmk4': '0.00',
                             'mrk4': '0.00', 'mk4': '0.00', 'mm': "%.2f" % scl['somme_t_mm'],
                             'mr': "%.2f" % scl['somme_t_mr'], 'mt': "%.2f" % scl['mt']}

        if '1' in sca:
            dico_recap_cat_cl['mmk1'] = "%.2f" % sca['1']['somme_k_mm']
            dico_recap_cat_cl['mrk1'] = "%.2f" % sca['1']['somme_k_mr']
            dico_recap_cat_cl['mk1'] = "%.2f" % sca['1']['mk']
        if '2' in sca:
            dico_recap_cat_cl['mmk2'] = "%.2f" % sca['2']['somme_k_mm']
            dico_recap_cat_cl['mrk2'] = "%.2f" % sca['2']['somme_k_mr']
            dico_recap_cat_cl['mk2'] = "%.2f" % sca['2']['mk']
        if '3' in sca:
            dico_recap_cat_cl['mmk3'] = "%.2f" % sca['3']['somme_k_mm']
            dico_recap_cat_cl['mrk3'] = "%.2f" % sca['3']['somme_k_mr']
            dico_recap_cat_cl['mk3'] = "%.2f" % sca['3']['mk']
        if '4' in sca:
            dico_recap_cat_cl['mmk4'] = "%.2f" % sca['4']['somme_k_mm']
            dico_recap_cat_cl['mrk4'] = "%.2f" % sca['4']['somme_k_mr']
            dico_recap_cat_cl['mk4'] = "%.2f" % sca['4']['mk']

        contenu_recap_cat_cl = r'''
            \hline
             & \multicolumn{1}{l|}{Montant} & \multicolumn{1}{l|}{Rabais} & \multicolumn{1}{l|}{Total} \\
            \hline
            Coût procédés (catégorie Utilisateur) & %(mmk1)s & %(mrk1)s & %(mk1)s \\
            \hline
            Coût procédés (catégorie Etudiant en projet Master) & %(mmk2)s & %(mrk2)s & %(mk2)s \\
            \hline
            Coût procédés (catégorie Etudiant en projet Semestre) & %(mmk3)s & %(mrk3)s & %(mk3)s \\
            \hline
            Coût procédés (catégorie Client) & %(mmk4)s & %(mrk4)s & %(mk4)s \\
            \hline
            Total & %(mm)s & %(mr)s & %(mt)s \\
            \hline
            ''' % dico_recap_cat_cl

        contenu += Latex.tableau(contenu_recap_cat_cl, structure_recap_cat_cl, legende_recap_cat_cl)

        contenu += contenu_compte

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
