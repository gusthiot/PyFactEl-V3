from outils import Outils
from .rabais import Rabais


class Sommes(object):
    """
    Classe contenant les méthodes pour le calcul des sommes par compte, catégorie et client
    """

    cles_somme_compte = ['somme_j_ai', 'somme_j_bi', 'somme_j_ci', 'somme_j_oi', 'somme_j_mai', 'somme_j_moi',
                         'somme_j_dsi', 'somme_j_dhi', 'somme_j_mm', 'somme_j_mr', 'somme_j_mb', 'mj', 'si_facture',
                         'res', 'mu1', 'mu2', 'mu3', 'mmo']

    cles_somme_client = ['somme_t_ai', 'somme_t_bi', 'somme_t_ci', 'somme_t_oi', 'mat', 'mot', 'dst', 'dht',
                         'somme_t_mm', 'somme_t_mr', 'somme_t_mb', 'mt', 'somme_eq', 'somme_t', 'em', 'er0', 'er', 'e',
                         'res', 'rm', 'rr', 'r']

    def __init__(self, verification, generaux):
        """
        initialisation des sommes, et vérification si données utilisées correctes
        :param verification: pour vérifier si les dates et les cohérences sont correctes
        :param generaux: paramètres généraux
        """

        self.verification = verification
        self.sommes_comptes = {}
        self.sco = 0
        self.sommes_clients = {}
        self.calculees = 0
        self.categories = generaux.codes_d3()
        self.min_fact_rese = generaux.min_fact_rese

    def calculer_toutes(self, livraisons, reservations, acces, prestations, clients, machines):
        """
        calculer toutes les sommes, par compte et par client
        :param livraisons: livraisons importées et vérifiées
        :param reservations: réservations importées et vérifiées
        :param acces: accès machines importés et vérifiés
        :param prestations: prestations importées et vérifiées
        :param clients: clients importés et vérifiés
        :param machines: machines importées et vérifiées
        """
        self.somme_par_compte(livraisons, acces, prestations)
        self.somme_par_client(clients, reservations, machines, acces)

    def nouveau_somme(self, cles):
        """
        créé un nouveau dictionnaire avec les clés entrées
        :param cles: clés pour le dictionnaire
        :return: dictionnaire indexé par les clés données, avec valeurs à zéro
        """
        somme = {}
        for cle in cles:
            somme[cle] = 0
        somme['sommes_cat_m'] = {}
        somme['sommes_cat_m_x'] = {}
        somme['sommes_cat_r'] = {}
        somme['tot_cat'] = {}
        somme['tot_cat_x'] = {}
        for categorie in self.categories:
            somme['sommes_cat_m'][categorie] = 0
            somme['sommes_cat_m_x'][categorie] = 0
            somme['sommes_cat_r'][categorie] = 0
            somme['tot_cat'][categorie] = 0
            somme['tot_cat_x'][categorie] = 0
        return somme

    def somme_par_compte(self, livraisons, acces, prestations):
        """
        calcule les sommes par comptes sous forme de dictionnaire : client->compte->clés_sommes
        :param livraisons: livraisons importées et vérifiées
        :param acces: accès machines importés et vérifiés
        :param prestations: prestations importées et vérifiées
        """

        if self.verification.a_verifier != 0:
            info = "Sommes :  vous devez faire les vérifications avant de calculer les sommes"
            print(info)
            Outils.affiche_message(info)
            return

        spco = {}
        for acce in acces.donnees:
            id_compte = acce['id_compte']
            code_client = acce['code_client']
            if code_client not in spco:
                spco[code_client] = {}
            spco_cl = spco[code_client]
            if id_compte not in spco_cl:
                spco_cl[id_compte] = self.nouveau_somme(Sommes.cles_somme_compte)
            somme = spco_cl[id_compte]
            somme['somme_j_ai'] += acce['ai']
            somme['somme_j_bi'] += acce['bi']
            somme['somme_j_ci'] += acce['ci']
            somme['somme_j_oi'] += acce['oi']
            somme['somme_j_mai'] += acce['mai']
            somme['somme_j_moi'] += acce['moi']
            somme['somme_j_dsi'] += acce['dsi']
            somme['somme_j_dhi'] += acce['dhi']
            somme['somme_j_mm'] += acce['mm']
            somme['somme_j_mr'] += acce['mr']
            somme['somme_j_mb'] += acce['mb']
            somme['mj'] += acce['m']

        for livraison in livraisons.donnees:
            id_compte = livraison['id_compte']
            code_client = livraison['code_client']
            if code_client not in spco:
                spco[code_client] = {}
            spco_cl = spco[code_client]
            if id_compte not in spco_cl:
                spco_cl[id_compte] = self.nouveau_somme(Sommes.cles_somme_compte)
            somme = spco_cl[id_compte]

            id_prestation = livraison['id_prestation']
            prestation = prestations.donnees[id_prestation]

            somme['sommes_cat_m'][prestation['categorie']] += livraison['montant']
            somme['sommes_cat_m_x'][prestation['categorie']] += livraison['montantx']
            somme['sommes_cat_r'][prestation['categorie']] += livraison['rabais_r']
            somme['tot_cat'][prestation['categorie']] += livraison['montant'] - livraison['rabais_r']
            somme['tot_cat_x'][prestation['categorie']] += livraison['montantx'] - livraison['rabais_r']


        for code_client in spco:
            if code_client in acces.sommes:
                for id_compte in spco[code_client]:
                    somme = spco[code_client][id_compte]

                    tot = somme['somme_j_mm']
                    for categorie in self.categories:
                        tot += somme['sommes_cat_m'][categorie]
                    if tot > 0:
                        somme['si_facture'] = 1

                    ac_som = acces.sommes[code_client]['categories']
                    if id_compte in ac_som:
                        for id_cout, som in ac_som[id_compte].items():
                            somme['mu1'] += som['mu1']
                            somme['mu2'] += som['mu2']
                            somme['mu3'] += som['mu3']
                            somme['mmo'] += som['mmo']

        # print("")
        # print("spco")
        # for code in spco:
        #     if code != "220208":
        #         continue
        #     print(code)
        #     spco_cl = spco[code]
        #     for id in spco_cl:
        #         somme = spco_cl[id]
        #         print("   ", id, somme['somme_j_mai'])

        self.sco = 1
        self.sommes_comptes = spco

    def somme_par_client(self, clients, reservations, machines, acces):
        """
        calcule les sommes par clients sous forme de dictionnaire : client->clés_sommes
        :param clients: clients importés et vérifiés
        :param reservations: réservations importées et vérifiées
        :param machines: machines importées et vérifiées
        :param acces: accès machines importés et vérifiés
        """

        if self.verification.a_verifier != 0:
            info = "Sommes :  vous devez faire les vérifications avant de calculer les sommes"
            print(info)
            Outils.affiche_message(info)
            return

        if self.sco != 0:
            spcl = {}
            for code_client, spco_cl in self.sommes_comptes.items():
                spcl[code_client] = self.nouveau_somme(Sommes.cles_somme_client)
                somme = spcl[code_client]
                somme['res'] = {}
                somme['rm'] = 0
                somme['rr'] = 0
                somme['r'] = 0
                for id_compte, som_co in spco_cl.items():

                    somme['somme_t_ai'] += som_co['somme_j_ai']
                    somme['somme_t_bi'] += som_co['somme_j_bi']
                    somme['somme_t_ci'] += som_co['somme_j_ci']
                    somme['somme_t_oi'] += som_co['somme_j_oi']
                    somme['mat'] += som_co['somme_j_mai']
                    somme['mot'] += som_co['somme_j_moi']
                    somme['dst'] += som_co['somme_j_dsi']
                    somme['dht'] += som_co['somme_j_dhi']
                    somme['somme_t_mm'] += som_co['somme_j_mm']
                    somme['somme_t_mr'] += som_co['somme_j_mr']
                    somme['somme_t_mb'] += som_co['somme_j_mb']
                    somme['mt'] += som_co['mj']

                    for categorie in self.categories:
                        somme['sommes_cat_m'][categorie] += som_co['sommes_cat_m'][categorie]
                        somme['sommes_cat_r'][categorie] += som_co['sommes_cat_r'][categorie]
                        somme['tot_cat'][categorie] += som_co['tot_cat'][categorie]

            # réservations
            for code_client in reservations.sommes:
                if code_client not in spcl:
                    spcl[code_client] = self.nouveau_somme(Sommes.cles_somme_client)
                    spcl[code_client]['res'] = {}
                somme = spcl[code_client]
                somme_res = reservations.sommes[code_client]

                somme_cae = {}
                if code_client in acces.sommes:
                    somme_cae = acces.sommes[code_client]['machines']

                for id_machine in somme_res.keys():
                    re_hp = somme_res[id_machine]['res_hp']
                    re_hc = somme_res[id_machine]['res_hc']
                    pu_hp = somme_res[id_machine]['pu_hp']
                    pu_hc = somme_res[id_machine]['pu_hc']
                    tx_hp = machines.donnees[id_machine]['tx_occ_eff_hp']
                    tx_hc = machines.donnees[id_machine]['tx_occ_eff_hc']
                    ok_hp = False
                    ok_hc = False
                    if re_hp > 0 and pu_hp > 0 and tx_hp > 0:
                        ok_hp = True
                    if re_hc > 0 and pu_hc > 0 and tx_hc > 0:
                        ok_hc = True

                    if ok_hp or ok_hc:
                        somme['res'][id_machine] = {'tot_hp': 0, 'tot_hc': 0, 'users': {}, 'mont_hp': 0, 'mont_hc': 0}

                        users = somme['res'][id_machine]['users']
                        for id_user, s_u in somme_res[id_machine]['users'].items():
                            if id_user not in users:
                                mini_hp = round(s_u['res_hp'] * tx_hp / 100)
                                mini_hc = round(s_u['res_hc'] * tx_hc / 100)
                                users[id_user] = {'ac_hp': 0, 'ac_hc': 0, 're_hp': s_u['res_hp'],
                                                  're_hc': s_u['res_hc'], 'mini_hp': mini_hp, 'mini_hc': mini_hc,
                                                  'tot_hp': 0, 'tot_hc': 0}

                        if id_machine in somme_cae:
                            for id_user, s_u in somme_cae[id_machine]['users'].items():
                                if id_user not in users:
                                    users[id_user] = {'ac_hp': s_u['duree_hp'], 'ac_hc': s_u['duree_hc'], 're_hp': 0,
                                                      're_hc': 0, 'mini_hp': 0, 'mini_hc': 0, 'tot_hp': 0, 'tot_hc': 0}
                                else:
                                    users[id_user]['ac_hp'] = s_u['duree_hp']
                                    users[id_user]['ac_hc'] = s_u['duree_hc']
                        for id_user, s_u in users.items():
                            s_u['tot_hp'] = s_u['mini_hp'] - s_u['ac_hp']
                            somme['res'][id_machine]['tot_hp'] += s_u['tot_hp']
                            s_u['tot_hc'] = s_u['mini_hc'] - s_u['ac_hc']
                            somme['res'][id_machine]['tot_hc'] += s_u['tot_hc']
                        somme['res'][id_machine]['tot_hp'] = max(0, somme['res'][id_machine]['tot_hp'])
                        somme['res'][id_machine]['tot_hc'] = max(0, somme['res'][id_machine]['tot_hc'])

                        somme['res'][id_machine]['mont_hp'] = round(somme['res'][id_machine]['tot_hp'] * pu_hp / 60, 2)
                        somme['res'][id_machine]['mont_hc'] = round(somme['res'][id_machine]['tot_hc'] * pu_hc / 60, 2)
                        somme['rm'] += somme['res'][id_machine]['mont_hp'] + somme['res'][id_machine]['mont_hc']

                somme['rr'] = Rabais.rabais_reservation_petit_montant(somme['rm'], self.min_fact_rese)
                somme['r'] = somme['rm'] - somme['rr']

                client = clients.donnees[code_client]
                somme['somme_eq'], somme['somme_t'], somme['em'], somme['er0'], somme['er'] = \
                    Rabais.rabais_emolument(somme['r'], somme['mt'], somme['mot'], somme['tot_cat'],
                                            client['emol_base_mens'], client['emol_fixe'], client['coef'],
                                            client['emol_sans_activite'])
                somme['e'] = somme['em'] - somme['er']

            # print("")
            # print("spcl")
            # for code in spcl:
            #     if code != "220208":
            #         continue
            #     somme = spcl[code]
            #     print(code, somme['mat'])

            self.calculees = 1
            self.sommes_clients = spcl

        else:
            info = "Vous devez d'abord faire la somme par catégorie, avant la somme par client"
            print(info)
            Outils.affiche_message(info)
