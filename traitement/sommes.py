from outils import Outils
from .rabais import Rabais


class Sommes(object):
    """
    Classe contenant les méthodes pour le calcul des sommes par compte, catégorie et client
    """

    cles_somme_compte = ['somme_j_ai', 'somme_j_bi', 'somme_j_ci', 'somme_j_oi', 'somme_j_mai', 'somme_j_moi',
                         'somme_j_dsi', 'somme_j_dhi', 'somme_j_mm', 'somme_j_mr', 'mj', 'si_facture', 'res']

    cles_somme_client = ['somme_t_ai', 'somme_t_bi', 'somme_t_ci', 'somme_t_oi', 'mat', 'mot', 'dst', 'dht',
                         'somme_t_mm', 'somme_t_mr', 'mt', 'somme_eq', 'somme_t', 'em', 'er0', 'er', 'e', 'res', 'rm',
                         'rr', 'r']

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
        somme['sommes_cat_r'] = {}
        somme['tot_cat'] = {}
        for categorie in self.categories:
            somme['sommes_cat_m'][categorie] = 0
            somme['sommes_cat_r'][categorie] = 0
            somme['tot_cat'][categorie] = 0
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
            somme['sommes_cat_r'][prestation['categorie']] += livraison['rabais_r']
            somme['tot_cat'][prestation['categorie']] += livraison['montant'] - livraison['rabais_r']

            tot = somme['somme_j_mm']
            for categorie in self.categories:
                tot += somme['sommes_cat_m'][categorie]
            if tot > 0:
                somme['si_facture'] = 1

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
                machines_utilisees = []
                somme_res = reservations.sommes[code_client]['client']
                for key in somme_res.keys():
                    machines_utilisees.append(key)

                somme_cae = {}
                if code_client in acces.sommes:
                    if id_compte in acces.sommes[code_client]:
                        somme_cae_compte = acces.sommes[code_client][id_compte]
                        for id_mach in somme_cae_compte:
                            if id_mach not in somme_cae:
                                somme_cae[id_mach] = {'duree_hp': 0, 'duree_hc': 0}
                            somme_cae[id_mach]['duree_hp'] += somme_cae_compte[id_mach]['duree_hp']
                            somme_cae[id_mach]['duree_hc'] += somme_cae_compte[id_mach]['duree_hc']
                for key in somme_cae.keys():
                    if key not in machines_utilisees:
                        machines_utilisees.append(key)

                for mach_u in machines_utilisees:
                    if mach_u in somme_res:
                        mini_hp = (somme_res[mach_u]['res_hp'] + somme_res[mach_u]['ann_hp']) * \
                                  machines.donnees[mach_u]['tx_occ_eff_hp'] / 100
                        mini_hc = (somme_res[mach_u]['res_hc'] + somme_res[mach_u]['ann_hc']) * \
                                  machines.donnees[mach_u]['tx_occ_eff_hc'] / 100
                    else:
                        mini_hp = 0
                        mini_hc = 0
                    if mach_u in somme_cae:
                        pen_hp = mini_hp - somme_cae[mach_u]['duree_hp']
                        pen_hc = mini_hc - somme_cae[mach_u]['duree_hc']
                    else:
                        pen_hp = mini_hp
                        pen_hc = mini_hc
                    if mach_u in somme_res:
                        m_hp = round(pen_hp / 60, 1) * reservations.sommes[code_client]['machines'][mach_u]['pu_hp']
                        m_hc = round(pen_hc / 60, 1) * reservations.sommes[code_client]['machines'][mach_u]['pu_hc']
                    else:
                        m_hp = 0
                        m_hc = 0

                    somme['res'][mach_u] = {'pen_hp': round(pen_hp / 60, 1),
                                            'pen_hc': round(pen_hc / 60, 1),
                                            'm_hp': m_hp, 'm_hc': m_hc}
                    somme['rm'] += m_hp + m_hc

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
