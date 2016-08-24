from importes import Fichier
from outils import Outils
from traitement import Rabais


class Reservation(Fichier):
    """
    Classe pour l'importation des données de Réservations
    """

    cles = ['annee', 'mois', 'code_client', 'abrev_labo', 'id_user', 'nom_user',
            'prenom_user', 'id_machine', 'nom_machine', 'date_debut', 'duree_hp', 'duree_hc', 'si_supprime',
            'duree_ouvree', 'date_reservation', 'date_suppression']
    nom_fichier = "res.csv"
    libelle = "Réservation Equipement"

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.sommes = {}

    def est_coherent(self, clients, machines):
        """
        vérifie que les données du fichier importé sont cohérentes (code client parmi clients,
        id machine parmi machines), et efface les colonnes mois et année
        :param clients: clients importés
        :param machines: machines importées
        :return: 1 s'il y a une erreur, 0 sinon
        """
        if self.verifie_date == 0:
            info = self.libelle + ". vous devez vérifier la date avant de vérifier la cohérence"
            print(info)
            Outils.affiche_message(info)
            return 1

        if self.verifie_coherence == 1:
            print(self.libelle + ": cohérence déjà vérifiée")
            return 0

        msg = ""
        ligne = 1
        donnees_list = []

        for donnee in self.donnees:
            if donnee['code_client'] == "":
                msg += "le code client de la ligne " + str(ligne) + " ne peut être vide\n"
            elif donnee['code_client'] not in clients.obtenir_codes():
                msg += "le code client '" + donnee['code_client'] + "' de la ligne " + str(ligne) + " n'est pas référencé\n"

            if donnee['id_machine'] == "":
                msg += "le machine id de la ligne " + str(ligne) + " ne peut être vide\n"
            elif machines.contient_id(donnee['id_machine']) == 0:
                msg += "le machine id '" + donnee['id_machine'] + "' de la ligne " + str(ligne) \
                       + " n'est pas référencé\n"

            donnee['duree_hp'], info = Outils.est_un_nombre(donnee['duree_hp'], "la durée réservée HP", ligne)
            msg += info
            donnee['duree_hc'], info = Outils.est_un_nombre(donnee['duree_hc'], "la durée réservée HC", ligne)
            msg += info
            donnee['duree_ouvree'], info = Outils.est_un_nombre(donnee['duree_ouvree'], "la durée ouvrée", ligne)
            msg += info

            del donnee['annee']
            del donnee['mois']
            donnees_list.append(donnee)

            ligne += 1

        self.donnees = donnees_list
        self.verifie_coherence = 1

        if msg != "":
            msg = self.libelle + "\n" + msg
            print("msg : " + msg)
            Outils.affiche_message(msg)
            return 1
        return 0

    def calcul_montants(self, machines, coefmachines, clients, verification):
        """
        calcule les montants 'pv' et 'qv' et les ajoute aux données
        :param machines: machines importées et vérifiées
        :param coefmachines: coefficients machines importés et vérifiés
        :param clients: clients importés et vérifiés
        :param verification: pour vérifier si les dates et les cohérences sont correctes
        """
        if verification.a_verifier != 0:
            info = self.libelle + ". vous devez faire les vérifications avant de calculer les montants"
            print(info)
            Outils.affiche_message(info)
            return

        donnees_list = []
        pos = 0
        for donnee in self.donnees:
            code_client = donnee['code_client']
            id_machine = donnee['id_machine']
            id_user = donnee['id_user']
            machine = machines.donnees[id_machine]
            client = clients.donnees[code_client]
            coefmachine = coefmachines.donnees[client['id_classe_tarif'] + machine['categorie']]
            duree_fact_hp, duree_fact_hc = Rabais.rabais_reservation(machine['delai_sans_frais'],
                                                                     donnee['duree_ouvree'],
                                                                     donnee['duree_hp'],
                                                                     donnee['duree_hc'])

            donnee['duree_fact_hp'] = duree_fact_hp
            donnee['duree_fact_hc'] = duree_fact_hc

            if code_client not in self.sommes:
                self.sommes[code_client] = {'client': {}, 'machines': {}}
            scl = self.sommes[code_client]
            if id_machine not in scl['client']:
                scl['client'][id_machine] = {'res_hp': 0, 'ann_hp': 0, 'res_hc': 0, 'ann_hc': 0}

            scl['client'][id_machine]['res_hp'] += duree_fact_hp
            scl['client'][id_machine]['res_hc'] += duree_fact_hc

            if id_machine not in scl['machines']:
                pu_hp = round(coefmachine['coef_r'] * machine['t_h_reservation_hp'], 2)
                pu_hc = round(coefmachine['coef_r'] * machine['t_h_reservation_hc'], 2)
                scl['machines'][id_machine] = {'res_hp': 0, 'res_hc': 0, 'pu_hp': pu_hp, 'pu_hc': pu_hc, 'users': {}}
            scm = scl['machines'][id_machine]

            scm['res_hp'] += duree_fact_hp
            scm['res_hc'] += duree_fact_hc

            if id_user not in scm['users']:
                scm['users'][id_user] = {'nom': donnee['nom_user'], 'prenom': donnee['prenom_user'],
                                         'res_hp': 0, 'res_hc': 0, 'data': []}

            scm['users'][id_user]['res_hp'] += duree_fact_hp
            scm['users'][id_user]['res_hc'] += duree_fact_hc
            scm['users'][id_user]['data'].append(pos)

            donnees_list.append(donnee)
            pos += 1

        self.donnees = donnees_list

    def reservations_pour_compte(self, id_compte, code_client):
        """
        retourne toutes les données réservations pour un compte donné
        :param id_compte: l'id du compte
        :param code_client: le code du client
        :return: toutes les données réservations d'un compte donné
        """
        donnees_list = []
        for donnee in self.donnees:
            if (donnee['id_compte'] == id_compte) and (donnee['code_client'] == code_client):
                donnees_list.append(donnee)
        return donnees_list
