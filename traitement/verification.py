import sys
from outils import Outils


class Verification(object):
    """
    Classe servant à vérifier les dates et la cohérence de toutes les données importées
    """

    def __init__(self):
        """
        initialisation à 2 séries de vérification (date et cohérence)
        """
        self.a_verifier = 2

    def verification_date(self, edition, acces, clients, coefmachines, coefprests, comptes, livraisons, machines,
                          prestations, reservations):
        """
        vérifie les dates de toutes les données importées
        :param edition: paramètres d'édition
        :param acces: accès importés
        :param clients: clients importés
        :param coefmachines: coefficients machines importés
        :param coefprests: coefficients prestations importés
        :param comptes: comptes importés
        :param livraisons: livraisons importées
        :param machines: machines importées
        :param prestations: prestations importées
        :param reservations: réservations importées
        :return: 0 si ok, sinon le nombre d'échecs à la vérification
        """
        verif = 0
        verif += acces.verification_date(edition.annee, edition.mois)
        verif += clients.verification_date(edition.annee, edition.mois)
        verif += coefmachines.verification_date(edition.annee, edition.mois)
        verif += coefprests.verification_date(edition.annee, edition.mois)
        verif += comptes.verification_date(edition.annee, edition.mois)
        verif += livraisons.verification_date(edition.annee, edition.mois)
        verif += machines.verification_date(edition.annee, edition.mois)
        verif += prestations.verification_date(edition.annee, edition.mois)
        verif += reservations.verification_date(edition.annee, edition.mois)
        self.a_verifier = 1
        return verif

    def verification_cohérence(self, generaux, edition, acces, clients, coefmachines, coefprests, comptes, livraisons,
                               machines, prestations, reservations):
        """
        vérifie la cohérence des données importées
        :param generaux: paramètres généraux
        :param edition: paramètres d'édition
        :param acces: accès importés
        :param clients: clients importés
        :param coefmachines: coefficients machines importés
        :param coefprests: coefficients prestations importés
        :param comptes: comptes importés
        :param livraisons: livraisons importées
        :param machines: machines importées
        :param prestations: prestations importées
        :param reservations: réservations importées
        :return: 0 si ok, sinon le nombre d'échecs à la vérification
        """
        verif = 0
        verif += acces.est_coherent(comptes, machines)
        verif += reservations.est_coherent(comptes, machines)
        verif += livraisons.est_coherent(comptes, prestations)
        verif += machines.est_coherent(coefmachines)
        verif += prestations.est_coherent(generaux, coefprests)
        verif += coefmachines.est_coherent()
        verif += coefprests.est_coherent(generaux)
        verif += clients.est_coherent(coefmachines, coefprests, generaux)

        comptes_actifs, clients_actifs = Verification.obtenir_comptes_clients_actifs(acces, reservations, livraisons)

        if (edition.version != '0') and (len(clients_actifs) > 1):
            Outils.affiche_message("Si version différente de 0, un seul client autorisé")
            sys.exit("Trop de clients pour version > 0")

        verif += comptes.est_coherent(comptes_actifs)
        self.a_verifier = 0
        if len(clients_actifs) == 1:
            edition.client_unique = clients_actifs[0]
        return verif

    @staticmethod
    def obtenir_comptes_clients_actifs(acces, reservations, livraisons):
        """
        retourne la liste des comptes utilisés , par clients, pour les accès, les réservation et les livraisons
        :param acces: accès importés
        :param reservations: réservations importées
        :param livraisons: livraisons importées
        :return: comptes utilisés mappés par clients
        """
        comptes_actifs = []
        clients_actifs = []
        for client, comptes in livraisons.obtenir_comptes().items():
            if client not in clients_actifs:
                clients_actifs.append(client)
            for compte in comptes:
                if compte not in comptes_actifs:
                    comptes_actifs.append(compte)
        for client, comptes in reservations.obtenir_comptes().items():
            if client not in clients_actifs:
                clients_actifs.append(client)
            for compte in comptes:
                if compte not in comptes_actifs:
                    comptes_actifs.append(compte)
        for client, comptes in acces.obtenir_comptes().items():
            if client not in clients_actifs:
                clients_actifs.append(client)
            for compte in comptes:
                if compte not in comptes_actifs:
                    comptes_actifs.append(compte)

        return comptes_actifs, clients_actifs
