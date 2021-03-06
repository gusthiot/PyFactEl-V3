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
                          prestations, reservations, couts, users):
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
        :param couts: catégories couts importées
        :param users: users importés
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
        verif += couts.verification_date(edition.annee, edition.mois)
        verif += users.verification_date(edition.annee, edition.mois)
        self.a_verifier = 1
        return verif

    def verification_coherence(self, generaux, edition, acces, clients, coefmachines, coefprests, comptes, livraisons,
                               machines, prestations, reservations, couts, users):
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
        :param couts: catégories couts importées
        :param users: users importés
        :return: 0 si ok, sinon le nombre d'échecs à la vérification
        """
        verif = 0
        verif += acces.est_coherent(comptes, machines, users)
        verif += livraisons.est_coherent(comptes, prestations, users)
        verif += couts.est_coherent()
        verif += users.est_coherent()
        verif += machines.est_coherent(coefmachines, couts)
        verif += prestations.est_coherent(generaux, coefprests)
        verif += coefmachines.est_coherent()
        verif += coefprests.est_coherent(generaux)
        verif += clients.est_coherent(coefmachines, coefprests, generaux)
        verif += reservations.est_coherent(clients, machines, users)

        comptes_actifs = Verification.obtenir_comptes_actifs(acces, livraisons)

        verif += comptes.est_coherent(comptes_actifs)

        clients_actifs = Verification.obtenir_clients_actifs(comptes_actifs, comptes)

        if (edition.version > 0) and (len(clients_actifs) > 1):
            Outils.affiche_message("Si version différente de 0, un seul client autorisé")
            sys.exit("Trop de clients pour version > 0")

        self.a_verifier = 0
        if len(clients_actifs) == 1:
            edition.client_unique = clients_actifs[0]
        return verif

    @staticmethod
    def obtenir_comptes_actifs(acces, livraisons):
        """
        retourne la liste des comptes utilisés, pour les accès et les livraisons
        :param acces: accès importés
        :param livraisons: livraisons importées
        :return: comptes utilisés mappés par clients
        """
        comptes_actifs = []
        for id_compte in livraisons.obtenir_comptes():
            if id_compte not in comptes_actifs:
                comptes_actifs.append(id_compte)
        for id_compte in acces.obtenir_comptes():
            if id_compte not in comptes_actifs:
                comptes_actifs.append(id_compte)

        return comptes_actifs

    @staticmethod
    def obtenir_clients_actifs(comptes_actifs, comptes):
        """
        retourne la liste des clients des comptes actifs
        :param comptes_actifs: comptes actifs
        :param comptes: comptes importés
        :return: comptes utilisés mappés par clients
        """
        clients_actifs = []
        for id_compte in comptes_actifs:
            code_client = comptes.donnees[id_compte]['code_client']
            if code_client not in clients_actifs:
                clients_actifs.append(code_client)
        return clients_actifs
