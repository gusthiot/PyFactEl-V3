from importes import Fichier
from outils import Outils


class Livraison(Fichier):
    """
    Classe pour l'importation des données de Livraisons
    """

    cles = ['annee', 'mois', 'id_compte', 'intitule_compte', 'code_client', 'abrev_labo', 'id_user', 'nom_user',
            'prenom_user', 'id_prestation', 'designation', 'date_livraison',
            'quantite', 'unite', 'rabais', 'responsable', 'id_livraison', 'date_commande', 'date_prise', 'remarque']
    nom_fichier = "lvr.csv"
    libelle = "Livraison Prestations"

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.comptes = {}

    def obtenir_comptes(self):
        """
        retourne la liste de tous les comptes clients
        :return: liste des comptes clients présents dans les données livraisons importées
        """
        if self.verifie_coherence == 0:
            info = self.libelle + ". vous devez vérifier la cohérence avant de pouvoir obtenir les comptes"
            print(info)
            Outils.affiche_message(info)
            return []
        return self.comptes

    def est_coherent(self, comptes, prestations):
        """
        vérifie que les données du fichier importé sont cohérentes (id compte parmi comptes,
        id prestation parmi prestations), et efface les colonnes mois et année
        :param comptes: comptes importés
        :param prestations: prestations importées
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
            if donnee['id_compte'] == "":
                msg += "le compte id de la ligne " + str(ligne) + " ne peut être vide\n"
            elif comptes.contient_id(donnee['id_compte']) == 0:
                msg += "le compte id '" + donnee['id_compte'] + "' de la ligne " + str(ligne) + " n'est pas référencé\n"
            elif donnee['code_client'] not in self.comptes:
                self.comptes[donnee['code_client']] = [donnee['id_compte']]
            elif donnee['id_compte'] not in self.comptes[donnee['code_client']]:
                self.comptes[donnee['code_client']].append(donnee['id_compte'])

            if donnee['id_prestation'] == "":
                msg += "le prestation id de la ligne " + str(ligne) + " ne peut être vide\n"
            elif prestations.contient_id(donnee['id_prestation']) == 0:
                msg += "le prestation id '" + donnee['id_prestation'] + "' de la ligne " + str(ligne) +\
                       " n'est pas référencé\n"

            donnee['quantite'], info = Outils.est_un_nombre(donnee['quantite'], "la quantité", ligne)
            msg += info
            donnee['rabais'], info = Outils.est_un_nombre(donnee['rabais'], "le rabais", ligne)
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

    def calcul_montants(self, prestations, coefprests, comptes, clients, verification):
        """
        calcule le 'montant' et le 'rabais_r' et les ajoute aux données
        :param prestations: prestations importées et vérifiées
        :param coefprests: coefficients prestations importés et vérifiés
        :param comptes: comptes importés et vérifiés
        :param clients: clients importés et vérifiés
        :param verification: pour vérifier si les dates et les cohérences sont correctes
        """
        if verification.a_verifier != 0:
            info = self.libelle + ". vous devez faire les vérifications avant de calculer les montants"
            print(info)
            Outils.affiche_message(info)
            return

        donnees_list = []
        for donnee in self.donnees:
            prestation = prestations.donnees[donnee['id_prestation']]
            compte = comptes.donnees[donnee['id_compte']]
            client = clients.donnees[compte['code_client']]
            coefprest = coefprests.donnees[client['id_classe_tarif'] + prestation['categorie']]
            donnee['prix_unit_client'] = round(prestation['prix_unit'] * coefprest['coefficient'], 2)
            donnee['montant'] = round(donnee['quantite'] * donnee['prix_unit_client'], 2)
            donnee['rabais_r'] = round(donnee['rabais'], 2)
            donnees_list.append(donnee)
        self.donnees = donnees_list

    def livraisons_pour_compte_par_categorie(self, id_compte, code_client, prestations):
        """
        retourne les livraisons pour un compte donné, pour une catégorie de prestations donnée
        :param id_compte: l'id du compte
        :param code_client: le code du client
        :param prestations: prestations importées et vérifiées
        :return: les livraisons pour le projet donné, pour une catégorie de prestations donnée
        """

        if prestations.verifie_coherence == 0:
            info = self.libelle + ". vous devez vérifier la cohérence des prestations avant de pouvvoir sélectionner " \
                                  "les livraisons par catégorie"
            print(info)
            Outils.affiche_message(info)
            return {}

        donnees_dico = {}
        for donnee in self.donnees:
            if (donnee['id_compte'] == id_compte) and (donnee['code_client'] == code_client):
                categorie = prestations.donnees[donnee['id_prestation']]['categorie']
                if categorie not in donnees_dico:
                    donnees_dico[categorie] = []
                liste = donnees_dico[categorie]
                liste.append(donnee)
        return donnees_dico
