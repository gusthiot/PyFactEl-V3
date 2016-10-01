from outils import Outils
import math

class BilanMensuel(object):
    """
    Classe pour la création du bilan mensuel
    """

    @staticmethod
    def bilan(dossier_destination, edition, sommes, clients, generaux, acces, livraisons):
        """
        création du bilan

        :param dossier_destination: Une instance de la classe dossier.DossierDestination
        :param edition: paramètres d'édition
        :param sommes: sommes calculées
        :param clients: clients importés
        :param generaux: paramètres généraux
        :param acces: accès importés
        :param livraisons: livraisons importées
        """

        if sommes.calculees == 0:
            info = "Vous devez d'abord faire toutes les sommes avant de pouvoir créer le bilan mensuel"
            print(info)
            Outils.affiche_message(info)
            return

        nom = "bilan_" + str(edition.annee) + "_" + Outils.mois_string(edition.mois) + "_" + \
              str(edition.version) + ".csv"

        with dossier_destination.writer(nom) as fichier_writer:

            ligne = ["année", "mois", "référence", "code client", "code client sap", "abrév. labo", "nom labo",
                     "type client", "nature client", "Em base", "somme EQ", "rabais Em", "règle", "nb utilisateurs",
                     "nb comptes", "MAt", "MOt", "DSt", "DHt", "Et", "Rt", "Mt"]
            for categorie in generaux.codes_d3():
                ligne.append(categorie + "t")
            ligne += ["total facturé HT", "Bonus St", "Bonus Ht", "Montant Bonus"]
            fichier_writer.writerow(ligne)

            for code_client in sorted(sommes.sommes_clients.keys()):
                scl = sommes.sommes_clients[code_client]
                client = clients.donnees[code_client]
                nature = generaux.nature_client_par_code_n(client['type_labo'])
                reference = nature + str(edition.annee)[2:] + Outils.mois_string(edition.mois) + "." + code_client
                if edition.version != "0":
                    reference += "-" + edition.version
                nb_u = len(BilanMensuel.utilisateurs(acces, livraisons, code_client))
                nb_c = len(BilanMensuel.comptes(acces, livraisons, code_client))

                total = scl['somme_t'] + scl['e']
                bst = client['bs'] * scl['dst']
                bht = client['bh'] * scl['dht']

                ligne = [edition.annee, edition.mois, reference, code_client, client['code_sap'], client['abrev_labo'],
                         client['nom_labo'], 'U', client['type_labo'], scl['em'], Outils.format_2_dec(scl['somme_eq']),
                         scl['er'], client['emol_sans_activite'], nb_u, nb_c, Outils.format_2_dec(scl['mat']),
                         Outils.format_2_dec(scl['mot']), Outils.format_2_dec(scl['dst']),
                         Outils.format_2_dec(scl['dht']), scl['e'], Outils.format_2_dec(scl['r']),
                         Outils.format_2_dec(scl['mt'])]
                for categorie in generaux.codes_d3():
                    ligne.append(Outils.format_2_dec(scl['tot_cat'][categorie]))
                ligne += [Outils.format_2_dec(total), math.ceil(bst), math.ceil(bht),
                          scl['somme_t_mb']]
                fichier_writer.writerow(ligne)

    @staticmethod
    def utilisateurs(acces, livraisons, code_client):
        """
        retourne la liste de tous les utilisateurs concernés pour les accès, les réservations et les livraisons
        pour un client donné
        :param acces: accès importés
        :param livraisons: livraisons importées
        :param code_client: client donné
        :return: liste des utilisateurs
        """
        utilisateurs = []
        for cae in acces.donnees:
            if cae['code_client'] == code_client:
                if cae['id_user'] not in utilisateurs:
                    utilisateurs.append(cae['id_user'])
        for lvr in livraisons.donnees:
            if lvr['code_client'] == code_client:
                if lvr['id_user'] not in utilisateurs:
                    utilisateurs.append(lvr['id_user'])
        return utilisateurs

    @staticmethod
    def comptes(acces, livraisons, code_client):
        """
        retourne la liste de tous les comptes concernés pour les accès, les réservations et les livraisons
        pour un client donné
        :param acces: accès importés
        :param livraisons: livraisons importées
        :param code_client: client donné
        :return: liste des comptes
        """
        comptes = []
        for cae in acces.donnees:
            if cae['code_client'] == code_client:
                if cae['id_compte'] not in comptes:
                    comptes.append(cae['id_compte'])
        for lvr in livraisons.donnees:
            if lvr['code_client'] == code_client:
                if lvr['id_compte'] not in comptes:
                    comptes.append(lvr['id_compte'])
        return comptes
