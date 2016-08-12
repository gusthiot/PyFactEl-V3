from outils import Outils


class BilanMensuel(object):
    """
    Classe pour la création du bilan mensuel
    """

    @staticmethod
    def bilan(dossier_destination, edition, sommes, clients, generaux, acces, reservations, livraisons,
              comptes):
        """
        création du bilan

        :param dossier_destination: Une instance de la classe dossier.DossierDestination
        :param edition: paramètres d'édition
        :param sommes: sommes calculées
        :param clients: clients importés
        :param generaux: paramètres généraux
        :param acces: accès importés
        :param reservations: réservations importés
        :param livraisons: livraisons importées
        :param comptes: comptes importés
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
                     "nb tot comptes", "nb comptes cat 1", "nb comptes cat 2", "nb comptes cat 3", "nb comptes cat 4",
                     "total M cat 1", "total M cat 2", "total M cat 3", "total M cat 4", "MAt", "MOt", "DSt", "DHt",
                     "Et", "Rt", "Mt"]
            for categorie in generaux.codes_d3():
                ligne.append(categorie + "t")
            ligne.append("total facturé HT")
            fichier_writer.writerow(ligne)

            for code_client in sorted(sommes.sommes_clients.keys()):
                scl = sommes.sommes_clients[code_client]
                sca = sommes.sommes_categories[code_client]
                client = clients.donnees[code_client]
                nature = generaux.nature_client_par_code_n(client['type_labo'])
                reference = nature + str(edition.annee)[2:] + Outils.mois_string(edition.mois) + "." + code_client
                nb_u = len(BilanMensuel.utilisateurs(acces, livraisons, reservations, code_client))
                cptes = BilanMensuel.comptes(acces, livraisons, reservations, code_client)
                cat = {'1': 0, '2': 0, '3': 0, '4': 0}
                nb_c = 0
                for cpte in cptes:
                    nb_c += 1
                    cat[comptes.donnees[cpte]['categorie']] += 1

                mk = {'1': 0, '2': 0, '3': 0, '4': 0}
                for num in mk.keys():
                    if num in sca:
                        mk[num] = sca[num]['mk']

                total = scl['somme_t'] + scl['e']

                ligne = [edition.annee, edition.mois, reference, code_client, client['code_sap'], client['abrev_labo'],
                         client['nom_labo'], 'U', client['type_labo'], scl['em'], "%.2f" % scl['somme_eq'], scl['er'],
                         client['emol_sans_activite'], nb_u, nb_c, cat['1'], cat['2'], cat['3'], cat['4'],
                         "%.2f" % mk['1'], "%.2f" % mk['2'], "%.2f" % mk['3'], "%.2f" % mk['4'], "%.2f" % scl['mat'],
                         scl['mot'], scl['dst'], scl['dht'], scl['e'], scl['r'], "%.2f" % scl['mt']]
                for categorie in generaux.codes_d3():
                    ligne.append(scl['tot_cat'][categorie])
                ligne.append("%.2f" % total)
                fichier_writer.writerow(ligne)

    @staticmethod
    def utilisateurs(acces, livraisons, reservations, code_client):
        """
        retourne la liste de tous les utilisateurs concernés pour les accès, les réservations et les livraisons
        pour un client donné
        :param acces: accès importés
        :param livraisons: livraisons importées
        :param reservations: réservations importées
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
        for res in reservations.donnees:
            if res['code_client'] == code_client:
                if res['id_user'] not in utilisateurs:
                    utilisateurs.append(res['id_user'])
        return utilisateurs

    @staticmethod
    def comptes(acces, livraisons, reservations, code_client):
        """
        retourne la liste de tous les comptes concernés pour les accès, les réservations et les livraisons
        pour un client donné
        :param acces: accès importés
        :param livraisons: livraisons importées
        :param reservations: réservations importées
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
        for res in reservations.donnees:
            if res['code_client'] == code_client:
                if res['id_compte'] not in comptes:
                    comptes.append(res['id_compte'])
        return comptes
