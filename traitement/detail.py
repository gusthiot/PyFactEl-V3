from outils import Outils
import math


class Detail(object):
    """
    Classe pour la création du détail des coûts
    """

    @staticmethod
    def detail(dossier_destination, edition, sommes, clients, generaux, acces, livraisons, comptes, couts):
        """
        création du détail des coûts

        :param dossier_destination: Une instance de la classe dossier.DossierDestination
        :param edition: paramètres d'édition
        :param sommes: sommes calculées
        :param clients: clients importés
        :param generaux: paramètres généraux
        :param acces: accès importés
        :param livraisons: livraisons importées
        """

        if sommes.calculees == 0:
            info = "Vous devez d'abord faire toutes les sommes avant de pouvoir créer le détail des coûts"
            print(info)
            Outils.affiche_message(info)
            return

        nom = "detail_" + str(edition.annee) + "_" + Outils.mois_string(edition.mois) + "_" + \
              str(edition.version) + ".csv"

        with dossier_destination.writer(nom) as fichier_writer:

            ligne = ["année", "mois", "code client", "code client sap", "abrév. labo", "id compte", "numéro",
                     "intitulé", "id catégorie", "catégorie", "machine U1", "machine U2", "machine U3",
                     "main d'oeuvre MO", "prestation", "montant prestation"]
            fichier_writer.writerow(ligne)

            for code_client in sorted(sommes.sommes_clients.keys()):
                client = clients.donnees[code_client]

                if code_client in sommes.sommes_comptes:
                    sclo = sommes.sommes_comptes[code_client]
                    comptes_utilises = Outils.comptes_in_somme(sclo, comptes)

                    for num_compte, id_compte in sorted(comptes_utilises.items()):
                        compte = comptes.donnees[id_compte]

                        if code_client in acces.sommes and id_compte in acces.sommes[code_client]['categories']:
                            som_cats = acces.sommes[code_client]['categories'][id_compte]
                            for id_cout, som_cat in sorted(som_cats.items()):
                                ligne = [edition.annee, edition.mois, code_client, client['code_sap'],
                                         client['abrev_labo'], id_compte, num_compte, compte['intitule'], id_cout,
                                         couts.donnees[id_cout]['intitule'], Outils.format_2_dec(som_cat['mu1']),
                                         Outils.format_2_dec(som_cat['mu2']), Outils.format_2_dec(som_cat['mu3']),
                                         Outils.format_2_dec(som_cat['mmo']), "", ""]
                                fichier_writer.writerow(ligne)

                        if code_client in livraisons.sommes and id_compte in livraisons.sommes[code_client]:
                            somme = livraisons.sommes[code_client][id_compte]

                            for article in generaux.articles_d3:
                                if article.code_d in somme:
                                    elu1 = article.eligible_U1
                                    elu2 = article.eligible_U2
                                    elu3 = article.eligible_U3
                                    if elu1 == "NON" and elu2 == "NON" and elu3 == "NON":
                                        continue
                                    ligne = [edition.annee, edition.mois, code_client, client['code_sap'],
                                             client['abrev_labo'], id_compte, num_compte, compte['intitule'], "", "",
                                             "", "", "", "", article.code_d,
                                             Outils.format_2_dec(sclo[id_compte]['tot_cat_x'][article.code_d])]
                                    fichier_writer.writerow(ligne)
