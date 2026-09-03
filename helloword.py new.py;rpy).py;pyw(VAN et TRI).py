
import math


# ==============================
# 1. CALCUL DE LA VAN
# ==============================

def calcul_van(flux, taux):
    van = 0

    for t in range(len(flux)):
        van += flux[t] / (1 + taux) ** t

    return van


# ==============================
# 2. CALCUL DU TRI
# ==============================
# ==============================


def calcul_tri(flux, precision=0.000001):

    # ------------------------------
    # Vérification de la précision
    # ------------------------------

    if precision <= 0:
        raise ValueError(
            "La précision doit être strictement positive."
        )

    if precision >= 1:
        raise ValueError(
            "La précision doit être inférieure à 1."
        )

    # ------------------------------
    # Vérification du type
    # ------------------------------

    if not isinstance(flux, list):
        raise ValueError(
            "Les flux doivent être fournis sous forme de liste."
        )

    if len(flux) < 2:
        raise ValueError(
            "Il faut au moins deux flux pour calculer un TRI."
        )

    # ------------------------------
    # Vérification des valeurs
    # ------------------------------

    if not all(
        isinstance(f, (int, float))
        and math.isfinite(f)
        for f in flux
    ):
        raise ValueError(
            "Tous les flux doivent être des nombres réels finis."
        )

    # ------------------------------
    # Flux négatif et positif
    # ------------------------------

    if not any(f < 0 for f in flux):
        raise ValueError(
            "Aucun flux négatif n'a été détecté."
        )

    if not any(f > 0 for f in flux):
        raise ValueError(
            "Aucun flux positif n'a été détecté."
        )

    # ------------------------------
    # Détection des changements de signe
    # ------------------------------

    changements_signe = 0

    for i in range(len(flux) - 1):

        if flux[i] * flux[i + 1] < 0:
            changements_signe += 1

    if changements_signe > 1:

        print("\n⚠️ ATTENTION :")
        print(
            "Plusieurs changements de signe ont été détectés."
        )
        print(
            "Le projet peut posséder plusieurs TRI."
        )

    # ------------------------------
    # Recherche des bornes
    # ------------------------------

    taux_bas = -0.99
    taux_haut = 0.10

    van_bas = calcul_van(flux, taux_bas)
    van_haut = calcul_van(flux, taux_haut)

    nombre_essais = 0

    while van_bas * van_haut > 0:

        taux_haut *= 2
        nombre_essais += 1

        # Protection contre des taux trop grands
        if taux_haut > 1_000_000:

            raise ValueError(
                "Impossible de trouver un intervalle contenant "
                "un TRI."
            )

        van_haut = calcul_van(
            flux,
            taux_haut
        )

        if nombre_essais > 100:

            raise ValueError(
                "Impossible de déterminer le TRI."
            )

    # ------------------------------
    # Recherche par dichotomie
    # ------------------------------

    nombre_iterations = 0

    while taux_haut - taux_bas > precision:

        taux_milieu = (
            taux_bas + taux_haut
        ) / 2

        van_milieu = calcul_van(
            flux,
            taux_milieu
        )

        if van_milieu > 0:

            taux_bas = taux_milieu

        else:

            taux_haut = taux_milieu

        nombre_iterations += 1

        if nombre_iterations > 1000:

            raise ValueError(
                "Le calcul du TRI n'a pas convergé."
            )

    tri = (
        taux_bas + taux_haut
    ) / 2

    # ------------------------------
    # Vérification finale
    # ------------------------------

    van_tri = calcul_van(
        flux,
        tri
    )

    if not math.isfinite(van_tri):

        raise ValueError(
            "Le TRI calculé n'est pas valide."
        )

    return tri


# ==============================
# 3. ANALYSE DU PROJET
# ==============================

def analyser_projet(flux, taux):

    try:

        # ==============================
        # CALCUL DU TRI
        # ==============================

        tri_precis = calcul_tri(
            flux,
            precision=0.00000001
        )

        # ==============================
        # CALCUL DE LA VAN
        # ==============================

        van = calcul_van(flux, taux)

        # ==============================
        # INVESTISSEMENT INITIAL
        # ==============================

        investissement_initial = abs(flux[0])

        # ==============================
        # INDICE DE RENTABILITÉ
        # ==============================

        valeur_actuelle_flux = van + investissement_initial

        indice_rentabilite = (
            valeur_actuelle_flux / investissement_initial
        )

        # ==============================
        # DÉLAI DE RÉCUPÉRATION
        # ==============================

        cumul = flux[0]
        delai_recuperation = None

        for annee in range(1, len(flux)):

            cumul_precedent = cumul
            cumul += flux[annee]

            if cumul >= 0:

                if flux[annee] != 0:

                    fraction = (
                        -cumul_precedent / flux[annee]
                    )

                    delai_recuperation = (
                        annee - 1 + fraction
                    )

                else:

                    delai_recuperation = annee

                break

        # ==============================
        # AFFICHAGE DES RÉSULTATS
        # ==============================

        print("\n==========================================")
        print("          RÉSULTATS DU PROJET")
        print("==========================================")

        print(
            "Investissement initial :",
            f"{investissement_initial:,.2f}"
            .replace(",", " ")
            .replace(".", ","),
            "FCFA"
        )

        print(
            "Nombre d'années :",
            len(flux) - 1
        )

        print(
            "Taux d'actualisation :",
            f"{taux * 100:.2f}".replace(".", ","),
            "%"
        )

        print(
            "VAN :",
            f"{van:,.2f}"
            .replace(",", " ")
            .replace(".", ","),
            "FCFA"
        )

        print(
            "TRI :",
            f"{tri_precis * 100:.2f}".replace(".", ","),
            "%"
        )

        print(
            "Indice de rentabilité :",
            f"{indice_rentabilite:.2f}".replace(".", ",")
        )

        if delai_recuperation is not None:

            print(
                "Délai de récupération :",
                f"{delai_recuperation:.2f}"
                .replace(".", ","),
                "ans"
            )

        else:

            print(
                "Délai de récupération :",
                "Non récupéré"
            )

        # ==============================
        # ANALYSE VAN
        # ==============================

        print("\n------------------------------------------")
        print("              ANALYSE")
        print("------------------------------------------")

        if van > 0:
            print("VAN : POSITIVE")
        elif van < 0:
            print("VAN : NÉGATIVE")
        else:
            print("VAN : NULLE")

        # ==============================
        # COMPARAISON TRI / TAUX
        # ==============================

        if tri_precis > taux:

            print(
                "TRI : SUPÉRIEUR au taux d'actualisation"
            )

        elif tri_precis < taux:

            print(
                "TRI : INFÉRIEUR au taux d'actualisation"
            )

        else:

            print(
                "TRI : ÉGAL au taux d'actualisation"
            )

        # ==============================
        # DÉCISION
        # ==============================

        print("\n------------------------------------------")
        print("          DÉCISION FINALE")
        print("------------------------------------------")

        if van > 0 and tri_precis > taux:

            print("✅ PROJET RENTABLE ET ACCEPTABLE")

        elif van < 0 and tri_precis < taux:

            print("❌ PROJET NON RENTABLE ET À REJETER")

        else:

            print("⚠️ PROJET À ANALYSER AVEC PRUDENCE")

        # ==============================
        # VÉRIFICATION DU TRI
        # ==============================

        van_verification = calcul_van(
            flux,
            tri_precis
        )

        if abs(van_verification) <= 1:

            print("✅ TRI suffisamment précis.")

        else:

            print("⚠️ Vérification du TRI nécessaire.")

        # ==============================
        # RETOUR DES RÉSULTATS
        # ==============================

        return {
            "investissement": investissement_initial,
            "annees": len(flux) - 1,
            "taux": taux,
            "van": van,
            "tri": tri_precis,
            "indice": indice_rentabilite,
            "recuperation": delai_recuperation
        }

    except ValueError as erreur:

        print("\nErreur :", erreur)

        return None


# ==============================
# 4. DONNÉES DU PROJET
# ==============================

projets = []
numero_projet = 1

print("\n==========================================")
print("       CALCULATEUR FINANCIER")
print("              VAN & TRI")
print("==========================================")
print("Analyse et comparaison de projets")
print("------------------------------------------")


while True:

    print(f"\n              PROJET N° {numero_projet}")
    print("==========================================")

    # ------------------------------
    # Investissement initial
    # ------------------------------

    while True:

        try:

            investissement = float(
                input(
                    "Entrez l'investissement initial : "
                )
            )

            if investissement <= 0:

                print(
                    "Erreur : l'investissement doit être supérieur à 0."
                )

            else:

                break

        except ValueError:

            print(
                "Erreur : veuillez entrer un nombre valide."
            )

    # ------------------------------
    # Nombre d'années
    # ------------------------------

    while True:

        try:

            nombre_annees = int(
                input(
                    "Entrez le nombre d'années du projet : "
                )
            )

            if nombre_annees <= 0:

                print(
                    "Erreur : le nombre d'années doit être supérieur à 0."
                )

            else:

                break

        except ValueError:

            print(
                "Erreur : veuillez entrer un nombre entier valide."
            )

    # ------------------------------
    # Flux annuels
    # ------------------------------

    flux = [-investissement]

    for annee in range(1, nombre_annees + 1):

        while True:

            try:

                flux_annee = float(
                    input(
                        f"Entrez le flux de l'année {annee} : "
                    )
                )

                break

            except ValueError:

                print(
                    "Erreur : veuillez entrer un nombre valide."
                )

        flux.append(flux_annee)

    # ------------------------------
    # Taux d'actualisation
    # ------------------------------

    while True:

        try:

            taux = float(
                input(
                    "Entrez le taux d'actualisation (en %) : "
                )
            )

            if taux <= -100:

                print(
                    "Erreur : le taux doit être supérieur à -100 %."
                )

            else:

                break

        except ValueError:

            print(
                "Erreur : veuillez entrer un nombre valide."
            )

    taux = taux / 100

    # ==============================
    # 5. ANALYSE
    # ==============================

    resultat = analyser_projet(
        flux,
        taux
    )

    if resultat is not None:

        resultat["numero"] = numero_projet

        projets.append(resultat)

    # ==============================
    # MENU
    # ==============================

    print("\n==========================================")
    print("              MENU")
    print("==========================================")
    print("1 - Analyser un autre projet")
    print("2 - Comparer les projets")
    print("3 - Quitter")
    print("==========================================")

    while True:

        choix = input(
            "Votre choix (1, 2 ou 3) : "
        ).strip()

        if choix == "1":

            numero_projet += 1
            break

        elif choix == "2":

            break

        elif choix == "3":

            break

        else:

            print(
                "Erreur : choisissez 1, 2 ou 3."
            )

    # ==============================
    # COMPARAISON DES PROJETS
    # ==============================

    if choix == "2":

        if len(projets) < 2:

            print("\n⚠️ Il faut au moins deux projets pour comparer.")

        else:

            print("\n")
            print("==============================================================")
            print("                 COMPARAISON DES PROJETS")
            print("==============================================================")

            print(
                f"{'Projet':<10}"
                f"{'VAN (FCFA)':>18}"
                f"{'TRI':>12}"
                f"{'Indice':>12}"
                f"{'Récupération':>16}"
            )

            print("-" * 68)

            for projet in projets:

                van_formatee = (
                    f"{projet['van']:,.2f}"
                    .replace(",", " ")
                    .replace(".", ",")
                )

                tri_formate = (
                    f"{projet['tri'] * 100:.2f}"
                    .replace(".", ",")
                    + " %"
                )

                indice_formate = (
                    f"{projet['indice']:.2f}"
                    .replace(".", ",")
                )

                if projet["recuperation"] is not None:

                    recuperation_formatee = (
                        f"{projet['recuperation']:.2f}"
                        .replace(".", ",")
                        + " ans"
                    )

                else:

                    recuperation_formatee = "Non récupéré"

                print(
                    f"{'Projet ' + str(projet['numero']):<10}"
                    f"{van_formatee:>18}"
                    f"{tri_formate:>12}"
                    f"{indice_formate:>12}"
                    f"{recuperation_formatee:>16}"
                )

            # ==============================
            # CLASSEMENT PAR VAN
            # ==============================

            projets_classes = sorted(
                projets,
                key=lambda p: p["van"],
                reverse=True
            )

            print("\n------------------------------------------")
            print("       CLASSEMENT SELON LA VAN")
            print("------------------------------------------")

            for position, projet in enumerate(
                projets_classes,
                start=1
            ):

                print(
                    f"{position}. Projet {projet['numero']} "
                    f"→ VAN = "
                    f"{projet['van']:,.2f}"
                    .replace(",", " ")
                    .replace(".", ",")
                    + " FCFA"
                )

            # ==============================
            # MEILLEURS INDICATEURS
            # ==============================

            meilleur_van = max(
                projets,
                key=lambda p: p["van"]
            )

            meilleur_tri = max(
                projets,
                key=lambda p: p["tri"]
            )

            meilleur_indice = max(
                projets,
                key=lambda p: p["indice"]
            )

            # ==============================
            # RECOMMANDATION
            # ==============================

            print("\n------------------------------------------")
            print("          RECOMMANDATION")
            print("------------------------------------------")

            print(
                f"🏆 Projet recommandé : "
                f"PROJET {meilleur_van['numero']}"
            )

            print(
                "Motif principal : VAN la plus élevée."
            )

            print(
                f"VAN : "
                f"{meilleur_van['van']:,.2f}"
                .replace(",", " ")
                .replace(".", ",")
                + " FCFA"
            )

            print(
                f"TRI : "
                f"{meilleur_van['tri'] * 100:.2f}"
                .replace(".", ",")
                + " %"
            )

            print(
                f"Indice de rentabilité : "
                f"{meilleur_van['indice']:.2f}"
                .replace(".", ",")
            )

            print("\nAutres meilleurs indicateurs :")

            print(
                f"📈 Meilleur TRI : "
                f"Projet {meilleur_tri['numero']}"
            )

            print(
                f"💰 Meilleur indice : "
                f"Projet {meilleur_indice['numero']}"
            )

            print("==============================================================")

        break

    # ==============================
    # QUITTER
    # ==============================

    if choix == "3":

        print("\n==========================================")
        print("       MERCI D'AVOIR UTILISÉ")
        print("       LE CALCULATEUR VAN & TRI")
        print("==========================================")

        break