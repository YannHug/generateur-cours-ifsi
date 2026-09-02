import streamlit as st
import requests
from bs4 import BeautifulSoup
from google import genai
from google.genai import errors as genai_errors
import os
import time
import urllib.parse
import io
import re
import subprocess
import shutil
from docx import Document
from docx.shared import Pt


# ============================================================
# CONFIGURATION
# ============================================================
#
# Le modèle Gemini n'est plus fixé ici : il se choisit dans
# l'interface (menu déroulant "🤖 Modèle Gemini").


# ============================================================
# PAGE STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Générateur de Cours IFSI",
    page_icon="📚",
    layout="wide"
)


# ============================================================
# FONCTION : AUTO-SCROLL PENDANT LE TRAITEMENT
# ============================================================
#
# Streamlit ne fait pas défiler la page automatiquement quand
# du nouveau contenu apparaît. On observe les changements du
# DOM et on fait défiler vers le bas à chaque ajout, le temps
# du traitement.
# ============================================================

def activer_autoscroll():

    st.iframe(
        """
        <script>
        function trouverConteneurDefilant(doc) {
            const selecteurs = [
                '[data-testid="stAppViewContainer"]',
                '[data-testid="stMain"]',
                'section.main',
                'div.main'
            ];

            for (const sel of selecteurs) {
                const el = doc.querySelector(sel);
                if (el && el.scrollHeight > el.clientHeight) {
                    return el;
                }
            }

            return null;
        }

        function defilerVersLeBas() {
            const doc = window.parent.document;

            const conteneur = trouverConteneurDefilant(doc);

            if (conteneur) {
                conteneur.scrollTop = conteneur.scrollHeight;
            }

            // Repli : on fait aussi défiler la fenêtre entière,
            // au cas où le vrai conteneur défilant serait
            // ailleurs (structure Streamlit qui peut varier).
            window.parent.scrollTo(
                0,
                doc.body.scrollHeight
            );
        }

        const observateur = new MutationObserver(
            defilerVersLeBas
        );

        observateur.observe(
            window.parent.document.body,
            { childList: true, subtree: true }
        );

        defilerVersLeBas();
        </script>
        """,
        height=1
    )


# ============================================================
# FONCTION : TÉLÉCHARGER LE MP3
# ============================================================

# ============================================================
# FONCTION : TROUVER LE VRAI DOMAINE (og:url)
# ============================================================

def obtenir_base_canonique(soup, url_page):

    # Certaines pages (ex : mediacenter.univ-lyon1.fr) embarquent
    # le lecteur Nudgis/MyVideo mais ne sont pas elles-mêmes le
    # serveur qui sert les liens /api/v2/... ou /downloads/...
    # La balise <meta property="og:url"> contient l'URL canonique
    # réelle (ex : myvideo.univ-lyon1.fr) — on l'utilise comme
    # base pour résoudre les liens relatifs si elle existe.

    meta = soup.find(
        "meta",
        attrs={"property": "og:url"}
    )

    if meta and meta.get("content"):

        parsed = urllib.parse.urlparse(
            meta["content"]
        )

        if parsed.scheme and parsed.netloc:

            base = f"{parsed.scheme}://{parsed.netloc}/"

            return base

    return url_page


# ============================================================
# FONCTION : DÉTECTER / EXTRAIRE LES LIENS VIDÉO D'UNE PAGE
# ============================================================
#
# Permet à l'utilisateur de coller l'URL d'une page de cours
# (Moodle "Livre", chapitre, etc.) plutôt que de devoir
# récupérer chaque lien vidéo à la main. On distingue :
# - un lien qui pointe déjà directement vers une vidéo
#   (mediacenter/myvideo) → traité tel quel ;
# - toute autre page → on la scanne pour en extraire tous
#   les liens vidéo qu'elle contient.
# ============================================================

DOMAINES_VIDEO = [
    "mediacenter.univ-lyon1.fr",
    "myvideo.univ-lyon1.fr",
]


def nettoyer_nom_fichier(titre):

    if not titre:

        return "Cours_IFSI"

    nom = re.sub(r'[\\/:*?"<>|]', "", titre)

    nom = re.sub(r"\s+", " ", nom).strip()

    nom = nom.replace(" - ", " - ")

    if len(nom) > 120:

        nom = nom[:120].rstrip()

    return nom or "Cours_IFSI"


def est_lien_video_direct(url):

    try:

        analyse = urllib.parse.urlparse(url)

    except Exception:

        return False

    if not any(
        domaine in analyse.netloc
        for domaine in DOMAINES_VIDEO
    ):

        return False

    if "video=" in analyse.query:

        return True

    if "/permalink/" in analyse.path:

        return True

    return False


def extraire_titre_page(soup):

    h1 = soup.find("h1")

    titre_h1 = (
        h1.get_text(strip=True)
        if h1
        else None
    )


    conteneur_chapitre = soup.find(
        "div",
        id="mod_book-chapter"
    )

    titre_h2 = None

    if conteneur_chapitre:

        h2 = conteneur_chapitre.find("h2")

        if h2:

            titre_h2 = h2.get_text(strip=True)


    parties = [
        t for t in (titre_h1, titre_h2)
        if t
    ]

    if not parties:

        return None

    return " - ".join(parties)


def extraire_liens_videos_page(url_page, session):

    try:

        response = session.get(
            url_page,
            timeout=30
        )

        response.raise_for_status()

    except Exception as e:

        st.error(
            f"❌ Impossible d'accéder à la page : {e}"
        )

        return [], None


    soup = BeautifulSoup(
        response.text,
        "html.parser"
    )


    liens_trouves = []
    deja_vus = set()

    for lien in soup.find_all(
        "a",
        href=True
    ):

        href_absolu = urllib.parse.urljoin(
            url_page,
            lien["href"]
        )

        if (
            est_lien_video_direct(href_absolu)
            and href_absolu not in deja_vus
        ):

            deja_vus.add(href_absolu)

            liens_trouves.append(href_absolu)


    titre = extraire_titre_page(soup)

    return liens_trouves, titre


def developper_urls(urls, session):

    urls_finales = []
    titre_cours = None

    for url in urls:

        if est_lien_video_direct(url):

            urls_finales.append(url)

            continue


        st.write(
            f"🔎 Page de cours détectée, recherche des "
            f"vidéos sur : {url}"
        )

        liens, titre = extraire_liens_videos_page(
            url,
            session
        )

        if titre_cours is None and titre:

            titre_cours = titre


        if liens:

            st.write(
                f"✅ {len(liens)} vidéo(s) trouvée(s) sur "
                f"cette page."
            )

            urls_finales.extend(liens)

        else:

            st.warning(
                f"⚠️ Aucun lien vidéo trouvé sur : {url}"
            )


    return urls_finales, titre_cours


def telecharger_mp3(url_page, index, session):

    try:

        response = session.get(
            url_page,
            timeout=30
        )

        response.raise_for_status()

    except Exception as e:

        st.error(
            f"❌ Impossible d'accéder à la page : {e}"
        )

        return None


    soup = BeautifulSoup(
        response.text,
        "html.parser"
    )


    # Recherche du MP3
    lien_mp3 = None

    for lien in soup.find_all(
        "a",
        href=True
    ):

        href = lien["href"]

        if ".mp3" in href.lower():

            lien_mp3 = href

            break


    if not lien_mp3:

        return None


    # Transforme une URL relative en URL absolue, en se basant
    # sur le vrai domaine (og:url) plutôt que sur l'URL collée
    # par l'utilisateur.
    base = obtenir_base_canonique(soup, url_page)

    url_mp3 = urllib.parse.urljoin(
        base,
        lien_mp3
    )


    st.write(
        "🔗 Fichier MP3 trouvé"
    )


    try:

        audio_response = session.get(
            url_mp3,
            timeout=120
        )

        audio_response.raise_for_status()

    except Exception as e:

        st.error(
            f"❌ Erreur pendant le téléchargement du MP3 : {e}"
        )

        return None


    if not audio_response.content:

        st.error(
            "❌ Le fichier MP3 est vide."
        )

        return None


    # ------------------------------------------------------
    # VÉRIFICATION : est-ce vraiment un fichier audio ?
    # ------------------------------------------------------
    # Si la session n'est pas authentifiée, le serveur peut
    # répondre 200 OK mais avec une page HTML d'erreur ou de
    # connexion à la place du MP3. On vérifie donc le
    # Content-Type et la taille avant d'accepter le fichier.

    content_type = audio_response.headers.get(
        "Content-Type", ""
    ).lower()

    taille_octets = len(audio_response.content)

    if "audio" not in content_type or taille_octets < 100_000:

        return None


    nom_fichier = (
        f"cours_ifsi_{index}.mp3"
    )


    try:

        with open(
            nom_fichier,
            "wb"
        ) as fichier:

            fichier.write(
                audio_response.content
            )

    except Exception as e:

        st.error(
            f"❌ Impossible de sauvegarder le MP3 : {e}"
        )

        return None


    taille_mo = taille_octets / 1024 / 1024

    st.write(
        f"📦 Taille : {taille_mo:.2f} Mo"
    )


    return nom_fichier


# ============================================================
# FONCTION : EXTRAIRE L'AUDIO D'UN FLUX HLS AVEC FFMPEG
# ============================================================

def extraire_audio_flux(url_flux, index, session):

    if shutil.which("ffmpeg") is None:

        st.error(
            """
❌ ffmpeg n'est pas installé sur cette machine.

Sur macOS :

brew install ffmpeg
"""
        )

        return None


    nom_fichier = (
        f"cours_ifsi_{index}.mp3"
    )


    # On ne transmet des en-têtes à ffmpeg QUE si on a un
    # vrai cookie de session. Un en-tête "Cookie:" vide suffit
    # à faire échouer la requête sur certains serveurs — d'où
    # l'importance de ne pas l'envoyer quand il n'y a rien.
    cookie_header = "; ".join(
        f"{c.name}={c.value}"
        for c in session.cookies
    )

    commande = [
        "ffmpeg",
        "-y",
        "-i", url_flux,
        "-vn",
        "-acodec", "libmp3lame",
        "-ar", "44100",
        "-b:a", "128k",
        nom_fichier
    ]

    if cookie_header.strip():

        user_agent = session.headers.get(
            "User-Agent",
            ""
        )

        entetes_ffmpeg = (
            f"Cookie: {cookie_header}\r\n"
            f"User-Agent: {user_agent}\r\n"
        )

        # On insère -headers juste après "-y", avant -i.
        commande.insert(1, entetes_ffmpeg)
        commande.insert(1, "-headers")


    st.write(
        "🎬 Extraction de l'audio depuis le flux "
        "(ffmpeg)..."
    )


    try:

        resultat = subprocess.run(
            commande,
            capture_output=True,
            text=True,
            timeout=900
        )

    except subprocess.TimeoutExpired:

        st.error(
            "⏱️ ffmpeg a mis trop de temps à extraire l'audio."
        )

        return None

    except Exception as e:

        st.error(
            f"❌ Erreur lors de l'exécution de ffmpeg : {e}"
        )

        return None


    if resultat.returncode != 0:

        st.error(
            "❌ ffmpeg n'a pas réussi à extraire l'audio "
            "du flux :"
        )

        st.code(
            resultat.stderr[-3000:]
        )

        return None


    if (
        not os.path.exists(nom_fichier)
        or os.path.getsize(nom_fichier) < 100_000
    ):

        st.error(
            "❌ Le fichier audio extrait est invalide ou "
            "trop petit."
        )

        return None


    taille_mo = (
        os.path.getsize(nom_fichier)
        / 1024
        / 1024
    )

    st.write(
        f"📦 Taille (via ffmpeg) : {taille_mo:.2f} Mo"
    )


    return nom_fichier


# ============================================================
# FONCTION : TROUVER L'OID DE LA VIDÉO
# ============================================================

def obtenir_oid(html_texte):

    correspondance = re.search(
        r'mediaOID\s*:\s*["\']([^"\']+)["\']',
        html_texte
    )

    if correspondance:

        return correspondance.group(1)

    return None


# ============================================================
# FONCTION : RÉCUPÉRER L'AUDIO VIA L'API "modes"
# ============================================================
#
# Le lien HLS statique présent dans le panneau de partage est
# parfois un point d'entrée générique qui ne fonctionne pas.
# Le lecteur JS, lui, appelle en réalité l'API /api/v2/medias/
# modes/ qui renvoie un JSON avec les vraies URLs signées
# (token "st"/"e") pour chaque qualité disponible — que ce
# soit un MP4 progressif ou un flux HLS. C'est cette méthode
# que l'on reproduit ici, car elle est beaucoup plus fiable.
# ============================================================

def recuperer_audio_via_modes(url_page, index, session):

    try:

        response = session.get(
            url_page,
            timeout=30
        )

        response.raise_for_status()

    except Exception as e:

        st.error(
            f"❌ Impossible d'accéder à la page : {e}"
        )

        return None


    soup = BeautifulSoup(
        response.text,
        "html.parser"
    )

    base = obtenir_base_canonique(
        soup,
        url_page
    )

    oid = obtenir_oid(
        response.text
    )

    if not oid:

        st.error(
            "❌ Impossible de trouver l'identifiant (oid) "
            "de la vidéo sur cette page."
        )

        return None


    url_modes = urllib.parse.urljoin(
        base,
        f"/api/v2/medias/modes/"
        f"?oid={oid}"
        f"&html5=webm_ogg_ogv_oga_mp4_m4a_mp3_m3u8"
        f"&yt=yt&embed=embed&maxheight=1080"
    )


    try:

        reponse_modes = session.get(
            url_modes,
            timeout=30
        )

        reponse_modes.raise_for_status()

        data = reponse_modes.json()

    except Exception as e:

        st.error(
            f"❌ Impossible de récupérer les informations "
            f"vidéo (API modes) : {e}"
        )

        return None


    noms_qualites = data.get("names", [])

    if not noms_qualites:

        st.error(
            "❌ Aucune qualité vidéo disponible pour "
            "cette vidéo."
        )

        return None


    # On choisit la qualité avec le débit le plus faible :
    # on n'a besoin que de l'audio, pas de la meilleure image.
    meilleure_qualite = min(
        noms_qualites,
        key=lambda q: data.get(q, {})
                          .get("resource", {})
                          .get("bitrate", float("inf"))
    )

    ressource = data.get(
        meilleure_qualite, {}
    ).get("resource")

    if not ressource or not ressource.get("url"):

        st.error(
            "❌ Aucune URL de média trouvée dans la "
            "réponse de l'API modes."
        )

        return None


    st.write(
        f"🔗 Média trouvé (qualité {meilleure_qualite})"
    )


    return extraire_audio_flux(
        ressource["url"],
        index,
        session
    )


# ============================================================
# FONCTION : RÉCUPÉRER L'AUDIO (MP3 direct, sinon API modes)
# ============================================================

def recuperer_audio(url_page, index, session):

    st.write(
        "⬇️ Recherche du fichier audio..."
    )

    fichier_local = telecharger_mp3(
        url_page,
        index,
        session
    )

    if fichier_local:

        return fichier_local


    st.write(
        "↪️ Pas de MP3 direct exploitable — tentative via "
        "l'API modes (MP4/HLS)."
    )

    return recuperer_audio_via_modes(
        url_page,
        index,
        session
    )


# ============================================================
# FONCTION : ACCÉLÉRER L'AUDIO (réduit le coût Gemini,
# facturé au nombre de secondes d'audio)
# ============================================================

FACTEUR_ACCELERATION = 1.3
ACCELERATION_ACTIVEE = False


def accelerer_audio(fichier_entree, index, facteur=FACTEUR_ACCELERATION):

    if shutil.which("ffmpeg") is None:

        st.write(
            "⚠️ ffmpeg indisponible — envoi de l'audio à "
            "vitesse normale."
        )

        return fichier_entree


    fichier_sortie = f"cours_ifsi_{index}_x{facteur}.mp3"

    commande = [
        "ffmpeg",
        "-y",
        "-i", fichier_entree,
        "-filter:a", f"atempo={facteur}",
        "-acodec", "libmp3lame",
        "-ar", "44100",
        "-b:a", "128k",
        fichier_sortie
    ]

    try:

        resultat = subprocess.run(
            commande,
            capture_output=True,
            text=True,
            timeout=300
        )

    except Exception as e:

        st.write(
            f"⚠️ Accélération audio impossible ({e}) — "
            f"envoi à vitesse normale."
        )

        return fichier_entree


    if (
        resultat.returncode != 0
        or not os.path.exists(fichier_sortie)
        or os.path.getsize(fichier_sortie) < 10_000
    ):

        st.write(
            "⚠️ Accélération audio échouée — envoi à "
            "vitesse normale."
        )

        return fichier_entree


    try:

        os.remove(fichier_entree)

    except Exception:

        pass


    st.write(
        f"⏩ Audio accéléré ×{facteur} (réduit le coût Gemini)"
    )


    return fichier_sortie


# ============================================================
# FONCTION : ATTENDRE LE TRAITEMENT GEMINI
# ============================================================

def attendre_fichier(client, fichier):

    maximum = 300

    temps = 0

    while temps < maximum:

        try:

            info = client.files.get(
                name=fichier.name
            )

        except Exception as e:

            st.error(
                f"❌ Impossible de vérifier le fichier : {e}"
            )

            return None


        if hasattr(info.state, "name"):

            statut = info.state.name

        else:

            statut = str(info.state)


        st.write(
            f"État Gemini : `{statut}`"
        )


        if statut == "ACTIVE":

            # ⚠️ CORRECTIF : on retourne l'objet ORIGINAL renvoyé par
            # client.files.upload() (celui reçu en paramètre : `fichier`),
            # et non "info" (celui renvoyé par client.files.get()).
            # Passer "info" à generate_content() déclenchait un
            # 400 INVALID_ARGUMENT côté API Gemini.
            return fichier


        if statut in [
            "FAILED",
            "ERROR"
        ]:

            st.error(
                f"❌ Gemini n'a pas réussi à traiter "
                f"{fichier.name}"
            )

            return None


        time.sleep(3)

        temps += 3


    st.error(
        "⏱️ Gemini a mis trop de temps à traiter le fichier."
    )

    return None


# ============================================================
# FONCTION : ANALYSER UN AUDIO
# ============================================================

# ============================================================
# FONCTION : APPELER GEMINI AVEC REPRISE AUTOMATIQUE
# ============================================================

def appeler_gemini_avec_reprise(client, model, contents, tentatives=4):

    delais = [10, 30, 60, 120]

    for essai in range(tentatives):

        try:

            return client.models.generate_content(
                model=model,
                contents=contents
            )

        except genai_errors.ServerError as e:

            if essai == tentatives - 1:

                raise

            delai = delais[
                min(essai, len(delais) - 1)
            ]

            st.warning(
                f"⚠️ Gemini est momentanément surchargé "
                f"(503) — nouvelle tentative dans {delai}s "
                f"({essai + 1}/{tentatives})..."
            )

            time.sleep(delai)


def uploader_audio_gemini(
    client,
    fichier_local,
    numero
):

    st.write(
        f"⬆️ Envoi de l'audio {numero} à Gemini..."
    )


    try:

        fichier = client.files.upload(
            file=fichier_local
        )

    except Exception as e:

        st.error(
            f"❌ Erreur d'upload : {e}"
        )

        return None


    st.write(
        f"📤 Fichier envoyé : `{fichier.name}`"
    )


    # Attendre que Gemini ait terminé de traiter le fichier
    fichier_pret = attendre_fichier(
        client,
        fichier
    )


    if fichier_pret is None:

        try:

            client.files.delete(
                name=fichier.name
            )

        except Exception:

            pass

        return None


    st.success(
        f"✅ Audio {numero} envoyé et prêt."
    )


    return fichier_pret


# ============================================================
# FONCTION : CRÉER LA FICHE FINALE (un seul appel Gemini,
# avec tous les fichiers audio directement en entrée)
# ============================================================
#
# Fusionne ce qui était avant deux étapes (transcription par
# audio, puis synthèse à partir du texte) en un seul appel :
# Gemini écoute directement tous les audios et rédige la
# fiche finale en une passe. Ça évite de payer deux fois le
# même contenu (une fois en sortie de transcription, une fois
# en entrée de la synthèse).
#
# Compromis assumé : si cet appel échoue, tout est à refaire
# (pas de résultat partiel par audio comme avant).
# ============================================================

def creer_fiche_finale(
    client,
    fichiers_geminis,
    model_name
):

    st.write(
        "### 🧠 Analyse des audios et création de la "
        "fiche de révision"
    )


    prompt = """
Tu es un formateur expert en Institut de Formation
en Soins Infirmiers (IFSI).

Tu vas écouter un ou plusieurs enregistrements audio
appartenant au même cours (éventuellement en plusieurs
parties). Écoute-les directement — ne produis PAS de
retranscription intermédiaire, rédige directement la fiche
de révision finale.

À partir de l'écoute de ces audios, crée une fiche de
révision complète, claire et pédagogique destinée à des
étudiants infirmiers.

==============================
STRUCTURE OBLIGATOIRE
==============================

# 1. L'essentiel de la physiopathologie

Explique simplement :

- les mécanismes ;
- les causes ;
- les conséquences ;
- les notions essentielles.

Définis les termes médicaux importants.


# 2. Les signes cliniques d'alerte

Présente :

- les signes cliniques ;
- les signes de gravité ;
- les complications ;
- les situations nécessitant une alerte.


# 3. La surveillance infirmière et le rôle propre

Présente clairement :

- les paramètres à surveiller ;
- les observations cliniques ;
- la surveillance de l'évolution ;
- les risques ;
- les transmissions ;
- le rôle propre infirmier.


# 4. Les grands principes de traitement

Présente :

- les traitements évoqués dans le cours ;
- leur objectif ;
- les principales précautions ;
- la surveillance infirmière associée.

Ne donne pas de posologie qui n'apparaît pas dans les audios.


# 5. Points clés à retenir

Termine par une section :

"À retenir pour l'examen"

avec les éléments indispensables à mémoriser.


==============================
RÈGLES
==============================

- Reste fidèle au contenu des audios, n'invente aucune
  information.
- Si plusieurs audios sont fournis, fusionne les
  informations et supprime les répétitions.
- Ne perds aucune information importante.
- Si deux passages semblent contradictoires,
  signale-le plutôt que d'inventer une réponse.
- Si un passage est incompréhensible, indique-le brièvement.
- Utilise des tableaux quand ils facilitent
  la compréhension.
- Utilise des listes à puces.
- Utilise un vocabulaire adapté à des étudiants IFSI.
- La fiche doit être suffisamment détaillée pour réviser
  efficacement.
"""


    contenu_requete = [prompt] + list(fichiers_geminis)


    try:

        reponse = appeler_gemini_avec_reprise(
            client,
            model_name,
            contenu_requete
        )

    except Exception as e:

        st.error(
            "❌ Erreur Gemini lors de la création "
            "de la fiche finale."
        )

        st.exception(e)

        return None


    try:

        return reponse.text

    except Exception:

        return None


# ============================================================
# FONCTION : WORD (conversion Markdown → docx)
# ============================================================

def nettoyer_latex(texte):

    remplacements = {
        r"\rightarrow": "→",
        r"\leftarrow": "←",
        r"\times": "×",
        r"\geq": "≥",
        r"\leq": "≤",
        r"\approx": "≈",
        r"\pm": "±",
        r"\alpha": "α",
        r"\beta": "β",
        r"\gamma": "γ",
    }

    for cle, valeur in remplacements.items():

        texte = texte.replace(cle, valeur)

    # Retire les délimiteurs $...$ ou $$...$$ sans supprimer
    # le contenu (on n'a pas de vrai rendu mathématique, mais
    # on évite au moins les "$" qui polluent le texte).
    texte = re.sub(r"\${1,2}", "", texte)

    return texte


def ajouter_texte_formate(paragraphe, texte):

    # Découpe le texte en segments gras (**...**) / italique
    # (*...* ou _..._) / normal, et ajoute chaque segment
    # comme un "run" avec la bonne mise en forme.

    texte = nettoyer_latex(texte)

    motif = re.compile(
        r"(\*\*.+?\*\*|\*.+?\*|__.+?__|_.+?_)"
    )

    segments = motif.split(texte)

    for segment in segments:

        if not segment:

            continue

        if segment.startswith("**") and segment.endswith("**"):

            run = paragraphe.add_run(segment[2:-2])

            run.bold = True

        elif segment.startswith("__") and segment.endswith("__"):

            run = paragraphe.add_run(segment[2:-2])

            run.bold = True

        elif segment.startswith("*") and segment.endswith("*"):

            run = paragraphe.add_run(segment[1:-1])

            run.italic = True

        elif segment.startswith("_") and segment.endswith("_"):

            run = paragraphe.add_run(segment[1:-1])

            run.italic = True

        else:

            paragraphe.add_run(segment)


def est_ligne_tableau(ligne):

    return (
        ligne.strip().startswith("|")
        and ligne.strip().endswith("|")
    )


def est_ligne_separateur_tableau(ligne):

    contenu = ligne.strip().strip("|")

    cellules = contenu.split("|")

    return all(
        re.fullmatch(r"\s*:?-{2,}:?\s*", cellule)
        for cellule in cellules
    ) and len(cellules) > 0


def parser_ligne_tableau(ligne):

    contenu = ligne.strip().strip("|")

    return [
        cellule.strip()
        for cellule in contenu.split("|")
    ]


def ajouter_tableau(document, lignes_tableau):

    entetes = parser_ligne_tableau(lignes_tableau[0])

    lignes_donnees = [
        parser_ligne_tableau(ligne)
        for ligne in lignes_tableau[2:]
    ]

    table = document.add_table(
        rows=1,
        cols=len(entetes)
    )

    table.style = "Light Grid Accent 1"

    cellules_entete = table.rows[0].cells

    for i, texte_entete in enumerate(entetes):

        cellules_entete[i].paragraphs[0].text = ""

        ajouter_texte_formate(
            cellules_entete[i].paragraphs[0],
            texte_entete
        )

        for run in cellules_entete[i].paragraphs[0].runs:

            run.bold = True


    for ligne_donnees in lignes_donnees:

        cellules = table.add_row().cells

        for i, valeur in enumerate(ligne_donnees):

            if i >= len(cellules):

                break

            cellules[i].paragraphs[0].text = ""

            ajouter_texte_formate(
                cellules[i].paragraphs[0],
                valeur
            )


def creer_word(texte):

    document = Document()

    document.add_heading(
        "Fiche de Révision IFSI",
        0
    )

    lignes = texte.split("\n")

    dans_bloc_code = False
    tampon_code = []

    dans_tableau = False
    tampon_tableau = []

    i = 0

    while i < len(lignes):

        ligne_brute = lignes[i]
        ligne = ligne_brute.rstrip()
        ligne_strip = ligne.strip()


        # ------------------------------------------------
        # BLOCS DE CODE ```...```
        # ------------------------------------------------

        if ligne_strip.startswith("```"):

            if dans_bloc_code:

                # Fin du bloc : on écrit le contenu en
                # police à chasse fixe.
                if tampon_code:

                    p = document.add_paragraph()

                    run = p.add_run(
                        "\n".join(tampon_code)
                    )

                    run.font.name = "Courier New"
                    run.font.size = Pt(9)

                tampon_code = []
                dans_bloc_code = False

            else:

                dans_bloc_code = True

            i += 1

            continue

        if dans_bloc_code:

            tampon_code.append(ligne_brute)

            i += 1

            continue


        # ------------------------------------------------
        # TABLEAUX MARKDOWN
        # ------------------------------------------------

        if est_ligne_tableau(ligne_strip):

            if (
                not dans_tableau
                and i + 1 < len(lignes)
                and est_ligne_separateur_tableau(lignes[i + 1])
            ):

                dans_tableau = True
                tampon_tableau = [ligne_strip]

            elif dans_tableau:

                tampon_tableau.append(ligne_strip)

            i += 1

            continue

        elif dans_tableau:

            # Fin du tableau (ligne courante n'en fait
            # plus partie) : on le construit.
            ajouter_tableau(document, tampon_tableau)

            tampon_tableau = []
            dans_tableau = False

            # ne pas "continue" ici : on retraite la
            # ligne courante normalement ci-dessous


        if not ligne_strip:

            i += 1

            continue


        # ------------------------------------------------
        # LIGNES HORIZONTALES (---, ***, ___)
        # ------------------------------------------------

        if re.fullmatch(r"(-{3,}|\*{3,}|_{3,})", ligne_strip):

            i += 1

            continue


        # ------------------------------------------------
        # TITRES MARKDOWN
        # ------------------------------------------------

        if ligne_strip.startswith("#"):

            niveau = len(
                ligne_strip
            ) - len(
                ligne_strip.lstrip("#")
            )

            niveau = max(1, min(niveau, 4))

            # Retire TOUS les groupes de # en tête de ligne,
            # même s'il y en a plusieurs séparés par des
            # espaces (ex : "## # Titre" → "Titre"), au cas
            # où le modèle produirait un formatage imparfait.
            titre_texte = re.sub(
                r"^(#+\s*)+",
                "",
                ligne_strip
            ).strip()

            p = document.add_heading(
                "",
                level=niveau
            )

            ajouter_texte_formate(p, titre_texte)

            i += 1

            continue


        # ------------------------------------------------
        # LISTES À PUCES
        # ------------------------------------------------

        correspondance_puce = re.match(
            r"^[\-\*•]\s+(.*)",
            ligne_strip
        )

        if correspondance_puce:

            p = document.add_paragraph(
                style="List Bullet"
            )

            ajouter_texte_formate(
                p,
                correspondance_puce.group(1)
            )

            i += 1

            continue


        # ------------------------------------------------
        # LISTES NUMÉROTÉES
        # ------------------------------------------------

        correspondance_numero = re.match(
            r"^\d+[\.\)]\s+(.*)",
            ligne_strip
        )

        if correspondance_numero:

            p = document.add_paragraph(
                style="List Number"
            )

            ajouter_texte_formate(
                p,
                correspondance_numero.group(1)
            )

            i += 1

            continue


        # ------------------------------------------------
        # PARAGRAPHE NORMAL
        # ------------------------------------------------

        p = document.add_paragraph()

        ajouter_texte_formate(p, ligne_strip)

        i += 1


    # Si le texte se termine par un tableau non encore flushé
    if dans_tableau and tampon_tableau:

        ajouter_tableau(document, tampon_tableau)


    buffer = io.BytesIO()

    document.save(
        buffer
    )

    buffer.seek(0)

    return buffer


# ============================================================
# INTERFACE
# ============================================================

st.title(
    "📚 Générateur de Cours IFSI"
)


st.write(
    """
Colle l'URL d'une page de cours Lyon 1 (Moodle, "Livre", etc.)
— l'appli trouve automatiquement les vidéos qu'elle contient.
Tu peux aussi coller directement un ou plusieurs liens vidéo.
Une URL par ligne.
"""
)


urls_input = st.text_area(
    "URL de la page de cours (ou des vidéos)",
    height=200,
    placeholder=(
        "https://moodle.univ-lyon1.fr/mod/book/view.php?"
        "id=6424&chapterid=252\n"
        "https://...\n"
        "https://..."
    )
)


cle_api_utilisateur = st.text_input(
    "🔑 Ta clé API Gemini (gratuite)",
    type="password",
    placeholder="ex : AIzaSyD-9xY2kLmN3pQ4rS5tU6vW7xY8zA9bC0",
    help=(
        "Obtiens une clé gratuite sur "
        "https://aistudio.google.com/apikey (aucune carte "
        "bancaire requise). Elle n'est jamais enregistrée "
        "par cette appli : utilisée uniquement le temps de "
        "cette session, dans ton navigateur.\n\n"
        "⚠️ Le site ne la réaffiche qu'une seule fois à sa "
        "création — pense à la copier dans un fichier texte "
        "de ton côté pour pouvoir la réutiliser la prochaine "
        "fois sans avoir à en regénérer une."
    )
)


MODELES_DISPONIBLES = {
    "gemini-3.7-flash": (
        "Gemini 3.7 Flash — recommandé"
    ),
    "gemini-3.5-flash-lite": (
        "Gemini 3.5 Flash-Lite — moins de limitations"
    ),
}

modele_choisi = st.selectbox(
    "🤖 Modèle Gemini",
    options=list(MODELES_DISPONIBLES.keys()),
    format_func=lambda cle: MODELES_DISPONIBLES[cle],
    index=0,
    help=(
        "D'après nos tests sur des cours à plusieurs "
        "sujets :\n\n"
        "• **Gemini 3.7 Flash** couvre systématiquement "
        "l'intégralité du contenu (aucun sujet oublié). "
        "C'est le choix recommandé pour ne rien perdre.\n\n"
        "• **Gemini 3.5 Flash-Lite** est soumis à moins de "
        "limitations d'usage, mais a parfois oublié une "
        "partie des sujets sur des cours couvrant plusieurs "
        "thèmes distincts lors de nos tests. À réserver aux "
        "cours courts / à un seul sujet."
    )
)


# Note : pas d'authentification nécessaire — les liens MP3
# publics et l'API "modes" (URLs signées) fonctionnent sans
# cookie de session sur cette plateforme.


# ============================================================
# BOUTON
# ============================================================

if st.button(
    "🚀 Générer la fiche complète",
    type="primary"
):


    # --------------------------------------------------------
    # URL
    # --------------------------------------------------------

    urls = [
        url.strip()
        for url in urls_input.splitlines()
        if url.strip()
    ]


    if not urls:

        st.warning(
            "⚠️ Ajoute au moins une URL."
        )

        st.stop()


    # --------------------------------------------------------
    # SESSION HTTP (créée tôt : nécessaire pour développer
    # les pages de cours en liens vidéo, avant même de créer
    # le client Gemini)
    # --------------------------------------------------------

    session = requests.Session()

    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 "
            "(KHTML, like Gecko) "
            "Chrome/140.0 Safari/537.36"
        )
    })


    # --------------------------------------------------------
    # DÉVELOPPEMENT DES URLS (pages de cours → liens vidéo)
    # --------------------------------------------------------

    urls, titre_cours = developper_urls(
        urls,
        session
    )

    if not urls:

        st.error(
            "❌ Aucun lien vidéo n'a pu être trouvé à partir "
            "des URL fournies."
        )

        st.stop()


    # --------------------------------------------------------
    # API KEY
    # --------------------------------------------------------

    api_key = (
        cle_api_utilisateur.strip()
        if cle_api_utilisateur
        else None
    )

    if not api_key:

        try:

            api_key = st.secrets[
                "GEMINI_API_KEY"
            ]

        except Exception:

            api_key = None

    if not api_key:

        st.error(
            """
❌ Aucune clé API Gemini renseignée.

Colle ta clé dans le champ "🔑 Ta clé API Gemini" ci-dessus.

Tu peux en obtenir une gratuitement (sans carte bancaire) sur :

https://aistudio.google.com/apikey
"""
        )

        st.stop()


    # --------------------------------------------------------
    # CLIENT
    # --------------------------------------------------------

    try:

        client = genai.Client(
            api_key=api_key
        )

    except Exception as e:

        st.error(
            f"❌ Impossible de créer le client Gemini : {e}"
        )

        st.stop()


    # Liste des fichiers audio envoyés à Gemini (prêts,
    # état ACTIVE), en attendant l'appel final unique
    fichiers_geminis = []


    # ========================================================
    # TRAITEMENT
    # ========================================================

    activer_autoscroll()

    debut_total = time.time()

    with st.status(
        "Traitement des cours...",
        expanded=True
    ) as status:


        # ----------------------------------------------------
        # POUR CHAQUE URL : récupération + upload seulement
        # (l'analyse se fait plus bas, en un seul appel)
        # ----------------------------------------------------

        for i, url in enumerate(urls):


            st.write(
                f"## 🎧 Cours {i + 1}/{len(urls)}"
            )


            # ----------------------------------------------
            # Téléchargement
            # ----------------------------------------------

            debut_audio = time.time()

            fichier_local = recuperer_audio(
                url,
                i,
                session
            )

            duree_audio = time.time() - debut_audio


            if not fichier_local:

                st.warning(
                    f"⚠️ Aucun audio récupérable (ni MP3 direct, "
                    f"ni HLS) pour : {url} "
                    f"(⏱️ {duree_audio:.0f}s)"
                )

                continue

            st.write(
                f"⏱️ Récupération de l'audio : {duree_audio:.0f}s"
            )


            # ----------------------------------------------
            # Accélération (réduit le coût Gemini)
            # ----------------------------------------------

            if ACCELERATION_ACTIVEE:

                fichier_local = accelerer_audio(
                    fichier_local,
                    i + 1
                )


            # ----------------------------------------------
            # Upload vers Gemini (pas d'analyse ici)
            # ----------------------------------------------

            debut_upload = time.time()

            fichier_gemini = uploader_audio_gemini(
                client,
                fichier_local,
                i + 1
            )

            duree_upload = time.time() - debut_upload

            st.write(
                f"⏱️ Upload de l'audio {i + 1} : "
                f"{duree_upload:.0f}s"
            )


            # ----------------------------------------------
            # Suppression locale
            # ----------------------------------------------

            try:

                os.remove(
                    fichier_local
                )

            except Exception:

                pass


            if fichier_gemini:

                fichiers_geminis.append(
                    fichier_gemini
                )


        # ====================================================
        # VÉRIFICATION
        # ====================================================

        if not fichiers_geminis:

            status.update(
                label="❌ Aucun cours analysé",
                state="error",
                expanded=True
            )

            st.error(
                """
Aucun fichier audio n'a pu être envoyé à Gemini.

Vérifie notamment que les URL contiennent bien
un lien vers un fichier MP3 accessible.
"""
            )

            st.stop()


        # ====================================================
        # FICHE FINALE (un seul appel Gemini avec tous
        # les audios directement en entrée)
        # ====================================================

        debut_fiche = time.time()

        fiche = creer_fiche_finale(
            client,
            fichiers_geminis,
            modele_choisi
        )

        duree_fiche = time.time() - debut_fiche

        st.write(
            f"⏱️ Analyse + création de la fiche : "
            f"{duree_fiche:.0f}s"
        )


        # Nettoyage des fichiers Gemini, qu'il y ait eu
        # succès ou échec
        for fichier_gemini in fichiers_geminis:

            try:

                client.files.delete(
                    name=fichier_gemini.name
                )

            except Exception:

                pass


        if not fiche:

            status.update(
                label="❌ Impossible de créer la fiche",
                state="error",
                expanded=True
            )

            st.stop()


        duree_totale = time.time() - debut_total

        st.write(
            f"⏱️ **Durée totale du traitement : "
            f"{duree_totale / 60:.1f} min "
            f"({duree_totale:.0f}s)**"
        )

        status.update(
            label=(
                f"✅ Fiche générée avec succès en "
                f"{duree_totale / 60:.1f} min !"
            ),
            state="complete",
            expanded=True
        )


    # ========================================================
    # SAUVEGARDE EN SESSION (survit au clic sur "Télécharger")
    # ========================================================

    st.session_state["fiche_generee"] = fiche

    st.session_state["nom_fichier_docx"] = (
        nettoyer_nom_fichier(titre_cours) + ".docx"
    )


# ============================================================
# AFFICHAGE DU RÉSULTAT (hors du bloc bouton, pour survivre
# au clic sur "Télécharger" qui relance le script)
# ============================================================

if st.session_state.get("fiche_generee"):

    st.success(
        "🎉 Ta fiche de révision est prête !"
    )

    st.markdown(
        st.session_state["fiche_generee"]
    )


    st.write(
        "### 📄 Télécharger la fiche"
    )

    buffer = creer_word(
        st.session_state["fiche_generee"]
    )

    colonne_telecharger, colonne_recharger = st.columns(2)

    with colonne_telecharger:

        st.download_button(
            label="📥 Télécharger la fiche (.docx)",
            data=buffer,
            file_name=st.session_state["nom_fichier_docx"],
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.wordprocessingml.document"
            )
        )

    with colonne_recharger:

        if st.button("🔄 Nouveau cours"):

            st.session_state.pop("fiche_generee", None)
            st.session_state.pop("nom_fichier_docx", None)

            st.rerun()