// =====================================================================
// Page « Détection des irritants dans les mails » (irritants.html)
// Chargé APRÈS app.js : réutilise ses réglages (fournisseur, modèle, clé), sa connexion Gmail
// (lecture seule), sa lecture des mails (.eml, .zip, Gmail) et son appel au modèle d'IA.
// Rien n'est enregistré (pas de Firestore), aucun mail n'est envoyé.
// =====================================================================

// ---------- 1. Instructions de l'agent (prompt) ----------
const INSTRUCTIONS_IRRITANTS_MAILS = `RÔLE
Tu es l'agent de détection des irritants d'une équipe de services. Tu lis les mails d'une boîte de travail et tu fais émerger, à partir de ces seuls mails, les irritants : les frictions récurrentes qui font perdre du temps, de l'énergie ou de la qualité à l'équipe. Tu ne connais pas à l'avance la liste des irritants : c'est à toi de les découvrir, de les nommer et de les justifier. Tu proposes ; un humain valide.

DONNÉES DISPONIBLES
Les mails de la boîte (réception et envoyés), lus via le connecteur Gmail. Les mails envoyés par la boîte sont signés au nom du service. Les autres viennent de collègues, clients, plateformes ou partenaires.
Lis tous les mails de la période demandée avant de conclure, et lis chaque fil en entier : la réponse suivante peut montrer qu'un problème était réglé en une heure, ou au contraire qu'il a coûté cher.

SIGNAUX DE FRICTION
Un irritant est rarement écrit en toutes lettres. Cherche des signaux, puis remonte à la cause :
- Relances (même demande envoyée plusieurs fois, « je reviens vers toi », « deuxième (ou troisième) demande », escalade vers un responsable).
- Urgence et pression sur les délais (dates qui se rapprochent ou sont avancées, réunions placées tôt le matin, tard le soir ou le week-end, travail rendu à des heures inhabituelles).
- Retards et indisponibilités (contributions en retard, personnes prises ailleurs, excuses pour le délai).
- Allers-retours et reprises (travail refait ou corrigé, demandes reformulées plusieurs fois).
- Informations manquantes ou contradictoires (élément attendu absent, donnée introuvable ou périmée, deux sources contradictoires).
- Surprises tardives (problème découvert au dernier moment ou après coup alors qu'il était connaissable plus tôt).
- Effort perdu (travail fait pour rien ou sans suite).

MÉFIE-TOI DES FAUX POSITIFS
Un signal seul ne suffit pas. N'est pas un irritant :
- un ton ferme ou un mot « urgent » quand la demande est banale et traitée normalement ;
- une formule standard ou réglementaire d'un client ou d'une plateforme (rappel des règles, délai de réponse, mise en demeure type) ;
- un désaccord ou un arbitrage rendu à temps par la personne qui en a la charge ;
- plusieurs allers-retours qui convergent vite vers une solution.
Retiens un irritant quand il y a un coût visible (temps perdu, travail refait, risque pour le résultat, marge, qualité) et une cause qui pourrait se reproduire.

MÉTHODE
1. Parcours les mails et relève chaque signal avec : identifiant ou date du mail, expéditeur, objet, phrase courte qui le montre.
2. Écarte les faux positifs et garde ceux qui ont un coût et une cause.
3. Regroupe les signaux par cause commune et donne à chaque groupe un nom court que tu choisis toi-même, qui décrit la cause (pas le symptôme). Ne force pas un mail dans un groupe : un irritant vu une seule fois peut rester seul.
4. Pour chaque catégorie, compte la fréquence (nombre de mails distincts) et attribue un niveau de confiance :
   - Élevé : plusieurs mails, ou un mail où le coût et la cause sont explicites ;
   - Moyen : un seul mail clair, ou plusieurs mails ambigus ;
   - Faible : interprétation possible, à confirmer par l'équipe.
5. Propose pour chaque catégorie une piste d'amélioration (processus, référentiel, outil ou agent IA) et dis ce qu'il faudrait vérifier auprès de l'équipe.

FORMAT DE RÉPONSE
- Un tableau récapitulatif trié par fréquence puis par confiance : Catégorie (ton nom) | Fréquence | Confiance | Piste d'amélioration.
- Puis, pour chaque catégorie :
  - une phrase qui décrit la cause ;
  - les mails concernés, chacun avec la date, l'expéditeur et un extrait court et exact (entre guillemets, sans reformuler) ;
  - le coût observé (temps, risque, marge, qualité), sans inventer de chiffre absent des mails.
- Une courte section « Signaux écartés » : mails qui avaient l'air tendus mais non retenus, avec la raison en une ligne.
- Une section « Limites » : ce que les mails ne permettent pas de dire (fréquence réelle, causes à confirmer, mails absents).

RÈGLES
- N'invente rien. Chaque catégorie s'appuie sur au moins un extrait réel de mail. Si tu doutes, baisse la confiance et dis-le.
- Cite exactement : l'extrait doit se retrouver mot pour mot dans le mail.
- Ne compte pas deux fois le même problème quand un fil le répète : compte les mails qui apportent un signal nouveau, et précise-le.
- N'exprime pas de jugement sur les personnes : un irritant décrit un fonctionnement, pas un coupable.
- N'expose pas le contenu des mails au-delà de ce qui est nécessaire pour justifier une catégorie.
- Tu ne réponds à aucun mail et tu n'envoies rien.`;

// ---------- 2. Les mails, en texte, pour le modèle ----------
// Chaque mail reçoit un identifiant (M1, M2…) dans l'ordre chronologique, et la mention « envoyé » ou « reçu ».
function mailEnvoye(m) {
  const compte = (etat.compteGoogle || "").toLowerCase();
  return (compte && m.de.toLowerCase().includes(compte)) || /Service facturation/i.test(m.de);
}
function mailsIrritantsEnTexte() {
  return etat.mails.map((m, i) =>
    `### M${i + 1} · ${mailEnvoye(m) ? "ENVOYÉ par la boîte" : "REÇU"} · ${m.dateTexte}\nDe : ${m.de}\nÀ : ${m.a}${m.cc ? "\nCc : " + m.cc : ""}\nObjet : ${m.objet}\n\n${m.corps}`).join("\n\n");
}
function periodeMails() {
  const dates = etat.mails.filter((m) => m.horodatage).map((m) => m.horodatage);
  if (!dates.length) return "dates inconnues";
  const f = (t) => new Date(t).toLocaleDateString("fr-FR");
  return `du ${f(Math.min(...dates))} au ${f(Math.max(...dates))}`;
}
function systemeIrritants() {
  const envoyes = etat.mails.filter(mailEnvoye).length;
  const boite = etat.compteGoogle ? `la boîte Gmail ${etat.compteGoogle}` : "une boîte de démonstration (mails fictifs chargés depuis des fichiers .eml)";
  return [
    INSTRUCTIONS_IRRITANTS_MAILS,
    `## CONTEXTE DE CETTE ANALYSE
Date du jour : ${new Date().toLocaleDateString("fr-FR")}.
Source : ${boite}. Période demandée : tous les mails fournis ci-dessous, ${periodeMails()}.
${etat.mails.length} mails : ${envoyes} envoyés par la boîte, ${etat.mails.length - envoyes} reçus. Brouillons, spam, corbeille et mails envoyés par Google sont exclus${etat.mails.length >= 300 ? " ; seuls les 300 mails les plus récents ont été lus (signale-le dans les Limites)" : ""}.
Le texte complet de chaque mail est fourni. Chaque mail porte un identifiant (M1, M2…) : utilise-le devant chaque extrait, avec la date et l'expéditeur. Ne mets pas ces identifiants dans le tableau récapitulatif, qui n'a pas de colonne de mails. Les fils se reconnaissent à leur objet (« Re : », « TR : »…).
Les pièces jointes ne sont pas fournies : si un mail y renvoie, dis-le dans les Limites.
Réponds en français, directement en Markdown (pas de bloc de code).`,
    "## MAILS\n" + mailsIrritantsEnTexte(),
  ].join("\n\n");
}

// ---------- 3. Bouton « Analyser les irritants des mails » ----------
let analyseEnCours = false;
function majPageIrritants() {
  const n = etat.mails.length;
  $("btnIrritants").disabled = analyseEnCours || !n;
  $("btnIrritants").title = n ? "" : "Connectez-vous d'abord à Gmail ou chargez des mails";
  $("aideIrritants").hidden = n > 0;
  const envoyes = etat.mails.filter(mailEnvoye).length;
  $("irrMails").textContent = !n ? ""
    : `${etat.demo ? "Mails de démonstration (fictifs)" : etat.compteGoogle ? "Boîte " + etat.compteGoogle : "Mails chargés depuis l'ordinateur"} : ${n} mails (${envoyes} envoyés, ${n - envoyes} reçus) ${periodeMails()}.`
      + ($("provider").value === "groq" ? " Attention : avec Groq gratuit, l'envoi de tous les mails dépassera probablement sa limite de 8 000 tokens par minute. Choisissez Claude." : "");
}
document.addEventListener("mailsCharges", majPageIrritants);
$("provider").addEventListener("change", majPageIrritants);
// La déconnexion (app.js) vide les mails : on remet la page à jour juste après.
$("btnDeconnexion").addEventListener("click", () => setTimeout(majPageIrritants, 0));

$("btnIrritants").addEventListener("click", async () => {
  if (!etat.mails.length) { $("diagStatus").textContent = "Aucun mail chargé : connectez-vous à Gmail ou chargez des mails depuis l'ordinateur."; return; }
  const zone = $("diagResultat");
  analyseEnCours = true; majPageIrritants();
  const arreter = demarrerAttente($("diagStatus"), `L'assistant lit ${etat.mails.length} mails…`);
  zone.hidden = true;
  let dernierRendu = 0;
  try {
    const texte = await appelerModele((partiel) => {
      if (zone.hidden) { zone.hidden = false; arreter("L'assistant écrit…"); }
      if (Date.now() - dernierRendu > 150) { dernierRendu = Date.now(); zone.innerHTML = DOMPurify.sanitize(marked.parse(partiel), { FORBID_TAGS: ["img"] }); }
    }, { systeme: systemeIrritants(), messages: [{ role: "user", content: `Analyse les irritants de l'équipe à partir de tous les mails fournis (${periodeMails()}).` }] });
    zone.innerHTML = DOMPurify.sanitize(marked.parse(texte), { FORBID_TAGS: ["img"] });
    zone.hidden = false;
    afficherConso("Irritants dans les mails");
    arreter("");
  } catch (e) {
    arreter("");
    zone.hidden = false;
    zone.replaceChildren(Object.assign(document.createElement("div"), { className: "error", textContent: e.message }));
  } finally {
    analyseEnCours = false; majPageIrritants();
  }
});

majPageIrritants();
