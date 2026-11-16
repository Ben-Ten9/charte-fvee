/*
 * Panneau « Charte FVEE » (Word, PowerPoint, Excel).
 * Contenu : catalogue.json, généré par deploiement\scripts\complement\Build-Complement.ps1
 * à partir de Catalogue.ps1 et de sources\couleurs\ (même source que QuickPart et Images de marque).
 */
'use strict';

let hote = null;        // 'Word' | 'PowerPoint' | 'Excel' | null (aperçu dans un navigateur)
let catalogue = null;
let onglet = null;
let mode = 'texte';     // couleurs : 'texte' | 'remplissage'
let echelle = 1;        // taille d'insertion : 0.6 / 1 / 1.6

const $ = (s) => document.querySelector(s);
const el = (tag, props = {}, ...enfants) => {
  const n = Object.assign(document.createElement(tag), props);
  for (const e of enfants) if (e != null) n.append(e);
  return n;
};
const normaliser = (s) => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
const cmEnPoints = (cm) => Math.round(cm * 28.35);

// ---------------------------------------------------------------- démarrage
if (typeof Office !== 'undefined') {
  Office.onReady((info) => {
    const H = Office.HostType;
    hote = { [H.Word]: 'Word', [H.PowerPoint]: 'PowerPoint', [H.Excel]: 'Excel' }[info.host] || null;
    demarrer();
  });
} else {
  demarrer();   // office.js injoignable : aperçu seul
}

async function demarrer() {
  try {
    const r = await fetch('catalogue.json', { cache: 'no-cache' });
    if (!r.ok) throw new Error(`catalogue.json : HTTP ${r.status}`);
    catalogue = await r.json();
  } catch (e) {
    $('#contenu').append(el('p', { className: 'vide', textContent: `Catalogue introuvable (${e.message}).` }));
    return;
  }
  if (hote === 'Excel') mode = 'remplissage';
  onglet = catalogue.categories[0].id;
  dessinerOnglets();
  dessiner();
  $('#recherche').addEventListener('input', dessiner);
  if (!hote) etat('Aperçu hors Office : l\'insertion ne fonctionne que depuis Word, PowerPoint ou Excel.');
}

function tousLesOnglets() {
  return [...catalogue.categories.map((c) => ({ id: c.id, nom: c.nom })),
          { id: 'couleurs', nom: 'Couleurs' }, { id: 'polices', nom: 'Polices' }];
}

function dessinerOnglets() {
  const nav = $('#onglets');
  nav.replaceChildren(...tousLesOnglets().map((o) => {
    const b = el('button', { type: 'button', textContent: o.nom, role: 'tab' });
    b.setAttribute('aria-selected', String(o.id === onglet));
    b.onclick = () => { onglet = o.id; $('#recherche').value = ''; dessinerOnglets(); dessiner(); };
    return b;
  }));
}

function dessiner() {
  const zone = $('#contenu');
  const q = normaliser($('#recherche').value.trim());
  zone.replaceChildren();
  zone.scrollTop = 0;

  if (q) {                                   // recherche : toutes les catégories d'images
    let total = 0;
    for (const c of catalogue.categories) {
      const res = c.elements.filter((e) => normaliser(`${e.nom} ${c.nom}`).includes(q));
      if (!res.length) continue;
      total += res.length;
      zone.append(el('h2', { textContent: `${c.nom} (${res.length})` }), grille(c, res));
    }
    if (!total) zone.append(el('p', { className: 'vide', textContent: 'Aucun résultat.' }));
    return;
  }
  if (onglet === 'couleurs') return dessinerCouleurs(zone);
  if (onglet === 'polices') return dessinerPolices(zone);

  const c = catalogue.categories.find((x) => x.id === onglet);
  zone.append(barreTaille());
  if (c.id === 'Logos') zone.append(el('p', { className: 'aide',
    textContent: 'Les logos blancs sont affichés sur fond bleu : pose-les sur un fond de couleur.' }));
  zone.append(grille(c, c.elements));
}

// ---------------------------------------------------------------- images
function barreTaille() {
  const tailles = [['Petite', 0.6], ['Normale', 1], ['Grande', 1.6]];
  const seg = el('div', { className: 'segments' }, ...tailles.map(([nom, v]) => {
    const b = el('button', { type: 'button', textContent: nom });
    b.setAttribute('aria-pressed', String(v === echelle));
    b.onclick = () => { echelle = v; dessiner(); };
    return b;
  }));
  return el('div', { className: 'barre' }, 'Taille à l\'insertion :', seg);
}

function grille(categorie, elements) {
  const g = el('div', { className: 'grille' + (categorie.id === 'Logos' ? ' logos' : '') });
  for (const e of elements) {
    const t = el('button', { type: 'button', className: 'tuile', title: `Insérer « ${e.nom} »` },
      el('div', { className: 'vignette' + (e.blanc ? ' blanc' : '') },
        el('img', { src: e.fichier, alt: '', loading: 'lazy' })),
      el('span', { className: 'nom', textContent: e.nom }));
    t.onclick = () => inserer(e);
    g.append(t);
  }
  return g;
}

async function inserer(e) {
  if (!hote) return etat('Ouvre ce panneau depuis Word, PowerPoint ou Excel pour insérer.', true);
  const largeur = Math.round(cmEnPoints(e.largeurCm) * echelle);
  try {
    if (e.type === 'svg' && Office.context.requirements.isSetSupported('ImageCoercion', '1.2')) {
      const svg = await (await fetch(e.fichier)).text();
      await definirSelection(nettoyerSvg(svg), { coercionType: Office.CoercionType.XmlSvg, imageWidth: largeur });
    } else {
      const b64 = e.type === 'svg' ? await svgEnPng(e.fichier, largeur) : await enBase64(e.fichier);
      await definirSelection(b64, { coercionType: Office.CoercionType.Image, imageWidth: largeur });
    }
    etat(`Inséré : ${e.nom}`);
  } catch (err) {
    etat(`Insertion impossible : ${message(err)}`, true);
  }
}

function definirSelection(donnees, options) {
  return new Promise((ok, ko) => Office.context.document.setSelectedDataAsync(donnees, options,
    (r) => (r.status === Office.AsyncResultStatus.Succeeded ? ok() : ko(r.error))));
}

// Office n'accepte pas toujours le prologue XML, le DOCTYPE et les commentaires d'Illustrator
function nettoyerSvg(svg) {
  return svg.replace(/<\?xml[\s\S]*?\?>/g, '').replace(/<!DOCTYPE[\s\S]*?>/gi, '').replace(/<!--[\s\S]*?-->/g, '').trim();
}

async function enBase64(url) {
  const blob = await (await fetch(url)).blob();
  return new Promise((ok, ko) => {
    const fr = new FileReader();
    fr.onload = () => ok(String(fr.result).split(',')[1]);
    fr.onerror = () => ko(fr.error);
    fr.readAsDataURL(blob);
  });
}

// Repli si l'hôte ne sait pas insérer du SVG : image PNG haute définition
async function svgEnPng(url, largeurPt) {
  const img = new Image();
  img.src = url;
  await img.decode();
  const l = Math.round(largeurPt * 4);
  const h = Math.round(l * (img.naturalHeight || 1) / (img.naturalWidth || 1));
  const c = Object.assign(document.createElement('canvas'), { width: l, height: h });
  c.getContext('2d').drawImage(img, 0, 0, l, h);
  return c.toDataURL('image/png').split(',')[1];
}

// ---------------------------------------------------------------- couleurs
function dessinerCouleurs(zone) {
  const seg = el('div', { className: 'segments' }, ...[['texte', 'Texte'], ['remplissage', 'Remplissage']].map(([v, nom]) => {
    const b = el('button', { type: 'button', textContent: nom });
    b.setAttribute('aria-pressed', String(v === mode));
    b.onclick = () => { mode = v; dessiner(); };
    return b;
  }));
  const aides = {
    Word: { texte: 'Sélectionne du texte, puis clique une couleur.',
            remplissage: 'Sélectionne une forme ou une cellule de tableau (fond pour un logo blanc).' },
    PowerPoint: { texte: 'Sélectionne du texte ou une zone de texte.',
                  remplissage: 'Sélectionne une ou plusieurs formes (ex. un rectangle derrière un logo blanc).' },
    Excel: { texte: 'Sélectionne des cellules : couleur du texte.', remplissage: 'Sélectionne des cellules : couleur de fond.' },
  };
  zone.append(el('div', { className: 'barre' }, 'Appliquer à :', seg),
    el('p', { className: 'aide', textContent: hote ? aides[hote][mode] : 'Clique une couleur pour l\'appliquer à la sélection.' }));

  const entete = el('div', { className: 'entete-nuances' }, el('span'));
  for (const n of catalogue.couleurs[0].nuances) entete.append(el('span', { textContent: n.nom }));
  zone.append(entete);

  for (const f of catalogue.couleurs) {
    const ligne = el('div', { className: 'famille' }, el('span', { className: 'libelle', textContent: f.nom }));
    for (const n of f.nuances) {
      const p = el('button', { type: 'button', className: 'pastille', title: `${f.nom} ${n.nom} - ${n.hex}`, textContent: n.hex.slice(1) });
      p.style.background = n.hex;
      p.style.color = estClair(n.hex) ? '#242424' : '#fff';
      p.onclick = () => appliquerCouleur(n.hex, `${f.nom} ${n.nom.toLowerCase()}`);
      ligne.append(p);
    }
    zone.append(ligne);
  }
  zone.append(el('p', { className: 'aide', textContent:
    'Code couleur aussi utilisable dans Couleurs > Autres couleurs > Personnalisées > Hex.' }));
}

function estClair(hex) {
  const [r, g, b] = [1, 3, 5].map((i) => parseInt(hex.substr(i, 2), 16) / 255);
  return 0.2126 * r + 0.7152 * g + 0.0722 * b > 0.55;
}

async function appliquerCouleur(hex, libelle) {
  if (!hote) return etat(`${libelle} : ${hex}`);
  try {
    const cible = mode === 'texte'
      ? await surTexte((police) => { police.color = hex; })
      : await surRemplissage(hex);
    etat(`${libelle} (${hex}) appliqué : ${cible}`);
  } catch (err) {
    etat(message(err), true);
  }
}

// Applique une modification à la police de la sélection (couleur ou nom de police)
async function surTexte(modifier) {
  if (hote === 'Word') {
    return Word.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      modifier(sel.font);
      await ctx.sync();
      return 'texte sélectionné';
    });
  }
  if (hote === 'Excel') {
    return Excel.run(async (ctx) => {
      modifier(ctx.workbook.getSelectedRange().format.font);
      await ctx.sync();
      return 'cellules sélectionnées';
    });
  }
  return PowerPoint.run(async (ctx) => {
    // 1) texte surligné à l'intérieur d'une zone de texte
    try {
      const tr = ctx.presentation.getSelectedTextRange();
      tr.load('length');
      await ctx.sync();
      if (tr.length > 0) { modifier(tr.font); await ctx.sync(); return 'texte sélectionné'; }
    } catch (_) { /* pas de texte sélectionné : on passe aux formes */ }
    // 2) zones de texte / formes sélectionnées entières
    const formes = ctx.presentation.getSelectedShapes();
    formes.load('items');
    await ctx.sync();
    if (!formes.items.length) throw new Error('Sélectionne du texte ou une zone de texte.');
    for (const f of formes.items) modifier(f.textFrame.textRange.font);
    await ctx.sync();
    return `${formes.items.length} forme(s)`;
  });
}

async function surRemplissage(hex) {
  if (hote === 'Excel') {
    return Excel.run(async (ctx) => {
      ctx.workbook.getSelectedRange().format.fill.color = hex;
      await ctx.sync();
      return 'fond des cellules';
    });
  }
  if (hote === 'PowerPoint') {
    return PowerPoint.run(async (ctx) => {
      const formes = ctx.presentation.getSelectedShapes();
      formes.load('items/type');
      await ctx.sync();
      if (!formes.items.length) throw new Error('Sélectionne une forme (ex. un rectangle placé derrière le logo).');
      for (const f of formes.items) f.fill.setSolidColor(hex);
      await ctx.sync();
      return `${formes.items.length} forme(s)`;
    });
  }
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    const apiFormes = Office.context.requirements.isSetSupported('WordApiDesktop', '1.2');
    let indice = '';
    // Forme sélectionnée (API récente de Word)
    if (apiFormes) {
      try {
        let formes = await formesWord(ctx, sel);
        // Forme flottante : la sélection se réduit souvent à un point, sans l'ancre de la forme.
        // On regarde alors le paragraphe, seulement s'il porte une seule forme (pour ne pas en colorer d'autres).
        if (!formes.racine.length) {
          const para = await formesWord(ctx, sel.paragraphs.getFirst().getRange('Whole'));
          if (para.racine.length === 1) formes = para;
          else if (para.racine.length > 1) indice = 'Plusieurs formes sont ancrées dans ce paragraphe : sélectionne la forme elle-même.';
        }
        if (formes.feuilles.length) {
          for (const f of formes.feuilles) f.fill.setSolidColor(hex);
          await ctx.sync();
          return `${formes.feuilles.length} forme(s)`;
        }
        if (formes.racine.some((f) => f.type === 'Picture')) {
          indice = 'C\'est une image, pas une forme : clic droit > Convertir en forme, puis réessaie.';
        }
      } catch (err) {
        indice = `Forme non modifiable par le complément (${message(err)}).`;
      }
    }
    // Formes libres (SVG converti en forme) : absentes de l'API, on modifie le XML du passage
    const viaOoxml = await remplissageOoxml(ctx, sel, hex);
    if (viaOoxml) return viaOoxml;
    const cellule = sel.parentTableCellOrNullObject;
    cellule.load('isNullObject');
    await ctx.sync();
    if (!cellule.isNullObject) {
      cellule.shadingColor = hex;
      await ctx.sync();
      return 'cellule de tableau';
    }
    if (indice) throw new Error(indice);
    if (!apiFormes) {
      throw new Error(`Cette version de Word ne permet pas au complément de colorer une forme : `
        + `Format de la forme > Remplissage > Autres couleurs > Hex ${hex.slice(1)}.`);
    }
    throw new Error('Aucune forme trouvée dans la sélection ni dans son paragraphe, et pas de cellule de tableau. '
      + 'Clique sur le bord de la forme pour la sélectionner, puis réessaie.');
  });
}

// Remplissage par OOXML : sélection d'abord, puis paragraphe d'ancrage (s'il ne porte qu'un dessin avec formes)
async function remplissageOoxml(ctx, sel, hex) {
  const plages = [sel, sel.paragraphs.getFirst().getRange('Content')];
  for (let i = 0; i < plages.length; i++) {
    const ox = plages[i].getOoxml();
    await ctx.sync();
    const r = recolorerOoxml(ox.value, hex);
    if (!r || !r.formes) continue;
    if (i === 1 && r.dessins > 1) {
      throw new Error('Plusieurs formes sont ancrées dans ce paragraphe : le complément ne sait pas laquelle est sélectionnée. '
        + `Format de la forme > Remplissage > Autres couleurs > Hex ${hex.slice(1)}.`);
    }
    plages[i].insertOoxml(r.xml, 'Replace');
    await ctx.sync();
    return `${r.formes} forme(s)`;
  }
  return null;
}

const NS = {
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  v: 'urn:schemas-microsoft-com:vml',
};
const REMPLISSAGES = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'];
const enfantsA = (n, noms) => [...n.children].filter((c) => c.namespaceURI === NS.a && noms.includes(c.localName));

// Remplace le remplissage des formes (wps:wsp) d'un paquet OOXML. Ne touche pas aux contours.
// Si certaines formes sont remplies, seules celles-là changent (une forme « sans remplissage » le reste).
function recolorerOoxml(xml, hex) {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length) return null;
  const spPrs = [...doc.getElementsByTagNameNS(NS.wps, 'wsp')]
    .map((f) => [...f.children].find((c) => c.namespaceURI === NS.wps && c.localName === 'spPr'))
    .filter(Boolean);
  if (!spPrs.length) return { formes: 0 };
  const remplies = spPrs.filter((sp) => enfantsA(sp, REMPLISSAGES.filter((n) => n !== 'noFill')).length);
  const cibles = remplies.length ? remplies : spPrs;
  for (const sp of cibles) {
    for (const c of enfantsA(sp, REMPLISSAGES)) sp.removeChild(c);
    const fill = doc.createElementNS(NS.a, 'a:solidFill');
    const clr = doc.createElementNS(NS.a, 'a:srgbClr');
    clr.setAttribute('val', hex.slice(1).toUpperCase());
    fill.appendChild(clr);
    // Ordre imposé par le schéma : xfrm, géométrie, remplissage, contour...
    const avant = enfantsA(sp, ['custGeom', 'prstGeom'])[0] || enfantsA(sp, ['xfrm'])[0];
    sp.insertBefore(fill, avant ? avant.nextSibling : sp.firstChild);
  }
  // Version VML de secours (anciens Word), si présente
  for (const s of doc.getElementsByTagNameNS(NS.v, 'shape')) {
    if (s.getAttribute('filled') !== 'f' && s.getAttribute('filled') !== 'false') s.setAttribute('fillcolor', hex);
  }
  const dessins = [...doc.getElementsByTagNameNS(NS.wp, 'anchor'), ...doc.getElementsByTagNameNS(NS.wp, 'inline')]
    .filter((d) => d.getElementsByTagNameNS(NS.wps, 'wsp').length).length;
  return { xml: new XMLSerializer().serializeToString(doc), formes: cibles.length, dessins };
}

// Formes d'une plage Word : « racine » = formes ancrées, « feuilles » = formes colorables
// (les groupes et zones de dessin, typiques d'un SVG converti en forme, sont dépliés)
async function formesWord(ctx, plage) {
  const racine = plage.shapes;
  racine.load('items/type');
  await ctx.sync();
  return { racine: racine.items, feuilles: await deplierFormes(ctx, racine.items) };
}

async function deplierFormes(ctx, formes) {
  const feuilles = [];
  const conteneurs = [];
  for (const f of formes) {
    if (f.type === 'Group' || f.type === 'Canvas') {
      const s = (f.type === 'Group' ? f.group : f.canvas).shapes;
      s.load('items/type');
      conteneurs.push(s);
    } else if (f.type === 'GeometricShape' || f.type === 'TextBox') {
      feuilles.push(f);
    }
  }
  if (conteneurs.length) {
    await ctx.sync();
    for (const c of conteneurs) feuilles.push(...await deplierFormes(ctx, c.items));
  }
  return feuilles;
}

// ---------------------------------------------------------------- polices
function dessinerPolices(zone) {
  zone.append(el('p', { className: 'aide', textContent:
    'Sélectionne du texte, puis applique la police. Si une police n\'est pas installée sur le poste, Office affiche une police de remplacement.' }));
  for (const p of catalogue.polices) {
    const apercu = el('div', { className: 'apercu', textContent: p.nom });
    apercu.style.fontFamily = `"${p.nom}", "Segoe UI", sans-serif`;
    const b = el('button', { type: 'button', className: 'bouton', textContent: 'Appliquer' });
    b.onclick = () => appliquerPolice(p.nom);
    zone.append(el('div', { className: 'police' }, el('div', {}, apercu, el('div', { className: 'role', textContent: p.role })), b));
  }
}

async function appliquerPolice(nom) {
  if (!hote) return etat(`Police : ${nom}`);
  try {
    const cible = await surTexte((police) => { police.name = nom; });
    etat(`${nom} appliquée : ${cible}`);
  } catch (err) {
    etat(message(err), true);
  }
}

// ---------------------------------------------------------------- état
let minuterie = null;
function etat(texte, erreur = false) {
  const z = $('#etat');
  z.textContent = texte;
  z.className = 'visible' + (erreur ? ' erreur' : '');
  clearTimeout(minuterie);
  minuterie = setTimeout(() => { z.className = ''; }, erreur ? 8000 : 4000);
}

function message(err) {
  if (!err) return 'erreur inconnue';
  const detail = err.message || err.code || String(err);
  if (err.code === 'InvalidSelection') return `Sélection non compatible avec cette action (${detail}).`;
  return detail;
}
