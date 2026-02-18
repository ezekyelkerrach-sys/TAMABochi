// ============================================================================
//  SUIVI 56 — Code.gs — Backend complet
//  Phases 1-10 intégrées depuis TAMABochi
//  Utopia 56 — Février 2026
// ============================================================================

// ============================================================================
//                               CONFIGURATION
// ============================================================================

var SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// Phase 9 : Cache serveur (30 min TTL)
var CACHE_TTL = 1800; // 30 minutes en secondes

// ============================================================================
//                                ROUTAGE & HTML
// ============================================================================

// Phase 3 : Titre dynamique depuis BDD_ANTENNE
function doGet(e) {
  var antenneConfig = getAntenneConfig_();
  var title = (antenneConfig.NOM_APP || 'SUIVI 56') + ' - ' + (antenneConfig.NOM_ANTENNE || 'Utopia 56');
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle(title)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================================
//                            SECURITE & CONNEXION
// ============================================================================

// Phase 8 : Hash SHA-256 des mots de passe
function hashSHA256_(input) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input);
  return raw.map(function(byte) {
    var v = (byte < 0) ? byte + 256 : byte;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

function loginUser(email, password) {
  const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('USERS');
  const data = ws.getDataRange().getValues();
  var hashedInput = hashSHA256_(password);
  for (let i = 1; i < data.length; i++) {
    var storedPass = String(data[i][1]);
    // Compatibilité : accepter mot de passe en clair OU hashé
    var match = (storedPass === password) || (storedPass === hashedInput);
    if (data[i][0] == email && match) {
      // Si le mot de passe était en clair, le hasher automatiquement
      if (storedPass === password && storedPass !== hashedInput) {
        ws.getRange(i + 1, 2).setValue(hashedInput);
      }
      return { 
        success: true, 
        token: Utilities.base64Encode(email + "_" + new Date().getTime()),
        user: { email: data[i][0], role: data[i][2], nom: data[i][3] }
      };
    }
  }
  return { success: false, message: "ID ou mot de passe incorrect." };
}

// ============================================================================
//                   Phase 3 : BDD_ANTENNE — Config par antenne
// ============================================================================

function getAntenneConfig_() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('antenne_config');
    if (cached) return JSON.parse(cached);
    
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var ws = ss.getSheetByName('BDD_ANTENNE');
    if (!ws) return {};
    var data = ws.getDataRange().getValues();
    var config = {};
    for (var i = 0; i < data.length; i++) {
      var key = String(data[i][0] || '').trim();
      var val = String(data[i][1] || '').trim();
      if (key) config[key] = val;
    }
    cache.put('antenne_config', JSON.stringify(config), CACHE_TTL);
    return config;
  } catch(e) {
    return {};
  }
}

// Fonction exposée au client pour récupérer la config antenne
function getAntenneConfig() {
  return getAntenneConfig_();
}

// ============================================================================
//                  API : CONFIGURATION & SCENARIOS
//  Phase 2 : Ajout colonnes GENRES, TYPES_PROFILS, ROLES_FOYER, TRANCHES_AGE
//  Phase 5 : Scénarios filtrés par profil (colonne Profils_Visibles)
//  Phase 9 : Cache serveur CacheService
// ============================================================================

function getConfigData() {
  // Phase 9 : Vérifier le cache d'abord
  var cache = CacheService.getScriptCache();
  var cached = cache.get('config_data');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) { /* cache corrompu, on recharge */ }
  }
  
  const ss = SpreadsheetApp.openById(SHEET_ID);
  
  const wsConfig = ss.getSheetByName('BDD_CONFIG');
  const dataConfig = wsConfig.getDataRange().getValues();
  let config = {};
  
  if (dataConfig.length > 0) {
    const headers = dataConfig[0];
    for (let col = 0; col < headers.length; col++) {
      let key = headers[col];
      let values = [];
      for (let row = 1; row < dataConfig.length; row++) {
        if (dataConfig[row][col] !== "") {
          values.push(dataConfig[row][col]);
        }
      }
      config[key] = values;
    }
  }

  // Phase 5 : Scénarios avec colonne Profils_Visibles (col N = index 13)
  const wsScenarios = ss.getSheetByName('BDD_SCENARIOS');
  if (wsScenarios) { 
    const dataScenarios = wsScenarios.getDataRange().getValues();
    let scenarios = {};
    for (let i = 1; i < dataScenarios.length; i++) {
      let type = dataScenarios[i][0];
      if(type) {
        scenarios[type] = {
          showLieuEvent: !!dataScenarios[i][1],
          showStatutPresence: !!dataScenarios[i][2],
          showLieuVie: !!dataScenarios[i][3],
          showStatutInstitution: !!dataScenarios[i][4],
          showMotifStatutInstitution: !!dataScenarios[i][5],
          showMateriel: !!dataScenarios[i][6],
          showVulnerabilites: !!dataScenarios[i][7],
          updateStatutPresence: !!dataScenarios[i][8],
          updateLieuVie: !!dataScenarios[i][9],
          updateStatutInstitution: !!dataScenarios[i][10],
          updateMotifStatutInstitution: !!dataScenarios[i][11],
          updateVulnerabilites: !!dataScenarios[i][12],
          // Phase 5 : Profils visibles
          profilsVisibles: String(dataScenarios[i][13] || '').trim()
        };
      }
    }
    config['SCENARIOS'] = scenarios;
  }
  
  // Phase 3 : Injecter la config antenne
  config['_ANTENNE'] = getAntenneConfig_();

  // Phase 9 : Mettre en cache (max 100KB par clé, on tronque si nécessaire)
  try {
    var jsonStr = JSON.stringify(config);
    if (jsonStr.length < 100000) {
      cache.put('config_data', jsonStr, CACHE_TTL);
    }
  } catch(e) { /* cache trop gros, pas grave */ }
  
  return config;
}

// Phase 9 : Fonction pour invalider le cache manuellement
function invalidateCache() {
  var cache = CacheService.getScriptCache();
  cache.removeAll(['config_data', 'antenne_config']);
  return { success: true, message: "Cache invalidé." };
}

// ============================================================================
//                                UTILITAIRES
// ============================================================================

function normalizeTel(tel) {
  if (!tel || tel === "Inconnu") return "";
  return String(tel).replace(/[^0-9]/g, ''); 
}

function formatDate(dateObj) {
  if (!dateObj || dateObj === "") return "";
  try { 
    return Utilities.formatDate(new Date(dateObj), "Europe/Paris", "dd/MM/yyyy"); 
  } catch (e) { 
    return ""; 
  }
}

function parseDateSecure(input) {
  if (!input) return new Date();
  if (input instanceof Date) return input;
  if (typeof input === 'string' && input.includes('/')) {
    let parts = input.split(' ')[0].split('/'); 
    if (parts.length === 3) {
      return new Date(parts[2], parseInt(parts[1]) - 1, parts[0]);
    }
  }
  return new Date(input);
}

// Phase 1 : Renommé de getJeuneStatus → getBenefStatus
function getBenefStatus(id) {
   const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_BENEFICIAIRES');
   const data = ws.getDataRange().getValues();
   for(let i=1; i<data.length; i++) {
     if(data[i][0] == id) {
       // Phase 2 : Colonnes mises à jour (M=12 statut_pres, N=13 lieu_vie, O=14 statut_inst)
       return { pres: data[i][12], inst: data[i][14], lieu: data[i][13] };
     }
   }
   return { pres:"", inst:"", lieu:"" };
}

// Rétrocompatibilité
function getJeuneStatus(id) { return getBenefStatus(id); }

// ============================================================================
//                          GESTION DES RAPPELS
//  Phase 6 : Rappels conditionnels par profil
// ============================================================================

function getRappelsAll() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsRappels = ss.getSheetByName('BDD_RAPPELS');
  const wsBenef = ss.getSheetByName('BDD_BENEFICIAIRES');
  
  const dataRappels = wsRappels.getDataRange().getValues();
  const dataBenef = wsBenef.getDataRange().getValues();
  
  let benefMap = {};
  for(let i=1; i<dataBenef.length; i++) {
    benefMap[dataBenef[i][0]] = { 
      nom: dataBenef[i][1] + ' ' + (dataBenef[i][2] || ''), 
      surnom: dataBenef[i][3],
      typeProfil: dataBenef[i][9] || '' // Phase 2 : col J = Type_Profil
    };
  }
  
  const today = new Date(); today.setHours(0,0,0,0);
  let results = [];
  
  for(let i=1; i<dataRappels.length; i++) {
    let statut = String(dataRappels[i][5] || "").toLowerCase().trim();
    if(statut.indexOf("faire") !== -1) {
      let echeance = parseDateSecure(dataRappels[i][3]);
      let idBenef = dataRappels[i][0];
      let b = benefMap[idBenef] || { nom: "Inconnu", surnom: "", typeProfil: "" };
      results.push({
        idRappel: dataRappels[i][1],
        idBenef: idBenef,
        nomBenef: b.nom,
        surnomBenef: b.surnom,
        typeRappel: dataRappels[i][6] || "",
        titre: dataRappels[i][6] || "Sans titre",
        details: dataRappels[i][3] || "",
        date: formatDate(echeance),
        isLate: echeance < today,
        typeProfil: b.typeProfil
      });
    }
  }
  
  results.sort(function(a, b) {
    if (a.isLate && !b.isLate) return -1;
    if (!a.isLate && b.isLate) return 1;
    return 0;
  });
  
  return results;
}

// Alias pour rétrocompatibilité
function getRappels() { return getRappelsAll(); }

function closeRappel(idRappel, auteurEmail, statusLabel) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsRappels = ss.getSheetByName('BDD_RAPPELS');
  const wsEvents = ss.getSheetByName('BDD_EVENTS');
  const data = wsRappels.getDataRange().getValues();
  const now = new Date();

  for(let i=1; i<data.length; i++) {
    if(data[i][1] == idRappel) {
      wsRappels.getRange(i+1, 6).setValue(statusLabel); 
      let idBenef = data[i][0];
      let titre = data[i][6];
      let details = data[i][3];
      let status = getBenefStatus(idBenef);
      wsEvents.appendRow([Utilities.getUuid(), idBenef, now, "Suivi Rappel", "Distance", status.lieu, "Rappel traité (" + statusLabel + ") : " + titre + " - " + details, status.pres, status.inst, "", "", "", status.lieu, auteurEmail]);
      return { success: true };
    }
  }
  return { success: false, message: "Rappel introuvable" };
}

// Phase 6 : Rappels conditionnels par profil
function createAutomaticReminders(idBenef, vulnsArray, eventType, details) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsRappels = ss.getSheetByName('BDD_RAPPELS');
  const wsRegles = ss.getSheetByName('BDD_REGLES_RAPPELS');
  const now = new Date();
  let rappelsToAdd = [];
  let rules = [];
  
  // Phase 6 : Récupérer le Type_Profil du bénéficiaire
  var benefProfil = '';
  try {
    var wsBenef = ss.getSheetByName('BDD_BENEFICIAIRES');
    var dataBenef = wsBenef.getDataRange().getValues();
    for (var b = 1; b < dataBenef.length; b++) {
      if (dataBenef[b][0] == idBenef) {
        benefProfil = String(dataBenef[b][9] || '').trim(); // Col J = Type_Profil
        break;
      }
    }
  } catch(e) {}
  
  if (wsRegles) {
    const dataRules = wsRegles.getDataRange().getValues();
    for (let i = 1; i < dataRules.length; i++) {
      if(dataRules[i][0] !== "") {
        // Phase 6 : Colonne Profils_Visibles dans les règles de rappels
        var profilsVisibles = String(dataRules[i][4] || '').trim();
        rules.push({ 
          trigger: String(dataRules[i][0]).toLowerCase(), 
          delay: dataRules[i][1], 
          title: dataRules[i][2], 
          msg: dataRules[i][3],
          profils: profilsVisibles
        });
      }
    }
  }

  let triggersToCheck = [];
  if (vulnsArray && vulnsArray.length > 0) vulnsArray.forEach(v => triggersToCheck.push(String(v).toLowerCase()));
  if (eventType) triggersToCheck.push(String(eventType).toLowerCase());

  rules.forEach(rule => {
    // Phase 6 : Vérifier si la règle s'applique au profil du bénéficiaire
    if (rule.profils) {
      var profilsList = rule.profils.split(',').map(function(p) { return p.trim(); });
      if (profilsList.length > 0 && !profilsList.includes(benefProfil)) return;
    }
    
    let match = triggersToCheck.some(t => t.includes(rule.trigger));
    if (match) {
      let ech = new Date(); ech.setDate(now.getDate() + parseInt(rule.delay));
      rappelsToAdd.push({ titre: rule.title, details: rule.msg, date: ech });
    }
  });

  rappelsToAdd.forEach(r => { 
    wsRappels.appendRow([idBenef, Utilities.getUuid(), "", r.details, r.date, "A faire", r.titre, now, "Automatique"]); 
  });
}

function saveManualReminder(form) {
  const wsRappels = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_RAPPELS');
  const now = new Date();
  wsRappels.appendRow([form.idBenef, Utilities.getUuid(), "", form.details, new Date(form.date), "A faire", form.titre, now, form.typeRappel || ""]);
  return { success: true, message: "Rappel ajouté." };
}

// ============================================================================
//      AJOUT COMMENTAIRE SUR RAPPEL
// ============================================================================

function addCommentToRappel(idRappel, comment) {
  const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_RAPPELS');
  const data = ws.getDataRange().getValues();
  const now = new Date();
  let jour = ("0" + now.getDate()).slice(-2);
  let mois = ("0" + (now.getMonth() + 1)).slice(-2);
  let dateTag = jour + "/" + mois;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == idRappel) {
      let currentDetails = String(data[i][3] || "");
      let newDetails = currentDetails + " // " + dateTag + " : '" + comment + "'";
      ws.getRange(i + 1, 4).setValue(newDetails);
      return { success: true, message: "Commentaire ajouté.", newDetails: newDetails };
    }
  }
  return { success: false, message: "Rappel introuvable." };
}

// ============================================================================
//      RECHERCHE BÉNÉFICIAIRE POUR RAPPEL
// ============================================================================

function searchBenefForRappel(query) {
  const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_BENEFICIAIRES');
  const data = ws.getDataRange().getValues();
  let results = [];
  let q = String(query).toLowerCase().trim();
  
  for (let i = 1; i < data.length; i++) {
    let prenom = String(data[i][1]).toLowerCase();
    let nom = String(data[i][2]).toLowerCase();
    let surnom = String(data[i][3]).toLowerCase();
    let telClean = normalizeTel(data[i][8]); // Col I = Telephone
    let qClean = normalizeTel(q);
    
    let matchText = prenom.includes(q) || nom.includes(q) || surnom.includes(q);
    let matchTel = (qClean.length > 3) && telClean.includes(qClean);
    
    if (matchText || matchTel) {
      results.push({
        id: data[i][0],
        nom: data[i][1] + ' ' + (data[i][2] || ''),
        surnom: data[i][3],
        tel: data[i][8]
      });
    }
    if (results.length >= 20) break;
  }
  return results;
}

// Alias rétrocompatibilité
function searchJeunesForRappel(query) { return searchBenefForRappel(query); }

// ============================================================================
//                          ENREGISTREMENT & DOUBLONS
//  Phase 1 : Renommage Jeune → Bénéficiaire
//  Phase 2 : Nouvelles colonnes Type_Profil, ID_Foyer, Role_Foyer, Genre via config
// ============================================================================

function checkDoublonTel(tel) {
  if (!tel) return [];
  const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_BENEFICIAIRES');
  const data = ws.getDataRange().getValues();
  let doublons = [];
  let inputTelClean = normalizeTel(tel);
  if (inputTelClean.length < 4) return []; 
  for (let i = 1; i < data.length; i++) {
    let dbTel = normalizeTel(data[i][8]); // Col I = Telephone
    if (dbTel === inputTelClean && dbTel !== "") {
      doublons.push({ nom: data[i][1] + ' ' + (data[i][2] || ''), surnom: data[i][3], id: data[i][0] });
    }
  }
  return doublons;
}

function saveBenefSmart(form, auteurEmail, modeForce) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsBenef = ss.getSheetByName('BDD_BENEFICIAIRES');
  const wsEvents = ss.getSheetByName('BDD_EVENTS');
  const wsModifs = ss.getSheetByName('BDD_MODIFS');
  const telPropre = "'" + form.tel; 
  const now = new Date();
  
  let finalIdFoyer = form.idFoyer || ''; // Phase 2 : ID_Foyer
  let historiqueMsg = "Création fiche";

  if (form.forceGroupId) { 
    finalIdFoyer = form.forceGroupId; 
    historiqueMsg = "Ajout au foyer (Enregistrement groupé)"; 
  } 
  else if (!modeForce) {
    let doublons = checkDoublonTel(form.tel);
    if (doublons.length > 0) return { status: "CONFLICT", doublons: doublons, message: "Ce numéro existe déjà." };
  }
  else if (modeForce === 'JOIN_GROUP') {
    const data = wsBenef.getDataRange().getValues();
    let inputTelClean = normalizeTel(form.tel);
    let groupFound = null;
    for (let i = 1; i < data.length; i++) {
      if (normalizeTel(data[i][8]) === inputTelClean) {
        if (data[i][10]) { // Col K = ID_Foyer
          groupFound = data[i][10]; 
        } else { 
          groupFound = 'FOY' + Utilities.getUuid().substring(0, 8); 
          wsBenef.getRange(i + 1, 11).setValue(groupFound); // Col K
        }
        break; 
      }
    }
    if (groupFound) finalIdFoyer = groupFound;
  }

  const idBenef = 'BEN' + Utilities.getUuid().substring(0, 10);
  const idEvent = Utilities.getUuid();
  const dateRencontre = form.dateRencontre ? new Date(form.dateRencontre) : now;
  let eventType = form.typeEvenement || "Première rencontre";
  let vulnString = Array.isArray(form.vulnerabilites) ? form.vulnerabilites.join(", ") : (form.vulnerabilites || "");
  let initStatutPres = form.statutPresence || "";
  let initStatutInst = form.statutInstitution || "";
  let langueStr = Array.isArray(form.langue) ? form.langue.join(", ") : (form.langue || "");
  let contactTypeStr = Array.isArray(form.typeContact) ? form.typeContact.join(", ") : (form.typeContact || "Non connu");
  let labelsStr = Array.isArray(form.labelsAutres) ? form.labelsAutres.join(", ") : (form.labelsAutres || "");

  // Phase 1+2 : Structure BDD_BENEFICIAIRES — 24 colonnes
  wsBenef.appendRow([
    idBenef,              // A - ID_Beneficiaire
    form.prenom,          // B - Prenom
    form.nom,             // C - Nom
    form.surnom,          // D - Surnom
    form.dateNaissance,   // E - Date_Naissance
    form.genre,           // F - Genre (Phase 2 : via config)
    form.nationalite,     // G - Nationalite
    langueStr,            // H - Langue
    telPropre,            // I - Telephone
    form.typeProfil || "",// J - Type_Profil (Phase 2 : NOUVEAU)
    finalIdFoyer,         // K - ID_Foyer (Phase 2 : NOUVEAU)
    form.roleFoyer || "", // L - Role_Foyer (Phase 2 : NOUVEAU)
    initStatutPres,       // M - Statut_Presence
    form.lieuVie,         // N - Lieu_Vie
    initStatutInst,       // O - Statut_Institution (Phase 1 : renommé)
    form.motifStatutInstitution || "", // P - Motif_Statut_Institution
    vulnString,           // Q - Vulnerabilites
    form.typeContact || "",// R - Type_Contact
    form.contactsAutres || "", // S - Contacts_Autres (JSON)
    form.notes || "",     // T - Notes
    now,                  // U - Date_Creation
    auteurEmail,          // V - Cree_Par
    now,                  // W - Date_Modification
    auteurEmail           // X - Modifie_Par
  ]);
  
  wsEvents.appendRow([idEvent, idBenef, dateRencontre, eventType, contactTypeStr, "", form.details || "Première rencontre", initStatutPres, initStatutInst, "", vulnString, "", form.lieuVie, auteurEmail]);
  
  createAutomaticReminders(idBenef, form.vulnerabilites, eventType, "");
  
  return { success: true, id: idBenef, groupId: finalIdFoyer, message: "Enregistrement réussi !" };
}

// Alias rétrocompatibilité
function saveJeuneSmart(form, auteurEmail, modeForce) { return saveBenefSmart(form, auteurEmail, modeForce); }

// ============================================================================
//                          MISE A JOUR & EVENEMENTS
//  Phase 1 : Renommage
//  Phase 2 : Nouvelles colonnes
// ============================================================================

function updateBenefProfile(form, auteurEmail) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsBenef = ss.getSheetByName('BDD_BENEFICIAIRES');
  const wsModifs = ss.getSheetByName('BDD_MODIFS');
  const data = wsBenef.getDataRange().getValues();
  const now = new Date();

  let rowIndex = -1;
  let current = {};

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == form.id) {
      rowIndex = i + 1;
      current = { 
        prenom: data[i][1], nom: data[i][2], surnom: data[i][3],
        dateNaissance: data[i][4], genre: data[i][5],
        nationalite: data[i][6], langue: data[i][7],
        tel: data[i][8], typeProfil: data[i][9],
        idFoyer: data[i][10], roleFoyer: data[i][11],
        statutPresence: data[i][12], lieuVie: data[i][13],
        statutInstitution: data[i][14], motifStatutInstitution: data[i][15],
        vulnerabilites: data[i][16], notes: data[i][19],
        labels: data[i][17] || ""
      };
      break;
    }
  }

  if (rowIndex === -1) return { success: false, message: "Bénéficiaire introuvable." };

  let changes = [];

  // Fonction helper pour les mises à jour
  function checkAndUpdate(fieldName, formValue, colIndex, currentValue) {
    if (formValue !== undefined && String(formValue) !== String(currentValue || '')) {
      wsModifs.appendRow([form.id, "Modification " + fieldName, currentValue, formValue, now, auteurEmail]);
      wsBenef.getRange(rowIndex, colIndex).setValue(formValue);
      changes.push(fieldName);
    }
  }

  // Champs texte simples
  if (form.prenom !== undefined) checkAndUpdate('Prénom', form.prenom, 2, current.prenom);
  if (form.nom !== undefined) checkAndUpdate('Nom', form.nom, 3, current.nom);
  if (form.surnom !== undefined) checkAndUpdate('Surnom', form.surnom, 4, current.surnom);
  if (form.tel !== undefined) {
    var newTel = "'" + form.tel;
    if (normalizeTel(form.tel) !== normalizeTel(current.tel)) {
      checkAndUpdate('Téléphone', newTel, 9, current.tel);
    }
  }
  if (form.dateNaissance !== undefined) checkAndUpdate('Date_Naissance', form.dateNaissance, 5, current.dateNaissance);
  if (form.genre !== undefined) checkAndUpdate('Genre', form.genre, 6, current.genre);
  if (form.nationalite !== undefined) checkAndUpdate('Nationalité', form.nationalite, 7, current.nationalite);
  
  // Phase 2 : Nouveaux champs
  if (form.typeProfil !== undefined) checkAndUpdate('Type_Profil', form.typeProfil, 10, current.typeProfil);
  if (form.idFoyer !== undefined) checkAndUpdate('ID_Foyer', form.idFoyer, 11, current.idFoyer);
  if (form.roleFoyer !== undefined) checkAndUpdate('Role_Foyer', form.roleFoyer, 12, current.roleFoyer);

  // Dropdowns
  if (form.lieu !== undefined) checkAndUpdate('Lieu_Vie', form.lieu, 14, current.lieuVie);
  if (form.statutInstitution !== undefined) checkAndUpdate('Statut_Institution', form.statutInstitution, 15, current.statutInstitution);
  if (form.statutPresence !== undefined) checkAndUpdate('Statut_Presence', form.statutPresence, 13, current.statutPresence);

  // Notes
  if (form.notesFixes !== undefined && form.notesFixes !== (current.notes || '')) {
    wsModifs.appendRow([form.id, "Modification Notes", current.notes, form.notesFixes, now, auteurEmail]);
    wsBenef.getRange(rowIndex, 20).setValue(form.notesFixes); // Col T
    changes.push("Notes");
  }

  // Langue (multi-select)
  if (form.langue !== undefined) {
    var newLangStr = Array.isArray(form.langue) ? form.langue.join(", ") : (form.langue || "");
    if (newLangStr !== String(current.langue || '')) {
      checkAndUpdate('Langue', newLangStr, 8, current.langue);
    }
  }

  // Labels
  if (form.labelsAutres !== undefined) {
    var newLabelsStr = Array.isArray(form.labelsAutres) ? form.labelsAutres.join(", ") : (form.labelsAutres || "");
    if (newLabelsStr !== current.labels) {
      wsModifs.appendRow([form.id, "Modification Labels", current.labels, newLabelsStr, now, auteurEmail]);
      wsBenef.getRange(rowIndex, 18).setValue(newLabelsStr); // Col R = Type_Contact → non, col S?
      changes.push("Labels");
    }
  }

  // Vulnérabilités
  if (form.vulnerabilites !== undefined) {
    let newVulnStr = Array.isArray(form.vulnerabilites) ? form.vulnerabilites.join(", ") : (form.vulnerabilites || "");
    let currentVulns = String(data[rowIndex - 1][16] || "");
    if (newVulnStr !== currentVulns) {
      wsModifs.appendRow([form.id, "Modification Vulnérabilités", currentVulns, newVulnStr, now, auteurEmail]);
      wsBenef.getRange(rowIndex, 17).setValue(newVulnStr); // Col Q
      changes.push("Vulnérabilités");
    }
  }

  // Mettre à jour Date_Modification et Modifie_Par
  if (changes.length > 0) {
    wsBenef.getRange(rowIndex, 23).setValue(now);  // Col W
    wsBenef.getRange(rowIndex, 24).setValue(auteurEmail); // Col X
  }

  if (changes.length === 0) return { success: false, message: "Aucune modification détectée." };
  return { success: true, message: "Modifications enregistrées : " + changes.join(", ") };
}

// Alias rétrocompatibilité
function updateJeuneProfile(form, auteurEmail) { return updateBenefProfile(form, auteurEmail); }

function saveNewEvent(form, auteurEmail) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsEvents = ss.getSheetByName('BDD_EVENTS');
  const wsBenef = ss.getSheetByName('BDD_BENEFICIAIRES');
  const dataBenef = wsBenef.getDataRange().getValues();
  const now = new Date(); 
  const dateEvent = form.date ? new Date(form.date) : now; 
  const idEvent = Utilities.getUuid();
  
  let rowIndex = -1; 
  let current = {};
  for (let i = 1; i < dataBenef.length; i++) { 
    if (dataBenef[i][0] == form.idBenef) { 
      rowIndex = i + 1; 
      current = { 
        pres: dataBenef[i][12],      // M = Statut_Presence
        inst: dataBenef[i][14],      // O = Statut_Institution
        motifInst: dataBenef[i][15], // P = Motif_Statut_Institution
        lieuVie: dataBenef[i][13],   // N = Lieu_Vie
        vulns: dataBenef[i][16]      // Q = Vulnerabilites
      }; 
      break; 
    } 
  }
  if (rowIndex === -1) return { success: false, message: "Bénéficiaire introuvable." };
  
  // Mise à jour Date_Modification
  wsBenef.getRange(rowIndex, 23).setValue(dateEvent); // Col W
  wsBenef.getRange(rowIndex, 24).setValue(auteurEmail); // Col X
  
  let eventTypeLower = String(form.type).toLowerCase();
  
  let finalLieuVie = current.lieuVie; 
  if (form.lieuVie && form.lieuVie !== "" && form.lieuVie !== current.lieuVie) { 
    wsBenef.getRange(rowIndex, 14).setValue(form.lieuVie); // Col N
    finalLieuVie = form.lieuVie; 
  }
  
  let finalStatutPres = current.pres; 
  if (form.statutPres && form.statutPres !== "" && form.statutPres !== current.pres) { 
    wsBenef.getRange(rowIndex, 13).setValue(form.statutPres); // Col M
    finalStatutPres = form.statutPres; 
  }
  
  let finalStatutInst = current.inst; 
  if (form.statutInstitution && form.statutInstitution !== "" && form.statutInstitution !== current.inst) { 
    wsBenef.getRange(rowIndex, 15).setValue(form.statutInstitution); // Col O
    finalStatutInst = form.statutInstitution; 
  }

  let finalMotifInst = current.motifInst;
  if (form.motifStatutInstitution && form.motifStatutInstitution !== "") {
    wsBenef.getRange(rowIndex, 16).setValue(form.motifStatutInstitution); // Col P
    finalMotifInst = form.motifStatutInstitution;
  }
  
  let newVulnString = ""; 
  if (form.vulns && form.vulns.length > 0) { 
    let currentVList = current.vulns ? String(current.vulns).split(',').map(s => s.trim()) : []; 
    let addedVulns = []; 
    form.vulns.forEach(v => { 
      if (!currentVList.includes(v)) { 
        currentVList.push(v); 
        addedVulns.push(v); 
      } 
    }); 
    if (addedVulns.length > 0) { 
      newVulnString = currentVList.join(", "); 
      wsBenef.getRange(rowIndex, 17).setValue(newVulnString); // Col Q
    } 
  }
  
  let fullDetails = form.note || ""; 
  let materielStr = Array.isArray(form.materiel) ? form.materiel.join(", ") : (form.materiel || "");
  let contactStr = Array.isArray(form.typeContact) ? form.typeContact.join(", ") : (form.typeContact || "Physique");
  let motifStr = Array.isArray(form.motifStatutInstitution) ? form.motifStatutInstitution.join(", ") : (form.motifStatutInstitution || "");
  
  wsEvents.appendRow([idEvent, form.idBenef, dateEvent, form.type, contactStr, form.lieuEvent || "", fullDetails, finalStatutPres, finalStatutInst, motifStr, newVulnString, materielStr, finalLieuVie, auteurEmail]);
  
  createAutomaticReminders(form.idBenef, form.vulns, form.type, form.note);
  
  return { success: true, message: "Événement enregistré." };
}

function getEventData(idEvent) {
  const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_EVENTS');
  const data = ws.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (String(data[i][0]) === String(idEvent)) {
      var rawDate = data[i][2];
      var dateRawStr = "";
      try { dateRawStr = (rawDate instanceof Date) ? rawDate.toISOString() : String(rawDate || ""); } catch(e) { dateRawStr = ""; }
      return { id: data[i][0], dateRaw: dateRawStr, type: data[i][3], contact: data[i][4], lieu: data[i][5], details: data[i][6] };
    }
  }
  return null;
}

function saveEditedEvent(form, auteurEmail) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsEvents = ss.getSheetByName('BDD_EVENTS');
  const wsModifs = ss.getSheetByName('BDD_MODIFS');
  const data = wsEvents.getDataRange().getValues();
  
  for (let i=1; i<data.length; i++) {
    if (String(data[i][0]) === String(form.idEvent)) {
      wsModifs.appendRow([data[i][1], "Modif Event (" + data[i][3] + ")", "", "Modifié par user", new Date(), auteurEmail]);
      wsEvents.getRange(i+1, 3).setValue(new Date(form.date));
      wsEvents.getRange(i+1, 4).setValue(form.type);
      wsEvents.getRange(i+1, 7).setValue(form.details);
      return { success: true, message: "Événement mis à jour" };
    }
  }
  return { success: false, message: "Événement introuvable" };
}

function saveMassEvent(form, auteurEmail) {
  if (!form.ids || form.ids.length === 0) return { success: false, message: "Aucun bénéficiaire sélectionné." };
  let successCount = 0;
  form.ids.forEach(id => {
    let singleForm = { 
      idBenef: id, type: form.type, date: form.date, lieuVie: form.lieuVie, lieuEvent: form.lieuEvent, 
      statutInstitution: form.statutInstitution, statutPres: form.statutPres, 
      motifStatutInstitution: form.motifStatutInstitution,
      materiel: form.materiel, typeContact: form.typeContact, vulns: form.vulns, note: form.note 
    };
    try { saveNewEvent(singleForm, auteurEmail); successCount++; } catch(e) { console.error(e); }
  });
  return { success: true, message: successCount + " bénéficiaires mis à jour avec succès." };
}

// ============================================================================
//              ÉVICTION RAPIDE PAR LIEU DE VIE
// ============================================================================

function getBenefByLieuVie(lieuVie) {
  const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_BENEFICIAIRES');
  const data = ws.getDataRange().getValues();
  let results = [];
  
  for (let i = 1; i < data.length; i++) {
    let statutPres = String(data[i][12] || "").toLowerCase();
    let lieu = String(data[i][13] || "");
    
    if (lieu === lieuVie && (statutPres.includes("présent") || statutPres.includes("foyer") || statutPres.includes("hospitalisé"))) {
      results.push({
        id: data[i][0],
        nom: data[i][1] + ' ' + (data[i][2] || ''),
        surnom: data[i][3],
        age: data[i][4], // Date_Naissance
        tel: data[i][8],
        lieu: data[i][13]
      });
    }
  }
  return results;
}

// Alias rétrocompatibilité
function getJeunesByLieuVie(lieuVie) { return getBenefByLieuVie(lieuVie); }

// ============================================================================
//              RECHERCHE AVANCÉE AVEC FILTRES
//  Phase 1 : Renommage
//  Phase 2 : Filtrage par Type_Profil
// ============================================================================

function searchBenefAdvanced(query, filters) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsBenef = ss.getSheetByName('BDD_BENEFICIAIRES');
  const wsModifs = ss.getSheetByName('BDD_MODIFS');
  const dataBenef = wsBenef.getDataRange().getValues();
  const dataModifs = wsModifs.getDataRange().getValues();
  
  // Phase 2 : Compteur par foyer au lieu de groupes
  let foyerCounts = {};
  for(let i=1; i<dataBenef.length; i++) {
    let fid = dataBenef[i][10]; // Col K = ID_Foyer
    if(fid) foyerCounts[fid] = (foyerCounts[fid] || 0) + 1;
  }

  let results = [];
  let foundIDs = new Set();
  let q = query ? String(query).toLowerCase().trim() : "";
  let qClean = normalizeTel(q);
  
  let filterPresence = (filters && filters.statutPresence) ? filters.statutPresence : "";
  let filterInst = (filters && filters.statutInstitution) ? filters.statutInstitution : "";
  let filterVuln = (filters && filters.vulnerabilite) ? filters.vulnerabilite.toLowerCase() : "";
  let filterLieu = (filters && filters.lieuVie) ? filters.lieuVie : "";
  let filterProfil = (filters && filters.typeProfil) ? filters.typeProfil : ""; // Phase 2
  let showAllActifs = (filters && filters.showAllActifs) ? true : false;
  let showAll = (filters && filters.showAll) ? true : false;
  
  for (let i = 1; i < dataBenef.length; i++) {
    let row = dataBenef[i];
    let id = row[0];
    let prenom = String(row[1] || "").toLowerCase();
    let nom = String(row[2] || "").toLowerCase();
    let surnom = String(row[3] || "").toLowerCase();
    let telClean = normalizeTel(row[8]);
    let statutPres = String(row[12] || "");
    let statutInst = String(row[14] || "");
    let vulns = String(row[16] || "").toLowerCase();
    let lieuVie = String(row[13] || "");
    let typeProfil = String(row[9] || ""); // Phase 2
    
    // Filtres
    if (filterPresence && statutPres !== filterPresence) continue;
    if (filterInst && statutInst !== filterInst) continue;
    if (filterVuln && !vulns.includes(filterVuln)) continue;
    if (filterLieu && lieuVie !== filterLieu) continue;
    if (filterProfil && typeProfil !== filterProfil) continue; // Phase 2
    
    if (!showAll && !showAllActifs) {
      let matchText = q ? (prenom.includes(q) || nom.includes(q) || surnom.includes(q) || String(id).toLowerCase().includes(q)) : false;
      let matchTel = (qClean.length > 3) ? telClean.includes(qClean) : false;
      let hasFilter = filterPresence || filterInst || filterVuln || filterLieu || filterProfil;
      if (!matchText && !matchTel && !hasFilter) continue;
    }
    
    if (showAllActifs) {
      let presLower = statutPres.toLowerCase();
      if (!presLower.includes("présent") && !presLower.includes("foyer") && !presLower.includes("hospitalisé")) continue;
    }
    
    if (!foundIDs.has(id)) {
      foundIDs.add(id);
      let fid = row[10];
      let isFoyer = fid && foyerCounts[fid] && foyerCounts[fid] > 1;
      results.push({
        id: id,
        nom: row[1] + ' ' + (row[2] || ''),
        surnom: row[3],
        age: row[4], // Date_Naissance
        tel: row[8],
        nationalite: row[6],
        statut_institution: statutInst,
        lieu: lieuVie,
        is_history: false,
        statut_presence: statutPres,
        is_grouped: isFoyer,
        genre: String(row[5] || ""),
        vulns: vulns,
        typeProfil: typeProfil // Phase 2
      });
    }
    if (results.length >= 200) break;
  }
  
  // Recherche dans l'historique des modifications
  if (q && results.length < 200) {
    for (let i = 1; i < dataModifs.length; i++) {
      let oldVal = String(dataModifs[i][4] || "").toLowerCase();
      let newVal = String(dataModifs[i][5] || "").toLowerCase();
      if (oldVal.includes(q) || newVal.includes(q)) {
        let idBenef = dataModifs[i][1];
        if (!foundIDs.has(idBenef)) {
          foundIDs.add(idBenef);
          for (let k = 1; k < dataBenef.length; k++) {
            if (dataBenef[k][0] === idBenef) {
              results.push({ id: dataBenef[k][0], nom: dataBenef[k][1] + ' ' + (dataBenef[k][2] || ''), surnom: dataBenef[k][3], age: dataBenef[k][4], nationalite: dataBenef[k][6], statut_institution: String(dataBenef[k][14] || ""), lieu: String(dataBenef[k][13] || ""), is_history: true, statut_presence: String(dataBenef[k][12] || ""), is_grouped: false, genre: "", vulns: "", typeProfil: String(dataBenef[k][9] || "") });
              break;
            }
          }
        }
      }
      if (results.length >= 200) break;
    }
  }
  return results;
}

function searchBenef(query) { return searchBenefAdvanced(query, null); }

// Alias rétrocompatibilité
function searchJeunesAdvanced(query, filters) { return searchBenefAdvanced(query, filters); }
function searchJeunes(query) { return searchBenef(query); }

// ============================================================================
//              RAPPELS EN RETARD PAR BÉNÉFICIAIRE
// ============================================================================

function getRappelsEnRetardParBenef() {
  const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_RAPPELS');
  const data = ws.getDataRange().getValues();
  const today = new Date();
  today.setHours(0,0,0,0);
  
  let retardMap = {};
  for (let i = 1; i < data.length; i++) {
    let statut = String(data[i][5] || "").toLowerCase().trim();
    if (statut.indexOf("faire") !== -1) {
      let echeance = parseDateSecure(data[i][4]);
      if (echeance < today) {
        let idBenef = data[i][0];
        retardMap[idBenef] = (retardMap[idBenef] || 0) + 1;
      }
    }
  }
  return retardMap;
}

// Alias
function getRappelsEnRetardParJeune() { return getRappelsEnRetardParBenef(); }

// ============================================================================
//                          FICHE BÉNÉFICIAIRE & FOYER
//  Phase 1 : Renommage
//  Phase 2 : Nouvelles colonnes
//  Phase 10 : Contexte foyer dans la fiche
// ============================================================================

function getFicheBenef(idBenef) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const dataBenef = ss.getSheetByName('BDD_BENEFICIAIRES').getDataRange().getValues();
  const dataEvents = ss.getSheetByName('BDD_EVENTS').getDataRange().getValues();
  const dataModifs = ss.getSheetByName('BDD_MODIFS').getDataRange().getValues();
  
  let infoBenef = null;
  for (let i = 1; i < dataBenef.length; i++) {
    if (dataBenef[i][0] == idBenef) {
      infoBenef = {
        id: dataBenef[i][0],
        prenom: dataBenef[i][1],
        nom: dataBenef[i][2] || '',
        surnom: dataBenef[i][3],
        date_naissance: dataBenef[i][4],
        genre: dataBenef[i][5],
        nationalite: dataBenef[i][6],
        langue: dataBenef[i][7],
        tel: dataBenef[i][8],
        type_profil: dataBenef[i][9] || '',       // Phase 2
        id_foyer: dataBenef[i][10] || '',          // Phase 2
        role_foyer: dataBenef[i][11] || '',        // Phase 2
        statut_presence: dataBenef[i][12],
        lieu_vie: dataBenef[i][13],
        statut_institution: dataBenef[i][14],
        motif_statut_institution: dataBenef[i][15],
        vulnerabilites: dataBenef[i][16],
        type_contact: dataBenef[i][17],
        contacts_autres: dataBenef[i][18],
        notes: dataBenef[i][19],
        date_creation: dataBenef[i][20],
        cree_par: dataBenef[i][21]
      };
      break;
    }
  }
  
  if (!infoBenef) return { success: false, message: "Bénéficiaire introuvable." };
  
  // Historique combiné (events + modifications)
  let history = [];
  
  for (let i = 1; i < dataEvents.length; i++) {
    if (dataEvents[i][1] == idBenef) {
      var rawDate = dataEvents[i][2];
      var dateRawStr = "";
      try { dateRawStr = (rawDate instanceof Date) ? rawDate.toISOString() : String(rawDate || ""); } catch(e) {}
      history.push({
        id: dataEvents[i][0],
        date: formatDate(rawDate),
        dateRaw: dateRawStr,
        type: dataEvents[i][3],
        contact: dataEvents[i][4],
        lieu: dataEvents[i][5],
        details: dataEvents[i][6],
        auteur: dataEvents[i][13]
      });
    }
  }
  
  for (let i = 1; i < dataModifs.length; i++) {
    if (dataModifs[i][0] == idBenef || dataModifs[i][1] == idBenef) {
      var mDate = dataModifs[i][4] || dataModifs[i][2];
      var dateRawStr2 = "";
      try { dateRawStr2 = (mDate instanceof Date) ? mDate.toISOString() : String(mDate || ""); } catch(e) {}
      history.push({
        date: formatDate(mDate),
        dateRaw: dateRawStr2,
        type: "Modification",
        details: (dataModifs[i][1] || dataModifs[i][3]) + " : " + (dataModifs[i][4] || dataModifs[i][2] || "") + " → " + (dataModifs[i][5] || dataModifs[i][3] || ""),
        auteur: dataModifs[i][5] || dataModifs[i][6] || ""
      });
    }
  }
  
  history.sort(function(a, b) {
    try {
      var dateA = a.dateRaw ? new Date(a.dateRaw) : new Date(0);
      var dateB = b.dateRaw ? new Date(b.dateRaw) : new Date(0);
      return dateB.getTime() - dateA.getTime();
    } catch(e) { return 0; }
  });

  // Dernier motif statut institution depuis les events
  var lastMotifInst = "";
  for (var ev = dataEvents.length - 1; ev >= 1; ev--) {
    if (dataEvents[ev][1] == idBenef) {
      var motif = String(dataEvents[ev][9] || "").trim();
      if (motif) { lastMotifInst = motif; break; }
    }
  }
  infoBenef.motif_statut_institution = infoBenef.motif_statut_institution || lastMotifInst;

  return { success: true, benef: infoBenef, history: history };
}

// Alias rétrocompatibilité
function getFicheJeune(idJeune) {
  var result = getFicheBenef(idJeune);
  if (result.success) {
    // Mapper pour compatibilité avec le code existant
    result.jeune = result.benef;
  }
  return result;
}

// Phase 10 : Membres du foyer
function getFoyerMembers(foyerId, excludeId) {
  const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_BENEFICIAIRES');
  const data = ws.getDataRange().getValues();
  let members = [];
  
  for(let i=1; i<data.length; i++) {
    if(data[i][10] === foyerId && data[i][0] !== excludeId) {
      members.push({
        id: data[i][0],
        prenom: data[i][1],
        nom: data[i][2],
        surnom: data[i][3],
        age: data[i][4],
        genre: data[i][5],
        typeProfil: data[i][9],
        roleFoyer: data[i][11]
      });
    }
  }
  return members;
}

// Alias rétrocompatibilité (ancien système de groupes)
function getGroupMembers(groupId, excludeId) { return getFoyerMembers(groupId, excludeId); }

// ============================================================================
//    LIAISON FOYER (Phase 2 + Phase 10)
// ============================================================================

function searchBenefForFoyerLink(query) {
  const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_BENEFICIAIRES');
  const data = ws.getDataRange().getValues();
  let results = [];
  let q = String(query).toLowerCase().trim();
  let qClean = normalizeTel(q);
  
  for (let i = 1; i < data.length; i++) {
    let prenom = String(data[i][1]).toLowerCase();
    let nom = String(data[i][2]).toLowerCase();
    let surnom = String(data[i][3]).toLowerCase();
    let telClean = normalizeTel(data[i][8]);
    
    let matchText = prenom.includes(q) || nom.includes(q) || surnom.includes(q);
    let matchTel = (qClean.length > 3) && telClean.includes(qClean);
    
    if (matchText || matchTel) {
      results.push({
        id: data[i][0],
        nom: data[i][1] + ' ' + (data[i][2] || ''),
        surnom: data[i][3],
        tel: data[i][8],
        age: data[i][4],
        lieu: data[i][13],
        id_foyer: data[i][10] || "",
        typeProfil: data[i][9] || "",
        roleFoyer: data[i][11] || ""
      });
    }
    if (results.length >= 30) break;
  }
  return results;
}

// Alias
function searchJeunesForGroupLink(query) { return searchBenefForFoyerLink(query); }

function linkBenefToFoyer(sourceBenefId, targetBenefIds) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsBenef = ss.getSheetByName('BDD_BENEFICIAIRES');
  const data = wsBenef.getDataRange().getValues();
  const now = new Date();
  
  if (!targetBenefIds || targetBenefIds.length === 0) {
    return { success: false, message: "Aucun bénéficiaire cible sélectionné." };
  }
  
  let foyerId = null;
  let sourceRowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === sourceBenefId) {
      sourceRowIndex = i + 1;
      foyerId = data[i][10] || null; // Col K = ID_Foyer
      break;
    }
  }
  
  if (sourceRowIndex === -1) return { success: false, message: "Bénéficiaire source introuvable." };
  
  if (!foyerId) {
    foyerId = 'FOY' + Utilities.getUuid().substring(0, 8);
    wsBenef.getRange(sourceRowIndex, 11).setValue(foyerId); // Col K
  }
  
  let linkedCount = 0;
  for (let i = 1; i < data.length; i++) {
    if (targetBenefIds.includes(data[i][0])) {
      if (data[i][10] !== foyerId) {
        wsBenef.getRange(i + 1, 11).setValue(foyerId); // Col K
        linkedCount++;
      }
    }
  }
  
  return { success: true, message: linkedCount + " bénéficiaire(s) lié(s) au foyer.", foyerId: foyerId };
}

// Alias
function linkJeunesToGroup(sourceId, targetIds) { return linkBenefToFoyer(sourceId, targetIds); }

// Phase 10 : Formulaire simplifié ajout membre foyer
function addFoyerMember(form, auteurEmail) {
  // Pré-remplir l'ID_Foyer et créer un nouveau bénéficiaire lié
  form.idFoyer = form.foyerId;
  form.forceGroupId = form.foyerId;
  return saveBenefSmart(form, auteurEmail, null);
}

// ============================================================================
//              PLAIDOYER & STATS
// ============================================================================

function getStatistics(startStr, endStr) {
  return getStatisticsDetailed(startStr, endStr);
}

function getStatisticsDetailed(startStr, endStr) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const wsBenef = ss.getSheetByName('BDD_BENEFICIAIRES');
    const wsEvents = ss.getSheetByName('BDD_EVENTS');
    const wsConfig = ss.getSheetByName('BDD_CONFIG');
    
    let start = parseDateSecure(startStr);
    let end = parseDateSecure(endStr);
    end.setHours(23, 59, 59); 

    const dataBenef = wsBenef.getDataRange().getValues();
    const dataEvents = wsEvents.getDataRange().getValues();

    const dataConfig = wsConfig.getDataRange().getValues();
    let configLists = {};
    if (dataConfig.length > 0) {
      const headers = dataConfig[0];
      for (let col = 0; col < headers.length; col++) {
        let key = headers[col];
        if (!key) continue;
        let values = [];
        for (let row = 1; row < dataConfig.length; row++) {
          if (dataConfig[row][col] !== "") values.push(String(dataConfig[row][col]));
        }
        configLists[key] = values;
      }
    }

    let presents = 0, nouveaux = 0;
    let countNationalites = {}, countLangues = {}, countStatutPresence = {};
    let countStatutInstitution = {}, countVulnerabilitesTags = {}, countLieuxVie = {};
    let countTypesEvenements = {}, countMotifStatutInstitution = {}, countMateriel = {}, countTypesContact = {};
    // Phase 2 : Stats par profil, genre, foyer
    let countTypesProfils = {}, countGenres = {}, countRolesFoyer = {};
    let foyerSet = new Set();

    function incrementMulti(counterObj, rawValue) {
      if (!rawValue) return;
      String(rawValue).split(',').forEach(function(part) {
        var trimmed = part.trim();
        if (trimmed !== "" && !trimmed.startsWith('---')) counterObj[trimmed] = (counterObj[trimmed] || 0) + 1;
      });
    }

    function incrementSingle(counterObj, rawValue) {
      if (!rawValue) return;
      var trimmed = String(rawValue).trim();
      if (trimmed !== "" && !trimmed.startsWith('---')) counterObj[trimmed] = (counterObj[trimmed] || 0) + 1;
    }

    // Stats bénéficiaires (stock)
    let total_actifs = 0;
    let jeunes_with_events = new Set();
    
    for (let i = 1; i < dataBenef.length; i++) {
      let presLower = String(dataBenef[i][12] || "").toLowerCase();
      let isActif = presLower.includes("présent") || presLower.includes("foyer") || presLower.includes("hospitalisé");
      if (isActif) { presents++; total_actifs++; }
      
      // Création dans la période ?
      let dateCreation = dataBenef[i][20];
      if (dateCreation) {
        try {
          let dc = new Date(dateCreation);
          if (dc >= start && dc <= end) nouveaux++;
        } catch(e) {}
      }
      
      incrementSingle(countNationalites, dataBenef[i][6]);
      incrementMulti(countLangues, dataBenef[i][7]);
      incrementSingle(countStatutPresence, dataBenef[i][12]);
      incrementSingle(countStatutInstitution, dataBenef[i][14]);
      incrementSingle(countLieuxVie, dataBenef[i][13]);
      incrementMulti(countVulnerabilitesTags, dataBenef[i][16]);
      // Phase 2
      incrementSingle(countTypesProfils, dataBenef[i][9]);
      incrementSingle(countGenres, dataBenef[i][5]);
      incrementSingle(countRolesFoyer, dataBenef[i][11]);
      if (dataBenef[i][10]) foyerSet.add(dataBenef[i][10]);
    }

    // Stats événements (période)
    let total_events = 0;
    for (let i = 1; i < dataEvents.length; i++) {
      let evDate = null;
      try { evDate = new Date(dataEvents[i][2]); } catch(e) { continue; }
      if (!evDate || isNaN(evDate.getTime())) continue;
      if (evDate < start || evDate > end) continue;
      
      total_events++;
      jeunes_with_events.add(dataEvents[i][1]);
      incrementSingle(countTypesEvenements, dataEvents[i][3]);
      incrementMulti(countTypesContact, dataEvents[i][4]);
      incrementSingle(countMotifStatutInstitution, dataEvents[i][9]);
      incrementMulti(countMateriel, dataEvents[i][11]);
    }

    let stats = {
      periode_debut: startStr, periode_fin: endStr,
      total_actifs: total_actifs,
      presents: presents, nouveaux: nouveaux,
      total_events: total_events,
      total_foyers: foyerSet.size, // Phase 2
      nationalites: countNationalites, langues: countLangues,
      statut_presence: countStatutPresence,
      statut_institution: countStatutInstitution, // Phase 1 : renommé
      lieux_vie: countLieuxVie, vulnerabilites_tags: countVulnerabilitesTags,
      types_evenements: countTypesEvenements, motif_statut_institution: countMotifStatutInstitution,
      materiel: countMateriel, types_contact: countTypesContact,
      // Phase 2
      types_profils: countTypesProfils, genres: countGenres, roles_foyer: countRolesFoyer,
      config_lists: configLists,
      jeunes_with_events: jeunes_with_events.size
    };

    return stats;

  } catch(e) {
    throw new Error("Erreur calcul stats: " + e.toString());
  }
}

// ============================================================================
//     Phase 4 : RAPPORT IP / SIGNALEMENT DYNAMIQUE PAR TYPE DE PROFIL
// ============================================================================

function generateReportText(idBenef) {
  const info = getFicheBenef(idBenef); 
  if (!info.success) return "Erreur: Bénéficiaire introuvable.";
  
  const b = info.benef;
  const hist = info.history;
  const antenneConfig = getAntenneConfig_();
  
  // Phase 4 : Vérifier si le rapport IP est actif pour cette antenne
  if (antenneConfig.RAPPORT_IP_ACTIF && antenneConfig.RAPPORT_IP_ACTIF.toUpperCase() !== 'OUI') {
    return "Le rapport IP n'est pas activé pour cette antenne.";
  }
  
  const aujourdhui = formatDate(new Date());
  const nomAntenne = antenneConfig.NOM_ANTENNE || 'Utopia 56';
  const contactSignalement = antenneConfig.CONTACT_SIGNALEMENT || '[À compléter]';
  
  let chronoHist = [...hist].reverse();
  let premierEvt = chronoHist.length > 0 ? chronoHist[0].date : "Inconnue";
  
  // Compteurs d'événements
  let nbEvictions = 0, nbViolencesPolice = 0, nbViolencesTiers = 0;
  let nbTentatives = 0, nbUrgencesSante = 0, nbRefus = 0;
  
  chronoHist.forEach(function(evt) {
    let typeLower = String(evt.type).toLowerCase();
    if (typeLower.includes("éviction") || typeLower.includes("expulsion")) nbEvictions++;
    if (typeLower.includes("violence") && typeLower.includes("polic")) nbViolencesPolice++;
    if (typeLower.includes("violence") && (typeLower.includes("tiers") || typeLower.includes("passeur"))) nbViolencesTiers++;
    if (typeLower.includes("traversée") || typeLower.includes("départ")) nbTentatives++;
    if (typeLower.includes("santé") || typeLower.includes("urgence") || typeLower.includes("hospitalisation")) nbUrgencesSante++;
    if (typeLower.includes("refus")) nbRefus++;
  });
  
  // Phase 4 : Adapter le rapport selon le Type_Profil
  var typeProfil = b.type_profil || '';
  var isMNA = typeProfil.toLowerCase().includes('mna');
  var isFamille = typeProfil.toLowerCase().includes('famille') || typeProfil.toLowerCase().includes('couple');
  var isMineur = isMNA; // On peut aussi vérifier l'âge
  
  let txt = "";
  txt += "=========================================================================\n";
  
  if (isMNA) {
    txt += "              INFORMATION PRÉOCCUPANTE (IP)\n";
    txt += "              SIGNALEMENT - MINEUR NON ACCOMPAGNÉ\n";
  } else if (isFamille) {
    txt += "              SIGNALEMENT\n";
    txt += "              FAMILLE EN SITUATION DE DANGER\n";
  } else {
    txt += "              SIGNALEMENT\n";
    txt += "              PERSONNE EN SITUATION DE VULNÉRABILITÉ\n";
  }
  txt += "=========================================================================\n\n";
  
  txt += "Émetteur : Association Utopia 56 - Antenne " + nomAntenne + "\n";
  txt += "Date du signalement : " + aujourdhui + "\n";
  
  if (isMNA) {
    txt += "Destinataire : Cellule de Recueil des Informations Préoccupantes (CRIP)\n";
    txt += "               Procureur de la République (le cas échéant)\n\n";
  } else {
    txt += "Destinataire : " + contactSignalement + "\n\n";
  }
  
  if (isMNA) {
    txt += "-------------------------------------------------------------------------\n";
    txt += "FONDEMENT JURIDIQUE\n";
    txt += "-------------------------------------------------------------------------\n";
    txt += "Le présent signalement est effectué sur le fondement des articles :\n";
    txt += "- L.112-3 du Code de l'Action Sociale et des Familles (CASF)\n";
    txt += "- L.223-2 du CASF relatif à la mise à l'abri immédiate des mineurs\n";
    txt += "  non accompagnés.\n";
    txt += "- Article 375 du Code Civil relatif à l'assistance éducative.\n";
    txt += "- Convention Internationale des Droits de l'Enfant (CIDE)\n\n";
  }
  
  txt += "=========================================================================\n";
  txt += "1. IDENTIFICATION\n";
  txt += "=========================================================================\n\n";
  txt += "Prénom / Nom     : " + (b.prenom || "") + " " + (b.nom || "(inconnu)") + "\n";
  if (b.surnom) txt += "Surnom           : " + b.surnom + "\n";
  txt += "Date de naissance: " + (b.date_naissance ? formatDate(b.date_naissance) : "Non connue") + "\n";
  txt += "Genre            : " + (b.genre || "Non connu") + "\n";
  txt += "Nationalité      : " + (b.nationalite || "Non connue") + "\n";
  txt += "Type de profil   : " + (b.type_profil || "Non précisé") + "\n";
  txt += "Téléphone        : " + (b.tel || "Non connu") + "\n";
  txt += "Premier contact  : " + premierEvt + "\n\n";
  
  // Phase 10 : Si foyer, ajouter les membres
  if (b.id_foyer) {
    var membres = getFoyerMembers(b.id_foyer, b.id);
    if (membres.length > 0) {
      txt += "COMPOSITION DU FOYER :\n";
      membres.forEach(function(m) {
        txt += "  - " + m.prenom + " " + (m.nom || '') + " (" + (m.roleFoyer || 'membre') + ", " + (m.genre || '?') + ")\n";
      });
      txt += "\n";
    }
  }
  
  txt += "=========================================================================\n";
  txt += "2. SITUATION ACTUELLE\n";
  txt += "=========================================================================\n\n";
  txt += "Lieu de vie       : " + (b.lieu_vie || "Inconnu") + "\n";
  txt += "Statut présence   : " + (b.statut_presence || "Non renseigné") + "\n";
  txt += "Statut institution: " + (b.statut_institution || "Aucun") + "\n";
  if (b.vulnerabilites) txt += "Vulnérabilités    : " + b.vulnerabilites + "\n";
  txt += "\n";
  
  txt += "INDICATEURS DE DANGER :\n";
  if (nbEvictions > 0) txt += "  ⚠ " + nbEvictions + " éviction(s) / expulsion(s) documentée(s)\n";
  if (nbViolencesPolice > 0) txt += "  ⚠ " + nbViolencesPolice + " violence(s) policière(s) documentée(s)\n";
  if (nbViolencesTiers > 0) txt += "  ⚠ " + nbViolencesTiers + " violence(s) de tiers documentée(s)\n";
  if (nbTentatives > 0) txt += "  ⚠ " + nbTentatives + " tentative(s) de traversée\n";
  if (nbUrgencesSante > 0) txt += "  ⚠ " + nbUrgencesSante + " urgence(s) de santé\n";
  if (nbRefus > 0) txt += "  ⚠ " + nbRefus + " refus institutionnel(s)\n";
  txt += "\n";
  
  txt += "=========================================================================\n";
  txt += "3. HISTORIQUE CHRONOLOGIQUE DU PARCOURS\n";
  txt += "=========================================================================\n\n";
  
  chronoHist.forEach(function(evt) {
    if (evt.type !== "Modification") {
      txt += "Le " + evt.date + " - " + evt.type;
      if (evt.lieu && evt.lieu !== "-" && evt.lieu !== "") txt += " (" + evt.lieu + ")";
      txt += "\n";
      if (evt.details && evt.details.trim() !== "") txt += "   → " + evt.details + "\n";
      if (evt.auteur) txt += "   [Source : " + evt.auteur + "]\n";
      txt += "\n";
    }
  });
  
  txt += "=========================================================================\n";
  txt += "4. DEMANDE\n";
  txt += "=========================================================================\n\n";
  
  if (isMNA) {
    txt += "Au regard de la situation de danger décrite ci-dessus, l'association \n";
    txt += "Utopia 56 demande :\n\n";
    txt += "  1. La mise à l'abri IMMÉDIATE de ce mineur conformément à l'article \n";
    txt += "     L.223-2 du CASF.\n";
    txt += "  2. L'évaluation de sa minorité et de son isolement dans des conditions \n";
    txt += "     respectueuses de ses droits fondamentaux.\n";
    txt += "  3. L'ouverture d'une mesure de protection adaptée.\n\n";
    txt += "Nous rappelons que l'article L.112-3 du CASF impose aux autorités \n";
    txt += "compétentes de garantir la prise en compte des besoins fondamentaux \n";
    txt += "de l'enfant et que tout mineur présent sur le territoire français, \n";
    txt += "quelle que soit sa nationalité, bénéficie de la protection de l'enfance.\n\n";
  } else {
    txt += "Au regard de la situation de vulnérabilité décrite ci-dessus, \n";
    txt += "l'association Utopia 56 signale cette situation et demande une prise \n";
    txt += "en charge adaptée dans les meilleurs délais.\n\n";
  }
  
  txt += "=========================================================================\n";
  txt += "Signalement transmis par :\n";
  txt += "Association Utopia 56 - Antenne " + nomAntenne + "\n";
  txt += "Contact : " + contactSignalement + "\n";
  txt += "Date : " + aujourdhui + "\n";
  txt += "=========================================================================\n";
  
  return txt;
}

// ============================================================================
//          WIDGET CONFIG ITEMS POUR PLAIDOYER
// ============================================================================

function getWidgetConfigItems() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsConfig = ss.getSheetByName('BDD_CONFIG');
  const dataConfig = wsConfig.getDataRange().getValues();
  
  let result = {};
  if (dataConfig.length > 0) {
    const headers = dataConfig[0];
    for (let col = 0; col < headers.length; col++) {
      let key = headers[col];
      if (!key) continue;
      let items = [];
      for (let row = 1; row < dataConfig.length; row++) {
        let val = String(dataConfig[row][col] || "").trim();
        if (val !== "") {
          let isSeparator = val.startsWith('---') && val.endsWith('---');
          items.push({ value: val, isSeparator: isSeparator, label: isSeparator ? val.replace(/---/g, '').trim() : val });
        }
      }
      result[key] = items;
    }
  }
  return result;
}

// ============================================================================
//          COMPUTE ALL WIDGETS
// ============================================================================

function computeAllWidgets(widgets, startStr, endStr) {
  var stats = getStatisticsDetailed(startStr, endStr);
  var subcards = { 'presents': 'Présents', 'nouveaux': 'Nouveaux', 'total_foyers': 'Foyers' };
  var results = {};
  var ss = SpreadsheetApp.openById(SHEET_ID);
  
  var wsBenef = ss.getSheetByName('BDD_BENEFICIAIRES');
  var wsEvents = ss.getSheetByName('BDD_EVENTS');
  var dB = wsBenef.getDataRange().getValues();
  var dE = wsEvents.getDataRange().getValues();
  var start = parseDateSecure(startStr);
  var end = parseDateSecure(endStr);
  end.setHours(23, 59, 59);
  
  // Phase 1+2 : Index des colonnes mis à jour pour BDD_BENEFICIAIRES
  var benefFields = { 
    'statut_presence': 12, 'statut_institution': 14, 'nationalite': 6, 'genre': 5, 
    'lieu_vie': 13, 'vulnerabilites': 16, 'langue': 7, 'type_profil': 9, 'role_foyer': 11
  };
  var eventsFields = { 
    'type_evenement': 3, 'type_contact': 4, 'lieu_event': 5, 
    'motif_statut_institution': 9, 'materiel': 11 
  };
  
  widgets.forEach(function(widget) {
    try {
      if (widget.type === 'single_item') {
        if (subcards[widget.category]) {
          results[widget.id] = { value: stats[widget.category] || 0 };
        } else {
          var data = stats[widget.category];
          results[widget.id] = { value: (data && data[widget.item] !== undefined) ? data[widget.item] : 0 };
        }
        
      } else if (widget.type === 'item_list') {
        var items = widget.items || [];
        var listResults = [];
        items.forEach(function(it) {
          if (subcards[it.category]) {
            listResults.push({ name: it.name || subcards[it.category], count: stats[it.category] || 0 });
          } else {
            var data = stats[it.category];
            var val = (data && data[it.name] !== undefined) ? data[it.name] : 0;
            listResults.push({ name: it.name, count: val });
          }
        });
        results[widget.id] = { items: listResults, total: listResults.reduce(function(s, i) { return s + i.count; }, 0) };
        
      } else if (widget.type === 'custom_stat') {
        var conditions = widget.conditions || [];
        var resultMode = widget.resultMode || 'jeunes';
        var matchedBenef = new Set();
        var matchedEvents = 0;
        
        var benefConditions = conditions.filter(function(c) { return c.source === 'jeunes' || c.source === 'benef'; });
        var eventConditions = conditions.filter(function(c) { return c.source === 'events'; });
        
        for (var bi = 1; bi < dB.length; bi++) {
          var allBenefMatch = true;
          for (var ci = 0; ci < benefConditions.length; ci++) {
            var cond = benefConditions[ci];
            var colIdx = benefFields[cond.field];
            if (colIdx === undefined) { allBenefMatch = false; break; }
            var cellVal = String(dB[bi][colIdx] || '').toLowerCase();
            var condVal = String(cond.value || '').toLowerCase();
            if (cond.operator === 'equals' && cellVal !== condVal) { allBenefMatch = false; break; }
            if (cond.operator === 'contains' && !cellVal.includes(condVal)) { allBenefMatch = false; break; }
            if (cond.operator === 'not_equals' && cellVal === condVal) { allBenefMatch = false; break; }
          }
          if (allBenefMatch || benefConditions.length === 0) {
            if (eventConditions.length === 0 && benefConditions.length > 0) matchedBenef.add(dB[bi][0]);
          }
        }
        
        for (var ei = 1; ei < dE.length; ei++) {
          var evDate = null;
          try { evDate = new Date(dE[ei][2]); } catch(e) { continue; }
          if (!evDate || isNaN(evDate.getTime()) || evDate < start || evDate > end) continue;
          
          var allEventMatch = true;
          for (var ci2 = 0; ci2 < eventConditions.length; ci2++) {
            var cond2 = eventConditions[ci2];
            var colIdx2 = eventsFields[cond2.field];
            if (colIdx2 === undefined) { allEventMatch = false; break; }
            var cellVal2 = String(dE[ei][colIdx2] || '').toLowerCase();
            var condVal2 = String(cond2.value || '').toLowerCase();
            if (cond2.operator === 'equals' && cellVal2 !== condVal2) { allEventMatch = false; break; }
            if (cond2.operator === 'contains' && !cellVal2.includes(condVal2)) { allEventMatch = false; break; }
            if (cond2.operator === 'not_equals' && cellVal2 === condVal2) { allEventMatch = false; break; }
          }
          if (allEventMatch && eventConditions.length > 0) {
            matchedEvents++;
            matchedBenef.add(dE[ei][1]);
          }
        }
        
        var value = resultMode === 'events' ? matchedEvents : matchedBenef.size;
        var total = resultMode === 'events' ? stats.total_events : stats.total_actifs;
        var pct = total > 0 ? Math.round((value / total) * 100) : 0;
        
        results[widget.id] = { value: value, percentage: pct, jeunesCount: matchedBenef.size, eventsCount: matchedEvents };
      }
    } catch(e) {
      results[widget.id] = { value: 'ERR', error: e.toString() };
    }
  });
  
  results._fullStats = stats;
  return results;
}

// ============================================================================
//          DASHBOARD RAPPELS SUMMARY
// ============================================================================

function getDashboardRappelsSummary() {
  const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_RAPPELS');
  const wsBenef = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_BENEFICIAIRES');
  const data = ws.getDataRange().getValues();
  const dataBenef = wsBenef.getDataRange().getValues();
  const today = new Date(); today.setHours(0,0,0,0);
  
  let benefMap = {};
  for(let i=1; i<dataBenef.length; i++) {
    benefMap[dataBenef[i][0]] = dataBenef[i][1] + ' ' + (dataBenef[i][2] || '');
  }
  
  let enRetard = 0, aFaire = 0;
  let prochains = [];
  
  for (let i = 1; i < data.length; i++) {
    try {
      let statut = String(data[i][5] || '').toLowerCase().trim();
      if (statut.indexOf('faire') !== -1) {
        aFaire++;
        let echeance = parseDateSecure(data[i][3] || data[i][4]);
        if (echeance < today) enRetard++;
        let nom = benefMap[data[i][0]] || 'Inconnu';
        prochains.push({
          idBenef: data[i][0],
          nom: nom,
          titre: String(data[i][6] || "Sans titre"),
          date: formatDate(echeance),
          isLate: echeance < today
        });
      }
    } catch(e) {}
  }
  
  prochains.sort(function(a,b) {
    if (a.isLate && !b.isLate) return -1;
    if (!a.isLate && b.isLate) return 1;
    return 0;
  });

  return { enRetard: enRetard, aFaire: aFaire, prochains: prochains };
}

// ============================================================================
//  Phase 7 : EXPORT CSV DES RECHERCHES ET STATISTIQUES
// ============================================================================

function exportSearchResultsCSV(query, filters) {
  var results = searchBenefAdvanced(query, filters);
  if (!results || results.length === 0) return { success: false, message: "Aucun résultat à exporter." };
  
  var headers = ['ID', 'Prénom', 'Nom', 'Surnom', 'Genre', 'Nationalité', 'Type_Profil', 'Statut_Presence', 'Statut_Institution', 'Lieu_Vie', 'Vulnérabilités', 'ID_Foyer'];
  var csv = headers.join(';') + '\n';
  
  results.forEach(function(r) {
    var line = [
      r.id || '',
      (r.nom || '').split(' ')[0] || '',
      (r.nom || '').split(' ').slice(1).join(' ') || '',
      r.surnom || '',
      r.genre || '',
      r.nationalite || '',
      r.typeProfil || '',
      r.statut_presence || '',
      r.statut_institution || '',
      r.lieu || '',
      (r.vulns || '').replace(/;/g, ','),
      ''
    ].map(function(v) { return '"' + String(v).replace(/"/g, '""') + '"'; });
    csv += line.join(';') + '\n';
  });
  
  return { success: true, csv: csv, count: results.length, filename: 'export_recherche_' + new Date().toISOString().substring(0, 10) + '.csv' };
}

function exportStatisticsCSV(startStr, endStr) {
  var stats = getStatisticsDetailed(startStr, endStr);
  var csv = "Catégorie;Item;Valeur\n";
  
  csv += '"Résumé";"Présents";"' + stats.presents + '"\n';
  csv += '"Résumé";"Nouveaux";"' + stats.nouveaux + '"\n';
  csv += '"Résumé";"Total événements";"' + stats.total_events + '"\n';
  csv += '"Résumé";"Total foyers";"' + stats.total_foyers + '"\n';
  
  var categories = {
    'Nationalités': stats.nationalites, 'Langues': stats.langues,
    'Statut Présence': stats.statut_presence, 'Statut Institution': stats.statut_institution,
    'Lieux de Vie': stats.lieux_vie, 'Vulnérabilités': stats.vulnerabilites_tags,
    'Types Événements': stats.types_evenements, 'Types Contact': stats.types_contact,
    'Types Profils': stats.types_profils, 'Genres': stats.genres
  };
  
  for (var cat in categories) {
    var data = categories[cat];
    if (!data) continue;
    for (var item in data) {
      if (data.hasOwnProperty(item)) {
        csv += '"' + cat + '";"' + String(item).replace(/"/g, '""') + '";"' + data[item] + '"\n';
      }
    }
  }
  
  return { success: true, csv: csv, filename: 'export_stats_' + startStr + '_' + endStr + '.csv' };
}

// ============================================================================
//  STOCKAGE PRÉFÉRENCES UTILISATEUR (spreadsheet)
// ============================================================================

function saveUserPref(userEmail, cle, valeur) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('BDD_PREFS');
  if (!sheet) {
    sheet = ss.insertSheet('BDD_PREFS');
    sheet.appendRow(['email', 'cle', 'valeur']);
  }
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === userEmail && data[i][1] === cle) {
      sheet.getRange(i + 1, 3).setValue(valeur);
      return;
    }
  }
  sheet.appendRow([userEmail, cle, valeur]);
}

function loadUserPref(userEmail, cle) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('BDD_PREFS');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === userEmail && data[i][1] === cle) {
      return data[i][2] || null;
    }
  }
  return null;
}
