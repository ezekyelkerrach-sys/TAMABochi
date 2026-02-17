// ============================================================================
//                               CONFIGURATION
// ============================================================================

var SHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// ============================================================================
//                                ROUTAGE & HTML
// ============================================================================

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('TAMABochi - Utopia 56')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================================
//                            SECURITE & CONNEXION
// ============================================================================

function loginUser(email, password) {
  const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('USERS');
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == email && data[i][1] == password) {
      return { 
        success: true, 
        token: Utilities.base64Encode(email + "_" + new Date().getTime()),
        user: { email: data[i][0], role: data[i][2], nom: data[i][3] }
      };
    }
  }
  return { success: false, message: "Email ou mot de passe incorrect." };
}

// ============================================================================
//                          API : CONFIGURATION & SCENARIOS
// ============================================================================

function getConfigData() {
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

  const wsScenarios = ss.getSheetByName('BDD_SCENARIOS');
  if (wsScenarios) { 
    const dataScenarios = wsScenarios.getDataRange().getValues();
    let scenarios = {};
    for (let i = 1; i < dataScenarios.length; i++) {
      let type = dataScenarios[i][0];
      if(type) {
        scenarios[type] = {
          showLieuVie: !!dataScenarios[i][1], 
          showLieuEvent: !!dataScenarios[i][2],
          showPres: !!dataScenarios[i][3], 
          showMab: !!dataScenarios[i][4],
          showRefus: !!dataScenarios[i][5], 
          showVuln: !!dataScenarios[i][6],
          showContact: !!dataScenarios[i][7], 
          showMat: !!dataScenarios[i][8],
          showNote: !!dataScenarios[i][9]
        };
      }
    }
    config['SCENARIOS'] = scenarios;
  }
  
  return config;
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

function getJeuneStatus(id) {
   const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_JEUNES');
   const data = ws.getDataRange().getValues();
   for(let i=1; i<data.length; i++) {
     if(data[i][0] == id) {
       return { pres: data[i][9], mab: data[i][10], lieu: data[i][14] };
     }
   }
   return { pres:"", mab:"", lieu:"" };
}

// ============================================================================
//                          GESTION DES RAPPELS
// ============================================================================

// NOUVEAU : getRappelsAll retourne 3 catégories : enRetard, aFaire, archives
function getRappelsAll() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsRappels = ss.getSheetByName('BDD_RAPPELS');
  const wsJeunes = ss.getSheetByName('BDD_JEUNES');
  
  const dataRappels = wsRappels.getDataRange().getValues();
  const dataJeunes = wsJeunes.getDataRange().getValues();
  
  let jeunesMap = {};
  for(let i=1; i<dataJeunes.length; i++) {
    jeunesMap[dataJeunes[i][0]] = { nom: dataJeunes[i][3], tel: dataJeunes[i][2], lieu: dataJeunes[i][14] };
  }
  
  let enRetard = [];
  let aFaire = [];
  let archives = [];
  const today = new Date(); 
  today.setHours(0,0,0,0);

  for (let i = 1; i < dataRappels.length; i++) {
    try {
      let statut = String(dataRappels[i][5] || "").trim();
      let statutLower = statut.toLowerCase();
      let idJeune = dataRappels[i][0];
      let jInfo = jeunesMap[idJeune] || { nom: "Inconnu", tel: "", lieu: "" };
      let rawDate = dataRappels[i][4];
      let echeance = parseDateSecure(rawDate);
      
      let item = {
        idRappel: dataRappels[i][1], 
        idJeune: idJeune, 
        nomJeune: jInfo.nom, 
        telJeune: jInfo.tel, 
        lieuJeune: jInfo.lieu,
        titre: String(dataRappels[i][6] || "Sans titre"), 
        details: String(dataRappels[i][3] || ""), 
        dateRaw: echeance.toISOString(), 
        date: formatDate(echeance), 
        isLate: false,
        typeRappel: String(dataRappels[i][8] || ""),
        statut: statut
      };

      if (statutLower.indexOf("faire") !== -1) {
        // C'est un rappel "À faire"
        if (echeance < today) {
          item.isLate = true;
          enRetard.push(item);
        } else {
          aFaire.push(item);
        }
      } else {
        // Archivé (Fait, Abandonné, ou tout autre statut terminé)
        archives.push(item);
      }
    } catch (e) { console.log("Erreur ligne rappel " + i + ": " + e.toString()); }
  }
  
  // Tri : en retard par date croissante, à faire par date croissante, archives par date décroissante
  enRetard.sort(function(a,b) { return new Date(a.dateRaw) - new Date(b.dateRaw); });
  aFaire.sort(function(a,b) { return new Date(a.dateRaw) - new Date(b.dateRaw); });
  archives.sort(function(a,b) { return new Date(b.dateRaw) - new Date(a.dateRaw); });
  
  return { enRetard: enRetard, aFaire: aFaire, archives: archives };
}

// CONSERVEE pour compatibilité (dashboard, etc.)
function getRappels() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsRappels = ss.getSheetByName('BDD_RAPPELS');
  const wsJeunes = ss.getSheetByName('BDD_JEUNES');
  
  const dataRappels = wsRappels.getDataRange().getValues();
  const dataJeunes = wsJeunes.getDataRange().getValues();
  
  let jeunesMap = {};
  for(let i=1; i<dataJeunes.length; i++) {
    jeunesMap[dataJeunes[i][0]] = { nom: dataJeunes[i][3], tel: dataJeunes[i][2], lieu: dataJeunes[i][14] };
  }
  
  let result = [];
  const today = new Date(); 
  today.setHours(0,0,0,0);

  for (let i = 1; i < dataRappels.length; i++) {
    try {
      let statut = String(dataRappels[i][5] || "").toLowerCase().trim();
      if (statut.indexOf("faire") !== -1) {
        let idJeune = dataRappels[i][0];
        let jInfo = jeunesMap[idJeune] || { nom: "Inconnu", tel: "", lieu: "" };
        let rawDate = dataRappels[i][4];
        let echeance = parseDateSecure(rawDate);
        let isLate = echeance < today;
        
        result.push({
          idRappel: dataRappels[i][1], 
          idJeune: idJeune, 
          nomJeune: jInfo.nom, 
          telJeune: jInfo.tel, 
          lieuJeune: jInfo.lieu,
          titre: String(dataRappels[i][6] || "Sans titre"), 
          details: String(dataRappels[i][3] || ""), 
          dateRaw: echeance.toISOString(), 
          date: formatDate(echeance), 
          isLate: isLate,
          typeRappel: String(dataRappels[i][8] || "")
        });
      }
    } catch (e) { console.log("Erreur ligne rappel " + i + ": " + e.toString()); }
  }
  
  result.sort(function(a,b) {
    if (a.isLate && !b.isLate) return -1;
    if (!a.isLate && b.isLate) return 1;
    return new Date(a.dateRaw) - new Date(b.dateRaw);
  });
  return result;
}

function updateRappel(form) {
  const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_RAPPELS');
  const data = ws.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][1] == form.idRappel) {
      ws.getRange(i+1, 4).setValue(form.details); 
      ws.getRange(i+1, 5).setValue(new Date(form.date)); 
      return { success: true, message: "Rappel mis à jour." };
    }
  }
  return { success: false, message: "Rappel introuvable." };
}

function closeRappel(idRappel, auteurEmail, statusLabel) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsRappels = ss.getSheetByName('BDD_RAPPELS');
  const wsEvents = ss.getSheetByName('BDD_EVENTS');
  const data = wsRappels.getDataRange().getValues();
  const now = new Date();

  for(let i=1; i<data.length; i++) {
    if(data[i][1] == idRappel) {
      wsRappels.getRange(i+1, 6).setValue(statusLabel); 
      let idJeune = data[i][0];
      let titre = data[i][6];
      let details = data[i][3];
      let status = getJeuneStatus(idJeune);
      wsEvents.appendRow([Utilities.getUuid(), idJeune, now, "Suivi Rappel", "Distance", status.lieu, "Rappel traité (" + statusLabel + ") : " + titre + " - " + details, status.pres, status.mab, "", "", "", status.lieu, auteurEmail]);
      return { success: true };
    }
  }
  return { success: false, message: "Rappel introuvable" };
}

function createAutomaticReminders(idJeune, vulnsArray, eventType, details) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsRappels = ss.getSheetByName('BDD_RAPPELS');
  const wsRegles = ss.getSheetByName('BDD_REGLES_RAPPELS');
  const now = new Date();
  let rappelsToAdd = [];
  let rules = [];
  
  if (wsRegles) {
    const dataRules = wsRegles.getDataRange().getValues();
    for (let i = 1; i < dataRules.length; i++) {
      if(dataRules[i][0] !== "") {
        rules.push({ trigger: String(dataRules[i][0]).toLowerCase(), delay: dataRules[i][1], title: dataRules[i][2], msg: dataRules[i][3] });
      }
    }
  }

  let triggersToCheck = [];
  if (vulnsArray && vulnsArray.length > 0) vulnsArray.forEach(v => triggersToCheck.push(String(v).toLowerCase()));
  if (eventType) triggersToCheck.push(String(eventType).toLowerCase());

  rules.forEach(rule => {
    let match = triggersToCheck.some(t => t.includes(rule.trigger));
    if (match) {
      let ech = new Date(); ech.setDate(now.getDate() + parseInt(rule.delay));
      rappelsToAdd.push({ titre: rule.title, details: rule.msg, date: ech });
    }
  });

  // BDD_RAPPELS : A=ID_Jeune, B=ID_Rappel, C=Infos_Jeune, D=Details, E=Echeance, F=Statut, G=Titre, H=DateCreation, I=TypeRappel
  rappelsToAdd.forEach(r => { 
    wsRappels.appendRow([idJeune, Utilities.getUuid(), "", r.details, r.date, "A faire", r.titre, now, "Automatique"]); 
  });
}

// M7 : saveManualReminder accepte maintenant un typeRappel
function saveManualReminder(form) {
  const wsRappels = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_RAPPELS');
  const now = new Date();
  // BDD_RAPPELS : A=ID_Jeune, B=ID_Rappel, C=Infos_Jeune, D=Details, E=Echeance, F=Statut, G=Titre, H=DateCreation, I=TypeRappel
  wsRappels.appendRow([form.idJeune, Utilities.getUuid(), "", form.details, new Date(form.date), "A faire", form.titre, now, form.typeRappel || ""]);
  return { success: true, message: "Rappel ajouté." };
}

// ============================================================================
//                  C16 : RAPPEL MANUEL DEPUIS VUE RAPPELS
// ============================================================================

function searchJeunesForRappel(query) {
  const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_JEUNES');
  const data = ws.getDataRange().getValues();
  let results = [];
  let q = String(query).toLowerCase().trim();
  
  for (let i = 1; i < data.length; i++) {
    let nom = String(data[i][3]).toLowerCase();
    let surnom = String(data[i][4]).toLowerCase();
    let telClean = normalizeTel(data[i][2]);
    let qClean = normalizeTel(q);
    
    let matchText = nom.includes(q) || surnom.includes(q);
    let matchTel = (qClean.length > 3) && telClean.includes(qClean);
    
    if (matchText || matchTel) {
      results.push({
        id: data[i][0],
        nom: data[i][3],
        surnom: data[i][4],
        tel: data[i][2]
      });
    }
    if (results.length >= 20) break;
  }
  return results;
}

// ============================================================================
//      NOUVEAU : AJOUT COMMENTAIRE SUR RAPPEL (depuis vue Rappels)
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
//                          ENREGISTREMENT & DOUBLONS
// ============================================================================

function checkDoublonTel(tel) {
  if (!tel) return [];
  const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_JEUNES');
  const data = ws.getDataRange().getValues();
  let doublons = [];
  let inputTelClean = normalizeTel(tel);
  if (inputTelClean.length < 4) return []; 
  for (let i = 1; i < data.length; i++) {
    let dbTel = normalizeTel(data[i][2]);
    if (dbTel === inputTelClean && dbTel !== "") {
      doublons.push({ nom: data[i][3], surnom: data[i][4], id: data[i][0] });
    }
  }
  return doublons;
}

function saveJeuneSmart(form, auteurEmail, modeForce) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsJeunes = ss.getSheetByName('BDD_JEUNES');
  const wsEvents = ss.getSheetByName('BDD_EVENTS');
  const wsModifs = ss.getSheetByName('BDD_MODIFS');
  const telPropre = "'" + form.tel; 
  const now = new Date();
  
  let finalIdGroupe = Utilities.getUuid();
  let historiqueMsg = "Création fiche";
  // C8 : Historique Groupe
  let historiqueGroupeMsg = "";

  if (form.forceGroupId) { 
    finalIdGroupe = form.forceGroupId; 
    historiqueMsg = "Ajout au groupe (Enregistrement groupé)"; 
    // C8 : Tracer l'ajout au groupe
    historiqueGroupeMsg = "Ajouté au groupe (enreg. groupé) le " + formatDate(now);
  } 
  else if (!modeForce) {
    let doublons = checkDoublonTel(form.tel);
    if (doublons.length > 0) return { status: "CONFLICT", doublons: doublons, message: "Ce numéro existe déjà." };
  }
  else if (modeForce === 'JOIN_GROUP') {
    const data = wsJeunes.getDataRange().getValues();
    let inputTelClean = normalizeTel(form.tel);
    let groupFound = null;
    for (let i = 1; i < data.length; i++) {
      if (normalizeTel(data[i][2]) === inputTelClean) { 
        if (data[i][1]) {
          groupFound = data[i][1]; 
        } else { 
          groupFound = Utilities.getUuid(); 
          wsJeunes.getRange(i + 1, 2).setValue(groupFound); 
          // C8 : Mettre à jour l'historique groupe du jeune existant aussi
          let existingHist = String(data[i][17] || "");
          let newHist = existingHist ? existingHist + " | Groupe créé le " + formatDate(now) : "Groupe créé le " + formatDate(now);
          wsJeunes.getRange(i + 1, 18).setValue(newHist);
        } 
        break; 
      }
    }
    if (groupFound) finalIdGroupe = groupFound;
    historiqueMsg = "Ajout au groupe existant";
    // C8 : Tracer l'ajout au groupe
    historiqueGroupeMsg = "Rejoint groupe existant le " + formatDate(now);
  }
  else if (modeForce === 'NEW_OWNER') {
    const data = wsJeunes.getDataRange().getValues();
    let inputTelClean = normalizeTel(form.tel);
    for (let i = 1; i < data.length; i++) {
      if (normalizeTel(data[i][2]) === inputTelClean) { 
        wsModifs.appendRow([data[i][0], "Num_Tel (Rotation SIM)", data[i][2], "Inconnu", now, auteurEmail]); 
        wsJeunes.getRange(i + 1, 3).setValue("Inconnu"); 
      }
    }
    historiqueMsg = "Récupération numéro (Rotation SIM)";
  }

  const idJeune = Utilities.getUuid();
  const idEvent = Utilities.getUuid();
  const dateRencontre = form.dateRencontre ? new Date(form.dateRencontre) : now;
  const vulnString = form.vulnerabilites ? form.vulnerabilites.join(", ") : "";
  
  let initStatutPres = "Présent"; let initStatutMab = "En attente";
  const eventType = form.eventType || "Rencontre";
  if (eventType.includes("Départ UK") || eventType.includes("Parti")) initStatutPres = "Parti";
  else if (eventType.includes("Demande MAB")) initStatutMab = "Demandée";
  else if (eventType.includes("MAB") && eventType.includes("Foyer")) { initStatutMab = "Validée"; initStatutPres = "MAB (Foyer)"; }

  let labelsStr = Array.isArray(form.labelsAutres) ? form.labelsAutres.join(", ") : (form.labelsAutres || "");
  
  // M6 : Langues peut être multi-select, stocker en chaîne séparée par virgules
  let langueStr = Array.isArray(form.langue) ? form.langue.join(", ") : (form.langue || "");
  // M6 : TypeContact peut être multi-select
  let contactTypeStr = Array.isArray(form.typeContact) ? form.typeContact.join(", ") : (form.typeContact || "Non connu");

  // Colonnes BDD_JEUNES : A=ID, B=IDGroupe, C=Tel, D=Nom, E=Surnom, F=Age, G=Genre, H=Nationalite, I=Langue, 
  // J=StatutPresence, K=StatutMab, L=DateDernierContact, M=DateDerniereRencontre, N=DateDerniereDistrib, 
  // O=LieuVie, P=VulnTags, Q=HistoriqueJeune, R=HistoriqueGroupe, S=DateCreation, T=NotesFixes, U=LabelsAutres
  wsJeunes.appendRow([
    idJeune,              // A - ID_Jeune
    finalIdGroupe,        // B - ID_Groupe
    telPropre,            // C - Num_Tel
    form.nom,             // D - Prenom_Nom
    form.surnom,          // E - Surnom
    form.age,             // F - Age
    form.genre,           // G - Genre
    form.nationalite,     // H - Nationalite
    langueStr,            // I - Langue (M6: peut être multi)
    initStatutPres,       // J - Statut_Presence
    initStatutMab,        // K - Statut_Mab
    dateRencontre,        // L - Date_Dernier_Contact
    dateRencontre,        // M - Date_Derniere_Rencontre
    "",                   // N - Date_Derniere_Distrib
    form.lieuVie,         // O - Lieu_Vie
    vulnString,           // P - Vulnerabilite_Tags
    historiqueMsg,        // Q - Historique_Jeune
    historiqueGroupeMsg,  // R - Historique_Groupe (C8)
    now,                  // S - Date_Creation
    form.notesFixes || "",// T - Notes_Fixes
    labelsStr             // U - Labels_Autres
  ]);
  
  wsEvents.appendRow([idEvent, idJeune, dateRencontre, eventType, contactTypeStr, "", form.details || "Première rencontre", initStatutPres, initStatutMab, "", vulnString, "", form.lieuVie, auteurEmail]);
  
  createAutomaticReminders(idJeune, form.vulnerabilites, eventType, "");
  
  return { success: true, id: idJeune, groupId: finalIdGroupe, message: "Enregistrement réussi !" };
}

// ============================================================================
//                          MISE A JOUR & EVENEMENTS
// ============================================================================

function updateJeuneProfile(form, auteurEmail) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsJeunes = ss.getSheetByName('BDD_JEUNES');
  const wsModifs = ss.getSheetByName('BDD_MODIFS');
  const data = wsJeunes.getDataRange().getValues();
  const now = new Date();

  let rowIndex = -1;
  let current = {};

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == form.id) {
      rowIndex = i + 1;
      current = { 
        tel: data[i][2], nom: data[i][3], surnom: data[i][4], age: data[i][5], 
        lieu: data[i][14], pres: data[i][9], mab: data[i][10],
        notes: data[i][19] || "",
        labels: data[i][20] || ""
      };
      break;
    }
  }

  if (rowIndex === -1) return { success: false, message: "Jeune introuvable." };

  let changes = [];
  if (form.nom && form.nom !== current.nom) { wsModifs.appendRow([form.id, "Modification Nom", current.nom, form.nom, now, auteurEmail]); wsJeunes.getRange(rowIndex, 4).setValue(form.nom); changes.push("Nom"); }
  if (form.surnom !== current.surnom) { wsModifs.appendRow([form.id, "Modification Surnom", current.surnom, form.surnom, now, auteurEmail]); wsJeunes.getRange(rowIndex, 5).setValue(form.surnom); changes.push("Surnom"); }
  if (normalizeTel(form.tel) !== normalizeTel(current.tel)) { wsModifs.appendRow([form.id, "Modification Tel", current.tel, "'" + form.tel, now, auteurEmail]); wsJeunes.getRange(rowIndex, 3).setValue("'" + form.tel); changes.push("Téléphone"); }
  if (form.age != current.age) { wsModifs.appendRow([form.id, "Modification Age", current.age, form.age, now, auteurEmail]); wsJeunes.getRange(rowIndex, 6).setValue(form.age); changes.push("Age"); }
  if (form.lieu && form.lieu !== current.lieu) { wsModifs.appendRow([form.id, "Modification Lieu", current.lieu, form.lieu, now, auteurEmail]); wsJeunes.getRange(rowIndex, 15).setValue(form.lieu); changes.push("Lieu"); }
  
  if (form.statutPres && form.statutPres !== current.pres) { 
    wsModifs.appendRow([form.id, "Modification Statut Presence", current.pres, form.statutPres, now, auteurEmail]); 
    wsJeunes.getRange(rowIndex, 10).setValue(form.statutPres); 
    changes.push("Statut Présence"); 
  }
  if (form.statutMab && form.statutMab !== current.mab) { 
    wsModifs.appendRow([form.id, "Modification Statut MAB", current.mab, form.statutMab, now, auteurEmail]); 
    wsJeunes.getRange(rowIndex, 11).setValue(form.statutMab); 
    changes.push("Statut MAB"); 
  }

  if (form.notesFixes !== undefined && form.notesFixes !== current.notes) {
    wsModifs.appendRow([form.id, "Modification Notes Infos", current.notes, form.notesFixes, now, auteurEmail]);
    wsJeunes.getRange(rowIndex, 20).setValue(form.notesFixes);
    changes.push("Notes Fixes");
  }

  let newLabelsStr = Array.isArray(form.labelsAutres) ? form.labelsAutres.join(", ") : (form.labelsAutres || "");
  if (newLabelsStr !== current.labels) {
    wsModifs.appendRow([form.id, "Modification Labels Autres", current.labels, newLabelsStr, now, auteurEmail]);
    wsJeunes.getRange(rowIndex, 21).setValue(newLabelsStr);
    changes.push("Labels Autres");
  }

  if (changes.length === 0) return { success: false, message: "Aucune modification détectée." };
  return { success: true, message: "Modifications enregistrées : " + changes.join(", ") };
}

function saveNewEvent(form, auteurEmail) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsEvents = ss.getSheetByName('BDD_EVENTS');
  const wsJeunes = ss.getSheetByName('BDD_JEUNES');
  const dataJeunes = wsJeunes.getDataRange().getValues();
  const now = new Date(); 
  const dateEvent = form.date ? new Date(form.date) : now; 
  const idEvent = Utilities.getUuid();
  
  let rowIndex = -1; 
  let current = {};
  for (let i = 1; i < dataJeunes.length; i++) { 
    if (dataJeunes[i][0] == form.idJeune) { 
      rowIndex = i + 1; 
      current = { pres: dataJeunes[i][9], mab: dataJeunes[i][10], lieuVie: dataJeunes[i][14], vulns: dataJeunes[i][15] }; 
      break; 
    } 
  }
  if (rowIndex === -1) return { success: false, message: "Jeune introuvable." };
  
  // C10 : Mettre à jour Date_Dernier_Contact (col L) pour TOUS les types d'événements
  wsJeunes.getRange(rowIndex, 12).setValue(dateEvent);
  // C10 : Mettre à jour Date_Derniere_Rencontre (col M) pour TOUS les types d'événements
  wsJeunes.getRange(rowIndex, 13).setValue(dateEvent);
  
  // B1 : Mettre à jour Date_Derniere_Distrib (col N) si événement de type Distrib
  let eventTypeLower = String(form.type).toLowerCase();
  if (eventTypeLower.includes("distrib")) {
    wsJeunes.getRange(rowIndex, 14).setValue(dateEvent);
  }
  
  let finalLieuVie = current.lieuVie; 
  if (form.lieuVie && form.lieuVie !== "" && form.lieuVie !== current.lieuVie) { 
    wsJeunes.getRange(rowIndex, 15).setValue(form.lieuVie); 
    finalLieuVie = form.lieuVie; 
  }
  
  let finalStatutPres = current.pres; 
  if (form.statutPres && form.statutPres !== "" && form.statutPres !== current.pres) { 
    wsJeunes.getRange(rowIndex, 10).setValue(form.statutPres); 
    finalStatutPres = form.statutPres; 
  }
  
  let finalStatutMab = current.mab; 
  if (form.statutMab && form.statutMab !== "" && form.statutMab !== current.mab) { 
    wsJeunes.getRange(rowIndex, 11).setValue(form.statutMab); 
    finalStatutMab = form.statutMab; 
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
      wsJeunes.getRange(rowIndex, 16).setValue(newVulnString); 
    } 
  }
  
  let fullDetails = form.note || ""; 
  let materielStr = Array.isArray(form.materiel) ? form.materiel.join(", ") : (form.materiel || "");
  // M6 : TypeContact peut être multi-select
  let contactStr = Array.isArray(form.typeContact) ? form.typeContact.join(", ") : (form.typeContact || "Physique");
  // M1 : motifStatutMab (anciennement motifRefus)
  let motifStr = Array.isArray(form.motifStatutMab) ? form.motifStatutMab.join(", ") : (form.motifStatutMab || "");
  
  wsEvents.appendRow([idEvent, form.idJeune, dateEvent, form.type, contactStr, form.lieuEvent || "", fullDetails, finalStatutPres, finalStatutMab, motifStr, newVulnString, materielStr, finalLieuVie, auteurEmail]);
  
  createAutomaticReminders(form.idJeune, form.vulns, form.type, form.note);
  
  return { success: true, message: "Événement enregistré." };
}

function getEventData(idEvent) {
  const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_EVENTS');
  const data = ws.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (data[i][0] === idEvent) {
      return { 
        id: data[i][0], 
        dateRaw: data[i][2],
        type: data[i][3], 
        contact: data[i][4], 
        lieu: data[i][5], 
        details: data[i][6]
      };
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
    if (data[i][0] === form.idEvent) {
      // Log modification
      wsModifs.appendRow([data[i][1], "Modif Event (" + data[i][3] + ")", "", "Modifié par user", new Date(), auteurEmail]);
      
      wsEvents.getRange(i+1, 3).setValue(new Date(form.date)); // Date
      wsEvents.getRange(i+1, 4).setValue(form.type); // Type
      wsEvents.getRange(i+1, 7).setValue(form.details); // Note
      
      return { success: true, message: "Événement mis à jour" };
    }
  }
  return { success: false, message: "Événement introuvable" };
}

function saveMassEvent(form, auteurEmail) {
  if (!form.ids || form.ids.length === 0) return { success: false, message: "Aucun jeune sélectionné." };
  let successCount = 0;
  form.ids.forEach(id => {
    let singleForm = { 
      idJeune: id, type: form.type, date: form.date, lieuVie: form.lieuVie, lieuEvent: form.lieuEvent, 
      statutMab: form.statutMab, statutPres: form.statutPres, 
      motifStatutMab: form.motifStatutMab,  // M1
      materiel: form.materiel, typeContact: form.typeContact, vulns: form.vulns, note: form.note 
    };
    try { saveNewEvent(singleForm, auteurEmail); successCount++; } catch(e) { console.error(e); }
  });
  return { success: true, message: successCount + " jeunes mis à jour avec succès." };
}

// ============================================================================
//              C6 : SELECTION PAR LIEU DE VIE (EVICTION RAPIDE)
// ============================================================================

function getJeunesByLieuVie(lieuVie) {
  const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_JEUNES');
  const data = ws.getDataRange().getValues();
  let results = [];
  
  for (let i = 1; i < data.length; i++) {
    let statutPres = String(data[i][9] || "").toLowerCase();
    let lieu = String(data[i][14] || "");
    
    // Ne sélectionner que les jeunes actifs/présents sur ce lieu
    if (lieu === lieuVie && (statutPres.includes("présent") || statutPres.includes("mab (foyer)") || statutPres.includes("hospitalisé"))) {
      results.push({
        id: data[i][0],
        nom: data[i][3],
        surnom: data[i][4],
        age: data[i][5],
        tel: data[i][2],
        lieu: data[i][14]
      });
    }
  }
  return results;
}

// ============================================================================
//              C11 & C12 & M5 : RECHERCHE AVANCEE AVEC FILTRES
// ============================================================================

function searchJeunesAdvanced(query, filters) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsJeunes = ss.getSheetByName('BDD_JEUNES');
  const wsModifs = ss.getSheetByName('BDD_MODIFS');
  const dataJeunes = wsJeunes.getDataRange().getValues();
  const dataModifs = wsModifs.getDataRange().getValues();
  
  let groupCounts = {};
  for(let i=1; i<dataJeunes.length; i++) {
    let gid = dataJeunes[i][1]; 
    if(gid) groupCounts[gid] = (groupCounts[gid] || 0) + 1;
  }

  let results = [];
  let foundIDs = new Set();
  let q = query ? String(query).toLowerCase().trim() : "";
  let qClean = normalizeTel(q);
  
  // Extraction des filtres
  let filterPresence = (filters && filters.statutPresence) ? filters.statutPresence : "";
  let filterMab = (filters && filters.statutMab) ? filters.statutMab : "";
  let filterVuln = (filters && filters.vulnerabilite) ? filters.vulnerabilite.toLowerCase() : "";
  let filterLieu = (filters && filters.lieuVie) ? filters.lieuVie : "";
  let showAllActifs = (filters && filters.showAllActifs) ? true : false;
  // M5 : Mode afficher TOUS les jeunes sans filtre de présence
  let showAll = (filters && filters.showAll) ? true : false;
  
  for (let i = 1; i < dataJeunes.length; i++) {
    let rawTel = String(dataJeunes[i][2]);
    let telClean = normalizeTel(rawTel);
    let nom = String(dataJeunes[i][3]).toLowerCase();
    let surnom = String(dataJeunes[i][4]).toLowerCase();
    let lieu = String(dataJeunes[i][14]).toLowerCase();
    let statutPres = String(dataJeunes[i][9] || "");
    let statutMab = String(dataJeunes[i][10] || "");
    let vulns = String(dataJeunes[i][15] || "").toLowerCase();
    let lieuVie = String(dataJeunes[i][14] || "");
    
    // Filtre texte (recherche classique)
    let matchText = false;
    if (q === "" && (showAll || showAllActifs || filterPresence || filterMab || filterVuln || filterLieu)) {
      // Pas de texte tapé mais des filtres actifs ou mode "tous"
      matchText = true;
    } else if (q !== "") {
      matchText = nom.includes(q) || surnom.includes(q) || lieu.includes(q);
      let matchTel = (qClean.length > 3) && telClean.includes(qClean);
      matchText = matchText || matchTel;
    } else {
      continue; // Ni texte ni filtre, on skip
    }
    
    if (!matchText) continue;
    
    // C12 : Mode "tous les actifs"
    if (showAllActifs) {
      let presLower = statutPres.toLowerCase();
      if (!(presLower.includes("présent") || presLower.includes("mab (foyer)") || presLower.includes("hospitalisé"))) {
        continue;
      }
    }
    
    // M5 : Mode "tous les jeunes" — pas de filtre de présence appliqué
    // (showAll ne filtre rien, on passe tout)
    
    // C11 : Filtres avancés
    if (filterPresence && statutPres !== filterPresence) continue;
    if (filterMab && statutMab !== filterMab) continue;
    if (filterVuln && !vulns.includes(filterVuln.toLowerCase())) continue;
    if (filterLieu && lieuVie !== filterLieu) continue;
    
    foundIDs.add(dataJeunes[i][0]);
    let gid = dataJeunes[i][1];
    let isGrouped = (gid && groupCounts[gid] > 1);
    
    let labels = dataJeunes[i][20] || "";
    let genre = String(dataJeunes[i][6]).toUpperCase();

    results.push({
      id: dataJeunes[i][0], tel: dataJeunes[i][2], nom: dataJeunes[i][3], surnom: dataJeunes[i][4],
      age: dataJeunes[i][5], nationalite: dataJeunes[i][7], statut_mab: dataJeunes[i][10], lieu: dataJeunes[i][14], is_history: false,
      statut_presence: dataJeunes[i][9],
      is_grouped: isGrouped,
      labels: labels,
      genre: genre,
      vulns: dataJeunes[i][15] || ""
    });
    
    if (results.length >= 200) break; 
  }

  // Recherche dans historique des anciens numéros (seulement si recherche texte)
  if (q !== "" && qClean.length > 3) {
    for (let j = 1; j < dataModifs.length; j++) {
      let oldVal = normalizeTel(dataModifs[j][2]); 
      if (oldVal.includes(qClean)) {
        let idAncien = dataModifs[j][0];
        if (!foundIDs.has(idAncien)) {
          for (let k = 1; k < dataJeunes.length; k++) {
            if (dataJeunes[k][0] == idAncien) {
              foundIDs.add(idAncien);
              results.push({ id: dataJeunes[k][0], tel: dataJeunes[k][2], nom: dataJeunes[k][3], surnom: dataJeunes[k][4], age: dataJeunes[k][5], nationalite: dataJeunes[k][7], statut_mab: dataJeunes[k][10], lieu: dataJeunes[k][14], is_history: true, statut_presence: dataJeunes[k][9], is_grouped: false, labels:"", genre:"", vulns:"" });
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

// On conserve l'ancienne fonction pour compatibilité
function searchJeunes(query) {
  return searchJeunesAdvanced(query, null);
}

// ============================================================================
//              U4 : DONNEES RAPPELS PAR JEUNE (pour indicateurs visuels)
// ============================================================================

function getRappelsEnRetardParJeune() {
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
        let idJeune = data[i][0];
        retardMap[idJeune] = (retardMap[idJeune] || 0) + 1;
      }
    }
  }
  return retardMap;
}

// ============================================================================
//                          RECHERCHE & FICHE & GROUPE
// ============================================================================

function getFicheJeune(idJeune) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const dataJeunes = ss.getSheetByName('BDD_JEUNES').getDataRange().getValues();
  let infoJeune = null;
  
  let groupCounts = {};
  for(let i=1; i<dataJeunes.length; i++) {
    let gid = dataJeunes[i][1];
    if(gid) groupCounts[gid] = (groupCounts[gid] || 0) + 1;
  }

  for (let i = 1; i < dataJeunes.length; i++) {
    if (dataJeunes[i][0] == idJeune) {
      let gid = dataJeunes[i][1];
      let hasGroup = (gid && groupCounts[gid] > 1);
      
      infoJeune = {
        id: dataJeunes[i][0], id_groupe: gid, tel: dataJeunes[i][2], nom: dataJeunes[i][3],
        surnom: dataJeunes[i][4], age: dataJeunes[i][5], genre: dataJeunes[i][6], nationalite: dataJeunes[i][7],
        langue: dataJeunes[i][8], statut_presence: dataJeunes[i][9], statut_mab: dataJeunes[i][10],
        date_contact: formatDate(dataJeunes[i][11]), 
        date_rencontre: formatDate(dataJeunes[i][12]),
        // C9 : Date dernière distribution
        date_distrib: formatDate(dataJeunes[i][13]),
        lieu_vie: dataJeunes[i][14], vulnerabilites: dataJeunes[i][15],
        has_group: hasGroup,
        historique_groupe: dataJeunes[i][17] || "",
        notes_fixes: dataJeunes[i][19] || "",
        labels_autres: dataJeunes[i][20] || ""
      };
      break;
    }
  }
  if (!infoJeune) return { success: false, message: "Jeune introuvable" };

  const dataEvents = ss.getSheetByName('BDD_EVENTS').getDataRange().getValues();
  let history = [];
  for (let j = 1; j < dataEvents.length; j++) {
    if (dataEvents[j][1] == idJeune) {
      history.push({ idEvent: dataEvents[j][0], date: formatDate(dataEvents[j][2]), type: dataEvents[j][3], lieu: dataEvents[j][5], details: dataEvents[j][6], statut_mab_event: dataEvents[j][8], auteur: dataEvents[j][13] });
    }
  }
  const dataModifs = ss.getSheetByName('BDD_MODIFS').getDataRange().getValues();
  for (let m = 1; m < dataModifs.length; m++) {
    if (dataModifs[m][0] == idJeune) {
       history.push({ date: formatDate(dataModifs[m][4]), type: "Modification", lieu: "-", details: dataModifs[m][1] + " : " + dataModifs[m][2] + " -> " + dataModifs[m][3], auteur: dataModifs[m][5] });
    }
  }
  history.reverse(); 
  return { success: true, jeune: infoJeune, history: history };
}

function getGroupMembers(groupId, excludeId) {
  const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_JEUNES');
  const data = ws.getDataRange().getValues();
  let members = [];
  
  for(let i=1; i<data.length; i++) {
    if(data[i][1] === groupId && data[i][0] !== excludeId) {
      members.push({
        id: data[i][0],
        nom: data[i][3],
        surnom: data[i][4],
        age: data[i][5]
      });
    }
  }
  return members;
}

// ============================================================================
//    NOUVEAU : LIAISON GROUPE PAR ID_GROUPE (sans partage de numéro)
// ============================================================================

// Recherche de jeunes pour la liaison groupe (par nom, surnom, tel)
function searchJeunesForGroupLink(query) {
  const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_JEUNES');
  const data = ws.getDataRange().getValues();
  let results = [];
  let q = String(query).toLowerCase().trim();
  let qClean = normalizeTel(q);
  
  for (let i = 1; i < data.length; i++) {
    let nom = String(data[i][3]).toLowerCase();
    let surnom = String(data[i][4]).toLowerCase();
    let telClean = normalizeTel(data[i][2]);
    
    let matchText = nom.includes(q) || surnom.includes(q);
    let matchTel = (qClean.length > 3) && telClean.includes(qClean);
    
    if (matchText || matchTel) {
      let gid = data[i][1];
      results.push({
        id: data[i][0],
        nom: data[i][3],
        surnom: data[i][4],
        tel: data[i][2],
        age: data[i][5],
        lieu: data[i][14],
        id_groupe: gid || ""
      });
    }
    if (results.length >= 30) break;
  }
  return results;
}

// Lier un ensemble de jeunes à un groupe commun
// targetJeuneIds : tableau d'IDs de jeunes à lier au groupe du sourceJeuneId
// Si sourceJeuneId a déjà un groupe, on y ajoute les targets
// Sinon on crée un nouveau groupe pour tout le monde
function linkJeunesToGroup(sourceJeuneId, targetJeuneIds) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsJeunes = ss.getSheetByName('BDD_JEUNES');
  const data = wsJeunes.getDataRange().getValues();
  const now = new Date();
  
  if (!targetJeuneIds || targetJeuneIds.length === 0) {
    return { success: false, message: "Aucun jeune cible sélectionné." };
  }
  
  // 1. Trouver le groupe du source (ou en créer un)
  let groupId = null;
  let sourceRowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === sourceJeuneId) {
      sourceRowIndex = i + 1;
      groupId = data[i][1] || null;
      break;
    }
  }
  
  if (sourceRowIndex === -1) {
    return { success: false, message: "Jeune source introuvable." };
  }
  
  // Si pas de groupe existant, en créer un
  if (!groupId) {
    groupId = Utilities.getUuid();
    wsJeunes.getRange(sourceRowIndex, 2).setValue(groupId);
    // Mettre à jour l'historique groupe du source
    let existingHist = String(data[sourceRowIndex - 1][17] || "");
    let newHist = existingHist ? existingHist + " | Groupe créé le " + formatDate(now) : "Groupe créé le " + formatDate(now);
    wsJeunes.getRange(sourceRowIndex, 18).setValue(newHist);
  }
  
  // 2. Mettre à jour tous les jeunes cibles
  let linkedCount = 0;
  for (let i = 1; i < data.length; i++) {
    if (targetJeuneIds.includes(data[i][0])) {
      let currentGroupId = data[i][1];
      if (currentGroupId !== groupId) {
        wsJeunes.getRange(i + 1, 2).setValue(groupId);
        // Mettre à jour l'historique groupe
        let existingHist = String(data[i][17] || "");
        let newHist = existingHist ? existingHist + " | Rejoint groupe le " + formatDate(now) : "Rejoint groupe le " + formatDate(now);
        wsJeunes.getRange(i + 1, 18).setValue(newHist);
        linkedCount++;
      }
    }
  }
  
  return { success: true, message: linkedCount + " jeune(s) lié(s) au groupe.", groupId: groupId };
}

// Lier pendant l'enregistrement : on passe les IDs des jeunes existants à lier
// Le nouveau jeune sera enregistré avec le même groupId
function getGroupIdForLinking(targetJeuneIds) {
  if (!targetJeuneIds || targetJeuneIds.length === 0) return null;
  
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsJeunes = ss.getSheetByName('BDD_JEUNES');
  const data = wsJeunes.getDataRange().getValues();
  const now = new Date();
  
  // Chercher si un des jeunes cibles a déjà un groupId
  let existingGroupId = null;
  for (let i = 1; i < data.length; i++) {
    if (targetJeuneIds.includes(data[i][0]) && data[i][1]) {
      existingGroupId = data[i][1];
      break;
    }
  }
  
  // Si aucun n'a de groupe, en créer un
  if (!existingGroupId) {
    existingGroupId = Utilities.getUuid();
  }
  
  // Mettre à jour tous les jeunes cibles avec ce groupId
  for (let i = 1; i < data.length; i++) {
    if (targetJeuneIds.includes(data[i][0])) {
      if (data[i][1] !== existingGroupId) {
        wsJeunes.getRange(i + 1, 2).setValue(existingGroupId);
        let existingHist = String(data[i][17] || "");
        let newHist = existingHist ? existingHist + " | Groupe lié le " + formatDate(now) : "Groupe lié le " + formatDate(now);
        wsJeunes.getRange(i + 1, 18).setValue(newHist);
      }
    }
  }
  
  return existingGroupId;
}

// ============================================================================
//          PLAIDOYER & STATS — NOUVELLE VERSION DETAILLEE
// ============================================================================
// Nouvelle fonction qui compte chaque item de chaque colonne BDD_CONFIG
// dans les données jeunes (stock) et events (période).
// Détecte aussi les items "orphelins" (supprimés de BDD_CONFIG mais encore
// présents dans les données).
// ============================================================================

function getStatisticsDetailed(startStr, endStr) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const wsJeunes = ss.getSheetByName('BDD_JEUNES');
    const wsEvents = ss.getSheetByName('BDD_EVENTS');
    const wsConfig = ss.getSheetByName('BDD_CONFIG');
    
    // Dates
    let start = parseDateSecure(startStr);
    let end = parseDateSecure(endStr);
    end.setHours(23, 59, 59); 

    const dataJeunes = wsJeunes.getDataRange().getValues();
    const dataEvents = wsEvents.getDataRange().getValues();

    // ---- Charger les listes BDD_CONFIG ----
    const dataConfig = wsConfig.getDataRange().getValues();
    let configLists = {};
    if (dataConfig.length > 0) {
      const headers = dataConfig[0];
      for (let col = 0; col < headers.length; col++) {
        let key = headers[col];
        if (!key) continue;
        let values = [];
        for (let row = 1; row < dataConfig.length; row++) {
          if (dataConfig[row][col] !== "") {
            values.push(String(dataConfig[row][col]));
          }
        }
        configLists[key] = values;
      }
    }

    // ---- Compteurs de base (sous-cartes Présence) ----
    let presents = 0;
    let mineurs_moins_15 = 0;
    let filles = 0;
    let nouveaux = 0;

    // ---- Compteurs détaillés par colonne BDD_CONFIG ----
    // Données issues de BDD_JEUNES (stock actuel, toutes fiches)
    let countNationalites = {};
    let countLangues = {};
    let countStatutPresence = {};
    let countStatutsMab = {};
    let countVulnerabilitesTags = {};
    let countLieuxVie = {};

    // Données issues de BDD_EVENTS (sur la période)
    let countTypesEvenements = {};
    let countMotifStatutMab = {};
    let countMateriel = {};
    let countTypesContact = {};

    // ---- Fonctions utilitaires pour incrémenter les compteurs ----
    function incrementMulti(counterObj, rawValue) {
      if (!rawValue) return;
      var str = String(rawValue);
      str.split(',').forEach(function(part) {
        var trimmed = part.trim();
        if (trimmed !== "" && !trimmed.startsWith('---')) {
          counterObj[trimmed] = (counterObj[trimmed] || 0) + 1;
        }
      });
    }

    function incrementSingle(counterObj, rawValue) {
      if (!rawValue) return;
      var trimmed = String(rawValue).trim();
      if (trimmed !== "" && !trimmed.startsWith('---')) {
        counterObj[trimmed] = (counterObj[trimmed] || 0) + 1;
      }
    }

    // ============================================================
    // 1. Analyse BDD_JEUNES (toutes les fiches)
    // ============================================================
    for (let i = 1; i < dataJeunes.length; i++) {
      let dateCreation = parseDateSecure(dataJeunes[i][18]); // Col S
      let statutPres = String(dataJeunes[i][9] || "").trim(); // Col J
      let genre = String(dataJeunes[i][6]).toUpperCase(); // Col G
      let age = dataJeunes[i][5]; // Col F
      let nationalite = String(dataJeunes[i][7] || ""); // Col H
      let langue = String(dataJeunes[i][8] || ""); // Col I
      let statutMab = String(dataJeunes[i][10] || ""); // Col K
      let vulns = String(dataJeunes[i][15] || ""); // Col P
      let lieuVie = String(dataJeunes[i][14] || ""); // Col O

      // Compteurs par colonne (toutes fiches)
      incrementSingle(countStatutPresence, statutPres);
      incrementSingle(countStatutsMab, statutMab);
      incrementSingle(countNationalites, nationalite);
      incrementMulti(countLangues, langue);
      incrementMulti(countVulnerabilitesTags, vulns);
      incrementSingle(countLieuxVie, lieuVie);

      // Sous-carte : Présents (statut contenant "présent")
      let presLower = statutPres.toLowerCase();
      if (presLower.includes("présent")) {
        presents++;
      }

      // Sous-carte : Moins de 15 ans
      if (age && parseInt(age) < 15) {
        mineurs_moins_15++;
      }

      // Sous-carte : Filles
      if (genre === "F" || genre === "FEMININ" || genre === "FILLE") {
        filles++;
      }

      // Sous-carte : Nouveaux dans la période
      if (dateCreation >= start && dateCreation <= end) {
        nouveaux++;
      }
    }

    // ============================================================
    // 2. Analyse BDD_EVENTS (sur la période sélectionnée)
    // ============================================================
    for (let j = 1; j < dataEvents.length; j++) {
      let dateEvt = parseDateSecure(dataEvents[j][2]); // Col C
      if (dateEvt >= start && dateEvt <= end) {
        let typeEvent = String(dataEvents[j][3] || ""); // Col D
        let typeContact = String(dataEvents[j][4] || ""); // Col E
        let motif = String(dataEvents[j][9] || ""); // Col J
        let materiel = String(dataEvents[j][11] || ""); // Col L

        incrementSingle(countTypesEvenements, typeEvent);
        incrementMulti(countTypesContact, typeContact);
        incrementMulti(countMotifStatutMab, motif);
        incrementMulti(countMateriel, materiel);
      }
    }

    // ============================================================
    // 3. Détection des items "orphelins"
    // ============================================================
    let orphelins = {};

    function detectOrphelins(counterObj, configKey) {
      var configItems = configLists[configKey] || [];
      var cleanConfig = configItems.filter(function(item) {
        return !String(item).startsWith('---');
      });
      for (var item in counterObj) {
        if (counterObj.hasOwnProperty(item) && counterObj[item] > 0) {
          var found = false;
          for (var c = 0; c < cleanConfig.length; c++) {
            if (cleanConfig[c] === item) {
              found = true;
              break;
            }
          }
          if (!found) {
            orphelins[item + " (" + configKey + ")"] = counterObj[item];
          }
        }
      }
    }

    detectOrphelins(countNationalites, "NATIONALITES");
    detectOrphelins(countLangues, "LANGUES");
    detectOrphelins(countTypesEvenements, "TYPES_EVENEMENTS");
    detectOrphelins(countStatutPresence, "STATUT_PRESENCE");
    detectOrphelins(countStatutsMab, "STATUTS_MAB");
    detectOrphelins(countMotifStatutMab, "MOTIF_STATUT_MAB");
    detectOrphelins(countMateriel, "MATERIEL");
    detectOrphelins(countVulnerabilitesTags, "VULNERABILITES_TAGS");
    detectOrphelins(countTypesContact, "TYPES_CONTACT");
    detectOrphelins(countLieuxVie, "LIEUX_VIE");

    // ============================================================
    // 4. Calculer les totaux
    // ============================================================
    function sumValues(obj) {
      var total = 0;
      for (var k in obj) {
        if (obj.hasOwnProperty(k)) total += obj[k];
      }
      return total;
    }

    // ============================================================
    // 5. Résultat final — NOUVEAU : inclure aussi configLists pour le front
    // ============================================================
    return {
      // Sous-cartes Présence
      presents: presents,
      mineurs_moins_15: mineurs_moins_15,
      filles: filles,
      nouveaux: nouveaux,

      // Listes détaillées (compteurs par item)
      nationalites: countNationalites,
      langues: countLangues,
      types_evenements: countTypesEvenements,
      total_types_evenements: sumValues(countTypesEvenements),
      statut_presence: countStatutPresence,
      statuts_mab: countStatutsMab,
      motif_statut_mab: countMotifStatutMab,
      total_motif_statut_mab: sumValues(countMotifStatutMab),
      materiel: countMateriel,
      total_materiel: sumValues(countMateriel),
      vulnerabilites_tags: countVulnerabilitesTags,
      total_vulnerabilites_tags: sumValues(countVulnerabilitesTags),
      types_contact: countTypesContact,
      total_types_contact: sumValues(countTypesContact),
      lieux_vie: countLieuxVie,

      // Items orphelins
      orphelins: orphelins,

      // NOUVEAU : Listes BDD_CONFIG brutes (pour que le front puisse afficher les items à 0 dans le sélecteur favoris)
      config_lists: configLists,

      // Métadonnées période
      periode_debut: formatDate(start),
      periode_fin: formatDate(end)
    };

  } catch(e) {
    throw new Error("Erreur calcul stats détaillées: " + e.toString());
  }
}

// ============================================================================
// getStatistics() CONSERVEE pour compatibilité avec le dashboard
// ============================================================================

function getStatistics(startStr, endStr) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const wsJeunes = ss.getSheetByName('BDD_JEUNES');
    const wsEvents = ss.getSheetByName('BDD_EVENTS');
    
    // Dates
    let start = parseDateSecure(startStr);
    let end = parseDateSecure(endStr);
    end.setHours(23, 59, 59); 

    const dataJeunes = wsJeunes.getDataRange().getValues();
    const dataEvents = wsEvents.getDataRange().getValues();

    // Initialisation Compteurs
    let stats = {
      nouveaux: 0,
      total_actifs: 0,
      mineurs_moins_15: 0,
      filles: 0,
      sous_emprise: 0,
      
      referencements_instit: 0,
      refus_police: 0,
      refus_dept: 0,
      mab_validees: 0,
      
      departs_uk: 0,
      retours_try: 0,
      evictions: 0,
      urgences_sante: 0,
      tentes_distribuees: 0,
      
      total_events_period: 0,
      jeunes_with_events: new Set(),
      
      nationalites: {},
      ages: {},
      vulnerabilites: {},
      // C14 : Vulnérabilités ajoutées sur la période via événements
      vulnerabilites_periode: {}
    };

    // 1. Analyse Jeunes
    for (let i = 1; i < dataJeunes.length; i++) {
      let dateCreation = parseDateSecure(dataJeunes[i][18]); // Col S
      let statutPres = String(dataJeunes[i][9] || "").toLowerCase().trim(); // Col J
      let genre = String(dataJeunes[i][6]).toUpperCase(); // Col G
      let vulns = String(dataJeunes[i][15]).toLowerCase(); // Col P
      let age = dataJeunes[i][5];

      if (statutPres.includes("présent") || statutPres.includes("mab (foyer)")) {
        stats.total_actifs++;
        
        if (age && parseInt(age) < 15) stats.mineurs_moins_15++;
        if (genre === "F" || genre === "FEMININ") stats.filles++;
        if (vulns.includes("emprise")) stats.sous_emprise++;
        
        // Vuln breakdown
        if(vulns) {
          vulns.split(',').forEach(v => {
            let vt = v.trim();
            if(vt) stats.vulnerabilites[vt] = (stats.vulnerabilites[vt] || 0) + 1;
          });
        }
      }

      // Nouveaux dans la période
      if (dateCreation >= start && dateCreation <= end) {
        stats.nouveaux++;
        let nat = dataJeunes[i][7] || "Inconnue";
        stats.nationalites[nat] = (stats.nationalites[nat] || 0) + 1;
        
        let tranche = "Inconnu";
        if (age) {
          let a = parseInt(age);
          if (a < 15) tranche = "-15 ans";
          else if (a < 18) tranche = "15-17 ans";
          else tranche = "18+ ans";
        }
        stats.ages[tranche] = (stats.ages[tranche] || 0) + 1;
      }
    }

    // 2. Analyse Events
    for (let j = 1; j < dataEvents.length; j++) {
      let dateEvt = parseDateSecure(dataEvents[j][2]); // Col C
      if (dateEvt >= start && dateEvt <= end) {
        let type = String(dataEvents[j][3]).toLowerCase(); // Col D
        let details = String(dataEvents[j][6]).toLowerCase(); // Col G
        let motif = String(dataEvents[j][9]).toLowerCase(); // Col J (M1: maintenant MOTIF_STATUT_MAB)
        let materiel = String(dataEvents[j][11]).toLowerCase(); // Col L
        let idJeune = dataEvents[j][1];
        // C14 : Vulnérabilités ajoutées par cet événement
        let newVulnsEvt = String(dataEvents[j][10] || "").toLowerCase(); // Col K

        stats.total_events_period++;
        stats.jeunes_with_events.add(idJeune);

        if (type.includes("départ uk") || type.includes("parti")) stats.departs_uk++;
        if (type.includes("retour de try")) stats.retours_try++;
        if (type.includes("éviction")) stats.evictions++;
        if (type.includes("santé") || type.includes("urgence")) stats.urgences_sante++;
        
        if (type.includes("demande mab") || type.includes("référencement")) stats.referencements_instit++;
        if (type.includes("prise en charge") || type.includes("validée")) stats.mab_validees++;

        if (type.includes("refus") || motif.includes("refus")) {
          if (motif.includes("police") || motif.includes("paf")) stats.refus_police++;
          else if (motif.includes("département") || motif.includes("ase") || motif.includes("social")) stats.refus_dept++;
        }

        if (materiel.includes("tente")) stats.tentes_distribuees++;
        
        // C14 : Comptage des vulnérabilités apparues sur la période
        if (newVulnsEvt) {
          newVulnsEvt.split(',').forEach(v => {
            let vt = v.trim();
            if(vt) stats.vulnerabilites_periode[vt] = (stats.vulnerabilites_periode[vt] || 0) + 1;
          });
        }
      }
    }

    // C13 : Nombre de jeunes ayant vécu au moins un événement sur la période
    stats.nb_jeunes_concernes = stats.jeunes_with_events.size;
    stats.avg_event_per_youth = stats.nb_jeunes_concernes > 0 ? (stats.total_events_period / stats.nb_jeunes_concernes).toFixed(1) : 0;
    stats.avg_eviction_per_youth = stats.total_actifs > 0 ? (stats.evictions / stats.total_actifs).toFixed(2) : 0; 

    stats.jeunes_with_events = stats.jeunes_with_events.size;

    return stats;

  } catch(e) {
    throw new Error("Erreur calcul stats: " + e.toString());
  }
}

// ============================================================================
//                     B3 : RAPPORT IP AMELIORE (PLAIDOYER)
// ============================================================================

function generateReportText(idJeune) {
  const info = getFicheJeune(idJeune); 
  if (!info.success) return "Erreur: Jeune introuvable.";
  
  const j = info.jeune;
  const hist = info.history;
  
  // Date du jour formatée
  const aujourdhui = formatDate(new Date());
  
  // Calcul de l'ancienneté sur le terrain
  let chronoHist = [...hist].reverse();
  let premierEvt = chronoHist.length > 0 ? chronoHist[0].date : "Inconnue";
  
  // Comptage des événements significatifs pour le rapport
  let nbEvictions = 0;
  let nbViolencesPolice = 0;
  let nbViolencesTiers = 0;
  let nbTentatives = 0;
  let nbUrgencesSante = 0;
  let nbRefus = 0;
  
  chronoHist.forEach(function(evt) {
    let typeLower = String(evt.type).toLowerCase();
    if (typeLower.includes("éviction")) nbEvictions++;
    if (typeLower.includes("violences policières") || typeLower.includes("police")) nbViolencesPolice++;
    if (typeLower.includes("violences tiers") || typeLower.includes("passeur")) nbViolencesTiers++;
    if (typeLower.includes("retour de try") || typeLower.includes("départ uk")) nbTentatives++;
    if (typeLower.includes("santé") || typeLower.includes("urgence") || typeLower.includes("hospitalisation")) nbUrgencesSante++;
    if (typeLower.includes("refus")) nbRefus++;
  });
  
  let txt = "";
  txt += "=========================================================================\n";
  txt += "              INFORMATION PRÉOCCUPANTE (IP)\n";
  txt += "              SIGNALEMENT - MINEUR NON ACCOMPAGNÉ\n";
  txt += "=========================================================================\n\n";
  
  txt += "Émetteur : Association Utopia 56 - Antenne Grande-Synthe\n";
  txt += "Date du signalement : " + aujourdhui + "\n";
  txt += "Destinataire : Cellule de Recueil des Informations Préoccupantes (CRIP)\n";
  txt += "               Procureur de la République (le cas échéant)\n\n";
  
  txt += "-------------------------------------------------------------------------\n";
  txt += "FONDEMENT JURIDIQUE\n";
  txt += "-------------------------------------------------------------------------\n";
  txt += "Le présent signalement est effectué sur le fondement des articles :\n";
  txt += "- L.112-3 du Code de l'Action Sociale et des Familles (CASF) : \n";
  txt += "  \"La protection de l'enfance vise à garantir la prise en compte des \n";
  txt += "  besoins fondamentaux de l'enfant.\"\n";
  txt += "- L.223-2 du CASF relatif à la mise à l'abri immédiate des mineurs \n";
  txt += "  non accompagnés.\n";
  txt += "- Article 375 du Code Civil relatif à l'assistance éducative.\n";
  txt += "- Convention Internationale des Droits de l'Enfant (CIDE), notamment \n";
  txt += "  l'article 3 (intérêt supérieur de l'enfant) et l'article 20 \n";
  txt += "  (protection de l'enfant privé de son milieu familial).\n\n";
  
  txt += "=========================================================================\n";
  txt += "1. IDENTIFICATION DU MINEUR\n";
  txt += "=========================================================================\n\n";
  txt += "Nom / Alias        : " + j.nom + "\n";
  txt += "Surnom              : " + (j.surnom || "Non renseigné") + "\n";
  txt += "Âge déclaré         : " + j.age + " ans\n";
  txt += "Genre               : " + (j.genre || "Non renseigné") + "\n";
  txt += "Nationalité         : " + j.nationalite + "\n";
  txt += "Langue(s) parlée(s) : " + (j.langue || "Non renseignée") + "\n";
  txt += "Téléphone           : " + (j.tel || "Non renseigné") + "\n";
  txt += "Premier contact     : " + premierEvt + "\n\n";
  
  txt += "=========================================================================\n";
  txt += "2. SITUATION ACTUELLE - ÉLÉMENTS DE DANGER\n";
  txt += "=========================================================================\n\n";
  txt += "Lieu de vie actuel  : " + j.lieu_vie + "\n";
  txt += "Statut présence     : " + j.statut_presence + "\n";
  txt += "Statut MAB          : " + j.statut_mab + "\n\n";
  
  txt += "Vulnérabilités identifiées :\n";
  if (j.vulnerabilites) {
    j.vulnerabilites.split(',').forEach(function(v) {
      txt += "  • " + v.trim() + "\n";
    });
  } else {
    txt += "  Aucune vulnérabilité spécifique notée à ce jour.\n";
  }
  txt += "\n";
  
  if (j.notes_fixes) {
    txt += "Informations complémentaires importantes :\n";
    txt += "  " + j.notes_fixes + "\n\n";
  }
  
  if (j.labels_autres) {
    txt += "Suivi par d'autres acteurs : " + j.labels_autres + "\n\n";
  }
  
  txt += "ÉLÉMENTS CARACTÉRISANT LA SITUATION DE DANGER :\n\n";
  txt += "Ce mineur non accompagné vit actuellement en extérieur sur le site de \n";
  txt += j.lieu_vie + ", sans hébergement stable ni protection adulte. ";
  txt += "Il/elle est exposé(e) aux intempéries, aux violences, et ne bénéficie \n";
  txt += "d'aucun cadre protecteur adapté à sa minorité.\n\n";
  
  if (nbEvictions > 0) {
    txt += "- Le/la jeune a subi " + nbEvictions + " éviction(s) de son lieu de vie, \n";
    txt += "  aggravant sa précarité et son errance.\n";
  }
  if (nbViolencesPolice > 0) {
    txt += "- " + nbViolencesPolice + " épisode(s) de violences policières ont été \n";
    txt += "  documenté(s), constituant des atteintes graves à son intégrité physique.\n";
  }
  if (nbViolencesTiers > 0) {
    txt += "- " + nbViolencesTiers + " épisode(s) de violences par des tiers ou passeurs \n";
    txt += "  ont été rapporté(s), exposant ce mineur à un risque d'exploitation.\n";
  }
  if (nbUrgencesSante > 0) {
    txt += "- " + nbUrgencesSante + " urgence(s) de santé ont nécessité une intervention.\n";
  }
  if (nbRefus > 0) {
    txt += "- " + nbRefus + " refus de mise à l'abri ont été opposés à ce mineur, \n";
    txt += "  le/la maintenant dans une situation de danger.\n";
  }
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
  
  txt += "=========================================================================\n";
  txt += "Signalement transmis par :\n";
  txt += "Association Utopia 56 - Antenne Grande-Synthe\n";
  txt += "Contact : [À compléter]\n";
  txt += "Date : " + aujourdhui + "\n";
  txt += "=========================================================================\n";
  
  return txt;
}
