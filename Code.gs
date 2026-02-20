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
  return { success: false, message: "ID ou mot de passe incorrect." };
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
  if (!input || input === "") return null;
  var d;
  if (input instanceof Date) {
    d = input;
  } else {
    var str = String(input).trim();
    if (str === "") return null;
    if (str.includes('/')) {
      var parts = str.split(' ')[0].split('/');
      if (parts.length === 3) {
        d = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
      }
    }
    if (!d) d = new Date(str);
  }
  if (!d || isNaN(d.getTime())) return null;
  // Normaliser à minuit heure locale pour éviter les décalages de fuseau
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
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
        if (echeance < today) {
          item.isLate = true;
          enRetard.push(item);
        } else {
          aFaire.push(item);
        }
      } else {
        archives.push(item);
      }
    } catch (e) { console.log("Erreur ligne rappel " + i + ": " + e.toString()); }
  }
  
  enRetard.sort(function(a,b) { return new Date(a.dateRaw) - new Date(b.dateRaw); });
  aFaire.sort(function(a,b) { return new Date(a.dateRaw) - new Date(b.dateRaw); });
  archives.sort(function(a,b) { return new Date(b.dateRaw) - new Date(a.dateRaw); });
  
  return { enRetard: enRetard, aFaire: aFaire, archives: archives };
}

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
        let echeance = parseDateSecure(rawDate) || new Date();
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

  rappelsToAdd.forEach(r => { 
    wsRappels.appendRow([idJeune, Utilities.getUuid(), "", r.details, r.date, "A faire", r.titre, now, "Automatique"]); 
  });
}

function saveManualReminder(form) {
  const wsRappels = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_RAPPELS');
  const now = new Date();
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
  let historiqueGroupeMsg = "";

  if (form.forceGroupId) { 
    finalIdGroupe = form.forceGroupId; 
    historiqueMsg = "Ajout au groupe (Enregistrement groupé)"; 
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
          let existingHist = String(data[i][17] || "");
          let newHist = existingHist ? existingHist + " | Groupe créé le " + formatDate(now) : "Groupe créé le " + formatDate(now);
          wsJeunes.getRange(i + 1, 18).setValue(newHist);
        } 
        break; 
      }
    }
    if (groupFound) finalIdGroupe = groupFound;
    historiqueMsg = "Ajout au groupe existant";
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
  
  let initStatutPres = "Présent pas équipé"; let initStatutMab = "Non demandée";
  const eventType = form.eventType || "Rencontre";
  if (eventType.includes("Départ UK") || eventType.includes("Parti")) initStatutPres = "Parti";
  else if (eventType.includes("Demande MAB")) initStatutMab = "Demandée";
  else if (eventType.includes("MAB") && eventType.includes("Foyer")) { initStatutMab = "Validée"; initStatutPres = "MAB (Foyer)"; }

  let labelsStr = Array.isArray(form.labelsAutres) ? form.labelsAutres.join(", ") : (form.labelsAutres || "");
  let langueStr = Array.isArray(form.langue) ? form.langue.join(", ") : (form.langue || "");
  let contactTypeStr = Array.isArray(form.typeContact) ? form.typeContact.join(", ") : (form.typeContact || "Non connu");

  wsJeunes.appendRow([
    idJeune,              // A - ID_Jeune
    finalIdGroupe,        // B - ID_Groupe
    telPropre,            // C - Num_Tel
    form.nom,             // D - Prenom_Nom
    form.surnom,          // E - Surnom
    form.age,             // F - Age
    form.genre,           // G - Genre
    form.nationalite,     // H - Nationalite
    langueStr,            // I - Langue
    initStatutPres,       // J - Statut_Presence
    initStatutMab,        // K - Statut_Mab
    dateRencontre,        // L - Date_Dernier_Contact
    dateRencontre,        // M - Date_Derniere_Rencontre
    "",                   // N - Date_Derniere_Distrib
    form.lieuVie,         // O - Lieu_Vie
    vulnString,           // P - Vulnerabilite_Tags
    historiqueMsg,        // Q - Historique_Jeune
    historiqueGroupeMsg,  // R - Historique_Groupe
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
  if (normalizeTel(form.tel) !== normalizeTel(current.tel)) { 
    var telForSheet = "'" + String(form.tel).replace(/^'+/, '');
    wsModifs.appendRow([form.id, "Modification Tel", current.tel, telForSheet, now, auteurEmail]); 
    wsJeunes.getRange(rowIndex, 3).setNumberFormat('@').setValue(String(form.tel).replace(/^'+/, '')); 
    changes.push("Téléphone"); 
  }
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

  // Modification des vulnérabilités
  if (form.vulnerabilites !== undefined) {
    let newVulnStr = Array.isArray(form.vulnerabilites) ? form.vulnerabilites.join(", ") : (form.vulnerabilites || "");
    let currentVulns = String(data[rowIndex - 1][15] || "");
    if (newVulnStr !== currentVulns) {
      wsModifs.appendRow([form.id, "Modification Vulnérabilités", currentVulns, newVulnStr, now, auteurEmail]);
      wsJeunes.getRange(rowIndex, 16).setValue(newVulnStr);
      changes.push("Vulnérabilités");
    }
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
  
  wsJeunes.getRange(rowIndex, 12).setValue(dateEvent);
  wsJeunes.getRange(rowIndex, 13).setValue(dateEvent);
  
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
  let contactStr = Array.isArray(form.typeContact) ? form.typeContact.join(", ") : (form.typeContact || "Physique");
  let motifStr = Array.isArray(form.motifStatutMab) ? form.motifStatutMab.join(", ") : (form.motifStatutMab || "");
  
  wsEvents.appendRow([idEvent, form.idJeune, dateEvent, form.type, contactStr, form.lieuEvent || "", fullDetails, finalStatutPres, finalStatutMab, motifStr, newVulnString, materielStr, finalLieuVie, auteurEmail]);
  
  createAutomaticReminders(form.idJeune, form.vulns, form.type, form.note);
  
  return { success: true, message: "Événement enregistré." };
}

function getEventData(idEvent) {
  const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_EVENTS');
  const data = ws.getDataRange().getValues();
  for (let i=1; i<data.length; i++) {
    if (String(data[i][0]) === String(idEvent)) {
      var rawDate = data[i][2];
      var dateRawStr = "";
      try {
        dateRawStr = (rawDate instanceof Date) ? rawDate.toISOString() : String(rawDate || "");
      } catch(e) {
        dateRawStr = "";
      }
      return { 
        id: data[i][0], 
        dateRaw: dateRawStr,
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
  if (!form.ids || form.ids.length === 0) return { success: false, message: "Aucun jeune sélectionné." };
  let successCount = 0;
  form.ids.forEach(id => {
    let singleForm = { 
      idJeune: id, type: form.type, date: form.date, lieuVie: form.lieuVie, lieuEvent: form.lieuEvent, 
      statutMab: form.statutMab, statutPres: form.statutPres, 
      motifStatutMab: form.motifStatutMab,
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
  
  let filterPresence = (filters && filters.statutPresence) ? filters.statutPresence : "";
  let filterMab = (filters && filters.statutMab) ? filters.statutMab : "";
  let filterVuln = (filters && filters.vulnerabilite) ? filters.vulnerabilite.toLowerCase() : "";
  let filterLieu = (filters && filters.lieuVie) ? filters.lieuVie : "";
  let showAllActifs = (filters && filters.showAllActifs) ? true : false;
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
    
    let matchText = false;
    if (q === "" && (showAll || showAllActifs || filterPresence || filterMab || filterVuln || filterLieu)) {
      matchText = true;
    } else if (q !== "") {
      matchText = nom.includes(q) || surnom.includes(q) || lieu.includes(q);
      let matchTel = (qClean.length > 3) && telClean.includes(qClean);
      matchText = matchText || matchTel;
    } else {
      continue;
    }
    
    if (!matchText) continue;
    
    if (showAllActifs) {
      let presLower = statutPres.toLowerCase();
      if (!(presLower.includes("présent") || presLower.includes("mab (foyer)") || presLower.includes("hospitalisé"))) {
        continue;
      }
    }
    
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

function searchJeunes(query) {
  return searchJeunesAdvanced(query, null);
}

// ============================================================================
//              U4 : DONNEES RAPPELS PAR JEUNE
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
  var infoJeune = null;
  
  var groupCounts = {};
  for (var i = 1; i < dataJeunes.length; i++) {
    var gid = dataJeunes[i][1];
    if (gid) groupCounts[gid] = (groupCounts[gid] || 0) + 1;
  }

  for (var i = 1; i < dataJeunes.length; i++) {
    if (dataJeunes[i][0] == idJeune) {
      var gid = dataJeunes[i][1];
      var hasGroup = (gid && groupCounts[gid] > 1);
      
      infoJeune = {
        id: dataJeunes[i][0], id_groupe: gid, tel: dataJeunes[i][2], nom: dataJeunes[i][3],
        surnom: dataJeunes[i][4], age: dataJeunes[i][5], genre: dataJeunes[i][6], nationalite: dataJeunes[i][7],
        langue: dataJeunes[i][8], statut_presence: dataJeunes[i][9], statut_mab: dataJeunes[i][10],
        date_contact: formatDate(dataJeunes[i][11]), 
        date_rencontre: formatDate(dataJeunes[i][12]),
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

  var dataEvents = ss.getSheetByName('BDD_EVENTS').getDataRange().getValues();
  var history = [];
  for (var j = 1; j < dataEvents.length; j++) {
    if (dataEvents[j][1] == idJeune) {
      // Convertir dateRaw en string ISO pour éviter problème de sérialisation
      var rawEvtDate = dataEvents[j][2];
      var dateRawStr = "";
      try {
        dateRawStr = (rawEvtDate instanceof Date) ? rawEvtDate.toISOString() : String(rawEvtDate || "");
      } catch(e) {
        dateRawStr = "";
      }
      history.push({
        idEvent: dataEvents[j][0],
        date: formatDate(dataEvents[j][2]),
        dateRaw: dateRawStr,
        type: dataEvents[j][3],
        lieu: dataEvents[j][5],
        details: dataEvents[j][6],
        statut_mab_event: dataEvents[j][8],
        auteur: dataEvents[j][13]
      });
    }
  }

  var dataModifs = ss.getSheetByName('BDD_MODIFS').getDataRange().getValues();
  for (var m = 1; m < dataModifs.length; m++) {
    if (dataModifs[m][0] == idJeune) {
      // Convertir dateRaw en string ISO pour éviter problème de sérialisation
      var rawModDate = dataModifs[m][4];
      var dateRawModStr = "";
      try {
        dateRawModStr = (rawModDate instanceof Date) ? rawModDate.toISOString() : String(rawModDate || "");
      } catch(e) {
        dateRawModStr = "";
      }
      history.push({
        date: formatDate(dataModifs[m][4]),
        dateRaw: dateRawModStr,
        type: "Modification",
        lieu: "-",
        details: dataModifs[m][1] + " : " + dataModifs[m][2] + " -> " + dataModifs[m][3],
        auteur: dataModifs[m][5]
      });
    }
  }

  // Trier par date d'événement décroissante (plus récent en haut)
  history.sort(function(a, b) {
    try {
      var dateA = a.dateRaw ? new Date(a.dateRaw) : new Date(0);
      var dateB = b.dateRaw ? new Date(b.dateRaw) : new Date(0);
      if (isNaN(dateA.getTime())) dateA = new Date(0);
      if (isNaN(dateB.getTime())) dateB = new Date(0);
      return dateB.getTime() - dateA.getTime();
    } catch(e) {
      return 0;
    }
  });

  // Récupérer le dernier motif statut MAB depuis les events
  var lastMotifMab = "";
  for (var ev = dataEvents.length - 1; ev >= 1; ev--) {
    if (dataEvents[ev][1] == idJeune) {
      var motif = String(dataEvents[ev][9] || "").trim();
      if (motif) { lastMotifMab = motif; break; }
    }
  }
  infoJeune.motif_statut_mab = lastMotifMab;

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
//    LIAISON GROUPE PAR ID_GROUPE (sans partage de numéro)
// ============================================================================

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

function linkJeunesToGroup(sourceJeuneId, targetJeuneIds) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsJeunes = ss.getSheetByName('BDD_JEUNES');
  const data = wsJeunes.getDataRange().getValues();
  const now = new Date();
  
  if (!targetJeuneIds || targetJeuneIds.length === 0) {
    return { success: false, message: "Aucun jeune cible sélectionné." };
  }
  
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
  
  if (!groupId) {
    groupId = Utilities.getUuid();
    wsJeunes.getRange(sourceRowIndex, 2).setValue(groupId);
    let existingHist = String(data[sourceRowIndex - 1][17] || "");
    let newHist = existingHist ? existingHist + " | Groupe créé le " + formatDate(now) : "Groupe créé le " + formatDate(now);
    wsJeunes.getRange(sourceRowIndex, 18).setValue(newHist);
  }
  
  let linkedCount = 0;
  for (let i = 1; i < data.length; i++) {
    if (targetJeuneIds.includes(data[i][0])) {
      let currentGroupId = data[i][1];
      if (currentGroupId !== groupId) {
        wsJeunes.getRange(i + 1, 2).setValue(groupId);
        let existingHist = String(data[i][17] || "");
        let newHist = existingHist ? existingHist + " | Rejoint groupe le " + formatDate(now) : "Rejoint groupe le " + formatDate(now);
        wsJeunes.getRange(i + 1, 18).setValue(newHist);
        linkedCount++;
      }
    }
  }
  
  return { success: true, message: linkedCount + " jeune(s) lié(s) au groupe.", groupId: groupId };
}

function getGroupIdForLinking(targetJeuneIds) {
  if (!targetJeuneIds || targetJeuneIds.length === 0) return null;
  
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsJeunes = ss.getSheetByName('BDD_JEUNES');
  const data = wsJeunes.getDataRange().getValues();
  const now = new Date();
  
  let existingGroupId = null;
  for (let i = 1; i < data.length; i++) {
    if (targetJeuneIds.includes(data[i][0]) && data[i][1]) {
      existingGroupId = data[i][1];
      break;
    }
  }
  
  if (!existingGroupId) {
    existingGroupId = Utilities.getUuid();
  }
  
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

function getStatisticsDetailed(startStr, endStr) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const wsJeunes = ss.getSheetByName('BDD_JEUNES');
    const wsEvents = ss.getSheetByName('BDD_EVENTS');
    const wsConfig = ss.getSheetByName('BDD_CONFIG');
    
    let start = parseDateSecure(startStr);
    let end = parseDateSecure(endStr);
    end.setHours(23, 59, 59); 

    const dataJeunes = wsJeunes.getDataRange().getValues();
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
          if (dataConfig[row][col] !== "") {
            values.push(String(dataConfig[row][col]));
          }
        }
        configLists[key] = values;
      }
    }

    let presents = 0;
    let mineurs_moins_15 = 0;
    let filles = 0;
    let nouveaux = 0;

    let countNationalites = {};
    let countLangues = {};
    let countStatutPresence = {};
    let countStatutsMab = {};
    let countVulnerabilitesTags = {};
    let countLieuxVie = {};

    let countTypesEvenements = {};
    let countMotifStatutMab = {};
    let countMateriel = {};
    let countTypesContact = {};

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

    for (let i = 1; i < dataJeunes.length; i++) {
      let dateCreation = parseDateSecure(dataJeunes[i][18]);
      let statutPres = String(dataJeunes[i][9] || "").trim();
      let genre = String(dataJeunes[i][6]).toUpperCase();
      let age = dataJeunes[i][5];
      let nationalite = String(dataJeunes[i][7] || "");
      let langue = String(dataJeunes[i][8] || "");
      let statutMab = String(dataJeunes[i][10] || "");
      let vulns = String(dataJeunes[i][15] || "");
      let lieuVie = String(dataJeunes[i][14] || "");

      incrementSingle(countStatutPresence, statutPres);
      incrementSingle(countStatutsMab, statutMab);
      incrementSingle(countNationalites, nationalite);
      incrementMulti(countLangues, langue);
      incrementMulti(countVulnerabilitesTags, vulns);
      incrementSingle(countLieuxVie, lieuVie);

      let presLower = statutPres.toLowerCase();
      if (presLower.includes("présent")) {
        presents++;
      }

      if (age && parseInt(age) < 15) {
        mineurs_moins_15++;
      }

      if (genre === "F" || genre === "FEMININ" || genre === "FILLE") {
        filles++;
      }

      if (dateCreation && dateCreation >= start && dateCreation <= end) {
        nouveaux++;
      }
    }

    for (let j = 1; j < dataEvents.length; j++) {
      let dateEvt = parseDateSecure(dataEvents[j][2]);
      if (!dateEvt) continue;
      if (dateEvt >= start && dateEvt <= end) {
        let typeEvent = String(dataEvents[j][3] || "");
        let typeContact = String(dataEvents[j][4] || "");
        let motif = String(dataEvents[j][9] || "");
        let materiel = String(dataEvents[j][11] || "");

        incrementSingle(countTypesEvenements, typeEvent);
        incrementMulti(countTypesContact, typeContact);
        incrementMulti(countMotifStatutMab, motif);
        incrementMulti(countMateriel, materiel);
      }
    }

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

    function sumValues(obj) {
      var total = 0;
      for (var k in obj) {
        if (obj.hasOwnProperty(k)) total += obj[k];
      }
      return total;
    }

    return {
      presents: presents,
      mineurs_moins_15: mineurs_moins_15,
      filles: filles,
      nouveaux: nouveaux,

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

      orphelins: orphelins,

      config_lists: configLists,

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
    
    let start = parseDateSecure(startStr);
    let end = parseDateSecure(endStr);
    end.setHours(23, 59, 59); 

    const dataJeunes = wsJeunes.getDataRange().getValues();
    const dataEvents = wsEvents.getDataRange().getValues();

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
      vulnerabilites_periode: {}
    };

    for (let i = 1; i < dataJeunes.length; i++) {
      let dateCreation = parseDateSecure(dataJeunes[i][18]);
      let statutPres = String(dataJeunes[i][9] || "").toLowerCase().trim();
      let genre = String(dataJeunes[i][6]).toUpperCase();
      let vulns = String(dataJeunes[i][15]).toLowerCase();
      let age = dataJeunes[i][5];

      if (statutPres.includes("présent") || statutPres.includes("mab (foyer)")) {
        stats.total_actifs++;
        
        if (age && parseInt(age) < 15) stats.mineurs_moins_15++;
        if (genre === "F" || genre === "FEMININ") stats.filles++;
        if (vulns.includes("emprise")) stats.sous_emprise++;
        
        if(vulns) {
          vulns.split(',').forEach(v => {
            let vt = v.trim();
            if(vt) stats.vulnerabilites[vt] = (stats.vulnerabilites[vt] || 0) + 1;
          });
        }
      }

      if (dateCreation && dateCreation >= start && dateCreation <= end) {
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

    for (let j = 1; j < dataEvents.length; j++) {
      let dateEvt = parseDateSecure(dataEvents[j][2]);
      if (!dateEvt) continue;
      if (dateEvt >= start && dateEvt <= end) {
        let type = String(dataEvents[j][3]).toLowerCase();
        let details = String(dataEvents[j][6]).toLowerCase();
        let motif = String(dataEvents[j][9]).toLowerCase();
        let materiel = String(dataEvents[j][11]).toLowerCase();
        let idJeune = dataEvents[j][1];
        let newVulnsEvt = String(dataEvents[j][10] || "").toLowerCase();

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
        
        if (newVulnsEvt) {
          newVulnsEvt.split(',').forEach(v => {
            let vt = v.trim();
            if(vt) stats.vulnerabilites_periode[vt] = (stats.vulnerabilites_periode[vt] || 0) + 1;
          });
        }
      }
    }

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
  
  const aujourdhui = formatDate(new Date());
  
  let chronoHist = [...hist].reverse();
  let premierEvt = chronoHist.length > 0 ? chronoHist[0].date : "Inconnue";
  
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
// ============================================================================
//          PHASE 3 : WIDGET CONFIG ITEMS POUR PLAIDOYER
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
//  CHERCHER dans Code.gs la fonction computeAllWidgets (ajoutée en Phase 3)
//  ET LA REMPLACER ENTIÈREMENT par celle-ci.
//  
//  La fonction getWidgetConfigItems() reste inchangée.
// ============================================================================

function computeAllWidgets(widgets, startStr, endStr) {
  var stats = getStatisticsDetailed(startStr, endStr);
  var subcards = { 'presents': 'Présents', 'mineurs_moins_15': 'Moins de 15 ans', 'filles': 'Filles', 'nouveaux': 'Nouveaux' };
  var results = {};
  var ss = SpreadsheetApp.openById(SHEET_ID);
  
  // Charger les données une seule fois
  var wsJ = ss.getSheetByName('BDD_JEUNES');
  var wsE = ss.getSheetByName('BDD_EVENTS');
  var dJ = wsJ.getDataRange().getValues();
  var dE = wsE.getDataRange().getValues();
  var start = parseDateSecure(startStr);
  var end = parseDateSecure(endStr);
  end.setHours(23, 59, 59);
  
  var jeunesFields = { 
    'statut_presence': 9, 'statut_mab': 10, 'nationalite': 7, 'genre': 6, 
    'lieu_vie': 14, 'vulnerabilites': 15, 'langue': 8, 'age': 5, 'labels': 20 
  };
  var eventsFields = { 
    'type_evenement': 3, 'type_contact': 4, 'lieu_event': 5, 
    'motif_statut_mab': 9, 'materiel': 11 
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
        results[widget.id] = { items: listResults };
        
      } else if (widget.type === 'custom_stat') {
        // NOUVEAU : Support conditions multiples (jeunes ET/OU events)
        var conditions = widget.conditions || [];
        if (conditions.length === 0 && widget.query) {
          // Rétrocompatibilité ancien format single query
          conditions = [widget.query];
        }
        
        // Séparer les conditions jeunes et events
        var condJeunes = [];
        var condEvents = [];
        conditions.forEach(function(c) {
          if (c.source === 'events') condEvents.push(c);
          else condJeunes.push(c);
        });
        
        // Calculer les IDs de jeunes matchant les conditions jeunes
        var jeuneIdsPool = null; // null = pas de filtre jeunes (tous)
        
        if (condJeunes.length > 0) {
          jeuneIdsPool = new Set();
          for (var i = 1; i < dJ.length; i++) {
            var matchAll = true;
            
            // Filtre actifs uniquement (si au moins une condition le demande)
            var needActive = condJeunes.some(function(c) { return c.activeOnly; });
            if (needActive) {
              var sP = String(dJ[i][9] || "").toLowerCase();
              if (!(sP.includes("présent") || sP.includes("mab (foyer)") || sP.includes("hospitalisé"))) {
                continue;
              }
            }
            
            for (var ci = 0; ci < condJeunes.length; ci++) {
              var c = condJeunes[ci];
              var colIdx = jeunesFields[c.field];
              if (colIdx === undefined) { matchAll = false; break; }
              var cV = String(dJ[i][colIdx] || "").toLowerCase();
              var qV = String(c.value || "").toLowerCase();
              
              if (!matchCondition(cV, c.operator, qV)) {
                matchAll = false;
                break;
              }
            }
            if (matchAll) jeuneIdsPool.add(dJ[i][0]);
          }
        }
        
        // Si pas de conditions events, le résultat = taille du pool jeunes
        if (condEvents.length === 0) {
          var count = jeuneIdsPool ? jeuneIdsPool.size : 0;
          var total = 0;
          for (var t = 1; t < dJ.length; t++) total++;
          var pct = total > 0 ? Math.round((count / total) * 100) : 0;
          results[widget.id] = { value: count, total: total, percentage: pct };
        } else {
          // Filtrer les events par conditions events + pool jeunes
          var matchedJeunes = new Set();
          var totalEventsChecked = 0;
          var matchedEvents = 0;
          
          for (var j = 1; j < dE.length; j++) {
            var dateEvt = parseDateSecure(dE[j][2]);
            if (!dateEvt || dateEvt < start || dateEvt > end) continue;
            
            // Si on a un pool jeunes, vérifier que cet event concerne un jeune du pool
            var evtJeuneId = dE[j][1];
            if (jeuneIdsPool && !jeuneIdsPool.has(evtJeuneId)) continue;
            
            totalEventsChecked++;
            
            var matchAllEvt = true;
            for (var ei = 0; ei < condEvents.length; ei++) {
              var ce = condEvents[ei];
              var eColIdx = eventsFields[ce.field];
              if (eColIdx === undefined) { matchAllEvt = false; break; }
              var eV = String(dE[j][eColIdx] || "").toLowerCase();
              var eqV = String(ce.value || "").toLowerCase();
              
              if (!matchCondition(eV, ce.operator, eqV)) {
                matchAllEvt = false;
                break;
              }
            }
            if (matchAllEvt) {
              matchedJeunes.add(evtJeuneId);
              matchedEvents++;
            }
          }
          
          // Le mode résultat : nombre de jeunes uniques matchant
          var useJeuneCount = widget.resultMode !== 'events';
          var finalCount = useJeuneCount ? matchedJeunes.size : matchedEvents;
          var finalTotal = useJeuneCount ? (jeuneIdsPool ? jeuneIdsPool.size : matchedJeunes.size) : totalEventsChecked;
          var finalPct = finalTotal > 0 ? Math.round((finalCount / finalTotal) * 100) : 0;
          
          results[widget.id] = { 
            value: finalCount, 
            total: finalTotal, 
            percentage: finalPct,
            jeunesCount: matchedJeunes.size,
            eventsCount: matchedEvents
          };
        }
      }
    } catch(e) {
      results[widget.id] = { value: 0, error: e.toString() };
    }
  });
  
  results['_fullStats'] = stats;
  return results;
}

// Fonction utilitaire pour évaluer une condition
function matchCondition(cellValue, operator, queryValue) {
  switch(operator) {
    case 'equals': return cellValue === queryValue;
    case 'contains': return cellValue.includes(queryValue);
    case 'not_contains': return !cellValue.includes(queryValue);
    case 'not_empty': return cellValue !== "";
    case 'empty': return cellValue === "";
    case 'less_than': return parseInt(cellValue) < parseInt(queryValue);
    case 'greater_than': return parseInt(cellValue) > parseInt(queryValue);
    default: return false;
  }
}

// Récupère les données d'un jeune épinglé (carte résumée pour dashboard)
function getDashboardJeuneCard(idJeune) {
  const ws = SpreadsheetApp.openById(SHEET_ID).getSheetByName('BDD_JEUNES');
  const data = ws.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == idJeune) {
      return {
        success: true,
        jeune: {
          id: data[i][0],
          nom: data[i][3],
          surnom: data[i][4],
          age: data[i][5],
          tel: data[i][2],
          nationalite: data[i][7],
          lieu_vie: data[i][14],
          statut_presence: data[i][9],
          statut_mab: data[i][10],
          vulnerabilites: data[i][15] || "",
          date_contact: formatDate(data[i][11])
        }
      };
    }
  }
  return { success: false, message: "Jeune introuvable" };
}

// Récupère un résumé des rappels pour le widget dashboard
function getDashboardRappelsSummary() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const wsRappels = ss.getSheetByName('BDD_RAPPELS');
  const wsJeunes = ss.getSheetByName('BDD_JEUNES');
  
  const dataRappels = wsRappels.getDataRange().getValues();
  const dataJeunes = wsJeunes.getDataRange().getValues();
  
  let jeunesMap = {};
  for(let i=1; i<dataJeunes.length; i++) {
    jeunesMap[dataJeunes[i][0]] = { nom: dataJeunes[i][3] };
  }
  
  const today = new Date(); 
  today.setHours(0,0,0,0);
  let enRetard = 0;
  let aFaire = 0;
  let prochains = []; // Les 5 prochains rappels

  for (let i = 1; i < dataRappels.length; i++) {
    try {
      let statut = String(dataRappels[i][5] || "").toLowerCase().trim();
      if (statut.indexOf("faire") !== -1) {
        let echeance = parseDateSecure(dataRappels[i][4]);
        let idJeune = dataRappels[i][0];
        let jInfo = jeunesMap[idJeune] || { nom: "Inconnu" };
        
        if (echeance < today) {
          enRetard++;
        } else {
          aFaire++;
        }
        
        if (prochains.length < 8) {
          prochains.push({
            idRappel: dataRappels[i][1],
            idJeune: idJeune,
            nomJeune: jInfo.nom,
            titre: String(dataRappels[i][6] || "Sans titre"),
            date: formatDate(echeance),
            isLate: echeance < today
          });
        }
      }
    } catch(e) {}
  }
  
  // Trier: en retard d'abord
  prochains.sort(function(a,b) {
    if (a.isLate && !b.isLate) return -1;
    if (!a.isLate && b.isLate) return 1;
    return 0;
  });

  return { enRetard: enRetard, aFaire: aFaire, prochains: prochains };
}

// ============================================================================
// PHASE 4 FIX : Stockage préférences PAR UTILISATEUR TAMABochi (spreadsheet)
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
