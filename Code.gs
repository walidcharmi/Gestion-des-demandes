// Définir les IDs des fichiers associés
var ACTIVITE_ESTUDIANTINE = '1f1Be7Lq-larL_Jva8KEDdsiNkJhxmMbj2GndITPySJk';
var DEMANDE_INTERVENTION = '1Bt_GLCPjMmQgKjamFu83joLOjKr1Vo9TH9-kpRTCUto';
var DEMANDE_MATERIEL_ETUDIANT = '1l2J5Fn4RLKdWkJZOrc16CRzPrP1TvxhFG12pFBULrv4';
var DEMANDE_MATERIEL_ENSEIGNANT_ADMIN = '15H3ZAVjT6G2IBAsvKHuEn6UmdK_WqIR-G04O-yiefqw';
var DEPOT_PFE = '1jyn_qKXCBze7dv3tA794oL5TnnfXXKuKVxCyItCqv4k';
var CERTIFICATION_MICROSOFT = '1xbbuBJEdH2p3Si6Bt92FWlBgQOSMV1Bwa2rFfCXF7Eo';
var COMPTE_INSTITUTIONNEL = '1Y-50gArs_Y9uZEY4g73OQu-QLbGZGyyzK49_n-KrZqU';
var DOSSIER_ACTIVITE_ESTUDIANTINE = '1mJ3SkX3HpnA0FXgkNS3EtzykZtkx8qEr';
var DEMANDES_SERVICE = '1bUwhRlfMdcinfd5Kw8LTtZc6x3V8a7Ob2hl_395AOPc';
var DOSSIER_DEPOT_PFE = '1cZ1EYcryAI_nnIZqxvsGvx9Sfd-UPR6_';

// Définir les dates de début et de fin pour la vérification des demandes
var START_DATE = new Date('2024-05-17T09:13:00'); // Exemple de date de début
var END_DATE = new Date('2024-12-27T08:41:40'); // Exemple de date de fin

// Définir le nombre maximum de réponses
var MAX_RESPONSES = 1;

function getCertificationMicrosoftCount() {
  var ss = SpreadsheetApp.openById(CERTIFICATION_MICROSOFT);
  var sheet = ss.getSheetByName('Feuille 1');
  var data = sheet.getDataRange().getValues();
  var count = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var timestamp = new Date(row[0]);
    if (timestamp >= START_DATE && timestamp <= END_DATE) {
      count++;
    }
  }

  return count;
}

function updateCertificationMicrosoftOption() {
  var currentDate = new Date();
  var certificationCount = getCertificationMicrosoftCount();

  if ((currentDate < START_DATE || currentDate > END_DATE) || certificationCount >= MAX_RESPONSES) {
    return false; // Inactif
  } else {
    return true; // Actif
  }
}

function sendResetPasswordEmail(email, nom, prenom, telephone, fonction) {
  const recipient = 'walid.charmi@enetcom.usf.tn';
  const subject = 'Demande d\'accès';

  // Créer un lien mailto avec l'e-mail pré-rempli
  const mailtoLink = `mailto:${email}?subject=Acc%C3%A8s%20Plateforme%20ENET'Com&body=Bonjour%20${prenom}%20${nom}%2C%0A%0AVeuillez%20trouver%20ci-dessous%20le%20mot%20de%20passe%20vous%20permettant%20d'acc%C3%A9der%20%C3%A0%20la%20plateforme%20de%20gestion%20des%20demandes%20des%20parties%20prenantes%20de%20l'ENET'Com%20:%0A%0Aenetcom%40ens%0A%0ACordialement%2C%0A`;

  const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Demande d'accès</title>
  <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700&display=swap" rel="stylesheet">
</head>
<body style="margin: 0; padding: 0; font-family: 'Montserrat', Arial, sans-serif; background-color: #F9FAFB; color: #1F2937;">
  <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 16px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06); margin-top: 20px; margin-bottom: 20px;">
    <!-- En-tête -->
    <div style="background: linear-gradient(120deg, #0EA5E9 0%, #0284C7 60%, #0369A1 100%); padding: 40px 30px; text-align: center; color: white;">
      <h1 style="margin: 0; font-size: 24px; font-weight: 700; letter-spacing: -0.5px;">Demande d'accès</h1>
      <p style="margin: 12px 0 0 0; font-size: 16px; font-weight: 300; letter-spacing: 1px; text-transform: uppercase;">Nouvelle notification</p>
    </div>

    <!-- Contenu principal -->
    <div style="padding: 32px 24px;">
      <p style="font-size: 16px; line-height: 1.6; color: #1F2937;">Bonjour,</p>
      <p style="font-size: 16px; line-height: 1.6; color: #374151;">Une demande d'accès à la plateforme de gestion des demandes des parties prenantes de l'ENET'Com a été reçue avec les détails suivants :</p>

      <!-- Détails de la demande -->
      <div style="background-color: #ffffff; border-radius: 12px; padding: 24px; margin: 24px 0; border: 1px solid #E5E7EB;">
        <div style="margin: 16px 0;">
          <p style="margin: 4px 0; font-size: 14px; color: #6B7280; font-weight: 500;">Nom & prénom :</p>
          <p style="margin: 0; font-size: 16px; color: #111827; font-weight: 600;">${nom} ${prenom}</p>
        </div>
        <div style="margin: 16px 0;">
          <p style="margin: 4px 0; font-size: 14px; color: #6B7280; font-weight: 500;">Fonction :</p>
          <p style="margin: 0; font-size: 16px; color: #111827; font-weight: 600;">${fonction}</p>
        </div>
        <div style="margin: 16px 0;">
          <p style="margin: 4px 0; font-size: 14px; color: #6B7280; font-weight: 500;">Téléphone :</p>
          <p style="margin: 0; font-size: 16px; color: #111827; font-weight: 600;">${telephone}</p>
        </div>
        <div style="margin: 16px 0;">
          <p style="margin: 4px 0; font-size: 14px; color: #6B7280; font-weight: 500;">E-mail :</p>
          <p style="margin: 0; font-size: 16px; color: #111827; font-weight: 600;">${email}</p>
        </div>
      </div>

      <!-- Bouton Envoyer un e-mail -->
      <div style="text-align: center; margin-bottom: 24px;">
        <a href="${mailtoLink}" style="display: inline-block; background-color: #0EA5E9; color: white; padding: 12px 24px; text-decoration: none; border-radius: 8px; font-weight: 600;">
          Envoyer un e-mail
        </a>
      </div>

      <div style="margin-top: 30px;">
        <p style="font-size: 16px; color: #374151; margin-bottom: 16px;">Cordialement,</p>
        <p>&nbsp;</p>
        <p>&nbsp;</p>
      </div>
    </div>

    <!-- Pied de page -->
    <div style="background-color: #4761a1; padding: 16px; text-align: center; font-size: 12px; color: #F3F4F6;">
      <p style="margin: 0;">Walid CHARMI<br>Service informatique - ENET'Com</p>
    </div>
  </div>
</body>
</html>
`;

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: htmlBody,
    name: 'Walid CHARMI'
  });
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Gestion des demandes des parties prenantes de l\'ENET\'Com');
}

// Fonction pour obtenir la source de l'appareil
function getDeviceSource() {
  var userAgent = HtmlService.getUserAgent();
  
  // Analyse de l'user-agent pour identifier l'appareil
  if (userAgent.indexOf('Mobile') > -1) {
    if (userAgent.indexOf('iPhone') > -1 || userAgent.indexOf('iPad') > -1) {
      return 'iOS';
    } else if (userAgent.indexOf('Android') > -1) {
      return 'Android';
    } else {
      return 'Mobile autre';
    }
  } else {
    if (userAgent.indexOf('Windows') > -1) {
      return 'Windows';
    } else if (userAgent.indexOf('Macintosh') > -1) {
      return 'MacOS';
    } else if (userAgent.indexOf('Linux') > -1) {
      return 'Linux';
    } else {
      return 'Desktop autre';
    }
  }
}

function processFormData(jsonData) {
  var formObject = JSON.parse(jsonData);
  var deviceSource = getDeviceSource();

  // Enregistrement des informations dans DEMANDES_SERVICE avec la source
  var ssService = SpreadsheetApp.openById(DEMANDES_SERVICE);
  var serviceSheet = ssService.getSheetByName('Feuille 1') || ssService.getSheets()[0];
  
  var serviceValues = [
    deviceSource,
    new Date(),
    "'" + (formObject.nationalite === 'Tunisienne' ? formObject.cin : formObject.codeMinistere),
    formObject.nom,
    formObject.prenom,
    formObject.telephone,
    formObject.email,
    formObject.fonction,
    formObject.typeService,
  ];
  
  // Ajouter les valeurs à la feuille
  serviceSheet.appendRow(serviceValues);
  var lastRow = serviceSheet.getLastRow();
  
  // Appliquer les styles de police et d'alignement
  var range = serviceSheet.getRange(lastRow, 1, 1, serviceSheet.getLastColumn());
  range.setFontFamily('Calibri').setFontSize(11).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);  

  // Fonction pour générer les valeurs communes
  function getCommonValues() {
    return [
      new Date(), // Timestamp
      formObject.nationalite,
      formObject.nationalite === 'Tunisienne' ? "'" + formObject.cin : "'" + formObject.codeMinistere,
      formObject.nom,
      formObject.prenom,
      formObject.telephone,
      formObject.email,
    ];
  }

  var commonValues = getCommonValues();

  // Traitement spécial pour "Dépôt final – PFE et Thèse"
// Traitement spécial pour "Dépôt final – PFE et Thèse"
if (formObject.typeService === "Dépôt final – PFE et Thèse") {
    var ss = SpreadsheetApp.openById(DEPOT_PFE);
    var depotSheet;
    // Sélectionner la feuille en fonction de la formation et de la filière
    switch(formObject.formation) {
        case 'Ingénieur':
            depotSheet = ss.getSheetByName('Ingénieur');
            break;
        case 'Licence Unifiée':
            depotSheet = ss.getSheetByName('Licence Unifiée');
            break;
        case 'Mastère':
            // Sélectionner la feuille en fonction de la filière pour Mastère
            switch(formObject.filiere) {
                case 'MR':
                    depotSheet = ss.getSheetByName('MR');
                    break;
                case 'MP (Formation initiale)':
                    depotSheet = ss.getSheetByName('MP (Formation initiale)');
                    break;
                case 'MP (Formation continue)':
                    depotSheet = ss.getSheetByName('MP (Formation continue)');
                    break;
                default:
                    throw new Error("Filière non reconnue pour Mastère: " + formObject.filiere);
            }
            break;
        case 'Doctorat en STIC':
            depotSheet = ss.getSheetByName('Doctorat en STIC');
            break;
        default:
            throw new Error("Formation non reconnue: " + formObject.formation);
    }

    // Créer un dossier nommé par la carte d'identité nationale ou le code ministère
    var folderName = formObject.nationalite === 'Tunisienne' ? formObject.cin : formObject.codeMinistere;
    var parentFolder = DriveApp.getFolderById(DOSSIER_DEPOT_PFE);
    var userFolder = parentFolder.createFolder(folderName);

    // Formater la date de soutenance en fonction de la langue
    var dateSoutenance;
    if (formObject.langueProjet === 'Français') {
        dateSoutenance = formObject.dateSoutenance ? Utilities.formatDate(new Date(formObject.dateSoutenance), Session.getScriptTimeZone(), "dd/MM/yyyy") : '';
    } else if (formObject.langueProjet === 'Anglais') {
        dateSoutenance = formObject.defenseDate ? Utilities.formatDate(new Date(formObject.defenseDate), Session.getScriptTimeZone(), "dd/MM/yyyy") : '';
    }

    // Initialiser les liens des fichiers
    var rapportCompletLink = '';
    var tableMatiereLink = '';
    var pageGardeResumeLink = '';

    // Gérer les fichiers joints pour "Dépôt final – PFE et Thèse"
    if (formObject.rapportComplet) {
        var rapportCompletBlob = Utilities.newBlob(Utilities.base64Decode(formObject.rapportComplet.split(',')[1]), 'application/pdf', 'RapportComplet.pdf');
        var rapportCompletFile = userFolder.createFile(rapportCompletBlob);
        rapportCompletLink = rapportCompletFile.getUrl();
    }

    if (formObject.tableMatiere) {
        var tableMatiereBlob = Utilities.newBlob(Utilities.base64Decode(formObject.tableMatiere.split(',')[1]), 'application/pdf', 'TableMatiere.pdf');
        var tableMatiereFile = userFolder.createFile(tableMatiereBlob);
        tableMatiereLink = tableMatiereFile.getUrl();
    }

    if (formObject.pageGardeResume) {
        var pageGardeResumeBlob = Utilities.newBlob(Utilities.base64Decode(formObject.pageGardeResume.split(',')[1]), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'PageGardeResume.docx');
        var pageGardeResumeFile = userFolder.createFile(pageGardeResumeBlob);
        pageGardeResumeLink = pageGardeResumeFile.getUrl();
    }

    // Combiner les valeurs avec des sauts de ligne
    var nomComplet = formObject.nom;
    var prenomComplet = formObject.prenom;
    var telephoneComplet = formObject.telephone;
    var emailComplet = formObject.email;
    var formationComplet = formObject.formation;
    var niveauFiliereComplet = formObject.niveau + ' ' + formObject.filiere;
    var nationaliteComplet = formObject.nationalite;

    // Gestion des identifiants combinés
    var identifiantComplet;
    if (formObject.travailEnBinome === "Non") {
        identifiantComplet = formObject.nationalite === 'Tunisienne' ? formObject.cin : formObject.codeMinistere;
    } else {
        if (formObject.nationalite === 'Tunisienne' && formObject.binomeNationalite === 'Tunisienne') {
            identifiantComplet = formObject.cin + '\n' + formObject.binomeCin;
        } else if (formObject.nationalite === 'Étrangère' && formObject.binomeNationalite === 'Étrangère') {
            identifiantComplet = formObject.codeMinistere + '\n' + formObject.binomeCodeMinistere;
        } else if (formObject.nationalite === 'Tunisienne' && formObject.binomeNationalite === 'Étrangère') {
            identifiantComplet = formObject.cin + '\n' + formObject.binomeCodeMinistere;
        } else if (formObject.nationalite === 'Étrangère' && formObject.binomeNationalite === 'Tunisienne') {
            identifiantComplet = formObject.codeMinistere + '\n' + formObject.binomeCin;
        }
    }

    if (formObject.binomeNom) {
        nomComplet += '\n' + formObject.binomeNom;
    }

    if (formObject.binomePrenom) {
        prenomComplet += '\n' + formObject.binomePrenom;
    }

    if (formObject.binomeTelephone) {
        telephoneComplet += '\n' + formObject.binomeTelephone;
    }

    if (formObject.binomeEmail) {
        emailComplet += '\n' + formObject.binomeEmail;
    }

    if (formObject.binomeFormation) {
        formationComplet += '\n' + formObject.binomeFormation;
    }

    if (formObject.binomeNiveau && formObject.binomeFiliere) {
        niveauFiliereComplet += '\n' + formObject.binomeNiveau + ' ' + formObject.binomeFiliere;
    }

    if (formObject.binomeNationalite) {
        nationaliteComplet += '\n' + formObject.binomeNationalite;
    }

    var commonValues = [
        new Date(), // Timestamp
        nationaliteComplet,
        identifiantComplet,
        nomComplet,
        prenomComplet,
        telephoneComplet,
        emailComplet,
        formationComplet,
        niveauFiliereComplet,
    ];

    var depotValues = commonValues.concat([
        formObject.encadrant1 || '', // Colonne J
        formObject.encadrant2 || '', // Colonne K
        formObject.encadrant3 || '', // Colonne L
        formObject.nomSociete || '', // Colonne M
        formObject.langueProjet,
        dateSoutenance,
        formObject.titreProjet || formObject.projectTitle,
        formObject.motsCles || formObject.keywords,
        rapportCompletLink,
        tableMatiereLink,
        pageGardeResumeLink
    ]);

    // Ajouter les valeurs à la feuille
    depotSheet.appendRow(depotValues);
    var lastRow = depotSheet.getLastRow();

    // Appliquer les styles de police et d'alignement
    var range = depotSheet.getRange(lastRow, 1, 1, depotSheet.getLastColumn());
    range.setFontFamily('Calibri').setFontSize(11).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);

    // Appliquer un fond bleu clair si "Travail en binôme" est "Oui"
    if (formObject.travailEnBinome === "Oui") {
        range.setBackground('#d4f1ff'); // Code couleur pour bleu clair
    }

    // Utiliser la liste des encadrants envoyée depuis le HTML
    var encadrants = formObject.encadrants || [];

    // Liste des encadrants et des cellules correspondantes
    const encadrantsList = [
        { encadrant: formObject.encadrant1, cell: depotSheet.getRange(lastRow, 10) }, // Colonne J
        { encadrant: formObject.encadrant2, cell: depotSheet.getRange(lastRow, 11) }, // Colonne K
        { encadrant: formObject.encadrant3, cell: depotSheet.getRange(lastRow, 12) }  // Colonne L
    ];

    // Parcourir la liste des encadrants et appliquer le style de couleur rouge si nécessaire
    encadrantsList.forEach(item => {
        if (item.encadrant && !encadrants.includes(item.encadrant)) {
            item.cell.setFontColor('#FF0000'); // Couleur rouge
        }
    });
  }
  
  // Traitement spécial pour "Demande d'activité estudiantine"
  if (formObject.typeService === "Demande d'activité estudiantine") {
    var ss = SpreadsheetApp.openById(ACTIVITE_ESTUDIANTINE);
    var activiteSheet = ss.getSheetByName('Feuille 1');

    var fileUrl = '';
    if (formObject.fichierJointContenu && formObject.fichierJointNom) {
      var fileBlob = Utilities.newBlob(Utilities.base64Decode(formObject.fichierJointContenu.split(',')[1]), 'application/pdf', formObject.fichierJointNom);
      var folder = DriveApp.getFolderById(DOSSIER_ACTIVITE_ESTUDIANTINE);
      var file = folder.createFile(fileBlob);
      fileUrl = file.getUrl();
    }

    var activiteValues = commonValues.concat([
      formObject.fonction === 'Étudiant' ? formObject.formation : formObject.departement,
      formObject.fonction === 'Étudiant' ? (formObject.niveau + ' ' + formObject.filiere) : formObject.grade,
      formObject.nomClub,
      formObject.sujetDemande,
      formObject.descriptionDetailleActivite,
      fileUrl
    ]);

    // Ajouter les valeurs à la feuille
    activiteSheet.appendRow(activiteValues);
    var lastRow = activiteSheet.getLastRow();

    // Appliquer les styles de police et d'alignement à toute la plage de cellules de la dernière ligne
    var range = activiteSheet.getRange(lastRow, 1, 1, activiteSheet.getLastColumn());
    range.setFontFamily('Calibri').setFontSize(11).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);

    // Ajouter une case à cocher dans la colonne "O" avec les styles de police et de taille
    var colonneO = activiteSheet.getRange(lastRow, 15);
    colonneO.insertCheckboxes()
                 .setHorizontalAlignment('center');

    // Envoyer la notification par e-mail
    sendStudentActivityRequestNotification(formObject);               
  }

  // Traitement spécial pour "Demande d'intervention"
  if (formObject.typeService === "Demande d'intervention") {
    var ss = SpreadsheetApp.openById(DEMANDE_INTERVENTION);
    var interventionSheet = ss.getSheetByName('Feuille 1');

    var interventionValues = commonValues.concat([
      formObject.fonction,
      formObject.fonction === 'Étudiant' ? formObject.formation : formObject.departement,
      formObject.fonction === 'Étudiant' ? (formObject.niveau + ' ' + formObject.filiere) : formObject.grade,
      formObject.natureDemande,
      formObject.lieuImplantation,
      formObject.descriptionDetailleeIntervention
    ]);

    Logger.log(interventionValues); // Ajouter cette ligne pour vérifier les valeurs ajoutées

    // Ajouter les valeurs à la feuille
    interventionSheet.appendRow(interventionValues);
    var lastRow = interventionSheet.getLastRow();

    // Appliquer les styles de police et d'alignement à toute la plage de cellules de la dernière ligne
    var range = interventionSheet.getRange(lastRow, 1, 1, interventionSheet.getLastColumn());
    range.setFontFamily('Calibri').setFontSize(11).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);

    // Ajouter une case à cocher dans la colonne Q et la centrer
    var colonneQ = interventionSheet.getRange(lastRow, 17);
    colonneQ.insertCheckboxes()
            .setHorizontalAlignment('center');

    // Ajouter une liste dans la colonne P
    var colonneP = interventionSheet.getRange(lastRow, 16);
    var validationP = SpreadsheetApp.newDataValidation().requireValueInList([
      'M. Walid CHARMI', 'M. Mourad BENJEDDOU', 'Mme Feten OUSSAIFI', 'M. Hassen BACCAR', 'Mme Emna MEGDICH'
    ], true).build();
    colonneP.setDataValidation(validationP);

    // Envoyer la notification par e-mail
    sendInter3t24NpUrJMNunMMASmhAM953bFGeLXzN7(formObject);
  }

  // Traitement spécial pour "Demande de matériels"
  if (formObject.typeService === "Demande de matériels") {
    var ss;
    var materielSheet;
    var materielValues;

    if (formObject.fonction === 'Étudiant') {
      ss = SpreadsheetApp.openById(DEMANDE_MATERIEL_ETUDIANT);
      materielSheet = ss.getSheetByName('Étudiant');
      materielValues = commonValues.concat([
        formObject.fonction,
        formObject.formation,
        formObject.niveau + ' ' + formObject.filiere,
        formObject.deuxiemeEtudiant,
        formObject.departementMateriel,
        formObject.typeProjet,
        formObject.nomProjet,
        formObject.encadreur
      ]);

      // Formatage des matériels pour la colonne P avec des parenthèses autour des quantités
      var materielsFormatted = formObject.materiels.map(function(materiel) {
        return '(' + materiel.quantite + ') ' + materiel.designation;
      }).join("\n");

      materielValues.push(materielsFormatted);

      // Ajouter des valeurs vides pour les colonnes Q à V
      materielValues = materielValues.concat(['', '', '', '', '', '']);
      // Ajouter "Non rendu" dans la colonne W
      materielValues.push('Non rendu');
      // Ajouter une valeur vide pour la colonne X
      materielValues.push('');

      materielSheet.appendRow(materielValues);
      var lastRow = materielSheet.getLastRow();

      // Envoyer la notification par e-mail
      sendMaterialRequestNotification(formObject);      

      // Appliquer les styles de police et d'alignement
      var range = materielSheet.getRange(lastRow, 1, 1, materielSheet.getLastColumn());
      range.setFontFamily('Calibri').setFontSize(11).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);

      // Insérer une case à cocher dans la colonne T et la centrer
      var colonneT = materielSheet.getRange(lastRow, 20);
      colonneT.insertCheckboxes()
        .setHorizontalAlignment('center');

      // Ajouter une liste dans la colonne Q
      var colonneQ = materielSheet.getRange(lastRow, 17);
      var validationQ = SpreadsheetApp.newDataValidation().requireValueInList(['M. Karim JABER', 'M. Hassen BACCAR', 'M. Walid CHARMI', 'M. Mourad BENJEDDOU', 'Mme Emna MEGDICH', 'Mme Feten OUSSAIFI'], true).build();
      colonneQ.setDataValidation(validationQ);

      // Ajouter une liste dans la colonne W avec "Rendu" en vert et "Non rendu" en rouge
      var colonneW = materielSheet.getRange(lastRow, 23);
      var validationW = SpreadsheetApp.newDataValidation().requireValueInList(['Rendu', 'Non rendu'], true).build();
      colonneW.setDataValidation(validationW);

      // Ajouter un agenda dans la colonne X pour insérer une date
      var colonneX = materielSheet.getRange(lastRow, 24);
      var validationX = SpreadsheetApp.newDataValidation().requireDate().build();
      colonneX.setDataValidation(validationX);

      // Appliquer le formatage conditionnel pour la colonne W
      var rulesW = materielSheet.getConditionalFormatRules();

      var ruleVRendu = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Rendu')
        .setBackground('#00FF00')
        .setRanges([colonneW])
        .build();

      var ruleVNotYet = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo('Non rendu')
        .setBackground('#FF0000')
        .setRanges([colonneW])
        .build();

      rulesW.push(ruleVRendu);
      rulesW.push(ruleVNotYet);
      materielSheet.setConditionalFormatRules(rulesW);

    } else if (formObject.fonction === 'Enseignant' || formObject.fonction === 'Administratif') {
      ss = SpreadsheetApp.openById(DEMANDE_MATERIEL_ENSEIGNANT_ADMIN);

  // Déterminer la feuille en fonction de la catégorie
  if (formObject.categorie === 'Consommables') {
    materielSheet = ss.getSheetByName('Consommables');
  } else if (formObject.categorie === 'Matériels') {
    materielSheet = ss.getSheetByName('Matériels');
    // Envoyer un e-mail de notification à Walid CHARMI
    sendMaterielsNotificationToWalid(formObject);
  } else {
    throw new Error("Catégorie non reconnue : " + formObject.categorie);
  }

    // Collecte des valeurs communes
    var justificationValue = (formObject.justification === 'Autre') ? formObject.precisezJustification : formObject.justification;
    materielValues = commonValues.concat([
        formObject.fonction,
        formObject.departement,
        formObject.grade,
        justificationValue,
    ]);

    // Formatage des matériels pour inclure la quantité, la désignation et le prix
    var materielsFormatted = formObject.materiels.map(function(materiel) {
        return '(' + materiel.quantite + ') ' + materiel.designation + ' - ' + materiel.prix + ' DT';
    }).join("\n");

    materielValues.push(materielsFormatted);

    // Ajouter des valeurs vides pour les colonnes supplémentaires (si nécessaire)
    materielValues = materielValues.concat(['', '', '', '']);

    // Ajouter les valeurs à la feuille
    materielSheet.appendRow(materielValues);

    // Appliquer les styles de police et d'alignement
    var lastRow = materielSheet.getLastRow();
    var range = materielSheet.getRange(lastRow, 1, 1, materielSheet.getLastColumn());
    range.setFontFamily('Calibri').setFontSize(11).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);

}
  }

  // Traitement spécial pour "Certification Microsoft"
  if (formObject.typeService === "Certification Microsoft") {
    var ss = SpreadsheetApp.openById(CERTIFICATION_MICROSOFT);
    var certificationSheet = ss.getSheetByName('Feuille 1');
    var data = certificationSheet.getDataRange().getValues();

    // Vérification des demandes existantes
    var id = formObject.nationalite === 'Tunisienne' ? formObject.cin : formObject.codeMinistere;
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var timestamp = new Date(row[0]);
      var existingId = row[2]; // Colonne contenant CIN ou Code Ministère

      if (timestamp >= START_DATE && timestamp <= END_DATE && existingId == id) {
        throw new Error('Une demande a déjà été soumise.');
      }
    }

    // Ajouter la fonction (Étudiant, Enseignant, Administratif) dans la colonne H
    var certificationValues = commonValues.slice(0, 7).concat([
      formObject.fonction,
      formObject.fonction === 'Étudiant' ? formObject.formation : formObject.departement,
      formObject.fonction === 'Étudiant' ? (formObject.niveau + ' ' + formObject.filiere) : formObject.grade,
      formObject.langueExamen,
      formObject.typeCertification
    ]);

    certificationValues.push(formObject.exam);

    // Ajouter les valeurs à la feuille
    certificationSheet.appendRow(certificationValues);
    var lastRow = certificationSheet.getLastRow();

    // Appliquer les styles de police et d'alignement
    var range = certificationSheet.getRange(lastRow, 1, 1, certificationSheet.getLastColumn());
    range.setFontFamily('Calibri').setFontSize(11).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);

    // Insérer une case à cocher dans la colonne Q et la centrer
    var colonneQ = certificationSheet.getRange(lastRow, 17);
    colonneQ.insertCheckboxes()
            .setHorizontalAlignment('center');

    // Ajouter une liste dans la colonne S avec "Achraf MTIBAA" et "Walid CHARMI"
    var colonneS = certificationSheet.getRange(lastRow, 19);
    var validationS = SpreadsheetApp.newDataValidation().requireValueInList(['Achraf MTIBAA', 'Walid CHARMI'], true).build();
    colonneS.setDataValidation(validationS);

    // Envoyer un e-mail personnalisé à l'adresse e-mail personnelle de l'utilisateur
    var firstName = formObject.prenom;
    var lastName = formObject.nom;
    var email = formObject.email;

var emailSubject = 'Préparation à la certification Microsoft';
var emailBody = `
  <div style="font-family: 'Inter', system-ui, sans-serif; margin: 0; background: #ffffff; width: 100%; padding: 20px; box-sizing: border-box; font-size: 15px;">
    <!-- Main Content -->
    <div style="padding: 40px 32px; max-width: 100%; box-sizing: border-box;">
      <!-- Greeting -->
      <p style="color: #334155; margin: 0 0 16px 0; line-height: 1.6;">
        Bonjour ${firstName} ${lastName.toUpperCase()},
      </p>
      <p style="color: #334155; margin: 0 0 32px 0; line-height: 1.6;">
        Votre demande a bien été reçue avec les détails suivants :
      </p>
      <!-- Details List -->
      <div style="display: grid; grid-template-columns: 1fr; gap: 16px; margin-bottom: 40px;">
        <div style="background: #f7f5f7; border-radius: 16px; padding: 24px; border: 1px solid #F1F5F9; width: 100%; box-sizing: border-box; word-break: break-word; overflow-wrap: break-word; border-left: 4px solid #0EA5E9;">
          <div style="display: grid; grid-template-columns: 1fr; gap: 20px;">
            <div style="padding-left: 16px; width: 100%; box-sizing: border-box;">
              <p style="color: #94A3B8; text-transform: uppercase; letter-spacing: 0.05em; margin: 0 0 4px 0; font-size: 14px;">Langue de l'examen</p>
              <p style="color: #0F172A; font-weight: 600; margin: 0; font-size: 14px;">${formObject.langueExamen}</p>
              <br>
            </div>
            <div style="padding-left: 16px; width: 100%; box-sizing: border-box;">
              <p style="color: #94A3B8; text-transform: uppercase; letter-spacing: 0.05em; margin: 0 0 4px 0; font-size: 14px;">Type de certification</p>
              <p style="color: #0F172A; font-weight: 600; margin: 0; font-size: 14px;">${formObject.typeCertification}</p>
              <br>
            </div>
            <div style="padding-left: 16px; width: 100%; box-sizing: border-box;">
              <p style="color: #94A3B8; text-transform: uppercase; letter-spacing: 0.05em; margin: 0 0 4px 0; font-size: 14px;">Examen</p>
              <p style="color: #0F172A; font-weight: 600; margin: 0; font-size: 14px;">${formObject.exam}</p>
              <br>
            </div>
          </div>
        </div>
      </div>
      <!-- Note -->
      <div style="background: #FEF9C3; border-left: 4px solid #EAB308; padding: 16px 20px; margin-bottom: 32px; border-radius: 8px; word-break: break-word; overflow-wrap: break-word;">
        <p style="color: #854D0E; margin: 0; line-height: 1.6;">
          <ul style="list-style-type: none; padding: 0; margin: 0;">
            <li style="color: #854D0E; display: flex; align-items: flex-start; margin-bottom: 8px;">
              <span style="margin-right: 8px; font-size: 16px;">&#x2713;</span>
              <span style="flex: 1;">Pour bien se préparer à la certification, veuillez cliquer sur le lien suivant
                <a href="https://drive.google.com/drive/folders/1HypJLd84r_FydQWI6ikJNdxCMlRyy9WH?usp=drive_link" style="color: #0EA5E9; text-decoration: none;">
                  <span style="display: inline-block; font-family: 'Arial', sans-serif; color: #0078D7; border: 1px solid #0078D7; padding: 2px 4px; border-radius: 4px; white-space: nowrap;">
                    Réussir
                  </span>
                </a>
              </span>
            </li>
            <li style="color: #854D0E; display: flex; align-items: flex-start; margin-bottom: 8px;">
              <span style="margin-right: 8px; font-size: 16px;">&#x2713;</span>
              <span style="flex: 1;">Vous devez avoir un compte Certiport pour pouvoir passer votre examen de certification. Si vous n'avez pas encore de compte, veuillez en créer un
                <a href="https://certiport.pearsonvue.com/" style="color: #0EA5E9; text-decoration: none;">
                  <span style="display: inline-block; font-family: 'Arial', sans-serif; color: #0078D7; border: 1px solid #0078D7; padding: 2px 4px; border-radius: 4px; white-space: nowrap;">
                    S'inscrire
                  </span>
                </a>
              </span>
            </li>
            <li style="color: #854D0E; display: flex; align-items: flex-start;">
              <span style="margin-right: 8px; font-size: 16px;">&#x2713;</span>
              <span style="flex: 1;">Vous allez recevoir un message par e-mail contenant la date de passage du certificat, ainsi que toutes les informations nécessaires.</span>
            </li>
          </ul>
        </p>
      </div>
      <!-- Conclusion -->
      <p style="color: #334155; margin: 0 0 32px 0; line-height: 1.6;">
        Cordialement,
      </p>
    </div>
  </div>
        <!-- Signature -->
        <div style="font-family: 'Inter', system-ui, sans-serif; padding: 16px 20px; border-top: 1px solid #0EA5E9; margin-top: 20px; text-align: center;">
          <table style="width: 100%; max-width: 600px; margin: 0 auto; text-align: center;">
            <tr>
              <td>
                <h3 style="margin: 0; color: #0F172A; font-size: 15px; font-weight: 600;">Achraf MTIBAA & Walid CHARMI</h3>
                <p style="margin: 4px 0 0 0; color: #64748B; font-size: 13px; line-height: 1.4;">
                  Administrateurs centre Certiport 4C-ENET'Com
                </p>
              </td>
            </tr>
          </table>
        </div>
      </div>
    `;

    var emailOptions = {
      name: 'Centre Certiport 4C-ENET\'Com',
      htmlBody: emailBody,
      noReply: true
    };
    MailApp.sendEmail(email, emailSubject, '', emailOptions);
}

  // Traitement spécial pour "Compte institutionnel"
  if (formObject.typeService === "Compte institutionnel") {
    var ss = SpreadsheetApp.openById(COMPTE_INSTITUTIONNEL);
    var sheetName = '';

    // Déterminer la feuille en fonction du type de compte
    if (formObject.typeCompte === "Office 365") {
        sheetName = "Walid CHARMI";
    } else if (formObject.typeCompte === "Extranet" || formObject.typeCompte === "Gmail professionnel") {
        sheetName = "Emna MEGDICH";
    }

    var compteSheet = ss.getSheetByName(sheetName);

    var compteValues = commonValues.concat([
      formObject.fonction,
      formObject.fonction === 'Étudiant' ? formObject.formation : formObject.departement,
      formObject.fonction === 'Étudiant' ? (formObject.niveau + ' ' + formObject.filiere) : formObject.grade,
      formObject.typeCompte,
      formObject.natureDemandeCompte,
      formObject.nomUtilisateur || ''
    ]);

    // Ajouter les valeurs à la feuille
    compteSheet.appendRow(compteValues);
    var lastRow = compteSheet.getLastRow();

    // Appliquer les styles de police et d'alignement
    var range = compteSheet.getRange(lastRow, 1, 1, compteSheet.getLastColumn());
    range.setFontFamily('Calibri').setFontSize(11).setHorizontalAlignment('left').setVerticalAlignment('middle').setWrap(true);

    // Insérer une case à cocher dans la colonne "O" pour "Walid CHARMI" et dans la colonne "N" pour "Emna MEGDICH"
    if (sheetName === "Walid CHARMI") {
        var colonneO = compteSheet.getRange(lastRow, 15);
        colonneO.insertCheckboxes()
                .setHorizontalAlignment('center');

    } else if (sheetName === "Emna MEGDICH") {
        var colonneN = compteSheet.getRange(lastRow, 14);
        colonneN.insertCheckboxes()
                .setHorizontalAlignment('center');

// Construction de l'email pour les demandes dans la feuille "Emna MEGDICH"
var htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Nouvelle demande de compte institutionnel'</title>
  <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700&display=swap" rel="stylesheet">
  <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;700&display=swap" rel="stylesheet">
  <link href="https://fonts.googleapis.com/css2?family=Pacifico&display=swap" rel="stylesheet">
</head>
<body style="margin: 0; padding: 0; font-family: 'Montserrat', Arial, sans-serif; background-color: #F9FAFB; color: #1F2937;">
  <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 16px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06); margin-top: 20px; margin-bottom: 20px;">
    <!-- En-tête -->
    <div style="background: linear-gradient(120deg, #0EA5E9 0%, #0284C7 60%, #0369A1 100%); padding: 40px 30px; text-align: center; color: white;">
      <h1 style="margin: 0; font-size: 24px; font-weight: 700; letter-spacing: -0.5px;">Demande de compte institutionnel</h1>
      <p style="margin: 12px 0 0 0; font-size: 16px; font-weight: 300; letter-spacing: 1px; text-transform: uppercase;">Nouvelle notification</p>
    </div>

    <!-- Contenu principal -->
    <div style="padding: 32px 24px;">
      <p style="font-size: 16px; line-height: 1.6; color: #1F2937;">Bonjour,</p>

      <p style="font-size: 16px; line-height: 1.6; color: #374151;">Une nouvelle demande de compte institutionnel a été soumise sur la plateforme "Gestion des demandes des parties prenantes de l'ENET'Com".</p>

      <p style="font-size: 16px; line-height: 1.6; color: #374151;">Pour la traiter, veuillez cliquer sur le lien suivant :
      <a href="https://docs.google.com/spreadsheets/d/1Y-50gArs_Y9uZEY4g73OQu-QLbGZGyyzK49_n-KrZqU/edit?gid=1316351680#gid=1316351680"
         style="color: #0284C7; text-decoration: none; font-weight: 600;">Accéder au document</a></p>

      <div style="margin-top: 30px;">
        <p style="font-size: 16px; color: #374151; margin-bottom: 16px;">Cordialement,</p>
        <p>&nbsp;</p>
        <p>&nbsp;</p>
      </div>
    </div>

    <!-- Pied de page -->
    <div style="background-color: #4761a1; padding: 16px; text-align: center; font-size: 12px; color: #F3F4F6;">
      <p style="margin: 0 0 10px 0;">Walid CHARMI<br>Service informatique - ENET'Com</p>
    </div>
  </div>

  <!-- Message automatique -->
  <div style="text-align: center; font-size: 11px; font-style: italic; color: #000000; margin-top: 20px;">
    <p style="margin: 0;">Cet e-mail est généré automatiquement, merci de ne pas y répondre.</p>
  </div>
</body>
</html>
  `;

// Envoie le courriel
MailApp.sendEmail({
  to: "emna.megdich@enetcom.usf.tn",
  subject: "Nouvelle demande de compte institutionnel",
  name: "Walid CHARMI",
  htmlBody: htmlBody
});
}
}  
}

function sendMaterialRequestNotification(_formObject) {
  var recipients = [
    'karim.jaber@enetcom.usf.tn',
    'hassen.baccar@enetcom.usf.tn',
    'mourad.benjeddou@enetcom.usf.tn'
  ];

  var subject = 'Nouvelle demande de matériels';

  const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Nouvelle demande de matériels</title>
  <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700&display=swap" rel="stylesheet">
  <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;700&display=swap" rel="stylesheet">
  <link href="https://fonts.googleapis.com/css2?family=Pacifico&display=swap" rel="stylesheet">
</head>
<body style="margin: 0; padding: 0; font-family: 'Montserrat', Arial, sans-serif; background-color: #F9FAFB; color: #1F2937;">
  <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 16px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06); margin-top: 20px; margin-bottom: 20px;">
    <!-- En-tête -->
    <div style="background: linear-gradient(120deg, #0EA5E9 0%, #0284C7 60%, #0369A1 100%); padding: 40px 30px; text-align: center; color: white;">
      <h1 style="margin: 0; font-size: 24px; font-weight: 700; letter-spacing: -0.5px;">Demande de matériels</h1>
      <p style="margin: 12px 0 0 0; font-size: 16px; font-weight: 300; letter-spacing: 1px; text-transform: uppercase;">Nouvelle notification</p>
    </div>

    <!-- Contenu principal -->
    <div style="padding: 32px 24px;">
      <p style="font-size: 16px; line-height: 1.6; color: #1F2937;">Bonjour,</p>

      <p style="font-size: 16px; line-height: 1.6; color: #374151;">Une nouvelle demande de matériels a été soumise sur la plateforme "Gestion des demandes des parties prenantes de l'ENET'Com".</p>

      <p style="font-size: 16px; line-height: 1.6; color: #374151;">Pour la traiter, veuillez cliquer sur le lien suivant :
      <a href="https://docs.google.com/spreadsheets/d/1l2J5Fn4RLKdWkJZOrc16CRzPrP1TvxhFG12pFBULrv4/edit?gid=0#gid=0"
         style="color: #0284C7; text-decoration: none; font-weight: 600;">Accéder au document</a></p>

      <div style="margin-top: 30px;">
        <p style="font-size: 16px; color: #374151; margin-bottom: 16px;">Cordialement,</p>
        <p>&nbsp;</p>
        <p>&nbsp;</p>
      </div>
    </div>

    <!-- Pied de page -->
    <div style="background-color: #4761a1; padding: 16px; text-align: center; font-size: 12px; color: #F3F4F6;">
      <p style="margin: 0;">Walid CHARMI<br>Service informatique - ENET'Com</p>
    </div>
  </div>

  <!-- Message automatique -->
  <div style="text-align: center; font-size: 11px; font-style: italic; color: #000000; margin-top: 20px;">
    <p style="margin: 0;">Cet e-mail est généré automatiquement, merci de ne pas y répondre.</p>
  </div>
</body>
</html>
`;

  MailApp.sendEmail({
    to: recipients.join(','),
    subject: subject,
    htmlBody: htmlBody,
    name: 'Walid CHARMI'
  });
}

function sendStudentActivityRequestNotification(_formObject) {
  var recipients = [
    'ali.regaieg@enetcom.usf.tn'
  ];

  var subject = 'Nouvelle demande d’activité estudiantine';

  const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Nouvelle demande d’activité estudiantine</title>
  <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700&display=swap" rel="stylesheet">
  <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;700&display=swap" rel="stylesheet">
  <link href="https://fonts.googleapis.com/css2?family=Pacifico&display=swap" rel="stylesheet">
</head>
<body style="margin: 0; padding: 0; font-family: 'Montserrat', Arial, sans-serif; background-color: #F9FAFB; color: #1F2937;">
  <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 16px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06); margin-top: 20px; margin-bottom: 20px;">
    <!-- En-tête -->
    <div style="background: linear-gradient(120deg, #0EA5E9 0%, #0284C7 60%, #0369A1 100%); padding: 40px 30px; text-align: center; color: white;">
      <h1 style="margin: 0; font-size: 24px; font-weight: 700; letter-spacing: -0.5px;">Demande d'activité estudiantine</h1>
      <p style="margin: 12px 0 0 0; font-size: 16px; font-weight: 300; letter-spacing: 1px; text-transform: uppercase;">Nouvelle notification</p>
    </div>

    <!-- Contenu principal -->
    <div style="padding: 32px 24px;">
      <p style="font-size: 16px; line-height: 1.6; color: #1F2937;">Bonjour,</p>

      <p style="font-size: 16px; line-height: 1.6; color: #374151;">Une nouvelle demande d'activité estudiantine a été soumise sur la plateforme "Gestion des demandes des parties prenantes de l'ENET'Com".</p>

      <p style="font-size: 16px; line-height: 1.6; color: #374151;">Pour la traiter, veuillez cliquer sur le lien suivant :
      <a href="https://docs.google.com/spreadsheets/d/1f1Be7Lq-larL_Jva8KEDdsiNkJhxmMbj2GndITPySJk/edit?gid=0#gid=0"
         style="color: #0284C7; text-decoration: none; font-weight: 600;">Accéder au document</a></p>

      <div style="margin-top: 30px;">
        <p style="font-size: 16px; color: #374151; margin-bottom: 16px;">Cordialement,</p>
        <p>&nbsp;</p>
        <p>&nbsp;</p>
      </div>
    </div>

    <!-- Pied de page -->
    <div style="background-color: #4761a1; padding: 16px; text-align: center; font-size: 12px; color: #F3F4F6;">
      <p style="margin: 0 0 10px 0;">Walid CHARMI<br>Service informatique - ENET'Com</p>
    </div>
  </div>

  <!-- Message automatique -->
  <div style="text-align: center; font-size: 11px; font-style: italic; color: #000000; margin-top: 20px;">
    <p style="margin: 0;">Cet e-mail est généré automatiquement, merci de ne pas y répondre.</p>
  </div>
</body>
</html>
  `;

  MailApp.sendEmail({
    to: recipients.join(','),
    subject: subject,
    htmlBody: htmlBody,
    name: 'Walid CHARMI'
  });
}

function sendInter3t24NpUrJMNunMMASmhAM953bFGeLXzN7(_formObject) {
  var recipients = [
    'walid.charmi@enetcom.usf.tn',
    'mourad.benjeddou@enetcom.usf.tn',
    'hassen.baccar@enetcom.usf.tn',
    'feten.oussaifi@enetcom.usf.tn',
    'emna.megdich@enetcom.usf.tn'
  ];
  var subject = 'Nouvelle demande d\'intervention';

  var htmlBody = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Nouvelle demande d'intervention</title>
      <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    </head>
    <body style="margin: 0; padding: 0; font-family: 'Montserrat', Arial, sans-serif; background-color: #F9FAFB; color: #1F2937;">
      <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 16px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06); margin-top: 20px; margin-bottom: 20px;">
        <!-- En-tête -->
        <div style="background: linear-gradient(120deg, #0EA5E9 0%, #0284C7 60%, #0369A1 100%); padding: 40px 30px; text-align: center; color: white;">
          <h1 style="margin: 0; font-size: 24px; font-weight: 700; letter-spacing: -0.5px;">Demande d'intervention</h1>
          <p style="margin: 12px 0 0 0; font-size: 16px; font-weight: 300; letter-spacing: 1px; text-transform: uppercase;">Nouvelle notification</p>
        </div>

        <!-- Contenu principal -->
        <div style="padding: 32px 24px;">
          <p style="font-size: 16px; line-height: 1.6; color: #1F2937;">Bonjour,</p>
          <p style="font-size: 16px; line-height: 1.6; color: #374151;">Une nouvelle demande d'intervention a été soumise sur la plateforme "Gestion des demandes des parties prenantes de l'ENET'Com".</p>
          <p style="font-size: 16px; line-height: 1.6; color: #374151;">Pour la traiter, veuillez cliquer sur le lien suivant :
          <a href="https://docs.google.com/spreadsheets/d/1Bt_GLCPjMmQgKjamFu83joLOjKr1Vo9TH9-kpRTCUto/edit?gid=0#gid=0"
             style="color: #0284C7; text-decoration: none; font-weight: 600;">Accéder au document</a></p>

          <div style="margin-top: 30px;">
            <p style="font-size: 16px; color: #374151; margin-bottom: 16px;">Cordialement,</p>
            <p>&nbsp;</p>
            <p>&nbsp;</p>
          </div>
        </div>

        <!-- Pied de page -->
        <div style="background-color: #4761a1; padding: 16px; text-align: center; font-size: 12px; color: #F3F4F6;">
          <p style="margin: 0;">Walid CHARMI<br>Service informatique - ENET'Com</p>
        </div>
      </div>

      <!-- Message automatique -->
      <div style="text-align: center; font-size: 11px; font-style: italic; color: #000000; margin-top: 20px;">
        <p style="margin: 0;">Cet e-mail est généré automatiquement, merci de ne pas y répondre.</p>
      </div>
    </body>
    </html>
  `;

  MailApp.sendEmail({
    to: recipients.join(','),
    subject: subject,
    htmlBody: htmlBody,
    name: 'Walid CHARMI'
  });
}

function sendMaterielsNotificationToWalid(formObject) {
  const recipient = 'haytham.ghariani@enetcom.usf.tn';
  const subject = 'Nouvelle demande de matériel';

  const htmlBody = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <title>Nouvelle demande de matériels</title>
      <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; }
        .container { max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #ddd; border-radius: 5px; }
        .header { background-color: #4285f4; color: white; padding: 10px; text-align: center; border-radius: 5px 5px 0 0; }
        .content { padding: 20px; }
        .footer { text-align: center; font-size: 12px; color: #777; margin-top: 20px; }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h2>Nouvelle demande de matériels (Matériels)</h2>
        </div>
        <div class="content">
          <p>Bonjour,</p>
          <p>Une nouvelle demande de matériels a été soumise dans la catégorie <strong>Matériels</strong>.</p>
          <p><strong>Nom et prénom :</strong> ${formObject.nom} ${formObject.prenom}</p>
          <p><strong>Fonction :</strong> ${formObject.fonction}</p>
          <p><strong>Département :</strong> ${formObject.departement}</p>
          <p><strong>Grade :</strong> ${formObject.grade}</p>
          <p><strong>Justification :</strong> ${formObject.justification === 'Autre' ? formObject.precisezJustification : formObject.justification}</p>
          <p><strong>Matériels demandés :</strong></p>
          <ul>
            ${formObject.materiels.map(materiel => `<li>${materiel.quantite} x ${materiel.designation} (${materiel.prix} DT)</li>`).join('')}
          </ul>
          <p>Pour traiter cette demande, veuillez consulter le document : <a href="https://docs.google.com/spreadsheets/d/15H3ZAVjT6G2IBAsvKHuEn6UmdK_WqIR-G04O-yiefqw/edit#gid=123456789">Accéder au document</a>.</p>
          <p>Cordialement,</p>
        </div>
        <div class="footer">
          <p>© 2025 ENET'Com – Développé par Walid CHARMI</p>
        </div>
      </div>
    </body>
    </html>
  `;

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    htmlBody: htmlBody,
    name: 'Walid CHARMI'
  });
}
