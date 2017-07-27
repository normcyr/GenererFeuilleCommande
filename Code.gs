function keys(obj) {
  var keys = [];
  for(var key in obj)
  {
    if(obj.hasOwnProperty(key))
    {
      keys.push(key);
    }
  }

  return keys;
};

function confirmOverwrite() {
  /* Vérifier si ok d'écraser la feuille Babac déjà construite */
  var confirm = Browser.msgBox('Feuille existante','La feuille Babac existe déjà. Écraser?', Browser.Buttons.OK_CANCEL);
  return confirm == 'ok';
};

function ecraserAncienneCommande(sheetBabac) {
  if (sheetBabac) {
    if (!confirmOverwrite()) {
      return;
    }

    Logger.log('Effaçage de la feuille Babac.');
    sheetBabac.clear();
  } else {
    Logger.log('Création de la feuille Babac.');
    sheetBabac = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Babac');
  }
};

function fixProductNumbers() {
  num_babac = (String(num_babac).replace(/(\d\d)(\d\d\d)/, '$1-$2'));
  Logger.log('No. de produit reformaté: ' + num_babac);
  return num_babac;
};

function entrerFeuilleBabac() {
  for (row in range) {
    row = range[row]
    num_babac = row[NUM_BABAC_COL];
    fixProductNumbers(num_babac);

    nb_paquets = row[QUANTITE_COL];
    nb_par_paquet = row[PAQUETS_DE_COL];
    nom = row[NOM_COL] + " " + row[REFERENCE_COL] + " " + row[CARACTERISTIQUE_COL];

    if (!num_babac)
      continue;

    if (!(num_babac in commande)) {
      commande[num_babac] = {};
      commande[num_babac]['quantite'] = 0;
      commande[num_babac]['nb_par_paquet'] = nb_par_paquet;
      commande[num_babac]['nom'] = nom;
    }

    commande[num_babac]['quantite'] += nb_paquets;
  }
};

function listePiecesAtelier(sheetAtelier) {
  /* Déterminer le contenu de chaque colonne pour feuille Atelier */
  NUM_BICIKLO_COL = 0;
  NUM_BABAC_COL = 1;
  CATEGORIE_COL = 2;
  NOM_COL = 3;
  REFERENCE_COL = 4;
  CARACTERISTIQUE_COL = 5;
  PRIX_BABAC_COL = 6;
  PAQUETS_DE_COL = 7;
  QUANTITE_COL = 8;

  /* extraire les données de la feuille */
  range = sheetAtelier.getSheetValues(2, 1, -1, -1);

  /* mettre dans feuille Babac */
  entrerFeuilleBabac(range);
};

function listePiecesPerso(sheetPerso) {
  /* Déterminer le contenu de chaque colonne pour feuille Perso */
  NUM_BABAC_COL = 1;
  NOM_COL = 2;
  PAQUETS_DE_COL = 4;
  QUANTITE_COL = 5;
  REFERENCE_COL = 9;
  CARACTERISTIQUE_COL = 10;

  /* extraire les données de la feuille */
  range = sheetPerso.getSheetValues(2, 1, -1, -1);

  /* entrer les données dans la feuille Babac */
  entrerFeuilleBabac(range);
};

function formaterFeuilleBabac(sheetBabac) {
  /* Formater la feuille Babac */
  row = sheetBabac.getRange(1, 1, 1, 3);
  row.setValues([["# Babac", "Nom", "Quantité"]]);
  row.setFontWeight("bold");

  curRow = 2;
  nums_babac = keys(commande).sort();
  for (num_babac in nums_babac) {
    num_babac = nums_babac[num_babac];

    item = commande[num_babac];
    nom = item['nom'];
    nb_par_paquet = item['nb_par_paquet'];
    quantite = item['quantite'];
    row = sheetBabac.getRange(curRow, 1, 1, 3);

    if (nb_par_paquet > 1) {
      paquets_str = quantite > 1 ? "paquets": "paquet";
      quantite_str = "" + quantite + " " + paquets_str + " de " + nb_par_paquet;
    } else {
      quantite_str = quantite;
    }
    row.setValues([[num_babac, nom, quantite_str]]);
    row.setHorizontalAlignments([['center','left','left']]);

    if (curRow % 2) {
      row.setBackground('#eee');
    }

    curRow++;
  }

  sheetBabac.autoResizeColumn(1);
  sheetBabac.autoResizeColumn(2);
  sheetBabac.autoResizeColumn(3);
};

function generateBabac() {
  Logger.log('Bonjour le monde');

  var sheetBabac = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Babac');
  var sheetAtelier = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Atelier");
  var sheetPerso = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Perso");

  ecraserAncienneCommande(sheetBabac);
  commande = {};

  listePiecesAtelier(sheetAtelier);
  listePiecesPerso(sheetPerso);
  formaterFeuilleBabac(sheetBabac);

  Logger.log('Done');
};

function myFunction() {
  /* Fonction à Simon - à mettre à jour */
  var numero = Browser.inputBox("Numéro babac");
  var response = UrlFetchApp.fetch("http://nova.polymtl.ca/~simark/cgi-bin/babac.py?piece=" + numero);
  Logger.log(response.getContentText());
};

function onOpen() {
  /* Ajouter un menu pour initier les scripts */
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [];
  entries.push({
    name : "Générer feuille Babac",
    functionName : "generateBabac"
  });
  entries.push({
    name : "Test",
    functionName : "myFunction"
  });
  Logger.log(entries);
  spreadsheet.addMenu("Scripts", entries);
};
