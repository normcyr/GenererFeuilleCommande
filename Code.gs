function keys(obj)
{
  var keys = [];

  for(var key in obj)
  {
    if(obj.hasOwnProperty(key))
    {
      keys.push(key);
    }
  }

  return keys;
}

function confirmOverwrite() {
  var confirm = Browser.msgBox('Feuille existante','La feuille Babac existe déjà. Écraser?', Browser.Buttons.OK_CANCEL);
  return confirm == 'ok';
}

function fixProductNumbers() {
  num_babac = (String(num_babac).replace(/(\d\d)(\d\d\d)/, '$1-$2'));
    return num_babac;
}

function ecraserAncienneCommande(sheetBabac) {
  if (sheetBabac) {
    if (!confirmOverwrite()) {
      return;
    }

    Logger.log('Clearing sheet Babac.');
    sheetBabac.clear();
  } else {
    Logger.log('Creating sheet Babac.');
    sheetBabac = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Babac');
  }
}

function listePiecesAtelier(sheetAtelier) {
  /* Déterminer le contenu de chaque colonne */
  NUM_BICIKLO_COL = 0;
  NUM_BABAC_COL = 1;
  CATEGORIE_COL = 2;
  NOM_COL = 3;
  REFERENCE_COL = 4;
  CARACTERISTIQUE_COL = 5;
  PRIX_BABAC_COL = 6;
  PAQUETS_DE_COL = 7;
  QUANTITE_COL = 8;

  range = sheetAtelier.getSheetValues(2, 1, -1, -1);

  for (row in range) {
    row = range[row]
    num_babac = row[NUM_BABAC_COL];

    nb_paquets = row[QUANTITE_COL];
    nb_par_paquet = row[PAQUETS_DE_COL];
    nom = row[NOM_COL] + " " + row[REFERENCE_COL] + " " + row[CARACTERISTIQUE_COL];

    if (!num_babac)
      continue;

    if (!(num_babac in commande)) {
      fixProductNumbers(num_babac);
      commande[num_babac] = {};
      commande[num_babac]['quantite'] = 0;
      commande[num_babac]['nb_par_paquet'] = nb_par_paquet;
      commande[num_babac]['nom'] = nom;
    }

    commande[num_babac]['quantite'] += nb_paquets;
  }
}

/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */

function generateBabac() {
  Logger.log("Hello world");

  var sheetBabac = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Babac');
  ecraserAncienneCommande(sheetBabac);

  var sheetAtelier = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Atelier");
  var sheetPerso = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Perso");

  commande = {};

  listePiecesAtelier(sheetAtelier);
/*  listePiecesPerso(sheetPerso);*/

  /* Ajouter pièces perso */
  range = sheetPerso.getSheetValues(2, 1, -1, -1);

  PERSO_NUM_BABAC_COL = 1;
  PERSO_NOM_COL = 2;
  PERSO_PAQUETS_DE_COL = 4;
  PERSO_QUANTITE_COL = 5;

  for (row in range) {
    row = range[row]
    num_babac = row[PERSO_NUM_BABAC_COL];
    nb_paquets = row[PERSO_QUANTITE_COL];
    nb_par_paquet = row[PERSO_PAQUETS_DE_COL];
    nom = row[PERSO_NOM_COL];

    if (!num_babac)
      continue;

    if (!(num_babac in commande)) {
      fixProductNumbers(num_babac);
      commande[num_babac] = {};
      commande[num_babac]['quantite'] = 0;
      commande[num_babac]['nb_par_paquet'] = nb_par_paquet;
      commande[num_babac]['nom'] = nom;
    }

    commande[num_babac]['quantite'] += nb_paquets;
  }


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

  Logger.log('Done');

};

function myFunction() {

 var numero = Browser.inputBox("Numéro babac");
 var response = UrlFetchApp.fetch("http://nova.polymtl.ca/~simark/cgi-bin/babac.py?piece=" + numero);
 Logger.log(response.getContentText());

}


/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
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
