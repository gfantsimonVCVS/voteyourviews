// Paste into Apps Script as a new file named "renamePhotos"
// Select runHaysPhotoCleanup in the dropdown, then Run.
// Upload your background-removed photos to the same Drive folder first.
// If you have duplicates (same filename), this script keeps the newest version.

function runHaysPhotoCleanup() {
  var folderId = '1EBlVluSyl5weEG0X-9WIAqZ5Kvl7X8cW';
  var folder = DriveApp.getFolderById(folderId);

  // Map: base filename (no extension) → new name, or null = trash (lost primary)
  var fileMap = {
    // ── WINNERS ──────────────────────────────────────────────────────────────
    'james-talarico':          'Hays_Dem_USSen_James_Talarico',
    'kristin-hook':            'Hays_Dem_USHouse21_Kristin_Hook',
    'tanya-lloyd':             'Hays_Dem_USHouse27_Tanya_Lloyd',
    'Gina_Hinojosa_New':       'Hays_Dem_Gov_Gina_Hinojosa',
    'vikki-goodwin':           'Hays_Dem_LtGov_Vikki_Goodwin',
    'nathan-johnson':          'Hays_Dem_AG_Nathan_Johnson',
    'sarah-eckhardt':          'Hays_Dem_Comptroller_Sarah_Eckhardt',
    'benjamin-flores':         'Hays_Dem_LandComm_Benjamin_Flores',
    'clayton-tucker':          'Hays_Dem_AgComm_Clayton_Tucker',
    'jon-rosenthal':           'Hays_Dem_RRComm_Jon_Rosenthal',
    'allison-bush':            'Hays_Dem_SBOE5_Allison_Bush',
    'erin-zwiener':            'Hays_Dem_TXHouse45_Erin_Zwiener',
    'merrie-fox':              'Hays_Dem_TXHouse73_Merrie_Fox',
    'michelle-gutierrez-cohen':'Hays_Dem_CountyJudge_Michelle_GutierrezCohen',
    'alfonso-o-salazar':       'Hays_Dem_CrimDA_Alfonso_Salazar',
    'amanda-k-calvert':        'Hays_Dem_DistClerk_Amanda_Calvert',
    'daphne-sanchez-tenorio':  'Hays_Dem_CountyTreasurer_Daphne_Tenorio',
    'johnny-flores':           'Hays_Dem_CommPct2_Johnny_Flores',
    'angie-unger':             'Hays_Dem_CommPct4_Angie_Unger',
    'maggie-ellis':            'Hays_Dem_TXSCChiefJustice_Maggie_Ellis',
    'chari-kelly':             'Hays_Dem_TXSCPl2_Chari_Kelly',
    'kristen-hawkins':         'Hays_Dem_TXSCPl7_Kristen_Hawkins',
    'gisela-d-triana':         'Hays_Dem_TXSCPl8_Gisela_Triana',
    'okey-anyiam':             'Hays_Dem_CCAPl3_Okey_Anyiam',
    'audra-riley':             'Hays_Dem_CCAPl4_Audra_Riley',
    'holly-taylor':            'Hays_Dem_CCAPl9_Holly_Taylor',
    'jerry-zimmerer':          'Hays_Dem_15thCourtCJ_Jerry_Zimmerer',
    'tom-baker':               'Hays_Dem_15thCourtPl2_Tom_Baker',
    'marc-m-meyer':            'Hays_Dem_15thCourtPl3_Marc_Meyer',
    'darlene-byrne':           'Hays_Dem_3rdCourtCJ_Darlene_Byrne',
    'chris-johnson':           'Hays_Dem_CCLNo2_Chris_Johnson',

    // ── LOSERS / NOT IN HAYS RACES — TRASH ───────────────────────────────────
    'ahmad-r-hassan':          null,
    'elizabeth-trevino-amaya': null,
    'anthony-tony-box':        null,
    'willie-tenorio-jr':       null,
    'joel-w-martin':           null,
    'neto-longoria':           null,
    'beth-smith':              null,
    'victor-sampson':          null,
    'laurie-brown':            null,
    'thomas-just':             null,
    'gary-taylor':             null,
    'jose-navarro-balbuena':   null,
    'stephanie-limon-bazan':   null,
    'gordon-goodman':          null,
    'jimmy-alan-hall':         null,
    'kevin-jackson':           null,
    'maggie-hernandez-moreno': null,
    'courtney-head':           null,
    'cynthia-a-millonzi':      null,
    'bobby-cole':              null,
    'jasmine-crockett':        null,
    'landon-bryan-campbell':   null,
    'cassie-benoist-templeton':null,
    'cory-l-carlyle':          null,
    'cortney-jones':           null,
    'sandra-bryant':           null,
    'bill-henry':              null,
    'jose-loya':               null,
    'marcos-isaias-velez':     null,
    'zach-vance':              null,
    'joe-jaworski':            null,
    'joe-pool':                null,
    'nicholas-nico-costilla':  null,
    'judith-zaffirini':        null,
    'michael-lange':           null,
    'w-david-friesenhahn':     null,
    'andrew-white':            null,
    'regina-vanburg':          null,
    'abigail-gray':            null,
    'deedee-rodgers':          null,
    'ruben-becerra':           null,
    'bryan-escobar':           null,
    'savant-moore':            null,
    'chris-bell':              null
  };

  // ── PASS 1: Group files by base name, trash older duplicates ─────────────
  var groups = {};
  var allFiles = folder.getFiles();

  while (allFiles.hasNext()) {
    var f = allFiles.next();
    var title = f.getName();
    if (title === '.DS_Store') { f.setTrashed(true); continue; }
    var dot = title.lastIndexOf('.');
    var base = dot > -1 ? title.substring(0, dot) : title;
    if (!groups[base]) groups[base] = [];
    groups[base].push(f);
  }

  // For any base name with multiple files, keep the newest, trash the rest
  for (var base in groups) {
    var group = groups[base];
    if (group.length > 1) {
      group.sort(function(a, b) { return b.getDateCreated() - a.getDateCreated(); });
      for (var i = 1; i < group.length; i++) {
        group[i].setTrashed(true);
        Logger.log('DEDUPED (older removed): ' + group[i].getName());
      }
      groups[base] = [group[0]]; // keep only newest
    }
  }

  // ── PASS 2: Rename winners, trash losers ─────────────────────────────────
  var renamed = 0, trashed = 0, skipped = [];

  for (var base in groups) {
    var file = groups[base][0];
    var fullName = file.getName();
    var dot = fullName.lastIndexOf('.');
    var ext = dot > -1 ? fullName.substring(dot) : '';

    if (base in fileMap) {
      var newBase = fileMap[base];
      if (newBase === null) {
        file.setTrashed(true);
        Logger.log('TRASHED: ' + fullName);
        trashed++;
      } else {
        file.setName(newBase + ext);
        Logger.log('RENAMED: ' + fullName + '  →  ' + newBase + ext);
        renamed++;
      }
    } else {
      skipped.push(fullName);
    }
  }

  Logger.log('────────────────────────────────────────');
  Logger.log('DONE.  Renamed: ' + renamed + '  |  Trashed: ' + trashed);
  if (skipped.length > 0) {
    Logger.log('UNRECOGNIZED (not touched): ' + skipped.join(', '));
  }
}
