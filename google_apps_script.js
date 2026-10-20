/**
 * Google Apps Script pro stezku "30 let pobočky Fantazie" (SVČ Helceletka)
 * 
 * NÁVOD K NASAZENÍ:
 * 1. Založte novou Google Tabulku (např. na Disku Google) a pojmenujte ji: "Fantazie 30 - Registrace týmů".
 * 2. V horním menu klikněte na: Rozšíření (Extensions) -> Apps Script.
 * 3. Smažte veškerý kód, který tam je, a vložte celý tento soubor.
 * 4. Nahoře nastavte email kolegyně (proměnná NOTIFY_EMAIL).
 * 5. Klikněte na ikonu diskety (Uložit).
 * 6. Vpravo nahoře klikněte na modré tlačítko: "Nasadit" (Deploy) -> "Nové nasazení" (New deployment).
 * 7. Zvolte typ: "Webová aplikace" (Web app).
 * 8. Vyplňte:
 *    - Popis: Fantazie 30 API
 *    - Spustit jako: Já (váš účet)
 *    - Kdo má přístup: Kdokoli (Anyone) -- DŮLEŽITÉ pro funkčnost z mobilů bez přihlašování!
 * 9. Klikněte na "Nasadit", schvalte oprávnění a zkopírujte vygenerovanou "URL webové aplikace".
 * 10. Tuto URL vložte do aplikace v index.html (konstanta GOOGLE_SHEET_API_URL).
 */

// ZDE ZADEJTE EMAIL KOLEGYNĚ PRO UPOZORNĚNÍ NA NOVÉ TÝMY:
const NOTIFY_EMAIL = "kolegyne@helceletka.cz"; // Upravte na skutečný email

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Týmy");
  if (!sheet) {
    sheet = ss.insertSheet("Týmy");
    // Vytvoříme záhlaví
    const headers = [
      "ID týmu",
      "Název týmu",
      "Tajné heslo",
      "Dětí",
      "Dospělých",
      "Vyřešeno bodů",
      "Složeno dílků",
      "Výherní kód",
      "Stav",
      "Datum registrace",
      "Poslední aktivita",
      "Data (JSON)"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9ead3");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doPost(e) {
  try {
    const rawData = e.postData ? e.postData.contents : "";
    const data = JSON.parse(rawData);
    const action = data.action;
    const sheet = getOrCreateSheet();
    const rows = sheet.getDataRange().getValues();

    if (action === "register") {
      const team = data.team;
      const normName = (team.teamName || "").trim().toLowerCase();

      // Kontrola, zda tým se stejným názvem již neexistuje
      for (let i = 1; i < rows.length; i++) {
        const existingName = String(rows[i][1] || "").trim().toLowerCase();
        if (existingName === normName) {
          return createJsonResponse({ success: false, error: "Tým s tímto názvem již v databázi existuje. Přihlaste se prosím heslem." });
        }
      }

      const teamId = team.id || ("team_" + new Date().getTime());
      const now = new Date().toLocaleString("cs-CZ");
      const visitedCount = Object.keys(team.visited || {}).filter(k => team.visited[k].correct).length;
      const placedCount = (team.placedPieces || []).length;
      const isWon = visitedCount === 12;

      sheet.appendRow([
        teamId,
        team.teamName,
        team.password,
        team.kidsCount,
        team.adultsCount,
        visitedCount,
        placedCount,
        team.uniqueCode || "",
        isWon ? "VÝHRA 🏆" : "Na trase 🏃",
        now,
        now,
        JSON.stringify(team)
      ]);

      // Odeslání emailové notifikace kolegyni
      if (NOTIFY_EMAIL && NOTIFY_EMAIL.includes("@")) {
        try {
          const subject = `🎉 Nový tým na stezce Fantazie: ${team.teamName}`;
          const body = 
            `Ahoj,\n\n` +
            `na výroční stezce "30 let pobočky Fantazie" se právě zaregistroval nový tým!\n\n` +
            `• Název týmu: ${team.teamName}\n` +
            `• Počet dětí: ${team.kidsCount}\n` +
            `• Počet dospělých: ${team.adultsCount}\n` +
            `• Tajné heslo týmu: ${team.password}\n` +
            `• Čas registrace: ${now}\n\n` +
            `Průběžný stav a výsledky můžete sledovat přímo v připojené Google Tabulce.\n\n` +
            `SVČ Helceletka – Pobočka Fantazie`;

          MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
        } catch (mailErr) {
          console.error("Chyba při odesílání emailu:", mailErr);
        }
      }

      return createJsonResponse({ success: true, team: team });

    } else if (action === "syncProgress") {
      const team = data.team;
      const normName = (team.teamName || "").trim().toLowerCase();
      let rowIndex = -1;

      for (let i = 1; i < rows.length; i++) {
        const idMatch = String(rows[i][0] || "") === team.id;
        const nameMatch = String(rows[i][1] || "").trim().toLowerCase() === normName;
        if (idMatch || nameMatch) {
          rowIndex = i + 1; // 1-based index
          break;
        }
      }

      const now = new Date().toLocaleString("cs-CZ");
      const visitedCount = Object.keys(team.visited || {}).filter(k => team.visited[k].correct).length;
      const placedCount = (team.placedPieces || []).length;
      const isWon = visitedCount === 12;

      if (rowIndex > 0) {
        // Aktualizujeme řádek
        sheet.getRange(rowIndex, 4).setValue(team.kidsCount);
        sheet.getRange(rowIndex, 5).setValue(team.adultsCount);
        sheet.getRange(rowIndex, 6).setValue(visitedCount);
        sheet.getRange(rowIndex, 7).setValue(placedCount);
        sheet.getRange(rowIndex, 8).setValue(team.uniqueCode || "");
        sheet.getRange(rowIndex, 9).setValue(isWon ? "VÝHRA 🏆" : "Na trase 🏃");
        sheet.getRange(rowIndex, 11).setValue(now);
        sheet.getRange(rowIndex, 12).setValue(JSON.stringify(team));
      } else {
        // Pokud tým v tabulce ještě nebyl, přidáme ho
        sheet.appendRow([
          team.id,
          team.teamName,
          team.password,
          team.kidsCount,
          team.adultsCount,
          visitedCount,
          placedCount,
          team.uniqueCode || "",
          isWon ? "VÝHRA 🏆" : "Na trase 🏃",
          now,
          now,
          JSON.stringify(team)
        ]);
      }

      return createJsonResponse({ success: true });

    } else if (action === "login") {
      const name = (data.teamName || "").trim().toLowerCase();
      const password = (data.password || "").trim();

      for (let i = 1; i < rows.length; i++) {
        const rowName = String(rows[i][1] || "").trim().toLowerCase();
        if (rowName === name) {
          const rowPass = String(rows[i][2] || "").trim();
          if (rowPass === password) {
            let fullTeam = null;
            try {
              fullTeam = JSON.parse(rows[i][11]);
            } catch (err) {}

            if (!fullTeam) {
              fullTeam = {
                id: String(rows[i][0]),
                teamName: String(rows[i][1]),
                password: rowPass,
                kidsCount: Number(rows[i][3]) || 1,
                adultsCount: Number(rows[i][4]) || 1,
                uniqueCode: String(rows[i][7] || ""),
                visited: {},
                placedPieces: []
              };
            }
            return createJsonResponse({ success: true, team: fullTeam });
          } else {
            return createJsonResponse({ success: false, error: "Zadané tajné heslo nesouhlasí." });
          }
        }
      }

      return createJsonResponse({ success: false, error: "Tým s tímto názvem nebyl v centrální databázi nalezen." });
    }

    return createJsonResponse({ success: false, error: "Neznámá akce." });

  } catch (err) {
    return createJsonResponse({ success: false, error: err.toString() });
  }
}

function doGet(e) {
  try {
    const action = e.parameter ? e.parameter.action : "";
    const sheet = getOrCreateSheet();
    const rows = sheet.getDataRange().getValues();

    if (action === "getAllTeams") {
      const teams = [];
      for (let i = 1; i < rows.length; i++) {
        teams.push({
          id: rows[i][0],
          teamName: rows[i][1],
          kidsCount: rows[i][3],
          adultsCount: rows[i][4],
          score: rows[i][5],
          placed: rows[i][6],
          code: rows[i][7],
          status: rows[i][8],
          updatedAt: rows[i][10]
        });
      }
      return createJsonResponse({ success: true, teams: teams });
    }

    return createJsonResponse({ success: true, message: "Fantazie 30 API běží v pořádku." });
  } catch (err) {
    return createJsonResponse({ success: false, error: err.toString() });
  }
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
