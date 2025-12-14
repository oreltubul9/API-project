// --- DOM ---
const urlInput = document.getElementById("urlInput");
const scanBtn = document.getElementById("scanBtn");
const statusEl = document.getElementById("status");
const resultsSummaryEl = document.getElementById("resultsSummary");
const resultsWrapperEl = document.getElementById("resultsTableWrapper");

// דוגמה (לא חובה)
const demoUrl = "http://base0010.sites.airnet/sites/Yaba528/Minhala/DocLib29/Forms/AllItems.aspx";
const loadFromExistingBtn = document.getElementById("loadFromExisting");
if (loadFromExistingBtn) {
  loadFromExistingBtn.addEventListener("click", () => {
    urlInput.value = demoUrl;
    statusEl.textContent = "נטען URL לדוגמה.";
    statusEl.className = "status info";
  });
}

scanBtn.addEventListener("click", async () => {
  const raw = urlInput.value.trim();
  if (!raw) {
    statusEl.textContent = "יש להזין כתובת URL.";
    statusEl.className = "status error";
    return;
  }
  await checkSharePointLibrary(raw);
});

function setStatus(msg, type = "info") {
  statusEl.textContent = msg;
  statusEl.className = `status ${type}`;
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/**
 * מקבל URL של ספרייה/דף בתוך ספרייה, ומחזיר:
 * siteUrl = עד רמת האתר (למשל .../Minhala)
 * libraryServerRelativeUrl = הנתיב של הספרייה (למשל /sites/.../Minhala/DocLib29)
 */
function parseSharePointUrls(inputUrl) {
  const u = new URL(inputUrl);

  const parts = u.pathname.split("/").filter(Boolean);
  // דוגמה: ["sites","Yaba528","Minhala","DocLib29","Forms","AllItems.aspx"]

  // נניח שבסביבה שלך יש תמיד /sites/<SiteCollection>/<Subsite>/...
  // האתר הוא 3 החלקים הראשונים:
  const siteParts = parts.slice(0, 3); // ["sites","Yaba528","Minhala"]
  const siteUrl = `${u.origin}/${siteParts.join("/")}`;

  // הספרייה היא החלק הרביעי (DocLibXX)
  const libName = parts[3];
  if (!libName) throw new Error("לא הצלחתי לזהות שם ספרייה מה-URL.");

  const libraryServerRelativeUrl = `/${siteParts.join("/")}/${libName}`;

  return { siteUrl, libraryServerRelativeUrl };
}

async function spFetchJson(url) {
  const res = await fetch(url, {
    method: "GET",
    headers: { "Accept": "application/json;odata=verbose" },
    credentials: "include"
  });

  if (!res.ok) {
    // לפעמים SharePoint מחזיר HTML/טקסט – נקרא כדי לראות מה באמת קורה
    const text = await res.text();
    throw new Error(`HTTP ${res.status} from ${url}\n${text.slice(0, 200)}`);
  }

  return res.json();
}

async function checkSharePointLibrary(inputUrl) {
  try {
    setStatus("בודק את ה-URL ומחשב Site/Library…", "info");
    resultsSummaryEl.textContent = "";
    resultsWrapperEl.innerHTML = "";

    const { siteUrl, libraryServerRelativeUrl } = parseSharePointUrls(inputUrl);

    console.log("Site URL =", siteUrl);
    console.log("Library SR URL =", libraryServerRelativeUrl);

    setStatus(`Site URL: ${siteUrl} | Library: ${libraryServerRelativeUrl}`, "info");

    // 1) מביאים Metadata של הרשימה לפי URL של ספרייה
    // זה חשוב: לא מחפשים לפי Title כי DocLib29 הוא "שם פנימי" ולא תמיד Title
    const listMeta = await spFetchJson(
      `${siteUrl}/_api/web/GetList(@v)?@v='${encodeURIComponent(libraryServerRelativeUrl)}'`
    );

    const listId = listMeta?.d?.Id;
    const listTitle = listMeta?.d?.Title;

    if (!listId) {
      throw new Error("לא הצלחתי לקבל List Id מהספרייה.");
    }

    setStatus(`נמצאה ספרייה: ${listTitle} (Id: ${listId}) — מושך פריטים…`, "info");

    // 2) מביאים פריטים (מסמכים) מהספרייה
    const itemsJson = await spFetchJson(
      `${siteUrl}/_api/web/lists(guid'${listId}')/items?$top=5000&$select=Id,Title,FileLeafRef,FileRef,Modified,Editor/Title&$expand=Editor`
    );

    const items = itemsJson?.d?.results || [];
    renderResults(items, listTitle);

    setStatus(`הצלחה! נמצאו ${items.length} פריטים בספרייה "${listTitle}".`, "info");
  } catch (err) {
    console.error(err);
    setStatus("שגיאה בביצוע סריקה. פתח Console לפרטים.", "error");
    resultsSummaryEl.textContent = "";
    resultsWrapperEl.innerHTML = `<pre style="white-space:pre-wrap">${escapeHtml(err.message || String(err))}</pre>`;
  }
}

function renderResults(items, libraryTitle) {
  resultsSummaryEl.textContent = `ספרייה: ${libraryTitle} | כמות פריטים: ${items.length}`;

  if (!items.length) {
    resultsWrapperEl.innerHTML = "<div>לא נמצאו פריטים.</div>";
    return;
  }

  const rows = items.map(it => {
    const name = it.FileLeafRef || it.Title || "";
    const path = it.FileRef || "";
    const modified = it.Modified ? new Date(it.Modified).toLocaleString() : "";
    const editor = it.Editor?.Title || "";
    return `
      <tr>
        <td>${escapeHtml(name)}</td>
        <td>${escapeHtml(modified)}</td>
        <td>${escapeHtml(editor)}</td>
        <td>${escapeHtml(path)}</td>
      </tr>
    `;
  }).join("");

  resultsWrapperEl.innerHTML = `
    <table class="resultsTable">
      <thead>
        <tr>
          <th>שם</th>
          <th>עודכן</th>
          <th>עודכן ע"י</th>
          <th>נתיב</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}
