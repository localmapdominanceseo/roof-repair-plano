function publishPendingArticles() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Articles");
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];

  const col = {};
  headers.forEach((h, i) => col[h] = i);

  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("GITHUB_TOKEN");
  const owner = props.getProperty("GITHUB_OWNER");
  const repo = props.getProperty("GITHUB_REPO");
  const branch = props.getProperty("GITHUB_BRANCH") || "main";
  const siteBase = props.getProperty("SITE_BASE_URL");

  if (!token || !owner || !repo || !siteBase) {
    throw new Error("Missing Script Properties: GITHUB_TOKEN, GITHUB_OWNER, GITHUB_REPO, SITE_BASE_URL");
  }

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    const status = String(row[col["Status"]] || "").trim();

    if (status !== "Pending") continue;

    const title = String(row[col["Title"]] || "").trim();
    let slug = String(row[col["Slug"]] || "").trim();
    const content = String(row[col["Content"]] || "").trim();

    if (!title || !content) {
      sheet.getRange(r + 1, col["Notes"] + 1).setValue("Missing title or content");
      continue;
    }

    if (!slug) slug = makeSlug(title);

    const html = buildArticleHtml(title, slug, content, siteBase);
    const path = "articles/" + slug + "/index.html";

    commitFileToGitHub(owner, repo, branch, path, html, "Publish article: " + title, token);

    const liveUrl = siteBase.replace(/\/$/, "") + "/articles/" + slug + "/";

    sheet.getRange(r + 1, col["Slug"] + 1).setValue(slug);
    sheet.getRange(r + 1, col["Status"] + 1).setValue("Published");
    sheet.getRange(r + 1, col["URL"] + 1).setValue(liveUrl);
    sheet.getRange(r + 1, col["Date Published"] + 1).setValue(new Date());
    sheet.getRange(r + 1, col["Notes"] + 1).setValue("Committed to GitHub");
  }
}

function commitFileToGitHub(owner, repo, branch, path, content, message, token) {
  const apiUrl = "https://api.github.com/repos/" + owner + "/" + repo + "/contents/" + path;

  let sha = null;
  const getResp = UrlFetchApp.fetch(apiUrl + "?ref=" + branch, {
    method: "get",
    headers: {
      Authorization: "token " + token,
      Accept: "application/vnd.github+json"
    },
    muteHttpExceptions: true
  });

  if (getResp.getResponseCode() === 200) {
    sha = JSON.parse(getResp.getContentText()).sha;
  }

  const payload = {
    message: message,
    content: Utilities.base64Encode(content, Utilities.Charset.UTF_8),
    branch: branch
  };

  if (sha) payload.sha = sha;

  const putResp = UrlFetchApp.fetch(apiUrl, {
    method: "put",
    contentType: "application/json",
    headers: {
      Authorization: "token " + token,
      Accept: "application/vnd.github+json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (![200, 201].includes(putResp.getResponseCode())) {
    throw new Error("GitHub commit failed: " + putResp.getResponseCode() + " " + putResp.getContentText());
  }
}

function buildArticleHtml(title, slug, content, siteBase) {
  const safeTitle = escapeHtml(title);
  const body = convertSimpleContentToHtml(content);
  const canonical = siteBase.replace(/\/$/, "") + "/articles/" + slug + "/";

  return '<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>' + safeTitle + '</title>' +
    '<meta name="description" content="' + safeTitle + '">' +
    '<link rel="canonical" href="' + canonical + '">' +
    '<link rel="stylesheet" href="/assets/css/style.css">' +
    '</head><body><main class="page">' +
    '<p class="meta">DMV Guide</p><h1>' + safeTitle + '</h1>' +
    '<article class="card">' + body + '</article>' +
    '<a class="back" href="/">← Back to DMV index</a>' +
    '<div class="footer">Unofficial guide. Always verify information with the official state DMV.</div>' +
    '</main></body></html>';
}

function convertSimpleContentToHtml(text) {
  if (/<[a-z][\s\S]*>/i.test(text)) return text;

  return text
    .split(/\n\s*\n/)
    .map(p => {
      const line = p.trim();
      if (!line) return "";
      if (line.startsWith("# ")) return "<h2>" + escapeHtml(line.replace("# ", "")) + "</h2>";
      if (line.startsWith("## ")) return "<h2>" + escapeHtml(line.replace("## ", "")) + "</h2>";
      return "<p>" + escapeHtml(line).replace(/\n/g, "<br>") + "</p>";
    })
    .join("\n");
}

function makeSlug(title) {
  return title.toLowerCase()
    .replace(/[’']/g, "")
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .substring(0, 90);
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}