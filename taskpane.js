/* ================================================================
   NDL Vendor News — Outlook Add-in Task Pane
   Reads selected email → sends to Power Automate HTTP trigger
   ================================================================ */

// ── CONFIGURATION ─────────────────────────────────────────────────
// Replace this URL with your Power Automate "When an HTTP request is received" trigger URL
const FLOW_ENDPOINT = "https://default5e8309eec8d04bc7b5b85a37a4eb10.7b.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/83bfdcddda9744a1975b6bb497f37fab/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=6pc-G6UCRs2Z6Flg--0IqVsBdIf_P7DQc9XzPxGrEho";
// ──────────────────────────────────────────────────────────────────

let emailData = {
  subject: "",
  from: "",
  body: "",
  bodyText: "",
  attachments: [],
  messageId: "",
};

// ── Office ready ──────────────────────────────────────────────────
Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    loadEmailData();
    document.getElementById("vendorSelect").addEventListener("change", onVendorChange);
    document.getElementById("submitBtn").addEventListener("click", onSubmit);
  }
});

// ── Load email data from the selected message ─────────────────────
function loadEmailData() {
  var item = Office.context.mailbox.item;

  // Subject
  emailData.subject = item.subject || "(No subject)";
  document.getElementById("emailSubject").textContent = emailData.subject;
  document.getElementById("newsTitle").value = emailData.subject;

  // From
  if (item.from) {
    emailData.from = item.from.displayName + " <" + item.from.emailAddress + ">";
  }
  document.getElementById("emailFrom").textContent = emailData.from;

  // Message ID for attachment retrieval
  emailData.messageId = item.itemId;

  // Get HTML body (preserves formatting + images)
  item.body.getAsync(Office.CoercionType.Html, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      emailData.body = result.value;
      var plainText = htmlToCleanText(result.value);
      emailData.bodyText = plainText;
      var preview = plainText.substring(0, 200);
      if (plainText.length > 200) preview += "...";
      document.getElementById("emailBodyPreview").textContent = preview;

      // Resolve inline images (cid: references → base64 data URLs)
      resolveInlineImages(item, function (resolvedHtml) {
        emailData.body = resolvedHtml;
      });
    } else {
      item.body.getAsync(Office.CoercionType.Text, function (textResult) {
        if (textResult.status === Office.AsyncResultStatus.Succeeded) {
          emailData.bodyText = textResult.value;
          var preview = textResult.value.substring(0, 200);
          if (textResult.value.length > 200) preview += "...";
          document.getElementById("emailBodyPreview").textContent = preview;
        }
      });
    }
  });

  // Attachments
  loadAttachments(item);
}

// ── Load attachments ──────────────────────────────────────────────
function loadAttachments(item) {
  var attachments = item.attachments;
  var attachmentsContainer = document.getElementById("emailAttachments");
  var toggleRow = document.getElementById("attachToggleRow");

  if (!attachments || attachments.length === 0) {
    attachmentsContainer.style.display = "none";
    toggleRow.style.display = "none";
    return;
  }

  attachmentsContainer.style.display = "flex";
  toggleRow.style.display = "flex";
  attachmentsContainer.innerHTML = "";

  emailData.attachments = [];

  for (var i = 0; i < attachments.length; i++) {
    var att = attachments[i];
    if (att.isInline) continue;

    emailData.attachments.push({
      id: att.id,
      name: att.name,
      size: att.size,
      contentType: att.contentType,
    });

    var chip = document.createElement("div");
    chip.className = "email-att-chip";
    chip.innerHTML = getFileIcon(att.name) + " " + att.name +
      " <span style='color:#9ca3af;font-size:10px;'>(" + formatSize(att.size) + ")</span>";
    attachmentsContainer.appendChild(chip);
  }

  if (emailData.attachments.length === 0) {
    attachmentsContainer.style.display = "none";
    toggleRow.style.display = "none";
  }
}

// ── Vendor selection change ──────────────────────────────────────
function onVendorChange() {
  var vendor = document.getElementById("vendorSelect").value;
  var btn = document.getElementById("submitBtn");
  if (vendor) {
    btn.disabled = false;
    btn.textContent = "Submit Vendor News";
  } else {
    btn.disabled = true;
    btn.textContent = "Select a vendor to continue";
  }
}

// ── Submit: send to Power Automate ────────────────────────────────
function onSubmit() {
  var vendor = document.getElementById("vendorSelect").value;
  var category = document.getElementById("categorySelect").value;
  var title = document.getElementById("newsTitle").value.trim();
  var notes = document.getElementById("newsNotes").value.trim();
  var includeAtt = document.getElementById("includeAttachments").checked;

  if (!vendor || !title) return;

  setStatus("loading", "Submitting vendor news...");
  var btn = document.getElementById("submitBtn");
  btn.disabled = true;
  btn.textContent = "Submitting...";

  // If we need to include attachments, fetch their content first
  if (includeAtt && emailData.attachments.length > 0) {
    fetchAttachmentContents(function (attachmentContents) {
      sendToFlow(vendor, category, title, notes, attachmentContents);
    });
  } else {
    sendToFlow(vendor, category, title, notes, []);
  }
}

// ── Fetch attachment binary content via Office JS ─────────────────
function fetchAttachmentContents(callback) {
  var item = Office.context.mailbox.item;
  var results = [];
  var remaining = emailData.attachments.length;

  if (remaining === 0) {
    callback([]);
    return;
  }

  for (var i = 0; i < emailData.attachments.length; i++) {
    (function (att) {
      item.getAttachmentContentAsync(att.id, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          results.push({
            name: att.name,
            contentType: att.contentType,
            content: result.value.content,
            format: result.value.format,
          });
        }
        remaining--;
        if (remaining === 0) {
          callback(results);
        }
      });
    })(emailData.attachments[i]);
  }
}

// ── Send payload to Power Automate HTTP trigger ───────────────────
function sendToFlow(vendor, category, title, notes, attachments) {
  var payload = {
    vendorName: vendor,
    category: category,
    title: title,
    notes: notes,
    htmlBody: sanitizeOutlookHtml(emailData.body) || emailData.bodyText,
    plainBody: emailData.bodyText,
    emailFrom: emailData.from,
    emailSubject: emailData.subject,
    submittedBy: Office.context.mailbox.userProfile.displayName,
    submittedByEmail: Office.context.mailbox.userProfile.emailAddress,
    newsDate: new Date().toISOString().split("T")[0],
    attachments: attachments.map(function (a) {
      return {
        fileName: a.name,
        contentType: a.contentType,
        contentBytes: a.content,
      };
    }),
  };

  fetch(FLOW_ENDPOINT, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  })
    .then(function (response) {
      if (response.ok) {
        return response.json().catch(function () { return {}; });
      }
      throw new Error("Flow returned " + response.status);
    })
    .then(function (data) {
      setStatus("success", "Vendor news submitted! ✓");
      var btn = document.getElementById("submitBtn");
      btn.textContent = "Submitted ✓";

      setTimeout(function () {
        setStatus("", "");
        btn.disabled = false;
        btn.textContent = "Submit Another";
      }, 3000);
    })
    .catch(function (err) {
      setStatus("error", "Failed: " + err.message);
      var btn = document.getElementById("submitBtn");
      btn.disabled = false;
      btn.textContent = "Retry";
    });
}

// ── Sanitize Outlook HTML → clean simple HTML for SharePoint ─────
function sanitizeOutlookHtml(rawHtml) {
  if (!rawHtml) return "";
  try {
    var parser = new DOMParser();
    var doc = parser.parseFromString(rawHtml, "text/html");

    // Remove <style> blocks
    var styles = doc.querySelectorAll("style");
    for (var i = 0; i < styles.length; i++) styles[i].remove();

    // Remove <script> blocks
    var scripts = doc.querySelectorAll("script");
    for (var i = 0; i < scripts.length; i++) scripts[i].remove();

    // Remove <meta>, <link>, <title>
    var junk = doc.querySelectorAll("meta, link, title");
    for (var i = 0; i < junk.length; i++) junk[i].remove();

    // Remove comments
    var walker = doc.createTreeWalker(doc.body || doc, NodeFilter.SHOW_COMMENT, null, false);
    var comments = [];
    while (walker.nextNode()) comments.push(walker.currentNode);
    for (var i = 0; i < comments.length; i++) comments[i].remove();

    // Strip class and style attributes from all elements
    var allEls = (doc.body || doc).querySelectorAll("*");
    for (var i = 0; i < allEls.length; i++) {
      allEls[i].removeAttribute("class");
      allEls[i].removeAttribute("style");
      allEls[i].removeAttribute("lang");
      allEls[i].removeAttribute("dir");
      var attrs = allEls[i].attributes;
      var toRemove = [];
      for (var j = 0; j < attrs.length; j++) {
        if (attrs[j].name.indexOf("data-") === 0 || attrs[j].name.indexOf("o:") === 0) {
          toRemove.push(attrs[j].name);
        }
      }
      for (var j = 0; j < toRemove.length; j++) allEls[i].removeAttribute(toRemove[j]);
    }

    // Unwrap <font> and <span> tags
    var wrappers = (doc.body || doc).querySelectorAll("font, span");
    for (var i = 0; i < wrappers.length; i++) {
      var el = wrappers[i];
      while (el.firstChild) el.parentNode.insertBefore(el.firstChild, el);
      el.remove();
    }

    // Constrain images
    var imgs = (doc.body || doc).querySelectorAll("img");
    for (var i = 0; i < imgs.length; i++) {
      imgs[i].setAttribute("style", "max-width:100%;height:auto;");
    }

    var cleaned = (doc.body || doc).innerHTML || "";
    cleaned = cleaned.replace(/\n\s*\n\s*\n/g, "\n\n").trim();
    return cleaned;
  } catch (e) {
    console.warn("[NDL] HTML sanitize failed, using plain text fallback", e);
    return htmlToCleanText(rawHtml);
  }
}

// ── Helpers ────────────────────────────────────────────────────────
function setStatus(type, msg) {
  var el = document.getElementById("statusMsg");
  el.className = "status" + (type ? " " + type : "");
  el.textContent = msg;
  el.style.display = msg ? "block" : "none";
}

function getFileIcon(filename) {
  var ext = (filename || "").split(".").pop().toLowerCase();
  var icons = {
    pdf: "📄", doc: "📝", docx: "📝", xls: "📊", xlsx: "📊",
    ppt: "📽️", pptx: "📽️", png: "🖼️", jpg: "🖼️", jpeg: "🖼️",
    gif: "🖼️", zip: "📦", rar: "📦", txt: "📃", csv: "📊",
  };
  return icons[ext] || "📎";
}

// ── Resolve inline images: replace cid: with base64 data URLs ─────
function resolveInlineImages(item, callback) {
  var html = emailData.body;
  var attachments = item.attachments;
  if (!attachments || attachments.length === 0) {
    callback(html);
    return;
  }

  var inlineAtts = [];
  for (var i = 0; i < attachments.length; i++) {
    if (attachments[i].isInline && attachments[i].contentType && attachments[i].contentType.indexOf("image") === 0) {
      inlineAtts.push(attachments[i]);
    }
  }

  if (inlineAtts.length === 0) {
    callback(html);
    return;
  }

  var remaining = inlineAtts.length;

  for (var j = 0; j < inlineAtts.length; j++) {
    (function (att) {
      item.getAttachmentContentAsync(att.id, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          var dataUrl = "data:" + att.contentType + ";base64," + result.value.content;
          html = html.replace(new RegExp('src=["\']cid:' + escapeRegex(att.name) + '["\']', 'gi'), 'src="' + dataUrl + '"');
          if (att.id) {
            html = html.replace(new RegExp('src=["\']cid:[^"\']*' + escapeRegex(att.name.split('.')[0]) + '[^"\']*["\']', 'gi'), 'src="' + dataUrl + '"');
          }
        }
        remaining--;
        if (remaining === 0) {
          callback(html);
        }
      });
    })(inlineAtts[j]);
  }
}

function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function htmlToCleanText(html) {
  if (!html) return "";
  var text = html;
  // Remove style/script blocks (tag + content) before stripping tags
  text = text.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "");
  text = text.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "");
  text = text.replace(/<br\s*\/?>/gi, "\n");
  text = text.replace(/<\/p>/gi, "\n\n");
  text = text.replace(/<\/div>/gi, "\n");
  text = text.replace(/<\/h[1-6]>/gi, "\n\n");
  text = text.replace(/<\/li>/gi, "\n");
  text = text.replace(/<\/tr>/gi, "\n");
  text = text.replace(/<hr[^>]*>/gi, "\n---\n");
  text = text.replace(/<[^>]+>/g, "");
  text = text.replace(/&nbsp;/gi, " ");
  text = text.replace(/&amp;/gi, "&");
  text = text.replace(/&lt;/gi, "<");
  text = text.replace(/&gt;/gi, ">");
  text = text.replace(/&quot;/gi, '"');
  text = text.replace(/&#39;/gi, "'");
  text = text.replace(/&rsquo;/gi, "'");
  text = text.replace(/&lsquo;/gi, "'");
  text = text.replace(/&rdquo;/gi, '"');
  text = text.replace(/&ldquo;/gi, '"');
  text = text.replace(/&mdash;/gi, "—");
  text = text.replace(/&ndash;/gi, "–");
  text = text.replace(/&#58;/gi, ":");
  text = text.replace(/[ \t]+/g, " ");
  text = text.replace(/\n /g, "\n");
  text = text.replace(/ \n/g, "\n");
  text = text.replace(/\n{3,}/g, "\n\n");
  return text.trim();
}

function formatSize(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1048576) return (bytes / 1024).toFixed(0) + " KB";
  return (bytes / 1048576).toFixed(1) + " MB";
}
