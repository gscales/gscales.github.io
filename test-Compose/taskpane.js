// taskpane.js — Quick Recipients Outlook Add-in
// Adds pre-configured (or custom) recipients to the To field in compose mode.

"use strict";

// ── Pre-configured contacts ──────────────────────────────────────────────────
// Edit this array to populate the contacts shown in the pane.
// Supported address types:
//   smtp  — standard internet email address
//   ex    — Exchange EX / X.500 DN  (/O=.../OU=.../CN=...)
//   notes — Lotus Notes routable address (CN=.../O=...@domain)
//
// The displayAddress field is what Office.js receives; for EX addresses this is
// the raw DN string.  Outlook will resolve it against the GAL when the message
// is sent, exactly as if the user had typed it into the To field manually.
// ─────────────────────────────────────────────────────────────────────────────
const CONTACTS = [
  {
    name:           "John Doe",
    displayAddress: "john.doe@contoso.com",
    type:           "smtp",
    group:          "Internal",
  },
  {
    name:           "Sarah Jones",
    displayAddress: "sarah.jones@fabrikam.com",
    type:           "smtp",
    group:          "Partners",
  },
  {
    name:           "Support Team",
    displayAddress: "support@contoso.com",
    type:           "smtp",
    group:          "Distribution Lists",
  },
  {
    name:           "All Staff",
    displayAddress: "allstaff@contoso.com",
    type:           "smtp",
    group:          "Distribution Lists",
  },
];

// Avatar background colours — cycled by index so each contact gets a distinct colour
const AVATAR_COLOURS = [
  "#0078d4", "#107c10", "#c19c00", "#a4262c", "#8764b8",
  "#038387", "#ca5010", "#4f6bed", "#881798", "#00b294",
];

// ── State ────────────────────────────────────────────────────────────────────
let officeReady = false;
let addedSet    = new Set(); // tracks displayAddress values added this session

// ── Office.js initialisation ─────────────────────────────────────────────────
Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) {
    showStatus("This add-in only works in Outlook.", "error");
    return;
  }

  officeReady = true;
  renderContacts(CONTACTS);
  wireUp();
});

// ── Render ────────────────────────────────────────────────────────────────────
function renderContacts(contacts) {
  const list  = document.getElementById("recipientList");
  const empty = document.getElementById("emptyState");
  list.innerHTML = "";

  if (contacts.length === 0) {
    empty.classList.add("show");
    return;
  }
  empty.classList.remove("show");

  // Group contacts by their group property
  const groups = [];
  const seen   = new Map();
  for (const c of contacts) {
    if (!seen.has(c.group)) { seen.set(c.group, []); groups.push(c.group); }
    seen.get(c.group).push(c);
  }

  for (const groupName of groups) {
    // Group label (only show when more than one group is visible)
    if (groups.length > 1) {
      const lbl = document.createElement("div");
      lbl.className   = "group-label";
      lbl.textContent = groupName;
      list.appendChild(lbl);
    }

    for (const contact of seen.get(groupName)) {
      list.appendChild(buildCard(contact));
    }
  }
}

function buildCard(contact) {
  const idx      = CONTACTS.indexOf(contact) % AVATAR_COLOURS.length;
  const initials = contact.name.split(" ").map(w => w[0]).slice(0, 2).join("").toUpperCase();
  const isAdded  = addedSet.has(contact.displayAddress);

  const card = document.createElement("div");
  card.className   = "recipient-card" + (isAdded ? " added" : "");
  card.dataset.addr = contact.displayAddress;
  card.setAttribute("role", "button");
  card.setAttribute("tabindex", "0");
  card.setAttribute("aria-label", `Add ${contact.name} to To field`);

  card.innerHTML = `
    <div class="avatar" style="background:${AVATAR_COLOURS[idx]}">${initials}</div>
    <div class="recipient-info">
      <div class="recipient-name">${escHtml(contact.name)}</div>
      <div class="recipient-address">${escHtml(contact.displayAddress)}</div>
    </div>
    <span class="recipient-tag tag-${contact.type}">${contact.type.toUpperCase()}</span>
    <div class="add-icon${isAdded ? " done" : ""}">
      ${isAdded ? "" : '<svg viewBox="0 0 16 16"><path d="M8 2a.75.75 0 01.75.75v4.5h4.5a.75.75 0 010 1.5h-4.5v4.5a.75.75 0 01-1.5 0v-4.5h-4.5a.75.75 0 010-1.5h4.5v-4.5A.75.75 0 018 2z"/></svg>'}
    </div>
  `;

  card.addEventListener("click",  () => addRecipient(contact.displayAddress, contact.name, card));
  card.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") addRecipient(contact.displayAddress, contact.name, card); });

  return card;
}

// ── Wire up search and custom input ──────────────────────────────────────────
function wireUp() {
  // Live search / filter
  document.getElementById("search").addEventListener("input", function () {
    const q = this.value.toLowerCase().trim();
    const filtered = q
      ? CONTACTS.filter(c =>
          c.name.toLowerCase().includes(q) ||
          c.displayAddress.toLowerCase().includes(q) ||
          c.group.toLowerCase().includes(q))
      : CONTACTS;
    renderContacts(filtered);
  });

  // Custom address button
  const input = document.getElementById("customAddress");
  const btn   = document.getElementById("btnAddCustom");

  btn.addEventListener("click", () => {
    const addr = input.value.trim();
    if (!addr) { showStatus("Please enter an email address.", "error"); return; }
    addRecipient(addr, addr, null);
    input.value = "";
  });

  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") btn.click();
  });

  input.addEventListener("input", () => {
    btn.disabled = input.value.trim().length === 0;
  });
  btn.disabled = true;
}

// ── Core: add recipient via Office.js ─────────────────────────────────────────
function addRecipient(address, displayName, cardEl) {
  if (!officeReady) { showStatus("Office.js is not ready yet.", "error"); return; }

  const item = Office.context.mailbox.item;

  if (!item || !item.to) {
    showStatus("No compose message is active.", "error");
    return;
  }

  if (cardEl) {
    cardEl.classList.add("adding");
    cardEl.setAttribute("aria-busy", "true");
  }

  // Build the recipient object.
  // For EX addresses (slash-delimited DNs), Office.js accepts them verbatim as
  // the emailAddress — Outlook resolves the DN against the GAL on send.
  // For standard SMTP, using EmailUser with both name and address gives the
  // best rendering in the To field (shows "Display Name <addr>").
  const recipient = address.startsWith("/")
    ? address                                    // EX DN — pass as plain string
    : { displayName: displayName, emailAddress: address }; // SMTP — named object

  item.to.addAsync(
    [recipient],
    { asyncContext: { address, cardEl } },
    function (asyncResult) {
      const ctx = asyncResult.asyncContext;

      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        addedSet.add(ctx.address);

        // Update card appearance
        if (ctx.cardEl) {
          ctx.cardEl.classList.remove("adding");
          ctx.cardEl.classList.add("added");
          ctx.cardEl.removeAttribute("aria-busy");
          const icon = ctx.cardEl.querySelector(".add-icon");
          if (icon) {
            icon.classList.add("done");
            icon.innerHTML = ""; // checkmark rendered via CSS ::after
          }
        }

        showStatus(`Added: ${ctx.address}`, "success");

      } else {
        if (ctx.cardEl) {
          ctx.cardEl.classList.remove("adding");
          ctx.cardEl.removeAttribute("aria-busy");
        }

        const msg = asyncResult.error
          ? asyncResult.error.message
          : "Unknown error.";
        showStatus(`Failed to add recipient: ${msg}`, "error");
      }
    }
  );
}

// ── Status banner ─────────────────────────────────────────────────────────────
let statusTimer = null;

function showStatus(message, type) {
  const el = document.getElementById("status");
  el.textContent  = message;
  el.className    = `show ${type}`;

  clearTimeout(statusTimer);
  statusTimer = setTimeout(() => {
    el.className = "";
    el.textContent = "";
  }, 4000);
}

// ── Utility ───────────────────────────────────────────────────────────────────
function escHtml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
