// =============================================
// CONFIG - ISI SESUAI DATA ANDA
// =============================================
var CONFIG = {
  clientId:      "c87df5fb-eb2f-40bc-8c33-d0c47dedfe10",
  tenantId:      "c1cc324d-9fda-4a48-97ee-ba6157ba9c67",
  siteUrl:       "https://gosgroup.sharepoint.com/sites/procurement",
  masterList:    "MasterPengadaan",
  resultList:    "HasilPengajuan",
  approverEmail: "muhamad.parid@gos.co.id",
  redirectUri:   window.location.origin + window.location.pathname
};

// =============================================
// APPROVERS - yang boleh update status
// =============================================
var APPROVERS = [
  "muhamad.parid@gos.co.id",
  "toro.sod@gos.co.id",
  "mis@gos.co.id"
];

// =============================================
// APPROVAL CONFIG - Sesuaikan dengan kebutuhan
// =============================================
var APPROVAL_CONFIG = {
  l2Threshold: 5000000,
  l2: { email: "", name: "" }, // Isi via Setting Approval di portal
  rules:      [],  // { project, l1Email, l1Name, l2Email, l2Name }
  submitters: []   // email yang boleh submit
};

// =============================================
// APPROVAL HELPERS
// =============================================
function myEmail() { return currentUser ? currentUser.userPrincipalName.toLowerCase() : ""; }
function amGA()    { return APPROVERS.indexOf(myEmail()) > -1; }
function amL2() {
  var email = myEmail();
  return APPROVAL_CONFIG.rules.some(function(r){ return r.l2Email && r.l2Email.toLowerCase()===email; });
}
function amL1()    {
  return APPROVAL_CONFIG.rules.some(function(r) {
    return r.l1Email.toLowerCase() === myEmail();
  });
}
function getL1ByProject(projectName, cabang) {
  var p = (projectName||"").toLowerCase().trim();
  var c = (cabang||"").toLowerCase().trim();
  // 1. Cari rule project + cabang exact match
  var r = APPROVAL_CONFIG.rules.filter(function(x){
    return (x.project||"").toLowerCase().trim() === p &&
           (x.cabang||"").toLowerCase().trim() === c;
  })[0];
  // 2. Fallback: project match tanpa cabang (catch-all)
  if (!r && c) r = APPROVAL_CONFIG.rules.filter(function(x){
    return (x.project||"").toLowerCase().trim() === p && !(x.cabang||"").trim();
  })[0];
  // 3. Fallback: project match saja (ignore cabang)
  if (!r) r = APPROVAL_CONFIG.rules.filter(function(x){
    return (x.project||"").toLowerCase().trim() === p;
  })[0];
  return r || null;
}
function hasSubmitAccess() {
  var email = myEmail();
  if (amGA()) return true;
  return APPROVAL_CONFIG.submitters.some(function(s){
    return s.toLowerCase() === email;
  });
}
function needsL2(amount, rule) {
  var r = rule || {};
  if (!r.l2Email) return false;
  return (parseInt(amount)||0) >= APPROVAL_CONFIG.l2Threshold;
}
function canApproveL1(item) {
  return amL1() && item.Status==="Pending L1" &&
         (item.L1ApproverEmail||"").toLowerCase()===myEmail();
}
function canApproveL2(item) {
  return item.Status==="Pending L2" &&
         (item.L2ApproverEmail||"").toLowerCase()===myEmail();
}
function canUpdateGA(item)  {
  return amGA() && ["Approved","Submitted to Finance","Delivered","Rejected"].indexOf(item.Status)>-1;
}

function isApprover() {
  if (!currentUser) return false;
  return APPROVERS.indexOf(currentUser.userPrincipalName.toLowerCase()) > -1;
}
var _msal = null;
var currentUser = null;
var siteId = null;
var SCOPES = [
  "https://graph.microsoft.com/Sites.ReadWrite.All",
  "https://graph.microsoft.com/Mail.Send",
  "https://graph.microsoft.com/User.Read"
];

function initMsal() {
  var cfg = {
    auth: {
      clientId: CONFIG.clientId,
      authority: "https://login.microsoftonline.com/" + CONFIG.tenantId,
      redirectUri: CONFIG.redirectUri
    },
    cache: { cacheLocation: "sessionStorage" }
  };
  _msal = new msal.PublicClientApplication(cfg);
  return _msal.initialize();
}

function signIn() {
  _msal.loginPopup({ scopes: SCOPES }).then(function(r) {
    _msal.setActiveAccount(r.account);
    initApp();
  }).catch(function(e) {
    // Fallback ke redirect jika popup diblokir (email iframe, browser setting, dll)
    if (e.errorCode === "popup_window_error" || e.errorCode === "empty_window_error") {
      _msal.loginRedirect({ scopes: SCOPES });
    } else {
      showToast("Login gagal: " + (e.message || e), "er");
    }
  });
}

function signOut() {
  _msal.logoutPopup();
  document.getElementById("ls").style.display = "flex";
  document.getElementById("app").classList.remove("on");
}

function getToken() {
  var acc = _msal.getActiveAccount();
  if (!acc) return Promise.reject(new Error("Tidak ada akun aktif"));
  return _msal.acquireTokenSilent({ scopes: SCOPES, account: acc })
    .then(function(r) { return r.accessToken; })
    .catch(function() {
      return _msal.acquireTokenPopup({ scopes: SCOPES })
        .then(function(r) { return r.accessToken; });
    });
}

// =============================================
// GRAPH API
// =============================================
function gGet(url) {
  return getToken().then(function(t) {
    return fetch(url, { headers: { Authorization: "Bearer " + t } });
  }).then(function(r) {
    if (!r.ok) return r.text().then(function(e) { throw new Error("GET " + r.status + ": " + e); });
    return r.json();
  });
}

function gPost(url, body) {
  return getToken().then(function(t) {
    return fetch(url, {
      method: "POST",
      headers: { Authorization: "Bearer " + t, "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });
  }).then(function(r) {
    if (!r.ok) return r.text().then(function(e) { throw new Error("POST " + r.status + ": " + e); });
    return (r.status === 204 || r.status === 202) ? {} : r.json();
  });
}

function gPatch(url, body) {
  return getToken().then(function(t) {
    return fetch(url, {
      method: "PATCH",
      headers: { Authorization: "Bearer " + t, "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });
  }).then(function(r) {
    if (!r.ok) return r.text().then(function(e) { throw new Error("PATCH " + r.status + ": " + e); });
    return (r.status === 204 || r.status === 202) ? {} : r.json();
  });
}

function getSid() {
  if (siteId) return Promise.resolve(siteId);
  var u = new URL(CONFIG.siteUrl);
  return gGet("https://graph.microsoft.com/v1.0/sites/" + u.hostname + ":" + u.pathname)
    .then(function(r) { siteId = r.id; return siteId; });
}

function getItems(list) {
  return getSid().then(function(sid) {
    var items = [];
    function fetchPage(url) {
      return gGet(url).then(function(r) {
        items = items.concat(r.value || []);
        if (r["@odata.nextLink"]) return fetchPage(r["@odata.nextLink"]);
        return items;
      });
    }
    return fetchPage("https://graph.microsoft.com/v1.0/sites/" + sid + "/lists/" + list + "/items?$expand=fields&$top=999");
  });
}

function createItem(list, fields) {
  return getSid().then(function(sid) {
    return gPost("https://graph.microsoft.com/v1.0/sites/" + sid + "/lists/" + list + "/items", { fields: fields });
  });
}

function patchItem(list, id, fields) {
  return getSid().then(function(sid) {
    return gPatch("https://graph.microsoft.com/v1.0/sites/" + sid + "/lists/" + list + "/items/" + id + "/fields", fields);
  });
}

// =============================================
// MASTER DATA
// =============================================
var masterData = [];
var CODES = {
  "PT. PRIMA RAYA SOLUSINDO": "PRS",
  "PT. OTSINDO PRIMA RAYA":   "OPR",
  "PT. DINAMIKA NUANSA ABSOLUTE": "DNA",
  "PT. FAS INDO RAYA":        "FAS"
};
var ROMAN = ["","I","II","III","IV","V","VI","VII","VIII","IX","X","XI","XII"];

function getCompanyCode(co) {
  // Cari dari masterData (CompanyCode field) dulu
  var r = (masterData||[]).filter(function(d){ return d.Company === co; })[0];
  if (r && r.CompanyCode) return r.CompanyCode;
  // Fallback ke CODES dictionary (case-insensitive)
  var key = Object.keys(CODES).filter(function(k){ return k.toUpperCase() === (co||"").toUpperCase(); })[0];
  return key ? CODES[key] : "___";
}

function loadMaster() {
  return getItems(CONFIG.masterList).then(function(items) {
    masterData = items.map(function(i) { return i.fields; })
                      .filter(function(r) { return r.Aktif !== false && r.Aktif !== 0; });
  });
}

function getCos() {
  var seen = {};
  return masterData.map(function(i) { return i.Company; })
    .filter(function(c) { return c && !seen[c] && (seen[c] = true); }).sort();
}
function getCls(co) {
  var seen = {};
  return masterData.filter(function(i) { return i.Company === co; })
    .map(function(i) { return i.Client; })
    .filter(function(c) { return c && !seen[c] && (seen[c] = true); }).sort();
}
function getPrs(co, cl) {
  var seen = {};
  return masterData.filter(function(i) { return i.Company === co && i.Client === cl; })
    .map(function(i) { return i.Project; })
    .filter(function(p) { return p && !seen[p] && (seen[p] = true); }).sort();
}

// =============================================
// FORM
// =============================================
var uploadedFiles = [];

function initDate() {
  var d = new Date();
  document.getElementById("f-tgl").value =
    pad(d.getDate()) + "/" + pad(d.getMonth()+1) + "/" + d.getFullYear();
}

function pad(n) { return n < 10 ? "0" + n : "" + n; }

// Searchable Dropdown
function populateCos() {
  renderSDDOpts("co", getCos(), "");
}

function renderSDDOpts(key, items, q) {
  var opts = document.getElementById("sdd-" + key + "-opts");
  if (!opts) return;
  var f = q ? items.filter(function(i){ return i.toLowerCase().indexOf(q.toLowerCase()) > -1; }) : items;
  if (!f.length) { opts.innerHTML = '<div class="sdd-none">Tidak ditemukan</div>'; return; }
  var cur = document.getElementById("f-" + key).value;
  opts.innerHTML = f.map(function(i) {
    return '<div class="sdd-item' + (i===cur?' ssel':'') + '" onmousedown="selectSDD(\'' + key + '\',\'' + i.replace(/'/g,"\\'") + '\')">' + i + '</div>';
  }).join("");
}

function openSDD(key) {
  // Guard: cek prerequisite
  if (key === "cl" && !document.getElementById("f-co").value) {
    showToast("Pilih Company terlebih dahulu", "er"); return;
  }
  if (key === "pr" && !document.getElementById("f-cl").value) {
    showToast("Pilih Client terlebih dahulu", "er"); return;
  }
  var list  = document.getElementById("sdd-" + key + "-list");
  var srch  = document.getElementById("sdd-" + key + "-search");
  // Close others
  ["co","cl","pr"].forEach(function(k) { if (k !== key) closeSDD(k); });
  var items = key==="co" ? getCos() : key==="cl" ? getCls(document.getElementById("f-co").value) : getPrs(document.getElementById("f-co").value, document.getElementById("f-cl").value);
  renderSDDOpts(key, items, "");
  list.style.display = "block";
  document.getElementById("f-" + key + "-txt").classList.add("open");
  if (srch) { srch.value = ""; srch.focus(); }
}

function closeSDD(key) {
  var list  = document.getElementById("sdd-" + key + "-list");
  var input = document.getElementById("f-" + key + "-txt");
  if (list)  list.style.display = "none";
  if (input) input.classList.remove("open");
}

function filterSDD(key) {
  var q = document.getElementById("sdd-" + key + "-search").value;
  var items = key==="co" ? getCos() : key==="cl" ? getCls(document.getElementById("f-co").value) : getPrs(document.getElementById("f-co").value, document.getElementById("f-cl").value);
  renderSDDOpts(key, items, q);
}

function selectSDD(key, val) {
  document.getElementById("f-" + key).value          = val;
  document.getElementById("f-" + key + "-txt").value = val;
  closeSDD(key);
  if (key === "co") {
    document.getElementById("f-cl").value     = "";
    document.getElementById("f-cl-txt").value = "";
    document.getElementById("f-pr").value     = "";
    document.getElementById("f-pr-txt").value = "";
    document.getElementById("f-cabang").value = "";
    document.getElementById("cabang-dd").style.display = "none";
    updateRn();
  }
  if (key === "cl") {
    document.getElementById("f-pr").value     = "";
    document.getElementById("f-pr-txt").value = "";
    document.getElementById("f-cabang").value = "";
    document.getElementById("cabang-dd").style.display = "none";
  }
  if (key === "pr") {
    document.getElementById("f-cabang").value = "";
    document.getElementById("cabang-dd").style.display = "none";
  }
}

function onSDDKey(e, key) {
  if (e.key === "Escape") closeSDD(key);
}

// Close dropdowns when clicking outside
document.addEventListener("click", function(e) {
  ["co","cl","pr"].forEach(function(key) {
    var sdd = document.getElementById("sdd-" + key);
    if (sdd && !sdd.contains(e.target)) closeSDD(key);
  });
});

function onCo() { /* kept for compatibility */ }
function onCl() { /* kept for compatibility */ }



function genSuffix() {
  var chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  var s = "";
  for (var i = 0; i < 3; i++) s += chars[Math.floor(Math.random() * chars.length)];
  return s;
}

function updateRn() {
  var seq = (document.getElementById("f-seq").value || "").replace(/\D/g,"");
  while (seq.length < 4) seq = "0" + seq;
  var co = document.getElementById("f-co").value;
  var j  = document.getElementById("f-jenis").value;
  var d  = new Date();
  document.getElementById("rnd").textContent =
    seq + "/" + getCompanyCode(co) + "/" + (j || "___") + "/" + ROMAN[d.getMonth()+1] + "/" + d.getFullYear();
}

function copyRn() {
  var v = document.getElementById("rnd").textContent;
  navigator.clipboard.writeText(v).then(function() { showToast("Nomor disalin!", "ok"); });
}

function autoSeq() {
  return getSid().then(function(sid) {
    var url = "https://graph.microsoft.com/v1.0/sites/" + sid + "/lists/" + CONFIG.resultList +
              "/items?$expand=fields($select=NomorUrut)&$orderby=fields/NomorUrut%20desc&$top=1";
    return gGet(url);
  }).then(function(r) {
    var last = (r.value && r.value[0] && r.value[0].fields && r.value[0].fields.NomorUrut) || 0;
    var next = "" + (parseInt(last, 10) + 1);
    while (next.length < 4) next = "0" + next;
    document.getElementById("f-seq").value = next;
    updateRn();
  }).catch(function(e) {
    console.warn("Auto seq gagal:", e.message);
  });
}

function fmtHg(el) {
  var v = el.value.replace(/\D/g, "");
  el.value = v ? parseInt(v, 10).toLocaleString("id-ID") : "";
}

function onFiles(files) {
  Array.from(files).forEach(function(f) {
    if (f.size > 10485760) { showToast(f.name + " terlalu besar", "er"); return; }
    if (!uploadedFiles.find(function(x) { return x.name === f.name; })) uploadedFiles.push(f);
  });
  renderFiles();
}
function onDrop(e) {
  e.preventDefault();
  document.getElementById("dz").classList.remove("drag");
  onFiles(e.dataTransfer.files);
}
function rmFile(i) { uploadedFiles.splice(i, 1); renderFiles(); }
function renderFiles() {
  document.getElementById("flist").innerHTML = uploadedFiles.map(function(f, i) {
    var s = f.size > 1048576 ? (f.size/1048576).toFixed(1) + " MB" : Math.round(f.size/1024) + " KB";
    return '<div class="fi"><span class="fn">&#128196; ' + f.name + '</span>' +
           '<span style="display:flex;align-items:center;gap:8px"><span style="font-size:11px;color:var(--g500)">' + s + '</span>' +
           '<button class="fr" onclick="rmFile(' + i + ')">x</button></span></div>';
  }).join("");
}

function validate() {
  var ok = true;
  [["f-jenis","Pilih jenis pengadaan"],["f-co","Pilih company"],["f-cl","Pilih client"],
   ["f-pr","Pilih project"],["f-tj","Wajib diisi"],["f-hg","Wajib diisi"]].forEach(function(r) {
    var el = document.getElementById(r[0]);
    var fd = el.closest(".fd");
    var fe = fd.querySelector(".fe");
    if (!el.value.trim()) {
      fd.classList.add("err"); if (fe) fe.textContent = r[1]; ok = false;
    } else {
      fd.classList.remove("err");
    }
  });
  if (!document.querySelector('input[name="jp"]:checked')) {
    document.querySelector(".rg").closest(".fd").classList.add("err"); ok = false;
  } else {
    document.querySelector(".rg").closest(".fd").classList.remove("err");
  }
  // Dokumen tidak wajib — reset border jika ada
  var fz = document.getElementById("dz");
  if (fz) { fz.style.borderColor = ""; fz.style.background = ""; }
  var jp = document.querySelector('input[name="jp"]:checked');
  if (jp && jp.value === "Barang") {
    var filledB = rBarang.filter(function(r) { return r.nama && r.jumlah; });
    if (!filledB.length) { showToast("Tambahkan minimal 1 barang dengan nama dan jumlah", "er"); ok = false; }
  }
  if (jp && jp.value === "Jasa") {
    var filledJ = rJasa.filter(function(r) { return r.vendor && r.nominal; });
    if (!filledJ.length) { showToast("Tambahkan minimal 1 vendor dengan nama dan nominal", "er"); ok = false; }
  }
  return ok;
}

function submitForm() {
  if (!validate()) {
    showToast("Lengkapi semua field yang wajib diisi", "er");
    var first = document.querySelector(".fd.err");
    if (first) first.scrollIntoView({ behavior: "smooth", block: "center" });
    return;
  }
  var btn = document.getElementById("bsub");
  btn.innerHTML = '<span class="spin"></span> Mengupload file...';
  btn.disabled = true;

  var d      = new Date();
  var suffix = genSuffix();
  var seqNum = parseInt((document.getElementById("f-seq").value||"1").replace(/\D/g,""),10) || 1;
  var seq    = String(seqNum); while (seq.length < 4) seq = "0" + seq;
  var co     = document.getElementById("f-co").value;
  var j      = document.getElementById("f-jenis").value;
  var rn     = seq + "-" + suffix + "/" + getCompanyCode(co) + "/" + (j||"___") + "/" + ROMAN[d.getMonth()+1] + "/" + d.getFullYear();
  var hgRaw = document.getElementById("f-hg").value.replace(/\./g,"").replace(/,/g,"");
  var hg  = parseInt(hgRaw, 10) || 0;
  var jp  = document.querySelector('input[name="jp"]:checked');
  var selectedProject = document.getElementById("f-pr").value;
  var selectedCabang  = document.getElementById("f-cabang").value.trim();
  var l1rule = getL1ByProject(selectedProject, selectedCabang);

  var storedDok = "";
  uploadAllFiles(uploadedFiles, rn).then(function(fileResults) {
    btn.innerHTML = '<span class="spin"></span> Menyimpan...';
    storedDok = fileResults.length ? JSON.stringify(fileResults) : "";
    var dokStr = storedDok;
    if (!l1rule) { showToast("Project ini belum memiliki rule approval" + (selectedCabang ? " untuk cabang " + selectedCabang : "") + ". Hubungi Admin.", "er"); btn.innerHTML="Kirim Pengajuan"; btn.disabled=false; return; }
    var fields = {
      Title:            rn,
      NomorPengajuan:   rn,
      NomorUrut:        parseInt(document.getElementById("f-seq").value, 10) || 1,
      TanggalPengajuan: d.toISOString().split("T")[0],
      JenisPengadaan:   document.getElementById("f-jenis").value,
      Company:          document.getElementById("f-co").value,
      Client:           document.getElementById("f-cl").value,
      Project:          document.getElementById("f-pr").value,
      Cabang:           selectedCabang,
      TujuanPermintaan: document.getElementById("f-tj").value,
      JenisProduk:      jp ? jp.value : "",
      EstimasiHarga:    hg,
      DokumenPendukung: dokStr,
      Status:           "Pending L1",
      SubmittedBy:      currentUser ? currentUser.displayName       : "",
      SubmittedByEmail: currentUser ? currentUser.userPrincipalName : "",
      L1ApproverEmail:  l1rule ? l1rule.l1Email : "",
      L1ApproverName:   l1rule ? l1rule.l1Name  : ""
    };
    if (l1rule && l1rule.l2Email) {
      fields.L2ApproverEmail = l1rule.l2Email;
      fields.L2ApproverName  = l1rule.l2Name || "";
    }
    var itemJson = getItemsJSON();
    if (itemJson) fields.DetailItem = itemJson;
    return createItem(CONFIG.resultList, fields);
  }).then(function(res) {
    // Cek duplikat NomorUrut setelah submit
    var submittedSeq = parseInt(document.getElementById("f-seq").value, 10) || 1;
    var newItemId = res && res.id ? res.id : null;
    fixDuplicate(submittedSeq, newItemId);

    var l1Email  = l1rule ? l1rule.l1Email : CONFIG.approverEmail;
    sendNotif("new", rn, {
      co: document.getElementById("f-co").value,
      cl: document.getElementById("f-cl").value,
      pr: document.getElementById("f-pr").value,
      tj: document.getElementById("f-tj").value,
      jp: jp ? jp.value : "",
      hg: hg, to: l1Email,
      l1Name:    l1rule ? l1rule.l1Name : "",
      submitter: currentUser ? currentUser.displayName : "",
      itemId:    newItemId,
      dokStr:    storedDok,
      tgl:       new Date().toLocaleDateString("id-ID")
    });
    showToast("Pengajuan " + rn + " berhasil dikirim!", "ok");
    resetForm();
    return autoSeq();
  }).catch(function(e) {
    showToast("Gagal: " + e.message, "er");
    console.error(e);
  }).finally(function() {
    btn.innerHTML = "Kirim Pengajuan";
    btn.disabled = false;
  });
}

// Auto-fix duplicate NomorUrut setelah submit
function fixDuplicate(seq, newId) {
  if (!newId) return;
  getSid().then(function(sid) {
    // Ambil semua item dengan NomorUrut yang sama
    var url = "https://graph.microsoft.com/v1.0/sites/" + sid + "/lists/" + CONFIG.resultList +
              "/items?$expand=fields($select=NomorUrut,NomorPengajuan,id)&$filter=fields/NomorUrut eq " + seq +
              "&$top=10";
    return gGet(url);
  }).then(function(r) {
    var dupes = (r.value || []).filter(function(i) {
      return i.id !== newId && i.fields && i.fields.NomorUrut == seq;
    });
    if (dupes.length === 0) return; // Tidak ada duplikat

    // Ada duplikat - cari NomorUrut terbesar lalu increment untuk item baru
    return getSid().then(function(sid) {
      var url2 = "https://graph.microsoft.com/v1.0/sites/" + sid + "/lists/" + CONFIG.resultList +
                 "/items?$expand=fields($select=NomorUrut)&$orderby=fields/NomorUrut desc&$top=1";
      return gGet(url2);
    }).then(function(r2) {
      var maxSeq = (r2.value && r2.value[0] && r2.value[0].fields && r2.value[0].fields.NomorUrut) || seq;
      var newSeq = parseInt(maxSeq) + 1;
      var newSeqStr = String(newSeq).padStart ? String(newSeq) : ("000" + newSeq).slice(-4);
      while (newSeqStr.length < 4) newSeqStr = "0" + newSeqStr;

      // Update item yang baru dibuat dengan NomorUrut yang benar
      return getSid().then(function(sid) {
        var d = new Date();
        var co = "";
        // Rebuild request number dengan seq baru
        // Ambil company code dari NomorPengajuan lama
        var oldRn = document.getElementById("rnd") ? document.getElementById("rnd").textContent : "";
        var parts  = oldRn.split("/");
        if (parts.length >= 5) {
          parts[0] = newSeqStr;
          var newRn = parts.join("/");
          return patchItem(CONFIG.resultList, newId, { NomorUrut: newSeq, NomorPengajuan: newRn, Title: newRn });
        }
      });
    });
  }).catch(function(e) {
    console.warn("fixDuplicate gagal:", e.message);
  });
}

function resetForm() {
  ["f-jenis","f-tj","f-hg","f-cabang"].forEach(function(id) {
    document.getElementById(id).value = "";
  });
  document.querySelectorAll('input[name="jp"]').forEach(function(r) { r.checked = false; });
  document.getElementById("f-co").value = ""; document.getElementById("f-co-txt").value = "";
  document.getElementById("f-cl").value = ""; document.getElementById("f-cl-txt").value = "";
  document.getElementById("f-pr").value = ""; document.getElementById("f-pr-txt").value = "";
  closeSDD("co"); closeSDD("cl"); closeSDD("pr");
  document.querySelectorAll(".fd.err").forEach(function(fd) { fd.classList.remove("err"); });
  rBarang = []; rJasa = [];
  document.getElementById("rows-barang").innerHTML = "";
  document.getElementById("rows-jasa").innerHTML   = "";
  document.getElementById("tot-barang").textContent = "Rp 0";
  document.getElementById("tot-jasa").textContent   = "Rp 0";
  document.getElementById("sec-barang").style.display = "none";
  document.getElementById("sec-jasa").style.display   = "none";
  uploadedFiles = []; renderFiles(); updateRn();
}

// =============================================
// ITEM TABLES - BARANG & JASA
// =============================================
var mBarang = [], mJasa = [];
var rBarang = [], rJasa = [];
var SB = ["pcs","unit","box","set","lusin","kodi","rim","rol","meter","kg","liter","buah","pasang","lembar"];
var SJ = ["paket","hari","minggu","bulan","tahun","unit","kali","jam","orang","lot","lokasi"];

function loadMasterBarang() {
  return getItems("MasterBarang").then(function(items) {
    mBarang = items.map(function(i) { return i.fields; })
                   .filter(function(b) { return b.Aktif !== false && b.Aktif !== 0; });
    var dl = document.getElementById("dl-barang");
    if (dl) dl.innerHTML = mBarang.map(function(b) {
      return '<option value="' + eA(b.NamaBarang||"") + '">';
    }).join("");
  }).catch(function(e) { console.warn("MasterBarang gagal:", e.message); });
}
function loadMasterJasa() {
  return getItems("MasterJasa").then(function(items) {
    mJasa = items.map(function(i) { return i.fields; })
                 .filter(function(j) { return j.Aktif !== false && j.Aktif !== 0; });
    var dl = document.getElementById("dl-jasa");
    if (dl) dl.innerHTML = mJasa.map(function(j) {
      return '<option value="' + eA(j.NamaVendor||"") + '">';
    }).join("");
  }).catch(function(e) { console.warn("MasterJasa gagal:", e.message); });
}

function eA(s) { return (s||"").replace(/"/g,"&quot;").replace(/'/g,"&#39;"); }
function eQ(s) { return (s||"").replace(/\\/g,"\\\\").replace(/'/g,"\\'"); }

function onJP() {
  var jp = document.querySelector('input[name="jp"]:checked');
  var v  = jp ? jp.value : "";
  document.getElementById("sec-barang").style.display = v === "Barang" ? "block" : "none";
  document.getElementById("sec-jasa").style.display   = v === "Jasa"   ? "block" : "none";
  if (v === "Barang" && rBarang.length === 0) addRow("barang");
  if (v === "Jasa"   && rJasa.length   === 0) addRow("jasa");
}

function addRow(type) {
  if (type === "barang") { rBarang.push({ nama:"", kategori:"", subkat:"", satuan:"", hargaEst:0, jumlah:"", total:0 }); renderRows("barang"); }
  else { rJasa.push({ vendor:"", kategori:"", domisili:"", pic:"", telp:"", nominal:0, keterangan:"" }); renderRows("jasa"); }
}

function delRow(type, i) {
  if (type === "barang") rBarang.splice(i,1); else rJasa.splice(i,1);
  renderRows(type); calcTot(type);
}

function sopts(list, sel) {
  return list.map(function(s) { return '<option value="'+s+'"'+(s===sel?' selected':'')+'>'+s+'</option>'; }).join("");
}

function renderRows(type) {
  var rows  = type === "barang" ? rBarang : rJasa;
  var tbody = document.getElementById("rows-" + type);
  if (!rows.length) { tbody.innerHTML = '<tr><td colspan="'+(type==="barang"?11:9)+'" style="text-align:center;padding:16px;color:var(--g500);font-style:italic">Belum ada item. Klik "+ Tambah" untuk menambahkan.</td></tr>'; calcTot(type); return; }

  tbody.innerHTML = rows.map(function(r, i) {
    var sub  = "";
    var hEst = r.hargaEst ? "Rp "+r.hargaEst.toLocaleString("id-ID") : "-";
    if (type === "barang") {
      sub = r.total ? "Rp "+r.total.toLocaleString("id-ID") : "-";
      return '<tr>'+
        '<td class="td-idx">'+(i+1)+'</td>'+
        '<td><input list="dl-barang" class="ti" value="'+eA(r.nama)+'" placeholder="Cari nama barang..." oninput="oii(\'barang\','+i+',\'nama\',this.value)"></td>'+
        '<td><input class="ti" value="'+eA(r.kategori)+'" readonly style="background:var(--g50);color:var(--g700)" placeholder="-"></td>'+
        '<td><input class="ti" value="'+eA(r.subkat)+'" readonly style="background:var(--g50);color:var(--g700)" placeholder="-"></td>'+
        '<td><input class="ti" value="'+eA(r.satuan)+'" readonly style="background:var(--g50);color:var(--g700)" placeholder="-"></td>'+
        '<td><input class="ti tr" value="'+hEst+'" readonly style="background:var(--g50);color:var(--navy);font-weight:600" placeholder="-"></td>'+
        '<td><input class="ti tr" type="number" min="1" value="'+(r.jumlah||"")+'" placeholder="0" oninput="oiq(\'barang\','+i+',this.value)"></td>'+
        '<td class="td-sub">'+sub+'</td>'+
        '<td><input class="ti" value="'+eA(r.keterangan||"")+'" placeholder="Keterangan..." oninput="oii(\'barang\','+i+',\'keterangan\',this.value)"></td>'+
        '<td class="td-del"><button class="del-r" onclick="delRow(\'barang\','+i+')" title="Hapus">x</button></td>'+
        '</tr>';
    } else {
      var nom = r.nominal ? "Rp "+r.nominal.toLocaleString("id-ID") : "";
      return '<tr>'+
        '<td class="td-idx">'+(i+1)+'</td>'+
        '<td><input list="dl-jasa" class="ti" value="'+eA(r.vendor)+'" placeholder="Cari vendor..." oninput="oii(\'jasa\','+i+',\'vendor\',this.value)"></td>'+
        '<td><input class="ti" value="'+eA(r.kategori)+'" readonly style="background:var(--g50);color:var(--g700)" placeholder="-"></td>'+
        '<td><input class="ti" value="'+eA(r.domisili)+'" readonly style="background:var(--g50);color:var(--g700)" placeholder="-"></td>'+
        '<td><input class="ti" value="'+eA(r.pic)+'" readonly style="background:var(--g50);color:var(--g700)" placeholder="-"></td>'+
        '<td><input class="ti" value="'+eA(r.telp)+'" readonly style="background:var(--g50);color:var(--g700)" placeholder="-"></td>'+
        '<td><input class="ti tr" type="text" value="'+nom+'" placeholder="0" onblur="oih(\'jasa\','+i+',this.value)"></td>'+
        '<td><input class="ti" value="'+eA(r.keterangan||"")+'" placeholder="Keterangan..." oninput="oii(\'jasa\','+i+',\'keterangan\',this.value)"></td>'+
        '<td class="td-del"><button class="del-r" onclick="delRow(\'jasa\','+i+')" title="Hapus">x</button></td>'+
        '</tr>';
    }
  }).join("");
}

function oii(type, i, field, val) {
  var rows = type === "barang" ? rBarang : rJasa;
  rows[i][field] = val;
  if (field === "nama" || field === "vendor") {
    if (type === "barang") {
      var m = mBarang.filter(function(b){ return b.NamaBarang===val; })[0];
      if (m) {
        rows[i].kategori = m.KategoriBarang || "";
        rows[i].subkat   = m.Subkategori    || "";
        rows[i].satuan   = m.Satuan         || "";
        rows[i].hargaEst = parseInt(m.HargaEstimasi) || 0;
        rows[i].total    = rows[i].hargaEst * (parseInt(rows[i].jumlah) || 0);
        renderRows("barang");
        calcTot("barang");
      }
    } else {
      var m = mJasa.filter(function(j){ return j.NamaVendor===val; })[0];
      if (m) {
        rows[i].kategori = m.Kategori       || "";
        rows[i].domisili = m.DomisiliVendor || "";
        rows[i].pic      = m.PICVendor      || "";
        rows[i].telp     = m.NoTelpon       || "";
        renderRows("jasa");
      }
    }
  }
}

function oiq(type, i, val) {
  var rows = type === "barang" ? rBarang : rJasa;
  rows[i].jumlah = val;
  rows[i].total  = (parseInt(val)||0) * (rows[i].hargaEst||0);
  var sub = rows[i].total ? "Rp "+rows[i].total.toLocaleString("id-ID") : "-";
  var tr  = document.getElementById("rows-barang").querySelectorAll("tr")[i];
  if (tr) { var cells=tr.querySelectorAll("td"); cells[cells.length-3].textContent=sub; }
  calcTot("barang");
}

function oih(type, i, val) {
  var rows  = type === "barang" ? rBarang : rJasa;
  var clean = parseInt((val||"").replace(/\D/g,"")) || 0;
  if (type === "jasa") {
    rows[i].nominal = clean;
    var tr = document.getElementById("rows-jasa").querySelectorAll("tr")[i];
    if (tr) {
      var nomInput = tr.querySelectorAll(".ti.tr");
      if (nomInput[0]) nomInput[0].value = clean ? clean.toLocaleString("id-ID") : "";
    }
    calcTot("jasa");
  } else {
    rows[i].harga = clean;
    rows[i].total = (parseInt(rows[i].jumlah)||0) * clean;
    var sub = rows[i].total ? "Rp "+rows[i].total.toLocaleString("id-ID") : "-";
    var tr  = document.getElementById("rows-barang").querySelectorAll("tr")[i];
    if (tr) {
      var cells = tr.querySelectorAll("td");
      cells[cells.length-3].textContent = sub;
      var hInput = tr.querySelectorAll(".ti.tr");
      if (hInput[0]) hInput[0].value = clean ? clean.toLocaleString("id-ID") : "";
    }
    calcTot("barang");
  }
}

function calcTot(type) {
  var total = 0;
  if (type === "barang") {
    total = rBarang.reduce(function(s,r){ return s+(r.total||0); }, 0);
  } else {
    total = rJasa.reduce(function(s,r){ return s+(r.nominal||0); }, 0);
  }
  document.getElementById("tot-"+type).textContent = "Rp "+total.toLocaleString("id-ID");
  var grand = rBarang.reduce(function(s,r){return s+(r.total||0);},0) +
              rJasa.reduce(function(s,r){return s+(r.nominal||0);},0);
  if (grand > 0) document.getElementById("f-hg").value = grand.toLocaleString("id-ID");
}

function getItemsJSON() {
  var jp = document.querySelector('input[name="jp"]:checked');
  var v  = jp ? jp.value : "";
  if (v === "Barang") return JSON.stringify(rBarang);
  if (v === "Jasa")   return JSON.stringify(rJasa);
  return "";
}

function buildItemDetail(jsonStr, type, realStr) {
  if (!jsonStr) return "-";
  try {
    var rows = JSON.parse(jsonStr);
    if (!rows.length) return "-";
    var realData = [];
    if (realStr) { try { realData = JSON.parse(realStr); } catch(e){} }
    var hasReal = realData.length > 0;
    if (type === "Barang") {
      return '<table style="width:100%;border-collapse:collapse;font-size:12px;margin-top:6px">' +
        '<tr style="background:var(--navy);color:white">' +
        '<th style="padding:5px 8px;text-align:left">Nama Barang</th>' +
        '<th style="padding:5px 8px">Kategori</th><th style="padding:5px 8px">Subkategori</th>' +
        '<th style="padding:5px 8px">Satuan</th>' +
        '<th style="padding:5px 8px;text-align:right">Harga Est.</th>' +
        '<th style="padding:5px 8px;text-align:right">Jumlah</th>' +
        '<th style="padding:5px 8px;text-align:right">Est. Total</th>' +
        (hasReal ? '<th style="padding:5px 8px;text-align:right;background:#107856">Harga Real</th><th style="padding:5px 8px;text-align:right;background:#107856">Total Real</th>' : '') +
        '<th style="padding:5px 8px">Keterangan</th></tr>' +
        rows.map(function(r,i) {
          var bg = i%2===0?"white":"var(--g50)";
          var hEst = r.hargaEst ? "Rp "+parseInt(r.hargaEst).toLocaleString("id-ID") : "-";
          var tot  = r.total ? "Rp "+parseInt(r.total).toLocaleString("id-ID") : "-";
          var rr   = realData[i] || {};
          var hR   = rr.hargaReal ? "Rp "+parseInt(rr.hargaReal).toLocaleString("id-ID") : "-";
          var tR   = (rr.hargaReal&&r.jumlah) ? "Rp "+(parseInt(rr.hargaReal)*parseInt(r.jumlah)).toLocaleString("id-ID") : "-";
          return '<tr style="background:'+bg+'">'+
            '<td style="padding:5px 8px">'+(r.nama||"-")+'</td>'+
            '<td style="padding:5px 8px;text-align:center">'+(r.kategori||"-")+'</td>'+
            '<td style="padding:5px 8px;text-align:center">'+(r.subkat||"-")+'</td>'+
            '<td style="padding:5px 8px;text-align:center">'+(r.satuan||"-")+'</td>'+
            '<td style="padding:5px 8px;text-align:right">'+hEst+'</td>'+
            '<td style="padding:5px 8px;text-align:right">'+(r.jumlah||0)+'</td>'+
            '<td style="padding:5px 8px;text-align:right;font-weight:600;color:var(--navy)">'+tot+'</td>'+
            (hasReal?'<td style="padding:5px 8px;text-align:right;color:var(--teal);font-weight:600">'+hR+'</td><td style="padding:5px 8px;text-align:right;color:var(--teal);font-weight:700">'+tR+'</td>':'')+
            '<td style="padding:5px 8px">'+(r.keterangan||"-")+'</td></tr>';
        }).join("")+'</table>';
    } else {
      return '<table style="width:100%;border-collapse:collapse;font-size:12px;margin-top:6px">' +
        '<tr style="background:var(--navy);color:white">' +
        '<th style="padding:5px 8px;text-align:left">Nama Vendor</th>' +
        '<th style="padding:5px 8px">Kategori</th><th style="padding:5px 8px">Domisili</th>' +
        '<th style="padding:5px 8px">PIC Vendor</th><th style="padding:5px 8px">No. Telpon</th>' +
        '<th style="padding:5px 8px;text-align:right">Nominal Est.</th>' +
        (hasReal ? '<th style="padding:5px 8px;text-align:right;background:#107856">Nominal Real</th>' : '') +
        '<th style="padding:5px 8px">Keterangan</th></tr>' +
        rows.map(function(r,i) {
          var bg = i%2===0?"white":"var(--g50)";
          var nom = r.nominal ? "Rp "+parseInt(r.nominal).toLocaleString("id-ID") : "-";
          var rr  = realData[i] || {};
          var nR  = rr.hargaReal ? "Rp "+parseInt(rr.hargaReal).toLocaleString("id-ID") : "-";
          return '<tr style="background:'+bg+'">'+
            '<td style="padding:5px 8px">'+(r.vendor||"-")+'</td>'+
            '<td style="padding:5px 8px;text-align:center">'+(r.kategori||"-")+'</td>'+
            '<td style="padding:5px 8px;text-align:center">'+(r.domisili||"-")+'</td>'+
            '<td style="padding:5px 8px">'+(r.pic||"-")+'</td>'+
            '<td style="padding:5px 8px">'+(r.telp||"-")+'</td>'+
            '<td style="padding:5px 8px;text-align:right;font-weight:600;color:var(--navy)">'+nom+'</td>'+
            (hasReal?'<td style="padding:5px 8px;text-align:right;color:var(--teal);font-weight:700">'+nR+'</td>':'')+
            '<td style="padding:5px 8px">'+(r.keterangan||"-")+'</td></tr>';
        }).join("")+'</table>';
    }
  } catch(e) { return jsonStr || "-"; }
}


function uploadFile(file, nomor) {
  return getSid().then(function(sid) {
    var safeName  = file.name.replace(/[#%*:<>?\/\\|]/g, "_");
    var safeNomor = nomor.replace(/\//g, "-");
    var path      = "Pengadaan Attachments/" + safeNomor + "/" + safeName;
    var url       = "https://graph.microsoft.com/v1.0/sites/" + sid +
                    "/drive/root:/" + encodeURIComponent(path).replace(/%2F/g, "/") + ":/content";
    return getToken().then(function(t) {
      return fetch(url, {
        method: "PUT",
        headers: {
          Authorization: "Bearer " + t,
          "Content-Type": file.type || "application/octet-stream"
        },
        body: file
      });
    }).then(function(r) {
      if (!r.ok) return r.text().then(function(e) {
        throw new Error("Upload gagal (" + file.name + "): " + e);
      });
      return r.json();
    }).then(function(data) {
      return { name: file.name, url: data.webUrl, size: file.size };
    });
  });
}

function uploadAllFiles(files, nomor) {
  if (!files || files.length === 0) return Promise.resolve([]);
  return Promise.all(files.map(function(f) { return uploadFile(f, nomor); }));
}

var realRows = [];

function onStatusChange() {
  var st = document.getElementById("mst").value;
  document.getElementById("delivered-fields").style.display   = st === "Delivered"            ? "block" : "none";
  document.getElementById("real-price-section").style.display = st === "Submitted to Finance" ? "block" : "none";
  if (st !== "Delivered")            { document.getElementById("m-penerima").value = ""; document.getElementById("m-tgl-terima").value = ""; }
  if (st !== "Submitted to Finance") { realRows = []; }
  if (st === "Submitted to Finance") { populateRealPriceTable(); }
}

function populateRealPriceTable() {
  var item = allSubs.filter(function(i){ return i.id===curId; })[0];
  if (!item) return;
  var type = item.JenisProduk || "Barang";
  var rows = [];
  if (item.DetailItem) {
    try { rows = JSON.parse(item.DetailItem); } catch(e) { rows = []; }
  }
  // Pre-fill with existing real prices if any
  var existing = {};
  if (item.DetailItemReal) {
    try {
      var ex = JSON.parse(item.DetailItemReal);
      ex.forEach(function(r,i){ existing[i] = r.hargaReal || 0; });
    } catch(e){}
  }
  realRows = rows.map(function(r,i){ return Object.assign({},r,{hargaReal: existing[i]||0}); });
  renderRealTable(type);
}

function renderRealTable(type) {
  var thead = document.getElementById("real-thead");
  var tbody = document.getElementById("real-tbody");
  if (!realRows.length) {
    thead.innerHTML = ""; tbody.innerHTML = '<tr><td style="padding:12px;color:var(--g500);text-align:center">Tidak ada detail item</td></tr>'; return;
  }
  if (type === "Barang") {
    thead.innerHTML = '<tr><th style="width:25%">Nama Barang</th><th style="width:8%">Jml</th><th style="width:8%">Satuan</th><th style="width:18%">Harga Est.</th><th style="width:20%">Harga Real *</th><th style="width:18%">Total Real</th></tr>';
    tbody.innerHTML = realRows.map(function(r,i){
      var est = r.hargaEst ? "Rp "+parseInt(r.hargaEst).toLocaleString("id-ID") : "-";
      var tot = (r.hargaReal && r.jumlah) ? "Rp "+(parseInt(r.hargaReal)*parseInt(r.jumlah)).toLocaleString("id-ID") : "-";
      return '<tr>'+
        '<td style="padding:4px 6px;font-size:12px">'+(r.nama||"-")+'</td>'+
        '<td style="padding:4px 6px;text-align:center;font-size:12px">'+(r.jumlah||0)+'</td>'+
        '<td style="padding:4px 6px;text-align:center;font-size:12px">'+(r.satuan||"-")+'</td>'+
        '<td style="padding:4px 6px;text-align:right;font-size:12px;color:var(--g500)">'+est+'</td>'+
        '<td style="padding:4px 6px"><input class="ti tr" type="text" value="'+(r.hargaReal?parseInt(r.hargaReal).toLocaleString("id-ID"):"")+'" placeholder="0" onblur="onRealHarga('+i+',this.value)"></td>'+
        '<td style="padding:4px 6px;text-align:right;font-weight:600;color:var(--teal);font-size:12px" id="rr-tot-'+i+'">'+tot+'</td>'+
        '</tr>';
    }).join("");
  } else {
    thead.innerHTML = '<tr><th style="width:30%">Nama Vendor</th><th style="width:20%">Nominal Est.</th><th style="width:25%">Nominal Real *</th><th style="width:22%">Total Real</th></tr>';
    tbody.innerHTML = realRows.map(function(r,i){
      var est = r.nominal ? "Rp "+parseInt(r.nominal).toLocaleString("id-ID") : "-";
      var tot = r.hargaReal ? "Rp "+parseInt(r.hargaReal).toLocaleString("id-ID") : "-";
      return '<tr>'+
        '<td style="padding:4px 6px;font-size:12px">'+(r.vendor||r.nama||"-")+'</td>'+
        '<td style="padding:4px 6px;text-align:right;font-size:12px;color:var(--g500)">'+est+'</td>'+
        '<td style="padding:4px 6px"><input class="ti tr" type="text" value="'+(r.hargaReal?parseInt(r.hargaReal).toLocaleString("id-ID"):"")+'" placeholder="0" onblur="onRealHarga('+i+',this.value)"></td>'+
        '<td style="padding:4px 6px;text-align:right;font-weight:600;color:var(--teal);font-size:12px" id="rr-tot-'+i+'">'+tot+'</td>'+
        '</tr>';
    }).join("");
  }
  calcRealTotal();
}

function onRealHarga(i, val) {
  var clean = parseInt((val||"").replace(/\D/g,"")) || 0;
  realRows[i].hargaReal = clean;
  var item = allSubs.filter(function(x){ return x.id===curId; })[0];
  var type = item ? item.JenisProduk : "Barang";
  var tot = type === "Barang"
    ? (clean * (parseInt(realRows[i].jumlah)||0))
    : clean;
  var totEl = document.getElementById("rr-tot-"+i);
  if (totEl) totEl.textContent = tot ? "Rp "+tot.toLocaleString("id-ID") : "-";
  // Update input display
  var inputs = document.getElementById("real-tbody").querySelectorAll("tr")[i];
  if (inputs) {
    var inp = inputs.querySelectorAll(".ti.tr")[0];
    if (inp) inp.value = clean ? clean.toLocaleString("id-ID") : "";
  }
  calcRealTotal();
}

function calcRealTotal() {
  var item = allSubs.filter(function(x){ return x.id===curId; })[0];
  var type = item ? item.JenisProduk : "Barang";
  var total = realRows.reduce(function(s,r){
    return s + (type==="Barang" ? (r.hargaReal||0)*(parseInt(r.jumlah)||0) : (r.hargaReal||0));
  },0);
  document.getElementById("real-total").textContent = "Rp "+total.toLocaleString("id-ID");
  return total;
}

function getRealTotal() { return calcRealTotal(); }



// =============================================
// EMAIL NOTIFICATIONS
// =============================================
function deepLink(id) {
  return (window.location.origin + window.location.pathname) + (id ? "?item=" + id : "");
}

function emailDocLinks(dokStr) {
  if (!dokStr) return "";
  try {
    var files = JSON.parse(dokStr);
    if (!files.length) return "";
    return "<div style='margin:10px 0 0;padding:10px;background:#f8f9fa;border-radius:6px'>" +
      "<p style='margin:0 0 6px;font-size:12px;color:#666;font-weight:bold'>&#128206; Dokumen Pendukung:</p>" +
      files.map(function(f){ return "<a href='" + f.url + "' style='display:block;color:#0B4F7E;font-size:13px;margin:2px 0'>&#8594; " + f.name + "</a>"; }).join("") +
      "</div>";
  } catch(e) { return ""; }
}

function emailDetail(item, data) {
  var hg = "Rp " + (parseInt(item.EstimasiHarga || data.hg || 0)).toLocaleString("id-ID");
  var tgl = item.TanggalPengajuan
    ? item.TanggalPengajuan.split("T")[0].split("-").reverse().join("/")
    : (data.tgl || new Date().toLocaleDateString("id-ID"));
  var rows = [
    ["Company",        item.Company        || data.co || "-"],
    ["Client",         item.Client         || data.cl || "-"],
    ["Project",        item.Project        || data.pr || "-"],
    ["Jenis",          (item.JenisProduk   || data.jp || "-") + " / " + (item.JenisPengadaan || "")],
    ["Tujuan",         item.TujuanPermintaan || data.tj || "-"],
    ["Estimasi Harga", "<strong style='color:#0B4F7E'>" + hg + "</strong>"],
    ["Diajukan Oleh",  item.SubmittedBy    || data.submitter || "-"],
    ["Tanggal",        tgl]
  ];
  return "<table style='width:100%;border-collapse:collapse;font-size:13px;font-family:Arial;margin-top:10px'>" +
    rows.map(function(r, i){
      return "<tr style='background:" + (i%2===0?"white":"#f8f9fa") + "'>" +
        "<td style='padding:7px 14px;color:#888;width:35%'>" + r[0] + "</td>" +
        "<td style='padding:7px 14px;color:#333'>" + r[1] + "</td></tr>";
    }).join("") + "</table>";
}

function actionBtn(url, label, color) {
  return "<div style='margin:20px 0 6px;text-align:center'>" +
    "<a href='" + url + "' target='_blank' rel='noopener noreferrer' style='display:inline-block;padding:13px 32px;background:" + color + ";color:white;border-radius:8px;text-decoration:none;font-family:Arial;font-size:15px;font-weight:bold;letter-spacing:.3px'>" +
    label + " &rarr;</a></div>" +
    "<p style='text-align:center;font-size:11px;color:#bbb;margin:4px 0 0;font-family:Arial'>atau buka: <a href='" + url + "' target='_blank' style='color:#bbb'>" + url + "</a></p>";
}

function emailWrap(accentColor, icon, title, sub, bodyHtml) {
  return "<div style='max-width:600px;margin:0 auto;font-family:Arial'>" +
    "<div style='background:" + accentColor + ";padding:18px 24px;border-radius:10px 10px 0 0'>" +
    "<h2 style='color:white;margin:0;font-size:17px'>&#127970; GOS Group &mdash; Procurement Portal</h2>" +
    "<p style='color:rgba(255,255,255,.75);margin:3px 0 0;font-size:12px'>Notifikasi Pengajuan Pengadaan</p></div>" +
    "<div style='background:white;border:1px solid #e2e6ea;border-top:none;border-radius:0 0 10px 10px;padding:24px'>" +
    "<h3 style='color:" + accentColor + ";margin:0 0 6px;font-size:16px'>" + icon + " " + title + "</h3>" +
    "<p style='color:#666;font-size:13px;margin:0 0 12px'>" + sub + "</p>" +
    bodyHtml +
    "<hr style='border:none;border-top:1px solid #eee;margin:20px 0'>" +
    "<p style='font-size:11px;color:#bbb;margin:0'>Pesan ini dikirim otomatis dari GOS Procurement Portal. Jangan balas email ini.</p>" +
    "</div></div>";
}

function sendNotif(type, rn, data) {
  var to   = data.to || CONFIG.approverEmail;
  var item = data.item || {};
  var id   = data.itemId || item.id || "";
  var url  = deepLink(id);
  var docs = emailDocLinks(data.dokStr || item.DokumenPendukung || "");
  var detail = emailDetail(item, data);
  var subject = "", body = "";

  if (type === "new") {
    subject = "[Approval L1] " + rn + " — Menunggu Persetujuan Anda";
    body = emailWrap("#D68910", "&#9888;", "Pengajuan Baru Perlu Approval L1 Anda",
      "Pengajuan berikut telah dikirim oleh <strong>" + (data.submitter || "pengaju") + "</strong> dan menunggu persetujuan Anda.",
      "<p style='background:#FEF9E7;border-left:4px solid #D68910;padding:10px 14px;border-radius:4px;font-size:13px;color:#666;margin:0 0 10px'>" +
      "Nomor: <strong style='color:#0B4F7E;font-family:monospace'>" + rn + "</strong></p>" +
      detail + docs + actionBtn(url, "Buka &amp; Berikan Approval", "#D68910"));

    var me = currentUser ? currentUser.userPrincipalName : null;
    if (me) {
      var bs = emailWrap("#1A8C6E","&#10003;","Pengajuan Berhasil Dikirim",
        "Pengajuan Anda sedang menunggu approval dari <strong>" + (data.l1Name||"atasan") + "</strong>.",
        detail + docs);
      gPost("https://graph.microsoft.com/v1.0/me/sendMail",{
        message:{subject:"[Terkirim] "+rn,body:{contentType:"HTML",content:bs},
        toRecipients:[{emailAddress:{address:me}}]},saveToSentItems:false}).catch(function(){});
    }

  } else if (type === "l2-needed") {
    subject = "[Approval L2] " + rn + " — Nilai Melebihi Threshold";
    body = emailWrap("#7B4FBF","&#128274;","Pengajuan Memerlukan Approval L2 Anda",
      "Pengajuan telah disetujui oleh <strong>" + (data.l1Name || "Approver L1") + "</strong> dan memerlukan approval final karena nilai melebihi Rp " +
      APPROVAL_CONFIG.l2Threshold.toLocaleString("id-ID") + ".",
      detail + docs + actionBtn(url, "Buka &amp; Berikan Approval L2", "#7B4FBF"));

  } else if (type === "ready-ga") {
    subject = "[Siap Diproses] " + rn + " — Semua Approval Selesai";
    body = emailWrap("#1A8C6E","&#10003;","Pengajuan Siap Diproses GA",
      "Semua approval telah diberikan. Silakan proses pengajuan berikut.",
      detail + docs + actionBtn(url, "Buka &amp; Proses di Portal", "#1A8C6E"));

  } else if (type === "rejected") {
    subject = "[Ditolak] " + rn;
    body = emailWrap("#C0392B","&#10007;","Pengajuan Anda Ditolak",
      "Pengajuan ditolak oleh Approver " + (data.level||"").toUpperCase() + ".",
      (data.notes ? "<div style='background:#fdedec;border-left:4px solid #C0392B;padding:10px 14px;border-radius:4px;margin-bottom:10px'><strong style='color:#C0392B'>Alasan:</strong> <span style='color:#666'>" + data.notes + "</span></div>" : "") +
      detail + docs + actionBtn(url, "Lihat Detail di Portal", "#C0392B"));
  }

  gPost("https://graph.microsoft.com/v1.0/me/sendMail", {
    message: { subject: subject, body: { contentType: "HTML", content: body },
               toRecipients: [{ emailAddress: { address: to } }] },
    saveToSentItems: false
  }).catch(function(e){ console.warn("Email gagal:", e.message); });
}


// =============================================
// DASHBOARD
// =============================================
var allSubs = [], filterMyApv = false;
function toggleMyApv(){
  filterMyApv=!filterMyApv;
  var b=document.getElementById("btn-my-apv");
  b.style.background=filterMyApv?"var(--amber-bd)":"";
  b.style.color=filterMyApv?"white":"";
  renderTbl();
}
var curId   = null;

function loadDash() {
  document.getElementById("tbl").innerHTML = '<tr class="lr"><td colspan="8">Memuat data...</td></tr>';
  getItems(CONFIG.resultList).then(function(items) {
    allSubs = items.map(function(i) {
      var f = i.fields;
      f.id = i.id;
      return f;
    });
    updateStats();
    populateCoFilter();
    renderTbl();
  }).catch(function(e) {
    document.getElementById("tbl").innerHTML = '<tr class="lr"><td colspan="8">Gagal: ' + e.message + '</td></tr>';
  });
}

function updateStats() {
  document.getElementById("st").textContent  = allSubs.length;
  document.getElementById("sp").textContent  = allSubs.filter(function(i){ return i.Status==="Pending L1"; }).length;
  document.getElementById("sp2").textContent = allSubs.filter(function(i){ return i.Status==="Pending L2"; }).length;
  document.getElementById("sa").textContent  = allSubs.filter(function(i){ return ["Approved","Submitted to Finance","Delivered"].indexOf(i.Status)>-1; }).length;
  document.getElementById("sr").textContent  = allSubs.filter(function(i){ return i.Status==="Rejected"; }).length;
}

function populateCoFilter() {
  var sel = document.getElementById("fco");
  var prev = sel.value;
  sel.innerHTML = '<option value="">Semua Company</option>';
  var seen = {};
  allSubs.forEach(function(i) {
    if (i.Company && !seen[i.Company]) {
      seen[i.Company] = true;
      var o = document.createElement("option"); o.value = i.Company; o.textContent = i.Company; sel.appendChild(o);
    }
  });
  sel.value = prev;
}

function renderTbl() {
  var q  = document.getElementById("srch").value.toLowerCase();
  var st = document.getElementById("fst").value;
  var co = document.getElementById("fco").value;
  var fr = document.getElementById("ffr").value;
  var to = document.getElementById("fto").value;

  var f = allSubs.filter(function(i) {
    if (q && ([i.NomorPengajuan,i.Company,i.Client,i.Project,i.TujuanPermintaan].join(" ").toLowerCase().indexOf(q) < 0)) return false;
    if (st && i.Status !== st) return false;
    if (co && i.Company !== co) return false;
    if (fr && (i.TanggalPengajuan || "") < fr) return false;
    if (to && (i.TanggalPengajuan || "") > to) return false;
    if (filterMyApv) {
      if (!(canApproveL1(i) || canApproveL2(i))) return false;
    }
    return true;
  }).sort(function(a,b) { return (b.TanggalPengajuan||"").localeCompare(a.TanggalPengajuan||""); });

  var tb = document.getElementById("tbl");
  if (!f.length) {
    tb.innerHTML = '<tr><td colspan="10" style="text-align:center;padding:32px;color:var(--g500)">Tidak ada data</td></tr>';
    return;
  }
  tb.innerHTML = f.map(function(item, idx) {
    var h    = item.EstimasiHarga ? "Rp " + parseInt(item.EstimasiHarga).toLocaleString("id-ID") : "-";
    var hr   = item.HargaReal     ? '<span style="color:var(--teal);font-weight:600">Rp ' + parseInt(item.HargaReal).toLocaleString("id-ID") + '</span>' : '<span style="color:var(--g400)">-</span>';
    var tgl  = item.TanggalPengajuan ? item.TanggalPengajuan.split("T")[0].split("-").reverse().join("/") : "-";
    var jBg  = item.JenisPengadaan === "EXT" ? "#E8F2FB" : "#E6F5F0";
    var jCo  = item.JenisPengadaan === "EXT" ? "var(--navy)" : "var(--teal)";
    var cl   = (item.Client || "-");
    var cls  = cl.length > 22 ? cl.slice(0, 22) + "..." : cl;
    var pr   = (item.Project || "-");
    var prs  = pr.length > 25 ? pr.slice(0, 25) + "..." : pr;
    return '<tr onclick="openMo(\'' + item.id + '\')">' +
      '<td style="color:var(--g500)">' + (idx+1) + '</td>' +
      '<td style="font-weight:600;color:var(--navy);font-family:\'Courier New\',monospace;font-size:12px">' + (item.NomorPengajuan||"-") + '</td>' +
      '<td>' + tgl + '</td>' +
      '<td style="font-size:12px">' + (item.Company||"-") + '</td>' +
      '<td style="font-size:12px">' + cls + '</td>' +
      '<td style="font-size:12px">' + prs + '</td>' +
      '<td><span style="background:' + jBg + ';color:' + jCo + ';padding:2px 7px;border-radius:10px;font-size:11px;font-weight:600">' + (item.JenisPengadaan||"-") + '</span></td>' +
      '<td style="text-align:right;font-weight:500">' + h + '</td>' +
      '<td style="text-align:right">' + hr + '</td>' +
      '<td>' + sBadge(item.Status) + '</td></tr>';
  }).join("");
}

function sBadge(s) {
  var m = {
    "Pending L1":          "pl1",
    "Pending L2":          "pl2",
    "Approved":            "apv",
    "Submitted to Finance":"app",
    "Delivered":           "dlv",
    "Rejected":            "rej"
  };
  return '<span class="sb '+(m[s]||"pl1")+'"><span class="sd"></span>'+(s||"Pending L1")+'</span>';
}

// =============================================
// MODAL
// =============================================
function buildFileLinks(dokStr) {
  if (!dokStr) return "-";
  try {
    var files = JSON.parse(dokStr);
    if (!files.length) return "-";
    return files.map(function(f) {
      var size = f.size > 1048576 ? (f.size/1048576).toFixed(1)+" MB" : Math.round((f.size||0)/1024)+" KB";
      return '<a href="' + f.url + '" target="_blank" ' +
             'style="display:inline-flex;align-items:center;gap:5px;background:#E8F2FB;color:#0B4F7E;' +
             'padding:4px 10px;border-radius:4px;font-size:12px;font-weight:500;text-decoration:none;' +
             'margin:2px 4px 2px 0;border:1px solid #BAD4EE">' +
             '&#128196; ' + f.name + ' <span style="opacity:.6;font-size:10px">(' + size + ')</span>' +
             '</a>';
    }).join("");
  } catch(e) {
    return dokStr || "-";
  }
}

function openMo(id) {
  curId = id;
  var item = allSubs.filter(function(i) { return i.id === id; })[0];
  if (!item) return;
  document.getElementById("mtt").textContent = item.NomorPengajuan || "Detail Pengajuan";
  var h   = item.EstimasiHarga ? "Rp " + parseInt(item.EstimasiHarga).toLocaleString("id-ID") : "-";
  var tgl = item.TanggalPengajuan ? item.TanggalPengajuan.split("T")[0].split("-").reverse().join("/") : "-";
  document.getElementById("mcnt").innerHTML =
    '<div class="df"><div class="dl">Nomor Pengajuan</div><div class="dv" style="font-family:\'Courier New\',monospace;color:var(--navy)">' + (item.NomorPengajuan||"-") + '</div></div>' +
    '<div class="df"><div class="dl">Tanggal</div><div class="dv">' + tgl + '</div></div>' +
    '<div class="df"><div class="dl">Jenis Pengadaan</div><div class="dv">' + (item.JenisPengadaan||"-") + '</div></div>' +
    '<div class="df"><div class="dl">Jenis Produk</div><div class="dv">' + (item.JenisProduk||"-") + '</div></div>' +
    '<div class="df"><div class="dl">Company</div><div class="dv">' + (item.Company||"-") + '</div></div>' +
    '<div class="df"><div class="dl">Client</div><div class="dv">' + (item.Client||"-") + '</div></div>' +
    '<div class="df full"><div class="dl">Project</div><div class="dv">' + (item.Project||"-") + '</div></div>' +
    (item.Cabang ? '<div class="df"><div class="dl">Cabang</div><div class="dv"><span style="background:#E8F2FB;color:var(--navy);padding:2px 8px;border-radius:10px;font-size:12px">' + item.Cabang + '</span></div></div>' : '') +
    '<div class="df full"><div class="dl">Detail Item</div><div class="dv">' + buildItemDetail(item.DetailItem, item.JenisProduk, item.DetailItemReal) + '</div></div>' +
    '<div class="df full"><div class="dl">Tujuan Permintaan</div><div class="dv">' + (item.TujuanPermintaan||"-") + '</div></div>' +
    '<div class="df"><div class="dl">Estimasi Harga</div><div class="dv" style="font-weight:700;color:var(--navy)">' + h + '</div></div>' +
    '<div class="df"><div class="dl">Diajukan Oleh</div><div class="dv">' + (item.SubmittedBy||"-") + '</div></div>' +
    '<div class="df full"><div class="dl">Dokumen Pendukung</div><div class="dv">' + buildFileLinks(item.DokumenPendukung) + '</div></div>' +
    '<div class="df"><div class="dl">Status</div><div class="dv">' + sBadge(item.Status) + '</div></div>' +
    '<div class="df"><div class="dl">Estimasi Harga</div><div class="dv" style="font-weight:600;color:var(--navy)">Rp ' + (parseInt(item.EstimasiHarga)||0).toLocaleString("id-ID") + '</div></div>' +
    (item.HargaReal ? '<div class="df"><div class="dl">Harga Real</div><div class="dv" style="font-weight:700;color:var(--teal);font-size:15px">Rp ' + (parseInt(item.HargaReal)||0).toLocaleString("id-ID") + '</div></div>' : '') +
    (item.TanggalSubmittedToFinance ? '<div class="df"><div class="dl">Tgl. Submitted to Finance</div><div class="dv">' + item.TanggalSubmittedToFinance.split("T")[0].split("-").reverse().join("/") + '</div></div>' : '') +
    '<div class="df"><div class="dl">Catatan GA</div><div class="dv">' + (item.ApproverNotes||"-") + '</div></div>' +
    (item.Status === "Delivered" ? (
      '<div class="df"><div class="dl">Nama Penerima</div><div class="dv" style="font-weight:600;color:#7B4FBF">' + (item.NamaPenerima||"-") + '</div></div>' +
      '<div class="df"><div class="dl">Tanggal Terima</div><div class="dv" style="font-weight:600;color:#7B4FBF">' + (item.TanggalTerima ? item.TanggalTerima.split("T")[0].split("-").reverse().join("/") : "-") + '</div></div>'
    ) : '') +
    '<div class="df full"><div class="dl">Riwayat Approval</div><div class="dv">' + buildApvHistory(item) + '</div></div>';
  document.getElementById("mst").value  = item.Status || "Pending L1";
  document.getElementById("mnt").value  = item.ApproverNotes || "";
  document.getElementById("m-penerima").value   = item.NamaPenerima  || "";
  document.getElementById("m-tgl-terima").value = item.TanggalTerima ? item.TanggalTerima.split("T")[0] : "";
  onStatusChange();

  // Show/hide action sections
  document.getElementById("m-l1").style.display   = canApproveL1(item) ? "block" : "none";
  document.getElementById("m-l2").style.display   = canApproveL2(item) ? "block" : "none";
  document.getElementById("m-ga").style.display   = canUpdateGA(item)  ? "block" : "none";
  document.getElementById("m-save-btn").style.display = canUpdateGA(item) ? "inline-flex" : "none";
  var showPdf = ["Submitted to Finance","Delivered"].indexOf(item.Status) > -1;
  document.getElementById("m-pdf-btn").style.display = showPdf ? "inline-flex" : "none";

  // Reset notes
  document.getElementById("m-l1-notes").value = "";
  document.getElementById("m-l2-notes").value = "";

  document.getElementById("modal").classList.add("on");
}

function closeMo() {
  document.getElementById("modal").classList.remove("on");
  curId = null;
}

function buildApvHistory(item) {
  var html = '<div class="apv-hist">';
  var hasHistory = false;

  if (item.L1ApproverName || item.L1ApproverEmail) {
    hasHistory = true;
    var l1Status = item.Status === "Pending L1" ? "Menunggu" :
                   item.Status === "Rejected" && !item.L1ApprovalDate ? "Menunggu" :
                   item.L1ApprovalDate ? "Approved" : "Menunggu";
    if (item.Status === "Rejected" && !item.L2ApprovalDate && !item.L2Notes) l1Status = "Rejected";
    html += '<div class="apv-hist-row"><span class="apv-key">L1 :</span><span>' +
            (item.L1ApproverName||item.L1ApproverEmail||"-") + ' — ' + l1Status +
            (item.L1ApprovalDate ? " (" + item.L1ApprovalDate.split("T")[0].split("-").reverse().join("/") + ")" : "") +
            (item.L1Notes ? "<br><em style='color:var(--g500)'>" + item.L1Notes + "</em>" : "") + '</span></div>';
  }
  if (item.L2Notes || ["Pending L2","Approved","Submitted to Finance","Delivered"].indexOf(item.Status)>-1) {
    hasHistory = true;
    var l2s = item.Status === "Pending L2" ? "Menunggu" : item.L2ApprovalDate ? "Approved" : "-";
    if (item.Status === "Rejected" && item.L2Notes) l2s = "Rejected";
    html += '<div class="apv-hist-row"><span class="apv-key">L2 :</span><span>' +
            APPROVAL_CONFIG.l2.name + ' — ' + l2s +
            (item.L2ApprovalDate ? " (" + item.L2ApprovalDate.split("T")[0].split("-").reverse().join("/") + ")" : "") +
            (item.L2Notes ? "<br><em style='color:var(--g500)'>" + item.L2Notes + "</em>" : "") + '</span></div>';
  }
  if (!hasHistory) html += '<span style="color:var(--g500);font-size:11px">Belum ada riwayat approval</span>';
  html += '</div>';
  return html;
}

function doApprove(level) {
  if (!curId) return;
  var item    = allSubs.filter(function(i){ return i.id===curId; })[0];
  if (!item) return;
  var notes   = document.getElementById("m-"+level+"-notes").value;
  var today   = new Date().toISOString().split("T")[0];
  var fields  = {};
  var submitterEmail = "";

  if (level === "l1") {
    var req2 = needsL2(item.EstimasiHarga, { l2Email: item.L2ApproverEmail });
    fields.Status         = req2 ? "Pending L2" : "Approved";
    fields.L1Notes        = notes;
    fields.L1ApprovalDate = today;
    if (req2 && item.L2ApproverEmail) {
      sendNotif("l2-needed", item.NomorPengajuan, { to: item.L2ApproverEmail, item: item, l1Name: item.L1ApproverName || "" });
    } else {
      APPROVERS.forEach(function(ga) {
        sendNotif("ready-ga", item.NomorPengajuan, { to: ga, item: item });
      });
    }
  } else {
    fields.Status         = "Approved";
    fields.L2Notes        = notes;
    fields.L2ApprovalDate = today;
    APPROVERS.forEach(function(ga) {
      sendNotif("ready-ga", item.NomorPengajuan, { to: ga, item: item });
    });
  }

  patchItem(CONFIG.resultList, curId, fields).then(function() {
    Object.assign(item, fields);
    updateStats(); renderTbl(); closeMo();
    showToast("Pengajuan berhasil di-Approve!", "ok");
  }).catch(function(e){ showToast("Gagal: "+e.message,"er"); });
}

function doReject(level) {
  if (!curId) return;
  var item  = allSubs.filter(function(i){ return i.id===curId; })[0];
  if (!item) return;
  var notes = document.getElementById("m-"+level+"-notes").value;
  var today = new Date().toISOString().split("T")[0];
  var fields = { Status: "Rejected" };
  if (level === "l1") { fields.L1Notes = notes; fields.L1ApprovalDate = today; }
  else                { fields.L2Notes = notes; fields.L2ApprovalDate = today; }

  // Notify submitter
  sendNotif("rejected", item.NomorPengajuan, {
    to:    item.SubmittedByEmail || item.SubmittedBy || CONFIG.approverEmail,
    item:  item, notes: notes, level: level
  });

  patchItem(CONFIG.resultList, curId, fields).then(function() {
    Object.assign(item, fields);
    updateStats(); renderTbl(); closeMo();
    showToast("Pengajuan ditolak.", "er");
  }).catch(function(e){ showToast("Gagal: "+e.message,"er"); });
}

function updStatus() {
  if (!curId) return;
  if (!amGA()) { showToast("Tidak ada akses", "er"); return; }
  var item = allSubs.filter(function(i){ return i.id===curId; })[0];
  if (!item) return;
  var st       = document.getElementById("mst").value;
  var nt       = document.getElementById("mnt").value;
  var penerima = document.getElementById("m-penerima").value.trim();
  var tglTrm   = document.getElementById("m-tgl-terima").value;

  if (st === "Delivered") {
    if (!penerima){ showToast("Nama penerima wajib diisi","er"); return; }
    if (!tglTrm)  { showToast("Tanggal terima wajib diisi","er"); return; }
  }

  var fields = { Status: st, ApproverNotes: nt };

  if (st === "Submitted to Finance") {
    // Validasi harga real
    var unfilled = realRows.filter(function(r){ return !r.hargaReal; });
    if (realRows.length && unfilled.length) {
      showToast("Harga real untuk semua item wajib diisi", "er"); return;
    }
    var totalReal = getRealTotal();
    fields.HargaReal                 = totalReal;
    fields.DetailItemReal            = JSON.stringify(realRows);
    fields.TanggalSubmittedToFinance = new Date().toISOString().split("T")[0];
  }

  if (st === "Delivered") {
    fields.NamaPenerima  = penerima;
    fields.TanggalTerima = tglTrm;
  }

  patchItem(CONFIG.resultList, curId, fields).then(function() {
    Object.assign(item, fields);
    updateStats(); renderTbl(); closeMo();
    showToast("Status diperbarui: "+st, "ok");
  }).catch(function(e){ showToast("Gagal: "+e.message,"er"); });
}

// =============================================
// EXPORT EXCEL
// =============================================
function expXls() {
  if (!allSubs.length) { showToast("Tidak ada data", "er"); return; }
  var fmtTgl = function(v) {
    if (!v) return "";
    var s = typeof v === "string" ? v : String(v);
    var d = s.split("T")[0]; // "2026-05-19"
    var p = d.split("-");
    return p.length === 3 ? p[2]+"/"+p[1]+"/"+p[0] : d;
  };
  // Helper: ambil field dengan coba beberapa nama (fallback)
  var fld = function(item, names) {
    for (var n = 0; n < names.length; n++) {
      if (item[names[n]] !== undefined && item[names[n]] !== null) return item[names[n]];
    }
    return "";
  };
  var rows = allSubs.map(function(i) {
    return {
      "Nomor Pengajuan":             fld(i,["NomorPengajuan"]),
      "Tanggal Pengajuan":           fmtTgl(fld(i,["TanggalPengajuan"])),
      "Company":                     fld(i,["Company"]),
      "Client":                      fld(i,["Client"]),
      "Project":                     fld(i,["Project"]),
      "Cabang":                      fld(i,["Cabang"]),
      "Jenis Pengadaan":             fld(i,["JenisPengadaan"]),
      "Jenis Produk":                fld(i,["JenisProduk"]),
      "Tujuan":                      fld(i,["TujuanPermintaan"]),
      "Estimasi Harga":              fld(i,["EstimasiHarga"]) || 0,
      "Harga Real":                  fld(i,["HargaReal"]) || 0,
      "Status":                      fld(i,["Status"]),
      "Diajukan Oleh":               fld(i,["SubmittedBy"]),
      "L1 Approver":                 fld(i,["L1ApproverName"]),
      "Tgl. L1 Approval":            fmtTgl(fld(i,["L1ApprovalDate"])),
      "L2 Approver":                 fld(i,["L2ApproverName"]),
      "Tgl. L2 Approval":            fmtTgl(fld(i,["L2ApprovalDate"])),
      "Tgl. Submitted to Finance":   fmtTgl(fld(i,["TanggalSubmittedToFinance"])),
      "Nama Penerima":               fld(i,["NamaPenerima"]),
      "Tgl. Terima":                 fmtTgl(fld(i,["TanggalTerima","Tgl_x002e_Terima","TglTerima"])),
      "Catatan GA":                  fld(i,["ApproverNotes"])
    };
  });
  var ws = XLSX.utils.json_to_sheet(rows);
  // Auto-width kolom
  var cols = Object.keys(rows[0] || {}).map(function(k) {
    var max = k.length;
    rows.forEach(function(r) { var v = String(r[k]||""); if (v.length > max) max = v.length; });
    return { wch: Math.min(max + 2, 40) };
  });
  ws["!cols"] = cols;
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Pengajuan");
  var n = new Date();
  XLSX.writeFile(wb, "Pengajuan_GOS_" + n.getFullYear() + pad(n.getMonth()+1) + pad(n.getDate()) + ".xlsx");
  showToast("File Excel berhasil didownload", "ok");
}

// =============================================
// UI HELPERS
// =============================================
function switchView(v) {
  document.querySelectorAll(".view").forEach(function(x) { x.classList.remove("on"); });
  document.getElementById("view-" + v).classList.add("on");
  document.querySelectorAll(".nb").forEach(function(x) { x.classList.remove("on"); });
  document.getElementById("nav-" + v).classList.add("on");
  if (v === "dash")   loadDash();
  if (v === "master") { mmTab="entitas"; switchMTab("entitas"); }
}

var toastT = null;
function showToast(msg, type) {
  clearTimeout(toastT);
  var el = document.getElementById("toast");
  el.textContent = msg;
  el.className = "toast on" + (type ? " " + type : "");
  toastT = setTimeout(function() { el.classList.remove("on"); }, 3500);
}

// =============================================
// EXPORT PDF
// =============================================
function exportPDF() {
  var item = allSubs.filter(function(i){ return i.id===curId; })[0];
  if (!item) return;

  var hEst  = item.EstimasiHarga ? "Rp "+parseInt(item.EstimasiHarga).toLocaleString("id-ID") : "-";
  var hReal = item.HargaReal     ? "Rp "+parseInt(item.HargaReal).toLocaleString("id-ID")     : "-";
  var tgl   = item.TanggalPengajuan ? item.TanggalPengajuan.split("T")[0].split("-").reverse().join("/") : "-";
  var itemTable = buildItemDetail(item.DetailItem, item.JenisProduk, item.DetailItemReal);

  // Approval history
  var apvHtml = "";
  if (item.L1ApproverName) apvHtml += "<tr><td style='padding:6px 14px;color:#888;width:160px'>L1 Approver</td><td style='padding:6px 14px'>" + (item.L1ApproverName||"-") + " &mdash; " + (item.L1ApprovalDate?item.L1ApprovalDate.split("T")[0].split("-").reverse().join("/"):"-") + "</td></tr>";
  if (item.L2ApproverName) apvHtml += "<tr><td style='padding:6px 14px;color:#888'>L2 Approver</td><td style='padding:6px 14px'>" + (item.L2ApproverName||"-") + " &mdash; " + (item.L2ApprovalDate?item.L2ApprovalDate.split("T")[0].split("-").reverse().join("/"):"-") + "</td></tr>";

  var html = "<!DOCTYPE html><html><head><meta charset='utf-8'><title>Pengajuan " + (item.NomorPengajuan||"") + "</title>" +
    "<style>body{font-family:Arial,sans-serif;font-size:13px;color:#333;padding:24px;max-width:900px;margin:0 auto}" +
    "h1{font-size:18px;color:#0B4F7E;margin:0}h2{font-size:13px;color:#888;font-weight:normal;margin:4px 0 16px}" +
    ".header{background:#0B4F7E;color:white;padding:16px 20px;border-radius:8px;margin-bottom:20px}" +
    ".header h1{color:white}.header h2{color:rgba(255,255,255,.75)}" +
    ".info{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:16px}" +
    ".box{background:#f8f9fa;border-radius:6px;padding:12px 16px}" +
    ".label{font-size:11px;color:#888;text-transform:uppercase;letter-spacing:.4px;margin-bottom:3px}" +
    ".val{font-size:14px;font-weight:600;color:#333}" +
    ".val.navy{color:#0B4F7E}.val.teal{color:#1A8C6E;font-size:16px}" +
    "table{width:100%;border-collapse:collapse;margin-top:8px}" +
    "th{background:#0B4F7E;color:white;padding:7px 10px;text-align:left;font-size:12px}" +
    "td{padding:7px 10px;border-bottom:1px solid #eee;font-size:12px}" +
    "tr:last-child td{border-bottom:none}" +
    "tr:nth-child(even) td{background:#f8f9fa}" +
    ".section{margin-bottom:20px;padding-bottom:16px;border-bottom:1px solid #eee}" +
    ".section-title{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#888;margin-bottom:8px}" +
    ".stamp{margin-top:40px;display:grid;grid-template-columns:1fr 1fr 1fr;gap:20px;text-align:center}" +
    ".stamp-box{border:1px solid #ccc;border-radius:6px;padding:12px;min-height:80px}" +
    ".stamp-label{font-size:11px;color:#888;margin-bottom:4px}" +
    "@media print{body{padding:0}}" +
    "</style></head><body>" +

    "<div class='header'><h1>&#127970; GOS Group &mdash; Procurement Portal</h1>" +
    "<h2>Bukti Pengajuan Pengadaan &mdash; " + (item.NomorPengajuan||"") + "</h2></div>" +

    "<div class='section'><div class='section-title'>Informasi Pengajuan</div>" +
    "<div class='info'>" +
    "<div class='box'><div class='label'>Nomor Pengajuan</div><div class='val navy'>" + (item.NomorPengajuan||"-") + "</div></div>" +
    "<div class='box'><div class='label'>Tanggal</div><div class='val'>" + tgl + "</div></div>" +
    "<div class='box'><div class='label'>Company</div><div class='val'>" + (item.Company||"-") + "</div></div>" +
    "<div class='box'><div class='label'>Client</div><div class='val'>" + (item.Client||"-") + "</div></div>" +
    "<div class='box' style='grid-column:span 2'><div class='label'>Project</div><div class='val'>" + (item.Project||"-") + "</div></div>" +
    "<div class='box'><div class='label'>Jenis Pengadaan</div><div class='val'>" + (item.JenisPengadaan||"-") + " / " + (item.JenisProduk||"-") + "</div></div>" +
    "<div class='box'><div class='label'>Status</div><div class='val'>" + (item.Status||"-") + "</div></div>" +
    "<div class='box'><div class='label'>Diajukan Oleh</div><div class='val'>" + (item.SubmittedBy||"-") + "</div></div>" +
    "<div class='box'><div class='label'>Tujuan</div><div class='val'>" + (item.TujuanPermintaan||"-") + "</div></div>" +
    "</div></div>" +

    "<div class='section'><div class='section-title'>Ringkasan Harga</div>" +
    "<div class='info'>" +
    "<div class='box'><div class='label'>Estimasi Harga</div><div class='val navy'>" + hEst + "</div></div>" +
    (item.HargaReal ? "<div class='box'><div class='label'>Harga Real (disetujui)</div><div class='val teal'>" + hReal + "</div></div>" : "") +
    "</div></div>" +

    "<div class='section'><div class='section-title'>Detail Item</div>" + itemTable + "</div>" +

    (apvHtml ? "<div class='section'><div class='section-title'>Riwayat Approval</div><table>" + apvHtml + "</table></div>" : "") +

    "<script>window.onload=function(){window.print()}<\/script>" +
    "</body></html>";

  var w = window.open("", "_blank");
  if (w) { w.document.write(html); w.document.close(); }
  else showToast("Popup diblokir browser. Izinkan popup untuk export PDF.", "er");
}
var apvConfigId = null; // ID of L2Setting record
var apvRules    = [];   // All rule records

function loadApprovalConfig() {
  return getItems("KonfigurasiApproval").then(function(items) {
    var records = items.map(function(i){ return Object.assign({id:i.id}, i.fields); });

    var l2rec = records.filter(function(r){ return r.TipeKonfigurasi==="L2Setting"; })[0];
    if (l2rec) {
      apvConfigId = l2rec.id;
      APPROVAL_CONFIG.l2Threshold  = parseInt(l2rec.L2Threshold)||5000000;
      APPROVAL_CONFIG.l2.email     = l2rec.L2Email || APPROVAL_CONFIG.l2.email;
      APPROVAL_CONFIG.l2.name      = l2rec.L2Name  || APPROVAL_CONFIG.l2.name;
    }

    apvRules = records.filter(function(r){ return r.TipeKonfigurasi==="ApprovalRule"; });
    APPROVAL_CONFIG.rules = apvRules.map(function(r){
      return {
        project: r.Project||"",
        cabang:  r.Cabang||"",
        l1Email: r.L1Email||"",
        l1Name:  r.L1Name||"",
        l2Email: r.L2Email||"",
        l2Name:  r.L2Name||""
      };
    });
    // Load daftar pengaju yang diizinkan
    var submRecs = records.filter(function(r){ return r.TipeKonfigurasi==="SubmitterAccess"; });
    APPROVAL_CONFIG.submitters = submRecs.map(function(r){ return (r.SubmitterEmail||"").toLowerCase(); });
    submitterRecs = submRecs;
  }).catch(function(e){ console.warn("KonfigurasiApproval gagal:", e.message); });
}

// =============================================
// SETTINGS UI
// =============================================
var l2EditId  = null;
var ruleEditId = null;

function editL2() {
  var t = APPROVAL_CONFIG.l2Threshold;
  document.getElementById("l2-t-input").value = t ? t.toLocaleString("id-ID") : "";
  document.getElementById("l2-e-input").value = APPROVAL_CONFIG.l2.email;
  document.getElementById("l2-n-input").value = APPROVAL_CONFIG.l2.name;
  document.getElementById("l2-edit").style.display = "block";
}
function cancelL2() { document.getElementById("l2-edit").style.display = "none"; }

function saveL2() {
  var t = parseInt((document.getElementById("l2-t-input").value||"").replace(/\D/g,""))||0;
  var e = document.getElementById("l2-e-input").value.trim();
  var n = document.getElementById("l2-n-input").value.trim();
  if (!t||!e||!n) { showToast("Semua field wajib diisi","er"); return; }

  var fields = { Title:"L2Setting", TipeKonfigurasi:"L2Setting", L2Threshold:t, L2Email:e, L2Name:n };
  var promise = apvConfigId ? patchItem("KonfigurasiApproval", apvConfigId, fields)
                            : createItem("KonfigurasiApproval", fields);
  promise.then(function(res){
    if (!apvConfigId && res && res.id) apvConfigId = res.id;
    APPROVAL_CONFIG.l2Threshold = t;
    APPROVAL_CONFIG.l2.email    = e;
    APPROVAL_CONFIG.l2.name     = n;
    refreshL2Display();
    cancelL2();
    showToast("Konfigurasi L2 disimpan!","ok");
  }).catch(function(e){ showToast("Gagal: "+e.message,"er"); });
}

function refreshL2Display() {
  document.getElementById("l2-threshold-val").value = APPROVAL_CONFIG.l2Threshold.toLocaleString("id-ID");
  document.getElementById("l2-email-val").value      = APPROVAL_CONFIG.l2.email;
  document.getElementById("l2-name-val").value       = APPROVAL_CONFIG.l2.name;
}

function renderRules() {
  var tbody = document.getElementById("rules-tbody");
  if (!apvRules.length) {
    tbody.innerHTML = '<tr><td colspan="6" style="text-align:center;padding:20px;color:var(--g500)">Belum ada rule. Klik "+ Tambah Rule" untuk menambahkan.</td></tr>';
    return;
  }
  // Sort: specific rules first, default (*) last
  tbody.innerHTML = apvRules.map(function(r,i){
    var editBtn = '<button class="btn btn-o bsm" style="font-size:11px" onclick="openRuleModal(\'' + r.id + '\')">Edit</button>';
    var delBtn  = '<button class="btn bsm" style="font-size:11px;color:var(--red);border:1.5px solid var(--red);background:white;border-radius:var(--rsm);cursor:pointer;font-family:inherit" onclick="deleteRule(\'' + r.id + '\')">Hapus</button>';
    return '<tr>' +
      '<td style="color:var(--g500)">' + (i+1) + '</td>' +
      '<td style="font-weight:500">' + eA(r.Project || "-") + '</td>' +
      '<td>' + (r.Cabang ? '<span style="background:#E8F2FB;color:var(--navy);padding:2px 8px;border-radius:10px;font-size:11px">' + eA(r.Cabang) + '</span>' : '<span style="color:var(--g400);font-size:11px">semua cabang</span>') + '</td>' +
      '<td>' + eA(r.L1Name || "-") + '</td>' +
      '<td style="font-size:12px">' + eA(r.L1Email || "-") + '</td>' +
      '<td>' + (r.L2Name ? eA(r.L2Name) : '<span style="color:var(--g500);font-size:11px">-</span>') + '</td>' +
      '<td style="font-size:12px">' + (r.L2Email ? eA(r.L2Email) : '<span style="color:var(--g500)">-</span>') + '</td>' +
      '<td><div style="display:flex;gap:4px">' + editBtn + delBtn + '</div></td>' +
      '</tr>';
  }).join("");
}

function filterCabangDD() {
  var dd   = document.getElementById("cabang-dd");
  var q    = (document.getElementById("f-cabang").value||"").toLowerCase();
  var proj = document.getElementById("f-pr").value;

  // Ambil cabang unik dari rules untuk project yang dipilih
  var options = [];
  APPROVAL_CONFIG.rules.forEach(function(r) {
    if (r.cabang && (!proj || (r.project||"").toLowerCase().trim() === proj.toLowerCase().trim())) {
      if (options.indexOf(r.cabang) < 0) options.push(r.cabang);
    }
  });
  options.sort();

  var filtered = q ? options.filter(function(c){ return c.toLowerCase().indexOf(q)>-1; }) : options;

  if (!filtered.length) {
    dd.style.display = "none"; return;
  }

  dd.innerHTML = filtered.map(function(c) {
    return '<div onmousedown="selectCabangDD(\'' + c.replace(/'/g,"\\'") + '\')" ' +
      'style="padding:8px 11px;font-size:13px;cursor:pointer;color:var(--g900)" ' +
      'onmouseover="this.style.background=\'var(--navy-lt)\'" ' +
      'onmouseout="this.style.background=\'\'">' + c + '</div>';
  }).join("");
  dd.style.display = "block";
}

function selectCabangDD(val) {
  document.getElementById("f-cabang").value = val;
  document.getElementById("cabang-dd").style.display = "none";
}


function filterProjectDD() {
  var input = document.getElementById("rule-project");
  var dd    = document.getElementById("project-dd");
  var q     = (input.value || "").toLowerCase();

  // Ambil semua project dari masterData
  var allProjects = [];
  (masterData || []).forEach(function(r) {
    if (r.Project && allProjects.indexOf(r.Project) < 0) allProjects.push(r.Project);
  });
  allProjects.sort();

  // Tampilkan semua project (project boleh punya multiple rule untuk cabang berbeda)
  var filtered = allProjects.filter(function(p) {
    if (q) return p.toLowerCase().indexOf(q) > -1;
    return true;
  });

  if (!filtered.length) {
    dd.innerHTML = '<div style="padding:8px 11px;font-size:12px;color:var(--g500);font-style:italic">' +
      (q ? "Tidak ditemukan / sudah dikonfigurasi" : "Semua project sudah dikonfigurasi") + '</div>';
  } else {
    dd.innerHTML = filtered.map(function(p) {
      return '<div onmousedown="selectProjectDD(\'' + p.replace(/'/g, "\\'") + '\')" ' +
        'style="padding:8px 11px;font-size:13px;cursor:pointer;color:var(--g900)" ' +
        'onmouseover="this.style.background=\'var(--navy-lt)\'" ' +
        'onmouseout="this.style.background=\'\'">' + p + '</div>';
    }).join("");
  }
  dd.style.display = "block";
}

function selectProjectDD(val) {
  document.getElementById("rule-project").value = val;
  document.getElementById("project-dd").style.display = "none";
}


function openRuleModal(id) {
  ruleEditId = id;
  var r = id ? apvRules.filter(function(x){ return x.id===id; })[0] : {};
  if (!r) r = {};
  document.getElementById("rule-modal-title").textContent = id ? "Edit Rule Project" : "Tambah Rule Project";
  document.getElementById("rule-project").value  = r.Project  || "";
  document.getElementById("project-dd").style.display = "none";
  document.getElementById("rule-cabang").value   = r.Cabang   || "";
  document.getElementById("rule-l1name").value   = r.L1Name   || "";
  document.getElementById("rule-l1email").value  = r.L1Email  || "";
  document.getElementById("rule-l2name").value   = r.L2Name   || "";
  document.getElementById("rule-l2email").value  = r.L2Email  || "";
  document.getElementById("rule-modal").classList.add("on");
}
function closeRuleModal() { document.getElementById("rule-modal").classList.remove("on"); }

function saveRule() {
  var proj = document.getElementById("rule-project").value.trim();
  var cab  = document.getElementById("rule-cabang").value.trim();
  var l1n  = document.getElementById("rule-l1name").value.trim();
  var l1e  = document.getElementById("rule-l1email").value.trim();
  var l2n  = document.getElementById("rule-l2name").value.trim();
  var l2e  = document.getElementById("rule-l2email").value.trim();
  if (!proj||!l1n||!l1e) { showToast("Project, Nama L1, dan Email L1 wajib diisi","er"); return; }

  var fields = { Title: proj+(cab?"/"+cab:""), TipeKonfigurasi:"ApprovalRule", Project:proj, Cabang:cab,
                 L1Name:l1n, L1Email:l1e, L2Name:l2n, L2Email:l2e };
  var promise = ruleEditId ? patchItem("KonfigurasiApproval", ruleEditId, fields)
                           : createItem("KonfigurasiApproval", fields);
  promise.then(function(res){
    if (ruleEditId) {
      var r = apvRules.filter(function(x){ return x.id===ruleEditId; })[0];
      if (r) Object.assign(r, fields);
    } else {
      var newId = res && res.id ? res.id : "tmp-"+Date.now();
      apvRules.push(Object.assign({id:newId}, fields));
    }
    APPROVAL_CONFIG.rules = apvRules.map(function(r){
      return { project:r.Project||"", cabang:r.Cabang||"", l1Email:r.L1Email||"", l1Name:r.L1Name||"",
               l2Email:r.L2Email||"", l2Name:r.L2Name||"" };
    });
    renderRules(); closeRuleModal();
    showToast((ruleEditId?"Rule diperbarui":"Rule ditambahkan")+"!","ok");
  }).catch(function(e){ showToast("Gagal: "+e.message,"er"); });
}

function deleteRule(id) {
  if (!confirm("Hapus rule ini?")) return;
  patchItem("KonfigurasiApproval", id, { TipeKonfigurasi:"Deleted" }).then(function(){
    apvRules = apvRules.filter(function(r){ return r.id!==id; });
    APPROVAL_CONFIG.rules = apvRules.map(function(r){
      return { project:r.Project||"", cabang:r.Cabang||"", l1Email:r.L1Email||"", l1Name:r.L1Name||"",
               l2Email:r.L2Email||"", l2Name:r.L2Name||"" };
    });
    renderRules();
    showToast("Rule dihapus","ok");
  }).catch(function(e){ showToast("Gagal: "+e.message,"er"); });
}


// =============================================
// SUBMITTER ACCESS MANAGEMENT
// =============================================
var submitterRecs = [];

function renderSubmitters() {
  var tbody = document.getElementById("submitters-tbody");
  if (!tbody) return;
  if (!submitterRecs.length) {
    tbody.innerHTML = '<tr><td colspan="3" style="text-align:center;padding:20px;color:var(--g500)">Belum ada pengaju terdaftar.</td></tr>';
    return;
  }
  tbody.innerHTML = submitterRecs.map(function(r,i){
    var delBtn = '<button class="btn bsm" style="font-size:11px;color:var(--red);border:1.5px solid var(--red);background:white;border-radius:var(--rsm);cursor:pointer;font-family:inherit" onclick="deleteSubmitter(\'' + r.id + '\')">Hapus</button>';
    return '<tr><td style="color:var(--g500)">'+(i+1)+'</td><td>'+eA(r.SubmitterEmail||"")+'</td><td>'+delBtn+'</td></tr>';
  }).join("");
}

function openSubmitterModal() {
  document.getElementById("submitter-email").value = "";
  document.getElementById("submitter-modal").classList.add("on");
}
function closeSubmitterModal() { document.getElementById("submitter-modal").classList.remove("on"); }

function saveSubmitter() {
  var email = document.getElementById("submitter-email").value.trim().toLowerCase();
  if (!email) { showToast("Email wajib diisi","er"); return; }
  if (submitterRecs.some(function(r){ return (r.SubmitterEmail||"").toLowerCase()===email; })) {
    showToast("Email sudah terdaftar","er"); return;
  }
  createItem("KonfigurasiApproval", { Title:email, TipeKonfigurasi:"SubmitterAccess", SubmitterEmail:email }).then(function(res){
    var newId = res && res.id ? res.id : "tmp-"+Date.now();
    submitterRecs.push({ id:newId, SubmitterEmail:email });
    APPROVAL_CONFIG.submitters.push(email);
    renderSubmitters(); closeSubmitterModal();
    showToast("Pengaju ditambahkan!","ok");
  }).catch(function(e){ showToast("Gagal: "+e.message,"er"); });
}

function deleteSubmitter(id) {
  if (!confirm("Hapus akses pengaju ini?")) return;
  patchItem("KonfigurasiApproval", id, { TipeKonfigurasi:"Deleted" }).then(function(){
    var r = submitterRecs.filter(function(x){ return x.id===id; })[0];
    if (r) {
      var em = (r.SubmitterEmail||"").toLowerCase();
      APPROVAL_CONFIG.submitters = APPROVAL_CONFIG.submitters.filter(function(s){ return s!==em; });
    }
    submitterRecs = submitterRecs.filter(function(x){ return x.id!==id; });
    renderSubmitters();
    showToast("Akses pengaju dihapus","ok");
  }).catch(function(e){ showToast("Gagal: "+e.message,"er"); });
}

var mData = { entitas: [], barang: [], jasa: [] };
var mmTab = "entitas";
var mmEditId = null;

var ML = {
  entitas: "MasterPengadaan",
  barang:  "MasterBarang",
  jasa:    "MasterJasa"
};

function switchMTab(tab) {
  mmTab = tab;
  ["entitas","barang","jasa","setting"].forEach(function(t) {
    document.getElementById("msub-"+t).style.display = t===tab ? "block" : "none";
    document.getElementById("mt-"+t).classList.toggle("on", t===tab);
  });
  if (tab === "setting") {
    refreshL2Display();
    renderRules();
    renderSubmitters();
    if (!apvRules.length) loadApprovalConfig().then(function(){
      refreshL2Display(); renderRules(); renderSubmitters();
    });
  } else if (!mData[tab].length) loadMT(tab);
}

function loadMT(tab) {
  var tbody = document.getElementById("mt-"+tab+"-tbody");
  var cols  = tab==="entitas" ? 7 : 8;
  tbody.innerHTML = '<tr><td colspan="'+cols+'" style="text-align:center;padding:20px;color:var(--g500)">Memuat...</td></tr>';
  getItems(ML[tab]).then(function(items) {
    mData[tab] = items.map(function(i){ return Object.assign({id:i.id}, i.fields); });
    renderMT(tab);
  }).catch(function(e){
    tbody.innerHTML = '<tr><td colspan="'+cols+'" style="text-align:center;padding:20px;color:var(--red)">'+e.message+'</td></tr>';
  });
}

function isAktif(r) { return r.Aktif !== false && r.Aktif !== 0 && r.Aktif !== "0"; }

function renderMT(tab) {
  var data   = mData[tab];
  var search = (document.getElementById("ms-"+tab+"-search")||{}).value || "";
  var fAktif = (document.getElementById("ms-"+tab+"-aktif")||{}).value || "";
  var q = search.toLowerCase();

  var f = data.filter(function(r) {
    var ak = isAktif(r);
    if (fAktif === "1" && !ak) return false;
    if (fAktif === "0" &&  ak) return false;
    if (q) {
      var str = [r.Company,r.Client,r.Project,r.NamaBarang,r.KategoriBarang,r.Subkategori,r.NamaVendor,r.Kategori,r.DomisiliVendor,r.PICVendor].join(" ").toLowerCase();
      if (str.indexOf(q) < 0) return false;
    }
    return true;
  });

  var tbody = document.getElementById("mt-"+tab+"-tbody");
  if (!f.length) { tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:24px;color:var(--g500)">Tidak ada data</td></tr>'; return; }

  var rows = f.map(function(r, i) {
    var ak  = isAktif(r);
    var badge = ak
      ? '<span class="sb apv"><span class="sd"></span>Aktif</span>'
      : '<span class="sb rej"><span class="sd"></span>Nonaktif</span>';
    var editBtn = '<button onclick="openMM(\''+tab+'\',\''+r.id+'\')" class="btn btn-o bsm" style="font-size:11px">Edit</button>';
    var togBtn  = ak
      ? '<button onclick="toggleAktif(\''+tab+'\',\''+r.id+'\',true)" class="btn bsm" style="font-size:11px;color:var(--red);border:1.5px solid var(--red);background:white;border-radius:var(--rsm);cursor:pointer;font-family:inherit">Nonaktifkan</button>'
      : '<button onclick="toggleAktif(\''+tab+'\',\''+r.id+'\',false)" class="btn bsm" style="font-size:11px;color:var(--teal);border:1.5px solid var(--teal);background:white;border-radius:var(--rsm);cursor:pointer;font-family:inherit">Aktifkan</button>';
    var acts = '<div style="display:flex;gap:4px">'+editBtn+togBtn+'</div>';

    if (tab === "entitas") {
      return '<tr><td style="color:var(--g500)">'+(i+1)+'</td><td>'+eA(r.Company||"")+'</td><td style="text-align:center">'+eA(r.CompanyCode||"")+'</td><td>'+eA(r.Client||"")+'</td><td style="font-size:12px">'+eA((r.Project||"").slice(0,35))+((r.Project||"").length>35?"...":"")+'</td><td>'+badge+'</td><td>'+acts+'</td></tr>';
    } else if (tab === "barang") {
      var hg = r.HargaEstimasi ? "Rp "+parseInt(r.HargaEstimasi).toLocaleString("id-ID") : "-";
      return '<tr><td style="color:var(--g500)">'+(i+1)+'</td><td>'+eA(r.NamaBarang||"")+'</td><td>'+eA(r.KategoriBarang||"")+'</td><td>'+eA(r.Subkategori||"")+'</td><td style="text-align:center">'+eA(r.Satuan||"")+'</td><td style="text-align:right;font-weight:500">'+hg+'</td><td>'+badge+'</td><td>'+acts+'</td></tr>';
    } else {
      return '<tr><td style="color:var(--g500)">'+(i+1)+'</td><td>'+eA(r.NamaVendor||"")+'</td><td>'+eA(r.Kategori||"")+'</td><td>'+eA(r.DomisiliVendor||"")+'</td><td>'+eA(r.PICVendor||"")+'</td><td>'+eA(r.NoTelpon||"")+'</td><td>'+badge+'</td><td>'+acts+'</td></tr>';
    }
  });
  tbody.innerHTML = rows.join("");
}

function mf2(id, label, val, req, type) {
  type = type || "text";
  return '<div class="mfield"><label>'+label+(req?' <span class="rq">*</span>':'')+'</label>'+
         '<input type="'+type+'" id="'+id+'" value="'+eA(val||"")+'" placeholder="'+label+'..."></div>';
}

function openMM(tab, id) {
  mmTab = tab; mmEditId = id || null;
  var isEdit = !!id;
  var r = id ? (mData[tab]||[]).filter(function(x){ return x.id===id; })[0] : {};
  if (!r) r = {};

  document.getElementById("mm-title").textContent = (isEdit?"Edit ":"Tambah ") +
    (tab==="entitas"?"Entitas":tab==="barang"?"Barang":"Vendor");

  var body = "";
  if (tab === "entitas") {
    body = mf2("mm-co","Company",r.Company,true) +
           mf2("mm-cc","Company Code",r.CompanyCode,true) +
           mf2("mm-cl","Client",r.Client,true) +
           '<div class="mfield" style="grid-column:span 2"><label>Project <span class="rq">*</span></label>'+
           '<input type="text" id="mm-pr" value="'+eA(r.Project||"")+'" placeholder="Nama project..."></div>';
  } else if (tab === "barang") {
    body = mf2("mm-nb","Nama Barang",r.NamaBarang,true) +
           mf2("mm-kb","Kategori",r.KategoriBarang,true) +
           mf2("mm-sb","Subkategori",r.Subkategori,false) +
           mf2("mm-st","Satuan",r.Satuan,true) +
           '<div class="mfield" style="grid-column:span 2"><label>Harga Estimasi (Rp) <span class="rq">*</span></label>'+
           '<div class="hw"><span class="hp">Rp</span>'+
           '<input type="text" id="mm-hg" class="hi" value="'+(r.HargaEstimasi?parseInt(r.HargaEstimasi).toLocaleString("id-ID"):"")+'" placeholder="0" oninput="fmtHg(this)"></div></div>';
  } else {
    body = mf2("mm-nv","Nama Vendor",r.NamaVendor,true) +
           mf2("mm-kv","Kategori",r.Kategori,false) +
           mf2("mm-dv","Domisili Vendor",r.DomisiliVendor,false) +
           mf2("mm-pv","PIC Vendor",r.PICVendor,false) +
           mf2("mm-av","Alamat",r.Alamat,false) +
           mf2("mm-tv","No. Telpon",r.NoTelpon,false);
  }
  document.getElementById("mm-body").innerHTML = body;
  document.getElementById("mm-overlay").classList.add("on");
}

function closeMM() { document.getElementById("mm-overlay").classList.remove("on"); }

function saveMM() {
  var fields = {};
  var tab = mmTab;

  if (tab === "entitas") {
    fields.Company     = document.getElementById("mm-co").value.trim();
    fields.CompanyCode = document.getElementById("mm-cc").value.trim();
    fields.Client      = document.getElementById("mm-cl").value.trim();
    fields.Project     = document.getElementById("mm-pr").value.trim();
    if (!fields.Company||!fields.CompanyCode||!fields.Client||!fields.Project) { showToast("Semua field wajib diisi","er"); return; }
  } else if (tab === "barang") {
    fields.NamaBarang    = document.getElementById("mm-nb").value.trim();
    fields.KategoriBarang= document.getElementById("mm-kb").value.trim();
    fields.Subkategori   = document.getElementById("mm-sb").value.trim();
    fields.Satuan        = document.getElementById("mm-st").value.trim();
    fields.HargaEstimasi = parseInt((document.getElementById("mm-hg").value||"").replace(/\D/g,""))||0;
    if (!fields.NamaBarang||!fields.KategoriBarang||!fields.Satuan) { showToast("Field bertanda * wajib diisi","er"); return; }
  } else {
    fields.NamaVendor    = document.getElementById("mm-nv").value.trim();
    fields.Kategori      = document.getElementById("mm-kv").value.trim();
    fields.DomisiliVendor= document.getElementById("mm-dv").value.trim();
    fields.PICVendor     = document.getElementById("mm-pv").value.trim();
    fields.Alamat        = document.getElementById("mm-av").value.trim();
    fields.NoTelpon      = document.getElementById("mm-tv").value.trim();
    if (!fields.NamaVendor) { showToast("Nama Vendor wajib diisi","er"); return; }
  }

  if (!fields.Title) fields.Title = fields.Company||fields.NamaBarang||fields.NamaVendor||"";

  var promise = mmEditId
    ? patchItem(ML[tab], mmEditId, fields)
    : createItem(ML[tab], fields);

  promise.then(function(res) {
    if (mmEditId) {
      var idx = mData[tab].findIndex ? mData[tab].findIndex(function(r){ return r.id===mmEditId; }) :
                mData[tab].map(function(r){ return r.id; }).indexOf(mmEditId);
      if (idx > -1) Object.assign(mData[tab][idx], fields);
    } else {
      var newId = res && res.id ? res.id : "tmp-"+Date.now();
      var newR  = Object.assign({id: newId, Aktif: true}, fields);
      mData[tab].push(newR);
      // Refresh master dropdowns
      if (tab==="barang") loadMasterBarang();
      if (tab==="jasa")   loadMasterJasa();
      if (tab==="entitas") loadMaster().then(populateCos);
    }
    renderMT(tab); closeMM();
    showToast((mmEditId?"Data diperbarui":"Data ditambahkan")+" berhasil!","ok");
  }).catch(function(e){ showToast("Gagal menyimpan: "+e.message,"er"); });
}

function toggleAktif(tab, id, currentAktif) {
  var newVal = !currentAktif;
  patchItem(ML[tab], id, { Aktif: newVal }).then(function() {
    var r = (mData[tab]||[]).filter(function(x){ return x.id===id; })[0];
    if (r) r.Aktif = newVal;
    renderMT(tab);
    showToast(newVal ? "Data diaktifkan" : "Data dinonaktifkan", "ok");
    // Refresh dropdowns
    if (tab==="barang") loadMasterBarang();
    if (tab==="jasa")   loadMasterJasa();
    if (tab==="entitas") loadMaster().then(populateCos);
  }).catch(function(e){ showToast("Gagal: "+e.message,"er"); });
}


function hasAccess() {
  // Semua user Microsoft org bisa masuk (minimal lihat Dashboard)
  // Pengajuan Baru hanya untuk yang terdaftar (dicek di initApp)
  var email = myEmail();
  return !!email;
}

function showAccessDenied() {
  document.getElementById("denied-email").textContent = myEmail();
  document.getElementById("ls").style.display = "none";
  document.getElementById("no-access").style.display = "flex";
}

function initApp() {
  gGet("https://graph.microsoft.com/v1.0/me").then(function(me) {
    currentUser = me;
    return loadApprovalConfig();
  }).then(function() {
    var me = currentUser;
    var ini = (me.displayName||"?").split(" ").map(function(w){ return w[0]; }).join("").slice(0,2).toUpperCase();
    document.getElementById("av").textContent = ini;
    document.getElementById("un").textContent  = me.displayName || me.userPrincipalName;

    // Master Data hanya GA
    if (amGA()) document.getElementById("nav-master").style.display = "inline-flex";

    // Master Data hanya GA
    if (amGA()) document.getElementById("nav-master").style.display = "inline-flex";
    // Pengajuan Baru hanya untuk submitter terdaftar
    if (!hasSubmitAccess()) {
      document.getElementById("nav-form").style.display = "none";
      switchView("dash");
    }

    loadMaster().then(function() {
      populateCos();
      loadMasterBarang();
      loadMasterJasa();
      return autoSeq();
    }).catch(function(e){ showToast("Master data gagal: "+e.message,"er"); });

    initDate(); updateRn();
    document.getElementById("ls").style.display = "none";
    document.getElementById("app").classList.add("on");

    // Deep link: auto-open item dari URL parameter ?item=ID
    var urlParams = new URLSearchParams(window.location.search);
    var deepLinkId = urlParams.get("item");
    if (deepLinkId) {
      switchView("dash");
      // Tunggu data dashboard selesai dimuat
      var dlTry = 0;
      var dlInterval = setInterval(function() {
        dlTry++;
        var found = allSubs.filter(function(i){ return i.id === deepLinkId; })[0];
        if (found) { clearInterval(dlInterval); openMo(deepLinkId); }
        else if (dlTry > 20) { clearInterval(dlInterval); showToast("Pengajuan tidak ditemukan","er"); }
      }, 300);
    }
  }).catch(function(e) {
    console.warn("Init gagal:", e);
    showToast("Gagal: "+e.message,"er");
  });
}

// Safe init - no async at top level
window.onload = function() {
  initMsal().then(function() {
    // Handle redirect response (fallback dari loginRedirect saat popup diblokir)
    return _msal.handleRedirectPromise();
  }).then(function(r) {
    if (r && r.account) {
      _msal.setActiveAccount(r.account);
      initApp();
      return;
    }
    var accs = _msal.getAllAccounts();
    if (accs.length > 0) {
      _msal.setActiveAccount(accs[0]);
      initApp();
    }
  }).catch(function(e) {
    console.error("MSAL init gagal:", e);
  });
};

document.getElementById("modal").addEventListener("click", function(e) {
  if (e.target === this) closeMo();
});
