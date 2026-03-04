/**
 * Documentos manager: "Gerar todos" loading state + drawer na lista de ofícios
 */
(function () {
  "use strict";

  // ----- Gerar todos (Central de Documentos) -----
  var formGerarTodos = document.querySelector("[data-gerar-todos]");
  if (formGerarTodos) {
    var btn = formGerarTodos.querySelector(".doc-manager-btn-gerar-todos");
    if (btn) {
      formGerarTodos.addEventListener("submit", function () {
        btn.setAttribute("data-loading", "true");
        btn.disabled = true;
      });
    }
  }

  // ----- Drawer na lista de ofícios -----
  var drawerOverlay = document.getElementById("docDrawerOverlay");
  var drawer = document.getElementById("docDrawer");
  var drawerBody = drawer && document.getElementById("docDrawerBody");
  var drawerTitle = drawer && document.getElementById("docDrawerTitle");
  var openDrawerBtn = document.querySelector("[data-doc-drawer-open]");
  var closeDrawerBtn = document.querySelector("[data-doc-drawer-close]");

  function openDrawer(oficioId, oficioLabel, baseUrl) {
    if (!drawer || !drawerBody) return;
    if (drawerTitle) drawerTitle.textContent = "Documentos — " + (oficioLabel || "Ofício " + oficioId);
    drawerBody.innerHTML = '<p class="doc-drawer-loading-text">Carregando…</p>';
    drawerBody.classList.add("is-loading");
    drawer.classList.add("is-open");
    if (drawerOverlay) drawerOverlay.classList.add("is-open");

    var url = (baseUrl || "/oficios/" + oficioId + "/documentos/").replace(/\/?$/, "") + "/fragment/";
    fetch(url, { headers: { "X-Requested-With": "XMLHttpRequest" } })
      .then(function (r) {
        if (!r.ok) throw new Error("Falha ao carregar");
        return r.text();
      })
      .then(function (html) {
        drawerBody.classList.remove("is-loading");
        drawerBody.innerHTML = html;
      })
      .catch(function () {
        drawerBody.classList.remove("is-loading");
        var centralUrl = url.replace("/fragment/", "/");
        drawerBody.innerHTML = '<p class="doc-drawer-error">Não foi possível carregar os documentos. <a href="' + centralUrl + '">Abrir Central de Documentos</a>.</p>';
      });
  }

  function closeDrawer() {
    if (drawer) drawer.classList.remove("is-open");
    if (drawerOverlay) drawerOverlay.classList.remove("is-open");
  }

  document.querySelectorAll("[data-oficio-id][data-doc-drawer]").forEach(function (el) {
    el.addEventListener("click", function (e) {
      e.preventDefault();
      var id = el.getAttribute("data-oficio-id");
      var label = el.getAttribute("data-oficio-label") || "";
      var baseUrl = el.getAttribute("href");
      openDrawer(id, label, baseUrl);
    });
  });

  // Fallback: if no drawer markup, let the link navigate
  if (!drawer) {
    document.querySelectorAll("[data-doc-drawer]").forEach(function (el) {
      el.removeAttribute("data-doc-drawer");
    });
  }

  if (closeDrawerBtn) closeDrawerBtn.addEventListener("click", closeDrawer);
  if (drawerOverlay) drawerOverlay.addEventListener("click", closeDrawer);
})();
