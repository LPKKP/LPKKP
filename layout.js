"use strict";

(function () {
  function onReady(callback) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", callback);
      return;
    }
    callback();
  }

  function setHidden(el, hidden) {
    if (!el) return;
    el.classList.toggle("lpkkp-hidden", !!hidden);
  }

  function getCurrentPage() {
    var path = (window.location.pathname || "").split("/").pop() || "index.html";
    return path.toLowerCase();
  }

  function markActiveNav() {
    var page = getCurrentPage();
    document.querySelectorAll(".sidebar-nav .nav-link[data-nav]").forEach(function (link) {
      var target = (link.getAttribute("data-nav") || "").toLowerCase();
      var active = target === page;
      link.classList.toggle("active", active);
      if (active) link.setAttribute("aria-current", "page");
      else link.removeAttribute("aria-current");
    });
  }

  function bindLogoutButtons() {
    document.querySelectorAll("[data-logout]").forEach(function (btn) {
      btn.addEventListener("click", function () {
        var auth = typeof getAuth === "function" ? getAuth() : null;
        if (!auth) {
          window.location.href = "login.html?logout=1";
          return;
        }
        auth.signOut().finally(function () {
          try { localStorage.removeItem("userEmail"); } catch (e) {}
          window.location.href = "login.html?logout=1";
        });
      });
    });
  }

  function updateShellUser(user, profile) {
    var email = user && user.email ? user.email.toLowerCase() : "";
    var name = (profile && profile.username) || email || "Pengguna";
    var workspace = (profile && profile.division) || "LPKKP";

    window.adminHMDUser = {
      name: name,
      workspace: workspace,
      avatar: null
    };

    document.querySelectorAll(".sidebar-user strong").forEach(function (el) {
      el.textContent = name;
    });
    document.querySelectorAll(".sidebar-user small").forEach(function (el) {
      el.textContent = workspace;
    });
    document.querySelectorAll(".profile-name").forEach(function (el) {
      el.textContent = name;
    });
    document.querySelectorAll("[data-user-email]").forEach(function (el) {
      el.textContent = email;
    });
  }

  function applyRoleVisibility(role) {
    var isAdmin = role === "admin";
    document.querySelectorAll(".admin-only-nav, .admin-only").forEach(function (el) {
      setHidden(el, !isAdmin);
    });
  }

  window.initLpkkpShell = function (options) {
    options = options || {};
    markActiveNav();
    bindLogoutButtons();

    if (typeof options.role === "string") {
      applyRoleVisibility(options.role);
    }

    if (options.user) {
      updateShellUser(options.user, options.profile || {});
    }
  };

  onReady(function () {
    markActiveNav();
    bindLogoutButtons();
  });
})();
