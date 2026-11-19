/* Shared Excel-like data grid behavior for all tables. */
(function () {
  function cellsInRow(row) {
    return Array.prototype.slice.call(row.querySelectorAll("td")).filter(function (td) {
      return !td.hidden;
    });
  }

  function visibleRows(table) {
    return Array.prototype.slice.call(table.tBodies[0] ? table.tBodies[0].rows : []).filter(function (row) {
      return !row.hidden && row.offsetParent !== null;
    });
  }

  function clearActive(table) {
    table.querySelectorAll("td.is-cell-on").forEach(function (td) {
      td.classList.remove("is-cell-on");
    });
  }

  function activate(td) {
    if (!td || td.tagName !== "TD") return;
    var table = td.closest("table");
    if (!table) return;
    clearActive(table);
    td.classList.add("is-cell-on");
    td.scrollIntoView({ block: "nearest", inline: "nearest" });
  }

  function move(td, dRow, dCol) {
    var table = td.closest("table");
    var rows = visibleRows(table);
    var row = td.parentElement;
    var rowIndex = rows.indexOf(row);
    if (rowIndex < 0) return;
    var cols = cellsInRow(row);
    var colIndex = cols.indexOf(td);
    if (colIndex < 0) return;
    var nextRow = rows[rowIndex + dRow] || rows[rowIndex];
    var nextCols = cellsInRow(nextRow);
    var next = nextCols[Math.min(Math.max(colIndex + dCol, 0), nextCols.length - 1)];
    activate(next);
    var focusable = next.querySelector("input, select, textarea, button, a");
    if (focusable && dRow === 0 && dCol === 0) focusable.focus();
  }

  function enhance(table) {
    if (table.dataset.gridReady === "1") return;
    table.dataset.gridReady = "1";
    table.classList.add("grid");
    if (!table.hasAttribute("tabindex")) table.setAttribute("tabindex", "0");
    var wrap = table.closest(".table-wrap, .sheet-scroll");
    if (wrap) wrap.classList.add("is-grid");

    table.addEventListener("click", function (event) {
      var td = event.target.closest("td");
      if (!td || !table.contains(td)) return;
      activate(td);
      if (!event.target.closest("input, select, textarea, button, a, label")) {
        table.focus({ preventScroll: true });
      }
    });

    table.addEventListener("keydown", function (event) {
      var active = table.querySelector("td.is-cell-on");
      if (!active) return;
      if (event.target.closest("input, select, textarea")) return;
      if (event.key === "ArrowRight") {
        event.preventDefault();
        move(active, 0, 1);
      } else if (event.key === "ArrowLeft") {
        event.preventDefault();
        move(active, 0, -1);
      } else if (event.key === "ArrowDown") {
        event.preventDefault();
        move(active, 1, 0);
      } else if (event.key === "ArrowUp") {
        event.preventDefault();
        move(active, -1, 0);
      } else if (event.key === "Enter" && !event.target.closest("button, a")) {
        var focusable = active.querySelector("input, select, textarea, button, a");
        if (focusable) {
          event.preventDefault();
          focusable.focus();
        }
      } else if (event.key === "Escape") {
        clearActive(table);
      }
    });
  }

  function boot() {
    document.querySelectorAll(".sheet-scroll table, .table-wrap table, table.grid, table.people").forEach(enhance);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }

  window.PayrollDataGrid = { refresh: boot, activate: activate };
})();
