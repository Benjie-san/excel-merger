(function (root, factory) {
  const report = factory();
  if (typeof module === "object" && module.exports) module.exports = report;
  if (root) root.CandataReport = report;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  const PAGE_SIZE = 25;
  const TAX_FIELDS = ["duty", "gst", "pst", "hst"];

  function isDiffering(group) {
    return TAX_FIELDS.some((field) => !group.equality[field]);
  }

  function matchingGroups(groups, query, filter) {
    const term = String(query || "").trim().toLocaleLowerCase();
    return groups.filter((group) =>
      (filter !== "differing" || isDiffering(group)) &&
      (!term || String(group.ccn ?? "").toLocaleLowerCase().includes(term) ||
        group.rows.some((row) => String(row.transaction ?? "").toLocaleLowerCase().includes(term)))
    );
  }

  function pageOf(items, page) {
    const totalPages = Math.max(1, Math.ceil(items.length / PAGE_SIZE));
    const currentPage = Math.min(Math.max(1, page), totalPages);
    return {
      page: currentPage,
      totalPages,
      start: (currentPage - 1) * PAGE_SIZE,
      items: items.slice((currentPage - 1) * PAGE_SIZE, currentPage * PAGE_SIZE)
    };
  }

  function element(tag, className, content) {
    const node = document.createElement(tag);
    if (className) node.className = className;
    if (content !== undefined) node.textContent = String(content);
    return node;
  }

  function table(headers, rows, numericColumns) {
    const scroll = element("div", "candata-table-scroll");
    scroll.tabIndex = 0;
    scroll.setAttribute("role", "region");
    scroll.setAttribute("aria-label", "Scrollable source values");
    const tableNode = element("table", "candata-report-table");
    const head = tableNode.createTHead().insertRow();
    headers.forEach((header, index) => {
      const cell = element("th", numericColumns.includes(index) ? "candata-numeric" : "", header);
      cell.scope = "col";
      head.appendChild(cell);
    });
    const body = tableNode.createTBody();
    rows.forEach((values) => {
      const row = body.insertRow();
      values.forEach((value, index) => {
        const cell = row.insertCell();
        cell.textContent = String(value ?? "");
        if (numericColumns.includes(index)) cell.className = "candata-numeric";
      });
    });
    scroll.appendChild(tableNode);
    return scroll;
  }

  function pager(parent, page, total, onPageChange) {
    const nav = element("nav", "candata-pager");
    nav.setAttribute("aria-label", "Report pages");
    const previous = element("button", "candata-page-button", "Previous");
    previous.type = "button";
    previous.disabled = page <= 1;
    function changePage(target, fallback) {
      onPageChange(target);
      const buttons = parent.querySelectorAll(".candata-pager button");
      const focusTarget = buttons[fallback] && !buttons[fallback].disabled ? buttons[fallback] : buttons[1 - fallback];
      if (focusTarget) focusTarget.focus();
    }
    previous.addEventListener("click", () => changePage(page - 1, 0));
    const count = element("span", "", `Page ${page} of ${total}`);
    const next = element("button", "candata-page-button", "Next");
    next.type = "button";
    next.disabled = page >= total;
    next.addEventListener("click", () => changePage(page + 1, 1));
    nav.append(previous, count, next);
    parent.appendChild(nav);
  }

  function render(container, collisions, duplicateGroups) {
    container.replaceChildren();
    const differingCount = duplicateGroups.filter(isDiffering).length;
    const title = element("h3", "", "Validation report");
    const explanation = element("p", "candata-report-intro", "Findings do not prevent downloads. Duplicate CCNs are shown for review; a duplicate is not itself an error.");
    const counts = element("div", "candata-report-counts");
    [[collisions.length, "existing-HST collisions"], [duplicateGroups.length, "duplicate Header CCNs"],
      [differingCount, "groups with differing taxes"]].forEach(([number, label]) => {
      const item = element("div", "candata-report-count");
      item.append(element("strong", "", number), element("span", "", label));
      counts.appendChild(item);
    });
    container.append(title, explanation, counts);

    const collisionSection = element("section", "candata-report-section");
    collisionSection.appendChild(element("h4", "", "Existing HST when PST was moved"));
    collisionSection.appendChild(element("p", "candata-report-help", "These source rows already had HST before provincial PST was routed into it. Before and after values are shown for review."));
    const collisionResults = element("div", "candata-report-results");
    collisionSection.appendChild(collisionResults);
    container.appendChild(collisionSection);
    let collisionPage = 1;
    function showCollisions() {
      collisionResults.replaceChildren();
      if (!collisions.length) {
        collisionResults.appendChild(element("p", "candata-report-empty", "No existing-HST collisions found in Header or Item."));
        return;
      }
      const current = pageOf(collisions, collisionPage);
      collisionPage = current.page;
      collisionResults.appendChild(element("p", "candata-result-count", `Showing ${current.start + 1}–${current.start + current.items.length} of ${collisions.length} findings`));
      collisionResults.appendChild(table(
        ["Source", "Input row", "CCN", "Transaction", "Province", "PST before", "HST before", "PST after", "HST after"],
        current.items.map((finding) => [finding.type, finding.row, finding.ccn, finding.transaction, finding.province,
          finding.beforePst, finding.beforeHst, finding.afterPst, finding.afterHst]), [1, 5, 6, 7, 8]
      ));
      if (current.totalPages > 1) pager(collisionResults, current.page, current.totalPages, (page) => {
        collisionPage = page;
        showCollisions();
      });
    }
    showCollisions();

    const duplicateSection = element("section", "candata-report-section");
    duplicateSection.appendChild(element("h4", "", "Duplicate Header CCN consistency"));
    duplicateSection.appendChild(element("p", "candata-report-help", "Tax fields compare the original source amounts. The first Header row supplies Duty, GST, PST and HST after routing; only SIMA and Surtax are summed across duplicate rows. “No HST column” identifies a field absent from the source workbook."));
    const controls = element("div", "candata-report-controls");
    const searchLabel = element("label", "", "Search CCN or transaction");
    const search = element("input");
    search.type = "search";
    search.placeholder = "Search duplicate groups";
    searchLabel.appendChild(search);
    const filterLabel = element("label", "", "Show groups");
    const filter = element("select");
    [["differing", "Differing taxes"], ["all", "All duplicates"]].forEach(([value, label]) => {
      const option = element("option", "", label);
      option.value = value;
      filter.appendChild(option);
    });
    filter.value = differingCount ? "differing" : "all";
    filterLabel.appendChild(filter);
    controls.append(searchLabel, filterLabel);
    duplicateSection.appendChild(controls);
    const duplicateResults = element("div", "candata-report-results");
    duplicateSection.appendChild(duplicateResults);
    container.appendChild(duplicateSection);
    let duplicatePage = 1;
    function showGroups() {
      duplicateResults.replaceChildren();
      const matching = matchingGroups(duplicateGroups, search.value, filter.value);
      if (!matching.length) {
        duplicateResults.appendChild(element("p", "candata-report-empty", duplicateGroups.length ?
          "No duplicate groups match this search and filter." : "No duplicate Header CCNs found."));
        return;
      }
      const current = pageOf(matching, duplicatePage);
      duplicatePage = current.page;
      duplicateResults.appendChild(element("p", "candata-result-count", `Showing ${current.start + 1}–${current.start + current.items.length} of ${matching.length} groups`));
      const list = element("div", "candata-group-list");
      current.items.forEach((group) => {
        const details = element("details", "candata-report-group");
        const summary = element("summary");
        summary.append(element("span", "candata-group-ccn", `CCN ${group.ccn}`),
          element("span", "candata-group-meta", `${group.rows.length} input rows`),
          element("span", isDiffering(group) ? "candata-group-status is-different" : "candata-group-status", isDiffering(group) ? "Differing taxes" : "Same taxes"));
        details.appendChild(summary);
        details.addEventListener("toggle", () => {
          if (!details.open || details.childElementCount > 1) return;
          const fields = element("div", "candata-field-tags");
          TAX_FIELDS.forEach((field) => {
            fields.appendChild(element("span", group.equality[field] ? "candata-field-tag" : "candata-field-tag is-different",
              `${field.toUpperCase()} ${group.equality[field] ? "SAME" : "DIFFERENT"}`));
          });
          details.appendChild(fields);
          details.appendChild(table(["Input row", "Transaction", "Province", "Duty", "GST", "PST", "HST"],
            group.rows.map((row) => [row.row, row.transaction, row.province,
              row.raw.duty, row.raw.gst, row.raw.pst, row.raw.hst]), [0, 3, 4, 5, 6]));
        });
        list.appendChild(details);
      });
      duplicateResults.appendChild(list);
      if (current.totalPages > 1) pager(duplicateResults, current.page, current.totalPages, (page) => {
        duplicatePage = page;
        showGroups();
      });
    }
    search.addEventListener("input", () => { duplicatePage = 1; showGroups(); });
    filter.addEventListener("change", () => { duplicatePage = 1; showGroups(); });
    showGroups();
  }

  return { PAGE_SIZE, isDiffering, matchingGroups, pageOf, render };
});
