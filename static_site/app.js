// ===================================================
// 과천도시공사 인사·ERP 통합검색 프로그램 (정적 HTML 버전)
// app.js : 화면을 그리고, 검색을 처리하는 자바스크립트
//
// data.js 안에 들어있는 DOCS 배열(문서 데이터)을 가지고
// 화면에서 직접 검색/분류/문서보기를 처리합니다. 서버는 필요 없습니다.
// ===================================================

const CATEGORY_ORDER = ["조례", "규정", "세칙", "지침", "기준"];

// 검색어 동의어 처리 (app.py와 동일)
const SYNONYMS = {
    "휴가종류": "휴가구분"
};
const HWPX_SYNONYMS = {
    "휴가정의": "휴가",
    "휴가종류": "휴가"
};

const appEl = document.getElementById("app");

// -----------------------------------------------------
// 공통 유틸 함수
// -----------------------------------------------------
function noExt(filename) {
    const idx = filename.lastIndexOf(".");
    return idx === -1 ? filename : filename.slice(0, idx);
}

// 화면에 보여줄 때는, 파일명 앞에 붙은 "31-1." 같은 내부 관리번호를 떼고 보여준다
// (파일명 자체는 그대로 두고, 화면 표시에만 적용)
function displayTitle(filename) {
    return noExt(filename).replace(/^\d+(-\d+)?\.\s*/, "");
}

function escapeHtml(text) {
    return String(text)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
}

// 검색 결과 안에서, 찾은 키워드를 노란색으로 강조 표시한다.
// html을 태그(<...>)와 일반 글자 부분으로 나눠서, 태그 안쪽은 건드리지 않고
// 일반 글자 부분에서만 강조 표시를 입혀서, 기존 <b>, <table> 같은 태그가 깨지지 않게 한다.
function highlightText(html, keywords) {
    if (!html || !keywords || keywords.length === 0) return html;

    // 띄어쓰기를 무시하고 강조하기 위해,
    // "연차 휴가" → "연차휴가"처럼 붙인 버전도 같이 강조 패턴에 추가한다
    const allPatterns = [];
    keywords.forEach(k => {
        const trimmed = k.trim();
        if (!trimmed) return;
        // 원래 키워드
        allPatterns.push(trimmed.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
        // 띄어쓰기 없앤 버전 (원래랑 다를 때만 추가)
        const noSpace = trimmed.replace(/\s+/g, "");
        if (noSpace !== trimmed) {
            allPatterns.push(noSpace.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"));
        }
    });

    if (allPatterns.length === 0) return html;
    const pattern = new RegExp("(" + allPatterns.join("|") + ")", "gi");
    return html
        .split(/(<[^>]+>)/g)
        .map(part => {
            if (part.startsWith("<") && part.endsWith(">")) return part;
            return part.replace(pattern, '<span class="highlight">$1</span>');
        })
        .join("");
}

function normalizeText(text) {
    return (text || "").replace(/\s+/g, "").toLowerCase();
}

function splitKeywords(keyword) {
    return keyword.toLowerCase().split(/\s+/).filter(Boolean);
}

// -----------------------------------------------------
// 분류 트리 만들기 (app.py의 build_category_tree와 동일한 역할)
// -----------------------------------------------------
function buildCategoryTree() {
    const tree = {};
    DOCS.forEach(doc => {
        if (!tree[doc.category]) tree[doc.category] = [];
        tree[doc.category].push(doc);
    });
    Object.keys(tree).forEach(cat => {
        tree[cat].sort((a, b) => a.docPriority - b.docPriority);
    });
    const ordered = CATEGORY_ORDER.filter(c => tree[c]);
    if (tree["기타"]) ordered.push("기타");
    return { tree, ordered };
}

// -----------------------------------------------------
// 엑셀 표 만들기 (app.py의 build_excel_table_html / build_full_table_html과 동일)
// -----------------------------------------------------
function getMainColumns(header) {
    const SKIP_HEADERS = new Set(["번"]);
    const indices = [];
    for (let i = 0; i < header.length; i++) {
        const h = header[i];
        if (!h) break;
        if (SKIP_HEADERS.has(h)) continue;
        indices.push(i);
    }
    return indices;
}

function formatExcelValue(value) {
    if (value === null || value === undefined) return "";
    if (typeof value === "number") {
        if (Number.isInteger(value)) return String(value);
        return value.toFixed(2);
    }
    return String(value);
}

function buildExcelTableHtml(header, rowValues) {
    const visible = getMainColumns(header);
    let html = '<table class="excel-preview"><tr>';
    visible.forEach(i => { html += `<th>${escapeHtml(header[i])}</th>`; });
    html += "</tr><tr>";
    visible.forEach(i => {
        const value = i < rowValues.length ? rowValues[i] : null;
        html += `<td>${escapeHtml(formatExcelValue(value))}</td>`;
    });
    html += "</tr></table>";
    return html;
}

function buildFullTableHtml(header, rows) {
    const visible = getMainColumns(header);
    let html = '<table class="excel-preview"><tr>';
    visible.forEach(i => { html += `<th>${escapeHtml(header[i])}</th>`; });
    html += "</tr>";
    rows.forEach(row => {
        if (row.every(v => v === null || v === undefined)) return;
        html += "<tr>";
        visible.forEach(i => {
            const value = i < row.length ? row[i] : null;
            html += `<td>${escapeHtml(formatExcelValue(value))}</td>`;
        });
        html += "</tr>";
    });
    html += "</table>";
    return html;
}

// -----------------------------------------------------
// 문서 종류별 "전체 내용 보기" (분류별 둘러보기 화면에서 사용)
// -----------------------------------------------------
function renderFullContent(doc) {
    if (doc.type === "hwpx") {
        return doc.fullHtml;
    }
    if (doc.type === "pdf") {
        return doc.pages.map(p =>
            `<div class="hwpx-line">${p.pageNumber}페이지</div>` +
            `<img src="${p.image}" class="page-preview">`
        ).join("");
    }
    if (doc.type === "xlsx") {
        return doc.sheets.map(sheet =>
            `<div class="hwpx-heading">${escapeHtml(sheet.title)}</div>` +
            buildFullTableHtml(sheet.header, sheet.rows)
        ).join("");
    }
    if (doc.type === "txt") {
        return doc.lines
            .filter(l => l.trim())
            .map(l => `<div class="hwpx-line">${escapeHtml(l.trim())}</div>`)
            .join("");
    }
    return "<p>지원하지 않는 형식입니다.</p>";
}

// -----------------------------------------------------
// 문서 종류별 검색 (app.py의 search_txt / search_pdf / search_xlsx / search_hwpx와 동일)
// -----------------------------------------------------
function searchTxt(doc, keyword) {
    const keywordLower = keyword.toLowerCase();
    const results = [];
    doc.lines.forEach((line, idx) => {
        if (line.toLowerCase().includes(keywordLower)) {
            results.push({ number: idx + 1, text: line.trim() });
        }
    });
    return results;
}

function searchPdf(doc, keyword) {
    const keywords = splitKeywords(keyword);
    const results = [];
    doc.pages.forEach(page => {
        const text = page.text || "";
        if (!text) return;
        const head = text.slice(0, 50);
        if (head.includes("차례") || head.includes("목차")) return;
        const textLower = text.toLowerCase();
        if (keywords.every(w => textLower.includes(w))) {
            const lines = text.split("\n");
            let previewLine = "";
            for (const line of lines) {
                if (keywords.some(w => line.toLowerCase().includes(w))) {
                    previewLine = line.trim();
                    break;
                }
            }
            if (!previewLine) previewLine = (lines[0] || "").trim();
            results.push({ number: page.pageNumber, text: previewLine, image: page.image });
        }
    });
    return results;
}

function searchXlsx(doc, keyword) {
    const rawKeywords = splitKeywords(keyword);
    const keywords = rawKeywords.map(w => SYNONYMS[w] || w);
    const results = [];
    doc.sheets.forEach(sheet => {
        const header = sheet.header;
        const headerLower = header.map(h => h.toLowerCase());
        const headerMatch = keywords.some(word => headerLower.some(h => h.includes(word)));
        if (headerMatch) {
            const tableHtml = buildFullTableHtml(header, sheet.rows);
            results.push({
                location: `${sheet.title} 시트 전체 표`,
                text: `'${keyword}' 관련 전체 표`,
                table: tableHtml
            });
            return; // 이 시트는 표 전체를 보여줬으니, 행 단위 검색은 건너뜀
        }
        sheet.rows.forEach((row, idx) => {
            const rowIndex = idx + 2;
            const rowText = row
                .filter(v => v !== null && v !== undefined)
                .map(v => String(v))
                .join(" ")
                .toLowerCase();
            if (!rowText.trim()) return;
            if (keywords.every(w => rowText.includes(w))) {
                const tableHtml = buildExcelTableHtml(header, row);
                const previewValues = row
                    .filter(v => v !== null && v !== undefined && v !== "")
                    .map(v => String(v));
                const preview = previewValues.join(" / ").slice(0, 80);
                results.push({
                    location: `${sheet.title} 시트, ${rowIndex}행`,
                    text: preview,
                    table: tableHtml
                });
            }
        });
    });
    return results;
}

function searchHwpx(doc, keyword) {
    const rawKeywords = splitKeywords(keyword);
    const keywords = rawKeywords.map(w => HWPX_SYNONYMS[w] || w);
    const results = [];

    const docTitleNorm = normalizeText(doc.docTitlePlain);
    if (docTitleNorm && keywords.every(w => docTitleNorm.includes(normalizeText(w)))) {
        let html = doc.docTitleExtraHtml || "";
        doc.sections.forEach(sec => {
            html += `<div class="hwpx-heading">${sec.headingHtml}</div>` + sec.bodyHtml;
        });
        results.push({ number: doc.docTitlePlain.trim(), text: html, isIndex: false });
        return results;
    }

    doc.sections.forEach(sec => {
        // 조항 제목(헤딩)뿐 아니라, 그 조항의 본문 내용까지 같이 확인한다
        // (제목에는 없지만 본문 안에 검색어가 있는 경우에도, 그 조항 전체를 보여줄 수 있도록)
        const headingNorm = normalizeText(sec.headingPlain);
        const bodyNorm = normalizeText(sec.bodyPlain || "");
        const combinedNorm = headingNorm + bodyNorm;
        if (keywords.every(w => combinedNorm.includes(normalizeText(w)))) {
            results.push({ number: sec.headingPlain, text: sec.bodyHtml, isIndex: false });
        }
    });

    if (results.length === 0 && doc.allBlocks) {
        doc.allBlocks.forEach((block, idx) => {
            if (keywords.every(w => normalizeText(block.plain).includes(normalizeText(w)))) {
                results.push({ number: idx + 1, text: block.html, isIndex: true });
            }
        });
    }

    return results;
}

function searchDoc(doc, keyword) {
    if (doc.type === "txt") return searchTxt(doc, keyword);
    if (doc.type === "pdf") return searchPdf(doc, keyword);
    if (doc.type === "xlsx") return searchXlsx(doc, keyword);
    if (doc.type === "hwpx") return searchHwpx(doc, keyword);
    return [];
}

// -----------------------------------------------------
// 화면 HTML 조립용 도우미
// -----------------------------------------------------
function renderCategoryBrowserHtml(tree, ordered, options) {
    options = options || {};
    const matchedFilenames = options.matchedFilenames || null; // null이면 강조 표시 안 함
    const currentFilename = options.currentFilename || null;

    let html = "";
    ordered.forEach(cat => {
        const docs = tree[cat];
        const hasHit = matchedFilenames
            ? docs.some(d => matchedFilenames.has(d.filename))
            : null;
        const isCurrent = currentFilename ? docs.some(d => d.filename === currentFilename) : false;
        const openAttr = (hasHit || isCurrent) ? "open" : "";

        html += `<details class="cat-tree" ${openAttr}>`;
        html += `<summary>${escapeHtml(cat)}<span class="match-count">(${docs.length}건)</span></summary>`;
        html += '<ul class="doc-list">';
        docs.forEach(d => {
            let cls = "";
            if (matchedFilenames && matchedFilenames.has(d.filename)) cls = "doc-hit";
            if (currentFilename && d.filename === currentFilename) cls = "doc-hit";
            html += `<li><a class="${cls}" href="#browse/${encodeURIComponent(d.filename)}">${escapeHtml(displayTitle(d.filename))}</a></li>`;
        });
        html += "</ul></details>";
    });
    return html;
}

// -----------------------------------------------------
// 화면 1: 메인 페이지
// -----------------------------------------------------
function renderMainView() {
    const { tree, ordered } = buildCategoryTree();

    appEl.innerHTML = `
        <header>
            <div class="header-top">
                <img src="logo.png" alt="과천도시공사 로고" class="header-logo">
                <div class="header-text">
                    <h1>과천도시공사 통합 규정 검색</h1>
                    <p>조례・규정・세칙・지침・기준을 한 번에 검색합니다</p>
                </div>
            </div>
        </header>

        <main>
            <section class="hero-search">
                <h2>찾으시는 규정 내용을 입력하세요</h2>
                <p class="hint">예: 취업규정, 가족돌봄휴가, 재정보증</p>
                <form id="hero-search-form" class="search-form-big">
                    <input type="text" id="hero-search-input" placeholder="검색어를 입력하세요" required>
                    <button type="submit">검색</button>
                </form>
            </section>

            <section class="card category-browser">
                <h3>📚 분류별 둘러보기</h3>
                ${renderCategoryBrowserHtml(tree, ordered)}
            </section>
        </main>

        <footer>
            <img src="logo_full.png" alt="과천도시공사 로고" class="footer-logo">
        </footer>
    `;

    document.getElementById("hero-search-form").addEventListener("submit", e => {
        e.preventDefault();
        const keyword = document.getElementById("hero-search-input").value.trim();
        if (keyword) {
            location.hash = "search/" + encodeURIComponent(keyword);
        }
    });
}

// -----------------------------------------------------
// 화면 2: 검색 결과
// -----------------------------------------------------
function renderSearchView(keyword) {
    const trimmed = (keyword || "").trim();
    if (!trimmed) {
        location.hash = "";
        return;
    }

    const allResults = [];
    const matchedFilenames = new Set();
    const keywords = splitKeywords(trimmed);

    DOCS.forEach(doc => {
        const matches = searchDoc(doc, trimmed);
        if (matches.length > 0) {
            allResults.push({ doc, matches });
            matchedFilenames.add(doc.filename);
        }
    });

    const { tree, ordered } = buildCategoryTree();
    const sidebarHtml = renderCategoryBrowserHtml(tree, ordered, { matchedFilenames });

    let resultsHtml = "";
    if (allResults.length > 0) {
        allResults.forEach(item => {
            resultsHtml += `<section class="card"><details open>`;
            resultsHtml += `<summary>📄 ${escapeHtml(displayTitle(item.doc.filename))}<span class="match-count">(${item.matches.length}건)</span></summary>`;
            resultsHtml += '<ul class="file-list">';
            item.matches.forEach(m => {
                resultsHtml += renderResultItem(item.doc, m, keywords);
            });
            resultsHtml += "</ul></details></section>";
        });
    } else {
        resultsHtml = `
            <section class="card">
                <p class="hint">
                    업로드된 자료에서 해당 내용을 확인하지 못했습니다.<br>
                    정확한 안내를 위해 인사부 담당자에게 확인해 주시기 바랍니다.
                </p>
            </section>`;
    }

    appEl.innerHTML = `
        <header>
            <div class="header-search-top">
                <div>
                    <h1>검색 결과</h1>
                    <p>검색어: "${escapeHtml(trimmed)}"</p>
                    <form id="header-search-form" class="search-form-header">
                        <input type="text" id="header-search-input" placeholder="다른 검색어를 입력하세요" value="${escapeHtml(trimmed)}">
                        <button type="submit">검색</button>
                    </form>
                </div>
                <img src="logo_full.png" alt="과천도시공사 로고" class="header-logo-full">
            </div>
        </header>

        <main>
            <p><a class="back-link" href="#">← 메인으로 돌아가기</a></p>
            <div class="search-layout">
                <aside class="category-filter">
                    <h3>분류</h3>
                    ${sidebarHtml}
                </aside>
                <div class="search-results">
                    ${resultsHtml}
                </div>
            </div>
        </main>
    `;

    document.getElementById("header-search-form").addEventListener("submit", e => {
        e.preventDefault();
        const newKeyword = document.getElementById("header-search-input").value.trim();
        if (newKeyword) {
            location.hash = "search/" + encodeURIComponent(newKeyword);
        }
    });
}

function renderResultItem(doc, m, keywords) {
    let label = "";
    if (doc.type === "pdf") {
        label = `${m.number}페이지`;
    } else if (doc.type === "xlsx") {
        label = escapeHtml(m.location);
    } else if (doc.type === "hwpx") {
        label = m.isIndex ? `${m.number}번째 문단` : escapeHtml(String(m.number));
    } else {
        label = `${m.number}번째 줄 : ${highlightText(escapeHtml(m.text), keywords)}`;
    }

    let html = `<li><strong>${label}</strong>`;

    if (doc.type === "hwpx") {
        html += `<div class="hwpx-section-text">${highlightText(m.text, keywords)}</div>`;
    }
    if (m.image) {
        html += `<br><img src="${m.image}" class="page-preview">`;
    }
    if (m.table) {
        html += highlightText(m.table, keywords);
    }

    html += "</li>";
    return html;
}

// -----------------------------------------------------
// 화면 3: 분류별 둘러보기 - 문서 전체 보기
// -----------------------------------------------------
function renderBrowseView(filename) {
    const doc = DOCS.find(d => d.filename === filename);
    const { tree, ordered } = buildCategoryTree();
    const sidebarHtml = renderCategoryBrowserHtml(tree, ordered, { currentFilename: filename });

    const contentHtml = doc ? renderFullContent(doc) : "<p>문서를 찾을 수 없습니다.</p>";
    const title = doc ? displayTitle(doc.filename) : "문서 없음";

    appEl.innerHTML = `
        <header>
            <div class="header-search-top">
                <div>
                    <h1>${escapeHtml(title)}</h1>
                    <form id="browse-search-form" class="search-form-header">
                        <input type="text" id="browse-search-input" placeholder="검색어를 입력하세요">
                        <button type="submit">검색</button>
                    </form>
                </div>
                <img src="logo_full.png" alt="과천도시공사 로고" class="header-logo-full">
            </div>
        </header>

        <main>
            <p><a class="back-link" href="#">← 메인으로 돌아가기</a></p>
            <div class="search-layout">
                <aside class="category-filter">
                    <h3>분류</h3>
                    ${sidebarHtml}
                </aside>
                <section class="card search-results">
                    ${contentHtml}
                </section>
            </div>
        </main>
    `;

    document.getElementById("browse-search-form").addEventListener("submit", e => {
        e.preventDefault();
        const keyword = document.getElementById("browse-search-input").value.trim();
        if (keyword) {
            location.hash = "search/" + encodeURIComponent(keyword);
        }
    });
}

// -----------------------------------------------------
// 라우팅: 주소(#) 뒷부분을 보고 어떤 화면을 그릴지 결정
// -----------------------------------------------------
function route() {
    const hash = decodeURIComponent(location.hash.replace(/^#/, ""));
    if (hash.startsWith("search/")) {
        renderSearchView(hash.slice("search/".length));
    } else if (hash.startsWith("browse/")) {
        renderBrowseView(hash.slice("browse/".length));
    } else {
        renderMainView();
    }
    window.scrollTo(0, 0);
}

window.addEventListener("hashchange", route);
window.addEventListener("DOMContentLoaded", route);
