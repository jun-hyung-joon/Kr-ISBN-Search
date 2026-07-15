// ======================================================================
// API 인증키
//
// 주소 뒤의 ?key=xxxx 값을 자동으로 읽어옵니다.
// 예) https://.../?key=abcdefg
// ======================================================================
const urlParams = new URLSearchParams(window.location.search);
const API_KEY = urlParams.get('key');

let collectedIsbns = [];
let codeReader = new ZXing.BrowserMultiFormatReader();
let generatedExcelFile = null;
let listState = {
    activeDeleteIndex: null,
    swipeIndex: null,
    pointerId: null,
    pointerActive: false,
    startX: 0,
    startY: 0,
    isDragging: false
};

const UI_TEXT = {
    requiredApiKeyMessage: "접속 주소 뒤에 '?key=본인의API키'를 붙여서 접속해야 정상 작동합니다.\n\n예시:\nhttps://아이디.github.io/저장소/?key=자신의인증키",
    cameraPermissionMessage: "카메라 권한을 승인해주세요.",
    downloadErrorMessage: "파일 다운로드에 실패했습니다. Chrome이나 Safari 브라우저에서 실행해 주세요.",
    analysisCompleteTitle: "분석 완료",
    rescanText: "다시 스캔",
    shareTextIOS: "엑셀 공유 / 저장",
    shareTextPC: "엑셀 다운로드",
    loadingStatusPrefix: "데이터 조회 중"
};

const CONFIG = {
    apiBaseUrl: 'https://www.nl.go.kr/seoji/SearchApi.do',
    scanDelayMs: 300
};

const DOM = {
    scannerSection: document.getElementById('scanner-section'),
    listSection: document.getElementById('list-section'),
    loadingScreen: document.getElementById('loading-screen'),
    completionModal: document.getElementById('completion-modal'),
    scanCount: document.getElementById('scan-count'),
    isbnUl: document.getElementById('isbn-ul'),
    loadingStatus: document.getElementById('loading-status'),
    completionTitle: document.getElementById('completion-title'),
    completionFileName: document.getElementById('completion-file'),
    shareButton: document.getElementById('ios-final-share-btn'),
    cancelButton: document.getElementById('ios-cancel-btn'),
    rescanButton: document.getElementById('rescan-btn'),
    processButton: document.getElementById('process-btn')
};

// ======================================================================
// [엑셀 항목 설정] 
//
// 1. [기본] API에서 가져온 값을 가공 없이 그대로 엑셀에 넣고 싶을 때
//     형식: { header: '국립중앙도서관_API_항목명', width: 엑셀_열_너비 }
//     예시: 정가(PRE_PRICE)를 넣고 싶다면?
//              { header: 'PRE_PRICE', width: 12 },  <-- 이 한 줄만 아래 리스트에 쏙 끼워 넣으세요!
//
// 2. [응용] 가져온 값을 내 입맛대로 넣고 싶을 때
//     형식: { header: '항목명', width: 너비, parser: (원본값) => { 가공 처리 } }
//     예시: 페이지(PAGE) 값 뒤에 '쪽'이라는 글자를 붙여서 넣고 싶다면?
//              { 
//                  header: 'PAGE', 
//                  width: 8, 
//                  parser: (val) => val ? val + '쪽' : '-' 
//              }
//
// 주의 및 약속 사항:
//  - 첫 번째 줄인 'ISBN'은 스캔한 바코드 번호를 담는 필수 자리이므로 절대 지우거나 순서를 바꾸지 마세요.
//  - 데이터가 없을 때 대시('-')를 채워 넣는 작업은 프로그램이 알아서 해줍니다. 편하게 추가해 보세요!
// ======================================================================
const EXCEL_COLUMNS = [
    { header: 'ISBN', width: 16 }, // 첫 컬럼은 스캔된 기본 ISBN이 들어갑니다.
    { header: 'TITLE', width: 25 },
    { header: 'VOL', width: 8 },
    { header: 'AUTHOR', width: 18 },
    { header: 'PUBLISHER', width: 18 },
    { header: 'EA_ISBN', width: 16 },
    {
        header: 'PUBLISH_PREDATE',
        width: 14,
        // 날짜 포맷팅(YYYY-MM-DD) 가공 로직 완벽 유지
        parser: (val) => {
            let dateVal = val || "-";
            if (dateVal !== "-" && dateVal.length === 8 && !isNaN(dateVal)) {
                return `${dateVal.substring(0, 4)}-${dateVal.substring(4, 6)}-${dateVal.substring(6, 8)}`;
            }
            return dateVal;
        }
    },
    {
        header: 'PAGE',
        width: 8,
        // 숫자만 추출하는 가공 로직 완벽 유지
        parser: (val) => {
            if (!val) return "-";
            const match = String(val).match(/\d+/);
            return match ? match[0] : "-";
        }
    }
];

// 화면 제어용 안전 함수
function showView(viewName) {
    const viewMap = {
        scanner: DOM.scannerSection,
        list: DOM.listSection,
        loading: DOM.loadingScreen,
        completion: DOM.completionModal
    };

    Object.entries(viewMap).forEach(([name, el]) => {
        if (el) el.classList.toggle('is-hidden', name !== viewName);
    });
}

function updateScanCount() {
    if (DOM.scanCount) DOM.scanCount.innerText = collectedIsbns.length;
}

function updateLoadingStatus(index, total) {
    if (DOM.loadingStatus) DOM.loadingStatus.innerText = `${UI_TEXT.loadingStatusPrefix} (${index + 1}/${total})`;
}

function setButtonPressed(button, isPressed) {
    if (button) button.classList.toggle('is-pressed', isPressed);
}

function hideCompletionModal() {
    if (DOM.completionModal) DOM.completionModal.classList.add('is-hidden');
}

function initApp() {
    // [안전장치] 주소창에 key 파라미터가 누락된 경우 경고를 띄우고 중단합니다.
    if (!API_KEY) {
        alert(UI_TEXT.requiredApiKeyMessage);
        return;
    }

    const hints = new Map();
    hints.set(ZXing.DecodeHintType.POSSIBLE_FORMATS, [ZXing.BarcodeFormat.EAN_13]);
    codeReader.hints = hints;
    showView('scanner');
    startZxingScanner();
}

function startZxingScanner() {
    hideCompletionModal();
    showView('scanner');

    const constraints = { video: { facingMode: "environment" } };

    codeReader.decodeFromConstraints(constraints, 'video', (result, err) => {
        if (result) {
            try { codeReader.reset(); } catch (e) { }
            const cleanIsbn = String(result.text || '').trim();
            collectedIsbns.push(cleanIsbn);

            updateScanCount();
            openListSection();
        }
    }).catch(err => {
        alert(UI_TEXT.cameraPermissionMessage);
    });
}

function renderList() {
    if (!DOM.isbnUl) return;

    DOM.isbnUl.innerHTML = "";

    if (collectedIsbns.length === 0) {
        const emptyItem = document.createElement('li');
        emptyItem.className = 'isbn-item isbn-item--empty';
        emptyItem.textContent = '아직 스캔된 ISBN이 없습니다.';
        DOM.isbnUl.appendChild(emptyItem);
        return;
    }

    for (let i = collectedIsbns.length - 1; i >= 0; i--) {
        const isbn = collectedIsbns[i];
        const li = document.createElement('li');
        li.className = 'isbn-item';
        li.dataset.index = i;

        if (listState.activeDeleteIndex === i) {
            li.classList.add('is-active-delete');
        }

        const content = document.createElement('div');
        content.className = 'isbn-item__content';

        const value = document.createElement('div');
        value.className = 'isbn-item__value';
        value.textContent = isbn;
        content.appendChild(value);

        const deleteButton = document.createElement('button');
        deleteButton.className = 'isbn-item__delete';
        deleteButton.type = 'button';
        deleteButton.setAttribute('aria-label', '삭제');
        deleteButton.innerHTML = '<svg class="isbn-item__delete-icon" xmlns="http://www.w3.org/2000/svg" x="0px" y="0px" width="48" height="48" viewBox="0 0 64 64" aria-hidden="true"><path fill="currentColor" d="M 28 7 C 25.243 7 23 9.243 23 12 L 23 15 L 13 15 C 11.896 15 11 15.896 11 17 C 11 18.104 11.896 19 13 19 L 15.109375 19 L 16.792969 49.332031 C 16.970969 52.510031 19.600203 55 22.783203 55 L 41.216797 55 C 44.398797 55 47.029031 52.510031 47.207031 49.332031 L 48.890625 19 L 51 19 C 52.104 19 53 18.104 53 17 C 53 15.896 52.104 15 51 15 L 41 15 L 41 12 C 41 9.243 38.757 7 36 7 L 28 7 z M 28 11 L 36 11 C 36.552 11 37 11.449 37 12 L 37 15 L 27 15 L 27 12 C 27 11.449 27.448 11 28 11 z M 19.113281 19 L 44.886719 19 L 43.212891 49.109375 C 43.153891 50.169375 42.277797 51 41.216797 51 L 22.783203 51 C 21.723203 51 20.846109 50.170328 20.787109 49.111328 L 19.113281 19 z M 32 23.25 C 31.033 23.25 30.25 24.034 30.25 25 L 30.25 45 C 30.25 45.966 31.033 46.75 32 46.75 C 32.967 46.75 33.75 45.966 33.75 45 L 33.75 25 C 33.75 24.034 32.967 23.25 32 23.25 z M 24.642578 23.251953 C 23.677578 23.285953 22.922078 24.094547 22.955078 25.060547 L 23.652344 45.146484 C 23.685344 46.091484 24.462391 46.835938 25.400391 46.835938 C 25.421391 46.835938 25.441891 46.835938 25.462891 46.835938 C 26.427891 46.801938 27.183391 45.991391 27.150391 45.025391 L 26.453125 24.939453 C 26.419125 23.974453 25.606578 23.228953 24.642578 23.251953 z M 39.355469 23.251953 C 38.388469 23.224953 37.580875 23.974453 37.546875 24.939453 L 36.849609 45.025391 C 36.815609 45.991391 37.571109 46.801938 38.537109 46.835938 C 38.558109 46.836938 38.578609 46.835938 38.599609 46.835938 C 39.537609 46.835938 40.314656 46.091484 40.347656 45.146484 L 41.044922 25.060547 C 41.078922 24.094547 40.321469 23.285953 39.355469 23.251953 z"></path></svg>';
        deleteButton.addEventListener('click', (event) => {
            event.stopPropagation();
            deleteItems([i]);
        });

        content.appendChild(deleteButton);
        li.appendChild(content);

        li.addEventListener('pointerdown', (event) => startSwipe(event, i, li));
        li.addEventListener('pointermove', (event) => moveSwipe(event, i, li));
        li.addEventListener('pointerup', (event) => endSwipe(event, i, li));
        li.addEventListener('pointercancel', () => resetSwipe(li));

        DOM.isbnUl.appendChild(li);
    }
}

function openListSection() {
    showView('list');
    renderList();
}

function activateDeleteItem(index) {
    if (listState.activeDeleteIndex === index) {
        listState.activeDeleteIndex = null;
    } else {
        listState.activeDeleteIndex = index;
    }
    renderList();
}

function deleteItems(indexes) {
    const sortedIndexes = Array.from(indexes).sort((a, b) => b - a);
    sortedIndexes.forEach(index => {
        collectedIsbns.splice(index, 1);
    });

    listState.activeDeleteIndex = null;
    updateScanCount();
    renderList();
}

function startSwipe(event, index, item) {
    if (event.button !== undefined && event.button !== 0) return;
    if (event.target.closest('.isbn-item__delete')) return;

    listState.pointerId = event.pointerId;
    listState.pointerActive = true;
    listState.swipeIndex = index;
    listState.startX = event.clientX;
    listState.startY = event.clientY;
    listState.isDragging = false;
    item.dataset.swipeStart = 'true';
}

function moveSwipe(event, index, item) {
    if (!listState.pointerActive || listState.pointerId !== event.pointerId || listState.swipeIndex !== index) return;

    const deltaX = event.clientX - listState.startX;
    const deltaY = event.clientY - listState.startY;

    if (Math.abs(deltaX) < 8 && Math.abs(deltaY) < 8) return;
    if (Math.abs(deltaX) <= Math.abs(deltaY)) return;

    event.preventDefault();
    listState.isDragging = true;
}

function endSwipe(event, index, item) {
    if (!listState.pointerActive || listState.pointerId !== event.pointerId || listState.swipeIndex !== index) return;

    if (event.target.closest('.isbn-item__delete')) {
        resetSwipe(item);
        return;
    }

    const deltaX = event.clientX - listState.startX;
    const deltaY = event.clientY - listState.startY;

    if (listState.isDragging && deltaX < -42) {
        listState.activeDeleteIndex = index;
        renderList();
    } else if (Math.abs(deltaX) < 8 && Math.abs(deltaY) < 8) {
        activateDeleteItem(index);
    } else if (listState.activeDeleteIndex === index) {
        listState.activeDeleteIndex = null;
        renderList();
    }

    resetSwipe(item);
}

function resetSwipe(item) {
    listState.pointerActive = false;
    listState.pointerId = null;
    listState.swipeIndex = null;
    listState.isDragging = false;
}

// ======================================================================
// 국립중앙도서관 API 조회
// ======================================================================
async function fetchBookInfo(isbn) {
    const targetUrl = `${CONFIG.apiBaseUrl}?cert_key=${API_KEY}&result_style=json&page_no=1&page_size=1&isbn=${isbn}`;

    try {
        const response = await fetch(targetUrl);
        if (response.ok) {
            const book = await response.json();
            const docs = book.docs || book.items || book.result || [];
            if (docs.length > 0) {
                return parseBookData(docs[0]);
            }
        }
    } catch (e) {
        console.error("ISBN 직접 조회 실패: " + isbn, e);
    }
    return null;
}

// ======================================================================
// API 응답 → Excel 행(Row) 데이터 변환 (EXCEL_COLUMNS 기반 자동 매핑)
// ======================================================================
function parseBookData(book) {
    const parsedRow = {};
    
    EXCEL_COLUMNS.forEach(col => {
        // 'ISBN' 컬럼은 스캔 데이터(바코드 값)이므로 API 응답에서 파싱하지 않고 건너뜁니다.
        if (col.header === 'ISBN') return;

        // 별도의 apiKey 지정이 없으면 header 이름을 그대로 API 키로 사용합니다.
        const apiKey = col.apiKey || col.header;
        const rawVal = book[apiKey];

        if (col.parser) {
            parsedRow[col.header] = col.parser(rawVal, book);
        } else {
            parsedRow[col.header] = (rawVal !== undefined && rawVal !== null) ? String(rawVal).trim() : "-";
        }
    });

    return parsedRow;
}

// ======================================================================
// 조회 실패 시 사용할 기본 행 (EXCEL_COLUMNS 기반 자동 생성)
// ======================================================================
function createEmptyRow(isbn) {
    const emptyRow = {};

    EXCEL_COLUMNS.forEach(col => {
        if (col.header === 'ISBN') {
            emptyRow[col.header] = isbn; // 첫 열은 무조건 스캔된 바코드 번호 배치
        } else {
            emptyRow[col.header] = "-";  // 나머지는 기본값인 대시(-) 지정
        }
    });

    return emptyRow;
}

async function processAllResults() {
    if (collectedIsbns.length === 0) return;

    showView('loading');

    const finalRows = [];

    for (let idx = 0; idx < collectedIsbns.length; idx++) {
        updateLoadingStatus(idx, collectedIsbns.length);

        const currentIsbn = collectedIsbns[idx];
        const info = await fetchBookInfo(currentIsbn);

        // 기본 뼈대 생성 (ISBN 번호 보존 및 나머지 '-' 세팅)
        const rowData = createEmptyRow(currentIsbn);

        // 조회 성공 시 API 데이터로 업데이트
        if (info) {
            Object.assign(rowData, info);
        }

        finalRows.push(rowData);
        await new Promise(resolve => setTimeout(resolve, CONFIG.scanDelayMs));
    }

    const worksheet = XLSX.utils.json_to_sheet(finalRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

    // 엑셀 열 너비 자동 부여
    worksheet['!cols'] = EXCEL_COLUMNS.map(col => ({ wch: col.width }));

    const now = new Date();
    const timestamp = now.getFullYear() +
        String(now.getMonth() + 1).padStart(2, '0') +
        String(now.getDate()).padStart(2, '0') +
        String(now.getHours()).padStart(2, '0') +
        String(now.getMinutes()).padStart(2, '0') +
        String(now.getSeconds()).padStart(2, '0');
    const fileName = `isbn_result_${timestamp}.xlsx`;

    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });

    generatedExcelFile = new File([wbout], fileName, {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });

    showIosCompletionModal(fileName);
}

function getDeviceType() {
    const ua = navigator.userAgent.toLowerCase();
    const isIOS = /iphone|ipad|ipod/.test(ua) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
    const isAndroid = /android/.test(ua);
    const isPC = !isIOS && !isAndroid;

    return { isIOS, isAndroid, isPC };
}

function showIosCompletionModal(fileName) {
    const { isIOS, isAndroid, isPC } = getDeviceType();
    const buttonText = isPC ? UI_TEXT.shareTextPC : UI_TEXT.shareTextIOS;

    if (!DOM.completionModal || !DOM.completionTitle || !DOM.completionFileName || !DOM.shareButton || !DOM.cancelButton) return;

    DOM.completionTitle.textContent = UI_TEXT.analysisCompleteTitle;
    DOM.completionFileName.textContent = fileName;
    DOM.shareButton.textContent = buttonText;
    DOM.cancelButton.textContent = UI_TEXT.rescanText;
    showView('completion');

    bindCompletionModalEvents(fileName, { isIOS, isAndroid, isPC });
}

function bindCompletionModalEvents(fileName, deviceInfo) {
    DOM.shareButton.onclick = async () => {
        if (!generatedExcelFile) return;

        if (deviceInfo.isIOS) {
            if (navigator.share) {
                try {
                    await navigator.share({ files: [generatedExcelFile] });
                    resetAppAndRestart();
                    return;
                } catch (err) {
                    console.log("iOS 공유 실패 또는 취소 -> 일반 다운로드 시도");
                }
            }
            triggerDirectDownload(fileName);
            return;
        }

        if (deviceInfo.isAndroid) {
            if (navigator.share) {
                try {
                    await navigator.share({ files: [generatedExcelFile] });
                    resetAppAndRestart();
                    return;
                } catch (err) {
                    console.log("안드로이드 공유 실패 또는 취소 -> 일반 다운로드 자동 전환");
                }
            }
            triggerDirectDownload(fileName);
            return;
        }

        if (deviceInfo.isPC) {
            triggerDirectDownload(fileName);
            return;
        }
    };

    DOM.cancelButton.onclick = () => {
        startZxingScanner();
    };

    DOM.shareButton.onmousedown = () => setButtonPressed(DOM.shareButton, true);
    DOM.shareButton.onmouseup = () => setButtonPressed(DOM.shareButton, false);
    DOM.shareButton.onmouseleave = () => setButtonPressed(DOM.shareButton, false);
    DOM.shareButton.ontouchstart = () => setButtonPressed(DOM.shareButton, true);
    DOM.shareButton.ontouchend = () => setButtonPressed(DOM.shareButton, false);

    DOM.cancelButton.onmousedown = () => setButtonPressed(DOM.cancelButton, true);
    DOM.cancelButton.onmouseup = () => setButtonPressed(DOM.cancelButton, false);
    DOM.cancelButton.onmouseleave = () => setButtonPressed(DOM.cancelButton, false);
    DOM.cancelButton.ontouchstart = () => setButtonPressed(DOM.cancelButton, true);
    DOM.cancelButton.ontouchend = () => setButtonPressed(DOM.cancelButton, false);
}

function triggerDirectDownload(downloadFileName) {
    try {
        const blobUrl = URL.createObjectURL(generatedExcelFile);
        const a = document.createElement('a');
        a.href = blobUrl;
        a.download = downloadFileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(blobUrl);

        resetAppAndRestart();
    } catch (e) {
        alert(UI_TEXT.downloadErrorMessage);
    }
}

function resetAppAndRestart() {
    collectedIsbns = [];
    updateScanCount();
    startZxingScanner();
}

if (DOM.rescanButton) DOM.rescanButton.onclick = startZxingScanner;
if (DOM.processButton) DOM.processButton.onclick = processAllResults;

initApp();