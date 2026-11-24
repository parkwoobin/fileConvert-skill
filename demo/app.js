'use strict';
const $ = id => document.getElementById(id);
const fileInput = $('file');
const drop = $('drop');
const format = $('format');
const submit = $('submit');
const status = $('status');
const download = $('download');
let settings = null;
let file = null;
let source = '';
let active = false;
let objectUrl = '';
function message(text, error = false) {
    status.textContent = text;
    status.hidden = !text;
    status.className = error ? 'error' : '';
}
function resetDownload() {
    if (objectUrl) URL.revokeObjectURL(objectUrl);
    objectUrl = '';
    download.removeAttribute('href');
    download.hidden = true;
}
function choices(select, values) {
    select.replaceChildren(...values.map(([value, label]) => new Option(label, value)));
}
function selectFile(next) {
    if (active || !settings) return;
    resetDownload();
    message('');
    const ext = next?.name.split('.').pop().toLowerCase();
    const destinations = settings.pairs.filter(([from]) => from === ext);
    if (!next || !next.size || next.size > settings.maxBytes || !destinations.length) {
        file = null;
        source = '';
        choices(format, [['', '파일을 먼저 선택하세요']]);
        $('file-label').textContent = '파일 선택 또는 여기에 끌어 놓기';
        $('file-detail').textContent = 'HWPX · HWP · PDF · DOCX · DOC · PPTX · PPT / 최대 32MB';
        format.disabled = submit.disabled = true;
        if (next) message('지원 형식의 32MB 이하 파일을 선택해 주세요.', true);
        return;
    }
    file = next;
    source = ext;
    $('file-label').textContent = file.name;
    $('file-detail').textContent = `${source.toUpperCase()} · ${(file.size / 1024).toLocaleString('ko-KR', {maximumFractionDigits:1})} KB · 다시 선택 가능`;
    destinations.sort((a,b) => (a[1] === 'docx' ? -1 : b[1] === 'docx' ? 1 : a[1].localeCompare(b[1])));
    choices(format, destinations.map(([,to]) => [to, to.toUpperCase()]));
    format.disabled = submit.disabled = false;
}
fileInput.addEventListener('change', () => selectFile(fileInput.files[0]));
format.addEventListener('change', () => {resetDownload(); message('');});
for (const event of ['dragenter', 'dragover']) drop.addEventListener(event, e => {
    e.preventDefault(); if (!active) drop.classList.add('drag');
});
drop.addEventListener('dragleave', () => drop.classList.remove('drag'));
drop.addEventListener('drop', e => {
    e.preventDefault();
    drop.classList.remove('drag');
    if (!active && e.dataTransfer.files.length) {
        fileInput.files = e.dataTransfer.files;
        selectFile(e.dataTransfer.files[0]);
    }
});
$('convert-form').addEventListener('submit', async e => {
    e.preventDefault();
    if (active || !file || !format.value) return;
    resetDownload();
    active = true;
    fileInput.disabled = format.disabled = submit.disabled = true;
    submit.textContent = '변환 중';
    message('문서를 변환하고 페이지를 확인하고 있습니다.');
    try {
        const query = new URLSearchParams({
            from: source, to: format.value,
        });
        const response = await fetch('/api/convert?' + query, {
            method: 'POST', headers: {'Content-Type': 'application/octet-stream'}, body: file,
        });
        if (!response.ok) {
            const failure = await response.json();
            const pages = failure.source_pages == null ? '' : ` 원본 ${failure.source_pages}페이지, 결과 ${failure.output_pages}페이지.`;
            message(failure.message + pages, true);
            return;
        }
        const result = JSON.parse(response.headers.get('X-Conversion-Report') || '{}');
        objectUrl = URL.createObjectURL(await response.blob());
        download.href = objectUrl;
        download.download = file.name.replace(/\.[^.]+$/, '') + '_변환.' + format.value;
        download.hidden = false;
        message(result.visual_match
            ? `변환 완료 · ${result.output_pages}페이지의 화면이 비교 기준에서 일치했습니다.`
            : `변환 완료 · ${result.output_pages}페이지. 원본과 일부 배치나 글꼴이 다를 수 있습니다.`);
    } catch {
        message('서버와 연결되지 않았습니다. 데모 서버가 실행 중인지 확인해 주세요.', true);
    } finally {
        active = false;
        fileInput.disabled = format.disabled = submit.disabled = false;
        submit.textContent = '변환하기';
    }
});
fileInput.disabled = true;
fetch('/api/formats').then(response => {
    if (!response.ok) throw new Error('formats');
    return response.json();
}).then(config => {
    settings = config;
    fileInput.disabled = false;
    if (fileInput.files[0]) selectFile(fileInput.files[0]);
}).catch(() => message('지원 형식 목록을 불러올 수 없습니다.', true));
