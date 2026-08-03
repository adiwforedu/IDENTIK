// --- State & Constants ---
const APP_STATE_KEY = 'inklusi_pintar_state';
let workbook = null;
let databaseConfig = {
    identitas: {},
    categories: []
};
let answers = {}; // { sheetName_rowIndex: "Ya", ... }
let currentCategoryIndex = -1;
let currentQuestionIndex = 0;

// --- Initialization ---
document.addEventListener('DOMContentLoaded', async () => {
    loadSavedState();
    await initDatabase();
    setupEventListeners();
    initSignature();
});

window.saveCatatanUmum = function() {
    databaseConfig.catatanUmum = document.getElementById('catatan_umum') ? document.getElementById('catatan_umum').value : '';
    saveState();
};

// --- Core Logic: Load Excel ---
async function initDatabase() {
    try {
        // base64 loaded from database_base64.js
        workbook = XLSX.read(database_base64, { type: 'base64' });
        
        parseWorkbookStructure(workbook);
        
        document.getElementById('loader').classList.add('hidden');
        document.getElementById('mainContent').classList.remove('hidden');
    } catch (error) {
        console.error("Gagal memuat database:", error);
        document.getElementById('loader').innerHTML = `
            <div style="color:red; text-align:left; padding:20px; word-break:break-all;">
                <h3>Gagal memuat database Excel</h3>
                <p><strong>Pesan:</strong> ${error.message}</p>
                <p><strong>Detail:</strong> ${error.stack}</p>
                <p>Mohon sampaikan teks ini kepada AI.</p>
            </div>
        `;
    }
}

// --- Parse Excel Structure ---
function parseWorkbookStructure(wb) {
    // Mapping nama sheet Excel -> nama tampil kategori di frontend
    // Urutan sesuai permintaan: 10 kategori
    const SHEET_CATEGORY_MAP = {
        "H. Penglihatan":  "Hambatan Penglihatan",
        "H. Pendengaran":  "Hambatan Pendengaran",
        "H. Intelektual":  "Hambatan Intelektual",
        "H. Fisik Motorik":"Hambatan Fisik Motorik",
        "H. Emosional":    "Hambatan Emosional",
        "AUTISM":          "AUTISM",
        "ADHD":            "ADHD",
        "Slow Leaner":     "Slow Learner",   // typo di file Excel (Leaner bukan Learner)
        "Kesulitan Belajar":"Kesulitan Belajar",
        "CIBI":            "CIBI (Cerdas Istimewa Bakat Istimewa)"
    };

    databaseConfig.categories = [];
    const processedSheets = new Set(); // Cegah sheet diproses lebih dari sekali

    // Loop semua sheet sesuai urutan di Excel
    for (let sheetIdx = 0; sheetIdx < wb.SheetNames.length; sheetIdx++) {
        const sheetName = wb.SheetNames[sheetIdx];

        // Skip jika sheet ini tidak ada di mapping atau sudah diproses
        if (!SHEET_CATEGORY_MAP[sheetName] || processedSheets.has(sheetName)) continue;
        processedSheets.add(sheetName);

        const categoryDisplayName = SHEET_CATEGORY_MAP[sheetName];
        const sheet = wb.Sheets[sheetName];
        if (!sheet || !sheet['!ref']) continue;

        const range = XLSX.utils.decode_range(sheet['!ref']);
        let endRow = Math.min(range.e.r, 2000);
        let endCol = Math.min(range.e.c, 20);

        let questions = [];
        let currentSubCat = "Umum";
        let catInstruction = null;
        let headerRowPassed = false; // Tandai sudah lewati baris header (NO/PERTANYAAN/BOBOT)

        for (let R = range.s.r; R <= endRow; ++R) {
            let rowValues = [];
            for (let C = range.s.c; C <= endCol; ++C) {
                let cell = sheet[XLSX.utils.encode_cell({c: C, r: R})];
                rowValues.push(cell ? cell.v : undefined);
            }

            // === URUTAN CHECK PENTING: header DULU, baru instruksi ===
            const rowString = rowValues.filter(v => v != null).join(" ");
            const rowUpper = rowString.toUpperCase();

            // 1. Deteksi baris header kolom (NO, PERTANYAAN, BOBOT) — set flag lalu skip
            //    HARUS dicek SEBELUM filter YA=1 karena header row F berisi "YA=1, TIDAK=0"
            if (rowUpper.includes("PERTANYAAN") && (rowUpper.includes("BOBOT") || rowUpper.includes("NO"))) {
                headerRowPassed = true;
                continue;
            }

            // 2. Deteksi baris instruksi pengisian (hanya jika bukan header)
            if (!headerRowPassed && (rowUpper.includes("KETIK ANGKA") || rowUpper.includes("KOLOM WARNA KUNING"))) {
                catInstruction = rowString.replace(/\s+/g, ' ').trim();
                continue;
            }

            // Sebelum header ditemukan, skip baris
            if (!headerRowPassed) continue;

            // --- Parsing pertanyaan ---
            let textCandidate = "";
            let noCandidate = "";
            let foundColIndex = -1;

            // Cari teks pertanyaan di kolom B, C, atau D (index 1-3)
            for (let i = 1; i <= 4; i++) {
                if (rowValues[i] && typeof rowValues[i] === 'string' && rowValues[i].length > 8) {
                    textCandidate = rowValues[i];
                    foundColIndex = i;
                    if (rowValues[i-1] !== null && rowValues[i-1] !== undefined && String(rowValues[i-1]).length <= 5) {
                        noCandidate = rowValues[i-1];
                    }
                    break;
                }
            }

            // Fallback: pertanyaan di kolom A (panjang > 20)
            if (!textCandidate && rowValues[0] && typeof rowValues[0] === 'string' && rowValues[0].length > 20) {
                textCandidate = rowValues[0];
                foundColIndex = 0;
            }

            // Deteksi sub-kategori dari kolom A
            if (foundColIndex > 0 && rowValues[0] && typeof rowValues[0] === 'string') {
                const colAText = rowValues[0].trim();
                if (colAText.length > 0 && colAText.length < 50 && colAText.toUpperCase() !== "KATEGORI") {
                    const upperA = colAText.toUpperCase();
                    // Jangan jadikan sub-kategori jika mengandung nama kategori utama atau kata header
                    if (!upperA.includes(categoryDisplayName.toUpperCase().split(' ')[0]) &&
                        upperA !== "NO" && upperA !== "BOBOT" && upperA !== "TEKNIK") {
                        currentSubCat = colAText;
                    }
                }
            }

            if (!textCandidate) continue;

            // Filter teks yang bukan pertanyaan
            const upperText = textCandidate.toUpperCase().trim();
            if (upperText === "PETUNJUK" || upperText === "PERTANYAAN" || upperText === "KATEGORI" ||
                upperText === "KESIMPULAN" || upperText === "NO" || upperText === "BOBOT" ||
                upperText === "TEKNIK" || upperText === "AUTIS" || upperText === "AUTISM" ||
                upperText === "TIDAK TERIDENTIFIKASI" || upperText === "TERIDENTIFIKASI" ||
                upperText === "DIDUGA" || upperText === "TUNANETRA TOTAL" ||
                upperText.startsWith("YA=1") || upperText.startsWith("TIDAK=0") ||
                upperText.includes("KETIK ANGKA") || upperText.includes("KOLOM WARNA KUNING") ||
                upperText.includes("TEMUAN LAIN") || upperText.includes("TULISKAN TEMUAN")) {
                // Simpan sebagai instruksi jika mengandung petunjuk pengisian
                if (upperText.includes("KETIK ANGKA")) catInstruction = textCandidate;
                continue;
            }

            // Ambil bobot (kolom setelah teks) dan teknik (kolom setelah bobot)
            let bobot = 0;
            let teknik = null;
            if (foundColIndex > -1) {
                const bobotVal = rowValues[foundColIndex + 1];
                if (bobotVal !== null && bobotVal !== undefined) bobot = parseFloat(bobotVal) || 0;
                const teknikVal = rowValues[foundColIndex + 2];
                if (teknikVal !== null && teknikVal !== undefined && !isNaN(parseFloat(teknikVal))) {
                    teknik = parseFloat(teknikVal);
                }
            }

            // Gabungkan teks pendek tanpa bobot dengan pertanyaan sebelumnya
            // (menangani kasus teks pertanyaan terpecah di beberapa sel, mis. Slow Learner "Memiliki" + "self image")
            if (bobot === 0 && teknik === null && textCandidate.trim().length < 25 && questions.length > 0) {
                // Gabungkan ke pertanyaan terakhir
                questions[questions.length - 1].text += ' ' + textCandidate.trim();
                continue;
            }

            // Khusus Penglihatan: batasi maks 18 pertanyaan (ada pertanyaan ke-19 yang di-drop)
            if (categoryDisplayName === "Hambatan Penglihatan" && questions.length === 18) continue;

            questions.push({
                id: `q_${sheetName}_${R}`,
                no: noCandidate || (questions.length + 1),
                text: textCandidate,
                subCategory: currentSubCat,
                bobot: bobot,
                teknik: teknik,
                row: R,
                sheetName: sheetName
            });
        }

        databaseConfig.categories.push({
            name: categoryDisplayName,
            title: categoryDisplayName,
            sheetName: sheetName,
            instruction: catInstruction,
            questions: questions
        });
    }

    // Fallback jika tidak ada kategori terdeteksi
    if (databaseConfig.categories.length === 0) {
        Object.values(SHEET_CATEGORY_MAP).forEach((name, idx) => {
            databaseConfig.categories.push({
                name: name, title: name,
                questions: [
                    { id: `fallback_${idx}_1`, no: 1, text: `Apakah peserta didik menunjukkan ciri-ciri utama ${name}?`, row: -1 },
                    { id: `fallback_${idx}_2`, no: 2, text: `Apakah hambatan ini mengganggu aktivitas belajarnya?`, row: -1 }
                ]
            });
        });
    }

    renderSidebar();
}

// --- UI Rendering ---
function renderSidebar() {
    const list = document.getElementById('kategoriList');
    list.innerHTML = '';
    
    databaseConfig.categories.forEach((cat, index) => {
        const li = document.createElement('li');
        li.className = `kategori-item ${index === currentCategoryIndex ? 'active' : ''}`;
        li.innerHTML = `
            <span>${cat.name}</span>
            <i class="fa-solid fa-circle-check status-icon"></i>
        `;
        li.onclick = () => {
            saveCurrentQuestionData();
            currentCategoryIndex = index;
            currentQuestionIndex = 0;
            renderSidebar();
            renderCategory();
        };
        list.appendChild(li);
    });
    
    renderCategory();
    updateProgress();
}

function renderCategory() {
    const container = document.getElementById('pertanyaanContainer');
    const headerPlaceholder = document.getElementById('kategoriHeaderPlaceholder');
    const headerActive = document.getElementById('kategoriHeaderActive');
    const navButtons = document.getElementById('navButtonsContainer');

    const mainArea = document.getElementById('mainPertanyaanArea');

    if (currentCategoryIndex === -1) {
        // State belum memilih kategori
        headerPlaceholder.classList.remove('hidden');
        if (mainArea) mainArea.classList.add('hidden');
        headerActive.classList.add('hidden');
        container.innerHTML = '';
        navButtons.style.display = 'none';
        return;
    }

    // State kategori aktif
    headerPlaceholder.classList.add('hidden');
    if (mainArea) mainArea.classList.remove('hidden');
    headerActive.classList.remove('hidden');
    navButtons.style.display = 'flex';

    const cat = databaseConfig.categories[currentCategoryIndex];
    document.getElementById('kategoriTitle').innerText = cat.title || cat.name;
    
    const subtitleEl = document.getElementById('kategoriPetunjuk');
    if (subtitleEl) {
        if (cat.instruction) {
            subtitleEl.innerText = cat.instruction;
        } else {
            subtitleEl.innerText = "Pilih Ya dan Tidak sesuai dengan gejala yang tampak/ diperoleh";
        }
    }
    
    // Cek apakah sudah di halaman kesimpulan (index 1)
    if (currentQuestionIndex >= 1) {
        renderKesimpulan(cat);
        return;
    }
    
    let html = '';
    let currentSub = null;
    
    let answeredCount = 0;
    cat.questions.forEach(q => {
        if (answers[q.id] && answers[q.id].value) answeredCount++;
    });
    let progressPercent = cat.questions.length > 0 ? (answeredCount / cat.questions.length) * 100 : 0;
    
    html += `
        <div class="progress-info" style="margin-bottom: 1rem;">
            <span id="catProgressText">Terjawab ${answeredCount} dari ${cat.questions.length} Pertanyaan</span>
            <div class="progress-bar"><div id="catProgressFill" class="progress-fill" style="width: ${progressPercent}%"></div></div>
        </div>
    `;

    cat.questions.forEach((q) => {
        const savedAnswer = answers[q.id] || {};
        
        if (q.subCategory && q.subCategory !== "Umum" && q.subCategory !== currentSub) {
            html += `<div style="background: #e2e8f0; padding: 0.75rem 1rem; border-radius: 6px; font-weight: bold; color: #1e293b; margin-bottom: 1rem; margin-top: 1rem; border-left: 4px solid var(--primary-color);">Kategori: ${q.subCategory}</div>`;
            currentSub = q.subCategory;
        }
        
        html += `
        <div class="pertanyaan-card" style="margin-bottom: 0.5rem;">
            <div class="pertanyaan-text">${q.no && q.no !== '-' ? q.no + '. ' : ''}${q.text}</div>
            <div style="display: flex; align-items: center; gap: 0.75rem; margin-top: 0.25rem; flex-wrap: wrap;">
                ${q.bobot ? `<span style="font-size: 0.75rem; color: #64748b; background: #f1f5f9; padding: 2px 8px; border-radius: 4px;">Bobot: <strong>${q.bobot}</strong></span>` : ''}
                ${q.teknik ? `<span style="font-size: 0.75rem; color: #64748b; background: #f1f5f9; padding: 2px 8px; border-radius: 4px;">Teknik: <strong>${q.teknik}</strong></span>` : ''}
            </div>
            <div class="jawaban-options">
                <div class="radio-btn">
                    <input type="radio" name="jawaban_${q.id}" id="ya_${q.id}" value="Ya" ${savedAnswer.value === 'Ya' ? 'checked' : ''} onchange="saveCurrentQuestionData()">
                    <label class="radio-label" for="ya_${q.id}">Ya</label>
                </div>
                <div class="radio-btn">
                    <input type="radio" name="jawaban_${q.id}" id="tidak_${q.id}" value="Tidak" ${savedAnswer.value === 'Tidak' ? 'checked' : ''} onchange="saveCurrentQuestionData()">
                    <label class="radio-label" for="tidak_${q.id}">Tidak</label>
                </div>
            </div>
        </div>
        `;
    });
    
    container.innerHTML = html;

    // Tombol Navigasi
    document.getElementById('btnPrevCat').style.visibility = 'hidden';
    document.getElementById('btnPrevCat').style.display = 'none';
    const btnNext = document.getElementById('btnNextCat');
    btnNext.innerHTML = 'Lihat Kesimpulan <i class="fa-solid fa-flag-checkered"></i>';
    btnNext.style.display = 'inline-flex';
}

function renderKesimpulan(cat) {
    const container = document.getElementById('pertanyaanContainer');
    const isAutismCategory = cat.name && cat.name.toUpperCase().includes('AUTISM');
    
    // Kelompokkan berdasarkan sub-kategori
    const subCategories = {};
    cat.questions.forEach(q => {
        let sub = q.subCategory || "Umum";
        if (!subCategories[sub]) {
            subCategories[sub] = { total: 0, count: 0, bobotTotal: 0, bobotMax: 0 };
        }
        subCategories[sub].total++;
        let b = q.bobot || 0;
        subCategories[sub].bobotMax += b;
        
        if (answers[q.id] && answers[q.id].value === 'Ya') {
            subCategories[sub].count++;
            subCategories[sub].bobotTotal += b;
        }
    });

    const subCatKeys = Object.keys(subCategories);
    let tableRows = '';

    if (isAutismCategory) {
        // Hitung bobot max secara dinamis dari pertanyaan (seharusnya 190 untuk 13 pertanyaan)
        const dynamicBobotMax = cat.questions.reduce((sum, q) => sum + (q.bobot || 0), 0) || 190;
        const stats = {
            count: cat.questions.filter(q => answers[q.id] && answers[q.id].value === 'Ya').length,
            bobotTotal: cat.questions.reduce((sum, q) => sum + ((answers[q.id] && answers[q.id].value === 'Ya') ? (q.bobot || 0) : 0), 0),
            bobotMax: dynamicBobotMax
        };
        const isDanger = stats.bobotTotal >= 100;
        const bobotInfo = stats.bobotMax > 0 ? `<br><span style="font-size: 0.9rem; color: #64748b; font-weight: normal;">(Skor: ${stats.bobotTotal} / ${stats.bobotMax})</span>` : '';
        
        // Simpan kesimpulan Autism ke database config
        databaseConfig.autismKesimpulan = {
            isDiagnosed: isDanger,
            skor: stats.bobotTotal,
            bobotMax: dynamicBobotMax,
            status: isDanger ? 'Teridentifikasi' : 'Tidak Teridentifikasi',
            resultText: isDanger ? 'Diduga Autism' : 'Tidak Teridentifikasi'
        };

        const conclusionText = databaseConfig.autismKesimpulan.resultText;
        const conclusionStyle = isDanger ? 'color: var(--danger); font-weight: bold;' : 'color: inherit; font-weight: normal;';
        tableRows += `
        <tr>
            <td rowspan="1" style="border: 1px solid #cbd5e1; padding: 1rem; text-align: center; font-weight: bold; width: 30%; font-size: 1.1rem; background: #f8fafc;">KESIMPULAN</td>
            <td colspan="2" style="border: 1px solid #cbd5e1; padding: 1rem; ${conclusionStyle};">${conclusionText}${bobotInfo}</td>
        </tr>`;
    } else {
        subCatKeys.forEach((subCat, idx) => {
            const stats = subCategories[subCat];
            let status = stats.count > 0 ? "Teridentifikasi" : "Tidak teridentifikasi";
            let isDanger = stats.count > 0;

            if (subCat.toUpperCase().includes("A. BUTA")) {
                if (stats.bobotTotal >= 100) {
                    status = "Tunanetra Total";
                    isDanger = true;
                } else {
                    status = "Tidak Teridentifikasi";
                    isDanger = false;
                }
            } else if (subCat.toUpperCase().includes("B. LOW VISION")) {
                if (stats.bobotTotal >= 100) {
                    status = "Low Vision";
                    isDanger = true;
                } else {
                    status = "Tidak Teridentifikasi";
                    isDanger = false;
                }
            }

            let label = "Diduga";
            if (subCat !== "Umum") {
                let prefixMatch = subCat.match(/^([A-Z])\./);
                if (prefixMatch) {
                    label = `${prefixMatch[1]}. Diduga`;
                } else {
                    label = `${subCat} Diduga`;
                }
            }

            let bobotInfo = stats.bobotMax > 0 ? `<br><span style="font-size: 0.9rem; color: #64748b; font-weight: normal;">(Skor: ${stats.bobotTotal} / ${stats.bobotMax})</span>` : '';
            
            // Tambahkan skor ke status jika ada
            let displayStatus = status;
            if (isDanger && bobotInfo) {
                displayStatus = status + bobotInfo;
            }

            if (idx === 0) {
                tableRows += `
                <tr>
                    <td rowspan="${subCatKeys.length}" style="border: 1px solid #cbd5e1; padding: 1rem; text-align: center; font-weight: bold; width: 30%; font-size: 1.1rem; background: #f8fafc;">KESIMPULAN</td>
                    <td style="border: 1px solid #cbd5e1; padding: 1rem; width: 40%;">${label}</td>
                    <td style="border: 1px solid #cbd5e1; padding: 1rem; color: ${isDanger ? 'var(--danger)' : 'inherit'}; font-weight: ${isDanger ? 'bold' : 'normal'};">${displayStatus}</td>
                </tr>`;
            } else {
                tableRows += `
                <tr>
                    <td style="border: 1px solid #cbd5e1; padding: 1rem;">${label}</td>
                    <td style="border: 1px solid #cbd5e1; padding: 1rem; color: ${isDanger ? 'var(--danger)' : 'inherit'}; font-weight: ${isDanger ? 'bold' : 'normal'};">${displayStatus}</td>
                </tr>`;
            }
        });
    }

    container.innerHTML = `
        <div class="pertanyaan-card" style="padding: 2rem 1.5rem;">
            <div style="text-align: center; margin-bottom: 2rem;">
                <i class="fa-solid fa-flag-checkered" style="font-size: 3rem; color: var(--success); margin-bottom: 1rem;"></i>
                <h3 style="font-size: 1.5rem; margin-bottom: 0.5rem;">Hasil Identifikasi: ${cat.name}</h3>
                <p style="color: var(--text-muted);">Berikut adalah kesimpulan sementara berdasarkan instrumen yang telah diisi.</p>
            </div>
            
            <table style="width: 100%; border-collapse: collapse; margin-top: 1.5rem; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                <tbody>
                    ${tableRows}
                </tbody>
            </table>
            
            <div style="margin-top: 2rem; text-align: left;">
                <label style="font-weight: bold; margin-bottom: 0.5rem; display: block; color: #1e293b;">Tuliskan temuan lain (jika ada) tentang kondisi anak yang berhubungan dengan hambatan di bawah ini:</label>
                <textarea id="catatan_umum" style="width: 100%; height: 100px; padding: 12px; border: 1px solid #cbd5e1; border-radius: 6px; font-family: inherit; font-size: 1rem; resize: vertical;" placeholder="Ketik temuan lain di sini..." onchange="saveCatatanUmum()">${databaseConfig.catatanUmum || ''}</textarea>
            </div>
            
            <div style="margin-top: 2rem; padding: 1rem; background: rgba(244, 185, 66, 0.1); border-radius: 8px; text-align: center;">
                <p style="margin-bottom: 1rem;">Silakan lanjut ke kategori hambatan berikutnya di menu sidebar (kiri) untuk melanjutkan observasi, atau klik tombol di bawah ini jika sudah selesai.</p>
                <button id="btnEksporKategori" onclick="prosesEkspor()" class="btn btn-success" style="padding: 0.75rem 1.5rem; font-size: 1rem;">
                    <i class="fa-solid fa-file-excel"></i> Ekspor ke Excel
                </button>
            </div>
        </div>
    `;

    document.getElementById('btnPrevCat').style.visibility = 'visible';
    document.getElementById('btnPrevCat').style.display = 'inline-flex';
    document.getElementById('btnNextCat').style.display = 'none';
    
    // Tandai kategori selesai jika kesimpulan dicapai
    const items = document.querySelectorAll('.kategori-item');
    if(items[currentCategoryIndex]) {
        items[currentCategoryIndex].classList.add('completed');
    }
    updateProgress();
    saveState();
}

// --- Event Listeners & State Management ---
function saveIdentitas() {
    const inputs = ['nama', 'ttl', 'jk', 'sekolah', 'kelas', 'alamat_rumah', 'alamat_sekolah', 'tanggal', 'asesor'];
    inputs.forEach(id => {
        const el = document.getElementById('identitas_' + id);
        if (el) {
            databaseConfig.identitas[id] = el.value;
        }
    });
}

function setupEventListeners() {
    document.getElementById('btnMulaiIdentifikasi').addEventListener('click', () => {
        saveIdentitas();
        document.getElementById('stepIdentitas').classList.remove('active');
        document.getElementById('stepIdentitas').classList.add('hidden');
        document.getElementById('stepInstrumen').classList.remove('hidden');
        
        if (currentCategoryIndex === -1 && databaseConfig.categories.length > 0) {
            selectCategory(0);
        }
    });



    document.getElementById('btnEditIdentitas').addEventListener('click', () => {
        document.getElementById('stepInstrumen').classList.add('hidden');
        document.getElementById('stepIdentitas').classList.remove('hidden');
        document.getElementById('stepIdentitas').classList.add('active');
        document.getElementById('btnMulaiIdentifikasi').innerHTML = 'Simpan & Kembali <i class="fa-solid fa-arrow-right"></i>';
    });

    document.getElementById('btnPrevCat').addEventListener('click', () => {
        if (currentQuestionIndex === 1) {
            currentQuestionIndex = 0;
            renderCategory();
        }
    });

    document.getElementById('btnNextCat').addEventListener('click', () => {
        saveCurrentQuestionData();
        currentQuestionIndex = 1;
        renderCategory();
    });

    document.getElementById('btnReset').addEventListener('click', () => {
        if(confirm("Apakah Anda yakin ingin mereset semua data?")) {
            localStorage.removeItem('answers_inklusi_pintar');
            localStorage.removeItem('identitas_inklusi_pintar');
            localStorage.removeItem('catatan_umum_inklusi_pintar');
            localStorage.removeItem('current_cat_inklusi_pintar');
            localStorage.removeItem('current_q_inklusi_pintar');
            location.reload();
        }
    });
}

window.saveCurrentQuestionData = function() {
    if (currentCategoryIndex === -1) return;
    const cat = databaseConfig.categories[currentCategoryIndex];
    if (currentQuestionIndex >= 1) return; // Sedang di layar kesimpulan
    
    let answeredCount = 0;
    cat.questions.forEach(q => {
        const radios = document.getElementsByName('jawaban_' + q.id);
        let val = null;
        for (let r of radios) {
            if (r.checked) val = r.value;
        }
        if (val) {
            answers[q.id] = { value: val, row: q.row, sheetName: q.sheetName };
            answeredCount++;
        }
    });
    
    // Update progress bar secara real-time
    const progressText = document.getElementById('catProgressText');
    const progressFill = document.getElementById('catProgressFill');
    if (progressText && progressFill) {
        progressText.innerText = `Terjawab ${answeredCount} dari ${cat.questions.length} Pertanyaan`;
        progressFill.style.width = `${(answeredCount / cat.questions.length) * 100}%`;
    }
    
    saveState();
}

window.saveCatatanUmum = function() {
    const el = document.getElementById('catatan_umum');
    if (el) {
        databaseConfig.catatanUmum = el.value;
        saveState();
    }
}

function updateProgress() {
    const total = databaseConfig.categories.length;
    const currentIndex = currentCategoryIndex === -1 ? 0 : currentCategoryIndex + 1;
    const percentage = total === 0 ? 0 : (currentIndex / total) * 100;
    
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    
    if (progressFill) progressFill.style.width = `${percentage}%`;
    if (progressText) progressText.innerText = `${currentIndex}/${total}`;
}

window.prosesEkspor = function() {
    if (window.saveCatatanUmum) saveCatatanUmum();
    // Pastikan autism kesimpulan tersimpan sebelum export
    saveState();
    exportToExcel();
}

function saveState() {
    localStorage.setItem('answers_inklusi_pintar', JSON.stringify(answers));
    localStorage.setItem('identitas_inklusi_pintar', JSON.stringify(databaseConfig.identitas));
    localStorage.setItem('catatan_umum_inklusi_pintar', databaseConfig.catatanUmum || '');
    localStorage.setItem('autism_kesimpulan_inklusi_pintar', JSON.stringify(databaseConfig.autismKesimpulan || {}));
    localStorage.setItem('current_cat_inklusi_pintar', currentCategoryIndex);
    localStorage.setItem('current_q_inklusi_pintar', currentQuestionIndex);
}

function loadSavedState() {
    const saved = localStorage.getItem('answers_inklusi_pintar');
    if (saved) {
        try { answers = JSON.parse(saved); } catch(e){}
    }
    const savedIdentitas = localStorage.getItem('identitas_inklusi_pintar');
    if (savedIdentitas) {
        try { 
            databaseConfig.identitas = JSON.parse(savedIdentitas); 
            // Restore UI
            Object.keys(databaseConfig.identitas).forEach(key => {
                const el = document.getElementById('identitas_' + key);
                if (el) el.value = databaseConfig.identitas[key];
            });
        } catch(e){}
    }
    databaseConfig.catatanUmum = localStorage.getItem('catatan_umum_inklusi_pintar') || '';
    
    const savedAutismKesimpulan = localStorage.getItem('autism_kesimpulan_inklusi_pintar');
    if (savedAutismKesimpulan) {
        try { databaseConfig.autismKesimpulan = JSON.parse(savedAutismKesimpulan); } catch(e){}
    }
    
    let catIdx = localStorage.getItem('current_cat_inklusi_pintar');
    if (catIdx !== null) currentCategoryIndex = parseInt(catIdx);
    
    let qIdx = localStorage.getItem('current_q_inklusi_pintar');
    if (qIdx !== null) currentQuestionIndex = parseInt(qIdx);
}

// --- Excel Export (The Magic) ---
async function exportToExcel() {
    try {
        if (!workbook) {
            alert("Database Excel belum termuat sempurna.");
            return;
        }

        // 1. Injeksi Identitas (Tergantung format, misal di Sheet 1)
        const sheetName = workbook.SheetNames[0];
        const btn = document.getElementById('btnEksporKategori');
        let originalText = "Ekspor ke Excel";
        if (btn) {
            originalText = btn.innerHTML;
            btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Mengekspor...';
            btn.disabled = true;
        }

        const excelJsWorkbook = new ExcelJS.Workbook();
        
        // Convert base64 string ke ArrayBuffer (bersihkan whitespace untuk mencegah DOMException)
        const cleanBase64 = database_base64.replace(/\s/g, '');
        const binaryString = window.atob(cleanBase64);
        const len = binaryString.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
            bytes[i] = binaryString.charCodeAt(i);
        }
        
        await excelJsWorkbook.xlsx.load(bytes.buffer);
        
        // Identifikasi sheet yang aktif
        let targetSheetName = null;
        let catNameSafe = "Kategori";

        if (currentCategoryIndex !== -1) {
            const currentCat = databaseConfig.categories[currentCategoryIndex];
            catNameSafe = (currentCat.name || "Kategori").replace(/[^a-zA-Z0-9]/g, '_');
            if (currentCat && currentCat.questions.length > 0) {
                targetSheetName = currentCat.questions[0].sheetName;
            }
        }
        
        // 1. Injeksi Identitas (di sheet yang relevan)
        const firstSheet = targetSheetName ? excelJsWorkbook.getWorksheet(targetSheetName) : excelJsWorkbook.worksheets[0];
        
        // Scan dinamis untuk lokasi label identitas agar kompatibel dengan semua template
        const identitasMap = {
            nama: null, jk: null, ttl: null,
            kelas: null, sekolah: null, alamat_rumah: null, alamat_sekolah: null, tanggal: null, asesor: null
        };

        firstSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                let cellValue = "";
                if (cell.value) {
                    if (typeof cell.value === 'string') cellValue = cell.value;
                    else if (cell.value.richText) cellValue = cell.value.richText.map(rt => rt.text).join('');
                    else cellValue = cell.value.toString();
                }
                
                if (cellValue) {
                    let val = cellValue.toLowerCase().replace(/[^a-z0-9]/g, '');
                    
                    let targetCol = colNumber + 1;
                    const nextCell = row.getCell(colNumber + 1);
                    let nextValue = "";
                    if (nextCell && nextCell.value) {
                        if (typeof nextCell.value === 'string') nextValue = nextCell.value;
                        else if (nextCell.value.richText) nextValue = nextCell.value.richText.map(rt => rt.text).join('');
                        else nextValue = nextCell.value.toString();
                    }
                    if (nextValue.trim() === ':') {
                        targetCol = colNumber + 2;
                    }

                    if (val.includes('namasiswa') || (val.includes('nama') && !val.includes('sekolah') && !val.includes('asesor') && !identitasMap.nama)) identitasMap.nama = { row: rowNumber, col: targetCol };
                    else if ((val.includes('jeniskelamin') || val === 'jk') && !identitasMap.jk) identitasMap.jk = { row: rowNumber, col: targetCol };
                    else if ((val.includes('tanggallahir') || val.includes('ttl') || val.includes('tgllahir') || val.includes('lahir')) && !identitasMap.ttl) identitasMap.ttl = { row: rowNumber, col: targetCol };
                    else if (val.includes('kelas') && val.length < 10 && !identitasMap.kelas) identitasMap.kelas = { row: rowNumber, col: targetCol };
                    else if ((val.includes('namasekolah') || val.includes('madrasah') || (val.includes('sekolah') && !val.includes('alamat'))) && !identitasMap.sekolah) identitasMap.sekolah = { row: rowNumber, col: targetCol };
                    else if (val.includes('alamatrumah') && !identitasMap.alamat_rumah) identitasMap.alamat_rumah = { row: rowNumber, col: targetCol };
                    else if (val.includes('alamatsekolah') && !identitasMap.alamat_sekolah) identitasMap.alamat_sekolah = { row: rowNumber, col: targetCol };
                    else if ((val.includes('tanggalpengisian') || val.includes('tglpemeriksaan') || val === 'tanggal') && !identitasMap.tanggal) identitasMap.tanggal = { row: rowNumber, col: targetCol };
                    else if ((val.includes('namaasesor') || val.includes('asesor')) && !identitasMap.asesor) identitasMap.asesor = { row: rowNumber, col: targetCol };
                }
            });
        });

        Object.keys(identitasMap).forEach(key => {
            if (databaseConfig.identitas[key] && identitasMap[key]) {
                const rowObj = firstSheet.getRow(identitasMap[key].row);
                const cell = rowObj.getCell(identitasMap[key].col);
                cell.value = databaseConfig.identitas[key];
            }
        });

        // 1.5 Bersihkan sel bawaan dari template: set kolom F=0, hapus kolom H (catatan lama)
        //     JANGAN timpa formula di kolom G — biarkan Excel menghitung sendiri via formula
        databaseConfig.categories.forEach(cat => {
            cat.questions.forEach(q => {
                if (q.row > -1 && q.sheetName) {
                    const sheet = excelJsWorkbook.getWorksheet(q.sheetName);
                    if (sheet) {
                        const rowObj = sheet.getRow(q.row + 1);
                        rowObj.getCell(6).value = 0; // Kolom F = 0 (default Tidak)
                        rowObj.getCell(8).value = null; // Hapus kolom H (catatan lama)
                        
                        // Update result-cache formula kolom G agar konsisten dengan F=0
                        const skorCell = rowObj.getCell(7);
                        if (skorCell.value && skorCell.value.formula) {
                            skorCell.value = { formula: skorCell.value.formula, result: 0 };
                        } else if (typeof skorCell.value === 'number') {
                            skorCell.value = { formula: `D${q.row + 1}*F${q.row + 1}`, result: 0 };
                        }
                    }
                }
            });
            
            // Bersihkan juga row yang di-drop (seperti pertanyaan ke-19 Penglihatan)
            if (cat.droppedRows) {
                cat.droppedRows.forEach(droppedRow => {
                    if (cat.questions.length > 0 && cat.questions[0].sheetName) {
                        const sheet = excelJsWorkbook.getWorksheet(cat.questions[0].sheetName);
                        if (sheet) {
                            const rowObj = sheet.getRow(droppedRow + 1);
                            rowObj.getCell(6).value = 0;
                            rowObj.getCell(8).value = null;
                            const skorCell = rowObj.getCell(7);
                            if (skorCell.value && skorCell.value.formula) {
                                skorCell.value = { formula: skorCell.value.formula, result: 0 };
                            }
                        }
                    }
                });
            }
        });

        // 2. Injeksi Jawaban Pengguna — hanya isi kolom F (1/0), update result-cache G
        Object.keys(answers).forEach(qId => {
            const ans = answers[qId];
            if (ans.row > -1 && ans.sheetName) {
                const sheet = excelJsWorkbook.getWorksheet(ans.sheetName);
                if (!sheet) return;
                
                const excelRow = ans.row + 1;
                const rowObj = sheet.getRow(excelRow);
                
                // Isi kolom F dengan 1 atau 0
                const fCell = rowObj.getCell(6);
                const yaValue = (ans.value === 'Ya') ? 1 : 0;
                fCell.value = yaValue;
                
                // Ambil bobot dari kolom D template
                let bobot = 0;
                const dCell = rowObj.getCell(4);
                if (dCell.value !== null && dCell.value !== undefined) {
                    bobot = parseFloat(dCell.value) || 0;
                }
                const calculatedSkor = bobot * yaValue;
                
                // Update result-cache formula kolom G agar formula SUM parent tahu nilainya
                const gCell = rowObj.getCell(7);
                if (gCell.value && gCell.value.formula) {
                    // Pertahankan formula asli, update result-cache
                    gCell.value = { formula: gCell.value.formula, result: calculatedSkor };
                } else {
                    // Buat formula baru jika tidak ada
                    gCell.value = { formula: `D${excelRow}*F${excelRow}`, result: calculatedSkor };
                }
            }
        });

        // 2.2 Injeksi Catatan Umum Temuan Lain
        if (databaseConfig.catatanUmum) {
            excelJsWorkbook.worksheets.forEach(sheet => {
                let foundRow = -1;
                let foundCol = -1;
                sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
                    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                        if (cell.value && typeof cell.value === 'string' && cell.value.toLowerCase().includes('temuan lain')) {
                            foundRow = rowNumber;
                            foundCol = colNumber;
                        }
                    });
                });
                
                if (foundRow > -1) {
                    // Masukkan ke baris di bawahnya (karena biasanya bentuknya teks pengantar panjang)
                    const noteCell = sheet.getRow(foundRow + 1).getCell(foundCol);
                    noteCell.value = databaseConfig.catatanUmum;
                    noteCell.alignment = { wrapText: true, vertical: 'top', horizontal: 'left' };
                }
            });
        }
        
        // 2.3 Update baris summary dan kesimpulan AUTISM secara khusus
        if (targetSheetName && targetSheetName.toUpperCase().includes('AUTISM')) {
            const autismSheet = excelJsWorkbook.getWorksheet(targetSheetName);
            if (autismSheet) {
                const cat = currentCategoryIndex !== -1 ? databaseConfig.categories[currentCategoryIndex] : null;
                
                // Hitung total skor aktual dari jawaban
                let totalSkor = 0;
                if (cat) {
                    cat.questions.forEach(q => {
                        if (answers[q.id] && answers[q.id].value === 'Ya') {
                            totalSkor += (q.bobot || 0);
                        }
                    });
                }
                const isDiagnosed = totalSkor >= 100;
                
                // Scan sheet untuk menemukan:
                // - Baris SUM total (biasanya berisi formula SUM(G...))
                // - Baris KESIMPULAN
                let sumRow = -1;       // baris yang punya formula SUM(G20:G32)
                let kesimpulanRow = -1; // baris yang punya KESIMPULAN
                
                autismSheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
                    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                        const cellVal = cell.value;
                        // Cek formula SUM di kolom G
                        if (colNumber === 7) {
                            if (cellVal && cellVal.formula && cellVal.formula.toUpperCase().includes('SUM(G')) {
                                sumRow = rowNumber;
                            }
                        }
                        // Cek teks KESIMPULAN di kolom A
                        if (colNumber === 1) {
                            let txt = '';
                            if (typeof cellVal === 'string') txt = cellVal;
                            else if (cellVal && cellVal.richText) txt = cellVal.richText.map(rt => rt.text).join('');
                            if (txt.toUpperCase().trim() === 'KESIMPULAN') kesimpulanRow = rowNumber;
                        }
                        // Cek formula IF di kolom E (baris kesimpulan)
                        if (colNumber === 5) {
                            if (cellVal && cellVal.formula && cellVal.formula.toUpperCase().includes('IF(G')) {
                                kesimpulanRow = rowNumber;
                            }
                        }
                    });
                });
                
                // Jika tidak ketemu via scan, gunakan posisi hardcoded dari struktur template
                // (berdasarkan inspeksi: SUM ada di G33, KESIMPULAN ada di row 38)
                if (sumRow === -1) sumRow = 33;
                if (kesimpulanRow === -1) kesimpulanRow = 38;
                
                // Update result-cache formula SUM di baris total
                const sumRowObj = autismSheet.getRow(sumRow);
                const sumCell = sumRowObj.getCell(7); // kolom G
                if (sumCell.value && sumCell.value.formula) {
                    sumCell.value = { formula: sumCell.value.formula, result: totalSkor };
                } else {
                    sumCell.value = { formula: `SUM(G20:G32)`, result: totalSkor };
                }
                
                // Hapus angka 1 di F33 (kolom F baris summary — bawaan template, bukan input user)
                sumRowObj.getCell(6).value = null;
                
                // Update baris KESIMPULAN
                const kesimpulanRowObj = autismSheet.getRow(kesimpulanRow);
                
                // A38 = "Kesimpulan" (set langsung — merged cell A38:C39, hanya isi master cell)
                const cellA = kesimpulanRowObj.getCell(1);
                cellA.value = 'Kesimpulan';
                
                // D38 = "Diduga" (merged cell D38:D39)
                const cellD = kesimpulanRowObj.getCell(4);
                cellD.value = 'Diduga';
                
                // E38 = update result-cache formula IF
                const cellE = kesimpulanRowObj.getCell(5);
                const resultText = isDiagnosed ? 'AUTIS' : 'Tidak teridentifikasi';
                if (cellE.value && cellE.value.formula) {
                    cellE.value = { formula: cellE.value.formula, result: resultText };
                } else {
                    cellE.value = { formula: `IF(G${sumRow}<100,"Tidak teridentifikasi",IF(G${sumRow}>=100,"AUTIS"))`, result: resultText };
                }
                
                // Hapus nilai B38 yang berisi skor lama dari template (jangan null — set ke empty string agar tidak corrupt merged)
                // B38 ada dalam merged range A38:C39, jadi sebaiknya tidak disentuh
                // Cukup pastikan D33 (bobot total) juga di-update
                const d33Cell = autismSheet.getRow(sumRow).getCell(4);
                if (d33Cell.value && typeof d33Cell.value === 'number') {
                    // Biarkan saja nilai bobot total template
                }
            }
        }


        // 2.5 Hapus Warna Kuning dan Biru (agar ramah print)
        excelJsWorkbook.worksheets.forEach(sheet => {
            sheet.eachRow({ includeEmpty: true }, row => {
                row.eachCell({ includeEmpty: true }, cell => {
                    if (cell.fill && cell.fill.type === 'pattern' && cell.fill.fgColor) {
                        const argb = cell.fill.fgColor.argb;
                        if (argb) {
                            const argbUpper = argb.toUpperCase();
                            // Hapus warna kuning (FFFF00) dan biru (0066FF, 0070C0, dll)
                            if (argbUpper.includes('FFFF00') || argbUpper.includes('0066FF') || argbUpper.includes('0070C0') || argbUpper.includes('00B0F0')) {
                                cell.fill = { type: 'pattern', pattern: 'none' };
                            }
                        }
                    }
                });
            });
            // Hapus juga fill dari kolom F secara utuh jika ada
            const colF = sheet.getColumn(6);
            if (colF.fill) {
                colF.fill = undefined;
            }
        });
        
        // 2.6 Bersihkan teks yang tidak diperlukan dari data
        excelJsWorkbook.worksheets.forEach(sheet => {
            sheet.eachRow({ includeEmpty: true }, row => {
                row.eachCell({ includeEmpty: true }, cell => {
                    if (cell.value && typeof cell.value === 'string') {
                        // Bersihkan "belum ada respon" dan "khais"
                        if (cell.value.toLowerCase().includes('belum ada respon') || 
                            cell.value.toLowerCase().includes('khais')) {
                            cell.value = '';
                        }
                    }
                });
            });
        });

        // 3. Hapus Worksheet (Tab/Sheet) Kategori Lain
        if (targetSheetName) {
            const sheetIdsToRemove = [];
            excelJsWorkbook.eachSheet(function(worksheet, sheetId) {
                if (worksheet.name !== targetSheetName) {
                    sheetIdsToRemove.push(sheetId);
                }
            });
            // Eksekusi penghapusan tab sheet yang tidak relevan
            sheetIdsToRemove.forEach(id => {
                excelJsWorkbook.removeWorksheet(id);
            });
        }

        // 4. Generate File
        // Paksa Excel (MS Excel) untuk mengkalkulasi semua formula saat file dibuka
        if (!excelJsWorkbook.calcProperties) excelJsWorkbook.calcProperties = {};
        excelJsWorkbook.calcProperties.fullCalcOnLoad = true;

        const buffer = await excelJsWorkbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        
        // Buat nama file yang relevan (Siswa_Kategori.xlsx)
        let safeName = "Siswa";
        if (databaseConfig.identitas && databaseConfig.identitas.nama) {
            safeName = databaseConfig.identitas.nama.replace(/[^a-zA-Z0-9]/g, '_');
        }
        const fileName = `IDENTIK_${safeName}_${catNameSafe}.xlsx`;

        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
        
        if (btn) {
            btn.innerHTML = '<i class="fa-solid fa-check"></i> Berhasil Diunduh';
            setTimeout(() => {
                btn.innerHTML = originalText;
                btn.disabled = false;
            }, 3000);
        }

    } catch (error) {
        console.error(error);
        alert('Terjadi kesalahan saat mengekspor: ' + error.message);
        document.getElementById('btnSelesai').innerHTML = 'Selesai & Ekspor';
        document.getElementById('btnSelesai').disabled = false;
    }
}

// Inisialisasi awal
loadSavedState();

// Init empty answers if first load
if (Object.keys(answers).length === 0) {
    document.getElementById('stepIdentitas').classList.remove('hidden');
    document.getElementById('stepInstrumen').classList.add('hidden');
}
