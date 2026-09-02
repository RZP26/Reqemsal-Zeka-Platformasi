<!DOCTYPE html>
<html lang="az">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Müəllimin Aİ Köməkçisi</title>
    <!-- Mətn və Fayl Kitabxanaları -->
    <script src="https://cdn.jsdelivr.net/npm/marked/marked.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js"></script>
    <!-- PPTX Generator Kitabxanası -->
    <script src="https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@3.12.0/dist/pptxgen.bundle.js"></script>
    <!-- Mermaid Diaqram Kitabxanası -->
    <script src="https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        :root {
            --bg-dark: #1e293b;
            --sidebar-bg: #ffffff;
            --main-bg: #f8fafc;
            --text-main: #0f172a;
            --text-muted: #64748b;
            --accent-blue: #1e3a8a;
            --accent-green: #0d9488;
            --accent-orange: #d97706;
            --border-color: #e2e8f0;
        }

        * {
            box-sizing: border-box;
            margin: 0;
            padding: 0;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }

        body {
            background-color: var(--main-bg);
            color: var(--text-main);
            display: flex;
            flex-direction: column;
            height: 100vh;
            overflow: hidden;
        }

        header {
            background-color: var(--bg-dark);
            color: white;
            padding: 15px 30px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
            z-index: 10;
        }

        .logo-area {
            display: flex;
            align-items: center;
            gap: 12px;
        }

        .logo-area i {
            font-size: 28px;
            color: #38bdf8;
        }

        .logo-text h1 {
            font-size: 18px;
            font-weight: 700;
        }

        .logo-text p {
            font-size: 11px;
            color: #94a3b8;
            text-transform: uppercase;
        }

        .config-bar {
            display: flex;
            align-items: center;
            gap: 20px;
            background: rgba(255, 255, 255, 0.05);
            padding: 6px 15px;
            border-radius: 8px;
            border: 1px solid rgba(255,255,255,0.1);
        }

        .config-item {
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 13px;
        }

        .config-item label {
            color: #94a3b8;
            font-weight: 600;
        }

        .config-item select, .config-item input {
            background: rgba(15, 23, 42, 0.6);
            border: 1px solid rgba(255,255,255,0.2);
            color: white;
            padding: 6px 12px;
            border-radius: 6px;
            font-size: 13px;
            outline: none;
        }

        .config-item input {
            width: 250px;
        }

        .wrapper {
            display: flex;
            flex: 1;
            overflow: hidden;
        }

        sidebar {
            width: 320px;
            background-color: var(--sidebar-bg);
            border-right: 1px solid var(--border-color);
            padding: 20px 10px;
            display: flex;
            flex-direction: column;
            gap: 6px;
            overflow-y: auto;
        }

        .sidebar-title {
            font-size: 11px;
            font-weight: 700;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-left: 10px;
            margin-bottom: 8px;
            margin-top: 10px;
        }

        .menu-item {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 10px 12px;
            border-radius: 8px;
            cursor: pointer;
            transition: all 0.2s ease;
            color: inherit;
            text-decoration: none;
        }

        .menu-item:hover {
            background-color: #f1f5f9;
        }

        .menu-item.active {
            background: linear-gradient(135deg, #1e3a8a, #2563eb);
            color: white;
            box-shadow: 0 4px 12px rgba(37, 99, 235, 0.15);
        }

        .menu-left {
            display: flex;
            align-items: center;
            gap: 12px;
            width: 100%;
        }

        .menu-left i {
            font-size: 18px;
            width: 22px;
            text-align: center;
        }

        .menu-info h2 {
            font-size: 13px;
            font-weight: 600;
            line-height: 1.3;
        }

        .menu-info p {
            font-size: 11px;
            color: var(--text-muted);
        }

        .menu-item.active .menu-info p {
            color: #93c5fd;
        }

        main {
            flex: 1;
            padding: 25px;
            overflow-y: auto;
            display: flex;
            flex-direction: column;
            gap: 20px;
            height: 100%;
        }

        .content-card {
            background: white;
            border-radius: 12px;
            padding: 25px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            border: 1px solid var(--border-color);
        }

        .section-tag {
            font-size: 11px;
            font-weight: 700;
            color: var(--accent-green);
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 5px;
        }

        .section-title {
            font-size: 20px;
            font-weight: 700;
            color: var(--text-main);
            margin-bottom: 6px;
        }

        .section-desc {
            font-size: 13px;
            color: var(--text-muted);
            margin-bottom: 20px;
        }

        .form-grid {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 15px;
            margin-bottom: 15px;
        }

        .form-group {
            display: flex;
            flex-direction: column;
            gap: 6px;
        }

        .form-group.full-width {
            grid-column: span 3;
        }

        .form-group label {
            font-size: 13px;
            font-weight: 600;
        }

        .form-group input, .form-group select, .form-group textarea {
            width: 100%;
            padding: 10px;
            border: 1px solid var(--border-color);
            border-radius: 6px;
            font-size: 13px;
            background-color: #f8fafc;
            outline: none;
        }

        .form-group textarea {
            height: 100px;
            resize: none;
        }

        .submit-btn {
            background-color: var(--accent-green);
            color: white;
            border: none;
            padding: 10px 20px;
            border-radius: 6px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 8px;
        }

        .submit-btn:hover {
            background-color: #0f766e;
        }

        .result-container {
            background: white;
            border-radius: 12px;
            border: 1px solid var(--border-color);
            box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
            display: flex;
            flex-direction: column;
            margin-bottom: 50px;
        }

        .result-header {
            background-color: #1e3a8a;
            color: white;
            padding: 14px 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-size: 13px;
            font-weight: 600;
        }

        .result-actions {
            display: flex;
            gap: 10px;
        }

        .result-btn {
            background: rgba(255, 255, 255, 0.15);
            border: none;
            color: white;
            padding: 6px 14px;
            border-radius: 6px;
            font-size: 12px;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 6px;
            transition: background 0.2s;
        }

        .result-btn:hover {
            background: rgba(255, 255, 255, 0.25);
        }

        .result-btn.word-btn {
            background-color: #2b579a;
        }

        .result-btn.pptx-btn {
            background-color: #d24726;
        }

        .result-body {
            padding: 30px;
            background-color: white;
            line-height: 1.8;
            font-size: 14px;
            color: #1e293b;
        }

        .result-body p {
            margin-bottom: 12px;
        }

        .result-body table {
            width: 100%;
            border-collapse: collapse;
            margin: 15px 0;
        }

        .result-body th, .result-body td {
            border: 1px solid #cbd5e1;
            padding: 8px 12px;
            text-align: left;
        }

        .result-body th {
            background-color: #f1f5f9;
        }

        .result-body img {
            max-width: 100%;
            border-radius: 8px;
            box-shadow: 0 4px 10px rgba(0,0,0,0.1);
            margin: 15px 0;
        }

        .placeholder-box {
            text-align: center;
            color: var(--text-muted);
            padding: 40px 0;
        }

        .placeholder-box i {
            font-size: 40px;
            color: #cbd5e1;
            margin-bottom: 10px;
        }

        .loader {
            display: none;
            align-items: center;
            gap: 10px;
            color: #2563eb;
            font-weight: 600;
            font-size: 13px;
            margin-top: 10px;
        }

        .welcome-box {
            text-align: center;
            padding: 30px 10px;
        }

        .welcome-box i {
            font-size: 50px;
            color: #3b82f6;
            margin-bottom: 15px;
        }

        .author-card {
            margin-top: 25px;
            padding: 15px 20px;
            background-color: #f8fafc;
            border: 1px dashed #cbd5e1;
            border-radius: 10px;
            display: inline-block;
            text-align: left;
        }

        .author-card h3 {
            font-size: 14px;
            font-weight: 700;
            color: var(--text-muted);
            text-transform: uppercase;
            margin-bottom: 8px;
            letter-spacing: 0.5px;
        }

        .author-card p {
            font-size: 13px;
            color: var(--text-main);
            margin-bottom: 4px;
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .author-card p i {
            font-size: 14px;
            color: var(--accent-blue);
            width: 16px;
        }

        .scale-info-box {
            background-color: #eff6ff;
            border-left: 4px solid #2563eb;
            padding: 12px;
            border-radius: 6px;
            margin-bottom: 15px;
            font-size: 12px;
        }

        .scale-table {
            width: 100%;
            margin-top: 5px;
            border-collapse: collapse;
        }

        .scale-table th, .scale-table td {
            border: 1px solid #dbeafe;
            padding: 6px;
            text-align: center;
        }

        .scale-table th {
            background-color: #dbeafe;
        }
    </style>
</head>
<body>

    <header>
        <div class="logo-area">
            <i class="fa-solid fa-graduation-cap"></i>
            <div class="logo-text">
                <h1>Müəllimin Aİ Köməkçisi</h1>
                <p>Genişləndirilmiş Tədris & Vizualization Modulu</p>
            </div>
        </div>
        <div class="config-bar">
            <div class="config-item">
                <label>MODEL</label>
                <select id="modelSelect">
                    <option value="gemini-3.1-flash-lite">Gemini 3.1 Flash Lite ★</option>
                    <option value="gemini-2.5-flash">Gemini 2.5 Flash</option>
                    <option value="gemini-1.5-flash">Gemini 1.5 Flash</option>
                </select>
            </div>
            <div class="config-item">
                <label>API AÇARI</label>
                <input type="password" id="apiKey" placeholder="AIzaSy ilə başlayan API kodunu daxil edin">
            </div>
        </div>
    </header>

    <div class="wrapper">
        
        <sidebar id="menuSidebar">
            <div class="sidebar-title">Əsas Menyu</div>
            
            <div class="menu-item active" onclick="switchTab('ana-sehife', this)">
                <div class="menu-left">
                    <i class="fa-solid fa-house" style="color: #3b82f6;"></i>
                    <div class="menu-info"><h2>Ana Səhifə</h2><p>Sistemə ümumi baxış</p></div>
                </div>
            </div>

            <div class="sidebar-title">Planlaşdırma Modulları</div>

            <div class="menu-item" onclick="switchTab('teqvim-tematik', this)">
                <div class="menu-left">
                    <i class="fa-solid fa-calendar-days" style="color: #8b5cf6;"></i>
                    <div class="menu-info"><h2>Təqvim-tematik plan</h2><p>İllik/Yarımillik planlaşdırma</p></div>
                </div>
            </div>

            <div class="menu-item" onclick="switchTab('ders-plani', this)">
                <div class="menu-left">
                    <i class="fa-solid fa-book-bookmark" style="color: #1e3a8a;"></i>
                    <div class="menu-info"><h2>Dərs planı</h2><p>Gündəlik dərs icmalı</p></div>
                </div>
            </div>

            <div class="menu-item" onclick="switchTab('ders-ssenarisi', this)">
                <div class="menu-left">
                    <i class="fa-solid fa-clapperboard" style="color: #06b6d4;"></i>
                    <div class="menu-info"><h2>Dərs ssenarisi</h2><p>Addım-addım dərs gedişatı</p></div>
                </div>
            </div>

            <div class="sidebar-title">Vizualization & İnteraktiv</div>

            <div class="menu-item" onclick="switchTab('pptx-generator', this)">
                <div class="menu-left">
                    <i class="fa-solid fa-file-powerpoint" style="color: #d24726;"></i>
                    <div class="menu-info"><h2>PPTX Təqdimat</h2><p>PowerPoint slaydları yaradın</p></div>
                </div>
            </div>

            <div class="menu-item" onclick="switchTab('diaqram-sxem', this)">
                <div class="menu-left">
                    <i class="fa-solid fa-diagram-project" style="color: #059669;"></i>
                    <div class="menu-info"><h2>Diaqram & Sxemlər</h2><p>Metodik və ağac sxemləri</p></div>
                </div>
            </div>

            <div class="menu-item" onclick="switchTab('ai-sekil', this)">
                <div class="menu-left">
                    <i class="fa-solid fa-image" style="color: #e11d48;"></i>
                    <div class="menu-info"><h2>Aİ Şəkil & İllüstrasiya</h2><p>Əyani dərs materialı çəkdirin</p></div>
                </div>
            </div>

            <div class="sidebar-title">Qiymətləndirmə & Metodika</div>

            <div class="menu-item" onclick="switchTab('interaktiv-metodlar', this)">
                <div class="menu-left">
                    <i class="fa-solid fa-comments" style="color: #f59e0b;"></i>
                    <div class="menu-info"><h2>İnteraktiv metodlar</h2><p>Beyin həmləsi, qrup işləri</p></div>
                </div>
            </div>

            <div class="menu-item" onclick="switchTab('ksq', this)">
                <div class="menu-left">
                    <i class="fa-solid fa-file-signature" style="color: #ea580c;"></i>
                    <div class="menu-info"><h2>Kiçik Summativ (KSQ)</h2><p>Təlim nəticəsi (20 Sual)</p></div>
                </div>
            </div>

            <div class="menu-item" onclick="switchTab('bsq', this)">
                <div class="menu-left">
                    <i class="fa-solid fa-bullseye" style="color: #e11d48;"></i>
                    <div class="menu-info"><h2>Böyük Summativ (BSQ)</h2><p>Səriştə qiymətləndirmə (20 Sual)</p></div>
                </div>
            </div>

            <div class="menu-item" onclick="switchTab('yoxlanma', this)">
                <div class="menu-left">
                    <i class="fa-solid fa-square-check" style="color: #16a34a;"></i>
                    <div class="menu-info"><h2>Nəticələrin yoxlanması</h2><p>Yazı, PDF və Şəkil analizi</p></div>
                </div>
            </div>

            <div class="menu-item" onclick="switchTab('ev-tapsiriqlari', this)">
                <div class="menu-left">
                    <i class="fa-solid fa-house-laptop" style="color: #6366f1;"></i>
                    <div class="menu-info"><h2>Ev tapşırıqları</h2><p>Yaradıcı və möhkəmləndirici</p></div>
                </div>
            </div>

            <div class="sidebar-title">Açıq Rejim</div>

            <div class="menu-item" onclick="switchTab('ai-cat', this)">
                <div class="menu-left">
                    <i class="fa-solid fa-robot" style="color: #0d9488;"></i>
                    <div class="menu-info"><h2>AI Çat</h2><p>Sərbəst rejimdə köməkçi</p></div>
                </div>
            </div>
        </sidebar>

        <main>
            
            <div id="dynamicFormCard" class="content-card">
                <div class="welcome-box">
                    <i class="fa-solid fa-wand-magic-sparkles"></i>
                    <h2 class="section-title">Müəllimin Süni İntellekt Assistantına Xoş Gəldiniz!</h2>
                    <p class="section-desc">Bütün alt menyular tənzimləndi və aktivləşdirildi. İşə başlamaq üçün sol paneldən keçid edin.</p>

                    <div class="author-card">
                        <h3>Təsisçi və Təlimçi</h3>
                        <p><i class="fa-solid fa-user"></i> Tacəddin İbadov</p>
                        <p><i class="fa-solid fa-phone"></i> +99455-670-71-16</p>
                        <p><i class="fa-solid fa-envelope"></i> t.ibadov88@gmail.com</p>
                    </div>
                </div>
            </div>

            <div class="result-container">
                <div class="result-header">
                    <span id="resultPanelTitle"><i class="fa-solid fa-robot"></i> Aİ Cavab Paneli</span>
                    <div class="result-actions">
                        <button class="result-btn pptx-btn" onclick="exportToPowerPoint()"><i class="fa-solid fa-file-powerpoint"></i> PPTX Yüklə</button>
                        <button class="result-btn word-btn" onclick="exportToWord()"><i class="fa-solid fa-file-word"></i> Word Sənədi Yüklə</button>
                        <button class="result-btn" onclick="copyResult()"><i class="fa-solid fa-copy"></i> Mətni Kopyala</button>
                    </div>
                </div>
                <div class="result-body" id="outputBody">
                    <div class="placeholder-box" id="placeholderBox">
                        <i class="fa-regular fa-clipboard"></i>
                        <p>Menyulardan birini seçin, lazımi xanaları doldurun və sorğunu göndərin.</p>
                    </div>
                </div>
            </div>

            <div class="loader" id="globalLoader">
                <i class="fa-solid fa-circle-notch fa-spin"></i> Süni İntellekt rəsmi sənəd və vizual struktura uyğun cavab hazırlayır, zəhmət olmasa gözləyin...
            </div>

        </main>
    </div>

<script>
    // PDF.js worker və Mermaid inicializasiyası
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';
    mermaid.initialize({ startOnLoad: false, theme: 'default' });

    let activeTab = 'ana-sehife';
    let base64Image = "";
    let imageMimeType = "image/jpeg";
    let extractedPdfText = "";
    let generatedPptxSlides = [];

    function switchTab(tabName, element) {
        activeTab = tabName;
        base64Image = "";
        imageMimeType = "image/jpeg";
        extractedPdfText = "";
        generatedPptxSlides = [];
        
        document.querySelectorAll('.menu-item').forEach(item => item.classList.remove('active'));
        if(element) {
            element.classList.add('active');
        } else {
            document.querySelectorAll('.menu-item').forEach(item => {
                if(item.getAttribute('onclick') && item.getAttribute('onclick').includes(`'${tabName}'`)) {
                    item.classList.add('active');
                }
            });
        }

        const formCard = document.getElementById('dynamicFormCard');
        if (!formCard) return;

        if(tabName === 'ana-sehife') {
            formCard.innerHTML = `
                <div class="welcome-box">
                    <i class="fa-solid fa-wand-magic-sparkles"></i>
                    <h2 class="section-title">Müəllimin Süni İntellekt Assistantına Xoş Gəldiniz!</h2>
                    <p class="section-desc">Bütün menyular rəsmi standartlara uyğun tənzimlənib. İşinizə başlamaq üçün sol paneldən bir mövzu seçin.</p>
                    
                    <div class="author-card">
                        <h3>Təsisçi və Təlimçi</h3>
                        <p><i class="fa-solid fa-user"></i> Tacəddin İbadov</p>
                        <p><i class="fa-solid fa-phone"></i> +99455-670-71-16</p>
                        <p><i class="fa-solid fa-envelope"></i> t.ibadov88@gmail.com</p>
                    </div>
                </div>`;
            return;
        }

        if(tabName === 'pptx-generator') {
            formCard.innerHTML = `
                <div class="section-tag"><i class="fa-solid fa-file-powerpoint"></i> VİZUAL MODUL</div>
                <h2 class="section-title">PowerPoint (.pptx) Slayd Generatoru</h2>
                <p class="section-desc">Mövzunu daxil edin, Aİ dərsi slayd-slayd struktura salacaq və sistem birbaşa .pptx faylı yükləməyə imkan verəcək.</p>
                <div class="form-grid">
                    <div class="form-group">
                        <label>Fənn / Modul <span>*</span></label>
                        <input type="text" id="inpFenn" placeholder="məs. Biologiya, Tarix...">
                    </div>
                    <div class="form-group">
                        <label>Sinif / Qrup <span>*</span></label>
                        <input type="text" id="inpSinif" placeholder="məs. 9-cu sinif">
                    </div>
                    <div class="form-group">
                        <label>Slayd Sayı</label>
                        <input type="text" id="inpParam" placeholder="məs. 5 slaydlıq">
                    </div>
                    <div class="form-group full-width">
                        <label>Slayd Mövzusu və Əsas Qeydlər <span>*</span></label>
                        <textarea id="inpMovzu" placeholder="Slaydda yer almasını istədiyiniz əsas faktları yazın..."></textarea>
                    </div>
                </div>
                <button class="submit-btn" onclick="fireGemini()"><i class="fa-solid fa-sliders"></i> PowerPoint Slaydları Hazırla</button>`;
        }
        else if(tabName === 'diaqram-sxem') {
            formCard.innerHTML = `
                <div class="section-tag"><i class="fa-solid fa-diagram-project"></i> VİZUAL MODUL</div>
                <h2 class="section-title">İnteraktiv Diaqram və Sxem Generatoru</h2>
                <p class="section-desc">Mövzuya uyğun ağac sxemləri, anlayış xəritələri və blok-sxemlər hazırlayır.</p>
                <div class="form-grid">
                    <div class="form-group">
                        <label>Fənn <span>*</span></label>
                        <input type="text" id="inpFenn" placeholder="məs. Fizika, Coğrafiya...">
                    </div>
                    <div class="form-group">
                        <label>Sxem Növü</label>
                        <input type="text" id="inpParam" placeholder="məs. Ağac sxemi, Təsnifat, Proses">
                    </div>
                    <div class="form-group full-width">
                        <label>Mövzu və Struktura Dair Qeydlər <span>*</span></label>
                        <textarea id="inpMovzu" placeholder="Sxemə çevriləcək anlayışları qeyd edin..."></textarea>
                    </div>
                </div>
                <button class="submit-btn" onclick="fireGemini()"><i class="fa-solid fa-network-wired"></i> Sxem Çək və Vizuallaşdır</button>`;
        }
        else if(tabName === 'ai-sekil') {
            formCard.innerHTML = `
                <div class="section-tag"><i class="fa-solid fa-image"></i> VİZUAL MODUL</div>
                <h2 class="section-title">Aİ Şəkil və Əyani Dərs İllüstrasiyası</h2>
                <p class="section-desc">Dərslər üçün pedaqoji və tədris yönümlü illüstrasiya və vizual görüntülər yaradın.</p>
                <div class="form-grid">
                    <div class="form-group full-width">
                        <label>Yaratmaq istədiyiniz visual görüntünün təsviri <span>*</span></label>
                        <textarea id="inpMovzu" placeholder="məs. Kosmosda planetlərin günəş ətrafında dövr etməsini əks etdirən rəngli tədris illüstrasiyası..."></textarea>
                    </div>
                </div>
                <button class="submit-btn" onclick="fireGemini()"><i class="fa-solid fa-paint-brush"></i> Şəkli Generator İlə Çək</button>`;
        }
        else if(tabName === 'ksq' || tabName === 'bsq') {
            const title = tabName === 'ksq' ? 'Kiçik Summativ Qiymətləndirmə (KSQ)' : 'Böyük Summativ Qiymətləndirmə (BSQ)';
            const desc = tabName === 'ksq' ? 'Təlim nəticələrinin qiymətləndirilməsi rəsmi sənədi' : 'Səriştələrin qiymətləndirilməsi rəsmi sənədi';
            formCard.innerHTML = `
                <div class="section-tag"><i class="fa-solid fa-file-invoice"></i> RƏSMİ STANDART</div>
                <h2 class="section-title">${title}</h2>
                <p class="section-desc">${desc}. Sistem avtomatik olaraq asandan çətinə doğru 20 unikal sual hazırlayacaq.</p>
                
                <div class="scale-info-box">
                    <strong>Meyar və 5 Ballıq Qiymətləndirmə Şkalası:</strong>
                    <table class="scale-table">
                        <tr><th>Düzgün Faiz Aralığı</th><th>Qiymət</th><th>Səviyyə</th></tr>
                        <tr><td>0% - 30%</td><td>2 (Qeyri-kafi)</td><td>I Səviyyə</td></tr>
                        <tr><td>30% - 60%</td><td>3 (Kafi)</td><td>II Səviyyə</td></tr>
                        <tr><td>60% - 80%</td><td>4 (Yaxşı)</td><td>III Səviyyə</td></tr>
                        <tr><td>80% - 100%</td><td>5 (Əla)</td><td>IV Səviyyə</td></tr>
                    </table>
                </div>

                <div class="form-grid">
                    <div class="form-group">
                        <label>Fənn / Modul Adı <span>*</span></label>
                        <input type="text" id="inpFenn" placeholder="məs. Riyaziyyat, Sahibkarlıq modulu...">
                    </div>
                    <div class="form-group">
                        <label>Sinif / Qrup <span>*</span></label>
                        <input type="text" id="inpSinif" placeholder="məs. 8-ci sinif, Aşpaz qrupu...">
                    </div>
                    <div class="form-group">
                        <label>KSQ/BSQ Nömrəsi</label>
                        <input type="text" id="inpParam" placeholder="məs. KSQ-1, Yarımillik BSQ">
                    </div>
                    <div class="form-group full-width">
                        <label>Əhatə olunacaq mövzular və ya alt-standartlar <span>*</span></label>
                        <textarea id="inpMovzu" placeholder="Sualların hansı mövzuları əhatə etməsini istəyirsinizsə qeyd edin..."></textarea>
                    </div>
                    <div class="form-group full-width">
                        <label>Və ya Fayl Seçin (Şəkil - PNG, JPG və ya PDF Sənədi)</label>
                        <input type="file" id="fileInp" accept="image/*,application/pdf" onchange="handleFileUpload(this)">
                    </div>
                </div>
                <button class="submit-btn" onclick="fireGemini()"><i class="fa-solid fa-print"></i> Metodiki Materialı Hazırla</button>`;
        } 
        else if(tabName === 'yoxlanma') {
            formCard.innerHTML = `
                <div class="section-tag"><i class="fa-solid fa-magnifying-glass-chart"></i> AVTOMATİK YOXLLAMA</div>
                <h2 class="section-title">Nəticələrin Yoxlanması və Analiz</h2>
                <p class="section-desc">Tələbənin cavablarını bura mətn kimi yazın və ya yazılı işin şəklini / PDF sənədini yükləyin. Sistem işi analiz edib 5 ballıq şkala ilə rəy verəcək.</p>
                <div class="form-grid">
                    <div class="form-group full-width">
                        <label>Tələbənin Yazılı Cavab Mətni</label>
                        <textarea id="inpMovzu" placeholder="Tələbənin cavabını bura yapışdıra bilərsiniz (Və ya aşağıdan şəkil / PDF yükləyin)..."></textarea>
                    </div>
                    <div class="form-group full-width">
                        <label>Və ya Fayl Seçin (Şəkil - PNG, JPG və ya PDF Sənədi)</label>
                        <input type="file" id="fileInp" accept="image/*,application/pdf" onchange="handleFileUpload(this)">
                    </div>
                </div>
                <button class="submit-btn" onclick="fireGemini()"><i class="fa-solid fa-calculator"></i> Analiz Et və Qiymətləndir</button>`;
        }
        else if(tabName === 'ai-cat') {
            formCard.innerHTML = `
                <div class="section-tag"><i class="fa-solid fa-comments"></i> AZAD REJİM</div>
                <h2 class="section-title">AI Çat (Açıq Köməkçi)</h2>
                <p class="section-desc">Menyuda xüsusi olaraq qeyd olunmayan digər bütün xidmətlər üçün bura sərbəst şəkildə sorğu yaza bilərsiniz.</p>
                <div class="form-grid">
                    <div class="form-group full-width">
                        <label>Süni İntellektə sualınız və ya tapşırığınız:</label>
                        <textarea id="inpMovzu" style="height: 120px;" placeholder="Bura istənilən mövzuda sorğu yaza bilərsiniz..."></textarea>
                    </div>
                </div>
                <button class="submit-btn" onclick="fireGemini()"><i class="fa-solid fa-paper-plane"></i> Çata Göndər</button>`;
        }
        else if(tabName === 'teqvim-tematik') {
            formCard.innerHTML = `
                <div class="section-tag"><i class="fa-solid fa-layer-group"></i> METODİKİ MODUL</div>
                <h2 class="section-title">Təqvim-tematik plan</h2>
                <p class="section-desc">Mövzu üzrə illik/rüblük mövzu ardıcıllığı planı. Şəkil və ya PDF yükləyərək proqramı avtomatik strukturlaşdıra bilərsiniz.</p>
                <div class="form-grid">
                    <div class="form-group">
                        <label>Fənn / İxtisas Modulu <span>*</span></label>
                        <input type="text" id="inpFenn" placeholder="məs. Coğrafiya, Avtoçilingər sənəti...">
                    </div>
                    <div class="form-group">
                        <label>Sinif / Auditoriya <span>*</span></label>
                        <input type="text" id="inpSinif" placeholder="məs. 10-cu sinif, Mexanik qrupu...">
                    </div>
                    <div class="form-group">
                        <label>Xüsusi Şərtlər (Müddət/Resurs)</label>
                        <input type="text" id="inpParam" placeholder="məs. 68 saatlıq, laboratoriya dərsi">
                    </div>
                    <div class="form-group full-width">
                        <label>Mövzu və ya Əsas Tələblər <span>*</span></label>
                        <textarea id="inpMovzu" placeholder="Mövzu adını daxil edin və ya aşağıdan fayl seçin..."></textarea>
                    </div>
                    <div class="form-group full-width">
                        <label>Və ya Fayl Seçin (Şəkil - PNG, JPG və ya PDF Sənədi)</label>
                        <input type="file" id="fileInp" accept="image/*,application/pdf" onchange="handleFileUpload(this)">
                    </div>
                </div>
                <button class="submit-btn" onclick="fireGemini()"><i class="fa-solid fa-gears"></i> Metodiki Materialı Hazırla</button>`;
        }
        else {
            let titleStr = "";
            let descStr = "";
            if(tabName === 'ders-plani') { titleStr = "Dərs planı"; descStr = "Gündəlik dərsin məqsədi, standartı və inteqrasiyası."; }
            if(tabName === 'ders-ssenarisi') { titleStr = "Dərs ssenarisi"; descStr = "Dərsin hər bir dəqiqəsinin addım-addım ssenari forması."; }
            if(tabName === 'interaktiv-metodlar') { titleStr = "İnteraktiv təlim metodları"; descStr = "Mövzuya uyğun qrup işləri, maraqlı analogiyalar və metodlar."; }
            if(tabName === 'ev-tapsiriqlari') { titleStr = "Ev tapşırıqları"; descStr = "Mövzunu möhkəmləndirən yaradıcı müstəqil işlər."; }

            formCard.innerHTML = `
                <div class="section-tag"><i class="fa-solid fa-layer-group"></i> METODİKİ MODUL</div>
                <h2 class="section-title">${titleStr}</h2>
                <p class="section-desc">${descStr}</p>
                <div class="form-grid">
                    <div class="form-group">
                        <label>Fənn / İxtisas Modulu <span>*</span></label>
                        <input type="text" id="inpFenn" placeholder="məs. Coğrafiya, Avtoçilingər sənəti...">
                    </div>
                    <div class="form-group">
                        <label>Sinif / Auditoriya <span>*</span></label>
                        <input type="text" id="inpSinif" placeholder="məs. 10-cu sinif, Mexanik qrupu...">
                    </div>
                    <div class="form-group">
                        <label>Xüsusi Şərtlər (Müddət/Resurs)</label>
                        <input type="text" id="inpParam" placeholder="məs. 45 dəqiqəlik, laboratoriya dərsi">
                    </div>
                    <div class="form-group full-width">
                        <label>Mövzu və ya Əsas Tələblər <span>*</span></label>
                        <textarea id="inpMovzu" placeholder="Mövzu adını daxil edin..."></textarea>
                    </div>
                    <div class="form-group full-width">
                        <label>Və ya Fayl Seçin (Şəkil - PNG, JPG və ya PDF Sənədi)</label>
                        <input type="file" id="fileInp" accept="image/*,application/pdf" onchange="handleFileUpload(this)">
                    </div>
                </div>
                <button class="submit-btn" onclick="fireGemini()"><i class="fa-solid fa-gears"></i> Metodiki Materialı Hazırla</button>`;
        }
    }

    function handleFileUpload(element) {
        const file = element.files[0];
        if (!file) return;

        base64Image = "";
        extractedPdfText = "";

        if (file.type === "application/pdf") {
            const reader = new FileReader();
            reader.onload = function(e) {
                const typedarray = new Uint8Array(e.target.result);
                pdfjsLib.getDocument(typedarray).promise.then(async function(pdf) {
                    let fullText = "";
                    for (let i = 1; i <= pdf.numPages; i++) {
                        const page = await pdf.getPage(i);
                        const textContent = await page.getTextContent();
                        const pageText = textContent.items.map(item => item.str).join(" ");
                        fullText += pageText + "\n";
                    }
                    extractedPdfText = fullText;
                    alert("PDF sənədi uğurla oxundu və mətni əlavə olundu!");
                }).catch(err => {
                    alert("PDF faylını oxuyarkən xəta baş verdi: " + err.message);
                });
            };
            reader.readAsArrayBuffer(file);
        } else if (file.type.startsWith("image/")) {
            imageMimeType = file.type || "image/jpeg";
            const reader = new FileReader();
            reader.onloadend = function() {
                base64Image = reader.result.split(',')[1];
                alert("Şəkil uğurla yükləndi!");
            }
            reader.readAsDataURL(file);
        }
    }

    async function fireGemini() {
        const apiKey = document.getElementById('apiKey').value.trim();
        const modelName = document.getElementById('modelSelect').value;
        const outputBody = document.getElementById('outputBody');
        const loader = document.getElementById('globalLoader');

        if (!apiKey) {
            alert('Zəhmət olmasa, yuxarı sağ küncdən Gemini API Key daxil edin!');
            return;
        }

        const fenn = document.getElementById('inpFenn') ? document.getElementById('inpFenn').value.trim() : '';
        const sinif = document.getElementById('inpSinif') ? document.getElementById('inpSinif').value.trim() : '';
        let movzu = document.getElementById('inpMovzu') ? document.getElementById('inpMovzu').value.trim() : '';
        const param = document.getElementById('inpParam') ? document.getElementById('inpParam').value.trim() : '';

        if (extractedPdfText) {
            movzu += "\n\n[YÜKLƏNMİŞ PDF-DƏN ÇIXARILAN MƏTN]:\n" + extractedPdfText;
        }

        if(!movzu && !base64Image) {
            alert('Mətn daxil edin və ya fayl (şəkil/PDF) yükləyin!');
            return;
        }

        loader.style.display = 'flex';
        outputBody.innerHTML = '<p style="color:#2563eb;"><i class="fa-solid fa-spinner fa-spin"></i> Sistem rəsmi təhsil və vizual emal bazası ilə əlaqə qurur...</p>';

        let systemInstructions = `Sən müəllimlər üçün peşəkar, Azərbaycan Respublikasının təhsil və kurikulum standartlarına mükəmməl bələd olan rəsmi Aİ asistansan.
MÜTLƏQ QAYDA: Riyazi simvollar, kəsrlər və kökaltı ifadələr üçün Heç bir halda LaTeX ($...$, \\frac, \\sqrt) İSTİFADƏ ETMƏ. Bütün riyazi düsturları, kökləri və kəsrləri standart Unicode simvolları ilə yaz (məsələn: √10, x², ≤, ±).`;

        if (activeTab === 'pptx-generator') {
            systemInstructions += `\n\nSənin vəzifən təqdimat slaydları hazırlamaqdır. Mətn cavabında həm aydın Markdown formatında slaydları göstər, həm də cavabın ən sonuna mütləq aşağıdakı JSON bloqunu daxil et:
\`\`\`json
{
  "slides": [
    {"title": "Slayd 1 Başlığı", "bullets": ["Bənd 1", "Bənd 2", "Bənd 3"]},
    {"title": "Slayd 2 Başlığı", "bullets": ["Bənd 1", "Bənd 2"]}
  ]
}
\`\`\``;
        } else if (activeTab === 'diaqram-sxem') {
            systemInstructions += `\n\nSənin vəzifən dərsin mövzusunu Mermaid.js sxeminə çevirməkdir. Mütləq cavabın içində təmiz \`\`\`mermaid bloqu daxil et (məsələn: graph TD və ya flowchart LR).`;
        } else if (activeTab === 'ai-sekil') {
            systemInstructions += `\n\nSən vizual illüstrasiya mütəxəssisisən. Daxil edilən təsvir əsasında yüksək keyfiyyətli, tədris yönümlü ingiliscə dəqiq şəkil promptu yaz və bunu istifadəçiyə izah et. Cavabın sonuna mütləq bu formatda ingiliscə prompt yerləşdir: [PROMPT: your translated English prompt here]`;
        } else if (activeTab === 'ksq' || activeTab === 'bsq') {
            const novu = activeTab === 'ksq' ? "Kiçik Summativ Qiymətləndirmə (KSQ)" : "Böyük Summativ Qiymətləndirmə (BSQ)";
            systemInstructions += `\n\nSənin əsas vəzifən ${novu} sənədi tərtib etməkdir.
MƏTLƏBƏN TƏLƏBLƏR:
1. Tam rəsmi, məktəb/mərkəz blankı formatında sənəd başlığı qur.
2. Dəqiq olaraq tam 20 unikal sual hazırla.
3. Suallar mütləq asandan çətinə doğru sıralanmalıdır (1-5 çox asan, 6-12 orta, 13-17 çətin, 18-20 mürəkkəb).
4. Sənədin sonuna qiymətləndirmə şkalasını standart Markdown cədvəli kimi əlavə et.
5. Sənədin ən sonuna müəllim üçün Düzgün Cavablar Açarı əlavə et.`;
        } else if (activeTab === 'yoxlanma') {
            systemInstructions += `\n\nSən tələbə nəticələrini yoxlayan ekspertsən. Daxil edilən mətni, PDF faylını və ya şəkildəki yazıları oxu. Onu analiz et, səhvlərini bənd-bənd göstər və 5 ballıq şkala ilə qiymətləndirib rəy yaz.`;
        } else if (activeTab === 'ai-cat') {
            systemInstructions += `\n\nSən azad AI Çatsan. Fərqli təhsil suallarını və pedaqoji məsələləri rəsmi üslubda cavablandır.`;
        } else {
            systemInstructions += `\n\nFunksiya: [${activeTab}]. Verilən fənn, sinif və mövzu üzrə strukturlu metodiki vəsait materialı tərtib et.`;
        }

        let userPrompt = `Fənn: ${fenn}\nSinif/Qrup: ${sinif}\nMövzu/Tələb: ${movzu}\nƏlavə parametr: ${param}`;
        
        let requestBody = {
            contents: [{
                parts: [{ text: systemInstructions + "\n\nİstifadəçi sorğusu:\n" + userPrompt }]
            }]
        };

        if (base64Image) {
            requestBody.contents[0].parts.push({
                inlineData: { mimeType: imageMimeType, data: base64Image }
            });
        }

        try {
            const url = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;
            
            const response = await fetch(url, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(requestBody)
            });

            const data = await response.json();
            
            if (data.error) {
                outputBody.innerHTML = `<div style="color:#ef4444; background:#fee2e2; padding:15px; border-radius:8px;"><i class="fa-solid fa-triangle-exclamation"></i> API Xətası: ${data.error.message}</div>`;
                return;
            }

            if (data.candidates && data.candidates[0].content.parts[0].text) {
                let markdownText = data.candidates[0].content.parts[0].text;
                
                // PPTX JSON parse
                if (activeTab === 'pptx-generator') {
                    const jsonMatch = markdownText.match(/```json\s*([\s\S]*?)\s*```/);
                    if (jsonMatch) {
                        try {
                            const parsed = JSON.parse(jsonMatch[1]);
                            generatedPptxSlides = parsed.slides || [];
                        } catch(e) { console.error("PPTX JSON xətası:", e); }
                    }
                }

                // AI Şəkil generasiyası (Pollinations API)
                if (activeTab === 'ai-sekil') {
                    const promptMatch = markdownText.match(/\[PROMPT:\s*(.*?)\]/i);
                    let imageUrl = "";
                    if (promptMatch && promptMatch[1]) {
                        const cleanPrompt = encodeURIComponent(promptMatch[1].trim());
                        imageUrl = `https://pollinations.ai/p/${cleanPrompt}?width=1024&height=768&seed=42&model=flux`;
                        markdownText += `\n\n### 🖼️ Generasiya olunan Vizual İllüstrasiya:\n![Əyani Material](${imageUrl})`;
                    }
                }

                outputBody.innerHTML = marked.parse(markdownText);

                // Mermaid diaqramlarının emalı
                if (activeTab === 'diaqram-sxem' || markdownText.includes('```mermaid')) {
                    const mermaidElements = outputBody.querySelectorAll('.language-mermaid');
                    mermaidElements.forEach((el, index) => {
                        const code = el.innerText;
                        const div = document.createElement('div');
                        div.className = 'mermaid';
                        div.innerHTML = code;
                        el.parentNode.replaceChild(div, el);
                    });
                    mermaid.run();
                }

                document.getElementById('resultPanelTitle').innerHTML = `<i class="fa-solid fa-circle-check" style="color:#10b981;"></i> İş tamamlandı!`;
            } else {
                outputBody.innerText = 'Cavab strukturunda xəta baş verdi və ya boş cavab gəldi.';
            }
        } catch (error) {
            outputBody.innerHTML = `<div style="color:#ef4444; background:#fee2e2; padding:15px; border-radius:8px;"><i class="fa-solid fa-bug"></i> Qoşulma xətası baş verdi: ${error.message}</div>`;
        } finally {
            loader.style.display = 'none';
        }
    }

    function copyResult() {
        const bodyText = document.getElementById('outputBody').innerText;
        if(bodyText.includes('Menyulardan birini seçin') || bodyText.trim() === "") return;
        
        navigator.clipboard.writeText(bodyText).then(() => {
            alert('Mətn bütün simvolları və xüsusi işarələri qoruyaraq kopyalandı!');
        }).catch(err => {
            alert('Kopyalama zamanı gözlənilməz xəta oldu.');
        });
    }

    function exportToWord() {
        const bodyHtml = document.getElementById('outputBody').innerHTML;
        const bodyText = document.getElementById('outputBody').innerText;
        
        if(bodyText.includes('Menyulardan birini seçin') || bodyText.trim() === "") {
            alert('Yükləmək üçün hər hansı bir material generasiya olunmayıb!');
            return;
        }

        const header = "<html xmlns:o='urn:schemas-microsoft-com:office:office' "+
            "xmlns:w='urn:schemas-microsoft-com:office:word' "+
            "xmlns='[http://www.w3.org/TR/REC-html40](http://www.w3.org/TR/REC-html40)'>"+
            "<head><meta charset='utf-8'><title>Metodiki Material</title>"+
            "<style>body {font-family:'Segoe UI', Arial, sans-serif; line-height: 1.6;} table {border-collapse:collapse; width:100%;} th, td {border:1px solid #000000; padding:6px; text-align:left;}</style></head><body>";
        const footer = "</body></html>";
        
        const sourceHtml = header + bodyHtml + footer;
        const blob = new Blob(['\ufeff' + sourceHtml], { type: 'application/msword;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        
        let fileName = activeTab.toUpperCase() + "_Materiali.doc";
        a.download = fileName;
        
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    function exportToPowerPoint() {
        if (!generatedPptxSlides || generatedPptxSlides.length === 0) {
            alert("PowerPoint slaydları yaradılmayıb. Zəhmət olmasa sol menyudan 'PPTX Təqdimat' modulu ilə material hazırlayın.");
            return;
        }

        let pptx = new PptxGenJS();
        
        // Titul Slaydı
        let titleSlide = pptx.addSlide();
        titleSlide.background = { color: "1E293B" };
        titleSlide.addText("Müəllimin Aİ Köməkçisi", { x: 0.5, y: 1.5, w: "90%", fontSize: 32, color: "38BDF8", bold: true, align: "center" });
        titleSlide.addText("Avtomatik Tədris Təqdimatı", { x: 0.5, y: 2.5, w: "90%", fontSize: 18, color: "94A3B8", align: "center" });

        // Əsas Slaydlar
        generatedPptxSlides.forEach((slideData) => {
            let slide = pptx.addSlide();
            
            // Başlıq
            slide.addText(slideData.title || "Slayd", { 
                x: 0.5, y: 0.5, w: "90%", fontSize: 22, color: "1E3A8A", bold: true, border: { pt: "0 0 2 0", color: "2563EB" } 
            });

            // Mətn Bəndləri
            if (slideData.bullets && slideData.bullets.length > 0) {
                let bulletText = slideData.bullets.map(b => ({ text: b, options: { fontSize: 15, color: "0F172A", breakLine: true, bullet: true } }));
                slide.addText(bulletText, { x: 0.8, y: 1.5, w: "85%", h: 4.5, lineSpacing: 28 });
            }
        });

        pptx.writeFile({ fileName: "Tədris_Təqdimatı.pptx" });
    }
</script>
</body>
</html>
