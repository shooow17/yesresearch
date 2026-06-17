document.addEventListener('DOMContentLoaded', () => {
    const chapterList = document.getElementById('chapterList');
    const pdfViewer = document.getElementById('pdfViewer');
    const pdfMobileViewer = document.getElementById('pdfMobileViewer');
    const welcomeScreen = document.getElementById('welcomeScreen');
    const currentDocTitle = document.getElementById('currentDocTitle');
    const downloadBtn = document.getElementById('downloadBtn');
    const searchInput = document.getElementById('searchInput');
    const courseTag = document.getElementById('courseTag');

    let allDocs = [];

    const categoryNames = {
        '0': '基本與封面',
        '1': '學校現況分析',
        '2': '願景與目標',
        '3': '總體課程架構',
        '4': '課程實施說明',
        '5': '領域課程計畫',
        '6': '彈性學習課程',
        '7': '發展與評鑑'
    };

    if (typeof siteData !== 'undefined') {
        allDocs = siteData;
        document.querySelector('.chapter-meta').textContent = `總計 ${allDocs.length} 份文件`;
        renderChapters(allDocs);
    } else {
        chapterList.innerHTML = `<div class="loading" style="color: #ef4444; padding: 24px; text-align: center;">資料載入失敗。請確認 data.js 是否存在。</div>`;
    }

    function renderChapters(dataList) {
        chapterList.innerHTML = '';
        
        if (dataList.length === 0) {
            chapterList.innerHTML = `<div class="loading" style="padding: 24px; text-align: center;">找不到相符的文件</div>`;
            return;
        }

        // Group by first digit
        const groups = {};
        dataList.forEach(item => {
            const match = item.name.match(/^(\d)/);
            const cat = match ? match[1] : '其他';
            if (!groups[cat]) groups[cat] = [];
            groups[cat].push(item);
        });

        // Sort categories logically
        const sortedCategories = Object.keys(groups).sort((a, b) => {
            if (a === '其他') return 1;
            if (b === '其他') return -1;
            return parseInt(a) - parseInt(b);
        });
        
        sortedCategories.forEach(cat => {
            // Render Category Header
            const header = document.createElement('div');
            header.className = 'category-header';
            
            const catTitle = categoryNames[cat] || '其他文件';
            
            header.innerHTML = `
                <span class="cat-num">${cat}</span>
                <span class="cat-title">${catTitle}</span>
            `;
            chapterList.appendChild(header);

            // Render items in category
            groups[cat].forEach((item) => {
                const node = document.createElement('div');
                node.className = 'chapter-item';
                node.title = item.name;
                
                node.innerHTML = `
                    <div class="chapter-icon">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline></svg>
                    </div>
                    <div class="chapter-title">${item.name}</div>
                `;
                
                node.addEventListener('click', () => {
                    document.querySelectorAll('.chapter-item.active').forEach(el => el.classList.remove('active'));
                    node.classList.add('active');
                    openPdf(item);
                });
                
                chapterList.appendChild(node);
            });
        });
    }

    // Setup PDF.js worker
    if (window.pdfjsLib) {
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    }

    // Detect mobile and tablets, including iPads pretending to be Macs
    const isIOS = /iPhone|iPad|iPod/i.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
    const isMobile = isIOS || /Android/i.test(navigator.userAgent) || window.innerWidth <= 992;

    function openPdf(item) {
        welcomeScreen.classList.add('hidden');
        courseTag.classList.remove('hidden');
        
        currentDocTitle.textContent = item.name;
        currentDocTitle.title = item.name;
        
        downloadBtn.href = item.pdf;
        downloadBtn.download = `${item.name}.pdf`;
        downloadBtn.classList.remove('disabled');

        if (isMobile && window.pdfjsLib) {
            pdfViewer.classList.add('hidden');
            pdfMobileViewer.classList.remove('hidden');
            renderPdfWithPdfJs(item.pdf);
        } else {
            pdfMobileViewer.classList.add('hidden');
            pdfViewer.classList.remove('hidden');
            // Use view=FitH to force the PDF to fill the width, reducing black bars
            pdfViewer.src = item.pdf + '#toolbar=0&navpanes=0&view=FitH';
        }
    }

    async function renderPdfWithPdfJs(url) {
        pdfMobileViewer.innerHTML = '<div class="loading" style="padding: 40px; text-align: center;">載入文件中...</div>';
        
        try {
            const encodedUrl = encodeURI(url);
            const loadingTask = pdfjsLib.getDocument(encodedUrl);
            const pdf = await loadingTask.promise;
            
            pdfMobileViewer.innerHTML = ''; // Clear loading
            
            for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
                const page = await pdf.getPage(pageNum);
                
                // Get the viewport for scaling
                const containerWidth = pdfMobileViewer.clientWidth || window.innerWidth - 32;
                const unscaledViewport = page.getViewport({ scale: 1.0 });
                // Calculate scale to fit width (with some padding)
                const scale = (containerWidth - 32) / unscaledViewport.width; 
                const viewport = page.getViewport({ scale: Math.max(scale, 1.0) });
                
                const canvas = document.createElement('canvas');
                const ctx = canvas.getContext('2d');
                
                canvas.width = viewport.width;
                canvas.height = viewport.height;
                pdfMobileViewer.appendChild(canvas);
                
                const renderContext = {
                    canvasContext: ctx,
                    viewport: viewport
                };
                
                await page.render(renderContext).promise;
            }
        } catch (error) {
            console.error('Error rendering PDF:', error);
            pdfMobileViewer.innerHTML = '<div class="loading" style="padding: 40px; text-align: center; color: #ef4444;">PDF 載入失敗，請嘗試直接下載。</div>';
        }
    }

    // Search functionality
    searchInput.addEventListener('input', (e) => {
        const term = e.target.value.toLowerCase();
        const filteredDocs = allDocs.filter(doc => doc.name.toLowerCase().includes(term));
        renderChapters(filteredDocs);
    });
});
