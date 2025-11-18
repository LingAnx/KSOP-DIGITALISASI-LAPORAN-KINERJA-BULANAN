// Google Apps Script Deployment URL
        const GOOGLE_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyY5sVpo-ukvZNZN1eCvruGDJwxtw71KrC1cFn-f-Z8hVv7N1atgwoq0O_sykjkb9iWhw/exec";

        let allData = [];
        let currentPage = 1;  // ← TAMBAH INI
        const itemsPerPage = 10;  // ← TAMBAH INI
        let dataPendukung = [];

        // Handle Login
        function handleLogin(event) {
            event.preventDefault();
            const username = document.getElementById('username').value;
            const password = document.getElementById('password').value;

            // Simple demo authentication
            if (username === 'admin' && password === '123456') {
                localStorage.setItem('loggedIn', 'true');
                localStorage.setItem('username', username);
                showPage('dashboard');
                updateNav();
            } else {
                alert('Username atau password salah!');
            }
        }

        // Handle Logout
        function logout() {
            if (confirm('Apakah Anda yakin ingin logout?')) {
                localStorage.removeItem('loggedIn');
                localStorage.removeItem('username');
                document.getElementById('username').value = '';
                document.getElementById('password').value = '';
                showPage('login');
                updateNav();
            }
        }

        // Show/Hide Pages
        function showPage(page) {
            // Sembunyikan semua halaman
            document.getElementById('loginPage').style.display = 'none';
            document.getElementById('dashboardPage').style.display = 'none';
            document.getElementById('formPage').style.display = 'none';
            document.getElementById('graphicPage').style.display = 'none';
            document.getElementById('dataPendukungPage').style.display = 'none'; // pastikan ini ada

            // Tampilkan halaman sesuai argumen
            if (page === 'login') {
                document.getElementById('loginPage').style.display = 'flex';
            } else if (page === 'dashboard') {
                document.getElementById('dashboardPage').style.display = 'block';
            } else if (page === 'form') {
                document.getElementById('formPage').style.display = 'block';
            } else if (page === 'graphic') {
                document.getElementById('graphicPage').style.display = 'block';
                loadDataFromSheet();
            } else if (page === 'dataPendukung') {
                document.getElementById('dataPendukungPage').style.display = 'block';
                populatePilihSeksi();
                populatePilihLaporan();
                displayDataPendukung();
            }
        }

        // Update Navigation
        function updateNav() {
            const isLoggedIn = localStorage.getItem('loggedIn') === 'true';
            document.getElementById('navDash').style.display = isLoggedIn ? 'inline-block' : 'none';
            document.getElementById('navForm').style.display = isLoggedIn ? 'inline-block' : 'none';
            document.getElementById('navGraphic').style.display = isLoggedIn ? 'inline-block' : 'none';
            document.getElementById('navPendukung').style.display = isLoggedIn ? 'inline-block' : 'none';
            document.getElementById('navLogout').style.display = isLoggedIn ? 'inline-block' : 'none';
        }

        // Handle Form Submit
        function handleFormSubmit(event) {
            event.preventDefault();

            const formData = {
                tanggal: new Date().toLocaleDateString('id-ID'),
                bulan: document.getElementById('bulan').value,
                tahun: document.getElementById('tahun').value,
                seksi: document.getElementById('seksi').value,
                indikator: document.getElementById('indikator').value,
                target: parseFloat(document.getElementById('target').value),
                realisasi: parseFloat(document.getElementById('realisasi').value),
                catatan: document.getElementById('catatan').value
            };

            console.log('Form submitted with data:', formData);

            // ✅ JANGAN overwrite, TAMBAHKAN ke array yang sudah ada
            let savedData = JSON.parse(localStorage.getItem('ksopData') || '[]');
            console.log('Current savedData count before add:', savedData.length);
            
            savedData.push(formData);
            localStorage.setItem('ksopData', JSON.stringify(savedData));
            
            console.log('Saved to localStorage, total count:', savedData.length);

            // Kirim ke Google Apps Script
            sendToGoogleSheet(formData);
        }

        // Send to Google Sheet
        async function sendToGoogleSheet(data) {
            try {
                const response = await fetch(GOOGLE_APPS_SCRIPT_URL, {
                    method: 'POST',
                    body: JSON.stringify(data),
                    headers: {
                        'Content-Type': 'application/json'
                    }
                });
                if (!response.ok) {
                    const txt = await response.text();
                    throw new Error('HTTP' + response.status + ':' + txt);
                }
                const result = await response.json().catch(() => { throw new Error('Response is not valid JSON'); });
                const alertDiv = document.getElementById('formAlert');
                if (result.status === 'success') {
                    alertDiv.className = 'alert success';
                    alertDiv.textContent = '✅ Data berhasil disimpan ke spreadsheet!';
                    alertDiv.style.display = 'block';
                    document.getElementById('laporanForm').reset();
                    setTimeout(() => {
                        alertDiv.style.display = 'none';
                    }, 3000);
                } else {
                    alertDiv.className = 'alert error';
                    alertDiv.textContent = '❌ Gagal menyimpan data. Periksa URL Google Apps Script.';
                    alertDiv.style.display = 'block';
                }
            } catch (error) {
                console.error('Error:', error);
                const alertDiv = document.getElementById('formAlert');
                alertDiv.className = 'alert error';
                alertDiv.textContent = '⚠️ Terjadi kesalahan koneksi. Data disimpan ke lokal.';
                alertDiv.style.display = 'block';
            }
        }

        // Load Data from Google Sheet — jangan auto-refresh
        async function loadDataFromSheet() {
            console.log('loadDataFromSheet() called');
            
            // 1) Load dari localStorage terlebih dahulu
            try {
                const localRaw = JSON.parse(localStorage.getItem('ksopData') || '[]');
                console.log('Local storage data (count):', localRaw.length);
                allData = Array.isArray(localRaw) ? localRaw.slice() : [];
            } catch (e) {
                console.error('Error parsing localStorage:', e);
                allData = [];
            }

            console.log('After local load, allData count:', allData.length);
            
            populateFilters();
            displayData();

            // 2) JANGAN coba load dari Google Sheet otomatis — hanya jika user minta
            // (Hapus blok try-catch Google Sheets yang lama)
        }

        // Tambah fungsi terpisah untuk manual sync dari Google Sheets (optional)
        async function syncFromGoogleSheets() {
            console.log('syncFromGoogleSheets() called');
            try {
                const response = await fetch(GOOGLE_APPS_SCRIPT_URL);
                if (!response.ok) throw new Error('HTTP ' + response.status);
                const result = await response.json().catch(() => { throw new Error('Respon bukan JSON'); });
                
                if (result.status === 'success' && result.data && Array.isArray(result.data)) {
                    console.log('Dari Google Sheets, data count:', result.data.length);
                    const newData = result.data.slice(1).map(row => ({
                        tanggal: row[0] || '',
                        seksi: String(row[1] || '').trim(),
                        bulan: String(row[2] || '').trim(),
                        tahun: String(row[3] || '').trim(),
                        indikator: row[4] || '',
                        target: row[5] || '',
                        realisasi: row[6] || '',
                        catatan: row[7] || ''
                    }));
                    allData = newData;
                    localStorage.setItem('ksopData', JSON.stringify(newData));
                    populateFilters();
                    displayData();
                    alert('✅ Data berhasil di-sync dari Google Sheets!');
                }
            } catch (err) {
                console.error('Error sync:', err);
                alert('❌ Gagal sync dari Google Sheets. Gunakan data lokal.');
            }
        }

        // Populate filter selects
        function populateFilters() {
            console.log('=== populateFilters() START ===');
            console.log('allData length:', (allData || []).length);
            
            const bulanSelect = document.getElementById('filterBulan');
            const seksiSelect = document.getElementById('filterSeksi');
            const tahunSelect = document.getElementById('filterTahun');
            
            if (!bulanSelect || !seksiSelect || !tahunSelect) return;

            const bulanSet = new Set();
            const seksiSet = new Set();
            const tahunSet = new Set();

            (allData || []).forEach((row, idx) => {
                const bulan = String(row.bulan || row[2] || '').trim();
                const seksi = String(row.seksi || row[1] || '').trim();
                const tahun = String(row.tahun || row[3] || '').trim();

                if (bulan) bulanSet.add(bulan);
                if (seksi) seksiSet.add(seksi);
                if (tahun) tahunSet.add(tahun);
            });

            console.log('✅ Unique bulan count:', bulanSet.size, 'values:', Array.from(bulanSet));
            console.log('✅ Unique seksi count:', seksiSet.size, 'values:', Array.from(seksiSet));
            console.log('✅ Unique tahun count:', tahunSet.size, 'values:', Array.from(tahunSet));

            function buildOptions(selectEl, setValues, defaultLabel) {
                selectEl.innerHTML = '';
                const optAll = document.createElement('option');
                optAll.value = '';
                optAll.textContent = defaultLabel;
                selectEl.appendChild(optAll);

                Array.from(setValues).sort().forEach(v => {
                    const opt = document.createElement('option');
                    opt.value = String(v);
                    opt.textContent = v;
                    selectEl.appendChild(opt);
                });
            }

            buildOptions(bulanSelect, bulanSet, 'Semua Bulan');
            buildOptions(seksiSelect, seksiSet, 'Semua Seksi');
            buildOptions(tahunSelect, tahunSet, 'Semua Tahun');
            
            console.log('=== populateFilters() END ===');
        }

        // Display Data in Table with Pagination
        function displayData() {
            console.log('displayData() called');
            
            const filterBulanEl = document.getElementById('filterBulan');
            const filterSeksiEl = document.getElementById('filterSeksi');
            const filterTahunEl = document.getElementById('filterTahun');
            const tableBody = document.getElementById('dataTableBody');
            const dataTable = document.getElementById('dataTable');
            const paginationDiv = document.getElementById('paginationControls');

            // Jika elemen tidak ada, jangan lanjut
            if (!filterBulanEl || !filterSeksiEl || !filterTahunEl || !tableBody || !dataTable || !paginationDiv) {
                console.error('❌ Salah satu elemen HTML tidak ditemukan');
                return;
            }

            const filterBulan = filterBulanEl.value || '';
            const filterSeksi = filterSeksiEl.value || '';
            const filterTahun = filterTahunEl.value || '';

            console.log('Active filters:', { filterBulan, filterSeksi, filterTahun });

            let filteredData = (allData || []).slice();

            if (filterBulan || filterSeksi || filterTahun) {
                filteredData = filteredData.filter(row => {
                    const bulan = String(row.bulan || row[2] || '').trim();
                    const seksi = String(row.seksi || row[1] || '').trim();
                    const tahun = String(row.tahun || row[3] || '').trim();

                    const matchBulan = filterBulan ? bulan === filterBulan : true;
                    const matchSeksi = filterSeksi ? seksi === filterSeksi : true;
                    const matchTahun = filterTahun ? tahun === filterTahun : true;
                    return matchBulan && matchSeksi && matchTahun;
                });
            }

            console.log('Filtered data count:', filteredData.length);

            tableBody.innerHTML = '';

            if (!filteredData || filteredData.length === 0) {
                dataTable.style.display = 'none';
                paginationDiv.innerHTML = '';
                renderChart([]);
                return;
            }

            // Reset ke halaman 1 saat filter berubah
            currentPage = 1;

            // Hitung total halaman
            const totalPages = Math.ceil(filteredData.length / itemsPerPage);
            const startIdx = (currentPage - 1) * itemsPerPage;
            const endIdx = startIdx + itemsPerPage;
            const pageData = filteredData.slice(startIdx, endIdx);

            console.log('Page data count:', pageData.length, 'Total pages:', totalPages);

            // Render baris tabel untuk halaman saat ini
            pageData.forEach(row => {
                const r = Array.isArray(row)
                    ? { tanggal: row[0], seksi: row[1], bulan: row[2], tahun: row[3], indikator: row[4], target: row[5], realisasi: row[6], catatan: row[7] }
                    : row;

                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>${r.tanggal || '-'}</td>
                    <td>${r.seksi || '-'}</td>
                    <td>${r.bulan || '-'}</td>
                    <td>${r.tahun || '-'}</td>
                    <td>${r.indikator || '-'}</td>
                    <td>${r.target || '-'}</td>
                    <td>${r.realisasi || '-'}</td>
                    <td>${r.catatan || '-'}</td>
                `;
                tableBody.appendChild(tr);
            });

            dataTable.style.display = 'table';

            // Render pagination controls
            renderPagination(totalPages, filteredData);

            // Render grafik berdasarkan data yang tampil di tabel
            renderChart(filteredData);
        }

        // Render Pagination Controls
        function renderPagination(totalPages, filteredData) {
            const paginationDiv = document.getElementById('paginationControls');
            if (!paginationDiv) return;

            paginationDiv.innerHTML = '';

            if (totalPages <= 1) return; // Jika hanya 1 halaman, jangan tampilkan pagination

            const container = document.createElement('div');
            container.style.cssText = 'display: flex; justify-content: center; gap: 8px; margin-top: 16px; flex-wrap: wrap;';

            // Tombol Previous
            const prevBtn = document.createElement('button');
            prevBtn.textContent = '← Sebelumnya';
            prevBtn.disabled = currentPage === 1;
            prevBtn.style.cssText = 'padding: 8px 16px; cursor: ' + (currentPage === 1 ? 'not-allowed' : 'pointer') + '; opacity: ' + (currentPage === 1 ? '0.5' : '1') + ';';
            prevBtn.onclick = () => {
                if (currentPage > 1) {
                    currentPage--;
                    displayData();
                }
            };
            container.appendChild(prevBtn);

            // Tombol halaman
            for (let i = 1; i <= totalPages; i++) {
                const pageBtn = document.createElement('button');
                pageBtn.textContent = i;
                pageBtn.style.cssText = 'padding: 8px 12px; cursor: pointer;' + (i === currentPage ? ' background-color: #0056b3; color: white; font-weight: bold;' : ' background-color: #e0e0e0;');
                pageBtn.onclick = () => {
                    currentPage = i;
                    displayData();
                };
                container.appendChild(pageBtn);
            }

            // Tombol Next
            const nextBtn = document.createElement('button');
            nextBtn.textContent = 'Selanjutnya →';
            nextBtn.disabled = currentPage === totalPages;
            nextBtn.style.cssText = 'padding: 8px 16px; cursor: ' + (currentPage === totalPages ? 'not-allowed' : 'pointer') + '; opacity: ' + (currentPage === totalPages ? '0.5' : '1') + ';';
            nextBtn.onclick = () => {
                if (currentPage < totalPages) {
                    currentPage++;
                    displayData();
                }
            };
            container.appendChild(nextBtn);

            // Info halaman
            const infoSpan = document.createElement('span');
            infoSpan.textContent = `Halaman ${currentPage} dari ${totalPages}`;
            infoSpan.style.cssText = 'padding: 8px 16px; align-self: center; font-weight: bold;';
            container.appendChild(infoSpan);

            paginationDiv.appendChild(container);
        }

        // Attach change listeners to filters — SEKALI saja
        function attachFilterListeners() {
            console.log('attachFilterListeners() called');
            ['filterBulan', 'filterSeksi', 'filterTahun'].forEach(id => {
                const el = document.getElementById(id);
                if (el) {
                    // Hapus listener lama (jika ada)
                    el.removeEventListener('change', displayData);
                    // Tambah listener baru
                    el.addEventListener('change', () => {
                        console.log(`Filter ${id} changed to:`, el.value);
                        currentPage = 1; // reset ke halaman 1
                        displayData();
                    });
                } else {
                    console.warn(`⚠️ Element #${id} tidak ditemukan`);
                }
            });
        }

        // --- Chart functions (Chart.js loader + renderer) ---
        function ensureChartJs() {
            return new Promise((resolve, reject) => {
                if (window.Chart) return resolve(window.Chart);
                const s = document.createElement('script');
                s.src = 'https://cdn.jsdelivr.net/npm/chart.js';
                s.onload = () => resolve(window.Chart);
                s.onerror = () => reject(new Error('Gagal memuat Chart.js'));
                document.head.appendChild(s);
            });
        }

        function normalizeRow(row) {
            if (!row) return null;
            if (Array.isArray(row)) {
                // asumsi urutan spreadsheet: [tanggal, seksi, bulan, tahun, indikator, target, realisasi, catatan]
                return {
                    tanggal: row[0] || '',
                    seksi: String(row[1] || '').trim(),
                    bulan: String(row[2] || '').trim(),
                    tahun: row[3] || '',
                    indikator: String(row[4] || '').trim(),
                    target: Number(row[5] || 0),
                    realisasi: Number(row[6] || 0),
                    catatan: row[7] || ''
                };
            } else {
                return {
                    tanggal: row.tanggal || '',
                    seksi: String(row.seksi || '').trim(),
                    bulan: String(row.bulan || '').trim(),
                    tahun: row.tahun || '',
                    indikator: String(row.indikator || '').trim(),
                    target: Number(row.target || 0),
                    realisasi: Number(row.realisasi || 0),
                    catatan: row.catatan || ''
                };
            }
        }

        function prepareChartData(dataArray) {
            const map = new Map(); // key: indikator, value: { target, realisasi }
            (dataArray || []).forEach(r => {
                const n = normalizeRow(r);
                if (!n || !n.indikator) return;
                const key = n.indikator;
                if (!map.has(key)) map.set(key, { target: 0, realisasi: 0 });
                const agg = map.get(key);
                agg.target += isFinite(n.target) ? n.target : 0;
                agg.realisasi += isFinite(n.realisasi) ? n.realisasi : 0;
            });
            const labels = Array.from(map.keys());
            const targets = labels.map(l => map.get(l).target);
            const realisasis = labels.map(l => map.get(l).realisasi);
            return { labels, targets, realisasis };
        }

        // renderChart menerima data (filtered) atau gunakan allData
        function renderChart(dataParam) {
            const dataForChart = Array.isArray(dataParam) ? dataParam : (allData || []);
            const prepared = prepareChartData(dataForChart);

            const chartContainer = document.getElementById('chartContainer');
            if (!chartContainer) return;

            if (!prepared.labels.length) {
                chartContainer.innerHTML = '<div class="no-data">📊 Tidak ada data untuk grafik.</div>';
                if (window._ksopChart && window._ksopChart.destroy) window._ksopChart.destroy();
                return;
            }

            // ensure canvas
            let canvas = document.getElementById('chartCanvas');
            if (!canvas) {
                chartContainer.innerHTML = '<canvas id="chartCanvas"></canvas>';
                canvas = document.getElementById('chartCanvas');
            }

            ensureChartJs()
                .then(() => {
                    const ctx = canvas.getContext('2d');
                    if (window._ksopChart && window._ksopChart.destroy) window._ksopChart.destroy();

                    window._ksopChart = new Chart(ctx, {
                        type: 'bar',
                        data: {
                            labels: prepared.labels,
                            datasets: [
                                { label: 'Target', data: prepared.targets, backgroundColor: 'rgba(54,162,235,0.5)', borderColor: 'rgba(54,162,235,1)', borderWidth: 1 },
                                { label: 'Realisasi', data: prepared.realisasis, backgroundColor: 'rgba(75,192,192,0.6)', borderColor: 'rgba(75,192,192,1)', borderWidth: 1 }
                            ]
                        },
                        options: {
                            responsive: true,
                            maintainAspectRatio: true,
                            indexAxis: 'y',
                            scales: { 
                                x: { beginAtZero: true }
                            },
                            plugins: { 
                                legend: { position: 'top' },
                                tooltip: { mode: 'index', intersect: false }
                            }
                        }
                    });
                })
                .catch(err => {
                    console.error('Chart load error:', err);
                    chartContainer.innerHTML = '<div class="no-data">⚠️ Gagal memuat grafik.</div>';
                });
        }

        // Fungsi untuk export data ke Google Sheets
        function exportToGoogleSheets() {
            if (!allData || allData.length === 0) {
                alert('❌ Tidak ada data untuk diekspor');
                return;
            }

            // Siapkan data dalam format spreadsheet (array of arrays)
            const sheetData = [
                ['Tanggal', 'Seksi', 'Bulan', 'Tahun', 'Indikator', 'Target', 'Realisasi', 'Catatan']
            ];

            allData.forEach(row => {
                const r = Array.isArray(row)
                    ? { tanggal: row[0], seksi: row[1], bulan: row[2], tahun: row[3], indikator: row[4], target: row[5], realisasi: row[6], catatan: row[7] }
                    : row;

                sheetData.push([
                    r.tanggal || '',
                    r.seksi || '',
                    r.bulan || '',
                    r.tahun || '',
                    r.indikator || '',
                    r.target || '',
                    r.realisasi || '',
                    r.catatan || ''
                ]);
            });

            // Buat CSV string
            const csv = sheetData.map(row => 
                row.map(cell => `"${String(cell || '').replace(/"/g, '""')}"`).join(',')
            ).join('\n');

            // Download sebagai file CSV
            const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement('a');
            const url = URL.createObjectURL(blob);
            link.setAttribute('href', url);
            link.setAttribute('download', `KSOP_Kinerja_${new Date().toLocaleDateString('id-ID').replace(/\//g, '-')}.csv`);
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);

            alert('✅ Data berhasil diunduh sebagai file CSV');
        }

        // Fungsi untuk print tabel
        function printTable() {
            if (!allData || allData.length === 0) {
                alert('❌ Tidak ada data untuk dicetak');
                return;
            }

            const printWindow = window.open('', '', 'height=600,width=900');
            const tableHTML = document.getElementById('dataTable').outerHTML;

            const htmlContent = `
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="UTF-8">
                    <title>KSOP Kelas II Cirebon - Data Kinerja</title>
                    <style>
                        body { font-family: Arial, sans-serif; margin: 20px; }
                        h1 { text-align: center; color: #0056b3; }
                        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
                        th, td { border: 1px solid #333; padding: 10px; text-align: left; }
                        th { background-color: #0056b3; color: white; font-weight: bold; }
                        tr:nth-child(even) { background-color: #f2f2f2; }
                        .footer { text-align: center; margin-top: 30px; font-size: 12px; color: #666; }
                    </style>
                </head>
                <body>
                    <h1>🏛️ KSOP Kelas II Cirebon</h1>
                    <h2 style="text-align: center;">Laporan Kinerja Bulanan</h2>
                    <p style="text-align: center;">Tanggal Cetak: ${new Date().toLocaleDateString('id-ID')}</p>
                    ${tableHTML}
                    <div class="footer">
                        <p>© 2025 KSOP Kelas II Cirebon. Sistem Pelaporan Kinerja Bulanan</p>
                    </div>
                </body>
                </html>
            `;
            printWindow.document.open();
            printWindow.document.innerHTML = htmlContent;
            printWindow.document.close();
            printWindow.print();
        }

        // Load data pendukung dari localStorage
        function loadDataPendukung() {
            try {
                const raw = JSON.parse(localStorage.getItem('dataPendukung') || '[]');
                dataPendukung = Array.isArray(raw) ? raw : [];
                console.log('Data pendukung loaded:', dataPendukung.length);
            } catch (e) {
                console.error('Error loading data pendukung:', e);
                dataPendukung = [];
            }
        }

        // Populate pilihan seksi
        function populatePilihSeksi() {
            const select = document.getElementById('pilihSeksiPendukung');
            if (!select) return;

            select.innerHTML = '<option value="">-- Pilih Seksi --</option>';

            // Daftar seksi (hardcode)
            const daftarSeksi = [
                'Tata Usaha',
                'Status Hukum dan Kapal',
                'Keselamatan Berlayar Penjagaan dan Patrol',
                'Lalu Lintas Angkutan Laut'
            ];

            daftarSeksi.forEach(seksi => {
                const opt = document.createElement('option');
                opt.value = seksi;
                opt.textContent = seksi;
                select.appendChild(opt);
            });

            // Tambah event listener
            select.addEventListener('change', () => {
                populatePilihLaporan();
            });

            console.log('Seksi populated:', daftarSeksi.length);
        }

        // Populate pilihan laporan berdasarkan seksi yang dipilih
        function populatePilihLaporan() {
            const selectSeksi = document.getElementById('pilihSeksiPendukung');
            const selectLaporan = document.getElementById('pilihLaporan');
            
            if (!selectLaporan) return;

            const seksiTerpilih = selectSeksi.value;
            selectLaporan.innerHTML = '<option value="">-- Pilih Laporan --</option>';

            const laporanSet = new Set();
            
            (allData || []).forEach(row => {
                const seksi = String(row.seksi || row[1] || '').trim();
                const indikator = String(row.indikator || row[4] || '').trim();
                const bulan = String(row.bulan || row[2] || '').trim();
                const tahun = String(row.tahun || row[3] || '').trim();

                // Jika seksi kosong atau sesuai dengan pilihan
                if (!seksiTerpilih || seksi === seksiTerpilih) {
                    const laporan = `${indikator} - ${bulan} ${tahun}`;
                    if (laporan.trim() !== ' - ') {
                        laporanSet.add(laporan);
                    }
                }
            });

            Array.from(laporanSet).sort().forEach(laporan => {
                const opt = document.createElement('option');
                opt.value = laporan;
                opt.textContent = laporan;
                selectLaporan.appendChild(opt);
            });

            console.log('Laporan updated for seksi:', seksiTerpilih);
        }

        // Handle form submit data pendukung
        async function handleDataPendukungSubmit(event) {
            event.preventDefault();

            const pilihSeksi = document.getElementById('pilihSeksiPendukung').value;
            const pilihLaporan = document.getElementById('pilihLaporan').value;
            const jenisDokumen = document.getElementById('jenisDokumen').value;
            const deskripsi = document.getElementById('deskripsiPendukung').value;
            const fileInput = document.getElementById('filePendukung');
            const linkPendukung = document.getElementById('linkPendukung').value;

            if (!pilihSeksi) {
                alert('❌ Pilih seksi terlebih dahulu!');
                return;
            }

            if (!pilihLaporan) {
                alert('❌ Pilih laporan terlebih dahulu!');
                return;
            }

            let fileName = '';
            let fileSize = 0;
            let fileType = '';

            if (fileInput.files.length > 0) {
                const file = fileInput.files[0];
                
                // Cek ukuran file (max 10MB)
                if (file.size > 10 * 1024 * 1024) {
                    alert('❌ File terlalu besar! Max 10MB');
                    return;
                }

                fileName = file.name;
                fileSize = (file.size / 1024 / 1024).toFixed(2); // Convert to MB
                fileType = file.type;

                console.log('File:', { name: fileName, size: fileSize + 'MB', type: fileType });
            }

            const dataBaru = {
                id: Date.now(),
                tanggal: new Date().toLocaleDateString('id-ID'),
                seksi: pilihSeksi,
                laporan: pilihLaporan,
                jenisDokumen: jenisDokumen,
                deskripsi: deskripsi,
                namaFile: fileName,
                ukuranFile: fileSize,
                link: linkPendukung,
                createdAt: new Date().toISOString()
            };

            dataPendukung.push(dataBaru);
            localStorage.setItem('dataPendukung', JSON.stringify(dataPendukung));

            const alertDiv = document.getElementById('pendukungAlert');
            alertDiv.className = 'alert success';
            alertDiv.textContent = '✅ Data pendukung berhasil disimpan!';
            alertDiv.style.display = 'block';

            document.getElementById('dataPendukungForm').reset();
            displayDataPendukung();

            setTimeout(() => {
                alertDiv.style.display = 'none';
            }, 3000);
        }

        // Display data pendukung di tabel
        function displayDataPendukung() {
            const tbody = document.getElementById('tabelPendukungBody');
            if (!tbody) return;

            loadDataPendukung();
            tbody.innerHTML = '';

            if (dataPendukung.length === 0) {
                tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; padding: 20px; color: #666;">Belum ada data pendukung</td></tr>';
                updateStatistikPendukung();
                return;
            }

            dataPendukung.forEach((data, idx) => {
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td style="border: 1px solid #ddd; padding: 10px; text-align: center;">${idx + 1}</td>
                    <td style="border: 1px solid #ddd; padding: 10px;">${data.tanggal}</td>
                    <td style="border: 1px solid #ddd; padding: 10px; font-weight: bold; color: #0056b3;">${data.seksi}</td>
                    <td style="border: 1px solid #ddd; padding: 10px;">
                        <small>${data.laporan.substring(0, 40)}${data.laporan.length > 40 ? '...' : ''}</small>
                    </td>
                    <td style="border: 1px solid #ddd; padding: 10px;">${data.jenisDokumen}</td>
                    <td style="border: 1px solid #ddd; padding: 10px;">
                        <small>${data.deskripsi.substring(0, 50)}${data.deskripsi.length > 50 ? '...' : ''}</small>
                    </td>
                    <td style="border: 1px solid #ddd; padding: 10px;">
                        ${data.namaFile ? `
                            <a href="#" onclick="downloadFile('${data.namaFile}', '${data.id}')" style="color: #0056b3; text-decoration: none;">
                                📥 ${data.namaFile.substring(0, 20)}...
                            </a><br/>
                            <small style="color: #666;">${data.ukuranFile} MB</small>
                        ` : 'Tanpa file'}
                        ${data.link ? `<br/><a href="${data.link}" target="_blank" style="color: #0056b3; text-decoration: none;">🔗 Link</a>` : ''}
                    </td>
                    <td style="border: 1px solid #ddd; padding: 10px; text-align: center;">
                        <button onclick="viewDetailPendukung(${data.id})" style="background-color: #0056b3; color: white; padding: 5px 10px; border: none; border-radius: 3px; cursor: pointer; margin-right: 5px;">👁️</button>
                        <button onclick="deleteDataPendukung(${data.id})" style="background-color: #dc3545; color: white; padding: 5px 10px; border: none; border-radius: 3px; cursor: pointer;">🗑️</button>
                    </td>
                `;
                tbody.appendChild(tr);
            });

            updateStatistikPendukung();
        }

        // Update statistik
        function updateStatistikPendukung() {
            loadDataPendukung();
            
            const totalFile = dataPendukung.length;
            const totalSize = dataPendukung.reduce((sum, d) => sum + parseFloat(d.ukuranFile || 0), 0).toFixed(2);
            const totalJenis = new Set(dataPendukung.map(d => d.jenisDokumen)).size;

            const elTotalFile = document.getElementById('totalFilePendukung');
            const elTotalSize = document.getElementById('totalSizePendukung');
            const elTotalJenis = document.getElementById('totalJenisPendukung');

            if (elTotalFile) elTotalFile.textContent = totalFile;
            if (elTotalSize) elTotalSize.textContent = totalSize + ' MB';
            if (elTotalJenis) elTotalJenis.textContent = totalJenis;
        }

        // View detail data pendukung
        function viewDetailPendukung(id) {
            const data = dataPendukung.find(d => d.id === id);
            if (!data) return;

            alert(`
📋 Detail Data Pendukung:

Seksi: ${data.seksi}
Laporan: ${data.laporan}
Jenis: ${data.jenisDokumen}
Tanggal: ${data.tanggal}

Deskripsi:
${data.deskripsi}

File: ${data.namaFile || 'Tidak ada'}
Ukuran: ${data.ukuranFile} MB

Link: ${data.link || 'Tidak ada'}
            `);
        }

        // Delete data pendukung
        function deleteDataPendukung(id) {
            if (confirm('Hapus data pendukung ini?')) {
                dataPendukung = dataPendukung.filter(d => d.id !== id);
                localStorage.setItem('dataPendukung', JSON.stringify(dataPendukung));
                displayDataPendukung();
                alert('✅ Data pendukung dihapus');
            }
        }

        // Download file (demo)
        function downloadFile(fileName, id) {
            alert(`📥 Fitur download untuk: ${fileName}\n\nNota: Untuk production, simpan file di cloud storage (Google Drive, AWS S3, dll)`);
        }

        // Update showPage untuk data pendukung
        function showPage(page) {
            // Sembunyikan semua halaman
            document.getElementById('loginPage').style.display = 'none';
            document.getElementById('dashboardPage').style.display = 'none';
            document.getElementById('formPage').style.display = 'none';
            document.getElementById('graphicPage').style.display = 'none';
            document.getElementById('dataPendukungPage').style.display = 'none'; // pastikan ini ada

            // Tampilkan halaman sesuai argumen
            if (page === 'login') {
                document.getElementById('loginPage').style.display = 'flex';
            } else if (page === 'dashboard') {
                document.getElementById('dashboardPage').style.display = 'block';
            } else if (page === 'form') {
                document.getElementById('formPage').style.display = 'block';
            } else if (page === 'graphic') {
                document.getElementById('graphicPage').style.display = 'block';
                loadDataFromSheet();
            } else if (page === 'dataPendukung') {
                document.getElementById('dataPendukungPage').style.display = 'block';
                populatePilihSeksi();
                populatePilihLaporan();
                displayDataPendukung();
            }
        }

        // Initialize
        window.addEventListener('load', () => {
            loadDataPendukung();
        }, { once: true });