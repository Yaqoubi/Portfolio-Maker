(function () {
    if (window.location.href !== "https://sida.medu.ir/#/studentSearch") {
        alert("لطفا در صفحه مشاهده مشخصات دانش آموزان از این اسکریپت استفاده کنید");
        return;
    }

    if (typeof jQuery == 'undefined') {
        var script = document.createElement('script');
        script.src = 'https://code.jquery.com/jquery-3.6.0.min.js';
        document.head.appendChild(script);
        script.onload = function () {
            main();
        };
    } else {
        main();
    }

    function main() {
        const tableData = [];
        let previousRowCount = 0;
        let currentPage = 1;

        function getTableData() {
            const currentPageData = [];
            $('table tbody tr').each(function () {
                const rowData = [];
                $(this).find('td').each(function () {
                    rowData.push($(this).text().trim());
                });
                currentPageData.push(rowData);
            });
            return currentPageData;
        }

        function getColumnHeaders() {
            const headers = [];
            $('table thead th').each(function () {
                const headerText = $(this).text().trim();
                if (headerText) {
                    headers.push(headerText);
                }
            });
            return headers;
        }

        function waitForPageLoad() {
            return new Promise(resolve => {
                const observer = new MutationObserver((mutations, obs) => {
                    const rows = $('table tbody tr');
                    if (rows.length > 0) {
                        obs.disconnect();
                        resolve(true);
                    }
                });
                observer.observe(document.querySelector('table tbody'), { childList: true });
            });
        }

        function clickNextPage(page) {
            return new Promise(resolve => {
                const nextPageLink = $(`a[data-page="${page}"]`);
                if (nextPageLink.length === 0) {
                    resolve(false);
                } else {
                    nextPageLink[0].click();
                    waitForPageLoad().then(() => resolve(true));
                }
            });
        }

        function goToFirstPage() {
            return new Promise((resolve, reject) => {
                const firstPageLink = $('a.k-pager-first');
                if (firstPageLink.hasClass('k-state-disabled')) {
                    resolve();
                } else {
                    firstPageLink[0].click();
                    waitForPageLoad().then(() => resolve());
                }
            });
        }

        function showLoadingScreen() {
            const loadingOverlay = document.createElement('div');
            loadingOverlay.id = 'loadingOverlay';
            loadingOverlay.style.position = 'fixed';
            loadingOverlay.style.top = '0';
            loadingOverlay.style.left = '0';
            loadingOverlay.style.width = '100%';
            loadingOverlay.style.height = '100%';
            loadingOverlay.style.backgroundColor = 'rgba(255,255,255,0.5)';
            loadingOverlay.style.zIndex = '10000';
            loadingOverlay.style.display = 'flex';
            loadingOverlay.style.alignItems = 'center';
            loadingOverlay.style.justifyContent = 'center';
            loadingOverlay.innerHTML = '<h1 style="font-size:30px;color:black;">لطفا کمی صبر کنید...</h1>';
            document.body.appendChild(loadingOverlay);
        }

        function hideLoadingScreen() {
            const loadingOverlay = document.getElementById('loadingOverlay');
            if (loadingOverlay) {
                loadingOverlay.remove();
            }
        }

        async function scrapeAllPages() {
            showLoadingScreen();
            await goToFirstPage();
            const headers = getColumnHeaders();
            tableData.push(headers);
            while (true) {
                const currentPageData = getTableData();
                const currentRowCount = currentPageData.length;
                if (currentRowCount < previousRowCount && previousRowCount !== 0) {
                    break;
                }
                tableData.push(...currentPageData);
                previousRowCount = currentRowCount;
                currentPage++;
                const morePages = await clickNextPage(currentPage);
                if (!morePages) break;
            }
            tableData.push(...getTableData());
            hideLoadingScreen();
            exportToExcel(tableData);
        }

        function exportToExcel(data) {
            const bom = '\uFEFF';
            let csvContent = bom + data.map(row =>
                row.map(cell => '"' + cell.replace(/"/g, '""') + '"').join(",")
            ).join("\n");

            let blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8' });
            let link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = 'students_data_with_headers.csv';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }

        scrapeAllPages();
    }
})();
