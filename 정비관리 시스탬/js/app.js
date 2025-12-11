document.addEventListener('DOMContentLoaded', () => {
    const navLinks = document.querySelectorAll('.nav-links li');
    const pageTitle = document.getElementById('page-title');
    const contentArea = document.getElementById('content-area');

    // URLs for existing systems
    const systems = {
        dashboard: 'dashboard',
        sootblower: 'https://kdy871.github.io/soot-blower-manager/',
        brush: 'https://kdy871.github.io/brush-manager/inspector.html'
    };

    navLinks.forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();

            // Remove active class from all
            navLinks.forEach(l => l.classList.remove('active'));
            // Add active class to clicked
            link.classList.add('active');

            const target = link.getAttribute('data-target');
            loadContent(target);
        });
    });

    function loadContent(target) {
        let title = '';
        let content = '';

        switch (target) {
            case 'dashboard':
                title = '대시보드';
                content = `
                    <div class="dashboard-widgets">
                        <div class="widget">
                            <h3>시스템 상태</h3>
                            <p>정상 가동 중</p>
                        </div>
                        <div class="widget">
                            <h3>바로가기</h3>
                            <p>왼쪽 메뉴를 통해 각 시스템에 접속하세요.</p>
                        </div>
                    </div>
                `;
                break;
            case 'sootblower':
                title = 'Soot Blower 관리';
                content = `<iframe src="${systems.sootblower}"></iframe>`;
                break;
            case 'brush':
                title = '브러쉬 이력 조회';
                content = `<iframe src="${systems.brush}"></iframe>`;
                break;
        }

        pageTitle.textContent = title;
        contentArea.innerHTML = content;
    }
});
