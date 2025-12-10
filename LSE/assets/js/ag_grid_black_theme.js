/**
 * AG Grid Black Theme (내부망 대응 - 완전 버전)
 * 
 * AG Grid의 강력한 CSS 우선순위를 JavaScript로 덮어씀
 * 로컬 파일 사용하지만, 스타일 적용은 JavaScript 필요
 */

(function() {
    'use strict';
    
    const CONFIG = {
        BG_COLOR: '#000000',
        HOVER_COLOR: '#1a3a4a',  // 파란 톤 hover
        TEXT_COLOR: '#fea029',
        STYLED_FLAG: 'data-ag-black-styled'
    };
    
    // CSS 변수 목록
    const CSS_VARS = [
        '--ag-background-color',
        '--ag-header-background-color',
        '--ag-odd-row-background-color',
        '--ag-tooltip-background-color',
        '--ag-control-panel-background-color',
        '--ag-subheader-background-color',
        '--ag-row-hover-color'
    ];
    
    // 배경색 적용 대상 선택자
    const BG_SELECTORS = [
        '.ag-root',
        '.ag-root-wrapper',
        '.ag-root-wrapper-body',
        '.ag-header',
        '.ag-header-viewport',
        '.ag-header-container',
        '.ag-header-cell',
        '.ag-body-viewport',
        '.ag-center-cols-viewport',
        '.ag-center-cols-container',
        '.ag-row',
        '.ag-row-odd',
        '.ag-row-even',
        '.ag-cell',
        '.ag-tooltip',
        '.ag-popup',
        '.ag-menu',
        '.ag-filter',
        '.ag-paging-panel',
        '.ag-tool-panel'
    ];
    
    /**
     * 스타일 적용 메인 함수
     */
    function applyBlackTheme() {
        const gridElement = document.querySelector('.ag-theme-alpine-dark');
        if (!gridElement) return;
        
        // 이미 적용됨 체크
        if (gridElement.getAttribute(CONFIG.STYLED_FLAG) === 'true') {
            return;
        }
        
        // 1. CSS 변수 설정
        CSS_VARS.forEach(varName => {
            gridElement.style.setProperty(varName, CONFIG.BG_COLOR, 'important');
        });
        
        // 2. 그리드 자체 배경
        gridElement.style.setProperty('background-color', CONFIG.BG_COLOR, 'important');
        
        // 3. 모든 하위 요소 배경
        BG_SELECTORS.forEach(selector => {
            const elements = gridElement.querySelectorAll(selector);
            elements.forEach(el => {
                el.style.setProperty('background-color', CONFIG.BG_COLOR, 'important');
            });
        });
        
        // 플래그 설정
        gridElement.setAttribute(CONFIG.STYLED_FLAG, 'true');
        
        console.log('[AG Grid] Black theme applied');
    }
    
    /**
     * MutationObserver로 동적 요소 감지
     * ★ 단, 컬럼 리사이즈 등 사용자 액션은 방해하지 않도록 제한
     */
    function observeGridChanges() {
        const gridElement = document.querySelector('.ag-theme-alpine-dark');
        if (!gridElement) {
            setTimeout(observeGridChanges, 100);
            return;
        }
        
        // ★ 새로운 행(row)이 추가될 때만 스타일 적용 (리사이즈는 무시)
        const observer = new MutationObserver((mutations) => {
            mutations.forEach(mutation => {
                if (mutation.addedNodes.length > 0) {
                    mutation.addedNodes.forEach(node => {
                        if (node.nodeType === 1) { // Element 노드만
                            // ★ .ag-row만 처리 (헤더/컬럼 변경은 무시)
                            if (node.classList && node.classList.contains('ag-row')) {
                                node.style.setProperty('background-color', CONFIG.BG_COLOR, 'important');
                                node.querySelectorAll('.ag-cell').forEach(cell => {
                                    cell.style.setProperty('background-color', CONFIG.BG_COLOR, 'important');
                                });
                            }
                        }
                    });
                }
            });
        });
        
        observer.observe(gridElement, {
            childList: true,
            subtree: true
        });
    }
    
    /**
     * Hover 효과 추가 (JavaScript로 직접 처리)
     */
    function setupHoverEffect() {
        const gridElement = document.querySelector('.ag-theme-alpine-dark');
        if (!gridElement) {
            setTimeout(setupHoverEffect, 100);
            return;
        }
        
        // 이벤트 위임: 그리드 전체에 리스너 추가
        gridElement.addEventListener('mouseover', (e) => {
            const row = e.target.closest('.ag-row');
            if (row) {
                row.style.setProperty('background-color', CONFIG.HOVER_COLOR, 'important');
                // 모든 셀도 같이 변경
                row.querySelectorAll('.ag-cell').forEach(cell => {
                    cell.style.setProperty('background-color', CONFIG.HOVER_COLOR, 'important');
                });
            }
        });
        
        gridElement.addEventListener('mouseout', (e) => {
            const row = e.target.closest('.ag-row');
            if (row) {
                row.style.setProperty('background-color', CONFIG.BG_COLOR, 'important');
                // 모든 셀도 원래대로
                row.querySelectorAll('.ag-cell').forEach(cell => {
                    cell.style.setProperty('background-color', CONFIG.BG_COLOR, 'important');
                });
            }
        });
        
        console.log('[AG Grid] Hover effect enabled');
    }
    
    /**
     * 초기화
     */
    function init() {
        // DOM 로드 후 즉시 적용
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', () => {
                applyBlackTheme();
                setupHoverEffect();
            });
        } else {
            applyBlackTheme();
            setupHoverEffect();
        }
        
        // ★ MutationObserver는 완전히 제거 (컬럼 리사이즈 간섭 방지)
        // clientside callback에서 처리하므로 불필요
    }
    
    // 전역 함수로 노출 (clientside callback에서 호출 가능)
    window.applyAgGridBlackTheme = applyBlackTheme;
    
    init();
    
})();
