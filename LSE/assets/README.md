# Assets 폴더 구조

NAV 대시보드의 모든 정적 파일을 관리합니다.

---

## 📁 폴더 구조

```
assets/
├── css/                          # 스타일시트
│   ├── vendor/                   # 외부 라이브러리 (원본, 수정 금지)
│   │   ├── ag-grid.css          # AG Grid 기본 CSS (235KB)
│   │   └── ag-theme-alpine.css  # AG Grid Alpine 테마 (32KB)
│   └── dark_theme_override.css  # 다크 테마 커스터마이징 (4KB)
│
└── js/                           # JavaScript
    └── ag_grid_black_theme.js   # 배경색 강제 적용 (5KB)
```

---

## 📄 파일 설명

### CSS 파일

#### Vendor 폴더 (외부 라이브러리 원본)

##### 1. `vendor/ag-grid.css` (원본, 수정 금지)
- **출처**: AG Grid Community v31.0.0
- **역할**: AG Grid 기본 구조 및 스타일
- **위치**: `css/vendor/` (원본 파일 전용 폴더)
- **업데이트**: AG Grid 버전 업그레이드 시 교체

##### 2. `vendor/ag-theme-alpine.css` (원본, 수정 금지)
- **출처**: AG Grid Alpine 테마
- **역할**: Alpine 테마 기본 스타일
- **위치**: `css/vendor/` (원본 파일 전용 폴더)
- **업데이트**: AG Grid 버전 업그레이드 시 교체

#### 커스텀 파일

##### 3. `dark_theme_override.css` (커스텀, 수정 가능)
- **역할**: 다크 테마 커스터마이징
- **내용**:
  - CSS 변수 Override
  - 배경색 강제 적용
  - 호버 색상 (파란 톤)
  - 텍스트 색상 (주황색 #fea029)
- **수정**: 색상 변경 시 이 파일만 수정

---

### JavaScript 파일

#### 1. `ag_grid_black_theme.js`
- **역할**: AG Grid 배경색을 #000000(검은색)으로 강제 적용
- **필요성**: AG Grid의 강력한 CSS 우선순위 때문에 JavaScript 필수
- **작동 방식**:
  1. 초기 로드 시 배경색 적용
  2. MutationObserver로 동적 요소 감시
  3. 새 요소 추가 시 즉시 배경색 적용
- **성능**: < 10ms 오버헤드 (HFT에 무시할 수 있는 수준)

---

## 🔧 유지보수 가이드

### 색상 변경

#### 텍스트 색상 변경
```css
/* css/dark_theme_override.css */
.ag-theme-alpine-dark .ag-cell.cell-price {
    color: #fea029 !important;  /* ← 여기만 수정 */
}
```

#### 배경색 변경
```javascript
// js/ag_grid_black_theme.js
const CONFIG = {
    BG_COLOR: '#000000',  // ← 여기만 수정
};
```

#### 호버 색상 변경
```css
/* css/dark_theme_override.css */
.ag-theme-alpine-dark {
    --ag-row-hover-color: #1a3a4a !important;  /* ← 여기만 수정 */
}
```

---

### AG Grid 버전 업그레이드

#### Step 1: 새 버전 다운로드
```powershell
cd assets/css/vendor

# 기존 파일 백업
copy ag-grid.css ag-grid_v31.css
copy ag-theme-alpine.css ag-theme-alpine_v31.css

# 새 버전 다운로드
Invoke-WebRequest -Uri "https://cdn.jsdelivr.net/npm/ag-grid-community@32.0.0/styles/ag-grid.css" -OutFile "ag-grid.css"
Invoke-WebRequest -Uri "https://cdn.jsdelivr.net/npm/ag-grid-community@32.0.0/styles/ag-theme-alpine.css" -OutFile "ag-theme-alpine.css"
```

#### Step 2: 테스트
```powershell
cd ../../../
python scripts/nav_dashboard.py
```

#### Step 3: 문제 발생 시
```powershell
cd assets/css/vendor

# 백업에서 복원
copy ag-grid_v31.css ag-grid.css
copy ag-theme-alpine_v31.css ag-theme-alpine.css
```

---

## ⚠️ 주의사항

### DO ✅
- `dark_theme_override.css`만 수정
- CSS 변수 활용
- `vendor/` 폴더 원본 파일 보존
- 업데이트 전 백업

### DON'T ❌
- `vendor/ag-grid.css` 직접 수정
- `vendor/ag-theme-alpine.css` 직접 수정
- `vendor/` 폴더 밖으로 원본 파일 이동
- CDN 링크 추가
- JavaScript 없이 배경색 변경 시도

---

## 📚 관련 문서

- **전체 가이드**: `docs/INTERNAL_NETWORK_SETUP.md`
- **AG Grid 교훈**: `docs/AG_GRID_LESSONS_LEARNED.md`
- **컬럼 설정**: `docs/NAV_DASHBOARD_COLUMNS.md`

---

**마지막 업데이트**: 2025-12-07  
**AG Grid 버전**: v31.0.0  
**관리자**: LP Analysis Team

