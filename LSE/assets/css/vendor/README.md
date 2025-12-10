# Vendor 폴더 (외부 라이브러리 원본)

이 폴더는 **외부 라이브러리의 원본 파일**만 보관합니다.

---

## ⚠️ 중요 규칙

### ❌ 절대 수정 금지
이 폴더의 파일들은 **외부에서 다운로드한 원본**입니다.
- 직접 수정하지 마세요
- 스타일을 변경하고 싶으면 `../dark_theme_override.css` 사용

### ✅ 업데이트 시에만 교체
- AG Grid 새 버전이 나왔을 때만 교체
- 교체 전 반드시 백업

---

## 📄 파일 목록

### 1. `ag-grid.css` (235KB)
- **출처**: [AG Grid Community](https://www.ag-grid.com/)
- **버전**: v31.0.0
- **다운로드**: 
  ```powershell
  Invoke-WebRequest -Uri "https://cdn.jsdelivr.net/npm/ag-grid-community@31.0.0/styles/ag-grid.css" -OutFile "ag-grid.css"
  ```

### 2. `ag-theme-alpine.css` (32KB)
- **출처**: [AG Grid Alpine Theme](https://www.ag-grid.com/)
- **버전**: v31.0.0
- **다운로드**:
  ```powershell
  Invoke-WebRequest -Uri "https://cdn.jsdelivr.net/npm/ag-grid-community@31.0.0/styles/ag-theme-alpine.css" -OutFile "ag-theme-alpine.css"
  ```

---

## 🔄 업데이트 방법

### 1. 백업
```powershell
copy ag-grid.css ag-grid_v31_backup.css
copy ag-theme-alpine.css ag-theme-alpine_v31_backup.css
```

### 2. 새 버전 다운로드
```powershell
# v32.0.0으로 업데이트 예시
Invoke-WebRequest -Uri "https://cdn.jsdelivr.net/npm/ag-grid-community@32.0.0/styles/ag-grid.css" -OutFile "ag-grid.css"
Invoke-WebRequest -Uri "https://cdn.jsdelivr.net/npm/ag-grid-community@32.0.0/styles/ag-theme-alpine.css" -OutFile "ag-theme-alpine.css"
```

### 3. 테스트
```powershell
cd ../../../
python scripts/nav_dashboard.py
```

### 4. 문제 발생 시 롤백
```powershell
copy ag-grid_v31_backup.css ag-grid.css
copy ag-theme-alpine_v31_backup.css ag-theme-alpine.css
```

---

## 📚 참고

- 상위 폴더 README: `../README.md`
- AG Grid 공식 문서: https://www.ag-grid.com/
- AG Grid CDN: https://cdn.jsdelivr.net/npm/ag-grid-community/

---

**마지막 업데이트**: 2025-12-07  
**현재 버전**: AG Grid Community v31.0.0

