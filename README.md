# 세탁 서비스 관리 앱

작업(장당제/포대/월정액) · 정산(명세서 엑셀) · 기사 관리 · 고객사 관리 · 운영 메모

데이터는 Firebase Firestore에 저장되어 **모든 사용자가 같은 데이터를 실시간으로 공유**합니다.

---

## 배포 순서

### 1. Firestore 규칙 설정 (필수! 안 하면 앱이 안 됨)

Firebase 콘솔 → Firestore Database → **규칙** 탭에서 아래 내용으로 교체 후 "게시":

```
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /app/{document} {
      allow read, write: if true;
    }
  }
}
```

⚠️ 이 규칙은 링크를 아는 누구나 읽고 쓸 수 있는 상태입니다.
테스트/초기 운영용이며, 나중에 로그인(권한) 기능을 붙일 때 강화합니다.
"테스트 모드"로 만들었다면 30일 후 자동 차단되므로 반드시 위 규칙으로 바꿔두세요.

### 2. GitHub 업로드

1. github.com → New repository → 이름 예: `laundry-app` → Create
2. "uploading an existing file" 클릭
3. 이 폴더 안의 모든 파일/폴더를 드래그해서 업로드 (node_modules는 없음, 그대로 전부 올리면 됨)
4. Commit changes

### 3. Vercel 배포

1. vercel.com → Continue with GitHub 로그인
2. Add New → Project → `laundry-app` 저장소 Import
3. 설정은 건드릴 것 없이 **Deploy** 클릭 (Vite 자동 인식)
4. 1~2분 후 `xxx.vercel.app` 주소 생성 완료

### 4. 기사님들에게 공유

- 주소를 카톡으로 전달
- 폰에서 열고 브라우저 메뉴 → **"홈 화면에 추가"** → 앱처럼 사용

---

## 업데이트 방법

1. 수정된 파일을 GitHub 저장소에서 교체 (해당 파일 열기 → 연필 아이콘 → 붙여넣기 → Commit)
2. Vercel이 자동으로 감지해서 1~2분 안에 새 버전 배포
3. 사용자는 새로고침만 하면 최신 버전

데이터는 Firebase에 따로 저장되므로 코드를 업데이트해도 데이터는 유지됩니다.

## 로컬에서 실행해보기 (선택)

```bash
npm install
npm run dev
```
