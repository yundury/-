# 캐릭터 소개 웹앱

5명의 캐릭터를 소개하는 간단한 웹앱입니다. 순수 HTML/CSS/JS로 만들어져 있고, 서버나 빌드 과정 없이 바로 동작합니다.

## 로컬에서 확인하기

```bash
python3 -m http.server 8000
```

브라우저에서 `http://localhost:8000` 접속.

## 캐릭터 내용 수정하기

`data/characters.js` 파일 하나만 수정하면 됩니다. 이름, 특징, 색상, 이모지, 설명 등을 자유롭게 바꿀 수 있습니다.

## GitHub Pages로 배포하기

1. GitHub 저장소 → **Settings → Pages**
2. **Source**를 `GitHub Actions`로 설정
3. `main` 브랜치에 push하면 `.github/workflows/deploy.yml`이 자동으로 배포합니다
4. 배포 후 `https://<사용자명>.github.io/<저장소명>/` 주소로 접속 가능

## 아이폰에 앱처럼 설치하기

1. 배포된 주소를 아이폰 Safari로 접속
2. 공유 버튼 → **홈 화면에 추가**
3. 홈 화면 아이콘을 누르면 전체화면 앱처럼 실행됩니다 (PWA)
