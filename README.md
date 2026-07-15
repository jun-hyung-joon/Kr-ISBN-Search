# ISBN 조회 시스템

도서의 ISBN 바코드를 카메라로 스캔하여 국립중앙도서관 서지정보 API에서 정보를 조회하고, Excel 파일(.xlsx)로 저장하는 정적 웹 애플리케이션입니다.

별도의 설치 없이 PC 및 모바일 브라우저에서 사용할 수 있습니다.


## 사용 방법

### 사용자

1. 국립중앙도서관 서지정보 API에서 인증키를 발급받습니다.
2. 아래 주소 뒤에 인증키를 붙여 접속합니다.

```
https://jun-hyung-joon.github.io/Kr-ISBN-Search/?key=본인의_API인증키
```

3. 카메라 권한을 허용합니다.
4. ISBN 바코드를 스캔합니다.
5. **결과조회**를 눌러 Excel 파일을 저장합니다.


### 개발자

저장소를 Fork 또는 Clone한 뒤 GitHub Pages를 활성화합니다.

접속 주소

```
https://본인아이디.github.io/저장소명/?key=본인의_API인증키
```

## 커스터마이징

엑셀에 출력할 항목은 `app.js` 상단의 `EXCEL_COLUMNS` 배열만 수정하면 됩니다. 

```javascript
const EXCEL_COLUMNS = [
    { header: 'ISBN', width: 16 },
    { header: 'TITLE', width: 25 },
    { header: 'AUTHOR', width: 18 },
    { header: 'PUBLISHER', width: 18 }
];
```

### 규칙

- `header` : 국립중앙도서관 API 필드명
- `width` : Excel 열 너비
- 배열 순서 = Excel 열 순서
- `ISBN` 열은 필수입니다.
- 조회 결과가 없는 항목은 자동으로 `-`가 입력됩니다.

필요한 API 필드를 추가하거나 삭제하여 원하는 형식의 Excel 파일을 만들 수 있습니다.


## 기술 스택

- HTML5
- CSS3
- Vanilla JavaScript (ES6+)
- ZXing
- SheetJS (xlsx)


## 라이선스

MIT License

조회되는 도서 정보의 저작권 및 이용 조건은 국립중앙도서관 서지정보 API 이용약관을 따릅니다.