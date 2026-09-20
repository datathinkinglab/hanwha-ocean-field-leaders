/*
 * 실시간 QR 투표 설정 — index.html(강의 화면)과 vote.html(휴대폰 투표 화면)이 함께 사용합니다.
 * Apps Script 웹 앱을 배포한 뒤, 아래 POLL_API_URL 에 ".../exec" 로 끝나는 주소를 붙여 넣으세요.
 */
window.POLL_CONFIG = {
  // Google Apps Script 웹 앱 URL (예: https://script.google.com/macros/s/AKfy.../exec)
  POLL_API_URL: 'https://script.google.com/macros/s/AKfycbw5rC7K8f-1k1syStvmvgQoKPaGC0h8kYRsYV2_a5sQopMBxd0AdNxd87-eg5MnETAc/exec',

  // 집계 초기화 비밀번호 — Code.gs 의 ADMIN_KEY 와 같게 맞춰 주세요.
  POLL_ADMIN_KEY: 'hanwha-2026',

  // 강의 화면이 집계를 새로 읽는 간격(밀리초)
  REFRESH_MS: 2500,

  // 휴대폰 투표 화면에 보이는 보기 (강의 화면 투표 슬라이드와 순서 동일)
  OPTIONS: [
    '매일 교대 때마다 작업일지·안전체크 수기 작성 및 중복 정리',
    '도면 변경점 확인이나 장비 매뉴얼에서 에러코드 찾느라 헤매기',
    '신규 작업자나 외국인 근로자에게 같은 작업·안전 지시 반복 설명',
    '용접 결함, 부품 단차, 라벨 오타 놓칠까 봐 눈 아프게 대조하기'
  ]
};
