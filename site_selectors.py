"""
KOTRA 보고서 생성 화면의 selector를 한곳에서 관리합니다.

아래 값은 1차 구현용 임시 selector입니다. 실제 사이트 DOM이 다르면
README의 codegen 안내를 참고해 이 파일만 교체하면 됩니다.
"""

SELECTORS = {
    "recommend_analysis_button": "text=유망 시장 추천 받기",
    "direct_analysis_button": "text=희망 국가 직접 분석",

    "hs_code_input": "input[placeholder*='6자리']",
    "product_name_input": "input[placeholder*='음료']",
    "export_scale_dropdown": "xpath=//label[contains(normalize-space(.), '수출액 규모')]/following-sibling::button[@role='combobox'][1]",

    "export_experience_first": "text=처음입니다",
    "export_experience_has": "text=수출 경험 있음",

    "target_country_input": "input[placeholder*='미국']",
    "excluded_country_input": "xpath=//label[contains(normalize-space(.), '분석 제외 국가')]/following::input[1]",
    "excluded_country_add_button": "xpath=//input[@placeholder='제외할 국가명 입력']/following-sibling::button[1]",
    "market_analysis_section": "xpath=//*[contains(normalize-space(.), '국가별 시장 분석')]/ancestor::*[self::div or self::section][1]",
    "market_analysis_card": "xpath=//*[contains(normalize-space(.), '국가별 시장 분석')]/following::*[.//button][self::div or self::section][position() <= 8]",
    "market_country_chip": "button",

    "generate_button": "button:has-text('보고서 생성')",
    "download_button": "button:has-text('PDF 저장')",
    "download_button_fallback": "button:has-text('다운로드')",
    # 문구가 충분히 고유한 상태 전환 버튼은 button 외 구현(a, role=button)도 허용한다.
    "retry_button": "button:has-text('다시 시도하기'), a:has-text('다시 시도하기'), [role='button']:has-text('다시 시도하기')",
    "streaming_error_text": "text=Unexpected server Streaming error occurred",
    "reset_button": "button:has-text('초기화')",
    "new_analysis_button": "button:has-text('새로운 분석 시작하기'), a:has-text('새로운 분석 시작하기'), [role='button']:has-text('새로운 분석 시작하기')",
}
