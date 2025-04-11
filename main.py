from web_automation import WebAutomation
import os
from config import LOGIN_INFO, EXCEL_SETTINGS

def main():
    try:
        # 웹 자동화 시작
        web_auto = WebAutomation()
        web_auto.login()
        
        print("로그인이 완료되었습니다.")
        
        # 대시보드 URL로 이동
        dashboard_url = input("▶ 이동할 대시보드 URL을 입력하세요:\n> ").strip()
        chapter_num = input("▶ 챕터 번호를 입력하세요:\n> ").strip()
        
        web_auto.navigate_to_chapter_report(dashboard_url, chapter_num)
        
        input("계속하려면 아무 키나 누르세요...")  # 브라우저 확인용 대기
        
    except Exception as e:
        print(f"오류 발생: {str(e)}")
    finally:
        web_auto.close()

if __name__ == "__main__":
    main()
