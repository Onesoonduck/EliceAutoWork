import pandas as pd
import os
from config import EXCEL_SETTINGS

class ExcelProcessor:
    def __init__(self, chapter_number):
        self.chapter_number = chapter_number
        self.download_path = EXCEL_SETTINGS['download_path']
        
    def get_latest_report(self):
        """가장 최근에 다운로드된 리포트 파일 찾기"""
        files = [f for f in os.listdir(self.download_path) if f.endswith('.xlsx')]
        if not files:
            raise FileNotFoundError("다운로드된 리포트 파일이 없습니다.")
        
        # 가장 최근 파일 선택
        latest_file = max(files, key=lambda x: os.path.getctime(os.path.join(self.download_path, x)))
        return os.path.join(self.download_path, latest_file)
    
    def process_report(self):
        """리포트 파일 처리"""
        report_file = self.get_latest_report()
        
        # 리포트 데이터 읽기
        report_df = pd.read_excel(report_file, sheet_name=EXCEL_SETTINGS['report_sheet_name'])
        
        # 필요한 데이터 추출 (유저명과 진도율)
        progress_data = report_df[['유저명', '진도율']].copy()
        
        # 레이서 정보 시트 업데이트
        racer_file = os.path.join(self.download_path, f'chapter_{self.chapter_number}_racer_info.xlsx')
        racer_df = pd.read_excel(racer_file, sheet_name=EXCEL_SETTINGS['racer_sheet_name'])
        
        # 진도율 업데이트
        for _, row in progress_data.iterrows():
            username = row['유저명']
            progress = row['진도율']
            
            # 해당 유저의 행 찾기
            mask = racer_df['유저명'] == username
            if mask.any():
                racer_df.loc[mask, EXCEL_SETTINGS['progress_column']] = progress
        
        # 업데이트된 파일 저장
        with pd.ExcelWriter(racer_file) as writer:
            racer_df.to_excel(writer, sheet_name=EXCEL_SETTINGS['racer_sheet_name'], index=False)
            
        return progress_data 