import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import matplotlib.patches as patches
import numpy as np
import datetime
import os
import traceback
from scipy.interpolate import make_interp_spline
import io
import base64
import tempfile

# 한글 폰트 설정 함수
def set_korean_font():
    """한글 폰트 설정"""
    try:
        # Streamlit 환경에서는 시스템 폰트 사용
        plt.rc('font', family='NanumGothic')
        plt.rcParams['axes.unicode_minus'] = False
        
        # 폰트 확인
        fonts = [f.name for f in fm.fontManager.ttflist]
        if 'NanumGothic' not in fonts:
            st.warning("나눔고딕 폰트가 설치되어 있지 않아 일부 한글이 제대로 표시되지 않을 수 있습니다.")
            # 대체 폰트 설정
            for font in fonts:
                if any(korean_font in font for korean_font in ['Malgun', 'Gulim', 'Batang', 'Dotum']):
                    plt.rc('font', family=font)
                    return True
        
        return True
    except Exception as e:
        st.error(f"한글 폰트 설정 실패: {str(e)}")
        return False

# 시간 변환 함수들
def convert_to_time_str(time_val):
    """시간값을 HH:MM:SS 형식의 문자열로 변환"""
    if isinstance(time_val, datetime.time):
        return f"{time_val.hour:02d}:{time_val.minute:02d}:{time_val.second:02d}"
    elif isinstance(time_val, str):
        parts = time_val.split(':')
        if len(parts) >= 3:
            return f"{parts[0].zfill(2)}:{parts[1].zfill(2)}:{parts[2].zfill(2)}"
        elif len(parts) == 2:
            return f"{parts[0].zfill(2)}:{parts[1].zfill(2)}:00"
    return str(time_val)

def time_to_seconds(time_str):
    """HH:MM:SS 형식의 시간 문자열을 초 단위로 변환"""
    try:
        parts = time_str.split(':')
        hours = int(parts[0])
        minutes = int(parts[1])
        seconds = int(float(parts[2])) if len(parts) >= 3 else 0
        return hours * 3600 + minutes * 60 + seconds
    except Exception:
        return None

def seconds_to_time_str(seconds):
    """초를 시:분:초 형식의 문자열로 변환"""
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    seconds = seconds % 60
    return f"{int(hours):02d}:{int(minutes):02d}:{int(seconds):02d}"

# 프로그램 데이터 검색 함수
def find_program_data_improved(sheet, program_names):
    """개선된 프로그램 데이터 검색"""
    # 데이터프레임의 모든 값을 문자열로 변환
    df_str = sheet.fillna('').astype(str)
    
    with st.expander("프로그램 검색 상세 로그", expanded=False):
        st.write(f"프로그램 검색 중: {program_names}")

    for program_name in program_names:
        # 전체 데이터프레임에서 프로그램명 검색
        for col in range(1, 5):  # B~E열 (인덱스 1~4)만 검색
            for row in range(df_str.shape[0]):
                cell_value = df_str.iloc[row, col].strip()
                if cell_value and program_name in cell_value:
                    if any(char.isdigit() for char in cell_value):  # 숫자가 포함된 셀만 처리
                        with st.expander("프로그램 검색 상세 로그", expanded=False):
                            st.write(f"프로그램명 '{program_name}' 발견: 행={row}, 열={col}")
                            st.write(f"전체 내용: {cell_value}")

                        try:
                            # 시청률 추출 (프로그램명 다음의 숫자)
                            rating_str = cell_value.split(program_name)[1].strip().split('(')[0].strip()
                            rating = float(rating_str)
                            return rating, row, col
                        except (IndexError, ValueError) as e:
                            with st.expander("프로그램 검색 상세 로그", expanded=False):
                                st.write(f"시청률 추출 실패: {str(e)}")
                            continue

    return None, None, None

def get_program_ratings(sheet_paid, sheet_2049, date_str):
    """종편 뉴스 프로그램들의 시청률 데이터를 추출"""
    # 주말 여부 확인
    date = datetime.datetime.strptime(date_str, '%y%m%d')
    is_weekend = date.weekday() >= 5
    
    with st.expander("프로그램 데이터 검색 로그", expanded=False):
        st.write("=== 프로그램 데이터 검색 시작 ===")
        st.write(f"검색 모드: {'주말' if is_weekend else '평일'}")

    programs = {
        'news_a': ['뉴스A', '특집뉴스A'],
        'jtbc': ['JTBC뉴스룸', '특집JTBC뉴스룸'],
        'mbn': ['MBN뉴스센터', '특집MBN뉴스센터'] if is_weekend else ['MBN뉴스7', '특집MBN뉴스7'],
        'tv_chosun': ['TV조선뉴스7', '특집TV조선뉴스7'] if is_weekend else ['TV조선뉴스9', '특집TV조선뉴스9']
    }

    with st.expander("프로그램별 데이터 검색 결과", expanded=False):
        st.write("=== 프로그램별 데이터 확인 ===")
        st.markdown("| 프로그램명 | 수도권 유료가구 | 수도권 20-49 |")
        st.markdown("|------------|--------------|------------|")

    ratings = {}

    for program_key, program_names in programs.items():
        # 수도권 유료가구 시청률 찾기
        paid_rating, paid_row, paid_col = find_program_data_improved(sheet_paid, program_names)
        paid_info = f"찾음 (행 {paid_row+1}, 열 {paid_col+1}, 값: {paid_rating}%)" if paid_rating is not None else "찾지 못함"

        # 수도권 2049 시청률 찾기
        rating_2049, r2049_row, r2049_col = find_program_data_improved(sheet_2049, program_names)
        rating_2049_info = f"찾음 (행 {r2049_row+1}, 열 {r2049_col+1}, 값: {rating_2049}%)" if rating_2049 is not None else "찾지 못함"

        with st.expander("프로그램별 데이터 검색 결과", expanded=False):
            st.markdown(f"| {program_names[0]} | {paid_info} | {rating_2049_info} |")

        ratings[program_key] = {
            'name': program_names[0],
            'paid': paid_rating,
            'rating_2049': rating_2049
        }

    return ratings

def find_news_a_data(df):
    """뉴스A 데이터가 있는 영역 찾기"""
    with st.expander("데이터 구조 상세 확인", expanded=False):
        st.write("=== 데이터 구조 상세 확인 ===")
        st.write(f"행 수: {df.shape[0]}, 열 수: {df.shape[1]}")

    program_row = None
    for row in range(df.shape[0]):
        if pd.notna(df.iloc[row, 0]) and str(df.iloc[row, 0]).strip() == "프로그램":
            program_row = row
            break

    if program_row is None:
        raise ValueError("프로그램 행을 찾을 수 없습니다.")

    program_cols = []
    for col in range(df.shape[1]):
        if pd.notna(df.iloc[program_row, col]) and str(df.iloc[program_row, col]).strip() == "프로그램":
            program_cols.append(col)

    news_a_row = None
    found_program_col = None
    found_time_col = None
    found_rating_col = None
    found_2049_col = None

    for program_col in program_cols:
        for row in range(program_row + 1, df.shape[0]):
            if pd.notna(df.iloc[row, program_col]):
                val = str(df.iloc[row, program_col]).strip()
                if val in ["뉴스A", "특집뉴스A"]:
                    news_a_row = row
                    found_program_col = program_col
                    found_time_col = program_col + 1
                    found_rating_col = program_col + 2

                    # 2049 데이터 열 찾기
                    for col in range(program_col, min(program_col + 10, df.shape[1])):
                        if pd.notna(df.iloc[0, col]) and "수도권 2049" in str(df.iloc[0, col]):
                            found_2049_col = col
                            break

                    if found_2049_col is None:
                        found_2049_col