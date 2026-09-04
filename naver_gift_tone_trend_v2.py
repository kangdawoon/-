# -*- coding: utf-8 -*-
"""
선물 '톤앤매너' 키워드 검색 트렌드 확장 수집 스크립트
- naver_gift_tone_trend.py에서 데이터가 비어있던 "성의있는 선물"을
  유의어("정성스러운 선물", "진심 선물")로 교체해 재조사합니다.
- 오쏘몰이 비타민 브랜드라는 점을 고려해, 카테고리 직결 키워드
  (건강 선물·효도 선물·비타민 선물)와 톤 대비 키워드
  (센스있는 선물·실속 선물·가성비 선물), 시즌 표현 변형(명절 선물세트),
  감성 키워드(감동 선물)를 추가로 조사합니다.
- 기존 naver_gift_tone_trend.csv의 유효 데이터(오쏘몰·선물 추천·
  부담없는 선물·고급 선물)와 합쳐서 naver_gift_tone_trend.csv를 갱신합니다.

[인증키]
.env 파일의 NAVER_CLIENT_ID / NAVER_CLIENT_SECRET을 그대로 사용합니다.

[실행 방법]
pip install requests python-dotenv pandas
python naver_gift_tone_trend_v2.py
"""

import requests
import json
import time
import sys
from datetime import datetime, timedelta

import pandas as pd
from dotenv import load_dotenv
import os

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
load_dotenv()

CLIENT_ID = os.getenv("NAVER_CLIENT_ID")
CLIENT_SECRET = os.getenv("NAVER_CLIENT_SECRET")

if not CLIENT_ID or not CLIENT_SECRET:
    raise ValueError(
        ".env 파일에서 NAVER_CLIENT_ID / NAVER_CLIENT_SECRET을 찾을 수 없습니다.\n"
        "같은 폴더에 .env 파일이 있는지, 값이 정확히 들어있는지 확인해주세요."
    )

URL = "https://openapi.naver.com/v1/datalab/search"

# 네이버 데이터랩 API는 요청 1회당 keywordGroups 최대 5개까지 허용 → 배치로 나눠서 호출
KEYWORD_BATCHES = [
    ["정성스러운 선물", "진심 선물", "건강 선물", "효도 선물", "비타민 선물"],
    ["센스있는 선물", "실속 선물", "가성비 선물", "명절 선물세트", "감동 선물"],
]


def get_date_range(months_back=12):
    end_date = datetime.today()
    start_date = end_date - timedelta(days=30 * months_back)
    return start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d")


def fetch_search_trend(keywords, start_date, end_date, time_unit="week"):
    headers = {
        "X-Naver-Client-Id": CLIENT_ID,
        "X-Naver-Client-Secret": CLIENT_SECRET,
        "Content-Type": "application/json",
    }
    body = {
        "startDate": start_date,
        "endDate": end_date,
        "timeUnit": time_unit,
        "keywordGroups": [{"groupName": kw, "keywords": [kw]} for kw in keywords],
    }

    print(f"요청 중: {keywords}")
    try:
        response = requests.post(URL, headers=headers, data=json.dumps(body), timeout=15)
    except requests.exceptions.RequestException as e:
        print(f"  [오류] 요청 중 문제가 발생했습니다: {e}")
        return None

    print(f"  응답 수신 완료 (상태 코드: {response.status_code})")

    if response.status_code != 200:
        print(f"  [오류] {response.text[:200]}")
        return None

    return response.json()


def result_to_rows(result):
    rows = []
    if not result or "results" not in result:
        return rows
    for group in result["results"]:
        group_name = group["title"]
        if not group["data"]:
            print(f"  ※ '{group_name}' — 검색량 부족으로 데이터 없음")
            continue
        for point in group["data"]:
            rows.append({
                "키워드그룹": group_name,
                "기간": point["period"],
                "검색관심도(상대값)": point["ratio"],
            })
    return rows


if __name__ == "__main__":
    start_date, end_date = get_date_range(months_back=12)
    print(f"수집 기간: {start_date} ~ {end_date} (주 단위)\n")

    all_rows = []
    for batch in KEYWORD_BATCHES:
        result = fetch_search_trend(batch, start_date, end_date, time_unit="week")
        all_rows.extend(result_to_rows(result))
        time.sleep(0.3)

    new_df = pd.DataFrame(all_rows)

    existing_path = "naver_gift_tone_trend.csv"
    if os.path.exists(existing_path):
        old_df = pd.read_csv(existing_path, encoding="utf-8-sig")
        old_df = old_df[old_df["키워드그룹"] != "성의있는 선물"]  # 빈 데이터였던 그룹 제거
        combined_df = pd.concat([old_df, new_df], ignore_index=True)
    else:
        combined_df = new_df

    combined_df.to_csv(existing_path, index=False, encoding="utf-8-sig")
    print(f"\n저장 완료: {existing_path} (총 {len(combined_df)}행, {combined_df['키워드그룹'].nunique()}개 키워드그룹)")
    print("\n키워드그룹별 행 수:")
    print(combined_df["키워드그룹"].value_counts().to_string())
