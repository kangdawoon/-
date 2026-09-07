# -*- coding: utf-8 -*-
"""
대상별 선물 키워드의 연령대 검색 비중 수집 스크립트
- "상사 선물", "거래처 선물", "부모님 선물", "동료 선물" 등
  '누구에게'가 명시된 선물 키워드를 실제로 어느 연령대가 검색하는지 확인합니다.
- 목적: "30대는 다양한 대상(상사/거래처/부모님/동료 등)에게 선물해야 하는
  상황이 많다"는 주장을 실제 검색 행동 데이터로 검증하기 위함
  (우리가 만든 질문 카테고리가 아닌, 실제 소비자 검색 데이터 기반)

[인증키]
.env 파일의 NAVER_CLIENT_ID / NAVER_CLIENT_SECRET을 그대로 사용합니다.

[실행 방법]
pip install requests python-dotenv
python naver_gift_recipient_demographics.py

[참고] 네이버 데이터랩 연령대 코드
1: 0~12세   2: 13~18세  3: 19~24세  4: 25~29세  5: 30~34세  6: 35~39세
7: 40~44세  8: 45~49세  9: 50~54세  10: 55~60세  11: 60세 이상

※ 키워드 5개 x 연령대(11) = 55회 호출 (성별 구분 없이 전체 기준으로 조회)
"""

import requests
import json
import csv
import time
from datetime import datetime, timedelta

from dotenv import load_dotenv
import os

load_dotenv()

CLIENT_ID = os.getenv("NAVER_CLIENT_ID")
CLIENT_SECRET = os.getenv("NAVER_CLIENT_SECRET")

if not CLIENT_ID or not CLIENT_SECRET:
    raise ValueError(
        ".env 파일에서 NAVER_CLIENT_ID / NAVER_CLIENT_SECRET을 찾을 수 없습니다.\n"
        "같은 폴더에 .env 파일이 있는지, 값이 정확히 들어있는지 확인해주세요."
    )

URL = "https://openapi.naver.com/v1/datalab/search"

AGE_LABELS = {
    "1": "0~12세", "2": "13~18세", "3": "19~24세", "4": "25~29세",
    "5": "30~34세", "6": "35~39세", "7": "40~44세", "8": "45~49세",
    "9": "50~54세", "10": "55~60세", "11": "60세 이상",
}

# '누구에게'가 명시된 선물 키워드 (대상 다양성 검증용)
RECIPIENT_KEYWORDS = ["상사 선물", "거래처 선물", "부모님 선물", "동료 선물", "첫만남 선물"]


def get_date_range(months_back=12):
    end_date = datetime.today()
    start_date = end_date - timedelta(days=30 * months_back)
    return start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d")


def fetch_age_trend(keyword, age_code, start_date, end_date):
    headers = {
        "X-Naver-Client-Id": CLIENT_ID,
        "X-Naver-Client-Secret": CLIENT_SECRET,
        "Content-Type": "application/json",
    }
    body = {
        "startDate": start_date,
        "endDate": end_date,
        "timeUnit": "month",
        "keywordGroups": [{"groupName": keyword, "keywords": [keyword]}],
        "ages": [age_code],
    }

    try:
        response = requests.post(URL, headers=headers, data=json.dumps(body), timeout=15)
    except requests.exceptions.RequestException as e:
        print(f"    [오류] 요청 실패: {e}")
        return None

    if response.status_code != 200:
        print(f"    [오류] 상태 코드: {response.status_code} / {response.text[:200]}")
        return None

    return response.json()


def collect_all():
    start_date, end_date = get_date_range(months_back=12)
    rows = []

    for keyword in RECIPIENT_KEYWORDS:
        for age_code, age_label in AGE_LABELS.items():
            print(f"수집 중: {keyword} / {age_label}")
            result = fetch_age_trend(keyword, age_code, start_date, end_date)
            time.sleep(0.2)

            if result is None or "results" not in result or not result["results"]:
                continue

            data_points = result["results"][0]["data"]
            avg_ratio = sum(p["ratio"] for p in data_points) / len(data_points) if data_points else 0

            rows.append({
                "키워드": keyword,
                "연령대": age_label,
                "평균검색관심도": round(avg_ratio, 3),
            })

    return rows


def save_csv(rows, filename="naver_gift_recipient_demographics.csv"):
    with open(filename, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["키워드", "연령대", "평균검색관심도"])
        writer.writeheader()
        writer.writerows(rows)
    print(f"\n저장 완료: {filename} (총 {len(rows)}행)")


if __name__ == "__main__":
    print(f"대상 키워드: {RECIPIENT_KEYWORDS}")
    print(f"총 예상 호출 수: {len(RECIPIENT_KEYWORDS)} x {len(AGE_LABELS)} = {len(RECIPIENT_KEYWORDS)*len(AGE_LABELS)}회\n")

    rows = collect_all()
    save_csv(rows)
    print("완료되었습니다. naver_gift_recipient_demographics.csv 파일을 확인해주세요.")
