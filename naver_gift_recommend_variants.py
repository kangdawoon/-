# -*- coding: utf-8 -*-
"""
상황별 '선물추천' 키워드의 연령대 검색 비중 수집 스크립트
- "부모님 선물추천", "상사 선물추천" 등 상황마다 다르게 검색되는
  "~선물추천" 계열 키워드를 20~30대가 실제로 얼마나 검색하는지 확인합니다.
- 목적: "매번 다른 상황마다 새로 추천을 찾아봐야 한다"는 것을
  (=매번 고민한다는 것을) 실제 검색 행동으로 보여주기 위함

[인증키]
.env 파일의 NAVER_CLIENT_ID / NAVER_CLIENT_SECRET을 그대로 사용합니다.

[실행 방법]
pip install requests python-dotenv
python naver_gift_recommend_variants.py

[참고] 네이버 데이터랩 연령대 코드
3: 19~24세  4: 25~29세  5: 30~34세  6: 35~39세

※ 키워드 7개 x 연령대(4) = 28회 호출
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
    "3": "19~24세", "4": "25~29세", "5": "30~34세", "6": "35~39세",
}

RECOMMEND_KEYWORDS = [
    "부모님 선물추천", "상사 선물추천", "생일선물 추천",
    "첫만남 선물추천", "지인 선물추천", "센스있는 선물추천", "적당한 선물추천",
]


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

    for keyword in RECOMMEND_KEYWORDS:
        for age_code, age_label in AGE_LABELS.items():
            print(f"수집 중: {keyword} / {age_label}")
            result = fetch_age_trend(keyword, age_code, start_date, end_date)
            time.sleep(0.2)

            if result is None or "results" not in result or not result["results"]:
                print(f"    -> 데이터 없음 (검색량이 너무 적을 수 있음)")
                continue

            data_points = result["results"][0]["data"]
            avg_ratio = sum(p["ratio"] for p in data_points) / len(data_points) if data_points else 0

            rows.append({
                "키워드": keyword,
                "연령대": age_label,
                "평균검색관심도": round(avg_ratio, 3),
            })

    return rows


def save_csv(rows, filename="naver_gift_recommend_variants.csv"):
    with open(filename, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["키워드", "연령대", "평균검색관심도"])
        writer.writeheader()
        writer.writerows(rows)
    print(f"\n저장 완료: {filename} (총 {len(rows)}행)")


if __name__ == "__main__":
    print(f"대상 키워드: {RECOMMEND_KEYWORDS}")
    print(f"총 예상 호출 수: {len(RECOMMEND_KEYWORDS)} x {len(AGE_LABELS)} = {len(RECOMMEND_KEYWORDS)*len(AGE_LABELS)}회\n")

    rows = collect_all()
    save_csv(rows)
    print("완료되었습니다. naver_gift_recommend_variants.csv 파일을 확인해주세요.")
