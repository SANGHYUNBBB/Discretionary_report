"""세금 환급은 각종세금의 음수로 표시한다. 금액의 원천/환율은 증권사별 코드에서 결정."""
from decimal import Decimal
import re


def is_tax_refund(transaction):
    text = re.sub(r"\s+", "", str(transaction or ""))
    # 환급 취소나 정정 거래를 일반 환급으로 바꾸지 않는다.
    return ("환급" in text
            and any(word in text for word in ("세금", "원천세", "소득세", "제세", "배당세", "지방세"))
            and not any(word in text for word in ("취소", "정정", "추징", "재징수")))


def first_nonzero(*values):
    return next((value for value in values if value != 0), Decimal("0"))


def negative_refund(amount):
    if not amount:
        raise ValueError("세금 환급액을 원본 금액 열에서 확인할 수 없습니다.")
    # 원본이 이미 음수인 경우에도 양수로 반전시키지 않는다.
    return -abs(amount)
