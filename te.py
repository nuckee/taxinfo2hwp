def number_to_korean_amount(number):
    """
    주어진 숫자를 한글로 변환하여 금액 표현과 함께 리턴하는 함수입니다.

    Parameters:
        number (int): 변환할 숫자

    Returns:
        str: 한글로 변환된 금액 표현 (예: "128,600원(금일십이만팔천육백원정)")
    """
    korean_numbers = ["일", "이", "삼", "사", "오", "육", "칠", "팔", "구"]
    korean_units = ["", "십", "백", "천"]
    korean_big_units = ["", "만", "억", "조", "경", "해", "경", "조", "억", "만"]

    number_str = str(number)
    result = []
    length = len(number_str)

    for i, digit in enumerate(number_str):
        num = int(digit)
        unit_index = (length - 1 - i) % 4
        if num != 0:
            if i > 0 and unit_index == 0 and number_str[i - 1] == '1':
                result.append(korean_numbers[0])
            else:
                result.append(korean_numbers[num - 1])
            result.append(korean_units[unit_index])
        if unit_index == 0 and i < length - 1:
            result.append(korean_big_units[(length - 1 - i) // 4])

    return "".join(result) + "원(금" + "".join(result) + "정)"

if __name__ == "__main__":
    input_number = 128600

    # 함수 실행
    korean_amount = number_to_korean_amount(input_number)
    print(korean_amount)

