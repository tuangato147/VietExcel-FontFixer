# -*- coding: utf-8 -*-
"""
Vietnamese Encoding Conversion Module
Supports:
- VNI-Windows to Unicode
- TCVN3 (ABC) to Unicode
- Auto-detection based on font name and character statistics
"""

import re

# Complete VNI-Windows to Unicode mapping table
VNI_TO_UNICODE = {
    # 3-char sequences (horn + vowel combinations)
    '\xf6\xf4\xf9': 'ướ',
    '\xf6\xf4\xf8': 'ườ',
    '\xf6\xf4\xfb': 'ưở',
    '\xf6\xf4\xfa': 'ưở',
    '\xf6\xf4\xf5': 'ưỡ',
    '\xf6\xf4\xef': 'ượ',
    '\xf6\xf4': 'ươ',
    '\xd6\xd4\xd9': 'ƯỚ',
    '\xd6\xd4\xd8': 'ƯỜ',
    '\xd6\xd4\xdb': 'ƯỞ',
    '\xd6\xd4\xda': 'ƯỞ',
    '\xd6\xd4\xd5': 'ƯỠ',
    '\xd6\xd4\xdf': 'ƯỢ',
    '\xd6\xd4': 'ƯƠ',

    # Horn combinations (ơ, ư)
    '\xf4\xf9': 'ớ',
    '\xf4\xf8': 'ờ',
    '\xf4\xfb': 'ở',
    '\xf4\xfa': 'ở',
    '\xf4\xf5': 'ỡ',
    '\xf4\xef': 'ợ',
    '\xf4': 'ơ',

    '\xf6\xf9': 'ứ',
    '\xf6\xf8': 'ừ',
    '\xf6\xfb': 'ử',
    '\xf6\xfa': 'ử',
    '\xf6\xf5': 'ữ',
    '\xf6\xef': 'ự',
    '\xf6': 'ư',

    '\xd4\xd9': 'Ớ',
    '\xd4\xd8': 'Ờ',
    '\xd4\xdb': 'Ở',
    '\xd4\xda': 'Ở',
    '\xd4\xd5': 'Ỡ',
    '\xd4\xdf': 'Ợ',
    '\xd4': 'Ơ',

    '\xd6\xd9': 'Ứ',
    '\xd6\xd8': 'Ừ',
    '\xd6\xdb': 'Ử',
    '\xd6\xda': 'Ử',
    '\xd6\xd5': 'Ữ',
    '\xd6\xdf': 'Ự',
    '\xd6': 'Ư',

    # Breve combinations (ă)
    'a\xe9': 'ắ',
    'a\xe8': 'ằ',
    'a\xec': 'ẳ',
    'a\xed': 'ẵ',
    'a\xeb': 'ặ',
    'a\xea': 'ă',
    'A\xc9': 'Ắ',
    'A\xc8': 'Ằ',
    'A\xcc': 'Ẳ',
    'A\xcd': 'Ẵ',
    'A\xcb': 'Ặ',
    'A\xca': 'Ă',

    # Circumflex combinations (â, ê, ô)
    'a\xe1': 'ấ',
    'a\xe0': 'ầ',
    'a\xe5': 'ẩ',
    'a\xe3': 'ẫ',
    'a\xe4': 'ậ',
    'a\xe2': 'â',

    'A\xd9': 'Á',
    'A\xd8': 'À',
    'A\xdb': 'Ả',
    'A\xd5': 'Ã',
    'A\xcf': 'Ạ',
    'A\xc1': 'Ấ',
    'A\xc0': 'Ầ',
    'A\xc5': 'Ẩ',
    'A\xc3': 'Ẫ',
    'A\xc4': 'Ậ',
    'A\xc2': 'Â',

    'e\xe1': 'ế',
    'e\xe0': 'ề',
    'e\xe5': 'ể',
    'e\xe3': 'ễ',
    'e\xe4': 'ệ',
    'e\xe2': 'ê',

    'E\xc1': 'Ế',
    'E\xc0': 'Ề',
    'E\xc5': 'Ể',
    'E\xc3': 'Ễ',
    'E\xc4': 'Ệ',
    'E\xc2': 'Ê',

    'o\xe1': 'ố',
    'o\xe0': 'ồ',
    'o\xe5': 'ổ',
    'o\xe3': 'ỗ',
    'o\xe4': 'ộ',
    'o\xe2': 'ô',

    'O\xc1': 'Ố',
    'O\xc0': 'Ồ',
    'O\xc5': 'Ổ',
    'O\xc3': 'Ỗ',
    'O\xc4': 'Ộ',
    'O\xc2': 'Ô',

    # Standard vowel + acute (sắc)
    'a\xf9': 'á',
    'e\xf9': 'é',
    'o\xf9': 'ó',
    'u\xf9': 'ú',
    'y\xf9': 'ý',

    # Standard vowel + grave (huyền)
    'a\xf8': 'à',
    'e\xf8': 'è',
    'o\xf8': 'ò',
    'u\xf8': 'ù',
    'y\xf8': 'ỳ',

    # Standard vowel + hook (hỏi)
    'a\xfb': 'ả',
    'e\xfb': 'ẻ',
    'o\xfb': 'ỏ',
    'u\xfb': 'ủ',
    'y\xfb': 'ỷ',

    # Standard vowel + tilde (ngã)
    'a\xf5': 'ã',
    'e\xf5': 'ẽ',
    'o\xf5': 'õ',
    'u\xf5': 'ũ',
    'y\xf5': 'ỹ',

    # Standard vowel + dot below (nặng)
    'a\xef': 'ạ',
    'e\xef': 'ẹ',
    'o\xef': 'ọ',
    'u\xef': 'ụ',
    'y\xef': 'ỵ',

    # Capital U + hook / accents
    'U\xdb': 'Ủ',
    'U\xd9': 'Ú',
    'U\xd8': 'Ù',
    'U\xd5': 'Ũ',
    'U\xdf': 'Ụ',

    # Standalone i and d
    '\xf1': 'đ',
    '\xd1': 'Đ',
    '\xec': 'ì',
    '\xed': 'í',
    '\xf2': 'ị',
    '\xe6': 'ỉ',
    '\xcc': 'Ì',
    '\xcd': 'Í',
}

# TCVN3 (ABC) to Unicode mapping table
TCVN3_TO_UNICODE = {
    # Thường
    '\xb8': 'à', '\xb5': 'á', '\xb6': 'ả', '\xb7': 'ã', '\xb9': 'ạ',
    '\xa8': 'ă', '\xbb': 'ằ', '\xbe': 'ắ', '\xbc': 'ẳ', '\xbd': 'ẵ', '\xc6': 'ặ',
    '\xa9': 'â', '\xc7': 'ầ', '\xca': 'ấ', '\xc8': 'ẩ', '\xc9': 'ẫ', '\xcb': 'ậ',
    '\xcc': 'è', '\xd0': 'é', '\xce': 'ẻ', '\xcf': 'ẽ', '\xd1': 'ẹ',
    '\xaa': 'ê', '\xd2': 'ề', '\xd5': 'ế', '\xd3': 'ể', '\xd4': 'ễ', '\xd6': 'ệ',
    '\xd7': 'ì', '\xdd': 'í', '\xd8': 'ỉ', '\xdc': 'ĩ', '\xde': 'ị',
    '\xdf': 'ò', '\xe1': 'ó', '\xe2': 'ỏ', '\xe3': 'õ', '\xe4': 'ọ',
    '\xab': 'ô', '\xe5': 'ồ', '\xe8': 'ố', '\xe6': 'ổ', '\xe7': 'ỗ', '\xe9': 'ộ',
    '\xac': 'ơ', '\xea': 'ờ', '\xed': 'ớ', '\xeb': 'ở', '\xec': 'ỡ', '\xee': 'ợ',
    '\xef': 'ù', '\xf3': 'ú', '\xf1': 'ủ', '\xf2': 'ũ', '\xf4': 'ụ',
    '\xad': 'ư', '\xf5': 'ừ', '\xf8': 'ứ', '\xf6': 'ử', '\xf7': 'ữ', '\xf9': 'ự',
    '\xfa': 'ỳ', '\xfd': 'ý', '\xfb': 'ỷ', '\xfc': 'ỹ', '\xfe': 'ỵ',
    '\xae': 'đ',

    # TCVN3 Chữ Hoa (.VnTimeH)
    '\xa1': 'Ă', '\xa2': 'Â', '\xa3': 'Đ', '\xa4': 'Ê', '\xa5': 'Ô', '\xa6': 'Ơ', '\xa7': 'Ư',
}

VNI_REGEX = re.compile('|'.join(re.escape(k) for k in sorted(VNI_TO_UNICODE.keys(), key=len, reverse=True)))
TCVN3_REGEX = re.compile('|'.join(re.escape(k) for k in sorted(TCVN3_TO_UNICODE.keys(), key=len, reverse=True)))

def convert_vni(text):
    if not isinstance(text, str):
        return text
    return VNI_REGEX.sub(lambda m: VNI_TO_UNICODE[m.group(0)], text)

def convert_tcvn3(text):
    if not isinstance(text, str):
        return text
    return TCVN3_REGEX.sub(lambda m: TCVN3_TO_UNICODE[m.group(0)], text)

def detect_font_target(font_name):
    """
    Returns (detected_encoding, target_font_family)
    detected_encoding in ('VNI', 'TCVN3', 'UNICODE')
    """
    if not font_name:
        return 'UNICODE', 'Times New Roman'

    f_upper = font_name.upper().strip()

    # VNI fonts
    if 'VNI' in f_upper:
        if 'HELVE' in f_upper or 'ARIAL' in f_upper:
            return 'VNI', 'Arial'
        return 'VNI', 'Times New Roman'

    # TCVN3 / ABC fonts (.Vn...)
    if f_upper.startswith('.VN') or f_upper.startswith('VN'):
        if 'ARIAL' in f_upper:
            return 'TCVN3', 'Arial'
        return 'TCVN3', 'Times New Roman'

    return 'UNICODE', font_name
