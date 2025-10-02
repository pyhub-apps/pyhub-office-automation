"""
Location name converter for Excel Map Chart (Issue #72)

Converts Korean location names to Excel-recognizable formats:
- Seoul districts: "gangnamgu" -> "Seoul City gangnamgu"
- Other regions: Support for nationwide administrative divisions  
- English names: "Gangnam-gu" -> "Seoul City gangnamgu"

Supports validation and provides user-friendly suggestions for incorrect formats.
"""

from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple


@dataclass
class LocationMatchResult:
    """Location name matching result"""

    original: str
    matched: Optional[str]
    status: str  # "exact", "converted", "fuzzy", "not_found"
    confidence: float  # 0.0 ~ 1.0
    suggestions: List[str]


# Seoul 25 Districts - Full List
# Format recognized by Excel Map Chart: "Seoul [District Name]"
SEOUL_DISTRICTS_KR = {
    # Korean district name (without "구") -> Excel format
    "강남": "Seoul Gangnam",
    "강동": "Seoul Gangdong",
    "강북": "Seoul Gangbuk",
    "강서": "Seoul Gangseo",
    "관악": "Seoul Gwanak",
    "광진": "Seoul Gwangjin",
    "구로": "Seoul Guro",
    "금천": "Seoul Geumcheon",
    "노원": "Seoul Nowon",
    "도봉": "Seoul Dobong",
    "동대문": "Seoul Dongdaemun",
    "동작": "Seoul Dongjak",
    "마포": "Seoul Mapo",
    "서대문": "Seoul Seodaemun",
    "서초": "Seoul Seocho",
    "성동": "Seoul Seongdong",
    "성북": "Seoul Seongbuk",
    "송파": "Seoul Songpa",
    "양천": "Seoul Yangcheon",
    "영등포": "Seoul Yeongdeungpo",
    "용산": "Seoul Yongsan",
    "은평": "Seoul Eunpyeong",
    "종로": "Seoul Jongno",
    "중": "Seoul Jung",
    "중랑": "Seoul Jungnang",
}

# English romanization variants
SEOUL_DISTRICTS_EN = {
    "gangnam": "Seoul Gangnam",
    "gangdong": "Seoul Gangdong",
    "gangbuk": "Seoul Gangbuk",
    "gangseo": "Seoul Gangseo",
    "gwanak": "Seoul Gwanak",
    "gwangjin": "Seoul Gwangjin",
    "guro": "Seoul Guro",
    "geumcheon": "Seoul Geumcheon",
    "nowon": "Seoul Nowon",
    "dobong": "Seoul Dobong",
    "dongdaemun": "Seoul Dongdaemun",
    "dongjak": "Seoul Dongjak",
    "mapo": "Seoul Mapo",
    "seodaemun": "Seoul Seodaemun",
    "seocho": "Seoul Seocho",
    "seongdong": "Seoul Seongdong",
    "seongbuk": "Seoul Seongbuk",
    "songpa": "Seoul Songpa",
    "yangcheon": "Seoul Yangcheon",
    "yeongdeungpo": "Seoul Yeongdeungpo",
    "yongsan": "Seoul Yongsan",
    "eunpyeong": "Seoul Eunpyeong",
    "jongno": "Seoul Jongno",
    "jung": "Seoul Jung",
    "jungnang": "Seoul Jungnang",
}

# Common variations and patterns
SEOUL_VARIATIONS = {
    # With "구" suffix
    "강남구": "강남",
    "강동구": "강동",
    "강북구": "강북",
    "강서구": "강서",
    "관악구": "관악",
    "광진구": "광진",
    "구로구": "구로",
    "금천구": "금천",
    "노원구": "노원",
    "도봉구": "도봉",
    "동대문구": "동대문",
    "동작구": "동작",
    "마포구": "마포",
    "서대문구": "서대문",
    "서초구": "서초",
    "성동구": "성동",
    "성북구": "성북",
    "송파구": "송파",
    "양천구": "양천",
    "영등포구": "영등포",
    "용산구": "용산",
    "은평구": "은평",
    "종로구": "종로",
    "중구": "중",
    "중랑구": "중랑",
    # English with -gu suffix
    "gangnam-gu": "gangnam",
    "gangdong-gu": "gangdong",
    "gangbuk-gu": "gangbuk",
    "gangseo-gu": "gangseo",
    "gwanak-gu": "gwanak",
    "gwangjin-gu": "gwangjin",
    "guro-gu": "guro",
    "geumcheon-gu": "geumcheon",
    "nowon-gu": "nowon",
    "dobong-gu": "dobong",
    "dongdaemun-gu": "dongdaemun",
    "dongjak-gu": "dongjak",
    "mapo-gu": "mapo",
    "seodaemun-gu": "seodaemun",
    "seocho-gu": "seocho",
    "seongdong-gu": "seongdong",
    "seongbuk-gu": "seongbuk",
    "songpa-gu": "songpa",
    "yangcheon-gu": "yangcheon",
    "yeongdeungpo-gu": "yeongdeungpo",
    "yongsan-gu": "yongsan",
    "eunpyeong-gu": "eunpyeong",
    "jongno-gu": "jongno",
    "jung-gu": "jung",
    "jungnang-gu": "jungnang",
    # With "서울" prefix
    "서울 강남": "강남",
    "서울 강동": "강동",
    "서울 강북": "강북",
    "서울 강서": "강서",
    "서울 관악": "관악",
    "서울 광진": "광진",
    "서울 구로": "구로",
    "서울 금천": "금천",
    "서울 노원": "노원",
    "서울 도봉": "도봉",
    "서울 동대문": "동대문",
    "서울 동작": "동작",
    "서울 마포": "마포",
    "서울 서대문": "서대문",
    "서울 서초": "서초",
    "서울 성동": "성동",
    "서울 성북": "성북",
    "서울 송파": "송파",
    "서울 양천": "양천",
    "서울 영등포": "영등포",
    "서울 용산": "용산",
    "서울 은평": "은평",
    "서울 종로": "종로",
    "서울 중": "중",
    "서울 중랑": "중랑",
    "서울 강남구": "강남",
    "서울 강동구": "강동",
    "서울 강북구": "강북",
    "서울 강서구": "강서",
    "서울 관악구": "관악",
    "서울 광진구": "광진",
    "서울 구로구": "구로",
    "서울 금천구": "금천",
    "서울 노원구": "노원",
    "서울 도봉구": "도봉",
    "서울 동대문구": "동대문",
    "서울 동작구": "동작",
    "서울 마포구": "마포",
    "서울 서대문구": "서대문",
    "서울 서초구": "서초",
    "서울 성동구": "성동",
    "서울 성북구": "성북",
    "서울 송파구": "송파",
    "서울 양천구": "양천",
    "서울 영등포구": "영등포",
    "서울 용산구": "용산",
    "서울 은평구": "은평",
    "서울 종로구": "종로",
    "서울 중구": "중",
    "서울 중랑구": "중랑",
}


class LocationConverter:
    """
    Convert location names to Excel Map Chart compatible format
    """

    def __init__(self):
        self.seoul_kr = SEOUL_DISTRICTS_KR
        self.seoul_en = SEOUL_DISTRICTS_EN
        self.variations = SEOUL_VARIATIONS

    def convert_seoul_district(self, location: str) -> LocationMatchResult:
        """
        Convert Seoul district name to Excel-compatible format

        Args:
            location: Location name (e.g., "강남구", "gangnam", "Gangnam-gu")

        Returns:
            LocationMatchResult with conversion status and suggestions
        """
        location_clean = location.strip()

        # 1. Check if already in Excel format (exact match)
        all_excel_formats = set(self.seoul_kr.values()) | set(self.seoul_en.values())
        if location_clean in all_excel_formats:
            return LocationMatchResult(
                original=location,
                matched=location_clean,
                status="exact",
                confidence=1.0,
                suggestions=[],
            )

        # 2. Direct Korean lookup (without "구")
        if location_clean in self.seoul_kr:
            matched = self.seoul_kr[location_clean]
            return LocationMatchResult(
                original=location,
                matched=matched,
                status="converted",
                confidence=1.0,
                suggestions=[matched],
            )

        # 3. Direct English lookup (lowercase)
        location_lower = location_clean.lower()
        if location_lower in self.seoul_en:
            matched = self.seoul_en[location_lower]
            return LocationMatchResult(
                original=location,
                matched=matched,
                status="converted",
                confidence=1.0,
                suggestions=[matched],
            )

        # 4. Variation lookup (with "구", "서울" prefix, etc.)
        # Check both exact case and lowercase for English variations
        if location_clean in self.variations:
            intermediate = self.variations[location_clean]
        elif location_lower in self.variations:
            intermediate = self.variations[location_lower]
        else:
            intermediate = None

        if intermediate:
            # Convert intermediate to final Excel format
            if intermediate in self.seoul_kr:
                matched = self.seoul_kr[intermediate]
            elif intermediate in self.seoul_en:
                matched = self.seoul_en[intermediate]
            else:
                matched = None

            if matched:
                return LocationMatchResult(
                    original=location,
                    matched=matched,
                    status="converted",
                    confidence=0.95,
                    suggestions=[matched],
                )

        # 5. Fuzzy matching (contains check)
        fuzzy_matches = self._fuzzy_match(location_clean)
        if fuzzy_matches:
            return LocationMatchResult(
                original=location,
                matched=None,
                status="fuzzy",
                confidence=0.5,
                suggestions=fuzzy_matches[:3],
            )

        # 6. Not found
        return LocationMatchResult(
            original=location,
            matched=None,
            status="not_found",
            confidence=0.0,
            suggestions=list(self.seoul_kr.values())[:5],
        )

    def _fuzzy_match(self, location: str) -> List[str]:
        """
        Find fuzzy matches for location name

        Args:
            location: Input location name

        Returns:
            List of similar Excel-format district names
        """
        matches = []
        location_lower = location.lower()

        # Check Korean districts
        for kr_name, excel_format in self.seoul_kr.items():
            if kr_name in location or location in kr_name:
                if excel_format not in matches:
                    matches.append(excel_format)
                    if len(matches) >= 5:
                        break

        # Check English districts
        for en_name, excel_format in self.seoul_en.items():
            if en_name in location_lower or location_lower in en_name:
                if excel_format not in matches:
                    matches.append(excel_format)
                    if len(matches) >= 5:
                        break

        return matches

    def convert_batch(self, locations: List[str]) -> Dict[str, LocationMatchResult]:
        """
        Convert multiple location names in batch

        Args:
            locations: List of location names

        Returns:
            Dictionary mapping original names to conversion results
        """
        results = {}
        for location in locations:
            results[location] = self.convert_seoul_district(location)
        return results

    def validate_data_range(self, locations: List[str]) -> Tuple[bool, List[str], List[LocationMatchResult]]:
        """
        Validate a list of location names for Excel Map Chart

        Args:
            locations: List of location names from data range

        Returns:
            Tuple of (all_valid, converted_locations, problematic_results)
        """
        results = self.convert_batch(locations)
        converted = []
        problematic = []

        for location, result in results.items():
            if result.status in ["exact", "converted"]:
                converted.append(result.matched)
            else:
                converted.append(location)
                problematic.append(result)

        all_valid = len(problematic) == 0
        return all_valid, converted, problematic

    def get_guidance(self, region: str = "seoul") -> Dict[str, any]:
        """
        Get location name format guidance

        Args:
            region: Region name (currently only "seoul" supported)

        Returns:
            Dictionary with format examples and tips
        """
        if region.lower() == "seoul":
            return {
                "region": "Seoul",
                "total_districts": 25,
                "correct_formats": [
                    "Seoul Gangnam",
                    "Seoul Seocho",
                    "Seoul Jongno",
                ],
                "accepted_inputs": [
                    "강남구 → Seoul Gangnam",
                    "gangnam → Seoul Gangnam",
                    "Gangnam-gu → Seoul Gangnam",
                    "서울 강남구 → Seoul Gangnam",
                ],
                "tips": [
                    "Korean district names are automatically converted (e.g., 강남구 → Seoul Gangnam)",
                    "English romanization is supported (e.g., gangnam, seocho)",
                    "Variations with -gu suffix are recognized (e.g., Gangnam-gu)",
                    "Prefix '서울' is optional and will be removed",
                    "Excel Map Chart requires Microsoft 365 subscription",
                    "Internet connection required (Bing Maps integration)",
                ],
                "all_districts": sorted(set(self.seoul_kr.values())),
            }
        else:
            return {
                "error": f"Region '{region}' not supported yet",
                "supported_regions": ["seoul"],
            }
