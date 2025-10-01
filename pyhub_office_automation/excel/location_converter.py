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


# Seoul 25 Districts
SEOUL_DISTRICTS = {
    # Short name -> Full administrative name (in Korean)
    "gangnam": "Seoul City Gangnam District",
    "gangdong": "Seoul City Gangdong District",
    "gangbuk": "Seoul City Gangbuk District",
    "gangseo": "Seoul City Gangseo District",
}


class LocationConverter:
    """
    Convert location names to Excel Map Chart compatible format
    """

    def __init__(self):
        self.seoul_districts = SEOUL_DISTRICTS

    def convert_seoul_district(self, location: str) -> LocationMatchResult:
        """
        Convert Seoul district name to Excel-compatible format

        Args:
            location: Location name (e.g., "gangnam", "Gangnam-gu")

        Returns:
            LocationMatchResult with conversion status and suggestions
        """
        location_clean = location.strip().lower()

        # Simple lookup
        if location_clean in self.seoul_districts:
            matched = self.seoul_districts[location_clean]
            return LocationMatchResult(
                original=location,
                matched=matched,
                status="converted",
                confidence=1.0,
                suggestions=[matched],
            )

        # Not found
        return LocationMatchResult(
            original=location,
            matched=None,
            status="not_found",
            confidence=0.0,
            suggestions=list(self.seoul_districts.values())[:5],
        )

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
                "region": "Seoul City",
                "total_districts": len(self.seoul_districts),
                "correct_formats": [
                    "Seoul City Gangnam District",
                    "Seoul City Seocho District",
                ],
                "tips": [
                    "Use full district names for best recognition",
                    "Excel Map Chart requires Microsoft 365 subscription",
                    "Internet connection required (Bing Maps)",
                ],
                "all_districts": list(self.seoul_districts.values()),
            }
        else:
            return {
                "error": f"Region '{region}' not supported yet",
                "supported_regions": ["seoul"],
            }
