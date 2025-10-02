"""
Map visualization using Python (Issue #72 Phase 3)

Create interactive maps with Seoul district data using folium library.
No Excel dependency - pure Python visualization.
"""

import json
from pathlib import Path
from typing import Dict, List, Optional, Union

import folium
import pandas as pd
from folium.plugins import MarkerCluster

from .location_converter import LocationConverter


class MapVisualizer:
    """
    Create interactive maps for Seoul district data
    """

    # Seoul center coordinates
    SEOUL_CENTER = [37.5665, 126.9780]

    # Seoul district center coordinates (approximate)
    DISTRICT_COORDS = {
        "Seoul Gangnam": [37.5172, 127.0473],
        "Seoul Gangdong": [37.5301, 127.1238],
        "Seoul Gangbuk": [37.6398, 127.0257],
        "Seoul Gangseo": [37.5509, 126.8495],
        "Seoul Gwanak": [37.4784, 126.9516],
        "Seoul Gwangjin": [37.5384, 127.0822],
        "Seoul Guro": [37.4955, 126.8874],
        "Seoul Geumcheon": [37.4519, 126.8955],
        "Seoul Nowon": [37.6542, 127.0568],
        "Seoul Dobong": [37.6688, 127.0471],
        "Seoul Dongdaemun": [37.5744, 127.0399],
        "Seoul Dongjak": [37.5124, 126.9393],
        "Seoul Mapo": [37.5663, 126.9019],
        "Seoul Seodaemun": [37.5791, 126.9368],
        "Seoul Seocho": [37.4837, 127.0324],
        "Seoul Seongdong": [37.5634, 127.0371],
        "Seoul Seongbuk": [37.5894, 127.0167],
        "Seoul Songpa": [37.5145, 127.1059],
        "Seoul Yangcheon": [37.5170, 126.8664],
        "Seoul Yeongdeungpo": [37.5264, 126.8962],
        "Seoul Yongsan": [37.5326, 126.9904],
        "Seoul Eunpyeong": [37.6176, 126.9227],
        "Seoul Jongno": [37.5730, 126.9794],
        "Seoul Jung": [37.5636, 126.9976],
        "Seoul Jungnang": [37.6063, 127.0925],
    }

    def __init__(self):
        self.converter = LocationConverter()

    def create_choropleth_map(
        self,
        data: Union[Dict[str, float], pd.DataFrame],
        value_column: Optional[str] = None,
        location_column: str = "location",
        output_file: str = "seoul_map.html",
        title: str = "Seoul District Map",
        color_scheme: str = "YlOrRd",
    ) -> str:
        """
        Create choropleth map with district data

        Args:
            data: Dictionary {location: value} or DataFrame with location and value columns
            value_column: Column name for values (if DataFrame)
            location_column: Column name for locations (if DataFrame)
            output_file: Output HTML file path
            title: Map title
            color_scheme: Color scheme (YlOrRd, YlGnBu, RdYlGn, etc.)

        Returns:
            Path to generated HTML file
        """
        # Convert data to dictionary format
        if isinstance(data, pd.DataFrame):
            if value_column is None:
                raise ValueError("value_column required for DataFrame input")
            data_dict = dict(zip(data[location_column], data[value_column]))
        else:
            data_dict = data

        # Normalize location names
        normalized_data = {}
        for location, value in data_dict.items():
            result = self.converter.convert_seoul_district(location)
            if result.matched:
                normalized_data[result.matched] = value
            else:
                # Keep original if no match
                normalized_data[location] = value

        # Create base map
        m = folium.Map(location=self.SEOUL_CENTER, zoom_start=11, tiles="OpenStreetMap")

        # Add title
        title_html = f"""
        <div style="position: fixed;
                    top: 10px; left: 50px; width: 300px; height: 50px;
                    background-color: white; border:2px solid grey; z-index:9999;
                    font-size:20px; font-weight:bold; padding: 10px">
            {title}
        </div>
        """
        m.get_root().html.add_child(folium.Element(title_html))

        # Add circle markers for each district with color based on value
        if normalized_data:
            min_val = min(normalized_data.values())
            max_val = max(normalized_data.values())
            value_range = max_val - min_val if max_val != min_val else 1

            for district, coords in self.DISTRICT_COORDS.items():
                if district in normalized_data:
                    value = normalized_data[district]

                    # Calculate color intensity (0-1)
                    intensity = (value - min_val) / value_range if value_range > 0 else 0.5

                    # Color mapping
                    if color_scheme == "YlOrRd":
                        color = self._get_color_ylord(intensity)
                    elif color_scheme == "YlGnBu":
                        color = self._get_color_ylgnbu(intensity)
                    else:
                        color = self._get_color_ylord(intensity)

                    # Add circle marker
                    folium.CircleMarker(
                        location=coords,
                        radius=15,
                        popup=f"<b>{district}</b><br>Value: {value:,.2f}",
                        tooltip=f"{district}: {value:,.2f}",
                        color=color,
                        fill=True,
                        fillColor=color,
                        fillOpacity=0.7,
                        weight=2,
                    ).add_to(m)

        # Save map
        output_path = Path(output_file)
        m.save(str(output_path))

        return str(output_path.absolute())

    def create_marker_map(
        self,
        data: Union[Dict[str, float], pd.DataFrame],
        value_column: Optional[str] = None,
        location_column: str = "location",
        output_file: str = "seoul_markers.html",
        title: str = "Seoul District Markers",
        use_cluster: bool = False,
    ) -> str:
        """
        Create map with markers for each location

        Args:
            data: Dictionary {location: value} or DataFrame
            value_column: Column name for values (if DataFrame)
            location_column: Column name for locations (if DataFrame)
            output_file: Output HTML file path
            title: Map title
            use_cluster: Use marker clustering

        Returns:
            Path to generated HTML file
        """
        # Convert data to dictionary format
        if isinstance(data, pd.DataFrame):
            if value_column is None:
                raise ValueError("value_column required for DataFrame input")
            data_dict = dict(zip(data[location_column], data[value_column]))
        else:
            data_dict = data

        # Normalize location names
        normalized_data = {}
        for location, value in data_dict.items():
            result = self.converter.convert_seoul_district(location)
            if result.matched:
                normalized_data[result.matched] = value

        # Create base map
        m = folium.Map(location=self.SEOUL_CENTER, zoom_start=11)

        # Add title
        title_html = f"""
        <div style="position: fixed;
                    top: 10px; left: 50px; width: 300px; height: 50px;
                    background-color: white; border:2px solid grey; z-index:9999;
                    font-size:20px; font-weight:bold; padding: 10px">
            {title}
        </div>
        """
        m.get_root().html.add_child(folium.Element(title_html))

        # Create marker cluster if requested
        marker_cluster = MarkerCluster() if use_cluster else None

        # Add markers
        for district, coords in self.DISTRICT_COORDS.items():
            if district in normalized_data:
                value = normalized_data[district]

                marker = folium.Marker(
                    location=coords,
                    popup=f"<b>{district}</b><br>Value: {value:,.2f}",
                    tooltip=f"{district}: {value:,.2f}",
                    icon=folium.Icon(color="blue", icon="info-sign"),
                )

                if marker_cluster:
                    marker.add_to(marker_cluster)
                else:
                    marker.add_to(m)

        if marker_cluster:
            marker_cluster.add_to(m)

        # Save map
        output_path = Path(output_file)
        m.save(str(output_path))

        return str(output_path.absolute())

    def _get_color_ylord(self, intensity: float) -> str:
        """Yellow-Orange-Red color gradient"""
        if intensity < 0.25:
            return "#FFEDA0"  # Light Yellow
        elif intensity < 0.5:
            return "#FED976"  # Yellow
        elif intensity < 0.75:
            return "#FD8D3C"  # Orange
        else:
            return "#E31A1C"  # Red

    def _get_color_ylgnbu(self, intensity: float) -> str:
        """Yellow-Green-Blue color gradient"""
        if intensity < 0.25:
            return "#FFFFCC"  # Light Yellow
        elif intensity < 0.5:
            return "#A1DAB4"  # Light Green
        elif intensity < 0.75:
            return "#41B6C4"  # Cyan
        else:
            return "#225EA8"  # Blue

    def get_supported_districts(self) -> List[str]:
        """Get list of supported district names"""
        return list(self.DISTRICT_COORDS.keys())

    def validate_data(self, data: Union[Dict[str, float], pd.DataFrame]) -> Dict[str, any]:
        """
        Validate input data and provide suggestions

        Returns:
            Validation report with matched/unmatched locations
        """
        if isinstance(data, pd.DataFrame):
            locations = data.iloc[:, 0].tolist()
        else:
            locations = list(data.keys())

        matched = []
        unmatched = []

        for location in locations:
            result = self.converter.convert_seoul_district(location)
            if result.matched:
                matched.append({"input": location, "matched": result.matched})
            else:
                unmatched.append({"input": location, "suggestions": result.suggestions[:3]})

        return {
            "total_locations": len(locations),
            "matched_count": len(matched),
            "unmatched_count": len(unmatched),
            "matched": matched,
            "unmatched": unmatched,
        }
