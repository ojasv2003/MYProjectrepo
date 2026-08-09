"""Master CSV template columns and Pydantic extraction schema for Cursor SDK."""

from __future__ import annotations

from typing import Optional

from pydantic import BaseModel, Field

TEMPLATE_COLUMNS: list[str] = [
    "Hub_Name",
    "Hub_Type",
    "Location",
    "State_Region",
    "Geographical_Zone",
    "Area_In_Acres",
    "Target_Capacity_TEU_Annum",
    "Capacity_Utilization_Pct",
    "Utilized_Volume_TEU_Annual",
    "Investment_Cost_Crores",
    "Development_Phase",
    "Connectivity_Modes",
    "Connectivity_Matrix",
    "Rail_Physical_Destinations",
    "Road_Physical_Destinations",
    "Port_Physical_Destinations",
    "Airport_Physical_Destinations",
    "Outbound_Rail_Pct",
    "Outbound_Road_Pct",
    "Outbound_Port_Pct",
    "Outbound_Airport_Pct",
    "Heavy_Machinery_Fleet",
    "Catchment_Score",
    "Avg_Freight_Speed_KmH",
    "Real_Anchor_Industries",
    "Industry_Driving_Distance_KM",
    "Pipeline_Tracker",
    "Regulatory_Layer",
    "Transaction_Proxies_Bilty",
    "Data_Source_Verification_Link",
]


class LogisticsHubRecord(BaseModel):
    """One Indian multi-modal logistics hub / park row matching the master CSV."""

    Hub_Name: Optional[str] = Field(
        default=None,
        description=(
            "Official name of the multi-modal logistics hub, park, or MMLP "
            "(e.g. 'Multi Modal Logistics Park Nagpur')."
        ),
    )
    Hub_Type: Optional[str] = Field(
        default=None,
        description=(
            "Facility classification such as MMLP, MMLH, ICD, CFS, dry port, "
            "freight village, or integrated logistics hub."
        ),
    )
    Location: Optional[str] = Field(
        default=None,
        description=(
            "City, district, or site locality where the hub is located "
            "(e.g. 'Nagpur', 'Nangal Chaudhary')."
        ),
    )
    State_Region: Optional[str] = Field(
        default=None,
        description="Indian state or union territory (e.g. 'Maharashtra', 'Haryana').",
    )
    Geographical_Zone: Optional[str] = Field(
        default=None,
        description=(
            "Broader geographic/planning zone if stated (North, South, East, West, "
            "Central, or a named economic corridor)."
        ),
    )
    Area_In_Acres: Optional[str] = Field(
        default=None,
        description="Total land area of the hub in acres (numeric string if known).",
    )
    Target_Capacity_TEU_Annum: Optional[str] = Field(
        default=None,
        description=(
            "Designed / target annual throughput capacity in TEU (or equivalent "
            "capacity metric if TEU is not stated)."
        ),
    )
    Capacity_Utilization_Pct: Optional[str] = Field(
        default=None,
        description="Current capacity utilization percentage, if reported.",
    )
    Utilized_Volume_TEU_Annual: Optional[str] = Field(
        default=None,
        description="Actual annual utilized / handled volume in TEU, if reported.",
    )
    Investment_Cost_Crores: Optional[str] = Field(
        default=None,
        description=(
            "Project investment / project cost in INR crores. Include currency "
            "context only if needed; prefer a numeric string."
        ),
    )
    Development_Phase: Optional[str] = Field(
        default=None,
        description=(
            "Status / phase such as Proposed, Under Construction, Operational, "
            "Commissioned, or Phase-1 / Phase-2. Include commissioning / "
            "commercial-operations year when stated (e.g. 'Operational; "
            "commissioned 2024'). Also note page last-updated or data as-of date "
            "if the source mentions one."
        ),
    )
    Connectivity_Modes: Optional[str] = Field(
        default=None,
        description=(
            "Comma-separated transport modes available (rail, road, port/sea, air)."
        ),
    )
    Connectivity_Matrix: Optional[str] = Field(
        default=None,
        description=(
            "Short summary of how modes interconnect at the hub (e.g. rail siding "
            "linked to NH corridor and nearest seaport)."
        ),
    )
    Rail_Physical_Destinations: Optional[str] = Field(
        default=None,
        description=(
            "Named rail destinations, corridors, or terminals served / planned."
        ),
    )
    Road_Physical_Destinations: Optional[str] = Field(
        default=None,
        description=(
            "Named highways, NH/SH links, or road destinations connected to the hub."
        ),
    )
    Port_Physical_Destinations: Optional[str] = Field(
        default=None,
        description="Named seaports / inland waterway ports linked to the hub.",
    )
    Airport_Physical_Destinations: Optional[str] = Field(
        default=None,
        description="Named airports or air-cargo facilities linked to the hub.",
    )
    Outbound_Rail_Pct: Optional[str] = Field(
        default=None,
        description="Share of outbound freight moving by rail (%), if stated.",
    )
    Outbound_Road_Pct: Optional[str] = Field(
        default=None,
        description="Share of outbound freight moving by road (%), if stated.",
    )
    Outbound_Port_Pct: Optional[str] = Field(
        default=None,
        description="Share of outbound freight moving via port/sea (%), if stated.",
    )
    Outbound_Airport_Pct: Optional[str] = Field(
        default=None,
        description="Share of outbound freight moving by air (%), if stated.",
    )
    Heavy_Machinery_Fleet: Optional[str] = Field(
        default=None,
        description=(
            "Material-handling equipment / fleet notes (RTGs, reach stackers, "
            "cranes, warehouses, cold chain, etc.)."
        ),
    )
    Catchment_Score: Optional[str] = Field(
        default=None,
        description=(
            "Catchment / hinterland strength indicator or qualitative score if given."
        ),
    )
    Avg_Freight_Speed_KmH: Optional[str] = Field(
        default=None,
        description="Average freight movement speed in km/h, if reported.",
    )
    Real_Anchor_Industries: Optional[str] = Field(
        default=None,
        description=(
            "Anchor industries or major cargo types (auto, steel, FMCG, pharma, "
            "agriculture, e-commerce, containers, etc.). Prefix with operator/"
            "owner when known (e.g. 'Operator: NHLML / NHAI; anchors: auto, FMCG')."
        ),
    )
    Industry_Driving_Distance_KM: Optional[str] = Field(
        default=None,
        description=(
            "Typical driving distance (km) from key industrial clusters to the hub."
        ),
    )
    Pipeline_Tracker: Optional[str] = Field(
        default=None,
        description=(
            "Upcoming phases, expansion plans, or pipeline milestones for the hub."
        ),
    )
    Regulatory_Layer: Optional[str] = Field(
        default=None,
        description=(
            "Policy / regulatory context (PM Gati Shakti, Bharatmala, NICDC, SEZ, "
            "customs notified area, PPP framework, etc.)."
        ),
    )
    Transaction_Proxies_Bilty: Optional[str] = Field(
        default=None,
        description=(
            "Any transaction, bilty, EXIM, customs, or throughput proxy metrics "
            "mentioned on the page."
        ),
    )
    Data_Source_Verification_Link: Optional[str] = Field(
        default=None,
        description=(
            "Canonical source URL used for this extraction. Prefer the page URL "
            "that was fetched. Optionally append '; as-of YYYY-MM-DD' when the "
            "page states a last-updated or publication date."
        ),
    )


class ExtractionResult(BaseModel):
    """Structured-output wrapper: zero or more hub rows from one page."""

    hubs: list[LogisticsHubRecord] = Field(
        default_factory=list,
        description=(
            "List of distinct Indian multi-modal logistics hubs / MMLPs found in "
            "the source text. Return an empty list if the page has no usable hub data."
        ),
    )
