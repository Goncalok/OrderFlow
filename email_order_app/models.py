from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime


@dataclass(slots=True)
class OrderItem:
    article_number: str
    description: str
    quantity_boxes: float
    unit: str


@dataclass(slots=True)
class CustomerOrder:
    customer_dc: str
    fatrans_dc: str
    reference: str
    items: list[OrderItem]
    slotboeking_id: str | None = None


@dataclass(slots=True)
class ParsedOrderEmail:
    source_file: str
    subject: str
    sender: str
    received_at: datetime | None
    delivery_date_to_dc: str
    leaving_date_venlo: str
    leaving_time_venlo: str
    orders: list[CustomerOrder]
