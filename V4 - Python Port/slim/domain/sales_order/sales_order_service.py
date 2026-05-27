from datetime import date

from sqlalchemy.orm import Session

from slim.domain.sales_order.sales_order import SalesOrder
from slim.domain.sales_order.sales_order_repository import SalesOrderRepository


class SalesOrderService:
    """Sole public API for the sales order domain. No cache — SOs are transactional records."""

    def __init__(self, session: Session) -> None:
        self._repo = SalesOrderRepository(session)

    def create(self, so: SalesOrder) -> SalesOrder:
        """Validate, persist, and return a sealed SalesOrder with DB-assigned IDs."""
        if not so.validate():
            raise ValueError("Sales order failed validation before persist.")
        return self._repo.insert(so)

    def load(self, sales_order_id: int) -> SalesOrder | None:
        return self._repo.load(sales_order_id)

    def load_by_date_range(self, start: date, end: date) -> list[SalesOrder]:
        return self._repo.load_by_date_range(start, end)

    def save(self, so: SalesOrder) -> None:
        """Update header fields only. Line items are not modified."""
        if not so.validate():
            raise ValueError("Sales order failed validation before save.")
        self._repo.update(so)

    def delete(self, sales_order_id: int) -> None:
        self._repo.delete(sales_order_id)
