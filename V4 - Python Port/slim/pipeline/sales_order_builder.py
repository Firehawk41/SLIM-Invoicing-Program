from datetime import date

from slim.domain.customer.customer_service import CustomerService
from slim.domain.quote.enums import PaymentType
from slim.domain.sales_order.sales_order import SalesOrder
from slim.domain.sales_order.sales_order_service import SalesOrderService
from slim.domain.tr.tr_submission import TRSubmission
from slim.pipeline.sales_order_line_item_builder import SalesOrderLineItemBuilder


class SalesOrderBuilder:
    """
    Top-level pipeline builder.
    TRSubmission -> priced groups -> line items -> SalesOrder -> persisted SalesOrder.
    """

    DEFAULT_PAYMENT_TERMS = 30

    def __init__(
        self,
        line_item_builder: SalesOrderLineItemBuilder,
        customer_svc: CustomerService,
        so_svc: SalesOrderService,
    ) -> None:
        self._line_item_builder = line_item_builder
        self._customer_svc = customer_svc
        self._so_svc = so_svc

    def build_from_submission(
        self, submission: TRSubmission, today: date | None = None
    ) -> SalesOrder:
        priced_groups = self._line_item_builder.price_submission(submission, today=today)
        payment_terms = (
            priced_groups[0].payment_terms
            if priced_groups
            else self.DEFAULT_PAYMENT_TERMS
        )
        line_items = self._line_item_builder.build_from_priced_groups(priced_groups)
        customer_name = self._customer_svc.get_name_by_id(submission.customer_id)

        so = SalesOrder(
            id=0,
            customer_name=customer_name,
            customer_id=submission.customer_id,
            submission_id=submission.id,
            submission_ref=submission.file_name,
            date_received=submission.date_received,
            po_number=submission.po_information,
            payment_type=(
                PaymentType.CREDIT_CARD
                if submission.credit_card_information
                else PaymentType.PO_NUMBER
            ),
            payment_terms=payment_terms,
            requester_name=submission.customer_contact,
            requester_email=submission.results_email_main_first,
            service_date=submission.effective_service_date,
            line_items=tuple(line_items),
        )
        return self._so_svc.create(so)
