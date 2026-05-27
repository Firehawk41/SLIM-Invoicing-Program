from datetime import date

from slim.domain.analysis.analysis_service import AnalysisService
from slim.domain.sales_order.sales_order_line_item import SalesOrderLineItem
from slim.domain.tr.tr_submission import TRSubmission
from slim.pipeline.line_item_data import LineItemData
from slim.pipeline.priced_group import PricedGroup
from slim.pipeline.sales_order_pricing_engine import SalesOrderPricingEngine


class SalesOrderLineItemBuilder:
    """
    Bridges pricing and entity construction.
    PricedGroup -> LineItemData (analysis ID resolution) -> SalesOrderLineItem.
    """

    def __init__(
        self,
        pricing_engine: SalesOrderPricingEngine,
        analysis_svc: AnalysisService,
    ) -> None:
        self._pricing_engine = pricing_engine
        self._analysis_svc = analysis_svc

    def price_submission(
        self, submission: TRSubmission, today: date | None = None
    ) -> list[PricedGroup]:
        return self._pricing_engine.calculate_pricing_information(submission, today=today)

    def build_from_priced_groups(
        self, groups: list[PricedGroup]
    ) -> list[SalesOrderLineItem]:
        result = []
        for group in groups:
            li_data = self._to_line_item_data(group)
            li = SalesOrderLineItem(
                id=0,
                item=li_data.item,
                description=li_data.description,
                quantity=li_data.quantity,
                rate=li_data.price,
                quote_price_id=li_data.quote_price_id,
                processing_time=li_data.processing_time,
            )
            if not li.validate():
                raise ValueError(
                    f"Line item failed validation: item={li.item!r}, "
                    f"qty={li.quantity}, rate={li.rate}"
                )
            result.append(li)
        return result

    def _to_line_item_data(self, group: PricedGroup) -> LineItemData:
        names = []
        descriptions = []
        for aid in sorted(group.analysis_ids):
            analysis = self._analysis_svc.load_analysis(aid)
            if analysis is None:
                raise ValueError(f"Analysis ID not found: {aid}")
            names.append(analysis.name)
            descriptions.append(analysis.description)

        pt_label = group.processing_time.label
        return LineItemData(
            item=" + ".join(names) + " - " + pt_label,
            description=" + ".join(descriptions) + " - " + pt_label,
            quantity=group.quantity,
            price=group.price,
            quote_price_id=group.quote_price_id,
            processing_time=group.processing_time,
            payment_terms=group.payment_terms,
        )
