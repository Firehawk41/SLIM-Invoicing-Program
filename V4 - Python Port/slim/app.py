from dataclasses import dataclass

from sqlalchemy.orm import Session

from slim.domain.analysis.analysis_service import AnalysisService
from slim.domain.chemical.chemical_service import ChemicalService
from slim.domain.customer.customer_service import CustomerService
from slim.domain.element.element_service import ElementService
from slim.domain.quote.quote_cache import QuoteCache
from slim.domain.quote.quote_repository import QuoteRepository
from slim.domain.quote.quote_service import QuoteService
from slim.domain.sales_order.sales_order_service import SalesOrderService
from slim.domain.tr.tr_form_input_resolver import TRFormInputResolver
from slim.domain.tr.tr_submission_service import TRSubmissionService
from slim.pipeline.sales_order_builder import SalesOrderBuilder
from slim.pipeline.sales_order_line_item_builder import SalesOrderLineItemBuilder
from slim.pipeline.sales_order_pricing_engine import SalesOrderPricingEngine
from slim.pipeline.submission_loader import SubmissionLoader

_DEFAULT_PAYMENT_TERMS = 30


@dataclass(frozen=True)
class App:
    loader: SubmissionLoader
    builder: SalesOrderBuilder


def create_app(session: Session, default_payment_terms: int = _DEFAULT_PAYMENT_TERMS) -> App:
    """Wire all services and return a ready-to-use App.

    Importing this module registers all ORM rows with Base.metadata, so
    Base.metadata.create_all() called before this function will create every table.
    """
    element_svc  = ElementService(session)
    analysis_svc = AnalysisService(session)
    chemical_svc = ChemicalService(session)
    customer_svc = CustomerService(session)
    so_svc       = SalesOrderService(session)

    quote_repo   = QuoteRepository(session)
    quote_cache  = QuoteCache(quote_repo)
    quote_svc    = QuoteService(quote_repo, quote_cache)
    quote_cache.build()  # QuoteCache.__init__ leaves the cache empty; populate it now

    resolver = TRFormInputResolver(customer_svc, chemical_svc)
    tr_svc   = TRSubmissionService(session, chemical_svc, analysis_svc, element_svc, resolver)

    pricing_engine = SalesOrderPricingEngine(quote_svc, default_payment_terms)
    li_builder     = SalesOrderLineItemBuilder(pricing_engine, analysis_svc)
    builder        = SalesOrderBuilder(li_builder, customer_svc, so_svc)
    loader         = SubmissionLoader(tr_svc)

    return App(loader=loader, builder=builder)
