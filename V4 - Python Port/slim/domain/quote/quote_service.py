from slim.domain.quote.quote import Quote
from slim.domain.quote.quote_cache import QuoteCache
from slim.domain.quote.quote_repository import QuoteRepository


class QuoteService:
    """Sole public API for the quote domain. Reads from cache; writes hit repo then rebuild."""

    def __init__(self, repo: QuoteRepository, cache: QuoteCache) -> None:
        self._repo = repo
        self._cache = cache

    # --- reads (cache) ---

    def all_quotes(self) -> list[Quote]:
        return self._cache.all_customer_quotes()

    def all_default_quotes(self) -> list[Quote]:
        return self._cache.get_default_quotes()

    def get_by_customer_id(self, customer_id: int) -> list[Quote]:
        return self._cache.get_by_customer_id(customer_id)

    def exists_by_customer_id(self, customer_id: int) -> bool:
        return self._cache.exists_by_customer_id(customer_id)

    # --- writes ---

    def create_quote(self, quote: Quote) -> int:
        if not quote.validate():
            raise ValueError("Quote failed validation before persist.")
        new_id = self._repo.insert_quote(quote)
        self._cache.build()
        return new_id

    def create_default_quote(self, quote: Quote) -> int:
        if not quote.validate():
            raise ValueError("Default quote failed validation before persist.")
        new_id = self._repo.insert_default_quote(quote)
        self._cache.build()
        return new_id

    def save_quote(self, quote: Quote) -> None:
        """Update quote header fields only. Line items are not modified."""
        if not quote.validate():
            raise ValueError("Quote failed validation before save.")
        self._repo.update_quote(quote)
        self._cache.build()

    def delete_quote(self, quote_id: int) -> None:
        self._repo.delete_quote(quote_id)
        self._cache.build()

    def delete_default_quote(self, quote_id: int) -> None:
        self._repo.delete_default_quote(quote_id)
        self._cache.build()
