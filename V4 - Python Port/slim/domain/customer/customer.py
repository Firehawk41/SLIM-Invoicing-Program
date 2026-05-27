from pydantic import BaseModel, ConfigDict


class Customer(BaseModel):
    model_config = ConfigDict(frozen=True)

    id: int
    name: str
    street_address: str
    city: str
    state: str
    postal_code: str
    country: str

    @property
    def formatted_address(self) -> str:
        """street_address, city [postal_code][, country] — state omitted (matches VBA)."""
        addr = f"{self.street_address}, {self.city}"
        if self.postal_code:
            addr += f" {self.postal_code}"
        if self.country:
            addr += f", {self.country}"
        return addr
