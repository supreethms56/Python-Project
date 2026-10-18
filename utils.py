"""
utils.py: Helper functions and calculation logic.
Fulfills:
- AC-4: Calculate Total Amount (Price * Quantity).
- AC-6: Apply 10% discount when Total Amount exceeds $100.
- AC-9 & AC-10: Error detection and logging payload formation without crashing.
- AC-12: Separation of concerns.
"""

from typing import Any, Optional, Tuple


def calculate_transaction_totals(
    price: Any, quantity: Any
) -> Tuple[Optional[float], Optional[float], Optional[str]]:
    """
    Validates numeric price and quantity, then calculates Total Amount and Discounted Total.

    Returns:
        (total_amount, discounted_total, error_message)
    """
    # AC-9 & AC-10: Catch missing fields
    if price is None or quantity is None:
        return None, None, "Missing Price or Quantity value."

    # Validate numeric types
    try:
        clean_price = float(price)
        clean_quantity = float(quantity)
    except (ValueError, TypeError):
        return None, None, f"Non-numeric data found (Price: '{price}', Quantity: '{quantity}')."

    # Validate business rules (non-negative price, positive non-zero quantity)
    if clean_price < 0:
        return None, None, f"Abnormal negative price encountered: {clean_price}."
    if clean_quantity <= 0:
        return None, None, f"Abnormal quantity encountered (must be > 0): {clean_quantity}."

    # AC-4: Total Amount as Price multiplied by Quantity
    total_amount = round(clean_price * clean_quantity, 2)

    # AC-6: 10% loyalty discount when Total Amount exceeds $100
    if total_amount > 100.0:
        discounted_total = round(total_amount * 0.90, 2)
    else:
        discounted_total = total_amount

    return total_amount, discounted_total, None