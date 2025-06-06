"""Exceptions thrown"""


class WindowDidNotAppearException(Exception):
    """Main windows didn't show up - possible pop-up window"""


class AttachException(Exception):
    """Error with attaching - connection or session"""


class ActionException(Exception):
    """Error performing action - click, select ..."""


class SapGuiComException(Exception):
    """Error interacting with SAP GUI COM object"""

class ElementNotFoundException(SapGuiComException):
    """GUI element not found by ID"""

# --- НОВЫЙ КОД ---
class PropertyNotFoundException(SapGuiComException):
    """Element does not have the requested property"""

class InvalidElementTypeException(SapGuiComException):
    """Element is not of the expected type for the operation"""
# --- КОНЕЦ НОВОГО КОДА ---


class MenuNotFoundException(ElementNotFoundException):
    """Menu item not found by name"""

class StatusBarException(SapGuiComException):
    """Error reading status bar"""

class TransactionException(Exception):
    """Base exception for transaction errors"""

class TransactionNotFoundError(TransactionException):
    """Transaction code does not exist"""

class AuthorizationError(TransactionException):
    """User not authorized for the transaction or action"""

class ActionBlockedError(TransactionException):
    """Action blocked within the transaction (e.g., locked object)"""

class SapLogonConfigError(Exception):
    """Error related to saplogon.ini configuration"""
class StatusBarAssertionError(SapGuiComException):
     """Raised when status bar content does not match expectations."""
     pass
