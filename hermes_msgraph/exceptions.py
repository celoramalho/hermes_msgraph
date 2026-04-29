
class HermesMSGraphError(Exception):
    """Custom exception for errors related to HermesGraphAPI."""
    def __init__(self, message: str, error_code: int = None):
        super().__init__(message)
        self.error_code = error_code

    def __str__(self):
        if self.error_code:
            return f"[Error {self.error_code}] {self.args[0]}"
        return self.args[0]