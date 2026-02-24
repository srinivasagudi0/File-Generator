from __future__ import annotations

import logging

# Logging is intentionally disabled by default for now to keep the code path simple.
# Set this to True later when detailed logging is needed again.
LOGGING_ENABLED = False


def configure_logging(logger_name: str | None = None, enabled: bool | None = None) -> logging.Logger:
    use_logging = LOGGING_ENABLED if enabled is None else enabled
    name = logger_name if logger_name else __name__
    logger = logging.getLogger(name)

    if use_logging:
        if not logging.getLogger().handlers:
            logging.basicConfig(
                level=logging.INFO,
                format='%(asctime)s %(levelname)s %(name)s - %(message)s',
                filename='app.log',
                filemode='a',
            )
        logger.disabled = False
    else:
        logger.disabled = True
        if not logger.handlers:
            logger.addHandler(logging.NullHandler())

    return logger


def preview_text(value: object, limit: int = 300) -> str:
    text = '' if value is None else str(value)
    text = text.replace('\n', '\\n')
    if len(text) <= limit:
        return text
    return f'{text[:limit]}...'

