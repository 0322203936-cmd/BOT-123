"""Send the separately configured Wholesale mass announcement."""

from __future__ import annotations

import os

from email_sender import load_email_config, send_mass_announcement_email


def main() -> None:
    config = load_email_config(
        os.environ.get("CAJAS_MASS_EMAIL_CONFIG"),
        "CAJAS_MASS_EMAIL_CONFIG",
    )
    if not config:
        raise RuntimeError(
            "No existe una configuración guardada para el envío masivo."
        )
    send_mass_announcement_email(
        config,
        os.environ.get("MAIL_SENDER", ""),
    )


if __name__ == "__main__":
    main()
