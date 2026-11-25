"""Future module contracts shown in the prototype without enabling operations."""

from dataclasses import dataclass


@dataclass(frozen=True)
class Capability:
    key: str
    label: str
    status: str
    description: str


FUTURE_CAPABILITIES = (
    Capability("iot", "IoT", "NOT CONFIGURED", "No device adapters are connected."),
    Capability(
        "cybersecurity",
        "Cybersecurity",
        "DISABLED",
        "No security actions are enabled in the prototype.",
    ),
    Capability(
        "robotics",
        "Robotics",
        "NOT CONFIGURED",
        "No robot or motion-control hardware is required or connected.",
    ),
)
