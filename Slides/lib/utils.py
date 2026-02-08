"""Shared utility functions"""

from pathlib import Path


def find_path(name, max_levels=3):
    """Find path in current or parent directories"""
    for level in range(max_levels):
        path = Path("../" * level + name)
        if path.exists():
            return path
    return Path(name)

