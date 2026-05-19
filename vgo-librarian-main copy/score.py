from collections import OrderedDict
from dataclasses import dataclass
from typing import Pattern, TypedDict

@dataclass
class ScoreVersion:
    title: str
    version: str
    regex: Pattern
    cover: bool
    parts: dict[str, str] # map of slugs to paths

@dataclass
class Binder:
    title: str
    names: str
    parts: dict[str, list[str]] # mapping of score to slug(s)
    is_virtual: bool

class ScoreVersionMetadata(TypedDict):
    regex_str: str
    created_at: str
    last_downloaded: str

class ShelfMetadata(TypedDict):
    sharepoint_path: str
    # TODO: binders_url
    scores: dict[str, OrderedDict[str, ScoreVersionMetadata]]

