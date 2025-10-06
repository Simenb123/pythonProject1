from __future__ import annotations
from abc import ABC, abstractmethod
from typing import Any, Dict, Optional

class DocProfile(ABC):
    """
    Basisklasse for dokumentprofiler. En profil beskriver:
    - hvordan gjenkjenne dokumenttypen (detect)
    - hvilke felter/tabeller som skal ut
    - hvilken normalisert JSON-struktur som returneres
    """
    name: str = "base"
    description: str = "Base profile"

    @abstractmethod
    def detect(self, page1_text: str, full_text: str) -> float:
        """Returner score 0..1 for hvor sannsynlig dokumentet matcher profilen."""
        ...

    @abstractmethod
    def parse(self, path: str, page1_text: str, full_text: str) -> Dict[str, Any]:
        """
        Returner strukturert dict. Hver profil bestemmer sin egen 'schema' (document_type + payload).
        Eksempel:
          {
            "document_type": "invoice",
            "payload": { ... }
          }
        """
        ...

def as_result(doc_type: str, payload: Dict[str, Any]) -> Dict[str, Any]:
    return {"document_type": doc_type, "payload": payload}
