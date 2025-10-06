import os
import sys

# Sørg for at "src" (to nivå opp) ligger på sys.path når vi kjører som script
BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../"))
if BASE not in sys.path:
    sys.path.insert(0, BASE)

from app.dokumentreader.invoice_reader import main

if __name__ == "__main__":
    sys.exit(main())
