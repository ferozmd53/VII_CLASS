# Start the integrated Finvasia/Shoonya Excel terminal.
from pathlib import Path
import sys

workbook = Path(__file__).resolve().with_name("Finvasia_Trade_Terminal_v3.xlsm")
if len(sys.argv) == 1:
    sys.argv.append(str(workbook))

import Finvasia_Core_V3_002  # noqa: F401 - core starts its WebSocket and workers
