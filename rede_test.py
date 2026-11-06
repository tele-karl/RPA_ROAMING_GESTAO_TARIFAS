import os
from dotenv import load_dotenv
load_dotenv()

print("REDE_COMFONE:", os.getenv("REDE_COMFONE"))
print("PULSEFORCE_FILE:", os.getenv("PULSEFORCE_FILE"))
REDE_COMFONE = os.getenv("REDE_COMFONE")
import os
full = os.path.join(REDE_COMFONE, os.getenv("PULSEFORCE_FILE"))
print("Full path:", full)
print("Exists?", os.path.exists(full))