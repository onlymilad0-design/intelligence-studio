#!/usr/bin/env bash
# Installs the machine-level tools (Headroom, OmniRoute).
# The Claude Code plugins (claude-mem, claude-code-setup) and the
# task-observer skill are declared in .claude/ and load with the project.
set -euo pipefail
pip install -U "headroom-ai[code]"   # Headroom: context compression proxy
npm install -g omniroute             # OmniRoute: local AI gateway (http://localhost:20128)
echo "Done. Try:  headroom wrap claude   |   OMNIROUTE_SERVER_HOST=127.0.0.1 omniroute serve"
