#!/bin/bash
# Azure App Service startup command.
# Set this exact line as the "Startup Command" in Azure Portal:
#   Configuration -> General settings -> Startup Command
#
#   bash startup.sh
#
# Azure's Python builder (Oryx) already runs `pip install -r requirements.txt`
# during deployment, so this script just launches the app on the port Azure
# expects (8000, matched by the WEBSITES_PORT app setting -- see README.md).

python -m streamlit run app.py --server.port 8000 --server.address 0.0.0.0
