#!/bin/bash
python -m streamlit run Home.py \
    --server.port 8000 \
    --server.address 0.0.0.0 \
    --server.enableCORS false \
    --server.enableXsrfProtection false
