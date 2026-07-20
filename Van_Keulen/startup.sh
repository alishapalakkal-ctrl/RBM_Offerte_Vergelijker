#!/bin/bash
python -m streamlit run offerte_vergelijker_web.py \
    --server.port 8000 \
    --server.address 0.0.0.0 \
    --server.enableCORS false \
    --server.enableXsrfProtection false
