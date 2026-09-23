#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Alias entrypoint for app_ead_scoring.py
Ensures both `streamlit run app_lead_scoring.py` and `streamlit run app_ead_scoring.py` execute properly.
"""
import runpy

if __name__ == "__main__":
    runpy.run_path("app_ead_scoring.py")
