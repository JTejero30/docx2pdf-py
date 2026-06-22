#!/usr/bin/env python3
"""Arranca la API HTTP de conversión .docx -> PDF en la red local.

    python serve.py

Escucha en 0.0.0.0:8000, accesible desde otros equipos de la misma wifi en
http://<IP-de-esta-maquina>:8000/docs
"""
from docx2pdf_py.webserver import run

if __name__ == "__main__":
    run()
