"""Modules embarqués pour exécuter pubmed_search.py avec le Python intégré.

PyInstaller n'inclut que les modules importés : on importe ici tout ce dont le
script (et ses futures versions probables) a besoin. Ajouter une ligne ici puis
recompiler si une nouvelle version du script importe une autre bibliothèque.
"""
# bibliothèques tierces
import requests  # noqa: F401
import urllib3  # noqa: F401
import certifi  # noqa: F401
import charset_normalizer  # noqa: F401
import idna  # noqa: F401
import openpyxl  # noqa: F401
import openpyxl.cell.cell  # noqa: F401
import openpyxl.styles  # noqa: F401
import openpyxl.utils  # noqa: F401
import pypdf  # noqa: F401

# bibliothèque standard
import argparse, base64, csv, datetime, email, glob, gzip, hashlib, html, html.parser, http.client  # noqa
import io, json, logging, math, pathlib, platform, random, re, runpy, shutil, signal, socket  # noqa
import sqlite3, ssl, string, subprocess, tarfile, tempfile, textwrap, threading, time  # noqa
import unicodedata, urllib.parse, urllib.request, uuid, zipfile, zlib  # noqa
import xml.etree.ElementTree, xml.dom.minidom  # noqa
import collections, concurrent.futures, dataclasses, functools, itertools, operator, typing  # noqa
