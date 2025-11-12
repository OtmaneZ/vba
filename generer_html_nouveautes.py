#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Convertit NOUVEAUTES_PHASE_2.md en HTML avec style professionnel
"""

import markdown2
import os

# Lire le fichier Markdown
with open('NOUVEAUTES_PHASE_2.md', 'r', encoding='utf-8') as f:
    md_content = f.read()

# Convertir en HTML
html_body = markdown2.markdown(md_content, extras=['tables', 'fenced-code-blocks', 'header-ids'])

# Template HTML avec le même style que GUIDE_UTILISATEUR_COMPLET.html
html_template = '''<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" lang="fr" xml:lang="fr">
<head>
  <meta charset="utf-8" />
  <meta name="generator" content="Python markdown2" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0, user-scalable=yes" />
  <title>Nouveautés et Améliorations - Phase 2</title>
  <style>
    code{{
      white-space: pre-wrap;
    }}
    span.smallcaps{{
      font-variant: small-caps;
    }}
    div.columns{{
      display: flex;
      gap: min(4vw, 1.5em);
    }}
    div.column{{
      flex: auto;
      overflow-x: auto;
    }}
    div.hanging-indent{{
      margin-left: 1.5em;
      text-indent: -1.5em;
    }}
    /* The extra [class] is a hack that increases specificity enough to
       override a similar rule in reveal.js */
    ul.task-list[class]{{
      list-style: none;
    }}
    ul.task-list li input[type="checkbox"] {{
      font-size: inherit;
      width: 0.8em;
      margin: 0 0.8em 0.2em -1.6em;
      vertical-align: middle;
    }}
    .display.math{{
      display: block;
      text-align: center;
      margin: 0.5rem auto;
    }}
    /* Style personnalisé */
    body {{
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif;
      line-height: 1.6;
      max-width: 900px;
      margin: 40px auto;
      padding: 0 20px;
      color: #24292e;
      background-color: #ffffff;
    }}
    h1 {{
      color: #0366d6;
      border-bottom: 3px solid #0366d6;
      padding-bottom: 10px;
      font-size: 2.5em;
    }}
    h2 {{
      color: #0366d6;
      border-bottom: 2px solid #e1e4e8;
      padding-bottom: 8px;
      margin-top: 40px;
      font-size: 2em;
    }}
    h3 {{
      color: #24292e;
      font-size: 1.5em;
      margin-top: 30px;
    }}
    h4 {{
      color: #586069;
      font-size: 1.2em;
    }}
    table {{
      border-collapse: collapse;
      width: 100%;
      margin: 20px 0;
      background-color: #fff;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }}
    th {{
      background-color: #0366d6;
      color: white;
      padding: 12px;
      text-align: left;
      font-weight: 600;
    }}
    td {{
      padding: 10px 12px;
      border-bottom: 1px solid #e1e4e8;
    }}
    tr:hover {{
      background-color: #f6f8fa;
    }}
    code {{
      background-color: #f6f8fa;
      padding: 2px 6px;
      border-radius: 3px;
      font-family: "SFMono-Regular", Consolas, "Liberation Mono", Menlo, monospace;
      font-size: 0.9em;
      color: #e83e8c;
    }}
    pre {{
      background-color: #f6f8fa;
      padding: 16px;
      border-radius: 6px;
      overflow-x: auto;
      border: 1px solid #e1e4e8;
    }}
    pre code {{
      background-color: transparent;
      padding: 0;
      color: #24292e;
    }}
    blockquote {{
      border-left: 4px solid #0366d6;
      padding-left: 20px;
      margin-left: 0;
      color: #586069;
      font-style: italic;
      background-color: #f6f8fa;
      padding: 15px 20px;
      border-radius: 3px;
    }}
    a {{
      color: #0366d6;
      text-decoration: none;
    }}
    a:hover {{
      text-decoration: underline;
    }}
    ul, ol {{
      margin: 15px 0;
      padding-left: 30px;
    }}
    li {{
      margin: 8px 0;
    }}
    hr {{
      border: none;
      border-top: 2px solid #e1e4e8;
      margin: 40px 0;
    }}
    .info-box {{
      background-color: #e7f3ff;
      border-left: 4px solid #0366d6;
      padding: 15px;
      margin: 20px 0;
      border-radius: 4px;
    }}
    .success-box {{
      background-color: #d4edda;
      border-left: 4px solid #28a745;
      padding: 15px;
      margin: 20px 0;
      border-radius: 4px;
    }}
    .warning-box {{
      background-color: #fff3cd;
      border-left: 4px solid #ffc107;
      padding: 15px;
      margin: 20px 0;
      border-radius: 4px;
    }}
    @media print {{
      body {{
        max-width: 100%;
        font-size: 12pt;
      }}
      h1 {{
        font-size: 22pt;
      }}
      h2 {{
        font-size: 18pt;
        page-break-before: always;
      }}
      table {{
        page-break-inside: avoid;
      }}
      pre {{
        page-break-inside: avoid;
      }}
    }}
  </style>
</head>
<body>
{body}
</body>
</html>
'''

# Insérer le contenu
html_final = html_template.format(body=html_body)

# Sauvegarder
output_file = 'NOUVEAUTES_PHASE_2.html'
with open(output_file, 'w', encoding='utf-8') as f:
    f.write(html_final)

print(f"✅ Fichier HTML généré : {output_file}")
print(f"📄 Taille : {len(html_final):,} caractères")
print()
print("📋 Prochaines étapes :")
print("   1. Ouvre NOUVEAUTES_PHASE_2.html dans Safari/Chrome")
print("   2. Cmd+P (Imprimer)")
print("   3. Choisis 'Enregistrer en PDF'")
print("   4. Le PDF est prêt à envoyer à la cliente !")
