# Project-WiseFins

## Import

```
    import pandas as pd 
    import json
    import numpy as np
    from fuzzywuzzy import process
    from pathlib import Path
```

## Architecture Nécessaire

* Project-WIseFIns
    * data
        * client_data
            * manual
            * software
        * menu_sales_analysis_pos.xlsx
        * Agribalyse 2023 3.1 with WiseFins categorisation.xlsx
    * data_converted
        * birchstreet
        * checkscm
        * iscala
        * manual
        * transform
    * data_output
        * birchstreet
        * checkscm
        * iscala
        * manual 
    * main.py

** Penssez à modifier les chemins de vos data d'origine dans le main.py ** 

## Remarques

Le temps d'exécution pour certaine fonction peut-être long.

## Faire fonctionner le programme

Pour faire fonctionner ce programme il suffit juste de lancer le main.py en ayant bien modifier les chemins du dossier data