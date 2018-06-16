## Excel Split Screen

Modifier les clées de registres comme suit :

- Appuyez sur `[Windows]+[R]`, saisissez `REGEDIT` et cliquez sur `OK`

### Première séquence

- Déployez la clé `HKEY_CLASSES_ROOT \ Excel.Sheet.12 \ Shell \ Open`
- Supprimez la clé `DDEEXEC` en cliquant dessus du bouton droit et en choisissant `Supprimer`
- Entrez dans clé `COMMAND`
- Remarquez la présence d'une valeur `(par défaut)` et d'une valeur `command`
- Cliquez du bouton droit sur la valeur `command` et choisissez `Supprimer`
- Double-cliquez sur la valeur `(par défaut)`
- Ajoutez un espace puis `%1` (avec les guillemets) en fin de ligne pour que la donnée ressemble à:
```bat
C:\Program Files\Microsoft Office\Office12\EXCEL.EXE" /e "%1"
```

### Deuxième séquence

- Déployez la clé `HKEY_CLASSES_ROOT \ Excel.Sheet.8 \ Shell \ Open`
- Supprimez la clé `DDEEXEC` en cliquant dessus du bouton droit et en choisissant `Supprimer`
- Entrez dans clé `COMMAND`
- Remarquez la présence d'une valeur `(par défaut)` et d'une valeur `command`
- Cliquez du bouton droit sur la valeur "command" et choisissez Supprimer
- Double-cliquez sur la valeur `(par défaut)`
- Ajoutez un espace puis `%1` (avec les guillemets) en fin de ligne pour que la donnée ressemble à:
```bat
"C:\Program Files\Microsoft Office\Office12\EXCEL.EXE" /e "%1"
```

### Finalement

- Fermez `REGEDIT`

Maintenant, si vous double-cliquez sur deux fichiers XLS ou XSLX sur le bureau ou l'explorateur, ils s'ouvriront bien dans deux fenêtres différentes.
