# Predictive Fault Weather Engine v2

Ανέβασε στο GitHub ΜΟΝΟ:
- app.py
- requirements.txt

Μετά στο Streamlit Cloud κάνε Deploy με:
- Branch: main
- Main file path: app.py

Νέα v2:
- Default scope: Μόνο Βλάβες
- Dropdown για Όλα / GroupA / GroupB
- Dropdown περιοχής καιρού
- Adjustable thresholds για βροχή, άνεμο, stock recovery
- Σύγκριση κακοκαιρίας με εισροή/stock
- Πρόγνωση 7 ημερών με LOW/MEDIUM/HIGH/CRITICAL

Σημείωση:
Το αρχείο Chart.v4.xlsb περιέχει weighted split ανά GroupA/GroupB. Η εφαρμογή χρησιμοποιεί αυτά τα βάρη για να απομονώσει τις βλάβες από το συνολικό ημερήσιο pivot.
