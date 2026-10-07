# Listy płac → Excel (dodatek Office)

Panel boczny Excela: upuszczasz listy płac PDF (Comarch ERP Optima „Szczegółowa lista płac”, InsERT Gratyfikant), a dane trafiają do tabeli `zlistplac` (Pracownik, wartość L4, PPK pracownika, PPK pracodawcy, Pensja do przelania).

- `taskpane.html`, `app.js` – panel i integracja z Excelem (Office.js)
- `core.js` – parsowanie PDF (pdf.js) i dopasowanie nazwisk do tabel `XLS` / `Pracownicy_Etat`
- `manifest.xml` – manifest dodatku (instrukcja instalacji: `index.html`)

Przetwarzanie odbywa się w całości lokalnie; repozytorium nie zawiera żadnych danych płacowych.
