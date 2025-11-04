## 📦 Muuto Varenummer Konvertering (Mapping Lookup Værktøj)

Dette er et brugervenligt webværktøj designet til at give **hurtig og præcis konvertering** fra dine gamle Muuto Vare-varianter og EAN-numre til de nye varenumre.

Værktøjet bruger en sikker, centraliseret database for at sikre, at du altid modtager de korrekte og opdaterede produktoplysninger.

***

## 🚀 Sådan bruges værktøjet

### Adgang

Applikationen er hostet på Streamlit Cloud og kan tilgås via følgende link:

[**https://muuto-mapping.streamlit.app/**](https://muuto-mapping.streamlit.app/)

### 🛠️ Den Simple 2-Trins Konverteringsproces

Værktøjet er optimeret for maksimal hastighed og enkelhed for vores kunder.

#### **Trin 1: Indsæt ID'er**

1.  **Input:** Kopier listen over **gamle Vare-varianter** eller **EAN-numre**, du skal konvertere fra dit system.
2.  **Indsæt:** Indtast ID'erne i den store tekstboks i applikationen. ID'erne kan adskilles af linjeskift, kommaer eller mellemrum.
3.  **Bemærk:** Applikationen indlæser straks mappingdatabasen i baggrunden, når ID'erne er indsat.

#### **Trin 2: Se Resultater og Download**

1.  **Se Konvertering:** Resultattabellen viser alle matchede varer, inklusive det **Nye Varenummer**, **Beskrivelse**, **Familie** og **Kategori**.
2.  **Manglende ID'er:** ID'er, der ikke kunne findes, vises tydeligt i en advarselsboks, så du hurtigt kan tjekke for tastefejl.
3.  **Download:** Klik på knappen **Download Resultat som Excel-fil (.xlsx)** for at gemme den komplette tabel. En statusindikator vises, mens filen genereres.

***

## 🎯 Output Datafelter

Den resulterende Excel-fil indeholder altid de konverterede data i dette konsistente format:

| Kolonne Navn | Formål |
| :--- | :--- |
| **New Item No.** | Det **nye** Muuto varenummer (resultatet af konverteringen). |
| **OLD Item-variant** | Den oprindelige Vare-variant brugt til opslaget. |
| **Ean no.** | EAN/Stregkodenummeret forbundet med varen. |
| **Description** | Produktbeskrivelse. |
| **Family** | Den nye produktfamilie (f.eks. Sofaer, Belysning). |
| **Category** | Produktkategorien. |

***

## 📞 Support

For support vedrørende fejl i værktøjet, dataunøjagtigheder eller spørgsmål om de nye varenumre, bedes du kontakte din **Muuto salgsrepræsentant.**
